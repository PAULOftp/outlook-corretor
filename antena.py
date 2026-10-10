#!/usr/bin/env python3
"""
Antena da A.P.IAv1: servidor local (só 127.0.0.1) que serve a página e vai buscar
agenda, e-mails e notícias em nome do browser. Apenas a biblioteca padrão do Python.

Rotas:
  GET  /                    -> apiav1.html
  GET  /api/ping            -> estado da antena e da sincronização Microsoft 365
  GET  /proxy?url=...       -> busca ICS/RSS (allowlist de domínios)
  POST /emails              -> e-mails via IMAP (só leitura, BODY.PEEK)
  GET  /api/ms/status       -> estado da sincronização (ficheiro m365.json)
  GET  /api/ms/mail         -> e-mails Office 365 lidos do m365.json
  GET  /api/ms/calendar     -> agenda Outlook lida do m365.json
"""
import base64
import email
import html
import imaplib
import json
import os
import re
import socket
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request
from email.header import decode_header, make_header
from datetime import datetime, timedelta, timezone
from email.utils import parseaddr, parsedate_to_datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HOST, PORT = "127.0.0.1", 8080
BASE = os.path.dirname(os.path.abspath(__file__))

PROXY_HOSTS = {"calendar.google.com", "news.google.com", "outlook.office365.com", "outlook.live.com"}
IMAP_HOSTS = {"imap.gmail.com"}
ALLOWED_HOSTS = {"localhost:%d" % PORT, "127.0.0.1:%d" % PORT}
ALLOWED_ORIGINS = {"http://localhost:%d" % PORT, "http://127.0.0.1:%d" % PORT}
STATIC = {".html": "text/html; charset=utf-8", ".png": "image/png", ".ico": "image/x-icon",
          ".svg": "image/svg+xml", ".css": "text/css; charset=utf-8", ".js": "text/javascript; charset=utf-8"}
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36"


try:
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:
    pass


def say(msg):
    print(msg, flush=True)


# ---------------------------------------------------------------- HTTP de saída
class _AllowlistRedirect(urllib.request.HTTPRedirectHandler):
    max_redirections = 3

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if urllib.parse.urlsplit(newurl).hostname not in PROXY_HOSTS:
            raise urllib.error.HTTPError(newurl, 403, "Redirecionamento para domínio não autorizado", headers, fp)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


_proxy_opener = urllib.request.build_opener(_AllowlistRedirect)


# ---------------------------------------------------------------- Microsoft 365 (sem Azure)
# Uma tarefa agendada na app Claude (conector Microsoft 365, já aprovado pela empresa) grava
# o ficheiro m365.json nesta pasta. A antena só o lê: não é precisa permissão de administrador.
SNAP_FILE = os.path.join(BASE, "m365.json")
NO_SNAP = "Ainda não há dados do Outlook. Configura a tarefa agendada no Claude (⚙ → Microsoft 365)."


def snap_load():
    try:
        with open(SNAP_FILE, "r", encoding="utf-8-sig") as f:
            txt = f.read().strip()
        txt = re.sub(r"^```(?:json)?\s*", "", txt, flags=re.I)
        txt = re.sub(r"```$", "", txt).strip()
        j = json.loads(txt)
        return j if isinstance(j, dict) else None
    except Exception:
        return None


def ms_state():
    s = snap_load()
    if not s:
        return {"state": "idle", "account": None, "atualizado": None}
    return {"state": "connected", "account": s.get("conta"), "atualizado": s.get("atualizado")}


def _dt(v):
    """ISO 8601 (com ou sem fuso, ou só data) -> datetime em UTC."""
    if not v:
        return None
    v = str(v).strip().replace("Z", "+00:00")
    v = re.sub(r"(\.\d{6})\d+", r"\1", v)  # Python antigo só aceita 6 casas decimais
    try:
        d = datetime.fromisoformat(v) if "T" in v else datetime.fromisoformat(v + "T00:00:00+00:00")
    except ValueError:
        return None
    if d.tzinfo is None:
        d = d.astimezone()  # hora local deste PC
    return d.astimezone(timezone.utc)


def _naive(d):
    return d.strftime("%Y-%m-%dT%H:%M:%S.000")


def ms_mail(top):
    s = snap_load()
    if not s:
        return 404, {"erro": NO_SNAP}
    out = []
    for m in (s.get("emails") or [])[:top]:
        d = _dt(m.get("data"))
        balde = m.get("balde") if m.get("balde") in ("acao", "info", "ruido") else None
        out.append({"id": str(m.get("id") or m.get("internetMessageId") or "%s|%s" % (m.get("assunto"), m.get("data"))),
                    "remetente": str(m.get("remetente") or m.get("email") or ""), "email": str(m.get("email") or ""),
                    "assunto": str(m.get("assunto") or "(sem assunto)"),
                    "data": d.isoformat().replace("+00:00", "Z") if d else None,
                    "trecho": re.sub(r"\s+", " ", str(m.get("trecho") or "")).strip()[:500],
                    "lido": bool(m.get("lido")), "balde": balde,
                    "resumo": str(m.get("resumo"))[:220] if m.get("resumo") else None})
    conhecimento = [c for c in (s.get("conhecimento") or []) if isinstance(c, dict)][:20]  # factos sobre a FTP Porto extraídos pelo Claude
    return 200, {"emails": out, "conhecimento": conhecimento, "atualizado": s.get("atualizado")}


def ms_calendar(start, end):
    s = snap_load()
    if not s:
        return 404, {"erro": NO_SNAP}
    w0, w1 = _dt(start), _dt(end)
    events = []
    for e in s.get("eventos") or []:
        a = _dt(e.get("inicio"))
        if not a:
            continue
        b = _dt(e.get("fim")) or a
        if b <= a:  # dia inteiro com fim = início
            b = a + timedelta(days=1 if e.get("diaTodo") else 0)
        if w0 and w1 and not (b > w0 and a < w1) and not (w0 <= a < w1):
            continue
        events.append({"titulo": str(e.get("titulo") or "(sem título)"), "inicio": _naive(a), "fim": _naive(b),
                       "diaTodo": bool(e.get("diaTodo")), "local": str(e.get("local") or "")})
    return 200, {"eventos": events, "atualizado": s.get("atualizado")}


# ---------------------------------------------------------------- IMAP (Gmail)
def _dec(value):
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return str(value)


def _text_of(msg):
    plain, htm = None, None
    for part in msg.walk():
        if part.is_multipart() or part.get_content_disposition() == "attachment":
            continue
        ctype = part.get_content_type()
        if ctype not in ("text/plain", "text/html"):
            continue
        try:
            if (part.get("Content-Transfer-Encoding") or "").strip().lower() == "base64":
                # A mensagem pode vir cortada a meio (lemos só 60 KB): descodifica os blocos completos
                b64 = re.sub(r"[^A-Za-z0-9+/]", "", str(part.get_payload(decode=False) or ""))
                payload = base64.b64decode(b64[:len(b64) // 4 * 4])
            else:
                payload = part.get_payload(decode=True) or b""
        except Exception:
            continue
        cs = part.get_content_charset() or "utf-8"
        try:
            txt = payload.decode(cs, errors="replace")
        except LookupError:
            txt = payload.decode("latin-1", errors="replace")
        if ctype == "text/plain" and plain is None:
            plain = txt
        elif ctype == "text/html" and htm is None:
            htm = txt
    if plain is None and htm:
        t = re.sub(r"(?is)<(script|style|head).*?</\1>", " ", htm)
        t = re.sub(r"(?s)<[^>]+>", " ", t)
        plain = html.unescape(t)
    return re.sub(r"\s+", " ", plain or "").strip()[:500]


def imap_fetch(host, user, password, count):
    count = max(1, min(int(count or 20), 50))
    M = imaplib.IMAP4_SSL(host, 993, ssl_context=ssl.create_default_context(), timeout=25)
    try:
        try:
            M.login(user, password)
        except imaplib.IMAP4.error:
            return 401, {"erro": "Senha de app inválida ou verificação em 2 passos desativada. Cria uma nova em myaccount.google.com/apppasswords."}
        M.select("INBOX", readonly=True)  # EXAMINE: caixa só de leitura
        typ, data = M.uid("search", None, "ALL")
        uids = data[0].split()[-count:] if data and data[0] else []
        if not uids:
            return 200, {"emails": []}
        # BODY.PEEK: nunca marca nada como lido. Só os primeiros 60 KB de cada mensagem.
        typ, data = M.uid("fetch", b",".join(uids), "(FLAGS BODY.PEEK[]<0.60000>)")
        items, cur = [], None
        for part in data:
            if isinstance(part, tuple):
                cur = {"meta": part[0], "raw": part[1]}
                items.append(cur)
            elif isinstance(part, bytes) and cur is not None:
                cur["meta"] += part
        out = []
        for it in items:
            meta = it["meta"].decode("latin-1", errors="replace")
            uid = (re.search(r"UID (\d+)", meta) or [None, "?"])[1]
            flags = (re.search(r"FLAGS \(([^)]*)\)", meta) or [None, ""])[1]
            msg = email.message_from_bytes(it["raw"] or b"")
            try:
                dt = parsedate_to_datetime(msg.get("Date")).isoformat()
            except Exception:
                dt = None
            name, addr = parseaddr(_dec(msg.get("From")))
            out.append({"id": (msg.get("Message-ID") or "").strip() or "%s:%s" % (user, uid),
                        "remetente": name or addr, "email": addr,
                        "assunto": _dec(msg.get("Subject")) or "(sem assunto)",
                        "data": dt, "trecho": _text_of(msg), "lido": "\\Seen" in flags})
        out.sort(key=lambda e: e["data"] or "", reverse=True)
        return 200, {"emails": out}
    finally:
        try:
            M.logout()
        except Exception:
            pass


# ---------------------------------------------------------------- Servidor
def _mask(path):
    return path.split("?", 1)[0] + ("?…" if "?" in path else "")


class Antena(BaseHTTPRequestHandler):
    server_version = "AntenaAPIAv1/1.0"

    def log_message(self, fmt, *args):  # nunca escreve links secretos nem credenciais
        sys.stderr.write("  · %s %s\n" % (self.command, _mask(self.path)))

    def _guard(self):
        """Só aceita pedidos feitos a localhost e vindos da própria página (protege contra sites maliciosos)."""
        if self.headers.get("Host", "") not in ALLOWED_HOSTS:
            self._json(403, {"erro": "Host não autorizado"})
            return False
        origin = self.headers.get("Origin")
        if origin and origin not in ALLOWED_ORIGINS:
            self._json(403, {"erro": "Origem não autorizada"})
            return False
        return True

    def _cors(self):
        origin = self.headers.get("Origin")
        if origin in ALLOWED_ORIGINS:
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Vary", "Origin")

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > 65536:
            raise ValueError("pedido demasiado grande")
        return json.loads(self.rfile.read(n).decode("utf-8") or "{}")

    def do_OPTIONS(self):
        if not self._guard():
            return
        self.send_response(204)
        self._cors()
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Max-Age", "600")
        self.end_headers()

    def do_GET(self):
        if not self._guard():
            return
        u = urllib.parse.urlsplit(self.path)
        q = urllib.parse.parse_qs(u.query)
        try:
            if u.path == "/api/ping":
                return self._json(200, {"ok": True, "versao": 2, "ms": ms_state()})
            if u.path == "/proxy":
                return self._proxy((q.get("url") or [""])[0])
            if u.path == "/api/ms/status":
                return self._json(200, ms_state())
            if u.path == "/api/ms/mail":
                return self._json(*ms_mail(max(1, min(int((q.get("top") or ["20"])[0]), 50))))
            if u.path == "/api/ms/calendar":
                return self._json(*ms_calendar((q.get("start") or [""])[0], (q.get("end") or [""])[0]))
            return self._static(u.path)
        except (socket.timeout, urllib.error.URLError) as e:
            return self._json(502, {"erro": "Sem ligação ao serviço externo (%s)." % type(e).__name__})
        except Exception as e:
            return self._json(500, {"erro": "Erro na antena: %s" % type(e).__name__})

    def do_POST(self):
        if not self._guard():
            return
        path = urllib.parse.urlsplit(self.path).path
        try:
            data = self._body()
            if path in ("/emails", "/api/imap"):
                host = (data.get("host") or "imap.gmail.com").strip().lower()
                if host not in IMAP_HOSTS:
                    return self._json(403, {"erro": "Servidor IMAP não autorizado: %s" % host})
                return self._json(*imap_fetch(host, data.get("usuario", ""), data.get("senhaApp", ""), data.get("quantidade", 20)))
            return self._json(404, {"erro": "Rota desconhecida"})
        except (socket.timeout, OSError) as e:
            return self._json(502, {"erro": "Sem ligação ao servidor de e-mail (%s)." % type(e).__name__})
        except Exception as e:
            return self._json(500, {"erro": "Erro na antena: %s" % type(e).__name__})

    def _proxy(self, url):
        p = urllib.parse.urlsplit(url)
        if p.scheme != "https" or p.hostname not in PROXY_HOSTS:
            return self._json(403, {"erro": "Domínio não autorizado na antena"})
        req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
        try:
            with _proxy_opener.open(req, timeout=20) as r:
                body = r.read(5 * 1024 * 1024)
                ctype = r.headers.get("Content-Type", "text/plain; charset=utf-8")
        except urllib.error.HTTPError as e:
            return self._json(e.code if e.code >= 400 else 502, {"erro": "A fonte respondeu %d" % e.code})
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def _static(self, path):
        name = urllib.parse.unquote(path.lstrip("/")) or "apiav1.html"
        ext = os.path.splitext(name)[1].lower()
        full = os.path.join(BASE, name)
        if "/" in name or "\\" in name or name.startswith(".") or ext not in STATIC or not os.path.isfile(full):
            return self._json(404, {"erro": "Ficheiro não encontrado: %s" % name})
        with open(full, "rb") as f:
            body = f.read()
        self.send_response(200)
        self.send_header("Content-Type", STATIC[ext])
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)


def main():
    try:
        srv = ThreadingHTTPServer((HOST, PORT), Antena)
    except OSError:
        say("\n  ✗ A porta %d já está ocupada." % PORT)
        say("    Provavelmente a antena já está a correr noutra janela (ou ficou aberto o servidor antigo).")
        say("    Fecha essa janela e corre o iniciar-apiav1.bat outra vez.\n")
        try:
            input("  Carrega em Enter para fechar...")
        except EOFError:
            pass
        sys.exit(1)
    say("")
    say("  ⚡ ANTENA DA A.P.IAv1 ONLINE — porta %d (só neste PC)." % PORT)
    say("     Abre http://localhost:%d/apiav1.html  ·  Mantém esta janela aberta." % PORT)
    say("")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        say("\n  Antena desligada.")


if __name__ == "__main__":
    main()
