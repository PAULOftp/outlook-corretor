#!/usr/bin/env python3
"""
Antena da PauloIA: servidor local (só 127.0.0.1) que serve a página e vai buscar
agenda, e-mails e notícias em nome do browser. Apenas a biblioteca padrão do Python.

Rotas:
  GET  /                    -> rissa.html
  GET  /api/ping            -> estado da antena e da ligação Microsoft 365
  GET  /proxy?url=...       -> busca ICS/RSS (allowlist de domínios)
  POST /api/imap            -> e-mails via IMAP (só leitura, BODY.PEEK)
  POST /api/ms/start        -> inicia login Microsoft (device code)
  GET  /api/ms/status       -> estado do login Microsoft
  POST /api/ms/logout       -> esquece os tokens Microsoft
  GET  /api/ms/mail         -> e-mails Office 365 (Microsoft Graph, só leitura)
  GET  /api/ms/calendar     -> agenda Outlook (Microsoft Graph, só leitura)
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
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from email.header import decode_header, make_header
from email.utils import parseaddr, parsedate_to_datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HOST, PORT = "127.0.0.1", 8080
BASE = os.path.dirname(os.path.abspath(__file__))
TOKENS_FILE = os.path.join(BASE, "antena_tokens.json")

PROXY_HOSTS = {"calendar.google.com", "news.google.com", "outlook.office365.com", "outlook.live.com"}
IMAP_HOSTS = {"imap.gmail.com"}
ALLOWED_HOSTS = {"localhost:%d" % PORT, "127.0.0.1:%d" % PORT}
ALLOWED_ORIGINS = {"http://localhost:%d" % PORT, "http://127.0.0.1:%d" % PORT}
STATIC = {".html": "text/html; charset=utf-8", ".png": "image/png", ".ico": "image/x-icon",
          ".svg": "image/svg+xml", ".css": "text/css; charset=utf-8", ".js": "text/javascript; charset=utf-8"}
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36"

MS_SCOPES = "offline_access User.Read Mail.Read Calendars.Read"
GRAPH = "https://graph.microsoft.com/v1.0"

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


def http_json(url, data=None, headers=None, form=False, timeout=20):
    """Pedido HTTP que devolve (status, json) sem lançar exceção em 4xx/5xx."""
    body = None
    h = {"User-Agent": UA, "Accept": "application/json"}
    h.update(headers or {})
    if data is not None:
        if form:
            body = urllib.parse.urlencode(data).encode()
            h["Content-Type"] = "application/x-www-form-urlencoded"
        else:
            body = json.dumps(data).encode()
            h["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=body, headers=h)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, json.loads(r.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode("utf-8") or "{}")
        except Exception:
            return e.code, {}


# ---------------------------------------------------------------- Microsoft 365
_ms_lock = threading.Lock()
MS = {"state": "idle", "user_code": None, "uri": None, "error": None, "account": None}


def _ms_load():
    try:
        with open(TOKENS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _ms_save(d):
    with open(TOKENS_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f)


def _ms_friendly(j):
    desc = (j.get("error_description") or "") + " " + (j.get("error") or "")
    if "AADSTS65001" in desc or "AADSTS90094" in desc or "AADSTS90095" in desc or "admin" in desc.lower():
        return "A FTP Porto exige aprovação do administrador para esta app. Pede à TI para aprovar a app PauloIA (só leitura)."
    if "AADSTS7000218" in desc:
        return "Ativa 'Permitir fluxos de cliente público' na app do Azure (separador Autenticação) e tenta outra vez."
    if "AADSTS700016" in desc:
        return "Não encontrei esta app no diretório. Confirma o ID da aplicação (cliente) e o ID do diretório (inquilino)."
    if "AADSTS50194" in desc or "AADSTS90002" in desc:
        return "Confirma o ID do diretório (inquilino) da app no portal Azure."
    if "expired_token" in desc:
        return "O código expirou. Clica LIGAR outra vez."
    if "access_denied" in desc or "authorization_declined" in desc:
        return "O pedido de acesso foi recusado."
    if "invalid_grant" in desc:
        return "A sessão Microsoft expirou. Clica LIGAR outra vez."
    return (j.get("error_description") or j.get("error") or "Erro desconhecido").splitlines()[0]


def _ms_poll(client_id, tenant, device_code, interval, expires_in):
    url = "https://login.microsoftonline.com/%s/oauth2/v2.0/token" % urllib.parse.quote(tenant)
    deadline = time.time() + expires_in
    while time.time() < deadline:
        time.sleep(interval)
        st, j = http_json(url, {"grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                                "client_id": client_id, "device_code": device_code}, form=True)
        if st == 200 and j.get("access_token"):
            tok = {"client_id": client_id, "tenant": tenant, "refresh_token": j.get("refresh_token"),
                   "access_token": j["access_token"], "expires_at": time.time() + int(j.get("expires_in", 3600))}
            _, me = http_json(GRAPH + "/me?$select=displayName,userPrincipalName",
                              headers={"Authorization": "Bearer " + tok["access_token"]})
            tok["account"] = me.get("userPrincipalName") or me.get("displayName") or "conta Microsoft"
            _ms_save(tok)
            with _ms_lock:
                MS.update(state="connected", user_code=None, uri=None, error=None, account=tok["account"])
            say("  ✓ Microsoft 365 ligado")
            return
        err = j.get("error")
        if err == "authorization_pending":
            continue
        if err == "slow_down":
            interval += 5
            continue
        with _ms_lock:
            MS.update(state="error", user_code=None, uri=None, error=_ms_friendly(j))
        return
    with _ms_lock:
        MS.update(state="error", user_code=None, uri=None, error="O código expirou. Clica LIGAR outra vez.")


def ms_start(client_id, tenant):
    client_id = (client_id or "").strip()
    tenant = (tenant or "").strip() or "organizations"
    if not re.fullmatch(r"[0-9a-fA-F-]{36}", client_id):
        return 400, {"erro": "O ID da aplicação (cliente) tem de ter o formato xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx."}
    if not re.fullmatch(r"[A-Za-z0-9.-]{3,64}", tenant):
        return 400, {"erro": "ID do diretório (inquilino) inválido."}
    st, j = http_json("https://login.microsoftonline.com/%s/oauth2/v2.0/devicecode" % tenant,
                      {"client_id": client_id, "scope": MS_SCOPES}, form=True)
    if st != 200 or "device_code" not in j:
        with _ms_lock:
            MS.update(state="error", error=_ms_friendly(j))
        return 200, dict(MS)
    with _ms_lock:
        MS.update(state="pending", user_code=j["user_code"], uri=j.get("verification_uri"), error=None)
    threading.Thread(target=_ms_poll, daemon=True,
                     args=(client_id, tenant, j["device_code"], int(j.get("interval", 5)), int(j.get("expires_in", 900)))).start()
    return 200, dict(MS)


def ms_token():
    tok = _ms_load()
    if not tok.get("refresh_token") and not tok.get("access_token"):
        return None, "Microsoft 365 não está ligado. Abre ⚙ e clica LIGAR."
    if tok.get("access_token") and tok.get("expires_at", 0) - 120 > time.time():
        return tok["access_token"], None
    st, j = http_json("https://login.microsoftonline.com/%s/oauth2/v2.0/token" % tok["tenant"],
                      {"grant_type": "refresh_token", "client_id": tok["client_id"],
                       "refresh_token": tok["refresh_token"], "scope": MS_SCOPES}, form=True)
    if st != 200 or not j.get("access_token"):
        with _ms_lock:
            MS.update(state="error", error=_ms_friendly(j))
        return None, _ms_friendly(j)
    tok.update(access_token=j["access_token"], expires_at=time.time() + int(j.get("expires_in", 3600)))
    if j.get("refresh_token"):
        tok["refresh_token"] = j["refresh_token"]
    _ms_save(tok)
    return tok["access_token"], None


def graph_get(path, extra_headers=None):
    token, err = ms_token()
    if not token:
        return 401, {"erro": err}
    h = {"Authorization": "Bearer " + token}
    h.update(extra_headers or {})
    st, j = http_json(path if path.startswith("http") else GRAPH + path, headers=h)
    if st >= 400:
        msg = (j.get("error") or {}).get("message") or ("Erro %d do Microsoft Graph" % st)
        return st, {"erro": msg}
    return st, j


def ms_mail(top):
    sel = "id,internetMessageId,from,subject,receivedDateTime,bodyPreview,isRead"
    st, j = graph_get("/me/mailFolders/inbox/messages?$top=%d&$select=%s&$orderby=receivedDateTime%%20desc" % (top, sel))
    if st != 200:
        return st, j
    out = []
    for m in j.get("value", []):
        fr = (m.get("from") or {}).get("emailAddress") or {}
        out.append({"id": m.get("internetMessageId") or m.get("id"),
                    "remetente": fr.get("name") or fr.get("address") or "",
                    "email": fr.get("address") or "",
                    "assunto": m.get("subject") or "(sem assunto)",
                    "data": m.get("receivedDateTime"),
                    "trecho": re.sub(r"\s+", " ", m.get("bodyPreview") or "").strip()[:500],
                    "lido": bool(m.get("isRead"))})
    return 200, {"emails": out}


def ms_calendar(start, end):
    q = urllib.parse.urlencode({"startDateTime": start, "endDateTime": end, "$top": "100",
                                "$select": "subject,start,end,isAllDay,location,isCancelled",
                                "$orderby": "start/dateTime"})
    url = GRAPH + "/me/calendarView?" + q
    events = []
    for _ in range(5):
        st, j = graph_get(url, {"Prefer": 'outlook.timezone="UTC"'})
        if st != 200:
            return st, j
        for e in j.get("value", []):
            if e.get("isCancelled"):
                continue
            events.append({"titulo": e.get("subject") or "(sem título)",
                           "inicio": (e.get("start") or {}).get("dateTime"),
                           "fim": (e.get("end") or {}).get("dateTime"),
                           "diaTodo": bool(e.get("isAllDay")),
                           "local": ((e.get("location") or {}).get("displayName") or "")})
        url = j.get("@odata.nextLink")
        if not url:
            break
    return 200, {"eventos": events}


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
    server_version = "AntenaPauloIA/1.0"

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
                with _ms_lock:
                    st = dict(MS)
                return self._json(200, {"ok": True, "versao": 1, "ms": st})
            if u.path == "/proxy":
                return self._proxy((q.get("url") or [""])[0])
            if u.path == "/api/ms/status":
                with _ms_lock:
                    return self._json(200, dict(MS))
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
            if path == "/api/imap":
                host = (data.get("host") or "imap.gmail.com").strip().lower()
                if host not in IMAP_HOSTS:
                    return self._json(403, {"erro": "Servidor IMAP não autorizado: %s" % host})
                return self._json(*imap_fetch(host, data.get("usuario", ""), data.get("senhaApp", ""), data.get("quantidade", 20)))
            if path == "/api/ms/start":
                return self._json(*ms_start(data.get("clientId"), data.get("tenant")))
            if path == "/api/ms/logout":
                try:
                    os.remove(TOKENS_FILE)
                except FileNotFoundError:
                    pass
                with _ms_lock:
                    MS.update(state="idle", user_code=None, uri=None, error=None, account=None)
                return self._json(200, dict(MS))
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
        name = urllib.parse.unquote(path.lstrip("/")) or "rissa.html"
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
    tok = _ms_load()
    if tok.get("refresh_token"):
        MS.update(state="connected", account=tok.get("account"))
    try:
        srv = ThreadingHTTPServer((HOST, PORT), Antena)
    except OSError:
        say("\n  ✗ A porta %d já está ocupada." % PORT)
        say("    Provavelmente a antena já está a correr noutra janela (ou ficou aberto o servidor antigo).")
        say("    Fecha essa janela e corre o iniciar-rissa.bat outra vez.\n")
        try:
            input("  Carrega em Enter para fechar...")
        except EOFError:
            pass
        sys.exit(1)
    say("")
    say("  ⚡ ANTENA DA PAULOIA ONLINE — porta %d (só neste PC)." % PORT)
    say("     Abre http://localhost:%d/rissa.html  ·  Mantém esta janela aberta." % PORT)
    say("")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        say("\n  Antena desligada.")


if __name__ == "__main__":
    main()
