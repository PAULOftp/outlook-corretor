#!/usr/bin/env node
/*
  ANTENA DO JARVIS (PauloIA) — Node puro, sem npm install.
  Corre com:  node server.js   (e deixa o terminal aberto)

  Escuta só em 127.0.0.1:4242. Rotas:
    GET  /api/ping          estado da antena e da ligação Microsoft 365
    GET  /proxy?url=...     ICS / RSS (allowlist: calendar.google.com, news.google.com, Outlook ICS)
    POST /emails            e-mails via IMAP, só leitura (EXAMINE + BODY.PEEK)
    POST /api/ms/start      inicia o login Microsoft (device code)
    GET  /api/ms/status     estado do login Microsoft
    POST /api/ms/logout     esquece a sessão Microsoft
    GET  /api/ms/mail       e-mails Office 365 (Microsoft Graph, só leitura)
    GET  /api/ms/calendar   agenda Outlook (Microsoft Graph, só leitura)
*/
"use strict";
const http = require("http");
const https = require("https");
const tls = require("tls");
const fs = require("fs");
const path = require("path");

const HOST = "127.0.0.1", PORT = 4242;
const TOKENS_FILE = path.join(__dirname, "antena_tokens.json");
const PROXY_HOSTS = new Set(["calendar.google.com", "news.google.com", "outlook.office365.com", "outlook.live.com"]);
const IMAP_HOSTS = new Set(["imap.gmail.com"]);
// "null" é a origem de uma página aberta com duplo clique (file://)
const ALLOWED_ORIGINS = new Set(["null", "http://localhost:" + PORT, "http://127.0.0.1:" + PORT]);
const ALLOWED_HOSTS = new Set(["localhost:" + PORT, "127.0.0.1:" + PORT]);
const STATIC = { ".html": "text/html; charset=utf-8", ".png": "image/png", ".ico": "image/x-icon", ".svg": "image/svg+xml" };
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36";
const MS_SCOPES = "offline_access User.Read Mail.Read Calendars.Read";
const GRAPH = "https://graph.microsoft.com/v1.0";

/* ------------------------------------------------------------ HTTP de saída */
function request(url, opts = {}) {
  const { method = "GET", headers = {}, body = null, maxRedirects = 0, allow = null, timeout = 20000 } = opts;
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const h = { "User-Agent": UA, ...headers };
    if (body) h["Content-Length"] = Buffer.byteLength(body);
    const req = https.request(u, { method, headers: h, timeout }, res => {
      if ([301, 302, 303, 307, 308].includes(res.statusCode) && res.headers.location && maxRedirects > 0) {
        res.resume();
        const next = new URL(res.headers.location, u);
        if (allow && !allow.has(next.hostname)) return reject(Object.assign(new Error("Redirecionamento para domínio não autorizado"), { status: 403 }));
        const keep = res.statusCode === 307 || res.statusCode === 308;
        return resolve(request(next.href, { ...opts, method: keep ? method : "GET", body: keep ? body : null, maxRedirects: maxRedirects - 1 }));
      }
      const chunks = [];
      let size = 0;
      res.on("data", c => { size += c.length; if (size > 5 * 1024 * 1024) req.destroy(new Error("resposta demasiado grande")); else chunks.push(c); });
      res.on("end", () => resolve({ status: res.statusCode, headers: res.headers, body: Buffer.concat(chunks) }));
      res.on("error", reject);
    });
    req.on("timeout", () => req.destroy(new Error("tempo esgotado")));
    req.on("error", reject);
    if (body) req.write(body);
    req.end();
  });
}
async function httpJson(url, opts) {
  const r = await request(url, opts);
  let j = {};
  try { j = JSON.parse(r.body.toString("utf8") || "{}"); } catch (e) {}
  return [r.status, j];
}
const form = o => new URLSearchParams(o).toString();
const FORM = { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" };

/* ------------------------------------------------------------ Microsoft 365 */
const MS = { state: "idle", user_code: null, uri: null, error: null, account: null };
function msLoad() { try { return JSON.parse(fs.readFileSync(TOKENS_FILE, "utf8")); } catch (e) { return {}; } }
function msSave(d) { fs.writeFileSync(TOKENS_FILE, JSON.stringify(d)); }
function msFriendly(j) {
  const d = (j.error_description || "") + " " + (j.error || "");
  if (/AADSTS65001|AADSTS90094|AADSTS90095/.test(d) || /admin/i.test(d)) return "A FTP Porto exige aprovação do administrador para esta app. Pede à TI para aprovar a app PauloIA (só leitura).";
  if (/AADSTS7000218/.test(d)) return "Ativa 'Permitir fluxos de cliente público' na app do Azure (separador Autenticação) e tenta outra vez.";
  if (/AADSTS700016/.test(d)) return "Não encontrei esta app no diretório. Confirma o ID da aplicação (cliente) e o ID do diretório (inquilino).";
  if (/AADSTS50194|AADSTS90002/.test(d)) return "Confirma o ID do diretório (inquilino) da app no portal Azure.";
  if (/expired_token/.test(d)) return "O código expirou. Clica LIGAR outra vez.";
  if (/access_denied|authorization_declined/.test(d)) return "O pedido de acesso foi recusado.";
  if (/invalid_grant/.test(d)) return "A sessão Microsoft expirou. Clica LIGAR outra vez.";
  return String(j.error_description || j.error || "Erro desconhecido").split(/\r?\n/)[0];
}
const sleep = ms => new Promise(r => setTimeout(r, ms));
async function msPoll(clientId, tenant, deviceCode, interval, expiresIn) {
  const url = "https://login.microsoftonline.com/" + encodeURIComponent(tenant) + "/oauth2/v2.0/token";
  const deadline = Date.now() + expiresIn * 1000;
  while (Date.now() < deadline) {
    await sleep(interval * 1000);
    let st, j;
    try { [st, j] = await httpJson(url, { method: "POST", headers: FORM, body: form({ grant_type: "urn:ietf:params:oauth:grant-type:device_code", client_id: clientId, device_code: deviceCode }) }); }
    catch (e) { continue; }
    if (st === 200 && j.access_token) {
      const tok = { client_id: clientId, tenant, refresh_token: j.refresh_token, access_token: j.access_token, expires_at: Date.now() / 1000 + Number(j.expires_in || 3600) };
      try { const [, me] = await httpJson(GRAPH + "/me?$select=displayName,userPrincipalName", { headers: { Authorization: "Bearer " + tok.access_token } }); tok.account = me.userPrincipalName || me.displayName; } catch (e) {}
      tok.account = tok.account || "conta Microsoft";
      msSave(tok);
      Object.assign(MS, { state: "connected", user_code: null, uri: null, error: null, account: tok.account });
      console.log("  ✓ Microsoft 365 ligado");
      return;
    }
    if (j.error === "authorization_pending") continue;
    if (j.error === "slow_down") { interval += 5; continue; }
    Object.assign(MS, { state: "error", user_code: null, uri: null, error: msFriendly(j) });
    return;
  }
  Object.assign(MS, { state: "error", user_code: null, uri: null, error: "O código expirou. Clica LIGAR outra vez." });
}
async function msStart(clientId, tenant) {
  clientId = String(clientId || "").trim();
  tenant = String(tenant || "").trim() || "organizations";
  if (!/^[0-9a-fA-F-]{36}$/.test(clientId)) return [400, { erro: "O ID da aplicação (cliente) tem de ter o formato xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx." }];
  if (!/^[A-Za-z0-9.-]{3,64}$/.test(tenant)) return [400, { erro: "ID do diretório (inquilino) inválido." }];
  const [st, j] = await httpJson("https://login.microsoftonline.com/" + tenant + "/oauth2/v2.0/devicecode", { method: "POST", headers: FORM, body: form({ client_id: clientId, scope: MS_SCOPES }) });
  if (st !== 200 || !j.device_code) { Object.assign(MS, { state: "error", error: msFriendly(j) }); return [200, { ...MS }]; }
  Object.assign(MS, { state: "pending", user_code: j.user_code, uri: j.verification_uri, error: null });
  msPoll(clientId, tenant, j.device_code, Number(j.interval || 5), Number(j.expires_in || 900)).catch(() => {});
  return [200, { ...MS }];
}
async function msToken() {
  const tok = msLoad();
  if (!tok.refresh_token && !tok.access_token) return [null, "Microsoft 365 não está ligado. Abre ⚙ e clica LIGAR."];
  if (tok.access_token && (tok.expires_at || 0) - 120 > Date.now() / 1000) return [tok.access_token, null];
  const [st, j] = await httpJson("https://login.microsoftonline.com/" + tok.tenant + "/oauth2/v2.0/token",
    { method: "POST", headers: FORM, body: form({ grant_type: "refresh_token", client_id: tok.client_id, refresh_token: tok.refresh_token, scope: MS_SCOPES }) });
  if (st !== 200 || !j.access_token) { Object.assign(MS, { state: "error", error: msFriendly(j) }); return [null, msFriendly(j)]; }
  Object.assign(tok, { access_token: j.access_token, expires_at: Date.now() / 1000 + Number(j.expires_in || 3600) });
  if (j.refresh_token) tok.refresh_token = j.refresh_token;
  msSave(tok);
  return [tok.access_token, null];
}
async function graphGet(pathOrUrl, extra = {}) {
  const [token, err] = await msToken();
  if (!token) return [401, { erro: err }];
  const [st, j] = await httpJson(pathOrUrl.startsWith("http") ? pathOrUrl : GRAPH + pathOrUrl, { headers: { Authorization: "Bearer " + token, Accept: "application/json", ...extra } });
  if (st >= 400) return [st, { erro: (j.error && j.error.message) || "Erro " + st + " do Microsoft Graph" }];
  return [st, j];
}
async function msMail(top) {
  const sel = "id,internetMessageId,from,subject,receivedDateTime,bodyPreview,isRead";
  const [st, j] = await graphGet("/me/mailFolders/inbox/messages?$top=" + top + "&$select=" + sel + "&$orderby=receivedDateTime%20desc");
  if (st !== 200) return [st, j];
  return [200, { emails: (j.value || []).map(m => {
    const fr = (m.from && m.from.emailAddress) || {};
    return { id: m.internetMessageId || m.id, remetente: fr.name || fr.address || "", email: fr.address || "", assunto: m.subject || "(sem assunto)",
      data: m.receivedDateTime, trecho: String(m.bodyPreview || "").replace(/\s+/g, " ").trim().slice(0, 500), lido: !!m.isRead };
  }) }];
}
async function msCalendar(start, end) {
  const q = new URLSearchParams({ startDateTime: start, endDateTime: end, $top: "100", $select: "subject,start,end,isAllDay,location,isCancelled", $orderby: "start/dateTime" });
  let url = GRAPH + "/me/calendarView?" + q.toString();
  const eventos = [];
  for (let i = 0; i < 5 && url; i++) {
    const [st, j] = await graphGet(url, { Prefer: 'outlook.timezone="UTC"' });
    if (st !== 200) return [st, j];
    for (const e of j.value || []) {
      if (e.isCancelled) continue;
      eventos.push({ titulo: e.subject || "(sem título)", inicio: e.start && e.start.dateTime, fim: e.end && e.end.dateTime, diaTodo: !!e.isAllDay, local: (e.location && e.location.displayName) || "" });
    }
    url = j["@odata.nextLink"];
  }
  return [200, { eventos }];
}

/* ------------------------------------------------------------ MIME */
function decodeCharset(buf, cs) {
  cs = String(cs || "utf-8").toLowerCase().replace(/^"|"$/g, "");
  try { return new TextDecoder(cs).decode(buf); } catch (e) {}
  try { return new TextDecoder("windows-1252").decode(buf); } catch (e) { return buf.toString("latin1"); }
}
function decodeQP(s) {
  s = s.replace(/=\r?\n/g, "");
  return Buffer.from(s.replace(/=([0-9A-Fa-f]{2})/g, (m, h) => String.fromCharCode(parseInt(h, 16))), "latin1");
}
function decodeB64(s) { // tolera base64 cortado a meio (lemos só o início da mensagem)
  const c = s.replace(/[^A-Za-z0-9+/]/g, "");
  return Buffer.from(c.slice(0, Math.floor(c.length / 4) * 4), "base64");
}
function decodeWords(v) {
  if (!v) return "";
  if (/[\x80-\xff]/.test(v)) v = decodeCharset(Buffer.from(v, "latin1"), "utf-8"); // cabeçalho em UTF-8 cru
  return v.replace(/(=\?[^?]+\?[BbQq]\?[^?]*\?=)\s+(?==\?)/g, "$1")
    .replace(/=\?([^?]+)\?([BbQq])\?([^?]*)\?=/g, (m, cs, enc, txt) =>
      decodeCharset(enc.toUpperCase() === "B" ? Buffer.from(txt, "base64") : decodeQP(txt.replace(/_/g, " ")), cs.split("*")[0]));
}
function splitHeaders(s) {
  let i = s.indexOf("\r\n\r\n"), sep = 4;
  if (i < 0) { i = s.indexOf("\n\n"); sep = 2; }
  const head = i < 0 ? s : s.slice(0, i), body = i < 0 ? "" : s.slice(i + sep);
  const h = {};
  for (const line of head.replace(/\r?\n[ \t]+/g, " ").split(/\r?\n/)) {
    const k = line.indexOf(":");
    if (k > 0) { const key = line.slice(0, k).trim().toLowerCase(); if (!(key in h)) h[key] = line.slice(k + 1).trim(); }
  }
  return { h, body };
}
function param(v, name) {
  const m = String(v || "").match(new RegExp(name + '\\s*=\\s*(?:"([^"]*)"|([^;\\s]+))', "i"));
  return m ? (m[1] !== undefined ? m[1] : m[2]) : null;
}
function textParts(s, depth = 0) {
  const { h, body } = splitHeaders(s);
  const ctRaw = h["content-type"] || "text/plain", ct = ctRaw.toLowerCase();
  if (ct.startsWith("multipart/")) {
    const b = param(ctRaw, "boundary");
    if (!b || depth > 6) return {};
    const r = {};
    for (const p of body.split("--" + b).slice(1)) {
      if (p.startsWith("--")) break;
      const sub = textParts(p.replace(/^\r?\n/, ""), depth + 1);
      r.plain = r.plain || sub.plain;
      r.html = r.html || sub.html;
      if (r.plain) break;
    }
    return r;
  }
  if (/^attachment/i.test(h["content-disposition"] || "")) return {};
  if (!ct.startsWith("text/plain") && !ct.startsWith("text/html")) return {};
  const cte = String(h["content-transfer-encoding"] || "").trim().toLowerCase();
  const buf = cte === "base64" ? decodeB64(body) : cte === "quoted-printable" ? decodeQP(body) : Buffer.from(body, "latin1");
  const txt = decodeCharset(buf, param(ctRaw, "charset"));
  return ct.startsWith("text/html") ? { html: txt } : { plain: txt };
}
const ENT = { nbsp: " ", amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", euro: "€", hellip: "…", ndash: "–", mdash: "—", laquo: "«", raquo: "»", ordm: "º", ordf: "ª" };
function stripHtml(t) {
  return t.replace(/<(script|style|head)[\s\S]*?<\/\1>/gi, " ").replace(/<[^>]+>/g, " ")
    .replace(/&(#\d+|#x[0-9a-f]+|[a-z]+);/gi, (m, e) => e[0] === "#" ? String.fromCodePoint(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : (ENT[e.toLowerCase()] || m));
}
function parseMessage(raw, user, uid, flags) {
  const { h } = splitHeaders(raw);
  const from = decodeWords(h.from || "");
  const m = from.match(/^\s*"?([^"<]*?)"?\s*<([^>]+)>/);
  const tp = textParts(raw);
  const texto = tp.plain != null ? tp.plain : tp.html ? stripHtml(tp.html) : "";
  const d = h.date ? new Date(h.date) : null;
  return {
    id: (h["message-id"] || "").trim() || user + ":" + uid,
    remetente: m ? (m[1].trim() || m[2]) : from.trim(), email: m ? m[2].trim() : from.trim(),
    assunto: decodeWords(h.subject || "") || "(sem assunto)",
    data: d && !isNaN(d) ? d.toISOString() : null,
    trecho: texto.replace(/\s+/g, " ").trim().slice(0, 500),
    lido: /\\Seen/i.test(flags || "")
  };
}

/* ------------------------------------------------------------ Cliente IMAP mínimo (tls nativo) */
class Imap {
  constructor(host) { this.host = host; this.buf = Buffer.alloc(0); this.n = 0; this.scan = 0; this.respStart = 0; this.pending = null; }
  connect() {
    return new Promise((resolve, reject) => {
      this.pending = { greeting: true, resolve, reject };
      this.sock = tls.connect({ host: this.host, port: 993, servername: this.host, timeout: 25000 });
      this.sock.on("data", d => { this.buf = Buffer.concat([this.buf, d]); this.check(); });
      this.sock.on("timeout", () => this.sock.destroy(new Error("O servidor de e-mail não respondeu a tempo.")));
      this.sock.on("error", e => { if (this.pending) { const p = this.pending; this.pending = null; p.reject(e); } });
      this.sock.on("close", () => { if (this.pending) { const p = this.pending; this.pending = null; p.reject(new Error("Ligação IMAP fechada.")); } });
    });
  }
  // Lê respostas completas, saltando literais {n} (o corpo dos e-mails pode conter qualquer texto)
  check() {
    if (!this.pending) return;
    let pos = this.scan;
    for (;;) {
      const nl = this.buf.indexOf("\r\n", pos);
      if (nl < 0) break;
      const line = this.buf.slice(pos, nl).toString("latin1");
      const lit = line.match(/\{(\d+)\}$/);
      if (lit) {
        const end = nl + 2 + parseInt(lit[1], 10);
        if (this.buf.length < end) break;
        pos = end;
        continue;
      }
      pos = nl + 2;
      const first = this.buf.slice(this.respStart, Math.min(this.respStart + 64, nl)).toString("latin1");
      const p = this.pending;
      if (p.greeting ? first.startsWith("* ") : first.startsWith(p.tag + " ")) {
        const data = this.buf.slice(0, pos);
        this.buf = this.buf.slice(pos);
        this.scan = 0; this.respStart = 0; this.pending = null;
        const status = (p.greeting ? first.slice(2) : first.slice(p.tag.length + 1)).split(" ")[0].toUpperCase();
        p.resolve({ status, data, line: first });
        return;
      }
      this.respStart = pos;
    }
    this.scan = pos;
  }
  cmd(text) {
    return new Promise((resolve, reject) => {
      const tag = "A" + (++this.n);
      this.pending = { tag, resolve, reject };
      this.sock.write(tag + " " + text + "\r\n");
      this.check();
    });
  }
  close() { try { this.sock.write("Z LOGOUT\r\n"); this.sock.end(); } catch (e) {} }
}
const q = s => '"' + String(s).replace(/[\\"]/g, m => "\\" + m) + '"';
async function imapFetch(host, user, pass, count) {
  count = Math.max(1, Math.min(parseInt(count, 10) || 20, 50));
  const c = new Imap(host);
  await c.connect();
  try {
    const login = await c.cmd("LOGIN " + q(user) + " " + q(pass));
    if (login.status !== "OK") return [401, { erro: "Senha de app inválida ou verificação em 2 passos desativada. Cria uma nova em myaccount.google.com/apppasswords." }];
    const ex = await c.cmd("EXAMINE INBOX"); // modo só de leitura
    if (ex.status !== "OK") return [502, { erro: "Não consegui abrir a caixa de entrada." }];
    const se = await c.cmd("UID SEARCH ALL");
    const m = se.data.toString("latin1").match(/\* SEARCH([\d ]*)/);
    const uids = m ? m[1].trim().split(/\s+/).filter(Boolean).slice(-count) : [];
    if (!uids.length) return [200, { emails: [] }];
    // BODY.PEEK: nunca marca nada como lido. Só os primeiros 60 KB de cada mensagem.
    const fe = await c.cmd("UID FETCH " + uids.join(",") + " (FLAGS BODY.PEEK[]<0.60000>)");
    const s = fe.data.toString("latin1"); // latin1 = 1 byte por carácter, os tamanhos dos literais batem certo
    const out = [];
    const start = /\* \d+ FETCH \(/g;
    let mm;
    while ((mm = start.exec(s))) {
      const litRe = /BODY\[\](?:<\d+>)? \{(\d+)\}\r\n/g;
      litRe.lastIndex = mm.index;
      const lm = litRe.exec(s);
      if (!lm) break;
      const bodyStart = lm.index + lm[0].length, len = parseInt(lm[1], 10), after = bodyStart + len;
      const lineEnd = s.indexOf("\r\n", after);
      const meta = s.slice(mm.index, lm.index) + s.slice(after, lineEnd < 0 ? s.length : lineEnd);
      const uid = (meta.match(/UID (\d+)/) || [])[1] || "?";
      const flags = (meta.match(/FLAGS \(([^)]*)\)/) || [])[1] || "";
      try { out.push(parseMessage(s.slice(bodyStart, after), user, uid, flags)); } catch (e) {}
      start.lastIndex = after;
    }
    out.sort((a, b) => String(b.data || "").localeCompare(String(a.data || "")));
    return [200, { emails: out }];
  } finally { c.close(); }
}

/* ------------------------------------------------------------ Servidor */
const mask = p => p.split("?")[0] + (p.includes("?") ? "?…" : ""); // nunca escreve links secretos nem credenciais
function cors(req, res) {
  const o = req.headers.origin;
  if (o && ALLOWED_ORIGINS.has(o)) { res.setHeader("Access-Control-Allow-Origin", o); res.setHeader("Vary", "Origin"); }
}
function send(req, res, code, obj) {
  const body = Buffer.from(JSON.stringify(obj), "utf8");
  cors(req, res);
  res.writeHead(code, { "Content-Type": "application/json; charset=utf-8", "Content-Length": body.length, "Cache-Control": "no-store" });
  res.end(body);
}
function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", c => { size += c.length; if (size > 65536) { reject(new Error("pedido demasiado grande")); req.destroy(); } else chunks.push(c); });
    req.on("end", () => { try { resolve(JSON.parse(Buffer.concat(chunks).toString("utf8") || "{}")); } catch (e) { reject(e); } });
    req.on("error", reject);
  });
}
async function proxy(req, res, target) {
  let u;
  try { u = new URL(target); } catch (e) { return send(req, res, 400, { erro: "URL inválido" }); }
  if (u.protocol !== "https:" || !PROXY_HOSTS.has(u.hostname)) return send(req, res, 403, { erro: "Domínio não autorizado na antena" });
  const r = await request(u.href, { headers: { Accept: "*/*" }, maxRedirects: 3, allow: PROXY_HOSTS });
  if (r.status >= 400) return send(req, res, r.status, { erro: "A fonte respondeu " + r.status });
  cors(req, res);
  res.writeHead(200, { "Content-Type": r.headers["content-type"] || "text/plain; charset=utf-8", "Content-Length": r.body.length, "Cache-Control": "no-store" });
  res.end(r.body);
}
function serveStatic(req, res, p) {
  const name = decodeURIComponent(p.replace(/^\/+/, "")) || "jarvis.html";
  const ext = path.extname(name).toLowerCase(), full = path.join(__dirname, name);
  if (/[\/\\]/.test(name) || name.startsWith(".") || !STATIC[ext] || !fs.existsSync(full)) return send(req, res, 404, { erro: "Ficheiro não encontrado: " + name });
  const body = fs.readFileSync(full);
  res.writeHead(200, { "Content-Type": STATIC[ext], "Content-Length": body.length, "Cache-Control": "no-store" });
  res.end(body);
}

const server = http.createServer(async (req, res) => {
  console.log("  · " + req.method + " " + mask(req.url));
  // Só aceita pedidos feitos a localhost e vindos do próprio Jarvis (protege contra sites maliciosos)
  if (!ALLOWED_HOSTS.has(req.headers.host || "")) return send(req, res, 403, { erro: "Host não autorizado" });
  if (req.headers.origin && !ALLOWED_ORIGINS.has(req.headers.origin)) return send(req, res, 403, { erro: "Origem não autorizada" });
  if (req.method === "OPTIONS") {
    cors(req, res);
    res.writeHead(204, { "Access-Control-Allow-Methods": "GET, POST, OPTIONS", "Access-Control-Allow-Headers": "Content-Type", "Access-Control-Max-Age": "600" });
    return res.end();
  }
  const u = new URL(req.url, "http://" + HOST + ":" + PORT);
  try {
    if (req.method === "GET") {
      if (u.pathname === "/api/ping") return send(req, res, 200, { ok: true, versao: 1, ms: { ...MS } });
      if (u.pathname === "/proxy") return await proxy(req, res, u.searchParams.get("url") || "");
      if (u.pathname === "/api/ms/status") return send(req, res, 200, { ...MS });
      if (u.pathname === "/api/ms/mail") return send(req, res, ...(await msMail(Math.max(1, Math.min(parseInt(u.searchParams.get("top"), 10) || 20, 50)))));
      if (u.pathname === "/api/ms/calendar") return send(req, res, ...(await msCalendar(u.searchParams.get("start") || "", u.searchParams.get("end") || "")));
      return serveStatic(req, res, u.pathname);
    }
    if (req.method === "POST") {
      const data = await readBody(req);
      if (u.pathname === "/emails" || u.pathname === "/api/imap") {
        const host = String(data.host || "imap.gmail.com").trim().toLowerCase();
        if (!IMAP_HOSTS.has(host)) return send(req, res, 403, { erro: "Servidor IMAP não autorizado: " + host });
        return send(req, res, ...(await imapFetch(host, data.usuario || "", data.senhaApp || "", data.quantidade)));
      }
      if (u.pathname === "/api/ms/start") return send(req, res, ...(await msStart(data.clientId, data.tenant)));
      if (u.pathname === "/api/ms/logout") {
        try { fs.unlinkSync(TOKENS_FILE); } catch (e) {}
        Object.assign(MS, { state: "idle", user_code: null, uri: null, error: null, account: null });
        return send(req, res, 200, { ...MS });
      }
    }
    return send(req, res, 404, { erro: "Rota desconhecida" });
  } catch (e) {
    if (e.status === 403) return send(req, res, 403, { erro: e.message });
    const net = /ENOTFOUND|ECONNREFUSED|ECONNRESET|ETIMEDOUT|EAI_AGAIN|tempo esgotado|não respondeu|fechada/.test(e.code + " " + e.message);
    return send(req, res, 502, { erro: net ? "Sem ligação ao serviço externo." : "Erro na antena: " + e.message });
  }
});

server.on("error", e => {
  if (e.code === "EADDRINUSE") {
    console.log("\n  ✗ A porta " + PORT + " já está ocupada.");
    console.log("    Provavelmente a antena já está a correr noutro terminal. Fecha-o e corre outra vez: node server.js\n");
  } else console.log("\n  ✗ Não consegui arrancar a antena: " + e.message + "\n");
  process.exit(1);
});

const tok = msLoad();
if (tok.refresh_token) Object.assign(MS, { state: "connected", account: tok.account || null });
server.listen(PORT, HOST, () => {
  console.log("");
  console.log("  ⚡ ANTENA DO JARVIS ONLINE — porta " + PORT + ". Já podes abrir o jarvis.html.");
  console.log("     Mantém este terminal aberto (Ctrl+C para desligar).");
  console.log("");
});
