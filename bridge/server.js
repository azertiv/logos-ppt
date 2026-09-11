"use strict";
const http = require("node:http");
const https = require("node:https");
const fs = require("node:fs");
const path = require("node:path");
const os = require("node:os");
const crypto = require("node:crypto");
const { spawn } = require("node:child_process");
const { CodexRpc, BridgeError } = require("./codex-rpc.js");
const { CodexService } = require("./codex-service.js");
const root = path.resolve(__dirname, "..");

function runtimeOptions(env = process.env) {
  const dataDir = path.resolve(env.PICTOS_DATA_DIR || path.join(env.LOCALAPPDATA || os.homedir(), "AtelierPictos"));
  const cwd = path.join(dataDir, "workspace");
  const codexHome = path.join(dataDir, "codex");
  fs.mkdirSync(cwd, { recursive: true, mode: 0o700 });
  fs.mkdirSync(codexHome, { recursive: true, mode: 0o700 });
  const executable = env.PICTOS_CODEX_PATH || (fs.existsSync(path.join(root, "runtime", "codex.exe")) ? path.join(root, "runtime", "codex.exe") : "codex");
  const childEnv = {};
  for (const key of ["PATH", "Path", "SystemRoot", "SYSTEMROOT", "WINDIR", "COMSPEC", "PATHEXT", "USERPROFILE", "APPDATA", "LOCALAPPDATA", "HOME", "TEMP", "TMP", "TMPDIR", "LANG", "LC_ALL", "HTTP_PROXY", "HTTPS_PROXY", "NO_PROXY", "http_proxy", "https_proxy", "no_proxy", "SSL_CERT_FILE", "SSL_CERT_DIR", "NODE_EXTRA_CA_CERTS"]) if (env[key]) childEnv[key] = env[key];
  childEnv.CODEX_HOME = codexHome;
  const config = {
    model_provider: "openai", forced_login_method: "chatgpt", approval_policy: "on-request", approvals_reviewer: "user", sandbox_mode: "read-only",
    "features.shell_tool": false, "features.unified_exec": false, "features.apply_patch_freeform": false,
    "features.multi_agent": false, "agents.enabled": false, "features.shell_snapshot": false, "features.code_mode": false,
    "apps._default.enabled": false, web_search: "disabled", project_doc_max_bytes: 0,
    "history.persistence": "none", "analytics.enabled": false, "windows.sandbox": "unelevated"
  };
  const args = ["app-server", "--listen", "stdio://"];
  for (const [key, value] of Object.entries(config)) args.push("-c", `${key}=${JSON.stringify(value)}`);
  return { executable, args, cwd, env: childEnv, dataDir };
}

function createBridge({ service, token = crypto.randomBytes(32).toString("hex"), port = 43129, allowedOrigins = [], tls } = {}) {
  const scheme = tls ? "https" : "http";
  const origins = new Set(["https://azertiv.github.io", "https://localhost:3000", ...allowedOrigins]);
  const publicFiles = new Map([
    ["/", [path.join(__dirname, "ui/index.html"), "text/html; charset=utf-8"]],
    ["/dashboard.js", [path.join(__dirname, "ui/dashboard.js"), "text/javascript; charset=utf-8"]],
    ["/dashboard.css", [path.join(__dirname, "ui/dashboard.css"), "text/css; charset=utf-8"]],
    ["/ai-providers.js", [path.join(root, "public/ai-providers.js"), "text/javascript; charset=utf-8"]]
  ]);
  let actualPort = port;
  const attempts = new Map();
  function json(res, status, body) { if (!res.destroyed) { res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" }); res.end(JSON.stringify(body)); } }
  async function handle(req, res) {
    res.setHeader("Cache-Control", "no-store"); res.setHeader("X-Content-Type-Options", "nosniff");
    res.setHeader("Referrer-Policy", "no-referrer"); res.setHeader("X-Frame-Options", "DENY");
    res.setHeader("Content-Security-Policy", "default-src 'none'; script-src 'self'; style-src 'self'; connect-src 'self'; img-src 'self'; base-uri 'none'; frame-ancestors 'none'; form-action 'none'");
    const validHosts = new Set([`127.0.0.1:${actualPort}`, `localhost:${actualPort}`]);
    if (!validHosts.has(req.headers.host)) return json(res, 403, { error: "Hôte non autorisé." });
    const origin = req.headers.origin;
    const localOrigins = new Set([`${scheme}://127.0.0.1:${actualPort}`, `${scheme}://localhost:${actualPort}`]);
    if (origin && !origins.has(origin) && !localOrigins.has(origin)) return json(res, 403, { error: "Origine non autorisée." });
    if (origin) { res.setHeader("Access-Control-Allow-Origin", origin); res.setHeader("Vary", "Origin"); }
    if (req.method === "OPTIONS") {
      if (!origin) return json(res, 403, { error: "Origine requise." });
      res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
      res.setHeader("Access-Control-Allow-Headers", "Authorization, Content-Type");
      res.setHeader("Access-Control-Allow-Private-Network", "true");
      res.writeHead(204); return res.end();
    }
    let pathname;
    try { pathname = new URL(req.url, `${scheme}://127.0.0.1:${actualPort}`).pathname; } catch { return json(res, 400, { error: "Adresse invalide." }); }
    if (req.method === "GET" && publicFiles.has(pathname)) {
      const [file, type] = publicFiles.get(pathname);
      try { const content = await fs.promises.readFile(file); res.writeHead(200, { "Content-Type": type }); return res.end(content); }
      catch { return json(res, 500, { error: "Fichier du compagnon manquant." }); }
    }
    if (req.method === "GET" && pathname === "/health") return json(res, 200, { application: "atelier-pictos", version: "0.2.0" });
    if (!pathname.startsWith("/v1/")) return json(res, 404, { error: "Adresse inconnue." });
    const supplied = Buffer.from(String(req.headers.authorization || ""));
    const expected = Buffer.from(`Bearer ${token}`);
    if (supplied.length !== expected.length || !crypto.timingSafeEqual(supplied, expected)) return json(res, 401, { error: "Code de liaison absent ou expiré. Copiez le code de la session actuelle du compagnon.", code: "PAIRING_REQUIRED" });
    const identity = origin || "local-cli";
    const now = Date.now();
    const recent = (attempts.get(identity) || []).filter(t => now - t < 60000);
    if (recent.length >= 120) return json(res, 429, { error: "Trop de demandes. Réessayez dans une minute.", code: "THROTTLED" });
    recent.push(now); attempts.set(identity, recent);
    const controller = new AbortController();
    res.on("close", () => { if (!res.writableEnded) controller.abort(); });
    try {
      if (pathname === "/v1/status" && req.method === "GET") return json(res, 200, await service.status());
      if (req.method !== "POST") return json(res, 405, { error: "Méthode non autorisée." });
      if (!/^application\/json(?:;|$)/i.test(req.headers["content-type"] || "")) return json(res, 415, { error: "Un contenu JSON est requis." });
      const body = await readBody(req);
      let result;
      switch (pathname) {
        case "/v1/login/start": result = await service.startLogin(body.deviceCode === true); break;
        case "/v1/login/cancel": result = await service.cancelLogin(); break;
        case "/v1/logout": result = await service.logout(); break;
        case "/v1/search": result = await service.search(body, controller.signal); break;
        default: return json(res, 404, { error: "Opération inconnue." });
      }
      json(res, 200, result);
    } catch (error) {
      json(res, error.status || 500, { error: error instanceof BridgeError ? error.message : "Le compagnon n’a pas pu traiter la demande.", code: error.code || "INTERNAL_ERROR" });
    }
  }
  const server = tls ? https.createServer(tls, handle) : http.createServer(handle);
  server.requestTimeout = 110000; server.headersTimeout = 10000;
  return {
    token, server,
    async listen() {
      await new Promise((resolve, reject) => { server.once("error", reject); server.listen(port, "127.0.0.1", () => { server.removeListener("error", reject); resolve(); }); });
      actualPort = server.address().port;
      return `${scheme}://127.0.0.1:${actualPort}`;
    },
    async close() { service.stop(); server.closeAllConnections(); if (server.listening) await new Promise(resolve => server.close(resolve)); }
  };
}
async function readBody(req) {
  let size = 0;
  const chunks = [];
  for await (const chunk of req) {
    size += chunk.length;
    if (size > 128 * 1024) throw new BridgeError("Demande trop volumineuse.", "BODY_TOO_LARGE", 413);
    chunks.push(chunk);
  }
  try { return JSON.parse(Buffer.concat(chunks).toString("utf8")); }
  catch { throw new BridgeError("JSON invalide.", "INVALID_JSON", 400); }
}
async function main() {
  if (Number(process.versions.node.split(".")[0]) < 20) throw new Error("Node.js 20 ou plus récent est requis.");
  const options = runtimeOptions();
  const port = Number(process.env.PICTOS_PORT || 43129);
  if (!Number.isInteger(port) || port < 1024 || port > 65535) throw new Error("Port PICTOS_PORT invalide.");
  let tls;
  if (process.env.PICTOS_TLS_CERT || process.env.PICTOS_TLS_KEY) tls = { cert: fs.readFileSync(process.env.PICTOS_TLS_CERT), key: fs.readFileSync(process.env.PICTOS_TLS_KEY) };
  const service = new CodexService(new CodexRpc(options), options);
  const bridge = createBridge({ service, port, tls, allowedOrigins: (process.env.PICTOS_ALLOWED_ORIGINS || "").split(",").map(s => s.trim()).filter(Boolean) });
  const url = await bridge.listen();
  console.log(`\nAtelier Pictos — compagnon Codex\n\nOuvrir : ${url}/#${bridge.token}\nCode de liaison : ${bridge.token}\n\nCette fenêtre doit rester ouverte. Ctrl+C pour arrêter.\nLa connexion ChatGPT est propre à ce compagnon. Aucun appel API payant automatique.\n`);
  if (process.argv.includes("--open")) {
    const fullUrl = `${url}/#${bridge.token}`;
    const command = process.platform === "win32" ? "explorer.exe" : process.platform === "darwin" ? "open" : "xdg-open";
    const browser = spawn(command, [fullUrl], { shell: false, stdio: "ignore", windowsHide: true });
    browser.on("error", () => console.log("Ouvrez manuellement le lien ci-dessus dans votre navigateur."));
    browser.unref();
  }
  let closing = false;
  const stop = async () => { if (closing) return; closing = true; await bridge.close(); };
  process.once("SIGINT", stop); process.once("SIGTERM", stop);
}
if (require.main === module) main().catch(error => {
  console.error(error.code === "EADDRINUSE" ? "Le compagnon est déjà ouvert ou le port 43129 est occupé. Fermez l’autre fenêtre ou définissez PICTOS_PORT." : error.message);
  process.exitCode = 1;
});
module.exports = { createBridge, runtimeOptions };
