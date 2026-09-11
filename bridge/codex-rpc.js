"use strict";
const { spawn } = require("node:child_process");
const { EventEmitter } = require("node:events");

class BridgeError extends Error {
  constructor(message, code = "CODEX_ERROR", status = 502) { super(message); this.code = code; this.status = status; }
}

class CodexRpc extends EventEmitter {
  constructor({ executable, args = [], cwd, env, timeoutMs = 20000, spawnImpl = spawn }) {
    super();
    Object.assign(this, { executable, args, cwd, env, timeoutMs, spawnImpl });
    this.pending = new Map();
    this.nextId = 1;
    this.child = null;
    this.startPromise = null;
  }
  async start() {
    if (this.startPromise) return this.startPromise;
    this.startPromise = this.initialize().catch(error => { this.stop(); throw error; });
    return this.startPromise;
  }
  async initialize() {
    const child = this.spawnImpl(this.executable, this.args, {
      cwd: this.cwd, env: this.env, shell: false, windowsHide: true, stdio: ["pipe", "pipe", "pipe"]
    });
    this.child = child;
    let buffer = "";
    child.stdout.setEncoding("utf8");
    child.stdout.on("data", chunk => {
      buffer += chunk;
      if (Buffer.byteLength(buffer) > 4 * 1024 * 1024) { this.fail(new BridgeError("Réponse Codex trop volumineuse.")); return; }
      let index;
      while ((index = buffer.indexOf("\n")) >= 0) {
        const line = buffer.slice(0, index); buffer = buffer.slice(index + 1);
        if (!line.trim()) continue;
        let message;
        try { message = JSON.parse(line); } catch { this.fail(new BridgeError("Protocole Codex illisible.")); return; }
        this.receive(message);
      }
    });
    // Drain stderr; never log account credentials, prompts or filesystem details.
    child.stderr.on("data", () => {});
    child.stdin.on("error", () => this.fail(new BridgeError("Communication avec Codex interrompue.")));
    child.on("error", error => this.fail(new BridgeError(
      error.code === "ENOENT" ? "Codex introuvable. Vérifiez le dossier runtime ou PICTOS_CODEX_PATH." : "Impossible de lancer Codex. Vérifiez les restrictions du poste.", "CODEX_START_FAILED", 503
    )));
    child.on("exit", () => {
      if (this.child === child) this.fail(new BridgeError("Codex s’est arrêté. Réessayez pour le relancer.", "CODEX_STOPPED", 503));
    });
    const result = await this.request("initialize", { clientInfo: { name: "atelier_pictos", title: "Atelier Pictos", version: "0.2.0" }, capabilities: { experimentalApi: false } });
    this.write({ method: "initialized" });
    return result;
  }
  receive(message) {
    if (message.method && message.id !== undefined) {
      // This companion never authorizes tools, shell commands, file changes or permission escalation.
      if (["item/commandExecution/requestApproval", "item/fileChange/requestApproval"].includes(message.method)) this.write({ id: message.id, result: { decision: "cancel" } });
      else this.write({ id: message.id, error: { code: -32601, message: "Tools and permission requests are not available in Atelier Pictos." } });
      this.emit("toolDenied", message.params);
      return;
    }
    if (message.id !== undefined) {
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id); clearTimeout(pending.timer);
      if (message.error) pending.reject(new BridgeError(safeError(message.error.message)));
      else pending.resolve(message.result);
    } else if (message.method) this.emit("notification", message.method, message.params);
  }
  write(message) {
    if (!this.child || this.child.stdin.destroyed) throw new BridgeError("Codex n’est pas disponible.", "CODEX_STOPPED", 503);
    this.child.stdin.write(JSON.stringify(message) + "\n");
  }
  request(method, params = {}, timeoutMs = this.timeoutMs) {
    return new Promise((resolve, reject) => {
      const id = this.nextId++;
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new BridgeError("Codex ne répond pas. Réessayez.", "CODEX_TIMEOUT", 504));
      }, timeoutMs);
      this.pending.set(id, { resolve, reject, timer });
      try { this.write({ id, method, params }); } catch (error) { clearTimeout(timer); this.pending.delete(id); reject(error); }
    });
  }
  fail(error) {
    const child = this.child;
    this.child = null; this.startPromise = null;
    for (const { reject, timer } of this.pending.values()) { clearTimeout(timer); reject(error); }
    this.pending.clear();
    child?.kill();
    this.emit("stopped", error);
  }
  stop() { this.fail(new BridgeError("Compagnon arrêté.", "CODEX_STOPPED", 503)); }
}
function safeError(message) {
  const text = String(message || "").toLowerCase();
  if (/rate.?limit|quota|usage limit|credits/.test(text)) return "Limite Codex atteinte. Consultez les limites du compte puis réessayez après leur réinitialisation.";
  if (/unauthorized|authentication|401|login|sign in/.test(text)) return "Connexion ChatGPT requise ou expirée. Reconnectez le compagnon.";
  if (/model/.test(text)) return "Modèle Codex indisponible ou configuration incompatible. Actualisez la liste des modèles.";
  if (/sandbox|permission|denied/.test(text)) return "Codex est bloqué par les permissions ou la configuration du poste.";
  return "Codex n’a pas pu terminer cette opération. Vérifiez la connexion et la version du programme, puis réessayez.";
}
module.exports = { CodexRpc, BridgeError, safeError };
