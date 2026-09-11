(function (root, factory) {
  const value = factory();
  if (typeof module === "object" && module.exports) module.exports = value;
  else root.PictosAiProviders = value;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";
  const DEFAULT_URL = "http://127.0.0.1:43129";
  function localUrl(value) {
    let url;
    try { url = new URL(value); } catch { throw new Error("Adresse du compagnon invalide."); }
    if (!["http:", "https:"].includes(url.protocol) || !["127.0.0.1", "localhost"].includes(url.hostname) || url.username || url.password || url.search || url.hash || url.pathname !== "/") throw new Error("Le compagnon doit avoir une adresse locale (127.0.0.1 ou localhost).");
    return url.origin;
  }
  async function request(url, options, timeoutMs, fetchImpl) {
    const controller = new AbortController();
    const outer = options.signal;
    const abort = () => controller.abort();
    if (outer?.aborted) controller.abort();
    else outer?.addEventListener("abort", abort, { once: true });
    let timedOut = false;
    const timer = setTimeout(() => { timedOut = true; controller.abort(); }, timeoutMs);
    try {
      const response = await fetchImpl(url, { ...options, signal: controller.signal, cache: "no-store", credentials: "omit", redirect: "error" });
      const data = await response.json().catch(() => null);
      if (!response.ok) {
        const error = new Error(data?.error?.message || data?.error || `Service indisponible (${response.status}).`);
        error.code = data?.code || `HTTP_${response.status}`;
        throw error;
      }
      if (!data) throw new Error("Réponse du service illisible.");
      return data;
    } catch (error) {
      if (timedOut) throw new Error("Le service a mis trop de temps à répondre. Réessayez.");
      throw error;
    } finally {
      clearTimeout(timer);
      outer?.removeEventListener("abort", abort);
    }
  }
  class CodexClient {
    constructor({ url = DEFAULT_URL, token = "", fetchImpl = globalThis.fetch } = {}) {
      this.url = localUrl(url);
      this.token = token.trim();
      this.fetchImpl = fetchImpl;
    }
    async call(path, body, signal) {
      if (!/^[a-f0-9]{64}$/.test(this.token)) throw new Error("Copiez le code de liaison affiché par le compagnon local.");
      try {
        return await request(this.url + "/v1/" + path, {
          method: body === undefined ? "GET" : "POST",
          headers: { Authorization: `Bearer ${this.token}`, ...(body === undefined ? {} : { "Content-Type": "application/json" }) },
          ...(body === undefined ? {} : { body: JSON.stringify(body) }), signal
        }, path === "search" ? 100000 : 25000, this.fetchImpl);
      } catch (error) {
        if (error instanceof TypeError) throw new Error("Compagnon inaccessible. Lancez-le puis vérifiez son adresse. Le navigateur ou Office peut aussi bloquer l’accès au réseau local.");
        throw error;
      }
    }
    status(signal) { return this.call("status", undefined, signal); }
    search(input, model, signal) { return this.call("search", { ...input, model }, signal); }
  }
  async function apiJson({ task, model, key, user, signal, fetchImpl = globalThis.fetch }) {
    const data = await request("https://api.openai.com/v1/chat/completions", {
      method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${key}` }, signal,
      body: JSON.stringify({
        model, reasoning_effort: "none", max_completion_tokens: 1800, store: false, user,
        response_format: { type: "json_schema", json_schema: { name: task.schemaName, strict: true, schema: task.schema } },
        messages: [{ role: "system", content: task.systemPrompt }, { role: "user", content: task.userPrompt }]
      })
    }, 90000, fetchImpl);
    const content = data?.choices?.[0]?.message?.content;
    if (!content) throw new Error("Réponse IA vide.");
    return { parsed: JSON.parse(content), usage: data.usage || null };
  }
  function formatLimits(payload) {
    const buckets = payload?.rateLimitsByLimitId ? Object.entries(payload.rateLimitsByLimitId) : payload?.rateLimits ? [["codex", payload.rateLimits]] : [];
    const lines = [];
    for (const [id, bucket] of buckets) {
      for (const key of ["primary", "secondary"]) {
        const w = bucket?.[key];
        if (!w || typeof w.usedPercent !== "number" || !Number.isFinite(w.usedPercent)) continue;
        const mins = w.windowDurationMins;
        const duration = typeof mins === "number" ? (mins >= 1440 ? `${Math.round(mins / 1440)} j` : mins >= 60 ? `${Math.round(mins / 60)} h` : `${mins} min`) : key;
        const reset = typeof w.resetsAt === "number" && w.resetsAt > 0 ? ` · réinitialisation ${new Date(w.resetsAt * 1000).toLocaleString("fr-FR")}` : "";
        lines.push(`${bucket.limitName || id} — ${Math.max(0, Math.min(100, 100 - w.usedPercent)).toFixed(0)} % restants (${duration})${reset}`);
      }
    }
    return lines.length ? lines.join("\n") : "Limites non disponibles pour le moment.";
  }
  return { DEFAULT_URL, localUrl, CodexClient, apiJson, formatLimits };
});
