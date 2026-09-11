"use strict";
const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const http = require("node:http");
const { CodexRpc } = require("../bridge/codex-rpc");
const { CodexService } = require("../bridge/codex-service");
const { createBridge, runtimeOptions } = require("../bridge/server");
const { CodexClient } = require("../public/ai-providers");
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
async function fixture(t, { account = "chatgpt", timeoutMs = 3000 } = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "pictos-test-"));
  const trace = path.join(dir, "trace.jsonl");
  const rpc = new CodexRpc({ executable: process.execPath, args: [path.join(__dirname, "fixtures/fake-codex.js")], cwd: dir, env: { ...process.env, PICTOS_TEST_ACCOUNT: account, PICTOS_TEST_TRACE: trace }, timeoutMs: 1000 });
  const service = new CodexService(rpc, { cwd: dir, timeoutMs });
  const bridge = createBridge({ service, port: 0 });
  const url = await bridge.listen();
  t.after(async () => { await bridge.close(); await delay(30); fs.rmSync(dir, { recursive: true, force: true }); });
  return { rpc, service, bridge, url, client: new CodexClient({ url, token: bridge.token }), trace: () => fs.existsSync(trace) ? fs.readFileSync(trace, "utf8").trim().split("\n").map(JSON.parse) : [] };
}
const rank = query => ({ task: "rank", query, candidates: [{ id: 4, label: "Target" }, { id: 9, label: "Rocket" }] });
test("HTTP boundary: token, host, origin, preflight, body and route restrictions", async t => {
  const f = await fixture(t);
  assert.equal((await fetch(f.url + "/health")).status, 200);
  assert.equal((await fetch(f.url + "/v1/status")).status, 401);
  assert.equal((await fetch(f.url + "/v1/status", { headers: { Origin: "https://evil.invalid", Authorization: `Bearer ${f.bridge.token}` } })).status, 403);
  const response = await fetch(f.url + "/v1/search", { method: "OPTIONS", headers: { Origin: "https://azertiv.github.io", "Access-Control-Request-Private-Network": "true" } });
  assert.equal(response.status, 204); assert.equal(response.headers.get("access-control-allow-origin"), "https://azertiv.github.io"); assert.equal(response.headers.get("access-control-allow-private-network"), "true");
  const hostStatus = await new Promise(resolve => { http.get(f.url + "/", { headers: { Host: "evil.invalid" } }, res => { res.resume(); resolve(res.statusCode); }); });
  assert.equal(hostStatus, 403);
  assert.equal((await fetch(f.url + "/v1/search", { method: "POST", headers: { Authorization: `Bearer ${f.bridge.token}`, "Content-Type": "text/plain" }, body: "{}" })).status, 415);
  assert.equal((await fetch(f.url + "/v1/search", { method: "POST", headers: { Authorization: `Bearer ${f.bridge.token}`, "Content-Type": "application/json" }, body: "{" })).status, 400);
  const large = await fetch(f.url + "/v1/search", { method: "POST", headers: { Authorization: `Bearer ${f.bridge.token}`, "Content-Type": "application/json" }, body: JSON.stringify({ query: "x".repeat(140000) }) });
  assert.equal(large.status, 413);
  assert.equal((await fetch(f.url + "/codex/auth.json")).status, 404);
  assert.equal(f.trace().length, 0, "unauthorized requests never start Codex");
});
test("account, model and quota reads never start a model turn", async t => {
  const f = await fixture(t); const state = await f.client.status();
  assert.equal(state.connected, true); assert.equal(state.models[0].id, "test-model"); assert.equal(state.limits.rateLimits.primary.usedPercent, 12);
  assert.equal(f.trace().filter(m => m.method === "turn/start").length, 0);
  assert.deepEqual(f.trace().slice(0, 2).map(m => m.method), ["initialize", "initialized"]);
});
test("rank uses ephemeral read-only thread and tolerates final-before-reply ordering", async t => {
  const f = await fixture(t); const result = await f.client.search(rank("ambition"), "test-model");
  assert.deepEqual(result.parsed.ordered_ids, [9, 4]); assert.equal(result.provider, "codex"); assert.equal(result.usage, null);
  const start = f.trace().find(m => m.method === "thread/start").params;
  assert.equal(start.ephemeral, true); assert.equal(start.sandbox, "read-only"); assert.equal(start.config["features.shell_tool"], false);
  assert.equal(f.trace().find(m => m.method === "turn/start").params.outputSchema.type, "object");
});
test("subscription mode refuses API-key identity and anonymous identity", async t => {
  for (const account of ["apiKey", "none"]) { const f = await fixture(t, { account }); await assert.rejects(f.client.search(rank("ambition")), /Connectez votre compte ChatGPT/); assert.ok(!f.trace().some(m => m.method === "turn/start")); }
});
test("unknown tasks and hallucinated candidate IDs are rejected", async t => {
  const f = await fixture(t);
  await assert.rejects(f.client.search({ task: "shell", query: "hi" }), /inconnue/);
  await assert.rejects(f.client.search(rank("invalid")), /identifiants/);
  await assert.rejects(f.client.search(rank("ambition"), "not-in-catalog"), /Modèle non disponible/);
});
test("quota failure is explicit and never falls back to an API provider", async t => {
  const f = await fixture(t); await assert.rejects(f.client.search(rank("quota")), /Limite Codex atteinte/);
  assert.equal(f.trace().filter(m => m.method === "turn/start").length, 1);
});
function nextTurnStarted(rpc) {
  return new Promise((resolve, reject) => {
    const listener = method => {
      if (method !== "turn/started") return;
      clearTimeout(timer); rpc.off("notification", listener); resolve();
    };
    const timer = setTimeout(() => {
      rpc.off("notification", listener); reject(new Error("Fake Codex did not start a turn."));
    }, 5000);
    rpc.on("notification", listener);
  });
}
async function waitUntil(predicate) {
  const deadline = Date.now() + 5000;
  while (!predicate()) {
    if (Date.now() > deadline) throw new Error("Test condition did not become true.");
    await delay(10);
  }
}
test("cancel, timeout and busy guard stop or interrupt work", async t => {
  const f = await fixture(t, { timeoutMs: 500 });
  const started = nextTurnStarted(f.rpc);
  const a = f.client.search(rank("slow"));
  const rejection = assert.rejects(a, /trop longue/);
  await started;
  await assert.rejects(f.client.search(rank("ambition")), /déjà en cours/);
  await rejection;
  assert.ok(f.trace().some(m => m.method === "turn/interrupt"));
  f.service.timeoutMs = 3000;
  const controller = new AbortController();
  const restarted = nextTurnStarted(f.rpc);
  const b = f.client.search(rank("slow"), "test-model", controller.signal);
  const cancelled = assert.rejects(b, error => error.name === "AbortError");
  await restarted;
  controller.abort(); await cancelled;
  await waitUntil(() => !f.service.busy);
  assert.equal(f.trace().filter(m => m.method === "turn/interrupt").length, 2);
});
test("unexpected tool requests are declined, process crashes can recover", async t => {
  const f = await fixture(t);
  await assert.rejects(f.client.search(rank("tool")), /outil non autorisé/);
  await delay(50);
  assert.ok(f.trace().some(m => m.id === "approval-1" && m.result?.decision === "cancel"));
  await assert.rejects(f.client.search(rank("crash")), /arrêté/);
  assert.equal((await f.client.status()).connected, true);
});
test("managed ChatGPT login and logout operate only on companion account", async t => {
  const f = await fixture(t, { account: "none" });
  const result = await f.client.call("login/start", {}); assert.equal(result.authUrl, "https://auth.openai.com/test");
  await delay(60); assert.equal((await f.client.status()).connected, true);
  await f.client.call("logout", {}); assert.equal((await f.client.status()).connected, false);
});
test("runtime isolates Codex configuration, excludes API secrets and selects unelevated Windows sandbox", t => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "pictos-runtime-"));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const options = runtimeOptions({ PICTOS_DATA_DIR: dir, OPENAI_API_KEY: "not-a-real-key", CODEX_HOME: "/should-not-use", PATH: "/bin", PICTOS_CODEX_PATH: "/with spaces/codex" });
  assert.equal(options.executable, "/with spaces/codex"); assert.equal(options.env.OPENAI_API_KEY, undefined); assert.equal(options.env.CODEX_HOME, path.join(dir, "codex"));
  assert.ok(options.args.includes('windows.sandbox="unelevated"')); assert.ok(options.args.includes('forced_login_method="chatgpt"'));
});
