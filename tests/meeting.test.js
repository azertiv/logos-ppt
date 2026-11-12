"use strict";
// Meeting notes for the Syllens extension: HTTP boundary, contract and Codex turn.
const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { CodexRpc } = require("../bridge/codex-rpc");
const { CodexService } = require("../bridge/codex-service");
const { createBridge } = require("../bridge/server");
const meetingTasks = require("../bridge/meeting-tasks");

const SYLLENS = "chrome-extension://ncoilhmgdphmchbcdlajbjnfkbmbfogk";
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

async function fixture(t) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "pictos-meeting-"));
  const trace = path.join(dir, "trace.jsonl");
  const rpc = new CodexRpc({ executable: process.execPath, args: [path.join(__dirname, "fixtures/fake-codex.js")], cwd: dir, env: { ...process.env, PICTOS_TEST_ACCOUNT: "chatgpt", PICTOS_TEST_TRACE: trace }, timeoutMs: 1000 });
  const service = new CodexService(rpc, { cwd: dir, timeoutMs: 3000, meetingTimeoutMs: 3000 });
  const bridge = createBridge({ service, port: 0 });
  const url = await bridge.listen();
  t.after(async () => { await bridge.close(); await delay(30); fs.rmSync(dir, { recursive: true, force: true }); });
  const post = (route, body, headers = {}) => fetch(url + route, {
    method: "POST", body: typeof body === "string" ? body : JSON.stringify(body),
    headers: { Authorization: `Bearer ${bridge.token}`, "Content-Type": "application/json", Origin: SYLLENS, ...headers }
  });
  return { url, bridge, service, post, trace: () => fs.existsSync(trace) ? fs.readFileSync(trace, "utf8").trim().split("\n").map(JSON.parse) : [] };
}

const meeting = { title: "Point hebdo", platform: "Google Meet", startedAt: "2026-11-12T13:00:00.000Z", participants: ["Alice", "Bob"] };
const summarizeBody = (extra = {}) => ({ language: "fr-FR", meeting, transcript: "[14:02:13] Alice: On valide le budget.\n[14:02:20] Bob: D'accord.", ...extra });

test("health announces meeting notes", async t => {
  const f = await fixture(t);
  const health = await (await fetch(f.url + "/health")).json();
  assert.equal(health.version, "1.2.0");
  assert.ok(health.capabilities.includes("meeting-notes"));
  assert.ok(health.capabilities.includes("pictos-search"));
});

test("only the Syllens extension origin is accepted among extensions", async t => {
  const f = await fixture(t);
  const preflight = await fetch(f.url + "/v1/meeting/summarize", { method: "OPTIONS", headers: { Origin: SYLLENS } });
  assert.equal(preflight.status, 204);
  assert.equal(preflight.headers.get("access-control-allow-origin"), SYLLENS);
  const other = await f.post("/v1/meeting/summarize", summarizeBody(), { Origin: "chrome-extension://aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa" });
  assert.equal(other.status, 403);
  const unpaired = await f.post("/v1/meeting/summarize", summarizeBody(), { Authorization: "Bearer " + "0".repeat(64) });
  assert.equal(unpaired.status, 401);
  assert.equal(f.trace().filter(m => m.method === "turn/start").length, 0);
});

test("rolling summary runs one ephemeral, tool-less, schema-constrained turn", async t => {
  const f = await fixture(t);
  const first = await (await f.post("/v1/meeting/summarize", summarizeBody({ context: "Budget Q1" }))).json();
  assert.equal(first.provider, "codex");
  assert.match(first.parsed.summary, /\(fr-FR\) \[14:02:13\] Alice: On valide le budget\./);
  assert.deepEqual(first.parsed.key_points.map(k => k.id), ["point-1"]);
  const start = f.trace().find(m => m.method === "thread/start").params;
  assert.equal(start.ephemeral, true);
  assert.equal(start.sandbox, "read-only");
  assert.equal(start.config["features.shell_tool"], false);
  assert.match(start.baseInstructions, /meeting-notes component of the Syllens/);
  assert.match(start.developerInstructions, /live notes of an ongoing meeting/);
  const turn = f.trace().find(m => m.method === "turn/start").params;
  assert.deepEqual(Object.keys(turn.outputSchema.properties), ["summary", "key_points"]);
  const prompt = JSON.parse(turn.input[0].text);
  assert.equal(prompt.task, "meeting-summary");
  assert.equal(prompt.context, "Budget Q1");

  // Next update: previous notes go back with the new lines only; ids are kept.
  const second = await (await f.post("/v1/meeting/summarize", summarizeBody({
    transcript: "[14:05:00] Bob: Prochaine étape, le planning.",
    previous: { summary: first.parsed.summary, keyPoints: first.parsed.key_points }
  }))).json();
  assert.deepEqual(second.parsed.key_points.map(k => k.id), ["point-1", "point-2"]);
});

test("key point expansion returns sub-points", async t => {
  const f = await fixture(t);
  const res = await (await f.post("/v1/meeting/expand", {
    language: "fr-FR", meeting, summary: "Budget validé.", keyPoint: { id: "budget", title: "Budget", detail: "Validé" },
    transcript: "[14:02:13] Alice: On valide le budget."
  })).json();
  assert.equal(res.parsed.sub_points[0].title, "Détail : Budget");
  assert.equal(JSON.parse(f.trace().find(m => m.method === "turn/start").params.input[0].text).task, "meeting-expand");
});

test("meeting input and output are bounded and validated", async t => {
  const f = await fixture(t);
  const missing = await f.post("/v1/meeting/summarize", { language: "fr", meeting });
  assert.equal(missing.status, 400);
  assert.match((await missing.json()).error, /Transcription manquant/);
  const tooLong = await f.post("/v1/meeting/summarize", summarizeBody({ transcript: "x".repeat(meetingTasks.LIMITS.transcript + 1) }));
  assert.equal(tooLong.status, 400);
  const tooBig = await f.post("/v1/meeting/summarize", summarizeBody({ transcript: "x".repeat(1100 * 1024) }));
  assert.equal(tooBig.status, 413);
  assert.equal(f.trace().filter(m => m.method === "turn/start").length, 0, "rejected requests never reach Codex");
  const bad = await f.post("/v1/meeting/summarize", summarizeBody({ transcript: "invalid-output" }));
  assert.equal(bad.status, 502);
  assert.match((await bad.json()).error, /inexploitable/);
  // Pictogram search keeps its own small limit.
  const search = await f.post("/v1/search", { task: "expand", query: "x".repeat(140000) });
  assert.equal(search.status, 413);
});

test("meeting and pictogram requests share the single Codex slot", async t => {
  const f = await fixture(t);
  const slow = f.post("/v1/search", { task: "expand", query: "slow" });
  await delay(150);
  const busy = await f.post("/v1/meeting/summarize", summarizeBody());
  assert.equal(busy.status, 409);
  assert.equal((await busy.json()).code, "BUSY");
  assert.equal((await slow).status, 200);
  assert.equal((await f.post("/v1/meeting/summarize", summarizeBody())).status, 200);
});

test("contract: unknown operations and malformed key points are refused", () => {
  assert.throws(() => meetingTasks.build("shell", summarizeBody()), /inconnue/);
  assert.throws(() => meetingTasks.build("summarize", summarizeBody({ previous: { keyPoints: [{ id: "a" }] } })), /Titre manquant/);
  assert.equal(JSON.parse(meetingTasks.build("summarize", summarizeBody({ language: "<script>" })).userPrompt).language, "en");
});
