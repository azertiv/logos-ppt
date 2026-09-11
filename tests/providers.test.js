const { test } = require("node:test");
const assert = require("node:assert/strict");
const { apiJson, CodexClient, localUrl, formatLimits } = require("../public/ai-providers");
const tasks = require("../public/ai-tasks");
test("API keeps original endpoint, key, structured output and usage", async () => {
  const task = tasks.build({ task: "expand", query: "ambition" });
  const result = await apiJson({ task, model: "existing-model", key: "test-key", user: "test-user", fetchImpl: async (url, options) => {
    assert.equal(url, "https://api.openai.com/v1/chat/completions"); assert.equal(options.headers.Authorization, "Bearer test-key");
    const body = JSON.parse(options.body); assert.equal(body.model, "existing-model"); assert.equal(body.store, false); assert.equal(body.response_format.json_schema.strict, true);
    return new Response(JSON.stringify({ choices: [{ message: { content: '{"ok":true}' } }], usage: { prompt_tokens: 100 } }));
  } });
  assert.equal(result.parsed.ok, true); assert.equal(result.usage.prompt_tokens, 100);
});
test("Codex pairing credentials can only be sent to loopback and never follow redirects", async () => {
  for (const url of ["https://evil.invalid", "http://127.0.0.1.evil.invalid", "http://user@localhost", "http://localhost/path", "file:///tmp/x", "http://localhost#token"]) assert.throws(() => localUrl(url));
  const client = new CodexClient({ token: "a".repeat(64), fetchImpl: async (url, opts) => { assert.equal(opts.redirect, "error"); assert.equal(opts.credentials, "omit"); assert.equal(opts.headers.Authorization, `Bearer ${"a".repeat(64)}`); return new Response("{}"); } });
  await client.status();
  await assert.rejects(new CodexClient({ token: "sk-foo" }).status(), /code de liaison/);
});
test("limits: missing differs from zero, and distinct buckets are preserved", () => {
  assert.match(formatLimits({ rateLimits: { primary: { usedPercent: null } } }), /non disponibles/);
  assert.match(formatLimits({ rateLimits: { primary: { usedPercent: 0, windowDurationMins: 300 } } }), /100 % restants/);
  assert.match(formatLimits({ rateLimitsByLimitId: { first: { primary: { usedPercent: 100 } }, second: { primary: { usedPercent: 10 } } } }), /0 % restants.*\nsecond — 90 %/);
});
test("bounded task contract validates IDs, sizes and structured responses", () => {
  assert.throws(() => tasks.build({ task: "expand", query: "a".repeat(501) }));
  assert.throws(() => tasks.build({ task: "rank", query: "x", candidates: [{ id: 1, label: "x" }, { id: 1, label: "y" }] }));
  assert.throws(() => tasks.validate({ task: "expand" }, { core_concepts: "not an array" }));
  const input = { task: "rank", query: "test", candidates: [{ id: 4, label: "A" }] };
  assert.throws(() => tasks.validate(input, { ordered_ids: [5], note: "" }));
  assert.throws(() => tasks.validate(input, { ordered_ids: [4, 4], note: "" }));
  assert.deepEqual(tasks.validate(input, { ordered_ids: [], note: "No match" }), { ordered_ids: [], note: "No match" });
});
