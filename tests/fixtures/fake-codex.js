"use strict";
const readline = require("node:readline");
const fs = require("node:fs");
let account = process.env.PICTOS_TEST_ACCOUNT === "none" ? null : { type: process.env.PICTOS_TEST_ACCOUNT || "chatgpt", email: "test@example.invalid", planType: "test" };
const tasks = new Map();
const timers = new Map();
const send = msg => process.stdout.write(JSON.stringify(msg) + "\n");
let count = 0;
readline.createInterface({ input: process.stdin }).on("line", line => {
  const message = JSON.parse(line);
  if (process.env.PICTOS_TEST_TRACE) fs.appendFileSync(process.env.PICTOS_TEST_TRACE, line + "\n");
  if (!message.method || message.id === undefined) return;
  const { id, method, params: p } = message;
  const reply = result => send({ id, result });
  switch (method) {
    case "initialize": return reply({ userAgent: "pictos-test" });
    case "account/read": return reply({ account });
    case "account/rateLimits/read": return reply({ rateLimits: { primary: { usedPercent: 12, windowDurationMins: 300, resetsAt: 2000000000 } } });
    case "model/list": return reply({ data: [{ id: "test-model", model: "test-model", displayName: "Test model", hidden: false, isDefault: true, supportedReasoningEfforts: [{ reasoningEffort: "low" }], defaultReasoningEffort: "low" }], nextCursor: null });
    case "account/login/start": {
      reply({ type: "chatgpt", loginId: "test-login", authUrl: "https://auth.openai.com/test" });
      setTimeout(() => { account = { type: "chatgpt" }; send({ method: "account/login/completed", params: { success: true } }); }, 30); return;
    }
    case "account/login/cancel": return reply({ status: "cancelled" });
    case "account/logout": account = null; return reply({});
    case "thread/start": { const threadId = `thread-${++count}`; tasks.set(threadId, p); return reply({ thread: { id: threadId }, model: p.model }); }
    case "thread/unsubscribe": return reply({ status: "unsubscribed" });
    case "turn/interrupt": {
      clearTimeout(timers.get(p.threadId));
      reply({}); send({ method: "turn/completed", params: { threadId: p.threadId, turn: { id: p.turnId, status: "interrupted", items: [] } } }); return;
    }
    case "turn/start": {
      const input = JSON.parse(p.input[0].text), turnId = `turn-${p.threadId}`;
      if (input.query === "crash") process.exit(4);
      send({ method: "turn/started", params: { threadId: p.threadId, turn: { id: turnId } } });
      // Deliberately deliver final notification before the RPC reply to exercise the race.
      const finish = () => {
        if (input.query === "tool") { send({ id: "approval-1", method: "item/commandExecution/requestApproval", params: { threadId: p.threadId, turnId } }); reply({ turn: { id: turnId } }); return; }
        if (input.query === "quota") { reply({ turn: { id: turnId } }); send({ method: "turn/completed", params: { threadId: p.threadId, turn: { id: turnId, status: "failed", error: { message: "usage limit exceeded" }, items: [] } } }); return; }
        const output = input.candidates ? { ordered_ids: input.query === "invalid" ? [9999] : input.candidates.map(c => c.id).reverse().slice(0, 24), note: "" } : { core_concepts: ["ambition"], visual_metaphors: ["sommet"], concrete_objects: ["fusée"], related_keywords: ["cible"] };
        const item = { id: "answer", type: "agentMessage", phase: "final_answer", text: JSON.stringify(output) };
        send({ method: "item/completed", params: { threadId: p.threadId, turnId, item } });
        send({ method: "turn/completed", params: { threadId: p.threadId, turn: { id: turnId, status: "completed", items: [] } } });
        reply({ turn: { id: turnId } });
      };
      if (input.query === "slow") { reply({ turn: { id: turnId } }); timers.set(p.threadId, setTimeout(finish, 1000)); }
      else finish();
      return;
    }
    default: return send({ id, error: { message: "Unknown method" } });
  }
});
