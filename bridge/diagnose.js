"use strict";
const fs = require("node:fs");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const { runtimeOptions, createBridge } = require("./server");
const { CodexRpc } = require("./codex-rpc");
const { CodexService } = require("./codex-service");
const { CodexClient } = require("../public/ai-providers");

async function main() {
  const report = { date: new Date().toISOString(), platform: process.platform, architecture: process.arch, node: process.versions.node, inferenceRequested: process.argv.includes("--search"), checks: [] };
  const options = runtimeOptions();
  const version = spawnSync(options.executable, ["--version"], { env: options.env, cwd: options.cwd, shell: false, windowsHide: true, encoding: "utf8", timeout: 10000 });
  report.codex = version.status === 0 ? version.stdout.trim().split("\n").at(-1).slice(0, 100) : "indisponible";
  const service = new CodexService(new CodexRpc(options), options);
  const bridge = createBridge({ service, port: 0 });
  try {
    const url = await bridge.listen();
    if (!(await fetch(url + "/health")).ok) throw new Error("Le serveur local ne répond pas.");
    report.checks.push("Serveur local accessible");
    const client = new CodexClient({ url, token: bridge.token });
    const state = await client.status();
    report.checks.push("Protocole Codex opérationnel");
    report.chatgptConnected = state.connected;
    report.availableModels = state.models.map(m => m.id);
    if (report.inferenceRequested) {
      if (!state.connected) throw new Error("Ouvrez Demarrer.cmd, connectez ChatGPT dans le navigateur, puis relancez Test-Recherche.cmd.");
      const start = Date.now();
      const expansion = await client.search({ task: "expand", query: "ambition" });
      const p = expansion.parsed;
      const result = await client.search({ task: "rank", query: "ambition", expansion: { coreConcepts: p.core_concepts, visualMetaphors: p.visual_metaphors, concreteObjects: p.concrete_objects, relatedKeywords: p.related_keywords }, candidates: [{ id: 1, label: "Sommet", keywords: ["mountain", "climb"] }, { id: 2, label: "Fusée", keywords: ["rocket", "growth"] }, { id: 3, label: "Chien", keywords: ["dog", "animal"] }] }, expansion.model);
      report.search = { model: result.model, durationMs: Date.now() - start, orderedIds: result.parsed.ordered_ids, concepts: p };
      report.checks.push("Recherche réelle avec quota Codex réussie");
    }
    report.ok = true;
  } catch (error) { report.ok = false; report.error = error.message; process.exitCode = 1; }
  finally {
    await bridge.close();
    const destination = path.join(options.dataDir, "diagnostic.json");
    fs.writeFileSync(destination, JSON.stringify(report, null, 2) + "\n", { mode: 0o600 });
    console.log(JSON.stringify(report, null, 2));
    console.log(`\nDiagnostic enregistré : ${destination}\nLe rapport ne contient ni code de liaison, ni clé, ni adresse e-mail.`);
  }
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
