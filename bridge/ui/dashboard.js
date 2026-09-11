/* global PictosAiProviders */
"use strict";
const $ = id => document.getElementById(id);
let client, loginTimer, loginDeadline = 0, activeSearch, refreshing = false;
let token = location.hash.slice(1);
try { token ||= sessionStorage.getItem("pictosPairing") || ""; } catch {}
history.replaceState(null, "", location.pathname);
$("token").value = token;
$("address").textContent = location.origin;
const demoCandidates = [
  { id: 1, label: "Sommet", keywords: ["mountain", "peak", "climb"] },
  { id: 2, label: "Cible", keywords: ["target", "goal", "focus"] },
  { id: 3, label: "Fusée", keywords: ["rocket", "launch", "growth"] },
  { id: 4, label: "Trophée", keywords: ["trophy", "win", "achievement"] },
  { id: 5, label: "Équipe", keywords: ["team", "people", "together"] },
  { id: 6, label: "Ampoule", keywords: ["idea", "lightbulb", "innovation"] },
  { id: 7, label: "Feuille", keywords: ["leaf", "nature", "sustainability"] },
  { id: 8, label: "Bouclier", keywords: ["shield", "security", "protection"] },
  { id: 9, label: "Chien", keywords: ["dog", "animal", "pet"] },
  { id: 10, label: "Horloge", keywords: ["clock", "time", "deadline"] }
];
function notice(message, error = false) { $("notice").textContent = message; $("notice").classList.toggle("error", error); }
function pair() {
  token = $("token").value.trim();
  client = new PictosAiProviders.CodexClient({ url: location.origin, token });
  try { sessionStorage.setItem("pictosPairing", token); } catch {}
}
async function refresh() {
  if (refreshing) return;
  refreshing = true;
  try {
    if (!client) pair();
    const state = await client.status();
    $("account").textContent = state.connected ? `Connecté${state.account.email ? ` : ${state.account.email}` : ""}${state.account.planType ? ` · ${state.account.planType}` : ""}` : "Connectez un compte ChatGPT pour utiliser son quota Codex.";
    const previous = $("model").value;
    $("model").replaceChildren();
    for (const m of state.models) $("model").add(new Option(m.name || m.id, m.id));
    if (state.models.some(m => m.id === previous)) $("model").value = previous;
    else $("model").value = (state.models.find(m => /luna|mini|nano|spark/i.test(m.id)) || state.models.find(m => m.isDefault) || state.models[0])?.id || "";
    $("limits").textContent = PictosAiProviders.formatLimits(state.limits);
    $("search").disabled = !state.connected || Boolean(activeSearch);
    $("logout").disabled = !state.connected || Boolean(activeSearch);
    $("login").disabled = Boolean(activeSearch);
    if (state.login?.pending) showLogin(state.login);
    else {
      clearTimeout(loginTimer); $("login-flow").hidden = true;
      if (state.login?.error) notice(state.login.error, true);
      else if (!activeSearch) notice(state.connected ? "Compagnon prêt. Vous pouvez essayer une recherche." : "Compagnon accessible. La prochaine étape est la connexion ChatGPT.");
    }
  } catch (error) { notice(error.message, true); $("search").disabled = true; }
  finally { refreshing = false; }
}
function showLogin(result) {
  $("login-flow").hidden = false;
  $("auth-link").href = result.authUrl || result.verificationUrl;
  $("device-code").textContent = result.userCode || "";
  if (!loginDeadline) loginDeadline = Date.now() + 10 * 60 * 1000;
  clearTimeout(loginTimer);
  if (Date.now() < loginDeadline) loginTimer = setTimeout(refresh, 2500);
}
async function login(deviceCode) {
  if (!client) pair();
  notice("Préparation de la connexion…");
  const result = await client.call("login/start", { deviceCode });
  loginDeadline = Date.now() + 10 * 60 * 1000;
  showLogin(result);
  notice("Ouvrez le lien de connexion ci-dessous, puis revenez ici.");
}
function action(id, fn) { $(id).addEventListener("click", async () => { const b = $(id); b.disabled = true; try { await fn(); } catch (e) { notice(e.message, true); } finally { b.disabled = false; } }); }
action("pair", async () => { pair(); await refresh(); });
action("copy", async () => { await navigator.clipboard.writeText($("token").value.trim()); notice("Code copié. Collez-le dans les réglages du complément."); });
action("refresh", refresh);
action("login", () => login(false));
action("device-login", () => login(true));
action("cancel-login", async () => { await client.call("login/cancel", {}); loginDeadline = 0; await refresh(); });
action("logout", async () => { await client.call("logout", {}); await refresh(); });
$("cancel-search").addEventListener("click", () => activeSearch?.abort());
$("demo").addEventListener("submit", async event => {
  event.preventDefault();
  if (activeSearch) return;
  activeSearch = new AbortController();
  $("search").disabled = true; $("cancel-search").hidden = false;
  $("results").replaceChildren(); $("concepts").textContent = ""; $("timing").textContent = "";
  const start = performance.now();
  try {
    pair();
    const query = $("query").value.trim(), model = $("model").value, signal = activeSearch.signal;
    notice("Recherche des concepts…");
    const expanded = await client.search({ task: "expand", query }, model, signal);
    const p = expanded.parsed;
    const expansion = { coreConcepts: p.core_concepts, visualMetaphors: p.visual_metaphors, concreteObjects: p.concrete_objects, relatedKeywords: p.related_keywords };
    $("concepts").textContent = `Concepts proposés : ${[...expansion.visualMetaphors, ...expansion.concreteObjects].join(", ")}`;
    notice("Classement des pictogrammes…");
    const ranked = await client.search({ task: "rank", query, expansion, candidates: demoCandidates }, model, signal);
    for (const id of ranked.parsed.ordered_ids) {
      const candidate = demoCandidates.find(c => c.id === id);
      if (!candidate) continue;
      const li = document.createElement("li"); li.textContent = candidate.label; $("results").append(li);
    }
    $("timing").textContent = `${((performance.now() - start) / 1000).toFixed(1)} s · ${ranked.model} · deux opérations Codex`;
    notice("Recherche terminée. Le compagnon et Codex ont répondu.");
  } catch (error) { notice(error.name === "AbortError" ? "Recherche annulée." : error.message, error.name !== "AbortError"); }
  finally { activeSearch = null; $("search").disabled = false; $("cancel-search").hidden = true; }
});
if (token) refresh();
else { notice("Copiez le code de liaison depuis la fenêtre du compagnon."); $("search").disabled = true; }
