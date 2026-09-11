/* global PictosAiProviders */
"use strict";
const $ = id => document.getElementById(id);
let client = null, token = "", state = null, timer, refreshing = false, acting = false;
let loginDeadline = 0;
const fragment = location.hash.slice(1);
history.replaceState(null, "", location.pathname);

function notice(message, error = false) { $("notice").textContent = message; $("notice").classList.toggle("error", error); }
function setToken(value) {
  token = /^[a-f0-9]{64}$/.test(value || "") ? value : "";
  client = token ? new PictosAiProviders.CodexClient({ url: location.origin, token }) : null;
  $("token").value = token;
  try { token ? sessionStorage.setItem("pictosPairing", token) : sessionStorage.removeItem("pictosPairing"); } catch {}
  syncButtons();
}
function syncButtons() {
  $("copy").disabled = !token || acting;
  $("login").disabled = !client || acting;
  $("device-login").disabled = !client || acting;
  $("cancel-login").disabled = acting;
  $("refresh").disabled = acting || refreshing;
  $("pair").disabled = acting;
  $("login").hidden = Boolean(state?.connected);
  $("logout").hidden = !state?.connected;
  $("logout").disabled = acting;
  $("device-login").hidden = Boolean(state?.connected);
  $("manual-pairing").hidden = Boolean(client);
}
function scheduleRefresh() {
  clearTimeout(timer);
  if (!client || document.hidden) return;
  const pending = state?.login?.pending && Date.now() < loginDeadline;
  timer = setTimeout(refresh, pending ? 2500 : 30000);
}
function showLogin(login) {
  const raw = login.authUrl || login.verificationUrl;
  let url;
  try { url = new URL(raw); } catch { throw new Error("Lien de connexion indisponible. Réessayez."); }
  if (url.protocol !== "https:" || !["auth.openai.com", "chatgpt.com", "auth.chatgpt.com"].includes(url.hostname)) throw new Error("Lien de connexion ChatGPT inattendu.");
  $("auth-link").href = url.href;
  $("device-code").textContent = login.userCode || "";
  $("login-flow").hidden = false;
  loginDeadline ||= Date.now() + 10 * 60 * 1000;
  return url.href;
}
async function refresh() {
  if (refreshing || acting) return;
  if (!client) { $("connection-title").textContent = "Ouvrez la liaison depuis l’icône du compagnon"; $("state-dot").dataset.state = "idle"; return; }
  refreshing = true; syncButtons();
  try {
    state = await client.connection();
    $("connection-title").textContent = state.connected ? "ChatGPT connecté" : state.login?.pending ? "Connexion ChatGPT en cours…" : "Connectez ChatGPT";
    $("state-dot").dataset.state = state.connected ? "ready" : "idle";
    $("account").textContent = state.connected ? (state.account?.email || "Le compagnon est prêt pour vos recherches dans PowerPoint.") : "Utilisez le compte ChatGPT de votre abonnement.";
    $("pairing-note").textContent = state.pairingPersistent ? "À faire une seule fois sur ce profil. Le code reste le même après redémarrage." : "La liaison sera mémorisée dans votre profil PowerPoint.";
    if (state.login?.pending) showLogin(state.login);
    else { $("login-flow").hidden = true; loginDeadline = 0; }
    notice(state.login?.error || "", Boolean(state.login?.error));
  } catch (error) {
    state = null; $("connection-title").textContent = "Connexion indisponible"; $("state-dot").dataset.state = "error";
    notice(error.message, true);
    if (error.code === "PAIRING_REQUIRED") { setToken(""); $("options").open = true; }
  } finally { refreshing = false; syncButtons(); scheduleRefresh(); }
}
async function login(deviceCode) {
  // Open only in direct response to the user's click; never on application startup.
  const authWindow = deviceCode ? null : window.open("about:blank", "_blank");
  if (authWindow) authWindow.opener = null;
  try {
    const login = await client.call("login/start", { deviceCode });
    state = { ...state, login }; loginDeadline = Date.now() + 10 * 60 * 1000;
    const url = showLogin(login);
    if (authWindow) authWindow.location.href = url;
    notice(deviceCode ? "Saisissez le code sur la page de connexion ChatGPT." : "Terminez la connexion dans la page ChatGPT.");
    $("connection-title").textContent = "Connexion ChatGPT en cours…";
  } catch (error) { authWindow?.close(); throw error; }
}
function action(id, callback) {
  $(id).addEventListener("click", async () => {
    if (acting) return;
    acting = true; syncButtons(); clearTimeout(timer);
    try { await callback(); } catch (error) { notice(error.message, true); }
    finally { acting = false; syncButtons(); scheduleRefresh(); }
  });
}
action("pair", async () => { setToken($("token").value.trim()); state = null; acting = false; await refresh(); });
action("copy", async () => { await navigator.clipboard.writeText(token); notice("Code copié. Collez-le dans les réglages PowerPoint."); });
action("refresh", async () => { acting = false; await refresh(); });
action("login", () => login(false));
action("device-login", () => login(true));
action("cancel-login", async () => { await client.call("login/cancel", {}); acting = false; await refresh(); });
action("logout", async () => { await client.call("logout", {}); acting = false; await refresh(); });
document.addEventListener("visibilitychange", () => { if (document.hidden) clearTimeout(timer); else refresh(); });
async function init() {
  try {
    if (/^link=[a-f0-9]{48}$/.test(fragment)) {
      const response = await fetch("/dashboard/session", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ ticket: fragment.slice(5) }), cache: "no-store", credentials: "omit", redirect: "error" });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Ouvrez à nouveau la liaison depuis l’icône du compagnon.");
      setToken(result.token);
    } else {
      let saved = ""; try { saved = sessionStorage.getItem("pictosPairing") || ""; } catch {}
      setToken(fragment || saved);
    }
    await refresh();
  } catch (error) { setToken(""); $("connection-title").textContent = "Ouvrez à nouveau la liaison depuis l’icône"; notice(error.message, true); }
}
init();
