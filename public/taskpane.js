/* global Office, JSZip, PictosAiTasks, PictosAiProviders */

const grid = document.getElementById("logo-grid");
const libraryScroll = document.getElementById("library-scroll");
const statusEl = document.getElementById("status");
const searchInput = document.getElementById("search-input");
const searchClear = document.getElementById("search-clear");
const refreshBtn = document.getElementById("refresh-btn");
const logoCount = document.getElementById("logo-count");
const settingsButton = document.getElementById("settings-button");
const settingsPanel = document.getElementById("settings-panel");
const keywordToggle = document.getElementById("keyword-toggle");
const densityRange = document.getElementById("density-range");
const densityValue = document.getElementById("density-value");
const sortSelect = document.getElementById("sort-select");
const zipArea = document.getElementById("zip-area");
const zipSummary = document.getElementById("zip-summary");
const zipToggle = document.getElementById("zip-toggle");
const zipStatusText = document.getElementById("zip-status-text");
const zipDrop = document.getElementById("zip-drop");
const zipInput = document.getElementById("zip-input");
const zipButton = document.getElementById("zip-button");
const zipMeta = document.getElementById("zip-meta");
const replaceToggle = document.getElementById("replace-toggle");
const aiToggle = document.getElementById("ai-toggle");
const aiApiKeyInput = document.getElementById("ai-api-key-input");
const aiApiSave = document.getElementById("ai-api-save");
const aiApiClear = document.getElementById("ai-api-clear");
const aiSettingsStatus = document.getElementById("ai-settings-status");
const aiUsageCalls = document.getElementById("ai-usage-calls");
const aiUsageInput = document.getElementById("ai-usage-input");
const aiUsageCached = document.getElementById("ai-usage-cached");
const aiUsageOutput = document.getElementById("ai-usage-output");
const aiUsageLastCost = document.getElementById("ai-usage-last-cost");
const aiUsageTotalCost = document.getElementById("ai-usage-total-cost");
const aiPricingNote = document.getElementById("ai-pricing-note");
const aiUsageReset = document.getElementById("ai-usage-reset");

let allLogos = [];
let keywordsMap = new Map();
let wordnetMap = new Map();
let keywordFilterState = "all";
let sortMode = "az";
let replaceSelectionEnabled = false;
let settingsPanelOpen = false;
let aiEnabled = false;
let aiApiKey = "";
let aiProvider = "api";
let codexReady = false;
let codexClient = null;
let codexModel = "";
let codexConnectionVersion = 0;
let codexConnectionState = "unpaired";
let codexConnectionMessage = "";
let codexCheckPromise = null;
let codexCheckController = null;
let codexConnectionTimer = null;
let codexPairingDirty = false;
let aiRequestController = null;
let localLogosCache = null;
let localZipRecord = null;
let keywordsPromise = null;
let wordnetPromise = null;
let logoById = new Map();
let tokenIndex = new Map();
let searchCache = new Map();
let searchTimer = null;
let renderFrame = null;
let lazyObserver = null;
let zipSession = null;
let zipPanelExpanded = true;
let zipPanelToggled = false;
let insertQueue = Promise.resolve();
let cachedSlideId = null;
let cachedSlideIdAt = 0;
const insertStateBySlide = new Map();
const localObjectUrls = new Set();
const favoriteSet = new Set();
const recentMap = new Map();
let aiSearchState = {
  query: "",
  resultIds: [],
  loading: false,
  error: "",
  requestId: 0,
  source: "local"
};
let aiSearchCache = new Map();
let gridLayoutKey = "";
let aiUsageSummary = {
  callCount: 0,
  inputTokens: 0,
  cachedInputTokens: 0,
  outputTokens: 0,
  totalCostUsd: 0,
  lastCostUsd: 0
};

const STORAGE_KEYS = {
  density: "logosPptDensity",
  keywordFilter: "logosPptKeywordFilter",
  sortMode: "logosPptSortMode",
  favorites: "logosPptFavorites",
  recents: "logosPptRecents",
  replaceSelection: "logosPptReplaceSelection",
  aiEnabled: "logosPptAiEnabled",
  aiApiKey: "logosPptAiApiKey",
  aiProvider: "logosPptAiProvider",
  codexUrl: "logosPptCodexUrl",
  codexModel: "logosPptCodexModel",
  codexToken: "logosPptCodexToken",
  aiSearchCache: "logosPptAiSearchCache",
  aiAnonId: "logosPptAiAnonId",
  aiUsage: "logosPptAiUsage"
};
const ZIP_CACHE_KEY = "logosPptZipCache";
const DB_NAME = "logosPptCache";
const DB_VERSION = 1;
const DB_STORE = "assets";
const TRANSPARENT_PIXEL =
  "data:image/gif;base64,R0lGODlhAQABAAAAACw=";
const MIN_SEARCH_PREFIX = 3;
const WORDNET_SYNONYMS_URL = "wordnet-synonyms.json";
const SYNONYM_LIMIT = 10;
const SCORE_DIRECT = 100;
const SCORE_SYNONYM = 12;
const SEARCH_DEBOUNCE_MS = 140;
const SEARCH_CACHE_LIMIT = 50;
let displayedLogos = [];
let displayedSearchKey = "";
let displayedMaxScore = 0;
let gridColumns = 3;
let gridWindowFrame = null;
let isComposing = false;
let aiPending = null;
let libraryBusy = false;
let libraryGeneration = 0;
const previewCache = new Map();
const PREVIEW_CACHE_LIMIT = 256;
const PREVIEW_CACHE_BYTES = 16 * 1024 * 1024;
let previewCacheBytes = 0;
let initialization;
let shortcutBusy = false;
let selectionActionRegistered = false;
let selectionActionError = "";
let shortcutCheckPromise = null;
const SLIDE_ID_CACHE_MS = 1000;
const INSERT_BASE_POSITION = { left: 48, top: 48 };
const INSERT_OFFSET_STEP = { x: 18, y: 18 };
const INSERT_OFFSET_STEPS = 8;
const INSERT_RESET_MS = 60000;
const RECENT_LIMIT = 80;
const AI_MODEL = "gpt-5.6-luna";
const AI_SEARCH_DEBOUNCE_MS = 2000;
const AI_FULL_SCAN_LIMIT = 180;
const AI_SCAN_FALLBACK_LIMIT = 320;
const AI_CANDIDATE_LIMIT = 72;
const AI_RESULT_LIMIT = 24;
const AI_CACHE_LIMIT = 24;
const AI_KEY_WARNING_MESSAGE = "Ajoutez une clé API OpenAI dans Réglages pour activer le mode AI.";
const AI_PRICING_PER_MILLION = {
  input: 0.2,
  cachedInput: 0.02,
  output: 1.2
};

registerSelectionAction();

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("Ouvrez cet add-in dans PowerPoint pour insérer les logos.", "error");
    return;
  }

  registerSelectionAction();
  initialization = init();
  initialization.catch((error) => {
    console.error(error);
    setStatus("Erreur d'initialisation de l'add-in.", "error");
  });
});

async function init() {
  restorePreferences();
  refreshBtn.addEventListener("click", () => loadLogos({ force: true }));
  initShortcutControls();
  document.getElementById("selection-insert").addEventListener("click", () => searchSelectedPictogram(undefined, { revealPane: false }));
  document.getElementById("search-form").addEventListener("submit", event => event.preventDefault());
  searchInput.addEventListener("input", () => {
    updateSearchClear();
    if (!isComposing) scheduleSearch();
  });
  searchInput.addEventListener("compositionstart", () => { isComposing = true; clearTimeout(searchTimer); clearAiSearchState(); });
  searchInput.addEventListener("compositionend", () => { isComposing = false; scheduleSearch(); });
  searchInput.addEventListener("keydown", event => {
    if (event.key === "Enter" && !event.isComposing && !isComposing) {
      event.preventDefault(); scheduleSearch({ immediate: true });
    }
  });
  searchInput.addEventListener("blur", () => {
    if (!isComposing && searchTimer) scheduleSearch({ immediate: true });
  });
  if (searchClear) {
    searchClear.addEventListener("click", () => {
      searchInput.value = "";
      updateSearchClear();
      scheduleSearch({ immediate: true });
      searchInput.focus();
    });
  }
  if (keywordToggle) {
    keywordToggle.addEventListener("change", () => {
      keywordFilterState = ["all", "with", "without"].includes(keywordToggle.value) ? keywordToggle.value : "all";
      persistKeywordFilter();
      syncKeywordToggle();
      scheduleSearch({ immediate: true });
    });
    syncKeywordToggle();
  }
  if (densityRange) {
    densityRange.addEventListener("input", () => {
      const value = Number(densityRange.value);
      updateGridColumns(value);
      persistDensity(value);
    });
    updateGridColumns(Number(densityRange.value));
  }
  if (sortSelect) {
    sortSelect.addEventListener("change", () => {
      const value = sortSelect.value;
      sortMode = isValidSortMode(value) ? value : "az";
      persistSortMode();
      clearSearchCache();
      scheduleSearch({ immediate: true });
    });
    syncSortSelect();
  }
  if (replaceToggle) {
    replaceToggle.addEventListener("change", () => {
      replaceSelectionEnabled = Boolean(replaceToggle.checked);
      persistReplaceSelection();
    });
  }
  if (grid) {
    grid.addEventListener("click", handleGridClick);
    grid.addEventListener("keydown", handleGridKeydown);
  }
  initSettingsPanel();
  initZipDropzone();
  initZipSummaryToggle();
  initAiControls();
  libraryScroll.addEventListener("scroll", scheduleGridWindow, { passive: true });
  window.addEventListener("resize", scheduleGridWindow, { passive: true });
  if (typeof ResizeObserver !== "undefined") new ResizeObserver(scheduleGridWindow).observe(libraryScroll);

  updateSearchClear();
  syncAiToggle();
  syncAiSettingsStatus();
  syncAiUsageView();
  await loadLogos();
}

function restorePreferences() {
  const storedFilter = safeStorageGet(STORAGE_KEYS.keywordFilter);
  if (storedFilter && ["all", "with", "without"].includes(storedFilter)) {
    keywordFilterState = storedFilter;
  }
  const storedSort = safeStorageGet(STORAGE_KEYS.sortMode);
  if (storedSort && isValidSortMode(storedSort)) {
    sortMode = storedSort;
  }
  if (densityRange) {
    const storedDensity = Number.parseInt(
      safeStorageGet(STORAGE_KEYS.density) || "",
      10
    );
    if (Number.isFinite(storedDensity)) {
      const clamped = Math.min(6, Math.max(1, storedDensity));
      densityRange.value = String(clamped);
    }
  }
  if (sortSelect) {
    sortSelect.value = sortMode;
  }
  loadFavoritesFromStorage();
  loadRecentsFromStorage();
  const storedReplace = safeStorageGet(STORAGE_KEYS.replaceSelection);
  replaceSelectionEnabled = storedReplace === "1";
  if (replaceToggle) {
    replaceToggle.checked = replaceSelectionEnabled;
  }
  aiProvider = safeStorageGet(STORAGE_KEYS.aiProvider) === "codex" ? "codex" : "api";
  document.getElementById("ai-provider").value = aiProvider;
  document.getElementById("ai-codex-url").value = safeStorageGet(STORAGE_KEYS.codexUrl) || PictosAiProviders.DEFAULT_URL;
  codexModel = safeStorageGet(STORAGE_KEYS.codexModel) || "";
  let savedToken = safeStorageGet(STORAGE_KEYS.codexToken) || "";
  try { savedToken ||= sessionStorage.getItem(STORAGE_KEYS.codexToken) || ""; } catch {}
  document.getElementById("ai-codex-token").value = savedToken;
  if (savedToken) safeStorageSet(STORAGE_KEYS.codexToken, savedToken);
  const storedAiEnabled = safeStorageGet(STORAGE_KEYS.aiEnabled);
  aiEnabled = storedAiEnabled === "1";
  aiApiKey = safeStorageGet(STORAGE_KEYS.aiApiKey) || "";
  if (aiApiKeyInput) {
    aiApiKeyInput.value = aiApiKey;
  }
  loadAiCacheFromStorage();
  loadAiUsageFromStorage();
  clearAiStatusWarningIfConfigured();
}

function persistKeywordFilter() {
  safeStorageSet(STORAGE_KEYS.keywordFilter, keywordFilterState);
}

function persistDensity(value) {
  if (!Number.isFinite(value)) return;
  safeStorageSet(STORAGE_KEYS.density, String(value));
}

function persistSortMode() {
  safeStorageSet(STORAGE_KEYS.sortMode, sortMode);
}

function persistReplaceSelection() {
  safeStorageSet(
    STORAGE_KEYS.replaceSelection,
    replaceSelectionEnabled ? "1" : "0"
  );
}

function persistAiEnabled() {
  safeStorageSet(STORAGE_KEYS.aiEnabled, aiEnabled ? "1" : "0");
}

function persistAiApiKey() {
  if (!aiApiKey) {
    safeStorageRemove(STORAGE_KEYS.aiApiKey);
    return;
  }
  safeStorageSet(STORAGE_KEYS.aiApiKey, aiApiKey);
}

function safeStorageGet(key) {
  try {
    return window.localStorage ? window.localStorage.getItem(key) : null;
  } catch (error) {
    return null;
  }
}

function safeStorageSet(key, value) {
  try {
    if (!window.localStorage) return false;
    window.localStorage.setItem(key, value);
    return true;
  } catch (error) {
    return false;
  }
}

function safeStorageRemove(key) {
  try {
    if (!window.localStorage) return;
    window.localStorage.removeItem(key);
  } catch (error) {
    // Ignore storage errors.
  }
}

function initSettingsPanel() {
  if (!settingsButton || !settingsPanel) return;
  const tips = [...settingsPanel.querySelectorAll(".info-tip")];
  for (const tip of tips) {
    tip.addEventListener("toggle", () => { if (tip.open) for (const other of tips) if (other !== tip) other.open = false; });
  }
  document.addEventListener("click", event => { for (const tip of tips) if (!tip.contains(event.target)) tip.open = false; });
  document.getElementById("settings-back").addEventListener("click", () => setSettingsPanelOpen(false));
  settingsButton.addEventListener("click", (event) => {
    event.stopPropagation();
    setSettingsPanelOpen(!settingsPanelOpen);
  });
  document.addEventListener("click", (event) => {
    if (!settingsPanelOpen) return;
    if (
      settingsPanel.contains(event.target) ||
      settingsButton.contains(event.target)
    ) {
      return;
    }
    setSettingsPanelOpen(false);
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && settingsPanelOpen) {
      const tip = settingsPanel.querySelector(".info-tip[open]");
      if (tip) { tip.open = false; tip.querySelector("summary").focus(); }
      else setSettingsPanelOpen(false);
    }
  });
}

function setSettingsPanelOpen(isOpen) {
  settingsPanelOpen = Boolean(isOpen);
  document.querySelector(".search-dock").inert = settingsPanelOpen;
  libraryScroll.inert = settingsPanelOpen;
  document.querySelector(".header").inert = settingsPanelOpen;
  document.body.classList.toggle("settings-open", settingsPanelOpen);
  if (settingsButton) {
    settingsButton.setAttribute("aria-expanded", settingsPanelOpen ? "true" : "false");
  }
  if (settingsPanel) {
    settingsPanel.classList.toggle("hidden", !settingsPanelOpen);
    settingsPanel.setAttribute("aria-hidden", settingsPanelOpen ? "false" : "true");
    if (settingsPanelOpen) { document.getElementById("settings-content").scrollTop = 0; document.getElementById("settings-back").focus(); refreshShortcutStatus(); }
    else settingsButton?.focus();
  }
}

function isAiProviderReady() {
  return aiProvider === "codex" ? codexReady && Boolean(codexClient && codexModel) : Boolean(aiApiKey);
}

function providerWarning() {
  return aiProvider === "codex" ? codexConnectionMessage || "Reliez le compagnon Codex dans Réglages." : AI_KEY_WARNING_MESSAGE;
}

function syncProviderPanels() {
  const codex = aiProvider === "codex";
  document.getElementById("ai-api-settings").classList.toggle("hidden", codex);
  document.getElementById("ai-api-usage").classList.toggle("hidden", codex);
  document.getElementById("ai-codex-settings").classList.toggle("hidden", !codex);
  syncCodexConnectionBadge();
}

function initProviderControls() {
  const providerSelect = document.getElementById("ai-provider");
  const urlInput = document.getElementById("ai-codex-url");
  const tokenInput = document.getElementById("ai-codex-token");
  const modelSelect = document.getElementById("ai-codex-model");
  providerSelect.addEventListener("change", () => {
    clearAiSearchState();
    aiProvider = providerSelect.value === "codex" ? "codex" : "api";
    safeStorageSet(STORAGE_KEYS.aiProvider, aiProvider);
    syncProviderPanels();
    syncAiToggle();
    syncAiSettingsStatus();
    scheduleSearch({ immediate: true });
    if (aiProvider === "codex") checkCodexConnection({ silent: true, full: true });
    else { codexConnectionVersion++; clearTimeout(codexConnectionTimer); codexCheckController?.abort(); codexCheckPromise = null; document.getElementById("ai-codex-connect").disabled = false; }
  });
  for (const input of [urlInput, tokenInput]) input.addEventListener("input", () => {
    codexConnectionVersion++;
    codexCheckController?.abort(); codexCheckPromise = null;
    document.getElementById("ai-codex-connect").disabled = false;
    codexPairingDirty = true;
    clearTimeout(codexConnectionTimer);
    codexReady = false;
    if (aiProvider === "codex") clearAiSearchState();
    setCodexConnectionState("unpaired", "Enregistrez le code pour conserver la liaison avec ce compagnon.");
    syncAiToggle();
    syncAiSettingsStatus();
  });
  modelSelect.addEventListener("change", () => {
    clearAiSearchState();
    codexModel = modelSelect.value;
    safeStorageSet(STORAGE_KEYS.codexModel, codexModel);
    syncAiSettingsStatus();
    scheduleSearch({ immediate: true });
  });
  document.getElementById("ai-codex-connect").addEventListener("click", () => checkCodexConnection({ full: true }));
  document.getElementById("ai-codex-forget").addEventListener("click", () => {
    codexConnectionVersion++; codexCheckController?.abort(); codexCheckPromise = null;
    document.getElementById("ai-codex-connect").disabled = false;
    clearTimeout(codexConnectionTimer); codexReady = false; codexClient = null; codexPairingDirty = false;
    tokenInput.value = ""; safeStorageRemove(STORAGE_KEYS.codexToken);
    try { sessionStorage.removeItem(STORAGE_KEYS.codexToken); } catch {}
    if (aiProvider === "codex") clearAiSearchState();
    setCodexConnectionState("unpaired", "Liaison oubliée sur ce profil PowerPoint."); syncAiToggle();
  });
  document.getElementById("ai-codex-open").addEventListener("click", async () => {
    let target;
    try {
      const url = PictosAiProviders.localUrl(urlInput.value.trim());
      const token = tokenInput.value.trim();
      target = window.open("about:blank", "_blank");
      if (target) target.opener = null;
      let destination = `${url}/`;
      if (/^[a-f0-9]{64}$/.test(token)) {
        try {
          const result = await new PictosAiProviders.CodexClient({ url, token }).call("dashboard-ticket", {});
          if (!/^[a-f0-9]{48}$/.test(result.ticket || "")) throw new Error("Lien du compagnon invalide.");
          destination += `#link=${result.ticket}`;
        } catch (error) {
          if (error.code !== "HTTP_404") throw error;
          destination += `#${token}`; // Compatibility with the previous console companion.
        }
      }
      if (target) target.location.href = destination;
      else syncAiSettingsStatus("Ouvrez la liaison depuis l’icône du compagnon près de l’horloge Windows.");
    } catch (error) { target?.close(); syncAiSettingsStatus(error.message); }
  });
  syncProviderPanels();
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) clearTimeout(codexConnectionTimer);
    else if (aiProvider === "codex" && !codexPairingDirty) checkCodexConnection({ silent: true });
  });
  window.addEventListener("focus", () => {
    if (aiProvider === "codex" && !codexPairingDirty) checkCodexConnection({ silent: true });
  });
  if (aiProvider === "codex" && tokenInput.value) checkCodexConnection({ silent: true, full: true });
}

function setCodexConnectionState(state, message = "") {
  codexConnectionState = state;
  codexConnectionMessage = message;
  if (state !== "ready") codexReady = false;
  syncCodexConnectionBadge();
  syncAiSettingsStatus();
}

function syncCodexConnectionBadge() {
  const badge = document.getElementById("codex-connection-badge");
  if (!badge) return;
  badge.classList.toggle("hidden", aiProvider !== "codex");
  badge.dataset.state = codexConnectionState;
  badge.textContent = ({ ready: "Codex connecté", checking: "Connexion…", offline: "Compagnon indisponible", login: "ChatGPT à connecter", unpaired: "Liaison à configurer", error: "Connexion à vérifier" })[codexConnectionState] || "Connexion à vérifier";
  badge.title = codexConnectionMessage || badge.textContent;
}

function scheduleCodexConnectionCheck() {
  clearTimeout(codexConnectionTimer);
  if (document.hidden || aiProvider !== "codex" || codexPairingDirty || codexConnectionState === "unpaired") return;
  codexConnectionTimer = setTimeout(() => checkCodexConnection({ silent: true }), codexConnectionState === "login" ? 5000 : 30000);
}

function checkCodexConnection({ silent = false, full = false } = {}) {
  if (silent && codexPairingDirty) return Promise.resolve();
  if (codexCheckPromise) return codexCheckPromise;
  const version = ++codexConnectionVersion;
  const controller = new AbortController(); codexCheckController = controller;
  const promise = performCodexConnectionCheck({ silent, full, version, controller });
  codexCheckPromise = promise;
  promise.finally(() => { if (codexCheckPromise === promise) codexCheckPromise = null; });
  return promise;
}

async function performCodexConnectionCheck({ silent, full, version, controller }) {
  const button = document.getElementById("ai-codex-connect");
  if (!silent || !codexReady) { button.disabled = true; setCodexConnectionState("checking", "Vérification du compagnon et de ChatGPT…"); }
  try {
    const url = document.getElementById("ai-codex-url").value.trim();
    const token = document.getElementById("ai-codex-token").value.trim();
    const client = new PictosAiProviders.CodexClient({ url, token });
    let state = full ? await client.status(controller.signal) : await client.connection(controller.signal);
    if (state.connected && !state.models && !codexReady) state = await client.status(controller.signal);
    if (version !== codexConnectionVersion) return;
    codexClient = client;
    codexPairingDirty = false;
    const urlSaved = safeStorageSet(STORAGE_KEYS.codexUrl, client.url);
    const tokenSaved = safeStorageSet(STORAGE_KEYS.codexToken, token);
    const remembered = urlSaved && tokenSaved;
    try { if (remembered) sessionStorage.removeItem(STORAGE_KEYS.codexToken); else sessionStorage.setItem(STORAGE_KEYS.codexToken, token); } catch {}
    if (state.models) {
      document.getElementById("ai-codex-limits").textContent = PictosAiProviders.formatLimits(state.limits);
      const select = document.getElementById("ai-codex-model"); select.replaceChildren();
      for (const model of state.models) select.add(new Option(model.name || model.id, model.id));
      codexModel = state.models.find(m => m.id === codexModel)?.id || state.models.find(m => /luna|mini|nano|spark/i.test(m.id))?.id || state.models.find(m => m.isDefault)?.id || state.models[0]?.id || "";
      select.value = codexModel; safeStorageSet(STORAGE_KEYS.codexModel, codexModel);
    }
    codexReady = Boolean(state.connected && codexModel);
    setCodexConnectionState(codexReady ? "ready" : state.connected ? "error" : "login", codexReady ? remembered ? `Codex connecté · ${codexModel}. Liaison mémorisée.` : "Codex connecté pour cette session. Office n’a pas pu mémoriser la liaison." : state.connected ? "Aucun modèle disponible. Actualisez la connexion." : "Compagnon ouvert. Connectez ChatGPT depuis son icône ; la liaison sera actualisée automatiquement.");
    clearAiStatusWarningIfConfigured();
  } catch (error) {
    if (version === codexConnectionVersion && error.name !== "AbortError") {
      const state = error.code === "PAIRING_REQUIRED" ? "unpaired" : error.code === "COMPANION_OFFLINE" || error.code === "SERVICE_TIMEOUT" ? "offline" : error.code === "LOGIN_REQUIRED" ? "login" : "error";
      setCodexConnectionState(state, error.message);
    }
  } finally {
    if (version === codexConnectionVersion) { button.disabled = false; syncAiToggle(); scheduleCodexConnectionCheck(); }
  }
}

function initAiControls() {
  initProviderControls();
  if (aiToggle) {
    aiToggle.addEventListener("click", handleAiToggle);
  }
  if (aiApiSave) {
    aiApiSave.addEventListener("click", () => {
      const nextValue = aiApiKeyInput ? aiApiKeyInput.value.trim() : "";
      if (aiProvider === "api") clearAiSearchState();
      aiApiKey = nextValue;
      persistAiApiKey();
      syncAiSettingsStatus(aiApiKey ? "Clé API enregistrée localement." : "Clé supprimée.");
      clearAiStatusWarningIfConfigured();
      syncAiToggle();
      scheduleSearch({ immediate: true });
    });
  }
  if (aiApiClear) {
    aiApiClear.addEventListener("click", () => {
      aiApiKey = "";
      if (aiApiKeyInput) {
        aiApiKeyInput.value = "";
      }
      persistAiApiKey();
      if (aiProvider === "api") clearAiSearchState({ keepCache: false });
      syncAiSettingsStatus("Clé API supprimée.");
      syncAiToggle();
      requestRender();
    });
  }
  if (aiUsageReset) {
    aiUsageReset.addEventListener("click", () => {
      aiUsageSummary = {
        callCount: 0,
        inputTokens: 0,
        cachedInputTokens: 0,
        outputTokens: 0,
        totalCostUsd: 0,
        lastCostUsd: 0
      };
      persistAiUsageSummary();
      syncAiUsageView();
    });
  }
}

function handleAiToggle() {
  const nextState = !aiEnabled;
  if (nextState && !isAiProviderReady()) {
    aiEnabled = false;
    persistAiEnabled();
    syncAiToggle();
    syncAiSettingsStatus(providerWarning());
    setSettingsPanelOpen(true);
    setStatus(providerWarning(), "error");
    return;
  }
  aiEnabled = nextState;
  persistAiEnabled();
  if (!aiEnabled) {
    clearAiSearchState();
  }
  clearAiStatusWarningIfConfigured();
  syncAiToggle();
  syncAiSettingsStatus();
  scheduleSearch({ immediate: true });
}

function syncAiToggle() {
  if (!aiToggle) return;
  aiToggle.classList.toggle("is-active", aiEnabled);
  aiToggle.classList.toggle("is-loading", aiSearchState.loading);
  aiToggle.classList.toggle("is-disabled", aiEnabled && !isAiProviderReady());
  aiToggle.setAttribute("aria-pressed", aiEnabled ? "true" : "false");
  aiToggle.setAttribute(
    "aria-label",
    aiEnabled ? "Désactiver le mode AI" : "Activer le mode AI"
  );
  aiToggle.title = aiEnabled ? "Mode AI activé" : "Mode AI désactivé";
}

function syncAiSettingsStatus(message = "") {
  if (!aiSettingsStatus) return;
  if (message) {
    aiSettingsStatus.textContent = message;
    return;
  }
  if (aiProvider === "codex") {
    aiSettingsStatus.textContent = codexConnectionMessage || (codexReady ? `Codex prêt · ${codexModel}.` : "Collez le code de liaison du compagnon pour connecter PowerPoint.");
    return;
  }
  if (!aiApiKey) {
    aiSettingsStatus.textContent = "Clé absente. Le mode AI restera désactivé.";
    return;
  }
  if (aiEnabled) {
    aiSettingsStatus.textContent = `Prêt · ${AI_MODEL}.`;
    return;
  }
  aiSettingsStatus.textContent = `Clé enregistrée · ${AI_MODEL}.`;
}

function clearAiSearchState(options = {}) {
  aiRequestController?.abort();
  aiRequestController = null;
  aiPending = null;
  const { keepCache = true } = options;
  aiSearchState = {
    query: "",
    resultIds: [],
    loading: false,
    error: "",
    requestId: aiSearchState.requestId + 1,
    source: "local"
  };
  if (!keepCache) {
    aiSearchCache = new Map();
    persistAiSearchCache();
  }
  syncAiToggle();
}

function loadAiCacheFromStorage() {
  aiSearchCache = new Map();
  const raw = safeStorageGet(STORAGE_KEYS.aiSearchCache);
  if (!raw) return;
  try {
    const entries = JSON.parse(raw);
    if (!Array.isArray(entries)) return;
    for (const entry of entries) {
      if (!entry || typeof entry.key !== "string" || !Array.isArray(entry.resultIds)) {
        continue;
      }
      aiSearchCache.set(entry.key, {
        resultIds: entry.resultIds.filter(Number.isFinite),
        source: entry.source || "cache",
        createdAt: Number(entry.createdAt) || Date.now()
      });
    }
  } catch (error) {
    aiSearchCache = new Map();
  }
}

function persistAiSearchCache() {
  const entries = Array.from(aiSearchCache.entries())
    .slice(-AI_CACHE_LIMIT)
    .map(([key, value]) => ({
      key,
      resultIds: Array.isArray(value.resultIds) ? value.resultIds.slice(0, AI_RESULT_LIMIT) : [],
      source: value.source || "cache",
      createdAt: Number(value.createdAt) || Date.now()
    }));
  safeStorageSet(STORAGE_KEYS.aiSearchCache, JSON.stringify(entries));
}

function saveAiCacheEntry(key, value) {
  if (!key || !value) return;
  aiSearchCache.set(key, {
    resultIds: Array.isArray(value.resultIds) ? value.resultIds.slice(0, AI_RESULT_LIMIT) : [],
    source: value.source || "ai",
    createdAt: Date.now()
  });
  while (aiSearchCache.size > AI_CACHE_LIMIT) {
    const oldestKey = aiSearchCache.keys().next().value;
    aiSearchCache.delete(oldestKey);
  }
  persistAiSearchCache();
}

function loadAiUsageFromStorage() {
  const raw = safeStorageGet(STORAGE_KEYS.aiUsage);
  if (!raw) return;
  try {
    const data = JSON.parse(raw);
    if (!data || typeof data !== "object") return;
    aiUsageSummary = {
      callCount: Number(data.callCount) || 0,
      inputTokens: Number(data.inputTokens) || 0,
      cachedInputTokens: Number(data.cachedInputTokens) || 0,
      outputTokens: Number(data.outputTokens) || 0,
      totalCostUsd: Number(data.totalCostUsd) || 0,
      lastCostUsd: Number(data.lastCostUsd) || 0
    };
  } catch (error) {
    // Ignore invalid persisted usage.
  }
}

function persistAiUsageSummary() {
  safeStorageSet(STORAGE_KEYS.aiUsage, JSON.stringify(aiUsageSummary));
}

function syncAiUsageView() {
  if (aiUsageCalls) aiUsageCalls.textContent = formatInteger(aiUsageSummary.callCount);
  if (aiUsageInput) aiUsageInput.textContent = formatInteger(aiUsageSummary.inputTokens);
  if (aiUsageCached) aiUsageCached.textContent = formatInteger(aiUsageSummary.cachedInputTokens);
  if (aiUsageOutput) aiUsageOutput.textContent = formatInteger(aiUsageSummary.outputTokens);
  if (aiUsageLastCost) aiUsageLastCost.textContent = formatUsd(aiUsageSummary.lastCostUsd);
  if (aiUsageTotalCost) aiUsageTotalCost.textContent = formatUsd(aiUsageSummary.totalCostUsd);
  if (aiPricingNote) {
    aiPricingNote.textContent =
      `Estimation ${AI_MODEL}: input $${AI_PRICING_PER_MILLION.input.toFixed(2)}/M, cached input ` +
      `$${AI_PRICING_PER_MILLION.cachedInput.toFixed(2)}/M, output $${AI_PRICING_PER_MILLION.output.toFixed(2)}/M.`;
  }
}

function recordAiUsage(usage) {
  if (!usage) return;
  const promptTokens = Number(usage.prompt_tokens) || 0;
  const cachedInputTokens = Number(usage.prompt_tokens_details?.cached_tokens) || 0;
  const outputTokens = Number(usage.completion_tokens) || 0;
  const billableInputTokens = Math.max(0, promptTokens - cachedInputTokens);
  const costUsd =
    (billableInputTokens * AI_PRICING_PER_MILLION.input) / 1_000_000 +
    (cachedInputTokens * AI_PRICING_PER_MILLION.cachedInput) / 1_000_000 +
    (outputTokens * AI_PRICING_PER_MILLION.output) / 1_000_000;

  aiUsageSummary.callCount += 1;
  aiUsageSummary.inputTokens += promptTokens;
  aiUsageSummary.cachedInputTokens += cachedInputTokens;
  aiUsageSummary.outputTokens += outputTokens;
  aiUsageSummary.lastCostUsd = costUsd;
  aiUsageSummary.totalCostUsd += costUsd;
  persistAiUsageSummary();
  syncAiUsageView();
}

function clearAiStatusWarningIfConfigured() {
  if (!isAiProviderReady() || !statusEl) return;
  if ([AI_KEY_WARNING_MESSAGE, "Vérifiez la connexion au compagnon Codex dans Réglages."].includes(statusEl.textContent)) {
    setStatus("");
  }
}

function formatInteger(value) {
  return new Intl.NumberFormat("fr-FR").format(Number(value) || 0);
}

function formatUsd(value) {
  const amount = Number(value) || 0;
  if (amount >= 1) {
    return `$${amount.toFixed(2)}`;
  }
  return `$${amount.toFixed(6)}`;
}

function isValidSortMode(value) {
  return ["az", "recent", "favorites"].includes(value);
}

function syncSortSelect() {
  if (!sortSelect) return;
  sortSelect.value = sortMode;
}

function getLogoStorageKey(logo) {
  return logo?.storageKey || logo?.entryName || logo?.name || "";
}

function loadFavoritesFromStorage() {
  favoriteSet.clear();
  const raw = safeStorageGet(STORAGE_KEYS.favorites);
  if (!raw) return;
  try {
    const items = JSON.parse(raw);
    if (Array.isArray(items)) {
      items.forEach((name) => {
        if (typeof name === "string" && name.trim()) {
          favoriteSet.add(name);
        }
      });
    }
  } catch (error) {
    // Ignore invalid favorites storage.
  }
}

function loadRecentsFromStorage() {
  recentMap.clear();
  const raw = safeStorageGet(STORAGE_KEYS.recents);
  if (!raw) return;
  try {
    const items = JSON.parse(raw);
    if (Array.isArray(items)) {
      items.forEach((entry) => {
        if (!entry) return;
        const key =
          typeof entry.key === "string" && entry.key.trim()
            ? entry.key
            : typeof entry.name === "string"
              ? entry.name
              : "";
        if (!key) return;
        const usedAt = Number(entry.usedAt) || 0;
        if (usedAt > 0) {
          recentMap.set(key, usedAt);
        }
      });
    }
  } catch (error) {
    // Ignore invalid recents storage.
  }
}

function persistFavorites() {
  safeStorageSet(
    STORAGE_KEYS.favorites,
    JSON.stringify(Array.from(favoriteSet))
  );
}

function persistRecents() {
  const entries = Array.from(recentMap.entries())
    .map(([key, usedAt]) => ({ key, usedAt }))
    .sort((a, b) => b.usedAt - a.usedAt)
    .slice(0, RECENT_LIMIT);
  safeStorageSet(STORAGE_KEYS.recents, JSON.stringify(entries));
}

function scheduleSearch(options = {}) {
  const { immediate = false } = options;
  if (searchTimer) {
    clearTimeout(searchTimer);
    searchTimer = null;
  }
  if (immediate) {
    void runSearchCycle();
    return;
  }
  const query = getNormalizedSearchQuery();
  if (aiSearchState.loading && aiSearchState.query !== query) clearAiSearchState();
  const delay = shouldUseAiSearch(query) ? AI_SEARCH_DEBOUNCE_MS : SEARCH_DEBOUNCE_MS;
  searchTimer = setTimeout(() => {
    searchTimer = null;
    void runSearchCycle();
  }, delay);
}

function requestRender() {
  if (renderFrame) {
    cancelAnimationFrame(renderFrame);
  }
  renderFrame = requestAnimationFrame(() => {
    renderFrame = null;
    renderLogos(getRenderableLogos());
  });
}

async function runSearchCycle() {
  const query = getNormalizedSearchQuery();
  if (!shouldUseAiSearch(query)) {
    clearAiSearchState();
    requestRender();
    return;
  }
  requestRender();
  await requestAiSearch(query);
}

function getNormalizedSearchQuery() {
  return normalizeSearchText(searchInput?.value.trim() || "");
}

function shouldUseAiSearch(query) {
  return Boolean(aiEnabled && isAiProviderReady() && query && query.length >= 2);
}

function getRenderableLogos() {
  const query = getNormalizedSearchQuery();
  if (
    shouldUseAiSearch(query) &&
    aiSearchState.query === query &&
    Array.isArray(aiSearchState.resultIds) &&
    aiSearchState.resultIds.length
  ) {
    return applyKeywordFilter(resolveAiResultIds(aiSearchState.resultIds));
  }
  return filterLogos();
}

function handleGridClick(event) {
  const favButton = event.target.closest(".favorite-toggle");
  if (favButton && grid.contains(favButton)) {
    event.preventDefault();
    event.stopPropagation();
    const logoId = Number(favButton.dataset.logoId);
    toggleFavoriteById(logoId, favButton);
    return;
  }
  const card = event.target.closest(".logo-card");
  if (!card || !grid.contains(card)) return;
  const logoId = Number(card.dataset.logoId);
  const logo = logoById.get(logoId);
  if (logo) {
    insertLogo(logo);
  }
}

function handleGridKeydown(event) {
  if (event.target.closest(".favorite-toggle")) return;
  const card = event.target.closest(".logo-card");
  if (!card || !grid.contains(card)) return;
  const index = Number(card.dataset.index);
  const steps = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: gridColumns, ArrowUp: -gridColumns };
  if (event.key in steps || event.key === "Home" || event.key === "End") {
    event.preventDefault();
    const next = event.key === "Home" ? 0 : event.key === "End" ? displayedLogos.length - 1
      : Math.min(displayedLogos.length - 1, Math.max(0, index + steps[event.key]));
    const rect = grid.getBoundingClientRect();
    const view = libraryScroll.getBoundingClientRect();
    const layout = PictosGrid.layout({ count: displayedLogos.length, columns: gridColumns, width: rect.width, top: view.top - rect.top, viewport: libraryScroll.clientHeight });
    const y = rect.top - view.top + Math.floor(next / gridColumns) * layout.stride;
    if (y < 0 || y + layout.size > libraryScroll.clientHeight) libraryScroll.scrollTo({ top: libraryScroll.scrollTop + y - 12, behavior: "instant" });
    renderGridWindow();
    grid.querySelector(`[data-index="${next}"]`)?.focus({ preventScroll: true });
    return;
  }
  if (event.key !== "Enter" && event.key !== " ") return;
  event.preventDefault();
  const logo = logoById.get(Number(card.dataset.logoId));
  if (logo) insertLogo(logo);
}

function initZipDropzone() {
  if (!zipDrop || !zipInput) return;

  zipDrop.addEventListener("click", () => zipInput.click());
  zipDrop.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      zipInput.click();
    }
  });
  if (zipButton) {
    zipButton.addEventListener("click", (event) => {
      event.stopPropagation();
      zipInput.click();
    });
  }
  zipInput.addEventListener("change", async () => {
    const [file] = zipInput.files || [];
    zipInput.value = "";
    if (file) {
      await handleZipFile(file);
    }
  });

  zipDrop.addEventListener("dragover", (event) => {
    event.preventDefault();
    zipDrop.classList.add("is-dragover");
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = "copy";
    }
  });
  zipDrop.addEventListener("dragleave", () => {
    zipDrop.classList.remove("is-dragover");
  });
  zipDrop.addEventListener("dragend", () => {
    zipDrop.classList.remove("is-dragover");
  });
  zipDrop.addEventListener("drop", async (event) => {
    event.preventDefault();
    zipDrop.classList.remove("is-dragover");
    const [file] = event.dataTransfer?.files || [];
    if (file) {
      await handleZipFile(file);
    }
  });
}

function initZipSummaryToggle() {
  if (!zipToggle) return;
  zipToggle.addEventListener("click", () => {
    const expanded = zipToggle.getAttribute("aria-expanded") === "true";
    setZipPanelExpanded(!expanded, { userInitiated: true });
  });
}

function setZipPanelExpanded(isExpanded, options = {}) {
  if (zipArea) {
    zipArea.classList.toggle("is-collapsed", !isExpanded);
  }
  if (zipToggle) {
    zipToggle.setAttribute("aria-expanded", isExpanded ? "true" : "false");
  }
  zipPanelExpanded = isExpanded;
  if (options.userInitiated) {
    zipPanelToggled = true;
  }
}

function setZipSummaryVisible(hasMeta) {
  if (!zipSummary) return;
  zipSummary.classList.toggle("hidden", !hasMeta);
}

async function loadLogos(options = {}) {
  if (libraryBusy) return;
  libraryBusy = true;
  refreshBtn.disabled = true;
  if (options.force) { keywordsPromise = null; wordnetPromise = null; }
  setStatus("Chargement des logos locaux…");
  try {
    if (!localLogosCache) {
      const record = localZipRecord?.buffer ? localZipRecord : await readZipCache();
      if (!record?.buffer) {
        renderEmptyState("Aucun ZIP local chargé. Ouvrez les réglages pour importer votre bibliothèque.");
        setStatus("");
        return;
      }
      const parsed = await loadZipBuffer(record.buffer);
      clearLogoCaches();
      localLogosCache = parsed.items;
      localZipRecord = { meta: buildZipMeta(record, parsed.items.length) };
      updateZipMeta(localZipRecord.meta);
    }
    await prepareLibrary();
  } catch (error) {
    console.error(error);
    setStatus("Impossible de charger la bibliothèque. Réessayez ou importez votre ZIP dans les réglages.", "error");
  } finally { libraryBusy = false; refreshBtn.disabled = false; }
}

async function prepareLibrary() {
  // Show filenames immediately; dictionary downloads must not delay the first preview.
  allLogos = attachKeywords(localLogosCache, keywordsMap);
  buildSearchIndex(allLogos);
  clearSearchCache();
  clearAiSearchState();
  requestRender();
  setStatus("");
  const [map, synonyms] = await Promise.all([getKeywordsMap(), getWordnetMap()]);
  keywordsMap = map;
  wordnetMap = synonyms;
  allLogos = attachKeywords(localLogosCache, keywordsMap);
  buildSearchIndex(allLogos);
  clearSearchCache();
  requestRender();
}

async function handleZipFile(file) {
  if (libraryBusy) { setStatus("Un chargement est déjà en cours."); return; }
  if (!file || !/\.zip$/i.test(file.name)) {
    setStatus("Merci de sélectionner un fichier .zip contenant des SVG.", "error"); return;
  }
  libraryBusy = true;
  refreshBtn.disabled = true;
  let candidate;
  setStatus(`Import du ZIP « ${file.name} »…`);
  try {
    const buffer = await file.arrayBuffer();
    candidate = createZipSession();
    const parsed = await candidate.load(buffer);
    if (!parsed.items.length) throw new Error("Aucun SVG trouvé dans le ZIP.");
    // Only replace the active archive once the candidate has been read successfully.
    resetZipSession({ terminate: true });
    zipSession = candidate; candidate = null;
    clearLogoCaches();
    localLogosCache = parsed.items;
    const meta = { name: file.name, size: file.size, count: parsed.items.length, updatedAt: Date.now() };
    localZipRecord = { meta };
    updateZipMeta(meta);
    const saved = saveZipCache(buffer, meta);
    await prepareLibrary();
    if (await saved) setStatus(buildZipStatsMessage(parsed.stats), "success");
    else setStatus("Bibliothèque chargée pour cette session. Le stockage local est indisponible : réimportez le ZIP à la prochaine ouverture.", "error");
  } catch (error) {
    setStatus(`Import impossible : ${error.message || error}`, "error");
  } finally { candidate?.terminate?.(); libraryBusy = false; refreshBtn.disabled = false; if (zipInput) zipInput.value = ""; }
}

function normalizeSynonymList(list) {
  const seen = new Set();
  const output = [];
  for (const item of list) {
    const normalized = normalizeSearchText(item);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    output.push(normalized);
  }
  return output;
}

function depluralizeToken(token) {
  if (!token || token.length < 4) return token;
  if (token.endsWith("ies") && token.length > 4) {
    return `${token.slice(0, -3)}y`;
  }
  if (token.endsWith("es") && token.length > 3) {
    return token.slice(0, -2);
  }
  if (token.endsWith("s") && token.length > 3) {
    return token.slice(0, -1);
  }
  return token;
}

function getSynonymsForToken(token) {
  if (!token || !wordnetMap || wordnetMap.size === 0) return [];
  const direct = wordnetMap.get(token);
  if (direct && direct.length) return normalizeSynonymList(direct);
  const singular = depluralizeToken(token);
  if (singular && singular !== token) {
    const fallback = wordnetMap.get(singular);
    if (fallback && fallback.length) return normalizeSynonymList(fallback);
  }
  return [];
}

function buildQueryGroups(tokens) {
  return tokens.map((token) => {
    const synonyms = getSynonymsForToken(token)
      .filter((item) => item && item !== token)
      .slice(0, SYNONYM_LIMIT);
    const terms = [token, ...synonyms].filter(Boolean);
    const termTokens = terms.map((term) => tokenizeSearchText(term));
    return { token, synonyms, terms, termTokens };
  });
}

function buildCandidateSetForGroup(group) {
  const union = new Set();
  for (const tokens of group.termTokens) {
    for (const term of tokens) {
      const set = tokenIndex.get(term);
      if (!set) continue;
      for (const id of set) {
        union.add(id);
      }
    }
  }
  return union;
}

function matchesGroup(searchText, group) {
  if (!searchText) return false;
  return group.terms.some((term) => searchText.includes(term));
}

function scoreLogo(logo, queryTokens, groups) {
  const text = logo.searchText || "";
  let directMatches = 0;
  let synonymMatches = 0;
  const matchedSynonyms = new Set();
  for (const token of queryTokens) {
    if (text.includes(token)) {
      directMatches += 1;
    }
  }
  for (const group of groups) {
    for (const synonym of group.synonyms) {
      if (matchedSynonyms.has(synonym)) continue;
      if (text.includes(synonym)) {
        synonymMatches += 1;
        matchedSynonyms.add(synonym);
      }
    }
  }
  return directMatches * SCORE_DIRECT + synonymMatches * SCORE_SYNONYM;
}

function compareBySortMode(a, b) {
  if (sortMode === "recent") {
    const delta = (b.lastUsedAt || 0) - (a.lastUsedAt || 0);
    if (delta !== 0) return delta;
    return a.name.localeCompare(b.name);
  }
  if (sortMode === "favorites") {
    const favDelta = (b.isFavorite ? 1 : 0) - (a.isFavorite ? 1 : 0);
    if (favDelta !== 0) return favDelta;
    return a.name.localeCompare(b.name);
  }
  return a.name.localeCompare(b.name);
}

function applyKeywordFilter(logos) {
  return logos.filter((logo) => {
    if (keywordFilterState === "with") return logo.hasKeywords;
    if (keywordFilterState === "without") return !logo.hasKeywords;
    return true;
  });
}

function filterLogos() {
  const query = normalizeSearchText(searchInput.value.trim());
  const cacheKey = `${keywordFilterState}|${sortMode}|${query}`;
  const cached = searchCache.get(cacheKey);
  if (cached) return cached;

  let candidates = allLogos;
  const queryTokens = query ? tokenizeSearchText(query) : [];
  const groups = queryTokens.length ? buildQueryGroups(queryTokens) : [];

  if (query) {
    if (queryTokens.length) {
      const sets = groups.map((group) => buildCandidateSetForGroup(group));
      if (sets.some((set) => set.size === 0)) {
        const empty = [];
        cacheSearchResult(cacheKey, empty);
        return empty;
      }
      const ids = intersectSets(sets);
      candidates = [];
      for (const id of ids) {
        const logo = logoById.get(id);
        if (logo) {
          candidates.push(logo);
        }
      }
      candidates = candidates.filter((logo) =>
        groups.every((group) => matchesGroup(logo.searchText || "", group))
      );
    } else {
      candidates = candidates.filter((logo) => (logo.searchText || "").includes(query));
    }
  }

  const filtered = applyKeywordFilter(candidates);

  let sorted = [];
  if (queryTokens.length) {
    sorted = filtered.map(logo => ({ ...logo }));
    for (const logo of sorted) {
      logo.relevanceScore = scoreLogo(logo, queryTokens, groups);
    }
    sorted.sort((a, b) => {
      const delta = (b.relevanceScore || 0) - (a.relevanceScore || 0);
      if (delta !== 0) return delta;
      return compareBySortMode(a, b);
    });
  } else {
    sorted = sortLogos(filtered).map(logo => ({ ...logo }));
    for (const logo of sorted) {
      logo.relevanceScore = 0;
    }
  }
  cacheSearchResult(cacheKey, sorted);
  return sorted;
}

function sortLogos(logos) {
  const list = logos.slice();
  if (sortMode === "recent") {
    list.sort((a, b) => {
      const delta = (b.lastUsedAt || 0) - (a.lastUsedAt || 0);
      if (delta !== 0) return delta;
      return a.name.localeCompare(b.name);
    });
    return list;
  }
  if (sortMode === "favorites") {
    list.sort((a, b) => {
      const favDelta = (b.isFavorite ? 1 : 0) - (a.isFavorite ? 1 : 0);
      if (favDelta !== 0) return favDelta;
      return a.name.localeCompare(b.name);
    });
    return list;
  }
  list.sort((a, b) => a.name.localeCompare(b.name));
  return list;
}

function requestAiSearch(query) {
  const key = buildAiSearchCacheKey(query);
  if (aiPending?.key === key && !aiRequestController?.signal.aborted) return aiPending.promise;
  const promise = performAiSearch(query);
  aiPending = { key, promise };
  promise.finally(() => { if (aiPending?.promise === promise) aiPending = null; });
  return promise;
}

async function performAiSearch(query) {
  if (!query || !aiEnabled || !isAiProviderReady() || !allLogos.length) {
    return;
  }
  aiRequestController?.abort();
  aiRequestController = new AbortController();
  const context = { provider: aiProvider, model: codexModel, client: codexClient, apiKey: aiApiKey, signal: aiRequestController.signal };
  const cacheKey = buildAiSearchCacheKey(query);
  const cached = aiSearchCache.get(cacheKey);
  if (cached && Array.isArray(cached.resultIds)) {
    aiSearchState = {
      query,
      resultIds: cached.resultIds,
      loading: false,
      error: "",
      requestId: aiSearchState.requestId + 1,
      source: cached.source || "cache"
    };
    syncAiToggle();
    requestRender();
    return;
  }

  const requestId = aiSearchState.requestId + 1;
  aiSearchState = {
    query,
    resultIds: aiSearchState.query === query ? aiSearchState.resultIds : [],
    loading: true,
    error: "",
    requestId,
    source: "pending"
  };
  syncAiToggle();

  try {
    const expansion =
      allLogos.length <= AI_FULL_SCAN_LIMIT
        ? createFallbackAiExpansion(query)
        : await expandQueryWithAi(query, context);
    if (requestId !== aiSearchState.requestId) {
      return;
    }
    const candidates = collectAiCandidates(query, expansion);
    const ranking = await rankAiCandidates(query, expansion, candidates, context);
    if (requestId !== aiSearchState.requestId) {
      return;
    }

    let resultIds = Array.isArray(ranking.orderedIds)
      ? ranking.orderedIds.filter(Number.isFinite)
      : [];
    resultIds = completeAiResultIds(resultIds, candidates);

    aiSearchState = {
      query,
      resultIds,
      loading: false,
      error: "",
      requestId,
      source: "ai"
    };
    saveAiCacheEntry(cacheKey, { resultIds, source: "ai" });
    setStatus("");
    syncAiSettingsStatus(
      resultIds.length
        ? `Mode AI actif sur ${context.provider === "codex" ? context.model : AI_MODEL}. ${resultIds.length} résultat${resultIds.length > 1 ? "s" : ""} reranké${resultIds.length > 1 ? "s" : ""}.`
        : `Mode AI actif sur ${context.provider === "codex" ? context.model : AI_MODEL}, sans résultat exploitable pour cette requête.`
    );
    syncAiToggle();
    requestRender();
  } catch (error) {
    if (requestId !== aiSearchState.requestId) {
      return;
    }
    if (context.provider === "codex" && ["COMPANION_OFFLINE", "SERVICE_TIMEOUT", "PAIRING_REQUIRED", "LOGIN_REQUIRED"].includes(error.code)) {
      setCodexConnectionState(error.code === "PAIRING_REQUIRED" ? "unpaired" : error.code === "LOGIN_REQUIRED" ? "login" : "offline", error.message);
      scheduleCodexConnectionCheck();
    }
    aiSearchState = {
      query,
      resultIds: [],
      loading: false,
      error: error?.message || String(error || "Erreur IA."),
      requestId,
      source: "error"
    };
    const message = `Erreur IA : ${aiSearchState.error}. La recherche locale reste disponible.`;
    syncAiSettingsStatus(message);
    setStatus(message, "error");
    syncAiToggle();
    requestRender();
  }
}

function createFallbackAiExpansion(query) {
  const tokens = tokenizeSearchText(query).slice(0, 6);
  return {
    coreConcepts: tokens,
    visualMetaphors: [],
    concreteObjects: tokens,
    relatedKeywords: tokens
  };
}

async function expandQueryWithAi(query, context) {
  const { parsed } = await callAiJson({ task: "expand", query }, context);
  return {
    coreConcepts: normalizeAiTerms(parsed.core_concepts),
    visualMetaphors: normalizeAiTerms(parsed.visual_metaphors),
    concreteObjects: normalizeAiTerms(parsed.concrete_objects),
    relatedKeywords: normalizeAiTerms(parsed.related_keywords)
  };
}

function collectAiCandidates(query, expansion) {
  if (allLogos.length <= AI_FULL_SCAN_LIMIT) {
    return allLogos.slice();
  }

  const buckets = [
    { terms: [query], weight: 160 },
    { terms: expansion.coreConcepts, weight: 120 },
    { terms: expansion.concreteObjects, weight: 92 },
    { terms: expansion.visualMetaphors, weight: 72 },
    { terms: expansion.relatedKeywords, weight: 56 }
  ];

  const phrases = buckets.flatMap(bucket => (bucket.terms || []).map(term => {
    const normalized = normalizeSearchText(term);
    return { normalized, tokens: tokenizeSearchText(normalized), weight: bucket.weight };
  }));
  const scored = [];
  for (const logo of allLogos) {
    const text = logo.searchText || "";
    let score = 0;
    for (const phrase of phrases) {
      score += phrase.weight * scorePreparedAiPhraseMatch(text, phrase.normalized, phrase.tokens);
    }
    if (score <= 0) continue;
    if (logo.isFavorite) score += 6;
    if (logo.lastUsedAt) score += 2;
    scored.push({ logo, score });
  }

  scored.sort((a, b) => {
    const delta = b.score - a.score;
    if (delta !== 0) return delta;
    return compareBySortMode(a.logo, b.logo);
  });

  if (!scored.length) {
    return sortLogos(allLogos).slice(0, Math.min(AI_CANDIDATE_LIMIT, AI_SCAN_FALLBACK_LIMIT));
  }

  return scored.slice(0, AI_CANDIDATE_LIMIT).map((entry) => entry.logo);
}

function scoreAiPhraseMatch(searchText, phrase) {
  const normalized = normalizeSearchText(phrase);
  if (!normalized || !searchText) return 0;
  const tokens = tokenizeSearchText(normalized);
  return scorePreparedAiPhraseMatch(searchText, normalized, tokens);
}

function scorePreparedAiPhraseMatch(searchText, normalized, tokens) {
  if (!tokens.length || !searchText) return 0;
  let matched = 0;
  for (const token of tokens) {
    if (searchText.includes(token)) {
      matched += 1;
    }
  }
  if (!matched) return 0;
  const ratio = matched / tokens.length;
  const phraseBonus = searchText.includes(normalized) ? 0.35 : 0;
  return ratio + phraseBonus;
}

async function rankAiCandidates(query, expansion, candidates, context) {
  if (!candidates.length) return { orderedIds: [] };
  if (candidates.length === 1) return { orderedIds: [candidates[0].id] };
  const input = {
    task: "rank", query, expansion,
    candidates: candidates.map(logo => ({ id: logo.id, label: stripSvgExtension(logo.name).slice(0, 300), keywords: Array.isArray(logo.keywords) ? logo.keywords.slice(0, 5).map(k => String(k).slice(0, 100)) : [] }))
  };
  const { parsed } = await callAiJson(input, context);
  return { orderedIds: parsed.ordered_ids, note: parsed.note || "" };
}

async function callAiJson(input, context) {
  const task = PictosAiTasks.build(input);
  let result;
  if (context.provider === "codex") {
    result = await context.client.search(input, context.model, context.signal);
  } else {
    result = await PictosAiProviders.apiJson({ task, model: AI_MODEL, key: context.apiKey, user: getOrCreateAiAnonId(), signal: context.signal });
    recordAiUsage(result.usage);
  }
  PictosAiTasks.validate(input, result.parsed);
  return result;
}

function normalizeAiTerms(list) {
  if (!Array.isArray(list)) return [];
  const seen = new Set();
  const output = [];
  for (const item of list) {
    const normalized = normalizeSearchText(item);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    output.push(normalized);
  }
  return output;
}

function buildAiSearchCacheKey(query) {
  return `${PictosAiTasks.VERSION}|${aiProvider}|${aiProvider === "codex" ? codexModel : AI_MODEL}|${getZipFingerprint()}|${query}`;
}

function getOrCreateAiAnonId() {
  const existing = safeStorageGet(STORAGE_KEYS.aiAnonId);
  if (existing) return existing;
  const next =
    typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
      ? `logos-ppt-${crypto.randomUUID()}`
      : `logos-ppt-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  safeStorageSet(STORAGE_KEYS.aiAnonId, next);
  return next;
}

function getZipFingerprint() {
  const meta = localZipRecord?.meta || {};
  return [
    meta.name || "zip",
    meta.updatedAt || 0,
    meta.size || 0,
    meta.count || allLogos.length
  ].join("|");
}

function resolveAiResultIds(ids) {
  const seen = new Set();
  const logos = [];
  for (const id of ids) {
    if (!Number.isFinite(id) || seen.has(id)) continue;
    const logo = logoById.get(id);
    if (!logo) continue;
    seen.add(id);
    logos.push(logo);
  }
  return logos;
}

function completeAiResultIds(ids, candidates) {
  const allowedIds = new Set(candidates.map(logo => logo.id));
  const seen = new Set();
  const output = [];
  for (const id of ids) {
    if (!Number.isFinite(id) || seen.has(id) || !allowedIds.has(id)) continue;
    const logo = logoById.get(id);
    if (!logo) continue;
    seen.add(id);
    output.push(id);
    if (output.length >= AI_RESULT_LIMIT) {
      return output;
    }
  }
  for (const logo of candidates) {
    if (!logo || !Number.isFinite(logo.id) || seen.has(logo.id)) continue;
    seen.add(logo.id);
    output.push(logo.id);
    if (output.length >= AI_RESULT_LIMIT) {
      break;
    }
  }
  return output;
}

function stripSvgExtension(value) {
  return String(value || "").replace(/\.svg$/i, "");
}

function attachKeywords(items, map) {
  return items.map((logo) => {
    const keywords = map.get(logo.name) || [];
    const searchLabel = [logo.name, logo.displayName, logo.entryName]
      .filter(Boolean)
      .join(" ");
    const searchText = buildSearchText(searchLabel, keywords);
    const storageKey = getLogoStorageKey(logo);
    const lastUsedAt = recentMap.get(storageKey) || recentMap.get(logo.name) || 0;
    const isFavorite = favoriteSet.has(storageKey) || favoriteSet.has(logo.name);
    return {
      ...logo,
      storageKey,
      keywords,
      hasKeywords: Array.isArray(keywords) && keywords.length > 0,
      searchText,
      lastUsedAt,
      isFavorite
    };
  });
}

function buildSearchIndex(logos) {
  logoById = new Map();
  tokenIndex = new Map();
  clearSearchCache();
  logos.forEach((logo, index) => {
    const id = Number.isFinite(logo.id) ? logo.id : index;
    logo.id = id;
    logoById.set(id, logo);
    const tokens = tokenizeSearchText(logo.searchText || "");
    const uniqueTokens = new Set(tokens);
    logo.searchTokens = tokens;
    logo.searchTokenSet = new Set(tokens);
    for (const token of uniqueTokens) {
      indexToken(token, id);
    }
  });
}

function indexToken(token, id) {
  if (!token) return;
  const maxLength = token.length;
  if (maxLength <= MIN_SEARCH_PREFIX) {
    addTokenToIndex(token, id);
    return;
  }
  for (let len = MIN_SEARCH_PREFIX; len <= maxLength; len += 1) {
    addTokenToIndex(token.slice(0, len), id);
  }
}

function addTokenToIndex(token, id) {
  if (!token) return;
  let set = tokenIndex.get(token);
  if (!set) {
    set = new Set();
    tokenIndex.set(token, set);
  }
  set.add(id);
}

function buildSearchText(name, keywords) {
  const safeName = name || "";
  const baseName = safeName.replace(/\.svg$/i, "");
  const parts = [safeName];
  if (baseName && baseName !== safeName) {
    parts.push(baseName);
  }
  if (Array.isArray(keywords)) {
    parts.push(...keywords);
  }
  return normalizeSearchText(parts.join(" "));
}

function normalizeSearchText(value) {
  const text = String(value || "").toLowerCase();
  if (!text) return "";
  try {
    return text
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  } catch (error) {
    return text.replace(/[^a-z0-9]+/g, " ").trim();
  }
}

function tokenizeSearchText(text) {
  if (!text) return [];
  return text.split(/\s+/).filter(Boolean);
}

function intersectSets(sets) {
  if (!sets.length) return new Set();
  const sorted = [...sets].sort((a, b) => a.size - b.size);
  const [first, ...rest] = sorted;
  const result = new Set();
  for (const value of first) {
    if (rest.every((set) => set.has(value))) {
      result.add(value);
    }
  }
  return result;
}

function cacheSearchResult(key, value) {
  searchCache.set(key, value);
  if (searchCache.size > SEARCH_CACHE_LIMIT) {
    clearSearchCache();
  }
}

function clearSearchCache() {
  searchCache.clear();
}

function renderEmptyState(message) {
  gridLayoutKey = "";
  displayedLogos = [];
  resetLazyObserver();
  grid.replaceChildren();
  grid.style.height = "auto";
  updateLogoCount(0);
  const empty = document.createElement("div");
  empty.className = "status";
  empty.textContent = message;
  grid.appendChild(empty);
}

function renderLogos(logos) {
  gridLayoutKey = "";
  const searchKey = `${libraryGeneration}|${getNormalizedSearchQuery()}|${keywordFilterState}|${sortMode}|${aiSearchState.source}`;
  if (searchKey !== displayedSearchKey) {
    displayedSearchKey = searchKey;
    libraryScroll.scrollTo({ top: 0, behavior: "instant" });
  }
  resetLazyObserver();
  grid.replaceChildren();
  displayedLogos = logos;
  displayedMaxScore = logos.reduce((max, logo) => Math.max(max, logo.relevanceScore || 0), 0);
  updateLogoCount(logos.length);
  if (!logos.length) { renderEmptyState("Aucun résultat pour cette recherche."); return; }
  renderGridWindow();
}

function scheduleGridWindow() {
  if (gridWindowFrame !== null) return;
  gridWindowFrame = requestAnimationFrame(() => { gridWindowFrame = null; renderGridWindow(); });
}

function renderGridWindow() {
  if (!displayedLogos.length) return;
  const rect = grid.getBoundingClientRect();
  if (!rect.width) return;
  const view = libraryScroll.getBoundingClientRect();
  const layout = PictosGrid.layout({ count: displayedLogos.length, columns: gridColumns,
    width: rect.width, top: view.top - rect.top, viewport: libraryScroll.clientHeight, overscanPixels: libraryScroll.clientHeight });
  const focusedCard = document.activeElement?.closest?.(".logo-card");
  const layoutKey = `${layout.start}|${layout.end}|${layout.size}|${layout.columns}|${layout.height}|${focusedCard?.dataset.index || ""}`;
  if (layoutKey === gridLayoutKey) return;
  gridLayoutKey = layoutKey;
  grid.style.height = `${layout.height}px`;
  const wanted = new Set();
  for (let index = layout.start; index < layout.end; index++) wanted.add(index);
  // Retain a focused card while scrolling so keyboard focus never vanishes.
  const focused = document.activeElement?.closest?.(".logo-card");
  if (focused && grid.contains(focused)) wanted.add(Number(focused.dataset.index));
  const existing = new Map();
  for (const card of Array.from(grid.children)) {
    const index = Number(card.dataset.index);
    if (!wanted.has(index)) { const img = card.querySelector("img"); if (img) lazyObserver?.unobserve(img); card.remove(); }
    else existing.set(index, card);
  }
  for (const index of wanted) {
    if (!displayedLogos[index]) continue;
    let card = existing.get(index);
    if (!card) { card = createLogoCard(displayedLogos[index], index, displayedMaxScore); grid.appendChild(card); }
    card.style.width = `${layout.size}px`;
    card.style.height = `${layout.size}px`;
    card.style.left = `${(index % layout.columns) * layout.stride}px`;
    card.style.top = `${Math.floor(index / layout.columns) * layout.stride}px`;
  }
  trimPreviewCache();
}

function createLogoCard(logo, index, maxScore) {
  const card = document.createElement("div");
  card.className = "logo-card";
  if (logo.isFavorite) {
    card.classList.add("is-favorite");
  }
  card.setAttribute("role", "button");
  card.setAttribute("tabindex", "0");
  card.setAttribute("aria-label", `Insérer ${logo.displayName || logo.name}`);
  card.dataset.index = String(index);
  card.title = logo.displayName || logo.name;
  card.dataset.logoId = String(logo.id ?? index);
  const score = Number(logo.relevanceScore) || 0;
  const normalized = maxScore > 0 ? Math.min(1, score / maxScore) : 0;
  const weighted = Math.pow(normalized, 1.5);
  const baseAlpha = maxScore > 0 ? 0.06 : 0.05;
  const alpha = baseAlpha + weighted * 0.5;
  const alphaStrong = baseAlpha + weighted * 0.7;
  card.style.setProperty("--relevance-alpha", alpha.toFixed(3));
  card.style.setProperty("--relevance-alpha-strong", alphaStrong.toFixed(3));

  const favButton = document.createElement("button");
  favButton.type = "button";
  favButton.className = "favorite-toggle";
  if (logo.isFavorite) {
    favButton.classList.add("is-active");
  }
  favButton.dataset.logoId = String(logo.id ?? index);
  favButton.setAttribute(
    "aria-label",
    logo.isFavorite ? "Retirer des favoris" : "Ajouter aux favoris"
  );
  favButton.setAttribute("aria-pressed", logo.isFavorite ? "true" : "false");
  favButton.textContent = "★";

  const preview = document.createElement("div");
  preview.className = "logo-preview";

  const img = document.createElement("img");
  img.loading = "lazy";
  img.decoding = "async";
  img.alt = logo.name;
  if (previewCache.get(logo.id)?.url) {
    img.src = previewCache.get(logo.id).url;
  } else {
    img.src = TRANSPARENT_PIXEL;
    img.dataset.logoId = String(logo.id ?? index);
    observeLazyImage(img);
  }

  preview.appendChild(img);
  card.appendChild(favButton);
  card.appendChild(preview);
  return card;
}

function observeLazyImage(img) {
  const observer = getLazyObserver();
  if (!observer) {
    loadLogoPreview(img);
    return;
  }
  observer.observe(img);
}

function getLazyObserver() {
  if (typeof IntersectionObserver === "undefined") {
    return null;
  }
  if (lazyObserver) {
    return lazyObserver;
  }
  lazyObserver = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const target = entry.target;
        lazyObserver.unobserve(target);
        loadLogoPreview(target);
      }
    },
    { root: libraryScroll, rootMargin: "600px", threshold: 0 }
  );
  return lazyObserver;
}

function resetLazyObserver() {
  if (lazyObserver) {
    lazyObserver.disconnect();
  }
}

function loadLogoPreview(img) {
  if (!img || img.dataset.loading === "true") return;
  const logoId = Number(img.dataset.logoId);
  if (!Number.isFinite(logoId)) return;
  const logo = logoById.get(logoId);
  if (!logo) return;
  img.dataset.loading = "true";
  ensureLogoUrl(logo)
    .then((url) => {
      if (!img.isConnected) return;
      img.src = url;
      img.removeAttribute("data-logo-id");
    })
    .catch((error) => {
      console.error(error);
    })
    .finally(() => {
      if (img.isConnected) {
        img.removeAttribute("data-loading");
      }
      trimPreviewCache();
    });
}

function toggleFavoriteById(logoId, button) {
  if (!Number.isFinite(logoId)) return;
  const logo = logoById.get(logoId);
  if (!logo || !logo.name) return;
  const nextState = !logo.isFavorite;
  const key = getLogoStorageKey(logo);
  logo.isFavorite = nextState;
  if (nextState) {
    favoriteSet.add(key);
  } else {
    favoriteSet.delete(key);
    favoriteSet.delete(logo.name);
  }
  persistFavorites();
  updateFavoriteButton(button, nextState);
  clearSearchCache();
  requestRender();
}

function updateFavoriteButton(button, isFavorite) {
  if (!button) return;
  button.classList.toggle("is-active", isFavorite);
  button.setAttribute("aria-pressed", isFavorite ? "true" : "false");
  button.setAttribute(
    "aria-label",
    isFavorite ? "Retirer des favoris" : "Ajouter aux favoris"
  );
}

function recordRecent(logo) {
  if (!logo || !logo.name) return;
  const usedAt = Date.now();
  recentMap.set(getLogoStorageKey(logo), usedAt);
  logo.lastUsedAt = usedAt;
  const canonical = logoById.get(logo.id);
  if (canonical) canonical.lastUsedAt = usedAt;
  persistRecents();
  clearSearchCache();
  if (sortMode === "recent") {
    requestRender();
  }
}

function cachedPreview(logo) {
  let entry = previewCache.get(logo.id);
  if (entry) previewCache.delete(logo.id);
  else entry = { bytes: 0, text: null, url: null, promise: null };
  previewCache.set(logo.id, entry);
  return entry;
}

function trimPreviewCache() {
  const pinned = new Set(Array.from(grid.children, card => Number(card.dataset.logoId)));
  for (const [id, entry] of previewCache) {
    if (previewCache.size <= PREVIEW_CACHE_LIMIT && previewCacheBytes <= PREVIEW_CACHE_BYTES) break;
    if (pinned.has(id) || entry.promise) continue;
    if (entry.url) { URL.revokeObjectURL(entry.url); localObjectUrls.delete(entry.url); }
    previewCacheBytes -= entry.bytes;
    previewCache.delete(id);
  }
}

async function ensureLogoUrl(logo) {
  const generation = libraryGeneration;
  const text = await getSvgText(logo);
  if (generation !== libraryGeneration) throw new Error("Bibliothèque remplacée.");
  const entry = cachedPreview(logo);
  if (!entry.url) entry.url = createSvgUrl(text);
  return entry.url;
}

async function getSvgText(logo) {
  const entry = cachedPreview(logo);
  if (entry.text) return entry.text;
  if (entry.promise) return entry.promise;
  const generation = libraryGeneration;
  entry.promise = fetchSvgTextFromZip(logo.entryName || logo.name).then(text => {
    if (generation !== libraryGeneration) throw new Error("Bibliothèque remplacée.");
    entry.text = text;
    entry.bytes = text.length * 2;
    previewCacheBytes += entry.bytes;
    return text;
  }).finally(() => { entry.promise = null; });
  return entry.promise;
}

async function fetchSvgTextFromZip(name) {
  const session = getZipSession();
  if (!session || !session.getSvg) {
    throw new Error("Lecteur ZIP indisponible.");
  }
  const result = await session.getSvg(name);
  if (!result || !result.svgText) {
    throw new Error("SVG introuvable en local.");
  }
  return result.svgText;
}

async function getPreparedSvg(logo) {
  const svgText = await getSvgText(logo);
  const normalized = normalizeSvg(svgText);
  return normalized;
}

async function getReplaceSelectionTarget() {
  if (!replaceSelectionEnabled) return null;
  if (typeof PowerPoint === "undefined" || typeof PowerPoint.run !== "function") {
    return null;
  }
  let info = null;
  try {
    await PowerPoint.run(async (context) => {
      if (typeof context.presentation.getSelectedShapes !== "function") {
        return;
      }
      const shapes = context.presentation.getSelectedShapes();
      const slides = context.presentation.getSelectedSlides();
      shapes.load("items");
      slides.load("items");
      await context.sync();
      if (!shapes.items || shapes.items.length !== 1) {
        return;
      }
      const slide = slides.items && slides.items[0] ? slides.items[0] : null;
      if (!slide) {
        return;
      }
      const shape = shapes.items[0];
      shape.load(["id", "left", "top", "width", "height", "type"]);
      shape.fill.load(["type", "foregroundColor"]);
      slide.load("id");
      await context.sync();
      const fillType = shape.fill?.type;
      const fillColor =
        fillType && String(fillType).toLowerCase() === "solid"
          ? shape.fill.foregroundColor
          : null;
      info = {
        shapeId: shape.id,
        slideId: slide.id,
        imageLeft: shape.left,
        imageTop: shape.top,
        imageWidth: shape.width,
        imageHeight: shape.height,
        fillColor
      };
    });
  } catch (error) {
    console.warn("Impossible de récupérer la sélection.", error);
    return null;
  }
  return info;
}

async function deleteShapeById(shapeId, slideId, fillColor) {
  if (!shapeId || !slideId) return;
  if (typeof PowerPoint === "undefined" || typeof PowerPoint.run !== "function") {
    return;
  }
  try {
    await PowerPoint.run(async (context) => {
      if (fillColor) {
        const selection = context.presentation.getSelectedShapes();
        selection.load("items/type,items/fill/type");
        await context.sync();
        const newShape = selection.items && selection.items[0];
        if (newShape) {
          const fillType = newShape.fill?.type;
          const isPictureFill =
            fillType && String(fillType).toLowerCase().includes("picture");
          const isGeometric =
            newShape.type &&
            String(newShape.type).toLowerCase().includes("geometric");
          if (!isPictureFill && isGeometric) {
            newShape.fill.setSolidColor(fillColor);
          }
        }
      }
      const slide = context.presentation.slides.getItem(slideId);
      const shape = slide.shapes.getItem(shapeId);
      shape.delete();
      await context.sync();
    });
  } catch (error) {
    console.warn("Impossible de supprimer la forme sélectionnée.", error);
  }
}

function insertLogo(logo) {
  if (!logo) return Promise.resolve();
  insertQueue = insertQueue
    .then(() => insertLogoNow(logo))
    .catch((error) => {
      console.error(error);
      setStatus(
        `Erreur d'insertion : ${error.message || error}. Vérifiez que le SVG est valide.`,
        "error"
      );
    });
  return insertQueue;
}

async function insertLogoNow(logo) {
  if (!Office.context.requirements.isSetSupported("ImageCoercion", "1.2")) {
    setStatus(
      "Votre version de PowerPoint ne supporte pas l'insertion SVG (ImageCoercion 1.2).",
      "error"
    );
    return;
  }
  setStatus(`Insertion de ${logo.name}…`);

  try {
    const svg = await getPreparedSvg(logo);
    const target = replaceSelectionEnabled
      ? await getReplaceSelectionTarget()
      : null;
    if (target) {
      await insertSvg(svg, target);
      await deleteShapeById(target.shapeId, target.slideId, target.fillColor);
    } else {
      const fallbackPosition = await getNextInsertPosition();
      await insertSvg(svg, fallbackPosition);
    }

    recordRecent(logo);
    setStatus(`Logo inséré : ${logo.name}`, "success");
  } catch (error) {
    console.error(error);
    setStatus(
      `Erreur d'insertion : ${error.message || error}. Vérifiez que le SVG est valide.`,
      "error"
    );
  }
}

function setSelectedData(data, options) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(data, options, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(result.error);
      }
    });
  });
}

async function insertSvg(svg, position = {}) {
  const left = Number.isFinite(position.imageLeft)
    ? position.imageLeft
    : INSERT_BASE_POSITION.left;
  const top = Number.isFinite(position.imageTop)
    ? position.imageTop
    : INSERT_BASE_POSITION.top;
  const width = Number.isFinite(position.imageWidth) ? position.imageWidth : null;
  const height = Number.isFinite(position.imageHeight) ? position.imageHeight : null;
  const options = {
    coercionType: Office.CoercionType.XmlSvg,
    imageLeft: left,
    imageTop: top
  };
  if (width) {
    options.imageWidth = width;
  }
  if (height) {
    options.imageHeight = height;
  }

  try {
    await setSelectedData(svg, options);
  } catch (error) {
    if (isSelectionError(error)) {
      await forceSlideSelection();
      await setSelectedData(svg, options);
      return;
    }
    throw error;
  }
}

async function getNextInsertPosition() {
  const slideId = await getCachedSlideId();
  const key = slideId || "default";
  const state = getInsertState(key);
  const position = {
    imageLeft: INSERT_BASE_POSITION.left + state.offsetIndex * INSERT_OFFSET_STEP.x,
    imageTop: INSERT_BASE_POSITION.top + state.offsetIndex * INSERT_OFFSET_STEP.y
  };
  state.offsetIndex = (state.offsetIndex + 1) % INSERT_OFFSET_STEPS;
  state.lastUsedAt = Date.now();
  insertStateBySlide.set(key, state);
  return position;
}

function getInsertState(key) {
  const now = Date.now();
  const existing = insertStateBySlide.get(key);
  if (!existing || now - existing.lastUsedAt > INSERT_RESET_MS) {
    return { offsetIndex: 0, lastUsedAt: now };
  }
  return existing;
}

async function getCachedSlideId() {
  const now = Date.now();
  if (cachedSlideId && now - cachedSlideIdAt < SLIDE_ID_CACHE_MS) {
    return cachedSlideId;
  }
  const slideId = await getSelectedSlideId();
  cachedSlideId = slideId || null;
  cachedSlideIdAt = now;
  return slideId;
}

function isSelectionError(error) {
  const message = (error && error.message) ? error.message : String(error || "");
  return (
    /current selection/i.test(message) ||
    /sélection actuelle/i.test(message) ||
    /selection/i.test(message)
  );
}

async function forceSlideSelection() {
  if (typeof PowerPoint === "undefined" || typeof PowerPoint.run !== "function") {
    return;
  }

  try {
    const slideId = await getSelectedSlideId();
    if (!slideId) {
      return;
    }
    await goToSlide(slideId);
  } catch (error) {
    // Ignore selection forcing errors and fall back to default behavior.
  }
}

async function getSelectedSlideId() {
  if (typeof PowerPoint === "undefined" || typeof PowerPoint.run !== "function") {
    return null;
  }
  let slideId = null;
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items");
    await context.sync();
    const slide = slides.items.length === 1 ? slides.items[0] : null;
    if (!slide) {
      return;
    }
    slide.load("id");
    await context.sync();
    slideId = slide.id;
  });
  cachedSlideId = slideId || null;
  cachedSlideIdAt = Date.now();
  return slideId;
}

function goToSlide(slideId) {
  return new Promise((resolve, reject) => {
    Office.context.document.goToByIdAsync(
      slideId,
      Office.GoToType.Slide,
      result => result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(result.error)
    );
  });
}

function getKeywordsMap() {
  if (!keywordsPromise) {
    keywordsPromise = fetchKeywords();
  }
  return keywordsPromise;
}

function getWordnetMap() {
  if (!wordnetPromise) {
    wordnetPromise = fetchWordnetSynonyms();
  }
  return wordnetPromise;
}

async function fetchKeywords() {
  try {
    const response = await fetch("keywords.json", { cache: "force-cache", signal: AbortSignal.timeout(12000) });
    if (!response.ok) {
      return new Map();
    }
    const data = await response.json();
    const items = Array.isArray(data.items) ? data.items : [];
    const map = new Map();
    for (const item of items) {
      if (item?.file) {
        map.set(item.file, Array.isArray(item.keywords) ? item.keywords : []);
      }
    }
    return map;
  } catch (error) {
    return new Map();
  }
}

async function fetchWordnetSynonyms() {
  try {
    const response = await fetch(WORDNET_SYNONYMS_URL, { cache: "force-cache", signal: AbortSignal.timeout(12000) });
    if (!response.ok) {
      return new Map();
    }
    const data = await response.json();
    const items = data?.items || data?.synonyms || {};
    const map = new Map();
    if (Array.isArray(items)) {
      for (const item of items) {
        if (!item) continue;
        const term = normalizeSearchText(item.term || item.word || "");
        const synonyms = Array.isArray(item.synonyms) ? item.synonyms : [];
        if (term && synonyms.length) {
          map.set(term, synonyms);
        }
      }
    } else if (items && typeof items === "object") {
      let processed = 0;
      for (const [term, synonyms] of Object.entries(items)) {
        if (++processed % 1024 === 0) await new Promise(resolve => setTimeout(resolve, 0));
        if (!Array.isArray(synonyms) || !synonyms.length) continue;
        const normalized = normalizeSearchText(term);
        if (!normalized) continue;
        map.set(normalized, synonyms);
      }
    }
    return map;
  } catch (error) {
    return new Map();
  }
}

function updateSearchClear() {
  if (!searchClear) return;
  const hasQuery = searchInput.value.trim().length > 0;
  searchClear.classList.toggle("hidden", !hasQuery);
}

function syncKeywordToggle() {
  if (keywordToggle) keywordToggle.value = keywordFilterState;
}

function updateGridColumns(columns) {
  if (!grid || !columns) return;
  const value = Math.min(6, Math.max(1, Math.round(Number(columns)) || 3));
  gridColumns = value;
  grid.dataset.columns = String(value);
  scheduleGridWindow();
  grid.style.setProperty("--grid-columns", value);
  if (densityValue) {
    densityValue.textContent = String(value);
  }
}

function updateLogoCount(count) {
  if (!logoCount) return;
  const label = count === 1 ? "logo" : "logos";
  logoCount.textContent = `${count} ${label}`;
}

function revokeLocalUrls() {
  resetLazyObserver();
  for (const url of localObjectUrls) {
    URL.revokeObjectURL(url);
  }
  localObjectUrls.clear();
}

function clearLogoCaches() {
  libraryGeneration += 1;
  previewCache.clear();
  previewCacheBytes = 0;
  revokeLocalUrls();
  clearSearchCache();
  clearAiSearchState();
  allLogos = [];
  logoById = new Map();
  tokenIndex = new Map();
  localLogosCache = null;
}

async function loadZipBuffer(buffer) {
  const session = getZipSession();
  return session.load(buffer);
}

function getZipSession() {
  if (!zipSession) zipSession = createZipSession();
  return zipSession;
}

function createZipSession() {
  let active;
  try { active = typeof Worker !== "undefined" ? createWorkerZipSession() : createMainZipSession(); }
  catch { active = createMainZipSession(); }
  return {
    async load(buffer) {
      try { return await active.load(buffer.slice(0)); }
      catch (error) {
        if (active.type !== "worker") throw error;
        active.terminate();
        active = createMainZipSession();
        return active.load(buffer);
      }
    },
    getSvg: name => active.getSvg(name),
    terminate: () => active.terminate ? active.terminate() : active.reset()
  };
}

function resetZipSession() {
  zipSession?.terminate?.();
  zipSession = null;
}

function createWorkerZipSession() {
  const worker = new Worker("zip-worker.js");
  const pending = new Map();
  let nextId = 0, failure = null;
  const fail = error => {
    failure = error;
    worker.terminate();
    for (const { reject, timer } of pending.values()) { clearTimeout(timer); reject(error); }
    pending.clear();
  };
  worker.onmessage = event => {
    const { id, ok, payload, error } = event.data || {};
    const request = pending.get(id);
    if (!request) return;
    clearTimeout(request.timer); pending.delete(id);
    if (ok) request.resolve(payload);
    else request.reject(new Error(error?.message || "Erreur du lecteur ZIP."));
  };
  worker.onerror = () => fail(new Error("Le lecteur ZIP ne répond plus."));
  function call(type, payload, buffer) {
    return new Promise((resolve, reject) => {
      if (failure) { reject(failure); return; }
      const id = ++nextId;
      const timer = setTimeout(() => fail(new Error("Le chargement ZIP a dépassé le délai de 30 secondes.")), 30000);
      pending.set(id, { resolve, reject, timer });
      try { worker.postMessage({ id, type, payload }, buffer ? [buffer] : []); }
      catch (error) { clearTimeout(timer); pending.delete(id); reject(error); }
    });
  }
  return { type: "worker", load: buffer => call("loadZip", { buffer }, buffer),
    getSvg: name => call("getSvg", { name }), terminate: () => fail(new Error("Bibliothèque remplacée.")) };
}

function createMainZipSession() {
  let zip = null;
  let entryMap = new Map();
  return {
    type: "main",
    async load(buffer) {
      if (typeof JSZip === "undefined") {
        throw new Error("JSZip indisponible.");
      }
      zip = await JSZip.loadAsync(buffer);
      const result = collectZipEntries(zip);
      entryMap = result.entryMap;
      return { items: result.items, stats: result.stats };
    },
    async getSvg(name) {
      if (!zip) {
        throw new Error("ZIP non chargé.");
      }
      const entryName = entryMap.get(name) || name;
      if (!entryName || !zip.file(entryName)) {
        throw new Error("SVG introuvable dans le ZIP.");
      }
      const svgText = await zip.file(entryName).async("text");
      return { name, svgText };
    },
    reset() {
      zip = null;
      entryMap = new Map();
    }
  };
}

function collectZipEntries(zip) {
  const rawItems = [];
  const entryMap = new Map();
  let ignored = 0;
  let duplicates = 0;

  zip.forEach((_, entry) => {
    if (entry.dir) {
      return;
    }
    if (!/\.svg$/i.test(entry.name)) {
      ignored += 1;
      return;
    }
    const name = extractFileName(entry.name);
    if (!name) {
      ignored += 1;
      return;
    }
    const entryName = normalizeZipEntryName(entry.name);
    if (entryMap.has(entryName)) {
      duplicates += 1;
      return;
    }
    entryMap.set(entryName, entry.name);
    rawItems.push({
      name,
      displayName: "",
      entryName,
      storageKey: entryName,
      ext: "svg",
      url: null,
      svgText: null,
      normalizedSvg: null,
      source: "local"
    });
  });

  rawItems.sort((a, b) => a.entryName.localeCompare(b.entryName));
  const nameCount = new Map();
  rawItems.forEach((item) => {
    nameCount.set(item.name, (nameCount.get(item.name) || 0) + 1);
  });
  const items = rawItems.map((item) => ({
    ...item,
    displayName:
      (nameCount.get(item.name) || 0) > 1
        ? buildZipDisplayName(item.entryName)
        : stripSvgExtension(item.name)
  }));
  items.forEach((item, index) => {
    item.id = index;
  });

  return {
    items,
    entryMap,
    stats: {
      total: items.length,
      duplicates,
      ignored
    }
  };
}

function extractFileName(filePath) {
  if (!filePath) return "";
  const normalized = filePath.replace(/\\/g, "/");
  return normalized.split("/").pop();
}

function normalizeZipEntryName(filePath) {
  return String(filePath || "").replace(/\\/g, "/");
}

function buildZipDisplayName(entryName) {
  const normalized = normalizeZipEntryName(entryName);
  const parts = normalized.split("/").filter(Boolean);
  if (!parts.length) {
    return stripSvgExtension(entryName);
  }
  const last = stripSvgExtension(parts.pop());
  if (!parts.length) {
    return last;
  }
  return `${last} (${parts.slice(-1)[0]})`;
}

function createSvgUrl(svgText) {
  const blob = new Blob([svgText], { type: "image/svg+xml" });
  const url = URL.createObjectURL(blob);
  localObjectUrls.add(url);
  return url;
}

function buildZipStatsMessage(stats) {
  if (!stats) return "ZIP chargé.";
  const parts = [`ZIP chargé (${stats.total} SVG)`];
  if (stats.duplicates) {
    parts.push(
      `${stats.duplicates} doublon${stats.duplicates > 1 ? "s" : ""} ignoré${stats.duplicates > 1 ? "s" : ""}`
    );
  }
  if (stats.ignored) {
    parts.push(
      `${stats.ignored} fichier${stats.ignored > 1 ? "s" : ""} non SVG ignoré${stats.ignored > 1 ? "s" : ""}`
    );
  }
  return parts.join(" · ");
}

function buildZipMeta(record, count) {
  const meta = record?.meta || {};
  return {
    name: meta.name || "ZIP local",
    size: meta.size || record?.buffer?.byteLength || 0,
    count: Number.isFinite(count) ? count : meta.count || 0,
    updatedAt: meta.updatedAt || Date.now()
  };
}

function updateZipMeta(meta) {
  if (!zipMeta) return;
  if (!meta) {
    setZipSummaryVisible(false);
    setZipPanelExpanded(true);
    zipPanelToggled = false;
    if (zipStatusText) {
      zipStatusText.textContent = "Aucun ZIP chargé";
    }
    zipMeta.textContent = "Aucun ZIP local chargé.";
    return;
  }
  setZipSummaryVisible(true);
  if (zipStatusText) {
    zipStatusText.textContent = `${meta.name || "Bibliothèque"} · ${meta.count || 0} SVG`;
  }
  if (!zipPanelToggled && zipPanelExpanded) {
    setZipPanelExpanded(false);
  }
  const details = [];
  if (meta.name) {
    details.push(meta.name);
  }
  if (Number.isFinite(meta.count)) {
    details.push(`${meta.count} SVG`);
  }
  if (meta.size) {
    details.push(formatBytes(meta.size));
  }
  if (meta.updatedAt) {
    details.push(formatDate(meta.updatedAt));
  }
  zipMeta.textContent = details.length
    ? `ZIP local : ${details.join(" • ")}`
    : "ZIP local chargé.";
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return "0 KB";
  }
  const units = ["B", "KB", "MB", "GB"];
  let size = bytes;
  let unitIndex = 0;
  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex += 1;
  }
  const precision = size < 10 && unitIndex > 0 ? 1 : 0;
  return `${size.toFixed(precision)} ${units[unitIndex]}`;
}

function formatDate(timestamp) {
  try {
    return new Date(timestamp).toLocaleString("fr-FR");
  } catch (error) {
    return "";
  }
}

async function readZipCache() {
  try {
    const db = await openZipCache();
    if (!db) return null;
    return await new Promise((resolve) => {
      const tx = db.transaction(DB_STORE, "readonly");
      const store = tx.objectStore(DB_STORE);
      const request = store.get(ZIP_CACHE_KEY);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => resolve(null);
      tx.oncomplete = () => db.close();
      tx.onerror = () => db.close();
      tx.onabort = () => db.close();
    });
  } catch (error) {
    return null;
  }
}

async function saveZipCache(buffer, meta) {
  try {
    const db = await openZipCache();
    if (!db) return false;
    await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readwrite");
      const store = tx.objectStore(DB_STORE);
      store.put({ buffer, meta }, ZIP_CACHE_KEY);
      tx.oncomplete = () => {
        db.close();
        resolve();
      };
      tx.onerror = () => {
        db.close();
        reject(tx.error);
      };
      tx.onabort = () => {
        db.close();
        reject(tx.error);
      };
    });
    return true;
  } catch (error) {
    return false;
  }
}

function openZipCache() {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") {
      resolve(null);
      return;
    }
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(DB_STORE)) {
        db.createObjectStore(DB_STORE);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function normalizeSvg(text) {
  let svg = text.replace(/^\uFEFF/, "").trim();
  svg = svg.replace(/<\?xml[^>]*>\s*/i, "");
  svg = svg.replace(/<!DOCTYPE[^>]*>\s*/i, "");
  if (!/xmlns=/.test(svg)) {
    svg = svg.replace(
      /<svg(\s|>)/i,
      '<svg xmlns=\"http://www.w3.org/2000/svg\"$1'
    );
  }
  return svg;
}

function setStatus(message, tone = "") {
  statusEl.textContent = message;
  statusEl.className = `status ${tone}`.trim();
}


function registerSelectionAction() {
  if (selectionActionRegistered || typeof Office.actions?.associate !== "function") return;
  try {
    Office.actions.associate("SearchSelectedPictogram", searchSelectedPictogram);
    selectionActionRegistered = true;
    selectionActionError = "";
  } catch (error) { selectionActionError = error.message || String(error); }
}

function supportsOfficeSet(name) {
  try { return Boolean(Office.context?.requirements?.isSetSupported(name, "1.1")); }
  catch { return false; }
}

function defaultSelectionShortcut() {
  return String(Office.context?.platform || "").toLowerCase() === "mac" ? "Cmd+Alt+P" : "Ctrl+Alt+P";
}

function officeCallWithTimeout(operation, message) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(message)), 5000);
    Promise.resolve().then(operation).then(resolve, reject).finally(() => clearTimeout(timer));
  });
}

async function checkShortcutRegistration() {
  const status = document.getElementById("shortcut-status");
  const key = document.getElementById("shortcut-key");
  const reset = document.getElementById("shortcut-reset");
  const manifest = document.getElementById("shortcut-manifest");
  const defaultKey = defaultSelectionShortcut();
  key.textContent = defaultKey.replaceAll("+", " + ");
  reset.classList.toggle("hidden", true);
  manifest.classList.toggle("hidden", true);
  status.dataset.state = "checking";
  status.textContent = "Vérification du raccourci auprès de PowerPoint…";
  const update = (state, message) => { status.dataset.state = state; status.textContent = message; };
  if (!supportsOfficeSet("KeyboardShortcuts")) {
    const version = Office.context?.diagnostics?.version;
    update("unsupported", `Cette version de PowerPoint ne propose pas les raccourcis de compléments${version ? ` (${version})` : ""}. Utilisez « Insérer depuis la sélection » dans le volet.`);
    return;
  }
  if (!supportsOfficeSet("SharedRuntime") || typeof Office.actions?.getShortcuts !== "function") {
    update("unavailable", "La vérification du raccourci est indisponible. Mettez à jour le complément avec le manifeste actuel, puis rouvrez-le.");
    manifest.classList.toggle("hidden", false);
    return;
  }
  registerSelectionAction();
  if (!selectionActionRegistered) {
    update("error", `L’action n’a pas pu être chargée. Rouvrez le volet.${selectionActionError ? ` ${selectionActionError}` : ""}`);
    return;
  }
  try {
    const shortcuts = await officeCallWithTimeout(() => Office.actions.getShortcuts(), "PowerPoint n’a pas répondu à la vérification. Réessayez.");
    if (!shortcuts || !Object.prototype.hasOwnProperty.call(shortcuts, "SearchSelectedPictogram")) {
      update("missing", "PowerPoint n’a pas enregistré cette action. Mettez à jour le complément avec le manifeste actuel, puis ouvrez son volet une fois.");
      manifest.classList.toggle("hidden", false);
      return;
    }
    const actual = shortcuts.SearchSelectedPictogram;
    if (actual === null) {
      update("conflict", "Ce raccourci a été attribué à une autre action. Rétablissez-le, puis choisissez Atelier Pictos si PowerPoint vous demande quelle action utiliser.");
    } else if (typeof actual === "string" && actual.trim()) {
      key.textContent = actual.replaceAll("+", " + ");
      update("registered", "Raccourci enregistré par PowerPoint. Sélectionnez le texte dans la diapositive, puis appuyez sur les touches indiquées.");
    } else {
      update("error", "PowerPoint n’a pas renvoyé de raccourci utilisable. Rouvrez le volet, puis vérifiez à nouveau.");
      return;
    }
    reset.classList.toggle("hidden", actual === defaultKey || typeof Office.actions?.replaceShortcuts !== "function");
  } catch (error) {
    update("error", `Raccourci non vérifié : ${error.message || error}`);
    manifest.classList.toggle("hidden", false);
  }
}

function refreshShortcutStatus() {
  if (shortcutCheckPromise) return shortcutCheckPromise;
  const button = document.getElementById("shortcut-refresh");
  button.disabled = true;
  shortcutCheckPromise = checkShortcutRegistration().finally(() => { button.disabled = false; shortcutCheckPromise = null; });
  return shortcutCheckPromise;
}

async function restoreSelectionShortcut() {
  const button = document.getElementById("shortcut-reset");
  button.disabled = true;
  try {
    if (shortcutCheckPromise) await shortcutCheckPromise;
    await officeCallWithTimeout(() => Office.actions.replaceShortcuts({ SearchSelectedPictogram: defaultSelectionShortcut() }), "PowerPoint n’a pas confirmé le changement. Vérifiez le raccourci à nouveau.");
    await refreshShortcutStatus();
  } catch (error) {
    const status = document.getElementById("shortcut-status");
    status.dataset.state = "error";
    status.textContent = `Raccourci non rétabli : ${error.message || error}`;
  } finally { button.disabled = false; }
}

function initShortcutControls() {
  document.getElementById("shortcut-refresh").addEventListener("click", refreshShortcutStatus);
  document.getElementById("shortcut-reset").addEventListener("click", restoreSelectionShortcut);
  refreshShortcutStatus();
  const status = document.getElementById("shortcut-startup-status");
  const startup = document.getElementById("shortcut-startup");
  const supported = supportsOfficeSet("SharedRuntime") && typeof Office.addin?.getStartupBehavior === "function" && typeof Office.addin?.setStartupBehavior === "function";
  startup.disabled = true;
  if (!supported) { status.textContent = "Le chargement en arrière-plan n’est pas disponible dans ce contexte PowerPoint."; return; }
  officeCallWithTimeout(() => Office.addin.getStartupBehavior(), "Le chargement automatique n’a pas pu être vérifié.")
    .then(value => { startup.checked = value === Office.StartupBehavior.load; startup.disabled = false; })
    .catch(error => { status.textContent = error.message || String(error); });
  startup.addEventListener("change", async () => {
    startup.disabled = true;
    try {
      await officeCallWithTimeout(() => Office.addin.setStartupBehavior(startup.checked ? Office.StartupBehavior.load : Office.StartupBehavior.none), "PowerPoint n’a pas confirmé le changement. Rouvrez le volet pour vérifier cette option.");
      status.textContent = startup.checked ? "Le complément se chargera à la prochaine ouverture de cette présentation." : "Ouvrez le volet pour charger le complément dans cette présentation.";
    } catch (error) { startup.checked = !startup.checked; status.textContent = `Option non confirmée : ${error.message || error}`; }
    finally { startup.disabled = false; }
  });
}

function getSelectedText() {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(String(result.value || "").trim());
      else reject(result.error);
    });
  });
}

async function searchSelectedPictogram(event, { revealPane = true } = {}) {
  if (shortcutBusy) { event?.completed?.(); return; }
  shortcutBusy = true;
  const button = document.getElementById("selection-insert");
  button.disabled = true;
  setStatus("Lecture du texte sélectionné…");
  try {
    // Capture the selection before opening the pane changes keyboard focus.
    const [text, slideId] = await officeCallWithTimeout(() => Promise.all([getSelectedText(), getSelectedSlideId()]), "PowerPoint n’a pas renvoyé la sélection. Sélectionnez le texte dans la diapositive et réessayez.");
    // A visible-pane button remains usable without the optional shared runtime.
    if (revealPane && supportsOfficeSet("SharedRuntime")) await Office.addin?.showAsTaskpane?.();
    await initialization;
    setSettingsPanelOpen(false);
    if (!text) throw new Error("Sélectionnez d’abord un mot ou une expression dans la diapositive.");
    if (!slideId) throw new Error("Sélectionnez une seule diapositive pour insérer un pictogramme.");
    if (text.length > 240) throw new Error("Sélectionnez une expression de 240 caractères maximum.");
    clearTimeout(searchTimer); searchTimer = null;
    searchInput.value = text;
    updateSearchClear();
    if (libraryBusy || !allLogos.length) throw new Error("Chargez votre bibliothèque ZIP avant l’insertion depuis la sélection.");
    if (aiEnabled && !isAiProviderReady()) throw new Error(providerWarning());
    const query = getNormalizedSearchQuery(), generation = libraryGeneration;
    await runSearchCycle();
    if (aiSearchState.error) throw new Error("La recherche IA a échoué. Aucun logo n’a été inséré.");
    const top = getRenderableLogos()[0];
    if (!top) throw new Error("Aucun pictogramme trouvé pour cette sélection.");
    const validate = async () => {
      if (getNormalizedSearchQuery() !== query || libraryGeneration !== generation || await getSelectedSlideId() !== slideId) {
        throw new Error("La recherche ou la diapositive a changé. Relancez le raccourci pour insérer le logo.");
      }
    };
    const operation = insertQueue.then(async () => {
      const svg = await getPreparedSvg(top);
      await validate();
      if (!Office.context.requirements.isSetSupported("ImageCoercion", "1.2")) throw new Error("Cette version de PowerPoint ne prend pas en charge l’insertion SVG.");
      // A shortcut adds a new pictogram, preserving the selected source text.
      await goToSlide(slideId);
      await insertSvg(svg, await getNextInsertPosition());
      recordRecent(top);
      setStatus(`Logo inséré : ${top.displayName || top.name}`, "success");
    });
    insertQueue = operation.catch(() => {});
    await operation;
  } catch (error) {
    try { if (revealPane) await Office.addin?.showAsTaskpane?.(); } catch {}
    setStatus(error.message || String(error), "error");
  } finally { shortcutBusy = false; button.disabled = false; event?.completed?.(); }
}
