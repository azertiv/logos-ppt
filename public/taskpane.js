/* global Office, JSZip */

const grid = document.getElementById("logo-grid");
const statusEl = document.getElementById("status");
const searchInput = document.getElementById("search-input");
const searchClear = document.getElementById("search-clear");
const refreshBtn = document.getElementById("refresh-btn");
const logoCount = document.getElementById("logo-count");
const keywordToggle = document.getElementById("keyword-toggle");
const densityRange = document.getElementById("density-range");
const densityValue = document.getElementById("density-value");
const sortSelect = document.getElementById("sort-select");
const zipDrop = document.getElementById("zip-drop");
const zipInput = document.getElementById("zip-input");
const zipButton = document.getElementById("zip-button");
const zipMeta = document.getElementById("zip-meta");
const replaceToggle = document.getElementById("replace-toggle");

let allLogos = [];
let keywordsMap = new Map();
let wordnetMap = new Map();
let keywordFilterState = "all";
let sortMode = "az";
let replaceSelectionEnabled = false;
let localLogosCache = null;
let localZipRecord = null;
let keywordsPromise = null;
let wordnetPromise = null;
let logoById = new Map();
let tokenIndex = new Map();
let searchCache = new Map();
let searchTimer = null;
let renderToken = 0;
let renderFrame = null;
let lazyObserver = null;
let zipSession = null;
let zipWorker = null;
let zipWorkerRequestId = 0;
const zipWorkerRequests = new Map();
let insertQueue = Promise.resolve();
let cachedSlideId = null;
let cachedSlideIdAt = 0;
const insertStateBySlide = new Map();
const localObjectUrls = new Set();
const favoriteSet = new Set();
const recentMap = new Map();

const STORAGE_KEYS = {
  density: "logosPptDensity",
  keywordFilter: "logosPptKeywordFilter",
  sortMode: "logosPptSortMode",
  favorites: "logosPptFavorites",
  recents: "logosPptRecents",
  replaceSelection: "logosPptReplaceSelection"
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
const RENDER_BATCH_SIZE = 72;
const LAZY_ROOT_MARGIN = "160px";
const SLIDE_ID_CACHE_MS = 1000;
const INSERT_BASE_POSITION = { left: 48, top: 48 };
const INSERT_OFFSET_STEP = { x: 18, y: 18 };
const INSERT_OFFSET_STEPS = 8;
const INSERT_RESET_MS = 60000;
const RECENT_LIMIT = 80;

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("Ouvrez cet add-in dans PowerPoint pour insérer les logos.", "error");
    return;
  }

  init().catch((error) => {
    console.error(error);
    setStatus("Erreur d'initialisation de l'add-in.", "error");
  });
});

async function init() {
  restorePreferences();
  refreshBtn.addEventListener("click", () => loadLogos({ force: true }));
  searchInput.addEventListener("input", () => {
    updateSearchClear();
    scheduleSearch();
  });
  searchClear.addEventListener("click", () => {
    searchInput.value = "";
    updateSearchClear();
    scheduleSearch({ immediate: true });
    searchInput.focus();
  });
  if (keywordToggle) {
    keywordToggle.addEventListener("click", () => {
      keywordFilterState =
        keywordFilterState === "all"
          ? "with"
          : keywordFilterState === "with"
            ? "without"
            : "all";
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
      requestRender();
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
  initZipDropzone();

  updateSearchClear();
  await hydrateLocalCacheMeta();
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
      const clamped = Math.min(4, Math.max(1, storedDensity));
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

function safeStorageGet(key) {
  try {
    return window.localStorage ? window.localStorage.getItem(key) : null;
  } catch (error) {
    return null;
  }
}

function safeStorageSet(key, value) {
  try {
    if (!window.localStorage) return;
    window.localStorage.setItem(key, value);
  } catch (error) {
    // Ignore storage errors (private mode, quota).
  }
}

function isValidSortMode(value) {
  return ["az", "recent", "favorites"].includes(value);
}

function syncSortSelect() {
  if (!sortSelect) return;
  sortSelect.value = sortMode;
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
        if (!entry || typeof entry.name !== "string") return;
        const usedAt = Number(entry.usedAt) || 0;
        if (usedAt > 0) {
          recentMap.set(entry.name, usedAt);
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
    .map(([name, usedAt]) => ({ name, usedAt }))
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
    requestRender();
    return;
  }
  searchTimer = setTimeout(() => {
    searchTimer = null;
    requestRender();
  }, SEARCH_DEBOUNCE_MS);
}

function requestRender() {
  if (renderFrame) {
    cancelAnimationFrame(renderFrame);
  }
  renderFrame = requestAnimationFrame(() => {
    renderFrame = null;
    renderLogos(filterLogos());
  });
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
  if (event.key !== "Enter" && event.key !== " ") return;
  if (event.target.closest(".favorite-toggle")) return;
  const card = event.target.closest(".logo-card");
  if (!card || !grid.contains(card)) return;
  event.preventDefault();
  const logoId = Number(card.dataset.logoId);
  const logo = logoById.get(logoId);
  if (logo) {
    insertLogo(logo);
  }
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

async function hydrateLocalCacheMeta() {
  const record = await readZipCache();
  if (record && record.meta) {
    localZipRecord = { meta: record.meta };
    updateZipMeta(record.meta);
  }
}

async function getZipRecordWithBuffer() {
  if (localZipRecord && localZipRecord.buffer) {
    return localZipRecord;
  }
  return await readZipCache();
}

async function loadLogos(options = {}) {
  const { force = false } = options;
  if (force) {
    keywordsPromise = null;
    wordnetPromise = null;
  }
  await loadLocalLogos({ force });
}

async function loadLocalLogos(options = {}) {
  const { force = false } = options;
  setStatus("Chargement des logos locaux…");
  try {
    if (!localLogosCache || force) {
      clearLogoCaches();
      resetZipSession();
      const record = await getZipRecordWithBuffer();
      if (!record || !record.buffer) {
        renderEmptyState("Aucun ZIP local chargé. Glissez un fichier .zip pour commencer.");
        setStatus("Aucun ZIP local disponible.", "error");
        return;
      }
      localZipRecord = record;
      const parsed = await loadZipBuffer(record.buffer);
      localLogosCache = parsed.items;
      const meta = buildZipMeta(record, localLogosCache.length);
      localZipRecord = { meta };
      updateZipMeta(meta);
      if (record.buffer) {
        await saveZipCache(record.buffer, meta);
      }
    }
    const map = await getKeywordsMap();
    keywordsMap = map;
    wordnetMap = await getWordnetMap();
    allLogos = attachKeywords(localLogosCache, keywordsMap);
    buildSearchIndex(allLogos);
    requestRender();
    setStatus(allLogos.length ? "" : "Aucun logo SVG trouvé dans le ZIP local.");
  } catch (error) {
    console.error(error);
    renderEmptyState("Impossible de charger les logos locaux.");
    setStatus("Impossible de charger les logos locaux.", "error");
  }
}

async function handleZipFile(file) {
  if (!file || !/\.zip$/i.test(file.name)) {
    setStatus("Merci de sélectionner un fichier .zip contenant des SVG.", "error");
    return;
  }
  if (typeof JSZip === "undefined" && typeof Worker === "undefined") {
    setStatus("JSZip n'est pas chargé. Vérifiez la connexion réseau.", "error");
    return;
  }

  setStatus(`Import du ZIP "${file.name}"…`);

  try {
    const buffer = await file.arrayBuffer();
    clearLogoCaches();
    resetZipSession();
    const parsed = await loadZipBuffer(buffer);
    if (!parsed.items.length) {
      renderEmptyState("Aucun SVG trouvé dans le ZIP.");
      setStatus("Aucun SVG trouvé dans le ZIP.", "error");
      return;
    }
    const meta = {
      name: file.name,
      size: file.size,
      count: parsed.items.length,
      updatedAt: Date.now()
    };
    localZipRecord = { meta };
    localLogosCache = parsed.items;
    updateZipMeta(meta);
    await saveZipCache(buffer, meta);

    await loadLocalLogos();

    const note = buildZipStatsMessage(parsed.stats);
    setStatus(note, "success");
  } catch (error) {
    console.error(error);
    setStatus("Erreur lors de l'import du ZIP.", "error");
  }
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
  if (direct && direct.length) return direct;
  const singular = depluralizeToken(token);
  if (singular && singular !== token) {
    const fallback = wordnetMap.get(singular);
    if (fallback && fallback.length) return fallback;
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

  const filtered = candidates.filter((logo) => {
    if (keywordFilterState === "with") return logo.hasKeywords;
    if (keywordFilterState === "without") return !logo.hasKeywords;
    return true;
  });

  let sorted = [];
  if (queryTokens.length) {
    sorted = filtered.slice();
    for (const logo of sorted) {
      logo.relevanceScore = scoreLogo(logo, queryTokens, groups);
    }
    sorted.sort((a, b) => {
      const delta = (b.relevanceScore || 0) - (a.relevanceScore || 0);
      if (delta !== 0) return delta;
      return compareBySortMode(a, b);
    });
  } else {
    sorted = sortLogos(filtered);
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

function attachKeywords(items, map) {
  return items.map((logo) => {
    const keywords = map.get(logo.name) || [];
    const searchText = buildSearchText(logo.name, keywords);
    const lastUsedAt = recentMap.get(logo.name) || 0;
    const isFavorite = favoriteSet.has(logo.name);
    return {
      ...logo,
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
  resetLazyObserver();
  grid.replaceChildren();
  updateLogoCount(0);
  const empty = document.createElement("div");
  empty.className = "status";
  empty.textContent = message;
  grid.appendChild(empty);
}

function renderLogos(logos) {
  renderToken += 1;
  const token = renderToken;
  resetLazyObserver();

  if (!logos.length) {
    renderEmptyState("Aucun résultat pour cette recherche.");
    return;
  }

  grid.replaceChildren();
  updateLogoCount(logos.length);

  const maxScore = logos.reduce(
    (max, logo) => Math.max(max, logo.relevanceScore || 0),
    0
  );

  let index = 0;
  const total = logos.length;

  const renderChunk = () => {
    if (token !== renderToken) return;
    const fragment = document.createDocumentFragment();
    const end = Math.min(index + RENDER_BATCH_SIZE, total);
    for (; index < end; index += 1) {
      fragment.appendChild(createLogoCard(logos[index], index, maxScore));
    }
    grid.appendChild(fragment);
    if (index < total) {
      requestAnimationFrame(renderChunk);
    }
  };

  renderChunk();
}

function createLogoCard(logo, index, maxScore) {
  const card = document.createElement("div");
  card.className = "logo-card";
  if (logo.isFavorite) {
    card.classList.add("is-favorite");
  }
  card.setAttribute("role", "button");
  card.setAttribute("tabindex", "0");
  card.setAttribute("aria-label", `Insérer ${logo.name}`);
  card.style.animationDelay = `${index * 20}ms`;
  card.dataset.logoId = String(logo.id ?? index);
  const score = Number(logo.relevanceScore) || 0;
  const normalized = maxScore > 0 ? Math.min(1, score / maxScore) : 0;
  const baseAlpha = maxScore > 0 ? 0.08 : 0.04;
  const alpha = baseAlpha + normalized * 0.45;
  const alphaStrong = baseAlpha + normalized * 0.7;
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
  if (logo.url) {
    img.src = logo.url;
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
    { root: null, rootMargin: LAZY_ROOT_MARGIN, threshold: 0.1 }
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
    });
}

function toggleFavoriteById(logoId, button) {
  if (!Number.isFinite(logoId)) return;
  const logo = logoById.get(logoId);
  if (!logo || !logo.name) return;
  const nextState = !logo.isFavorite;
  logo.isFavorite = nextState;
  if (nextState) {
    favoriteSet.add(logo.name);
  } else {
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
  recentMap.set(logo.name, usedAt);
  logo.lastUsedAt = usedAt;
  persistRecents();
  clearSearchCache();
  if (sortMode === "recent") {
    requestRender();
  }
}

async function ensureLogoUrl(logo) {
  if (logo.url) return logo.url;
  if (logo.urlPromise) return logo.urlPromise;
  logo.urlPromise = (async () => {
    const svgText = await getSvgText(logo);
    const url = createSvgUrl(svgText);
    logo.url = url;
    return url;
  })();
  try {
    return await logo.urlPromise;
  } finally {
    logo.urlPromise = null;
  }
}

async function getSvgText(logo) {
  if (logo.svgText) return logo.svgText;
  if (logo.svgPromise) return logo.svgPromise;
  if (!logo.name) {
    throw new Error("SVG introuvable en local.");
  }
  logo.svgPromise = fetchSvgTextFromZip(logo.name).then((text) => {
    logo.svgText = text;
    return text;
  });
  try {
    return await logo.svgPromise;
  } finally {
    logo.svgPromise = null;
  }
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
  if (logo.normalizedSvg) return logo.normalizedSvg;
  const svgText = await getSvgText(logo);
  const normalized = normalizeSvg(svgText);
  logo.normalizedSvg = normalized;
  return normalized;
}

async function getReplaceSelectionPosition() {
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
      shapes.load("items");
      await context.sync();
      if (!shapes.items || shapes.items.length !== 1) {
        return;
      }
      const shape = shapes.items[0];
      shape.load(["left", "top", "width", "height"]);
      await context.sync();
      info = {
        imageLeft: shape.left,
        imageTop: shape.top,
        imageWidth: shape.width,
        imageHeight: shape.height
      };
    });
  } catch (error) {
    console.warn("Impossible de récupérer la sélection.", error);
    return null;
  }
  return info;
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
    const position = replaceSelectionEnabled
      ? await getReplaceSelectionPosition()
      : null;
    if (position) {
      await insertSvg(svg, position);
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
    const slide = slides.items[0];
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
  return new Promise((resolve) => {
    Office.context.document.goToByIdAsync(
      slideId,
      Office.GoToType.Slide,
      () => resolve()
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
    const response = await fetch("keywords.json", { cache: "force-cache" });
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
    const response = await fetch(WORDNET_SYNONYMS_URL, { cache: "force-cache" });
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
          map.set(term, normalizeSynonymList(synonyms));
        }
      }
    } else if (items && typeof items === "object") {
      for (const [term, synonyms] of Object.entries(items)) {
        if (!Array.isArray(synonyms) || !synonyms.length) continue;
        const normalized = normalizeSearchText(term);
        if (!normalized) continue;
        map.set(normalized, normalizeSynonymList(synonyms));
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
  if (!keywordToggle) return;
  const isWith = keywordFilterState === "with";
  const isWithout = keywordFilterState === "without";
  keywordToggle.classList.toggle("is-active", isWith);
  keywordToggle.classList.toggle("is-negative", isWithout);
  keywordToggle.setAttribute(
    "aria-pressed",
    isWith ? "true" : isWithout ? "mixed" : "false"
  );
  keywordToggle.setAttribute(
    "aria-label",
    isWith
      ? "Keywords activés"
      : isWithout
        ? "Keywords désactivés"
        : "Keywords sans filtre"
  );
}

function updateGridColumns(columns) {
  if (!grid || !columns) return;
  const value = Math.min(4, Math.max(1, Number(columns)));
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
  revokeLocalUrls();
  clearSearchCache();
  allLogos = [];
  logoById = new Map();
  tokenIndex = new Map();
  localLogosCache = null;
}

async function loadZipBuffer(buffer) {
  const session = getZipSession();
  if (!session) {
    throw new Error("Impossible d'initialiser le lecteur ZIP.");
  }
  const useWorker = session.type === "worker";
  const payload = useWorker ? buffer.slice(0) : buffer;
  try {
    return await session.load(payload);
  } catch (error) {
    if (useWorker) {
      console.warn("ZIP worker indisponible, repli sur le thread principal.", error);
      resetZipSession({ terminate: true });
      zipSession = createMainZipSession();
      return await zipSession.load(buffer);
    }
    throw error;
  }
}

function getZipSession() {
  if (zipSession) {
    return zipSession;
  }
  if (typeof Worker !== "undefined") {
    try {
      zipSession = createWorkerZipSession();
      return zipSession;
    } catch (error) {
      console.warn("Impossible d'initialiser le worker ZIP.", error);
    }
  }
  zipSession = createMainZipSession();
  return zipSession;
}

function resetZipSession(options = {}) {
  const { terminate = false } = options;
  if (!zipSession) return;
  if (zipSession.reset) {
    try {
      zipSession.reset();
    } catch (error) {
      // Ignore reset errors.
    }
  }
  if (terminate && zipSession.terminate) {
    zipSession.terminate();
  }
  zipSession = null;
}

function createWorkerZipSession() {
  if (!zipWorker) {
    zipWorker = new Worker("zip-worker.js");
    zipWorker.onmessage = handleZipWorkerMessage;
    zipWorker.onerror = handleZipWorkerError;
  }
  return {
    type: "worker",
    load: (buffer) => postZipWorkerMessage("loadZip", { buffer }, buffer),
    getSvg: (name) => postZipWorkerMessage("getSvg", { name }),
    reset: () => postZipWorkerMessage("reset", {}).catch(() => {}),
    terminate: () => {
      zipWorker.terminate();
      zipWorker = null;
      rejectZipWorkerRequests(new Error("Worker terminé."));
    }
  };
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
      const entryName = entryMap.get(name);
      if (!entryName) {
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
  const items = [];
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
    if (entryMap.has(name)) {
      duplicates += 1;
      return;
    }
    entryMap.set(name, entry.name);
    items.push({
      id: 0,
      name,
      ext: "svg",
      url: null,
      svgText: null,
      normalizedSvg: null,
      source: "local"
    });
  });

  items.sort((a, b) => a.name.localeCompare(b.name));
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

function postZipWorkerMessage(type, payload, transfer) {
  return new Promise((resolve, reject) => {
    if (!zipWorker) {
      reject(new Error("Worker ZIP indisponible."));
      return;
    }
    const id = ++zipWorkerRequestId;
    zipWorkerRequests.set(id, { resolve, reject });
    try {
      if (transfer) {
        zipWorker.postMessage({ id, type, payload }, [transfer]);
      } else {
        zipWorker.postMessage({ id, type, payload });
      }
    } catch (error) {
      zipWorkerRequests.delete(id);
      reject(error);
    }
  });
}

function handleZipWorkerMessage(event) {
  const { id, ok, payload, error } = event.data || {};
  const pending = zipWorkerRequests.get(id);
  if (!pending) return;
  zipWorkerRequests.delete(id);
  if (ok) {
    pending.resolve(payload);
  } else {
    pending.reject(new Error(error?.message || "Erreur worker ZIP."));
  }
}

function handleZipWorkerError(event) {
  const error = event?.message
    ? new Error(event.message)
    : new Error("Erreur du worker ZIP.");
  rejectZipWorkerRequests(error);
}

function rejectZipWorkerRequests(error) {
  zipWorkerRequests.forEach(({ reject }) => reject(error));
  zipWorkerRequests.clear();
}

function extractFileName(filePath) {
  if (!filePath) return "";
  const normalized = filePath.replace(/\\/g, "/");
  return normalized.split("/").pop();
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
    zipMeta.textContent = "Aucun ZIP local chargé.";
    return;
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
    if (!db) return;
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
  } catch (error) {
    // Ignore cache write failures.
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
