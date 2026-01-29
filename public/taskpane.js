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
const sourceOnlineBtn = document.getElementById("source-online");
const sourceLocalBtn = document.getElementById("source-local");
const zipDrop = document.getElementById("zip-drop");
const zipInput = document.getElementById("zip-input");
const zipButton = document.getElementById("zip-button");
const zipMeta = document.getElementById("zip-meta");

let allLogos = [];
let keywordsMap = new Map();
let keywordFilterState = "all";
let sourceMode = "online";
let onlineLogosCache = null;
let localLogosCache = null;
let localZipRecord = null;
let keywordsPromise = null;
const localObjectUrls = new Set();

const SOURCE_PREFERENCE_KEY = "logosPptSourcePreference";
const ZIP_CACHE_KEY = "logosPptZipCache";
const DB_NAME = "logosPptCache";
const DB_VERSION = 1;
const DB_STORE = "assets";

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
  refreshBtn.addEventListener("click", () => loadLogos({ force: true }));
  searchInput.addEventListener("input", () => {
    updateSearchClear();
    renderLogos(filterLogos());
  });
  searchClear.addEventListener("click", () => {
    searchInput.value = "";
    updateSearchClear();
    renderLogos(filterLogos());
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
      syncKeywordToggle();
      renderLogos(filterLogos());
    });
    syncKeywordToggle();
  }
  if (densityRange) {
    densityRange.addEventListener("input", () => {
      updateGridColumns(Number(densityRange.value));
    });
    updateGridColumns(Number(densityRange.value));
  }
  initSourceSwitch();
  initZipDropzone();

  updateSearchClear();
  await hydrateLocalCacheMeta();
  const preferredSource = getStoredSourcePreference();
  setSource(preferredSource || "online", { persist: false });
  await loadLogos();
}

function initSourceSwitch() {
  if (sourceOnlineBtn) {
    sourceOnlineBtn.addEventListener("click", () => {
      setSource("online", { persist: true, load: true });
    });
  }
  if (sourceLocalBtn) {
    sourceLocalBtn.addEventListener("click", () => {
      setSource("local", { persist: true, load: true });
    });
  }
  syncSourceSwitch();
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

function setSource(source, options = {}) {
  const { persist = true, load = false, force = false } = options;
  sourceMode = source === "local" ? "local" : "online";
  syncSourceSwitch();
  if (persist) {
    setStoredSourcePreference(sourceMode);
  }
  if (load) {
    loadLogos({ force });
  }
}

function syncSourceSwitch() {
  if (sourceOnlineBtn) {
    const isOnline = sourceMode === "online";
    sourceOnlineBtn.classList.toggle("is-active", isOnline);
    sourceOnlineBtn.setAttribute("aria-pressed", isOnline ? "true" : "false");
  }
  if (sourceLocalBtn) {
    const isLocal = sourceMode === "local";
    sourceLocalBtn.classList.toggle("is-active", isLocal);
    sourceLocalBtn.setAttribute("aria-pressed", isLocal ? "true" : "false");
  }
}

function getStoredSourcePreference() {
  try {
    return localStorage.getItem(SOURCE_PREFERENCE_KEY);
  } catch (error) {
    return null;
  }
}

function setStoredSourcePreference(value) {
  try {
    localStorage.setItem(SOURCE_PREFERENCE_KEY, value);
  } catch (error) {
    // Ignore storage errors.
  }
}

async function hydrateLocalCacheMeta() {
  const record = await readZipCache();
  if (record && record.meta) {
    localZipRecord = record;
    updateZipMeta(record.meta);
  }
}

async function loadLogos(options = {}) {
  const { force = false } = options;
  if (sourceMode === "local") {
    await loadLocalLogos({ force });
    return;
  }
  await loadOnlineLogos({ force });
}

async function loadOnlineLogos(options = {}) {
  const { force = false } = options;
  setStatus("Chargement des logos en ligne…");
  try {
    const map = await getKeywordsMap();
    if (onlineLogosCache && !force) {
      keywordsMap = map;
      allLogos = attachKeywords(onlineLogosCache, keywordsMap);
      renderLogos(filterLogos());
      setStatus("");
      return;
    }
    const response = await fetch("logos.json", {
      cache: force ? "no-store" : "default"
    });
    const data = await response.json();
    const items = Array.isArray(data.items) ? data.items : [];
    onlineLogosCache = items.map((logo) => ({ ...logo }));
    keywordsMap = map;
    allLogos = attachKeywords(onlineLogosCache, keywordsMap);
    renderLogos(filterLogos());
    setStatus(allLogos.length ? "" : "Aucun logo disponible en ligne.");
  } catch (error) {
    console.error(error);
    setStatus("Impossible de charger la liste des logos en ligne.", "error");
    renderEmptyState("Aucun logo en ligne disponible.");
  }
}

async function loadLocalLogos(options = {}) {
  const { force = false } = options;
  setStatus("Chargement des logos locaux…");
  try {
    if (!localLogosCache || force) {
      revokeLocalUrls();
      const record = localZipRecord || await readZipCache();
      if (!record || !record.buffer) {
        renderEmptyState("Aucun ZIP local chargé. Glissez un fichier .zip pour commencer.");
        setStatus("Aucun ZIP local disponible.", "error");
        return;
      }
      localZipRecord = record;
      const parsed = await parseZipBuffer(record.buffer);
      localLogosCache = parsed.items;
      const meta = buildZipMeta(record, localLogosCache.length);
      localZipRecord.meta = meta;
      updateZipMeta(meta);
      await saveZipCache(record.buffer, meta);
    }
    const map = await getKeywordsMap();
    keywordsMap = map;
    allLogos = attachKeywords(localLogosCache, keywordsMap);
    renderLogos(filterLogos());
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
  if (typeof JSZip === "undefined") {
    setStatus("JSZip n'est pas chargé. Vérifiez la connexion réseau.", "error");
    return;
  }

  setStatus(`Import du ZIP "${file.name}"…`);

  try {
    const buffer = await file.arrayBuffer();
    revokeLocalUrls();
    const parsed = await parseZipBuffer(buffer);
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
    localZipRecord = { buffer, meta };
    localLogosCache = parsed.items;
    updateZipMeta(meta);
    await saveZipCache(buffer, meta);

    if (sourceMode !== "local") {
      setSource("local", { persist: true, load: false });
    }
    await loadLocalLogos();

    const note = buildZipStatsMessage(parsed.stats);
    setStatus(note, "success");
  } catch (error) {
    console.error(error);
    setStatus("Erreur lors de l'import du ZIP.", "error");
  }
}

function filterLogos() {
  const query = searchInput.value.trim().toLowerCase();

  return allLogos.filter((logo) => {
    const matchesQuery = query
      ? logo.name.toLowerCase().includes(query) ||
        (logo.keywords || []).some((kw) => kw.toLowerCase().includes(query))
      : true;
    const matchesFilter =
      keywordFilterState === "with"
        ? logo.hasKeywords
        : keywordFilterState === "without"
          ? !logo.hasKeywords
          : true;
    return matchesQuery && matchesFilter;
  });
}

function attachKeywords(items, map) {
  return items.map((logo) => {
    const keywords = map.get(logo.name) || [];
    return {
      ...logo,
      keywords,
      hasKeywords: Array.isArray(keywords) && keywords.length > 0
    };
  });
}

function renderEmptyState(message) {
  grid.innerHTML = "";
  updateLogoCount(0);
  const empty = document.createElement("div");
  empty.className = "status";
  empty.textContent = message;
  grid.appendChild(empty);
}

function renderLogos(logos) {
  if (!logos.length) {
    renderEmptyState("Aucun résultat pour cette recherche.");
    return;
  }

  grid.innerHTML = "";
  updateLogoCount(logos.length);

  logos.forEach((logo, index) => {
    const card = document.createElement("div");
    card.className = "logo-card";
    card.setAttribute("role", "button");
    card.setAttribute("tabindex", "0");
    card.setAttribute("aria-label", `Insérer ${logo.name}`);
    card.style.animationDelay = `${index * 20}ms`;

    const preview = document.createElement("div");
    preview.className = "logo-preview";

    const img = document.createElement("img");
    img.loading = "lazy";
    img.src = logo.url;
    img.alt = logo.name;
    preview.appendChild(img);

    card.appendChild(preview);

    card.addEventListener("click", () => insertLogo(logo));
    card.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        insertLogo(logo);
      }
    });

    grid.appendChild(card);
  });
}

async function insertLogo(logo) {
  if (!logo) return;
  if (!Office.context.requirements.isSetSupported("ImageCoercion", "1.2")) {
    setStatus(
      "Votre version de PowerPoint ne supporte pas l'insertion SVG (ImageCoercion 1.2).",
      "error"
    );
    return;
  }
  setStatus(`Insertion de ${logo.name}…`);

  try {
    const svgText = logo.svgText
      ? logo.svgText
      : await fetch(logo.url).then((res) => res.text());
    const svg = normalizeSvg(svgText);
    await forceSlideSelection();
    await insertSvg(svg);

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

async function insertSvg(svg) {
  const options = {
    coercionType: Office.CoercionType.XmlSvg,
    imageLeft: 48,
    imageTop: 48
  };

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

async function fetchKeywords() {
  try {
    const response = await fetch("keywords.json", { cache: "no-store" });
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
  for (const url of localObjectUrls) {
    URL.revokeObjectURL(url);
  }
  localObjectUrls.clear();
}

async function parseZipBuffer(buffer) {
  if (typeof JSZip === "undefined") {
    throw new Error("JSZip indisponible.");
  }

  const zip = await JSZip.loadAsync(buffer);
  const svgEntries = [];
  let ignored = 0;

  zip.forEach((_, entry) => {
    if (entry.dir) {
      return;
    }
    if (!/\.svg$/i.test(entry.name)) {
      ignored += 1;
      return;
    }
    svgEntries.push(entry);
  });

  const items = [];
  const seen = new Set();
  let duplicates = 0;

  for (const entry of svgEntries) {
    const name = extractFileName(entry.name);
    if (!name) {
      ignored += 1;
      continue;
    }
    if (seen.has(name)) {
      duplicates += 1;
      continue;
    }
    seen.add(name);
    const svgText = await entry.async("text");
    const url = createSvgUrl(svgText);
    items.push({
      name,
      ext: "svg",
      url,
      svgText,
      source: "local"
    });
  }

  items.sort((a, b) => a.name.localeCompare(b.name));

  return {
    items,
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
