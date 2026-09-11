/* global JSZip */

importScripts("vendor/jszip-3.10.1.min.js");

let zipInstance = null;
let entryMap = new Map();

self.onmessage = async (event) => {
  const { id, type, payload } = event.data || {};
  if (!id || !type) return;

  try {
    if (type === "loadZip") {
      const result = await handleLoadZip(payload?.buffer);
      respond(id, result);
      return;
    }
    if (type === "getSvg") {
      const result = await handleGetSvg(payload?.name);
      respond(id, result);
      return;
    }
    if (type === "reset") {
      zipInstance = null;
      entryMap = new Map();
      respond(id, { ok: true });
      return;
    }
    throw new Error(`Type de message inconnu: ${type}`);
  } catch (error) {
    respondError(id, error);
  }
};

async function handleLoadZip(buffer) {
  if (!buffer) {
    throw new Error("Buffer ZIP manquant.");
  }
  if (typeof JSZip === "undefined") {
    throw new Error("JSZip indisponible dans le worker.");
  }

  zipInstance = await JSZip.loadAsync(buffer);
  entryMap = new Map();
  const rawItems = [];
  let ignored = 0;
  let duplicates = 0;

  zipInstance.forEach((_, entry) => {
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

  return {
    items,
    stats: {
      total: items.length,
      duplicates,
      ignored
    }
  };
}

async function handleGetSvg(name) {
  if (!zipInstance) {
    throw new Error("ZIP non chargé.");
  }
  if (!name) {
    throw new Error("Nom de fichier manquant.");
  }
  const entryName = entryMap.get(name) || name;
  if (!entryName || !zipInstance.file(entryName)) {
    throw new Error("SVG introuvable dans le ZIP.");
  }
  const svgText = await zipInstance.file(entryName).async("text");
  return { name, svgText };
}

function extractFileName(filePath) {
  if (!filePath) return "";
  const normalized = filePath.replace(/\\/g, "/");
  return normalized.split("/").pop();
}

function normalizeZipEntryName(filePath) {
  return String(filePath || "").replace(/\\/g, "/");
}

function stripSvgExtension(value) {
  return String(value || "").replace(/\.svg$/i, "");
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

function respond(id, payload) {
  self.postMessage({ id, ok: true, payload });
}

function respondError(id, error) {
  self.postMessage({
    id,
    ok: false,
    error: serializeError(error)
  });
}

function serializeError(error) {
  if (!error) {
    return { message: "Erreur inconnue." };
  }
  if (typeof error === "string") {
    return { message: error };
  }
  return {
    message: error.message || String(error)
  };
}
