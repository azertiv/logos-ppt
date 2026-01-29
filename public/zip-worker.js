/* global JSZip */

importScripts("https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js");

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
  const items = [];
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
    if (entryMap.has(name)) {
      duplicates += 1;
      return;
    }
    entryMap.set(name, entry.name);
    items.push({
      name,
      ext: "svg",
      source: "local"
    });
  });

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

async function handleGetSvg(name) {
  if (!zipInstance) {
    throw new Error("ZIP non chargé.");
  }
  if (!name) {
    throw new Error("Nom de fichier manquant.");
  }
  const entryName = entryMap.get(name);
  if (!entryName) {
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
