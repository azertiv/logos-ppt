/* global Office */

const grid = document.getElementById("logo-grid");
const statusEl = document.getElementById("status");
const searchInput = document.getElementById("search-input");
const refreshBtn = document.getElementById("refresh-btn");
const dropZone = document.getElementById("drop-zone");
const logoCount = document.getElementById("logo-count");

let allLogos = [];
let draggingLogo = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("Ouvrez cet add-in dans PowerPoint pour insérer les pictogrammes.", "error");
    return;
  }

  init();
});

function init() {
  refreshBtn.addEventListener("click", () => loadLogos());
  searchInput.addEventListener("input", () => renderLogos(filterLogos()));

  dropZone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropZone.classList.add("over");
  });

  dropZone.addEventListener("dragleave", () => dropZone.classList.remove("over"));

  dropZone.addEventListener("drop", async (event) => {
    event.preventDefault();
    dropZone.classList.remove("over");
    if (draggingLogo) {
      await insertLogo(draggingLogo);
      draggingLogo = null;
    }
  });

  loadLogos();
}

async function loadLogos() {
  setStatus("Chargement des pictogrammes…");
  try {
    const response = await fetch("logos.json", { cache: "no-store" });
    const data = await response.json();
    allLogos = data.items || [];
    logoCount.textContent = `${data.count || allLogos.length} logos`;
    renderLogos(filterLogos());
    setStatus(allLogos.length ? "" : "Aucun logo trouvé dans media/logos.");
  } catch (error) {
    console.error(error);
    setStatus("Impossible de charger la liste des logos.", "error");
  }
}

function filterLogos() {
  const query = searchInput.value.trim().toLowerCase();
  if (!query) return allLogos;
  return allLogos.filter((logo) => logo.name.toLowerCase().includes(query));
}

function renderLogos(logos) {
  grid.innerHTML = "";

  if (!logos.length) {
    const empty = document.createElement("div");
    empty.className = "status";
    empty.textContent = "Aucun résultat pour cette recherche.";
    grid.appendChild(empty);
    return;
  }

  logos.forEach((logo, index) => {
    const card = document.createElement("div");
    card.className = "logo-card";
    card.setAttribute("role", "button");
    card.setAttribute("tabindex", "0");
    card.setAttribute("aria-label", `Insérer ${logo.name}`);
    card.style.animationDelay = `${index * 20}ms`;
    card.draggable = true;

    const preview = document.createElement("div");
    preview.className = "logo-preview";

    if (logo.ext === "emf") {
      preview.classList.add("placeholder");
      preview.textContent = "EMF";
    } else {
      const img = document.createElement("img");
      img.loading = "lazy";
      img.src = logo.url;
      img.alt = logo.name;
      preview.appendChild(img);
    }

    const meta = document.createElement("div");
    meta.className = "logo-meta";

    const name = document.createElement("span");
    name.className = "logo-name";
    name.textContent = logo.name;

    const ext = document.createElement("span");
    ext.className = "logo-ext";
    ext.textContent = logo.ext || "?";

    meta.appendChild(name);
    meta.appendChild(ext);

    card.appendChild(preview);
    card.appendChild(meta);

    card.addEventListener("click", () => insertLogo(logo));
    card.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        insertLogo(logo);
      }
    });

    card.addEventListener("dragstart", (event) => {
      draggingLogo = logo;
      event.dataTransfer.effectAllowed = "copy";
      event.dataTransfer.setData("text/plain", logo.url);
      dropZone.classList.add("active");
    });

    card.addEventListener("dragend", () => {
      dropZone.classList.remove("active");
      dropZone.classList.remove("over");
    });

    grid.appendChild(card);
  });
}

async function insertLogo(logo) {
  if (!logo) return;
  setStatus(`Insertion de ${logo.name}…`);

  try {
    if (logo.ext === "svg") {
      const svgText = await fetch(logo.url).then((res) => res.text());
      const svgBase64 = toBase64(svgText);
      await setSelectedData(svgBase64, { coercionType: Office.CoercionType.XmlSvg });
    } else {
      const blob = await fetch(logo.url).then((res) => res.blob());
      const base64 = await blobToBase64(blob);
      await setSelectedData(base64, { coercionType: Office.CoercionType.Image });
    }

    setStatus(`Logo inséré : ${logo.name}`, "success");
  } catch (error) {
    console.error(error);
    setStatus(`Erreur d'insertion : ${error.message || error}`, "error");
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

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result || "";
      const comma = dataUrl.indexOf(",");
      resolve(comma >= 0 ? dataUrl.slice(comma + 1) : dataUrl);
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

function toBase64(text) {
  return btoa(unescape(encodeURIComponent(text)));
}

function setStatus(message, tone = "") {
  statusEl.textContent = message;
  statusEl.className = `status ${tone}`.trim();
}
