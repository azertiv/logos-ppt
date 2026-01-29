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

    const img = document.createElement("img");
    img.loading = "lazy";
    img.src = logo.url;
    img.alt = logo.name;
    preview.appendChild(img);

    const meta = document.createElement("div");
    meta.className = "logo-meta";

    const name = document.createElement("span");
    name.className = "logo-name";
    name.textContent = logo.name;

    const ext = document.createElement("span");
    ext.className = "logo-ext";
    ext.textContent = "svg";

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
  if (!Office.context.requirements.isSetSupported("ImageCoercion", "1.2")) {
    setStatus(
      "Votre version de PowerPoint ne supporte pas l'insertion SVG (ImageCoercion 1.2).",
      "error"
    );
    return;
  }
  setStatus(`Insertion de ${logo.name}…`);

  try {
    const svgText = await fetch(logo.url).then((res) => res.text());
    const svg = normalizeSvg(svgText);
    await setSelectedData(svg, { coercionType: Office.CoercionType.XmlSvg });

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
