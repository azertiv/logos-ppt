/* global Office */

const grid = document.getElementById("logo-grid");
const statusEl = document.getElementById("status");
const searchInput = document.getElementById("search-input");
const searchClear = document.getElementById("search-clear");
const refreshBtn = document.getElementById("refresh-btn");
const logoCount = document.getElementById("logo-count");
const keywordToggle = document.getElementById("keyword-toggle");
const densityRange = document.getElementById("density-range");
const densityValue = document.getElementById("density-value");

let allLogos = [];
let keywordsMap = new Map();
let keywordFilterState = "all";

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("Ouvrez cet add-in dans PowerPoint pour insérer les logos.", "error");
    return;
  }

  init();
});

function init() {
  refreshBtn.addEventListener("click", () => loadLogos());
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

  updateSearchClear();
  loadLogos();
}

async function loadLogos() {
  setStatus("Chargement des logos…");
  try {
    const [response, map] = await Promise.all([
      fetch("logos.json", { cache: "no-store" }),
      loadKeywords()
    ]);
    const data = await response.json();
    keywordsMap = map;
    allLogos = (data.items || []).map((logo) => {
      const keywords = keywordsMap.get(logo.name) || [];
      return {
        ...logo,
        keywords,
        hasKeywords: Array.isArray(keywords) && keywords.length > 0
      };
    });
    renderLogos(filterLogos());
    setStatus(allLogos.length ? "" : "Aucun logo trouvé dans media/logos.");
  } catch (error) {
    console.error(error);
    setStatus("Impossible de charger la liste des logos.", "error");
    updateLogoCount(0);
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

function renderLogos(logos) {
  grid.innerHTML = "";
  updateLogoCount(logos.length);

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
    const svgText = await fetch(logo.url).then((res) => res.text());
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

async function loadKeywords() {
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
