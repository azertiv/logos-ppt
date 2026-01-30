import fs from "node:fs";
import path from "node:path";

const root = path.resolve(".");
const dictDir = path.join(root, "media", "wordnet", "dict");
const outputPath = path.join(root, "public", "wordnet-synonyms.json");
const dataFiles = ["data.noun", "data.verb", "data.adj", "data.adv"];

const synonymsMap = new Map();

for (const file of dataFiles) {
  const filePath = path.join(dictDir, file);
  if (!fs.existsSync(filePath)) {
    console.error(`Missing WordNet file: ${filePath}`);
    process.exit(1);
  }
  const content = fs.readFileSync(filePath, "utf8");
  const lines = content.split(/\r?\n/);
  for (const line of lines) {
    if (!line || line.startsWith("  ")) continue;
    const pipeIndex = line.indexOf(" | ");
    const dataLine = pipeIndex >= 0 ? line.slice(0, pipeIndex) : line;
    const fields = dataLine.trim().split(/\s+/);
    if (fields.length < 5) continue;
    const wordCount = Number.parseInt(fields[3], 16);
    if (!Number.isFinite(wordCount) || wordCount <= 0) continue;
    const words = [];
    let index = 4;
    for (let i = 0; i < wordCount; i += 1) {
      const rawWord = fields[index];
      index += 2; // Skip lex_id
      const normalized = normalizeWordnetToken(rawWord);
      if (normalized) {
        words.push(normalized);
      }
    }
    if (words.length < 2) continue;
    const unique = Array.from(new Set(words));
    for (const word of unique) {
      let set = synonymsMap.get(word);
      if (!set) {
        set = new Set();
        synonymsMap.set(word, set);
      }
      for (const synonym of unique) {
        if (synonym !== word) {
          set.add(synonym);
        }
      }
    }
  }
}

const items = {};
for (const [term, set] of synonymsMap.entries()) {
  const list = Array.from(set).filter(Boolean);
  if (!list.length) continue;
  list.sort();
  items[term] = list;
}

const payload = {
  generatedAt: new Date().toISOString(),
  source: "WordNet 3.1 (wordnet-db 3.1.14)",
  items
};

fs.writeFileSync(outputPath, JSON.stringify(payload));
console.log(`Wrote ${Object.keys(items).length} WordNet entries to ${outputPath}`);

function normalizeWordnetToken(value) {
  if (!value) return "";
  const text = String(value).toLowerCase().replace(/_/g, " ");
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
