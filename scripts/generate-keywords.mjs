import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = path.resolve(".");
const logosDir = path.join(root, "media", "logos");
const promptPath = path.join(root, "prompts", "keywords.txt");
const outputPath = path.join(root, "media", "keywords.json");
const tmpDir = path.join(root, ".tmp", "raster");

const model = process.env.GEMINI_MODEL || "gemini-flash-lite-latest";
const apiKey = process.env.GEMINI_API_KEY;
const force = process.env.FORCE === "1";

if (!apiKey) {
  console.error("Missing GEMINI_API_KEY.");
  process.exit(1);
}

if (!fs.existsSync(promptPath)) {
  console.error(`Missing prompt file: ${promptPath}`);
  process.exit(1);
}

const prompt = fs.readFileSync(promptPath, "utf8").trim();

const svgFiles = fs.existsSync(logosDir)
  ? fs
      .readdirSync(logosDir)
      .filter((file) => path.extname(file).toLowerCase() === ".svg")
      .sort((a, b) => a.localeCompare(b))
  : [];

if (!svgFiles.length) {
  console.error("No SVG files found in media/logos.");
  process.exit(1);
}

fs.mkdirSync(tmpDir, { recursive: true });

let existing = { items: [] };
if (!force && fs.existsSync(outputPath)) {
  try {
    existing = JSON.parse(fs.readFileSync(outputPath, "utf8"));
  } catch (error) {
    existing = { items: [] };
  }
}

const existingMap = new Map(
  (existing.items || []).map((item) => [item.file, item])
);

const results = [];

for (const file of svgFiles) {
  if (!force && existingMap.has(file)) {
    results.push(existingMap.get(file));
    continue;
  }

  const svgPath = path.join(logosDir, file);
  const pngPath = path.join(tmpDir, file.replace(/\\.svg$/i, ".png"));

  try {
    execFileSync("rsvg-convert", [
      svgPath,
      "-w",
      "512",
      "-h",
      "512",
      "-o",
      pngPath
    ]);
  } catch (error) {
    console.error(`Failed to rasterize ${file}:`, error.message || error);
    continue;
  }

  const base64 = fs.readFileSync(pngPath, "base64");
  const text = await callGemini(model, apiKey, prompt, base64);
  const keywords = parseKeywords(text);

  results.push({
    file,
    keywords,
    raw: text
  });
}

const payload = {
  generatedAt: new Date().toISOString(),
  model,
  items: results
};

fs.writeFileSync(outputPath, JSON.stringify(payload, null, 2));
console.log(`Wrote ${results.length} entries to ${outputPath}`);

async function callGemini(modelName, key, promptText, imageBase64) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${key}`;
  const body = {
    contents: [
      {
        role: "user",
        parts: [
          { text: promptText },
          { inline_data: { mime_type: "image/png", data: imageBase64 } }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.2,
      topP: 0.9,
      maxOutputTokens: 1024
    }
  };

  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Gemini API error ${response.status}: ${errText}`);
  }

  const data = await response.json();
  const parts = data?.candidates?.[0]?.content?.parts || [];
  const text = parts.map((part) => part.text || "").join("").trim();
  if (!text) {
    throw new Error("Empty response from Gemini.");
  }
  return text;
}

function parseKeywords(text) {
  const items = text
    .split(";")
    .map((item) => item.trim())
    .filter(Boolean);

  const seen = new Set();
  const unique = [];
  for (const item of items) {
    const key = item.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push(item);
  }
  return unique;
}
