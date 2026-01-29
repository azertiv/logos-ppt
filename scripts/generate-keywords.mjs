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
const concurrency = Math.max(
  1,
  Number.parseInt(process.env.CONCURRENCY || "10", 10) || 10
);

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

const results = new Array(svgFiles.length).fill(null);
const tasks = [];

svgFiles.forEach((file, index) => {
  if (!force && existingMap.has(file)) {
    results[index] = existingMap.get(file);
    return;
  }
  tasks.push({ file, index });
});

if (tasks.length) {
  console.log(`Processing ${tasks.length} new file(s) with concurrency ${concurrency}.`);
}

await runPool(tasks, concurrency, async ({ file, index }) => {
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
    return;
  }

  const base64 = fs.readFileSync(pngPath, "base64");
  let text = "";
  try {
    text = await callGeminiWithRetry(model, apiKey, prompt, base64);
  } catch (error) {
    console.error(`Gemini failed for ${file}:`, error.message || error);
    return;
  }

  const keywords = parseKeywords(text);
  results[index] = {
    file,
    keywords,
    raw: text
  };
});

const payload = {
  generatedAt: new Date().toISOString(),
  model,
  items: results.filter(Boolean)
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

async function callGeminiWithRetry(modelName, key, promptText, imageBase64) {
  const maxAttempts = 3;
  const delayMs = 10_000;
  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      return await callGemini(modelName, key, promptText, imageBase64);
    } catch (error) {
      lastError = error;
      if (attempt < maxAttempts) {
        console.warn(
          `Gemini attempt ${attempt} failed, retrying in ${delayMs / 1000}s...`
        );
        await sleep(delayMs);
      }
    }
  }

  throw lastError || new Error("Gemini failed after retries.");
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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

async function runPool(items, limit, worker) {
  if (!items.length) return;
  const queue = items.slice();
  const workers = Array.from(
    { length: Math.min(limit, queue.length) },
    async () => {
      while (queue.length) {
        const item = queue.shift();
        if (!item) return;
        await worker(item);
      }
    }
  );
  await Promise.all(workers);
}
