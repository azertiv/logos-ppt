const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const publicDir = path.join(root, "public");
const mediaDir = path.join(root, "media", "logos");
const distDir = path.join(root, "dist");
const distMediaDir = path.join(distDir, "media", "logos");
const keywordsPath = path.join(root, "media", "keywords.json");

const allowed = new Set([".svg"]);

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function copyFileSync(src, dest) {
  ensureDir(path.dirname(dest));
  fs.copyFileSync(src, dest);
}

function copyDirSync(srcDir, destDir) {
  if (!fs.existsSync(srcDir)) {
    return;
  }
  ensureDir(destDir);
  for (const entry of fs.readdirSync(srcDir, { withFileTypes: true })) {
    const srcPath = path.join(srcDir, entry.name);
    const destPath = path.join(destDir, entry.name);
    if (entry.isDirectory()) {
      copyDirSync(srcPath, destPath);
    } else if (entry.isFile()) {
      copyFileSync(srcPath, destPath);
    }
  }
}

function buildLogos() {
  let items = [];
  if (fs.existsSync(mediaDir)) {
    items = fs
      .readdirSync(mediaDir)
      .filter((file) => allowed.has(path.extname(file).toLowerCase()))
      .map((file) => ({
        name: file,
        ext: path.extname(file).toLowerCase().replace(".", ""),
        url: `media/logos/${encodeURIComponent(file)}`
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }

  const payload = JSON.stringify({ count: items.length, items }, null, 2);
  ensureDir(distDir);
  fs.writeFileSync(path.join(distDir, "logos.json"), payload);
}

function copyLogos() {
  if (!fs.existsSync(mediaDir)) {
    ensureDir(distMediaDir);
    return;
  }

  ensureDir(distMediaDir);
  for (const file of fs.readdirSync(mediaDir)) {
    const ext = path.extname(file).toLowerCase();
    if (!allowed.has(ext)) {
      continue;
    }
    copyFileSync(path.join(mediaDir, file), path.join(distMediaDir, file));
  }
}

function build() {
  fs.rmSync(distDir, { recursive: true, force: true });
  copyDirSync(publicDir, distDir);
  buildLogos();
  copyLogos();
  if (fs.existsSync(keywordsPath)) {
    copyFileSync(keywordsPath, path.join(distDir, "keywords.json"));
  }
  fs.writeFileSync(path.join(distDir, ".nojekyll"), "");
  console.log("Built dist/ for GitHub Pages.");
}

build();
