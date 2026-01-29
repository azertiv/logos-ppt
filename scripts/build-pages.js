const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const publicDir = path.join(root, "public");
const distDir = path.join(root, "dist");
const keywordsPath = path.join(root, "media", "keywords.json");

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


function build() {
  fs.rmSync(distDir, { recursive: true, force: true });
  copyDirSync(publicDir, distDir);
  if (fs.existsSync(keywordsPath)) {
    copyFileSync(keywordsPath, path.join(distDir, "keywords.json"));
  }
  fs.writeFileSync(path.join(distDir, ".nojekyll"), "");
  console.log("Built dist/ for GitHub Pages.");
}

build();
