const express = require("express");
const https = require("https");
const path = require("path");
const fs = require("fs");
const devcerts = require("office-addin-dev-certs");

const app = express();
const port = process.env.PORT || 3000;

const publicDir = path.join(__dirname, "public");
const mediaDir = path.join(__dirname, "media");
const logosDir = path.join(mediaDir, "logos");

app.use("/", express.static(publicDir));
app.use("/media", express.static(mediaDir));

app.get("/logos.json", (req, res) => {
  const allowed = new Set([
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".svg",
    ".emf",
    ".bmp",
    ".webp"
  ]);

  let items = [];
  if (fs.existsSync(logosDir)) {
    const files = fs.readdirSync(logosDir);
    items = files
      .filter((file) => allowed.has(path.extname(file).toLowerCase()))
      .map((file) => ({
        name: file,
        ext: path.extname(file).toLowerCase().replace(".", ""),
        url: `media/logos/${encodeURIComponent(file)}`
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }

  res.json({
    count: items.length,
    items
  });
});

async function start() {
  const httpsOptions = await devcerts.getHttpsServerOptions();

  https.createServer(httpsOptions, app).listen(port, () => {
    console.log(`Logos add-in server running on port ${port}`);
  });
}

start().catch((error) => {
  console.error("Failed to start server", error);
  process.exit(1);
});
