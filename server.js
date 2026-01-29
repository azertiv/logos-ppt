const express = require("express");
const https = require("https");
const path = require("path");
const fs = require("fs");
const devcerts = require("office-addin-dev-certs");

const app = express();
const port = process.env.PORT || 3000;

const publicDir = path.join(__dirname, "public");
const mediaDir = path.join(__dirname, "media");
const keywordsPath = path.join(mediaDir, "keywords.json");

app.use("/", express.static(publicDir));
app.use("/media", express.static(mediaDir));

app.get("/keywords.json", (req, res) => {
  if (!fs.existsSync(keywordsPath)) {
    res.status(404).json({ items: [] });
    return;
  }
  res.sendFile(keywordsPath);
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
