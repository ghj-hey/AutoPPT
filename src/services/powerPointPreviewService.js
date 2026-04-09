const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const { execFile } = require("node:child_process");
const { promisify } = require("node:util");
const { ensureDir } = require("../utils/fileUtils");
const { REFERENCE_PREVIEW_CACHE } = require("../utils/pathConfig");

const execFileAsync = promisify(execFile);

function sleep(ms = 0) {
  return new Promise((resolve) => setTimeout(resolve, Math.max(0, Number(ms) || 0)));
}

function escapePowerShellString(value) {
  return String(value || "").replace(/'/g, "''");
}

function slideIndexFromName(filePath) {
  const name = path.basename(String(filePath || ""));
  const match = name.match(/(\d+)/);
  return Number(match?.[1] || 0);
}

function sortPreviewFiles(files) {
  return [...files].sort((left, right) => slideIndexFromName(left) - slideIndexFromName(right));
}

function fileHash(filePath) {
  const buffer = fs.readFileSync(filePath);
  return crypto.createHash("sha1").update(buffer).digest("hex").slice(0, 16);
}

function listExportedImages(outputDir) {
  if (!fs.existsSync(outputDir)) return [];
  return sortPreviewFiles(
    fs
      .readdirSync(outputDir)
      .filter((name) => /\.(png|jpg|jpeg)$/i.test(name))
      .map((name) => path.join(outputDir, name)),
  );
}

async function waitForExportedImages(outputDir, options = {}) {
  const timeoutMs = Math.max(1000, Number(options.timeoutMs || 6000));
  const intervalMs = Math.max(200, Number(options.intervalMs || 300));
  const startedAt = Date.now();
  while (Date.now() - startedAt <= timeoutMs) {
    const images = listExportedImages(outputDir);
    if (images.length) return images;
    await sleep(intervalMs);
  }
  return listExportedImages(outputDir);
}

async function exportPresentationPngs(pptPath, outputDir, options = {}) {
  ensureDir(outputDir);
  const width = Number(options.width || 1600);
  const height = Number(options.height || 900);
  const cacheable = options.cache !== false;

  if (cacheable) {
    ensureDir(REFERENCE_PREVIEW_CACHE);
    const hash = fileHash(pptPath);
    const cacheDir = path.join(REFERENCE_PREVIEW_CACHE, `${hash}_${width}x${height}`);
    const cached = listExportedImages(cacheDir);
    if (cached.length) {
      return cached;
    }
  }

  const script = `
$ErrorActionPreference = 'Stop'
$pptPath = '${escapePowerShellString(path.resolve(pptPath))}'
$outDir = '${escapePowerShellString(path.resolve(outputDir))}'
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$app = $null
$presentation = $null
try {
  $app = New-Object -ComObject PowerPoint.Application
  $app.Visible = 1
  $presentation = $app.Presentations.Open($pptPath, $false, $false, $false)
  $presentation.Export($outDir, 'PNG', ${width}, ${height})
}
finally {
  if ($presentation -ne $null) { $presentation.Close() }
  if ($app -ne $null) { $app.Quit() }
}
`;

  await execFileAsync("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script], {
    windowsHide: true,
    timeout: options.timeoutMs || 180000,
    maxBuffer: 1024 * 1024 * 8,
  });

  let exported = await waitForExportedImages(outputDir, {
    timeoutMs: options.exportReadyTimeoutMs || 8000,
    intervalMs: options.exportReadyIntervalMs || 400,
  });

  if (!exported.length) {
    await sleep(options.retryDelayMs || 1200);
    exported = listExportedImages(outputDir);
  }

  if (cacheable && exported.length) {
    const hash = fileHash(pptPath);
    const cacheDir = path.join(REFERENCE_PREVIEW_CACHE, `${hash}_${width}x${height}`);
    ensureDir(cacheDir);
    exported.forEach((imagePath) => {
      const dest = path.join(cacheDir, path.basename(imagePath));
      if (!fs.existsSync(dest)) {
        fs.copyFileSync(imagePath, dest);
      }
    });
    return listExportedImages(cacheDir);
  }

  return exported;
}

function readPreviewImages(imagePaths, titles = []) {
  return imagePaths.map((imagePath, index) => {
    const ext = path.extname(imagePath).toLowerCase();
    const mime = ext === ".jpg" || ext === ".jpeg" ? "image/jpeg" : "image/png";
    const data = fs.readFileSync(imagePath).toString("base64");
    return {
      page: index + 1,
      title: titles[index] || `幻灯片 ${index + 1}`,
      imagePath,
      dataUrl: `data:${mime};base64,${data}`,
    };
  });
}

async function buildRenderedPreviews(pptPath, outputDir, titles = [], options = {}) {
  const imagePaths = await exportPresentationPngs(pptPath, outputDir, options);
  return readPreviewImages(imagePaths, titles);
}

module.exports = {
  exportPresentationPngs,
  readPreviewImages,
  buildRenderedPreviews,
};
