const fs = require("node:fs");
const path = require("node:path");
const { imageSize } = require("image-size");

const { extractReferenceLibrary } = require("../cli/build_reference_library");
const {
  DEFAULT_REFERENCE_LIBRARY,
  MASTER_REFERENCE_DIR,
  LIBRARIES_ROOT,
  REFERENCE_ROOT,
} = require("../utils/pathConfig");
const { ensureDir, readJson, writeJson, slugifyName, listDirectories } = require("../utils/fileUtils");

const CATEGORY_LABELS = {
  branding: "品牌素材",
  icons: "图标",
  "vector-icons": "矢量图标",
  illustrations: "插画",
  "screenshots-charts": "截图图表",
  "mixed-media": "混合素材",
  decorations: "装饰素材",
  other: "其他",
};

const USAGE_KEYWORDS = [
  { pattern: /(summary|metric|badge|chip)/i, tag: "summary-card" },
  { pattern: /(table|grid|matrix)/i, tag: "table-analysis" },
  { pattern: /(cover|title)/i, tag: "cover" },
  { pattern: /(process|flow|bridge|timeline)/i, tag: "process-flow" },
  { pattern: /(dashboard|wall|takeaway)/i, tag: "dashboard" },
  { pattern: /(logo|brand)/i, tag: "branding" },
  { pattern: /(image|gallery|storyboard|screenshot)/i, tag: "image-story" },
];

function unique(list = []) {
  return [...new Set((list || []).filter(Boolean))];
}

function safeStat(filePath) {
  try {
    return fs.statSync(filePath);
  } catch {
    return null;
  }
}

function safeImageDimensions(filePath) {
  try {
    if (!filePath || !fs.existsSync(filePath)) return {};
    const dims = imageSize(filePath);
    return {
      width: Number(dims.width || 0),
      height: Number(dims.height || 0),
      type: String(dims.type || path.extname(filePath).replace(/^\./, "") || "").toLowerCase(),
    };
  } catch {
    return {};
  }
}

function previewUrlFor(filePath = "") {
  if (!filePath) return "";
  return `/api/library-preview?path=${encodeURIComponent(path.resolve(filePath))}`;
}

function inferOrientation(width = 0, height = 0) {
  if (!width || !height) return "unknown";
  if (width / height >= 1.18) return "landscape";
  if (width / height <= 0.86) return "portrait";
  return "square";
}

function inferMaterialTags(asset = {}) {
  const tags = [];
  const category = String(asset.category || "").toLowerCase();
  const width = Number(asset.width || asset.dimensions?.width || 0);
  const height = Number(asset.height || asset.dimensions?.height || 0);
  const ext = String(path.extname(asset.name || asset.path || "") || "").toLowerCase();
  const orientation = inferOrientation(width, height);

  if (category) tags.push(category);
  if (ext === ".svg") tags.push("vector");
  if ([".png", ".jpg", ".jpeg", ".webp", ".gif"].includes(ext)) tags.push("bitmap");
  if (width && height) tags.push(orientation);
  if (Math.max(width, height) <= 420) tags.push("small");
  if (Math.max(width, height) >= 1200) tags.push("wide");
  if (category === "branding") tags.push("logo");
  if (category === "icons" || category === "vector-icons") tags.push("icon");
  if (category === "screenshots-charts") tags.push("screenshot");
  if (category === "mixed-media" || category === "illustrations" || category === "decorations") tags.push("illustration");
  return unique(tags);
}

function inferUsageTags(asset = {}) {
  const source = `${asset.name || ""} ${asset.id || ""} ${(asset.tags || []).join(" ")} ${(asset.usageTags || []).join(" ")}`;
  const tags = [];
  USAGE_KEYWORDS.forEach(({ pattern, tag }) => {
    if (pattern.test(source)) tags.push(tag);
  });
  if (!tags.length && (asset.category === "icons" || asset.category === "vector-icons")) {
    tags.push("summary-card", "process-flow");
  }
  if (!tags.length && asset.category === "branding") {
    tags.push("branding", "cover");
  }
  if (!tags.length && asset.category === "screenshots-charts") {
    tags.push("image-story", "table-analysis");
  }
  return unique(tags);
}

function normalizeAsset(asset = {}, sourceLibrary = "", categoryHint = "") {
  const filePath = path.resolve(String(asset.path || ""));
  const stats = safeStat(filePath);
  const dims = {
    ...safeImageDimensions(filePath),
    ...(asset.dimensions || {}),
  };
  const width = Number(asset.width || dims.width || 0);
  const height = Number(asset.height || dims.height || 0);
  const category = String(asset.category || categoryHint || "other").toLowerCase();

  const normalized = {
    id: String(asset.id || path.basename(filePath, path.extname(filePath)) || `${category}-${Date.now()}`),
    name: String(asset.name || path.basename(filePath) || "素材"),
    category,
    path: filePath,
    dimensions: {
      width,
      height,
      type: String(dims.type || path.extname(filePath).replace(/^\./, "")),
    },
    slides: Array.isArray(asset.slides) ? asset.slides : [],
    width,
    height,
    size: Number(asset.size || stats?.size || 0),
    previewUrl: asset.previewUrl || previewUrlFor(filePath),
    previewDataUrl: asset.previewDataUrl || "",
    sourceLibrary: asset.sourceLibrary || sourceLibrary || "",
  };

  normalized.tags = unique([...(asset.tags || []), ...inferMaterialTags(normalized)]);
  normalized.usageTags = unique([...(asset.usageTags || []), ...inferUsageTags(normalized)]);
  return normalized;
}

function normalizeAssetCollections(raw = {}, sourceLibrary = "") {
  const mapped = {
    iconAssets: [],
    brandingAssets: [],
    illustrationAssets: [],
    screenshotAssets: [],
    allAssets: [],
  };

  const sources = {
    iconAssets: [...(raw.iconAssets || []), ...(raw.icons || []), ...(raw.vectorIcons || [])],
    brandingAssets: [...(raw.brandingAssets || []), ...(raw.branding || [])],
    illustrationAssets: [...(raw.illustrationAssets || []), ...(raw.illustrations || []), ...(raw.decorations || [])],
    screenshotAssets: [...(raw.screenshotAssets || []), ...(raw.screenshots || []), ...(raw.screenshotsCharts || [])],
    allAssets: raw.allAssets || raw.assets || [],
  };

  mapped.iconAssets = sources.iconAssets.map((item) => normalizeAsset(item, sourceLibrary, item.category || "icons"));
  mapped.brandingAssets = sources.brandingAssets.map((item) => normalizeAsset(item, sourceLibrary, "branding"));
  mapped.illustrationAssets = sources.illustrationAssets.map((item) => normalizeAsset(item, sourceLibrary, item.category || "mixed-media"));
  mapped.screenshotAssets = sources.screenshotAssets.map((item) => normalizeAsset(item, sourceLibrary, item.category || "screenshots-charts"));

  const allCombined = [
    ...mapped.iconAssets,
    ...mapped.brandingAssets,
    ...mapped.illustrationAssets,
    ...mapped.screenshotAssets,
    ...sources.allAssets.map((item) => normalizeAsset(item, sourceLibrary, item.category || "other")),
  ];

  const deduped = [];
  const seen = new Set();
  allCombined.forEach((item) => {
    const key = `${item.path}::${item.id}`;
    if (seen.has(key)) return;
    seen.add(key);
    deduped.push(item);
  });
  mapped.allAssets = deduped;
  return mapped;
}

function normalizeReusableLibrary(json = {}, libraryPath = "") {
  const sourceLibrary = String(
    json.id ||
      json.libraryId ||
      json.name ||
      json.displayName ||
      path.basename(path.dirname(libraryPath || ""), "") ||
      "reference",
  );
  const rawCollections = json.assetCollections || {};
  if (json.assets && typeof json.assets === "object") {
    rawCollections.icons = json.assets.icons || [];
    rawCollections.branding = json.assets.branding || [];
    rawCollections.illustrations = json.assets.illustrations || [];
    rawCollections.screenshotsCharts = json.assets.screenshots || json.assets["screenshots-charts"] || [];
  }

  const assetCollections = normalizeAssetCollections(rawCollections, sourceLibrary);
  const componentPresets = {
    header: json.componentPresets?.header || json.header || {},
    summaryBand: json.componentPresets?.summaryBand || json.summaryBand || {},
    sectionRibbons: json.componentPresets?.sectionRibbons || [],
    badges: json.componentPresets?.badges || [],
    callouts: json.componentPresets?.callouts || [],
    cards: json.componentPresets?.cards || [],
    dividers: json.componentPresets?.dividers || [],
    tables: json.componentPresets?.tables || [],
    iconAssets: assetCollections.iconAssets,
  };

  return {
    id: sourceLibrary,
    name: String(json.name || json.displayName || sourceLibrary),
    displayName: String(json.displayName || json.name || sourceLibrary),
    sourcePptx: String(json.sourcePptx || ""),
    paletteTokens: json.paletteTokens || {},
    tagTaxonomy: json.tagTaxonomy || { materialTags: [], usageTags: [] },
    assetCollections,
    componentPresets,
  };
}

function libraryDirFrom(inputPath = "") {
  if (!inputPath) return "";
  const resolved = path.resolve(inputPath);
  try {
    const stats = fs.statSync(resolved);
    if (stats.isDirectory()) return resolved;
  } catch {
    return "";
  }
  return path.dirname(resolved);
}

function reusablePathFrom(inputPath = "") {
  if (!inputPath) return "";
  const resolved = path.resolve(inputPath);
  try {
    const stats = fs.statSync(resolved);
    if (stats.isFile()) return resolved;
  } catch {
    return "";
  }
  const candidate = path.join(resolved, "reusable_materials.json");
  return fs.existsSync(candidate) ? candidate : "";
}

function categorizeForSummary(library = {}) {
  const buckets = new Map();
  (library.assetCollections?.allAssets || []).forEach((asset) => {
    const category = String(asset.category || "other").toLowerCase();
    if (!buckets.has(category)) {
      buckets.set(category, {
        category,
        label: CATEGORY_LABELS[category] || "其他",
        count: 0,
        samples: [],
      });
    }
    const bucket = buckets.get(category);
    bucket.count += 1;
    if (bucket.samples.length < 6) bucket.samples.push(asset);
  });
  return [...buckets.values()].sort((a, b) => b.count - a.count);
}

function summarizeLibrary(inputPath = "") {
  const dir = libraryDirFrom(inputPath);
  const reusablePath = reusablePathFrom(inputPath || dir);
  if (!reusablePath || !fs.existsSync(reusablePath)) return null;
  const normalized = normalizeReusableLibrary(readJson(reusablePath, {}), reusablePath);
  const counts = {
    media: normalized.assetCollections.allAssets.length,
    icons: normalized.assetCollections.iconAssets.length,
    branding: normalized.assetCollections.brandingAssets.length,
    illustrations: normalized.assetCollections.illustrationAssets.length,
    screenshotsOrCharts: normalized.assetCollections.screenshotAssets.length,
  };
  const id = path.basename(dir || path.dirname(reusablePath));
  return {
    id,
    name: normalized.name || id,
    displayName: normalized.displayName || normalized.name || id,
    sourceName: normalized.name || id,
    dir,
    reusablePath,
    sourcePptx: normalized.sourcePptx || "",
    counts,
    categories: categorizeForSummary(normalized),
    paletteTokens: normalized.paletteTokens || {},
    componentPresets: normalized.componentPresets || {},
  };
}

function getDefaultLibrary() {
  return summarizeLibrary(DEFAULT_REFERENCE_LIBRARY);
}

function getMasterLibrary() {
  return summarizeLibrary(MASTER_REFERENCE_DIR);
}

function listReferenceLibraries() {
  const fixed = [
    summarizeLibrary(DEFAULT_REFERENCE_LIBRARY),
    summarizeLibrary(MASTER_REFERENCE_DIR),
  ].filter(Boolean);

  const dynamic = listDirectories(LIBRARIES_ROOT)
    .map((dir) => summarizeLibrary(dir))
    .filter(Boolean)
    .sort((a, b) => String(a.displayName || a.id).localeCompare(String(b.displayName || b.id), "zh-CN"));

  const seen = new Set();
  return [...fixed, ...dynamic].filter((item) => {
    if (!item?.id) return false;
    if (seen.has(item.id)) return false;
    seen.add(item.id);
    return true;
  });
}

function resolveLibraryById(id = "") {
  const target = String(id || "").trim();
  if (!target) return null;
  return listReferenceLibraries().find((item) => item.id === target) || null;
}

function writeLibraryReadme(targetDir, summary) {
  const lines = [
    "# 参考素材库",
    "",
    `- 名称：${summary.displayName || summary.name || summary.id}`,
    `- 素材总数：${summary.counts?.media || 0}`,
    `- 图标：${summary.counts?.icons || 0}`,
    `- 品牌素材：${summary.counts?.branding || 0}`,
    `- 插画：${summary.counts?.illustrations || 0}`,
    `- 截图/图表：${summary.counts?.screenshotsOrCharts || 0}`,
  ];
  fs.writeFileSync(path.join(targetDir, "README.md"), `${lines.join("\n")}\n`, "utf8");
}

function buildMergedLibrary(sourceSummaries = [], name = "当前会话参考库", sourcePptx = "") {
  const allAssets = [];
  const paletteTokens = {};
  const componentPresets = {
    header: {},
    summaryBand: {},
    sectionRibbons: [],
    badges: [],
    callouts: [],
    cards: [],
    dividers: [],
    tables: [],
    iconAssets: [],
  };

  sourceSummaries.forEach((summary) => {
    const reusablePath = summary?.reusablePath;
    if (!reusablePath || !fs.existsSync(reusablePath)) return;
    const normalized = normalizeReusableLibrary(readJson(reusablePath, {}), reusablePath);
    Object.assign(paletteTokens, normalized.paletteTokens || {});
    ["sectionRibbons", "badges", "callouts", "cards", "dividers", "tables"].forEach((key) => {
      componentPresets[key].push(...(normalized.componentPresets?.[key] || []));
    });
    if (!componentPresets.header || !Object.keys(componentPresets.header).length) {
      componentPresets.header = normalized.componentPresets?.header || {};
    }
    if (!componentPresets.summaryBand || !Object.keys(componentPresets.summaryBand).length) {
      componentPresets.summaryBand = normalized.componentPresets?.summaryBand || {};
    }
    allAssets.push(...(normalized.assetCollections?.allAssets || []));
  });

  const deduped = [];
  const seen = new Set();
  allAssets.forEach((asset) => {
    const normalized = normalizeAsset(asset, asset.sourceLibrary || "");
    const key = `${normalized.path}::${normalized.id}`;
    if (seen.has(key)) return;
    seen.add(key);
    deduped.push(normalized);
  });

  const assetCollections = normalizeAssetCollections(
    {
      allAssets: deduped,
      icons: deduped.filter((item) => ["icons", "vector-icons"].includes(item.category)),
      branding: deduped.filter((item) => item.category === "branding"),
      illustrations: deduped.filter((item) => ["mixed-media", "illustrations", "decorations"].includes(item.category)),
      screenshotsCharts: deduped.filter((item) => item.category === "screenshots-charts"),
    },
    slugifyName(name, "reference"),
  );

  componentPresets.iconAssets = assetCollections.iconAssets;
  ["sectionRibbons", "badges", "callouts", "cards", "dividers", "tables"].forEach((key) => {
    componentPresets[key] = unique(componentPresets[key].map((item) => JSON.stringify(item))).map((item) => JSON.parse(item));
  });

  return {
    name,
    displayName: name,
    sourcePptx,
    paletteTokens,
    tagTaxonomy: { materialTags: [], usageTags: [] },
    assetCollections,
    componentPresets,
  };
}

function composeReferenceLibraries({ sourceDirs = [], targetDir = "", readmeTitle = "当前会话参考库", sourcePptx = "" } = {}) {
  ensureDir(targetDir);
  const summaries = sourceDirs.map((dir) => summarizeLibrary(dir)).filter(Boolean);
  const merged = buildMergedLibrary(summaries, readmeTitle, sourcePptx);
  const reusablePath = path.join(targetDir, "reusable_materials.json");
  writeJson(reusablePath, merged);
  const summary = summarizeLibrary(reusablePath);
  writeLibraryReadme(targetDir, summary);
  return { dir: targetDir, reusablePath, summary };
}

function mergeReferenceIntoMaster(referenceDir = "") {
  const targetDir = MASTER_REFERENCE_DIR;
  ensureDir(targetDir);
  const merged = composeReferenceLibraries({
    sourceDirs: unique([MASTER_REFERENCE_DIR, referenceDir]).filter(Boolean),
    targetDir,
    readmeTitle: "主素材库",
  });
  return merged.summary;
}

async function createReferenceLibraryFromPpt(filePath = "", targetDir = "", name = "") {
  const safeName = slugifyName(name || path.basename(filePath, path.extname(filePath)) || "reference", "reference");
  const outDir = targetDir || path.join(LIBRARIES_ROOT, `${safeName}-${Date.now()}`);
  ensureDir(outDir);
  const result = await extractReferenceLibrary({
    ppt: filePath,
    previewDir: "",
    out: outDir,
  });
  const summary = summarizeLibrary(result.reusablePath || outDir);
  return {
    ...summary,
    dir: outDir,
    reusablePath: result.reusablePath || path.join(outDir, "reusable_materials.json"),
    catalogPath: result.catalogPath || path.join(outDir, "catalog.json"),
  };
}

module.exports = {
  listReferenceLibraries,
  summarizeLibrary,
  mergeReferenceIntoMaster,
  composeReferenceLibraries,
  resolveLibraryById,
  getDefaultLibrary,
  getMasterLibrary,
  normalizeReusableLibrary,
  normalizeAsset,
  inferMaterialTags,
  inferUsageTags,
  createReferenceLibraryFromPpt,
};
