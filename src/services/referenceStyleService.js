const fs = require("node:fs");
const path = require("node:path");
const { imageSize } = require("image-size");

function fileToDataUrl(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return "";
  const ext = path.extname(filePath).toLowerCase();
  const mime = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".svg": "image/svg+xml",
  }[ext] || "application/octet-stream";
  const data = fs.readFileSync(filePath).toString("base64");
  return `data:${mime};base64,${data}`;
}

function safeImageMeta(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return null;
  try {
    const stats = fs.statSync(filePath);
    const dims = imageSize(fs.readFileSync(filePath));
    const width = Math.max(1, dims.width || 1);
    const height = Math.max(1, dims.height || 1);
    const pixels = width * height;
    return {
      path: filePath,
      name: path.basename(filePath),
      size: stats.size,
      width,
      height,
      pixels,
      megapixels: Number((pixels / 1000000).toFixed(2)),
      aspectRatio: Number((width / height).toFixed(3)),
      orientation: width / height >= 1.18 ? "landscape" : width / height <= 0.86 ? "portrait" : "square",
      highResolution: width >= 1800 || height >= 1000 || pixels >= 1800000,
      previewDataUrl: fileToDataUrl(filePath),
    };
  } catch {
    return {
      path: filePath,
      name: path.basename(filePath),
      size: fs.existsSync(filePath) ? fs.statSync(filePath).size : 0,
      width: 0,
      height: 0,
      pixels: 0,
      megapixels: 0,
      aspectRatio: 0,
      orientation: "unknown",
      highResolution: false,
      previewDataUrl: fileToDataUrl(filePath),
    };
  }
}

function ratio(count, total) {
  return total ? Number((count / total).toFixed(3)) : 0;
}

function unique(list = []) {
  return [...new Set((list || []).filter(Boolean))];
}

function pickSuggestedLayoutSet(profile) {
  if (!profile?.count) return "";
  if (profile.pageRhythm === "dense" || profile.tablePreference === "dense") return "bank_finance_dense";
  if (profile.styleFamily === "boardroom-report") return "bank_finance_boardroom";
  if (profile.styleFamily === "visual-report") return "bank_finance_visual";
  if (profile.cardStyle === "dashboard-card" || profile.tablePreference === "dashboard") return "bank_finance_dynamic";
  return "bank_finance_default";
}

function inferStyleFamily(profile) {
  if (profile.landscapeRatio >= 0.72 && profile.highResolutionRatio >= 0.66 && profile.averageMegapixels >= 1.8) {
    return "dense-report";
  }
  if (profile.portraitRatio >= 0.22 || profile.gallerySignal >= 0.4) {
    return "visual-report";
  }
  if (profile.squareRatio >= 0.25 || profile.count >= 3) {
    return "boardroom-report";
  }
  return "balanced-report";
}

function inferDensityBias(profile) {
  if (profile.averageMegapixels >= 2.2 || (profile.landscapeRatio >= 0.66 && profile.highResolutionRatio >= 0.68)) {
    return "high";
  }
  if (profile.count >= 2 || profile.averageMegapixels >= 1.2) {
    return "medium";
  }
  return "low";
}

function inferHeaderStyle(profile, styleFamily) {
  if (styleFamily === "dense-report") return "boardroom-strip";
  if (styleFamily === "visual-report") return "badge-band";
  if (profile.highResolutionRatio >= 0.7) return "formal-line";
  return "formal-line";
}

function inferSummaryBandStyle(profile, styleFamily) {
  if (styleFamily === "visual-report") return "card-band";
  if (styleFamily === "boardroom-report") return "chip-band";
  if (profile.count >= 3 && profile.highResolutionRatio >= 0.6) return "accent-band";
  return "solid-left-bar";
}

function inferTablePreference(profile, densityBias, styleFamily) {
  if (densityBias === "high") return "dense";
  if (styleFamily === "visual-report") return "picture";
  if (styleFamily === "boardroom-report") return "compare";
  if (profile.squareRatio >= 0.22 && profile.count >= 2) return "dashboard";
  return "compare";
}

function inferCardStyle(styleFamily, densityBias) {
  if (styleFamily === "visual-report") return "soft-card";
  if (densityBias === "high") return "dashboard-card";
  if (styleFamily === "boardroom-report") return "ribbon-card";
  return "classic-card";
}

function inferPageRhythm(densityBias, styleFamily) {
  if (densityBias === "high") return "dense";
  if (styleFamily === "visual-report") return "visual";
  return "balanced";
}

function inferImagePlacement(profile, styleFamily) {
  if (styleFamily === "visual-report") return profile.portraitRatio >= 0.25 ? "gallery" : "hero";
  if (profile.landscapeRatio >= 0.7) return "split";
  return "hero";
}

function inferIconDiversity(profile, styleFamily) {
  if (profile.count >= 3 || styleFamily === "boardroom-report" || styleFamily === "visual-report") return "high";
  return profile.highResolutionRatio >= 0.6 ? "medium" : "low";
}

function buildTemplateBias({ styleFamily, densityBias, tablePreference, pageRhythm, imagePlacement, cardStyle }) {
  const bias = [];

  if (densityBias === "high") bias.push("table_dense_v1", "summary_dense_v1", "action_dashboard_v1");
  if (tablePreference === "dense") bias.push("table_dense_v1", "table_dashboard_v1");
  if (tablePreference === "compare") bias.push("table_compare_v1", "table_sidecallout_v1");
  if (tablePreference === "picture") bias.push("table_picture_v1", "table_visual_v1");
  if (tablePreference === "dashboard") bias.push("table_dashboard_v1", "table_matrix_v1");

  if (styleFamily === "visual-report") {
    bias.push("summary_spread_v1", "image_storyboard_v1", "process_bridge_v1");
  } else if (styleFamily === "boardroom-report") {
    bias.push("summary_bands_v1", "bullet_masonry_v1", "action_dashboard_v1");
  } else if (styleFamily === "dense-report") {
    bias.push("summary_dense_v1", "bullet_triple_v1", "process_ladder_v1");
  }

  if (pageRhythm === "visual") bias.push("summary_spread_v1", "image_gallery_v1");
  if (pageRhythm === "dense") bias.push("summary_dense_v1", "table_dense_v1");
  if (imagePlacement === "gallery") bias.push("image_gallery_v1", "image_storyboard_v1");
  if (imagePlacement === "hero") bias.push("image_focus_v1", "summary_spread_v1");
  if (cardStyle === "dashboard-card") bias.push("table_dashboard_v1", "action_dashboard_v1");

  return unique(bias);
}

function buildPreferredVariants({ styleFamily, densityBias, tablePreference, imagePlacement, cardStyle, pageRhythm }) {
  const variants = [];

  if (styleFamily === "dense-report") variants.push("dense", "dashboard", "matrix", "dense_grid", "triple", "ladder");
  if (styleFamily === "visual-report") variants.push("visual", "picture", "storyboard", "gallery", "focus", "spread");
  if (styleFamily === "boardroom-report") variants.push("compare", "sidecallout", "masonry", "stacked", "bands");
  if (styleFamily === "balanced-report") variants.push("split", "grid", "cards", "timeline");

  if (tablePreference === "dense") variants.push("dense", "dashboard", "matrix");
  if (tablePreference === "compare") variants.push("compare", "sidecallout", "highlight");
  if (tablePreference === "picture") variants.push("picture", "visual", "storyboard");
  if (tablePreference === "dashboard") variants.push("dashboard", "matrix", "compare");

  if (imagePlacement === "gallery") variants.push("gallery", "storyboard");
  if (imagePlacement === "hero") variants.push("focus", "spread");
  if (imagePlacement === "split") variants.push("split", "visual");

  if (cardStyle === "dashboard-card") variants.push("dashboard", "wall", "dense_grid");
  if (cardStyle === "ribbon-card") variants.push("bands", "compare", "timeline");
  if (cardStyle === "soft-card") variants.push("mosaic", "storyboard", "cards");

  if (pageRhythm === "dense") variants.push("dense_grid", "dense", "triple");
  if (pageRhythm === "visual") variants.push("spread", "gallery", "storyboard");
  if (densityBias === "high") variants.push("dashboard", "dense", "matrix");

  return unique(variants);
}

function analyzeReferenceImages(filePaths = []) {
  const images = [...new Set((filePaths || []).filter(Boolean))].map(safeImageMeta).filter(Boolean);
  const count = images.length;
  const landscapeCount = images.filter((item) => item.orientation === "landscape").length;
  const portraitCount = images.filter((item) => item.orientation === "portrait").length;
  const squareCount = images.filter((item) => item.orientation === "square").length;
  const highResolutionCount = images.filter((item) => item.highResolution).length;
  const averageAspectRatio = count
    ? Number((images.reduce((sum, item) => sum + (item.aspectRatio || 0), 0) / count).toFixed(3))
    : 0;
  const averageSizeKb = count
    ? Number((images.reduce((sum, item) => sum + (item.size || 0), 0) / count / 1024).toFixed(1))
    : 0;
  const averageMegapixels = count
    ? Number((images.reduce((sum, item) => sum + (item.megapixels || 0), 0) / count).toFixed(2))
    : 0;
  const landscapeRatio = ratio(landscapeCount, count);
  const portraitRatio = ratio(portraitCount, count);
  const squareRatio = ratio(squareCount, count);
  const highResolutionRatio = ratio(highResolutionCount, count);
  const gallerySignal = ratio(portraitCount + squareCount, count);

  const baseProfile = {
    count,
    averageAspectRatio,
    averageSizeKb,
    averageMegapixels,
    landscapeRatio,
    portraitRatio,
    squareRatio,
    highResolutionRatio,
    gallerySignal,
  };

  const densityBias = inferDensityBias(baseProfile);
  const styleFamily = inferStyleFamily(baseProfile);
  const headerStyle = inferHeaderStyle(baseProfile, styleFamily);
  const summaryBandStyle = inferSummaryBandStyle(baseProfile, styleFamily);
  const tablePreference = inferTablePreference(baseProfile, densityBias, styleFamily);
  const cardStyle = inferCardStyle(styleFamily, densityBias);
  const pageRhythm = inferPageRhythm(densityBias, styleFamily);
  const imagePlacement = inferImagePlacement(baseProfile, styleFamily);
  const iconDiversity = inferIconDiversity(baseProfile, styleFamily);
  const templateBias = buildTemplateBias({
    styleFamily,
    densityBias,
    tablePreference,
    pageRhythm,
    imagePlacement,
    cardStyle,
  });
  const preferredVariants = buildPreferredVariants({
    styleFamily,
    densityBias,
    tablePreference,
    imagePlacement,
    cardStyle,
    pageRhythm,
  });
  const suggestedLayoutSet = pickSuggestedLayoutSet({
    ...baseProfile,
    styleFamily,
    densityBias,
    tablePreference,
    cardStyle,
    pageRhythm,
  });

  return {
    ...baseProfile,
    styleFamily,
    densityBias,
    suggestedLayoutSet,
    headerStyle,
    summaryBandStyle,
    tablePreference,
    tableStyleBias:
      tablePreference === "dense"
        ? ["dark-header", "zebra-rows", "compact-font"]
        : tablePreference === "picture"
          ? ["light-header", "accent-highlight", "image-rail"]
          : tablePreference === "dashboard"
            ? ["dashboard-header", "metric-accent", "soft-grid"]
            : ["contrast-header", "highlight-row", "balanced-grid"],
    cardStyle,
    pageRhythm,
    imagePlacement,
    iconDiversity,
    spacingBias: pageRhythm === "dense" ? "tight" : pageRhythm === "visual" ? "open" : "balanced",
    colorMood: styleFamily === "visual-report" ? "fresh-green" : styleFamily === "boardroom-report" ? "formal-green" : "bank-green",
    repeatedPatterns:
      styleFamily === "dense-report"
        ? ["top-brand-strip", "dense-data-table", "callout-side-panel"]
        : styleFamily === "visual-report"
          ? ["hero-image", "spread-summary", "soft-card"]
          : ["top-brand-strip", "section-ribbon", "comparison-card"],
    preferredVariants,
    templateBias,
    extractionConfidence: highResolutionRatio >= 0.66 ? "high" : count >= 1 ? "medium" : "low",
    images,
    summary: count
      ? `共分析 ${count} 张参考页图片，当前风格更接近 ${styleFamily}，标题结构偏向 ${headerStyle}，表格倾向 ${tablePreference}，页面节奏偏向 ${pageRhythm}。`
      : "当前未提供参考页图片，系统将沿用默认金融汇报风格。",
  };
}

module.exports = {
  analyzeReferenceImages,
  fileToDataUrl,
};
