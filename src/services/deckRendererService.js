const fs = require("node:fs");
const path = require("node:path");
const { imageSize } = require("image-size");
const PptxGenJS = require("pptxgenjs");

const BADGE_SHAPES = [
  "roundRect",
  "ellipse",
  "diamond",
  "hexagon",
  "teardrop",
  "chevron",
  "homePlate",
  "parallelogram",
  "triangle",
  "rtTriangle",
  "octagon",
  "star5",
  "star7",
  "rightArrow",
  "leftRightArrow",
  "upArrow",
  "downArrow",
];
const RENDERABLE_IMAGE_EXTENSIONS = new Set([".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"]);
const BADGE_INNER_SHAPES = [
  "ellipse",
  "diamond",
  "hexagon",
  "triangle",
  "rtTriangle",
  "star5",
  "star7",
  "oval",
  "cloud",
];

function textLength(value) {
  return String(value || "").replace(/\s+/g, "").length;
}

function autoFontSize(text, base, min = 8, step = 24) {
  const len = textLength(text);
  const penalty = Math.max(0, Math.floor((len - 18) / step));
  return Math.max(min, base - penalty);
}

function clipText(value, maxLen = 120) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  if (!text) return "";
  if (text.length <= maxLen) return text;
  return `${text.slice(0, Math.max(16, maxLen - 1)).trim()}…`;
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function isRenderableBox(x, y, w, h) {
  return [x, y, w, h].every((value) => Number.isFinite(value)) && w > 0.02 && h > 0.02;
}

function canRenderImageFile(filePath) {
  return Boolean(filePath) && fs.existsSync(filePath) && RENDERABLE_IMAGE_EXTENSIONS.has(String(path.extname(filePath)).toLowerCase());
}

function centeredRowBoxes(count, x, y, totalWidth, height, options = {}) {
  const gap = options.gap ?? 0.16;
  const maxPerRow = options.maxPerRow || count || 1;
  const safeCount = Math.max(1, count);
  const perRow = Math.min(maxPerRow, safeCount);
  const rows = Math.ceil(safeCount / perRow);
  const boxes = [];

  for (let row = 0; row < rows; row += 1) {
    const rowCount = row === rows - 1 ? safeCount - row * perRow : perRow;
    const rowWidth = totalWidth - gap * Math.max(0, rowCount - 1);
    const boxWidth = rowWidth / Math.max(1, rowCount);
    const startX = x + (totalWidth - (boxWidth * rowCount + gap * Math.max(0, rowCount - 1))) / 2;
    for (let index = 0; index < rowCount; index += 1) {
      boxes.push({
        x: startX + index * (boxWidth + gap),
        y: y + row * (height + (options.rowGap ?? gap)),
        w: boxWidth,
        h: height,
      });
    }
  }

  return boxes;
}

function gridBoxes(count, x, y, w, h, options = {}) {
  const safeCount = Math.max(1, count);
  const cols =
    options.cols ||
    (safeCount === 1 ? 1 : safeCount === 2 ? 2 : safeCount <= 4 ? 2 : 3);
  const rows = Math.ceil(safeCount / cols);
  const gapX = options.gapX ?? 0.18;
  const gapY = options.gapY ?? 0.18;
  const boxW = (w - gapX * Math.max(0, cols - 1)) / cols;
  const boxH = (h - gapY * Math.max(0, rows - 1)) / rows;
  const boxes = [];

  for (let index = 0; index < safeCount; index += 1) {
    const row = Math.floor(index / cols);
    const col = index % cols;
    boxes.push({
      x: x + col * (boxW + gapX),
      y: y + row * (boxH + gapY),
      w: boxW,
      h: boxH,
    });
  }

  return boxes;
}

function metricBoxesForMode(count, x, y, w, mode = "compact") {
  const safeCount = Math.max(1, count);
  const maxPerRow = mode === "hero" ? 2 : mode === "balanced" ? 3 : 4;
  const height = mode === "hero" ? 0.92 : mode === "balanced" ? 0.82 : 0.74;
  const gap = mode === "hero" ? 0.2 : mode === "balanced" ? 0.18 : 0.16;
  const rowGap = mode === "hero" ? 0.18 : mode === "balanced" ? 0.16 : 0.12;
  return centeredRowBoxes(safeCount, x, y, w, height, {
    gap,
    rowGap,
    maxPerRow,
  });
}

function isNumericLike(value) {
  const text = String(value || "").trim();
  if (!text) return false;
  const compact = text.replace(/[,\s]/g, "");
  return /^[-+]?\d+(?:\.\d+)?(?:%|万元|亿元|万人|户|次|笔|个|bp|ms)?$/.test(compact);
}

function normalizeTableForRender(table) {
  const header = Array.isArray(table?.header) ? [...table.header] : [];
  const rows = Array.isArray(table?.rows) ? table.rows.map((row) => (Array.isArray(row) ? [...row] : [])) : [];
  const caption = String(table?.caption || "");
  const colCount = Math.max(header.length, ...rows.map((row) => row.length), table?.colCount || 0, 1);
  let normalizedHeader = header;
  let normalizedRows = rows;

  if (header.length === 1 && rows.length && rows[0].length > 1) {
    normalizedHeader = rows[0];
    normalizedRows = rows.slice(1);
  }

  if (!normalizedHeader.length) {
    normalizedHeader = Array.from({ length: colCount }, (_, index) => `列${index + 1}`);
  }

  if (normalizedHeader.length < colCount) {
    normalizedHeader = [...normalizedHeader, ...Array(colCount - normalizedHeader.length).fill("")];
  }

  normalizedRows = normalizedRows
    .filter((row) => row.some((cell) => String(cell || "").trim()))
    .map((row) => [...row, ...Array(Math.max(0, colCount - row.length)).fill("")]);

  return {
    caption,
    highlight: table?.highlight || "",
    header: normalizedHeader,
    rows: normalizedRows,
    colCount,
    rowCount: normalizedRows.length,
  };
}

function buildColumnWidths(tableModel, totalWidth) {
  const weights = Array.from({ length: tableModel.colCount }, (_, col) => {
    const values = [tableModel.header[col], ...tableModel.rows.map((row) => row[col])].map((item) => String(item || ""));
    const maxLen = Math.max(...values.map((item) => textLength(item)), 1);
    const numericRatio = values.filter(isNumericLike).length / Math.max(1, values.length);
    let weight = numericRatio >= 0.8 ? 0.92 : maxLen >= 18 ? 1.9 : maxLen >= 10 ? 1.45 : 1.15;
    if (col === 0 && numericRatio < 0.6) weight += 0.28;
    return weight;
  });
  const sum = weights.reduce((acc, value) => acc + value, 0) || 1;
  return weights.map((weight) => (totalWidth * weight) / sum);
}

function fitImage(imagePath, box) {
  let width = 1;
  let height = 1;
  try {
    const dims = imageSize(fs.readFileSync(imagePath));
    width = dims.width || 1;
    height = dims.height || 1;
  } catch {
    width = 1;
    height = 1;
  }
  const scale = Math.min(box.w / width, box.h / height);
  return {
    path: imagePath,
    x: box.x + (box.w - width * scale) / 2,
    y: box.y + (box.h - height * scale) / 2,
    w: width * scale,
    h: height * scale,
  };
}

function fitImageByAspect(imagePath, region) {
  let width = 1;
  let height = 1;
  try {
    const dims = imageSize(fs.readFileSync(imagePath));
    width = Math.max(1, dims.width || 1);
    height = Math.max(1, dims.height || 1);
  } catch {
    width = 1;
    height = 1;
  }
  const ratio = width / height;

  let box;
  if (ratio >= 1.18) {
    const maxHeight = Math.min(region.h, region.w / ratio);
    box = {
      x: region.x,
      y: region.y + (region.h - maxHeight) / 2,
      w: region.w,
      h: maxHeight,
    };
  } else {
    const maxWidth = Math.min(region.w, region.h * ratio);
    box = {
      x: region.x + (region.w - maxWidth) / 2,
      y: region.y,
      w: maxWidth,
      h: region.h,
    };
  }

  return fitImage(imagePath, box);
}

function getReferenceVisual(style = {}) {
  const profile = style.referenceStyleProfile || {};
  return {
    headerStyle: String(profile.headerStyle || "formal-line").toLowerCase(),
    summaryBandStyle: String(profile.summaryBandStyle || "solid-left-bar").toLowerCase(),
    tablePreference: String(profile.tablePreference || "compare").toLowerCase(),
    cardStyle: String(profile.cardStyle || "classic-card").toLowerCase(),
    pageRhythm: String(profile.pageRhythm || "balanced").toLowerCase(),
    imagePlacement: String(profile.imagePlacement || "split").toLowerCase(),
    iconDiversity: String(profile.iconDiversity || profile.iconDiversityPolicy || "medium").toLowerCase(),
  };
}

function iconCategoriesForStyle(style, fallback = ["icons", "vector-icons", "illustrations"]) {
  const visual = getReferenceVisual(style);
  if (visual.iconDiversity === "high") {
    return ["vector-icons", "icons", "illustrations"];
  }
  if (visual.iconDiversity === "low") {
    return ["vector-icons", "icons"];
  }
  return fallback;
}

function tablePaletteForStyle(style) {
  const visual = getReferenceVisual(style);
  if (visual.tablePreference === "dense") {
    return {
      headerFill: style.palette.deepGreen,
      headerText: style.palette.white,
      rowA: style.palette.lightGreen,
      rowB: style.palette.white,
      line: style.palette.softLine,
    };
  }
  if (visual.tablePreference === "picture") {
    return {
      headerFill: style.palette.oliveGreen,
      headerText: style.palette.white,
      rowA: style.palette.white,
      rowB: style.palette.subtleGray,
      line: style.palette.borderGreen,
    };
  }
  if (visual.tablePreference === "dashboard") {
    return {
      headerFill: style.palette.accentBlue,
      headerText: style.palette.white,
      rowA: style.palette.subtleGray,
      rowB: style.palette.white,
      line: style.palette.softLine,
    };
  }
  return {
    headerFill: style.palette.darkGreen,
    headerText: style.palette.white,
    rowA: style.palette.lightGreen,
    rowB: style.palette.white,
    line: style.palette.softLine,
  };
}

function getReferenceAssets(style) {
  const collections = style.materials?.assetCollections || {};
  return [
    ...(collections.iconAssets || []),
    ...(collections.brandingAssets || []),
    ...(collections.illustrationAssets || []),
    ...(collections.screenshotAssets || []),
    ...(collections.allAssets || []),
    ...(style.materials?.componentPresets?.iconAssets || []),
  ].filter(Boolean);
}

function resolveBrandAsset(style) {
  const branding = getReferenceAssets(style).find((item) => item.category === "branding" && canRenderImageFile(item.path));
  return branding?.path || style.assets?.brand || "";
}

function resolveAsset(style, tags = [], categories = []) {
  const assets = getReferenceAssets(style).filter((item) => canRenderImageFile(item?.path));
  const tagged = assets.find(
    (item) =>
      tags.some((tag) => (item.tags || []).includes(tag) || (item.usageTags || []).includes(tag)) &&
      (!categories.length || categories.includes(item.category)),
  );
  if (tagged) return tagged;
  return assets.find((item) => !categories.length || categories.includes(item.category)) || null;
}

function stableHash(value = "") {
  let hash = 0;
  const text = String(value || "");
  for (let index = 0; index < text.length; index += 1) {
    hash = (hash * 31 + text.charCodeAt(index)) % 2147483647;
  }
  return Math.abs(hash);
}

function pickDecorAsset(style, tags = [], categories = ["icons", "vector-icons", "illustrations"], seed = "") {
  const assets = getReferenceAssets(style).filter(
    (item) => canRenderImageFile(item?.path) && (!categories.length || categories.includes(item.category)),
  );
  if (!assets.length) return null;

  const tokens = [...new Set((tags || []).map((item) => String(item || "").toLowerCase()).filter(Boolean))];
  const ranked = assets.map((asset) => {
    const assetTokens = new Set([
      String(asset.category || "").toLowerCase(),
      String(asset.id || "").toLowerCase(),
      String(asset.name || "").toLowerCase(),
      ...(asset.tags || []).map((item) => String(item || "").toLowerCase()),
      ...(asset.usageTags || []).map((item) => String(item || "").toLowerCase()),
    ]);
    let score = 0;
    tokens.forEach((token) => {
      if (assetTokens.has(token)) {
        score += 6;
        return;
      }
      if ([...assetTokens].some((candidate) => candidate.includes(token) || token.includes(candidate))) {
        score += 2;
      }
    });
    const usageCounts = style.assetUsage || {};
    const usageKey = asset.id || asset.path;
    const useCount = Number(usageCounts[usageKey] || 0);
    score -= Math.min(6, useCount * 2.25);
    const categoryCounts = style.assetCategoryUsage || {};
    const categoryKey = asset.category || "other";
    const categoryCount = Number(categoryCounts[categoryKey] || 0);
    score -= Math.min(4, categoryCount * 0.65);
    if (style.lastDecorAssetId && style.lastDecorAssetId === usageKey) {
      score -= 4.5;
    }
    if (style.lastDecorCategory && style.lastDecorCategory === categoryKey) {
      score -= 2.4;
    }
    score += (stableHash(`${seed}|${asset.id || asset.path}`) % 997) / 997;
    return { asset, score };
  });

  ranked.sort((left, right) => right.score - left.score);
  const topScore = ranked[0]?.score ?? 0;
  const nearTop = ranked
    .filter((item) => topScore - item.score <= 1.18)
    .slice(0, 4);
  const pool = nearTop.length >= 2 ? nearTop : ranked.slice(0, Math.min(5, ranked.length));
  if (!pool.length) return ranked[0]?.asset || null;
  const index = stableHash(`${seed}|${tokens.join("|")}|${pool.length}`) % pool.length;
  const chosen = pool[index]?.asset || ranked[0]?.asset || null;
  if (chosen) {
    const usageCounts = style.assetUsage || (style.assetUsage = {});
    const usageKey = chosen.id || chosen.path;
    usageCounts[usageKey] = Number(usageCounts[usageKey] || 0) + 1;
    const categoryCounts = style.assetCategoryUsage || (style.assetCategoryUsage = {});
    const categoryKey = chosen.category || "other";
    categoryCounts[categoryKey] = Number(categoryCounts[categoryKey] || 0) + 1;
    style.lastDecorAssetId = usageKey;
    style.lastDecorCategory = categoryKey;
  }
  return chosen;
}

function addDecorativeIcon(slide, style, x, y, size, accent, tags = [], categories = ["icons", "vector-icons", "illustrations"]) {
  const asset = pickDecorAsset(style, tags, iconCategoriesForStyle(style, categories), `${x}|${y}|${size}|${tags.join("|")}`);
  if (!asset) return false;
  const outerShape = badgeShape(tags);
  const innerShape = badgeInnerShape(tags);
  slide.addShape(outerShape, {
    x,
    y,
    w: size,
    h: size,
    fill: { color: accent, transparency: 82 },
    line: { color: accent, transparency: 100 },
  });
  const padding = clamp(size * 0.16, 0.04, 0.14);
  slide.addShape(innerShape, {
    x: x + padding * 0.6,
    y: y + padding * 0.6,
    w: size - padding * 1.2,
    h: size - padding * 1.2,
    fill: { color: style.palette.white, transparency: 8 },
    line: { color: style.palette.white, transparency: 100 },
  });
  slide.addImage(fitImage(asset.path, { x: x + padding, y: y + padding, w: size - padding * 2, h: size - padding * 2 }));
  return true;
}

function badgeShape(tags = []) {
  const key = tags.join("|");
  let hash = 0;
  for (let index = 0; index < key.length; index += 1) {
    hash = (hash + key.charCodeAt(index)) % BADGE_SHAPES.length;
  }
  return BADGE_SHAPES[hash] || "roundRect";
}

function badgeInnerShape(tags = []) {
  const key = `${tags.join("|")}|inner`;
  let hash = 0;
  for (let index = 0; index < key.length; index += 1) {
    hash = (hash * 33 + key.charCodeAt(index)) % BADGE_INNER_SHAPES.length;
  }
  return BADGE_INNER_SHAPES[hash] || "ellipse";
}

function addText(slide, text, x, y, w, h, style, options = {}) {
  if (!isRenderableBox(x, y, w, h)) return;
  slide.addText(String(text || ""), {
    x,
    y,
    w,
    h,
    fontFace: options.fontFace || style.fonts.body,
    fontSize: options.fontSize || 10,
    bold: Boolean(options.bold),
    color: options.color || style.palette.textPrimary,
    align: options.align || "left",
    valign: options.valign || "top",
    margin: options.margin ?? 0,
    fit: options.fit || "shrink",
    breakLine: options.breakLine,
  });
}

function buildEmphasisRuns(text, style, options = {}) {
  const source = String(text || "");
  if (!source) return [{ text: "", options: {} }];

  const tokens = [
    {
      regex: /\d+(?:\.\d+)?%?/g,
      options: { bold: true, color: options.numberColor || style.palette.accentRed },
    },
    {
      regex: /(提升|下降|增长|压降|预警|监控|闭环|重点|风险|拦截|覆盖|优化|可视化|精准|提速|缩短|加强|扩大|提升至|达到|超过|新增|减少)/g,
      options: { bold: true, color: options.keywordColor || style.palette.darkGreen },
    },
  ];

  const matches = [];
  tokens.forEach((token) => {
    let match = null;
    token.regex.lastIndex = 0;
    while ((match = token.regex.exec(source))) {
      matches.push({
        start: match.index,
        end: match.index + match[0].length,
        text: match[0],
        options: token.options,
      });
    }
  });

  matches.sort((left, right) => left.start - right.start || right.end - left.end);

  const runs = [];
  let cursor = 0;
  matches.forEach((match) => {
    if (match.start < cursor) return;
    if (match.start > cursor) {
      runs.push({ text: source.slice(cursor, match.start), options: {} });
    }
    runs.push({ text: match.text, options: match.options });
    cursor = match.end;
  });
  if (cursor < source.length) {
    runs.push({ text: source.slice(cursor), options: {} });
  }

  return runs.length ? runs : [{ text: source, options: {} }];
}

function addRichText(slide, text, x, y, w, h, style, options = {}) {
  if (!isRenderableBox(x, y, w, h)) return;
  const runs = buildEmphasisRuns(text, style, options);
  slide.addText(runs, {
    x,
    y,
    w,
    h,
    fontFace: options.fontFace || style.fonts.body,
    fontSize: options.fontSize || 10,
    color: options.color || style.palette.textPrimary,
    bold: Boolean(options.bold),
    align: options.align || "left",
    valign: options.valign || "top",
    margin: options.margin ?? 0,
    fit: options.fit || "shrink",
  });
}

function addPanel(slide, x, y, w, h, style, options = {}) {
  if (!isRenderableBox(x, y, w, h)) return;
  slide.addShape("roundRect", {
    x,
    y,
    w,
    h,
    rectRadius: options.rectRadius ?? 0.03,
    fill: { color: options.fill || style.palette.white, transparency: options.fillTransparency || 0 },
    line: {
      color: options.line || style.palette.borderGreen,
      width: options.lineWidth || 0.8,
      dashType: options.dash ? "dash" : "solid",
      transparency: options.lineTransparency || 0,
    },
  });
}

function addRibbon(slide, label, x, y, w, h, style, variant = "homePlate", fillColor = "") {
  if (!isRenderableBox(x, y, w, h)) return;
  const fill = fillColor || style.palette.deepGreen;
  const shape = ["homePlate", "parallelogram", "chevron", "rightArrow"].includes(variant) ? variant : "homePlate";
  slide.addShape(shape, {
    x,
    y,
    w,
    h,
    fill: { color: fill },
    line: { color: fill, transparency: 100 },
  });
  addText(slide, label, x + 0.14, y + 0.07, w - 0.28, h - 0.14, style, {
    fontFace: style.fonts.heading,
    fontSize: 10,
    bold: true,
    color: style.palette.white,
    align: "center",
    valign: "mid",
  });
}

function addMetricBlock(slide, metric, x, y, w, h, style, accent) {
  if (!isRenderableBox(x, y, w, h)) return;
  addPanel(slide, x, y, w, h, style, { line: style.palette.borderGreen });
  addText(slide, metric.label, x + 0.08, y + 0.1, w - 0.16, 0.16, style, {
    fontSize: clamp(autoFontSize(metric.label, 9.6, 8.2, 16), 8.2, 10),
    bold: true,
    color: style.palette.textMuted,
    align: "center",
  });
  addText(slide, metric.value, x + 0.04, y + 0.3, w - 0.08, h - 0.38, style, {
    fontFace: style.fonts.heading,
    fontSize: autoFontSize(metric.value, 17.4, 11, 8),
    bold: true,
    color: accent || style.palette.darkGreen,
    align: "center",
    valign: "mid",
  });
}

function addIconBadge(slide, style, x, y, size, accent, tags = [], categories = ["icons", "vector-icons", "illustrations"]) {
  if (!isRenderableBox(x, y, size, size)) return;
  if (addDecorativeIcon(slide, style, x, y, size, accent, tags, categories)) {
    return;
  }
  const outerShape = badgeShape(tags);
  const innerShape = badgeInnerShape(tags);
  slide.addShape(outerShape, {
    x,
    y,
    w: size,
    h: size,
    fill: { color: accent },
    line: { color: accent, transparency: 100 },
  });
  slide.addShape(innerShape, {
    x: x + size * 0.24,
    y: y + size * 0.24,
    w: size * 0.52,
    h: size * 0.52,
    fill: { color: style.palette.white },
    line: { color: style.palette.white, transparency: 100 },
  });
}

function addInfoCard(slide, card, x, y, w, h, style, options = {}) {
  if (!isRenderableBox(x, y, w, h)) return;
  const accent = options.accent || style.palette.darkGreen;
  const visual = getReferenceVisual(style);
  const titleText = clipText(card.title || "", w <= 2.6 ? 18 : w <= 3.8 ? 26 : 34);
  const bodyLimit =
    h <= 0.8
      ? 28
      : h <= 1.0
        ? 42
        : h <= 1.3
          ? 68
          : h <= 1.8
            ? 110
            : 150;
  const bodyText = clipText(card.body || card.detail || "", bodyLimit);
  const iconSize = clamp(Math.min(0.38, h * 0.26), 0.24, 0.38);
  const titleFontSize = autoFontSize(titleText, w <= 2.2 ? 9.4 : 11.2, 8.0, 8);
  const titleHeight = clamp(0.2 + Math.max(0, Math.ceil(textLength(titleText) / Math.max(9, Math.floor(w * 7.5))) - 1) * 0.11, 0.2, 0.46);
  const bodyY = y + 0.16 + titleHeight;
  const iconTags = [
    ...(options.tags || []),
    style.currentTemplateVariant || "",
    style.currentSlideType || "",
    options.variantHint || "",
    accent,
    titleText,
    bodyText.slice(0, 36),
    bodyText.slice(36, 72),
  ].map((item) => String(item || "").toLowerCase()).filter(Boolean);
  const fillColor =
    options.fill ||
    (visual.cardStyle === "soft-card"
      ? style.palette.subtleGray
      : visual.cardStyle === "dashboard-card"
        ? style.palette.lightGreen
        : style.palette.white);
  addPanel(slide, x, y, w, h, style, {
    fill: fillColor,
    line: options.line || style.palette.borderGreen,
    rectRadius: visual.cardStyle === "soft-card" ? 0.05 : 0.03,
  });
  if (visual.cardStyle === "dashboard-card") {
    slide.addShape("rect", {
      x: x,
      y: y,
      w: 0.08,
      h,
      fill: { color: accent },
      line: { color: accent, transparency: 100 },
    });
  } else if (visual.cardStyle === "ribbon-card") {
    slide.addShape("rect", {
      x: x + 0.12,
      y: y + 0.08,
      w: Math.min(1.34, w * 0.42),
      h: 0.04,
      fill: { color: accent },
      line: { color: accent, transparency: 100 },
    });
  }
  addIconBadge(slide, style, x + 0.12, y + 0.12, iconSize, accent, iconTags, options.iconCategories);
  addText(slide, titleText, x + 0.18 + iconSize, y + 0.08, w - (0.32 + iconSize), titleHeight, style, {
    fontFace: style.fonts.heading,
    fontSize: titleFontSize,
    bold: true,
    color: accent,
  });
  addRichText(slide, bodyText, x + 0.18 + iconSize, bodyY, w - (0.34 + iconSize), Math.max(0.24, h - (bodyY - y) - 0.1), style, {
    fontSize: autoFontSize(bodyText, h <= 0.84 ? 8.4 : 9.6, 7.2, 18),
    color: style.palette.textPrimary,
  });
}

function addBulletBlock(slide, title, bullets, x, y, w, h, style, accent) {
  if (!isRenderableBox(x, y, w, h)) return;
  addPanel(slide, x, y, w, h, style, {
    fill: style.palette.white,
    line: style.palette.borderGreen,
  });
  addText(slide, title, x + 0.14, y + 0.1, w - 0.28, 0.2, style, {
    fontFace: style.fonts.heading,
    fontSize: autoFontSize(title, 11.4, 9.4, 8),
    bold: true,
    color: accent || style.palette.darkGreen,
  });
  const runs = (bullets || []).slice(0, 6).map((item, index, list) => ({
    text: item,
    options: {
      bullet: { indent: 10 },
      breakLine: index < list.length - 1,
    },
  }));
  slide.addText(runs, {
    x: x + 0.16,
    y: y + 0.42,
    w: w - 0.32,
    h: h - 0.52,
    fontFace: style.fonts.body,
    fontSize: autoFontSize((bullets || []).join(" "), h <= 2.2 ? 8.2 : 8.8, 7.0, 26),
    color: style.palette.textPrimary,
    fit: "shrink",
    margin: 0,
  });
}

function addGridTable(slide, table, x, y, w, h, style) {
  if (!isRenderableBox(x, y, w, h)) return;
  const model = normalizeTableForRender(table);
  const tablePalette = tablePaletteForStyle(style);
  let offsetY = y;
  let remainingHeight = h;

  if (model.caption) {
    const captionH = clamp(h * 0.12, 0.24, 0.34);
    slide.addShape("rect", {
      x,
      y,
      w,
      h: captionH,
      fill: { color: tablePalette.headerFill },
      line: { color: tablePalette.line, width: 0.5 },
    });
    addText(slide, model.caption, x + 0.06, y + 0.04, w - 0.12, captionH - 0.08, style, {
      fontSize: autoFontSize(model.caption, 8.6, 7.2, 12),
      bold: true,
      color: tablePalette.headerText,
      align: "left",
      valign: "mid",
    });
    offsetY += captionH;
    remainingHeight -= captionH;
  }

  const colWidths = buildColumnWidths(model, w);
  const headerH = clamp(remainingHeight * 0.14, 0.28, 0.38);
  const visibleRows = model.rows.slice(0, Math.max(1, Math.min(model.rows.length, Math.floor((remainingHeight - headerH) / 0.34) || model.rows.length)));
  const rowH = clamp((remainingHeight - headerH) / Math.max(1, visibleRows.length), 0.22, 0.42);
  const headerFont = clamp(8.6 - Math.max(0, model.colCount - 4) * 0.4, 6.8, 8.6);
  const bodyFont = clamp(8.2 - Math.max(0, model.colCount - 4) * 0.3 - Math.max(0, visibleRows.length - 5) * 0.12, 6.6, 8.2);

  let currentX = x;
  for (let col = 0; col < model.colCount; col += 1) {
    const colW = colWidths[col];
    slide.addShape("rect", {
      x: currentX,
      y: offsetY,
      w: colW,
      h: headerH,
      fill: { color: tablePalette.headerFill },
      line: { color: tablePalette.line, width: 0.5 },
    });
    addText(slide, model.header[col] || "", currentX + 0.03, offsetY + 0.04, colW - 0.06, headerH - 0.08, style, {
      fontSize: headerFont,
      bold: true,
      color: tablePalette.headerText,
      align: isNumericLike(model.header[col]) ? "right" : "center",
      valign: "mid",
    });
    currentX += colW;
  }

  visibleRows.forEach((row, rowIndex) => {
    const rowY = offsetY + headerH + rowIndex * rowH;
    const fill = rowIndex % 2 === 0 ? tablePalette.rowA : tablePalette.rowB;
    const highlight = model.highlight && row.some((cell) => String(cell || "").includes(model.highlight));
    let cellX = x;
    for (let col = 0; col < model.colCount; col += 1) {
      const colW = colWidths[col];
      slide.addShape("rect", {
        x: cellX,
        y: rowY,
        w: colW,
        h: rowH,
        fill: { color: fill },
        line: { color: tablePalette.line, width: 0.5 },
      });
      addText(slide, row[col] || "", cellX + 0.03, rowY + 0.03, colW - 0.06, rowH - 0.06, style, {
        fontSize: bodyFont,
        bold: highlight,
        color: highlight ? style.palette.accentRed : style.palette.textPrimary,
        align: isNumericLike(row[col]) ? "right" : col === 0 ? "left" : "center",
        valign: "mid",
      });
      cellX += colW;
    }
  });
}

function addBars(slide, bars, x, y, w, h, style, title = "横向对比") {
  if (!isRenderableBox(x, y, w, h) || h < 0.28) return;
  addPanel(slide, x, y, w, h, style, { line: style.palette.softLine, dash: true });
  addText(slide, title, x + 0.12, y + 0.1, w - 0.24, 0.16, style, {
    fontFace: style.fonts.heading,
    fontSize: 8.8,
    bold: true,
  });

  const list = (bars || []).filter((item) => item.value > 0).slice(0, 5);
  if (!list.length) {
    addText(slide, "暂无可展示数据", x + 0.14, y + 0.34, w - 0.28, Math.max(0.12, h - 0.38), style, {
      fontSize: 7.2,
      color: style.palette.textMuted,
      align: "center",
      valign: "mid",
    });
    return;
  }
  const max = Math.max(...list.map((item) => item.value), 1);
  const maxRows = Math.max(1, Math.min(list.length, Math.floor((h - 0.38) / 0.12)));
  const visible = list.slice(0, maxRows);
  const rowGap = visible.length >= 4 ? 0.03 : 0.05;
  const rowH = clamp((h - 0.38 - rowGap * Math.max(0, visible.length - 1)) / visible.length, 0.06, 0.14);
  const labelW = clamp(Math.max(0.72, Math.min(1.18, w * 0.18)), 0.72, 1.18);
  const valueW = clamp(Math.max(0.28, Math.min(0.44, w * 0.08)), 0.28, 0.44);
  const barX = x + labelW + 0.12;
  const barW = Math.max(0.14, w - labelW - valueW - 0.22);
  const baseY = y + 0.32;

  visible.forEach((item, index) => {
    const rowY = baseY + index * (rowH + rowGap);
    addText(slide, item.name, x + 0.12, rowY + 0.01, labelW - 0.12, rowH, style, {
      fontSize: clamp(6.8 - Math.max(0, visible.length - 3) * 0.15, 6.2, 7.1),
      color: style.palette.textPrimary,
      valign: "mid",
    });
    slide.addShape("rect", {
      x: barX,
      y: rowY + Math.max(0.02, rowH * 0.34),
      w: barW,
      h: Math.max(0.04, rowH * 0.24),
      fill: { color: style.palette.lightGreen },
      line: { color: style.palette.lightGreen, transparency: 100 },
    });
    slide.addShape("rect", {
      x: barX,
      y: rowY + Math.max(0.02, rowH * 0.34),
      w: Math.max(0.14, (barW * item.value) / max),
      h: Math.max(0.04, rowH * 0.24),
      fill: { color: index === 0 ? style.palette.accentOrange : style.palette.oliveGreen },
      line: { color: style.palette.white, transparency: 100 },
    });
    addText(slide, item.display || String(item.value), x + w - valueW, rowY, valueW - 0.04, rowH, style, {
      fontSize: clamp(6.8, 6.2, 7.1),
      bold: true,
      color: style.palette.accentRed,
      align: "right",
      valign: "mid",
    });
  });
}

function addBrandHeader(slide, title, style) {
  const visual = getReferenceVisual(style);
  slide.background = { color: style.palette.white };
  addText(slide, title, 0.78, 0.18, 8.8, 0.4, style, {
    fontFace: style.fonts.heading,
    fontSize: autoFontSize(title, 18, 14, 10),
    bold: true,
    color: style.palette.titleBlack,
  });
  const brand = resolveBrandAsset(style);
  if (brand && fs.existsSync(brand)) {
    slide.addImage({ path: brand, x: 10.35, y: 0.18, w: 2.35, h: 0.38 });
  }
  if (visual.headerStyle === "boardroom-strip") {
    slide.addShape("rect", {
      x: 0.72,
      y: 0.6,
      w: 12.12,
      h: 0.05,
      fill: { color: style.palette.deepGreen },
      line: { color: style.palette.deepGreen, transparency: 100 },
    });
    slide.addShape("rect", {
      x: 11.2,
      y: 0.14,
      w: 0.32,
      h: 0.18,
      fill: { color: style.palette.oliveGreen },
      line: { color: style.palette.oliveGreen, transparency: 100 },
    });
    slide.addShape("rect", {
      x: 11.54,
      y: 0.14,
      w: 0.32,
      h: 0.18,
      fill: { color: style.palette.accentLime },
      line: { color: style.palette.accentLime, transparency: 100 },
    });
    slide.addShape("rect", {
      x: 11.88,
      y: 0.14,
      w: 0.32,
      h: 0.18,
      fill: { color: style.palette.lightGreen },
      line: { color: style.palette.lightGreen, transparency: 100 },
    });
  } else if (visual.headerStyle === "badge-band") {
    slide.addShape("roundRect", {
      x: 11.08,
      y: 0.2,
      w: 1.58,
      h: 0.24,
      rectRadius: 0.04,
      fill: { color: style.palette.deepGreen },
      line: { color: style.palette.deepGreen, transparency: 100 },
    });
    slide.addShape("rect", {
      x: 0.72,
      y: 0.62,
      w: 12.12,
      h: 0.04,
      fill: { color: style.palette.darkGreen },
      line: { color: style.palette.darkGreen, transparency: 100 },
    });
  } else {
    slide.addShape("rect", {
      x: 0.72,
      y: 0.62,
      w: 12.12,
      h: 0.04,
      fill: { color: style.palette.deepGreen },
      line: { color: style.palette.deepGreen, transparency: 100 },
    });
  }
  slide.addShape("rect", {
    x: 0,
    y: 6.52,
    w: style.page.widthInches,
    h: 0.16,
    fill: { color: style.palette.darkGreen },
    line: { color: style.palette.darkGreen, transparency: 100 },
  });
}

function addSummaryBand(slide, text, style, y = 0.78, h = 0.7) {
  const visual = getReferenceVisual(style);
  if (visual.summaryBandStyle === "card-band") {
    addPanel(slide, 0.72, y, 12.12, h, style, {
      fill: style.palette.subtleGray,
      line: style.palette.borderGreen,
      rectRadius: 0.06,
    });
    addText(slide, text, 1.0, y + 0.1, 11.52, h - 0.18, style, {
      fontSize: autoFontSize(text, 10.8, 8.6, 22),
      bold: true,
    });
    return;
  }
  if (visual.summaryBandStyle === "chip-band") {
    slide.addShape("rect", {
      x: 0.72,
      y,
      w: 12.12,
      h,
      fill: { color: style.palette.lightGreen },
      line: { color: style.palette.lightGreen, transparency: 100 },
    });
    slide.addShape("roundRect", {
      x: 0.92,
      y: y + 0.14,
      w: 2.2,
      h: 0.24,
      rectRadius: 0.04,
      fill: { color: style.palette.deepGreen },
      line: { color: style.palette.deepGreen, transparency: 100 },
    });
    addText(slide, text, 1.18, y + 0.12, 11.08, h - 0.2, style, {
      fontSize: autoFontSize(text, 10.6, 8.6, 22),
      bold: true,
    });
    return;
  }
  slide.addShape("rect", {
    x: 0.72,
    y,
    w: 12.12,
    h,
    fill: { color: visual.summaryBandStyle === "accent-band" ? style.palette.subtleGray : style.palette.lightGreen },
    line: { color: style.palette.lightGreen, transparency: 100 },
  });
  slide.addShape("rect", {
    x: 0.72,
    y,
    w: 0.28,
    h,
    fill: { color: style.palette.deepGreen },
    line: { color: style.palette.deepGreen, transparency: 100 },
  });
  if (visual.summaryBandStyle === "accent-band") {
    slide.addShape("rect", {
      x: 1.02,
      y: y + 0.12,
      w: 10.9,
      h: 0.04,
      fill: { color: style.palette.oliveGreen },
      line: { color: style.palette.oliveGreen, transparency: 100 },
    });
  }
  addText(slide, text, 1.06, y + 0.1, 11.4, h - 0.18, style, {
    fontSize: autoFontSize(text, 11, 8.8, 22),
    bold: true,
  });
}

function addPageNumber(slide, page, style) {
  addText(slide, String(page).padStart(2, "0"), 11.48, 6.52, 0.42, 0.18, style, {
    fontSize: 10.2,
    bold: true,
    color: style.palette.darkGreen,
    align: "right",
  });
}

function resolveTemplate(style, slideDef) {
  const templates = style.layoutLibrary?.templates || {};
  const defaultId = style.layoutLibrary?.set?.[slideDef.type] || style.layoutLibrary?.defaultsByType?.[slideDef.type];
  const templateId = slideDef.templateId || defaultId || "";
  return {
    id: templateId,
    pageType: slideDef.type,
    variant: templates[templateId]?.variant || "default",
    displayName: templates[templateId]?.displayName || templateId,
  };
}

function renderCover(slide, slideDef, style, template) {
  slide.background = { color: style.palette.white };
  const brand = resolveBrandAsset(style);
  if (brand && fs.existsSync(brand)) {
    slide.addImage({ path: brand, x: 9.52, y: 0.48, w: 2.8, h: 0.42 });
  }

  if (template.variant === "centered") {
    slide.addShape("rect", {
      x: 4.42,
      y: 1.42,
      w: 3.96,
      h: 0.04,
      fill: { color: style.palette.deepGreen },
      line: { color: style.palette.deepGreen, transparency: 100 },
    });
    addText(slide, slideDef.title, 1.24, 1.84, 10.3, 1.02, style, {
      fontFace: style.fonts.heading,
      fontSize: autoFontSize(slideDef.title, 24, 18, 8),
      bold: true,
      color: style.palette.deepGreen,
      align: "center",
      valign: "mid",
    });
  } else {
    slide.addShape("rect", {
      x: 1.18,
      y: template.variant === "clean" ? 1.36 : 1.54,
      w: 10.3,
      h: 0.04,
      fill: { color: style.palette.deepGreen },
      line: { color: style.palette.deepGreen, transparency: 100 },
    });
    addText(slide, slideDef.title, 1.1, 1.9, 10.5, 1.0, style, {
      fontFace: style.fonts.heading,
      fontSize: autoFontSize(slideDef.title, template.variant === "clean" ? 24 : 26, 18, 8),
      bold: true,
      color: style.palette.deepGreen,
      align: "center",
      valign: "mid",
    });
  }

  if (slideDef.subtitle) {
    addText(slide, slideDef.subtitle, 2.0, 2.92, 8.8, 0.3, style, {
      fontSize: 12.6,
      bold: true,
      color: style.palette.textMuted,
      align: "center",
    });
  }

  slide.addShape("rect", {
    x: 4.46,
    y: 3.24,
    w: 3.84,
    h: 0.03,
    fill: { color: style.palette.borderGreen },
    line: { color: style.palette.borderGreen, transparency: 100 },
  });

  [
    `部门：${slideDef.department || style.department || ""}`,
    `汇报人：${slideDef.presenter || style.presenter || ""}`,
    `日期：${slideDef.reportDate || style.reportDate || ""}`,
  ].forEach((line, index) => {
    addText(slide, line, 4.0, 3.56 + index * 0.42, 4.7, 0.24, style, {
      fontSize: 12.2,
      bold: index < 2,
      color: style.palette.deepGreen,
      align: "center",
    });
  });
}

function renderSummaryCards(slide, slideDef, style, template) {
  addBrandHeader(slide, slideDef.title, style);
  addSummaryBand(slide, slideDef.headline, style);

  const metrics = (slideDef.metrics || []).slice(
    0,
    template.variant === "bands" ? 2 : slideDef.density === "low" ? 3 : 4,
  );
  const cards = (slideDef.cards || []).slice(0, template.variant === "dense_grid" ? 6 : 4);
  const takeaways = (slideDef.takeaways || []).slice(0, 3);
  const metricMode =
    template.variant === "bands"
      ? "hero"
      : template.variant === "dense_grid" || slideDef.density === "high"
        ? "compact"
        : slideDef.density === "low"
          ? "hero"
          : "balanced";
  const metricBoxes = metricBoxesForMode(metrics.length, 0.82, 1.58, 11.0, metricMode);
  const metricColors =
    metricMode === "hero"
      ? [style.palette.accentBlue, style.palette.oliveGreen, style.palette.accentOrange, style.palette.darkGreen]
      : metricMode === "balanced"
        ? [style.palette.darkGreen, style.palette.oliveGreen, style.palette.accentOrange, style.palette.accentBlue]
        : [style.palette.accentRed, style.palette.accentOrange, style.palette.oliveGreen, style.palette.darkGreen];

  metrics.forEach((metric, index) => {
    addMetricBlock(
      slide,
      metric,
      metricBoxes[index].x,
      metricBoxes[index].y,
      metricBoxes[index].w,
      metricBoxes[index].h,
      style,
      metricColors[index % metricColors.length],
    );
  });

  const cardsTop = metrics.length ? metricBoxes[metricBoxes.length - 1].y + metricBoxes[metricBoxes.length - 1].h + 0.24 : 1.72;
  const cardsHeight = clamp(5.84 - cardsTop, 2.1, 3.86);

  if (template.variant === "spread") {
    const heroLeft = cards[0];
    const heroRight = cards[1];
    if (heroLeft) {
      addInfoCard(slide, heroLeft, 0.82, cardsTop, 6.22, 1.52, style, {
        accent: style.palette[heroLeft.accent] || style.palette.darkGreen,
        fill: style.palette.white,
        tags: ["summary", "spread", "hero-left"],
      });
    }
    if (heroRight) {
      addInfoCard(slide, heroRight, 7.22, cardsTop + 0.02, 4.6, 1.48, style, {
        accent: style.palette[heroRight.accent] || style.palette.darkGreen,
        fill: style.palette.lightGreen,
        tags: ["summary", "spread", "hero-right"],
      });
    }
    const lowerCards = cards.slice(2, 4);
    const bottomBoxes = gridBoxes(Math.max(1, lowerCards.length || 1), 0.82, cardsTop + 1.72, 11.0, 1.2, {
      cols: Math.max(1, lowerCards.length || 1),
      gapX: 0.22,
    });
    lowerCards.forEach((card, index) => {
      addInfoCard(slide, card, bottomBoxes[index].x, bottomBoxes[index].y, bottomBoxes[index].w, bottomBoxes[index].h, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.subtleGray : style.palette.white,
        tags: ["summary", "spread", `lower-${index}`],
      });
    });
    if (cards[4]) {
      addRibbon(slide, cards[4].title || "关键结论", 0.96, cardsTop + 3.12, 2.7, 0.34, style, "chevron", style.palette.oliveGreen);
      addText(slide, cards[4].body || "", 3.84, cardsTop + 3.08, 7.24, 0.22, style, {
        fontSize: autoFontSize(cards[4].body || "", 10.2, 8.2, 12),
        color: style.palette.darkGreen,
        bold: true,
      });
    }
  } else if (template.variant === "mosaic") {
    const boxes = gridBoxes(cards.length || 1, 0.82, cardsTop, 11.0, cardsHeight, {
      cols: cards.length >= 5 ? 3 : 2,
      gapX: 0.18,
      gapY: 0.16,
    });
    cards.forEach((card, index) => {
      const short = index % 3 === 0;
      addInfoCard(slide, card, boxes[index].x, boxes[index].y, boxes[index].w, short ? boxes[index].h * 0.9 : boxes[index].h, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["summary", "mosaic", `card-${index}`],
      });
    });
  } else if (template.variant === "bands") {
    const boxes = centeredRowBoxes(cards.length || 1, 0.82, cardsTop, 11.0, clamp(cardsHeight / Math.max(1, cards.length), 0.7, 1.0), {
      gap: 0.14,
      maxPerRow: 1,
    });
    cards.forEach((card, index) => {
      addInfoCard(slide, card, boxes[index].x, boxes[index].y, boxes[index].w, boxes[index].h, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["summary", `band-${index}`],
      });
    });
  } else if (template.variant === "dense_grid") {
    const cols = cards.length >= 5 ? 3 : 2;
    const boxes = gridBoxes(cards.length || 1, 0.82, cardsTop, 11.0, cardsHeight, {
      cols,
      gapX: 0.18,
      gapY: 0.18,
    });
    cards.forEach((card, index) => {
      const row = Math.floor(index / cols);
      const col = index % cols;
      addInfoCard(slide, card, boxes[index].x, boxes[index].y, boxes[index].w, boxes[index].h, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: (row + col) % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["summary", "dense", `grid-${index}`],
      });
    });
  } else {
    const boxes = gridBoxes(cards.length || 1, 0.82, cardsTop, 11.0, cardsHeight, {
      cols: cards.length === 1 ? 1 : 2,
      gapX: 0.26,
      gapY: 0.22,
    });
    cards.forEach((card, index) => {
      addInfoCard(slide, card, boxes[index].x, boxes[index].y, boxes[index].w, boxes[index].h, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        tags: ["summary", `grid-${index}`],
      });
    });
  }

  const railCards =
    takeaways.length >= 3
      ? takeaways.slice(0, 3)
      : [
          takeaways[0] || {
            title: "核心结论",
            body: slideDef.footer || "围绕关键结论继续细化为可执行动作。",
            accent: "darkGreen",
          },
          takeaways[1] || {
            title: "推进重点",
            body: takeaways[0]?.body || "把重点事项拆成短周期动作。",
            accent: "accentOrange",
          },
          takeaways[2] || {
            title: "落地节奏",
            body: "建立跟踪、复盘和迭代机制。",
            accent: "oliveGreen",
          },
        ].slice(0, 3);

  if (template.variant !== "wall" && railCards.length) {
    const railY = template.variant === "closing" ? 5.62 : 5.02;
    const railH = template.variant === "closing" ? 0.7 : 0.76;
    const railBoxes = gridBoxes(Math.min(3, railCards.length), 0.82, railY, 11.0, railH, {
      cols: Math.min(3, railCards.length),
      gapX: 0.14,
    });
    railCards.forEach((card, index) => {
      addInfoCard(slide, card, railBoxes[index].x, railBoxes[index].y, railBoxes[index].w, railBoxes[index].h, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index === 1 ? style.palette.lightGreen : style.palette.white,
        tags: ["takeaway", "rail", `card-${index}`],
      });
    });
  }

  addPageNumber(slide, slideDef.page, style);
}

function renderTableAnalysis(slide, slideDef, style, template) {
  addBrandHeader(slide, slideDef.title, style);
  addSummaryBand(slide, slideDef.headline, style);

  const tableModel = normalizeTableForRender(slideDef.table || {});
  const screenshots = Array.isArray(slideDef.screenshots) ? slideDef.screenshots.filter((item) => item?.path && fs.existsSync(item.path)) : [];
  const imagePath = slideDef.image?.path && fs.existsSync(slideDef.image.path) ? slideDef.image.path : "";
  const previewAsset = screenshots[0] || (imagePath ? slideDef.image : null);
  const previewPath = previewAsset?.path || "";
  const previewAspect = Number(previewAsset?.aspectRatio || slideDef.image?.aspectRatio || 0);
  const hasPreview = Boolean(previewPath);
  const hasBars = Boolean((slideDef.bars || []).filter((item) => Number(item?.value || 0) > 0).length);
  const denseTableSignal = tableModel.rowCount >= 6 || tableModel.colCount >= 5 || (tableModel.rowCount >= 4 && tableModel.colCount >= 4);
  const requestedVariant = String(template.variant || "").toLowerCase();
  const insights = (slideDef.insights || []).slice(0, 4);
  const fallbackCards = [
    { title: slideDef.table?.highlight || "关键判断", body: slideDef.footer || slideDef.headline || "请聚焦本页的主要发现。", accent: "accentOrange" },
    { title: "补充说明", body: slideDef.summary || "结合表格和样本情况补充说明。", accent: "darkGreen" },
    { title: "后续动作", body: "将核心结论转成后续动作或跟进建议。", accent: "oliveGreen" },
  ];
  const cards = (insights.length ? insights : fallbackCards).slice(0, 3);

  let effectiveVariant = requestedVariant || "compare";
  if (["dense", "dashboard"].includes(effectiveVariant) && !denseTableSignal) effectiveVariant = hasPreview ? "compare" : "sidecallout";
  if (["visual", "picture"].includes(effectiveVariant) && !hasPreview && hasBars) effectiveVariant = "compare";
  if (!["dashboard", "dense", "compare", "sidecallout", "visual", "picture", "highlight", "matrix", "stack"].includes(effectiveVariant)) {
    effectiveVariant = denseTableSignal ? "dense" : hasPreview ? "compare" : "sidecallout";
  }

  const metricLimit =
    effectiveVariant === "dense" || effectiveVariant === "dashboard" ? 2 :
    effectiveVariant === "compare" ? 2 :
    effectiveVariant === "sidecallout" ? 1 :
    effectiveVariant === "visual" || effectiveVariant === "picture" ? 1 :
    2;
  const metrics = (slideDef.metrics || []).slice(0, metricLimit);
  const metricMode =
    effectiveVariant === "dense" || effectiveVariant === "dashboard" ? "compact" :
    effectiveVariant === "sidecallout" || effectiveVariant === "visual" || effectiveVariant === "picture" ? "hero" :
    "balanced";
  const metricBoxes = metricBoxesForMode(metrics.length, 0.82, 1.58, 11.0, metricMode);
  const metricColors =
    metricMode === "hero"
      ? [style.palette.accentBlue, style.palette.oliveGreen, style.palette.accentOrange]
      : metricMode === "balanced"
        ? [style.palette.darkGreen, style.palette.oliveGreen, style.palette.accentOrange]
        : [style.palette.accentRed, style.palette.accentOrange, style.palette.darkGreen];
  metrics.forEach((metric, index) => {
    addMetricBlock(
      slide,
      metric,
      metricBoxes[index].x,
      metricBoxes[index].y,
      metricBoxes[index].w,
      metricBoxes[index].h,
      style,
      metricColors[index % metricColors.length],
    );
  });

  const contentY = metrics.length ? metricBoxes[metricBoxes.length - 1].y + metricBoxes[metricBoxes.length - 1].h + 0.18 : 1.68;
  const availableHeight = clamp(5.9 - contentY, 3.0, 4.2);
  const tableY = contentY + 0.46;

  const renderCards = (x, y, w, h, cardList, opts = {}) => {
    const items = (cardList || []).slice(0, 3);
    if (!items.length || h <= 0.32) return;
    const boxes = gridBoxes(items.length, x, y, w, h, {
      cols: 1,
      gapY: opts.gapY || 0.12,
    });
    items.forEach((card, index) => {
      addInfoCard(slide, card, boxes[index].x, boxes[index].y, boxes[index].w, boxes[index].h, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index % 2 === 1 ? style.palette.lightGreen : style.palette.white,
        tags: ["table", effectiveVariant, `card-${index}`],
      });
    });
  };

  if (effectiveVariant === "dense" || effectiveVariant === "dashboard") {
    addRibbon(slide, tableModel.caption || "核心数据表", 0.82, contentY, 3.18, 0.4, style, "homePlate", style.palette.deepGreen);
    const footerCards = cards.slice(0, denseTableSignal ? 2 : 3);
    const footerH = footerCards.length ? (denseTableSignal ? 0.86 : 1.02) : 0;
    const tableH = availableHeight - footerH;
    addPanel(slide, 0.82, tableY, 11.0, tableH, style, { line: style.palette.softLine });
    addGridTable(slide, tableModel, 0.96, tableY + 0.12, 10.72, Math.max(1.7, tableH - 0.24), style);
    if (hasBars && tableH >= 2.7) {
      addBars(slide, slideDef.bars || [], 0.96, tableY + tableH - 0.68, 10.72, 0.54, style, "横向趋势");
    }
    if (footerCards.length) {
      const boxes = gridBoxes(footerCards.length, 0.82, tableY + tableH + 0.08, 11.0, footerH - 0.08, {
        cols: footerCards.length >= 3 ? 3 : footerCards.length,
        gapX: 0.16,
      });
      footerCards.forEach((card, index) => {
        addInfoCard(slide, card, boxes[index].x, boxes[index].y, boxes[index].w, boxes[index].h, style, {
          accent: style.palette[card.accent] || style.palette.darkGreen,
          fill: index === 1 ? style.palette.lightGreen : style.palette.white,
          tags: ["table", effectiveVariant, `footer-${index}`],
        });
      });
    }
  } else if (effectiveVariant === "compare" || effectiveVariant === "sidecallout") {
    addRibbon(slide, tableModel.caption || "核心数据表", 0.82, contentY, 3.0, 0.4, style, "chevron", style.palette.deepGreen);
    addRibbon(slide, "结论判断", 8.5, contentY, 2.5, 0.4, style, "rightArrow", style.palette.accentBlue);
    const leftW = effectiveVariant === "sidecallout" ? 7.28 : 6.86;
    const gap = 0.22;
    const rightX = 0.82 + leftW + gap;
    const rightW = 11.82 - rightX;
    addPanel(slide, 0.82, tableY, leftW, availableHeight, style, { line: style.palette.softLine });
    const barReserve = hasBars ? 0.78 : 0;
    addGridTable(slide, tableModel, 0.96, tableY + 0.12, leftW - 0.28, Math.max(1.5, availableHeight - 0.24 - barReserve), style);
    if (hasBars && availableHeight > 1.5) {
      addBars(slide, slideDef.bars || [], 0.96, tableY + availableHeight - 0.66, leftW - 0.28, 0.5, style, "横向趋势");
    }
    addPanel(slide, rightX, tableY, rightW, availableHeight, style, { line: style.palette.softLine, fill: style.palette.white });
    let cardStartY = tableY + 0.12;
    let cardHeight = availableHeight - 0.24;
    if (hasPreview) {
      const previewH = clamp(previewAspect >= 1.45 ? 1.22 : 0.98, 0.98, 1.28);
      slide.addImage(fitImageByAspect(previewPath, { x: rightX + 0.14, y: tableY + 0.14, w: rightW - 0.28, h: previewH }));
      cardStartY += previewH + 0.12;
      cardHeight -= previewH + 0.12;
      renderCards(rightX + 0.1, cardStartY, rightW - 0.2, cardHeight, cards.slice(0, 2));
    } else {
      const focusTitle = slideDef.table?.highlight || slideDef.footer || slideDef.headline || "重点结论";
      addText(slide, focusTitle, rightX + 0.16, tableY + 0.14, rightW - 0.32, 0.28, style, {
        fontFace: style.fonts.heading,
        fontSize: autoFontSize(focusTitle, 10.8, 8.8, 16),
        bold: true,
        color: style.palette.accentOrange,
      });
      const detailText = slideDef.summary || cards.map((item) => item.body || item.detail || "").filter(Boolean).join(" ");
      addRichText(slide, detailText || "请结合表格和分行反馈，形成关键判断与后续动作。", rightX + 0.16, tableY + 0.46, rightW - 0.32, 0.62, style, {
        fontSize: autoFontSize(detailText || "", 8.6, 7.6, 22),
      });
      cardStartY = tableY + 1.18;
      cardHeight = availableHeight - 1.3;
      renderCards(rightX + 0.1, cardStartY, rightW - 0.2, cardHeight, cards.slice(0, 3));
    }
  } else if (effectiveVariant === "visual" || effectiveVariant === "picture") {
    addRibbon(slide, tableModel.caption || "表格与图像", 0.82, contentY, 2.96, 0.4, style, "parallelogram", style.palette.oliveGreen);
    addRibbon(slide, "结论判断", 8.32, contentY, 2.84, 0.4, style, "rightArrow", style.palette.accentBlue);
    const tableW = hasPreview ? 5.32 : 6.78;
    const mediaW = hasPreview ? clamp(previewAspect >= 1.6 ? 3.1 : 2.74, 2.5, 3.2) : 0;
    const gap = 0.18;
    const mediaX = 0.82 + tableW + gap;
    const cardX = hasPreview ? mediaX + mediaW + gap : 0.82 + tableW + gap;
    const cardW = 11.82 - cardX;
    addPanel(slide, 0.82, tableY, tableW, availableHeight, style, { line: style.palette.softLine });
    addGridTable(slide, tableModel, 0.96, tableY + 0.12, tableW - 0.28, Math.min(availableHeight - 0.24, hasBars ? 2.18 : 2.46), style);
    if (hasBars) {
      addBars(slide, slideDef.bars || [], 0.96, tableY + availableHeight - 0.76, tableW - 0.28, 0.58, style, "横向趋势");
    }
    if (hasPreview) {
      addPanel(slide, mediaX, tableY, mediaW, availableHeight, style, { line: style.palette.softLine, fill: style.palette.subtleGray });
      slide.addImage(fitImageByAspect(previewPath, { x: mediaX + 0.08, y: tableY + 0.08, w: mediaW - 0.16, h: availableHeight - 0.16 }));
    }
    if (!hasPreview && hasBars) {
      const insightTitle = slideDef.table?.highlight || "补充对比";
      addPanel(slide, cardX, tableY, cardW, availableHeight, style, { line: style.palette.softLine, fill: style.palette.white });
      addText(slide, insightTitle, cardX + 0.16, tableY + 0.16, cardW - 0.32, 0.2, style, {
        fontFace: style.fonts.heading,
        fontSize: 10.2,
        bold: true,
        color: style.palette.accentOrange,
      });
      addBars(slide, slideDef.bars || [], cardX + 0.12, tableY + 0.48, cardW - 0.24, 1.02, style, "横向趋势");
      renderCards(cardX + 0.08, tableY + 1.64, cardW - 0.16, availableHeight - 1.72, cards.slice(0, 2));
    } else {
      renderCards(cardX, tableY, cardW, availableHeight, cards.slice(0, hasPreview ? 2 : 3));
    }
  } else {
    addRibbon(slide, tableModel.caption || "重点结果与表格", 0.82, contentY, 3.0, 0.4, style, "homePlate", style.palette.deepGreen);
    const railW = 1.86;
    const tableX = 0.82 + railW + 0.18;
    const tableW = 6.16;
    const rightX = tableX + tableW + 0.18;
    const rightW = 11.82 - rightX;
    addPanel(slide, 0.82, tableY, railW, availableHeight, style, { fill: style.palette.darkGreen, line: style.palette.darkGreen });
    addText(slide, slideDef.table?.highlight || "重点对象", 0.96, tableY + 0.14, railW - 0.28, 0.2, style, {
      fontFace: style.fonts.heading,
      fontSize: 10.2,
      bold: true,
      color: style.palette.white,
      align: "center",
    });
    const leftMetrics = metrics.slice(0, 2);
    const metricRailBoxes = centeredRowBoxes(Math.max(1, leftMetrics.length), 0.94, tableY + 0.52, railW - 0.24, 0.58, {
      gap: 0.08,
      rowGap: 0.16,
      maxPerRow: 1,
    });
    leftMetrics.forEach((metric, index) => {
      addText(slide, metric.label, metricRailBoxes[index].x, metricRailBoxes[index].y, metricRailBoxes[index].w, 0.12, style, {
        fontSize: 6.9,
        color: style.palette.white,
        align: "center",
      });
      addText(slide, metric.value, metricRailBoxes[index].x, metricRailBoxes[index].y + 0.16, metricRailBoxes[index].w, 0.18, style, {
        fontFace: style.fonts.heading,
        fontSize: 11.2,
        bold: true,
        color: index === 0 ? style.palette.accentLime : style.palette.white,
        align: "center",
      });
    });
    addPanel(slide, tableX, tableY, tableW, availableHeight, style, { line: style.palette.softLine });
    addGridTable(slide, tableModel, tableX + 0.12, tableY + 0.12, tableW - 0.24, availableHeight - 0.24, style);
    addPanel(slide, rightX, tableY, rightW, availableHeight, style, { line: style.palette.softLine, fill: style.palette.white });
    renderCards(rightX + 0.08, tableY + 0.08, rightW - 0.16, availableHeight - 0.16, cards.slice(0, 3));
  }

  addPageNumber(slide, slideDef.page, style);
}

function renderProcessFlow(slide, slideDef, style, template) {
  addBrandHeader(slide, slideDef.title, style);
  addSummaryBand(slide, slideDef.headline || slideDef.summary || slideDef.footer || "", style);
  const stages = (slideDef.stages || []).slice(0, 5);
  const notes = (slideDef.notes || slideDef.callouts || []).slice(0, 4);
  const stepBullets = stages
    .map((stage) => stage.detail || stage.body || stage.title || "")
    .filter(Boolean)
    .slice(0, 4);
  const previewImages = [slideDef.image, ...(slideDef.screenshots || [])]
    .map((item) => item?.path)
    .filter((item) => canRenderImageFile(item));
  const primaryPreviewPath = previewImages[0] || "";

  if (template.variant === "cards") {
    stages.slice(0, 4).forEach((stage, index) => {
      const row = Math.floor(index / 2);
      const col = index % 2;
      addInfoCard(slide, stage, 0.82 + col * 5.42, 1.74 + row * 1.42, 4.94, 1.08, style, {
        accent: style.palette[stage.accent] || style.palette.darkGreen,
        fill: row % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["process", `card-${index}`],
      });
    });
    notes.slice(0, 2).forEach((card, index) => {
      addInfoCard(slide, card, 0.82 + index * 5.42, 4.68, 4.94, 1.02, style, {
        accent: style.palette[card.accent] || style.palette.accentOrange,
        fill: style.palette.subtleGray,
        tags: ["process-note", `card-${index}`],
      });
    });
  } else if (template.variant === "bridge") {
    addRibbon(slide, "问题与方法桥接", 0.82, 1.66, 3.0, 0.42, style, "chevron", style.palette.deepGreen);
    addRibbon(slide, "关键路径", 4.0, 1.66, 4.0, 0.42, style, "parallelogram", style.palette.oliveGreen);
    addRibbon(slide, "反馈与闭环", 8.28, 1.66, 3.0, 0.42, style, "homePlate", style.palette.accentBlue);
    addPanel(slide, 0.82, 2.16, 3.0, 3.2, style, { fill: style.palette.white, line: style.palette.softLine });
    addBulletBlock(
      slide,
      "建模出发点",
      stepBullets.length ? stepBullets : [slideDef.summary || "围绕问题定义、路径设计与闭环推进展开。"],
      0.96,
      2.32,
      2.72,
      2.86,
      style,
      style.palette.darkGreen,
    );
    addPanel(slide, 4.04, 2.16, 4.0, 3.2, style, { fill: style.palette.subtleGray, line: style.palette.softLine });
    stages.forEach((stage, index) => {
      const y = 2.34 + index * 0.56;
      addRibbon(
        slide,
        stage.title,
        4.22,
        y,
        3.44,
        0.26,
        style,
        ["rightArrow", "chevron", "parallelogram", "homePlate", "rightArrow"][index % 5],
        [style.palette.oliveGreen, style.palette.darkGreen, style.palette.accentBlue, style.palette.accentOrange, style.palette.deepGreen][index % 5],
      );
      addText(slide, stage.detail, 4.28, y + 0.3, 3.3, 0.18, style, {
        fontSize: 7.9,
        color: style.palette.textMuted,
      });
    });
    addPanel(slide, 8.28, 2.16, 2.8, 3.2, style, { fill: style.palette.white, line: style.palette.softLine });
    if (primaryPreviewPath) {
      slide.addImage(fitImageByAspect(primaryPreviewPath, { x: 8.42, y: 2.3, w: 2.52, h: 1.36 }));
      notes.slice(0, 2).forEach((card, index) => {
        addInfoCard(slide, card, 8.42, 3.84 + index * 0.72, 2.52, 0.58, style, {
          accent: style.palette[card.accent] || style.palette.accentOrange,
          fill: index === 0 ? style.palette.white : style.palette.lightGreen,
          tags: ["process-note", "bridge", `note-${index}`],
        });
      });
    } else {
      addBulletBlock(
        slide,
        "反馈与闭环",
        notes.length ? notes.map((card) => card.body || card.detail || card.title || "").filter(Boolean) : ["补充关键反馈、闭环动作与风险提示。"],
        8.42,
        2.3,
        2.52,
        2.86,
        style,
        style.palette.accentOrange,
      );
    }
  } else if (template.variant === "ladder") {
    addRibbon(slide, "方法框架", 0.82, 1.66, 2.3, 0.4, style, "homePlate", style.palette.deepGreen);
    addPanel(slide, 0.82, 2.1, 11.0, 3.8, style, { fill: style.palette.subtleGray, line: style.palette.softLine });
    stages.forEach((stage, index) => {
      const x = 1.02 + index * 2.08;
      const y = 2.38 + (index % 2) * 0.42;
      addInfoCard(slide, stage, x, y, 1.84, 1.48, style, {
        accent: style.palette[stage.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["process", "ladder", `stage-${index}`],
      });
      if (index < Math.min((slideDef.stages || []).length, 5) - 1) {
        slide.addShape("rightArrow", {
          x: x + 1.9,
          y: y + 0.52,
          w: 0.16,
          h: 0.12,
          fill: { color: style.palette.borderGreen },
          line: { color: style.palette.borderGreen, transparency: 100 },
        });
      }
    });
    notes.slice(0, 2).forEach((card, index) => {
      addInfoCard(slide, card, 8.04, 4.32 + index * 0.82, 3.52, 0.68, style, {
        accent: style.palette[card.accent] || style.palette.accentOrange,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["process-note", "ladder", `note-${index}`],
      });
    });
  } else {
    addRibbon(slide, "建模出发点", 0.82, 1.66, 2.34, 0.4, style, "parallelogram", style.palette.oliveGreen);
    addRibbon(slide, "核心流程", 3.56, 1.66, 4.08, 0.4, style, "chevron", style.palette.darkGreen);
    addRibbon(slide, "反馈与关注点", 8.28, 1.66, 3.0, 0.4, style, "homePlate", style.palette.deepGreen);
    addBulletBlock(
      slide,
      "建模出发点",
      stepBullets.length ? stepBullets : [slideDef.summary || "围绕问题定义、建模路径与落地反馈展开。"],
      0.82,
      2.14,
      2.62,
      3.72,
      style,
      style.palette.darkGreen,
    );
    addPanel(slide, 3.62, 2.14, 4.1, 3.74, style, { fill: style.palette.subtleGray, line: style.palette.softLine });
    stages.forEach((stage, index) => {
      const y = 2.38 + index * 0.64;
      addIconBadge(slide, style, 3.84, y, 0.26, style.palette[stage.accent] || style.palette.darkGreen, ["process", `index-${index}`]);
      addRibbon(
        slide,
        stage.title,
        4.28,
        y - 0.02,
        2.74,
        0.34,
        style,
        index % 2 === 0 ? "chevron" : "parallelogram",
        index % 2 === 0 ? style.palette.oliveGreen : style.palette.accentBlue,
      );
      addText(slide, stage.detail, 4.28, y + 0.3, 2.88, 0.2, style, {
        fontSize: 7.6,
        color: style.palette.textMuted,
      });
    });
    if (primaryPreviewPath) {
      addPanel(slide, 8.2, 2.18, 3.62, 2.04, style, {
        fill: style.palette.subtleGray,
        line: style.palette.softLine,
      });
      slide.addImage(fitImageByAspect(primaryPreviewPath, { x: 8.28, y: 2.26, w: 3.46, h: 1.88 }));
      if (notes[0]) {
        addInfoCard(slide, notes[0], 8.2, 4.42, 3.62, 0.86, style, {
          accent: style.palette[notes[0].accent] || style.palette.accentOrange,
          fill: style.palette.white,
          tags: ["process-note", "preview-note"],
        });
      } else {
        addBulletBlock(slide, "反馈与闭环", [slideDef.summary || "请结合试点反馈与后续闭环动作继续推进。"], 8.2, 4.42, 3.62, 0.86, style, style.palette.accentOrange);
      }
    } else {
      addBulletBlock(
        slide,
        "反馈与关注点",
        notes.length ? notes.map((card) => card.body || card.detail || card.title || "").filter(Boolean) : ["补充关键反馈、落地抓手与后续闭环动作。"],
        8.2,
        2.18,
        3.62,
        3.1,
        style,
        style.palette.accentOrange,
      );
    }
    (slideDef.focusItems || slideDef.metrics || []).slice(0, 3).forEach((item, index) => {
      const value = item.value || item.label || "";
      addPanel(slide, 8.22 + index * 1.16, 5.34, 1.02, 0.42, style, {
        fill: style.palette.lightGreen,
        line: style.palette.lightGreen,
      });
      addText(slide, value, 8.28 + index * 1.16, 5.46, 0.9, 0.1, style, {
        fontSize: 7.1,
        bold: true,
        color: style.palette.textMuted,
        align: "center",
      });
    });
  }

  addPageNumber(slide, slideDef.page, style);
}

function renderBulletColumns(slide, slideDef, style, template) {
  addBrandHeader(slide, slideDef.title, style);
  addSummaryBand(slide, slideDef.headline, style);

  const columns = slideDef.columns || [];
  if (template.variant === "staggered") {
    columns.slice(0, 3).forEach((column, index) => {
      addBulletBlock(
        slide,
        column.title,
        column.bullets || [],
        0.82 + index * 3.84,
        1.72 + (index % 2) * 0.28,
        3.44,
        2.74,
        style,
        [style.palette.darkGreen, style.palette.oliveGreen, style.palette.accentOrange][index],
      );
    });
    (slideDef.cards || []).slice(0, 2).forEach((card, index) => {
      addInfoCard(slide, card, 1.04 + index * 5.62, 4.9, 5.08, 0.92, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["bullet", "staggered", `footer-${index}`],
      });
    });
  } else if (template.variant === "masonry") {
    const left = columns[0];
    const middle = columns[1];
    const right = columns[2];
    if (left) {
      addBulletBlock(slide, left.title, left.bullets || [], 0.82, 1.72, 4.0, 3.0, style, style.palette.darkGreen);
    }
    if (middle) {
      addBulletBlock(slide, middle.title, middle.bullets || [], 4.96, 1.94, 3.4, 2.66, style, style.palette.oliveGreen);
    }
    if (right) {
      addBulletBlock(slide, right.title, right.bullets || [], 8.52, 1.72, 3.0, 3.0, style, style.palette.accentOrange);
    }
    (slideDef.cards || []).slice(0, 2).forEach((card, index) => {
      addInfoCard(slide, card, 0.98 + index * 5.62, 4.98, 5.08, 0.84, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["bullet", "masonry", `footer-${index}`],
      });
    });
  } else {
    const triple = template.variant === "triple";
    const width = triple ? 3.72 : 5.42;
    const step = triple ? 4.04 : 5.66;
    columns.slice(0, triple ? 3 : 2).forEach((column, index) => {
      addBulletBlock(slide, column.title, column.bullets || [], 0.82 + index * step, 1.72, width, 3.2, style, [style.palette.darkGreen, style.palette.oliveGreen, style.palette.accentOrange][index]);
    });
    (slideDef.cards || []).slice(0, 3).forEach((card, index) => {
      addInfoCard(slide, card, 0.82 + index * 4.0, 5.18, 3.72, 0.82, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["bullet", `footer-${index}`],
      });
    });
  }

  addPageNumber(slide, slideDef.page, style);
}

function renderImageStory(slide, slideDef, style, template) {
  addBrandHeader(slide, slideDef.title, style);
  addSummaryBand(slide, slideDef.headline, style);

  const imagePath = slideDef.image?.path && fs.existsSync(slideDef.image.path) ? slideDef.image.path : "";
  const screenshots = (slideDef.screenshots || []).filter((item) => item?.path && fs.existsSync(item.path));
  const primaryImage = screenshots[0]?.path || imagePath;
  const primaryAspect = Number(screenshots[0]?.aspectRatio || slideDef.image?.aspectRatio || 0);
  const callouts =
    (slideDef.callouts || []).length
      ? (slideDef.callouts || []).slice(0, 3)
      : (slideDef.textBlocks || []).slice(0, 3).map((text, index) => ({
          title: ["重点内容", "补充说明", "落地动作"][index] || `要点${index + 1}`,
          body: text,
          accent: ["accentOrange", "darkGreen", "oliveGreen"][index % 3],
        }));
  const bullets = (slideDef.bullets || []).slice(0, 5);
  const narrative = (slideDef.textBlocks || []).slice(0, 4);

  if (template.variant === "storyboard") {
    addPanel(slide, 0.82, 1.72, 6.2, 3.94, style, { line: style.palette.softLine });
    if (primaryImage) {
      const imageRegion =
        primaryAspect >= 1.2
          ? { x: 0.98, y: 1.9, w: 5.86, h: 2.84 }
          : primaryAspect > 0 && primaryAspect <= 0.9
            ? { x: 1.7, y: 1.9, w: 4.2, h: 3.42 }
            : { x: 1.1, y: 1.9, w: 5.5, h: 3.18 };
      slide.addImage(fitImageByAspect(primaryImage, imageRegion));
    }
    const boxes = gridBoxes(Math.max(1, callouts.length), 7.28, 1.82, 4.54, 3.86, {
      cols: 1,
      gapY: 0.12,
    });
    callouts.forEach((card, index) => {
      addInfoCard(slide, card, boxes[index].x, boxes[index].y, boxes[index].w, boxes[index].h, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["image", "storyboard", `callout-${index}`],
      });
    });
    addPanel(slide, 0.82, 5.86, 11.0, 0.58, style, { fill: style.palette.lightGreen, line: style.palette.borderGreen });
    addText(slide, slideDef.footer || "", 1.0, 5.98, 10.64, 0.18, style, {
      fontSize: autoFontSize(slideDef.footer || "", 10.4, 9, 16),
      bold: true,
      color: style.palette.darkGreen,
      align: "center",
    });
  } else if (template.variant === "focus" && imagePath) {
    addPanel(slide, 0.82, 1.72, 11.0, 2.72, style, { line: style.palette.softLine });
    slide.addImage(
      fitImageByAspect(imagePath, {
        x: 0.96,
        y: 1.88,
        w: 10.72,
        h: 2.38,
      }),
    );
    callouts.slice(0, 3).forEach((card, index) => {
      addInfoCard(slide, card, 0.82 + index * 3.72, 4.72, 3.4, 1.06, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["image", "focus", `callout-${index}`],
      });
    });
  } else if (template.variant === "gallery") {
    const galleryShots = screenshots.length ? screenshots.slice(0, 3) : imagePath ? [{ path: imagePath, aspectRatio: primaryAspect }] : [];
    if (galleryShots.length >= 3) {
      galleryShots.forEach((shot, index) => {
        const boxX = 0.82 + index * 3.66;
        addPanel(slide, boxX, 1.88, 3.32, 2.1, style, { line: style.palette.softLine });
        slide.addImage(
          fitImageByAspect(shot.path, {
            x: boxX + 0.16,
            y: 2.0,
            w: 3.0,
            h: 1.52,
          }),
        );
        const card = (slideDef.callouts || [])[index] || {
          title: `场景${index + 1}`,
          body: (slideDef.bullets || [])[index] || "",
        };
        addInfoCard(slide, card, boxX, 4.22, 3.32, 1.18, style, {
          accent: style.palette[card.accent] || style.palette.darkGreen,
          fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
          tags: ["image", "gallery", `card-${index}`],
        });
      });
    } else {
      const heroShot = galleryShots[0]?.path || imagePath;
      const heroAspect = Number(galleryShots[0]?.aspectRatio || primaryAspect || slideDef.image?.aspectRatio || 0);
      const heroBox =
        heroAspect >= 1.45
          ? { x: 0.82, y: 1.88, w: 6.34, h: 3.0 }
          : heroAspect > 0 && heroAspect <= 0.9
            ? { x: 0.82, y: 1.88, w: 5.1, h: 3.38 }
            : { x: 0.82, y: 1.88, w: 6.0, h: 3.2 };
      addPanel(slide, heroBox.x, heroBox.y, heroBox.w, heroBox.h, style, { line: style.palette.softLine });
      if (heroShot) {
        slide.addImage(fitImageByAspect(heroShot, { x: heroBox.x + 0.14, y: heroBox.y + 0.12, w: heroBox.w - 0.28, h: heroBox.h - 0.24 }));
      }

      const rightX = 7.42;
      const rightW = 4.0;
      if (galleryShots[1]?.path) {
        addPanel(slide, rightX, 1.88, rightW, 1.16, style, { fill: style.palette.subtleGray, line: style.palette.softLine });
        slide.addImage(fitImageByAspect(galleryShots[1].path, { x: rightX + 0.08, y: 2.0, w: rightW - 0.16, h: 0.92 }));
      }
      const cards = callouts.slice(0, 3);
      const cardStartY = galleryShots[1]?.path ? 3.18 : 2.06;
      const cardBoxes = gridBoxes(Math.max(1, cards.length || (galleryShots[1]?.path ? 2 : 3)), rightX, cardStartY, rightW, 2.72, {
        cols: 1,
        gapY: 0.12,
      });
      const fallbackCards = cards.length
        ? cards
        : [
            { title: "关键看点", body: bullets[0] || narrative[0] || "突出关键要点。", accent: "accentOrange" },
            { title: "辅助说明", body: bullets[1] || narrative[1] || "补充背景和逻辑。", accent: "darkGreen" },
            { title: "下一步动作", body: bullets[2] || narrative[2] || "落到后续推进。", accent: "oliveGreen" },
          ];
      fallbackCards.slice(0, 3).forEach((card, index) => {
        addInfoCard(slide, card, cardBoxes[index].x, cardBoxes[index].y, cardBoxes[index].w, 0.82, style, {
          accent: style.palette[card.accent] || style.palette.darkGreen,
          fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
          tags: ["image", "gallery", `card-${index}`],
        });
      });
    }
  } else {
    const leftW = primaryImage ? 5.36 : 4.82;
    const rightX = primaryImage ? 6.42 : 5.88;
    const rightW = primaryImage ? 5.4 : 5.94;
    addPanel(slide, 0.82, 1.72, leftW, 4.18, style, { line: style.palette.softLine });
    if (primaryImage) {
      slide.addImage(fitImageByAspect(primaryImage, { x: 0.96, y: 1.9, w: leftW - 0.28, h: 2.8 }));
      addPanel(slide, 0.96, 4.86, leftW - 0.28, 0.86, style, { fill: style.palette.subtleGray, line: style.palette.softLine });
      addText(slide, narrative[0] || slideDef.footer || "请结合系统成果说明关键变化。", 1.12, 5.0, leftW - 0.6, 0.5, style, {
        fontSize: autoFontSize(narrative[0] || "", 8.8, 7.8, 20),
        bold: true,
        color: style.palette.deepGreen,
      });
    } else {
      addBulletBlock(slide, "重点内容", bullets.length ? bullets : narrative, 0.98, 1.9, leftW - 0.32, 3.82, style, style.palette.darkGreen);
    }
    addBulletBlock(slide, "重点内容", bullets.length ? bullets : narrative, rightX, 1.72, rightW, 2.26, style, style.palette.darkGreen);
    callouts.slice(0, 2).forEach((card, index) => {
      addInfoCard(slide, card, rightX, 4.16 + index * 0.9, rightW, 0.74, style, {
        accent: style.palette[card.accent] || style.palette.accentOrange,
        fill: index === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["image", `callout-${index}`],
      });
    });
  }

  addPageNumber(slide, slideDef.page, style);
}

function renderActionPlan(slide, slideDef, style, template) {
  addBrandHeader(slide, slideDef.title, style);
  addSummaryBand(slide, slideDef.headline, style);

  const steps = (slideDef.steps || slideDef.actions || []).slice(0, 4);
  const screenshots = (slideDef.screenshots || []).filter((item) => item?.path && fs.existsSync(item.path));
  const primaryPreviewPath = screenshots[0]?.path || slideDef.image?.path || "";
  const primaryPreviewAspect = Number(screenshots[0]?.aspectRatio || slideDef.image?.aspectRatio || 0);
  const timelineItems = (slideDef.timeline || steps || []).slice(0, 4);

  if (template.variant === "dashboard") {
    const stepBoxes = gridBoxes(Math.min(3, steps.length || 3), 0.82, 1.78, 11.0, 1.02, {
      cols: Math.min(3, Math.max(1, steps.length || 3)),
      gapX: 0.16,
    });
    steps.slice(0, 3).forEach((step, index) => {
      addInfoCard(slide, step, stepBoxes[index].x, stepBoxes[index].y, stepBoxes[index].w, stepBoxes[index].h, style, {
        accent: style.palette[step.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["action", "dashboard", `step-${index}`],
      });
    });
    const previewBox =
      primaryPreviewAspect >= 1.45
        ? { x: 0.98, y: 3.1, w: 6.1, h: 2.0 }
        : primaryPreviewAspect > 0 && primaryPreviewAspect <= 0.85
          ? { x: 1.18, y: 3.02, w: 3.92, h: 2.5 }
          : { x: 1.04, y: 3.1, w: 5.88, h: 2.02 };
    addPanel(slide, 0.82, 3.0, 6.4, 2.24, style, { fill: hasPreview ? style.palette.subtleGray : style.palette.white, line: style.palette.softLine });
    if (primaryPreviewPath) {
      slide.addImage(fitImageByAspect(primaryPreviewPath, previewBox));
    } else {
      addText(slide, "推进重点", 1.06, 3.18, 1.2, 0.18, style, {
        fontFace: style.fonts.heading,
        fontSize: 10.2,
        bold: true,
        color: style.palette.darkGreen,
      });
      timelineItems.slice(0, 4).forEach((item, index) => {
        addPanel(slide, 1.04, 3.5 + index * 0.38, 5.76, 0.28, style, {
          fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
          line: style.palette.softLine,
        });
        addText(slide, `${index + 1}. ${clipText(item.title || item.label || item.detail || "", 44)}`, 1.22, 3.59 + index * 0.38, 5.36, 0.08, style, {
          fontSize: 8.0,
          bold: index === 0,
          color: index === 0 ? style.palette.darkGreen : style.palette.textPrimary,
        });
      });
    }
    addPanel(slide, 7.48, 3.0, 4.34, 2.24, style, { fill: style.palette.white, line: style.palette.softLine });
    addText(slide, "实施节奏", 7.72, 3.16, 1.2, 0.18, style, {
      fontFace: style.fonts.heading,
      fontSize: 10.4,
      bold: true,
      color: style.palette.accentOrange,
    });
    timelineItems.forEach((item, index) => {
      addText(slide, `${index + 1}. ${item.title || item.label || ""}`, 7.76, 3.44 + index * 0.36, 3.8, 0.16, style, {
        fontSize: 8.1,
        bold: index === 0,
        color: index % 2 === 0 ? style.palette.deepGreen : style.palette.textPrimary,
      });
    });
    addPanel(slide, 0.82, 5.42, 11.0, 0.62, style, { fill: style.palette.lightGreen, line: style.palette.borderGreen });
    addText(slide, slideDef.systemNote || "围绕系统、流程和组织三条线持续推进闭环管理，形成更稳定的落地节奏。", 1.04, 5.56, 10.56, 0.16, style, {
      fontSize: autoFontSize(slideDef.systemNote || "", 10.4, 9, 18),
      bold: true,
      color: style.palette.darkGreen,
      align: "center",
    });
  } else if (template.variant === "matrix") {
    const topBoxes = gridBoxes(Math.min(2, steps.length || 1), 0.82, 1.76, 7.06, 1.76, {
      cols: Math.min(2, steps.length || 1),
      gapX: 0.22,
    });
    steps.slice(0, 2).forEach((step, index) => {
      addInfoCard(slide, step, topBoxes[index].x, topBoxes[index].y, topBoxes[index].w, topBoxes[index].h, style, {
        accent: style.palette[step.accent] || style.palette.darkGreen,
        tags: ["action", "matrix", `top-${index}`],
      });
    });
    if (steps[2]) {
      addInfoCard(slide, steps[2], 0.82, 3.74, 7.06, 1.54, style, {
        accent: style.palette[steps[2].accent] || style.palette.accentOrange,
        fill: style.palette.lightGreen,
        tags: ["action", "matrix", "bottom"],
      });
    }
    const previewPanelH = screenshots.length >= 2 ? 3.84 : 3.68;
    addPanel(slide, 8.0, 1.76, 3.82, previewPanelH, style, { line: style.palette.softLine, fill: screenshots.length || primaryPreviewPath ? style.palette.subtleGray : style.palette.white });
    addText(slide, screenshots.length || primaryPreviewPath ? "系统预览" : "实施要点", 8.24, 1.92, 1.4, 0.18, style, {
      fontFace: style.fonts.heading,
      fontSize: 10.2,
      bold: true,
      color: style.palette.accentOrange,
    });
    if (screenshots.length) {
      const showShots = screenshots.slice(0, 2);
      showShots.forEach((shot, index) => {
        const slot = primaryPreviewAspect >= 1.45
          ? { x: 8.18, y: 2.32 + index * 1.38, w: 3.46, h: 1.24 }
          : { x: 8.18, y: 2.26 + index * 1.46, w: 3.46, h: 1.3 };
        slide.addImage(fitImageByAspect(shot.path, slot));
      });
    } else if (primaryPreviewPath) {
      slide.addImage(fitImageByAspect(primaryPreviewPath, { x: 8.22, y: 2.28, w: 3.34, h: 1.82 }));
      addText(slide, slideDef.systemNote || "围绕组织协同、系统落地和应用闭环持续推进。", 8.26, 4.18, 3.28, 0.72, style, {
        fontSize: autoFontSize(slideDef.systemNote || "", 9.0, 8.0, 22),
        bold: true,
      });
    } else {
      const matrixNotes = [slideDef.systemNote, ...(timelineItems || []).map((item) => item.title || item.label || item.detail || "")]
        .filter(Boolean)
        .slice(0, 4);
      addBulletBlock(slide, "实施要点", matrixNotes, 8.14, 2.24, 3.54, 2.94, style, style.palette.accentOrange);
    }
  } else if (template.variant === "stacked") {
    steps.slice(0, 3).forEach((step, index) => {
      addInfoCard(slide, step, 0.82, 1.82 + index * 1.2, 7.08, 0.92, style, {
        accent: style.palette[step.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["action", "stacked", `step-${index}`],
      });
    });
    const stackedPreviewH = screenshots.length || primaryPreviewPath ? 4.28 : 4.18;
    addPanel(slide, 8.08, 1.82, 3.74, stackedPreviewH, style, { line: style.palette.softLine, fill: screenshots.length || primaryPreviewPath ? style.palette.subtleGray : style.palette.white });
    addText(slide, screenshots.length || primaryPreviewPath ? "系统预览 / 落地抓手" : "落地抓手", 8.32, 2.0, 2.4, 0.2, style, {
      fontFace: style.fonts.heading,
      fontSize: 10.2,
      bold: true,
      color: style.palette.accentOrange,
    });
    if (screenshots.length) {
      screenshots.slice(0, 2).forEach((shot, index) => {
        const slot = primaryPreviewAspect >= 1.45
          ? { x: 8.24, y: 2.34 + index * 1.42, w: 3.38, h: 1.18 }
          : { x: 8.24, y: 2.36 + index * 1.46, w: 3.38, h: 1.28 };
        slide.addImage(fitImageByAspect(shot.path, slot));
      });
    } else if (primaryPreviewPath) {
      slide.addImage(fitImageByAspect(primaryPreviewPath, { x: 8.28, y: 2.34, w: 3.3, h: 1.86 }));
    }
    if (screenshots.length || primaryPreviewPath) {
      addText(slide, slideDef.systemNote || "围绕组织协同、系统落地和应用闭环持续推进。", 8.28, 4.44, 3.32, 1.0, style, {
        fontSize: autoFontSize(slideDef.systemNote || "", 8.8, 7.8, 24),
        bold: true,
      });
    } else {
      addBulletBlock(
        slide,
        "落地抓手",
        [slideDef.systemNote || "围绕组织协同、系统落地和应用闭环持续推进。", ...timelineItems.map((item) => item.title || item.label || item.detail || "")].filter(Boolean).slice(0, 4),
        8.2,
        2.28,
        3.5,
        3.34,
        style,
        style.palette.accentOrange,
      );
    }
  } else {
    const boxes = gridBoxes(Math.min(3, steps.length || 1), 0.82, 1.82, 11.0, 2.08, {
      cols: Math.min(3, steps.length || 1),
      gapX: 0.18,
    });
    steps.slice(0, 3).forEach((step, index) => {
      addInfoCard(slide, step, boxes[index].x, boxes[index].y, boxes[index].w, boxes[index].h, style, {
        accent: style.palette[step.accent] || style.palette.darkGreen,
        tags: ["action", "timeline", `step-${index}`],
      });
    });
    addPanel(slide, 0.82, 4.28, 6.86, 1.2, style, { fill: style.palette.subtleGray, line: style.palette.softLine });
    addText(slide, "推进节奏", 1.06, 4.44, 1.0, 0.18, style, {
      fontFace: style.fonts.heading,
      fontSize: 10.2,
      bold: true,
      color: style.palette.darkGreen,
    });
    (slideDef.timeline || []).slice(0, 3).forEach((item, index) => {
      addPanel(slide, 1.06 + index * 1.84, 4.78, 1.52, 0.32, style, {
        fill: style.palette.lightGreen,
        line: style.palette.lightGreen,
      });
      addText(slide, item.label, 1.14 + index * 1.84, 4.88, 1.36, 0.1, style, {
        fontSize: 7.6,
        bold: true,
        color: style.palette.textMuted,
        align: "center",
      });
      if (index < Math.min((slideDef.timeline || []).length, 3) - 1) {
        slide.addShape("rightArrow", {
          x: 2.64 + index * 1.84,
          y: 4.88,
          w: 0.22,
          h: 0.1,
          fill: { color: style.palette.borderGreen },
          line: { color: style.palette.borderGreen, transparency: 100 },
        });
      }
    });
    addPanel(slide, 7.9, 4.28, 3.92, 1.2, style, { line: style.palette.softLine, fill: screenshots[0] ? style.palette.subtleGray : style.palette.white });
    if (screenshots[0]) {
      slide.addImage(fitImageByAspect(screenshots[0].path, { x: 8.08, y: 4.42, w: 1.26, h: 0.9 }));
      addText(slide, slideDef.systemNote || "围绕系统预览与组织协同持续推进。", 9.42, 4.48, 2.12, 0.64, style, {
        fontSize: autoFontSize(slideDef.systemNote || "", 8.6, 7.8, 22),
        bold: true,
      });
    } else {
      addBulletBlock(
        slide,
        "配套动作",
        [slideDef.systemNote || "后续围绕落地节奏、系统闭环与组织协同持续推进。", ...timelineItems.map((item) => item.title || item.label || item.detail || "")].filter(Boolean).slice(0, 3),
        8.02,
        4.36,
        3.68,
        1.04,
        style,
        style.palette.accentOrange,
      );
    }
  }

  addPageNumber(slide, slideDef.page, style);
}

function renderKeyTakeaways(slide, slideDef, style, template) {
  addBrandHeader(slide, slideDef.title, style);
  addSummaryBand(slide, slideDef.headline, style);

  const takeaways = (slideDef.takeaways || []).slice(0, 4);
  if (template.variant === "wall") {
    const topBoxes = gridBoxes(Math.min(4, takeaways.length || 1), 0.82, 1.82, 11.0, 1.04, {
      cols: 2,
      gapX: 0.18,
      gapY: 0.16,
    });
    takeaways.slice(0, 4).forEach((card, index) => {
      addInfoCard(slide, card, topBoxes[index].x, topBoxes[index].y, topBoxes[index].w, topBoxes[index].h, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["takeaway", "wall", `card-${index}`],
      });
    });
    addPanel(slide, 0.82, 4.42, 11.0, 1.02, style, {
      fill: style.palette.lightGreen,
      line: style.palette.borderGreen,
    });
    addText(slide, slideDef.footer || "建议将关键信息继续沉淀为可复用模板与管理动作。", 1.04, 4.68, 10.56, 0.34, style, {
      fontSize: autoFontSize(slideDef.footer || "", 12, 10, 20),
      bold: true,
      color: style.palette.darkGreen,
      align: "center",
      valign: "mid",
    });
  } else if (template.variant === "closing") {
    const topBoxes = gridBoxes(Math.min(2, takeaways.length || 1), 0.82, 1.82, 11.0, 1.58, {
      cols: Math.min(2, takeaways.length || 1),
      gapX: 0.22,
    });
    takeaways.slice(0, 2).forEach((card, index) => {
      addInfoCard(slide, card, topBoxes[index].x, topBoxes[index].y, topBoxes[index].w, topBoxes[index].h, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["takeaway", "closing", `card-${index}`],
      });
    });
    addPanel(slide, 0.82, 3.68, 11.0, 1.86, style, {
      fill: style.palette.lightGreen,
      line: style.palette.borderGreen,
    });
    addText(slide, slideDef.footer || "通过本次汇报形成清晰的结论、抓手与后续安排。", 1.02, 4.1, 10.6, 1.02, style, {
      fontSize: autoFontSize(slideDef.footer || "", 12, 10, 20),
      bold: true,
      align: "center",
      valign: "mid",
    });
  } else {
    const topBoxes = gridBoxes(Math.min(3, takeaways.length || 1), 0.82, 1.92, 11.0, 1.7, {
      cols: Math.min(3, takeaways.length || 1),
      gapX: 0.18,
    });
    takeaways.slice(0, 3).forEach((card, index) => {
      addInfoCard(slide, card, topBoxes[index].x, topBoxes[index].y, topBoxes[index].w, topBoxes[index].h, style, {
        accent: style.palette[card.accent] || style.palette.darkGreen,
        fill: index % 2 === 0 ? style.palette.white : style.palette.lightGreen,
        tags: ["takeaway", `card-${index}`],
      });
    });
    addPanel(slide, 0.82, 4.1, 11.0, 0.72, style, {
      fill: style.palette.subtleGray,
      line: style.palette.softLine,
    });
    addText(slide, slideDef.footer || "建议围绕关键结论继续细化实施路径、复核数据并沉淀为长期机制。", 1.0, 4.46, 10.64, 0.48, style, {
      fontSize: autoFontSize(slideDef.footer || "", 10.4, 9, 24),
      bold: true,
      align: "center",
    });
  }

  addPageNumber(slide, slideDef.page, style);
}

function normalizeTableForRender(table) {
  const header = Array.isArray(table?.header) ? [...table.header] : [];
  const rows = Array.isArray(table?.rows) ? table.rows.map((row) => (Array.isArray(row) ? [...row] : [])) : [];
  const caption = String(table?.caption || "");
  const colCount = Math.max(header.length, ...rows.map((row) => row.length), table?.colCount || 0, 1);
  let normalizedHeader = header;
  let normalizedRows = rows;

  if (header.length === 1 && rows.length && rows[0].length > 1) {
    normalizedHeader = rows[0];
    normalizedRows = rows.slice(1);
  }

  if (!normalizedHeader.length) {
    normalizedHeader = Array.from({ length: colCount }, (_, index) => `列${index + 1}`);
  }

  if (normalizedHeader.length < colCount) {
    normalizedHeader = [...normalizedHeader, ...Array(colCount - normalizedHeader.length).fill("")];
  }

  normalizedRows = normalizedRows
    .filter((row) => row.some((cell) => String(cell || "").trim()))
    .map((row) => [...row, ...Array(Math.max(0, colCount - row.length)).fill("")]);

  return {
    caption,
    highlight: table?.highlight || "",
    header: normalizedHeader,
    rows: normalizedRows,
    colCount,
    rowCount: normalizedRows.length,
  };
}

function renderSlide(slide, slideDef, style) {
  const template = resolveTemplate(style, slideDef);
  style.currentTemplateVariant = template.variant || "";
  style.currentSlideType = slideDef.type || "";

  switch (slideDef.type) {
    case "cover":
      renderCover(slide, slideDef, style, template);
      break;
    case "summary_cards":
      renderSummaryCards(slide, slideDef, style, template);
      break;
    case "table_analysis":
      renderTableAnalysis(slide, slideDef, style, template);
      break;
    case "process_flow":
      renderProcessFlow(slide, slideDef, style, template);
      break;
    case "bullet_columns":
      renderBulletColumns(slide, slideDef, style, template);
      break;
    case "image_story":
      renderImageStory(slide, slideDef, style, template);
      break;
    case "action_plan":
      renderActionPlan(slide, slideDef, style, template);
      break;
    case "key_takeaways":
    default:
      renderKeyTakeaways(slide, slideDef, style, template);
      break;
  }
}

async function renderDeck(outline, style, outputPath) {
  const pptx = new PptxGenJS();
  pptx.defineLayout({
    name: "PSBC_CUSTOM",
    width: style.page.widthInches,
    height: style.page.heightInches,
  });
  pptx.layout = "PSBC_CUSTOM";
  pptx.author = "OpenAI Codex";
  pptx.company = "PSBC";
  pptx.subject = outline.meta.title;
  pptx.title = outline.meta.title;
  pptx.lang = "zh-CN";
  pptx.theme = {
    headFontFace: style.fonts.heading,
    bodyFontFace: style.fonts.body,
  };
  pptx.background = { color: style.palette.white };

  (outline.slides || []).forEach((slideDef) => {
    const slide = pptx.addSlide();
    renderSlide(slide, slideDef, style);
  });

  await pptx.writeFile({ fileName: outputPath, compression: true });
}

module.exports = {
  renderDeck,
};
