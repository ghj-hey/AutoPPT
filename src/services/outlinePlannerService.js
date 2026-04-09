const fs = require("node:fs");
const { imageSize } = require("image-size");
const {
  buildDocumentStructure,
  collectPlainText,
  normalizeText,
  stripHeadingPrefix,
} = require("./documentStructureService");

const CN_TOP_LEVEL_RE = /^[一二三四五六七八九十百千]+[、.]\s*/;
const TABLE_CAPTION_RE = /^表\s*\d+\s*[:：]/;
const IMAGE_CAPTION_RE = /^图\s*\d+\s*[:：]/;
const ACTION_KEYWORDS = ["下一步", "计划", "建议", "安排", "优化", "推进", "推广", "落地", "实施", "协同"];
const PROCESS_KEYWORDS = ["模型", "流程", "机制", "路径", "方法", "方案", "架构", "迭代", "建模", "思路"];
const RESULT_KEYWORDS = ["效果", "成效", "结果", "验证", "试点", "测试", "实验", "A/B", "成本", "压降", "收益"];

function clampPageCount(value) {
  if (value == null || value === "" || Number.isNaN(Number(value))) return 0;
  const numeric = Math.round(Number(value));
  if (numeric <= 0) return 0;
  return Math.max(2, Math.min(10, numeric));
}

function splitSentences(text) {
  return String(text || "")
    .split(/[。！？；\n]/)
    .map((item) => normalizeText(item))
    .filter(Boolean);
}

function compactTitle(text, maxLength = 30) {
  const value = stripHeadingPrefix(normalizeText(text || ""));
  if (!value) return "";
  return value.length <= maxLength ? value : `${value.slice(0, Math.max(10, maxLength - 1)).trim()}…`;
}

function compactBody(text, maxLength = 96) {
  const value = normalizeText(text || "");
  if (!value) return "";
  return value.length <= maxLength ? value : `${value.slice(0, Math.max(24, maxLength - 1)).trim()}…`;
}

function uniqueByKey(items = [], keyFn = (item) => item) {
  const seen = new Set();
  return (items || []).filter((item) => {
    const key = keyFn(item);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function chunkArray(items = [], parts = 2) {
  const source = Array.isArray(items) ? items.filter(Boolean) : [];
  if (!source.length) return [];
  const bucketCount = Math.max(1, Math.min(parts, source.length));
  const size = Math.ceil(source.length / bucketCount);
  const output = [];
  for (let index = 0; index < source.length; index += size) {
    output.push(source.slice(index, index + size));
  }
  return output;
}

function flattenHeadings(section = {}) {
  const headings = [];
  const walk = (node) => {
    (node.children || []).forEach((child) => {
      if (child.heading) headings.push(child.heading);
      walk(child);
    });
  };
  walk(section);
  return headings;
}

function flattenParagraphs(section = {}) {
  const paragraphs = [];
  const walk = (node) => {
    (node.paragraphs || []).forEach((paragraph) => paragraphs.push(paragraph));
    (node.children || []).forEach((child) => walk(child));
  };
  walk(section);
  return paragraphs.filter(Boolean);
}

function collectSectionText(section = {}) {
  return [section.heading, section.title, ...flattenHeadings(section), ...flattenParagraphs(section)]
    .filter(Boolean)
    .join("\n");
}

function imageKey(image = {}) {
  return normalizeText(image.path || image.source || image.caption || image.title || image.name || "");
}

function uniqueImages(images = []) {
  return uniqueByKey(images, (image) => imageKey(image));
}

function normalizeTableStructure(table) {
  const rawRows = (table?.rows || [])
    .map((row) => (Array.isArray(row) ? row : []).map((cell) => normalizeText(cell)))
    .filter((row) => row.some(Boolean));
  if (!rawRows.length) return null;

  const colCount = Math.max(...rawRows.map((row) => row.length), 1);
  let caption = normalizeText(table?.caption || "");
  let header = rawRows[0];
  let rows = rawRows.slice(1);

  if (rawRows.length >= 2) {
    const firstFilled = rawRows[0].filter(Boolean).length;
    const secondFilled = rawRows[1].filter(Boolean).length;
    if (firstFilled <= 1 || firstFilled < secondFilled) {
      caption = caption || rawRows[0].filter(Boolean).join(" ");
      header = rawRows[1];
      rows = rawRows.slice(2);
    }
  }

  header = [...header, ...Array(Math.max(0, colCount - header.length)).fill("")];
  rows = rows
    .map((row) => [...row, ...Array(Math.max(0, colCount - row.length)).fill("")])
    .filter((row) => row.some(Boolean));

  if (!caption) {
    caption = header.filter(Boolean).join(" / ").slice(0, 48) || "核心数据表";
  }

  return {
    caption,
    title: caption,
    header,
    rows,
    colCount,
    rowCount: rows.length,
    highlight: rows[0]?.filter(Boolean).join(" / ") || "",
  };
}

function allTables(section = {}) {
  return (section.tables || []).map((table) => normalizeTableStructure(table)).filter(Boolean);
}

function hasRealTable(section = {}) {
  return allTables(section).some((table) => table.colCount >= 2 && table.rowCount >= 1);
}

function primaryTable(section = {}) {
  return (
    allTables(section).sort((left, right) => {
      const leftScore = (left.rowCount || 0) * 10 + (left.colCount || 0);
      const rightScore = (right.rowCount || 0) * 10 + (right.colCount || 0);
      return rightScore - leftScore;
    })[0] || null
  );
}

function sliceTable(table, partIndex = 0, totalParts = 1) {
  if (!table) return null;
  if (totalParts <= 1 || !Array.isArray(table.rows) || table.rows.length <= 6) return table;
  const size = Math.ceil(table.rows.length / totalParts);
  const start = partIndex * size;
  const rows = table.rows.slice(start, start + size);
  return {
    ...table,
    rows,
    rowCount: rows.length,
    highlight: rows[0]?.filter(Boolean).join(" / ") || table.highlight || "",
  };
}

function splitTablesForPart(section, partIndex = 0, totalParts = 1) {
  const tables = allTables(section);
  if (!tables.length) return [];
  if (totalParts <= 1) return tables;
  if (tables.length > 1) {
    const buckets = chunkArray(tables, totalParts);
    return buckets[partIndex] || [];
  }
  const table = tables[0];
  if (table.rowCount <= 6) {
    return partIndex === 0 ? [table] : [];
  }
  const sliced = sliceTable(table, partIndex, totalParts);
  return sliced ? [sliced] : [];
}

function splitImagesForPart(section, partIndex = 0, totalParts = 1) {
  const images = uniqueImages(section.images || []);
  if (!images.length) return { mainImage: null, screenshots: [] };
  if (totalParts <= 1) return { mainImage: images[0], screenshots: images.slice(1, 3) };
  const buckets = chunkArray(images, totalParts);
  const bucket = buckets[partIndex] || [];
  return {
    mainImage: bucket[0] || null,
    screenshots: bucket.slice(1, 3),
  };
}

function estimateImageAspect(image) {
  if (!image?.path || !fs.existsSync(image.path)) return Number(image?.aspectRatio || 1.6);
  try {
    const size = imageSize(image.path);
    const width = Number(size?.width || image.width || 0);
    const height = Number(size?.height || image.height || 0);
    if (width > 0 && height > 0) return width / height;
  } catch {
    // ignore
  }
  return Number(image?.aspectRatio || 1.6);
}

function sectionLinesForPart(section = {}, partIndex = 0, totalParts = 1) {
  const lines = [];
  const push = (value) => {
    const text = normalizeText(value);
    if (text) lines.push(text);
  };

  push(section.heading);
  push(section.title);
  flattenHeadings(section).forEach(push);
  flattenParagraphs(section).forEach(push);
  (section.tables || []).forEach((table) => {
    const normalized = normalizeTableStructure(table);
    if (normalized?.caption) push(normalized.caption);
    if (normalized?.highlight) push(normalized.highlight);
    normalized?.rows?.slice(0, 2).forEach((row) => push(row.filter(Boolean).join(" ")));
  });
  uniqueImages(section.images || []).forEach((image) => {
    push(image?.caption);
    push(image?.title);
    push(image?.name);
  });

  const uniqueLines = uniqueByKey(
    lines.flatMap((line) => splitSentences(line)).filter((line) => line.length >= 4),
    (line) => line,
  );
  if (totalParts <= 1) return uniqueLines.slice(0, 14);
  const size = Math.max(4, Math.ceil(uniqueLines.length / totalParts));
  const start = partIndex * size;
  const sliced = uniqueLines.slice(start, start + size);
  return sliced.length ? sliced : uniqueLines.slice(Math.max(0, uniqueLines.length - size));
}

function makeBullets(lines = [], limit = 6) {
  return uniqueByKey(lines, (line) => line)
    .map((line) => compactBody(line, 86))
    .filter(Boolean)
    .slice(0, limit);
}

function makeCards(lines = [], limit = 4, section = {}) {
  const bullets = makeBullets(lines, limit * 2);
  const palette = ["darkGreen", "oliveGreen", "accentOrange", "accentBlue"];
  return bullets.slice(0, limit).map((line, index) => {
    const parts = splitSentences(line);
    const title = compactTitle(parts[0] || line, 22);
    const body = compactBody(parts.slice(1).join(" ") || line, 92);
    return {
      title,
      body,
      detail: body,
      accent: palette[index % palette.length],
      source: section.heading || section.title || "",
    };
  });
}

function makeColumns(lines = [], limit = 2) {
  const buckets = chunkArray(makeBullets(lines, 12), limit);
  return buckets.map((chunk, index) => ({
    title: index === 0 ? "重点内容" : "补充说明",
    bullets: chunk,
    items: chunk,
  }));
}

function makeStages(lines = [], limit = 4) {
  const buckets = chunkArray(makeBullets(lines, 12), limit);
  const titles = ["问题识别", "路径设计", "落地推进", "闭环优化"];
  return buckets.map((chunk, index) => ({
    title: titles[index] || `步骤${index + 1}`,
    detail: compactBody(chunk.join("，"), 86),
    body: compactBody(chunk.join("，"), 86),
    accent: ["darkGreen", "oliveGreen", "accentBlue", "accentOrange"][index % 4],
  }));
}

function extractHighlights(lines = [], limit = 4) {
  const results = [];
  const seen = new Set();
  (lines || []).forEach((line) => {
    const text = normalizeText(line);
    const matches = text.match(/[0-9]+(?:\.[0-9]+)?%|[0-9]+(?:\.[0-9]+)?亿元|[0-9]+(?:\.[0-9]+)?万元|[0-9]+(?:\.[0-9]+)?bp|[0-9]+(?:\.[0-9]+)?ms|[0-9]+xx/g) || [];
    matches.forEach((value) => {
      if (seen.has(value) || results.length >= limit) return;
      const idx = text.indexOf(value);
      const rawLabel = text.slice(Math.max(0, idx - 14), idx).replace(/[，,。；：: ]+$/, "").trim();
      results.push({
        label: compactTitle(rawLabel || "关键指标", 16),
        value,
      });
      seen.add(value);
    });
  });
  return results;
}

function sectionDensityScore(section = {}) {
  const charCount = Number(section.meta?.charCount || 0);
  const paragraphCount = Number(section.meta?.paragraphCount || 0);
  const tableCount = allTables(section).length;
  const imageCount = uniqueImages(section.images || []).length;
  const childCount = Number(section.meta?.subsectionCount || 0);
  return charCount / 220 + paragraphCount * 0.85 + tableCount * 1.5 + imageCount * 1.2 + childCount * 0.35;
}

function shouldSplitSection(section = {}) {
  return sectionDensityScore(section) >= 8.5;
}

function allocateContentPages(sections = [], contentBudget = 0) {
  const source = (sections || []).filter(Boolean);
  if (!source.length) return [];
  const target = Math.max(source.length, Math.round(Number(contentBudget) || source.length));
  const densities = source.map((section) => ({
    section,
    density: Math.max(0.1, sectionDensityScore(section)),
    pages: 1,
  }));

  const totalDensity = densities.reduce((sum, item) => sum + item.density, 0) || densities.length;
  densities.forEach((item) => {
    const ratio = item.density / totalDensity;
    item.pages = Math.max(1, Math.round(target * ratio));
  });

  const pageSum = () => densities.reduce((sum, item) => sum + item.pages, 0);
  const rankedDesc = () => [...densities].sort((left, right) => right.density - left.density);
  const rankedAsc = () => [...densities].sort((left, right) => left.density - right.density);

  let current = pageSum();
  while (current < target) {
    const candidate = rankedDesc().find((item) => item.pages < 10) || densities[0];
    if (!candidate) break;
    candidate.pages += 1;
    current += 1;
  }
  while (current > target) {
    const candidate = rankedAsc().find((item) => item.pages > 1);
    if (!candidate) break;
    candidate.pages -= 1;
    current -= 1;
  }

  const pagesBySectionId = new Map(densities.map((item) => [item.section.id, item.pages]));
  return source.map((section) => ({ section, pages: pagesBySectionId.get(section.id) || 1 }));
}

function hasStructuredProcessSignals(section = {}, text = "") {
  const paragraphCount = Number(section.meta?.paragraphCount || section.allParagraphs?.length || 0);
  const subsectionCount = Number(section.meta?.subsectionCount || section.allHeadings?.length || 0);
  const stepMarkers = [
    "第一",
    "第二",
    "第三",
    "第四",
    "第五",
    "第六",
    "第七",
    "首先",
    "其次",
    "再次",
    "最后",
    "步骤",
    "流程",
    "路径",
    "机制",
    "闭环",
    "推进",
    "实施",
    "落地",
  ];
  const hasExplicitSteps = stepMarkers.some((item) => text.includes(item));
  return paragraphCount >= 2 || subsectionCount >= 1 || hasExplicitSteps;
}

function classifySection(section = {}) {
  const heading = normalizeText(section.heading || "");
  const text = [heading, collectPlainText(section)].join("\n");
  const tableCount = allTables(section).length;
  const imageCount = uniqueImages(section.images || []).length;
  const hasAction = ACTION_KEYWORDS.some((item) => heading.includes(item) || text.includes(item));
  const hasProcess = PROCESS_KEYWORDS.some((item) => heading.includes(item) || text.includes(item));
  const hasResult = RESULT_KEYWORDS.some((item) => heading.includes(item) || text.includes(item));

  if (tableCount > 0) return "table_analysis";
  if (imageCount >= 2 && hasResult) return "image_story";
  if (hasProcess && hasStructuredProcessSignals(section, text)) return "process_flow";
  if (ACTION_KEYWORDS.some((item) => heading.includes(item))) return "action_plan";
  if (imageCount > 0) return "image_story";
  if (hasResult) return "key_takeaways";
  if (hasAction) return "action_plan";
  return "bullet_columns";
}

function chooseLayoutBias(sectionType, section, partIndex, totalParts) {
  const density = sectionDensityScore(section);
  if (sectionType === "table_analysis") {
    if (totalParts > 1) return partIndex === 0 ? "compare" : "sidecallout";
    return density >= 10 ? "dashboard" : "compare";
  }
  if (sectionType === "image_story") {
    if (totalParts > 1) return partIndex === 0 ? "storyboard" : "gallery";
    return "picture";
  }
  if (sectionType === "process_flow") return totalParts > 1 ? (partIndex === 0 ? "bridge" : "timeline") : "bridge";
  if (sectionType === "action_plan") return totalParts > 1 ? (partIndex === 0 ? "timeline" : "stack") : "dashboard";
  if (sectionType === "key_takeaways") return "wall";
  return density >= 8 ? "masonry" : "cards";
}

function chooseTableMode(sectionType, section, partIndex, totalParts) {
  if (sectionType !== "table_analysis") return "";
  if (totalParts > 1) return partIndex === 0 ? "compare" : "sidecallout";
  return sectionDensityScore(section) >= 11 ? "dense" : "compare";
}

function chooseImageMode(sectionType, section, partIndex, totalParts) {
  if (sectionType !== "image_story") return "";
  return totalParts > 1 ? (partIndex === 0 ? "storyboard" : "gallery") : "focus";
}

function buildTableInsights(section, lines, tableModel, screenshots = []) {
  const insights = makeCards(lines, 3, section);
  const caption = tableModel?.caption || "";
  if (caption && !insights.some((item) => item.title.includes("表"))) {
    insights.unshift({
      title: compactTitle(caption, 22),
      body: compactBody(tableModel?.highlight || caption, 92),
      detail: compactBody(tableModel?.highlight || caption, 92),
      accent: "darkGreen",
    });
  }
  if (screenshots[0]?.caption || screenshots[0]?.title) {
    const text = normalizeText(screenshots[0].caption || screenshots[0].title);
    if (text) {
      insights.push({
        title: compactTitle(text, 24),
        body: compactBody(text, 92),
        detail: compactBody(text, 92),
        accent: "oliveGreen",
      });
    }
  }
  return insights.slice(0, 3);
}

function buildSlideForSection(page, section, partIndex = 0, totalParts = 1) {
  const sectionType = classifySection(section);
  const chapterTitle = normalizeText(section.heading || section.title || "");
  const partTitle = totalParts > 1 ? `${chapterTitle} (${partIndex + 1}/${totalParts})` : chapterTitle;
  const lines = sectionLinesForPart(section, partIndex, totalParts);
  const summary = compactBody(lines[0] || collectPlainText(section), 120);
  const metrics = extractHighlights(lines, sectionType === "table_analysis" ? 4 : 3);
  const cards = makeCards(lines, sectionType === "table_analysis" ? 3 : 2, section);
  const { mainImage, screenshots } = splitImagesForPart(section, partIndex, totalParts);
  const tables = splitTablesForPart(section, partIndex, totalParts);
  const primaryTableModel = tables[0] || primaryTable(section);
  const densityScore = sectionDensityScore(section);
  const density = densityScore >= 12 ? "high" : densityScore >= 7 ? "medium" : "low";
  const pageRole =
    sectionType === "table_analysis"
      ? "table"
      : sectionType === "image_story"
        ? "image"
        : sectionType === "process_flow"
          ? "process"
          : sectionType === "action_plan"
            ? "action"
            : "bullet";
  const layoutBias = chooseLayoutBias(sectionType, section, partIndex, totalParts);
  const tableMode = chooseTableMode(sectionType, section, partIndex, totalParts);
  const imageMode = chooseImageMode(sectionType, section, partIndex, totalParts);
  const preferredFamilies =
    pageRole === "table"
      ? ["table", "comparison", "dashboard", "visual"]
      : pageRole === "image"
        ? ["image", "storyboard", "gallery"]
        : pageRole === "process"
          ? ["process", "bridge", "timeline"]
          : pageRole === "action"
            ? ["action", "timeline", "cards"]
            : ["bullet", "cards", "masonry"];

  if (sectionType === "table_analysis") {
    const insights = buildTableInsights(section, lines, primaryTableModel, screenshots);
    return {
      page,
      type: "table_analysis",
      title: partTitle,
      sectionHeading: section.heading,
      sectionId: section.id,
      sectionIndex: section.index,
      partIndex,
      partCount: totalParts,
      pageRole,
      density,
      layoutTier: density === "high" || totalParts > 1 ? "high" : "medium",
      layoutBias,
      tableMode,
      imageMode,
      preferredFamilies,
      summary,
      headline: summary,
      footer: compactBody(section.heading || section.title || "", 64),
      keyPoints: metrics.map((item) => `${item.label}:${item.value}`),
      metrics,
      focusItems: metrics,
      tables,
      table: primaryTableModel,
      image: mainImage,
      images: mainImage ? [mainImage] : [],
      screenshots,
      insights,
      callouts: insights,
      cards: insights,
      takeaways: insights,
      bars: metrics.map((metric) => ({
        label: metric.label,
        value: metric.value,
      })),
      textBlocks: makeBullets(lines, 6),
      contentMode:
        primaryTableModel && (mainImage || screenshots.length)
          ? "mixed"
          : primaryTableModel
            ? "table-heavy"
            : mainImage || screenshots.length
              ? "image-heavy"
              : "balanced",
      note: `本页围绕“${section.title || chapterTitle}”展开，优先保留表格、截图和结论的对应关系。`,
    };
  }

  if (sectionType === "image_story") {
    return {
      page,
      type: "image_story",
      title: partTitle,
      sectionHeading: section.heading,
      sectionId: section.id,
      sectionIndex: section.index,
      partIndex,
      partCount: totalParts,
      pageRole,
      density,
      layoutTier: density === "high" ? "medium" : "low",
      layoutBias,
      imageMode,
      preferredFamilies,
      summary,
      headline: summary,
      footer: compactBody(section.heading || section.title || "", 64),
      metrics,
      keyPoints: metrics.map((item) => `${item.label}:${item.value}`),
      image: mainImage,
      images: mainImage ? [mainImage] : [],
      screenshots,
      callouts: cards,
      cards,
      bullets: makeBullets(lines, 5),
      textBlocks: makeBullets(lines, 6),
      contentMode: screenshots.length || mainImage ? "image-heavy" : "balanced",
      note: `本页围绕“${section.title || chapterTitle}”展开，优先保留图片与说明的配合关系。`,
    };
  }

  if (sectionType === "process_flow") {
    const stages = makeStages(lines, 4);
    return {
      page,
      type: "process_flow",
      title: partTitle,
      sectionHeading: section.heading,
      sectionId: section.id,
      sectionIndex: section.index,
      partIndex,
      partCount: totalParts,
      pageRole,
      density,
      layoutTier: density === "high" ? "high" : "medium",
      layoutBias,
      preferredFamilies,
      summary,
      headline: summary,
      footer: compactBody(section.heading || section.title || "", 64),
      stages,
      steps: stages,
      metrics,
      focusItems: metrics,
      image: mainImage,
      images: mainImage ? [mainImage] : [],
      screenshots,
      callouts: cards,
      notes: cards,
      contentMode: mainImage || screenshots.length ? "mixed" : "balanced",
      note: `本页围绕“${section.title || chapterTitle}”展开，优先表达方法路径和推进步骤。`,
    };
  }

  if (sectionType === "action_plan") {
    const timeline = makeStages(lines, 3);
    return {
      page,
      type: "action_plan",
      title: partTitle,
      sectionHeading: section.heading,
      sectionId: section.id,
      sectionIndex: section.index,
      partIndex,
      partCount: totalParts,
      pageRole,
      density,
      layoutTier: density === "high" ? "high" : "medium",
      layoutBias,
      preferredFamilies,
      summary,
      headline: summary,
      footer: compactBody(section.heading || section.title || "", 64),
      timeline,
      steps: timeline,
      metrics,
      focusItems: metrics,
      image: mainImage,
      images: mainImage ? [mainImage] : [],
      screenshots,
      callouts: cards,
      cards,
      contentMode: "action",
      note: `本页围绕“${section.title || chapterTitle}”展开，优先表达行动安排与闭环节奏。`,
    };
  }

  if (sectionType === "key_takeaways") {
    return {
      page,
      type: "key_takeaways",
      title: partTitle,
      sectionHeading: section.heading,
      sectionId: section.id,
      sectionIndex: section.index,
      partIndex,
      partCount: totalParts,
      pageRole,
      density,
      layoutTier: "medium",
      layoutBias,
      preferredFamilies,
      summary,
      headline: summary,
      footer: compactBody(section.heading || section.title || "", 64),
      metrics,
      takeaways: cards,
      cards,
      image: mainImage,
      images: mainImage ? [mainImage] : [],
      screenshots,
      contentMode: "balanced",
      note: `本页围绕“${section.title || chapterTitle}”展开，优先收束核心结论。`,
    };
  }

  const columns = makeColumns(lines, totalParts > 1 ? 3 : 2);
  return {
    page,
    type: "bullet_columns",
    title: partTitle,
    sectionHeading: section.heading,
    sectionId: section.id,
    sectionIndex: section.index,
    partIndex,
    partCount: totalParts,
    pageRole,
    density,
    layoutTier: density === "high" ? "medium" : "low",
    layoutBias,
    preferredFamilies,
    summary,
    headline: summary,
    footer: compactBody(section.heading || section.title || "", 64),
    columns,
    metrics,
    focusItems: metrics,
    image: mainImage,
    images: mainImage ? [mainImage] : [],
    screenshots,
    callouts: cards,
    cards,
    takeaways: cards.slice(0, 3),
    contentMode: "text-heavy",
    note: `本页围绕“${section.title || chapterTitle}”展开，优先归纳文字要点和分栏内容。`,
  };
}

function detectSections(documentStructure, doc = {}) {
  if (documentStructure?.topLevel?.length) return documentStructure.topLevel;
  return buildDocumentStructure(doc).topLevel || [];
}

function buildOutline(doc = {}, options = {}) {
  const pageCount = clampPageCount(options.pages ?? options.pageCount);
  const documentStructure = options.documentStructure || buildDocumentStructure(doc);
  const seenImageKeys = new Set();
  const sections = detectSections(documentStructure, doc).map((section, index) => {
    const heading = normalizeText(section.heading || section.title || "");
    const images = uniqueImages(section.images || []).filter((image) => {
      const key = imageKey(image);
      if (!key || seenImageKeys.has(key)) return false;
      seenImageKeys.add(key);
      return true;
    });
    return {
      ...section,
      index: Number(section.index || index + 1),
      title: stripHeadingPrefix(heading) || heading,
      heading,
      images,
      allHeadings: flattenHeadings(section),
      allParagraphs: flattenParagraphs(section),
      markdown: collectPlainText(section),
    };
  });

  const coverTitle = normalizeText(options.title || doc.title || documentStructure.title || sections[0]?.title || "汇报材料");
  const coverSubtitle = normalizeText(options.subtitle || doc.subtitle || "");
  const coverDepartment = normalizeText(options.department || "");
  const coverPresenter = normalizeText(options.presenter || "");
  const coverDate = normalizeText(options.date || options.reportDate || "");

  const requestedContentPages = pageCount > 0 ? Math.max(1, pageCount - 1) : Math.max(1, sections.length);
  const targetContentPages = Math.max(requestedContentPages, sections.length);
  const allocations = allocateContentPages(sections, targetContentPages);

  let page = 1;
  const slides = [
    {
      page: page++,
      type: "cover",
      title: coverTitle,
      subtitle: coverSubtitle,
      department: coverDepartment,
      presenter: coverPresenter,
      reportDate: coverDate,
      note: "封面页",
      pageRole: "cover",
      density: "low",
      layoutBias: "cover",
      preferredFamilies: ["cover"],
    },
  ];

  allocations.forEach(({ section, pages }) => {
    for (let partIndex = 0; partIndex < pages; partIndex += 1) {
      slides.push(buildSlideForSection(page++, section, partIndex, pages));
    }
  });

  return {
    meta: {
      title: coverTitle,
      pages: slides.length,
      requestedPages: pageCount || slides.length,
      contentPages: slides.length - 1,
      sectionCount: sections.length,
      topLevelCount: sections.length,
    },
    slides,
  };
}

function buildNotes(outline = {}) {
  return (outline.slides || [])
    .map((slide) => `第${slide.page}页：${slide.title}${slide.note ? `\n${slide.note}` : ""}`)
    .join("\n\n");
}

module.exports = {
  clampPageCount,
  normalizeTableStructure,
  detectSections,
  buildOutline,
  buildNotes,
  estimateImageAspect,
};
