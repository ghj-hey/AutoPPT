const path = require("node:path");

const CN_TOP_LEVEL = "\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u5343";
const TOP_LEVEL_HEADING_RE = new RegExp(`^[${CN_TOP_LEVEL}]+(?:\u3001|\.)\\s*`);
const NUMERIC_HEADING_RE = /^\d+(?:\.\d+){0,4}(?:[.)\u3001]|[\uff1a:]\s*|\s+)/;
const FIGURE_CAPTION_RE = /^\u56fe\s*\d+\s*[:\uff1a]/;
const TABLE_CAPTION_RE = /^\u8868\s*\d+\s*[:\uff1a]/;
const PAGE_NOISE_RE = /^(?:\u7b2c\s*\d+\s*\u9875|page\s*\d+)$/i;
const NOISE_TITLE_RE = /^(?:\u76ee\u5f55|\u9644\u5f55|\u53c2\u8003\u6587\u732e|\u56fe\u76ee\u5f55|\u8868\u76ee\u5f55|\u81f4\u8c22|\u6458\u8981|\u5173\u952e\u5b57)$/;
const IMAGE_HINT_RE = /(?:\u56fe\s*\d+|\u622a\u56fe|\u754c\u9762|\u7cfb\u7edf|\u793a\u610f\u56fe|\u6d41\u7a0b\u56fe|\u53cd\u9988|\u540d\u5355|\u8bd5\u70b9)/;

function normalizeText(text) {
  return String(text || "")
    .replace(/\u0000/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function stripHeadingPrefix(text) {
  return normalizeText(text)
    .replace(TOP_LEVEL_HEADING_RE, "")
    .replace(NUMERIC_HEADING_RE, "")
    .trim();
}

function isNoiseLine(text) {
  const line = normalizeText(text);
  if (!line) return true;
  if (PAGE_NOISE_RE.test(line)) return true;
  if (NOISE_TITLE_RE.test(line)) return true;
  if (/^(TOC|HYPERLINK|PAGEREF)\b/i.test(line)) return true;
  if (/^\d+$/.test(line)) return true;
  return false;
}

function inferHeadingLevel(node = {}) {
  const line = normalizeText(node.text || node);
  const styleId = String(node.styleId || "");
  const outlineLevel = Number.isFinite(node.outlineLevel) ? Number(node.outlineLevel) : -1;

  if (!line || isNoiseLine(line)) return null;
  if (/^[-\u2014]{1,}/.test(line)) return null;
  if (FIGURE_CAPTION_RE.test(line) || TABLE_CAPTION_RE.test(line)) return null;
  if (line.length > 90) return null;

  if (/heading|title|\u6807\u9898/i.test(styleId)) {
    return Math.max(1, Math.min(5, outlineLevel >= 0 ? outlineLevel + 1 : 1));
  }
  if (outlineLevel >= 0 && outlineLevel <= 4) return outlineLevel + 1;
  if (TOP_LEVEL_HEADING_RE.test(line)) return 1;

  if (NUMERIC_HEADING_RE.test(line)) {
    const match = line.match(/^(\d+(?:\.\d+){0,4})/);
    const depth = match ? match[1].split(".").length : 1;
    return Math.min(5, depth + 1);
  }

  const standalone = !/[\uff0c\u3002\uff1b\uff1a\u3001\u201c\u201d\u2018\u2019\uff08\uff09\u3010\u3011]/.test(line);
  if (standalone && line.length >= 4 && line.length <= 28) return null;
  return null;
}

function looksLikeImplicitTopLevelHeading(line, previousBlock, nextBlock) {
  const text = normalizeText(line);
  const previousText = normalizeText(previousBlock?.text || "");
  const nextText = normalizeText(nextBlock?.text || "");

  if (!text) return false;
  if (TOP_LEVEL_HEADING_RE.test(text)) return true;
  if (FIGURE_CAPTION_RE.test(text) || TABLE_CAPTION_RE.test(text)) return false;
  if (/^[-\u2014]{1,}/.test(text)) return false;
  if (/[\uff0c\u3002\uff1b\uff1a\u3001\u201c\u201d\u2018\u2019\uff08\uff09\u3010\u3011]/.test(text)) return false;
  if (text.length < 4 || text.length > 26) return false;
  if (!nextText || nextText.length < 18) return false;

  if (!previousText) return true;
  if (/^[-\u2014]{1,}/.test(previousText)) return true;
  if (!/[\uff0c\u3002\uff1b\uff1a\u3001\u201c\u201d\u2018\u2019\uff08\uff09\u3010\u3011]/.test(previousText) && previousText.length <= 24) return true;
  return previousText.length >= 20 || TABLE_CAPTION_RE.test(previousText);
}

function createSectionNode({ id, index, level, heading, parent = null, sourceIndex = 0 }) {
  const cleanHeading = normalizeText(heading);
  return {
    id,
    index,
    level,
    heading: cleanHeading,
    title: stripHeadingPrefix(cleanHeading) || cleanHeading,
    parentId: parent?.id || null,
    sourceIndex,
    path: [],
    paragraphs: [],
    children: [],
    tables: [],
    images: [],
    meta: {
      startIndex: sourceIndex,
      endIndex: sourceIndex,
      charCount: 0,
      paragraphCount: 0,
      subsectionCount: 0,
      imageHints: 0,
    },
  };
}

function walkSections(section, visitor) {
  if (!section) return;
  visitor(section);
  (section.children || []).forEach((child) => walkSections(child, visitor));
}

function collectAllParagraphs(section) {
  const paragraphs = [];
  walkSections(section, (node) => {
    (node.paragraphs || []).forEach((item) => paragraphs.push(item));
  });
  return paragraphs.filter(Boolean);
}

function collectAllHeadings(section) {
  const headings = [];
  walkSections(section, (node) => {
    if (node !== section && node.heading) headings.push(node.heading);
  });
  return headings.filter(Boolean);
}

function collectPlainText(section) {
  return [section.heading, ...collectAllHeadings(section), ...collectAllParagraphs(section)]
    .filter(Boolean)
    .join("\n");
}

function buildSectionMarkdown(section, depth = 1) {
  const lines = [];
  if (section.heading) {
    lines.push(`${"#".repeat(Math.min(6, depth))} ${section.heading}`);
    lines.push("");
  }
  (section.paragraphs || []).forEach((paragraph) => {
    lines.push(paragraph);
    lines.push("");
  });
  (section.tables || []).forEach((table, index) => {
    lines.push(`- [TABLE ${index + 1}] rows=${(table.rows || []).length}`);
  });
  if ((section.tables || []).length) lines.push("");
  (section.images || []).forEach((image, index) => {
    lines.push(`- [IMAGE ${index + 1}] ${path.basename(image.path || "")}`);
  });
  if ((section.images || []).length) lines.push("");
  (section.children || []).forEach((child) => {
    lines.push(buildSectionMarkdown(child, depth + 1));
    lines.push("");
  });
  return lines.join("\n").trim();
}

function collectSourceBlocks(doc = {}) {
  if (Array.isArray(doc.blocks) && doc.blocks.length) {
    return doc.blocks.map((block, index) => ({
      ...block,
      sourceIndex: Number.isFinite(block.sourceIndex) ? block.sourceIndex : index,
      text: normalizeText(block.text || ""),
    }));
  }

  const paragraphs = (doc.paragraphNodes || []).map((node, index) => ({
    kind: "paragraph",
    sourceIndex: Number.isFinite(node.sourceIndex) ? node.sourceIndex : index,
    text: normalizeText(node.text || ""),
    styleId: node.styleId || "",
    outlineLevel: Number.isFinite(node.outlineLevel) ? node.outlineLevel : -1,
  }));
  const tables = (doc.tables || []).map((table, index) => ({
    kind: "table",
    sourceIndex: Number.isFinite(table.sourceIndex) ? table.sourceIndex : 100000 + index,
    index: table.index || index + 1,
    rows: table.rows || [],
    caption: table.caption || "",
  }));
  return [...paragraphs, ...tables].sort((left, right) => left.sourceIndex - right.sourceIndex);
}

function extractReferencedNumbers(text, prefix) {
  const regexp = new RegExp(`${prefix}\\s*([0-9]+)`, "g");
  const matches = [...normalizeText(text).matchAll(regexp)];
  return [...new Set(matches.map((item) => Number(item[1])).filter((value) => Number.isFinite(value) && value > 0))];
}

function imageKey(image = {}) {
  return normalizeText(image.path || image.source || image.caption || image.title || image.name || "");
}

function uniqueImages(images = []) {
  const seen = new Set();
  return (images || []).filter((image) => {
    const key = imageKey(image);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function assignExplicitImages(topLevel, images = []) {
  const normalizedImages = uniqueImages(images);
  const byIndex = new Map(normalizedImages.map((image, index) => [Number(image.index || index + 1), image]));
  const used = new Set();

  (topLevel || []).forEach((section) => {
    const refs = extractReferencedNumbers(collectPlainText(section), "\u56fe");
    refs.forEach((refNo) => {
      const image = byIndex.get(refNo);
      const key = imageKey(image);
      if (!image || !key || used.has(key)) return;
      section.images.push(image);
      used.add(key);
    });
  });

  return normalizedImages.filter((image) => !used.has(imageKey(image)));
}

function assignSequentialImages(topLevel, images = []) {
  if (!(topLevel || []).length || !(images || []).length) return;
  const normalizedImages = uniqueImages(images);

  const rankedSections = (topLevel || []).map((section, index) => ({
    section,
    index,
    score:
      Number(section.meta?.imageHints || 0) * 3 +
      (IMAGE_HINT_RE.test(section.heading || "") ? 2 : 0) +
      Math.min(3, Math.floor(Number(section.meta?.charCount || 0) / 200)),
  }));

  let cursor = 0;
  normalizedImages.forEach((image) => {
    const ranked = rankedSections
      .slice(cursor)
      .map((item, offset) => ({ ...item, absoluteIndex: cursor + offset }))
      .sort((left, right) => right.score - left.score || left.absoluteIndex - right.absoluteIndex);
    const chosen = ranked[0] || rankedSections[cursor] || rankedSections[rankedSections.length - 1];
    if (!chosen) return;
    const alreadyExists = (chosen.section.images || []).some((item) => imageKey(item) === imageKey(image));
    if (!alreadyExists) {
      chosen.section.images.push(image);
    }
    cursor = Math.min(rankedSections.length - 1, chosen.absoluteIndex);
  });
}

function addSyntheticTopLevelPrefixes(topLevel = []) {
  const numerals = CN_TOP_LEVEL.split("");
  (topLevel || []).forEach((section, index) => {
    const heading = normalizeText(section.heading || "");
    if (TOP_LEVEL_HEADING_RE.test(heading)) return;
    const numeral = numerals[index] || `${index + 1}`;
    const prefixed = `${numeral}\u3001${stripHeadingPrefix(heading) || heading}`;
    section.heading = prefixed;
    section.title = stripHeadingPrefix(prefixed) || prefixed;
  });
}

function findParentNode(roots, targetParentId) {
  if (!targetParentId) return null;
  const queue = [...roots];
  while (queue.length) {
    const current = queue.shift();
    if (current.id === targetParentId) return current;
    queue.push(...(current.children || []));
  }
  return null;
}

function buildDocumentStructure(doc = {}) {
  const blocks = collectSourceBlocks(doc);
  const roots = [];
  const stack = [];
  let sectionCounter = 0;

  blocks.forEach((block, index) => {
    const previous = blocks[index - 1];
    const next = blocks[index + 1];

    if (block.kind === "table") {
      const current = stack[stack.length - 1];
      if (current) {
        current.tables.push(block);
        current.meta.endIndex = block.sourceIndex;
      }
      return;
    }

    const text = normalizeText(block.text || "");
    if (!text || isNoiseLine(text)) return;

    let level = inferHeadingLevel(block);
    if (looksLikeImplicitTopLevelHeading(text, previous, next)) {
      level = 1;
    }

    if (level) {
      while (stack.length && stack[stack.length - 1].level >= level) stack.pop();
      const parent = stack[stack.length - 1] || null;
      const section = createSectionNode({
        id: `section-${++sectionCounter}`,
        index: sectionCounter,
        level,
        heading: text,
        parent,
        sourceIndex: block.sourceIndex,
      });
      if (parent) {
        parent.children.push(section);
        parent.meta.subsectionCount += 1;
      } else {
        roots.push(section);
      }
      stack.push(section);
      return;
    }

    const current = stack[stack.length - 1];
    if (!current) return;
    current.paragraphs.push(text);
    current.meta.endIndex = block.sourceIndex;
    current.meta.paragraphCount += 1;
    current.meta.charCount += text.length;
    if (IMAGE_HINT_RE.test(text)) current.meta.imageHints += 1;
  });

  const explicitTopLevel = roots.filter((item) => item.level === 1 && TOP_LEVEL_HEADING_RE.test(normalizeText(item.heading || "")));
  const inferredTopLevel = roots.filter((item) => item.level === 1);
  const coverLikeLead =
    inferredTopLevel.length > 1 &&
    inferredTopLevel[0] &&
    Number(inferredTopLevel[0].meta?.paragraphCount || 0) <= 1 &&
    !(inferredTopLevel[0].tables || []).length &&
    !(inferredTopLevel[0].images || []).length;
  let topLevel = explicitTopLevel.length >= 2 ? explicitTopLevel : inferredTopLevel;
  if (coverLikeLead && explicitTopLevel.length < 2 && inferredTopLevel.length > 1) {
    topLevel = inferredTopLevel.slice(1);
  }
  if (topLevel.length && explicitTopLevel.length < 2) addSyntheticTopLevelPrefixes(topLevel);

  const explicitRemaining = assignExplicitImages(topLevel, doc.images || []);
  assignSequentialImages(topLevel, explicitRemaining);

  const allSections = [];
  roots.forEach((root) => {
    walkSections(root, (section) => {
      const segments = [];
      let cursor = section;
      while (cursor) {
        if (cursor.heading) segments.unshift(cursor.heading);
        cursor = findParentNode(roots, cursor.parentId);
      }
      section.path = segments;
      allSections.push(section);
    });
  });

  const markdown = roots.map((item) => buildSectionMarkdown(item, 1)).filter(Boolean).join("\n\n");

  return {
    sections: roots,
    topLevel,
    markdown,
    counts: {
      sections: allSections.length,
      topLevelSections: topLevel.length,
      paragraphs: allSections.reduce((sum, item) => sum + Number(item.meta?.paragraphCount || 0), 0),
      tables: allSections.reduce((sum, item) => sum + (item.tables || []).length, 0),
      images: allSections.reduce((sum, item) => sum + (item.images || []).length, 0),
    },
  };
}

module.exports = {
  normalizeText,
  stripHeadingPrefix,
  inferHeadingLevel,
  buildDocumentStructure,
  buildSectionMarkdown,
  collectAllParagraphs,
  collectAllHeadings,
  collectPlainText,
};
