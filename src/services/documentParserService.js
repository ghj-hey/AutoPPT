const fs = require("node:fs");
const path = require("node:path");
const JSZip = require("jszip");
const iconv = require("iconv-lite");
const { XMLParser } = require("fast-xml-parser");

let pdfParseModule = null;

function normalizeMatrixInit(init) {
  if (Array.isArray(init) && init.length >= 6) return init;
  if (init && typeof init === "object") {
    return [
      Number(init.a ?? 1),
      Number(init.b ?? 0),
      Number(init.c ?? 0),
      Number(init.d ?? 1),
      Number(init.e ?? 0),
      Number(init.f ?? 0),
    ];
  }
  return [1, 0, 0, 1, 0, 0];
}

function ensurePdfCanvasGlobals() {
  if (typeof globalThis.DOMMatrix === "undefined") {
    globalThis.DOMMatrix = class DOMMatrix {
      constructor(init) {
        const [a, b, c, d, e, f] = normalizeMatrixInit(init);
        this.a = a;
        this.b = b;
        this.c = c;
        this.d = d;
        this.e = e;
        this.f = f;
      }

      multiplySelf(other) {
        const matrix = normalizeMatrixInit(other);
        const a = this.a * matrix[0] + this.c * matrix[1];
        const b = this.b * matrix[0] + this.d * matrix[1];
        const c = this.a * matrix[2] + this.c * matrix[3];
        const d = this.b * matrix[2] + this.d * matrix[3];
        const e = this.a * matrix[4] + this.c * matrix[5] + this.e;
        const f = this.b * matrix[4] + this.d * matrix[5] + this.f;
        this.a = a;
        this.b = b;
        this.c = c;
        this.d = d;
        this.e = e;
        this.f = f;
        return this;
      }

      preMultiplySelf(other) {
        const matrix = normalizeMatrixInit(other);
        const a = matrix[0] * this.a + matrix[2] * this.b;
        const b = matrix[1] * this.a + matrix[3] * this.b;
        const c = matrix[0] * this.c + matrix[2] * this.d;
        const d = matrix[1] * this.c + matrix[3] * this.d;
        const e = matrix[0] * this.e + matrix[2] * this.f + matrix[4];
        const f = matrix[1] * this.e + matrix[3] * this.f + matrix[5];
        this.a = a;
        this.b = b;
        this.c = c;
        this.d = d;
        this.e = e;
        this.f = f;
        return this;
      }

      translate(tx = 0, ty = 0) {
        this.e += Number(tx) || 0;
        this.f += Number(ty) || 0;
        return this;
      }

      scale(scaleX = 1, scaleY = scaleX) {
        this.a *= Number(scaleX) || 1;
        this.b *= Number(scaleX) || 1;
        this.c *= Number(scaleY) || 1;
        this.d *= Number(scaleY) || 1;
        return this;
      }

      invertSelf() {
        const det = this.a * this.d - this.b * this.c;
        if (!det) return this;
        const a = this.d / det;
        const b = -this.b / det;
        const c = -this.c / det;
        const d = this.a / det;
        const e = (this.c * this.f - this.d * this.e) / det;
        const f = (this.b * this.e - this.a * this.f) / det;
        this.a = a;
        this.b = b;
        this.c = c;
        this.d = d;
        this.e = e;
        this.f = f;
        return this;
      }
    };
  }

  if (typeof globalThis.ImageData === "undefined") {
    globalThis.ImageData = class ImageData {
      constructor(dataOrWidth, widthOrHeight, height) {
        if (ArrayBuffer.isView(dataOrWidth)) {
          this.data = dataOrWidth;
          this.width = Number(widthOrHeight) || 0;
          this.height = Number(height) || 0;
          return;
        }
        this.width = Number(dataOrWidth) || 0;
        this.height = Number(widthOrHeight) || 0;
        this.data = new Uint8ClampedArray(this.width * this.height * 4);
      }
    };
  }

  if (typeof globalThis.Path2D === "undefined") {
    globalThis.Path2D = class Path2D {
      constructor() {}
      addPath() {}
      closePath() {}
      moveTo() {}
      lineTo() {}
      bezierCurveTo() {}
      quadraticCurveTo() {}
      rect() {}
      arc() {}
    };
  }
}

function loadPdfParse() {
  ensurePdfCanvasGlobals();
  if (!pdfParseModule) {
    pdfParseModule = require("pdf-parse");
  }
  return pdfParseModule;
}

const EMU_PER_INCH = 914400;
const XML_OPTIONS = {
  ignoreAttributes: false,
  trimValues: false,
};
const ORDERED_XML_OPTIONS = {
  ignoreAttributes: false,
  trimValues: false,
  preserveOrder: true,
};

function decodeXml(text) {
  const named = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
  return String(text || "")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, num) => String.fromCodePoint(parseInt(num, 10)))
    .replace(/&([a-zA-Z]+);/g, (_, name) => named[name] || `&${name};`);
}

function cleanText(text) {
  return decodeXml(String(text || ""))
    .replace(/\r/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\n+/g, " ")
    .trim();
}

function toArray(value) {
  if (value == null) return [];
  return Array.isArray(value) ? value : [value];
}

function collectText(node) {
  if (node == null) return "";
  if (typeof node === "string" || typeof node === "number") return String(node);
  if (Array.isArray(node)) return node.map((item) => collectText(item)).join("");
  return Object.entries(node)
    .filter(([key]) => !key.startsWith("@_"))
    .map(([, value]) => collectText(value))
    .join("");
}

function collectOrderedText(node) {
  if (node == null) return "";
  if (typeof node === "string" || typeof node === "number") return String(node);
  if (Array.isArray(node)) return node.map((item) => collectOrderedText(item)).join("");
  return Object.entries(node)
    .filter(([key]) => !key.startsWith("@_"))
    .map(([key, value]) => {
      if (key === "#text") return String(value || "");
      return collectOrderedText(value);
    })
    .join("");
}

async function loadZip(filePath) {
  return JSZip.loadAsync(fs.readFileSync(filePath));
}

async function readEntry(zip, entryPath) {
  const file = zip.file(entryPath);
  if (!file) throw new Error(`Missing ZIP entry: ${entryPath}`);
  const buffer = await file.async("nodebuffer");
  if (!buffer || !buffer.length) return "";
  let text = "";
  if (iconv.encodingExists("utf8")) {
    text = iconv.decode(buffer, "utf8");
  } else {
    text = Buffer.from(buffer).toString("utf8");
  }
  return text.replace(/^\uFEFF/, "");
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function normalizeParagraphRecord(paragraphNode, index) {
  const text = cleanText(collectText(paragraphNode));
  const pPr = paragraphNode["w:pPr"] || {};
  const styleId = pPr["w:pStyle"]?.["@_w:val"] || "";
  const outlineLevel = Number(pPr["w:outlineLvl"]?.["@_w:val"] ?? -1);
  return {
    kind: "paragraph",
    index,
    sourceIndex: index,
    text,
    styleId,
    outlineLevel: Number.isFinite(outlineLevel) ? outlineLevel : -1,
  };
}

function getOrderedChildren(nodes, name) {
  return toArray(nodes).flatMap((item) =>
    Object.prototype.hasOwnProperty.call(item || {}, name) ? toArray(item[name]) : [],
  );
}

function normalizeTableCellText(cellNode) {
  if (!cellNode) return "";
  const paragraphs = toArray(cellNode["w:p"]);
  const parts = paragraphs
    .map((paragraph) => cleanText(collectText(paragraph)))
    .filter(Boolean);
  if (parts.length) return parts.join(" ");
  return cleanText(collectText(cellNode));
}

function parseTableNode(tableNode, index) {
  const rawRows = toArray(tableNode?.["w:tr"]);
  const rows = rawRows
    .map((rowNode) => {
      const cells = toArray(rowNode?.["w:tc"]).map((cellNode) => normalizeTableCellText(cellNode));
      return cells.filter((cell, cellIndex) => cell !== "" || cellIndex < cells.length - 1);
    })
    .filter((row) => row.some(Boolean));

  return {
    kind: "table",
    index,
    sourceIndex: index,
    rows,
  };
}

async function extractDocx(wordPath, assetDir) {
  const zip = await loadZip(wordPath);
  const parser = new XMLParser(XML_OPTIONS);
  const orderedParser = new XMLParser(ORDERED_XML_OPTIONS);
  const xmlText = await readEntry(zip, "word/document.xml");
  const xml = parser.parse(xmlText);
  const orderedXml = orderedParser.parse(xmlText);
  const body = xml["w:document"]?.["w:body"] || {};
  const orderedDocument = getOrderedChildren(orderedXml, "w:document");
  const orderedBody = getOrderedChildren(orderedDocument, "w:body");
  const tableNodes = toArray(body["w:tbl"]);

  const paragraphNodes = toArray(body["w:p"])
    .map((node, index) => normalizeParagraphRecord(node, index))
    .filter((item) => item.text);
  const paragraphs = paragraphNodes.map((item) => item.text);

  const tables = [];
  const blocks = [];
  let paragraphCursor = 0;
  let tableCursor = 0;

  orderedBody.forEach((entry, blockIndex) => {
    if (entry["w:p"]) {
      const paragraph = paragraphNodes[paragraphCursor];
      paragraphCursor += 1;
      if (paragraph?.text) {
        blocks.push({
          ...paragraph,
          sourceIndex: blockIndex,
        });
      }
      return;
    }

    if (entry["w:tbl"]) {
      const tableNode = tableNodes[tableCursor] || {};
      const table = parseTableNode(tableNode, tableCursor + 1);
      tableCursor += 1;
      if (table.rows.length) {
        table.sourceIndex = blockIndex;
        tables.push(table);
        blocks.push(table);
      }
    }
  });

  ensureDir(assetDir);
  const images = [];
  const mediaEntries = Object.keys(zip.files)
    .filter((name) => name.startsWith("word/media/") && !zip.files[name].dir)
    .sort();
  for (let i = 0; i < mediaEntries.length; i += 1) {
    const src = mediaEntries[i];
    const ext = path.extname(src).toLowerCase() || ".bin";
    const dest = path.join(assetDir, `figure-${i + 1}${ext}`);
    fs.writeFileSync(dest, await zip.file(src).async("nodebuffer"));
    images.push({ index: i + 1, source: src, path: dest });
  }

  return {
    sourceType: "docx",
    sourcePath: wordPath,
    blocks,
    paragraphNodes,
    paragraphs,
    tables,
    images,
  };
}

function chunkPdfLines(lines) {
  const paragraphs = [];
  let current = [];

  lines.forEach((line) => {
    const text = cleanText(line);
    if (!text) {
      if (current.length) {
        paragraphs.push(current.join(""));
        current = [];
      }
      return;
    }
    if (/^(第\s*\d+\s*页|page\s*\d+)$/i.test(text)) return;
    current.push(text);
    if (/[。！？；]$/.test(text) || text.length >= 40) {
      paragraphs.push(current.join(""));
      current = [];
    }
  });

  if (current.length) paragraphs.push(current.join(""));
  return paragraphs.filter(Boolean);
}

function detectPdfTables(lines) {
  const tables = [];
  let current = [];

  lines.forEach((line) => {
    const parts = String(line || "")
      .split(/\s{2,}|\t+/)
      .map((item) => cleanText(item))
      .filter(Boolean);
    const looksLikeRow = parts.length >= 3 && parts.some((item) => /[\d.%xX]/.test(item));
    if (looksLikeRow) {
      current.push(parts);
      return;
    }
    if (current.length >= 2) {
      tables.push({ index: tables.length + 1, rows: current });
    }
    current = [];
  });

  if (current.length >= 2) {
    tables.push({ index: tables.length + 1, rows: current });
  }
  return tables;
}

async function extractPdf(pdfPath, assetDir) {
  const buffer = fs.readFileSync(pdfPath);
  const { PDFParse } = loadPdfParse();
  const parser = new PDFParse({ data: buffer });

  try {
    const [textResult, tableResult, imageResult] = await Promise.all([
      parser.getText(),
      parser.getTable().catch(() => ({ mergedTables: [] })),
      parser
        .getImage({
          imageThreshold: 120,
          imageDataUrl: false,
          imageBuffer: true,
        })
        .catch(() => ({ pages: [] })),
    ]);

    ensureDir(assetDir);

    const rawLines = String(textResult?.text || "")
      .split(/\r?\n/)
      .map((line) => cleanText(line))
      .filter(Boolean);

    const paragraphs = chunkPdfLines(rawLines);
    const paragraphNodes = paragraphs.map((text, index) => ({
      kind: "paragraph",
      index,
      sourceIndex: index,
      text,
      styleId: "",
      outlineLevel: -1,
    }));

    const mergedTables = toArray(tableResult?.mergedTables)
      .map((rows, index) => ({
        kind: "table",
        index: index + 1,
        sourceIndex: Number.MAX_SAFE_INTEGER - (toArray(tableResult?.mergedTables).length - index),
        rows: toArray(rows)
          .map((row) =>
            toArray(row)
              .map((cell) => cleanText(cell))
              .filter(Boolean),
          )
          .filter((row) => row.length),
      }))
      .filter((table) => table.rows.length >= 2);
    const tables = mergedTables.length ? mergedTables : detectPdfTables(rawLines);

    const images = [];
    toArray(imageResult?.pages).forEach((page) => {
      toArray(page?.images).forEach((image, index) => {
        if (!image?.data?.length) return;
        const ext = image.kind ? `.${String(image.kind).toLowerCase()}` : ".png";
        const dest = path.join(assetDir, `pdf-figure-p${page.pageNumber || 1}-${index + 1}${ext}`);
        fs.writeFileSync(dest, Buffer.from(image.data));
        images.push({
          index: images.length + 1,
          page: page.pageNumber || 1,
          path: dest,
          width: image.width || 0,
          height: image.height || 0,
          kind: image.kind || ext.slice(1),
        });
      });
    });

    const blocks = [
      ...paragraphNodes,
      ...tables.map((table, index) => ({
        kind: "table",
        index: index + 1,
        sourceIndex: paragraphNodes.length + index,
        rows: table.rows,
      })),
    ].sort((left, right) => Number(left.sourceIndex || 0) - Number(right.sourceIndex || 0));

    return {
      sourceType: "pdf",
      sourcePath: pdfPath,
      rawText: textResult?.text || "",
      pageTexts: toArray(textResult?.pages),
      blocks,
      paragraphNodes,
      paragraphs,
      tables,
      images,
    };
  } finally {
    await parser.destroy().catch(() => {});
  }
}

async function extractInputDocument(inputPath, assetDir) {
  const ext = path.extname(inputPath).toLowerCase();
  if (ext === ".pdf") return extractPdf(inputPath, assetDir);
  return extractDocx(inputPath, assetDir);
}

async function extractTemplate(templatePath) {
  const zip = await loadZip(templatePath);
  const parser = new XMLParser(XML_OPTIONS);
  const presentation = parser.parse(await readEntry(zip, "ppt/presentation.xml"));
  const theme = parser.parse(await readEntry(zip, "ppt/theme/theme1.xml"));
  const size = presentation["p:presentation"]["p:sldSz"];
  const fontScheme = theme["a:theme"]?.["a:themeElements"]?.["a:fontScheme"] || {};
  return {
    path: templatePath,
    widthInches: Number(size?.["@_cx"] || 0) / EMU_PER_INCH,
    heightInches: Number(size?.["@_cy"] || 0) / EMU_PER_INCH,
    headFontFace: fontScheme["a:majorFont"]?.["a:latin"]?.["@_typeface"] || "",
    bodyFontFace: fontScheme["a:minorFont"]?.["a:latin"]?.["@_typeface"] || "",
  };
}

function buildDocumentSummary(doc, inputPath = "") {
  const sampleParagraphs = (doc.paragraphs || []).filter(Boolean).slice(0, 10);
  return {
    sourceType: doc.sourceType || path.extname(inputPath || doc.sourcePath || "").replace(/^\./, "") || "unknown",
    sourcePath: inputPath || doc.sourcePath || "",
    counts: {
      paragraphs: (doc.paragraphs || []).length,
      tables: (doc.tables || []).length,
      images: (doc.images || []).length,
      sections: 0,
    },
    paragraphNodes: (doc.paragraphNodes || []).slice(0, 20),
    sampleParagraphs,
    tables: (doc.tables || []).map((table) => ({
      index: table.index,
      rows: (table.rows || []).length,
      columns: Math.max(...(table.rows || []).map((row) => row.length), 0),
      preview: (table.rows || []).slice(0, 4),
    })),
    images: (doc.images || []).map((image) => ({
      index: image.index,
      page: image.page || null,
      path: image.path,
    })),
  };
}

module.exports = {
  decodeXml,
  cleanText,
  toArray,
  collectText,
  ensureDir,
  extractDocx,
  extractPdf,
  extractInputDocument,
  extractTemplate,
  buildDocumentSummary,
};
