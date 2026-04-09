const fs = require("node:fs");
const path = require("node:path");
const JSZip = require("jszip");
const { XMLParser } = require("fast-xml-parser");
const { imageSize } = require("image-size");
const { buildRenderedPreviews } = require("../services/powerPointPreviewService");

const XML_OPTIONS = { ignoreAttributes: false, trimValues: false };
const parser = new XMLParser(XML_OPTIONS);
const EMU_PER_INCH = 914400;

function parseArgs(argv) {
  const args = {
    ppt: "",
    previewDir: "",
    out: path.resolve(process.cwd(), "reference_library", "default"),
  };

  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) continue;
    const key = token.slice(2);
    const value = argv[i + 1];
    if (!value || value.startsWith("--")) continue;
    i += 1;
    if (key === "out") args.out = path.resolve(process.cwd(), value);
    else if (key === "preview-dir") args.previewDir = path.resolve(process.cwd(), value);
    else args[key] = value;
  }

  if (!args.ppt) throw new Error("Missing required argument: --ppt");
  return args;
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function toArray(value) {
  if (value == null) return [];
  return Array.isArray(value) ? value : [value];
}

function decodeXml(text) {
  const named = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
  return String(text || "")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, num) => String.fromCodePoint(parseInt(num, 10)))
    .replace(/&([a-zA-Z]+);/g, (_, name) => named[name] || `&${name};`);
}

function cleanText(text) {
  return decodeXml(text)
    .replace(/\r/g, "")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function collectText(node) {
  if (node == null) return "";
  if (typeof node === "string" || typeof node === "number") return String(node);
  if (Array.isArray(node)) return node.map(collectText).join("");
  return Object.entries(node)
    .filter(([key]) => !key.startsWith("@_"))
    .map(([, value]) => collectText(value))
    .join("");
}

function getHex(node) {
  if (!node || typeof node !== "object") return null;
  if (node["a:srgbClr"]?.["@_val"]) return node["a:srgbClr"]["@_val"].toUpperCase();
  if (node["a:schemeClr"]?.["@_val"]) return `scheme:${node["a:schemeClr"]["@_val"]}`;
  return null;
}

function getFillHex(spPr) {
  if (!spPr) return null;
  return getHex(spPr["a:solidFill"]);
}

function getLineHex(spPr) {
  if (!spPr?.["a:ln"]) return null;
  return getHex(spPr["a:ln"]["a:solidFill"]);
}

function getTransform(node) {
  const xfrm = node?.["a:xfrm"];
  if (!xfrm) return null;
  return {
    x: Number(xfrm["a:off"]?.["@_x"] || 0),
    y: Number(xfrm["a:off"]?.["@_y"] || 0),
    cx: Number(xfrm["a:ext"]?.["@_cx"] || 0),
    cy: Number(xfrm["a:ext"]?.["@_cy"] || 0),
    xIn: Number(xfrm["a:off"]?.["@_x"] || 0) / EMU_PER_INCH,
    yIn: Number(xfrm["a:off"]?.["@_y"] || 0) / EMU_PER_INCH,
    wIn: Number(xfrm["a:ext"]?.["@_cx"] || 0) / EMU_PER_INCH,
    hIn: Number(xfrm["a:ext"]?.["@_cy"] || 0) / EMU_PER_INCH,
  };
}

function loadZip(filePath) {
  return JSZip.loadAsync(fs.readFileSync(filePath));
}

async function readText(zip, entryPath) {
  const entry = zip.file(entryPath);
  if (!entry) return "";
  return entry.async("string");
}

async function readBuffer(zip, entryPath) {
  const entry = zip.file(entryPath);
  if (!entry) return null;
  return entry.async("nodebuffer");
}

function countMapToSorted(map, limit = Infinity) {
  return [...map.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, limit)
    .map(([name, count]) => ({ name, count }));
}

function bump(map, key, amount = 1) {
  if (!key) return;
  map.set(key, (map.get(key) || 0) + amount);
}

function normalizeTarget(baseFile, relTarget) {
  const baseDir = path.posix.dirname(baseFile);
  return path.posix.normalize(path.posix.join(baseDir, relTarget));
}

function readSvgDims(text) {
  const viewBox = text.match(/viewBox="([^"]+)"/i);
  if (viewBox) {
    const parts = viewBox[1].trim().split(/\s+/).map(Number);
    if (parts.length === 4) {
      return { width: parts[2], height: parts[3] };
    }
  }
  const width = text.match(/\bwidth="([0-9.]+)"/i);
  const height = text.match(/\bheight="([0-9.]+)"/i);
  if (width && height) {
    return { width: Number(width[1]), height: Number(height[1]) };
  }
  return { width: 0, height: 0 };
}

function detectMediaCategory(media) {
  const ext = path.extname(media.name).toLowerCase();
  const slides = media.slides?.length || 0;
  const dims = media.dimensions || { width: 0, height: 0 };
  const maxSide = Math.max(dims.width || 0, dims.height || 0);

  if (/image1\.png|image2\.png|image23\.png|image24\.png/i.test(media.name)) return "branding";
  if (ext === ".svg") return "vector-icons";
  if (slides >= 3 && maxSide <= 600) return "decorations";
  if (maxSide >= 1200 || media.size >= 200000) return "screenshots-charts";
  if (maxSide <= 400) return "icons";
  return "mixed-media";
}

function unique(list) {
  return [...new Set((list || []).filter(Boolean))];
}

function deriveMaterialTags(media) {
  const tags = [media.category];
  const ext = path.extname(media.name).toLowerCase();
  const dims = media.dimensions || { width: 0, height: 0 };
  const maxSide = Math.max(dims.width || 0, dims.height || 0);

  if (ext === ".svg") tags.push("vector");
  if ([".png", ".jpg", ".jpeg", ".webp"].includes(ext)) tags.push("bitmap");
  if ((media.slides || []).length >= 2) tags.push("reusable");
  if (maxSide >= 1200) tags.push("wide");
  if (maxSide <= 420) tags.push("small");
  if (media.category === "branding") tags.push("logo");
  if (media.category === "screenshots-charts") tags.push("screenshot");
  if (media.category === "vector-icons" || media.category === "icons") tags.push("icon");
  if (media.category === "mixed-media" || media.category === "decorations") tags.push("illustration");
  return unique(tags);
}

function deriveUsageTags(media, slideSummaries) {
  const tags = [];
  const relatedSlides = slideSummaries.filter((slide) => (media.slides || []).includes(slide.slide));
  const textCorpus = relatedSlides
    .flatMap((slide) => slide.topTexts || [])
    .map((item) => item.text || "")
    .join(" ");

  if (media.category === "branding") {
    tags.push("header-logo");
    if (/底部|页脚|footer/i.test(textCorpus)) tags.push("footer-brand");
  }
  if (media.category === "screenshots-charts") {
    tags.push("data-analysis");
    if (/系统|页面|流程|截图/.test(textCorpus)) tags.push("system-preview");
    if (/效果|结果|成效|试点/.test(textCorpus)) tags.push("evidence");
  }
  if (media.category === "vector-icons" || media.category === "icons") {
    tags.push("summary-card");
    tags.push("process-flow");
    if (/行动|计划|推进|安排/.test(textCorpus)) tags.push("action-plan");
  }
  if (/封面|标题|专题|汇报/.test(textCorpus)) tags.push("cover");
  if (/表|指标|分析|对比/.test(textCorpus)) tags.push("data-analysis");
  if (/流程|路径|模型|机制/.test(textCorpus)) tags.push("process-flow");
  return unique(tags);
}

function guessComponentFamilies(shapeCounts, topColors) {
  return {
    header: {
      titleText: "Top-left black title with thin dark-green underline and top-right brand logo",
      underlineColor: topColors.find((item) => item.name === "006544" || item.name === "086346")?.name || "006544",
      commonShapes: ["rect", "line"],
    },
    summaryBand: {
      description: "Full-width light-green summary strip with dark-green left accent bar",
      fills: topColors.filter((item) => ["E2F0D9", "FFFFFF", "006544", "086346"].includes(item.name)),
      commonShapes: ["rect"],
    },
    sectionLabels: [
      { preset: "homePlate", description: "Title tag with notch tail" },
      { preset: "parallelogram", description: "Tilted ribbon tag" },
      { preset: "chevron", description: "Step/flow ribbon" },
      { preset: "rightArrow", description: "Directional arrow banner" },
    ].filter((item) => shapeCounts.some((shape) => shape.name === `prst="${item.preset}"` || shape.name === item.preset)),
    callouts: [
      { preset: "wedgeRoundRectCallout", description: "Speech-bubble callout card" },
      { preset: "wedgeRectCallout", description: "Sharp speech callout" },
      { preset: "teardrop", description: "Accent droplet/marker badge" },
    ],
    badges: [
      { preset: "ellipse", description: "Solid circular badge" },
      { preset: "round2DiagRect", description: "Rounded diagonal badge" },
      { preset: "star5", description: "Highlight/star marker" },
    ],
    containers: [
      { preset: "rect", description: "Plain content box / table cell / divider" },
      { preset: "roundRect", description: "Soft card container" },
    ],
  };
}

function collectTextRuns(txBody) {
  const runs = [];
  toArray(txBody?.["a:p"]).forEach((para) => {
    toArray(para?.["a:r"]).forEach((run) => {
      const text = cleanText(collectText(run?.["a:t"]));
      if (!text) return;
      const rPr = run?.["a:rPr"] || {};
      runs.push({
        text,
        font: rPr["a:latin"]?.["@_typeface"] || rPr["a:ea"]?.["@_typeface"] || "",
        size: Number(rPr["@_sz"] || 0) / 100,
        color: getHex(rPr["a:solidFill"]) || "",
        bold: rPr["@_b"] === "1",
      });
    });
  });
  return runs;
}

function summarizeTable(tableNode) {
  const rows = toArray(tableNode?.["a:tr"]);
  const table = {
    rows: rows.length,
    cols: 0,
    fills: [],
    textColors: [],
  };
  rows.forEach((row) => {
    const cells = toArray(row?.["a:tc"]);
    table.cols = Math.max(table.cols, cells.length);
    cells.forEach((cell) => {
      const fill = getHex(cell?.["a:tcPr"]?.["a:solidFill"]);
      const runs = collectTextRuns(cell?.["a:txBody"]);
      if (fill) table.fills.push(fill);
      runs.forEach((run) => {
        if (run.color) table.textColors.push(run.color);
      });
    });
  });
  return table;
}

async function extractReferenceLibrary(args) {
  ensureDir(args.out);
  ensureDir(path.join(args.out, "media"));
  ensureDir(path.join(args.out, "previews"));
  ensureDir(path.join(args.out, "categorized"));

  const zip = await loadZip(args.ppt);
  const presentationXml = parser.parse(await readText(zip, "ppt/presentation.xml"));
  const themeXml = parser.parse(await readText(zip, "ppt/theme/theme1.xml"));
  const presentationRelsXml = parser.parse(await readText(zip, "ppt/_rels/presentation.xml.rels"));
  const slideSize = presentationXml["p:presentation"]["p:sldSz"];
  const presentationRels = Object.fromEntries(
    toArray(presentationRelsXml.Relationships.Relationship).map((rel) => [rel["@_Id"], rel["@_Target"]]),
  );

  const slideIdList = toArray(presentationXml["p:presentation"]["p:sldIdLst"]?.["p:sldId"]);
  const slideTargets = slideIdList.map((item) => normalizeTarget("ppt/presentation.xml", presentationRels[item["@_r:id"]]));
  const slideFiles = slideTargets.length
    ? slideTargets
    : Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort();

  const mediaEntries = Object.keys(zip.files)
    .filter((name) => name.startsWith("ppt/media/") && !zip.files[name].dir)
    .sort();

  const mediaCatalog = new Map();
  for (const mediaPath of mediaEntries) {
    const fileName = path.basename(mediaPath);
    const buffer = await readBuffer(zip, mediaPath);
    const outPath = path.join(args.out, "media", fileName);
    fs.writeFileSync(outPath, buffer);
    let dimensions = { width: 0, height: 0 };
    try {
      dimensions = imageSize(buffer);
    } catch {
      if (fileName.endsWith(".svg")) {
        dimensions = readSvgDims(buffer.toString("utf8"));
      }
    }
    mediaCatalog.set(fileName, {
      id: path.parse(fileName).name,
      name: fileName,
      source: mediaPath,
      ext: path.extname(fileName).toLowerCase(),
      size: buffer.length,
      dimensions,
      slides: [],
      uses: [],
    });
  }

  if (args.previewDir && fs.existsSync(args.previewDir)) {
    for (const entry of fs.readdirSync(args.previewDir)) {
      const src = path.join(args.previewDir, entry);
      if (!fs.statSync(src).isFile()) continue;
      fs.copyFileSync(src, path.join(args.out, "previews", entry));
    }
  } else {
    await buildRenderedPreviews(args.ppt, path.join(args.out, "previews"), [], {
      width: 1600,
      height: 900,
      timeoutMs: 180000,
    }).then((result) => result?.previews || []).catch(() => []);
  }

  const shapeCounts = new Map();
  const colorCounts = new Map();
  const fontCounts = new Map();
  const textStyleCounts = new Map();
  const tableStyleCounts = new Map();
  const slideSummaries = [];

  for (const slideFile of slideFiles) {
    const slideXmlText = await readText(zip, slideFile);
    const slideXml = parser.parse(slideXmlText);
    const slideName = path.basename(slideFile, ".xml");
    const relFile = path.posix.join(path.posix.dirname(slideFile), "_rels", `${path.basename(slideFile)}.rels`);
    const relText = await readText(zip, relFile);
    const relXml = relText ? parser.parse(relText) : { Relationships: {} };
    const relMap = Object.fromEntries(
      toArray(relXml.Relationships?.Relationship).map((rel) => [rel["@_Id"], normalizeTarget(slideFile, rel["@_Target"])]),
    );

    const spTree = slideXml["p:sld"]?.["p:cSld"]?.["p:spTree"] || {};
    const slideShapes = [];
    const slideTexts = [];
    const slideTables = [];
    const slideMedia = [];

    toArray(spTree["p:sp"]).forEach((shapeNode) => {
      const spPr = shapeNode["p:spPr"] || {};
      const prst = spPr["a:prstGeom"]?.["@_prst"] || "custom";
      const fill = getFillHex(spPr);
      const line = getLineHex(spPr);
      const text = cleanText(collectText(shapeNode["p:txBody"]));
      const runs = collectTextRuns(shapeNode["p:txBody"]);
      const xfrm = getTransform(spPr);
      slideShapes.push({ prst, fill, line, text, xfrm });
      bump(shapeCounts, prst);
      if (fill) bump(colorCounts, fill);
      if (line) bump(colorCounts, line);
      runs.forEach((run) => {
        if (run.font) bump(fontCounts, run.font);
        if (run.color) bump(colorCounts, run.color);
        bump(textStyleCounts, `${run.font || "unknown"}|${run.size || 0}|${run.color || "none"}|${run.bold ? "bold" : "regular"}`);
        slideTexts.push(run);
      });
    });

    toArray(spTree["p:pic"]).forEach((picNode) => {
      const rid = picNode["p:blipFill"]?.["a:blip"]?.["@_r:embed"];
      const target = relMap[rid] || "";
      const fileName = path.basename(target);
      const spPr = picNode["p:spPr"] || {};
      const xfrm = getTransform(spPr);
      slideMedia.push({ file: fileName, xfrm });
      const media = mediaCatalog.get(fileName);
      if (media) {
        media.slides.push(slideName);
        media.uses.push({ slide: slideName, xfrm });
      }
    });

    toArray(spTree["p:graphicFrame"]).forEach((frameNode) => {
      const graphicData = frameNode["a:graphic"]?.["a:graphicData"];
      const tableNode = graphicData?.["a:tbl"];
      if (tableNode) {
        const table = summarizeTable(tableNode);
        slideTables.push(table);
        const signature = `${table.rows}x${table.cols}|${table.fills.slice(0, 3).join(",")}`;
        bump(tableStyleCounts, signature);
        table.fills.forEach((fill) => bump(colorCounts, fill));
        table.textColors.forEach((fill) => bump(colorCounts, fill));
      }
    });

    slideSummaries.push({
      slide: slideName,
      source: slideFile,
      shapes: slideShapes.length,
      pictures: slideMedia.length,
      tables: slideTables.length,
      dominantShapePresets: countMapToSorted(
        slideShapes.reduce((map, item) => {
          map.set(item.prst, (map.get(item.prst) || 0) + 1);
          return map;
        }, new Map()),
        8,
      ),
      topTexts: slideTexts.slice(0, 10).map((item) => ({
        text: item.text.slice(0, 40),
        font: item.font,
        size: item.size,
        color: item.color,
      })),
      media: slideMedia,
    });
  }

  const mediaList = [...mediaCatalog.values()].map((item) => ({
    ...item,
    slides: [...new Set(item.slides)].sort(),
  }));

  mediaList.forEach((item) => {
    item.category = detectMediaCategory(item);
    item.tags = deriveMaterialTags(item);
    item.usageTags = deriveUsageTags(item, slideSummaries);
    const destDir = path.join(args.out, "categorized", item.category);
    ensureDir(destDir);
    fs.copyFileSync(path.join(args.out, "media", item.name), path.join(destDir, item.name));
  });

  const tagTaxonomy = {
    materialTags: unique(mediaList.flatMap((item) => item.tags || [])).sort(),
    usageTags: unique(mediaList.flatMap((item) => item.usageTags || [])).sort(),
  };

  const topColors = countMapToSorted(colorCounts, 20);
  const topShapes = countMapToSorted(shapeCounts, 20);
  const topFonts = countMapToSorted(fontCounts, 20);
  const topTextStyles = countMapToSorted(textStyleCounts, 20).map((item) => {
    const [font, size, color, weight] = item.name.split("|");
    return { font, size: Number(size), color, weight, count: item.count };
  });

  const themeFonts = themeXml["a:theme"]?.["a:themeElements"]?.["a:fontScheme"];
  const output = {
    sourcePptx: path.resolve(args.ppt),
    generatedAt: new Date().toISOString(),
    slideSize: {
      widthInches: Number(slideSize["@_cx"]) / EMU_PER_INCH,
      heightInches: Number(slideSize["@_cy"]) / EMU_PER_INCH,
    },
    themeFonts: {
      majorLatin: themeFonts?.["a:majorFont"]?.["a:latin"]?.["@_typeface"] || "",
      minorLatin: themeFonts?.["a:minorFont"]?.["a:latin"]?.["@_typeface"] || "",
      majorEastAsia: themeFonts?.["a:majorFont"]?.["a:ea"]?.["@_typeface"] || "",
      minorEastAsia: themeFonts?.["a:minorFont"]?.["a:ea"]?.["@_typeface"] || "",
    },
    counts: {
      slides: slideFiles.length,
      media: mediaList.length,
      vectorIcons: mediaList.filter((item) => item.category === "vector-icons").length,
      screenshotsOrCharts: mediaList.filter((item) => item.category === "screenshots-charts").length,
      icons: mediaList.filter((item) => item.category === "icons").length,
      branding: mediaList.filter((item) => item.category === "branding").length,
    },
    palette: {
      topColors,
    },
    shapes: {
      topPresets: topShapes,
    },
    typography: {
      topFonts,
      topTextStyles,
    },
    tables: {
      topTableSignatures: countMapToSorted(tableStyleCounts, 10),
    },
    media: mediaList,
    slideSummaries,
    tagTaxonomy,
    reusableFamilies: guessComponentFamilies(topShapes, topColors),
  };

  const buildAsset = (item) => ({
    id: item.id,
    name: item.name,
    category: item.category,
    path: path.join(args.out, "categorized", item.category, item.name),
    dimensions: item.dimensions,
    slides: item.slides,
    tags: item.tags || [],
    usageTags: item.usageTags || [],
  });

  const assetCollections = {
    iconAssets: mediaList
      .filter((item) => item.category === "vector-icons" || item.category === "icons")
      .map(buildAsset),
    brandingAssets: mediaList
      .filter((item) => item.category === "branding")
      .map(buildAsset),
    illustrationAssets: mediaList
      .filter((item) => item.category === "decorations" || item.category === "mixed-media")
      .map(buildAsset),
    screenshotAssets: mediaList
      .filter((item) => item.category === "screenshots-charts")
      .map(buildAsset),
    allAssets: mediaList.map(buildAsset),
  };

  const reusable = {
    sourcePptx: output.sourcePptx,
    tagTaxonomy,
    paletteTokens: {
      deepGreen: topColors.find((item) => item.name === "006544")?.name || "006544",
      darkGreen: topColors.find((item) => item.name === "086346")?.name || "086346",
      limeGreen: topColors.find((item) => item.name === "70AD47")?.name || "70AD47",
      accentOrange: topColors.find((item) => item.name === "ED7D31")?.name || "ED7D31",
      accentGold: topColors.find((item) => item.name === "FFC000")?.name || "FFC000",
      lightGreen: topColors.find((item) => item.name === "E2F0D9")?.name || "E2F0D9",
      white: "FFFFFF",
      black: "000000",
    },
    assetCollections,
    componentPresets: {
      header: {
        titleAlign: "left",
        titleColor: "000000",
        underlineShape: "rect",
        underlineColor: topColors.find((item) => item.name === "006544")?.name || "006544",
        logoCategory: "branding",
      },
      summaryBand: {
        bandShape: "rect",
        bandFill: topColors.find((item) => item.name === "E2F0D9")?.name || "E2F0D9",
        accentBarFill: topColors.find((item) => item.name === "006544")?.name || "006544",
      },
      sectionRibbons: [
        { id: "home-plate-green", shape: "homePlate", fill: "006544", textColor: "FFFFFF" },
        { id: "parallelogram-lime", shape: "parallelogram", fill: "B4CC27", textColor: "FFFFFF" },
        { id: "chevron-green", shape: "chevron", fill: "70AD47", textColor: "FFFFFF" },
        { id: "right-arrow-blue", shape: "rightArrow", fill: "2F5597", textColor: "FFFFFF" },
      ],
      badges: [
        { id: "ellipse-solid", shape: "ellipse", fill: "006544", textColor: "FFFFFF" },
        { id: "teardrop-gold", shape: "teardrop", fill: "FFC000", textColor: "FFFFFF" },
        { id: "diag-rect-lime", shape: "round2DiagRect", fill: "70AD47", textColor: "FFFFFF" },
        { id: "star-highlight", shape: "star5", fill: "ED7D31", textColor: "FFFFFF" },
      ],
      callouts: [
        { id: "soft-speech", shape: "wedgeRoundRectCallout", fill: "E2F0D9", line: "70AD47", textColor: "000000" },
        { id: "sharp-speech", shape: "wedgeRectCallout", fill: "FFFFFF", line: "006544", textColor: "000000" },
      ],
      cards: [
        { id: "plain-card", shape: "rect", fill: "FFFFFF", line: "70AD47" },
        { id: "soft-round-card", shape: "roundRect", fill: "FFFFFF", line: "B7D7A8" },
      ],
      dividers: [
        { id: "thin-green-line", shape: "rect", fill: "006544" },
        { id: "soft-green-strip", shape: "rect", fill: "E2F0D9" },
      ],
      tables: [
        {
          id: "dark-header-striped-body",
          headerFill: "006544",
          headerText: "FFFFFF",
          bodyFillA: "E2F0D9",
          bodyFillB: "FFFFFF",
          emphasisText: "ED7D31",
        },
      ],
      iconAssets: assetCollections.iconAssets,
    },
  };

  fs.writeFileSync(path.join(args.out, "catalog.json"), JSON.stringify(output, null, 2), "utf8");
  fs.writeFileSync(path.join(args.out, "reusable_materials.json"), JSON.stringify(reusable, null, 2), "utf8");

  const summaryLines = [
    "# 参考 PPT 素材库",
    "",
    `- 来源 PPT: ${path.resolve(args.ppt)}`,
    `- 生成时间: ${output.generatedAt}`,
    `- 幻灯片数量: ${output.counts.slides}`,
    `- 提取媒体总数: ${output.counts.media}`,
    `- 矢量图标数量: ${output.counts.vectorIcons}`,
    `- 常规图标数量: ${output.counts.icons}`,
    `- 截图/图表数量: ${output.counts.screenshotsOrCharts}`,
    `- 品牌素材数量: ${output.counts.branding}`,
    "",
    "## 颜色与形状统计",
    ...topColors.slice(0, 10).map((item) => `- 颜色 ${item.name}: ${item.count}`),
    ...topShapes.slice(0, 10).map((item) => `- 形状 ${item.name}: ${item.count}`),
    "",
    "## 可复用组件建议",
    "- 页眉: 左标题 + 顶部细绿线 + 右上品牌标识。",
    "- 摘要带: 浅绿横条 + 深绿左侧强调条。",
    "- 标题签: `homePlate` / `parallelogram` / `chevron` / `rightArrow` 交替使用。",
    "- 说明框: `rect` / `roundRect` / `wedgeRoundRectCallout` 组合。",
    "- 表格: 深绿表头 + 浅绿/白色斑马纹 + 橙红重点数字。",
    "- 图标: 优先使用 `categorized/vector-icons/` 与 `categorized/icons/`，并保留品牌、插画、截图等独立资产集合。",
    "",
    "## 素材集合",
    `- 图标资产: ${assetCollections.iconAssets.length}`,
    `- 品牌资产: ${assetCollections.brandingAssets.length}`,
    `- 插画/装饰资产: ${assetCollections.illustrationAssets.length}`,
    `- 截图/图表资产: ${assetCollections.screenshotAssets.length}`,
    `- 全量资产索引: ${assetCollections.allAssets.length}`,
    "",
    "## 输出目录",
    "- media/: 原始媒体完整提取。",
    "- categorized/: 按 branding / vector-icons / icons / screenshots-charts / decorations / mixed-media 分类后的素材。",
    "- previews/: 参考 PPT 导出的整页预览图。",
    "- catalog.json: 全量结构化归纳。",
    "- reusable_materials.json: 可复用组件、资产集合与色板预设。",
    "",
  ];
  fs.writeFileSync(path.join(args.out, "README.md"), `${summaryLines.join("\n")}\n`, "utf8");

  console.log(`Reference library extracted to: ${args.out}`);
  console.log(`Catalog: ${path.join(args.out, "catalog.json")}`);
  console.log(`Reusable materials: ${path.join(args.out, "reusable_materials.json")}`);

  return {
    outDir: args.out,
    catalogPath: path.join(args.out, "catalog.json"),
    reusablePath: path.join(args.out, "reusable_materials.json"),
    readmePath: path.join(args.out, "README.md"),
  };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  return extractReferenceLibrary(args);
}

module.exports = {
  parseArgs,
  ensureDir,
  toArray,
  decodeXml,
  cleanText,
  collectText,
  loadZip,
  readText,
  readBuffer,
  detectMediaCategory,
  extractReferenceLibrary,
  main,
};

if (require.main === module) {
  main().catch((error) => {
    console.error(error.stack || error.message || error);
    process.exitCode = 1;
  });
}
