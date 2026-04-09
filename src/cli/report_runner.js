const fs = require("node:fs");
const path = require("node:path");

const {
  decodeXml,
  cleanText,
  toArray,
  collectText,
  ensureDir,
  extractDocx,
  extractPdf,
  extractInputDocument,
  extractTemplate,
  buildDocumentSummary: buildRawDocumentSummary,
} = require("../services/documentParserService");
const { buildDocumentStructure } = require("../services/documentStructureService");
const { clampPageCount, detectSections, buildOutline, buildNotes } = require("../services/outlinePlannerService");
const { renderDeck } = require("../services/deckRendererService");
const { normalizeReusableLibrary } = require("../services/referenceLibraryService");
const { templateVariantFamily } = require("../services/layoutSelectionService");
const { analyzeReferenceImages } = require("../services/referenceStyleService");
const { readJson, writeJson } = require("../utils/fileUtils");
const { DEFAULT_LAYOUT_LIBRARY, DEFAULT_REFERENCE_LIBRARY } = require("../utils/pathConfig");

const DEFAULT_LAYOUTS_BY_TYPE = {
  cover: "cover_formal_v1",
  summary_cards: "summary_grid_v1",
  table_analysis: "table_split_v1",
  process_flow: "process_three_lane_v1",
  bullet_columns: "bullet_dual_v1",
  image_story: "image_split_v1",
  action_plan: "action_timeline_v1",
  key_takeaways: "takeaway_cards_v1",
};

function parseArgs(argv) {
  const args = {
    word: process.env.WORD_PATH || process.env.DOC_PATH || "",
    doc: process.env.DOC_PATH || process.env.WORD_PATH || "",
    template: process.env.TEMPLATE_PATH || "",
    refImage: process.env.REF_IMAGE_PATH || "",
    referenceLibrary: process.env.REFERENCE_LIBRARY_PATH || DEFAULT_REFERENCE_LIBRARY,
    layoutLibrary: process.env.LAYOUT_LIBRARY_PATH || DEFAULT_LAYOUT_LIBRARY,
    layoutSet: process.env.LAYOUT_SET || "",
    pages: process.env.PAGES ? clampPageCount(process.env.PAGES) : 0,
    out: path.resolve(process.cwd(), "output"),
    title: process.env.REPORT_TITLE || "",
    department: process.env.DEPARTMENT || "",
    presenter: process.env.PRESENTER || "",
    date: process.env.REPORT_DATE || "",
  };

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) continue;
    const key = token.slice(2);
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) continue;
    index += 1;

    if (key === "pages") {
      args.pages = value === "auto" ? 0 : clampPageCount(value);
      continue;
    }
    if (key === "word" || key === "doc") {
      args.word = value;
      args.doc = value;
      continue;
    }
    if (key === "out") {
      args.out = path.resolve(process.cwd(), value);
      continue;
    }
    if (key === "ref-image") {
      args.refImage = value;
      continue;
    }
    if (key === "reference-library") {
      args.referenceLibrary = path.resolve(process.cwd(), value);
      continue;
    }
    if (key === "layout-library") {
      args.layoutLibrary = path.resolve(process.cwd(), value);
      continue;
    }
    args[key] = value;
  }

  if (!args.word && args.doc) {
    args.word = args.doc;
  }
  if (!args.word || !args.template) {
    throw new Error("Missing required arguments: --word/--doc and --template");
  }
  return args;
}

function createEmptyReferenceLibrary(libraryPath = "") {
  return {
    path: libraryPath || null,
    sourcePptx: null,
    paletteTokens: {},
    tagTaxonomy: {
      materialTags: [],
      usageTags: [],
    },
    assetCollections: {
      iconAssets: [],
      brandingAssets: [],
      illustrationAssets: [],
      screenshotAssets: [],
      allAssets: [],
    },
    componentPresets: {
      header: {},
      summaryBand: {},
      sectionRibbons: [],
      badges: [],
      callouts: [],
      cards: [],
      dividers: [],
      tables: [],
      iconAssets: [],
    },
  };
}

function stableHash(value = "") {
  let hash = 0;
  const text = String(value || "");
  for (let index = 0; index < text.length; index += 1) {
    hash = (hash * 31 + text.charCodeAt(index)) % 2147483647;
  }
  return Math.abs(hash);
}

function normalizeReferenceLibrary(json, libraryPath) {
  const fallback = createEmptyReferenceLibrary(libraryPath);
  const normalized = normalizeReusableLibrary(json);
  return {
    ...fallback,
    ...normalized,
    path: libraryPath,
    tagTaxonomy: normalized.tagTaxonomy || fallback.tagTaxonomy,
    assetCollections: normalized.assetCollections || fallback.assetCollections,
    componentPresets: {
      ...fallback.componentPresets,
      ...(normalized.componentPresets || {}),
      iconAssets: normalized.assetCollections?.iconAssets || normalized.componentPresets?.iconAssets || [],
    },
  };
}

function loadReferenceLibrary(libraryPath) {
  if (!libraryPath || !fs.existsSync(libraryPath)) {
    return createEmptyReferenceLibrary(libraryPath);
  }
  return normalizeReferenceLibrary(readJson(libraryPath, {}), libraryPath);
}

function createEmptyLayoutLibrary(layoutPath = "") {
  return {
    path: layoutPath || null,
    version: "2.0.0",
    defaultSet: "bank_finance_default",
    setName: "bank_finance_default",
    setMetadata: {},
    defaultsByType: { ...DEFAULT_LAYOUTS_BY_TYPE },
    sets: {
      bank_finance_default: { ...DEFAULT_LAYOUTS_BY_TYPE },
    },
    set: { ...DEFAULT_LAYOUTS_BY_TYPE },
    templates: {},
  };
}

function loadLayoutLibrary(layoutPath, layoutSet = "") {
  const fallback = createEmptyLayoutLibrary(layoutPath);
  if (!layoutPath || !fs.existsSync(layoutPath)) {
    return fallback;
  }

  const json = readJson(layoutPath, {});
  const defaultsByType = {
    ...DEFAULT_LAYOUTS_BY_TYPE,
    ...(json.defaultsByType || {}),
  };
  const setName =
    layoutSet ||
    json.defaultSet ||
    Object.keys(json.sets || {})[0] ||
    fallback.defaultSet;

  return {
    path: layoutPath,
    version: json.version || fallback.version,
    defaultSet: json.defaultSet || fallback.defaultSet,
    setName,
    setMetadata: json.setMetadata || {},
    defaultsByType,
    sets: json.sets || fallback.sets,
    set: json.sets?.[setName] || {},
    templates: json.templates || {},
  };
}

function getReferenceAssets(materials) {
  const collections = materials?.assetCollections || {};
  return [
    ...(collections.iconAssets || []),
    ...(collections.brandingAssets || []),
    ...(collections.illustrationAssets || []),
    ...(collections.screenshotAssets || []),
    ...(collections.allAssets || []),
    ...(materials?.componentPresets?.iconAssets || []),
  ].filter(Boolean);
}

function findAssetByUsage(materials, usageTag) {
  return getReferenceAssets(materials).find((item) => (item.usageTags || []).includes(usageTag)) || null;
}

function findAssetByCategory(materials, category) {
  return getReferenceAssets(materials).find((item) => item.category === category) || null;
}

function resolveHex(value, fallback) {
  const normalized = String(value || "").trim().replace(/^#/, "");
  return /^[0-9a-fA-F]{6}$/.test(normalized) ? normalized.toUpperCase() : fallback;
}

function buildStyle(template, refImage, materials, layoutLibrary, options = {}) {
  const palette = materials?.paletteTokens || {};
  const headerLogo = findAssetByUsage(materials, "header-logo") || findAssetByCategory(materials, "branding");
  const footerBrand = findAssetByUsage(materials, "footer-strip") || findAssetByUsage(materials, "footer-brand");

  return {
    referenceImage: refImage || null,
    referenceLibrary: materials?.path || null,
    referenceStyleProfile: options.referenceStyleProfile || null,
    layoutLibraryPath: layoutLibrary?.path || null,
    materials,
    layoutLibrary,
    template,
    assetUsage: {},
    page: {
      widthInches: template.widthInches,
      heightInches: template.heightInches,
    },
    department: options.department || "",
    presenter: options.presenter || "",
    reportDate: options.date || "",
    assets: {
      brand: headerLogo?.path || "",
      footerBar: footerBrand?.path || "",
    },
    fonts: {
      heading: "Microsoft YaHei",
      body: "Microsoft YaHei",
      fallbackHeading: template.headFontFace || "Calibri",
      fallbackBody: template.bodyFontFace || "Calibri",
    },
    palette: {
      deepGreen: resolveHex(palette.deepGreen, "006544"),
      darkGreen: resolveHex(palette.darkGreen, "0F6D53"),
      oliveGreen: resolveHex(palette.limeGreen, "70AD47"),
      lightGreen: resolveHex(palette.lightGreen, "E2F0D9"),
      paleGreen: "F6FBF3",
      borderGreen: "BCD7B0",
      softLine: "D7E7D4",
      accentRed: "DB2C1F",
      accentOrange: resolveHex(palette.accentOrange, "D96A1D"),
      accentLime: "A8C94A",
      accentGold: resolveHex(palette.accentGold, "D9A319"),
      accentBlue: "2F5597",
      textPrimary: resolveHex(palette.black, "111111"),
      textMuted: "4C5F52",
      titleBlack: resolveHex(palette.black, "111111"),
      white: "FFFFFF",
      subtleGray: "F3F6F2",
    },
  };
}

function resolveTemplateId(slide, layoutLibrary, slideSelections = {}, context = {}) {
  const templatesByType = Object.entries(layoutLibrary?.templates || {})
    .filter(([, template]) => template.pageType === slide.type)
    .map(([id, template]) => ({ id, ...template }));
  const contentMode = String(slide.contentMode || slide.density || "balanced").toLowerCase();
  const slideTextLength = [
    slide.headline || "",
    ...(slide.cards || []).map((item) => item.body || item.detail || ""),
    ...(slide.insights || []).map((item) => item.body || ""),
    ...(slide.columns || []).flatMap((item) => item.bullets || []),
    ...(slide.stages || []).map((item) => item.detail || ""),
    ...(slide.steps || []).map((item) => item.detail || ""),
    ...(slide.takeaways || []).map((item) => item.body || ""),
  ]
    .join(" ")
    .length;
  const density = slide.density || "medium";
  const rowCount = slide.table?.rowCount || slide.table?.rows?.length || 0;
  const colCount = slide.table?.colCount || slide.table?.header?.length || 0;
  const cardCount = slide.cards?.length || 0;
  const insightCount = slide.insights?.length || 0;
  const stageCount = slide.stages?.length || 0;
  const stepCount = slide.steps?.length || 0;
  const takeawaysCount = slide.takeaways?.length || 0;
  const hasImage = Boolean(slide.image?.path);
  const imageAspect = Number(slide.image?.aspectRatio || slide.screenshots?.[0]?.aspectRatio || 0);
  const hasScreenshots = Boolean(slide.screenshots?.length);
  const previousTemplate = context.previousTemplateByType?.[slide.type] || "";
  const recentFamilies = Array.isArray(context.recentFamilies) ? context.recentFamilies.slice(-4) : [];
  const recentTemplates = Array.isArray(context.recentTemplates) ? context.recentTemplates.slice(-4) : [];

  function findVariant(variant) {
    return templatesByType.find((template) => template.variant === variant)?.id || "";
  }

  function pickTemplate(variants, fallbackVariants = []) {
    const orderedIds = [
      ...variants.map((variant) => findVariant(variant)).filter(Boolean),
      ...fallbackVariants.map((variant) => findVariant(variant)).filter(Boolean),
    ];
    const uniqueIds = [...new Set(orderedIds)];
    if (!uniqueIds.length) return "";
    const scored = uniqueIds.map((id, index) => {
      const template = layoutLibrary?.templates?.[id] || {};
      const family = template.family || templateVariantFamily(template.variant);
      let score = index * 100;
      if (previousTemplate && id === previousTemplate) score += 1000;
      if (recentTemplates.includes(id)) score += 420;
      if (recentTemplates.slice(-2).includes(id)) score += 220;
      if (recentFamilies.includes(family)) score += 520;
      if (recentFamilies.slice(-2).includes(family)) score += 260;
      score += (stableHash(`${slide.page}|${slide.type}|${id}`) % 97) / 100;
      return { id, score };
    });
    scored.sort((left, right) => left.score - right.score);
    return scored[0]?.id || uniqueIds[0] || "";
  }

  function recommendTemplateId() {
    switch (slide.type) {
      case "summary_cards":
        if (contentMode === "sparse" || (density === "low" && cardCount >= 4 && slideTextLength <= 220)) {
          return pickTemplate(["bands", "spread", "grid", "dense_grid", "mosaic"]);
        }
        if (contentMode === "dense" || density === "high" || cardCount >= 5 || slideTextLength >= 260) {
          return pickTemplate(["mosaic", "dense_grid", "spread", "grid", "bands"]);
        }
        if (contentMode === "visual") {
          return pickTemplate(["spread", "grid", "bands", "mosaic", "dense_grid"]);
        }
        if ((slide.metrics?.length || 0) <= 1 && slideTextLength <= 180) {
          return pickTemplate(["bands", "grid", "spread", "dense_grid", "mosaic"]);
        }
        return pickTemplate(["grid", "spread", "bands", "mosaic", "dense_grid"]);
      case "table_analysis":
        if (contentMode === "visual" || hasImage || hasScreenshots) {
          if (imageAspect >= 1.8) {
            return pickTemplate(["picture", "visual", "split", "dashboard", "sidecallout", "highlight", "dense"]);
          }
          if (imageAspect > 0 && imageAspect < 1.0) {
            return pickTemplate(["visual", "split", "picture", "dashboard", "sidecallout", "highlight", "dense"]);
          }
          return pickTemplate(["visual", "picture", "sidecallout", "dashboard", "highlight", "split", "dense"]);
        }
        if (contentMode === "dense" || rowCount >= 7 || colCount >= 5 || (density === "high" && rowCount >= 5)) {
          return pickTemplate(["dense", "dashboard", "highlight", "sidecallout", "split", "visual"]);
        }
        if (contentMode === "sparse" || rowCount <= 3 || slideTextLength <= 160) {
          return pickTemplate(["highlight", "split", "sidecallout", "visual", "dashboard", "dense"]);
        }
        if (insightCount >= 3) {
          return pickTemplate(["highlight", "sidecallout", "split", "dashboard", "visual", "dense"]);
        }
        if (rowCount >= 5 && colCount >= 4 && slideTextLength >= 220) {
          return pickTemplate(["dashboard", "split", "highlight", "sidecallout", "visual", "dense"]);
        }
        if ((slide.metrics?.length || 0) >= 2 && rowCount <= 5) {
          return pickTemplate(["split", "highlight", "sidecallout", "dashboard", "visual", "dense"]);
        }
        return pickTemplate(["split", "highlight", "visual", "dashboard", "sidecallout", "dense"]);
      case "process_flow":
        if (contentMode === "dense" || stageCount >= 5 || density === "high") return pickTemplate(["ladder", "bridge", "cards", "three_lane"]);
        if (contentMode === "visual" || (slide.notes?.length || 0) >= 3) return pickTemplate(["bridge", "cards", "ladder", "three_lane"]);
        if (stageCount >= 3 && stageCount <= 5 && (slide.notes?.length || 0) >= 2) return pickTemplate(["bridge", "cards", "ladder", "three_lane"]);
        if (stageCount <= 4 && (slide.notes?.length || 0) <= 2) return pickTemplate(["cards", "bridge", "three_lane", "ladder"]);
        return pickTemplate(["three_lane", "bridge", "cards", "ladder"]);
      case "bullet_columns":
        if (contentMode === "dense" || ((slide.columns?.length || 0) >= 3 && (slide.cards?.length || 0) >= 2 && density === "high")) return pickTemplate(["masonry", "triple", "staggered", "dual"]);
        if (contentMode === "sparse" || cardCount >= 2 && slideTextLength <= 260) return pickTemplate(["staggered", "dual", "triple", "masonry"]);
        if ((slide.columns?.length || 0) >= 3 || density === "high") return pickTemplate(["triple", "masonry", "staggered", "dual"]);
        return pickTemplate(["dual", "staggered", "triple", "masonry"]);
      case "image_story":
        if (contentMode === "visual" || (hasImage && (slide.callouts?.length || 0) >= 2 && slideTextLength >= 260)) return pickTemplate(["storyboard", "gallery", "focus", "split"]);
        if (hasImage && (slide.callouts?.length || 0) >= 3 && slideTextLength <= 240) return pickTemplate(["focus", "gallery", "storyboard", "split"]);
        if ((slide.callouts?.length || 0) >= 3 && density === "high") return pickTemplate(["gallery", "storyboard", "focus", "split"]);
        return pickTemplate(["split", "gallery", "focus", "storyboard"]);
      case "action_plan":
        if (contentMode === "visual" || hasScreenshots && stepCount >= 2 && slideTextLength <= 260) return pickTemplate(["dashboard", "matrix", "timeline", "stacked"]);
        if (hasScreenshots && stepCount >= 3) return pickTemplate(["matrix", "dashboard", "timeline", "stacked"]);
        if (contentMode === "dense" || density === "high" || slideTextLength >= 240) return pickTemplate(["stacked", "dashboard", "matrix", "timeline"]);
        return pickTemplate(["timeline", "dashboard", "matrix", "stacked"]);
      case "key_takeaways":
        if (contentMode === "dense" || takeawaysCount >= 4 || slideTextLength >= 220) return pickTemplate(["wall", "closing", "cards"]);
        if (contentMode === "sparse" || takeawaysCount <= 2 && (slide.footer || "").length >= 48) return pickTemplate(["closing", "cards", "wall"]);
        return pickTemplate(["cards", "closing", "wall"]);
      case "cover":
      default:
        return "";
    }
  }

  return (
    slideSelections[String(slide.page)] ||
    slide.templateId ||
    recommendTemplateId() ||
    layoutLibrary?.set?.[slide.type] ||
    layoutLibrary?.defaultsByType?.[slide.type] ||
    DEFAULT_LAYOUTS_BY_TYPE[slide.type] ||
    ""
  );
}

function applyTemplatesToOutline(outline, layoutLibrary, slideSelections = {}) {
  const previousTemplateByType = {};
  const recentFamilies = [];
  const recentTemplates = [];
  return {
    ...outline,
    slides: (outline.slides || []).map((slide) => {
      const templateId = resolveTemplateId(slide, layoutLibrary, slideSelections, {
        previousTemplateByType,
        recentFamilies,
        recentTemplates,
      });
      previousTemplateByType[slide.type] = templateId;
      const resolvedTemplate = layoutLibrary?.templates?.[templateId] || {};
      const family = resolvedTemplate.family || templateVariantFamily(resolvedTemplate.variant);
      recentTemplates.push(templateId);
      recentFamilies.push(family);
      if (recentTemplates.length > 4) recentTemplates.shift();
      if (recentFamilies.length > 4) recentFamilies.shift();
      return {
        ...slide,
        templateId,
      };
    }),
  };
}

function buildDocumentSummary(doc, inputPath = "") {
  const summary = buildRawDocumentSummary(doc, inputPath);
  const structure = buildDocumentStructure(doc);
  return {
    ...summary,
    counts: {
      ...(summary.counts || {}),
      sections: structure.topLevel.length,
    },
    sectionTitles: structure.topLevel.map((item) => item.heading || item.title).slice(0, 60),
    sectionTree: structure.topLevel.map((item) => ({
      id: item.id,
      level: item.level,
      heading: item.heading,
      title: item.title,
      paragraphCount: item.meta?.paragraphCount || 0,
      charCount: item.meta?.charCount || 0,
      subsectionCount: item.meta?.subsectionCount || 0,
      tables: item.tables?.length || 0,
      images: item.images?.length || 0,
      children: (item.children || []).map((child) => ({
        id: child.id,
        level: child.level,
        heading: child.heading,
        title: child.title,
      })),
    })),
    frontMatter: structure.frontMatter || { headings: [], paragraphs: [] },
    markdown: structure.markdown,
  };
}

function buildLayoutTemplateManifest(style, outline) {
  const layoutLibrary = style.layoutLibrary || createEmptyLayoutLibrary(style.layoutLibraryPath);
  return {
    path: layoutLibrary.path,
    version: layoutLibrary.version,
    activeSet: layoutLibrary.setName,
    defaultSet: layoutLibrary.defaultSet,
    defaultsByType: layoutLibrary.defaultsByType,
    slides: (outline?.slides || []).map((slide) => ({
      page: slide.page,
      type: slide.type,
      title: slide.title,
      templateId: slide.templateId || resolveTemplateId(slide, layoutLibrary),
      displayName: layoutLibrary.templates?.[slide.templateId || resolveTemplateId(slide, layoutLibrary)]?.displayName || "",
    })),
  };
}

function buildMaterialRecipeManifest(style) {
  const materials = style.materials || createEmptyReferenceLibrary(style.referenceLibrary);
  return {
    sourcePptx: materials.sourcePptx || null,
    referenceLibrary: style.referenceLibrary || null,
    paletteTokens: materials.paletteTokens || {},
    tagTaxonomy: materials.tagTaxonomy || { materialTags: [], usageTags: [] },
    assetCounts: {
      icons: materials.assetCollections?.iconAssets?.length || 0,
      branding: materials.assetCollections?.brandingAssets?.length || 0,
      illustrations: materials.assetCollections?.illustrationAssets?.length || 0,
      screenshots: materials.assetCollections?.screenshotAssets?.length || 0,
      allAssets: materials.assetCollections?.allAssets?.length || 0,
    },
    componentCounts: {
      ribbons: materials.componentPresets?.sectionRibbons?.length || 0,
      badges: materials.componentPresets?.badges?.length || 0,
      callouts: materials.componentPresets?.callouts?.length || 0,
      cards: materials.componentPresets?.cards?.length || 0,
      tables: materials.componentPresets?.tables?.length || 0,
    },
  };
}

async function generateReport(args) {
  ensureDir(args.out);
  const assetDir = path.join(args.out, "assets");
  ensureDir(assetDir);

  const doc = await extractInputDocument(args.word, assetDir);
  const template = await extractTemplate(args.template);
  const materials = loadReferenceLibrary(args.referenceLibrary);
  const layoutLibrary = loadLayoutLibrary(args.layoutLibrary, args.layoutSet);
  const referenceStyleProfile =
    args.referenceStyleProfile ||
    (args.refImage
      ? analyzeReferenceImages(
          String(args.refImage)
            .split(",")
            .map((item) => item.trim())
            .filter(Boolean),
        )
      : null);
  const style = buildStyle(template, args.refImage, materials, layoutLibrary, {
    ...args,
    referenceStyleProfile,
  });
  const outline = applyTemplatesToOutline(buildOutline(doc, args), layoutLibrary);
  const summary = buildDocumentSummary(doc, args.word);
  const notes = buildNotes(outline);
  const materialManifest = buildMaterialRecipeManifest(style);
  const layoutManifest = buildLayoutTemplateManifest(style, outline);

  const outlinePath = path.join(args.out, "outline.json");
  const stylePath = path.join(args.out, "style.json");
  const summaryPath = path.join(args.out, "document_summary.json");
  const structurePath = path.join(args.out, "document_structure.json");
  const structureMarkdownPath = path.join(args.out, "document_structure.md");
  const notesPath = path.join(args.out, "page_notes.md");
  const materialPath = path.join(args.out, "material_recipes.json");
  const layoutPath = path.join(args.out, "layout_templates_resolved.json");

  writeJson(outlinePath, outline);
  writeJson(stylePath, style);
  writeJson(summaryPath, summary);
  writeJson(structurePath, outline.documentStructure || buildDocumentStructure(doc));
  fs.writeFileSync(structureMarkdownPath, summary.markdown || "", "utf8");
  writeJson(materialPath, materialManifest);
  writeJson(layoutPath, layoutManifest);
  fs.writeFileSync(notesPath, notes, "utf8");

  const deckName = `generated_report_${outline.meta.pages}slides.pptx`;
  const deckPath = path.join(args.out, deckName);
  await renderDeck(outline, style, deckPath);

  return {
    outline,
    style,
    documentSummary: summary,
    recipeManifest: materialManifest,
    layoutManifest,
    deckPath,
    outlinePath,
    stylePath,
    documentSummaryPath: summaryPath,
    documentStructurePath: structurePath,
    documentStructureMarkdownPath: structureMarkdownPath,
    recipePath: materialPath,
    layoutPath,
    notesPath,
  };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const result = await generateReport(args);
  console.log(`Generated deck: ${result.deckPath}`);
  console.log(`Generated outline: ${result.outlinePath}`);
  console.log(`Generated style: ${result.stylePath}`);
  console.log(`Generated summary: ${result.documentSummaryPath}`);
  return result;
}

module.exports = {
  parseArgs,
  ensureDir,
  decodeXml,
  cleanText,
  toArray,
  collectText,
  extractDocx,
  extractPdf,
  extractInputDocument,
  extractTemplate,
  detectSections,
  buildDocumentStructure,
  loadReferenceLibrary,
  loadLayoutLibrary,
  buildStyle,
  buildOutline,
  buildDocumentSummary,
  buildNotes,
  buildLayoutTemplateManifest,
  buildMaterialRecipeManifest,
  applyTemplatesToOutline,
  renderDeck,
  writeJson,
  generateReport,
  main,
};

if (require.main === module) {
  main().catch((error) => {
    console.error(error.stack || error.message || error);
    process.exitCode = 1;
  });
}
