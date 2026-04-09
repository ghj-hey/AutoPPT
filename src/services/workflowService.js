const fs = require("node:fs");
const path = require("node:path");
const reportGenerator = require("../../report_runner");
const {
  createSession,
  prepareFinalDeliveryWorkspace,
  saveSessionMeta,
  updateSessionMeta,
  loadSessionMeta,
  destroySessionWorkspace,
  stageUploadedFile,
  buildStagedFileName,
} = require("./workflowSessionService");
const {
  createReferenceLibraryFromPpt,
  mergeReferenceIntoMaster,
  composeReferenceLibraries,
  resolveLibraryById,
  getDefaultLibrary,
  getMasterLibrary,
  summarizeLibrary,
} = require("./referenceLibraryService");
const { analyzeReferenceImages } = require("./referenceStyleService");
const {
  analyzeDocumentSemantics,
  analyzeSectionSemantics,
  analyzeReferenceStyle,
  reviewRenderedDeck,
} = require("./semanticModelService.stable");
const { detectPreviewSupport, buildRenderedPreviews, exportPresentationPngs } = require("./powerPointPreviewService");
const { DEFAULT_LAYOUT_LIBRARY } = require("../utils/pathConfig");
const { ensureDir, readJson, writeJson } = require("../utils/fileUtils");
const {
  buildLayoutOptions: buildSmartLayoutOptions,
  templateVariantFamily,
} = require("./layoutSelectionService");
const {
  readableTypeLabel,
  buildReadableWorkflowStagesCn,
  buildReadableDownloadManifestCn,
  buildReadableFinalDownloadManifestCn,
  buildArchivedDeliveryManifestCn,
} = require("./workflowTextHelpers");

function unique(list) {
  return [...new Set((list || []).filter(Boolean))];
}

function toProgressPercent(value = 0) {
  const numeric = Number(value) || 0;
  return Math.max(0, Math.min(100, Math.round(numeric <= 1 ? numeric * 100 : numeric)));
}

function normalizePreviewState(previewResult = {}, stage = "draft") {
  const previewState = previewResult.previewState || previewResult.availability || {};
  const previews = Array.isArray(previewResult.previews) ? previewResult.previews : [];
  const status = String(previewState.status || "").toLowerCase();
  const supported = previewState.supported ?? previewState.available ?? status === "available";
  return {
    stage,
    supported: Boolean(supported),
    status: status || (previews.length ? "available" : supported ? "failed" : "unsupported"),
    reason: String(previewState.reason || "").trim(),
    platform: String(previewState.platform || process.platform),
    previewCount: Number(previewState.previewCount || previews.length || 0),
  };
}

function previewStatusMessage(stageName, previewState = {}) {
  const stageLabel = String(stageName || "预览").trim();
  const count = Number(previewState.previewCount || 0);
  const reason = String(previewState.reason || "").trim();
  switch (String(previewState.status || "").toLowerCase()) {
    case "available":
      return `${stageLabel}已完成，共 ${count} 张真实渲染图。`;
    case "unsupported":
      return `${stageLabel}已生成，但当前环境不支持真实预览，已跳过预览导出。`;
    case "failed":
      return `${stageLabel}已生成，但预览导出失败，已继续后续流程。${reason ? `原因：${reason}` : ""}`;
    default:
      return `${stageLabel}已生成。`;
  }
}

function normalizeSemanticApiKey(value = "") {
  return String(value || "").trim().replace(/^Bearer\s+/i, "");
}

function resolveSemanticApiKey(options = {}) {
  return normalizeSemanticApiKey(
    options.semanticApiKey ||
      process.env.SEMANTIC_MODEL_API_KEY ||
      process.env.MINIMAX_API_KEY ||
      process.env.OPENAI_API_KEY ||
      "",
  );
}

function resolveSemanticModelName(options = {}) {
  const explicitProvider = String(options.semanticProvider || "").toLowerCase();
  const provider = explicitProvider || String(process.env.SEMANTIC_MODEL_PROVIDER || "").toLowerCase();
  const hasExplicitConfig = explicitProvider || options.semanticBaseUrl || options.semanticApiKey;
  const explicit = String(options.semanticModel || "").trim();
  if (explicit) return explicit;
  if (!hasExplicitConfig && !provider) return "";
  if (provider === "minimax") return "MiniMax-M2.7";
  if (provider === "local" || options.semanticBaseUrl) {
    return String(process.env.SEMANTIC_MODEL_NAME || "deepseek-r1:14b").trim();
  }
  if (provider === "openai") {
    return String(process.env.OPENAI_MODEL_NAME || process.env.SEMANTIC_MODEL_NAME || "gpt-5.4").trim();
  }
  return String(process.env.SEMANTIC_MODEL_NAME || "").trim();
}

function createProgressWriter(sessionDir) {
  return (patch = {}) => updateSessionMeta(sessionDir, patch);
}

function sampleReferenceImages(imagePaths = [], limit = 12) {
  const files = unique(imagePaths);
  if (files.length <= limit) return files;

  const picks = [];
  const addIndex = (index) => {
    if (index >= 0 && index < files.length) picks.push(files[index]);
  };

  [0, 1, 2, Math.floor(files.length / 2) - 1, Math.floor(files.length / 2), Math.floor(files.length / 2) + 1, files.length - 3, files.length - 2, files.length - 1].forEach(addIndex);

  const step = Math.max(1, Math.floor(files.length / Math.max(3, limit)));
  for (let index = 0; index < files.length && picks.length < limit; index += step) {
    picks.push(files[index]);
  }

  return unique(picks).slice(0, limit);
}

function stageReferenceFiles(sessionDir, files = [], kind = "file", fallbackExt = "") {
  return (files || [])
    .filter((file) => {
      const name = String(file?.originalname || file?.path || "");
      return name && !path.basename(name).startsWith("~$");
    })
    .map((file, index) =>
      stageUploadedFile(
        sessionDir,
        file,
        buildStagedFileName(kind, index + 1, file.originalname || file.path || "", fallbackExt),
      ),
    );
}

function resolveSemanticProvider(options = {}) {
  const explicitProvider = String(options.semanticProvider || "").toLowerCase();
  if (explicitProvider) return explicitProvider;
  if (options.semanticBaseUrl) return "local";
  if (normalizeSemanticApiKey(options.semanticApiKey)) return "minimax";
  return String(process.env.SEMANTIC_MODEL_PROVIDER || "").toLowerCase() || "";
}

function resolveSemanticAvailability(options = {}) {
  const provider = resolveSemanticProvider(options);
  if (!provider) return false;
  if (provider === "local") return true;
  if (provider === "minimax") return Boolean(resolveSemanticApiKey(options));
  return Boolean(resolveSemanticApiKey(options));
}

function pickPrimaryUpload(files = []) {
  return (files || []).find((file) => {
    const name = String(file?.originalname || file?.path || "");
    return name && !path.basename(name).startsWith("~$");
  }) || null;
}

async function exportReferencePptImages(sessionDir, pptPaths = []) {
  const outputRoot = path.join(sessionDir, "intermediate", "reference_ppt_images");
  ensureDir(outputRoot);
  const exported = [];

  for (let index = 0; index < (pptPaths || []).length; index += 1) {
    const pptPath = pptPaths[index];
    if (!pptPath) continue;
    const baseName = path.basename(pptPath, path.extname(pptPath));
    const targetDir = path.join(outputRoot, `${String(index + 1).padStart(2, "0")}_${baseName}`);
    try {
      const images = await exportPresentationPngs(pptPath, targetDir, {
        width: 1600,
        height: 900,
        timeoutMs: 120000,
      });
      exported.push(...images);
    } catch {
      // If one reference PPT fails to export, keep the rest of the workflow available.
    }
  }

  return exported;
}

async function buildSectionSemanticPlans(sections = [], modelOptions = {}) {
  const candidates = (sections || []).slice(0, 3);
  const settled = await Promise.allSettled(
    candidates.map((section) =>
      analyzeSectionSemantics(section, {
        ...modelOptions,
        slideCap: 2,
      }),
    ),
  );

  return settled
    .map((result, index) => {
      if (result.status !== "fulfilled" || !result.value) return null;
      const section = candidates[index];
      return {
        sectionId: section.id,
        heading: section.heading,
        ...result.value,
      };
    })
    .filter(Boolean);
}

function buildLayoutOptions(layoutLibraryPath, outline, activeSet = "", currentSelection = {}, referenceStyleProfile = null) {
  const layoutLibrary = reportGenerator.loadLayoutLibrary(layoutLibraryPath, activeSet);
  const layoutOptions = buildSmartLayoutOptions(layoutLibrary, outline, currentSelection, referenceStyleProfile);
  return {
    ...layoutOptions,
    layoutLibraryPath,
    activeSet: layoutLibrary.setName,
    slideLayouts: (layoutOptions.slideLayouts || []).map((slide) => ({
      ...slide,
      typeLabel: readableTypeLabel(slide.type),
    })),
  };
}

function normalizeIssueCategories(semanticReview = {}) {
  return new Set((semanticReview.issues || []).map((item) => String(item?.category || "").toLowerCase()).filter(Boolean));
}

function tableModeVariantOrder(tableMode = "", pageParity = 0, contentMode = "") {
  const mode = String(tableMode || "").toLowerCase();
  const parity = Number(pageParity) % 2;
  const balance = String(contentMode || "").toLowerCase();

  if (["dashboard", "dense", "matrix"].includes(mode)) {
    return unique([
      ...(parity === 0 ? ["dashboard", "dense", "matrix"] : ["dense", "matrix", "dashboard"]),
      "compare",
      "split",
      "visual",
      "picture",
      "highlight",
      "sidecallout",
    ]);
  }
  if (["compare", "sidecallout", "split"].includes(mode)) {
    return unique([
      ...(parity === 0 ? ["compare", "sidecallout", "split"] : ["sidecallout", "compare", "split"]),
      "highlight",
      "visual",
      "picture",
      "dashboard",
      "matrix",
      "dense",
    ]);
  }
  if (["visual", "picture", "storyboard", "gallery", "focus"].includes(mode)) {
    return unique([
      ...(parity === 0 ? ["visual", "picture", "storyboard"] : ["picture", "visual", "storyboard"]),
      "gallery",
      "focus",
      "split",
      "compare",
      "sidecallout",
      "highlight",
      "dashboard",
      "matrix",
      "dense",
    ]);
  }
  if (["stack", "highlight", "cards", "closing"].includes(mode)) {
    return unique([
      ...(parity === 0 ? ["stack", "highlight", "cards"] : ["highlight", "stack", "cards"]),
      "closing",
      "sidecallout",
      "compare",
      "visual",
      "picture",
    ]);
  }

  if (balance === "image-heavy") return ["visual", "picture", "gallery", "focus", "split"];
  if (balance === "table-heavy" || balance === "chart-heavy") return ["dashboard", "matrix", "dense", "compare", "split"];
  if (balance === "text-heavy") return ["stack", "highlight", "cards", "closing", "compare"];
  return [];
}

function buildSemanticRefinedSelection(layoutOptions, semanticReview = {}, baseSelection = {}) {
  if (!layoutOptions?.slideLayouts?.length) {
    return {
      selection: { ...(baseSelection || {}) },
      changed: false,
    };
  }

  const issueCategories = normalizeIssueCategories(semanticReview);
  const hasRepeat = issueCategories.has("repeat") || issueCategories.has("icon_repeat") || (semanticReview.repeatedPatterns || []).length > 0;
  const hasTableWidth = issueCategories.has("table_width");
  const hasWhitespace = issueCategories.has("whitespace");
  const hasLayoutBalance = issueCategories.has("layout_balance");
  const hasTextSize = issueCategories.has("text_size");
  const hints = semanticReview.refinementHints || {};
  const parsePageSet = (values = []) =>
    new Set((Array.isArray(values) ? values : []).map((value) => Number(value)).filter((value) => Number.isFinite(value) && value > 0));
  const densePages = parsePageSet([...(hints.densePages || []), ...(hints.tablePages || [])]);
  const spaciousPages = parsePageSet([...(hints.spaciousPages || []), ...(hints.summaryPages || [])]);
  const imagePages = parsePageSet(hints.imagePages || []);
  const tablePages = parsePageSet(hints.tablePages || []);
  const processPages = parsePageSet(hints.processPages || []);
  const pageAdviceList = Array.isArray(semanticReview.pageAdvice) ? semanticReview.pageAdvice : [];
  const pageAdviceMap = new Map(pageAdviceList.map((item) => [Number(item?.page || 0), item]).filter(([page]) => Number.isFinite(page) && page > 0));
  let lastTableVariant = "";
  let lastTableFamily = "";

  const desiredFamiliesByType = {
    summary_cards: ["summary", "dense", "stack", "visual"],
    table_analysis: ["dense", "visual", "split", "compare", "picture"],
    process_flow: ["process", "action", "summary"],
    bullet_columns: ["bullet", "summary"],
    image_story: ["visual", "split", "summary"],
    action_plan: ["action", "process", "dense"],
    key_takeaways: ["stack", "summary"],
  };

  const selection = { ...(baseSelection || {}) };
  const familyUsage = {};
  let changed = false;

  (layoutOptions.slideLayouts || []).forEach((slide, index) => {
    const pageKey = String(slide.page);
    const currentTemplate = selection[pageKey] || slide.currentTemplate || slide.recommendedTemplate || "";
    const options = slide.options || [];
    const currentOption = options.find((option) => option.id === currentTemplate) || options[0];
    if (!currentOption) return;

    const pageNo = Number(slide.page || 0);
    const pageParity = pageNo % 2;
    const currentFamily = currentOption.family || templateVariantFamily(currentOption.variant);
    const tableVariantOrder = slide.type === "table_analysis" ? tableModeVariantOrder(slide.tableMode || "", pageParity, slide.contentMode || slide.density || "") : [];

    const preferredVariantsByType = {
      summary_cards: [
        ...(hasWhitespace ? ["spread", "mosaic"] : []),
        ...(hasRepeat ? ["spread", "mosaic"] : []),
        ...(spaciousPages.has(pageNo) ? ["spread", "bands"] : []),
        ...(densePages.has(pageNo) ? ["dense_grid", "grid"] : []),
        ...(pageParity === 0 ? ["dense_grid", "mosaic"] : ["spread", "bands"]),
        "dense_grid",
        "spread",
        "mosaic",
        "grid",
        "bands",
      ],
      table_analysis: [
        ...(hasTableWidth ? ["dashboard", "dense", "matrix"] : []),
        ...(hasLayoutBalance ? ["compare", "split", "visual"] : []),
        ...(hasRepeat ? ["visual", "compare", "dashboard"] : []),
        ...(hasWhitespace ? ["compare", "highlight", "sidecallout"] : []),
        ...(slide.hasImage || slide.hasScreenshots || imagePages.has(pageNo) ? ["visual", "storyboard", "gallery"] : []),
        ...(tablePages.has(pageNo) || densePages.has(pageNo) || slide.tableRowCount >= 5 || slide.tableColCount >= 4 ? ["dashboard", "dense", "matrix"] : []),
        ...(pageParity === 0 ? ["dashboard", "matrix"] : ["visual", "compare", "highlight"]),
        ...(tableVariantOrder.length ? tableVariantOrder : []),
        "visual",
        "dashboard",
        "compare",
        "sidecallout",
        "highlight",
        "dense",
        "matrix",
        "split",
        "stack",
      ],
      process_flow: [
        ...(processPages.has(pageNo) ? ["bridge", "cards"] : []),
        ...(densePages.has(pageNo) ? ["ladder", "bridge"] : []),
        ...(pageParity === 0 ? ["bridge", "cards"] : ["cards", "bridge", "ladder"]),
        "bridge",
        "cards",
        "ladder",
        "three_lane",
      ],
      bullet_columns: [
        ...(hasTextSize || hasWhitespace ? ["dual", "staggered"] : []),
        ...(pageParity === 0 ? ["masonry", "triple"] : ["staggered", "dual"]),
        "masonry",
        "triple",
        "staggered",
        "dual",
      ],
      image_story: [
        ...(imagePages.has(pageNo) ? ["storyboard", "gallery", "focus"] : []),
        ...(slide.hasImage ? ["storyboard", "focus"] : []),
        ...(pageParity === 0 ? ["storyboard", "gallery"] : ["focus", "split"]),
        "storyboard",
        "gallery",
        "focus",
        "split",
      ],
      action_plan: [
        ...(pageParity === 0 ? ["dashboard", "matrix"] : ["timeline", "stacked"]),
        ...(densePages.has(pageNo) ? ["dashboard", "matrix"] : []),
        "dashboard",
        "matrix",
        "timeline",
        "stacked",
      ],
      key_takeaways: [
        ...(spaciousPages.has(pageNo) ? ["wall", "closing"] : []),
        ...(pageParity === 0 ? ["wall", "closing"] : ["cards", "closing"]),
        "wall",
        "closing",
        "cards",
      ],
    };

    const preferredFamilies = unique([
      ...(desiredFamiliesByType[slide.type] || []),
      ...(Array.isArray(slide.preferredFamilies) ? slide.preferredFamilies : []),
    ]);
    const preferredVariants = unique([
      ...(tableVariantOrder || []),
      ...(Array.isArray(slide.preferredVariants) ? slide.preferredVariants : []),
      ...(slide.type === "table_analysis" ? ["dashboard", "dense", "matrix", "compare", "sidecallout", "highlight", "split", "visual", "picture"] : []),
    ]);

    const advice = pageAdviceMap.get(Number(slide.page));
    if (advice) {
      const preferredFamilies = unique([
        String(advice.preferredFamily || "").toLowerCase(),
        String(advice.preferredVariant || "").toLowerCase(),
      ].filter(Boolean));
      const preferredOption =
        options.find((option) => preferredFamilies.includes(String(option.family || "").toLowerCase())) ||
        options.find((option) => preferredFamilies.includes(String(option.variant || "").toLowerCase())) ||
        options.find((option) => preferredFamilies.includes(String(option.id || "").toLowerCase()));
      if (preferredOption && preferredOption.id && preferredOption.id !== currentTemplate) {
        selection[pageKey] = preferredOption.id;
        changed = true;
        familyUsage[preferredOption.family || currentOption.family || "general"] = (familyUsage[preferredOption.family || currentOption.family || "general"] || 0) + 1;
        if (slide.type === "table_analysis") {
          lastTableFamily = preferredOption.family || currentOption.family || "";
          lastTableVariant = preferredOption.variant || currentOption.variant || "";
        }
        return;
      }
    }

    if (slide.type === "table_analysis" && (slide.hasImage || slide.hasScreenshots)) {
      preferredFamilies.unshift("picture", "visual");
    }
    const familyRank = (family) => {
      const position = preferredFamilies.indexOf(family);
      return position >= 0 ? position : preferredFamilies.length + 8;
    };
    const variantRank = (variant) => {
      const position = preferredVariants.indexOf(variant);
      return position >= 0 ? position : preferredVariants.length + 8;
    };
    const hasImage = Boolean(slide.hasImage || slide.hasScreenshots);
    const ranked = options.map((option) => {
      const family = option.family || templateVariantFamily(option.variant);
      const base = familyRank(family) * 100 + variantRank(option.variant) * 85;
      const repeatPenalty = family === currentFamily ? (slide.type === "table_analysis" ? 240 : 160) : 0;
      const tableRepeatPenalty =
        slide.type === "table_analysis" && (option.variant === lastTableVariant || family === lastTableFamily)
          ? option.variant === lastTableVariant
            ? 300
            : 160
          : 0;
      const usagePenalty = (familyUsage[family] || 0) * (slide.type === "table_analysis" || slide.type === "process_flow" ? 55 : 35);
      const denseBonus =
        (slide.contentMode === "dense" || slide.density === "high" || densePages.has(pageNo)) && ["dense", "matrix", "dashboard", "ladder"].includes(family)
          ? -38
          : (slide.contentMode === "sparse" || slide.density === "low" || spaciousPages.has(pageNo)) && ["summary", "stack", "visual", "cards"].includes(family)
            ? -24
            : 0;
      const imageBonus =
        hasImage && ["visual", "storyboard", "gallery", "focus", "picture"].includes(family)
          ? -30
          : slide.type === "image_story" && ["visual", "storyboard", "gallery", "focus"].includes(family)
            ? -22
            : 0;
      const pageBonus =
        slide.type === "table_analysis"
          ? pageParity === 0
            ? ["dense", "matrix", "dashboard"].includes(family)
              ? -26
              : 0
            : ["visual", "compare", "highlight", "sidecallout"].includes(family)
              ? -22
              : 0
          : slide.type === "process_flow"
            ? pageParity === 0
              ? ["bridge", "cards"].includes(family)
                ? -18
                : 0
              : ["ladder", "three_lane"].includes(family)
                ? -18
                : 0
            : 0;
      const seedBonus = (index + 1) * 0.01;
      return {
        id: option.id,
        family,
        variant: option.variant,
        score: base + repeatPenalty + tableRepeatPenalty + usagePenalty + denseBonus + imageBonus + pageBonus + seedBonus,
      };
    })
      .sort((left, right) => left.score - right.score);

    const best = ranked[0];
    const currentScore = ranked.find((item) => item.id === currentTemplate)?.score ?? Number.POSITIVE_INFINITY;
    const alternateTableOption =
      slide.type === "table_analysis"
        ? ranked.find((item) => item.family && item.family !== lastTableFamily && item.family !== currentFamily)
        : null;
    if (
      slide.type === "table_analysis" &&
      best &&
      lastTableFamily &&
      best.family === lastTableFamily &&
      alternateTableOption &&
      alternateTableOption.score <= best.score + 40
    ) {
      selection[pageKey] = alternateTableOption.id;
      changed = true;
      familyUsage[alternateTableOption.family || currentFamily || "general"] = (familyUsage[alternateTableOption.family || currentFamily || "general"] || 0) + 1;
      lastTableFamily = alternateTableOption.family || currentFamily || "";
      lastTableVariant = alternateTableOption.variant || currentOption.variant || "";
      return;
    }
    if (best && best.id && best.id !== currentTemplate && best.score + 12 < currentScore) {
      selection[pageKey] = best.id;
      changed = true;
    }
    familyUsage[(best?.family || currentFamily) || "general"] = (familyUsage[(best?.family || currentFamily) || "general"] || 0) + 1;
    if (slide.type === "table_analysis") {
      lastTableFamily = best?.family || currentFamily || "";
      lastTableVariant = best?.variant || currentOption.variant || "";
    }
  });

  return { selection, changed };
}

async function resolveReferenceLibraries(materialFiles, options = {}, sessionDir = "") {
  const selectedBaseLibrary = options.referenceLibraryId
    ? resolveLibraryById(options.referenceLibraryId)
    : getDefaultLibrary();

  const uploadedLibraries = [];
  const mergeSummaries = [];

  for (let index = 0; index < (materialFiles || []).length; index += 1) {
    const filePath = materialFiles[index];
    const libraryName = options.referenceNames?.[index] || options.referenceName || path.basename(filePath, path.extname(filePath));
    const extracted = await createReferenceLibraryFromPpt(filePath, undefined, libraryName);
    uploadedLibraries.push(extracted);
    if (options.mergeReference !== false) {
      mergeSummaries.push(mergeReferenceIntoMaster(extracted.dir));
    }
  }

  const masterLibrary = getMasterLibrary();
  const sourceDirs = unique([
    selectedBaseLibrary?.dir,
    masterLibrary?.dir,
    ...uploadedLibraries.map((item) => item.dir),
  ]);

  let effectiveLibrary = selectedBaseLibrary?.dir ? summarizeLibrary(selectedBaseLibrary.dir) : selectedBaseLibrary;
  let reusablePath = selectedBaseLibrary?.reusablePath || "";

  if (sourceDirs.length > 1 && sessionDir) {
    const combinedDir = path.join(sessionDir, "reference_effective");
    const combined = composeReferenceLibraries({
      sourceDirs,
      targetDir: combinedDir,
      readmeTitle: "当前会话参考库",
    });
    effectiveLibrary = combined.summary;
    reusablePath = combined.reusablePath;
  } else if (sourceDirs.length === 1 && sourceDirs[0]) {
    effectiveLibrary = summarizeLibrary(sourceDirs[0]);
    reusablePath = effectiveLibrary.reusablePath;
  }

  return {
    uploadedLibraries,
    selectedBaseLibrary: selectedBaseLibrary?.dir ? summarizeLibrary(selectedBaseLibrary.dir) : selectedBaseLibrary,
    masterLibrary,
    effectiveLibrary,
    mergeSummaries,
    reusablePath,
  };
}

async function renderDraftArtifacts({ sessionDir, outline, style }) {
  const draftDeckPath = path.join(sessionDir, "intermediate", "draft_generated.pptx");
  const previewDir = path.join(sessionDir, "intermediate", "rendered_preview");
  await reportGenerator.renderDeck(outline, style, draftDeckPath);
  const previewResult = await buildRenderedPreviews(
    draftDeckPath,
    previewDir,
    (outline.slides || []).map((slide) => slide.title),
    { width: 1280, height: 720, timeoutMs: 90000 },
  );
  const previewState = normalizePreviewState(previewResult, "draft");
  return { draftDeckPath, previewDir, previews: previewResult.previews || [], previewState };
}

async function createWorkflowSession({ files, options = {} }) {
  const providedSessionId = options.sessionId || "";
  const providedSessionDir = options.sessionDir || "";
  const session = providedSessionId && providedSessionDir ? { sessionId: providedSessionId, sessionDir: providedSessionDir } : createSession();
  const { sessionId, sessionDir } = session;
  const writeProgress = createProgressWriter(sessionDir);
  const previewEnvironment = detectPreviewSupport();
    const stagedFiles = {
      materialPpts: stageReferenceFiles(sessionDir, files.materialPpt || [], "material", ".pptx"),
      referencePpts: stageReferenceFiles(sessionDir, files.referencePpt || [], "reference-ppt", ".pptx"),
      referenceImages: stageReferenceFiles(sessionDir, files.referenceImages || [], "reference-image", ".png"),
    };

  saveSessionMeta(sessionDir, {
    sessionId,
    sessionDir,
    createdAt: new Date().toISOString(),
    status: "running",
    progress: 2,
    currentStage: "upload",
    currentStageLabel: "上传整理",
    message: "正在整理上传文件",
    previewEnvironment,
    files: stagedFiles,
  });

  const primaryTemplate = pickPrimaryUpload(files.template || []);
  const primaryRequirementDoc = pickPrimaryUpload(files.requirementDoc || []);

    if (primaryTemplate) {
      stagedFiles.template = stageUploadedFile(
        sessionDir,
        primaryTemplate,
        buildStagedFileName("template", 0, primaryTemplate.originalname || primaryTemplate.path || "", ".pptx"),
      );
    }
    if (primaryRequirementDoc) {
      stagedFiles.requirementDoc = stageUploadedFile(
        sessionDir,
        primaryRequirementDoc,
        buildStagedFileName("requirement", 0, primaryRequirementDoc.originalname || primaryRequirementDoc.path || "", ".docx"),
      );
    }

  if (!stagedFiles.template || !stagedFiles.requirementDoc) {
    throw new Error("请同时上传空白模板 PPT 和需求文档，并确保不要选择以 ~$ 开头的临时 Office 文件。");
  }

  writeProgress({
    status: "running",
    progress: 6,
    currentStage: "upload",
    currentStageLabel: "上传整理",
    message: "上传文件已成功归档到当前会话。",
    files: stagedFiles,
  });

  const [reference, derivedReferenceImages] = await Promise.all([
    resolveReferenceLibraries(stagedFiles.materialPpts, options, sessionDir),
    exportReferencePptImages(sessionDir, stagedFiles.referencePpts),
  ]);
  const effectiveReferenceImages = unique([...(stagedFiles.referenceImages || []), ...(derivedReferenceImages || [])]);
  writeProgress({
    status: "running",
    progress: 16,
    currentStage: "reference",
    currentStageLabel: "参考素材融合",
    message: `已加载 ${reference.uploadedLibraries?.length || 0} 份素材 PPT，并拆分出 ${derivedReferenceImages.length || 0} 张参考页图。`,
    workflowStages: buildReadableWorkflowStagesCn({
      documentSummary: null,
      effectiveLibrary: reference.effectiveLibrary,
      referenceStyleProfile: { count: derivedReferenceImages.length },
      semanticAnalysis: null,
      semanticReview: null,
      pageCount: 0,
      referencePptCount: stagedFiles.referencePpts.length,
      derivedReferenceImageCount: derivedReferenceImages.length,
    }),
  });
  const assetDir = path.join(sessionDir, "intermediate", "assets");
  ensureDir(assetDir);

  const [doc, template] = await Promise.all([
    reportGenerator.extractInputDocument(stagedFiles.requirementDoc, assetDir),
    reportGenerator.extractTemplate(stagedFiles.template),
  ]);
  writeProgress({
    status: "running",
    progress: 26,
    currentStage: "parse",
    currentStageLabel: "文档解析",
    message: "需求文档和模板已完成解析。",
  });
  const documentStructure = reportGenerator.buildDocumentStructure(doc);
  const detectedSections = documentStructure.sections || reportGenerator.detectSections(doc);
  const documentSummary = reportGenerator.buildDocumentSummary(doc, stagedFiles.requirementDoc);
  documentSummary.counts.sections = documentStructure.topLevel?.length || detectedSections.length;
  writeProgress({
    status: "running",
    progress: 28,
    currentStage: "parse",
    currentStageLabel: "文档解析",
    message: `已识别 ${documentSummary.counts.sections || 0} 个一级标题，将严格按“一、二、三、四”顺序规划页面。`,
    workflowStages: buildReadableWorkflowStagesCn({
      documentSummary,
      effectiveLibrary: reference.effectiveLibrary,
      referenceStyleProfile: null,
      semanticAnalysis: null,
      semanticReview: null,
      pageCount: 0,
      referencePptCount: stagedFiles.referencePpts.length,
      derivedReferenceImageCount: derivedReferenceImages.length,
    }),
  });

  const referenceStyleProfile = {
    ...analyzeReferenceImages(effectiveReferenceImages),
  };
  const referenceStyleSampleImages = sampleReferenceImages(effectiveReferenceImages, 12);
  referenceStyleProfile.sampledReferenceImageCount = referenceStyleSampleImages.length;
  const resolvedSemanticProvider = resolveSemanticProvider(options);
  const semanticModelOptions = {
    provider: resolvedSemanticProvider,
    baseUrl: options.semanticBaseUrl || "",
    apiKey: resolveSemanticApiKey(options),
    model: resolveSemanticModelName(options),
    supportsImages: options.semanticSupportsImages,
    allowRepair: true,
    compactPayload: resolvedSemanticProvider === "minimax",
    timeoutMs: Math.max(8000, Number(options.semanticTimeoutMs || 60000)),
  };
  const semanticDocumentPayload = {
    doc,
    documentSummary,
    sections: documentStructure.topLevel || documentStructure.sections || detectedSections,
    referenceStyleProfile,
    requestedPages: options.pageCount || 0,
    title: options.title || "",
    department: options.department || "",
    presenter: options.presenter || "",
  };
  writeProgress({
    status: "running",
    progress: 34,
    currentStage: "semantic",
    currentStageLabel: "语义规划",
    message: `语义模型调用中：${semanticModelOptions.provider || "local"} / ${semanticModelOptions.model || "deepseek-r1:14b"}`,
    semanticModelProvider: semanticModelOptions.provider || "local",
    semanticModelName: semanticModelOptions.model || "deepseek-r1:14b",
    semanticModelAvailable: semanticModelOptions.provider === "local" ? true : Boolean(semanticModelOptions.apiKey),
    semanticAnalysis: {
      provider: semanticModelOptions.provider || "local",
      model: semanticModelOptions.model || "deepseek-r1:14b",
      available: semanticModelOptions.provider === "local" ? true : Boolean(semanticModelOptions.apiKey),
      parseStatus: "running",
      summary: "语义模型正在分析文档结构与参考风格。",
    },
    workflowStages: buildReadableWorkflowStagesCn({
      documentSummary,
      effectiveLibrary: reference.effectiveLibrary,
      referenceStyleProfile,
      semanticAnalysis: {
        provider: semanticModelOptions.provider || "local",
        model: semanticModelOptions.model || "deepseek-r1:14b",
        available: semanticModelOptions.provider === "local" ? true : Boolean(semanticModelOptions.apiKey),
        parseStatus: "running",
      },
      semanticReview: null,
      pageCount: 0,
      referencePptCount: stagedFiles.referencePpts.length,
      derivedReferenceImageCount: derivedReferenceImages.length,
    }),
  });
  const [semanticReferenceStyle, semanticAnalysis, sectionSemanticPlans] = await Promise.all([
    analyzeReferenceStyle(referenceStyleSampleImages, referenceStyleProfile, semanticModelOptions),
    analyzeDocumentSemantics(semanticDocumentPayload, semanticModelOptions),
    buildSectionSemanticPlans((documentStructure.topLevel || []).slice(0, 3), semanticModelOptions),
  ]);
  writeProgress({
    status: "running",
    progress: 38,
    currentStage: "semantic",
    currentStageLabel: "语义规划",
    message: `语义模型已返回：${semanticModelOptions.provider || "local"} / ${semanticModelOptions.model || "deepseek-r1:14b"}`,
    semanticAnalysis,
    workflowStages: buildReadableWorkflowStagesCn({
      documentSummary,
      effectiveLibrary: reference.effectiveLibrary,
      referenceStyleProfile,
      semanticAnalysis,
      semanticReview: null,
      pageCount: 0,
      referencePptCount: stagedFiles.referencePpts.length,
      derivedReferenceImageCount: derivedReferenceImages.length,
    }),
  });
  if (semanticReferenceStyle?.available) {
    referenceStyleProfile.semanticAnalysis = semanticReferenceStyle;
    referenceStyleProfile.semanticSummary = semanticReferenceStyle.summary || semanticReferenceStyle.styleFamily || "";
    if (semanticReferenceStyle.styleFamily) {
      referenceStyleProfile.styleFamily = semanticReferenceStyle.styleFamily;
    }
    if (semanticReferenceStyle.densityBias) {
      referenceStyleProfile.densityBias = semanticReferenceStyle.densityBias;
    }
    if (Array.isArray(semanticReferenceStyle.layoutBias) && semanticReferenceStyle.layoutBias.length) {
      referenceStyleProfile.templateBias = unique([...(referenceStyleProfile.templateBias || []), ...semanticReferenceStyle.layoutBias]);
    }
    if (Array.isArray(semanticReferenceStyle.preferredVariants) && semanticReferenceStyle.preferredVariants.length) {
      referenceStyleProfile.preferredVariants = unique([...(referenceStyleProfile.preferredVariants || []), ...semanticReferenceStyle.preferredVariants]);
    }
    if (Array.isArray(semanticReferenceStyle.tableStyleBias) && semanticReferenceStyle.tableStyleBias.length) {
      referenceStyleProfile.tableStyleBias = unique([...(referenceStyleProfile.tableStyleBias || []), ...semanticReferenceStyle.tableStyleBias]);
    }
    if (semanticReferenceStyle.headerStyle) {
      referenceStyleProfile.headerStyle = semanticReferenceStyle.headerStyle;
    }
    if (semanticReferenceStyle.summaryBandStyle) {
      referenceStyleProfile.summaryBandStyle = semanticReferenceStyle.summaryBandStyle;
    }
    if (semanticReferenceStyle.tablePreference) {
      referenceStyleProfile.tablePreference = semanticReferenceStyle.tablePreference;
    }
    if (semanticReferenceStyle.cardStyle) {
      referenceStyleProfile.cardStyle = semanticReferenceStyle.cardStyle;
    }
    if (semanticReferenceStyle.pageRhythm) {
      referenceStyleProfile.pageRhythm = semanticReferenceStyle.pageRhythm;
    }
    if (semanticReferenceStyle.imagePlacement) {
      referenceStyleProfile.imagePlacement = semanticReferenceStyle.imagePlacement;
    }
    if (semanticReferenceStyle.iconDiversity || semanticReferenceStyle.iconDiversityPolicy) {
      referenceStyleProfile.iconDiversity = semanticReferenceStyle.iconDiversity || semanticReferenceStyle.iconDiversityPolicy;
    }
  }
  const layoutLibraryPath = options.layoutLibraryPath || DEFAULT_LAYOUT_LIBRARY;
  const requestedLayoutSet = options.layoutSet || referenceStyleProfile.suggestedLayoutSet || "";
  const layoutLibrary = reportGenerator.loadLayoutLibrary(layoutLibraryPath, requestedLayoutSet);
  const materials = reportGenerator.loadReferenceLibrary(reference.reusablePath);

  const rawOutline = reportGenerator.buildOutline(doc, {
    ...options,
    word: stagedFiles.requirementDoc,
    template: stagedFiles.template,
    pages: options.pageCount || 0,
    documentStructure,
    sectionSemanticPlans,
    semanticAnalysis,
  });
  writeProgress({
    status: "running",
    progress: 52,
    currentStage: "planning",
    currentStageLabel: "动态页面规划",
    message: `已根据文档结构规划出 ${rawOutline.meta.pages || 0} 页内容。`,
    workflowStages: buildReadableWorkflowStagesCn({
      documentSummary,
      effectiveLibrary: reference.effectiveLibrary,
      referenceStyleProfile,
      semanticAnalysis,
      semanticReview: null,
      pageCount: rawOutline.meta.pages || 0,
      referencePptCount: stagedFiles.referencePpts.length,
      derivedReferenceImageCount: derivedReferenceImages.length,
    }),
  });
  let layoutOptions = buildLayoutOptions(layoutLibraryPath, rawOutline, layoutLibrary.setName, {}, referenceStyleProfile);
  const draftLayoutSelection = layoutOptions.initialSelection || {};
  let outline = reportGenerator.applyTemplatesToOutline(rawOutline, layoutLibrary, draftLayoutSelection);

  const style = {
    ...reportGenerator.buildStyle(template, "", materials, layoutLibrary, {
      ...options,
      referenceStyleProfile,
    }),
    layoutLibrary,
    layoutLibraryPath,
    referenceLibrary: reference.reusablePath,
    referenceStyleProfile,
  };

  let notes = reportGenerator.buildNotes(outline);
  let { draftDeckPath, previewDir, previews, previewState: draftPreviewState } = await renderDraftArtifacts({ sessionDir, outline, style });
  writeProgress({
    status: "running",
    progress: 74,
    currentStage: "draft",
    currentStageLabel: "草稿输出",
    message: previewStatusMessage("草稿 PPT", draftPreviewState),
    previewState: draftPreviewState,
    workflowStages: buildReadableWorkflowStagesCn({
      documentSummary,
      effectiveLibrary: reference.effectiveLibrary,
      referenceStyleProfile,
      semanticAnalysis,
      semanticReview: null,
      pageCount: outline.meta.pages,
      referencePptCount: stagedFiles.referencePpts.length,
      derivedReferenceImageCount: derivedReferenceImages.length,
    }),
  });
  let semanticReview = {
    available: false,
    provider: resolvedSemanticProvider || "",
    model: resolveSemanticModelName(options) || "",
    parseStatus: "skipped",
    summary: "Initial draft preview has been skipped to shorten wait time.",
    issues: [],
  };

  const intermediateDir = path.join(sessionDir, "intermediate");
  const semanticRefinedSelectionPath = path.join(intermediateDir, "semantic_refined_selection.json");
  const semanticRefinedSelection = draftLayoutSelection;
  layoutOptions = buildLayoutOptions(layoutLibraryPath, outline, layoutLibrary.setName, semanticRefinedSelection, referenceStyleProfile);
  writeProgress({
    status: "running",
    progress: 84,
    currentStage: "review",
    currentStageLabel: "效果复核",
    message: semanticReview.summary || "效果复核已完成。",
    workflowStages: buildReadableWorkflowStagesCn({
      documentSummary,
      effectiveLibrary: reference.effectiveLibrary,
      referenceStyleProfile,
      semanticAnalysis,
      semanticReview,
      pageCount: outline.meta.pages,
      referencePptCount: stagedFiles.referencePpts.length,
      derivedReferenceImageCount: derivedReferenceImages.length,
    }),
  });
  const outlinePath = path.join(intermediateDir, "outline.json");
  const stylePath = path.join(intermediateDir, "style.json");
  const notesPath = path.join(intermediateDir, "page_notes.md");
  const summaryPath = path.join(intermediateDir, "document_summary.json");
  const structurePath = path.join(intermediateDir, "document_structure.json");
  const structureMarkdownPath = path.join(intermediateDir, "document_structure.md");
  const referencePath = path.join(intermediateDir, "reference_summary.json");
  const referenceStylePath = path.join(intermediateDir, "reference_style.json");
  const semanticAnalysisPath = path.join(intermediateDir, "semantic_analysis.json");
  const sectionSemanticPlansPath = path.join(intermediateDir, "section_semantic_plans.json");
  const semanticReviewPath = path.join(intermediateDir, "semantic_review.json");
  const layoutOptionsPath = path.join(intermediateDir, "layout_options.json");

  writeJson(outlinePath, outline);
  writeJson(stylePath, style);
  writeJson(summaryPath, documentSummary);
  writeJson(structurePath, documentStructure);
  writeJson(referencePath, {
    uploadedLibraries: reference.uploadedLibraries,
    selectedBaseLibrary: reference.selectedBaseLibrary,
    masterLibrary: reference.masterLibrary,
    effectiveLibrary: reference.effectiveLibrary,
    mergeSummaries: reference.mergeSummaries,
  });
  writeJson(referenceStylePath, referenceStyleProfile);
  writeJson(semanticAnalysisPath, semanticAnalysis || {});
  writeJson(sectionSemanticPlansPath, sectionSemanticPlans || []);
  writeJson(semanticReviewPath, semanticReview || {});
  writeJson(semanticRefinedSelectionPath, semanticRefinedSelection || {});
  writeJson(layoutOptionsPath, layoutOptions);
  fs.writeFileSync(notesPath, notes, "utf8");
  fs.writeFileSync(structureMarkdownPath, documentStructure.markdown || "", "utf8");

  const meta = {
    sessionId,
    sessionDir,
    createdAt: new Date().toISOString(),
    status: "running",
    progress: 86,
    currentStage: "review",
    currentStageLabel: "效果复核",
    message: "草稿已完成，正在整理输出文件。",
    files: stagedFiles,
    pageCount: outline.meta.pages,
    referenceLibraryPath: reference.reusablePath,
    referenceStylePath,
    referenceStyleProfile,
    referenceImagePaths: effectiveReferenceImages,
    uploadedReferenceImages: stagedFiles.referenceImages,
    referencePptPaths: stagedFiles.referencePpts,
    derivedReferenceImageCount: derivedReferenceImages.length,
    previewState: draftPreviewState,
    layoutLibraryPath,
    layoutSet: layoutLibrary.setName,
    draftLayoutSelection: semanticRefinedSelection,
    semanticRefinedSelection,
    semanticModelProvider: resolveSemanticProvider(options),
    semanticModelBaseUrl: options.semanticBaseUrl || process.env.SEMANTIC_MODEL_BASE_URL || "",
    semanticModelName: resolveSemanticModelName(options),
    semanticModelAvailable: semanticAnalysis?.available ?? resolveSemanticAvailability(options),
    semanticModelSupportsImages: typeof options.semanticSupportsImages === "boolean" ? options.semanticSupportsImages : undefined,
    semanticModelApiKey: resolveSemanticApiKey(options),
    layoutOptions,
    intermediate: {
      outlinePath,
      stylePath,
      notesPath,
      documentSummaryPath: summaryPath,
      documentStructurePath: structurePath,
      documentStructureMarkdownPath: structureMarkdownPath,
      referenceSummaryPath: referencePath,
      referenceStylePath,
      semanticAnalysisPath,
      sectionSemanticPlansPath,
      semanticReviewPath,
      semanticRefinedSelectionPath,
      layoutOptionsPath,
      draftDeckPath,
      previewDir,
      previewState: draftPreviewState,
    },
  };
  updateSessionMeta(sessionDir, meta);

  return {
    sessionId,
    outline,
    style,
    notes,
    documentSummary,
    previewEnvironment,
    previews,
    previewState: draftPreviewState,
    uploadedMaterialLibraries: reference.uploadedLibraries,
    uploadedReferenceLibraries: [],
    uploadedReferenceImages: stagedFiles.referenceImages,
    uploadedReferencePpts: stagedFiles.referencePpts,
    derivedReferenceImageCount: derivedReferenceImages.length,
    selectedReferenceLibrary: reference.selectedBaseLibrary,
    masterReferenceLibrary: reference.masterLibrary,
    effectiveReferenceLibrary: reference.effectiveLibrary,
    referenceStyleProfile,
    referenceImagePaths: effectiveReferenceImages,
    mergeSummaries: reference.mergeSummaries,
    layoutOptions,
    draftLayoutSelection: semanticRefinedSelection,
    semanticRefinedSelection,
    workflowStages: buildReadableWorkflowStagesCn({
      documentSummary,
      effectiveLibrary: reference.effectiveLibrary,
      referenceStyleProfile,
      semanticAnalysis,
      semanticReview,
      pageCount: outline.meta.pages,
      referencePptCount: stagedFiles.referencePpts.length,
      derivedReferenceImageCount: derivedReferenceImages.length,
    }),
    draftDownloads: buildReadableDownloadManifestCn(sessionId),
    outputFiles: {
      outlinePath,
      stylePath,
      notesPath,
      summaryPath,
      structurePath,
      structureMarkdownPath,
      referencePath,
      referenceStylePath,
      semanticAnalysisPath,
      sectionSemanticPlansPath,
      semanticReviewPath,
      semanticRefinedSelectionPath,
      layoutOptionsPath,
      draftDeckPath,
    },
    semanticAnalysis,
    sectionSemanticPlans,
    semanticReview,
    semanticRefinedSelection,
  };
}

async function startWorkflowSession({ files, options = {} }) {
  const { sessionId, sessionDir } = createSession();
  const previewEnvironment = detectPreviewSupport();
  saveSessionMeta(sessionDir, {
    sessionId,
    sessionDir,
    createdAt: new Date().toISOString(),
    status: "queued",
    progress: 2,
    currentStage: "upload",
    currentStageLabel: "上传整理",
    message: "任务已进入队列。",
    semanticModelProvider: resolveSemanticProvider(options) || (options.semanticBaseUrl ? "local" : "") || (process.env.MINIMAX_API_KEY ? "minimax" : ""),
    semanticModelBaseUrl: options.semanticBaseUrl || process.env.SEMANTIC_MODEL_BASE_URL || "",
    semanticModelName: resolveSemanticModelName(options),
    semanticModelAvailable: resolveSemanticAvailability(options),
    previewEnvironment,
  });

  setImmediate(async () => {
    try {
      const result = await createWorkflowSession({
        files,
        options: {
          ...options,
          sessionId,
          sessionDir,
        },
      });
      updateSessionMeta(sessionDir, {
        ...result,
        status: "completed",
        progress: 100,
        currentStage: "completed",
        currentStageLabel: "处理完成",
        message: "草稿已就绪，可以继续生成最终 PPT。",
        semanticModelProvider: result.semanticAnalysis?.provider || resolveSemanticProvider(options),
        semanticModelName: result.semanticAnalysis?.model || resolveSemanticModelName(options),
        semanticModelAvailable: result.semanticAnalysis?.available ?? resolveSemanticAvailability(options),
        result,
      });
    } catch (error) {
      updateSessionMeta(sessionDir, {
        status: "failed",
        progress: 100,
        currentStage: "failed",
        currentStageLabel: "处理失败",
        message: error.message || String(error),
        error: error.message || String(error),
        workflowStages: buildReadableWorkflowStagesCn({
          documentSummary: null,
          effectiveLibrary: null,
          referenceStyleProfile: null,
          semanticAnalysis: {
            provider: resolveSemanticProvider(options) || "未启用",
            model: resolveSemanticModelName(options) || "未配置",
            available: false,
            parseStatus: "error",
            summary: error.message || String(error),
          },
          semanticReview: {
            available: false,
            parseStatus: "unknown",
          },
          pageCount: 0,
          referencePptCount: 0,
          derivedReferenceImageCount: 0,
        }),
      });
    }
  });

  return {
    sessionId,
    sessionDir,
    status: "queued",
    progress: 2,
    currentStage: "upload",
    currentStageLabel: "上传整理",
    message: "任务已进入队列。",
    semanticModelProvider: resolveSemanticProvider(options),
    semanticModelBaseUrl: options.semanticBaseUrl || process.env.SEMANTIC_MODEL_BASE_URL || "",
    semanticModelName: resolveSemanticModelName(options),
    semanticModelAvailable: resolveSemanticAvailability(options),
    previewEnvironment,
  };
}

async function generateWorkflowDeck({ sessionId, outline, style, layoutSelection = {}, layoutSet = "" }) {
  const session = loadSessionMeta(sessionId);
  if (!session.meta) {
    throw new Error(`未找到会话：${sessionId}`);
  }

  const finalOutlineInput = outline || readJson(session.meta.intermediate.outlinePath, null);
  const finalStyleInput = style || readJson(session.meta.intermediate.stylePath, null);
  if (!finalOutlineInput || !finalStyleInput) {
    throw new Error("Missing intermediate Outline or Style.");
  }

  const delivery = prepareFinalDeliveryWorkspace(sessionId);
  const outputDir = delivery.outputDir;

  const referenceStyleProfile = session.meta.referenceStyleProfile || finalStyleInput.referenceStyleProfile || null;
  const layoutLibrary = reportGenerator.loadLayoutLibrary(
    session.meta.layoutLibraryPath,
    layoutSet || session.meta.layoutSet || referenceStyleProfile?.suggestedLayoutSet || "",
  );
  const savedLayoutOptions = readJson(session.meta.intermediate.layoutOptionsPath, {});
  const savedSemanticRefinedSelection =
    readJson(session.meta.intermediate.semanticRefinedSelectionPath, null) ||
    session.meta.semanticRefinedSelection ||
    null;
  const semanticBaseSelection = Object.keys(savedSemanticRefinedSelection || {}).length
    ? savedSemanticRefinedSelection
    : savedLayoutOptions.initialSelection || session.meta.draftLayoutSelection || {};
  const effectiveLayoutSelection = Object.keys(layoutSelection || {}).length
    ? { ...semanticBaseSelection, ...(layoutSelection || {}) }
    : semanticBaseSelection;
  const finalOutline = reportGenerator.applyTemplatesToOutline(finalOutlineInput, layoutLibrary, effectiveLayoutSelection);
  const finalStyle = {
    ...finalStyleInput,
    materials: reportGenerator.loadReferenceLibrary(session.meta.referenceLibraryPath),
    referenceLibrary: session.meta.referenceLibraryPath,
    referenceStyleProfile,
    layoutLibrary,
    layoutLibraryPath: session.meta.layoutLibraryPath,
  };

  const deckPath = path.join(outputDir, "workflow_generated.pptx");
  await reportGenerator.renderDeck(finalOutline, finalStyle, deckPath);
  const previewDir = path.join(outputDir, "rendered_preview");
  const finalPreviewResult = await buildRenderedPreviews(
    deckPath,
    previewDir,
    (finalOutline.slides || []).map((slide) => slide.title),
    { width: 1280, height: 720, timeoutMs: 90000 },
  );
  const finalPreviewState = normalizePreviewState(finalPreviewResult, "final");
  const finalPreviews = finalPreviewResult.previews || [];
  const finalSemanticReview = await reviewRenderedDeck(
    finalPreviews || [],
    finalOutline,
    finalStyle,
    {
      provider: session.meta.semanticModelProvider || process.env.SEMANTIC_MODEL_PROVIDER || "",
      baseUrl: session.meta.semanticModelBaseUrl || process.env.SEMANTIC_MODEL_BASE_URL || "",
      apiKey:
        session.meta.semanticModelApiKey ||
        process.env.SEMANTIC_MODEL_API_KEY ||
        process.env.MINIMAX_API_KEY ||
        process.env.OPENAI_API_KEY ||
        "",
      model: session.meta.semanticModelName || process.env.SEMANTIC_MODEL_NAME || "",
      supportsImages: session.meta.semanticModelSupportsImages,
    },
  );

  const notes = reportGenerator.buildNotes(finalOutline);
  const layoutManifest = reportGenerator.buildLayoutTemplateManifest(finalStyle, finalOutline);
  const sourceDocumentSummary = readJson(session.meta.intermediate.documentSummaryPath, {});
  const sourceDocumentStructure = readJson(session.meta.intermediate.documentStructurePath, {});
  const sourceDocumentStructureMarkdown = fs.existsSync(session.meta.intermediate.documentStructureMarkdownPath)
    ? fs.readFileSync(session.meta.intermediate.documentStructureMarkdownPath, "utf8")
    : sourceDocumentStructure.markdown || "";
  const sourceSemanticAnalysis = readJson(session.meta.intermediate.semanticAnalysisPath, {});

  const outlinePath = path.join(outputDir, "outline.final.json");
  const stylePath = path.join(outputDir, "style.final.json");
  const layoutPath = path.join(outputDir, "layout.final.json");
  const summaryPath = path.join(outputDir, "document_summary.final.json");
  const structurePath = path.join(outputDir, "document_structure.final.json");
  const structureMarkdownPath = path.join(outputDir, "document_structure.final.md");
  const referenceStylePath = path.join(outputDir, "reference_style.final.json");
  const semanticAnalysisPath = path.join(outputDir, "semantic_analysis.final.json");
  const archivedManifestPath = path.join(outputDir, "archived_manifest.json");
  const notesPath = path.join(outputDir, "page_notes.final.md");
  const semanticReviewPath = path.join(outputDir, "semantic_review.final.json");
  const semanticRefinedSelectionPath = path.join(outputDir, "semantic_refined_selection.final.json");

  writeJson(outlinePath, finalOutline);
  writeJson(stylePath, finalStyle);
  writeJson(layoutPath, layoutManifest);
  writeJson(summaryPath, sourceDocumentSummary || {});
  writeJson(structurePath, sourceDocumentStructure || {});
  fs.writeFileSync(structureMarkdownPath, sourceDocumentStructureMarkdown || "", "utf8");
  writeJson(referenceStylePath, referenceStyleProfile || {});
  writeJson(semanticAnalysisPath, sourceSemanticAnalysis || {});
  writeJson(semanticReviewPath, finalSemanticReview || {});
  writeJson(semanticRefinedSelectionPath, effectiveLayoutSelection || {});
  fs.writeFileSync(notesPath, notes, "utf8");

  const finalMeta = {
    ...session.meta,
    generatedAt: new Date().toISOString(),
    status: "completed",
    progress: 100,
    currentStage: "completed",
    currentStageLabel: "处理完成",
    message: `${previewStatusMessage("最终 PPT", finalPreviewState)} 原始会话工作区已销毁。`,
    previewState: finalPreviewState,
    layoutSet: layoutLibrary.setName,
    layoutOptions: buildLayoutOptions(
      session.meta.layoutLibraryPath,
      finalOutline,
      layoutLibrary.setName,
      effectiveLayoutSelection,
      referenceStyleProfile,
    ),
    referenceStyleProfile,
    semanticRefinedSelection: effectiveLayoutSelection,
    semanticReview: finalSemanticReview,
    archivedAt: new Date().toISOString(),
    deliveryDir: delivery.deliveryDir,
    sourceSessionDir: session.sessionDir,
    sessionDir: delivery.deliveryDir,
    intermediate: null,
    output: {
      deckPath,
      outlinePath,
      stylePath,
      layoutPath,
      summaryPath,
      structurePath,
      structureMarkdownPath,
      referenceStylePath,
      semanticAnalysisPath,
      archivedManifestPath,
      semanticReviewPath,
      semanticRefinedSelectionPath,
      notesPath,
      previewDir,
      previewState: finalPreviewState,
    },
  };

  const archivedMeta = {
    ...finalMeta,
    files: undefined,
    outputFiles: undefined,
    draftDownloads: undefined,
    uploadedReferenceImages: undefined,
    uploadedReferencePpts: undefined,
    referencePptPaths: undefined,
    referenceImagePaths: undefined,
    intermediate: null,
  };
  delete archivedMeta.files;
  delete archivedMeta.outputFiles;
  delete archivedMeta.draftDownloads;
  delete archivedMeta.uploadedReferenceImages;
  delete archivedMeta.uploadedReferencePpts;
  delete archivedMeta.referencePptPaths;
  delete archivedMeta.referenceImagePaths;

  const archivedManifest = buildArchivedDeliveryManifestCn({
    sessionId,
    archivedAt: finalMeta.archivedAt,
    deliveryDir: delivery.deliveryDir,
    sourceSessionDir: session.sessionDir,
    files: [
      { type: "deck", title: "最终 PPT", path: deckPath },
      { type: "outline", title: "最终 Outline", path: outlinePath },
      { type: "style", title: "最终 Style", path: stylePath },
      { type: "layout", title: "最终布局", path: layoutPath },
      { type: "summary", title: "文档摘要", path: summaryPath },
      { type: "structure", title: "标题结构图", path: structurePath },
      { type: "structure-md", title: "Markdown 结构稿", path: structureMarkdownPath },
      { type: "notes", title: "最终备注", path: notesPath },
    ],
  });
  writeJson(archivedManifestPath, archivedManifest);
  archivedMeta.output.archivedManifestPath = archivedManifestPath;

  const archivedSnapshot = updateSessionMeta(session.sessionDir, archivedMeta);
  saveSessionMeta(delivery.deliveryDir, archivedSnapshot);
  const destroyResult = destroySessionWorkspace(session.sessionDir);

  return {
    sessionId,
    deckPath,
    finalPreviews,
    previewState: finalPreviewState,
    layoutOptions: buildLayoutOptions(
      session.meta.layoutLibraryPath,
      finalOutline,
      layoutLibrary.setName,
      effectiveLayoutSelection,
      referenceStyleProfile,
    ),
    downloadManifest: buildReadableFinalDownloadManifestCn(sessionId),
    files: {
      deckPath,
      outlinePath,
      stylePath,
      layoutPath,
      summaryPath,
      structurePath,
      structureMarkdownPath,
      referenceStylePath,
      semanticAnalysisPath,
      archivedManifestPath,
      semanticReviewPath,
      semanticRefinedSelectionPath,
      notesPath,
      previewDir,
    },
    semanticReview: finalSemanticReview,
    archivedAt: archivedSnapshot.archivedAt || finalMeta.archivedAt,
    deliveryDir: delivery.deliveryDir,
    sourceSessionDir: session.sessionDir,
    destroyResult,
  };
}

module.exports = {
  createWorkflowSession,
  startWorkflowSession,
  generateWorkflowDeck,
  buildLayoutOptions,
  normalizePreviewState,
  previewStatusMessage,
};
