const { fileToDataUrl } = require("./referenceStyleService");
const { templateVariantFamily } = require("./layoutSelectionService");
const {
  buildStylePromptV2,
  buildImagePromptContentV2,
  buildReviewPromptV2,
  buildReviewTextContentV2,
} = require("./semanticPrompts");

function unique(list) {
  return [...new Set((list || []).filter(Boolean))];
}

function normalizeText(text) {
  return String(text || "")
    .replace(/\s+/g, " ")
    .replace(/[\u0000-\u001f]+/g, " ")
    .trim();
}

function clip(text, maxLength = 160) {
  const value = normalizeText(text);
  if (value.length <= maxLength) return value;
  return `${value.slice(0, Math.max(24, maxLength - 1)).trim()}…`;
}

function buildProviderConfig(overrides = {}) {
  const explicitBaseUrl = overrides.baseUrl || process.env.SEMANTIC_MODEL_BASE_URL || "";
  const explicitProvider = String(overrides.provider || process.env.SEMANTIC_MODEL_PROVIDER || "").toLowerCase();
  const provider = String(
    explicitProvider ||
      (explicitBaseUrl ? "local" : "") ||
      (process.env.MINIMAX_API_KEY ? "minimax" : process.env.OPENAI_API_KEY ? "openai" : "off"),
  ).toLowerCase();
  const baseUrl =
    provider === "minimax"
      ? "https://api.minimax.io/v1"
      : provider === "openai"
        ? explicitBaseUrl || "https://api.openai.com/v1"
        : provider === "local"
          ? explicitBaseUrl || "http://127.0.0.1:11434/v1"
          : explicitBaseUrl || process.env.SEMANTIC_MODEL_BASE_URL || "";
  const apiKey =
    overrides.apiKey ||
    process.env.SEMANTIC_MODEL_API_KEY ||
    process.env.MINIMAX_API_KEY ||
    process.env.OPENAI_API_KEY ||
    "";
  const model =
    provider === "minimax"
      ? "MiniMax-M2.7"
      : overrides.model ||
        process.env.SEMANTIC_MODEL_NAME ||
        (provider === "local" ? "deepseek-r1:14b" : provider === "openai" ? "gpt-5.4" : "gpt-5.4");
  const supportsImages =
    typeof overrides.supportsImages === "boolean"
      ? overrides.supportsImages
      : provider === "openai" ||
        provider === "local" ||
        String(process.env.SEMANTIC_MODEL_SUPPORTS_IMAGES || "").toLowerCase() === "true";

  return { provider, baseUrl, apiKey, model, supportsImages };
}

function isSemanticModelEnabled(config = buildProviderConfig()) {
  if (!config || config.provider === "off") return false;
  if (!config.baseUrl) return false;
  if (config.provider === "local") return true;
  return Boolean(config.apiKey);
}

function extractJsonText(text = "") {
  const value = normalizeText(text);
  if (!value) return "";
  const codeBlock = value.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (codeBlock?.[1]) return codeBlock[1].trim();
  const firstBrace = value.indexOf("{");
  const lastBrace = value.lastIndexOf("}");
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    return value.slice(firstBrace, lastBrace + 1).trim();
  }
  return value;
}

function parseJsonResponse(text = "") {
  const raw = extractJsonText(text);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function buildChatMessages(system, userContent, providerConfig) {
  const messages = [];
  if (system) {
    messages.push({ role: "system", content: system });
  }
  if (providerConfig.supportsImages && Array.isArray(userContent)) {
    messages.push({ role: "user", content: userContent });
    return messages;
  }

  const text = Array.isArray(userContent)
    ? userContent
        .map((item) => {
          if (typeof item === "string") return item;
          if (!item) return "";
          if (item.type === "text") return item.text || "";
          if (item.type === "image_url") {
            return `[图片:${item.image_url?.url ? "已提供" : "缺失"}]`;
          }
          return "";
        })
        .filter(Boolean)
        .join("\n")
    : String(userContent || "");
  messages.push({ role: "user", content: text });
  return messages;
}

function buildLocalChatUrl(baseUrl) {
  const normalized = String(baseUrl || "").replace(/\/v1\/?$/, "");
  return `${normalized || "http://127.0.0.1:11434"}/api/chat`;
}

async function callChatCompletion(providerConfig, system, userContent, options = {}) {
  if (!isSemanticModelEnabled(providerConfig)) return null;
  const messages = buildChatMessages(system, userContent, providerConfig);
  const temperature = Math.max(0.1, Math.min(1, Number(options.temperature || 0.2)));

  if (providerConfig.provider === "local") {
    const url = buildLocalChatUrl(providerConfig.baseUrl);
    const body = {
      model: providerConfig.model,
      messages,
      stream: false,
      format: "json",
      options: {
        temperature,
        num_predict: Math.max(256, Number(options.maxCompletionTokens || options.maxTokens || 1200)),
      },
    };
    if (options.topP != null) {
      body.options.top_p = Math.max(0.01, Math.min(1, Number(options.topP)));
    }
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });
    const text = await response.text();
    if (!response.ok) {
      throw new Error(`语义模型调用失败 (${response.status}): ${clip(text, 240)}`);
    }
    try {
      const json = JSON.parse(text);
      const content = json?.message?.content || json?.choices?.[0]?.message?.content || json?.response || "";
      return normalizeText(Array.isArray(content) ? content.map((item) => item?.text || "").join("") : content);
    } catch {
      return normalizeText(text);
    }
  }

  const baseUrl = String(providerConfig.baseUrl || "").replace(/\/$/, "");
  const url = `${baseUrl}/chat/completions`;
  const body = {
    model: providerConfig.model,
    messages,
    temperature,
  };
  if (providerConfig.provider === "minimax") {
    body.max_completion_tokens = Math.max(256, Number(options.maxCompletionTokens || 1200));
  } else {
    body.max_tokens = Math.max(256, Number(options.maxTokens || 1200));
  }
  if (options.topP != null) {
    body.top_p = Math.max(0.01, Math.min(1, Number(options.topP)));
  }
  if (options.extraBody && typeof options.extraBody === "object") {
    Object.assign(body, options.extraBody);
  }

  const response = await fetch(url, {
    method: "POST",
    headers: {
      ...(providerConfig.apiKey ? { Authorization: `Bearer ${providerConfig.apiKey}` } : {}),
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`语义模型调用失败 (${response.status}): ${clip(text, 240)}`);
  }

  try {
    const json = JSON.parse(text);
    const content = json?.choices?.[0]?.message?.content || json?.choices?.[0]?.text || json?.output_text || "";
    return normalizeText(Array.isArray(content) ? content.map((item) => item?.text || "").join("") : content);
  } catch {
    return normalizeText(text);
  }
}

function summarizeParagraphNodes(section = {}, limit = 2) {
  return (section.paragraphs || [])
    .filter(Boolean)
    .slice(0, limit)
    .map((item) => clip(item, 180));
}

function buildDocPayload({ doc = {}, sections = [], documentSummary = {}, referenceStyleProfile = {}, requestedPages = 0, title = "", department = "", presenter = "" } = {}) {
  return {
    title,
    department,
    presenter,
    requestedPages,
    sourceType: doc.sourceType || documentSummary.sourceType || "",
    counts: documentSummary.counts || {
      paragraphs: (doc.paragraphs || []).length,
      tables: (doc.tables || []).length,
      images: (doc.images || []).length,
    },
    sections: (sections || []).slice(0, 28).map((section) => ({
      index: section.index,
      level: section.level,
      title: clip(section.title || "", 60),
      heading: clip(section.heading || "", 80),
      paragraphCount: (section.paragraphs || []).length,
      text: summarizeParagraphNodes(section, 2),
    })),
    tables: (documentSummary.tables || []).slice(0, 10).map((table) => ({
      index: table.index,
      rows: table.rows,
      columns: table.columns,
      preview: (table.preview || []).slice(0, 3).map((row) => row.slice(0, 5)),
    })),
    images: (documentSummary.images || []).slice(0, 10).map((image) => ({
      index: image.index,
      page: image.page || null,
      path: image.path || "",
      aspectRatio: image.aspectRatio || null,
    })),
    referenceStyle: referenceStyleProfile
      ? {
          count: referenceStyleProfile.count || 0,
          styleFamily: referenceStyleProfile.styleFamily || "",
          densityBias: referenceStyleProfile.densityBias || "",
          suggestedLayoutSet: referenceStyleProfile.suggestedLayoutSet || "",
          templateBias: referenceStyleProfile.templateBias || [],
          summary: clip(referenceStyleProfile.summary || "", 220),
        }
      : {},
  };
}

function buildDocumentSystemPrompt() {
  return [
    "你是银行金融PPT的语义规划助手。",
    "任务：根据需求文档、章节结构、表格和图片，规划一份可编辑 PPT 的页面分组、布局风格和信息密度。",
    "要求：",
    "1. 语义相关的内容要合并到同一页，只有内容过多时再拆分。",
    "2. 避免所有页面都长得像同一种模板，要主动拉开密度差异，并给出页面家族建议。",
    "3. 如果一页表格很多，优先建议高密度表格型、看板型、矩阵型或高亮型布局。",
    "4. 如果一页图多文字少，优先建议图文故事型、图廊型或大图聚焦型布局。",
    "5. 如果是一组步骤、措施、闭环动作，优先建议流程型或行动计划型布局。",
    "6. 如果页面内容较少，优先建议更舒展的留白和更大的文字层级；如果内容较多，优先建议更紧凑但清晰的分区布局。",
    "7. 给出尽量具体的页面蓝图，说明哪些内容要放在同一页，哪些内容要拆到下一页。",
    "8. 如果需要拆分，优先把同一主题拆成相邻页面，但避免相邻页面使用完全相同的版式家族。",
    "9. 输出必须是合法 JSON，不要附加解释，不要附加 Markdown。",
  ].join("\n");
}

function buildDocumentUserPrompt(payload) {
  return JSON.stringify(
    {
      instruction: "请输出一个适合 PPT 生成器使用的语义规划 JSON。",
      expectedKeys: {
        summary: "一句话总结文档核心主题",
        pageCountSuggestion: "2到10之间的建议页数，无法判断可给 null",
        layoutProfile: {
          density: "sparse/balanced/dense/visual/process",
          spacing: "airy/balanced/tight",
          iconVariety: "low/medium/high",
          tablePreference: "split/dense/dashboard/highlight/visual",
          familyBias: ["summary", "split", "dense", "visual", "process", "action"],
          avoidPatterns: ["不要让所有页面都使用同一种左右分栏结构"],
        },
        blocks: [
          {
            title: "页面标题或聚合标题",
            sectionIndexes: [0, 1],
            summary: "页面要点",
            preferredPageType: "summary_cards/table_analysis/process_flow/bullet_columns/image_story/action_plan/key_takeaways",
            density: "sparse/balanced/dense/visual/process",
            keepTogether: true,
            layoutBias: "left_table_right_conclusion/top_summary_bottom_detail/full_width_table/image_first/cards_first",
            pageRole: "summary/process/table/image/action/closing",
            familyHint: "summary/split/dense/visual/process/action/stack",
            contentBalance: "text-heavy/table-heavy/chart-heavy/image-heavy/mixed",
            iconVariety: "low/medium/high",
            spacingBias: "airy/balanced/tight",
            tableKeywords: ["关键词"],
            imageKeywords: ["关键词"],
            needsTable: false,
            needsImage: false,
            mergeHint: "should-merge/should-split",
            splitHint: "split-after/none",
          },
        ],
        styleHints: {
          titleSize: "large/medium/small",
          bodySize: "large/medium/small",
          emphasizeNumbers: true,
          useMoreIconVariety: true,
          useMoreWhitespace: false,
        },
      },
      source: payload,
      pageBlueprints: payload.blocks || [],
      avoidAdjacentFamilies: true,
    },
    null,
    2,
  );
}

function buildDocumentSystemPromptV2() {
  return [
    "你是银行PPT语义规划器。",
    "只输出一个合法 JSON 对象，不要 Markdown，不要代码块，不要解释。",
    "任务：根据需求文档的章节、表格、图片和段落结构，规划适合做 PPT 的页面划分与风格方向。",
    "规则：",
    "1. 相邻页面尽量不要使用同一种布局家族。",
    "2. 内容少则合并，内容多则拆分；同一主题尽量放在同一页，除非内容过多。",
    "3. 如果表格多或列多，优先选择高密度或对比型页面；如果图片多，优先选择图文型页面；如果是步骤/措施，优先选择流程型页面。",
    "4. pageCountSuggestion 只能是 2 到 10 的整数。",
    "5. blocks 最多输出 10 个，每个 block 只保留最关键的 sectionIndexes。",
    "6. 只输出这些字段：summary, pageCountSuggestion, layoutProfile, blocks, styleHints, globalHints。",
  ].join("\n");
}

function buildDocumentUserPromptV2(payload) {
  return [
    "请基于下面的文档内容输出页面规划 JSON。",
    JSON.stringify(payload, null, 2),
  ].join("\n");
}

function normalizeSemanticLayoutProfile(profile = {}, payload = {}) {
  const counts = payload.counts || {};
  const objectProfile = profile && typeof profile === "object" && !Array.isArray(profile) ? profile : {};

  const familyBias = unique([...(Array.isArray(objectProfile.familyBias) ? objectProfile.familyBias : []), ...(Array.isArray(objectProfile.preferredFamilies) ? objectProfile.preferredFamilies : [])]);
  const preferredVariants = unique([
    ...(Array.isArray(objectProfile.preferredVariants) ? objectProfile.preferredVariants : []),
    ...(Array.isArray(objectProfile.tableVariants) ? objectProfile.tableVariants : []),
    ...(Array.isArray(objectProfile.processVariants) ? objectProfile.processVariants : []),
    ...(Array.isArray(objectProfile.summaryVariants) ? objectProfile.summaryVariants : []),
    ...(Array.isArray(objectProfile.imageVariants) ? objectProfile.imageVariants : []),
  ]);
  const preferredPageTypes = unique([
    ...(Array.isArray(objectProfile.preferredPageTypes) ? objectProfile.preferredPageTypes : []),
    ...(Array.isArray(objectProfile.pageTypes) ? objectProfile.pageTypes : []),
  ]);
  const avoidPatterns = unique([...(Array.isArray(objectProfile.avoidPatterns) ? objectProfile.avoidPatterns : [])]);

  let density = String(objectProfile.density || "").toLowerCase();
  let spacing = String(objectProfile.spacing || "").toLowerCase();
  let iconVariety = String(objectProfile.iconVariety || "").toLowerCase();
  let tablePreference = String(objectProfile.tablePreference || "").toLowerCase();
  let source = objectProfile.source || "";

  const addFamily = (...items) => familyBias.push(...items);
  const addVariant = (...items) => preferredVariants.push(...items);
  const addPageType = (...items) => preferredPageTypes.push(...items);

  const items = Array.isArray(profile) ? profile : [];
  items.forEach((item) => {
    const type = String(item?.type || "").toLowerCase();
    const styleType = String(item?.styleType || "").toLowerCase();
    if (type.includes("title")) {
      addFamily("summary", "stack");
      addVariant("spread", "grid", "closing");
      addPageType("summary_cards", "key_takeaways");
    }
    if (type.includes("chart") || type.includes("graph") || styleType.includes("high_density")) {
      density = density || "high";
      tablePreference = tablePreference || "dashboard";
      addFamily("dense", "visual");
      addVariant("dashboard", "dense", "matrix");
      addPageType("table_analysis");
    }
    if (type.includes("text_only")) {
      spacing = spacing || "tight";
      addFamily("bullet", "stack");
      addVariant("staggered", "dual", "triple");
      addPageType("bullet_columns", "summary_cards");
    }
    if (type.includes("table")) {
      tablePreference = tablePreference || "dense";
      addFamily("dense", "split");
      addVariant("dense", "dashboard", "compare", "split");
      addPageType("table_analysis");
    }
    if (type.includes("image")) {
      iconVariety = iconVariety || "high";
      spacing = spacing || "balanced";
      addFamily("visual", "picture");
      addVariant("visual", "storyboard", "gallery", "focus", "split");
      addPageType("image_story");
    }
    if (type.includes("flowchart") || type.includes("process")) {
      spacing = spacing || "balanced";
      addFamily("process", "action");
      addVariant("bridge", "cards", "ladder", "three_lane");
      addPageType("process_flow", "action_plan");
    }
    if (type.includes("comparison")) {
      tablePreference = tablePreference || "compare";
      addFamily("split", "compare");
      addVariant("compare", "split", "sidecallout");
      addPageType("table_analysis");
    }
    if (type.includes("conclusion")) {
      addFamily("stack", "closing");
      addVariant("wall", "closing", "cards");
      addPageType("key_takeaways");
    }
  });

  if (counts.tables >= 3) {
    density = density || "high";
    tablePreference = tablePreference || "dashboard";
    addFamily("dense", "split");
    addVariant("dashboard", "dense", "matrix", "compare", "visual");
    addPageType("table_analysis");
  }
  if (counts.images >= 2) {
    iconVariety = iconVariety || "high";
    spacing = spacing || "balanced";
    addFamily("visual", "picture");
    addVariant("visual", "storyboard", "gallery", "focus", "split");
    addPageType("image_story");
  }
  if (counts.paragraphs >= 18) {
    spacing = spacing || "tight";
    addPageType("summary_cards", "bullet_columns");
  }
  if (counts.sections >= 7) {
    addFamily("summary", "dense", "process");
    addVariant("spread", "dense", "bridge", "cards");
    addPageType("summary_cards", "process_flow");
  }
  if (counts.sections <= 3 && counts.tables <= 1 && counts.images <= 1) {
    spacing = spacing || "airy";
    addVariant("spread", "cards", "closing");
    addPageType("summary_cards", "key_takeaways");
  }

  const normalized = {
    ...objectProfile,
    density: density || (counts.tables >= 3 ? "high" : counts.images >= 2 ? "medium" : "balanced"),
    spacing: spacing || (counts.paragraphs >= 18 ? "tight" : "balanced"),
    iconVariety: iconVariety || (counts.images >= 1 ? "medium" : "low"),
    tablePreference: tablePreference || (counts.tables >= 2 ? "dense" : "split"),
    familyBias: unique(familyBias).slice(0, 10),
    preferredVariants: unique(preferredVariants).slice(0, 12),
    preferredPageTypes: unique(preferredPageTypes).slice(0, 10),
    avoidPatterns,
    source: source || "semantic-array",
  };

  return normalized;
}

async function analyzeDocumentSemantics(context = {}, options = {}) {
  const providerConfig = buildProviderConfig(options);
  if (!isSemanticModelEnabled(providerConfig)) return null;

  const payload = buildDocPayload(context);
  try {
    const text = await callChatCompletion(providerConfig, buildDocumentSystemPromptV2(), buildDocumentUserPromptV2(payload), {
      temperature: 0.05,
      maxCompletionTokens: 900,
      maxTokens: 900,
      extraBody: providerConfig.provider === "minimax" ? { top_p: 0.9 } : undefined,
    });
    const parsed = parseJsonResponse(text);
    if (!parsed) {
      return {
        provider: providerConfig.provider,
        model: providerConfig.model,
        available: true,
        summary: "语义模型已启用，但返回结果无法解析，已回退到规则规划。",
        rawText: text || "",
      };
    }
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: true,
      summary: normalizeText(parsed.summary || parsed.overview || ""),
      pageCountSuggestion: Number(parsed.pageCountSuggestion || parsed.pageCount || 0) || 0,
      layoutProfile: normalizeSemanticLayoutProfile(parsed.layoutProfile || {}, payload),
      blocks: Array.isArray(parsed.blocks) ? parsed.blocks.slice(0, 12) : [],
      styleHints: parsed.styleHints || {},
      globalHints: parsed.globalHints || {},
      raw: parsed,
    };
  } catch (error) {
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: false,
      summary: `语义模型调用失败，已回退到规则规划：${error.message || String(error)}`,
      error: error.message || String(error),
    };
  }
}

function buildStylePrompt(payload) {
  return JSON.stringify(
    {
      instruction: "请分析参考 PPT 图片的视觉风格、布局风格、字体感受和图标风格，并输出 JSON。",
      expectedKeys: {
        styleFamily: "dense-report/visual-report/boardroom-report/balanced-report",
        densityBias: "low/medium/high",
        layoutBias: ["summary_grid_v1", "table_dashboard_v1"],
        iconStyle: ["outline", "solid", "mixed"],
        iconDiversityPolicy: "low/medium/high",
        typography: {
          titleWeight: "bold/extraBold",
          bodyWeight: "regular/medium",
          sizeBias: "large/medium/small",
        },
        spacingBias: "airy/balanced/tight",
        tableStyleBias: ["dense", "split", "dashboard", "highlight", "visual"],
        colorMood: "green-heavy / green+blue / formal / etc",
        repeatedPatterns: ["top metrics", "right conclusions", "table + callout"],
      },
      source: payload,
    },
    null,
    2,
  );
}

function buildReferencePayload(referenceImages = [], referenceStyleProfile = {}) {
  return {
    count: referenceImages.length,
    images: referenceImages.slice(0, 6).map((item) => ({
      path: item.path || "",
      name: item.name || "",
      width: item.width || 0,
      height: item.height || 0,
      aspectRatio: item.aspectRatio || null,
      orientation: item.orientation || "",
      previewDataUrl: item.previewDataUrl || "",
    })),
    heuristicProfile: {
      styleFamily: referenceStyleProfile.styleFamily || "",
      densityBias: referenceStyleProfile.densityBias || "",
      suggestedLayoutSet: referenceStyleProfile.suggestedLayoutSet || "",
      averageAspectRatio: referenceStyleProfile.averageAspectRatio || 0,
      averageSizeKb: referenceStyleProfile.averageSizeKb || 0,
      templateBias: referenceStyleProfile.templateBias || [],
    },
  };
}

function buildImagePromptContent(referenceImages = [], heuristicProfile = {}) {
  const textItems = [
    {
      type: "text",
      text: "请分析这些参考 PPT 图片的布局、字体、卡片感、图标风格、留白和表格密度，输出 JSON，不要解释。",
    },
    {
      type: "text",
      text: JSON.stringify(
        {
          expectedKeys: {
            styleFamily: "dense-report/visual-report/boardroom-report/balanced-report",
            densityBias: "low/medium/high",
            layoutBias: ["summary_grid_v1", "table_dashboard_v1"],
            iconStyle: ["outline", "solid", "mixed"],
            typography: {
              titleWeight: "bold/extraBold",
              bodyWeight: "regular/medium",
              sizeBias: "large/medium/small",
            },
            spacingBias: "airy/balanced/tight",
            tableStyleBias: ["dense", "split", "dashboard", "highlight", "visual"],
            colorMood: "green-heavy / green+blue / formal / etc",
            repeatedPatterns: ["top metrics", "right conclusions", "table + callout"],
          },
          heuristicProfile,
        },
        null,
        2,
      ),
    },
  ];

  referenceImages
    .slice(0, 4)
    .filter((item) => item?.previewDataUrl)
    .forEach((item, index) => {
      textItems.push({
        type: "text",
        text: `参考图${index + 1}：${clip(item.name || item.path || "", 40)}，尺寸 ${item.width || 0}x${item.height || 0}，长宽比 ${item.aspectRatio || 0}`,
      });
      textItems.push({
        type: "image_url",
        image_url: {
          url: item.previewDataUrl,
        },
      });
    });

  return textItems;
}

async function analyzeReferenceStyle(referenceImages = [], referenceStyleProfile = {}, options = {}) {
  const providerConfig = buildProviderConfig(options);
  if (!isSemanticModelEnabled(providerConfig)) return null;
  if (!providerConfig.supportsImages) {
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: true,
      visionSupported: false,
      summary: "当前语义模型仅支持文本，参考图片风格将以本地启发式结果为主。",
    };
  }

  const payload = buildReferencePayload(referenceImages, referenceStyleProfile);
  try {
    const text = await callChatCompletion(providerConfig, buildStylePromptV2(payload), buildImagePromptContentV2(referenceImages, referenceStyleProfile), {
      temperature: 0.05,
      maxCompletionTokens: 800,
      maxTokens: 800,
    });
    const parsed = parseJsonResponse(text);
    if (!parsed) {
      return {
        provider: providerConfig.provider,
        model: providerConfig.model,
        available: true,
        visionSupported: true,
        summary: "参考图片语义分析已执行，但结果无法解析，已保留启发式风格判断。",
        rawText: text || "",
      };
    }
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: true,
      visionSupported: true,
      ...parsed,
      raw: parsed,
    };
  } catch (error) {
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: false,
      visionSupported: true,
      summary: `参考图片语义分析失败：${error.message || String(error)}`,
      error: error.message || String(error),
    };
  }
}

function summarizeSlideForReview(slide = {}) {
  const templateFamily = templateVariantFamily(slide.templateId || "");
  return {
    page: slide.page,
    title: clip(slide.title || "", 48),
    type: slide.type || "",
    density: slide.density || "",
    templateId: slide.templateId || "",
    templateFamily,
    contentMode: slide.contentMode || "",
    metrics: Array.isArray(slide.metrics) ? slide.metrics.length : 0,
    cards: Array.isArray(slide.cards) ? slide.cards.length : 0,
    columns: Array.isArray(slide.columns) ? slide.columns.length : 0,
    insights: Array.isArray(slide.insights) ? slide.insights.length : 0,
    stages: Array.isArray(slide.stages) ? slide.stages.length : 0,
    steps: Array.isArray(slide.steps) ? slide.steps.length : 0,
    takeaways: Array.isArray(slide.takeaways) ? slide.takeaways.length : 0,
    hasTable: Boolean(slide.table?.rowCount || slide.table?.rows?.length),
    tableRowCount: slide.table?.rowCount || slide.table?.rows?.length || 0,
    tableColCount: slide.table?.colCount || slide.table?.header?.length || 0,
    hasImage: Boolean(slide.image?.path),
    hasScreenshots: Boolean(slide.screenshots?.length),
  };
}

function buildReviewPayload(previews = [], outline = {}, style = {}) {
  const slides = (outline.slides || []).map((slide) => summarizeSlideForReview(slide));
  const templateUsage = slides.reduce((acc, slide) => {
    const key = slide.templateId || "unknown";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
  const typeUsage = slides.reduce((acc, slide) => {
    const key = slide.type || "unknown";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
  const densityUsage = slides.reduce((acc, slide) => {
    const key = slide.density || "unknown";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
  const familyUsage = slides.reduce((acc, slide) => {
    const key = slide.templateFamily || "unknown";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  const repeatedPatterns = [];
  Object.entries(templateUsage).forEach(([key, count]) => {
    if (count >= 2) repeatedPatterns.push(`模板 ${key} 使用 ${count} 次`);
  });
  Object.entries(familyUsage).forEach(([key, count]) => {
    if (count >= 2) repeatedPatterns.push(`布局家族 ${key} 使用 ${count} 次`);
  });

  return {
    pages: slides,
    previewCount: previews.length,
    styleHints: {
      family: style.referenceStyleProfile?.styleFamily || "",
      densityBias: style.referenceStyleProfile?.densityBias || "",
      layoutSet: style.layoutLibrary?.setName || "",
      emphasis: style.referenceStyleProfile?.semanticSummary || "",
    },
    pageHints: slides.map((slide) => ({
      page: slide.page,
      title: slide.title,
      type: slide.type,
      layoutBias: slide.layoutBias || "",
      familyHint: slide.templateFamily || "",
      contentMode: slide.contentMode || "",
      hasImage: slide.hasImage,
      hasScreenshots: slide.hasScreenshots,
      tableRowCount: slide.tableRowCount,
      tableColCount: slide.tableColCount,
    })),
    repeatedPatterns,
    typeUsage,
    densityUsage,
    familyUsage,
    templateUsage,
    potentialRiskPages: slides.filter((slide) => slide.metrics >= 4 || slide.tableColCount >= 5 || slide.hasImage || slide.contentMode === "dense").slice(0, 8),
  };
}

function buildPreviewContent(previews = [], outline = {}) {
  return previews.slice(0, 4).flatMap((preview) => [
    {
      type: "text",
      text: `第${preview.page}页：${clip(preview.title || outline.slides?.[preview.page - 1]?.title || "", 48)}`,
    },
    {
      type: "image_url",
      image_url: {
        url: preview.dataUrl || "",
      },
    },
  ]);
}

function buildReviewPrompt(payload) {
  return JSON.stringify(
    {
      instruction:
        "请分析这份 PPT 草稿的生成效果，重点关注页面重复、文字过小、表格过宽、留白过多、图标重复和左右布局失衡。如果 source 里没有图片，请基于结构化布局摘要复核。",
      expectedKeys: {
        overallScore: "0-100",
        summary: "一句话总结",
        issues: [
          {
            page: 3,
            severity: "low/medium/high",
            category: "repeat/table_width/whitespace/text_size/icon_repeat/layout_balance/encoding",
            issue: "问题描述",
            suggestion: "修改建议",
          },
        ],
        pageAdvice: [
          {
            page: 3,
            preferredFamily: "dense/split/visual/process/action/stack",
            preferredVariant: "dashboard/compare/picture/bridge/timeline",
            reason: "为什么这样更合适",
          },
        ],
        repeatedPatterns: ["重复的版式模式"],
        globalSuggestions: ["全局优化建议"],
        refinementHints: {
          densePages: [3, 4],
          spaciousPages: [5],
          imagePages: [4],
          tablePages: [3, 4],
        },
        familyUsage: {
          dense: 2,
          split: 1,
          visual: 1,
        },
      },
      source: payload,
    },
    null,
    2,
  );
}

function buildReviewTextContent(payload) {
  return JSON.stringify(
    {
      instruction: "请基于结构化摘要复核 PPT 效果，指出重复、文字过密、留白过多、表格过宽、图标重复和布局失衡等问题，并给出可执行优化建议。",
      expectedKeys: {
        overallScore: "0-100",
        summary: "一句话总结",
        issues: [
          {
            page: 3,
            severity: "low/medium/high",
            category: "repeat/table_width/whitespace/text_size/icon_repeat/layout_balance/encoding",
            issue: "问题描述",
            suggestion: "修改建议",
          },
        ],
        pageAdvice: [
          {
            page: 3,
            preferredFamily: "dense/split/visual/process/action/stack",
            preferredVariant: "dashboard/compare/picture/bridge/timeline",
            reason: "为什么这样更合适",
          },
        ],
        repeatedPatterns: ["重复的版式模式"],
        globalSuggestions: ["全局优化建议"],
        familyUsage: {
          dense: 2,
          split: 1,
          visual: 1,
        },
      },
      source: payload,
    },
    null,
    2,
  );
}

async function reviewRenderedDeck(previews = [], outline = {}, style = {}, options = {}) {
  const providerConfig = buildProviderConfig(options);
  if (!isSemanticModelEnabled(providerConfig)) return null;

  const payload = buildReviewPayload(previews, outline, style);
  const useImages = providerConfig.supportsImages && previews.some((preview) => preview.dataUrl);

  try {
    const text = await callChatCompletion(
      providerConfig,
      buildReviewPromptV2(payload),
      useImages ? buildPreviewContent(previews, outline, style) : buildReviewTextContentV2(payload),
      {
        temperature: 0.05,
        maxCompletionTokens: 900,
        maxTokens: 900,
        extraBody: providerConfig.provider === "minimax" ? { reasoning_split: true, top_p: 0.9 } : undefined,
      },
    );
    const parsed = parseJsonResponse(text);
    if (!parsed) {
      return {
        provider: providerConfig.provider,
        model: providerConfig.model,
        available: true,
        visionSupported: useImages,
        summary: useImages ? "效果复核已执行，但结果无法解析。" : "文本结构复核已执行，但结果无法解析。",
        rawText: text || "",
      };
    }
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: true,
      visionSupported: useImages,
      ...parsed,
      raw: parsed,
    };
  } catch (error) {
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: false,
      visionSupported: useImages,
      summary: `效果复核失败：${error.message || String(error)}`,
      error: error.message || String(error),
    };
  }
}

function stripThinkBlocks(text = "") {
  return String(text || "")
    .replace(/<think>[\s\S]*?<\/think>/gi, " ")
    .replace(/<thinking>[\s\S]*?<\/thinking>/gi, " ")
    .replace(/<analysis>[\s\S]*?<\/analysis>/gi, " ")
    .trim();
}

function clip(text, maxLength = 160) {
  const value = normalizeText(text);
  if (value.length <= maxLength) return value;
  return `${value.slice(0, Math.max(24, maxLength - 1)).trim()}…`;
}

function buildProviderConfig(overrides = {}) {
  const explicitBaseUrl = String(overrides.baseUrl || process.env.SEMANTIC_MODEL_BASE_URL || "").trim();
  const explicitProvider = String(overrides.provider || process.env.SEMANTIC_MODEL_PROVIDER || "").toLowerCase();
  const provider = String(
    explicitProvider ||
      (explicitBaseUrl ? "local" : "") ||
      (process.env.MINIMAX_API_KEY ? "minimax" : process.env.OPENAI_API_KEY ? "openai" : "off"),
  ).toLowerCase();
  const baseUrl =
    provider === "minimax"
      ? explicitBaseUrl || process.env.MINIMAX_BASE_URL || "https://api.minimaxi.com/v1"
      : provider === "openai"
        ? explicitBaseUrl || "https://api.openai.com/v1"
        : provider === "local"
          ? explicitBaseUrl || "http://127.0.0.1:11434/v1"
          : explicitBaseUrl || "";
  const apiKey =
    overrides.apiKey ||
    process.env.SEMANTIC_MODEL_API_KEY ||
    process.env.MINIMAX_API_KEY ||
    process.env.OPENAI_API_KEY ||
    "";
  const model =
    provider === "minimax"
      ? overrides.model || process.env.SEMANTIC_MODEL_NAME || "MiniMax-M2.7"
      : overrides.model ||
        process.env.SEMANTIC_MODEL_NAME ||
        (provider === "local" ? "deepseek-r1:14b" : "gpt-5.4");
  const supportsImagesEnv = String(process.env.SEMANTIC_MODEL_SUPPORTS_IMAGES || "").toLowerCase() === "true";
  const supportsImages =
    provider === "local" || provider === "minimax"
      ? false
      : typeof overrides.supportsImages === "boolean"
        ? overrides.supportsImages
        : provider === "openai"
          ? true
          : provider === "custom"
            ? supportsImagesEnv
            : false;

  return { provider, baseUrl, apiKey, model, supportsImages };
}

function isSemanticModelEnabled(config = buildProviderConfig()) {
  if (!config || config.provider === "off") return false;
  if (!config.baseUrl) return false;
  if (config.provider === "local") return true;
  return Boolean(config.apiKey);
}

function extractJsonText(text = "") {
  const value = normalizeText(stripThinkBlocks(text));
  if (!value) return "";
  const codeBlock = value.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (codeBlock?.[1]) return codeBlock[1].trim();
  const firstBrace = value.indexOf("{");
  const lastBrace = value.lastIndexOf("}");
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    return value.slice(firstBrace, lastBrace + 1).trim();
  }
  return value;
}

function parseJsonResponse(text = "") {
  const raw = extractJsonText(text);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function buildChatMessages(system, userContent, providerConfig) {
  const messages = [];
  if (system) {
    messages.push({ role: "system", content: system });
  }
  if (providerConfig.supportsImages && Array.isArray(userContent)) {
    messages.push({ role: "user", content: userContent });
    return messages;
  }

  const text = Array.isArray(userContent)
    ? userContent
        .map((item, index) => {
          if (typeof item === "string") return item;
          if (!item) return "";
          if (item.type === "text") return item.text || "";
          if (item.type === "image_url") {
            return `[图片${index + 1}：${item.image_url?.url ? "已省略" : "缺失"}]`;
          }
          return "";
        })
        .filter(Boolean)
        .join("\n")
    : String(userContent || "");
  messages.push({ role: "user", content: text });
  return messages;
}

function buildOpenAIChatUrl(baseUrl) {
  const normalized = String(baseUrl || "").trim().replace(/\/+$/, "");
  if (!normalized) return "https://api.openai.com/v1/chat/completions";
  if (/\/chat\/completions$/i.test(normalized)) return normalized;
  return `${normalized.replace(/\/v1$/i, "")}/v1/chat/completions`;
}

function buildOllamaChatUrl(baseUrl) {
  const normalized = String(baseUrl || "").trim().replace(/\/+$/, "");
  if (!normalized) return "http://127.0.0.1:11434/api/chat";
  const base = normalized.replace(/\/v1$/i, "");
  return `${base.replace(/\/api\/chat$/i, "")}/api/chat`;
}

function buildChatHeaders(providerConfig = {}) {
  const headers = {
    "Content-Type": "application/json",
  };
  if (!providerConfig.apiKey) return headers;
  if (providerConfig.provider === "minimax") {
    headers.Authorization = `Bearer ${providerConfig.apiKey}`;
    headers["api-key"] = providerConfig.apiKey;
    headers["x-api-key"] = providerConfig.apiKey;
  } else {
    headers.Authorization = `Bearer ${providerConfig.apiKey}`;
  }
  return headers;
}

async function postChatCompletion(url, body, headers = {}) {
  const response = await fetch(url, {
    method: "POST",
    headers: {
      ...headers,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`语义模型调用失败 (${response.status}): ${clip(text, 240)}`);
  }
  return text;
}

function extractChatContent(text = "") {
  try {
    const json = JSON.parse(text);
    const content = json?.choices?.[0]?.message?.content || json?.choices?.[0]?.text || json?.message?.content || json?.response || json?.output_text || "";
    return normalizeText(
      Array.isArray(content)
        ? content
            .map((item) => {
              if (typeof item === "string") return item;
              if (!item) return "";
              return item.text || item.content || "";
            })
            .join("")
        : content,
    );
  } catch {
    return normalizeText(stripThinkBlocks(text));
  }
}

async function callChatCompletion(providerConfig, system, userContent, options = {}) {
  if (!isSemanticModelEnabled(providerConfig)) return null;
  const messages = buildChatMessages(system, userContent, providerConfig);
  const temperature = Math.max(0.1, Math.min(1, Number(options.temperature || 0.2)));
  const maxTokens = Math.max(256, Number(options.maxCompletionTokens || options.maxTokens || 1200));
  const baseBody = {
    model: providerConfig.model,
    messages,
    temperature,
    max_tokens: maxTokens,
  };
  if (options.topP != null) {
    baseBody.top_p = Math.max(0.01, Math.min(1, Number(options.topP)));
  }
  if (options.extraBody && typeof options.extraBody === "object") {
    Object.assign(baseBody, options.extraBody);
  }

  if (providerConfig.provider === "local") {
    try {
      const text = await postChatCompletion(buildOpenAIChatUrl(providerConfig.baseUrl), baseBody, buildChatHeaders(providerConfig));
      return extractChatContent(text);
    } catch (error) {
      const ollamaBody = {
        model: providerConfig.model,
        messages,
        stream: false,
        format: "json",
        options: {
          temperature,
          num_predict: maxTokens,
        },
      };
      if (options.topP != null) {
        ollamaBody.options.top_p = Math.max(0.01, Math.min(1, Number(options.topP)));
      }
      if (options.extraBody && typeof options.extraBody === "object") {
        Object.assign(ollamaBody.options, options.extraBody);
      }
      const text = await postChatCompletion(buildOllamaChatUrl(providerConfig.baseUrl), ollamaBody, { "Content-Type": "application/json" });
      return extractChatContent(text);
    }
  }

  const text = await postChatCompletion(buildOpenAIChatUrl(providerConfig.baseUrl), baseBody, buildChatHeaders(providerConfig));
  return extractChatContent(text);
}

function buildChatMessages(system, userContent, providerConfig) {
  const messages = [];
  if (system) {
    messages.push({ role: "system", content: system });
  }
  if (providerConfig.supportsImages && Array.isArray(userContent)) {
    messages.push({ role: "user", content: userContent });
    return messages;
  }

  const text = Array.isArray(userContent)
    ? userContent
        .map((item, index) => {
          if (typeof item === "string") return item;
          if (!item) return "";
          if (item.type === "text") return item.text || "";
          if (item.type === "image_url") {
            return `[图片${index + 1}：${item.image_url?.url ? "已提供" : "缺失"}]`;
          }
          return "";
        })
        .filter(Boolean)
        .join("\n")
    : String(userContent || "");
  messages.push({ role: "user", content: text });
  return messages;
}

function buildChatHeaders(providerConfig = {}) {
  const headers = {
    "Content-Type": "application/json",
  };
  if (providerConfig.apiKey) {
    headers.Authorization = `Bearer ${providerConfig.apiKey}`;
  }
  return headers;
}

function buildMinimaxChatUrls(baseUrl) {
  const normalized = String(baseUrl || "").trim().replace(/\/+$/, "");
  const primary = buildOpenAIChatUrl(normalized || "https://api.minimaxi.com/v1");
  const alternateBase = /minimaxi\.com/i.test(normalized) ? "https://api.minimax.io/v1" : "https://api.minimaxi.com/v1";
  const alternate = buildOpenAIChatUrl(alternateBase);
  return [...new Set([primary, alternate].filter(Boolean))];
}

async function callChatCompletion(providerConfig, system, userContent, options = {}) {
  if (!isSemanticModelEnabled(providerConfig)) return null;
  const messages = buildChatMessages(system, userContent, providerConfig);
  const temperature = Math.max(0.1, Math.min(1, Number(options.temperature || 0.2)));
  const maxTokens = Math.max(256, Number(options.maxCompletionTokens || options.maxTokens || 1200));
  const baseBody = {
    model: providerConfig.model,
    messages,
    temperature,
    max_tokens: maxTokens,
  };
  if (options.topP != null) {
    baseBody.top_p = Math.max(0.01, Math.min(1, Number(options.topP)));
  }
  if (options.extraBody && typeof options.extraBody === "object") {
    Object.assign(baseBody, options.extraBody);
  }

  if (providerConfig.provider === "local") {
    try {
      const text = await postChatCompletion(buildOpenAIChatUrl(providerConfig.baseUrl), baseBody, buildChatHeaders(providerConfig));
      return extractChatContent(text);
    } catch (error) {
      const ollamaBody = {
        model: providerConfig.model,
        messages,
        stream: false,
        format: "json",
        options: {
          temperature,
          num_predict: maxTokens,
        },
      };
      if (options.topP != null) {
        ollamaBody.options.top_p = Math.max(0.01, Math.min(1, Number(options.topP)));
      }
      if (options.extraBody && typeof options.extraBody === "object") {
        Object.assign(ollamaBody.options, options.extraBody);
      }
      const text = await postChatCompletion(buildOllamaChatUrl(providerConfig.baseUrl), ollamaBody, { "Content-Type": "application/json" });
      return extractChatContent(text);
    }
  }

  const candidateUrls = providerConfig.provider === "minimax" ? buildMinimaxChatUrls(providerConfig.baseUrl) : [buildOpenAIChatUrl(providerConfig.baseUrl)];
  let lastError = null;
  for (const url of candidateUrls) {
    try {
      const text = await postChatCompletion(url, baseBody, buildChatHeaders(providerConfig));
      return extractChatContent(text);
    } catch (error) {
      lastError = error;
    }
  }

  if (lastError) throw lastError;
  return null;
}

function clip(text, maxLength = 160) {
  const value = normalizeText(text);
  if (value.length <= maxLength) return value;
  return `${value.slice(0, Math.max(24, maxLength - 1)).trim()}…`;
}

async function postChatCompletion(url, body, headers = {}) {
  const response = await fetch(url, {
    method: "POST",
    headers: {
      ...headers,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`语义模型调用失败 (${response.status}): ${clip(text, 240)}`);
  }
  return text;
}

module.exports = {
  analyzeDocumentSemantics,
  analyzeReferenceStyle,
  buildProviderConfig,
  isSemanticModelEnabled,
  reviewRenderedDeck,
  parseJsonResponse,
};
