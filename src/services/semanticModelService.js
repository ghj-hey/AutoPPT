const { fileToDataUrl } = require("./referenceStyleService");
const { templateVariantFamily } = require("./layoutSelectionService");

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
    explicitBaseUrl ||
    (provider === "minimax"
      ? "https://api.minimax.io/v1"
      : provider === "openai"
        ? "https://api.openai.com/v1"
        : provider === "local"
          ? "http://127.0.0.1:11434/v1"
        : process.env.SEMANTIC_MODEL_BASE_URL || "");
  const apiKey =
    overrides.apiKey ||
    process.env.SEMANTIC_MODEL_API_KEY ||
    process.env.MINIMAX_API_KEY ||
    process.env.OPENAI_API_KEY ||
    "";
  const model =
    overrides.model ||
    process.env.SEMANTIC_MODEL_NAME ||
    (provider === "minimax" ? "MiniMax-M2.7" : "gpt-5.4");
  const supportsImages =
    typeof overrides.supportsImages === "boolean"
      ? overrides.supportsImages
      : provider === "openai" ||
        provider === "local" ||
        String(process.env.SEMANTIC_MODEL_SUPPORTS_IMAGES || "").toLowerCase() === "true";

  return {
    provider,
    baseUrl,
    apiKey,
    model,
    supportsImages,
  };
}

function isSemanticModelEnabled(config = buildProviderConfig()) {
  if (!config || config.provider === "off") return false;
  if (!config.baseUrl) return false;
  if (config.provider === "local") return true;
  return Boolean(config.apiKey);
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

function buildChatMessages(system, userContent, providerConfig, options = {}) {
  const messages = [];
  if (system) {
    messages.push({ role: "system", content: system });
  }
  if (providerConfig.supportsImages && Array.isArray(userContent)) {
    messages.push({
      role: "user",
      content: userContent,
    });
  } else {
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
  }
  return messages;
}

async function callChatCompletion(providerConfig, system, userContent, options = {}) {
  if (!isSemanticModelEnabled(providerConfig)) return null;
  const baseUrl = String(providerConfig.baseUrl || "").replace(/\/$/, "");
  const url = `${baseUrl}/chat/completions`;
  const messages = buildChatMessages(system, userContent, providerConfig, options);
  const body = {
    model: providerConfig.model,
    messages,
    temperature: Math.max(0.1, Math.min(1, Number(options.temperature || 0.2))),
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

function buildDocumentSystemPrompt() {
  return [
    "你是银行金融PPT的语义规划助手。",
    "任务：根据需求文档、章节结构、表格和图片，规划一份可编辑PPT的页面分组与版式方向。",
    "要求：",
    "1. 把语义相关的内容合并到同一页，除非内容过多才拆分。",
    "2. 避免生成每页都相似的结构，主动拉开密度差异，并给出页面家族差异建议。",
    "3. 如果一页里表格很多，优先给出高密度表格型、看板型、矩阵型或高亮型建议。",
    "4. 如果一页图多文字少，优先给出图文故事型、图廊型或大图聚焦型建议。",
    "5. 如果是一组步骤、措施、闭环动作，优先给出流程型或行动计划型建议。",
    "6. 如果页面内容较少，优先建议更舒展的留白和更大的文字层级；如果内容较多，优先建议更密集但清晰的分区布局。",
    "6. 输出必须是合法 JSON，不要加解释，不要加 Markdown。",
  ].join("\n");
}

function buildDocumentUserPrompt(payload) {
  return JSON.stringify(
    {
      instruction: "请输出一个适合PPT生成器使用的语义规划JSON。",
      expectedKeys: {
        summary: "一句话总结文档核心主题",
        pageCountSuggestion: "2到10之间的建议页数，无法判断可给null",
        layoutProfile: {
          density: "sparse/balanced/dense/visual/process",
          spacing: "airy/balanced/tight",
          iconVariety: "low/medium/high",
          tablePreference: "split/dense/dashboard/highlight/visual",
          familyBias: ["summary", "split", "dense", "visual", "process", "action"],
          avoidPatterns: ["不要让所有页面都用同一个左右分栏结构"],
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
            tableKeywords: ["关键词"],
            imageKeywords: ["关键词"],
            needsTable: false,
            needsImage: false,
            contentBalance: "text-heavy/chart-heavy/image-heavy/mixed",
            familyHint: "summary/split/dense/visual/process/action",
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
    },
    null,
    2,
  );
}

async function analyzeDocumentSemantics(context = {}, options = {}) {
  const providerConfig = buildProviderConfig(options);
  if (!isSemanticModelEnabled(providerConfig)) return null;

  const payload = buildDocPayload(context);
  const system = buildDocumentSystemPrompt();
  const user = buildDocumentUserPrompt(payload);

  try {
    const text = await callChatCompletion(providerConfig, system, user, {
      temperature: 0.2,
      maxCompletionTokens: 1400,
      maxTokens: 1400,
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
      layoutProfile: parsed.layoutProfile || {},
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
      instruction: "请分析参考PPT图片的视觉风格、布局风格、字体感受和图标风格，并输出JSON。",
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
    { type: "text", text: "请分析这些参考PPT图片的布局、字体、卡片感、图标风格、留白和表格密度，输出JSON，不要解释。" },
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
        text: `参考图${index + 1}: ${clip(item.name || item.path || "", 40)}，${item.width || 0}x${item.height || 0}，纵横比 ${item.aspectRatio || 0}`,
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
      summary: "当前语义模型仅支持文本，参考图片风格仍以本地启发式结果为主。",
    };
  }

  const payload = buildReferencePayload(referenceImages, referenceStyleProfile);
  const system = buildStylePrompt(payload);
  const userContent = buildImagePromptContent(referenceImages, referenceStyleProfile);

  try {
    const text = await callChatCompletion(providerConfig, system, userContent, {
      temperature: 0.2,
      maxCompletionTokens: 1000,
      maxTokens: 1000,
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

function buildReviewPrompt(payload) {
  return JSON.stringify(
    {
      instruction:
        "请分析这份PPT草稿的生成效果，重点关注页面重复、文字过小、表格过宽、留白过多、图标重复和左右布局失衡。如果 source 里没有图片，请基于结构化布局摘要评审。",
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
    repeatedPatterns,
    typeUsage,
    densityUsage,
    familyUsage,
    templateUsage,
    potentialRiskPages: slides.filter((slide) => slide.metrics >= 4 || slide.tableColCount >= 5 || slide.hasImage || slide.contentMode === "dense").slice(0, 8),
  };
}

function buildPreviewContent(previews = [], outline = {}, style = {}) {
  return previews.slice(0, 4).flatMap((preview) => [
    {
      type: "text",
      text: `第${preview.page}页：${clip(preview.title || outline.slides?.[preview.page - 1]?.title || "", 48)}。`,
    },
    {
      type: "image_url",
      image_url: {
        url: preview.dataUrl || "",
      },
    },
  ]);
}

function buildReviewTextContent(payload) {
  return JSON.stringify(
    {
      instruction: "请基于结构化摘要评审PPT效果，指出重复、文字过密、留白过多、表格过宽、图标重复、布局失衡等问题，并给出可执行优化建议。",
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
      buildReviewPrompt(payload),
      useImages ? buildPreviewContent(previews, outline, style) : buildReviewTextContent(payload),
      {
        temperature: 0.2,
        maxCompletionTokens: 1200,
        maxTokens: 1200,
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
        summary: useImages ? "效果评审已执行，但结果无法解析。" : "文本结构评审已执行，但结果无法解析。",
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
      summary: `效果评审失败：${error.message || String(error)}`,
      error: error.message || String(error),
    };
  }
}

module.exports = {
  analyzeDocumentSemantics,
  analyzeReferenceStyle,
  buildProviderConfig,
  isSemanticModelEnabled,
  reviewRenderedDeck,
  parseJsonResponse,
};
