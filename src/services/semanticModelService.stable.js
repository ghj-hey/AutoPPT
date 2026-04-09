const { fileToDataUrl } = require("./referenceStyleService");

function unique(list) {
  return [...new Set((list || []).filter(Boolean))];
}

function normalizeApiKey(value = "") {
  return String(value || "").trim().replace(/^Bearer\s+/i, "");
}

function normalizeText(text) {
  return String(text || "")
    .replace(/<think>[\s\S]*?<\/think>/gi, " ")
    .replace(/<thinking>[\s\S]*?<\/thinking>/gi, " ")
    .replace(/<analysis>[\s\S]*?<\/analysis>/gi, " ")
    .replace(/\s+/g, " ")
    .replace(/[\u0000-\u001f]+/g, " ")
    .trim();
}

function escapeSuspiciousJsonQuotes(text = "") {
  const value = String(text || "");
  if (!value) return "";

  let result = "";
  let inString = false;
  let escaped = false;

  const nextNonWhitespace = (index) => {
    for (let i = index; i < value.length; i += 1) {
      if (!/\s/.test(value[i])) return value[i];
    }
    return "";
  };

  for (let i = 0; i < value.length; i += 1) {
    const char = value[i];
    result += char;

    if (escaped) {
      escaped = false;
      continue;
    }
    if (char === "\\") {
      if (inString) escaped = true;
      continue;
    }
    if (char !== '"') {
      continue;
    }

    if (!inString) {
      inString = true;
      continue;
    }

    const lookahead = nextNonWhitespace(i + 1);
    if (!lookahead || lookahead === "," || lookahead === "}" || lookahead === "]") {
      inString = false;
      continue;
    }

    result = result.slice(0, -1) + '\\"';
  }

  return result;
}

function clip(text, maxLength = 160) {
  const value = normalizeText(text);
  if (value.length <= maxLength) return value;
  return `${value.slice(0, Math.max(24, maxLength - 1)).trim()}...`;
}

function buildProviderConfig(overrides = {}) {
  const explicitProvider = String(overrides.provider || process.env.SEMANTIC_MODEL_PROVIDER || "").toLowerCase();
  const explicitBaseUrl = String(overrides.baseUrl || process.env.SEMANTIC_MODEL_BASE_URL || "").trim();
  const envModel = String(process.env.SEMANTIC_MODEL_NAME || "").trim();
  const semanticKey = normalizeApiKey(overrides.apiKey || process.env.SEMANTIC_MODEL_API_KEY || "");
  const provider = String(
    explicitProvider ||
      (explicitBaseUrl ? "local" : "") ||
      (semanticKey ? "minimax" : "") ||
      (process.env.MINIMAX_API_KEY ? "minimax" : process.env.OPENAI_API_KEY ? "openai" : "off"),
  ).toLowerCase();

  const baseUrl =
    provider === "minimax"
      ? explicitBaseUrl || process.env.MINIMAX_BASE_URL || "https://api.minimaxi.com/v1"
      : provider === "openai"
        ? explicitBaseUrl || process.env.OPENAI_BASE_URL || "https://api.openai.com/v1"
        : provider === "local"
          ? explicitBaseUrl || "http://127.0.0.1:11434/v1"
          : explicitBaseUrl || process.env.SEMANTIC_MODEL_BASE_URL || "";

  const apiKeyByProvider =
    provider === "minimax"
      ? semanticKey || String(process.env.MINIMAX_API_KEY || "").trim()
      : provider === "openai"
        ? semanticKey || String(process.env.OPENAI_API_KEY || "").trim()
        : provider === "local"
          ? semanticKey
          : semanticKey || String(process.env.MINIMAX_API_KEY || "").trim() || String(process.env.OPENAI_API_KEY || "").trim();

  const model =
    overrides.model ||
    (provider === "local"
      ? "deepseek-r1:14b"
      : provider === "minimax"
        ? "MiniMax-M2.7"
        : provider === "openai"
          ? String(process.env.OPENAI_MODEL_NAME || envModel || "gpt-5.4").trim()
          : envModel || String(process.env.OPENAI_MODEL_NAME || "gpt-5.4").trim());

  const supportsImages =
    typeof overrides.supportsImages === "boolean"
      ? overrides.supportsImages
      : provider === "openai" ||
        provider === "minimax" ||
        String(process.env.SEMANTIC_MODEL_SUPPORTS_IMAGES || "").toLowerCase() === "true";

  return { provider, baseUrl, apiKey: apiKeyByProvider, model, supportsImages };
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

function buildDocPayload(
  { doc = {}, sections = [], documentSummary = {}, referenceStyleProfile = {}, requestedPages = 0, title = "", department = "", presenter = "" } = {},
  options = {},
) {
  const compact = Boolean(options.compactPayload);
  const sectionLimit = compact ? 8 : 14;
  const markdownLimit = compact ? 240 : 720;
  const paragraphLimit = compact ? 1 : 2;
  const tablePreviewLimit = compact ? 2 : 3;
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
      sections: sections.length,
    },
     sections: (sections || []).slice(0, sectionLimit).map((section) => ({
       index: section.index,
       level: section.level,
       title: clip(section.title || "", 60),
       heading: clip(section.heading || "", 80),
       paragraphCount: (section.paragraphs || []).length,
       subsectionCount: Number(section.meta?.subsectionCount || section.allHeadings?.length || 0),
       tableCount: Number(section.tables?.length || 0),
       imageCount: Number(section.images?.length || 0),
       markdown: clip(section.markdown || "", markdownLimit),
       text: summarizeParagraphNodes(section, paragraphLimit),
     })),
    tables: (documentSummary.tables || []).slice(0, compact ? 4 : 6).map((table) => ({
      index: table.index,
      rows: table.rows,
      columns: table.columns,
      preview: (table.preview || []).slice(0, tablePreviewLimit).map((row) => row.slice(0, compact ? 4 : 5)),
    })),
    images: (documentSummary.images || []).slice(0, compact ? 4 : 6).map((image) => ({
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

function buildSectionPayload(section = {}, options = {}) {
  return {
    sectionId: section.id || "",
    title: section.heading || section.title || "",
    slideCap: Math.max(1, Math.min(2, Number(options.slideCap || 2))),
    paragraphCount: Number(section.meta?.paragraphCount || section.allParagraphs?.length || 0),
    charCount: Number(section.meta?.charCount || 0),
    subsectionCount: Number(section.meta?.subsectionCount || section.allHeadings?.length || 0),
    tableCount: Number(section.tables?.length || 0),
    imageCount: Number(section.images?.length || 0),
    headings: (section.allHeadings || []).slice(0, 12),
    markdown: clip(section.markdown || "", 3200),
    tables: (section.tables || []).slice(0, 4).map((table, index) => ({
      index: index + 1,
      rows: (table.rows || []).length,
      columns: Math.max(...(table.rows || []).map((row) => row.length), 0),
      preview: (table.rows || []).slice(0, 3).map((row) => row.slice(0, 5)),
    })),
    images: (section.images || []).slice(0, 4).map((image, index) => ({
      index: index + 1,
      page: image.page || null,
      aspectRatio: image.aspectRatio || null,
      path: image.path || "",
    })),
  };
}

function buildStylePayload(referenceImages = [], referenceStyleProfile = {}) {
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
      averageMegapixels: referenceStyleProfile.averageMegapixels || 0,
      highResolutionRatio: referenceStyleProfile.highResolutionRatio || 0,
      headerStyle: referenceStyleProfile.headerStyle || "",
      summaryBandStyle: referenceStyleProfile.summaryBandStyle || "",
      tablePreference: referenceStyleProfile.tablePreference || "",
      cardStyle: referenceStyleProfile.cardStyle || "",
      pageRhythm: referenceStyleProfile.pageRhythm || "",
      imagePlacement: referenceStyleProfile.imagePlacement || "",
      iconDiversity: referenceStyleProfile.iconDiversity || "",
      preferredVariants: referenceStyleProfile.preferredVariants || [],
      tableStyleBias: referenceStyleProfile.tableStyleBias || [],
      templateBias: referenceStyleProfile.templateBias || [],
    },
  };
}

function buildSectionPrompt(payload = {}) {
  return [
    "You are a PowerPoint planner for structured Chinese business documents.",
    "Return one strict JSON object only. No markdown. No explanations. Do not echo the input.",
    "Task: summarize one top-level section and convert it into 1 or 2 PPT page plans.",
    "Rules:",
    "1. Use the section markdown and heading tree as the primary evidence.",
    "2. Keep 1 / 1.1 / 1.1.1 style related subsections together unless the content is clearly too dense for one slide.",
    "3. If the section contains important tables, keep the table and its explanation together when possible.",
    "4. If the section contains screenshots or UI-like images, prefer a visual or split layout.",
    "5. slideCountSuggestion must be 1 or 2 only.",
    "6. pagePlans length must equal slideCountSuggestion.",
    "7. Each pagePlan must contain only these keys: title, pageRole, density, layoutBias, keyPoints, focusHeadings, mustUseTable, mustUseImage.",
    "8. pageRole must be one of: table, process, image, action, closing, bullet, mixed.",
    "9. density must be one of: low, medium, high.",
    "10. layoutBias should use reusable keywords such as dense, dashboard, matrix, compare, split, visual, picture, focus, bridge, cards, timeline, wall.",
    "11. Use the input JSON as evidence; do not return the input schema itself.",
  ].join("\n");
}

function buildReviewPayload(previews = [], outline = {}, style = {}) {
  return {
    previewCount: previews.length,
    pages: (outline.slides || []).map((slide) => ({
      page: slide.page,
      title: slide.title || "",
      type: slide.type || "",
      density: slide.density || "",
      templateId: slide.templateId || "",
      tableRowCount: slide.table?.rowCount || slide.table?.rows?.length || 0,
      tableColCount: slide.table?.colCount || slide.table?.header?.length || 0,
      hasImage: Boolean(slide.image?.path),
      hasScreenshots: Boolean(slide.screenshots?.length),
    })),
    riskPages: (outline.slides || [])
      .filter((slide) => slide.type === "table_analysis" || slide.type === "image_story")
      .slice(0, 8)
      .map((slide) => ({
        page: slide.page,
        title: slide.title || "",
        type: slide.type || "",
        density: slide.density || "",
        tableRowCount: slide.table?.rowCount || slide.table?.rows?.length || 0,
        tableColCount: slide.table?.colCount || slide.table?.header?.length || 0,
      })),
    repeatedPatterns: outline.slides?.length
      ? unique(
          outline.slides
            .map((slide) => `${slide.type || ""}:${slide.templateId || ""}`)
            .filter(Boolean)
            .filter((item, index, array) => array.indexOf(item) !== index),
        )
      : [],
    styleHints: {
      family: style.referenceStyleProfile?.styleFamily || "",
      density: style.referenceStyleProfile?.densityBias || "",
      colorMood: style.colorMood || "",
      iconStyle: style.iconStyle || "",
    },
  };
}

function stripThinkBlocks(text = "") {
  return String(text || "")
    .replace(/<think>[\s\S]*?<\/think>/gi, " ")
    .replace(/<thinking>[\s\S]*?<\/thinking>/gi, " ")
    .replace(/<analysis>[\s\S]*?<\/analysis>/gi, " ")
    .trim();
}

function extractJsonCandidate(text = "") {
  const value = normalizeText(stripThinkBlocks(text));
  if (!value) return "";
  const codeBlock = value.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (codeBlock?.[1]) return codeBlock[1].trim();

  const startIndex = value.search(/[\[{]/);
  if (startIndex >= 0) {
    const openChar = value[startIndex];
    const closeChar = openChar === "{" ? "}" : "]";
    let depth = 0;
    let inString = false;
    let escape = false;
    for (let index = startIndex; index < value.length; index += 1) {
      const char = value[index];
      if (inString) {
        if (escape) {
          escape = false;
          continue;
        }
        if (char === "\\") {
          escape = true;
          continue;
        }
        if (char === '"') {
          inString = false;
        }
        continue;
      }

      if (char === '"') {
        inString = true;
        continue;
      }
      if (char === openChar) depth += 1;
      if (char === closeChar) {
        depth -= 1;
        if (depth === 0) {
          return value.slice(startIndex, index + 1).trim();
        }
      }
    }
  }

  const firstBrace = value.indexOf("{");
  const lastBrace = value.lastIndexOf("}");
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    return value.slice(firstBrace, lastBrace + 1).trim();
  }
  const firstBracket = value.indexOf("[");
  const lastBracket = value.lastIndexOf("]");
  if (firstBracket >= 0 && lastBracket > firstBracket) {
    return value.slice(firstBracket, lastBracket + 1).trim();
  }
  return value;
}

function repairJsonCandidate(text = "") {
  return String(text || "")
    .replace(/^[^\[{]*/, "")
    .replace(/[^\]}]*$/, "")
    .replace(/([{,]\s*)([A-Za-z_][A-Za-z0-9_-]*)(\s*:)/g, '$1"$2"$3')
    .replace(/,\s*([}\]])/g, '$1')
    .replace(/},\s*"\{/g, '},{')
    .replace(/}\s*"\{/g, '},{')
    .replace(/]\s*"\{/g, '],{');
}

function closeUnbalancedJsonCandidate(text = "") {
  const source = String(text || "");
  if (!source) return "";

  let result = "";
  let inString = false;
  let escape = false;
  const stack = [];

  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    result += char;

    if (inString) {
      if (escape) {
        escape = false;
        continue;
      }
      if (char === "\\") {
        escape = true;
        continue;
      }
      if (char === '"') {
        inString = false;
      }
      continue;
    }

    if (char === '"') {
      inString = true;
      continue;
    }
    if (char === "{" || char === "[") {
      stack.push(char);
      continue;
    }
    if (char === "}" || char === "]") {
      const expected = stack[stack.length - 1];
      if ((char === "}" && expected === "{") || (char === "]" && expected === "[")) {
        stack.pop();
      }
    }
  }

  if (inString) {
    result += '"';
  }
  while (stack.length) {
    const open = stack.pop();
    result += open === "{" ? "}" : "]";
  }
  return result;
}

function extractBalancedJsonCandidate(text = "") {
  const value = normalizeText(stripThinkBlocks(text));
  if (!value) return "";

  const candidates = [];
  const firstObject = value.indexOf("{");
  const lastObject = value.lastIndexOf("}");
  if (firstObject >= 0 && lastObject > firstObject) {
    candidates.push(value.slice(firstObject, lastObject + 1));
  }

  const firstArray = value.indexOf("[");
  const lastArray = value.lastIndexOf("]");
  if (firstArray >= 0 && lastArray > firstArray) {
    candidates.push(value.slice(firstArray, lastArray + 1));
  }

  return candidates.find(Boolean) || "";
}

function parseJsonResponse(text = "") {
  const raw = extractJsonCandidate(text);
  if (!raw) return null;
  const bounded = extractBalancedJsonCandidate(raw);
  const candidates = [
    raw,
    bounded,
    escapeSuspiciousJsonQuotes(raw),
    escapeSuspiciousJsonQuotes(bounded),
    repairJsonCandidate(raw),
    repairJsonCandidate(bounded),
    closeUnbalancedJsonCandidate(raw),
    closeUnbalancedJsonCandidate(bounded),
    closeUnbalancedJsonCandidate(repairJsonCandidate(raw)),
    closeUnbalancedJsonCandidate(repairJsonCandidate(bounded)),
    repairJsonCandidate(stripThinkBlocks(raw)),
    closeUnbalancedJsonCandidate(escapeSuspiciousJsonQuotes(raw)),
    closeUnbalancedJsonCandidate(escapeSuspiciousJsonQuotes(bounded)),
  ];
  for (const candidate of candidates) {
    try {
      return JSON.parse(candidate);
    } catch {
      // try next candidate
    }
  }
  return null;
}

function buildChatMessages(system, userContent, providerConfig) {
  const messages = [];
  if (system) messages.push({ role: "system", content: system });
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
            return `[image-${index + 1}:${item.image_url?.url ? "available" : "missing"}]`;
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
  if (/\/v1$/i.test(normalized)) return `${normalized}/chat/completions`;
  return `${normalized}/v1/chat/completions`;
}

function buildOllamaChatUrl(baseUrl) {
  const normalized = String(baseUrl || "").trim().replace(/\/+$/, "");
  if (!normalized) return "http://127.0.0.1:11434/api/chat";
  if (/\/api\/chat$/i.test(normalized)) return normalized;
  const base = normalized.replace(/\/v1$/i, "");
  return `${base}/api/chat`;
}

function buildMinimaxTextChatUrl(baseUrl) {
  const normalized = String(baseUrl || "").trim().replace(/\/+$/, "");
  if (!normalized) return "https://api.minimaxi.com/v1/text/chatcompletion_v2";
  if (/\/text\/chatcompletion_v2$/i.test(normalized)) return normalized;
  if (/\/v1\/text\/chatcompletion_v2$/i.test(normalized)) return normalized;
  if (/\/v1$/i.test(normalized)) return `${normalized}/text/chatcompletion_v2`;
  if (/\/text$/i.test(normalized)) return `${normalized}/chatcompletion_v2`;
  return `${normalized}/v1/text/chatcompletion_v2`;
}

function buildMinimaxChatUrls(baseUrl) {
  const normalized = String(baseUrl || "").trim().replace(/\/+$/, "");
  const candidates = [];
  if (normalized) {
    candidates.push(buildOpenAIChatUrl(normalized));
    candidates.push(buildMinimaxTextChatUrl(normalized));
  }
  candidates.push(buildOpenAIChatUrl("https://api.minimaxi.com/v1"));
  candidates.push(buildOpenAIChatUrl("https://api.minimax.io/v1"));
  candidates.push(buildMinimaxTextChatUrl("https://api.minimaxi.com/v1"));
  candidates.push(buildMinimaxTextChatUrl("https://api.minimax.io/v1"));
  return [...new Set(candidates)];
}

function buildChatHeaders(providerConfig = {}) {
  const headers = {
    "Content-Type": "application/json",
  };
  const apiKey = normalizeApiKey(providerConfig.apiKey || "");
  if (!apiKey) return headers;
  headers.Authorization = `Bearer ${apiKey}`;
  if (providerConfig.provider === "minimax") {
    headers["api-key"] = apiKey;
    headers["x-api-key"] = apiKey;
  }
  return headers;
}

async function postChatCompletion(url, body, headers = {}, options = {}) {
  const timeoutMs = Math.max(8000, Number(options.timeoutMs || 24000));
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(new Error(`Semantic model request timed out after ${timeoutMs}ms`)), timeoutMs);
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        ...headers,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
      signal: controller.signal,
    });
    const text = await response.text();
    if (!response.ok) {
      throw new Error(`Semantic model call failed (${response.status}): ${clip(text, 260)}`);
    }
    return text;
  } catch (error) {
    if (String(error?.name || "").toLowerCase() === "aborterror") {
      throw new Error(`Semantic model call timed out after ${timeoutMs}ms`);
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
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

function buildJsonBody(providerConfig, messages, options = {}) {
  const temperature = Math.max(0.02, Math.min(1, Number(options.temperature || 0.2)));
  const maxTokens = Math.max(256, Number(options.maxCompletionTokens || options.maxTokens || 1200));
  const body = {
    model: providerConfig.model,
    messages,
    temperature,
    max_tokens: maxTokens,
  };
  if (providerConfig.provider === "openai") {
    body.max_completion_tokens = maxTokens;
  }
  if (options.topP != null) {
    body.top_p = Math.max(0.01, Math.min(1, Number(options.topP)));
  }
  if (providerConfig.provider === "minimax" && options.expectJson !== false) {
    body.extra_body = { ...(body.extra_body || {}), reasoning_split: true };
  }
  if (options.extraBody && typeof options.extraBody === "object") {
    Object.assign(body, options.extraBody);
  }
  if (options.expectJson !== false && providerConfig.provider === "openai") {
    body.response_format = { type: "json_object" };
  }
  return body;
}

function buildMinimaxModelCandidates(providerConfig = {}) {
  const baseModel = String(providerConfig.model || "").trim();
  return unique([
    baseModel || "MiniMax-M2.7",
    "MiniMax-M2.5-highspeed",
    "MiniMax-M2.5",
  ]);
}

async function callJsonChatOnce(url, providerConfig, messages, options = {}) {
  const body = buildJsonBody(providerConfig, messages, options);
  const text = await postChatCompletion(url, body, buildChatHeaders(providerConfig), options);
  return extractChatContent(text);
}

async function callChatCompletion(providerConfig, system, userContent, options = {}) {
  if (!isSemanticModelEnabled(providerConfig)) return null;
  const messages = buildChatMessages(system, userContent, providerConfig);
  const timeoutMs = Math.max(
    8000,
    Number(
      options.timeoutMs ||
        (providerConfig.provider === "minimax" ? 26000 : providerConfig.provider === "local" ? 22000 : 20000),
    ),
  );
  const urls =
    providerConfig.provider === "minimax"
      ? buildMinimaxChatUrls(providerConfig.baseUrl)
      : providerConfig.provider === "local"
        ? [buildOpenAIChatUrl(providerConfig.baseUrl), buildOllamaChatUrl(providerConfig.baseUrl)]
        : [buildOpenAIChatUrl(providerConfig.baseUrl)];
  const modelCandidates =
    providerConfig.provider === "minimax"
      ? buildMinimaxModelCandidates(providerConfig)
      : [providerConfig.model];

  let lastError = null;
  for (const model of modelCandidates) {
    const modelConfig = { ...providerConfig, model };
    for (const url of urls) {
      try {
        return await callJsonChatOnce(url, modelConfig, messages, { ...options, timeoutMs });
      } catch (error) {
        lastError = error;
        if (String(error?.message || "").includes("response_format")) {
          try {
            return await callJsonChatOnce(url, modelConfig, messages, { ...options, expectJson: false, timeoutMs });
          } catch (retryError) {
            lastError = retryError;
          }
        }
      }
    }
  }

  if (lastError) throw lastError;
  return null;
}

function buildDocumentPrompt(payload) {
  return [
    "You are a semantic planner for Chinese banking PPT generation.",
    "Return one strict JSON object only. No markdown. No explanations. Do not echo the input.",
    "Required top-level keys only: summary, pageCountSuggestion, layoutProfile, blocks, styleHints, globalHints.",
    "Constraints:",
    "1. Split pages by top-level heading boundaries, not by arbitrary paragraphs.",
    "2. Each top-level heading should usually become 1 slide, or 2 slides if it is dense in text, tables, or images.",
    "3. Do not create an extra generic core-summary page.",
    "4. Use heading hierarchy such as ?? ?? 1 1.1 1.1.1 to keep related subsections on the same page group.",
    "5. If there are many tables, prefer dense table, compare table, picture-table, or conclusion-table styles.",
    "6. If there are many images, prefer visual layouts.",
    "7. Try to diversify page families so consecutive slides are not visually identical.",
    "8. pageCountSuggestion must be between 2 and 10, or 0 if uncertain.",
    "9. Each block should describe a page candidate and keep related subsections together.",
  ].join("\n");
}

function buildStylePrompt(payload) {
  return [
    "You analyze reference PPT images for banking-style slide generation.",
    "Return one strict JSON object only. No markdown. No explanations. Do not echo the input.",
    "Required keys: styleFamily, densityBias, layoutBias, preferredVariants, headerStyle, summaryBandStyle, tablePreference, cardStyle, pageRhythm, imagePlacement, iconStyle, iconDiversityPolicy, typography, spacingBias, tableStyleBias, colorMood, repeatedPatterns, summary.",
    "styleFamily must be one of: dense-report, visual-report, boardroom-report, balanced-report.",
    "layoutBias and tableStyleBias should use reusable keywords, not long prose.",
    "preferredVariants should use reusable words such as dense, dashboard, matrix, compare, sidecallout, visual, picture, spread, mosaic, bridge, ladder, storyboard, gallery, timeline, wall.",
    "headerStyle must be one of: formal-line, boardroom-strip, badge-band.",
    "summaryBandStyle must be one of: solid-left-bar, accent-band, card-band, chip-band.",
    "tablePreference must be one of: dense, compare, picture, dashboard.",
    "cardStyle must be one of: classic-card, ribbon-card, soft-card, dashboard-card.",
    "pageRhythm must be one of: dense, balanced, visual.",
    "imagePlacement must be one of: split, hero, gallery.",
  ].join("\n");
}

function buildReviewPrompt(payload) {
  return [
    "You review rendered banking PPT pages for layout quality.",
    "Return one strict JSON object only. No markdown. No explanations. Do not echo the input.",
    "Check for repeated layouts, tiny text, overly narrow images, tables that are too wide or too small, weak spacing, excessive blank areas, and repeated icons.",
    "Required keys only: overallScore, summary, issues, pageAdvice, repeatedPatterns, globalSuggestions, refinementHints, familyUsage.",
    "issues and pageAdvice must be arrays of JSON objects.",
  ].join("\n");
}

function buildReviewTextPayload(payload) {
  const pageLines = (payload.pages || [])
    .map(
      (page) =>
        `P${page.page}: ${page.title || ""} | type=${page.type || ""} | density=${page.density || ""} | template=${page.templateId || ""} | table=${page.tableRowCount || 0}x${page.tableColCount || 0} | image=${page.hasImage ? "yes" : "no"} | shots=${page.hasScreenshots ? "yes" : "no"}`,
    )
    .join("\n");

  const riskLines = (payload.riskPages || [])
    .map((page) => `P${page.page}: ${page.title || ""} | type=${page.type || ""} | density=${page.density || ""} | table=${page.tableRowCount || 0}x${page.tableColCount || 0}`)
    .join("\n");

  return [
    "Please review the rendered PPT pages and return strict JSON only.",
    "Use the page lines and risk lines below as the primary evidence.",
    "Focus on repeated layouts, tiny text, blank areas, narrow images, and table fit issues.",
    pageLines || "No page lines.",
    "Risk pages:",
    riskLines || "No risk pages.",
    "Style hints:",
    JSON.stringify(payload.styleHints || {}, null, 2),
    `Preview count: ${payload.previewCount || 0}`,
  ].join("\n");
}

function buildImagePromptContent(referenceImages = [], heuristicProfile = {}) {
  const textItems = [
    {
      type: "text",
      text: "???????PPT?????????????????????????????? JSON?",
    },
    {
      type: "text",
      text: JSON.stringify({ heuristicProfile }, null, 2),
    },
  ];

  referenceImages
    .slice(0, 4)
    .filter((item) => item?.previewDataUrl)
    .forEach((item, index) => {
      textItems.push({
        type: "text",
        text: `???${index + 1}: ${String(item.name || item.path || "").slice(0, 40)}; size=${item.width || 0}x${item.height || 0}; ratio=${item.aspectRatio || 0}`,
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

function normalizeSemanticLayoutProfile(profile = {}, payload = {}) {
  const counts = payload.counts || {};
  const source = profile && typeof profile === "object" && !Array.isArray(profile) ? profile : {};
  const familyBias = unique([...(Array.isArray(source.familyBias) ? source.familyBias : []), ...(Array.isArray(source.preferredFamilies) ? source.preferredFamilies : [])]);
  const preferredVariants = unique([
    ...(Array.isArray(source.preferredVariants) ? source.preferredVariants : []),
    ...(Array.isArray(source.tableVariants) ? source.tableVariants : []),
    ...(Array.isArray(source.processVariants) ? source.processVariants : []),
    ...(Array.isArray(source.summaryVariants) ? source.summaryVariants : []),
    ...(Array.isArray(source.imageVariants) ? source.imageVariants : []),
  ]);
  const preferredPageTypes = unique([...(Array.isArray(source.preferredPageTypes) ? source.preferredPageTypes : []), ...(Array.isArray(source.pageTypes) ? source.pageTypes : [])]);
  const avoidPatterns = unique([...(Array.isArray(source.avoidPatterns) ? source.avoidPatterns : [])]);

  const density = String(source.density || (counts.tables >= 3 ? "high" : counts.images >= 2 ? "medium" : "balanced")).toLowerCase();
  const spacing = String(source.spacing || (counts.paragraphs >= 18 ? "tight" : "balanced")).toLowerCase();
  const iconVariety = String(source.iconVariety || (counts.images >= 1 ? "medium" : "low")).toLowerCase();
  const tablePreference = String(source.tablePreference || (counts.tables >= 2 ? "dense" : "split")).toLowerCase();

  return {
    ...source,
    density,
    spacing,
    iconVariety,
    tablePreference,
    familyBias: unique(familyBias).slice(0, 10),
    preferredVariants: unique(preferredVariants).slice(0, 12),
    preferredPageTypes: unique(preferredPageTypes).slice(0, 10),
    avoidPatterns,
    source: source.source || "semantic-object",
  };
}

async function repairJsonWithModel(providerConfig, rawText, repairLabel, payload, options = {}) {
  const timeoutMs = Math.max(8000, Number(options.timeoutMs || 15000));
  const prompt = [
    "You are a JSON repair tool.",
    `Target label: ${repairLabel}`,
    "Return one strict JSON object or array only. No markdown. No explanations.",
    "Repair the broken model output below so it becomes valid JSON.",
    "Do not echo the schema unless the broken output is completely unusable.",
    "Reference payload:",
    JSON.stringify(payload, null, 2),
    "Broken output:",
    rawText,
  ].join("\n");

  const repairedText = await callChatCompletion(
    providerConfig,
    "You repair malformed JSON and return only valid JSON.",
    prompt,
    {
      ...options,
      temperature: 0.02,
      expectJson: true,
      timeoutMs,
    },
  );
  return parseJsonResponse(repairedText) || parseJsonResponse(repairJsonCandidate(repairedText || ""));
}

function buildSemanticResult({ providerConfig, parsed, rawText, summary, parseStatus, available = true, extra = {} }) {
  return {
    provider: providerConfig.provider,
    model: providerConfig.model,
    available,
    parseStatus,
    summary,
    rawText: rawText || "",
    ...extra,
    raw: parsed || extra.raw || null,
  };
}

function isObjectLike(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function isMeaningfulDocumentSemanticResult(parsed) {
  if (!isObjectLike(parsed)) return false;
  return Boolean(
    normalizeText(parsed.summary || parsed.overview || "").length >= 6 ||
      Array.isArray(parsed.blocks) && parsed.blocks.length > 0 ||
      isObjectLike(parsed.layoutProfile) && Object.keys(parsed.layoutProfile).length > 0 ||
      Number(parsed.pageCountSuggestion || parsed.pageCount || 0) > 0,
  );
}

function isMeaningfulSectionSemanticResult(parsed) {
  if (Array.isArray(parsed)) return parsed.length > 0;
  if (!isObjectLike(parsed)) return false;
  const pagePlans = Array.isArray(parsed.pagePlans) ? parsed.pagePlans : [];
  return Boolean(
    normalizeText(parsed.summary || "").length >= 6 ||
      pagePlans.length > 0 ||
      Number(parsed.slideCountSuggestion || 0) > 0,
  );
}

function isMeaningfulStyleResult(parsed) {
  if (!isObjectLike(parsed)) return false;
  return Boolean(
    normalizeText(parsed.styleFamily || "").length > 0 ||
      normalizeText(parsed.summary || "").length >= 6 ||
      normalizeText(parsed.headerStyle || "").length > 0 ||
      normalizeText(parsed.tablePreference || "").length > 0 ||
      Array.isArray(parsed.preferredVariants) && parsed.preferredVariants.length > 0 ||
      Array.isArray(parsed.layoutBias) && parsed.layoutBias.length > 0 ||
      Array.isArray(parsed.tableStyleBias) && parsed.tableStyleBias.length > 0,
  );
}

function isMeaningfulReviewResult(parsed) {
  if (!isObjectLike(parsed)) return false;
  return Boolean(
    Number(parsed.overallScore || 0) > 0 ||
      normalizeText(parsed.summary || "").length >= 6 ||
      Array.isArray(parsed.issues) && parsed.issues.length > 0 ||
      Array.isArray(parsed.pageAdvice) && parsed.pageAdvice.length > 0,
  );
}

function isAuthLikeSemanticError(error) {
  const message = normalizeText(error?.message || error || "").toLowerCase();
  return (
    message.includes("401") ||
    message.includes("403") ||
    message.includes("invalid api key") ||
    message.includes("unauthorized") ||
    message.includes("forbidden") ||
    message.includes("authentication")
  );
}

function extractQuotedJsonField(text = "", key = "") {
  const source = String(text || "");
  if (!source || !key) return "";
  const match = source.match(new RegExp(`"${key}"\\s*:\\s*"((?:\\\\.|[^"])*)"`, "i"));
  if (!match) return "";
  try {
    return JSON.parse(`"${match[1].replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`);
  } catch {
    return String(match[1] || "")
      .replace(/\\"/g, '"')
      .replace(/\\\\/g, "\\")
      .trim();
  }
}

function extractNumericJsonField(text = "", key = "") {
  const source = String(text || "");
  if (!source || !key) return 0;
  const match = source.match(new RegExp(`"${key}"\\s*:\\s*(-?\\d+(?:\\.\\d+)?)`, "i"));
  if (!match) return 0;
  return Number(match[1] || 0) || 0;
}

function normalizeReviewResult(parsed, rawText = "", payload = {}) {
  const base = isObjectLike(parsed) ? { ...parsed } : {};
  const summary =
    normalizeText(base.summary || base.overview || extractQuotedJsonField(rawText, "summary") || "") ||
    (rawText ? "Rendered-page review completed." : "");
  const overallScore = Number(base.overallScore || extractNumericJsonField(rawText, "overallScore") || 0);
  const issues = Array.isArray(base.issues) ? base.issues.filter(Boolean) : [];
  const pageAdvice = Array.isArray(base.pageAdvice) ? base.pageAdvice.filter(Boolean) : [];
  const repeatedPatterns = Array.isArray(base.repeatedPatterns) ? base.repeatedPatterns.filter(Boolean) : [];
  const globalSuggestions = Array.isArray(base.globalSuggestions) ? base.globalSuggestions.filter(Boolean) : [];
  const refinementHints = Array.isArray(base.refinementHints) ? base.refinementHints.filter(Boolean) : [];
  const familyUsage = isObjectLike(base.familyUsage) ? base.familyUsage : {};
  return {
    ...base,
    overallScore: overallScore > 0 ? overallScore : Number(payload?.previewCount || 0) > 0 ? 60 : 0,
    summary,
    issues,
    pageAdvice,
    repeatedPatterns,
    globalSuggestions,
    refinementHints,
    familyUsage,
  };
}

async function analyzeDocumentSemantics(context = {}, options = {}) {
  const providerConfig = buildProviderConfig(options);
  if (!isSemanticModelEnabled(providerConfig)) return null;

  const payload = buildDocPayload(context, {
    compactPayload: Boolean(options.compactPayload),
  });
  try {
    const text = await callChatCompletion(providerConfig, buildDocumentPrompt(payload), JSON.stringify(payload, null, 2), {
      temperature: 0.04,
      maxCompletionTokens: 900,
      maxTokens: 900,
      timeoutMs: Number(options.timeoutMs || 0) || undefined,
    });
    const parsed = parseJsonResponse(text) || (options.allowRepair === false ? null : await repairJsonWithModel(providerConfig, text || "", "document_semantics", payload, { temperature: 0.02 }));
    if (!isMeaningfulDocumentSemanticResult(parsed)) {
      return buildSemanticResult({
        providerConfig,
        rawText: text,
        available: true,
        parseStatus: "fallback",
        summary: "Semantic planning fell back to rule-based planning because the model output was not stable enough to parse.",
      });
    }
    return buildSemanticResult({
      providerConfig,
      parsed,
      rawText: text || "",
      available: true,
      parseStatus: text && parsed ? "parsed" : "repaired",
      summary: normalizeText(parsed.summary || parsed.overview || "Semantic planning completed."),
      extra: {
        pageCountSuggestion: Number(parsed.pageCountSuggestion || parsed.pageCount || 0) || 0,
        layoutProfile: normalizeSemanticLayoutProfile(parsed.layoutProfile || {}, payload),
        blocks: Array.isArray(parsed.blocks) ? parsed.blocks.slice(0, 12) : [],
        styleHints: parsed.styleHints || {},
        globalHints: parsed.globalHints || {},
      },
    });
  } catch (error) {
    const fallbackAvailable = !isAuthLikeSemanticError(error);
    return buildSemanticResult({
      providerConfig,
      available: fallbackAvailable,
      parseStatus: fallbackAvailable ? "fallback" : "error",
      summary: fallbackAvailable
        ? `Semantic model call returned an unstable response and the workflow fell back to rules: ${error.message || String(error)}`
        : `Semantic model call failed and the workflow fell back to rules: ${error.message || String(error)}`,
      rawText: "",
      extra: { error: error.message || String(error) },
    });
  }
}
async function analyzeSectionSemantics(section = {}, options = {}) {
  const providerConfig = buildProviderConfig(options);
  if (!isSemanticModelEnabled(providerConfig)) return null;

  const payload = buildSectionPayload(section, options);
  try {
    const text = await callChatCompletion(providerConfig, buildSectionPrompt(payload), JSON.stringify(payload, null, 2), {
      temperature: 0.04,
      maxCompletionTokens: 1400,
      maxTokens: 1400,
    });
    const parsed = parseJsonResponse(text) || (options.allowRepair === false ? null : await repairJsonWithModel(providerConfig, text || "", "section_semantics", payload, { temperature: 0.02 }));
    const normalizedParsed = Array.isArray(parsed)
      ? {
          summary: "Section semantic planning completed.",
          slideCountSuggestion: Math.max(1, Math.min(2, parsed.length || 1)),
          pagePlans: parsed,
        }
      : parsed;
    if (!isMeaningfulSectionSemanticResult(normalizedParsed)) {
      return buildSemanticResult({
        providerConfig,
        rawText: text,
        available: true,
        parseStatus: "fallback",
        summary: "Section semantic planning fell back to rules because the model output could not be parsed reliably.",
      });
    }
    const pagePlans = Array.isArray(normalizedParsed.pagePlans) ? normalizedParsed.pagePlans.slice(0, 2) : [];
    return buildSemanticResult({
      providerConfig,
      parsed: normalizedParsed,
      rawText: text || "",
      available: true,
      parseStatus: text && normalizedParsed ? "parsed" : "repaired",
      summary: normalizeText(normalizedParsed.summary || "Section semantic planning completed."),
      extra: {
        slideCountSuggestion: Math.max(1, Math.min(2, Number(normalizedParsed.slideCountSuggestion || pagePlans.length || 1))),
        layoutProfile: String(normalizedParsed.layoutProfile || "").toLowerCase(),
        pagePlans,
      },
    });
  } catch (error) {
    const fallbackAvailable = !isAuthLikeSemanticError(error);
    return buildSemanticResult({
      providerConfig,
      available: fallbackAvailable,
      parseStatus: fallbackAvailable ? "fallback" : "error",
      summary: fallbackAvailable
        ? `Section semantic planning returned an unstable response and fell back to rules: ${error.message || String(error)}`
        : `Section semantic planning failed: ${error.message || String(error)}`,
      rawText: "",
      extra: { error: error.message || String(error) },
    });
  }
}
async function analyzeReferenceStyle(referenceImages = [], referenceStyleProfile = {}, options = {}) {
  const providerConfig = buildProviderConfig(options);
  if (!isSemanticModelEnabled(providerConfig)) return null;
  if (!referenceImages.length) {
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: true,
      parseStatus: "skipped",
      summary: "Reference-style semantic analysis was skipped because no reference images were provided.",
    };
  }

  const payload = buildStylePayload(referenceImages, referenceStyleProfile);
  const images = referenceImages
    .slice(0, 4)
    .map((item) => ({
      ...item,
      previewDataUrl: item.previewDataUrl || (item.path ? fileToDataUrl(item.path) : ""),
    }))
    .filter((item) => item.previewDataUrl);

  try {
    const userContent = providerConfig.supportsImages ? buildImagePromptContent(images, payload.heuristicProfile || {}) : JSON.stringify(payload, null, 2);
    const text = await callChatCompletion(providerConfig, buildStylePrompt(payload), userContent, {
      temperature: 0.06,
      maxCompletionTokens: 900,
      maxTokens: 900,
    });
    const parsed = parseJsonResponse(text) || (options.allowRepair === false ? null : await repairJsonWithModel(providerConfig, text || "", "style_analysis", payload, { temperature: 0.02 }));
    if (!isMeaningfulStyleResult(parsed)) {
      return {
        provider: providerConfig.provider,
        model: providerConfig.model,
        available: true,
        parseStatus: "fallback",
        summary: "Reference-style semantic analysis could not be parsed reliably, so heuristic style hints were kept.",
        rawText: text || "",
      };
    }
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: true,
      parseStatus: text && parsed ? "parsed" : "repaired",
      summary: normalizeText(parsed.summary || parsed.styleFamily || "Reference-style analysis completed."),
      styleFamily: parsed.styleFamily || referenceStyleProfile.styleFamily || "",
      densityBias: parsed.densityBias || referenceStyleProfile.densityBias || "",
      layoutBias: Array.isArray(parsed.layoutBias) ? parsed.layoutBias : [],
      preferredVariants: Array.isArray(parsed.preferredVariants) ? parsed.preferredVariants : [],
      headerStyle: parsed.headerStyle || "",
      summaryBandStyle: parsed.summaryBandStyle || "",
      tablePreference: parsed.tablePreference || "",
      cardStyle: parsed.cardStyle || "",
      pageRhythm: parsed.pageRhythm || "",
      imagePlacement: parsed.imagePlacement || "",
      iconStyle: parsed.iconStyle || "",
      iconDiversityPolicy: parsed.iconDiversityPolicy || parsed.iconDiversity || "",
      typography: parsed.typography || {},
      spacingBias: parsed.spacingBias || "",
      tableStyleBias: Array.isArray(parsed.tableStyleBias) ? parsed.tableStyleBias : [],
      colorMood: parsed.colorMood || "",
      repeatedPatterns: Array.isArray(parsed.repeatedPatterns) ? parsed.repeatedPatterns : [],
      raw: parsed,
      rawText: text || "",
    };
  } catch (error) {
    const fallbackAvailable = !isAuthLikeSemanticError(error);
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: fallbackAvailable,
      parseStatus: fallbackAvailable ? "fallback" : "error",
      summary: fallbackAvailable
        ? `Reference-style semantic analysis returned an unstable response, so heuristic style hints were kept: ${error.message || String(error)}`
        : `Reference-style semantic analysis failed: ${error.message || String(error)}`,
      error: error.message || String(error),
    };
  }
}
function previewContentToUser(previews = [], outline = {}, style = {}) {
  const lines = (outline.slides || [])
    .map((slide) => `P${slide.page}: ${slide.title || ""} | type=${slide.type || ""} | template=${slide.templateId || ""}`)
    .join("\n");
  const items = [
    {
      type: "text",
      text: lines || "No page summary available.",
    },
  ];

  previews
    .slice(0, 4)
    .filter((item) => item?.dataUrl)
    .forEach((item, index) => {
      items.push({
        type: "text",
        text: `Preview ${index + 1}: page ${item.page} ${item.title || ""}` ,
      });
      items.push({
        type: "image_url",
        image_url: {
          url: item.dataUrl,
        },
      });
    });

  if (!previews.length) {
    items.push({
      type: "text",
      text: JSON.stringify({ outline: { pages: outline.slides?.length || 0 }, style: { palette: style.palette || {} } }, null, 2),
    });
  }

  return items;
}

async function reviewRenderedDeck(previews = [], outline = {}, style = {}, options = {}) {
  const providerConfig = buildProviderConfig(options);
  if (!isSemanticModelEnabled(providerConfig)) return null;

  const payload = buildReviewPayload(previews, outline, style);
  const useImages = providerConfig.supportsImages && previews.some((preview) => preview.dataUrl);
  const userContent = useImages ? previewContentToUser(previews, outline, style) : buildReviewTextPayload(payload);

  try {
    const text = await callChatCompletion(providerConfig, buildReviewPrompt(payload), userContent, {
      temperature: 0.05,
      maxCompletionTokens: 1400,
      maxTokens: 1400,
    });
    const parsed = parseJsonResponse(text) || (options.allowRepair === false ? null : await repairJsonWithModel(providerConfig, text || "", "render_review", payload, { temperature: 0.02 }));
    const normalized = normalizeReviewResult(parsed, text || "", payload);
    if (!isMeaningfulReviewResult(normalized)) {
      return {
        provider: providerConfig.provider,
        model: providerConfig.model,
        available: true,
        parseStatus: "fallback",
        visionSupported: useImages,
        summary: useImages ? "Rendered-page review was executed, but the result could not be parsed reliably." : "Text-only review was executed, but the result could not be parsed reliably.",
        rawText: text || "",
        raw: normalized,
      };
    }
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: true,
      parseStatus: text && parsed ? "parsed" : "repaired",
      visionSupported: useImages,
      rawText: text || "",
      ...normalized,
      raw: normalized,
    };
  } catch (error) {
    const fallbackAvailable = !isAuthLikeSemanticError(error);
    return {
      provider: providerConfig.provider,
      model: providerConfig.model,
      available: fallbackAvailable,
      parseStatus: fallbackAvailable ? "fallback" : "error",
      visionSupported: useImages,
      summary: fallbackAvailable
        ? `Rendered-page review returned an unstable response and fell back to heuristic review: ${error.message || String(error)}`
        : `Rendered-page review failed: ${error.message || String(error)}`,
      error: error.message || String(error),
    };
  }
}
module.exports = {
  analyzeDocumentSemantics,
  analyzeSectionSemantics,
  analyzeReferenceStyle,
  buildProviderConfig,
  isSemanticModelEnabled,
  reviewRenderedDeck,
  parseJsonResponse,
};
