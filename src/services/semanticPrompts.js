function buildDocumentSystemPromptV2() {
  return [
    "你是银行PPT语义规划器。",
    "只输出一个合法 JSON 对象，不要 Markdown，不要代码块，不要解释。",
    "任务：根据需求文档的章节、表格、图片和段落结构，规划适合做成 PPT 的页面切分与风格方向。",
    "规则：",
    "1. 相邻页面尽量不要使用同一种布局家族。",
    "2. 内容少则合并，内容多则拆分；同一主题尽量放在同一页，除非内容过多。",
    "3. 如果表格多或列多，优先选择高密度或对比型页面；如果图片多，优先选择图文型页面；如果是步骤/措施，则优先选择流程型页面。",
    "4. pageCountSuggestion 只能是 2 到 10 的整数。",
    "5. blocks 最多输出 10 个，每个 block 只保留最关键的 sectionIndexes。",
    "6. 只输出这些字段：summary, pageCountSuggestion, layoutProfile, blocks, styleHints, globalHints。",
  ].join("\n");
}

function buildDocumentUserPromptV2(payload) {
  const sectionLines = (payload.sections || [])
    .map((section) => {
      const snippets = (section.text || []).slice(0, 2).join(" / ");
      return `${section.index + 1}. ${section.title || section.heading || "未命名章节"} | level=${section.level || 1} | paragraphs=${section.paragraphCount || 0}${snippets ? ` | ${snippets}` : ""}`;
    })
    .join("\n");

  const tableLines = (payload.tables || [])
    .map(
      (table) =>
        `${table.index + 1}. rows=${table.rows || 0}, cols=${table.columns || 0}${(table.preview || []).length ? ` | ${table.preview[0]?.join(" / ") || ""}` : ""}`,
    )
    .join("\n");

  const imageLines = (payload.images || [])
    .map((image) => `${image.index + 1}. page=${image.page || "-"} | aspect=${image.aspectRatio || "-"} | ${image.path || ""}`)
    .join("\n");

  return [
    "请把下面的需求文档摘要转换为 PPT 页面规划 JSON，不要复述原文内容。",
    "只输出 JSON 对象，不要解释，不要代码块。",
    `标题：${payload.title || ""}`,
    `部门：${payload.department || ""}`,
    `汇报人：${payload.presenter || ""}`,
    `建议页数：${payload.requestedPages || 0}（0 表示自动）`,
    `正文段落数：${payload.counts?.paragraphs || 0}，表格数：${payload.counts?.tables || 0}，图片数：${payload.counts?.images || 0}，章节数：${payload.counts?.sections || 0}`,
    "章节摘要：",
    sectionLines || "无",
    "表格摘要：",
    tableLines || "无",
    "图片摘要：",
    imageLines || "无",
    "输出要求：",
    "1. 只输出 summary, pageCountSuggestion, layoutProfile, blocks, styleHints, globalHints。",
    "2. summary 必须是一句话。",
    "3. blocks 最多 8 个，每个 block 只保留最关键的 sectionIndexes 和布局建议。",
    "4. 内容少就合并，内容多就拆分，但不要让每页内容过散。",
  ].join("\n");
}

function buildStylePromptV2(payload) {
  return [
    "你是银行PPT风格分析器。",
    "只输出一个合法 JSON 对象，不要 Markdown，不要代码块，不要解释。",
    "请根据参考PPT图片判断 styleFamily、densityBias、layoutBias、iconStyle、iconDiversityPolicy、typography、spacingBias、tableStyleBias、colorMood、repeatedPatterns。",
    "styleFamily 只能从 dense-report、visual-report、boardroom-report、balanced-report 中选择。",
    "densityBias 只能从 low、medium、high 中选择。",
    "layoutBias 请输出适合复用的布局关键词，而不是描述句子。",
    "只输出这些字段：styleFamily, densityBias, layoutBias, iconStyle, iconDiversityPolicy, typography, spacingBias, tableStyleBias, colorMood, repeatedPatterns, summary。",
    JSON.stringify(payload, null, 2),
  ].join("\n");
}

function buildImagePromptContentV2(referenceImages = [], heuristicProfile = {}) {
  const textItems = [
    {
      type: "text",
      text: "请分析这些参考PPT图片的布局、字体、图标风格、留白和表格密度，只输出 JSON 对象，不要解释。",
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
        text: `参考图${index + 1}：${String(item.name || item.path || "").slice(0, 40)}，尺寸 ${item.width || 0}x${item.height || 0}，长宽比 ${item.aspectRatio || 0}`,
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

function buildReviewPromptV2(payload) {
  return [
    "你是银行PPT版式审阅器。",
    "只输出一个合法 JSON 对象，不要 Markdown，不要代码块，不要解释。",
    "请检查：页面重复、文字过小、表格过宽、留白过多、图标重复、左中右失衡、编码异常、图片比例异常。",
    "输出字段只保留：overallScore, summary, issues, pageAdvice, repeatedPatterns, globalSuggestions, refinementHints, familyUsage。",
    "issues 最多 8 条；pageAdvice 最多 6 条；每条 pageAdvice 只保留一个 preferredFamily 和 preferredVariant。",
    "preferredFamily 只能从 dense、split、visual、process、action、stack 中选择。",
    JSON.stringify(payload, null, 2),
  ].join("\n");
}

function buildReviewTextContentV2(payload) {
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
    "请根据下面的 PPT 页面摘要输出审阅 JSON，重点检查重复、表格过宽、文字过小、留白过多、图标重复和左右失衡。",
    "不要复述输入内容，不要 Markdown，只输出 JSON 对象。",
    "页面摘要：",
    pageLines || "无",
    "重复模式：",
    (payload.repeatedPatterns || []).join(" | ") || "无",
    "风险页面：",
    riskLines || "无",
    "风格提示：",
    JSON.stringify(payload.styleHints || {}, null, 2),
    `预览数量：${payload.previewCount || 0}`,
    "输出要求：overallScore、summary、issues、pageAdvice、repeatedPatterns、globalSuggestions、refinementHints、familyUsage。",
  ].join("\n");
}

module.exports = {
  buildDocumentSystemPromptV2,
  buildDocumentUserPromptV2,
  buildStylePromptV2,
  buildImagePromptContentV2,
  buildReviewPromptV2,
  buildReviewTextContentV2,
};
