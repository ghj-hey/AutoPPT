function readableTypeLabel(type) {
  return (
    {
      cover: '封面页',
      summary_cards: '摘要卡片页',
      table_analysis: '表格分析页',
      process_flow: '流程方法页',
      bullet_columns: '分栏要点页',
      image_story: '图文展示页',
      action_plan: '行动计划页',
      key_takeaways: '结论收束页',
    }[type] || type || '页面'
  );
}

function semanticStatusParts(semanticAnalysis = null, semanticReview = null) {
  const provider = semanticAnalysis?.provider || semanticReview?.provider || '未启用';
  const model = semanticAnalysis?.model || semanticReview?.model || '未配置';
  const available = semanticAnalysis?.available ?? semanticReview?.available ?? false;
  const analysisParse = semanticAnalysis?.parseStatus || 'unknown';
  const reviewParse = semanticReview?.parseStatus || 'unknown';
  return { provider, model, available, analysisParse, reviewParse };
}

function semanticStatusText(semanticAnalysis = null, semanticReview = null) {
  const parts = semanticStatusParts(semanticAnalysis, semanticReview);
  return `语义模型：${parts.provider} / ${parts.model}；可用性：${parts.available ? '可用' : '不可用'}；规划解析：${parts.analysisParse}；效果复核：${parts.reviewParse}`;
}

function buildReadableWorkflowStagesCn({
  documentSummary,
  effectiveLibrary,
  referenceStyleProfile,
  semanticAnalysis,
  semanticReview,
  pageCount,
  referencePptCount = 0,
  derivedReferenceImageCount = 0,
}) {
  const semantic = semanticStatusParts(semanticAnalysis, semanticReview);
  const sectionCount = documentSummary?.sectionTree?.length || documentSummary?.counts?.sections || 0;

  return [
    {
      id: 'upload',
      title: '上传整理',
      detail: '需求文档、模板、素材 PPT 和参考 PPT / 图片已归档到当前会话。',
      progress: 0.08,
    },
    {
      id: 'parse',
      title: '文档解析',
      detail: `识别到 ${documentSummary?.counts?.paragraphs || 0} 段正文、${documentSummary?.counts?.tables || 0} 个表格、${documentSummary?.counts?.images || 0} 张图片，以及 ${sectionCount} 个一级章节。`,
      progress: 0.22,
    },
    {
      id: 'structure',
      title: '标题树与 Markdown',
      detail: '系统已按标题层级构建结构树，并生成 Markdown 中间层，用于后续分页和语义总结。',
      progress: 0.36,
    },
    {
      id: 'reference',
      title: '参考素材融合',
      detail:
        referencePptCount || referenceStyleProfile?.count
          ? `已处理 ${referencePptCount} 份参考 PPT，并拆分出 ${derivedReferenceImageCount} 张参考页图，与素材库风格合并使用。`
          : '未上传参考 PPT 时，将沿用当前素材库和默认金融汇报风格。',
      progress: 0.52,
    },
    {
      id: 'semantic',
      title: '语义规划',
      detail: semanticAnalysis?.summary || semanticStatusText(semanticAnalysis, semanticReview),
      progress: 0.68,
    },
    {
      id: 'semantic-availability',
      title: '语义模型状态',
      detail: `当前调用：${semantic.provider} / ${semantic.model}；可用性：${semantic.available ? '可用' : '不可用'}；规划解析：${semantic.analysisParse}；效果复核：${semantic.reviewParse}`,
      progress: 0.78,
    },
    {
      id: 'layout',
      title: '动态页面规划',
      detail: `系统已根据一级标题、正文密度、表格和图片自动规划为 ${pageCount || 0} 页，并对相邻页面进行差异化布局分配。`,
      progress: 0.9,
    },
    {
      id: 'draft',
      title: '草稿输出',
      detail: effectiveLibrary?.name
        ? `已输出可编辑草稿，并结合参考库 ${effectiveLibrary.name} 完成页面结构和素材映射。`
        : '已输出可编辑草稿。',
      progress: 1,
    },
  ];
}

function buildReadableDownloadManifestCn(sessionId) {
  return [
    { type: 'draft-deck', title: '草稿 PPT' },
    { type: 'outline-draft', title: '草稿 Outline' },
    { type: 'style-draft', title: '草稿 Style' },
    { type: 'structure-draft', title: '标题结构图' },
    { type: 'structure-md-draft', title: 'Markdown 中间稿' },
    { type: 'layout-draft', title: '布局规划' },
    { type: 'summary-draft', title: '文档摘要' },
    { type: 'notes-draft', title: '草稿备注' },
    { type: 'reference-draft', title: '参考库摘要' },
    { type: 'reference-style-draft', title: '参考风格分析' },
    { type: 'semantic-draft', title: '语义分析' },
    { type: 'semantic-review-draft', title: '效果复核' },
    { type: 'semantic-refined-draft', title: '语义优化布局' },
  ].map((item) => ({ ...item, url: `/api/download/${sessionId}/${item.type}` }));
}

function buildReadableFinalDownloadManifestCn(sessionId) {
  return [
    { type: 'deck', title: '最终 PPT' },
    { type: 'outline', title: '最终 Outline' },
    { type: 'style', title: '最终 Style' },
    { type: 'structure', title: '标题结构图' },
    { type: 'structure-md', title: 'Markdown 结构稿' },
    { type: 'layout', title: '最终布局' },
    { type: 'notes', title: '最终备注' },
    { type: 'summary', title: '文档摘要' },
    { type: 'reference-style', title: '参考风格分析' },
    { type: 'semantic', title: '语义分析' },
    { type: 'semantic-review', title: '效果复核' },
    { type: 'semantic-refined', title: '语义优化布局' },
  ].map((item) => ({ ...item, url: `/api/download/${sessionId}/${item.type}` }));
}

function buildArchivedDeliveryManifestCn({
  sessionId,
  archivedAt = '',
  deliveryDir = '',
  sourceSessionDir = '',
  files = [],
} = {}) {
  return {
    sessionId,
    archivedAt,
    deliveryDir,
    sourceSessionDir,
    outputDir: deliveryDir ? `${deliveryDir}/output` : '',
    files: (files || []).map((item) => ({
      title: item.title,
      type: item.type,
      url: `/api/download/${sessionId}/${item.type}`,
      path: item.path || '',
    })),
  };
}

module.exports = {
  readableTypeLabel,
  semanticStatusText,
  buildReadableWorkflowStagesCn,
  buildReadableDownloadManifestCn,
  buildReadableFinalDownloadManifestCn,
  buildArchivedDeliveryManifestCn,
};
