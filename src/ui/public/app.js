const state = {
  sessionId: '',
  pollingTimer: null,
  libraries: [],
  libraryLayouts: null,
  selectedLibraryId: '',
  currentMeta: null,
  currentLayouts: {},
  outlineText: '',
  styleText: '',
  notesText: '',
  layoutData: null,
  loadingArtifacts: false,
  generating: false,
};

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function clip(value, max = 120) {
  const text = String(value || '').trim();
  if (text.length <= max) return text;
  return `${text.slice(0, Math.max(24, max - 1)).trim()}...`;
}

function byId(id) {
  return document.getElementById(id);
}

function setHtml(id, html) {
  const el = byId(id);
  if (el) el.innerHTML = html;
}

function renderProgressBar(containerId, progress = 0, label = '', detail = '', tone = 'active') {
  const container = byId(containerId);
  if (!container) return;
  const pct = Math.max(0, Math.min(100, Number(progress) || 0));
  container.innerHTML = `
    <div class="progress-shell ${tone}">
      <div class="progress-head">
        <span>${escapeHtml(label)}</span>
        <strong>${Math.round(pct)}%</strong>
      </div>
      <div class="progress-track"><div class="progress-fill" style="width:${pct}%"></div></div>
      <div class="progress-detail">${escapeHtml(detail)}</div>
    </div>
  `;
}

function renderStatusItems(items = []) {
  setHtml(
    'status-list',
    (items || [])
      .map(
        (item) => `
          <div class="status-item">
            <strong>${escapeHtml(item.title || '')}</strong>
            <div>${escapeHtml(item.body || '')}</div>
          </div>
        `,
      )
      .join(''),
  );
}

function semanticLabel(available) {
  return available ? '可用' : '不可用';
}

function renderSemanticStatus(meta = {}) {
  const analysis = meta.semanticAnalysis || {};
  const review = meta.semanticReview || {};
  const provider = analysis.provider || review.provider || meta.semanticModelProvider || '未启用';
  const model = analysis.model || review.model || meta.semanticModelName || '未配置';
  const available = analysis.available ?? review.available ?? meta.semanticModelAvailable ?? false;
  const analysisParse = analysis.parseStatus || 'unknown';
  const reviewParse = review.parseStatus || 'unknown';

  setHtml(
    'semantic-status',
    `
      <div class="semantic-status-card">
        <div class="semantic-status-head">
          <strong>语义模型状态</strong>
          <span class="semantic-pill ${available ? 'ok' : 'off'}">${semanticLabel(available)}</span>
        </div>
        <div class="semantic-status-grid">
          <div><span>提供方</span><strong>${escapeHtml(provider)}</strong></div>
          <div><span>模型</span><strong>${escapeHtml(model)}</strong></div>
          <div><span>规划解析</span><strong>${escapeHtml(analysisParse)}</strong></div>
          <div><span>效果复核</span><strong>${escapeHtml(reviewParse)}</strong></div>
        </div>
      </div>
    `,
  );
}

function buildStatusCards(meta = {}) {
  const items = [];
  if (meta.sessionId) items.push({ title: '草稿已生成', body: `会话 ID：${meta.sessionId}` });
  if (meta.pageCount) items.push({ title: '动态页数', body: `本次共规划 ${meta.pageCount} 页 PPT。` });
  if (meta.referencePptPaths?.length) items.push({ title: '参考 PPT 已处理', body: `本次共处理 ${meta.referencePptPaths.length} 份参考 PPT。` });
  if (meta.derivedReferenceImageCount) items.push({ title: '参考页图片已拆分', body: `后台共拆分 ${meta.derivedReferenceImageCount} 张参考页图片。` });
  if (meta.archivedAt) items.push({ title: '会话已归档', body: '最终文件已迁移到独立交付目录，原始会话工作区已销毁。' });
  if (meta.semanticAnalysis?.summary) items.push({ title: '语义模型摘要', body: meta.semanticAnalysis.summary });
  if (meta.semanticReview?.summary) items.push({ title: '效果评审摘要', body: meta.semanticReview.summary });
  if (meta.referenceStyleProfile?.styleFamily || meta.referenceStyleProfile?.suggestedLayoutSet) {
    items.push({
      title: '本次生效参考风格',
      body: `${meta.referenceStyleProfile?.styleFamily || '未识别'} / ${meta.referenceStyleProfile?.suggestedLayoutSet || '默认布局'}`,
    });
  }
  return items.map((item) => ({
    ...item,
    body: humanizeStatusText(item.body || ''),
  }));
}

function previewSrc(sample = {}) {
  if (sample.previewDataUrl) return sample.previewDataUrl;
  if (sample.previewUrl) return sample.previewUrl;
  if (sample.path) return `/api/library-preview?path=${encodeURIComponent(sample.path)}`;
  return '';
}

function friendlyTemplateLabel(id = '') {
  const value = String(id || '').trim();
  const map = {
    cover_formal_v1: '正式封面',
    cover_clean_v1: '简洁封面',
    summary_grid_v1: '摘要网格',
    summary_spread_v1: '摘要铺陈',
    summary_mosaic_v1: '摘要拼接',
    table_compare_v1: '左右对比表格',
    table_visual_v1: '图表结合',
    table_split_v1: '表格分栏',
    table_sidecallout_v1: '表格 + 结论卡',
    table_dashboard_v1: '高密度数据看板',
    table_dense_v1: '高密度表格',
    process_three_lane_v1: '三段流程',
    process_bridge_v1: '桥接流程',
    bullet_dual_v1: '双栏要点',
    bullet_masonry_v1: '瀑布要点',
    image_split_v1: '图文分栏',
    image_storyboard_v1: '图文故事板',
    image_gallery_v1: '图片画廊',
    action_dashboard_v1: '行动看板',
    action_stacked_v1: '堆叠行动卡',
    action_timeline_v1: '时间推进',
    takeaway_cards_v1: '结论卡片',
    takeaway_wall_v1: '结论墙',
  };
  return map[value] || '';
}

function friendlyLayoutSetLabel(id = '') {
  const value = String(id || '').trim();
  const map = {
    bank_finance_default: '标准金融汇报',
    bank_finance_dense: '高密度汇报',
    bank_finance_visual: '图文强化',
    bank_finance_highlight: '重点高亮',
    bank_finance_boardroom: '正式工作会',
    bank_finance_reporting: '综合材料型',
    bank_finance_dynamic: '动态组合型',
  };
  return map[value] || '';
}

function friendlyPageTypeLabel(type = '') {
  const value = String(type || '').trim();
  const map = {
    cover: '封面页',
    summary_cards: '摘要页',
    table_analysis: '表格分析页',
    process_flow: '流程方法页',
    bullet_columns: '分栏要点页',
    image_story: '图文展示页',
    action_plan: '行动计划页',
    key_takeaways: '结论收束页',
  };
  return map[value] || '';
}

function containsChineseText(value = '') {
  return /[\u3400-\u9fff]/.test(String(value || ''));
}

function sanitizeVisibleChineseLabel(value = '') {
  const text = String(value || '').trim();
  if (!text || /[?]{2,}|[�]/.test(text)) return '';
  return containsChineseText(text) ? text : '';
}

function normalizeTemplateOptionLabel(option = {}, index = 0) {
  const label = sanitizeVisibleChineseLabel(option.label);
  const displayName = sanitizeVisibleChineseLabel(option.displayName);
  const id = String(option.id || '').trim();
  const friendly = friendlyTemplateLabel(id);
  return displayName || label || friendly || `模板方案 ${index + 1}`;
}

function normalizeLayoutSetLabel(option = {}, index = 0) {
  const displayName = sanitizeVisibleChineseLabel(option.displayName);
  const label = sanitizeVisibleChineseLabel(option.label);
  const friendly = friendlyLayoutSetLabel(option.id);
  return displayName || label || friendly || `布局集 ${index + 1}`;
}

function humanizeStatusText(text = '') {
  const value = String(text || '').trim();
  if (!value) return '';
  if (/Text-only review was executed, but the result could not be parsed reliably\./i.test(value)) {
    return '已执行文本结构复核，但返回结果不够稳定，系统已按规则复核结果继续运行。';
  }
  if (/Semantic model call failed and the workflow fell back to rules:/i.test(value)) {
    if (/invalid api key/i.test(value)) {
      return '语义模型调用失败：当前 Minimax API Key 无效或未生效，系统已自动回退到规则规划。请刷新页面后重新填写 API Key 再运行。';
    }
    return '语义模型调用失败，系统已自动回退到规则规划。';
  }
  if (/invalid api key/i.test(value)) {
    return '当前 API Key 无效或未生效，请刷新页面后重新填写。';
  }
  return value;
}

function renderLibrarySummary() {
  const container = byId('library-summary');
  if (!container) return;
  const library = state.libraries.find((item) => item.id === state.selectedLibraryId) || state.libraries[0];
  if (!library) {
    container.innerHTML = '<div class="status-item"><strong>暂无参考库</strong><div>请先上传参考 PPT 或素材 PPT。</div></div>';
    return;
  }

  const categories = Array.isArray(library.categories) ? library.categories : [];
  const chips = [
    { label: '素材', value: library.counts?.media || 0 },
    { label: '图标', value: library.counts?.icons || 0 },
    { label: '品牌', value: library.counts?.branding || 0 },
    { label: '截图', value: library.counts?.screenshotsOrCharts || 0 },
  ];

  container.innerHTML = `
    <div class="library-overview-head">
      <div>
        <strong>${escapeHtml(library.displayName || library.sourceName || library.name || library.id)}</strong>
        <p>${escapeHtml(library.sourcePptx || '当前参考库已按类别整理，可横向滚动查看缩略素材。')}</p>
      </div>
      <div class="library-overview-stats">
        ${chips.map((chip) => `<span>${escapeHtml(chip.label)} ${chip.value}</span>`).join('')}
      </div>
    </div>
    <div class="library-scroll">
      ${categories
        .map((category) => {
          const samples = (category.samples || []).slice(0, 4);
          return `
            <div class="material-card">
              <div class="material-card-head">
                <div>
                  <strong>${escapeHtml(category.label || category.category || '素材')}</strong>
                  <span>${escapeHtml(`${category.count || 0} 个素材`)}</span>
                </div>
              </div>
              <div class="material-thumb-grid">
                ${samples
                  .map((sample) => {
                    const src = previewSrc(sample);
                    return `
                      <div class="material-thumb">
                        ${src ? `<img src="${src}" alt="${escapeHtml(sample.name || '')}" loading="lazy" />` : '<div class="preview-empty">无预览</div>'}
                        <div class="material-thumb-badge">${escapeHtml(sample.name || '素材')}</div>
                      </div>
                    `;
                  })
                  .join('')}
              </div>
              <div class="material-tag-section">
                <div class="material-tag-line">
                  ${(category.usageTags || []).slice(0, 6).map((tag) => `<span class="tag-chip tag-usage">${escapeHtml(tag)}</span>`).join('')}
                </div>
              </div>
            </div>
          `;
        })
        .join('')}
    </div>
  `;
}

function renderLibraries(payload = {}) {
  state.libraries = payload.libraries || [];
  state.libraryLayouts = payload.layouts || null;

  const select = byId('reference-library-select');
  if (select) {
    const current = state.selectedLibraryId || state.libraries[0]?.id || '';
    state.selectedLibraryId = current;
    select.innerHTML = ['<option value="">不使用已有参考库</option>']
      .concat(
        state.libraries.map(
          (item) => `<option value="${escapeHtml(item.id)}" ${item.id === current ? 'selected' : ''}>${escapeHtml(item.displayName || item.sourceName || item.name || item.id)}</option>`,
        ),
      )
      .join('');
  }

  const layoutSelect = byId('layout-set-select');
  if (layoutSelect) {
    const setOptions = payload.layouts?.setOptions || [];
    layoutSelect.innerHTML = setOptions
      .map(
        (item, index) => `<option value="${escapeHtml(item.id)}" ${item.id === payload.layouts?.activeSet ? 'selected' : ''}>${escapeHtml(normalizeLayoutSetLabel(item, index))}</option>`,
      )
      .join('');
  }

  renderLibrarySummary();
  if ((payload.layouts?.slideLayouts || []).length) {
    renderLayoutSelectors(payload.layouts?.slideLayouts || [], payload.layouts?.initialSelection || {});
  }
}

function renderWorkflowStages(stages = []) {
  const fallbackStages = (!stages.length && state.currentMeta)
    ? [
        {
          title: state.currentMeta.currentStageLabel || '当前阶段',
          detail: state.currentMeta.message || '后台正在处理，请稍候。',
        },
      ]
    : stages;
  setHtml(
    'workflow-stages',
    (fallbackStages || [])
      .map(
        (stage) => `
          <div class="stage-card">
            <strong>${escapeHtml(stage.title || stage.id || '阶段')}</strong>
            <div>${escapeHtml(stage.detail || '')}</div>
          </div>
        `,
      )
      .join('') || '<div class="status-item"><strong>暂无流程信息</strong><div>生成草稿后会在这里显示详细阶段。</div></div>',
  );
}

function renderDocumentSummary(summary = null) {
  const container = byId('document-summary');
  if (!container) return;
  if (!summary) {
    container.innerHTML = '<div class="status-item"><strong>暂无文档摘要</strong><div>草稿生成后会自动载入文档统计和标题结构。</div></div>';
    return;
  }

  const counts = summary.counts || {};
  const sections = summary.sectionTree || [];
  container.innerHTML = `
    <div class="status-item"><strong>正文统计</strong><div>段落 ${counts.paragraphs || 0}｜表格 ${counts.tables || 0}｜图片 ${counts.images || 0}</div></div>
    <div class="status-item"><strong>标题结构</strong><div>${escapeHtml((sections || []).map((item) => item.title).filter(Boolean).slice(0, 8).join(' ｜ ') || '未识别到一级标题')}</div></div>
  `;
}

function renderDownloadGrid(containerId, items = [], emptyText = '暂无可下载文件') {
  const container = byId(containerId);
  if (!container) return;
  if (!items.length) {
    container.innerHTML = `<div class="status-item"><strong>${escapeHtml(emptyText)}</strong></div>`;
    return;
  }
  container.innerHTML = items
    .map(
      (item) => `
        <div class="download-card">
          <strong>${escapeHtml(item.title || item.type || '下载项')}</strong>
          <a class="download-button" href="${item.url}" download>下载</a>
        </div>
      `,
    )
    .join('');
}

function renderArchiveSummary(meta = {}, manifest = null) {
  const container = byId('archive-summary');
  if (!container) return;
  if (!meta.archivedAt) {
    container.innerHTML = '<div class="status-item"><strong>暂无归档信息</strong><div>最终 PPT 生成后，这里会显示独立归档路径和下载文件列表。</div></div>';
    return;
  }

  const outputDir = meta.deliveryDir ? `${meta.deliveryDir}/output` : '';
  const items = Array.isArray(manifest?.files) ? manifest.files : [];
  container.innerHTML = `
    <div class="status-item">
      <strong>已归档路径</strong>
      <div>${escapeHtml(outputDir || '未记录归档路径')}</div>
    </div>
    <div class="status-item">
      <strong>可下载文件</strong>
      <div>${escapeHtml(items.length ? items.map((item) => item.title).join('｜') : '暂无可下载文件')}</div>
    </div>
  `;
}

function renderPreviewGrid(containerId, items = [], emptyText = '暂无预览') {
  const container = byId(containerId);
  if (!container) return;
  if (!items.length) {
    container.innerHTML = `<div class="status-item"><strong>${escapeHtml(emptyText)}</strong><div>如果预览仍为空，请先等待后台导出完成。</div></div>`;
    return;
  }
  container.innerHTML = items
    .map(
      (item) => `
        <div class="preview-card">
          <img src="${item.url}" alt="${escapeHtml(item.title || '')}" loading="lazy" />
          <div class="preview-meta">
            <strong>${escapeHtml(item.title || `第${item.page}页`)}</strong>
          </div>
        </div>
      `,
    )
    .join('');
}

function previewStateLabel(state = {}) {
  const status = String(state.status || '').toLowerCase();
  if (status === 'pending') return '待检测';
  if (status === 'available') return '可用';
  if (status === 'unsupported') return '不可用';
  if (status === 'failed') return '失败';
  return '未知';
}

function renderPreviewStatus(containerId, state = null, emptyText = '暂无预览状态') {
  const container = byId(containerId);
  if (!container) return;
  if (!state) {
    container.innerHTML = `<div class="status-item preview-status-card"><strong>预览状态待检测</strong><div>${escapeHtml(emptyText)}</div></div>`;
    return;
  }

  const status = String(state.status || '').toLowerCase();
  const platform = String(state.platform || '').trim() || '未知平台';
  const reason = String(state.reason || '').trim();
  const previewCount = Number(state.previewCount || 0);
  const detail =
    status === 'available'
      ? `已生成 ${previewCount} 张真实渲染图。`
      : reason || '当前预览不可用。';

  container.innerHTML = `
    <div class="status-item preview-status-card preview-status-${escapeHtml(status || 'unknown')}">
      <strong>预览环境：${escapeHtml(previewStateLabel(state))}</strong>
      <div>${escapeHtml(detail)}</div>
      <div>${escapeHtml(`平台：${platform}｜${status === 'available' ? `预览数量：${previewCount}` : '已跳过预览导出'}`)}</div>
    </div>
  `;
}

function renderPreviewSection(prefix, state = null, items = [], emptyText = '暂无预览') {
  renderPreviewStatus(`${prefix}-preview-status`, state, emptyText);
  const grid = byId(`${prefix}-preview-grid`);
  if (!grid) return;
  const status = String(state?.status || '').toLowerCase();
  if (status === 'unsupported' || status === 'failed') {
    grid.style.display = 'none';
    grid.innerHTML = '';
    return;
  }
  grid.style.display = '';
  renderPreviewGrid(`${prefix}-preview-grid`, items, emptyText);
}

function fallbackPreviewEnvironment(meta = {}) {
  const env = meta.previewEnvironment || null;
  if (!env) return null;
  return {
    ...env,
    previewCount: Number(env.previewCount || 0),
  };
}

function effectiveLayoutPayload(explicitPayload = null) {
  if (explicitPayload?.slideLayouts?.length) return explicitPayload;
  if (state.layoutData?.slideLayouts?.length) return state.layoutData;
  if (state.currentMeta?.layoutOptions?.slideLayouts?.length) return state.currentMeta.layoutOptions;
  return explicitPayload || state.layoutData || state.currentMeta?.layoutOptions || null;
}

function renderLayoutSelectors(slides = [], selection = {}) {
  const container = byId('layout-selectors');
  if (!container) return;
  if (!slides.length) {
    const fallback = effectiveLayoutPayload();
    if ((fallback?.slideLayouts || []).length) {
      renderLayoutSelectors(fallback.slideLayouts, fallback.initialSelection || {});
      return;
    }
    const pageCount = Number(fallback?.slideLayouts?.length || state.currentMeta?.pageCount || 0);
    const activeSet = friendlyLayoutSetLabel(fallback?.activeSet || state.currentMeta?.layoutSet || '') || '默认布局集';
    container.innerHTML = `
      <div class="status-item">
        <strong>${pageCount ? '布局规划已生成' : '暂无布局信息'}</strong>
        <div>${escapeHtml(pageCount ? `当前共有 ${pageCount} 页布局规划，使用布局集：${activeSet}。若下方仍未列出每页选项，请等待草稿清单回写完成。` : '生成草稿后会在这里显示每页可切换的模板。')}</div>
      </div>
    `;
    return;
  }

  state.currentLayouts = { ...(selection || {}) };
  container.innerHTML = slides
    .map((slide) => {
      const current = state.currentLayouts[String(slide.page)] || slide.currentTemplate || slide.recommendedTemplate || '';
      return `
        <div class="status-item">
          <strong>第${slide.page}页 · ${escapeHtml(slide.typeLabel || friendlyPageTypeLabel(slide.type) || '页面')}</strong>
          <div>${escapeHtml(slide.title || '')}</div>
          <select class="layout-page-select" data-page="${slide.page}">
            ${(slide.options || [])
              .map(
                (option, index) => `<option value="${escapeHtml(option.id)}" ${option.id === current ? 'selected' : ''}>${escapeHtml(normalizeTemplateOptionLabel(option, index))}</option>`,
              )
              .join('')}
          </select>
          <div>${escapeHtml(slide.description || slide.note || '用于拉开不同页面的版式差异。')}</div>
        </div>
      `;
    })
    .join('');

  container.querySelectorAll('.layout-page-select').forEach((select) => {
    select.addEventListener('change', (event) => {
      state.currentLayouts[String(event.target.dataset.page)] = event.target.value;
    });
  });
}

async function fetchJson(url) {
  const response = await fetch(url, { cache: 'no-store' });
  if (!response.ok) throw new Error(await response.text());
  return response.json();
}

async function fetchText(url) {
  const response = await fetch(url, { cache: 'no-store' });
  if (!response.ok) throw new Error(await response.text());
  return response.text();
}

async function maybeFetchJson(url) {
  try {
    const text = await fetchText(url);
    return JSON.parse(String(text || '').replace(/^\uFEFF/, ''));
  } catch {
    return null;
  }
}

async function maybeFetchText(url) {
  try {
    return await fetchText(url);
  } catch {
    return '';
  }
}

function getSemanticField(name) {
  return document.querySelector(`[data-semantic-field="${name}"]`);
}

function setFieldVisibility(field, visible) {
  if (!field) return;
  field.style.display = visible ? '' : 'none';
  field.querySelectorAll('input, select, textarea').forEach((node) => {
    node.disabled = !visible;
  });
}

function syncSemanticConfigVisibility() {
  const provider = document.querySelector('[name="semanticProvider"]')?.value || '';
  const modelInput = document.querySelector('[name="semanticModel"]');
  const apiKeyInput = document.querySelector('[name="semanticApiKey"]');

  setFieldVisibility(getSemanticField('baseUrl'), provider === 'local' || provider === 'custom');
  setFieldVisibility(getSemanticField('model'), provider === 'local' || provider === 'openai' || provider === 'custom');
  setFieldVisibility(getSemanticField('apiKey'), provider === 'minimax' || provider === 'openai' || provider === 'custom');
  setFieldVisibility(getSemanticField('supportsImages'), provider === 'local' || provider === 'minimax' || provider === 'openai' || provider === 'custom');

  if (provider === 'local' && modelInput && !String(modelInput.value || '').trim()) modelInput.value = 'deepseek-r1:14b';
  if (provider === 'minimax' && modelInput) {
    const current = String(modelInput.value || '').trim();
    if (!current || /MiniMax-M2\.[0-9]/i.test(current)) {
      modelInput.value = 'MiniMax-M2.7';
    }
  }
  if (provider === 'minimax' && apiKeyInput) apiKeyInput.placeholder = '请输入 Minimax API Key';
}

function resetOutputs() {
  renderStatusItems([]);
  renderWorkflowStages([]);
  renderDocumentSummary(null);
  renderDownloadGrid('draft-download-links', [], '暂无草稿文件');
  renderPreviewSection('draft', null, [], '草稿预览尚未生成');
  renderArchiveSummary({}, null);
  renderDownloadGrid('archive-download-links', [], '暂无归档文件');
  renderDownloadGrid('download-links', [], '暂无最终文件');
  renderPreviewSection('final', null, [], '最终预览尚未生成');
  renderLayoutSelectors([], {});
  byId('outline-editor').value = '';
  byId('style-editor').value = '';
  byId('notes-viewer').textContent = '';
  byId('generate-button').disabled = true;
}

async function loadArtifacts(sessionId) {
  if (!sessionId || state.loadingArtifacts) return;
  state.loadingArtifacts = true;
  try {
  const archived = Boolean(state.currentMeta?.archivedAt);
    let draftOutline = '';
    let draftStyle = '';
    let draftNotes = '';
    let draftSummary = null;
    let draftLayout = null;
    let draftPreviews = { items: [] };
    let finalPreviews = { items: [] };
    let archiveManifest = null;
    const draftPreviewState = archived
      ? state.currentMeta?.intermediate?.previewState || state.currentMeta?.previewState || fallbackPreviewEnvironment(state.currentMeta || {})
      : state.currentMeta?.intermediate?.previewState || state.currentMeta?.previewState || fallbackPreviewEnvironment(state.currentMeta || {});
    const finalPreviewState =
      state.currentMeta?.output?.previewState || state.currentMeta?.previewState || fallbackPreviewEnvironment(state.currentMeta || {});

    if (archived) {
      [draftOutline, draftStyle, draftNotes, draftSummary, draftLayout, finalPreviews, archiveManifest] = await Promise.all([
        maybeFetchText(`/api/download/${sessionId}/outline`),
        maybeFetchText(`/api/download/${sessionId}/style`),
        maybeFetchText(`/api/download/${sessionId}/notes`),
        maybeFetchJson(`/api/download/${sessionId}/summary`),
        maybeFetchJson(`/api/download/${sessionId}/layout`),
        fetchJson(`/api/workflow/${sessionId}/previews/final`).catch(() => ({ items: [] })),
        maybeFetchJson(`/api/download/${sessionId}/archive-manifest`),
      ]);
    } else {
      [draftOutline, draftStyle, draftNotes, draftSummary, draftLayout, draftPreviews] = await Promise.all([
        maybeFetchText(`/api/download/${sessionId}/outline-draft`),
        maybeFetchText(`/api/download/${sessionId}/style-draft`),
        maybeFetchText(`/api/download/${sessionId}/notes-draft`),
        maybeFetchJson(`/api/download/${sessionId}/summary-draft`),
        maybeFetchJson(`/api/download/${sessionId}/layout-draft`),
        fetchJson(`/api/workflow/${sessionId}/previews/draft`).catch(() => ({ items: [] })),
      ]);
      finalPreviews = await fetchJson(`/api/workflow/${sessionId}/previews/final`).catch(() => ({ items: [] }));
    }

    state.outlineText = draftOutline || '';
    state.styleText = draftStyle || '';
    state.notesText = draftNotes || '';
    state.layoutData = draftLayout || state.currentMeta?.layoutOptions || null;

    byId('outline-editor').value = state.outlineText;
    byId('style-editor').value = state.styleText;
    byId('notes-viewer').textContent = state.notesText || '暂无备注';

    renderDocumentSummary(draftSummary);
    renderPreviewSection(
      'draft',
      draftPreviewState,
      archived ? [] : draftPreviews.items || [],
      archived ? '草稿工作区已销毁。' : '首轮草稿已输出。若需要最终渲染图，请点击“生成 PPT”。',
    );
    renderPreviewSection('final', finalPreviewState, finalPreviews.items || [], '最终 PPT 尚未生成。');
    renderArchiveSummary(state.currentMeta || {}, archiveManifest);
    renderDownloadGrid(
      'archive-download-links',
      archived
        ? [
            { title: '归档清单', url: `/api/download/${sessionId}/archive-manifest` },
            ...((archiveManifest?.files || []).map((item) => ({ title: item.title || item.type || '文件', url: item.url || '' }))),
          ]
        : [],
      archived ? '暂无归档文件' : '暂无归档文件',
    );

      const effectiveLayout = effectiveLayoutPayload(draftLayout || null);
      renderLayoutSelectors(effectiveLayout?.slideLayouts || [], effectiveLayout?.initialSelection || {});

    byId('generate-button').disabled = !state.outlineText || !state.styleText;
  } finally {
    state.loadingArtifacts = false;
  }
}

function updateProgress(meta = {}) {
  const progress = meta.progress ?? 0;
  const label = meta.currentStageLabel || '处理中';
  const detail = meta.message || '后台正在执行工作流。';
  renderProgressBar('processing-progress', progress, label, detail, meta.status === 'error' ? 'error' : progress >= 100 ? 'done' : 'active');

  const generationDone = Boolean(meta.output?.deckPath);
  if (generationDone) {
    renderProgressBar('generation-progress', 100, '生成完成', '最终 PPT 与真实渲染图已生成，可直接下载。', 'done');
  } else if (state.generating) {
    renderProgressBar('generation-progress', 40, '生成 PPT', '正在根据草稿和布局选择渲染最终 PPT。', 'active');
  } else {
    renderProgressBar('generation-progress', 0, '生成 PPT', '确认草稿无误后，可点击“生成 PPT”。', 'active');
  }
}

function draftDownloadManifest(sessionId) {
  const items = [
    ['draft-deck', '草稿 PPT'],
    ['outline-draft', '草稿 Outline'],
    ['style-draft', '草稿 Style'],
    ['layout-draft', '布局规划'],
    ['notes-draft', '草稿备注'],
    ['summary-draft', '文档摘要'],
  ];
  return items.map(([type, title]) => ({ title, url: `/api/download/${sessionId}/${type}` }));
}

function finalDownloadManifest(sessionId) {
  const items = [
    ['deck', '最终 PPT'],
    ['outline', '最终 Outline'],
    ['style', '最终 Style'],
    ['layout', '最终布局'],
    ['notes', '最终备注'],
    ['semantic', '语义分析'],
    ['semantic-review', '效果复核'],
  ];
  return items.map(([type, title]) => ({ title, url: `/api/download/${sessionId}/${type}` }));
}

function renderMeta(meta = {}) {
  const mergedMeta = { ...(meta || {}), ...(meta.result || {}) };
  state.currentMeta = mergedMeta;
  updateProgress(mergedMeta);
  renderSemanticStatus(mergedMeta);
  renderStatusItems(buildStatusCards(mergedMeta));
  renderWorkflowStages(mergedMeta.workflowStages || []);
    renderPreviewSection(
      'draft',
      mergedMeta.archivedAt ? null : mergedMeta.intermediate?.previewState || mergedMeta.previewState || fallbackPreviewEnvironment(mergedMeta),
      [],
      mergedMeta.archivedAt ? '草稿工作区已销毁。' : '草稿预览尚未生成',
    );
    renderPreviewSection('final', mergedMeta.output?.previewState || mergedMeta.previewState || fallbackPreviewEnvironment(mergedMeta), [], '最终预览尚未生成');
  if (!state.loadingArtifacts) {
    const effectiveLayout = effectiveLayoutPayload(mergedMeta.layoutOptions || null);
    renderLayoutSelectors(effectiveLayout?.slideLayouts || [], effectiveLayout?.initialSelection || {});
  }
  if (mergedMeta.sessionId) {
    const draftFilesAvailable = Boolean(mergedMeta.intermediate?.outlinePath && !mergedMeta.archivedAt);
    renderDownloadGrid(
      'draft-download-links',
      draftFilesAvailable ? draftDownloadManifest(mergedMeta.sessionId) : [],
      mergedMeta.archivedAt ? '草稿文件已销毁' : '暂无草稿文件',
    );
    renderArchiveSummary(mergedMeta, null);
    renderDownloadGrid(
      'archive-download-links',
      mergedMeta.archivedAt
        ? [
            { title: '归档清单', url: `/api/download/${mergedMeta.sessionId}/archive-manifest` },
            { title: '最终 PPT', url: `/api/download/${mergedMeta.sessionId}/deck` },
            { title: '最终 Outline', url: `/api/download/${mergedMeta.sessionId}/outline` },
            { title: '最终 Style', url: `/api/download/${mergedMeta.sessionId}/style` },
            { title: '最终布局', url: `/api/download/${mergedMeta.sessionId}/layout` },
            { title: '文档摘要', url: `/api/download/${mergedMeta.sessionId}/summary` },
            { title: '标题结构图', url: `/api/download/${mergedMeta.sessionId}/structure` },
            { title: 'Markdown 结构稿', url: `/api/download/${mergedMeta.sessionId}/structure-md` },
            { title: '最终备注', url: `/api/download/${mergedMeta.sessionId}/notes` },
          ]
        : [],
      mergedMeta.archivedAt ? '暂无归档文件' : '暂无归档文件',
    );
    renderDownloadGrid('download-links', mergedMeta.output?.deckPath ? finalDownloadManifest(mergedMeta.sessionId) : [], '暂无最终文件');
  }
}

async function pollSession(sessionId) {
  if (!sessionId) return;
  if (state.pollingTimer) clearInterval(state.pollingTimer);

  const tick = async () => {
    try {
      const meta = await fetchJson(`/api/workflow/${sessionId}`);
      renderMeta(meta);
      await loadArtifacts(sessionId);
      if (meta.status === 'completed') {
        try {
          const payload = await fetchJson('/api/libraries');
          renderLibraries(payload);
        } catch {}
      }
      if (meta.output?.deckPath && state.generating) state.generating = false;
      if (meta.status === 'completed' || meta.status === 'error' || meta.status === 'failed') {
        clearInterval(state.pollingTimer);
        state.pollingTimer = null;
      }
    } catch (error) {
      renderStatusItems([{ title: '会话状态读取失败', body: error.message || String(error) }]);
      clearInterval(state.pollingTimer);
      state.pollingTimer = null;
    }
  };

  await tick();
  state.pollingTimer = setInterval(tick, 2500);
}

async function createWorkflow(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const formData = new FormData(form);
  resetOutputs();
  renderProgressBar('processing-progress', 10, '正在创建会话', '材料已提交，后台正在开始解析。', 'active');

  try {
    const response = await fetch('/api/workflow/create', {
      method: 'POST',
      body: formData,
    });
    const raw = await response.text();
    let payload = {};
    try {
      payload = raw ? JSON.parse(raw) : {};
    } catch {
      payload = { error: raw || '创建工作流失败' };
    }
    if (!response.ok) throw new Error(payload.error || raw || '创建工作流失败');
    state.sessionId = payload.sessionId;
    renderMeta(payload);
    await pollSession(payload.sessionId);
  } catch (error) {
    renderProgressBar('processing-progress', 100, '处理失败', error.message || String(error), 'error');
    renderStatusItems([{ title: '工作流创建失败', body: error.message || String(error) }]);
  }
}

async function generateFinalDeck() {
  if (!state.sessionId || state.generating) return;
  const outlineEditor = byId('outline-editor');
  const styleEditor = byId('style-editor');
  let outline;
  let style;
  try {
    outline = JSON.parse(outlineEditor.value || '{}');
    style = JSON.parse(styleEditor.value || '{}');
  } catch (error) {
    renderStatusItems([{ title: 'JSON 解析失败', body: '请先修正 Outline 或 Style JSON 的格式，再继续生成。' }]);
    return;
  }

  state.generating = true;
  updateProgress(state.currentMeta || {});

  try {
    const response = await fetch('/api/workflow/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sessionId: state.sessionId,
        outline,
        style,
        layoutSelection: state.currentLayouts,
        layoutSet: byId('layout-set-select')?.value || '',
      }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || '生成最终 PPT 失败');
    if (payload.sessionId) state.sessionId = payload.sessionId;
    await pollSession(state.sessionId);
  } catch (error) {
    state.generating = false;
    renderProgressBar('generation-progress', 100, '生成失败', error.message || String(error), 'error');
    renderStatusItems([{ title: '最终 PPT 生成失败', body: error.message || String(error) }]);
  }
}

function applyJsonEditor(sourceId) {
  const source = byId(sourceId);
  if (!source) return;
  try {
    JSON.parse(source.value || '{}');
    renderStatusItems([{ title: 'JSON 已更新', body: `${sourceId === 'outline-editor' ? 'Outline' : 'Style'} JSON 已通过格式校验，可继续生成。` }]);
  } catch {
    renderStatusItems([{ title: 'JSON 格式有误', body: '请修正 JSON 语法后再继续。' }]);
  }
}

function bindCopyButtons() {
  document.querySelectorAll('[data-copy]').forEach((button) => {
    button.addEventListener('click', async () => {
      const target = byId(button.dataset.copy);
      if (!target) return;
      try {
        await navigator.clipboard.writeText(target.value || target.textContent || '');
        renderStatusItems([{ title: '已复制', body: '内容已复制到剪贴板。' }]);
      } catch {
        renderStatusItems([{ title: '复制失败', body: '当前环境不支持自动复制，请手动复制。' }]);
      }
    });
  });
}

async function init() {
  resetOutputs();
  renderProgressBar('processing-progress', 0, '待开始', '请上传材料并点击“生成草稿”。', 'active');
  renderProgressBar('generation-progress', 0, '生成 PPT', '确认草稿无误后，可点击“生成 PPT”。', 'active');
  syncSemanticConfigVisibility();

  document.querySelector('[name="semanticProvider"]')?.addEventListener('change', syncSemanticConfigVisibility);
  byId('workflow-form')?.addEventListener('submit', createWorkflow);
  byId('generate-button')?.addEventListener('click', generateFinalDeck);
  byId('apply-outline-json')?.addEventListener('click', () => applyJsonEditor('outline-editor'));
  byId('apply-style-json')?.addEventListener('click', () => applyJsonEditor('style-editor'));
  byId('reference-library-select')?.addEventListener('change', (event) => {
    state.selectedLibraryId = event.target.value;
    renderLibrarySummary();
  });
  bindCopyButtons();

  try {
    const payload = await fetchJson('/api/libraries');
    renderLibraries(payload);
  } catch (error) {
    renderStatusItems([{ title: '初始化失败', body: error.message || String(error) }]);
  }
}

document.addEventListener('DOMContentLoaded', init);
