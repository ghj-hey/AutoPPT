const test = require("node:test");
const assert = require("node:assert/strict");

const { buildLayoutOptions } = require("../src/services/layoutSelectionService");
const {
  buildReadableDownloadManifestCn,
  buildReadableFinalDownloadManifestCn,
  buildArchivedDeliveryManifestCn,
} = require("../src/services/workflowTextHelpers");

function sampleTable(rows, cols) {
  return {
    header: Array.from({ length: cols }, (_, index) => `列${index + 1}`),
    rows: Array.from({ length: rows }, (_, rowIndex) =>
      Array.from({ length: cols }, (_, colIndex) => `R${rowIndex + 1}-${colIndex + 1}`),
    ),
  };
}

function createTestLayoutLibrary() {
  return {
    path: "/virtual/layout-library.json",
    version: "2.0.0",
    defaultSet: "bank_finance_dynamic",
    setName: "bank_finance_dynamic",
    setMetadata: {
      bank_finance_dynamic: {
        displayName: "动态组合型",
        description: "测试用布局库，覆盖主要页面类型。",
      },
    },
    defaultsByType: {
      cover: "cover_formal_v1",
      summary_cards: "summary_grid_v1",
      table_analysis: "table_compare_v1",
      process_flow: "process_three_lane_v1",
      bullet_columns: "bullet_dual_v1",
      image_story: "image_split_v1",
      action_plan: "action_timeline_v1",
      key_takeaways: "takeaway_cards_v1",
    },
    sets: {
      bank_finance_dynamic: {
        cover: "cover_formal_v1",
        summary_cards: "summary_grid_v1",
        table_analysis: "table_compare_v1",
        process_flow: "process_three_lane_v1",
        bullet_columns: "bullet_dual_v1",
        image_story: "image_split_v1",
        action_plan: "action_timeline_v1",
        key_takeaways: "takeaway_cards_v1",
      },
    },
    set: {
      cover: "cover_formal_v1",
      summary_cards: "summary_grid_v1",
      table_analysis: "table_compare_v1",
      process_flow: "process_three_lane_v1",
      bullet_columns: "bullet_dual_v1",
      image_story: "image_split_v1",
      action_plan: "action_timeline_v1",
      key_takeaways: "takeaway_cards_v1",
    },
    templates: {
      cover_formal_v1: { pageType: "cover", variant: "formal", displayName: "正式封面", family: "cover" },
      cover_clean_v1: { pageType: "cover", variant: "clean", displayName: "简洁封面", family: "cover" },
      summary_grid_v1: { pageType: "summary_cards", variant: "grid", displayName: "摘要四卡", family: "summary" },
      summary_dense_v1: { pageType: "summary_cards", variant: "dense_grid", displayName: "高密度摘要", family: "summary" },
      summary_spread_v1: { pageType: "summary_cards", variant: "spread", displayName: "横向摘要铺陈", family: "summary" },
      table_compare_v1: { pageType: "table_analysis", variant: "compare", displayName: "对比表格", family: "compare_split" },
      table_split_v1: { pageType: "table_analysis", variant: "split", displayName: "拆分表格", family: "compare_split" },
      table_visual_v1: { pageType: "table_analysis", variant: "visual", displayName: "图表混排", family: "media_panel" },
      table_dashboard_v1: { pageType: "table_analysis", variant: "dashboard", displayName: "看板表格", family: "data_wall" },
      table_matrix_v1: { pageType: "table_analysis", variant: "matrix", displayName: "矩阵分析表格", family: "matrix_grid" },
      table_picture_v1: { pageType: "table_analysis", variant: "picture", displayName: "图片佐证表格", family: "media_panel" },
      table_highlight_v1: { pageType: "table_analysis", variant: "highlight", displayName: "重点高亮表格", family: "highlight_focus" },
      table_stack_v1: { pageType: "table_analysis", variant: "stack", displayName: "堆叠结论表格", family: "stacked_story" },
      image_split_v1: { pageType: "image_story", variant: "split", displayName: "图文分栏", family: "media_panel" },
      image_focus_v1: { pageType: "image_story", variant: "focus", displayName: "主图聚焦", family: "media_panel" },
      image_storyboard_v1: { pageType: "image_story", variant: "storyboard", displayName: "图文故事板", family: "media_panel" },
      image_gallery_v1: { pageType: "image_story", variant: "gallery", displayName: "图片画廊", family: "media_panel" },
      process_three_lane_v1: { pageType: "process_flow", variant: "three_lane", displayName: "三段流程", family: "process" },
      bullet_dual_v1: { pageType: "bullet_columns", variant: "dual", displayName: "双栏要点", family: "bullet" },
      action_timeline_v1: { pageType: "action_plan", variant: "timeline", displayName: "时间推进", family: "action" },
      takeaway_cards_v1: { pageType: "key_takeaways", variant: "cards", displayName: "结论卡片", family: "summary" },
    },
  };
}

test("buildLayoutOptions keeps machine ids stable while exposing Chinese layout labels", () => {
  const layoutLibrary = createTestLayoutLibrary();
  const outline = {
    slides: [
      {
        page: 1,
        type: "summary_cards",
        title: "经营摘要",
        cards: [{ body: "A" }, { body: "B" }, { body: "C" }, { body: "D" }],
      },
      {
        page: 2,
        type: "table_analysis",
        title: "规模分析",
        table: sampleTable(6, 4),
        insights: [{ body: "结论一" }, { body: "结论二" }, { body: "结论三" }],
        screenshots: [{ path: "/tmp/chart.png" }],
        bars: [{ value: 18 }, { value: 9 }],
      },
      {
        page: 3,
        type: "table_analysis",
        title: "结构分析",
        table: sampleTable(3, 2),
        insights: [{ body: "结论一" }],
        image: { path: "/tmp/mock.png" },
      },
      {
        page: 4,
        type: "table_analysis",
        title: "趋势分析",
        table: sampleTable(7, 3),
        metrics: [
          { label: "指标1", value: "10%" },
          { label: "指标2", value: "20%" },
          { label: "指标3", value: "30%" },
          { label: "指标4", value: "40%" },
        ],
      },
      {
        page: 5,
        type: "image_story",
        title: "场景呈现",
        image: { path: "/tmp/story.png" },
        cards: [{ body: "A" }, { body: "B" }],
      },
    ],
  };

  const options = buildLayoutOptions(layoutLibrary, outline, {}, {
    styleFamily: "visual-report",
    pageRhythm: "visual",
    tablePreference: "picture",
  });

  assert.equal(options.activeSet, "bank_finance_dynamic");
  assert.ok(options.setOptions.length > 0);
  assert.ok(options.setOptions.every((item) => /[\u4e00-\u9fff]/.test(item.displayName)));

  for (const slide of options.slideLayouts) {
    assert.ok(slide.typeLabel);
    assert.ok(slide.currentTemplate);
    assert.ok(slide.recommendedTemplate);
    assert.ok(slide.options.length > 0);
    assert.ok(slide.options.every((item) => item.id && /^[a-z0-9_]+_v\d+$/i.test(item.id)));
    assert.ok(slide.options.every((item) => /[\u4e00-\u9fff]/.test(item.displayName)));
    assert.ok(slide.options.every((item) => item.displayName !== item.id));
  }

  const tableSlides = options.slideLayouts.filter((slide) => slide.type === "table_analysis");
  assert.equal(tableSlides.length, 3);
  assert.ok(
    new Set(tableSlides.map((slide) => slide.recommendedTemplate)).size >= 2,
    "table_analysis pages should not all collapse to one template",
  );

  assert.deepEqual(
    Object.keys(options.initialSelection),
    options.slideLayouts.map((slide) => String(slide.page)),
  );
});

test("download manifests keep Chinese titles while archive links preserve stable machine types", () => {
  const sessionId = "session-demo";
  const draftManifest = buildReadableDownloadManifestCn(sessionId);
  const finalManifest = buildReadableFinalDownloadManifestCn(sessionId);
  const archiveManifest = buildArchivedDeliveryManifestCn({
    sessionId,
    archivedAt: "2026-04-09T02:30:00.000Z",
    deliveryDir: "/tmp/archive/session-demo",
    sourceSessionDir: "/tmp/session/session-demo",
    files: [
      { type: "deck", title: "最终 PPT", path: "/tmp/archive/session-demo/output/workflow_generated.pptx" },
      { type: "layout", title: "最终布局", path: "/tmp/archive/session-demo/output/layout.final.json" },
    ],
  });

  assert.ok(draftManifest.every((item) => /[\u4e00-\u9fffA-Za-z]/.test(item.title)));
  assert.ok(finalManifest.every((item) => /[\u4e00-\u9fffA-Za-z]/.test(item.title)));
  assert.ok(draftManifest.some((item) => item.type === "layout-draft" && item.title === "布局规划"));
  assert.ok(finalManifest.some((item) => item.type === "layout" && item.title === "最终布局"));
  assert.equal(archiveManifest.outputDir, "/tmp/archive/session-demo/output");
  assert.deepEqual(
    archiveManifest.files.map((item) => item.type),
    ["deck", "layout"],
  );
  assert.ok(archiveManifest.files.every((item) => item.url === `/api/download/${sessionId}/${item.type}`));
});
