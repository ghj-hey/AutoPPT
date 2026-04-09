const test = require("node:test");
const assert = require("node:assert/strict");

const { loadLayoutLibrary } = require("../src/cli/report_runner");
const { buildLayoutOptions } = require("../src/services/layoutSelectionService");
const {
  buildReadableDownloadManifestCn,
  buildReadableFinalDownloadManifestCn,
  buildArchivedDeliveryManifestCn,
} = require("../src/services/workflowTextHelpers");
const { DEFAULT_LAYOUT_LIBRARY } = require("../src/utils/pathConfig");

function sampleTable(rows, cols) {
  return {
    header: Array.from({ length: cols }, (_, index) => `列${index + 1}`),
    rows: Array.from({ length: rows }, (_, rowIndex) =>
      Array.from({ length: cols }, (_, colIndex) => `R${rowIndex + 1}-${colIndex + 1}`),
    ),
  };
}

test("buildLayoutOptions keeps machine ids stable while exposing Chinese layout labels", () => {
  const layoutLibrary = loadLayoutLibrary(DEFAULT_LAYOUT_LIBRARY, "bank_finance_dynamic");
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
