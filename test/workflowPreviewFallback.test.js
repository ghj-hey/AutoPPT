const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const {
  buildRenderedPreviews,
} = require("../src/services/powerPointPreviewService");
const {
  normalizePreviewState,
  previewStatusMessage,
} = require("../src/services/workflowService");

function makeTempDir(prefix) {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

function writeTempFile(dir, name, content = "preview") {
  const filePath = path.join(dir, name);
  fs.writeFileSync(filePath, content);
  return filePath;
}

test("buildRenderedPreviews skips unsupported environments without invoking exporter", async () => {
  let exporterCalled = false;
  const result = await buildRenderedPreviews("/tmp/demo.pptx", makeTempDir("ppt-preview-unsupported-"), ["第一页"], {
    previewSupport: {
      supported: false,
      status: "unsupported",
      reason: "当前运行环境不支持 PowerPoint 预览。",
      platform: "linux",
    },
    exportPresentationPngs: async () => {
      exporterCalled = true;
      return [];
    },
  });

  assert.equal(exporterCalled, false);
  assert.deepEqual(result.previews, []);
  assert.equal(result.previewState.supported, false);
  assert.equal(result.previewState.status, "unsupported");
  assert.match(result.previewState.reason, /不支持/);
});

test("buildRenderedPreviews converts exporter failures into a non-blocking failed state", async () => {
  const result = await buildRenderedPreviews("/tmp/demo.pptx", makeTempDir("ppt-preview-failed-"), ["第一页"], {
    previewSupport: {
      supported: true,
      status: "available",
      reason: "",
      platform: "win32",
    },
    exportPresentationPngs: async () => {
      throw new Error("boom");
    },
  });

  assert.deepEqual(result.previews, []);
  assert.equal(result.previewState.supported, true);
  assert.equal(result.previewState.status, "failed");
  assert.match(result.previewState.reason, /boom/);
});

test("buildRenderedPreviews returns preview images when the exporter succeeds", async () => {
  const outDir = makeTempDir("ppt-preview-success-");
  const imagePath = writeTempFile(outDir, "slide-1.png", "fake-png-data");

  const result = await buildRenderedPreviews("/tmp/demo.pptx", outDir, ["第一页"], {
    previewSupport: {
      supported: true,
      status: "available",
      reason: "",
      platform: "win32",
    },
    exportPresentationPngs: async () => [imagePath],
  });

  assert.equal(result.previewState.supported, true);
  assert.equal(result.previewState.status, "available");
  assert.equal(result.previewState.previewCount, 1);
  assert.equal(result.previews.length, 1);
  assert.equal(result.previews[0].title, "第一页");
  assert.match(result.previews[0].dataUrl, /^data:image\/png;base64,/);
});

test("workflow preview state normalization preserves readable fallback messaging", () => {
  const unsupported = normalizePreviewState(
    {
      previewState: {
        supported: false,
        status: "unsupported",
        reason: "当前环境不支持真实预览。",
        platform: "linux",
      },
    },
    "draft",
  );
  const failed = normalizePreviewState(
    {
      previewState: {
        supported: true,
        status: "failed",
        reason: "无法启动 powershell.exe。",
        platform: "win32",
      },
    },
    "final",
  );

  assert.equal(unsupported.supported, false);
  assert.equal(unsupported.status, "unsupported");
  assert.match(previewStatusMessage("草稿 PPT", unsupported), /已跳过预览导出/);

  assert.equal(failed.supported, true);
  assert.equal(failed.status, "failed");
  assert.match(previewStatusMessage("最终 PPT", failed), /预览导出失败/);
});
