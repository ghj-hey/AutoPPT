const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");

const {
  prepareFinalDeliveryWorkspace,
  saveSessionMeta,
  loadSessionMeta,
  destroySessionWorkspace,
} = require("../src/services/workflowSessionService");
const { SESSION_ROOT, DELIVERABLE_ROOT } = require("../src/utils/pathConfig");
const { buildArchivedDeliveryManifestCn } = require("../src/services/workflowTextHelpers");

function removeTree(targetPath) {
  fs.rmSync(targetPath, { recursive: true, force: true });
}

test("prepareFinalDeliveryWorkspace creates an independent archive output root", () => {
  const sessionId = `session-${crypto.randomUUID()}`;
  const deliveryDir = path.join(DELIVERABLE_ROOT, sessionId);
  removeTree(deliveryDir);

  const result = prepareFinalDeliveryWorkspace(sessionId);

  assert.equal(result.deliveryDir, deliveryDir);
  assert.equal(path.basename(result.outputDir), "output");
  assert.equal(fs.existsSync(result.deliveryDir), true);
  assert.equal(fs.existsSync(result.outputDir), true);

  removeTree(deliveryDir);
});

test("loadSessionMeta falls back to archived deliverables after session destruction", () => {
  const sessionId = `session-${crypto.randomUUID()}`;
  const sessionDir = path.join(SESSION_ROOT, sessionId);
  const deliveryDir = path.join(DELIVERABLE_ROOT, sessionId);

  removeTree(sessionDir);
  removeTree(deliveryDir);
  fs.mkdirSync(sessionDir, { recursive: true });
  fs.mkdirSync(path.join(sessionDir, "intermediate"), { recursive: true });

  const archivedMeta = {
    sessionId,
    sessionDir: deliveryDir,
    sourceSessionDir: sessionDir,
    deliveryDir,
    archivedAt: new Date().toISOString(),
    status: "completed",
    output: {
      deckPath: path.join(deliveryDir, "output", "workflow_generated.pptx"),
    },
  };
  saveSessionMeta(deliveryDir, archivedMeta);

  const liveMeta = {
    sessionId,
    sessionDir,
    status: "running",
  };
  saveSessionMeta(sessionDir, liveMeta);

  let loaded = loadSessionMeta(sessionId);
  assert.equal(loaded.meta.status, "running");
  assert.equal(loaded.archived, false);

  destroySessionWorkspace(sessionDir);

  loaded = loadSessionMeta(sessionId);
  assert.equal(loaded.meta.status, "completed");
  assert.equal(loaded.archived, true);
  assert.equal(loaded.sessionDir, deliveryDir);

  removeTree(sessionDir);
  removeTree(deliveryDir);
});

test("destroySessionWorkspace removes only session-root directories", () => {
  const sessionId = `session-${crypto.randomUUID()}`;
  const sessionDir = path.join(SESSION_ROOT, sessionId);
  const deliveryDir = path.join(DELIVERABLE_ROOT, sessionId);

  removeTree(sessionDir);
  removeTree(deliveryDir);
  fs.mkdirSync(sessionDir, { recursive: true });
  fs.mkdirSync(deliveryDir, { recursive: true });

  const result = destroySessionWorkspace(sessionDir);

  assert.equal(result.destroyed, true);
  assert.equal(fs.existsSync(sessionDir), false);
  assert.equal(fs.existsSync(deliveryDir), true);

  removeTree(deliveryDir);
});

test("buildArchivedDeliveryManifestCn keeps archive metadata compact and linkable", () => {
  const manifest = buildArchivedDeliveryManifestCn({
    sessionId: "session-demo",
    archivedAt: "2026-04-09T02:30:00.000Z",
    deliveryDir: "/tmp/archive/session-demo",
    sourceSessionDir: "/tmp/session/session-demo",
    files: [
      { type: "deck", title: "最终 PPT", path: "/tmp/archive/session-demo/output/workflow_generated.pptx" },
      { type: "archive-manifest", title: "归档清单", path: "/tmp/archive/session-demo/output/archived_manifest.json" },
    ],
  });

  assert.equal(manifest.sessionId, "session-demo");
  assert.equal(manifest.outputDir, "/tmp/archive/session-demo/output");
  assert.equal(manifest.files.length, 2);
  assert.equal(manifest.files[0].url, "/api/download/session-demo/deck");
  assert.equal(manifest.files[1].title, "归档清单");
});
