const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const { SESSION_ROOT, DELIVERABLE_ROOT } = require("../utils/pathConfig");
const { ensureDir, readJson, writeJson, copyFile } = require("../utils/fileUtils");
const { UPLOAD_ROOT } = require("../utils/pathConfig");

function safeExtension(fileName = "", fallback = "") {
  const ext = path.extname(String(fileName || "")).toLowerCase();
  return ext || fallback || "";
}

function buildStagedFileName(kind = "file", index = 0, originalName = "", fallbackExt = "") {
  const ext = safeExtension(originalName, fallbackExt);
  const suffix = Number.isFinite(index) && index > 0 ? `-${String(index).padStart(2, "0")}` : "";
  return `${kind}${suffix}${ext}`;
}

function createSession() {
  ensureDir(SESSION_ROOT);
  const sessionId = crypto.randomUUID();
  const sessionDir = path.join(SESSION_ROOT, sessionId);
  ensureDir(sessionDir);
  ensureDir(path.join(sessionDir, "input"));
  ensureDir(path.join(sessionDir, "intermediate"));
  ensureDir(path.join(sessionDir, "output"));
  return { sessionId, sessionDir };
}

function prepareFinalDeliveryWorkspace(sessionId) {
  ensureDir(DELIVERABLE_ROOT);
  const deliveryDir = path.join(DELIVERABLE_ROOT, sessionId);
  if (fs.existsSync(deliveryDir)) {
    fs.rmSync(deliveryDir, { recursive: true, force: true });
  }
  ensureDir(deliveryDir);
  const outputDir = path.join(deliveryDir, "output");
  ensureDir(outputDir);
  return { deliveryDir, outputDir };
}

function saveSessionMeta(sessionDir, meta) {
  const metaPath = path.join(sessionDir, "session.json");
  writeJson(metaPath, meta);
  return metaPath;
}

function updateSessionMeta(sessionDir, patch = {}) {
  const metaPath = path.join(sessionDir, "session.json");
  const current = readJson(metaPath, {});
  const next = {
    ...current,
    ...patch,
    updatedAt: new Date().toISOString(),
  };
  writeJson(metaPath, next);
  return next;
}

function loadSessionMeta(sessionId) {
  const sessionDir = path.join(SESSION_ROOT, sessionId);
  const metaPath = path.join(sessionDir, "session.json");
  const meta = readJson(metaPath, null);
  if (meta) {
    return {
      sessionDir,
      metaPath,
      meta,
      archived: false,
    };
  }

  const deliveryDir = path.join(DELIVERABLE_ROOT, sessionId);
  const deliveryMetaPath = path.join(deliveryDir, "session.json");
  const archivedMeta = readJson(deliveryMetaPath, null);
  return {
    sessionDir: deliveryDir,
    metaPath: deliveryMetaPath,
    meta: archivedMeta,
    archived: Boolean(archivedMeta),
  };
}

function destroySessionWorkspace(sessionDir) {
  if (!sessionDir) {
    return { destroyed: false, sessionDir: "" };
  }

  const resolvedSessionDir = path.resolve(sessionDir);
  const resolvedSessionRoot = path.resolve(SESSION_ROOT);
  const relative = path.relative(resolvedSessionRoot, resolvedSessionDir);
  const withinRoot = relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));

  if (!withinRoot || resolvedSessionDir === resolvedSessionRoot) {
    return { destroyed: false, sessionDir: resolvedSessionDir };
  }

  try {
    fs.rmSync(resolvedSessionDir, { recursive: true, force: true });
    return { destroyed: true, sessionDir: resolvedSessionDir };
  } catch {
    return { destroyed: false, sessionDir: resolvedSessionDir };
  }
}

function stageUploadedFile(sessionDir, file, targetName) {
  const destPath = path.join(sessionDir, "input", targetName || file.originalname);
  copyFile(file.path, destPath);
  const sourcePath = path.resolve(String(file.path || ""));
  const uploadIncomingRoot = path.resolve(path.join(UPLOAD_ROOT, "incoming"));
  if (sourcePath && sourcePath.startsWith(uploadIncomingRoot)) {
    try {
      fs.unlinkSync(sourcePath);
    } catch {
      // temp uploads are best-effort cleanup only
    }
  }
  return destPath;
}

module.exports = {
  createSession,
  prepareFinalDeliveryWorkspace,
  saveSessionMeta,
  updateSessionMeta,
  loadSessionMeta,
  destroySessionWorkspace,
  stageUploadedFile,
  buildStagedFileName,
};
