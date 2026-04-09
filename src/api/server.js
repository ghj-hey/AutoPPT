const fs = require("node:fs");
const path = require("node:path");
const express = require("express");
const multer = require("multer");
const { startWorkflowSession, generateWorkflowDeck, buildLayoutOptions } = require("../services/workflowService");
const { listReferenceLibraries } = require("../services/referenceLibraryService");
const { UPLOAD_ROOT, DEFAULT_LAYOUT_LIBRARY, ROOT, REFERENCE_ROOT, WORKSPACE_ROOT } = require("../utils/pathConfig");
const { ensureDir } = require("../utils/fileUtils");
const { loadSessionMeta } = require("../services/workflowSessionService");

ensureDir(path.join(UPLOAD_ROOT, "incoming"));

const upload = multer({
  dest: path.join(UPLOAD_ROOT, "incoming"),
});

function listPreviewImages(previewDir) {
  if (!previewDir || !fs.existsSync(previewDir)) return [];
  return fs
    .readdirSync(previewDir)
    .filter((name) => /\.(png|jpg|jpeg|webp)$/i.test(name))
    .sort((left, right) => {
      const leftNum = Number((left.match(/(\d+)/) || [])[1] || 0);
      const rightNum = Number((right.match(/(\d+)/) || [])[1] || 0);
      return leftNum - rightNum;
    })
    .map((name, index) => ({
      page: index + 1,
      name,
      path: path.join(previewDir, name),
    }));
}

function createApp() {
  const app = express();
  app.use(express.json({ limit: "15mb" }));
  app.use(express.urlencoded({ extended: true, limit: "15mb" }));
  app.disable("etag");

  function isWithinRoot(filePath, rootDir) {
    const resolvedFile = path.resolve(filePath);
    const resolvedRoot = path.resolve(rootDir);
    const relative = path.relative(resolvedRoot, resolvedFile);
    return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
  }

  function sendFreshIndex(res) {
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");
    res.sendFile(path.join(ROOT, "src", "ui", "public", "index.html"));
  }

  app.get("/", (req, res) => {
    sendFreshIndex(res);
  });

  app.get("/index.html", (req, res) => {
    sendFreshIndex(res);
  });

  app.use(
    express.static(path.join(ROOT, "src", "ui", "public"), {
      etag: false,
      lastModified: false,
      maxAge: 0,
      setHeaders(res, filePath) {
        if (/\.html?$/i.test(filePath)) {
          res.setHeader("Content-Type", "text/html; charset=utf-8");
          res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
          res.setHeader("Pragma", "no-cache");
          res.setHeader("Expires", "0");
          return;
        }
        if (/\.js$/i.test(filePath)) {
          res.setHeader("Content-Type", "application/javascript; charset=utf-8");
          res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
          res.setHeader("Pragma", "no-cache");
          res.setHeader("Expires", "0");
          return;
        }
        if (/\.css$/i.test(filePath)) {
          res.setHeader("Content-Type", "text/css; charset=utf-8");
          res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
          res.setHeader("Pragma", "no-cache");
          res.setHeader("Expires", "0");
        }
      },
    }),
  );

  app.get("/api/health", (req, res) => {
    res.json({ ok: true, now: new Date().toISOString() });
  });

  app.get("/api/libraries", (req, res) => {
    res.json({
      libraries: listReferenceLibraries(),
      layouts: buildLayoutOptions(DEFAULT_LAYOUT_LIBRARY, null),
    });
  });

  app.get("/api/library-preview", (req, res) => {
    const rawPath = String(req.query.path || "");
    if (!rawPath) {
      res.status(404).json({ error: "Missing path" });
      return;
    }

    const filePath = path.resolve(rawPath);
    const allowedRoot = isWithinRoot(filePath, REFERENCE_ROOT) || isWithinRoot(filePath, WORKSPACE_ROOT);
    if (!allowedRoot || !fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
      res.status(404).json({ error: "Preview not found" });
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const mime = {
      ".png": "image/png",
      ".jpg": "image/jpeg",
      ".jpeg": "image/jpeg",
      ".gif": "image/gif",
      ".webp": "image/webp",
      ".svg": "image/svg+xml",
    }[ext];
    if (!mime) {
      res.status(415).json({ error: "Unsupported preview type" });
      return;
    }

    res.setHeader("Content-Type", mime);
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");
    fs.createReadStream(filePath)
      .on("error", () => {
        if (!res.headersSent) res.status(404).end();
      })
      .pipe(res);
  });

  app.post(
    "/api/workflow/create",
    upload.fields([
      { name: "materialPpt", maxCount: 6 },
      { name: "referencePpt", maxCount: 6 },
      { name: "referenceImages", maxCount: 12 },
      { name: "template", maxCount: 1 },
      { name: "requirementDoc", maxCount: 1 },
    ]),
    async (req, res) => {
      try {
        const payload = await startWorkflowSession({
          files: req.files || {},
          options: {
            referenceName: req.body.referenceName || "",
            referenceLibraryId: req.body.referenceLibraryId || "",
            mergeReference: req.body.mergeReference !== "false",
            department: req.body.department || "",
            presenter: req.body.presenter || "",
            date: req.body.reportDate || "",
            layoutSet: req.body.layoutSet || "",
            pageCount: Number(req.body.pageCount || 0) || 0,
            layoutLibraryPath: DEFAULT_LAYOUT_LIBRARY,
            semanticProvider: req.body.semanticProvider || "",
            semanticBaseUrl: req.body.semanticBaseUrl || "",
            semanticModel: req.body.semanticModel || "",
            semanticApiKey: req.body.semanticApiKey || "",
            semanticSupportsImages: req.body.semanticSupportsImages === "on" || req.body.semanticSupportsImages === "true",
          },
        });
        res.status(202).json(payload);
      } catch (error) {
        res.status(400).json({ error: error.message || String(error) });
      }
    },
  );

  app.post("/api/workflow/generate", async (req, res) => {
    try {
      const payload = await generateWorkflowDeck({
        sessionId: req.body.sessionId,
        outline: req.body.outline,
        style: req.body.style,
        layoutSelection: req.body.layoutSelection || {},
        layoutSet: req.body.layoutSet || "",
      });
      res.json(payload);
    } catch (error) {
      res.status(400).json({ error: error.message || String(error) });
    }
  });

  app.get("/api/workflow/:sessionId", (req, res) => {
    const session = loadSessionMeta(req.params.sessionId);
    if (!session.meta) {
      res.status(404).json({ error: "Session not found" });
      return;
    }
    res.json(session.meta);
  });

  app.get("/api/workflow/:sessionId/previews/:stage", (req, res) => {
    const session = loadSessionMeta(req.params.sessionId);
    if (!session.meta) {
      res.status(404).json({ error: "Session not found" });
      return;
    }

    const stage = String(req.params.stage || "").toLowerCase();
    const previewDir =
      stage === "final" ? session.meta.output?.previewDir : session.meta.intermediate?.previewDir;
    const previewState =
      stage === "final" ? session.meta.output?.previewState || session.meta.previewState || null : session.meta.intermediate?.previewState || session.meta.previewState || null;

    if (!previewDir || !path.isAbsolute(previewDir) || !isWithinRoot(previewDir, WORKSPACE_ROOT)) {
      res.json({ items: [], previewState });
      return;
    }

    const items = listPreviewImages(previewDir).map((item) => ({
      page: item.page,
      title: `第${item.page}页`,
      url: `/api/library-preview?path=${encodeURIComponent(item.path)}`,
      path: item.path,
    }));
    res.json({ items, previewState });
  });

  app.get("/api/download/:sessionId/:fileType", (req, res) => {
    const session = loadSessionMeta(req.params.sessionId);
    if (!session.meta) {
      res.status(404).json({ error: "Session not found" });
      return;
    }
      const fileMap = {
        deck: session.meta.output?.deckPath,
        outline: session.meta.output?.outlinePath || session.meta.intermediate?.outlinePath,
        style: session.meta.output?.stylePath || session.meta.intermediate?.stylePath,
        structure: session.meta.output?.structurePath || session.meta.intermediate?.documentStructurePath,
        "structure-md": session.meta.output?.structureMarkdownPath || session.meta.intermediate?.documentStructureMarkdownPath,
        layout: session.meta.output?.layoutPath || session.meta.intermediate?.layoutOptionsPath,
        notes: session.meta.output?.notesPath || session.meta.intermediate?.notesPath,
        summary: session.meta.output?.summaryPath || session.meta.intermediate?.documentSummaryPath,
        "reference-style": session.meta.output?.referenceStylePath || session.meta.intermediate?.referenceStylePath,
        semantic: session.meta.output?.semanticAnalysisPath || session.meta.intermediate?.semanticAnalysisPath,
        "semantic-review": session.meta.output?.semanticReviewPath || session.meta.intermediate?.semanticReviewPath,
        "semantic-refined": session.meta.output?.semanticRefinedSelectionPath || session.meta.intermediate?.semanticRefinedSelectionPath,
        "archive-manifest": session.meta.output?.archivedManifestPath,
        "draft-deck": session.meta.intermediate?.draftDeckPath,
        "outline-draft": session.meta.intermediate?.outlinePath,
        "style-draft": session.meta.intermediate?.stylePath,
        "structure-draft": session.meta.intermediate?.documentStructurePath,
        "structure-md-draft": session.meta.intermediate?.documentStructureMarkdownPath,
        "layout-draft": session.meta.intermediate?.layoutOptionsPath,
        "notes-draft": session.meta.intermediate?.notesPath,
        "summary-draft": session.meta.intermediate?.documentSummaryPath,
        "reference-draft": session.meta.intermediate?.referenceSummaryPath,
        "reference-style-draft": session.meta.intermediate?.referenceStylePath,
        "semantic-draft": session.meta.intermediate?.semanticAnalysisPath,
        "semantic-review-draft": session.meta.intermediate?.semanticReviewPath,
        "semantic-refined-draft": session.meta.intermediate?.semanticRefinedSelectionPath,
      };
    const filePath = fileMap[req.params.fileType];
    if (!filePath) {
      res.status(404).json({ error: "Unknown file type" });
      return;
    }
    if (!path.isAbsolute(filePath)) {
      res.status(404).json({ error: "File path is invalid" });
      return;
    }
    res.download(filePath, (error) => {
      if (error && !res.headersSent) {
        res.status(404).json({ error: "Download file not found" });
      }
    });
  });

  return app;
}

function startServer(port = Number(process.env.PORT || 3210)) {
  const app = createApp();
  app.listen(port, () => {
    console.log(`Workflow UI running at http://127.0.0.1:${port}`);
  });
}

if (require.main === module) {
  startServer();
}

module.exports = {
  createApp,
  startServer,
};
