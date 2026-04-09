const path = require("node:path");

const ROOT = path.resolve(__dirname, "..", "..");
const REFERENCE_ROOT = path.join(ROOT, "reference_library");
const WORKSPACE_ROOT = path.join(ROOT, "workspace");
const SESSION_ROOT = path.join(WORKSPACE_ROOT, "sessions");
const DELIVERABLE_ROOT = path.join(WORKSPACE_ROOT, "deliverables");
const UPLOAD_ROOT = path.join(WORKSPACE_ROOT, "uploads");
const CACHE_ROOT = path.join(WORKSPACE_ROOT, "cache");
const REFERENCE_PREVIEW_CACHE = path.join(CACHE_ROOT, "reference_ppt_images");
const DEFAULT_REFERENCE_LIBRARY = path.join(REFERENCE_ROOT, "work_meeting", "reusable_materials.json");
const DEFAULT_LAYOUT_LIBRARY = path.join(REFERENCE_ROOT, "work_meeting", "layout_templates.json");
const MASTER_REFERENCE_DIR = path.join(REFERENCE_ROOT, "master");
const LIBRARIES_ROOT = path.join(REFERENCE_ROOT, "libraries");

module.exports = {
  ROOT,
  REFERENCE_ROOT,
  WORKSPACE_ROOT,
  SESSION_ROOT,
  DELIVERABLE_ROOT,
  UPLOAD_ROOT,
  CACHE_ROOT,
  REFERENCE_PREVIEW_CACHE,
  DEFAULT_REFERENCE_LIBRARY,
  DEFAULT_LAYOUT_LIBRARY,
  MASTER_REFERENCE_DIR,
  LIBRARIES_ROOT,
};
