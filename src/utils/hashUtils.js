const fs = require("node:fs");
const crypto = require("node:crypto");

function hashBuffer(buffer) {
  return crypto.createHash("sha1").update(buffer).digest("hex");
}

function hashFile(filePath) {
  return hashBuffer(fs.readFileSync(filePath));
}

module.exports = {
  hashBuffer,
  hashFile,
};
