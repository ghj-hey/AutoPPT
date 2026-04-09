const extractor = require("./src/cli/build_reference_library");

module.exports = extractor;

if (require.main === module) {
  extractor.main().catch((error) => {
    console.error(error.stack || error.message || error);
    process.exitCode = 1;
  });
}
