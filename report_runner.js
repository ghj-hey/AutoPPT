const runner = require("./src/cli/report_runner");

module.exports = runner;

if (require.main === module) {
  runner.main().catch((error) => {
    console.error(error.stack || error.message || error);
    process.exitCode = 1;
  });
}
