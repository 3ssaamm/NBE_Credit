const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const testsDir = __dirname;
const testFiles = fs.readdirSync(testsDir)
  .filter(f => f.endsWith(".test.js"))
  .sort();

let failed = 0;
console.log(`Running ${testFiles.length} test suite(s)...`);

for (const file of testFiles) {
  const fullPath = path.join(testsDir, file);
  try {
    const out = execFileSync(process.execPath, [fullPath], { encoding: "utf8" });
    const summaryLine = out.trim().split(/\r?\n/).pop() || "passed";
    console.log(`  PASS  ${file} — ${summaryLine}`);
  } catch (err) {
    failed++;
    console.error(`  FAIL  ${file}`);
    if (err.stdout) console.error(err.stdout.trim());
    if (err.stderr) console.error(err.stderr.trim());
  }
}

if (failed > 0) {
  console.error(`\n${failed} of ${testFiles.length} test suite(s) failed.`);
  process.exit(1);
}

console.log(`\nAll ${testFiles.length} test suite(s) passed.`);
