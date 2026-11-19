const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

function main() {
  let raw = "";
  try {
    raw = fs.readFileSync(0, "utf8");
  } catch (_) {
    process.stdout.write("{}\n");
    return;
  }

  try {
    const payload = raw ? JSON.parse(raw) : {};
    if (payload.error) {
      process.stdout.write("{}\n");
      return;
    }

    const args = (payload.toolCall && payload.toolCall.args) || {};
    const targetPath = args.Path || args.path || args.file_path || "";
    if (!targetPath) {
      process.stdout.write("{}\n");
      return;
    }

    const defaultRoot = path.resolve(__dirname, "..", "..");
    const projectRoot = payload.cwd && fs.existsSync(path.join(payload.cwd, ".clasp.json"))
      ? payload.cwd
      : defaultRoot;

    if (!fs.existsSync(path.join(projectRoot, ".clasp.json"))) {
      process.stdout.write("{}\n");
      return;
    }

    const resolvedTarget = path.resolve(projectRoot, targetPath);
    const relPath = path.relative(projectRoot, resolvedTarget).replace(/\\/g, "/");
    const isAppsScriptFile =
      relPath === "Code.js" ||
      relPath === "appsscript.json" ||
      /^src\/.+\.(js|gs)$/i.test(relPath) ||
      (!relPath.startsWith("..") && /\.gs$/i.test(relPath));

    if (!isAppsScriptFile) {
      process.stdout.write("{}\n");
      return;
    }

    if ((relPath === "Code.js" || /^src\/.+\.(js|gs)$/i.test(relPath)) && fs.existsSync(resolvedTarget)) {
      execSync(`node -c "${resolvedTarget}"`, { cwd: projectRoot, stdio: "pipe", timeout: 5000 });
    } else if (relPath === "appsscript.json" && fs.existsSync(resolvedTarget)) {
      JSON.parse(fs.readFileSync(resolvedTarget, "utf8"));
    }

    if (fs.existsSync(path.join(projectRoot, "package.json"))) {
      execSync("npm test", { cwd: projectRoot, stdio: "pipe", timeout: 25000 });
    }

    execSync("clasp push -f", { cwd: projectRoot, stdio: "pipe", timeout: 25000 });
  } catch (_) {
    // Never block or break tool execution if syntax is mid-edit or clasp fails
  }

  process.stdout.write("{}\n");
}

main();
