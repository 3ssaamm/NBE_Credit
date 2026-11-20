const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

function computeVersionInfo(rootDir) {
  const cwd = rootDir || path.resolve(__dirname, "..");
  let gitSha = "unknown";
  try {
    gitSha = execSync("git rev-parse --short HEAD", { cwd, encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }).trim() || "unknown";
  } catch (_) {
    gitSha = "unknown";
  }

  let dirty = false;
  try {
    const statusRaw = execSync("git status --porcelain", { cwd, encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] });
    const filteredLines = statusRaw
      .split(/\r?\n/)
      .filter(line => {
        if (!line || !line.trim()) return false;
        const p = line.slice(3).trim().replace(/\\/g, "/");
        return p !== "src/Version.js";
      });
    dirty = filteredLines.length > 0;
  } catch (_) {
    dirty = false;
  }

  return {
    gitSha,
    dirty,
    stampedAt: new Date().toISOString()
  };
}

function formatVersionFile(info) {
  const safeInfo = info || { gitSha: "unknown", dirty: false, stampedAt: new Date().toISOString() };
  const sha = String(safeInfo.gitSha || "unknown").replace(/["\\]/g, "");
  const dirty = Boolean(safeInfo.dirty);
  const stampedAt = String(safeInfo.stampedAt || new Date().toISOString()).replace(/["\\]/g, "");

  return [
    "/**",
    " * Auto-stamped code version metadata for the Health Check panel.",
    " * Updated automatically by tools/stamp_version.js before clasp push.",
    " */",
    "function getCodeVersionInfo() {",
    "  return {",
    `    gitSha: "${sha}",`,
    `    dirty: ${dirty},`,
    `    stampedAt: "${stampedAt}"`,
    "  };",
    "}",
    "",
    "function getCodeVersion() {",
    "  const info = getCodeVersionInfo();",
    "  const sha = (info && info.gitSha) ? String(info.gitSha).trim() : \"unknown\";",
    "  return info && info.dirty ? (sha + \"*\") : sha;",
    "}",
    ""
  ].join("\n");
}

function main() {
  const rootDir = path.resolve(__dirname, "..");
  const versionPath = path.join(rootDir, "src", "Version.js");
  const force = process.argv.includes("--force");
  const info = computeVersionInfo(rootDir);

  if (!force && fs.existsSync(versionPath)) {
    const existing = fs.readFileSync(versionPath, "utf8");
    const shaMatch = existing.match(/gitSha:\s*"([^"]+)"/);
    const dirtyMatch = existing.match(/dirty:\s*(true|false)/);
    if (shaMatch && dirtyMatch) {
      const existingSha = shaMatch[1];
      const existingDirty = dirtyMatch[1] === "true";
      if (existingSha === info.gitSha && existingDirty === info.dirty) {
        return;
      }
    }
  }

  fs.mkdirSync(path.dirname(versionPath), { recursive: true });
  fs.writeFileSync(versionPath, formatVersionFile(info), "utf8");
}

if (require.main === module) {
  main();
}

module.exports = {
  computeVersionInfo,
  formatVersionFile
};
