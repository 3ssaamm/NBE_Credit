/**
 * Auto-stamped code version metadata for the Health Check panel.
 * Updated automatically by tools/stamp_version.js before clasp push.
 */
function getCodeVersionInfo() {
  return {
    gitSha: "3d59ba4",
    dirty: true,
    stampedAt: "2026-11-20T20:06:43.521Z"
  };
}

function getCodeVersion() {
  const info = getCodeVersionInfo();
  const sha = (info && info.gitSha) ? String(info.gitSha).trim() : "unknown";
  return info && info.dirty ? (sha + "*") : sha;
}
