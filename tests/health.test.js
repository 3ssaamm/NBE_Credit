const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { createGasSandbox, populateSpreadsheetFromFixture, MockSpreadsheet } = require("./helpers/gas_sandbox");
const { computeVersionInfo, formatVersionFile } = require("../tools/stamp_version");

const tz = "Africa/Cairo";

function colLetterToIdx(letters) {
  let res = 0;
  for (let i = 0; i < letters.length; i++) {
    res = res * 26 + (letters.charCodeAt(i) - 64);
  }
  return res;
}

// ==========================================
// Test 1: Real Fixture (2026-11-real.json) All-Green Health Panel & Layout Anchors
// ==========================================
{
  const realFixture = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "2026-11-real.json"), "utf8"));
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);
  populateSpreadsheetFromFixture(ss, realFixture, CONFIG);

  const dash = sandbox.updateLiveDashboard({ ss });
  assert.strictEqual(dash.blocked, false, "Healthy real fixture must not be blocked");
  assert.ok(dash.health, "updateLiveDashboard must return health model");
  assert.strictEqual(dash.health.allHealthy, true, "Real fixture must be HEALTHY");
  assert.strictEqual(dash.health.overallStatus, "HEALTHY");
  assert.strictEqual(dash.health.parsePassed, true, "Parse tie-out must pass");
  assert.strictEqual(dash.health.unresolvedCount, 0, "Unresolved bank lines must be 0");
  assert.strictEqual(dash.health.orphanCount, 0, "Active orphan logs must be 0");
  assert.strictEqual(dash.health.unlinkedInstCount, 0, "Unlinked installments must be 0");
  assert.strictEqual(dash.health.statementCheckOk, true, "Statement check within +-0.50 must pass");

  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  assert.ok(ar, "Active Bill & Reconciliation sheet must exist");

  // Verify Health Panel cells (Rows 9-11)
  const a9 = String(ar._getCell(9, 1) || "");
  const a10 = String(ar._getCell(10, 1) || "");
  const c10 = String(ar._getCell(10, 3) || "");
  const e10 = String(ar._getCell(10, 5) || "");
  const g10 = String(ar._getCell(10, 7) || "");
  const a11 = String(ar._getCell(11, 1) || "");
  const c11 = String(ar._getCell(11, 3) || "");
  const e11 = String(ar._getCell(11, 5) || "");
  const g11 = String(ar._getCell(11, 7) || "");

  assert.ok(a9.includes("🩺 Reconciliation Health Check") && a9.includes("All Checks Passed"), `Unexpected A9: ${a9}`);
  assert.strictEqual(a10, "Parse Tie-Out: ✅ Passed");
  assert.ok(c10.includes("Unresolved Bank Lines:"), `Unexpected C10: ${c10}`);
  assert.ok(e10.includes("Orphan Logs:"), `Unexpected E10: ${e10}`);
  assert.ok(g10.includes("Unlinked Installments:"), `Unexpected G10: ${g10}`);
  assert.ok(
    g10.includes("COUNTIF(") && g10.includes('"🚫*"') && g10.includes('"🔗*"') && !g10.includes('"<>🚫*"'),
    `G10 must subtract 🚫* and 🔗* rather than using <>🚫*: ${g10}`
  );
  assert.ok(a11.includes("Stmt Check:"), `Unexpected A11: ${a11}`);
  assert.ok(c11.includes("Carryover:") && c11.includes("Pending:"), `Unexpected C11: ${c11}`);
  assert.match(e11, /^Last Refresh: \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/, `Unexpected E11: ${e11}`);
  assert.strictEqual(g11, `Code Version: ${sandbox.getCodeVersion()}`, `Unexpected G11: ${g11}`);

  // Green backgrounds when healthy
  assert.strictEqual(ar.backgrounds["9_1"], "#e8f5e9");
  assert.strictEqual(ar.backgrounds["10_1"], "#e8f5e9");
  assert.strictEqual(ar.backgrounds["10_3"], "#e8f5e9");
  assert.strictEqual(ar.backgrounds["10_5"], "#e8f5e9");
  assert.strictEqual(ar.backgrounds["10_7"], "#e8f5e9");
  assert.strictEqual(ar.backgrounds["11_1"], "#e8f5e9");

  // Conditional format rules registered & no broken LEFT(,1) emoji surrogate check anywhere
  const cfRules = ar.getConditionalFormatRules();
  assert.ok(cfRules.length >= 6, `Expected >= 6 conditional format rules, got ${cfRules.length}`);
  for (let r = 1; r <= ar.getLastRow(); r++) {
    for (let c = 1; c <= 8; c++) {
      const val = String(ar._getCell(r, c) || "");
      assert.ok(!/LEFT\([^)]+,\s*1\)/.test(val), `Found forbidden LEFT(,1) formula at (${r},${c}): ${val}`);
    }
  }

  // Verify Overview anchors B5:H7 remain unshifted
  assert.strictEqual(ar._getCell(5, 2), "2026-10-31", "B5 must keep Statement Date");
  assert.strictEqual(ar._getCell(6, 2), "November 27, 2026", "B6 must keep Payment Due Date");
  assert.strictEqual(ar._getCell(7, 2), "November 2026", "B7 must keep Active Billing Cycle");
  assert.strictEqual(ar._getCell(5, 4), 32618.95, "D5 must keep Bank Closing Bill");
  assert.ok(String(ar._getCell(6, 4)).startsWith("=H"), "D6 must link Family Shares Total");
  assert.ok(String(ar._getCell(5, 6)).includes("Balanced"), "F5 must keep Recon Status formula");
  assert.ok(String(ar._getCell(6, 6)).includes("D5"), "F6 must keep Statement Check formula");
  assert.ok(String(ar._getCell(7, 8)).includes("ISBLANK(H6)"), "H7 must keep Payment Status formula");

  // Verify Table 1 moved cleanly to startShareRow = 15 & cross-sheet formulas point to $H$15:$H$...
  assert.strictEqual(sandbox.getActiveReconStartShareRow(), 15);
  assert.ok(String(ar._getCell(13, 1)).includes("Family Member Active Debt Shares"), "Row 13 must be Table 1 banner");
  assert.strictEqual(ar._getCell(14, 1), "#", "Row 14 must be Table 1 header");
  assert.strictEqual(ar._getCell(15, 1), 1, "Row 15 must be first family member row");

  const totShareRow = 15 + dash.ledger.personShares.length;
  const forecastSheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY_FORECAST);
  const fHeaders = forecastSheet.getRange(4, 1, 1, forecastSheet.getLastColumn()).getValues()[0];
  const activeColIdx = fHeaders.findIndex(h => String(h).includes("(Active)")) + 1;
  assert.ok(activeColIdx > 1, "Monthly History & Forecast must include Active cycle column");
  const activeFormula = String(forecastSheet._getCell(5, activeColIdx) || "");
  assert.ok(activeFormula.includes("$H$15:$H$"), `Forecast Active formula must reference $H$15:$H$: ${activeFormula}`);

  const memberSheet = ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
  const a6Formula = String(memberSheet._getCell(6, 1) || "");
  assert.ok(
    a6Formula.includes(`$H$${totShareRow}`) && a6Formula.includes("$H$15:$H$"),
    `Member Breakdown A6 formula must reference row 15 and total row ${totShareRow}: ${a6Formula}`
  );
}

// ==========================================
// Test 2: Warning States (Unresolved, Orphan, Unlinked Installment) & Live Recovery
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 450, "Total Debit:", 450],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 450]
  ]);
  stmt.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Post Date", "Description", "Type", "Amount", "Status"]]);
  stmt.getRange(7, 1, 2, 7).setValues([
    [1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "MYSTERY SHOP", "DEBIT", 150, ""],
    [2, new Date(2026, 9, 12, 12, 0, 0), new Date(2026, 9, 12, 12, 0, 0), "PHONE PLAN 2 OF 6", "INSTALLMENT", 300, ""]
  ]);

  const txSheet = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  txSheet.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  txSheet.getRange(2, 1, 1, 5).setValues([
    [1, new Date(2026, 9, 15, 12, 0, 0), "CAFE SPOT", "P1", 90]
  ]);

  const instSheet = ss.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
  instSheet.getRange(1, 1, 1, 11).setValues([
    ["#", "Start Date", "Description", "Total Price", "Months", "C6", "C7", "EMI", "Payer", "Paid", "Status"]
  ]);

  // Phase 2a: All three warning categories active
  const res1 = sandbox.updateLiveDashboard({ ss });
  assert.strictEqual(res1.health.overallStatus, "WARNING");
  assert.strictEqual(res1.health.allHealthy, false);
  assert.strictEqual(res1.health.parsePassed, true);
  assert.strictEqual(res1.health.unresolvedCount, 2);
  assert.strictEqual(res1.health.unresolvedAmount, 450);
  assert.strictEqual(res1.health.unlinkedInstCount, 1);
  assert.strictEqual(res1.health.unlinkedInstAmount, 300);
  assert.strictEqual(res1.health.orphanCount, 1);
  assert.strictEqual(res1.health.orphanAmount, 90);

  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  assert.strictEqual(ar.backgrounds["9_1"], "#fff2cc");
  assert.strictEqual(ar.backgrounds["10_3"], "#fff8e1");
  assert.strictEqual(ar.backgrounds["10_5"], "#fff8e1");
  assert.strictEqual(ar.backgrounds["10_7"], "#fff8e1");

  // Phase 2b: Assign I-1 directly to payer P1 (NOT linked to plan), assign M-1 to P1, dismiss ORPH-1
  let rowM1 = -1;
  let rowI1 = -1;
  let rowOrph1 = -1;
  for (let r = 15; r <= ar.getLastRow(); r++) {
    const id = String(ar._getCell(r, 1) || "");
    if (id === "M-1") rowM1 = r;
    if (id === "I-1") rowI1 = r;
    if (id === "ORPH-1") rowOrph1 = r;
  }
  assert.ok(rowM1 > 0 && rowI1 > 0 && rowOrph1 > 0, "Expected M-1, I-1, ORPH-1 rows on sheet");

  ar._setCell(rowM1, 6, "P1");
  sandbox.handleActiveReconEdit({ range: ar.getRange(rowM1, 6) });
  ar._setCell(rowI1, 6, "P1");
  sandbox.handleActiveReconEdit({ range: ar.getRange(rowI1, 6) });
  ar._setCell(rowOrph1, 6, "🚫 Dismiss / Duplicate Log");
  sandbox.handleActiveReconEdit({ range: ar.getRange(rowOrph1, 6) });

  const res2 = sandbox.updateLiveDashboard({ ss });
  assert.strictEqual(res2.health.unresolvedCount, 0, "Assigning I-1 & M-1 makes unresolved bank lines = 0");
  assert.strictEqual(res2.health.orphanCount, 0, "Dismissing ORPH-1 makes active orphan logs = 0");
  assert.strictEqual(res2.health.unlinkedInstCount, 1, "Unlinked installment metric must stay 1 until linked to a plan");
  assert.strictEqual(res2.health.overallStatus, "WARNING", "Panel must still warn while unlinkedInstCount > 0");

  // Phase 2c: Add Installment plan & link S-00002 -> I-00001
  instSheet.getRange(2, 1, 1, 11).setValues([
    [1, new Date(2026, 7, 15, 12, 0, 0), "Phone Plan", 1800, 6, 0, 0, 300, "P1", 1, "Ongoing"]
  ]);
  const decisions = sandbox.loadDecisions(ss);
  const idInfo = sandbox.ensureAllSheetRowIds(ss, decisions);
  const phoneStmtId = idInfo.stmt.ids[1];
  const phoneInstId = idInfo.inst.ids[0];
  sandbox.setStmtInstLink(decisions, phoneStmtId, [phoneInstId], "Phone Plan", "user");
  sandbox.saveDecisions(ss, decisions);

  const res3 = sandbox.updateLiveDashboard({ ss });
  assert.strictEqual(res3.health.unlinkedInstCount, 0, "Linked installment clears unlinked count");
  assert.strictEqual(res3.health.unresolvedCount, 0);
  assert.strictEqual(res3.health.orphanCount, 0);
  assert.strictEqual(res3.health.statementCheckOk, true);
  assert.strictEqual(res3.health.allHealthy, true);
  assert.strictEqual(res3.health.overallStatus, "HEALTHY");
  assert.strictEqual(ar.backgrounds["9_1"], "#e8f5e9");
}

// ==========================================
// Test 3: Parse Tie-Out Failure (Inline Error State + Blocked Parse Gate)
// ==========================================
{
  // 3a: Manual Bank Statement without META_IMPORT where math fails tie-out
  const envInline = createGasSandbox();
  const ssInline = new MockSpreadsheet();
  envInline.setActiveSpreadsheet(ssInline);

  const stmtInline = ssInline.insertSheet(envInline.CONFIG.SHEETS.BANK_STATEMENT);
  stmtInline.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 100, "Closing Balance:", 500, "Total Debit:", 100],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 500]
  ]);
  stmtInline.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Post Date", "Description", "Type", "Amount", "Status"]]);
  stmtInline.getRange(7, 1, 1, 7).setValues([
    [1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "TEST SHOP", "DEBIT", 100, ""]
  ]);
  const txInline = ssInline.insertSheet(envInline.CONFIG.SHEETS.TRANSACTIONS);
  txInline.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  txInline.getRange(2, 1, 1, 5).setValues([[1, new Date(2026, 9, 10, 12, 0, 0), "Test Shop", "P1", 100]]);

  const resErr = envInline.sandbox.updateLiveDashboard({ ss: ssInline });
  assert.strictEqual(resErr.blocked, false, "Without META_IMPORT inline dashboard renders with parse error banner");
  assert.strictEqual(resErr.health.parsePassed, false);
  assert.strictEqual(resErr.health.overallStatus, "ERROR");
  const arInline = ssInline.getSheetByName(envInline.CONFIG.SHEETS.ACTIVE_RECON);
  assert.ok(String(arInline._getCell(10, 1)).startsWith("Parse Tie-Out: ❌ Failed"), `Expected A10 Failed, got: ${arInline._getCell(10, 1)}`);
  assert.strictEqual(arInline.backgrounds["10_1"], "#fce8e6");
  assert.strictEqual(arInline.backgrounds["9_1"], "#fce8e6");

  // 3b: Statement import WITH META_IMPORT & failed tie-out triggers blocked error screen with Row 3 diagnostic strip
  const envBlocked = createGasSandbox();
  const ssBlocked = new MockSpreadsheet();
  envBlocked.setActiveSpreadsheet(ssBlocked);

  const badText = [
    "National Bank of Egypt",
    "Card Number: 5264****1234",
    "Statement Date: 30-09-2026",
    "Due Date: 25-10-2026\tCredit Limit: 100,000.00",
    "Opening Balance: 5,000.00\tClosing Balance: 2,264.17",
    "Total of Credit: 2,835.83\tTotal of Debit: 100.00",
    "10/09/2026 11/09/2026 COFFEE SHOP ALEX 100.00",
    "15/09/2026 16/09/2026 SAMPLE STORE EGY 2,835.83"
  ].join("\n");
  const badParsed = envBlocked.sandbox.parseNBEStatementText(badText);
  envBlocked.sandbox.writeStatementToSheet(ssBlocked, badParsed.meta, badParsed.transactions);
  const blockedRes = envBlocked.sandbox.updateLiveDashboard({ ss: ssBlocked });
  assert.strictEqual(blockedRes.blocked, true);

  const arBlocked = ssBlocked.getSheetByName(envBlocked.CONFIG.SHEETS.ACTIVE_RECON);
  const row3 = String(arBlocked._getCell(3, 1) || "");
  assert.ok(row3.includes("Parse Tie-Out: ❌ FAILED"), `Blocked error screen Row 3 must show Parse Tie-Out ❌ FAILED: ${row3}`);
  assert.ok(row3.includes("Last Refresh:"), `Blocked error screen Row 3 must include Last Refresh: ${row3}`);
  assert.ok(row3.includes("Code Version:"), `Blocked error screen Row 3 must include Code Version: ${row3}`);
}

// ==========================================
// Test 4: Version Stamping Script & getCodeVersion()
// ==========================================
{
  const stampedClean = formatVersionFile({
    gitSha: "abc1234",
    dirty: false,
    stampedAt: "2026-11-20T00:00:00.000Z"
  });
  assert.ok(stampedClean.includes('gitSha: "abc1234"') && stampedClean.includes("dirty: false"));

  const stampedDirty = formatVersionFile({
    gitSha: "abc1234",
    dirty: true,
    stampedAt: "2026-11-20T00:00:00.000Z"
  });
  assert.ok(stampedDirty.includes("dirty: true"));

  const liveInfo = computeVersionInfo(path.resolve(__dirname, ".."));
  assert.ok(/^[0-9a-f]{7,}$/i.test(liveInfo.gitSha) || liveInfo.gitSha === "unknown", `Invalid live gitSha: ${liveInfo.gitSha}`);

  const { sandbox } = createGasSandbox();
  const ver = sandbox.getCodeVersion();
  assert.ok(/^(?:[0-9a-f]{7,}|unknown)\*?$/i.test(ver), `getCodeVersion returned unexpected format: ${ver}`);
}

// ==========================================
// Test 5: Editing Inside Health Panel Rows (Rows 9-11) Is Safe
// ==========================================
{
  const realFixture = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "2026-11-real.json"), "utf8"));
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);
  populateSpreadsheetFromFixture(ss, realFixture, CONFIG);
  sandbox.updateLiveDashboard({ ss });

  const decSheet = ss.getSheetByName(CONFIG.SHEETS.DECISIONS);
  const decBefore = JSON.stringify(decSheet.getRange(1, 1, decSheet.getLastRow(), 7).getValues());

  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  assert.doesNotThrow(() => {
    sandbox.handleActiveReconEdit({ range: ar.getRange(10, colLetterToIdx("C")) });
    sandbox.handleActiveReconEdit({ range: ar.getRange(11, colLetterToIdx("A")) });
  }, "Editing cells inside Health Check panel (Rows 9-11) must not throw");

  const decAfter = JSON.stringify(decSheet.getRange(1, 1, decSheet.getLastRow(), 7).getValues());
  assert.strictEqual(decAfter, decBefore, "Editing Health Check panel rows must not mutate _Decisions");
}

console.log("health.test.js: all Health Check panel, live formula, tie-out, version stamp, and edit safety assertions passed.");
