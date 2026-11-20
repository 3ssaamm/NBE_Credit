const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { createGasSandbox, populateSpreadsheetFromFixture, MockSpreadsheet } = require("./helpers/gas_sandbox");

const tz = "Africa/Cairo";

// ==========================================
// Test 1: Close Gates (Unresolved, Statement Check Mismatch, Parse Tie-Out Failure)
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 300, "Total Debit:", 300],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 300]
  ]);
  stmt.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Post Date", "Description", "Type", "Amount", "Status"]]);
  stmt.getRange(7, 1, 1, 7).setValues([
    [1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "UNASSIGNED SHOP", "DEBIT", 300, ""]
  ]);
  ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS).getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);

  // 1a. Block when unresolved > 0
  let resUnres = sandbox.closeActiveStatement(ss, tz);
  assert.strictEqual(resUnres.closed, false, "Must block close when unresolvedCount > 0");
  assert.ok(resUnres.blockers.some(b => b.includes("Unresolved bank lines")));
  assert.strictEqual(ss.getSheetByName(CONFIG.SHEETS.HISTORY), null, "History must NOT be created on blocked close");

  // Resolve unassigned line -> check closes vs statement check mismatch (+100 diff in Closing)
  stmt.getRange(3, 4).setValue(200); // Closing 200 vs 300 DEBIT assigned -> +100 Statement Check diff
  const decisions = sandbox.loadDecisions(ss);
  const idInfo = sandbox.ensureAllSheetRowIds(ss, decisions);
  sandbox.setStmtAssign(decisions, idInfo.stmt.ids[0], "Mido", "test", "user");
  sandbox.saveDecisions(ss, decisions);

  let resDiff = sandbox.closeActiveStatement(ss, tz);
  assert.strictEqual(resDiff.closed, false, "Must block close when |auditDifference| >= 0.50");
  assert.ok(resDiff.blockers.some(b => b.includes("Statement check difference")));
  assert.strictEqual(ss.getSheetByName(CONFIG.SHEETS.HISTORY), null);

  // 1b. Block when Phase 4 parse tie-out failed
  const badParse = sandbox.parseNBEStatementText([
    "Statement Date: 30-09-2026",
    "Due Date: 25-10-2026\tCredit Limit: 100,000.00",
    "Opening Balance: 5,000.00\tClosing Balance: 2,264.17",
    "Total of Credit: 2,835.83\tTotal of Debit: 100.00",
    "10/09/2026 11/09/2026 COFFEE SHOP 100.00",
    "15/09/2026 16/09/2026 SHOP UNKNOWN 2,835.83"
  ].join("\n"));
  sandbox.writeStatementToSheet(ss, badParse.meta, badParse.transactions);
  let resTieOut = sandbox.closeActiveStatement(ss, tz);
  assert.strictEqual(resTieOut.closed, false, "Must block close when parse tie-out failed");
  assert.ok(resTieOut.blockers.some(b => b.includes("parse tie-out failed")));
  assert.strictEqual(ss.getSheetByName(CONFIG.SHEETS.HISTORY), null);
}

// ==========================================
// Test 2: Real October (2026-11-real.json) Close, Snapshot Fidelity & Immutability
// ==========================================
{
  const realFixture = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "2026-11-real.json"), "utf8"));
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);
  populateSpreadsheetFromFixture(ss, realFixture, CONFIG);

  sandbox.updateLiveDashboard();

  const closeRes = sandbox.closeActiveStatement(ss, tz);
  assert.strictEqual(closeRes.closed, true, `October statement (+0.49 check, 0 unresolved) must close cleanly: ${JSON.stringify(closeRes.blockers)}`);
  assert.strictEqual(closeRes.cycleKey, "2026-11");
  assert.strictEqual(closeRes.status, "Closed");
  assert.ok(closeRes.decisionCount > 0, "Close must snapshot decisions");

  const histSheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY);
  assert.ok(histSheet, "History sheet must be created");

  const archive = sandbox.readHistoryArchive(ss);
  assert.ok(archive["2026-11"], "History archive must contain 2026-11");
  assert.strictEqual(archive["2026-11"].label, "November 2026");
  assert.strictEqual(archive["2026-11"].status, "Closed");
  assert.strictEqual(Number(archive["2026-11"].summary.meta.closingBill.toFixed(2)), realFixture.expectedTotals.closingBill);
  assert.strictEqual(Number(archive["2026-11"].summary.meta.auditDifference.toFixed(2)), 0.49);

  // Verify every person share in History matches expectedShares to exact 0.00 EGP
  Object.keys(realFixture.expectedShares).forEach(p => {
    const expNet = realFixture.expectedShares[p].netTotal;
    if (expNet !== 0 || realFixture.expectedShares[p].purchases !== 0 || realFixture.expectedShares[p].installments !== 0) {
      assert.strictEqual(
        Number((archive["2026-11"].people[p] || 0).toFixed(2)),
        expNet,
        `History member netTotal mismatch for ${p}`
      );
      assert.strictEqual(
        Number((archive["2026-11"].memberBreakdown[p].purchases || 0).toFixed(2)),
        realFixture.expectedShares[p].purchases,
        `History purchases mismatch for ${p}`
      );
      assert.strictEqual(
        Number((archive["2026-11"].memberBreakdown[p].installments || 0).toFixed(2)),
        realFixture.expectedShares[p].installments,
        `History installments mismatch for ${p}`
      );
    }
  });

  // Decision snapshot contains both explicit user decisions and auto-matches
  const snapTypes = new Set(archive["2026-11"].decisions.map(d => d.type));
  assert.ok(snapTypes.has("STMT_ASSIGN"), "Snapshot must include STMT_ASSIGN decisions");
  assert.ok(snapTypes.has("STMT_MATCH") || snapTypes.has("AUTO_MATCH"), "Snapshot must include statement debit pairings");
  assert.ok(snapTypes.has("AUTO_INST_LINK") || snapTypes.has("STMT_INST_LINK"), "Snapshot must include installment links");

  // Active Bill banner marks closed cycle
  const arSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  assert.ok(String(arSheet._getCell(1, 1)).includes("Closed in History"), "Active Bill banner must display Closed in History badge");

  // Monthly History & Forecast shows November 2026 (Closed) with NO duplicate (Active) column
  const forecastSheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY_FORECAST);
  const fLastCol = forecastSheet.getLastColumn();
  const fHeaders = forecastSheet.getRange(4, 1, 1, fLastCol).getValues()[0].map(h => String(h));
  assert.ok(fHeaders.some(h => h.includes("November 2026") && h.includes("(Closed)")), `Forecast must render November 2026 (Closed): ${JSON.stringify(fHeaders)}`);
  assert.ok(!fHeaders.some(h => h.includes("November 2026") && h.includes("(Active)")), "Closed active month must NOT duplicate as (Active)");

  // Attempting to close a second time is rejected cleanly (immutable archive)
  const histRowsBefore = histSheet.getLastRow();
  const closeAgain = sandbox.closeActiveStatement(ss, tz);
  assert.strictEqual(closeAgain.closed, false, "Already-closed cycle must not re-close");
  assert.ok(closeAgain.blockers.some(b => b.includes("already closed")));
  assert.strictEqual(histSheet.getLastRow(), histRowsBefore, "History row count must remain unchanged");

  // Mutating Transactions afterward does NOT alter closed 2026-11 values in History or Monthly History & Forecast
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  txSheet._setCell(2, 5, 99999);
  sandbox.updateLiveDashboard();
  const archiveAfterEdit = sandbox.readHistoryArchive(ss);
  Object.keys(realFixture.expectedShares).forEach(p => {
    if (realFixture.expectedShares[p].netTotal !== 0) {
      assert.strictEqual(archiveAfterEdit["2026-11"].people[p], archive["2026-11"].people[p], `History must stay frozen for ${p}`);
    }
  });
}

// ==========================================
// Test 3: One-Time Legacy Migration From Monthly History & Forecast (Mar-Aug)
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  // Seed legacy Monthly History & Forecast sheet with closed Paid/Settled columns + Active + Forecast
  const legacyHist = ss.insertSheet(CONFIG.SHEETS.HISTORY_FORECAST);
  legacyHist.getRange(4, 1, 1, 6).setValues([[
    "Family Member",
    "March 2026\n(Paid)",
    "April 2026\n(Paid)",
    "August 2026\n(Settled)",
    "November 2026\n(Active)",
    "December 2026\n(Forecast)"
  ]]);
  legacyHist.getRange(5, 1, 4, 6).setValues([
    ["Alice", 1200.50, 0, 500.00, 999, 100],
    ["Bob", 300.00, 450.25, 0, 888, 200],
    ["Carol", 0, 210.00, 750.75, 777, 300],
    ["TOTAL BILL (EGP)", 1500.50, 660.25, 1250.75, 2664, 600]
  ]);

  // Seed a minimal clean statement & transactions for active cycle (November 2026)
  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 100, "Total Debit:", 100],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 100]
  ]);
  stmt.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Post Date", "Description", "Type", "Amount", "Status"]]);
  stmt.getRange(7, 1, 1, 7).setValues([[1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "CAFE", "DEBIT", 100, ""]]);
  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, 1, 5).setValues([[1, new Date(2026, 9, 10, 0, 0, 0), "Cafe", "Alice", 100]]);

  // Two refreshes -> migrates Mar/Apr/Aug once, ignores (Active)/(Forecast), never duplicates
  sandbox.updateLiveDashboard();
  const histAfter1 = ss.getSheetByName(CONFIG.SHEETS.HISTORY).getLastRow();
  sandbox.updateLiveDashboard();
  const histAfter2 = ss.getSheetByName(CONFIG.SHEETS.HISTORY).getLastRow();
  assert.strictEqual(histAfter2, histAfter1, "Second refresh must NEVER duplicate migrated History rows");

  const dec = sandbox.loadDecisions(ss);
  assert.strictEqual(dec.isHistoryMigrated, true, "V1_HISTORY_ARCHIVE migration flag must persist");

  const archive = sandbox.readHistoryArchive(ss);
  assert.deepStrictEqual(Object.keys(archive).sort(), ["2026-03", "2026-04", "2026-08"]);
  assert.strictEqual(archive["2026-03"].people["Alice"], 1200.50);
  assert.strictEqual(archive["2026-03"].people["Bob"], 300.00);
  assert.strictEqual(archive["2026-03"].people["Carol"], undefined, "0.00 entries should not clutter History");
  assert.strictEqual(archive["2026-04"].people["Bob"], 450.25);
  assert.strictEqual(archive["2026-08"].status, "Settled");
  assert.strictEqual(archive["2026-08"].people["Carol"], 750.75);

  const forecastOut = ss.getSheetByName(CONFIG.SHEETS.HISTORY_FORECAST);
  const outHeaders = forecastOut.getRange(4, 1, 1, forecastOut.getLastColumn()).getValues()[0];
  assert.ok(outHeaders.some(h => String(h).includes("March 2026\n(Paid)")));
  assert.ok(outHeaders.some(h => String(h).includes("August 2026\n(Settled)")));
  assert.ok(outHeaders.some(h => String(h).includes("November 2026\n(Active)")));
}

// ==========================================
// Test 4: Fresh Template Has 0 Hardcoded History & Static Guard
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 50, "Total Debit:", 50],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 50]
  ]);
  stmt.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Post Date", "Description", "Type", "Amount", "Status"]]);
  stmt.getRange(7, 1, 1, 7).setValues([[1, new Date(2026, 9, 5, 12, 0, 0), new Date(2026, 9, 5, 12, 0, 0), "SHOP", "DEBIT", 50, ""]]);
  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, 1, 5).setValues([[1, new Date(2026, 9, 5, 0, 0, 0), "Shop", "UserOne", 50]]);

  sandbox.updateLiveDashboard();
  const archive = sandbox.readHistoryArchive(ss);
  assert.deepStrictEqual(Object.keys(archive), [], "Fresh template must have 0 closed history months");

  const forecastSheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY_FORECAST);
  const headers = forecastSheet.getRange(4, 1, 1, forecastSheet.getLastColumn()).getValues()[0].map(String);
  assert.ok(!headers.some(h => h.includes("March 2026")), "Fresh template must NOT inject hardcoded March 2026 history");

  const rootDir = path.resolve(__dirname, "..");
  const codeSrc = fs.readFileSync(path.join(rootDir, "Code.js"), "utf8");
  const histSrc = fs.readFileSync(path.join(rootDir, "src", "History.js"), "utf8");
  [codeSrc, histSrc].forEach(src => {
    assert.ok(!src.includes("getHistoricalMonthsArchive"), "Hardcoded getHistoricalMonthsArchive must be deleted");
    assert.ok(!src.includes("defaultHistory"), "Hardcoded defaultHistory must be deleted");
    assert.ok(!src.includes("17143"), "Hardcoded May 2026 historical figure must be deleted");
    assert.ok(!src.includes("21696.49"), "Hardcoded July 2026 historical figure must be deleted");
    assert.ok(!src.includes("6233.32"), "Hardcoded March 2026 historical figure must be deleted");
  });
}

console.log("history.test.js: all Close Statement gating, immutable History archive, legacy migration, and zero-hardcoded-history assertions passed.");
