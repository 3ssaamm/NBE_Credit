const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { createGasSandbox, populateSpreadsheetFromFixture, MockSpreadsheet } = require("./helpers/gas_sandbox");

const tz = "Africa/Cairo";

// ==========================================
// Test 1: Confidence Tiers (Auto vs Suggested on Lag & Unknown Merchant)
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 850.00, "Total Debit:", 850.00],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 850.00]
  ]);
  stmt.getRange(7, 1, 3, 7).setValues([
    [1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "CARIBOU COFFEE", "DEBIT", 150.00, ""],
    [2, new Date(2026, 9, 14, 12, 0, 0), new Date(2026, 9, 14, 12, 0, 0), "UBER TRIP", "DEBIT", 200.00, ""],
    [3, new Date(2026, 9, 20, 12, 0, 0), new Date(2026, 9, 20, 12, 0, 0), "POS TERMINAL 9921", "DEBIT", 500.00, ""]
  ]);

  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, 3, 5).setValues([
    [1, new Date(2026, 9, 8, 0, 0, 0), "Caribou", "Mido", 150.00],          // 2d + merchant -> Auto
    [2, new Date(2026, 9, 10, 0, 0, 0), "Uber", "Mai", 200.00],            // 4d lag -> Suggested only
    [3, new Date(2026, 9, 20, 0, 0, 0), "Corner Pharmacy", "Dad", 500.00]  // unknown merchant -> Suggested only
  ]);

  const recon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(recon.stmtDebits.filter(s => s.matched).length, 1, "Only Caribou (<=3d + merchant) should auto-match");
  assert.strictEqual(recon.stmtDebits[0].matchTier, "auto");
  assert.strictEqual(recon.missingInSheet.length, 2, "4-day lag Uber and unknown-merchant POS must stay Suggested");

  const uberRow = recon.missingInSheet.find(s => s.stmtId === "S-00002");
  const posRow = recon.missingInSheet.find(s => s.stmtId === "S-00003");
  assert.ok(uberRow && uberRow.suggestedPairs.length > 0 && uberRow.suggestedPairs[0].txIds[0] === "T-00002");
  assert.ok(posRow && posRow.suggestedPairs.length > 0 && posRow.suggestedPairs[0].txIds[0] === "T-00003");
}

// ==========================================
// Test 2: Auto Same-Merchant Split (0.06 diff) vs Suggested Cross-Merchant Split
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 389.94, "Total Debit:", 389.94],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 389.94]
  ]);
  stmt.getRange(7, 1, 2, 7).setValues([
    [1, new Date(2026, 9, 15, 12, 0, 0), new Date(2026, 9, 15, 12, 0, 0), "UBER EGYPT", "DEBIT", 89.94, ""],
    [2, new Date(2026, 9, 22, 12, 0, 0), new Date(2026, 9, 22, 12, 0, 0), "MYSTERY GATEWAY", "DEBIT", 300.00, ""]
  ]);

  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, 4, 5).setValues([
    [1, new Date(2026, 9, 15, 0, 0, 0), "Uber", "Mai", 65.00],
    [2, new Date(2026, 9, 15, 0, 0, 0), "Uber", "Muhanad", 25.00],
    [3, new Date(2026, 9, 22, 0, 0, 0), "Bakery", "Mido", 120.00],
    [4, new Date(2026, 9, 22, 0, 0, 0), "Bookshop", "Dad", 180.00]
  ]);

  const recon = sandbox.reconcileActiveStatement(ss, tz);
  const uberStmt = recon.stmtDebits.find(s => s.stmtId === "S-00001");
  const mysteryStmt = recon.stmtDebits.find(s => s.stmtId === "S-00002");

  assert.ok(uberStmt.matched, "Same-merchant split 65+25 vs 89.94 (diff 0.06) must auto-match");
  assert.strictEqual(uberStmt.matchTier, "auto-split");
  assert.strictEqual(mysteryStmt.matched, false, "Cross-merchant split (Bakery + Bookshop) must NEVER auto-match silently");
  assert.ok(mysteryStmt.suggestedPairs.length > 0, "Cross-merchant split must appear in suggestedPairs");
  assert.deepStrictEqual(
    Array.from(mysteryStmt.suggestedPairs[0].txIds).sort(),
    ["T-00003", "T-00004"]
  );
}

// ==========================================
// Test 3: 1-Click Confirmation + Learned Merchant Aliases + Generic Gateway Guard
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 340.00, "Total Debit:", 340.00],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 340.00]
  ]);
  stmt.getRange(7, 1, 2, 7).setValues([
    [1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "PAYMOB*COFFEENOOK", "DEBIT", 140.00, ""],
    [2, new Date(2026, 9, 12, 12, 0, 0), new Date(2026, 9, 12, 12, 0, 0), "FAWRY", "DEBIT", 200.00, ""]
  ]);

  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, 2, 5).setValues([
    [1, new Date(2026, 9, 10, 0, 0, 0), "BrewPlace", "Mido", 140.00],
    [2, new Date(2026, 9, 12, 0, 0, 0), "CornerGrocer", "Mai", 200.00]
  ]);

  let recon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(recon.missingInSheet.length, 2, "Neither should auto-match before learning");
  let ledger = sandbox.buildDebtLedger(ss, tz, recon);
  sandbox.renderActiveReconciliationSheet(ss, tz, ledger, recon);

  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  let rM1 = -1, rM2 = -1;
  for (let r = 12; r <= ar.getLastRow(); r++) {
    if (ar._getCell(r, 1) === "M-1") rM1 = r;
    if (ar._getCell(r, 1) === "M-2") rM2 = r;
  }

  // Confirm PAYMOB*COFFEENOOK <-> BrewPlace -> learns alias coffeenook <-> brewplace
  ar._setCell(rM1, 6, recon.missingInSheet[0].suggestedPairs[0].dropdownLabel);
  sandbox.handleActiveReconEdit({ range: ar.getRange(rM1, 6) });

  // Locate remaining M-1 (was FAWRY) and confirm FAWRY <-> CornerGrocer -> must NOT learn generic "fawry" alias!
  recon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(recon.missingInSheet.length, 1);
  let rRemaining = -1;
  for (let r = 12; r <= ar.getLastRow(); r++) {
    if (ar._getCell(r, 1) === "M-1") { rRemaining = r; break; }
  }
  ar._setCell(rRemaining, 6, recon.missingInSheet[0].suggestedPairs[0].dropdownLabel);
  sandbox.handleActiveReconEdit({ range: ar.getRange(rRemaining, 6) });

  const dec = sandbox.loadDecisions(ss);
  assert.ok(dec.merchantAliases["coffeenook::brewplace"], "Should learn specific alias coffeenook <-> brewplace");
  assert.strictEqual(Object.keys(dec.merchantAliases).length, 1, "Generic FAWRY must NOT be learned as a merchant alias!");

  // Import next month's statement (November 2026) with a new PAYMOB*COFFEENOOK charge -> should auto-match now!
  sandbox.writeStatementToSheet(ss, {
    statementDate: "30-11-2026",
    dueDate: "27-12-2026",
    creditLimit: 100000,
    openingBalance: 0,
    closingBalance: 350.00,
    totalDebit: 350.00,
    totalCredit: 0,
    card: "1234"
  }, [
    { txDate: "10-11-2026", postDate: "10-11-2026", desc: "PAYMOB*COFFEENOOK", type: "DEBIT", amount: 150.00 },
    { txDate: "12-11-2026", postDate: "12-11-2026", desc: "FAWRY", type: "DEBIT", amount: 200.00 }
  ]);
  tx.getRange(4, 1, 2, 5).setValues([
    [3, new Date(2026, 10, 10, 0, 0, 0), "BrewPlace", "Mido", 150.00],
    [4, new Date(2026, 10, 12, 0, 0, 0), "UnrelatedShop", "Dad", 200.00]
  ]);

  const novRecon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(novRecon.stmtDebits[0].matched, true, "Learned alias must auto-match COFFEENOOK <-> BrewPlace in next month!");
  assert.strictEqual(novRecon.stmtDebits[0].matchTier, "auto");
  assert.strictEqual(novRecon.stmtDebits[1].matched, false, "Generic FAWRY must NOT auto-match UnrelatedShop without user confirmation!");
}

// ==========================================
// Test 4: Reverse Pairing & Dismissing from Orphan Logs (ORPH-*)
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 250.00, "Total Debit:", 250.00],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 250.00]
  ]);
  stmt.getRange(7, 1, 1, 7).setValues([
    [1, new Date(2026, 9, 18, 12, 0, 0), new Date(2026, 9, 18, 12, 0, 0), "PAYMOB*UNKNOWNPOS", "DEBIT", 250.00, ""]
  ]);

  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, 2, 5).setValues([
    [1, new Date(2026, 9, 18, 0, 0, 0), "Alpha Bistro", "Mido", 250.00],
    [2, new Date(2026, 9, 19, 0, 0, 0), "Duplicate Note", "Mai", 99.00]
  ]);

  let recon = sandbox.reconcileActiveStatement(ss, tz);
  let ledger = sandbox.buildDebtLedger(ss, tz, recon);
  sandbox.renderActiveReconciliationSheet(ss, tz, ledger, recon);

  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  let rOrph1 = -1, rOrph2 = -1;
  for (let r = 12; r <= ar.getLastRow(); r++) {
    if (ar._getCell(r, 1) === "ORPH-1") rOrph1 = r;
    if (ar._getCell(r, 1) === "ORPH-2") rOrph2 = r;
  }
  assert.ok(rOrph1 > 0 && rOrph2 > 0, "Both orphan logs should render");

  // Dismiss ORPH-2 cleanly
  ar._setCell(rOrph2, 6, "🚫 Dismiss / Duplicate Log");
  sandbox.handleActiveReconEdit({ range: ar.getRange(rOrph2, 6) });

  // Pair ORPH-1 from Orphan Logs panel -> triggers STMT_MATCH + refresh
  const orph1Hint = recon.orphanLogs[0].suggestedStmts[0].dropdownLabel;
  ar._setCell(rOrph1, 6, orph1Hint);
  sandbox.handleActiveReconEdit({ range: ar.getRange(rOrph1, 6) });

  recon = sandbox.reconcileActiveStatement(ss, tz);
  ledger = sandbox.buildDebtLedger(ss, tz, recon);
  assert.strictEqual(recon.missingInSheet.length, 0, "Statement line paired via Orphan Logs must resolve");
  assert.strictEqual(recon.orphanLogs.length, 1, "Only dismissed ORPH-2 remains in orphanLogs");
  assert.strictEqual(recon.orphanLogs[0].isNeglected, true, "Dismissed orphan persists across refresh");
  assert.strictEqual(Number(ledger.personShares.find(p => p.person === "Mido").purchases.toFixed(2)), 250.00);
  assert.strictEqual(Number(ledger.personShares.find(p => p.person === "Mai").purchases.toFixed(2)), 0.00);
}

// ==========================================
// Test 5: Purchase-Date Rule (Q5c / Q13) & Transactions A:E Immutability
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 300.00, "Total Debit:", 300.00],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 300.00]
  ]);
  stmt.getRange(7, 1, 1, 7).setValues([
    [1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "AMAZON", "DEBIT", 300.00, ""]
  ]);

  const txRowsOrig = [
    [1, new Date(2026, 9, 10, 0, 0, 0), "Amazon", "Mido", 300.00],       // Matched -> billed 300
    [2, new Date(2026, 9, 25, 0, 0, 0), "Missed Entry", "Mai", 450.00],  // In-period unmatched -> Orphan Log -> 0 billed
    [3, new Date(2026, 9, 31, 0, 0, 0), "Oct 31 Dinner", "Dad", 600.00], // Cutoff rollover -> Pending next statement -> 0 billed
    [4, new Date(2026, 10, 2, 0, 0, 0), "Nov 2 Groceries", "Mum", 250.00] // Post-closing -> Pending next statement -> 0 billed
  ];
  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, txRowsOrig.length, 5).setValues(txRowsOrig.map(r => r.slice()));

  const recon = sandbox.reconcileActiveStatement(ss, tz);
  const ledger = sandbox.buildDebtLedger(ss, tz, recon);
  sandbox.renderActiveReconciliationSheet(ss, tz, ledger, recon);

  assert.strictEqual(recon.orphanLogs.length, 1, "Only in-period unmatched log is in orphanLogs");
  assert.strictEqual(recon.orphanLogs[0].txId, "T-00002");
  assert.strictEqual(recon.pendingNextItems.length, 2, "Oct 31 rollover and Nov 2 purchase belong in pendingNextItems");
  assert.strictEqual(Number(ledger.pendingTotal.toFixed(2)), 850.00);
  assert.strictEqual(Number(ledger.familySharesTotal.toFixed(2)), 300.00, "Neither orphans nor pending items may be billed on active bill");
  assert.strictEqual(Number(ledger.auditDifference.toFixed(2)), 0.00);

  // Verify Transactions Cols A:E were completely untouched
  for (let i = 0; i < txRowsOrig.length; i++) {
    assert.strictEqual(tx._getCell(i + 2, 1), txRowsOrig[i][0]);
    assert.strictEqual(tx._getCell(i + 2, 3), txRowsOrig[i][2]);
    assert.strictEqual(tx._getCell(i + 2, 4), txRowsOrig[i][3]);
    assert.strictEqual(tx._getCell(i + 2, 5), txRowsOrig[i][4]);
  }
}

// ==========================================
// Test 6: Real October 2026 Fixture Unlock-and-Confirm Replay (P1..P11 to 0.00 EGP)
// ==========================================
{
  const realFixturePath = path.join(__dirname, "fixtures", "2026-11-real.json");
  assert.ok(fs.existsSync(realFixturePath), "Expected tests/fixtures/2026-11-real.json to exist");
  const realFixture = JSON.parse(fs.readFileSync(realFixturePath, "utf8"));

  // Strip S-00039 and S-00045 locked matches to simulate fresh Phase 2 encounter on real October data
  const unlockedFixture = Object.assign({}, realFixture, {
    decisions: (realFixture.decisions || []).filter(d => !(d[0] === "STMT_MATCH" && (d[1] === "S-00039" || d[1] === "S-00045")))
  });

  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);
  populateSpreadsheetFromFixture(ss, unlockedFixture, CONFIG);

  let recon = sandbox.reconcileActiveStatement(ss, tz);
  const s39 = recon.missingInSheet.find(s => s.stmtId === "S-00039");
  const s45 = recon.missingInSheet.find(s => s.stmtId === "S-00045");
  assert.ok(s39, "S-00039 (unknown merchant 195 split) must NOT auto-match silently in Phase 2");
  assert.ok(s45, "S-00045 (unknown merchant 125) must NOT auto-match silently in Phase 2");
  assert.deepStrictEqual(
    Array.from(s39.suggestedPairs[0].txIds).sort(),
    ["T-00332", "T-00333"],
    "S-00039 must surface T-00332+T-00333 as top Suggested split pair"
  );
  assert.deepStrictEqual(
    Array.from(s45.suggestedPairs[0].txIds),
    ["T-00325"],
    "S-00045 must surface T-00325 as top Suggested pair"
  );

  // Confirm S-00039 and S-00045 via UI Dropdown Edits
  let ledger = sandbox.buildDebtLedger(ss, tz, recon);
  sandbox.renderActiveReconciliationSheet(ss, tz, ledger, recon);
  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);

  function findStmtRowOnAr(stmtId) {
    for (let r = 12; r <= ar.getLastRow(); r++) {
      if (ar._getCell(r, 9) === stmtId && /^[FCM]-/.test(String(ar._getCell(r, 1)))) {
        return r;
      }
    }
    return -1;
  }

  const r39 = findStmtRowOnAr("S-00039");
  assert.ok(r39 > 0, "S-00039 must render in Needs Pairing");
  ar._setCell(r39, 6, s39.suggestedPairs[0].dropdownLabel);
  sandbox.handleActiveReconEdit({ range: ar.getRange(r39, 6) });

  recon = sandbox.reconcileActiveStatement(ss, tz);
  const s45After = recon.missingInSheet.find(s => s.stmtId === "S-00045");
  const r45 = findStmtRowOnAr("S-00045");
  assert.ok(r45 > 0 && s45After, "S-00045 must still be in Needs Pairing until confirmed");
  ar._setCell(r45, 6, s45After.suggestedPairs[0].dropdownLabel);
  sandbox.handleActiveReconEdit({ range: ar.getRange(r45, 6) });

  recon = sandbox.reconcileActiveStatement(ss, tz);
  ledger = sandbox.buildDebtLedger(ss, tz, recon);

  assert.strictEqual(ledger.unresolvedCount, 0, "All bank lines must be resolved after confirming both suggestions");
  assert.strictEqual(ledger.orphanActiveCount, 0, "Zero active orphans after pairing October fixtures");
  assert.strictEqual(Number(ledger.auditDifference.toFixed(2)), 0.49, "Statement Check must equal +0.49 EGP");
  assert.strictEqual(ledger.isExactMatch, true, "isExactMatch must be true (+0.49 < 0.50)");

  Object.keys(realFixture.expectedShares).forEach(person => {
    const ps = ledger.personShares.find(item => item.person === person);
    assert.ok(ps, `Missing person ${person}`);
    const exp = realFixture.expectedShares[person];
    ["purchases", "installments", "fees", "unrecorded", "refunds", "netTotal"].forEach(k => {
      assert.strictEqual(
        Number(ps[k].toFixed(2)),
        exp[k],
        `Real October unlock+confirm mismatch for ${person}.${k}: got ${ps[k]}, expected ${exp[k]}`
      );
    });
  });
}

// ==========================================
// Test 7: 3-Way Split Pairing from Orphan Logs (Partial Match Protection, Add All, Sequential/Batch Append, Full Replace)
// ==========================================
{
  function makeSplitFixtureSpreadsheet(sandbox, CONFIG, setActiveSpreadsheet) {
    const ss = new MockSpreadsheet();
    setActiveSpreadsheet(ss);
    const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
    stmt.getRange(2, 1, 3, 6).setValues([
      ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
      ["Opening Balance:", 0, "Closing Balance:", 667.00, "Total Debit:", 667.00],
      ["Total Credit:", 0, "Card Number:", "1234", "Net:", 667.00]
    ]);
    stmt.getRange(7, 1, 1, 7).setValues([
      [1, new Date(2026, 9, 14, 12, 0, 0), new Date(2026, 9, 15, 12, 0, 0), "TE DATA", "DEBIT", 667.00, ""]
    ]);
    const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
    tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
    tx.getRange(2, 1, 3, 5).setValues([
      [1, new Date(2026, 9, 14, 0, 0, 0), "We Internet", "Mai", 100.00],
      [2, new Date(2026, 9, 14, 0, 0, 0), "We Internet", "Dad", 200.00],
      [3, new Date(2026, 9, 14, 0, 0, 0), "We Internet", "Mido", 367.00]
    ]);
    return ss;
  }

  // 7a. Partial match (Mido 367 vs Bank 667): Mido NOT overbilled, split remainder hints surfaced, 1-click Add all completes split
  {
    const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
    const ss = makeSplitFixtureSpreadsheet(sandbox, CONFIG, setActiveSpreadsheet);
    sandbox.updateLiveDashboard({ ss });

    const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
    const startRow = sandbox.getActiveReconStartShareRow();
    let rowOrphMido = -1;
    for (let r = startRow; r <= ar.getLastRow(); r++) {
      if (String(ar._getCell(r, 1)).startsWith("ORPH-") && ar._getCell(r, 5) === "Mido") {
        rowOrphMido = r;
        break;
      }
    }
    assert.ok(rowOrphMido > 0, "Should find Mido orphan row");

    ar._setCell(rowOrphMido, 6, "🔗 Pair with: S-00001 (667.00 · TE DATA)");
    sandbox.handleActiveReconEdit({ range: ar.getRange(rowOrphMido, 6) });

    let recon = sandbox.reconcileActiveStatement(ss, tz);
    let ledger = sandbox.buildDebtLedger(ss, tz, recon);

    const st1 = recon.stmtDebits.find(s => s.stmtId === "S-00001");
    assert.strictEqual(st1.isPartialMatch, true, "Single 367 tx on 667 stmt line must flag isPartialMatch");
    assert.strictEqual(st1.matchedSum, 367.00);
    assert.strictEqual(st1.remainingAmount, 300.00);
    assert.strictEqual(Number(ledger.personShares.find(p => p.person === "Mido").purchases.toFixed(2)), 367.00, "Partial match must bill Mido 367.00 (NOT 667.00!)");
    assert.strictEqual(Number(ledger.auditDifference.toFixed(2)), -300.00, "Unpaired split remainder shows -300.00 Statement Check");

    const midoShare = sandbox.shareMemberStatement(ss, tz, "Mido", { skipPdf: true, context: { recon, ledger } });
    assert.strictEqual(midoShare.memberData.totals.purchases, 367.00, "WhatsApp share must also bill Mido only 367.00 during partial match");

    assert.strictEqual(recon.orphanLogs.length, 2, "Mai (100) and Dad (200) must remain in Orphan Logs");
    const maiOrph = recon.orphanLogs.find(o => o.person === "Mai");
    const dadOrph = recon.orphanLogs.find(o => o.person === "Dad");
    const maiLabels = maiOrph.suggestedStmts.map(s => s.dropdownLabel);
    const dadLabels = dadOrph.suggestedStmts.map(s => s.dropdownLabel);
    assert.ok(maiLabels.some(l => l.startsWith("🔗 Add all to: S-00001")), `Mai should get Add all hint: ${JSON.stringify(maiLabels)}`);
    assert.ok(maiLabels.some(l => l.startsWith("🔗 Add to: S-00001 (rem 300.00")), `Mai should get Add to hint: ${JSON.stringify(maiLabels)}`);
    assert.ok(dadLabels.some(l => l.startsWith("🔗 Add all to: S-00001")), `Dad should get Add all hint: ${JSON.stringify(dadLabels)}`);

    let rowOrphMai = -1;
    for (let r = startRow; r <= ar.getLastRow(); r++) {
      if (String(ar._getCell(r, 1)).startsWith("ORPH-") && ar._getCell(r, 5) === "Mai") {
        rowOrphMai = r;
        break;
      }
    }
    const addAllLabel = maiLabels.find(l => l.startsWith("🔗 Add all to: S-00001"));
    const flushBefore = sandbox.__flushCount || 0;
    ss.toasts.length = 0;
    // Call full onEdit entry point with e.value overlay WITHOUT manually calling updateLiveDashboard afterward
    sandbox.onEdit({ range: ar.getRange(rowOrphMai, 6), value: addAllLabel });

    assert.ok((sandbox.__flushCount || 0) > flushBefore, "onEdit auto-refresh must flush updated sheet");
    let remainingOrphRowsOnAr = 0;
    const arTable1Purchases = {};
    for (let r = startRow; r <= ar.getLastRow(); r++) {
      const colA = String(ar._getCell(r, 1) || "");
      if (/^ORPH-/i.test(colA)) remainingOrphRowsOnAr++;
      if (/^\d+$/.test(colA)) {
        const memberName = ar._getCell(r, 2);
        if (memberName) arTable1Purchases[memberName] = ar._getCell(r, 3);
      }
    }
    assert.strictEqual(remainingOrphRowsOnAr, 0, "onEdit must automatically re-render Active Bill & clear Orphan rows immediately");
    assert.strictEqual(arTable1Purchases["Mai"], "=100.00", "onEdit auto-refresh must immediately update Mai Table 1 purchases");
    assert.strictEqual(arTable1Purchases["Dad"], "=200.00", "onEdit auto-refresh must immediately update Dad Table 1 purchases");
    assert.strictEqual(arTable1Purchases["Mido"], "=367.00", "onEdit auto-refresh must immediately update Mido Table 1 purchases");
    assert.ok(ss.toasts.some(t => t.includes("⏳") && t.includes("S-00001")), "Should show immediate progress toast");
    assert.ok(ss.toasts.some(t => !t.startsWith("⏳") && t.includes("🔗 Paired S-00001")), "Should show completion toast after refresh");

    recon = sandbox.reconcileActiveStatement(ss, tz);
    ledger = sandbox.buildDebtLedger(ss, tz, recon);
    const dec = sandbox.loadDecisions(ss);
    assert.strictEqual(dec.stmtMatch["S-00001"].targetIds.length, 3, "All 3 txs must be stored on S-00001");
    assert.strictEqual(recon.orphanLogs.length, 0, "Zero orphans after Add all");
    assert.strictEqual(recon.stmtDebits[0].isPartialMatch, false);
    assert.strictEqual(Number(ledger.auditDifference.toFixed(2)), 0.00);
    assert.strictEqual(Number(ledger.personShares.find(p => p.person === "Mai").purchases.toFixed(2)), 100.00);
    assert.strictEqual(Number(ledger.personShares.find(p => p.person === "Dad").purchases.toFixed(2)), 200.00);
    assert.strictEqual(Number(ledger.personShares.find(p => p.person === "Mido").purchases.toFixed(2)), 367.00);
    assert.ok(Object.keys(dec.merchantAliases).length >= 1, "Merchant alias te data <-> we internet should be learned");
  }

  // 7b. Batch paste AND sequential "🔗 Pair with: S-00001" across all 3 rows appends without overwriting
  {
    const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
    const ss = makeSplitFixtureSpreadsheet(sandbox, CONFIG, setActiveSpreadsheet);
    sandbox.updateLiveDashboard({ ss });
    const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
    const startRow = sandbox.getActiveReconStartShareRow();

    const orphRows = [];
    for (let r = startRow; r <= ar.getLastRow(); r++) {
      if (String(ar._getCell(r, 1)).startsWith("ORPH-")) orphRows.push(r);
    }
    assert.strictEqual(orphRows.length, 3);
    const pairStr = "🔗 Pair with: S-00001 (667.00 · TE DATA)";
    orphRows.forEach(r => ar._setCell(r, 6, pairStr));
    sandbox.handleActiveReconEdit({ range: ar.getRange(orphRows[0], 6, 3, 1) });

    const decBatch = sandbox.loadDecisions(ss);
    assert.strictEqual(decBatch.stmtMatch["S-00001"].targetIds.length, 3, "Batch Pair with on 3 rows must append all 3 IDs");
    const reconBatch = sandbox.reconcileActiveStatement(ss, tz);
    const ledgerBatch = sandbox.buildDebtLedger(ss, tz, reconBatch);
    assert.strictEqual(Number(ledgerBatch.personShares.find(p => p.person === "Mai").purchases.toFixed(2)), 100.00);
    assert.strictEqual(Number(ledgerBatch.personShares.find(p => p.person === "Dad").purchases.toFixed(2)), 200.00);
    assert.strictEqual(Number(ledgerBatch.personShares.find(p => p.person === "Mido").purchases.toFixed(2)), 367.00);

    // Sequential one-by-one test on fresh sheet
    const envSeq = createGasSandbox();
    const ssSeq = makeSplitFixtureSpreadsheet(envSeq.sandbox, envSeq.CONFIG, envSeq.setActiveSpreadsheet);
    envSeq.sandbox.updateLiveDashboard({ ss: ssSeq });
    const arSeq = ssSeq.getSheetByName(envSeq.CONFIG.SHEETS.ACTIVE_RECON);
    const startSeq = envSeq.sandbox.getActiveReconStartShareRow();
    for (let step = 0; step < 3; step++) {
      let targetRow = -1;
      for (let r = startSeq; r <= arSeq.getLastRow(); r++) {
        if (String(arSeq._getCell(r, 1)).startsWith("ORPH-")) { targetRow = r; break; }
      }
      assert.ok(targetRow > 0, `Sequential step ${step} must locate remaining ORPH row`);
      arSeq._setCell(targetRow, 6, pairStr);
      envSeq.sandbox.handleActiveReconEdit({ range: arSeq.getRange(targetRow, 6) });
    }
    const decSeq = envSeq.sandbox.loadDecisions(ssSeq);
    assert.strictEqual(decSeq.stmtMatch["S-00001"].targetIds.length, 3, "Sequential Pair with across 3 turns must keep all 3 IDs");
  }

  // 7c. Full 1-to-1 match replacement still works when statement line is already full
  {
    const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
    const ss = new MockSpreadsheet();
    setActiveSpreadsheet(ss);
    const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
    stmt.getRange(2, 1, 3, 6).setValues([
      ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
      ["Opening Balance:", 0, "Closing Balance:", 100.00, "Total Debit:", 100.00],
      ["Total Credit:", 0, "Card Number:", "1234", "Net:", 100.00]
    ]);
    stmt.getRange(7, 1, 1, 7).setValues([
      [1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "ONLINE SHOP", "DEBIT", 100.00, ""]
    ]);
    const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
    tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
    tx.getRange(2, 1, 2, 5).setValues([
      [1, new Date(2026, 9, 10, 0, 0, 0), "Shop A", "Mai", 100.00],
      [2, new Date(2026, 9, 10, 0, 0, 0), "Shop B", "Mido", 100.00]
    ]);

    const decInit = sandbox.loadDecisions(ss);
    const idInfo = sandbox.ensureAllSheetRowIds(ss, decInit);
    sandbox.setStmtMatch(decInit, idInfo.stmt.ids[0], [idInfo.tx.ids[0]], "Initial full match", "user");
    sandbox.saveDecisions(ss, decInit);

    sandbox.updateLiveDashboard({ ss });
    const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
    const startRow = sandbox.getActiveReconStartShareRow();
    let rowOrphMido = -1;
    for (let r = startRow; r <= ar.getLastRow(); r++) {
      if (String(ar._getCell(r, 1)).startsWith("ORPH-")) { rowOrphMido = r; break; }
    }
    assert.ok(rowOrphMido > 0);
    ar._setCell(rowOrphMido, 6, "🔗 Pair with: S-00001 (100.00 · ONLINE SHOP)");
    sandbox.handleActiveReconEdit({ range: ar.getRange(rowOrphMido, 6) });

    const decAfter = sandbox.loadDecisions(ss);
    assert.deepStrictEqual(Array.from(decAfter.stmtMatch["S-00001"].targetIds), [idInfo.tx.ids[1]], "Full 100 match must replace T-00001 with T-00002");
  }
}

console.log("matcher.test.js: all 7 confidence-tiered matching, split orphan pairing, learning, Q5c, and October replay assertions passed.");
