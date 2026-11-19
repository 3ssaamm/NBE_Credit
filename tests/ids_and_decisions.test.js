const assert = require("assert");
const { createGasSandbox, MockSpreadsheet } = require("./helpers/gas_sandbox");

const tz = "Africa/Cairo";

// ==========================================
// 1. Hidden _Row_ID Creation & User-Column Protection
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 500, "Total Debit:", 500],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 500]
  ]);
  stmt.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Posting Date", "Description", "Type", "Amount (EGP)", "Status"]]);
  stmt.getRange(7, 1, 2, 7).setValues([
    [1, new Date(2026, 9, 5, 12, 0, 0), new Date(2026, 9, 5, 12, 0, 0), "COFFEE SHOP", "DEBIT", 100, ""],
    [2, new Date(2026, 9, 6, 12, 0, 0), new Date(2026, 9, 6, 12, 0, 0), "GADGET 1 OF 3", "INSTALLMENT", 400, ""]
  ]);

  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 6).setValues([["#", "Date", "Description", "Person", "Amount", "My Custom Notes"]]);
  tx.getRange(2, 1, 2, 6).setValues([
    [1, new Date(2026, 9, 5, 0, 0, 0), "Coffee", "Mido", 100, "User note row 2"],
    [2, new Date(2026, 9, 6, 0, 0, 0), "Lunch", "Mai", 200, ""]
  ]);

  const inst = ss.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
  inst.getRange(1, 1, 1, 11).setValues([["#", "Start", "Desc", "Total", "Months", "C6", "C7", "EMI", "Payer", "Paid", "Status"]]);
  inst.getRange(2, 1, 1, 11).setValues([
    [1, new Date(2026, 9, 1, 12, 0, 0), "Gadget", 1200, 3, 0, 0, 400, "Mido", 0, "Ongoing"]
  ]);

  const res = sandbox.ensureAllSheetRowIds(ss);

  assert.strictEqual(res.stmt.idCol, 8, "Bank Statement _Row_ID should use Col H (8)");
  assert.ok(stmt.hiddenCols.has(8), "Bank Statement Col H must be hidden");
  assert.deepStrictEqual(Array.from(res.stmt.ids), ["S-00001", "S-00002"]);
  assert.strictEqual(stmt._getCell(6, 8), "_Row_ID");

  // Non-interference check: Transactions Col F had user notes -> must use Col G (7)
  assert.strictEqual(res.tx.idCol, 7, "Transactions with Col F note must place _Row_ID in Col G (7)");
  assert.ok(tx.hiddenCols.has(7) && !tx.hiddenCols.has(6), "Transactions Col G must be hidden and Col F left visible");
  assert.strictEqual(tx._getCell(2, 6), "User note row 2", "User note in Col F must be untouched");
  assert.strictEqual(tx._getCell(1, 6), "My Custom Notes", "User header in Col F must be untouched");
  assert.deepStrictEqual(Array.from(res.tx.ids), ["T-00001", "T-00002"]);

  assert.strictEqual(res.inst.idCol, 12, "Installments _Row_ID should use Col L (12)");
  assert.ok(inst.hiddenCols.has(12), "Installments Col L must be hidden");
  assert.deepStrictEqual(Array.from(res.inst.ids), ["I-00001"]);
}

// ==========================================
// 2. Monotonic Non-Reuse & Single-Row onEdit Allocation
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, 3, 5).setValues([
    [1, new Date(2026, 9, 1, 0, 0, 0), "A", "Mido", 10],
    [2, new Date(2026, 9, 2, 0, 0, 0), "B", "Mai", 20],
    [3, new Date(2026, 9, 3, 0, 0, 0), "C", "Dad", 30]
  ]);

  sandbox.ensureAllSheetRowIds(ss);
  assert.strictEqual(tx._getCell(4, 6), "T-00003");

  // Delete row 4 completely (simulating user deleting last transaction)
  tx.getRange(4, 1, 1, 6).clearContent();

  // Add a brand new row at row 4 and trigger onEdit
  tx.getRange(4, 1, 1, 5).setValues([[3, new Date(2026, 9, 4, 0, 0, 0), "New Item", "mido", 40]]);
  sandbox.onEdit({ range: tx.getRange(4, 4, 1, 1) });

  assert.strictEqual(tx._getCell(4, 4), "Mido", "onEdit should still normalize payer casing");
  assert.strictEqual(tx._getCell(4, 6), "T-00004", "New row must NEVER reuse deleted T-00003 ID");

  // Copy-pasting an existing ID onto row 5 must auto-assign a fresh ID on ensureAllSheetRowIds
  tx.getRange(5, 1, 1, 6).setValues([[4, new Date(2026, 9, 5, 0, 0, 0), "Copied", "Mai", 50, "T-00001"]]);
  const resAfterDup = sandbox.ensureAllSheetRowIds(ss);
  assert.strictEqual(resAfterDup.tx.ids[0], "T-00001");
  assert.strictEqual(resAfterDup.tx.ids[3], "T-00005", "Duplicate copied ID must receive fresh T-00005");
}

// ==========================================
// 3. Same-Day Duplicate NUTOPIA Stability When Rows Above Disappear + Hidden Col I
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 580, "Total Debit:", 580],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 580]
  ]);
  stmt.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Posting Date", "Description", "Type", "Amount (EGP)", "Status"]]);
  stmt.getRange(7, 1, 3, 7).setValues([
    [1, new Date(2026, 9, 20, 12, 0, 0), new Date(2026, 9, 20, 12, 0, 0), "FAWRY STORE", "DEBIT", 500, ""],
    [2, new Date(2026, 9, 23, 12, 0, 0), new Date(2026, 9, 23, 12, 0, 0), "NUTOPIA-SUTHERLAND ALEX. -02G", "DEBIT", 40, ""],
    [3, new Date(2026, 9, 23, 12, 0, 0), new Date(2026, 9, 23, 12, 0, 0), "NUTOPIA-SUTHERLAND ALEX. -02G", "DEBIT", 40, ""]
  ]);

  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);

  sandbox.updateLiveDashboard();
  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  assert.ok(ar.hiddenCols.has(9), "Active Bill & Reconciliation Col I (_Subject_ID) must be hidden");

  // Locate M-2 and M-3 (the two identical same-day NUTOPIA rows)
  let rM2 = -1, rM3 = -1;
  for (let r = 12; r <= ar.getLastRow(); r++) {
    const id = ar._getCell(r, 1);
    if (id === "M-2") rM2 = r;
    if (id === "M-3") rM3 = r;
  }
  assert.strictEqual(ar._getCell(rM2, 9), "S-00002");
  assert.strictEqual(ar._getCell(rM3, 9), "S-00003");

  ar._setCell(rM2, 6, "Mido");
  sandbox.handleActiveReconEdit({ range: ar.getRange(rM2, 6) });
  ar._setCell(rM3, 6, "Mai");
  sandbox.handleActiveReconEdit({ range: ar.getRange(rM3, 6) });

  // Now log FAWRY STORE (S-00001) so it matches and disappears from Table 2 (renumbering Table 2!)
  tx.getRange(2, 1, 1, 5).setValues([[1, new Date(2026, 9, 20, 0, 0, 0), "Fawry", "Dad", 500]]);
  sandbox.updateLiveDashboard();

  let newM1 = -1, newM2 = -1;
  for (let r = 12; r <= ar.getLastRow(); r++) {
    if (ar._getCell(r, 1) === "M-1") newM1 = r;
    if (ar._getCell(r, 1) === "M-2") newM2 = r;
  }
  assert.strictEqual(ar._getCell(newM1, 9), "S-00002");
  assert.strictEqual(ar._getCell(newM1, 6), "Mido", "S-00002 must stay assigned to Mido after M-1 above matched");
  assert.strictEqual(ar._getCell(newM2, 9), "S-00003");
  assert.strictEqual(ar._getCell(newM2, 6), "Mai", "S-00003 must stay assigned to Mai after row renumbering");
}

// ==========================================
// 4. writeStatementToSheet Same-Date Preservation vs New-Statement Isolation
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const inst = ss.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
  inst.getRange(1, 1, 1, 11).setValues([["#", "Start", "Desc", "Total", "Months", "C6", "C7", "EMI", "Payer", "Paid", "Status"]]);
  inst.getRange(2, 1, 1, 11).setValues([[1, new Date(2026, 5, 1, 12, 0, 0), "Gym", 600, 6, 0, 0, 100, "Mido", 2, "Ongoing"]]);

  const octMeta = {
    statementDate: "31-10-2026",
    dueDate: "27-11-2026",
    creditLimit: 100000,
    openingBalance: 0,
    closingBalance: 100,
    totalDebit: 100,
    totalCredit: 0,
    card: "1234"
  };
  const octTx = [{ txDate: "31-10-2026", postDate: "31-10-2026", desc: "GYM 3 OF 6", type: "INSTALLMENT", amount: 100 }];

  sandbox.writeStatementToSheet(ss, octMeta, octTx);
  const stmt = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  assert.strictEqual(stmt._getCell(7, 8), "S-00001");

  // Reconcile and mark INST-1 skipped for October cycle
  let recon = sandbox.reconcileActiveStatement(ss, tz);
  let ledger = sandbox.buildDebtLedger(ss, tz, recon);
  sandbox.renderActiveReconciliationSheet(ss, tz, ledger, recon);
  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  let instRow = -1;
  for (let r = 12; r <= ar.getLastRow(); r++) {
    if (ar._getCell(r, 1) === "INST-1") { instRow = r; break; }
  }
  ar._setCell(instRow, 6, "🚫 Neglect / Skip This Cycle");
  sandbox.handleActiveReconEdit({ range: ar.getRange(instRow, 6) });

  // Re-import SAME October statement -> S-00001 preserved, skip override preserved
  sandbox.writeStatementToSheet(ss, octMeta, octTx);
  assert.strictEqual(stmt._getCell(7, 8), "S-00001", "Re-importing same statement must preserve S-00001");
  recon = sandbox.reconcileActiveStatement(ss, tz);
  assert.ok(recon.installmentReconItems[0].isNeglected, "Same-statement re-import must keep skip decision");

  // Import NEW November statement -> fresh S-00002, October's S-00001 skip must NOT leak!
  const novMeta = {
    statementDate: "30-11-2026",
    dueDate: "27-12-2026",
    creditLimit: 100000,
    openingBalance: 100,
    closingBalance: 200,
    totalDebit: 100,
    totalCredit: 0,
    card: "1234"
  };
  const novTx = [{ txDate: "30-11-2026", postDate: "30-11-2026", desc: "GYM 4 OF 6", type: "INSTALLMENT", amount: 100 }];
  sandbox.writeStatementToSheet(ss, novMeta, novTx);
  assert.strictEqual(stmt._getCell(7, 8), "S-00002", "New month statement must issue fresh S-00002");
  recon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(recon.installmentReconItems[0].action, "✅ Billed to Mido", "October's S-00001 skip must NOT leak onto November's S-00002");
}

// ==========================================
// 5. One-Time Legacy Migration (Clean + Ambiguous Report) & Backup Retention
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
  stmt.getRange(7, 1, 2, 7).setValues([
    [1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "SHOP A", "DEBIT", 100, ""],
    [2, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "SHOP A", "DEBIT", 100, ""]
  ]);

  // Seed legacy store with a colliding orphan entry on 2026-10-10 / 100.00 (_#3 doesn't match occ 1 or occ 2)
  const storeSheet = sandbox.getReconStoreSheet(ss);
  storeSheet.getRange(2, 1, 2, 10).setValues([
    ["TABLE2_ASSIGNMENT", "2026-10-10_100.00_shop a_#1", "2026-10-10", "SHOP A", 100, "Mido", "", false, false, "2026-11-01"],
    ["TABLE2_ASSIGNMENT", "2026-10-10_100.00_shop a_#3", "2026-10-10", "SHOP A", 100, "Mai", "", false, false, "2026-11-01"]
  ]);

  const recon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(recon.missingInSheet[0].assignedPayer, "Mido", "_#1 should migrate cleanly to Mido");
  assert.ok(ss.getSheetByName(CONFIG.SHEETS.RECON_STORE), "Legacy _Recon_Store must remain intact as read-only backup");
  const dec = sandbox.loadDecisions(ss);
  assert.strictEqual(dec.isMigrated, true);
  const report = ss.getSheetByName(CONFIG.SHEETS.MIGRATION_REPORT);
  assert.ok(report, "Ambiguous unconsumed legacy entry must generate Migration Report sheet");
  assert.strictEqual(report._getCell(3, 1), "TABLE2_ASSIGNMENT");
  assert.strictEqual(report._getCell(3, 6), "Mai");
}

console.log("ids_and_decisions.test.js: all permanent ID, isolation, and migration assertions passed.");
