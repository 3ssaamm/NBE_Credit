const assert = require("assert");
const { createGasSandbox, populateSpreadsheetFromFixture, MockSpreadsheet } = require("./helpers/gas_sandbox");

const tz = "Africa/Cairo";
const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();

const ss = new MockSpreadsheet();
setActiveSpreadsheet(ss);

const stmtSheet = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
stmtSheet.getRange(2, 1, 3, 6).setValues([
  ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
  ["Opening Balance:", 10000, "Closing Balance:", 3679.21, "Total Debit:", 3729.21],
  ["Total Credit:", 10050, "Card Number:", "5264****1234", "Net:", 3679.21]
]);
stmtSheet.getRange(7, 1, 11, 7).setValues([
  [1, new Date(2026, 9, 2, 12, 0, 0), new Date(2026, 9, 2, 12, 0, 0), "FAWRY*BAZLYKWA ALEX EGY \"MOBILE TOKEN\"", "DEBIT", 1447.80, "Secret Ref 999"],
  [2, new Date(2026, 9, 8, 12, 0, 0), new Date(2026, 9, 8, 12, 0, 0), "AMAZON", "DEBIT", 833.10, ""],
  [3, new Date(2026, 9, 24, 12, 0, 0), new Date(2026, 9, 24, 12, 0, 0), "NUTOPIA-SUTHERLAND ALEX. -02G", "DEBIT", 40.00, ""],
  [4, new Date(2026, 9, 28, 12, 0, 0), new Date(2026, 9, 28, 12, 0, 0), "NUTOPIA-SUTHERLAND ALEX. -02G", "DEBIT", 40.00, ""],
  [5, new Date(2026, 9, 23, 12, 0, 0), new Date(2026, 9, 23, 12, 0, 0), "Google One 650-2530000 USACA", "DEBIT", 179.99, ""],
  [6, new Date(2026, 9, 23, 12, 0, 0), new Date(2026, 9, 23, 12, 0, 0), "Google One 650-2530000 USACA", "DEBIT", 179.99, ""],
  [7, new Date(2026, 9, 5, 12, 0, 0), new Date(2026, 9, 5, 12, 0, 0), "INSTAPAY PAYMENT - THANK YOU", "CREDIT", 10000.00, ""],
  [8, new Date(2026, 9, 12, 12, 0, 0), new Date(2026, 9, 12, 12, 0, 0), "MICROSOFT*MICROSOFT 365 REFUND", "CREDIT", 50.00, ""],
  [9, new Date(2026, 9, 31, 12, 0, 0), new Date(2026, 9, 31, 12, 0, 0), "STAMP TAX", "FEE", 41.01, ""],
  [10, new Date(2026, 9, 31, 12, 0, 0), new Date(2026, 9, 31, 12, 0, 0), "MUHANAD'S EARPHONES 5 OF 9", "INSTALLMENT", 222.21, ""],
  [11, new Date(2026, 9, 31, 12, 0, 0), new Date(2026, 9, 31, 12, 0, 0), "KETTLE 10 OF 12", "INSTALLMENT", 123.00, ""]
]);

const txSheet = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
txSheet.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
txSheet.getRange(2, 1, 6, 5).setValues([
  [1, new Date(2026, 8, 30, 0, 0, 0), "Old Fawry", "Mido", 1320.00],
  [2, new Date(2026, 9, 2, 0, 0, 0), "Basilico", "Mido", 1447.80],
  [3, new Date(2026, 9, 8, 0, 0, 0), "Amazon", "Mido", 333.10],
  [4, new Date(2026, 9, 8, 0, 0, 0), "Amazon", "Dad", 300.00],
  [5, new Date(2026, 9, 8, 0, 0, 0), "Amazon", "Muhanad", 200.00],
  [6, new Date(2026, 9, 15, 0, 0, 0), "Unmatched Shop", "Mai", 250.00]
]);

const instSheet = ss.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
instSheet.getRange(1, 1, 1, 11).setValues([["#", "Start", "Desc", "Total", "Months", "C6", "C7", "EMI", "Payer", "Paid", "Status"]]);
instSheet.getRange(2, 1, 5, 11).setValues([
  [1, new Date(2026, 3, 18, 12, 0, 0), "Muhanad's Earphones", 666.63, 9, 0, 0, 74.07, "Mido", 4, "Ongoing"],
  [2, new Date(2026, 3, 18, 12, 0, 0), "Muhanad's Earphones", 666.63, 9, 0, 0, 74.07, "Mai", 4, "Ongoing"],
  [3, new Date(2026, 3, 18, 12, 0, 0), "Muhanad's Earphones", 666.63, 9, 0, 0, 74.07, "Mum", 4, "Ongoing"],
  [4, new Date(2025, 10, 7, 12, 0, 0), "Kettle", 738.00, 12, 0, 0, 61.50, "Mido", 9, "Ongoing"],
  [5, new Date(2025, 10, 7, 12, 0, 0), "Kettle", 738.00, 12, 0, 0, 61.50, "Mai", 9, "Ongoing"]
]);

const storeSheet = sandbox.getReconStoreSheet(ss);
storeSheet.getRange(2, 1, 7, 10).setValues([
  ["TABLE2_ASSIGNMENT", "2026-10-24_40.00_nutopia-sutherland alex. -02g_#1", "2026-10-24", "NUTOPIA-SUTHERLAND ALEX. -02G", 40.00, "Mai", "private note 1", false, false, "2026-11-17"],
  ["TABLE2_ASSIGNMENT", "2026-10-28_40.00_nutopia-sutherland alex. -02g_#1", "2026-10-28", "NUTOPIA-SUTHERLAND ALEX. -02G", 40.00, "Mai", "private note 2", false, false, "2026-11-17"],
  ["TABLE2_ASSIGNMENT", "2026-10-23_179.99_google one 650-2530000 usaca_#1", "2026-10-23", "Google One 650-2530000 USACA", 179.99, "Mido", "private note 3", false, false, "2026-11-17"],
  ["TABLE2_ASSIGNMENT", "2026-10-23_179.99_google one 650-2530000 usaca_#2", "2026-10-23", "Google One 650-2530000 USACA", 179.99, "Muhanad", "private note 4", false, false, "2026-11-17"],
  ["TABLE2_ASSIGNMENT", "2026-10-12_50.00_microsoft*microsoft 365 refund_#1", "2026-10-12", "MICROSOFT*MICROSOFT 365 REFUND", 50.00, "Mido", "private note 5", false, false, "2026-11-17"],
  ["TABLE4_INSTALLMENT", "2026-10-31_74.07_muhanad's earphones (5 of 9)_#2", "2026-10-31", "Muhanad's Earphones (5 of 9)", 74.07, "🚫 Neglect / Skip This Cycle", "private note 6", false, true, "2026-11-17"],
  ["TABLE3_NEGLECTED", "2026-09-30_1320.00_old fawry_#1", "", "", 0, "", "", false, true, "2026-11-17"]
]);

const fixture = sandbox.buildAnonymizedFixture(ss, tz);
const serialized = JSON.stringify(fixture);

// 1. Strict Privacy / Anonymization Assertions
const forbiddenTokens = [
  "mido", "muhanad", "dad", "mum", "mai",
  "bazlykwa", "basilico", "amazon", "nutopia", "sutherland",
  "google", "microsoft", "instapay", "fawry", "kettle", "earphones",
  "5264", "650-2530000", "2530000", "private note", "secret ref"
];
forbiddenTokens.forEach(tok => {
  assert.ok(
    !serialized.toLowerCase().includes(tok),
    `Anonymized fixture leaked sensitive token "${tok}"!`
  );
});
assert.strictEqual(fixture.bankStatement.meta.card, "****0000");

// 2. Lossless Replay Assertion in a Fresh Sandbox
const replayEnv = createGasSandbox();
const ssReplay = new MockSpreadsheet();
replayEnv.setActiveSpreadsheet(ssReplay);
populateSpreadsheetFromFixture(ssReplay, fixture, replayEnv.CONFIG);

const replayRecon = replayEnv.sandbox.reconcileActiveStatement(ssReplay, tz);
const replayLedger = replayEnv.sandbox.buildDebtLedger(ssReplay, tz, replayRecon);

assert.strictEqual(Number(replayLedger.closingBill.toFixed(2)), fixture.expectedTotals.closingBill);
assert.strictEqual(Number(replayLedger.purchasesTotal.toFixed(2)), fixture.expectedTotals.purchasesTotal);
assert.strictEqual(Number(replayLedger.installmentsTotal.toFixed(2)), fixture.expectedTotals.installmentsTotal);
assert.strictEqual(Number(replayLedger.feesTotal.toFixed(2)), fixture.expectedTotals.feesTotal);
assert.strictEqual(Number(replayLedger.unrecordedTotal.toFixed(2)), fixture.expectedTotals.unrecordedTotal);
assert.strictEqual(Number(replayLedger.refundsTotal.toFixed(2)), fixture.expectedTotals.refundsTotal);
assert.strictEqual(Number(replayLedger.familySharesTotal.toFixed(2)), fixture.expectedTotals.familySharesTotal);
assert.strictEqual(Number(replayLedger.auditDifference.toFixed(2)), fixture.expectedTotals.auditDifference);

Object.keys(fixture.expectedShares).forEach(p => {
  const actual = replayLedger.personShares.find(ps => ps.person === p);
  assert.ok(actual, `Replayed ledger missing anonymized person ${p}`);
  const exp = fixture.expectedShares[p];
  ["purchases", "installments", "fees", "unrecorded", "refunds", "netTotal"].forEach(k => {
    assert.strictEqual(
      Number(actual[k].toFixed(2)),
      exp[k],
      `Replay mismatch for ${p}.${k}: expected ${exp[k]}, got ${actual[k]}`
    );
  });
});

console.log("fixture_export.test.js: anonymization and lossless replay assertions passed.");
