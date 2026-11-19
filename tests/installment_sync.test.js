const assert = require("assert");
const { createGasSandbox, MockSpreadsheet } = require("./helpers/gas_sandbox");

const { sandbox, CONFIG } = createGasSandbox();
const tz = "Africa/Cairo";

// ==========================================
// 1. Real matchStatementInstallments Across 21 Statement Lines (Splits & Guard Protection)
// ==========================================
const sheetRows = [
  { index: 0, desc: "Kettle", emi: 61.50, duration: 12, payer: "Mido", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 1, desc: "Kettle", emi: 61.50, duration: 12, payer: "Mai", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 2, desc: "Hair dryer", emi: 22.88, duration: 12, payer: "Mai", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 3, desc: "Hair dryer", emi: 22.88, duration: 12, payer: "Mido", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 4, desc: "Hand Wrest", emi: 8.33, duration: 12, payer: "Mido", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 5, desc: "Hard drive", emi: 154.39, duration: 12, payer: "Mido", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 6, desc: "Trimmer", emi: 52.42, duration: 12, payer: "Mido", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 7, desc: "Deepfreezer", emi: 400.00, duration: 12, payer: "Dad", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 8, desc: "Deepfreezer", emi: 405.00, duration: 12, payer: "Mum", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 9, desc: "Deepfreezer", emi: 360.00, duration: 12, payer: "Mido", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 10, desc: "Deepfreezer", emi: 308.00, duration: 12, payer: "Mai", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 11, desc: "Airbods", emi: 106.17, duration: 12, payer: "Muhanad", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 12, desc: "Hand blender", emi: 33.00, duration: 12, payer: "Mai", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 13, desc: "Hand blender", emi: 15.00, duration: 12, payer: "Mum", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 14, desc: "Hand blender", emi: 45.67, duration: 12, payer: "Mido", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 15, desc: "Mobilaty", emi: 833.25, duration: 11, payer: "Mai", paymentsMade: 9, status: "Ongoing", isCompleted: false },
  { index: 16, desc: "Noon 2 of 6", emi: 132.00, duration: 2, payer: "Mido", paymentsMade: 2, status: "Completed", isCompleted: true },
  { index: 17, desc: "Dad's Phone", emi: 610.33, duration: 12, payer: "Dad", paymentsMade: 8, status: "Ongoing", isCompleted: false },
  { index: 18, desc: "Headphones", emi: 194.33, duration: 9, payer: "Mido", paymentsMade: 6, status: "Ongoing", isCompleted: false },
  { index: 19, desc: "Cosmetics", emi: 212.46, duration: 12, payer: "Mai", paymentsMade: 5, status: "Ongoing", isCompleted: false },
  { index: 20, desc: "Shaving Razors", emi: 97.50, duration: 6, payer: "Mai", paymentsMade: 4, status: "Ongoing", isCompleted: false },
  { index: 21, desc: "Muhanad's Earphones", emi: 74.07, duration: 9, payer: "Mido", paymentsMade: 4, status: "Ongoing", isCompleted: false },
  { index: 22, desc: "Muhanad's Earphones", emi: 74.07, duration: 9, payer: "Mai", paymentsMade: 4, status: "Ongoing", isCompleted: false },
  { index: 23, desc: "Muhanad's Earphones", emi: 74.07, duration: 9, payer: "Mum", paymentsMade: 4, status: "Ongoing", isCompleted: false },
  { index: 24, desc: "Watch organaizer", emi: 85.21, duration: 6, payer: "Abdo", paymentsMade: 2, status: "Ongoing", isCompleted: false },
  { index: 25, desc: "Iron", emi: 316.50, duration: 6, payer: "Abdo", paymentsMade: 2, status: "Ongoing", isCompleted: false },
  { index: 26, desc: "Heater", emi: 474.83, duration: 6, payer: "Abdo", paymentsMade: 2, status: "Ongoing", isCompleted: false },
  { index: 27, desc: "Power bank", emi: 234.33, duration: 6, payer: "Abdo", paymentsMade: 2, status: "Ongoing", isCompleted: false },
  { index: 28, desc: "Creatine", emi: 252.45, duration: 6, payer: "Mido", paymentsMade: 2, status: "Ongoing", isCompleted: false },
  { index: 29, desc: "Power bank", emi: 174.07, duration: 6, payer: "Mido", paymentsMade: 1, status: "Ongoing", isCompleted: false },
  { index: 30, desc: "Cosmatics", emi: 472.64, duration: 6, payer: "Mai", paymentsMade: 1, status: "Ongoing", isCompleted: false },
  { index: 31, desc: "Fan", emi: 80.00, duration: 6, payer: "Dad", paymentsMade: 0, status: "Ongoing", isCompleted: false },
  { index: 32, desc: "Fan", emi: 84.33, duration: 6, payer: "Mum", paymentsMade: 0, status: "Ongoing", isCompleted: false },
  { index: 33, desc: "Fan", emi: 84.33, duration: 6, payer: "Mai", paymentsMade: 0, status: "Ongoing", isCompleted: false }
];

const stmtList = [
  { desc: "KETTLE 10 OF 12", amount: 123.00, instCurrent: 10, instTotal: 12 },
  { desc: "HAIR DRYER 10 OF 12", amount: 45.76, instCurrent: 10, instTotal: 12 },
  { desc: "HAND WREST 10 OF 12", amount: 8.33, instCurrent: 10, instTotal: 12 },
  { desc: "HARD DRIVE 10 OF 12", amount: 154.39, instCurrent: 10, instTotal: 12 },
  { desc: "TRIMMER 10 OF 12", amount: 52.42, instCurrent: 10, instTotal: 12 },
  { desc: "DEEPFREEZER 10 OF 12", amount: 1473.00, instCurrent: 10, instTotal: 12 },
  { desc: "AIRBODS 10 OF 12", amount: 106.17, instCurrent: 10, instTotal: 12 },
  { desc: "HAND BLENDER 10 OF 12", amount: 93.67, instCurrent: 10, instTotal: 12 },
  { desc: "MOBILATY 10 OF 11", amount: 833.25, instCurrent: 10, instTotal: 11 },
  { desc: "DAD'S PHONE 9 OF 12", amount: 610.33, instCurrent: 9, instTotal: 12 },
  { desc: "HEADPHONES 7 OF 9", amount: 194.33, instCurrent: 7, instTotal: 9 },
  { desc: "COSMETICS 6 OF 12", amount: 212.46, instCurrent: 6, instTotal: 12 },
  { desc: "SHAVING RAZORS 5 OF 6", amount: 97.50, instCurrent: 5, instTotal: 6 },
  { desc: "MUHANAD'S EARPHONES 5 OF 9", amount: 222.21, instCurrent: 5, instTotal: 9 },
  { desc: "WATCH ORGANAIZER 3 OF 6", amount: 85.21, instCurrent: 3, instTotal: 6 },
  { desc: "IRON 3 OF 6", amount: 316.50, instCurrent: 3, instTotal: 6 },
  { desc: "HEATER 3 OF 6", amount: 474.83, instCurrent: 3, instTotal: 6 },
  { desc: "POWER BANK 3 OF 6", amount: 234.33, instCurrent: 3, instTotal: 6 },
  { desc: "CREATINE 3 OF 6", amount: 252.45, instCurrent: 3, instTotal: 6 },
  { desc: "POWER BANK 2 OF 6", amount: 174.07, instCurrent: 2, instTotal: 6 },
  { desc: "FAN 1 OF 6", amount: 248.66, instCurrent: 1, instTotal: 6 }
];

const matches = sandbox.matchStatementInstallments(stmtList, sheetRows);
assert.strictEqual(matches.length, stmtList.length, "All 21 statement installments must match");

const hairDryerMatch = matches.find(m => m.stmt.desc === "HAIR DRYER 10 OF 12");
assert.strictEqual(hairDryerMatch.type, "SPLIT_SAME_DESC", "Hair dryer must match its 2-row split group");
assert.ok(hairDryerMatch.sheetRows.every(r => r.desc === "Hair dryer"), "Hair dryer must not steal Hand blender");

const handBlenderMatch = matches.find(m => m.stmt.desc === "HAND BLENDER 10 OF 12");
assert.strictEqual(handBlenderMatch.type, "SPLIT_SAME_DESC", "Hand blender must match its 3-row split group");
assert.strictEqual(handBlenderMatch.sheetRows.length, 3);
assert.ok(handBlenderMatch.sheetRows.every(r => r.desc === "Hand blender"), "Hand blender group must stay intact");

// ==========================================
// 2. Real syncInstallmentsFromStatement (Paid, Pending, Idempotency & Completed Protection)
// ==========================================
const ss = new MockSpreadsheet();
const stmtSheet = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
stmtSheet.getRange(2, 1, 3, 6).setValues([
  ["Statement Date:", new Date(2026, 9, 30, 12, 0, 0), "Due Date:", new Date(2026, 10, 25, 12, 0, 0), "Credit Limit:", 100000],
  ["Opening Balance:", 0, "Closing Balance:", 5000, "Total Debit:", 5000],
  ["Total Credit:", 0, "Card Number:", "1234", "Net:", 5000]
]);
stmtSheet.getRange(7, 1, stmtList.length, 7).setValues(
  stmtList.map((s, idx) => [idx + 1, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), s.desc, "INSTALLMENT", s.amount, ""])
);

const instSheet = ss.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
instSheet.getRange(1, 1, 1, 11).setValues([["#", "Start Date", "Item", "Total Price", "Months", "C6", "C7", "EMI", "Payer", "Paid", "Status"]]);
instSheet.getRange(2, 1, sheetRows.length, 11).setValues(
  sheetRows.map((sh, idx) => [idx + 1, new Date(2025, 10, 15, 12, 0, 0), sh.desc, sh.emi * sh.duration, sh.duration, 0, 0, sh.emi, sh.payer, sh.paymentsMade, sh.status])
);

// Paid state sync
sandbox.syncInstallmentsFromStatement(ss, tz, true);
assert.strictEqual(instSheet._getCell(2, 10), 10, "Kettle Mido paid -> 10");
assert.strictEqual(instSheet._getCell(18, 10), 2, "Completed Noon 2 of 6 must stay 2");
assert.strictEqual(instSheet._getCell(18, 11), "Completed", "Completed Noon 2 of 6 must stay Completed");
assert.strictEqual(instSheet._getCell(32, 10), 0, "Unmatched Cosmatics with 1 -> reset to 0");
assert.strictEqual(instSheet._getCell(33, 10), 1, "Fan (1 of 6) paid -> 1");

// Idempotency check
const snapshotPaid = instSheet.getRange(2, 10, sheetRows.length, 2).getValues();
sandbox.syncInstallmentsFromStatement(ss, tz, true);
const snapshotPaidAgain = instSheet.getRange(2, 10, sheetRows.length, 2).getValues();
assert.deepStrictEqual(snapshotPaidAgain, snapshotPaid, "syncInstallmentsFromStatement must be idempotent");

// Pending state rollback
sandbox.syncInstallmentsFromStatement(ss, tz, false);
assert.strictEqual(instSheet._getCell(2, 10), 9, "Kettle Mido pending -> 9");
assert.strictEqual(instSheet._getCell(18, 10), 2, "Completed Noon 2 of 6 unaffected by pending rollback");
assert.strictEqual(instSheet._getCell(33, 10), 0, "Fan (1 of 6) pending -> 0");

console.log("installment_sync.test.js: all installment matching and sync assertions passed.");
