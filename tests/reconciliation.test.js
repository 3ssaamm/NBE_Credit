const assert = require("assert");
const { createGasSandbox, MockSpreadsheet, evalTable4Sumifs } = require("./helpers/gas_sandbox");

const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
const tz = "Africa/Cairo";

// ==========================================
// Build Realistic Spreadsheet Fixtures
// ==========================================
const ss = new MockSpreadsheet();

const stmtSheet = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
stmtSheet.getRange(2, 1, 3, 6).setValues([
  ["Statement Date:", new Date("2026-09-30T12:00:00"), "Due Date:", new Date("2026-10-25T12:00:00"), "Credit Limit:", 100000],
  ["Opening Balance:", 0, "Closing Balance:", 5264.61, "Total Debit:", 5264.61],
  ["Total Credit:", 0, "Card Number:", "1234", "Net:", 5264.61]
]);

const stmtRows = [
  [1, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "HARD DRIVE 10 OF 12", "INSTALLMENT", 154.39, ""],
  [2, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "DEEPFREEZER 10 OF 12", "INSTALLMENT", 1473.00, ""],
  [3, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "HAND BLENDER 10 OF 12", "INSTALLMENT", 93.67, ""],
  [4, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "MUHANAD'S EARPHONES 5 OF 9", "INSTALLMENT", 222.21, ""],
  [5, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "MOBILATY 10 OF 11", "INSTALLMENT", 833.25, ""],
  [6, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "MYSTERY GADGET 1 OF 6", "INSTALLMENT", 350.00, ""]
];
stmtSheet.getRange(7, 1, stmtRows.length, 7).setValues(stmtRows);

const instSheet = ss.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
const instRows = [
  [1, new Date("2025-11-15T12:00:00"), "Hard drive", 1852.68, 12, 0, 0, 154.39, "Mido", 9, "Ongoing"],
  [2, new Date("2025-11-15T12:00:00"), "Deepfreezer", 4800, 12, 0, 0, 400.00, "Dad", 9, "Ongoing"],
  [3, new Date("2025-11-15T12:00:00"), "Deepfreezer", 4860, 12, 0, 0, 405.00, "Mum", 9, "Ongoing"],
  [4, new Date("2025-11-15T12:00:00"), "Deepfreezer", 4320, 12, 0, 0, 360.00, "Mido", 9, "Ongoing"],
  [5, new Date("2025-11-15T12:00:00"), "Deepfreezer", 3696, 12, 0, 0, 308.00, "Mai", 9, "Ongoing"],
  [6, new Date("2025-11-15T12:00:00"), "Hand blender", 396, 12, 0, 0, 33.00, "Mai", 9, "Ongoing"],
  [7, new Date("2025-11-15T12:00:00"), "Hand blender", 180, 12, 0, 0, 15.00, "Mum", 9, "Ongoing"],
  [8, new Date("2025-11-15T12:00:00"), "Hand blender", 548.04, 12, 0, 0, 45.67, "Mido", 9, "Ongoing"],
  [9, new Date("2026-04-10T12:00:00"), "Muhanad's Earphones", 666.63, 9, 0, 0, 74.07, "Mido", 4, "Ongoing"],
  [10, new Date("2026-04-10T12:00:00"), "Muhanad's Earphones", 666.63, 9, 0, 0, 74.07, "Mai", 4, "Ongoing"],
  [11, new Date("2026-04-10T12:00:00"), "Muhanad's Earphones", 666.63, 9, 0, 0, 74.07, "Mum", 4, "Ongoing"],
  [12, new Date("2025-11-15T12:00:00"), "Mobilaty", 9165.20, 11, 0, 0, 833.20, "Mai", 9, "Ongoing"],
  [13, new Date("2026-09-01T12:00:00"), "Cosmatics", 2835.84, 6, 0, 0, 472.64, "Mai", 0, "Ongoing"],
  [14, new Date("2025-12-01T12:00:00"), "Noon 2 of 6", 264.00, 2, 0, 0, 132.00, "Mido", 2, "Completed"]
];
instSheet.getRange(2, 1, instRows.length, 11).setValues(instRows);

// ==========================================
// Test 1: Reconciliation, Discrepancy Ordering & Unequal Splits
// ==========================================
const recon1 = sandbox.reconcileActiveStatement(ss, tz);
const items1 = recon1.installmentReconItems;

assert.strictEqual(items1.length, 14, `Expected 14 Table 4 items, got ${items1.length}`);
assert.strictEqual(items1[0].sourceCategory, "⚠️ Statement Only", "Discrepancy Statement Only must be first");
assert.strictEqual(items1[0].action, "", "Statement Only default action must be blank");
assert.strictEqual(items1[1].sourceCategory, "⏳ Sheet Only", "Sheet Only must appear before matched items");
assert.strictEqual(items1[1].action, "🚫 Not on Statement (Skip)", "Sheet Only default action must skip");
assert.ok(items1[1].desc.includes("Cosmatics"), "Cosmatics must be in Sheet Only");
assert.ok(!items1.some(it => it.desc.includes("Noon 2 of 6")), "Completed item must NOT be in Sheet Only");

const mobilatyItem = items1.find(it => it.desc.includes("Mobilaty"));
assert.ok(mobilatyItem, "Mobilaty must be present");
assert.strictEqual(mobilatyItem.amount, 833.25, "Mobilaty amount should adjust to statement 833.25");
assert.ok(mobilatyItem.note.includes("⚠️ Stmt: 833.25 vs Sheet: 833.20"), "Mobilaty note should surface Stmt vs Sheet diff");

const ledger1 = sandbox.buildDebtLedger(ss, tz, recon1);
const shareMap1 = {};
ledger1.personShares.forEach(ps => { shareMap1[ps.person] = Number(ps.installments.toFixed(2)); });

assert.strictEqual(shareMap1["Dad"], 400.00, `Dad share mismatch: ${shareMap1["Dad"]}`);
assert.strictEqual(shareMap1["Mum"], 494.07, `Mum share mismatch: ${shareMap1["Mum"]}`);
assert.strictEqual(shareMap1["Mido"], 634.13, `Mido share mismatch: ${shareMap1["Mido"]}`);
assert.strictEqual(shareMap1["Mai"], 1248.32, `Mai share mismatch: ${shareMap1["Mai"]}`);

// ==========================================
// Test 2: Rendering Table 4 & Live SUMIFS Equivalence
// ==========================================
sandbox.renderActiveReconciliationSheet(ss, tz, ledger1, recon1);
const reconSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);

let startInstRow = -1;
let endInstRow = -1;
for (let r = 1; r <= reconSheet.getLastRow(); r++) {
  const colA = String(reconSheet._getCell(r, 1) || "");
  if (/^INST-\d+/i.test(colA)) {
    if (startInstRow === -1) startInstRow = r;
    endInstRow = r;
  }
}
assert.ok(startInstRow > 19, "Table 4 should be rendered below Table 1 & Table 2");
assert.strictEqual(endInstRow - startInstRow + 1, 14, "Table 4 should have 14 rows");

for (let i = 0; i < ledger1.personShares.length; i++) {
  const r = 12 + i;
  const person = reconSheet._getCell(r, 2);
  const formula = String(reconSheet._getCell(r, 4) || "");
  assert.ok(formula.startsWith("=SUMIFS("), `Table 1 Col D row ${r} must be SUMIFS formula, got: ${formula}`);
  assert.ok(formula.includes(`$D$${startInstRow}:$D$${endInstRow}`), "SUMIFS must reference Table 4 Col D");
  const evalSum = evalTable4Sumifs(reconSheet, person, startInstRow, endInstRow);
  assert.strictEqual(evalSum, shareMap1[person], `SUMIFS eval for ${person} (${evalSum}) must match ledger (${shareMap1[person]})`);
}

// ==========================================
// Test 3: onEdit Handling & Occurrence-Safe Persistence
// ==========================================
const rowInst1 = startInstRow;
reconSheet._setCell(rowInst1, 6, "Abdo");
sandbox.handleActiveReconEdit({ range: reconSheet.getRange(rowInst1, 6) });
assert.strictEqual(reconSheet._getCell(rowInst1, 6), "✅ Billed to Abdo", "Bare name should normalize to '✅ Billed to Abdo'");

let earphonesRow2 = -1;
for (let r = startInstRow; r <= endInstRow; r++) {
  if (String(reconSheet._getCell(r, 3)).includes("Muhanad's Earphones") &&
      String(reconSheet._getCell(r, 6)).includes("Mai")) {
    earphonesRow2 = r;
    break;
  }
}
assert.ok(earphonesRow2 > 0, "Should locate Mai's Earphones row");
reconSheet._setCell(earphonesRow2, 6, "🚫 Neglect / Skip This Cycle");
sandbox.handleActiveReconEdit({ range: reconSheet.getRange(earphonesRow2, 6) });

const rowCosmatics = startInstRow + 1;
reconSheet._setCell(rowCosmatics, 6, "✅ Billed to Mai");
sandbox.handleActiveReconEdit({ range: reconSheet.getRange(rowCosmatics, 6) });

const recon2 = sandbox.reconcileActiveStatement(ss, tz);
const ledger2 = sandbox.buildDebtLedger(ss, tz, recon2);
const shareMap2 = {};
ledger2.personShares.forEach(ps => { shareMap2[ps.person] = Number(ps.installments.toFixed(2)); });

const earphonesItems2 = recon2.installmentReconItems.filter(it => it.desc.includes("Muhanad's Earphones"));
assert.strictEqual(earphonesItems2.length, 3);
assert.strictEqual(earphonesItems2[0].action, "✅ Billed to Mido", "occ=1 Mido must remain billed!");
assert.strictEqual(earphonesItems2[1].action, "🚫 Neglect / Skip This Cycle", "occ=2 Mai must persist as neglected!");
assert.strictEqual(earphonesItems2[2].action, "✅ Billed to Mum", "occ=3 Mum must remain billed!");

assert.strictEqual(shareMap2["Abdo"], 350.00, `Abdo share after assign mismatch: ${shareMap2["Abdo"]}`);
assert.strictEqual(shareMap2["Mai"], 1646.89, `Mai share after neglect+unskip mismatch: ${shareMap2["Mai"]}`);
assert.strictEqual(shareMap2["Mido"], 634.13, `Mido share should be unchanged: ${shareMap2["Mido"]}`);
assert.strictEqual(shareMap2["Mum"], 494.07, `Mum share should be unchanged: ${shareMap2["Mum"]}`);

sandbox.renderActiveReconciliationSheet(ss, tz, ledger2, recon2);
let startInstRow2 = -1, endInstRow2 = -1;
for (let r = 1; r <= reconSheet.getLastRow(); r++) {
  if (/^INST-\d+/i.test(String(reconSheet._getCell(r, 1) || ""))) {
    if (startInstRow2 === -1) startInstRow2 = r;
    endInstRow2 = r;
  }
}
for (let i = 0; i < ledger2.personShares.length; i++) {
  const r = 12 + i;
  const person = reconSheet._getCell(r, 2);
  const evalSum = evalTable4Sumifs(reconSheet, person, startInstRow2, endInstRow2);
  assert.strictEqual(evalSum, shareMap2[person], `Post-refresh SUMIFS for ${person} (${evalSum}) != ${shareMap2[person]}`);
}

// ==========================================
// Test 4: Co-existence with Table 3, Rounding Residual, Clearing Override & syncInstallments
// ==========================================
const txSheet = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
txSheet.getRange(2, 1, 1, 5).setValues([
  [1, new Date("2026-09-15T12:00:00"), "Unmatched Coffee", "Mido", 120.00]
]);

stmtSheet.getRange(13, 1, 1, 7).setValues([
  [7, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "SHARED GYM 2 OF 3", "INSTALLMENT", 100.00, ""]
]);
instSheet.getRange(16, 1, 1, 11).setValues([
  [15, new Date("2026-07-15T12:00:00"), "Shared Gym", 300.00, 3, 0, 0, 100.00, "Mido + Mai + Abdo", 1, "Ongoing"]
]);

reconSheet._setCell(startInstRow2, 6, "");
sandbox.handleActiveReconEdit({ range: reconSheet.getRange(startInstRow2, 6) });

const recon4 = sandbox.reconcileActiveStatement(ss, tz);
const ledger4 = sandbox.buildDebtLedger(ss, tz, recon4);
sandbox.renderActiveReconciliationSheet(ss, tz, ledger4, recon4);

let orphanRowsCount = 0;
let instRowsCount = 0;
let startInst4 = -1, endInst4 = -1;
for (let r = 1; r <= reconSheet.getLastRow(); r++) {
  const id = String(reconSheet._getCell(r, 1) || "");
  if (/^ORPH-/i.test(id)) orphanRowsCount++;
  if (/^INST-\d+/i.test(id)) {
    if (startInst4 === -1) startInst4 = r;
    endInst4 = r;
    instRowsCount++;
  }
}
assert.strictEqual(orphanRowsCount, 1, "Orphan Logs panel must have 1 row");
assert.strictEqual(instRowsCount, 17, `Table 4 must have 17 rows (14 + 3 Shared Gym), got ${instRowsCount}`);
const mido4 = ledger4.personShares.find(ps => ps.person === "Mido");
assert.strictEqual(Number(mido4.purchases.toFixed(2)), 0.00, "Unpaired in-period orphan log must NOT be billed on current statement (Q5c)");

const gymItems = recon4.installmentReconItems.filter(it => it.desc.includes("Shared Gym"));
assert.strictEqual(gymItems.length, 3);
const gymSum = Number(gymItems.reduce((s, it) => s + it.amount, 0).toFixed(2));
assert.strictEqual(gymSum, 100.00, `Shared Gym 3-way split must sum to 100.00, got ${gymSum}`);
assert.ok(gymItems.every(it => !it.note.includes("⚠️")), "Exact 100/3 rounding should not emit false ⚠️ diff warning");

assert.strictEqual(recon4.installmentReconItems[0].action, "", "Cleared INST-1 override should revert to blank");

for (let i = 0; i < ledger4.personShares.length; i++) {
  const r = 12 + i;
  const person = reconSheet._getCell(r, 2);
  const expected = Number(ledger4.personShares[i].installments.toFixed(2));
  const evalSum = evalTable4Sumifs(reconSheet, person, startInst4, endInst4);
  assert.strictEqual(evalSum, expected, `Table 3+4 SUMIFS mismatch for ${person}: ${evalSum} != ${expected}`);
}

sandbox.syncInstallmentsFromStatement(ss, tz, false);
assert.strictEqual(instSheet._getCell(2, 10), 9, "Hard drive (10 of 12, unpaid) should sync to 9 payments made");

// ==========================================
// Test 5: Multi-Date NUTOPIA Persistence, Legacy _#k Recovery & Same-Date Isolation
// ==========================================
const ss5 = new MockSpreadsheet();
setActiveSpreadsheet(ss5);

const stmt5 = ss5.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
stmt5.getRange(2, 1, 3, 6).setValues([
  ["Statement Date:", new Date(2026, 9, 30, 12, 0, 0), "Due Date:", new Date(2026, 10, 25, 12, 0, 0), "Credit Limit:", 100000],
  ["Opening Balance:", 0, "Closing Balance:", 3200.00, "Total Debit:", 3209.27],
  ["Total Credit:", 9.27, "Card Number:", "1234", "Net:", 3200.00]
]);
stmt5.getRange(7, 1, 12, 7).setValues([
  [1, new Date(2026, 9, 28, 12, 0, 0), new Date(2026, 9, 28, 12, 0, 0), "NUTOPIA-SUTHERLAND ALEX. -02G", "DEBIT", 40.00, ""],
  [2, new Date(2026, 9, 24, 12, 0, 0), new Date(2026, 9, 24, 12, 0, 0), "NUTOPIA-SUTHERLAND ALEX. -02G", "DEBIT", 40.00, ""],
  [3, new Date(2026, 9, 21, 12, 0, 0), new Date(2026, 9, 21, 12, 0, 0), "NUTOPIA-SUTHERLAND ALEX. -02G", "DEBIT", 40.00, ""],
  [4, new Date(2026, 9, 28, 12, 0, 0), new Date(2026, 9, 28, 12, 0, 0), "NUTOPIA-SUTHERLAND ALEX. -02G", "DEBIT", 80.00, ""],
  [5, new Date(2026, 9, 12, 12, 0, 0), new Date(2026, 9, 12, 12, 0, 0), "NUTOPIA-SUTHERLAND ALEX. -02G", "DEBIT", 80.00, ""],
  [6, new Date(2026, 9, 22, 12, 0, 0), new Date(2026, 9, 22, 12, 0, 0), "NUTOPIA-SUTHERLAND ALEX. -02G", "DEBIT", 85.00, ""],
  [7, new Date(2026, 9, 20, 12, 0, 0), new Date(2026, 9, 20, 12, 0, 0), "NUTOPIA-SUTHERLAND ALEX. -02G", "DEBIT", 85.00, ""],
  [8, new Date(2026, 9, 23, 12, 0, 0), new Date(2026, 9, 23, 12, 0, 0), "Google One 650-2530000 USACA", "DEBIT", 179.99, ""],
  [9, new Date(2026, 9, 23, 12, 0, 0), new Date(2026, 9, 23, 12, 0, 0), "Google One 650-2530000 USACA", "DEBIT", 179.99, ""],
  [10, new Date(2026, 9, 12, 12, 0, 0), new Date(2026, 9, 12, 12, 0, 0), "MICROSOFT*MICROSOFT 365 BMSE", "CREDIT", 9.00, ""],
  [11, new Date(2026, 9, 12, 12, 0, 0), new Date(2026, 9, 12, 12, 0, 0), "DCC Markup Refund-MICROSO*MICR", "CREDIT", 0.27, ""],
  [12, new Date(2026, 9, 3, 12, 0, 0), new Date(2026, 9, 3, 12, 0, 0), "FAWRY*THE H CAF", "DEBIT", 2316.00, ""]
]);

const tx5 = ss5.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
tx5.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
tx5.getRange(2, 1, 1, 5).setValues([[307, new Date(2026, 9, 3, 0, 0, 0), "Fawry", "Mido", 2316.00]]);

const store5 = sandbox.getReconStoreSheet(ss5);
store5.getRange(2, 1, 5, 10).setValues([
  ["TABLE2_ASSIGNMENT", "2026-10-24_40.00_nutopia-sutherland alex. -02g_#2", "2026-10-24", "NUTOPIA-SUTHERLAND ALEX. -02G", 40.00, "Mai", "nutopia", false, false, "2026-11-17"],
  ["TABLE2_ASSIGNMENT", "2026-10-21_40.00_nutopia-sutherland alex. -02g_#3", "2026-10-21", "NUTOPIA-SUTHERLAND ALEX. -02G", 40.00, "Mai", "nutopia", false, false, "2026-11-17"],
  ["TABLE2_ASSIGNMENT", "2026-10-12_80.00_nutopia-sutherland alex. -02g_#2", "2026-10-12", "NUTOPIA-SUTHERLAND ALEX. -02G", 80.00, "Mai", "nutopia", false, false, "2026-11-17"],
  ["TABLE2_ASSIGNMENT", "2026-10-20_85.00_nutopia-sutherland alex. -02g_#2", "2026-10-20", "NUTOPIA-SUTHERLAND ALEX. -02G", 85.00, "Mai", "nutopia", false, false, "2026-11-17"],
  ["TABLE2_ASSIGNMENT", "2026-10-23_179.99_google one 650-2530000 usaca_#2", "2026-10-23", "Google One 650-2530000 USACA", 179.99, "Muhanad", "google one", false, false, "2026-11-17"]
]);

sandbox.updateLiveDashboard();
const reconSheet5 = ss5.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);

function readTable2Rows(sheet) {
  const rows = [];
  for (let r = 12; r <= sheet.getLastRow(); r++) {
    const id = String(sheet._getCell(r, 1) || "");
    if (/^[FCM]-\d+/i.test(id)) {
      rows.push({
        row: r,
        id,
        date: sheet._getCell(r, 2),
        desc: String(sheet._getCell(r, 3) || ""),
        amount: parseFloat(sheet._getCell(r, 4)) || 0,
        category: String(sheet._getCell(r, 5) || ""),
        payer: String(sheet._getCell(r, 6) || "")
      });
    }
  }
  return rows;
}

const t2AfterRecovery = readTable2Rows(reconSheet5);
assert.ok(!t2AfterRecovery.some(r => Math.abs(r.amount - 2316.00) < 0.01), "FAWRY*THE H CAF must match Fawry via Pass 1");

const recoveredNutopia = t2AfterRecovery.filter(r => r.desc.includes("NUTOPIA"));
assert.strictEqual(recoveredNutopia.length, 7, "All 7 NUTOPIA charges must appear in Table 2");
["2026-10-24", "2026-10-21", "2026-10-12", "2026-10-20"].forEach(dStr => {
  const match = recoveredNutopia.find(r => sandbox.makeNormDateStr(r.date, tz) === dStr);
  assert.ok(match, `Should find NUTOPIA row for ${dStr}`);
  assert.strictEqual(match.payer, "Mai", `Legacy _#k NUTOPIA row on ${dStr} should auto-recover to Mai`);
});

const googleRows5 = t2AfterRecovery.filter(r => r.desc.includes("Google One"));
assert.strictEqual(googleRows5.length, 2);
assert.strictEqual(googleRows5[0].payer, "", "Google One occ=1 must NOT be clobbered by occ=2");
assert.strictEqual(googleRows5[1].payer, "Muhanad", "Google One occ=2 must remain Muhanad");

const nutopiaOct28_40 = recoveredNutopia.find(r => sandbox.makeNormDateStr(r.date, tz) === "2026-10-28" && r.amount === 40);
const nutopiaOct28_80 = recoveredNutopia.find(r => sandbox.makeNormDateStr(r.date, tz) === "2026-10-28" && r.amount === 80);
const nutopiaOct22_85 = recoveredNutopia.find(r => sandbox.makeNormDateStr(r.date, tz) === "2026-10-22" && r.amount === 85);
reconSheet5._setCell(nutopiaOct28_40.row, 6, "Mai");
reconSheet5._setCell(nutopiaOct28_80.row, 6, "Mai");
reconSheet5._setCell(nutopiaOct22_85.row, 6, "Mai");
sandbox.handleActiveReconEdit({ range: reconSheet5.getRange(nutopiaOct28_40.row, 6) });
sandbox.handleActiveReconEdit({ range: reconSheet5.getRange(nutopiaOct28_80.row, 6) });
sandbox.handleActiveReconEdit({ range: reconSheet5.getRange(nutopiaOct22_85.row, 6) });

const creditRows5 = t2AfterRecovery.filter(r => r.category === "Refund / Credit");
assert.strictEqual(creditRows5.length, 2);
reconSheet5._setCell(creditRows5[0].row, 6, "Mido");
reconSheet5._setCell(creditRows5[1].row, 6, "Mido");
sandbox.handleActiveReconEdit({ range: reconSheet5.getRange(creditRows5[0].row, 6, 2, 1) });

reconSheet5._setCell(googleRows5[0].row, 6, "Mido");
sandbox.handleActiveReconEdit({ range: reconSheet5.getRange(googleRows5[0].row, 6) });

sandbox.updateLiveDashboard();
const t2AfterRefresh = readTable2Rows(reconSheet5);
const nutopiaFinal = t2AfterRefresh.filter(r => r.desc.includes("NUTOPIA"));
assert.strictEqual(nutopiaFinal.length, 7);
assert.ok(nutopiaFinal.every(r => r.payer === "Mai"), "Every single NUTOPIA row across all dates must retain Mai after refresh!");

const googleFinal = t2AfterRefresh.filter(r => r.desc.includes("Google One"));
assert.strictEqual(googleFinal[0].payer, "Mido", "Google One occ=1 must retain Mido");
assert.strictEqual(googleFinal[1].payer, "Muhanad", "Google One occ=2 must retain Muhanad");

const creditsFinal = t2AfterRefresh.filter(r => r.category === "Refund / Credit");
assert.ok(creditsFinal.every(r => r.payer === "Mido"), "Both credit rows from multi-row edit must retain Mido");

// ==========================================
// Test 6: Member Breakdown Midnight Date Boundary, Installment Parity (1,484.61) & Section 3
// ==========================================
const ss6 = new MockSpreadsheet();
setActiveSpreadsheet(ss6);

const stmt6 = ss6.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
stmt6.getRange(2, 1, 3, 6).setValues([
  ["Statement Date:", new Date(2026, 9, 30, 12, 0, 0), "Due Date:", new Date(2026, 10, 25, 12, 0, 0), "Credit Limit:", 100000],
  ["Opening Balance:", 0, "Closing Balance:", 9000.00, "Total Debit:", 9000.00],
  ["Total Credit:", 0, "Card Number:", "1234", "Net:", 9000.00]
]);
stmt6.getRange(7, 1, 23, 7).setValues([
  [1, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "KETTLE 10 OF 12", "INSTALLMENT", 123.00, ""],
  [2, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "HAIR DRYER 10 OF 12", "INSTALLMENT", 45.76, ""],
  [3, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "HAND WREST 10 OF 12", "INSTALLMENT", 8.33, ""],
  [4, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "HARD DRIVE 10 OF 12", "INSTALLMENT", 154.39, ""],
  [5, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "TRIMMER 10 OF 12", "INSTALLMENT", 52.42, ""],
  [6, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "DEEPFREEZER 10 OF 12", "INSTALLMENT", 1473.00, ""],
  [7, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "HAND BLENDER 10 OF 12", "INSTALLMENT", 93.67, ""],
  [8, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "HEADPHONES 7 OF 9", "INSTALLMENT", 194.33, ""],
  [9, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "MUHANAD'S EARPHONES 5 OF 9", "INSTALLMENT", 222.21, ""],
  [10, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "CREATINE 3 OF 6", "INSTALLMENT", 252.45, ""],
  [11, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "POWER BANK 2 OF 6", "INSTALLMENT", 174.07, ""],
  [12, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "FAN 1 OF 6", "INSTALLMENT", 248.81, ""],
  [13, new Date(2026, 9, 23, 12, 0, 0), new Date(2026, 9, 23, 12, 0, 0), "Google One 650-2530000 USACA", "DEBIT", 179.99, ""],
  [14, new Date(2026, 8, 30, 12, 0, 0), new Date(2026, 8, 30, 12, 0, 0), "FAWRY ROLLOVER", "DEBIT", 1320.00, ""],
  [15, new Date(2026, 9, 2, 12, 0, 0), new Date(2026, 9, 2, 12, 0, 0), "FAWRY*BAZLYKWA", "DEBIT", 1447.80, ""],
  [16, new Date(2026, 9, 3, 12, 0, 0), new Date(2026, 9, 3, 12, 0, 0), "WAVE STORE", "DEBIT", 570.00, ""],
  [17, new Date(2026, 9, 3, 12, 0, 0), new Date(2026, 9, 3, 12, 0, 0), "FAWRY STORE", "DEBIT", 2316.00, ""],
  [18, new Date(2026, 9, 6, 12, 0, 0), new Date(2026, 9, 6, 12, 0, 0), "MOSAAB CAFE", "DEBIT", 80.00, ""],
  [19, new Date(2026, 9, 8, 12, 0, 0), new Date(2026, 9, 8, 12, 0, 0), "AMAZON", "DEBIT", 333.10, ""],
  [20, new Date(2026, 9, 14, 12, 0, 0), new Date(2026, 9, 14, 12, 0, 0), "WE-FBB INTERNET", "DEBIT", 367.00, ""],
  [21, new Date(2026, 9, 23, 12, 0, 0), new Date(2026, 9, 23, 12, 0, 0), "AMAZON", "DEBIT", 164.99, ""],
  [22, new Date(2026, 9, 29, 12, 0, 0), new Date(2026, 9, 29, 12, 0, 0), "EL ABD PASTRY", "DEBIT", 250.00, ""],
  [23, new Date(2026, 9, 20, 12, 0, 0), new Date(2026, 9, 20, 12, 0, 0), "VIA CAFFE", "DEBIT", 97.00, ""]
]);

const inst6 = ss6.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
inst6.getRange(1, 1, 1, 11).setValues([["#", "Start Date", "Description", "Total Price", "Months", "C6", "C7", "EMI", "Payer", "Paid", "Status"]]);
inst6.getRange(2, 1, 21, 11).setValues([
  [1, new Date(2025, 10, 7, 12, 0, 0), "Kettle", 738, 12, 0, 0, 61.50, "Mido", 9, "Ongoing"],
  [2, new Date(2025, 10, 7, 12, 0, 0), "Kettle", 738, 12, 0, 0, 61.50, "Mai", 9, "Ongoing"],
  [3, new Date(2025, 10, 7, 12, 0, 0), "Hair dryer", 274.50, 12, 0, 0, 22.88, "Mai", 9, "Ongoing"],
  [4, new Date(2025, 10, 7, 12, 0, 0), "Hair dryer", 274.50, 12, 0, 0, 22.88, "Mido", 9, "Ongoing"],
  [5, new Date(2025, 10, 7, 12, 0, 0), "Hand Wrest", 100, 12, 0, 0, 8.33, "Mido", 9, "Ongoing"],
  [6, new Date(2025, 10, 7, 12, 0, 0), "Hard drive", 1852.70, 12, 0, 0, 154.39, "Mido", 9, "Ongoing"],
  [7, new Date(2025, 10, 7, 12, 0, 0), "Trimmer", 629, 12, 0, 0, 52.42, "Mido", 9, "Ongoing"],
  [8, new Date(2025, 10, 11, 12, 0, 0), "Deepfreezer", 4800, 12, 0, 0, 400.00, "Dad", 9, "Ongoing"],
  [9, new Date(2025, 10, 11, 12, 0, 0), "Deepfreezer", 4860, 12, 0, 0, 405.00, "Mum", 9, "Ongoing"],
  [10, new Date(2025, 10, 11, 12, 0, 0), "Deepfreezer", 4320, 12, 0, 0, 360.00, "Mido", 9, "Ongoing"],
  [11, new Date(2025, 10, 11, 12, 0, 0), "Deepfreezer", 3696, 12, 0, 0, 308.00, "Mai", 9, "Ongoing"],
  [12, new Date(2025, 10, 27, 12, 0, 0), "Hand blender", 396, 12, 0, 0, 33.00, "Mai", 9, "Ongoing"],
  [13, new Date(2025, 10, 27, 12, 0, 0), "Hand blender", 180, 12, 0, 0, 15.00, "Mum", 9, "Ongoing"],
  [14, new Date(2025, 10, 27, 12, 0, 0), "Hand blender", 548, 12, 0, 0, 45.67, "Mido", 9, "Ongoing"],
  [15, new Date(2026, 1, 3, 12, 0, 0), "Headphones", 1748.97, 9, 0, 0, 194.33, "Mido", 6, "Ongoing"],
  [16, new Date(2026, 3, 18, 12, 0, 0), "Muhanad's Earphones", 666.67, 9, 0, 0, 74.07, "Mido", 4, "Ongoing"],
  [17, new Date(2026, 3, 18, 12, 0, 0), "Muhanad's Earphones", 666.67, 9, 0, 0, 74.07, "Mai", 4, "Ongoing"],
  [18, new Date(2026, 3, 18, 12, 0, 0), "Muhanad's Earphones", 666.67, 9, 0, 0, 74.07, "Mum", 4, "Ongoing"],
  [19, new Date(2026, 5, 26, 12, 0, 0), "Creatine", 1514.70, 6, 0, 0, 252.45, "Mido", 2, "Ongoing"],
  [20, new Date(2026, 6, 18, 12, 0, 0), "Power bank", 1044.44, 6, 0, 0, 174.07, "Mido", 1, "Ongoing"],
  [21, new Date(2026, 7, 28, 12, 0, 0), "Fan", 507, 6, 0, 0, 84.50, "Mido", 0, "Ongoing"]
]);
inst6.getRange(23, 1, 2, 11).setValues([
  [22, new Date(2026, 7, 28, 12, 0, 0), "Fan", 480, 6, 0, 0, 80.00, "Dad", 0, "Ongoing"],
  [23, new Date(2026, 7, 28, 12, 0, 0), "Fan", 506, 6, 0, 0, 84.33, "Mum", 0, "Ongoing"]
]);

const tx6 = ss6.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
tx6.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
tx6.getRange(2, 1, 11, 5).setValues([
  [302, new Date(2026, 8, 30, 0, 0, 0), "Sep 30 Rollover", "Mido", 1320.00],
  [304, new Date(2026, 9, 2, 0, 0, 0), "Basilico", "Mido", 1447.80],
  [305, new Date(2026, 9, 3, 0, 0, 0), "Wave", "Mido", 570.00],
  [307, new Date(2026, 9, 3, 0, 0, 0), "Fawry", "Mido", 2316.00],
  [309, new Date(2026, 9, 6, 0, 0, 0), "Mosaab", "Mido", 80.00],
  [310, new Date(2026, 9, 8, 0, 0, 0), "Amazon", "Mido", 333.10],
  [320, new Date(2026, 9, 14, 0, 0, 0), "We Internet", "Mido", 367.00],
  [328, new Date(2026, 9, 23, 0, 0, 0), "Amazon", "Mido", 164.99],
  [330, new Date(2026, 9, 29, 0, 0, 0), "ElAbd", "Mido", 250.00],
  [332, new Date(2026, 9, 20, 0, 0, 0), "Via caffe", "Mido", 97.00],
  [399, new Date(2026, 9, 25, 0, 0, 0), "Orphan Snack", "Dad", 50.00]
]);

const recon6 = sandbox.reconcileActiveStatement(ss6, tz);
const ledger6 = sandbox.buildDebtLedger(ss6, tz, recon6);
const midoShare6 = ledger6.personShares.find(ps => ps.person === "Mido");
assert.strictEqual(Number(midoShare6.purchases.toFixed(2)), 6945.89, `Mido purchases must equal 6945.89, got ${midoShare6.purchases}`);
assert.strictEqual(Number(midoShare6.installments.toFixed(2)), 1484.61, `Mido installments must equal exact 1484.61, got ${midoShare6.installments}`);
assert.strictEqual(recon6.orphanLogs.length, 1, "Dad's unpaired snack must be in orphanLogs with 0 billed");
assert.strictEqual(Number(ledger6.personShares.find(ps => ps.person === "Dad").purchases.toFixed(2)), 0.00);

const cycleRange6 = sandbox.getCycleDateRange("2026-11", tz);
const midnightSep30 = new Date(2026, 8, 30, 0, 0, 0);
assert.ok(midnightSep30 >= cycleRange6.start && midnightSep30 <= cycleRange6.end, "Sep 30 00:00:00 midnight must fall inside Nov 2026 cycle date range!");

sandbox.updateLiveDashboard();
const mbSheet6 = ss6.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
assert.strictEqual(mbSheet6._getCell(1, 26), "=DATE(2026,9,30)", "Z1 must be midnight DATE(2026,9,30)");
assert.strictEqual(mbSheet6._getCell(2, 26), "=DATE(2026,10,30)", "Z2 must be midnight DATE(2026,10,30)");

const purchFormula6 = String(mbSheet6._getCell(10, 1) || "");
assert.ok(
  purchFormula6.startsWith("=VSTACK(") &&
  purchFormula6.includes("ARRAYFORMULA(") &&
  purchFormula6.includes("<($Z$2+1)") &&
  purchFormula6.includes("COUNTIFS(") &&
  !purchFormula6.includes('"🚫*"') &&
  purchFormula6.includes("Unrecorded Statement Items, Fees & Refunds") &&
  purchFormula6.includes("Active Bill & Reconciliation"),
  "A10 must exclude all Orphan Logs via COUNTIFS without restricting to 🚫* and VSTACK exceptions"
);

const instFormula6 = String(mbSheet6._getCell(10, 8) || "");
assert.ok(instFormula6.includes("MAP(") && instFormula6.includes("LAMBDA(") && instFormula6.includes("COUNTIFS("), "Installments spill must use MAP + LAMBDA + scalar COUNTIFS");

// ==========================================
// Test 7: Multi-Pass Cross-Gateway Matching, Multi-Row Uneven Split & No Duplication
// ==========================================
const ss7 = new MockSpreadsheet();
setActiveSpreadsheet(ss7);

const stmt7 = ss7.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
stmt7.getRange(2, 1, 3, 6).setValues([
  ["Statement Date:", new Date(2026, 9, 30, 12, 0, 0), "Due Date:", new Date(2026, 10, 25, 12, 0, 0), "Credit Limit:", 100000],
  ["Opening Balance:", 0, "Closing Balance:", 2375.90, "Total Debit:", 2375.90],
  ["Total Credit:", 0, "Card Number:", "1234", "Net:", 2375.90]
]);
stmt7.getRange(7, 1, 4, 7).setValues([
  [1, new Date(2026, 9, 2, 12, 0, 0), new Date(2026, 9, 2, 12, 0, 0), "FAWRY*BAZLYKWA ALEX EGY \"MOBILE TOKEN\"", "DEBIT", 1447.80, ""],
  [2, new Date(2026, 9, 4, 12, 0, 0), new Date(2026, 9, 4, 12, 0, 0), "SAHL GIZA EGY", "DEBIT", 80.00, ""],
  [3, new Date(2026, 9, 3, 12, 0, 0), new Date(2026, 9, 3, 12, 0, 0), "BASATA PAY MAADI EGY \"MOBILE TOKEN\"", "DEBIT", 15.00, ""],
  [4, new Date(2026, 9, 8, 12, 0, 0), new Date(2026, 9, 8, 12, 0, 0), "AMAZON", "DEBIT", 833.10, ""]
]);

const tx7 = ss7.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
tx7.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
tx7.getRange(2, 1, 7, 5).setValues([
  [1, new Date(2026, 9, 2, 0, 0, 0), "Basilico", "Mido", 1447.80],
  [2, new Date(2026, 9, 4, 0, 0, 0), "rabbit", "Muhanad", 80.00],
  [3, new Date(2026, 9, 6, 0, 0, 0), "rabbit", "Muhanad", 15.00],
  [4, new Date(2026, 9, 8, 0, 0, 0), "Amazon", "Mido", 333.10],
  [5, new Date(2026, 9, 8, 0, 0, 0), "Amazon", "Dad", 200.00],
  [6, new Date(2026, 9, 8, 0, 0, 0), "Amazon", "Mum", 200.00],
  [7, new Date(2026, 9, 8, 0, 0, 0), "Amazon", "Muhanad", 100.00]
]);

let recon7 = sandbox.reconcileActiveStatement(ss7, tz);
// High-confidence pairs auto-match (FAWRY*BAZLYKWA, SAHL, AMAZON split); unknown BASATA PAY <-> rabbit is Suggested only!
assert.strictEqual(recon7.stmtDebits.filter(st => st.matched).length, 3, "3 merchant-matched statement lines must auto-match");
assert.strictEqual(recon7.missingInSheet.length, 1, `Expected 1 suggested item in Needs Pairing (BASATA PAY), got ${recon7.missingInSheet.length}`);
assert.strictEqual(recon7.orphanLogs.length, 1, `Expected 1 unpaired log in Orphan Logs (rabbit 15.00), got ${recon7.orphanLogs.length}`);
assert.ok(recon7.missingInSheet[0].suggestedPairs && recon7.missingInSheet[0].suggestedPairs.length > 0, "BASATA PAY must surface a suggested pair");
assert.deepStrictEqual(Array.from(recon7.missingInSheet[0].suggestedPairs[0].txIds), ["T-00003"]);

let ledger7 = sandbox.buildDebtLedger(ss7, tz, recon7);
sandbox.renderActiveReconciliationSheet(ss7, tz, ledger7, recon7);
const ar7 = ss7.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
let rM1_7 = -1;
for (let r = 12; r <= ar7.getLastRow(); r++) {
  if (ar7._getCell(r, 1) === "M-1") { rM1_7 = r; break; }
}
assert.ok(rM1_7 > 0, "Should find M-1 (BASATA PAY) row in Needs Pairing");
ar7._setCell(rM1_7, 6, recon7.missingInSheet[0].suggestedPairs[0].dropdownLabel);
sandbox.handleActiveReconEdit({ range: ar7.getRange(rM1_7, 6) });

recon7 = sandbox.reconcileActiveStatement(ss7, tz);
assert.strictEqual(recon7.missingInSheet.length, 0, "Expected 0 missing in Needs Pairing after 1-click confirmation");
assert.strictEqual(recon7.orphanLogs.length, 0, "Expected 0 orphans after 1-click confirmation");
assert.ok(recon7.stmtDebits.every(st => st.matched), "All 4 statement debits must match!");

ledger7 = sandbox.buildDebtLedger(ss7, tz, recon7);
const shares7 = {};
ledger7.personShares.forEach(ps => { shares7[ps.person] = Number(ps.purchases.toFixed(2)); });
assert.strictEqual(shares7["Mido"], 1780.90, `Mido purchases mismatch: ${shares7["Mido"]}`);
assert.strictEqual(shares7["Dad"], 200.00, `Dad purchases mismatch: ${shares7["Dad"]}`);
assert.strictEqual(shares7["Mum"], 200.00, `Mum purchases mismatch: ${shares7["Mum"]}`);
assert.strictEqual(shares7["Muhanad"], 195.00, `Muhanad purchases mismatch: ${shares7["Muhanad"]}`);

// ==========================================
// Test 8: All Statement Credits Shown in Table 2 & Neglect/Use Persistence
// ==========================================
const ss8 = new MockSpreadsheet();
setActiveSpreadsheet(ss8);
const stmt8 = ss8.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
stmt8.getRange(2, 1, 3, 6).setValues([
  ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
  ["Opening Balance:", 1000, "Closing Balance:", 0, "Total Debit:", 0],
  ["Total Credit:", 1600, "Card Number:", "1234", "Net:", 0]
]);
stmt8.getRange(7, 1, 3, 7).setValues([
  [1, new Date(2026, 9, 5, 12, 0, 0), new Date(2026, 9, 5, 12, 0, 0), "INSTAPAY PAYMENT - THANK YOU", "CREDIT", 1000.00, ""],
  [2, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "FAWRY SETTLEMENT CREDIT", "CREDIT", 400.00, ""],
  [3, new Date(2026, 9, 12, 12, 0, 0), new Date(2026, 9, 12, 12, 0, 0), "MICROSOFT*MICROSOFT 365 REFUND", "CREDIT", 200.00, ""]
]);
const tx8 = ss8.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
tx8.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
tx8.getRange(2, 1, 1, 5).setValues([[1, new Date(2026, 9, 15, 0, 0, 0), "Test", "Mido", 1000.00]]);

let recon8 = sandbox.reconcileActiveStatement(ss8, tz);
let ledger8 = sandbox.buildDebtLedger(ss8, tz, recon8);
sandbox.renderActiveReconciliationSheet(ss8, tz, ledger8, recon8);

const ar8 = ss8.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
const cRows8 = [];
for (let r = 12; r <= ar8.getLastRow(); r++) {
  const id = String(ar8._getCell(r, 1) || "");
  if (/^C-\d+/.test(id)) {
    cRows8.push({ row: r, id, desc: ar8._getCell(r, 3), amt: ar8._getCell(r, 4), payer: ar8._getCell(r, 6) });
  }
}
assert.strictEqual(cRows8.length, 3, `All 3 credits must be shown in Table 2, got ${cRows8.length}`);
assert.strictEqual(cRows8[0].payer, "🚫 Neglect / Ignore", "Payment credit C-1 should default to Neglect");
assert.strictEqual(cRows8[1].payer, "🚫 Neglect / Ignore", "Settlement credit C-2 should default to Neglect");
assert.strictEqual(cRows8[2].payer, CONFIG.DEFAULT_FEE_PAYER, "Refund credit C-3 should default to DEFAULT_FEE_PAYER");
assert.strictEqual(Number(ledger8.refundsTotal.toFixed(2)), 200.00, "Only non-neglected credit (200) should count initially");

ar8._setCell(cRows8[1].row, 6, "Mido");
sandbox.handleActiveReconEdit({ range: ar8.getRange(cRows8[1].row, 6, 1, 1) });
ar8._setCell(cRows8[2].row, 6, "🚫 Neglect / Ignore");
sandbox.handleActiveReconEdit({ range: ar8.getRange(cRows8[2].row, 6, 1, 1) });

sandbox.syncActiveReconSheetToStore(ss8, tz);
recon8 = sandbox.reconcileActiveStatement(ss8, tz);
ledger8 = sandbox.buildDebtLedger(ss8, tz, recon8);
sandbox.renderActiveReconciliationSheet(ss8, tz, ledger8, recon8);

const midoShare8 = ledger8.personShares.find(ps => ps.person === "Mido");
const maiShare8 = ledger8.personShares.find(ps => ps.person === CONFIG.DEFAULT_FEE_PAYER);
assert.strictEqual(Number(midoShare8.refunds.toFixed(2)), 400.00, "Mido should receive C-2 (400.00) credit after refresh");
assert.strictEqual(Number(maiShare8.refunds.toFixed(2)), 0.00, "Mai should have 0.00 credit after neglecting C-3");
assert.strictEqual(Number(ledger8.refundsTotal.toFixed(2)), 400.00, "Total refunds should be 400.00");
assert.strictEqual(Number(ledger8.billPaymentsTotal.toFixed(2)), 1000.00, "Assigned C-2 (400) must NOT double-count in billPaymentsTotal");
assert.strictEqual(Number(ledger8.carriedOver.toFixed(2)), 0.00, "carriedOver must be Opening (1000) - Neglected payment (1000) = 0.00");

console.log("reconciliation.test.js: all 8 reconciliation assertions passed.");
