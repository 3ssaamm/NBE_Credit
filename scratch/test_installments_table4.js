const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

// ==========================================
// Minimal Google Apps Script Mock Environment
// ==========================================

function formatDateMock(date, tz, fmt) {
  const d = new Date(date);
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  const ss = String(d.getSeconds()).padStart(2, "0");
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  if (fmt === "yyyy-MM-dd") return `${y}-${String(m).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  if (fmt === "yyyy-MM-dd HH:mm:ss") return `${y}-${String(m).padStart(2, "0")}-${String(day).padStart(2, "0")} ${hh}:${mm}:${ss}`;
  if (fmt === "yyyy-MM") return `${y}-${String(m).padStart(2, "0")}`;
  if (fmt === "MMMM yyyy") return `${months[m - 1]} ${y}`;
  if (fmt === "MMMM d, yyyy") return `${months[m - 1]} ${day}, ${y}`;
  if (fmt === "yyyy") return String(y);
  if (fmt === "M") return String(m);
  if (fmt === "d") return String(day);
  return d.toISOString();
}

function colLetterToIndex(letters) {
  let res = 0;
  for (let i = 0; i < letters.length; i++) {
    res = res * 26 + (letters.charCodeAt(i) - 64);
  }
  return res;
}

function parseA1(a1) {
  const clean = a1.replace(/\$/g, "");
  if (clean.includes(":")) {
    const [p1, p2] = clean.split(":");
    const m1 = p1.match(/^([A-Z]+)(\d+)?$/i);
    const m2 = p2.match(/^([A-Z]+)(\d+)?$/i);
    if (m1 && m2) {
      const c1 = colLetterToIndex(m1[1].toUpperCase());
      const r1 = m1[2] ? parseInt(m1[2], 10) : 1;
      const c2 = colLetterToIndex(m2[1].toUpperCase());
      const r2 = m2[2] ? parseInt(m2[2], 10) : 200;
      return { r: r1, c: c1, nr: r2 - r1 + 1, nc: c2 - c1 + 1 };
    }
    const rOnly = clean.match(/^(\d+):(\d+)$/);
    if (rOnly) {
      const r1 = parseInt(rOnly[1], 10);
      const r2 = parseInt(rOnly[2], 10);
      return { r: r1, c: 1, nr: r2 - r1 + 1, nc: 26 };
    }
  }
  const m = clean.match(/^([A-Z]+)(\d+)$/i);
  return { r: parseInt(m[2], 10), c: colLetterToIndex(m[1].toUpperCase()), nr: 1, nc: 1 };
}

class MockRange {
  constructor(sheet, r, c, nr, nc) {
    this.sheet = sheet;
    this.r = r;
    this.c = c;
    this.nr = nr;
    this.nc = nc;
  }
  getSheet() { return this.sheet; }
  getRow() { return this.r; }
  getColumn() { return this.c; }
  getLastRow() { return this.r + this.nr - 1; }
  getNumRows() { return this.nr; }
  getNumColumns() { return this.nc; }
  getValue() { return this.sheet._getCell(this.r, this.c); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) {
        row.push(this.sheet._getCell(this.r + i, this.c + j));
      }
      out.push(row);
    }
    return out;
  }
  setValue(v) {
    this.sheet._setCell(this.r, this.c, v);
    return this;
  }
  setFormula(v) {
    return this.setValue(v);
  }
  setValues(arr) {
    for (let i = 0; i < arr.length; i++) {
      for (let j = 0; j < arr[i].length; j++) {
        this.sheet._setCell(this.r + i, this.c + j, arr[i][j]);
      }
    }
    return this;
  }
  clearContent() {
    for (let i = 0; i < this.nr; i++) {
      for (let j = 0; j < this.nc; j++) {
        this.sheet._setCell(this.r + i, this.c + j, "");
      }
    }
    return this;
  }
  setBackground(bg) {
    for (let i = 0; i < this.nr; i++) {
      for (let j = 0; j < this.nc; j++) {
        this.sheet.backgrounds[`${this.r + i}_${this.c + j}`] = bg;
      }
    }
    return this;
  }
  setDataValidation(dv) {
    this.sheet.validations[`${this.r}_${this.c}`] = dv;
    return this;
  }
  merge() { return this; }
  setFontWeight() { return this; }
  setFontSize() { return this; }
  setFontStyle() { return this; }
  setFontColor() { return this; }
  setHorizontalAlignment() { return this; }
  setBorder() { return this; }
  setNumberFormat() { return this; }
  clearFormat() { return this; }
  setWrap() { return this; }
}

class MockSheet {
  constructor(name, ss) {
    this.name = name;
    this.ss = ss;
    this.grid = {};
    this.backgrounds = {};
    this.validations = {};
    this.maxRow = 0;
    this.maxCol = 12;
    this.colWidths = {};
  }
  getName() { return this.name; }
  getParent() { return this.ss; }
  _getCell(r, c) {
    const k = `${r}_${c}`;
    return Object.prototype.hasOwnProperty.call(this.grid, k) ? this.grid[k] : "";
  }
  _setCell(r, c, v) {
    this.grid[`${r}_${c}`] = v;
    if (v !== "" && v !== null && v !== undefined && r > this.maxRow) {
      this.maxRow = r;
    }
    if (c > this.maxCol) this.maxCol = c;
  }
  getLastRow() {
    let max = 0;
    for (const k of Object.keys(this.grid)) {
      const val = this.grid[k];
      if (val !== "" && val !== null && val !== undefined) {
        const r = parseInt(k.split("_")[0], 10);
        if (r > max) max = r;
      }
    }
    return max;
  }
  getMaxColumns() { return this.maxCol; }
  getMaxRows() { return Math.max(this.getLastRow(), 50); }
  clear() {
    this.grid = {};
    this.backgrounds = {};
    this.validations = {};
    this.maxRow = 0;
  }
  hideSheet() {}
  setFrozenRows() {}
  setHiddenGridlines() {}
  setRowHeight() {}
  setColumnWidth(c, w) { this.colWidths[c] = w; }
  getColumnWidth(c) { return this.colWidths[c] || 100; }
  autoResizeColumns() {}
  hideColumns() {}
  getRange(rOrA1, c, nr, nc) {
    if (typeof rOrA1 === "string") {
      const p = parseA1(rOrA1);
      return new MockRange(this, p.r, p.c, p.nr, p.nc);
    }
    return new MockRange(this, rOrA1, c, nr || 1, nc || 1);
  }
}

class MockSpreadsheet {
  constructor() {
    this.sheets = {};
    this.toasts = [];
  }
  getSpreadsheetTimeZone() { return "Africa/Cairo"; }
  getSheetByName(name) { return this.sheets[name] || null; }
  insertSheet(name) {
    const s = new MockSheet(name, this);
    this.sheets[name] = s;
    return s;
  }
  toast(msg) { this.toasts.push(msg); }
}

const codeSrc = fs.readFileSync(path.join(__dirname, "..", "Code.js"), "utf8");
const sandbox = {
  console,
  Logger: { log: () => {} },
  Utilities: { formatDate: formatDateMock },
  SpreadsheetApp: {
    newDataValidation: () => {
      const dv = { values: null, allowInvalid: true };
      const builder = {
        requireValueInList: (vals) => { dv.values = vals; return builder; },
        requireDate: () => builder,
        setAllowInvalid: (b) => { dv.allowInvalid = b; return builder; },
        build: () => dv
      };
      return builder;
    }
  }
};
vm.createContext(sandbox);
vm.runInContext(codeSrc + "\nthis.__CONFIG = CONFIG;", sandbox);
const CONFIG = sandbox.__CONFIG;

// ==========================================
// Helper: Evaluate Table 1 SUMIFS Formula in JS
// ==========================================
function wildcardMatches(cellVal, pattern) {
  const s = String(cellVal === null || cellVal === undefined ? "" : cellVal);
  if (pattern.startsWith("<>")) {
    const inner = pattern.slice(2);
    return !wildcardMatches(s, inner);
  }
  const escaped = pattern.replace(/[.+?^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*");
  const re = new RegExp(`^${escaped}$`, "i");
  return re.test(s);
}

function evalTable4Sumifs(sheet, personName, startInstRow, endInstRow) {
  let sum = 0;
  for (let r = startInstRow; r <= endInstRow; r++) {
    const emi = parseFloat(sheet._getCell(r, 4)) || 0;
    const action = String(sheet._getCell(r, 6) || "");
    if (wildcardMatches(action, `*${personName.trim()}*`) && wildcardMatches(action, "<>🚫*")) {
      sum += emi;
    }
  }
  return Number(sum.toFixed(2));
}

// ==========================================
// Build Realistic Spreadsheet Fixtures
// ==========================================
const ss = new MockSpreadsheet();
const tz = "Africa/Cairo";

// 1. Bank Statement sheet
const stmtSheet = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
stmtSheet.getRange(2, 1, 3, 6).setValues([
  ["Statement Date:", new Date("2026-09-30T12:00:00"), "Due Date:", new Date("2026-10-25T12:00:00"), "Credit Limit:", 100000],
  ["Opening Balance:", 0, "Closing Balance:", 5264.61, "Total Debit:", 5264.61],
  ["Total Credit:", 0, "Card Number:", "1234", "Net:", 5264.61]
]);

const stmtRows = [
  // Standard 1-to-1 matched
  [1, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "HARD DRIVE 10 OF 12", "INSTALLMENT", 154.39, ""],
  // Unequal multi-person split (SPLIT_SAME_DESC): Dad 400 + Mum 405 + Mido 360 + Mai 308 = 1473.00
  [2, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "DEEPFREEZER 10 OF 12", "INSTALLMENT", 1473.00, ""],
  // Unequal 3-person split (SPLIT_SAME_DESC): Mai 33 + Mum 15 + Mido 45.67 = 93.67
  [3, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "HAND BLENDER 10 OF 12", "INSTALLMENT", 93.67, ""],
  // Equal 3-person split (SPLIT_SAME_DESC): Mido 74.07 + Mai 74.07 + Mum 74.07 = 222.21
  [4, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "MUHANAD'S EARPHONES 5 OF 9", "INSTALLMENT", 222.21, ""],
  // Statement installment with slight 0.05 rounding diff vs sheet EMI (sheet = 833.20, stmt = 833.25)
  [5, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "MOBILATY 10 OF 11", "INSTALLMENT", 833.25, ""],
  // Unmatched statement installment (missing in Installments sheet)
  [6, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "MYSTERY GADGET 1 OF 6", "INSTALLMENT", 350.00, ""]
];
stmtSheet.getRange(7, 1, stmtRows.length, 7).setValues(stmtRows);

// 2. Installments sheet
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
  // Ongoing sheet item NOT on statement (should appear in Sheet Only, skipped by default)
  [13, new Date("2026-09-01T12:00:00"), "Cosmatics", 2835.84, 6, 0, 0, 472.64, "Mai", 0, "Ongoing"],
  // Completed sheet item (should NOT appear in Sheet Only)
  [14, new Date("2025-12-01T12:00:00"), "Noon 2 of 6", 264.00, 2, 0, 0, 132.00, "Mido", 2, "Completed"]
];
instSheet.getRange(2, 1, instRows.length, 11).setValues(instRows);

// ==========================================
// Test 1: Reconciliation, Discrepancy Ordering & Unequal Splits
// ==========================================
console.log("Running Test 1: Reconciliation, Discrepancy Ordering & Unequal Splits...");
const recon1 = sandbox.reconcileActiveStatement(ss, tz);
const items1 = recon1.installmentReconItems;

// Expected count:
// 1 Statement Only (MYSTERY GADGET)
// 1 Sheet Only (Cosmatics) -- Completed Noon 2 of 6 must be excluded
// 12 Matched entries (1 Hard drive + 4 Deepfreezer + 3 Hand blender + 3 Earphones + 1 Mobilaty)
assert.strictEqual(items1.length, 14, `Expected 14 Table 4 items, got ${items1.length}`);
assert.strictEqual(items1[0].sourceCategory, "⚠️ Statement Only", "Discrepancy Statement Only must be first");
assert.strictEqual(items1[0].action, "", "Statement Only default action must be blank");
assert.strictEqual(items1[1].sourceCategory, "⏳ Sheet Only", "Sheet Only must appear before matched items");
assert.strictEqual(items1[1].action, "🚫 Not on Statement (Skip)", "Sheet Only default action must skip");
assert.ok(items1[1].desc.includes("Cosmatics"), "Cosmatics must be in Sheet Only");
assert.ok(!items1.some(it => it.desc.includes("Noon 2 of 6")), "Completed item must NOT be in Sheet Only");

// Check Mobilaty 0.05 difference note and adjusted amount
const mobilatyItem = items1.find(it => it.desc.includes("Mobilaty"));
assert.ok(mobilatyItem, "Mobilaty must be present");
assert.strictEqual(mobilatyItem.amount, 833.25, "Mobilaty amount should adjust to statement 833.25");
assert.ok(mobilatyItem.note.includes("⚠️ Stmt: 833.25 vs Sheet: 833.20"), "Mobilaty note should surface Stmt vs Sheet diff");

// Check exact unequal per-person shares in buildDebtLedger!
const ledger1 = sandbox.buildDebtLedger(ss, tz, recon1);
const shareMap1 = {};
ledger1.personShares.forEach(ps => { shareMap1[ps.person] = Number(ps.installments.toFixed(2)); });

// Expected:
// Dad: Deepfreezer 400.00
// Mum: Deepfreezer 405.00 + Hand blender 15.00 + Earphones 74.07 = 494.07
// Mido: Hard drive 154.39 + Deepfreezer 360.00 + Hand blender 45.67 + Earphones 74.07 = 634.13
// Mai: Deepfreezer 308.00 + Hand blender 33.00 + Earphones 74.07 + Mobilaty 833.25 = 1248.32
assert.strictEqual(shareMap1["Dad"], 400.00, `Dad share mismatch: ${shareMap1["Dad"]}`);
assert.strictEqual(shareMap1["Mum"], 494.07, `Mum share mismatch: ${shareMap1["Mum"]}`);
assert.strictEqual(shareMap1["Mido"], 634.13, `Mido share mismatch: ${shareMap1["Mido"]}`);
assert.strictEqual(shareMap1["Mai"], 1248.32, `Mai share mismatch: ${shareMap1["Mai"]}`);
console.log("  PASS: Exact unequal split EMI attribution & discrepancy ordering verified!");

// ==========================================
// Test 2: Rendering Table 4 & Live SUMIFS Equivalence
// ==========================================
console.log("Running Test 2: Rendering Table 4 & Live SUMIFS Formula Equivalence...");
sandbox.renderActiveReconciliationSheet(ss, tz, ledger1, recon1);
const reconSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);

// Find where Table 4 starts in reconSheet
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

// Verify Table 1 Col D contains live SUMIFS formulas pointing to startInstRow..endInstRow
for (let i = 0; i < ledger1.personShares.length; i++) {
  const r = 12 + i;
  const person = reconSheet._getCell(r, 2);
  const formula = String(reconSheet._getCell(r, 4) || "");
  assert.ok(formula.startsWith("=SUMIFS("), `Table 1 Col D row ${r} must be SUMIFS formula, got: ${formula}`);
  assert.ok(formula.includes(`$D$${startInstRow}:$D$${endInstRow}`), "SUMIFS must reference Table 4 Col D");
  const evalSum = evalTable4Sumifs(reconSheet, person, startInstRow, endInstRow);
  assert.strictEqual(evalSum, shareMap1[person], `SUMIFS eval for ${person} (${evalSum}) must match ledger (${shareMap1[person]})`);
}
console.log("  PASS: Table 1 SUMIFS formula matches exact per-person EMI totals!");

// ==========================================
// Test 3: onEdit Handling & Occurrence-Safe Persistence
// ==========================================
console.log("Running Test 3: onEdit Handling & Occurrence-Safe Persistence...");

// 3a. Assign bare name "Abdo" to INST-1 (Statement Only MYSTERY GADGET 350.00)
const rowInst1 = startInstRow; // INST-1
reconSheet._setCell(rowInst1, 6, "Abdo");
sandbox.handleActiveReconEdit({ range: reconSheet.getRange(rowInst1, 6) });
assert.strictEqual(reconSheet._getCell(rowInst1, 6), "✅ Billed to Abdo", "Bare name should normalize to '✅ Billed to Abdo'");

// 3b. Toggle 2nd occurrence of Muhanad's Earphones (Mai's 74.07) to Neglect
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

// 3c. Toggle Cosmatics (Sheet Only) to "✅ Billed to Mai"
const rowCosmatics = startInstRow + 1;
reconSheet._setCell(rowCosmatics, 6, "✅ Billed to Mai");
sandbox.handleActiveReconEdit({ range: reconSheet.getRange(rowCosmatics, 6) });

// Refresh / Reload from scratch to verify _Recon_Store persistence!
const recon2 = sandbox.reconcileActiveStatement(ss, tz);
const ledger2 = sandbox.buildDebtLedger(ss, tz, recon2);
const shareMap2 = {};
ledger2.personShares.forEach(ps => { shareMap2[ps.person] = Number(ps.installments.toFixed(2)); });

// Verify Mido's Earphones (occ=1) was NOT clobbered by Mai's Earphones (occ=2) override!
const earphonesItems2 = recon2.installmentReconItems.filter(it => it.desc.includes("Muhanad's Earphones"));
assert.strictEqual(earphonesItems2.length, 3);
assert.strictEqual(earphonesItems2[0].action, "✅ Billed to Mido", "occ=1 Mido must remain billed!");
assert.strictEqual(earphonesItems2[1].action, "🚫 Neglect / Skip This Cycle", "occ=2 Mai must persist as neglected!");
assert.strictEqual(earphonesItems2[2].action, "✅ Billed to Mum", "occ=3 Mum must remain billed!");

// Verify Abdo now has 350.00, Mai has 1248.32 - 74.07 + 472.64 = 1646.89
assert.strictEqual(shareMap2["Abdo"], 350.00, `Abdo share after assign mismatch: ${shareMap2["Abdo"]}`);
assert.strictEqual(shareMap2["Mai"], 1646.89, `Mai share after neglect+unskip mismatch: ${shareMap2["Mai"]}`);
assert.strictEqual(shareMap2["Mido"], 634.13, `Mido share should be unchanged: ${shareMap2["Mido"]}`);
assert.strictEqual(shareMap2["Mum"], 494.07, `Mum share should be unchanged: ${shareMap2["Mum"]}`);

// Re-render and verify SUMIFS matches ledger2
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
console.log("  PASS: Persistence across refresh & occurrence-safe overrides verified!");

// ==========================================
// Test 4: Co-existence with Table 3, Rounding Residual, Clearing Override & syncInstallments
// ==========================================
console.log("Running Test 4: Co-existence with Table 3, Rounding Residual & Clearing Override...");

// Add an unmatched transaction so Table 3 renders above Table 4
const txSheet = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
txSheet.getRange(2, 1, 1, 5).setValues([
  [1, new Date("2026-09-15T12:00:00"), "Unmatched Coffee", "Mido", 120.00]
]);

// Add a 100.00 3-way equal split installment (33.3333 each)
stmtSheet.getRange(13, 1, 1, 7).setValues([
  [7, new Date("2026-09-30T12:00:00"), new Date("2026-09-30T12:00:00"), "SHARED GYM 2 OF 3", "INSTALLMENT", 100.00, ""]
]);
instSheet.getRange(16, 1, 1, 11).setValues([
  [15, new Date("2026-07-15T12:00:00"), "Shared Gym", 300.00, 3, 0, 0, 100.00, "Mido + Mai + Abdo", 1, "Ongoing"]
]);

// Clear INST-1 override via handleActiveReconEdit("") -> should revert to blank default on next reconcile
let inst1RowAfter = startInstRow2;
reconSheet._setCell(inst1RowAfter, 6, "");
sandbox.handleActiveReconEdit({ range: reconSheet.getRange(inst1RowAfter, 6) });

const recon4 = sandbox.reconcileActiveStatement(ss, tz);
const ledger4 = sandbox.buildDebtLedger(ss, tz, recon4);
sandbox.renderActiveReconciliationSheet(ss, tz, ledger4, recon4);

// Verify Table 3 AND Table 4 coexist cleanly without row collisions
let txRowsCount = 0;
let instRowsCount = 0;
let startInst4 = -1, endInst4 = -1;
for (let r = 1; r <= reconSheet.getLastRow(); r++) {
  const id = String(reconSheet._getCell(r, 1) || "");
  if (/^TX-/i.test(id)) txRowsCount++;
  if (/^INST-\d+/i.test(id)) {
    if (startInst4 === -1) startInst4 = r;
    endInst4 = r;
    instRowsCount++;
  }
}
assert.strictEqual(txRowsCount, 1, "Table 3 must have 1 row");
assert.strictEqual(instRowsCount, 17, `Table 4 must have 17 rows (14 + 3 Shared Gym), got ${instRowsCount}`);

// Verify Shared Gym 3-way split sums to 100.00 exactly and has NO false ⚠️ diff warning
const gymItems = recon4.installmentReconItems.filter(it => it.desc.includes("Shared Gym"));
assert.strictEqual(gymItems.length, 3);
const gymSum = Number(gymItems.reduce((s, it) => s + it.amount, 0).toFixed(2));
assert.strictEqual(gymSum, 100.00, `Shared Gym 3-way split must sum to 100.00, got ${gymSum}`);
assert.ok(gymItems.every(it => !it.note.includes("⚠️")), "Exact 100/3 rounding should not emit false ⚠️ diff warning");

// Verify INST-1 reverted to blank default after clearing
assert.strictEqual(recon4.installmentReconItems[0].action, "", "Cleared INST-1 override should revert to blank");

// Verify Table 1 SUMIFS matches ledger4 with Table 3 present
for (let i = 0; i < ledger4.personShares.length; i++) {
  const r = 12 + i;
  const person = reconSheet._getCell(r, 2);
  const expected = Number(ledger4.personShares[i].installments.toFixed(2));
  const evalSum = evalTable4Sumifs(reconSheet, person, startInst4, endInst4);
  assert.strictEqual(evalSum, expected, `Table 3+4 SUMIFS mismatch for ${person}: ${evalSum} != ${expected}`);
}

// Verify syncInstallmentsFromStatement executes cleanly with updated matchStatementInstallments
sandbox.syncInstallmentsFromStatement(ss, tz, false);
assert.strictEqual(instSheet._getCell(2, 10), 9, "Hard drive (10 of 12, unpaid) should sync to 9 payments made");

console.log("  PASS: Table 3+4 coexistence, 3-way rounding, clearing override, and syncInstallments verified!");

// ==========================================
// Test 5: Multi-Date NUTOPIA Persistence, Legacy _#k Recovery & Same-Date Isolation
// ==========================================
console.log("Running Test 5: Multi-Date NUTOPIA Persistence, Legacy _#k Recovery & Same-Date Isolation...");

const ss5 = new MockSpreadsheet();
sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ss5;

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

// Pre-seed _Recon_Store with legacy _#2/_#3 entries caused by date-agnostic occ counting,
// plus a same-date occ=2 entry for Google One (leaving occ=1 unassigned).
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
// FAWRY*THE H CAF (2316.00) must match Fawry (2316.00) via Pass 1 and NOT appear in Table 2
assert.ok(!t2AfterRecovery.some(r => Math.abs(r.amount - 2316.00) < 0.01), "FAWRY*THE H CAF must match Fawry via Pass 1");

const recoveredNutopia = t2AfterRecovery.filter(r => r.desc.includes("NUTOPIA"));
assert.strictEqual(recoveredNutopia.length, 7, "All 7 NUTOPIA charges must appear in Table 2");
// The 4 legacy _#2/_#3 dates (Oct 24, Oct 21, Oct 12, Oct 20) must be recovered automatically as Mai
["2026-10-24", "2026-10-21", "2026-10-12", "2026-10-20"].forEach(dStr => {
  const match = recoveredNutopia.find(r => sandbox.makeNormDateStr(r.date, tz) === dStr);
  assert.ok(match, `Should find NUTOPIA row for ${dStr}`);
  assert.strictEqual(match.payer, "Mai", `Legacy _#k NUTOPIA row on ${dStr} should auto-recover to Mai`);
});

// Same-date Google One (2026-10-23): occ=1 must stay blank, occ=2 must stay Muhanad
const googleRows5 = t2AfterRecovery.filter(r => r.desc.includes("Google One"));
assert.strictEqual(googleRows5.length, 2);
assert.strictEqual(googleRows5[0].payer, "", "Google One occ=1 must NOT be clobbered by occ=2");
assert.strictEqual(googleRows5[1].payer, "Muhanad", "Google One occ=2 must remain Muhanad");

// Now assign Mai to the remaining 3 NUTOPIA rows (Oct 28 40.00, Oct 28 80.00, Oct 22 85.00),
// assign Mido to the 2 CREDIT rows and Google One occ=1, and verify multi-row range edit + refresh!
const nutopiaOct28_40 = recoveredNutopia.find(r => sandbox.makeNormDateStr(r.date, tz) === "2026-10-28" && r.amount === 40);
const nutopiaOct28_80 = recoveredNutopia.find(r => sandbox.makeNormDateStr(r.date, tz) === "2026-10-28" && r.amount === 80);
const nutopiaOct22_85 = recoveredNutopia.find(r => sandbox.makeNormDateStr(r.date, tz) === "2026-10-22" && r.amount === 85);
reconSheet5._setCell(nutopiaOct28_40.row, 6, "Mai");
reconSheet5._setCell(nutopiaOct28_80.row, 6, "Mai");
reconSheet5._setCell(nutopiaOct22_85.row, 6, "Mai");
sandbox.handleActiveReconEdit({ range: reconSheet5.getRange(nutopiaOct28_40.row, 6) });
sandbox.handleActiveReconEdit({ range: reconSheet5.getRange(nutopiaOct28_80.row, 6) });
sandbox.handleActiveReconEdit({ range: reconSheet5.getRange(nutopiaOct22_85.row, 6) });

// Multi-row paste edit on the two Credit rows (C-1, C-2 -> Mido)
const creditRows5 = t2AfterRecovery.filter(r => r.category === "Refund / Credit");
assert.strictEqual(creditRows5.length, 2);
reconSheet5._setCell(creditRows5[0].row, 6, "Mido");
reconSheet5._setCell(creditRows5[1].row, 6, "Mido");
sandbox.handleActiveReconEdit({ range: reconSheet5.getRange(creditRows5[0].row, 6, 2, 1) });

// Assign Google One occ=1 to Mido (leaving occ=2 as Muhanad)
reconSheet5._setCell(googleRows5[0].row, 6, "Mido");
sandbox.handleActiveReconEdit({ range: reconSheet5.getRange(googleRows5[0].row, 6) });

// Refresh dashboard and verify NOTHING is cleared!
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
console.log("  PASS: Multi-date NUTOPIA persistence, legacy _#k recovery, and same-date isolation verified!");

// ==========================================
// Test 6: Member Breakdown Midnight Date Boundary, Installment Parity (1,484.61) & Section 3
// ==========================================
console.log("Running Test 6: Member Breakdown Midnight Date Boundary, Installment Parity & Section 3...");

const ss6 = new MockSpreadsheet();
sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ss6;

const stmt6 = ss6.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
stmt6.getRange(2, 1, 3, 6).setValues([
  ["Statement Date:", new Date(2026, 9, 30, 12, 0, 0), "Due Date:", new Date(2026, 10, 25, 12, 0, 0), "Credit Limit:", 100000],
  ["Opening Balance:", 0, "Closing Balance:", 9000.00, "Total Debit:", 9000.00],
  ["Total Credit:", 0, "Card Number:", "1234", "Net:", 9000.00]
]);
stmt6.getRange(7, 1, 13, 7).setValues([
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
  // Intentional 0.02 statement rounding diff on multi-row Fan (Dad 80 + Mum 84.33 + Mido 84.50 = 248.83 vs stmt 248.81)
  [12, new Date(2026, 9, 30, 12, 0, 0), new Date(2026, 9, 30, 12, 0, 0), "FAN 1 OF 6", "INSTALLMENT", 248.81, ""],
  [13, new Date(2026, 9, 23, 12, 0, 0), new Date(2026, 9, 23, 12, 0, 0), "Google One 650-2530000 USACA", "DEBIT", 179.99, ""]
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

// Transactions including 2026-09-30 stored at 00:00:00 midnight (1,320.00 EGP) + Oct purchases (5,625.89 EGP) = 6,945.89 EGP
const tx6 = ss6.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
tx6.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
tx6.getRange(2, 1, 10, 5).setValues([
  [302, new Date(2026, 8, 30, 0, 0, 0), "Sep 30 Rollover", "Mido", 1320.00],
  [304, new Date(2026, 9, 2, 0, 0, 0), "Basilico", "Mido", 1447.80],
  [305, new Date(2026, 9, 3, 0, 0, 0), "Wave", "Mido", 570.00],
  [307, new Date(2026, 9, 3, 0, 0, 0), "Fawry", "Mido", 2316.00],
  [309, new Date(2026, 9, 6, 0, 0, 0), "Mosaab", "Mido", 80.00],
  [310, new Date(2026, 9, 8, 0, 0, 0), "Amazon", "Mido", 333.10],
  [320, new Date(2026, 9, 14, 0, 0, 0), "We Internet", "Mido", 367.00],
  [328, new Date(2026, 9, 23, 0, 0, 0), "Amazon", "Mido", 164.99],
  [330, new Date(2026, 9, 29, 0, 0, 0), "ElAbd", "Mido", 250.00],
  [332, new Date(2026, 9, 20, 0, 0, 0), "Via caffe", "Mido", 97.00]
]);

const recon6 = sandbox.reconcileActiveStatement(ss6, tz);
const ledger6 = sandbox.buildDebtLedger(ss6, tz, recon6);
const midoShare6 = ledger6.personShares.find(ps => ps.person === "Mido");
assert.strictEqual(Number(midoShare6.purchases.toFixed(2)), 6945.89, `Mido purchases must equal 6945.89, got ${midoShare6.purchases}`);
assert.strictEqual(Number(midoShare6.installments.toFixed(2)), 1484.61, `Mido installments must equal exact 1484.61 without multi-row residual subtraction, got ${midoShare6.installments}`);

const cycleRange6 = sandbox.getCycleDateRange("2026-11", tz);
const midnightSep30 = new Date(2026, 8, 30, 0, 0, 0);
assert.ok(midnightSep30 >= cycleRange6.start && midnightSep30 <= cycleRange6.end, "Sep 30 00:00:00 midnight must fall inside Nov 2026 cycle date range!");

sandbox.updateLiveDashboard();
const mbSheet6 = ss6.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
assert.strictEqual(mbSheet6._getCell(1, 26), "=DATE(2026,9,30)", "Z1 must be midnight DATE(2026,9,30)");
assert.strictEqual(mbSheet6._getCell(2, 26), "=DATE(2026,10,30)", "Z2 must be midnight DATE(2026,10,30)");

const purchFormula6 = String(mbSheet6._getCell(10, 1) || "");
assert.ok(purchFormula6.includes("ARRAYFORMULA(") && purchFormula6.includes("<($Z$2+1)"), "Purchases spill must use ARRAYFORMULA and <($Z$2+1)");

const instFormula6 = String(mbSheet6._getCell(10, 8) || "");
assert.ok(instFormula6.includes("MAP(") && instFormula6.includes("LAMBDA(") && instFormula6.includes("COUNTIFS("), "Installments spill must use MAP + LAMBDA + scalar COUNTIFS");

// Verify Section 3 header and spill formula exist in Col H below Installments
let sec3BannerRow = -1;
for (let r = 12; r <= mbSheet6.getLastRow(); r++) {
  const val = String(mbSheet6._getCell(r, 8) || "");
  if (val.includes("Assigned Statement Items")) {
    sec3BannerRow = r;
    break;
  }
}
assert.ok(sec3BannerRow >= 27, `Section 3 banner should be placed safely below Installments, found at row ${sec3BannerRow}`);
assert.strictEqual(mbSheet6._getCell(sec3BannerRow + 1, 13), "Amount (EGP)", "Section 3 Col M must align Amount (EGP)");
const sec3Formula = String(mbSheet6._getCell(sec3BannerRow + 2, 8) || "");
assert.ok(sec3Formula.includes("FILTER(ARRAYFORMULA(") && sec3Formula.includes("Active Bill & Reconciliation"), "Section 3 spill must filter Table 2 from Active Bill & Reconciliation");

console.log("  PASS: Member Breakdown midnight date boundary, installment parity (1,484.61), and Section 3 verified!");
console.log("ALL TABLE 4 INSTALLMENT TESTS PASSED!");
