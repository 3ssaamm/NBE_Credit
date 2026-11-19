const fs = require("fs");
const path = require("path");
const vm = require("vm");

const ROOT_DIR = path.resolve(__dirname, "..", "..");

function formatDateMock(date, tz, fmt) {
  const d = date instanceof Date ? date : new Date(date);
  if (isNaN(d.getTime())) return "";
  const timeZone = tz || "Africa/Cairo";
  const parts = {};
  new Intl.DateTimeFormat("en-US", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  }).formatToParts(d).forEach(p => {
    if (p.type !== "literal") parts[p.type] = p.value;
  });

  const y = parts.year;
  const mm2 = parts.month;
  const dd2 = parts.day;
  const m = String(parseInt(mm2, 10));
  const day = String(parseInt(dd2, 10));
  const hh = parts.hour === "24" ? "00" : parts.hour;
  const min = parts.minute;
  const ss = parts.second;
  const months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  const monthName = months[parseInt(mm2, 10) - 1];

  if (fmt === "yyyy-MM-dd") return `${y}-${mm2}-${dd2}`;
  if (fmt === "yyyy-MM-dd HH:mm:ss") return `${y}-${mm2}-${dd2} ${hh}:${min}:${ss}`;
  if (fmt === "yyyy-MM") return `${y}-${mm2}`;
  if (fmt === "MMMM yyyy") return `${monthName} ${y}`;
  if (fmt === "MMMM d, yyyy") return `${monthName} ${day}, ${y}`;
  if (fmt === "yyyy") return y;
  if (fmt === "M") return m;
  if (fmt === "d") return day;
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
  clear() { return this.clearContent(); }
  setBackground(bg) {
    for (let i = 0; i < this.nr; i++) {
      for (let j = 0; j < this.nc; j++) {
        this.sheet.backgrounds[`${this.r + i}_${this.c + j}`] = bg;
      }
    }
    return this;
  }
  setDataValidation(dv) {
    for (let i = 0; i < this.nr; i++) {
      for (let j = 0; j < this.nc; j++) {
        this.sheet.validations[`${this.r + i}_${this.c + j}`] = dv;
      }
    }
    return this;
  }
  setRanges() { return this; }
  merge() { return this; }
  setFontWeight() { return this; }
  setFontSize() { return this; }
  setFontStyle() { return this; }
  setFontColor() { return this; }
  setHorizontalAlignment() { return this; }
  setVerticalAlignment() { return this; }
  setBorder() { return this; }
  setNumberFormat() { return this; }
  clearFormat() { return this; }
  setWrap() { return this; }
  setWrapStrategy() { return this; }
}

class MockSheet {
  constructor(name, ss) {
    this.name = name;
    this.ss = ss;
    this.grid = {};
    this.backgrounds = {};
    this.validations = {};
    this.cfRules = [];
    this.maxRow = 0;
    this.maxCol = 12;
    this.colWidths = {};
    this.hidden = false;
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
  getLastColumn() {
    let max = 0;
    for (const k of Object.keys(this.grid)) {
      const val = this.grid[k];
      if (val !== "" && val !== null && val !== undefined) {
        const c = parseInt(k.split("_")[1], 10);
        if (c > max) max = c;
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
  hideSheet() { this.hidden = true; }
  showSheet() { this.hidden = false; }
  setFrozenRows() {}
  setHiddenGridlines() {}
  setRowHeight() {}
  setRowHeights() {}
  setColumnWidth(c, w) { this.colWidths[c] = w; }
  setColumnWidths() {}
  getColumnWidth(c) { return this.colWidths[c] || 100; }
  autoResizeColumns() {}
  hideColumns() {}
  showColumns() {}
  setConditionalFormatRules(rules) { this.cfRules = rules || []; }
  getConditionalFormatRules() { return this.cfRules; }
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
    this.activeSheet = null;
  }
  getSpreadsheetTimeZone() { return "Africa/Cairo"; }
  getSheetByName(name) { return this.sheets[name] || null; }
  insertSheet(name) {
    const s = new MockSheet(name, this);
    this.sheets[name] = s;
    return s;
  }
  setActiveSheet(s) { this.activeSheet = s; }
  toast(msg) { this.toasts.push(msg); }
}

function wildcardMatches(cellVal, pattern) {
  const s = String(cellVal === null || cellVal === undefined ? "" : cellVal);
  if (pattern.startsWith("<>")) {
    return !wildcardMatches(s, pattern.slice(2));
  }
  const escaped = pattern.replace(/[.+?^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*");
  return new RegExp(`^${escaped}$`, "i").test(s);
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

function makeLocalNoonDate(isoStr) {
  if (!isoStr || typeof isoStr !== "string") return "";
  const m = isoStr.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return isoStr;
  return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10), 12, 0, 0);
}

function createGasSandbox() {
  let currentSs = null;
  const sandbox = {
    console,
    Logger: { log: () => {} },
    Utilities: { formatDate: formatDateMock },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => currentSs,
      newDataValidation: () => {
        const dv = { values: null, allowInvalid: true };
        const builder = {
          requireValueInList: (vals) => { dv.values = vals; return builder; },
          requireDate: () => builder,
          setAllowInvalid: (b) => { dv.allowInvalid = b; return builder; },
          build: () => dv
        };
        return builder;
      },
      newConditionalFormatRule: () => {
        const rule = {};
        const builder = {
          whenFormulaSatisfied: (f) => { rule.formula = f; return builder; },
          setBold: (b) => { rule.bold = b; return builder; },
          setBackground: (bg) => { rule.bg = bg; return builder; },
          setRanges: (r) => { rule.ranges = r; return builder; },
          build: () => rule
        };
        return builder;
      }
    }
  };
  vm.createContext(sandbox);

  const srcDir = path.join(ROOT_DIR, "src");
  const filesToLoad = [];
  if (fs.existsSync(srcDir)) {
    fs.readdirSync(srcDir)
      .filter(f => /\.(js|gs)$/i.test(f))
      .sort()
      .forEach(f => filesToLoad.push(path.join(srcDir, f)));
  }
  const mainCodePath = path.join(ROOT_DIR, "Code.js");
  if (fs.existsSync(mainCodePath)) {
    filesToLoad.push(mainCodePath);
  }

  for (const filePath of filesToLoad) {
    const src = fs.readFileSync(filePath, "utf8");
    vm.runInContext(src, sandbox, { filename: path.basename(filePath) });
  }
  vm.runInContext("this.__CONFIG = CONFIG;", sandbox);

  return {
    sandbox,
    CONFIG: sandbox.__CONFIG,
    MockSpreadsheet,
    evalTable4Sumifs,
    setActiveSpreadsheet: (ss) => {
      currentSs = ss;
      sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ss;
    }
  };
}

function populateSpreadsheetFromFixture(ss, fixture, CONFIG) {
  if (fixture.defaultFeePayer && CONFIG) {
    CONFIG.DEFAULT_FEE_PAYER = fixture.defaultFeePayer;
  }

  if (fixture.bankStatement) {
    const stmtSheet = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
    const meta = fixture.bankStatement.meta || {};
    stmtSheet.getRange(2, 1, 3, 6).setValues([
      [
        "Statement Date:", makeLocalNoonDate(meta.statementDate),
        "Due Date:", makeLocalNoonDate(meta.dueDate),
        "Credit Limit:", meta.creditLimit || CONFIG.CREDIT_LIMIT
      ],
      [
        "Opening Balance:", meta.openingBalance || 0,
        "Closing Balance:", meta.closingBalance || 0,
        "Total Debit:", meta.totalDebit || 0
      ],
      [
        "Total Credit:", meta.totalCredit || 0,
        "Card Number:", meta.card || "****0000",
        "Net:", meta.closingBalance || 0
      ]
    ]);
    const rows = (fixture.bankStatement.rows || []).map(r => [
      r[0],
      makeLocalNoonDate(r[1]),
      makeLocalNoonDate(r[2] || r[1]),
      r[3],
      r[4],
      r[5],
      r[6] || ""
    ]);
    if (rows.length > 0) {
      stmtSheet.getRange(7, 1, rows.length, 7).setValues(rows);
    }
  }

  if (fixture.transactions && fixture.transactions.length > 0) {
    const txSheet = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
    txSheet.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
    const rows = fixture.transactions.map(r => [r[0], makeLocalNoonDate(r[1]), r[2], r[3], r[4]]);
    txSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  if (fixture.installments && fixture.installments.length > 0) {
    const instSheet = ss.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
    instSheet.getRange(1, 1, 1, 11).setValues([
      ["#", "Start Date", "Description", "Total Price", "Months", "C6", "C7", "EMI", "Payer", "Paid", "Status"]
    ]);
    const rows = fixture.installments.map(r => [
      r[0], makeLocalNoonDate(r[1]), r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10]
    ]);
    instSheet.getRange(2, 1, rows.length, 11).setValues(rows);
  }

  if (fixture.reconStore && fixture.reconStore.length > 0) {
    const storeSheet = ss.insertSheet(CONFIG.SHEETS.RECON_STORE);
    storeSheet.getRange(1, 1, 1, 10).setValues([
      ["Type", "Primary_Key", "Date_Str", "Desc", "Amount", "Payer", "Note", "Confirmed", "Neglected", "Updated_At"]
    ]);
    storeSheet.getRange(2, 1, fixture.reconStore.length, 10).setValues(fixture.reconStore);
  }

  if (fixture.paymentDate) {
    const arSheet = ss.insertSheet(CONFIG.SHEETS.ACTIVE_RECON);
    arSheet.getRange(6, 8).setValue(fixture.paymentDate);
  }
}

module.exports = {
  createGasSandbox,
  populateSpreadsheetFromFixture,
  MockSpreadsheet,
  evalTable4Sumifs,
  formatDateMock
};
