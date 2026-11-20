/**
 * NBE Credit Card Automated Tracker & Reconciliation Engine
 * Consolidated, Streamlined, and High-Performance Architecture
 */

const CONFIG = {
  FOLDER_NAME: "NBE Statements",
  PROCESSED_SUBFOLDER: "Processed",
  CREDIT_LIMIT: 100000.00,
  DEFAULT_FEE_PAYER: "",
  SHEETS: {
    TRANSACTIONS: "Transactions",
    INSTALLMENTS: "Installments",
    BANK_STATEMENT: "Bank Statement",
    ACTIVE_RECON: "Active Bill & Reconciliation",
    HISTORY_FORECAST: "Monthly History & Forecast",
    HISTORY: "History",
    SETTINGS: "Settings",
    MEMBER_BREAKDOWN: "Member Breakdown",
    RECON_STORE: "_Recon_Store",
    DECISIONS: "_Decisions",
    MIGRATION_REPORT: "Migration Report"
  },
  COLORS: {
    PRIMARY: "#1b5e20",       // NBE dark green
    PRIMARY_LIGHT: "#e8f5e9", // Soft light green
    HEADER: "#f3f3f3",        // Neutral light header
    PAID: "#d9ead3",          // Soft green badge
    DUE_SOON: "#fff2cc",      // Soft yellow badge
    OVERDUE: "#fce8e6",       // Soft red badge
    TEXT_MUTED: "#555555",
    BORDER: "#e0e0e0",
    ROW_ALT: "#fafafa"
  }
};

// ==========================================
// 1. UTILITIES & STRING/DATE SANITIZERS
// ==========================================

function parseDateValue(rawDate, tz) {
  if (!rawDate) return null;
  const targetTz = tz || "Africa/Cairo";

  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    const y = parseInt(Utilities.formatDate(rawDate, targetTz, "yyyy"), 10);
    const m = parseInt(Utilities.formatDate(rawDate, targetTz, "M"), 10) - 1;
    const d = parseInt(Utilities.formatDate(rawDate, targetTz, "d"), 10);
    return new Date(y, m, d, 12, 0, 0);
  }

  if (typeof rawDate === "number" && rawDate > 30000) {
    const dObj = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
    if (!isNaN(dObj.getTime())) {
      const y = parseInt(Utilities.formatDate(dObj, targetTz, "yyyy"), 10);
      const m = parseInt(Utilities.formatDate(dObj, targetTz, "M"), 10) - 1;
      const d = parseInt(Utilities.formatDate(dObj, targetTz, "d"), 10);
      return new Date(y, m, d, 12, 0, 0);
    }
  }

  const s = String(rawDate).trim();
  if (!s) return null;

  const iso = s.match(/^(\d{4})[-\/\.](\d{1,2})[-\/\.](\d{1,2})/);
  if (iso) {
    return new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10), 12, 0, 0);
  }

  const dmy = s.match(/^(\d{1,2})[-\/\.](\d{1,2})[-\/\.](\d{4})/);
  if (dmy) {
    let day = parseInt(dmy[1], 10);
    let month = parseInt(dmy[2], 10) - 1;
    const y = parseInt(dmy[3], 10);
    if (parseInt(dmy[2], 10) > 12 && day <= 12) {
      day = parseInt(dmy[2], 10);
      month = parseInt(dmy[1], 10) - 1;
    }
    return new Date(y, month, day, 12, 0, 0);
  }

  const monthNames = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const mdyWord = s.match(/(?:[a-zA-Z]+,\s*)?([a-zA-Z]{3,9})\s+(\d{1,2})(?:st|nd|rd|th)?,\s*(\d{4})/i);
  if (mdyWord) {
    const mStr = mdyWord[1].toLowerCase().substring(0, 3);
    const mIdx = monthNames.indexOf(mStr);
    const day = parseInt(mdyWord[2], 10);
    const year = parseInt(mdyWord[3], 10);
    if (mIdx >= 0) {
      return new Date(year, mIdx, day, 12, 0, 0);
    }
  }

  const wordMatch = s.match(/(\d{1,2})[-\/\.\s]+([a-zA-Z]{3,9})[-\/\.\s]+(\d{2,4})/);
  if (wordMatch) {
    const day = parseInt(wordMatch[1], 10);
    const mStr = wordMatch[2].toLowerCase().substring(0, 3);
    const mIdx = monthNames.indexOf(mStr);
    let year = parseInt(wordMatch[3], 10);
    if (year < 100) year += 2000;
    if (mIdx >= 0) {
      return new Date(year, mIdx, day, 12, 0, 0);
    }
  }

  const timestamp = Date.parse(s);
  if (!isNaN(timestamp)) {
    const dObj = new Date(timestamp);
    return new Date(dObj.getFullYear(), dObj.getMonth(), dObj.getDate(), 12, 0, 0);
  }

  return null;
}

function makeNormDateStr(rawDate, tz) {
  if (!rawDate) return "";
  const targetTz = tz || "Africa/Cairo";
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    return Utilities.formatDate(rawDate, targetTz, "yyyy-MM-dd");
  }
  const parsed = parseDateValue(rawDate, targetTz);
  if (parsed && !isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, targetTz, "yyyy-MM-dd");
  }
  return String(rawDate).trim().substring(0, 10);
}

function makeNormDesc(desc) {
  return String(desc || "")
    .toLowerCase()
    .replace(/[\u00A0\s]+/g, " ")
    .trim();
}

function cleanMerchantName(str) {
  if (!str) return "";
  let s = String(str).toLowerCase();
  const noise = [
    "fawry*", "paymob-*", "geideae*", "basata pay", "dcc markup fees-",
    "\"mobile token\"", "mobile token", "token", "egy", "cairo", "alex",
    "giza", "alexandria", "matrouh", "downtown", "smouha", "60 giz"
  ];
  noise.forEach(n => {
    s = s.replace(new RegExp(n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"), " ");
  });
  return s.trim();
}

function hasMerchantKeywordOverlap(desc1, desc2) {
  const raw1 = String(desc1 || "").toLowerCase();
  const raw2 = String(desc2 || "").toLowerCase();
  const c1 = cleanMerchantName(desc1);
  const c2 = cleanMerchantName(desc2);

  const words1 = c1.match(/[a-z0-9\u0600-\u06FF]{4,}/g) || [];
  const words2 = c2.match(/[a-z0-9\u0600-\u06FF]{4,}/g) || [];
  if (
    words1.some(w => words2.includes(w) || c2.includes(w)) ||
    words2.some(w => words1.includes(w) || c1.includes(w))
  ) {
    return true;
  }

  const gateways = ["fawry", "paymob", "geidea", "basata"];
  const strip1 = c1.replace(/[^a-z0-9]/g, "");
  const strip2 = c2.replace(/[^a-z0-9]/g, "");
  if (gateways.includes(strip1) && raw2.includes(strip1)) return true;
  if (gateways.includes(strip2) && raw1.includes(strip2)) return true;

  return false;
}

function isNeglectedPayer(name) {
  if (!name) return false;
  const s = String(name).toLowerCase().trim();
  return s.includes("neglect") || s.includes("ignore") || s.startsWith("🚫") || s === "none" || s === "-";
}

function isPaymentCreditDesc(desc) {
  return /\b(?:PAYMENT|DIRECT\s*DEBIT|SETTLEMENT|DEPOSIT|سداد|إيداع|تسوية)\b/i.test(String(desc || ""));
}

function normalizePersonName(name, optSettings) {
  if (!name) return "";
  // Strip any parenthetical suffix like (split), (Split), (cart), etc.
  const cleaned = String(name).replace(/\s*\([^)]*\)/g, " ").replace(/[\u00A0\s]+/g, " ").trim();
  if (!cleaned) return "";

  if (isNeglectedPayer(cleaned)) return "";
  const lower = cleaned.toLowerCase();
  if (lower === "shared") return "";

  const aliasMap = (optSettings && optSettings.nameAliases) ? optSettings.nameAliases : getActiveNameAliasMap_();
  if (aliasMap && aliasMap[lower]) {
    return aliasMap[lower];
  }

  return cleaned
    .toLowerCase()
    .split(" ")
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ");
}

function splitPayerNames(payerStr) {
  if (!payerStr) return [];
  const s = String(payerStr).replace(/\s*\([^)]*\)/g, " ").replace(/[\u00A0\s]+/g, " ").trim();
  if (isNeglectedPayer(s)) return [];

  return s
    .replace(/\s+and\s+/gi, " + ")
    .split(/[\+\,\/\&]/)
    .map(p => normalizePersonName(p))
    .filter(p => p && p.toLowerCase() !== "shared" && !isNeglectedPayer(p));
}

function formatPayerCell(val) {
  if (!val) return "";
  const s = String(val).replace(/[\u00A0\s]+/g, " ").trim();
  if (!s || isNeglectedPayer(s) || s.toLowerCase() === "shared") return s;

  if (/[\+\,\/\&]|\band\b/i.test(s)) {
    const parts = splitPayerNames(s);
    return parts.length > 0 ? parts.join(" + ") : s;
  }
  return normalizePersonName(s) || s;
}

function sanitizeAllPayerNames(ss) {
  if (!ss) return;
  try {
    const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
    if (txSheet && txSheet.getLastRow() >= 2) {
      const range = txSheet.getRange(2, 4, txSheet.getLastRow() - 1, 1);
      const vals = range.getValues();
      let changed = false;
      for (let i = 0; i < vals.length; i++) {
        const orig = String(vals[i][0] || "");
        if (orig) {
          const clean = formatPayerCell(orig);
          if (clean !== orig) { vals[i][0] = clean; changed = true; }
        }
      }
      if (changed) range.setValues(vals);
    }

    const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
    if (instSheet && instSheet.getLastRow() >= 2) {
      const range = instSheet.getRange(2, 9, instSheet.getLastRow() - 1, 1);
      const vals = range.getValues();
      let changed = false;
      for (let i = 0; i < vals.length; i++) {
        const orig = String(vals[i][0] || "");
        if (orig) {
          const clean = formatPayerCell(orig);
          if (clean !== orig) { vals[i][0] = clean; changed = true; }
        }
      }
      if (changed) range.setValues(vals);
    }
  } catch (err) {
    Logger.log("[sanitizeAllPayerNames] Error: " + err);
  }
}

function sanitizeTransactionDates(ss) {
  if (!ss) return;
  try {
    const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
    if (!txSheet || txSheet.getLastRow() < 2) return;
    const numRows = txSheet.getLastRow() - 1;
    const range = txSheet.getRange(2, 2, numRows, 1);
    const vals = range.getValues();
    let changed = false;

    for (let i = 0; i < vals.length; i++) {
      const cellVal = vals[i][0];
      if (!cellVal) continue;
      if (cellVal instanceof Date && !isNaN(cellVal.getTime())) {
        continue;
      }
      const parsed = parseDateValue(cellVal);
      if (parsed && !isNaN(parsed.getTime())) {
        vals[i][0] = parsed;
        changed = true;
      }
    }

    if (changed) {
      range.setValues(vals);
    }
    range.setNumberFormat("dddd, MMMM d, yyyy");
  } catch (err) {
    Logger.log("[sanitizeTransactionDates] Error: " + err);
  }
}

function getAllUniquePayers(ss) {
  const peopleSet = new Set();

  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (txSheet && txSheet.getLastRow() >= 2) {
    const vals = txSheet.getRange(2, 4, txSheet.getLastRow() - 1, 1).getValues();
    vals.forEach(r => {
      splitPayerNames(r[0]).forEach(name => {
        if (name && !isNeglectedPayer(name) && name.toLowerCase() !== "shared") peopleSet.add(name);
      });
    });
  }

  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  if (instSheet && instSheet.getLastRow() >= 2) {
    const vals = instSheet.getRange(2, 9, instSheet.getLastRow() - 1, 1).getValues();
    vals.forEach(r => {
      splitPayerNames(r[0]).forEach(name => {
        if (name && !isNeglectedPayer(name) && name.toLowerCase() !== "shared") peopleSet.add(name);
      });
    });
  }

  try {
    const decisions = loadDecisions(ss);
    Object.values(decisions.stmtAssign || {}).forEach(a => {
      if (a && a.target && !a.isNeglected && !isNeglectedPayer(a.target)) {
        splitPayerNames(a.target).forEach(name => { if (name) peopleSet.add(name); });
      }
    });
    Object.values(decisions.instOverride || {}).forEach(o => {
      if (o && o.target && !o.isNeglected && !isNeglectedPayer(o.target)) {
        const cleaned = String(o.target).replace(/^✅\s*Billed\s+to\s+/i, "");
        splitPayerNames(cleaned).forEach(name => { if (name) peopleSet.add(name); });
      }
    });
  } catch (e) { }

  const defFeePayer = getDefaultFeePayer(ss);
  if (defFeePayer && !isNeglectedPayer(defFeePayer)) {
    peopleSet.add(defFeePayer);
  }

  return Array.from(peopleSet).sort((a, b) => a.localeCompare(b));
}

function getColumnLetter(colIndex) {
  let letter = "";
  let temp = colIndex;
  while (temp > 0) {
    const rem = (temp - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    temp = Math.floor((temp - 1) / 26);
  }
  return letter;
}

function applySheetLayoutFitting(sheet, options) {
  if (!sheet) return;
  const opts = options || {};
  const bRow = opts.bannerRow || 1;
  const sRow = opts.subtitleRow || 2;
  try { sheet.setRowHeight(bRow, opts.bannerHeight || 30); } catch (e) { }
  try { sheet.setRowHeight(sRow, opts.subtitleHeight || 20); } catch (e) { }

  if (opts.customRowHeights) {
    Object.keys(opts.customRowHeights).forEach(r => {
      try { sheet.setRowHeight(parseInt(r, 10), opts.customRowHeights[r]); } catch (e) { }
    });
  }

  if (opts.autoResizeStartCol && opts.autoResizeColCount && opts.autoResizeColCount > 0) {
    try { sheet.autoResizeColumns(opts.autoResizeStartCol, opts.autoResizeColCount); } catch (e) { }
  }

  if (opts.explicitColWidths) {
    Object.keys(opts.explicitColWidths).forEach(c => {
      try { sheet.setColumnWidth(parseInt(c, 10), opts.explicitColWidths[c]); } catch (e) { }
    });
  }

  if (opts.colWidthBounds) {
    const bounds = opts.colWidthBounds;
    const startCol = opts.autoResizeStartCol || 2;
    const endCol = startCol + (opts.autoResizeColCount || 0) - 1;
    for (let c = startCol; c <= endCol; c++) {
      const bound = bounds[c] || bounds.default;
      if (bound) {
        try {
          const currentWidth = sheet.getColumnWidth(c);
          if (bound.min && currentWidth < bound.min) sheet.setColumnWidth(c, bound.min);
          else if (bound.max && currentWidth > bound.max) sheet.setColumnWidth(c, bound.max);
        } catch (e) { }
      }
    }
  }
}

function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function cleanupDeprecatedSheets(ss) {
  const deprecated = [
    "Monthly Overview",
    "Debt Breakdown",
    "Reconciliation",
    "Audit & Differences",
    "Payment History",
    "Monthly Debts"
  ];
  deprecated.forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) {
      try { ss.deleteSheet(s); } catch (e) { Logger.log("Cleaned up " + name); }
    }
  });
}

// ==========================================
// 2. BILLING CYCLES & CALENDAR LOGIC
// ==========================================

function getBillingCycleForPurchase(purchaseDate, tz) {
  if (!purchaseDate || isNaN(purchaseDate.getTime())) return null;
  const targetTz = tz || "Africa/Cairo";

  let day = parseInt(Utilities.formatDate(purchaseDate, targetTz, "d"), 10);
  let month = parseInt(Utilities.formatDate(purchaseDate, targetTz, "M"), 10) - 1;
  let year = parseInt(Utilities.formatDate(purchaseDate, targetTz, "yyyy"), 10);

  let cycleYear = year;
  let cycleMonth = month;

  // Statement cut-off rollover rule:
  // Purchases on the last calendar day of month M (e.g. Aug 31, Sep 30, Oct 31, Feb 28) roll over
  // into the billing cycle closed at the end of month M+1 and due on the 25th of month M+2.
  const nextDay = new Date(year, month, day + 1, 12, 0, 0);
  const isLastDay = (nextDay.getMonth() !== month);

  if (isLastDay) {
    cycleMonth = month + 1;
    if (cycleMonth > 11) {
      cycleMonth = 0;
      cycleYear++;
    }
  }

  let dueYear = cycleYear;
  let dueMonth = cycleMonth + 1;
  if (dueMonth > 11) {
    dueMonth = 0;
    dueYear++;
  }
  const dueDate = new Date(dueYear, dueMonth, 25, 12, 0, 0);
  const cycleDate = new Date(dueYear, dueMonth, 1, 12, 0, 0);

  return {
    cycleSortKey: Utilities.formatDate(cycleDate, targetTz, "yyyy-MM"),
    cycleLabel: Utilities.formatDate(cycleDate, targetTz, "MMMM yyyy"),
    cycleDate: cycleDate,
    dueDate: dueDate,
    dueDateLabel: Utilities.formatDate(dueDate, targetTz, "MMMM d, yyyy"),
    isDay31Rollover: isLastDay
  };
}

function getBillingCycleForInstallment(purchaseDate, installmentIndex, tz) {
  if (!purchaseDate || isNaN(purchaseDate.getTime())) return null;
  const targetTz = tz || "Africa/Cairo";

  const month = parseInt(Utilities.formatDate(purchaseDate, targetTz, "M"), 10) - 1;
  const year = parseInt(Utilities.formatDate(purchaseDate, targetTz, "yyyy"), 10);

  let cycleMonth = month + 1 + installmentIndex;
  let cycleYear = year + Math.floor(cycleMonth / 12);
  cycleMonth = ((cycleMonth % 12) + 12) % 12;

  let dueMonth = cycleMonth + 1;
  let dueYear = cycleYear;
  if (dueMonth > 11) {
    dueMonth = 0;
    dueYear++;
  }

  const cycleDate = new Date(dueYear, dueMonth, 1, 12, 0, 0);
  const dueDate = new Date(dueYear, dueMonth, 25, 12, 0, 0);

  return {
    cycleSortKey: Utilities.formatDate(cycleDate, targetTz, "yyyy-MM"),
    cycleLabel: Utilities.formatDate(cycleDate, targetTz, "MMMM yyyy"),
    cycleDate: cycleDate,
    dueDate: dueDate,
    dueDateLabel: Utilities.formatDate(dueDate, targetTz, "MMMM d, yyyy")
  };
}

function getCycleDateRange(cycleSortKey, tz) {
  if (!cycleSortKey) return null;
  const parts = cycleSortKey.split("-");
  const y = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  if (isNaN(y) || isNaN(m)) return null;

  const anchor = new Date(y, m - 1, 1, 12, 0, 0);
  let start = null;
  let end = null;

  for (let offset = -70; offset <= 5; offset++) {
    const d = new Date(anchor.getTime() + offset * 86400000);
    const info = getBillingCycleForPurchase(d, tz);
    if (info && info.cycleSortKey === cycleSortKey) {
      if (!start) start = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
      end = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999);
    }
  }

  return {
    start: start,
    end: end,
    startYear: start ? start.getFullYear() : y,
    startMonth: start ? start.getMonth() + 1 : m - 1,
    startDay: start ? start.getDate() : 1,
    endYear: end ? end.getFullYear() : y,
    endMonth: end ? end.getMonth() + 1 : m,
    endDay: end ? end.getDate() : 30
  };
}

function makePurchaseSumFormula(personCellRef, dateRange) {
  const { startYear, startMonth, startDay, endYear, endMonth, endDay } = dateRange;
  const startExpr = `DATE(${startYear},${startMonth},${startDay})`;
  const endExpr = `DATE(${endYear},${endMonth},${endDay})`;
  const single = `SUMIFS(Transactions!$E:$E, Transactions!$D:$D, TRIM(${personCellRef}), Transactions!$B:$B, ">="&${startExpr}, Transactions!$B:$B, "<"&(${endExpr}+1))`;
  const split1 = `0.5 * SUMIFS(Transactions!$E:$E, Transactions!$D:$D, "*" & TRIM(${personCellRef}) & "*+*", Transactions!$B:$B, ">="&${startExpr}, Transactions!$B:$B, "<"&(${endExpr}+1))`;
  const split2 = `0.5 * SUMIFS(Transactions!$E:$E, Transactions!$D:$D, "*+*" & TRIM(${personCellRef}) & "*", Transactions!$B:$B, ">="&${startExpr}, Transactions!$B:$B, "<"&(${endExpr}+1))`;
  return `${single} + ${split1} + ${split2}`;
}

function getStatementPeriod(stmtDate, minStmtTxDate, maxStmtTxDate) {
  if (!stmtDate && !minStmtTxDate) return { start: null, end: null };
  let pStart = null;
  let pEnd = null;

  if (minStmtTxDate) {
    pStart = new Date(minStmtTxDate.getFullYear(), minStmtTxDate.getMonth(), minStmtTxDate.getDate(), 0, 0, 0, 0);
  } else if (stmtDate) {
    pStart = new Date(stmtDate.getFullYear(), stmtDate.getMonth(), 1, 0, 0, 0, 0);
  }

  if (stmtDate) {
    pEnd = new Date(stmtDate.getFullYear(), stmtDate.getMonth(), stmtDate.getDate(), 23, 59, 59, 999);
  }
  if (maxStmtTxDate && (!pEnd || maxStmtTxDate > pEnd)) {
    pEnd = new Date(maxStmtTxDate.getFullYear(), maxStmtTxDate.getMonth(), maxStmtTxDate.getDate(), 23, 59, 59, 999);
  }

  return { start: pStart, end: pEnd };
}

// ==========================================
// 3. LEGACY RECONCILIATION KEY UTILITY
// ==========================================

function makeReconKey(dateVal, amountVal, descVal, tz, occurrenceIndex) {
  const dStr = makeNormDateStr(dateVal, tz);
  const amtNum = parseFloat(amountVal);
  const aStr = isNaN(amtNum) ? "0.00" : amtNum.toFixed(2);
  const dClean = makeNormDesc(descVal).substring(0, 30);
  const occStr = (occurrenceIndex !== undefined && occurrenceIndex !== null) ? `_#${occurrenceIndex}` : "";
  return `${dStr}_${aStr}_${dClean}${occStr}`;
}

// ==========================================
// 4. INSTALLMENT SPLITS & MATCHING
// ==========================================

function findSubsetCombination(items, targetAmt, maxK, tolerance) {
  const tol = tolerance || 0.10;
  const n = items.length;
  if (n < 2) return null;
  const sorted = items.slice().sort((a, b) => b.amount - a.amount);

  function search(startIdx, currentCombo, currentSum, k) {
    if (currentCombo.length >= 2 && Math.abs(currentSum - targetAmt) <= tol) {
      return currentCombo.slice();
    }
    if (currentCombo.length >= k || startIdx >= n) return null;

    for (let i = startIdx; i < n; i++) {
      const nextSum = currentSum + sorted[i].amount;
      if (nextSum > targetAmt + tol) continue;
      currentCombo.push(sorted[i]);
      const res = search(i + 1, currentCombo, nextSum, k);
      if (res) return res;
      currentCombo.pop();
    }
    return null;
  }

  const limit = Math.min(n, maxK || 6);
  for (let k = 2; k <= limit; k++) {
    const res = search(0, [], 0, k);
    if (res) return res;
  }
  return null;
}

// Installment link matching & progress moved to src/Installments.js

// ==========================================
// 5. CARD BALANCE & STATEMENT INFO
// ==========================================

function calculateCardBalance(ss, tz) {
  const creditLimit = getEffectiveCreditLimit(ss);
  let billedBalance = 0;
  let statementDate = null;
  let statementDueDate = null;
  let openingBalance = 0;
  let totalDebit = 0;
  let totalCredit = 0;
  let minStmtTxDate = null;
  let maxStmtTxDate = null;

  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  if (stmtSheet && stmtSheet.getLastRow() >= 4) {
    const metaVals = stmtSheet.getRange(2, 1, 3, 6).getValues();
    statementDate = parseDateValue(metaVals[0][1], tz);
    statementDueDate = parseDateValue(metaVals[0][3], tz);

    const openVal = metaVals[1][1];
    if (!isNaN(parseFloat(openVal))) openingBalance = parseFloat(openVal);

    const closeVal = metaVals[1][3];
    if (!isNaN(parseFloat(closeVal))) billedBalance = parseFloat(closeVal);

    const totDebVal = metaVals[1][5];
    const hasTotDeb = (typeof totDebVal === "number" && !isNaN(totDebVal)) ||
      (typeof totDebVal === "string" && totDebVal.trim() !== "" && !totDebVal.trim().startsWith("=") && !isNaN(parseFloat(totDebVal)));
    if (hasTotDeb) totalDebit = parseFloat(totDebVal);

    const totCredVal = metaVals[2][1];
    const hasTotCred = (typeof totCredVal === "number" && !isNaN(totCredVal)) ||
      (typeof totCredVal === "string" && totCredVal.trim() !== "" && !totCredVal.trim().startsWith("=") && !isNaN(parseFloat(totCredVal)));
    if (hasTotCred) totalCredit = parseFloat(totCredVal);

    if (stmtSheet.getLastRow() >= 7) {
      const stmtNumRows = stmtSheet.getLastRow() - 6;
      const txDates = stmtSheet.getRange(7, 2, stmtNumRows, 2).getValues();
      txDates.forEach(r => {
        const d1 = parseDateValue(r[0], tz);
        const d2 = parseDateValue(r[1], tz);
        const d = d1 || d2;
        if (d) {
          if (!minStmtTxDate || d < minStmtTxDate) minStmtTxDate = d;
          if (!maxStmtTxDate || d > maxStmtTxDate) maxStmtTxDate = d;
        }
      });

      if (!hasTotDeb || !hasTotCred) {
        const typeAmtRows = stmtSheet.getRange(7, 5, stmtNumRows, 2).getValues();
        let fallbackDebit = 0;
        let fallbackCredit = 0;
        typeAmtRows.forEach(r => {
          const t = String(r[0] || "").trim().toUpperCase();
          const a = Math.abs(parseFloat(r[1]) || 0);
          if (t === "CREDIT") fallbackCredit += a;
          else if (a > 0) fallbackDebit += a;
        });
        if (!hasTotDeb) totalDebit = Number(fallbackDebit.toFixed(2));
        if (!hasTotCred) totalCredit = Number(fallbackCredit.toFixed(2));
      }
    }
  }

  const stmtPeriod = getStatementPeriod(statementDate, minStmtTxDate, maxStmtTxDate);
  if (!statementDueDate && statementDate) {
    statementDueDate = new Date(statementDate.getFullYear(), statementDate.getMonth() + 1, 25);
  }

  const activeAnchor = statementDueDate || new Date();
  const activeCycleKey = Utilities.formatDate(activeAnchor, tz, "yyyy-MM");
  const activeCycleLabel = Utilities.formatDate(activeAnchor, tz, "MMMM yyyy");

  let totalBlockedInstallments = 0;
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  if (instSheet && instSheet.getLastRow() >= 2) {
    const decisions = loadDecisions(ss);
    const numInstRows = instSheet.getLastRow() - 1;
    const instIdCol = findExistingIdColumn(instSheet, 1);
    const instIds = instIdCol
      ? instSheet.getRange(2, instIdCol, numInstRows, 1).getValues().map(r => r[0])
      : null;

    const sheetInstList = readSheetInstallmentRows(instSheet, instIds, decisions, activeCycleKey, tz);

    const cardStmtInstallments = [];
    if (stmtSheet && stmtSheet.getLastRow() >= 7) {
      const stmtLastRow = stmtSheet.getLastRow();
      const stmtRows = stmtSheet.getRange(7, 1, stmtLastRow - 6, 7).getValues();
      const stmtIdCol = findExistingIdColumn(stmtSheet, 6);
      const stmtIds = stmtIdCol ? stmtSheet.getRange(7, stmtIdCol, stmtLastRow - 6, 1).getValues() : null;
      stmtRows.forEach((row, idx) => {
        const type = String(row[4] || "").trim().toUpperCase();
        const desc = String(row[3] || "").trim();
        const amt = parseFloat(row[5]) || 0;
        if (type === "INSTALLMENT") {
          const parsed = parseInstallmentProgressText(desc);
          cardStmtInstallments.push({
            stmtId: stmtIds ? String(stmtIds[idx][0] || "").trim() : "",
            stmtRowIndex: idx,
            desc,
            amount: Math.abs(amt),
            instCurrent: parsed.instCurrent,
            instTotal: parsed.instTotal
          });
        }
      });
    }

    let isBillPaid = false;
    const arSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
    if (arSheet) {
      const h6Date = parseDateValue(arSheet.getRange(6, 8).getValue(), tz);
      if (h6Date && !isNaN(h6Date.getTime())) isBillPaid = true;
    }
    if (!isBillPaid && decisions.paymentRecords && decisions.paymentRecords[activeCycleKey] && decisions.paymentRecords[activeCycleKey].dateStr) {
      isBillPaid = true;
    }

    const progressMap = (cardStmtInstallments.length > 0 && sheetInstList.length > 0)
      ? computeBankDerivedInstallmentProgress(
          matchStatementInstallmentsByLink(cardStmtInstallments, sheetInstList, decisions).matches,
          sheetInstList,
          isBillPaid
        )
      : {};

    sheetInstList.forEach(sh => {
      const prog = progressMap[sh.index];
      if (prog && prog.onStatement) {
        if (!prog.isCompleted) totalBlockedInstallments += (prog.remainingMonths * sh.emi);
      } else if (!sh.isCompleted) {
        totalBlockedInstallments += Math.max(0, (sh.duration || 0) - sh.paymentsMade) * sh.emi;
      }
    });
  }

  let unbilledNewPurchases = 0;
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (txSheet && txSheet.getLastRow() >= 2) {
    const txData = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 5).getValues();
    txData.forEach(row => {
      const pDate = parseDateValue(row[1], tz);
      const amt = parseFloat(row[4]);
      if (pDate && !isNaN(amt) && amt > 0) {
        const cycleInfo = getBillingCycleForPurchase(pDate, tz);
        if (cycleInfo && cycleInfo.cycleSortKey > activeCycleKey) {
          unbilledNewPurchases += amt;
        }
      }
    });
  }

  const effectiveBilled = billedBalance;
  const totalUtilized = effectiveBilled + totalBlockedInstallments + unbilledNewPurchases;
  const availableBalanceNow = Math.max(0, creditLimit - totalUtilized);
  const availableAfterSettlement = Math.max(0, creditLimit - totalBlockedInstallments - unbilledNewPurchases);

  return {
    creditLimit,
    billedBalance,
    openingBalance,
    totalDebit,
    totalCredit,
    statementDate,
    statementDueDate,
    activeCycleKey,
    activeCycleLabel,
    minStmtTxDate,
    maxStmtTxDate,
    stmtPeriod,
    totalBlockedInstallments,
    unbilledNewPurchases,
    totalUtilized,
    availableBalanceNow,
    availableAfterSettlement
  };
}

// ==========================================
// 6. PDF OCR EXTRACTION, PARSER, TIE-OUT VALIDATION & IMPORT -> see src/Parser.js
// ==========================================

// ==========================================
// 7. ACTIVE RECONCILIATION & WATERFALL ENGINE
// ==========================================

function reconcileActiveStatement(ss, tz, optContext) {
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);

  const decisions = loadDecisions(ss);
  loadSettings(ss);
  syncSettingsToDecisions(ss, decisions);

  if (!stmtSheet || stmtSheet.getLastRow() < 7) {
    return { error: "Please upload and process a bank statement to view active bill reconciliation." };
  }

  const idInfo = ensureAllSheetRowIds(ss, decisions);
  const parseHealth = (optContext && optContext.parseHealth) || evaluateBankStatementHealth(ss, tz, decisions);

  // 1. Read Statement Rows (respect explicit Col E Type — no hardcoded amount/merchant flips)
  const stmtLastRow = stmtSheet.getLastRow();
  const stmtData = stmtSheet.getRange(7, 1, stmtLastRow - 6, 7).getValues();

  const stmtDebits = [];
  const stmtInstallments = [];
  const stmtFees = [];
  const stmtCredits = [];
  let minStmtTxDate = null;
  let maxStmtTxDate = null;

  const validStmtTypes = ["DEBIT", "CREDIT", "INSTALLMENT", "FEE"];

  stmtData.forEach((row, idx) => {
    const pDate = parseDateValue(row[1], tz);
    if (pDate) {
      if (!minStmtTxDate || pDate < minStmtTxDate) minStmtTxDate = pDate;
      if (!maxStmtTxDate || pDate > maxStmtTxDate) maxStmtTxDate = pDate;
    }

    let type = String(row[4] || "").trim().toUpperCase();
    const desc = String(row[3] || "").trim();
    const amt = parseFloat(row[5]) || 0;

    if (!validStmtTypes.includes(type)) {
      type = amt < 0 ? "CREDIT" : "DEBIT";
    }

    const item = {
      id: idx + 1,
      stmtId: (idInfo.stmt.ids && idInfo.stmt.ids[idx]) || "",
      stmtRowIndex: idx,
      dateStr: String(row[1]),
      date: pDate,
      desc: desc,
      amount: Math.abs(amt),
      type: type,
      matched: false,
      assignedPayer: "",
      note: ""
    };

    if (type === "INSTALLMENT") {
      let instCurrent = null, instTotal = null;
      const mInst = item.desc.match(/(\d+)\s+OF\s+(\d+)/i);
      if (mInst) { instCurrent = parseInt(mInst[1], 10); instTotal = parseInt(mInst[2], 10); }
      item.instCurrent = instCurrent;
      item.instTotal = instTotal;
      stmtInstallments.push(item);
    } else if (type === "FEE") {
      stmtFees.push(item);
    } else if (type === "CREDIT") {
      stmtCredits.push(item);
    } else {
      stmtDebits.push(item);
    }
  });

  const cardBal = calculateCardBalance(ss, tz);
  const stmtPeriod = cardBal.stmtPeriod || getStatementPeriod(cardBal.statementDate, minStmtTxDate, maxStmtTxDate);

  // 2. Read Transactions Sheet (Columns A:E only — never mutate user columns)
  const txDebits = [];
  const pendingNextTx = [];
  if (txSheet && txSheet.getLastRow() >= 2) {
    const txData = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 5).getValues();
    txData.forEach((row, idx) => {
      const pDate = parseDateValue(row[1], tz);
      const amt = parseFloat(row[4]);
      const person = String(row[3] || "").trim();
      const desc = String(row[2] || "").trim();
      if (pDate && !isNaN(amt) && amt > 0 && person) {
        const cycle = getBillingCycleForPurchase(pDate, tz);
        const belongsToCycle = (cycle && cycle.cycleSortKey === cardBal.activeCycleKey) ||
                               (!cycle && stmtPeriod.start && stmtPeriod.end && pDate >= stmtPeriod.start && pDate <= stmtPeriod.end);
        if (belongsToCycle) {
          txDebits.push({
            txId: (idInfo.tx.ids && idInfo.tx.ids[idx]) || "",
            rowNum: idx + 2,
            date: pDate,
            desc: desc,
            person: person,
            amount: amt,
            matched: false
          });
        } else if (cycle && cardBal.activeCycleKey && cycle.cycleSortKey > cardBal.activeCycleKey) {
          pendingNextTx.push({
            txId: (idInfo.tx.ids && idInfo.tx.ids[idx]) || "",
            rowNum: idx + 2,
            date: pDate,
            desc: desc,
            person: person,
            amount: amt,
            cycleSortKey: cycle.cycleSortKey,
            cycleLabel: cycle.cycleLabel
          });
        }
      }
    });
  }

  // 3. Read Installments Sheet
  const sheetInstallmentRows = readSheetInstallmentRows(instSheet, idInfo.inst.ids, decisions, cardBal.activeCycleKey, tz);

  // Stage 1 one-time migration from legacy _Recon_Store (pre-matching: payment records + confirmed matches)
  const legacyMigrationCtx = migrateLegacyPreMatch(ss, tz, decisions, { stmtDebits, txDebits });

  // 4. Match Installments by Link & Build Table 4 Reconciliation Items
  const instMatchResult = matchStatementInstallmentsByLink(stmtInstallments, sheetInstallmentRows, decisions);
  const instMatches = instMatchResult.matches;
  const instMatchMap = {};
  const matchedStmtIdxSet = instMatchResult.matchedStmtIdxSet;
  const matchedSheetIdxSet = instMatchResult.matchedSheetIdxSet;
  instMatches.forEach(m => {
    instMatchMap[m.stmtIdx] = m.payerLabel;
    stmtInstallments[m.stmtIdx].matched = true;
    stmtInstallments[m.stmtIdx].assignedPayer = m.payerLabel;
  });

  // Persist bank-observed installment progress into _Decisions (INST_PROGRESS)
  let isInstBillPaid = Boolean(
    decisions.paymentRecords &&
    decisions.paymentRecords[cardBal.activeCycleKey] &&
    decisions.paymentRecords[cardBal.activeCycleKey].dateStr
  );
  if (!isInstBillPaid) {
    const arSheetCheck = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
    if (arSheetCheck) {
      const h6Date = parseDateValue(arSheetCheck.getRange(6, 8).getValue(), tz);
      if (h6Date && !isNaN(h6Date.getTime())) isInstBillPaid = true;
    }
  }
  const instProgressMap = computeBankDerivedInstallmentProgress(instMatches, sheetInstallmentRows, isInstBillPaid);
  persistBankObservedInstallmentProgress(instProgressMap, decisions, cardBal.activeCycleKey);

  // Prune Table 2 assignment overrides only for linked/matched statement installment IDs
  stmtInstallments.forEach((st, sIdx) => {
    if (matchedStmtIdxSet.has(sIdx) && st.stmtId && decisions.stmtAssign && decisions.stmtAssign[st.stmtId]) {
      clearStmtAssign(decisions, st.stmtId);
    }
  });

  // Unlinked statement installments move to Needs Pairing (suggestions first, preserving row order)
  const rawUnlinkedStmtInst = instMatchResult.unmatchedStmtInstallments || [];
  const unlinkedStmtInstallments = [
    ...rawUnlinkedStmtInst.filter(st => st.suggestedInstLinks && st.suggestedInstLinks.length > 0),
    ...rawUnlinkedStmtInst.filter(st => !st.suggestedInstLinks || st.suggestedInstLinks.length === 0)
  ];

  const unmatchedSheetInstallments = sheetInstallmentRows.filter(
    sh => !matchedSheetIdxSet.has(sh.index) && !sh.isCompleted && sh.status.toLowerCase() !== "completed"
  );

  const unmatchedSheetItems = [];
  unmatchedSheetInstallments.forEach(sh => {
    const payers = splitPayerNames(sh.payer);
    const targetPayers = payers.length > 0 ? payers : [sh.payer];
    const shareEmi = sh.emi / targetPayers.length;
    targetPayers.forEach(p => {
      unmatchedSheetItems.push({
        subjectId: buildInstSubjectId(null, sh.instId, cardBal.activeCycleKey, p, targetPayers.length),
        date: sh.startDate || sh.dateStr || "",
        desc: `${sh.desc} (${sh.paymentsMade}/${sh.duration || "?"} Paid)`,
        amount: Number(shareEmi.toFixed(2)),
        sourceCategory: "⏳ Sheet Only",
        defaultPayer: p,
        defaultAction: "🚫 Not on Statement (Skip)",
        defaultNote: `In Installments sheet (${p}) — not billed on current statement`
      });
    });
  });

  const matchedInstItems = [];
  instMatches.forEach(m => {
    const isLockedLink = m.tier === "locked" || m.tier === "locked-plan";
    const entries = [];
    m.sheetRows.forEach(sh => {
      const payers = splitPayerNames(sh.payer);
      const targetPayers = payers.length > 0 ? payers : [sh.payer];
      const shareEmi = sh.emi / targetPayers.length;
      const progressStr = (m.stmt.instCurrent && m.stmt.instTotal)
        ? `${m.stmt.instCurrent} of ${m.stmt.instTotal}`
        : (sh.duration ? `${sh.paymentsMade + 1} of ${sh.duration}` : "Active");
      const isSplit = m.sheetRows.length > 1 || targetPayers.length > 1;
      const sourceCategory = isLockedLink
        ? (isSplit ? "✅ Linked Split" : "✅ Linked Plan")
        : (isSplit ? "✅ Matched Split" : "✅ Matched Statement");
      const rowEntries = [];
      targetPayers.forEach(p => {
        rowEntries.push({
          subjectId: buildInstSubjectId(m.stmt.stmtId, sh.instId, cardBal.activeCycleKey, p, targetPayers.length),
          date: m.stmt.date || sh.startDate || m.stmt.dateStr || "",
          desc: `${sh.desc} (${progressStr})`,
          rawEmi: shareEmi,
          amount: Number(shareEmi.toFixed(2)),
          sourceCategory: sourceCategory,
          defaultPayer: p,
          defaultAction: `✅ Billed to ${p}`,
          defaultNote: `${isLockedLink ? "Linked" : "Matched"} Stmt: ${m.stmt.desc}`
        });
      });

      if (targetPayers.length > 1 && rowEntries.length > 0) {
        const rowRoundedSum = rowEntries.reduce((s, e) => s + e.amount, 0);
        const rowResidual = Number((Number(sh.emi.toFixed(2)) - rowRoundedSum).toFixed(2));
        if (Math.abs(rowResidual) >= 0.01) {
          const lastRowEntry = rowEntries[rowEntries.length - 1];
          lastRowEntry.amount = Number((lastRowEntry.amount + rowResidual).toFixed(2));
        }
      }

      rowEntries.forEach(e => entries.push(e));
    });

    if (entries.length > 0) {
      const rawSum = entries.reduce((s, e) => s + e.rawEmi, 0);
      const roundedSum = entries.reduce((s, e) => s + e.amount, 0);
      const roundingResidual = Number((m.stmt.amount - roundedSum).toFixed(2));
      if (m.sheetRows.length === 1 && Math.abs(roundingResidual) >= 0.01) {
        const last = entries[entries.length - 1];
        last.amount = Number((last.amount + roundingResidual).toFixed(2));
      }
      const rawDiff = m.stmt.amount - rawSum;
      if (Math.abs(rawDiff) >= 0.01) {
        entries.forEach(e => {
          e.defaultNote += ` (⚠️ Stmt: ${m.stmt.amount.toFixed(2)} vs Sheet: ${rawSum.toFixed(2)})`;
        });
      }
      entries.forEach(e => matchedInstItems.push(e));
    }
  });

  const installmentReconItems = [...unmatchedSheetItems, ...matchedInstItems];

  // 5. Confidence-Tiered Matching for Purchases (DEBIT)
  const matchResult = matchStatementDebits(stmtDebits, txDebits, decisions, tz);
  const suggestedMatches = matchResult.suggestedMatches || [];

  const suggestedSet = new Set(suggestedMatches);
  const unmatchedDebits = stmtDebits.filter(st => !st.matched);
  const missingInSheet = [
    ...unmatchedDebits.filter(st => suggestedSet.has(st)),
    ...unmatchedDebits.filter(st => !suggestedSet.has(st))
  ];

  // Prune Table 2 assignment overrides only for matched statement IDs
  stmtDebits.forEach(st => {
    if (st.matched && st.stmtId && decisions.stmtAssign[st.stmtId]) {
      clearStmtAssign(decisions, st.stmtId);
    }
  });

  const unmatchedTx = txDebits.filter(tx => !tx.matched);

  // Build Orphan Logs (in-period transactions without a paired bank line — NOT billed on current statement)
  const orphanLogs = [];
  const txOccTracker = {};
  unmatchedTx.forEach((tx, idx) => {
    const payers = splitPayerNames(tx.person);
    const baseHints = tx.suggestedStmts || [];
    if (payers.length <= 1) {
      const p = payers[0] || tx.person;
      const exId = `ORPH-${idx + 1}`;
      const baseKey = makeReconKey(tx.date, tx.amount, tx.desc, tz);
      txOccTracker[baseKey] = (txOccTracker[baseKey] || 0) + 1;
      const occ = txOccTracker[baseKey];
      const pKeyWithOcc = `${baseKey}_#${occ}`;
      orphanLogs.push({
        id: exId,
        subjectId: buildTxSubjectId(tx.txId, p, 1),
        txId: tx.txId,
        date: tx.date,
        desc: tx.desc,
        amount: tx.amount,
        person: p,
        isNeglected: false,
        suggestedStmts: baseHints,
        _legacyKeys: [pKeyWithOcc, occ === 1 ? baseKey : null, `TX-${idx + 1}`, exId].filter(Boolean)
      });
    } else {
      const splitAmt = tx.amount / payers.length;
      const legacyWholeKey = makeReconKey(tx.date, tx.amount, tx.desc, tz);
      payers.forEach((p, pIdx) => {
        const exId = `ORPH-${idx + 1}.${pIdx + 1}`;
        const baseKey = makeReconKey(tx.date, splitAmt, tx.desc, tz);
        txOccTracker[baseKey] = (txOccTracker[baseKey] || 0) + 1;
        const occ = txOccTracker[baseKey];
        const pKeyWithOcc = `${baseKey}_#${occ}`;
        orphanLogs.push({
          id: exId,
          subjectId: buildTxSubjectId(tx.txId, p, payers.length),
          txId: tx.txId,
          date: tx.date,
          desc: `${tx.desc} (Split ${pIdx + 1}/${payers.length})`,
          amount: splitAmt,
          person: p,
          isNeglected: false,
          suggestedStmts: baseHints,
          _legacyKeys: [pKeyWithOcc, occ === 1 ? baseKey : null, `${legacyWholeKey}_#${pIdx + 1}`, `TX-${idx + 1}.${pIdx + 1}`, exId].filter(Boolean)
        });
      });
    }
  });

  const pendingNextItems = [];
  pendingNextTx.forEach((tx, idx) => {
    const payers = splitPayerNames(tx.person);
    if (payers.length <= 1) {
      pendingNextItems.push({
        id: `PEND-${idx + 1}`,
        txId: tx.txId,
        date: tx.date,
        desc: tx.desc,
        amount: tx.amount,
        person: payers[0] || tx.person,
        cycleLabel: tx.cycleLabel || ""
      });
    } else {
      const splitAmt = tx.amount / payers.length;
      payers.forEach((p, pIdx) => {
        pendingNextItems.push({
          id: `PEND-${idx + 1}.${pIdx + 1}`,
          txId: tx.txId,
          date: tx.date,
          desc: `${tx.desc} (Split ${pIdx + 1}/${payers.length})`,
          amount: splitAmt,
          person: p,
          cycleLabel: tx.cycleLabel || ""
        });
      });
    }
  });

  const expandedUnmatched = orphanLogs;

  // Stage 2 one-time migration from legacy _Recon_Store (Table 2, Table 3, Table 4)
  if (legacyMigrationCtx && !decisions.isMigrated) {
    migrateLegacyPostMatch(ss, tz, decisions, legacyMigrationCtx, {
      stmtFees,
      missingInSheet,
      stmtCredits,
      expandedUnmatched,
      installmentReconItems
    });
  }

  // Apply Table 4 Installment Overrides purely by permanent SubjectId
  installmentReconItems.forEach((item, idx) => {
    item.id = `INST-${idx + 1}`;
    item.primaryKey = item.subjectId;
    const saved = item.subjectId && decisions.instOverride && decisions.instOverride[item.subjectId];
    if (saved && saved.target !== undefined) {
      item.action = saved.target;
      item.note = saved.note || item.defaultNote;
    } else {
      item.action = item.defaultAction;
      item.note = item.defaultNote;
    }
    item.isNeglected = isNeglectedPayer(item.action);
  });

  // Apply Table 2 Assignments purely by permanent Statement ID
  function applyTable2Decisions(items, options) {
    items.forEach(it => {
      const saved = it.stmtId && decisions.stmtAssign && decisions.stmtAssign[it.stmtId];
      if (saved && saved.target) {
        it.assignedPayer = (saved.target === "NEGLECT" || saved.isNeglected)
          ? "🚫 Neglect / Ignore"
          : (formatPayerCell(saved.target) || saved.target);
        it.note = saved.note || options.defaultNote(it);
      } else if (options.defaultPayer !== undefined) {
        it.assignedPayer = typeof options.defaultPayer === "function"
          ? options.defaultPayer(it)
          : options.defaultPayer;
        it.note = options.defaultNote(it);
      }
    });
  }

  // 6. Fees Default to configured defaultFeePayer (or restored from _Decisions)
  const defaultFeePayer = getDefaultFeePayer(ss);
  applyTable2Decisions(stmtFees, {
    defaultPayer: defaultFeePayer,
    defaultNote: () => "Bank Fee / Tax"
  });

  // 7. Unlinked Statement Installments (or restored from _Decisions)
  applyTable2Decisions(unlinkedStmtInstallments, {
    defaultNote: () => "On bank statement — link to plan or assign payer"
  });

  // 8. Missing in sheet purchases (or restored from _Decisions)
  applyTable2Decisions(missingInSheet, {
    defaultNote: m => cleanMerchantName(m.desc)
  });

  // 9. Refunds / Credits (payment credits neglected by default unless user assigns payer)
  applyTable2Decisions(stmtCredits, {
    defaultPayer: cr => isPaymentCreditDesc(cr.desc) ? "🚫 Neglect / Ignore" : defaultFeePayer,
    defaultNote: cr => isPaymentCreditDesc(cr.desc)
      ? "Bill Payment / Settlement (Neglected by default)"
      : "Merchant Refund (Reduces Debt)"
  });

  // 10. Optional Assignable Prior Cycle Carryover Item (Opening Balance - Neglected Bill Payments)
  let neglectedBillPayments = 0;
  stmtCredits.forEach(cr => {
    if (isPaymentCreditDesc(cr.desc) && isNeglectedPayer(cr.assignedPayer)) {
      neglectedBillPayments += cr.amount;
    }
  });
  const carrySubjectId = "CARRY::" + (cardBal.activeCycleKey || "ACTIVE");
  const rawCarryover = Number(((cardBal.openingBalance || 0) - neglectedBillPayments).toFixed(2));
  const savedCarry = decisions.stmtAssign && decisions.stmtAssign[carrySubjectId];
  const hasSavedCarry = Boolean(
    savedCarry &&
    savedCarry.target &&
    savedCarry.target !== "NEGLECT" &&
    !isNeglectedPayer(savedCarry.target) &&
    !savedCarry.isNeglected
  );
  const carryoverItem = {
    id: "CARRY-1",
    subjectId: carrySubjectId,
    date: cardBal.statementDate || "",
    dateStr: cardBal.statementDate ? Utilities.formatDate(cardBal.statementDate, tz, "yyyy-MM-dd") : "",
    desc: "Prior Cycle Carryover (Opening − Neglected Payments)",
    amount: rawCarryover,
    category: "Prior Cycle Carryover",
    assignedPayer: hasSavedCarry
      ? (formatPayerCell(savedCarry.target) || savedCarry.target)
      : "🚫 Keep Unassigned (Prior Cycle)",
    isAssigned: hasSavedCarry,
    note: (savedCarry && savedCarry.note) ? savedCarry.note : "Optional: assign to a member, or leave unassigned",
    showRow: Math.abs(rawCarryover) >= 0.01 || hasSavedCarry
  };

  // Apply Orphan Log dismiss decisions purely by permanent SubjectId
  orphanLogs.forEach(tx => {
    tx.isNeglected = Boolean(tx.subjectId && decisions.txNeglect && decisions.txNeglect[tx.subjectId]);
    delete tx._legacyKeys;
  });

  if (decisions._dirty) {
    saveDecisions(ss, decisions);
  }

  return {
    stmtDebits,
    stmtInstallments,
    instMatches,
    instProgressMap,
    unlinkedStmtInstallments,
    unmatchedSheetInstallments,
    installmentReconItems,
    stmtFees,
    stmtCredits,
    carryoverItem,
    suggestedMatches,
    missingInSheet,
    txDebits,
    unmatchedTx,
    orphanLogs,
    expandedUnmatched,
    pendingNextItems,
    cardBal,
    parseHealth,
    decisions,
    reconStore: decisions,
    tz
  };
}

// ==========================================
// 8. DEBT LEDGER & WATERFALL CALCULATION
// ==========================================

function buildDebtLedger(ss, tz, recon) {
  const cardBal = recon.cardBal;
  const activeCycleKey = cardBal.activeCycleKey;
  const activeCycleLabel = cardBal.activeCycleLabel;
  const activeDueDate = cardBal.statementDueDate || new Date();

  const personShares = {};
  const allPayers = getAllUniquePayers(ss);
  allPayers.forEach(p => {
    personShares[p] = {
      person: p,
      matchedPurchases: 0,
      unmatchedPurchases: 0,
      purchases: 0,
      installments: 0,
      fees: 0,
      unrecorded: 0,
      refunds: 0,
      netTotal: 0
    };
  });

  function addShare(payerStr, amt, category) {
    if (!payerStr || isNeglectedPayer(payerStr)) return;
    const people = splitPayerNames(payerStr);
    if (people.length === 0) return;
    const splitAmt = amt / people.length;

    people.forEach(p => {
      if (!personShares[p]) {
        personShares[p] = {
          person: p,
          matchedPurchases: 0,
          unmatchedPurchases: 0,
          purchases: 0,
          installments: 0,
          fees: 0,
          unrecorded: 0,
          refunds: 0,
          netTotal: 0
        };
      }
      if (category === "MatchedPurchase") {
        personShares[p].matchedPurchases += splitAmt;
        personShares[p].purchases += splitAmt;
      } else if (category === "UnmatchedPurchase") {
        personShares[p].unmatchedPurchases += splitAmt;
        personShares[p].purchases += splitAmt;
      } else if (category === "Installment") {
        personShares[p].installments += splitAmt;
      } else if (category === "Fee") {
        personShares[p].fees += splitAmt;
      } else if (category === "Unrecorded") {
        personShares[p].unrecorded += splitAmt;
      } else if (category === "Refund") {
        personShares[p].refunds += splitAmt;
      }
    });
  }

  // 1a. Attributed Matched Purchases (Statement Debits matched to Sheet)
  recon.stmtDebits.forEach(st => {
    if (!st.matched) return;
    if (st.matchedItems && st.matchedItems.length > 1) {
      st.matchedItems.forEach(tx => {
        if (tx.person && !isNeglectedPayer(tx.person)) {
          addShare(tx.person, tx.amount, "MatchedPurchase");
        }
      });
    } else if (st.isPartialMatch && st.matchedItems && st.matchedItems.length === 1) {
      const tx = st.matchedItems[0];
      if (tx.person && !isNeglectedPayer(tx.person)) {
        addShare(tx.person, tx.amount, "MatchedPurchase");
      }
    } else if (st.assignedPayer && !isNeglectedPayer(st.assignedPayer)) {
      addShare(st.assignedPayer, st.amount, "MatchedPurchase");
    }
  });

  // Note (Q5c/Q13): In-period unpaired Orphan Logs are NOT billed on top of the bank statement;
  // purchases only count once paired with a statement line or assigned in Needs Pairing.

  // 2. Attributed Installments (From Table 4 Items + Assigned Unlinked Installments in Needs Pairing)
  let unresolvedInstCount = 0;
  let unresolvedInstTotal = 0;
  let neglectedTotal = 0;
  let neglectedCount = 0;

  (recon.installmentReconItems || []).forEach(inst => {
    const act = String(inst.action || "").trim();
    if (!act) {
      unresolvedInstCount++;
      unresolvedInstTotal += (parseFloat(inst.amount) || 0);
      return;
    }
    if (isNeglectedPayer(act)) return;
    const cleanPayer = act.replace(/^✅\s*Billed\s+to\s+/i, "").trim();
    if (cleanPayer && !isNeglectedPayer(cleanPayer)) {
      addShare(cleanPayer, inst.amount, "Installment");
    }
  });

  (recon.unlinkedStmtInstallments || []).forEach(u => {
    if (isNeglectedPayer(u.assignedPayer)) {
      neglectedTotal += u.amount;
      neglectedCount++;
    } else if (u.assignedPayer) {
      addShare(u.assignedPayer, u.amount, "Installment");
    } else {
      unresolvedInstCount++;
      unresolvedInstTotal += (parseFloat(u.amount) || 0);
    }
  });

  // 3. Attributed Fees (Default: configured defaultFeePayer or overridden)
  let unassignedTotal = 0;
  let unassignedCount = 0;
  let unresolvedFeesCount = 0;
  let unresolvedFeesTotal = 0;

  recon.stmtFees.forEach(fee => {
    if (isNeglectedPayer(fee.assignedPayer)) {
      neglectedTotal += fee.amount;
      neglectedCount++;
    } else if (fee.assignedPayer) {
      addShare(fee.assignedPayer, fee.amount, "Fee");
    } else {
      unresolvedFeesCount++;
      unresolvedFeesTotal += fee.amount;
    }
  });

  // 4. Attributed Unrecorded Missing Statement Debits
  recon.missingInSheet.forEach(m => {
    if (isNeglectedPayer(m.assignedPayer)) {
      neglectedTotal += m.amount;
      neglectedCount++;
    } else if (m.assignedPayer) {
      addShare(m.assignedPayer, m.amount, "Unrecorded");
    } else {
      unassignedTotal += m.amount;
      unassignedCount++;
    }
  });

  // 5. Attributed Refunds / Credits (only count neglected bill payments in carriedOver so assigned credits aren't double-counted)
  let billPaymentsTotal = 0;
  let unresolvedCreditsCount = 0;
  let unresolvedCreditsTotal = 0;
  recon.stmtCredits.forEach(cr => {
    if (isPaymentCreditDesc(cr.desc) && isNeglectedPayer(cr.assignedPayer)) {
      billPaymentsTotal += cr.amount;
    }
    if (!cr.assignedPayer) {
      unresolvedCreditsCount++;
      unresolvedCreditsTotal += cr.amount;
    } else if (!isNeglectedPayer(cr.assignedPayer)) {
      addShare(cr.assignedPayer, cr.amount, "Refund");
    }
  });

  // 5b. Assignable Prior Cycle Carryover (added to assigned member's shares without double-counting in F6)
  const openingBalance = cardBal.openingBalance || 0;
  const carriedOver = openingBalance - billPaymentsTotal;
  const carryItem = recon.carryoverItem || null;
  const carryAssignedPayer = (carryItem && carryItem.assignedPayer && !isNeglectedPayer(carryItem.assignedPayer))
    ? carryItem.assignedPayer
    : "";
  let assignedCarryover = 0;
  if (carryAssignedPayer && Math.abs(carriedOver) >= 0.005) {
    if (carriedOver > 0) {
      addShare(carryAssignedPayer, carriedOver, "Unrecorded");
    } else {
      addShare(carryAssignedPayer, Math.abs(carriedOver), "Refund");
    }
    assignedCarryover = carriedOver;
  }
  const unassignedCarriedOver = carriedOver - assignedCarryover;

  let purchasesTotal = 0;
  let installmentsTotal = 0;
  let feesTotal = 0;
  let unrecordedTotal = 0;
  let refundsTotal = 0;
  let familySharesTotal = 0;

  Object.values(personShares).forEach(ps => {
    ps.netTotal = ps.purchases + ps.installments + ps.fees + ps.unrecorded - ps.refunds;
    purchasesTotal += ps.purchases;
    installmentsTotal += ps.installments;
    feesTotal += ps.fees;
    unrecordedTotal += ps.unrecorded;
    refundsTotal += ps.refunds;
    familySharesTotal += ps.netTotal;
  });

  const orphanList = recon.orphanLogs || recon.expandedUnmatched || [];
  let orphanActiveCount = 0;
  let orphanActiveTotal = 0;
  let orphanTotal = 0;
  orphanList.forEach(o => {
    const amt = parseFloat(o.amount) || 0;
    orphanTotal += amt;
    if (!o.isNeglected) {
      orphanActiveCount++;
      orphanActiveTotal += amt;
    }
  });

  const pendingList = recon.pendingNextItems || [];
  const pendingCount = pendingList.length;
  const pendingTotal = pendingList.reduce((s, p) => s + (parseFloat(p.amount) || 0), 0);

  const unresolvedCount = unassignedCount + unresolvedFeesCount + unresolvedCreditsCount + unresolvedInstCount;
  const unresolvedTotal = unassignedTotal + unresolvedFeesTotal - unresolvedCreditsTotal + unresolvedInstTotal;

  const closingBill = cardBal.billedBalance || 0;
  const auditDifference = (familySharesTotal + unassignedCarriedOver) - closingBill;
  const isExactMatch = Math.abs(auditDifference) < 0.50;

  return {
    activeCycleKey,
    activeCycleLabel,
    activeDueDate,
    closingBill,
    openingBalance,
    billPaymentsTotal,
    carriedOver,
    carryoverAssignedPayer: carryAssignedPayer,
    assignedCarryover,
    unassignedCarriedOver,
    purchasesTotal,
    installmentsTotal,
    feesTotal,
    unrecordedTotal,
    neglectedTotal,
    neglectedCount,
    unassignedTotal,
    unassignedCount,
    unresolvedFeesCount,
    unresolvedFeesTotal,
    unresolvedCreditsCount,
    unresolvedCreditsTotal,
    unresolvedInstCount,
    unresolvedInstTotal,
    unresolvedCount,
    unresolvedTotal,
    orphanActiveCount,
    orphanActiveTotal,
    orphanTotal,
    pendingCount,
    pendingTotal,
    refundsTotal,
    familySharesTotal,
    auditDifference,
    isExactMatch,
    personShares: allPayers.map(p => personShares[p] || {
      person: p,
      matchedPurchases: 0,
      unmatchedPurchases: 0,
      purchases: 0,
      installments: 0,
      fees: 0,
      unrecorded: 0,
      refunds: 0,
      netTotal: 0
    })
  };
}

// ==========================================
// 9. UI: ACTIVE BILL & RECONCILIATION
// ==========================================

function getReconTableBounds(ledger, recon) {
  const numPayers = ledger.personShares.length;
  const startShareRow = getActiveReconStartShareRow();
  const endShareRow = startShareRow + numPayers - 1;
  const totShareRow = endShareRow + 1;

  const creditsCount = recon.stmtCredits ? recon.stmtCredits.length : 0;
  const unlinkedInstCount = recon.unlinkedStmtInstallments ? recon.unlinkedStmtInstallments.length : 0;
  const numBankExceptions = (recon.stmtFees ? recon.stmtFees.length : 0) + creditsCount + unlinkedInstCount + (recon.missingInSheet ? recon.missingInSheet.length : 0);
  const hasCarryoverRow = Boolean(recon && recon.carryoverItem && recon.carryoverItem.showRow);
  const numExceptions = numBankExceptions + (hasCarryoverRow ? 1 : 0);
  const startExHeaderRow = totShareRow + 3;
  const startExRow = startExHeaderRow + 2;
  const endExRow = startExRow + Math.max(1, numExceptions) - 1;
  const bankExEndRow = numBankExceptions > 0 ? (startExRow + numBankExceptions - 1) : (startExRow - 1);
  const carryRow = hasCarryoverRow ? endExRow : -1;

  const numOrphans = (recon.orphanLogs || recon.expandedUnmatched || []).length;
  const hasOrphanTable = numOrphans > 0;
  const startOrphanHeaderRow = endExRow + 3;
  const startOrphanRow = startOrphanHeaderRow + 2;
  const endOrphanRow = startOrphanRow + Math.max(1, numOrphans) - 1;

  const numInstRecon = (recon.installmentReconItems || []).length;
  const hasTable4 = numInstRecon > 0;
  const prevEndBeforeInst = hasOrphanTable ? endOrphanRow : endExRow;
  const startInstHeaderRow = prevEndBeforeInst + 3;
  const startInstRow = startInstHeaderRow + 2;
  const endInstRow = startInstRow + Math.max(1, numInstRecon) - 1;

  const numPending = (recon.pendingNextItems || []).length;
  const hasPendingTable = false;
  const prevEndBeforePending = hasTable4 ? endInstRow : prevEndBeforeInst;
  const startPendingHeaderRow = prevEndBeforePending + 3;
  const startPendingRow = startPendingHeaderRow + 2;
  const endPendingRow = startPendingRow + Math.max(1, numPending) - 1;

  return {
    numPayers, startShareRow, endShareRow, totShareRow,
    numBankExceptions, numExceptions, startExHeaderRow, startExRow, bankExEndRow, endExRow,
    hasCarryoverRow, carryRow,
    numOrphans, hasOrphanTable, startOrphanHeaderRow, startOrphanRow, endOrphanRow,
    numUnmatched: numOrphans, hasTable3: hasOrphanTable,
    startTxHeaderRow: startOrphanHeaderRow, startTxRow: startOrphanRow, endTxRow: endOrphanRow,
    numInstRecon, hasTable4, startInstHeaderRow, startInstRow, endInstRow,
    numPending, hasPendingTable, startPendingHeaderRow, startPendingRow, endPendingRow
  };
}

function resetActiveReconRowsAndGroups_(sheet) {
  if (!sheet) return;
  const prevLastRow = Math.max(1, sheet.getLastRow());
  try { sheet.showRows(1, prevLastRow); } catch (e) { }
  let savedGroups = [];
  try {
    const rawMeta = String(sheet.getRange(1, 9).getValue() || "").trim();
    if (rawMeta.startsWith("[[")) {
      const parsed = JSON.parse(rawMeta);
      if (Array.isArray(parsed)) savedGroups = parsed;
    }
  } catch (e) { }
  savedGroups.forEach(pair => {
    if (Array.isArray(pair) && pair[0] >= 1 && pair[1] >= 1) {
      try { sheet.getRange(pair[0], 1, pair[1], 1).shiftRowGroupDepth(-1); } catch (e) { }
    }
  });
}

function applyCollapsibleRowGroups_(sheet, groupSpecs) {
  if (!sheet) return;
  const specs = (groupSpecs || []).filter(g => g && g.startRow >= 1 && g.numRows > 0);
  if (specs.length === 0) {
    try { sheet.getRange(1, 9).setValue(""); } catch (e) { }
    return;
  }
  try {
    if (
      typeof SpreadsheetApp !== "undefined" &&
      SpreadsheetApp.GroupControlTogglePosition &&
      typeof sheet.setRowGroupControlPosition === "function"
    ) {
      sheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
    }
  } catch (e) { }

  const saved = [];
  specs.forEach(g => {
    try {
      sheet.getRange(g.startRow, 1, g.numRows, 1).shiftRowGroupDepth(1);
      saved.push([g.startRow, g.numRows]);
      if (g.collapse && typeof sheet.getRowGroup === "function") {
        const grp = sheet.getRowGroup(g.startRow, 1);
        if (grp && typeof grp.collapse === "function") grp.collapse();
      }
    } catch (e) { }
  });
  try { sheet.getRange(1, 9).setValue(saved.length > 0 ? JSON.stringify(saved) : ""); } catch (e) { }
}

function renderActiveReconciliationSheet(ss, tz, ledger, recon) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.ACTIVE_RECON);
  resetActiveReconRowsAndGroups_(sheet);
  sheet.clear();
  sheet.setHiddenGridlines(false);

  // 1. Title Banner
  const isCycleLocked = isCycleClosed(ss, ledger.activeCycleKey);
  sheet.getRange("A1:H1").merge()
    .setValue("💳 NBE Credit Card — Active Bill & Precision Reconciliation (" + ledger.activeCycleLabel + ")" + (isCycleLocked ? " [🔒 Closed in History]" : ""))
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  sheet.getRange("A2:H2").merge()
    .setValue("Live reactive statement reconciliation. Edits update instantly via native formulas without screen refresh.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);

  // 2. Prepare Data & Dynamic Bounds for Declarative Equations
  const bounds = getReconTableBounds(ledger, recon);
  const sharePeople = ledger.personShares;
  const numPayers = bounds.numPayers;
  const startShareRow = bounds.startShareRow;
  const endShareRow = bounds.endShareRow;
  const totShareRow = bounds.totShareRow;

  // Build Unified Table 2 Needs Pairing & Statement Review Items
  const defaultFeePayer = getDefaultFeePayer(ss);
  const exceptionItems = [];
  recon.stmtFees.forEach((f, idx) => {
    exceptionItems.push({
      id: `F-${idx + 1}`,
      subjectId: f.stmtId || "",
      date: f.date || f.dateStr,
      desc: f.desc,
      amount: f.amount,
      category: "Bank Fee / Tax",
      assignedPayer: f.assignedPayer || defaultFeePayer,
      suggestedPairs: [],
      note: f.note || "Auto-assigned fee"
    });
  });

  recon.stmtCredits.forEach((c, idx) => {
    const isPay = isPaymentCreditDesc(c.desc);
    exceptionItems.push({
      id: `C-${idx + 1}`,
      subjectId: c.stmtId || "",
      date: c.date || c.dateStr,
      desc: c.desc,
      amount: c.amount,
      category: isPay ? "Bill Payment (Credit)" : "Refund / Credit",
      assignedPayer: c.assignedPayer || "",
      suggestedPairs: [],
      suggestedInstLinks: [],
      note: c.note || (isPay ? "Bill Payment / Settlement (Neglected by default)" : "Merchant Refund (Reduces Debt)")
    });
  });

  (recon.unlinkedStmtInstallments || []).forEach((u, idx) => {
    const sugs = u.suggestedInstLinks || [];
    const hasUnresolvedSugs = sugs.length > 0 && !u.assignedPayer;
    const topSug = sugs[0];
    exceptionItems.push({
      id: `I-${idx + 1}`,
      subjectId: u.stmtId || "",
      date: u.date || u.dateStr,
      desc: u.desc,
      amount: u.amount,
      category: hasUnresolvedSugs ? "Installment (💡 Link)" : "Installment (Unlinked)",
      assignedPayer: u.assignedPayer || "",
      suggestedPairs: [],
      suggestedInstLinks: sugs,
      note: hasUnresolvedSugs
        ? `💡 ${topSug.reason} → ${topSug.dropdownLabel}`
        : (u.note || "On bank statement — link to plan or assign payer")
    });
  });

  recon.missingInSheet.forEach((m, idx) => {
    const sugs = m.suggestedPairs || [];
    const hasUnresolvedSugs = sugs.length > 0 && !m.assignedPayer;
    const topSug = sugs[0];
    exceptionItems.push({
      id: `M-${idx + 1}`,
      subjectId: m.stmtId || "",
      date: m.date || m.dateStr,
      desc: m.desc,
      amount: m.amount,
      category: hasUnresolvedSugs ? "Missing Purchase (💡)" : "Missing Purchase",
      assignedPayer: m.assignedPayer || "",
      suggestedPairs: sugs,
      suggestedInstLinks: [],
      note: hasUnresolvedSugs
        ? `💡 ${topSug.reason} → ${topSug.dropdownLabel}`
        : (m.note || cleanMerchantName(m.desc))
    });
  });

  const numBankExceptions = bounds.numBankExceptions;
  const numExceptions = bounds.numExceptions;
  const startExHeaderRow = bounds.startExHeaderRow;
  const startExRow = bounds.startExRow;
  const bankExEndRow = bounds.bankExEndRow;
  const endExRow = bounds.endExRow;
  const hasCarryoverRow = bounds.hasCarryoverRow;
  const carryRow = bounds.carryRow;

  // Orphan Logs (in-period transactions without a paired bank line)
  const orphanList = recon.orphanLogs || recon.expandedUnmatched || [];
  const numOrphans = orphanList.length;
  const hasOrphanTable = bounds.hasOrphanTable;
  const startOrphanHeaderRow = bounds.startOrphanHeaderRow;
  const startOrphanRow = bounds.startOrphanRow;
  const endOrphanRow = bounds.endOrphanRow;

  // Table 4 Installments Reconciliation Items
  const instReconList = recon.installmentReconItems || [];
  const numInstRecon = instReconList.length;
  const hasTable4 = bounds.hasTable4;
  const startInstHeaderRow = bounds.startInstHeaderRow;
  const startInstRow = bounds.startInstRow;
  const endInstRow = bounds.endInstRow;

  // Pending Next Statement items remain computed in ledger/forecast without a 31-row table on Active Bill
  const groupSpecs = [];

  // 3. Unified Executive Overview & Live 1-Line Health Strip (Rows 4-8)
  const stmtDateStr = recon.cardBal.statementDate ? Utilities.formatDate(recon.cardBal.statementDate, tz, "yyyy-MM-dd") : "N/A";
  const dueDateStr = Utilities.formatDate(ledger.activeDueDate, tz, "MMMM d, yyyy");

  const healthModel = buildHealthCheckModel(ss, tz, ledger, recon, { bounds });
  const {
    t2UnresolvedCount,
    t2UnresolvedAmt,
    unresolvedCountExpr,
    unresolvedAmtExpr,
    negPayExpr,
    unassignedCarriedOverExpr
  } = healthModel.expressions;

  const badgeFormula = `=IF(AND(${unresolvedCountExpr}=0, ABS(F6)<0.5), "✅ Balanced (0 unresolved · 0.00 diff)", IF(${unresolvedCountExpr}>0, "⚠️ Unresolved: " & ${unresolvedCountExpr} & " line(s) (" & TEXT(${unresolvedAmtExpr},"#,##0.00") & " EGP)", IF(F6>0, "ℹ️ Statement Check: +" & TEXT(F6,"#,##0.00") & " EGP", "⚠️ Statement Check: -" & TEXT(ABS(F6),"#,##0.00") & " EGP")))`;

  const kpiRows = [
    ["Statement Date:", stmtDateStr, "Bank Closing Bill:", ledger.closingBill, "Recon Status:", badgeFormula],
    ["Payment Due Date:", dueDateStr, "Family Shares Total:", `=H${totShareRow}`, "Statement Check:", `=(D6 + ${unassignedCarriedOverExpr}) - D5`],
    ["Active Billing Cycle:", ledger.activeCycleLabel, "Total Card Utilized:", recon.cardBal.totalUtilized, "Available Limit Now:", recon.cardBal.availableBalanceNow]
  ];

  sheet.getRange(5, 1, 3, 6).setValues(kpiRows).setFontSize(9);
  sheet.getRange("A5:A7").setFontWeight("bold");
  sheet.getRange("C5:C7").setFontWeight("bold");
  sheet.getRange("E5:E7").setFontWeight("bold");
  sheet.getRange("A5:H7").setBackground(CONFIG.COLORS.PRIMARY_LIGHT).setBorder(true, true, true, true, true, true);
  const maxCols = sheet.getMaxColumns();
  if (maxCols > 9) {
    sheet.getRange(1, 10, Math.min(sheet.getMaxRows(), 30), maxCols - 9).clearFormat();
  }
  sheet.getRange("F5:H5").merge().setValue(badgeFormula).setFontWeight("bold").setFontSize(9);

  // Payment Verification & Status (Row 6 & Row 7, Cols G-H)
  const dateCell = sheet.getRange(6, 8);
  const existingVal = dateCell.getValue();
  const existingDate = parseDateValue(existingVal, tz);
  const decisionsStore = (recon && (recon.decisions || recon.reconStore)) ? (recon.decisions || recon.reconStore) : loadDecisions(ss);
  const cycleRecord = (decisionsStore.paymentRecords && decisionsStore.paymentRecords[ledger.activeCycleKey]) || null;

  let paymentDateVal = "";
  if (existingDate && !isNaN(existingDate.getTime())) {
    paymentDateVal = Utilities.formatDate(existingDate, tz, "yyyy-MM-dd");
  } else if (cycleRecord && cycleRecord.dateStr) {
    paymentDateVal = cycleRecord.dateStr;
  }
  const isPaid = !!paymentDateVal;

  sheet.getRange(6, 7).setValue("Payment Date:").setFontWeight("bold").setFontSize(9);
  if (paymentDateVal) {
    dateCell.setValue(paymentDateVal);
  }
  // NEVER CALL dateCell.clearContent()! If the cell has an existing entry or is empty, preserve it!
  dateCell.setNumberFormat("yyyy-MM-dd")
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setFontWeight(isPaid ? "bold" : "normal")
    .setFontColor(isPaid ? "#1b5e20" : "#333333")
    .setBackground(isPaid ? CONFIG.COLORS.PAID : "#fffde7");

  const dateValidation = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  dateCell.setDataValidation(dateValidation);

  sheet.getRange(7, 7).setValue("Payment Status:").setFontWeight("bold").setFontSize(9);
  const statusCell = sheet.getRange(7, 8);
  const statusFormula = `=IF(ISBLANK(H6), "⏳ Pending Payment", "✅ Paid on " & TEXT(H6, "yyyy-MM-dd"))`;
  statusCell.setFormula(statusFormula)
    .setFontSize(9)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(isPaid ? CONFIG.COLORS.PAID : CONFIG.COLORS.DUE_SOON)
    .setFontColor(isPaid ? "#1b5e20" : "#7f6000");

  sheet.getRange("D5:D7").setNumberFormat("#,##0.00");
  sheet.getRange("F6:F7").setNumberFormat("#,##0.00");

  // Live Health Check Panel (Rows 9-11)
  renderHealthCheckPanel(sheet, healthModel);

  // 4. Family Member Shares Table (Table 1 - Fully Formula-Driven)
  sheet.getRange(startShareRow - 2, 1, 1, 8).merge()
    .setValue("👥 Family Member Active Debt Shares (Billed on Current Statement)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  const sharesHeaders = ["#", "Family Member", "Purchases (EGP)", "Installments (EGP)", "Bank Fees & Taxes (EGP)", "Unrecorded Assigned (EGP)", "Refunds / Credits (EGP)", "Net Total Due (EGP)"];
  sheet.getRange(startShareRow - 1, 1, 1, sharesHeaders.length).setValues([sharesHeaders])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER);

  if (numPayers > 0) {
    const shareRows = sharePeople.map((ps, idx) => {
      const r = startShareRow + idx;
      const purchFormula = `=${ps.matchedPurchases.toFixed(2)}`;
      let instFormula;
      if (hasTable4 && numBankExceptions > 0) {
        instFormula = `=SUMIFS($D$${startInstRow}:$D$${endInstRow}, $F$${startInstRow}:$F$${endInstRow}, "*" & TRIM($B${r}) & "*", $F$${startInstRow}:$F$${endInstRow}, "<>🚫*") + SUMIFS($D$${startExRow}:$D$${bankExEndRow}, $F$${startExRow}:$F$${bankExEndRow}, $B${r}, $E$${startExRow}:$E$${bankExEndRow}, "Installment*")`;
      } else if (hasTable4) {
        instFormula = `=SUMIFS($D$${startInstRow}:$D$${endInstRow}, $F$${startInstRow}:$F$${endInstRow}, "*" & TRIM($B${r}) & "*", $F$${startInstRow}:$F$${endInstRow}, "<>🚫*")`;
      } else if (numBankExceptions > 0) {
        instFormula = `=SUMIFS($D$${startExRow}:$D$${bankExEndRow}, $F$${startExRow}:$F$${bankExEndRow}, $B${r}, $E$${startExRow}:$E$${bankExEndRow}, "Installment*")`;
      } else {
        instFormula = `=${ps.installments.toFixed(2)}`;
      }
      const feesFormula = numBankExceptions > 0
        ? `=SUMIFS($D$${startExRow}:$D$${bankExEndRow}, $F$${startExRow}:$F$${bankExEndRow}, $B${r}, $E$${startExRow}:$E$${bankExEndRow}, "Bank Fee / Tax")`
        : "=0";

      let posCarryTerm = "";
      let negCarryTerm = "";
      if (hasCarryoverRow && carryRow > 0) {
        const carryNormPayer = `SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM($F$${carryRow}), " + ", "+"), ", ", "+"), ",", "+")`;
        const carryMemberMatch = `ISNUMBER(SEARCH("+" & TRIM($B${r}) & "+", "+" & ${carryNormPayer} & "+"))`;
        const carrySplitDenom = `(LEN(${carryNormPayer}) - LEN(SUBSTITUTE(${carryNormPayer}, "+", "")) + 1)`;
        posCarryTerm = `IF(AND($D$${carryRow}>0, ISERROR(SEARCH("🚫", $F$${carryRow})), ${carryMemberMatch}), $D$${carryRow} / ${carrySplitDenom}, 0)`;
        negCarryTerm = `IF(AND($D$${carryRow}<0, ISERROR(SEARCH("🚫", $F$${carryRow})), ${carryMemberMatch}), ABS($D$${carryRow}) / ${carrySplitDenom}, 0)`;
      }

      const unrecBase = numBankExceptions > 0
        ? `SUMIFS($D$${startExRow}:$D$${bankExEndRow}, $F$${startExRow}:$F$${bankExEndRow}, $B${r}, $E$${startExRow}:$E$${bankExEndRow}, "Missing Purchase*")`
        : "";
      const unrecFormula = unrecBase && posCarryTerm
        ? `=${unrecBase} + ${posCarryTerm}`
        : (unrecBase ? `=${unrecBase}` : (posCarryTerm ? `=${posCarryTerm}` : "=0"));

      const refBase = numBankExceptions > 0
        ? `SUMIFS($D$${startExRow}:$D$${bankExEndRow}, $F$${startExRow}:$F$${bankExEndRow}, $B${r}, $E$${startExRow}:$E$${bankExEndRow}, "*Credit*")`
        : "";
      const refFormula = refBase && negCarryTerm
        ? `=${refBase} + ${negCarryTerm}`
        : (refBase ? `=${refBase}` : (negCarryTerm ? `=${negCarryTerm}` : "=0"));

      const netFormula = `=C${r}+D${r}+E${r}+F${r}-G${r}`;

      return [
        idx + 1,
        ps.person,
        purchFormula,
        instFormula,
        feesFormula,
        unrecFormula,
        refFormula,
        netFormula
      ];
    });

    sheet.getRange(startShareRow, 1, shareRows.length, sharesHeaders.length).setValues(shareRows).setFontSize(9);
    sheet.getRange(startShareRow, 3, shareRows.length, 5).setNumberFormat('#,##0.00;-#,##0.00;"—"');
    sheet.getRange(startShareRow, 8, shareRows.length, 1).setNumberFormat("#,##0.00").setFontWeight("bold");

    sharePeople.forEach((ps, idx) => {
      const hasActivity = Math.abs(ps.purchases || 0) >= 0.005 ||
        Math.abs(ps.installments || 0) >= 0.005 ||
        Math.abs(ps.fees || 0) >= 0.005 ||
        Math.abs(ps.unrecorded || 0) >= 0.005 ||
        Math.abs(ps.refunds || 0) >= 0.005 ||
        Math.abs(ps.netTotal || 0) >= 0.005;
      if (!hasActivity) {
        try { sheet.hideRows(startShareRow + idx); } catch (e) { }
      }
    });
  }

  // Summary Row (Formula-Driven)
  const totalRow = [
    "TOTAL",
    "All Family Members",
    `=SUM(C${startShareRow}:C${endShareRow})`,
    `=SUM(D${startShareRow}:D${endShareRow})`,
    `=SUM(E${startShareRow}:E${endShareRow})`,
    `=SUM(F${startShareRow}:F${endShareRow})`,
    `=SUM(G${startShareRow}:G${endShareRow})`,
    `=SUM(H${startShareRow}:H${endShareRow})`
  ];
  sheet.getRange(totShareRow, 1, 1, sharesHeaders.length).setValues([totalRow])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT)
    .setBorder(true, true, true, true, true, true);
  sheet.getRange(totShareRow, 3, 1, 6).setNumberFormat("#,##0.00");

  // 5. Unified Needs Pairing & Statement Review Table (Table 2 — Unresolved first, resolved collapsed, Carryover row below)
  const unresolvedEx = exceptionItems.filter(ex => !String(ex.assignedPayer || "").trim());
  const resolvedEx = exceptionItems.filter(ex => Boolean(String(ex.assignedPayer || "").trim()));
  const orderedBankExceptions = [...unresolvedEx, ...resolvedEx];
  const orderedExceptions = orderedBankExceptions.slice();
  const carryEx = hasCarryoverRow ? Object.assign({}, recon.carryoverItem, { isCarryover: true }) : null;
  if (carryEx) {
    orderedExceptions.push(carryEx);
  }

  const t2BannerRange = sheet.getRange(startExHeaderRow, 1, 1, 8).merge()
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(unresolvedEx.length === 0 ? CONFIG.COLORS.PRIMARY_LIGHT : CONFIG.COLORS.DUE_SOON);
  if (numBankExceptions > 0) {
    const t2SuggestedCount = `COUNTIFS(E${startExRow}:E${bankExEndRow}, "*💡*", F${startExRow}:F${bankExEndRow}, "")`;
    const carryBannerSuffix = hasCarryoverRow ? ' & " + Carryover"' : "";
    t2BannerRange.setFormula(`="🔗 Needs Pairing & Statement Review (" & ${numBankExceptions} & " Items | Unresolved: " & ${t2UnresolvedCount} & " = " & TEXT(${t2UnresolvedAmt}, "#,##0.00") & " EGP | Suggested pairs: " & ${t2SuggestedCount}${carryBannerSuffix} & ")"`);
  } else if (hasCarryoverRow) {
    t2BannerRange.setValue("🔗 Needs Pairing & Statement Review (0 Bank Items · Prior Cycle Carryover)");
  } else {
    t2BannerRange.setValue("🔗 Needs Pairing & Statement Review (0 Items — Statement 100% Reconciled)");
  }

  const exHeaders = ["#", "Date", "Description / Merchant", "Amount (EGP)", "Category", "Resolve (Pair / Assign)", "Suggestion / Note", "Status", "_Subject_ID"];
  sheet.getRange(startExHeaderRow + 1, 1, 1, 8).setFontWeight("bold").setFontSize(9).setBackground(CONFIG.COLORS.HEADER);
  sheet.getRange(startExHeaderRow + 1, 1, 1, exHeaders.length).setValues([exHeaders]);

  if (numExceptions > 0) {
    const exRows = orderedExceptions.map((ex, idx) => {
      const r = startExRow + idx;
      if (ex.isCarryover) {
        const carryAmtFormula = numBankExceptions > 0
          ? `=${(ledger.openingBalance || 0).toFixed(2)} - ${negPayExpr}`
          : `=${(ledger.openingBalance || 0).toFixed(2)}`;
        const statusFormula = `=IF(OR(ISBLANK(F${r}), ISNUMBER(SEARCH("🚫", F${r}))), "ℹ️ Unassigned (Prior Cycle)", "✅ Assigned to " & F${r})`;
        return [
          ex.id,
          ex.date,
          ex.desc,
          carryAmtFormula,
          ex.category,
          ex.assignedPayer,
          ex.note,
          statusFormula,
          ex.subjectId || ""
        ];
      }
      const statusFormula = `=IF(ISBLANK(F${r}), IF(ISNUMBER(SEARCH("💡", E${r})), "💡 Confirm Link/Pair or Assign", "⚠️ Needs Action"), IF(ISNUMBER(SEARCH("🚫", F${r})), "🚫 Neglected", IF(ISNUMBER(SEARCH("🔗", F${r})), "🔗 Linking...", "✅ Assigned")))`;
      return [
        ex.id,
        ex.date,
        ex.desc,
        ex.amount,
        ex.category,
        ex.assignedPayer,
        ex.note,
        statusFormula,
        ex.subjectId || ""
      ];
    });

    const exRange = sheet.getRange(startExRow, 1, exRows.length, exHeaders.length);
    exRange.setValues(exRows);
    sheet.getRange(startExRow, 1, exRows.length, 8).setFontSize(9);
    sheet.getRange(startExRow, 2, exRows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(startExRow, 4, exRows.length, 1).setNumberFormat("#,##0.00");
    if (hasCarryoverRow && carryRow > 0) {
      sheet.getRange(carryRow, 4, 1, 1).setNumberFormat("+#,##0.00;-#,##0.00;0.00");
    }

    const dynamicPayers = getAllUniquePayers(ss);
    const standardOptions = [...dynamicPayers, "🚫 Neglect / Ignore"];
    const defaultValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(standardOptions, true)
      .setAllowInvalid(true)
      .build();

    const activeOrphanPool = orphanList.filter(o => !o.isNeglected && o.txId);
    const seenOrphanTx = new Set();
    const uniqueActiveOrphans = [];
    activeOrphanPool.forEach(o => {
      if (!seenOrphanTx.has(o.txId)) {
        seenOrphanTx.add(o.txId);
        uniqueActiveOrphans.push(o);
      }
    });

    const activeUnmatchedSheetInsts = (recon.unmatchedSheetInstallments || []).filter(sh => sh.instId);
    const exValidations = [];
    const exBgs = [];

    for (let i = 0; i < orderedExceptions.length; i++) {
      const ex = orderedExceptions[i];
      let rowVal = defaultValidation;

      if (ex.isCarryover) {
        rowVal = SpreadsheetApp.newDataValidation()
          .requireValueInList(["🚫 Keep Unassigned (Prior Cycle)", ...dynamicPayers], true)
          .setAllowInvalid(true)
          .build();
      } else if (ex.suggestedInstLinks && ex.suggestedInstLinks.length > 0) {
        const sugLabels = ex.suggestedInstLinks.map(sl => sl.dropdownLabel);
        rowVal = SpreadsheetApp.newDataValidation()
          .requireValueInList([...sugLabels, ...standardOptions], true)
          .setAllowInvalid(true)
          .build();
      } else if (/^I-/.test(ex.id) && !ex.assignedPayer && activeUnmatchedSheetInsts.length > 0) {
        const closest = activeUnmatchedSheetInsts
          .slice()
          .sort((a, b) => Math.abs(a.emi - ex.amount) - Math.abs(b.emi - ex.amount))
          .slice(0, 3)
          .map(sh => formatInstLinkDropdownLabel([sh]));
        if (closest.length > 0) {
          rowVal = SpreadsheetApp.newDataValidation()
            .requireValueInList([...standardOptions, ...closest], true)
            .setAllowInvalid(true)
            .build();
        }
      } else if (ex.suggestedPairs && ex.suggestedPairs.length > 0) {
        const sugLabels = ex.suggestedPairs.map(sp => sp.dropdownLabel);
        rowVal = SpreadsheetApp.newDataValidation()
          .requireValueInList([...sugLabels, ...standardOptions], true)
          .setAllowInvalid(true)
          .build();
      } else if (/^M-/.test(ex.id) && !ex.assignedPayer && uniqueActiveOrphans.length > 0) {
        const closest = uniqueActiveOrphans
          .slice()
          .sort((a, b) => Math.abs(a.amount - ex.amount) - Math.abs(b.amount - ex.amount))
          .slice(0, 3)
          .map(o => formatPairDropdownLabel([o]));
        if (closest.length > 0) {
          rowVal = SpreadsheetApp.newDataValidation()
            .requireValueInList([...standardOptions, ...closest], true)
            .setAllowInvalid(true)
            .build();
        }
      }
      exValidations.push([rowVal]);

      const cat = ex.category;
      const payer = ex.assignedPayer;
      const hasSugs = (ex.suggestedPairs && ex.suggestedPairs.length > 0) || (ex.suggestedInstLinks && ex.suggestedInstLinks.length > 0);
      let bg = "#ffffff";
      if (ex.isCarryover) {
        bg = (isNeglectedPayer(payer) || !payer) ? "#f3f4f6" : CONFIG.COLORS.PAID;
      } else if (isNeglectedPayer(payer)) {
        bg = "#eeeeee";
      } else if (cat.includes("Credit")) {
        bg = CONFIG.COLORS.PRIMARY_LIGHT;
      } else if (payer) {
        bg = CONFIG.COLORS.PAID;
      } else if (hasSugs) {
        bg = "#e8f0fe";
      } else if (/^I-/.test(ex.id)) {
        bg = CONFIG.COLORS.DUE_SOON;
      }
      exBgs.push([bg, bg, bg, bg, bg, bg, bg, bg]);
    }

    sheet.getRange(startExRow, 6, exRows.length, 1).setDataValidations(exValidations);
    sheet.getRange(startExRow, 1, exRows.length, 8).setBackgrounds(exBgs);
    if (resolvedEx.length > 0) {
      groupSpecs.push({
        startRow: startExRow + unresolvedEx.length,
        numRows: resolvedEx.length,
        collapse: true
      });
    }
  }

  // 6. Orphan Logs Panel (In-period transactions without a paired bank line — Active first, dismissed collapsed)
  if (hasOrphanTable) {
    const activeOrphans = orphanList.filter(o => !o.isNeglected);
    const dismissedOrphans = orphanList.filter(o => Boolean(o.isNeglected));
    const orderedOrphans = [...activeOrphans, ...dismissedOrphans];

    sheet.getRange(startOrphanHeaderRow, 1, 1, 8).merge()
      .setFormula(`="🔎 Orphan Logs — In-Period Transactions Without a Bank Line (" & ${numOrphans} & " Items | Unresolved: " & COUNTIF(F${startOrphanRow}:F${endOrphanRow}, "<>🚫*") & " = " & TEXT(SUMIFS(D${startOrphanRow}:D${endOrphanRow}, F${startOrphanRow}:F${endOrphanRow}, "<>🚫*"), "#,##0.00") & " EGP — NOT billed until paired above)"`)
      .setFontWeight("bold")
      .setFontSize(10)
      .setBackground(activeOrphans.length === 0 ? CONFIG.COLORS.PRIMARY_LIGHT : "#e8eaed");

    const orphanHeaders = ["#", "Date", "Description", "Amount (EGP)", "Person", "Resolution", "Pairing Hint / Note", "Status", "_Subject_ID"];
    sheet.getRange(startOrphanHeaderRow + 1, 1, 1, 8).setFontWeight("bold").setFontSize(9).setBackground(CONFIG.COLORS.HEADER);
    sheet.getRange(startOrphanHeaderRow + 1, 1, 1, orphanHeaders.length).setValues([orphanHeaders]);

    const orphanRows = orderedOrphans.map((tx, idx) => {
      const r = startOrphanRow + idx;
      const statusFormula = `=IF(ISNUMBER(SEARCH("🚫", F${r})), "🚫 Dismissed", IF(ISNUMBER(SEARCH("🔗", F${r})), "🔗 Pairing...", "🔎 Unpaired (Not Billed)"))`;
      const defaultAction = tx.isNeglected ? "🚫 Dismiss / Duplicate Log" : "⏳ Not on Statement Yet (Ignored)";
      const topHint = (tx.suggestedStmts && tx.suggestedStmts.length > 0) ? tx.suggestedStmts[0] : null;
      const hintNote = topHint
        ? (topHint.isPartialAppend
          ? `💡 ${topHint.reason} → ${topHint.stmtId} (rem ${Number(topHint.remainingAmount || 0).toFixed(2)})`
          : `💡 ${topHint.reason} → ${topHint.stmtId} (${topHint.stmt.amount.toFixed(2)})`)
        : "Not found on current statement";
      return [
        tx.id,
        tx.date,
        tx.desc,
        tx.amount,
        tx.person,
        defaultAction,
        hintNote,
        statusFormula,
        tx.subjectId || ""
      ];
    });

    sheet.getRange(startOrphanRow, 1, orphanRows.length, orphanHeaders.length).setValues(orphanRows);
    sheet.getRange(startOrphanRow, 1, orphanRows.length, 8).setFontSize(9);
    sheet.getRange(startOrphanRow, 2, orphanRows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(startOrphanRow, 4, orphanRows.length, 1).setNumberFormat("#,##0.00");

    const partialStmtDebits = (recon.stmtDebits || []).filter(st => st.matched && st.remainingAmount > 0.10 && st.stmtId);
    const unassignedStmtDebits = (recon.missingInSheet || []).filter(m => !m.assignedPayer);
    const orphanValidations = [];
    const orphanBgs = [];
    for (let i = 0; i < orderedOrphans.length; i++) {
      const orph = orderedOrphans[i];
      const sugLabels = (orph.suggestedStmts || []).map(s => s.dropdownLabel);
      if (sugLabels.length === 0) {
        partialStmtDebits
          .slice()
          .sort((a, b) => Math.abs(a.remainingAmount - orph.amount) - Math.abs(b.remainingAmount - orph.amount))
          .slice(0, 2)
          .forEach(st => sugLabels.push(formatOrphanAddDropdownLabel(st, st.remainingAmount)));
        unassignedStmtDebits
          .slice()
          .sort((a, b) => Math.abs(a.amount - orph.amount) - Math.abs(b.amount - orph.amount))
          .slice(0, 3)
          .forEach(st => sugLabels.push(formatOrphanPairDropdownLabel(st)));
      }
      const orphanValidation = SpreadsheetApp.newDataValidation()
        .requireValueInList([...sugLabels, "⏳ Not on Statement Yet (Ignored)", "🚫 Dismiss / Duplicate Log"], true)
        .setAllowInvalid(true)
        .build();
      orphanValidations.push([orphanValidation]);

      let bg = "#fff8e1";
      if (orph.isNeglected) {
        bg = "#eeeeee";
      } else if (orph.suggestedStmts && orph.suggestedStmts.length > 0) {
        bg = "#e8f0fe";
      }
      orphanBgs.push([bg, bg, bg, bg, bg, bg, bg, bg]);
    }
    sheet.getRange(startOrphanRow, 6, orphanRows.length, 1).setDataValidations(orphanValidations);
    sheet.getRange(startOrphanRow, 1, orphanRows.length, 8).setBackgrounds(orphanBgs);
    if (dismissedOrphans.length > 0) {
      groupSpecs.push({
        startRow: startOrphanRow + activeOrphans.length,
        numRows: dismissedOrphans.length,
        collapse: true
      });
    }
  }

  // 7. Active Installments & Statement EMI Reconciliation (Table 4)
  if (hasTable4) {
    sheet.getRange(startInstHeaderRow, 1, 1, 8).merge()
      .setValue(`📅 Active Installments & Statement EMI Reconciliation (${numInstRecon} Items — Live Formula Linked to Table 1)`)
      .setFontWeight("bold")
      .setFontSize(10)
      .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

    const instHeaders = ["#", "Date", "Installment Description", "Monthly EMI (EGP)", "Match Source", "Billing Action / Payer", "Custom Note", "Status", "_Subject_ID"];
    sheet.getRange(startInstHeaderRow + 1, 1, 1, 8).setFontWeight("bold").setFontSize(9).setBackground(CONFIG.COLORS.HEADER);
    sheet.getRange(startInstHeaderRow + 1, 1, 1, instHeaders.length).setValues([instHeaders]);

    const instRows = instReconList.map((inst, idx) => {
      const r = startInstRow + idx;
      const statusFormula = `=IF(ISBLANK(F${r}), "⚠️ Needs Payer", IF(ISNUMBER(SEARCH("🚫", F${r})), "🚫 Skipped This Cycle", IF(ISNUMBER(SEARCH("✅", F${r})), F${r}, "✅ Billed to " & F${r})))`;
      return [
        inst.id,
        inst.date,
        inst.desc,
        inst.amount,
        inst.sourceCategory,
        inst.action,
        inst.note,
        statusFormula,
        inst.subjectId || ""
      ];
    });

    sheet.getRange(startInstRow, 1, instRows.length, instHeaders.length).setValues(instRows);
    sheet.getRange(startInstRow, 1, instRows.length, 8).setFontSize(9);
    sheet.getRange(startInstRow, 2, instRows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(startInstRow, 4, instRows.length, 1).setNumberFormat("#,##0.00");

    const instPayersList = getAllUniquePayers(ss);
    const instValCache = {};
    const instValidations = [];
    const instBgs = [];
    for (let i = 0; i < instReconList.length; i++) {
      const inst = instReconList[i];
      const defKey = inst.defaultPayer || "";
      if (!instValCache[defKey]) {
        const options = [];
        if (inst.defaultPayer) {
          options.push(`✅ Billed to ${inst.defaultPayer}`);
        }
        instPayersList.forEach(p => {
          const opt = `✅ Billed to ${p}`;
          if (!options.includes(opt)) options.push(opt);
        });
        options.push("🚫 Neglect / Skip This Cycle");
        options.push("🚫 Not on Statement (Skip)");

        instValCache[defKey] = SpreadsheetApp.newDataValidation()
          .requireValueInList(options, true)
          .setAllowInvalid(true)
          .build();
      }
      instValidations.push([instValCache[defKey]]);

      let bg = "#f8f9fa";
      if (!inst.action) {
        bg = CONFIG.COLORS.DUE_SOON;
      } else if (isNeglectedPayer(inst.action)) {
        bg = "#eeeeee";
      } else if (inst.sourceCategory && inst.sourceCategory.includes("Statement Only")) {
        bg = CONFIG.COLORS.PAID;
      }
      instBgs.push([bg, bg, bg, bg, bg, bg, bg, bg]);
    }
    sheet.getRange(startInstRow, 6, instRows.length, 1).setDataValidations(instValidations);
    sheet.getRange(startInstRow, 1, instRows.length, 8).setBackgrounds(instBgs);

    const hasUnresolvedInst = instReconList.some(it => !String(it.action || "").trim());
    groupSpecs.push({
      startRow: startInstRow,
      numRows: instRows.length,
      collapse: !hasUnresolvedInst
    });
  }

  try { sheet.hideColumns(9); } catch (e) { }
  try {
    renderInstallmentProgressFormulas(ss, bounds, {
      decisions: recon.decisions,
      progressMap: recon.instProgressMap,
      activeCycleKey: ledger.activeCycleKey,
      tz
    });
  } catch (e) { }

  applySheetLayoutFitting(sheet, {
    bannerRow: 1,
    bannerHeight: 30,
    customRowHeights: { 2: 20, 3: 10, 4: 24, 8: 22, 9: 10, 10: 26, 11: 24 },
    explicitColWidths: {
      1: 55,  // #
      2: 125, // Date / Member
      3: 210, // Description / Category
      4: 105, // Amount / Purchases
      5: 115, // Installments / Category
      6: 145, // Resolve (Pair / Assign)
      7: 145, // Suggestion / Note
      8: 155  // Status
    }
  });

  applyCollapsibleRowGroups_(sheet, groupSpecs);

  return healthModel;
}

// ==========================================
// 10. UI: MONTHLY HISTORY & FORECAST
// ==========================================

// Immutable closed-statement archive & Close Statement workflow live in src/History.js

function renderMonthlyHistoryForecastSheet(ss, tz, ledger) {
  migrateHistoryFromForecastSheet(ss, tz);

  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.HISTORY_FORECAST);
  sheet.clear();
  sheet.setHiddenGridlines(false);

  const history = readHistoryArchive(ss);
  const histKeys = Object.keys(history).sort();

  // Future Forecasts (Next 4 cycles based on Installments sheet)
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  const futureForecasts = {};
  if (instSheet && instSheet.getLastRow() >= 2) {
    const instData = instSheet.getRange(2, 1, instSheet.getLastRow() - 1, 9).getValues();
    instData.forEach(row => {
      const pDate = parseDateValue(row[1], tz);
      const duration = parseInt(row[4], 10);
      const emi = parseFloat(row[7]);
      const payer = String(row[8] || "").trim();
      if (pDate && !isNaN(duration) && !isNaN(emi) && emi > 0 && payer) {
        const people = splitPayerNames(payer);
        if (people.length === 0) return;
        const splitEmi = emi / people.length;

        for (let i = 0; i < duration; i++) {
          const cyc = getBillingCycleForInstallment(pDate, i, tz);
          if (cyc && cyc.cycleSortKey > ledger.activeCycleKey) {
            if (!futureForecasts[cyc.cycleSortKey]) {
              futureForecasts[cyc.cycleSortKey] = { label: cyc.cycleLabel, total: 0, people: {} };
            }
            people.forEach(p => {
              futureForecasts[cyc.cycleSortKey].people[p] = (futureForecasts[cyc.cycleSortKey].people[p] || 0) + splitEmi;
              futureForecasts[cyc.cycleSortKey].total += splitEmi;
            });
          }
        }
      }
    });
  }

  const futureKeys = Object.keys(futureForecasts).sort().slice(0, 4);

  // Build union of active payers + historical members, hiding all-zero inactive members
  const payerSet = new Set(getAllUniquePayers(ss));
  histKeys.forEach(k => {
    Object.keys((history[k] && history[k].people) || {}).forEach(p => {
      if (p) payerSet.add(p);
    });
  });
  const allPayers = Array.from(payerSet).sort((a, b) => a.localeCompare(b));

  const futureTxPayers = new Set();
  const futureRanges = futureKeys.map(k => getCycleDateRange(k, tz)).filter(Boolean);
  if (futureRanges.length > 0) {
    const txSheetHist = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
    if (txSheetHist && txSheetHist.getLastRow() >= 2) {
      const txHistData = txSheetHist.getRange(2, 1, txSheetHist.getLastRow() - 1, 5).getValues();
      txHistData.forEach(row => {
        const pDate = parseDateValue(row[1], tz);
        const amt = parseFloat(row[4]);
        const payer = String(row[3] || "").trim();
        if (!pDate || isNaN(amt) || amt <= 0 || !payer) return;
        const inFuture = futureRanges.some(dr => pDate >= dr.start && pDate <= dr.end);
        if (inFuture) {
          splitPayerNames(payer).forEach(p => { if (p) futureTxPayers.add(p); });
        }
      });
    }
  }

  const activeNonZeroMap = {};
  (ledger.personShares || []).forEach(ps => {
    activeNonZeroMap[ps.person] = Math.abs(ps.purchases || 0) >= 0.005 ||
      Math.abs(ps.installments || 0) >= 0.005 ||
      Math.abs(ps.fees || 0) >= 0.005 ||
      Math.abs(ps.unrecorded || 0) >= 0.005 ||
      Math.abs(ps.refunds || 0) >= 0.005 ||
      Math.abs(ps.netTotal || 0) >= 0.005;
  });

  const filteredPayers = allPayers.filter(p => {
    if (activeNonZeroMap[p]) return true;
    if (histKeys.some(k => Math.abs((history[k] && history[k].people && history[k].people[p]) || 0) >= 0.005)) return true;
    if (futureKeys.some(k => Math.abs((futureForecasts[k] && futureForecasts[k].people && futureForecasts[k].people[p]) || 0) >= 0.005)) return true;
    if (futureTxPayers.has(p)) return true;
    return false;
  });
  const visiblePayers = filteredPayers.length > 0 ? filteredPayers : allPayers;

  const activeAlreadyClosed = Boolean(history[ledger.activeCycleKey]);
  const headers = [
    "Family Member",
    ...histKeys.map(k => `${history[k].label}\n(${history[k].status || "Closed"})`),
    ...(activeAlreadyClosed ? [] : [`${ledger.activeCycleLabel}\n(Active)`]),
    ...futureKeys.map(k => `${futureForecasts[k].label}\n(Forecast)`)
  ];

  const bannerEndCol = getColumnLetter(Math.max(headers.length, 4));
  sheet.getRange(`A1:${bannerEndCol}1`).merge()
    .setValue("📅 NBE Credit Card — Monthly History & Multi-Month Forecast Horizon")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  sheet.getRange(`A2:${bannerEndCol}2`).merge()
    .setValue("Closed billing cycles alongside active statement debt and future projected installment commitments.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);

  sheet.getRange(4, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER)
    .setWrap(true);

  const startDataRow = 5;
  const lastDataRow = startDataRow + Math.max(1, visiblePayers.length) - 1;
  const startShareRow = getActiveReconStartShareRow();
  const activeShareCount = (ledger.personShares && ledger.personShares.length) || allPayers.length;
  const endShareRow = startShareRow + Math.max(1, activeShareCount) - 1;

  const rows = visiblePayers.map((p, idx) => {
    const r = startDataRow + idx;
    const row = [p];

    // 1. Past Closed Cycles (Frozen in History)
    histKeys.forEach(k => row.push(history[k].people[p] || 0));

    // 2. Active Month (Formula linked to Active Bill & Reconciliation Table 1 Net Total, omitted if already closed)
    if (!activeAlreadyClosed) {
      row.push(`=SUMIFS('${CONFIG.SHEETS.ACTIVE_RECON}'!$H$${startShareRow}:$H$${endShareRow}, '${CONFIG.SHEETS.ACTIVE_RECON}'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($A${r}) & "*")`);
    }

    // 3. Multi-Month Future Forecasts (Installment Base + Live SUMIFS from Transactions)
    futureKeys.forEach(k => {
      const instAmt = (futureForecasts[k] && futureForecasts[k].people[p]) || 0;
      const dateRange = getCycleDateRange(k, tz);
      if (dateRange) {
        const purchSum = makePurchaseSumFormula(`$A${r}`, dateRange);
        if (instAmt > 0) {
          row.push(`=${instAmt.toFixed(2)} + ${purchSum}`);
        } else {
          row.push(`=${purchSum}`);
        }
      } else {
        row.push(instAmt > 0 ? instAmt : 0);
      }
    });

    return row;
  });

  // Equation-based Total Row
  const totalRow = ["TOTAL BILL (EGP)"];
  for (let c = 2; c <= headers.length; c++) {
    const colLetter = getColumnLetter(c);
    totalRow.push(`=SUM(${colLetter}${startDataRow}:${colLetter}${lastDataRow})`);
  }

  if (rows.length > 0) {
    sheet.getRange(5, 1, rows.length, headers.length).setValues(rows).setFontSize(9);
  }
  sheet.getRange(5 + rows.length, 1, 1, headers.length).setValues([totalRow])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

  if (headers.length > 1) {
    if (rows.length > 0) {
      sheet.getRange(5, 2, rows.length, headers.length - 1).setNumberFormat('#,##0.00;-#,##0.00;"—"');
    }
    sheet.getRange(5 + rows.length, 2, 1, headers.length - 1).setNumberFormat("#,##0.00");
  }

  try { sheet.setFrozenRows(4); } catch (e) { }

  applySheetLayoutFitting(sheet, {
    bannerRow: 1,
    bannerHeight: 30,
    customRowHeights: { 2: 20, 4: 32 },
    explicitColWidths: { 1: 140 },
    autoResizeStartCol: 2,
    autoResizeColCount: Math.max(0, headers.length - 1),
    colWidthBounds: { default: { min: 100, max: 125 } }
  });
}

function shouldRedrawMonthlyHistory_(ss, ledger) {
  const sheet = ss && ss.getSheetByName(CONFIG.SHEETS.HISTORY_FORECAST);
  if (!sheet || sheet.getLastRow() < 5 || sheet.getLastColumn() < 2) return true;
  const lastCol = sheet.getLastColumn();
  const headerVals = sheet.getRange(4, 1, 1, lastCol).getValues()[0].map(h => String(h || ""));
  const closedCycle = typeof isCycleClosed === "function" ? isCycleClosed(ss, ledger.activeCycleKey) : false;

  if (closedCycle) {
    const hasClosedCol = headerVals.some(h => h.includes(ledger.activeCycleLabel) && h.includes("(Closed)"));
    const hasDupActive = headerVals.some(h => h.includes("(Active)"));
    if (!hasClosedCol || hasDupActive) return true;
    return false;
  }

  const activeColIdx = headerVals.findIndex(h => h.includes(ledger.activeCycleLabel) && h.includes("(Active)")) + 1;
  if (activeColIdx <= 1) return true;

  const startShareRow = getActiveReconStartShareRow();
  const numPayers = (ledger.personShares && ledger.personShares.length) || 0;
  if (numPayers === 0) return true;
  const endShareRow = startShareRow + numPayers - 1;
  const expectedRange = `$H$${startShareRow}:$H$${endShareRow}`;

  const sampleFormula = String(sheet.getRange(5, activeColIdx).getFormula() || sheet.getRange(5, activeColIdx).getValue() || "");
  if (!sampleFormula.includes(expectedRange)) return true;

  const visibleRowCount = Math.max(0, sheet.getLastRow() - 5);
  const visibleMembers = new Set(
    visibleRowCount > 0
      ? sheet.getRange(5, 1, visibleRowCount, 1).getValues().map(r => String(r[0] || "").trim()).filter(Boolean)
      : []
  );
  return (ledger.personShares || []).some(ps =>
    (Math.abs(ps.netTotal || 0) >= 0.005 || Math.abs(ps.purchases || 0) >= 0.005) &&
    !visibleMembers.has(ps.person)
  );
}

// ==========================================
// 10B. UI: MEMBER BREAKDOWN (INTERACTIVE FILTER)
// ==========================================

function buildMemberBreakdownSpillFormulas_(ss, bounds) {
  const txRef = `'${CONFIG.SHEETS.TRANSACTIONS}'`;
  const instRef = `'${CONFIG.SHEETS.INSTALLMENTS}'`;
  const arRef = `'${CONFIG.SHEETS.ACTIVE_RECON}'`;

  const txColsExpr = `ARRAYFORMULA({${txRef}!$A$2:$C, IFERROR(${txRef}!$E$2:$E / IF($B$3="All Family Members", 1, LEN(${txRef}!$D$2:$D) - LEN(SUBSTITUTE(${txRef}!$D$2:$D, "+", "")) + 1), 0), ${txRef}!$D$2:$D, IF(${txRef}!$A$2:$A<>"","✅ Recorded","")})`;
  const txBaseCond = `(${txRef}!$A$2:$A<>"") * (${txRef}!$E$2:$E>0) * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), ${txRef}!$D$2:$D)))) * (${txRef}!$B$2:$B>=$Z$1) * (${txRef}!$B$2:$B<($Z$2+1))`;
  const txNotOrphanCond = bounds.hasOrphanTable
    ? ` * MAP(${txRef}!$B$2:$B, ${txRef}!$C$2:$C, ${txRef}!$D$2:$D, LAMBDA(dt, dc, pr, IF(dc="", FALSE, COUNTIFS(${arRef}!$B$${bounds.startOrphanRow}:$B$${bounds.endOrphanRow}, dt, ${arRef}!$C$${bounds.startOrphanRow}:$C$${bounds.endOrphanRow}, dc & "*", ${arRef}!$E$${bounds.startOrphanRow}:$E$${bounds.endOrphanRow}, IF($B$3="All Family Members", IF(ISNUMBER(FIND("+", pr)), "*", "*" & TRIM(pr) & "*"), "*" & TRIM($B$3) & "*")) < IF($B$3="All Family Members", LEN(pr) - LEN(SUBSTITUTE(pr, "+", "")) + 1, 1))))`
    : "";
  const purchFilterPart = `IFERROR(FILTER(${txColsExpr}, ${txBaseCond}${txNotOrphanCond}), {"—", "—", "No direct purchases recorded in this billing cycle", 0, "—", "—"})`;

  let exFilterPart;
  if (bounds.numExceptions > 0) {
    const exAmountExpr = `(IF(ISNUMBER(SEARCH("Credit", ${arRef}!$E$${bounds.startExRow}:$E$${bounds.endExRow})), -1, 1) * ${arRef}!$D$${bounds.startExRow}:$D$${bounds.endExRow}) / IF($B$3="All Family Members", 1, LEN(${arRef}!$F$${bounds.startExRow}:$F$${bounds.endExRow}) - LEN(SUBSTITUTE(${arRef}!$F$${bounds.startExRow}:$F$${bounds.endExRow}, "+", "")) + 1)`;
    const exColsExpr = `ARRAYFORMULA({${arRef}!$A$${bounds.startExRow}:$C$${bounds.endExRow}, ${exAmountExpr}, ${arRef}!$F$${bounds.startExRow}:$F$${bounds.endExRow}, ${arRef}!$E$${bounds.startExRow}:$E$${bounds.endExRow}})`;
    const exFilterCond = `(${arRef}!$A$${bounds.startExRow}:$A$${bounds.endExRow}<>"") * (${arRef}!$F$${bounds.startExRow}:$F$${bounds.endExRow}<>"") * (ISERROR(SEARCH("🚫", ${arRef}!$F$${bounds.startExRow}:$F$${bounds.endExRow}))) * (ISERROR(SEARCH("🔗", ${arRef}!$F$${bounds.startExRow}:$F$${bounds.endExRow}))) * (ISERROR(SEARCH("Installment", ${arRef}!$E$${bounds.startExRow}:$E$${bounds.endExRow}))) * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), ${arRef}!$F$${bounds.startExRow}:$F$${bounds.endExRow}))))`;
    exFilterPart = `IFERROR(FILTER(${exColsExpr}, ${exFilterCond}), {"—", "—", "No unrecorded statement items, fees, or refunds", 0, "—", "—"})`;
  } else {
    exFilterPart = `{"—", "—", "No statement adjustments for this cycle", 0, "—", "—"}`;
  }

  const purchSpillFormula = `=VSTACK(${purchFilterPart}, {"", "", "", "", "", ""}, {"#", "Tx Date", "🧾 Unrecorded Statement Items, Fees & Refunds", "Amount (EGP)", "Assigned Payer", "Category"}, ${exFilterPart})`;

  let instColsExpr;
  const instSheetObj = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  const instColsInfo = instSheetObj ? getOrCreateInstallmentProgressCols(instSheetObj) : null;
  if (instColsInfo && instColsInfo.progressCol === 10 && instColsInfo.statusCol === 11) {
    instColsExpr = `ARRAYFORMULA({${instRef}!$A$2:$E, ROUND(${instRef}!$H$2:$H, 2) / IF($B$3="All Family Members", 1, LEN(${instRef}!$I$2:$I) - LEN(SUBSTITUTE(${instRef}!$I$2:$I, "+", "")) + 1), ${instRef}!$I$2:$K})`;
  } else if (instColsInfo && instColsInfo.progressCol > 11 && instColsInfo.statusCol > 11) {
    const progLet = getColumnLetter(instColsInfo.progressCol);
    const statLet = getColumnLetter(instColsInfo.statusCol);
    instColsExpr = `ARRAYFORMULA({${instRef}!$A$2:$E, ROUND(${instRef}!$H$2:$H, 2) / IF($B$3="All Family Members", 1, LEN(${instRef}!$I$2:$I) - LEN(SUBSTITUTE(${instRef}!$I$2:$I, "+", "")) + 1), ${instRef}!$I$2:$I, ${instRef}!$${progLet}$2:$${progLet}, ${instRef}!$${statLet}$2:$${statLet}})`;
  } else {
    instColsExpr = `ARRAYFORMULA({${instRef}!$A$2:$E, ROUND(${instRef}!$H$2:$H, 2) / IF($B$3="All Family Members", 1, LEN(${instRef}!$I$2:$I) - LEN(SUBSTITUTE(${instRef}!$I$2:$I, "+", "")) + 1), ${instRef}!$I$2:$K})`;
  }
  const statColLet = getColumnLetter((instColsInfo && instColsInfo.statusCol) || 11);
  const instFilterCond = bounds.hasTable4
    ? `(${instRef}!$A$2:$A<>"") * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), ${instRef}!$I$2:$I)))) * (${instRef}!$B$2:$B<($Z$2+1)) * MAP(${instRef}!$C$2:$C, ${instRef}!$I$2:$I, LAMBDA(d, p, IF(d="", FALSE, COUNTIFS(${arRef}!$C$${bounds.startInstRow}:$C$${bounds.endInstRow}, d & " (*", ${arRef}!$F$${bounds.startInstRow}:$F$${bounds.endInstRow}, IF($B$3="All Family Members", IF(ISNUMBER(FIND("+", p)), "✅*", "✅*" & TRIM(p) & "*"), "✅*" & TRIM($B$3) & "*"))>0)))`
    : `(${instRef}!$A$2:$A<>"") * (${instRef}!$${statColLet}$2:$${statColLet}<>"Completed") * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), ${instRef}!$I$2:$I)))) * (${instRef}!$B$2:$B<($Z$2+1))`;
  const instSpillFormula = `=IFERROR(FILTER(${instColsExpr}, ${instFilterCond}), {"—", "—", "No active installments for this cycle", 0, 0, 0, "—", "—", "—"})`;

  return { purchSpillFormula, instSpillFormula };
}

function syncMemberBreakdownSpillFormulas_(ss, tz, ledger, recon) {
  const mbSheet = ss && ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
  if (!mbSheet || mbSheet.getLastRow() < 10) {
    renderMemberBreakdownSheet(ss, tz, ledger, recon);
    return;
  }

  const bounds = getReconTableBounds(ledger, recon);
  const a6Formula = String(mbSheet.getRange("A6").getFormula() || mbSheet.getRange("A6").getValue() || "");
  const expectedTotalRef = `$H$${bounds.totShareRow}`;
  const expectedShareRange = `$H$${bounds.startShareRow}:$H$${bounds.endShareRow}`;

  if (!a6Formula.includes(expectedTotalRef) || !a6Formula.includes(expectedShareRange)) {
    renderMemberBreakdownSheet(ss, tz, ledger, recon);
    return;
  }

  const spills = buildMemberBreakdownSpillFormulas_(ss, bounds);
  mbSheet.getRange("A10").setFormula(spills.purchSpillFormula);
  mbSheet.getRange("H10").setFormula(spills.instSpillFormula);
}

function renderMemberBreakdownSheet(ss, tz, ledger, recon) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.MEMBER_BREAKDOWN);

  // Preserve currently selected person if valid
  let currentSelection = "All Family Members";
  try {
    const existingVal = sheet.getRange("B3").getValue();
    if (existingVal && typeof existingVal === "string" && existingVal.trim()) {
      currentSelection = existingVal.trim();
    }
  } catch (e) { }

  sheet.clear();
  sheet.setHiddenGridlines(false);

  // 1. Title Banner
  sheet.getRange("A1:P1").merge()
    .setValue("💳 NBE Credit Card — Personal Member Expense & Bill Breakdown")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  sheet.getRange("A2:P2").merge()
    .setValue("Select any family member below or choose 'All Family Members' to inspect itemized purchases, monthly installment EMIs, and fees.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED)
    .setHorizontalAlignment("center");

  // 2. Filter & Metadata Bar (Row 3)
  const uniquePayers = getAllUniquePayers(ss);
  const payerOptions = ["All Family Members", ...uniquePayers];

  sheet.getRange("A3").setValue("Family Member:").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("right");
  const filterCell = sheet.getRange("B3:C3").merge();
  filterCell.setValue(payerOptions.includes(currentSelection) ? currentSelection : "All Family Members")
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setBackground("#fffde7")
    .setBorder(true, true, true, true, false, false);

  const valRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(payerOptions, true)
    .setAllowInvalid(true)
    .build();
  filterCell.setDataValidation(valRule);

  sheet.getRange("D3").setValue("Billing Cycle:").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("right");
  sheet.getRange("E3").setFormula(`='${CONFIG.SHEETS.ACTIVE_RECON}'!B7`).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");

  sheet.getRange("F3").setValue("Due Date:").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("right");
  sheet.getRange("G3:H3").merge().setFormula(`='${CONFIG.SHEETS.ACTIVE_RECON}'!B6`).setFontSize(9).setHorizontalAlignment("center");

  sheet.getRange("I3").setValue("Status:").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("right");
  sheet.getRange("J3:K3").merge().setFormula(`='${CONFIG.SHEETS.ACTIVE_RECON}'!H7`).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");

  // 3. Helper Dates for Period Filtering (Hidden in Col Z)
  const cardBal = recon.cardBal || calculateCardBalance(ss, tz);
  const cycleDateRange = getCycleDateRange(cardBal.activeCycleKey, tz);
  const pStart = cycleDateRange ? cycleDateRange.start : (cardBal.stmtPeriod && cardBal.stmtPeriod.start);
  const pEnd = cycleDateRange ? cycleDateRange.end : (cardBal.stmtPeriod && cardBal.stmtPeriod.end);
  if (pStart) {
    sheet.getRange("Z1")
      .setFormula(`=DATE(${pStart.getFullYear()},${pStart.getMonth() + 1},${pStart.getDate()})`)
      .setNumberFormat("yyyy-MM-dd");
  }
  if (pEnd) {
    sheet.getRange("Z2")
      .setFormula(`=DATE(${pEnd.getFullYear()},${pEnd.getMonth() + 1},${pEnd.getDate()})`)
      .setNumberFormat("yyyy-MM-dd");
  }
  try { sheet.hideColumns(26); } catch (e) { }

  // 4. Dynamic Row Indices from Active Bill & Reconciliation
  const bounds = getReconTableBounds(ledger, recon);
  const startShareRow = bounds.startShareRow;
  const endShareRow = bounds.endShareRow;
  const totShareRow = bounds.totShareRow;

  // 5. KPI Summary Ribbon (Rows 5 & 6)
  // Perfectly mirrors Table 1 of Active Bill & Reconciliation so Net Total Owed ALWAYS equals the actual payment!
  const kpiCards = [
    { title: "Net Total Owed", headerRange: "A5:C5", valRange: "A6:C6", formula: `=IF($B$3="All Family Members", '${CONFIG.SHEETS.ACTIVE_RECON}'!$H$${totShareRow}, IFERROR(SUMIFS('${CONFIG.SHEETS.ACTIVE_RECON}'!$H$${startShareRow}:$H$${endShareRow}, '${CONFIG.SHEETS.ACTIVE_RECON}'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` },
    { title: "Direct Purchases", headerRange: "D5:F5", valRange: "D6:F6", formula: `=IF($B$3="All Family Members", '${CONFIG.SHEETS.ACTIVE_RECON}'!$C$${totShareRow}, IFERROR(SUMIFS('${CONFIG.SHEETS.ACTIVE_RECON}'!$C$${startShareRow}:$C$${endShareRow}, '${CONFIG.SHEETS.ACTIVE_RECON}'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` },
    { title: "Monthly Installments", headerRange: "G5:I5", valRange: "G6:I6", formula: `=IF($B$3="All Family Members", '${CONFIG.SHEETS.ACTIVE_RECON}'!$D$${totShareRow}, IFERROR(SUMIFS('${CONFIG.SHEETS.ACTIVE_RECON}'!$D$${startShareRow}:$D$${endShareRow}, '${CONFIG.SHEETS.ACTIVE_RECON}'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` },
    { title: "Unrecorded Statement Items", headerRange: "J5:L5", valRange: "J6:L6", formula: `=IF($B$3="All Family Members", '${CONFIG.SHEETS.ACTIVE_RECON}'!$F$${totShareRow}, IFERROR(SUMIFS('${CONFIG.SHEETS.ACTIVE_RECON}'!$F$${startShareRow}:$F$${endShareRow}, '${CONFIG.SHEETS.ACTIVE_RECON}'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` },
    { title: "Bank Fees & Taxes", headerRange: "M5:N5", valRange: "M6:N6", formula: `=IF($B$3="All Family Members", '${CONFIG.SHEETS.ACTIVE_RECON}'!$E$${totShareRow}, IFERROR(SUMIFS('${CONFIG.SHEETS.ACTIVE_RECON}'!$E$${startShareRow}:$E$${endShareRow}, '${CONFIG.SHEETS.ACTIVE_RECON}'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` },
    { title: "Credits & Refunds", headerRange: "O5:P5", valRange: "O6:P6", formula: `=IF($B$3="All Family Members", '${CONFIG.SHEETS.ACTIVE_RECON}'!$G$${totShareRow}, IFERROR(SUMIFS('${CONFIG.SHEETS.ACTIVE_RECON}'!$G$${startShareRow}:$G$${endShareRow}, '${CONFIG.SHEETS.ACTIVE_RECON}'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` }
  ];

  kpiCards.forEach((kpi, idx) => {
    sheet.getRange(kpi.headerRange).merge()
      .setValue(kpi.title)
      .setFontWeight("bold")
      .setFontSize(8)
      .setHorizontalAlignment("center")
      .setBackground(idx === 0 ? CONFIG.COLORS.PRIMARY : CONFIG.COLORS.HEADER)
      .setFontColor(idx === 0 ? "#ffffff" : "#333333")
      .setBorder(true, true, false, true, false, false);

    const valCell = sheet.getRange(kpi.valRange).merge();
    valCell.setFormula(kpi.formula)
      .setFontWeight("bold")
      .setFontSize(idx === 0 ? 12 : 10)
      .setHorizontalAlignment("center")
      .setBackground(idx === 0 ? CONFIG.COLORS.PRIMARY_LIGHT : "#ffffff")
      .setFontColor(idx === 0 ? "#1b5e20" : "#333333")
      .setBorder(false, true, true, true, false, false)
      .setNumberFormat("#,##0.00 \"EGP\"");
  });

  // 6. Informative Note Strip (Row 7)
  const noteFormula = `="ℹ️ Active Bill Breakdown: Purchases (" & TEXT(D6, "#,##0.00") & ") + Installments (" & TEXT(G6, "#,##0.00") & ")" & IF(J6>0, " + Unrecorded Purchases (" & TEXT(J6, "#,##0.00") & ")", "") & IF(M6>0, " + Bank Fees (" & TEXT(M6, "#,##0.00") & ")", "") & IF(O6>0, " - Refunds (" & TEXT(O6, "#,##0.00") & ")", "") & " = Net Total Due: " & TEXT(A6, "#,##0.00") & " EGP"`;
  sheet.getRange("A7:P7").merge()
    .setFormula(noteFormula)
    .setFontSize(8)
    .setFontStyle("italic")
    .setFontColor(CONFIG.COLORS.TEXT_MUTED)
    .setBackground("#f9fbe7");

  // 7. Parallel Table 1 (Left, Cols A-F): Direct Purchases & Card Expenses (Starts Row 8, Spills Row 10+)
  sheet.getRange("A8:F8").merge()
    .setValue("🛍️ Direct Purchases & Card Expenses (Active Cycle)")
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

  const purchHeaders = ["#", "Tx Date", "Merchant / Description", "Amount (EGP)", "Assigned Payer", "Status"];
  sheet.getRange(9, 1, 1, purchHeaders.length).setValues([purchHeaders])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER);

  const spills = buildMemberBreakdownSpillFormulas_(ss, bounds);
  sheet.getRange("A10").setFormula(spills.purchSpillFormula).setFontSize(9);
  sheet.getRange("D10:D250").setNumberFormat("#,##0.00");
  sheet.getRange("B10:B250").setNumberFormat("yyyy-MM-dd");

  if (typeof SpreadsheetApp.newConditionalFormatRule === "function" && typeof sheet.setConditionalFormatRules === "function") {
    try {
      const subHeaderRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$C10="🧾 Unrecorded Statement Items, Fees & Refunds"')
        .setBold(true)
        .setBackground(CONFIG.COLORS.HEADER)
        .setRanges([sheet.getRange("A10:F250")])
        .build();
      sheet.setConditionalFormatRules([subHeaderRule]);
    } catch (e) { }
  }

  // 8. Parallel Table 2 (Right, Cols H-P): Active Monthly Installments & EMIs (Starts Row 8, Spills Row 10+)
  sheet.getRange("H8:P8").merge()
    .setValue("📅 Active Monthly Installments (Current Month EMIs)")
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

  const instHeaders = ["#", "Start Date", "Item Description", "Total Price (EGP)", "Months", "Monthly EMI (EGP)", "Payer", "Bank Progress", "Status"];
  sheet.getRange(9, 8, 1, instHeaders.length).setValues([instHeaders])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER);

  sheet.getRange("H10").setFormula(spills.instSpillFormula).setFontSize(9);
  sheet.getRange("K10:K150").setNumberFormat("#,##0.00");
  sheet.getRange("M10:M150").setNumberFormat("#,##0.00").setFontWeight("bold");
  sheet.getRange("I10:I150").setNumberFormat("yyyy-MM-dd");

  try { sheet.setFrozenRows(3); } catch (e) { }

  // 9. Layout Fitting & Ergonomics
  applySheetLayoutFitting(sheet, {
    bannerRow: 1,
    bannerHeight: 30,
    customRowHeights: { 2: 20, 3: 26, 5: 22, 6: 28, 7: 20, 8: 24, 9: 22 },
    explicitColWidths: {
      1: 45,   // A: #
      2: 100,  // B: Date / Dropdown
      3: 175,  // C: Description
      4: 95,   // D: Amount / Label
      5: 85,   // E: Payer / Cycle
      6: 85,   // F: Status / Category
      7: 25,   // G: Spacer
      8: 45,   // H: #
      9: 95,   // I: Start Date
      10: 165, // J: Description
      11: 85,  // K: Total
      12: 55,  // L: Months
      13: 95,  // M: EMI
      14: 75,  // N: Payer
      15: 75,  // O: Paid / Bank Progress
      16: 80   // P: Status
    }
  });
}

// ==========================================
// 11. FAST onEdit & WORKFLOW TRIGGERS
// ==========================================

function syncInstallmentsFromStatement(ss, tz, isBillPaid, optDecisions) {
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  if (!instSheet || instSheet.getLastRow() < 2) return null;
  if (!stmtSheet || stmtSheet.getLastRow() < 7) return null;

  const decisions = optDecisions || loadDecisions(ss);
  const idInfo = ensureAllSheetRowIds(ss, decisions);
  const cardBal = calculateCardBalance(ss, tz);
  const activeCycleKey = cardBal ? cardBal.activeCycleKey : "";

  // 1. Read Statement Installments (read-only)
  const stmtLastRow = stmtSheet.getLastRow();
  const stmtData = stmtSheet.getRange(7, 1, stmtLastRow - 6, 7).getValues();
  const stmtInstallments = [];

  stmtData.forEach((row, idx) => {
    const type = String(row[4] || "").trim().toUpperCase();
    const desc = String(row[3] || "").trim();
    const amt = parseFloat(row[5]) || 0;

    const parsed = parseInstallmentProgressText(desc);
    if (parsed.instCurrent || type === "INSTALLMENT") {
      stmtInstallments.push({
        id: idx + 1,
        stmtId: (idInfo.stmt.ids && idInfo.stmt.ids[idx]) || "",
        stmtRowIndex: idx,
        desc: desc,
        amount: Math.abs(amt),
        instCurrent: parsed.instCurrent,
        instTotal: parsed.instTotal
      });
    }
  });

  if (stmtInstallments.length === 0) return null;

  // 2. Read Installments Sheet (Cols A:I user inputs strictly read-only)
  const sheetInstallmentRows = readSheetInstallmentRows(instSheet, idInfo.inst.ids, decisions, activeCycleKey, tz);

  // 3. Match by Link, compute bank-derived progress, & persist INST_PROGRESS in _Decisions
  const matchResult = matchStatementInstallmentsByLink(stmtInstallments, sheetInstallmentRows, decisions);
  const progressMap = computeBankDerivedInstallmentProgress(matchResult.matches, sheetInstallmentRows, isBillPaid);
  persistBankObservedInstallmentProgress(progressMap, decisions, activeCycleKey);

  // 4. Keep Bank Progress (Col J) / Bank Status (Col K) formula columns synced to active Table 4 rows
  let arBounds = null;
  const startShareRow = getActiveReconStartShareRow();
  const arSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  if (arSheet && arSheet.getLastRow() >= startShareRow) {
    let startInstRow = -1;
    let endInstRow = -1;
    const colA = arSheet.getRange(1, 1, arSheet.getLastRow(), 1).getValues();
    for (let r = startShareRow; r <= colA.length; r++) {
      if (/^INST-\d+/i.test(String(colA[r - 1][0] || "").trim())) {
        if (startInstRow === -1) startInstRow = r;
        endInstRow = r;
      }
    }
    if (startInstRow > 0) {
      arBounds = { hasTable4: true, startInstRow, endInstRow };
    }
  }
  renderInstallmentProgressFormulas(ss, arBounds, { decisions, progressMap, activeCycleKey, tz });

  if (!optDecisions && decisions._dirty) {
    saveDecisions(ss, decisions);
  }

  return { matches: matchResult.matches, progressMap };
}

function handlePaymentDateEdit(sheet, range, tz, forceReapply) {
  const ss = sheet.getParent();
  const rawVal = range.getValue();
  const decisions = loadDecisions(ss);

  const cardBal = calculateCardBalance(ss, tz);
  const activeCycleKey = cardBal.activeCycleKey;
  const cycleLabel = cardBal.activeCycleLabel;
  const closingBill = cardBal.billedBalance;

  // Case 1: User cleared the cell (Unpaid / Undo)
  if (rawVal === "" || rawVal === null || rawVal === undefined) {
    syncInstallmentsFromStatement(ss, tz, false, decisions);

    clearCyclePayment(decisions, activeCycleKey);
    saveDecisions(ss, decisions);

    // Reset cell styling in Active Bill & Reconciliation
    range.setBackground("#fffde7").setFontColor("#333333").setFontWeight("normal");
    sheet.getRange(7, 8).setBackground(CONFIG.COLORS.DUE_SOON).setFontColor("#7f6000");

    ss.toast("Payment date cleared. Installment Bank Progress formulas updated to pending state.", "💳 NBE Tracker", 4);
    return;
  }

  // Case 2: User entered a date
  const parsedDate = parseDateValue(rawVal, tz);
  if (!parsedDate || isNaN(parsedDate.getTime())) {
    ss.toast("Please enter a valid date (e.g. YYYY-MM-DD).", "⚠️ Invalid Date", 4);
    return;
  }

  const dateStr = Utilities.formatDate(parsedDate, tz, "yyyy-MM-dd");
  range.setValue(dateStr);
  range.setNumberFormat("yyyy-MM-dd");
  range.setBackground(CONFIG.COLORS.PAID).setFontColor("#1b5e20").setFontWeight("bold");
  sheet.getRange(7, 8).setBackground(CONFIG.COLORS.PAID).setFontColor("#1b5e20");

  // Synchronize Installments sheet with Ground Truth (isBillPaid = true)
  syncInstallmentsFromStatement(ss, tz, true, decisions);

  // Save record in _Decisions
  setCyclePayment(decisions, activeCycleKey, {
    cycleKey: activeCycleKey,
    dateStr: dateStr,
    cycleLabel: cycleLabel,
    amount: closingBill,
    status: "Settled",
    isApplied: true
  });
  saveDecisions(ss, decisions);

  ss.toast("Payment recorded (" + dateStr + ")! Installment Bank Progress formulas updated to paid state.", "💳 NBE Tracker", 4);
}

function handleActiveReconEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const col = e.range.getColumn();
  const startRow = e.range.getRow();
  const numRows = typeof e.range.getNumRows === "function"
    ? e.range.getNumRows()
    : (e.range.getLastRow ? (e.range.getLastRow() - startRow + 1) : 1);
  const numCols = typeof e.range.getNumColumns === "function" ? e.range.getNumColumns() : 1;

  // Check if user edited the Payment Date cell (Row 6 Col 8, or adjacent to "Payment Date")
  if (numRows === 1 && numCols === 1) {
    let isPaymentDateCell = (startRow === 6 && col === 8);
    if (!isPaymentDateCell && col >= 2 && col <= 8 && startRow <= 8) {
      const prevCellVal = String(sheet.getRange(startRow, col - 1).getValue() || "").trim().toLowerCase();
      if (prevCellVal.includes("payment date")) {
        isPaymentDateCell = true;
      }
    }

    if (isPaymentDateCell) {
      const ss = sheet.getParent();
      const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
      handlePaymentDateEdit(sheet, e.range, tz, true);
      return;
    }
  }

  const lastR = sheet.getLastRow();
  const loopStart = Math.max(getActiveReconStartShareRow(), startRow);
  const loopEnd = Math.min(lastR, startRow + numRows - 1);
  if (loopStart > loopEnd) return;

  const ss = sheet.getParent();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const decisions = loadDecisions(ss);
  const colsToRead = Math.min(9, Math.max(8, sheet.getLastColumn() || 8));
  const rawRows = sheet.getRange(1, 1, lastR, colsToRead).getValues();
  const allRows = colsToRead >= 9
    ? rawRows
    : rawRows.map(r => [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], ""]);
  if (numRows === 1 && numCols === 1 && e.value !== undefined && allRows[startRow - 1] && col <= allRows[startRow - 1].length) {
    allRows[startRow - 1][col - 1] = e.value;
  }
  let modified = false;
  let needsRefresh = false;
  let toastMsg = "";

  for (let row = loopStart; row <= loopEnd; row++) {
    const res = processActiveReconRowChange(sheet, allRows, row, decisions, tz, { allowDelete: true, forceStyle: true });
    if (res && res.changed) {
      modified = true;
      if (res.needsRefresh) needsRefresh = true;
      if (res.toastMsg) toastMsg = res.toastMsg;
    }
  }

  if (modified) {
    saveDecisions(ss, decisions);
    if (!needsRefresh) {
      const targetPayers = new Set();
      for (let row = loopStart; row <= loopEnd; row++) {
        const rawPayer = String((allRows[row - 1] && allRows[row - 1][5]) || "").trim();
        if (rawPayer && !isNeglectedPayer(rawPayer) && !/^🔗/.test(rawPayer)) {
          const cleanStr = rawPayer.replace(/^✅\s*Billed\s+to\s+/i, "");
          splitPayerNames(cleanStr).forEach(p => { if (p) targetPayers.add(p); });
        }
      }
      if (targetPayers.size > 0) {
        const startShareRow = getActiveReconStartShareRow();
        for (let sr = startShareRow; sr <= allRows.length; sr++) {
          const rowArr = allRows[sr - 1];
          if (!rowArr || String(rowArr[0] || "") === "TOTAL") break;
          if (targetPayers.has(String(rowArr[1] || "").trim())) {
            try { sheet.showRows(sr); } catch (e) { }
          }
        }
      }
    }
  }
  if (needsRefresh) {
    try {
      ss.toast("⏳ " + (toastMsg || "Updating reconciliation..."), "💳 NBE Tracker", 10);
    } catch (err) { }
    try {
      updateLiveDashboard({ skipPDF: true, fromEdit: true, ss: ss });
      SpreadsheetApp.flush();
      if (toastMsg) {
        try { ss.toast(toastMsg, "💳 NBE Tracker", 4); } catch (err) { }
      }
    } catch (err) {
      Logger.log("[handleActiveReconEdit] Refresh error: " + err);
      try {
        ss.toast("⚠️ Decision saved, but auto-refresh hit an error (" + (err && err.message ? err.message : err) + "). Please click Refresh Dashboard.", "⚠️ NBE Tracker", 8);
      } catch (e2) { }
    }
  }
  // ZERO SCREEN WIPE on normal edits: Declarative formulas in Table 1, Totals, and KPIs recalculate in sub-50ms!
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const ss = sheet.getParent();
  const sheetName = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();

  if (sheetName === CONFIG.SHEETS.SETTINGS) {
    invalidateSettingsCache_();
    loadSettings(ss);
    if (row === 6 && col === 2) {
      const val = e.range.getValue();
      const clean = formatPayerCell(val);
      if (clean !== val) e.range.setValue(clean);
      invalidateSettingsCache_();
    } else if (row >= 6 && col >= 7 && col <= 9) {
      syncSettingsToDecisions(ss, null, { replaceAll: true });
    }
    return;
  }

  if (sheetName === CONFIG.SHEETS.TRANSACTIONS && row >= 2) {
    if (col >= 1 && col <= 5) {
      try { ensureSingleRowId(ss, sheet, 1, 6, row, "T"); } catch (err) { }
    }
    if (col === 2) {
      const val = e.range.getValue();
      if (val && !(val instanceof Date)) {
        const parsed = parseDateValue(val);
        if (parsed && !isNaN(parsed.getTime())) {
          e.range.setValue(parsed);
        }
      }
      e.range.setNumberFormat("dddd, MMMM d, yyyy");
      return;
    }
    if (col === 4) {
      loadSettings(ss);
      const val = e.range.getValue();
      const clean = formatPayerCell(val);
      if (clean !== val) e.range.setValue(clean);
      return;
    }
    return;
  }

  if (sheetName === CONFIG.SHEETS.INSTALLMENTS && row >= 2) {
    if (col >= 1 && col <= 9) {
      try { ensureSingleRowId(ss, sheet, 1, 12, row, "I"); } catch (err) { }
    }
    if (col === 9) {
      loadSettings(ss);
      const val = e.range.getValue();
      const clean = formatPayerCell(val);
      if (clean !== val) e.range.setValue(clean);
    }
    return;
  }

  if (sheetName === CONFIG.SHEETS.BANK_STATEMENT && row >= 7) {
    let stmtId = "";
    if (col >= 1 && col <= 7) {
      try { stmtId = ensureSingleRowId(ss, sheet, 6, 8, row, "S"); } catch (err) { }
    }
    if (col === 5) {
      const val = String(e.range.getValue() || "").trim().toUpperCase();
      if (["DEBIT", "CREDIT", "INSTALLMENT", "FEE"].includes(val) && stmtId) {
        const decisions = loadDecisions(ss);
        setStmtTypeConfirm(decisions, stmtId, val, "Confirmed by user on Bank Statement");
        const tzEdit = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
        const stmtDateCell = sheet.getRange(2, 2).getValue();
        const stmtDateKey = makeNormDateStr(parseDateValue(stmtDateCell, tzEdit) || stmtDateCell, tzEdit) || String(stmtDateCell || "").trim() || "ACTIVE";
        const imp = decisions.importMeta && (decisions.importMeta[stmtDateKey] || decisions.importMeta["ACTIVE"]);
        if (imp && Array.isArray(imp.ambiguousIds)) {
          imp.ambiguousIds = imp.ambiguousIds.filter(id => id !== stmtId);
          setImportMeta(decisions, imp.statementDate || stmtDateKey, imp);
        }
        if (decisions._dirty) saveDecisions(ss, decisions);
      }

      let bg = "#ffffff";
      let statusLabel = "⏳ Checking...";
      if (val === "CREDIT") {
        bg = CONFIG.COLORS.PAID;
        statusLabel = "💳 Payment / Refund";
      } else if (val === "INSTALLMENT") {
        bg = CONFIG.COLORS.PRIMARY_LIGHT;
        statusLabel = "✅ Type Confirmed";
      } else if (val === "FEE") {
        bg = "#fffde7";
        statusLabel = "🏛️ Bank Fee";
      } else if (val === "DEBIT") {
        statusLabel = "✅ Type Confirmed";
      }
      sheet.getRange(row, 7).setValue(statusLabel);
      sheet.getRange(row, 1, 1, 7).setBackground(bg);
      updateLiveDashboard({ skipPDF: true });
    }
    return;
  }

  if (sheetName === CONFIG.SHEETS.ACTIVE_RECON) {
    handleActiveReconEdit(e);
  }
}

function updateLiveDashboard(options) {
  const ss = (options && options.ss) || SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const fromEdit = Boolean(options && options.fromEdit);

  invalidateSettingsCache_();
  loadSettings(ss);
  if (!fromEdit) {
    cleanupDeprecatedSheets(ss);
    sanitizeAllPayerNames(ss);
    sanitizeTransactionDates(ss);
    syncActiveReconSheetToDecisions(ss, tz);
  }

  const parseHealth = evaluateBankStatementHealth(ss, tz);
  const enforceGate = Boolean((options && options.enforceParseGate) || (parseHealth && parseHealth.hasImportMeta));
  if (enforceGate && parseHealth && !parseHealth.passed) {
    renderStatementParseErrorSheet(ss, tz, parseHealth);
    return { blocked: true, parseHealth };
  }

  const recon = reconcileActiveStatement(ss, tz, { parseHealth });
  if (recon.error) {
    Logger.log(recon.error);
    return { blocked: false, parseHealth, error: recon.error };
  }

  const ledger = buildDebtLedger(ss, tz, recon);

  const health = renderActiveReconciliationSheet(ss, tz, ledger, recon);
  if (fromEdit) {
    syncMemberBreakdownSpillFormulas_(ss, tz, ledger, recon);
    if (shouldRedrawMonthlyHistory_(ss, ledger)) {
      renderMonthlyHistoryForecastSheet(ss, tz, ledger);
    }
  } else {
    renderMonthlyHistoryForecastSheet(ss, tz, ledger);
    renderMemberBreakdownSheet(ss, tz, ledger, recon);
  }

  return { blocked: false, parseHealth, recon, ledger, health };
}

function navToMemberBreakdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
  if (!s) {
    updateLiveDashboard();
    s = ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
  }
  if (s) ss.setActiveSheet(s);
}

function applyPaymentDateNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const reconSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  if (!reconSheet) {
    SpreadsheetApp.getUi().alert("Active Bill & Reconciliation sheet not found.");
    return;
  }

  const dateCell = reconSheet.getRange(6, 8);
  const val = dateCell.getValue();
  if (!val) {
    recordActiveBillPaymentPrompt();
    return;
  }

  handlePaymentDateEdit(reconSheet, dateCell, tz, true);
}

function recordActiveBillPaymentPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const reconSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  if (!reconSheet) {
    SpreadsheetApp.getUi().alert("Active Bill & Reconciliation sheet not found.");
    return;
  }

  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const currentVal = reconSheet.getRange(6, 8).getValue();
  const defaultVal = currentVal ? (currentVal instanceof Date ? Utilities.formatDate(currentVal, tz, "yyyy-MM-dd") : String(currentVal)) : todayStr;

  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "Verify & Record Bill Payment",
    `Enter the payment date for this cycle (${defaultVal}):\n(Leave blank to clear/un-pay)`,
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() === ui.Button.OK) {
    const text = resp.getResponseText().trim();
    const cellRange = reconSheet.getRange(6, 8);
    cellRange.setValue(text);
    handlePaymentDateEdit(reconSheet, cellRange, tz, true);
  }
}

function round2_(n) {
  return Number((Number(n) || 0).toFixed(2));
}

/**
 * Builds an anonymized JSON snapshot of the active billing cycle for golden regression tests.
 * Masks all person names (P1..Pn), merchant names (MERCH01..), card digits, and custom notes
 * while preserving exact matching topology, statement metadata, decisions, and baseline totals.
 */
function buildAnonymizedFixture(ss, tz) {
  const targetTz = tz || ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const activeReconSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  if (activeReconSheet) {
    syncActiveReconSheetToDecisions(ss, targetTz);
  }

  const recon = reconcileActiveStatement(ss, targetTz);
  if (recon.error) {
    throw new Error(recon.error);
  }
  const ledger = buildDebtLedger(ss, targetTz, recon);
  const idInfo = ensureAllSheetRowIds(ss);

  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  const decSheet = ss.getSheetByName(CONFIG.SHEETS.DECISIONS);

  const txRaw = (txSheet && txSheet.getLastRow() >= 2)
    ? txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 5).getValues()
    : [];
  const instRaw = (instSheet && instSheet.getLastRow() >= 2)
    ? instSheet.getRange(2, 1, instSheet.getLastRow() - 1, 9).getValues()
    : [];
  const stmtRaw = (stmtSheet && stmtSheet.getLastRow() >= 7)
    ? stmtSheet.getRange(7, 1, stmtSheet.getLastRow() - 6, 7).getValues()
    : [];
  const decRaw = (decSheet && decSheet.getLastRow() >= 2)
    ? decSheet.getRange(2, 1, decSheet.getLastRow() - 1, 7).getValues()
    : [];

  // 1. Deterministic Person Anonymizer (P1..Pn)
  const rawNames = new Set();
  const settings = loadSettings(ss);
  const aliasKeys = [];
  if (settings && settings.nameAliases) {
    Object.keys(settings.nameAliases).forEach(k => {
      aliasKeys.push(k);
      aliasKeys.push(settings.nameAliases[k]);
    });
  }
  function collectPayers(val) {
    splitPayerNames(val).forEach(p => {
      if (p) rawNames.add(p);
    });
  }
  const dataNames = new Set();
  function collectDataPayers(val) {
    splitPayerNames(val).forEach(p => {
      if (p) {
        rawNames.add(p);
        dataNames.add(p);
      }
    });
  }
  const activeDefaultFeePayer = getDefaultFeePayer(ss);
  collectPayers(activeDefaultFeePayer);
  (ledger.personShares || []).forEach(ps => collectPayers(ps.person));
  txRaw.forEach(r => collectDataPayers(r[3]));
  instRaw.forEach(r => collectDataPayers(r[8]));
  decRaw.forEach(r => {
    const type = String(r[0] || "").trim();
    const t = String(r[2] || "").trim();
    if ((type === "STMT_ASSIGN" || type === "INST_OVERRIDE") && t && !isNeglectedPayer(t)) {
      collectDataPayers(t.replace(/^✅\s*Billed\s+to\s+/i, ""));
    }
  });
  const fixtureStartShareRow = getActiveReconStartShareRow();
  if (activeReconSheet && activeReconSheet.getLastRow() >= fixtureStartShareRow) {
    const arVals = activeReconSheet.getRange(fixtureStartShareRow, 1, activeReconSheet.getLastRow() - fixtureStartShareRow + 1, 6).getValues();
    arVals.forEach(r => {
      const colA = String(r[0] || "").trim();
      if (/^[FCM]-\d+/i.test(colA) && r[5] && !isNeglectedPayer(r[5])) {
        collectDataPayers(r[5]);
      } else if (/^INST-\d+/i.test(colA) && /^✅\s*Billed\s+to\s+/i.test(String(r[5] || ""))) {
        collectDataPayers(String(r[5]).replace(/^✅\s*Billed\s+to\s+/i, ""));
      } else if (/^\d+$/.test(colA) && r[1] && String(r[1]) !== "All Family Members") {
        collectPayers(r[1]);
      }
    });
  }

  const sortedPeople = Array.from(rawNames).sort((a, b) => a.localeCompare(b));
  const personMap = {};
  sortedPeople.forEach((p, idx) => {
    personMap[p] = `P${idx + 1}`;
  });

  function mapPerson(name) {
    const norm = normalizePersonName(name, settings);
    if (!norm) return "";
    if (!personMap[norm]) {
      const nextIdx = Object.keys(personMap).length + 1;
      personMap[norm] = `P${nextIdx}`;
    }
    return personMap[norm];
  }

  function mapPayerField(val) {
    const s = String(val || "").trim();
    if (!s) return "";
    if (s === "NEGLECT" || isNeglectedPayer(s)) return s;
    if (s === "All Family Members") return s;
    const billedMatch = s.match(/^(✅\s*Billed\s+to\s+)(.+)$/i);
    if (billedMatch) {
      const people = splitPayerNames(billedMatch[2]).map(mapPerson).filter(Boolean);
      return people.length > 0 ? `${billedMatch[1]}${people.join(" + ")}` : s;
    }
    const people = splitPayerNames(s).map(mapPerson).filter(Boolean);
    return people.length > 0 ? people.join(" + ") : "";
  }

  // Build regex tokens for scrubbing person names embedded inside item descriptions
  const nameTokens = new Set([...sortedPeople, ...aliasKeys]);
  const escapedTokens = Array.from(nameTokens)
    .filter(t => t && t.length >= 3 && t !== "me")
    .sort((a, b) => b.length - a.length)
    .map(t => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  const embeddedNameRegex = escapedTokens.length > 0
    ? new RegExp(`(^|[^a-z0-9])(?:${escapedTokens.join("|")})(?:'s|’s)?(?=$|[^a-z0-9])`, "gi")
    : null;

  // 2. Lossless Merchant Anonymizer (MERCH01..)
  // Pre-seed equivalent CLEAN buckets for pairs that matched via overlap or merchant alias
  const cleanEquivalent = {};
  function normalizeForAnonBucket_(str) {
    let b = String(str || "");
    if (embeddedNameRegex) b = b.replace(embeddedNameRegex, "$1MEMBER");
    b = b.replace(/\b(\d+)\s+(?:OF|\/)\s+(\d+)\b/gi, " ")
      .replace(/\(\d+\s+of\s+\d+\)/gi, " ")
      .replace(/\(\d+\/\d+\s+Paid\)/gi, " ")
      .replace(/\(Split\s+\d+\/\d+\)/gi, " ")
      .replace(/\b(?:THANK\s*YOU|MOBILE\s*TOKEN)\b/gi, " ")
      .replace(/\b(?:PAYMENT|DIRECT\s*DEBIT|SETTLEMENT|DEPOSIT|REFUND|CASHBACK|CASH\s*BACK|REVERSAL|REVERSED|DCC\s*MARKUP\s*FEES?|DCC\s*MARKUP|STAMP\s*TAX|STAMP\s*DUTY|FINANCE\s*CHARGE|INTEREST\s*CHARGE|LATE\s*PAYMENT|ANNUAL\s*MEMBERSHIP|FEES?)\b/gi, " ")
      .replace(/\b\d{3,}[-\s]?\d{3,}\b/g, " ")
      .replace(/\b\d{5,}\b/g, " ");
    return cleanMerchantName(b)
      .replace(/\d+/g, " ")
      .replace(/[^\w\u0600-\u06FF]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }
  function cleanBaseForKey(str) {
    return normalizeForAnonBucket_(str) || makeNormDesc(str);
  }
  (recon.stmtDebits || []).forEach(st => {
    if (!st.matched || !st.matchedItems) return;
    const stClean = cleanBaseForKey(st.desc);
    st.matchedItems.forEach(tx => {
      if (hasMerchantMatch(st.desc, tx.desc, recon.decisions)) {
        const txClean = cleanBaseForKey(tx.desc);
        if (txClean && stClean && txClean !== stClean) {
          cleanEquivalent[txClean] = stClean;
        }
      }
    });
  });

  const merchantMap = {};
  let nextMerchNum = 1;

  function anonymizeDesc(rawDesc) {
    const raw = String(rawDesc || "").trim();
    if (!raw || raw === "—") return raw;

    // Preserve structural suffixes needed by matching/sync engines (extract paren before instSuffix so "(5 of 9)" isn't matched twice)
    let parenSuffix = "";
    const parenMatch = raw.match(/(\(\d+\s+of\s+\d+\)|\(\d+\/\d+\s+Paid\)|\(Split\s+\d+\/\d+\))/i);
    if (parenMatch) parenSuffix = ` ${parenMatch[1]}`;
    const rawWithoutParen = parenMatch ? raw.replace(parenMatch[0], " ") : raw;

    let instSuffix = "";
    const instMatch = rawWithoutParen.match(/\b(\d+\s+(?:OF|\/)\s+\d+)\b/i);
    if (instMatch) instSuffix = ` ${instMatch[1].toUpperCase()}`;

    // Preserve structural bank classification keywords (never containing merchant names)
    const bankKeywords = [];
    if (/\bDIRECT\s*DEBIT\b/i.test(raw)) bankKeywords.push("DIRECT DEBIT");
    else if (/\bPAYMENT\b/i.test(raw)) bankKeywords.push("PAYMENT");
    if (/\bSETTLEMENT\b/i.test(raw)) bankKeywords.push("SETTLEMENT");
    if (/\bDEPOSIT\b/i.test(raw)) bankKeywords.push("DEPOSIT");
    if (/\bDCC\s*MARKUP/i.test(raw)) bankKeywords.push("DCC MARKUP");
    if (/\bREFUND\b/i.test(raw)) bankKeywords.push("REFUND");
    if (/\bCASH\s*BACK\b/i.test(raw)) bankKeywords.push("CASHBACK");
    if (/\bREVERS(?:AL|ED)\b/i.test(raw)) bankKeywords.push("REVERSAL");
    if (/\bSTAMP\s*(?:TAX|DUTY)\b/i.test(raw)) bankKeywords.push("STAMP TAX");
    else if (/\b(?:FINANCE|INTEREST)\s*CHARGE\b/i.test(raw)) bankKeywords.push("FINANCE CHARGE");
    else if (/\bLATE\s*PAYMENT\b/i.test(raw)) bankKeywords.push("LATE PAYMENT");
    else if (/\bANNUAL\s*MEMBERSHIP\b/i.test(raw)) bankKeywords.push("ANNUAL MEMBERSHIP");

    const isUnr = /\bunr_/i.test(raw);
    const cleaned = normalizeForAnonBucket_(raw);

    // Pure bank fee or pure generic bank payment line with no merchant name
    if (!cleaned && bankKeywords.length > 0) {
      return `${bankKeywords.join(" ")}${instSuffix}${parenSuffix}`.trim();
    }

    const canonical = cleanEquivalent[cleaned] || cleaned;
    const bucketKey = canonical ? `CLEAN:${canonical}` : `RAW:${makeNormDesc(raw)}`;

    if (!merchantMap[bucketKey]) {
      const tag = `MERCH${String(nextMerchNum++).padStart(2, "0")}`;
      merchantMap[bucketKey] = isUnr ? `UNR_${tag}` : tag;
    }
    const merchCode = merchantMap[bucketKey];
    const prefix = bankKeywords.length > 0 ? `${bankKeywords.join(" ")} ` : "";
    return `${prefix}${merchCode}${instSuffix}${parenSuffix}`.trim();
  }

  // 3. Serialize Sheets (Dates as ISO yyyy-MM-dd)
  const cb = recon.cardBal || calculateCardBalance(ss, targetTz);
  const bankRows = stmtRaw.map((r, idx) => [
    r[0] || idx + 1,
    makeNormDateStr(r[1], targetTz),
    makeNormDateStr(r[2] || r[1], targetTz),
    anonymizeDesc(r[3]),
    String(r[4] || "").trim().toUpperCase(),
    round2_(parseFloat(r[5]) || 0),
    ""
  ]);

  const anonTxByDateAmt = [];
  const txRows = txRaw.map((r, idx) => {
    const dateStr = makeNormDateStr(r[1], targetTz);
    const origDesc = String(r[2] || "").trim();
    const anonDesc = anonymizeDesc(origDesc);
    const anonPayer = mapPayerField(r[3]);
    const amt = round2_(parseFloat(r[4]) || 0);
    anonTxByDateAmt.push({ dateStr, amt, origNorm: makeNormDesc(origDesc), anonDesc, splitCount: splitPayerNames(r[3]).length || 1 });
    return [r[0] || idx + 1, dateStr, anonDesc, anonPayer, amt];
  });

  const instRows = instRaw.map((r, idx) => {
    const instId = (idInfo.inst.ids && idInfo.inst.ids[idx]) || "";
    const prog = (recon.instProgressMap && ((instId && recon.instProgressMap[instId]) || recon.instProgressMap[idx])) || null;
    const duration = parseInt(r[4], 10) || 0;
    const base = prog
      ? { paymentsMade: prog.paidCount, status: prog.isCompleted ? "Completed" : "Ongoing" }
      : resolveInstallmentBaseline(
          { instId, startDate: parseDateValue(r[1], targetTz), duration },
          recon.decisions,
          cb.activeCycleKey,
          targetTz
        );
    return [
      r[0] || idx + 1,
      makeNormDateStr(r[1], targetTz),
      anonymizeDesc(r[2]),
      round2_(parseFloat(r[3]) || 0),
      duration,
      round2_(parseFloat(r[5]) || 0),
      round2_(parseFloat(r[6]) || 0),
      round2_(parseFloat(r[7]) || 0),
      mapPayerField(r[8]),
      parseInt(base.paymentsMade, 10) || 0,
      base.status || "Ongoing"
    ];
  });

  function anonymizeDecisionSubjectId(type, subjId) {
    const s = String(subjId || "").trim();
    if (type === "TX_NEGLECT") {
      const idx = s.indexOf(":");
      if (idx > 0) {
        return `${s.substring(0, idx)}:${mapPerson(s.substring(idx + 1))}`;
      }
    } else if (type === "INST_OVERRIDE") {
      const m = s.match(/^(.+?::[^:]+):([^:]+)$/);
      if (m) {
        return `${m[1]}:${mapPerson(m[2])}`;
      }
    }
    return s;
  }

  const anonDecisions = decRaw.map(r => {
    const type = String(r[0] || "").trim();
    const subjId = String(r[1] || "").trim();
    const target = String(r[2] || "").trim();
    const source = String(r[4] || "user").trim();
    const locked = r[5] === true || String(r[5]).toUpperCase() === "TRUE";
    const updatedAt = String(r[6] || "").trim();

    if (type === "META_MIGRATION" || type === "META_COUNTER" || type === "INST_PROGRESS" || type === "STMT_TYPE_CONFIRM") {
      return [type, subjId, target, "", source, locked, updatedAt];
    }
    if (type === "META_IMPORT") {
      let metaObj = {};
      try {
        if (target) metaObj = JSON.parse(target);
      } catch (e) { }
      const cleanMeta = {
        statementDate: metaObj.statementDate || subjId || "",
        printedOpening: round2_(metaObj.printedOpening),
        printedClosing: round2_(metaObj.printedClosing),
        printedTotalDebit: round2_(metaObj.printedTotalDebit),
        printedTotalCredit: round2_(metaObj.printedTotalCredit),
        counts: metaObj.counts || { DEBIT: 0, INSTALLMENT: 0, FEE: 0, CREDIT: 0, total: 0, ambiguous: 0 },
        tieOutPassed: Boolean(metaObj.tieOutPassed),
        mathPassed: Boolean(metaObj.mathPassed),
        closingDiff: round2_(metaObj.closingDiff),
        debitDiff: round2_(metaObj.debitDiff),
        creditDiff: round2_(metaObj.creditDiff),
        ambiguousIds: Array.isArray(metaObj.ambiguousIds) ? metaObj.ambiguousIds : []
      };
      return [type, subjId, JSON.stringify(cleanMeta), "", source, locked, updatedAt];
    }
    if (type === "CYCLE_PAYMENT") {
      return [type, subjId, target, "[]", source, locked, updatedAt];
    }
    if (type === "STMT_MATCH" || type === "STMT_INST_LINK") {
      return [type, subjId, target, "", source, locked, updatedAt];
    }
    if (type === "INST_LINK") {
      const parts = subjId.split("|");
      const anonBase = parts[0] ? makeNormDesc(anonymizeDesc(parts[0])) : "inst";
      return [type, [anonBase, ...parts.slice(1)].join("|"), target, "", source, locked, updatedAt];
    }
    if (type === "MERCHANT_ALIAS") {
      return [type, makeNormDesc(anonymizeDesc(subjId)), makeNormDesc(anonymizeDesc(target)), "", source, locked, updatedAt];
    }
    const anonSubj = anonymizeDecisionSubjectId(type, subjId);
    const anonTarget = (type === "TX_NEGLECT") ? target : mapPayerField(target);
    return [type, anonSubj, anonTarget, "", source, locked, updatedAt];
  });

  let paymentDate = "";
  if (activeReconSheet) {
    const h6Date = parseDateValue(activeReconSheet.getRange(6, 8).getValue(), targetTz);
    if (h6Date && !isNaN(h6Date.getTime())) {
      paymentDate = Utilities.formatDate(h6Date, targetTz, "yyyy-MM-dd");
    }
  }

  const expectedShares = {};
  (ledger.personShares || []).forEach(ps => {
    const hasActivity = ps.purchases !== 0 || ps.installments !== 0 || ps.fees !== 0 ||
      ps.unrecorded !== 0 || ps.refunds !== 0 || ps.netTotal !== 0 || dataNames.has(ps.person);
    if (!hasActivity) return;
    const anonP = mapPerson(ps.person);
    if (!anonP) return;
    expectedShares[anonP] = {
      purchases: round2_(ps.purchases),
      installments: round2_(ps.installments),
      fees: round2_(ps.fees),
      unrecorded: round2_(ps.unrecorded),
      refunds: round2_(ps.refunds),
      netTotal: round2_(ps.netTotal)
    };
  });

  return {
    schemaVersion: 2,
    cycleKey: ledger.activeCycleKey,
    cycleLabel: ledger.activeCycleLabel,
    timeZone: targetTz,
    defaultFeePayer: mapPerson(activeDefaultFeePayer),
    exportedAt: new Date().toISOString(),
    bankStatement: {
      meta: {
        statementDate: cb.statementDate ? Utilities.formatDate(cb.statementDate, targetTz, "yyyy-MM-dd") : "",
        dueDate: cb.statementDueDate ? Utilities.formatDate(cb.statementDueDate, targetTz, "yyyy-MM-dd") : "",
        creditLimit: round2_(cb.creditLimit),
        openingBalance: round2_(cb.openingBalance),
        closingBalance: round2_(cb.billedBalance),
        totalDebit: round2_(cb.totalDebit),
        totalCredit: round2_(cb.totalCredit),
        card: "****0000"
      },
      rows: bankRows
    },
    statementIds: idInfo.stmt.ids || [],
    transactions: txRows,
    transactionIds: idInfo.tx.ids || [],
    installments: instRows,
    installmentIds: idInfo.inst.ids || [],
    decisions: anonDecisions,
    reconStore: [],
    paymentDate: paymentDate,
    expectedTotals: {
      closingBill: round2_(ledger.closingBill),
      purchasesTotal: round2_(ledger.purchasesTotal),
      installmentsTotal: round2_(ledger.installmentsTotal),
      feesTotal: round2_(ledger.feesTotal),
      unrecordedTotal: round2_(ledger.unrecordedTotal),
      refundsTotal: round2_(ledger.refundsTotal),
      familySharesTotal: round2_(ledger.familySharesTotal),
      auditDifference: round2_(ledger.auditDifference)
    },
    expectedShares: expectedShares
  };
}

function menuExportAnonymizedFixture() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const ui = SpreadsheetApp.getUi();
  try {
    const fixture = buildAnonymizedFixture(ss, tz);
    const filename = `NBE_Fixture_${fixture.cycleKey}_anonymized.json`;
    const file = DriveApp.createFile(filename, JSON.stringify(fixture, null, 2), MimeType.PLAIN_TEXT);
    ui.alert(
      "Anonymized Test Fixture Exported",
      `Saved to your Google Drive root as:\n${filename}\n\n` +
      `Link: ${file.getUrl()}\n\n` +
      `Rows exported: ${fixture.bankStatement.rows.length} statement, ` +
      `${fixture.transactions.length} transactions, ${fixture.installments.length} installments, ` +
      `${fixture.decisions.length} decision entries.\n` +
      `All member names, merchant names, card numbers, and custom notes have been anonymized.`,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert("Export Failed", String(err && err.message ? err.message : err), ui.ButtonSet.OK);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💳 NBE Tracker")
    .addItem("🔄 Refresh Dashboard & Reconciliation", "updateLiveDashboard")
    .addItem("👥 View Member Breakdown Sheet", "navToMemberBreakdown")
    .addItem("📤 Share Member Statement...", "menuShareMemberStatement")
    .addItem("📥 Process New Bank Statements from Drive", "manualProcessStatements")
    .addSeparator()
    .addItem("⚡ Apply Payment Date to Installments Now", "applyPaymentDateNow")
    .addItem("✅ Verify & Record Bill Payment...", "recordActiveBillPaymentPrompt")
    .addItem("🔒 Close Active Statement to History...", "menuCloseActiveStatement")
    .addItem("⚙️ Open / Create Settings Sheet", "menuOpenSettingsSheet")
    .addSeparator()
    .addItem("🧪 Export Anonymized Test Fixture", "menuExportAnonymizedFixture")
    .addToUi();
}

function setupDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "checkAndProcessNewStatements") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("checkAndProcessNewStatements")
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
}
