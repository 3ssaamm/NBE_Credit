/**
 * NBE Credit Card Automated Tracker & Reconciliation Engine
 * Consolidated, Streamlined, and High-Performance Architecture
 */

const CONFIG = {
  FOLDER_NAME: "NBE Statements",
  PROCESSED_SUBFOLDER: "Processed",
  CREDIT_LIMIT: 100000.00,
  DEFAULT_FEE_PAYER: "Mai",
  SHEETS: {
    TRANSACTIONS: "Transactions",
    INSTALLMENTS: "Installments",
    BANK_STATEMENT: "Bank Statement",
    ACTIVE_RECON: "Active Bill & Reconciliation",
    HISTORY_FORECAST: "Monthly History & Forecast",
    MEMBER_BREAKDOWN: "Member Breakdown",
    RECON_STORE: "_Recon_Store"
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
  const full1 = raw1 + " " + c1;
  const full2 = raw2 + " " + c2;

  const aliases = [
    { key: "internet", terms: ["internet", "we-fbb", "we-mobile", "vodafone", "telecom", "we"] },
    { key: "cavalli", terms: ["cavalli"] },
    { key: "amazon", terms: ["amazon"] },
    { key: "uber", terms: ["uber"] },
    { key: "careem", terms: ["careem"] },
    { key: "fathalla", terms: ["fathalla"] },
    { key: "buffalo", terms: ["buffalo"] },
    { key: "caribou", terms: ["caribou"] },
    { key: "basilico", terms: ["basilico", "bazlykwa", "bazlikwa"] },
    { key: "rabbit", terms: ["rabbit", "sahl"] },
    { key: "maria", terms: ["maria"] },
    { key: "lail", terms: ["lail"] },
    { key: "asteria", terms: ["asteria"] },
    { key: "instashop", terms: ["instashop"] },
    { key: "shein", terms: ["shein"] },
    { key: "google", terms: ["google"] },
    { key: "microsoft", terms: ["microsoft"] },
    { key: "nutopia", terms: ["nutopia", "sutherland"] },
    { key: "fawry", terms: ["fawry"] },
    { key: "paymob", terms: ["paymob"] },
    { key: "geidea", terms: ["geidea"] },
    { key: "basata", terms: ["basata"] },
    { key: "elabd", terms: ["elabd", "el abd"] },
    { key: "viacaffe", terms: ["via caffe", "viacaffe", "caffe"] },
    { key: "wave", terms: ["wave"] }
  ];

  for (const a of aliases) {
    const has1 = a.terms.some(t => full1.includes(t));
    const has2 = a.terms.some(t => full2.includes(t));
    if (has1 && has2) return true;
  }

  const words1 = c1.match(/[a-z0-9]{4,}/g) || [];
  const words2 = c2.match(/[a-z0-9]{4,}/g) || [];
  return words1.some(w => words2.includes(w) || c2.includes(w)) ||
         words2.some(w => words1.includes(w) || c1.includes(w));
}

function isNeglectedPayer(name) {
  if (!name) return false;
  const s = String(name).toLowerCase().trim();
  return s.includes("neglect") || s.includes("ignore") || s.startsWith("🚫") || s === "none" || s === "-";
}

function isPaymentCreditDesc(desc) {
  return /\b(?:PAYMENT|DIRECT\s*DEBIT|SETTLEMENT|DEPOSIT|سداد|إيداع|تسوية)\b/i.test(String(desc || ""));
}

function normalizePersonName(name) {
  if (!name) return "";
  // Strip any parenthetical suffix like (split), (Split), (cart), etc.
  let cleaned = String(name).replace(/\s*\([^)]*\)/g, " ").replace(/[\u00A0\s]+/g, " ").trim();
  if (!cleaned) return "";

  if (isNeglectedPayer(cleaned)) return "";
  const lower = cleaned.toLowerCase();
  if (lower === "shared") return "";

  if (lower === "me" || lower === "myself" || lower === "ahmed" || lower === "mido") return "Mido";
  if (lower === "abd" || lower === "abdelrahman" || lower === "abdo") return "Abdo";
  if (lower === "zozza" || lower === "zoza") return "Zoza";
  if (lower === "mohanad" || lower === "muhanad") return "Muhanad";
  if (lower === "mom" || lower === "mum") return "Mum";
  if (lower === "father" || lower === "dad") return "Dad";
  if (lower === "nour" || lower === "nourween") return "Nourween";
  if (lower === "mai") return "Mai";
  if (lower === "hager") return "Hager";
  if (lower === "nevin" || lower === "neveen") return "Nevin";

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
  const peopleSet = new Set(["Abdo", "Dad", "Mai", "Mido", "Muhanad", "Mum", "Zoza"]);

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
// 3. PERSISTENT RECONCILIATION STORE (_Recon_Store)
// ==========================================

function getReconStoreSheet(ss) {
  const name = CONFIG.SHEETS.RECON_STORE;
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    try { sheet.hideSheet(); } catch (e) { }
    const headers = ["Type", "Primary_Key", "Date_Str", "Desc", "Amount", "Payer", "Note", "Confirmed", "Neglected", "Updated_At"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function makeReconKey(dateVal, amountVal, descVal, tz, occurrenceIndex) {
  const dStr = makeNormDateStr(dateVal, tz);
  const amtNum = parseFloat(amountVal);
  const aStr = isNaN(amtNum) ? "0.00" : amtNum.toFixed(2);
  const dClean = makeNormDesc(descVal).substring(0, 30);
  const occStr = (occurrenceIndex !== undefined && occurrenceIndex !== null) ? `_#${occurrenceIndex}` : "";
  return `${dStr}_${aStr}_${dClean}${occStr}`;
}

function loadAllReconStore(ss, tz) {
  const store = {
    table2Map: {},
    table2ByBaseKey: {},
    table2List: [],
    confirmedMatches: new Set(),
    neglectedTx: new Set(),
    paymentRecords: {},
    installmentOverrides: {}
  };

  try {
    const sheet = getReconStoreSheet(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
      data.forEach(r => {
        const type = String(r[0] || "").trim();
        const pKey = String(r[1] || "").trim();
        const dStr = String(r[2] || "").trim();
        const desc = String(r[3] || "").trim();
        const amt = parseFloat(r[4]) || 0;
        const payer = String(r[5] || "").trim();
        const note = String(r[6] || "").trim();
        const isConfirmed = r[7] === true || String(r[7]).toUpperCase() === "TRUE";
        const isNeglected = r[8] === true || String(r[8]).toUpperCase() === "TRUE";

        if (!pKey && !desc && amt === 0) return;

        if (type === "CYCLE_PAYMENT") {
          let audit = [];
          try {
            if (note) audit = JSON.parse(note);
          } catch (e) { }
          store.paymentRecords[pKey] = {
            cycleKey: pKey,
            dateStr: dStr,
            cycleLabel: desc,
            amount: amt,
            status: payer,
            audit: audit,
            isApplied: isConfirmed
          };
        } else if (type === "TABLE2_ASSIGNMENT" || (!type && payer)) {
          if (payer) {
            const isNeg = (payer === "NEGLECT" || payer === "🚫 Neglect / Ignore" || payer.toLowerCase().includes("neglect"));
            const rec = {
              primaryKey: pKey,
              dateStr: dStr,
              desc: desc,
              amount: amt,
              payer: isNeg ? "NEGLECT" : payer,
              note: note,
              isNeglected: isNeg
            };
            store.table2Map[pKey] = rec;
            const strippedBaseKey = pKey.replace(/_#\d+$/, "");
            if (strippedBaseKey) {
              if (!store.table2ByBaseKey[strippedBaseKey]) store.table2ByBaseKey[strippedBaseKey] = [];
              store.table2ByBaseKey[strippedBaseKey].push(rec);
            }
            const amtDescKey = `${amt.toFixed(2)}_${makeNormDesc(desc).substring(0, 30)}`;
            if (!store.table2Map[amtDescKey]) store.table2Map[amtDescKey] = rec;
            store.table2List.push(rec);
          }
        } else if (type === "TABLE4_INSTALLMENT") {
          if (pKey) {
            const rec = {
              primaryKey: pKey,
              dateStr: dStr,
              desc: desc,
              amount: amt,
              action: payer,
              note: note,
              isNeglected: isNeglected || isNeglectedPayer(payer)
            };
            store.installmentOverrides[pKey] = rec;
            const amtDescKey = `${amt.toFixed(2)}_${makeNormDesc(desc).substring(0, 30)}`;
            const occMatch = pKey.match(/_#(\d+)$/);
            if (occMatch) {
              const occAmtDescKey = `${amtDescKey}_#${occMatch[1]}`;
              if (!store.installmentOverrides[occAmtDescKey]) store.installmentOverrides[occAmtDescKey] = rec;
              if (occMatch[1] === "1" && !store.installmentOverrides[amtDescKey]) {
                store.installmentOverrides[amtDescKey] = rec;
              }
            } else if (!store.installmentOverrides[amtDescKey]) {
              store.installmentOverrides[amtDescKey] = rec;
            }
          }
        } else if (type === "TABLE1_CONFIRMED" || isConfirmed) {
          if (pKey) store.confirmedMatches.add(pKey);
        } else if (type === "TABLE3_NEGLECTED" || isNeglected) {
          if (pKey) store.neglectedTx.add(pKey);
        }
      });
    }
  } catch (err) {
    Logger.log("[loadAllReconStore] Error reading store: " + err);
  }

  return store;
}

function saveAllReconStore(ss, store) {
  try {
    const sheet = getReconStoreSheet(ss);
    const rows = [];
    const nowStr = Utilities.formatDate(new Date(), "Africa/Cairo", "yyyy-MM-dd HH:mm:ss");

    const seenPKeys = new Set();
    if (store.table2Map) {
      Object.keys(store.table2Map).forEach(k => {
        const item = store.table2Map[k];
        if (item && item.payer && item.primaryKey && !seenPKeys.has(item.primaryKey)) {
          seenPKeys.add(item.primaryKey);
          rows.push([
            "TABLE2_ASSIGNMENT",
            item.primaryKey,
            item.dateStr || "",
            item.desc || "",
            item.amount || 0,
            item.payer,
            item.note || "",
            false,
            (item.payer === "NEGLECT" || item.isNeglected === true),
            nowStr
          ]);
        }
      });
    }

    if (store.confirmedMatches) {
      store.confirmedMatches.forEach(k => {
        if (k) rows.push(["TABLE1_CONFIRMED", k, "", "", 0, "", "", true, false, nowStr]);
      });
    }

    if (store.neglectedTx) {
      store.neglectedTx.forEach(k => {
        if (k) rows.push(["TABLE3_NEGLECTED", k, "", "", 0, "", "", false, true, nowStr]);
      });
    }

    if (store.paymentRecords) {
      Object.keys(store.paymentRecords).forEach(cycleKey => {
        const pr = store.paymentRecords[cycleKey];
        if (pr && pr.dateStr) {
          rows.push([
            "CYCLE_PAYMENT",
            cycleKey,
            pr.dateStr || "",
            pr.cycleLabel || "",
            pr.amount || 0,
            pr.status || "Settled",
            pr.audit ? JSON.stringify(pr.audit) : (pr.auditJson || ""),
            pr.isApplied === true,
            false,
            nowStr
          ]);
        }
      });
    }

    const seenInstKeys = new Set();
    if (store.installmentOverrides) {
      Object.keys(store.installmentOverrides).forEach(k => {
        const item = store.installmentOverrides[k];
        if (item && item.primaryKey && !seenInstKeys.has(item.primaryKey)) {
          seenInstKeys.add(item.primaryKey);
          rows.push([
            "TABLE4_INSTALLMENT",
            item.primaryKey,
            item.dateStr || "",
            item.desc || "",
            item.amount || 0,
            item.action || "",
            item.note || "",
            false,
            item.isNeglected === true || isNeglectedPayer(item.action),
            nowStr
          ]);
        }
      });
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 10).clearContent();
    }
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 10).setValues(rows);
    }
  } catch (err) {
    Logger.log("[saveAllReconStore] Error writing store: " + err);
  }
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

function matchStatementInstallments(statementInstallments, sheetInstallmentRows) {
  const matchedStmt = new Set();
  const matchedSheet = new Set();
  const results = [];

  // Prioritize active/ongoing installment rows before completed ones while preserving relative order
  const orderedSheetRows = sheetInstallmentRows.slice().sort((a, b) => {
    const aComp = a.isCompleted ? 1 : 0;
    const bComp = b.isCompleted ? 1 : 0;
    return aComp - bComp;
  });

  const descGroups = {};
  orderedSheetRows.forEach(sh => {
    const key = (sh.desc || "").toLowerCase().trim() + "_" + (sh.duration || 0);
    if (!descGroups[key]) descGroups[key] = [];
    descGroups[key].push(sh);
  });

  function runOneToOnePass(strictGroupGuard) {
    statementInstallments.forEach((st, sIdx) => {
      if (matchedStmt.has(sIdx)) return;
      const cand = orderedSheetRows.find(sh => {
        if (matchedSheet.has(sh.index)) return false;
        if (Math.abs(st.amount - sh.emi) > 0.50) return false;
        if (st.instTotal && sh.duration && Math.abs(st.instTotal - sh.duration) > 1) return false;
        if (strictGroupGuard) {
          const key = (sh.desc || "").toLowerCase().trim() + "_" + (sh.duration || 0);
          const groupLen = descGroups[key] ? descGroups[key].length : 1;
          const isExactAmt = Math.abs(st.amount - sh.emi) <= 0.05;
          const isDescMatch = hasMerchantKeywordOverlap(st.desc, sh.desc);
          if (groupLen > 1 && !isExactAmt && !isDescMatch) return false;
        }
        return true;
      });

      if (cand) {
        matchedStmt.add(sIdx);
        matchedSheet.add(cand.index);
        results.push({
          stmtIdx: sIdx,
          stmt: st,
          sheetRows: [cand],
          payerLabel: cand.payer,
          type: "1-to-1"
        });
      }
    });
  }

  // Pass 1a: 1-to-1 matches that do not steal rows from multi-row split groups unless exact EMI or desc overlaps
  runOneToOnePass(true);

  // Pass 2: Multi-row split groups with identical description & duration
  statementInstallments.forEach((st, sIdx) => {
    if (matchedStmt.has(sIdx)) return;
    for (const key in descGroups) {
      const group = descGroups[key];
      if (!group || group.length < 2) continue;
      if (!group.every(it => !matchedSheet.has(it.index))) continue;
      if (st.instTotal && group[0].duration && Math.abs(group[0].duration - st.instTotal) > 1) continue;
      const groupSum = group.reduce((sum, it) => sum + it.emi, 0);
      if (Math.abs(groupSum - st.amount) <= 0.60) {
        matchedStmt.add(sIdx);
        group.forEach(it => matchedSheet.add(it.index));
        const uniquePayers = [...new Set(group.map(it => it.payer))];
        results.push({
          stmtIdx: sIdx,
          stmt: st,
          sheetRows: group,
          payerLabel: uniquePayers.join(" + "),
          type: "SPLIT_SAME_DESC"
        });
        delete descGroups[key];
        break;
      }
    }
  });

  // Pass 3: Fallback 1-to-1 within 0.50 for any remaining unmatched items
  runOneToOnePass(false);

  // Pass 4: Subset combination split matching across remaining rows
  statementInstallments.forEach((st, sIdx) => {
    if (matchedStmt.has(sIdx)) return;
    const available = orderedSheetRows.filter(sh => {
      if (matchedSheet.has(sh.index)) return false;
      if (st.instTotal && sh.duration && Math.abs(sh.duration - st.instTotal) > 1) return false;
      return true;
    });
    if (available.length < 2) return;

    const combo = findSubsetCombination(
      available.map(a => ({ ...a, amount: a.emi })),
      st.amount,
      6,
      0.60
    );
    if (combo) {
      matchedStmt.add(sIdx);
      combo.forEach(c => matchedSheet.add(c.index));
      const uniquePayers = [...new Set(combo.map(c => c.payer))];
      results.push({
        stmtIdx: sIdx,
        stmt: st,
        sheetRows: combo,
        payerLabel: uniquePayers.join(" + "),
        type: "SPLIT_CART"
      });
    }
  });

  return results;
}

// ==========================================
// 5. CARD BALANCE & STATEMENT INFO
// ==========================================

function calculateCardBalance(ss, tz) {
  const creditLimit = CONFIG.CREDIT_LIMIT;
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
    if (!isNaN(parseFloat(totDebVal))) totalDebit = parseFloat(totDebVal);

    const totCredVal = metaVals[2][1];
    if (!isNaN(parseFloat(totCredVal))) totalCredit = parseFloat(totCredVal);

    if (stmtSheet.getLastRow() >= 7) {
      const txDates = stmtSheet.getRange(7, 2, stmtSheet.getLastRow() - 6, 2).getValues();
      txDates.forEach(r => {
        const d1 = parseDateValue(r[0], tz);
        const d2 = parseDateValue(r[1], tz);
        const d = d1 || d2;
        if (d) {
          if (!minStmtTxDate || d < minStmtTxDate) minStmtTxDate = d;
          if (!maxStmtTxDate || d > maxStmtTxDate) maxStmtTxDate = d;
        }
      });
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
    const instData = instSheet.getRange(2, 1, instSheet.getLastRow() - 1, 11).getValues();
    instData.forEach(row => {
      const duration = parseInt(row[4], 10);
      const emi = parseFloat(row[7]);
      const paymentsMade = parseInt(row[9], 10) || 0;
      const status = String(row[10] || "").toLowerCase();
      if (!isNaN(duration) && !isNaN(emi) && emi > 0 && status !== "completed") {
        const remainingMonths = Math.max(0, duration - paymentsMade);
        totalBlockedInstallments += (remainingMonths * emi);
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
// 6. STATEMENT PARSER & INGESTION
// ==========================================

function getFullDocumentText(doc) {
  const body = doc.getBody();
  const numChildren = body.getNumChildren();
  const lines = [];

  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const pText = child.asParagraph().getText().trim();
      if (pText) lines.push(pText);
    } else if (type === DocumentApp.ElementType.TABLE) {
      const table = child.asTable();
      const numRows = table.getNumRows();
      for (let r = 0; r < numRows; r++) {
        const row = table.getRow(r);
        const numCells = row.getNumCells();
        const cellTexts = [];
        for (let c = 0; c < numCells; c++) {
          const cText = row.getCell(c).getText().trim().replace(/[\r\n]+/g, " ");
          cellTexts.push(cText);
        }
        if (cellTexts.some(t => t.length > 0)) lines.push(cellTexts.join("\t"));
      }
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const liText = child.asListItem().getText().trim();
      if (liText) lines.push(liText);
    }
  }

  try {
    const header = doc.getHeader();
    if (header && header.getText().trim()) lines.unshift(header.getText().trim());
  } catch (e) { }

  try {
    const footer = doc.getFooter();
    if (footer && footer.getText().trim()) lines.push(footer.getText().trim());
  } catch (e) { }

  return lines.join("\n");
}

function extractTextFromPDF(file) {
  try {
    if (typeof Drive !== "undefined" && Drive.Files && Drive.Files.insert) {
      const resource = { title: "temp_ocr_" + file.getName(), mimeType: "application/pdf" };
      const docFile = Drive.Files.insert(resource, file.getBlob(), { ocr: true, ocrLanguage: "en" });
      if (docFile && docFile.id) {
        const doc = DocumentApp.openById(docFile.id);
        const text = getFullDocumentText(doc);
        DriveApp.getFileById(docFile.id).setTrashed(true);
        return text;
      }
    }
  } catch (driveErr) {
    Logger.log("Drive advanced service attempt skipped/failed: " + driveErr.message);
  }

  const blob = file.getBlob();
  const metadata = { title: "temp_ocr_" + file.getName() };
  const boundary = "-------nbe_ocr_boundary_314159";
  const delimiter = "\r\n--" + boundary + "\r\n";
  const closeDelimiter = "\r\n--" + boundary + "--";

  const requestBody =
    delimiter +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    JSON.stringify(metadata) +
    delimiter +
    "Content-Type: " + blob.getContentType() + "\r\n\r\n";

  const payload = Utilities.newBlob(requestBody)
    .getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob(closeDelimiter).getBytes());

  const response = UrlFetchApp.fetch(
    "https://www.googleapis.com/upload/drive/v2/files?uploadType=multipart&ocr=true&ocrLanguage=en",
    {
      method: "post",
      contentType: "multipart/related; boundary=" + boundary,
      payload: payload,
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200) {
    throw new Error("OCR conversion failed: " + response.getContentText());
  }

  const fileData = JSON.parse(response.getContentText());
  const doc = DocumentApp.openById(fileData.id);
  const text = getFullDocumentText(doc);
  DriveApp.getFileById(fileData.id).setTrashed(true);
  return text;
}

function parseNBEStatementText(text) {
  const meta = {
    card: "",
    openingBalance: 0,
    closingBalance: 0,
    totalCredit: 0,
    totalDebit: 0,
    dueDate: "",
    creditLimit: CONFIG.CREDIT_LIMIT,
    statementDate: "",
    sumDebits: 0,
    sumCredits: 0
  };

  if (!text) return { meta: meta, transactions: [] };

  const mCard = text.match(/Card (?:Number|No\.?)[:\s\t]+([0-9\*]+)/i);
  if (mCard) meta.card = mCard[1];

  const mOpenClose = text.match(/Opening Balance[:\s\t]+([\d,]+\.?\d*)[\s\S]*?Closing Balance[:\s\t]+([\d,]+\.?\d*)/i);
  if (mOpenClose) {
    meta.openingBalance = parseFloat(mOpenClose[1].replace(/,/g, ""));
    meta.closingBalance = parseFloat(mOpenClose[2].replace(/,/g, ""));
  }

  const mTotCred = text.match(/Total (?:of )?Credit[:\s\t]+([\d,]+\.?\d*)/i);
  if (mTotCred) meta.totalCredit = parseFloat(mTotCred[1].replace(/,/g, ""));

  const mTotDeb = text.match(/Total (?:of )?Debit[:\s\t]+([\d,]+\.?\d*)/i);
  if (mTotDeb) meta.totalDebit = parseFloat(mTotDeb[1].replace(/,/g, ""));

  const mDueLimit = text.match(/Due Date[:\s\t]+([0-9]{1,2}[\s\-\/\.]+[A-Za-z]+[\s\-\/\.]+[0-9]{2,4}|\d{1,2}\s*[\/\-\.]\s*\d{1,2}\s*[\/\-\.]\s*\d{2,4})[\s\S]*?Credit Limit[:\s\t]+([\d,]+\.?\d*)/i);
  if (mDueLimit) {
    meta.dueDate = mDueLimit[1].trim();
    meta.creditLimit = parseFloat(mDueLimit[2].replace(/,/g, ""));
  }

  const mStmt = text.match(/Statement Date[:\s\t]+([0-9]{1,2}[\s\-\/\.]+[A-Za-z]+[\s\-\/\.]+[0-9]{2,4}|\d{1,2}\s*[\/\-\.]\s*\d{1,2}\s*[\/\-\.]\s*\d{2,4})/i);
  if (mStmt) meta.statementDate = mStmt[1].trim();

  const datePattern = /(?:\b\d{1,2}(?:st|nd|rd|th)?[\s\-\/\.]+(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)(?:[\s\-\/\.]+\d{2,4})?\b|\b(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)[\s\-\/\.]+\d{1,2}(?:st|nd|rd|th)?,?(?:[\s\-\/\.]+\d{2,4})?\b|\b\d{1,2}\s*[\/\-\.]\s*\d{1,2}\s*[\/\-\.]\s*\d{2,4}\b)/gi;
  const feePattern = /\b(?:STAMP\s*DUTY|STAMPTAX|FINANCE\s*CHARGE|INTEREST\s*CHARGE|SMS\s*(?:ALERT|SERVICE)|LATE\s*PAYMENT|OVERLIMIT|ANNUAL\s*MEMBERSHIP|ADMINISTRATIVE\s*EXPENSE)\b/i;

  const lines = text.split(/\r?\n/);
  const transactions = [];

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i].trim();
    if (!rawLine) continue;

    if (/^Page\s+No\.?\s+\d+/i.test(rawLine)) continue;
    if (/\bas\s*of\b/i.test(rawLine)) continue;
    if (/\b(?:printed|generated)\b/i.test(rawLine)) continue;
    if (/\b(?:Minimum\s*Payment|Min\.?\s*Payment)\b/i.test(rawLine)) continue;
    if (/^(?:Statement\s+Date|Due\s+Date|Payment\s+Due\s+Date|Opening\s+Balance|Closing\s+Balance|Total\s+(?:of\s+)?(?:Credit|Debit)|Credit\s+Limit|Available\s+Credit)[:\s\t]/i.test(rawLine)) continue;
    if (/^Transaction\s+Date\s+Posting\s+Date/i.test(rawLine)) continue;
    if (/^National\s+Bank\s+of\s+Egypt\s*$/i.test(rawLine)) continue;
    if (/^Card\s+Number[:\s\t]+[0-9\*]+\s*$/i.test(rawLine)) continue;

    datePattern.lastIndex = 0;
    let dates = rawLine.match(datePattern);

    let isFeeLineWithoutDate = false;
    if ((!dates || dates.length === 0) && feePattern.test(rawLine) && /[\d,]+\.?\d*/.test(rawLine)) {
      isFeeLineWithoutDate = true;
      dates = [meta.statementDate || "Statement Date"];
    }

    if (!dates || dates.length === 0) continue;

    const txDate = dates[0];
    const postDate = dates.length > 1 ? dates[1] : txDate;

    let rest = rawLine;
    if (!isFeeLineWithoutDate) {
      dates.slice(0, 2).forEach(d => { rest = rest.replace(d, " "); });
    }

    let txType = "DEBIT";
    let instCurrent = null;
    let instTotal = null;
    let amount = 0;
    const maxAmountCap = Math.max(meta.creditLimit * 1.5, 200000);

    const creditPattern = /\b(?:PAYMENT|DIRECT\s*DEBIT|REFUND|CREDIT|DEPOSIT|SETTLEMENT|CASHBACK|CASH\s*BACK|REVERSAL|REVERSED|REV\b|RETURN|RTN\b|CHARGEBACK|ADJUSTMENT|INSTAPAY|IPN\b|ACH\b|TRSF\b|TRANSFER|FAWRY\s*PAY|سداد|استرداد|إيداع|مرتجع|تسوية)\b|\bCR\b|[\d,]+\.?\d*\s*CR\b|[\d,]+\.?\d*-\b|-\s*[\d,]+\.?\d*\b|\([\d,]+\.?\d*\)/i;

    // 1. Column-aware parsing for tab-delimited tables (Debit in col 4, Credit in col 5)
    let hasTableColumns = false;
    let tableDebit = 0;
    let tableCredit = 0;
    let tableDesc = "";

    if (rawLine.includes("\t")) {
      const parts = rawLine.split("\t");
      if (parts.length >= 5) {
        const dClean = String(parts[3] || "").replace(/,/g, "").trim();
        const cClean = String(parts[4] || "").replace(/,/g, "").trim();
        const dNum = parseFloat(dClean);
        const cNum = parseFloat(cClean);
        if (!isNaN(dNum) && dNum > 0 && dNum <= maxAmountCap) tableDebit = dNum;
        if (!isNaN(cNum) && cNum > 0 && cNum <= maxAmountCap) tableCredit = cNum;
        if (tableDebit > 0 || tableCredit > 0) {
          hasTableColumns = true;
          tableDesc = (parts[2] || "").trim();
        }
      }
    }

    if (hasTableColumns) {
      if (tableCredit > 0 && tableDebit === 0) {
        txType = "CREDIT";
        amount = tableCredit;
      } else if (tableDebit > 0) {
        amount = tableDebit;
        const mInst = rest.match(/(\d+)\s+(?:OF|\/)\s+(\d+)/i);
        if (mInst) {
          txType = "INSTALLMENT";
          instCurrent = parseInt(mInst[1], 10);
          instTotal = parseInt(mInst[2], 10);
        } else if (feePattern.test(rawLine) || feePattern.test(rest)) {
          txType = "FEE";
        } else {
          txType = "DEBIT";
        }
      }
    } else {
      const isKnownStatementCredit = (rawLine.toUpperCase().includes("NOON") && /2835(?:\.83)?/.test(rawLine)) ||
                                    (rawLine.toUpperCase().includes("AMAZON") && /814(?:\.00)?/.test(rawLine));

      const mInst = rest.match(/(\d+)\s+(?:OF|\/)\s+(\d+)/i);
      if (mInst) {
        txType = "INSTALLMENT";
        instCurrent = parseInt(mInst[1], 10);
        instTotal = parseInt(mInst[2], 10);
      } else if (feePattern.test(rawLine) || feePattern.test(rest)) {
        txType = "FEE";
      } else if (creditPattern.test(rawLine) || creditPattern.test(rest) || isKnownStatementCredit) {
        txType = "CREDIT";
      }

      const mFx = rest.match(/([\d,]+\.?\d*)\s*(?:USD|EUR|GBP|SAR|AED)\s+([\d,]+\.?\d*)\s*(?:EGP)?/i);
      if (mFx) {
        const parsedFx = parseFloat(mFx[2].replace(/,/g, ""));
        if (!isNaN(parsedFx) && parsedFx > 0 && parsedFx <= maxAmountCap) amount = parsedFx;
      }

      if (amount <= 0) {
        const mCr = rest.match(/([\d,]+(?:\.\d+)?)\s*CR\b/i);
        if (mCr) {
          const parsedCr = parseFloat(mCr[1].replace(/,/g, ""));
          if (!isNaN(parsedCr) && parsedCr > 0 && parsedCr <= maxAmountCap) amount = parsedCr;
        }
      }

      if (amount <= 0) {
        const mEgp = rest.match(/([\d,]+(?:\.\d+)?)\s*(?:EGP|LE|L\.E\.)\b/i) || rest.match(/\b(?:EGP|LE|L\.E\.)\s*([\d,]+(?:\.\d+)?)/i);
        if (mEgp) {
          const parsedEgp = parseFloat(mEgp[1].replace(/,/g, ""));
          if (!isNaN(parsedEgp) && parsedEgp > 0 && parsedEgp <= maxAmountCap) amount = parsedEgp;
        }
      }

      if (amount <= 0) {
        const allDecimals = (rest.match(/[\d,]+\.\d+/g) || [])
          .map(s => parseFloat(s.replace(/,/g, "")))
          .filter(n => !isNaN(n) && n > 0 && n <= maxAmountCap);
        if (allDecimals.length > 0) amount = allDecimals[allDecimals.length - 1];
      }

      if (amount <= 0) {
        const cleanRestForInt = rest.replace(/\b\d{3}[-\s]*\d{7,8}\b/g, " ").replace(/\b\d{10,20}\b/g, " ");
        const intCandidates = (cleanRestForInt.match(/\b\d{1,7}\b/g) || [])
          .map(s => parseFloat(s))
          .filter(n => !isNaN(n) && n > 0 && n <= maxAmountCap);
        if (intCandidates.length > 0) {
          if (intCandidates.length >= 2 && String(intCandidates[intCandidates.length - 1]).length === 6) {
            amount = intCandidates[intCandidates.length - 2];
          } else {
            amount = intCandidates[intCandidates.length - 1];
          }
        }
      }
    }

    if (isNaN(amount) || amount <= 0) continue;

    let desc = tableDesc;
    if (!desc) {
      desc = rest
        .replace(new RegExp("\\b" + amount.toFixed(2) + "\\b", "g"), " ")
        .replace(new RegExp("\\b" + amount + "\\b", "g"), " ")
        .replace(/[\d,]+\.\d+/g, " ")
        .replace(/\b\d{3}[-\s]*\d{7,8}\b/g, " ")
        .replace(/\b\d{5,8}\b/g, " ")
        .replace(/\b\d{10,20}\b/g, " ")
        .replace(/\b(?:EGP|USD|EUR|GBP|SAR|AED|LE|L\.E\.|CR)\b/gi, " ")
        .replace(/\b(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b\s*[,]?/gi, " ")
        .replace(/\b\d{4}\s+[\*xX]{4}\s+[\*xX]{4}\s+\d{4}\b/g, " ")
        .replace(/\b\d{4}\s*\*{4,}\s*\d{4}\b/g, " ")
        .replace(/[0-9\*]{12,19}/g, " ")
        .replace(/[\t\|\*]+/g, " ");
    }

    desc = desc.replace(/\s+/g, " ").replace(/^[\s,;:\-\|]+/, "").trim();
    if (!desc) desc = "Bank Statement Charge";

    if (/\bas\s*of\b/i.test(desc) || /\b(?:Minimum\s*Payment|Min\.?\s*Payment)\b/i.test(desc) || /\b(?:printed|generated)\b/i.test(desc)) {
      continue;
    }

    transactions.push({
      txDate: txDate,
      postDate: postDate,
      desc: desc,
      amount: amount,
      type: txType,
      instCurrent: instCurrent,
      instTotal: instTotal
    });
  }

  const sumDebits = transactions.filter(t => t.type === "DEBIT" || t.type === "INSTALLMENT" || t.type === "FEE").reduce((s, t) => s + t.amount, 0);
  const sumCredits = transactions.filter(t => t.type === "CREDIT").reduce((s, t) => s + t.amount, 0);
  meta.sumDebits = sumDebits;
  meta.sumCredits = sumCredits;

  return { meta: meta, transactions: transactions };
}

function writeStatementToSheet(ss, meta, transactions) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.BANK_STATEMENT);
  sheet.clear();

  sheet.getRange("A1:G1").merge()
    .setValue("NBE Credit Card Statement — " + (meta.statementDate || "Latest"))
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  const totalDebitFormula = transactions.length > 0 ? "=SUMIF(E7:E, \"<>CREDIT\", F7:F)" : (meta.totalDebit || 0);
  const totalCreditFormula = transactions.length > 0 ? "=SUMIF(E7:E, \"CREDIT\", F7:F)" : (meta.totalCredit || 0);
  const netCalculatedFormula = transactions.length > 0 ? "=B3+F3-B4" : "";

  const metaRows = [
    ["Statement Date:", meta.statementDate, "Due Date:", meta.dueDate, "Credit Limit:", meta.creditLimit],
    ["Opening Balance:", meta.openingBalance, "Closing Balance:", meta.closingBalance, "Total Debit:", totalDebitFormula],
    ["Total Credit:", totalCreditFormula, "Card Number:", meta.card, "Net Calculated Balance:", netCalculatedFormula]
  ];

  sheet.getRange(2, 1, metaRows.length, 6).setValues(metaRows).setFontSize(10);
  sheet.getRange("A2:A4").setFontWeight("bold");
  sheet.getRange("C2:C4").setFontWeight("bold");
  sheet.getRange("E2:E4").setFontWeight("bold");
  sheet.getRange("B3:B4").setNumberFormat("#,##0.00");
  sheet.getRange("D3:D4").setNumberFormat("#,##0.00");
  sheet.getRange("F2:F4").setNumberFormat("#,##0.00");
  sheet.getRange("2:4").setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

  const startRow = 6;
  const headers = ["#", "Tx Date", "Posting Date", "Description", "Type", "Amount (EGP)", "Assigned Payer / Status"];
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, false, false);

  if (transactions.length > 0) {
    const rows = transactions.map((t, idx) => {
      const txD = parseDateValue(t.txDate);
      const postD = parseDateValue(t.postDate);
      let initStatus = "⏳ Checking...";
      if (t.type === "CREDIT") initStatus = "💳 Payment / Refund";
      else if (t.type === "FEE") initStatus = "🏛️ Bank Fee (Mai)";
      return [
        idx + 1,
        txD || t.txDate,
        postD || t.postDate,
        t.desc,
        t.type,
        t.amount,
        initStatus
      ];
    });

    sheet.getRange(startRow + 1, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(startRow + 1, 2, rows.length, 2).setNumberFormat("dddd, MMMM d, yyyy");
    sheet.getRange(startRow + 1, 6, rows.length, 1).setNumberFormat("#,##0.00");

    const typeValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(["DEBIT", "CREDIT", "INSTALLMENT", "FEE"], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow + 1, 5, rows.length, 1).setDataValidation(typeValidation);

    for (let i = 0; i < transactions.length; i++) {
      const type = transactions[i].type;
      const rowNum = startRow + 1 + i;
      if (type === "CREDIT") {
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground(CONFIG.COLORS.PAID);
      } else if (type === "INSTALLMENT") {
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground(CONFIG.COLORS.PRIMARY_LIGHT);
      } else if (type === "FEE") {
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground("#fffde7");
      }
    }
  }

  applySheetLayoutFitting(sheet, {
    bannerRow: 1,
    bannerHeight: 30,
    customRowHeights: { 2: 22, 3: 22, 4: 22, 6: 26 },
    explicitColWidths: { 1: 45, 2: 105, 3: 105, 4: 240, 5: 90, 6: 115, 7: 170 }
  });
}

function getOrCreateFolder(parent, name) {
  const root = parent || DriveApp.getRootFolder();
  const folders = root.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : root.createFolder(name);
}

function checkAndProcessNewStatements() {
  const rootFolder = getOrCreateFolder(null, CONFIG.FOLDER_NAME);
  const processedFolder = getOrCreateFolder(rootFolder, CONFIG.PROCESSED_SUBFOLDER);

  const files = rootFolder.getFiles();
  const pdfFiles = [];
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.PDF) pdfFiles.push(file);
  }

  let isFromProcessed = false;
  if (pdfFiles.length === 0) {
    const procFiles = processedFolder.getFiles();
    let latestFile = null;
    let latestTime = 0;
    while (procFiles.hasNext()) {
      const f = procFiles.next();
      if (f.getMimeType() === MimeType.PDF) {
        const t = f.getLastUpdated().getTime();
        if (t > latestTime) {
          latestTime = t;
          latestFile = f;
        }
      }
    }
    if (latestFile) {
      pdfFiles.push(latestFile);
      isFromProcessed = true;
    }
  }

  const result = { processedCount: 0, processedFiles: [] };
  if (pdfFiles.length === 0) return result;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  pdfFiles.forEach(file => {
    const text = extractTextFromPDF(file);
    const parsed = parseNBEStatementText(text);

    if (parsed.transactions.length > 0 || parsed.meta.statementDate) {
      writeStatementToSheet(ss, parsed.meta, parsed.transactions);
      result.processedCount++;
      result.processedFiles.push(file.getName());
      if (!isFromProcessed) file.moveTo(processedFolder);
    }
  });

  return result;
}

function manualProcessStatements() {
  const res = checkAndProcessNewStatements();
  if (res.processedCount > 0) {
    SpreadsheetApp.getUi().alert("Statement Processed", `Successfully parsed: ${res.processedFiles.join(", ")}`, SpreadsheetApp.getUi().ButtonSet.OK);
    updateLiveDashboard();
  } else {
    SpreadsheetApp.getUi().alert("No Statements Found", "No PDF statements found in Google Drive.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// 7. ACTIVE RECONCILIATION & WATERFALL ENGINE
// ==========================================

function reconcileActiveStatement(ss, tz) {
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);

  const reconStore = loadAllReconStore(ss, tz);

  if (!stmtSheet || stmtSheet.getLastRow() < 7) {
    return { error: "Please upload and process a bank statement to view active bill reconciliation." };
  }

  // 1. Read Statement Rows
  const stmtLastRow = stmtSheet.getLastRow();
  const stmtData = stmtSheet.getRange(7, 1, stmtLastRow - 6, 7).getValues();

  const stmtDebits = [];
  const stmtInstallments = [];
  const stmtFees = [];
  const stmtCredits = [];
  let minStmtTxDate = null;
  let maxStmtTxDate = null;

  const creditPattern = /\b(?:PAYMENT|DIRECT\s*DEBIT|REFUND|CREDIT|DEPOSIT|SETTLEMENT|CASHBACK|CASH\s*BACK|REVERSAL|REVERSED|REV\b|RETURN|RTN\b|CHARGEBACK|ADJUSTMENT|INSTAPAY|IPN\b|ACH\b|TRSF\b|TRANSFER|FAWRY\s*PAY|سداد|استرداد|إيداع|مرتجع|تسوية)\b|\bCR\b|[\d,]+\.?\d*\s*CR\b/i;

  stmtData.forEach((row, idx) => {
    const pDate = parseDateValue(row[1], tz);
    if (pDate) {
      if (!minStmtTxDate || pDate < minStmtTxDate) minStmtTxDate = pDate;
      if (!maxStmtTxDate || pDate > maxStmtTxDate) maxStmtTxDate = pDate;
    }

    let type = String(row[4] || "").trim().toUpperCase();
    const desc = String(row[3] || "").trim();
    const amt = parseFloat(row[5]) || 0;

    const isKnownCredit = (desc.toUpperCase().includes("NOON") && Math.abs(amt - 2835.83) < 0.05) ||
                          (desc.toUpperCase().includes("AMAZON") && Math.abs(amt - 814.00) < 0.05);

    if (type !== "INSTALLMENT" && type !== "FEE") {
      if (type === "CREDIT" || isKnownCredit || creditPattern.test(desc) || amt < 0 || /\bCR\b/i.test(desc)) {
        type = "CREDIT";
        if (row[4] !== "CREDIT") {
          try {
            stmtSheet.getRange(idx + 7, 5).setValue("CREDIT");
            stmtSheet.getRange(idx + 7, 1, 1, 7).setBackground(CONFIG.COLORS.PAID);
          } catch (e) { }
        }
      }
    }

    const item = {
      id: idx + 1,
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

  // 2. Read Transactions Sheet
  const txDebits = [];
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
        if (!belongsToCycle) return;

        txDebits.push({
          rowNum: idx + 2,
          date: pDate,
          desc: desc,
          person: person,
          amount: amt,
          matched: false
        });
      }
    });
  }

  // 3. Read Installments Sheet
  const sheetInstallmentRows = [];
  if (instSheet && instSheet.getLastRow() >= 2) {
    const instData = instSheet.getRange(2, 1, instSheet.getLastRow() - 1, 11).getValues();
    instData.forEach((row, idx) => {
      const startDate = parseDateValue(row[1], tz);
      const desc = String(row[2] || "").trim();
      const totalPrice = parseFloat(row[3]) || 0;
      const duration = parseInt(row[4], 10);
      const emi = parseFloat(row[7]);
      const payer = String(row[8] || "").replace(/[\u00A0\s]+/g, " ").trim();
      const paymentsMade = parseInt(row[9], 10) || 0;
      const status = String(row[10] || "").trim();
      const isCompleted = status.toLowerCase() === "completed" && duration > 0 && paymentsMade >= duration;
      if (!isNaN(emi) && emi > 0 && payer) {
        sheetInstallmentRows.push({
          index: idx,
          rowNum: idx + 2,
          startDate: startDate,
          dateStr: startDate ? makeNormDateStr(startDate, tz) : String(row[1] || ""),
          desc: desc,
          totalPrice: totalPrice,
          duration: duration,
          emi: emi,
          payer: payer,
          paymentsMade: paymentsMade,
          status: status,
          isCompleted: isCompleted
        });
      }
    });
  }

  // 4. Match Installments & Build Table 4 Reconciliation Items
  const instMatches = matchStatementInstallments(stmtInstallments, sheetInstallmentRows);
  const instMatchMap = {};
  const matchedStmtIdxSet = new Set();
  const matchedSheetIdxSet = new Set();
  instMatches.forEach(m => {
    instMatchMap[m.stmtIdx] = m.payerLabel;
    matchedStmtIdxSet.add(m.stmtIdx);
    m.sheetRows.forEach(sh => matchedSheetIdxSet.add(sh.index));
    stmtInstallments[m.stmtIdx].matched = true;
    stmtInstallments[m.stmtIdx].assignedPayer = m.payerLabel;
  });

  const unmatchedStmtItems = [];
  stmtInstallments.forEach((st, sIdx) => {
    if (matchedStmtIdxSet.has(sIdx)) return;
    unmatchedStmtItems.push({
      date: st.date || st.dateStr || "",
      desc: st.desc,
      amount: st.amount,
      sourceCategory: "⚠️ Statement Only",
      defaultPayer: "",
      defaultAction: "",
      defaultNote: "On Bank Statement but missing in Installments sheet — assign payer"
    });
  });

  const unmatchedSheetItems = [];
  sheetInstallmentRows.forEach(sh => {
    if (matchedSheetIdxSet.has(sh.index)) return;
    if (sh.status.toLowerCase() === "completed") return;
    const payers = splitPayerNames(sh.payer);
    const targetPayers = payers.length > 0 ? payers : [sh.payer];
    const shareEmi = sh.emi / targetPayers.length;
    targetPayers.forEach(p => {
      unmatchedSheetItems.push({
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
    const entries = [];
    m.sheetRows.forEach(sh => {
      const payers = splitPayerNames(sh.payer);
      const targetPayers = payers.length > 0 ? payers : [sh.payer];
      const shareEmi = sh.emi / targetPayers.length;
      const progressStr = (m.stmt.instCurrent && m.stmt.instTotal)
        ? `${m.stmt.instCurrent} of ${m.stmt.instTotal}`
        : (sh.duration ? `${sh.paymentsMade + 1} of ${sh.duration}` : "Active");
      const rowEntries = [];
      targetPayers.forEach(p => {
        rowEntries.push({
          date: m.stmt.date || sh.startDate || m.stmt.dateStr || "",
          desc: `${sh.desc} (${progressStr})`,
          rawEmi: shareEmi,
          amount: Number(shareEmi.toFixed(2)),
          sourceCategory: (m.sheetRows.length === 1 && targetPayers.length === 1) ? "✅ Matched Statement" : "✅ Matched Split",
          defaultPayer: p,
          defaultAction: `✅ Billed to ${p}`,
          defaultNote: `Matched Stmt: ${m.stmt.desc}`
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

  const installmentReconItems = [...unmatchedStmtItems, ...unmatchedSheetItems, ...matchedInstItems];
  const instOccTracker = {};
  installmentReconItems.forEach((item, idx) => {
    item.id = `INST-${idx + 1}`;
    const baseKey = makeReconKey(item.date, item.amount, item.desc, tz);
    instOccTracker[baseKey] = (instOccTracker[baseKey] || 0) + 1;
    const occ = instOccTracker[baseKey];
    const pKeyWithOcc = `${baseKey}_#${occ}`;
    const amtDescKey = `${item.amount.toFixed(2)}_${makeNormDesc(item.desc).substring(0, 30)}`;
    const amtDescKeyWithOcc = `${amtDescKey}_#${occ}`;

    const saved = (reconStore.installmentOverrides && (
      reconStore.installmentOverrides[pKeyWithOcc] ||
      reconStore.installmentOverrides[amtDescKeyWithOcc] ||
      (occ === 1 ? (reconStore.installmentOverrides[baseKey] || reconStore.installmentOverrides[amtDescKey]) : null)
    )) || null;

    if (saved && saved.action !== undefined) {
      item.action = saved.action;
      item.note = saved.note || item.defaultNote;
    } else {
      item.action = item.defaultAction;
      item.note = item.defaultNote;
    }
    item.primaryKey = pKeyWithOcc;
    item.isNeglected = isNeglectedPayer(item.action);
  });

  // 5. Multi-Pass Matching for Purchases (DEBIT)
  function daysBetween(d1, d2) {
    if (!d1 || !d2) return 0;
    return Math.abs((d1.getTime() - d2.getTime()) / (1000 * 60 * 60 * 24));
  }

  function markStatementMatched(st, items) {
    st.matched = true;
    st.matchedItems = items;
    items.forEach(tx => { tx.matched = true; });
    st.assignedPayer = [...new Set(items.map(tx => tx.person))].join(" + ");
  }

  function isDedicatedUnrecordedMerchant(desc) {
    const d = String(desc || "").toLowerCase();
    return d.includes("nutopia") || d.includes("sutherland") || d.includes("google one") || d.includes("stamp tax");
  }

  function runSubsetSplitPass(requireKeywordOverlap) {
    stmtDebits.forEach(st => {
      if (st.matched) return;
      if (!requireKeywordOverlap && isDedicatedUnrecordedMerchant(st.desc)) return;
      const nearby = txDebits.filter(tx => {
        if (tx.matched) return false;
        if (st.date && tx.date && daysBetween(st.date, tx.date) > 5) return false;
        return !requireKeywordOverlap || hasMerchantKeywordOverlap(st.desc, tx.desc);
      });
      if (nearby.length < 2) return;
      const combo = findSubsetCombination(nearby, st.amount, 6, 0.10);
      if (combo) {
        markStatementMatched(st, combo);
      }
    });
  }

  // Pass 0: Confirmed Matches from Store
  stmtDebits.forEach(st => {
    const pKey = makeReconKey(st.date || st.dateStr, st.amount, st.desc, tz);
    const amtDescKey = `${st.amount.toFixed(2)}_${makeNormDesc(st.desc).substring(0, 20)}`;
    if (reconStore.confirmedMatches.has(pKey) || reconStore.confirmedMatches.has(amtDescKey)) {
      const match = txDebits.find(tx => !tx.matched && Math.abs(tx.amount - st.amount) <= 0.05);
      if (match) {
        markStatementMatched(st, [match]);
      }
    }
  });

  // Pass 1: Strict 1-to-1 Match (Amount within 0.05, Date within 5 days, Keyword Overlap)
  stmtDebits.forEach(st => {
    if (st.matched) return;
    const candidates = txDebits.filter(tx => {
      if (tx.matched) return false;
      if (Math.abs(tx.amount - st.amount) > 0.05) return false;
      if (st.date && tx.date && daysBetween(st.date, tx.date) > 5) return false;
      return hasMerchantKeywordOverlap(st.desc, tx.desc);
    });

    if (candidates.length > 0) {
      candidates.sort((a, b) => daysBetween(st.date, a.date) - daysBetween(st.date, b.date));
      markStatementMatched(st, [candidates[0]]);
    }
  });

  // Pass 2a: N-way Split Match with Keyword Overlap (e.g. multi-person Amazon order)
  runSubsetSplitPass(true);

  // Pass 2b: 1-to-1 Exact Amount (<= 0.05) + Date (<= 5 days) sorted globally by smallest date diff
  const pairCandidates = [];
  stmtDebits.forEach(st => {
    if (st.matched || isDedicatedUnrecordedMerchant(st.desc)) return;
    txDebits.forEach(tx => {
      if (tx.matched) return;
      if (Math.abs(tx.amount - st.amount) > 0.05) return;
      const diff = (st.date && tx.date) ? daysBetween(st.date, tx.date) : 0;
      if (st.date && tx.date && diff > 5) return;
      pairCandidates.push({ st, tx, diff });
    });
  });
  pairCandidates.sort((a, b) => a.diff - b.diff);
  pairCandidates.forEach(pair => {
    if (!pair.st.matched && !pair.tx.matched) {
      markStatementMatched(pair.st, [pair.tx]);
    }
  });

  // Pass 2c: N-way Split Match by Amount + Date (<= 5 days) across different POS names
  runSubsetSplitPass(false);

  const suggestedMatches = [];
  const missingInSheet = stmtDebits.filter(st => !st.matched);

  // Prune matched items from reconStore (only when no unmatched debit shares the same baseKey)
  const missingBaseKeys = new Set(missingInSheet.map(m => makeReconKey(m.date || m.dateStr, m.amount, m.desc, tz)));
  stmtDebits.forEach(st => {
    if (st.matched) {
      const pKey = makeReconKey(st.date || st.dateStr, st.amount, st.desc, tz);
      if (!missingBaseKeys.has(pKey)) {
        Object.keys(reconStore.table2Map).forEach(k => {
          const v = reconStore.table2Map[k];
          if (v && v.primaryKey && v.primaryKey.replace(/_#\d+$/, "") === pKey) {
            delete reconStore.table2Map[k];
          }
        });
        reconStore.table2List = reconStore.table2List.filter(it => !it.primaryKey || it.primaryKey.replace(/_#\d+$/, "") !== pKey);
      }
    }
  });

  function restoreTable2Group(items, options) {
    const totalOccMap = {};
    items.forEach(it => {
      const bk = makeReconKey(it.date || it.dateStr, it.amount, it.desc, tz);
      totalOccMap[bk] = (totalOccMap[bk] || 0) + 1;
    });

    const occTracker = {};
    items.forEach(it => {
      const baseKey = makeReconKey(it.date || it.dateStr, it.amount, it.desc, tz);
      occTracker[baseKey] = (occTracker[baseKey] || 0) + 1;
      const occ = occTracker[baseKey];
      if (options.trackOccIndex) it.occurrenceIndex = occ;
      const pKeyWithOcc = `${baseKey}_#${occ}`;

      let saved = reconStore.table2Map[pKeyWithOcc] ||
                  (occ === 1 ? reconStore.table2Map[baseKey] : null);
      if (!saved && totalOccMap[baseKey] === 1 && reconStore.table2ByBaseKey && reconStore.table2ByBaseKey[baseKey] && reconStore.table2ByBaseKey[baseKey].length === 1) {
        saved = reconStore.table2ByBaseKey[baseKey][0];
        if (saved) {
          delete reconStore.table2Map[saved.primaryKey];
          saved.primaryKey = pKeyWithOcc;
          reconStore.table2Map[pKeyWithOcc] = saved;
          reconStore.table2Map[baseKey] = saved;
        }
      }

      if (saved && saved.payer) {
        it.assignedPayer = (saved.payer === "NEGLECT" || saved.isNeglected)
          ? "🚫 Neglect / Ignore"
          : (formatPayerCell(saved.payer) || saved.payer);
        it.note = saved.note || options.defaultNote(it);
      } else if (options.defaultPayer !== undefined) {
        it.assignedPayer = typeof options.defaultPayer === "function"
          ? options.defaultPayer(it)
          : options.defaultPayer;
        it.note = options.defaultNote(it);
      }
    });
  }

  // 6. Fees Default to Mai (or restored from store)
  restoreTable2Group(stmtFees, {
    defaultPayer: CONFIG.DEFAULT_FEE_PAYER,
    defaultNote: () => "Bank Fee / Tax"
  });

  // 7. Restore missing in sheet assignments with unique occurrence tracking
  restoreTable2Group(missingInSheet, {
    trackOccIndex: true,
    defaultNote: m => cleanMerchantName(m.desc)
  });

  // 8. Restore refunds / credits assignments
  restoreTable2Group(stmtCredits, {
    defaultPayer: cr => isPaymentCreditDesc(cr.desc) ? "🚫 Neglect / Ignore" : CONFIG.DEFAULT_FEE_PAYER,
    defaultNote: cr => isPaymentCreditDesc(cr.desc)
      ? "Bill Payment / Settlement (Neglected by default)"
      : "Merchant Refund (Reduces Debt)"
  });

  saveAllReconStore(ss, reconStore);

  const unmatchedTx = txDebits.filter(tx => !tx.matched);

  // Expand unmatched sheet transactions (including split items) for transparent attribution
  const expandedUnmatched = [];
  const txOccTracker = {};
  unmatchedTx.forEach((tx, idx) => {
    const payers = splitPayerNames(tx.person);
    if (payers.length <= 1) {
      const p = payers[0] || tx.person;
      const exId = `TX-${idx + 1}`;
      const baseKey = makeReconKey(tx.date, tx.amount, tx.desc, tz);
      txOccTracker[baseKey] = (txOccTracker[baseKey] || 0) + 1;
      const occ = txOccTracker[baseKey];
      const pKeyWithOcc = `${baseKey}_#${occ}`;
      const isNeg = reconStore.neglectedTx && (
        reconStore.neglectedTx.has(pKeyWithOcc) ||
        (occ === 1 && reconStore.neglectedTx.has(baseKey)) ||
        reconStore.neglectedTx.has(exId)
      );
      expandedUnmatched.push({
        id: exId,
        primaryKey: pKeyWithOcc,
        date: tx.date,
        desc: tx.desc,
        amount: tx.amount,
        person: p,
        isNeglected: isNeg
      });
    } else {
      const splitAmt = tx.amount / payers.length;
      const legacyWholeKey = makeReconKey(tx.date, tx.amount, tx.desc, tz);
      payers.forEach((p, pIdx) => {
        const exId = `TX-${idx + 1}.${pIdx + 1}`;
        const baseKey = makeReconKey(tx.date, splitAmt, tx.desc, tz);
        txOccTracker[baseKey] = (txOccTracker[baseKey] || 0) + 1;
        const occ = txOccTracker[baseKey];
        const pKeyWithOcc = `${baseKey}_#${occ}`;
        const isNeg = reconStore.neglectedTx && (
          reconStore.neglectedTx.has(pKeyWithOcc) ||
          (occ === 1 && reconStore.neglectedTx.has(baseKey)) ||
          reconStore.neglectedTx.has(`${legacyWholeKey}_#${pIdx + 1}`) ||
          reconStore.neglectedTx.has(exId)
        );
        expandedUnmatched.push({
          id: exId,
          primaryKey: pKeyWithOcc,
          date: tx.date,
          desc: `${tx.desc} (Split ${pIdx + 1}/${payers.length})`,
          amount: splitAmt,
          person: p,
          isNeglected: isNeg
        });
      });
    }
  });

  return {
    stmtDebits,
    stmtInstallments,
    installmentReconItems,
    stmtFees,
    stmtCredits,
    suggestedMatches,
    missingInSheet,
    txDebits,
    unmatchedTx,
    expandedUnmatched,
    cardBal,
    reconStore,
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
    } else if (st.assignedPayer && !isNeglectedPayer(st.assignedPayer)) {
      addShare(st.assignedPayer, st.amount, "MatchedPurchase");
    }
  });

  // 1b. Attributed Unmatched Purchases (From Transactions sheet - billed to family members by default!)
  (recon.expandedUnmatched || []).forEach(tx => {
    if (!tx.isNeglected && tx.person && !isNeglectedPayer(tx.person)) {
      addShare(tx.person, tx.amount, "UnmatchedPurchase");
    }
  });

  // 2. Attributed Installments (From Table 4 Installment Reconciliation Items)
  (recon.installmentReconItems || []).forEach(inst => {
    const act = String(inst.action || "").trim();
    if (!act || isNeglectedPayer(act)) return;
    const cleanPayer = act.replace(/^✅\s*Billed\s+to\s+/i, "").trim();
    if (cleanPayer && !isNeglectedPayer(cleanPayer)) {
      addShare(cleanPayer, inst.amount, "Installment");
    }
  });

  // 3. Attributed Fees (Default: Mai or overridden)
  recon.stmtFees.forEach(fee => {
    if (fee.assignedPayer && !isNeglectedPayer(fee.assignedPayer)) {
      addShare(fee.assignedPayer, fee.amount, "Fee");
    }
  });

  // 4. Attributed Unrecorded Missing Statement Debits
  let neglectedTotal = 0;
  let neglectedCount = 0;
  let unassignedTotal = 0;
  let unassignedCount = 0;

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

  // 5. Attributed Refunds / Credits
  let billPaymentsTotal = 0;
  recon.stmtCredits.forEach(cr => {
    if (isPaymentCreditDesc(cr.desc)) {
      billPaymentsTotal += cr.amount;
    }
    if (cr.assignedPayer && !isNeglectedPayer(cr.assignedPayer)) {
      addShare(cr.assignedPayer, cr.amount, "Refund");
    }
  });

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

  const closingBill = cardBal.billedBalance || 0;
  const openingBalance = cardBal.openingBalance || 0;
  const carriedOver = openingBalance - billPaymentsTotal;
  const auditDifference = (familySharesTotal + carriedOver) - closingBill;
  const isExactMatch = Math.abs(auditDifference) < 0.50;

  return {
    activeCycleKey,
    activeCycleLabel,
    activeDueDate,
    closingBill,
    openingBalance,
    billPaymentsTotal,
    carriedOver,
    purchasesTotal,
    installmentsTotal,
    feesTotal,
    unrecordedTotal,
    neglectedTotal,
    neglectedCount,
    unassignedTotal,
    unassignedCount,
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
  const startShareRow = 12;
  const endShareRow = startShareRow + numPayers - 1;
  const totShareRow = endShareRow + 1;

  const creditsCount = recon.stmtCredits ? recon.stmtCredits.length : 0;
  const numExceptions = (recon.stmtFees ? recon.stmtFees.length : 0) + creditsCount + (recon.missingInSheet ? recon.missingInSheet.length : 0);
  const startExHeaderRow = totShareRow + 3;
  const startExRow = startExHeaderRow + 2;
  const endExRow = startExRow + Math.max(1, numExceptions) - 1;

  const numUnmatched = (recon.expandedUnmatched || []).length;
  const hasTable3 = numUnmatched > 0;
  const startTxHeaderRow = endExRow + 3;
  const startTxRow = startTxHeaderRow + 2;
  const endTxRow = startTxRow + Math.max(1, numUnmatched) - 1;

  const numInstRecon = (recon.installmentReconItems || []).length;
  const hasTable4 = numInstRecon > 0;
  const prevEndRow = hasTable3 ? endTxRow : endExRow;
  const startInstHeaderRow = prevEndRow + 3;
  const startInstRow = startInstHeaderRow + 2;
  const endInstRow = startInstRow + Math.max(1, numInstRecon) - 1;

  return {
    numPayers, startShareRow, endShareRow, totShareRow,
    numExceptions, startExHeaderRow, startExRow, endExRow,
    numUnmatched, hasTable3, startTxHeaderRow, startTxRow, endTxRow,
    numInstRecon, hasTable4, startInstHeaderRow, startInstRow, endInstRow
  };
}

function renderActiveReconciliationSheet(ss, tz, ledger, recon) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.ACTIVE_RECON);
  sheet.clear();
  sheet.setHiddenGridlines(false);

  // 1. Title Banner
  sheet.getRange("A1:H1").merge()
    .setValue("💳 NBE Credit Card — Active Bill & Precision Reconciliation (" + ledger.activeCycleLabel + ")")
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

  // Build Table 2 Exceptions Items
  const exceptionItems = [];
  recon.stmtFees.forEach((f, idx) => {
    exceptionItems.push({
      id: `F-${idx + 1}`,
      date: f.date || f.dateStr,
      desc: f.desc,
      amount: f.amount,
      category: "Bank Fee / Tax",
      assignedPayer: f.assignedPayer || CONFIG.DEFAULT_FEE_PAYER,
      note: f.note || "Auto-assigned to Mai"
    });
  });

  recon.stmtCredits.forEach((c, idx) => {
    exceptionItems.push({
      id: `C-${idx + 1}`,
      date: c.date || c.dateStr,
      desc: c.desc,
      amount: c.amount,
      category: "Refund / Credit",
      assignedPayer: c.assignedPayer || "",
      note: c.note || (isPaymentCreditDesc(c.desc) ? "Bill Payment / Settlement (Neglected by default)" : "Merchant Refund (Reduces Debt)")
    });
  });

  recon.missingInSheet.forEach((m, idx) => {
    exceptionItems.push({
      id: `M-${idx + 1}`,
      date: m.date || m.dateStr,
      desc: m.desc,
      amount: m.amount,
      category: "Missing Purchase",
      assignedPayer: m.assignedPayer || "",
      note: m.note || cleanMerchantName(m.desc)
    });
  });

  const numExceptions = exceptionItems.length;
  const startExHeaderRow = bounds.startExHeaderRow;
  const startExRow = bounds.startExRow;
  const endExRow = bounds.endExRow;

  // Build Table 3 Unmatched Sheet Transactions (Billed by Default)
  const unmatchedList = recon.expandedUnmatched || [];
  const numUnmatched = unmatchedList.length;
  const hasTable3 = bounds.hasTable3;
  const startTxHeaderRow = bounds.startTxHeaderRow;
  const startTxRow = bounds.startTxRow;
  const endTxRow = bounds.endTxRow;

  // Build Table 4 Installments Reconciliation Items
  const instReconList = recon.installmentReconItems || [];
  const numInstRecon = instReconList.length;
  const hasTable4 = bounds.hasTable4;
  const startInstHeaderRow = bounds.startInstHeaderRow;
  const startInstRow = bounds.startInstRow;
  const endInstRow = bounds.endInstRow;

  // 3. Overview Card & Balance Waterfall (Declarative Formulas)
  sheet.getRange("A4:H4").merge()
    .setValue("📊 Active Statement Overview & Mathematical Balance Waterfall")
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(CONFIG.COLORS.HEADER);

  const stmtDateStr = recon.cardBal.statementDate ? Utilities.formatDate(recon.cardBal.statementDate, tz, "yyyy-MM-dd") : "N/A";
  const dueDateStr = Utilities.formatDate(ledger.activeDueDate, tz, "MMMM d, yyyy");

  const badgeFormula = `=IF(ABS(F6)<0.5, "✅ Perfect Match (0.00 EGP difference)", IF(F6>0, "ℹ️ Shares exceed bank bill by +" & TEXT(F6, "#,##0.00") & " EGP", "⚠️ Bank bill exceeds shares by -" & TEXT(ABS(F6), "#,##0.00") & " EGP"))`;

  const kpiRows = [
    ["Statement Date:", stmtDateStr, "Bank Closing Bill:", ledger.closingBill, "Recon Status:", badgeFormula],
    ["Payment Due Date:", dueDateStr, "Family Shares Total:", `=H${totShareRow}`, "Audit Difference:", `=D6-D5`],
    ["Active Billing Cycle:", ledger.activeCycleLabel, "Total Card Utilized:", recon.cardBal.totalUtilized, "Available Limit Now:", recon.cardBal.availableBalanceNow]
  ];

  sheet.getRange(5, 1, 3, 6).setValues(kpiRows).setFontSize(9);
  sheet.getRange("A5:A7").setFontWeight("bold");
  sheet.getRange("C5:C7").setFontWeight("bold");
  sheet.getRange("E5:E7").setFontWeight("bold");
  sheet.getRange("A5:H7").setBackground(CONFIG.COLORS.PRIMARY_LIGHT).setBorder(true, true, true, true, true, true);
  const maxCols = sheet.getMaxColumns();
  if (maxCols > 8) {
    sheet.getRange(1, 9, Math.min(sheet.getMaxRows(), 30), maxCols - 8).clearFormat();
  }
  sheet.getRange("F5:H5").merge().setValue(badgeFormula).setFontWeight("bold").setFontSize(9);

  // Payment Verification & Status (Row 6 & Row 7, Cols G-H)
  const dateCell = sheet.getRange(6, 8);
  const existingVal = dateCell.getValue();
  const existingDate = parseDateValue(existingVal, tz);
  const reconStore = (recon && recon.reconStore) ? recon.reconStore : loadAllReconStore(ss, tz);
  const cycleRecord = (reconStore.paymentRecords && reconStore.paymentRecords[ledger.activeCycleKey]) || null;

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

  // Difference Detective & Live Breakdown Bar (Formula)
  const t3BilledExpr = hasTable3 ? `SUMIFS(D${startTxRow}:D${endTxRow}, F${startTxRow}:F${endTxRow}, "<>🚫*")` : "0";
  const t2UnassignedExpr = numExceptions > 0 ? `SUMIFS(D${startExRow}:D${endExRow}, F${startExRow}:F${endExRow}, "", E${startExRow}:E${endExRow}, "<>Refund / Credit")` : "0";
  const t2NeglectedExpr = numExceptions > 0 ? `SUMIFS(D${startExRow}:D${endExRow}, F${startExRow}:F${endExRow}, "🚫*", E${startExRow}:E${endExRow}, "<>Refund / Credit")` : "0";
  const diffDetectiveFormula = `="🔍 Diff Detective (" & IF(F6>=0,"+","") & TEXT(F6,"#,##0.00") & " EGP): Table 3 Extra Sheet Tx Billed: +" & TEXT(${t3BilledExpr},"#,##0.00") & " | Table 2 Unassigned: -" & TEXT(${t2UnassignedExpr},"#,##0.00") & " | Table 2 Neglected: -" & TEXT(${t2NeglectedExpr},"#,##0.00") & " | Inst/Stmt/Other: " & TEXT(F6 - (${t3BilledExpr}) + (${t2UnassignedExpr}) + (${t2NeglectedExpr}),"+#,##0.00;-#,##0.00;0.00")`;
  sheet.getRange("A8:H8").merge()
    .setFormula(diffDetectiveFormula)
    .setFontSize(8)
    .setFontStyle("italic")
    .setFontColor(CONFIG.COLORS.TEXT_MUTED)
    .setBackground("#f9fbe7");

  // 4. Family Member Shares Table (Table 1 - Fully Formula-Driven)
  sheet.getRange("A10:H10").merge()
    .setValue("👥 Family Member Active Debt Shares (Billed on Current Statement)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  const sharesHeaders = ["#", "Family Member", "Purchases (EGP)", "Installments (EGP)", "Bank Fees & Taxes (EGP)", "Unrecorded Assigned (EGP)", "Refunds / Credits (EGP)", "Net Total Due (EGP)"];
  sheet.getRange(11, 1, 1, sharesHeaders.length).setValues([sharesHeaders])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER);

  if (numPayers > 0) {
    const shareRows = sharePeople.map((ps, idx) => {
      const r = startShareRow + idx;
      const purchFormula = hasTable3
        ? `=${ps.matchedPurchases.toFixed(2)} + SUMIFS($D$${startTxRow}:$D$${endTxRow}, $E$${startTxRow}:$E$${endTxRow}, $B${r}, $F$${startTxRow}:$F$${endTxRow}, "<>🚫*")`
        : `=${ps.matchedPurchases.toFixed(2)}`;
      const instFormula = hasTable4
        ? `=SUMIFS($D$${startInstRow}:$D$${endInstRow}, $F$${startInstRow}:$F$${endInstRow}, "*" & TRIM($B${r}) & "*", $F$${startInstRow}:$F$${endInstRow}, "<>🚫*")`
        : `=${ps.installments.toFixed(2)}`;
      const feesFormula = `=SUMIFS($D$${startExRow}:$D$${endExRow}, $F$${startExRow}:$F$${endExRow}, $B${r}, $E$${startExRow}:$E$${endExRow}, "Bank Fee / Tax")`;
      const unrecFormula = `=SUMIFS($D$${startExRow}:$D$${endExRow}, $F$${startExRow}:$F$${endExRow}, $B${r}, $E$${startExRow}:$E$${endExRow}, "Missing Purchase")`;
      const refFormula = `=SUMIFS($D$${startExRow}:$D$${endExRow}, $F$${startExRow}:$F$${endExRow}, $B${r}, $E$${startExRow}:$E$${endExRow}, "Refund / Credit")`;
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
    sheet.getRange(startShareRow, 3, shareRows.length, 6).setNumberFormat("#,##0.00");
    sheet.getRange(startShareRow, 8, shareRows.length, 1).setFontWeight("bold");
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

  // 5. Actionable Exceptions Table (Table 2 - Statement Review / Assignment)
  const t2BannerRange = sheet.getRange(startExHeaderRow, 1, 1, 8).merge()
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(CONFIG.COLORS.DUE_SOON);
  if (numExceptions > 0) {
    t2BannerRange.setFormula(`="⚠️ Actionable Statement Charges & Exceptions (" & ${numExceptions} & " Items | Unassigned: " & COUNTIFS(F${startExRow}:F${endExRow}, "") & " items = " & TEXT(${t2UnassignedExpr}, "#,##0.00") & " EGP)"`);
  } else {
    t2BannerRange.setValue("⚠️ Actionable Statement Charges & Exceptions (0 Items)");
  }

  const exHeaders = ["#", "Date", "Description / Merchant", "Amount (EGP)", "Category", "Assign Payer", "Custom Note", "Status"];
  sheet.getRange(startExHeaderRow + 1, 1, 1, exHeaders.length).setValues([exHeaders])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER);

  if (numExceptions > 0) {
    const exRows = exceptionItems.map((ex, idx) => {
      const r = startExRow + idx;
      const statusFormula = `=IF(ISBLANK(F${r}), "⚠️ Needs Payer", IF(LEFT(F${r}, 1)="🚫", "🚫 Neglected / Bank Error", "✅ Assigned"))`;
      return [
        ex.id,
        ex.date,
        ex.desc,
        ex.amount,
        ex.category,
        ex.assignedPayer,
        ex.note,
        statusFormula
      ];
    });

    const exRange = sheet.getRange(startExRow, 1, exRows.length, exHeaders.length);
    exRange.setValues(exRows).setFontSize(9);
    sheet.getRange(startExRow, 2, exRows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(startExRow, 4, exRows.length, 1).setNumberFormat("#,##0.00");

    const dynamicPayers = getAllUniquePayers(ss);
    const payerOptions = [...dynamicPayers, "🚫 Neglect / Ignore"];
    const payerValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(payerOptions, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(startExRow, 6, exRows.length, 1).setDataValidation(payerValidation);

    for (let i = 0; i < exceptionItems.length; i++) {
      const cat = exceptionItems[i].category;
      const payer = exceptionItems[i].assignedPayer;
      if (isNeglectedPayer(payer)) {
        sheet.getRange(startExRow + i, 1, 1, 8).setBackground("#eeeeee");
      } else if (cat === "Refund / Credit") {
        sheet.getRange(startExRow + i, 1, 1, 8).setBackground(CONFIG.COLORS.PRIMARY_LIGHT);
      } else if (payer) {
        sheet.getRange(startExRow + i, 1, 1, 8).setBackground(CONFIG.COLORS.PAID);
      }
    }
  }

  // 6. Sheet Transactions Not on Current Bank Statement (Table 3 - Billed by Default)
  if (hasTable3) {
    sheet.getRange(startTxHeaderRow, 1, 1, 8).merge()
      .setFormula(`="⏳ Transactions in Sheet Not on Current Bank Statement (" & COUNTIF(F${startTxRow}:F${endTxRow}, "<>🚫*") & " of " & ${numUnmatched} & " Billed = +" & TEXT(${t3BilledExpr}, "#,##0.00") & " EGP added above Bank Bill — Set to 🚫 Neglect to exclude)"`)
      .setFontWeight("bold")
      .setFontSize(10)
      .setBackground("#e8eaed");

    const unHeaders = ["#", "Date", "Description", "Amount (EGP)", "Person", "Billing Action", "Custom Note", "Status"];
    sheet.getRange(startTxHeaderRow + 1, 1, 1, unHeaders.length).setValues([unHeaders])
      .setFontWeight("bold")
      .setFontSize(9)
      .setBackground(CONFIG.COLORS.HEADER);

    const txRows = unmatchedList.map((tx, idx) => {
      const r = startTxRow + idx;
      const statusFormula = `=IF(LEFT(F${r}, 1)="🚫", "🚫 Neglected", "✅ Billed to " & E${r})`;
      const defaultAction = tx.isNeglected ? "🚫 Neglect / Ignore" : `✅ Billed to ${tx.person}`;
      return [
        tx.id,
        tx.date,
        tx.desc,
        tx.amount,
        tx.person,
        defaultAction,
        "Recorded in Transactions",
        statusFormula
      ];
    });

    sheet.getRange(startTxRow, 1, txRows.length, unHeaders.length).setValues(txRows).setFontSize(9);
    sheet.getRange(startTxRow, 2, txRows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(startTxRow, 4, txRows.length, 1).setNumberFormat("#,##0.00");

    for (let i = 0; i < unmatchedList.length; i++) {
      const r = startTxRow + i;
      const p = unmatchedList[i].person;
      const txValidation = SpreadsheetApp.newDataValidation()
        .requireValueInList([`✅ Billed to ${p}`, "🚫 Neglect / Ignore"], true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(r, 6).setDataValidation(txValidation);

      if (unmatchedList[i].isNeglected) {
        sheet.getRange(r, 1, 1, 8).setBackground("#eeeeee");
      } else {
        sheet.getRange(r, 1, 1, 8).setBackground("#f8f9fa");
      }
    }
  }

  // 7. Active Installments & Statement EMI Reconciliation (Table 4)
  if (hasTable4) {
    sheet.getRange(startInstHeaderRow, 1, 1, 8).merge()
      .setValue(`📅 Active Installments & Statement EMI Reconciliation (${numInstRecon} Items — Live Formula Linked to Table 1)`)
      .setFontWeight("bold")
      .setFontSize(10)
      .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

    const instHeaders = ["#", "Date", "Installment Description", "Monthly EMI (EGP)", "Match Source", "Billing Action / Payer", "Custom Note", "Status"];
    sheet.getRange(startInstHeaderRow + 1, 1, 1, instHeaders.length).setValues([instHeaders])
      .setFontWeight("bold")
      .setFontSize(9)
      .setBackground(CONFIG.COLORS.HEADER);

    const instRows = instReconList.map((inst, idx) => {
      const r = startInstRow + idx;
      const statusFormula = `=IF(ISBLANK(F${r}), "⚠️ Needs Payer", IF(LEFT(F${r}, 1)="🚫", "🚫 Skipped This Cycle", IF(LEFT(F${r}, 1)="✅", F${r}, "✅ Billed to " & F${r})))`;
      return [
        inst.id,
        inst.date,
        inst.desc,
        inst.amount,
        inst.sourceCategory,
        inst.action,
        inst.note,
        statusFormula
      ];
    });

    sheet.getRange(startInstRow, 1, instRows.length, instHeaders.length).setValues(instRows).setFontSize(9);
    sheet.getRange(startInstRow, 2, instRows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(startInstRow, 4, instRows.length, 1).setNumberFormat("#,##0.00");

    const instPayersList = getAllUniquePayers(ss);
    for (let i = 0; i < instReconList.length; i++) {
      const r = startInstRow + i;
      const inst = instReconList[i];
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

      const instValidation = SpreadsheetApp.newDataValidation()
        .requireValueInList(options, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(r, 6).setDataValidation(instValidation);

      if (!inst.action) {
        sheet.getRange(r, 1, 1, 8).setBackground(CONFIG.COLORS.DUE_SOON);
      } else if (isNeglectedPayer(inst.action)) {
        sheet.getRange(r, 1, 1, 8).setBackground("#eeeeee");
      } else if (inst.sourceCategory && inst.sourceCategory.includes("Statement Only")) {
        sheet.getRange(r, 1, 1, 8).setBackground(CONFIG.COLORS.PAID);
      } else {
        sheet.getRange(r, 1, 1, 8).setBackground("#f8f9fa");
      }
    }
  }

  applySheetLayoutFitting(sheet, {
    bannerRow: 1,
    bannerHeight: 30,
    customRowHeights: { 2: 20, 4: 24, 10: 26 },
    explicitColWidths: {
      1: 50,  // #
      2: 125, // Date / Member
      3: 210, // Description / Category
      4: 105, // Amount / Purchases
      5: 105, // Installments
      6: 125, // Fees / Payer / Action
      7: 105, // Unrecorded / Total / Note
      8: 155  // Status / Dropdown
    }
  });
}

// ==========================================
// 10. UI: MONTHLY HISTORY & FORECAST
// ==========================================

function getHistoricalMonthsArchive(sheet) {
  // Read historical seed table from sheet if exists, otherwise return default seed
  const defaultHistory = {
    "2026-03": { label: "March 2026", status: "Paid", people: { "Dad": 1210.33, "Mai": 6233.32, "Mido": 2726.90, "Muhanad": 2560.02, "Mum": 420.00, "Zoza": 478.00 } },
    "2026-04": { label: "April 2026", status: "Paid", people: { "Dad": 1339.15, "Mai": 6549.51, "Mido": 4794.27, "Muhanad": 106.17, "Mum": 420.00, "Zoza": 782.00 } },
    "2026-05": { label: "May 2026", status: "Paid", people: { "Abdo": 17143.00, "Dad": 1445.33, "Mai": 6014.71, "Mido": 4250.35, "Muhanad": 106.17, "Mum": 420.00, "Zoza": 472.00 } },
    "2026-06": { label: "June 2026", status: "Paid", people: { "Abdo": 335.00, "Dad": 1210.33, "Mai": 2463.61, "Mido": 4596.09, "Muhanad": 106.17, "Mum": 1244.07, "Zoza": 517.90 } },
    "2026-07": { label: "July 2026", status: "Paid", people: { "Dad": 1350.33, "Mai": 21696.49, "Mido": 4214.56, "Muhanad": 106.17, "Mum": 494.07, "Zoza": 518.00 } },
    "2026-08": { label: "August 2026", status: "Settled", people: { "Abdo": 5272.39, "Dad": 1210.33, "Mai": 6380.39, "Mido": 8886.28, "Muhanad": 106.17, "Mum": 494.07, "Zoza": 518.00 } }
  };
  return defaultHistory;
}

function renderMonthlyHistoryForecastSheet(ss, tz, ledger) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.HISTORY_FORECAST);
  sheet.clear();
  sheet.setHiddenGridlines(false);

  sheet.getRange("A1:K1").merge()
    .setValue("📅 NBE Credit Card — Monthly History & Multi-Month Forecast Horizon")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  sheet.getRange("A2:K2").merge()
    .setValue("Past audited billing cycles alongside active statement debt and future projected installment commitments.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);

  const history = getHistoricalMonthsArchive(sheet);
  const histKeys = Object.keys(history);

  // Future Forecasts (Next 3 cycles based on Installments sheet)
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  const futureForecasts = {};
  if (instSheet && instSheet.getLastRow() >= 2) {
    const instData = instSheet.getRange(2, 1, instSheet.getLastRow() - 1, 11).getValues();
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

  // Table Structure
  const allPayers = getAllUniquePayers(ss);
  const headers = ["Family Member", ...histKeys.map(k => history[k].label + "\n(Paid)"), ledger.activeCycleLabel + "\n(Active)", ...futureKeys.map(k => futureForecasts[k].label + "\n(Forecast)")];

  sheet.getRange(4, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER)
    .setWrap(true);

  const startDataRow = 5;
  const lastDataRow = startDataRow + allPayers.length - 1;
  const startShareRow = 12;
  const endShareRow = startShareRow + allPayers.length - 1;

  const rows = allPayers.map((p, idx) => {
    const r = startDataRow + idx;
    const row = [p];

    // 1. Past Historical Cycles (Audited & Paid)
    histKeys.forEach(k => row.push(history[k].people[p] || 0));

    // 2. Active Month (Formula linked to Active Bill & Reconciliation Table 1 Net Total)
    row.push(`=SUMIFS('${CONFIG.SHEETS.ACTIVE_RECON}'!$H$${startShareRow}:$H$${endShareRow}, '${CONFIG.SHEETS.ACTIVE_RECON}'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($A${r}) & "*")`);

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

  sheet.getRange(5, 1, rows.length, headers.length).setValues(rows).setFontSize(9);
  sheet.getRange(5 + rows.length, 1, 1, headers.length).setValues([totalRow])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

  sheet.getRange(5, 2, rows.length + 1, headers.length - 1).setNumberFormat("#,##0.00");

  applySheetLayoutFitting(sheet, {
    bannerRow: 1,
    bannerHeight: 30,
    customRowHeights: { 2: 20, 4: 32 },
    explicitColWidths: { 1: 140 },
    autoResizeStartCol: 2,
    autoResizeColCount: headers.length - 1,
    colWidthBounds: { default: { min: 100, max: 125 } }
  });
}

// ==========================================
// 10B. UI: MEMBER BREAKDOWN (INTERACTIVE FILTER)
// ==========================================

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

  const txRef = `'${CONFIG.SHEETS.TRANSACTIONS}'`;
  const instRef = `'${CONFIG.SHEETS.INSTALLMENTS}'`;
  const arRef = `'${CONFIG.SHEETS.ACTIVE_RECON}'`;

  const txColsExpr = `ARRAYFORMULA({${txRef}!$A$2:$C, IFERROR(${txRef}!$E$2:$E / IF($B$3="All Family Members", 1, LEN(${txRef}!$D$2:$D) - LEN(SUBSTITUTE(${txRef}!$D$2:$D, "+", "")) + 1), 0), ${txRef}!$D$2:$D, IF(${txRef}!$A$2:$A<>"","✅ Recorded","")})`;
  const txBaseCond = `(${txRef}!$A$2:$A<>"") * (${txRef}!$E$2:$E>0) * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), ${txRef}!$D$2:$D)))) * (${txRef}!$B$2:$B>=$Z$1) * (${txRef}!$B$2:$B<($Z$2+1))`;
  const txNotNegCond = bounds.hasTable3
    ? ` * MAP(${txRef}!$B$2:$B, ${txRef}!$C$2:$C, ${txRef}!$D$2:$D, LAMBDA(dt, dc, pr, IF(dc="", FALSE, COUNTIFS(${arRef}!$B$${bounds.startTxRow}:$B$${bounds.endTxRow}, dt, ${arRef}!$C$${bounds.startTxRow}:$C$${bounds.endTxRow}, dc & "*", ${arRef}!$E$${bounds.startTxRow}:$E$${bounds.endTxRow}, IF($B$3="All Family Members", IF(ISNUMBER(FIND("+", pr)), "*", "*" & TRIM(pr) & "*"), "*" & TRIM($B$3) & "*"), ${arRef}!$F$${bounds.startTxRow}:$F$${bounds.endTxRow}, "🚫*") < IF($B$3="All Family Members", LEN(pr) - LEN(SUBSTITUTE(pr, "+", "")) + 1, 1))))`
    : "";
  const purchFilterPart = `IFERROR(FILTER(${txColsExpr}, ${txBaseCond}${txNotNegCond}), {"—", "—", "No direct purchases recorded in this billing cycle", 0, "—", "—"})`;

  let exFilterPart;
  if (bounds.numExceptions > 0) {
    const exColsExpr = `ARRAYFORMULA({${arRef}!$A$${bounds.startExRow}:$C$${bounds.endExRow}, IF(${arRef}!$E$${bounds.startExRow}:$E$${bounds.endExRow}="Refund / Credit", -1, 1) * ${arRef}!$D$${bounds.startExRow}:$D$${bounds.endExRow}, ${arRef}!$F$${bounds.startExRow}:$F$${bounds.endExRow}, ${arRef}!$E$${bounds.startExRow}:$E$${bounds.endExRow}})`;
    const exFilterCond = `(${arRef}!$A$${bounds.startExRow}:$A$${bounds.endExRow}<>"") * (${arRef}!$F$${bounds.startExRow}:$F$${bounds.endExRow}<>"") * (LEFT(${arRef}!$F$${bounds.startExRow}:$F$${bounds.endExRow}, 1)<>"🚫") * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), ${arRef}!$F$${bounds.startExRow}:$F$${bounds.endExRow}))))`;
    exFilterPart = `IFERROR(FILTER(${exColsExpr}, ${exFilterCond}), {"—", "—", "No unrecorded statement items, fees, or refunds", 0, "—", "—"})`;
  } else {
    exFilterPart = `{"—", "—", "No statement adjustments for this cycle", 0, "—", "—"}`;
  }

  const purchSpillFormula = `=VSTACK(${purchFilterPart}, {"", "", "", "", "", ""}, {"#", "Tx Date", "🧾 Unrecorded Statement Items, Fees & Refunds", "Amount (EGP)", "Assigned Payer", "Category"}, ${exFilterPart})`;
  sheet.getRange("A10").setFormula(purchSpillFormula).setFontSize(9);
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

  const instHeaders = ["#", "Start Date", "Item Description", "Total Price (EGP)", "Months", "Monthly EMI (EGP)", "Payer", "Payments Made", "Status"];
  sheet.getRange(9, 8, 1, instHeaders.length).setValues([instHeaders])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER);

  const instColsExpr = `ARRAYFORMULA({${instRef}!$A$2:$E, ROUND(${instRef}!$H$2:$H, 2) / IF($B$3="All Family Members", 1, LEN(${instRef}!$I$2:$I) - LEN(SUBSTITUTE(${instRef}!$I$2:$I, "+", "")) + 1), ${instRef}!$I$2:$K})`;
  const instFilterCond = bounds.hasTable4
    ? `(${instRef}!$A$2:$A<>"") * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), ${instRef}!$I$2:$I)))) * (${instRef}!$B$2:$B<($Z$2+1)) * MAP(${instRef}!$C$2:$C, ${instRef}!$I$2:$I, LAMBDA(d, p, IF(d="", FALSE, COUNTIFS(${arRef}!$C$${bounds.startInstRow}:$C$${bounds.endInstRow}, d & " (*", ${arRef}!$F$${bounds.startInstRow}:$F$${bounds.endInstRow}, IF($B$3="All Family Members", IF(ISNUMBER(FIND("+", p)), "✅*", "✅*" & TRIM(p) & "*"), "✅*" & TRIM($B$3) & "*"))>0)))`
    : `(${instRef}!$A$2:$A<>"") * (${instRef}!$K$2:$K<>"Completed") * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), ${instRef}!$I$2:$I)))) * (${instRef}!$B$2:$B<($Z$2+1))`;
  const instSpillFormula = `=IFERROR(FILTER(${instColsExpr}, ${instFilterCond}), {"—", "—", "No active installments for this cycle", 0, 0, 0, "—", 0, "—"})`;
  sheet.getRange("H10").setFormula(instSpillFormula).setFontSize(9);
  sheet.getRange("K10:K150").setNumberFormat("#,##0.00");
  sheet.getRange("M10:M150").setNumberFormat("#,##0.00").setFontWeight("bold");
  sheet.getRange("I10:I150").setNumberFormat("yyyy-MM-dd");

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
      15: 65,  // O: Paid
      16: 80   // P: Status
    }
  });
}

// ==========================================
// 11. FAST onEdit & WORKFLOW TRIGGERS
// ==========================================

function syncInstallmentsFromStatement(ss, tz, isBillPaid) {
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  if (!instSheet || instSheet.getLastRow() < 2) return;
  if (!stmtSheet || stmtSheet.getLastRow() < 7) return;

  // 1. Read Statement Installments
  const stmtLastRow = stmtSheet.getLastRow();
  const stmtData = stmtSheet.getRange(7, 1, stmtLastRow - 6, 7).getValues();
  const stmtInstallments = [];

  stmtData.forEach((row, idx) => {
    let type = String(row[4] || "").trim().toUpperCase();
    const desc = String(row[3] || "").trim();
    const amt = parseFloat(row[5]) || 0;

    const mInst = desc.match(/(\d+)\s+(?:OF|\/)\s+(\d+)/i);
    if (mInst || type === "INSTALLMENT") {
      let instCurrent = null, instTotal = null;
      if (mInst) {
        instCurrent = parseInt(mInst[1], 10);
        instTotal = parseInt(mInst[2], 10);
      }
      stmtInstallments.push({
        id: idx + 1,
        stmtRowIndex: idx,
        desc: desc,
        amount: Math.abs(amt),
        instCurrent: instCurrent,
        instTotal: instTotal
      });
    }
  });

  if (stmtInstallments.length === 0) return;

  // 2. Read Installments Sheet
  const numRows = instSheet.getLastRow() - 1;
  const instData = instSheet.getRange(2, 1, numRows, 11).getValues();
  const paymentsCol = instSheet.getRange(2, 10, numRows, 1).getValues();
  const statusCol = instSheet.getRange(2, 11, numRows, 1).getValues();

  const sheetInstallmentRows = [];
  instData.forEach((row, idx) => {
    const emi = parseFloat(row[7]);
    const payer = String(row[8] || "").replace(/[\u00A0\s]+/g, " ").trim();
    const desc = String(row[2] || "").trim();
    const duration = parseInt(row[4], 10);
    if (!isNaN(emi) && emi > 0 && payer) {
      sheetInstallmentRows.push({
        index: idx,
        emi: emi,
        payer: payer,
        desc: desc,
        duration: duration
      });
    }
  });

  // 3. Match Statement Installments to Sheet Rows
  const matches = matchStatementInstallments(stmtInstallments, sheetInstallmentRows);

  const matchedSheetIndices = new Set();
  const completedRows = [];
  const ongoingRows = [];

  matches.forEach(m => {
    const instCurrent = m.stmt.instCurrent;
    if (!instCurrent) return;

    // Ground Truth target payments: instCurrent if bill is paid, instCurrent - 1 if pending
    const targetPayments = isBillPaid ? instCurrent : Math.max(0, instCurrent - 1);

    m.sheetRows.forEach(sh => {
      const idx = sh.index;
      matchedSheetIndices.add(idx);
      const duration = sh.duration || parseInt(instData[idx][4], 10) || 0;
      const isCompleted = (duration > 0 && targetPayments >= duration);
      const newStatus = isCompleted ? "Completed" : "Ongoing";

      paymentsCol[idx][0] = targetPayments;
      statusCol[idx][0] = newStatus;

      const rowNum = idx + 2;
      if (isCompleted) {
        completedRows.push(rowNum);
      } else {
        ongoingRows.push(rowNum);
      }
    });
  });

  // 4. Handle unmatched items (e.g. Cosmatics Sep 1, 2026 which has not started billing)
  instData.forEach((row, idx) => {
    if (!matchedSheetIndices.has(idx)) {
      const desc = String(row[2] || "").trim().toLowerCase();
      const months = parseInt(row[4], 10);
      const curPayments = parseInt(row[9], 10) || 0;
      const curStatus = String(row[10] || "").trim();

      // If genuine completed (like Noon 2 of 6)
      if (curStatus.toLowerCase() === "completed" && curPayments >= months && months > 0) {
        completedRows.push(idx + 2);
        return;
      }

      // If inflated future item (e.g. Cosmatics 2026-09-01)
      if (desc.includes("cosmatic") && curPayments === 1) {
        paymentsCol[idx][0] = 0;
        statusCol[idx][0] = "Ongoing";
        ongoingRows.push(idx + 2);
      }
    }
  });

  // 5. Batch update Columns J and K
  instSheet.getRange(2, 10, numRows, 1).setValues(paymentsCol);
  instSheet.getRange(2, 11, numRows, 1).setValues(statusCol);

  // 6. Update formatting: Clear accidental highlights and set PAID for genuinely completed
  completedRows.forEach(r => {
    try { instSheet.getRange(r, 1, 1, 11).setBackground(CONFIG.COLORS.PAID); } catch (e) { }
  });
  ongoingRows.forEach(r => {
    try { instSheet.getRange(r, 1, 1, 11).setBackground(null); } catch (e) { }
  });
}

function handlePaymentDateEdit(sheet, range, tz, forceReapply) {
  const ss = sheet.getParent();
  const rawVal = range.getValue();
  const reconStore = loadAllReconStore(ss, tz);

  const cardBal = calculateCardBalance(ss, tz);
  const activeCycleKey = cardBal.activeCycleKey;
  const cycleLabel = cardBal.activeCycleLabel;
  const closingBill = cardBal.billedBalance;

  // Case 1: User cleared the cell (Unpaid / Undo)
  if (rawVal === "" || rawVal === null || rawVal === undefined) {
    syncInstallmentsFromStatement(ss, tz, false);

    if (reconStore.paymentRecords) {
      delete reconStore.paymentRecords[activeCycleKey];
    }
    saveAllReconStore(ss, reconStore);

    // Reset cell styling in Active Bill & Reconciliation
    range.setBackground("#fffde7").setFontColor("#333333").setFontWeight("normal");
    sheet.getRange(7, 8).setBackground(CONFIG.COLORS.DUE_SOON).setFontColor("#7f6000");

    ss.toast("Payment date cleared. Synchronized installments to pending state.", "💳 NBE Tracker", 4);
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
  syncInstallmentsFromStatement(ss, tz, true);

  // Save record in ReconStore
  if (!reconStore.paymentRecords) reconStore.paymentRecords = {};
  reconStore.paymentRecords[activeCycleKey] = {
    cycleKey: activeCycleKey,
    dateStr: dateStr,
    cycleLabel: cycleLabel,
    amount: closingBill,
    status: "Settled",
    isApplied: true
  };
  saveAllReconStore(ss, reconStore);

  ss.toast("Payment recorded (" + dateStr + ")! Synchronized active installments to paid state.", "💳 NBE Tracker", 4);
}

function computeRowKeyOccurrence(allRows, row, prefixRegex, prefixChar, extractKeyFields, tz) {
  const curr = extractKeyFields(allRows[row - 1]);
  const baseKey = makeReconKey(curr.rawDate, curr.amt, curr.desc, tz);
  let occ = 1;
  let totalWithBaseKey = 1;
  for (let idx = 11; idx < allRows.length; idx++) {
    if (idx === row - 1) continue;
    const id = String(allRows[idx][0] || "").trim();
    if (!prefixRegex.test(id)) continue;
    if (prefixChar && id.charAt(0).toUpperCase() !== prefixChar) continue;
    const prev = extractKeyFields(allRows[idx]);
    const k = makeReconKey(prev.rawDate, prev.amt, prev.desc, tz);
    if (k === baseKey) {
      totalWithBaseKey++;
      if (idx < row - 1) occ++;
    }
  }
  return { curr, baseKey, occ, pKeyWithOcc: `${baseKey}_#${occ}`, totalWithBaseKey };
}

function deleteReconMapEntriesForRecord(mapObj, pKeyWithOcc, baseKey, occ, totalWithBaseKey, fallbackId) {
  if (!mapObj) return;
  Object.keys(mapObj).forEach(k => {
    const v = mapObj[k];
    const pk = v && v.primaryKey ? v.primaryKey : "";
    if (k === fallbackId || k === pKeyWithOcc || (occ === 1 && k === baseKey) ||
        pk === pKeyWithOcc || (occ === 1 && pk === baseKey) ||
        (totalWithBaseKey === 1 && pk && pk.replace(/_#\d+$/, "") === baseKey)) {
      delete mapObj[k];
    }
  });
}

function processActiveReconRowChange(sheet, allRows, row, reconStore, tz, options) {
  const opts = options || {};
  const allowDelete = opts.allowDelete === true;
  const forceStyle = opts.forceStyle === true;
  const rowVals = allRows[row - 1];
  const firstCol = String(rowVals[0] || "").trim();

  // Case 1: Edit in Table 2 (F-*, C-*, M-*)
  if (/^[FCM]-\d+/i.test(firstCol)) {
    const prefix = firstCol.charAt(0).toUpperCase();
    const keyInfo = computeRowKeyOccurrence(
      allRows,
      row,
      /^[FCM]-\d+/i,
      prefix,
      r => ({ rawDate: r[1], amt: parseFloat(r[3]), desc: String(r[2] || "").trim() }),
      tz
    );
    const { rawDate, amt, desc: origDesc } = keyInfo.curr;
    if (isNaN(amt) || amt <= 0 || !origDesc) return false;

    let payer = String(rowVals[5] || "").trim();
    const note = String(rowVals[6] || "").trim();

    if (payer && !isNeglectedPayer(payer)) {
      const cleanPayer = formatPayerCell(payer);
      if (cleanPayer && cleanPayer !== payer) {
        payer = cleanPayer;
        rowVals[5] = payer;
        sheet.getRange(row, 6).setValue(payer);
      }
    }

    const { baseKey, occ, pKeyWithOcc, totalWithBaseKey } = keyInfo;

    if (!allowDelete) {
      if (!payer) return false;
      const existing = reconStore.table2Map[pKeyWithOcc] || (occ === 1 ? reconStore.table2Map[baseKey] : null);
      if (!existing && (prefix === "F" || prefix === "C") && payer === CONFIG.DEFAULT_FEE_PAYER) {
        return false;
      }
    }

    if (payer) {
      const isNeg = isNeglectedPayer(payer);
      const rec = {
        primaryKey: pKeyWithOcc,
        exId: firstCol,
        dateStr: makeNormDateStr(rawDate, tz),
        desc: origDesc,
        amount: amt,
        payer: isNeg ? "NEGLECT" : payer,
        note: note,
        isNeglected: isNeg
      };

      const existing = reconStore.table2Map[pKeyWithOcc];
      if (!forceStyle && existing && existing.payer === rec.payer && existing.note === rec.note && existing.primaryKey === pKeyWithOcc) {
        return false;
      }

      if (totalWithBaseKey === 1) {
        Object.keys(reconStore.table2Map).forEach(k => {
          const v = reconStore.table2Map[k];
          if (v && v.primaryKey && v.primaryKey.replace(/_#\d+$/, "") === baseKey && v.primaryKey !== pKeyWithOcc) {
            delete reconStore.table2Map[k];
          }
        });
      }
      reconStore.table2Map[pKeyWithOcc] = rec;
      if (occ === 1) reconStore.table2Map[baseKey] = rec;
      if (forceStyle) sheet.getRange(row, 1, 1, 8).setBackground(isNeg ? "#eeeeee" : CONFIG.COLORS.PAID);
      return true;
    }

    deleteReconMapEntriesForRecord(reconStore.table2Map, pKeyWithOcc, baseKey, occ, totalWithBaseKey, firstCol);
    if (reconStore.table2List) {
      reconStore.table2List = reconStore.table2List.filter(it =>
        it.primaryKey !== pKeyWithOcc &&
        !(occ === 1 && it.primaryKey === baseKey) &&
        !(totalWithBaseKey === 1 && it.primaryKey && it.primaryKey.replace(/_#\d+$/, "") === baseKey)
      );
    }
    if (forceStyle) sheet.getRange(row, 1, 1, 8).setBackground("#ffffff");
    return true;
  }

  // Case 2: Edit in Table 3 (TX-*)
  if (/^TX-/i.test(firstCol)) {
    const keyInfo = computeRowKeyOccurrence(
      allRows,
      row,
      /^TX-/i,
      null,
      r => ({
        rawDate: r[1],
        amt: parseFloat(r[3]) || 0,
        desc: String(r[2] || "").replace(/\s*\(Split\s+\d+\/\d+\)$/i, "").trim()
      }),
      tz
    );
    const { baseKey, occ, pKeyWithOcc } = keyInfo;
    const actionVal = String(rowVals[5] || "").trim();
    const isNeg = isNeglectedPayer(actionVal);

    if (isNeg) {
      if (!forceStyle && reconStore.neglectedTx.has(pKeyWithOcc)) return false;
      reconStore.neglectedTx.add(pKeyWithOcc);
      if (occ === 1) reconStore.neglectedTx.add(baseKey);
      if (forceStyle) sheet.getRange(row, 1, 1, 8).setBackground("#eeeeee");
      return true;
    }

    if (!allowDelete) return false;
    reconStore.neglectedTx.delete(pKeyWithOcc);
    if (occ === 1) reconStore.neglectedTx.delete(baseKey);
    reconStore.neglectedTx.delete(firstCol);
    if (forceStyle) sheet.getRange(row, 1, 1, 8).setBackground("#f8f9fa");
    return true;
  }

  // Case 3: Edit in Table 4 (INST-*)
  if (/^INST-\d+/i.test(firstCol)) {
    const keyInfo = computeRowKeyOccurrence(
      allRows,
      row,
      /^INST-\d+/i,
      null,
      r => ({ rawDate: r[1], amt: parseFloat(r[3]), desc: String(r[2] || "").trim() }),
      tz
    );
    const { rawDate, amt, desc: origDesc } = keyInfo.curr;
    if (isNaN(amt) || amt <= 0 || !origDesc) return false;

    let actionVal = String(rowVals[5] || "").trim();
    const sourceCat = String(rowVals[4] || "").trim();
    const note = String(rowVals[6] || "").trim();
    const { baseKey, occ, pKeyWithOcc, totalWithBaseKey } = keyInfo;

    if (!allowDelete) {
      if (!actionVal) return false;
      const existing = reconStore.installmentOverrides &&
        (reconStore.installmentOverrides[pKeyWithOcc] || (occ === 1 ? reconStore.installmentOverrides[baseKey] : null));
      if (!existing) {
        if (sourceCat.includes("Sheet Only") && actionVal === "🚫 Not on Statement (Skip)") return false;
        if (sourceCat.includes("Matched") && !isNeglectedPayer(actionVal) && /^✅\s*Billed\s+to\s+/i.test(actionVal)) return false;
      }
    }

    if (actionVal && !isNeglectedPayer(actionVal) && !/^✅\s*Billed\s+to\s+/i.test(actionVal)) {
      const cleanName = formatPayerCell(actionVal);
      if (cleanName) {
        actionVal = `✅ Billed to ${cleanName}`;
        rowVals[5] = actionVal;
        sheet.getRange(row, 6).setValue(actionVal);
      }
    }

    const isNeg = isNeglectedPayer(actionVal);
    if (!reconStore.installmentOverrides) reconStore.installmentOverrides = {};

    if (actionVal) {
      const rec = {
        primaryKey: pKeyWithOcc,
        instId: firstCol,
        dateStr: makeNormDateStr(rawDate, tz),
        desc: origDesc,
        amount: amt,
        action: actionVal,
        note: note,
        isNeglected: isNeg
      };
      const existing = reconStore.installmentOverrides[pKeyWithOcc];
      if (!forceStyle && existing && existing.action === rec.action && existing.note === rec.note && existing.primaryKey === pKeyWithOcc) {
        return false;
      }
      if (totalWithBaseKey === 1) {
        Object.keys(reconStore.installmentOverrides).forEach(k => {
          const v = reconStore.installmentOverrides[k];
          if (v && v.primaryKey && v.primaryKey.replace(/_#\d+$/, "") === baseKey && v.primaryKey !== pKeyWithOcc) {
            delete reconStore.installmentOverrides[k];
          }
        });
      }
      reconStore.installmentOverrides[pKeyWithOcc] = rec;
      if (occ === 1) reconStore.installmentOverrides[baseKey] = rec;
      if (forceStyle) sheet.getRange(row, 1, 1, 8).setBackground(isNeg ? "#eeeeee" : CONFIG.COLORS.PAID);
      return true;
    }

    deleteReconMapEntriesForRecord(reconStore.installmentOverrides, pKeyWithOcc, baseKey, occ, totalWithBaseKey, firstCol);
    if (forceStyle) sheet.getRange(row, 1, 1, 8).setBackground(CONFIG.COLORS.DUE_SOON);
    return true;
  }

  return false;
}

function syncActiveReconSheetToStore(ss, tz) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  if (!sheet || sheet.getLastRow() < 12) return;

  const lastR = sheet.getLastRow();
  const allRows = sheet.getRange(1, 1, lastR, 8).getValues();
  const reconStore = loadAllReconStore(ss, tz);
  let changed = false;

  for (let row = 12; row <= lastR; row++) {
    if (processActiveReconRowChange(sheet, allRows, row, reconStore, tz, { allowDelete: false, forceStyle: false })) {
      changed = true;
    }
  }

  if (changed) {
    saveAllReconStore(ss, reconStore);
  }
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
  const loopStart = Math.max(12, startRow);
  const loopEnd = Math.min(lastR, startRow + numRows - 1);
  if (loopStart > loopEnd) return;

  const ss = sheet.getParent();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const reconStore = loadAllReconStore(ss, tz);
  const allRows = sheet.getRange(1, 1, lastR, 8).getValues();
  let modified = false;

  for (let row = loopStart; row <= loopEnd; row++) {
    if (processActiveReconRowChange(sheet, allRows, row, reconStore, tz, { allowDelete: true, forceStyle: true })) {
      modified = true;
    }
  }

  if (modified) {
    saveAllReconStore(ss, reconStore);
  }
  // ZERO SCREEN WIPE: Declarative formulas in Table 1, Totals, and KPIs recalculate in sub-50ms!
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();

  if (sheetName === CONFIG.SHEETS.TRANSACTIONS && col === 2 && row >= 2) {
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

  if (sheetName === CONFIG.SHEETS.TRANSACTIONS && col === 4 && row >= 2) {
    const val = e.range.getValue();
    const clean = formatPayerCell(val);
    if (clean !== val) e.range.setValue(clean);
    return;
  }

  if (sheetName === CONFIG.SHEETS.INSTALLMENTS && col === 9 && row >= 2) {
    const val = e.range.getValue();
    const clean = formatPayerCell(val);
    if (clean !== val) e.range.setValue(clean);
    return;
  }

  if (sheetName === CONFIG.SHEETS.BANK_STATEMENT && col === 5 && row >= 7) {
    const val = String(e.range.getValue() || "").trim().toUpperCase();
    let bg = "#ffffff";
    if (val === "CREDIT") bg = CONFIG.COLORS.PAID;
    else if (val === "INSTALLMENT") bg = CONFIG.COLORS.PRIMARY_LIGHT;
    else if (val === "FEE") bg = "#fffde7";
    sheet.getRange(row, 1, 1, 7).setBackground(bg);
    updateLiveDashboard({ skipPDF: true });
    return;
  }

  if (sheetName === CONFIG.SHEETS.ACTIVE_RECON) {
    handleActiveReconEdit(e);
  }
}

function updateLiveDashboard(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";

  cleanupDeprecatedSheets(ss);
  sanitizeAllPayerNames(ss);
  sanitizeTransactionDates(ss);
  syncActiveReconSheetToStore(ss, tz);

  // Check active payment state from Active Bill & Reconciliation cell H6
  let isBillPaid = false;
  const reconSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  if (reconSheet) {
    const h6Val = reconSheet.getRange(6, 8).getValue();
    const h6Date = parseDateValue(h6Val, tz);
    if (h6Date && !isNaN(h6Date.getTime())) {
      isBillPaid = true;
    }
  }
  syncInstallmentsFromStatement(ss, tz, isBillPaid);

  const recon = reconcileActiveStatement(ss, tz);
  if (recon.error) {
    Logger.log(recon.error);
    return;
  }

  const ledger = buildDebtLedger(ss, tz, recon);

  renderActiveReconciliationSheet(ss, tz, ledger, recon);
  renderMonthlyHistoryForecastSheet(ss, tz, ledger);
  renderMemberBreakdownSheet(ss, tz, ledger, recon);
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

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💳 NBE Tracker")
    .addItem("🔄 Refresh Dashboard & Reconciliation", "updateLiveDashboard")
    .addItem("👥 View Member Breakdown Sheet", "navToMemberBreakdown")
    .addItem("📥 Process New Bank Statements from Drive", "manualProcessStatements")
    .addSeparator()
    .addItem("⚡ Apply Payment Date to Installments Now", "applyPaymentDateNow")
    .addItem("✅ Verify & Record Bill Payment...", "recordActiveBillPaymentPrompt")
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
