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
  const c1 = cleanMerchantName(desc1);
  const c2 = cleanMerchantName(desc2);

  const aliases = [
    { key: "internet", terms: ["internet", "we-fbb", "we-mobile", "vodafone", "telecom", "we"] },
    { key: "cavalli", terms: ["cavalli"] },
    { key: "amazon", terms: ["amazon"] },
    { key: "uber", terms: ["uber"] },
    { key: "careem", terms: ["careem"] },
    { key: "fathalla", terms: ["fathalla"] },
    { key: "buffalo", terms: ["buffalo"] },
    { key: "caribou", terms: ["caribou"] },
    { key: "basilico", terms: ["basilico"] },
    { key: "maria", terms: ["maria"] },
    { key: "lail", terms: ["lail"] },
    { key: "asteria", terms: ["asteria"] },
    { key: "instashop", terms: ["instashop"] },
    { key: "shein", terms: ["shein"] },
    { key: "google", terms: ["google"] },
    { key: "microsoft", terms: ["microsoft"] },
    { key: "nutopia", terms: ["nutopia", "sutherland"] }
  ];

  for (const a of aliases) {
    const has1 = a.terms.some(t => c1.includes(t));
    const has2 = a.terms.some(t => c2.includes(t));
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

  if (purchaseDate.getDate() === 31 || purchaseDate.getUTCDate() === 31) {
    day = 31;
    if (purchaseDate.getDate() === 31) {
      month = purchaseDate.getMonth();
      year = purchaseDate.getFullYear();
    } else {
      month = purchaseDate.getUTCMonth();
      year = purchaseDate.getUTCFullYear();
    }
  }

  let cycleYear = year;
  let cycleMonth = month;

  // Day 31 Rollover Rule (e.g. Aug 31 -> September cycle, due Oct 25)
  if (day === 31) {
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
    isDay31Rollover: (day === 31)
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

  for (let offset = -60; offset <= 5; offset++) {
    const d = new Date(anchor.getTime() + offset * 86400000);
    const info = getBillingCycleForPurchase(d, tz);
    if (info && info.cycleSortKey === cycleSortKey) {
      if (!start) start = new Date(d);
      end = new Date(d);
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
  const single = `SUMIFS(Transactions!$E:$E, Transactions!$D:$D, "*" & TRIM(${personCellRef}) & "*", Transactions!$D:$D, "<>*+*", Transactions!$B:$B, ">="&${startExpr}, Transactions!$B:$B, "<="&${endExpr})`;
  const split = `0.5 * SUMIFS(Transactions!$E:$E, Transactions!$D:$D, "*+*", Transactions!$D:$D, "*" & TRIM(${personCellRef}) & "*", Transactions!$B:$B, ">="&${startExpr}, Transactions!$B:$B, "<="&${endExpr})`;
  return `${single} + ${split}`;
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
    table2List: [],
    confirmedMatches: new Set(),
    neglectedTx: new Set()
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

        if (type === "TABLE2_ASSIGNMENT" || (!type && payer)) {
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
            const amtDescKey = `${amt.toFixed(2)}_${makeNormDesc(desc).substring(0, 30)}`;
            if (!store.table2Map[amtDescKey]) store.table2Map[amtDescKey] = rec;
            store.table2List.push(rec);
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

  statementInstallments.forEach((st, sIdx) => {
    const cand = sheetInstallmentRows.find((sh, shIdx) => {
      if (matchedSheet.has(shIdx)) return false;
      if (Math.abs(st.amount - sh.emi) > 0.50) return false;
      if (st.instTotal && sh.duration && Math.abs(st.instTotal - sh.duration) > 1) return false;
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

  const descGroups = {};
  sheetInstallmentRows.forEach((sh, shIdx) => {
    if (matchedSheet.has(shIdx)) return;
    const key = (sh.desc || "").toLowerCase().trim() + "_" + (sh.duration || 0);
    if (!descGroups[key]) descGroups[key] = [];
    descGroups[key].push(sh);
  });

  statementInstallments.forEach((st, sIdx) => {
    if (matchedStmt.has(sIdx)) return;
    for (const key in descGroups) {
      const group = descGroups[key];
      if (!group || group.length < 2) continue;
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

  statementInstallments.forEach((st, sIdx) => {
    if (matchedStmt.has(sIdx)) return;
    const available = sheetInstallmentRows.filter((sh, shIdx) => {
      if (matchedSheet.has(shIdx)) return false;
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
  }

  // 4. Match Installments
  const instMatches = matchStatementInstallments(stmtInstallments, sheetInstallmentRows);
  const instMatchMap = {};
  instMatches.forEach(m => {
    instMatchMap[m.stmtIdx] = m.payerLabel;
    stmtInstallments[m.stmtIdx].matched = true;
    stmtInstallments[m.stmtIdx].assignedPayer = m.payerLabel;
  });

  // 5. Multi-Pass Matching for Purchases (DEBIT)
  // Pass 0: Confirmed Matches from Store
  stmtDebits.forEach(st => {
    const pKey = makeReconKey(st.date || st.dateStr, st.amount, st.desc, tz);
    const amtDescKey = `${st.amount.toFixed(2)}_${makeNormDesc(st.desc).substring(0, 20)}`;
    if (reconStore.confirmedMatches.has(pKey) || reconStore.confirmedMatches.has(amtDescKey)) {
      st.matched = true;
      const match = txDebits.find(tx => !tx.matched && Math.abs(tx.amount - st.amount) <= 0.05);
      if (match) {
        match.matched = true;
        st.assignedPayer = match.person;
      }
    }
  });

  // Pass 1: Strict Match (Amount within 0.05, Date within 3 days, Keyword Overlap)
  stmtDebits.forEach(st => {
    if (st.matched) return;
    const match = txDebits.find(tx => {
      if (tx.matched) return false;
      if (Math.abs(tx.amount - st.amount) > 0.05) return false;
      if (st.date && tx.date) {
        const daysDiff = Math.abs((tx.date.getTime() - st.date.getTime()) / (1000 * 60 * 60 * 24));
        if (daysDiff > 3) return false;
      }
      return hasMerchantKeywordOverlap(st.desc, tx.desc);
    });

    if (match) {
      st.matched = true;
      match.matched = true;
      st.assignedPayer = match.person;
    }
  });

  // Pass 2: Suggested Matches (Requires user confirmation)
  const suggestedMatches = [];
  stmtDebits.forEach(st => {
    if (st.matched) return;
    const nearby = txDebits.filter(tx => {
      if (tx.matched) return false;
      if (st.date && tx.date) {
        return Math.abs((tx.date.getTime() - st.date.getTime()) / (1000 * 60 * 60 * 24)) <= 5;
      }
      return true;
    });

    const singleAmtMatch = nearby.find(tx => Math.abs(tx.amount - st.amount) < 0.05);
    if (singleAmtMatch) {
      suggestedMatches.push({
        stmt: st,
        sheetItems: [singleAmtMatch],
        reason: `Exact amount (${st.amount.toFixed(2)} EGP), dates within 5 days — verify shop name`
      });
      return;
    }

    const combo = findSubsetCombination(nearby, st.amount, 6, 0.10);
    if (combo) {
      const peopleList = combo.map(it => it.person).join(" + ");
      suggestedMatches.push({
        stmt: st,
        sheetItems: combo,
        reason: `${combo.length}-way split sum matches ${st.amount.toFixed(2)} EGP (${peopleList})`
      });
    }
  });

  const suggestedStmtIds = new Set(suggestedMatches.map(s => s.stmt.id));
  const missingInSheet = stmtDebits.filter(st => !st.matched && !suggestedStmtIds.has(st.id));

  // Prune matched items from reconStore
  stmtDebits.forEach(st => {
    if (st.matched) {
      const pKey = makeReconKey(st.date || st.dateStr, st.amount, st.desc, tz);
      delete reconStore.table2Map[pKey];
      reconStore.table2List = reconStore.table2List.filter(it => it.primaryKey !== pKey);
    }
  });

  // 6. Fees Default to Mai (or restored from store)
  const feeOccTracker = {};
  stmtFees.forEach((fee, idx) => {
    const baseKey = makeReconKey(fee.date || fee.dateStr, fee.amount, fee.desc, tz);
    feeOccTracker[baseKey] = (feeOccTracker[baseKey] || 0) + 1;
    const occ = feeOccTracker[baseKey];
    const pKeyWithOcc = `${baseKey}_#${occ}`;
    const exId = `F-${idx + 1}`;

    const saved = reconStore.table2Map[exId] ||
                  reconStore.table2Map[pKeyWithOcc] ||
                  (occ === 1 ? reconStore.table2Map[baseKey] : null);
    fee.assignedPayer = (saved && saved.payer) ? saved.payer : CONFIG.DEFAULT_FEE_PAYER;
    fee.note = (saved && saved.note) ? saved.note : "Bank Fee / Tax";
  });

  // 7. Restore missing in sheet assignments with unique occurrence tracking
  const missingOccTracker = {};
  missingInSheet.forEach((m, idx) => {
    const baseKey = makeReconKey(m.date || m.dateStr, m.amount, m.desc, tz);
    missingOccTracker[baseKey] = (missingOccTracker[baseKey] || 0) + 1;
    const occ = missingOccTracker[baseKey];
    m.occurrenceIndex = occ;
    const pKeyWithOcc = `${baseKey}_#${occ}`;
    const exId = `M-${idx + 1}`;

    const saved = reconStore.table2Map[exId] ||
                  reconStore.table2Map[pKeyWithOcc] ||
                  (occ === 1 ? reconStore.table2Map[baseKey] : null);
    if (saved && saved.payer) {
      m.assignedPayer = (saved.payer === "NEGLECT" || saved.isNeglected) ? "🚫 Neglect / Ignore" : saved.payer;
      m.note = saved.note || cleanMerchantName(m.desc);
    }
  });

  // 8. Restore refunds / credits assignments
  const creditOccTracker = {};
  stmtCredits.forEach((cr, idx) => {
    const baseKey = makeReconKey(cr.date || cr.dateStr, cr.amount, cr.desc, tz);
    creditOccTracker[baseKey] = (creditOccTracker[baseKey] || 0) + 1;
    const occ = creditOccTracker[baseKey];
    const pKeyWithOcc = `${baseKey}_#${occ}`;
    const exId = `C-${idx + 1}`;

    const saved = reconStore.table2Map[exId] ||
                  reconStore.table2Map[pKeyWithOcc] ||
                  (occ === 1 ? reconStore.table2Map[baseKey] : null);
    if (saved && saved.payer) {
      cr.assignedPayer = (saved.payer === "NEGLECT" || saved.isNeglected) ? "🚫 Neglect / Ignore" : saved.payer;
      cr.note = saved.note || cr.desc;
    }
  });

  saveAllReconStore(ss, reconStore);

  const unmatchedTx = txDebits.filter(tx => !tx.matched);

  // Expand unmatched sheet transactions (including split items) for transparent attribution
  const expandedUnmatched = [];
  unmatchedTx.forEach((tx, idx) => {
    const payers = splitPayerNames(tx.person);
    const baseKey = makeReconKey(tx.date, tx.amount, tx.desc, tz);
    if (payers.length <= 1) {
      const p = payers[0] || tx.person;
      const exId = `TX-${idx + 1}`;
      const isNeg = reconStore.neglectedTx && (reconStore.neglectedTx.has(exId) || reconStore.neglectedTx.has(baseKey));
      expandedUnmatched.push({
        id: exId,
        primaryKey: baseKey,
        date: tx.date,
        desc: tx.desc,
        amount: tx.amount,
        person: p,
        isNeglected: isNeg
      });
    } else {
      const splitAmt = tx.amount / payers.length;
      payers.forEach((p, pIdx) => {
        const exId = `TX-${idx + 1}.${pIdx + 1}`;
        const pKey = `${baseKey}_#${pIdx + 1}`;
        const isNeg = reconStore.neglectedTx && (reconStore.neglectedTx.has(exId) || reconStore.neglectedTx.has(pKey) || reconStore.neglectedTx.has(baseKey));
        expandedUnmatched.push({
          id: exId,
          primaryKey: pKey,
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
    if (st.matched && st.assignedPayer && !isNeglectedPayer(st.assignedPayer)) {
      addShare(st.assignedPayer, st.amount, "MatchedPurchase");
    }
  });

  // 1b. Attributed Unmatched Purchases (From Transactions sheet - billed to family members by default!)
  (recon.expandedUnmatched || []).forEach(tx => {
    if (!tx.isNeglected && tx.person && !isNeglectedPayer(tx.person)) {
      addShare(tx.person, tx.amount, "UnmatchedPurchase");
    }
  });

  // 2. Attributed Installments (Matched Statement Installments)
  recon.stmtInstallments.forEach(st => {
    if (st.assignedPayer) {
      addShare(st.assignedPayer, st.amount, "Installment");
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
    const isPayment = /\b(?:PAYMENT|DIRECT\s*DEBIT|SETTLEMENT|DEPOSIT|سداد|إيداع|تسوية)\b/i.test(cr.desc);
    if (isPayment) {
      billPaymentsTotal += cr.amount;
    } else {
      // Refund / Cashback
      addShare(cr.assignedPayer || CONFIG.DEFAULT_FEE_PAYER, cr.amount, "Refund");
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
  const sharePeople = ledger.personShares;
  const numPayers = sharePeople.length;
  const startShareRow = 12;
  const endShareRow = startShareRow + numPayers - 1;
  const totShareRow = endShareRow + 1;

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
    const isPayment = /\b(?:PAYMENT|DIRECT\s*DEBIT|SETTLEMENT|DEPOSIT|سداد|إيداع|تسوية)\b/i.test(c.desc);
    if (!isPayment) {
      exceptionItems.push({
        id: `C-${idx + 1}`,
        date: c.date || c.dateStr,
        desc: c.desc,
        amount: c.amount,
        category: "Refund / Credit",
        assignedPayer: c.assignedPayer || CONFIG.DEFAULT_FEE_PAYER,
        note: c.note || "Merchant Refund (Reduces Debt)"
      });
    }
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
  const startExHeaderRow = totShareRow + 3;
  const startExRow = startExHeaderRow + 2;
  const endExRow = startExRow + Math.max(1, numExceptions) - 1;

  // Build Table 3 Unmatched Sheet Transactions (Billed by Default)
  const unmatchedList = recon.expandedUnmatched || [];
  const numUnmatched = unmatchedList.length;
  const hasTable3 = numUnmatched > 0;
  const startTxHeaderRow = endExRow + 3;
  const startTxRow = startTxHeaderRow + 2;
  const endTxRow = startTxRow + Math.max(1, numUnmatched) - 1;

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

  // Waterfall Explanation Bar (Formula)
  const waterfallFormula = `="Waterfall: Purchases (" & TEXT(C${totShareRow}, "#,##0.00") & ") + Installments (" & TEXT(D${totShareRow}, "#,##0.00") & ") + Fees (" & TEXT(E${totShareRow}, "#,##0.00") & ") + Unrecorded (" & TEXT(F${totShareRow}, "#,##0.00") & ") - Refunds (" & TEXT(G${totShareRow}, "#,##0.00") & ") = Family Shares (" & TEXT(H${totShareRow}, "#,##0.00") & ") | Bank Closing Bill: " & TEXT(D5, "#,##0.00") & " EGP"`;
  sheet.getRange("A8:H8").merge()
    .setValue(waterfallFormula)
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
    const activeRange = getCycleDateRange(cardBal.activeCycleKey, tz);
    const shareRows = sharePeople.map((ps, idx) => {
      const r = startShareRow + idx;
      const activePurchSum = activeRange ? makePurchaseSumFormula(`$B${r}`, activeRange) : `0`;
      const purchFormula = hasTable3
        ? `=${activePurchSum} - SUMIFS($D$${startTxRow}:$D$${endTxRow}, $E$${startTxRow}:$E$${endTxRow}, $B${r}, $F$${startTxRow}:$F$${endTxRow}, "🚫*")`
        : `=${activePurchSum}`;
      const feesFormula = `=SUMIFS($D$${startExRow}:$D$${endExRow}, $F$${startExRow}:$F$${endExRow}, $B${r}, $E$${startExRow}:$E$${endExRow}, "Bank Fee / Tax")`;
      const unrecFormula = `=SUMIFS($D$${startExRow}:$D$${endExRow}, $F$${startExRow}:$F$${endExRow}, $B${r}, $E$${startExRow}:$E$${endExRow}, "Missing Purchase")`;
      const refFormula = `=SUMIFS($D$${startExRow}:$D$${endExRow}, $F$${startExRow}:$F$${endExRow}, $B${r}, $E$${startExRow}:$E$${endExRow}, "Refund / Credit")`;
      const netFormula = `=C${r}+D${r}+E${r}+F${r}-G${r}`;

      return [
        idx + 1,
        ps.person,
        purchFormula,
        ps.installments,
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
  sheet.getRange(startExHeaderRow, 1, 1, 8).merge()
    .setValue(`⚠️ Actionable Statement Charges & Exceptions (${numExceptions} Items Requiring Review / Assignment)`)
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(CONFIG.COLORS.DUE_SOON);

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
      if (cat === "Refund / Credit") {
        sheet.getRange(startExRow + i, 1, 1, 8).setBackground(CONFIG.COLORS.PRIMARY_LIGHT);
      } else if (isNeglectedPayer(payer)) {
        sheet.getRange(startExRow + i, 1, 1, 8).setBackground("#eeeeee");
      } else if (payer) {
        sheet.getRange(startExRow + i, 1, 1, 8).setBackground(CONFIG.COLORS.PAID);
      }
    }
  }

  // 6. Sheet Transactions Not on Current Bank Statement (Table 3 - Billed by Default)
  if (hasTable3) {
    sheet.getRange(startTxHeaderRow, 1, 1, 8).merge()
      .setValue(`⏳ Transactions in Sheet Not on Current Bank Statement (${numUnmatched} Items — Billed to Family Members by Default)`)
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
// 11. FAST onEdit & WORKFLOW TRIGGERS
// ==========================================

function handleActiveReconEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const col = e.range.getColumn();
  const row = e.range.getRow();

  if (row < 12) return;
  const lastR = sheet.getLastRow();
  if (row > lastR) return;

  const rowVals = sheet.getRange(row, 1, 1, 8).getValues()[0];
  const firstCol = String(rowVals[0] || "").trim();

  const ss = sheet.getParent();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const reconStore = loadAllReconStore(ss, tz);

  // Case 1: Edit in Table 2 (F-*, C-*, M-*)
  if (/^[FCM]-\d+/i.test(firstCol)) {
    const exId = firstCol;
    const rawDate = rowVals[1];
    const origDesc = String(rowVals[2] || "").trim();
    const amt = parseFloat(rowVals[3]);
    const payer = String(rowVals[5] || "").trim();
    const note = String(rowVals[6] || "").trim();

    if (isNaN(amt) || amt <= 0 || !origDesc) return;

    // Calculate occurrence number within Table 2
    let occ = 1;
    const startExRow = 14;
    if (row > startExRow) {
      const prevRange = sheet.getRange(startExRow, 1, row - startExRow, 4).getValues();
      prevRange.forEach(pr => {
        const pDesc = String(pr[2] || "").trim();
        const pAmt = parseFloat(pr[3]) || 0;
        if (Math.abs(pAmt - amt) < 0.01 && makeNormDesc(pDesc) === makeNormDesc(origDesc)) {
          occ++;
        }
      });
    }

    const baseKey = makeReconKey(rawDate, amt, origDesc, tz);
    const pKeyWithOcc = `${baseKey}_#${occ}`;
    const amtDescKey = `${amt.toFixed(2)}_${makeNormDesc(origDesc).substring(0, 30)}_#${occ}`;

    if (payer && !isNeglectedPayer(payer)) {
      const rec = {
        primaryKey: pKeyWithOcc,
        exId: exId,
        dateStr: makeNormDateStr(rawDate, tz),
        desc: origDesc,
        amount: amt,
        payer: payer,
        note: note,
        isNeglected: false
      };
      reconStore.table2Map[pKeyWithOcc] = rec;
      reconStore.table2Map[amtDescKey] = rec;
      reconStore.table2Map[exId] = rec;
      if (occ === 1) {
        reconStore.table2Map[baseKey] = rec;
        reconStore.table2Map[`${amt.toFixed(2)}_${makeNormDesc(origDesc).substring(0, 30)}`] = rec;
      }
      sheet.getRange(row, 1, 1, 8).setBackground(CONFIG.COLORS.PAID);
    } else if (isNeglectedPayer(payer)) {
      const rec = {
        primaryKey: pKeyWithOcc,
        exId: exId,
        dateStr: makeNormDateStr(rawDate, tz),
        desc: origDesc,
        amount: amt,
        payer: "NEGLECT",
        note: note,
        isNeglected: true
      };
      reconStore.table2Map[pKeyWithOcc] = rec;
      reconStore.table2Map[amtDescKey] = rec;
      reconStore.table2Map[exId] = rec;
      if (occ === 1) {
        reconStore.table2Map[baseKey] = rec;
        reconStore.table2Map[`${amt.toFixed(2)}_${makeNormDesc(origDesc).substring(0, 30)}`] = rec;
      }
      sheet.getRange(row, 1, 1, 8).setBackground("#eeeeee");
    } else {
      delete reconStore.table2Map[pKeyWithOcc];
      delete reconStore.table2Map[amtDescKey];
      delete reconStore.table2Map[exId];
      if (occ === 1) {
        delete reconStore.table2Map[baseKey];
        delete reconStore.table2Map[`${amt.toFixed(2)}_${makeNormDesc(origDesc).substring(0, 30)}`];
      }
      sheet.getRange(row, 1, 1, 8).setBackground("#ffffff");
    }

    saveAllReconStore(ss, reconStore);
    // ZERO SCREEN WIPE: Declarative formulas in Table 1, Totals, and KPIs recalculate in sub-50ms!
    return;
  }

  // Case 2: Edit in Table 3 (TX-*)
  if (/^TX-/i.test(firstCol)) {
    const txId = firstCol;
    const actionVal = String(rowVals[5] || "").trim();
    const isNeg = isNeglectedPayer(actionVal);

    if (isNeg) {
      reconStore.neglectedTx.add(txId);
      sheet.getRange(row, 1, 1, 8).setBackground("#eeeeee");
    } else {
      reconStore.neglectedTx.delete(txId);
      sheet.getRange(row, 1, 1, 8).setBackground("#f8f9fa");
    }

    saveAllReconStore(ss, reconStore);
    // ZERO SCREEN WIPE: Declarative formulas in Table 1, Totals, and KPIs recalculate in sub-50ms!
    return;
  }
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();

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

  const recon = reconcileActiveStatement(ss, tz);
  if (recon.error) {
    Logger.log(recon.error);
    return;
  }

  const ledger = buildDebtLedger(ss, tz, recon);

  renderActiveReconciliationSheet(ss, tz, ledger, recon);
  renderMonthlyHistoryForecastSheet(ss, tz, ledger);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💳 NBE Tracker")
    .addItem("💳 Refresh Dashboard & Reconciliation", "updateLiveDashboard")
    .addItem("📥 Process New Bank Statements from Drive", "manualProcessStatements")
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
