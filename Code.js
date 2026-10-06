/**
 * NBE Credit Card Automated Tracker & Reconciliation Engine
 * 
 * Features:
 * 1. Google Drive PDF Ingestion: Auto-detects & parses statements from "NBE Statements" folder.
 * 2. Automated Daily Schedule: Checks for new statements daily at 6:00 AM (or on-demand via menu).
 * 3. 100k Available Balance Tracker: Computes real-time available limit, blocked installments principal, and post-payment projection.
 * 4. Smart Reconciliation: Identifies missing/unrecorded transactions, flags unbilled debits, and syncs installment payment numbers.
 * 5. Dynamic Monthly Debts: Color-coded status (Paid, Due Soon, Overdue), clean numeric formatting (#,##0.00).
 */

const CONFIG = {
  FOLDER_NAME: "NBE Statements",
  PROCESSED_SUBFOLDER: "Processed",
  CREDIT_LIMIT: 100000.00,
  SHEETS: {
    TRANSACTIONS: "Transactions",
    INSTALLMENTS: "Installments",
    DEBT_BREAKDOWN: "Debt Breakdown",
    MONTHLY_OVERVIEW: "Monthly Overview",
    PAYMENT_HISTORY: "Payment History",
    BANK_STATEMENT: "Bank Statement",
    RECONCILIATION: "Reconciliation",
    AUDIT_DIFFERENCES: "Audit & Differences"
  },
  COLORS: {
    PAID: "#d9ead3",       // Soft green
    DUE_SOON: "#fff2cc",   // Soft yellow
    OVERDUE: "#fce8e6",    // Soft red
    UPCOMING: "#f8f9fa",   // Light grey
    HEADER: "#f3f3f3",     // Light neutral grey
    PRIMARY: "#1b5e20",    // NBE dark green
    PRIMARY_LIGHT: "#e8f5e9",
    TEXT_MUTED: "#555555",
    BORDER: "#e0e0e0",
    ACCENT_BLUE: "#e8f0fe", // Soft blue highlight for installments
    ROW_ALT: "#fafafa"
  }
};

// ==========================================
// UTILITIES & SAFE SCRIPT PROPERTIES
// ==========================================

function safeSetScriptProperty(key, val) {
  try {
    let strVal = typeof val === "string" ? val : JSON.stringify(val);
    if (strVal.length > 8500 && key === "ASSIGNED_STATEMENT_CHARGES") {
      try {
        const parsed = JSON.parse(strVal);
        const compacted = {};
        const keys = Object.keys(parsed);
        const recentKeys = keys.slice(-40);
        recentKeys.forEach(k => {
          const item = parsed[k];
          if (item) {
            compacted[k] = {
              payer: item.payer,
              note: (item.note || "").substring(0, 40),
              amount: item.amount,
              desc: (item.desc || "").substring(0, 30)
            };
          }
        });
        strVal = JSON.stringify(compacted);
      } catch (compactErr) { }
    }
    PropertiesService.getScriptProperties().setProperty(key, strVal);
  } catch (err) {
    Logger.log(`[SafeProperties] Could not set '${key}': ` + (err ? (err.message || String(err)) : "Quota limit"));
  }
}


function parseDateValue(rawDate, tz) {
  if (!rawDate) return null;
  const targetTz = tz || "Africa/Cairo";

  // Case 1: Already a Date object (use Utilities.formatDate with target timezone to prevent UTC midnight shift!)
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    const y = parseInt(Utilities.formatDate(rawDate, targetTz, "yyyy"), 10);
    const m = parseInt(Utilities.formatDate(rawDate, targetTz, "M"), 10) - 1;
    const d = parseInt(Utilities.formatDate(rawDate, targetTz, "d"), 10);
    return new Date(y, m, d, 12, 0, 0);
  }

  // Case 2: Numeric Google Sheets serial date (e.g. 46265)
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

  // Case 3: ISO format YYYY-MM-DD
  const iso = s.match(/^(\d{4})[-\/\.](\d{1,2})[-\/\.](\d{1,2})/);
  if (iso) {
    return new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10), 12, 0, 0);
  }

  // Case 4: DD/MM/YYYY or MM/DD/YYYY with smart disambiguation
  const dmy = s.match(/^(\d{1,2})[-\/\.](\d{1,2})[-\/\.](\d{4})/);
  if (dmy) {
    const p1 = parseInt(dmy[1], 10);
    const p2 = parseInt(dmy[2], 10);
    const y = parseInt(dmy[3], 10);
    let day = p1;
    let month = p2 - 1;
    if (p2 > 12 && p1 <= 12) {
      // MM/DD/YYYY (e.g. 8/31/2026)
      month = p1 - 1;
      day = p2;
    } else if (p1 > 12 && p2 <= 12) {
      // DD/MM/YYYY (e.g. 31/8/2026)
      day = p1;
      month = p2 - 1;
    }
    return new Date(y, month, day, 12, 0, 0);
  }

  // Case 5: Text month with optional weekday prefix (e.g. "Monday, August 31, 2026" or "31 August 2026")
  const monthMap = {
    jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2, apr: 3, april: 3,
    may: 4, jun: 5, june: 5, jul: 6, july: 6, aug: 7, august: 7,
    sep: 8, sept: 8, september: 8, oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11
  };
  const named1 = s.match(/(?:[a-zA-Z]+,?\s+)?(\d{1,2})(?:st|nd|rd|th)?[\s\-\/\.]+([a-zA-Z]+)[\s\-\/\.]+(\d{4})/i);
  if (named1 && monthMap[named1[2].toLowerCase()] !== undefined) {
    return new Date(parseInt(named1[3], 10), monthMap[named1[2].toLowerCase()], parseInt(named1[1], 10), 12, 0, 0);
  }
  const named2 = s.match(/(?:[a-zA-Z]+,?\s+)?([a-zA-Z]+)[\s\-\/\.]+(\d{1,2})(?:st|nd|rd|th)?,?[\s\-\/\.]+(\d{4})/i);
  if (named2 && monthMap[named2[1].toLowerCase()] !== undefined) {
    return new Date(parseInt(named2[3], 10), monthMap[named2[1].toLowerCase()], parseInt(named2[2], 10), 12, 0, 0);
  }

  // Case 6: General date string fallback
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const y = parseInt(Utilities.formatDate(d, targetTz, "yyyy"), 10);
    const m = parseInt(Utilities.formatDate(d, targetTz, "M"), 10) - 1;
    const day = parseInt(Utilities.formatDate(d, targetTz, "d"), 10);
    return new Date(y, m, day, 12, 0, 0);
  }
  return null;
}

/**
 * Determines the NBE Credit Card Billing Cycle and Payment Due Date for a purchase.
 * 
 * NBE Credit Card Rules:
 * - Purchases made on days 1 to 30 of month M belong to Month M Billing Cycle (billed month-end M, due on Month M+1, 25th).
 * - Purchases made on day 31 of month M (e.g. August 31st) roll over into Month M+1 Billing Cycle (September),
 *   which is billed at the end of Month M+1 and due on Month M+2, 25th (October 25th).
 */
function getBillingCycleForPurchase(purchaseDate, tz) {
  if (!purchaseDate || isNaN(purchaseDate.getTime())) return null;
  const targetTz = tz || "Africa/Cairo";

  let day = parseInt(Utilities.formatDate(purchaseDate, targetTz, "d"), 10);
  let month = parseInt(Utilities.formatDate(purchaseDate, targetTz, "M"), 10) - 1; // 0-11
  let year = parseInt(Utilities.formatDate(purchaseDate, targetTz, "yyyy"), 10);

  // Timezone safeguard: if date object is on the 31st locally or in UTC, ensure day is 31
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
  let cycleMonth = month; // 0-11

  // Day 31 Rollover Rule:
  // August 31st belongs to September billing cycle (cycleMonth = 8, year = 2026), due October 25th!
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
  // Months are keyed & labeled by PAYMENT DUE month (e.g. Aug 31 purchase -> "October 2026", due Oct 25)
  const cycleDate = new Date(dueYear, dueMonth, 1, 12, 0, 0);

  const cycleSortKey = Utilities.formatDate(cycleDate, targetTz, "yyyy-MM");
  const cycleLabel = Utilities.formatDate(cycleDate, targetTz, "MMMM yyyy");
  const dueDateLabel = Utilities.formatDate(dueDate, targetTz, "MMMM d, yyyy");

  return {
    cycleSortKey: cycleSortKey,
    cycleLabel: cycleLabel,
    cycleDate: cycleDate,
    dueDate: dueDate,
    dueDateLabel: dueDateLabel,
    isDay31Rollover: (day === 31)
  };
}

/**
 * Determines the NBE Credit Card Billing Cycle and Due Date for an Installment EMI.
 * 
 * NBE 55-Day Installment Policy:
 * Purchases made on installment in month M have a full billing cycle grace period (~55 days).
 * The first EMI is billed on month M+1 statement, due on the 25th of month M+2.
 * Example: Purchased in August -> 1st EMI billed on Sep 30 statement (September cycle), due Oct 25.
 */
function getBillingCycleForInstallment(purchaseDate, installmentIndex, tz) {
  if (!purchaseDate || isNaN(purchaseDate.getTime())) return null;
  const targetTz = tz || "Africa/Cairo";

  const month = parseInt(Utilities.formatDate(purchaseDate, targetTz, "M"), 10) - 1; // 0-11
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

  // Keyed & labeled by PAYMENT DUE month
  const cycleDate = new Date(dueYear, dueMonth, 1, 12, 0, 0);
  const dueDate = new Date(dueYear, dueMonth, 25, 12, 0, 0);

  const cycleSortKey = Utilities.formatDate(cycleDate, targetTz, "yyyy-MM");
  const cycleLabel = Utilities.formatDate(cycleDate, targetTz, "MMMM yyyy");
  const dueDateLabel = Utilities.formatDate(dueDate, targetTz, "MMMM d, yyyy");

  return {
    cycleSortKey: cycleSortKey,
    cycleLabel: cycleLabel,
    cycleDate: cycleDate,
    dueDate: dueDate,
    dueDateLabel: dueDateLabel
  };
}

// Backwards compatibility wrappers
function getDueDateForPurchase(purchaseDate, tz) {
  const info = getBillingCycleForPurchase(purchaseDate, tz);
  return info ? info.dueDate : null;
}

function getDueDateForInstallment(purchaseDate, tz) {
  const info = getBillingCycleForInstallment(purchaseDate, 0, tz);
  return info ? info.dueDate : null;
}

// ==========================================
// 1. MENU & TRIGGER SETUP
// ==========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💳 NBE Tracker")
    .addItem("📥 Process Latest Statement (Drive)", "menuProcessStatement")
    .addItem("🔄 Refresh Dashboard & Reconciliation", "updateLiveDashboard")
    .addSeparator()
    .addItem("📊 Go to Debt Breakdown Sheet", "menuGoToDebtBreakdown")
    .addItem("📅 Go to Monthly Overview Sheet", "menuGoToMonthlyOverview")
    .addItem("🔍 Go to Audit & Differences Sheet", "menuGoToAuditDifferences")
    .addSeparator()
    .addItem("⏰ Setup Daily Auto-Check", "setupDailyTrigger")
    .addToUi();
}

function menuGoToDebtBreakdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DEBT_BREAKDOWN);
  if (!sheet) {
    updateLiveDashboard();
    sheet = ss.getSheetByName(CONFIG.SHEETS.DEBT_BREAKDOWN);
  }
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

function menuGoToMonthlyOverview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.MONTHLY_OVERVIEW);
  if (!sheet) {
    updateLiveDashboard();
    sheet = ss.getSheetByName(CONFIG.SHEETS.MONTHLY_OVERVIEW);
  }
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

function menuGoToAuditDifferences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT_DIFFERENCES);
  if (!sheet) {
    updateLiveDashboard();
    sheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT_DIFFERENCES);
  }
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // If user edits Transactions: native equations in Monthly Overview update spontaneously in-browser (sub-50ms)!
  // No need to wipe and rewrite Monthly Overview on every keystroke.
  if (sheetName === CONFIG.SHEETS.TRANSACTIONS) {
    return;
  }

  // If user edits Installments or Payment History: re-sync static installment schedules in the background
  if (sheetName === CONFIG.SHEETS.INSTALLMENTS || sheetName === CONFIG.SHEETS.PAYMENT_HISTORY) {
    if (e.range.getRow() >= 2) {
      try {
        updateLiveDashboard({ skipRecon: true });
      } catch (err) {
        Logger.log("[onEdit] Auto-update error: " + (err ? (err.message || String(err)) : ""));
      }
    }
    return;
  }

  if (sheetName !== CONFIG.SHEETS.RECONCILIATION) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row < 5) return;

  const lastRow = sheet.getLastRow();
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  let table1Start = -1;
  let table2Start = -1;
  let table3Start = -1;

  for (let i = 0; i < colA.length; i++) {
    const text = String(colA[i][0] || "");
    if (text.includes("Suggested Matches")) table1Start = i + 1;
    if (text.includes("MISSING from Sheet")) table2Start = i + 1;
    if (text.includes("NOT Found on Bank Statement") || text.includes("Transactions in Sheet NOT")) table3Start = i + 1;
  }

  // Case 1: Table 2 edit (Col 5 = Assign Payer, Col 6 = Custom Note)
  if ((col === 5 || col === 6) && table2Start > 0 && row > table2Start + 1 && (table3Start === -1 || row < table3Start)) {
    const rowValues = sheet.getRange(row, 1, 1, 6).getValues()[0];
    const itemNum = String(rowValues[0] || (row - table2Start - 1));
    const rawDate = rowValues[1];
    const origDesc = String(rowValues[2] || "").trim();
    const amt = parseFloat(rowValues[3]);
    const payer = String(rowValues[4] || "").trim();
    const note = String(rowValues[5] || "").trim();

    if (isNaN(amt) || amt <= 0) return;

    const dateStr = String(rawDate);
    const chargeKey = `debit_${itemNum}_${amt.toFixed(2)}_${origDesc.substring(0, 30)}`;
    const legacyKey = `${dateStr}_${amt.toFixed(2)}_${origDesc.substring(0, 30)}`;

    const scriptProps = PropertiesService.getScriptProperties();
    let assignedMap = {};
    try {
      assignedMap = JSON.parse(scriptProps.getProperty("ASSIGNED_STATEMENT_CHARGES") || "{}");
    } catch (err) { }

    if (payer) {
      const isNeglected = (payer === "🚫 Neglect / Ignore" || payer.toLowerCase().includes("neglect"));
      const record = {
        itemNum: itemNum,
        dateStr: dateStr,
        desc: origDesc,
        amount: amt,
        payer: isNeglected ? "NEGLECT" : payer,
        note: note
      };
      assignedMap[chargeKey] = record;
      assignedMap[legacyKey] = record;
      safeSetScriptProperty("ASSIGNED_STATEMENT_CHARGES", JSON.stringify(assignedMap));

      if (isNeglected) {
        sheet.getRange(row, 1, 1, 6).setBackground("#eeeeee");
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Marked "${origDesc}" (${amt.toFixed(2)} EGP) as Neglected/Ignored!`,
          "🚫 Charge Neglected",
          4
        );
      } else {
        sheet.getRange(row, 1, 1, 6).setBackground(CONFIG.COLORS.PAID);
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Assigned "${origDesc}" (${amt.toFixed(2)} EGP) to ${payer} directly in Debt Breakdown (⚠️ Missing from Transactions)!`,
          "💳 Assigned to Debt Breakdown",
          4
        );
      }

      updateBankStatementStatusRow(origDesc, amt, isNeglected ? "🚫 Neglected / Ignored" : `⚠️ Unrecorded (Assigned to ${payer})`, isNeglected ? "#eeeeee" : CONFIG.COLORS.DUE_SOON);
      updateLiveDashboard({ skipRecon: true });
    } else {
      delete assignedMap[chargeKey];
      delete assignedMap[legacyKey];
      safeSetScriptProperty("ASSIGNED_STATEMENT_CHARGES", JSON.stringify(assignedMap));
      sheet.getRange(row, 1, 1, 6).setBackground(CONFIG.COLORS.OVERDUE);
      updateBankStatementStatusRow(origDesc, amt, "⚠️ Unassigned (Not in Sheet)", CONFIG.COLORS.OVERDUE);
      updateLiveDashboard({ skipRecon: true });
    }
    return;
  }

  // Case 2: Table 1 confirm match (Col 7 checkbox)
  if (col === 7 && table1Start > 0 && row > table1Start + 1 && (table2Start === -1 || row < table2Start)) {
    if (e.value === "TRUE" || e.range.getValue() === true) {
      const rowValues = sheet.getRange(row, 1, 1, 7).getValues()[0];
      const rawDate = rowValues[1];
      const stmtDesc = String(rowValues[2] || "").trim();
      const amt = parseFloat(rowValues[3]);

      if (!isNaN(amt) && amt > 0) {
        const dateStr = String(rawDate);
        const stKey = `${dateStr}_${amt.toFixed(2)}_${stmtDesc.substring(0, 20)}`;

        const scriptProps = PropertiesService.getScriptProperties();
        const rawJson = scriptProps.getProperty("CONFIRMED_MATCHES") || "[]";
        let setKeys = [];
        try { setKeys = JSON.parse(rawJson); } catch (err) { }
        if (!setKeys.includes(stKey)) {
          setKeys.push(stKey);
          safeSetScriptProperty("CONFIRMED_MATCHES", JSON.stringify(setKeys));
        }

        sheet.getRange(row, 1, 1, 7).setBackground(CONFIG.COLORS.PAID);
        sheet.getRange(row, 7).clearDataValidations().setValue("✅ Confirmed");

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        ss.toast(`Match confirmed for ${stmtDesc}!`, "💳 Match Confirmed", 3);
        updateBankStatementStatusRow(stmtDesc, amt, "✅ Confirmed Match", CONFIG.COLORS.PAID);
        updateLiveDashboard({ skipRecon: true });
      }
    }
    return;
  }

  // Case 3: Table 3 neglect/exclude checkbox (Col 7)
  if (col === 7 && table3Start > 0 && row > table3Start + 1) {
    const rowValues = sheet.getRange(row, 1, 1, 8).getValues()[0];
    const rawDate = rowValues[1];
    const person = String(rowValues[2] || "").trim();
    const origDesc = String(rowValues[3] || "").trim();
    const amt = parseFloat(rowValues[4]);
    const isChecked = (e.value === "TRUE" || e.range.getValue() === true);

    if (isNaN(amt) || amt <= 0) return;

    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || "Africa/Cairo";
    const pDate = parseDateValue(rawDate, tz);
    const dStr = pDate ? Utilities.formatDate(pDate, tz, "yyyy-MM-dd") : String(rawDate);
    const txKeyNorm = `${dStr}_${amt.toFixed(2)}_${origDesc.substring(0, 30)}_${person}`;
    const txKeyLegacy = `${String(rawDate)}_${amt.toFixed(2)}_${origDesc.substring(0, 30)}_${person}`;

    const scriptProps = PropertiesService.getScriptProperties();
    let neglectedKeys = [];
    try {
      neglectedKeys = JSON.parse(scriptProps.getProperty("NEGLECTED_TRANSACTIONS") || "[]");
    } catch (err) { }

    if (isChecked) {
      if (!neglectedKeys.includes(txKeyNorm)) neglectedKeys.push(txKeyNorm);
      if (!neglectedKeys.includes(txKeyLegacy)) neglectedKeys.push(txKeyLegacy);
      safeSetScriptProperty("NEGLECTED_TRANSACTIONS", JSON.stringify(neglectedKeys));
      sheet.getRange(row, 1, 1, 8).setBackground("#eeeeee");
      sheet.getRange(row, 8).setValue("🚫 Neglected (Excluded)");
      updateLiveDashboard({ skipRecon: true });
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Neglected "${origDesc}" (${amt.toFixed(2)} EGP) from active calculations!`,
        "🚫 Transaction Neglected",
        4
      );
    } else {
      neglectedKeys = neglectedKeys.filter(k => k !== txKeyNorm && k !== txKeyLegacy);
      safeSetScriptProperty("NEGLECTED_TRANSACTIONS", JSON.stringify(neglectedKeys));
      sheet.getRange(row, 1, 1, 8).setBackground(null);
      sheet.getRange(row, 8).setValue("Active in Bill");
      updateLiveDashboard({ skipRecon: true });
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Restored "${origDesc}" (${amt.toFixed(2)} EGP) to active calculations!`,
        "✅ Transaction Restored",
        4
      );
    }
  }
}

function updateBankStatementStatusRow(desc, amt, statusText, color) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
    if (!stmtSheet || stmtSheet.getLastRow() < 7) return;

    const lastRow = stmtSheet.getLastRow();
    const rows = stmtSheet.getRange(7, 4, lastRow - 6, 3).getValues(); // Col 4: desc, Col 5: type, Col 6: amt
    for (let i = 0; i < rows.length; i++) {
      const rDesc = String(rows[i][0] || "").trim();
      const rAmt = parseFloat(rows[i][2]);
      if (Math.abs(rAmt - amt) < 0.05 && (rDesc === desc || rDesc.includes(desc.substring(0, 15)) || desc.includes(rDesc.substring(0, 15)))) {
        const targetCell = stmtSheet.getRange(7 + i, 7);
        targetCell.setValue(statusText);
        if (color) targetCell.setBackground(color);
        break;
      }
    }
  } catch (err) { }
}

function setupDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "dailyStatementCheck") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("dailyStatementCheck")
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getUi().alert(
    "Daily Auto-Check Active",
    "The script will now check the '" + CONFIG.FOLDER_NAME + "' Google Drive folder every morning at 6:00 AM.\n\n" +
    "Whenever you upload a statement PDF from your phone, it will be automatically processed!",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function dailyStatementCheck() {
  const result = checkAndProcessNewStatements();
  if (result.processedCount > 0) {
    updateLiveDashboard();
  }
}

function menuProcessStatement() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = checkAndProcessNewStatements();
    updateLiveDashboard();

    if (result && result.processedCount > 0) {
      const meta = result.lastMeta || {};
      const debVal = Number(meta.sumDebits || 0);
      const credVal = Number(meta.sumCredits || 0);
      const stmtDeb = Number(meta.totalDebit || 0);
      const stmtCred = Number(meta.totalCredit || 0);

      const debitSummary = debVal > 0 ? `\n• Total Debits: ${debVal.toFixed(2)} EGP` + (stmtDeb > 0 ? ` (Statement: ${stmtDeb.toFixed(2)} EGP)` : "") : "";
      const creditSummary = credVal > 0 ? `\n• Total Credits: ${credVal.toFixed(2)} EGP` + (stmtCred > 0 ? ` (Statement: ${stmtCred.toFixed(2)} EGP)` : "") : "";

      ui.alert(
        "Statement Processed Successfully",
        `Processed ${result.processedCount} statement file(s):\n` +
        `• File: ${(result.processedFiles || []).join(", ")}\n` +
        `• Statement Date: ${meta.statementDate || "N/A"}\n` +
        `• Due Date: ${meta.dueDate || "N/A"}\n` +
        `• Closing Balance: ${meta.closingBalance ? Number(meta.closingBalance).toFixed(2) + " EGP" : "N/A"}\n` +
        `• Transactions Found: ${result.totalTxCount}${debitSummary}${creditSummary}\n\n` +
        `The Bank Statement, Reconciliation, Debt Breakdown, and Audit & Differences tabs have all been updated!`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        "No New Statements Found",
        `No PDF files were found in '${CONFIG.FOLDER_NAME}' or '${CONFIG.PROCESSED_SUBFOLDER}'.\n\n` +
        `Please upload your NBE Credit Card statement PDF into Google Drive.\n` +
        `The dashboard was refreshed using the current sheet data.`,
        ui.ButtonSet.OK
      );
    }
  } catch (err) {
    const errText = (err && (err.message || err.toString())) ? (err.message || err.toString()) : "An error occurred during statement processing.";
    Logger.log("Error in menuProcessStatement: " + errText + (err && err.stack ? "\n" + err.stack : ""));
    try {
      ui.alert("Error Processing Statement", errText, ui.ButtonSet.OK);
    } catch (uiErr) {
      Logger.log("Failed to show alert dialog: " + uiErr);
    }
  }
}

// ==========================================
// 2. GOOGLE DRIVE PDF DETECTION & OCR
// ==========================================

function getOrCreateFolder(parent, name) {
  const folders = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

function getFullDocumentText(doc) {
  if (!doc) return "";
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
          if (cText) cellTexts.push(cText);
        }
        if (cellTexts.length > 0) {
          // Join cells with tabs so OCR table columns are cleanly separated
          lines.push(cellTexts.join("\t"));
        }
      }
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const liText = child.asListItem().getText().trim();
      if (liText) lines.push(liText);
    }
  }

  // Also include headers and footers if present
  try {
    const header = doc.getHeader();
    if (header) {
      const hText = header.getText().trim();
      if (hText) lines.unshift(hText);
    }
  } catch (e) { }

  try {
    const footer = doc.getFooter();
    if (footer) {
      const fText = footer.getText().trim();
      if (fText) lines.push(fText);
    }
  } catch (e) { }

  return lines.join("\n");
}

function extractTextFromPDF(file) {
  // Method 1: Drive Advanced Service (if enabled)
  try {
    if (typeof Drive !== "undefined" && Drive.Files && Drive.Files.insert) {
      const resource = {
        title: "temp_ocr_" + file.getName(),
        mimeType: "application/pdf"
      };
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

  // Method 2: UrlFetchApp Drive API v2 multipart upload
  const blob = file.getBlob();
  const metadata = {
    title: "temp_ocr_" + file.getName()
  };

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
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200) {
    throw new Error("OCR conversion failed with status " + response.getResponseCode() + ": " + response.getContentText());
  }

  const fileData = JSON.parse(response.getContentText());
  const docId = fileData.id;
  const doc = DocumentApp.openById(docId);
  const text = getFullDocumentText(doc);
  DriveApp.getFileById(docId).setTrashed(true); // Delete temporary Google Doc
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

  // 1. Parse Statement Metadata
  const mCard = text.match(/Card (?:Number|No\.?)[:\s\t]+([0-9\*]+)/i);
  if (mCard) meta.card = mCard[1];

  const mOpenClose = text.match(/Opening Balance[:\s\t]+([\d,]+\.?\d*)[\s\S]*?Closing Balance[:\s\t]+([\d,]+\.?\d*)/i);
  if (mOpenClose) {
    meta.openingBalance = parseFloat(mOpenClose[1].replace(/,/g, ""));
    meta.closingBalance = parseFloat(mOpenClose[2].replace(/,/g, ""));
  }

  const mTotCred = text.match(/Total (?:of )?Credit[:\s\t]+([\d,]+\.?\d*)/i);
  if (mTotCred) {
    meta.totalCredit = parseFloat(mTotCred[1].replace(/,/g, ""));
  }

  const mTotDeb = text.match(/Total (?:of )?Debit[:\s\t]+([\d,]+\.?\d*)/i);
  if (mTotDeb) {
    meta.totalDebit = parseFloat(mTotDeb[1].replace(/,/g, ""));
  }

  const mDueLimit = text.match(/Due Date[:\s\t]+([0-9]{1,2}[\s\-\/\.]+[A-Za-z]+[\s\-\/\.]+[0-9]{2,4}|\d{1,2}\s*[\/\-\.]\s*\d{1,2}\s*[\/\-\.]\s*\d{2,4})[\s\S]*?Credit Limit[:\s\t]+([\d,]+\.?\d*)/i);
  if (mDueLimit) {
    meta.dueDate = mDueLimit[1].trim();
    meta.creditLimit = parseFloat(mDueLimit[2].replace(/,/g, ""));
  }

  const mStmt = text.match(/Statement Date[:\s\t]+([0-9]{1,2}[\s\-\/\.]+[A-Za-z]+[\s\-\/\.]+[0-9]{2,4}|\d{1,2}\s*[\/\-\.]\s*\d{1,2}\s*[\/\-\.]\s*\d{2,4})/i);
  if (mStmt) {
    meta.statementDate = mStmt[1].trim();
  }

  // 2. Enhanced Universal Date Pattern:
  // Supports tabs, multi-spaces, optional year for months, dotted/hyphenated/slashed formats
  const datePattern = /(?:\b\d{1,2}(?:st|nd|rd|th)?[\s\-\/\.]+(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)(?:[\s\-\/\.]+\d{2,4})?\b|\b(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)[\s\-\/\.]+\d{1,2}(?:st|nd|rd|th)?,?(?:[\s\-\/\.]+\d{2,4})?\b|\b\d{1,2}\s*[\/\-\.]\s*\d{1,2}\s*[\/\-\.]\s*\d{2,4}\b)/gi;

  const feePattern = /\b(?:STAMP\s*DUTY|STAMPTAX|FINANCE\s*CHARGE|INTEREST\s*CHARGE|SMS\s*(?:ALERT|SERVICE)|LATE\s*PAYMENT|OVERLIMIT|ANNUAL\s*MEMBERSHIP|ADMINISTRATIVE\s*EXPENSE)\b/i;

  const lines = text.split(/\r?\n/);
  const transactions = [];

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i].trim();
    if (!rawLine) continue;

    // Skip pure page headers, metadata labels, print timestamps, and column header titles
    if (/^Page\s+No\.?\s+\d+/i.test(rawLine)) continue;
    if (/\bas\s*of\b/i.test(rawLine)) continue;
    if (/\b(?:printed|generated)\b/i.test(rawLine)) continue;
    if (/\b(?:Minimum\s*Payment|Min\.?\s*Payment)\b/i.test(rawLine)) continue;
    if (/^(?:Statement\s+Date|Due\s+Date|Payment\s+Due\s+Date|Opening\s+Balance|Closing\s+Balance|Total\s+(?:of\s+)?(?:Credit|Debit)|Credit\s+Limit|Available\s+Credit)[:\s\t]/i.test(rawLine)) continue;
    if (/^Transaction\s+Date\s+Posting\s+Date/i.test(rawLine)) continue;
    if (/^National\s+Bank\s+of\s+Egypt\s*$/i.test(rawLine)) continue;
    if (/^Card\s+Number[:\s\t]+[0-9\*]+\s*$/i.test(rawLine)) continue;

    // Reset datePattern lastIndex before matching
    datePattern.lastIndex = 0;
    let dates = rawLine.match(datePattern);

    // If line has no date but is a known bank fee/charge with a numeric amount, use statementDate
    let isFeeLineWithoutDate = false;
    if ((!dates || dates.length === 0) && feePattern.test(rawLine) && /[\d,]+\.?\d*/.test(rawLine)) {
      isFeeLineWithoutDate = true;
      dates = [meta.statementDate || "Statement Date"];
    }

    if (!dates || dates.length === 0) continue;

    const txDate = dates[0];
    const postDate = dates.length > 1 ? dates[1] : txDate;

    // Remove dates from rawLine to parse transaction details
    let rest = rawLine;
    if (!isFeeLineWithoutDate) {
      dates.slice(0, 2).forEach(d => {
        rest = rest.replace(d, " ");
      });
    }

    let txType = "DEBIT";
    let instCurrent = null;
    let instTotal = null;
    let authCode = "";
    let amount = 0;
    const maxAmountCap = Math.max(meta.creditLimit * 1.5, 200000);

    // Check for installment indicator (e.g. "11 OF 12", "09 OF 12", "02/06")
    const mInst = rest.match(/(\d+)\s+(?:OF|\/)\s+(\d+)/i);
    if (mInst) {
      txType = "INSTALLMENT";
      instCurrent = parseInt(mInst[1], 10);
      instTotal = parseInt(mInst[2], 10);
    } else if (/\b(?:PAYMENT|DIRECT DEBIT|REFUND|CREDIT|DEPOSIT|SETTLEMENT)\b|\bCR\b|[\d,]+\.?\d*\s*CR\b/i.test(rest)) {
      txType = "CREDIT";
    }

    // Amount extraction:
    // Case 1: Foreign currency converted to EGP (e.g. "50.00 USD 2,457.58 EGP")
    const mFx = rest.match(/([\d,]+\.?\d*)\s*(?:USD|EUR|GBP|SAR|AED)\s+([\d,]+\.?\d*)\s*(?:EGP)?/i);
    if (mFx) {
      const parsedFx = parseFloat(mFx[2].replace(/,/g, ""));
      if (!isNaN(parsedFx) && parsedFx > 0 && parsedFx <= maxAmountCap) {
        amount = parsedFx;
      }
    }

    // Case 2: Credit with explicit CR only (e.g. "10000 CR", "16,200.00 CR")
    // Do NOT match hyphens in phone numbers ("650-2530000") or store codes ("-02GEGY")!
    if (amount <= 0) {
      const mCr = rest.match(/([\d,]+(?:\.\d+)?)\s*CR\b/i);
      if (mCr) {
        const parsedCr = parseFloat(mCr[1].replace(/,/g, ""));
        if (!isNaN(parsedCr) && parsedCr > 0 && parsedCr <= maxAmountCap) {
          amount = parsedCr;
        }
      }
    }

    // Case 3: EGP or LE explicitly tagged (e.g. "384.00 EGP", "EGP 384.00", "12.5 EGP", "2.90 EGP")
    if (amount <= 0) {
      const mEgp = rest.match(/([\d,]+(?:\.\d+)?)\s*(?:EGP|LE|L\.E\.)\b/i) || rest.match(/\b(?:EGP|LE|L\.E\.)\s*([\d,]+(?:\.\d+)?)/i);
      if (mEgp) {
        const parsedEgp = parseFloat(mEgp[1].replace(/,/g, ""));
        if (!isNaN(parsedEgp) && parsedEgp > 0 && parsedEgp <= maxAmountCap) {
          amount = parsedEgp;
        }
      }
    }

    // Case 4: Any numbers with decimal point (e.g. "833.25", "12.5", "104.99")
    if (amount <= 0) {
      const allDecimals = (rest.match(/[\d,]+\.\d+/g) || [])
        .map(s => parseFloat(s.replace(/,/g, "")))
        .filter(n => !isNaN(n) && n > 0 && n <= maxAmountCap);
      if (allDecimals.length > 0) {
        amount = allDecimals[allDecimals.length - 1];
      }
    }

    // Case 5: Integer amount (e.g. "10000", "100", "150", "8", "75")
    if (amount <= 0) {
      const cleanRestForInt = rest
        .replace(/\b\d{3}[-\s]*\d{7,8}\b/g, " ")
        .replace(/\b\d{10,20}\b/g, " ");

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

    if (isNaN(amount) || amount <= 0) continue;

    // Clean up description: strip amounts, currencies, delimiters, card tokens, reference numbers
    let desc = rest
      .replace(new RegExp("\\b" + amount.toFixed(2) + "\\b", "g"), " ")
      .replace(new RegExp("\\b" + amount + "\\b", "g"), " ")
      .replace(/[\d,]+\.\d+/g, " ")
      .replace(/\b\d{3}[-\s]*\d{7,8}\b/g, " ") // strip phone numbers like 650-2530000
      .replace(/\b\d{5,8}\b/g, " ") // strip 5-8 digit auth/reference codes
      .replace(/\b\d{10,20}\b/g, " ") // strip long reference numbers
      .replace(/\b(?:EGP|USD|EUR|GBP|SAR|AED|LE|L\.E\.|CR)\b/gi, " ")
      .replace(/\b(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b\s*[,]?/gi, " ")
      .replace(/\b\d{4}\s+[\*xX]{4}\s+[\*xX]{4}\s+\d{4}\b/g, " ") // strip spaced card numbers
      .replace(/\b\d{4}\s*\*{4,}\s*\d{4}\b/g, " ")
      .replace(/[0-9\*]{12,19}/g, " ")
      .replace(/[\t\|\*]+/g, " ");

    desc = desc.replace(/\s+/g, " ")
      .replace(/^[\s,;:\-\|]+/, "")
      .trim();
    if (!desc) desc = "Bank Statement Charge";

    // Double safeguard: Skip if the extracted description is a print timestamp or statement header
    if (/\bas\s*of\b/i.test(desc) || /\b(?:Minimum\s*Payment|Min\.?\s*Payment)\b/i.test(desc) || /\b(?:printed|generated)\b/i.test(desc)) {
      continue;
    }

    // NO DEDUPLICATION OF DISTINCT DOCUMENT ROWS!
    // Every separate row in the statement is an authentic transaction!
    // (e.g. multiple recharges on the same day for the same amount are preserved!)
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

  // 3. Mathematical Reconciliation Check
  const sumDebits = transactions.filter(t => t.type === "DEBIT" || t.type === "INSTALLMENT").reduce((s, t) => s + t.amount, 0);
  const sumCredits = transactions.filter(t => t.type === "CREDIT").reduce((s, t) => s + t.amount, 0);
  meta.sumDebits = sumDebits;
  meta.sumCredits = sumCredits;

  Logger.log(`[Statement Parser] Extracted ${transactions.length} items. Total Debits: ${sumDebits.toFixed(2)} (Stmt Meta: ${Number(meta.totalDebit || 0).toFixed(2)}), Total Credits: ${sumCredits.toFixed(2)} (Stmt Meta: ${Number(meta.totalCredit || 0).toFixed(2)})`);

  return { meta: meta, transactions: transactions };
}

function checkAndProcessNewStatements() {
  const rootFolder = getOrCreateFolder(null, CONFIG.FOLDER_NAME);
  const processedFolder = getOrCreateFolder(rootFolder, CONFIG.PROCESSED_SUBFOLDER);

  const files = rootFolder.getFiles();
  const pdfFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.PDF) {
      pdfFiles.push(file);
    }
  }

  // If no PDF found in rootFolder, automatically check processedFolder for the most recent statement
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

  const result = {
    processedCount: 0,
    processedFiles: [],
    lastMeta: {},
    totalTxCount: 0
  };

  if (pdfFiles.length === 0) {
    return result;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  pdfFiles.forEach(file => {
    const text = extractTextFromPDF(file);
    const parsed = parseNBEStatementText(text);

    if (parsed.transactions.length > 0 || parsed.meta.statementDate) {
      writeStatementToSheet(ss, parsed.meta, parsed.transactions);
      syncInstallmentsFromStatement(ss, parsed.transactions);

      result.processedCount++;
      result.processedFiles.push(file.getName());
      result.lastMeta = parsed.meta;
      result.totalTxCount += parsed.transactions.length;

      // Move file to Processed folder if not already there
      if (!isFromProcessed) {
        file.moveTo(processedFolder);
      }
    }
  });

  return result;
}

// ==========================================
// 3. BANK STATEMENT SHEET MANAGEMENT
// ==========================================

function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function writeStatementToSheet(ss, meta, transactions) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.BANK_STATEMENT);
  sheet.clear();

  // Summary Banner
  sheet.getRange("A1:G1").merge()
    .setValue("NBE Credit Card Statement — " + (meta.statementDate || "Latest"))
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  const metaRows = [
    ["Statement Date:", meta.statementDate, "Due Date:", meta.dueDate, "Credit Limit:", meta.creditLimit],
    ["Opening Balance:", meta.openingBalance, "Closing Balance:", meta.closingBalance, "Total Debit:", meta.totalDebit],
    ["Total Credit:", meta.totalCredit, "Card Number:", meta.card, "", ""]
  ];

  const metaRange = sheet.getRange(2, 1, metaRows.length, 6);
  metaRange.setValues(metaRows);
  sheet.getRange("A2:F4").setFontSize(10);
  sheet.getRange("A2:A4").setFontWeight("bold");
  sheet.getRange("C2:C4").setFontWeight("bold");
  sheet.getRange("E2:E4").setFontWeight("bold");
  sheet.getRange("B2").setNumberFormat("@");
  sheet.getRange("D2").setNumberFormat("@");
  sheet.getRange("B3:B4").setNumberFormat("#,##0.00");
  sheet.getRange("D3:D4").setNumberFormat("#,##0.00");
  sheet.getRange("F2:F4").setNumberFormat("#,##0.00");
  sheet.getRange("2:4").setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

  // Table Headers
  const startRow = 6;
  const headers = ["#", "Tx Date", "Posting Date", "Description", "Type", "Amount (EGP)", "Assigned Payer / Status"];
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, false, false);

  if (transactions.length > 0) {
    const rows = transactions.map((t, idx) => {
      const txD = new Date(t.txDate);
      const postD = new Date(t.postDate);
      let initStatus = "⏳ Checking...";
      if (t.type === "CREDIT") initStatus = "💳 Bank Payment Settled";
      return [
        idx + 1,
        isNaN(txD.getTime()) ? t.txDate : txD,
        isNaN(postD.getTime()) ? t.postDate : postD,
        t.desc,
        t.type,
        t.amount,
        initStatus
      ];
    });

    const dataRange = sheet.getRange(startRow + 1, 1, rows.length, headers.length);
    dataRange.setValues(rows);
    sheet.getRange(startRow + 1, 2, rows.length, 2).setNumberFormat("dddd, MMMM d, yyyy");
    sheet.getRange(startRow + 1, 6, rows.length, 1).setNumberFormat("#,##0.00");

    // Color rows by type
    for (let i = 0; i < transactions.length; i++) {
      const type = transactions[i].type;
      const rowNum = startRow + 1 + i;
      if (type === "CREDIT") {
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground(CONFIG.COLORS.PAID);
      } else if (type === "INSTALLMENT") {
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground(CONFIG.COLORS.PRIMARY_LIGHT);
      }
    }
  }

  sheet.autoResizeColumns(1, headers.length);
}

// ==========================================
// 4. INSTALLMENTS SPLIT MATCHING & AUTO-SYNC
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
    if (currentCombo.length >= k || startIdx >= n) {
      return null;
    }

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

  // Pass 1: 1-to-1 exact match by EMI
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

  // Pass 2: Split match on same description & duration (e.g. Deepfreezer, Kettle, Earphones, Hand blender)
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
          payerLabel: uniquePayers.join(" + ") + " (Split)",
          type: "SPLIT_SAME_DESC"
        });
        delete descGroups[key];
        break;
      }
    }
  });

  // Pass 3: General subset matching for remaining items (e.g. multi-item cart)
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
        payerLabel: uniquePayers.join(" + ") + " (Split)",
        type: "SPLIT_CART"
      });
    }
  });

  return results;
}

function syncInstallmentsFromStatement(ss, transactions) {
  const installmentsSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  if (!installmentsSheet || installmentsSheet.getLastRow() < 2) return [];

  const statementInstallments = transactions.filter(t => t.type === "INSTALLMENT");
  if (statementInstallments.length === 0) return [];

  const numRows = installmentsSheet.getLastRow() - 1;
  const range = installmentsSheet.getRange(2, 1, numRows, 11);
  const values = range.getValues();

  const sheetInstallmentRows = [];
  values.forEach((row, idx) => {
    const desc = String(row[2] || "").trim();
    const duration = parseInt(row[4], 10);
    const emi = parseFloat(row[7]);
    const payer = String(row[8] || "").trim();
    if (!isNaN(emi) && emi > 0) {
      sheetInstallmentRows.push({
        index: idx,
        desc: desc,
        duration: duration,
        emi: emi,
        payer: payer
      });
    }
  });

  const matches = matchStatementInstallments(statementInstallments, sheetInstallmentRows);
  // NOTE: Preserving Installments sheet untouched as user-managed input
  return matches;
}

// ==========================================
// ==========================================
// 5. SMART RECONCILIATION ENGINE
// ==========================================

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
  let cleaned = String(name).trim().replace(/\s+/g, " ");
  if (!cleaned) return "";

  if (isNeglectedPayer(cleaned)) return "";
  const lower = cleaned.toLowerCase();
  if (lower === "shared") return "";

  // Common aliases & typo unification
  if (lower === "me" || lower === "myself") return "Mido";
  if (lower === "abd") return "Abdo";
  if (lower === "zozza") return "Zoza";
  if (lower === "mohanad") return "Muhanad";

  return cleaned
    .toLowerCase()
    .split(" ")
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ");
}

function splitPayerNames(payerStr) {
  if (!payerStr) return [];
  const s = String(payerStr).trim();
  if (isNeglectedPayer(s)) return [];

  // Support splitting by +, &, /, and " and " (e.g. "Me and Abdo")
  return s
    .replace(/\s+and\s+/gi, " + ")
    .split(/[\+\,\/\&]/)
    .map(p => normalizePersonName(p))
    .filter(p => p && p.toLowerCase() !== "shared" && !isNeglectedPayer(p));
}

function getAllUniquePayers(ss) {
  const peopleSet = new Set();

  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (txSheet && txSheet.getLastRow() >= 2) {
    const numRows = txSheet.getLastRow() - 1;
    const vals = txSheet.getRange(2, 4, numRows, 1).getValues();
    vals.forEach(r => {
      splitPayerNames(r[0]).forEach(name => {
        if (name && !isNeglectedPayer(name) && name.toLowerCase() !== "shared") peopleSet.add(name);
      });
    });
  }

  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  if (instSheet && instSheet.getLastRow() >= 2) {
    const numRows = instSheet.getLastRow() - 1;
    const vals = instSheet.getRange(2, 9, numRows, 1).getValues();
    vals.forEach(r => {
      splitPayerNames(r[0]).forEach(name => {
        if (name && !isNeglectedPayer(name) && name.toLowerCase() !== "shared") peopleSet.add(name);
      });
    });
  }

  // Also include payers from stored statement assignments
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const assignedMap = JSON.parse(scriptProps.getProperty("ASSIGNED_STATEMENT_CHARGES") || "{}");
    Object.keys(assignedMap).forEach(k => {
      const item = assignedMap[k];
      if (item && item.payer) {
        splitPayerNames(item.payer).forEach(name => {
          if (name && !isNeglectedPayer(name) && name.toLowerCase() !== "shared") peopleSet.add(name);
        });
      }
    });
  } catch (e) { }

  // Also include payers from visible Reconciliation sheet Table 2
  try {
    const reconSheet = ss.getSheetByName(CONFIG.SHEETS.RECONCILIATION);
    if (reconSheet && reconSheet.getLastRow() >= 5) {
      const lastR = reconSheet.getLastRow();
      const colA = reconSheet.getRange(1, 1, lastR, 1).getValues();
      let t2 = -1;
      let t3 = -1;
      for (let i = 0; i < colA.length; i++) {
        const txt = String(colA[i][0] || "");
        if (txt.includes("MISSING from Sheet")) t2 = i + 1;
        if (txt.includes("NOT Found on Bank Statement") || txt.includes("Transactions in Sheet NOT")) t3 = i + 1;
      }
      if (t2 > 0) {
        const endRow = (t3 > t2) ? t3 - 1 : lastR;
        const numRows = endRow - (t2 + 1);
        if (numRows > 0) {
          const pVals = reconSheet.getRange(t2 + 2, 5, numRows, 1).getValues();
          pVals.forEach(r => {
            splitPayerNames(r[0]).forEach(name => {
              if (name && !isNeglectedPayer(name) && name.toLowerCase() !== "shared") peopleSet.add(name);
            });
          });
        }
      }
    }
  } catch (e) { }

  const result = Array.from(peopleSet)
    .map(p => normalizePersonName(p))
    .filter(p => p && !isNeglectedPayer(p) && p.toLowerCase() !== "shared");
  const uniqueResult = Array.from(new Set(result)).sort((a, b) => a.localeCompare(b));
  return uniqueResult.length > 0 ? uniqueResult : ["Abdo", "Dad", "Hager", "Mai", "Mido", "Muhanad", "Mum", "Nourween", "Zoza"];
}

function getStatementPeriod(stmtDate, minStmtTxDate, maxStmtTxDate) {
  if (!stmtDate && !minStmtTxDate) return { start: null, end: null };

  let pStart = null;
  let pEnd = null;

  // STRICT RULE: If the bank statement has transactions, the cycle start date IS
  // the earliest transaction date found in the PDF (e.g. July 30).
  // "YOU FOLLOW THE DATES IN THE PDF!!"
  if (minStmtTxDate) {
    pStart = new Date(minStmtTxDate.getFullYear(), minStmtTxDate.getMonth(), minStmtTxDate.getDate(), 0, 0, 0, 0);
  } else if (stmtDate) {
    // Fallback only if statement has no transactions: 1st of statement month
    pStart = new Date(stmtDate.getFullYear(), stmtDate.getMonth(), 1, 0, 0, 0, 0);
  }

  // Cycle end date is the statement cutoff date (end of day)
  if (stmtDate) {
    pEnd = new Date(stmtDate.getFullYear(), stmtDate.getMonth(), stmtDate.getDate(), 23, 59, 59, 999);
  }
  if (maxStmtTxDate) {
    const maxEnd = new Date(maxStmtTxDate.getFullYear(), maxStmtTxDate.getMonth(), maxStmtTxDate.getDate(), 23, 59, 59, 999);
    if (!pEnd || maxEnd > pEnd) {
      pEnd = maxEnd;
    }
  }

  return { start: pStart, end: pEnd };
}

function runReconciliation(ss, tz) {
  const reconSheet = getOrCreateSheet(ss, CONFIG.SHEETS.RECONCILIATION);

  // CRITICAL: Preserve any existing user assignments currently on the Reconciliation sheet BEFORE clearing!
  const scriptProps = PropertiesService.getScriptProperties();
  let assignedMap = {};
  try {
    assignedMap = JSON.parse(scriptProps.getProperty("ASSIGNED_STATEMENT_CHARGES") || "{}");
  } catch (err) { }

  if (reconSheet.getLastRow() >= 5) {
    const lastR = reconSheet.getLastRow();
    const colA = reconSheet.getRange(1, 1, lastR, 1).getValues();
    let t2Start = -1;
    let t3Start = -1;
    for (let i = 0; i < colA.length; i++) {
      const txt = String(colA[i][0] || "");
      if (txt.includes("MISSING from Sheet")) t2Start = i + 1;
      if (txt.includes("NOT Found on Bank Statement") || txt.includes("Transactions in Sheet NOT")) t3Start = i + 1;
    }
    if (t2Start > 0) {
      const endR = (t3Start > t2Start) ? t3Start - 1 : lastR;
      const numR = endR - (t2Start + 1);
      if (numR > 0) {
        const existingRows = reconSheet.getRange(t2Start + 2, 1, numR, 6).getValues();
        existingRows.forEach(r => {
          const itemNum = String(r[0] || "");
          const rawD = r[1];
          const dText = String(r[2] || "").trim();
          const aNum = parseFloat(r[3]);
          const pText = String(r[4] || "").trim();
          const nText = String(r[5] || "").trim();
          if (!isNaN(aNum) && aNum > 0 && dText && pText) {
            const k1 = `debit_${itemNum}_${aNum.toFixed(2)}_${dText.substring(0, 30)}`;
            const k2 = `${String(rawD)}_${aNum.toFixed(2)}_${dText.substring(0, 30)}`;
            const isNeg = (pText === "🚫 Neglect / Ignore" || pText.toLowerCase().includes("neglect"));
            const rec = {
              itemNum: itemNum,
              dateStr: String(rawD),
              desc: dText,
              amount: aNum,
              payer: isNeg ? "NEGLECT" : pText,
              note: nText
            };
            assignedMap[k1] = rec;
            assignedMap[k2] = rec;
          }
        });
        safeSetScriptProperty("ASSIGNED_STATEMENT_CHARGES", JSON.stringify(assignedMap));
      }
    }
  }

  reconSheet.clear();

  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);

  if (!txSheet || !stmtSheet || stmtSheet.getLastRow() < 7) {
    reconSheet.getRange("A1").setValue("Please upload and process a bank statement to view reconciliation.");
    return;
  }

  // 1. Read Statement Debits & determine transaction date boundaries
  const stmtLastRow = stmtSheet.getLastRow();
  const stmtData = stmtSheet.getRange(7, 1, stmtLastRow - 6, 7).getValues();
  const stmtDebits = [];
  let minStmtTxDate = null;
  let maxStmtTxDate = null;

  stmtData.forEach((row, idx) => {
    const pDate = parseDateValue(row[1], tz);
    if (pDate) {
      if (!minStmtTxDate || pDate < minStmtTxDate) minStmtTxDate = pDate;
      if (!maxStmtTxDate || pDate > maxStmtTxDate) maxStmtTxDate = pDate;
    }

    const type = String(row[4]).trim();
    if (type === "DEBIT") {
      stmtDebits.push({
        id: idx + 1,
        stmtRowIndex: idx,
        dateStr: String(row[1]),
        date: pDate,
        desc: String(row[3]),
        amount: parseFloat(row[5]),
        matched: false,
        matchType: null,
        matchedItems: []
      });
    }
  });

  // Calculate Statement Period (Billing Cycle)
  const stmtDateVal = stmtSheet.getRange("B2").getValue();
  const stmtDate = parseDateValue(stmtDateVal, tz);
  const stmtPeriod = getStatementPeriod(stmtDate, minStmtTxDate, maxStmtTxDate);

  // Read Installments for payer lookup on Bank Statement
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  const sheetInstallmentRows = [];
  if (instSheet && instSheet.getLastRow() >= 2) {
    const instData = instSheet.getRange(2, 1, instSheet.getLastRow() - 1, 11).getValues();
    instData.forEach((iRow, idx) => {
      const emi = parseFloat(iRow[7]);
      const payer = String(iRow[8] || "").trim();
      const desc = String(iRow[2] || "").trim();
      const duration = parseInt(iRow[4], 10);
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

  // Extract statement installment rows for split matching
  const stmtInstallments = [];
  stmtData.forEach((row, idx) => {
    const type = String(row[4]).trim().toUpperCase();
    if (type === "INSTALLMENT") {
      const desc = String(row[3] || "");
      let instCurrent = null;
      let instTotal = null;
      const mInst = desc.match(/(\d+)\s+OF\s+(\d+)/i);
      if (mInst) {
        instCurrent = parseInt(mInst[1], 10);
        instTotal = parseInt(mInst[2], 10);
      }
      stmtInstallments.push({
        stmtRowIdx: idx,
        desc: desc,
        amount: parseFloat(row[5]),
        instCurrent: instCurrent,
        instTotal: instTotal
      });
    }
  });

  const instMatches = matchStatementInstallments(stmtInstallments, sheetInstallmentRows);
  const instMatchMap = {};
  instMatches.forEach(m => {
    instMatchMap[m.stmt.stmtRowIdx] = m.payerLabel;
  });

  // 2. Read Sheet Transactions (Strictly filtered to Statement Period)
  const txDebits = [];
  if (txSheet.getLastRow() >= 2) {
    const txData = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 5).getValues();
    txData.forEach((row, idx) => {
      const rawDate = row[1];
      const rawPerson = row[3];
      const rawAmount = row[4];
      if (!rawDate || !rawAmount) return;

      const pDate = parseDateValue(rawDate, tz);
      const amt = parseFloat(rawAmount);
      if (pDate && !isNaN(amt) && amt > 0) {
        // STRICT STATEMENT PERIOD FILTER:
        // Only include sheet transactions that fall within the statement's billing period!
        if (stmtPeriod.start && pDate < stmtPeriod.start) return;
        if (stmtPeriod.end && pDate > stmtPeriod.end) return;

        txDebits.push({
          rowNum: idx + 2,
          date: pDate,
          desc: String(row[2] || ""),
          person: String(rawPerson || ""),
          amount: amt,
          matched: false
        });
      }
    });
  }

  // 3. Multi-Pass Matching
  // Pass 0: Pre-confirmed Matches from ScriptProperties
  const confirmedMatchesJson = scriptProps.getProperty("CONFIRMED_MATCHES") || "[]";
  let confirmedMatchKeys = new Set();
  try {
    confirmedMatchKeys = new Set(JSON.parse(confirmedMatchesJson));
  } catch (e) { }

  stmtDebits.forEach(st => {
    const stKey = `${st.dateStr}_${st.amount.toFixed(2)}_${st.desc.substring(0, 20)}`;
    if (confirmedMatchKeys.has(stKey)) {
      st.matched = true;
      st.matchType = "CONFIRMED";
      // Find matching items in txDebits and mark them matched so they don't appear in Table 3!
      const singleMatch = txDebits.find(tx => !tx.matched && Math.abs(tx.amount - st.amount) <= 0.05);
      if (singleMatch) {
        singleMatch.matched = true;
        st.matchedItems.push(singleMatch);
      } else {
        const combo = findSubsetCombination(txDebits.filter(tx => !tx.matched), st.amount, 6, 0.10);
        if (combo) {
          combo.forEach(c => {
            c.matched = true;
            st.matchedItems.push(c);
          });
        }
      }
    }
  });

  // Pass 1: 1-to-1 Exact Match ONLY (Amount exact within 0.05, Date within 3 days, Keyword Overlap)
  stmtDebits.forEach(st => {
    if (st.matched) return;
    const match = txDebits.find(tx => {
      if (tx.matched) return false;
      const amtDiff = Math.abs(tx.amount - st.amount);
      if (amtDiff > 0.05) return false;

      if (st.date) {
        const daysDiff = Math.abs((tx.date.getTime() - st.date.getTime()) / (1000 * 60 * 60 * 24));
        if (daysDiff > 3) return false;
      }

      return hasMerchantKeywordOverlap(st.desc, tx.desc);
    });

    if (match) {
      st.matched = true;
      st.matchType = "EXACT";
      match.matched = true;
      st.matchedItems.push(match);
    }
  });

  // Pass 2: SUGGESTED MATCHES (ALL Splits & Approximate Matches MUST BE CONFIRMED BY USER!)
  // NO AUTO-MATCHING OR AUTO-ASSIGNING ALLOWED!
  const suggestedMatches = [];

  stmtDebits.forEach(st => {
    if (st.matched) return;

    const nearby = txDebits.filter(tx => {
      if (tx.matched) return false;
      if (st.date) {
        const daysDiff = Math.abs((tx.date.getTime() - st.date.getTime()) / (1000 * 60 * 60 * 24));
        return daysDiff <= 5;
      }
      return true;
    });

    if (nearby.length === 0) return;

    // Case A: 1-to-1 exact amount match, but different description / aggregator / typo
    const singleAmtMatch = nearby.find(tx => Math.abs(tx.amount - st.amount) < 0.05);
    if (singleAmtMatch) {
      suggestedMatches.push({
        stmt: st,
        sheetItems: [singleAmtMatch],
        reason: `Exact amount (${st.amount.toFixed(2)} EGP), dates within 5 days — check shop name / aggregator`
      });
      return;
    }

    // Case B: N-way split sum matches (2 to 6 items sum to 1 statement charge)
    const combo = findSubsetCombination(nearby, st.amount, 6, 0.10);
    if (combo) {
      const peopleList = combo.map(it => it.person).join(" + ");
      const keywordOverlap = combo.some(item => hasMerchantKeywordOverlap(st.desc, item.desc));
      const reasonSuffix = keywordOverlap ? ` (Merchant match: ${st.desc})` : "";
      suggestedMatches.push({
        stmt: st,
        sheetItems: combo,
        reason: `${combo.length}-way split sum matches ${st.amount.toFixed(2)} EGP (${peopleList})${reasonSuffix}`
      });
      return;
    }
  });

  const suggestedStmtIds = new Set(suggestedMatches.map(s => s.stmt.id));
  const missingInSheet = stmtDebits.filter(st => !st.matched && !suggestedStmtIds.has(st.id));

  // Identify unmatched sheet transactions
  const suggestedTxRowNums = new Set();
  suggestedMatches.forEach(sm => {
    sm.sheetItems.forEach(it => {
      if (it.rowNum) suggestedTxRowNums.add(it.rowNum);
    });
  });
  const unmatchedTxDebits = txDebits.filter(tx => !tx.matched && !suggestedTxRowNums.has(tx.rowNum));

  // 4. Render Reconciliation Sheet
  const periodStr = (stmtPeriod.start && stmtPeriod.end)
    ? ` (${Utilities.formatDate(stmtPeriod.start, tz, "MMM d, yyyy")} – ${Utilities.formatDate(stmtPeriod.end, tz, "MMM d, yyyy")})`
    : "";

  reconSheet.getRange("A1:H1").merge()
    .setValue("Statement Reconciliation — Smart Audit" + periodStr)
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  const confirmedMatchesCount = stmtDebits.filter(st => st.matched).length;
  const kpiRow = [
    "Confirmed Matched:", confirmedMatchesCount,
    "Needs Confirmation:", suggestedMatches.length,
    "On Statement MISSING from Sheet:", missingInSheet.length,
    "In Sheet NOT on Statement:", unmatchedTxDebits.length
  ];
  reconSheet.getRange("A2:H2").setValues([kpiRow])
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);
  reconSheet.getRange("B2").setFontColor(CONFIG.COLORS.PRIMARY);
  reconSheet.getRange("D2").setFontColor("#b06000");
  reconSheet.getRange("F2").setFontColor("#b71c1c");
  reconSheet.getRange("H2").setFontColor("#1565c0");

  let curRow = 4;

  // TABLE 1: ❓ Suggested Matches (Needs Confirmation)
  if (suggestedMatches.length > 0) {
    reconSheet.getRange(curRow, 1, 1, 7).merge()
      .setValue("❓ Suggested Matches (Check box to confirm match) — " + suggestedMatches.length + " items")
      .setFontWeight("bold")
      .setBackground(CONFIG.COLORS.DUE_SOON);
    curRow++;

    const sugHeaders = ["#", "Statement Date", "Statement Charge", "Amount (EGP)", "Suggested Sheet Items", "Match Reason", "Confirm Match?"];
    reconSheet.getRange(curRow, 1, 1, 7).setValues([sugHeaders])
      .setFontWeight("bold")
      .setBackground(CONFIG.COLORS.HEADER);
    curRow++;

    const sugRows = suggestedMatches.map((s, idx) => {
      const itemsDesc = s.sheetItems.map(it => `${it.desc} (${it.person}: ${it.amount.toFixed(2)})`).join(" + ");
      return [
        idx + 1,
        s.stmt.date || s.stmt.dateStr,
        s.stmt.desc,
        s.stmt.amount,
        itemsDesc,
        s.reason,
        false
      ];
    });

    const sugRange = reconSheet.getRange(curRow, 1, sugRows.length, 7);
    sugRange.setValues(sugRows);
    reconSheet.getRange(curRow, 2, sugRows.length, 1).setNumberFormat("dddd, MMMM d, yyyy");
    reconSheet.getRange(curRow, 4, sugRows.length, 1).setNumberFormat("#,##0.00");
    reconSheet.getRange(curRow, 7, sugRows.length, 1).insertCheckboxes();
    curRow += sugRows.length + 1;
  }

  // TABLE 2: ⚠️ Charges on Statement MISSING from Sheet
  reconSheet.getRange(curRow, 1, 1, 6).merge()
    .setValue("⚠️ Charges on Statement MISSING from Sheet (" + missingInSheet.length + " items)")
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.OVERDUE);
  curRow++;

  const missHeaders = ["#", "Date", "Description / Merchant", "Amount (EGP)", "Assign Payer", "Custom Note"];
  reconSheet.getRange(curRow, 1, 1, 6).setValues([missHeaders])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  if (missingInSheet.length === 0) {
    reconSheet.getRange(curRow, 1, 1, 6).merge()
      .setValue("✅ All statement debits are successfully recorded in Transactions!")
      .setFontColor(CONFIG.COLORS.PRIMARY);
    curRow += 2;
  } else {
    // Dynamically fetch all payers from Transactions and Installments (no "Shared")
    const dynamicPayers = getAllUniquePayers(ss);
    const payerOptions = [...dynamicPayers, "🚫 Neglect / Ignore"];
    const payerValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(payerOptions, true)
      .setAllowInvalid(true)
      .build();

    const scriptProps = PropertiesService.getScriptProperties();
    let assignedMap = {};
    try {
      assignedMap = JSON.parse(scriptProps.getProperty("ASSIGNED_STATEMENT_CHARGES") || "{}");
    } catch (err) { }

    const missRows = missingInSheet.map((m, idx) => {
      const itemNum = String(idx + 1);
      const chargeKey = `debit_${itemNum}_${m.amount.toFixed(2)}_${m.desc.substring(0, 30)}`;
      const legacyKey = `${String(m.date || m.dateStr)}_${m.amount.toFixed(2)}_${m.desc.substring(0, 30)}`;
      const saved = assignedMap[chargeKey] || assignedMap[legacyKey];
      let assignedPayer = (saved && saved.payer) ? saved.payer : "";
      if (assignedPayer === "NEGLECT") assignedPayer = "🚫 Neglect / Ignore";
      const note = (saved && saved.note) ? saved.note : cleanMerchantName(m.desc);
      return [
        idx + 1,
        m.date || m.dateStr,
        m.desc,
        m.amount,
        assignedPayer,
        note
      ];
    });

    const missRange = reconSheet.getRange(curRow, 1, missRows.length, 6);
    missRange.setValues(missRows);
    reconSheet.getRange(curRow, 2, missRows.length, 1).setNumberFormat("dddd, MMMM d, yyyy");
    reconSheet.getRange(curRow, 4, missRows.length, 1).setNumberFormat("#,##0.00");
    reconSheet.getRange(curRow, 5, missRows.length, 1).setDataValidation(payerValidation);

    // Color assigned rows
    for (let i = 0; i < missRows.length; i++) {
      if (missRows[i][4] === "🚫 Neglect / Ignore") {
        reconSheet.getRange(curRow + i, 1, 1, 6).setBackground("#eeeeee");
      } else if (missRows[i][4]) {
        reconSheet.getRange(curRow + i, 1, 1, 6).setBackground(CONFIG.COLORS.PAID);
      }
    }
    curRow += missRows.length + 1;
  }

  // TABLE 3: 📋 Transactions in Sheet NOT Found on Bank Statement
  reconSheet.getRange(curRow, 1, 1, 8).merge()
    .setValue("📋 Transactions in Sheet NOT Found on Bank Statement (" + unmatchedTxDebits.length + " items)")
    .setFontWeight("bold")
    .setBackground("#e3f2fd");
  curRow++;

  reconSheet.getRange(curRow, 1, 1, 8).merge()
    .setValue("Charges recorded in Transactions sheet within this statement period that were not billed on this statement. Check the box to Neglect/Exclude any item from active bill calculations.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);
  curRow++;

  const unmatHeaders = [
    "#",
    "Purchase Date",
    "Person",
    "Description / Merchant",
    "Amount (EGP)",
    "Period Status",
    "Neglect / Exclude from Active Bill?",
    "Status / Note"
  ];

  reconSheet.getRange(curRow, 1, 1, unmatHeaders.length).setValues([unmatHeaders])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  if (unmatchedTxDebits.length === 0) {
    reconSheet.getRange(curRow, 1, 1, 8).merge()
      .setValue("✅ All sheet transactions in this statement period are accounted for or matched with statement items!")
      .setFontColor(CONFIG.COLORS.PRIMARY);
    curRow += 2;
  } else {
    const scriptProps = PropertiesService.getScriptProperties();
    let neglectedKeys = [];
    try {
      neglectedKeys = JSON.parse(scriptProps.getProperty("NEGLECTED_TRANSACTIONS") || "[]");
    } catch (e) { }

    const unmatRows = unmatchedTxDebits.map((tx, idx) => {
      const isNearCutoff = (stmtPeriod.end && Math.abs((stmtPeriod.end.getTime() - tx.date.getTime()) / (1000 * 60 * 60 * 24)) <= 2);
      const cycleClass = isNearCutoff ? "⚠️ Near Cutoff (May roll to next statement)" : "⚠️ In-Period (Not on Bank Statement)";

      const dStr = Utilities.formatDate(tx.date, tz, "yyyy-MM-dd");
      const txKeyNorm = `${dStr}_${tx.amount.toFixed(2)}_${tx.desc.substring(0, 30)}_${tx.person}`;
      const txKeyLegacy = `${String(tx.date)}_${tx.amount.toFixed(2)}_${tx.desc.substring(0, 30)}_${tx.person}`;
      const isNeglected = neglectedKeys.includes(txKeyNorm) || neglectedKeys.includes(txKeyLegacy);

      let statusNote = "Active in Bill";
      if (isNeglected) {
        statusNote = "🚫 Neglected (Excluded)";
      } else if (isNearCutoff) {
        statusNote = "⚠️ Check Next Stmt";
      }

      return [
        idx + 1,
        tx.date,
        tx.person,
        tx.desc,
        tx.amount,
        cycleClass,
        isNeglected,
        statusNote
      ];
    });

    const unmatRange = reconSheet.getRange(curRow, 1, unmatRows.length, unmatHeaders.length);
    unmatRange.setValues(unmatRows);
    reconSheet.getRange(curRow, 2, unmatRows.length, 1).setNumberFormat("dddd, MMMM d, yyyy");
    reconSheet.getRange(curRow, 5, unmatRows.length, 1).setNumberFormat("#,##0.00");
    reconSheet.getRange(curRow, 7, unmatRows.length, 1).insertCheckboxes();

    for (let i = 0; i < unmatRows.length; i++) {
      const isNeg = unmatRows[i][6];
      const isNear = String(unmatRows[i][5]).includes("Near Cutoff");
      if (isNeg) {
        reconSheet.getRange(curRow + i, 1, 1, unmatHeaders.length).setBackground("#eeeeee");
        reconSheet.getRange(curRow + i, 4).setFontLine("line-through");
      } else if (isNear) {
        reconSheet.getRange(curRow + i, 1, 1, unmatHeaders.length).setBackground("#fffde7");
      } else {
        reconSheet.getRange(curRow + i, 1, 1, unmatHeaders.length).setBackground("#fff9c4"); // Soft warning yellow
      }
    }
    curRow += unmatRows.length + 1;
  }

  // 5. Update Column H ('Assigned Payer / Status') in Bank Statement Sheet
  const stmtRowStatus = new Array(stmtData.length);
  for (let i = 0; i < stmtData.length; i++) {
    const rType = String(stmtData[i][4]).trim().toUpperCase();
    const rAmt = parseFloat(stmtData[i][5]);

    if (rType === "CREDIT") {
      stmtRowStatus[i] = { text: "💳 Bank Payment Settled", color: "#e1f5fe" };
    } else if (rType === "INSTALLMENT") {
      const payerLabel = instMatchMap[i];
      if (payerLabel) {
        stmtRowStatus[i] = { text: "✅ " + payerLabel + " (Installment)", color: CONFIG.COLORS.PAID };
      } else {
        stmtRowStatus[i] = { text: "⚠️ Unassigned Installment", color: CONFIG.COLORS.OVERDUE };
      }
    } else {
      stmtRowStatus[i] = { text: "⚠️ Unassigned (Not in Sheet)", color: CONFIG.COLORS.OVERDUE };
    }
  }

  // Overlay matched debits and suggested matches
  stmtDebits.forEach(st => {
    const idx = st.stmtRowIndex;
    if (idx != null && idx >= 0 && idx < stmtData.length) {
      if (st.matched) {
        let payerStr = "";
        if (st.matchType === "EXACT" && st.matchedItems.length > 0) {
          payerStr = st.matchedItems[0].person;
        } else if (st.matchType === "SPLIT") {
          payerStr = st.matchedItems.map(it => it.person).join(" + ") + " (Split)";
        } else if (st.matchType === "CONFIRMED") {
          payerStr = "Confirmed Match";
        }
        stmtRowStatus[idx] = { text: "✅ " + (payerStr || "Assigned"), color: CONFIG.COLORS.PAID };
      } else if (suggestedStmtIds.has(st.id)) {
        stmtRowStatus[idx] = { text: "❓ Needs Confirmation", color: CONFIG.COLORS.DUE_SOON };
      } else {
        const cKey = `${String(st.dateStr)}_${st.amount.toFixed(2)}_${st.desc.substring(0, 30)}`;
        const scriptProps = PropertiesService.getScriptProperties();
        let assignedMap = {};
        try { assignedMap = JSON.parse(scriptProps.getProperty("ASSIGNED_STATEMENT_CHARGES") || "{}"); } catch (e) { }
        if (assignedMap[cKey] && assignedMap[cKey].payer) {
          if (assignedMap[cKey].payer === "NEGLECT") {
            stmtRowStatus[idx] = { text: "🚫 Neglected / Ignored", color: "#eeeeee" };
          } else {
            stmtRowStatus[idx] = { text: "⚠️ Unrecorded (Assigned to " + assignedMap[cKey].payer + ")", color: CONFIG.COLORS.DUE_SOON };
          }
        } else {
          stmtRowStatus[idx] = { text: "⚠️ Unassigned (Not in Sheet)", color: CONFIG.COLORS.OVERDUE };
        }
      }
    }
  });

  stmtSheet.getRange(6, 7).setValue("Assigned Payer / Status")
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, false, false);

  const statusValues = stmtRowStatus.map(s => [s.text]);
  const statusColors = stmtRowStatus.map(s => [s.color]);
  const hRange = stmtSheet.getRange(7, 7, stmtData.length, 1);
  hRange.setValues(statusValues);
  hRange.setBackgrounds(statusColors);
  hRange.setFontWeight("bold");
  hRange.setFontSize(9);
  stmtSheet.autoResizeColumns(1, 7);

  reconSheet.autoResizeColumns(1, 8);
}

function addAssignedChargesToTransactions() {
  SpreadsheetApp.getUi().alert(
    "Safe Mode Active",
    "Transactions and Installments sheets are kept untouched as user-managed sheets.\n\n" +
    "Any charges you assign in Reconciliation are automatically included in the Debt Breakdown sheet with a '⚠️ Missing from Transactions sheet' tag.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  updateLiveDashboard();
}

// ==========================================
// 6. DASHBOARD & 100K AVAILABLE BALANCE
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

  // Read latest statement closing balance if available (batch read B2:F4)
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  if (stmtSheet && stmtSheet.getLastRow() >= 4) {
    const metaVals = stmtSheet.getRange(2, 1, 3, 6).getValues();
    const stmtDateVal = metaVals[0][1]; // B2
    statementDate = parseDateValue(stmtDateVal, tz);
    const dueDateVal = metaVals[0][3]; // D2
    statementDueDate = parseDateValue(dueDateVal, tz);

    const openVal = metaVals[1][1]; // B3
    if (!isNaN(parseFloat(openVal))) openingBalance = parseFloat(openVal);

    const closeVal = metaVals[1][3]; // D3
    if (!isNaN(parseFloat(closeVal))) billedBalance = parseFloat(closeVal);

    const totDebVal = metaVals[1][5]; // F3
    if (!isNaN(parseFloat(totDebVal))) totalDebit = parseFloat(totDebVal);

    const totCredVal = metaVals[2][1]; // B4
    if (!isNaN(parseFloat(totCredVal))) totalCredit = parseFloat(totCredVal);

    if (stmtSheet.getLastRow() >= 7) {
      const lastRow = stmtSheet.getLastRow();
      const txDates = stmtSheet.getRange(7, 2, lastRow - 6, 2).getValues();
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

  // Active month is keyed by the statement's PAYMENT DUE month
  // (e.g. statement closing Aug 31, due Sep 25 -> "2026-09" / "September 2026")
  const activeAnchor = statementDueDate || new Date();
  const activeCycleKey = Utilities.formatDate(activeAnchor, tz, "yyyy-MM");
  const activeCycleLabel = Utilities.formatDate(activeAnchor, tz, "MMMM yyyy");

  // Calculate Remaining Installments Principal Blocked
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

  // Calculate Unbilled New Purchases (purchases belonging to upcoming billing cycles)
  let unbilledNewPurchases = 0;
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (txSheet && txSheet.getLastRow() >= 2) {
    const txData = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 5).getValues();
    txData.forEach(row => {
      const pDate = parseDateValue(row[1], tz);
      const amt = parseFloat(row[4]);
      if (pDate && !isNaN(amt) && amt > 0) {
        const cycleInfo = getBillingCycleForPurchase(pDate, tz);
        // Include purchases belonging to upcoming cycles (including day 31 purchases rolling over!)
        if (cycleInfo && cycleInfo.cycleSortKey > activeCycleKey) {
          unbilledNewPurchases += amt;
        }
      }
    });
  }

  // Check if current month bill has been paid in Payment History
  const historySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_HISTORY);
  let isCurrentBillPaid = false;
  if (historySheet && historySheet.getLastRow() >= 2) {
    const paidMonths = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 1).getValues().flat().map(m => {
      if (m instanceof Date) return Utilities.formatDate(m, tz, "MMMM yyyy").toLowerCase();
      return String(m || "").trim().toLowerCase();
    });

    const currentMonthLabel = Utilities.formatDate(new Date(), tz, "MMMM yyyy").toLowerCase();
    const activeCycleMonthLabel = activeCycleLabel.toLowerCase();
    let stmtDueMonthLabel = "";
    if (statementDueDate) {
      stmtDueMonthLabel = Utilities.formatDate(statementDueDate, tz, "MMMM yyyy").toLowerCase();
    }

    isCurrentBillPaid = paidMonths.includes(currentMonthLabel) ||
      (activeCycleMonthLabel && paidMonths.includes(activeCycleMonthLabel)) ||
      (stmtDueMonthLabel && paidMonths.includes(stmtDueMonthLabel));
  }

  const effectiveBilled = isCurrentBillPaid ? 0 : billedBalance;
  const totalUtilized = effectiveBilled + totalBlockedInstallments + unbilledNewPurchases;
  const availableBalanceNow = Math.max(0, creditLimit - totalUtilized);
  const availableAfterSettlement = Math.max(0, creditLimit - totalBlockedInstallments - unbilledNewPurchases);

  return {
    creditLimit: creditLimit,
    billedBalance: billedBalance,
    openingBalance: openingBalance,
    totalDebit: totalDebit,
    totalCredit: totalCredit,
    statementDate: statementDate,
    statementDueDate: statementDueDate,
    activeCycleKey: activeCycleKey,
    activeCycleLabel: activeCycleLabel,
    minStmtTxDate: minStmtTxDate,
    maxStmtTxDate: maxStmtTxDate,
    stmtPeriod: stmtPeriod,
    isCurrentBillPaid: isCurrentBillPaid,
    totalBlockedInstallments: totalBlockedInstallments,
    unbilledNewPurchases: unbilledNewPurchases,
    totalUtilized: totalUtilized,
    availableBalanceNow: availableBalanceNow,
    availableAfterSettlement: availableAfterSettlement
  };
}

function updateLiveDashboard(options) {
  const opts = (options && typeof options === "object") ? options : {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";

  const transactionsSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const installmentsSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  const historySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_HISTORY);

  // Permanently delete legacy 'Monthly Debts' sheet if it exists
  const oldDebtsSheet = ss.getSheetByName("Monthly Debts");
  if (oldDebtsSheet) {
    try { ss.deleteSheet(oldDebtsSheet); } catch (e) { }
  }

  if (!transactionsSheet || !installmentsSheet || !historySheet) {
    try {
      SpreadsheetApp.getUi().alert(
        "Missing Sheet",
        "Please make sure 'Transactions', 'Installments', and 'Payment History' sheets exist.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (e) {
      Logger.log("Missing required sheets: Transactions, Installments, or Payment History");
    }
    return;
  }

  // NOTE: 'Transactions' and 'Installments' sheets are user-managed inputs.
  // The script NEVER modifies, writes to, or cleans them.

  // Run reconciliation ONLY when not skipped (e.g. avoid erasing/rewriting Reconciliation sheet during onEdit)
  if (!opts.skipRecon) {
    runReconciliation(ss, tz);
  }

  // Read Card Balance & Active Statement info
  // Read Card Balance & Active Statement info
  const cardBal = calculateCardBalance(ss, tz);
  const activeCycleKey = cardBal.activeCycleKey;
  const activeCycleLabel = cardBal.activeCycleLabel;
  const activeStatementDueDate = cardBal.statementDueDate || (cardBal.statementDate ? new Date(cardBal.statementDate.getFullYear(), cardBal.statementDate.getMonth() + 1, 25) : new Date());
  const activeStatementSortKey = activeCycleKey;

  // Read Paid Months Set
  const paidMonthsSet = new Set();
  const lastHistoryRow = historySheet.getLastRow();
  if (lastHistoryRow >= 2) {
    const rawHistory = historySheet.getRange(2, 1, lastHistoryRow - 1, 1).getValues();
    rawHistory.forEach(row => {
      const val = row[0];
      if (!val) return;
      if (val instanceof Date && !isNaN(val.getTime())) {
        paidMonthsSet.add(Utilities.formatDate(val, tz, "MMMM yyyy").toLowerCase());
      } else {
        paidMonthsSet.add(String(val).trim().toLowerCase());
      }
    });
  }

  const allDebts = {};
  const debtLineItems = [];
  // Separate forecast bucket for Monthly Overview (calendar rule: day 31 rolls into the following due month)
  const forecastDebts = {};
  const forecastLineItems = [];

  function addDebtItem(args) {
    addItemToBucket(allDebts, debtLineItems, args);
  }

  function addForecastItem(args) {
    addItemToBucket(forecastDebts, forecastLineItems, args);
  }

  function addItemToBucket(allDebts, debtLineItems, { cycleSortKey, cycleLabel, dueDate, dueDateLabel, person, category, desc, installmentInfo, purchaseDate, originalAmount, amount, isMissingFromSheet, note }) {
    if (!cycleSortKey || !person) return;
    const normalizedPerson = normalizePersonName(person);
    if (!normalizedPerson || normalizedPerson.toLowerCase() === "shared") return;

    if (!allDebts[cycleSortKey]) {
      allDebts[cycleSortKey] = {
        label: cycleLabel,
        cycleSortKey: cycleSortKey,
        dueDate: dueDate,
        dueDateLabel: dueDateLabel,
        total: 0,
        purchasesTotal: 0,
        installmentsTotal: 0,
        missingFromSheetTotal: 0,
        people: {},
        peopleBreakdown: {}
      };
    }

    allDebts[cycleSortKey].total += amount;
    if (category === "Purchase") {
      allDebts[cycleSortKey].purchasesTotal += amount;
    } else if (category === "Installment") {
      allDebts[cycleSortKey].installmentsTotal += amount;
    } else if (category === "MissingStatementDebit" || isMissingFromSheet) {
      allDebts[cycleSortKey].missingFromSheetTotal += amount;
    }

    allDebts[cycleSortKey].people[normalizedPerson] = (allDebts[cycleSortKey].people[normalizedPerson] || 0) + amount;

    if (!allDebts[cycleSortKey].peopleBreakdown[normalizedPerson]) {
      allDebts[cycleSortKey].peopleBreakdown[normalizedPerson] = { purchases: 0, installments: 0, missingFromSheet: 0, total: 0 };
    }
    if (category === "Purchase") {
      allDebts[cycleSortKey].peopleBreakdown[normalizedPerson].purchases += amount;
    } else if (category === "Installment") {
      allDebts[cycleSortKey].peopleBreakdown[normalizedPerson].installments += amount;
    } else if (category === "MissingStatementDebit" || isMissingFromSheet) {
      allDebts[cycleSortKey].peopleBreakdown[normalizedPerson].missingFromSheet += amount;
    }
    allDebts[cycleSortKey].peopleBreakdown[normalizedPerson].total += amount;

    debtLineItems.push({
      sortKey: cycleSortKey,
      cycleSortKey: cycleSortKey,
      cycleLabel: cycleLabel,
      dueMonthLabel: cycleLabel,
      dueDate: dueDate,
      dueDateLabel: dueDateLabel,
      person: normalizedPerson,
      category: category,
      desc: desc,
      installmentInfo: installmentInfo || "-",
      purchaseDate: purchaseDate,
      originalAmount: originalAmount,
      amount: amount,
      isMissingFromSheet: !!isMissingFromSheet,
      note: note || ""
    });
  }

  // 1. Process One-Time Payments (Read-only from Transactions)
  const scriptProps = PropertiesService.getScriptProperties();
  let neglectedTxKeys = new Set();
  try {
    const rawNeglected = JSON.parse(scriptProps.getProperty("NEGLECTED_TRANSACTIONS") || "[]");
    if (Array.isArray(rawNeglected)) {
      // Auto-clean any day-31 / rollover purchases that were mistakenly flagged as neglected from the active statement
      const cleaned = rawNeglected.filter(k => !k.includes("2026-08-31") && !k.includes("1010.00"));
      if (cleaned.length !== rawNeglected.length) {
        safeSetScriptProperty("NEGLECTED_TRANSACTIONS", JSON.stringify(cleaned));
      }
      neglectedTxKeys = new Set(cleaned);
    }
  } catch (err) { }

  const stmtPeriod = cardBal.stmtPeriod || getStatementPeriod(cardBal.statementDate, cardBal.minStmtTxDate, cardBal.maxStmtTxDate);
  const hasStatement = !!cardBal.statementDate;
  const existingTransactionsForDedupe = [];
  if (transactionsSheet.getLastRow() >= 2) {
    const txData = transactionsSheet.getRange(2, 1, transactionsSheet.getLastRow() - 1, 5).getValues();
    txData.forEach(row => {
      const rawDate = row[1];
      const desc = String(row[2] || "").trim();
      const rawPerson = row[3];
      const rawAmount = row[4];
      if (!rawDate || !rawPerson) return;

      const purchaseDate = parseDateValue(rawDate, tz);
      const amount = parseFloat(rawAmount);
      if (!purchaseDate || isNaN(amount) || amount <= 0) return;

      // Calendar billing rule (day 31 rolls over: Aug 31 -> October 2026, due Oct 25)
      const cycleInfo = getBillingCycleForPurchase(purchaseDate, tz);
      if (!cycleInfo) return;

      // Check if user neglected/excluded this transaction in Reconciliation Table 3 for the active statement
      const dStr = Utilities.formatDate(purchaseDate, tz, "yyyy-MM-dd");
      const txKeyNorm = `${dStr}_${amount.toFixed(2)}_${desc.substring(0, 30)}_${rawPerson}`;
      const txKey1 = `${String(purchaseDate)}_${amount.toFixed(2)}_${desc.substring(0, 30)}_${rawPerson}`;
      const txKey2 = `${String(rawDate)}_${amount.toFixed(2)}_${desc.substring(0, 30)}_${rawPerson}`;
      const isNeglectedFromActiveStmt = (neglectedTxKeys.has(txKeyNorm) || neglectedTxKeys.has(txKey1) || neglectedTxKeys.has(txKey2));

      // Calendar billing rule: does this purchase belong to a future forecast cycle?
      const isForecast = cycleInfo.cycleSortKey > activeCycleKey;

      // If a transaction is marked neglected from the active statement, but is NOT a future forecast, skip it.
      // If it IS a future forecast, it MUST still be projected in the Monthly Overview forecast!
      if (isNeglectedFromActiveStmt && !isForecast) {
        return;
      }

      const people = splitPayerNames(rawPerson);
      if (people.length === 0) return;
      const splitAmount = amount / people.length;

      // A) DEBT BREAKDOWN (statement-based): follows the dates in the uploaded PDF statement
      const beforeStmt = !!(stmtPeriod.start && purchaseDate < stmtPeriod.start);
      const inStmtWindow = hasStatement && !beforeStmt && (!stmtPeriod.end || purchaseDate <= stmtPeriod.end);
      let stmtTarget = null;
      if (inStmtWindow && !isNeglectedFromActiveStmt) {
        stmtTarget = {
          cycleSortKey: activeCycleKey,
          cycleLabel: activeCycleLabel,
          dueDate: activeStatementDueDate,
          dueDateLabel: Utilities.formatDate(activeStatementDueDate, tz, "MMMM d, yyyy")
        };
        existingTransactionsForDedupe.push({
          date: purchaseDate,
          rawDate: rawDate,
          desc: desc,
          amount: amount,
          rawPerson: rawPerson
        });
      } else if (!beforeStmt && !isNeglectedFromActiveStmt && cycleInfo.cycleSortKey > activeCycleKey) {
        stmtTarget = cycleInfo;
      } else if (!hasStatement && !isNeglectedFromActiveStmt && cycleInfo.cycleSortKey === activeCycleKey) {
        stmtTarget = cycleInfo;
      }

      people.forEach(p => {
        const base = {
          person: p,
          category: "Purchase",
          desc: desc,
          installmentInfo: "-",
          purchaseDate: purchaseDate,
          originalAmount: amount,
          amount: splitAmount,
          isMissingFromSheet: false
        };
        if (stmtTarget) {
          addDebtItem(Object.assign({}, base, {
            cycleSortKey: stmtTarget.cycleSortKey,
            cycleLabel: stmtTarget.cycleLabel,
            dueDate: stmtTarget.dueDate,
            dueDateLabel: stmtTarget.dueDateLabel
          }));
        }
        if (isForecast) {
          const fItem = Object.assign({}, base, {
            cycleSortKey: cycleInfo.cycleSortKey,
            cycleLabel: cycleInfo.cycleLabel,
            dueDate: cycleInfo.dueDate,
            dueDateLabel: cycleInfo.dueDateLabel
          });
          addForecastItem(fItem);
          if (!stmtTarget || stmtTarget.cycleSortKey !== cycleInfo.cycleSortKey) {
            addDebtItem(fItem);
          }
        }
      });
    });
  }

  // 2. Process Installments (Read-only from Installments, NO WRITES)
  if (installmentsSheet.getLastRow() >= 2) {
    const instData = installmentsSheet.getRange(2, 1, installmentsSheet.getLastRow() - 1, 11).getValues();

    instData.forEach(row => {
      const rawDate = row[1];
      const desc = String(row[2] || "").trim();
      const durationMonths = parseInt(row[4], 10);
      const emi = parseFloat(row[7]);
      const rawPayer = row[8];

      if (!rawDate || isNaN(durationMonths) || durationMonths <= 0 || isNaN(emi) || emi <= 0 || !rawPayer) {
        return;
      }

      const purchaseDate = parseDateValue(rawDate, tz);
      if (!purchaseDate) return;

      const people = splitPayerNames(rawPayer);
      if (people.length === 0) return;
      const splitEmi = emi / people.length;

      for (let i = 0; i < durationMonths; i++) {
        const instCycle = getBillingCycleForInstallment(purchaseDate, i, tz);
        if (!instCycle) continue;
        if (instCycle.cycleSortKey < activeCycleKey) continue;

        const instInfo = `${i + 1} of ${durationMonths}`;

        people.forEach(p => {
          const instItem = {
            cycleSortKey: instCycle.cycleSortKey,
            cycleLabel: instCycle.cycleLabel,
            dueDate: instCycle.dueDate,
            dueDateLabel: instCycle.dueDateLabel,
            person: p,
            category: "Installment",
            desc: desc,
            installmentInfo: instInfo,
            purchaseDate: purchaseDate,
            originalAmount: emi,
            amount: splitEmi,
            isMissingFromSheet: false
          };
          addDebtItem(instItem);
          if (instCycle.cycleSortKey > activeCycleKey) {
            addForecastItem(Object.assign({}, instItem));
          }
        });
      }
    });
  }

  // 3. Process Assigned Statement Charges (From Reconciliation Table 2)
  // Read all assigned rows directly from visible Reconciliation sheet Table 2 as an ARRAY (NO Map collapsing!)
  const assignedChargesList = [];
  const reconSheet = ss.getSheetByName(CONFIG.SHEETS.RECONCILIATION);
  if (reconSheet && reconSheet.getLastRow() >= 5) {
    const lastRow = reconSheet.getLastRow();
    const colA = reconSheet.getRange(1, 1, lastRow, 1).getValues();
    let table2Start = -1;
    let table3Start = -1;
    for (let i = 0; i < colA.length; i++) {
      const txt = String(colA[i][0] || "");
      if (txt.includes("MISSING from Sheet")) table2Start = i + 1;
      if (txt.includes("NOT Found on Bank Statement") || txt.includes("Transactions in Sheet NOT")) table3Start = i + 1;
    }
    if (table2Start > 0) {
      const endRow = (table3Start > table2Start) ? table3Start - 1 : lastRow;
      const numRows = endRow - (table2Start + 1);
      if (numRows > 0) {
        const rows = reconSheet.getRange(table2Start + 2, 1, numRows, 6).getValues();
        rows.forEach(r => {
          const itemNum = String(r[0] || "");
          const rawDate = r[1];
          const desc = String(r[2] || "").trim();
          const amt = parseFloat(r[3]);
          const payer = String(r[4] || "").trim();
          const note = String(r[5] || "").trim();
          if (!isNaN(amt) && amt > 0 && desc && payer) {
            assignedChargesList.push({
              itemNum: itemNum,
              rawDate: rawDate,
              desc: desc,
              amount: amt,
              payer: payer,
              note: note
            });
          }
        });
      }
    }
  }

  // Fallback: If Reconciliation sheet had no rows, read from ScriptProperties
  if (assignedChargesList.length === 0) {
    let assignedMap = {};
    try {
      assignedMap = JSON.parse(scriptProps.getProperty("ASSIGNED_STATEMENT_CHARGES") || "{}");
    } catch (err) { }
    Object.keys(assignedMap).forEach(k => {
      const item = assignedMap[k];
      if (item && item.amount > 0 && item.desc && item.payer) {
        assignedChargesList.push(item);
      }
    });
  }

  let assignedMissingCount = 0;
  let assignedMissingTotal = 0;
  const reconAssignedList = [];

  assignedChargesList.forEach(item => {
    const payerStr = String(item.payer || "").trim();
    if (!payerStr || isNeglectedPayer(payerStr)) {
      return;
    }

    const people = splitPayerNames(payerStr);
    if (people.length === 0) return;

    assignedMissingCount++;
    assignedMissingTotal += item.amount;
    reconAssignedList.push(item);

    const chargeDueDate = activeStatementDueDate;
    const chargeDate = parseDateValue(item.rawDate || item.dateStr, tz) || cardBal.statementDate || chargeDueDate;
    const splitAmount = item.amount / people.length;

    people.forEach(p => {
      addDebtItem({
        cycleSortKey: activeCycleKey,
        cycleLabel: activeCycleLabel,
        dueDate: chargeDueDate,
        dueDateLabel: Utilities.formatDate(chargeDueDate, tz, "MMMM d, yyyy"),
        person: p,
        category: "MissingStatementDebit",
        desc: item.desc || "Statement Charge",
        installmentInfo: "-",
        purchaseDate: chargeDate,
        originalAmount: item.amount,
        amount: splitAmount,
        isMissingFromSheet: true,
        note: item.note || cleanMerchantName(item.desc)
      });
    });
  });

  // Collect and sort unique people: combine all known sheet payers + any in debtLineItems + historical people
  const allKnownPeople = getAllUniquePayers(ss);
  const allPeopleSet = new Set(allKnownPeople);
  debtLineItems.forEach(it => {
    if (it.person && it.person.toLowerCase() !== "shared" && !isNeglectedPayer(it.person)) {
      allPeopleSet.add(it.person);
    }
  });
  ["Abdo", "Dad", "Mai", "Mido", "Muhanad", "Mum", "Zoza"].forEach(p => allPeopleSet.add(p));
  const sortedPeople = Array.from(allPeopleSet)
    .filter(p => p && !isNeglectedPayer(p) && p.toLowerCase() !== "shared")
    .sort((a, b) => a.localeCompare(b));

  // 4. Render Dedicated 'Debt Breakdown' Sheet (Active Statement ONLY)
  renderDebtBreakdownSheet(ss, tz, allDebts, debtLineItems, sortedPeople, paidMonthsSet, cardBal, activeCycleKey, activeCycleLabel, activeStatementDueDate, assignedMissingCount, assignedMissingTotal);

  // 5. Render Dedicated 'Monthly Overview' Sheet (Upcoming Forecasts & Previous Cycles)
  renderMonthlyOverviewSheet(ss, tz, allDebts, debtLineItems, sortedPeople, paidMonthsSet, activeCycleKey, activeCycleLabel, activeStatementDueDate, forecastDebts);

  // 6. Render Dedicated 'Audit & Differences' Sheet
  renderAuditDifferencesSheet(ss, tz, cardBal, allDebts, debtLineItems, activeCycleKey, activeCycleLabel, activeStatementDueDate, reconAssignedList, neglectedTxKeys, existingTransactionsForDedupe, sortedPeople, assignedChargesList);
}

function renderDebtBreakdownSheet(ss, tz, allDebts, debtLineItems, sortedPeople, paidMonthsSet, cardBal, activeCycleKey, activeCycleLabel, activeStatementDueDate, assignedMissingCount, assignedMissingTotal) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.DEBT_BREAKDOWN);
  sheet.clear();
  sheet.setHiddenGridlines(false);

  const activeSortKey = activeCycleKey || Utilities.formatDate(activeStatementDueDate, tz, "yyyy-MM");
  const activeMonthLabel = activeCycleLabel || Utilities.formatDate(activeStatementDueDate, tz, "MMMM yyyy");
  const activeDueDateFormatted = Utilities.formatDate(activeStatementDueDate, tz, "EEEE, MMMM d, yyyy");
  const activeData = allDebts[activeSortKey] || {
    label: activeMonthLabel,
    total: 0,
    purchasesTotal: 0,
    installmentsTotal: 0,
    missingFromSheetTotal: 0,
    people: {},
    peopleBreakdown: {}
  };

  const isCurrentPaid = paidMonthsSet.has(activeMonthLabel.toLowerCase());
  const bankBill = cardBal.billedBalance || 0;
  const diff = activeData.total - bankBill;

  // Banner
  sheet.getRange("A1:I1").merge()
    .setValue("💳 NBE Credit Card — Active Statement Debt Breakdown (" + activeMonthLabel + ")")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  sheet.getRange("A2:I2").merge()
    .setValue("Active billing cycle payment breakdown for the latest uploaded statement. All debt divided by family member.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);

  let curRow = 4;

  // ==========================================
  // ACTIVE STATEMENT RECONCILIATION SUMMARY CARD
  // ==========================================
  sheet.getRange(curRow, 1, 1, 6).merge()
    .setValue("📊 Active Statement Overview & Audit Reconciliation")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  const stmtDateStr = cardBal.statementDate ? Utilities.formatDate(cardBal.statementDate, tz, "yyyy-MM-dd") : "N/A";
  let statusBadge = "✅ Perfect Match (0.00 EGP difference)";
  let statusColor = CONFIG.COLORS.PAID;
  if (Math.abs(diff) >= 1.00) {
    if (diff > 0) {
      statusBadge = "ℹ️ Family debts exceed bank bill by +" + diff.toFixed(2) + " EGP";
      statusColor = CONFIG.COLORS.DUE_SOON;
    } else {
      statusBadge = "⚠️ Bank bill exceeds recorded shares by -" + Math.abs(diff).toFixed(2) + " EGP (Missing items)";
      statusColor = CONFIG.COLORS.OVERDUE;
    }
  }

  const kpiRows = [
    ["Statement Date:", stmtDateStr, "Bank Closing Balance (Must Pay):", bankBill, "Reconciliation Status:", statusBadge],
    ["Payment Due Date:", activeDueDateFormatted, "Sum of Family Shares Accounted For:", activeData.total, "Audit Difference (Diff):", diff],
    ["Active Billing Cycle:", activeMonthLabel, "Total One-Time Purchases:", activeData.purchasesTotal, "Total Monthly Installments:", activeData.installmentsTotal]
  ];

  sheet.getRange(curRow, 1, kpiRows.length, 6).setValues(kpiRows);
  sheet.getRange(curRow, 1, kpiRows.length, 1).setFontWeight("bold").setFontSize(9);
  sheet.getRange(curRow, 3, kpiRows.length, 1).setFontWeight("bold").setFontSize(9);
  sheet.getRange(curRow, 5, kpiRows.length, 1).setFontWeight("bold").setFontSize(9);
  sheet.getRange(curRow, 4, kpiRows.length, 1).setNumberFormat("#,##0.00").setFontWeight("bold");
  sheet.getRange(curRow + 1, 6).setNumberFormat("#,##0.00").setFontWeight("bold");
  sheet.getRange(curRow + 2, 6).setNumberFormat("#,##0.00").setFontWeight("bold");
  sheet.getRange(curRow, 1, kpiRows.length, 6).setBackground(CONFIG.COLORS.PRIMARY_LIGHT).setBorder(true, true, true, true, true, true);
  sheet.getRange(curRow, 6).setBackground(statusColor);
  curRow += kpiRows.length + 1;

  if (assignedMissingCount > 0) {
    sheet.getRange(curRow, 1, 1, 7).merge()
      .setValue("⚠️ NOTICE: " + assignedMissingCount + " charge(s) totaling " + Utilities.formatString("%.2f", assignedMissingTotal) + " EGP are assigned from Bank Statement but MISSING from Transactions sheet!")
      .setFontWeight("bold")
      .setFontColor("#b71c1c")
      .setBackground(CONFIG.COLORS.DUE_SOON);
    curRow += 2;
  }

  if (Math.abs(diff) >= 1.00) {
    let diffNote = "";
    if (diff > 0) {
      diffNote = `🔍 Difference Detective (+${diff.toFixed(2)} EGP Exceeding): ` +
        `Family shares total (${activeData.total.toFixed(2)} EGP) exceeds the bank bill (${bankBill.toFixed(2)} EGP). ` +
        `Check: (1) Reconciliation Table 3 for unbilled sheet purchases that can be neglected, ` +
        `(2) Reconciliation Table 2 if any of the ${assignedMissingCount} assigned missing charges was already recorded in Transactions under another merchant name, ` +
        `or (3) Bank Statement credits/refunds (Total Credit: ${cardBal.totalCredit ? cardBal.totalCredit.toFixed(2) + " EGP" : "0.00 EGP"}).`;
    } else {
      diffNote = `🔍 Difference Detective (-${Math.abs(diff).toFixed(2)} EGP Shortfall): Bank bill is higher by ${Math.abs(diff).toFixed(2)} EGP. Check Reconciliation Table 2 for unassigned statement charges.`;
    }
    sheet.getRange(curRow, 1, 1, 7).merge()
      .setValue(diffNote)
      .setFontWeight("bold")
      .setFontSize(9)
      .setFontColor("#b06000")
      .setBackground("#fffde7")
      .setWrap(true);
    curRow += 2;
  }

  // ==========================================
  // TABLE 1: 👥 ACTIVE STATEMENT BILL (Divided by Person in One Table)
  // ==========================================
  sheet.getRange(curRow, 1, 1, 7).merge()
    .setValue("👥 Active Statement Bill (" + activeMonthLabel + ") — Family Shares Divided in One Table")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");
  curRow++;

  const summaryHeaders = [
    "Person",
    "🛒 Purchases (EGP)",
    "📦 Installments (EGP)",
    "⚠️ Missing from Sheet (EGP)",
    "Total Share Due (EGP)",
    "% of Bank Bill",
    "Payment Status"
  ];

  sheet.getRange(curRow, 1, 1, summaryHeaders.length).setValues([summaryHeaders])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, true, true);
  curRow++;

  const activePeople = sortedPeople.filter(p => {
    const bk = activeData.peopleBreakdown[p];
    return bk && bk.total > 0;
  });

  const sumRows = activePeople.map(p => {
    const bk = activeData.peopleBreakdown[p] || { purchases: 0, installments: 0, missingFromSheet: 0, total: 0 };
    const pct = bankBill > 0 ? (bk.total / bankBill) : (activeData.total > 0 ? bk.total / activeData.total : 0);
    const pStatus = isCurrentPaid ? "✅ Paid" : "🕒 Pending Payment";
    return [
      p,
      bk.purchases,
      bk.installments,
      bk.missingFromSheet,
      bk.total,
      pct,
      pStatus
    ];
  });

  if (sumRows.length > 0) {
    sheet.getRange(curRow, 1, sumRows.length, summaryHeaders.length).setValues(sumRows);
    sheet.getRange(curRow, 2, sumRows.length, 4).setNumberFormat("#,##0.00");
    sheet.getRange(curRow, 6, sumRows.length, 1).setNumberFormat("0.0%");
    sheet.getRange(curRow, 1, sumRows.length, 1).setFontWeight("bold");

    for (let i = 0; i < sumRows.length; i++) {
      if (sumRows[i][3] > 0) {
        // Highlight cell in Missing column
        sheet.getRange(curRow + i, 4).setBackground(CONFIG.COLORS.DUE_SOON).setFontWeight("bold");
      }
      sheet.getRange(curRow + i, 7).setBackground(isCurrentPaid ? CONFIG.COLORS.PAID : CONFIG.COLORS.UPCOMING);
    }
    curRow += sumRows.length;
  }

  // Total Row for Active Bill
  const totalPct = bankBill > 0 ? (activeData.total / bankBill) : 1.00;
  const activeTotalRow = [
    "TOTAL ACTIVE STATEMENT BILL",
    activeData.purchasesTotal,
    activeData.installmentsTotal,
    activeData.missingFromSheetTotal,
    activeData.total,
    totalPct,
    isCurrentPaid ? "✅ Settled with NBE" : "⚠️ Must Pay NBE"
  ];

  sheet.getRange(curRow, 1, 1, summaryHeaders.length).setValues([activeTotalRow])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT)
    .setBorder(true, true, true, true, true, true);
  sheet.getRange(curRow, 2, 1, 4).setNumberFormat("#,##0.00");
  sheet.getRange(curRow, 6, 1, 1).setNumberFormat("0.0%");
  curRow += 3;

  // ==========================================
  // TABLE 2: 📋 ACTIVE STATEMENT ITEMIZED CHARGES (Every Single Charge)
  // ==========================================
  sheet.getRange(curRow, 1, 1, 9).merge()
    .setValue("📋 Itemized Charges for " + activeMonthLabel + " Statement (Every Charge Making Up This Bill)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  const itemHeaders = [
    "#",
    "Person",
    "Category",
    "Merchant / Description",
    "Installment Progress",
    "Charge Date",
    "Person's Share (EGP)",
    "Original Amount (EGP)",
    "Source / Audit Note"
  ];

  sheet.getRange(curRow, 1, 1, itemHeaders.length).setValues([itemHeaders])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, true, true);
  curRow++;

  const headerRowForFilter = curRow - 1;

  // Filter items strictly for the active statement
  const activeItems = debtLineItems.filter(it => it.sortKey === activeSortKey);
  // Sort: Missing first (to highlight), then Installments, then Purchases
  activeItems.sort((a, b) => {
    if (a.isMissingFromSheet !== b.isMissingFromSheet) return a.isMissingFromSheet ? -1 : 1;
    if (a.person !== b.person) return a.person.localeCompare(b.person);
    if (a.category !== b.category) return a.category === "Installment" ? -1 : 1;
    return a.desc.localeCompare(b.desc);
  });

  const regRows = [];
  const regColors = [];

  activeItems.forEach((it, idx) => {
    let catLabel = "🛒 One-Time Purchase";
    let sourceNote = "Transactions Sheet";
    let rowColor = null;

    if (it.isMissingFromSheet) {
      catLabel = "⚠️ Statement Charge";
      sourceNote = "⚠️ MISSING from Transactions sheet (Assigned from Bank Statement)";
      rowColor = "#fff3cd"; // Alert yellow
    } else if (it.category === "Installment") {
      catLabel = "📦 Installment";
      sourceNote = "Installments Sheet";
      rowColor = CONFIG.COLORS.ACCENT_BLUE;
    }

    regRows.push([
      idx + 1,
      it.person,
      catLabel,
      it.desc,
      it.installmentInfo,
      it.purchaseDate,
      it.amount,
      it.originalAmount,
      sourceNote
    ]);
    regColors.push(rowColor);
  });

  if (regRows.length > 0) {
    const regRange = sheet.getRange(curRow, 1, regRows.length, itemHeaders.length);
    regRange.setValues(regRows);

    sheet.getRange(curRow, 6, regRows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(curRow, 7, regRows.length, 2).setNumberFormat("#,##0.00");
    sheet.getRange(curRow, 1, regRows.length, 1).setFontStyle("italic").setFontColor(CONFIG.COLORS.TEXT_MUTED);

    for (let r = 0; r < regRows.length; r++) {
      if (regColors[r]) {
        sheet.getRange(curRow + r, 1, 1, itemHeaders.length).setBackground(regColors[r]);
      }
      if (regRows[r][8].includes("MISSING")) {
        sheet.getRange(curRow + r, 9).setFontWeight("bold").setFontColor("#b71c1c");
      }
    }

    try {
      const existingFilter = sheet.getFilter();
      if (existingFilter) existingFilter.remove();
      sheet.getRange(headerRowForFilter, 1, regRows.length + 1, itemHeaders.length).createFilter();
    } catch (e) { }

    curRow += regRows.length + 3;
  } else {
    sheet.getRange(curRow, 1, 1, 9).merge().setValue("No charges found for this statement cycle.");
    curRow += 3;
  }

  // Footer Link to Monthly Overview Sheet
  sheet.getRange(curRow, 1, 1, 9).merge()
    .setValue("👉 Note: For upcoming billing cycle forecasts and previous months (March – August 2026), please switch to the 'Monthly Overview' sheet.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);
  curRow += 2;

  sheet.autoResizeColumns(1, 9);
}

function getPurchaseDateWindow(dueYear, dueMonth) {
  let endMonth = dueMonth - 1;
  let endYear = dueYear;
  if (endMonth < 1) {
    endMonth = 12;
    endYear--;
  }

  let endDay = 30;
  if (endMonth === 2) {
    const isLeap = (endYear % 4 === 0 && endYear % 100 !== 0) || (endYear % 400 === 0);
    endDay = isLeap ? 29 : 28;
  }

  let prevPrevMonth = endMonth - 1;
  let prevPrevYear = endYear;
  if (prevPrevMonth < 1) {
    prevPrevMonth = 12;
    prevPrevYear--;
  }

  const daysInPrevPrev = new Date(prevPrevYear, prevPrevMonth, 0).getDate();
  let startYear, startMonth, startDay;
  if (daysInPrevPrev === 31) {
    startYear = prevPrevYear;
    startMonth = prevPrevMonth;
    startDay = 31;
  } else {
    startYear = endYear;
    startMonth = endMonth;
    startDay = 1;
  }
  return { startYear, startMonth, startDay, endYear, endMonth, endDay };
}

function renderMonthlyOverviewSheet(ss, tz, allDebts, debtLineItems, sortedPeople, paidMonthsSet, activeCycleKey, activeCycleLabel, activeStatementDueDate, forecastDebts) {
  const upcomingSource = forecastDebts || allDebts;
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.MONTHLY_OVERVIEW);
  sheet.clear();
  sheet.setHiddenGridlines(false);

  const activeSortKey = activeCycleKey || Utilities.formatDate(activeStatementDueDate, tz, "yyyy-MM");
  const activeMonthLabel = activeCycleLabel || Utilities.formatDate(activeStatementDueDate, tz, "MMMM yyyy");
  const activeData = allDebts[activeSortKey] || { total: 0, purchasesTotal: 0, installmentsTotal: 0, people: {}, peopleBreakdown: {} };

  // Hardcoded historical months data (Audited March 2026 – August 2026)
  const HISTORICAL_MONTHS = {
    "2026-03": {
      label: "March 2026",
      dueDateLabel: "March 25, 2026",
      totalBill: 13628.57,
      status: "Paid",
      people: {
        "Dad": 1210.33,
        "Mai": 6233.32,
        "Mido": 2726.90,
        "Muhanad": 2560.02,
        "Mum": 420.00,
        "Zoza": 478.00
      }
    },
    "2026-04": {
      label: "April 2026",
      dueDateLabel: "April 25, 2026",
      totalBill: 13991.09,
      status: "Paid",
      people: {
        "Dad": 1339.15,
        "Mai": 6549.51,
        "Mido": 4794.27,
        "Muhanad": 106.17,
        "Mum": 420.00,
        "Zoza": 782.00
      }
    },
    "2026-05": {
      label: "May 2026",
      dueDateLabel: "May 25, 2026",
      totalBill: 29851.56,
      status: "Paid",
      people: {
        "Abdo": 17143.00,
        "Dad": 1445.33,
        "Mai": 6014.71,
        "Mido": 4250.35,
        "Muhanad": 106.17,
        "Mum": 420.00,
        "Zoza": 472.00
      }
    },
    "2026-06": {
      label: "June 2026",
      dueDateLabel: "June 25, 2026",
      totalBill: 10473.17,
      status: "Paid",
      people: {
        "Abdo": 335.00,
        "Dad": 1210.33,
        "Mai": 2463.61,
        "Mido": 4596.09,
        "Muhanad": 106.17,
        "Mum": 1244.07,
        "Zoza": 517.90
      }
    },
    "2026-07": {
      label: "July 2026",
      dueDateLabel: "July 25, 2026",
      totalBill: 28379.62,
      status: "Paid",
      people: {
        "Dad": 1350.33,
        "Mai": 21696.49,
        "Mido": 4214.56,
        "Muhanad": 106.17,
        "Mum": 494.07,
        "Zoza": 518.00
      }
    },
    "2026-08": {
      label: "August 2026",
      dueDateLabel: "September 25, 2026",
      totalBill: 22867.63,
      status: "Settled with NBE",
      people: {
        "Abdo": 5272.39,
        "Dad": 1210.33,
        "Mai": 6380.39,
        "Mido": 8886.28,
        "Muhanad": 106.17,
        "Mum": 494.07,
        "Zoza": 518.00
      }
    }
  };

  // Banner
  sheet.getRange("A1:K1").merge()
    .setValue("📅 NBE Credit Card — Monthly Overview (Upcoming Forecasts & Previous History)")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  sheet.getRange("A2:K2").merge()
    .setValue("Multi-month debt overview: upcoming future dues forecasted dynamically from Transactions and Installments, alongside audited historical billing cycles.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);

  let curRow = 4;

  // Upcoming keys (strictly future months)
  const upcomingKeys = Object.keys(upcomingSource).filter(k => k > activeSortKey).sort();
  const histKeys = Object.keys(HISTORICAL_MONTHS).filter(k => k < activeSortKey).sort();

  // ==========================================
  // SECTION 1: 📊 MULTI-MONTH COMPARISON MATRIX (Full Timeline View)
  // ==========================================
  const matrixCols = [
    { key: "label", header: "Family Member", type: "header" }
  ];

  histKeys.forEach(k => {
    matrixCols.push({
      key: k,
      header: HISTORICAL_MONTHS[k].label,
      type: "hist"
    });
  });

  matrixCols.push({
    key: activeSortKey,
    header: activeMonthLabel + " (Active)",
    type: "active"
  });

  upcomingKeys.forEach(k => {
    const src = upcomingSource[k] || allDebts[k] || {};
    matrixCols.push({
      key: k,
      header: (src.label || k) + " (Forecast)",
      type: "upcoming"
    });
  });

  sheet.getRange(curRow, 1, 1, matrixCols.length).merge()
    .setValue("📊 Section 1: Multi-Month Comparison Matrix (Timeline Overview)")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");
  curRow++;

  sheet.getRange(curRow, 1, 1, matrixCols.length).merge()
    .setValue("Side-by-side comparison of family shares across all billing cycles: historical audited statements, active statement, and future projections.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);
  curRow += 2;

  // Headers
  const matrixHeaderRow = matrixCols.map(c => c.header);
  sheet.getRange(curRow, 1, 1, matrixCols.length).setValues([matrixHeaderRow])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, true, true);
  curRow++;

  // Row 1: TOTAL MONTH BILL
  const totalBillRow = matrixCols.map((c, idx) => {
    if (idx === 0) return "TOTAL MONTH BILL (EGP)";
    if (c.type === "hist") return HISTORICAL_MONTHS[c.key].totalBill;
    if (c.type === "active") return activeData.total || (HISTORICAL_MONTHS[c.key] ? HISTORICAL_MONTHS[c.key].totalBill : 0);
    if (c.type === "upcoming") {
      const src = upcomingSource[c.key] || allDebts[c.key];
      return src ? src.total : 0;
    }
    return 0;
  });
  sheet.getRange(curRow, 1, 1, matrixCols.length).setValues([totalBillRow])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT)
    .setBorder(true, true, true, true, true, true);
  sheet.getRange(curRow, 2, 1, matrixCols.length - 1).setNumberFormat("#,##0.00");
  curRow++;

  // Row 2: Status
  const statusRow = matrixCols.map((c, idx) => {
    if (idx === 0) return "Cycle Status";
    if (c.type === "hist") return "✅ Settled";
    if (c.type === "active") return (paidMonthsSet.has(activeMonthLabel.toLowerCase()) ? "✅ Settled" : "💳 Active Statement");
    if (c.type === "upcoming") return "🕒 Forecast";
    return "";
  });
  sheet.getRange(curRow, 1, 1, matrixCols.length).setValues([statusRow])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBorder(true, true, true, true, true, true);
  for (let cIdx = 1; cIdx < matrixCols.length; cIdx++) {
    const colType = matrixCols[cIdx].type;
    const bg = colType === "hist" ? CONFIG.COLORS.PAID : (colType === "active" ? CONFIG.COLORS.DUE_SOON : CONFIG.COLORS.UPCOMING);
    sheet.getRange(curRow, cIdx + 1).setBackground(bg);
  }
  curRow++;

  // Person Rows
  const matrixPersonRows = [];
  sortedPeople.forEach(person => {
    const pRow = [person];
    for (let cIdx = 1; cIdx < matrixCols.length; cIdx++) {
      const col = matrixCols[cIdx];
      let val = 0;
      if (col.type === "hist") {
        val = HISTORICAL_MONTHS[col.key].people[person] || 0;
      } else if (col.type === "active") {
        const bk = activeData.peopleBreakdown[person];
        val = bk ? bk.total : (activeData.people[person] || (HISTORICAL_MONTHS[col.key] && HISTORICAL_MONTHS[col.key].people[person] ? HISTORICAL_MONTHS[col.key].people[person] : 0));
      } else if (col.type === "upcoming") {
        const mD = upcomingSource[col.key] || allDebts[col.key];
        if (mD) {
          const bk = mD.peopleBreakdown[person];
          val = bk ? bk.total : (mD.people[person] || 0);
        }
      }
      pRow.push(val);
    }
    matrixPersonRows.push(pRow);
  });

  if (matrixPersonRows.length > 0) {
    sheet.getRange(curRow, 1, matrixPersonRows.length, matrixCols.length).setValues(matrixPersonRows);
    sheet.getRange(curRow, 1, matrixPersonRows.length, 1).setFontWeight("bold");
    sheet.getRange(curRow, 2, matrixPersonRows.length, matrixCols.length - 1).setNumberFormat("#,##0.00");
    sheet.getRange(curRow, 1, matrixPersonRows.length, matrixCols.length).setBorder(true, true, true, true, true, true);
    curRow += matrixPersonRows.length + 3;
  }

  // ==========================================
  // SECTION 2: 🕒 UPCOMING BILLING CYCLES (Forecasted from Transactions & Installments)
  // ==========================================
  sheet.getRange(curRow, 1, 1, 8).merge()
    .setValue("🕒 Section 2: Upcoming Billing Cycles (Forecasted from Transactions & Installments)")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");
  curRow++;

  sheet.getRange(curRow, 1, 1, 8).merge()
    .setValue("Upcoming months dynamically projected from active installments and new purchases recorded in Transactions (including day 31 purchases rolling over).")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);
  curRow += 2;

  if (upcomingKeys.length === 0) {
    sheet.getRange(curRow, 1, 1, 8).merge()
      .setValue("ℹ️ No upcoming installment charges or post-cutoff purchases currently recorded.")
      .setFontStyle("italic")
      .setFontColor(CONFIG.COLORS.TEXT_MUTED);
    curRow += 3;
  } else {
    upcomingKeys.forEach(mKey => {
      const mData = upcomingSource[mKey] || allDebts[mKey] || { total: 0, purchasesTotal: 0, installmentsTotal: 0, peopleBreakdown: {} };
      const mLabel = mData.label || mKey;
      const mDueDateStr = mData.dueDateLabel || (mData.dueDate ? Utilities.formatDate(mData.dueDate, tz, "MMMM d, yyyy") : "25th");

      const mPeople = sortedPeople.filter(p => (mData.peopleBreakdown[p] && mData.peopleBreakdown[p].total > 0));
      const startDataRow = curRow + 2; // header title is curRow, headers are curRow + 1
      const numPeople = mPeople.length;
      const totalRowIndex = startDataRow + numPeople;

      // Header Bar for this month with live forecast total formula referencing the total row
      const titleFormula = numPeople > 0
        ? `="📅 " & "${mLabel} — Due ${mDueDateStr} — Forecast Total: " & TEXT(D${totalRowIndex}, "#,##0.00") & " EGP"`
        : `📅 ${mLabel} — Due ${mDueDateStr} — Forecast Total: 0.00 EGP`;
      sheet.getRange(curRow, 1, 1, 6).merge()
        .setValue(titleFormula)
        .setFontWeight("bold")
        .setFontSize(11)
        .setBackground(CONFIG.COLORS.ACCENT_BLUE);
      curRow++;

      const mHeaders = ["Person", "🛒 Purchases (EGP)", "📦 Installments (EGP)", "Total Due (EGP)", "% of Month Bill", "Status"];
      sheet.getRange(curRow, 1, 1, mHeaders.length).setValues([mHeaders])
        .setFontWeight("bold")
        .setBackground(CONFIG.COLORS.HEADER)
        .setBorder(true, true, true, true, true, true);
      curRow++;

      // Compute purchase cycle date window for this due month
      const mMatch = mKey.match(/^(\d{4})-(\d{2})$/);
      let pFormulaFunc = (personName, rIdx) => (mData.peopleBreakdown[personName]?.purchases || 0);
      if (mMatch) {
        const dYear = parseInt(mMatch[1], 10);
        const dMonth = parseInt(mMatch[2], 10);
        const w = getPurchaseDateWindow(dYear, dMonth);
        pFormulaFunc = (personName, rIdx) => `=SUMIFS('Transactions'!$E:$E, 'Transactions'!$D:$D, $A${rIdx}, 'Transactions'!$B:$B, ">="&DATE(${w.startYear},${w.startMonth},${w.startDay}), 'Transactions'!$B:$B, "<="&DATE(${w.endYear},${w.endMonth},${w.endDay}))`;
      }

      const mRows = mPeople.map((p, idx) => {
        const bk = mData.peopleBreakdown[p] || { purchases: 0, installments: 0, total: 0 };
        const currentRow = startDataRow + idx;
        const purchaseValOrFormula = pFormulaFunc(p, currentRow);
        const totalFormula = `=B${currentRow}+C${currentRow}`;
        const pctFormula = `=IF($D$${totalRowIndex}>0, D${currentRow}/$D$${totalRowIndex}, 0)`;
        return [p, purchaseValOrFormula, bk.installments, totalFormula, pctFormula, "🕒 Upcoming Forecast"];
      });

      if (mRows.length > 0) {
        sheet.getRange(curRow, 1, mRows.length, mHeaders.length).setValues(mRows);
        sheet.getRange(curRow, 2, mRows.length, 3).setNumberFormat("#,##0.00");
        sheet.getRange(curRow, 5, mRows.length, 1).setNumberFormat("0.0%");
        sheet.getRange(curRow, 1, mRows.length, 1).setFontWeight("bold");
        sheet.getRange(curRow, 6, mRows.length, 1).setBackground(CONFIG.COLORS.UPCOMING);
        curRow += mRows.length;
      }

      // Total Row for this month with live SUM formulas
      const purchasesSumFormula = numPeople > 0 ? `=SUM(B${startDataRow}:B${curRow - 1})` : mData.purchasesTotal;
      const installmentsSumFormula = numPeople > 0 ? `=SUM(C${startDataRow}:C${curRow - 1})` : mData.installmentsTotal;
      const totalSumFormula = numPeople > 0 ? `=SUM(D${startDataRow}:D${curRow - 1})` : mData.total;
      const mTotalRow = ["TOTAL (" + mLabel + ")", purchasesSumFormula, installmentsSumFormula, totalSumFormula, 1.00, "🕒 Upcoming Due"];
      sheet.getRange(curRow, 1, 1, mHeaders.length).setValues([mTotalRow])
        .setFontWeight("bold")
        .setBackground(CONFIG.COLORS.PRIMARY_LIGHT)
        .setBorder(true, true, true, true, true, true);
      sheet.getRange(curRow, 2, 1, 3).setNumberFormat("#,##0.00");
      sheet.getRange(curRow, 5, 1, 1).setNumberFormat("0.0%");
      curRow += 2;
    });
  }

  // ==========================================
  // SECTION 3: 📜 HISTORICAL BILLING CYCLES (Audited Statements: March 2026 – August 2026)
  // ==========================================
  sheet.getRange(curRow, 1, 1, 8).merge()
    .setValue("📜 Section 3: Historical Billing Cycles (Audited Statements: March 2026 – August 2026)")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");
  curRow++;

  sheet.getRange(curRow, 1, 1, 8).merge()
    .setValue("Finalized, audited statement numbers and family shares for past billing cycles.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);
  curRow += 2;

  const allHistKeys = Object.keys(HISTORICAL_MONTHS).sort();
  allHistKeys.forEach(hKey => {
    const hist = HISTORICAL_MONTHS[hKey];
    sheet.getRange(curRow, 1, 1, 5).merge()
      .setValue(`📜 ${hist.label} — Total Bill: ${hist.totalBill.toFixed(2)} EGP — ✅ [PAID / SETTLED]`)
      .setFontWeight("bold")
      .setFontSize(11)
      .setBackground(CONFIG.COLORS.PAID);
    curRow++;

    const histHeaders = ["Person", "Amount Owed (EGP)", "% of Total Bill", "Status"];
    sheet.getRange(curRow, 1, 1, histHeaders.length).setValues([histHeaders])
      .setFontWeight("bold")
      .setBackground(CONFIG.COLORS.HEADER)
      .setBorder(true, true, true, true, true, true);
    curRow++;

    const hPeople = Object.keys(hist.people).sort();
    const hRows = hPeople.map(p => {
      const amt = hist.people[p];
      const pct = hist.totalBill > 0 ? (amt / hist.totalBill) : 0;
      return [p, amt, pct, "✅ Paid"];
    });

    if (hRows.length > 0) {
      sheet.getRange(curRow, 1, hRows.length, histHeaders.length).setValues(hRows);
      sheet.getRange(curRow, 2, hRows.length, 1).setNumberFormat("#,##0.00");
      sheet.getRange(curRow, 3, hRows.length, 1).setNumberFormat("0.0%");
      sheet.getRange(curRow, 1, hRows.length, 1).setFontWeight("bold");
      sheet.getRange(curRow, 4, hRows.length, 1).setBackground(CONFIG.COLORS.PAID);
      sheet.getRange(curRow, 1, hRows.length, histHeaders.length).setBorder(true, true, true, true, true, true);
      curRow += hRows.length;
    }

    // Total Row
    const hTotalRow = ["TOTAL BILL (" + hist.label + ")", hist.totalBill, 1.00, "✅ Settled with NBE"];
    sheet.getRange(curRow, 1, 1, histHeaders.length).setValues([hTotalRow])
      .setFontWeight("bold")
      .setBackground(CONFIG.COLORS.PRIMARY_LIGHT)
      .setBorder(true, true, true, true, true, true);
    sheet.getRange(curRow, 2, 1, 1).setNumberFormat("#,##0.00");
    sheet.getRange(curRow, 3, 1, 1).setNumberFormat("0.0%");
    curRow += 3;
  });

  sheet.autoResizeColumns(1, Math.max(9, matrixCols.length));
}

function renderAuditDifferencesSheet(ss, tz, cardBal, allDebts, debtLineItems, activeCycleKey, activeCycleLabel, activeStatementDueDate, reconAssignedList, neglectedTxKeys, existingTxList, sortedPeople, assignedChargesList) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.AUDIT_DIFFERENCES);
  sheet.clear();
  sheet.setHiddenGridlines(false);

  const activeSortKey = activeCycleKey || Utilities.formatDate(activeStatementDueDate, tz, "yyyy-MM");
  const activeMonthLabel = activeCycleLabel || Utilities.formatDate(activeStatementDueDate, tz, "MMMM yyyy");
  const activeData = allDebts[activeSortKey] || {
    label: activeMonthLabel,
    total: 0,
    purchasesTotal: 0,
    installmentsTotal: 0,
    missingFromSheetTotal: 0,
    people: {},
    peopleBreakdown: {}
  };

  const bankBill = cardBal.billedBalance || 0;
  const familyTotal = activeData.total || 0;
  const netDiff = familyTotal - bankBill;

  // Title Banner
  sheet.getRange("A1:H1").merge()
    .setValue("🔍 NBE Credit Card — Cross-Sheet Audit & Discrepancy Detective (" + activeMonthLabel + ")")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  sheet.getRange("A2:H2").merge()
    .setValue("Automated reconciliation & audit engine comparing Bank Statement, Transactions, Installments, Reconciliation, and Debt Breakdown.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);

  let curRow = 4;

  // ==========================================
  // 1. EXECUTIVE SUMMARY & BALANCE EQUATION
  // ==========================================
  sheet.getRange(curRow, 1, 1, 8).merge()
    .setValue("📊 Executive Cross-Sheet Balance Summary")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  let neglectedDebitsTotal = 0;
  let neglectedDebitsCount = 0;
  (assignedChargesList || []).forEach(item => {
    if (isNeglectedPayer(item.payer)) {
      neglectedDebitsTotal += item.amount;
      neglectedDebitsCount++;
    }
  });

  let statusBadge = "✅ Perfect Match (0.00 EGP difference)";
  let statusColor = CONFIG.COLORS.PAID;
  if (Math.abs(netDiff) >= 0.05) {
    if (Math.abs(netDiff + neglectedDebitsTotal) < 0.50) {
      statusBadge = "✅ Balanced (Difference is " + neglectedDebitsTotal.toFixed(2) + " EGP Neglected Debits)";
      statusColor = CONFIG.COLORS.PAID;
    } else if (netDiff > 0) {
      statusBadge = "⚠️ Family Shares Exceed Bank Bill by +" + netDiff.toFixed(2) + " EGP";
      statusColor = CONFIG.COLORS.DUE_SOON;
    } else {
      statusBadge = "🚨 Bank Bill Exceeds Shares by -" + Math.abs(netDiff).toFixed(2) + " EGP (Missing Items!)";
      statusColor = CONFIG.COLORS.OVERDUE;
    }
  }

  const kpiData = [
    ["🏦 Bank Closing Balance (Must Pay):", bankBill, "👥 Total Family Shares in Debt Breakdown:", familyTotal, "⚖️ Net Discrepancy:", netDiff, "Status:", statusBadge],
    ["🛒 Purchases Total (Sheet):", activeData.purchasesTotal, "📦 Installments Total (Sheet):", activeData.installmentsTotal, "⚠️ Missing Assigned (Recon Table 2):", activeData.missingFromSheetTotal, "🚫 Neglected Debits (Table 2):", neglectedDebitsTotal],
    ["💳 Statement Total Debits:", cardBal.totalDebit || 0, "💰 Statement Credits / Payments:", cardBal.totalCredit || 0, "Available Balance Now:", cardBal.availableBalanceNow, "Available Post-Settlement:", cardBal.availableAfterSettlement]
  ];

  sheet.getRange(curRow, 1, kpiData.length, 8).setValues(kpiData);
  sheet.getRange(curRow, 1, kpiData.length, 8).setFontSize(9).setBackground(CONFIG.COLORS.PRIMARY_LIGHT).setBorder(true, true, true, true, true, true);
  for (let c of [1, 3, 5, 7]) {
    sheet.getRange(curRow, c, kpiData.length, 1).setFontWeight("bold");
  }
  for (let c of [2, 4, 6, 8]) {
    sheet.getRange(curRow, c, kpiData.length, 1).setNumberFormat("#,##0.00").setFontWeight("bold");
  }
  sheet.getRange(curRow, 8).setBackground(statusColor).setFontWeight("bold");
  curRow += kpiData.length + 2;

  // ==========================================
  // TABLE 1: 👥 PERSON-BY-PERSON CROSS-SHEET COMPARISON (Reconciliation vs Debt Breakdown)
  // ==========================================
  sheet.getRange(curRow, 1, 1, 8).merge()
    .setValue("👥 Table 1: Person-by-Person Cross-Sheet Comparison (Reconciliation vs Debt Breakdown)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");
  curRow++;

  sheet.getRange(curRow, 1, 1, 8).merge()
    .setValue("Compares what was assigned and recorded across input sheets (Transactions, Installments, and Reconciliation Table 2) against what is currently reflected in Debt Breakdown.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);
  curRow++;

  const t1Headers = [
    "Person",
    "🛒 Purchases (Transactions)",
    "📦 Installments (Installments)",
    "⚠️ Assigned Debits (Recon Table 2)",
    "Expected Total (EGP)",
    "Reflected in Debt Breakdown",
    "Discrepancy (Diff EGP)",
    "Cross-Sheet Audit Status"
  ];
  sheet.getRange(curRow, 1, 1, t1Headers.length).setValues([t1Headers]).setFontWeight("bold").setBackground(CONFIG.COLORS.HEADER).setBorder(true, true, true, true, true, true);
  curRow++;

  const personComparisonRows = [];
  const personRowColors = [];

  let grandExpPurchases = 0;
  let grandExpInstallments = 0;
  let grandExpReconAssigned = 0;
  let grandExpTotal = 0;
  let grandReflectedTotal = 0;
  let grandPersonDiff = 0;

  // Filter to real family members who have activity or expected debt (never include Neglect or Ignore!)
  const auditPeople = (sortedPeople || []).filter(p => {
    if (!p || isNeglectedPayer(p) || p.toLowerCase() === "shared") return false;
    const bk = (activeData.peopleBreakdown && activeData.peopleBreakdown[p]) || { purchases: 0, installments: 0, missingFromSheet: 0, total: 0 };
    let pReconAssigned = 0;
    (assignedChargesList || []).forEach(item => {
      if (isNeglectedPayer(item.payer)) return;
      const payers = splitPayerNames(item.payer);
      if (payers.includes(p)) {
        pReconAssigned += (item.amount / payers.length);
      }
    });
    return (bk.purchases > 0 || bk.installments > 0 || pReconAssigned > 0 || bk.total > 0);
  });

  auditPeople.forEach(p => {
    const bk = (activeData.peopleBreakdown && activeData.peopleBreakdown[p]) || { purchases: 0, installments: 0, missingFromSheet: 0, total: 0 };
    const pPurchases = bk.purchases || 0;
    const pInstallments = bk.installments || 0;

    // Calculate this person's assigned debits directly from assignedChargesList (Table 2)
    let pReconAssigned = 0;
    (assignedChargesList || []).forEach(item => {
      if (isNeglectedPayer(item.payer)) return;
      const payers = splitPayerNames(item.payer);
      if (payers.includes(p)) {
        pReconAssigned += (item.amount / payers.length);
      }
    });

    const pExpectedTotal = pPurchases + pInstallments + pReconAssigned;
    const pReflectedTotal = bk.total || 0;
    const pDiff = pReflectedTotal - pExpectedTotal;

    grandExpPurchases += pPurchases;
    grandExpInstallments += pInstallments;
    grandExpReconAssigned += pReconAssigned;
    grandExpTotal += pExpectedTotal;
    grandReflectedTotal += pReflectedTotal;
    grandPersonDiff += pDiff;

    let pStatus = "✅ In Sync (0.00 EGP)";
    let rowColor = CONFIG.COLORS.PAID;

    if (Math.abs(pDiff) >= 0.05) {
      if (pDiff < -0.05) {
        pStatus = `🚨 Missing in Debt Breakdown! (${Math.abs(pDiff).toFixed(2)} EGP unreflected)`;
        rowColor = CONFIG.COLORS.OVERDUE;
      } else {
        pStatus = `⚠️ Exceeds Expected by +${pDiff.toFixed(2)} EGP`;
        rowColor = CONFIG.COLORS.DUE_SOON;
      }
    }

    personComparisonRows.push([
      p,
      pPurchases,
      pInstallments,
      pReconAssigned,
      pExpectedTotal,
      pReflectedTotal,
      pDiff,
      pStatus
    ]);
    personRowColors.push(rowColor);
  });

  if (personComparisonRows.length > 0) {
    sheet.getRange(curRow, 1, personComparisonRows.length, t1Headers.length).setValues(personComparisonRows);
    sheet.getRange(curRow, 1, personComparisonRows.length, 1).setFontWeight("bold");
    sheet.getRange(curRow, 2, personComparisonRows.length, 6).setNumberFormat("#,##0.00");

    for (let r = 0; r < personComparisonRows.length; r++) {
      sheet.getRange(curRow + r, 8).setBackground(personRowColors[r]).setFontWeight("bold");
    }
    curRow += personComparisonRows.length;
  }

  // Summary Row for Table 1
  const t1TotalRow = [
    "TOTAL ALL FAMILY MEMBERS",
    grandExpPurchases,
    grandExpInstallments,
    grandExpReconAssigned,
    grandExpTotal,
    grandReflectedTotal,
    grandPersonDiff,
    Math.abs(grandPersonDiff) < 0.05 ? "✅ All Sheets 100% In Sync" : `🚨 Net Discrepancy: ${grandPersonDiff.toFixed(2)} EGP`
  ];
  sheet.getRange(curRow, 1, 1, t1Headers.length).setValues([t1TotalRow])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT)
    .setBorder(true, true, true, true, true, true);
  sheet.getRange(curRow, 2, 1, 6).setNumberFormat("#,##0.00");
  sheet.getRange(curRow, 8).setBackground(Math.abs(grandPersonDiff) < 0.05 ? CONFIG.COLORS.PAID : CONFIG.COLORS.OVERDUE);
  curRow += 3;

  // ==========================================
  // TABLE 2: 📋 RECONCILIATION TABLE 2 CHARGES VS DEBT BREAKDOWN VERIFICATION
  // ==========================================
  sheet.getRange(curRow, 1, 1, 7).merge()
    .setValue("📋 Table 2: Reconciliation Table 2 Charges vs Debt Breakdown Verification (" + (assignedChargesList ? assignedChargesList.length : 0) + " items)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  sheet.getRange(curRow, 1, 1, 7).merge()
    .setValue("Every single charge assigned in Reconciliation Table 2 verified against Debt Breakdown line items.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);
  curRow++;

  const t2Headers = ["#", "Date", "Description / Merchant", "Amount (EGP)", "Assigned Payer", "Custom Note", "Status in Debt Breakdown"];
  sheet.getRange(curRow, 1, 1, t2Headers.length).setValues([t2Headers]).setFontWeight("bold").setBackground(CONFIG.COLORS.HEADER).setBorder(true, true, true, true, true, true);
  curRow++;

  if (!assignedChargesList || assignedChargesList.length === 0) {
    sheet.getRange(curRow, 1, 1, 7).merge()
      .setValue("No charges are currently assigned in Reconciliation Table 2.")
      .setFontStyle("italic");
    curRow += 2;
  } else {
    const t2Rows = assignedChargesList.map((item, idx) => {
      const isNeg = isNeglectedPayer(item.payer);
      const payers = isNeg ? [] : splitPayerNames(item.payer);

      let statusText = "";
      if (isNeg) {
        statusText = "🚫 Neglected / Excluded from Bill (Correct)";
      } else {
        const isReflected = debtLineItems.some(it => {
          return it.isMissingFromSheet &&
            Math.abs(it.originalAmount - item.amount) < 0.05 &&
            payers.includes(it.person);
        });
        statusText = isReflected ? "✅ Reflected in Debt Breakdown" : "🚨 Missing from Debt Breakdown!";
      }

      return [
        item.itemNum || (idx + 1),
        item.rawDate || item.dateStr,
        item.desc,
        item.amount,
        isNeg ? "🚫 Neglect / Ignore" : item.payer,
        item.note || "-",
        statusText
      ];
    });

    sheet.getRange(curRow, 1, t2Rows.length, t2Headers.length).setValues(t2Rows);
    sheet.getRange(curRow, 2, t2Rows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(curRow, 4, t2Rows.length, 1).setNumberFormat("#,##0.00").setFontWeight("bold");

    for (let r = 0; r < t2Rows.length; r++) {
      if (t2Rows[r][6].includes("Neglected")) {
        sheet.getRange(curRow + r, 1, 1, t2Headers.length).setBackground("#eeeeee");
        sheet.getRange(curRow + r, 7).setFontStyle("italic").setFontColor(CONFIG.COLORS.TEXT_MUTED);
      } else if (t2Rows[r][6].includes("Reflected")) {
        sheet.getRange(curRow + r, 7).setBackground(CONFIG.COLORS.PAID).setFontWeight("bold");
      } else {
        sheet.getRange(curRow + r, 7).setBackground(CONFIG.COLORS.OVERDUE).setFontWeight("bold");
      }
    }
    curRow += t2Rows.length + 3;
  }

  // ==========================================
  // TABLE 3: 📦 INSTALLMENTS VERIFICATION (Installments Sheet vs Bank Statement)
  // ==========================================
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  let stmtInstallmentsBilledTotal = 0;
  let stmtInstallmentCount = 0;
  const stmtInstallments = [];

  if (stmtSheet && stmtSheet.getLastRow() >= 7) {
    const lastR = stmtSheet.getLastRow();
    const rows = stmtSheet.getRange(7, 1, lastR - 6, 7).getValues();
    rows.forEach((r, idx) => {
      const type = String(r[4] || "").trim().toUpperCase();
      const amt = parseFloat(r[5]);
      const desc = String(r[3] || "").trim();
      if (type === "INSTALLMENT" && !isNaN(amt) && amt > 0) {
        stmtInstallmentsBilledTotal += amt;
        stmtInstallmentCount++;

        let instCurrent = null;
        let instTotal = null;
        const mInst = desc.match(/(\d+)\s+OF\s+(\d+)/i);
        if (mInst) {
          instCurrent = parseInt(mInst[1], 10);
          instTotal = parseInt(mInst[2], 10);
        }
        stmtInstallments.push({
          stmtRowIdx: idx,
          desc: desc,
          amount: amt,
          instCurrent: instCurrent,
          instTotal: instTotal
        });
      }
    });
  }

  // Read all installment records from user's Installments sheet
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  const sheetInstallmentRows = [];
  if (instSheet && instSheet.getLastRow() >= 2) {
    const instData = instSheet.getRange(2, 1, instSheet.getLastRow() - 1, 11).getValues();
    instData.forEach((iRow, idx) => {
      const rawDate = iRow[1];
      const desc = String(iRow[2] || "").trim();
      const duration = parseInt(iRow[4], 10);
      const emi = parseFloat(iRow[7]);
      const payer = String(iRow[8] || "").trim();
      const paymentsMade = parseInt(iRow[9], 10) || 0;
      const status = String(iRow[10] || "").toLowerCase();
      const pDate = parseDateValue(rawDate, tz);

      if (!isNaN(emi) && emi > 0 && desc && payer) {
        sheetInstallmentRows.push({
          index: idx,
          rowIdx: idx + 2,
          date: pDate,
          rawDate: rawDate,
          desc: desc,
          duration: duration,
          emi: emi,
          payer: payer,
          paymentsMade: paymentsMade,
          status: status
        });
      }
    });
  }

  // Run matching engine
  const instMatches = matchStatementInstallments(stmtInstallments, sheetInstallmentRows);
  const matchedStmtIdxs = new Set(instMatches.map(m => m.stmtIdx));
  const matchedSheetIdxs = new Set();
  instMatches.forEach(m => {
    (m.sheetRows || []).forEach(r => matchedSheetIdxs.add(r.index));
  });

  // Identify unbilled sheet installments (e.g. Fan bought Aug 28 subject to 55-day policy)
  const unbilledSheetInstallments = sheetInstallmentRows.filter(sh => {
    return !matchedSheetIdxs.has(sh.index) && sh.status !== "completed";
  });

  // Identify statement installments not found in sheet
  const unrecordedStmtInstallments = stmtInstallments.filter((st, sIdx) => !matchedStmtIdxs.has(sIdx));

  const instDiff = stmtInstallmentsBilledTotal - activeData.installmentsTotal;
  const instStatusBadge = Math.abs(instDiff) < 0.50 ? "✅ 100% Matched (Bank Statement and Active Bill in Sync)" : `⚠️ Discrepancy: ${instDiff.toFixed(2)} EGP`;

  sheet.getRange(curRow, 1, 1, 7).merge()
    .setValue("📦 Table 3: Installments Verification (Installments Sheet vs Bank Statement)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  sheet.getRange(curRow, 1, 1, 7).merge()
    .setValue(`Bank Statement Billed EMIs: ${stmtInstallmentsBilledTotal.toFixed(2)} EGP (${stmtInstallmentCount} items) | ` +
      `Active Sheet Due: ${activeData.installmentsTotal.toFixed(2)} EGP | ` +
      `Unbilled Sheet Installments (55-Day Policy / Pending): ${unbilledSheetInstallments.length} item(s) | Status: ${instStatusBadge}`)
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(Math.abs(instDiff) < 0.50 ? CONFIG.COLORS.PRIMARY : "#b71c1c");
  curRow++;

  const t3Headers = ["#", "Installment Description", "Payer", "Progress", "EMI (EGP)", "Status on Bank Statement", "Audit & Policy Verification Note"];
  sheet.getRange(curRow, 1, 1, t3Headers.length).setValues([t3Headers]).setFontWeight("bold").setBackground(CONFIG.COLORS.HEADER).setBorder(true, true, true, true, true, true);
  curRow++;

  const t3Rows = [];
  const t3RowColors = [];

  // 1. Matched installments (Billed by Bank and in Sheet)
  instMatches.forEach((m, idx) => {
    const sheetDesc = m.sheetRows.map(r => r.desc).join(" + ");
    const payer = m.payerLabel;
    const progress = (m.stmt.instCurrent && m.stmt.instTotal)
      ? `${m.stmt.instCurrent} of ${m.stmt.instTotal}`
      : (m.sheetRows[0].duration ? `Active (${m.sheetRows[0].duration} mos)` : "-");

    t3Rows.push([
      idx + 1,
      sheetDesc || m.stmt.desc,
      payer,
      progress,
      m.stmt.amount,
      "✅ Billed by Bank (Verified)",
      `Matched statement debit: "${m.stmt.desc}"`
    ]);
    t3RowColors.push(CONFIG.COLORS.PAID);
  });

  // 2. Unbilled Sheet Installments (In Sheet but NOT on Statement, e.g. Fan bought Aug 28)
  unbilledSheetInstallments.forEach(sh => {
    const pDate = sh.date;
    const firstDueDate = pDate ? getDueDateForInstallment(pDate, tz) : null;
    const isFutureGrace = firstDueDate && firstDueDate > activeStatementDueDate;

    let statusText = "";
    let noteText = "";
    let rowColor = "#fffde7"; // Soft yellow

    if (isFutureGrace) {
      const dueLabel = Utilities.formatDate(firstDueDate, tz, "MMMM yyyy");
      statusText = "⏳ NOT Billed (55-Day Grace Policy)";
      noteText = `Bought ${pDate ? Utilities.formatDate(pDate, tz, "MMM d, yyyy") : ""}. Under NBE 55-day policy, 1st EMI starts in ${dueLabel}. Excluded from this bill.`;
    } else {
      statusText = "⚠️ Missing from Bank Statement";
      noteText = "Recorded in Installments sheet, but NBE did NOT bill this EMI on this statement.";
      rowColor = CONFIG.COLORS.OVERDUE;
    }

    t3Rows.push([
      t3Rows.length + 1,
      sh.desc,
      sh.payer,
      sh.duration ? `Pending 1st of ${sh.duration}` : "Pending",
      sh.emi,
      statusText,
      noteText
    ]);
    t3RowColors.push(rowColor);
  });

  // 3. Statement installments not found in Sheet
  unrecordedStmtInstallments.forEach(st => {
    t3Rows.push([
      t3Rows.length + 1,
      st.desc,
      "⚠️ Unknown (Unrecorded)",
      (st.instCurrent && st.instTotal) ? `${st.instCurrent} of ${st.instTotal}` : "-",
      st.amount,
      "🚨 Unrecorded on Sheet",
      `Bank billed this EMI (${st.amount.toFixed(2)} EGP), but it is missing from Installments sheet!`
    ]);
    t3RowColors.push(CONFIG.COLORS.OVERDUE);
  });

  if (t3Rows.length === 0) {
    sheet.getRange(curRow, 1, 1, t3Headers.length).merge().setValue("No installments found.");
    curRow += 2;
  } else {
    sheet.getRange(curRow, 1, t3Rows.length, t3Headers.length).setValues(t3Rows);
    sheet.getRange(curRow, 5, t3Rows.length, 1).setNumberFormat("#,##0.00").setFontWeight("bold");

    for (let r = 0; r < t3Rows.length; r++) {
      sheet.getRange(curRow + r, 6).setBackground(t3RowColors[r]).setFontWeight("bold");
      if (t3Rows[r][5].includes("55-Day")) {
        sheet.getRange(curRow + r, 1, 1, t3Headers.length).setBackground("#fffde7");
      } else if (t3Rows[r][5].includes("Missing") || t3Rows[r][5].includes("Unrecorded")) {
        sheet.getRange(curRow + r, 1, 1, t3Headers.length).setBackground("#ffebee");
      }
    }
    curRow += t3Rows.length + 3;
  }

  // ==========================================
  // TABLE 4: ⚠️ STATEMENT DEBITS UNASSIGNED IN RECONCILIATION TABLE 2
  // ==========================================
  // Scan Bank Statement debits that have not been assigned in Table 2
  const unassignedDebits = [];
  if (stmtSheet && stmtSheet.getLastRow() >= 7) {
    const lastR = stmtSheet.getLastRow();
    const rows = stmtSheet.getRange(7, 1, lastR - 6, 7).getValues();
    rows.forEach((r, idx) => {
      const type = String(r[4] || "").trim().toUpperCase();
      const amt = parseFloat(r[5]);
      const desc = String(r[3] || "").trim();
      const rawDate = r[1];
      const pDate = parseDateValue(rawDate, tz);
      const status = String(r[6] || "").trim();

      if (type === "DEBIT" && !isNaN(amt) && amt > 0) {
        // Check if assigned in assignedChargesList
        const isAssigned = (assignedChargesList || []).some(a => {
          return Math.abs(a.amount - amt) < 0.05 &&
            (cleanMerchantName(a.desc).toLowerCase() === cleanMerchantName(desc).toLowerCase() || a.desc.includes(desc.substring(0, 15)));
        });

        // Check if matched to a purchase in Transactions
        const isMatched = debtLineItems.some(it => {
          return it.category === "Purchase" &&
            it.sortKey === activeSortKey &&
            Math.abs(it.originalAmount - amt) < 0.05 &&
            hasMerchantKeywordOverlap(desc, it.desc);
        });

        if (!isAssigned && !isMatched && !status.includes("Confirmed")) {
          unassignedDebits.push({
            idx: idx + 1,
            date: pDate || rawDate,
            desc: desc,
            amount: amt,
            status: status || "⚠️ Unassigned in Table 2"
          });
        }
      }
    });
  }

  sheet.getRange(curRow, 1, 1, 6).merge()
    .setValue("⚠️ Table 4: Statement Debits Unassigned in Reconciliation Table 2 (" + unassignedDebits.length + " items)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(unassignedDebits.length > 0 ? CONFIG.COLORS.OVERDUE : CONFIG.COLORS.PAID);
  curRow++;

  sheet.getRange(curRow, 1, 1, 6).merge()
    .setValue("Debits billed by NBE on the statement that are not in Transactions and have no payer assigned in Reconciliation Table 2.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);
  curRow++;

  const t4Headers = ["Statement #", "Date", "Description / Merchant", "Amount (EGP)", "Status", "Action Required"];
  sheet.getRange(curRow, 1, 1, t4Headers.length).setValues([t4Headers]).setFontWeight("bold").setBackground(CONFIG.COLORS.HEADER).setBorder(true, true, true, true, true, true);
  curRow++;

  if (unassignedDebits.length === 0) {
    sheet.getRange(curRow, 1, 1, 6).merge()
      .setValue("✅ All statement debits are assigned to family members or matched with transactions!")
      .setFontColor(CONFIG.COLORS.PRIMARY)
      .setFontWeight("bold");
    curRow += 2;
  } else {
    const t4Rows = unassignedDebits.map(d => [
      d.idx,
      d.date,
      d.desc,
      d.amount,
      d.status,
      "Select a payer in Reconciliation Table 2 to add to Debt Breakdown"
    ]);

    sheet.getRange(curRow, 1, t4Rows.length, t4Headers.length).setValues(t4Rows);
    sheet.getRange(curRow, 2, t4Rows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(curRow, 4, t4Rows.length, 1).setNumberFormat("#,##0.00").setFontWeight("bold");
    sheet.getRange(curRow, 1, t4Rows.length, t4Headers.length).setBackground(CONFIG.COLORS.OVERDUE);
    curRow += t4Rows.length + 3;
  }

  // ==========================================
  // TABLE 5: 📋 PURCHASES IN TRANSACTIONS NOT FOUND ON BANK STATEMENT
  // ==========================================
  const unbilledPurchases = [];
  const stmtPeriod = cardBal.stmtPeriod || getStatementPeriod(cardBal.statementDate, cardBal.minStmtTxDate, cardBal.maxStmtTxDate);
  (existingTxList || []).forEach(tx => {
    if (stmtPeriod.start && tx.date < stmtPeriod.start) return;
    if (stmtPeriod.end && tx.date > stmtPeriod.end) return;

    if (!cardBal.statementDate) {
      const cycleInfo = getBillingCycleForPurchase(tx.date, tz);
      if (!cycleInfo || cycleInfo.cycleSortKey !== activeSortKey) return;
    }

    // Check if matched any statement debit
    let matched = false;
    if (stmtSheet && stmtSheet.getLastRow() >= 7) {
      const lastR = stmtSheet.getLastRow();
      const rows = stmtSheet.getRange(7, 4, lastR - 6, 3).getValues(); // Col 4: desc, Col 5: type, Col 6: amt
      matched = rows.some(r => {
        const type = String(r[1] || "").trim().toUpperCase();
        const amt = parseFloat(r[2]);
        const desc = String(r[0] || "").trim();
        return type === "DEBIT" && Math.abs(amt - tx.amount) < 0.05 && hasMerchantKeywordOverlap(desc, tx.desc);
      });
    }

    if (!matched) {
      const dStr = Utilities.formatDate(tx.date, tz, "yyyy-MM-dd");
      const txKeyNorm = `${dStr}_${tx.amount.toFixed(2)}_${tx.desc.substring(0, 30)}_${tx.rawPerson}`;
      const isNeglected = (neglectedTxKeys && neglectedTxKeys.has(txKeyNorm));

      unbilledPurchases.push({
        date: tx.date,
        person: normalizePersonName(tx.rawPerson),
        desc: tx.desc,
        amount: tx.amount,
        status: isNeglected ? "🚫 Neglected / Excluded from Bill" : "Active in Bill (Unbilled by NBE)",
        action: isNeglected ? "Already excluded from bill" : "If bank has not billed this yet, check Neglect in Reconciliation Table 3"
      });
    }
  });

  sheet.getRange(curRow, 1, 1, 7).merge()
    .setValue("📋 Table 5: Purchases in Transactions NOT Found on Bank Statement (" + unbilledPurchases.length + " items)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground("#e3f2fd");
  curRow++;

  sheet.getRange(curRow, 1, 1, 7).merge()
    .setValue("Purchases recorded in Transactions sheet within this cycle that were not billed on this statement. Can be neglected via Reconciliation Table 3.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);
  curRow++;

  const t5Headers = ["#", "Purchase Date", "Person", "Description / Merchant", "Amount (EGP)", "Status", "Action Recommendation"];
  sheet.getRange(curRow, 1, 1, t5Headers.length).setValues([t5Headers]).setFontWeight("bold").setBackground(CONFIG.COLORS.HEADER).setBorder(true, true, true, true, true, true);
  curRow++;

  if (unbilledPurchases.length === 0) {
    sheet.getRange(curRow, 1, 1, 7).merge()
      .setValue("✅ All sheet transactions in this cycle were billed on the bank statement!")
      .setFontColor(CONFIG.COLORS.PRIMARY)
      .setFontWeight("bold");
    curRow += 2;
  } else {
    const t5Rows = unbilledPurchases.map((p, idx) => [
      idx + 1,
      p.date,
      p.person,
      p.desc,
      p.amount,
      p.status,
      p.action
    ]);

    sheet.getRange(curRow, 1, t5Rows.length, t5Headers.length).setValues(t5Rows);
    sheet.getRange(curRow, 2, t5Rows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(curRow, 5, t5Rows.length, 1).setNumberFormat("#,##0.00").setFontWeight("bold");
    for (let r = 0; r < t5Rows.length; r++) {
      if (t5Rows[r][5].includes("Neglected")) {
        sheet.getRange(curRow + r, 1, 1, t5Headers.length).setBackground("#eeeeee");
      }
    }
    curRow += t5Rows.length + 2;
  }

  sheet.autoResizeColumns(1, 8);
}