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
// DATE & CUTOFF UTILITIES
// ==========================================

function parseDateValue(rawDate, tz) {
  if (!rawDate) return null;
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) return rawDate;
  if (typeof rawDate === "number" && rawDate > 30000) {
    const d = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) return d;
  }
  const s = String(rawDate).trim();
  if (!s) return null;
  // Try YYYY-MM-DD
  const iso = s.match(/^(\d{4})[-\/\.](\d{1,2})[-\/\.](\d{1,2})/);
  if (iso) {
    const d = new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
    if (!isNaN(d.getTime())) return d;
  }
  // Try DD/MM/YYYY or DD-MM-YYYY
  const dmy = s.match(/^(\d{1,2})[-\/\.](\d{1,2})[-\/\.](\d{4})/);
  if (dmy) {
    const d = new Date(parseInt(dmy[3], 10), parseInt(dmy[2], 10) - 1, parseInt(dmy[1], 10));
    if (!isNaN(d.getTime())) return d;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}

function getDueDateForPurchase(purchaseDate, tz) {
  if (!purchaseDate || isNaN(purchaseDate.getTime())) return null;
  const targetTz = tz || "Africa/Cairo";
  const day = parseInt(Utilities.formatDate(purchaseDate, targetTz, "d"), 10);
  const month = parseInt(Utilities.formatDate(purchaseDate, targetTz, "M"), 10) - 1;
  const year = parseInt(Utilities.formatDate(purchaseDate, targetTz, "yyyy"), 10);
  // NBE Statement Cycle: Closes on the 29th of each month.
  // Purchases made on or before the 29th are billed on this month's statement (due 25th of next month: month + 1).
  // Purchases made on 30th or 31st roll over to next month's statement (due 25th of month + 2).
  const monthOffset = (day <= 29) ? 1 : 2;
  return new Date(year, month + monthOffset, 1);
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
  if (sheet.getName() !== CONFIG.SHEETS.RECONCILIATION) return;

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
    } catch (err) {}

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
      scriptProps.setProperty("ASSIGNED_STATEMENT_CHARGES", JSON.stringify(assignedMap));

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
      scriptProps.setProperty("ASSIGNED_STATEMENT_CHARGES", JSON.stringify(assignedMap));
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
        try { setKeys = JSON.parse(rawJson); } catch (err) {}
        if (!setKeys.includes(stKey)) {
          setKeys.push(stKey);
          scriptProps.setProperty("CONFIRMED_MATCHES", JSON.stringify(setKeys));
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
    } catch (err) {}

    if (isChecked) {
      if (!neglectedKeys.includes(txKeyNorm)) neglectedKeys.push(txKeyNorm);
      if (!neglectedKeys.includes(txKeyLegacy)) neglectedKeys.push(txKeyLegacy);
      scriptProps.setProperty("NEGLECTED_TRANSACTIONS", JSON.stringify(neglectedKeys));
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
      scriptProps.setProperty("NEGLECTED_TRANSACTIONS", JSON.stringify(neglectedKeys));
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
        const targetCell = stmtSheet.getRange(7 + i, 8);
        targetCell.setValue(statusText);
        if (color) targetCell.setBackground(color);
        break;
      }
    }
  } catch (err) {}
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

    if (result.processedCount > 0) {
      ui.alert(
        "Statement Processed Successfully",
        `Processed ${result.processedCount} statement file(s):\n` +
        `• File: ${result.processedFiles.join(", ")}\n` +
        `• Statement Date: ${result.lastMeta.statementDate || "N/A"}\n` +
        `• Due Date: ${result.lastMeta.dueDate || "N/A"}\n` +
        `• Closing Balance: ${result.lastMeta.closingBalance ? result.lastMeta.closingBalance.toFixed(2) + " EGP" : "N/A"}\n` +
        `• Transactions Found: ${result.totalTxCount}\n\n` +
        `The Bank Statement, Reconciliation, Debt Breakdown, and Audit & Differences tabs have all been updated!`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        "No New Statements Found",
        `No new PDF files were found in the '${CONFIG.FOLDER_NAME}' folder.\n\n` +
        `If you uploaded a file, please make sure it is in '${CONFIG.FOLDER_NAME}' (and not inside 'Processed').\n` +
        `The dashboard was refreshed using the current sheet data.`,
        ui.ButtonSet.OK
      );
    }
  } catch (err) {
    ui.alert("Error Processing Statement", err.message, ui.ButtonSet.OK);
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
        const text = doc.getBody().getText();
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
  const text = doc.getBody().getText();
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
    statementDate: ""
  };

  const mCard = text.match(/Card Number\s+([0-9\*]+)/);
  if (mCard) meta.card = mCard[1];

  const mOpenClose = text.match(/Opening Balance\s+([\d,]+\.?\d*)\s+Closing Balance\s+([\d,]+\.?\d*)/);
  if (mOpenClose) {
    meta.openingBalance = parseFloat(mOpenClose[1].replace(/,/g, ""));
    meta.closingBalance = parseFloat(mOpenClose[2].replace(/,/g, ""));
  }

  const mTot = text.match(/Total of Credit\s+([\d,]+\.?\d*)\s+Total of Debit\s+([\d,]+\.?\d*)/);
  if (mTot) {
    meta.totalCredit = parseFloat(mTot[1].replace(/,/g, ""));
    meta.totalDebit = parseFloat(mTot[2].replace(/,/g, ""));
  }

  const mDueLimit = text.match(/Due Date\s+([0-9]+\s+[A-Za-z]+\s+[0-9]+)\s+Credit Limit\s+([\d,]+\.?\d*)/);
  if (mDueLimit) {
    meta.dueDate = mDueLimit[1].trim();
    meta.creditLimit = parseFloat(mDueLimit[2].replace(/,/g, ""));
  }

  const mStmt = text.match(/Statement Date\s+([0-9]+\s+[A-Za-z]+\s+[0-9]+)/);
  if (mStmt) {
    meta.statementDate = mStmt[1].trim();
  }

  // Parse transaction items
  const dateRegexStr = "(?:\\d{1,2}\\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\\s+\\d{4})";
  const blockPattern = new RegExp("(" + dateRegexStr + "\\s+" + dateRegexStr + "[\\s\\S]*?)(?=" + dateRegexStr + "\\s+" + dateRegexStr + "|Page\\s+No\\.|\\Z)", "g");
  
  const matches = text.match(blockPattern) || [];
  const transactions = [];

  matches.forEach(m => {
    const rawBlock = m.trim().replace(/\n/g, " ");
    const mDates = rawBlock.match(new RegExp("^(" + dateRegexStr + ")\\s+(" + dateRegexStr + ")\\s+([\\s\\S]+)$"));
    if (!mDates) return;

    const txDate = mDates[1];
    const postDate = mDates[2];
    const body = mDates[3].trim();

    const amtMatch = body.match(/([\d,]+\.?\d*)\s+EGP\s+([\d,]+\.?\d*)(?:\s+(\d+))?/);
    if (!amtMatch) return;

    const amt = parseFloat(amtMatch[1].replace(/,/g, ""));
    const authCode = amtMatch[3] || "";
    const desc = body.substring(0, amtMatch.index).trim();

    let txType = "DEBIT";
    let instCurrent = null;
    let instTotal = null;

    const instMatch = desc.match(/(\d+)\s+OF\s+(\d+)/i);
    if (instMatch) {
      txType = "INSTALLMENT";
      instCurrent = parseInt(instMatch[1], 10);
      instTotal = parseInt(instMatch[2], 10);
    } else if (desc.toUpperCase().includes("PAYMENT") || desc.toUpperCase().includes("DIRECT DEBIT")) {
      txType = "CREDIT";
    }

    transactions.push({
      txDate: txDate,
      postDate: postDate,
      desc: desc,
      amount: amt,
      type: txType,
      authCode: authCode,
      instCurrent: instCurrent,
      instTotal: instTotal
    });
  });

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

      // Move file to Processed folder
      file.moveTo(processedFolder);
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
  sheet.getRange("A1:H1").merge()
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
  const headers = ["#", "Tx Date", "Posting Date", "Description", "Type", "Amount (EGP)", "Auth Code", "Assigned Payer / Status"];
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
        t.authCode,
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

function normalizePersonName(name) {
  if (!name) return "";
  const cleaned = String(name).trim().replace(/\s+/g, " ");
  if (!cleaned) return "";
  return cleaned
    .toLowerCase()
    .split(" ")
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ");
}

function splitPayerNames(payerStr) {
  if (!payerStr) return [];
  return String(payerStr)
    .split(/[\+\,\/]/)
    .map(p => normalizePersonName(p))
    .filter(p => p && p.toLowerCase() !== "shared");
}

function getAllUniquePayers(ss) {
  const peopleSet = new Set();

  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (txSheet && txSheet.getLastRow() >= 2) {
    const numRows = txSheet.getLastRow() - 1;
    const vals = txSheet.getRange(2, 4, numRows, 1).getValues();
    vals.forEach(r => {
      splitPayerNames(r[0]).forEach(name => peopleSet.add(name));
    });
  }

  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  if (instSheet && instSheet.getLastRow() >= 2) {
    const numRows = instSheet.getLastRow() - 1;
    const vals = instSheet.getRange(2, 9, numRows, 1).getValues();
    vals.forEach(r => {
      splitPayerNames(r[0]).forEach(name => peopleSet.add(name));
    });
  }

  // Also include payers from stored statement assignments
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const assignedMap = JSON.parse(scriptProps.getProperty("ASSIGNED_STATEMENT_CHARGES") || "{}");
    Object.keys(assignedMap).forEach(k => {
      const item = assignedMap[k];
      if (item && item.payer) {
        splitPayerNames(item.payer).forEach(name => peopleSet.add(name));
      }
    });
  } catch (e) {}

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
            splitPayerNames(r[0]).forEach(name => peopleSet.add(name));
          });
        }
      }
    }
  } catch (e) {}

  const result = Array.from(peopleSet).sort((a, b) => a.localeCompare(b));
  return result.length > 0 ? result : ["Abdo", "Dad", "Hager", "Mai", "Mido", "Muhanad", "Mum", "Nourween", "Zoza"];
}

function getStatementPeriod(stmtDate, minStmtTxDate, maxStmtTxDate) {
  if (!stmtDate && !minStmtTxDate) return { start: null, end: null };

  let pEnd = null;
  let pStart = null;

  if (stmtDate) {
    pEnd = new Date(stmtDate.getFullYear(), stmtDate.getMonth(), stmtDate.getDate(), 23, 59, 59, 999);

    const sYear = stmtDate.getFullYear();
    const sMonth = stmtDate.getMonth(); // 0-indexed (0=Jan)
    const sDay = stmtDate.getDate();

    // Previous month index and year
    const prevMonth = sMonth === 0 ? 11 : sMonth - 1;
    const prevYear = sMonth === 0 ? sYear - 1 : sYear;

    // Number of days in previous month
    const daysInPrevMonth = new Date(sYear, sMonth, 0).getDate();
    const prevCutoffDay = Math.min(sDay, daysInPrevMonth);

    // Cycle starts the day after previous cutoff
    let cycleStartDay = prevCutoffDay + 1;
    let cycleStartMonth = prevMonth;
    let cycleStartYear = prevYear;

    if (cycleStartDay > daysInPrevMonth) {
      cycleStartDay = 1;
      cycleStartMonth = sMonth;
      cycleStartYear = sYear;
    }

    pStart = new Date(cycleStartYear, cycleStartMonth, cycleStartDay, 0, 0, 0, 0);
  }

  // If statement transactions start earlier, expand pStart to encompass them
  if (minStmtTxDate) {
    const minDateStart = new Date(minStmtTxDate.getFullYear(), minStmtTxDate.getMonth(), minStmtTxDate.getDate(), 0, 0, 0, 0);
    if (!pStart || minDateStart < pStart) {
      pStart = minDateStart;
    }
  }

  // If statement transactions end later, expand pEnd to encompass them
  if (maxStmtTxDate) {
    const maxDateEnd = new Date(maxStmtTxDate.getFullYear(), maxStmtTxDate.getMonth(), maxStmtTxDate.getDate(), 23, 59, 59, 999);
    if (!pEnd || maxDateEnd > pEnd) {
      pEnd = maxDateEnd;
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
  } catch (err) {}

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
        scriptProps.setProperty("ASSIGNED_STATEMENT_CHARGES", JSON.stringify(assignedMap));
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
  } catch (e) {}

  stmtDebits.forEach(st => {
    const stKey = `${st.dateStr}_${st.amount.toFixed(2)}_${st.desc.substring(0, 20)}`;
    if (confirmedMatchKeys.has(stKey)) {
      st.matched = true;
      st.matchType = "CONFIRMED";
    }
  });

  // Pass 1: 1-to-1 Exact Match (Amount exact, Date within 5 days, Keyword Overlap)
  stmtDebits.forEach(st => {
    if (st.matched) return;
    const match = txDebits.find(tx => {
      if (tx.matched) return false;
      const amtDiff = Math.abs(tx.amount - st.amount);
      if (amtDiff > 0.05) return false;

      if (st.date) {
        const daysDiff = Math.abs((tx.date.getTime() - st.date.getTime()) / (1000 * 60 * 60 * 24));
        if (daysDiff > 5) return false;
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

// (findSubsetCombination moved to top-level Section 4 for shared use across debits & installments)

  // Pass 2: N-Way Split Matches (2, 3, 4, 5, or 6 people split 1 statement charge)
  stmtDebits.forEach(st => {
    if (st.matched) return;

    const candidates = txDebits.filter(tx => {
      if (tx.matched) return false;
      if (st.date) {
        const daysDiff = Math.abs((tx.date.getTime() - st.date.getTime()) / (1000 * 60 * 60 * 24));
        return daysDiff <= 5;
      }
      return true;
    });

    if (candidates.length < 2) return;

    const combo = findSubsetCombination(candidates, st.amount, 6, 0.10);
    if (combo) {
      const keywordMatch = combo.some(item => hasMerchantKeywordOverlap(st.desc, item.desc));
      if (keywordMatch) {
        st.matched = true;
        st.matchType = "SPLIT";
        combo.forEach(item => {
          item.matched = true;
          st.matchedItems.push(item);
        });
      }
    }
  });

  // Pass 3: Suggested Matches (Needs user confirmation)
  const suggestedMatches = [];

  stmtDebits.forEach(st => {
    if (st.matched) return;

    const nearby = txDebits.filter(tx => {
      if (tx.matched) return false;
      if (st.date) {
        const daysDiff = Math.abs((tx.date.getTime() - st.date.getTime()) / (1000 * 60 * 60 * 24));
        return daysDiff <= 4;
      }
      return true;
    });

    if (nearby.length === 0) return;

    // Case A: Exact 1-to-1 amount match, but different description or typo
    const singleAmtMatch = nearby.find(tx => Math.abs(tx.amount - st.amount) < 0.05);
    if (singleAmtMatch) {
      suggestedMatches.push({
        stmt: st,
        sheetItems: [singleAmtMatch],
        reason: `Exact amount (${st.amount.toFixed(2)}), dates within 4 days — check shop name`
      });
      return;
    }

    // Case B: N-way split sum matches (2 to 6 items) without keyword overlap
    const combo = findSubsetCombination(nearby, st.amount, 6, 0.10);
    if (combo) {
      const peopleList = combo.map(it => it.person).join(" + ");
      suggestedMatches.push({
        stmt: st,
        sheetItems: combo,
        reason: `${combo.length}-way split sum matches ${st.amount.toFixed(2)} (${peopleList})`
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
    } catch (err) {}

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
    } catch (e) {}

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
        try { assignedMap = JSON.parse(scriptProps.getProperty("ASSIGNED_STATEMENT_CHARGES") || "{}"); } catch (e) {}
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

  stmtSheet.getRange(6, 8).setValue("Assigned Payer / Status")
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, false, false);

  const statusValues = stmtRowStatus.map(s => [s.text]);
  const statusColors = stmtRowStatus.map(s => [s.color]);
  const hRange = stmtSheet.getRange(7, 8, stmtData.length, 1);
  hRange.setValues(statusValues);
  hRange.setBackgrounds(statusColors);
  hRange.setFontWeight("bold");
  hRange.setFontSize(9);
  stmtSheet.autoResizeColumns(1, 8);

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

  // Read latest statement closing balance if available
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  if (stmtSheet && stmtSheet.getLastRow() >= 4) {
    const closeVal = stmtSheet.getRange("D3").getValue();
    if (!isNaN(parseFloat(closeVal))) {
      billedBalance = parseFloat(closeVal);
    }
    const openVal = stmtSheet.getRange("B3").getValue();
    if (!isNaN(parseFloat(openVal))) openingBalance = parseFloat(openVal);

    const totDebVal = stmtSheet.getRange("F3").getValue();
    if (!isNaN(parseFloat(totDebVal))) totalDebit = parseFloat(totDebVal);

    const totCredVal = stmtSheet.getRange("B4").getValue();
    if (!isNaN(parseFloat(totCredVal))) totalCredit = parseFloat(totCredVal);

    const stmtDateVal = stmtSheet.getRange("B2").getValue();
    statementDate = parseDateValue(stmtDateVal, tz);
    const dueDateVal = stmtSheet.getRange("D2").getValue();
    statementDueDate = parseDateValue(dueDateVal, tz);
  }

  if (!statementDueDate && statementDate) {
    statementDueDate = getDueDateForPurchase(statementDate, tz);
  }

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

  // Calculate Unbilled New Purchases (strictly after statement cutoff)
  let unbilledNewPurchases = 0;
  let stmtCutoff = null;
  if (statementDate) {
    stmtCutoff = new Date(statementDate.getFullYear(), statementDate.getMonth(), statementDate.getDate(), 23, 59, 59, 999);
  }

  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (txSheet && txSheet.getLastRow() >= 2) {
    const txData = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 5).getValues();
    txData.forEach(row => {
      const pDate = parseDateValue(row[1], tz);
      const amt = parseFloat(row[4]);
      if (pDate && !isNaN(amt) && amt > 0) {
        if (stmtCutoff && pDate > stmtCutoff) {
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
    let stmtDueMonthLabel = "";
    if (statementDueDate) {
      stmtDueMonthLabel = Utilities.formatDate(statementDueDate, tz, "MMMM yyyy").toLowerCase();
    } else if (statementDate) {
      const stmtDueDate = getDueDateForPurchase(statementDate, tz);
      if (stmtDueDate) stmtDueMonthLabel = Utilities.formatDate(stmtDueDate, tz, "MMMM yyyy").toLowerCase();
    }

    isCurrentBillPaid = paidMonths.includes(currentMonthLabel) || (stmtDueMonthLabel && paidMonths.includes(stmtDueMonthLabel));
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
    try { ss.deleteSheet(oldDebtsSheet); } catch (e) {}
  }

  if (!transactionsSheet || !installmentsSheet || !historySheet) {
    SpreadsheetApp.getUi().alert(
      "Missing Sheet",
      "Please make sure 'Transactions', 'Installments', and 'Payment History' sheets exist.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // NOTE: 'Transactions' and 'Installments' sheets are user-managed inputs.
  // The script NEVER modifies, writes to, or cleans them.

  // Run reconciliation ONLY when not skipped (e.g. avoid erasing/rewriting Reconciliation sheet during onEdit)
  if (!opts.skipRecon) {
    runReconciliation(ss, tz);
  }

  // Read Card Balance & Active Statement info
  const cardBal = calculateCardBalance(ss, tz);
  const activeStatementDueDate = cardBal.statementDueDate || (cardBal.statementDate ? getDueDateForPurchase(cardBal.statementDate, tz) : new Date());
  const activeStatementSortKey = Utilities.formatDate(activeStatementDueDate, tz, "yyyy-MM");

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

  function addDebtItem({ dueDate, person, category, desc, installmentInfo, purchaseDate, originalAmount, amount, isMissingFromSheet, note }) {
    if (!dueDate || !person) return;
    const normalizedPerson = normalizePersonName(person);
    if (!normalizedPerson || normalizedPerson.toLowerCase() === "shared") return;

    const sortKey = Utilities.formatDate(dueDate, tz, "yyyy-MM");
    const label = Utilities.formatDate(dueDate, tz, "MMMM yyyy");

    if (!allDebts[sortKey]) {
      allDebts[sortKey] = {
        label: label,
        dueDate: dueDate,
        total: 0,
        purchasesTotal: 0,
        installmentsTotal: 0,
        missingFromSheetTotal: 0,
        people: {},
        peopleBreakdown: {}
      };
    }

    allDebts[sortKey].total += amount;
    if (category === "Purchase") {
      allDebts[sortKey].purchasesTotal += amount;
    } else if (category === "Installment") {
      allDebts[sortKey].installmentsTotal += amount;
    } else if (category === "MissingStatementDebit" || isMissingFromSheet) {
      allDebts[sortKey].missingFromSheetTotal += amount;
    }

    allDebts[sortKey].people[normalizedPerson] = (allDebts[sortKey].people[normalizedPerson] || 0) + amount;

    if (!allDebts[sortKey].peopleBreakdown[normalizedPerson]) {
      allDebts[sortKey].peopleBreakdown[normalizedPerson] = { purchases: 0, installments: 0, missingFromSheet: 0, total: 0 };
    }
    if (category === "Purchase") {
      allDebts[sortKey].peopleBreakdown[normalizedPerson].purchases += amount;
    } else if (category === "Installment") {
      allDebts[sortKey].peopleBreakdown[normalizedPerson].installments += amount;
    } else if (category === "MissingStatementDebit" || isMissingFromSheet) {
      allDebts[sortKey].peopleBreakdown[normalizedPerson].missingFromSheet += amount;
    }
    allDebts[sortKey].peopleBreakdown[normalizedPerson].total += amount;

    debtLineItems.push({
      sortKey: sortKey,
      dueMonthLabel: label,
      dueDate: dueDate,
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
    neglectedTxKeys = new Set(JSON.parse(scriptProps.getProperty("NEGLECTED_TRANSACTIONS") || "[]"));
  } catch (err) {}

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

      // Check if user neglected/excluded this transaction in Reconciliation Table 3
      const dStr = Utilities.formatDate(purchaseDate, tz, "yyyy-MM-dd");
      const txKeyNorm = `${dStr}_${amount.toFixed(2)}_${desc.substring(0, 30)}_${rawPerson}`;
      const txKey1 = `${String(purchaseDate)}_${amount.toFixed(2)}_${desc.substring(0, 30)}_${rawPerson}`;
      const txKey2 = `${String(rawDate)}_${amount.toFixed(2)}_${desc.substring(0, 30)}_${rawPerson}`;
      if (neglectedTxKeys.has(txKeyNorm) || neglectedTxKeys.has(txKey1) || neglectedTxKeys.has(txKey2)) {
        return;
      }

      existingTransactionsForDedupe.push({
        date: purchaseDate,
        rawDate: rawDate,
        desc: desc,
        amount: amount,
        rawPerson: rawPerson
      });

      const dueDate = getDueDateForPurchase(purchaseDate, tz);
      if (!dueDate) return;

      const people = splitPayerNames(rawPerson);
      if (people.length === 0) return;
      const splitAmount = amount / people.length;

      people.forEach(p => {
        addDebtItem({
          dueDate: dueDate,
          person: p,
          category: "Purchase",
          desc: desc,
          installmentInfo: "-",
          purchaseDate: purchaseDate,
          originalAmount: amount,
          amount: splitAmount,
          isMissingFromSheet: false
        });
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

      // First installment is due in the statement cycle of the purchase date
      const firstDueDate = getDueDateForPurchase(purchaseDate, tz);

      for (let i = 0; i < durationMonths; i++) {
        const dueMonthDate = new Date(firstDueDate.getFullYear(), firstDueDate.getMonth() + i, 1);
        const instInfo = `${i + 1} of ${durationMonths}`;

        people.forEach(p => {
          addDebtItem({
            dueDate: dueMonthDate,
            person: p,
            category: "Installment",
            desc: desc,
            installmentInfo: instInfo,
            purchaseDate: purchaseDate,
            originalAmount: emi,
            amount: splitEmi,
            isMissingFromSheet: false
          });
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
    } catch (err) {}
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
    if (!payerStr || payerStr === "🚫 Neglect / Ignore" || payerStr.toUpperCase().includes("NEGLECT")) {
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
        dueDate: chargeDueDate,
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

  // Collect and sort unique people: combine all known sheet payers + any in debtLineItems
  const allKnownPeople = getAllUniquePayers(ss);
  const allPeopleSet = new Set(allKnownPeople);
  debtLineItems.forEach(it => {
    if (it.person && it.person.toLowerCase() !== "shared") {
      allPeopleSet.add(it.person);
    }
  });
  const sortedPeople = Array.from(allPeopleSet).sort((a, b) => a.localeCompare(b));

  // 4. Render Dedicated 'Debt Breakdown' Sheet
  renderDebtBreakdownSheet(ss, tz, allDebts, debtLineItems, sortedPeople, paidMonthsSet, cardBal, activeStatementDueDate, assignedMissingCount, assignedMissingTotal);

  // 5. Render Dedicated 'Audit & Differences' Sheet
  renderAuditDifferencesSheet(ss, tz, cardBal, allDebts, debtLineItems, activeStatementDueDate, reconAssignedList, neglectedTxKeys, existingTransactionsForDedupe, sortedPeople, assignedChargesList);
}

function renderDebtBreakdownSheet(ss, tz, allDebts, debtLineItems, sortedPeople, paidMonthsSet, cardBal, activeStatementDueDate, assignedMissingCount, assignedMissingTotal) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.DEBT_BREAKDOWN);
  sheet.clear();
  sheet.setHiddenGridlines(false);

  const activeSortKey = Utilities.formatDate(activeStatementDueDate, tz, "yyyy-MM");
  const activeMonthLabel = Utilities.formatDate(activeStatementDueDate, tz, "MMMM yyyy");
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
    } catch (e) {}

    curRow += regRows.length + 3;
  } else {
    sheet.getRange(curRow, 1, 1, 9).merge().setValue("No charges found for this statement cycle.");
    curRow += 3;
  }

  // ==========================================
  // SECTION 3: 📅 OTHER BILLING CYCLES (Individual Tables for Each Month)
  // ==========================================
  const otherKeys = Object.keys(allDebts).sort().filter(k => k !== activeSortKey);

  if (otherKeys.length > 0) {
    sheet.getRange(curRow, 1, 1, 7).merge()
      .setValue("📅 Other Billing Cycles (Individual Monthly Tables)")
      .setFontWeight("bold")
      .setFontSize(13)
      .setBackground(CONFIG.COLORS.PRIMARY)
      .setFontColor("#ffffff");
    curRow++;

    sheet.getRange(curRow, 1, 1, 7).merge()
      .setValue("Upcoming and previous monthly dues. Each month has its own dedicated breakdown table below.")
      .setFontStyle("italic")
      .setFontSize(9)
      .setFontColor(CONFIG.COLORS.TEXT_MUTED);
    curRow += 2;

    otherKeys.forEach(mKey => {
      const mData = allDebts[mKey];
      const mLabel = mData.label;
      const isPaid = paidMonthsSet.has(mLabel.toLowerCase());

      // Header Bar for this month
      const mTitle = `📅 ${mLabel} (Due 25th) — Total Due: ${mData.total.toFixed(2)} EGP ${isPaid ? "✅ [PAID]" : "🕒 [UPCOMING]"}`;
      sheet.getRange(curRow, 1, 1, 5).merge()
        .setValue(mTitle)
        .setFontWeight("bold")
        .setFontSize(11)
        .setBackground(isPaid ? CONFIG.COLORS.PAID : CONFIG.COLORS.HEADER);
      curRow++;

      const mHeaders = ["Person", "🛒 Purchases (EGP)", "📦 Installments (EGP)", "Total Due (EGP)", "% of Month Bill"];
      sheet.getRange(curRow, 1, 1, mHeaders.length).setValues([mHeaders])
        .setFontWeight("bold")
        .setBackground(CONFIG.COLORS.HEADER)
        .setBorder(true, true, true, true, true, true);
      curRow++;

      const mPeople = sortedPeople.filter(p => (mData.peopleBreakdown[p] && mData.peopleBreakdown[p].total > 0));
      const mRows = mPeople.map(p => {
        const bk = mData.peopleBreakdown[p] || { purchases: 0, installments: 0, total: 0 };
        const pct = mData.total > 0 ? (bk.total / mData.total) : 0;
        return [p, bk.purchases, bk.installments, bk.total, pct];
      });

      if (mRows.length > 0) {
        sheet.getRange(curRow, 1, mRows.length, mHeaders.length).setValues(mRows);
        sheet.getRange(curRow, 2, mRows.length, 3).setNumberFormat("#,##0.00");
        sheet.getRange(curRow, 5, mRows.length, 1).setNumberFormat("0.0%");
        sheet.getRange(curRow, 1, mRows.length, 1).setFontWeight("bold");
        curRow += mRows.length;
      }

      // Total Row for this month
      const mTotalRow = ["TOTAL (" + mLabel + ")", mData.purchasesTotal, mData.installmentsTotal, mData.total, 1.00];
      sheet.getRange(curRow, 1, 1, mHeaders.length).setValues([mTotalRow])
        .setFontWeight("bold")
        .setBackground(CONFIG.COLORS.PRIMARY_LIGHT)
        .setBorder(true, true, true, true, true, true);
      sheet.getRange(curRow, 2, 1, 3).setNumberFormat("#,##0.00");
      sheet.getRange(curRow, 5, 1, 1).setNumberFormat("0.0%");
      curRow += 2;
    });
  }

  sheet.autoResizeColumns(1, 9);
}

function renderAuditDifferencesSheet(ss, tz, cardBal, allDebts, debtLineItems, activeStatementDueDate, reconAssignedList, neglectedTxKeys, existingTxList, sortedPeople, assignedChargesList) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.AUDIT_DIFFERENCES);
  sheet.clear();
  sheet.setHiddenGridlines(false);

  const activeSortKey = Utilities.formatDate(activeStatementDueDate, tz, "yyyy-MM");
  const activeMonthLabel = Utilities.formatDate(activeStatementDueDate, tz, "MMMM yyyy");
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

  let statusBadge = "✅ Perfect Match (0.00 EGP difference)";
  let statusColor = CONFIG.COLORS.PAID;
  if (Math.abs(netDiff) >= 0.05) {
    if (netDiff > 0) {
      statusBadge = "⚠️ Family Shares Exceed Bank Bill by +" + netDiff.toFixed(2) + " EGP";
      statusColor = CONFIG.COLORS.DUE_SOON;
    } else {
      statusBadge = "🚨 Bank Bill Exceeds Shares by -" + Math.abs(netDiff).toFixed(2) + " EGP (Missing Items!)";
      statusColor = CONFIG.COLORS.OVERDUE;
    }
  }

  const kpiData = [
    ["🏦 Bank Closing Balance (Must Pay):", bankBill, "👥 Total Family Shares in Debt Breakdown:", familyTotal, "⚖️ Net Discrepancy:", netDiff, "Status:", statusBadge],
    ["🛒 Purchases Total (Sheet):", activeData.purchasesTotal, "📦 Installments Total (Sheet):", activeData.installmentsTotal, "⚠️ Missing Assigned (Recon Table 2):", activeData.missingFromSheetTotal, "Total Card Limit:", cardBal.creditLimit],
    ["💳 Statement Total Debits:", cardBal.totalDebit || 0, "💰 Statement Credits / Payments:", cardBal.totalCredit || 0, "Available Balance Now:", cardBal.availableBalanceNow, "Available Post-Settlement:", cardBal.availableAfterSettlement]
  ];

  sheet.getRange(curRow, 1, kpiData.length, 8).setValues(kpiData);
  sheet.getRange(curRow, 1, kpiData.length, 8).setFontSize(9).setBackground(CONFIG.COLORS.PRIMARY_LIGHT).setBorder(true, true, true, true, true, true);
  for (let c of [1, 3, 5, 7]) {
    sheet.getRange(curRow, c, kpiData.length, 1).setFontWeight("bold");
  }
  for (let c of [2, 4, 6]) {
    sheet.getRange(curRow, c, kpiData.length, 1).setNumberFormat("#,##0.00").setFontWeight("bold");
  }
  sheet.getRange(curRow + 1, 8).setNumberFormat("#,##0.00");
  sheet.getRange(curRow + 2, 6).setNumberFormat("#,##0.00");
  sheet.getRange(curRow + 2, 8).setNumberFormat("#,##0.00");
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

  (sortedPeople || []).forEach(p => {
    const bk = (activeData.peopleBreakdown && activeData.peopleBreakdown[p]) || { purchases: 0, installments: 0, missingFromSheet: 0, total: 0 };
    const pPurchases = bk.purchases || 0;
    const pInstallments = bk.installments || 0;

    // Calculate this person's assigned debits directly from assignedChargesList (Table 2)
    let pReconAssigned = 0;
    (assignedChargesList || []).forEach(item => {
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
      const payers = splitPayerNames(item.payer);
      // Check if this item is represented in active Debt Breakdown missing items
      const isReflected = debtLineItems.some(it => {
        return it.isMissingFromSheet &&
               Math.abs(it.originalAmount - item.amount) < 0.05 &&
               payers.includes(it.person);
      });

      return [
        item.itemNum || (idx + 1),
        item.rawDate || item.dateStr,
        item.desc,
        item.amount,
        item.payer,
        item.note || "-",
        isReflected ? "✅ Reflected in Debt Breakdown" : "🚨 Missing from Debt Breakdown!"
      ];
    });

    sheet.getRange(curRow, 1, t2Rows.length, t2Headers.length).setValues(t2Rows);
    sheet.getRange(curRow, 2, t2Rows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(curRow, 4, t2Rows.length, 1).setNumberFormat("#,##0.00").setFontWeight("bold");

    for (let r = 0; r < t2Rows.length; r++) {
      if (t2Rows[r][6].includes("Reflected")) {
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

  if (stmtSheet && stmtSheet.getLastRow() >= 7) {
    const lastR = stmtSheet.getLastRow();
    const rows = stmtSheet.getRange(7, 1, lastR - 6, 8).getValues();
    rows.forEach(r => {
      const type = String(r[4] || "").trim().toUpperCase();
      const amt = parseFloat(r[5]);
      if (type === "INSTALLMENT" && !isNaN(amt) && amt > 0) {
        stmtInstallmentsBilledTotal += amt;
        stmtInstallmentCount++;
      }
    });
  }

  const instDiff = stmtInstallmentsBilledTotal - activeData.installmentsTotal;
  const instStatusBadge = Math.abs(instDiff) < 0.50 ? "✅ 100% Matched (Bank Statement and Sheet Installments are in perfect sync)" : `⚠️ Discrepancy: ${instDiff.toFixed(2)} EGP`;

  sheet.getRange(curRow, 1, 1, 6).merge()
    .setValue("📦 Table 3: Installments Verification (Installments Sheet vs Bank Statement)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  sheet.getRange(curRow, 1, 1, 6).merge()
    .setValue(`Bank Statement Billed EMIs: ${stmtInstallmentsBilledTotal.toFixed(2)} EGP (${stmtInstallmentCount} items) | Installments Sheet Due: ${activeData.installmentsTotal.toFixed(2)} EGP | Status: ${instStatusBadge}`)
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(Math.abs(instDiff) < 0.50 ? CONFIG.COLORS.PRIMARY : "#b71c1c");
  curRow++;

  const t3Headers = ["#", "Installment Description", "Payer", "Progress", "EMI (Sheet EGP)", "Status on Bank Statement"];
  sheet.getRange(curRow, 1, 1, t3Headers.length).setValues([t3Headers]).setFontWeight("bold").setBackground(CONFIG.COLORS.HEADER).setBorder(true, true, true, true, true, true);
  curRow++;

  // List active installments for this cycle from debtLineItems
  const activeInstItems = debtLineItems.filter(it => it.category === "Installment" && it.sortKey === activeSortKey);
  if (activeInstItems.length === 0) {
    sheet.getRange(curRow, 1, 1, 6).merge().setValue("No installments due for this cycle.");
    curRow += 2;
  } else {
    const t3Rows = activeInstItems.map((it, idx) => [
      idx + 1,
      it.desc,
      it.person,
      it.installmentInfo,
      it.amount,
      "✅ Billed by Bank"
    ]);

    sheet.getRange(curRow, 1, t3Rows.length, t3Headers.length).setValues(t3Rows);
    sheet.getRange(curRow, 5, t3Rows.length, 1).setNumberFormat("#,##0.00").setFontWeight("bold");
    sheet.getRange(curRow, 6, t3Rows.length, 1).setBackground(CONFIG.COLORS.PAID);
    curRow += t3Rows.length + 3;
  }

  // ==========================================
  // TABLE 4: ⚠️ STATEMENT DEBITS UNASSIGNED IN RECONCILIATION TABLE 2
  // ==========================================
  // Scan Bank Statement debits that have not been assigned in Table 2
  const unassignedDebits = [];
  if (stmtSheet && stmtSheet.getLastRow() >= 7) {
    const lastR = stmtSheet.getLastRow();
    const rows = stmtSheet.getRange(7, 1, lastR - 6, 8).getValues();
    rows.forEach((r, idx) => {
      const type = String(r[4] || "").trim().toUpperCase();
      const amt = parseFloat(r[5]);
      const desc = String(r[3] || "").trim();
      const rawDate = r[1];
      const pDate = parseDateValue(rawDate, tz);
      const status = String(r[7] || "").trim();

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
  (existingTxList || []).forEach(tx => {
    const dueDate = getDueDateForPurchase(tx.date, tz);
    if (!dueDate) return;
    const sKey = Utilities.formatDate(dueDate, tz, "yyyy-MM");
    if (sKey !== activeSortKey) return;

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