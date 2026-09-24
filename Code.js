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
    MONTHLY_DEBTS: "Monthly Debts",
    DEBT_BREAKDOWN: "Debt Breakdown",
    PAYMENT_HISTORY: "Payment History",
    BANK_STATEMENT: "Bank Statement",
    RECONCILIATION: "Reconciliation"
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

function cleanTransactionsSheet(ss) {
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!txSheet || txSheet.getLastRow() < 2) return 0;

  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const lastRow = txSheet.getLastRow();
  const data = txSheet.getRange(2, 1, lastRow - 1, 5).getValues();

  const seenExact = new Set();
  const cleanRows = [];
  let duplicatesRemoved = 0;

  data.forEach(r => {
    const rawDate = r[1];
    const d = parseDateValue(rawDate, tz);
    const dStr = d ? Utilities.formatDate(d, tz, "yyyy-MM-dd") : String(rawDate).trim();
    const desc = String(r[2] || "").trim().toLowerCase();
    const payer = String(r[3] || "").trim().toLowerCase();
    const amt = parseFloat(r[4]);

    if (!desc || !payer || isNaN(amt) || amt <= 0) return;

    const exactKey = `${dStr}_${desc}_${payer}_${amt.toFixed(2)}`;

    // Exact duplicate seen earlier in the sheet
    if (seenExact.has(exactKey)) {
      duplicatesRemoved++;
      return;
    }

    seenExact.add(exactKey);
    cleanRows.push(r);
  });

  if (duplicatesRemoved > 0) {
    // Renumber IDs sequentially
    cleanRows.forEach((r, idx) => {
      r[0] = idx + 1;
    });

    txSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    if (cleanRows.length > 0) {
      txSheet.getRange(2, 1, cleanRows.length, 5).setValues(cleanRows);
      txSheet.getRange(2, 2, cleanRows.length, 1).setNumberFormat("dddd, MMMM d, yyyy");
      txSheet.getRange(2, 5, cleanRows.length, 1).setNumberFormat("#,##0.00");
    }
  }

  return duplicatesRemoved;
}

function menuCleanDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const removed = cleanTransactionsSheet(ss);
  updateLiveDashboard();
  SpreadsheetApp.getUi().alert(
    "Cleanup Complete",
    `Removed ${removed} duplicate transaction(s)!\n\n` +
    `Monthly Debts and Reconciliation have been restored to their authentic numbers.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💳 NBE Tracker")
    .addItem("📥 Process Latest Statement (Drive)", "menuProcessStatement")
    .addItem("🔄 Refresh Dashboard & Reconciliation", "updateLiveDashboard")
    .addSeparator()
    .addItem("📊 Go to Debt Breakdown Sheet", "menuGoToDebtBreakdown")
    .addItem("🧹 Clean Duplicates & Restore Debts", "menuCleanDuplicates")
    .addItem("➕ Add Assigned Charges to Transactions", "addAssignedChargesToTransactions")
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

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEETS.RECONCILIATION) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row < 5) return;

  const validPayers = ["Mido", "Mai", "Abdo", "Dad", "Mum", "Zoza", "Shared"];

  // Table 2: User assigns a Payer in Column E (col 5)
  if (col === 5) {
    const val = String(e.value || e.range.getValue() || "").trim();
    if (!validPayers.includes(val)) return;

    // Check that row is strictly in Table 2 (under "MISSING from Sheet")
    const lastRow = sheet.getLastRow();
    const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
    let table2Start = -1;
    for (let i = 0; i < colA.length; i++) {
      if (String(colA[i][0]).includes("MISSING from Sheet")) {
        table2Start = i + 1;
        break;
      }
    }
    if (table2Start === -1 || row <= table2Start + 1) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
    if (!txSheet) return;

    // Read row values: Col 1 (#), Col 2 (Date), Col 3 (Desc), Col 4 (Amount), Col 5 (Payer), Col 6 (Custom Note)
    const rowValues = sheet.getRange(row, 1, 1, 6).getValues()[0];
    const rawDate = rowValues[1];
    const origDesc = String(rowValues[2] || "").trim();
    const amt = parseFloat(rowValues[3]);
    const payer = val;
    const note = String(rowValues[5] || "").trim();

    if (isNaN(amt) || amt <= 0) return;

    const dateObj = new Date(rawDate);
    const finalDate = isNaN(dateObj.getTime()) ? new Date() : dateObj;
    const finalDesc = note || origDesc;

    // Check if duplicate already exists in Transactions
    const txLastRow = txSheet.getLastRow();
    if (txLastRow >= 2) {
      const existingTx = txSheet.getRange(2, 1, txLastRow - 1, 5).getValues();
      const alreadyExists = existingTx.some(r => {
        const rAmt = parseFloat(r[4]);
        const rDesc = String(r[2] || "").trim().toLowerCase();
        return Math.abs(rAmt - amt) < 0.05 && rDesc === finalDesc.toLowerCase();
      });
      if (alreadyExists) {
        sheet.getRange(row, 1, 1, 6).setBackground(CONFIG.COLORS.PAID);
        sheet.getRange(row, 5).clearDataValidations().setValue("✅ Already in Sheet (" + payer + ")");
        ss.toast(`Charge "${finalDesc}" is already in Transactions!`, "ℹ️ Already Recorded", 3);
        return;
      }
    }

    // Determine next ID in Transactions
    let nextId = 1;
    if (txLastRow >= 2) {
      const lastIdVal = txSheet.getRange(txLastRow, 1).getValue();
      if (!isNaN(parseInt(lastIdVal, 10))) {
        nextId = parseInt(lastIdVal, 10) + 1;
      }
    }

    // Append to Transactions
    const newTxRow = [nextId, finalDate, finalDesc, payer, amt];
    txSheet.getRange(txLastRow + 1, 1, 1, 5).setValues([newTxRow]);
    txSheet.getRange(txLastRow + 1, 2, 1, 1).setNumberFormat("dddd, MMMM d, yyyy");
    txSheet.getRange(txLastRow + 1, 5, 1, 1).setNumberFormat("#,##0.00");

    // Mark row as Added in Reconciliation sheet
    sheet.getRange(row, 1, 1, 6).setBackground(CONFIG.COLORS.PAID);
    sheet.getRange(row, 5).clearDataValidations().setValue("✅ Added (" + payer + ")");

    // Toast notification
    ss.toast(`Added "${finalDesc}" (${amt.toFixed(2)} EGP) for ${payer} to Transactions!`, "💳 Transaction Added", 4);
  }

  // Table 1: User confirms a Suggested Match in Column G (col 7)
  if (col === 7 && (e.value === "TRUE" || e.range.getValue() === true)) {
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
    }
  }
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
        `The Bank Statement, Reconciliation, and Monthly Debts tabs have all been updated!`,
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
  if (!installmentsSheet || installmentsSheet.getLastRow() < 2) return;

  const statementInstallments = transactions.filter(t => t.type === "INSTALLMENT");
  if (statementInstallments.length === 0) return;

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

  matches.forEach(m => {
    const instCurrent = m.stmt.instCurrent;
    if (instCurrent != null) {
      m.sheetRows.forEach(sh => {
        values[sh.index][9] = instCurrent; // Payments Made (Column J)
        values[sh.index][10] = (instCurrent >= sh.duration) ? "Completed" : "Ongoing"; // Status (Column K)
      });
    }
  });

  range.setValues(values);
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

function runReconciliation(ss, tz) {
  const reconSheet = getOrCreateSheet(ss, CONFIG.SHEETS.RECONCILIATION);
  reconSheet.clear();

  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);

  if (!txSheet || !stmtSheet || stmtSheet.getLastRow() < 7) {
    reconSheet.getRange("A1").setValue("Please upload and process a bank statement to view reconciliation.");
    return;
  }

  // 1. Read Statement Debits
  const stmtLastRow = stmtSheet.getLastRow();
  const stmtData = stmtSheet.getRange(7, 1, stmtLastRow - 6, 7).getValues();
  const stmtDebits = [];

  stmtData.forEach((row, idx) => {
    const type = String(row[4]).trim();
    if (type === "DEBIT") {
      const pDate = parseDateValue(row[1], tz);
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

  // 2. Read Sheet Transactions
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
      if (pDate && !isNaN(amt)) {
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
  const scriptProps = PropertiesService.getScriptProperties();
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

  // 4. Render Reconciliation Sheet
  reconSheet.getRange("A1:G1").merge()
    .setValue("Statement Reconciliation — Smart Audit & Discrepancies")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  const confirmedMatchesCount = stmtDebits.filter(st => st.matched).length;
  const kpiRow = [
    "Confirmed Matched:", confirmedMatchesCount,
    "Needs Confirmation:", suggestedMatches.length,
    "Missing in Sheet:", missingInSheet.length
  ];
  reconSheet.getRange("A2:F2").setValues([kpiRow])
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);
  reconSheet.getRange("B2").setFontColor(CONFIG.COLORS.PRIMARY);
  reconSheet.getRange("D2").setFontColor("#b06000");
  reconSheet.getRange("F2").setFontColor("#b71c1c");

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

  // TABLE 2: ⚠️ Missing Charges from Sheet
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
    const payerOptions = ["Mido", "Mai", "Abdo", "Dad", "Mum", "Zoza", "Shared"];
    const payerValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(payerOptions, true)
      .setAllowInvalid(true)
      .build();

    const missRows = missingInSheet.map((m, idx) => [
      idx + 1,
      m.date || m.dateStr,
      m.desc,
      m.amount,
      "",
      cleanMerchantName(m.desc)
    ]);

    const missRange = reconSheet.getRange(curRow, 1, missRows.length, 6);
    missRange.setValues(missRows);
    reconSheet.getRange(curRow, 2, missRows.length, 1).setNumberFormat("dddd, MMMM d, yyyy");
    reconSheet.getRange(curRow, 4, missRows.length, 1).setNumberFormat("#,##0.00");
    reconSheet.getRange(curRow, 5, missRows.length, 1).setDataValidation(payerValidation);
    curRow += missRows.length + 1;
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
        stmtRowStatus[idx] = { text: "⚠️ Unassigned (Not in Sheet)", color: CONFIG.COLORS.OVERDUE };
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

  reconSheet.autoResizeColumns(1, 7);
}

function addAssignedChargesToTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reconSheet = ss.getSheetByName(CONFIG.SHEETS.RECONCILIATION);
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!reconSheet || !txSheet) return;

  const lastRow = reconSheet.getLastRow();
  if (lastRow < 5) return;

  // Locate Table 2: "MISSING from Sheet"
  const colA = reconSheet.getRange(1, 1, lastRow, 1).getValues();
  let table2Start = -1;
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).includes("MISSING from Sheet")) {
      table2Start = i + 1;
      break;
    }
  }

  if (table2Start === -1 || table2Start + 2 > lastRow) {
    SpreadsheetApp.getUi().alert(
      "No Missing Charges Found",
      "No unassigned charges found in the Reconciliation sheet.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const startRow = table2Start + 2;
  const numRows = lastRow - startRow + 1;
  const data = reconSheet.getRange(startRow, 1, numRows, 6).getValues();
  const toAdd = [];
  const validPayers = ["Mido", "Mai", "Abdo", "Dad", "Mum", "Zoza", "Shared"];

  // Read existing transactions to prevent duplicates
  const existingTx = txSheet.getLastRow() >= 2 ? txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 5).getValues() : [];

  data.forEach((row, idx) => {
    const amt = parseFloat(row[3]);
    const payer = String(row[4] || "").trim();
    const note = String(row[5] || "").trim();

    if (!isNaN(amt) && amt > 0 && validPayers.includes(payer)) {
      const rawDate = row[1];
      const desc = note || String(row[2]);
      const dateObj = new Date(rawDate);
      const finalDate = isNaN(dateObj.getTime()) ? new Date() : dateObj;

      const alreadyExists = existingTx.some(r => {
        const rAmt = parseFloat(r[4]);
        const rDesc = String(r[2] || "").trim().toLowerCase();
        return Math.abs(rAmt - amt) < 0.05 && rDesc === desc.toLowerCase();
      });

      if (!alreadyExists) {
        toAdd.push({
          date: finalDate,
          desc: desc,
          payer: payer,
          amount: amt
        });
      }
    }
  });

  if (toAdd.length === 0) {
    SpreadsheetApp.getUi().alert(
      "No New Charges Selected",
      "Please select a name in 'Assign Payer' (Column E) for any missing charges you want to add to Transactions.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  let txLastRow = txSheet.getLastRow();
  let nextId = 1;
  if (txLastRow >= 2) {
    const lastIdVal = txSheet.getRange(txLastRow, 1).getValue();
    if (!isNaN(parseInt(lastIdVal, 10))) {
      nextId = parseInt(lastIdVal, 10) + 1;
    }
  }

  const newRows = toAdd.map(item => [
    nextId++,
    item.date,
    item.desc,
    item.payer,
    item.amount
  ]);

  txSheet.getRange(txLastRow + 1, 1, newRows.length, 5).setValues(newRows);
  txSheet.getRange(txLastRow + 1, 2, newRows.length, 1).setNumberFormat("dddd, MMMM d, yyyy");
  txSheet.getRange(txLastRow + 1, 5, newRows.length, 1).setNumberFormat("#,##0.00");

  updateLiveDashboard();

  SpreadsheetApp.getUi().alert(
    "Charges Added Successfully",
    `Successfully added ${toAdd.length} charge(s) to the Transactions sheet!\n` +
    `The Reconciliation and Monthly Debts have been updated.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==========================================
// 6. DASHBOARD & 100K AVAILABLE BALANCE
// ==========================================

function calculateCardBalance(ss, tz) {
  const creditLimit = CONFIG.CREDIT_LIMIT;
  let billedBalance = 0;
  let statementDate = null;

  // Read latest statement closing balance if available
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  if (stmtSheet && stmtSheet.getLastRow() >= 4) {
    const closeVal = stmtSheet.getRange("D3").getValue();
    if (!isNaN(parseFloat(closeVal))) {
      billedBalance = parseFloat(closeVal);
    }
    const stmtDateVal = stmtSheet.getRange("B2").getValue();
    statementDate = parseDateValue(stmtDateVal, tz);
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
    if (statementDate) {
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
    statementDate: statementDate,
    isCurrentBillPaid: isCurrentBillPaid,
    totalBlockedInstallments: totalBlockedInstallments,
    unbilledNewPurchases: unbilledNewPurchases,
    totalUtilized: totalUtilized,
    availableBalanceNow: availableBalanceNow,
    availableAfterSettlement: availableAfterSettlement
  };
}

function updateLiveDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";

  const transactionsSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  const installmentsSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  const debtsSheet = getOrCreateSheet(ss, CONFIG.SHEETS.MONTHLY_DEBTS);
  const historySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_HISTORY);

  if (!transactionsSheet || !installmentsSheet || !historySheet) {
    SpreadsheetApp.getUi().alert(
      "Missing Sheet",
      "Please make sure 'Transactions', 'Installments', and 'Payment History' sheets exist.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Clean duplicate transactions if any were added
  cleanTransactionsSheet(ss);

  // Run reconciliation
  runReconciliation(ss, tz);

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

  function addDebtItem({ dueDate, person, category, desc, installmentInfo, purchaseDate, originalAmount, amount }) {
    const sortKey = Utilities.formatDate(dueDate, tz, "yyyy-MM");
    const label = Utilities.formatDate(dueDate, tz, "MMMM yyyy");
    const rawP = String(person || "").trim();
    const normalizedPerson = rawP.charAt(0).toUpperCase() + rawP.slice(1);

    if (!allDebts[sortKey]) {
      allDebts[sortKey] = {
        label: label,
        dueDate: dueDate,
        total: 0,
        purchasesTotal: 0,
        installmentsTotal: 0,
        people: {},
        peopleBreakdown: {}
      };
    }

    allDebts[sortKey].total += amount;
    if (category === "Purchase") {
      allDebts[sortKey].purchasesTotal += amount;
    } else {
      allDebts[sortKey].installmentsTotal += amount;
    }

    allDebts[sortKey].people[normalizedPerson] = (allDebts[sortKey].people[normalizedPerson] || 0) + amount;

    if (!allDebts[sortKey].peopleBreakdown[normalizedPerson]) {
      allDebts[sortKey].peopleBreakdown[normalizedPerson] = { purchases: 0, installments: 0, total: 0 };
    }
    if (category === "Purchase") {
      allDebts[sortKey].peopleBreakdown[normalizedPerson].purchases += amount;
    } else {
      allDebts[sortKey].peopleBreakdown[normalizedPerson].installments += amount;
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
      amount: amount
    });
  }

  // 1. Process One-Time Payments
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

      const dueDate = getDueDateForPurchase(purchaseDate, tz);
      if (!dueDate) return;

      const rawPayerStr = String(rawPerson).trim();
      const people = rawPayerStr.split(/[\+\,\/]/).map(p => p.trim()).filter(p => p);
      const splitAmount = amount / Math.max(1, people.length);

      people.forEach(p => {
        addDebtItem({
          dueDate: dueDate,
          person: p,
          category: "Purchase",
          desc: desc,
          installmentInfo: "-",
          purchaseDate: purchaseDate,
          originalAmount: amount,
          amount: splitAmount
        });
      });
    });
  }

  // 2. Process Installments (Exact NBE Cutoff Rule)
  if (installmentsSheet.getLastRow() >= 2) {
    const instData = installmentsSheet.getRange(2, 1, installmentsSheet.getLastRow() - 1, 11).getValues();
    const newInstallmentStatuses = [];

    instData.forEach(row => {
      const rawDate = row[1];
      const desc = String(row[2] || "").trim();
      const durationMonths = parseInt(row[4], 10);
      const emi = parseFloat(row[7]);
      const rawPayer = row[8];
      const existingPaymentsMade = parseInt(row[9], 10) || 0;
      const existingStatus = String(row[10] || "").trim().toLowerCase();

      if (!rawDate || isNaN(durationMonths) || durationMonths <= 0 || isNaN(emi) || emi <= 0 || !rawPayer) {
        newInstallmentStatuses.push([row[9] || "", row[10] || ""]);
        return;
      }

      const purchaseDate = parseDateValue(rawDate, tz);
      if (!purchaseDate) {
        newInstallmentStatuses.push([row[9] || "", row[10] || ""]);
        return;
      }

      const rawPayerStr = String(rawPayer).trim();
      const people = rawPayerStr.split(/[\+\,\/]/).map(p => p.trim()).filter(p => p);
      const splitEmi = emi / Math.max(1, people.length);

      // First installment is due in the statement cycle of the purchase date
      const firstDueDate = getDueDateForPurchase(purchaseDate, tz);

      let paymentsCounted = 0;
      for (let i = 0; i < durationMonths; i++) {
        const dueMonthDate = new Date(firstDueDate.getFullYear(), firstDueDate.getMonth() + i, 1);
        const monthKey = Utilities.formatDate(dueMonthDate, tz, "MMMM yyyy").toLowerCase();

        if (paidMonthsSet.has(monthKey)) {
          paymentsCounted++;
        }

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
            amount: splitEmi
          });
        });
      }

      // Preserve existing payments made if higher (e.g. from bank statement sync)
      const finalPaymentsMade = Math.max(existingPaymentsMade, paymentsCounted);
      const isCompleted = (finalPaymentsMade >= durationMonths) || (existingStatus === "completed");
      newInstallmentStatuses.push([finalPaymentsMade, isCompleted ? "Completed" : "Ongoing"]);
    });

    if (newInstallmentStatuses.length > 0) {
      installmentsSheet.getRange(2, 10, newInstallmentStatuses.length, 2).setValues(newInstallmentStatuses);
    }
  }

  // 3. Render Dashboard in 'Monthly Debts'
  debtsSheet.clear();

  // Overview Card
  const cardBal = calculateCardBalance(ss, tz);

  debtsSheet.getRange("A1:B1").merge()
    .setValue("💳 NBE Credit Card Overview (Limit: 100,000 EGP)")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  const overviewRows = [
    ["Total Credit Limit:", cardBal.creditLimit],
    ["Available Balance Right Now:", cardBal.availableBalanceNow],
    ["Available After Month Settled:", cardBal.availableAfterSettlement],
    ["Blocked Remaining Installments:", cardBal.totalBlockedInstallments],
    ["Billed Statement Balance:", cardBal.billedBalance],
    ["Unbilled New Purchases:", cardBal.unbilledNewPurchases]
  ];

  debtsSheet.getRange(2, 1, overviewRows.length, 2).setValues(overviewRows);
  debtsSheet.getRange(2, 1, overviewRows.length, 1).setFontWeight("bold").setFontSize(10);
  debtsSheet.getRange(2, 2, overviewRows.length, 1).setFontWeight("bold").setFontSize(10).setNumberFormat("#,##0.00");
  debtsSheet.getRange("A2:B7").setBackground(CONFIG.COLORS.PRIMARY_LIGHT);
  debtsSheet.getRange("A3:B3").setBackground(CONFIG.COLORS.DUE_SOON);

  let currentRow = 9;

  const sortedKeys = Object.keys(allDebts).sort();
  if (sortedKeys.length === 0) {
    debtsSheet.getRange(currentRow, 1).setValue("No transaction or installment data found.");
    return;
  }

  // Collect and sort unique people
  const allPeopleSet = new Set();
  debtLineItems.forEach(it => allPeopleSet.add(it.person));
  const sortedPeople = Array.from(allPeopleSet).sort((a, b) => {
    if (a.toLowerCase() === "shared") return 1;
    if (b.toLowerCase() === "shared") return -1;
    return a.localeCompare(b);
  });

  const today = new Date();
  const currentMonthSortKey = Utilities.formatDate(today, tz, "yyyy-MM");
  const todayDay = parseInt(Utilities.formatDate(today, tz, "d"), 10);

  let stmtDueMonthSortKey = null;
  if (cardBal.statementDate) {
    const sDueDate = getDueDateForPurchase(cardBal.statementDate, tz);
    if (sDueDate) {
      stmtDueMonthSortKey = Utilities.formatDate(sDueDate, tz, "yyyy-MM");
    }
  }

  // ==========================================
  // MASTER MATRIX TABLE (DIVIDED BY PERSON IN ONE TABLE)
  // ==========================================
  const totalCols = 5 + sortedPeople.length;

  debtsSheet.getRange(currentRow, 1, 1, totalCols).merge()
    .setValue("📊 Master Monthly Debt Matrix (All Persons Divided in One Table)")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");
  currentRow++;

  debtsSheet.getRange(currentRow, 1, 1, totalCols).merge()
    .setValue("Complete overview of all months, due dates, bank statement bills, and individual family shares side-by-side.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);
  currentRow++;

  const matrixHeaders = ["Due Month", "Status", "Family Total Due", "Bank Bill (EGP)", "Audit Diff (EGP)", ...sortedPeople];
  debtsSheet.getRange(currentRow, 1, 1, matrixHeaders.length)
    .setValues([matrixHeaders])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, true, true);
  currentRow++;

  const matrixRows = [];
  const matrixRowColors = [];

  let grandTotalDebt = 0;
  let grandBankBill = 0;
  const personTotals = {};
  sortedPeople.forEach(p => { personTotals[p] = 0; });

  sortedKeys.forEach(key => {
    const data = allDebts[key];
    const monthLabel = data.label;
    const isPaid = paidMonthsSet.has(monthLabel.toLowerCase());
    let statusLabel = "🕒 Upcoming";
    let statusColor = CONFIG.COLORS.UPCOMING;

    if (isPaid) {
      statusLabel = "✅ Paid";
      statusColor = CONFIG.COLORS.PAID;
    } else if (key < currentMonthSortKey || (key === currentMonthSortKey && todayDay > 25)) {
      statusLabel = "🚨 Overdue";
      statusColor = CONFIG.COLORS.OVERDUE;
    } else if (key === currentMonthSortKey && todayDay <= 25) {
      statusLabel = "⚠️ Due Soon (25th)";
      statusColor = CONFIG.COLORS.DUE_SOON;
    }

    const isStatementMonth = (key === stmtDueMonthSortKey) || (!stmtDueMonthSortKey && (statusLabel.includes("Due Soon") || key === currentMonthSortKey));
    const bankBillVal = (cardBal.billedBalance > 0 && isStatementMonth) ? cardBal.billedBalance : "";
    const diffVal = (typeof bankBillVal === "number") ? (data.total - bankBillVal) : "";

    grandTotalDebt += data.total;
    if (typeof bankBillVal === "number") grandBankBill += bankBillVal;

    const row = [
      monthLabel,
      statusLabel,
      data.total,
      bankBillVal,
      diffVal
    ];

    sortedPeople.forEach(p => {
      const pAmt = data.people[p] || 0;
      row.push(pAmt);
      personTotals[p] += pAmt;
    });

    matrixRows.push(row);
    matrixRowColors.push(statusColor);
  });

  // Render Matrix Data Rows
  const matrixRange = debtsSheet.getRange(currentRow, 1, matrixRows.length, totalCols);
  matrixRange.setValues(matrixRows);

  for (let r = 0; r < matrixRows.length; r++) {
    debtsSheet.getRange(currentRow + r, 2).setBackground(matrixRowColors[r]).setFontWeight("bold");
    if (matrixRows[r][1].includes("Paid")) {
      debtsSheet.getRange(currentRow + r, 1).setFontLine("line-through");
    }
  }

  // Format currency columns
  debtsSheet.getRange(currentRow, 3, matrixRows.length, totalCols - 2).setNumberFormat("#,##0.00");
  debtsSheet.getRange(currentRow, 1, matrixRows.length, 1).setFontWeight("bold");
  currentRow += matrixRows.length;

  // Grand Total Summary Row
  const totalSummaryRow = ["TOTAL ALL MONTHS", "", grandTotalDebt, (grandBankBill > 0 ? grandBankBill : ""), ""];
  sortedPeople.forEach(p => totalSummaryRow.push(personTotals[p]));

  debtsSheet.getRange(currentRow, 1, 1, totalCols)
    .setValues([totalSummaryRow])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT)
    .setBorder(true, true, true, true, true, true);
  debtsSheet.getRange(currentRow, 3, 1, totalCols - 2).setNumberFormat("#,##0.00");
  currentRow += 3;

  // ==========================================
  // DETAILED MONTHLY DUES CARDS (WITH SPLIT BREAKDOWN)
  // ==========================================
  debtsSheet.getRange(currentRow, 1, 1, 3).merge()
    .setValue("📑 Detailed Monthly Dues (Cards & Category Breakdown)")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(CONFIG.COLORS.HEADER);
  currentRow += 2;

  sortedKeys.forEach(key => {
    const data = allDebts[key];
    const monthLabel = data.label;
    const isPaid = paidMonthsSet.has(monthLabel.toLowerCase());

    let blockColor = null;
    let isOverdue = false;
    let isDueSoon = false;

    if (isPaid) {
      blockColor = CONFIG.COLORS.PAID;
    } else if (key < currentMonthSortKey || (key === currentMonthSortKey && todayDay > 25)) {
      isOverdue = true;
      blockColor = CONFIG.COLORS.OVERDUE;
    } else if (key === currentMonthSortKey && todayDay <= 25) {
      isDueSoon = true;
      blockColor = CONFIG.COLORS.DUE_SOON;
    }

    // Title Row
    const titleRange = debtsSheet.getRange(currentRow, 1, 1, 3);
    titleRange.merge();
    let headerText = monthLabel;
    if (isOverdue) headerText += " (OVERDUE)";
    else if (isDueSoon) headerText += " (DUE BY 25th)";
    else if (isPaid) headerText += " (PAID)";

    titleRange.setValue(headerText)
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground(blockColor || CONFIG.COLORS.HEADER);
    if (isPaid) titleRange.setFontLine("line-through");
    currentRow++;

    // Data Rows
    const rows = [];
    const isStatementMonth = (key === stmtDueMonthSortKey) || (!stmtDueMonthSortKey && (isDueSoon || key === currentMonthSortKey));
    let bankDueOffset = -1;

    if (cardBal.billedBalance > 0 && isStatementMonth) {
      bankDueOffset = rows.length;
      rows.push(["🏦 BANK STATEMENT DUE (Must Pay):", cardBal.billedBalance, "Exact statement closing balance from NBE"]);
      rows.push(["👥 Sum of Individual Family Shares:", data.total, "Total purchases & installments in this cycle"]);

      const diff = data.total - cardBal.billedBalance;
      if (Math.abs(diff) < 1.00) {
        rows.push(["⚖️ Audit & Reconciliation Status:", "✅ Matched (0.00 EGP difference)", "All family shares perfectly match bank bill"]);
      } else if (diff > 1.00) {
        rows.push(["⚖️ Audit Difference:", "+" + diff.toFixed(2) + " EGP", "Family debts exceed bank bill (prior overpayment applied)"]);
        rows.push(["👉 What to Pay:", cardBal.billedBalance, "Transfer billed amount to NBE to settle card"]);
      } else {
        rows.push(["⚠️ Audit Difference:", "-" + Math.abs(diff).toFixed(2) + " EGP", "Bank bill is higher! Check Reconciliation for unassigned items"]);
      }
    } else {
      rows.push(["Total Bill (Purchases & EMIs):", data.total, `Purchases: ${data.purchasesTotal.toFixed(2)} | EMIs: ${data.installmentsTotal.toFixed(2)}`]);
    }

    const personHeaderIdx = rows.length;
    rows.push(["Person", "Owes (EGP)", "Category Breakdown"]);

    const activePeople = Object.keys(data.people).sort();
    activePeople.forEach(p => {
      const pTotal = data.people[p];
      const bk = data.peopleBreakdown[p] || { purchases: 0, installments: 0 };
      const bkStr = `Purchases: ${bk.purchases.toFixed(2)} | EMIs: ${bk.installments.toFixed(2)}`;
      rows.push([p, pTotal, bkStr]);
    });

    const dataRange = debtsSheet.getRange(currentRow, 1, rows.length, 3);
    dataRange.setValues(rows);

    // Number formatting
    for (let r = 0; r < rows.length; r++) {
      if (typeof rows[r][1] === "number") {
        debtsSheet.getRange(currentRow + r, 2).setNumberFormat("#,##0.00");
      } else {
        debtsSheet.getRange(currentRow + r, 2).setNumberFormat("@");
      }
    }

    // Typography & Styling
    debtsSheet.getRange(currentRow, 1, personHeaderIdx, 3).setFontWeight("bold");
    debtsSheet.getRange(currentRow + personHeaderIdx, 1, 1, 3).setFontStyle("italic").setFontColor(CONFIG.COLORS.TEXT_MUTED);

    if (blockColor) dataRange.setBackground(blockColor);

    if (bankDueOffset >= 0) {
      debtsSheet.getRange(currentRow + bankDueOffset, 1, 1, 3)
        .setBackground(CONFIG.COLORS.PRIMARY_LIGHT)
        .setFontWeight("bold")
        .setFontColor(CONFIG.COLORS.PRIMARY);
    }

    if (isPaid) dataRange.setFontLine("line-through");

    currentRow += rows.length + 1;
  });

  debtsSheet.autoResizeColumns(1, totalCols);

  // 4. Render Dedicated 'Debt Breakdown' Sheet
  renderDebtBreakdownSheet(ss, tz, allDebts, debtLineItems, sortedPeople, paidMonthsSet, cardBal);
}

function renderDebtBreakdownSheet(ss, tz, allDebts, debtLineItems, sortedPeople, paidMonthsSet, cardBal) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.DEBT_BREAKDOWN);
  sheet.clear();

  // Banner
  sheet.getRange("A1:J1").merge()
    .setValue("💳 NBE Credit Card — Complete Itemized Debt Breakdown & Audit Register")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  sheet.getRange("A2:J2").merge()
    .setValue("Every single purchase and installment installment divided by person and due month. Filter or sort by any column to audit.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor(CONFIG.COLORS.TEXT_MUTED);

  let curRow = 4;

  // ==========================================
  // SECTION 1: PERSON TOTALS & DEBT COMPOSITION
  // ==========================================
  sheet.getRange(curRow, 1, 1, 6).merge()
    .setValue("👥 Family Members — Debt Summary & Composition")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  const summaryHeaders = ["Person", "One-Time Purchases (EGP)", "Installments (EGP)", "Total Debt (EGP)", "% of Total Debt", "Due / Overdue Now (EGP)"];
  sheet.getRange(curRow, 1, 1, summaryHeaders.length).setValues([summaryHeaders])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, true, true);
  curRow++;

  const today = new Date();
  const currentMonthSortKey = Utilities.formatDate(today, tz, "yyyy-MM");

  let totalPurchasesAll = 0;
  let totalInstallmentsAll = 0;
  let grandTotalAll = 0;
  let grandDueNow = 0;

  const personStats = {};
  sortedPeople.forEach(p => {
    personStats[p] = { purchases: 0, installments: 0, total: 0, dueNow: 0 };
  });

  debtLineItems.forEach(it => {
    if (!personStats[it.person]) {
      personStats[it.person] = { purchases: 0, installments: 0, total: 0, dueNow: 0 };
    }
    const isPaid = paidMonthsSet.has(it.dueMonthLabel.toLowerCase());
    const isDueOrOverdue = !isPaid && (it.sortKey <= currentMonthSortKey);

    if (it.category === "Purchase") {
      personStats[it.person].purchases += it.amount;
      totalPurchasesAll += it.amount;
    } else {
      personStats[it.person].installments += it.amount;
      totalInstallmentsAll += it.amount;
    }
    personStats[it.person].total += it.amount;
    grandTotalAll += it.amount;

    if (isDueOrOverdue) {
      personStats[it.person].dueNow += it.amount;
      grandDueNow += it.amount;
    }
  });

  const sumRows = sortedPeople.map(p => {
    const st = personStats[p];
    const pct = grandTotalAll > 0 ? (st.total / grandTotalAll) : 0;
    return [p, st.purchases, st.installments, st.total, pct, st.dueNow];
  });

  if (sumRows.length > 0) {
    sheet.getRange(curRow, 1, sumRows.length, summaryHeaders.length).setValues(sumRows);
    sheet.getRange(curRow, 2, sumRows.length, 3).setNumberFormat("#,##0.00");
    sheet.getRange(curRow, 5, sumRows.length, 1).setNumberFormat("0.0%");
    sheet.getRange(curRow, 6, sumRows.length, 1).setNumberFormat("#,##0.00");
    sheet.getRange(curRow, 1, sumRows.length, 1).setFontWeight("bold");
    curRow += sumRows.length;
  }

  // Family Total Row
  const totalRow = ["FAMILY TOTAL", totalPurchasesAll, totalInstallmentsAll, grandTotalAll, 1.00, grandDueNow];
  sheet.getRange(curRow, 1, 1, summaryHeaders.length).setValues([totalRow])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.PRIMARY_LIGHT)
    .setBorder(true, true, true, true, true, true);
  sheet.getRange(curRow, 2, 1, 3).setNumberFormat("#,##0.00");
  sheet.getRange(curRow, 5, 1, 1).setNumberFormat("0.0%");
  sheet.getRange(curRow, 6, 1, 1).setNumberFormat("#,##0.00");
  curRow += 3;

  // ==========================================
  // SECTION 2: ITEMIZED REGISTER (EVERY CHARGE ACCOUNTED FOR)
  // ==========================================
  sheet.getRange(curRow, 1, 1, 10).merge()
    .setValue("📋 Itemized Dues Register (Every Single Charge Divided & Accounted For)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  const itemHeaders = ["#", "Due Month", "Person", "Category", "Merchant / Description", "Installment Progress", "Purchase Date", "Person's Share (EGP)", "Original Charge (EGP)", "Month Status"];
  sheet.getRange(curRow, 1, 1, itemHeaders.length).setValues([itemHeaders])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, true, true);
  curRow++;

  const headerRowForFilter = curRow - 1;

  // Sort debtLineItems: sortKey ascending, person ascending, category ("Purchase" first)
  debtLineItems.sort((a, b) => {
    if (a.sortKey !== b.sortKey) return a.sortKey.localeCompare(b.sortKey);
    if (a.person !== b.person) return a.person.localeCompare(b.person);
    if (a.category !== b.category) return a.category === "Purchase" ? -1 : 1;
    return a.desc.localeCompare(b.desc);
  });

  const todayDay = parseInt(Utilities.formatDate(today, tz, "d"), 10);
  const regRows = [];
  const regRowColors = [];

  debtLineItems.forEach((it, idx) => {
    const isPaid = paidMonthsSet.has(it.dueMonthLabel.toLowerCase());
    let statusLabel = "🕒 Upcoming";

    if (isPaid) {
      statusLabel = "✅ Paid";
    } else if (it.sortKey < currentMonthSortKey || (it.sortKey === currentMonthSortKey && todayDay > 25)) {
      statusLabel = "🚨 Overdue";
    } else if (it.sortKey === currentMonthSortKey && todayDay <= 25) {
      statusLabel = "⚠️ Due Soon (25th)";
    }

    regRows.push([
      idx + 1,
      it.dueMonthLabel,
      it.person,
      it.category === "Purchase" ? "🛒 Purchase" : "📦 Installment",
      it.desc,
      it.installmentInfo,
      it.purchaseDate,
      it.amount,
      it.originalAmount,
      statusLabel
    ]);

    if (it.category === "Installment") {
      regRowColors.push(CONFIG.COLORS.ACCENT_BLUE);
    } else if (isPaid) {
      regRowColors.push(CONFIG.COLORS.PAID);
    } else {
      regRowColors.push(null);
    }
  });

  if (regRows.length > 0) {
    const regRange = sheet.getRange(curRow, 1, regRows.length, itemHeaders.length);
    regRange.setValues(regRows);

    sheet.getRange(curRow, 7, regRows.length, 1).setNumberFormat("yyyy-MM-dd");
    sheet.getRange(curRow, 8, regRows.length, 2).setNumberFormat("#,##0.00");
    sheet.getRange(curRow, 1, regRows.length, 1).setFontStyle("italic").setFontColor(CONFIG.COLORS.TEXT_MUTED);

    for (let r = 0; r < regRows.length; r++) {
      if (regRowColors[r]) {
        sheet.getRange(curRow + r, 1, 1, itemHeaders.length).setBackground(regRowColors[r]);
      }
    }

    const filterRange = sheet.getRange(headerRowForFilter, 1, regRows.length + 1, itemHeaders.length);
    try {
      const existingFilter = sheet.getFilter();
      if (existingFilter) existingFilter.remove();
      filterRange.createFilter();
    } catch (e) {}
  }

  sheet.autoResizeColumns(1, 10);
}