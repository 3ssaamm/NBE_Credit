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
    PAYMENT_HISTORY: "Payment History",
    BANK_STATEMENT: "Bank Statement",
    RECONCILIATION: "Reconciliation"
  },
  COLORS: {
    PAID: "#d9ead3",       // Soft green
    DUE_SOON: "#fff2cc",   // Soft yellow
    OVERDUE: "#fce8e6",    // Soft red
    HEADER: "#f3f3f3",     // Light neutral grey
    PRIMARY: "#1b5e20",    // NBE dark green
    PRIMARY_LIGHT: "#e8f5e9",
    TEXT_MUTED: "#555555",
    BORDER: "#e0e0e0"
  }
};

// ==========================================
// 1. MENU & TRIGGER SETUP
// ==========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💳 NBE Tracker")
    .addItem("📥 Process Latest Statement (Drive)", "menuProcessStatement")
    .addItem("🔄 Refresh Dashboard & Reconciliation", "updateLiveDashboard")
    .addSeparator()
    .addItem("➕ Add Checked Charges to Transactions", "addAssignedChargesToTransactions")
    .addSeparator()
    .addItem("⏰ Setup Daily Auto-Check", "setupDailyTrigger")
    .addToUi();
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
  sheet.getRange("B2:B4").setNumberFormat("#,##0.00");
  sheet.getRange("D2:D4").setNumberFormat("#,##0.00");
  sheet.getRange("F2:F4").setNumberFormat("#,##0.00");
  sheet.getRange("2:4").setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

  // Table Headers
  const startRow = 6;
  const headers = ["#", "Tx Date", "Posting Date", "Description", "Type", "Amount (EGP)", "Auth Code"];
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, false, false);

  if (transactions.length > 0) {
    const rows = transactions.map((t, idx) => [
      idx + 1,
      t.txDate,
      t.postDate,
      t.desc,
      t.type,
      t.amount,
      t.authCode
    ]);

    const dataRange = sheet.getRange(startRow + 1, 1, rows.length, headers.length);
    dataRange.setValues(rows);
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
// 4. INSTALLMENTS AUTO-SYNC FROM STATEMENT
// ==========================================

function syncInstallmentsFromStatement(ss, transactions) {
  const installmentsSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  if (!installmentsSheet || installmentsSheet.getLastRow() < 2) return;

  const statementInstallments = transactions.filter(t => t.type === "INSTALLMENT");
  if (statementInstallments.length === 0) return;

  const numRows = installmentsSheet.getLastRow() - 1;
  const range = installmentsSheet.getRange(2, 1, numRows, 11);
  const values = range.getValues();

  values.forEach(row => {
    const desc = String(row[2] || "").toLowerCase().trim();
    const emi = parseFloat(row[7]);
    const duration = parseInt(row[4], 10);

    if (isNaN(emi) || emi <= 0) return;

    // Search statement installments by matching EMI (within 0.50 EGP margin)
    const match = statementInstallments.find(st => {
      const emiMatch = Math.abs(st.amount - emi) < 0.50;
      if (!emiMatch) return false;
      // If duration is specified on statement, match total
      if (st.instTotal && duration && Math.abs(st.instTotal - duration) > 1) {
        return false;
      }
      return true;
    });

    if (match && match.instCurrent != null) {
      row[9] = match.instCurrent; // Payments Made (Column J)
      row[10] = (match.instCurrent >= duration) ? "Completed" : "Ongoing"; // Status (Column K)
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
      const pDate = new Date(row[1]);
      stmtDebits.push({
        id: idx + 1,
        dateStr: String(row[1]),
        date: isNaN(pDate.getTime()) ? null : pDate,
        desc: String(row[3]),
        amount: parseFloat(row[5]),
        matched: false,
        matchType: null,
        matchedItems: []
      });
    }
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

      const pDate = new Date(rawDate);
      const amt = parseFloat(rawAmount);
      if (!isNaN(pDate.getTime()) && !isNaN(amt)) {
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

  // Pass 2: Split Matches (2 or 3 sheet items sum to 1 statement charge)
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

    // Check pairs (2-way split)
    let foundPair = false;
    for (let i = 0; i < candidates.length; i++) {
      for (let j = i + 1; j < candidates.length; j++) {
        const sum = candidates[i].amount + candidates[j].amount;
        if (Math.abs(sum - st.amount) < 0.10) {
          const keywordMatch = hasMerchantKeywordOverlap(st.desc, candidates[i].desc) ||
                               hasMerchantKeywordOverlap(st.desc, candidates[j].desc);
          if (keywordMatch) {
            st.matched = true;
            st.matchType = "SPLIT";
            candidates[i].matched = true;
            candidates[j].matched = true;
            st.matchedItems.push(candidates[i], candidates[j]);
            foundPair = true;
            break;
          }
        }
      }
      if (foundPair) break;
    }

    if (foundPair) return;

    // Check triplets (3-way split)
    let foundTriplet = false;
    for (let i = 0; i < candidates.length; i++) {
      for (let j = i + 1; j < candidates.length; j++) {
        for (let k = j + 1; k < candidates.length; k++) {
          const sum = candidates[i].amount + candidates[j].amount + candidates[k].amount;
          if (Math.abs(sum - st.amount) < 0.10) {
            const keywordMatch = hasMerchantKeywordOverlap(st.desc, candidates[i].desc) ||
                                 hasMerchantKeywordOverlap(st.desc, candidates[j].desc) ||
                                 hasMerchantKeywordOverlap(st.desc, candidates[k].desc);
            if (keywordMatch) {
              st.matched = true;
              st.matchType = "SPLIT";
              candidates[i].matched = true;
              candidates[j].matched = true;
              candidates[k].matched = true;
              st.matchedItems.push(candidates[i], candidates[j], candidates[k]);
              foundTriplet = true;
              break;
            }
          }
        }
        if (foundTriplet) break;
      }
      if (foundTriplet) break;
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
        return daysDiff <= 3;
      }
      return true;
    });

    const singleAmtMatch = nearby.find(tx => Math.abs(tx.amount - st.amount) < 0.05);
    if (singleAmtMatch) {
      suggestedMatches.push({
        stmt: st,
        sheetItems: [singleAmtMatch],
        reason: `Exact amount (${st.amount.toFixed(2)}), dates within 3 days — check shop name`
      });
      return;
    }

    for (let i = 0; i < nearby.length; i++) {
      for (let j = i + 1; j < nearby.length; j++) {
        const sum = nearby[i].amount + nearby[j].amount;
        if (Math.abs(sum - st.amount) < 0.10) {
          suggestedMatches.push({
            stmt: st,
            sheetItems: [nearby[i], nearby[j]],
            reason: `Split sum matches ${st.amount.toFixed(2)} (${nearby[i].person} + ${nearby[j].person})`
          });
          return;
        }
      }
    }
  });

  const suggestedStmtIds = new Set(suggestedMatches.map(s => s.stmt.id));
  const missingInSheet = stmtDebits.filter(st => !st.matched && !suggestedStmtIds.has(st.id));
  const unmatchedInSheet = txDebits.filter(tx => !tx.matched);

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
        s.stmt.dateStr,
        s.stmt.desc,
        s.stmt.amount,
        itemsDesc,
        s.reason,
        false
      ];
    });

    const sugRange = reconSheet.getRange(curRow, 1, sugRows.length, 7);
    sugRange.setValues(sugRows);
    reconSheet.getRange(curRow, 4, sugRows.length, 1).setNumberFormat("#,##0.00");
    reconSheet.getRange(curRow, 7, sugRows.length, 1).insertCheckboxes();
    curRow += sugRows.length + 1;
  }

  // TABLE 2: ⚠️ Missing Charges from Sheet
  reconSheet.getRange(curRow, 1, 1, 7).merge()
    .setValue("⚠️ Charges on Statement MISSING from Sheet (" + missingInSheet.length + " items)")
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.OVERDUE);
  curRow++;

  const missHeaders = ["#", "Date", "Description / Merchant", "Amount (EGP)", "Assign Payer", "Custom Note", "Add to Sheet?"];
  reconSheet.getRange(curRow, 1, 1, 7).setValues([missHeaders])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  if (missingInSheet.length === 0) {
    reconSheet.getRange(curRow, 1, 1, 7).merge()
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
      m.dateStr,
      m.desc,
      m.amount,
      "",
      cleanMerchantName(m.desc),
      false
    ]);

    const missRange = reconSheet.getRange(curRow, 1, missRows.length, 7);
    missRange.setValues(missRows);
    reconSheet.getRange(curRow, 4, missRows.length, 1).setNumberFormat("#,##0.00");
    reconSheet.getRange(curRow, 5, missRows.length, 1).setDataValidation(payerValidation);
    reconSheet.getRange(curRow, 7, missRows.length, 1).insertCheckboxes();
    curRow += missRows.length + 1;
  }

  // TABLE 3: 🕒 Unmatched Sheet Transactions
  reconSheet.getRange(curRow, 1, 1, 6).merge()
    .setValue("🕒 Logged Transactions NOT on Statement (" + unmatchedInSheet.length + " items)")
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  const unHeaders = ["#", "Date", "Description", "Amount (EGP)", "Payer", "Status"];
  reconSheet.getRange(curRow, 1, 1, 6).setValues([unHeaders])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER);
  curRow++;

  if (unmatchedInSheet.length === 0) {
    reconSheet.getRange(curRow, 1, 1, 6).merge()
      .setValue("✅ All sheet transactions match statement records!");
  } else {
    let stmtCutoffDate = null;
    if (stmtSheet && stmtSheet.getLastRow() >= 2) {
      const sVal = stmtSheet.getRange("B2").getValue();
      if (sVal) stmtCutoffDate = new Date(sVal);
    }

    const unRows = unmatchedInSheet.map((u, idx) => {
      let status = "🕒 Pending (Next Statement)";
      if (stmtCutoffDate && u.date <= stmtCutoffDate) {
        status = "⚠️ Unbilled Discrepancy (before cutoff)";
      }
      return [
        idx + 1,
        Utilities.formatDate(u.date, tz, "yyyy-MM-dd"),
        u.desc,
        u.amount,
        u.person,
        status
      ];
    });

    reconSheet.getRange(curRow, 1, unRows.length, 6).setValues(unRows);
    reconSheet.getRange(curRow, 4, unRows.length, 1).setNumberFormat("#,##0.00");
  }

  reconSheet.autoResizeColumns(1, 7);
}

function addAssignedChargesToTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reconSheet = ss.getSheetByName(CONFIG.SHEETS.RECONCILIATION);
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);

  if (!reconSheet || !txSheet) return;

  const lastRow = reconSheet.getLastRow();
  if (lastRow < 5) return;

  const data = reconSheet.getRange(1, 1, lastRow, 7).getValues();
  const toAdd = [];

  data.forEach(row => {
    const amt = parseFloat(row[3]);
    const payer = String(row[4] || "").trim();
    const note = String(row[5] || "").trim();
    const isChecked = row[6] === true;

    if (!isNaN(amt) && amt > 0 && payer !== "" && isChecked) {
      const rawDate = row[1];
      const dateObj = new Date(rawDate);
      toAdd.push({
        date: isNaN(dateObj.getTime()) ? new Date() : dateObj,
        desc: note || String(row[2]),
        payer: payer,
        amount: amt
      });
    }
  });

  if (toAdd.length === 0) {
    SpreadsheetApp.getUi().alert(
      "No Charges Selected",
      "Please assign a Payer (Column E) and check the 'Add to Sheet?' box (Column G) for any charges you want to add.",
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
    if (stmtDateVal) {
      statementDate = new Date(stmtDateVal);
    }
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

  // Calculate Unbilled New Purchases (after statement date)
  let unbilledNewPurchases = 0;
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (txSheet && txSheet.getLastRow() >= 2) {
    const txData = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 5).getValues();
    txData.forEach(row => {
      const pDate = new Date(row[1]);
      const amt = parseFloat(row[4]);
      if (!isNaN(pDate.getTime()) && !isNaN(amt) && amt > 0) {
        if (!statementDate || pDate > statementDate) {
          unbilledNewPurchases += amt;
        }
      }
    });
  }

  // Check if current month bill has been paid in Payment History
  const historySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_HISTORY);
  let isCurrentBillPaid = false;
  if (historySheet && historySheet.getLastRow() >= 2) {
    const paidMonths = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 1).getValues().flat();
    const currentMonthLabel = Utilities.formatDate(new Date(), tz, "MMMM yyyy").toLowerCase();
    isCurrentBillPaid = paidMonths.some(m => String(m).trim().toLowerCase() === currentMonthLabel);
  }

  const effectiveBilled = isCurrentBillPaid ? 0 : billedBalance;
  const totalUtilized = effectiveBilled + totalBlockedInstallments + unbilledNewPurchases;
  const availableBalanceNow = Math.max(0, creditLimit - totalUtilized);
  const availableAfterSettlement = Math.max(0, creditLimit - totalBlockedInstallments - unbilledNewPurchases);

  return {
    creditLimit: creditLimit,
    billedBalance: billedBalance,
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

  function addDebt(dueDate, person, amount) {
    const sortKey = Utilities.formatDate(dueDate, tz, "yyyy-MM");
    const label = Utilities.formatDate(dueDate, tz, "MMMM yyyy");
    if (!allDebts[sortKey]) {
      allDebts[sortKey] = { label: label, dueDate: dueDate, total: 0, people: {} };
    }
    allDebts[sortKey].total += amount;
    allDebts[sortKey].people[person] = (allDebts[sortKey].people[person] || 0) + amount;
  }

  // Cutoff calculation: 29th cutoff
  function getDueDateForPurchase(purchaseDate) {
    const day = parseInt(Utilities.formatDate(purchaseDate, tz, "d"), 10);
    const month = parseInt(Utilities.formatDate(purchaseDate, tz, "M"), 10) - 1;
    const year = parseInt(Utilities.formatDate(purchaseDate, tz, "yyyy"), 10);
    const monthOffset = (day <= 29) ? 1 : 2;
    return new Date(year, month + monthOffset, 1);
  }

  // 1. Process One-Time Payments
  if (transactionsSheet.getLastRow() >= 2) {
    const txData = transactionsSheet.getRange(2, 1, transactionsSheet.getLastRow() - 1, 5).getValues();
    txData.forEach(row => {
      const rawDate = row[1], rawPerson = row[3], rawAmount = row[4];
      if (!rawDate || !rawPerson) return;

      const purchaseDate = new Date(rawDate);
      const amount = parseFloat(rawAmount);
      if (isNaN(purchaseDate.getTime()) || isNaN(amount) || amount <= 0) return;

      const person = String(rawPerson).trim().toLowerCase();
      const dueDate = getDueDateForPurchase(purchaseDate);
      addDebt(dueDate, person, amount);
    });
  }

  // 2. Process Installments
  if (installmentsSheet.getLastRow() >= 2) {
    const instData = installmentsSheet.getRange(2, 1, installmentsSheet.getLastRow() - 1, 11).getValues();
    instData.forEach(row => {
      const rawDate = row[1];
      const durationMonths = parseInt(row[4], 10);
      const emi = parseFloat(row[7]);
      const rawPayer = row[8];

      if (!rawDate || isNaN(durationMonths) || durationMonths <= 0 || isNaN(emi) || emi <= 0 || !rawPayer) {
        return;
      }

      const purchaseDate = new Date(rawDate);
      if (isNaN(purchaseDate.getTime())) return;

      const person = String(rawPayer).trim().toLowerCase();
      const firstDueDate = getDueDateForPurchase(purchaseDate);

      for (let i = 0; i < durationMonths; i++) {
        const instDueDate = new Date(firstDueDate.getFullYear(), firstDueDate.getMonth() + i, 1);
        addDebt(instDueDate, person, emi);
      }
    });
  }

  // 3. Render Dashboard in 'Monthly Debts'
  debtsSheet.clear();

  // Card Overview & 100k Available Balance Card
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
  debtsSheet.getRange("A3:B3").setBackground(CONFIG.COLORS.DUE_SOON); // Highlight Available Now

  let currentRow = 9;

  const sortedKeys = Object.keys(allDebts).sort();
  if (sortedKeys.length === 0) {
    debtsSheet.getRange(currentRow, 1).setValue("No transaction or installment data found.");
    return;
  }

  const today = new Date();
  const currentMonthSortKey = Utilities.formatDate(today, tz, "yyyy-MM");
  const todayDay = parseInt(Utilities.formatDate(today, tz, "d"), 10);

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
    const titleRange = debtsSheet.getRange(currentRow, 1, 1, 2);
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
    const rows = [
      ["Total Bill:", data.total],
      ["Person", "Owes"]
    ];

    const sortedPeople = Object.keys(data.people).sort();
    sortedPeople.forEach(p => {
      const formattedName = p.charAt(0).toUpperCase() + p.slice(1);
      rows.push([formattedName, data.people[p]]);
    });

    const dataRange = debtsSheet.getRange(currentRow, 1, rows.length, 2);
    dataRange.setValues(rows);

    // Number formatting for amounts (Column B)
    debtsSheet.getRange(currentRow, 2, 1, 1).setNumberFormat("#,##0.00");
    if (sortedPeople.length > 0) {
      debtsSheet.getRange(currentRow + 2, 2, sortedPeople.length, 1).setNumberFormat("#,##0.00");
    }

    // Typography & Styling
    debtsSheet.getRange(currentRow, 1, 1, 2).setFontWeight("bold");
    debtsSheet.getRange(currentRow + 1, 1, 1, 2).setFontStyle("italic").setFontColor(CONFIG.COLORS.TEXT_MUTED);

    if (blockColor) dataRange.setBackground(blockColor);
    if (isPaid) dataRange.setFontLine("line-through");

    currentRow += rows.length + 1;
  });

  debtsSheet.autoResizeColumns(1, 2);
}