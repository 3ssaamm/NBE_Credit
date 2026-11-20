/**
 * Statement PDF OCR, Parsing, Mathematical Tie-Out & Import Pipeline
 *
 * Parses NBE Credit Card PDF statements, enforces opening + debits - credits = closing
 * tie-out validation and line-count verification, routes ambiguous / suspected flipped
 * lines to explicit "Confirm Type" review (zero hardcoded merchant/amount hacks), gates
 * bill generation on tie-out pass, and preserves statement S-IDs idempotently on re-import.
 *
 * NOTE: Apps Script V8 load order across files is arbitrary — never reference CONFIG or Code.js
 * helpers at top-level evaluation time; only call them inside function bodies.
 */

function roundParser2_(n) {
  return Number((Number(n) || 0).toFixed(2));
}

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

function findHalfDiffFlipCandidates_(transactions, closingDiff) {
  const halfDiff = roundParser2_(Math.abs(closingDiff) / 2);
  if (halfDiff <= 0.02) return [];

  const suspectType = closingDiff > 0 ? "DEBIT" : "CREDIT";
  const candidates = (transactions || [])
    .filter(t => String(t && t.type || "").toUpperCase() === suspectType && (parseFloat(t && t.amount) || 0) > 0)
    .slice()
    .sort((a, b) => (Boolean(a.hasTableColumns) === Boolean(b.hasTableColumns) ? 0 : (a.hasTableColumns ? 1 : -1)));

  const singleHits = candidates.filter(t => Math.abs((parseFloat(t.amount) || 0) - halfDiff) <= 0.02);
  if (singleHits.length > 0) {
    const reason = suspectType === "DEBIT"
      ? "Suspected CREDIT read as DEBIT (amount equals half of tie-out diff)"
      : "Suspected DEBIT read as CREDIT (amount equals half of tie-out diff)";
    return singleHits.map(tx => ({ tx, reason }));
  }

  const pool = candidates.slice(0, 40);
  const pairs = [];
  for (let i = 0; i < pool.length; i++) {
    for (let j = i + 1; j < pool.length; j++) {
      const sum = roundParser2_((parseFloat(pool[i].amount) || 0) + (parseFloat(pool[j].amount) || 0));
      if (Math.abs(sum - halfDiff) <= 0.02) {
        const rank = (!pool[i].hasTableColumns && !pool[j].hasTableColumns) ? 0 : 1;
        pairs.push({ items: [pool[i], pool[j]], rank });
      }
    }
  }

  if (pairs.length > 0) {
    pairs.sort((a, b) => a.rank - b.rank);
    const reason = "Suspected flipped DEBIT↔CREDIT pair (pair sum equals half of tie-out diff)";
    return pairs[0].items.map(tx => ({ tx, reason }));
  }

  return [];
}

function validateStatementTieOut(meta, transactions) {
  const m = meta || {};
  const list = transactions || [];
  const openingBalance = roundParser2_(m.openingBalance);
  const closingBalance = roundParser2_(m.closingBalance);
  const printedTotalDebit = roundParser2_(m.totalDebit);
  const printedTotalCredit = roundParser2_(m.totalCredit);

  const counts = { DEBIT: 0, INSTALLMENT: 0, FEE: 0, CREDIT: 0, total: list.length, ambiguous: 0 };
  let sumDebits = 0;
  let sumCredits = 0;

  list.forEach(t => {
    const type = String(t && t.type || "").trim().toUpperCase();
    const amt = Math.abs(parseFloat(t && t.amount) || 0);
    if (type === "CREDIT") {
      counts.CREDIT++;
      sumCredits += amt;
    } else if (type === "INSTALLMENT") {
      counts.INSTALLMENT++;
      sumDebits += amt;
    } else if (type === "FEE") {
      counts.FEE++;
      sumDebits += amt;
    } else {
      if (type === "DEBIT") counts.DEBIT++;
      sumDebits += amt;
    }
  });

  const parsedDebits = roundParser2_(sumDebits);
  const parsedCredits = roundParser2_(sumCredits);
  const calculatedClosing = roundParser2_(openingBalance + parsedDebits - parsedCredits);

  const closingDiff = roundParser2_(calculatedClosing - closingBalance);
  const closingOk = Math.abs(closingDiff) <= 0.02;

  const debitDiff = printedTotalDebit > 0 ? roundParser2_(parsedDebits - printedTotalDebit) : 0;
  const debitOk = printedTotalDebit <= 0 || Math.abs(debitDiff) <= 0.02;

  const creditDiff = printedTotalCredit > 0 ? roundParser2_(parsedCredits - printedTotalCredit) : 0;
  const creditOk = printedTotalCredit <= 0 || Math.abs(creditDiff) <= 0.02;

  const emptyFail = list.length === 0 && (
    Math.abs(closingBalance - openingBalance) > 0.02 ||
    printedTotalDebit > 0 ||
    printedTotalCredit > 0
  );
  const mathPassed = closingOk && debitOk && creditOk && !emptyFail;

  if (!mathPassed && Math.abs(closingDiff) > 0.02) {
    const flipCandidates = findHalfDiffFlipCandidates_(list, closingDiff);
    flipCandidates.forEach(hit => {
      if (hit && hit.tx) {
        hit.tx.needsTypeConfirm = true;
        if (!hit.tx.typeConfirmReason) {
          hit.tx.typeConfirmReason = hit.reason;
        }
      }
    });
  }

  const validTypes = ["DEBIT", "CREDIT", "INSTALLMENT", "FEE"];
  const ambiguousLines = [];
  list.forEach((t, idx) => {
    const typeUpper = String(t && t.type || "").trim().toUpperCase();
    const isInvalidType = !validTypes.includes(typeUpper);
    if ((t && t.needsTypeConfirm === true) || isInvalidType) {
      if (isInvalidType && t && !t.typeConfirmReason) {
        t.typeConfirmReason = "Invalid or blank transaction Type — set DEBIT, CREDIT, INSTALLMENT, or FEE";
      }
      ambiguousLines.push({
        lineIndex: idx + 1,
        rowNum: (t && t.rowNum) || (idx + 7),
        stmtId: (t && t.stmtId) || "",
        txDate: (t && (t.txDateStr || t.txDate)) || "",
        desc: (t && t.desc) || "",
        amount: roundParser2_(t && t.amount),
        type: typeUpper || "(BLANK)",
        reason: (t && t.typeConfirmReason) || "Confirm transaction Type (DEBIT vs CREDIT)"
      });
    }
  });

  counts.ambiguous = ambiguousLines.length;

  const reasons = [];
  if (emptyFail) {
    reasons.push("0 statement lines parsed despite non-zero statement balances/totals");
  }
  if (!closingOk) {
    reasons.push(`Closing balance tie-out diff: ${closingDiff > 0 ? "+" : ""}${closingDiff.toFixed(2)} EGP`);
  }
  if (!debitOk) {
    reasons.push(`Parsed vs Printed Debit diff: ${debitDiff > 0 ? "+" : ""}${debitDiff.toFixed(2)} EGP`);
  }
  if (!creditOk) {
    reasons.push(`Parsed vs Printed Credit diff: ${creditDiff > 0 ? "+" : ""}${creditDiff.toFixed(2)} EGP`);
  }
  if (ambiguousLines.length > 0) {
    reasons.push(`${ambiguousLines.length} line(s) require Type confirmation`);
  }

  return {
    passed: mathPassed && ambiguousLines.length === 0,
    mathPassed,
    closingOk,
    debitOk,
    creditOk,
    openingBalance,
    closingBalance,
    printedTotalDebit,
    printedTotalCredit,
    parsedDebits,
    parsedCredits,
    calculatedClosing,
    closingDiff,
    debitDiff,
    creditDiff,
    counts,
    ambiguousLines,
    reasons
  };
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

  if (!text) {
    return { meta: meta, transactions: [], tieOut: validateStatementTieOut(meta, []) };
  }

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
  const strongCreditPattern = /\b(?:DIRECT\s*DEBIT\s*PAYMENT|DIRECT\s*DEBIT|CARD\s*PAYMENT|INSTAPAY\s*PAYMENT|PAYMENT\s*-\s*THANK\s*YOU|REFUND|CASHBACK|CASH\s*BACK|REVERSAL|REVERSED|CHARGEBACK|INSTAPAY|IPN\b|ACH\b|سداد|استرداد|مرتجع)\b|\bCR\b|[\d,]+\.?\d*\s*CR\b|[\d,]+\.?\d*-\b|(?:^|[\s\t])-\s*[\d,]+\.?\d*\b|\([\d,]+\.?\d*\)/i;
  const ambiguousCreditPattern = /\b(?:PAYMENT|CREDIT|DEPOSIT|SETTLEMENT|RETURN|RTN\b|REV\b|ADJUSTMENT|TRSF\b|TRANSFER|FAWRY\s*PAY|إيداع|تسوية)\b/i;

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
    let needsTypeConfirm = false;
    let typeConfirmReason = "";
    const maxAmountCap = Math.max(meta.creditLimit * 1.5, 200000);

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
      if (tableDebit > 0 && tableCredit > 0) {
        amount = Math.max(tableDebit, tableCredit);
        txType = tableDebit >= tableCredit ? "DEBIT" : "CREDIT";
        needsTypeConfirm = true;
        typeConfirmReason = "Both Debit and Credit columns non-zero in OCR table";
      } else if (tableCredit > 0 && tableDebit === 0) {
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
      const mInst = rest.match(/(\d+)\s+(?:OF|\/)\s+(\d+)/i);
      if (mInst) {
        txType = "INSTALLMENT";
        instCurrent = parseInt(mInst[1], 10);
        instTotal = parseInt(mInst[2], 10);
      } else if (feePattern.test(rawLine) || feePattern.test(rest)) {
        txType = "FEE";
      } else if (strongCreditPattern.test(rawLine) || strongCreditPattern.test(rest)) {
        txType = "CREDIT";
      } else if (ambiguousCreditPattern.test(rawLine) || ambiguousCreditPattern.test(rest)) {
        txType = "DEBIT";
        needsTypeConfirm = true;
        typeConfirmReason = "Ambiguous keyword on non-tabbed OCR line — confirm DEBIT vs CREDIT";
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
      instTotal: instTotal,
      hasTableColumns: hasTableColumns,
      needsTypeConfirm: needsTypeConfirm,
      typeConfirmReason: typeConfirmReason
    });
  }

  const tieOut = validateStatementTieOut(meta, transactions);
  meta.sumDebits = tieOut.parsedDebits;
  meta.sumCredits = tieOut.parsedCredits;

  return { meta: meta, transactions: transactions, tieOut: tieOut };
}

function applyStatementHealthBanner_(sheet, tieOut) {
  if (!sheet || !tieOut) return;
  const c = tieOut.counts || { total: 0, DEBIT: 0, INSTALLMENT: 0, FEE: 0, CREDIT: 0 };
  const bannerRange = sheet.getRange("A5:G5");
  bannerRange.merge().setFontWeight("bold").setFontSize(9);

  if (tieOut.passed) {
    bannerRange
      .setValue(
        `✅ Parse Tie-Out Passed | Opening + Debits − Credits = Closing (0.00 diff) | Lines: ${c.total} (${c.DEBIT} Debit, ${c.INSTALLMENT} Inst, ${c.FEE} Fee, ${c.CREDIT} Credit)`
      )
      .setBackground(CONFIG.COLORS.PAID)
      .setFontColor("#1b5e20");
  } else {
    bannerRange
      .setValue(
        `❌ PARSE TIE-OUT FAILED | ${(tieOut.reasons || []).join(" | ")} — Fix Type (Col E) or missing lines, then Refresh`
      )
      .setBackground(CONFIG.COLORS.OVERDUE)
      .setFontColor("#b71c1c");
  }
}

function evaluateBankStatementHealth(ss, tz, optDecisions) {
  const targetTz = tz || (ss && ss.getSpreadsheetTimeZone()) || "Africa/Cairo";
  const stmtSheet = ss && ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  if (!stmtSheet || stmtSheet.getLastRow() < 4) {
    return Object.assign(
      { hasStatement: false, hasImportMeta: false, statementDateStr: "" },
      validateStatementTieOut({ openingBalance: 0, closingBalance: 0, totalDebit: 0, totalCredit: 0 }, [])
    );
  }

  const decisions = optDecisions || loadDecisions(ss);
  const metaVals = stmtSheet.getRange(2, 1, 3, 6).getValues();
  const parsedStmtDate = parseDateValue(metaVals[0][1], targetTz);
  const stmtDateKey = makeNormDateStr(parsedStmtDate || metaVals[0][1], targetTz) || String(metaVals[0][1] || "").trim() || "ACTIVE";
  const importMeta = (decisions.importMeta && (decisions.importMeta[stmtDateKey] || decisions.importMeta["ACTIVE"])) || null;

  if (stmtSheet.getLastRow() < 7 && !importMeta) {
    return Object.assign(
      { hasStatement: false, hasImportMeta: false, statementDateStr: stmtDateKey },
      validateStatementTieOut({ openingBalance: 0, closingBalance: 0, totalDebit: 0, totalCredit: 0 }, [])
    );
  }

  const parseMetaNum = val => {
    if (typeof val === "number" && !isNaN(val)) return val;
    const s = String(val || "").trim();
    if (!s || s.startsWith("=")) return NaN;
    const n = parseFloat(s.replace(/,/g, ""));
    return isNaN(n) ? NaN : n;
  };

  const rawOpen = parseMetaNum(metaVals[1][1]);
  const rawClose = parseMetaNum(metaVals[1][3]);
  const rawDeb = parseMetaNum(metaVals[1][5]);
  const rawCred = parseMetaNum(metaVals[2][1]);

  const evalMeta = {
    statementDate: stmtDateKey,
    openingBalance: importMeta ? importMeta.printedOpening : (isNaN(rawOpen) ? 0 : rawOpen),
    closingBalance: importMeta ? importMeta.printedClosing : (isNaN(rawClose) ? 0 : rawClose),
    totalDebit: importMeta ? importMeta.printedTotalDebit : (isNaN(rawDeb) ? 0 : rawDeb),
    totalCredit: importMeta ? importMeta.printedTotalCredit : (isNaN(rawCred) ? 0 : rawCred)
  };

  const stmtLastRow = stmtSheet.getLastRow();
  const numRows = Math.max(0, stmtLastRow - 6);
  const stmtData = numRows > 0 ? stmtSheet.getRange(7, 1, numRows, 7).getValues() : [];
  const idCol = numRows > 0 ? findExistingIdColumn(stmtSheet, 6) : 0;
  const idVals = (idCol > 0 && numRows > 0) ? stmtSheet.getRange(7, idCol, numRows, 1).getValues() : null;

  const importAmbiguousSet = new Set(
    (importMeta && Array.isArray(importMeta.ambiguousIds) ? importMeta.ambiguousIds : [])
      .map(id => String(id || "").trim().toUpperCase())
      .filter(Boolean)
  );

  const rows = [];
  for (let i = 0; i < stmtData.length; i++) {
    const r = stmtData[i];
    const desc = String(r[3] || "").trim();
    const amt = parseFloat(r[5]);
    if (!desc && (isNaN(amt) || amt === 0)) continue;

    const stmtId = idVals ? String(idVals[i][0] || "").trim().toUpperCase() : "";
    const type = String(r[4] || "").trim().toUpperCase();
    const isConfirmed = Boolean(stmtId && decisions.stmtTypeConfirm && decisions.stmtTypeConfirm[stmtId]);
    const wasAmbiguous = Boolean(stmtId && importAmbiguousSet.has(stmtId) && !isConfirmed);
    const txD = parseDateValue(r[1], targetTz);

    rows.push({
      rowNum: i + 7,
      stmtId: stmtId,
      txDate: r[1],
      txDateStr: makeNormDateStr(txD || r[1], targetTz) || String(r[1] || ""),
      desc: desc,
      type: type,
      amount: Math.abs(isNaN(amt) ? 0 : amt),
      hasTableColumns: false,
      needsTypeConfirm: wasAmbiguous,
      typeConfirmReason: wasAmbiguous ? "Flagged on import for manual Type confirmation" : ""
    });
  }

  const tieOut = validateStatementTieOut(evalMeta, rows);

  if (importMeta) {
    const currentAmbiguousIds = tieOut.ambiguousLines
      .map(a => a.stmtId)
      .filter(Boolean);

    setImportMeta(decisions, importMeta.statementDate || stmtDateKey, {
      statementDate: importMeta.statementDate || stmtDateKey,
      printedOpening: evalMeta.openingBalance,
      printedClosing: evalMeta.closingBalance,
      printedTotalDebit: evalMeta.totalDebit,
      printedTotalCredit: evalMeta.totalCredit,
      counts: tieOut.counts,
      tieOutPassed: tieOut.passed,
      mathPassed: tieOut.mathPassed,
      closingDiff: tieOut.closingDiff,
      debitDiff: tieOut.debitDiff,
      creditDiff: tieOut.creditDiff,
      ambiguousIds: currentAmbiguousIds,
      reasons: tieOut.reasons,
      sourceFile: importMeta.sourceFile || "",
      note: tieOut.passed ? "PASSED" : "FAILED"
    });

    if (!optDecisions && decisions._dirty) {
      saveDecisions(ss, decisions);
    }

    applyStatementHealthBanner_(stmtSheet, tieOut);
  }

  return Object.assign(
    {
      hasStatement: true,
      hasImportMeta: Boolean(importMeta),
      statementDateStr: stmtDateKey
    },
    tieOut
  );
}

function renderStatementParseErrorSheet(ss, tz, health) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.ACTIVE_RECON);
  sheet.clear();
  sheet.setHiddenGridlines(false);

  sheet.getRange("A1:H1").merge()
    .setValue("🚨 NBE Credit Card — Statement Import Health Check FAILED (Bill Blocked)")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground("#b71c1c")
    .setFontColor("#ffffff");

  sheet.getRange("A2:H2").merge()
    .setValue("No family bill was produced. Fix Type (Col E) or missing lines on \"Bank Statement\", then click 🔄 Refresh Dashboard Now.")
    .setFontStyle("italic")
    .setFontSize(9)
    .setFontColor("#b71c1c")
    .setBackground(CONFIG.COLORS.OVERDUE);

  sheet.getRange("A4:H4").merge()
    .setValue("📊 Mathematical Tie-Out & Line Count Diagnostics")
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(CONFIG.COLORS.OVERDUE)
    .setFontColor("#b71c1c");

  const c = (health && health.counts) || { DEBIT: 0, INSTALLMENT: 0, FEE: 0, CREDIT: 0, total: 0, ambiguous: 0 };
  const diagHeaders = ["Metric", "Printed (PDF)", "Parsed (Sheet)", "Diff", "Line Type", "Count", "Status", "Notes"];
  sheet.getRange(5, 1, 1, diagHeaders.length).setValues([diagHeaders])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER);

  const diagRows = [
    [
      "Opening Balance",
      health.openingBalance || 0,
      health.openingBalance || 0,
      0,
      "DEBIT",
      c.DEBIT,
      "✅",
      "Standard purchases"
    ],
    [
      "Total Debit (D + Inst + Fee)",
      health.printedTotalDebit || 0,
      health.parsedDebits || 0,
      health.debitDiff || 0,
      "INSTALLMENT",
      c.INSTALLMENT,
      health.debitOk ? "✅ Match" : "❌ Mismatch",
      "X OF Y installment lines"
    ],
    [
      "Total Credit",
      health.printedTotalCredit || 0,
      health.parsedCredits || 0,
      health.creditDiff || 0,
      "FEE",
      c.FEE,
      health.creditOk ? "✅ Match" : "❌ Mismatch",
      "Bank fees & stamp tax"
    ],
    [
      "Closing (Open + Deb − Cred)",
      health.closingBalance || 0,
      health.calculatedClosing || 0,
      health.closingDiff || 0,
      "CREDIT / Ambiguous",
      `${c.CREDIT} / ${c.ambiguous}`,
      health.passed ? "✅ Passed" : "❌ BLOCKED",
      (health.reasons || []).join(" | ") || "Check Bank Statement Col E"
    ]
  ];

  sheet.getRange(6, 1, diagRows.length, diagHeaders.length).setValues(diagRows).setFontSize(9);
  sheet.getRange(6, 2, diagRows.length, 3).setNumberFormat("#,##0.00");

  const ambList = (health && health.ambiguousLines) || [];
  sheet.getRange("A11:H11").merge()
    .setValue(`⚠️ Confirm Type — Ambiguous / Suspected Flipped Lines (${ambList.length} Item(s))`)
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(CONFIG.COLORS.DUE_SOON)
    .setFontColor("#7f6000");

  const ambHeaders = ["Sheet Row", "S-ID", "Date", "Description", "Amount (EGP)", "Current Type", "Why Flagged", "How to Resolve"];
  sheet.getRange(12, 1, 1, ambHeaders.length).setValues([ambHeaders])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER);

  if (ambList.length > 0) {
    const ambRows = ambList.map(a => [
      a.rowNum,
      a.stmtId || "—",
      a.txDate,
      a.desc,
      a.amount,
      a.type,
      a.reason,
      "Change Col E on 'Bank Statement' (e.g. CREDIT ↔ DEBIT)"
    ]);
    sheet.getRange(13, 1, ambRows.length, ambHeaders.length).setValues(ambRows).setFontSize(9).setBackground("#fff8e1");
    sheet.getRange(13, 5, ambRows.length, 1).setNumberFormat("#,##0.00");
  } else {
    sheet.getRange("A13:H13").merge()
      .setValue("No single flipped candidate matched half-diff automatically — compare Bank Statement rows against the PDF for missing or misread amounts.")
      .setFontSize(9)
      .setBackground("#fff8e1");
  }

  return sheet;
}

function writeStatementToSheet(ss, meta, transactions, optOptions) {
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.BANK_STATEMENT);
  const opts = optOptions || {};
  const list = transactions || [];

  // Idempotent S-ID preservation by statementDate + lineIndex + amount (with reconKey/date+amt fallback)
  const newStmtDateStr = makeNormDateStr(parseDateValue(meta.statementDate, tz) || meta.statementDate, tz) || String(meta.statementDate || "").trim();
  const preservedIds = new Array(list.length).fill("");
  let preservedCount = 0;

  if (sheet.getLastRow() >= 7 && newStmtDateStr) {
    const rawOldDate = sheet.getRange(2, 2).getValue();
    const oldStmtDateStr = makeNormDateStr(parseDateValue(rawOldDate, tz) || rawOldDate, tz) || String(rawOldDate || "").trim();
    const oldIdCol = findExistingIdColumn(sheet, 6);
    if (oldStmtDateStr && oldStmtDateStr === newStmtDateStr && oldIdCol > 0) {
      const oldCount = sheet.getLastRow() - 6;
      const oldRows = sheet.getRange(7, 1, oldCount, 7).getValues();
      const oldIds = sheet.getRange(7, oldIdCol, oldCount, 1).getValues();
      const oldEntries = [];
      for (let i = 0; i < oldCount; i++) {
        const oid = String(oldIds[i][0] || "").trim();
        if (!parseRowIdNum(oid, "S")) continue;
        const amt = roundParser2_(Math.abs(parseFloat(oldRows[i][5]) || 0));
        const dVal = parseDateValue(oldRows[i][1], tz);
        const dateStr = makeNormDateStr(dVal || oldRows[i][1], tz);
        const reconKey = makeReconKey(oldRows[i][1], amt, oldRows[i][3], tz);
        oldEntries.push({ idx: i, stmtId: oid, dateStr, amount: amt, reconKey, used: false });
      }

      // Tier 1: Exact same line index + amount on the same statement date
      for (let i = 0; i < list.length; i++) {
        const newAmt = roundParser2_(Math.abs(parseFloat(list[i].amount) || 0));
        if (i < oldEntries.length && !oldEntries[i].used && Math.abs(oldEntries[i].amount - newAmt) <= 0.01) {
          preservedIds[i] = oldEntries[i].stmtId;
          oldEntries[i].used = true;
          preservedCount++;
        }
      }

      // Tier 2: Fallback by reconKey or dateStr + amount
      for (let i = 0; i < list.length; i++) {
        if (preservedIds[i]) continue;
        const t = list[i];
        const newAmt = roundParser2_(Math.abs(parseFloat(t.amount) || 0));
        const txD = parseDateValue(t.txDate, tz);
        const newDateStr = makeNormDateStr(txD || t.txDate, tz);
        const k = makeReconKey(txD || t.txDate, newAmt, t.desc, tz);

        let hit = oldEntries.find(e => !e.used && e.reconKey === k);
        if (!hit) {
          hit = oldEntries.find(e => !e.used && e.dateStr === newDateStr && Math.abs(e.amount - newAmt) <= 0.01);
        }
        if (hit) {
          preservedIds[i] = hit.stmtId;
          hit.used = true;
          preservedCount++;
        }
      }
    }
  }

  // Load decisions before tie-out so previously user-confirmed Type overrides are reapplied idempotently
  const decisions = loadDecisions(ss);
  const workingTransactions = list.map((t, idx) => {
    const copy = Object.assign({}, t);
    const pid = preservedIds[idx];
    if (pid && decisions.stmtTypeConfirm && decisions.stmtTypeConfirm[pid]) {
      copy.type = decisions.stmtTypeConfirm[pid].target;
      copy.needsTypeConfirm = false;
      copy.typeConfirmReason = "";
    }
    return copy;
  });

  const tieOut = validateStatementTieOut(meta, workingTransactions);

  sheet.clear();

  sheet.getRange("A1:G1").merge()
    .setValue("NBE Credit Card Statement — " + (meta.statementDate || "Latest"))
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(CONFIG.COLORS.PRIMARY)
    .setFontColor("#ffffff");

  const totalDebitFormula = workingTransactions.length > 0 ? "=SUMIF(E7:E, \"<>CREDIT\", F7:F)" : (meta.totalDebit || 0);
  const totalCreditFormula = workingTransactions.length > 0 ? "=SUMIF(E7:E, \"CREDIT\", F7:F)" : (meta.totalCredit || 0);
  const netCalculatedFormula = workingTransactions.length > 0 ? "=B3+F3-B4" : "";

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

  applyStatementHealthBanner_(sheet, tieOut);

  const startRow = 6;
  const headers = ["#", "Tx Date", "Posting Date", "Description", "Type", "Amount (EGP)", "Assigned Payer / Status"];
  sheet.getRange(startRow, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setBorder(true, true, true, true, false, false);
  sheet.getRange(startRow, 8).setValue(ROW_ID_HEADER);

  if (workingTransactions.length > 0) {
    const idColVals = [];
    const rows = workingTransactions.map((t, idx) => {
      const txD = parseDateValue(t.txDate, tz);
      const postD = parseDateValue(t.postDate, tz);
      let initStatus = "⏳ Checking...";
      if (t.needsTypeConfirm) initStatus = "⚠️ Confirm Type (DEBIT/CREDIT)";
      else if (t.type === "CREDIT") initStatus = "💳 Payment / Refund";
      else if (t.type === "FEE") initStatus = "🏛️ Bank Fee";

      idColVals.push([preservedIds[idx] || ""]);

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
    sheet.getRange(startRow + 1, 8, idColVals.length, 1).setValues(idColVals);
    sheet.getRange(startRow + 1, 2, rows.length, 2).setNumberFormat("dddd, MMMM d, yyyy");
    sheet.getRange(startRow + 1, 6, rows.length, 1).setNumberFormat("#,##0.00");

    const typeValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(["DEBIT", "CREDIT", "INSTALLMENT", "FEE"], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow + 1, 5, rows.length, 1).setDataValidation(typeValidation);

    for (let i = 0; i < workingTransactions.length; i++) {
      const t = workingTransactions[i];
      const rowNum = startRow + 1 + i;
      if (t.needsTypeConfirm) {
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground("#ffe0b2");
      } else if (t.type === "CREDIT") {
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground(CONFIG.COLORS.PAID);
      } else if (t.type === "INSTALLMENT") {
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground(CONFIG.COLORS.PRIMARY_LIGHT);
      } else if (t.type === "FEE") {
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground("#fffde7");
      }
    }
  }

  const idInfo = ensureSheetRowIds(
    ss,
    CONFIG.SHEETS.BANK_STATEMENT,
    6,
    8,
    "S",
    decisions,
    r => Boolean(String(r[3] || "").trim() || (!isNaN(parseFloat(r[5])) && parseFloat(r[5]) !== 0))
  );

  const assignedIds = (idInfo && idInfo.ids) ? idInfo.ids : [];
  const ambiguousIds = [];
  workingTransactions.forEach((t, idx) => {
    if (t.needsTypeConfirm && assignedIds[idx]) {
      ambiguousIds.push(assignedIds[idx]);
    }
  });

  setImportMeta(decisions, newStmtDateStr || "ACTIVE", {
    statementDate: newStmtDateStr || String(meta.statementDate || "").trim() || "ACTIVE",
    printedOpening: meta.openingBalance || 0,
    printedClosing: meta.closingBalance || 0,
    printedTotalDebit: meta.totalDebit || 0,
    printedTotalCredit: meta.totalCredit || 0,
    counts: tieOut.counts,
    tieOutPassed: tieOut.passed,
    mathPassed: tieOut.mathPassed,
    closingDiff: tieOut.closingDiff,
    debitDiff: tieOut.debitDiff,
    creditDiff: tieOut.creditDiff,
    ambiguousIds: ambiguousIds,
    reasons: tieOut.reasons,
    sourceFile: opts.sourceFile || "",
    note: tieOut.passed ? "PASSED" : "FAILED"
  });

  if (decisions._dirty) saveDecisions(ss, decisions);
  try { sheet.hideColumns(8); } catch (e) { }

  applySheetLayoutFitting(sheet, {
    bannerRow: 1,
    bannerHeight: 30,
    customRowHeights: { 2: 22, 3: 22, 4: 22, 5: 22, 6: 26 },
    explicitColWidths: { 1: 45, 2: 105, 3: 105, 4: 240, 5: 90, 6: 115, 7: 170 }
  });

  return {
    tieOut,
    statementDateStr: newStmtDateStr,
    preservedCount,
    assignedIds
  };
}

function getOrCreateFolder(parent, name) {
  const root = parent || DriveApp.getRootFolder();
  const folders = root.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : root.createFolder(name);
}

function checkAndProcessNewStatements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folderName = getEffectiveFolderName(ss);
  const rootFolder = getOrCreateFolder(null, folderName);
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

  const result = { processedCount: 0, processedFiles: [], tieOutResults: [], hasTieOutFailure: false };
  if (pdfFiles.length === 0) return result;

  pdfFiles.forEach(file => {
    const text = extractTextFromPDF(file);
    const parsed = parseNBEStatementText(text);

    if (parsed.transactions.length > 0 || parsed.meta.statementDate) {
      const writeRes = writeStatementToSheet(ss, parsed.meta, parsed.transactions, { sourceFile: file.getName() });
      result.processedCount++;
      result.processedFiles.push(file.getName());
      if (writeRes && writeRes.tieOut) {
        result.tieOutResults.push({ file: file.getName(), tieOut: writeRes.tieOut });
        if (!writeRes.tieOut.passed) result.hasTieOutFailure = true;
      }
      if (!isFromProcessed) file.moveTo(processedFolder);
    }
  });

  return result;
}

function manualProcessStatements() {
  const res = checkAndProcessNewStatements();
  if (res.processedCount > 0) {
    if (res.hasTieOutFailure) {
      const firstFail = res.tieOutResults.find(r => r.tieOut && !r.tieOut.passed);
      const reasonsText = (firstFail && firstFail.tieOut.reasons ? firstFail.tieOut.reasons.join("\n• ") : "See Bank Statement Row 5");
      SpreadsheetApp.getUi().alert(
        "Statement Tie-Out Failed — Review Required",
        `Parsed ${res.processedFiles.join(", ")}, but mathematical tie-out failed or lines need Type confirmation:\n\n• ${reasonsText}\n\nCheck 'Active Bill & Reconciliation' and fix Col E on 'Bank Statement'.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert("Statement Processed", `Successfully parsed & verified: ${res.processedFiles.join(", ")}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    updateLiveDashboard();
  } else {
    SpreadsheetApp.getUi().alert("No Statements Found", "No PDF statements found in Google Drive.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
