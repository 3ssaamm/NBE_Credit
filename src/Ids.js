/**
 * Permanent Row ID Management
 * Assigns monotonic, never-reused hidden IDs (S-xxxxx, T-xxxxx, I-xxxxx)
 * to rows in Bank Statement, Transactions, and Installments without altering user columns.
 */

const ROW_ID_HEADER = "_Row_ID";

function formatRowId(prefix, num) {
  return `${prefix}-${String(num).padStart(5, "0")}`;
}

function parseRowIdNum(val, prefix) {
  const s = String(val || "").trim();
  const re = new RegExp(`^${prefix}-(\\d+)$`, "i");
  const m = s.match(re);
  return m ? parseInt(m[1], 10) : 0;
}

function findExistingIdColumn(sheet, headerRow) {
  if (!sheet || sheet.getLastRow() < headerRow) return 0;
  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  for (let c = 0; c < headers.length; c++) {
    if (String(headers[c] || "").trim() === ROW_ID_HEADER) {
      return c + 1;
    }
  }
  return 0;
}

function getOrCreateIdColumn(sheet, headerRow, defaultCol) {
  const existing = findExistingIdColumn(sheet, headerRow);
  if (existing > 0) {
    try { sheet.hideColumns(existing); } catch (e) { }
    return existing;
  }

  const maxCol = sheet.getMaxColumns();
  const lastRow = Math.max(headerRow, sheet.getLastRow());
  const numRows = lastRow - headerRow + 1;
  const lastCol = Math.max(defaultCol, sheet.getLastColumn());
  let targetCol = 0;

  const upperCheckCol = Math.min(lastCol + 1, maxCol);
  if (upperCheckCol >= defaultCol) {
    const scanWidth = upperCheckCol - defaultCol + 1;
    const grid = sheet.getRange(headerRow, defaultCol, numRows, scanWidth).getValues();
    for (let offset = 0; offset < scanWidth; offset++) {
      const colEmpty = grid.every(row => {
        const v = row[offset];
        return v === "" || v === null || v === undefined;
      });
      if (colEmpty) {
        targetCol = defaultCol + offset;
        break;
      }
    }
  }

  if (!targetCol) {
    targetCol = Math.max(defaultCol, lastCol + 1);
  }
  if (targetCol > maxCol && typeof sheet.insertColumnsAfter === "function") {
    sheet.insertColumnsAfter(maxCol, targetCol - maxCol);
  }

  sheet.getRange(headerRow, targetCol).setValue(ROW_ID_HEADER);
  try { sheet.hideColumns(targetCol); } catch (e) { }
  return targetCol;
}

function ensureSheetRowIds(ss, sheetName, headerRow, defaultCol, prefix, decisions, isDataRowFn) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < headerRow) {
    return { idCol: 0, ids: [] };
  }

  const lastRow = sheet.getLastRow();
  const numDataRows = lastRow - headerRow;
  if (numDataRows <= 0) {
    const idColOnly = getOrCreateIdColumn(sheet, headerRow, defaultCol);
    return { idCol: idColOnly, ids: [] };
  }

  const idCol = getOrCreateIdColumn(sheet, headerRow, defaultCol);
  const dataWidth = Math.max(1, defaultCol - 1);
  const dataRows = sheet.getRange(headerRow + 1, 1, numDataRows, dataWidth).getValues();
  const idVals = sheet.getRange(headerRow + 1, idCol, numDataRows, 1).getValues();

  let maxOnSheet = 0;
  for (let i = 0; i < numDataRows; i++) {
    const num = parseRowIdNum(idVals[i][0], prefix);
    if (num > maxOnSheet) maxOnSheet = num;
  }

  const storedCounter = (decisions && decisions.counters && decisions.counters[prefix]) || 0;
  let counter = Math.max(storedCounter, maxOnSheet);
  const seen = new Set();
  const ids = new Array(numDataRows);
  let sheetChanged = false;

  for (let i = 0; i < numDataRows; i++) {
    if (!isDataRowFn(dataRows[i])) {
      ids[i] = "";
      continue;
    }

    const rawId = String(idVals[i][0] || "").trim();
    const parsed = parseRowIdNum(rawId, prefix);
    if (parsed > 0) {
      const canonical = formatRowId(prefix, parsed);
      if (!seen.has(canonical)) {
        seen.add(canonical);
        ids[i] = canonical;
        if (rawId !== canonical) {
          idVals[i][0] = canonical;
          sheetChanged = true;
        }
        continue;
      }
    }

    counter++;
    const freshId = formatRowId(prefix, counter);
    seen.add(freshId);
    ids[i] = freshId;
    idVals[i][0] = freshId;
    sheetChanged = true;
  }

  if (sheetChanged) {
    sheet.getRange(headerRow + 1, idCol, numDataRows, 1).setValues(idVals);
  }
  try { sheet.hideColumns(idCol); } catch (e) { }

  if (decisions && decisions.counters && counter > (decisions.counters[prefix] || 0)) {
    decisions.counters[prefix] = counter;
    decisions._dirty = true;
  }

  return { idCol, ids };
}

function ensureAllSheetRowIds(ss, optDecisions) {
  const ownDecisions = !optDecisions;
  const decisions = optDecisions || loadDecisions(ss);

  const stmt = ensureSheetRowIds(
    ss,
    CONFIG.SHEETS.BANK_STATEMENT,
    6,
    8,
    "S",
    decisions,
    r => Boolean(String(r[3] || "").trim() || (!isNaN(parseFloat(r[5])) && parseFloat(r[5]) !== 0))
  );

  const tx = ensureSheetRowIds(
    ss,
    CONFIG.SHEETS.TRANSACTIONS,
    1,
    6,
    "T",
    decisions,
    r => Boolean(r[1] || String(r[2] || "").trim() || String(r[3] || "").trim() || parseFloat(r[4]) > 0)
  );

  const inst = ensureSheetRowIds(
    ss,
    CONFIG.SHEETS.INSTALLMENTS,
    1,
    12,
    "I",
    decisions,
    r => Boolean(String(r[2] || "").trim() || parseFloat(r[7]) > 0 || String(r[8] || "").trim())
  );

  if (ownDecisions && decisions._dirty) {
    saveDecisions(ss, decisions);
  }

  return { stmt, tx, inst, decisions };
}

function ensureSingleRowId(ss, sheet, headerRow, defaultCol, row, prefix) {
  if (!sheet || row <= headerRow) return "";
  const idCol = getOrCreateIdColumn(sheet, headerRow, defaultCol);
  const existing = String(sheet.getRange(row, idCol).getValue() || "").trim();
  const parsed = parseRowIdNum(existing, prefix);

  const lastRow = Math.max(row, sheet.getLastRow());
  const numRows = lastRow - headerRow;
  const colVals = numRows > 0 ? sheet.getRange(headerRow + 1, idCol, numRows, 1).getValues() : [];

  let maxOnCol = 0;
  let dupCount = 0;
  for (let i = 0; i < colVals.length; i++) {
    const n = parseRowIdNum(colVals[i][0], prefix);
    if (n > maxOnCol) maxOnCol = n;
    if (parsed > 0 && n === parsed) dupCount++;
  }

  if (parsed > 0 && dupCount <= 1) {
    return formatRowId(prefix, parsed);
  }

  const decisions = loadDecisions(ss);
  const nextNum = Math.max((decisions.counters && decisions.counters[prefix]) || 0, maxOnCol) + 1;
  const newId = formatRowId(prefix, nextNum);
  sheet.getRange(row, idCol).setValue(newId);
  decisions.counters[prefix] = nextNum;
  saveDecisions(ss, decisions);
  return newId;
}

function buildTxSubjectId(txId, person, totalPayers) {
  if (!txId) return "";
  return totalPayers > 1 ? `${txId}:${person}` : txId;
}

function buildInstSubjectId(stmtId, instId, cycleKey, person, totalPayers) {
  if (stmtId && !instId) return stmtId;
  if (stmtId && instId) {
    return totalPayers > 1 ? `${stmtId}::${instId}:${person}` : `${stmtId}::${instId}`;
  }
  if (!stmtId && instId) {
    const cKey = cycleKey || "ACTIVE";
    return totalPayers > 1 ? `${cKey}::${instId}:${person}` : `${cKey}::${instId}`;
  }
  return "";
}
