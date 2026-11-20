/**
 * Reconciliation Health Check Model & Panel Renderer
 *
 * Centralizes Active Bill & Reconciliation row layout coordinates and renders
 * the live Health Check panel (Rows 9-11) above Table 1 (startShareRow = 15).
 *
 * NOTE: Apps Script V8 load order across files is arbitrary — never reference CONFIG or
 * Code.js helpers at top-level evaluation time; only call them inside function bodies.
 */

function getActiveReconStartShareRow() {
  return 15;
}

function roundHealth2_(n) {
  return Number((Number(n) || 0).toFixed(2));
}

function buildHealthCheckModel(ss, tz, ledger, recon, optOptions) {
  const opts = optOptions || {};
  const targetTz = tz || (ss && ss.getSpreadsheetTimeZone()) || "Africa/Cairo";
  const safeRecon = recon || {};
  const safeLedger = ledger || {};
  const bounds = opts.bounds || getReconTableBounds(safeLedger, safeRecon);

  const refreshDate = opts.now instanceof Date && !isNaN(opts.now.getTime()) ? opts.now : new Date();
  const lastRefreshTime = Utilities.formatDate(refreshDate, targetTz, "yyyy-MM-dd HH:mm:ss");
  const codeVersion = opts.codeVersion || (typeof getCodeVersion === "function" ? getCodeVersion() : "unknown");
  const parseHealth = safeRecon.parseHealth || evaluateBankStatementHealth(ss, targetTz);

  // 1. Parse Tie-Out Snapshot
  const hasStatement = Boolean(parseHealth && parseHealth.hasStatement);
  const parsePassed = Boolean(hasStatement && parseHealth.passed);
  const parseClosingDiff = roundHealth2_((parseHealth && parseHealth.closingDiff) || 0);
  let parseLabel;
  if (!hasStatement) {
    parseLabel = "Parse Tie-Out: ❌ No Statement";
  } else if (parsePassed && Math.abs(parseClosingDiff) < 0.01) {
    parseLabel = "Parse Tie-Out: ✅ Passed";
  } else if (parsePassed) {
    parseLabel = "Parse Tie-Out: ✅ Passed (" + (parseClosingDiff > 0 ? "+" : "") + parseClosingDiff.toFixed(2) + ")";
  } else {
    parseLabel = "Parse Tie-Out: ❌ Failed (" + (parseClosingDiff > 0 ? "+" : "") + parseClosingDiff.toFixed(2) + ")";
  }

  // 2. Unresolved Bank Lines Snapshot
  let unresolvedCount = 0;
  let unresolvedRawAmt = 0;
  (safeRecon.stmtFees || []).forEach(f => {
    if (!String((f && f.assignedPayer) || "").trim()) {
      unresolvedCount++;
      unresolvedRawAmt += parseFloat(f.amount) || 0;
    }
  });
  (safeRecon.stmtCredits || []).forEach(cr => {
    if (!String((cr && cr.assignedPayer) || "").trim()) {
      unresolvedCount++;
      unresolvedRawAmt -= parseFloat(cr.amount) || 0;
    }
  });
  (safeRecon.unlinkedStmtInstallments || []).forEach(u => {
    if (!String((u && u.assignedPayer) || "").trim()) {
      unresolvedCount++;
      unresolvedRawAmt += parseFloat(u.amount) || 0;
    }
  });
  (safeRecon.missingInSheet || []).forEach(m => {
    if (!String((m && m.assignedPayer) || "").trim()) {
      unresolvedCount++;
      unresolvedRawAmt += parseFloat(m.amount) || 0;
    }
  });
  (safeRecon.installmentReconItems || []).forEach(inst => {
    const src = String((inst && inst.sourceCategory) || "");
    const act = String((inst && inst.action) || "").trim();
    if (src.startsWith("⚠️") && !act) {
      unresolvedCount++;
      unresolvedRawAmt += parseFloat(inst.amount) || 0;
    }
  });
  const unresolvedAmount = roundHealth2_(unresolvedRawAmt);

  // 3. Orphan Logs Snapshot (in-period transactions without a paired bank line, not dismissed)
  const orphanList = safeRecon.orphanLogs || safeRecon.expandedUnmatched || [];
  let orphanCount = 0;
  let orphanRawAmt = 0;
  orphanList.forEach(o => {
    if (o && !o.isNeglected) {
      orphanCount++;
      orphanRawAmt += parseFloat(o.amount) || 0;
    }
  });
  const orphanAmount = roundHealth2_(orphanRawAmt);

  // 4. Unlinked Installments Snapshot (statement installment lines not yet linked to a plan or neglected)
  const unlinkedList = safeRecon.unlinkedStmtInstallments || [];
  let unlinkedInstCount = 0;
  let unlinkedInstRawAmt = 0;
  unlinkedList.forEach(u => {
    if (u && !isNeglectedPayer(u.assignedPayer)) {
      unlinkedInstCount++;
      unlinkedInstRawAmt += parseFloat(u.amount) || 0;
    }
  });
  const unlinkedInstAmount = roundHealth2_(unlinkedInstRawAmt);

  // 5. Statement Check, Carryover, & Pending Snapshot
  const statementCheckDiff = roundHealth2_(safeLedger.auditDifference || 0);
  const statementCheckOk = Math.abs(statementCheckDiff) < 0.50;
  const rawCarryover = safeLedger.carriedOverBalance !== undefined
    ? safeLedger.carriedOverBalance
    : safeLedger.carriedOver;
  const carryoverAmount = roundHealth2_(rawCarryover || 0);
  const pendingCount = Number(safeLedger.pendingCount || 0);
  const pendingTotal = roundHealth2_(safeLedger.pendingTotal || 0);

  // 6. Overall Health Status
  let overallStatus;
  if (!parsePassed) {
    overallStatus = "ERROR";
  } else if (unresolvedCount === 0 && orphanCount === 0 && unlinkedInstCount === 0 && statementCheckOk) {
    overallStatus = "HEALTHY";
  } else {
    overallStatus = "WARNING";
  }
  const allHealthy = overallStatus === "HEALTHY";

  // 7. Live Reactive Formula Expressions Bound to Sheet Tables
  const numExceptions = bounds.numExceptions || 0;
  const startExRow = bounds.startExRow;
  const endExRow = bounds.endExRow;
  const hasTable4 = Boolean(bounds.hasTable4);
  const startInstRow = bounds.startInstRow;
  const endInstRow = bounds.endInstRow;
  const hasOrphanTable = Boolean(bounds.hasOrphanTable);
  const startOrphanRow = bounds.startOrphanRow;
  const endOrphanRow = bounds.endOrphanRow;

  const t2UnresolvedCount = numExceptions > 0 ? `COUNTIFS(F${startExRow}:F${endExRow}, "")` : "0";
  const t2UnresolvedAmt = numExceptions > 0
    ? `(SUMIFS(D${startExRow}:D${endExRow}, F${startExRow}:F${endExRow}, "", E${startExRow}:E${endExRow}, "<>*Credit*") - SUMIFS(D${startExRow}:D${endExRow}, F${startExRow}:F${endExRow}, "", E${startExRow}:E${endExRow}, "*Credit*"))`
    : "0";
  const t4UnresolvedCount = hasTable4 ? `COUNTIFS(E${startInstRow}:E${endInstRow}, "⚠️*", F${startInstRow}:F${endInstRow}, "")` : "0";
  const t4UnresolvedAmt = hasTable4 ? `SUMIFS(D${startInstRow}:D${endInstRow}, E${startInstRow}:E${endInstRow}, "⚠️*", F${startInstRow}:F${endInstRow}, "")` : "0";

  const unresolvedCountExpr = hasTable4 ? `(${t2UnresolvedCount} + ${t4UnresolvedCount})` : t2UnresolvedCount;
  const unresolvedAmtExpr = hasTable4 ? `(${t2UnresolvedAmt} + ${t4UnresolvedAmt})` : t2UnresolvedAmt;

  const negPayExpr = numExceptions > 0
    ? `SUMIFS(D${startExRow}:D${endExRow}, E${startExRow}:E${endExRow}, "Bill Payment (Credit)", F${startExRow}:F${endExRow}, "🚫*")`
    : "0";
  const carriedOverExpr = `(${(safeLedger.openingBalance || 0).toFixed(2)} - ${negPayExpr})`;

  const orphanActiveCountExpr = hasOrphanTable ? `COUNTIFS(F${startOrphanRow}:F${endOrphanRow}, "<>🚫*")` : "0";
  const orphanActiveAmtExpr = hasOrphanTable ? `SUMIFS(D${startOrphanRow}:D${endOrphanRow}, F${startOrphanRow}:F${endOrphanRow}, "<>🚫*")` : "0";

  // Subtract ignored/linking rows explicitly instead of "<>🚫*" so blank F cells remain included
  const unlinkedInstCountExpr = numExceptions > 0
    ? `(COUNTIF(E${startExRow}:E${endExRow}, "Installment*") - COUNTIFS(E${startExRow}:E${endExRow}, "Installment*", F${startExRow}:F${endExRow}, "🚫*") - COUNTIFS(E${startExRow}:E${endExRow}, "Installment*", F${startExRow}:F${endExRow}, "🔗*"))`
    : "0";
  const unlinkedInstAmtExpr = numExceptions > 0
    ? `(SUMIF(E${startExRow}:E${endExRow}, "Installment*", D${startExRow}:D${endExRow}) - SUMIFS(D${startExRow}:D${endExRow}, E${startExRow}:E${endExRow}, "Installment*", F${startExRow}:F${endExRow}, "🚫*") - SUMIFS(D${startExRow}:D${endExRow}, E${startExRow}:E${endExRow}, "Installment*", F${startExRow}:F${endExRow}, "🔗*"))`
    : "0";

  const headerFormula = !parsePassed
    ? `="🩺 Reconciliation Health Check — ❌ Parse Tie-Out Failed"`
    : `=IF(AND(${unresolvedCountExpr}=0, ${orphanActiveCountExpr}=0, ${unlinkedInstCountExpr}=0, ABS(F6)<0.5), "🩺 Reconciliation Health Check — ✅ All Checks Passed", "🩺 Reconciliation Health Check — ⚠️ " & (IF(${unresolvedCountExpr}>0,1,0) + IF(${orphanActiveCountExpr}>0,1,0) + IF(${unlinkedInstCountExpr}>0,1,0) + IF(ABS(F6)>=0.5,1,0)) & " Check(s) Need Attention")`;

  const unresolvedFormula = `=IF(${unresolvedCountExpr}=0, "Unresolved Bank Lines: ✅ 0 (0.00 EGP)", "Unresolved Bank Lines: ⚠️ " & ${unresolvedCountExpr} & " (" & TEXT(${unresolvedAmtExpr},"#,##0.00") & " EGP)")`;
  const orphanFormula = `=IF(${orphanActiveCountExpr}=0, "Orphan Logs: ✅ 0 (0.00 EGP)", "Orphan Logs: ⚠️ " & ${orphanActiveCountExpr} & " (" & TEXT(${orphanActiveAmtExpr},"#,##0.00") & " EGP)")`;
  const unlinkedInstFormula = `=IF(${unlinkedInstCountExpr}=0, "Unlinked Installments: ✅ 0 (0.00 EGP)", "Unlinked Installments: ⚠️ " & ${unlinkedInstCountExpr} & " (" & TEXT(${unlinkedInstAmtExpr},"#,##0.00") & " EGP)")`;
  const stmtCheckFormula = `=IF(ABS(F6)<0.5, "Stmt Check: ✅ " & TEXT(F6,"+#,##0.00;-#,##0.00;0.00") & " EGP", "Stmt Check: ⚠️ " & TEXT(F6,"+#,##0.00;-#,##0.00;0.00") & " EGP")`;
  const carryoverPendingFormula = `="Carryover: " & TEXT(${carriedOverExpr},"+#,##0.00;-#,##0.00;0.00") & " | Pending: ${pendingCount} (" & TEXT(${pendingTotal.toFixed(2)},"#,##0.00") & " EGP)"`;
  const lastRefreshLabel = "Last Refresh: " + lastRefreshTime;
  const codeVersionLabel = "Code Version: " + codeVersion;

  return {
    hasStatement,
    parsePassed,
    parseClosingDiff,
    parseLabel,
    parseHealth,
    unresolvedCount,
    unresolvedAmount,
    orphanCount,
    orphanAmount,
    unlinkedInstCount,
    unlinkedInstAmount,
    statementCheckDiff,
    statementCheckOk,
    carryoverAmount,
    pendingCount,
    pendingTotal,
    lastRefreshTime,
    codeVersion,
    lastRefreshLabel,
    codeVersionLabel,
    overallStatus,
    allHealthy,
    expressions: {
      t2UnresolvedCount,
      t2UnresolvedAmt,
      t4UnresolvedCount,
      t4UnresolvedAmt,
      unresolvedCountExpr,
      unresolvedAmtExpr,
      negPayExpr,
      carriedOverExpr,
      orphanActiveCountExpr,
      orphanActiveAmtExpr,
      unlinkedInstCountExpr,
      unlinkedInstAmtExpr
    },
    formulas: {
      headerFormula,
      unresolvedFormula,
      orphanFormula,
      unlinkedInstFormula,
      stmtCheckFormula,
      carryoverPendingFormula
    }
  };
}

function renderHealthCheckPanel(sheet, healthModel) {
  if (!sheet || !healthModel) return;

  const isError = healthModel.overallStatus === "ERROR";
  const isHealthy = healthModel.overallStatus === "HEALTHY";

  const headerBg = isError ? "#fce8e6" : (isHealthy ? "#e8f5e9" : "#fff2cc");
  const headerFg = isError ? "#b71c1c" : (isHealthy ? "#1b5e20" : "#7f6000");

  // Row 9: Health Check Panel Header Bar
  sheet.getRange("A9:H9").merge()
    .setFormula(healthModel.formulas.headerFormula)
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(headerBg)
    .setFontColor(headerFg);

  // Row 10: Parse Tie-Out, Unresolved Bank Lines, Orphan Logs, Unlinked Installments
  sheet.getRange("A10:B10").merge()
    .setValue(healthModel.parseLabel)
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(healthModel.parsePassed ? "#e8f5e9" : "#fce8e6")
    .setFontColor(healthModel.parsePassed ? "#1b5e20" : "#b71c1c");

  sheet.getRange("C10:D10").merge()
    .setFormula(healthModel.formulas.unresolvedFormula)
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(healthModel.unresolvedCount === 0 ? "#e8f5e9" : "#fff8e1")
    .setFontColor(healthModel.unresolvedCount === 0 ? "#1b5e20" : "#7f6000");

  sheet.getRange("E10:F10").merge()
    .setFormula(healthModel.formulas.orphanFormula)
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(healthModel.orphanCount === 0 ? "#e8f5e9" : "#fff8e1")
    .setFontColor(healthModel.orphanCount === 0 ? "#1b5e20" : "#7f6000");

  sheet.getRange("G10:H10").merge()
    .setFormula(healthModel.formulas.unlinkedInstFormula)
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(healthModel.unlinkedInstCount === 0 ? "#e8f5e9" : "#fff8e1")
    .setFontColor(healthModel.unlinkedInstCount === 0 ? "#1b5e20" : "#7f6000");

  // Row 11: Statement Check, Carryover & Pending, Last Refresh, Code Version
  sheet.getRange("A11:B11").merge()
    .setFormula(healthModel.formulas.stmtCheckFormula)
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(healthModel.statementCheckOk ? "#e8f5e9" : "#fff8e1")
    .setFontColor(healthModel.statementCheckOk ? "#1b5e20" : "#7f6000");

  sheet.getRange("C11:D11").merge()
    .setFormula(healthModel.formulas.carryoverPendingFormula)
    .setFontStyle("italic")
    .setFontSize(9)
    .setBackground("#f8f9fa")
    .setFontColor("#555555");

  sheet.getRange("E11:F11").merge()
    .setValue(healthModel.lastRefreshLabel)
    .setFontSize(9)
    .setBackground("#f8f9fa")
    .setFontColor("#555555");

  sheet.getRange("G11:H11").merge()
    .setValue(healthModel.codeVersionLabel)
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground("#f8f9fa")
    .setFontColor("#333333");

  sheet.getRange("A9:H11").setBorder(true, true, true, true, true, true);

  // Live Conditional Formatting Rules (using ISNUMBER(SEARCH(...)) — safe for UTF-16 emojis)
  try {
    const headerRange = sheet.getRange("A9:H9");
    const row10TilesRange = sheet.getRange("C10:H10");
    const stmtCheckRange = sheet.getRange("A11:B11");

    const rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISNUMBER(SEARCH("✅", $A$9))')
        .setBackground("#e8f5e9")
        .setFontColor("#1b5e20")
        .setBold(true)
        .setRanges([headerRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISNUMBER(SEARCH("❌", $A$9))')
        .setBackground("#fce8e6")
        .setFontColor("#b71c1c")
        .setBold(true)
        .setRanges([headerRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISNUMBER(SEARCH("⚠️", $A$9))')
        .setBackground("#fff2cc")
        .setFontColor("#7f6000")
        .setBold(true)
        .setRanges([headerRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISNUMBER(SEARCH("✅", C10))')
        .setBackground("#e8f5e9")
        .setFontColor("#1b5e20")
        .setBold(true)
        .setRanges([row10TilesRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISNUMBER(SEARCH("⚠️", C10))')
        .setBackground("#fff8e1")
        .setFontColor("#7f6000")
        .setBold(true)
        .setRanges([row10TilesRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISNUMBER(SEARCH("✅", A11))')
        .setBackground("#e8f5e9")
        .setFontColor("#1b5e20")
        .setBold(true)
        .setRanges([stmtCheckRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISNUMBER(SEARCH("⚠️", A11))')
        .setBackground("#fff8e1")
        .setFontColor("#7f6000")
        .setBold(true)
        .setRanges([stmtCheckRange])
        .build()
    ];
    sheet.setConditionalFormatRules(rules);
  } catch (e) { }
}
