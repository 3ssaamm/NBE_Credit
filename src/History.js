/**
 * Statement Closing & Immutable Archive (History Sheet)
 *
 * Manages closing an audited billing cycle into an immutable append-only `History` sheet
 * containing cycle summaries, per-member breakdowns, and line-level decisions, and handles
 * one-time migration of legacy closed-month columns from `Monthly History & Forecast`.
 *
 * NOTE: Apps Script V8 load order across files is arbitrary — never reference CONFIG or Code.js
 * helpers at top-level evaluation time; only call them inside function bodies.
 */

function historyHeaders_() {
  return [
    "CycleKey",
    "CycleLabel",
    "RecordType",
    "Status",
    "Member / SubjectId",
    "Purchases / DecisionType",
    "Installments / Target",
    "Fees (EGP)",
    "Unrecorded (EGP)",
    "Refunds (EGP)",
    "Net Total (EGP)",
    "Audit / Note",
    "ClosedAt"
  ];
}

function roundHist2_(n) {
  return Number((Number(n) || 0).toFixed(2));
}

function getOrCreateHistorySheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.HISTORY);
  if (sheet.getLastRow() < 1) {
    const headers = historyHeaders_();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setFontSize(9)
      .setBackground(CONFIG.COLORS.HEADER);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readHistoryArchive(ss) {
  const sheet = ss && ss.getSheetByName(CONFIG.SHEETS.HISTORY);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const rawMap = {};

  data.forEach(r => {
    const cycleKey = String(r[0] || "").trim();
    if (!cycleKey) return;

    const cycleLabel = String(r[1] || "").trim();
    const recordType = String(r[2] || "").trim().toUpperCase();
    const status = String(r[3] || "").trim();

    if (!rawMap[cycleKey]) {
      rawMap[cycleKey] = {
        cycleKey,
        label: cycleLabel || cycleKey,
        status: status || "Closed",
        summary: {
          purchases: 0,
          installments: 0,
          fees: 0,
          unrecorded: 0,
          refunds: 0,
          netTotal: 0,
          meta: {},
          closedAt: ""
        },
        people: {},
        memberBreakdown: {},
        decisions: []
      };
    }

    const entry = rawMap[cycleKey];
    if (cycleLabel && !entry.label) entry.label = cycleLabel;
    if (status && (!entry.status || recordType === "CYCLE_SUMMARY")) entry.status = status;

    if (recordType === "CYCLE_SUMMARY") {
      let meta = {};
      try {
        const rawMeta = String(r[11] || "").trim();
        if (rawMeta) meta = JSON.parse(rawMeta);
      } catch (e) { }
      entry.summary = {
        purchases: roundHist2_(r[5]),
        installments: roundHist2_(r[6]),
        fees: roundHist2_(r[7]),
        unrecorded: roundHist2_(r[8]),
        refunds: roundHist2_(r[9]),
        netTotal: roundHist2_(r[10]),
        meta,
        closedAt: String(r[12] || "").trim()
      };
      if (cycleLabel) entry.label = cycleLabel;
      if (status) entry.status = status;
    } else if (recordType === "MEMBER_SHARE") {
      const member = String(r[4] || "").trim();
      if (member) {
        const share = {
          purchases: roundHist2_(r[5]),
          installments: roundHist2_(r[6]),
          fees: roundHist2_(r[7]),
          unrecorded: roundHist2_(r[8]),
          refunds: roundHist2_(r[9]),
          netTotal: roundHist2_(r[10])
        };
        entry.people[member] = share.netTotal;
        entry.memberBreakdown[member] = share;
      }
    } else if (recordType === "DECISION") {
      entry.decisions.push({
        subjectId: String(r[4] || "").trim(),
        type: String(r[5] || "").trim(),
        target: String(r[6] || "").trim(),
        note: String(r[11] || "").trim(),
        closedAt: String(r[12] || "").trim()
      });
    }
  });

  const sorted = {};
  Object.keys(rawMap).sort().forEach(k => {
    sorted[k] = rawMap[k];
  });
  return sorted;
}

function isCycleClosed(ss, cycleKey) {
  const key = String(cycleKey || "").trim();
  if (!key || !ss) return false;
  const archive = readHistoryArchive(ss);
  return Boolean(archive[key]);
}

function parseLegacyHistoryHeader_(headerText) {
  const text = String(headerText || "").trim();
  if (!text) return null;

  const m = text.match(
    /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})[\s\r\n]*\((Paid|Settled|Closed)\)$/i
  );
  if (!m) return null;

  const months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  const monthIdx = months.findIndex(mon => mon.toLowerCase() === m[1].toLowerCase());
  if (monthIdx < 0) return null;

  const monthTitle = months[monthIdx];
  const year = m[2];
  const mm = String(monthIdx + 1).padStart(2, "0");
  const statusRaw = m[3];
  const statusTitle = statusRaw.charAt(0).toUpperCase() + statusRaw.slice(1).toLowerCase();

  return {
    cycleKey: `${year}-${mm}`,
    cycleLabel: `${monthTitle} ${year}`,
    status: statusTitle
  };
}

function appendCycleRowsToHistory_(ss, rows, summaryIndices, decisionStartIdx) {
  if (!rows || rows.length === 0) return;
  const sheet = getOrCreateHistorySheet(ss);
  const startRow = Math.max(2, sheet.getLastRow() + 1);

  sheet.getRange(startRow, 1, rows.length, 13).setValues(rows).setFontSize(9);

  const numericCount = Math.min(rows.length, decisionStartIdx !== undefined ? decisionStartIdx : rows.length);
  if (numericCount > 0) {
    sheet.getRange(startRow, 6, numericCount, 6).setNumberFormat("#,##0.00");
  }

  (summaryIndices || [0]).forEach(idx => {
    if (idx >= 0 && idx < rows.length) {
      sheet.getRange(startRow + idx, 1, 1, 13)
        .setFontWeight("bold")
        .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);
    }
  });

  if (decisionStartIdx !== undefined && decisionStartIdx < rows.length) {
    sheet.getRange(startRow + decisionStartIdx, 1, rows.length - decisionStartIdx, 13)
      .setFontSize(8)
      .setFontColor("#666666");
  }

  try {
    sheet.protect()
      .setWarningOnly(true)
      .setDescription("Immutable closed statement archive");
  } catch (e) { }

  applySheetLayoutFitting(sheet, {
    customRowHeights: { 1: 26 },
    explicitColWidths: {
      1: 80,   // CycleKey
      2: 120,  // CycleLabel
      3: 115,  // RecordType
      4: 75,   // Status
      5: 135,  // Member / SubjectId
      6: 120,  // Purchases / DecisionType
      7: 135,  // Installments / Target
      8: 90,   // Fees
      9: 100,  // Unrecorded
      10: 90,  // Refunds
      11: 105, // Net Total
      12: 240, // Audit / Note
      13: 140  // ClosedAt
    }
  });
}

function migrateHistoryFromForecastSheet(ss, tz) {
  const targetTz = tz || (ss && ss.getSpreadsheetTimeZone()) || "Africa/Cairo";
  const decisions = loadDecisions(ss);
  if (decisions.isHistoryMigrated) {
    return { migratedMonths: 0, alreadyMigrated: true, migratedKeys: [] };
  }

  const forecastSheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY_FORECAST);
  if (!forecastSheet || forecastSheet.getLastRow() < 5) {
    decisions.isHistoryMigrated = true;
    decisions.historyMigrationNote = "No legacy forecast history present";
    saveDecisions(ss, decisions);
    return { migratedMonths: 0, alreadyMigrated: false, migratedKeys: [] };
  }

  const lastCol = forecastSheet.getLastColumn();
  if (lastCol < 2) {
    decisions.isHistoryMigrated = true;
    decisions.historyMigrationNote = "No legacy forecast columns present";
    saveDecisions(ss, decisions);
    return { migratedMonths: 0, alreadyMigrated: false, migratedKeys: [] };
  }

  const headerRow = forecastSheet.getRange(4, 1, 1, lastCol).getValues()[0];
  const legacyCols = [];
  for (let c = 1; c < headerRow.length; c++) {
    const parsed = parseLegacyHistoryHeader_(headerRow[c]);
    if (parsed) {
      legacyCols.push(Object.assign({ colIdx: c }, parsed));
    }
  }

  if (legacyCols.length === 0) {
    decisions.isHistoryMigrated = true;
    decisions.historyMigrationNote = "No legacy closed columns in Monthly History & Forecast";
    saveDecisions(ss, decisions);
    return { migratedMonths: 0, alreadyMigrated: false, migratedKeys: [] };
  }

  legacyCols.sort((a, b) => a.cycleKey.localeCompare(b.cycleKey));

  const existingArchive = readHistoryArchive(ss);
  const grid = forecastSheet.getRange(5, 1, forecastSheet.getLastRow() - 4, lastCol).getValues();
  const memberRows = [];
  for (let i = 0; i < grid.length; i++) {
    const member = String(grid[i][0] || "").trim();
    if (!member || /^TOTAL\b/i.test(member)) break;
    memberRows.push({ member, row: grid[i] });
  }

  const migratedAt = Utilities.formatDate(new Date(), targetTz, "yyyy-MM-dd HH:mm:ss");
  const migratedKeys = [];

  legacyCols.forEach(col => {
    if (existingArchive[col.cycleKey]) return;

    const entries = [];
    let cycleTotal = 0;
    memberRows.forEach(mr => {
      const val = parseFloat(mr.row[col.colIdx]);
      if (!isNaN(val) && Math.abs(val) > 0) {
        const rounded = roundHist2_(val);
        entries.push({ member: mr.member, netTotal: rounded });
        cycleTotal += rounded;
      }
    });

    if (entries.length === 0 && roundHist2_(cycleTotal) === 0) return;

    const summaryMeta = JSON.stringify({
      source: "migrated_from_forecast",
      memberCount: entries.length
    });

    const cycleRows = [
      [
        col.cycleKey,
        col.cycleLabel,
        "CYCLE_SUMMARY",
        col.status,
        "ALL",
        0,
        0,
        0,
        0,
        0,
        roundHist2_(cycleTotal),
        summaryMeta,
        migratedAt
      ]
    ];

    entries.forEach(e => {
      cycleRows.push([
        col.cycleKey,
        col.cycleLabel,
        "MEMBER_SHARE",
        col.status,
        e.member,
        0,
        0,
        0,
        0,
        0,
        e.netTotal,
        "",
        migratedAt
      ]);
    });

    appendCycleRowsToHistory_(ss, cycleRows, [0], cycleRows.length);
    existingArchive[col.cycleKey] = true;
    migratedKeys.push(col.cycleKey);
  });

  decisions.isHistoryMigrated = true;
  decisions.historyMigrationNote = `Migrated ${migratedKeys.length} legacy closed month(s) to History`;
  saveDecisions(ss, decisions);

  return { migratedMonths: migratedKeys.length, alreadyMigrated: false, migratedKeys };
}

function evaluateCloseStatementReadiness(ss, tz) {
  const targetTz = tz || (ss && ss.getSpreadsheetTimeZone()) || "Africa/Cairo";
  const arSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  if (arSheet) {
    syncActiveReconSheetToDecisions(ss, targetTz);
  }

  const parseHealth = evaluateBankStatementHealth(ss, targetTz);
  const recon = reconcileActiveStatement(ss, targetTz);
  if (recon.error) {
    return { canClose: false, blockers: [recon.error], warnings: [], parseHealth };
  }

  const ledger = buildDebtLedger(ss, targetTz, recon);
  const blockers = [];
  const warnings = [];

  if (parseHealth && parseHealth.hasImportMeta && !parseHealth.passed) {
    blockers.push("Statement parse tie-out failed — fix Bank Statement first.");
  }

  if (ledger.unresolvedCount > 0) {
    blockers.push(`Unresolved bank lines: ${ledger.unresolvedCount} (${roundHist2_(ledger.unresolvedTotal).toFixed(2)} EGP) — must be 0.`);
  }

  if (!ledger.isExactMatch) {
    const diffStr = roundHist2_(ledger.auditDifference).toFixed(2);
    blockers.push(`Statement check difference is ${ledger.auditDifference > 0 ? "+" : ""}${diffStr} EGP — must be within ±0.50 EGP.`);
  }

  if (isCycleClosed(ss, ledger.activeCycleKey)) {
    blockers.push(`Cycle ${ledger.activeCycleLabel} (${ledger.activeCycleKey}) is already closed in History.`);
  }

  if (ledger.orphanActiveCount > 0) {
    warnings.push(`Note: ${ledger.orphanActiveCount} un-dismissed orphan log(s) (${roundHist2_(ledger.orphanActiveTotal).toFixed(2)} EGP) exist.`);
  }

  let paymentDateStr = "";
  if (arSheet) {
    const h6Date = parseDateValue(arSheet.getRange(6, 8).getValue(), targetTz);
    if (h6Date && !isNaN(h6Date.getTime())) {
      paymentDateStr = Utilities.formatDate(h6Date, targetTz, "yyyy-MM-dd");
    }
  }
  const decisions = recon.decisions || loadDecisions(ss);
  if (!paymentDateStr && decisions.paymentRecords && decisions.paymentRecords[ledger.activeCycleKey] && decisions.paymentRecords[ledger.activeCycleKey].dateStr) {
    paymentDateStr = decisions.paymentRecords[ledger.activeCycleKey].dateStr;
  }

  if (!paymentDateStr) {
    warnings.push('Note: Payment Date (H6) is not filled yet — cycle will close with status "Closed" instead of "Paid".');
  }

  return {
    canClose: blockers.length === 0,
    blockers,
    warnings,
    cycleKey: ledger.activeCycleKey,
    cycleLabel: ledger.activeCycleLabel,
    paymentDateStr,
    recon,
    ledger,
    parseHealth
  };
}

function collectCycleDecisionsSnapshot_(recon, ledger) {
  const decisions = (recon && recon.decisions) || {};
  const cycleKey = (ledger && ledger.activeCycleKey) || "";

  const activeStmtIds = new Set();
  const addStmtList = arr => {
    (arr || []).forEach(it => {
      if (it && it.stmtId) activeStmtIds.add(String(it.stmtId).trim().toUpperCase());
    });
  };
  addStmtList(recon.stmtDebits);
  addStmtList(recon.stmtInstallments);
  addStmtList(recon.stmtFees);
  addStmtList(recon.stmtCredits);

  const activeInstSubjIds = new Set();
  (recon.installmentReconItems || []).forEach(it => {
    if (it && it.subjectId) activeInstSubjIds.add(String(it.subjectId).trim());
  });

  const activeOrphanSubjIds = new Set();
  (recon.orphanLogs || recon.expandedUnmatched || []).forEach(it => {
    if (it && it.subjectId) activeOrphanSubjIds.add(String(it.subjectId).trim());
  });

  const snapshot = [];

  Object.keys(decisions.stmtTypeConfirm || {}).sort().forEach(id => {
    if (!activeStmtIds.has(id.toUpperCase())) return;
    const rec = decisions.stmtTypeConfirm[id];
    if (!rec) return;
    snapshot.push({
      subjectId: id,
      type: "STMT_TYPE_CONFIRM",
      target: rec.target || "",
      note: rec.note || ""
    });
  });

  Object.keys(decisions.stmtMatch || {}).sort().forEach(id => {
    if (!activeStmtIds.has(id.toUpperCase())) return;
    const rec = decisions.stmtMatch[id];
    if (!rec) return;
    snapshot.push({
      subjectId: id,
      type: "STMT_MATCH",
      target: (rec.targetIds && rec.targetIds.join(",")) || rec.target || "",
      note: rec.note || ""
    });
  });

  Object.keys(decisions.stmtAssign || {}).sort().forEach(id => {
    if (!activeStmtIds.has(id.toUpperCase())) return;
    const rec = decisions.stmtAssign[id];
    if (!rec) return;
    snapshot.push({
      subjectId: id,
      type: "STMT_ASSIGN",
      target: rec.target || "",
      note: rec.note || ""
    });
  });

  Object.keys(decisions.stmtInstLink || {}).sort().forEach(id => {
    if (!activeStmtIds.has(id.toUpperCase())) return;
    const rec = decisions.stmtInstLink[id];
    if (!rec) return;
    snapshot.push({
      subjectId: id,
      type: "STMT_INST_LINK",
      target: (rec.targetIds && rec.targetIds.join(",")) || rec.target || "",
      note: rec.note || ""
    });
  });

  Object.keys(decisions.instOverride || {}).sort().forEach(subjId => {
    const stmtPart = String(subjId || "").split("::")[0].trim().toUpperCase();
    const matchesCycle = activeInstSubjIds.has(subjId) ||
      (cycleKey && String(subjId).startsWith(cycleKey + "::")) ||
      (stmtPart && activeStmtIds.has(stmtPart));
    if (!matchesCycle) return;
    const rec = decisions.instOverride[subjId];
    if (!rec) return;
    snapshot.push({
      subjectId: subjId,
      type: "INST_OVERRIDE",
      target: rec.target || "",
      note: rec.note || ""
    });
  });

  Object.keys(decisions.txNeglect || {}).sort().forEach(subjId => {
    if (!activeOrphanSubjIds.has(subjId)) return;
    const rec = decisions.txNeglect[subjId];
    if (!rec) return;
    snapshot.push({
      subjectId: subjId,
      type: "TX_NEGLECT",
      target: "NEGLECT",
      note: rec.note || ""
    });
  });

  (recon.stmtDebits || []).forEach(st => {
    if (!st || !st.matched || !st.stmtId) return;
    const sid = String(st.stmtId).trim().toUpperCase();
    if (decisions.stmtMatch && decisions.stmtMatch[sid]) return;
    const txIds = (st.matchedItems || []).map(tx => tx.txId).filter(Boolean).join(",");
    snapshot.push({
      subjectId: sid,
      type: "AUTO_MATCH",
      target: txIds || st.assignedPayer || "",
      note: st.matchTier || "auto"
    });
  });

  (recon.instMatches || []).forEach(m => {
    const st = m && m.stmt;
    if (!st || !st.stmtId) return;
    const sid = String(st.stmtId).trim().toUpperCase();
    if (decisions.stmtInstLink && decisions.stmtInstLink[sid]) return;
    const instIds = (m.sheetRows || []).map(r => r.instId).filter(Boolean).join(",");
    snapshot.push({
      subjectId: sid,
      type: "AUTO_INST_LINK",
      target: instIds || m.payerLabel || "",
      note: m.tier || "auto"
    });
  });

  return snapshot;
}

function closeActiveStatement(ss, tz) {
  const targetTz = tz || (ss && ss.getSpreadsheetTimeZone()) || "Africa/Cairo";
  migrateHistoryFromForecastSheet(ss, targetTz);

  const readiness = evaluateCloseStatementReadiness(ss, targetTz);
  if (!readiness.canClose) {
    return {
      closed: false,
      blockers: readiness.blockers,
      warnings: readiness.warnings,
      cycleKey: readiness.cycleKey || ""
    };
  }

  const { recon, ledger, paymentDateStr } = readiness;
  const cycleKey = ledger.activeCycleKey;
  const cycleLabel = ledger.activeCycleLabel;
  const status = paymentDateStr ? "Paid" : "Closed";
  const closedAt = Utilities.formatDate(new Date(), targetTz, "yyyy-MM-dd HH:mm:ss");
  const decisionSnapshot = collectCycleDecisionsSnapshot_(recon, ledger);

  const cb = (recon && recon.cardBal) || {};
  const auditMeta = JSON.stringify({
    statementDate: cb.statementDate ? Utilities.formatDate(cb.statementDate, targetTz, "yyyy-MM-dd") : "",
    dueDate: cb.statementDueDate ? Utilities.formatDate(cb.statementDueDate, targetTz, "yyyy-MM-dd") : "",
    openingBalance: roundHist2_(ledger.openingBalance),
    closingBill: roundHist2_(ledger.closingBill),
    familySharesTotal: roundHist2_(ledger.familySharesTotal),
    carriedOver: roundHist2_(ledger.carriedOver),
    auditDifference: roundHist2_(ledger.auditDifference),
    paymentDate: paymentDateStr || "",
    decisionCount: decisionSnapshot.length,
    source: "close_statement"
  });

  const summaryRow = [
    cycleKey,
    cycleLabel,
    "CYCLE_SUMMARY",
    status,
    "ALL",
    roundHist2_(ledger.purchasesTotal),
    roundHist2_(ledger.installmentsTotal),
    roundHist2_(ledger.feesTotal),
    roundHist2_(ledger.unrecordedTotal),
    roundHist2_(ledger.refundsTotal),
    roundHist2_(ledger.familySharesTotal),
    auditMeta,
    closedAt
  ];

  const memberRows = [];
  (ledger.personShares || []).forEach(ps => {
    const hasActivity = ps.purchases !== 0 || ps.installments !== 0 || ps.fees !== 0 ||
      ps.unrecorded !== 0 || ps.refunds !== 0 || ps.netTotal !== 0;
    if (!hasActivity) return;
    memberRows.push([
      cycleKey,
      cycleLabel,
      "MEMBER_SHARE",
      status,
      ps.person,
      roundHist2_(ps.purchases),
      roundHist2_(ps.installments),
      roundHist2_(ps.fees),
      roundHist2_(ps.unrecorded),
      roundHist2_(ps.refunds),
      roundHist2_(ps.netTotal),
      "",
      closedAt
    ]);
  });

  const decisionRows = decisionSnapshot.map(d => [
    cycleKey,
    cycleLabel,
    "DECISION",
    status,
    d.subjectId,
    d.type,
    d.target,
    "",
    "",
    "",
    "",
    d.note || "",
    closedAt
  ]);

  const rows = [summaryRow, ...memberRows, ...decisionRows];
  appendCycleRowsToHistory_(ss, rows, [0], 1 + memberRows.length);

  renderMonthlyHistoryForecastSheet(ss, targetTz, ledger);

  const arSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  if (arSheet) {
    arSheet.getRange("A1:H1").setValue(
      `💳 NBE Credit Card — Active Bill & Precision Reconciliation (${cycleLabel}) [🔒 Closed in History]`
    );
  }

  return {
    closed: true,
    cycleKey,
    cycleLabel,
    status,
    memberCount: memberRows.length,
    decisionCount: decisionSnapshot.length,
    closingBill: roundHist2_(ledger.closingBill),
    familySharesTotal: roundHist2_(ledger.familySharesTotal),
    auditDifference: roundHist2_(ledger.auditDifference),
    warnings: readiness.warnings
  };
}

function menuCloseActiveStatement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const ui = SpreadsheetApp.getUi();

  const readiness = evaluateCloseStatementReadiness(ss, tz);
  if (!readiness.canClose) {
    ui.alert(
      "Cannot Close Statement Yet",
      "The active statement cannot be archived until all requirements pass:\n\n• " + readiness.blockers.join("\n• "),
      ui.ButtonSet.OK
    );
    return;
  }

  const { ledger, cycleLabel, cycleKey, warnings } = readiness;
  const activeMembers = (ledger.personShares || []).filter(ps => ps.netTotal !== 0);
  const memberLines = activeMembers.map(ps => `  • ${ps.person}: ${roundHist2_(ps.netTotal).toFixed(2)} EGP`).join("\n");
  const warnSection = warnings.length > 0 ? `\n\nNotes:\n• ${warnings.join("\n• ")}` : "";

  const msg =
    `Archive ${cycleLabel} (${cycleKey}) into the immutable History sheet?\n\n` +
    `• Bank Closing Bill: ${roundHist2_(ledger.closingBill).toFixed(2)} EGP\n` +
    `• Family Shares Total: ${roundHist2_(ledger.familySharesTotal).toFixed(2)} EGP\n` +
    `• Statement Check: ${ledger.auditDifference >= 0 ? "+" : ""}${roundHist2_(ledger.auditDifference).toFixed(2)} EGP\n\n` +
    `Member Totals:\n${memberLines}${warnSection}\n\n` +
    `Once closed, this cycle's frozen totals cannot be overwritten.`;

  const resp = ui.alert("Close Statement to Immutable History?", msg, ui.ButtonSet.OK_CANCEL);
  if (resp !== ui.Button.OK) return;

  const res = closeActiveStatement(ss, tz);
  if (res.closed) {
    ui.alert(
      "Statement Closed to History",
      `Archived ${res.cycleLabel} (${res.status}) with ${res.memberCount} member share(s) and ${res.decisionCount} decision snapshot row(s).`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      "Close Statement Failed",
      "• " + (res.blockers || ["Unknown error"]).join("\n• "),
      ui.ButtonSet.OK
    );
  }
}
