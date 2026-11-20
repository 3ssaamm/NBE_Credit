/**
 * Decision Store v2 (_Decisions) & One-Time Legacy Migration
 * Stores all user & confirmed reconciliation decisions explicitly by permanent row ID
 * instead of fragile occurrence-counted or positional keys.
 */

const DECISIONS_HEADERS = ["Type", "SubjectId", "Target", "Note", "Source", "Locked", "UpdatedAt"];

function nowCairoTimestamp_() {
  return Utilities.formatDate(new Date(), "Africa/Cairo", "yyyy-MM-dd HH:mm:ss");
}

function getDecisionsSheet(ss) {
  const name = CONFIG.SHEETS.DECISIONS;
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    try { sheet.hideSheet(); } catch (e) { }
    sheet.getRange(1, 1, 1, DECISIONS_HEADERS.length).setValues([DECISIONS_HEADERS]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

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

function loadDecisions(ss) {
  const store = {
    stmtAssign: {},
    stmtMatch: {},
    txNeglect: {},
    instOverride: {},
    instLinks: {},
    stmtInstLink: {},
    instProgress: {},
    importMeta: {},
    stmtTypeConfirm: {},
    merchantAliases: {},
    paymentRecords: {},
    counters: { S: 0, T: 0, I: 0 },
    isMigrated: false,
    migrationNote: "",
    isHistoryMigrated: false,
    historyMigrationNote: "",
    _dirty: false
  };

  try {
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DECISIONS);
    if (!sheet || sheet.getLastRow() < 2) return store;

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    data.forEach(r => {
      const type = String(r[0] || "").trim();
      const subjectId = String(r[1] || "").trim();
      const target = String(r[2] || "").trim();
      const note = String(r[3] || "").trim();
      const source = String(r[4] || "user").trim();
      const locked = r[5] === true || String(r[5]).toUpperCase() === "TRUE";
      const updatedAt = String(r[6] || "").trim();

      if (!type && !subjectId) return;

      if (type === "META_MIGRATION" && subjectId === "V1_RECON_STORE") {
        store.isMigrated = target.toUpperCase() === "TRUE";
        store.migrationNote = note;
      } else if (type === "META_MIGRATION" && subjectId === "V1_HISTORY_ARCHIVE") {
        store.isHistoryMigrated = target.toUpperCase() === "TRUE";
        store.historyMigrationNote = note;
      } else if (type === "META_COUNTER" && (subjectId === "S" || subjectId === "T" || subjectId === "I")) {
        store.counters[subjectId] = parseInt(target, 10) || 0;
      } else if (type === "STMT_ASSIGN" && subjectId) {
        const isNeg = target === "NEGLECT" || isNeglectedPayer(target);
        store.stmtAssign[subjectId] = {
          subjectId,
          target: isNeg ? "NEGLECT" : target,
          note,
          source,
          locked,
          isNeglected: isNeg,
          updatedAt
        };
      } else if (type === "STMT_MATCH" && subjectId) {
        const targetIds = target.split(",").map(s => s.trim()).filter(Boolean);
        store.stmtMatch[subjectId] = {
          subjectId,
          targetIds,
          target: targetIds.join(","),
          note,
          source,
          locked: locked || true,
          updatedAt
        };
      } else if (type === "TX_NEGLECT" && subjectId) {
        store.txNeglect[subjectId] = {
          subjectId,
          note,
          source,
          locked,
          updatedAt
        };
      } else if (type === "INST_OVERRIDE" && subjectId) {
        store.instOverride[subjectId] = {
          subjectId,
          target,
          note,
          source,
          locked,
          isNeglected: isNeglectedPayer(target),
          updatedAt
        };
      } else if (type === "INST_LINK" && subjectId && target) {
        const planKey = subjectId.toLowerCase();
        const instIds = target.split(",").map(s => s.trim().toUpperCase()).filter(Boolean);
        store.instLinks[planKey] = {
          planKey,
          instIds,
          target: instIds.join(","),
          note,
          source,
          locked: locked || true,
          updatedAt
        };
      } else if (type === "STMT_INST_LINK" && subjectId && target) {
        const targetIds = target.split(",").map(s => s.trim().toUpperCase()).filter(Boolean);
        store.stmtInstLink[subjectId] = {
          subjectId,
          targetIds,
          target: targetIds.join(","),
          note,
          source,
          locked: locked || true,
          updatedAt
        };
      } else if (type === "MERCHANT_ALIAS" && subjectId && target) {
        const stmtKey = subjectId.toLowerCase();
        const txKey = target.toLowerCase();
        store.merchantAliases[`${stmtKey}::${txKey}`] = {
          stmtKey,
          txKey,
          note,
          source,
          locked: locked || true,
          updatedAt
        };
      } else if (type === "CYCLE_PAYMENT" && subjectId) {
        let meta = {};
        try {
          if (target) meta = JSON.parse(target);
        } catch (e) { }
        store.paymentRecords[subjectId] = {
          cycleKey: subjectId,
          dateStr: meta.dateStr || "",
          cycleLabel: meta.cycleLabel || "",
          amount: parseFloat(meta.amount) || 0,
          status: meta.status || "Settled",
          isApplied: locked,
          updatedAt
        };
      } else if (type === "INST_PROGRESS" && subjectId) {
        let meta = {};
        try {
          if (target) meta = JSON.parse(target);
        } catch (e) { }
        const id = subjectId.toUpperCase();
        const rawPaid = parseInt(meta.paid, 10);
        const rawTotal = parseInt(meta.total, 10);
        const rawBilled = parseInt(meta.billed, 10);
        const paid = isNaN(rawPaid) ? 0 : rawPaid;
        const total = isNaN(rawTotal) ? 0 : rawTotal;
        const billed = isNaN(rawBilled) ? paid : rawBilled;
        store.instProgress[id] = {
          instId: id,
          paid,
          total,
          billed,
          cycle: String(meta.cycle || "").trim(),
          completed: meta.completed === true,
          note,
          source,
          locked,
          updatedAt
        };
      } else if (type === "META_IMPORT" && subjectId) {
        let meta = {};
        try {
          if (target) meta = JSON.parse(target);
        } catch (e) { }
        store.importMeta[subjectId] = Object.assign(
          { statementDate: subjectId, ambiguousIds: [], reasons: [] },
          meta,
          { note, source, locked, updatedAt }
        );
      } else if (type === "STMT_TYPE_CONFIRM" && subjectId && target) {
        const id = subjectId.toUpperCase();
        const txType = target.toUpperCase();
        store.stmtTypeConfirm[id] = {
          subjectId: id,
          target: txType,
          note,
          source,
          locked: true,
          updatedAt
        };
      }
    });
  } catch (err) {
    Logger.log("[loadDecisions] Error reading _Decisions: " + err);
  }

  return store;
}

function saveDecisions(ss, decisions) {
  try {
    const sheet = getDecisionsSheet(ss);
    const nowStr = nowCairoTimestamp_();
    const rows = [];

    if (decisions.isMigrated) {
      rows.push(["META_MIGRATION", "V1_RECON_STORE", "TRUE", decisions.migrationNote || "", "system", true, nowStr]);
    }

    if (decisions.isHistoryMigrated) {
      rows.push(["META_MIGRATION", "V1_HISTORY_ARCHIVE", "TRUE", decisions.historyMigrationNote || "", "system", true, nowStr]);
    }

    ["S", "T", "I"].forEach(prefix => {
      const val = (decisions.counters && decisions.counters[prefix]) || 0;
      if (val > 0) {
        rows.push(["META_COUNTER", prefix, String(val), `${prefix} monotonic sequence`, "system", true, nowStr]);
      }
    });

    Object.keys(decisions.paymentRecords || {}).sort().forEach(cycleKey => {
      const pr = decisions.paymentRecords[cycleKey];
      if (!pr || !pr.dateStr) return;
      const payload = JSON.stringify({
        dateStr: pr.dateStr,
        cycleLabel: pr.cycleLabel || "",
        amount: pr.amount || 0,
        status: pr.status || "Settled"
      });
      rows.push(["CYCLE_PAYMENT", cycleKey, payload, "[]", "user", pr.isApplied === true, pr.updatedAt || nowStr]);
    });

    Object.keys(decisions.stmtMatch || {}).sort().forEach(subjectId => {
      const m = decisions.stmtMatch[subjectId];
      if (!m || !m.targetIds || m.targetIds.length === 0) return;
      rows.push(["STMT_MATCH", subjectId, m.targetIds.join(","), m.note || "", m.source || "user", m.locked !== false, m.updatedAt || nowStr]);
    });

    Object.keys(decisions.stmtAssign || {}).sort().forEach(subjectId => {
      const a = decisions.stmtAssign[subjectId];
      if (!a || !a.target) return;
      rows.push(["STMT_ASSIGN", subjectId, a.target, a.note || "", a.source || "user", Boolean(a.locked), a.updatedAt || nowStr]);
    });

    Object.keys(decisions.txNeglect || {}).sort().forEach(subjectId => {
      const t = decisions.txNeglect[subjectId];
      if (!t) return;
      rows.push(["TX_NEGLECT", subjectId, "NEGLECT", t.note || "", t.source || "user", Boolean(t.locked), t.updatedAt || nowStr]);
    });

    Object.keys(decisions.instOverride || {}).sort().forEach(subjectId => {
      const o = decisions.instOverride[subjectId];
      if (!o || !o.target) return;
      rows.push(["INST_OVERRIDE", subjectId, o.target, o.note || "", o.source || "user", Boolean(o.locked), o.updatedAt || nowStr]);
    });

    Object.keys(decisions.stmtInstLink || {}).sort().forEach(subjectId => {
      const l = decisions.stmtInstLink[subjectId];
      if (!l || !l.targetIds || l.targetIds.length === 0) return;
      rows.push(["STMT_INST_LINK", subjectId, l.targetIds.join(","), l.note || "", l.source || "user", l.locked !== false, l.updatedAt || nowStr]);
    });

    Object.keys(decisions.instLinks || {}).sort().forEach(planKey => {
      const l = decisions.instLinks[planKey];
      if (!l || !l.instIds || l.instIds.length === 0) return;
      rows.push(["INST_LINK", l.planKey || planKey, l.instIds.join(","), l.note || "", l.source || "user", l.locked !== false, l.updatedAt || nowStr]);
    });

    Object.keys(decisions.merchantAliases || {}).sort().forEach(key => {
      const a = decisions.merchantAliases[key];
      if (!a || !a.stmtKey || !a.txKey) return;
      rows.push(["MERCHANT_ALIAS", a.stmtKey, a.txKey, a.note || "", a.source || "user", a.locked !== false, a.updatedAt || nowStr]);
    });

    Object.keys(decisions.instProgress || {}).sort().forEach(instId => {
      const p = decisions.instProgress[instId];
      if (!p) return;
      const paid = parseInt(p.paid, 10) || 0;
      const total = parseInt(p.total, 10) || 0;
      const billed = p.billed !== undefined && p.billed !== null ? (parseInt(p.billed, 10) || 0) : paid;
      const payload = JSON.stringify({
        paid,
        total,
        billed,
        cycle: String(p.cycle || "").trim(),
        completed: p.completed === true
      });
      rows.push(["INST_PROGRESS", instId, payload, p.note || "", p.source || "system", p.locked !== false, p.updatedAt || nowStr]);
    });

    Object.keys(decisions.stmtTypeConfirm || {}).sort().forEach(subjectId => {
      const c = decisions.stmtTypeConfirm[subjectId];
      if (!c || !c.target) return;
      rows.push(["STMT_TYPE_CONFIRM", subjectId, c.target, c.note || "", c.source || "user", true, c.updatedAt || nowStr]);
    });

    Object.keys(decisions.importMeta || {}).sort().forEach(stmtDateKey => {
      const m = decisions.importMeta[stmtDateKey];
      if (!m) return;
      const payload = JSON.stringify(buildImportMetaPayload_(m, stmtDateKey));
      rows.push(["META_IMPORT", stmtDateKey, payload, m.note || "", m.source || "parser", m.tieOutPassed === true, m.updatedAt || nowStr]);
    });

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
    }
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 7).setValues(rows);
    }
    decisions._dirty = false;
  } catch (err) {
    Logger.log("[saveDecisions] Error writing _Decisions: " + err);
  }
}

function setStmtAssign(decisions, subjectId, target, note, source) {
  if (!subjectId || !target) return;
  const isNeg = target === "NEGLECT" || isNeglectedPayer(target);
  decisions.stmtAssign[subjectId] = {
    subjectId,
    target: isNeg ? "NEGLECT" : target,
    note: note || "",
    source: source || "user",
    locked: true,
    isNeglected: isNeg,
    updatedAt: nowCairoTimestamp_()
  };
  decisions._dirty = true;
}

function clearStmtAssign(decisions, subjectId) {
  if (decisions.stmtAssign && decisions.stmtAssign[subjectId]) {
    delete decisions.stmtAssign[subjectId];
    decisions._dirty = true;
  }
}

function setStmtMatch(decisions, stmtId, txIds, note, source) {
  if (!stmtId || !txIds || txIds.length === 0) return;
  decisions.stmtMatch[stmtId] = {
    subjectId: stmtId,
    targetIds: txIds.slice(),
    target: txIds.join(","),
    note: note || "",
    source: source || "user",
    locked: true,
    updatedAt: nowCairoTimestamp_()
  };
  decisions._dirty = true;
}

function clearStmtMatch(decisions, stmtId) {
  if (decisions.stmtMatch && decisions.stmtMatch[stmtId]) {
    delete decisions.stmtMatch[stmtId];
    decisions._dirty = true;
  }
}

function setTxNeglect(decisions, subjectId, note, source) {
  if (!subjectId) return;
  decisions.txNeglect[subjectId] = {
    subjectId,
    note: note || "",
    source: source || "user",
    locked: true,
    updatedAt: nowCairoTimestamp_()
  };
  decisions._dirty = true;
}

function clearTxNeglect(decisions, subjectId) {
  if (decisions.txNeglect && decisions.txNeglect[subjectId]) {
    delete decisions.txNeglect[subjectId];
    decisions._dirty = true;
  }
}

function setInstOverride(decisions, subjectId, action, note, source) {
  if (!subjectId || !action) return;
  decisions.instOverride[subjectId] = {
    subjectId,
    target: action,
    note: note || "",
    source: source || "user",
    locked: true,
    isNeglected: isNeglectedPayer(action),
    updatedAt: nowCairoTimestamp_()
  };
  decisions._dirty = true;
}

function clearInstOverride(decisions, subjectId) {
  if (decisions.instOverride && decisions.instOverride[subjectId]) {
    delete decisions.instOverride[subjectId];
    decisions._dirty = true;
  }
}

function setCyclePayment(decisions, cycleKey, record) {
  if (!cycleKey || !record || !record.dateStr) return;
  decisions.paymentRecords[cycleKey] = Object.assign({}, record, {
    cycleKey,
    updatedAt: nowCairoTimestamp_()
  });
  decisions._dirty = true;
}

function clearCyclePayment(decisions, cycleKey) {
  if (decisions.paymentRecords && decisions.paymentRecords[cycleKey]) {
    delete decisions.paymentRecords[cycleKey];
    decisions._dirty = true;
  }
}

function setMerchantAlias(decisions, stmtKey, txKey, note, source) {
  if (!decisions) return;
  if (!decisions.merchantAliases) decisions.merchantAliases = {};
  const k1 = String(stmtKey || "").toLowerCase().trim();
  const k2 = String(txKey || "").toLowerCase().trim();
  if (!k1 || !k2 || k1 === k2) return;
  decisions.merchantAliases[`${k1}::${k2}`] = {
    stmtKey: k1,
    txKey: k2,
    note: note || "",
    source: source || "user",
    locked: true,
    updatedAt: nowCairoTimestamp_()
  };
  decisions._dirty = true;
}

function clearMerchantAlias(decisions, stmtKey, txKey) {
  if (!decisions || !decisions.merchantAliases) return;
  const k1 = String(stmtKey || "").toLowerCase().trim();
  const k2 = String(txKey || "").toLowerCase().trim();
  const mapKey = `${k1}::${k2}`;
  if (decisions.merchantAliases[mapKey]) {
    delete decisions.merchantAliases[mapKey];
    decisions._dirty = true;
  }
}

function setInstLink(decisions, planKey, instIds, note, source) {
  if (!decisions) return;
  if (!decisions.instLinks) decisions.instLinks = {};
  const pk = String(planKey || "").toLowerCase().trim();
  const cleanIds = (instIds || []).map(id => String(id || "").trim().toUpperCase()).filter(Boolean);
  if (!pk || cleanIds.length === 0) return;
  decisions.instLinks[pk] = {
    planKey: pk,
    instIds: cleanIds,
    target: cleanIds.join(","),
    note: note || "",
    source: source || "user",
    locked: true,
    updatedAt: nowCairoTimestamp_()
  };
  decisions._dirty = true;
}

function clearInstLink(decisions, planKey) {
  if (!decisions || !decisions.instLinks) return;
  const pk = String(planKey || "").toLowerCase().trim();
  if (decisions.instLinks[pk]) {
    delete decisions.instLinks[pk];
    decisions._dirty = true;
  }
}

function setStmtInstLink(decisions, stmtId, instIds, note, source) {
  if (!decisions || !stmtId) return;
  if (!decisions.stmtInstLink) decisions.stmtInstLink = {};
  const cleanIds = (instIds || []).map(id => String(id || "").trim().toUpperCase()).filter(Boolean);
  if (cleanIds.length === 0) return;
  decisions.stmtInstLink[stmtId] = {
    subjectId: stmtId,
    targetIds: cleanIds,
    target: cleanIds.join(","),
    note: note || "",
    source: source || "user",
    locked: true,
    updatedAt: nowCairoTimestamp_()
  };
  decisions._dirty = true;
}

function clearStmtInstLink(decisions, stmtId) {
  if (decisions && decisions.stmtInstLink && decisions.stmtInstLink[stmtId]) {
    delete decisions.stmtInstLink[stmtId];
    decisions._dirty = true;
  }
}

function setInstProgress(decisions, instId, entry) {
  const id = String(instId || "").trim().toUpperCase();
  if (!decisions || !id || !entry) return;
  if (!decisions.instProgress) decisions.instProgress = {};

  const rawPaid = parseInt(entry.paid, 10);
  const rawTotal = parseInt(entry.total, 10);
  const rawBilled = (entry.billed !== undefined && entry.billed !== null) ? parseInt(entry.billed, 10) : rawPaid;
  const paid = isNaN(rawPaid) ? 0 : rawPaid;
  const total = isNaN(rawTotal) ? 0 : rawTotal;
  const billed = isNaN(rawBilled) ? paid : rawBilled;
  const cycle = String(entry.cycle || "").trim();
  const completed = Boolean(entry.completed);

  const existing = decisions.instProgress[id];
  if (
    existing &&
    existing.paid === paid &&
    existing.total === total &&
    existing.billed === billed &&
    existing.cycle === cycle &&
    existing.completed === completed
  ) {
    return;
  }

  decisions.instProgress[id] = {
    instId: id,
    paid,
    total,
    billed,
    cycle,
    completed,
    note: entry.note || "",
    source: entry.source || "system",
    locked: entry.locked !== false,
    updatedAt: nowCairoTimestamp_()
  };
  decisions._dirty = true;
}

function buildImportMetaPayload_(m, statementDateKey) {
  return {
    statementDate: String((m && m.statementDate) || statementDateKey || "").trim(),
    printedOpening: Number((Number(m && m.printedOpening) || 0).toFixed(2)),
    printedClosing: Number((Number(m && m.printedClosing) || 0).toFixed(2)),
    printedTotalDebit: Number((Number(m && m.printedTotalDebit) || 0).toFixed(2)),
    printedTotalCredit: Number((Number(m && m.printedTotalCredit) || 0).toFixed(2)),
    counts: (m && m.counts) || { DEBIT: 0, INSTALLMENT: 0, FEE: 0, CREDIT: 0, total: 0, ambiguous: 0 },
    tieOutPassed: Boolean(m && m.tieOutPassed),
    mathPassed: Boolean(m && m.mathPassed),
    closingDiff: Number((Number(m && m.closingDiff) || 0).toFixed(2)),
    debitDiff: Number((Number(m && m.debitDiff) || 0).toFixed(2)),
    creditDiff: Number((Number(m && m.creditDiff) || 0).toFixed(2)),
    ambiguousIds: Array.isArray(m && m.ambiguousIds) ? m.ambiguousIds.map(id => String(id || "").trim().toUpperCase()).filter(Boolean) : [],
    reasons: Array.isArray(m && m.reasons) ? m.reasons.map(r => String(r || "")) : [],
    sourceFile: String((m && m.sourceFile) || "").trim()
  };
}

function setImportMeta(decisions, statementDateKey, metaObj) {
  const key = String(statementDateKey || "").trim();
  if (!decisions || !key || !metaObj) return;
  if (!decisions.importMeta) decisions.importMeta = {};

  const nextPayload = buildImportMetaPayload_(metaObj, key);
  const existing = decisions.importMeta[key];
  if (existing) {
    const prevPayload = buildImportMetaPayload_(existing, key);
    if (JSON.stringify(prevPayload) === JSON.stringify(nextPayload)) {
      return;
    }
  }

  decisions.importMeta[key] = Object.assign({}, nextPayload, {
    note: metaObj.note || (nextPayload.tieOutPassed ? "PASSED" : "FAILED"),
    source: metaObj.source || "parser",
    locked: nextPayload.tieOutPassed === true,
    updatedAt: nowCairoTimestamp_()
  });
  decisions._dirty = true;
}

function setStmtTypeConfirm(decisions, stmtId, txType, note) {
  const id = String(stmtId || "").trim().toUpperCase();
  const cleanType = String(txType || "").trim().toUpperCase();
  if (!decisions || !id) return;
  if (!["DEBIT", "CREDIT", "INSTALLMENT", "FEE"].includes(cleanType)) return;
  if (!decisions.stmtTypeConfirm) decisions.stmtTypeConfirm = {};

  const existing = decisions.stmtTypeConfirm[id];
  const cleanNote = String(note || "");
  if (existing && existing.target === cleanType && existing.note === cleanNote) {
    return;
  }

  decisions.stmtTypeConfirm[id] = {
    subjectId: id,
    target: cleanType,
    note: cleanNote,
    source: "user",
    locked: true,
    updatedAt: nowCairoTimestamp_()
  };
  decisions._dirty = true;
}

function lookupTxDescriptionsByIds_(ss, txIds) {
  if (!ss || !txIds || txIds.length === 0) return [];
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!txSheet || txSheet.getLastRow() < 2) return [];
  const idCol = findExistingIdColumn(txSheet, 1);
  if (!idCol) return [];
  const numRows = txSheet.getLastRow() - 1;
  const descs = txSheet.getRange(2, 3, numRows, 1).getValues();
  const ids = txSheet.getRange(2, idCol, numRows, 1).getValues();
  const byId = {};
  for (let i = 0; i < numRows; i++) {
    const id = String(ids[i][0] || "").trim().toUpperCase();
    if (id) byId[id] = String(descs[i][0] || "").trim();
  }
  return txIds.map(tid => byId[String(tid || "").trim().toUpperCase()] || "").filter(Boolean);
}

function lookupStmtDescriptionById_(ss, stmtId) {
  if (!ss || !stmtId) return "";
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  if (!stmtSheet || stmtSheet.getLastRow() < 7) return "";
  const idCol = findExistingIdColumn(stmtSheet, 6);
  if (!idCol) return "";
  const numRows = stmtSheet.getLastRow() - 6;
  const descs = stmtSheet.getRange(7, 4, numRows, 1).getValues();
  const ids = stmtSheet.getRange(7, idCol, numRows, 1).getValues();
  const targetId = String(stmtId || "").trim().toUpperCase();
  for (let i = 0; i < numRows; i++) {
    if (String(ids[i][0] || "").trim().toUpperCase() === targetId) {
      return String(descs[i][0] || "").trim();
    }
  }
  return "";
}

function lookupStmtAmountById_(ss, stmtId) {
  if (!ss || !stmtId) return 0;
  const stmtSheet = ss.getSheetByName(CONFIG.SHEETS.BANK_STATEMENT);
  if (!stmtSheet || stmtSheet.getLastRow() < 7) return 0;
  const idCol = findExistingIdColumn(stmtSheet, 6);
  if (!idCol) return 0;
  const numRows = stmtSheet.getLastRow() - 6;
  const amts = stmtSheet.getRange(7, 6, numRows, 1).getValues();
  const ids = stmtSheet.getRange(7, idCol, numRows, 1).getValues();
  const targetId = String(stmtId || "").trim().toUpperCase();
  for (let i = 0; i < numRows; i++) {
    if (String(ids[i][0] || "").trim().toUpperCase() === targetId) {
      return Math.abs(parseFloat(amts[i][0])) || 0;
    }
  }
  return 0;
}

function lookupTxAmountMapByIds_(ss, txIds) {
  const byId = {};
  if (!ss || !txIds || txIds.length === 0) return byId;
  const txSheet = ss.getSheetByName(CONFIG.SHEETS.TRANSACTIONS);
  if (!txSheet || txSheet.getLastRow() < 2) return byId;
  const idCol = findExistingIdColumn(txSheet, 1);
  if (!idCol) return byId;
  const numRows = txSheet.getLastRow() - 1;
  const amts = txSheet.getRange(2, 5, numRows, 1).getValues();
  const ids = txSheet.getRange(2, idCol, numRows, 1).getValues();
  const wanted = new Set(txIds.map(t => String(t || "").trim().toUpperCase()).filter(Boolean));
  for (let i = 0; i < numRows; i++) {
    const id = String(ids[i][0] || "").trim().toUpperCase();
    if (id && wanted.has(id)) {
      byId[id] = Math.abs(parseFloat(amts[i][0])) || 0;
    }
  }
  return byId;
}

function readLegacyReconStoreRaw_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RECON_STORE);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const store = {
    table2Map: {},
    table2ByBaseKey: {},
    table2List: [],
    confirmedMatches: new Set(),
    neglectedTx: new Set(),
    paymentRecords: {},
    installmentOverrides: {}
  };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  let rowCount = 0;
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
    rowCount++;

    if (type === "CYCLE_PAYMENT") {
      let audit = [];
      try { if (note) audit = JSON.parse(note); } catch (e) { }
      store.paymentRecords[pKey] = {
        cycleKey: pKey,
        dateStr: dStr,
        cycleLabel: desc,
        amount: amt,
        status: payer,
        audit,
        isApplied: isConfirmed
      };
    } else if (type === "TABLE2_ASSIGNMENT" || (!type && payer)) {
      if (payer) {
        const isNeg = payer === "NEGLECT" || isNeglectedPayer(payer);
        const rec = {
          primaryKey: pKey,
          dateStr: dStr,
          desc,
          amount: amt,
          payer: isNeg ? "NEGLECT" : payer,
          note,
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
    } else if (type === "TABLE4_INSTALLMENT" && pKey) {
      const rec = {
        primaryKey: pKey,
        dateStr: dStr,
        desc,
        amount: amt,
        action: payer,
        note,
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
    } else if (type === "TABLE1_CONFIRMED" || isConfirmed) {
      if (pKey) store.confirmedMatches.add(pKey);
    } else if (type === "TABLE3_NEGLECTED" || isNeglected) {
      if (pKey) store.neglectedTx.add(pKey);
    }
  });

  return rowCount > 0 ? store : null;
}

function migrateLegacyPreMatch(ss, tz, decisions, ctx) {
  if (!decisions || decisions.isMigrated) return null;
  const legacy = readLegacyReconStoreRaw_(ss);
  if (!legacy) return null;

  let migratedCount = 0;
  Object.keys(legacy.paymentRecords).forEach(cycleKey => {
    const pr = legacy.paymentRecords[cycleKey];
    if (pr && pr.dateStr && !decisions.paymentRecords[cycleKey]) {
      setCyclePayment(decisions, cycleKey, {
        dateStr: pr.dateStr,
        cycleLabel: pr.cycleLabel,
        amount: pr.amount,
        status: pr.status,
        isApplied: pr.isApplied
      });
      migratedCount++;
    }
  });

  const usedTxIds = new Set();
  (ctx.stmtDebits || []).forEach(st => {
    if (!st.stmtId || decisions.stmtMatch[st.stmtId]) return;
    const pKey = makeReconKey(st.date || st.dateStr, st.amount, st.desc, tz);
    const amtDescKey = `${st.amount.toFixed(2)}_${makeNormDesc(st.desc).substring(0, 20)}`;
    if (legacy.confirmedMatches.has(pKey) || legacy.confirmedMatches.has(amtDescKey)) {
      const match = (ctx.txDebits || []).find(tx =>
        tx.txId && !usedTxIds.has(tx.txId) && Math.abs(tx.amount - st.amount) <= 0.05
      );
      if (match) {
        usedTxIds.add(match.txId);
        setStmtMatch(decisions, st.stmtId, [match.txId], "Migrated confirmed match", "migrated");
        migratedCount++;
      }
    }
  });

  legacy._migratedCount = migratedCount;
  return legacy;
}

function writeMigrationReportSheet_(ss, ambiguous) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.MIGRATION_REPORT);
  sheet.clear();

  sheet.getRange("A1:G1").merge()
    .setValue("📋 Legacy Decision Migration Report — Review Ambiguous Items")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(CONFIG.COLORS.DUE_SOON);

  const headers = ["Legacy Type", "Legacy Key", "Date", "Description", "Amount (EGP)", "Legacy Payer", "Reason"];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER);

  if (ambiguous.length > 0) {
    sheet.getRange(3, 1, ambiguous.length, headers.length).setValues(ambiguous).setFontSize(9);
    sheet.getRange(3, 5, ambiguous.length, 1).setNumberFormat("#,##0.00");
  }
}

function migrateLegacyPostMatch(ss, tz, decisions, legacy, ctx) {
  if (!decisions || decisions.isMigrated || !legacy) return { migrated: 0, ambiguous: [] };

  let migratedCount = legacy._migratedCount || 0;
  const consumedLegacyRecs = new Set();

  function migrateTable2Bucket(items) {
    const totalOccMap = {};
    items.forEach(it => {
      const bk = makeReconKey(it.date || it.dateStr, it.amount, it.desc, tz);
      totalOccMap[bk] = (totalOccMap[bk] || 0) + 1;
    });

    const occTracker = {};
    items.forEach(it => {
      if (!it.stmtId) return;
      const baseKey = makeReconKey(it.date || it.dateStr, it.amount, it.desc, tz);
      occTracker[baseKey] = (occTracker[baseKey] || 0) + 1;
      const occ = occTracker[baseKey];
      const pKeyWithOcc = `${baseKey}_#${occ}`;

      let saved = legacy.table2Map[pKeyWithOcc] || (occ === 1 ? legacy.table2Map[baseKey] : null);
      if (!saved && totalOccMap[baseKey] === 1 && legacy.table2ByBaseKey[baseKey] && legacy.table2ByBaseKey[baseKey].length === 1) {
        saved = legacy.table2ByBaseKey[baseKey][0];
      }

      if (saved && saved.payer) {
        consumedLegacyRecs.add(saved);
        if (!decisions.stmtAssign[it.stmtId]) {
          const cleanPayer = (saved.payer === "NEGLECT" || saved.isNeglected)
            ? "NEGLECT"
            : (formatPayerCell(saved.payer) || saved.payer);
          setStmtAssign(decisions, it.stmtId, cleanPayer, saved.note || "", "migrated");
          migratedCount++;
        }
      }
    });
  }

  migrateTable2Bucket(ctx.stmtFees || []);
  migrateTable2Bucket(ctx.missingInSheet || []);
  migrateTable2Bucket(ctx.stmtCredits || []);

  // Migrate Table 3 neglected transactions
  (ctx.expandedUnmatched || []).forEach(tx => {
    if (!tx.subjectId || decisions.txNeglect[tx.subjectId]) return;
    const isNeg = tx._legacyKeys && tx._legacyKeys.some(k => k && legacy.neglectedTx.has(k));
    if (isNeg) {
      setTxNeglect(decisions, tx.subjectId, "Migrated from legacy neglect", "migrated");
      migratedCount++;
    }
  });

  // Migrate Table 4 installment overrides
  const instOccTracker = {};
  (ctx.installmentReconItems || []).forEach(item => {
    const baseKey = makeReconKey(item.date, item.amount, item.desc, tz);
    instOccTracker[baseKey] = (instOccTracker[baseKey] || 0) + 1;
    const occ = instOccTracker[baseKey];
    const pKeyWithOcc = `${baseKey}_#${occ}`;
    const amtDescKey = `${item.amount.toFixed(2)}_${makeNormDesc(item.desc).substring(0, 30)}`;
    const amtDescKeyWithOcc = `${amtDescKey}_#${occ}`;

    const saved = legacy.installmentOverrides[pKeyWithOcc] ||
                  legacy.installmentOverrides[amtDescKeyWithOcc] ||
                  (occ === 1 ? (legacy.installmentOverrides[baseKey] || legacy.installmentOverrides[amtDescKey]) : null);
    if (saved && saved.action !== undefined && item.subjectId && !decisions.instOverride[item.subjectId]) {
      setInstOverride(decisions, item.subjectId, saved.action, saved.note || "", "migrated");
      migratedCount++;
    }
  });

  // Detect any unconsumed Table 2 legacy entries that ambiguously collide with active statement items
  const allTable2Items = [...(ctx.stmtFees || []), ...(ctx.missingInSheet || []), ...(ctx.stmtCredits || [])];
  const ambiguous = [];
  legacy.table2List.forEach(rec => {
    if (consumedLegacyRecs.has(rec)) return;
    const recBase = String(rec.primaryKey || "").replace(/_#\d+$/, "");
    const candidateItems = allTable2Items.filter(it => {
      const itBase = makeReconKey(it.date || it.dateStr, it.amount, it.desc, tz);
      const itDate = makeNormDateStr(it.date || it.dateStr, tz);
      return itBase === recBase || (rec.dateStr && itDate === rec.dateStr && Math.abs(it.amount - rec.amount) <= 0.02);
    });
    if (candidateItems.length > 0) {
      ambiguous.push([
        "TABLE2_ASSIGNMENT",
        rec.primaryKey,
        rec.dateStr,
        rec.desc,
        rec.amount,
        rec.payer,
        `Matched ${candidateItems.length} statement row(s) without unique occurrence key`
      ]);
    }
  });

  decisions.isMigrated = true;
  decisions.migrationNote = `Migrated ${migratedCount} decision(s); ${ambiguous.length} ambiguous`;
  saveDecisions(ss, decisions);

  if (ambiguous.length > 0) {
    writeMigrationReportSheet_(ss, ambiguous);
  }

  return { migrated: migratedCount, ambiguous };
}

function processActiveReconRowChange(sheet, allRows, row, decisions, tz, options) {
  const noChange = { changed: false, needsRefresh: false, toastMsg: "" };
  const opts = options || {};
  const allowDelete = opts.allowDelete === true;
  const forceStyle = opts.forceStyle === true;
  const rowVals = allRows[row - 1];
  const firstCol = String(rowVals[0] || "").trim();
  const subjectId = String(rowVals[8] || "").trim();
  if (!subjectId) return noChange;

  // Case 1: Edit in Table 2 Needs Pairing (F-*, C-*, I-*, M-*)
  if (/^[FCMI]-\d+/i.test(firstCol)) {
    const prefix = firstCol.charAt(0).toUpperCase();
    const origDesc = String(rowVals[2] || "").trim();
    const amt = parseFloat(rowVals[3]);
    if (isNaN(amt) || amt <= 0 || !origDesc) return noChange;

    let payer = String(rowVals[5] || "").trim();
    const note = String(rowVals[6] || "").trim();

    // Handle 1-click pairing selection ("🔗 Pair: T-00332+T-00333 (...)")
    const pairMatch = payer.match(/^🔗\s*Pair:\s*((?:T-\d+)(?:\s*\+\s*T-\d+)*)/i);
    if (pairMatch) {
      const txIds = pairMatch[1].split("+").map(s => s.trim().toUpperCase()).filter(Boolean);
      if (txIds.length === 0) return noChange;

      const sortedNew = txIds.slice().sort().join(",");
      const existingMatch = decisions.stmtMatch && decisions.stmtMatch[subjectId];
      const existingSorted = existingMatch && existingMatch.targetIds
        ? existingMatch.targetIds.slice().sort().join(",")
        : "";
      if (!forceStyle && existingSorted === sortedNew) {
        return noChange;
      }

      // Remove any conflicting STMT_MATCH on another statement line claiming any of these txIds
      Object.keys(decisions.stmtMatch || {}).forEach(otherStmtId => {
        if (otherStmtId === subjectId) return;
        const other = decisions.stmtMatch[otherStmtId];
        if (other && other.targetIds && other.targetIds.some(tid => txIds.includes(tid))) {
          delete decisions.stmtMatch[otherStmtId];
          decisions._dirty = true;
        }
      });

      setStmtMatch(decisions, subjectId, txIds, note || "Paired by user", "user");
      clearStmtAssign(decisions, subjectId);
      Object.keys(decisions.txNeglect || {}).forEach(negSubj => {
        if (txIds.some(tid => negSubj === tid || negSubj.startsWith(tid + ":"))) {
          clearTxNeglect(decisions, negSubj);
        }
      });

      const ss = sheet.getParent();
      const txDescs = lookupTxDescriptionsByIds_(ss, txIds);
      const learned = maybeLearnMerchantAlias(
        decisions,
        origDesc,
        txDescs,
        "Learned from " + subjectId + " <-> " + txIds.join("+")
      );
      if (learned && learned.learned) {
        syncDecisionsMerchantAliasesToSettings(ss, decisions);
      }
      if (forceStyle) {
        try { sheet.getRange(row, 1, 1, 8).setBackground(CONFIG.COLORS.PAID); } catch (e) { }
      }
      return {
        changed: true,
        needsRefresh: true,
        toastMsg: "🔗 Paired " + subjectId + " ↔ " + txIds.join("+") +
          (learned && learned.learned ? " (learned alias: " + learned.stmtKey + " ↔ " + learned.txKey + ")" : "")
      };
    }

    // Handle 1-click installment link selection ("🔗 Link: I-00003+I-00004 (...)")
    const linkMatch = payer.match(/^🔗\s*Link:\s*((?:I-\d+)(?:\s*\+\s*I-\d+)*)/i);
    if (linkMatch) {
      const instIds = linkMatch[1].split("+").map(s => s.trim().toUpperCase()).filter(Boolean);
      if (instIds.length === 0) return noChange;

      const sortedNew = instIds.slice().sort().join(",");
      const existingLink = decisions.stmtInstLink && decisions.stmtInstLink[subjectId];
      const existingSorted = existingLink && existingLink.targetIds
        ? existingLink.targetIds.slice().sort().join(",")
        : "";
      if (!forceStyle && existingSorted === sortedNew) {
        return noChange;
      }

      Object.keys(decisions.stmtInstLink || {}).forEach(otherStmtId => {
        if (otherStmtId === subjectId) return;
        const other = decisions.stmtInstLink[otherStmtId];
        if (other && other.targetIds && other.targetIds.some(id => instIds.includes(id))) {
          delete decisions.stmtInstLink[otherStmtId];
          decisions._dirty = true;
        }
      });

      setStmtInstLink(decisions, subjectId, instIds, note || "Linked by user", "user");
      const planKey = buildInstPlanKey({ desc: origDesc, amount: amt });
      setInstLink(decisions, planKey, instIds, "Plan link from " + subjectId, "user");
      clearStmtAssign(decisions, subjectId);

      if (forceStyle) {
        try { sheet.getRange(row, 1, 1, 8).setBackground(CONFIG.COLORS.PAID); } catch (e) { }
      }
      return {
        changed: true,
        needsRefresh: true,
        toastMsg: "🔗 Linked " + subjectId + " ↔ " + instIds.join("+") + " (saved for future months)"
      };
    }

    if (payer && !isNeglectedPayer(payer)) {
      const cleanPayer = formatPayerCell(payer);
      if (cleanPayer && cleanPayer !== payer) {
        payer = cleanPayer;
        rowVals[5] = payer;
        sheet.getRange(row, 6).setValue(payer);
      }
    }

    if (!allowDelete) {
      if (!payer) return noChange;
      const existing = decisions.stmtAssign && decisions.stmtAssign[subjectId];
      if (!existing) {
        const defaultFeePayer = getDefaultFeePayer(sheet.getParent());
        if ((prefix === "F" || (prefix === "C" && !isPaymentCreditDesc(origDesc))) && defaultFeePayer && payer === defaultFeePayer) {
          return noChange;
        }
        if (prefix === "C" && isPaymentCreditDesc(origDesc) && isNeglectedPayer(payer)) {
          return noChange;
        }
      }
    }

    if (payer) {
      const isNeg = isNeglectedPayer(payer);
      const targetVal = isNeg ? "NEGLECT" : payer;
      const existing = decisions.stmtAssign[subjectId];
      if (!forceStyle && existing && existing.target === targetVal && existing.note === note) {
        return noChange;
      }
      setStmtAssign(decisions, subjectId, targetVal, note, "user");
      if (forceStyle) {
        sheet.getRange(row, 1, 1, 8).setBackground(isNeg ? "#eeeeee" : CONFIG.COLORS.PAID);
      }
      return { changed: true, needsRefresh: false, toastMsg: "" };
    }

    clearStmtAssign(decisions, subjectId);
    if (forceStyle) {
      sheet.getRange(row, 1, 1, 8).setBackground("#ffffff");
    }
    return { changed: true, needsRefresh: false, toastMsg: "" };
  }

  // Case 2: Edit in Orphan Logs panel (ORPH-* and legacy TX-*)
  if (/^(?:ORPH|TX)-/i.test(firstCol)) {
    const actionVal = String(rowVals[5] || "").trim();
    const groupPairMatch = actionVal.match(/^🔗\s*(Add\s+all\s+to|Pair\s+group):\s*(S-\d+)\s*\(((?:T-\d+)(?:\s*\+\s*T-\d+)*)/i);
    const singlePairMatch = !groupPairMatch ? actionVal.match(/^🔗\s*(Pair(?:\s+with)?|Add(?:\s+to)?):\s*(S-\d+)/i) : null;
    if (groupPairMatch || singlePairMatch) {
      const baseTxId = subjectId.split(":")[0].toUpperCase();
      if (!baseTxId) return noChange;

      let stmtId = "";
      let isExplicitAdd = false;
      let incomingIds = [];
      if (groupPairMatch) {
        isExplicitAdd = /^add/i.test(groupPairMatch[1]);
        stmtId = groupPairMatch[2].toUpperCase();
        const extraIds = groupPairMatch[3].split("+").map(s => s.trim().toUpperCase()).filter(Boolean);
        incomingIds = [...new Set([baseTxId, ...extraIds])];
      } else {
        isExplicitAdd = /^add/i.test(singlePairMatch[1]);
        stmtId = singlePairMatch[2].toUpperCase();
        incomingIds = [baseTxId];
      }
      if (!stmtId || incomingIds.length === 0) return noChange;

      const ss = sheet.getParent();
      const existingMatch = decisions.stmtMatch && decisions.stmtMatch[stmtId];
      const existingIds = (existingMatch && Array.isArray(existingMatch.targetIds))
        ? existingMatch.targetIds.map(id => String(id || "").trim().toUpperCase()).filter(Boolean)
        : [];

      if (!forceStyle && existingIds.length > 0 && incomingIds.every(id => existingIds.includes(id))) {
        return noChange;
      }

      let finalTxIds;
      if (existingIds.length === 0) {
        finalTxIds = incomingIds;
      } else if (isExplicitAdd) {
        finalTxIds = [...new Set([...existingIds, ...incomingIds])];
      } else {
        const stmtAmt = lookupStmtAmountById_(ss, stmtId);
        const amtMap = lookupTxAmountMapByIds_(ss, [...existingIds, ...incomingIds]);
        const newUnique = incomingIds.filter(id => !existingIds.includes(id));
        if (newUnique.length === 0) {
          finalTxIds = existingIds;
        } else {
          const existingSum = existingIds.reduce((s, id) => s + (amtMap[id] || 0), 0);
          const newSum = newUnique.reduce((s, id) => s + (amtMap[id] || 0), 0);
          if (stmtAmt > 0 && (existingSum + newSum) <= stmtAmt + 0.50) {
            finalTxIds = [...existingIds, ...newUnique];
          } else {
            finalTxIds = incomingIds;
          }
        }
      }

      const finalSet = new Set(finalTxIds);
      Object.keys(decisions.stmtMatch || {}).forEach(otherStmtId => {
        if (otherStmtId === stmtId) return;
        const other = decisions.stmtMatch[otherStmtId];
        if (other && other.targetIds && other.targetIds.some(tid => finalSet.has(tid))) {
          const remaining = other.targetIds.filter(tid => !finalSet.has(tid));
          if (remaining.length === 0) {
            delete decisions.stmtMatch[otherStmtId];
            decisions._dirty = true;
          } else {
            setStmtMatch(decisions, otherStmtId, remaining, other.note, other.source);
          }
        }
      });

      setStmtMatch(decisions, stmtId, finalTxIds, String(rowVals[6] || "").trim() || "Paired from orphan log", "user");
      clearStmtAssign(decisions, stmtId);
      Object.keys(decisions.txNeglect || {}).forEach(negSubj => {
        if (finalTxIds.some(tid => negSubj === tid || negSubj.startsWith(tid + ":"))) {
          clearTxNeglect(decisions, negSubj);
        }
      });

      const stmtDesc = lookupStmtDescriptionById_(ss, stmtId);
      const txDescs = lookupTxDescriptionsByIds_(ss, finalTxIds);
      const learned = maybeLearnMerchantAlias(
        decisions,
        stmtDesc,
        txDescs.length > 0 ? txDescs : [String(rowVals[2] || "").trim()],
        "Learned from " + stmtId + " <-> " + finalTxIds.join("+")
      );
      if (learned && learned.learned) {
        syncDecisionsMerchantAliasesToSettings(ss, decisions);
      }
      if (forceStyle) {
        try { sheet.getRange(row, 1, 1, 8).setBackground(CONFIG.COLORS.PAID); } catch (e) { }
      }
      return {
        changed: true,
        needsRefresh: true,
        toastMsg: "🔗 Paired " + stmtId + " ↔ " + finalTxIds.join("+") +
          (learned && learned.learned ? " (learned alias: " + learned.stmtKey + " ↔ " + learned.txKey + ")" : "")
      };
    }

    const isNeg = isNeglectedPayer(actionVal);
    if (isNeg) {
      if (!forceStyle && decisions.txNeglect[subjectId]) return noChange;
      setTxNeglect(decisions, subjectId, String(rowVals[6] || "").trim(), "user");
      if (forceStyle) sheet.getRange(row, 1, 1, 8).setBackground("#eeeeee");
      return { changed: true, needsRefresh: false, toastMsg: "" };
    }

    if (!allowDelete) return noChange;
    const hadNeglect = Boolean(decisions.txNeglect && decisions.txNeglect[subjectId]);
    clearTxNeglect(decisions, subjectId);
    if (forceStyle) sheet.getRange(row, 1, 1, 8).setBackground("#fff8e1");
    return { changed: hadNeglect, needsRefresh: false, toastMsg: "" };
  }

  // Case 3: Edit in Table 4 (INST-*)
  if (/^INST-\d+/i.test(firstCol)) {
    const origDesc = String(rowVals[2] || "").trim();
    const amt = parseFloat(rowVals[3]);
    if (isNaN(amt) || amt <= 0 || !origDesc) return noChange;

    let actionVal = String(rowVals[5] || "").trim();
    const sourceCat = String(rowVals[4] || "").trim();
    const note = String(rowVals[6] || "").trim();

    if (!allowDelete) {
      if (!actionVal) return noChange;
      const existing = decisions.instOverride && decisions.instOverride[subjectId];
      if (!existing) {
        if (sourceCat.includes("Sheet Only") && actionVal === "🚫 Not on Statement (Skip)") return noChange;
        if (sourceCat.includes("Matched") && !isNeglectedPayer(actionVal) && /^✅\s*Billed\s+to\s+/i.test(actionVal)) return noChange;
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
    if (actionVal) {
      const existing = decisions.instOverride[subjectId];
      if (!forceStyle && existing && existing.target === actionVal && existing.note === note) {
        return noChange;
      }
      setInstOverride(decisions, subjectId, actionVal, note, "user");
      if (forceStyle) {
        sheet.getRange(row, 1, 1, 8).setBackground(isNeg ? "#eeeeee" : CONFIG.COLORS.PAID);
      }
      return { changed: true, needsRefresh: false, toastMsg: "" };
    }

    clearInstOverride(decisions, subjectId);
    if (forceStyle) {
      sheet.getRange(row, 1, 1, 8).setBackground(CONFIG.COLORS.DUE_SOON);
    }
    return { changed: true, needsRefresh: false, toastMsg: "" };
  }

  return noChange;
}

function syncActiveReconSheetToDecisions(ss, tz) {
  const startRow = typeof getActiveReconStartShareRow === "function" ? getActiveReconStartShareRow() : 15;
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  if (!sheet || sheet.getLastRow() < startRow) return;

  const lastR = sheet.getLastRow();
  const colsToRead = Math.min(9, Math.max(8, sheet.getLastColumn() || 8));
  const rawRows = sheet.getRange(1, 1, lastR, colsToRead).getValues();
  const allRows = colsToRead >= 9
    ? rawRows
    : rawRows.map(r => [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], ""]);

  const decisions = loadDecisions(ss);
  let changed = false;

  for (let row = startRow; row <= lastR; row++) {
    const res = processActiveReconRowChange(sheet, allRows, row, decisions, tz, { allowDelete: false, forceStyle: false });
    if (res && res.changed) {
      changed = true;
    }
  }

  if (changed) {
    saveDecisions(ss, decisions);
  }
}

function syncActiveReconSheetToStore(ss, tz) {
  return syncActiveReconSheetToDecisions(ss, tz);
}
