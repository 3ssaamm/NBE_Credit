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
    paymentRecords: {},
    counters: { S: 0, T: 0, I: 0 },
    isMigrated: false,
    migrationNote: "",
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
  const opts = options || {};
  const allowDelete = opts.allowDelete === true;
  const forceStyle = opts.forceStyle === true;
  const rowVals = allRows[row - 1];
  const firstCol = String(rowVals[0] || "").trim();
  const subjectId = String(rowVals[8] || "").trim();
  if (!subjectId) return false;

  // Case 1: Edit in Table 2 (F-*, C-*, M-*)
  if (/^[FCM]-\d+/i.test(firstCol)) {
    const prefix = firstCol.charAt(0).toUpperCase();
    const origDesc = String(rowVals[2] || "").trim();
    const amt = parseFloat(rowVals[3]);
    if (isNaN(amt) || amt <= 0 || !origDesc) return false;

    let payer = String(rowVals[5] || "").trim();
    const note = String(rowVals[6] || "").trim();

    if (payer && !isNeglectedPayer(payer)) {
      const cleanPayer = formatPayerCell(payer);
      if (cleanPayer && cleanPayer !== payer) {
        payer = cleanPayer;
        rowVals[5] = payer;
        sheet.getRange(row, 6).setValue(payer);
      }
    }

    if (!allowDelete) {
      if (!payer) return false;
      const existing = decisions.stmtAssign && decisions.stmtAssign[subjectId];
      if (!existing) {
        if ((prefix === "F" || (prefix === "C" && !isPaymentCreditDesc(origDesc))) && payer === CONFIG.DEFAULT_FEE_PAYER) {
          return false;
        }
        if (prefix === "C" && isPaymentCreditDesc(origDesc) && isNeglectedPayer(payer)) {
          return false;
        }
      }
    }

    if (payer) {
      const isNeg = isNeglectedPayer(payer);
      const targetVal = isNeg ? "NEGLECT" : payer;
      const existing = decisions.stmtAssign[subjectId];
      if (!forceStyle && existing && existing.target === targetVal && existing.note === note) {
        return false;
      }
      setStmtAssign(decisions, subjectId, targetVal, note, "user");
      if (forceStyle) {
        sheet.getRange(row, 1, 1, 8).setBackground(isNeg ? "#eeeeee" : CONFIG.COLORS.PAID);
      }
      return true;
    }

    clearStmtAssign(decisions, subjectId);
    if (forceStyle) {
      sheet.getRange(row, 1, 1, 8).setBackground("#ffffff");
    }
    return true;
  }

  // Case 2: Edit in Table 3 (TX-*)
  if (/^TX-/i.test(firstCol)) {
    const actionVal = String(rowVals[5] || "").trim();
    const isNeg = isNeglectedPayer(actionVal);

    if (isNeg) {
      if (!forceStyle && decisions.txNeglect[subjectId]) return false;
      setTxNeglect(decisions, subjectId, "", "user");
      if (forceStyle) sheet.getRange(row, 1, 1, 8).setBackground("#eeeeee");
      return true;
    }

    if (!allowDelete) return false;
    clearTxNeglect(decisions, subjectId);
    if (forceStyle) sheet.getRange(row, 1, 1, 8).setBackground("#f8f9fa");
    return true;
  }

  // Case 3: Edit in Table 4 (INST-*)
  if (/^INST-\d+/i.test(firstCol)) {
    const origDesc = String(rowVals[2] || "").trim();
    const amt = parseFloat(rowVals[3]);
    if (isNaN(amt) || amt <= 0 || !origDesc) return false;

    let actionVal = String(rowVals[5] || "").trim();
    const sourceCat = String(rowVals[4] || "").trim();
    const note = String(rowVals[6] || "").trim();

    if (!allowDelete) {
      if (!actionVal) return false;
      const existing = decisions.instOverride && decisions.instOverride[subjectId];
      if (!existing) {
        if (sourceCat.includes("Sheet Only") && actionVal === "🚫 Not on Statement (Skip)") return false;
        if (sourceCat.includes("Matched") && !isNeglectedPayer(actionVal) && /^✅\s*Billed\s+to\s+/i.test(actionVal)) return false;
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
        return false;
      }
      setInstOverride(decisions, subjectId, actionVal, note, "user");
      if (forceStyle) {
        sheet.getRange(row, 1, 1, 8).setBackground(isNeg ? "#eeeeee" : CONFIG.COLORS.PAID);
      }
      return true;
    }

    clearInstOverride(decisions, subjectId);
    if (forceStyle) {
      sheet.getRange(row, 1, 1, 8).setBackground(CONFIG.COLORS.DUE_SOON);
    }
    return true;
  }

  return false;
}

function syncActiveReconSheetToDecisions(ss, tz) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  if (!sheet || sheet.getLastRow() < 12) return;

  const lastR = sheet.getLastRow();
  const colsToRead = Math.min(9, Math.max(8, sheet.getLastColumn() || 8));
  const rawRows = sheet.getRange(1, 1, lastR, colsToRead).getValues();
  const allRows = colsToRead >= 9
    ? rawRows
    : rawRows.map(r => [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], ""]);

  const decisions = loadDecisions(ss);
  let changed = false;

  for (let row = 12; row <= lastR; row++) {
    if (processActiveReconRowChange(sheet, allRows, row, decisions, tz, { allowDelete: false, forceStyle: false })) {
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
