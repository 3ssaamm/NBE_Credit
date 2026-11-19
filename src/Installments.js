/**
 * Installment Plan Linking & Bank-Derived Progress
 *
 * Links bank statement installment lines (e.g. "MERCH01 11 OF 12") to rows in the
 * Installments sheet by permanent I- IDs (persisted in _Decisions as INST_LINK / STMT_INST_LINK),
 * routes unlinked / ambiguous bundles to Needs Pairing for 1-click confirmation, and derives
 * progress reactively from bank "X OF Y" text without overwriting user columns A:K.
 *
 * NOTE: Apps Script V8 load order across files is arbitrary — never reference CONFIG or Code.js
 * helpers at top-level evaluation time; only call them inside function bodies.
 */

function parseInstallmentProgressText(desc) {
  const m = String(desc || "").match(/(\d+)\s+(?:OF|\/)\s+(\d+)/i);
  if (!m) return { instCurrent: null, instTotal: null };
  return {
    instCurrent: parseInt(m[1], 10),
    instTotal: parseInt(m[2], 10)
  };
}

function cleanInstBaseDesc(desc) {
  let s = String(desc || "")
    .replace(/\b\d+\s+(?:OF|\/)\s+\d+\b/gi, " ")
    .replace(/\(\d+\s+of\s+\d+\)/gi, " ");
  s = cleanMerchantName(s)
    .replace(/[^a-z0-9\u0600-\u06FF\s]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
  return s.substring(0, 30);
}

function buildInstPlanKey(st) {
  const parsed = parseInstallmentProgressText(st && st.desc);
  const total = (st && st.instTotal) || parsed.instTotal || 0;
  const base = cleanInstBaseDesc(st && st.desc) || "inst";
  const roundedAmt = Math.round(Math.abs(parseFloat(st && st.amount) || 0));
  return `${base}|${total}|${roundedAmt}`;
}

function formatInstLinkDropdownLabel(sheetRows) {
  const list = sheetRows || [];
  const ids = list.map(r => r.instId).filter(Boolean).join("+");
  const sum = list.reduce((s, r) => s + (parseFloat(r.emi) || 0), 0).toFixed(2);
  const descs = [...new Set(list.map(r => String(r.desc || "").trim()).filter(Boolean))].slice(0, 2).join("/").substring(0, 24);
  const payers = [...new Set(list.map(r => r.payer).filter(Boolean))].join("+");
  return `🔗 Link: ${ids} (${sum} · ${descs} · ${payers})`;
}

function matchStatementInstallmentsByLink(statementInstallments, sheetInstallmentRows, decisions) {
  const matchedStmt = new Set();
  const matchedSheet = new Set();
  const results = [];

  const stmtList = statementInstallments || [];
  const sheetList = sheetInstallmentRows || [];
  const byInstId = {};
  sheetList.forEach(sh => {
    if (sh.instId) byInstId[sh.instId] = sh;
  });

  // Prioritize active/ongoing installment rows before completed ones while preserving relative order
  const orderedSheetRows = sheetList.slice().sort((a, b) => {
    const aComp = a.isCompleted ? 1 : 0;
    const bComp = b.isCompleted ? 1 : 0;
    return aComp - bComp;
  });

  // Group by description + duration + completion state so ongoing split groups don't mix with completed ones
  const descGroups = {};
  orderedSheetRows.forEach(sh => {
    const key = `${(sh.desc || "").toLowerCase().trim()}_${sh.duration || 0}_${sh.isCompleted ? "C" : "O"}`;
    if (!descGroups[key]) descGroups[key] = [];
    descGroups[key].push(sh);
  });

  function recordMatch(sIdx, st, rows, type, tier) {
    matchedStmt.add(sIdx);
    rows.forEach(r => matchedSheet.add(r.index));
    const uniquePayers = [...new Set(rows.map(r => r.payer))];
    results.push({
      stmtIdx: sIdx,
      stmt: st,
      sheetRows: rows,
      payerLabel: uniquePayers.join(" + "),
      type,
      tier
    });
  }

  // Pass 0a: Exact Per-Statement Link (decisions.stmtInstLink[stmtId])
  stmtList.forEach((st, sIdx) => {
    if (matchedStmt.has(sIdx) || !st.stmtId) return;
    const saved = decisions && decisions.stmtInstLink && decisions.stmtInstLink[st.stmtId];
    if (!saved || !saved.targetIds || saved.targetIds.length === 0) return;
    const resolved = saved.targetIds.map(id => byInstId[id]).filter(sh => sh && !matchedSheet.has(sh.index));
    if (resolved.length === saved.targetIds.length) {
      recordMatch(sIdx, st, resolved, resolved.length === 1 ? "LINKED_1TO1" : "LINKED_SPLIT", "locked");
    }
  });

  // Pass 0b: Reusable Plan Link Across Months (decisions.instLinks)
  const savedLinks = decisions && decisions.instLinks ? Object.values(decisions.instLinks) : [];
  if (savedLinks.length > 0) {
    stmtList.forEach((st, sIdx) => {
      if (matchedStmt.has(sIdx)) return;
      const exactKey = buildInstPlanKey(st);
      const stBase = cleanInstBaseDesc(st.desc);
      const stTotal = st.instTotal || parseInstallmentProgressText(st.desc).instTotal || 0;
      const stCurr = st.instCurrent || parseInstallmentProgressText(st.desc).instCurrent || 0;

      const candidates = [];
      savedLinks.forEach(link => {
        if (!link || !link.instIds || link.instIds.length === 0) return;
        const rows = link.instIds.map(id => byInstId[id]).filter(sh => sh && !matchedSheet.has(sh.index));
        if (rows.length !== link.instIds.length) return;

        const sumEmi = rows.reduce((s, r) => s + (parseFloat(r.emi) || 0), 0);
        const emiDiff = Math.abs(sumEmi - st.amount);
        const isExactKey = link.planKey === exactKey;

        const parts = String(link.planKey || "").split("|");
        const linkBase = parts[0] || "";
        const linkTotal = parseInt(parts[1], 10) || 0;
        const durationOk = !stTotal || !linkTotal || Math.abs(stTotal - linkTotal) <= 1;
        const baseOk = (linkBase && stBase && (linkBase === stBase || hasMerchantKeywordOverlap(linkBase, stBase)));

        if (isExactKey || (baseOk && durationOk && emiDiff <= 1.00)) {
          const avgPaid = rows.reduce((s, r) => s + (parseInt(r.paymentsMade, 10) || 0), 0) / rows.length;
          const progDist = stCurr ? Math.abs((avgPaid + 1) - stCurr) : 0;
          candidates.push({ rows, isExactKey, emiDiff, progDist });
        }
      });

      if (candidates.length > 0) {
        candidates.sort((a, b) => {
          if (a.isExactKey !== b.isExactKey) return a.isExactKey ? -1 : 1;
          if (Math.abs(a.emiDiff - b.emiDiff) > 0.01) return a.emiDiff - b.emiDiff;
          return a.progDist - b.progDist;
        });
        const best = candidates[0];
        recordMatch(sIdx, st, best.rows, best.rows.length === 1 ? "LINKED_1TO1" : "LINKED_SPLIT", "locked-plan");
      }
    });
  }

  // Pass 1: High-Confidence Unique 1-to-1 (unambiguous EMI + duration match)
  stmtList.forEach((st, sIdx) => {
    if (matchedStmt.has(sIdx)) return;
    const candidates = orderedSheetRows.filter(sh => {
      if (matchedSheet.has(sh.index)) return false;
      if (Math.abs(st.amount - sh.emi) > 0.50) return false;
      if (st.instTotal && sh.duration && Math.abs(st.instTotal - sh.duration) > 1) return false;

      const key = `${(sh.desc || "").toLowerCase().trim()}_${sh.duration || 0}_${sh.isCompleted ? "C" : "O"}`;
      const groupLen = descGroups[key] ? descGroups[key].length : 1;
      const isExactAmt = Math.abs(st.amount - sh.emi) <= 0.05;
      const isDescMatch = hasMerchantKeywordOverlap(st.desc, sh.desc);
      if (groupLen > 1 && !isExactAmt && !isDescMatch) return false;
      return true;
    });

    if (candidates.length === 0) return;

    const ongoingPool = candidates.filter(sh => !sh.isCompleted);
    const pool = ongoingPool.length > 0 ? ongoingPool : candidates;

    let chosen = null;
    if (pool.length === 1) {
      chosen = pool[0];
    } else {
      const overlapPool = pool.filter(sh => hasMerchantKeywordOverlap(st.desc, sh.desc));
      if (overlapPool.length === 1) {
        chosen = overlapPool[0];
      } else {
        const sortedByDiff = pool.slice().sort((a, b) => Math.abs(st.amount - a.emi) - Math.abs(st.amount - b.emi));
        if (Math.abs(st.amount - sortedByDiff[0].emi) <= 0.05 && Math.abs(st.amount - sortedByDiff[1].emi) > 0.15) {
          chosen = sortedByDiff[0];
        }
      }
    }

    if (chosen) {
      recordMatch(sIdx, st, [chosen], "1-to-1", "auto");
    }
  });

  // Pass 2: High-Confidence Same-Description Split Group
  stmtList.forEach((st, sIdx) => {
    if (matchedStmt.has(sIdx)) return;
    const matchingGroups = [];
    for (const key in descGroups) {
      const group = descGroups[key];
      if (!group || group.length < 2) continue;
      if (!group.every(it => !matchedSheet.has(it.index))) continue;
      if (st.instTotal && group[0].duration && Math.abs(group[0].duration - st.instTotal) > 1) continue;
      const groupSum = group.reduce((sum, it) => sum + it.emi, 0);
      const diff = Math.abs(groupSum - st.amount);
      if (diff <= 0.60) {
        matchingGroups.push({
          key,
          group,
          diff,
          isOngoing: !group[0].isCompleted,
          overlap: hasMerchantKeywordOverlap(st.desc, group[0].desc)
        });
      }
    }

    if (matchingGroups.length === 0) return;

    const ongoingGroups = matchingGroups.filter(g => g.isOngoing);
    const pool = ongoingGroups.length > 0 ? ongoingGroups : matchingGroups;
    let chosenGroup = null;
    if (pool.length === 1) {
      chosenGroup = pool[0];
    } else {
      const overlapGroups = pool.filter(g => g.overlap);
      if (overlapGroups.length === 1) {
        chosenGroup = overlapGroups[0];
      }
    }

    if (chosenGroup) {
      recordMatch(sIdx, st, chosenGroup.group, "SPLIT_SAME_DESC", "auto-split");
    }
  });

  // Pass 3: Suggested Links (Ambiguous 1-to-1s, Ambiguous Splits & Cross-Item Cart Bundles — NEVER silently applied)
  stmtList.forEach(st => { st.suggestedInstLinks = []; });

  stmtList.forEach((st, sIdx) => {
    if (matchedStmt.has(sIdx)) return;

    const candByKey = {};
    function pushCandidate(rows, score, reason) {
      if (!rows || rows.length === 0) return;
      const instIds = rows.map(r => r.instId).filter(Boolean);
      const fallbackKey = instIds.length === rows.length
        ? instIds.slice().sort().join(",")
        : rows.map(r => r.index).sort((a, b) => a - b).join(",");
      const sumEmi = rows.reduce((s, r) => s + (parseFloat(r.emi) || 0), 0);
      const roundedScore = Math.round(score * 100) / 100;
      const existing = candByKey[fallbackKey];
      if (!existing || roundedScore > existing.score) {
        candByKey[fallbackKey] = {
          instIds,
          sheetRows: rows.slice(),
          sumEmi: Number(sumEmi.toFixed(2)),
          score: roundedScore,
          reason,
          dropdownLabel: formatInstLinkDropdownLabel(rows)
        };
      }
    }

    // (A) Remaining unclaimed 1-to-1 candidates
    orderedSheetRows.forEach(sh => {
      if (matchedSheet.has(sh.index)) return;
      const emiDiff = Math.abs(st.amount - sh.emi);
      if (emiDiff > 1.00) return;
      if (st.instTotal && sh.duration && Math.abs(st.instTotal - sh.duration) > 1) return;
      const overlap = hasMerchantKeywordOverlap(st.desc, sh.desc);
      const score = 100 - emiDiff * 40 + (overlap ? 20 : 0) + (sh.isCompleted ? -15 : 0);
      pushCandidate([sh], score, overlap ? "Same merchant & EMI" : "EMI & duration match");
    });

    // (B) Same-description split groups within 1.00 EGP
    for (const key in descGroups) {
      const group = descGroups[key];
      if (!group || group.length < 2) continue;
      if (!group.every(it => !matchedSheet.has(it.index))) continue;
      if (st.instTotal && group[0].duration && Math.abs(group[0].duration - st.instTotal) > 1) continue;
      const groupSum = group.reduce((s, r) => s + r.emi, 0);
      const emiDiff = Math.abs(groupSum - st.amount);
      if (emiDiff <= 1.00) {
        const overlap = hasMerchantKeywordOverlap(st.desc, group[0].desc);
        pushCandidate(group, 95 - emiDiff * 40 + (overlap ? 10 : 0), "Same-item split group");
      }
    }

    // (C) Cross-item cart subset bundle (SPLIT_CART) across active plans
    const activeAvailable = orderedSheetRows.filter(sh => {
      if (matchedSheet.has(sh.index)) return false;
      if (sh.isCompleted) return false;
      if (st.instTotal && sh.duration && Math.abs(sh.duration - st.instTotal) > 1) return false;
      return true;
    });
    if (activeAvailable.length >= 2) {
      const combo = findSubsetCombination(
        activeAvailable.map(a => Object.assign({}, a, { amount: a.emi })),
        st.amount,
        6,
        0.60
      );
      if (combo) {
        const comboRows = combo.map(c => sheetList.find(sh => sh.index === c.index) || c);
        const comboSum = comboRows.reduce((s, r) => s + r.emi, 0);
        const emiDiff = Math.abs(comboSum - st.amount);
        pushCandidate(comboRows, 88 - emiDiff * 40, `Bundled cart split (${comboRows.length} plans)`);
      }
    }

    st.suggestedInstLinks = Object.values(candByKey)
      .sort((a, b) => b.score - a.score)
      .slice(0, 3);
  });

  return {
    matches: results,
    unmatchedStmtInstallments: stmtList.filter((_, idx) => !matchedStmt.has(idx)),
    matchedStmtIdxSet: matchedStmt,
    matchedSheetIdxSet: matchedSheet
  };
}

function matchStatementInstallments(statementInstallments, sheetInstallmentRows, optDecisions) {
  return matchStatementInstallmentsByLink(statementInstallments, sheetInstallmentRows, optDecisions).matches;
}

function computeBankDerivedInstallmentProgress(matches, sheetInstallmentRows, isBillPaid) {
  const progressByKey = {};
  const matchedIndices = new Set();

  (matches || []).forEach(m => {
    const st = m.stmt || {};
    const parsed = parseInstallmentProgressText(st.desc);
    const cur = st.instCurrent || parsed.instCurrent;
    const tot = st.instTotal || parsed.instTotal;

    (m.sheetRows || []).forEach(sh => {
      matchedIndices.add(sh.index);
      const duration = sh.duration || tot || 0;
      let paidCount;
      if (cur && cur > 0) {
        paidCount = isBillPaid ? cur : Math.max(0, cur - 1);
      } else {
        paidCount = sh.paymentsMade || 0;
      }
      const isCompleted = duration > 0 && paidCount >= duration;
      const remainingMonths = Math.max(0, duration - paidCount);
      const entry = {
        index: sh.index,
        instId: sh.instId || "",
        stmtId: st.stmtId || "",
        stmtDesc: st.desc || "",
        billedCurrent: cur || null,
        billedTotal: tot || duration || null,
        paidCount,
        remainingMonths,
        isCompleted,
        onStatement: true
      };
      progressByKey[sh.index] = entry;
      if (sh.instId) progressByKey[sh.instId] = entry;
    });
  });

  (sheetInstallmentRows || []).forEach(sh => {
    if (matchedIndices.has(sh.index)) return;
    const duration = sh.duration || 0;
    const paidCount = sh.paymentsMade || 0;
    const isCompleted = Boolean(sh.isCompleted || (String(sh.status || "").toLowerCase() === "completed" && duration > 0 && paidCount >= duration));
    const remainingMonths = isCompleted ? 0 : Math.max(0, duration - paidCount);
    const entry = {
      index: sh.index,
      instId: sh.instId || "",
      stmtId: "",
      stmtDesc: "",
      billedCurrent: null,
      billedTotal: duration || null,
      paidCount,
      remainingMonths,
      isCompleted,
      onStatement: false
    };
    progressByKey[sh.index] = entry;
    if (sh.instId) progressByKey[sh.instId] = entry;
  });

  return progressByKey;
}

function parseInstallmentPaidCell(val) {
  if (typeof val === "number" && !isNaN(val)) {
    return Math.max(0, Math.floor(val));
  }
  const s = String(val || "").trim();
  if (!s || s.startsWith("=")) return 0;
  const m = s.match(/^(\d+)(?:\s*\/\s*\d+)?/);
  return m ? parseInt(m[1], 10) : 0;
}

function isLegacyManualInstallmentLayout(instSheet) {
  if (!instSheet || instSheet.getLastRow() < 1) return false;
  const scanWidth = Math.min(12, Math.max(10, instSheet.getLastColumn() || 10));
  const headers = instSheet.getRange(1, 1, 1, scanWidth).getValues()[0];
  const h10 = String(headers[9] || "").trim().toLowerCase();
  const h11 = String(headers[10] || "").trim().toLowerCase();
  if (h10 === "bank progress" && h11 === "bank status") return false;
  if (h10 === "_row_id") return false;
  return true;
}

function migrateLegacyInstallmentColsToProgress(instSheet, decisions, instIds) {
  if (!instSheet || instSheet.getLastRow() < 2 || !decisions) return 0;
  if (!isLegacyManualInstallmentLayout(instSheet)) return 0;

  const lastRow = instSheet.getLastRow();
  const numRows = lastRow - 1;
  const idCol = findExistingIdColumn(instSheet, 1);
  let idList = instIds;
  if ((!idList || idList.length === 0) && idCol > 0) {
    idList = instSheet.getRange(2, idCol, numRows, 1).getValues().map(r => r[0]);
  }
  if (!idList || idList.length === 0) return 0;

  const width = Math.min(11, Math.max(9, instSheet.getLastColumn() || 9));
  const dataRange = instSheet.getRange(2, 1, numRows, width);
  const rows = dataRange.getValues();
  const formulas = typeof dataRange.getFormulas === "function" ? dataRange.getFormulas() : null;

  let migrated = 0;
  for (let idx = 0; idx < rows.length; idx++) {
    const instId = idList[idx] ? String(idList[idx]).trim().toUpperCase() : "";
    if (!instId) continue;
    if (decisions.instProgress && decisions.instProgress[instId]) continue;

    const row = rows[idx];
    const hasDescOrEmi = Boolean(String(row[2] || "").trim() || parseFloat(row[7]) > 0);
    if (!hasDescOrEmi) continue;

    let rawJ = "";
    let rawK = "";
    let fJ = "";
    let fK = "";
    if (idCol === 11) {
      rawK = width >= 10 ? row[9] : "";
      fK = (formulas && formulas[idx] && width >= 10) ? formulas[idx][9] : "";
    } else {
      rawJ = width >= 10 ? row[9] : "";
      rawK = width >= 11 ? row[10] : "";
      fJ = (formulas && formulas[idx] && width >= 10) ? formulas[idx][9] : "";
      fK = (formulas && formulas[idx] && width >= 11) ? formulas[idx][10] : "";
    }

    if (
      String(rawJ || "").trim().startsWith("=") ||
      String(rawK || "").trim().startsWith("=") ||
      String(fJ || "").trim().startsWith("=") ||
      String(fK || "").trim().startsWith("=")
    ) {
      continue;
    }

    const duration = parseInt(row[4], 10) || 0;
    const oldPaid = parseInstallmentPaidCell(rawJ);
    const oldStatus = String(rawK || "").trim().toLowerCase();
    const oldCompleted = oldStatus === "completed" || (duration > 0 && oldPaid >= duration);

    if (oldCompleted || oldPaid > 0) {
      const finalPaid = (oldCompleted && duration > 0) ? Math.max(oldPaid, duration) : oldPaid;
      setInstProgress(decisions, instId, {
        paid: finalPaid,
        total: duration,
        billed: finalPaid,
        cycle: "",
        completed: oldCompleted,
        source: "migrated",
        note: "Migrated from legacy columns"
      });
      migrated++;
    }
  }

  return migrated;
}

function resolveInstallmentBaseline(sh, decisions, activeCycleKey, tz) {
  const instId = sh && sh.instId ? String(sh.instId).trim().toUpperCase() : "";
  const duration = (sh && parseInt(sh.duration, 10)) || 0;

  if (decisions && decisions.instProgress && instId && decisions.instProgress[instId]) {
    const rec = decisions.instProgress[instId];
    let effectivePaid;
    if (rec.cycle && activeCycleKey && activeCycleKey > rec.cycle) {
      effectivePaid = Math.max(rec.paid || 0, rec.billed || 0);
    } else {
      effectivePaid = rec.paid || 0;
    }
    const isCompleted = Boolean(rec.completed || (duration > 0 && effectivePaid >= duration));
    return {
      paymentsMade: effectivePaid,
      isCompleted,
      status: isCompleted ? "Completed" : "Ongoing",
      source: "decisions"
    };
  }

  if (sh) {
    const rawPaid = parseInstallmentPaidCell(sh.rawLegacyPaid);
    const rawStatus = String(sh.rawLegacyStatus || "").trim().toLowerCase();
    const legacyCompleted = rawStatus === "completed" || (duration > 0 && rawPaid >= duration);
    if (legacyCompleted || rawPaid > 0) {
      const finalPaid = (legacyCompleted && duration > 0) ? Math.max(rawPaid, duration) : rawPaid;
      return {
        paymentsMade: finalPaid,
        isCompleted: legacyCompleted,
        status: legacyCompleted ? "Completed" : "Ongoing",
        source: "legacy"
      };
    }

    if (sh.startDate && duration > 0 && activeCycleKey && typeof getBillingCycleForInstallment === "function") {
      const lastCyc = getBillingCycleForInstallment(sh.startDate, duration - 1, tz);
      if (lastCyc && lastCyc.cycleSortKey && lastCyc.cycleSortKey < activeCycleKey) {
        return {
          paymentsMade: duration,
          isCompleted: true,
          status: "Completed",
          source: "calendar"
        };
      }
    }
  }

  return {
    paymentsMade: 0,
    isCompleted: false,
    status: "Ongoing",
    source: "default"
  };
}

function readSheetInstallmentRows(instSheet, instIds, decisions, activeCycleKey, tz) {
  if (!instSheet || instSheet.getLastRow() < 2) return [];

  if (decisions) {
    migrateLegacyInstallmentColsToProgress(instSheet, decisions, instIds);
  }

  const legacyLayout = isLegacyManualInstallmentLayout(instSheet);
  const idCol = findExistingIdColumn(instSheet, 1);
  const numRows = instSheet.getLastRow() - 1;
  const width = legacyLayout ? Math.min(11, Math.max(9, instSheet.getLastColumn() || 9)) : 9;
  const dataRange = instSheet.getRange(2, 1, numRows, width);
  const instData = dataRange.getValues();
  const formulas = (legacyLayout && typeof dataRange.getFormulas === "function") ? dataRange.getFormulas() : null;

  const sheetInstallmentRows = [];
  instData.forEach((row, idx) => {
    const startDate = parseDateValue(row[1], tz);
    const desc = String(row[2] || "").trim();
    const totalPrice = parseFloat(row[3]) || 0;
    const duration = parseInt(row[4], 10);
    const emi = parseFloat(row[7]);
    const payer = String(row[8] || "").replace(/[\u00A0\s]+/g, " ").trim();
    if (isNaN(emi) || emi <= 0 || !payer) return;

    const instId = (instIds && instIds[idx]) ? String(instIds[idx]).trim().toUpperCase() : "";
    let rawLegacyPaid = "";
    let rawLegacyStatus = "";
    if (legacyLayout) {
      let candJ = "";
      let candK = "";
      let fJ = "";
      let fK = "";
      if (idCol === 11) {
        candK = width >= 10 ? row[9] : "";
        fK = (formulas && formulas[idx] && width >= 10) ? formulas[idx][9] : "";
      } else {
        candJ = width >= 10 ? row[9] : "";
        candK = width >= 11 ? row[10] : "";
        fJ = (formulas && formulas[idx] && width >= 10) ? formulas[idx][9] : "";
        fK = (formulas && formulas[idx] && width >= 11) ? formulas[idx][10] : "";
      }
      if (!String(candJ || "").trim().startsWith("=") && !String(fJ || "").trim().startsWith("=")) {
        rawLegacyPaid = candJ;
      }
      if (!String(candK || "").trim().startsWith("=") && !String(fK || "").trim().startsWith("=")) {
        rawLegacyStatus = candK;
      }
    }

    const base = resolveInstallmentBaseline(
      { instId, startDate, duration, rawLegacyPaid, rawLegacyStatus },
      decisions,
      activeCycleKey,
      tz
    );

    sheetInstallmentRows.push({
      index: idx,
      instId,
      rowNum: idx + 2,
      startDate,
      dateStr: startDate ? makeNormDateStr(startDate, tz) : String(row[1] || ""),
      desc,
      totalPrice,
      duration,
      emi,
      payer,
      paymentsMade: base.paymentsMade,
      status: base.status,
      isCompleted: base.isCompleted
    });
  });

  return sheetInstallmentRows;
}

function persistBankObservedInstallmentProgress(progressMap, decisions, activeCycleKey) {
  if (!progressMap || !decisions) return;
  const seen = new Set();
  Object.keys(progressMap).forEach(k => {
    const prog = progressMap[k];
    if (!prog || !prog.onStatement || !prog.instId) return;
    const id = String(prog.instId).trim().toUpperCase();
    if (!id || seen.has(id)) return;
    seen.add(id);

    setInstProgress(decisions, id, {
      paid: prog.paidCount,
      total: prog.billedTotal || 0,
      billed: prog.billedCurrent !== null && prog.billedCurrent !== undefined ? prog.billedCurrent : prog.paidCount,
      cycle: activeCycleKey || "",
      completed: prog.isCompleted,
      source: "bank",
      note: prog.stmtId ? ("From " + prog.stmtId) : ""
    });
  });
}

function normalizeInstallmentsColumns(instSheet) {
  if (!instSheet) return { idCol: 12, progressCol: 10, statusCol: 11 };

  const maxCol = instSheet.getMaxColumns();
  if (maxCol < 12 && typeof instSheet.insertColumnsAfter === "function") {
    instSheet.insertColumnsAfter(maxCol, 12 - maxCol);
  }

  const lastRow = Math.max(1, instSheet.getLastRow());
  const existingIdCol = findExistingIdColumn(instSheet, 1);
  if (existingIdCol > 0 && existingIdCol !== 12) {
    const idColVals = instSheet.getRange(1, existingIdCol, lastRow, 1).getValues();
    instSheet.getRange(1, 12, lastRow, 1).setValues(idColVals);
    if (existingIdCol === 10 || existingIdCol === 11) {
      instSheet.getRange(1, existingIdCol, lastRow, 1).clearContent();
    }
  } else if (existingIdCol === 0) {
    instSheet.getRange(1, 12).setValue(ROW_ID_HEADER);
  }

  try { instSheet.showColumns(10, 2); } catch (e) { }
  try { instSheet.hideColumns(12); } catch (e) { }

  const lastCol = instSheet.getLastColumn();
  for (let c = lastCol; c >= 13; c--) {
    const h = String(instSheet.getRange(1, c).getValue() || "").trim().toLowerCase();
    if (h === "bank progress" || h === "bank status" || h === "_row_id") {
      if (typeof instSheet.deleteColumn === "function") {
        instSheet.deleteColumn(c);
      } else {
        instSheet.getRange(1, c, lastRow, 1).clearContent();
      }
    }
  }

  return { idCol: 12, progressCol: 10, statusCol: 11 };
}

function getOrCreateInstallmentProgressCols(instSheet) {
  if (!instSheet) return { idCol: 12, progressCol: 10, statusCol: 11 };
  const idCol = findExistingIdColumn(instSheet, 1) || 12;
  const lastCol = Math.max(11, instSheet.getLastColumn() || 11);
  const headers = instSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let progressCol = 0;
  let statusCol = 0;
  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c] || "").trim().toLowerCase();
    if (h === "bank progress") progressCol = c + 1;
    else if (h === "bank status") statusCol = c + 1;
  }

  return {
    idCol,
    progressCol: progressCol || 10,
    statusCol: statusCol || 11
  };
}

function renderInstallmentProgressFormulas(ss, bounds, optContext) {
  if (!ss) return null;
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  if (!instSheet || instSheet.getLastRow() < 1) return null;

  const ctx = optContext || {};
  const tz = ctx.tz || ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const decisions = ctx.decisions || loadDecisions(ss);
  const activeCycleKey = ctx.activeCycleKey || "";
  const progressMap = ctx.progressMap || null;

  const lastRow = instSheet.getLastRow();
  if (lastRow >= 2) {
    const preIdCol = findExistingIdColumn(instSheet, 1);
    const preIds = preIdCol > 0
      ? instSheet.getRange(2, preIdCol, lastRow - 1, 1).getValues().map(r => r[0])
      : [];
    migrateLegacyInstallmentColsToProgress(instSheet, decisions, preIds);
    if (decisions._dirty && (!optContext || !optContext.decisions)) {
      saveDecisions(ss, decisions);
    }
  }

  const cols = normalizeInstallmentsColumns(instSheet);
  const { idCol, progressCol, statusCol } = cols;

  instSheet.getRange(1, progressCol, 1, 2)
    .setValues([["Bank Progress", "Bank Status"]])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER);

  if (lastRow < 2) return cols;

  const numRows = lastRow - 1;
  const rowMeta = instSheet.getRange(2, 1, numRows, 5).getValues();
  const idVals = instSheet.getRange(2, idCol, numRows, 1).getValues();
  const idColLet = getColumnLetter(idCol);
  const arRef = `'${CONFIG.SHEETS.ACTIVE_RECON}'`;
  const hasTable4 = Boolean(bounds && bounds.hasTable4 && bounds.startInstRow && bounds.endInstRow >= bounds.startInstRow);

  const progFormulas = [];
  const statusFormulas = [];

  for (let r = 2; r <= lastRow; r++) {
    const idx = r - 2;
    const startDate = parseDateValue(rowMeta[idx][1], tz);
    const duration = parseInt(rowMeta[idx][4], 10) || 0;
    const instId = String(idVals[idx][0] || "").trim().toUpperCase();
    const prog = (progressMap && ((instId && progressMap[instId]) || progressMap[idx])) || null;

    let fallbackPaid;
    let fallbackCompleted;
    if (prog && !prog.onStatement) {
      fallbackPaid = parseInt(prog.paidCount, 10) || 0;
      fallbackCompleted = Boolean(prog.isCompleted);
    } else {
      const base = resolveInstallmentBaseline({ instId, startDate, duration }, decisions, activeCycleKey, tz);
      fallbackPaid = base.paymentsMade;
      fallbackCompleted = base.isCompleted;
    }

    const curFallback = (prog && prog.billedCurrent) ? prog.billedCurrent : fallbackPaid;
    const offStmtStatusExpr = fallbackCompleted
      ? '"Completed"'
      : `IF(AND(E${r}>0, ${fallbackPaid}>=E${r}), "Completed", "Ongoing")`;

    if (hasTable4) {
      const idRange = `${arRef}!$I$${bounds.startInstRow}:$I$${bounds.endInstRow}`;
      const descRange = `${arRef}!$C$${bounds.startInstRow}:$C$${bounds.endInstRow}`;
      const paidCell = `${arRef}!$H$6`;
      const idxExpr = `IFERROR(MATCH("S-*::" & id, ${idRange}, 0), IFERROR(MATCH("S-*::" & id & ":*", ${idRange}, 0), 0))`;
      const curExpr = `IFERROR(VALUE(REGEXEXTRACT(d, "\\((\\d+)\\s+of\\s+\\d+\\)")), ${curFallback})`;
      const totExpr = `IFERROR(VALUE(REGEXEXTRACT(d, "\\(\\d+\\s+of\\s+(\\d+)\\)")), E${r})`;
      const paidExpr = `IF(ISBLANK(${paidCell}), MAX(0, cur-1), cur)`;

      progFormulas.push([
        `=IF(C${r}="", "", LET(id, TRIM(${idColLet}${r}), idx, ${idxExpr}, IF(idx=0, ${fallbackPaid} & " / " & E${r}, LET(d, INDEX(${descRange}, idx), cur, ${curExpr}, tot, ${totExpr}, paid, ${paidExpr}, paid & " / " & tot))))`
      ]);
      statusFormulas.push([
        `=IF(C${r}="", "", LET(id, TRIM(${idColLet}${r}), idx, ${idxExpr}, IF(idx=0, ${offStmtStatusExpr}, LET(d, INDEX(${descRange}, idx), cur, ${curExpr}, tot, ${totExpr}, paid, ${paidExpr}, IF(AND(tot>0, paid>=tot), "Completed", "Ongoing")))))`
      ]);
    } else {
      progFormulas.push([`=IF(C${r}="", "", ${fallbackPaid} & " / " & E${r})`]);
      statusFormulas.push([`=IF(C${r}="", "", ${offStmtStatusExpr})`]);
    }
  }

  const progRange = instSheet.getRange(2, progressCol, numRows, 1);
  const statusRange = instSheet.getRange(2, statusCol, numRows, 1);
  try { progRange.clearDataValidations(); } catch (e) { }
  try { statusRange.clearDataValidations(); } catch (e) { }
  progRange.setNumberFormat("@").setValues(progFormulas).setFontSize(9).setHorizontalAlignment("center");
  statusRange.setNumberFormat("@").setValues(statusFormulas).setFontSize(9).setHorizontalAlignment("center");

  return cols;
}
