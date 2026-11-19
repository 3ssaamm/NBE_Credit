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

function getOrCreateInstallmentProgressCols(instSheet) {
  if (!instSheet) return { idCol: 12, progressCol: 13, statusCol: 14 };
  const idCol = findExistingIdColumn(instSheet, 1) || 12;
  const lastCol = Math.max(idCol, instSheet.getLastColumn() || idCol);
  const headers = instSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let progressCol = 0;
  let statusCol = 0;
  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c] || "").trim().toLowerCase();
    if (h === "bank progress") progressCol = c + 1;
    else if (h === "bank status") statusCol = c + 1;
  }

  if (!progressCol) progressCol = Math.max(idCol + 1, 13);
  if (!statusCol) statusCol = progressCol + 1;

  const maxCol = instSheet.getMaxColumns();
  if (statusCol > maxCol && typeof instSheet.insertColumnsAfter === "function") {
    instSheet.insertColumnsAfter(maxCol, statusCol - maxCol);
  }

  return { idCol, progressCol, statusCol };
}

function renderInstallmentProgressFormulas(ss, bounds) {
  if (!ss) return null;
  const instSheet = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  if (!instSheet || instSheet.getLastRow() < 1) return null;

  const cols = getOrCreateInstallmentProgressCols(instSheet);
  const { idCol, progressCol, statusCol } = cols;

  instSheet.getRange(1, progressCol, 1, 2)
    .setValues([["Bank Progress", "Bank Status"]])
    .setFontWeight("bold")
    .setFontSize(9)
    .setBackground(CONFIG.COLORS.HEADER);

  const lastRow = instSheet.getLastRow();
  if (lastRow < 2) return cols;

  const numRows = lastRow - 1;
  const idColLet = getColumnLetter(idCol);
  const arRef = `'${CONFIG.SHEETS.ACTIVE_RECON}'`;
  const hasTable4 = Boolean(bounds && bounds.hasTable4 && bounds.startInstRow && bounds.endInstRow >= bounds.startInstRow);

  const progFormulas = [];
  const statusFormulas = [];

  for (let r = 2; r <= lastRow; r++) {
    if (hasTable4) {
      const idRange = `${arRef}!$I$${bounds.startInstRow}:$I$${bounds.endInstRow}`;
      const descRange = `${arRef}!$C$${bounds.startInstRow}:$C$${bounds.endInstRow}`;
      const paidCell = `${arRef}!$H$6`;
      const idxExpr = `IFERROR(MATCH("S-*::" & id, ${idRange}, 0), IFERROR(MATCH("S-*::" & id & ":*", ${idRange}, 0), 0))`;
      const curExpr = `IFERROR(VALUE(REGEXEXTRACT(d, "\\((\\d+)\\s+of\\s+\\d+\\)")), J${r})`;
      const totExpr = `IFERROR(VALUE(REGEXEXTRACT(d, "\\(\\d+\\s+of\\s+(\\d+)\\)")), E${r})`;
      const paidExpr = `IF(ISBLANK(${paidCell}), MAX(0, cur-1), cur)`;

      progFormulas.push([
        `=IF(C${r}="", "", LET(id, TRIM(${idColLet}${r}), idx, ${idxExpr}, IF(idx=0, J${r} & " / " & E${r}, LET(d, INDEX(${descRange}, idx), cur, ${curExpr}, tot, ${totExpr}, paid, ${paidExpr}, paid & " / " & tot))))`
      ]);
      statusFormulas.push([
        `=IF(C${r}="", "", LET(id, TRIM(${idColLet}${r}), idx, ${idxExpr}, IF(idx=0, K${r}, LET(d, INDEX(${descRange}, idx), cur, ${curExpr}, tot, ${totExpr}, paid, ${paidExpr}, IF(AND(tot>0, paid>=tot), "Completed", "Ongoing")))))`
      ]);
    } else {
      progFormulas.push([`=IF(C${r}="", "", J${r} & " / " & E${r})`]);
      statusFormulas.push([`=IF(C${r}="", "", K${r})`]);
    }
  }

  instSheet.getRange(2, progressCol, numRows, 1).setValues(progFormulas).setFontSize(9).setHorizontalAlignment("center");
  instSheet.getRange(2, statusCol, numRows, 1).setValues(statusFormulas).setFontSize(9).setHorizontalAlignment("center");

  return cols;
}
