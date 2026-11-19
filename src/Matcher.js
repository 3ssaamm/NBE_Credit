/**
 * Confidence-Tiered Statement Debit Matcher & Merchant Alias Learner
 *
 * Tiers:
 *   Pass 0 — Locked by ID: User-confirmed pairs persisted in _Decisions (STMT_MATCH).
 *   Pass 1 — Auto (1-to-1): amount within 0.05 EGP, <= 3 calendar days, AND merchant match
 *            (keyword overlap or learned MERCHANT_ALIAS).
 *   Pass 2 — Auto (Same-merchant split): <= 3 calendar days, merchant match, subset sums to
 *            bank line within 0.10 EGP.
 *   Pass 3 — Suggested (Medium confidence): amount+date without merchant match, wide date lag,
 *            small FX/tip diff, or subset split — NEVER silently applied; surfaced for 1-click confirm.
 *
 * NOTE: Apps Script V8 load order across files is arbitrary — never reference CONFIG or Code.js
 * helpers at top-level evaluation time; only call them inside function bodies.
 */

const GENERIC_MERCHANT_KEYS = new Set([
  "fawry", "paymob", "geidea", "geideae", "basata", "instapay", "pos", "online",
  "card", "payment", "transfer", "purchase", "cash", "atm", "bank",
  "mobile", "token", "egypt", "egy", "cairo", "alex", "giza", "maadi",
  "smouha", "downtown", "store", "shop", "pay"
]);

function daysBetweenDates(d1, d2) {
  if (!d1 || !d2) return 999;
  const a = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate(), 12, 0, 0);
  const b = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate(), 12, 0, 0);
  return Math.round(Math.abs(a.getTime() - b.getTime()) / 86400000);
}

function isDedicatedUnrecordedMerchant(desc) {
  const d = String(desc || "").toLowerCase();
  return d.includes("nutopia") || d.includes("sutherland") || d.includes("google one") || d.includes("stamp tax") || d.includes("unr_");
}

function extractLearnableMerchantKey(desc) {
  let s = cleanMerchantName(desc)
    .replace(/[^a-z0-9\u0600-\u06FF\s]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!s) return "";

  const words = s.split(" ").filter(Boolean);
  const withoutGateway = words.filter(w =>
    w !== "fawry" && w !== "paymob" && w !== "geidea" && w !== "geideae" && w !== "basata"
  );
  if (withoutGateway.length === 0) return "";

  s = withoutGateway.join(" ").trim();
  if (!s || s.length < 3 || /^\d+$/.test(s) || GENERIC_MERCHANT_KEYS.has(s)) {
    return "";
  }

  const meaningful = s.split(" ").filter(w =>
    w.length >= 2 && !/^\d+$/.test(w) && !GENERIC_MERCHANT_KEYS.has(w)
  );
  if (meaningful.length === 0) return "";

  return s.substring(0, 40);
}

function hasMerchantMatch(desc1, desc2, decisions) {
  if (hasMerchantKeywordOverlap(desc1, desc2)) return true;
  if (!decisions || !decisions.merchantAliases) return false;

  const aliases = Object.values(decisions.merchantAliases);
  if (aliases.length === 0) return false;

  const k1 = extractLearnableMerchantKey(desc1);
  const k2 = extractLearnableMerchantKey(desc2);
  const c1 = cleanMerchantName(desc1);
  const c2 = cleanMerchantName(desc2);

  for (let i = 0; i < aliases.length; i++) {
    const a = aliases[i];
    if (!a || !a.stmtKey || !a.txKey) continue;
    const exactMatch = (k1 && k2 && ((k1 === a.stmtKey && k2 === a.txKey) || (k1 === a.txKey && k2 === a.stmtKey)));
    const substrMatch = (c1.includes(a.stmtKey) && c2.includes(a.txKey)) || (c1.includes(a.txKey) && c2.includes(a.stmtKey));
    if (exactMatch || substrMatch) return true;
  }
  return false;
}

function maybeLearnMerchantAlias(decisions, stmtDesc, txDescs, note) {
  if (!decisions || !stmtDesc || !txDescs || txDescs.length === 0) {
    return { learned: false };
  }

  const stmtKey = extractLearnableMerchantKey(stmtDesc);
  if (!stmtKey) return { learned: false };

  const txKeys = [...new Set(txDescs.map(extractLearnableMerchantKey).filter(Boolean))];
  if (txKeys.length !== 1) return { learned: false };

  const txKey = txKeys[0];
  if (txKey === stmtKey) return { learned: false };
  if (hasMerchantMatch(stmtDesc, txDescs[0], decisions)) return { learned: false };

  setMerchantAlias(decisions, stmtKey, txKey, note || "", "user");
  return { learned: true, stmtKey, txKey };
}

function formatPairDropdownLabel(txs) {
  const list = txs || [];
  const ids = list.map(t => t.txId).join("+");
  const sum = list.reduce((s, t) => s + (parseFloat(t.amount) || 0), 0).toFixed(2);
  const shops = [...new Set(list.map(t => String(t.desc || "").trim()).filter(Boolean))].slice(0, 2).join("/").substring(0, 22);
  const payers = [...new Set(list.map(t => t.person).filter(Boolean))].join("+");
  return `🔗 Pair: ${ids} (${sum} · ${shops} · ${payers})`;
}

function formatOrphanPairDropdownLabel(st) {
  const descShort = String((st && st.desc) || "").trim().substring(0, 24);
  const amt = st ? (parseFloat(st.amount) || 0).toFixed(2) : "0.00";
  return `🔗 Pair with: ${(st && st.stmtId) || ""} (${amt} · ${descShort})`;
}

function matchStatementDebits(stmtDebits, txDebits, decisions, tz) {
  function markMatch(st, items, tier, reason) {
    st.matched = true;
    st.matchedItems = items;
    st.matchTier = tier;
    st.matchReason = reason;
    items.forEach(tx => {
      tx.matched = true;
      tx.matchedStmtId = st.stmtId;
    });
    st.assignedPayer = [...new Set(items.map(tx => tx.person))].join(" + ");
  }

  // Pass 0: Locked Explicit Matches by ID from _Decisions
  stmtDebits.forEach(st => {
    if (st.matched || !st.stmtId) return;
    const savedMatch = decisions && decisions.stmtMatch && decisions.stmtMatch[st.stmtId];
    if (!savedMatch || !savedMatch.targetIds || savedMatch.targetIds.length === 0) return;

    const matchedTxs = savedMatch.targetIds
      .map(tid => txDebits.find(tx => tx.txId === tid && !tx.matched))
      .filter(Boolean);
    if (matchedTxs.length === savedMatch.targetIds.length) {
      markMatch(st, matchedTxs, "locked", "Confirmed by user");
    }
  });

  // Pass 1: Auto (High-Confidence 1-to-1) — amount <= 0.05, <= 3 calendar days, AND merchant match
  const autoCandidates = [];
  stmtDebits.forEach(st => {
    if (st.matched) return;
    txDebits.forEach(tx => {
      if (tx.matched) return;
      const amtDiff = Math.abs(tx.amount - st.amount);
      if (amtDiff > 0.05) return;
      const dateDiff = daysBetweenDates(st.date, tx.date);
      if (dateDiff > 3) return;
      if (!hasMerchantMatch(st.desc, tx.desc, decisions)) return;
      autoCandidates.push({ st, tx, dateDiff, amtDiff });
    });
  });

  autoCandidates.sort((a, b) => {
    if (a.dateDiff !== b.dateDiff) return a.dateDiff - b.dateDiff;
    if (Math.abs(a.amtDiff - b.amtDiff) > 0.001) return a.amtDiff - b.amtDiff;
    if ((a.st.stmtRowIndex || 0) !== (b.st.stmtRowIndex || 0)) return (a.st.stmtRowIndex || 0) - (b.st.stmtRowIndex || 0);
    return (a.tx.rowNum || 0) - (b.tx.rowNum || 0);
  });

  autoCandidates.forEach(pair => {
    if (!pair.st.matched && !pair.tx.matched) {
      markMatch(pair.st, [pair.tx], "auto", "High-confidence 1-to-1");
    }
  });

  // Pass 2: Auto (High-Confidence Same-Merchant Split) — <= 3 calendar days, merchant match, subset sum <= 0.10
  stmtDebits.forEach(st => {
    if (st.matched) return;
    const nearby = txDebits.filter(tx => {
      if (tx.matched) return false;
      if (daysBetweenDates(st.date, tx.date) > 3) return false;
      return hasMerchantMatch(st.desc, tx.desc, decisions);
    });
    if (nearby.length < 2) return;
    const combo = findSubsetCombination(nearby, st.amount, 6, 0.10);
    if (combo) {
      markMatch(st, combo, "auto-split", "Same-merchant split match");
    }
  });

  // Pass 3: Suggested (Medium-Confidence Candidates — NEVER silently applied)
  stmtDebits.forEach(st => { st.suggestedPairs = []; });
  txDebits.forEach(tx => { tx.suggestedStmts = []; });

  stmtDebits.forEach(st => {
    if (st.matched || isDedicatedUnrecordedMerchant(st.desc)) return;

    const candByKey = {};
    function pushCandidate(txs, score, reason) {
      if (!txs || txs.length === 0) return;
      const txIds = txs.map(t => t.txId).filter(Boolean);
      if (txIds.length !== txs.length) return;
      const dedupKey = txIds.slice().sort().join(",");
      const sumAmount = txs.reduce((s, t) => s + t.amount, 0);
      const roundedScore = Math.round(score * 100) / 100;
      const existing = candByKey[dedupKey];
      if (!existing || roundedScore > existing.score) {
        candByKey[dedupKey] = {
          txIds,
          txs: txs.slice(),
          sumAmount: Number(sumAmount.toFixed(2)),
          score: roundedScore,
          reason,
          dropdownLabel: formatPairDropdownLabel(txs)
        };
      }
    }

    const unmatchedPool = txDebits.filter(tx => !tx.matched);

    // (A) 1-to-1 Amount + Date (unknown merchant)
    // (B) 1-to-1 Same Merchant + Wide Date Lag (4..10 days)
    // (C) 1-to-1 Same Merchant + Small Amount Diff (FX/tip)
    unmatchedPool.forEach(tx => {
      const amtDiff = Math.abs(tx.amount - st.amount);
      const dateDiff = daysBetweenDates(st.date, tx.date);
      const merchantMatches = hasMerchantMatch(st.desc, tx.desc, decisions);

      if (amtDiff <= 0.10 && dateDiff <= 5) {
        const score = 100 - dateDiff * 10 - amtDiff * 100 + (dateDiff === 0 ? 15 : 0);
        const reason = amtDiff <= 0.01
          ? (dateDiff === 0 ? "Exact amount & same date" : `Exact amount (±${dateDiff}d)`)
          : `Close amount (±${dateDiff}d)`;
        pushCandidate([tx], score, reason);
      }

      if (merchantMatches && amtDiff <= 0.10 && dateDiff > 3 && dateDiff <= 10) {
        pushCandidate([tx], 85 - dateDiff * 3, `Same merchant (${dateDiff}d lag)`);
      }

      const smallDiffCap = Math.max(5.0, st.amount * 0.05);
      if (merchantMatches && dateDiff <= 5 && amtDiff > 0.10 && amtDiff <= smallDiffCap) {
        const signedDiff = st.amount - tx.amount;
        const sign = signedDiff >= 0 ? "+" : "";
        pushCandidate([tx], 75 - dateDiff * 8, `Same merchant (diff ${sign}${signedDiff.toFixed(2)})`);
      }
    });

    // (D) Same-shop split (tx rows from same shop within <= 5 days summing to st.amount)
    const nearby5d = unmatchedPool.filter(tx => daysBetweenDates(st.date, tx.date) <= 5);
    const byShop = {};
    nearby5d.forEach(tx => {
      const shopKey = cleanMerchantName(tx.desc) || makeNormDesc(tx.desc);
      if (!shopKey) return;
      if (!byShop[shopKey]) byShop[shopKey] = [];
      byShop[shopKey].push(tx);
    });
    Object.values(byShop).forEach(group => {
      if (group.length < 2) return;
      const combo = findSubsetCombination(group, st.amount, 6, 0.10);
      if (combo) {
        const maxDateDiff = Math.max(...combo.map(t => daysBetweenDates(st.date, t.date)));
        pushCandidate(combo, 92 - maxDateDiff * 8, "Same-shop split sum");
      }
    });

    // (E) Cross-merchant subset split within <= 5 days
    if (nearby5d.length >= 2) {
      const crossCombo = findSubsetCombination(nearby5d, st.amount, 4, 0.10);
      if (crossCombo) {
        const maxDateDiff = Math.max(...crossCombo.map(t => daysBetweenDates(st.date, t.date)));
        pushCandidate(crossCombo, 70 - maxDateDiff * 8, "Split sum match");
      }
    }

    st.suggestedPairs = Object.values(candByKey)
      .sort((a, b) => b.score - a.score)
      .slice(0, 3);
  });

  // Reverse hints for Orphan Logs: which unresolved statement lines suggest each unmatched tx
  txDebits.forEach(tx => {
    if (tx.matched || !tx.txId) return;
    const hints = [];
    stmtDebits.forEach(st => {
      if (st.matched || !st.stmtId || !st.suggestedPairs) return;
      const matchCand = st.suggestedPairs.find(c => c.txIds.includes(tx.txId));
      if (matchCand) {
        hints.push({
          stmtId: st.stmtId,
          stmt: st,
          score: matchCand.score,
          reason: matchCand.reason,
          dropdownLabel: formatOrphanPairDropdownLabel(st)
        });
      }
    });
    hints.sort((a, b) => b.score - a.score);
    tx.suggestedStmts = hints.slice(0, 3);
  });

  return {
    suggestedMatches: stmtDebits.filter(st => !st.matched && st.suggestedPairs && st.suggestedPairs.length > 0)
  };
}
