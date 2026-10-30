// Test exact matchStatementInstallments from Code.js

function findSubsetCombination(items, targetAmt, maxK, tolerance) {
  const tol = tolerance || 0.10;
  const n = items.length;
  if (n < 2) return null;
  const sorted = items.slice().sort((a, b) => b.amount - a.amount);

  function backtrack(startIndex, currentSubset, currentSum) {
    if (Math.abs(currentSum - targetAmt) <= tol && currentSubset.length >= 2) {
      return currentSubset.slice();
    }
    if (currentSum > targetAmt + tol || currentSubset.length >= maxK) {
      return null;
    }
    for (let i = startIndex; i < n; i++) {
      currentSubset.push(sorted[i]);
      const res = backtrack(i + 1, currentSubset, currentSum + sorted[i].amount);
      if (res) return res;
      currentSubset.pop();
    }
    return null;
  }

  for (let k = 2; k <= Math.min(n, maxK); k++) {
    const res = backtrack(0, [], 0);
    if (res) return res;
  }
  return null;
}

function matchStatementInstallments(statementInstallments, sheetInstallmentRows) {
  const matchedStmt = new Set();
  const matchedSheet = new Set();
  const results = [];

  statementInstallments.forEach((st, sIdx) => {
    const cand = sheetInstallmentRows.find((sh, shIdx) => {
      if (matchedSheet.has(sh.index)) return false;
      if (Math.abs(st.amount - sh.emi) > 0.50) return false;
      if (st.instTotal && sh.duration && Math.abs(st.instTotal - sh.duration) > 1) return false;
      return true;
    });

    if (cand) {
      matchedStmt.add(sIdx);
      matchedSheet.add(cand.index);
      results.push({
        stmtIdx: sIdx,
        stmt: st,
        sheetRows: [cand],
        payerLabel: cand.payer,
        type: "1-to-1"
      });
    }
  });

  const descGroups = {};
  sheetInstallmentRows.forEach((sh, shIdx) => {
    if (matchedSheet.has(sh.index)) return;
    const key = (sh.desc || "").toLowerCase().trim() + "_" + (sh.duration || 0);
    if (!descGroups[key]) descGroups[key] = [];
    descGroups[key].push(sh);
  });

  statementInstallments.forEach((st, sIdx) => {
    if (matchedStmt.has(sIdx)) return;
    for (const key in descGroups) {
      const group = descGroups[key];
      if (!group || group.length < 2) continue;
      if (st.instTotal && group[0].duration && Math.abs(group[0].duration - st.instTotal) > 1) continue;
      const groupSum = group.reduce((sum, it) => sum + it.emi, 0);
      if (Math.abs(groupSum - st.amount) <= 0.60) {
        matchedStmt.add(sIdx);
        group.forEach(it => matchedSheet.add(it.index));
        const uniquePayers = [...new Set(group.map(it => it.payer))];
        results.push({
          stmtIdx: sIdx,
          stmt: st,
          sheetRows: group,
          payerLabel: uniquePayers.join(" + "),
          type: "SPLIT_SAME_DESC"
        });
        delete descGroups[key];
        break;
      }
    }
  });

  statementInstallments.forEach((st, sIdx) => {
    if (matchedStmt.has(sIdx)) return;
    const available = sheetInstallmentRows.filter((sh, shIdx) => {
      if (matchedSheet.has(sh.index)) return false;
      if (st.instTotal && sh.duration && Math.abs(sh.duration - st.instTotal) > 1) return false;
      return true;
    });
    if (available.length < 2) return;

    const combo = findSubsetCombination(
      available.map(a => ({ ...a, amount: a.emi })),
      st.amount,
      6,
      0.60
    );
    if (combo) {
      matchedStmt.add(sIdx);
      combo.forEach(c => matchedSheet.add(c.index));
      const uniquePayers = [...new Set(combo.map(c => c.payer))];
      results.push({
        stmtIdx: sIdx,
        stmt: st,
        sheetRows: combo,
        payerLabel: uniquePayers.join(" + "),
        type: "SPLIT_SUBSET"
      });
    }
  });

  return results;
}

const sheetRows = [
  { index: 0, desc: "Kettle", emi: 61.50, duration: 12, payer: "Mido" },
  { index: 1, desc: "Kettle", emi: 61.50, duration: 12, payer: "Mai" },
  { index: 2, desc: "Hair dryer", emi: 22.88, duration: 12, payer: "Mai" },
  { index: 3, desc: "Hair dryer", emi: 22.88, duration: 12, payer: "Mido" },
  { index: 4, desc: "Hand Wrest", emi: 8.33, duration: 12, payer: "Mido" },
  { index: 5, desc: "Hard drive", emi: 154.39, duration: 12, payer: "Mido" },
  { index: 6, desc: "Trimmer", emi: 52.42, duration: 12, payer: "Mido" },
  { index: 7, desc: "Deepfreezer", emi: 400.00, duration: 12, payer: "Dad" },
  { index: 8, desc: "Deepfreezer", emi: 405.00, duration: 12, payer: "Mum" },
  { index: 9, desc: "Deepfreezer", emi: 360.00, duration: 12, payer: "Mido" },
  { index: 10, desc: "Deepfreezer", emi: 308.00, duration: 12, payer: "Mai" },
  { index: 11, desc: "Airbods", emi: 106.17, duration: 12, payer: "Muhanad" },
  { index: 12, desc: "Hand blender", emi: 33.00, duration: 12, payer: "Mai" },
  { index: 13, desc: "Hand blender", emi: 15.00, duration: 12, payer: "Mum" },
  { index: 14, desc: "Hand blender", emi: 45.67, duration: 12, payer: "Mido" },
  { index: 15, desc: "Mobilaty", emi: 833.25, duration: 11, payer: "Mai" },
  { index: 16, desc: "Noon 2 of 6", emi: 132.00, duration: 2, payer: "Mido" },
  { index: 17, desc: "Dad's Phone", emi: 610.33, duration: 12, payer: "Dad" },
  { index: 18, desc: "Headphones", emi: 194.33, duration: 9, payer: "Mido" },
  { index: 19, desc: "Cosmetics", emi: 212.46, duration: 12, payer: "Mai" },
  { index: 20, desc: "Shaving Razors", emi: 97.50, duration: 6, payer: "Mai" },
  { index: 21, desc: "Muhanad's Earphones", emi: 74.07, duration: 9, payer: "Mido" },
  { index: 22, desc: "Muhanad's Earphones", emi: 74.07, duration: 9, payer: "Mai" },
  { index: 23, desc: "Muhanad's Earphones", emi: 74.07, duration: 9, payer: "Mum" },
  { index: 24, desc: "Watch organaizer", emi: 85.21, duration: 6, payer: "Abdo" },
  { index: 25, desc: "Iron", emi: 316.50, duration: 6, payer: "Abdo" },
  { index: 26, desc: "Heater", emi: 474.83, duration: 6, payer: "Abdo" },
  { index: 27, desc: "Power bank", emi: 234.33, duration: 6, payer: "Abdo" },
  { index: 28, desc: "Creatine", emi: 252.45, duration: 6, payer: "Mido" },
  { index: 29, desc: "Power bank", emi: 174.07, duration: 6, payer: "Mido" },
  { index: 30, desc: "Cosmatics", emi: 472.64, duration: 6, payer: "Mai" },
  { index: 31, desc: "Fan", emi: 80.00, duration: 6, payer: "Dad" },
  { index: 32, desc: "Fan", emi: 84.33, duration: 6, payer: "Mum" },
  { index: 33, desc: "Fan", emi: 84.33, duration: 6, payer: "Mai" }
];

const stmtList = [
  { desc: "KETTLE 10 OF 12", amount: 123.00, instCurrent: 10, instTotal: 12 },
  { desc: "HAIR DRYER 10 OF 12", amount: 45.76, instCurrent: 10, instTotal: 12 },
  { desc: "HAND WREST 10 OF 12", amount: 8.33, instCurrent: 10, instTotal: 12 },
  { desc: "HARD DRIVE 10 OF 12", amount: 154.39, instCurrent: 10, instTotal: 12 },
  { desc: "TRIMMER 10 OF 12", amount: 52.42, instCurrent: 10, instTotal: 12 },
  { desc: "DEEPFREEZER 10 OF 12", amount: 1473.00, instCurrent: 10, instTotal: 12 },
  { desc: "AIRBODS 10 OF 12", amount: 106.17, instCurrent: 10, instTotal: 12 },
  { desc: "HAND BLENDER 10 OF 12", amount: 93.67, instCurrent: 10, instTotal: 12 },
  { desc: "MOBILATY 10 OF 11", amount: 833.25, instCurrent: 10, instTotal: 11 },
  { desc: "DAD'S PHONE 9 OF 12", amount: 610.33, instCurrent: 9, instTotal: 12 },
  { desc: "HEADPHONES 7 OF 9", amount: 194.33, instCurrent: 7, instTotal: 9 },
  { desc: "COSMETICS 6 OF 12", amount: 212.46, instCurrent: 6, instTotal: 12 },
  { desc: "SHAVING RAZORS 5 OF 6", amount: 97.50, instCurrent: 5, instTotal: 6 },
  { desc: "MUHANAD'S EARPHONES 5 OF 9", amount: 222.21, instCurrent: 5, instTotal: 9 },
  { desc: "WATCH ORGANAIZER 3 OF 6", amount: 85.21, instCurrent: 3, instTotal: 6 },
  { desc: "IRON 3 OF 6", amount: 316.50, instCurrent: 3, instTotal: 6 },
  { desc: "HEATER 3 OF 6", amount: 474.83, instCurrent: 3, instTotal: 6 },
  { desc: "POWER BANK 3 OF 6", amount: 234.33, instCurrent: 3, instTotal: 6 },
  { desc: "CREATINE 3 OF 6", amount: 252.45, instCurrent: 3, instTotal: 6 },
  { desc: "POWER BANK 2 OF 6", amount: 174.07, instCurrent: 2, instTotal: 6 },
  { desc: "FAN 1 OF 6", amount: 248.66, instCurrent: 1, instTotal: 6 }
];

const matches = matchStatementInstallments(stmtList, sheetRows);
console.log(`Matched ${matches.length} statement installments out of ${stmtList.length}:`);
matches.forEach(m => {
  const rowDescs = m.sheetRows.map(r => `${r.desc} (${r.payer})`).join(", ");
  console.log(`- Stmt: ${m.stmt.desc} [instCurrent: ${m.stmt.instCurrent}] -> Rows: ${rowDescs} (${m.type})`);
});

console.assert(matches.length === stmtList.length, "All statement items must be matched!");
console.log("All statement items matched perfectly!");

