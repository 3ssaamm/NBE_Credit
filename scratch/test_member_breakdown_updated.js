const startShareRow = 12;
const endShareRow = 18;
const totShareRow = 19;

const kpiCards = [
  { title: "Net Total Owed", headerRange: "A5:C5", valRange: "A6:C6", formula: `=IF($B$3="All Family Members", 'Active Bill & Reconciliation'!$H$${totShareRow}, IFERROR(SUMIFS('Active Bill & Reconciliation'!$H$${startShareRow}:$H$${endShareRow}, 'Active Bill & Reconciliation'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` },
  { title: "Direct Purchases", headerRange: "D5:F5", valRange: "D6:F6", formula: `=IF($B$3="All Family Members", 'Active Bill & Reconciliation'!$C$${totShareRow}, IFERROR(SUMIFS('Active Bill & Reconciliation'!$C$${startShareRow}:$C$${endShareRow}, 'Active Bill & Reconciliation'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` },
  { title: "Monthly Installments", headerRange: "G5:I5", valRange: "G6:I6", formula: `=IF($B$3="All Family Members", 'Active Bill & Reconciliation'!$D$${totShareRow}, IFERROR(SUMIFS('Active Bill & Reconciliation'!$D$${startShareRow}:$D$${endShareRow}, 'Active Bill & Reconciliation'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` },
  { title: "Unrecorded Statement Debits", headerRange: "J5:L5", valRange: "J6:L6", formula: `=IF($B$3="All Family Members", 'Active Bill & Reconciliation'!$F$${totShareRow}, IFERROR(SUMIFS('Active Bill & Reconciliation'!$F$${startShareRow}:$F$${endShareRow}, 'Active Bill & Reconciliation'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` },
  { title: "Bank Fees & Taxes", headerRange: "M5:N5", valRange: "M6:N6", formula: `=IF($B$3="All Family Members", 'Active Bill & Reconciliation'!$E$${totShareRow}, IFERROR(SUMIFS('Active Bill & Reconciliation'!$E$${startShareRow}:$E$${endShareRow}, 'Active Bill & Reconciliation'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` },
  { title: "Credits & Refunds", headerRange: "O5:P5", valRange: "O6:P6", formula: `=IF($B$3="All Family Members", 'Active Bill & Reconciliation'!$G$${totShareRow}, IFERROR(SUMIFS('Active Bill & Reconciliation'!$G$${startShareRow}:$G$${endShareRow}, 'Active Bill & Reconciliation'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))` }
];

console.log("Card Ranges and Formulas:");
kpiCards.forEach(c => {
  console.log(`${c.title} -> Header: ${c.headerRange}, Value: ${c.valRange}\n  Formula: ${c.formula}`);
});

const noteFormula = `="ℹ️ Breakdown: Purchases (" & TEXT(D6, "#,##0.00") & ") + Installments (" & TEXT(G6, "#,##0.00") & ")" & IF(J6>0, " + Unrecorded (" & TEXT(J6, "#,##0.00") & ")", "") & IF(M6>0, " + Bank Fees (" & TEXT(M6, "#,##0.00") & ")", "") & IF(O6>0, " - Refunds (" & TEXT(O6, "#,##0.00") & ")", "") & " = Total Owed: " & TEXT(A6, "#,##0.00") & " EGP"`;
console.log("\nNote Formula:\n", noteFormula);

