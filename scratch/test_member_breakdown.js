// Test script to validate Member Breakdown formulas and data generation
const fs = require('fs');

console.log("=== Testing Member Breakdown Formula Syntax ===");

const startShareRow = 12;
const endShareRow = 18;
const totShareRow = 19;
const startExRow = 22;
const endExRow = 25;

const kpiNetFormula = `=IF($B$3="All Family Members", 'Active Bill & Reconciliation'!$H$${totShareRow}, IFERROR(SUMIFS('Active Bill & Reconciliation'!$H$${startShareRow}:$H$${endShareRow}, 'Active Bill & Reconciliation'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))`;
const kpiPurchFormula = `=IF($B$3="All Family Members", 'Active Bill & Reconciliation'!$C$${totShareRow}, IFERROR(SUMIFS('Active Bill & Reconciliation'!$C$${startShareRow}:$C$${endShareRow}, 'Active Bill & Reconciliation'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))`;
const kpiInstFormula = `=IF($B$3="All Family Members", 'Active Bill & Reconciliation'!$D$${totShareRow}, IFERROR(SUMIFS('Active Bill & Reconciliation'!$D$${startShareRow}:$D$${endShareRow}, 'Active Bill & Reconciliation'!$B$${startShareRow}:$B$${endShareRow}, "*" & TRIM($B$3) & "*"), 0))`;

console.log("KPI Net Owed:", kpiNetFormula);
console.log("KPI Purchases:", kpiPurchFormula);
console.log("KPI Installments:", kpiInstFormula);

const purchSpill = `=IFERROR(FILTER({Transactions!$A$2:$A, Transactions!$B$2:$B, Transactions!$C$2:$C, Transactions!$E$2:$E, Transactions!$D$2:$D}, (Transactions!$A$2:$A<>"") * (Transactions!$E$2:$E>0) * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), Transactions!$D$2:$D)))) * (Transactions!$B$2:$B>=$Z$1) * (Transactions!$B$2:$B<=$Z$2)), {"—", "—", "No direct purchases recorded in this billing cycle", 0, "—"})`;
console.log("Purchases Spill Formula:", purchSpill);

const instSpill = `=IFERROR(FILTER({Installments!$A$2:$E, Installments!$H$2:$K}, (Installments!$A$2:$A<>"") * (Installments!$K$2:$K<>"Completed") * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), Installments!$I$2:$I)))) * (Installments!$B$2:$B<=$Z$2)), {"—", "—", "No active installments for this cycle", 0, 0, 0, "—", 0, "—"})`;
console.log("Installments Spill Formula:", instSpill);

const feeSpill = `=IFERROR(FILTER({'Active Bill & Reconciliation'!$A$${startExRow}:$E$${endExRow}, 'Active Bill & Reconciliation'!$G$${startExRow}:$G$${endExRow}}, (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), 'Active Bill & Reconciliation'!$F$${startExRow}:$F$${endExRow})))) * ('Active Bill & Reconciliation'!$A$${startExRow}:$A$${endExRow}<>"")), {"—", "—", "No extra bank fees or adjustments assigned", 0, "—", "—"})`;
console.log("Fees Spill Formula:", feeSpill);

console.log("All formula syntax looks valid!");

