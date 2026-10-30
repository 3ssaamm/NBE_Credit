const startShareRow = 12;
const endShareRow = 18;
const totShareRow = 19;
const startExRow = 24;
const endExRow = 27;

const feeSpill = `=IFERROR(FILTER({'Active Bill & Reconciliation'!$A$${startExRow}:$E$${endExRow}, 'Active Bill & Reconciliation'!$G$${startExRow}:$G$${endExRow}}, (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), 'Active Bill & Reconciliation'!$F$${startExRow}:$F$${endExRow})))) * ('Active Bill & Reconciliation'!$A$${startExRow}:$A$${endExRow}<>"")), {"—", "—", "No extra bank fees or adjustments assigned", 0, "—", "—"})`;

const purchSpill = `=IFERROR(FILTER({Transactions!$A$2:$A, Transactions!$B$2:$B, Transactions!$C$2:$C, Transactions!$E$2:$E, Transactions!$D$2:$D, IF(Transactions!$A$2:$A<>"","✅ Recorded","")}, (Transactions!$A$2:$A<>"") * (Transactions!$E$2:$E>0) * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), Transactions!$D$2:$D)))) * (Transactions!$B$2:$B>=$Z$1) * (Transactions!$B$2:$B<=$Z$2)), {"—", "—", "No direct purchases recorded in this billing cycle", 0, "—", "—"})`;

const instSpill = `=IFERROR(FILTER({Installments!$A$2:$E, Installments!$H$2:$K}, (Installments!$A$2:$A<>"") * (Installments!$K$2:$K<>"Completed") * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), Installments!$I$2:$I)))) * (Installments!$B$2:$B<=$Z$2)), {"—", "—", "No active installments for this cycle", 0, 0, 0, "—", 0, "—"})`;

console.log("Fee Spill cols: 6 -> default has 6 items:", feeSpill.includes('{"—", "—", "No extra bank fees or adjustments assigned", 0, "—", "—"}'));
console.log("Purch Spill cols: 6 -> default has 6 items:", purchSpill.includes('{"—", "—", "No direct purchases recorded in this billing cycle", 0, "—", "—"}'));
console.log("Inst Spill cols: 9 -> default has 9 items:", instSpill.includes('{"—", "—", "No active installments for this cycle", 0, 0, 0, "—", 0, "—"}'));

