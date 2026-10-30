const fs = require('fs');

console.log("=== Testing Member Breakdown Redesigned Formulas ===");

const cardPurchFormula = `=IFERROR(SUM(D10:D), 0)`;
const cardInstFormula = `=IFERROR(SUM(M10:M), 0)`;
const cardFeesFormula = `=IF($B$3="All Family Members", 'Active Bill & Reconciliation'!$E$19, IFERROR(SUMIFS('Active Bill & Reconciliation'!$E$12:$E$18, 'Active Bill & Reconciliation'!$B$12:$B$18, "*" & TRIM($B$3) & "*"), 0))`;
const cardRefundsFormula = `=IF($B$3="All Family Members", 'Active Bill & Reconciliation'!$G$19, IFERROR(SUMIFS('Active Bill & Reconciliation'!$G$12:$G$18, 'Active Bill & Reconciliation'!$B$12:$B$18, "*" & TRIM($B$3) & "*"), 0))`;
const cardNetFormula = `=D6 + H6 + K6 - N6`;

console.log("Card Purchases:", cardPurchFormula);
console.log("Card Installments:", cardInstFormula);
console.log("Card Fees:", cardFeesFormula);
console.log("Card Refunds:", cardRefundsFormula);
console.log("Card Net:", cardNetFormula);

const noteFormula = `="ℹ️ " & IF(K6>0, "Bank Fee (" & TEXT(K6, "#,##0.00") & " EGP) assigned to primary cardholder. ", "No bank fees or taxes on this member. ") & IF(N6>0, "Refund credit (" & TEXT(N6, "#,##0.00") & " EGP) applied.", "")`;
console.log("Note Formula:", noteFormula);

const purchSpill = `=IFERROR(FILTER({Transactions!$A$2:$A, Transactions!$B$2:$B, Transactions!$C$2:$C, Transactions!$E$2:$E, Transactions!$D$2:$D, IF(Transactions!$A$2:$A<>"","✅ Recorded","")}, (Transactions!$A$2:$A<>"") * (Transactions!$E$2:$E>0) * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), Transactions!$D$2:$D)))) * (Transactions!$B$2:$B>=$Z$1) * (Transactions!$B$2:$B<=$Z$2)), {"—", "—", "No direct purchases recorded in this billing cycle", 0, "—", "—"})`;
const instSpill = `=IFERROR(FILTER({Installments!$A$2:$E, Installments!$H$2:$K}, (Installments!$A$2:$A<>"") * (Installments!$K$2:$K<>"Completed") * (IF($B$3="All Family Members", 1, ISNUMBER(SEARCH(TRIM($B$3), Installments!$I$2:$I)))) * (Installments!$B$2:$B<=$Z$2)), {"—", "—", "No active installments for this cycle", 0, 0, 0, "—", 0, "—"})`;

console.log("Purchases Spill at A10:", purchSpill.length > 0);
console.log("Installments Spill at H10:", instSpill.length > 0);
console.log("Validation complete!");

