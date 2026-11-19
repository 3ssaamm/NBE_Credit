const assert = require("assert");
const { createGasSandbox } = require("./helpers/gas_sandbox");

const { sandbox } = createGasSandbox();

const sampleText = [
  "National Bank of Egypt",
  "Card Number: 5264****1234",
  "Statement Date: 31-10-2026",
  "Due Date: 27-11-2026\tCredit Limit: 100,000.00",
  "Opening Balance: 15,000.00\tClosing Balance: 32,618.95",
  "Total of Credit: 15,531.27\tTotal of Debit: 33,150.22",
  "Transaction Date Posting Date Description Debit Credit",
  "Page No. 1 of 2",
  "as of 31/10/2026 14:22:01",
  "Minimum Payment Due: 1,630.95",
  "05/10/2026\t06/10/2026\tINSTAPAY PAYMENT - THANK YOU\t\t15,000.00",
  "12/10/2026\t13/10/2026\tMICROSOFT*MICROSOFT 365 BMSE\t\t9.00",
  "20/10/2026\t21/10/2026\tNUTOPIA-SUTHERLAND ALEX. -02G\t40.00\t",
  "20/10/2026\t21/10/2026\tNUTOPIA-SUTHERLAND ALEX. -02G\t40.00\t",
  "25/10/2026\t26/10/2026\tKETTLE 10 OF 12\t123.00\t",
  "02/10/2026 03/10/2026 FAWRY*BAZLYKWA ALEX EGY AUTH 987654321 1,447.80",
  "STAMP DUTY 41.01"
].join("\n");

const parsed = sandbox.parseNBEStatementText(sampleText);

assert.strictEqual(parsed.meta.card, "5264****1234");
assert.strictEqual(parsed.meta.openingBalance, 15000.00);
assert.strictEqual(parsed.meta.closingBalance, 32618.95);
assert.strictEqual(parsed.meta.totalCredit, 15531.27);
assert.strictEqual(parsed.meta.totalDebit, 33150.22);
assert.strictEqual(parsed.meta.creditLimit, 100000);
assert.strictEqual(parsed.meta.statementDate, "31-10-2026");

// Noise lines (Page No, as of timestamp, Minimum Payment, headers) must be excluded
assert.strictEqual(parsed.transactions.length, 7, `Expected 7 parsed rows (preserving identical duplicates), got ${parsed.transactions.length}`);

// Check tab-delimited credit vs debit classification
const pay = parsed.transactions[0];
assert.strictEqual(pay.type, "CREDIT");
assert.strictEqual(pay.amount, 15000.00);

const ms = parsed.transactions[1];
assert.strictEqual(ms.type, "CREDIT");
assert.strictEqual(ms.amount, 9.00);

// Identical same-day duplicate transactions must NOT be deduplicated
assert.strictEqual(parsed.transactions[2].amount, 40.00);
assert.strictEqual(parsed.transactions[2].type, "DEBIT");
assert.strictEqual(parsed.transactions[3].amount, 40.00);
assert.strictEqual(parsed.transactions[3].type, "DEBIT");

// Installment X OF Y detection
const kettle = parsed.transactions[4];
assert.strictEqual(kettle.type, "INSTALLMENT");
assert.strictEqual(kettle.instCurrent, 10);
assert.strictEqual(kettle.instTotal, 12);
assert.strictEqual(kettle.amount, 123.00);

// Space-separated line with large auth number (>creditLimit*1.5) picks actual amount 1,447.80
const fawry = parsed.transactions[5];
assert.strictEqual(fawry.type, "DEBIT");
assert.strictEqual(fawry.amount, 1447.80);

// Dateless fee line uses Statement Date and FEE type
const fee = parsed.transactions[6];
assert.strictEqual(fee.type, "FEE");
assert.strictEqual(fee.amount, 41.01);
assert.strictEqual(fee.txDate, "31-10-2026");

console.log("parser.test.js: all PDF statement text parsing assertions passed.");
