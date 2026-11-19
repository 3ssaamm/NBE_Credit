const assert = require("assert");
const { createGasSandbox } = require("./helpers/gas_sandbox");

const { sandbox } = createGasSandbox();
const tz = "Africa/Cairo";

// 1. Purchase Billing Cycle Rollover Rules (end-of-month day rolls into next cycle)
const purchaseCases = [
  { date: new Date(2026, 7, 30, 12, 0, 0), expected: "2026-09" }, // Aug 30 -> Sep due
  { date: new Date(2026, 7, 31, 12, 0, 0), expected: "2026-10" }, // Aug 31 rollover -> Oct due
  { date: new Date(2026, 8, 1, 12, 0, 0), expected: "2026-10" },  // Sep 1 -> Oct due
  { date: new Date(2026, 8, 29, 12, 0, 0), expected: "2026-10" }, // Sep 29 -> Oct due
  { date: new Date(2026, 8, 30, 12, 0, 0), expected: "2026-11" }, // Sep 30 rollover -> Nov due
  { date: new Date(2026, 9, 1, 12, 0, 0), expected: "2026-11" },  // Oct 1 -> Nov due
  { date: new Date(2026, 9, 31, 12, 0, 0), expected: "2026-12" }, // Oct 31 rollover -> Dec due
  { date: new Date(2026, 1, 28, 12, 0, 0), expected: "2026-04" }, // Feb 28 non-leap rollover -> Apr due
  { date: new Date(2024, 1, 28, 12, 0, 0), expected: "2024-03" }, // Feb 28 leap year stays -> Mar due
  { date: new Date(2024, 1, 29, 12, 0, 0), expected: "2024-04" }  // Feb 29 leap year rollover -> Apr due
];

purchaseCases.forEach(tc => {
  const res = sandbox.getBillingCycleForPurchase(tc.date, tz);
  assert.ok(res, `Expected cycle info for ${tc.date.toDateString()}`);
  assert.strictEqual(res.cycleSortKey, tc.expected, `Purchase ${tc.date.toDateString()} cycle mismatch`);
});

// 2. Cycle Date Range Midnight Start & End Bounds (00:00:00.000 .. 23:59:59.999)
const novRange = sandbox.getCycleDateRange("2026-11", tz);
assert.strictEqual(novRange.startYear, 2026);
assert.strictEqual(novRange.startMonth, 9);
assert.strictEqual(novRange.startDay, 30);
assert.strictEqual(novRange.endYear, 2026);
assert.strictEqual(novRange.endMonth, 10);
assert.strictEqual(novRange.endDay, 30);

const sep30Midnight = new Date(2026, 8, 30, 0, 0, 0, 0);
const oct30EndOfDay = new Date(2026, 9, 30, 23, 59, 59, 999);
assert.ok(sep30Midnight >= novRange.start && sep30Midnight <= novRange.end, "Sep 30 midnight must be inside Nov cycle");
assert.ok(oct30EndOfDay >= novRange.start && oct30EndOfDay <= novRange.end, "Oct 30 end-of-day must be inside Nov cycle");
assert.ok(new Date(2026, 9, 31, 0, 0, 0, 0) > novRange.end, "Oct 31 midnight must be outside Nov cycle");

// 3. Installment Billing Cycle Indexing
const instCycle0 = sandbox.getBillingCycleForInstallment(new Date(2026, 7, 28, 12, 0, 0), 0, tz);
assert.strictEqual(instCycle0.cycleSortKey, "2026-10", "Aug purchase installment #0 due Oct 2026");
const instCycle1 = sandbox.getBillingCycleForInstallment(new Date(2026, 7, 28, 12, 0, 0), 1, tz);
assert.strictEqual(instCycle1.cycleSortKey, "2026-11", "Aug purchase installment #1 due Nov 2026");

// 4. parseDateValue across formats (ISO, DD/MM/YYYY, MM/DD/YYYY ambiguous day>12, word formats, serial, Date)
const isoD = sandbox.parseDateValue("2026-10-31", tz);
assert.strictEqual(sandbox.makeNormDateStr(isoD, tz), "2026-10-31");

const dmyD = sandbox.parseDateValue("31/08/2026", tz);
assert.strictEqual(sandbox.makeNormDateStr(dmyD, tz), "2026-08-31");

const mdyD = sandbox.parseDateValue("08/31/2026", tz);
assert.strictEqual(sandbox.makeNormDateStr(mdyD, tz), "2026-08-31");

const wordD = sandbox.parseDateValue("November 27, 2026", tz);
assert.strictEqual(sandbox.makeNormDateStr(wordD, tz), "2026-11-27");

const dmyWordD = sandbox.parseDateValue("30 Sep 2026", tz);
assert.strictEqual(sandbox.makeNormDateStr(dmyWordD, tz), "2026-09-30");

// Google Sheets serial 46326 = 2026-10-31
const serialD = sandbox.parseDateValue(46326, tz);
assert.strictEqual(sandbox.makeNormDateStr(serialD, tz), "2026-10-31");

console.log("cycle_dates.test.js: all cycle date and parsing assertions passed.");
