const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { createGasSandbox, populateSpreadsheetFromFixture, MockSpreadsheet } = require("./helpers/gas_sandbox");

const fixturesDir = path.join(__dirname, "fixtures");
const fixtureFiles = fs.readdirSync(fixturesDir).filter(f => f.endsWith(".json")).sort();
assert.ok(fixtureFiles.length > 0, "Expected at least one golden fixture in tests/fixtures/");

fixtureFiles.forEach(file => {
  const fixture = JSON.parse(fs.readFileSync(path.join(fixturesDir, file), "utf8"));
  const tz = fixture.timeZone || "Africa/Cairo";
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();

  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);
  populateSpreadsheetFromFixture(ss, fixture, CONFIG);

  // First reconciliation pass + full sheet rendering
  const recon1 = sandbox.reconcileActiveStatement(ss, tz);
  const ledger1 = sandbox.buildDebtLedger(ss, tz, recon1);
  sandbox.renderActiveReconciliationSheet(ss, tz, ledger1, recon1);
  sandbox.renderMemberBreakdownSheet(ss, tz, ledger1, recon1);

  assert.strictEqual(Number(ledger1.closingBill.toFixed(2)), fixture.expectedTotals.closingBill, `${file}: closingBill mismatch`);
  assert.strictEqual(Number(ledger1.purchasesTotal.toFixed(2)), fixture.expectedTotals.purchasesTotal, `${file}: purchasesTotal mismatch`);
  assert.strictEqual(Number(ledger1.installmentsTotal.toFixed(2)), fixture.expectedTotals.installmentsTotal, `${file}: installmentsTotal mismatch`);
  assert.strictEqual(Number(ledger1.feesTotal.toFixed(2)), fixture.expectedTotals.feesTotal, `${file}: feesTotal mismatch`);
  assert.strictEqual(Number(ledger1.unrecordedTotal.toFixed(2)), fixture.expectedTotals.unrecordedTotal, `${file}: unrecordedTotal mismatch`);
  assert.strictEqual(Number(ledger1.refundsTotal.toFixed(2)), fixture.expectedTotals.refundsTotal, `${file}: refundsTotal mismatch`);
  assert.strictEqual(Number(ledger1.familySharesTotal.toFixed(2)), fixture.expectedTotals.familySharesTotal, `${file}: familySharesTotal mismatch`);
  assert.strictEqual(Number(ledger1.auditDifference.toFixed(2)), fixture.expectedTotals.auditDifference, `${file}: auditDifference mismatch`);

  Object.keys(fixture.expectedShares).forEach(person => {
    const ps = ledger1.personShares.find(item => item.person === person);
    assert.ok(ps, `${file}: missing expected person "${person}"`);
    const exp = fixture.expectedShares[person];
    ["purchases", "installments", "fees", "unrecorded", "refunds", "netTotal"].forEach(field => {
      assert.strictEqual(
        Number(ps[field].toFixed(2)),
        exp[field],
        `${file}: ${person}.${field} expected ${exp[field]}, got ${ps[field]}`
      );
    });
  });

  // Unlisted people must have 0.00 netTotal
  ledger1.personShares.forEach(ps => {
    if (!Object.prototype.hasOwnProperty.call(fixture.expectedShares, ps.person)) {
      assert.strictEqual(
        Number(ps.netTotal.toFixed(2)),
        0,
        `${file}: unlisted person "${ps.person}" has non-zero netTotal ${ps.netTotal}`
      );
    }
  });

  // Idempotency check: second refresh pass must keep all shares identical
  sandbox.syncActiveReconSheetToStore(ss, tz);
  const recon2 = sandbox.reconcileActiveStatement(ss, tz);
  const ledger2 = sandbox.buildDebtLedger(ss, tz, recon2);
  Object.keys(fixture.expectedShares).forEach(person => {
    const ps2 = ledger2.personShares.find(item => item.person === person);
    assert.strictEqual(
      Number(ps2.netTotal.toFixed(2)),
      fixture.expectedShares[person].netTotal,
      `${file}: second refresh altered ${person}.netTotal`
    );
  });
});

console.log(`golden_replay.test.js: verified ${fixtureFiles.length} fixture(s) + refresh idempotency.`);
