const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { createGasSandbox, populateSpreadsheetFromFixture, MockSpreadsheet } = require("./helpers/gas_sandbox");

const tz = "Africa/Cairo";

// ==========================================
// Test 1 & 2: Real October Bundled Cart Split (S-00004) Suggestion, 1-Click Confirm & Cross-Month Reuse
// ==========================================
{
  const realFixturePath = path.join(__dirname, "fixtures", "2026-11-real.json");
  const realFixture = JSON.parse(fs.readFileSync(realFixturePath, "utf8"));

  // Strip the S-00004 bundled cart INST_LINK so we exercise fresh suggestion -> 1-click confirm
  const strippedFixture = Object.assign({}, realFixture, {
    decisions: (realFixture.decisions || []).filter(d => d[0] !== "INST_LINK")
  });

  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);
  populateSpreadsheetFromFixture(ss, strippedFixture, CONFIG);

  let recon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(recon.unlinkedStmtInstallments.length, 1, "Unlinked 5-item cart bundle S-00004 must NOT auto-match silently");
  const s4 = recon.unlinkedStmtInstallments[0];
  assert.strictEqual(s4.stmtId, "S-00004");
  assert.ok(s4.suggestedInstLinks && s4.suggestedInstLinks.length > 0, "S-00004 must surface suggested installment link");
  assert.deepStrictEqual(
    Array.from(s4.suggestedInstLinks[0].instIds).sort(),
    ["I-00003", "I-00004", "I-00005", "I-00006", "I-00007"]
  );

  let ledger = sandbox.buildDebtLedger(ss, tz, recon);
  sandbox.renderActiveReconciliationSheet(ss, tz, ledger, recon);
  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);

  let rowI1 = -1;
  for (let r = 12; r <= ar.getLastRow(); r++) {
    if (ar._getCell(r, 1) === "I-1" && ar._getCell(r, 9) === "S-00004") {
      rowI1 = r;
      break;
    }
  }
  assert.ok(rowI1 > 0, "S-00004 must render in Needs Pairing as I-1");
  assert.strictEqual(ar._getCell(rowI1, 5), "Installment (💡 Link)");

  // Confirm 1-click link
  ar._setCell(rowI1, 6, s4.suggestedInstLinks[0].dropdownLabel);
  sandbox.handleActiveReconEdit({ range: ar.getRange(rowI1, 6) });

  const dec = sandbox.loadDecisions(ss);
  assert.ok(dec.stmtInstLink["S-00004"], "STMT_INST_LINK must persist for S-00004");
  assert.ok(dec.instLinks["merch01|12|261"], "Reusable INST_LINK plan key must persist for future months");

  recon = sandbox.reconcileActiveStatement(ss, tz);
  ledger = sandbox.buildDebtLedger(ss, tz, recon);

  assert.strictEqual(recon.unlinkedStmtInstallments.length, 0, "Zero unlinked installments after 1-click confirm");
  assert.strictEqual(ledger.unresolvedCount, 0);
  const linkedItems = recon.installmentReconItems.filter(it => it.subjectId.startsWith("S-00004::"));
  assert.strictEqual(linkedItems.length, 5, "All 5 bundled plan rows must appear in Table 4 under S-00004");
  assert.ok(linkedItems.every(it => it.sourceCategory === "✅ Linked Split"), "Bundled rows must show '✅ Linked Split'");

  Object.keys(realFixture.expectedShares).forEach(person => {
    const ps = ledger.personShares.find(item => item.person === person);
    const exp = realFixture.expectedShares[person];
    ["purchases", "installments", "fees", "unrecorded", "refunds", "netTotal"].forEach(k => {
      assert.strictEqual(
        Number(ps[k].toFixed(2)),
        exp[k],
        `October replay after S-00004 confirm mismatch for ${person}.${k}: got ${ps[k]}, expected ${exp[k]}`
      );
    });
  });

  // Test 2: Cross-Month Persistence — import November statement (12 OF 12, fresh S- ID) -> auto-links with ZERO clicks!
  sandbox.writeStatementToSheet(ss, {
    statementDate: "30-11-2026",
    dueDate: "27-12-2026",
    creditLimit: 100000,
    openingBalance: 0,
    closingBalance: 260.89,
    totalDebit: 260.89,
    totalCredit: 0,
    card: "1234"
  }, [
    { txDate: "30-11-2026", postDate: "30-11-2026", desc: "MERCH01 12 OF 12", type: "INSTALLMENT", amount: 260.89 }
  ]);

  const novRecon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(novRecon.unlinkedStmtInstallments.length, 0, "Next month's MERCH01 12 OF 12 must auto-link via INST_LINK with zero clicks");
  assert.strictEqual(novRecon.instMatches.length, 1);
  assert.strictEqual(novRecon.instMatches[0].tier, "locked-plan");
  assert.deepStrictEqual(
    Array.from(novRecon.instMatches[0].sheetRows.map(r => r.instId)).sort(),
    ["I-00003", "I-00004", "I-00005", "I-00006", "I-00007"]
  );
}

// ==========================================
// Test 3: Emoji Surrogate-Pair Formula Regression Check Across Both Sheets
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 300, "Total Debit:", 500],
    ["Total Credit:", 200, "Card Number:", "1234", "Net:", 300]
  ]);
  stmt.getRange(7, 1, 4, 7).setValues([
    [1, new Date(2026, 9, 27, 12, 0, 0), new Date(2026, 9, 27, 12, 0, 0), "DIRECT DEBIT PAYMENT", "CREDIT", 200, ""],
    [2, new Date(2026, 9, 31, 12, 0, 0), new Date(2026, 9, 31, 12, 0, 0), "UNLINKED PLAN 1 OF 3", "INSTALLMENT", 100, ""],
    [3, new Date(2026, 9, 31, 12, 0, 0), new Date(2026, 9, 31, 12, 0, 0), "LINKED PLAN 2 OF 3", "INSTALLMENT", 150, ""],
    [4, new Date(2026, 9, 15, 12, 0, 0), new Date(2026, 9, 15, 12, 0, 0), "NUTOPIA-SUTHERLAND", "DEBIT", 50, ""]
  ]);

  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, 1, 5).setValues([[1, new Date(2026, 9, 10, 0, 0, 0), "Orphan", "Mido", 80]]);

  const inst = ss.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
  inst.getRange(1, 1, 1, 11).setValues([["#", "Start", "Desc", "Total", "Months", "C6", "C7", "EMI", "Payer", "Paid", "Status"]]);
  inst.getRange(2, 1, 1, 11).setValues([[1, new Date(2026, 7, 1, 12, 0, 0), "Linked Plan", 450, 3, 0, 0, 150, "Mido", 1, "Ongoing"]]);

  sandbox.updateLiveDashboard();

  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  for (let r = 1; r <= ar.getLastRow(); r++) {
    const colA = String(ar._getCell(r, 1) || "");
    const statusFormula = String(ar._getCell(r, 8) || "");
    if (/^(?:[FCMI]|ORPH|INST)-\d+/.test(colA)) {
      assert.ok(
        !/LEFT\([^)]+,\s*1\)/.test(statusFormula),
        `Row ${r} (${colA}) must NEVER use LEFT(..., 1) for emoji checks: ${statusFormula}`
      );
      if (/^(?:[FCMI]|ORPH|INST)-/.test(colA)) {
        assert.ok(
          statusFormula.includes('SEARCH("🚫"'),
          `Row ${r} (${colA}) status formula should use SEARCH("🚫"...): ${statusFormula}`
        );
      }
    }
  }

  const mb = ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
  const mbLeftFormula = String(mb._getCell(10, 1) || "");
  assert.ok(!/LEFT\([^)]+,\s*1\)/.test(mbLeftFormula), `Member Breakdown A10 must NEVER use LEFT(..., 1) on emojis: ${mbLeftFormula}`);
  assert.ok(mbLeftFormula.includes('ISERROR(SEARCH("🚫"') && mbLeftFormula.includes('ISERROR(SEARCH("🔗"'));
  assert.strictEqual(mb._getCell(9, 15), "Bank Progress", "Member Breakdown Col O header must be Bank Progress");
}

// ==========================================
// Test 4: Anonymized Export of INST_LINK & STMT_INST_LINK + Lossless Replay
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 260.89, "Total Debit:", 260.89],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 260.89]
  ]);
  stmt.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Post", "Description", "Type", "Amount", "Status"]]);
  stmt.getRange(7, 1, 1, 7).setValues([
    [1, new Date(2026, 9, 31, 12, 0, 0), new Date(2026, 9, 31, 12, 0, 0), "AMAZON 11 OF 12", "INSTALLMENT", 260.89, ""]
  ]);

  const inst = ss.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
  inst.getRange(1, 1, 1, 11).setValues([["#", "Start", "Desc", "Total", "Months", "C6", "C7", "EMI", "Payer", "Paid", "Status"]]);
  inst.getRange(2, 1, 2, 11).setValues([
    [1, new Date(2025, 10, 1, 12, 0, 0), "Gadget A", 1200, 12, 0, 0, 100.00, "Mido", 10, "Ongoing"],
    [2, new Date(2025, 10, 1, 12, 0, 0), "Gadget B", 1930.68, 12, 0, 0, 160.89, "Mai", 10, "Ongoing"]
  ]);

  const decSheet = sandbox.getDecisionsSheet(ss);
  decSheet.getRange(2, 1, 2, 7).setValues([
    ["INST_LINK", "amazon|12|261", "I-00001,I-00002", "Confirmed link", "user", true, "2026-11-19"],
    ["STMT_INST_LINK", "S-00001", "I-00001,I-00002", "Confirmed link", "user", true, "2026-11-19"]
  ]);

  const fixture = sandbox.buildAnonymizedFixture(ss, tz);
  const serialized = JSON.stringify(fixture).toLowerCase();
  assert.ok(!serialized.includes("amazon"), "Exported INST_LINK must NOT leak real merchant 'amazon'");
  assert.ok(!serialized.includes("gadget"), "Exported fixture must NOT leak 'gadget'");
  assert.ok(fixture.decisions.some(d => d[0] === "INST_LINK"), "Exported fixture must preserve INST_LINK");
  assert.ok(fixture.decisions.some(d => d[0] === "STMT_INST_LINK"), "Exported fixture must preserve STMT_INST_LINK");

  const replayEnv = createGasSandbox();
  const ssReplay = new MockSpreadsheet();
  replayEnv.setActiveSpreadsheet(ssReplay);
  populateSpreadsheetFromFixture(ssReplay, fixture, replayEnv.CONFIG);
  const replayRecon = replayEnv.sandbox.reconcileActiveStatement(ssReplay, tz);
  const replayLedger = replayEnv.sandbox.buildDebtLedger(ssReplay, tz, replayRecon);

  assert.strictEqual(replayRecon.unlinkedStmtInstallments.length, 0);
  assert.strictEqual(Number(replayLedger.installmentsTotal.toFixed(2)), 260.89);
}

console.log("installments_link.test.js: all Phase 3 installment link, cross-month reuse, emoji, and replay assertions passed.");
