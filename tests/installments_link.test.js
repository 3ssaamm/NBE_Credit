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

// ==========================================
// Test 5: Real October Fixture In-Place Col J/K Replacement, Duplicate M/N Cleanup & Two-Pass Stability
// ==========================================
{
  const realFixturePath = path.join(__dirname, "fixtures", "2026-11-real.json");
  const realFixture = JSON.parse(fs.readFileSync(realFixturePath, "utf8"));

  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);
  populateSpreadsheetFromFixture(ss, realFixture, CONFIG);

  const inst = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  // Simulate leftover Phase 3 duplicate columns M & N on Installments
  inst.getRange(1, 13, 1, 2).setValues([["Bank Progress", "Bank Status"]]);
  inst.getRange(2, 13, 2, 2).setValues([
    ["=LET(...)", "=LET(...)"],
    ["=LET(...)", "=LET(...)"]
  ]);
  // Add sample manual data validation dropdown rule on Col K to verify clearDataValidations removes it
  inst.getRange(2, 11, 5, 1).setDataValidation({ dummy: true });

  // Pass 1 & Pass 2 full dashboard refresh
  sandbox.updateLiveDashboard();
  sandbox.updateLiveDashboard();

  assert.strictEqual(inst._getCell(1, 10), "Bank Progress", "Col J must become Bank Progress");
  assert.strictEqual(inst._getCell(1, 11), "Bank Status", "Col K must become Bank Status");
  assert.strictEqual(inst._getCell(1, 12), "_Row_ID", "Col L must stay hidden _Row_ID");
  assert.ok(inst.hiddenCols.has(12), "Installments Col L (_Row_ID) must be hidden");
  assert.ok(!inst.hiddenCols.has(10) && !inst.hiddenCols.has(11), "Cols J & K must be visible");
  assert.strictEqual(inst.getLastColumn(), 12, "Duplicate Phase 3 Cols M/N must be deleted");
  assert.strictEqual(inst.validations["2_11"], undefined, "Stale manual dropdown validation on Col K must be cleared");

  // Every data row 2..40 must have non-self-referencing formulas
  for (let r = 2; r <= 40; r++) {
    const fJ = String(inst._getCell(r, 10));
    const fK = String(inst._getCell(r, 11));
    assert.ok(fJ.startsWith("=") && !/\b[JK]\d+\b/.test(fJ), `Row ${r} Col J formula must not self-reference: ${fJ}`);
    assert.ok(fK.startsWith("=") && !/\b[JK]\d+\b/.test(fK), `Row ${r} Col K formula must not self-reference: ${fK}`);
  }

  // I-00016 & I-00017 stay Completed and excluded from Table 4 Sheet Only; I-00036..I-00039 stay Ongoing 0/12
  assert.ok(String(inst._getCell(17, 10)).includes('12 & " / " & E17') || String(inst._getCell(17, 10)).includes("12"), "I-00016 must keep 12/12 progress");
  assert.ok(String(inst._getCell(17, 11)).includes('"Completed"'), "I-00016 must stay Completed");
  assert.ok(String(inst._getCell(18, 11)).includes('"Completed"'), "I-00017 must stay Completed");
  assert.ok(String(inst._getCell(37, 10)).includes('0 & " / " & E37'), "Future I-00036 must show 0 / E37");
  assert.ok(String(inst._getCell(37, 11)).includes('"Ongoing"'), "Future I-00036 must stay Ongoing");

  const recon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(Number(recon.cardBal.totalBlockedInstallments.toFixed(2)), 19542.07);
  const sheetOnlyIds = recon.unmatchedSheetInstallments.map(x => x.instId).sort();
  assert.deepStrictEqual(Array.from(sheetOnlyIds), ["I-00036", "I-00037", "I-00038", "I-00039"]);

  const mb = ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
  const mbInstFormula = String(mb._getCell(10, 8) || "");
  assert.ok(mbInstFormula.includes("$I$2:$K"), `Member Breakdown right table must project $I$2:$K: ${mbInstFormula}`);

  // Confirm Test 5 migrated I-00016 & I-00017 into INST_PROGRESS
  const dec5 = sandbox.loadDecisions(ss);
  assert.ok(dec5.instProgress["I-00016"] && dec5.instProgress["I-00016"].completed === true, "I-00016 must persist Completed in INST_PROGRESS");
  assert.ok(dec5.instProgress["I-00017"] && dec5.instProgress["I-00017"].completed === true, "I-00017 must persist Completed in INST_PROGRESS");
}

// ==========================================
// Test 6: Recovery If User Already Manually Deleted Cols J & K
// ==========================================
{
  const realFixturePath = path.join(__dirname, "fixtures", "2026-11-real.json");
  const realFixture = JSON.parse(fs.readFileSync(realFixturePath, "utf8"));

  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);
  populateSpreadsheetFromFixture(ss, realFixture, CONFIG);

  const inst = ss.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  // Add Phase 3 M/N then simulate user deleting manual Cols J & K (cols 10..11) BEFORE running updateLiveDashboard
  inst.getRange(1, 13, 1, 2).setValues([["Bank Progress", "Bank Status"]]);
  inst.deleteColumns(10, 2);
  assert.strictEqual(inst._getCell(1, 10), "_Row_ID", "After user deletes J & K, _Row_ID shifted to Col 10");

  sandbox.updateLiveDashboard();

  assert.strictEqual(inst._getCell(1, 10), "Bank Progress", "Col J must be restored as Bank Progress");
  assert.strictEqual(inst._getCell(1, 11), "Bank Status", "Col K must be restored as Bank Status");
  assert.strictEqual(inst._getCell(1, 12), "_Row_ID", "Col L must be restored as hidden _Row_ID");
  assert.strictEqual(inst._getCell(2, 12), "I-00001", "Existing I-00001 ID must be preserved cleanly in Col L");
  assert.strictEqual(inst._getCell(17, 12), "I-00016", "Existing I-00016 ID must be preserved cleanly in Col L");
  assert.ok(inst.hiddenCols.has(12), "Col L (_Row_ID) must be hidden");
  assert.ok(!inst.hiddenCols.has(10) && !inst.hiddenCols.has(11), "Cols J & K must be unhidden/visible");
  assert.strictEqual(inst.getLastColumn(), 12);

  const recon = sandbox.reconcileActiveStatement(ss, tz);
  // I-00017 (2025-11-30, 2m -> last cycle 2026-02 < 2026-11) recovers Completed via calendar fallback even if manual cols were pre-deleted
  assert.ok(!recon.unmatchedSheetInstallments.some(x => x.instId === "I-00017"), "Past I-00017 must recover Completed via calendar fallback");

  // Also verify post-migration recovery: once INST_PROGRESS exists, even deleting J/K preserves exact 19542.07 blocked total
  const ss2 = new MockSpreadsheet();
  setActiveSpreadsheet(ss2);
  populateSpreadsheetFromFixture(ss2, realFixture, CONFIG);
  sandbox.updateLiveDashboard();
  const inst2 = ss2.getSheetByName(CONFIG.SHEETS.INSTALLMENTS);
  inst2.deleteColumns(10, 2);
  sandbox.updateLiveDashboard();
  const recon2 = sandbox.reconcileActiveStatement(ss2, tz);
  assert.strictEqual(inst2._getCell(1, 10), "Bank Progress");
  assert.strictEqual(inst2._getCell(1, 11), "Bank Status");
  assert.strictEqual(inst2._getCell(1, 12), "_Row_ID");
  assert.strictEqual(Number(recon2.cardBal.totalBlockedInstallments.toFixed(2)), 19542.07);
}

// ==========================================
// Test 7 & 8: Cross-Month Completion Without H6 & handlePaymentDateEdit Non-Clobber
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const inst = ss.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
  inst.getRange(1, 1, 1, 11).setValues([["#", "Start Date", "Description", "Total Price", "Months", "C6", "C7", "EMI", "Payer", "Payments Made", "Status"]]);
  inst.getRange(2, 1, 2, 11).setValues([
    [1, new Date(2026, 4, 10, 12, 0, 0), "Gym Plan", 600, 6, 0, 0, 100, "Mido", 5, "Ongoing"],
    [2, new Date(2026, 7, 10, 12, 0, 0), "Phone Plan", 1200, 6, 0, 0, 200, "Mai", 2, "Ongoing"]
  ]);

  sandbox.writeStatementToSheet(ss, {
    statementDate: "31-10-2026",
    dueDate: "27-11-2026",
    creditLimit: 100000,
    openingBalance: 0,
    closingBalance: 300,
    totalDebit: 300,
    totalCredit: 0,
    card: "1234"
  }, [
    { txDate: "31-10-2026", postDate: "31-10-2026", desc: "GYM PLAN 6 OF 6", type: "INSTALLMENT", amount: 100 },
    { txDate: "31-10-2026", postDate: "31-10-2026", desc: "PHONE PLAN 3 OF 6", type: "INSTALLMENT", amount: 200 }
  ]);

  sandbox.updateLiveDashboard();

  const gymId = inst._getCell(2, 12);
  const phoneId = inst._getCell(3, 12);
  let dec = sandbox.loadDecisions(ss);
  assert.strictEqual(dec.instProgress[gymId].paid, 5, "While H6 blank, Gym 6 OF 6 stores paid=5, billed=6");
  assert.strictEqual(dec.instProgress[gymId].billed, 6);
  assert.strictEqual(dec.instProgress[phoneId].paid, 2);

  // Test 8: handlePaymentDateEdit must save BOTH CYCLE_PAYMENT and updated INST_PROGRESS without clobbering
  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  ar.getRange(6, 8).setValue("2026-11-25");
  sandbox.handlePaymentDateEdit(ar, ar.getRange(6, 8), tz, true);

  dec = sandbox.loadDecisions(ss);
  assert.ok(dec.paymentRecords["2026-11"] && dec.paymentRecords["2026-11"].dateStr === "2026-11-25", "CYCLE_PAYMENT must persist");
  assert.strictEqual(dec.instProgress[phoneId].paid, 3, "INST_PROGRESS must NOT be clobbered by saveDecisions in handlePaymentDateEdit");
  assert.strictEqual(dec.instProgress[gymId].completed, true, "Gym becomes Completed when H6 is paid");

  // Clear H6 -> reverts both CYCLE_PAYMENT and INST_PROGRESS cleanly
  ar.getRange(6, 8).setValue("");
  sandbox.handlePaymentDateEdit(ar, ar.getRange(6, 8), tz, true);

  dec = sandbox.loadDecisions(ss);
  assert.strictEqual(dec.paymentRecords["2026-11"], undefined, "Clearing H6 must remove CYCLE_PAYMENT");
  assert.strictEqual(dec.instProgress[phoneId].paid, 2, "Clearing H6 must revert Phone INST_PROGRESS to paid=2");
  assert.strictEqual(dec.instProgress[gymId].completed, false, "Clearing H6 must revert Gym to unpaid (billed=6)");

  // Test 7 continued: Import next cycle (2026-12) WITHOUT Gym (since 6/6 finished even though H6 was never filled!)
  sandbox.writeStatementToSheet(ss, {
    statementDate: "30-11-2026",
    dueDate: "27-12-2026",
    creditLimit: 100000,
    openingBalance: 0,
    closingBalance: 200,
    totalDebit: 200,
    totalCredit: 0,
    card: "1234"
  }, [
    { txDate: "30-11-2026", postDate: "30-11-2026", desc: "PHONE PLAN 4 OF 6", type: "INSTALLMENT", amount: 200 }
  ]);

  sandbox.updateLiveDashboard();
  const decRecon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(decRecon.unmatchedSheetInstallments.length, 0, "Completed Gym Plan from 2026-11 must NOT appear as Sheet Only in 2026-12");
  assert.ok(String(inst._getCell(2, 11)).includes('"Completed"'), "Gym Col K formula must resolve Completed in 2026-12");
  assert.strictEqual(Number(decRecon.cardBal.totalBlockedInstallments.toFixed(2)), 600.00, "Only Phone Plan (3 remaining * 200 = 600) should be blocked");

  // Test 9: Export anonymized fixture from formula-based Col J/K + INST_PROGRESS & replay losslessly
  const exported = sandbox.buildAnonymizedFixture(ss, tz);
  const serializedExp = JSON.stringify(exported).toLowerCase();
  assert.ok(exported.decisions.some(d => d[0] === "INST_PROGRESS"), "Exported fixture must preserve INST_PROGRESS decisions");
  ["gym", "phone", "mido", "mai"].forEach(tok => {
    assert.ok(!serializedExp.includes(tok), `Exported fixture must not leak token "${tok}"`);
  });

  const replayEnv3 = createGasSandbox();
  const ssReplay3 = new MockSpreadsheet();
  replayEnv3.setActiveSpreadsheet(ssReplay3);
  populateSpreadsheetFromFixture(ssReplay3, exported, replayEnv3.CONFIG);
  const replayRecon3 = replayEnv3.sandbox.reconcileActiveStatement(ssReplay3, tz);
  const replayLedger3 = replayEnv3.sandbox.buildDebtLedger(ssReplay3, tz, replayRecon3);
  assert.strictEqual(replayRecon3.unmatchedSheetInstallments.length, 0, "Replayed 2026-12 fixture must keep Gym Completed (not Sheet Only)");
  assert.strictEqual(Number(replayRecon3.cardBal.totalBlockedInstallments.toFixed(2)), 600.00);
  assert.strictEqual(Number(replayLedger3.installmentsTotal.toFixed(2)), 200.00);
}

console.log("installments_link.test.js: all Phase 3 installment link, cross-month reuse, column replacement, emoji, and replay assertions passed.");
