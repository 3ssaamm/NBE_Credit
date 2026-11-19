const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { createGasSandbox, populateSpreadsheetFromFixture, MockSpreadsheet } = require("./helpers/gas_sandbox");

const { sandbox } = createGasSandbox();
const tz = "Africa/Cairo";

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

// ==========================================
// Phase 4 Tests — Mathematical Tie-Out, Health Gate, "Confirm Type", & Idempotent Re-Import
// ==========================================

// Test 1: Clean Mathematical Tie-Out & Green Row 5 Health Banner + Full Bill Generation
{
  const env = createGasSandbox();
  const ss = new MockSpreadsheet();
  env.setActiveSpreadsheet(ss);

  const cleanText = [
    "National Bank of Egypt",
    "Card Number: 5264****1234",
    "Statement Date: 31-10-2026",
    "Due Date: 27-11-2026\tCredit Limit: 100,000.00",
    "Opening Balance: 10,000.00\tClosing Balance: 355.01",
    "Total of Credit: 10,009.00\tTotal of Debit: 364.01",
    "Transaction Date\tPosting Date\tDescription\tDebit\tCredit",
    "05/10/2026\t06/10/2026\tINSTAPAY PAYMENT - THANK YOU\t\t10,000.00",
    "12/10/2026\t13/10/2026\tMICROSOFT*MICROSOFT 365\t\t9.00",
    "20/10/2026\t21/10/2026\tCOFFEE ROASTERS\t100.00\t",
    "22/10/2026\t23/10/2026\tBOOK STORE\t100.00\t",
    "25/10/2026\t26/10/2026\tKETTLE 10 OF 12\t123.00\t",
    "STAMP DUTY 41.01"
  ].join("\n");

  const res = env.sandbox.parseNBEStatementText(cleanText);
  assert.strictEqual(res.tieOut.passed, true, `Clean statement must pass tie-out: ${res.tieOut.reasons.join("; ")}`);
  assert.deepStrictEqual(Object.assign({}, res.tieOut.counts), {
    DEBIT: 2,
    INSTALLMENT: 1,
    FEE: 1,
    CREDIT: 2,
    total: 6,
    ambiguous: 0
  });
  assert.strictEqual(res.tieOut.closingDiff, 0);
  assert.strictEqual(res.tieOut.debitDiff, 0);
  assert.strictEqual(res.tieOut.creditDiff, 0);

  env.sandbox.writeStatementToSheet(ss, res.meta, res.transactions, { sourceFile: "october.pdf" });
  const dashRes = env.sandbox.updateLiveDashboard();
  assert.strictEqual(dashRes.blocked, false, "Clean tie-out must allow dashboard & bill generation");

  const stmtSheet = ss.getSheetByName(env.CONFIG.SHEETS.BANK_STATEMENT);
  assert.ok(String(stmtSheet._getCell(5, 1)).includes("Tie-Out Passed"), "Row 5 banner must show green Tie-Out Passed");
  assert.ok(ss.getSheetByName(env.CONFIG.SHEETS.MEMBER_BREAKDOWN), "Member Breakdown must render on passed tie-out");
  assert.ok(ss.getSheetByName(env.CONFIG.SHEETS.ACTIVE_RECON)._getCell(1, 1).includes("Active Bill"), "Active Bill must render");
}

// Test 2: NOON 2835.83 Hack Removed -> Half-Diff "Confirm Type" + Red Health Gate + Recovery via Col E onEdit
{
  const env = createGasSandbox();
  const ss = new MockSpreadsheet();
  env.setActiveSpreadsheet(ss);

  const noonText = [
    "National Bank of Egypt",
    "Card Number: 5264****1234",
    "Statement Date: 30-09-2026",
    "Due Date: 25-10-2026\tCredit Limit: 100,000.00",
    "Opening Balance: 5,000.00\tClosing Balance: 2,264.17",
    "Total of Credit: 2,835.83\tTotal of Debit: 100.00",
    "10/09/2026 11/09/2026 COFFEE SHOP ALEX 100.00",
    "15/09/2026 16/09/2026 NOON.COM CAIRO EGY 2,835.83"
  ].join("\n");

  const parsedNoon = env.sandbox.parseNBEStatementText(noonText);
  assert.strictEqual(parsedNoon.transactions[1].type, "DEBIT", "Must NEVER silently auto-flip NOON 2835.83 via hardcoded hack");
  assert.strictEqual(parsedNoon.tieOut.passed, false, "Mismatched closing/totals must fail tie-out");
  assert.strictEqual(parsedNoon.tieOut.closingDiff, 5671.66);
  assert.strictEqual(parsedNoon.transactions[1].needsTypeConfirm, true, "Half-diff math must flag NOON for Confirm Type");
  assert.strictEqual(parsedNoon.tieOut.counts.ambiguous, 1);

  env.sandbox.writeStatementToSheet(ss, parsedNoon.meta, parsedNoon.transactions, { sourceFile: "Sep_Private_Mido.pdf" });
  const blockedRes = env.sandbox.updateLiveDashboard();
  assert.strictEqual(blockedRes.blocked, true, "Failed tie-out must BLOCK bill generation");
  assert.strictEqual(ss.getSheetByName(env.CONFIG.SHEETS.MEMBER_BREAKDOWN), null, "Member Breakdown must NOT generate while tie-out fails");
  assert.strictEqual(ss.getSheetByName(env.CONFIG.SHEETS.HISTORY_FORECAST), null, "Monthly History must NOT generate while tie-out fails");

  const stmtSheet = ss.getSheetByName(env.CONFIG.SHEETS.BANK_STATEMENT);
  assert.ok(String(stmtSheet._getCell(5, 1)).includes("TIE-OUT FAILED"), "Bank Statement Row 5 must show red failure banner");
  assert.ok(String(stmtSheet._getCell(8, 7)).includes("Confirm Type"), "Row 8 (NOON) Col G must prompt Confirm Type");

  const arErr = ss.getSheetByName(env.CONFIG.SHEETS.ACTIVE_RECON);
  assert.ok(String(arErr._getCell(1, 1)).includes("FAILED (Bill Blocked)"), "Active Bill tab must display red Health error screen");
  assert.strictEqual(arErr._getCell(13, 1), 8, "Confirm Type table must point to sheet Row 8");
  assert.strictEqual(arErr._getCell(13, 5), 2835.83, "Confirm Type table must surface 2835.83 line");

  // User resolves Row 8 Col E from DEBIT -> CREDIT
  stmtSheet._setCell(8, 5, "CREDIT");
  env.sandbox.onEdit({ range: stmtSheet.getRange(8, 5) });

  const decAfterFix = env.sandbox.loadDecisions(ss);
  const noonId = stmtSheet._getCell(8, 8);
  assert.strictEqual(noonId, "S-00002");
  assert.ok(decAfterFix.stmtTypeConfirm[noonId] && decAfterFix.stmtTypeConfirm[noonId].target === "CREDIT", "Col E fix must persist in STMT_TYPE_CONFIRM");
  assert.ok(String(stmtSheet._getCell(5, 1)).includes("Tie-Out Passed"), "Row 5 banner must turn green after Col E fix");
  assert.ok(ss.getSheetByName(env.CONFIG.SHEETS.MEMBER_BREAKDOWN), "Member Breakdown must generate once tie-out passes");

  // Test 5 (continued on same ss): Idempotent Re-Import preserves S-IDs & STMT_TYPE_CONFIRM even with OCR text shifts
  const noonReimportText = [
    "National Bank of Egypt",
    "Card Number: 5264****1234",
    "Statement Date: 30-09-2026",
    "Due Date: 25-10-2026\tCredit Limit: 100,000.00",
    "Opening Balance: 5,000.00\tClosing Balance: 2,264.17",
    "Total of Credit: 2,835.83\tTotal of Debit: 100.00",
    "10/09/2026 11/09/2026 COFFEE SHOP - ALEX. 100.00",
    "15/09/2026 16/09/2026 NOON . COM - CAIRO EGY 2,835.83"
  ].join("\n");
  const reparsed = env.sandbox.parseNBEStatementText(noonReimportText);
  const writeRes2 = env.sandbox.writeStatementToSheet(ss, reparsed.meta, reparsed.transactions);
  assert.strictEqual(writeRes2.preservedCount, 2, "Re-importing same statement must preserve both S-IDs via lineIndex + amount");
  assert.deepStrictEqual(Array.from(writeRes2.assignedIds), ["S-00001", "S-00002"]);
  assert.strictEqual(writeRes2.tieOut.passed, true, "Previously confirmed STMT_TYPE_CONFIRM must reapply on re-import so tie-out stays GREEN");

  // Test 7: PII-Safe Export with META_IMPORT & STMT_TYPE_CONFIRM
  const txSheet = ss.insertSheet(env.CONFIG.SHEETS.TRANSACTIONS);
  txSheet.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  txSheet.getRange(2, 1, 1, 5).setValues([[1, new Date(2026, 8, 10, 0, 0, 0), "Coffee Spot", "Mido", 100.00]]);
  const exported = env.sandbox.buildAnonymizedFixture(ss, tz);
  const expStr = JSON.stringify(exported).toLowerCase();
  assert.ok(exported.decisions.some(d => d[0] === "META_IMPORT"), "Exported fixture must include META_IMPORT");
  assert.ok(exported.decisions.some(d => d[0] === "STMT_TYPE_CONFIRM"), "Exported fixture must include STMT_TYPE_CONFIRM");
  assert.ok(!expStr.includes("sep_private_mido"), "META_IMPORT sourceFile must be stripped in export");
  assert.ok(!expStr.includes("noon"), "Export must not leak NOON merchant name");

  // Replay exported fixture cleanly in a fresh sandbox
  const replayEnv = createGasSandbox();
  const ssReplay = new MockSpreadsheet();
  replayEnv.setActiveSpreadsheet(ssReplay);
  populateSpreadsheetFromFixture(ssReplay, exported, replayEnv.CONFIG);
  const replayHealth = replayEnv.sandbox.evaluateBankStatementHealth(ssReplay, tz);
  assert.strictEqual(replayHealth.passed, true, "Replayed fixture META_IMPORT health check must pass cleanly");

  // Import a NEW month (November 2026) -> fresh monotonic S-IDs (S-00003..S-00004), no leak from S-00002
  const novText = [
    "National Bank of Egypt",
    "Card Number: 5264****1234",
    "Statement Date: 31-10-2026",
    "Due Date: 27-11-2026\tCredit Limit: 100,000.00",
    "Opening Balance: 2,264.17\tClosing Balance: 3,078.17",
    "Total of Credit: 0.00\tTotal of Debit: 814.00",
    "15/10/2026 16/10/2026 ONLINE STORE EGY 814.00"
  ].join("\n");
  const novParsed = env.sandbox.parseNBEStatementText(novText);
  const novWrite = env.sandbox.writeStatementToSheet(ss, novParsed.meta, novParsed.transactions);
  assert.strictEqual(novWrite.preservedCount, 0, "Different statement date must NOT reuse old S-IDs");
  assert.deepStrictEqual(Array.from(novWrite.assignedIds), ["S-00003"]);
  assert.strictEqual(stmtSheet._getCell(7, 5), "DEBIT", "Previous statement's STMT_TYPE_CONFIRM must never leak onto fresh S-00003");
}

// Test 3: Two-Line Flipped Pair (2835.83 + 814.00) Detected Without Hardcoding
{
  const env = createGasSandbox();
  const pairText = [
    "National Bank of Egypt",
    "Card Number: 5264****1234",
    "Statement Date: 30-09-2026",
    "Due Date: 25-10-2026\tCredit Limit: 100,000.00",
    "Opening Balance: 10,000.00\tClosing Balance: 6,550.17",
    "Total of Credit: 3,649.83\tTotal of Debit: 200.00",
    "05/09/2026 06/09/2026 CAFE STORE 200.00",
    "15/09/2026 16/09/2026 VENDOR A CAIRO EGY 2,835.83",
    "20/09/2026 21/09/2026 VENDOR B CAIRO EGY 814.00"
  ].join("\n");

  const parsedPair = env.sandbox.parseNBEStatementText(pairText);
  assert.strictEqual(parsedPair.tieOut.passed, false);
  assert.strictEqual(parsedPair.tieOut.closingDiff, 7299.66);
  assert.strictEqual(parsedPair.transactions[0].needsTypeConfirm, false, "Unrelated 200 DEBIT must NOT be flagged");
  assert.strictEqual(parsedPair.transactions[1].needsTypeConfirm, true, "First line of flipped pair (2835.83) must be flagged");
  assert.strictEqual(parsedPair.transactions[2].needsTypeConfirm, true, "Second line of flipped pair (814.00) must be flagged");
}

// Test 4: Ambiguous Keyword & Dual-Column Tabbed Row Require Manual Type Confirmation Even If Math Sums
{
  const env = createGasSandbox();
  const ss = new MockSpreadsheet();
  env.setActiveSpreadsheet(ss);

  const ambText = [
    "National Bank of Egypt",
    "Card Number: 5264****1234",
    "Statement Date: 31-10-2026",
    "Due Date: 27-11-2026\tCredit Limit: 100,000.00",
    "Opening Balance: 0.00\tClosing Balance: 500.00",
    "Total of Credit: 0.00\tTotal of Debit: 500.00",
    "10/10/2026 11/10/2026 WALLET TRANSFER FEES 500.00"
  ].join("\n");

  const parsedAmb = env.sandbox.parseNBEStatementText(ambText);
  assert.strictEqual(parsedAmb.tieOut.mathPassed, true, "Math matches as DEBIT");
  assert.strictEqual(parsedAmb.transactions[0].needsTypeConfirm, true, "Ambiguous TRANSFER keyword on non-tabbed line must trigger Confirm Type");
  assert.strictEqual(parsedAmb.tieOut.passed, false, "Overall tie-out must block until ambiguous line is confirmed");

  env.sandbox.writeStatementToSheet(ss, parsedAmb.meta, parsedAmb.transactions);
  assert.strictEqual(env.sandbox.updateLiveDashboard().blocked, true, "Ambiguous line must block bill until Col E confirmed");

  const stmtSheet = ss.getSheetByName(env.CONFIG.SHEETS.BANK_STATEMENT);
  stmtSheet._setCell(7, 5, "DEBIT");
  env.sandbox.onEdit({ range: stmtSheet.getRange(7, 5) });
  assert.strictEqual(env.sandbox.evaluateBankStatementHealth(ss, tz).passed, true, "Confirming DEBIT in Col E clears Confirm Type gate");

  // Also check tabbed row with BOTH Debit & Credit non-zero
  const dualTabText = [
    "Statement Date: 31-10-2026",
    "Due Date: 27-11-2026\tCredit Limit: 100,000.00",
    "Opening Balance: 0.00\tClosing Balance: 150.00",
    "Total of Credit: 0.00\tTotal of Debit: 150.00",
    "10/10/2026\t11/10/2026\tDUAL COLUMN SHOP\t150.00\t50.00"
  ].join("\n");
  const dualParsed = env.sandbox.parseNBEStatementText(dualTabText);
  assert.strictEqual(dualParsed.transactions[0].needsTypeConfirm, true, "Both Debit and Credit non-zero in tabbed row must require confirmation");
}

// Test 6: Static Guard — Hardcoded NOON/AMAZON Amount Hacks Must Not Exist in Code.js or src/Parser.js
{
  const rootDir = path.resolve(__dirname, "..");
  const codeSrc = fs.readFileSync(path.join(rootDir, "Code.js"), "utf8");
  const parserSrc = fs.readFileSync(path.join(rootDir, "src", "Parser.js"), "utf8");
  [codeSrc, parserSrc].forEach(src => {
    assert.ok(!src.includes("isKnownStatementCredit"), "Hardcoded isKnownStatementCredit hack must be deleted");
    assert.ok(!src.includes("isKnownCredit"), "Hardcoded isKnownCredit hack must be deleted");
    assert.ok(!src.includes("2835"), "Hardcoded 2835.83 hack must be deleted");
    assert.ok(!src.includes("814.00"), "Hardcoded 814.00 hack must be deleted");
  });
}

console.log("parser.test.js: all PDF statement text parsing, tie-out validation, Confirm Type gating, and idempotent re-import assertions passed.");
