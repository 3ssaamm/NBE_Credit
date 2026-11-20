const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { createGasSandbox, MockSpreadsheet } = require("./helpers/gas_sandbox");

const tz = "Africa/Cairo";

// ==========================================
// Test 1: Zero-Config People Auto-Discovery & Optional Settings Absence
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 350, "Total Debit:", 350],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 350]
  ]);
  stmt.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Post", "Description", "Type", "Amount", "Status"]]);
  stmt.getRange(7, 1, 2, 7).setValues([
    [1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "SHOP ALPHA", "DEBIT", 300, ""],
    [2, new Date(2026, 9, 31, 12, 0, 0), new Date(2026, 9, 31, 12, 0, 0), "STAMP TAX", "FEE", 50, ""]
  ]);

  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, 2, 5).setValues([
    [1, new Date(2026, 9, 10, 0, 0, 0), "Shop Alpha", "Zara", 300],
    [2, new Date(2026, 10, 5, 0, 0, 0), "Future Book", "Omar", 120]
  ]);

  const inst = ss.insertSheet(CONFIG.SHEETS.INSTALLMENTS);
  inst.getRange(1, 1, 1, 9).setValues([["#", "Start", "Desc", "Total", "Months", "C6", "C7", "EMI", "Payer"]]);
  inst.getRange(2, 1, 1, 9).setValues([
    [1, new Date(2026, 10, 10, 12, 0, 0), "Laptop", 1200, 12, 0, 0, 100, "Lina"]
  ]);

  const decSheet = sandbox.getDecisionsSheet(ss);
  decSheet.getRange(2, 1, 1, 7).setValues([
    ["STMT_ASSIGN", "S-99999", "Sami", "", "user", true, "2026-11-19"]
  ]);

  assert.strictEqual(ss.getSheetByName(CONFIG.SHEETS.SETTINGS), null, "Settings sheet must stay optional and NOT auto-created");
  assert.strictEqual(sandbox.getDefaultFeePayer(ss), "", "Default fee payer must be blank until configured");

  const people = Array.from(sandbox.getAllUniquePayers(ss));
  assert.deepStrictEqual(people, ["Lina", "Omar", "Sami", "Zara"], "Must auto-discover people cleanly without hardcoded defaults");

  const recon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(recon.stmtFees[0].assignedPayer, "", "STAMP TAX must remain unassigned when no Default Fee Payer is set");
  const ledger = sandbox.buildDebtLedger(ss, tz, recon);
  assert.strictEqual(ledger.unresolvedFeesCount, 1, "Unassigned fee must surface in Needs Pairing as unresolved");
}

// ==========================================
// Test 2: Optional Settings Sheet Overrides (Default Payer, Credit Limit, Name Aliases)
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 50, "Total Debit:", 50],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 50]
  ]);
  stmt.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Post", "Description", "Type", "Amount", "Status"]]);
  stmt.getRange(7, 1, 1, 7).setValues([
    [1, new Date(2026, 9, 31, 12, 0, 0), new Date(2026, 9, 31, 12, 0, 0), "STAMP TAX", "FEE", 50, ""]
  ]);

  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, 2, 5).setValues([
    [1, new Date(2026, 10, 5, 0, 0, 0), "Snack", "me", 10],
    [2, new Date(2026, 10, 6, 0, 0, 0), "Coffee", "omr + z", 20]
  ]);

  const settingsSheet = sandbox.getOrCreateSettingsSheet(ss);
  settingsSheet.getRange("B6").setValue("Zara");
  settingsSheet.getRange("B7").setValue(150000);
  settingsSheet.getRange(6, 4, 2, 2).setValues([
    ["me, myself, z", "Zara"],
    ["omr", "Omar"]
  ]);
  sandbox.invalidateSettingsCache_();

  assert.strictEqual(sandbox.getDefaultFeePayer(ss), "Zara");
  assert.strictEqual(sandbox.getEffectiveCreditLimit(ss), 150000);
  assert.strictEqual(sandbox.normalizePersonName("MYSELF"), "Zara");
  assert.strictEqual(sandbox.formatPayerCell("omr + me"), "Omar + Zara");

  sandbox.updateLiveDashboard();
  assert.strictEqual(tx._getCell(2, 4), "Zara", "Name alias 'me' must sanitize to Zara");
  assert.strictEqual(tx._getCell(3, 4), "Omar + Zara", "Name aliases 'omr + z' must sanitize cleanly");

  const recon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(recon.stmtFees[0].assignedPayer, "Zara", "STAMP TAX must auto-assign to configured Default Fee Payer");
  assert.strictEqual(recon.cardBal.creditLimit, 150000, "calculateCardBalance must honor configured Credit Limit");
}

// ==========================================
// Test 3: Bi-Directional Merchant Alias Sync Between Settings & Learned _Decisions
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 350, "Total Debit:", 350],
    ["Total Credit:", 0, "Card Number:", "1234", "Net:", 350]
  ]);
  stmt.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Post", "Description", "Type", "Amount", "Status"]]);
  stmt.getRange(7, 1, 2, 7).setValues([
    [1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "PAYMOB*MERCHANTX", "DEBIT", 200, ""],
    [2, new Date(2026, 9, 12, 12, 0, 0), new Date(2026, 9, 12, 12, 0, 0), "FAWRY*COFFEENOOK", "DEBIT", 150, ""]
  ]);

  const tx = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  tx.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  tx.getRange(2, 1, 2, 5).setValues([
    [1, new Date(2026, 9, 10, 0, 0, 0), "MyShop", "Alice", 200],
    [2, new Date(2026, 9, 12, 0, 0, 0), "BrewPlace", "Bob", 150]
  ]);

  // 3a. Add merchant alias in Settings Col G:I -> onEdit syncs to _Decisions & auto-matches PAYMOB*MERCHANTX
  const settingsSheet = sandbox.getOrCreateSettingsSheet(ss);
  settingsSheet.getRange(6, 7, 1, 3).setValues([["merchantx", "myshop", "Manual alias"]]);
  sandbox.onEdit({ range: settingsSheet.getRange(6, 7, 1, 1) });

  let dec = sandbox.loadDecisions(ss);
  assert.ok(dec.merchantAliases["merchantx::myshop"], "Settings Col G:I edit must sync alias to _Decisions");

  sandbox.updateLiveDashboard();
  let recon = sandbox.reconcileActiveStatement(ss, tz);
  assert.strictEqual(recon.stmtDebits[0].matched, true, "MERCHANTX <-> MyShop must auto-match via Settings alias");
  assert.strictEqual(recon.missingInSheet.length, 1, "Unlearned COFFEENOOK stays in Needs Pairing");

  // 3b. Learn COFFEENOOK <-> BrewPlace via 1-click pairing -> auto-fills Settings Col G:I!
  const ar = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  let rowM1 = -1;
  for (let r = 12; r <= ar.getLastRow(); r++) {
    if (ar._getCell(r, 1) === "M-1") { rowM1 = r; break; }
  }
  assert.ok(rowM1 > 0);
  ar._setCell(rowM1, 6, recon.missingInSheet[0].suggestedPairs[0].dropdownLabel);
  sandbox.handleActiveReconEdit({ range: ar.getRange(rowM1, 6) });

  const settingsAliases = settingsSheet.getRange(6, 7, 2, 2).getValues();
  assert.ok(
    settingsAliases.some(r => r[0] === "coffeenook" && r[1] === "brewplace"),
    "Learned MERCHANT_ALIAS must auto-populate into Settings Col G:I"
  );

  // 3c. Deleting alias from Settings + onEdit removes it from _Decisions
  settingsSheet.getRange(7, 7, 1, 3).setValues([["", "", ""]]);
  sandbox.onEdit({ range: settingsSheet.getRange(7, 7, 1, 1) });
  dec = sandbox.loadDecisions(ss);
  assert.strictEqual(Object.keys(dec.merchantAliases).length, 1, "Deleting a row in Settings Col G:I must remove it from _Decisions");
}

// ==========================================
// Test 4: Static Audit Across Code.js + ALL src/*.js
// ==========================================
{
  const rootDir = path.resolve(__dirname, "..");
  const filesToAudit = [path.join(rootDir, "Code.js")];
  const srcDir = path.join(rootDir, "src");
  fs.readdirSync(srcDir)
    .filter(f => /\.(js|gs)$/i.test(f))
    .sort()
    .forEach(f => filesToAudit.push(path.join(srcDir, f)));

  const forbiddenTerms = [
    "MERCHANT_ALIAS_GROUPS",
    "getMerchantAliasKey",
    "Mido", "Abdo", "Muhanad", "Mohanad", "Zoza", "Zozza",
    "Nourween", "Hager", "Nevin", "Neveen", "Abdelrahman",
    "Mai", "Dad", "Mum",
    "bazlykwa", "bazlikwa", "basilico", "nutopia", "sutherland",
    "google one", "cavalli", "fathalla", "buffalo", "caribou",
    "instashop", "shein", "elabd", "viacaffe", "asteria"
  ];

  filesToAudit.forEach(filePath => {
    const src = fs.readFileSync(filePath, "utf8");
    const rel = path.basename(filePath);
    forbiddenTerms.forEach(term => {
      const re = new RegExp(`\\b${term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "i");
      assert.ok(
        !re.test(src),
        `Static audit failed: "${rel}" still contains hardcoded term "${term}"`
      );
    });
  });
}

console.log("config.test.js: all auto-discovery, Settings override, bi-directional merchant alias, and static cleanup assertions passed.");
