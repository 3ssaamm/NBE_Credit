const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { createGasSandbox, populateSpreadsheetFromFixture, MockSpreadsheet } = require("./helpers/gas_sandbox");

const tz = "Africa/Cairo";

function createDriveMock() {
  let idCounter = 1;

  function makeIterator(arr) {
    let idx = 0;
    return {
      hasNext: () => idx < arr.length,
      next: () => arr[idx++]
    };
  }

  function makeFolder(name) {
    const id = `folder-${idCounter++}`;
    const children = [];
    const files = [];
    return {
      name,
      id,
      children,
      files,
      getId: () => id,
      getName: () => name,
      getUrl: () => `https://drive.google.com/drive/folders/${id}`,
      getFoldersByName: target => makeIterator(children.filter(f => f.name === target && !f.trashed)),
      createFolder: childName => {
        const child = makeFolder(childName);
        children.push(child);
        return child;
      },
      getFilesByName: fileName => makeIterator(files.filter(f => f.name === fileName && !f.trashed)),
      createFile: blob => {
        const fid = `file-${idCounter++}`;
        const fname = blob && blob.name ? blob.name : "untitled.pdf";
        const fileObj = {
          id: fid,
          name: fname,
          blob,
          trashed: false,
          getId: () => fid,
          getName: () => fname,
          getUrl: () => `https://drive.google.com/file/d/${fid}/view`,
          setTrashed: flag => { fileObj.trashed = Boolean(flag); }
        };
        files.push(fileObj);
        return fileObj;
      }
    };
  }

  const root = makeFolder("ROOT");
  return {
    root,
    DriveApp: {
      getRootFolder: () => root
    }
  };
}

// ==========================================
// Test 1: Exact Mathematical Fidelity & WhatsApp Formatting on Real Fixture (2026-11-real.json)
// ==========================================
{
  const realFixture = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "2026-11-real.json"), "utf8"));
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);
  populateSpreadsheetFromFixture(ss, realFixture, CONFIG);

  const recon = sandbox.reconcileActiveStatement(ss, tz);
  const ledger = sandbox.buildDebtLedger(ss, tz, recon);

  Object.keys(realFixture.expectedShares).forEach(member => {
    const expected = realFixture.expectedShares[member];
    const data = sandbox.buildMemberStatementData(ss, tz, member, { recon, ledger });
    assert.strictEqual(data.person, member);
    assert.strictEqual(data.totals.purchases, expected.purchases, `${member} purchases mismatch`);
    assert.strictEqual(data.totals.installments, expected.installments, `${member} installments mismatch`);
    assert.strictEqual(data.totals.fees, expected.fees, `${member} fees mismatch`);
    assert.strictEqual(data.totals.unrecorded, expected.unrecorded, `${member} unrecorded mismatch`);
    assert.strictEqual(data.totals.refunds, expected.refunds, `${member} refunds mismatch`);
    assert.strictEqual(data.totals.netTotal, expected.netTotal, `${member} netTotal mismatch`);
  });

  // Member with purchases, installments, unrecorded items, and refunds (P5)
  const p5Data = sandbox.buildMemberStatementData(ss, tz, "P5", { recon, ledger });
  const p5Text = sandbox.formatWhatsAppStatement(p5Data, tz);
  assert.ok(p5Text.includes("💳 *NBE Credit Card — November 2026*"));
  assert.ok(p5Text.includes("👤 *Member:* P5"));
  assert.ok(p5Text.includes("📅 *Due Date:* November 27, 2026"));
  assert.ok(p5Text.includes("💰 *Total Due: 6,073.88 EGP*"));
  assert.ok(p5Text.includes("🛍️ *Purchases (504.96 EGP):*"));
  assert.ok(p5Text.includes("📅 *Installments (1,133.04 EGP):*"));
  assert.ok(p5Text.includes("🧾 *Unrecorded / Statement Items (7,451.88 EGP):*"));
  assert.ok(p5Text.includes("↩️ *Refunds / Credits (-3,016.00 EGP):*"));
  assert.ok(!p5Text.includes("Bank Fees & Taxes"), "Empty fees bucket must be omitted");
  assert.ok(!p5Text.includes("Draft:"), "Resolved statement must not include Draft warning");
  assert.ok(p5Text.includes("✅ *Net Total Due: 6,073.88 EGP*"));

  // Member with zero activity (P4)
  const p4Data = sandbox.buildMemberStatementData(ss, tz, "P4", { recon, ledger });
  const p4Text = sandbox.formatWhatsAppStatement(p4Data, tz);
  assert.ok(p4Text.includes("💰 *Total Due: 0.00 EGP*"));
  assert.ok(p4Text.includes("✨ _No charges billed in this cycle._"));
  assert.ok(!p4Text.includes("🛍️ *Purchases"));
  assert.ok(!p4Text.includes("📅 *Installments"));
}

// ==========================================
// Test 2: Splits, Grouped Purchases, Unlinked Installments, Fees & Draft Warning
// ==========================================
{
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);

  const stmt = ss.insertSheet(CONFIG.SHEETS.BANK_STATEMENT);
  stmt.getRange(2, 1, 3, 6).setValues([
    ["Statement Date:", new Date(2026, 9, 31, 12, 0, 0), "Due Date:", new Date(2026, 10, 27, 12, 0, 0), "Credit Limit:", 100000],
    ["Opening Balance:", 0, "Closing Balance:", 1090, "Total Debit:", 1140],
    ["Total Credit:", 50, "Card Number:", "1234", "Net:", 1090]
  ]);
  stmt.getRange(6, 1, 1, 7).setValues([["#", "Tx Date", "Post Date", "Description", "Type", "Amount", "Status"]]);
  stmt.getRange(7, 1, 6, 7).setValues([
    [1, new Date(2026, 9, 10, 12, 0, 0), new Date(2026, 9, 10, 12, 0, 0), "MARKET SPLIT", "DEBIT", 200, ""],
    [2, new Date(2026, 9, 12, 12, 0, 0), new Date(2026, 9, 12, 12, 0, 0), "STORE GROUP", "DEBIT", 300, ""],
    [3, new Date(2026, 9, 14, 12, 0, 0), new Date(2026, 9, 14, 12, 0, 0), "PHONE PLAN 3 OF 6", "INSTALLMENT", 400, ""],
    [4, new Date(2026, 9, 31, 12, 0, 0), new Date(2026, 9, 31, 12, 0, 0), "STAMP DUTY", "FEE", 40, ""],
    [5, new Date(2026, 9, 16, 12, 0, 0), new Date(2026, 9, 16, 12, 0, 0), "STORE REFUND", "CREDIT", 50, ""],
    [6, new Date(2026, 9, 18, 12, 0, 0), new Date(2026, 9, 18, 12, 0, 0), "MYSTERY LINE", "DEBIT", 200, ""]
  ]);

  const txSheet = ss.insertSheet(CONFIG.SHEETS.TRANSACTIONS);
  txSheet.getRange(1, 1, 1, 5).setValues([["#", "Date", "Description", "Person", "Amount"]]);
  txSheet.getRange(2, 1, 3, 5).setValues([
    [1, new Date(2026, 9, 10, 12, 0, 0), "Market Split", "P1 + P2", 200],
    [2, new Date(2026, 9, 12, 12, 0, 0), "Store Group A", "P1", 120],
    [3, new Date(2026, 9, 12, 12, 0, 0), "Store Group B", "P1", 180]
  ]);

  const decisions = sandbox.loadDecisions(ss);
  const idInfo = sandbox.ensureAllSheetRowIds(ss, decisions);
  sandbox.setStmtAssign(decisions, idInfo.stmt.ids[2], "P1", "Unlinked installment", "user");
  sandbox.setStmtAssign(decisions, idInfo.stmt.ids[3], "P1", "Fee", "user");
  sandbox.setStmtAssign(decisions, idInfo.stmt.ids[4], "P1", "Refund", "user");
  sandbox.saveDecisions(ss, decisions);

  const recon = sandbox.reconcileActiveStatement(ss, tz);
  const ledger = sandbox.buildDebtLedger(ss, tz, recon);

  const shareP1 = sandbox.shareMemberStatement(ss, tz, "P1", { skipPdf: true });
  const shareP2 = sandbox.shareMemberStatement(ss, tz, "P2", { skipPdf: true });

  const ledgerP1 = ledger.personShares.find(ps => ps.person === "P1");
  const ledgerP2 = ledger.personShares.find(ps => ps.person === "P2");

  assert.strictEqual(shareP1.memberData.totals.purchases, 400);
  assert.strictEqual(shareP1.memberData.totals.installments, 400);
  assert.strictEqual(shareP1.memberData.totals.fees, 40);
  assert.strictEqual(shareP1.memberData.totals.refunds, 50);
  assert.strictEqual(shareP1.memberData.totals.netTotal, Number(ledgerP1.netTotal.toFixed(2)));
  assert.strictEqual(shareP1.memberData.purchases.length, 3, "Grouped debit + split must produce 3 itemized purchases for P1");

  assert.strictEqual(shareP2.memberData.totals.purchases, 100);
  assert.strictEqual(shareP2.memberData.totals.netTotal, Number(ledgerP2.netTotal.toFixed(2)));
  assert.ok(shareP2.whatsappText.includes("(1/2 split)"), "Split item must carry (1/2 split) tag");
  assert.ok(shareP1.whatsappText.includes("⚠️ _Draft: 1 bank line(s) still unresolved"), "Unresolved lines must trigger Draft warning");
  assert.ok(shareP1.whatsappText.includes("-50.00 EGP"), "Refund line must show negative amount");
}

// ==========================================
// Test 3: PDF Export URL, B3 Flush/Restore, Drive Folder Hierarchy & Duplicate Replacement
// ==========================================
{
  const realFixture = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "2026-11-real.json"), "utf8"));
  const { sandbox, CONFIG, setActiveSpreadsheet } = createGasSandbox();
  const ss = new MockSpreadsheet();
  setActiveSpreadsheet(ss);
  populateSpreadsheetFromFixture(ss, realFixture, CONFIG);
  sandbox.updateLiveDashboard({ ss });

  const driveMock = createDriveMock();
  const fetchCalls = [];
  let nextHttpCode = 200;

  sandbox.DriveApp = driveMock.DriveApp;
  sandbox.ScriptApp = { getOAuthToken: () => "mock-oauth-token" };
  sandbox.UrlFetchApp = {
    fetch: (url, params) => {
      const mb = ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
      fetchCalls.push({
        url,
        headers: params && params.headers,
        b3AtFetchTime: mb.getRange("B3").getValue(),
        flushCountAtFetchTime: sandbox.__flushCount || 0
      });
      return {
        getResponseCode: () => nextHttpCode,
        getBlob: () => ({
          name: "",
          setName(n) { this.name = n; return this; }
        })
      };
    }
  };

  const mbSheet = ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
  mbSheet.getRange("B3").setValue("All Family Members");

  const res1 = sandbox.shareMemberStatement(ss, tz, "P5");
  assert.strictEqual(fetchCalls.length, 1);
  const call1 = fetchCalls[0];
  assert.strictEqual(call1.b3AtFetchTime, "P5", "Member Breakdown B3 must be set to target member before fetch");
  assert.ok(call1.flushCountAtFetchTime >= 1, "SpreadsheetApp.flush() must run before fetching PDF");
  assert.strictEqual(mbSheet.getRange("B3").getValue(), "All Family Members", "B3 must be restored after export");
  assert.strictEqual(call1.headers.Authorization, "Bearer mock-oauth-token");

  const expectedGid = mbSheet.getSheetId();
  assert.ok(call1.url.includes("format=pdf"));
  assert.ok(call1.url.includes(`gid=${expectedGid}`));
  assert.ok(call1.url.includes("portrait=false"));
  assert.ok(call1.url.includes("fitw=true"));
  assert.ok(call1.url.includes("size=A4"));
  assert.ok(call1.url.includes("c1=0") && call1.url.includes("c2=16"), "Export range must clip A:P (excluding hidden Col Z)");

  assert.strictEqual(res1.pdf.fileName, "NBE_2026-11_P5_Statement.pdf");
  assert.ok(res1.pdf.fileUrl.includes("drive.google.com/file/d/"));

  const baseFolder = driveMock.root.children.find(f => f.name === "NBE Statements");
  assert.ok(baseFolder, "Statements root folder must be created");
  const msFolder = baseFolder.children.find(f => f.name === "Member Statements");
  assert.ok(msFolder, "Member Statements subfolder must be created");
  const cycFolder = msFolder.children.find(f => f.name === "2026-11");
  assert.ok(cycFolder, "Cycle subfolder 2026-11 must be created");
  assert.strictEqual(cycFolder.files.length, 1);
  assert.strictEqual(cycFolder.files[0].name, "NBE_2026-11_P5_Statement.pdf");
  assert.strictEqual(cycFolder.files[0].trashed, false);

  // Re-exporting same member trashes previous duplicate
  const res2 = sandbox.shareMemberStatement(ss, tz, "P5");
  assert.strictEqual(cycFolder.files.length, 2);
  assert.strictEqual(cycFolder.files[0].trashed, true, "Older duplicate PDF must be trashed");
  assert.strictEqual(cycFolder.files[1].trashed, false, "Newly exported PDF must stay active");
  assert.notStrictEqual(res2.pdf.fileId, res1.pdf.fileId);

  // HTTP 500 error still restores B3 cleanly via finally block
  nextHttpCode = 500;
  assert.throws(
    () => sandbox.shareMemberStatement(ss, tz, "P6"),
    /HTTP 500/i,
    "HTTP 500 must throw descriptive error"
  );
  assert.strictEqual(mbSheet.getRange("B3").getValue(), "All Family Members", "B3 must still restore on fetch failure");
}

// ==========================================
// Test 4: Modal HTML Generation & Escaping (buildShareDialogHtml)
// ==========================================
{
  const { sandbox } = createGasSandbox();
  const maliciousName = 'Ali & <script>alert(1)</script>';
  const html = sandbox.buildShareDialogHtml({
    cycleLabel: "November 2026",
    dueDateStr: "November 27, 2026",
    unresolvedCount: 2,
    defaultPerson: maliciousName,
    members: [
      {
        person: maliciousName,
        netTotal: 1234.5,
        whatsappText: "Line with </script> and & < >"
      }
    ]
  });

  assert.ok(html.includes('<select id="memberSelect">'));
  assert.ok(html.includes('id="copyBtn"') && html.includes('id="pdfBtn"'));
  assert.ok(html.includes("generateMemberStatementPdfForUi"));
  assert.ok(html.includes("Draft: 2 bank line(s)"), "Unresolved draft warning must show in dialog");
  assert.ok(html.includes("Ali &amp; &lt;script&gt;alert(1)&lt;/script&gt;"), "Option label must HTML-escape member name");

  // Only the final closing </body></html> script tag may exist in HTML
  const scriptCloseCount = (html.match(/<\/script>/gi) || []).length;
  assert.strictEqual(scriptCloseCount, 1, "Embedded JSON must escape '</' so payload never closes <script>");
  assert.ok(html.includes("\\u003c/script>"));
}

// ==========================================
// Test 5: Invalid / Empty Member Input Guard
// ==========================================
{
  const { sandbox } = createGasSandbox();
  const ss = new MockSpreadsheet();
  assert.throws(() => sandbox.buildMemberStatementData(ss, tz, ""), /valid family member/i);
  assert.throws(() => sandbox.buildMemberStatementData(ss, tz, "All Family Members"), /valid family member/i);
  assert.throws(() => sandbox.buildMemberStatementData(ss, tz, "🚫 Neglect / Ignore"), /valid family member/i);
}

console.log("share.test.js: all per-member exact math, WhatsApp formatting, PDF export, modal HTML, and guard assertions passed.");
