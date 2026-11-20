/**
 * Per-Person Shareable Statement (WhatsApp Summary + Landscape A4 PDF Export)
 *
 * Generates itemized per-member statements matching buildDebtLedger exact figures,
 * formats WhatsApp-ready summaries, exports Member Breakdown PDFs to Google Drive,
 * and powers the interactive Share Member Statement modal dialog.
 *
 * NOTE: Apps Script V8 load order across files is arbitrary — never reference CONFIG or
 * Code.js helpers at top-level evaluation time; only call them inside function bodies.
 */

function roundShare2_(n) {
  const val = Number((Number(n) || 0).toFixed(2));
  return (Object.is(val, -0) || val === 0) ? 0 : val;
}

function formatShareMoney_(n) {
  const rounded = roundShare2_(n);
  if (rounded === 0) return "0.00";
  const sign = rounded < 0 ? "-" : "";
  const fixed = Math.abs(rounded).toFixed(2);
  const parts = fixed.split(".");
  const intPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return sign + intPart + "." + parts[1];
}

function formatShortItemDate_(dateVal, tz) {
  if (!dateVal) return "";
  const norm = makeNormDateStr(dateVal, tz);
  const m = String(norm || "").match(/^\d{4}-(\d{2})-(\d{2})$/);
  return m ? `${m[1]}/${m[2]}` : "";
}

function sanitizePdfFileToken_(str) {
  const cleaned = String(str || "")
    .replace(/[^A-Za-z0-9_-]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
  return cleaned || "Member";
}

function escapeShareHtml_(str) {
  return String(str === null || str === undefined ? "" : str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function pushMemberSplitItem_(list, canonicalPerson, payerStr, totalAmt, meta, tz) {
  if (!payerStr || isNeglectedPayer(payerStr)) return 0;
  const people = splitPayerNames(payerStr);
  if (people.length === 0) return 0;
  const splitAmt = (parseFloat(totalAmt) || 0) / people.length;
  const isSplit = people.length > 1;
  const dateObj = (meta && meta.date) || "";
  const dateStr = makeNormDateStr(dateObj, tz) || String((meta && meta.dateStr) || "").trim();
  let addedRaw = 0;

  people.forEach(p => {
    if (p === canonicalPerson) {
      addedRaw += splitAmt;
      list.push({
        date: dateObj,
        dateStr: dateStr,
        desc: String((meta && meta.desc) || "").trim(),
        rawAmount: splitAmt,
        amount: roundShare2_(splitAmt),
        isSplit: isSplit,
        splitCount: people.length
      });
    }
  });
  return addedRaw;
}

function buildMemberStatementData(ss, tz, personName, optContext) {
  if (!ss) throw new Error("Spreadsheet reference is required.");
  const targetTz = tz || ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const settings = loadSettings(ss);
  const canonicalPerson = normalizePersonName(personName, settings);
  if (!canonicalPerson || canonicalPerson === "All Family Members") {
    throw new Error("Please select a valid family member.");
  }

  let recon = optContext && optContext.recon;
  let ledger = optContext && optContext.ledger;
  if (!recon) {
    syncActiveReconSheetToDecisions(ss, targetTz);
    recon = reconcileActiveStatement(ss, targetTz);
    if (recon.error) {
      throw new Error(recon.error);
    }
  }
  if (!ledger) {
    ledger = buildDebtLedger(ss, targetTz, recon);
  }

  const purchases = [];
  const installments = [];
  const unrecorded = [];
  const fees = [];
  const refunds = [];
  const carryover = [];

  let rawPurchases = 0;
  let rawInstallments = 0;
  let rawUnrecorded = 0;
  let rawUnrecordedItems = 0;
  let rawFees = 0;
  let rawRefunds = 0;
  let rawRefundItems = 0;
  let rawCarryoverSigned = 0;

  // 1. Matched Purchases
  (recon.stmtDebits || []).forEach(st => {
    if (!st.matched) return;
    if (st.matchedItems && st.matchedItems.length > 1) {
      st.matchedItems.forEach(tx => {
        if (tx.person && !isNeglectedPayer(tx.person)) {
          rawPurchases += pushMemberSplitItem_(purchases, canonicalPerson, tx.person, tx.amount, {
            date: tx.date || st.date,
            dateStr: st.dateStr,
            desc: tx.desc || st.desc
          }, targetTz);
        }
      });
    } else if (st.isPartialMatch && st.matchedItems && st.matchedItems.length === 1) {
      const tx = st.matchedItems[0];
      if (tx.person && !isNeglectedPayer(tx.person)) {
        rawPurchases += pushMemberSplitItem_(purchases, canonicalPerson, tx.person, tx.amount, {
          date: tx.date || st.date,
          dateStr: st.dateStr,
          desc: tx.desc || st.desc
        }, targetTz);
      }
    } else if (st.assignedPayer && !isNeglectedPayer(st.assignedPayer)) {
      const firstTx = st.matchedItems && st.matchedItems[0];
      rawPurchases += pushMemberSplitItem_(purchases, canonicalPerson, st.assignedPayer, st.amount, {
        date: (firstTx && firstTx.date) || st.date,
        dateStr: st.dateStr,
        desc: (firstTx && firstTx.desc) || st.desc
      }, targetTz);
    }
  });
  purchases.sort((a, b) => (a.dateStr || "").localeCompare(b.dateStr || ""));

  // 2. Installments (Table 4 reconciled items + assigned unlinked statement installments)
  (recon.installmentReconItems || []).forEach(inst => {
    const act = String((inst && inst.action) || "").trim();
    if (!act || isNeglectedPayer(act)) return;
    const cleanPayer = act.replace(/^✅\s*Billed\s+to\s+/i, "").trim();
    if (cleanPayer && !isNeglectedPayer(cleanPayer)) {
      rawInstallments += pushMemberSplitItem_(installments, canonicalPerson, cleanPayer, inst.amount, {
        date: inst.date,
        desc: inst.desc
      }, targetTz);
    }
  });

  (recon.unlinkedStmtInstallments || []).forEach(u => {
    if (!u || !u.assignedPayer || isNeglectedPayer(u.assignedPayer)) return;
    rawInstallments += pushMemberSplitItem_(installments, canonicalPerson, u.assignedPayer, u.amount, {
      date: u.date || u.dateStr,
      dateStr: u.dateStr,
      desc: u.desc
    }, targetTz);
  });

  // 3. Unrecorded Missing Statement Purchases
  (recon.missingInSheet || []).forEach(m => {
    if (!m || !m.assignedPayer || isNeglectedPayer(m.assignedPayer)) return;
    const added = pushMemberSplitItem_(unrecorded, canonicalPerson, m.assignedPayer, m.amount, {
      date: m.date || m.dateStr,
      dateStr: m.dateStr,
      desc: m.desc
    }, targetTz);
    rawUnrecorded += added;
    rawUnrecordedItems += added;
  });

  // 4. Bank Fees & Taxes
  (recon.stmtFees || []).forEach(f => {
    if (!f || !f.assignedPayer || isNeglectedPayer(f.assignedPayer)) return;
    rawFees += pushMemberSplitItem_(fees, canonicalPerson, f.assignedPayer, f.amount, {
      date: f.date || f.dateStr,
      dateStr: f.dateStr,
      desc: f.desc
    }, targetTz);
  });

  // 5. Refunds / Credits
  (recon.stmtCredits || []).forEach(cr => {
    if (!cr || !cr.assignedPayer || isNeglectedPayer(cr.assignedPayer)) return;
    const added = pushMemberSplitItem_(refunds, canonicalPerson, cr.assignedPayer, cr.amount, {
      date: cr.date || cr.dateStr,
      dateStr: cr.dateStr,
      desc: cr.desc
    }, targetTz);
    rawRefunds += added;
    rawRefundItems += added;
  });

  // 6. Prior Cycle Carryover (when explicitly assigned to member)
  const carryItem = recon.carryoverItem;
  if (carryItem && carryItem.assignedPayer && !isNeglectedPayer(carryItem.assignedPayer)) {
    const carryAmt = Number((ledger && ledger.carriedOver !== undefined ? ledger.carriedOver : carryItem.amount) || 0);
    if (Math.abs(carryAmt) >= 0.005) {
      const carryMeta = {
        date: carryItem.date || carryItem.dateStr,
        dateStr: carryItem.dateStr,
        desc: carryItem.desc
      };
      const prevLen = carryover.length;
      const absAdded = pushMemberSplitItem_(carryover, canonicalPerson, carryItem.assignedPayer, Math.abs(carryAmt), carryMeta, targetTz);
      if (absAdded > 0) {
        if (carryAmt > 0) {
          rawUnrecorded += absAdded;
          rawCarryoverSigned += absAdded;
        } else {
          rawRefunds += absAdded;
          rawCarryoverSigned -= absAdded;
          for (let ci = prevLen; ci < carryover.length; ci++) {
            carryover[ci].isCredit = true;
          }
        }
      }
    }
  }

  const netTotal = roundShare2_(rawPurchases + rawInstallments + rawFees + rawUnrecorded - rawRefunds);
  const cardBal = recon.cardBal || {};
  const cycleKey = ledger.activeCycleKey || cardBal.activeCycleKey || "";
  const cycleLabel = ledger.activeCycleLabel || cardBal.activeCycleLabel || cycleKey;
  const statementDateStr = cardBal.statementDate
    ? Utilities.formatDate(cardBal.statementDate, targetTz, "yyyy-MM-dd")
    : "N/A";
  const dueDateObj = ledger.activeDueDate || cardBal.statementDueDate;
  const dueDateStr = dueDateObj
    ? Utilities.formatDate(dueDateObj, targetTz, "MMMM d, yyyy")
    : "N/A";

  let paymentDateStr = "";
  const arSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_RECON);
  if (arSheet) {
    const h6Date = parseDateValue(arSheet.getRange(6, 8).getValue(), targetTz);
    if (h6Date && !isNaN(h6Date.getTime())) {
      paymentDateStr = Utilities.formatDate(h6Date, targetTz, "yyyy-MM-dd");
    }
  }
  const decisions = recon.decisions || loadDecisions(ss);
  if (!paymentDateStr && cycleKey && decisions && decisions.paymentRecords && decisions.paymentRecords[cycleKey] && decisions.paymentRecords[cycleKey].dateStr) {
    paymentDateStr = decisions.paymentRecords[cycleKey].dateStr;
  }

  const isPaid = Boolean(paymentDateStr);
  const isClosed = typeof isCycleClosed === "function" ? isCycleClosed(ss, cycleKey) : false;
  const unresolvedCount = Number(ledger.unresolvedCount || 0);

  return {
    person: canonicalPerson,
    cycleKey,
    cycleLabel,
    statementDateStr,
    dueDateStr,
    paymentDateStr,
    isPaid,
    isClosed,
    unresolvedCount,
    purchases,
    installments,
    unrecorded,
    fees,
    refunds,
    carryover,
    totals: {
      purchases: roundShare2_(rawPurchases),
      installments: roundShare2_(rawInstallments),
      unrecorded: roundShare2_(rawUnrecorded),
      unrecordedItems: roundShare2_(rawUnrecordedItems),
      fees: roundShare2_(rawFees),
      refunds: roundShare2_(rawRefunds),
      refundItems: roundShare2_(rawRefundItems),
      carryover: roundShare2_(rawCarryoverSigned),
      netTotal: netTotal
    }
  };
}

function formatWhatsAppLine_(item, isInstallment, isRefund, tz) {
  const shortDate = isInstallment ? "" : formatShortItemDate_(item.dateStr || item.date, tz);
  const datePrefix = shortDate ? `${shortDate} ` : "";
  const splitTag = item.isSplit ? ` (1/${item.splitCount} split)` : "";
  const signedAmt = isRefund
    ? `-${formatShareMoney_(item.amount)}`
    : formatShareMoney_(item.amount);
  return `• ${datePrefix}${item.desc}${splitTag} — ${signedAmt} EGP`;
}

function formatWhatsAppStatement(memberData, optTz) {
  if (!memberData) throw new Error("Member statement data is required.");
  const targetTz = optTz || "Africa/Cairo";
  const t = memberData.totals || {};
  const lines = [
    `💳 *NBE Credit Card — ${memberData.cycleLabel}*`,
    `👤 *Member:* ${memberData.person}`,
    `📅 *Due Date:* ${memberData.dueDateStr}`,
    `💰 *Total Due: ${formatShareMoney_(t.netTotal)} EGP*`
  ];

  const purchases = memberData.purchases || [];
  const installments = memberData.installments || [];
  const unrecorded = memberData.unrecorded || [];
  const fees = memberData.fees || [];
  const refunds = memberData.refunds || [];
  const carryover = memberData.carryover || [];

  const hasAnyItems = purchases.length > 0 || installments.length > 0 ||
    unrecorded.length > 0 || fees.length > 0 || refunds.length > 0 || carryover.length > 0;

  lines.push("");

  if (!hasAnyItems) {
    lines.push("✨ _No charges billed in this cycle._");
  } else {
    if (purchases.length > 0) {
      lines.push(`🛍️ *Purchases (${formatShareMoney_(t.purchases)} EGP):*`);
      purchases.forEach(it => lines.push(formatWhatsAppLine_(it, false, false, targetTz)));
      lines.push("");
    }
    if (installments.length > 0) {
      lines.push(`📅 *Installments (${formatShareMoney_(t.installments)} EGP):*`);
      installments.forEach(it => lines.push(formatWhatsAppLine_(it, true, false, targetTz)));
      lines.push("");
    }
    if (unrecorded.length > 0) {
      const unrecSum = t.unrecordedItems !== undefined ? t.unrecordedItems : t.unrecorded;
      lines.push(`🧾 *Unrecorded / Statement Items (${formatShareMoney_(unrecSum)} EGP):*`);
      unrecorded.forEach(it => lines.push(formatWhatsAppLine_(it, false, false, targetTz)));
      lines.push("");
    }
    if (fees.length > 0) {
      lines.push(`🏛️ *Bank Fees & Taxes (${formatShareMoney_(t.fees)} EGP):*`);
      fees.forEach(it => lines.push(formatWhatsAppLine_(it, false, false, targetTz)));
      lines.push("");
    }
    if (refunds.length > 0) {
      const refSum = t.refundItems !== undefined ? t.refundItems : t.refunds;
      lines.push(`↩️ *Refunds / Credits (-${formatShareMoney_(refSum)} EGP):*`);
      refunds.forEach(it => lines.push(formatWhatsAppLine_(it, false, true, targetTz)));
      lines.push("");
    }
    if (carryover.length > 0) {
      const carryHeaderAmt = t.carryover < 0
        ? `-${formatShareMoney_(Math.abs(t.carryover))}`
        : formatShareMoney_(t.carryover);
      lines.push(`🔄 *Prior Cycle Carryover (${carryHeaderAmt} EGP):*`);
      carryover.forEach(it => lines.push(formatWhatsAppLine_(it, false, Boolean(it.isCredit), targetTz)));
      lines.push("");
    }
  }

  lines.push("━━━━━━━━━━━━━━━━━━");
  lines.push(`✅ *Net Total Due: ${formatShareMoney_(t.netTotal)} EGP*`);

  const statusParts = [
    `Statement Date: ${memberData.statementDateStr}`,
    `Due: ${memberData.dueDateStr}`
  ];
  if (memberData.isPaid && memberData.paymentDateStr) {
    statusParts.push(`Paid: ${memberData.paymentDateStr}`);
  }
  if (memberData.isClosed) {
    statusParts.push("Closed");
  }
  lines.push(`_${statusParts.join(" · ")}_`);

  if (memberData.unresolvedCount > 0) {
    lines.push(`⚠️ _Draft: ${memberData.unresolvedCount} bank line(s) still unresolved on statement._`);
  }

  return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

function buildMemberBreakdownPdfUrl(spreadsheetId, sheetId) {
  return (
    `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export` +
    `?format=pdf&gid=${sheetId}&size=A4&portrait=false&fitw=true` +
    `&sheetnames=false&printtitle=false&pagenum=UNDEFINED&gridlines=true&fzr=false` +
    `&top_margin=0.40&bottom_margin=0.40&left_margin=0.35&right_margin=0.35` +
    `&r1=0&c1=0&c2=16`
  );
}

function exportMemberBreakdownPdf(ss, tz, personName, optContext) {
  if (!ss) throw new Error("Spreadsheet reference is required.");
  const targetTz = tz || ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const settings = loadSettings(ss);
  const canonicalPerson = normalizePersonName(personName, settings);
  if (!canonicalPerson || canonicalPerson === "All Family Members") {
    throw new Error("Please select a valid family member.");
  }

  let recon = optContext && optContext.recon;
  let ledger = optContext && optContext.ledger;
  if (!recon) {
    syncActiveReconSheetToDecisions(ss, targetTz);
    recon = reconcileActiveStatement(ss, targetTz);
    if (recon.error) throw new Error(recon.error);
  }
  if (!ledger) {
    ledger = buildDebtLedger(ss, targetTz, recon);
  }

  let mbSheet = ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
  if (!mbSheet) {
    renderMemberBreakdownSheet(ss, targetTz, ledger, recon);
    mbSheet = ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
  }
  if (!mbSheet) {
    throw new Error("Member Breakdown sheet could not be initialized.");
  }

  const prevSelection = mbSheet.getRange("B3").getValue();
  const cycleKey = ledger.activeCycleKey || makeNormDateStr(new Date(), targetTz).substring(0, 7) || "Active";
  const cycleLabel = ledger.activeCycleLabel || cycleKey;

  try {
    mbSheet.getRange("B3").setValue(canonicalPerson);
    SpreadsheetApp.flush();

    const exportUrl = buildMemberBreakdownPdfUrl(ss.getId(), mbSheet.getSheetId());
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code !== 200) {
      throw new Error(`Failed to export Member Breakdown PDF (HTTP ${code}).`);
    }

    const baseFolderName = getEffectiveFolderName(ss);
    const rootFolder = getOrCreateFolder(null, baseFolderName);
    const memberStatementsFolder = getOrCreateFolder(rootFolder, "Member Statements");
    const cycleFolder = getOrCreateFolder(memberStatementsFolder, cycleKey);
    const fileName = `NBE_${cycleKey}_${sanitizePdfFileToken_(canonicalPerson)}_Statement.pdf`;

    const existing = cycleFolder.getFilesByName(fileName);
    while (existing && existing.hasNext()) {
      existing.next().setTrashed(true);
    }

    const blob = response.getBlob().setName(fileName);
    const file = cycleFolder.createFile(blob);

    return {
      person: canonicalPerson,
      cycleKey,
      cycleLabel,
      fileId: file.getId(),
      fileName,
      fileUrl: file.getUrl(),
      folderName: `${baseFolderName} / Member Statements / ${cycleKey}`,
      folderUrl: typeof cycleFolder.getUrl === "function" ? cycleFolder.getUrl() : "",
      exportUrl
    };
  } finally {
    if (
      prevSelection !== "" &&
      prevSelection !== null &&
      prevSelection !== undefined &&
      String(prevSelection) !== canonicalPerson
    ) {
      mbSheet.getRange("B3").setValue(prevSelection);
      SpreadsheetApp.flush();
    }
  }
}

function shareMemberStatement(ss, tz, personName, optOptions) {
  const opts = optOptions || {};
  const targetTz = tz || (ss && ss.getSpreadsheetTimeZone()) || "Africa/Cairo";
  let ctx = opts.context || null;

  if (!ctx || !ctx.recon || !ctx.ledger) {
    syncActiveReconSheetToDecisions(ss, targetTz);
    const recon = reconcileActiveStatement(ss, targetTz);
    if (recon.error) throw new Error(recon.error);
    const ledger = buildDebtLedger(ss, targetTz, recon);
    ctx = { recon, ledger };
  }

  const memberData = buildMemberStatementData(ss, targetTz, personName, ctx);
  const whatsappText = formatWhatsAppStatement(memberData, targetTz);
  const pdf = opts.skipPdf
    ? null
    : exportMemberBreakdownPdf(ss, targetTz, memberData.person, ctx);

  return {
    memberData,
    whatsappText,
    pdf
  };
}

function buildShareDialogHtml(payload) {
  const data = payload || {};
  const members = data.members || [];
  const cycleLabel = escapeShareHtml_(data.cycleLabel || "");
  const dueDateStr = escapeShareHtml_(data.dueDateStr || "");
  const unresolvedCount = Number(data.unresolvedCount || 0);
  const defaultPerson = String(data.defaultPerson || (members[0] && members[0].person) || "");

  const optionTags = members.map(m => {
    const selectedAttr = m.person === defaultPerson ? " selected" : "";
    const label = `${m.person} — ${formatShareMoney_(m.netTotal)} EGP`;
    return `<option value="${escapeShareHtml_(m.person)}"${selectedAttr}>${escapeShareHtml_(label)}</option>`;
  }).join("");

  const safeJson = JSON.stringify({
    defaultPerson,
    members
  }).replace(/</g, "\\u003c");

  const draftNotice = unresolvedCount > 0
    ? `<div class="warn">⚠️ Draft: ${unresolvedCount} bank line(s) are still unresolved on this statement.</div>`
    : "";

  return [
    "<!DOCTYPE html>",
    "<html>",
    "<head>",
    '<meta charset="utf-8">',
    "<style>",
    "  * { box-sizing: border-box; }",
    "  body { font-family: Arial, sans-serif; margin: 14px; color: #222; background: #fafafa; }",
    "  .header { background: #1b5e20; color: #fff; padding: 10px 12px; border-radius: 6px; margin-bottom: 10px; }",
    "  .header h2 { margin: 0 0 2px 0; font-size: 15px; }",
    "  .header .sub { font-size: 11px; opacity: 0.9; }",
    "  .warn { background: #fff2cc; color: #7f6000; border: 1px solid #f0c36d; border-radius: 4px; padding: 6px 8px; font-size: 11px; margin-bottom: 10px; }",
    "  .row { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }",
    "  label { font-weight: bold; font-size: 12px; min-width: 105px; }",
    "  select { flex: 1; padding: 6px 8px; font-size: 12px; border: 1px solid #ccc; border-radius: 4px; background: #fff; }",
    "  textarea { width: 100%; height: 340px; padding: 10px; font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 12px; line-height: 1.45; border: 1px solid #ccc; border-radius: 6px; background: #fff; resize: vertical; }",
    "  .actions { display: flex; gap: 8px; margin-top: 10px; }",
    "  button { flex: 1; padding: 9px 12px; font-size: 12px; font-weight: bold; border: none; border-radius: 5px; cursor: pointer; }",
    "  .btn-copy { background: #1b5e20; color: #fff; }",
    "  .btn-pdf { background: #1565c0; color: #fff; }",
    "  button:disabled { opacity: 0.6; cursor: wait; }",
    "  .status { margin-top: 10px; font-size: 12px; min-height: 36px; padding: 8px 10px; border-radius: 5px; display: none; }",
    "  .status.ok { display: block; background: #e8f5e9; color: #1b5e20; border: 1px solid #c8e6c9; }",
    "  .status.err { display: block; background: #fce8e6; color: #b71c1c; border: 1px solid #f5c6cb; }",
    "  .status a { color: #0d47a1; font-weight: bold; word-break: break-all; }",
    "</style>",
    "</head>",
    "<body>",
    '  <div class="header">',
    `    <h2>📤 Share Member Statement — ${cycleLabel}</h2>`,
    `    <div class="sub">Due Date: ${dueDateStr} · WhatsApp summary + landscape Member Breakdown PDF</div>`,
    "  </div>",
    draftNotice,
    '  <div class="row">',
    '    <label for="memberSelect">Family Member:</label>',
    `    <select id="memberSelect">${optionTags}</select>`,
    "  </div>",
    '  <textarea id="waText" readonly></textarea>',
    '  <div class="actions">',
    '    <button type="button" id="copyBtn" class="btn-copy">📋 Copy for WhatsApp</button>',
    '    <button type="button" id="pdfBtn" class="btn-pdf">📄 Generate PDF in Drive</button>',
    "  </div>",
    '  <div id="statusBox" class="status"></div>',
    "  <script>",
    `    var SHARE_PAYLOAD = ${safeJson};`,
    '    var memberMap = {};',
    '    (SHARE_PAYLOAD.members || []).forEach(function(m) { memberMap[m.person] = m; });',
    '    var selectEl = document.getElementById("memberSelect");',
    '    var textEl = document.getElementById("waText");',
    '    var copyBtn = document.getElementById("copyBtn");',
    '    var pdfBtn = document.getElementById("pdfBtn");',
    '    var statusBox = document.getElementById("statusBox");',
    '    function clearStatus() { statusBox.className = "status"; statusBox.textContent = ""; }',
    '    function showOkText(msg) { statusBox.className = "status ok"; statusBox.textContent = msg; }',
    '    function showErrText(msg) { statusBox.className = "status err"; statusBox.textContent = msg; }',
    '    function updatePreview() {',
    '      var person = selectEl.value;',
    '      var entry = memberMap[person];',
    '      textEl.value = entry ? entry.whatsappText : "";',
    '      copyBtn.textContent = "📋 Copy for WhatsApp";',
    '      clearStatus();',
    '    }',
    '    selectEl.addEventListener("change", updatePreview);',
    '    copyBtn.addEventListener("click", function() {',
    '      var content = textEl.value || "";',
    '      function onCopied() { copyBtn.textContent = "✅ Copied!"; showOkText("✅ WhatsApp summary copied to clipboard."); }',
    '      if (navigator.clipboard && navigator.clipboard.writeText) {',
    '        navigator.clipboard.writeText(content).then(onCopied).catch(function() {',
    '          textEl.select(); document.execCommand("copy"); onCopied();',
    '        });',
    '      } else {',
    '        textEl.select(); document.execCommand("copy"); onCopied();',
    '      }',
    '    });',
    '    pdfBtn.addEventListener("click", function() {',
    '      var person = selectEl.value;',
    '      if (!person) return;',
    '      pdfBtn.disabled = true;',
    '      pdfBtn.textContent = "⏳ Exporting PDF...";',
    '      clearStatus();',
    '      google.script.run',
    '        .withSuccessHandler(function(res) {',
    '          pdfBtn.disabled = false;',
    '          pdfBtn.textContent = "📄 Generate PDF in Drive";',
    '          if (!res || !res.ok) { showErrText("❌ Failed to generate PDF."); return; }',
    '          statusBox.className = "status ok";',
    '          statusBox.textContent = "";',
    '          var prefix = document.createTextNode("✅ Saved PDF to " + (res.folderName || "Drive") + ": ");',
    '          var link = document.createElement("a");',
    '          link.href = res.fileUrl;',
    '          link.target = "_blank";',
    '          link.rel = "noopener noreferrer";',
    '          link.textContent = res.fileName;',
    '          statusBox.appendChild(prefix);',
    '          statusBox.appendChild(link);',
    '          if (res.folderUrl) {',
    '            var sep = document.createTextNode(" · ");',
    '            var fLink = document.createElement("a");',
    '            fLink.href = res.folderUrl;',
    '            fLink.target = "_blank";',
    '            fLink.rel = "noopener noreferrer";',
    '            fLink.textContent = "Open Folder";',
    '            statusBox.appendChild(sep);',
    '            statusBox.appendChild(fLink);',
    '          }',
    '        })',
    '        .withFailureHandler(function(err) {',
    '          pdfBtn.disabled = false;',
    '          pdfBtn.textContent = "📄 Generate PDF in Drive";',
    '          showErrText("❌ PDF export error: " + (err && err.message ? err.message : String(err)));',
    '        })',
    '        .generateMemberStatementPdfForUi(person);',
    '    });',
    '    updatePreview();',
    "  </script>",
    "</body>",
    "</html>"
  ].join("\n");
}

function generateMemberStatementPdfForUi(personName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const res = exportMemberBreakdownPdf(ss, tz, personName);
  return Object.assign({ ok: true }, res);
}

function menuShareMemberStatement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || "Africa/Cairo";
  const ui = SpreadsheetApp.getUi();

  syncActiveReconSheetToDecisions(ss, tz);
  const recon = reconcileActiveStatement(ss, tz);
  if (recon.error) {
    ui.alert("Cannot Share Statement", recon.error, ui.ButtonSet.OK);
    return;
  }
  const ledger = buildDebtLedger(ss, tz, recon);

  if (!ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN)) {
    renderMemberBreakdownSheet(ss, tz, ledger, recon);
  }

  const payers = getAllUniquePayers(ss);
  if (!payers || payers.length === 0) {
    ui.alert("No Family Members Found", "No members were discovered in Transactions, Installments, or Settings.", ui.ButtonSet.OK);
    return;
  }

  const ctx = { recon, ledger };
  const members = payers.map(p => {
    const memberData = buildMemberStatementData(ss, tz, p, ctx);
    const whatsappText = formatWhatsAppStatement(memberData, tz);
    return {
      person: memberData.person,
      netTotal: memberData.totals.netTotal,
      whatsappText
    };
  });

  let defaultPerson = "";
  const mbSheet = ss.getSheetByName(CONFIG.SHEETS.MEMBER_BREAKDOWN);
  if (mbSheet) {
    const selectedRaw = mbSheet.getRange("B3").getValue();
    const normSelected = normalizePersonName(selectedRaw, loadSettings(ss));
    if (normSelected && normSelected !== "All Family Members" && members.some(m => m.person === normSelected)) {
      defaultPerson = normSelected;
    }
  }
  if (!defaultPerson) {
    const nonZero = members.find(m => Math.abs(m.netTotal) > 0.005);
    defaultPerson = nonZero ? nonZero.person : members[0].person;
  }

  const sampleDueObj = ledger.activeDueDate || (recon.cardBal && recon.cardBal.statementDueDate);
  const dueDateStr = sampleDueObj ? Utilities.formatDate(sampleDueObj, tz, "MMMM d, yyyy") : "N/A";
  const html = buildShareDialogHtml({
    cycleLabel: ledger.activeCycleLabel,
    dueDateStr,
    unresolvedCount: ledger.unresolvedCount || 0,
    defaultPerson,
    members
  });

  const output = HtmlService.createHtmlOutput(html).setWidth(560).setHeight(620);
  ui.showModalDialog(output, "Share Member Statement");
}
