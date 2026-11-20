/**
 * Configuration & Optional Settings Sheet (No People Profiles)
 *
 * People are auto-discovered from Transactions, Installments, and _Decisions.
 * This module manages an optional, 3-panel `Settings` sheet for only the things
 * that cannot be inferred from ledger rows:
 *   - Default Fee / Refund Payer (optional)
 *   - Credit Limit & Statements Drive Folder overrides
 *   - Optional Name Aliases (e.g. typos / nicknames -> canonical name)
 *   - Editable Merchant Aliases synced bi-directionally with _Decisions
 *
 * NOTE: Apps Script V8 load order across files is arbitrary — never reference CONFIG or Code.js
 * helpers at top-level evaluation time; only call them inside function bodies.
 */

let _settingsCache = { ss: null, settings: null };

function invalidateSettingsCache_() {
  _settingsCache = { ss: null, settings: null };
}

function loadSettings(ss) {
  const fallbackFeePayer = (typeof CONFIG !== "undefined" && CONFIG.DEFAULT_FEE_PAYER)
    ? String(CONFIG.DEFAULT_FEE_PAYER).trim()
    : "";
  const fallbackCreditLimit = (typeof CONFIG !== "undefined" && CONFIG.CREDIT_LIMIT) || 100000.00;
  const fallbackFolderName = (typeof CONFIG !== "undefined" && CONFIG.FOLDER_NAME) || "NBE Statements";

  if (!ss) {
    return {
      hasSettingsSheet: false,
      defaultFeePayer: fallbackFeePayer,
      creditLimit: fallbackCreditLimit,
      folderName: fallbackFolderName,
      nameAliases: {},
      merchantAliases: []
    };
  }

  const sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sheet) {
    const defaults = {
      hasSettingsSheet: false,
      defaultFeePayer: fallbackFeePayer,
      creditLimit: fallbackCreditLimit,
      folderName: fallbackFolderName,
      nameAliases: {},
      merchantAliases: []
    };
    _settingsCache = { ss, settings: defaults };
    return defaults;
  }

  const scanRows = Math.max(3, sheet.getLastRow() >= 6 ? sheet.getLastRow() - 5 : 3);
  const grid = sheet.getRange(6, 1, scanRows, 9).getValues();

  const rawDefaultPayer = grid[0] && grid[0][1] ? String(grid[0][1]).trim() : "";
  const rawLimit = grid[1] && grid[1][1] !== "" && grid[1][1] !== null && grid[1][1] !== undefined
    ? parseFloat(grid[1][1])
    : NaN;
  const rawFolder = grid[2] && grid[2][1] ? String(grid[2][1]).trim() : "";

  const nameAliases = {};
  const merchantAliases = [];

  for (let i = 0; i < grid.length; i++) {
    const row = grid[i];
    const aliasCell = String(row[3] || "").trim();
    const canonCell = String(row[4] || "").replace(/[\u00A0\s]+/g, " ").trim();
    if (aliasCell && canonCell) {
      aliasCell.split(/[,\/;]+/).forEach(tok => {
        const key = String(tok || "").toLowerCase().trim();
        if (key) nameAliases[key] = canonCell;
      });
    }

    const stmtKey = String(row[6] || "").toLowerCase().trim();
    const txKey = String(row[7] || "").toLowerCase().trim();
    const note = String(row[8] || "").trim();
    if (stmtKey && txKey) {
      merchantAliases.push({ stmtKey, txKey, note });
    }
  }

  const settings = {
    hasSettingsSheet: true,
    defaultFeePayer: rawDefaultPayer || fallbackFeePayer,
    creditLimit: (!isNaN(rawLimit) && rawLimit > 0) ? rawLimit : fallbackCreditLimit,
    folderName: rawFolder || fallbackFolderName,
    nameAliases,
    merchantAliases
  };

  _settingsCache = { ss, settings };
  return settings;
}

function getActiveNameAliasMap_(optSs) {
  try {
    const ss = optSs || (typeof SpreadsheetApp !== "undefined" && SpreadsheetApp.getActiveSpreadsheet && SpreadsheetApp.getActiveSpreadsheet());
    if (!ss) return {};
    if (_settingsCache.ss === ss && _settingsCache.settings) {
      return _settingsCache.settings.nameAliases || {};
    }
    const s = loadSettings(ss);
    return (s && s.nameAliases) || {};
  } catch (e) {
    return {};
  }
}

function getDefaultFeePayer(ss) {
  const settings = loadSettings(ss);
  const raw = settings && settings.defaultFeePayer ? String(settings.defaultFeePayer).trim() : "";
  if (!raw) return "";
  return normalizePersonName(raw, settings) || raw;
}

function getEffectiveCreditLimit(ss) {
  const settings = loadSettings(ss);
  return (settings && settings.creditLimit) || CONFIG.CREDIT_LIMIT;
}

function getEffectiveFolderName(ss) {
  const settings = loadSettings(ss);
  return (settings && settings.folderName) || CONFIG.FOLDER_NAME;
}

function syncSettingsToDecisions(ss, optDecisions, optOptions) {
  if (!ss) return false;
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!settingsSheet) return false;

  const settings = loadSettings(ss);
  const decisions = optDecisions || loadDecisions(ss);
  if (!decisions.merchantAliases) decisions.merchantAliases = {};

  const replaceAll = Boolean(optOptions && optOptions.replaceAll);
  if (replaceAll) {
    const nextAliases = {};
    (settings.merchantAliases || []).forEach(a => {
      if (!a.stmtKey || !a.txKey || a.stmtKey === a.txKey) return;
      const mapKey = `${a.stmtKey}::${a.txKey}`;
      const existing = decisions.merchantAliases[mapKey];
      nextAliases[mapKey] = {
        stmtKey: a.stmtKey,
        txKey: a.txKey,
        note: a.note || (existing && existing.note) || "From Settings",
        source: (existing && existing.source) || "settings",
        locked: true,
        updatedAt: (existing && existing.updatedAt) || nowCairoTimestamp_()
      };
    });

    const prevSig = Object.keys(decisions.merchantAliases).sort().map(k => {
      const x = decisions.merchantAliases[k];
      return `${k}|${x && x.note ? x.note : ""}`;
    }).join(";;");
    const nextSig = Object.keys(nextAliases).sort().map(k => {
      const x = nextAliases[k];
      return `${k}|${x && x.note ? x.note : ""}`;
    }).join(";;");

    if (prevSig !== nextSig) {
      decisions.merchantAliases = nextAliases;
      decisions._dirty = true;
    }
  } else {
    (settings.merchantAliases || []).forEach(a => {
      if (!a.stmtKey || !a.txKey || a.stmtKey === a.txKey) return;
      const mapKey = `${a.stmtKey}::${a.txKey}`;
      if (!decisions.merchantAliases[mapKey]) {
        setMerchantAlias(decisions, a.stmtKey, a.txKey, a.note || "From Settings", "settings");
      }
    });
  }

  if (!optDecisions && decisions._dirty) {
    saveDecisions(ss, decisions);
  }
  return Boolean(decisions._dirty);
}

function syncDecisionsMerchantAliasesToSettings(ss, decisions) {
  if (!ss || !decisions) return;
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sheet) return;

  const aliases = decisions.merchantAliases || {};
  const newRows = Object.keys(aliases).sort().map(k => {
    const a = aliases[k];
    return [a.stmtKey, a.txKey, a.note || (a.source === "user" ? "Auto-learned" : "Manual")];
  }).filter(r => r[0] && r[1]);

  const lastRow = Math.max(6, sheet.getLastRow());
  const oldGrid = sheet.getRange(6, 7, lastRow - 5, 3).getValues();
  const oldRows = oldGrid
    .map(r => [String(r[0] || "").toLowerCase().trim(), String(r[1] || "").toLowerCase().trim(), String(r[2] || "").trim()])
    .filter(r => r[0] && r[1]);

  if (JSON.stringify(oldRows) === JSON.stringify(newRows)) return;

  if (lastRow >= 6) {
    sheet.getRange(6, 7, lastRow - 5, 3).clearContent();
  }
  if (newRows.length > 0) {
    sheet.getRange(6, 7, newRows.length, 3).setValues(newRows).setFontSize(9);
  }
  invalidateSettingsCache_();
}

function getOrCreateSettingsSheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.SETTINGS);
  if (sheet.getLastRow() < 5) {
    sheet.clear();
    sheet.setHiddenGridlines(false);

    sheet.getRange("A1:I1").merge()
      .setValue("⚙️ NBE Credit Card — Tracker Settings (Optional Overrides Only)")
      .setFontWeight("bold")
      .setFontSize(13)
      .setBackground(CONFIG.COLORS.PRIMARY)
      .setFontColor("#ffffff");

    sheet.getRange("A2:I2").merge()
      .setValue("Family members are auto-discovered from Transactions & Installments (no People sheet needed). Use this tab only for defaults and overrides.")
      .setFontStyle("italic")
      .setFontSize(9)
      .setFontColor(CONFIG.COLORS.TEXT_MUTED);

    sheet.getRange("A4:B4").merge()
      .setValue("🔧 General Settings")
      .setFontWeight("bold")
      .setFontSize(10)
      .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

    sheet.getRange("A5:B5").setValues([["Setting", "Value"]])
      .setFontWeight("bold")
      .setFontSize(9)
      .setBackground(CONFIG.COLORS.HEADER);

    const existingFeePayer = (typeof CONFIG !== "undefined" && CONFIG.DEFAULT_FEE_PAYER) ? CONFIG.DEFAULT_FEE_PAYER : "";
    const genRows = [
      ["Default Fee / Refund Payer", existingFeePayer],
      ["Credit Limit (EGP)", CONFIG.CREDIT_LIMIT],
      ["Statements Folder Name", CONFIG.FOLDER_NAME]
    ];
    sheet.getRange(6, 1, genRows.length, 2).setValues(genRows).setFontSize(9);
    sheet.getRange("A6:A8").setFontWeight("bold");
    sheet.getRange("B6:B8").setBackground("#fffde7");
    sheet.getRange("B7").setNumberFormat("#,##0.00");

    sheet.getRange("D4:E4").merge()
      .setValue("🏷️ Name Aliases (Optional Overrides)")
      .setFontWeight("bold")
      .setFontSize(10)
      .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

    sheet.getRange("D5:E5").setValues([["Alias / Typo (comma-separated ok)", "Canonical Name"]])
      .setFontWeight("bold")
      .setFontSize(9)
      .setBackground(CONFIG.COLORS.HEADER);

    sheet.getRange("G4:I4").merge()
      .setValue("🏪 Merchant Aliases (Auto-Learned + Editable)")
      .setFontWeight("bold")
      .setFontSize(10)
      .setBackground(CONFIG.COLORS.PRIMARY_LIGHT);

    sheet.getRange("G5:I5").setValues([["Statement Merchant", "Transaction Log Name", "Source / Note"]])
      .setFontWeight("bold")
      .setFontSize(9)
      .setBackground(CONFIG.COLORS.HEADER);

    sheet.setFrozenRows(5);

    applySheetLayoutFitting(sheet, {
      bannerRow: 1,
      bannerHeight: 30,
      customRowHeights: { 2: 20, 4: 24, 5: 22 },
      explicitColWidths: {
        1: 185,
        2: 145,
        3: 20,
        4: 205,
        5: 135,
        6: 20,
        7: 165,
        8: 165,
        9: 185
      }
    });
  }

  const decisions = loadDecisions(ss);
  syncDecisionsMerchantAliasesToSettings(ss, decisions);

  try {
    const payers = getAllUniquePayers(ss);
    if (payers.length > 0 && typeof SpreadsheetApp !== "undefined" && SpreadsheetApp.newDataValidation) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(payers, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange("B6").setDataValidation(rule);
    }
  } catch (e) { }

  invalidateSettingsCache_();
  return sheet;
}

function menuOpenSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSettingsSheet(ss);
  ss.setActiveSheet(sheet);
}
