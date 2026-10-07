---
name: google-apps-script-patterns
description: >-
  Provides production-grade patterns, performance optimizations, and architectural best practices for Google Apps Script and Google Workspace automation (Sheets, Drive, Gmail). Use when building, debugging, or optimizing Apps Script projects to enforce batch I/O operations (getValues/setValues), avoid quota limitations, manage the 6-minute execution timeout, safely use PropertiesService and CacheService, handle Drive PDF ingestion, and design robust web endpoints.
---

# Google Apps Script & Workspace Automation Patterns

Google Apps Script runs in a hosted cloud container with strict quotas, execution limits, and high network latency for spreadsheet cell operations. This skill provides architectural guidelines and patterns to write fast, resilient, and enterprise-grade Apps Script solutions.

---

## 1. The Cardinal Rule: Batch Sheet I/O

Every call to `SpreadsheetApp` (e.g. `getValue()`, `setValue()`, `setBackground()`, `setNumberFormat()`) incurs an HTTP RPC network trip to Google servers (~50–200ms per call).

### Anti-Pattern: Cell-by-Cell Looping
```javascript
// ❌ SLOW: 100 rows * 3 calls = 300 network round trips (~30-60 seconds)
for (let i = 1; i <= lastRow; i++) {
  const val = sheet.getRange(i, 1).getValue();
  if (val > 100) {
    sheet.getRange(i, 2).setValue("High");
    sheet.getRange(i, 2).setBackground("#ffcccc");
  }
}
```

### Production Pattern: 2D Array In-Memory Batching
Read once, compute in-memory, write once:
```javascript
// ✅ FAST: 2 network calls total (~200ms)
const numRows = sheet.getLastRow();
if (numRows < 2) return;

// 1. Single read
const data = sheet.getRange(2, 1, numRows - 1, sheet.getLastColumn()).getValues();

const outputValues = [];
const outputColors = [];

for (let r = 0; r < data.length; r++) {
  const amount = parseFloat(data[r][0]) || 0;
  if (amount > 100) {
    outputValues.push(["High"]);
    outputColors.push(["#ffcccc"]);
  } else {
    outputValues.push(["Normal"]);
    outputColors.push([null]);
  }
}

// 2. Single write
const targetRange = sheet.getRange(2, 2, outputValues.length, 1);
targetRange.setValues(outputValues);
// Apply formatting only where ranges exist
for (let r = 0; r < outputColors.length; r++) {
  if (outputColors[r][0]) {
    sheet.getRange(2 + r, 2).setBackground(outputColors[r][0]);
  }
}
```

---

## 2. Managing Quotas & PropertiesService Storage

### 9KB Script Property Quota
- `PropertiesService.getScriptProperties().setProperty(key, value)` throws an exception if the serialized value exceeds **9,000 characters** (approx. 9 KB).
- **Safe Properties Wrapper Pattern**: Always implement length validation, auto-compaction, or eviction of stale entries before setting script properties:
  ```javascript
  function safeSetScriptProperty(key, val) {
    try {
      let strVal = typeof val === "string" ? val : JSON.stringify(val);
      if (strVal.length > 8500) {
        // Compact payload: retain only essential fields or the most recent items
        try {
          const parsed = JSON.parse(strVal);
          if (Array.isArray(parsed)) {
            strVal = JSON.stringify(parsed.slice(-50));
          } else if (typeof parsed === "object") {
            const keys = Object.keys(parsed).slice(-40);
            const compacted = {};
            keys.forEach(k => { compacted[k] = parsed[k]; });
            strVal = JSON.stringify(compacted);
          }
        } catch (e) {}
      }
      PropertiesService.getScriptProperties().setProperty(key, strVal);
    } catch (err) {
      Logger.log(`[SafeProperties] Could not write '${key}': ` + (err?.message || err));
    }
  }
  ```

---

## 3. Handling Execution Timeouts (6-Minute Limit)

Google Apps Script enforces a **6-minute execution limit** for standard consumer accounts (30 minutes for Google Workspace Enterprise).

### Checkpointing & Continuation Trigger Pattern
For batch migrations or large file processing that may exceed 5 minutes:
1. Track elapsed execution time with `Date.now()`.
2. Stop processing when elapsed time reaches 4.5 minutes (270,000 ms).
3. Save the current row/index checkpoint to `PropertiesService`.
4. Programmatically schedule a one-time time-driven trigger to resume from the checkpoint:
   ```javascript
   function processLargeDataset() {
     const startTime = Date.now();
     const props = PropertiesService.getScriptProperties();
     let startIndex = parseInt(props.getProperty("LAST_PROCESSED_ROW") || "0", 10);
     
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");
     const data = sheet.getDataRange().getValues();
     
     for (let i = startIndex; i < data.length; i++) {
       // Check time budget: abort before hard timeout at 4.5 minutes
       if (Date.now() - startTime > 270000) {
         props.setProperty("LAST_PROCESSED_ROW", String(i));
         ScriptApp.newTrigger("processLargeDataset")
           .timeBased()
           .after(60 * 1000) // Resume in 1 minute
           .create();
         return;
       }
       // Process row i...
     }
     
     // Clean up checkpoint upon full completion
     props.deleteProperty("LAST_PROCESSED_ROW");
     deleteExistingTriggers("processLargeDataset");
   }
   ```

---

## 4. Drive & PDF Ingestion Patterns

When processing uploaded bank statements or receipts from Google Drive:
- **Search Query**: Use specific Drive queries (e.g. `'<folderId>' in parents and mimeType = 'application/pdf' and trashed = false`).
- **OCR Ingestion**: When extracting text from scanned PDFs, use `Drive.Files.insert` with OCR enabled, or parse text using standard PDF parsers.
- **Archive Processed Files**: Never leave ingested files in the incoming folder. Move processed files to a dedicated `"Processed"` subfolder to prevent reprocessing loops:
  ```javascript
  function moveFileToFolder(file, targetFolder) {
    targetFolder.addFile(file);
    const parents = file.getParents();
    while (parents.hasNext()) {
      parents.next().removeFile(file);
    }
  }
  ```

---

## 5. Webhook & Web App Endpoints (`doGet` / `doPost`)

When receiving form submissions or webhook notifications:
- **LockService for Concurrency**: Use `LockService.getScriptLock()` to prevent race conditions when appending rows simultaneously:
  ```javascript
  function doPost(e) {
    const lock = LockService.getScriptLock();
    // Wait up to 10 seconds for concurrent requests
    if (!lock.tryLock(10000)) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "Server busy, please retry."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    try {
      const payload = JSON.parse(e.postData.contents || "{}");
      // Append row safely...
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        received: payload
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: err.message
      })).setMimeType(ContentService.MimeType.JSON);
    } finally {
      lock.releaseLock();
    }
  }
  ```

---

## 6. Clasp & Version Control Best Practices

- Always maintain a strict `.claspignore` file so only code files (`Code.js`, `.gs`, `appsscript.json`) are pushed to the Apps Script cloud project:
  ```text
  **/**
  !appsscript.json
  !Code.js
  ```
- Push changes with `npx @google/clasp push -f` after syntax-validating with `node -c Code.js`.

---

## 7. Spreadsheet Layout, Column Fitting & The Merged-Banner AutoResize Pitfall

### The Pitfall: Oversized Column 1
In Google Sheets, calling `sheet.autoResizeColumns(1, N)` on any sheet containing a merged header banner in Row 1 (e.g. `sheet.getRange("A1:K1").merge().setValue("...")`) causes Google Sheets to calculate Column 1's width against the **entire merged text string**.
- **Resulting Bug**: Column A (Column 1) blows out to 400–600 pixels wide, pushing all data tables off-screen and leaving massive empty whitespace next to names or `#` numbers.
- **Row 1 Height**: If left unconfigured, high font sizes (14pt+) in merged banners cause Row 1 to expand to an overly tall, awkward height.

### Production Pattern: Controlled Fitting & Explicit Overrides
1. **Never auto-resize Column 1 over merged banners**:
   Always auto-resize from column 2 onwards (`sheet.autoResizeColumns(2, numCols - 1)`), and set Column 1 explicitly with `sheet.setColumnWidth(1, targetWidth)`.
2. **Explicit Row Heights for Banners**:
   Set standard proportional row heights:
   - Row 1 (Title Banner): `sheet.setRowHeight(1, 30)` (or 32px max).
   - Row 2 (Subtitle): `sheet.setRowHeight(2, 20)`.
   - Section header rows: `sheet.setRowHeight(r, 26)`.
3. **Use a Universal Fitting Helper**:
   ```javascript
   function applySheetLayoutFitting(sheet, colWidthsMap, rowHeightsMap) {
     if (!sheet) return;
     // Set banner row heights
     sheet.setRowHeight(1, rowHeightsMap?.[1] || 30);
     sheet.setRowHeight(2, rowHeightsMap?.[2] || 20);
     // Apply explicit column widths
     if (colWidthsMap) {
       Object.keys(colWidthsMap).forEach(col => {
         sheet.setColumnWidth(parseInt(col, 10), colWidthsMap[col]);
       });
     }
   }
   ```

---

## 8. User Input Sanitization & Whitespace Resilience Patterns

### The Accidental Whitespace Bug
When users enter data in raw sheets (such as payer names, categories, or transaction IDs), leading/trailing spaces (`"Muhanad "`) and non-breaking spaces (`\u00A0` often introduced by copy-pasting from web portals or banking apps) cause downstream lookups and formula evaluations (`SUMIFS`, `VLOOKUP`, `MATCH`) to fail silently.

### Production Pattern: 3-Tier Defense

```mermaid
graph TD
    A[User Enters / Pastes Text in Input Sheet] --> B[Tier 1: onEdit Real-Time In-Place Sanitize]
    B --> C[Cleans \u00A0 & Trailing Spaces in Sheet Cell]
    C --> D[Tier 2: Formula-Level Wildcards]
    D --> E[SUMIFS / MATCH with * & TRIM Match Instantly]
    E --> F[Tier 3: Defensive Procedural Reads]
    F --> G[All Scripts Run .replace /[\u00A0\s]+/g, ' ' .trim()]
```

1. **Tier 1: In-Place Real-Time `onEdit` Range Interceptor**:
   Automatically clean any stray spaces as soon as user types or pastes into the column:
   ```javascript
   if (sheetName === CONFIG.SHEETS.TRANSACTIONS) {
     const startCol = e.range.getColumn();
     const endCol = startCol + (e.range.getNumColumns ? e.range.getNumColumns() : 1) - 1;
     if (startCol <= 4 && endCol >= 4 && e.range.getLastRow() >= 2) {
       const startRow = Math.max(2, e.range.getRow());
       const numRows = e.range.getLastRow() - startRow + 1;
       const targetRange = sheet.getRange(startRow, 4, numRows, 1);
       const values = targetRange.getValues();
       let changed = false;
       for (let r = 0; r < values.length; r++) {
         const orig = String(values[r][0] || "");
         if (orig) {
           const clean = orig.replace(/[\u00A0\s]+/g, " ").trim();
           if (clean !== orig) {
             values[r][0] = clean;
             changed = true;
           }
         }
       }
       if (changed) targetRange.setValues(values);
     }
     return;
   }
   ```

2. **Tier 2: Declarative Formula Wildcard Protection**:
   In summary views (`Monthly Overview`), generate formulas using `*` wildcards:
   ```excel
   =SUMIFS(Transactions!$E:$E, Transactions!$D:$D, "*" & TRIM($A4) & "*", ...)
   ```

3. **Tier 3: Batch Self-Healing Sanitizer**:
   Before updating dashboards or running reconciliation, execute a batch scrubber across all user sheets:
   ```javascript
   function sanitizeAllPayerNames(ss) {
     [
       { sheet: CONFIG.SHEETS.TRANSACTIONS, col: 4 },
       { sheet: CONFIG.SHEETS.INSTALLMENTS, col: 9 }
     ].forEach(cfg => {
       const s = ss.getSheetByName(cfg.sheet);
       if (!s || s.getLastRow() < 2) return;
       const range = s.getRange(2, cfg.col, s.getLastRow() - 1, 1);
       const vals = range.getValues();
       let mod = false;
       for (let i = 0; i < vals.length; i++) {
         const orig = String(vals[i][0] || "");
         if (orig) {
           const clean = orig.replace(/[\u00A0\s]+/g, " ").trim();
           if (clean !== orig) { vals[i][0] = clean; mod = true; }
         }
       }
       if (mod) range.setValues(vals);
     });
   }
   ```

