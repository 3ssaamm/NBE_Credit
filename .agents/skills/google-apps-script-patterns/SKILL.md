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
