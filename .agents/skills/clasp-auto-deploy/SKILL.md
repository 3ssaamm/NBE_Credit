---
name: clasp-auto-deploy
description: >-
  Automatically validates and pushes Google Apps Script changes (Code.js, .gs files,
  and appsscript.json) to the live Apps Script project using clasp. Use automatically at
  the end of every task that modifies Apps Script files without waiting for the user to ask.
---

# Clasp Auto-Deploy Workflow

All code changes to `Code.js`, `.gs` files, or `appsscript.json` in this workspace must be pushed to the bound Google Sheets Apps Script project immediately after editing and testing — never leave updated script code local-only.

## Steps

1. **Syntax Check**:
   Validate JavaScript syntax before pushing:
   ```powershell
   node -c Code.js
   ```
2. **Run Local Verification**:
   Run the automated test suite (`npm test`) whenever modifying logic in `Code.js` or `src/*.js`:
   ```powershell
   npm test
   ```
3. **Push to Google Apps Script**:
   Deploy directly from the repository root (where `.clasp.json` resides):
   ```powershell
   clasp push -f
   ```
   *(Fallback if `clasp` is not globally on PATH: `npx @google/clasp push -f`)*
4. **Confirm Deployment**:
   Verify `clasp push -f` exits with code `0` and lists the pushed files (`appsscript.json`, `Code.js`), then report the push status in the final response.
