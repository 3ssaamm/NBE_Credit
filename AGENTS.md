# Workspace Rules — NBE Credit Card Tracker

## Mandatory Automatic Apps Script Deployment (`clasp push`)

Whenever you create, edit, or modify `Code.js`, any `.gs` file, or `appsscript.json` in this repository:

1. Validate JavaScript syntax before deploying:
   ```powershell
   node -c Code.js
   ```
2. Run the project test suite before deploying (`npm test` must pass):
   ```powershell
   npm test
   ```
3. Automatically deploy the changes to the live Google Apps Script project without waiting to be asked:
   ```powershell
   clasp push -f
   ```
4. Confirm in your final response that `clasp push -f` succeeded and list the pushed files.
