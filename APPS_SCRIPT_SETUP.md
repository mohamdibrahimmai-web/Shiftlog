# ShiftLog Apps Script setup

1. Copy `Code.gs` into the Google Apps Script project attached to the ShiftLog spreadsheet.
2. In **Project Settings → Script properties**, add `SHIFTLOG_PIN` with the team PIN. This secret is not stored in `index.html` or this repository.
3. Optionally add `SESSION_TTL_SECONDS` (60–21600). The default is 14400 seconds (4 hours).
4. Deploy the script as a web app that runs as the spreadsheet owner and is accessible to the people who use ShiftLog. Copy its `/exec` URL into `API_URL` in `index.html` if it changed.
5. Redeploy after any Apps Script code change.

The browser holds the returned token only in `sessionStorage`, so closing the tab clears it. Every action other than `login`, including the three read operations, requires a valid short-lived token.
