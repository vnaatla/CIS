# Browser smoke test

`run.mjs` drives headless Chrome / Edge (puppeteer-core, the browser installed on the PC) against a running API (:5100) and UI (:4200):
login redirect, the desktop's wrong-password message (one attempt only: three suspend the account), login, layout (rail, home cards),
module overview and sidebar, unknown-screen fallback, logout.

```powershell
$env:E2E_PASSWORD = '<password of E2E_USER>'   # required
$env:E2E_USER = 'DEVADMIN'                      # default
$env:E2E_URL = 'http://localhost:4200'          # default
$env:E2E_BROWSER = 'C:\...\chrome.exe'          # optional; Chrome / Edge are found automatically
npm run e2e
```

While building a phase, run only its steps (one login, no wrong-password attempt, far less load on the local SQL Server):

```powershell
$env:E2E_STEPS = 'claim-bordx'   # modules in e2e/steps (comma-separated)
node e2e/only.mjs
```

It writes what a real login writes on the local databases: `USR.CIScnt` + 1 then - 1, one `Tbl_USRTrack` row, one `LoginCIS.txt`
line and `tblActionLog` rows. Run it against local / test databases only.
