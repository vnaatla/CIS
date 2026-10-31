# frmLogin + MDIMain session handling → `/login`, `POST /api/auth/*`

Phase 0. Source: `CIS VB/frmLogin.frm`, `MDIMain.frm` (Logout_Click, mnuCloseConnection_Click, MDIForm_Load), `Medix.bas`
(ActionLogUpdate, ErrorLog, CheckAccess), `SubMain.bas`.

## Controls

| Desktop | Web |
|---|---|
| txtUserName (Enter → password) | User Name (Enter → password), max 10 |
| txtPassword (Enter → OK) | Password (Enter submits), max 10 |
| cmdOK | **OK** |
| cmdCancel (Unload) | not needed (close the tab) |
| cmdChange → form grows to show txtnewpass / txtnewconfirm | **Change Password** opens the same two fields under the login box |
| cmdSave | **Save** (in change mode) |

## Behaviour checklist

| # | Desktop (cmdOK_Click / cmdSave_Click / cmdChange_Click) | Web | ✓ |
|---|---|---|---|
| 1 | Empty user name → "Please Enter User Name" | `LoginRequestValidator` | ✓ |
| 2 | `Select * from USR Where USRCode = ... And USRStatus = 'A'`; none → "Invalid User Name!! Please enter again.." | `cis.Login_GetActiveUser` | ✓ |
| 3 | `UCase(Crypt(pwd)) <> UCase(USRPassword)` → "Invalid Password!! Please enter again..", counter + 1 (same user) | `LegacyCrypt.Matches`, `LoginAttempts` | ✓ |
| 4 | counter = 3 → `UPDATE USR SET USRStatus = 'S'` whatever the password, "Your Account Been Suspended" | `cis.Login_Suspend` | ✓ |
| 5 | Success: Access = USRAccess (menu rights) | JWT claim `usr_access` | ✓ |
| 6 | `OutputToTxtfile`: `"Login","USER",#yyyy-mm-dd#,"hh:mm:ss am",` to `DocPath\FileLogin\LoginCIS.txt` | `LoginLogFile` (`Session:LoginLogPath`), same line | ✓ |
| 7 | `Update USR set CIScnt = ISNULL(CIScnt,0) + 1, actpass = '<plain password>'` | `cis.Login_Record`: CIScnt + 1 only — **actpass not written** (D1) | ✓ |
| 8 | `Insert into Tbl_USRTrack (USRCode, SYSName, LoginTime, LoginDate) Values (user, 'CIS', CStr(Time), CStr(Date))` | `cis.Login_Record`, server time as `h:mm:ss tt` / `dd/MM/yyyy` (production rows' format) | ✓ |
| 9 | MDIMain Form_Load → `ActionLogUpdate(12, "MDIMain")` | login page posts `screen-opened MDIMain` | ✓ |
| 10 | frmLogin Form_Load → `ActionLogUpdate(12, "frmLogin")` with no user yet | not reproduced (anonymous row of no use) (D3) | — |
| 11 | cmdChange with empty user and password → "Please Enter User ID and Password" | same, in the page | ✓ |
| 12 | cmdSave checks: "Please Enter User Name", "Please Enter Password", "Please Enter Valid New Password", "Please Enter Valid Re-Type Password", "Password Verification Is In Valid" | `ChangePasswordRequestValidator`, same order | ✓ |
| 13 | cmdSave: `Select * From USR Where USRPassword = '<plain>' And USRCode = ...` on **HISDB**, set `USRPassword = UCase(new)` → "No Record Match For User ID or Password" / "Password Update Sucessfully" | checks the scrambled password on HISMaintenance.USR and stores `UCase(Crypt(new))` (D2) | ✓ |
| 14 | Logout / Close Connection: `Update USR set CIScnt = CIScnt - 1 Where ... And USRStatus = 'A'` | `cis.Login_Logout` | ✓ |
| 15 | `Medix.ErrorLog`: `Tbl_ErrorLog ('CIS', form, method, error, GETDATE())`, CIScnt - 1, close the app, e-mail IT | `ApiExceptionHandler` → `cis.ErrorLog_Write`; the session survives, so no CIScnt - 1 and no e-mail (D4) | ✓ |

## Deliberate differences

- **D1** The desktop copied the typed password in clear text to `USR.actpass` on every login. Not reproduced (password leak).
- **D2** The desktop's change-password compared the typed password in clear text with the scrambled `USRPassword` of **HISDB**.USR and
  stored the new password unscrambled, so it either failed or locked the user out. The web does what the HIS desktop / GENESIS do.
- **D3** frmLogin's own action-log row (empty user) is not written.
- **D4** Errors no longer end the session, so the error log does not decrement CIScnt or e-mail IT (Serilog keeps the full error).
- **D5** The wrong-password counter lived in the open login window. The web keeps it per user code + client address for 30 minutes
  (memory of the API process); a successful login clears it.
- **D6** The desktop wrote `Tbl_USRTrack` time / date with the PC's regional settings (e.g. `21:10:35` / `31-10-2026` on this PC); the web
  always writes the format of the production rows (`9:10:35 PM` / `31/10/2026`).

## Tests

`LoginTests.cs` (rules, messages, file line, suspend after three), `DatabaseTests.cs` (procedures on the real tables, rolled back),
`ApiIntegrationTests.cs` (HTTP, Problem Details), `login.page.spec.ts`, `session.service.spec.ts`, `e2e/run.mjs`.
