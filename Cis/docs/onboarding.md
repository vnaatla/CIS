# Onboarding: run CIS locally

Prerequisites: .NET SDK 10, Node.js ≥ 22.22.3 (Angular 22), SQL Server with the CIS / HIS databases (see `AGENTS.md`: local
default instance `localhost`, Windows authentication, TCP disabled → use `(local)`).

## 1. Database (once, and after any change in `Cis/sql/main`)

```powershell
sqlcmd -S localhost -E -d CISDB -i Cis\sql\cis-main.sql
```

Additive only: creates the `cis` schema, synonyms and procedures in CISDB.

## 2. API (http://localhost:5100, Swagger at /swagger)

```powershell
cd Cis\src\Cis.Api
$env:Auth__JwtKey = '<any 32+ character secret>'   # keeps tokens valid across restarts
dotnet run --launch-profile http
```

`appsettings.json` points `ConnectionStrings:CisDb` at `(local)` / CISDB. `appsettings.Development.json` writes the login log to
`D:\CIS\LocalRun\FileLogin\LoginCIS.txt` (the file the VB6 LocalRun copy writes too) and e-mails to `D:\CIS\LocalRun\MailPickup`.
Health: `GET /api/health` (Healthy = CISDB reachable and `cis-main.sql` deployed).

## 3. UI (http://localhost:4200)

```powershell
cd CisUI
npm ci
npm start        # ng serve, proxies /api to :5100
```

Log in with a `HISMaintenance.dbo.USR` account (local dev: see AGENTS.md).

## 4. Checks

```powershell
# API: unit + architecture + integration; DB tests run when CIS_TEST_DB is set (they roll back)
$env:CIS_TEST_DB = 'Server=(local);Database=CISDB;Integrated Security=true;Encrypt=false;TrustServerCertificate=true;Connect Timeout=60'
dotnet test Cis\Cis.sln

# UI
cd CisUI; npm run lint; npm test; npm run build

# Browser smoke test (API + UI running)
$env:E2E_PASSWORD = '<local dev password>'; npm run e2e
```

After editing `Cis/sql/main/*.sql`: `$env:CIS_WRITE_SQL='1'; dotnet test Cis\Cis.sln --filter Combined_script_is_up_to_date`, then rerun
step 1. After editing `CisUI/src/app/layout/menu.ts`: `node scripts/menu-mapping.mjs` (in CisUI) to regenerate `docs/menu-mapping.md`.

## Configuration reference (`appsettings.json`)

| Key | Purpose |
|---|---|
| `ConnectionStrings:CisDb` | CISDB (everything else through synonyms) |
| `ConnectionStrings:Reporting` | Jupiter / Neptune replica for the "(Jupiter)" menu items; empty = CisDb |
| `ConnectionStrings:SunSystems` | SUN Systems accounting DB (A/C Export > SUN screens); empty = those screens answer "not configured" |
| `Auth:JwtKey` / `TokenHours` | token signing key (set via env `Auth__JwtKey`) and lifetime |
| `Session:LoginLogPath` | desktop's `DocPath\FileLogin\LoginCIS.txt`; empty = not written |
| `Reports:TemplateRoot` | copies of `\\svr-doc\HISREPORTS` templates as .xltx; empty = layouts generated in code |
| `Files:Folders:<name>` | server folders the desktop wrote to (`ClaimAscii`, `Bordx`, `McoInvoice`, `PaymentVoucher`, `Outpatient`, `ClinicPvs`, `PolicyWording`, `LgLetters`) |
| `Email:*` | SMTP (`Host`, `Port`, `User`, `Password` via env) or `PickupDirectory` |
| `Cors:Origins` | dev UI origin |
