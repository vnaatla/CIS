# Architecture

`Cis` (ASP.NET Core 10 Web API) and `CisUI` (Angular 22) replace the VB6 desktop **MedixCIS** (`CIS VB/medixcis.vbp`). They use the
same SQL Server databases as the desktop, so the desktop, the HIS desktop and GENESIS keep working during the migration. The design is
the GENESIS one (`GENESIS/trinath18/Genisis/docs/architecture.md`); this page records what CIS does the same and what differs.

## System context

```mermaid
flowchart LR
    user([Browser]) -->|HTTPS| ui[CisUI<br/>Angular build in wwwroot]
    ui -->|JSON + JWT| api[Cis.Api<br/>ASP.NET Core 10]
    api --> cis[(CISDB<br/>main)]
    cis -. synonyms cis.Maint_* .-> maint[(HISMaintenance)]
    cis -. synonyms cis.Hdb_* .-> his[(HISDB)]
    cis -. synonyms cis.Ws_* / Pay_* .-> ws[(CISWorksheet / CISPayment)]
    api -. optional .-> sun[(SUN Systems)]
    api -. optional .-> rep[(Jupiter reporting replica)]
    api --> files[/Server folders: Claim ASCII, Bordx, PV, LG/]
    desktop([VB6 MedixCIS]) --> cis
    desktop --> maint
    desktop --> his
```

- Stateless API: identity travels in a JWT (`usr_access` = `USR.USRAccess`), nothing is kept in memory between requests except the
  wrong-password counter (see `docs/screens/frmLogin.md`).
- One connection (CISDB) per request; the other databases through `cis.*` synonyms (`sql/main/000_setup.sql`), so a save that touches
  several databases is **one transaction** (the desktop used up to eight connections and no transaction).
- Production: the API serves the Angular build from `wwwroot` (one deployable). Development: `ng serve` proxies `/api` to the API.

## API layers

| Layer | Lives in | Knows about | Must not |
|---|---|---|---|
| Controller | `Controllers/<Screen>Controller.cs` | the service, request / response models, `[RequireAccess]` | touch Dapper / SqlClient / `DbConnections` / `IDbSessions`, validate by hand, `try/catch` |
| Validator | `<Module>/<Screen>/…Validators.cs` (`RequestValidator<T>`) | the request model | read the database |
| Service | `<Module>/<Screen>/…Service.cs` | repository **interfaces**, `IDbSessions`, rules, `ILogger` | reference a concrete repository or `HttpContext` |
| Repository | `<Module>/<Screen>/…Repository.cs` (interface + class) | `IDbSession`, procedure names | hold rules or commit |
| Cross-cutting | `Auth/`, `Common/`, `Data/`, `Audit/`, `Documents/`, `Files/`, `Email/`, `Lookups/` | everything above uses them | depend on a screen |

`tests/Cis.Api.Tests/ArchitectureTests.cs` enforces it for **all** code (CIS starts layered, unlike GENESIS where older modules are
exempt): no data access or `try/catch` in controllers, services depend on abstractions, every repository / service interface is
registered, every `…Request` has a validator, every screen endpoint has `[RequireAccess]`, only login / change password are anonymous.
XML documentation is mandatory for every public API member (`CS1591` is an error).

Errors: services throw `BusinessRuleException` (400), `NotFoundException` (404), `ConflictException` (409), `UnauthorizedException`
(401), `NotConfiguredException` (503); `ApiExceptionHandler` answers Problem Details with a `message` (the desktop's MsgBox text) and
writes unexpected errors to `HISMaintenance.dbo.Tbl_ErrorLog` (SysName 'CIS'), like `Medix.ErrorLog`.

## Cross-cutting services (Phase 0)

| Need | Use | Desktop equivalent |
|---|---|---|
| Login / logout / change password | `Auth/Login/LoginService` | frmLogin, MDIMain.Logout_Click / mnuCloseConnection_Click |
| Menu rights | `AccessRight` + `[RequireAccess(n)]`; UI `access.ts` + `accessGuard` | `Medix.CheckAccess(n)` |
| Action log | `POST /api/audit/screen-opened`, UI `ScreenAuditService` (route `data.legacyForm`) | `ActionLogUpdate(12, "frmName")` in every Form_Load |
| Reports / exports | `IReportExporter` (Excel via ClosedXML, PDF via MigraDoc, CSV) with optional legacy templates (`Reports:TemplateRoot`) | `ModListviewtoExcel.ExportToExcel*`, `.xlt` templates, Crystal / DataReport |
| Files to shares | `IFileDrop` (`Files:Folders`) | writes to `SRVPath`, `BordxPath`, `MCOPath`, `DocPath` |
| E-mail | `IEmailSender` (MailKit; SMTP or a pickup folder) | Outlook automation, CDO in `ErrorLog` |
| Drop-down lists | `GET /api/lookups/*` | the combo loaders (`select PAYCode, PAYCompanyName from PAY`, ...) |
| External DBs | `DbConnections.SunSystems()` / `Reporting()` | `MediConnSUNSystem.UDL`, the `MediNeptune*` (Jupiter) connections |

## SQL

`sql/main/NNN_<screen>.sql`, one file per screen, `CREATE OR ALTER PROCEDURE cis.<Screen>_<Action>`; `sql/cis-main.sql` is generated
(a test fails when it is stale). Additive only (a test rejects `ALTER/DROP TABLE` and non-`cis` objects). Legacy procedures that splice a
raw `@strAppend` into dynamic SQL are reproduced as parameterised procedures. See `sql/README.md`.

## UI

| Need | Use | Notes |
|---|---|---|
| Styling | `src/styles.css` = GENESIS `index.css` verbatim + a "CIS additions" section | same classes / variables, so both apps look the same; **no inline styles**, no component library |
| Button | `button[appButton]` (`variant`, `busy`) | GENESIS Button |
| Field | `app-field` wrapping the control | GENESIS Field |
| Date | `app-date-input` (ControlValueAccessor, dd/mm/yyyy + picker) | VB6 MaskEdBox `##/##/####` |
| Drop-down | `app-code-select` + `LookupsService` | "code - name" combos |
| Table | `app-data-grid` (sort, select, check-all, empty state) | VB6 ListView + ColumnClick |
| Tabs | `app-tabs` + `ng-template[appTab]` | VB6 SSTab |
| Dialog | `app-modal` (focus trap, Escape) | pop-up forms (frmSearchMBM, frmSearchICD, ...) |
| Messages | `NotifyService.notify()` / `ask()` | MsgBox / MsgBox vbYesNo |
| Screen chrome | `app-page-shell`, `app-action-bar`, `app-stat` | form caption, command-button frame, record counters |
| Pure helpers | `shared/lib` (masks, format, download / CSV) | `KeyNumber`, `Format`, `ExportListViewtoExcel` |

State: signals; forms: Reactive Forms (template-driven only for the login box); routes are lazy, one folder per desktop top menu, each
screen route carries `data: { access, legacyForm }`.

## Adding a screen

1. `docs/screens/<form>.md`: the parity checklist (controls, events, SQL, messages, outputs, pop-ups) from the VB6 form.
2. `sql/main/NNN_<screen>.sql` → rebuild `cis-main.sql`.
3. `<Module>/<Screen>/` models + validators + repository + service, register in `CisServiceCollectionExtensions`, controller with `[RequireAccess]`.
4. UI page under `src/app/features/<module>/`, route with `data: { access, legacyForm }`, set `migrated: true` in `layout/menu.ts`,
   regenerate `docs/menu-mapping.md` (`node scripts/menu-mapping.mjs`).
5. Tests: service (fake repositories), DB (rolled back), UI spec, e2e step. Then compare with the VB6 desktop (`LocalRun`).
