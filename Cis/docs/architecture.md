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

## Projects (Clean Architecture)

Unlike GENESIS (one API project, layers by folder), CIS separates the layers into assemblies, so the compiler enforces the dependency
direction **Domain ← Application ← Infrastructure ← Api**:

| Project | Contains | References |
|---|---|---|
| `Cis.Domain` | entities (`UserAccount`), business rules (`AccessRight`, `LegacyCrypt`), domain errors (`BusinessRuleException` family, `NotConfiguredException`) | nothing (BCL only) |
| `Cis.Application` | use cases (`…Service`), ports (`I…Repository`, `IDbSessions`, `ITokenIssuer`, `ILoginAttempts`, `ILoginLog`, `IErrorLog`, `IReportExporter`, `IFileDrop`, `IEmailSender`), request / response models, validators | Domain, FluentValidation, Microsoft.Extensions abstractions |
| `Cis.Infrastructure` | adapters: Dapper repositories, `SqlDbSessions` / `DbConnections`, stored-procedure helpers, Excel (ClosedXML) / PDF (MigraDoc) writers, MailKit, file drop, login attempt cache / log file, error log, health check | Application, Dapper, SqlClient, ClosedXML, PDFsharp, MailKit |
| `Cis.Api` | controllers, `Program.cs` (composition root), JWT (`TokenService` : `ITokenIssuer`), `[RequireAccess]`, `ApiExceptionHandler`, `ValidationFilter`, Swagger | Application (+ Infrastructure for registration only) |

`IDbSession` (Application) carries no SQL types; the infrastructure's `ISqlSession` adds the connection and transaction, and
repositories reach them with `session.Sql()`. Service tests therefore use a plain fake session.

## Layers inside a screen

| Layer | Lives in | Knows about | Must not |
|---|---|---|---|
| Controller | `Cis.Api/Controllers/<Screen>Controller.cs` | the application service, request / response models, `[RequireAccess]` | use Cis.Infrastructure, Dapper, SqlClient, sessions; validate by hand; `try/catch` |
| Validator | `Cis.Application/<Module>/<Screen>/…Validators.cs` (`RequestValidator<T>`) | the request model | read the database |
| Service | `Cis.Application/<Module>/<Screen>/…Service.cs` | repository **interfaces**, `IDbSessions`, domain rules, `ILogger` | know SQL, HTTP or file formats |
| Repository | interface in `Cis.Application/<Module>/<Screen>/`, class in `Cis.Infrastructure/<Module>/` | `IDbSession`, procedure names | hold rules or commit |

`tests/Cis.Api.Tests/ArchitectureTests.cs` enforces it: assembly references (Domain uses only the BCL; Application references no
Infrastructure / Api / SqlClient / Dapper / ClosedXML / PDFsharp / MailKit / ASP.NET; Infrastructure does not reference Api),
controllers take application services only and contain no data access or `try/catch`, services depend on abstractions, every
repository sits in Infrastructure behind an Application interface, every Application port is registered, every `…Request` lives in
Application and has a validator, every screen endpoint has `[RequireAccess]`, only login / change password are anonymous.
XML documentation is mandatory for every public member of every project (`CS1591` is an error).

Errors: services throw `BusinessRuleException` (400), `NotFoundException` (404), `ConflictException` (409), `UnauthorizedException`
(401), `NotConfiguredException` (503); `ApiExceptionHandler` answers Problem Details with a `message` (the desktop's MsgBox text) and
writes unexpected errors to `HISMaintenance.dbo.Tbl_ErrorLog` (SysName 'CIS'), like `Medix.ErrorLog`.

## Migration rules (agreed 2026-10-31)

- Screens keep the desktop's functionality and logic: same fields, buttons, validations, messages, calculations, tables read / written.
  Verified per screen with the `docs/screens/<form>.md` checklist and an SQL comparison against the VB6 app (LocalRun).
- No existing database object is changed; only `cis.*` procedures and synonyms are added.
- Layout follows the GENESIS web style (same fields and order, grouped in cards / tabs / dialogs); desktop-only mechanisms are replaced
  (Outlook → SMTP, Crystal / .xlt → generated Excel / PDF, shares → download + optional server copy).
- **Bugs found in the VB6 logic are reproduced as they are and logged** under "Known desktop issues" in the screen's doc, for the business
  to decide later. Only security problems and features that cannot work at all are fixed, and each is listed as a "Deliberate difference".

## Cross-cutting services (Phase 0)

| Need | Use | Desktop equivalent |
|---|---|---|
| Login / logout / change password | `Cis.Application/Auth/Login/LoginService` | frmLogin, MDIMain.Logout_Click / mnuCloseConnection_Click |
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
3. `Cis.Application/<Module>/<Screen>/` models + validators + repository interface + service (register in
   `Cis.Application/DependencyInjection.cs`); repository class in `Cis.Infrastructure/<Module>/` (register in
   `Cis.Infrastructure/DependencyInjection.cs`); controller in `Cis.Api/Controllers` with `[RequireAccess]`.
4. UI page under `src/app/features/<module>/`, route with `data: { access, legacyForm }`, set `migrated: true` in `layout/menu.ts`,
   regenerate `docs/menu-mapping.md` (`node scripts/menu-mapping.mjs`).
5. Tests: service (fake repositories), DB (rolled back), UI spec, e2e step. Then compare with the VB6 desktop (`LocalRun`).
