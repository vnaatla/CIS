---
agent: devin-local
session: swamp-oatmeal
created: 2026-10-31T16:00:11Z
---
# CIS VB6 → ASP.NET Core Web API + Angular migration

Rebuild every form of the MedixCIS VB6 project (138 forms, 26 DataReports, ~396k code lines) as a standalone ASP.NET Core 10 Web API (Cis.Api) plus Angular 22 UI (CisUI) in D:\CIS. Both work against the existing CIS/HIS databases, and the UI reuses the GENESIS architecture and CSS/markup conventions so the two apps look and behave the same.

## 1. Decisions (agreed + recommended)

| Topic | Decision |
|---|---|
| Location | Standalone in `D:\CIS`: `Cis\` (API solution) + `CisUI\` (Angular). The legacy `CIS VB\` and `LocalRun\` stay untouched for side-by-side parity checks. |
| Scope | **Everything**: all 138 forms + 26 DataReports in `medixcis.vbp`, including hidden, disabled, no-handler and never-opened forms (inventory in §6). Forms that sit in the folder but **not** in the `.vbp` (old copies: frmDentalInvoice, frmInvDentalCln, frmInvoiceProcessing, ...) are listed in §6.6 for confirmation, not built by default. |
| Variant screens | Keep every desktop menu item as its own route/menu entry (1:1). Variants (Simple / Jupiter / GST / Zeroised / per-client Print Bordx and Claim ASCII) share **one Angular component + one API service** and are selected by a `variant` parameter. |
| Reports / Excel ("which is best") | **Rebuild layouts in code, with an optional template override.** Default: `.xlsx` via ClosedXML and PDF via PDFsharp-MigraDoc (MigraDoc is already used by GENESIS), with column/row layouts taken from the VB `ExportToExcel*` / Crystal code. If the `\\svr-doc\HISREPORTS` templates are supplied later (saved as `.xltx`), the `ReportTemplateStore` fills the template instead. Why: the templates aren't on this PC, the Crystal 7 runtime can't run server-side, and layouts built in code can be versioned and unit-tested. |
| Backend stack | ASP.NET Core **10** (LTS; .NET 8 support ends 10 Nov 2026; SDK 10.0.400 is installed), Dapper, Microsoft.Data.SqlClient, FluentValidation, JWT, Serilog, Swagger, ExcelDataReader (receipt upload), ClosedXML, PDFsharp-MigraDoc, MailKit (replaces Outlook). Pin versions published at least 7 days ago. |
| Frontend stack | Angular **22.2.x** (pin 22.2.3, released 14 Oct 2026): standalone components, signals, Reactive Forms, functional guards and interceptors, lazy routes per module. **No component library**: GENESIS `index.css` is ported as the global stylesheet and its class names are reused, so CIS looks identical. |
| Architecture | Same patterns as GENESIS (Controller → Validator → Service → Repository → `cis.*` stored procedures; unit of work through `IDbSessions`; central `ApiExceptionHandler`), but split into 4 projects (Cis.Domain / Cis.Application / Cis.Infrastructure / Cis.Api) so the compiler enforces the layers; architecture tests check the project references. |
| Migration rules | Same functionality / logic / messages as each VB6 form; no existing DB object changed (only `cis.*` added). VB6 bugs are reproduced and logged ("Known desktop issues" in the screen doc); only security problems and features that cannot work are fixed ("Deliberate differences"). |

### Defaults I'll use unless you say otherwise (open questions, §10)
- **SUN Systems** (frmSUNSystem / SUNGSTRpt / Allocation*) read an external SUN accounting DB (`MediConnSUNSystem.UDL`, tables `SALFLDG<UHC|FW|MED|GV|TL|UHI>`, `VIEW_GST_P2`) that isn't on this PC. I'll add a separate `SunSystems` connection string. Until it's provided, those screens return a clear "SUN Systems database not configured" message.
- **"Jupiter" variants** read the Neptune/Jupiter replica (the `MediNeptune*` connections, which are commented out in `SubMain`, so the desktop items fail today). They'll get a `Reporting` connection string that defaults to the main databases.
- **Email** (3 screens, Outlook automation): SMTP via MailKit, configured under `Email:Smtp`, with a preview list before sending and a send log.

## 2. What was analysed (facts)

- **Startup** `SubMain.bas`: reads `SVRSuper.ini`, opens 8 UDL connections (HISDB, HISDMArchive, HISWorkSheet, HISPayment, HISMaintenance, CISDB, CISWorksheet, CISPayment), then shows frmLogin.
- **Login** `frmLogin`:
  - checks `HISMaintenance.dbo.USR` (`USRStatus='A'`, legacy `Crypt`; GENESIS `LegacyCrypt.cs` can be reused);
  - suspends the account after 3 bad tries (`USRStatus='S'`);
  - `CIScnt += 1`, `actpass`, inserts `Tbl_USRTrack` (SYSName 'CIS'), appends `DocPath\FileLogin\LoginCIS.txt`;
  - "change password" panel (`cmdSave`).
- **Logout / Close Connection**: `CIScnt -= 1`, close all connections.
- **Rights**: `Medix.CheckAccess(n)` = character *n* of `USR.USRAccess` (the same string GENESIS uses; CIS positions 113–179). `ActionLogUpdate` writes `hisdb..tblactionlog` (module 'CIS'). `ErrorLog` writes the error table.
- **Data access**: about 2,184 stored-procedure calls plus inline SQL. Usage by connection: HISDB 1,983, CISWorksheet 1,604, CISDB 1,161, CISPayment 280, SUN 140, HISMaintenance 90, HISWorkSheet 36. Three-part names (`CISDB..`, `HISDB..`) are common, so everything is on one instance. Compatibility levels: CISDB 120, CISWorksheet 100, CISPayment 160.
- **Outputs**:
  - about 90 Excel templates (`crdPath\CISTemplate\*`, `Template\{PaymentVoucher,MBMTracking,MCOInvoice,Bordx,Worksheet,...}`) and about 100 `ExportToExcel*` routines (`ModListviewtoExcel.bas`);
  - about 25 Crystal reports (ClaimBordx Summary/Detail variants, MCOInvoice/MCOFeeListing, UHC/MED LG Preview);
  - a Word `coverletter.dot`;
  - 26 DataReports (maintenance list printing);
  - ASCII files to `SRVPath\ClinicalClaimASCII`, PV and cheque printing, `DocPath\Outpatient\`, `ClinicPVS\`, `CLNPolicyWording\`, `BordxPath`, `MCOPath`.
- **Size hot-spots**:
  - MBM Tracking family (4 forms of 30–34k lines; FrmNewMBMTracking alone has 210 report-type entries over about 20 categories);
  - Claim Tracking family (7 forms of 13–16k lines each);
  - Invoice entry and adjustment family (3.5–6k lines each);
  - frmLGGeneration (6.6k), frmMaintenancePlan (5.6k).

## 3. Target solution layout

```
D:\CIS
  Cis\
    Cis.sln
    (Clean Architecture, 4 projects, decided 2026-10-31 — dependency direction Domain ← Application ← Infrastructure ← Api)
    src\Cis.Domain\          entities (UserAccount), rules (AccessRight, LegacyCrypt), domain errors — no dependencies
    src\Cis.Application\     <Module>\<Screen>\ services, ports (I...Repository, IDbSessions, ITokenIssuer, IReportExporter, IFileDrop, IEmailSender ...),
                             request/response models, validators; shared Lookups, Audit, Documents abstractions
    src\Cis.Infrastructure\  Dapper repositories on cis.* procs, SqlDbSessions/DbConnections (CisDb + Reporting + SunSystems), ClosedXML / MigraDoc writers,
                             MailKit, file drop, login attempts / LoginCIS.txt, Tbl_ErrorLog, health check
    src\Cis.Api\             Controllers (route api/<module>/<screen>), Program.cs (composition root), TokenService, RequireAccess, ApiExceptionHandler, ValidationFilter
    sql\
      main\NNN_<screen>.sql     CREATE OR ALTER PROCEDURE cis.<Screen>_<Action> (in CISDB)
      000_setup.sql             cis schema + synonyms cis.Hdb_*, cis.Ws_*, cis.Pay_*, cis.Maint_*, cis.HWs_*
      cis-main.sql              generated combined script (test fails if stale)
    tests\Cis.Api.Tests\        unit, service (fake repos), integration (WebApplicationFactory), architecture, DB (rolled-back transactions), SqlScript
    docs\                       architecture.md, menu-mapping.md (traceability), screens\<form>.md parity checklists, reports.md, onboarding.md
  CisUI\
    src\styles.css              port of GenisisUI/src/index.css (same variables + class names) + CIS additions
    src\app\core\               auth (session signal, access.ts mirror of AccessRight), http interceptors (JWT, 401 → login), notify/ask service, error handler
    src\app\shared\             Button, Field family (Text/Number/Date/Select/Combo/Check/Read/TextArea), DataGrid (sort/select/check/empty state),
                                Tabs, Modal, PageShell/ActionBar/Stat, DateInput (dd/mm/yyyy), CodeSelect, Icon set (copied SVGs), ReportRunner, FileDownload
    src\app\layout\             AppLayout = GENESIS topbar + icon rail + sidebar (collapsible, hover sections), open-windows list
    src\app\features\<module>\  one lazy route folder per desktop top menu
    e2e\                        headless Chrome suites (puppeteer-core, as GENESIS)
```

Ports: API `http://localhost:5100`, UI `http://localhost:4200` (`/api` proxied). This avoids clashing with GENESIS on 5000/5173. In production the API serves the Angular build from `wwwroot` (single deployable, as GENESIS does).

## 4. Cross-cutting design

1. **Database access**:
   - one connection per request to **CISDB**; other DBs through `cis.*` synonyms with configurable names, so multi-DB saves are **one transaction**. The desktop used up to 4 connections without a transaction.
   - Legacy procs that take `@StrAppend` dynamic SQL are never called; they are reproduced as parameterised `cis.*` procs. All inline VB SQL is rewritten with parameters, which removes the SQL injection in the desktop.
   - Procs that touch CISWorksheet tables avoid T-SQL newer than compatibility level 100 where they run in that DB's context. The procs themselves live in CISDB (level 120).
2. **Auth and session parity**:
   - `POST /api/auth/login` covers bad-password counting and the 3-strike suspend, `CIScnt+1`, `Tbl_USRTrack` and the optional `LoginCIS.txt` append (config `Session:LoginLogPath`). It returns a JWT with `usr_access`.
   - `POST /api/auth/change-password`; `POST /api/auth/logout` (`CIScnt-1`; also used for "Close Connection").
   - Screen open → `POST /api/audit/action` (tblactionlog, module 'CIS'). Errors → Serilog + legacy error table.
3. **Access rights**: `[RequireAccess(n)]` on every endpoint and a route guard in the UI. Positions come from the `MDIMain` click handlers (§6). Forms the desktop never opens get the position of their closest sibling, documented in `AccessRight.cs` and `access.ts`, as GENESIS did.
4. **Business rules**:
   - desktop message wording kept exactly (`MsgBox` → `notify()` / `ask()`);
   - KeyPress numeric/char masks → input masks;
   - LostFocus date checks → `DateInput` + validators;
   - server time replaces client `Date/Time`;
   - number generators (`GenerateCaseNo`, `GenerateCLMBordxNo`, `GenerateSPVNo`/`NPVNo`, `GenerateDebitNo`/`CreditNo`, `GeneratePostingINVNo`, `GenSUNPostingBatchNo`, sequence generators) become concurrency-safe procs (`UPDLOCK`/`HOLDLOCK`).
5. **Reports engine** (needed by roughly 60% of screens):
   - `ReportDefinition` = id, category, filters, `cis.*` query proc, columns (header, format, total), Excel layout (title rows, grouping, subtotals, template name) and an optional PDF layout.
   - Tracking screens = filter panel + category + report type + grid (paged, capped) + Export Excel/PDF/CSV + Print.
   - Every `CboReportType`/`CboRptSelection` entry and every `ExportToExcel*` routine is mapped to one definition, so none is lost.
6. **Files**: ASCII exports, PV/cheque/LG documents and upload files are streamed to the browser. When configured (`Files:*` paths matching SVRSuper.ini keys), they are also written to the server folders the desktop used.
7. **Printing**: "Print" buttons give a print-friendly view (`@media print` CSS) or a PDF. `frmPrintScreen` / File > Print = print the current screen. Page Setup = the browser's print dialog.
8. **Long operations** (bordx generation, posting, bulk payment, large tracking reports): synchronous with a request timeout plus row caps first. Background jobs with progress polling for operations that exceed about 30 s on the production data (decided per screen during the build).

## 5. Per-screen migration method (applied to every form, nothing skipped)

1. **Inventory** (`docs/screens/<form>.md`):
   - every control (label, type, mask, default, enabled/visible rules);
   - every event handler and button;
   - every SQL statement / proc call, with tables read and written;
   - every validation and `MsgBox` text, every output (Excel/Crystal/file/print), and every popup the form opens.
   - This becomes the parity checklist.
2. **SQL**: `sql/main/NNN_<screen>.sql` with parameterised `cis.<Screen>_<Action>` procs; regenerate `cis-main.sql`.
3. **API**: models, validator, repository (interface + class), service (rules, unit of work), controller, `AccessRight`, register in the composition root.
4. **UI**: page component(s) and popups as modals, route, sidebar entry, and access in `access.ts`.
5. **Tests**:
   - service tests with fake repositories (rules and messages);
   - DB tests that run the real save in a rolled-back transaction;
   - report-definition tests (columns/totals);
   - an e2e happy path.
6. **Parity check against VB6 LocalRun**: same input in both apps, then compare the affected rows with SQL diff scripts and compare the exported files. Tick off the checklist.

## 6. Complete inventory → web mapping

Route prefix = desktop top menu. *Pos* = `CheckAccess` position. **H** = hidden in the desktop, **D** = disabled, **NH** = menu item without a click handler, **NO** = form never opened by the desktop (built anyway, per the "Everything" scope).

### 6.1 Shell / File / Exit / Windows
| Desktop | Form | Pos | Web |
|---|---|---|---|
| Startup + Login + change password | SubMain, frmLogin | – | `/login`, `POST /api/auth/*` |
| Main window | MDIMain | – | AppLayout (topbar/rail/sidebar), access-filtered menu |
| File > Page Setup (NH) | – | – | browser print setup (documented) |
| File > Print (handler commented out) / frmPrintScreen (NO) | frmPrintScreen | – | "Print screen" action in the topbar |
| Windows (WindowList) | – | – | "Open screens" list (recent tabs) in the topbar |
| Exit > Logout / Close Connection | MDIMain | – | Logout (`CIScnt-1`) |

### 6.2 Membership (`/membership`)
| Menu | Form | Pos | Route |
|---|---|---|---|
| MBM Adjustment | FrmMBMCLNAdj | 113 | `/membership/adjustment` |
| MBM Enquiry | FrmMBMCLNEnq | 113 | `/membership/enquiry` |
| MBM Topup Adj | FrmMBMTopupLmt | 113 | `/membership/topup` |
| MBM Tracking | FrmNewMBMTracking | 114 | `/membership/tracking` |
| MBM Tracking (Simple) | frmNewMBMTrackingSimple | 114 | `/membership/tracking-simple` |
| MBM Tracking (Jupiter) | FrmNewMBMTracking2 | 160 | `/membership/tracking-jupiter` |
| MBM Tracking (Jupiter - Simple) | frmNewMBMTrackingRepSimple | 160 | `/membership/tracking-jupiter-simple` |
| (NO) old MBM Tracking | FrmMBMTracking | 114 | `/membership/tracking-legacy` |
| (NO) MBM Account | FrmMBMAccount | 113 | panel in MBM Enquiry |
| Popups: member search (name / 05 / IC), clinic member registration | frmSearchMBM, frmSearchMBM05, frmSearchMBMIC, frmMemberRegistrationCLN | – | shared modals |

### 6.3 MCO Fees (`/mco-fees`)
| Menu | Form | Pos | Route |
|---|---|---|---|
| MCO Invoice | frmMCOInvoiceNew | 115 | `/mco-fees/invoice` |
| MCO Invoice GST | frmMCOFeeGST | 115 | `/mco-fees/invoice-gst` |
| MCO Invoice GST Zeroised | frmMCOFeeGSTZeroised | 115 | `/mco-fees/invoice-gst-zeroised` |
| Fee Listing | frmMCOFeeListNew | 116 | `/mco-fees/fee-listing` |
| Refund Listing | frmMCORefundListing | 116 | `/mco-fees/refund-listing` |

### 6.4 Invoice (`/invoice`)
| Menu | Form | Pos | Route |
|---|---|---|---|
| Invoice Control (NH) | – | – | nothing to build (documented) |
| Clinic > Clinic | frmInvoice | 117 | `/invoice/clinic` |
| Clinic > Member | frmInvoiceMBM | 118 | `/invoice/member` |
| Clinic > Adjustment | frmInvoiceAdj | 119 | `/invoice/adjustment` |
| Clinic > Adjustment (After Posting) | frmInvoiceAdjAfterPost | 119 | `/invoice/adjustment-after-posting` |
| Clinic > Enquiry | FrmInvoiceEnq | 119 | `/invoice/enquiry` |
| Clinic > Batch Transfer | FrmInvoiceBatchNoUpd | 119 | `/invoice/batch-transfer` |
| Clinic > Adjustment (After Bordx) | frmInvoiceAdjAftBordx | 179 | `/invoice/adjustment-after-bordx` |
| Email Notification (Payment / Providers / Rejection) | frmEmailNotification, frmEmailNotificationProvider, frmEmailRejectedClaim | 117 | `/invoice/email/{payment,providers,rejection}` |
| General Exclusions / (Enquiry) | frmGeneralBenefitExclusion / frmGeneralBenefitExcEnq | 133 / 113 | `/invoice/general-exclusions[/enquiry]` |
| Claim ASCII (Bintulu, Toshiba, Singer, Utusan, OMNI, TIMS, CIMB) | frmClaimDetailsASCII, frmClaimASCIIToshiba, frmClaimASCIISinger, frmClaimASCIIUtusan, frmClaimDataOMNI, frmClaimASCIITIMS, frmClaimASCIICIMB | 119 | `/invoice/claim-ascii/{bintulu,toshiba,singer,utusan,omni,tims,cimb}` |
| (dead handler) Med Providers OMNI | frmMedProviderOMNI | 119 | `/invoice/claim-ascii/omni-providers` |
| (NO) Claims data / Sin Chew Excel | frmClaimsdata, frmClaimsExcelSinChew (layout only) | 119 | `/invoice/claim-ascii/{claims-data,sinchew}` |
| Dental > Dental / Member (H, NH), Adjustment (NH) | – (no form in project) | – | see §6.6 |
| Invoice Update (H) | frmUpdateCLNInvoice | 119 | `/invoice/update` |
| OP Batch/Invoice | frmUpdateInvoiceBatch | 119 | `/invoice/op-batch` |
| Case Transfer | frmMembershipTransfer | 162 | `/invoice/case-transfer` |
| Tracking / Tracking (Jupiter) | FrmInvoiceTracking / FrmInvoiceTracking2 | 161 / 162 | `/invoice/tracking[-jupiter]` |
| MMA/CPT (+ investigation popup) | frmMMASchedule, frmShowInvestigation | 150 | `/invoice/mma-cpt` |
| Worksheet calculation popup (from the invoice forms) | frmWorksheet | – | modal + `api/invoice/worksheet/*` (benefit, reimbursement and cashless calculation, saves) |
| (NO) Worksheet Reg / Adj | frmNewWSReg / frmNewWSAjd | 98→117 / 99→119 | `/invoice/worksheet-registration`, `/invoice/worksheet-adjustment` |
| Popups: ICD search, case list, case pending, member info, invoice detail | frmSearchICD, frmCasList, frmCasePending, FrmINVMBMinfo, frmInvoiceDetPopup | – | shared modals |

### 6.5 Claim / Claim Bordx / LG / Audit / Payment / Receipt / A/C Export / Tracking / Maintenance
| Menu | Form | Pos | Route |
|---|---|---|---|
| Claim > Pass to Claim | frmPass2Claim | 163 | `/claim/pass-to-claim` |
| Bordx > Worksheet > Insert Worksheet | frmWSBordx | 122 | `/claim-bordx/insert-worksheet` |
| Bordx > Worksheet > Insert Invoice (D) | frmWSBordx2 | 122 | `/claim-bordx/insert-invoice` |
| Bordx > Worksheet > Remove Worksheet | FrmBordxRevise | 123 | `/claim-bordx/remove-worksheet` |
| Update Claims Data | frmUpdateClaimsData | 124 | `/claim-bordx/update-claims-data` |
| Print Bordx (H, Crystal) | frmPrintBordx | 124 | `/claim-bordx/print/crystal` |
| Print Bordx (Excel) | frmPrintBordxExcel | 124 | `/claim-bordx/print/standard` |
| Print Bordx SEB/Indah Water, Prasarana, AXA Affin, Time Dotcom, BSN, Berjaya Sompo | frmPrintBordxIBM, frmPrintBordxSPNB, frmPrintBordxAXAAffin, frmPrintBordxTimeDotcom, frmPrintBordxBSN, frmPrintBordxBSI | 124 | `/claim-bordx/print/{seb,prasarana,axa-affin,time-dotcom,bsn,berjaya-sompo}` |
| BSN-PrintSummary (handler commented out → NO) | frmBSNPrintSummary | 124 | `/claim-bordx/print/bsn-summary` |
| Bordx Delivery > Send to Grp Company / Returned / Send to A/C | FrmSendtoPayor / FrmReturnPayor / FrmSendtoAC | 125 / 126 / 127 | `/claim-bordx/delivery/{send-to-payor,returned,send-to-ac}` |
| (NO) Send to MNRB | FrmSendtoMNRB | 125 | `/claim-bordx/delivery/send-to-mnrb` |
| Bor Acc Period | FrmACPeriod | 128 | `/claim-bordx/ac-period` |
| Monitoring > Bordx Control / Outstanding / Bordx & Worksheet | FrmBordxControl / FrmOSBordx / FrmBordxWks | 129 / 130 / 131 | `/claim-bordx/monitoring/{control,outstanding,bordx-worksheet}` |
| Claim Tracking / (Simple) / 2 / (Jupiter) / (Jupiter - Simple) / 2 (Jupiter) / (SUN System) | frmClaimRpt / frmClaimRptSimple / FrmClaimRpt2 / frmNewClaimRpt / frmNewClaimRptSimpleJup / FrmNewClaimRpt2 / frmClaimRptSUN | 132 / 132 / 120 / 163 / 163 / 164 / 132 | `/claim-bordx/claim-tracking/{standard,simple,v2,jupiter,jupiter-simple,v2-jupiter,sun}` |
| ClaimBordxDate | FrmBordxDate | 122 | `/claim-bordx/bordx-date` |
| LG > Generate LG (Crystal LG preview + Word cover letter) | frmLGGeneration | 119 | `/lg/generate` |
| Audit review / Audit Tracking | frmAuditInvoice / FrmAuditInvoiceTracking | 119 | `/audit/review`, `/audit/tracking` |
| Payment > To Clinic | frmPaymentClinic | 133 | `/payment/clinic` |
| Payment > To Clinic - Invoice (+ cheque and PV range print) | frmPaymentCLNInvoice, frmPrintChequeRange, frmPrintPVRange | 121 | `/payment/clinic-invoice` |
| Payment > To Member | frmPaymentMemberAuto08 | 134 | `/payment/member` |
| (NO) Payment to Member (old) | frmPaymentMember | 134 | `/payment/member-legacy` |
| Payment > To Member - Batch (H, NH) | – | – | nothing to build (documented) |
| Payment > Tracking / Tracking (Jupiter) | FrmPayment / frmPayment2 | 165 / 167 | `/payment/tracking[-jupiter]` |
| Payment > FFS Payment | FrmPaymentFFS | 121 | `/payment/ffs` |
| (NO) Pay MBM Post / Pay Clinic Post | frmPayMBMPost / frmPayClinicPost | 120 | `/payment/post/{member,clinic}` |
| Receipt > Fr Member / Fr Grp Company / Fr FFS | FrmReceiptMember / FrmReceiptPayor / FrmReceiptFFS | 135 / 136 / 121 | `/receipt/{member,payor,ffs}` |
| (NO) Receipt Fr MNRB | FrmReceiptMNRB | 136 | `/receipt/mnrb` |
| Upload Receipt / (Multiple Receipts) | frmReceiptUpload / frmReceiptUpMultiple | 136 | `/receipt/upload[-multiple]` |
| A/C Export > Invoice > Clinic / Clinic(GST) / Clinic(GST-2017) / Member | frmPostingInvoiceCln / frmPostingInvoiceClnGST / frmPostingInvClnGST2017 / frmPostingInvoiceMember | 120 | `/ac-export/invoice/{clinic,clinic-gst,clinic-gst-2017,member}` |
| (NO) Clinic GST-2017 Zeroised, Invoice Posting | frmPostingInvClnGST2017Zeroised, frmInvPosting | 120 | `/ac-export/invoice/{clinic-gst-2017-zeroised,posting}` |
| A/C Export > Payment > Clinic / Member | FrmPostingPay2Cln / FrmPostingPay2MBM | 120 | `/ac-export/payment/{clinic,member}` |
| A/C Export > Bordx > Grp Company (H) / (GST) | frmPostingBordxtoGrpComp / frmPostingBordxtoGrpGST | 120 | `/ac-export/bordx/{grp-company,grp-company-gst}` |
| A/C Export > Receipt > Grp Company | frmPostingRecGrpCompany | 120 | `/ac-export/receipt/grp-company` |
| SUN System GST / Report / Allocation (+ New (NO), account search) | frmSUNGSTRpt / frmSUNSystem / frmSUNSystemAllocation, frmSUNSystemAllocation_New, frmSUNSystemSearchAcc | 120 | `/ac-export/sun/{gst,report,allocation,allocation-new}` |
| Tracking > Clinic / Grp Com Clinic | FrmTracking / FrmTracking2 | 144 | `/tracking/{clinic,grp-company-clinic}` |
| Maintenance > Entity (+ clinic search; 4 list reports) | frmMaintenanceEntity, frmSearchCLN | 137 | `/maintenance/entity` |
| Maintenance > Plan (8 list reports) | frmMaintenancePlan | 138 | `/maintenance/plan` |
| Maintenance > Case (5 list reports) | frmMaintenanceOthers | 139 | `/maintenance/case` |
| Maintenance > Calculation (2 list reports) | frmMaintenanceCal | 140 | `/maintenance/calculation` |
| Maintenance > LG Maintenance | frmLGMaintenance | 138 | `/maintenance/lg` |
| Maintenance > Benefit > Member / Clinic / Dental | frmMaintenanceBenefits / frmMaintenanceBenefitsCln / frmMaintenanceBNFDental | 141 / 142 / 141 | `/maintenance/benefit/{member,clinic,dental}` |
| (NO) Benefit Clinic (new) | frmMaintenanceBenefitsClnNew | 142 | `/maintenance/benefit/clinic-new` |
| Maintenance > Check Item (1 list report) | frmCheckingItems | 143 | `/maintenance/check-item` |
| Maintenance > DataTransfer (NH) | – | – | nothing to build (documented) |
| Maintenance > Test Form | frmTest | none (desktop has no check) | `/maintenance/test-form`, gated by a config flag |
| 26 DataReports (Rpt*List etc.) | *.Dsr | as parent form | "Print list" PDF on each maintenance tab |

### 6.6 Outside the project (needs your confirmation; not built by default)
In the folder but not in `medixcis.vbp`:
- Dental: frmDentalInvoice, frmDentalWorksheet, frmInvDentalCln, frmInvDentalAdj, frmInvDentalProcessing (these would back the hidden Dental menu items).
- Old invoice-processing versions: frmInvoiceProcessing, frmInvProcessing, frmInvoiceOP, frmInvMedicine.
- Old payment batches: frmPayClinicBatch, frmPayMbmBatch, frmPaymentClnByBatch/ByInvoice, frmPaymentHospital.
- Other: frmSUNSystemUnallocate, frmMCOGSTInvoice, the posting Zeroised variants, frmUpdateInvoice, frmWorksheetAdj, frmWorksheetsAppOP, backups, and about 15 junk-named files.

## 7. Phases (order follows master data → transactions → finance → reporting)

| # | Phase | Contents | Exit criteria |
|---|---|---|---|
| 0 ✅ done 31 Oct 2026 | Foundation | Scaffold `Cis.sln` (copy/adapt GENESIS Auth/Common/Data), `sql/000_setup.sql` (cis schema + synonyms), login/logout/change-password/session tracking/action log, Angular workspace + ported CSS + shared components + AppLayout with the full CIS menu (unbuilt items shown disabled, "not migrated yet", as GENESIS does), lookups API, report engine (Excel/PDF/CSV writers, template store), file-drop, email sender, test projects, e2e harness, docs skeleton + `menu-mapping.md` | DEVADMIN can log in and sees the access-filtered menu; CIScnt and Tbl_USRTrack match VB6; architecture tests pass |
| 1 ✅ done 1 Nov 2026 | Maintenance | Entity, Plan, Case, Calculation, LG Maintenance, Benefit Member/Clinic/Clinic-new/Dental, Check Item, the 26 list reports, clinic search | CRUD parity on every tab; list PDFs. Result: all 11 screens migrated (code-table engine + per-screen services, `cis.*` procedures 100–220); 154 API tests and 33 browser checks pass; desktop defects fixed or kept per the bug policy and listed in `Cis/docs/screens/*.md` |
| 2 ✅ done 2 Nov 2026 | Membership | MBM Adjustment, Enquiry (+Account), Topup, General Exclusions + Enquiry, member search popups, clinic member registration, Case Transfer | Result: all migrated except the General Exclusions editing form, which is not in the VB6 source (the Enquiry is migrated; documented). `cis.*` procedures 300–365 (member search, Topup, Case Transfer, registration, exclusions, Adjustment, Enquiry incl. FlexCare copies without the production OPENROWSET); 190 API tests and 43 browser checks pass. Saves that write MBM / MBMTwo / MBMOthersTwo need the HISDB master key for the encryption triggers (missing on the restored local copy). Details and deliberate differences in `Cis/docs/screens/membership.md`, test data in `Cis/docs/testing/membership-test-guide.md` |
| 3 ✅ done 2 Nov 2026 | Invoice core | Worksheet calculation engine (frmWorksheet, shared by everything below), Clinic and Member invoice, the 3 Adjustment screens, Enquiry, Batch Transfer, OP Batch, Invoice Update, Worksheet Reg/Adj, ICD/case-list/pending/member-info popups, MMA/CPT, Pass to Claim, Update Claims Data | Same input gives identical CISWorksheet/CISDB rows as VB6 (SQL diff). Result: all migrated except Worksheet Registration / Adjustment (frmNewWSReg / frmNewWSAjd), which the desktop cannot open (menu handlers commented out) and whose views / tables (`VIEW_MBM_CAS_Claim`, `View_Invoice`, `CAS`, the old `CLMLimit` … tables) no longer exist — their work is done by the worksheet engine, the invoices and the Adjustment screens (documented). `cis.*` procedures 400–530; the desktop's save procedures are called with parameters (its pasted SQL broke for amounts ≥ 1,000); saves verified with rolled-back database tests on real cases; 229 API tests and 55 browser checks pass. The deferment letter (Word template not in the source) is not generated. Details in `Cis/docs/screens/invoice.md`, test data in `Cis/docs/testing/invoice-test-guide.md` |
| 4 ✅ done 8 Nov 2026 | LG and Audit | Generate LG (PDF letter + cover letter), Audit review, Audit Tracking | Result: all 3 migrated. Audit review (frmAuditInvoice) with its guarantee letter documents; Audit Tracking (FrmAuditInvoiceTracking) with every filter as a parameter and the Excel; Generate LG (frmLGGeneration): limits (CheckMBMLimit as pure rules), the Add / Edit limit trees translated statement by statement, Issue (SaveLG in one transaction), Decline, Preview, Copy File, company notification e-mail through the configured sender (the form's hard-coded SMTP account is not used). The 13 Word templates are not available, so the GL / decline / preview letters are PDFs with the same fields. `cis.*` procedures 600–630; 241 API tests and 58 browser checks pass. Details and reproduced desktop issues in `Cis/docs/screens/lg-audit.md`, test data in `Cis/docs/testing/lg-audit-test-guide.md` |
| 5 ✅ done 8 Nov 2026 | MCO Fees | Invoice, GST, GST Zeroised, Fee Listing, Refund Listing (Crystal → PDF, Excel invoices) | Result: all 5 migrated. MCO Invoice, GST and GST Zeroised (member lists by bordereaux batch / range / invoice, numbering from CLNMCOINVQ and member stamps in one transaction, Excel invoices with summary and detail sheets; GST Zeroised prints only the checked members), Fee Listing (fee listing and MCO invoice as PDFs instead of Crystal and the `\\PLUTO` staging table), Refund Listing (refunds with the 50 / 25 utilised-fee rule, Excel listing, credit note from CLNMCOCRNoteSeq). The `MCOFeeNew` / `MCOFeeModule` / GST procedures' pasted criteria are replaced by parameters; the `.rpt` / `.xlt` templates are not available, so layouts are generated with the same fields. `cis.*` procedures 700–740; 249 API tests and 63 browser checks pass. Details and reproduced desktop issues in `Cis/docs/screens/mco-fees.md`, test data in `Cis/docs/testing/mco-fees-test-guide.md` |
| 6 ✅ done 9 Nov 2026 | Claim Bordx | Insert Worksheet/Invoice, Remove, Bordx Date, 9 Print Bordx variants + BSN summary (Crystal summary/detail → PDF; Excel pay advice), Delivery (payor/returned/A/C/MNRB), A/C Period, Monitoring ×3 | Bordx numbers and totals match VB6. Result: 19 of 21 migrated — Insert Worksheet (ClaimBordxModule_isPureASO case 21 and the UOB cases copied with parameters; bordx number, case updates and Bordx row in one transaction), Remove Worksheet, ClaimBordxDate, the four deliveries and Bor Acc Period on one page, the three monitoring lists, Print Bordx and its seven client variants (generated pay-advice workbooks; the ~20 templates and the Word cover letter are not available). Not migrated: Insert Invoice (menu item disabled) and BSN-PrintSummary (never opened, unfinished). Fixed instead of reproduced: unticked worksheets being removed, Bordx Control with a subsidiary always failing, the blank A/C period failing. SQL Server memory pinned (min 2 GB / max 4 GB) and a per-phase browser runner (`e2e/only.mjs`) to avoid the timeouts; `cis.*` procedures 800–860; 257 API tests pass. Details in `Cis/docs/screens/claim-bordx.md`, test data in `Cis/docs/testing/claim-bordx-test-guide.md` |
| 7 ✅ done 9 Nov 2026 | Payment | To Clinic, Clinic-Invoice (+cheque/PV range print, invoice detail), To Member (Auto08 + legacy), FFS, Pay posts, Payment Tracking ×2 | PV/cheque output; amount in words. Result: To Clinic, To Clinic - Invoice (with the invoice popup, PV range and cheque range), To Member (Auto08, with member PV / cheque ranges), Tracking and Tracking (Jupiter), FFS migrated; the list procedures copied with parameters and checked against the legacy ones (same rows); PV numbers, payment rows and cheque details in one transaction with the PVSeq2013 row locked; CurrencyConversion.bas ported and checked against stored wording; PV / cheque templates not available → generated workbooks. Not migrated: To Member (old) and the two posting forms (never opened), To Member - Batch (no handler). `cis.*` procedures 900–940; 272 API tests pass. Details in `Cis/docs/screens/payment.md`, test data in `Cis/docs/testing/payment-test-guide.md` |
| 8 ✅ done 9 Nov 2026 | Receipt | Member, Payor, FFS, MNRB, Upload, Upload Multiple (ExcelDataReader) | Result: Fr Member, Fr Grp Company (bordereaux / worksheet), Fr FFS, Upload Receipt and Upload Receipt (Multiple Receipts) migrated; Fr MNRB migrated on request (never opened by the desktop; no MNRB data locally, checked for its messages). Lists copied from PaymentReceiptModule with parameters and checked against the legacy procedure; slow lists sped up with filters implied by their joins (worksheet receipts 27 s → 6 s, FFS receipt relist 25 s → 7 s); each save in one transaction; uploads read `.xlsx` on the server with the existing ClosedXML reader (no ExcelDataReader needed). The restored data has no member debit notes, so Fr Member is checked only for its messages. `cis.*` procedures 950–990; 279 API tests pass. Details in `Cis/docs/screens/receipt.md`, test data in `Cis/docs/testing/receipt-test-guide.md` |
| 9 | A/C Export | Invoice posting ×6, Payment posting ×2, Bordx posting ×2, Receipt posting, SUN GST/Report/Allocation/Allocation-New (+account search) | Posting batch numbers and exports match |
| 10 | Tracking and reports | MBM Tracking ×5 (210 report types → definitions), Claim Tracking ×7, Invoice Tracking ×2, Clinic and Grp-Com-Clinic tracking, Claim ASCII ×7 + OMNI providers + Claims data + Sin Chew, Email notifications ×3 | Every report type has a definition and a test; exports match VB6 for sample filters |
| 11 | Finish | Test Form, Print screen, open-screens list, performance pass (indexes, row caps, background jobs), security review, publish script (API serving the Angular build), onboarding docs, AGENTS.md | Full traceability: every form, button and report ticked in `menu-mapping.md` |

Each phase ends with: `dotnet build` + `dotnet test`, `ng lint/test/build`, e2e for the phase, and a parity sign-off against VB6 LocalRun.

## 8. Verification

- [ ] `dotnet build Cis\Cis.sln` and `dotnet test` (unit, service, architecture, SqlScript; DB tests against local `(local)` with `CIS_TEST_*` env connection strings, rolled back).
- [ ] `sqlcmd -S localhost -E -d CISDB -i Cis\sql\cis-main.sql` applies cleanly and can be rerun (CREATE OR ALTER, additive only).
- [ ] `npm run lint`, `npm test`, `npm run build` in `CisUI`; e2e `node e2e/run.mjs` against a running API.
- [ ] Parity: the per-screen checklist in `docs/screens/*.md` is fully ticked; SQL diff of rows written by VB6 vs web for scripted scenarios; exported Excel/PDF/ASCII compared with the VB6 output.
- [ ] Visual consistency: the same CSS variables and class names as GenisisUI; side-by-side screenshot review against GENESIS.
- [ ] Security: every endpoint has `[RequireAccess]`, no string-built SQL, and the JWT key comes from env (`Auth__JwtKey`).

## 9. Risks / considerations

- **Size**: about 396k VB lines. The tracking families are mostly repeated query builders; turning them into report definitions is what makes the work tractable. They are the biggest and riskiest part (Phase 10).
- **Hidden behaviour** in `On Error Resume Next` blocks (e.g. login, logout) and in SQL built with `strAppend`. The inventory step lists each one; behaviour is reproduced deliberately, never silently "fixed". Fixes are flagged in the screen doc.
- **External dependencies missing locally**: SUN Systems DB, Neptune/Jupiter replica, HISREPORTS templates, SMTP server, server folders (Outpatient, ClinicPVS, CLNPolicyWording, Bordx, MCO). All are configurable; screens degrade with clear messages.
- **Shared production databases**: the `cis` schema, procs and synonyms are additive. Nothing existing is altered, so VB6 CIS, VB6 HIS and GENESIS keep working side by side.
- **Concurrency**: number generators move to locked procs; the desktop could issue duplicate numbers under load.
- **Local machine** is short on RAM (~2 GB free). Running VB6 + API + Angular dev server + SQL together will be tight; DB tests may need a rerun after timeouts (same as GENESIS).
- **Angular vs GENESIS React**: components are re-implemented in Angular with the same markup/classes/CSS rather than shared. If GENESIS CSS changes later, `styles.css` must be resynced (documented).

## 10. Open questions: deferred (decision 31 Oct 2026)

None of these blocks Phases 0–4. The default applies until the user decides, and each is raised again before the phase that needs it.

| # | Question | Decide before | Default until decided |
|---|---|---|---|
| 1 | SUN Systems DB: connection details / a test copy? | Phase 9 (SUN screens), Phase 10 (Claim Tracking SUN) | `SunSystems` connection setting; screens return "SUN Systems database not configured" |
| 2 | Jupiter/Neptune: production reporting replica, or main DBs? | Phase 10 (Jupiter variants) | `Reporting` connection = main databases |
| 3 | SMTP server and sender for the email-notification screens | Phase 10 (email screens) | MailKit sender, tested with a local pickup folder |
| 4 | Can the `\\svr-doc\HISREPORTS` templates and `.rpt` files be supplied? | Phase 5 (first formatted outputs) | Layouts rebuilt in code (ClosedXML/MigraDoc); `ReportTemplateStore` picks up templates later without rework |
| 5 | Build the §6.6 out-of-project dental and old screens? | After Phase 11 | Not built |
