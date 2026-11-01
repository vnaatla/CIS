# Maintenance — developer test guide (Phase 1)

Manual test script for every Maintenance screen and tab of the web app, with real data from the local databases (values checked on
1 Nov 2026). Each tab lists: where it is, data that already exists (Search), data to add (Add / Save), the checks and their exact desktop
messages, Delete, and what is deliberately different from VB6. Desktop texts are quoted exactly, including odd spelling and spaces.

## 0. Before you start

| Step | Command / value |
|---|---|
| SQL objects | `sqlcmd -S localhost -E -d CISDB -i Cis\sql\cis-main.sql` (after any SQL change) |
| API | `cd Cis\src\Cis.Api; $env:Auth__JwtKey='<32+ chars>'; dotnet run --launch-profile http` → http://localhost:5100/api/health = `Healthy` |
| UI | `cd CisUI; npm start` → http://localhost:4200 |
| Login | `DEVADMIN` / `Gen12345` (wrong password → login refused with the desktop message) |
| Automated | API: `dotnet test Cis\Cis.sln` (154 tests, DB tests roll back) · UI: `npm run lint; npm test` · browser: `$env:E2E_PASSWORD='Gen12345'; npm run e2e` (35 checks) |

Rules for test data:

- Saves write to the real local tables (CISDB, and HISDB for Clinic Info / Clinic Access). Use the **`ZZ` / `Z` test values** below and
  delete them afterwards with the tab's Delete button, or with the clean-up SQL in section 9.
- If a request fails with a 504 / timeout, SQL Server is usually paged out (low RAM): wait a minute and retry before reporting a bug.
- **Add never calls the API**: it clears and unlocks the boxes. The API is called by **Save → "Yes"** in the confirmation.

## 1. Checks for every tab

| # | Action | Expected |
|---|---|---|
| G1 | Open the tab | Boxes empty, Search / Add / Clear / Refresh / Print enabled; Edit / Delete / Save disabled |
| G2 | Search with every box empty | Every row of the table (counts below); TOTAL RECORDS = that count |
| G3 | Search with a value that does not exist (e.g. `ZZZ`) | Grid **empties** and the "Record not found" message (exact text per tab) |
| G4 | Click a row | Its values fill the boxes, read-only; Edit / Delete enabled (where the desktop has them) |
| G5 | Add → Save with every box empty | The first required-field message of the tab (lists below give the order) |
| G6 | Add a new record → Save → **No** | Nothing saved, no request in DevTools › Network |
| G7 | Add the same key twice | `Record Duplicated! Please Change The Corresponding Field` (unless the tab says otherwise) |
| G8 | Delete → **No** | Nothing deleted |
| G9 | Clear | Boxes **and the grid** empty, TOTAL RECORDS 0 |
| G10 | Refresh | Lists / drop-downs reloaded, boxes empty |
| G11 | Print | Opens the listing PDF (Benefit Listing / Clinic Info export → Excel) |
| G12 | Paging: search a list longer than 50 rows | Pager under the grid: `Rows 1–50 of N`, First / Previous / page numbers / Next / Last, Rows per page 25/50/100/200; sorting a column sorts **all** rows; a new search returns to page 1 |
| G13 | Server paging (Clinic Group, Doctor, Clinic Info, Group Company Clinic, Plan Code, Annual Limit, MCO Fee) | DevTools › Network shows one request per page (`page=1&pageSize=50…`), each returning ≤ 50 rows and `total` |
| G14 | No access right | A user without the screen's USRAccess position does not see the menu item and gets 403 from the API |

## 2. Check Item — `/maintenance/check-item` (frmCheckingItems, right 143, table CheckItems)

| Item | Value |
|---|---|
| Existing data | 10 rows, e.g. `BILL` / `BILL ACCUTE` / `30` |
| Search | Category Code (first 3 characters =); empty → 10 rows |
| Add (test) | Category Code `ZZ1`, Category Name `ZZ TEST ITEM`, Amount `12.50` → `Record Saved...` |
| Checks in order | `Please Enter Category Code` · `Category Code must be 3 integer!` (code shorter than 3) · `Please Enter Category Description` · `Please Enter Amount` |
| Amount box | digits and one decimal point only |
| Edit | row `ZZ1` → Edit → Amount `15` → Save → `Record Updated...` |
| Delete | row `ZZ1` → Delete → Yes → `Record Deleted... ` |

## 3. Calculation — `/maintenance/calculation` (frmMaintenanceCal, right 140)

| Tab | Table | Existing | Add (test) | Checks in order |
|---|---|---|---|---|
| Co-Payment Cal. | CoPayment | 4 rows, `1` / `Lower of Act of % or Limit` | Code `ZZ1`, Description `ZZ TEST` | `Please Enter Code` · `Co-Payment Code must be 3 integer!` · `Please Enter Description` |
| Fee Payable 2 Clinic | FEE | 4 rows, `001` / `FLAT RATE` | Fee Code `ZZ1`, Description `ZZ TEST` | `Please Enter Fee Code` · `Fee Code must be 3 integer!` · `Please Enter Description` |

Remarks are optional. Edit / Delete as G4 / Check Item.

## 4. Case — `/maintenance/case` (frmMaintenanceOthers, right 139)

| Tab | Table | Existing (empty search) | Add (test) | Checks in order |
|---|---|---|---|---|
| Treatment Reason | TreatmentReason | 612 rows (pages!), `101` / `URTI/SORE THROAT` / `ACUTE` | Reason Code `ZZ1`, Description `ZZ TEST` | `Please Enter Reason Code` · `Reason Code must be 3 integer!` · `Please Enter Description` |
| Procedures | Procedures | 16 rows, `P001` / `X-RAY` | Code `ZZ01`, Description `ZZ TEST` | `Please Enter Procedure Code` · `Please Enter Description` |
| Drugs | Drug | 3 rows, `001` / `DRUG UNIT` / `1` | Code `ZZ1`, Unit `TAB`, Value `1234.5`, Description optional | `Please Enter Drug Code` · `Please Enter Drug Unit` · `Please Enter Drug Value` — the value is stored as text `1,234.50` (desktop `Format(…, "#,##0.00")`) |
| Diagnosis Not Payable | DiagNotPayable | 3 rows, `02` / `002` | Plan Code `ZZ1`, Diagnosis Code `ZZ1` | `Please Enter Plan Code` · `Please Enter Diagnosis Code` |
| Repudiation | Repudiation | 19 rows, `R001` / `EXPENSES IN RESPECT OF PREGNANCY…` | Code `ZZ01`, Description `ZZ TEST` | `Please Enter Repudiation Code` · `Please Enter Description` |

Treatment Reason with 612 rows is the quickest way to test the browser pager (G12).

## 5. LG Maintenance — `/maintenance/lg` (frmLGMaintenance, rights 139–142 by tab)

### 5.1 GL Template Binding (Tbl_OPGLMaster)

| Item | Value |
|---|---|
| Drop-downs | Group company (826), GL Type (14), Template (34), LG Format (4), VIP (2). Empty after a SQL stall → **Reload lists** |
| Search | Group company `B006 - BANK SIMPANAN NASIONAL` → 617 bindings (pages); without a group company → `Plese Select Group Company to Search` and the grid empties |
| Cascade | Picking a group company fills its Plan Codes and Pay Codes |
| Add (test) | Group company B006, a plan, GL Type, Template, LG Format → Save → `New Record Inserted Successfully..` |
| Edit | row → Edit → change template → `Record Updated Successfully..`; Edit without a row → `Please search and select the record to edit.` |
| Group name | stored as the part before the second "-" (`1001 - 1MILLENNIUM INTERNATIONAL SDN BHD-1` → `1MILLENNIUM INTERNATIONAL SDN BHD`) |
| Delete | none (as the desktop) |
| Different | Clear also empties the grid (desktop kept it) |

### 5.2 GL Type Binding / Template Name Binding / Decline Reason (code tables, no Delete)

| Tab | Table | Existing | Add (test) |
|---|---|---|---|
| GL Type Binding | Tbl_GLType | 14, `1` / `Normal GL` / `Standared GL` | GL TYPE `ZZ TEST`, description `ZZ` (ID is generated) |
| Template Name Binding | Tbl_GLTemplates | 34, `1003` / `HCBARAKAHPBJVGROUP.dot` | Name `ZZTEST.dot`, description `ZZ` |
| Decline Reason | Tbl_DeclineReason | 18, `1` / `Case decline due to limit exhausted…` | Reason `ZZ TEST`, remarks `ZZ` |

Saved → `New Record Inserted Successfully..`; edited → `Record Updated Successfully..`; search = text contains. No Delete button: remove
test rows with SQL (section 9).

## 6. Benefit — `/maintenance/benefit/...`

### 6.1 Member — Category (BNFCategoryNew, right 141)

31 rows (`101` / `GP FEE & Basic Med`). Add Category Code `ZZ1`, Description `ZZ TEST`. Checks: `Please Enter Code` ·
`Category Code must be 3 integer!` · `Please Enter Description`.

### 6.2 Member — Benefit Entry (BNF)

| Item | Value |
|---|---|
| Search (existing) | Health `N`, Insured `H`, Plan `A001`, Eff Date `01/01/2004` → 29 saved rows fill the grid |
| Checks (Search / Save) | `Please key in Health Type !` · `Please key in Insured Type !` · `Please key in Plan Code!` · `Please key in Effective Date!` / `Please key in VALID Effective Date!` |
| Add (test) | Health `N`, Insured `H`, Plan `A001`, Eff Date `01/11/2026`; grid row 101: Calculation 1, RM Per Visit `10` → Save → Yes → `Record Saved...` |
| Rows written | Calculation 1/2 only with RM Per Visit or Percentage; 3 always; 4 only with RM Per Visit |
| Duplicate | same Insured + Plan + Eff Date again → `Record Duplicated! Please Change The Corresponding Field` (ignores Health Type, as the desktop) |
| Visit counts | whole numbers (`No of Visit Per day must be a whole number.`) |
| No Edit / Delete | as the desktop (remove test rows with SQL) |

### 6.3 Member — Benefit Listing

Plan Code `A001` → Search → its rows (13 columns); a row copies into the filters; Print → Excel. Not found → `Record not found!`.

### 6.4 Clinic and Clinic (new) (STDClinicBNFSpe, right 142)

| Item | Value |
|---|---|
| Search (existing) | Card Type `Gold`, Eff Date `01/01/2002` → categories 101–103 show 50 / 60 / 60 |
| Checks | `Please key in BNF Card Type!` · `Please Enter (Valid )effectice Date!` |
| Add (test) | Card Type `Silver`, Eff Date `01/12/2026`, fee `10` on row 101 → `Record Saved...`; same type + date again → `Record Duplicated! …` |
| Different | Clear empties all 31 rows (desktop only 29); Print works in "Clinic (new)" too |

### 6.5 Dental — `/maintenance/benefit/dental` (right 141)

| Tab | Test |
|---|---|
| Tooth Category (ToothRefCategory) | 4 rows (`0` / `ADULT DENTITION UPPER TEETH`); add Code `Z`, Description `ZZ TEST`; checks `Please Enter Category Code` · `Please Enter Description` |
| Tooth Reference (ToothRef) | 52 rows (`11` / category `0` / `RIGHT UPPER CENTRAL INCISOR`); add Code `ZZ1`, Category `0`, Description `ZZ TEST`; checks `Please Enter Code` · `Please Enter Category` · `Please Enter Description` |
| Benefit Category (DentalBNFCategory) | 12 rows (`100` / `FILLINGS`); add `ZZ1` / `ZZ TEST` |
| Benefit Entry (DentalBNF) | Search Plan `01`, Insured `I`, Health `S`, Eff Date `24/09/2003` → 3 saved rows; same rules as 6.2; Print → "DENTAL BENEFIT ENTRY LISTING" |

## 7. Entity — `/maintenance/entity` (frmMaintenanceEntity)

### 7.1 Clinic Group (GRPClinic) — server paging

| Item | Value |
|---|---|
| Search | Cln Grp. Name `KLINIK` → matching groups; empty → 2,521 groups, 50 per page |
| Add (test) | Initial Group Name `ZZ TEST`, Cln Grp. Name `ZZ TEST GROUP`, Address 1 `JALAN TEST 1`, Eff. Date `01/11/2026` → `Record Saved...`, code `Z####` generated |
| Checks in order | `Please Enter Clinic Group Name` (Initial Group Name, Add only) · `Please Enter Clinic Group Address` · `Please Enter Effective Date` |
| Edit | row → Edit (Initial Group Name locked, not required) → Save → `Record Updated...` |
| Delete | refused while a HISDB clinic uses the group: `This record cannot be deleted because it is used in Clinic Table!`; test group → `Record Deleted...` |

### 7.2 Doctor — server paging

Add Doctor Full Name `ZZ TEST DOCTOR` → code `Z00001`-style; duplicate name refused; check `Please Enter Doctor Full Name`. Delete works
(desktop's never did).

### 7.3 Clinic Info (HISDB CLN, 8,304 clinics) — server paging

| Item | Value |
|---|---|
| Search | Clinic Name `KLINIK` → clinics (38 columns, pages); View/Back = larger list; Print → Excel of the whole result |
| Add (test) | Clinic Name `ZZ TEST CLINIC`, Initial `ZZ`, Type `G`, Address `JALAN ZZ TEST 1`, Benefit Type `1 - Standard`, Received Date `01/11/2026`, State + City, Post Code `50000`, Panel `Y`, 24 Hours `N`, FFS Status `N`, Location `KL`, Operating Hour `9-5` → `Record Saved...` |
| Checks in order | `Please Enter Clinic Name` · `Please Enter Clinic Initial` · `Please Enter Clinic Type` · `Please Enter Clinic Address` · ` Please Enter Benefit Type` · `Please Enter Received Date!` · ` Please Enter State` · ` Please Enter City` · ` Please Enter PostCode` · ` Please Enter Panel` · … (leading spaces are the desktop's) |
| Duplicate | same address on Add → `Duplicate Record!` |
| FFS | FFS Status `Y` without a percentage → ` Please Enter FFS Percentage` |
| Delete | clinic with cases → `There are case(s) register under this Clinic !`; test clinic → deleted |

### 7.4 Group Company Clinic (VIEW_GrpCom_Clinic, ~1.7 million rows) — paged in SQL

Search Group company `B006` → its clinics; empty search → `Rows 1–50 of 1,744,655` with every page reachable (no 1,000-row cap any
more). Add needs "Load the group company and clinic lists" ticked; an existing pair → `Record Duplicated! Please Change The Corresponding Field`
(also on Edit, as the desktop). Delete / Print as the desktop.

### 7.5 Clinic Access (HISDB CLNAccessControl)

Search / Refresh → latest 100 rows (e.g. clinic `A019`, 17/08/2020). Add → "..." → search `KLINIK` → pick a clinic → From `01/11/2026`,
To `31/12/2026` → Save → saved row first in the list. Checks: `Please Enter Clinic Code` · `Invalid Effective Date From` ·
`InvalidEffective Date To` · `Effective Date From cannot be more than Effective Date To!` · `Clinic Code does not exists, please enter again!`;
dates after 06/06/2079 refused. No Edit / Delete / Print (as the desktop).

## 8. Plan — `/maintenance/plan` (frmMaintenancePlan, right 138)

| Tab | Search (existing) | Add (test) | Checks / expected |
|---|---|---|---|
| HealthType (HLT) | empty → `C`, `N`, `S` | `Z` / `ZZ TEST` → stored upper-case | `Please Enter Health Type` · `Please Enter Health Description`; Delete `N` → `This record cannot be deleted because it is used in Plan table!`; Delete empty → `Please Enter Health Code` |
| Insured Type (INS) | empty → `F`, `G`, `H`, `I` | `Z` / `ZZ TEST` | `Please Enter Insured Type` · `Please Enter Insured Type Description`; Delete `H` → refused (`…used in Benefit table!` / `…PRM table!`) |
| Plan Code (PLN, server paging) | Plan Code box 1 `A001` → plan A001 (group A001 ASTANA INTERNATIONAL); empty → `Rows 1–50 of 2,690` | HLT `N`, CIS Pay `DC` (DC Status becomes Y), HIS Pay any, Grp Comp. a `CODE - NAME`, Cln Cont. `GOLD`, all Yes/No boxes, No of Plan Code `2`, descriptions `ZZ PLAN 1` / `ZZ PLAN 2`, OP Dependent `N` → `Record Saved...` and the group's plans listed with 2 new codes | 14 checks from `Please Enter Health Code ` to `Please Select OP Dependent`; HLT `S` → refused (deliberate); no Edit (as the desktop); Delete of a plan used in membership → `This record cannot be deleted because it is used in Membership table!` |
| Premium (PRM) | Insured `H` → rows such as H / N / WB / AA05 / 01/01/1999 | Insured `H`, Health `N`, Payor `DC`, Plan `ZZ9`, Eff Date `01/01/1901`, Version `1`, band 01 premium `100` → `Record Saved...` (7 rows) | `Please Enter Insured Code` · `Please Enter Health Code` · `Please Enter Plan Code` · `Please validate your Premium Effective date`; B / C enable the supplementary column; no Edit |
| Annual Limit (AnnualLimit, server paging) | Plan `A001` → G / N / 01/01/2004 version 1 | Plan `ZZ9`, Health `N`, Insured `H`, limits `1000` / `200` / `150` / `50`, Eff Date `01/01/1901`, Version `1` → `Record Saved...`; Edit → `Record Updated...` | 13 checks (`Please Enter Plan Code` … `Please Enter Annual Limit Version`; B / C add the supplementary ones); same keys again → `Record Duplicated!` / `Please Change The Corresponding Field` |
| Visit Limit (VisitLimit) | empty → 87 rows, e.g. plan `01` G / N 01/01/2008 v1 | Plan `ZZ9`, Health `N`, Insured `H`, Visit Limit `12`, Eff Date `01/01/1901`, Version `1` → `Record Saved...` | `Please Enter Visit Limit` · `Please Enter Visit Limit Effective Date` · `Please Enter Visit Limit Version` · B / C → `Please Enter Supp Visit Limit`; Edit works |
| MCO Fee (MCO, server paging) | Plan `A001` → its fees (G / N 01/01/2004) | Health `N`, Plan `ZZ9`, Eff Date `01/01/1901`, I Adult `10` → `Record Saved ! ` | `Please Enter Health Code` · `Please Enter Plan Code` · `Please Enter MCO Fees ` (all principal fees empty) · `Please validate your MCO Effective date`; same again → `Record Not Save for Insured Type 'I! Please Change The Corresponding Field`; Delete removes the row (deliberate) |
| Co-Payment (MBM) (COPaymentMBM) | empty → 76 rows, e.g. N / A067 / 20 | Health `N`, Plan `ZZ9`, Amount `5` → `Record Saved...` | `Please Enter Health Type` · `Please Enter Plan Code ` · `Please Enter Amount `; plan again → duplicate refused; Delete empty → `Please Enter Plan Code` |

## 9. Clean-up SQL (run in PowerShell; deletes only the test values above)

```sql
-- CISDB
DELETE dbo.CheckItems WHERE CatCode = 'ZZ1';
DELETE dbo.CoPayment WHERE COPCode = 'ZZ1';      DELETE dbo.FEE WHERE FeeCode = 'ZZ1';
DELETE dbo.TreatmentReason WHERE ReasonCode = 'ZZ1'; DELETE dbo.Procedures WHERE ProcedureCode = 'ZZ01';
DELETE dbo.Drug WHERE DrugCode = 'ZZ1';          DELETE dbo.DiagNotPayable WHERE PLNCode = 'ZZ1';
DELETE dbo.Repudiation WHERE RepCode = 'ZZ01';   DELETE dbo.BNFCategoryNew WHERE BNFCATCode = 'ZZ1';
DELETE dbo.ToothRefCategory WHERE ToothRefCatCode = 'Z'; DELETE dbo.ToothRef WHERE ToothRefCode = 'ZZ1';
DELETE dbo.DentalBNFCategory WHERE DentalBNFCATCode = 'ZZ1';
DELETE dbo.Tbl_GLType WHERE GLType = 'ZZ TEST';  DELETE dbo.Tbl_GLTemplates WHERE GLTemplateName = 'ZZTEST.dot';
DELETE dbo.Tbl_DeclineReason WHERE DeclineReason = 'ZZ TEST';
DELETE dbo.HLT WHERE HLTCode = 'Z';              DELETE dbo.INS WHERE INSCode = 'Z';
DELETE dbo.PRM WHERE PLNCode = 'ZZ9';            DELETE dbo.AnnualLimit WHERE PLNCode = 'ZZ9';
DELETE dbo.VisitLimit WHERE PLNCode = 'ZZ9';     DELETE dbo.MCO WHERE PLNCode = 'ZZ9';
DELETE dbo.COPaymentMBM WHERE PLNCode = 'ZZ9';
DELETE dbo.PLN WHERE PLNDesc IN ('ZZ PLAN 1', 'ZZ PLAN 2');
DELETE dbo.BNF WHERE PLNCode = 'A001' AND BNFEffDate = '2026-11-01';
DELETE dbo.STDClinicBNFSpe WHERE CLNBNFType = 'Silver' AND CLNBNFEffDate = '2026-12-01';
DELETE dbo.GRPClinic WHERE GRPClnName = 'ZZ TEST GROUP';
-- HISDB
DELETE HISDB.dbo.CLN WHERE CLNName = 'ZZ TEST CLINIC';
```

Check the row counts each statement reports before committing; the code counters (GRPClnSeq, CLNSeq, PLNSeq) are not rolled back.

## 10. Reporting a defect

Give: screen › tab, the exact steps and data, the expected text from this guide, the actual text, and the failing request from DevTools ›
Network (URL, status, response body). Compare with the desktop (`LocalRun`, see `AGENTS.md`) when the behaviour is unclear; deliberate
differences are listed in `Cis/docs/screens/*.md`.
