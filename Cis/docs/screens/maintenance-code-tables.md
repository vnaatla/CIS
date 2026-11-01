# Maintenance code tables → `/maintenance/check-item`, `/maintenance/calculation`, `/maintenance/case`

Phase 1. Source: `frmCheckingItems`, `frmMaintenanceCal`, `frmMaintenanceOthers` (+ the category tabs of `frmMaintenanceBNFDental` /
`frmMaintenanceBenefits`, which use the same engine and go live with those screens). API: `api/maintenance/code-tables/{key}`,
`api/maintenance/list-reports/{key}`. Definitions: `Cis.Application/Maintenance/CodeTables/CodeTableCatalog.cs`.

## Behaviour checklist (all tabs)

| Desktop | Web | ✓ |
|---|---|---|
| Search: strAppend on the searchable boxes (`Code = Trim(Mid(x,1,3))`, `= x`, `Like '%x%'`), no value = whole table, "Record not found!" | same matches (`SearchMatch`), toast when empty | ✓ |
| ListView + TOTAL RECORDS, ColumnClick sort | `app-data-grid` (sortable) + TOTAL RECORDS | ✓ |
| Click a row → boxes filled, disabled; Edit → code disabled | same (`selected` / `edit` modes) | ✓ |
| Save checks in form order with the MsgBox texts ("Please Enter Category Code", "… must be 3 integer!", …) | `CodeTableService.Validate`, same texts and order | ✓ |
| Add: duplicate check on the code → "Record Duplicated! Please Change The Corresponding Field" | same (409) | ✓ |
| "Do you want to save the record?" / "Do you want to delete this record?" | `ask()` confirmations | ✓ |
| Saves `…LastUpdateUser = Logon`, `…LastUpdateDate = Date` | user code + today's date | ✓ |
| "Record Saved..." / "Record Updated..." / "Record Deleted... ", list refreshed | same | ✓ |
| Delete without a searched row → "Please search the record" | same | ✓ |
| Drug Value stored as `Format(x, "#,##0.00")` text | same | ✓ |
| KeyNumber on numeric boxes, MaxLength on every box | input filter + max length (client and server) | ✓ |
| Print → DataReport (RptCheckList, RptCoPaymentList, RptFeeList, RptReasonList, rptProcedureslist, rptDruglist, rptDiagnosisList, rptRepudiationList) | PDF from `cis.ListReport_*` with the report's title and columns | ✓ |

## Known desktop issues (reproduced, for the business to decide)

- Codes are not unique in the data (e.g. CheckItems has three `BILL` rows). Edit updates the **first** row with the code; Delete removes
  **every** row with the code; Diagnosis Not Payable checks duplicates and deletes by **Plan Code only** (one diagnosis per plan through
  the screen, delete removes all of a plan's diagnoses).
- New codes must have 3 characters ("must be 3 integer!" — letters are accepted) although existing data has 4-character codes.
- Procedures / Drug search compare only the first 3 characters of the typed code (Procedure codes are up to 4).
- Tooth Reference stores only the first character of the chosen category although category codes can have 2.

## Deliberate differences

- **Print works.** The DataReports read from `DataEnvironment1`, which is not in `medixcis.vbp`, so the desktop's Print buttons could not
  run. RptAnnualList also sorted on `VisitIndex`, a column its view does not have; the web sorts on `AnnualIndex`.
- **FEE audit columns.** The desktop wrote FEE's user / date into `COPLastUpdateUser` / `COPLastUpdateDate` (not FEE columns, silently
  skipped); the web fills `FeeLastUpdateUser` / `FeeLastUpdateDate`.
- The list is ordered by code (the desktop showed rows in table order).
- One engine serves all simple tables: SQL is built only from catalog identifiers (checked, bracket-quoted) with parameterised values.

## LG Maintenance (frmLGMaintenance)

GL Type / Template Name / Decline Reason run on the engine (IDENTITY key, no Delete, no duplicate or required check, "New Record Inserted
Successfully.." / "Record Updated Successfully.."). GL Template Binding (Tbl_OPGLMaster) is its own panel (`cis.LgBinding_*`).

Known desktop issues (reproduced):
- Edit of a binding updates **every** Tbl_OPGLMaster row of the plan code, not only the selected one.
- GRPCompnay is the text between the first and second "-" of "CODE - NAME", so names with "-" are stored cut (e.g. "…SDN BHD-1" → "…SDN BHD").
- Search needs a group company ("Plese Select Group Company to Search", spelling kept); nothing is validated on Save.

Deliberate differences:
- Edit always writes the GL type name and Y/N VIP. The desktop's Edit cut the first 4 characters off the GL type unless it was re-picked
  (`Mid(cmbGLType, 5)` on a name) and failed on a re-picked VIP ("N-No" into char(1)), i.e. it corrupted or lost the save.
- Clear also empties the grid on all four tabs. The desktop's LG Clear kept the list (unlike the other Maintenance screens); after Save
  the grid is refilled by the same search as in the desktop.
- The drop-down lists retry once and offer "Reload lists" if the database does not answer.

## Benefit Clinic (frmMaintenanceBenefitsCln, frmMaintenanceBenefitsClnNew)

One page for both forms (`/maintenance/benefit/clinic`, `/clinic-new`; the second is never opened by MDIMain). Card types Classic /
Silver / Gold; the grid has every BNFCategoryNew row (Cat Code, Cat Name, RM Payable Fee). Search needs card type and a valid date
("Please key in BNF Card Type!" / "Please key in Effective Date!" / "Please key in VALID Effective Date!") and fills the fees from
STDClinicBNFSpe, else "Record not found!". Save checks "Please Enter Plan Code!" (desktop text for the card type) and "Please Enter
(Valid )effectice Date!", refuses a saved card type + date ("Record Duplicated! ..."), asks "Do you want to save this record?" and inserts
one row per category (empty fee = NULL, user / date audit columns) → "Record Saved...". Leaving Eff Date warns "Invalid date format!" /
"Date cannot greater than system date ! " but does not block, as in the desktop. Only numeric fees are kept (FGrid_Click).

Known desktop issues (reproduced):
- There is no Edit or Delete: a card type + date, once saved, can only be changed in the database.
- STDClinicBNF (no card type), read by the unused ClinicBNFlisting2, is not used by the screen.

Deliberate differences:
- Clear / Add / Search empty every fee cell. The desktop's ClearGrid stopped at row 29 while BNFCategoryNew has 31 rows, so the last
  rows kept the previous values and could be saved under the new card type.
- Save is one transaction (the desktop inserted row by row; a failure left part of the rows).
- Print works in both forms (in frmMaintenanceBenefitsClnNew PrintClinicBNF was empty).
- TOTAL RECORDS shows the rows found by Search (the desktop never updated it: always 0).
- Grid rows are ordered by category code (the desktop used table order).
