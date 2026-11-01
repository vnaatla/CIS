# frmMaintenancePlan (Maintenance > Plan)

Route `/maintenance/plan`, access 138 (MaintenancePlan). Tabs in desktop order: HealthType · Insured Type · Plan Code · Premium · Annual
Limit · Visit Limit · MCO Fee · Co-Payment (MBM). Status: HealthType, Insured Type and Plan Code migrated; the other five still to do (the
menu item stays "not migrated" until all eight are done).

## HealthType (HLT) / Insured Type (INS) — code-table engine (`health-type`, `insured-type`)

- One-letter code + description (100), values upper-cased (ChkStr); Search: code =; "Please Enter Health Type" / "Please Enter Health
  Description", "Please Enter Insured Type" / "Please Enter Insured Type Description"; duplicate code refused; Print RptHltList / RptInsList.
- Delete: "Please Enter Health Code" / "Please Enter Insured Type", then the desktop's usage checks in order — HLT: Member (MBM), Plan
  (PLN), Age (AGE), MCO, Premium (PRM); INS: Benefit (BNF), MBM, PRM — each "This record cannot be deleted because it is used in …
  table!", then "Do you want to delete this record?" → "Record Deleted... ".

## Plan Code (CISDB PLN, VIEW_PLNListing, PLNSeq — `cis.PlanCode_*`)

- Combos: HLT, CIS PAY (reloaded for the chosen health type), HISDB PAY, clinic contracts (CheckItems CSMD), group companies (not ZD01),
  Yes / No indicators, dental / spec / MEP status and limit lists, Insured / HR approval (N-No default) / OP Family Status (B-Both
  default) / OP Dependent. Picking payor DC sets DC Status Y; DC Status Y sets the payor to DC.
- Search: plan code LIKE (box 1), pay codes / health type / group company / clinic contract / every indicator =; the 17 ListView columns;
  "Record not found". A row shows the plan read-only; Edit stays disabled (the desktop disables it on a row click).
- Save (health type N): the 14 SavePLN checks in order ("Please Enter Health Code " … "Please Select OP Dependent"), "Record Duplicated!
  Please change the corresponding field" when the Grp Comp. text is a group company name, "Do you want to save the record?", then No of
  Plan Code (1–6) rows: code = first letter of Grp Comp. + PLNSeq "000", description upper-cased, codes cut before "-" (health 1, pay /
  group 4), indicators first letter, OP Family Status as shown, limits numeric else 0 → "Record Saved..." and a list of the group company.
- Delete: "Please search the record", "This record cannot be deleted because it is used in Membership table!" (HISDB MBMCLNOthers.PLNType),
  "Do you want to delete this record?" → "Record Deleted... ", list again. Print: RptPlanList.

## Known desktop issues (reproduced)

- Plan Code has no Edit. Save writes N for the ten dental / spec / MEP indicators whatever is chosen; the Diag combo is not saved (not
  shown). The group duplicate check compares the whole "CODE - NAME" text with GRPCoName, so it practically never fires.

## Deliberate differences

- **HLT / INS Delete works.** The first usage check reads MBM, which CISDB does not have, so the desktop's Delete always failed; a check
  whose table does not exist is skipped (where MBM exists it still runs, in the same order).
- **Health type S plans are refused** with a message. The desktop wrote typed codes into the HIS plan table, which has none of the
  columns the save sets (visit-limit flags, clinic contract, DC status, Insured, limits), so its save always failed. *To confirm.*
- **Search filters work.** VIEW_PLNListing lacks the indicator columns, so any indicator filter made the desktop's search fail; the web
  reads them from PLN (same list columns).
- Plan codes skip numbers already in PLN; the plans of one Save are one transaction.
