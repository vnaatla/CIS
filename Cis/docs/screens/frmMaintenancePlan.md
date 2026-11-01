# frmMaintenancePlan (Maintenance > Plan)

Route `/maintenance/plan`, access 138 (MaintenancePlan). Tabs in desktop order: HealthType · Insured Type · Plan Code · Premium · Annual
Limit · Visit Limit · MCO Fee · Co-Payment (MBM). Status: HealthType, Insured Type, Plan Code, Premium, Annual Limit and Visit Limit migrated; MCO Fee and Co-Payment (MBM) still to do (the
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

## Premium (CISDB PRM — `cis.PlanPremium_*`)

- Combos: INS, HLT, PAY (reloaded for the health type), AGE, PLN; Supplementary A - Share Limit for Family (default) / B - Share Limit for
  Supplementary / C - Separate Limit for Supplementary; age bands 01–07 with premium and supplementary premium (supplementary only for B / C).
- Search: insured / payor / health / age = (code before "-"), plan = (code before "-", else its first 2 characters); 11 columns; "Record
  not found". A row shows its keys and the band amounts of insured type + plan, read-only.
- Save: "Please Enter Insured Code", "Please Enter Health Code", "Please Enter Plan Code", "Please validate your Premium Effective date",
  duplicate check, "Do you want to save this record?", then seven PRM rows: codes ChkStr(Mid(…, 1, 2)) (plan before "-"), a band with a
  premium gets age code "0n" and its amount, an empty band keeps the Age combo's code and no amount; status, date, version →
  "Record Saved...". Delete removes the selected PRM row ("Please Enter Insured Code", "Please search the record", membership check,
  "Do you want to delete this record?" → "Record Deleted... "). Print: RptPremiumList.

## Annual Limit (CISDB AnnualLimit, VIEW_ANN_PAY_PLN — `cis.PlanAnnual_*`)

- Combos: PLN, HLT, INS, PAY, distinct PLN.GRPCompany; Supplementary A / B / C (default A; Supp Premium Limit only for B / C).
- Search: plan / health / insured / payor = (code before "-"), group company prefix; 17 columns; "Record not found". A row fills the boxes
  read-only; Edit works (updates the selected AnnualIndex).
- Save: the 13 SaveANN checks in order ("Please Enter Plan Code" … "Please Enter Supplementary MEP Annual Limit" for B / C), Add refuses
  insured + health + plan + version + date already saved ("Record Duplicated! / Please Change The Corresponding Field"), "Do you want
  to save the record?", supplementary limits written only for B / C (A leaves them as they are) → "Record Saved..." / "Record Updated...".
- Delete: "Please Enter Insured Code" (desktop text, checked on the plan box), "Please search the record", member check (MBM, only where
  it exists), "Do you want to delete the record?" → "Record Deleted...", then the whole list. Print: RptAnnualList.

## Visit Limit (CISDB VisitLimit, VIEW_VISIT_PAY_PLN — `cis.PlanVisit_*`)

- Same combos and search as Annual Limit (11 columns). Save: "Please Enter Plan Code", "Please Enter Health Code", "Please Enter Insured
  Code", "Please Enter Visit Limit", "Please Enter Visit Limit Effective Date", "Please Enter Visit Limit Version", "Please Enter Supp
  Visit Limit" (B / C); Add refuses insured + health + plan + version + date already saved; Supp. Visit Lmt written only for B / C →
  "Record Saved..." / "Record Updated...". Edit works (selected VisitIndex). Delete as Annual Limit (member check where MBM exists,
  "Record Deleted...", whole list). Print: RptVisitList.

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
- **Premium:** the duplicate check uses the effective date (the desktop formatted the age-01 premium as the date, so it never matched);
  for B / C the supplementary premium of a band is saved (0 when empty) — the desktop's condition was inverted: an entered amount saved 0
  and an empty one crashed the save (no B / C rows exist locally). Edit is not offered: the desktop's Edit re-opened the row and saved
  nothing. The membership check on Delete reads MBM only where CISDB has it (it does not locally, where the desktop's Delete failed).
  The seven rows of one Save are one transaction.
