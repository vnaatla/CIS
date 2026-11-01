# frmMaintenanceEntity (Maintenance > Entity) + frmSearchCLN

Route `/maintenance/entity`, access 137 (MaintenanceEntity). Tabs in desktop order: Clinic Group · Doctor · Clinic Info · Group Company
Clinic · Clinic Access. Status: all five tabs and frmSearchCLN migrated.

## Clinic Group (GRPClinic, `cis.EntityGroup_*`)

- Search: Code / Cln Grp. Name LIKE %x% (upper-cased), list Code + Name, "Record not found"; the boxes are cleared after a search.
- Add + Save: "Please Enter Clinic Group Name" (Initial Group Name, then Cln Grp. Name), "Please Enter Clinic Group Address" (line 1),
  "Please Enter Effective Date"; "Do you want to save the record?"; code = first letter of the Initial Group Name + the next GRPClnSeq
  number ("0000"); all text upper-cased (ChkStr), empty optional boxes NULL, Homepage "" when empty, Y/N combos as one letter, an
  invalid date saved as NULL → "Record Saved..."; then the full list again.
- Edit: code and Initial Group Name locked; same columns → "Record Updated...".
- Leaving Eff. Date warns "Invalid date format!" / "Date cannot greater than system date ! " without blocking. Tel / Fax take digits only.
- Delete: "Please Enter Clinic Group Code", "Do you want to delete the record?", "This record cannot be deleted because it is used in
  Clinic Table!", "Record Deleted... ". Print: RptGrpClinicList.

## Doctor (Doctor, `cis.EntityDoctor_*`)

- Search: Code / Name LIKE %x%; Add: "Please Enter Doctor Full Name", duplicate name refused ("Record Duplicated! ..."), code = first
  letter of the name + next DOCSeq number ("00000"), name upper-cased; Edit locks code and name (only the audit columns change, as in
  the desktop); Delete: "Please Enter Doctor Code" …; Print: RptDoctorList.

## Clinic Info (HISDB CLN, `cis.EntityClinic_*`)

- Combos: clinic groups (HISDB ClinicGrp, which mirrors CISDB GRPClinic), cities (HISMaintenance CTY), states (HISDB STA) and the fixed
  lists (Type G/S, Benefit Type 1/2/3 defaulting to 1 - Standard, Panel Y/N/D/C, X-Ray / 24 Hours / FFS / Online Y/N, GL Format M/H).
- Search: code =, initial / name / state / city / post code LIKE %x% (upper-cased), panel / type / 24 hours / FFS =, group prefix LIKE,
  Online Y vs. N-or-empty, GL format M vs. H; the 38-column list; "Record not found"; the boxes are cleared after a search. View / Back
  hides the boxes to show the list larger. Print exports the list to Excel ("Report cannot be printed without any record.").
- Save: the 15 desktop checks in order ("Please Enter Clinic Name" … " Please Operating hours", leading spaces kept); Add refuses an
  Address 1 already used ("Duplicate Record!"); "Do you want to save the record?"; code from HISMaintenance CLNSeq (first letter of the
  Clinic Initial + seq "000"); text upper-cased (ChkStr, "" when empty); the group's first 5 characters; combos as one letter; FFS %
  rounded into numeric(18,0) → "Record Saved..." / "Record Updated...", then a search on the saved boxes. Edit locks code and name.
- Delete: "Please Enter Clinic Name", "Please search the record", "There are case(s) register under this Clinic !" (CISDB CISCAS),
  "Do you want to delete the record?" → "Record Deleted... "; boxes and list cleared.

## Group Company Clinic (CISDB GRPCOMClinic / VIEW_GrpCom_Clinic, `cis.EntityGrpCom_*`)

- The "load lists" box (Check1) fills the group company (GRPCompany) and clinic (HISDB CLN) type-aheads; otherwise codes are typed.
  The code is the text before "-" (else its first 10 characters).
- Search: = group company / = clinic, 12 columns (address = 1 + 2 + 3), "Record not found". A row puts its codes back, read-only.
- Save: "Please Enter Group Company Name", "Please Enter Clinic Code", the pair must not exist ("Record Duplicated! ..." — also on Edit),
  "Do you want to save the record?"; Add inserts, Edit rewrites the selected pair → "Record Saved..." / "Record Updated...", then the whole
  list. Delete: "Please search the record", "Do you want to delete the record?" → "Record Deleted... ", then the whole list.
  Print: RptGrpComCLN.

## Clinic Access (HISDB CLNAccessControl, `cis.EntityAccess_*`) + frmSearchCLN

- Search / Refresh: the latest 100 rows (clinic code, name, from, to, created by / date). Add enables the boxes; "..." opens the clinic
  look-up (frmSearchCLN: "Please Enter Clinic Name to search", name LIKE %x%, "Record not found"; clicking a clinic fills code and name).
- Save: "Please Enter Clinic Code", "Invalid Effective Date From", "InvalidEffective Date To", "Effective Date From cannot be more than
  Effective Date To!", "Clinic Code does not exists, please enter again!", "Do you want to save the record?" → row with the upper-cased
  code, created by / now → "Record Saved...", list again. Leaving a date box warns "Invalid Effective Date From / To format!".
- No Edit, Delete or Print on this tab (the desktop has none).

## Known desktop issues (reproduced)

- Edit of a doctor cannot change anything but the audit columns (both boxes are locked).
- The Add duplicate check of Clinic Group looks up the (still empty) code, so it never finds a duplicate.
- Clinic Info: Pay. Post Code / City / State are shown from the clinic but never saved (the web shows them read-only). A letter with no
  CLNSeq row gets a doubled code (e.g. "ZZ001"). The duplicate check compares Address 1 exactly as typed.

## Deliberate differences

- **Delete works.** The desktop's Delete looked the record up in its own table and therefore always answered "This record cannot be
  deleted because it is used in …" (and for Clinic Group its DELETE named a non-existent table, ClinicGrp). The web refuses a clinic group
  only while a HISDB clinic uses it (CLN.ClinicGrpCode) and deletes from GRPClinic; a doctor is not referenced by code anywhere
  (CLN.CLNDocName is free text), so it is deleted. *To confirm with the business.*
- **Clinic Group Edit works without an Initial Group Name.** Edit locks that box, but SaveClinicGrp still required it first ("Please
  Enter Clinic Group Name"), so a group without an initial (all 2,521 locally) could never be edited. The web requires it only on Add,
  where it builds the code, and keeps the stored value on Edit.
- **Generated codes skip numbers already used.** GRPClnSeq / DOCSeq can lag behind the data (locally B is at 0001 while B0080 exists), so
  the desktop's next code hit the primary key and the save failed silently; the web takes the next free number and moves the counter.
- Clinic Info: an empty FFS % is saved as NULL (the desktop assigned "" to the numeric column and the save failed, so it only saved with a
  percentage). " Please Enter FFS Percentage" fires whenever FFS is Y (the desktop compared the combo text "Y - Yes" with "Y", so the
  check only ran after a list click). The clinic code box is not typed on Add (the desktop enabled it but ignored it).
- Group Company Clinic: the list shows at most 1,000 rows with the total ("Showing the first … of …"). The desktop listed the whole view
  (1,744,655 rows locally) on an empty search and after every Save / Delete.
- Clinic Access: dates outside smalldatetime (after 06/06/2079) are refused as invalid instead of failing the insert.
- Save / Delete are transactions; the counter update and the insert commit together.
