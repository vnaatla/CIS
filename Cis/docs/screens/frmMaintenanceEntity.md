# frmMaintenanceEntity (Maintenance > Entity) + frmSearchCLN

Route `/maintenance/entity`, access 137 (MaintenanceEntity). Tabs in desktop order: Clinic Group · Doctor · Clinic Info · Group Company
Clinic · Clinic Access. Status: Clinic Group and Doctor migrated; Clinic Info, Group Company Clinic and Clinic Access still to do (the
menu item stays "not migrated" until all five are done).

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

## Known desktop issues (reproduced)

- Edit of a doctor cannot change anything but the audit columns (both boxes are locked).
- The Add duplicate check of Clinic Group looks up the (still empty) code, so it never finds a duplicate.

## Deliberate differences

- **Delete works.** The desktop's Delete looked the record up in its own table and therefore always answered "This record cannot be
  deleted because it is used in …" (and for Clinic Group its DELETE named a non-existent table, ClinicGrp). The web refuses a clinic group
  only while a HISDB clinic uses it (CLN.ClinicGrpCode) and deletes from GRPClinic; a doctor is not referenced by code anywhere
  (CLN.CLNDocName is free text), so it is deleted. *To confirm with the business.*
- **Generated codes skip numbers already used.** GRPClnSeq / DOCSeq can lag behind the data (locally B is at 0001 while B0080 exists), so
  the desktop's next code hit the primary key and the save failed silently; the web takes the next free number and moves the counter.
- Save / Delete are transactions; the counter update and the insert commit together.
