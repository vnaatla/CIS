# Membership — developer test guide (Phase 2)

Manual test script for the Membership screens, with real data from the local databases (checked on 2 Nov 2026). Setup, login and the
general checks are the same as in `maintenance-test-guide.md` sections 0–1. Desktop texts are quoted exactly.

Test member used below: **DCHC9086608** (NUR RABIZAH BINTI ADENI, IC `820615136152`, plan `S079`, insured type `H`, eff 01/08/2021 –
31/12/2021, group company SACOFA SDN BHD, limit type `A - Share Limit for Family`, GP balance **3,129.00**; covers `01`–`03` are family members).

## 1. Member search pop-ups (frmSearchMBM / frmSearchMBMIC / frmSearchMBM05)

Open from MBM Topup Adj › **Search** (later also MBM Adjustment, MBM Enquiry and the invoices).

| # | Action | Expected |
|---|---|---|
| S1 | Search with `ab` | `Search String must be at least 3 characters!` |
| S2 | Mem Name (default) `NUR RABIZAH BINTI ADENI` → Search | DCHC9086608 / 00 / NUR RABIZAH BINTI ADENI / 820615136152 / 01/08/2021 / 31/12/2021 (about 2–5 s: name search scans ~1 million members) |
| S3 | Mem Num `DCHC9086608` | the principal **and** covers 01–03 (supplementary rows show PolSt / DataSt); Principal 1, Supp 3 |
| S4 | IC Num `820615136152` | the principal; `820615-13-6152` finds it too (12+ characters match with / without dashes) |
| S5 | Employee Num / Policy Sub Num | principals only (no supplementary search, as the desktop) |
| S6 | Mem Name `ZZZZZZ` | `No record found!`, list empty |
| S7 | Click a row | pop-up closes, the caller loads the member |
| S8 | Clear | text, list and counters empty, Mem Name selected |
| S9 | Member box with 8–9 characters (e.g. `DCHC9086`) → Go | frmSearchMBM05 lists every member number starting with it (one match is loaded directly); Print → CSV; none → `No record found.` |
| S10 | Principal rows | PolSt, DataSt, Group Company, Dept are blank (known desktop behaviour) |

## 2. MBM Topup Adj — `/membership/topup` (FrmMBMTopupLmt, right 113)

| # | Action | Expected |
|---|---|---|
| T1 | Go with an empty member box | `Please Enter Membership No` |
| T2 | `DCHC9086608` → Go | name, IC, plan S079, Ins Type H, dates, status `A - Active`, group company, Limit Type `A - Share Limit for Family`, GP selected, Balance Lmt `3,129.00` |
| T3 | `DCHC9086608-01` → Go | JAYA AFFANDI BINTI ZAINI (cover 01) with its own IC |
| T4 | Switch to SP | Balance Lmt reloads from the SP balance (`NA` → 0.00 for this member) |
| T5 | Topup Amt `abc` | only digits / one decimal point are kept |
| T6 | Topup Amt `100` | New Bal. `3,229.00` |
| T7 | Save with Topup Amt empty and IC unchanged | `Nothing to Adjust ! ` |
| T8 | Topup Amt `100` → Save | `Updated Successfully`; MBMCISTopupHistory row (TopupAmt 100, PrevLimit 3129, NewLimit 3229, UpdateBy DEVADMIN); HISDB MBMCLNOthers.MBMCLNBalLmt = 3229 |
| T9 | Prev / Next | the previous / next clinic member number is shown as `NUMBER- 00` and loaded |
| T10 | Change the IC and Save | **fails locally**: the restored HISDB has no database master key for the `MBM` IC-encryption trigger (works on a real server) |
| T11 | Unknown member `ZZ000000000` → Go | `No record found!` |

Clean-up after T8 (PowerShell, `sqlcmd -S localhost -E -d CISDB`):

```sql
DELETE dbo.MBMCISTopupHistory WHERE MBMNumber = 'DCHC9086608' AND UpdateBy = 'DEVADMIN' AND UpdateDate >= CAST(GETDATE() AS date);
UPDATE HISDB.dbo.MBMCLNOthers SET MBMCLNBalLmt = 3129.00 WHERE MBMNumber = 'DCHC9086608';  -- value on 2 Nov 2026
```

Check `SELECT MBMCLNBalLmt FROM HISDB.dbo.MBMCLNOthers WHERE MBMNumber = 'DCHC9086608'` **before** T8 (3129.00 on 2 Nov 2026) and
restore that value.

## 3. Case Transfer — `/invoice/case-transfer` (frmMembershipTransfer, right 162)

Test case **K001596118** (member DCGB3518731, cover 00, SHER BAHADUR TAMANG, visit 04/08/2014).

| # | Action | Expected |
|---|---|---|
| C1 | Go with Case No. empty | `Please select Case Number ! ` |
| C2 | `ZZNOCASE` → Go | `Record not found!` |
| C3 | `K001596118` → Go | Current Member DCGB3518731 / 00 / SHER BAHADUR TAMANG / 04/08/2014; Transfer disabled |
| C4 | Search with Member No. empty / Cov ID empty | `Please enter Member Number! ` / `Please select Cover ID! ` |
| C5 | Member `DCHC9086608`, Cov `00` → Search | Patient NUR RABIZAH BINTI ADENI, Transfer enabled; `01` → JAYA AFFANDI BINTI ZAINI |
| C6 | Change Member No. after Search | Transfer disabled again |
| C7 | Transfer → Yes | `Record Updated!`; CISCAS row now DCHC9086608 / 00 |

Clean-up after C7: `UPDATE CISDB.dbo.CISCAS SET MBMnumber = 'DCGB3518731', CoverID = '00' WHERE CASNumber = 'K001596118';`

## 4. Clinic member registration (frmMemberRegistrationCLN)

Opened from the clinic invoice (Phase 3). Until then test it through the API (`/swagger` › MemberRegistration):

| # | Call | Expected |
|---|---|---|
| R1 | `GET setup` | races C / I / M / O, staff categories (Crew = G007 / G, Management Family = G001 / H …), group companies |
| R2 | `GET member?ic=820615136152` | NUR RABIZAH BINTI ADENI, `active: false` (policy ended 31/12/2021) |
| R3 | `GET age?effDate=01/11/2026&dob=15/06/1982` | `43Y`, band `02 - …`, adult `A` (desktop GetAge quirk) |
| R4 | `GET limits?staffCategory=Crew&effDate=01/11/2026` | annual / visit limit of plan G007 |
| R5 | `GET departments?text=` | `Please enter Department! ` |
| R6 | `POST` with each required field empty in turn | the 11 messages in order (section above) |
| R7 | `POST` complete | locally: fails on the MBM encryption trigger (no master key); on a server with the key: `Membership Generated, Your Membership Number is : DCG…` |

## 5. General Exclusions (Enquiry) — `/invoice/general-exclusions/enquiry` (right 113)

| # | Action | Expected |
|---|---|---|
| E1 | Search without a group company | `Please Select Group Company!` |
| E2 | `1MILLENNIUM INTERNATIONAL SDN BHD` → Search | Dental and Optical ticked, Dermatitis not; remarks start `PAYOR : ST` |
| E3 | `ZZ NO COMPANY` → Search | every box unticked, remarks empty, no message |
| E4 | Benifit Schedule tab | the company's policy wording files, or the "not configured" notice when `Files:Folders:PolicyWording` is empty (local default) |

## 6. MBM Adjustment — `/membership/adjustment` (FrmMBMCLNAdj, right 113)

| # | Action | Expected |
|---|---|---|
| A1 | Show with an empty member | `Please Enter Membership No` |
| A2 | `DCHC9086608` → Show | principal NUR RABIZAH BINTI ADENI, H / S079, limit status `Share Limit For Family`, annual limit 7,500.00; Edit enabled |
| A3 | Supplementary tab | covers 01–03 (JAYA AFFANDI BINTI ZAINI …) |
| A4 | Adjustment History / MBM History tabs | the member's amendments / memberships of the same IC |
| A5 | Edit → Save without dates | `Please Enter Endt. Bordx Date!` / `Please Enter Endt. Effective Date!` |
| A6 | Edit → Endorse `Y` on an active member | warning `No Re-Activate allowed ! ` (not blocking) |
| A7 | Edit → `A`, change Eff. Date | premium / MCO recalculated; MCO differs → ` Do you want to change member's MCO Fee?` |
| A8 | Edit → `C` on a member with an open case | `There is at least an outstanding case under this member!` |
| A9 | Amendment save | locally fails on the HISDB master key (encryption triggers); on a server: `Updated Successfully` |

Clean-up after a test endorsement: delete the MBMCISAmendHistory / MBMCISRefund rows of the member created by DEVADMIN today and restore
the MBMCLNOthers / MBMCoveredPersonsCLN status columns (note them before testing).
