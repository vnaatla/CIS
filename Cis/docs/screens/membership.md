# Membership (Phase 2)

## Member search pop-ups — frmSearchMBM, frmSearchMBMIC, frmSearchMBM05 (`app-member-search-modal`, `api/membership/search`)

Opened from MBM Topup Adj, MBM Adjustment, MBM Enquiry and the invoices (rights 113 / 117 / 118). A click on a row returns it (the
desktop's double-click); the caller puts `NUMBER-ID` (Topup, invoices) or `NUMBER` (Adjustment, Enquiry) in its member box and runs Go / Show.

| Pop-up | Behaviour |
|---|---|
| frmSearchMBM (`mode="search"`) | Search box + Mem Num / Policy Num / **Mem Name** (default) / IC Num / Employee Num / Policy Sub Num. Fewer than 3 characters → `Search String must be at least 3 characters!`. Principals (HISDB MBMCLNOthers + HISMaintenance MBMCrossReference + MBMOthers), then supplementary members (MBMCoveredPersons) for Mem Num / Name / IC only. IC Num with 12+ characters also matches the IC with / without dashes. Nothing → `No record found!`. Counters Principal / Supp / Record(s). Clear resets the text, the list and Mem Name. |
| frmSearchMBMIC (`mode="ic"`) | Runs on open with the patient IC (exact), principals and supplementary members; Print → Excel (CSV). |
| frmSearchMBM05 (`mode="prefix"`) | Runs on open with the caller's member number: HISDB `Search_CISMBM_05` / `Search_CISMBMSupp05` (member numbers starting with it, HSCLNInd ≠ H), group company from CISDB PLN; `No record found.`; Print → Excel. |

Columns: Member No, ID, Member Name, IC No., PolSt, DataSt (05 / IC: Policy No, Status), Eff Date, Exp Date, Group Company, Dept.

Known desktop issues (reproduced): for principals frmSearchMBM fills only Member No, ID, Name, IC and the two dates (PolSt, DataSt, Group
Company, Dept stay blank although the query returns them).

Deliberate differences:
- **Parameterized search.** The desktop sent the filter as SQL text into HISDB `Search_MBMCLN2013` / `Search_MBMCovCLN2013`, which run it
  with `sp_executesql` (SQL injection through the search box). `cis.MemberSearch_Principals` / `_Supps` run the same joins and columns with
  parameters, one branch per search type.
- **No duplicate rows for a name search.** The desktop's second pass re-ran the first command, so every principal was listed twice.
- At most 5,000 principals and 5,000 supplementary members per search, with a notice (the desktop streamed every match with a Stop button).

## MBM Topup Adj — FrmMBMTopupLmt (`/membership/topup`, right 113, `api/membership/topup`)

| Item | Behaviour |
|---|---|
| Member No | `NUMBER` or `NUMBER-ID` (no dash = cover 00), upper-cased. Go with 8 or 9 characters opens frmSearchMBM05: one match is taken directly, several are listed. Empty → `Please Enter Membership No`. |
| Go (FindMBM) | HISDB `SearchCISMemberProc`: name, IC, plan, insured type, eff / exp date, status (`A - Active`, `C - Cancelled`, `Delete - Delete`); for a cover other than 00 the covered person's name, IC, plan, balance limit and status. Group company / department from MBMOthers. GP selected. |
| Limit type | AnnualLimit.SuppLimitStatus of the plan + insured type on or before the eff date (none = N); shown for F / H as `code - LimitDesc` (HISDB AnnualLimitType); `N - Unlimited Limit` disables Topup Amt. |
| Balance Lmt | CISDB `spUtilisationOP_test1`: GPBalance (GP) or SPBalance (SP) of the cover's row; switching GP / SP reloads it. New Bal. = Balance + Topup. |
| Prev / Next | Previous / next clinic member number (HSCLNInd ≠ H) → `NUMBER- 00` and Go; ends → `First Record Reached!` / `Not a valid Membership Number Format!`, ` Last Record Reached! `. |
| Search | frmSearchMBM. |
| Save | Topup 0 and IC unchanged → `Nothing to Adjust ! `. Otherwise: MBMCISTopupHistory row (GP: IC, previous IC, TopupAmt; SP: SP_TopupAmt); GP with a changed IC updates MBM + MBMCrossReference (cover 00) or MBMCoveredPersons; a new limit ≠ 0 goes to MBMCLNOthers.MBMCLNBalLmt (limit type A or cover 00) or MBMCoveredPersonsCLN.MBMCCLNBalLmt (cover 01 for type B, the cover otherwise; nothing for N) → `Updated Successfully`. |

Deliberate differences:
- One transaction for CISDB, HISDB and HISMaintenance (the desktop committed three connections one after the other).
- The SP history row stores the cover id as text (`00`); the desktop wrote it unquoted, so `00` became `0`.
- A member number that is not found → `No record found!` (the desktop left the boxes empty without a message).

Local environment: an IC change updates HISDB `MBM`, whose trigger `trgICEncrypt` opens a symmetric key; the restored local HISDB has no
database master key, so that save fails here (`Please create a master key in the database…`) as it does in the desktop. Top-ups without
an IC change work.

## Case Transfer — frmMembershipTransfer (`/invoice/case-transfer`, right 162, `api/invoice/case-transfer`)

| Item | Behaviour |
|---|---|
| Go | Case No. empty → `Please select Case Number ! `; CISDB `VIEW_CISMBMTransfer` → current Member No., Cov ID, Patient (principal name for cover 00, covered person otherwise), Visit Date; none → `Record not found!` |
| Search | Member No. empty → `Please enter Member Number! `; Cov ID empty → `Please select Cover ID! `; name from HISDB MBM (cover 00) or MBMCoveredPersons → shown, Transfer enabled; none → `Record not found!` |
| Transfer | CISCAS.MBMnumber / CoverID of the case → `Record Updated!`; any change to the case or member boxes disables it again |

Deliberate differences:
- **Transfer works.** In the desktop the visible Search button only ran HISWorkSheet `SP_GETcasetransfer` and discarded the result; the
  button that fills the name and enables Transfer (`cmdSearch`, ">") is hidden, and every edit disables Transfer, so a case could never
  be moved. The web Search runs the hidden button's logic; Save checks the case and member again and asks for confirmation.
- The Current / Transfer frames' Dept Code, Cost Center, Dept Name, Subsidiary, Company Name boxes and "If Data Correct" ticks are not
  shown: the desktop never fills or reads them.

## Clinic member registration — frmMemberRegistrationCLN (`app-member-registration-modal`, `api/membership/registration`)

Pop-up of the clinic invoice (rights 117 / 118); built in Phase 2, opened from the invoice in Phase 3.

| Item | Behaviour |
|---|---|
| Lists | Race (HISDB `SearchRace`), Gender M / F, Staff Category (CISDB WebCategory → plan + insured type), Group Company (GRPCompany names); Department: type part of a name → Search (GroupCompanyDept of G001; empty → `Please enter Department! `) |
| Search MBM | HISDB `SP_SearchCISMBM_IC` by the IC: fills name, race, gender, DOB, department, group company, employee no.; a policy still in force also fills the dates, Staff Category `CREW`, shows the member number and disables Register; none → `No record found!` |
| Effective Date | Expiry = 31/12 of that year; limits (CallCLNLimit) and age band (CLNMemberCallAge); invalid → `Invalid Date Format! ` |
| Limits | Annual: PLN.ANNLmtInd N → 0, else the latest AnnualLimit (health N); Visit: PLN.VISLmtInd N → 0, the DC plan row with N → `FULL`, else the latest VisitLimit |
| Age | Medix.bas GetAge, reproduced as written: up to 17 years the age is `-1D` / `0Y` / `-1Y`; a later birth day counts one year less even when the birth month has passed (15/06/1982 at 01/11/2026 = 43Y); band from HISDB AGE, band 01 = child (C), else adult (A) |
| Register | Checks in order: `Please enter IC No! ` · `Please enter Member Name! ` · `Please enter Race Code! ` · `Please enter Gender! ` · `Please enter Staff Category! ` · `Please enter Employee No! ` · `Please enter Effective Date! ` · `Please enter Expiry Date! ` · `Please enter Birth Date! ` · `Please enter Group Company! ` · `Please enter Department! `. Number = `DC` + insured type + MBMSEQ (letter + 5 digits) + MBMCheckDigits (2 digits); inserts MBM, MBMTwo, MBMOthers, MBMOthersTwo, MBMCrossReference, MBMCLNOthers → `Membership Generated, Your Membership Number is : …`; clicking the number returns it to the invoice. |

Deliberate differences:
- Number generation and the six inserts are one transaction with locks on MBMSEQ / MBMCheckDigits (the desktop took the numbers before
  its transaction, so a failed save used them up, and two users could get the same number).
- Limits are written as numbers: the desktop put the formatted annual limit (`1,000.00`) and a visit limit of `FULL` straight into the
  INSERT text, which failed for annual limits of 1,000 or more and for unlimited plans; `FULL` is stored as 0.
- The server recomputes plan, limits and age band at Register instead of trusting the labels.

Local environment: the MBM insert fires `trgICEncrypt` (symmetric key), which fails on the restored HISDB without a database master key,
so Register cannot complete locally (the test accepts that error); it works where the key exists.
