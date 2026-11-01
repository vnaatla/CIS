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
