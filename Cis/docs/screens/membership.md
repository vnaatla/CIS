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

## General Exclusions (Enquiry) — frmGeneralBenefitExcEnq (`/invoice/general-exclusions/enquiry`, right 113)

| Item | Behaviour |
|---|---|
| Group Company | type-ahead of GRPCompany names; Search without one → `Please Select Group Company!` |
| Remarks tab | the 24 exclusion check boxes (desktop captions, e.g. `Dermatitis/Skin Disorder`, `Tyroid Disorder`, `Prenancy Related/Test`) ticked where CISDB GeneralExclusions has `Y`, and the remarks; read-only. A company without a row shows everything unticked (no message, as the desktop). |
| Benifit Schedule tab | files of `SRVPath\CLNPolicyWording\<group company>` with the desktop's extensions (jpg, tif, gif, pdf, txt, xlsx, docx); a click opens the file. Server folder = `Files:Folders:PolicyWording`; not configured → a notice. |

Deliberate differences: the hidden Update button (which would add / overwrite the company's row) is not offered; a group company name that
would reach outside the policy wording folder is refused.

**General Exclusions (frmGeneralBenefitExclusion, right 133) cannot be migrated:** the menu opens a form that is not in the source folder
or in `medixcis.vbp`, so the desktop build cannot open it either. The menu item stays "not migrated" until the form (or its rules) is
supplied; the Enquiry above shows the same data read-only.

## MBM Adjustment — FrmMBMCLNAdj (`/membership/adjustment`, right 113, `api/membership/adjustment`)

| Item | Behaviour |
|---|---|
| Show | Member No empty → `Please Enter Membership No`; 8–9 characters → frmSearchMBM05; HISDB `SearchCISMemberProc` (principal; F / H covers); none → `No record found!`. Deleted → `Membership has been deleted. No further changes are allowed!` (no Edit); cancelled → `Membership has been canceled. No further changes are allowed!` (Edit stays for Re-Activate); suspended → `Membership has been Suspended. No further changes are allowed!` (no Edit). Prev / Next, Search (frmSearchMBM). |
| Principal tab | name, IC, policy no, payor, insured type, plan, group company / department, status, exp. / data status, eff / exp date, DOB + age band, card type (CheckItems CSMD), limit status text (AnnualLimit.SuppLimitStatus: `Share Limit For Family` …), e-mail, mobile, bank, exclusion, remarks, CIS remarks; premium, MCO, accum., annual / balance / visit limit (`FULL` when unlimited), MEP / dental / spectacles limits. |
| Supplementary tab | I / G → `This Plan does not have supplementary`; covers with their plan, card type, premium, MCO, refunds (cancelled memberships, MBMCISRefund), limits, age band. Update (A): `Please select covered persons to update !`, `Please Enter Covered Person Identification`, `Are you sure?`; Cancel Supp (K): `Please Choose a Record`, `Are you sure?`. |
| History tabs | Adjustment History (CISDB MBMCISAmendHistory), MBM History (HISDB `SearchCISMBMHistory` by IC). |
| Edit / Endorse Status | A Amendment · C Cancellation · K Cancel Supplementary · X Suspend · Y Re-Activate · D Delete Policy; `No Re-Activate allowed ! ` / `Only Re-Activate allow ! ` are warnings only, as in the desktop. Changing the policy dates recalculates premium / MCO (GetPremium, GetMCO, InsertMCOFees) and asks ` Do you want to change member's MCO Fee?`. |
| Save | `Please select a record to edit!`, `Please Enter Endt. Bordx Date!`, `Please Enter Endt. Effective Date!`, `Are you sure?`. C: an outstanding case → `There is at least an outstanding case under this member!`; else history + status C + covers C + refunds (MBMCISRefund) → `This Membership has been cancelled successfully`. K: refunds + status C for the flagged covers → `Supplementary has been cancelled successfully`. X → `Record is Suspended`; Y → `Record is Activated`; D → `Policy Deleted Successfully`; A: history + MBMCLNOthers, MBMTwo, MBMOthersTwo, covers → `Updated Successfully`. History = MBMCISAmendHistory row + MBMCLNOthers last endorsement, status, endt dates and MBMCancelDate (= endt. eff date, also for A, as the desktop). |

Deliberate differences:
- One transaction for CISDB and HISDB per save; the status of a deleted / suspended membership is also refused by the API.
- A cover's premium is saved as its premium; the desktop wrote the annual-limit column into MBMCCLNPrem.
- Cover refunds store the MCO refund in MCO Refund (the desktop showed the effective MCO in both columns).

Local environment: MBMTwo (`trgMbmContactEncrypt`), MBMOthersTwo (`trgBankAcEncrypt`) and MBM have encryption triggers that need the
HISDB database master key, which the restored local copy lacks, so Amendment saves fail here; the other endorsements do not touch them.

## MBM Enquiry — FrmMBMCLNEnq (`/membership/enquiry`, right 113, `api/membership/enquiry`)

| Item | Behaviour |
|---|---|
| Show | `Please Enter Membership No. Atleast first 8 characters required.` under 8 characters; numbers not starting DC / SL with at most 11 characters (and no `*` in position 12) go through frmSearchMBM05 (one match is taken); HISDB `SearchCISMemberProc_test`; none → `No record found!`. Prev / Next / Search as in MBM Adjustment. The plan's special remark (PLNSpecialRemark) is highlighted and shown as a message. |
| Exclusions | Empty group company → `Please Check Group Company! `; else General Exclusions (Enquiry) opens searched for the member's group company. |
| Principal Details | the member's fields, product (HISDB MBM + PLN: `ProductName(PLNDescription)`), limit status text, fees and the limits labels, plus the FPGridTopup grid: Item / GP / SP / Maternity / Dental / Optical × Limit / Utilised / Balance. |
| Limits | Top-up (CISDB MBMCISTopupHistory): limit status A → all covers, B–F → cover 00, none → `N/A`. CallLimit: ASO members (ASOInd Y) → Tbl_ASOMemberDetails (combined → ASO limit, else OP limit); else GetAnnualLimit / GetVisitLimit; GetSpecLimit (PLN.PSpecLimit F → FULL, Y → SpecialistLimit). Balance = annual + top-up − utilised (limit status E → annual `FULL`). Visit limit shows MBMCLNVisitLmt when numeric. |
| Utilisation | `CIS_FFS2018_test` claims of the member (parameterised copy `cis.MbmEnq_Utilisation`), filtered by limit status (shared GP or A → all covers; B–E, blank, F → cover 00). No specialist limit: per row the deduction of the plan's scenario (PDentalStatus / SDentalStatus / IsGPDentalShared / IsDentalOpticalShared / PSpecStatus / SSpecStatus, PDentalAnnLimit), maternity deducted for maternity plans; with a specialist limit: GP rows (all rows when PLNGPSPShared) minus others / typhoid / maternity / dental, specialist rows (CLNType S) = bordx − others − maternity, filtered by SpecialistLimit status. Maternity = AnnMaternityLmt − all maternity claims. Optical = principal optical claims; dental from `tmpDlimit_local_test1` or `tmpDentallimit_OneLimit` per scenario (`MbmEnquiryRules`). |
| Supplementary | I / G → `This Plan does not have supplementary` (back to the first tab). showMBMCov per cover: AnnualLimit / SpecialistLimit status of the cover plan on the policy date (the desktop keeps the last row it reads, i.e. the oldest), limits by CallLimit("S") (ASO: the cover's Tbl_ASOMemberDetails; status F: PLN + AnnualLimit, W/H → SuppLimit, S/D → ChildDepLimit, FULL when ANNLmtIndSupp is not Y); utilised from the claims filtered by status (B covers other than 00, C/F/blank this cover, D cover 01 only, E this cover; A → 0); limit / balance only for the first B (or D 01 / D 02) cover, balance = limit − utilised + cover 01's top-up; specialist, dental (tmpsuppdentallimit) and optical (Suppoptiocalimitlimit) per scenario. A click shows the cover's panel and grid (maternity N/A), FFS from CIS_FFS2013 and the top-ups by the principal plan's status. |
| Adjustment / MBM History | as MBM Adjustment. |
| Case | GP (VIEW_CISMBMInquiry_GPCASE) and specialist (VIEW_CISMBMInquiry_Spe_CASE) cases; totals: visits, bill, bordereaux and FFS of not-deleted GP cases; specialist bordereaux total adds the bill amount (desktop). |
| Benefit Schedule | `.jpg` / `.JPG` / `.pdf` / `.PDF` files of `CLNPolicyWording\<group company>`; a click opens the file. |

Deliberate differences:
- CIS_FFS2018_test, tmpDlimit_local_test1 and tmpDentallimit_OneLimit are called with parameters; the desktop concatenated the member number into SQL text.
- Dentallimit_OneLimit compares a blank ASOType with 2 / 3, which raises a VB type-mismatch and aborts the Show; the web treats blank as "not 2 / 3".
- CallLimit splits the member number at "-" only when the dash is the first character (desktop bug, reproduced: ASO look-ups use the whole number).
- lstCovAdd_Click writes the premium into a local variable that hides the label, so the desktop never shows a cover's premium; the web shows it.
- Known desktop issues kept: every B cover shows the shared utilised amount but only the first gets the limit; the cover balance always adds cover 01's top-up; with a specialist limit the specialist utilisation tests the GP status for blank.
- FlexCare / Family Benefits (UOB FlexCare plans, tblFlexibleBenefitPlanLimit_OP covering the expiry date): figures from `cis.MbmEnq_Flex`, a copy of spGet_UOB_Flex_New_Cal; the original reads spCalcFlexUOB through OPENROWSET(SQLNCLI, Server=sqlinstance1), an ad hoc distributed query to the production server that is disabled on SQL Server by default, so the copy (`cis.MbmEnq_CalcFlex`) fills a shared temp table instead (`365_mbm_enquiry_flex.sql`). FlexCare tab = cover 00 (capped items, items utilised, totals with the last three highlighted); Family Benefits = the covered person selected on Supplementary (family limit / utilised / balance for cover 01 only, else `Please select cover id 01 for available Limits & Balance`; out-patient and other items).
- The desktop overrides the principal's accumulated amount with the FlexCare SumOP using the previous Show's state (isUOBFlex is set after showMBM runs); the web always shows the computed amount.
- UB plans with PLN.PlncategoryInd: MessageDisplay notes on Principal Details (lblPLimitStatus) and Supplementary (lblDLimitStatus): the out-patient note of the AnnualLimit status plus the dental / optical note of the scenario, with Dental_Utilization_Limit_Sum utilisation and the CategoryAmountCal home-office limit (the desktop's concatenation that repeats the out-patient note before the home-office line is kept). NolimitAmount only fills a variable the desktop never shows, so it is not called.
