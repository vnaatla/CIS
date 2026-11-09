# Claim Bordx screens (Phase 6)

## Insert Worksheet — frmWSBordx (`/claim-bordx/insert-worksheet`, right 122, `api/claim-bordx/insert-worksheet`)

| Step | Behaviour |
|---|---|
| Search | Bordx type INSURER-KIS / ASO-KDC (Pure ASO, BTBG or ASO Type) / CREDIT NOTE-CN / SIHAT JOHOR-KSJ / VOID-KIV, entity Health Connect / Sihat Johor, group company (name\|code; its departments and subsidiaries), scan one worksheet no. (appended; `Data has been uploaded!` if already listed) or a case range, visit / registration dates, empty visit date, clinic type, panel clinic, OP batch range, cash type, treatment, OMNI. Checks: `Please select a Group Company Name`, `Please Enter Worksheet No` (scan), `Please select Bordx Type!`, `Please select Bordx Type whether ASO or BTBG..!`; `No record found!`. The list is ClaimBordxModule_isPureASO case 21 (copied as `cis.BordxWs_Search`); for the UOB companies the Outpatient / flexCare drop-down lists ClaimBordxModule case 24 / 23 (`cis.BordxWs_SearchUob`) and Search is disabled. |
| Tick | The desktop's checks, in order: `Date of Visit cannot be empty.`, `Clinic Name cannot be empty!`, `Group Company cannot be empty!`, `Batch Number is empty!`, `FFS is empty!`, `Group Company's FFS Status is empty!`, `Please check Bordx Amount!`, `Difference in Bill Amount and Payable to Clinic Amount!`, `Please check Bill Amount & Bordx Amount!` (not for BTBG plans). Check All ticks every row that passes. |
| Bordx | `Please select a record before proceeding!`, `Please select at least one record to be updated! `, `Please select Bordx Type!`, `Please select Flexcare or Outpatient Type!`; `Do you want to update the record?` → in one transaction: the bordx no. (`K` + payor + the next number of CLAIM_BordxSEQ, `KIV…` for void), per case Claim / ClaimInvoice bordx fields (BTBG fields for ASO-KDC / CN with a BTBG plan; void also writes a `VOID` Payment2MBM row for a member reimbursement), CISCas Pending = N, and the Bordx row (amount, number of claims, group company, by-department flag, subsidiary) → `Record Saved, Your Bordx No is : …`. |
| Print | The list as CSV (`Report cannot be printed without any record.`). |

Known desktop issues (reproduced):
- The BTBG check at Bordx uses the plan of the last ticked case for all cases.
- The Bordx row's payor code is `DC` whenever a group company was chosen (blank otherwise), whatever the bordx type.
- The UOB test applies "single company" only to UOB Malaysia (operator precedence): UOB Properties / Limited as the first company switch to the UOB lists even with several companies.
- The duplicate-GP-visit scan in the tick routine is computed but never used.

Deliberate differences: parameters instead of the criteria pasted into the procedures; the CISCas / Payment2MBM updates are in the same transaction (the desktop's CISDB / CISPayment writes were outside its CISWorksheet transaction); the list prints as CSV; double-clicking a row (opens Invoice Adjustment in the desktop) is not provided — open the case in Invoice Adjustment.
