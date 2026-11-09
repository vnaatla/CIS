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

## Remove Worksheet — FrmBordxRevise (`/claim-bordx/remove-worksheet`, right 123, `api/claim-bordx/remove-worksheet`)

| Step | Behaviour |
|---|---|
| Go | `Please select a Bordx No `; the bordereaux's worksheets (VIEW_CIS_BordxRevise, ClaimBordxModule case 3) with the bordx date, printed date, count and total approved; `No record found!`. |
| Reject | `You are going to remove a worksheet from current Bordereaux. Are you sure?` → per ticked worksheet: the claim count comes off the case's own bordereaux and the approved amount off the bordereaux on the screen, a BordxWSArchive row, ClaimInvoice and Claim bordx fields cleared (an ASCII-sent worksheet gets ASCII N / status D) → `Record updated...`. |

Known desktop issues (reproduced): the count and the amount are taken off different bordereaux rows when the case's BordxRefNo differs from the one typed; Print does nothing.

Deliberate differences: unticking a worksheet really leaves it out (the desktop compared the member no. with the case no., so unticked worksheets were still removed); all removals are in one transaction (the desktop had none).

## ClaimBordxDate — FrmBordxDate (`/claim-bordx/bordx-date`, right 122, `api/claim-bordx/bordx-date`)

| Step | Behaviour |
|---|---|
| Search | BordxDateChange: one line per posting status (bordx no., date, `YES` / `NO`, no. of cases); `Record not found`; a posted bordereaux shows `Posting is Done. Update is not allowed`. |
| Update | `Please Enter Bordx Date`; `Posting is Done. Update is not allowed`; Bordx date / by and the claims' bordx month / year → `BordxDate Updated Successfully..!!`. |

Deliberate differences: the posting check is made again at Update on the bordx no. being updated (the desktop used the status of the last search, even if the no. was changed afterwards); both updates are in one transaction.

## Bordx Delivery — FrmSendtoPayor, FrmReturnPayor, FrmSendtoAC, FrmSendtoMNRB (`/claim-bordx/delivery/send-to-payor` · `returned` · `send-to-ac` · `send-to-mnrb`, rights 125 / 126 / 127 / 125, `api/claim-bordx/delivery/{payor,returned,ac,mnrb}`)

One page and one service for the four screens; they differ only in their filters, status / date columns and update.

| Step | Behaviour |
|---|---|
| Search | Group company (Send to MNRB: payor `code - name`), scan one bordx no. (appended) or a bordx range, bordx date range, amount range, A/C month / year ranges (Send to A/C: date sent to A/C range instead), status 0 - Not Sent / 1 - Sent / 2 - Both (not on Send to A/C). Scan checks: `Please select a Group Company` (`Please select a Payor`), `Please Enter Bordx No`; `No record found!`. Lists Bordx (ClaimBordxModule cases 7 / 8 / 10, the MNRB select) with the total amount. |
| Tick | A bordereaux already sent / returned / sent to A/C / sent to MNRB: `Record has been Updated!` (not ticked). |
| Update | `Please select at least one record to be updated! `, `Please Enter in date`, `Do you want to update the record?` → Send to Grp. Company: BordxPayToStatus 1 + date sent; Returned: BordxPayDocFRStatus 1 + date received, except when the date is before the date sent (`Date Received from Payor cannot be ealier than Date Sent to Payor!`); Send to A/C: BordxSendToACStatus 1 + date; Send to MNRB: BordxMNRBToStatus 1 + date → `Record updated...`; the range list is searched again. |
| Print | The list as CSV. |

Known desktop issues (reproduced): amounts and A/C periods are compared as text; Send to MNRB is not on the desktop menu (never opened).

Deliberate differences: parameters instead of pasted criteria; the updates of the ticked bordereaux are in one transaction; the scan duplicate check uses the bordx no. (the desktop compared the typed no. with the group company column, so a bordereaux could be listed twice); the Excel print is a CSV.

## Bor Acc Period — FrmACPeriod (`/claim-bordx/ac-period`, right 128, `api/claim-bordx/delivery/acperiod`)

The same page as Bordx Delivery with the A/C period filters and update.

| Step | Behaviour |
|---|---|
| Search | Group company, scan / bordx range, bordx date, amount, A/C period from / to (`mm/yyyy`), without A/C period; ClaimBordxModule case 11; `No record found!`. |
| Tick | A bordereaux whose A/C period status is set: `Record has been Updated!`. |
| Update | `Please select at least one record to be updated! `, `Month cannot be Empty.` (1–12), `Year cannot be empty!` (1999–3000), `Do you want to update the record?` → PAYToACPeriodMth / Yr and ACPeriodStatus 1 → `Record updated...`. |

Known desktop issues (reproduced): the period filter compares the month as text, and applies the year only for months 10–12.

Deliberate differences: a blank period is simply not filtered (the desktop's masked box compared `_` with 0 and the search failed); scan mode uses the same list as the range (the desktop used case 10, which lacks the A/C period status column).

## Monitoring — FrmBordxControl, FrmOSBordx, FrmBordxWks (`/claim-bordx/monitoring/control` · `outstanding` · `bordx-worksheet`, rights 129 / 130 / 131, `api/claim-bordx/monitoring`)

| Screen | Behaviour |
|---|---|
| Bordx Control | Group company (its subsidiaries), bordx range, bordx date range, sent / not sent, returned paid / partial, payment paid / partial / unpaid, subsidiary, cash type C / M, entity Health Connect / Sihat Johor; ClaimBordxModule1 case 17 (copied as `cis.BordxControl_Search`): per bordereaux the bill, co-pay, approved, MCO fee, bordx and non-GP amounts, the sent / returned / paid dates and days (0 shown as 1); `No record found!`. The total is the sum of the MCO fee column (as the desktop). |
| Outstanding Bordx | Group company and O/S date (today); bordereaux sent but not fully paid (VIEW_OS_CISBordx); days = O/S date − date sent + 1, amount under `<= 14 Days` or `> 14 Days`; total due; `No record found for this Group Company!`. |
| Bordx & Worksheet | Group company, payor, bordx / worksheet / receipt no. ranges, date from payor range or empty; VIEW_CISBordx_Wks rows with a payor code; count and total approved; `No record found!`. |
| Print | The list as CSV. |

Known desktop issues (reproduced): Bordx Control's "total amount" is the MCO fee total; check boxes for opposite values (e.g. sent and not sent) both apply and find nothing.

Deliberate differences: parameters instead of pasted criteria; Bordx Control with a subsidiary uses the same query as without one (the desktop switched to case 16, whose view rejects the `Bordx.` / `CISCas.` filters it always adds, so it always failed); the check boxes are drop-downs.

## Print Bordx — frmPrintBordxExcel (`/claim-bordx/print/standard`, right 124, `api/claim-bordx/print`)

| Step | Behaviour |
|---|---|
| Search | `Please enter Bordx Reference No`; the bordereaux's worksheets (CLM_PayAdviceCLNBSN case 1) with plan, case, patient, member, clinic, bill / co-pay / approved / MCO fee / bordx amounts and the printed flag; `No record found `. The print month / year come from the bordereaux and can only be chosen while it is not printed. |
| Print Details / Preview | `Please Select company name`, `Please Select Bill Type`, `Please Select Health Connect Or Sihat Johor..!`, `Please keyin Month or Year !` (not printed yet); the pay advice (Excel): summary sheet with the group company / subsidiary block (VIEW_OP_GRPSummaryALL), the month line and one line per group of the report type (group company, department, relationship, cost centre, cost centre-branch, subsidiary), and a detail sheet. Print (not Preview) stamps the bordereaux printed (date, month, year, by). |
| Print Summary / Preview | The summary sheet only. |

Deliberate differences: the ~20 Excel templates (`PayAdviceCLNBordx*.xlt`, `BordxByCostCenter*.xlt`, credit-note variants) and the Word cover letter are not available, so the pay advice is a generated workbook with the same figures (title: Invoice / Disbursement Note, `CREDIT NOTE` for KCN bordereaux); the pasted criterion is replaced by the quoted bordx no.
