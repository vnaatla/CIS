# A/C Export screens (Phase 9)

Batch numbers: the posting screens use one of two desktop sequences, both locked until the post is saved:
`cis.SunPost_NextBatch` (CISWorksheet SunPostBatch, Medix.bas GenSUNPostingBatchNo: the last number + 1, 1001 when empty) and
`cis.SunPost_NextInvNo` (PostingINVSEQ, GeneratePostingINVNo: a letter and 6 digits, the next letter after 999999).

## Invoice Posting > Clinic — frmPostingInvoiceCln (`/ac-export/invoice/clinic`, right AcExport, `api/ac-export/invoice-clinic`)

| Step | Behaviour |
|---|---|
| Search / Search Debit / Search Credit | CISWorksheet SUN_CIS_Posting `InvCln` (VIEW_SUN_Posting_InvCln, as `cis.SunPost_InvCln_Search`): clinic, invoice date and no. ranges, post status 0 - No (not posted) / 1 - Yes, post date, batch no., bordx date (its month and year compared separately), the posted A/C period, PV date (`Invalid Date Format! `). Search lists two lines per invoice: credit `84` + clinic (clinic name) and debit `6000` (YYMM-clinic name), journal RC, ledger UHC, amount = payable to clinic rounded to 2 places, analysis `4` + clinic; Search Debit / Credit list only one side. Totals, record count, `No record selected ! `. |
| Tick / Check All | A posted list: `Record(s) has been posted ! `; an invoice's debit and credit lines are ticked together. |
| Post | `Please select the record`, `Record(s) has been posted ! `, `Please Enter Bordx Date! ` (from and to), `Please Select Accounting Period ! `; `Do you want to Post record(s) ?` → a batch no. (PostingINVSEQ) and a SUNPosting row per ticked invoice (InvoiceNo, clinic as TransRef, status 1, type InvCln, today, user, A/C month / year, batch no., invoice date) → the batch's lines are listed → the posting workbook `<batch> RC-CLNInv <bordx from yyyyMMdd>-<bordx to dd>` → `Record(s) Posted Sucessfully... ! `. |
| Print to File | The listed lines as a workbook (`Report cannot be printed without any record.`). |

Fixed desktop bugs: the bordx month and year are compared as one period (the desktop compared them separately, so 12/2020 – 01/2021 found nothing).

Deliberate differences: parameters instead of pasted criteria; the group company filter is not offered (VIEW_SUN_Posting_InvCln has no group company column, so the desktop's filter failed); the ticked invoices are checked to be still unposted on the server; one transaction with the sequence locked; the posting workbook is downloaded (the desktop saved it on `\\Mars\accounts\Account Journal Import\JIS CLN Not Posted\`).

## Invoice Posting > Clinic (GST) — frmPostingInvoiceClnGST (`/ac-export/invoice/clinic-gst`, right AcExport, `api/ac-export/invoice-clinic-gst`)

The clinic posting screen with route variant `gst`.

| Step | Behaviour |
|---|---|
| Search | SUN_CIS_Posting `InvClnGST` (VIEW_SUN_Posting_InvClnGST, as `cis.SunPost_InvClnGST_Search`): clinic, invoice date and no., post status, post date, batch no., SUN invoice no. (CLNInvoiceBatch), posted A/C period. Line date = today, the batch column shows the SUN invoice batch. No FFS fee: `84` + clinic C and `6000` D (payable). With an FFS fee (MCO processing fee): `6000` D (actual − co-pay, the line Post writes), `84` + clinic C (payable), `1002` C (the fee × 94.34 % from 01/04/2015, else the whole fee) and `8830` C (the GST, fee − 1002 amount). |
| Post | As the clinic screen, but `Please Enter Invoice Date! ` (invoice dates instead of bordx dates). Batch no. from SunPostBatch; each invoice gets a SUN invoice batch `T` (FFS fee with GST) or `I` + PostingINVSEQ; SUNPosting rows with type InvClnGST; file `<batch> RC-CLNInv <invoice from yyyyMMdd>-<invoice to dd>`. |

Fixed desktop bugs: each line adds its own amount to the debit / credit totals (the desktop added the payable amount for every line, so an FFS invoice counted it three times as credit).

Deliberate differences: as the clinic screen (no group company filter, the view has no such column).

## Invoice Posting > Clinic (GST-2017) — frmPostingInvClnGST2017 (`/ac-export/invoice/clinic-gst-2017`, right AcExport, `api/ac-export/invoice-clinic-gst-2017`)

The GST screen (route variant `gst2017`) over SUN_CIS_Posting `InvClnGST2017` (VIEW_SUN_Posting_InvClnGST_2017, as `cis.SunPost_InvClnGST2017_Search`). Differences:

| Step | Behaviour |
|---|---|
| Search | Posted / not posted is PostStatus only; the payor column is blank. |
| Post | One SUNPosting row per case of each ticked invoice (CISCAS by invoice no., clinic and invoice date; the case no. as InvoiceNo), type InvClnGST2017. |
| Generate Tax Invoice | For the selected line's SUN invoice batch (`Invoice posting is not done,Generate is not possible` when blank): the clinic (representative, name, address, phone), `FEE FOR SERVICE FOR <batch>`, the FFS fee without GST and the GST, the batch and today → `Tax Invoice Generated Successfully`. |

Deliberate differences: the `HC FFS Tax Invoice.xlt` template is not available, so the tax invoice is a generated workbook downloaded in the browser (the desktop saved it on `\svr-doc\Template\FFSTaxInv\<batch>\`).

## Invoice Posting > Clinic (GST-2017 Zeroised) — frmPostingInvClnGST2017Zeroised (`/ac-export/invoice/clinic-gst-2017-zeroised`, `api/ac-export/invoice-clinic-gst-2017-zeroised`)

The desktop menu never opens this form; it is migrated as GST-2017 with one difference: in the list the whole FFS fee goes to `1002` and the `8830` GST line is 0.00 (route variant `zeroised`). Post (type InvClnGST2017, a `T` SUN invoice batch for FFS invoices) and Generate Tax Invoice are those of GST-2017.

Fixed desktop bugs: Generate Tax Invoice shows the fee without GST, as the zeroised list (the desktop still split it into 94.34 % and GST).

## Invoice Posting > Member — frmPostingInvoiceMember (`/ac-export/invoice/member`, right AcExport, `api/ac-export/invoice-member`)

| Step | Behaviour |
|---|---|
| Search / Search Debit / Search Credit | SUN_CIS_Posting `InvMBM` (VIEW_SUN_Posting_InvMBM, as `cis.SunPost_InvMbm_Search`): group company, invoice date, case no., post status, post date, batch no., bordx date, bordx no., invoice registration period. Per case: `6000` D (actual = payable to member), `82` + group company code C (actual − postage) and, with postage, `4040` C (postage); journal RC, ledger UHC, analysis 1R, the patient (member or covered person, 25 characters), bordx period YYMM, bordx no., the member no.'s last 4 digits. Search Credit lists the `82` (actual − postage) and `4040` lines (the desktop listed `82` twice, with the actual and with actual − postage). |
| Post | `Please select the record`, `Record(s) has been posted ! `, `Please Enter Bordx Date! `, `Please Keyin for Accounting Period ! `; `Do you want to Post record(s) ?` → a SunPostBatch batch no. and a SUNPosting row per ticked case (TransRef = case no., InvoiceNo = the row's batch no., i.e. blank, type InvMBM) → file `<batch> RC-MBMInv <bordx from yyyyMMdd>-<bordx to dd>`. |

Deliberate differences: VIEW_SUN_Posting_InvMBM takes over 4 minutes for a 5-day search, so the procedure picks the cases through the most selective filter given (batch, invoice date / case no., batch range, bordx no.) and applies the view's joins and conditions to them (about 15 s; a search with none of these filters reads everything, as the desktop).

## Invoice Posting — frmInvPosting (`/ac-export/invoice/posting`, right AcExport, `api/ac-export/invoice-posting`)

The desktop menu never opens this form; it is migrated on request. It is the form Pay Clinic Post was copied from, on the invoice posting flag (CISCas.CASInvPost).

| Step | Behaviour |
|---|---|
| Search | Payor (first 2 characters), invoice no. range, register date range (`Invalid date format!`, `Invoice Date cannot greater than system date ! `), post status (default 0 - No) → VIEW_InvSummary_Posting (as `cis.InvPost_Search`): per summary row a credit line `84` + clinic (invoice period, payable, invoice date, RC, invoice no.), then per VIEW_Inv_Posting row of the same invoice and register date a debit line `60` + group company code (payor DC) or `60` + payor (bordx no.); register period YYMM, group company, bordx no. characters 4–8, posted Yes / No, post user / date; debit / credit totals. |
| Post | VIEW_Inv_Posting with the status and the criteria (`cis.InvPost_Rows`) as `6000` D (bordx no.) and `84` + clinic C (clinic name) lines in a workbook; with status 0 - No the cases of each invoice and clinic get CASInvPost 1, the user and today. |
| Print to File | The listed lines as a workbook. |

Fixed desktop bugs: the detail rows are matched on the register day of both rows (the desktop compared the detail's full register date with the day, so a registration with a time — most of them — had no debit line); the register date "to" includes the whole day.

Deliberate differences: parameters instead of pasted criteria; one transaction; the posting workbook is downloaded (the desktop filled the `CISTemplate\InvoicePosting` template, which is not available).

## Payment Posting > Clinic — FrmPostingPay2Cln (`/ac-export/payment/clinic`, right AcExport, `api/ac-export/payment-clinic`)

| Step | Behaviour |
|---|---|
| Search / Search Debit / Search Credit | SUN_CIS_Posting `Pay2Cln` (VIEW_SUN_Posting_Pay2Cln, as `cis.SunPost_Pay2Cln_Search`): entity (Health Connect = payors other than SJ, default; Sihat Johor = SJ), clinic, PV date, PV no., post status, post date, batch no. Per PV: credit `7610` (`7520` for SJ) with the clinic name and debit `84` + clinic (the line Post writes); journal CP, ledger UHC (MS for SJ), amount = paid to clinic rounded to 2 places, analysis `4` + clinic. |
| Post | `Please select the record`, `Record(s) has been posted ! `, `Please Enter PV Date! `, `Please Keyin for Accounting Period ! `; `Do you want to Post record(s) ?` → a SunPostBatch batch no. and a SUNPosting row per ticked PV (TransRef = PV no., InvoiceNo = `4` + clinic, type Pay2Cln) → file `<batch> CP-PayCLN <PV from yyyyMMdd>-<PV to dd>`. |

Kept as the desktop: SUNPosting.InvoiceNo holds the analysis code (`4` + clinic); nothing reads it for clinic payments (the views match on TransRef), so the stored value stays as in the existing postings.

Deliberate differences: the group company filter is not offered (the view has no such column, so the desktop's filter failed); after a post the batch is listed through VIEW_SUN_Posting_Pay2Cln (the desktop switched to VIEW_SUN_Posting_Pay2Cln_Done, which lists no posted rows, so its list came back empty).

## Payment Posting > Member — FrmPostingPay2MBM (`/ac-export/payment/member`, right AcExport, `api/ac-export/payment-member`)

| Step | Behaviour |
|---|---|
| Search / Search Debit / Search Credit | SUN_CIS_Posting `Pay2MBM` (VIEW_SUN_Posting_Pay2MBM, as `cis.SunPost_Pay2Mbm_Search`): group company, PV date, PV no., post status, post date, batch no. Per PV: credit `7610` with the payee (the line Post writes) and debit `82` + group company; journal CP, ledger UHC, analysis 1R; the full PV no. |
| Post | As the clinic payment screen (`Please Enter PV Date! `, `Please Keyin for Accounting Period ! `) → a SunPostBatch batch no. and a SUNPosting row per ticked PV (TransRef = the PV no., InvoiceNo blank, type Pay2MBM) → file `<batch> CP-PayMBM <PV from yyyyMMdd>-<PV to dd>`. |

Fixed desktop bugs: Post stores the full PV no. in SUNPosting.TransRef and the list shows the full PV no. The desktop dropped its first character, which the view never matches, so a posted member PV kept listing as not posted and could be posted again; postings made that way (e.g. batch 2654: `P100463241`) are now recognised as posted too. Search Debit adds its 82 lines to the debit total (the desktop added them to the credit total).

Deliberate differences: parameters instead of pasted criteria; one transaction; the posting workbook is downloaded.

## Bordx Posting > Grp Company (GST) — frmPostingBordxtoGrpGST (`/ac-export/bordx/grp-company-gst`, `api/ac-export/bordx-group-gst`) and Grp Company — frmPostingBordxtoGrpComp (`/ac-export/bordx/grp-company`, hidden in the desktop menu, `api/ac-export/bordx-group`)

One page (route variant `gst`).

| Step | Behaviour |
|---|---|
| Search / Search Debit / Search Credit | SUN_CIS_Posting `BordxGrp` (VIEW_Sun_Posting_BordxGrpComp, as `cis.SunPost_BordxGrp_Search`): group company, bordx no., bordx date (both dates), post date (both dates), post status, batch no.; GST also the entity (Health Connect / Sihat Johor). Per bordereaux a debit line `65` + account (KDC — and on GST KCN — bordereaux: the company code; else the payor's CIS account code, on the hidden screen its characters 3–6; on GST GAS MALAYSIA BERHAD `65A002`; else the payor) and a credit `6000` (the line Post writes); journal BI, ledger UHC, payor (or `1RE` when paid to member). Description: GST `year/month-` + subsidiary (debit) or group company; hidden `YYMM-` + group company from the print month / year (25 characters). The hidden screen adds a credit `1002` for the MCO fee of fee-for-service (F) companies and then credits 6000 with the payable to clinic. Search Debit (group account) / Credit (6000) list actual − co-pay. `Sorry, No Record Found! `. |
| Post | `Please select the record`, `Record(s) has been posted ! `, `Please Enter Bordx Date! `, `Please Keyin for Accounting Period ! `; `Do you want to Post record(s) ?` → a SunPostBatch batch no. and a SUNPosting row per ticked bordereaux (TransRef = bordx no., type BordxGrp) → the posting workbook. |

Fixed desktop bugs: the bordx no. boxes are used as labelled (the desktop's GST screen swapped them); each bordx / post date bound applies on its own (the desktop applied them only when both dates were given).

Deliberate differences: VIEW_Sun_Posting_BordxGrpComp takes minutes (over 4 for 5 days, even filtered by bordx no.), so the procedure picks the bordereaux first (batch, bordx date, bordx no.) and reads the view's two branches for them with the same joins and sums (about 16 s for 5 days; checked against the view for KIS18019); one transaction; the workbook is downloaded.

## Receipt Posting > Grp Company — frmPostingRecGrpCompany (`/ac-export/receipt/grp-company`, right AcExport, `api/ac-export/receipt-group`)

| Step | Behaviour |
|---|---|
| Search / Search Debit / Search Credit | SUN_CIS_Posting `RecGrp` (VIEW_Sun_Posting_RecFrGrpComp, as `cis.SunPost_RecGrp_Search`): entity (Health Connect / Sihat Johor), group company, receipt no., receipt date (both dates), post date (both dates), post status, batch no. Per receipt a debit `7610` and a credit on the company account: `65` + group company code for KDC bordereaux, else `65` + the payor's account code characters 3–6, else `65` + group company code; journal CR, ledger UHC, the receipt amount, the group company, the bordx no.; the credit line is the one Post writes. Search Credit lists `65` + group company, Search Debit `7610`. `Sorry, No Record Found! `. |
| Post | `Please select the record`, `Record(s) has been posted ! `, `Please Enter Receipt Date! `, `Please Keyin for Accounting Period ! `; `Do you want to Post record(s) ?` → a SunPostBatch batch no. and a SUNPosting row per ticked receipt (type RecGrp) → file `<batch> CR-GrpCompany <receipt from yyyyMMdd>-<receipt to dd>`. |

Fixed desktop bugs: the debit is the bank account `7610` (as Search Debit) and the credit the company account (the desktop put both lines on the company account); a receipt is listed and totalled once with all its bordereaux (the desktop listed it once per bordereaux with the whole amount each time); each receipt / post date bound applies on its own.

Deliberate differences: VIEW_Sun_Posting_RecFrGrpComp takes minutes (114 s for 3 rows of two months), so the procedure picks the receipts first (batch, receipt date, receipt no.) and applies the view's joins and grouping to them (about 10 s for two months); one transaction; the workbook is downloaded.

## Not migrated: SUN System GST, SUN System Report, SUN System Allocation (with its account search), SUN System Allocation (New)

frmSUNGSTRpt, frmSUNSystem, frmSUNSystemAllocation / frmSUNSystemSearchAcc and frmSUNSystemAllocation_New read only the SUN accounting database (the
`SALFLDG<UHC|FW|MED|GV|TL|UHI|MS>` ledgers, `SSRFACC` accounts and the `VIEW_GST_*` views) through the `medixmasterSUNSys` connection. That database is not among
the restored GENESIS databases, and the desktop does not open the connection either (the `medixmasterSUNSys.Open` line in SubMain.bas is commented out), so these
screens cannot run in the desktop today. Decision (12 Nov 2026): not migrated; the menu items show "Nothing to build" (desktop state `unavailable`). They can be
built against a configurable SUN connection once that database is available.
