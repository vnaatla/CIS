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

Known desktop issues (reproduced): the bordx month and year are compared separately (e.g. 12/2020 – 01/2021 finds nothing).

Deliberate differences: parameters instead of pasted criteria; the group company filter is not offered (VIEW_SUN_Posting_InvCln has no group company column, so the desktop's filter failed); the ticked invoices are checked to be still unposted on the server; one transaction with the sequence locked; the posting workbook is downloaded (the desktop saved it on `\\Mars\accounts\Account Journal Import\JIS CLN Not Posted\`).

## Invoice Posting > Clinic (GST) — frmPostingInvoiceClnGST (`/ac-export/invoice/clinic-gst`, right AcExport, `api/ac-export/invoice-clinic-gst`)

The clinic posting screen with route variant `gst`.

| Step | Behaviour |
|---|---|
| Search | SUN_CIS_Posting `InvClnGST` (VIEW_SUN_Posting_InvClnGST, as `cis.SunPost_InvClnGST_Search`): clinic, invoice date and no., post status, post date, batch no., SUN invoice no. (CLNInvoiceBatch), posted A/C period. Line date = today, the batch column shows the SUN invoice batch. No FFS fee: `84` + clinic C and `6000` D (payable). With an FFS fee (MCO processing fee): `6000` D (actual − co-pay, the line Post writes), `84` + clinic C (payable), `1002` C (the fee × 94.34 % from 01/04/2015, else the whole fee) and `8830` C (the GST, fee − 1002 amount). |
| Post | As the clinic screen, but `Please Enter Invoice Date! ` (invoice dates instead of bordx dates). Batch no. from SunPostBatch; each invoice gets a SUN invoice batch `T` (FFS fee with GST) or `I` + PostingINVSEQ; SUNPosting rows with type InvClnGST; file `<batch> RC-CLNInv <invoice from yyyyMMdd>-<invoice to dd>`. |

Known desktop issues (reproduced): the debit / credit totals add the payable amount for every line (an FFS invoice counts it three times as credit), not each line's amount.

Deliberate differences: as the clinic screen (no group company filter, the view has no such column).

## Invoice Posting > Clinic (GST-2017) — frmPostingInvClnGST2017 (`/ac-export/invoice/clinic-gst-2017`, right AcExport, `api/ac-export/invoice-clinic-gst-2017`)

The GST screen (route variant `gst2017`) over SUN_CIS_Posting `InvClnGST2017` (VIEW_SUN_Posting_InvClnGST_2017, as `cis.SunPost_InvClnGST2017_Search`). Differences:

| Step | Behaviour |
|---|---|
| Search | Posted / not posted is PostStatus only; the payor column is blank. |
| Post | One SUNPosting row per case of each ticked invoice (CISCAS by invoice no., clinic and invoice date; the case no. as InvoiceNo), type InvClnGST2017. |
| Generate Tax Invoice | For the selected line's SUN invoice batch (`Invoice posting is not done,Generate is not possible` when blank): the clinic (representative, name, address, phone), `FEE FOR SERVICE FOR <batch>`, the FFS fee without GST and the GST, the batch and today → `Tax Invoice Generated Successfully`. |

Deliberate differences: the `HC FFS Tax Invoice.xlt` template is not available, so the tax invoice is a generated workbook downloaded in the browser (the desktop saved it on `\svr-doc\Template\FFSTaxInv\<batch>\`).
