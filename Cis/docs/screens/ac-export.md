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
