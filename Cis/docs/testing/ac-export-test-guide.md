# A/C Export — manual test guide (local data)

Log in as DEVADMIN / Gen12345. Browser steps for this phase alone: `$env:E2E_STEPS='ac-export'; node e2e/only.mjs` (CisUI).

## 1. Invoice Posting > Clinic — `/ac-export/invoice/clinic` (frmPostingInvoiceCln)

| # | Action | Expected |
|---|---|---|
| C1 | Invoice date 01/07/2021 – 01/07/2021, status 0 - No → Search | 116 invoices, 232 lines (debit = credit), e.g. 0210152590 P320 983.00 as `84P320` C and `6000` D |
| C2 | C1 → Search Debit | only the 6000 lines |
| C3 | C1 → tick 0210152590 → Post without bordx dates | `Please Enter Bordx Date! ` |
| C4 | C3 with bordx dates, no accounting period → Post | `Please Select Accounting Period ! ` |

Clean-up after a real post (replace the batch no.; set PostingINVSEQ back to the number before the test):

```sql
DELETE CISWorksheet.dbo.SUNPosting WHERE BatchNo = 'A817066' AND PostType = 'InvCln';
UPDATE CISWorksheet.dbo.PostingINVSEQ SET SEQNumber = 'A817065';
```

## 2. Invoice Posting > Clinic (GST) — `/ac-export/invoice/clinic-gst` (frmPostingInvoiceClnGST)

| # | Action | Expected |
|---|---|---|
| G1 | Invoice date 01/08/2017 – 05/08/2017, status 0 - No → Search | 891 unposted invoices |
| G2 | Invoice date 04/08/2017, invoice no. 059809 → Search | 6000 D 640.90, 84D153 C 608.86, 1002 C 30.23, 8830 C 1.82 |
| G3 | Tick a line → Post without invoice dates | `Please Enter Invoice Date! ` |

Clean-up after a real post: `DELETE CISWorksheet.dbo.SUNPosting WHERE BatchNo = '<batch>' AND PostType = 'InvClnGST'`, then set SunPostBatch.BatchNo and PostingINVSEQ.SEQNumber back.
