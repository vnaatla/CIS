# Invoice screens (Phase 3)

Desktop forms under Invoice (and the Phase 3 claim screens). Case Transfer and General Exclusions (Enquiry) are documented in `membership.md`.

## Batch Transfer — FrmInvoiceBatchNoUpd (`/invoice/batch-transfer`, right 119, `api/invoice/batch-transfer`)

| Item | Behaviour |
|---|---|
| Batch No. 1 / 2 (LostFocus → ShowInfo) | upper-cased; CISDB `SP_SeachCasByOPBatch` (`cis.BatchTransfer_Batch`): no rows → `Invalid OP Batch Number ! ` and the batch is cleared; else No. of Cases (sum), Invoice No. (distinct values joined with ","), Invoice Amt (sum of the amounts, each rounded to cents) and Clinic Name (last row). More than one row → `Invoice No : <invoices>, Different Clinic Found !`. |
| Final | Final Batch No. = batch 2, total cases, total amount (FinalBatchCal). |
| Transfer | `Please confirm your Batch Number ! ` (a batch empty), `Different Invoice No. Must be SAME Invoice No. ! `, `Same Batch Number ! `, `Are you sure?`; one transaction: batch 1's CISCAS rows get OPBatchNo = batch 2 and BatchTransRemarks = `Transfer By : <user> <batch 1>-><batch 2> , yyyymmdd`, then every batch 2 row gets InvoiceStmtAmt = the total → `Records Transfered Successfully ! ` and the form clears. |
| Clear | empties the form. |

Deliberate differences:
- The API reads both batches again before saving (the desktop compared the text boxes ShowInfo filled); the updates are parameterised (the desktop concatenated the batch numbers into SQL).
- After a successful transfer the desktop also showed an empty message box (its error label runs after the commit); not reproduced.
- `ReSequence` (OPBatch sequence renumbering) is never called by the desktop; not migrated.

## Invoice Update — frmUpdateCLNInvoice (`/invoice/update`, right 119, `api/invoice/update`; hidden in the desktop menu)

| Item | Behaviour |
|---|---|
| Search | Filters, each optional: WS No. from / to (CASNumber), Clinic (panel clinics Y / C; the code before " - "), Payor (group company), Invoice No. (BillNo), Without Invoice No., visit dates from / to, invoice dates from / to, invoice amount from / to. Rows from CISWorksheet `VIEW_Invoice_Tracking_2016` (the desktop's ClaimBordxModule case 1, as `cis.InvUpdate_Search`). None → `No record found!`. Invalid dates → `Invalid Date Format! ` on leaving the field and the filter is ignored. |
| List | Case No, Member No, Cov, Patient Name, Visit Date, Clinic, Invoice No, Invoice Date, Entered Date / By, Invoice Amt, FFS, Final Approve, Group Company; check boxes; Check toggles every row (`Please select a record.` when empty). |
| Print | the list as Excel (`Report cannot be printed without any record.`). |
| Update | `Please select the record`, `Invalid Invoice Date!`, `Please Enter Cheque Amt!` (the desktop's text for a missing invoice number), `Do you want to update the record ? `; each checked case: CISCAS InvoiceNo, InvoiceDate, RegisterDate = today, RegisterBy = user, in one transaction → `Record Saved...` and the list shows the cases of the new invoice number and date. |

Deliberate differences: filters are parameters (the desktop built SQL text and passed it to ClaimBordxModule); an amount filter that is not a number is ignored (the desktop failed with an SQL error); the unreachable `Please Enter Invoice Date! ` check is not shown.
