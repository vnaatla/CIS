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
