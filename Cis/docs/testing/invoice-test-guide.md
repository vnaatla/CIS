# Invoice — developer test guide (Phase 3)

Setup, login and the general checks are in `maintenance-test-guide.md`. Every section lists real local data, the steps, and the exact expected messages.

## 1. Batch Transfer — `/invoice/batch-transfer` (FrmInvoiceBatchNoUpd, right 119)

| # | Action | Expected |
|---|---|---|
| B1 | Batch No. 1 `opy9978`, Tab | upper-cased; 4 cases, `H153/SEPT14 - SYASYA`, 184.00, KLINIK DAN SURGERI HILLTOP |
| B2 | Batch No. 2 `OPY9987`, Tab | 2 cases, `HC2014-09 - SYASYA`, 99.00, KLINIK MESRA BAYU; Final: OPY9987, 6, 283.00 |
| B3 | Transfer | `Different Invoice No. Must be SAME Invoice No. ! ` |
| B4 | Batch No. 2 = `OPY9978`, Transfer | `Same Batch Number ! ` |
| B5 | Batch No. 1 `ZZNOBATCH` | `Invalid OP Batch Number ! `, the field is cleared |
| B6 | Clear Batch No. 2, Transfer | `Please confirm your Batch Number ! ` |
| B7 | Two batches of the same invoice → Transfer → Yes | `Records Transfered Successfully ! `; batch 1's cases now in batch 2 with the transfer remark; batch 2 amount = total |

Clean-up after B7 (note the original values first):
```sql
UPDATE CISDB.dbo.CISCAS SET OPBatchNo = '<batch 1>', BatchTransRemarks = NULL, InvoiceStmtAmt = <old amount 1>
WHERE BatchTransRemarks LIKE 'Transfer By : DEVADMIN <batch 1>-><batch 2>%';
UPDATE CISDB.dbo.CISCAS SET InvoiceStmtAmt = <old amount 2> WHERE OPBatchNo = '<batch 2>';
```
