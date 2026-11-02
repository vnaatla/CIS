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

## 2. Invoice Update — `/invoice/update` (frmUpdateCLNInvoice, right 119)

| # | Action | Expected |
|---|---|---|
| U1 | Invoice No. `I00496-ERMA-M813` → Search | cases of clinic M813 incl. EA12582360 (47.00), EA12572229 (91.00), EA12580861 (46.00) |
| U2 | Invoice No. `ZZ-NO-SUCH` → Search | `No record found!` |
| U3 | Update with nothing checked | `Please select the record` |
| U4 | Check a row, Invoice Date empty → Update | `Invalid Invoice Date!` |
| U5 | Invoice Date 31/10/2026, Invoice No. empty → Update | `Please Enter Cheque Amt!` |
| U6 | Invoice No. `ZZTEST-INV`, 31/10/2026 → Update → Yes | `Record Saved...`; list shows the updated case |
| U7 | Print | InvoiceUpdate.xlsx |

Clean-up after U6 (note the case's InvoiceNo / InvoiceDate / RegisterDate / RegisterBy first):
```sql
UPDATE CISDB.dbo.CISCAS SET InvoiceNo = '<old>', InvoiceDate = '<old>', RegisterDate = '<old>', RegisterBy = '<old>' WHERE CASNumber = '<case>';
```
