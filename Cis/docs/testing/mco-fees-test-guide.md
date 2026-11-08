# MCO Fees — manual test guide (local data)

Log in as DEVADMIN / Gen12345.

## 1. Fee Listing — `/mco-fees/fee-listing` (frmMCOFeeListNew)

| # | Action | Expected |
|---|---|---|
| F1 | Search with no criteria | `Please Enter MCO Criteria!` |
| F2 | Bordereaux Batch, no date | `Please Enter Bordereaux Date and Batch No!` |
| F3 | Bordereaux Batch 06/08/2021 / 1, by Group Company `D005`, status N → Search | 6 members of D005, Print enabled |
| F4 | Invoice `I07600`, by Group Company `M284` → Search | 5,363 rows, all invoiced, Print disabled |
| F5 | F4 → Report Type Fee Listing → Print Preview | MCOFeeListing.pdf |
| F6 | F4 → Report Type MCO Invoice → Print Preview | MCOInvoice.pdf with invoice I07600 |
| F7 | (writes) F3 → Report Type MCO Invoice → Print | new invoice I0832x; the 6 members show it |

Clean-up after F7 (replace the number; restore the counter):

```sql
UPDATE HISDB.dbo.MBMCLNOthers SET MBMCInvoiceDate = NULL, MBMCInvoiceNo = NULL, MBMCInvoiceBy = NULL, MBMCInvoiced = 'N', MBMCLNBilledMCO = NULL WHERE MBMCInvoiceNo = 'I08329';
UPDATE HISDB.dbo.MBMCoveredPersonsCLN SET MBMCInvoiceDate = NULL, MBMCInvoiceNo = NULL, MBMCInvoiceBy = NULL, MBMCCLNBilledMCO = NULL WHERE MBMCInvoiceNo = 'I08329';
UPDATE CISDB.dbo.CLNMCOINVQ SET MCOInvoice_No = '08328';
```
