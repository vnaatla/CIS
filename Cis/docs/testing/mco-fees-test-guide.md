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

## 2. Refund Listing — `/mco-fees/refund-listing` (frmMCORefundListing)

| # | Action | Expected |
|---|---|---|
| R1 | Search without a version | `Please Enter Listing Version!` |
| R2 | by Group Company, no To date | `Please Enter Endt. Eff Date!` |
| R3 | by Group Company `RAYA AIRWAYS SDN BHD`, Endt. Eff Date 01/01/2021–28/07/2021, status N → Search | 34 principals, refund 748.95 |
| R4 | R3 → Print | MCORefundListing.xlsx |
| R5 | (writes) R3 → Print CR Note | credit note C00097 on the 34 members |

Clean-up after R5:

```sql
UPDATE HISDB.dbo.MBMCLNOthers SET MBMCLNCRNote = NULL, MBMCLNCRNoteDate = NULL, MBMCLNCRNoteby = NULL WHERE MBMCLNCRNote = 'C00097';
UPDATE HISDB.dbo.MBMCoveredPersonsCLN SET MBMCCLNCRNote = NULL, MBMCCLNCRNoteDate = NULL, MBMCCLNCRNoteby = NULL WHERE MBMCCLNCRNote = 'C00097';
UPDATE HISMaintenance.dbo.CLNMCOCRNoteSeq SET CLNCRNote = '00096';
```

## 3. MCO Invoice — `/mco-fees/invoice` (frmMCOInvoiceNew)

| # | Action | Expected |
|---|---|---|
| I1 | Bordereaux Batch, no date → Search | `Please Enter Bordereaux Date!` |
| I2 | Batch 06/08/2021 / 1, by Group Company `D048`, Print New → Search | 2 principals (STHB9440145*01, STHB9440246*01), 4 covered persons |
| I3 | I2 → Print Preview | MCOInvoice.xlsx: H - GROUP FAMILY, Child, principals 2, covered 6 |
| I4 | (writes) I2 → Print | new invoice I0832x on the 6 rows |

Clean-up after I4: as after F7 of the Fee Listing (replace the invoice number).

## 4. MCO Invoice GST — `/mco-fees/invoice-gst` (frmMCOFeeGST)

| # | Action | Expected |
|---|---|---|
| G1 | Batch 06/08/2021 / 1, by Group Company, no company → Search | `Please Select a Group Company!` |
| G2 | Batch 06/08/2021 / 1, by Group Company `D005` → Search | 6 members |
| G3 | G2 → Print Preview | MCOInvoiceGST.xlsx with the invoice and DETAILS sheets |
| G4 | (writes) G2 → Print | new invoice on the 6 members (clean up as F7) |
