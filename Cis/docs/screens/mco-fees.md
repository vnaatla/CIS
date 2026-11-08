# MCO Fees screens (Phase 5)

## Fee Listing — frmMCOFeeListNew (`/mco-fees/fee-listing`, `api/mco-fees/fee-listing`)

| Step | Behaviour |
|---|---|
| Search | MCO criteria Bordereaux Batch (bordx date + batch no) / Bordereaux Range / Invoice; listing version by Payor, Group Company or Department; print status N / Y (not used for an invoice no). Checks in the desktop's order: `Please Enter MCO Criteria!`, `Please Enter Invoice Number!`, `Please Enter Bordereaux Date and Batch No!`, `Please Enter Bordereaux Range!`, `Please Enter Listing Version!`, `Please Enter Payor!`, `Please Enter Group Company!`; rows of View_Clinic_MCO_ListingNew ordered by member (`No record found!`), principal / supplementary counts. Print is disabled when any row already has an invoice no. |
| Print Preview / Print | `Different Invoice Found in Same Batch or Range, Please Search by Invoice No ! `; `Please select Report Type ! `. Fee Listing: PDF sorted by member / cover with the payor or group company line and the criteria line. MCO Invoice: PDF with the company block, invoice no / date (from the first row that has one), rows by bordx date / insurer. Print of a batch without an invoice takes the next number of CLNMCOINVQ (`I` + 5 digits) and stamps the members in one transaction: cover 00 on MBMCLNOthers (invoice date / no / by, MBMCInvoiced = Y, billed MCO), other covers on MBMCoveredPersonsCLN (invoice date / no / by, billed MCO). |

Known desktop issues (reproduced):
- The invoice's company block is always read from GRPCompany by the payor / group code (PrintPAYInfo compares the listing version with "Invoice - by Payor", which this screen never has).
- Covered persons are not marked MBMCInvoiced = 'Y' when invoiced, so they stay in the "not invoiced" listing.

Deliberate differences: parameters instead of the criteria pasted into MCOFeeNew; the reports are PDFs built by the report engine (the Crystal `.rpt` files and the `\\PLUTO` staging copy `MCOInvoicePluto` are not used); the invoice number is taken under a lock inside the save (the desktop's recordset update could give two users the same number).
