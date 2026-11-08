# MCO Fees screens (Phase 5)

## Fee Listing — frmMCOFeeListNew (`/mco-fees/fee-listing`, `api/mco-fees/fee-listing`)

| Step | Behaviour |
|---|---|
| Search | MCO criteria Bordereaux Batch (bordx date + batch no) / Bordereaux Range / Invoice; listing version by Payor, Group Company or Department; print status N / Y (not used for an invoice no). Checks in the desktop's order: `Please Enter MCO Criteria!`, `Please Enter Invoice Number!`, `Please Enter Bordereaux Date and Batch No!`, `Please Enter Bordereaux Range!`, `Please Enter Listing Version!`, `Please Enter Payor!`, `Please Enter Group Company!`; rows of View_Clinic_MCO_ListingNew ordered by member (`No record found!`), principal / supplementary counts. Print is disabled when any row already has an invoice no. |
| Print Preview / Print | `Different Invoice Found in Same Batch or Range, Please Search by Invoice No ! `; `Please select Report Type ! `. Fee Listing: PDF sorted by member / cover with the payor or group company line and the criteria line. MCO Invoice: PDF with the company block, invoice no / date (from the first row that has one), rows by bordx date / insurer. Print of a batch without an invoice takes the next number of CLNMCOINVQ (`I` + 5 digits) and stamps the members in one transaction: cover 00 on MBMCLNOthers (invoice date / no / by, MBMCInvoiced = Y, billed MCO), other covers on MBMCoveredPersonsCLN (invoice date / no / by, billed MCO). |

Known desktop issues (reproduced):
- The invoice's company block is always read from GRPCompany by the payor / group code (PrintPAYInfo compares the listing version with "Invoice - by Payor", which this screen never has).
- Covered persons are not marked MBMCInvoiced = 'Y' when invoiced, so they stay in the "not invoiced" listing.

Deliberate differences (Fee Listing): parameters instead of the criteria pasted into MCOFeeNew; the reports are PDFs built by the report engine (the Crystal `.rpt` files and the `\\PLUTO` staging copy `MCOInvoicePluto` are not used); the invoice number is taken under a lock inside the save (the desktop's recordset update could give two users the same number).

## Refund Listing — frmMCORefundListing (`/mco-fees/refund-listing`, `api/mco-fees/refund-listing`)

| Step | Behaviour |
|---|---|
| Search | Listing version by Payor / Group Company / Department (the company's departments), Endt. Eff Date or Endt. Bordx Date range, print status N (no credit note) / Y. Checks: `Please Enter Listing Version!`, `Please Enter Endt. Eff Date!` / `Please Enter Endt. Bordx Date!` (To date), `Please Select a Payor!`, `Please Select a Group Company!`, `Please Select a Department!`; `No record found!`. Principals from VIEW_CLN_MBM_Refund_Search; covered persons of family plans (insured type F / H) from VIEW_CLN_Supp_Refund_Search. MCO fee utilised: utilised ≤ 50 (F / H and covered persons) or ≤ 25 → fee − threshold, else the utilised amount. Totals: principals, covered persons, principal refund, fee utilised, covered-person refund, total refund. |
| Print | The listing as Excel (`Report cannot be printed without any record.`). |
| Print CR Note | `Please Enter Listing Version!`, `Please Select Records!`, `Please Select Date Selection!`, `Please Select Payor!`, `Please Select Group Company!`; company block from PAY (by payor) or GRPCompany by name (department version shows the department); status N → next credit note (`C` + 5 digits of CLNMCOCRNoteSeq) stamped on each listed member and all its covers; status Y → the first row's credit note and date; total refund. |

Known desktop issues (reproduced):
- Only the To date is checked (the desktop tests it twice); the From date may be blank.
- The name is shown only for cancelled (`(Cancelled)`) or pending (`(Pending)`) members.
- Covered persons are listed whatever their credit-note status.

Deliberate differences (Refund Listing): parameters instead of the filters pasted into MCOFeeModule; the credit note number and the member stamps are saved in one transaction (the desktop updated row by row without one); the credit note is an Excel built by the report engine (the `Clinical_MCOCreditNote.xlt` template is not available); a principal's credit note stays on the principal's row (the desktop wrote it onto the last covered person's row).

## MCO Invoice — frmMCOInvoiceNew (`/mco-fees/invoice`, `api/mco-fees/invoice`)

| Step | Behaviour |
|---|---|
| Search | Criteria Bordereaux Batch / Bordereaux Range / Invoice No (invoice no switches to Reprint by group company with payor / group disabled); version Invoice - by Payor / Group Company / Department (the company's departments); Print New / Reprint. Checks: `Please Enter Bordereaux Date!`, `Please Enter Batch Number!`, `Please enter valid Bordereaux Range !`, `Please enter Invoice Number!`, `Please Select a Payor!`, `Please Select a Group Company!`, `Please Select a Department!`; `No record found!`. Principals (VIEW_CisMCO_Princ), then the covered persons of family plans (F / H), then every other matching covered person (VIEW_CisMCO_Supp); principal / supplementary counts and the invoice no. found. |
| Print Preview / Print | `Different Invoice Found in Same Batch or Range, Please Search by Invoice No ! `; the invoice (Excel): company block (PAY by payor, else GRPCompany), department for the department version, invoice no / date, one line per insurer (I - INDIVIDUAL, F - FAMILY, G - GROUP INDIVIDUAL, H - GROUP FAMILY) / bordx date / adult-child with principals, covered persons and MCO fee, totals. Print with no invoice yet: a new number from CLNMCOINVQ for a batch (print new) or a range, then every listed member without one is stamped (cover 00: MBMCLNOthers incl. MBMCInvoiced = Y; other covers: MBMCoveredPersonsCLN); afterwards Print is disabled and the option switches to Reprint. |

Known desktop issues (reproduced):
- The Reprint filter (`IS NOT NULL OR <> ''`) also matches members without an invoice no.
- Printing by invoice no or a batch with Reprint stamps the members with the screen's (blank) invoice no when none was found.
- Covered persons are not marked MBMCInvoiced = 'Y'.

Deliberate differences (MCO Invoice): parameters instead of the criteria pasted into MCOFeeNew; numbering and stamps in one transaction with the number taken under a lock; the invoice is an Excel built by the report engine (the `MCOInvoice.xlt` template is not available).
