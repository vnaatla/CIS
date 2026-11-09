# Payment screens (Phase 7)

## To Clinic — frmPaymentClinic (`/payment/clinic`, right PaymentClinic, `api/payment/clinic`)

| Step | Behaviour |
|---|---|
| Search | Clinic-payable (`INVPayTo = 'C'`), open cases (ClaimBordxModule case 26, copied as `cis.PayClinic_Search`). Filters: group company, clinic, payment 0 All / 1 Paid / 2 Unpaid, entity, scan one worksheet no. (`Please Enter Worksheet No`, `Data has been uploaded!`) or a range, append records, invoice no., invoice date range, OP batch range, invoice amount range, PV no. range, PV date range, A/C month / year ranges, SUN invoice batch range, and the "empty" boxes (invoice no., invoice date, OP batch, invoice amount, PV no., PV date, A/C month, A/C year); `No record found!`. Rows: status Full / Partial, payable (payable to clinic + non-GP fee), paid, balance (partial: payable to clinic − paid), and the PV, cheque, bank, period, SUN and claim fields. |
| Upload | A file of invoice numbers (first column, from row 2) lists each invoice's cases. |
| Tick | A fully paid case: `Payment has been made!`; the amount to pay is the row's balance. |
| Save (new PV) | `Invalid PV Date!`, `Please Enter Cheque Amt!`, cheque amount ≠ amount to pay without a remark: `Amount to paid is incorrect` / `Please key in remark for reason.`, `Please Enter Bank !`, `Please keyin Account Period Month ! `, `... Year ! `, bank without a PV code: `Please keyin Bank - PV Code ! `, nothing ticked: `Please select the record`; `Do you want to update the record ? ` → PV no. = bank's PV code + the next PVSeq2013 number; per ticked case the Payment2CLN row (a partial payment gets a new row and its old row the new status) with status F → `Record Saved...` and the PV's rows are listed. |
| Edit / Update | Clicking a paid row fills the PV frame; Edit → Update (`Please select the record`, `Please Confirm PV No!`, the same date / amount / bank / period checks; `Do you want to update the record?`) writes the frame's PV details on every row of the PV. |
| Print PV / Preview PV | `Please Enter or Generate PV No` / `Please select the record`; the payment voucher (Excel): clinic block (preview: the HQ clinic when the clinic has a group code), PV no., cheque no., PV date, prepared by, and per case bill no. / date, case no., receipt no. (REceiptFrPAY), bill amount, MCO fee and amount paid (preview: bordx no. and payable). |
| Print | The list as CSV. |

Known desktop issues (reproduced): the A/C year "from" filter only applies when "to" is also chosen; amounts are compared as text; the partial balance leaves out the non-GP fee.

Deliberate differences: parameters instead of pasted criteria; PV number, Payment2CLN rows and statuses are saved in one transaction with the sequence row locked (the desktop had none, so two users could get the same PV no.); the amounts to pay are recalculated on the server from the listed balance; the PV template (`PV-Clinic.xlt`) is not available, so the voucher is a generated workbook with the same fields; the upload takes a CSV / text file (the desktop opened the Excel file through Excel); double-click to open Invoice Adjustment is not provided.

## To Clinic - Invoice — frmPaymentCLNInvoice, frmInvoiceDetPopup, frmPrintPVRange, frmPrintChequeRange (`/payment/clinic-invoice`, right PaymentClinicInvoice, `api/payment/clinic-invoice`)

| Step | Behaviour |
|---|---|
| Search | One row per clinic / invoice month / year / PV (CISPayment SP_CISPayment `PaymentCLNV2`, copied as `cis.PayClnInv_Search`; same rows as the legacy procedure). Filters: clinic, invoice year and month ranges, payment All / Paid / Unpaid, entity, PV no. and PV date ranges, A/C month / year ranges, the "empty" boxes, append records; `No record found!`. Members EM / ES, non clinic-payable and zero payable are left out. Each row shows the clinic's payee, payee address, bank account, type, FFS, reg. no. and MCO % (NA for Sihat Johor). |
| Invoices | The selected row's invoices (`Pay2CLN`; frmInvoiceDetPopup). |
| Tick | `Please Check Payee! `, `Please Check Payee's Address 1 ! `, the same clinic and month twice: `Please Check Duplicate Records! `, a row already paid: `Payment has been made!`; the amount is the balance. |
| Save | `Please select the record`, `Please Enter Cheque Number!`, `Do you want to update the record ? ` → per run of ticked rows of the same clinic: a PV no. (bank's PV code + PVSeq2013; Sihat Johor: `CS` + the SJ sequence), cheque no. (the first as typed, then the first 3 characters + the next 6-digit number), PV and cheque date today, cheque amount = the clinic's total, the amount in words (CurrencyConversion.bas, split over two lines past 55 characters), and a Payment2CLN row (status F, paid = payable to clinic) for each unpaid case of the clinic, year and ticked months → `Record Saved...` and the new PVs are listed. |
| Edit / Update | Clicking a row fills the PV frame; Edit (nothing ticked) → Update: `Please Enter Pay to! `, `Please Confirm PV No!`, `Invalid PV Date!`, `Please Enter Cheque Number!`, `Please Enter Cheque Date!`, `Please Enter Cheque Amt!`, `Please Enter Wording Amount!`, `Please Enter Bank !`, A/C period month / year; `Do you want to update the record?` → every paid row of the PV gets the frame's details and the amount in words. |
| Print PV / Preview PV | The voucher: payee, clinic (code), payee address, PV no., cheque no., PV date, prepared by, and per invoice no., date, approved, MCO fee, amount paid (preview: the ticked rows' unpaid invoices with the payable). |
| PV range | `Please Enter PV No`; the PVs of the range (`No record found`); Print: one voucher sheet per PV. |
| Cheques | Print Cheque (the PV's cheque) and the cheque range: `Please Enter Cheque No`, `Please ensure cheque range entered correctly`, `No record found`, `Cheque range not in series`; one sheet per cheque (payee, the two wording lines, amount, DDMMYY date) and the cheques are marked printed (PVChequeStatus Y). |

Known desktop issues (reproduced): a clinic's PV takes the year of its last ticked row for every ticked month; the A/C year "from" filter only applies when "to" is chosen; the amount in words keeps the desktop's spacing ("Two Hundred  And Fifty").

Deliberate differences: parameters instead of pasted criteria; all PVs of a save are in one transaction with the sequence locked; the clinic's cheque amount is the total of the cases paid (the desktop used the ticked rows' payable, which can include cases already on another PV); Print PV / Preview PV work (the desktop read fields its list query does not return and always failed with "Export To Excel failed"); the PV (`PV-Clinic.xlt`) and cheque (`PaymentCheque.xlt`) templates are not available, so they are generated workbooks to print on the pre-printed forms; the cheque range checks still run in order but stop at the first failing one.

## FFS Payment — FrmPaymentFFS (`/payment/ffs`, right PaymentFfs, `api/payment/ffs`)

| Step | Behaviour |
|---|---|
| Search | One bordx no. (the default; `Please Enter Bordx Ref. No. `, `Data has been uploaded!`, appended) or a range, group company, payment All / Paid / Unpaid, bordx date, PV no., PV date, A/C month / year ranges and the "empty" boxes (CISPayment SP_CISPayment `PayFFS` / View_PaymentFFS); `No Record Found ! `. Rows: status, bordx no. / date / amount, FFS (MCO fee), paid, balance, group company, PV / cheque / bank / period / payee / receipt fields; total FFS. |
| Tick / Check All | A paid bordereaux: `Payment has been made!`; the amount is the balance. Check All toggles every row (`Please select a record.`). |
| Add | `Please checked at least 1 Claim  ! `; clears the PV frame and puts the ticked total in the cheque amount (payee HEALTH CONNECT SDN BHD, bank MB - MAYBANK). |
| Save | `Please select the record`, `Invalid PV Date!`, `Please Enter Cheque Amt!`, cheque ≠ total without a remark (`Amount to paid is incorrect` / `Please key in remark for reason.`), `Please Enter Bank !`, A/C month / year, `Please keyin Bank - PV Code ! `, `Please keyin Payee Name ! `; `Do you want to update the record ? ` → PV no. (bank's PV code + PVSeq2013) and per ticked bordereaux a PaymentFFS row (status F, the payee's group company code) → `Record Saved...` and the PV is listed. |
| Edit / Update | Clicking a paid row fills the frame; Update (`Please select the record`, `Please Confirm PV No!`, the date / amount / bank / period checks) overwrites the row. |
| Print PV / Preview PV | The FFS voucher: payee and address, PV no., cheque no. and date, prepared by; per bordereaux no., date and amount paid (preview: the FFS balance). |
| Print to File | The list as CSV. |

Known desktop issues (reproduced): Update writes the cheque amount as the row's paid amount.

Deliberate differences: parameters instead of pasted criteria; the amounts are re-read on the server; the preview shows each bordereaux's FFS (the desktop left the amount column empty); the PV template (`PV-FFS.xlt`) is not available, so the voucher is a generated workbook.

## To Member (old) — frmPaymentMember (`/payment/member-legacy`, right PaymentMember, `api/payment/member-legacy`)

The desktop menu never opens this form (it was replaced by frmPaymentMemberAuto08); it is migrated on request.

| Step | Behaviour |
|---|---|
| Search | View_Payment2MBM (as `cis.PayMbmOld_Search`): scan one worksheet no. (`Please Enter Worksheet No`, `Data has been uploaded!`; appended) or a range, member no., member name, group company, payment 0–5, bordx status (default 1 - Bordx), PV no. (each box an "equals" filter), PV date, cheque no., CR no., OP batch, A/C month / year ranges and the "empty" boxes; `No record found!`. Rows as To Member, with the payee's account no. |
| Tick / Check All / Add | As To Member (`Payment has been made!`; the amount to pay is the balance, or payable − postage); Add (`Please checked at least 1 Claim  ! `) clears the PV frame. |
| Cheque amount | Fills the amount in words. |
| Save | The user enters the PV frame: `Please select the record`, `Please Enter Pay To!`, `Please Enter PV Date!`, `Invalid PV Date!`, `Please Enter Cheque Number!`, `Please Enter Cheque Date!`, `Please Enter Cheque Amt!`, cheque amount ≠ selected without a remark (`Amount to paid is incorrect` / `Please key in remark for reason.`), `Please Enter Wording Amount!`, `Please Enter Bank !`, A/C month / year, `Please keyin Bank - PV Code ! `; `Do you want to update the record?` → a PV no. and per ticked case a Payment2MBM row with the frame's details (a partial case's row first gets the new status, then a new row is added); the wording is split at 55 characters into dollars / cents → the PV is listed → `Record Saved...`. |
| Edit / Update | Clicking a paid row fills the frame; Update (`Please select the record`, `Please Enter Pay to! `, the cheque / bank / period / wording checks) overwrites that row (keeping its status and paid amount) and the PV's cheque details → `Record Updated ! `. |
| Print PV / Preview PV / Print Cheque | The PV of the frame's PV no. (`Please Enter or Generate PV No`) with each case's patient; the preview and the cheque as To Member. |
| Print | The list as CSV. |

Known desktop issues (reproduced): the two PV no. boxes each filter on "equals"; the A/C year "from" filter only applies when "to" is chosen.

Deliberate differences: parameters instead of pasted criteria; Save with nothing ticked is refused before a PV no. is taken (the desktop took a PV no. and saved nothing); amounts re-read on the server; one transaction.

## Pay Clinic Post — frmPayClinicPost (`/payment/post/clinic`, right AcExport, `api/payment/post/clinic`)

The desktop menu never opens this form; it is migrated on request.

| Step | Behaviour |
|---|---|
| Search | CISPayment VIEW_PayClinic_Posting (as `cis.PayClnPost_Search`): payor (first 2 characters), invoice no. range, register date range (`Invalid date format!`, `Invoice Date cannot greater than system date ! `), post status (default 0 - No). Two posting lines per payment: debit `84` + clinic code (PV date period, amount, D, PV date, CP, PV no., invoice no., clinic, payor, register period YYMM, group company code, bordx no. characters 4–8, posted Yes / No, post user / date) and credit `7600` (C, invoice no., cheque no. - clinic name); debit and credit totals. |
| Post | The listed posting lines as a workbook; the cases of each listed invoice and clinic not yet posted get CASPayCLNPost 1, the user and today. |
| Print to File | The list as CSV. |

Deliberate differences: the desktop's Post read VIEW_PayClinic_Posting through the CISWorksheet connection (the view exists only in CISPayment), filtered on an invoice-posting column the view does not have, wrote the invoice-posting layout and set the invoice-posting flag, so it could never run. Post here exports the lines this screen lists and sets the clinic-payment posting flag the screen searches on; the parameters replace the pasted criteria; one transaction.

## Not migrated

- **To Member - Batch** — the menu item has no handler (nothing opens).
- **Pay Member Post (frmPayMBMPost)** — nothing to build (decision 9 Nov 2026): the form has a layout like Pay Clinic Post (payor, invoice no., register date, post status, list, Search / Post / Print to File) but no code behind it except the action log on opening, and member payments have no posting flag of their own. Member payment posting is part of A/C Export (VIEW_SUN_Posting_Pay2MBM, Phase 9). The menu item shows "Nothing to build" (desktop state `empty`).

## Tracking / Tracking (Jupiter) — FrmPayment, frmPayment2 (`/payment/tracking`, `/payment/tracking-jupiter`, rights PaymentTracking / PaymentTrackingJupiter, `api/payment/tracking`)

| Step | Behaviour |
|---|---|
| Search | `Please select a report type.` (Pay to Clinic / Pay to Member); selection By Claim, By Bordx, By PV No, By Clinic, By Group Company, General Listing (CISPayment PaymentReceiptModule cases 4–9 clinic / 10–15 member, as `cis.PayTracking_Search`). Filters: claim no., group company, member no., invoice no. and date ranges, bordx, OP batch, clinic (first 4 characters), payment All / Paid / Partial Paid, PV no., PV date, amount (payable), A/C month / year ranges; `No record found!`. The general listings show the status Full / Partial; By Bordx shows `Empty` for cases without a bordereaux; total paid. |
| View / Back | The list on the whole page. |
| Export | The list as CSV (`Report cannot be printed without any record.`). |

frmPayment2 is the same screen on the desktop's Jupiter ("Neptune") connections; the web app reads the one set of databases it is configured with.

Deliberate differences: parameters instead of pasted criteria (the amounts were pasted as numbers); the columns are those of the selected listing (the desktop's two header layouts).

## To Member — frmPaymentMemberAuto08 (with frmPrintPVRange / frmPrintChequeRange in member mode) (`/payment/member`, right PaymentMember, `api/payment/member`)

| Step | Behaviour |
|---|---|
| Search | Member-payable (`INVPayTo = 'M'`), open, claimable cases (CISPayment SP_CISPayment_TESTING `Pay2MBM`, copied as `cis.PayMbm_Search`; same rows as the legacy procedure). Filters: scan one worksheet no. (`Please Enter Worksheet No`, `Data has been uploaded!`; appended) or a range, member no., member name (contains), group company, payment 0–6 (All, Paid, Unpaid, Partial, Paid & Partial, Unpaid & Partial, Paid & Payment Void; except 6 the VOID payments are left out), bordx status (default 1 - Bordx) and range, BTBG bordx status and range, PV no., PV date, cheque no., CR no., OP batch, A/C month / year ranges and the "empty" boxes; `No record found!`. Rows: CR no., status, patient (member or covered person), payable, paid (partial: all the case's payments), balance, postage, PV / cheque / bank / period fields, bordx Yes / No, wording, member, group company, subsidiary, category, channel, treatment, BTBG fields. |
| Tick | Business channel starting with B: `Please check Payment towards borker`; a case already paid: `Payment has been made!`; the amount to pay is the balance (no payment yet: payable − paid − postage fee). |
| Save | `Please select record`, `Please Enter Cheque Number!` (not for VOID), `Please Enter Bank !`, `Please keyin Bank - PV Code ! `; `Do you want to update the record?` → one PV for all ticked cases (bank's PV code + PVSeq2013, or `VOID`): per case a Payment2MBM row (Savecispayment2mbm; a partial payment's row only gets the new status), PV and cheque date today, pay to = the member; then (not VOID) the PV's cheque no., cheque amount (total payable), pay to and the amount in words split into two lines at 50 characters → `Record Saved...` and the PV is listed. |
| Edit / Update | Clicking a paid row fills the PV frame; Edit (nothing ticked) → Update: `Please select the record`, `Please Enter Pay to! `, `Please Enter Cheque Number!`, `Please Enter Cheque Date!`, `Please Enter Cheque Amt!`, `Please Enter Bank !`, A/C month / year, `Please Enter Wording Amount! `; every row of the PV (or the case's VOID rows) gets the frame's details and the wording → `Record Updated ! `. |
| Preview PV / PV range | The voucher (pay to, member address, PV no., cheque no., PV date, prepared by; per case close date, bordx, payee, payable, postage, amount); the PV range lists the PVs (`Please Enter PV No`, `No record found`) and prints one sheet per PV. |
| Cheques | Print Cheque (the PV's cheque) and the cheque range (`No record found`): payee (IBM Malaysia: payee and account no.), the two wording lines, amount, DDMMYY date; marked printed (PVChequeStatus Y). Print Acc No: account no. and payee only (not marked). |

Known desktop issues (reproduced): a partial payment's next instalment only changes the status of the existing row (no new row is saved); the cheque amount is the total payable of the ticked cases, not the amount to pay; the A/C year "from" filter only applies when "to" is chosen.

Deliberate differences: parameters instead of pasted criteria; the PV no., payment rows and cheque details are saved in one transaction (the desktop committed each row separately); the amounts and statuses are re-read on the server; an empty BTBG bordx status is "not chosen" (the desktop compared an empty text with 0, a type mismatch that stopped the search); the PV (`PV-ClinicMBM.xlt`) and cheque (`PaymentCheque.xlt`) templates are not available, so they are generated workbooks; double-click to open Invoice Adjustment is not provided.
