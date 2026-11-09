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
