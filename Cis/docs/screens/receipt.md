# Receipt screens (Phase 8)

## Fr Member — FrmReceiptMember (`/receipt/member`, right ReceiptMember, `api/receipt/member`)

| Step | Behaviour |
|---|---|
| Search | Member debit notes (CISPayment PaymentReceiptModule case 1: VIEW_CISReceiptFrMBM, payable ≤ 0 and a `D…` debit note, as `cis.RcpMbm_Search`). Filters: scan one worksheet no. (`Please Enter Worksheet No`, `Data has been uploaded!`; appended) or a range, member no. / name (contains), payment 0–5 (All, Paid, Unpaid, Partial, Paid & Partial, Unpaid & Partial), receipt no., receipt date range, DN no. range, DN amount range, A/C month / year ranges and the "empty" boxes; `No record found!`. Rows: DN no., status, case, amount (the debit shown positive), received (partial: all the case's receipts), balance, receipt no. / date, A/C period. |
| Tick | A case already received in full: `Payment has been made!`; the amount to receive is the balance (or the whole amount). |
| Save | `Please Enter Receipt No`, `Please Enter Receipt Date`, `Please Enter Amount`, `Please Enter Cheque No`; `Do you want to update the record?` → per ticked case the case's ReceiptfrMBM row (a new one when none) gets the receipt no., date, cheque no., remarks, period, status F and the amount (added to a partial receipt) → `Record updated...`. |
| Edit / Update | Edit, then clicking a row loads its receipt (`Payment hasn't been made!` when it has none); Update: `Please select the record`, `Please Enter Receipt No!`, `Invalid Receipt Date!`, `Please Enter Receipt Date`, `Please Enter Cheque No!`, `Please Enter Cheque Amt!`; `Do you want to update the record?` → the receipt row is overwritten → `Record Updated...`. |
| Export | The list as CSV. |

Known desktop issues (reproduced): the receipt no. filter uses the "from" box for both ends; the group company drop-down is never applied; a partial case's next receipt overwrites the row's receipt no. and date.

Deliberate differences: parameters instead of pasted criteria (the amount range was compared as formatted text); the amounts are re-read on the server; the receipts of one Save are in one transaction; unticking a case really takes it out of the save (the desktop's untick cleared the whole selection).

## Fr Grp Company — FrmReceiptPayor (`/receipt/payor`, right ReceiptPayor, `api/receipt/payor`)

| Step | Behaviour |
|---|---|
| Search | By bordereaux (PaymentReceiptModule case 18) or by worksheet (case 19), as `cis.RcpPay_Search`. Scan one bordx no. / worksheet no. (`Please Enter Bordx No` / `Please Enter Worksheet No`, `Data has been uploaded!`; appended) or ranges, group company (worksheet: its subsidiaries), payor (CISDB PAY) and its insured clients, insured YES (KDC bordereaux) / NO, bordx status 0–5, worksheet status 0–2, receipt no. / date, client PV no., Medix receipt no., A/C month / year ranges; `No record found!` / `No record found !`. Bordereaux rows: payor, bordx no., claim amount, received, balance, status, bordx date, subsidiary, group company, receipt no., client PV, Medix receipt, insured. Worksheet rows add the case and receipt date. |
| Tick | A bordereaux or claim received in full: `Record has been paid !`; a partial bordereaux: `Record has been Partial Paid !` (it takes the claims not yet received). |
| Save | Bordereaux: `Please select at least one record to be updated! ` / `Please select group company! ` (with an OP claim or OP fee amount), `Please Enter Receipt No!`, `Please Enter Receipt Date!`, `Please Enter Cheque No!`, `Please Enter Receipt Amount!`, amount ≠ selected without a remark `Please Enter Remark!`; `Do you want to update the record ?` → per claim of each ticked bordereaux: claim received (F), a ReceiptfrPAY row (receipt amount, the claim's bordx amount as received), bordereaux F with the receipt date; the OP claim / OP fee lines go to ReceiptSummary → `Record Updated...`. Worksheet: the same checks with their own texts → per ticked claim a ReceiptfrPAY row when it has none, the claim F, then the bordereaux F when all its claims are received, else P → `Record updated...`. The receipt's rows are then listed. |
| Edit / Update | Edit, then clicking a row loads its receipt (`Payment hasn't been made!` when none); Update (`Please select the record`, `Please Enter Receipt No!`, `Invalid Receipt Date!`, `Please Enter Cheque No!`, `Please Enter Receipt Amt!`) overwrites the bordereaux's receipt rows (with the period) or the claim's row → `Record Updated...`. |
| Export | The list as CSV. |

Known desktop issues (reproduced): every ReceiptfrPAY row of a save carries the whole receipt amount; Update does not change the bordereaux's receipt date (the desktop never saved it); the worksheet Update does not change the period.

Deliberate differences: parameters instead of pasted criteria; the lists use an extra filter implied by their join so a bordereaux search takes about 11 s (the desktop's procedure 13 s) and a worksheet search 6 s (27 s); one transaction per save (the desktop wrote claim by claim); the selected amount and the statuses are re-read on the server; the claim update is limited to the ticked case (the desktop's `case AND bordx OR BTBG bordx` without brackets could pick another claim); bordx status 5 works (the desktop's filter was invalid SQL); in bordereaux mode clicking a row in edit mode loads its receipt (the desktop's check was inverted and always failed).

## Fr FFS — FrmReceiptFFS (`/receipt/ffs`, right ReceiptFfs, `api/receipt/ffs`)

| Step | Behaviour |
|---|---|
| Search | One bordx no. (the default; `Please Enter Bordx No`, `Data has been uploaded!`, appended) or a range, group company, payment All / Paid / Unpaid, bordx date, receipt no., receipt date and A/C month / year ranges (PaymentReceiptModule case 17: VIEW_CISReceiptFFS, as `cis.RcpFfs_Search`); `No record found!`. Rows: status, bordx no. / date / amount, FFS, received, balance, receipt no. / date / remark / period / cheque, total received, group company; total FFS. |
| Tick / Check All / Add | A received bordereaux: `FFS has been received!`; Check All toggles every row (`Please select a record.`); Add (`Please checked at least 1 Claim  ! `) clears the frame with the checked FFS total. |
| Save | `Please select the record`, `Please Enter Receipt No`, `Invalid Receipt Date!`, `Please Enter Receipt Amt!`, amount ≠ checked total without a remark (`Amount to paid is incorrect` / `Please key in remarks.`), `Please Cheque Number!`, A/C month / year; `Do you want to update the record ? ` → one ReceiptFrFFS row (status F, the bordereaux's FFS) per ticked bordereaux → the receipt's rows are listed → `Record Saved...`. |
| Edit / Update | Clicking a received row fills the frame; Update (same checks) overwrites that row with the frame's amount → `Record Saved...`. |
| Export | The list as CSV. |

Deliberate differences: parameters instead of pasted criteria; the FFS amounts are re-read on the server; listing a receipt's rows reads its bordereaux one at a time (the view sums every claim, so a receipt-no. filter took 25 s; now about 7 s for 9 bordereaux).

## Upload Receipt / Upload Receipt (Multiple Receipts) — frmReceiptUpload, frmReceiptUpMultiple (`/receipt/upload`, `/receipt/upload-multiple`, right ReceiptPayor, `api/receipt/upload`)

| Step | Behaviour |
|---|---|
| Upload | `Please select file to upload!`. Sheet 1 from row 2 (rows with a blank claim no. — column B for multiple — are skipped): B bordx no., C claim no., F receipt amount, G client PV no.; multiple also H receipt no., I Medix receipt no. Each row is compared with its claim: claim / bordx `Matched` or `Not Matched`, the claim's bordx amount and the difference. |
| Receipt date | Fills the A/C period month and year. |
| Update | `Please Upload file..! `; single: `Please Enter Receipt Number! `, `Please Enter Medix Receipt Number! `; `Do you want to update the record ?` → per matched claim a ReceiptfrPAY row (receipt no. and Medix no. from the form, or from the file for multiple; date, amount, cheque and period from the form; the file amount as received; the claim's bordx no.; client PV no.) and the claim received (F) → `Record Updated...`. |
| Print | The list as CSV. |

Known desktop issues (reproduced): the bordx column always shows `Matched` when the claim exists (the desktop compared the CIS bordx no. with itself); the bordereaux's receipt status is not updated.

Deliberate differences: the file is an `.xlsx` read on the server (the desktop opened it in Excel); the claims are matched again on the server when saving; one transaction for the upload (the desktop wrote row by row and could add an extra blank receipt row for a claim without one).

## Not migrated

- **Fr MNRB (FrmReceiptMNRB)** — never opened from the desktop menu.

The restored local data has no member debit notes (VIEW_CISReceiptFrMBM and ReceiptFrMBM are empty), so the list and the saves could only be checked for their messages and SQL, not against real rows.
