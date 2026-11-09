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

The restored local data has no member debit notes (VIEW_CISReceiptFrMBM and ReceiptFrMBM are empty), so the list and the saves could only be checked for their messages and SQL, not against real rows.
