# Receipt — manual test guide (local data)

Log in as DEVADMIN / Gen12345. Browser steps for this phase alone: `$env:E2E_STEPS='receipt'; node e2e/only.mjs` (CisUI).

## 1. Fr Member — `/receipt/member` (FrmReceiptMember)

The restored data has no member debit notes (VIEW_CISReceiptFrMBM is empty), so only the checks can be tried locally.

| # | Action | Expected |
|---|---|---|
| R1 | Scan, blank worksheet → Search | `Please Enter Worksheet No` |
| R2 | Untick Scan, payment 2 - Unpaid → Search | `No record found!` |
| R3 | Save with no receipt no | `Please Enter Receipt No` |

## 2. Fr Grp Company — `/receipt/payor` (FrmReceiptPayor)

| # | Action | Expected |
|---|---|---|
| P1 | Bordx, Scan, blank → Search | `Please Enter Bordx No` |
| P2 | Bordx `KIT00036` → Search (about 11 s) | ELANCO MALAYSIA SDN BHD, 99.00, Full, receipt RP37929 |
| P3 | P2 → tick | `Record has been paid !` |
| P4 | Worksheet, Scan `EA12523998` → Search | KIT00036, 99.00, RP37929, Full |
| P5 | Save with no receipt no | `Please Enter Receipt No!` |

## 3. Fr FFS — `/receipt/ffs` (FrmReceiptFFS)

| # | Action | Expected |
|---|---|---|
| F1 | Search with a blank bordx no | `Please Enter Bordx No` |
| F2 | `KDC01078` → Search → tick | received 99.30 (MR10245); `FFS has been received!` |
| F3 | `KDC55710` → Search → tick → Add | Receipt Amt 14.20 |
| F4 | F3 → receipt no `MRTEST`, date, amount 1, cheque → Save | `Amount to paid is incorrect` / `Please key in remarks.` |

Clean-up after an FFS save: `DELETE CISPayment.dbo.ReceiptFrFFS WHERE ReceiptNo = 'MRTEST';`

Clean-up after a real group company save (replace the receipt no. and bordx no.; the claims' PVRecPayorStatus and the bordereaux's BordxPAYFRStatus / BordxPAYRecDate must be set back to their values before the test):

```sql
DELETE CISPayment.dbo.ReceiptfrPAY WHERE ReceiptNo = 'RPTEST';
DELETE CISPayment.dbo.ReceiptSummary WHERE ReceiptNo = 'RPTEST';
```
