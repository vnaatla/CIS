# Payment — manual test guide (local data)

Log in as DEVADMIN / Gen12345. Browser steps for this phase alone: `$env:E2E_STEPS='payment'; node e2e/only.mjs` (CisUI).

## 1. To Clinic — `/payment/clinic` (frmPaymentClinic)

| # | Action | Expected |
|---|---|---|
| C1 | Scan ticked, blank worksheet no → Search | `Please Enter Worksheet No` |
| C2 | Untick Scan, payment 2 - Unpaid, invoice date 01/08/2021 – 05/08/2021 → Search | includes EA12565564 (S981, SIMON YII MEDICAL SPECIALIST CLINIC) |
| C3 | PV no `CP100472166` in PV No From / To, payment All → Search | 21 paid rows (e.g. EA12399500, F097) |
| C4 | C3 → tick a Full row | `Payment has been made!` |
| C5 | C2 → tick EA12565564 → Save with no PV date | `Invalid PV Date!` |
| C6 | C3 → click a row → Print PV | PV_CP100472166.xlsx |
| C7 | (writes) C2 → tick → PV date, cheque amt = amount to pay, bank `MB - MBB`, period 10 / 2021 → Save → Yes | `Record Saved...`, new PV `CP…` listed |

Clean-up after C7 (replace the PV no.; set PVSeq2013 back to the number before the test):

```sql
DELETE CISPayment.dbo.Payment2CLN WHERE PVCLNNo = 'CP100472168';
UPDATE CISPayment.dbo.PVSeq2013 SET PVSEQNO = '100472167' WHERE BankCode = 'MB';
```
