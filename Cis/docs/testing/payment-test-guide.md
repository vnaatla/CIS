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

## 2. To Clinic - Invoice — `/payment/clinic-invoice` (frmPaymentCLNInvoice)

| # | Action | Expected |
|---|---|---|
| I1 | Clinic `G029`, invoice year 2021–2021, month 05–05 → Search | the unpaid row and PV CP100468765 (cheque amt 7,254.15, wording "Seven Thousand Two Hundred  And Fifty Four" / " And Cents Fifteen Only") |
| I2 | I1 → select a row → Invoices | its invoices (e.g. O210082271, I210098357) |
| I3 | I1 → tick the PV row | `Payment has been made!` |
| I4 | I1 → tick the unpaid row → Save without cheque no | `Please Enter Cheque Number!` |
| I5 | PV range CP100468765 – CP100468765 → Search PV → Print PV Range | the PV's invoices; PV workbook |
| I6 | Cheque range CHQ000010 – CHQ000005 → Search Cheques | `Please ensure cheque range entered correctly` |

## 3. To Member — `/payment/member` (frmPaymentMemberAuto08)

| # | Action | Expected |
|---|---|---|
| M1 | Scan, blank worksheet → Search | `Please Enter Worksheet No` |
| M2 | Scan `SLK004743875` → Search | CHOOK SHENG CAI, payable 88.00, unpaid |
| M3 | Untick Scan, bordx status blank, PV CP100471601 – CP100471603 → Search | 7 paid rows |
| M4 | M3 → tick a row | `Payment has been made!` |
| M5 | M2 → tick → Save without cheque no | `Please Enter Cheque Number!` |
| M6 | PV range CP100471601 – CP100471603 → Search PV → Print PV | one PV sheet per PV |

## 4. Tracking — `/payment/tracking`, `/payment/tracking-jupiter`

| # | Action | Expected |
|---|---|---|
| T1 | Search with no report type | `Please select a report type.` |
| T2 | Pay to Clinic, General Listing, PV CP100472166 – CP100472166 | 21 cases (e.g. EA12401071, F097) |
| T3 | Pay to Clinic, By PV No, same PV | 1 row, 21 cases, 1,323.90 |
| T4 | Pay to Member, By Bordx, PV CP100471601 – CP100471603 | 1 row `Empty`, 7 cases, 1,091.00 |

## 5. FFS Payment — `/payment/ffs`

| # | Action | Expected |
|---|---|---|
| F1 | Search with a blank bordx no | `Please Enter Bordx Ref. No. ` |
| F2 | `KDC55955` → Search | unpaid FFS 215.98 |
| F3 | F2 → tick → Save without PV date | `Invalid PV Date!` |
| F4 | Untick One bordx no., PV CP100175200 – CP100175200 → Search → click a row → Print PV | the PV's bordereaux (e.g. KDC07917 161.85) |

Clean-up after C7 (replace the PV no.; set PVSeq2013 back to the number before the test):

```sql
DELETE CISPayment.dbo.Payment2CLN WHERE PVCLNNo = 'CP100472168';
UPDATE CISPayment.dbo.PVSeq2013 SET PVSEQNO = '100472167' WHERE BankCode = 'MB';
```

## 6. Pay Clinic Post — `/payment/post/clinic` (frmPayClinicPost)

| # | Action | Expected |
|---|---|---|
| C1 | Post status 0 - No, register date 01/08/2021 – 03/08/2021 → Search | 29 unposted clinic payments (two posting lines each: `84` + clinic code debit and the bank credit), including PV CP100470829 |
| C2 | Register date From 31/02/2021 | `Invalid date format!` |
| C3 | Register date From in the future | `Invoice Date cannot greater than system date ! ` |
| C4 | C1 → Print to File | the list as CSV |
| C5 | (writes) C1 → Post | a workbook of the listed lines; their cases get CASPayCLNPost 1 |

Clean-up after C5:

```sql
UPDATE CISDB.dbo.CISCas SET CASPayCLNPost = '0', CASPayCLNPostUser = NULL, CASPayCLNPostDate = NULL
WHERE CASPayCLNPostDate = CAST(GETDATE() AS date) AND CASPayCLNPostUser = '<user>';
```
