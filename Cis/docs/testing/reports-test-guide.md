# Tracking and reports — manual test guide (local data)

Log in as DEVADMIN / Gen12345. Browser steps for this phase alone: `$env:E2E_STEPS='reports'; node e2e/only.mjs` (CisUI).

## 1. Tracking > Clinic — `/tracking/clinic` (FrmTracking)

| # | Action | Expected |
|---|---|---|
| T1 | Clinic name BALA → Search | 8 clinics, e.g. B020 |
| T2 | Clinic code ZZZZ → Search | `Record not found` |
| T3 | Received from 31/02/2021 → Search | `Invalid Date Format! ` |

## 2. Tracking > Grp. Com. Clinic — `/tracking/grp-company-clinic` (FrmTracking2)

| # | Action | Expected |
|---|---|---|
| G1 | View "Group companies per clinic", clinic code B324 → Search | KLINIK BALA with its no. of group companies |
| G2 | View "By clinic", clinic name KLINIK BALA → Search | about 1900 rows |
| G3 | Group code ZZZZZZ → Search | `No record found!` |

## 3. Claim ASCII — `/invoice/claim-ascii/{format}`

| # | Screen | Action | Expected |
|---|---|---|---|
| A1 | Bintulu | Bordx KIV03274 – KIV03274, clinic NP → Generate | file `KIV03274(NP)<month><year>`, 78 visits |
| A2 | Toshiba | Bordx KIV03274 – KIV03274 → Generate | `KIV03274(M-MEDIX).txt`, lines start with `M-MEDIX KIV03274` |
| A3 | Singer | Bordx KIV03274, no "to" | `Please Enter Bordx No To ` |
| A4 | TIMS | Bordx KDC55481 → Generate | 18 visits |
| A5 | CIMB | Bordx KIS01540, Cashless → Generate | 571 visits |
| A6 | OMNI | Generate | zip; the provider file marks the clinics as sent (clean-up: `UPDATE HISDB.dbo.CLN SET CLNOMNISent = NULL WHERE ...` for the listed codes) |

## 4. Invoice > Tracking — `/invoice/tracking` (FrmInvoiceTracking)

| # | Action | Expected |
|---|---|---|
| I1 | Case no. EA12538600 – EA12538699 → Search | about 95 rows with invoice amount / approved / BTBG totals |
| I2 | Visit date 01/07/2021 – 01/07/2021 → Search | about 2,200 rows (about 6 s) |
| I3 | Bordx no. KIS18019 – KIS18019 → Search | 60 rows (about 35 s) |
| I4 | Double-click a row | Invoice Adjustment opens with that case |
| I5 | Visit date from 31/02/2021 | `Invalid Date Format! ` |

## 5. Invoice > Tracking (Jupiter) — `/invoice/tracking-jupiter` (FrmInvoiceTracking2)

| # | Action | Expected |
|---|---|---|
| J1 | Case no. EA12538600 – EA12538699 → Search | about 95 rows with the MS No column |
| J2 | Same, case status O - Open | only open cases |

## 6. Email Notification (Payment) — `/invoice/email/payment` (frmEmailNotification)

| # | Action | Expected |
|---|---|---|
| E1 | PV date 03/08/2021 – 03/08/2021, status N - New → Search | 161 PVs |
| E2 | No status → Search | `Please Enter Status!` |
| E3 | Tick CP100471624, send date 05/08/2021 → Send → Yes | `Emails have been sent out successfully!`; an .eml file in `D:\CIS\LocalRun\MailPickup` (development: nothing is sent); the PV now lists under R - Resend. Clean-up: `UPDATE CISPayment.dbo.Payment2MBM SET PVEmailStatus = NULL, PVEmailDate = NULL WHERE PVNo = 'CP100471624'` |

## 7. Email Notification (Providers) — `/invoice/email/providers` (frmEmailNotificationProvider)

| # | Action | Expected |
|---|---|---|
| P1 | PV date 03/08/2021 – 03/08/2021, status N - New → Search | 830 clinic PVs |
| P2 | Tick CP100471757 (KLINIK DR AKBAR), send date 05/08/2021 → Send → Yes | .eml in `D:\CIS\LocalRun\MailPickup`; clean-up: `UPDATE CISPayment.dbo.Payment2CLN SET PVEmailStatus = NULL WHERE PVCLNNo = 'CP100471757'` |

## 8. Email Notification (Rejection) — `/invoice/email/rejection` (frmEmailRejectedClaim)

| # | Action | Expected |
|---|---|---|
| R1 | Registration date 01/07/2021 – 31/07/2021, status N - New → Search | about 2,000 claims (about 8 s) |
| R2 | Registration 01/07/2021 – 02/07/2021, case EA12514425 → Search | Ooi Chee Leong, United Overseas Bank, "duplicated" |
| R3 | Tick EA12514425 → Send → Yes | .eml in `D:\CIS\LocalRun\MailPickup`; clean-up: `UPDATE CISDB.dbo.CISCas SET CASEmailRejStatus = NULL WHERE CASNumber = 'EA12514425'` |

## 9. Med Providers (OMNI) — `/invoice/claim-ascii/omni-providers` (frmMedProviderOMNI)

| # | Action | Expected |
|---|---|---|
| M1 | Generate | `<today>_OUTPATIENTPROVIDER.txt` with 131 clinics (nothing is updated) |
