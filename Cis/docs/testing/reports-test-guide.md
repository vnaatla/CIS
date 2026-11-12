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
