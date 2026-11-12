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
