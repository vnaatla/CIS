# Tracking and reports — manual test guide (local data)

Log in as DEVADMIN / Gen12345. Browser steps for this phase alone: `$env:E2E_STEPS='reports'; node e2e/only.mjs` (CisUI).

## 1. Tracking > Clinic — `/tracking/clinic` (FrmTracking)

| # | Action | Expected |
|---|---|---|
| T1 | Clinic name BALA → Search | 8 clinics, e.g. B020 |
| T2 | Clinic code ZZZZ → Search | `Record not found` |
| T3 | Received from 31/02/2021 → Search | `Invalid Date Format! ` |
