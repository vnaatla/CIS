# Tracking and report screens (Phase 10)

## Tracking > Clinic — FrmTracking (`/tracking/clinic`, right ClinicTracking, `api/tracking/clinic`)

| Step | Behaviour |
|---|---|
| Combos | Clinic groups (HISDB ClinicGrp, "code - name"), cities (HISMaintenance CTY), states (HISDB STA); BNF type 1 - Standard / 2 - Plan; clinic type G / S; panel Y / N / C; X-ray Y / N; FFS Y / N. |
| Search | HISDB MBMTracking case 182 (`SELECT * FROM CLN`, as `cis.TrkCln_Search`): clinic code, name (contains), group (code prefix), postcode, city, state, contact person, doctor name, MIX clinic (initial), BNF type, received date range (`Invalid Date Format! `), panel, type, X-ray, for ref, location (any address line contains), FFS status. Lists the clinic's code, name, group, initial, type, BNF type, address, postcode, city, state, received date, panel, X-ray, for ref, no. of doctors, phone, fax, email, homepage, bank account / name / branch, account person, 24 hours, remark, location, FFS; `Record not found`. |
| Print | The list as CSV. |

Deliberate differences: parameters instead of pasted criteria (the desktop pasted the typed text into SQL).
