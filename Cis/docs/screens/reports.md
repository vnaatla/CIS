# Tracking and report screens (Phase 10)

## Tracking > Clinic — FrmTracking (`/tracking/clinic`, right ClinicTracking, `api/tracking/clinic`)

| Step | Behaviour |
|---|---|
| Combos | Clinic groups (HISDB ClinicGrp, "code - name"), cities (HISMaintenance CTY), states (HISDB STA); BNF type 1 - Standard / 2 - Plan; clinic type G / S; panel Y / N / C; X-ray Y / N; FFS Y / N. |
| Search | HISDB MBMTracking case 182 (`SELECT * FROM CLN`, as `cis.TrkCln_Search`): clinic code, name (contains), group (code prefix), postcode, city, state, contact person, doctor name, MIX clinic (initial), BNF type, received date range (`Invalid Date Format! `), panel, type, X-ray, for ref, location (any address line contains), FFS status. Lists the clinic's code, name, group, initial, type, BNF type, address, postcode, city, state, received date, panel, X-ray, for ref, no. of doctors, phone, fax, email, homepage, bank account / name / branch, account person, 24 hours, remark, location, FFS; `Record not found`. |
| Print | The list as CSV. |

Deliberate differences: parameters instead of pasted criteria (the desktop pasted the typed text into SQL).

## Tracking > Grp. Com. Clinic — FrmTracking2 (`/tracking/grp-company-clinic`, right ClinicTracking, `api/tracking/grp-company-clinic`)

| Step | Behaviour |
|---|---|
| Combos | Group company names (CISDB GRPCompany) and clinic names (HISDB CLN). |
| Search | CISDB CLNTracking (VIEW_GrpCom_Clinic, as `cis.TrkGrpCln_Search`): group code and clinic code (contain), group company, clinic name. Four views: by group company (group code — `Empty` when blank —, group name, clinic code, clinic name); clinics per group company (group code, name, no. of clinics); by clinic (clinic code, name, group code, name); group companies per clinic (clinic code, name, no. of group companies). `No record found!`. |
| Print | The list as CSV (`Report cannot be printed without any record.`). |

Deliberate differences: parameters instead of pasted criteria.

## Claim ASCII — Bintulu, Toshiba, Singer, TIMS, CIMB, OMNI (`/invoice/claim-ascii/{format}`, right ClaimAscii, `api/reports/claim-ascii/{format}`)

One page (route data `format`); Generate downloads the file and shows the visit count and total. Bintulu, Toshiba and Singer ask `Are you sure?` first. Every screen
shows `No record found!` when nothing is selected and `File Generated Sucessfully...` after the download.

| Screen | Form | Data | File |
|---|---|---|---|
| Bintulu | frmClaimDetailsASCII | View_ClaimDetails_ASCII for the bordx range and the clinic (`cis.ClaimAscii_Bintulu`); the clinic list is group company B003's clinics (SearchGrpCLN). `Please Enter Bordx No Fr `, `Please Enter Bordx No To `, `Please Select Clinic Name `. | `<bordx to>(<clinic>)<month><year>`: "BPSB", employee no., patient, relationship (P → E, H / W → U), visit date and time, M, seven zeros, actual, treatment reason (199 when blank), MC days, the visit date when there are MC days, two blanks — "\|"-separated, each line ending in CR. |
| Toshiba | frmClaimASCIIToshiba | ClaimDetails_ASCII (bordx range) | `<bordx from>(M-MEDIX).txt`: fixed width (Medix `l()`: text left-aligned, numbers right-aligned): M-MEDIX, bordx no. and date, employee no., patient (dependants only), NO, visit date, Toshiba ICD and reason (else 52 Medication & Treatment), MC days, bordx amount without the decimal point, P, bordx date twice, posting id from 001001, time, blanks, 000000. |
| Singer | frmClaimASCIISinger | ClaimDetails_ASCII (bordx range) | `CC<yymmdd of the first bordx date>01`: the name as header, CSV lines (1, bordx date, 01, IC, employee no., patient, visit date, actual, clinic code and name) ending in CR, trailer `2,<last bordx date>,01,,,,,<total>`. |
| TIMS | frmClaimASCIITIMS | ClaimDetails_ASCII_TIMS (one bordx) and HISDB OPPlanAmount for the deductible (N/A when none) | `<today>_OUTPATIENTCLAIM.txt`: "~"-separated, 60 fields as the desktop (fees, medicines, balance, TI no. / cover id, med / injection types, reasons, days and quantities, BTBG amount, deductible, plan, marital status, G / S). |
| CIMB | frmClaimASCIICIMB | ClaimDetails_ASCII_CIMB (one bordx), cash type C / R | `<today>CIMBMEDIX - Cashless.txt` or `- Reimbursement.txt`: fixed width as the desktop; trailer = the total actual as 18 digits. |
| OMNI | frmClaimDataOMNI | ClaimDetails_ASCII_OMNI (the procedure ignores the bordx no.: the claims of OMNI companies not yet sent) and the panel clinics not yet sent | A zip with `<today>_OUTPATIENTCLAIM.txt` (header 01\|TPA, a 02 line per claim of an active member with a cover id, trailer 03 with count and totals) and `<today>_OUTPATIENTPROVIDER.txt` (the clinics, which are then marked CLNOMNISent = Y, in one transaction). |

Known desktop issues (reproduced): TIMS writes an empty cover id (the variable is never set); the TIMS and OMNI fields joined with + lose their separator when empty; CIMB cashless lines without a visit date lose that column; the CIMB trailer rounds the total to whole ringgit; the OMNI procedure ignores the bordx no.

Deliberate differences: files are downloaded (the desktop wrote them to the server's ClinicalClaimASCII folder and C:\ClinicalClaimsASCII / C:\MedicaGenClaimASCII); Bintulu / Toshiba / Singer require the "to" bordx no. (the desktop's check compared it with an empty date mask and never fired); TIMS writes a found deductible as text (the desktop added a Currency value to "~", a type mismatch that stopped the export); CIMB requires a cash type (the desktop silently wrote nothing) and shows the actual total (the desktop showed 0); the clinic can also be typed (the restored data has no clinics for B003, so Bintulu's list is empty here). The TIMS and CIMB provider-file code (CLNFile) is never called on the desktop and is not migrated.

## Claim ASCII (Utusan Melayu) — frmClaimASCIIUtusan (`/invoice/claim-ascii/utusan`, same page and API, format `utusan`)

| Step | Behaviour |
|---|---|
| Generate | `Please Enter Bordx No Fr `, `Please Enter Bordx No To `, `Are you sure?` → ClaimDetails_ASCII_Utusan cases 1–13 (as `cis.ClaimAscii_Utusan`: one view per claim type — 01A, 01B, 01C, 02A, 02B, 03A, 03B, 04A, 04B, 09, 10A, 11A, 11B — for the bordx range) in that order into one file `<today yyyyMMdd>_BPOCLMTRH.KZULF`: UMMB, employee no., claim type, visit date, bordx amount twice, FINANCE, clinic, blank, the treatment code for 10A (reason 304 → 03, specialist clinic → 02, else 01), invoice no. followed by ", ", blank, invoice date, two blanks, "Upload From MediExpress", the dependant's relationship (W → I, H → S, else A), bordx amount, the dependant's name, twelve blanks — each line ending in CR. |

Known desktop issues (reproduced): a member's own line has no name field, so its trailing columns shift by one; the file is written (empty) even when no view returns a row.

Fixed desktop bugs: the visit count covered types 01A and 01C only and the total amount was never added up; both now cover every line.

Local data: the restored databases have no Utusan claims, and view 01A (View_ClaimASCII_UtusanDelEO01A) refers to FWHSHISDB.dbo.MBMCoveredPersons, a database that is not restored, so Generate stops with that SQL error here (the desktop fails the same way). The line layout is covered by a unit test.
