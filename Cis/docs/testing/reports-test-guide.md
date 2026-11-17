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
| A1 | Bintulu (`/invoice/claim-ascii/bintulu`, frmClaimDetailsASCII) | Bordx KIV03274 – KIV03274, clinic NP → Generate | file `KIV03274(NP)<month><year>`, 78 visits |
| A2 | Toshiba (`/invoice/claim-ascii/toshiba`, frmClaimASCIIToshiba) | Bordx KIV03274 – KIV03274 → Generate | `KIV03274(M-MEDIX).txt`, lines start with `M-MEDIX KIV03274` |
| A3 | Singer (`/invoice/claim-ascii/singer`, frmClaimASCIISinger) | Bordx KIV03274, no "to" | `Please Enter Bordx No To ` |
| A4 | TIMS (`/invoice/claim-ascii/tims`, frmClaimASCIITIMS) | Bordx KDC55481 → Generate | 18 visits |
| A5 | CIMB (`/invoice/claim-ascii/cimb`, frmClaimASCIICIMB) | Bordx KIS01540, Cashless → Generate | 571 visits |
| A7 | Utusan Melayu (`/invoice/claim-ascii/utusan`, frmClaimASCIIUtusan) | Bordx KDC55481 – KDC55481 → Generate → Yes | the SQL error for `FWHSHISDB.dbo.MBMCoveredPersons` (view 01A needs a database that is not restored; the desktop fails the same way). The file layout is covered by the unit test |
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

## 10. MBM Tracking — `/membership/tracking` (FrmNewMBMTracking, in progress)

| # | Action | Expected |
|---|---|---|
| T1 | INS Type & MBM Type / By Payor, member type P, status 0 - Active, payor DC → Search | the same rows as `EXEC HISDB.dbo.CISMBMTracking N' AND PAYCode = ''DC'' And MBMStatus = ''A'' ', 113` (about 15 s) |
| T2 | No report type → Search | `Please select a report type.` |
| T3 | A report marked "not migrated yet" → Search | `This report is not migrated yet.` |

## 11. Claim Tracking — `/claim-bordx/claim-tracking/standard` (frmClaimRpt)

| # | Action | Expected |
|---|---|---|
| C1 | Ageband / Summary, payor DC, DOV 01/07/2021 – 01/07/2021 → Search | the same age bands and counts as `EXEC CISWorksheet.dbo.CISClaimTracking N' AND PAYCode = ''DC'' And DateVisit >= ''07/01/2021'' And DateVisit <= ''07/01/2021''', 14` (DC includes EA) |
| C2 | Admission / Monthly Report, DOV 30/06/2021 – 01/07/2021 | `Must key in same Month` |
| C3 | Admission / YTD Report without a year | `Please enter Year` |
| C4 | General Listing / Without Remarks, payor DC, DOV 01/07/2021 | one row per claim, 54 columns |

## 12. Claim Tracking 2 — `/claim-bordx/claim-tracking/v2` (FrmClaimRpt2)

| # | Action | Expected |
|---|---|---|
| T1 | Actual + Claim, payor EA, DOV 01/07/2021 – 01/07/2021 → Search | one row per claim of `CISWorksheet.dbo.VIEW_CLMReport_Actual` for EA on that day (PAYCode not starting with Z), 28 fee columns and the total |
| T2 | Cost Detail, same criteria | the 28 fee categories and Total; the Actual total = the sum of the fee columns of that view |
| T3 | Claim Tracking 2 (Jupiter) (`/claim-bordx/claim-tracking/v2-jupiter`, FrmNewClaimRpt2), Agent Code "A" | only agent code exactly "A" (the main screen: containing "A") |

## 13. MBM Tracking variants — `/membership/tracking-jupiter` (FrmNewMBMTracking2), `/membership/tracking-simple` (frmNewMBMTrackingSimple), `/membership/tracking-jupiter-simple` (frmNewMBMTrackingRepSimple)

| # | Screen | Action | Expected |
|---|---|---|---|
| V1 | Jupiter | Report Type General Listing | the choices offered by FrmNewMBMTracking2: no "Card Sent Date Listing"; no Billing combo effect |
| V2 | Jupiter | General Listing (Details), payor DC, member no. one member | the supplementaries' Eff. / Exp. dates are their payor dates (main screen: cover dates) |
| V3 | Simple | INS Type & MBM Type / By Payor, Status 0 - Active, payor DC → Search | 2 rows (DC N G 4,173 principals; DC N H 110,890), as `EXEC HISDB.dbo.CISMBMTracking_Simple N' AND PAYCode = ''DC'' And MBMStatus = ''A'' ', 113` |
| V4 | Simple | Report Type Head Count, Head Count From 01/07/2021, To 31/07/2021 → Search | July · DC · 1375 · 121 · 1496 (`EXEC HISDB.dbo.ActiveHeadcount_CIS '2021-07-01','2021-07-31'`) |
| V5 | Simple | Head Count → Print | the list as a file (the desktop's Excel export of it could never run) |
| V6 | Jupiter - Simple | open the Report Type list | Age Band … Race only (no States, no Head Count) |

## 14. Claim Tracking variants — `/claim-bordx/claim-tracking/simple` (frmClaimRptSimple), `jupiter` (frmNewClaimRpt), `jupiter-simple` (frmNewClaimRptSimpleJup), `sun` (frmClaimRptSUN)

| # | Screen | Action | Expected |
|---|---|---|---|
| W1 | Simple | open the Report Type list | 20 types: Admission, Ageband, Bordx, Clinic, … Utilization (no Agent No, Clinic Group, Invoice, Verification of Claims) |
| W2 | Simple | Ageband / Summary, payor `EA`, DOV 01/07/2021 – 01/07/2021 → Search | 9 age bands, the first (blank age band) 221 cases, 28,324.30 actual; as `EXEC CISWorksheet.dbo.CISClaimTracking_Simple N' AND PAYCode = ''EA'' And DateVisit >= ''07/01/2021'' And DateVisit <= ''07/01/2021''', 14` |
| W3 | Simple | payor `DC` instead of EA | no DC → DC / EA rewrite on this screen: only DC claims |
| W4 | Simple | General Listing / Without Remarks, payor EA, DOV 01/07/2021 | 42 columns (main screen: 54) |
| W5 | Jupiter | Ageband / Summary, payor DC, DOV 01/07/2021 | the same 9 rows as the main screen (DC includes EA); the Report Type list stops at State (21 types) |
| W6 | Jupiter - Simple | open the Report Type list | 19 types (Simple without Utilization) |
| W7 | SUN System | General Listing / Without Remarks, payor EA, DOV 01/07/2021 → Search | 2,207 rows (case 112) with Clinic Invoice Batch, PV No and PV Date before the BTBG columns; no Verification of Claims type |
