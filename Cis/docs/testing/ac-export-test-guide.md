# A/C Export — manual test guide (local data)

Log in as DEVADMIN / Gen12345. Browser steps for this phase alone: `$env:E2E_STEPS='ac-export'; node e2e/only.mjs` (CisUI).

## 1. Invoice Posting > Clinic — `/ac-export/invoice/clinic` (frmPostingInvoiceCln)

| # | Action | Expected |
|---|---|---|
| C1 | Invoice date 01/07/2021 – 01/07/2021, status 0 - No → Search | 116 invoices, 232 lines (debit = credit), e.g. 0210152590 P320 983.00 as `84P320` C and `6000` D |
| C2 | C1 → Search Debit | only the 6000 lines |
| C3 | C1 → tick 0210152590 → Post without bordx dates | `Please Enter Bordx Date! ` |
| C4 | C3 with bordx dates, no accounting period → Post | `Please Select Accounting Period ! ` |

Clean-up after a real post (replace the batch no.; set PostingINVSEQ back to the number before the test):

```sql
DELETE CISWorksheet.dbo.SUNPosting WHERE BatchNo = 'A817066' AND PostType = 'InvCln';
UPDATE CISWorksheet.dbo.PostingINVSEQ SET SEQNumber = 'A817065';
```

## 2. Invoice Posting > Clinic (GST) — `/ac-export/invoice/clinic-gst` (frmPostingInvoiceClnGST)

| # | Action | Expected |
|---|---|---|
| G1 | Invoice date 01/08/2017 – 05/08/2017, status 0 - No → Search | 891 unposted invoices |
| G2 | Invoice date 04/08/2017, invoice no. 059809 → Search | 6000 D 640.90, 84D153 C 608.86, 1002 C 30.23, 8830 C 1.82 |
| G3 | Tick a line → Post without invoice dates | `Please Enter Invoice Date! ` |

Clean-up after a real post: `DELETE CISWorksheet.dbo.SUNPosting WHERE BatchNo = '<batch>' AND PostType = 'InvClnGST'`, then set SunPostBatch.BatchNo and PostingINVSEQ.SEQNumber back.

## 3. Invoice Posting > Clinic (GST-2017) — `/ac-export/invoice/clinic-gst-2017` (frmPostingInvClnGST2017)

| # | Action | Expected |
|---|---|---|
| Y1 | Invoice date 01/07/2021 – 01/07/2021 → Search | 105 unposted invoices (payor column blank) |
| Y2 | Status 1 - Yes, batch 2660 – 2660 → Search | the posted invoice of C225 with SUN invoice batch IA817065 |
| Y3 | Y2 → select a line → Generate Tax Invoice | a workbook `IA817065<timestamp>.xlsx` |

## 4. Invoice Posting > Member — `/ac-export/invoice/member` (frmPostingInvoiceMember)

| # | Action | Expected |
|---|---|---|
| M1 | Invoice date 01/07/2021 – 05/07/2021 → Search | 1016 unposted cases (about 15 s) |
| M2 | Case EA12538622 → Search | 6000 D 178.00, 82U063 C 178.00, patient Nur Amilin Binti Man |
| M3 | Status 1 - Yes, batch 2649 – 2649 → Search | the posted batch (9343 cases, about 20 s) |
| M4 | Tick → Post with bordx dates but no period | `Please Keyin for Accounting Period ! ` |

## 5. Invoice Posting — `/ac-export/invoice/posting` (frmInvPosting)

| # | Action | Expected |
|---|---|---|
| P1 | Register date 01/08/2021 – 02/08/2021, status 0 - No → Search | about 1500 lines (about 13 s), e.g. `84A010` C 49.50 |
| P2 | Register date from 01/01/2999 | `Invoice Date cannot greater than system date ! ` |

Clean-up after a real post: `UPDATE CISDB.dbo.CISCas SET CASInvPost = '0', CASInvPostUser = NULL, CASInvPostDate = NULL WHERE CASInvPostDate = '<today>' AND CASInvPostUser = '<user>'`.

## 6. Payment Posting > Clinic — `/ac-export/payment/clinic` (FrmPostingPay2Cln)

| # | Action | Expected |
|---|---|---|
| K1 | PV date 01/07/2021 – 05/07/2021 → Search | 2858 unposted PVs (about 5 s) |
| K2 | PV date 02/07/2021, PV CP100465721 → Search | 7610 C 34.20 and 84B294 D 34.20 (KLINIK BALA) |
| K3 | Status 1 - Yes, batch 2647 – 2647 → Search | the posted batch (3188 PVs) |
| K4 | Tick → Post without PV dates | `Please Enter PV Date! ` |

## 7. Payment Posting > Member — `/ac-export/payment/member` (FrmPostingPay2MBM)

| # | Action | Expected |
|---|---|---|
| R1 | PV date 01/07/2021 – 05/07/2021 → Search | 529 PVs (about 4 s) |
| R2 | PV date 01/07/2021, PV CP100464273 → Search | 7610 C 260.00 (HATIBUL-UMAM BIN BASRI) and 82A014 D 260.00, PV shown as P100464273 |
| R3 | Tick → Post without a period | `Please Keyin for Accounting Period ! ` |

## 8. Bordx Posting > Grp Company (GST) / Grp Company — `/ac-export/bordx/grp-company-gst`, `/ac-export/bordx/grp-company`

| # | Action | Expected |
|---|---|---|
| B1 | GST, bordx date 01/07/2021 – 05/07/2021 → Search | 76 bordereaux (about 16 s) |
| B2 | GST, bordx no. KIS18019 – KIS18019 → Search | 65… D and 6000 C 5,077.34 (ASIAFLEX PRODUCTS SDN. BHD.) |
| B3 | Hidden screen, KIS18019 → Search | 6000 C 4,633.49 and 1002 C 443.85 (fee-for-service company) |
| B4 | Status 1 - Yes, batch 2648 – 2648 → Search | the posted batch (881 bordereaux, about 30 s) |

## 9. Receipt Posting > Grp Company — `/ac-export/receipt/grp-company` (frmPostingRecGrpCompany)

| # | Action | Expected |
|---|---|---|
| E1 | Receipt date 01/06/2021 – 31/07/2021 → Search | 975 rows (about 10 s) |
| E2 | Receipt no. RP37386 – RP37386 → Search | 4 rows (one per bordereaux), 103,547.24, PERBADANAN INSURANS DEPOSIT MALAYSIA |
| E3 | Tick → Post without receipt dates | `Please Enter Receipt Date! ` |
