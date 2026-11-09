# Claim Bordx — manual test guide (local data)

Log in as DEVADMIN / Gen12345. Browser steps for this phase alone: `$env:E2E_STEPS='claim-bordx'; node e2e/only.mjs` (CisUI).

## 1. Insert Worksheet — `/claim-bordx/insert-worksheet` (frmWSBordx)

| # | Action | Expected |
|---|---|---|
| W1 | Search without a bordx type | `Please select Bordx Type!` |
| W2 | ASO-KDC without an ASO option | `Please select Bordx Type whether ASO or BTBG..!` |
| W3 | Scan ticked, group company blank | `Please select a Group Company Name` |
| W4 | CREDIT NOTE-CN, group company `X\|M005` (or pick M005 from the list) → Search | EA12573444 (DCHC8949658) |
| W5 | W4 → Check All | Selected: 1 |
| W6 | (writes) W5 → Bordx → Yes | `Record Saved, Your Bordx No is : KCN…` |

Clean-up after W6 (replace the bordx no.; restore the counter to its value before the test):

```sql
UPDATE CISWorksheet.dbo.Claim SET BordxRefNo = NULL, ClaimBordxStatus = 0, CLMBordxMonth = NULL, CLMBordxYear = NULL WHERE BordxRefNo = 'KCN00000';
UPDATE CISWorksheet.dbo.ClaimInvoice SET INVBordxRefNo = NULL, INVBordxed = 0 WHERE INVBordxRefNo = 'KCN00000';
DELETE CISWorksheet.dbo.Bordx WHERE BordxRefNo = 'KCN00000';
UPDATE CISWorksheet.dbo.CLAIM_BordxSEQ SET CLMBordxNo = '<previous>' WHERE PAYCode = 'CN';
```

## 2. Remove Worksheet — `/claim-bordx/remove-worksheet` (FrmBordxRevise)

| # | Action | Expected |
|---|---|---|
| R1 | Go with a blank bordx no | `Please select a Bordx No ` |
| R2 | `KIV03270` → Go | EA12476058 (NADRAH BINTI ABD HAMID), approved 503.35; bordx date 02/08/2021 |
| R3 | (writes) R2 → tick EA12476058 → Reject → Yes | `Record updated...`; the worksheet is no longer listed |

## 3. ClaimBordxDate — `/claim-bordx/bordx-date` (FrmBordxDate)

| # | Action | Expected |
|---|---|---|
| D1 | `KZZ99999` → Search | `Record not found` |
| D2 | `KIV03270` → Search | KIV03270, 02/08/2021, posting NO |
| D3 | Update with no date | `Please Enter Bordx Date` |
| D4 | (writes) New date 05/08/2021 → Update | `BordxDate Updated Successfully..!!` (restore: set it back to 02/08/2021) |

## 4. Bordx Delivery — `/claim-bordx/delivery/send-to-payor`, `returned`, `send-to-ac`, `send-to-mnrb`

| # | Action | Expected |
|---|---|---|
| B1 | Send to Grp. Company: Scan ticked, no company → Search | `Please select a Group Company` |
| B2 | Group company `DAYANG ENTERPRISE SDN BHD`, bordx date from 01/01/2021, status 0 → Search | 17 bordereaux |
| B3 | B2 → To Grp. Company with nothing ticked | `Please select at least one record to be updated! ` |
| B4 | (writes) tick one → date → To Grp. Company → Yes | `Record updated...`; it disappears from status 0 |
| B5 | Returned: same bordereaux, date earlier than the date sent | `Date Received from Payor cannot be ealier than Date Sent to Payor!` |
| B6 | Send to A/C / Send to MNRB (payor `DC - …`): same search | the bordereaux with their A/C / MNRB dates |

| B7 | Bor Acc Period: same company / date → Search | 17 bordereaux with their A/C periods |
| B8 | B7 → tick one → Update without a month | `Month cannot be Empty.` |

## 5. Monitoring — `/claim-bordx/monitoring/control`, `outstanding`, `bordx-worksheet`

| # | Action | Expected |
|---|---|---|
| M1 | Bordx Control: `DAYANG ENTERPRISE SDN BHD`, bordx date from 01/01/2021 → Search | 17 bordereaux (e.g. KCN00799, 24/02/2021) |
| M2 | Outstanding Bordx: `ALAM TEKNOKRAT SDN BHD` → Search | 4 bordereaux with days and the ≤ 14 / > 14 split |
| M3 | Outstanding Bordx: `DAYANG ENTERPRISE SDN BHD` → Search | `No record found for this Group Company!` |
| M4 | Bordx & Worksheet: `DAYANG ENTERPRISE SDN BHD`, bordx KDC55049–KDC55049 → Search | its worksheets (e.g. EA12409034, 250.00) |

Clean-up after B4: `UPDATE CISWorksheet.dbo.Bordx SET BordxPayToStatus = 0, BordxPAYSentDate = NULL WHERE BordxRefNo = '<no>';`
