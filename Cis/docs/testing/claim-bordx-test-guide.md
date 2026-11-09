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
