# Receipt — manual test guide (local data)

Log in as DEVADMIN / Gen12345. Browser steps for this phase alone: `$env:E2E_STEPS='receipt'; node e2e/only.mjs` (CisUI).

## 1. Fr Member — `/receipt/member` (FrmReceiptMember)

The restored data has no member debit notes (VIEW_CISReceiptFrMBM is empty), so only the checks can be tried locally.

| # | Action | Expected |
|---|---|---|
| R1 | Scan, blank worksheet → Search | `Please Enter Worksheet No` |
| R2 | Untick Scan, payment 2 - Unpaid → Search | `No record found!` |
| R3 | Save with no receipt no | `Please Enter Receipt No` |
