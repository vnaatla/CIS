# Invoice — developer test guide (Phase 3)

Setup, login and the general checks are in `maintenance-test-guide.md`. Every section lists real local data, the steps, and the exact expected messages.

## 1. Batch Transfer — `/invoice/batch-transfer` (FrmInvoiceBatchNoUpd, right 119)

| # | Action | Expected |
|---|---|---|
| B1 | Batch No. 1 `opy9978`, Tab | upper-cased; 4 cases, `H153/SEPT14 - SYASYA`, 184.00, KLINIK DAN SURGERI HILLTOP |
| B2 | Batch No. 2 `OPY9987`, Tab | 2 cases, `HC2014-09 - SYASYA`, 99.00, KLINIK MESRA BAYU; Final: OPY9987, 6, 283.00 |
| B3 | Transfer | `Different Invoice No. Must be SAME Invoice No. ! ` |
| B4 | Batch No. 2 = `OPY9978`, Transfer | `Same Batch Number ! ` |
| B5 | Batch No. 1 `ZZNOBATCH` | `Invalid OP Batch Number ! `, the field is cleared |
| B6 | Clear Batch No. 2, Transfer | `Please confirm your Batch Number ! ` |
| B7 | Two batches of the same invoice → Transfer → Yes | `Records Transfered Successfully ! `; batch 1's cases now in batch 2 with the transfer remark; batch 2 amount = total |

Clean-up after B7 (note the original values first):
```sql
UPDATE CISDB.dbo.CISCAS SET OPBatchNo = '<batch 1>', BatchTransRemarks = NULL, InvoiceStmtAmt = <old amount 1>
WHERE BatchTransRemarks LIKE 'Transfer By : DEVADMIN <batch 1>-><batch 2>%';
UPDATE CISDB.dbo.CISCAS SET InvoiceStmtAmt = <old amount 2> WHERE OPBatchNo = '<batch 2>';
```

## 2. Invoice Update — `/invoice/update` (frmUpdateCLNInvoice, right 119)

| # | Action | Expected |
|---|---|---|
| U1 | Invoice No. `I00496-ERMA-M813` → Search | cases of clinic M813 incl. EA12582360 (47.00), EA12572229 (91.00), EA12580861 (46.00) |
| U2 | Invoice No. `ZZ-NO-SUCH` → Search | `No record found!` |
| U3 | Update with nothing checked | `Please select the record` |
| U4 | Check a row, Invoice Date empty → Update | `Invalid Invoice Date!` |
| U5 | Invoice Date 31/10/2026, Invoice No. empty → Update | `Please Enter Cheque Amt!` |
| U6 | Invoice No. `ZZTEST-INV`, 31/10/2026 → Update → Yes | `Record Saved...`; list shows the updated case |
| U7 | Print | InvoiceUpdate.xlsx |

Clean-up after U6 (note the case's InvoiceNo / InvoiceDate / RegisterDate / RegisterBy first):
```sql
UPDATE CISDB.dbo.CISCAS SET InvoiceNo = '<old>', InvoiceDate = '<old>', RegisterDate = '<old>', RegisterBy = '<old>' WHERE CASNumber = '<case>';
```

## 3. OP Batch/Invoice — `/invoice/op-batch` (frmUpdateInvoiceBatch, right 119)

| # | Action | Expected |
|---|---|---|
| O1 | Clinic `T301 - …`, visit 01/08/2021–31/08/2021 → Search | 3 cases: EA12585463 (OP batch OCA1133582), EA12586874, EA12586891 |
| O2 | Check EA12585463 after giving it an invoice | `Selected Case Already OP Batch And Invoice No generated..! ` |
| O3 | Invoice No. `ZZ9`, clear Visit Date To → Update | `Please select Visit Date.` |
| O4 | Update with nothing checked | `Please select the record` |
| O5 | Check EA12586874 → Update (no invoice no.) → Yes | `Record Saved...`; the case gets a new OCA… batch, OPNoEnterUser, OPBatchDate |
| O6 | Search again, check EA12586891, Invoice No. `ZZ9` → Update → Yes | invoice `ZZ9-T301`, date 31/08/2021, CASInvMonth 08, CASInvYear 2021 |
| O7 | Repeat O6 with `ZZ9` | `Given Invoice No is already used...!` |

Clean-up after O5 / O6:
```sql
UPDATE CISDB.dbo.CISCAS SET OPBatchNo = NULL, OPNoEnterUser = NULL, OPBatchDate = NULL, InvoiceNo = NULL, InvoiceDate = NULL, CASInvMonth = NULL, CASInvYear = NULL
WHERE CASNumber IN ('EA12586874', 'EA12586891');
-- The OPBatch 'OC' sequence is not rolled back (the desktop never reuses numbers either).
```

## 4. MMA/CPT — `/invoice/mma-cpt` (frmMMASchedule, right 150)

| # | Action | Expected |
|---|---|---|
| M1 | Main Body Part `brain`, OPCS `A01` → Search | A0110 (surgical cat `CMO5`, surgeon fee 7030, part B), A0180 (4750) … |
| M2 | Click A0110 | MMA Details: surgical fee 7030, page 4B, part B |
| M3 | OPCS `ZZZZ` → Search | `Record not found` |
| M4 | Surgical Category `SCMO5` | rows whose category contains `CMO5` (the first character is dropped, as in the desktop) |
| M5 | Print | MMASchedule.xlsx |

Read-only screen: nothing to clean up.

## 5. Pass to Claim — `/claim/pass-to-claim` (frmPass2Claim, right 163)

| # | Action | Expected |
|---|---|---|
| P1 | Scan on, Search with an empty case no. | `Please Enter Case No` |
| P2 | Scan `ea12586874` | one row: Not, KLINIK TAN (TAMAN SETIA INDAH) |
| P3 | Scan `EA12586874` again | `Data has been uploaded!` |
| P4 | Scan off, case no. EA1258680–EA1258690 → Search | the cases of that range |
| P5 | Check a case without register date | `Date of Register is Empty.` |
| P6 | Check EA12586874 → Pass → Yes | `Record updated...`; the case shows Passed |

Clean-up after P6: `DELETE FROM CISWorksheet.dbo.ClaimFileTransfer WHERE CASNumber = 'EA12586874' AND ClaimPassBy = 'DEVADMIN' AND ClaimPassDate = '<today>';`

## 6. Update Claims Data — `/claim-bordx/update-claims-data` (frmUpdateClaimsData, right 124)

| # | Action | Expected |
|---|---|---|
| D1 | Verify without a file | `Please select file to upload!` |
| D2 | Workbook row 2: A=`N1234567`, H=12-character code, I=date, J=amount → Verify FFS | `File verified successfully!`, FFS enabled |
| D3 | Same with A=`X1` | `Error encoutered. Please check the file!`; row 2 listed |
| D4 | Copay file: B=case no. (e.g. a test case starting S / N), C=amount, T=date → Verify Copay → Copay → Yes | `Records have been uploaded!`; Claim.CLMCoPayColl updated |

Clean-up after D4: restore the case's CLMCoPayColl / ClaimLastUpdateDate / ClaimLastUpdateUser to the values noted before.

## 7. Clinic invoice — `/invoice/clinic` (frmInvoice, right 117)

| # | Action | Expected |
|---|---|---|
| C1 | Inv. No `ZZTEST`, Inv Date 31/08/2021, Rec Date 01/09/2021, Clinic `t301`, Tot. Amt 45 | clinic KLINIK TAN (TAMAN SETIA INDAH) |
| C2 | Member `DCHC9086608` → Go | NUR RABIZAH BINTI ADENI, policy 01/08/2021–31/12/2021, limits shown |
| C3 | Visit 15/08/2021, GP Fee 30, X-Ray 15, co-pay 0, Bill Amt 45, Clm'lity A → Add | line Bill Amt 45.00; Compute shows the worksheet (actual 30 / 15) |
| C4 | Bill Amt 40 → Add | `Please check Bill Amount! ` |
| C5 | Visit 31/07/2021 → Add | `Visit Date cannot less than Membership Effective date ! ` |
| C6 | Verify | `45 Days LATE ! ` (old data) listed; Save is still enabled (desktop behaviour) |
| C7 | (data within 45 days only) Verify clean → Save → Are you sure ? → Yes | `Record Saved ! `, new case EA…; CISCAS, Claim, ClaimInvoice*, member balance updated |

Clean-up after C7: delete the new case's rows (CISCAS, CISWorksheet Claim / ClaimInvoice / ClaimInvoiceActual* / ClaimCharges) and restore the member's MBMCLNBalLmt / MBMAccumAmt / MBMAdjAnnLmt / MBMCLNMCDaysBal noted before.

## 8. Member invoice — `/invoice/member` (frmInvoiceMBM, right 118)

| # | Action | Expected |
|---|---|---|
| R1 | Header as C1 with a Rec Mode, Clinic `T301`, Member `DCHC9086608` → Go | member panel |
| R2 | Visit 15/08/2021, GP Fee 30, X-Ray 15, co-pay 0, Accepted Y → Add | line 45.00; Verify lists `45 Days LATE ! ` (old data) |
| R3 | Rec Mode blank → Add | `Please Enter Invoice Receipt Mode ! ` |
| R4 | Accepted N without reason → Add | `Please Enter Repudiation Reason` |
| R5 | Clinic `NP` without a name → Add | `Please keyin the NP Clinic ! ` |
| R6 | (data within the policy and 45 days) Verify → Save → Yes | `Record Saved ! ` with case EA…; CISCAS OPBatchNo OC…, ClaimInvoice pay to M |

Clean-up after R6: as C7, plus the member's MBMHealthlimit when the plan has a health limit.

## 9. Invoice enquiry — `/invoice/enquiry` (FrmInvoiceEnq, right 119)

| # | Action | Expected |
|---|---|---|
| E1 | Case `ea12586874` → Go | member DCHC… cover, clinic T301 - KLINIK TAN (TAMAN SETIA INDAH), status Open |
| E2 | Worksheet tab | stored worksheet with Final Approved |
| E3 | Bordx tab | SP_Claim_Bordx row (status 0, final approve 233.00) |
| E4 | Next / Prev | the following / preceding clinical case |
| E5 | Case `ZZNOCASE1` → Go | `Not Record Found ! ` |
| E6 | Documents tab (Files:Folders:Outpatient configured) | the case's scanned files; click downloads |

Read-only screen: nothing to clean up.

## 10. Invoice adjustment — `/invoice/adjustment` (frmInvoiceAdj, right 119)

| # | Action | Expected |
|---|---|---|
| J1 | Case `EA12586874` → Go | T301 KLINIK TAN, total 233.00, status O, accepted |
| J2 | Edit | form enabled; Update available |
| J3 | Bill Amt 200 → Update | `Please check Bill Amount! ` |
| J4 | Pending Y, no reasons → Update | `Case is in Pending State, Please select Pending Reasons...! ` |
| J5 | Compute | worksheet with the stored contracted rates |
| J6 | (writes) Remark changed → Update → Are you sure ? → Yes | `Record updated ! `; CISCAS remark, Claim, CLM*New tables updated |
| J7 | Audit: query + by + date → Update Audit | `Audit Details updated ! ` |

Clean-up after J6 / J7: restore CISCAS (CasRemark, LastUpdateBy / Date, Audit*), the Claim / ClaimInvoice / ClaimInvoiceActualNew / CLM*New rows and the member's MBMCLNBalLmt / MBMAccumAmt / MBMAdjAnnLmt from a backup taken before.


## 11. Invoice adjustment after bordereaux —  (frmInvoiceAdjAftBordx, right 179)

| # | Action | Expected |
|---|---|---|
| K1 | Case  → Go | the case loads; no Pending Reasons / Audit / Void Bordx / Claims App |
| K2 | Edit → Bill Amt 200 → Update |  |
| K3 | (writes) Remark changed → Update → Yes | ; only CISCAS and ClaimInvoice change; CISCAS.AddPendNote becomes NULL |

Clean-up after K3: restore CISCAS (CasRemark, LastUpdateBy / Date, AddPendNote) and ClaimInvoice from a backup taken before.
