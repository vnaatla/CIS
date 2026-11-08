# LG & Audit — manual test guide (local data)

Log in as DEVADMIN / Gen12345. Clean-up SQL is given after the steps that write.

## 1. Audit review — `/lg-audit/audit-review` (frmAuditInvoice, right 119)

| # | Action | Expected |
|---|---|---|
| A1 | Case `EA12586874` → Go | T301 - KLINIK TAN (TAMAN SETIA INDAH), member BSHA0449584*15-00, LATIFAH BTE HASSAN, diagnoses 202 / 101 |
| A2 | Case `ZZNOCASE1` → Go | `Not Record Found ! ` |
| A3 | Next / Prev | the following / preceding clinical case |
| A4 | Query blank → Update → Yes | `Please Enter Query! ` |
| A5 | Query + Audit By + Audit Date 31/02/2026 → Update | `Invalid date format!` |
| A6 | (writes) Query `TEST`, Audit By `DEVADMIN`, Audit Date today → Update → Yes | `Audit Details updated ! ` |
| A7 | Documents tab (Files:Folders:Outpatient configured) → choose a file → Copy | `File Transfered Successful...`; the file is listed |

Clean-up after A6:

```sql
UPDATE CISDB.dbo.CISCas SET AuditQuery = NULL, AuditAnswer = NULL, AuditDiscountReceived = NULL, AuditRemarks = NULL, AuditBy = NULL, AuditDate = NULL
WHERE CASNumber = 'EA12586874';
```

After A7: delete the copied file from `<Outpatient>\Guarantee Letter\EA12586874`.

## 2. Audit Tracking — `/lg-audit/audit-tracking` (FrmAuditInvoiceTracking, right 119; read-only)

| # | Action | Expected |
|---|---|---|
| T1 | Print to File before any search | `Please search again.!` |
| T2 | Case No From / To `EA12586874` → Search | 6 rows (LORATADINE 10MG TAB … ECONAZINE 20MG CREAM); the case columns (IC 710121016006, LATIFAH BTE HASSAN, T301, BANK SIMPANAN NASIONAL, DIABETES / URTI/SORE THROAT) only on the first |
| T3 | T2 + Pay To `M - Member` → Search | `No record found!` |
| T4 | T2 + OMNI Ref No From `A` → Search | `Invalid column name 'OMNIRefNo'.` |
| T5 | T2 → Print to File | InvoiceAuditTracking.xlsx with the 6 rows in columns A–Y |
| T6 | Group Company `BANK SIMPANAN NASIONAL` | Department list includes JABATAN OPERASI PERBANKAN |

## 3. Generate LG — `/lg-audit/generate-lg` (frmLGGeneration, right 119)

The restored members' policies ended in 2021, so Add / Edit / LG stop at the expiry checks on local data; the issue and decline writes are covered by the rolled-back API tests.

| # | Action | Expected |
|---|---|---|
| G1 | Search with a blank Membership No | `Please key the Membership No` |
| G2 | `UBHA0511637*05` → Search | covers 00–03 (Patricia Ooi Puay Hua …) |
| G3 | Click cover 00 | Annual / Balance FULL, LG Issued Y, LGs EA12590411, EA12590413 … |
| G4 | Add | `Letter of Guarantee is Not Allowed !!` (policy expired 31/12/2021) |
| G5 | Click LG EA12590411 | clinic SUNWAY MEDICAL CENTRE[S427], DERMATOLOGIST, 500.00, Normal GL, UOB template remarks |
| G6 | Edit | `THE ABOVE MEMBER EXPIRY..!!` |
| G7 | Decline without a reason | `Please Select Decline Reason`; with a reason: `Are you sure you want to decline [EA12590411]?` → answer No |
| G8 | Preview | PreviewEA12590411.pdf with the LG's fields |
| G9 | GP Cases / Specialist Cases tabs | the member's cases |

If an LG is issued or declined on a test copy, restore it with (values of EA12590411 before the test):

```sql
UPDATE CISDB.dbo.tblRecord SET IsLGissued = 1, IsRejected = 0, CaseProgress = 'LG Issued' WHERE CaseNo = 'EA12590411';
```
