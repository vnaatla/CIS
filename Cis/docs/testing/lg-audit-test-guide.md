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
