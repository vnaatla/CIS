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
