# LG & Audit screens (Phase 4)

## Audit review — frmAuditInvoice (`/lg-audit/audit-review`, right 119, `api/lg-audit/audit-review`)

The form file is `frmAuditInvoicej.frm` (VB_Name `frmAuditInvoice`).

| Step | Behaviour |
|---|---|
| Go | SP_CISAuditdetails (`Not Record Found ! `): status (O / C / D), claimability (A / R / C / S), pending, pay to, clinic, invoice no / dates, process month, member-cover, name, IC, department, bill no / amount, visit date / time, diagnoses 1–5 with their TreatmentReason descriptions (diagnosis 5 "OTHERS" shows the typed others diagnosis), preliminary diagnosis, medications 1–7 with their charges, and the audit details. All case fields are read-only. |
| Prev / Next / Search | VIEW_Cas_MBM next / previous clinical case of an active member (`First Record Reached! … Not a valid Case Number! `, ` Last Record Reached! `); case list pop-up. |
| Update | `Not Record Found ! ` (no case loaded); `Are you sure want update Audit Details? `; `Please Enter Query! `; `Please Enter Audit By! `; `Please Enter Audit Date! `; `Invalid date format!`; CISCas AuditQuery, AuditAnswer, AuditDiscountReceived, AuditRemarks, AuditBy, AuditDate → `Audit Details updated ! `. (Same update as the Adjustment screen's audit review.) |
| Documents | the case's folder `Outpatient\Guarantee Letter\<case>` (jpg, tif, gif, doc, pdf, png, msg, xlsx, dot, docx, txt); a file opens by download; Copy puts a chosen file into the folder (`Please Enter Case No` without a case) → `File Transfered Successful...`. |

Known desktop issues: the Update asks "Are you sure…" before checking the fields, and its answer is only used for the final update (reproduced: the question comes first).

Deliberate differences: parameters instead of pasted SQL; `.docx` files are always listed (the desktop matched `.docx` only for the first file of the folder); the upload is a browser file picker instead of the Windows open dialog and file copy.

## Audit Tracking — FrmAuditInvoiceTracking (`/lg-audit/audit-tracking`, right 119, `api/lg-audit/audit-tracking`)

| Step | Behaviour |
|---|---|
| Filters | Case, member, IC, invoice, Bordx, plan, OP batch, OMNI ref and invoice amount ranges; visit, register, invoice, invoice-registered and Bordx date ranges; pay to, clinic, payor, group company (→ its subsidiaries and departments), patient name (exact), case status, claimability, A/C period month / year, clinic type (0 S / 1 G / 2 blank), FFS, OMNI sent, OMNI status, TIMS, audited (AuditQuery not null). Each "blank" check box adds "is null or empty". Invalid dates are ignored, as the desktop's IsDate checks did. |
| Search | Sp_Audittraccking_2018's report (`cis.AuditTrack_Search`): cases of payors other than DZ / EM / ES that are not closed or deleted, one row per medication, with the case columns on the first row only. None → `No record found!`. |
| Print to File | `Please search again.!` before a search; otherwise the InvoiceAuditTrackingNEW columns A–Y (Bordx no, IC, name, case, visit date, diagnoses 1–5, medication / days / qty / charges, procedure fee, consultation fee, invoice amount, query, answer, remarks, discount received, audit by, clinic code / name, group company) as an Excel file. |

Known desktop issues (reproduced):
- The claimability filter adds `Or CISCas.Claimability = …` and the FFS "N" filter `AND CLNFFSStatus = 'N' OR (…)` without brackets, so they widen the search instead of narrowing it.
- Cases without any medication are never listed.
- The procedure always excludes closed and deleted cases, so Case Status C / D finds nothing.
- The OMNI Ref filters use the column `OMNIRefNo`, which none of the report's tables has, so they end in SQL Server's `Invalid column name 'OMNIRefNo'.` (the desktop showed that error).

Deliberate differences: every filter is a parameter (the desktop pasted the filter text into the procedure, an injection risk); the Excel is written by the report engine without the template's merged cells (the case columns are blank on a case's medication rows instead); Stop and View / Back (desktop layout buttons) are not needed in the browser.
