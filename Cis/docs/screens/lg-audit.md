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
