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

## Generate LG — frmLGGeneration (`/lg-audit/generate-lg`, right 119, `api/lg-audit/generate-lg`)

| Step | Behaviour |
|---|---|
| Search | `Please key the Membership No`; HISDB CLNLG1 lists the covers (`Record not found!`). For OPENET, SPAN, Communication Corporation, Saint-Gobain, PKNS, DiGi and three named members the member's MBMCLNWarranty text is shown. |
| Cover click | CheckMBMLimit: GetLGAnnLimitSpecLimit labels, then the utilisation recomputed from CIS_FFS2013 by limit status (ASO members: Tbl_ASOMemberDetails), specialist top-ups, status E / C rules, shared plans N/A, MERCHANTRADE M230 balance message. HR approval asks ` All MO/SP GL Request Must be Come From HR.`. ValidateLG (SP_ValidateLG1) sets the status (`Member is not active.`) and whether an LG may be issued. The member's LGs come from LGlst2020. |
| LG click | View_PrintLG, the pending notice (`The case is still in pending state. …`), the LG's alert message, GetLGAnnLimitSpecLimit labels as they are (no recalculation), the GL template and remarks (Sp_Gettemplatedetails_2), pharmacy cases (documents in OtherScannedDoc, actions disabled while a medication request is in progress). |
| Add | The desktop's limit tree (`Letter of Guarantee is Not Allowed !!`, `Limit Exceeded. …`, `Member Entitled Max LG Amount RM … Press OK to Add LG and Cancel to Abort.`, the SCOMI notice, `No LG For this Plan!!!!` …), then only opens the entry fields. |
| Edit | `Please Select the LG !!`; bordereaux / invoice no.; payment; clinic not in panel (spLGforSP_check, then CLN.CLNPanel); `THE ABOVE MEMBER EXPIRY..!!` (expiry older than 30 days); then the limit tree with "Edit LG" wording. |
| LG / Update | Pending question; bordereaux, payment, panel and expiry checks; the reason for delay when the treatment date is the LG's date and more than 30 minutes have passed (`Please Enter the reason for Delay.!`, Yes on `Are you want to Enter or Update the reason for Delay.?` stops); the balance test (`Insufficient Balance, Please Check`); SaveLG's checks in order (`Please select a Clinic` … `Treatment Date cannot be Greater Than Member Expiry date..!`, `Please Check NO GLTemplate is Assigned for Plan Code : …`); `Are you sure?`; then LG, Claim, CISCas (pending cleared), TblRecord (`LG Issued`), ClaimInvoice and ClaimInvoiceActualNew in one transaction → `Record Updated....`, the GL letter (`LG Generated Successfully`) and the company notification e-mail (Tbl_CISEmailMaster by group / subsidiary / department / cost centre). |
| Decline | `Please Select Decline Reason`; `Please Select a LG!`; `Are you sure you want to decline [LG]?`; UpdateLG2, the pharmacy request, TblRecord `LG Declined` → decline letter, `Decline Letter Generated Successfully`. |
| Preview / Copy File | The letter as it stands; Copy File puts a file in the LG's folder (pharmacy: bmp, gif, png, jpg, jpeg, pdf as RX_&lt;LG&gt;-&lt;n&gt;; `Specified File Copied Successfully`). |

Known desktop issues (reproduced):
- Add never creates an LG: it only opens the fields, and LG then says `Kindly Press Edit..Then Continue`. LGs are issued by loading an existing request and using Edit.
- In the "B" limit branches the balances are joined as text before being compared or shown (100 and 50 give 10050).
- A blank GL amount is a VB type mismatch in the balance test, so nothing happens; for an LG with a previous amount the test compares IsNumeric (−1) with the balance.
- With GP / SP not shared and no specialist balance, any positive amount gives `Insufficient Balance, Please Check`.
- The daily visit check never stops (its result is ignored), so `Specialist GL issued already for the same visit date...!!` is never shown.
- Clinics picked from the group company list are shown as `NAME     [ CODE]`, which never matches the panel check, so non-PEMC GLs to them stop with `The selected clinic is not in HC panel. GL not allowed..!!`.
- The private delivery amount compares the Normal and Caesarean boxes as text; ClaimInvoice always takes the GL amount.
- Decline passes the decline reason as the LG remarks (UpdateLG2).
- The e-mail is sent depending on the EmailInd of the last LG in the list, not the one issued.

Deliberate differences: parameters instead of pasted SQL (CLNLG1 / LGlst2020 filters are built from validated values); the Word templates (`\\Svr-doc\Template\CIS LG\*.dot`) are not available, so the GL, decline and preview letters are PDFs carrying the same fields (the AIA additional page, SEB health assessment form and excess letter are not produced); the notification e-mail uses the configured mail sender instead of the SMTP account and password hard-coded in the form; questions are asked one at a time by the page and answered back to the API.
