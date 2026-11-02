# Invoice screens (Phase 3)

Desktop forms under Invoice (and the Phase 3 claim screens). Case Transfer and General Exclusions (Enquiry) are documented in `membership.md`.

## Batch Transfer — FrmInvoiceBatchNoUpd (`/invoice/batch-transfer`, right 119, `api/invoice/batch-transfer`)

| Item | Behaviour |
|---|---|
| Batch No. 1 / 2 (LostFocus → ShowInfo) | upper-cased; CISDB `SP_SeachCasByOPBatch` (`cis.BatchTransfer_Batch`): no rows → `Invalid OP Batch Number ! ` and the batch is cleared; else No. of Cases (sum), Invoice No. (distinct values joined with ","), Invoice Amt (sum of the amounts, each rounded to cents) and Clinic Name (last row). More than one row → `Invoice No : <invoices>, Different Clinic Found !`. |
| Final | Final Batch No. = batch 2, total cases, total amount (FinalBatchCal). |
| Transfer | `Please confirm your Batch Number ! ` (a batch empty), `Different Invoice No. Must be SAME Invoice No. ! `, `Same Batch Number ! `, `Are you sure?`; one transaction: batch 1's CISCAS rows get OPBatchNo = batch 2 and BatchTransRemarks = `Transfer By : <user> <batch 1>-><batch 2> , yyyymmdd`, then every batch 2 row gets InvoiceStmtAmt = the total → `Records Transfered Successfully ! ` and the form clears. |
| Clear | empties the form. |

Deliberate differences:
- The API reads both batches again before saving (the desktop compared the text boxes ShowInfo filled); the updates are parameterised (the desktop concatenated the batch numbers into SQL).
- After a successful transfer the desktop also showed an empty message box (its error label runs after the commit); not reproduced.
- `ReSequence` (OPBatch sequence renumbering) is never called by the desktop; not migrated.

## Invoice Update — frmUpdateCLNInvoice (`/invoice/update`, right 119, `api/invoice/update`; hidden in the desktop menu)

| Item | Behaviour |
|---|---|
| Search | Filters, each optional: WS No. from / to (CASNumber), Clinic (panel clinics Y / C; the code before " - "), Payor (group company), Invoice No. (BillNo), Without Invoice No., visit dates from / to, invoice dates from / to, invoice amount from / to. Rows from CISWorksheet `VIEW_Invoice_Tracking_2016` (the desktop's ClaimBordxModule case 1, as `cis.InvUpdate_Search`). None → `No record found!`. Invalid dates → `Invalid Date Format! ` on leaving the field and the filter is ignored. |
| List | Case No, Member No, Cov, Patient Name, Visit Date, Clinic, Invoice No, Invoice Date, Entered Date / By, Invoice Amt, FFS, Final Approve, Group Company; check boxes; Check toggles every row (`Please select a record.` when empty). |
| Print | the list as Excel (`Report cannot be printed without any record.`). |
| Update | `Please select the record`, `Invalid Invoice Date!`, `Please Enter Cheque Amt!` (the desktop's text for a missing invoice number), `Do you want to update the record ? `; each checked case: CISCAS InvoiceNo, InvoiceDate, RegisterDate = today, RegisterBy = user, in one transaction → `Record Saved...` and the list shows the cases of the new invoice number and date. |

Deliberate differences: filters are parameters (the desktop built SQL text and passed it to ClaimBordxModule); an amount filter that is not a number is ignored (the desktop failed with an SQL error); the unreachable `Please Enter Invoice Date! ` check is not shown.

## OP Batch/Invoice — frmUpdateInvoiceBatch (`/invoice/op-batch`, right 119, `api/invoice/op-batch`)

| Item | Behaviour |
|---|---|
| Search | Open cases (CASStatus O, payor not DZ / EM / ES, not DEMO COMPANY BERHAD) of CISCAS + CISWorksheet Claim (the desktop's ClaimBordxModule case 27 as `cis.OpBatch_Search`). Filters, each optional: WS No., clinic, payor, visit / invoice / register dates, invoice no., amount, OP batch and bordereaux / BTBG ranges, the "without OP batch / invoice / bordx / BTBG bordx" boxes, OMNI, clinic type (0 S, 1 G, 2 none), FFS clinic, invoiced Y / N, pay to, referral letter, entity (SJ cases or not). Cases with Pending Z are left out. None → `No record found!`. Totals: records, invoice amount, FFS. |
| Check | a case with an OP batch and an invoice → `Selected Case Already OP Batch And Invoice No generated..! `; pending → `Selected Case is Pending..! `; claimability not A → `Please check Case Claimability..! `. Check (all) checks the others and warns `Un Selected Few Cases Already OP Batch or Case Pending Or Case Claimability Or Invoice No generated ..! `. |
| Print | the list as Excel. |
| Update with an invoice no. | `Please Search Data..!`, `Please select Visit Date.`, `Above claims must belong to same clinic code.` (no clinic filter), `All the claims visit date must belong to same Month & Year.`; invoice no. = typed no. + "-" + the first four characters of the clinic; invoice date = last day of the month of the last listed case's visit; used already → `Given Invoice No is already used...!`. |
| Update | `Please select the record`, `Do you want to update the record ? `. Checked cases without an OP batch share one new batch (OPBatch "OC" sequence: `OC` + letter + 7 digits, `cis.OpBatch_Next` with a lock) with OPNoEnterUser = user and OPBatchDate = today; with an invoice no. every checked case also gets InvoiceNo, InvoiceDate, CASInvMonth (MM) and CASInvYear → `Record Saved...` and the list clears. Without an invoice no. cases that already have a batch are not changed. |

Deliberate differences / known desktop issues:
- The search is parameterised; the desktop built SQL text for ClaimBordxModule.
- The desktop's check rules test the selected row instead of each row (Check all and ItemCheck); the web applies them to each row, and the API re-checks them.
- All checked cases are saved in one transaction (the desktop committed per case and, for a case with a batch but no invoice no., left a transaction open).
- The OP batch number is generated inside the save transaction with a lock (the desktop could hand the same number to two users).
- The "pending state" message after listing can never appear in the desktop (it tests a Pending "Z" value for "Y"); not shown. Double-click to open the case in Adjustment comes with that screen.

## Invoice pop-ups (`api/invoice/popups`, open to rights 117 / 118 / 119 / 179)

| Pop-up | Behaviour |
|---|---|
| ICD search — frmSearchICD (`IcdSearchModalComponent`) | Code (starts with), Description (contains), Criteria AND / OR, Description 2 (contains); CISDB TreatmentReason (`cis.Popup_Icd`); none → `Record not found`; a click returns code and description to the calling form's diagnosis field. With OR the desktop's unbracketed SQL applies the code / first description only to the first part; reproduced. |
| Case list — frmCasList (`CaseListModalComponent`) | at least 3 characters (`Search String must be at least 3 characters!`); search by Member No (member or covered person number, starts with), Policy No, Name (contains, default), IC/BC/PP, Case No; invoice (C) or dental (D) cases of active members in VIEW_Cas_MBM (`cis.Popup_CaseList`); none → `No Record Available`; patient blank → `Member may has been Deleted!`; clinic shown as "code - name" for panel clinics. |

Built with their callers (they read or write the caller's state): case pending (frmCasePending, with Adjustment), member info (FrmINVMBMinfo, uses Medix.bas MemberSuppInfo, with the worksheet / invoice screens), invoice detail (frmInvoiceDetPopup, with its caller).

## MMA/CPT — frmMMASchedule + frmShowInvestigation (`/invoice/mma-cpt`, right 150, `api/invoice/mma-cpt`)

| Item | Behaviour |
|---|---|
| Search | Filters (each optional): Main Body Part, OPCS Code, Procedures (contains, upper-cased), Surgical / Anaesthetic Category (contains the text **without its first character**, as the desktop's `Mid(x, 2, 8)`), MMA Part (exact), Page from / to, CPT Code (exact), CPT Description (contains); HISMaintenance `VIEW_MMA_Schedule` (`cis.Mma_Search`). None → `Record not found`. |
| List | main body part, OPCS, procedure, surgical / anaesthetic category (without the first character), surgeon / anaesthetist fee (Surgeon13thNew / Anesthetist13thNew, not part A), part A fee (part A only), page, part, CPT code / description. |
| Row | opens the details (frmShowInvestigation, Source 1): body part, OPCS, categories and fees, part A fee, page, part, description. |
| Print | the list as Excel. |

frmShowInvestigation's second use (Source 2, the OPCS of an invoice line via Case_MMASchedule) comes with the invoice screens.

## Pass to Claim — frmPass2Claim (`/claim/pass-to-claim`, right 163, `api/claim/pass-to-claim`)

| Item | Behaviour |
|---|---|
| Search | Scan mode (default): one case number (`Please Enter Case No`; a case already listed → `Data has been uploaded!`; the list shows the scanned case). Otherwise case no. from / to. Filters: case status (0 Open, 1 Delete, 2 Open & Close — default 2), claimability A / R, Pass By (HISDB SearchCASToken), status P passed / N not, empty register / admission / pass date, and their date ranges. CISWorksheet `VIEW_ClaimPass2Claim` grouped (the desktop's ClaimPass2Claim case 1 as `cis.Pass2Claim_Search`); only rows whose cover id equals the member's cover id; None → `No record found!`. |
| List | Passed / Not, case, member, patient, dates of register / admission / pass, no. of days ("0" shows 1), OP batch, pass by, clinic, claim status. |
| Check | register date empty → `Date of Register is Empty.`; admission empty → `Date of Admitted is Empty.`; clinic empty → `Clinic Name is Empty.`; claim status 1 → `File has been send out.`. Check (all) checks the rest. |
| Pass | list empty → `Please select a payor before proceeding!` (desktop text); `Do you want to update the record?`; each checked case: ClaimPass2ClaimUpdate (ClaimFileTransfer row, status P, user, today) in one transaction → `Record updated...` and the list shows the passed cases. |
| Print | the list as Excel. |

Deliberate differences: the search is parameterised (with `OPTION (RECOMPILE)` so a single case uses the index); after Pass the desktop's unbracketed "AND CASNumber = a OR CASNumber = b" relist could show unrelated cases, the web lists exactly the passed ones; a scan code starting "LO" opened the first legacy screen the user may access (a scan-login shortcut), not migrated.

## Update Claims Data — frmUpdateClaimsData (`/claim-bordx/update-claims-data`, right 124, `api/claim-bordx/update-claims-data`)

| Item | Behaviour |
|---|---|
| File | an Excel workbook; the first sheet's rows from row 2 while column A is filled (at most 60,000). No file → `Please select file to upload!`. |
| Verify FFS | column 1 bordx no. (8 or 10 characters, starting N or S), column 8 approval code (12 characters), column 9 approval date, column 10 amount (numeric). |
| Verify Copay | column 2 bordx no., column 3 amount, column 20 return date. |
| Verify Bordx Amt | column 2 bordx no., column 3 amount, column 20 send date. |
| Verify Bordx Date | column 2 bordx no., column 20 date. |
| Result | rejected rows (row, bordx no., code, approval date, amount, date) → `Error encoutered. Please check the file!`; none → `File verified successfully!` and the matching upload button is enabled. |
| Upload | `Do you want to Upload the records?`; the file is verified again, then in one CISWorksheet transaction: FFS → Claim.CLMMCOProcFee of case (column 1) = column 10; Copay → CLMCoPayColl of case (column 2) = column 3; Bordx Amt → CLMBordxAmt of case (column 2) = column 3; Bordx Date → CLMBordxMonth / CLMBordxYear of the claims of bordereaux (column 2) and Bordx.BordxDate = column 20; ClaimLastUpdateDate / User set → `Records have been uploaded!`. |
| Print | the rejected rows as CSV. |

**Deliberate difference (agreed 2 Nov 2026):** the desktop's uploads could not work — FFS and Copay read column 2 as both the case number and the amount, Bordx Date used an empty date (always rolled back), Bordx Amt's update was commented out, "Verify Bordx Date" had no code and "Verify Bordx Amt" checked the wrong list (always verified). The web keeps the verification rules and messages and makes the uploads work with the columns each verification checks.
