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
| Case list — frmCasList (`CaseListModalComponent`) | at least 3 characters (`Search String must be at least 3 characters!`); search by Member No (member or covered person number, starts with), Policy No, Name (contains, default), IC/BC/PP, Case No; invoice (C) or dental (D) cases of active members in VIEW_Cas_MBM (`cis.Popup_CaseList`); none → `No Record Available`; patient blank → `Member may has been Deleted!`; clinic shown as "code - name" for panel clinics. Each search option runs as its own indexed query (one catch-all query took about 9 s per search on the 8.5 million cases). Like the desktop, Name is the default option: a member or case number typed without choosing Member No / Case No finds nothing. |

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

## Worksheet — frmWorksheet (`api/invoice/worksheet/calculate`, rights 117 / 118 / 119 / 179)

The worksheet pop-up does **not** write to the database: on opening it computes a benefit grid from the calling invoice row and copies the results
back into that row (`SaveCLMActual` … `SaveClaim`); the invoice / adjustment screen saves them. The web port is a pure calculation
(`WorksheetEngine`, unit-tested) fed by the invoice row (`WorksheetInput`) and `cis.Worksheet_Lookups` (BNFCategoryNew, CLN, STDClinicBNF,
STDClinicBNFSPE, PLN, BNFNew, CLNRedAmt). Callers: clinic invoice (1), member invoice (2), adjustment (3), after bordx (4), after posting (5).

| Column / line | Rule |
|---|---|
| Actual | the item amount (I1–I31: row 15–20, 24–42, 21–23, 101 / 95, 474 / 473, 475); 0 → "-". |
| Contracted Rate | cashless only. Clinic type 1: the actual. Type 2 (clinic invoice): codes 101–106 from STDClinicBNFSPE of the clinic's card, the rest from STDClinicBNF. Adjustments: the stored rates (row 195–225). |
| Payable to CLN | type 1: actual; type 2 / 3: GetMin(actual, actual). |
| Not Payable to CLN | type ≠ 1: rate − actual when the actual is higher (also for reimbursement, where the rate is 0). |
| Benefit Lmt | invoices: PLN.BNFStatus Y → BNFNew per code: 1 GetMin(% × actual, RM per visit), 2 GetMax (returns the RM per visit — desktop quirk), 3 actual, 4 RM per visit; status N → the actual. Adjustments: stored limits (row 297–327). Display only. |
| Approved Benefit | GetMin(actual, L) with L = the actual (the desktop sets L = I), i.e. the actual. |
| Excess | cashless: approved − actual when the actual is higher; reimbursement: the approved amount. |
| Final payable to CLN | total payable − co-pay collected (only when that minus the reduction is not 0) − reduction (CLNRedAmt: status 1 amount, 2 % of the payable); MCO proc. fee = CLN.CLNMCOProcfeePerc % when the row has MCO status Y and a GP amount. |
| Final approved | balance = annual indicator off → total approved, else row 70 + adjusted annual limit; visit = indicator off / FULL → total approved, else the visit limit; visit > balance → (approved > balance ? balance : approved) − co-pay collected; else approved − co-pay collected − co-pay. |
| Row values | 168–198 actual; 211–241, 426–456, 297–327, 254–284 also the actual (desktop copies the Actual column into the contracted-rate, payable, limit and approved slots — reproduced); 383–413 not payable; 340–370 excess; 150–166 claim totals / audit (159 is always 0 — never set by the desktop); 468 MCO fee; member invoice (476) / adjustment pay-to-M (475) override the approved amount when > 0. |

`CalculateReim` / `CalculateCashless` are never called by the desktop and are not migrated. The grid is shown by the invoice screens' worksheet modal (with those screens).

## Member info — Medix.bas MemberSuppInfo + FrmINVMBMinfo (`api/invoice/member-info/{member}/{cover}`, `MemberInfoModalComponent`)

`cis.MemberSupp_Info` returns the raw rows (VIEW_MBM_Clinical without deleted data, GRPCompany, MBMCLNOthers, PLN by plan / HLT / payor,
CoPaymentMBM, AnnualLimit.SuppLimitStatus, the covered person and MBMCoveredPersonsCLN of the cover and cover 01); `MemberSuppService` applies
the desktop's rules: `Membership No not found ! `, `Clinical Plan not found ! `, `No such Cover Person ! `; supplementary limit type A unless
ANNLmtIndSupp = Y and AnnualLimit gives B / C / D; principal of type D → balance FULL and annual indicator N; AnnLmtIND N → balance FULL;
VISLmtInd N → visit FULL; principal P* / cover S* dental, specialist and MEP statuses; covers use their own (C) or cover 01's (B / D) clinical
limits (supplementary indicators N → FULL), card type and status; no cover limit row → status D. The pop-up shows the member, policy and limits.

## Clinic invoice — frmInvoice (`/invoice/clinic`, right 117, `api/invoice/clinic`)

| Step | Behaviour |
|---|---|
| Header | Inv. No, Inv Date, Rec Date, Rec Mode (InvRecMode), Clinic (CLN: `No such Clinic Code ! `; a clinic that is not panel C / Y — the desktop does nothing (its branch is commented out); the web says `Non-Panel Clinic !`), Tot. Amt. |
| Go | `Please Key Member No Or IC Number ! `; "member-cover" or member + cover; MemberSuppInfo; ASO plans: Tbl_ASOMemberDetails (combined → ASO limit and IP + OP utilised, else OP); otherwise GetAnnualLimit (limit status), top-ups, GetVisitLimit (shown as the member's visit limit); specialist limit (GetSpecLimit with insured type = the member number's 3rd character, payor DC); utilisation from CIS_FFS2013 by limit status (A member, B / D principal or covers, C / E / F / blank the cover) = bordereaux amount − others − typhoid − maternity / dental by the plan's flags, specialist claims to the specialist total unless GP/SP shared; ASO → the ASO utilised; limit F covers (relationship D / S, ChildLimitInd Y) use ChildDepLimit; balance = annual + top-up − utilised (blank when the annual limit is FULL); `MC Days Exceeded! `. |
| Items | GP fee B1–B4 → A3 if B2, A2 if B3 and B4, else A1 (SP B5–B8 → A6 / A5 / A4); procedures, laboratory, immunisation, dental, maternity, PRV, vision, room & board, typhoid, others; co-pay A29 (Prince Court C014: 5 when a GP fee is entered); Total Bill = all amounts except the co-pay. |
| Add | visit-date checks (`Visit Date cannot greater than system date ! `, `Membership has been terminated before the entered visit date.`, `Visit date entered is not in range of Policy Effective Date and Expiry Date`); ChkData in the desktop order (`Please keyin Member No ! ` … `Please check Bill Amount! ` (Bill Amt must equal the total), `Please Enter Claimability `); same member / cover / visit date / clinic already registered → `Invoice has been entered !, Press 'Yes' to Proceed, 'No' to Cancel`; then the worksheet (Compute) and Verify. |
| Verify | DiagNotPayable, `Invoice No Can't be null ! `, invoice / visit date against the policy (+ 14 days), `45 Days LATE ! `, `Bill Amt > Balance Limit ! `, `Bill Amt > Visit Limit ! `, `Invoice Amt is different from Total Bill Amt`; no errors → `Verifed Sucessfully...`; errors are listed and can be exported. Save is enabled after Verify either way (as in the desktop). |
| Save | `No Balance Limit to Register Case. Please Check!!!` (GP: balance − bill + typhoid + others < 0; SP: specialist balance); `Are you sure ? ` / `ASO Limit already Exceed. Still you want to Proceed... `; one transaction: case number (WebSearchCASNo_New2013test, "EA" + 8 digits), Save_Case_2013, Save_Invoice, Save_ClaimActual, SaveActualOthers, SaveINVClaim_2015, the seven SaveINV* worksheet procedures, Save_Charges, member balance / accumulated / adjusted annual limit and MC days (HISDB), duplicate numbering (DupClaimSeq); ASO → sp_ASOClaimsUtilization; `Record Saved ! ` with the case number. |

Known desktop issues (reproduced): DiagNotPayable is checked against the plan code taken from row 64, which is diagnosis 3; the clinic card (row 72) is never set, so the worksheet's card rates and the CSMD check never apply; TxtCoPay is never filled (worksheet co-pay 0); TotalFinalCLNExc is always 0.

Deliberate differences: the desktop pasted every value into SQL text for the save procedures — amounts of 1,000 or more ("1,000.00") broke the call; the web passes parameters. The Add visit-limit check sets "isVisit" in every branch, so its message can never appear; not shown. A gross-amount status other than Y / N leaves the balance unchanged (the desktop wrote an empty value). The case is saved in one transaction with the server date.

## Member invoice — frmInvoiceMBM (`/invoice/member`, right 118, `api/invoice/member`)

Reimbursement claims (pay to M, payee MEMBER). Go, set-up and diagnoses are the clinic invoice's.

| Step | Behaviour |
|---|---|
| Clinic | CLN: `No such Clinic Code ! `; NP / NPS take a typed clinic name; another non-panel clinic → `Non-Panel Clinic !, Press 'Yes' to Proceed, 'No' to Cancel` (the desktop then opens Entity maintenance to register it). |
| Fields | Clinic type (G-General Practitioner / S-Specialist), Accepted (Y / N), repudiation reason / action / amount (Accepted N), referral letter, pharmacy (IBM members). No FFS, process month or claimability. |
| Add | `Bill Amount exceeded Balance Limit entitlement! Please check the charged Bill Amount should not more than Balance Limit.` (balance − Tot. Amt < 0); visit-date checks; ChkData order: member, clinic, invoice no / date, receipt date, `Please Enter Invoice Receipt Mode ! `, amount (Accepted Y), visit date, zero total (Accepted Y), `Please keyin the NP Clinic ! `, others diagnosis, `Please Enter Clinic type. ` (NP), `Please Enter Accepted status `, `Please Enter Repudiation Reason`, `Please Enter Action` (action and amount), `Please enter status for referral letter` (NPS). |
| Co-pay | referral letter N and pharmacy N → −(total × (1 − PLN MM %)); IBM members (sp_IBMMemberInformation): visit from 2014 — pharmacy Y −10 %, GP accepted −5, GP otherwise 0, specialist −25 %; 2013 specialist without referral −25 %. The total includes the (negative) co-pay. |
| Worksheet | reimbursement: approved = actual; the group company's bordereaux fee (when > 0) replaces the approved amount (desktop row 476). |
| Verify | as the clinic invoice, plus: balance check includes others without deduction; visit after expiry → `Visit date IS GREATER THAN Expiry date` (Save hidden); balance exceeded → `Please check balance limit` (Save hidden); same-day cases at the clinic → `Invoice has been entered . MBMNo =… , Date Visit = …` (Save hidden unless a negative amount is confirmed by `Is this Reversal Claim? `); Tot. Amt + co-pay must equal the line total (`Invoice Amt is different from Total Bill Amt`). Other listed errors do not block Save. |
| Save | balance numeric and not > 0 (unless amount and balance are both 0) → `No Balance Limit to Register Case. Please Check!!!`; `Are you sure ? `; one transaction: an OC batch (OPBatch letter + 7 digits) for the save, case number, CISCAS (accepted, NP clinic, clinic type, repudiation, pharmacy), ClaimInvoice (amount less co-pay), ClaimInvoiceActualNew (incl. co-payment), ClaimInvoiceActualOthers, Claim (postage 0, bordereaux memo, referral letter), the worksheet slot tables, member balance / accumulated / adjusted annual limit / health limit (PLN health limit Y); ASO → sp_ASOClaimsUtilization. |

Known desktop issues (reproduced): the ChkData duplicate check compares keys in two different formats and never finds a saved case (the Verify same-day check is the effective one); the health-limit block of Add is empty. Deliberate differences: the web saves one line at a time (the desktop could list several before saving); the inserts are parameterised; the SaveINVMBMApp procedure (shared with the clinic invoice) has no Pap Smear column, so that one slot is not stored.

## Invoice enquiry — FrmInvoiceEnq (`/invoice/enquiry`, right 119, `api/invoice/enquiry`)

Read-only view of one case. Go: CISWorksheet SP_INV_CLAIM (`Not Record Found ! `) with the GP / SP fee parts; the member panel (Go of the
clinic invoice); Prev / Next (VIEW_Cas_MBM clinical cases of active members: `First Record Reached! … Not a valid Case Number! ` /
`Last Record Reached! `); Search (case list); MBM Info; Exclusions (`Please Check Group Company! ` or the General Exclusions enquiry of the
case's group company). Tabs: Items (the stored actual amounts), Worksheet (Claim + the seven worksheet tables), Bordx (SP_Claim_Bordx incl. Medix
payment), Payment (CISPayment Payment2CLN / Payment2MBM and ReceiptFrMBM), Documents (Outpatient folder: Guarantee Letter\<case>,
SP_ScannedDoc\<case>.pdf / .jpg, OtherScannedDoc\<case>, Reim_ScannedDoc\<case>; files open by download), Medication Request
(sp_CISMedRequestInfo, sp_CISMedReqDOMedicine; Print exports the requests).

Deliberate difference: the member panel uses the clinic invoice's Go (top-ups included); the desktop's enquiry left top-ups out of the balances
and only counted specialist claims of the same cover side.

## Invoice adjustment — frmInvoiceAdj + frmCasePending (`/invoice/adjustment`, right 119, `api/invoice/adjustment`)

| Step | Behaviour |
|---|---|
| Go | SearchCASOP (`Not Record Found ! `), the case's Claim row, payments, SUN posting, the latest pending reasons (spGetPendingReasonsByCase), the specialist co-pay (SP_SpecialistCopayamt), the member panel and the stored worksheet; same-visit cases → `Multiple cases Found for this Member on same VistDate: …`; audited → `This case already audited..!`. |
| Edit | posted to SUN (not bordered / paid, or paid and bordered and posted) → `This case posted into SUN System Only can amend Diagnosis and Remark ! ` (diagnoses and remark only); bordered → `This case already Bordx ! `; paid → `This case already Payment done ! `; otherwise everything; the plan's special remark is shown. |
| Pending | frmCasePending: reasons of the pay-to category (spGetPendingReasons_OP), questions to the hospital (no duplicates; `You have typed the question but not added to the list. Please add it or remove it!`), additional remark; pending changed to N with details left → `Case pending status changed to 'N', but pending details not cleared! Please check`. |
| Update | diagnosis-only: CISCAS diagnoses, others and remark. Otherwise, in the desktop order: member, visit date, `Please keyin the Clinic code ! `, NP / NPS clinic names, others diagnosis, `Invalid Clinic Code` (pay C with NP), `Please Enter Bill Amount ! `, `Please check Bill Amount! ` (must equal the total), referral letter (NPS), `Please enter invoice date`, the balance check excluding this case (`Bill Amount exceeded Balance Limit entitlement! …`, GL not issued), `Please check case status.. ! `, zero total, pending reasons, UB duplicate invoices (`Duplicate case found: …`), UB / BS rejection without a reason and pending without an action, `Please enter BTBG Remark ! ` (BTBG plans G018 / G019), `This Case Number information seems created before! …` (specialist, pay C). Save: INVAdj_SaveCase_2013V3 (pending reason history `|date time-ids`), QA, LG amount, medication charges, INVAdj_SaveInvoice, INVAdj_saveActual (with GST), INVAdj_saveActualOthers, INVAdj_SaveINVClaim_2015, the seven worksheet procedures, the member balance (previous approval undone, new one applied); an OC batch when the invoice number is set, open, accepted and not pending (cleared when the invoice number is blank); UB / BS reimbursement rejection / pending e-mails (sp_EmailNotifyPendingReject_LatestModified) with the desktop's prompts and messages (`Record updated and Email Sent! `, `Record updated..!Since no change No email is Sent. ` …). |
| Other | Audit review (`Please Enter Query! `, `Please Enter Audit By! `, `Please Enter Audit Date! ` → `Audit Details updated ! `); Claims App (CASClaimsApp = Y); Void Bordx for reimbursements (`VOID Bordx Only For Reimbursement Cases..!` … → `Void Bordx Generated! Bordx No: KIV… `). |

Known desktop issues (reproduced): the pending reasons list is empty when the case already has reasons; the reimbursement payable reads the non-GP fee before it is set (payable to member has no non-GP part); a gross-amount status Y on a clinic case leaves the balance unchanged (the formula cancels); "rejection" means status D with claimability C (list positions 3 / 3).

Deliberate differences: parameters instead of pasted SQL (formatted amounts ≥ 1,000 broke it); the member balance is updated in the same transaction (the desktop's HISDB update was outside it); a shared-limit cover updates cover 01 (the desktop took an unordered "top 1" cover); the deferment letter (Word template `HCT_GL_deferment_letter.dot`, not in the source) is not generated; the diagnosis-only update shows a short confirmation (the desktop showed nothing).
