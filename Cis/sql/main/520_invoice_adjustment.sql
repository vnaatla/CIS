-- Invoice > Clinic > Adjustment (frmInvoiceAdj, with frmCasePending): load a saved case, check and update it, audit review, void bordereaux.
-- The save calls the desktop's INVAdj_* procedures with parameters (the desktop pasted formatted amounts into SQL text).

-- InvoiceINFO + CmdGo_Click: SearchCASOP, the case's Claim row, payments, SUN posting, the latest pending reasons, the specialist co-pay,
-- same-visit cases and the plan flags.
CREATE OR ALTER PROCEDURE cis.InvAdj_Load @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Ws_SearchCASOP @CASNumber = @casNumber;
    SELECT TOP (1) CLMCoPayColl, CLMFinalApprove, ClmNonGPFee, CLMAdjAnnLmt, CLMRefLetter, CLMLimitExceeded, CLMPayableByMedix2C, CLMMCOProcStatus,
           CLMMCOProcFee, CLMProcMth, CLMProcYR, CLMCoPayCLNColl, CLMBordxMemo, CLMPayableByMedix2M
    FROM cis.Ws_Claim WHERE CASNumber = @casNumber;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM cis.Pay_Payment2CLN WHERE CASNumber = @casNumber)
                       OR EXISTS (SELECT 1 FROM cis.Pay_Payment2MBM WHERE CASNumber = @casNumber) THEN 1 ELSE 0 END AS bit) AS Paid,
           CAST(CASE WHEN EXISTS (SELECT 1 FROM cis.Ws_SUNPosting WHERE InvoiceNo = @casNumber) THEN 1 ELSE 0 END AS bit) AS Posted;
    DECLARE @ids TABLE (ReasonID nvarchar(400));
    INSERT INTO @ids EXEC dbo.spGetPendingReasonsByCase @casNo = @casNumber, @latest = 1;
    SELECT ReasonID FROM @ids;
    EXEC dbo.SP_SpecialistCopayamt @CASNumber = @casNumber;
    -- The member's e-mail (MBMTwo), the plan's special remark (shown on Edit) and BTBG flag.
    DECLARE @mbm nvarchar(50) = (SELECT TOP (1) MBMnumber FROM dbo.CISCAS WHERE CASNumber = @casNumber);
    SELECT TOP (1) RTRIM(MBMEmail) AS MBMEmail FROM cis.Hdb_MBMTwo WHERE MBMNumber = @mbm;
    SELECT TOP (1) RTRIM(p.PLNSpecialRemark) AS PLNSpecialRemark, p.PLNBTBG FROM cis.Hdb_MBMCLNOthers o INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    WHERE o.MBMNumber = @mbm;
END
GO

-- PopulatePendingReasons: the active reasons of the category (pay to C / M).
CREATE OR ALTER PROCEDURE cis.InvAdj_PendingReasons @category char(1) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.spGetPendingReasons_OP @Category = @category, @Active = 1;
END
GO

-- duplicatecasesIdentify (on Go) / duplicateinvoicesIdentify (UB members) / CheckDoubleEntry.
CREATE OR ALTER PROCEDURE cis.InvAdj_SameVisitCases @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.spMemberduplicatecases_Identify @CasNumber = @casNumber;
END
GO

CREATE OR ALTER PROCEDURE cis.InvAdj_DuplicateInvoices @casNumber nvarchar(20), @invoiceDate date, @billAmt nvarchar(30), @invoiceNo nvarchar(50),
    @patientName nvarchar(200), @patientIc nvarchar(50), @userCode nvarchar(50), @coverId nvarchar(2)
AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.spMemberduplicateInvoices_Identify @CaseNo = @casNumber, @Invreceiptdate = @invoiceDate, @InvBillamt = @billAmt, @Invno = @invoiceNo,
        @PatientName = @patientName, @PatientIC = @patientIc, @Userid = @userCode, @Coverid = @coverId;
END
GO

CREATE OR ALTER PROCEDURE cis.InvAdj_DoubleEntry @mbmNumber nvarchar(50), @coverId nvarchar(2), @clinicCode nvarchar(10), @invoiceNo nvarchar(200),
    @visitTime nvarchar(20), @dateVisit datetime, @totalActual money, @casNumber nvarchar(100)
AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_chkDoubleCaseEntry @MBMNumber = @mbmNumber, @CoverID = @coverId, @ClnCode = @clinicCode, @InvoiceNo = @invoiceNo, @VisitTime = @visitTime,
        @DateVisit = @dateVisit, @TotalActualAmt = @totalActual, @CasNumber = @casNumber;
END
GO

-- CheckExceedLimit: the member's claims (CIS_FFS2018 filter) without this case.
CREATE OR ALTER PROCEDURE cis.InvAdj_OtherClaims @mbmNumber nvarchar(50), @mode int, @covId nvarchar(2), @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @rows TABLE (FFS money, CLNType nvarchar(5), CLMMaternity money, CoverID nvarchar(2), CLMBordxAmt money, PLNGPSPShared nvarchar(2),
        CLMOthersNoDed money, CLMTyphiadJabNoDed money, CLMDental money, CASNumber nvarchar(20));
    INSERT INTO @rows
    SELECT cl.CLMMCOProcFee, RTRIM(cln.CLNType), COALESCE(ia.CLMMaternity, 0), c.CoverID,
           CASE WHEN p.CoPayInd = 'Y' AND COALESCE(cl.CLMCoPayColl, 0) <> 0 THEN cl.CLMTotalActual - cl.CLMCoPayColl ELSE cl.CLMTotalActual END,
           RTRIM(p.PLNGPSPShared), CASE WHEN COALESCE(cl.CLMBTBGAmt, 0) <> 0 THEN COALESCE(cl.CLMBTBGAmt, 0) ELSE COALESCE(ia.CLMOthersNoDed, 0) END,
           ia.CLMTyphiadJabNoDed, COALESCE(ia.CLMDental, 0), c.CASNumber
    FROM dbo.CISCas c WITH (NOLOCK)
    INNER JOIN cis.Ws_Claim cl WITH (NOLOCK) ON c.CASNumber = cl.CASNumber
    INNER JOIN cis.Hdb_MBMCLNOthers o ON o.MBMNumber = c.MBMnumber
    INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    INNER JOIN cis.Ws_ClaimInvoiceActualNew ia WITH (NOLOCK) ON cl.CASNumber = ia.CASNumber
    LEFT JOIN cis.Hdb_CLN cln WITH (NOLOCK) ON c.ClinicCode = cln.CLNCode
    WHERE c.CASStatus <> 'D' AND c.MBMnumber = @mbmNumber AND c.CASNumber <> @casNumber
      AND (@mode = 0 OR (@mode = 1 AND c.CoverID = '00') OR (@mode = 2 AND c.CoverID <> '00') OR (@mode = 3 AND c.CoverID = @covId));
    SELECT FFS, CLNType, CLMMaternity, CoverID, CLMBordxAmt, PLNGPSPShared, CLMOthersNoDed, CLMTyphiadJabNoDed, CLMDental FROM @rows;
END
GO

-- SaveRecord's CISDB / CISWorksheet writes besides the INVAdj procedures: QA, LG amount, medication charges, OP batch.
CREATE OR ALTER PROCEDURE cis.InvAdj_SaveExtras @casNumber nvarchar(20), @qa nvarchar(max), @glAmount money,
    @c1 money, @c2 money, @c3 money, @c4 money, @c5 money, @c6 money, @c7 money
AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.spUpdateQA @casNumber = @casNumber, @QA = @qa;
    UPDATE dbo.LG SET GLAmt = @glAmount WHERE LGNo = @casNumber;
    UPDATE cis.Ws_CLMMedCharges SET MedCharges1 = @c1, MedCharges2 = @c2, MedCharges3 = @c3, MedCharges4 = @c4, MedCharges5 = @c5, MedCharges6 = @c6,
        MedCharges7 = @c7 WHERE CASNumber = @casNumber;
END
GO

-- UpdateRecord / UpdateEmptyOPbatchno.
CREATE OR ALTER PROCEDURE cis.InvAdj_Batch @casNumber nvarchar(20), @opBatchNo nvarchar(30) = NULL, @userCode nvarchar(50), @today datetime AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCas SET OPBatchNo = @opBatchNo, OPNoEnterUser = CASE WHEN @opBatchNo IS NULL THEN NULL ELSE @userCode END,
        OPBatchDate = CASE WHEN @opBatchNo IS NULL THEN NULL ELSE @today END
    WHERE CASNumber = @casNumber;
END
GO

-- DiagnosisUpdate (a case posted to SUN: only the diagnoses and the remark).
CREATE OR ALTER PROCEDURE cis.InvAdj_Diagnosis @casNumber nvarchar(20), @d1 nvarchar(20), @d2 nvarchar(20), @d3 nvarchar(20), @d4 nvarchar(20),
    @d5 nvarchar(20), @other nvarchar(200), @remark nvarchar(500)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET CASTreatReason1 = @d1, CASTreatReason2 = @d2, CASTreatReason3 = @d3, CASTreatReason4 = @d4, CASTreatReason5 = @d5,
        CASDiagnosis5 = @other, CasRemark = @remark
    WHERE CASNumber = @casNumber;
END
GO

-- UpdateMBMBalance of the adjustment: undo the previous approval, apply the new one (row: principal, cover 01 for shared covers B / D, or the cover).
CREATE OR ALTER PROCEDURE cis.InvAdj_Balance @mbmNumber nvarchar(50), @limitCover nvarchar(2) = NULL, @adjDelta money, @accumDelta money,
    @balanceDelta money = NULL
AS
BEGIN
    SET NOCOUNT ON;
    IF @limitCover IS NULL
        UPDATE cis.Hdb_MBMCLNOthers SET MBMAdjAnnLmt = COALESCE(MBMAdjAnnLmt, 0) + @adjDelta, MBMAccumAmt = COALESCE(MBMAccumAmt, 0) + @accumDelta,
            MBMCLNBalLmt = CASE WHEN @balanceDelta IS NULL THEN MBMCLNBalLmt WHEN COALESCE(MBMCLNBalLmt, 0) + @balanceDelta < 0 THEN 0
                                ELSE COALESCE(MBMCLNBalLmt, 0) + @balanceDelta END
        WHERE MBMNumber = @mbmNumber;
    ELSE
        UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCAdjAnnLmt = COALESCE(TRY_CONVERT(money, MBMCAdjAnnLmt), 0) + @adjDelta,
            MBMCAccumAmt = COALESCE(TRY_CONVERT(money, MBMCAccumAmt), 0) + @accumDelta,
            MBMCCLNBalLmt = CASE WHEN @balanceDelta IS NULL THEN MBMCCLNBalLmt WHEN COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) + @balanceDelta < 0 THEN 0
                                 ELSE COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) + @balanceDelta END
        WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @limitCover;
END
GO

-- cmdauditreviewupdate_Click.
CREATE OR ALTER PROCEDURE cis.InvAdj_Audit @casNumber nvarchar(20), @query nvarchar(max), @answer nvarchar(max), @discount nvarchar(50),
    @remarks nvarchar(max), @auditBy nvarchar(100), @auditDate datetime
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCas SET AuditQuery = @query, AuditAnswer = @answer, AuditDiscountReceived = @discount, AuditRemarks = @remarks, AuditBy = @auditBy,
        AuditDate = @auditDate
    WHERE CASNumber = @casNumber;
END
GO

-- cmdCLMBordxApp_Click.
CREATE OR ALTER PROCEDURE cis.InvAdj_ClaimsApp @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET CASClaimsApp = 'Y' WHERE CASNumber = @casNumber;
END
GO

-- cmdvoidbordx_Click: a "KIV" bordereaux number (Claim_BordxSEQ of pay code IV), the invoice and claim marked bordered, a BORDX row and a VOID
-- member payment.
CREATE OR ALTER PROCEDURE cis.InvAdj_VoidBordx @casNumber nvarchar(20), @groupCompany nvarchar(100), @userCode nvarchar(50), @today datetime AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @n int = (SELECT TOP (1) CAST(CLMBordxNo AS int) FROM cis.Ws_Claim_BordxSEQ WITH (UPDLOCK, HOLDLOCK) WHERE PAYCode = 'IV');
    IF @n IS NULL
    BEGIN
        SET @n = 1;
        INSERT INTO cis.Ws_Claim_BordxSEQ (HLTCode, PAYCode, CLMBordxNo) VALUES ('', 'IV', '00001');
    END
    ELSE
    BEGIN
        SET @n += 1;
        UPDATE cis.Ws_Claim_BordxSEQ SET CLMBordxNo = RIGHT('00000' + CAST(@n AS varchar(10)), 5) WHERE PAYCode = 'IV';
    END
    DECLARE @no nvarchar(20) = 'KIV' + RIGHT('00000' + CAST(@n AS varchar(10)), 5);
    UPDATE cis.Ws_ClaimInvoice SET INVBordxed = 1, INVBordxRefNo = @no WHERE CASNumber = @casNumber;
    UPDATE cis.Ws_Claim SET BordxRefNo = @no, ClaimBordxStatus = 1 WHERE CASNumber = @casNumber;
    INSERT INTO cis.Ws_Bordx (BordxRefNo, BordxPAYCode, BordxDate, BordxBy, GroupCompany) VALUES (@no, 'IV', @today, @userCode, @groupCompany);
    INSERT INTO cis.Pay_Payment2MBM (PVNo, PVDate, PVChequeDate, PVChequeNo, CASNumber, Pay2MBMLastUpdateDate, Pay2MBMLastUpdateUser, PAYAccPeriodMth, PAYAccPeriodYr)
    VALUES ('VOID', @today, @today, 'VOID', @casNumber, @today, @userCode, MONTH(@today), YEAR(@today));
    SELECT @no;
END
GO

-- CheckRejectionPending: the desktop's database e-mail notification (1 = sent).
CREATE OR ALTER PROCEDURE cis.InvAdj_Notify @casNumber nvarchar(20), @notes varchar(2000), @pendingReasons varchar(2000), @email varchar(2000) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_EmailNotifyPendingReject_LatestModified @Casnumber = @casNumber, @tempReason = @notes, @PendingReason = @pendingReasons, @tempemail = @email;
END
GO
