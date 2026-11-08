-- LG & Audit > Generate LG (frmLGGeneration): the checks of Edit / LG, issuing (SaveLG), declining (cmdDecline) and the LG e-mail recipients.
SET QUOTED_IDENTIFIER ON;
GO

-- CmdEdit_Click / cmdLG_Click checks: bordereaux / invoice, payments, spLGforSP_check, spclnforpanel_check (by the clinic text, as the desktop),
-- statusInfo (expiry within 30 days of today), pending, and the LG's case date / time.
CREATE OR ALTER PROCEDURE cis.Lg_Checks @lgNo nvarchar(20), @member nvarchar(30), @cover nvarchar(2), @clinicCode nvarchar(20), @clinicText nvarchar(300) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(COALESCE(c.ClaimBordxStatus, 0) AS bit) AS Bordx, LTRIM(RTRIM(COALESCE(cas.InvoiceNo, ''))) AS InvoiceNo,
        CAST(COALESCE(c.BTBGBordxStatus, 0) AS bit) AS BtbgBordx
    FROM cis.Ws_Claim c INNER JOIN dbo.CISCas cas ON cas.CASNumber = c.CASNumber WHERE c.CASNumber = @lgNo;
    SELECT (SELECT COUNT(*) FROM cis.Pay_Payment2CLN WHERE CASNumber = @lgNo) + (SELECT COUNT(*) FROM cis.Pay_Payment2MBM WHERE CASNumber = @lgNo) AS Payments;
    DECLARE @sp TABLE (Result nvarchar(20));
    INSERT @sp EXEC dbo.spLGforSP_check @mbmno = @member, @clncode = @clinicCode;
    SELECT TOP 1 Result FROM @sp;
    SELECT CLNPanel FROM cis.Hdb_CLN WHERE LTRIM(RTRIM(CLNName)) + '[' + CONVERT(varchar, LTRIM(RTRIM(CLNCode))) + ']' = @clinicText;
    IF @cover = '00'
        SELECT TOP 1 o.MBMPolExpDate AS ExpiryDate FROM dbo.CISCas cas INNER JOIN cis.Hdb_MBMCLNOthers o ON o.MBMNumber = cas.MBMNumber
        WHERE o.MBMPolExpDate >= GETDATE() - 30 AND cas.CASNumber = @lgNo;
    ELSE
        SELECT TOP 1 p.MBMCCLNExpDate AS ExpiryDate FROM dbo.CISCas cas
        INNER JOIN cis.Hdb_MBMCoveredPersonsCLN p ON p.MBMCNumber = cas.MBMNumber AND p.MBMCCoverID = cas.CoverID
        WHERE p.MBMCCLNExpDate >= GETDATE() - 30 AND cas.CASNumber = @lgNo AND cas.CoverID = @cover;
    SELECT Pending FROM dbo.CISCas WHERE CASNumber = @lgNo;
END
GO

-- SaveLG (the update of the selected LG): LG, Claim, CISCas, TblRecord, ClaimInvoice and ClaimInvoiceActualNew in one transaction.
CREATE OR ALTER PROCEDURE cis.Lg_Issue @lgNo nvarchar(20), @user nvarchar(50), @treatmentDate datetime, @clinicCode nvarchar(20), @remarks nvarchar(max),
    @diagnosis nvarchar(max), @lgAmount money, @glAmount money, @internalRemarks nvarchar(max), @doctor nvarchar(200), @glType nvarchar(100),
    @maternity money, @pending bit, @delivery bit, @delayReason nvarchar(500) AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRAN;
    UPDATE dbo.LG SET LGDate = @treatmentDate, LGIssuedBy = @user, CLNName = @clinicCode, LGRemarks = @remarks, Diagnosis = @diagnosis, GLAmt = @lgAmount,
        InternalRemarks = @internalRemarks, DoctorName = @doctor, GLType = @glType, LGIssueDate = GETDATE()
    WHERE LGNo = @lgNo;
    UPDATE cis.Ws_Claim SET ClaimOpenDate = CAST(GETDATE() AS date), ClaimOpenBy = @user, ClaimCloseDate = CAST(GETDATE() AS date), ClaimClosedBy = @user,
        CLMTotalActual = @lgAmount, CLMFinalApprove = @lgAmount, CLMPaidBYMedix2C = @lgAmount, CLMTotalApproved = @lgAmount, CLMBordxAmt = @lgAmount
    WHERE ClaimBordxStatus = 0 AND CASNumber = @lgNo;
    IF @pending = 1
        UPDATE dbo.CISCas SET CasDelayReason = @delayReason, Pending = 'N', CASTreatReason1 = CASE WHEN @delivery = 1 THEN '302' END, PendingReminderDate = NULL,
            CASRepAction = CASRepAction + '|' + CONVERT(nvarchar, GETDATE(), 103) + ' ' + CONVERT(varchar, GETDATE(), 108) + '-', DateVisit = @treatmentDate,
            Claimability = 'A'
        WHERE CASNumber = @lgNo;
    ELSE
        UPDATE dbo.CISCas SET CasDelayReason = @delayReason, DateVisit = @treatmentDate, CASTreatReason1 = CASE WHEN @delivery = 1 THEN '302' END, Claimability = 'A'
        WHERE CASNumber = @lgNo;
    UPDATE dbo.tblRecord SET IsLGissued = 1, IsRejected = 0, CaseProgress = 'LG Issued' WHERE CaseNo = @lgNo;
    UPDATE cis.Ws_ClaimInvoice SET INVAmount = @glAmount, INVBillAmt = @glAmount WHERE InvBordxed = 0 AND CASNumber = @lgNo;
    UPDATE cis.Ws_ClaimInvoiceActualNew SET CLMMaternity = @maternity WHERE CASNumber = @lgNo;
    COMMIT;
END
GO

-- cmdDecline_Click: UpdateLG2 (its remarks parameter gets the decline reason, as the desktop passed it), the pharmacy request, TblRecord.
CREATE OR ALTER PROCEDURE cis.Lg_Decline @lgNo nvarchar(20), @doctor nvarchar(200), @clinicCode nvarchar(20), @reason varchar(500), @user nvarchar(50),
    @diagnosis nvarchar(max), @pharmacy bit AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRAN;
    DECLARE @now datetime = CAST(GETDATE() AS date);
    EXEC dbo.UpdateLG2 @LGNo = @lgNo, @LGDate = @now, @DoctorName = @doctor, @CLNName = @clinicCode, @LGRemarks = @reason, @LGUpdateBy = @user,
        @Diagnosis = @diagnosis, @DeclineReason = @reason;
    IF @pharmacy = 1 EXEC dbo.sp_UpdateMedReqDeclineReason @CaseNo = @lgNo, @DeclineReason = @reason;
    UPDATE dbo.tblRecord SET IsLGissued = 0, IsRejected = 1, CaseProgress = 'LG Declined' WHERE CaseNo = @lgNo;
    COMMIT;
END
GO

-- SaveLG e-mail: the group company's mail type (first row), the member's subsidiary / cost centre / department, and the first matching recipient row.
CREATE OR ALTER PROCEDURE cis.Lg_EmailRecipients @company nvarchar(200), @member nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @cnt int, @type nvarchar(5);
    SELECT TOP 1 @cnt = COUNT(*), @type = Mailtype FROM cis.Maint_CISEmailMaster
    WHERE COALESCE(Mailtype, '') <> 'E' AND COALESCE(IsCIS, '') = 'Y' AND Grpcompany = @company GROUP BY Mailtype;
    IF COALESCE(@cnt, 0) < 1 RETURN;
    DECLARE @sub nvarchar(200), @cost nvarchar(200), @dept nvarchar(200);
    SELECT TOP 1 @sub = COALESCE(Subsidiary, ''), @cost = COALESCE(MBMCostCenter, ''), @dept = COALESCE(MBMDepartment, '') FROM cis.Hdb_MBMOthers WHERE MBMNumber = @member;
    SELECT TOP 1 EmailAddress, BCC, CC, @sub AS Subsidiary FROM cis.Maint_CISEmailMaster
    WHERE LGInd = 1 AND IsCIS = 'Y' AND Grpcompany = @company
      AND (LTRIM(RTRIM(@type)) = 'G' OR (LTRIM(RTRIM(@type)) = 'S' AND Subsidary = @sub) OR (LTRIM(RTRIM(@type)) = 'D' AND Department = @dept)
           OR (LTRIM(RTRIM(@type)) = 'C' AND Costcenter = @cost));
END
GO
