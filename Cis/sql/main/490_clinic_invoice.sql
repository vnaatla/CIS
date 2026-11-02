-- Invoice > Clinic > Clinic (frmInvoice): look-ups, the member's claims (CIS_FFS2013) and the save. Limits / top-ups / ASO / specialist
-- limits reuse cis.MbmAdj_Limits, cis.MbmEnq_Topup, cis.MbmEnq_Aso and cis.MbmEnq_SpecLimit (Phase 2).

-- ComboListing: receipt modes and process months.
CREATE OR ALTER PROCEDURE cis.InvClinic_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(RecModeCode) AS Code, RTRIM(RecModeDesc) AS Name FROM dbo.InvRecMode;
    SELECT RTRIM(INVProcMth) + '/' + RTRIM(INVProcYR) AS Month FROM dbo.INVProcMonth;
END
GO

-- TxtClinicCode_LostFocus (CLN) and GetClinicLimit (Tbl_CLNGPLimits).
CREATE OR ALTER PROCEDURE cis.InvClinic_Clinic @clinicCode nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) RTRIM(CLNPanel) AS Panel, RTRIM(CLNCode) AS Code, RTRIM(CLNName) AS Name, RTRIM(CLNType) AS ClinicType, RTRIM(CLNFFSStatus) AS FfsStatus,
           CLNMCOProcfeePerc AS McoProcFeePercent
    FROM cis.Hdb_CLN WHERE CLNCode = @clinicCode;
END
GO

-- CmdGo_Click look-ups after MemberSuppInfo: 1 AnnualLimit + PLN visit limits, 2 PLN maternity / dental / ASO, 3 GP / SP limits,
-- 4 SpecialistLimit status, 5 MBMOthers + GRPCompany (claim cost centre etc.), 6 Tbl_ASOMemberDetails.
CREATE OR ALTER PROCEDURE cis.InvClinic_MemberExtra @mbmNumber nvarchar(50), @covId nvarchar(2), @planCode nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) a.VisitLimit, a.prepostmatellimit AS PrePostMatLimit, p.PDentalStatus, a.PDentalAnnLimit, a.SPVisitLimit, a.MEVisitLimit, RTRIM(p.VISLmtInd) AS VisLmtInd,
           RTRIM(p.ChildLimitInd) AS ChildLimitInd, a.ChildDepLimit
    FROM dbo.AnnualLimit a INNER JOIN dbo.PLN p ON p.PLNCode = a.PLNCode WHERE a.PLNCode = @planCode;
    SELECT TOP (1) RTRIM(PLNMaternityInd) AS MaternityInd, RTRIM(PDentalStatus) AS DentalInd, RTRIM(ASOInd) AS AsoInd FROM dbo.PLN WHERE PLNCode = @planCode;
    IF @covId = '00'
        SELECT TOP (1) p.GRPCompany, p.PrinGPLimit AS GPLimit, p.PrinSPLimit AS SPLimit
        FROM cis.Hdb_MBMCLNOthers o INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType WHERE o.MBMNumber = @mbmNumber;
    ELSE
        SELECT TOP (1) p.GRPCompany, p.DepGPLimit AS GPLimit, p.DepSPLimit AS SPLimit
        FROM cis.Hdb_MBMCoveredPersonsCLN c INNER JOIN dbo.PLN p ON p.PLNCode = c.MBMCCLNPLNCode WHERE c.MBMCNumber = @mbmNumber AND c.MBMCCoverID = @covId;
    SELECT TOP (1) RTRIM(SuppLimitStatus) AS SuppLimitStatus FROM dbo.SpecialistLimit WHERE PLNCode = @planCode;
    SELECT TOP (1) LTRIM(RTRIM(COALESCE(o.MBMDepartment, ''))) AS MBMDepartment, o.MBMCostCenter, LTRIM(RTRIM(COALESCE(o.MBMDeptCode, ''))) AS MBMDeptCode,
           LTRIM(RTRIM(COALESCE(o.Subsidiary, ''))) AS Subsidiary, LTRIM(RTRIM(g.GRPCoCode)) AS GRPCoCode
    FROM cis.Hdb_MBMOthers o INNER JOIN dbo.GRPCompany g ON g.GRPCoName = o.MBMGroupCompany WHERE o.MBMNumber = @mbmNumber;
    SELECT MBMASOLimit, MBMOPLimit, MBMIPUtilized, MBMOPUtilized, RTRIM(IsIPOPCombined) AS IsIPOPCombined
    FROM cis.Hdb_ASOMemberDetails WHERE MBMNumber = @mbmNumber AND MBMCoverID = @covId;
END
GO

-- CIS_FFS2013 with the dynamic filter replaced: @mode 0 the member, 1 cover 00, 2 covers other than 00, 3 one cover.
CREATE OR ALTER PROCEDURE cis.InvClinic_Claims @mbmNumber nvarchar(50), @mode int, @covId nvarchar(2) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT cl.CLMMCOProcFee AS FFS, RTRIM(cln.CLNType) AS CLNType, COALESCE(ia.CLMMaternity, 0) AS CLMMaternity, c.CoverID,
           CASE WHEN p.CoPayInd = 'Y' AND COALESCE(cl.CLMCoPayColl, 0) <> 0 THEN cl.CLMTotalActual - cl.CLMCoPayColl ELSE cl.CLMTotalActual END AS CLMBordxAmt,
           RTRIM(p.PLNGPSPShared) AS PLNGPSPShared,
           CASE WHEN COALESCE(cl.CLMBTBGAmt, 0) <> 0 THEN COALESCE(cl.CLMBTBGAmt, 0) ELSE COALESCE(ia.CLMOthersNoDed, 0) END AS CLMOthersNoDed,
           ia.CLMTyphiadJabNoDed, COALESCE(ia.CLMDental, 0) AS CLMDental
    FROM dbo.CISCas c WITH (NOLOCK)
    INNER JOIN cis.Ws_Claim cl WITH (NOLOCK) ON c.CASNumber = cl.CASNumber
    INNER JOIN cis.Hdb_MBMCLNOthers o ON o.MBMNumber = c.MBMnumber
    INNER JOIN cis.Hdb_MBMTwo t ON t.MBMNumber = c.MBMnumber
    INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    INNER JOIN cis.Ws_ClaimInvoiceActualNew ia WITH (NOLOCK) ON cl.CASNumber = ia.CASNumber
    LEFT JOIN cis.Hdb_CLN cln WITH (NOLOCK) ON c.ClinicCode = cln.CLNCode
    WHERE c.CASStatus <> 'D' AND c.MBMnumber = @mbmNumber
      AND (@mode = 0 OR (@mode = 1 AND c.CoverID = '00') OR (@mode = 2 AND c.CoverID <> '00') OR (@mode = 3 AND c.CoverID = @covId));
END
GO

-- ChkData: cases of the member / cover on the visit date at the clinic (the duplicate check); cmdVerify: CheckItems CSMD and DiagNotPayable.
CREATE OR ALTER PROCEDURE cis.InvClinic_Checks @mbmNumber nvarchar(50), @covId nvarchar(2), @visitDate datetime, @clinicCode nvarchar(10), @card nvarchar(50),
    @planCode nvarchar(20), @diag1 nvarchar(20) = NULL, @diag2 nvarchar(20) = NULL, @diag3 nvarchar(20) = NULL, @diag4 nvarchar(20) = NULL, @diag5 nvarchar(20) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT CoverID, DateVisit, ClinicCode, InvoiceStmtAmt FROM dbo.CISCAS
    WHERE MBMnumber = @mbmNumber AND CoverID = @covId AND DateVisit = @visitDate AND ClinicCode = @clinicCode;
    SELECT Amount FROM dbo.CheckItems WHERE CatCode = 'CSMD' AND CatName = @card;
    SELECT RTRIM(DiagCode) AS DiagCode FROM dbo.DiagNotPayable WHERE PLNCode = @planCode AND DiagCode IN (@diag1, @diag2, @diag3, @diag4, @diag5);
END
GO

-- GenerateCISCaseNo: WebSearchCASNo_New2013test (pay code EA) → "EA" + 8 digits.
CREATE OR ALTER PROCEDURE cis.InvClinic_CaseNo AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @no nvarchar(2000);
    DECLARE @rows TABLE (v sql_variant);
    -- The procedure also returns the number as a result set; it is swallowed so only the formatted number comes back.
    INSERT INTO @rows EXEC dbo.WebSearchCASNo_New2013test @PayCode = 'EA', @CaseNo1 = @no OUTPUT;
    SELECT 'EA' + RIGHT('00000000' + CAST(CAST(@no AS bigint) AS varchar(20)), 8);
END
GO

-- SaveDup_ori: number the cases of the same member / cover / visit date / clinic when there is more than one.
CREATE OR ALTER PROCEDURE cis.InvClinic_Duplicates @mbmNumber nvarchar(50), @covId nvarchar(2), @visitDate datetime, @clinicCode nvarchar(10),
    @userCode nvarchar(50), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF (SELECT COUNT(*) FROM dbo.CISCas WHERE MBMnumber = @mbmNumber AND CoverID = @covId AND DateVisit = @visitDate AND ClinicCode = @clinicCode) <= 1 RETURN;
    DECLARE @seq nvarchar(20) = (SELECT TOP (1) RTRIM(DupSeqNo) FROM dbo.CISCas
        WHERE MBMnumber = @mbmNumber AND CoverID = @covId AND DateVisit = @visitDate AND ClinicCode = @clinicCode GROUP BY DupSeqNo ORDER BY DupSeqNo DESC);
    IF ISNULL(@seq, '') = ''
    BEGIN
        DECLARE @n int = (SELECT TOP (1) CAST(DupNo AS int) FROM dbo.DupClaimSeq WITH (UPDLOCK, HOLDLOCK) WHERE Code = 'DUP');
        IF @n IS NULL
        BEGIN
            SET @n = 1;
            INSERT INTO dbo.DupClaimSeq (Code, DupNo) VALUES ('DUP', '000001');
        END
        ELSE
        BEGIN
            SET @n += 1;
            UPDATE dbo.DupClaimSeq SET DupNo = RIGHT('000000' + CAST(@n AS varchar(10)), 6) WHERE Code = 'DUP';
        END
        SET @seq = 'DUP' + RIGHT('000000' + CAST(@n AS varchar(10)), 6);
    END
    UPDATE dbo.CISCas SET DupSeqNo = @seq, DupUpdateUser = @userCode, DupUpdateDate = @today, DupAppInd = 'N'
    WHERE MBMnumber = @mbmNumber AND CoverID = @covId AND DateVisit = @visitDate AND ClinicCode = @clinicCode;
END
GO

-- UpdateMBMBalance / UpdateMBMMCDays (HISDB): the clinical limit row of the principal, cover 01 (shared limit B / D) or the cover (C).
CREATE OR ALTER PROCEDURE cis.InvClinic_Balance @mbmNumber nvarchar(50), @limitCover nvarchar(2) = NULL, @adjAnnLmt money, @approved money,
    @deduct money, @full bit, @mcDays money
AS
BEGIN
    SET NOCOUNT ON;
    IF @limitCover IS NULL
        UPDATE cis.Hdb_MBMCLNOthers SET MBMAdjAnnLmt = COALESCE(TRY_CONVERT(money, MBMAdjAnnLmt), 0) + @adjAnnLmt,
            MBMAccumAmt = COALESCE(TRY_CONVERT(money, MBMAccumAmt), 0) + @approved,
            MBMCLNBalLmt = CASE WHEN @full = 1 THEN MBMCLNBalLmt
                                WHEN COALESCE(TRY_CONVERT(money, MBMCLNBalLmt), 0) - @deduct < 0 THEN 0 ELSE COALESCE(TRY_CONVERT(money, MBMCLNBalLmt), 0) - @deduct END
        WHERE MBMNumber = @mbmNumber;
    ELSE
        UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCAdjAnnLmt = COALESCE(TRY_CONVERT(money, MBMCAdjAnnLmt), 0) + @adjAnnLmt,
            MBMCAccumAmt = COALESCE(TRY_CONVERT(money, MBMCAccumAmt), 0) + @approved,
            MBMCCLNBalLmt = CASE WHEN @full = 1 THEN MBMCCLNBalLmt
                                 WHEN COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) - @deduct < 0 THEN 0 ELSE COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) - @deduct END
        WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @limitCover;
    UPDATE cis.Hdb_MBMCLNOthers SET MBMCLNMCDaysBal = CASE WHEN COALESCE(TRY_CONVERT(money, MBMCLNMCDaysBal), 0) > @mcDays
        THEN COALESCE(TRY_CONVERT(money, MBMCLNMCDaysBal), 0) - @mcDays ELSE 0 END
    WHERE MBMNumber = @mbmNumber;
END
GO

-- VerifyDiagnosis + TreatmentReason.ReasonOtherDiag (diagnosis 5 "others" indicator).
CREATE OR ALTER PROCEDURE cis.InvClinic_Diagnosis @code nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) RTRIM(ReasonCode) AS Code, RTRIM(ReasonDesc) AS Description, RTRIM(ReasonOtherDiag) AS OtherIndicator
    FROM dbo.TreatmentReason WHERE ReasonCode = @code;
END
GO
