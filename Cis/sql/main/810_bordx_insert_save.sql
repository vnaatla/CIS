-- Claim Bordx > Insert Worksheet (frmWSBordx): look-ups, the BTBG plan check and CmdBordx_Click's save.
SET QUOTED_IDENTIFIER ON;
GO

-- PayorComboListing (GRPCoName|GRPCoCode), the treatment codes and the panel clinics (ChkClinic).
CREATE OR ALTER PROCEDURE cis.BordxWs_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoCode) AS Code, RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany;
    SELECT RTRIM(ReasonCode) AS Code, RTRIM(ReasonDesc) AS Name FROM dbo.TreatmentReason;
    SELECT RTRIM(CLNCode) AS Code, RTRIM(CLNName) AS Name FROM cis.Hdb_CLN WHERE (CLNPanel = 'Y' OR CLNPanel = 'C') AND LEN(CLNName) > 0;
END
GO

-- CboPayor_LostFocus: SearchDepartment / SearchSubsidiary_bygrp of the group company name.
CREATE OR ALTER PROCEDURE cis.BordxWs_Company @company nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @d TABLE (MBMDepartment nvarchar(200));
    DECLARE @s TABLE (MBMGRPCompany nvarchar(200), MBMGRPSubsidiary nvarchar(200));
    INSERT @d EXEC HISDB.dbo.SearchDepartment @company;
    INSERT @s EXEC HISDB.dbo.SearchSubsidiary_bygrp @company;
    SELECT DISTINCT LTRIM(RTRIM(MBMDepartment)) AS Name FROM @d WHERE COALESCE(LTRIM(RTRIM(MBMDepartment)), '') <> '';
    SELECT DISTINCT LTRIM(RTRIM(MBMGRPSubsidiary)) AS Name FROM @s WHERE COALESCE(LTRIM(RTRIM(MBMGRPSubsidiary)), '') <> '';
END
GO

-- CheckIsBTBG: sp_GetPLNBTBGInd.
CREATE OR ALTER PROCEDURE cis.BordxWs_IsBtbg @plnCode nvarchar(8), @isBtbg bit OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    SET @isBtbg = 0;
    EXEC dbo.sp_GetPLNBTBGInd @PLNCode = @plnCode, @RetOut = @isBtbg OUTPUT;
END
GO

-- GenerateCLMBordxNo / GenerateCLMBordxNoVOID: K + payor + the next 5-digit number of CLAIM_BordxSEQ (KIV + number for void), locked for the save.
CREATE OR ALTER PROCEDURE cis.BordxWs_NextNo @payCode nvarchar(4), @bordxNo nvarchar(20) OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @n nvarchar(10);
    IF NOT EXISTS (SELECT 1 FROM CISWorksheet.dbo.CLAIM_BordxSEQ WITH (UPDLOCK, HOLDLOCK) WHERE PAYCode = @payCode)
    BEGIN
        INSERT CISWorksheet.dbo.CLAIM_BordxSEQ (HLTCode, PAYCode, CLMBordxNo)
        VALUES (CASE WHEN @payCode = 'IV' THEN NULL ELSE (SELECT TOP 1 HLTCode FROM dbo.PAY WHERE PAYCode = @payCode) END, @payCode, '00001');
        SET @n = '00001';
    END
    ELSE
        UPDATE TOP (1) CISWorksheet.dbo.CLAIM_BordxSEQ
        SET @n = CLMBordxNo = RIGHT('00000' + CAST(CAST(CLMBordxNo AS int) + 1 AS varchar(10)), CASE WHEN LEN(CAST(CAST(CLMBordxNo AS int) + 1 AS varchar(10))) > 5 THEN LEN(CAST(CAST(CLMBordxNo AS int) + 1 AS varchar(10))) ELSE 5 END)
        WHERE PAYCode = @payCode;
    SET @bordxNo = CASE WHEN @payCode = 'IV' THEN 'KIV' ELSE 'K' + @payCode END + @n;
END
GO

-- UpdateRecord / UpdateRecordvoid / UpdateBTBGRecord for one case; returns the amount the desktop adds to the bordx total.
-- @mode 0 normal (CLMFinalApprove), 1 void (+ Payment2MBM 'VOID' row for a member reimbursement), 2 BTBG (CLMBTBGAmt).
CREATE OR ALTER PROCEDURE cis.BordxWs_UpdateCase @casNumber nvarchar(30), @bordxNo nvarchar(20), @mode int, @user nvarchar(50), @amount money OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @today datetime = CAST(GETDATE() AS date);
    SELECT @amount = CASE WHEN @mode = 2 THEN COALESCE(CLMBTBGAmt, 0) ELSE CLMFinalApprove END FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber;
    IF @mode = 2
    BEGIN
        UPDATE CISWorksheet.dbo.Claim SET BTBGBordxRefNo = @bordxNo, BTBGBordxStatus = 1, CLMPending = 'N' WHERE CASNumber = @casNumber;
        UPDATE CISWorksheet.dbo.ClaimInvoice SET BTBGINVBordxRefNo = @bordxNo, BTBGINVBordxed = 1 WHERE CASNumber = @casNumber;
    END
    ELSE
    BEGIN
        UPDATE CISWorksheet.dbo.Claim SET BordxRefNo = @bordxNo, ClaimBordxStatus = 1, CLMPending = 'N', CLMBordxMonth = FORMAT(@today, 'MM'),
            CLMBordxYear = FORMAT(@today, 'yyyy') WHERE CASNumber = @casNumber;
        UPDATE CISWorksheet.dbo.ClaimInvoice SET INVBordxRefNo = @bordxNo, INVBordxed = 1 WHERE CASNumber = @casNumber;
    END
    UPDATE dbo.CISCas SET Pending = 'N' WHERE CASNumber = @casNumber;
    IF @mode = 1 AND EXISTS (SELECT 1 FROM CISWorksheet.dbo.ClaimInvoice WHERE CASNumber = @casNumber AND RTRIM(INVPayTo) = 'M')
        INSERT CISPayment.dbo.Payment2MBM (PVNo, PVDate, PVChequeDate, PVChequeNo, CASNumber, Pay2MBMLastUpdateDate, Pay2MBMLastUpdateUser, PAYAccPeriodMth,
            PAYAccPeriodYr, PVPayMBMStatus)
        VALUES ('VOID', @today, @today, 'VOID', @casNumber, @today, @user, MONTH(@today), YEAR(@today), 'F');
END
GO

-- SaveCLMBordx.
CREATE OR ALTER PROCEDURE cis.BordxWs_SaveBordx @bordxNo nvarchar(20), @payCode nvarchar(4), @user nvarchar(50), @amount money, @claims int,
    @company nvarchar(200), @rptType nvarchar(2), @subsidiary nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISWorksheet.dbo.Bordx (BordxPAYCode, BordxRefNo, BordxDate, BordxBy, BordxClaimAmt, BordxNoClaims, GroupCompany, RptTypeByDept, BordxSubsidary)
    VALUES (@payCode, @bordxNo, CAST(GETDATE() AS date), @user, @amount, @claims, @company, @rptType, @subsidiary);
END
GO
