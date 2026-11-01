-- Membership > MBM Adjustment (FrmMBMCLNAdj): look-ups (Medix.bas GetPremium / GetMCO / GetAnnualLimit / GetVisitLimit), history tabs and
-- the endorsement writes (CISDB MBMCISAmendHistory, MBMCISRefund; HISDB MBMCLNOthers, MBMCoveredPersonsCLN, MBMTwo, MBMOthersTwo).
-- Member rows: HISDB SearchCISMemberProc (cis.Hdb_SearchCISMemberProc), group company: cis.MbmTopup_Others, age band: cis.MbmReg_AgeBand.

-- ComboListing: card types (CheckItems CSMD), relationships (HISDB REL).
CREATE OR ALTER PROCEDURE cis.MbmAdj_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(CatName) AS CardType FROM dbo.CheckItems WHERE CatCode = 'CSMD';
    SELECT RTRIM(RELCode) AS Code, RTRIM(RELDescription) AS Name FROM cis.Hdb_REL;
END
GO

-- showMBM: SuppLimitStatus of plan + insured type (any date, first row; blank = A; none = N).
CREATE OR ALTER PROCEDURE cis.MbmAdj_LimitStatus @plnCode nvarchar(10), @insCode nvarchar(5) AS
BEGIN
    SET NOCOUNT ON;
    SELECT ISNULL((SELECT TOP (1) ISNULL(NULLIF(RTRIM(SuppLimitStatus), ''), 'A') FROM dbo.AnnualLimit WHERE PLNCode = @plnCode AND INSCode = @insCode), 'N');
END
GO

-- showMBMCov: the latest SuppLimitStatus on or before the policy date (blank / none = A).
CREATE OR ALTER PROCEDURE cis.MbmAdj_CoverLimitStatus @plnCode nvarchar(10), @insCode nvarchar(5), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT ISNULL((SELECT TOP (1) NULLIF(RTRIM(SuppLimitStatus), '') FROM dbo.AnnualLimit WHERE PLNCode = @plnCode AND INSCode = @insCode
                   AND AnnEffDate <= @effDate ORDER BY AnnEffDate DESC), 'A');
END
GO

-- GetPremium: PLN.PRMInd must be Y for plan + payor; PRM row (first, as the desktop) on or before the date; P = PRMAmount, S = SuppPRMAmt.
-- Found = 0 → "Age Out of Premium Range!".
CREATE OR ALTER PROCEDURE cis.MbmAdj_Premium
    @mbmType char(1), @insCode nvarchar(5), @payCode nvarchar(5), @ageCode nvarchar(5), @plnCode nvarchar(10), @hltCode nvarchar(2), @effDate datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM dbo.PLN WHERE PLNCode = @plnCode AND PAYCode = @payCode AND RTRIM(PRMInd) = 'Y')
    BEGIN SELECT CAST(0 AS money) AS Amount, CAST(1 AS bit) AS Found; RETURN; END
    SELECT TOP (1) CAST(ISNULL(CASE WHEN @mbmType = 'S' THEN SuppPRMAmt ELSE PRMAmount END, 0) AS money) AS Amount, CAST(1 AS bit) AS Found
    FROM dbo.PRM
    WHERE INSCode = @insCode AND HLTCode = @hltCode AND PLNCode = @plnCode AND AGECode = @ageCode AND PRMEffDate <= @effDate
      AND (@hltCode = 'S' OR PAYCode = @payCode);
    IF @@ROWCOUNT = 0 SELECT CAST(0 AS money) AS Amount, CAST(0 AS bit) AS Found;
END
GO

-- GetMCO: PLN.MCOInd must be Y; the plan's latest MCO row, else the insured type's latest; column by adult / child and P / S. Found = 0 → message.
CREATE OR ALTER PROCEDURE cis.MbmAdj_Mco
    @mbmType char(1), @insCode nvarchar(5), @hltCode nvarchar(2), @adultChild char(1), @effDate datetime, @plnCode nvarchar(10), @payCode nvarchar(5)
AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM dbo.PLN WHERE PLNCode = @plnCode AND PAYCode = @payCode AND RTRIM(MCOInd) = 'Y')
    BEGIN SELECT CAST(0 AS money) AS Amount, CAST(1 AS bit) AS Found; RETURN; END
    DECLARE @amt money, @found bit = 0;
    SELECT TOP (1) @amt = CASE WHEN @adultChild = 'A' THEN CASE WHEN @mbmType = 'S' THEN MCOSuppAdultAmt ELSE MCOAmtAdult END
                               ELSE CASE WHEN @mbmType = 'S' THEN MCOSuppChildAmt ELSE MCOAmtChild END END, @found = 1
    FROM dbo.MCO WHERE INSCode = @insCode AND HLTCode = @hltCode AND PLNCode = @plnCode AND MCOEffDate <= @effDate ORDER BY MCOEffDate DESC;
    IF @found = 0
        SELECT TOP (1) @amt = CASE WHEN @adultChild = 'A' THEN CASE WHEN @mbmType = 'S' THEN MCOSuppAdultAmt ELSE MCOAmtAdult END
                                   ELSE CASE WHEN @mbmType = 'S' THEN MCOSuppChildAmt ELSE MCOAmtChild END END, @found = 1
        FROM dbo.MCO WHERE INSCode = @insCode AND HLTCode = @hltCode AND MCOEffDate <= @effDate ORDER BY MCOEffDate DESC;
    SELECT CAST(ISNULL(@amt, 0) AS money) AS Amount, @found AS Found;
END
GO

-- GetAnnualLimit / GetVisitLimit (health N): Full = 1 is the desktop's "FULL"; Status = the AnnualLimit SuppLimitStatus it returns.
CREATE OR ALTER PROCEDURE cis.MbmAdj_Limits
    @mbmType char(1), @insCode nvarchar(5), @plnCode nvarchar(10), @effDate datetime, @payCode nvarchar(5), @limitStatus varchar(1)
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @annInd varchar(1), @annSupp varchar(1), @shared bit, @visInd varchar(1);
    SELECT TOP (1) @annInd = RTRIM(ANNLmtInd), @annSupp = RTRIM(ANNLmtIndSupp), @shared = ISNULL(IsPrinDepGPShared, 0), @visInd = RTRIM(VISLmtInd)
    FROM dbo.PLN WHERE PLNCode = @plnCode AND PAYCode = @payCode;
    DECLARE @annFull bit = CASE WHEN (@mbmType = 'P' AND @annInd = 'N') OR (@mbmType = 'S' AND ISNULL(@limitStatus, '') <> 'B' AND @annSupp = 'N') THEN 1 ELSE 0 END;
    DECLARE @ann money = 0, @annOk bit = 1, @status varchar(1) = @limitStatus;
    IF @annFull = 0
    BEGIN
        DECLARE @annLimit money, @supp money, @aso money, @st varchar(1), @hasRow bit = 0;
        SELECT TOP (1) @annLimit = AnnLimit, @supp = SuppLimit, @aso = ISNULL(ASOLimit, 0), @st = RTRIM(SuppLimitStatus), @hasRow = 1
        FROM dbo.AnnualLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND HLTCode = 'N' AND AnnEffDate <= @effDate ORDER BY AnnEffDate DESC;
        IF @hasRow = 1
        BEGIN
            SET @status = @st;
            IF ISNULL(@st, '') = '' SELECT @ann = 0, @annOk = 0;
            ELSE IF @st = 'D' SET @annFull = 1;
            ELSE SET @ann = ISNULL(@annLimit, 0) + @aso;
            IF @mbmType = 'S'
                IF @shared = 1 SELECT @ann = ISNULL(@annLimit, 0) + @aso, @annFull = 0;
                ELSE IF ISNULL(@st, '') <> 'A' SELECT @ann = ISNULL(@supp, 0), @annFull = 0;
        END
    END
    DECLARE @visFull bit = CASE WHEN @visInd = 'N' THEN 1 ELSE 0 END, @visit money = 0;
    IF @visFull = 0
        SELECT TOP (1) @visit = CASE WHEN @mbmType = 'S' AND ISNULL(RTRIM(SuppVisitLimitStatus), '') <> 'A' THEN ISNULL(SuppVisitLimit, 0) ELSE ISNULL(VisitLimit, 0) END
        FROM dbo.VisitLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND HLTCode = 'N' AND VisitEffDate <= @effDate ORDER BY VisitEffDate DESC;
    SELECT @ann AS AnnualLimit, @annFull AS AnnualFull, @annOk AS AnnualOk, @visit AS VisitLimit, @visFull AS VisitFull, ISNULL(@status, '') AS Status;
END
GO

-- showMBMCov: refund figures of a cancelled membership.
CREATE OR ALTER PROCEDURE cis.MbmAdj_Refunds @mbmNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(MBMCoveredID) AS CoverId, MBMPremReduction AS PremReduction, MBMPremEffective AS PremEffective, MBMMCORefund AS McoRefund,
           MBMMCOEffective AS McoEffective
    FROM dbo.MBMCISRefund WHERE MBMNumber = @mbmNumber;
END
GO

-- CmdSave_Click (C): an outstanding case blocks the cancellation.
CREATE OR ALTER PROCEDURE cis.MbmAdj_OpenCase @mbmNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.CISCAS WHERE MBMnumber = @mbmNumber AND CASStatus = 'O') THEN 1 ELSE 0 END AS bit);
END
GO

-- ADJListing: CISDB MBMCISAmendHistory.
CREATE OR ALTER PROCEDURE cis.MbmAdj_AmendHistory @mbmNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT MBMAEndtBordxDate AS BordxDate, MBMAmendDate AS AmendDate, MBMAmendEdtType AS EndtType, MBMAEndtEffDate AS EndtEffDate, RTRIM(MBMAmendUser) AS AmendUser
    FROM dbo.MBMCISAmendHistory WHERE MBMNumber = @mbmNumber ORDER BY MBMAmendID;
END
GO

-- Write2MBMnHistory: history row from the current MBMCLNOthers, then the endorsement on MBMCLNOthers.
CREATE OR ALTER PROCEDURE cis.MbmAdj_Endorse
    @mbmNumber nvarchar(30), @endtType char(1), @status nvarchar(1), @bordxDate datetime, @endtEffDate datetime, @user nvarchar(50), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.MBMCISAmendHistory (MBMNumber, MBMPolicyNo, MBMEffDate, MBMExpiryDate, MBMAmendEdtType, MBMAEndtBordxDate, MBMAEndtEffDate, MBMAmendUser, MBMAmendDate)
    SELECT TOP (1) MBMNumber, MBMPolNumber, MBMPolEffDate, MBMPolExpDate, @endtType, @bordxDate, @endtEffDate, @user, @today
    FROM cis.Hdb_MBMCLNOthers WHERE MBMNumber = @mbmNumber;
    UPDATE cis.Hdb_MBMCLNOthers
    SET MBMCLNLastEndtStatus = @endtType, MBMCLNStatus = @status, MBMCLNEndtEffDate = @endtEffDate, MBMCancelDate = @endtEffDate, MBMCLNEndtBordxDate = @bordxDate
    WHERE MBMNumber = @mbmNumber;
END
GO

-- WriteAdjStatus2Cov: every cover's status, or one cover's (K, with today's cancel date).
CREATE OR ALTER PROCEDURE cis.MbmAdj_CoverStatus @mbmNumber nvarchar(30), @status nvarchar(1), @covId nvarchar(2) = NULL, @today datetime AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Hdb_MBMCoveredPersonsCLN
    SET MBMCCLNStatus = @status, MBMCCancelDate = CASE WHEN @covId IS NULL THEN MBMCCancelDate ELSE @today END
    WHERE MBMCNumber = @mbmNumber AND (@covId IS NULL OR MBMCCoverID = @covId);
END
GO

-- MemberMCORefund: replace the cover's MBMCISRefund row.
CREATE OR ALTER PROCEDURE cis.MbmAdj_Refund
    @mbmNumber nvarchar(30), @covId nvarchar(2), @premEffective money, @premReduction money, @mcoEffective money, @mcoRefund money
AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.MBMCISRefund WHERE MBMNumber = @mbmNumber AND MBMCoveredID = @covId;
    INSERT INTO dbo.MBMCISRefund (MBMNumber, MBMCoveredID, MBMPremEffective, MBMPremReduction, MBMMCOEffective, MBMMCORefund)
    VALUES (@mbmNumber, @covId, @premEffective, @premReduction, @mcoEffective, @mcoRefund);
END
GO

-- Amendment (A): SaveCLNOthers + SaveMBMTwo + SaveMBMOthersTwo.
CREATE OR ALTER PROCEDURE cis.MbmAdj_SavePrincipal
    @mbmNumber nvarchar(30), @policyNo nvarchar(100), @effDate datetime = NULL, @expDate datetime = NULL, @remarks nvarchar(max), @cardType nvarchar(50),
    @mcoFee money, @premium money, @dataStatus nvarchar(1), @status nvarchar(1), @warranty nvarchar(max), @plnCode nvarchar(10),
    @email nvarchar(200), @mobile nvarchar(50), @bankName nvarchar(200), @bankAccount nvarchar(100)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Hdb_MBMCLNOthers
    SET MBMPolNumber = @policyNo, MBMPolEffDate = ISNULL(@effDate, MBMPolEffDate), MBMPolExpDate = ISNULL(@expDate, MBMPolExpDate), MBMRemarks = @remarks,
        MBMCardType = @cardType, MBMCLNMCOFee = @mcoFee, MBMCLNPremium = @premium, MBMCLNDataStatus = @dataStatus, MBMCLNStatus = @status,
        MBMCLNWarranty = @warranty, PLNType = @plnCode
    WHERE MBMNumber = @mbmNumber;
    UPDATE cis.Hdb_MBMTwo SET MBMEmail = @email, MBMTelNoM = @mobile WHERE MBMNumber = @mbmNumber;
    UPDATE cis.Hdb_MBMOthersTwo SET MBMBankName = @bankName, MBMBankACNo = @bankAccount WHERE MBMNumber = @mbmNumber;
END
GO

-- SaveCisCovPersons: one cover.
CREATE OR ALTER PROCEDURE cis.MbmAdj_SaveCover
    @mbmNumber nvarchar(30), @covId nvarchar(2), @remarks nvarchar(max), @mco money, @premium money, @warranty nvarchar(max), @cardType nvarchar(50),
    @plnCode nvarchar(10), @effDate datetime = NULL, @expDate datetime = NULL
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Hdb_MBMCoveredPersonsCLN
    SET MBMCCLNRemarks = @remarks, MBMCCLNMCO = @mco, MBMCCLNPrem = @premium, MBMCCLNWarranty = @warranty, MBMCCardType = @cardType,
        MBMCCLNPLNCode = @plnCode, MBMCCLNEffDate = @effDate, MBMCCLNExpDate = @expDate
    WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @covId;
END
GO
