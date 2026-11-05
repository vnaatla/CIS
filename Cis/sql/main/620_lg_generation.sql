-- LG & Audit > Generate LG (frmLGGeneration): read side. Member search, limits, the member's LGs, one LG, the case tabs and the drop-downs.
SET QUOTED_IDENTIFIER ON;
GO

-- SearchRecords: CLNLG1 (one row per cover) and, for the listed companies, the member's warranty text.
CREATE OR ALTER PROCEDURE cis.Lg_Search @member nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Hdb_CLNLG1 @StrAppend = @member;
    SELECT DISTINCT TOP 1 MBMCLNWarranty FROM cis.Hdb_MBMCLNOthers WHERE MBMNumber LIKE @member + N'%';
END
GO

-- CheckMBMLimit / ValidateLG inputs: GetLGAnnLimitSpecLimit, SP_ValidateLG1, the specialist top-ups, AnnualLimit.SuppLimitStatus and the plan flags.
CREATE OR ALTER PROCEDURE cis.Lg_Member @member nvarchar(30), @cover nvarchar(2), @ic nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @lim TABLE (
        SchMBMNumber nvarchar(50), CoverId nvarchar(10), PolEffDate nvarchar(20), PolExpDate nvarchar(20), PLNCode nvarchar(20), HLTCode nvarchar(10),
        PAYCode nvarchar(10), INSType nvarchar(10), ANNLmtInd nvarchar(5), AnnLimit nvarchar(30), SuppLimit nvarchar(30), SuppLimitStatus nvarchar(5),
        SpecLmtInd nvarchar(5), SpecAnnLimit nvarchar(30), SpecSuppLimit nvarchar(30), SpecSuppLimitStatus nvarchar(5), TopupAmt nvarchar(30),
        GBordxAmt nvarchar(30), GFFS nvarchar(30), GNoLimitDed1 nvarchar(30), GNoLimitDed2 nvarchar(30), GNoLimitDedTotal nvarchar(30), GUtilized nvarchar(30),
        GBalance nvarchar(30), SBordxAmt nvarchar(30), SFFS nvarchar(30), SNoLimitDed1 nvarchar(30), SNoLimitDed1b nvarchar(30), SNoLimitDedTotal nvarchar(30),
        SUtilized nvarchar(30), SBalance nvarchar(30), Error nvarchar(500), OriSpecAnnLimit nvarchar(30), OriAnnLimit nvarchar(30), LGIssuedInd nvarchar(5),
        LGIssuedSuppInd nvarchar(5), LGIssuedAfterLimitExInd nvarchar(5), LGUsingLimit nvarchar(5), LGUsingLimitDesc nvarchar(50), LGHosTypeTInd nvarchar(5),
        PLNGLNo nvarchar(50), HRApproval nvarchar(5));
    INSERT @lim EXEC cis.Ws_GetLGAnnLimitSpecLimit @SchMBMNumber = @member, @CoverId = @cover;
    SELECT * FROM @lim;
    EXEC cis.Hdb_SP_ValidateLG1 @ICNo = @ic, @MBMNumber = @member;
    SELECT SUM(COALESCE(SP_TopupAmt, 0)) AS SpTopup FROM dbo.MBMCISTopupHistory WHERE MBMNumber = @member;
    DECLARE @plan nvarchar(20) = (SELECT TOP 1 RTRIM(PLNCode) FROM @lim);
    SELECT TOP 1 SuppLimitStatus FROM dbo.AnnualLimit WHERE PLNCode = @plan;
    SELECT TOP 1 PLNMaternityInd, PDentalStatus, COALESCE(IsPrinDepGPShared, 0) AS IsPrinDepGPShared FROM dbo.PLN WHERE PLNCode = @plan;
    SELECT MBMASOLimit, MBMOPLimit, MBMIPUtilized, MBMOPUtilized, IsIPOPCombined FROM cis.Hdb_ASOMemberDetails WHERE MBMNumber = @member AND MBMCoverID = @cover;
END
GO

-- spMERCHANTRADE (plan M230 of MERCHANTRADE ASIA SDN BHD): this month's balance.
CREATE OR ALTER PROCEDURE cis.Lg_Merchantrade @member nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.spMERCHANTRADE @mbmno = @member;
END
GO

-- LGListing: LGlst2020 with its filter built from the parameters (the desktop pasted it).
CREATE OR ALTER PROCEDURE cis.Lg_List @member nvarchar(30), @cover nvarchar(2) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @f varchar(2000) = ' AND LG.MBMNumber = ''' + REPLACE(@member, '''', '''''') + ''''
        + CASE WHEN @cover IS NULL THEN '' ELSE ' AND LG.CoverID = ''' + REPLACE(@cover, '''', '''''') + '''' END;
    EXEC dbo.LGlst2020 @StrAppend = @f;
END
GO

-- lstLG_Click: View_PrintLG, pending, pharmacy (clinic type D), the master case of a follow-up, PEMC offline GL data.
CREATE OR ALTER PROCEDURE cis.Lg_Load @lgNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM dbo.View_PrintLG WHERE LGNo = @lgNo;
    SELECT Pending FROM dbo.CISCas WHERE CASNumber = @lgNo;
    SELECT COUNT(*) AS Pharmacy FROM dbo.LG INNER JOIN cis.Hdb_CLN CLN ON CLN.CLNCode = LG.CLNName WHERE LGNo = @lgNo AND CLN.CLNType = 'D';
    SELECT MainCaseNo FROM dbo.tblMedRequestFollowup WHERE FollowupCaseNo = @lgNo;
    SELECT PatientName, IcNo FROM dbo.tblOfflineGL WHERE CasNo = @lgNo;
END
GO

-- IsMedRequestInProgress.
CREATE OR ALTER PROCEDURE cis.Lg_MedRequestInProgress @caseNo nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @r bit = 0;
    EXEC dbo.sp_HCMedRequestInProgress @CaseNo = @caseNo, @RetOut = @r OUTPUT;
    SELECT @r AS InProgress;
END
GO

-- ValidateLG1: the clinic's GL format / address and Sp_Gettemplatedetails_2 (template, dependant template, remarks).
CREATE OR ALTER PROCEDURE cis.Lg_Template @clinic nvarchar(20), @company nvarchar(200), @plan nvarchar(20), @glType nvarchar(100), @ic nvarchar(50),
    @category nvarchar(20), @caseNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @format nvarchar(5) = (SELECT TOP 1 RTRIM(CLNGLFormat) FROM cis.Hdb_CLN WHERE CLNCode = @clinic);
    SELECT TOP 1 CLNGLFormat, COALESCE(CLNName, '') AS CLNName, CLNAdd1, COALESCE(CLNAdd2, '') + ',' + COALESCE(CLNAdd3, '') AS CLNAdd2,
        COALESCE(CLNState, '') + ',' + COALESCE(CLNPostCode, '') AS CLNState, CLNPanel
    FROM cis.Hdb_CLN WHERE CLNCode = @clinic;
    DECLARE @f varchar(10) = CASE WHEN @format = 'M' THEN 'M' ELSE 'H' END;
    EXEC dbo.Sp_Gettemplatedetails_2 @CompanyName = @company, @plncode = @plan, @GLFormat = @f, @GLType = @glType, @IC = @ic, @PEMC = @category, @caseno = @caseNo;
END
GO

-- ListCases / ListSpecialistCases.
CREATE OR ALTER PROCEDURE cis.Lg_Cases @member nvarchar(30), @cover nvarchar(2) = NULL, @specialist bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    IF @specialist = 0
        SELECT CASNumber, MBMCName, ReasonDesc, CoverID, CLMMCOProcFee, DateVisit, MBMCRELCode, CLMTotalActual, CLMBordxAmt, CASStatus, MBMName, MBMNumber, Claimability
        FROM dbo.VIEW_CISMBMInquiry_GPCASE WHERE MBMNumber = @member AND (@cover IS NULL OR CoverID = @cover);
    ELSE
        SELECT CASNumber, MBMCName, ReasonDesc, CoverID, CLMMCOProcFee, DateVisit, MBMCRELCode, CLMTotalActual, CLMBordxAmt, CASStatus, MBMName, MBMNumber, Claimability
        FROM dbo.VIEW_CISMBMInquiry_Spe_CASE WHERE MBMNumber = @member AND (@cover IS NULL OR CoverID = @cover);
END
GO

-- Form_Load: BindGLType, BindDecline, BindDelayreasons.
CREATE OR ALTER PROCEDURE cis.Lg_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT GLID, GLType, GLDesc FROM dbo.Tbl_GLType;
    SELECT ID, DeclineReason FROM dbo.Tbl_DeclineReason WHERE COALESCE(DeclineReason, '') <> '';
    SELECT MessageContent FROM cis.Hdb_MessageComboList
    WHERE MessageModule = 'LG Generation' AND MessageForm = 'Generate Guarantee Letter-CIS' AND MessageTag = 'Reason for delay';
END
GO

-- chkCLN_Click: the group company's clinics.
CREATE OR ALTER PROCEDURE cis.Lg_Clinics @company nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.SearchGrpComCLN @company;
END
GO
