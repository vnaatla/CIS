-- Maintenance > Plan (frmMaintenancePlan) tab Plan Code: CISDB PLN, VIEW_PLNListing, PLNSeq.

-- PLNComboListing / Check2_Click: health types, CIS and HIS payors, clinic contracts (CheckItems CSMD), group companies (not ZD01).
CREATE OR ALTER PROCEDURE cis.PlanCode_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(HLTCode) AS Code, ISNULL(HLTDesc, N'') AS Name FROM dbo.HLT ORDER BY HLTCode;
    SELECT RTRIM(PAYCode) AS Code, ISNULL(PAYCompanyName, N'') AS Name FROM dbo.PAY ORDER BY PAYCode;
    SELECT RTRIM(PAYCode) AS Code, ISNULL(PAYCompanyName, N'') AS Name FROM cis.Hdb_PAY ORDER BY PAYCode;
    SELECT RTRIM(CatName) AS Code, RTRIM(CatName) AS Name FROM dbo.CheckItems WHERE CatCode = 'CSMD' ORDER BY CatName;
    SELECT RTRIM(GRPCoCode) AS Code, RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoCode <> 'ZD01'
    GROUP BY GRPCoCode, GRPCoName ORDER BY GRPCoCode, GRPCoName;
END
GO

-- cboPLNHLTCode_Click: CIS payors of the health type.
CREATE OR ALTER PROCEDURE cis.PlanCode_Payors @hltCode nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    SELECT UPPER(RTRIM(PAYCode)) AS Code, UPPER(ISNULL(PAYCompanyName, N'')) AS Name FROM dbo.PAY WHERE HLTCode = @hltCode ORDER BY PAYCode;
END
GO

-- SearchPLN → PLNListing (VIEW_PLNListing); the indicator filters the view lacks are read from PLN.
CREATE OR ALTER PROCEDURE cis.PlanCode_Search
    @plnCode nvarchar(10) = NULL, @payCode nvarchar(2) = NULL, @hisPayCode nvarchar(2) = NULL, @hltCode nchar(1) = NULL, @grpCode nvarchar(4) = NULL,
    @clnContract nvarchar(50) = NULL, @annLmt nchar(1) = NULL, @visLmt nchar(1) = NULL, @prm nchar(1) = NULL, @mco nchar(1) = NULL,
    @pDentalStatus nchar(1) = NULL, @pDentalLimit nchar(1) = NULL, @pSpecStatus nchar(1) = NULL, @pSpecLimit nchar(1) = NULL,
    @pMepStatus nchar(1) = NULL, @sDentalStatus nchar(1) = NULL, @sDentalLimit nchar(1) = NULL, @sSpecStatus nchar(1) = NULL,
    @sSpecLimit nchar(1) = NULL, @sMepStatus nchar(1) = NULL, @sAnnLmt nchar(1) = NULL, @sVisLmt nchar(1) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT v.HLTCode, v.PLNCode, v.PLNDesc, RTRIM(v.PAYCode) AS PAYCode, RTRIM(v.HISPAYCode) AS HISPAYCode, v.CoGRPCode, v.CoGRPName, v.ClnContract,
           v.Insured, v.HRApproval, v.PrinSPLimit, v.PrinGPLimit, v.DepSPLimit, v.DepGPLimit, v.OPFamilyStatus, v.PLNSpecialRemark, v.isopdependent
    FROM dbo.VIEW_PLNListing v
    INNER JOIN dbo.PLN p ON p.PLNCode = v.PLNCode
    WHERE (@plnCode IS NULL OR v.PLNCode LIKE N'%' + @plnCode + N'%') AND (@payCode IS NULL OR v.PAYCode = @payCode)
      AND (@hisPayCode IS NULL OR v.HISPAYCode = @hisPayCode) AND (@hltCode IS NULL OR v.HLTCode = @hltCode)
      AND (@grpCode IS NULL OR v.CoGRPCode = @grpCode) AND (@clnContract IS NULL OR v.ClnContract = @clnContract)
      AND (@annLmt IS NULL OR p.ANNLmtInd = @annLmt) AND (@visLmt IS NULL OR p.VISLmtInd = @visLmt) AND (@prm IS NULL OR p.PRMInd = @prm)
      AND (@mco IS NULL OR p.MCOInd = @mco) AND (@pDentalStatus IS NULL OR p.PDentalStatus = @pDentalStatus)
      AND (@pDentalLimit IS NULL OR p.PDentalLimit = @pDentalLimit) AND (@pSpecStatus IS NULL OR p.PSpecStatus = @pSpecStatus)
      AND (@pSpecLimit IS NULL OR p.PSpecLimit = @pSpecLimit) AND (@pMepStatus IS NULL OR p.PMEPStatus = @pMepStatus)
      AND (@sDentalStatus IS NULL OR p.SDentalStatus = @sDentalStatus) AND (@sDentalLimit IS NULL OR p.SDentalLimit = @sDentalLimit)
      AND (@sSpecStatus IS NULL OR p.SSpecStatus = @sSpecStatus) AND (@sSpecLimit IS NULL OR p.SSpecLimit = @sSpecLimit)
      AND (@sMepStatus IS NULL OR p.SMEPStatus = @sMepStatus) AND (@sAnnLmt IS NULL OR p.ANNLmtIndSupp = @sAnnLmt)
      AND (@sVisLmt IS NULL OR p.VISLmtIndSupp = @sVisLmt)
    ORDER BY v.PLNCode;
END
GO

-- lstPlanList_Click: the plan's columns.
CREATE OR ALTER PROCEDURE cis.PlanCode_Get @code nvarchar(4) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PLNCode) AS PlnCode, PLNDesc AS PlnDesc, RTRIM(HLTCode) AS HltCode, RTRIM(PAYCode) AS PayCode, RTRIM(HISPAYCode) AS HisPayCode,
           RTRIM(GRPCompany) AS GrpCompany, ClnContract, RTRIM(ANNLmtInd) AS AnnLmt, RTRIM(VISLmtInd) AS VisLmt, RTRIM(PRMInd) AS Prm,
           RTRIM(MCOInd) AS Mco, RTRIM(DCStatus) AS DcStatus, RTRIM(PDentalStatus) AS PDentalStatus, RTRIM(PDentalLimit) AS PDentalLimit,
           RTRIM(PSpecStatus) AS PSpecStatus, RTRIM(PSpecLimit) AS PSpecLimit, RTRIM(PMEPStatus) AS PMepStatus, RTRIM(SDentalStatus) AS SDentalStatus,
           RTRIM(SDentalLimit) AS SDentalLimit, RTRIM(SSpecStatus) AS SSpecStatus, RTRIM(SSpecLimit) AS SSpecLimit, RTRIM(SMEPStatus) AS SMepStatus,
           RTRIM(ANNLmtIndSupp) AS SAnnLmt, RTRIM(VISLmtIndSupp) AS SVisLmt, RTRIM(Insured) AS Insured, RTRIM(HRApproval) AS HrApproval,
           OPFamilyStatus AS OpFamilyStatus, CAST(PrinSPLimit AS nvarchar(30)) AS PrinSpLimit, CAST(PrinGPLimit AS nvarchar(30)) AS PrinGpLimit,
           CAST(DepSPLimit AS nvarchar(30)) AS DepSpLimit, CAST(DepGPLimit AS nvarchar(30)) AS DepGpLimit, PLNSpecialRemark AS SpecialRemark,
           RTRIM(isOPDependent) AS OpDependent
    FROM dbo.PLN WHERE PLNCode = @code;
END
GO

-- SavePLN (N): "Record Duplicated!" when the Grp Comp. text is a group company name.
CREATE OR ALTER PROCEDURE cis.PlanCode_GroupNameExists @name nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.GRPCompany WHERE GRPCoName = @name) THEN 1 ELSE 0 END AS bit);
END
GO

-- GeneratePLNSeq (Medix.bas): PLNSeq row of the letter, +1 "000" (a new letter starts at 001); codes already in PLN are skipped.
CREATE OR ALTER PROCEDURE cis.PlanCode_NextCode @letter nchar(1) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seq int = (SELECT TOP (1) TRY_CAST(PLNNoSeq AS int) FROM dbo.PLNSeq WITH (UPDLOCK, HOLDLOCK) WHERE PLNSeqCode = @letter);
    DECLARE @new bit = CASE WHEN @seq IS NULL THEN 1 ELSE 0 END;
    SET @seq = ISNULL(@seq, 0) + 1;
    WHILE EXISTS (SELECT 1 FROM dbo.PLN WHERE PLNCode = @letter + RIGHT(N'000' + CAST(@seq AS nvarchar(10)), 3)) SET @seq += 1;
    IF @seq > 999 THROW 51000, N'No plan codes left for this letter.', 1;
    DECLARE @text nvarchar(3) = RIGHT(N'000' + CAST(@seq AS nvarchar(10)), 3);
    IF @new = 1 INSERT INTO dbo.PLNSeq (PLNSeqCode, PLNNoSeq) VALUES (@letter, @text);
    ELSE UPDATE dbo.PLNSeq SET PLNNoSeq = @text WHERE PLNSeqCode = @letter;
    SELECT @letter + @text;
END
GO

-- SavePLN (N): one PLN row per plan; dental / spec / MEP indicators always N, as the desktop wrote them.
CREATE OR ALTER PROCEDURE cis.PlanCode_Insert
    @code nvarchar(4), @desc nvarchar(200), @hltCode nvarchar(2), @payCode nvarchar(2), @hisPayCode nvarchar(2), @grpCompany nvarchar(200),
    @annLmt nchar(1), @visLmt nchar(1), @sAnnLmt nchar(1), @sVisLmt nchar(1), @prm nchar(1), @mco nchar(1), @clnContract nvarchar(50),
    @dcStatus nchar(1), @insured nchar(1), @hrApproval nchar(1), @opFamilyStatus nvarchar(50), @prinSpLimit float, @prinGpLimit float,
    @depSpLimit float, @depGpLimit float, @specialRemark varchar(1500), @opDependent nchar(1), @user nvarchar(100), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.PLN (PLNCode, PLNDesc, HLTCode, PAYCode, HISPAYCode, GRPCompany, ANNLmtInd, VISLmtInd, ANNLmtIndSupp, VISLmtIndSupp, PRMInd, MCOInd,
                         ClnContract, DCStatus, PDentalStatus, PSpecStatus, PMEPStatus, PDentalLimit, PSpecLimit, SDentalStatus, SSpecStatus, SMEPStatus,
                         SDentalLimit, SSpecLimit, Insured, HRApproval, OPFamilyStatus, PrinSPLimit, PrinGPLimit, DepSPLimit, DepGPLimit, PLNSpecialRemark,
                         isOPDependent, PLNLastUpdateUser, PLNLastUpdateDate)
    VALUES (@code, @desc, @hltCode, @payCode, @hisPayCode, @grpCompany, @annLmt, @visLmt, @sAnnLmt, @sVisLmt, @prm, @mco, @clnContract, @dcStatus,
            'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', @insured, @hrApproval, @opFamilyStatus, @prinSpLimit, @prinGpLimit, @depSpLimit,
            @depGpLimit, @specialRemark, @opDependent, @user, @today);
END
GO

-- DeletePLN: refused while membership (HISDB MBMCLNOthers.PLNType) uses the plan.
CREATE OR ALTER PROCEDURE cis.PlanCode_InUse @code nvarchar(4) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM cis.Hdb_MBMCLNOthers WHERE PLNType = @code) THEN 1 ELSE 0 END AS bit);
END
GO

CREATE OR ALTER PROCEDURE cis.PlanCode_Delete @code nvarchar(4) AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.PLN WHERE PLNCode = @code;
    SELECT @@ROWCOUNT;
END
GO
