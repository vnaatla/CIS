-- Maintenance > Plan (frmMaintenancePlan) tab Annual Limit: CISDB AnnualLimit, VIEW_ANN_PAY_PLN.

-- AnnualComboListing: plans, insured types, health types, payors, distinct PLN.GRPCompany.
CREATE OR ALTER PROCEDURE cis.PlanAnnual_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PLNCode) AS Code, ISNULL(PLNDesc, N'') AS Name FROM dbo.PLN ORDER BY PLNCode;
    SELECT RTRIM(INSCode) AS Code, ISNULL(INSDesc, N'') AS Name FROM dbo.INS ORDER BY INSCode;
    SELECT RTRIM(HLTCode) AS Code, ISNULL(HLTDesc, N'') AS Name FROM dbo.HLT ORDER BY HLTCode;
    SELECT RTRIM(PAYCode) AS Code, ISNULL(PAYCompanyName, N'') AS Name FROM dbo.PAY ORDER BY PAYCode;
    SELECT DISTINCT RTRIM(GRPCompany) AS Code, RTRIM(GRPCompany) AS Name FROM dbo.PLN WHERE LEN(GRPCompany) > 0 ORDER BY Code;
END
GO

-- SearchANN → AnnualListing (VIEW_ANN_PAY_PLN).
CREATE OR ALTER PROCEDURE cis.PlanAnnual_Search
    @plnCode nvarchar(4) = NULL, @hltCode nvarchar(4) = NULL, @insCode nvarchar(4) = NULL, @payCode nvarchar(4) = NULL, @grpCompany nvarchar(4) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT HLTCode, AnnualIndex, PLNCode, INSCode, AnnLimit, PDentalAnnLimit, PSpecAnnLimit, PMEPAnnLimit, SuppLimit, SDentalAnnLimit, SSpecAnnLimit,
           SMEPAnnLimit, AnnEffDate, AnnLimitVersion, PAYCode, GRPCompany, SuppLimitStatus
    FROM dbo.VIEW_ANN_PAY_PLN
    WHERE (@plnCode IS NULL OR PLNCode = @plnCode) AND (@hltCode IS NULL OR HLTCode = @hltCode) AND (@insCode IS NULL OR INSCode = @insCode)
      AND (@payCode IS NULL OR PAYCode = @payCode) AND (@grpCompany IS NULL OR GRPCompany LIKE @grpCompany + N'%')
    ORDER BY AnnualIndex;
END
GO

-- SaveANN duplicate check (Add only): insured + health + plan + version + effective date.
CREATE OR ALTER PROCEDURE cis.PlanAnnual_Exists @insCode nvarchar(2), @hltCode nvarchar(2), @plnCode nvarchar(4), @version int, @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.AnnualLimit WHERE INSCode = @insCode AND HLTCode = @hltCode AND PLNCode = @plnCode
                                  AND AnnLimitVersion = @version AND AnnEffDate = @effDate) THEN 1 ELSE 0 END AS bit);
END
GO

-- SaveANN: Add inserts; Edit updates the selected AnnualIndex. Status A leaves the supplementary columns as they are (NULL on Add).
CREATE OR ALTER PROCEDURE cis.PlanAnnual_Save
    @index int = NULL, @plnCode nvarchar(4), @insCode nvarchar(2), @hltCode nvarchar(2), @annLimit money, @pDental money, @pSpec money, @pMep money,
    @status varchar(1), @suppLimit money = NULL, @sDental money = NULL, @sSpec money = NULL, @sMep money = NULL, @effDate datetime,
    @version int, @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF @index IS NULL
        INSERT INTO dbo.AnnualLimit (PLNCode, INSCode, HLTCode, AnnLimit, PDentalAnnLimit, PSpecAnnLimit, PMEPAnnLimit, SuppLimitStatus, SuppLimit,
                                     SDentalAnnLimit, SSpecAnnLimit, SMEPAnnLimit, AnnEffDate, AnnLimitVersion, AnnLastUpdateUser, AnnLastUpdateDate)
        VALUES (@plnCode, @insCode, @hltCode, @annLimit, @pDental, @pSpec, @pMep, @status,
                CASE WHEN @status = 'A' THEN NULL ELSE @suppLimit END, CASE WHEN @status = 'A' THEN NULL ELSE @sDental END,
                CASE WHEN @status = 'A' THEN NULL ELSE @sSpec END, CASE WHEN @status = 'A' THEN NULL ELSE @sMep END, @effDate, @version, @user, @today);
    ELSE
        UPDATE dbo.AnnualLimit
        SET PLNCode = @plnCode, INSCode = @insCode, HLTCode = @hltCode, AnnLimit = @annLimit, PDentalAnnLimit = @pDental, PSpecAnnLimit = @pSpec,
            PMEPAnnLimit = @pMep, SuppLimitStatus = @status,
            SuppLimit = CASE WHEN @status = 'A' THEN SuppLimit ELSE @suppLimit END, SDentalAnnLimit = CASE WHEN @status = 'A' THEN SDentalAnnLimit ELSE @sDental END,
            SSpecAnnLimit = CASE WHEN @status = 'A' THEN SSpecAnnLimit ELSE @sSpec END, SMEPAnnLimit = CASE WHEN @status = 'A' THEN SMEPAnnLimit ELSE @sMep END,
            AnnEffDate = @effDate, AnnLimitVersion = @version, AnnLastUpdateUser = @user, AnnLastUpdateDate = @today
        WHERE AnnualIndex = @index;
    SELECT @@ROWCOUNT;
END
GO

-- DeleteANN: refused while MBM (members) uses the plan; runs only where CISDB has MBM.
CREATE OR ALTER PROCEDURE cis.PlanAnnual_InUse @plnCode nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    IF OBJECT_ID(N'dbo.MBM') IS NULL BEGIN SELECT CAST(0 AS bit); RETURN; END
    EXEC sp_executesql N'SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.MBM WHERE PLNCode = @plnCode) THEN 1 ELSE 0 END AS bit)',
        N'@plnCode nvarchar(10)', @plnCode = @plnCode;
END
GO

CREATE OR ALTER PROCEDURE cis.PlanAnnual_Delete @index int AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.AnnualLimit WHERE AnnualIndex = @index;
    SELECT @@ROWCOUNT;
END
GO
