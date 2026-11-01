-- Maintenance > Plan (frmMaintenancePlan) tab Premium: CISDB PRM (one row per age band 01–07).

-- PRMComboListing: insured types, health types, payors, age codes, plans.
CREATE OR ALTER PROCEDURE cis.PlanPremium_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(INSCode) AS Code, ISNULL(INSDesc, N'') AS Name FROM dbo.INS ORDER BY INSCode;
    SELECT RTRIM(HLTCode) AS Code, ISNULL(HLTDesc, N'') AS Name FROM dbo.HLT ORDER BY HLTCode;
    SELECT RTRIM(PAYCode) AS Code, ISNULL(PAYCompanyName, N'') AS Name FROM dbo.PAY ORDER BY PAYCode;
    SELECT RTRIM(AgeCode) AS Code, ISNULL(AGEDesc, N'') AS Name FROM dbo.AGE ORDER BY AgeCode;
    SELECT RTRIM(PLNCode) AS Code, ISNULL(PLNDesc, N'') AS Name FROM dbo.PLN ORDER BY PLNCode;
END
GO

-- SearchPRM → PRMListing.
CREATE OR ALTER PROCEDURE cis.PlanPremium_Search
    @insCode nvarchar(4) = NULL, @payCode nvarchar(4) = NULL, @hltCode nvarchar(4) = NULL, @ageCode nvarchar(4) = NULL, @plnCode nvarchar(4) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT PRMCode, INSCode, PAYCode, HLTCode, AGECode, PLNCode, PRMAmount, SuppPRMAmt, PRMEffDate, PRMVersion, SuppPRMStatus
    FROM dbo.PRM
    WHERE (@insCode IS NULL OR INSCode = @insCode) AND (@payCode IS NULL OR PAYCode = @payCode) AND (@hltCode IS NULL OR HLTCode = @hltCode)
      AND (@ageCode IS NULL OR AGECode = @ageCode) AND (@plnCode IS NULL OR PLNCode = @plnCode)
    ORDER BY PRMCode;
END
GO

-- lstPMPRMList_Click: the age-band amounts of insured type + plan (every effective date; the last row of an age code wins, as in the desktop).
CREATE OR ALTER PROCEDURE cis.PlanPremium_Bands @insCode nvarchar(2), @plnCode nvarchar(4) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(AGECode) AS AgeCode, PRMAmount AS Amount, SuppPRMAmt AS SuppAmount FROM dbo.PRM WHERE INSCode = @insCode AND PLNCode = @plnCode ORDER BY PRMCode;
END
GO

-- SavePRM duplicate check: insured + health + plan + age + effective date (+ payor unless health type S).
CREATE OR ALTER PROCEDURE cis.PlanPremium_Exists
    @insCode nvarchar(2), @hltCode nvarchar(2), @plnCode nvarchar(4), @ageCode nvarchar(5), @effDate datetime, @payCode nvarchar(2) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.PRM WHERE INSCode = @insCode AND HLTCode = @hltCode AND PLNCode = @plnCode AND AGECode = @ageCode
                                  AND PRMEffDate = @effDate AND (@payCode IS NULL OR PAYCode = @payCode)) THEN 1 ELSE 0 END AS bit);
END
GO

-- SavePRM: one row (called seven times).
CREATE OR ALTER PROCEDURE cis.PlanPremium_Insert
    @insCode nvarchar(2), @payCode nvarchar(2), @hltCode nvarchar(2), @ageCode nvarchar(5), @plnCode nvarchar(4), @amount money = NULL,
    @suppAmount money = NULL, @suppStatus varchar(1), @effDate datetime, @version int = NULL, @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.PRM (INSCode, PAYCode, HLTCode, AGECode, PLNCode, PRMAmount, SuppPRMAmt, SuppPRMStatus, PRMEffDate, PRMVersion, PRMLastUpdateUser, PRMLastUpdateDate)
    VALUES (@insCode, @payCode, @hltCode, @ageCode, @plnCode, @amount, @suppAmount, @suppStatus, @effDate, @version, @user, @today);
END
GO

-- DeletePRM: refused while MBM (members) uses the premium's keys. CISDB has no MBM locally, so the check only runs where it exists.
CREATE OR ALTER PROCEDURE cis.PlanPremium_InUse
    @insCode nvarchar(2), @hltCode nvarchar(2), @plnCode nvarchar(4), @ageCode nvarchar(5), @payCode nvarchar(2) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    IF OBJECT_ID(N'dbo.MBM') IS NULL BEGIN SELECT CAST(0 AS bit); RETURN; END
    EXEC sp_executesql N'SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.MBM WHERE INSCode = @insCode AND HLTCode = @hltCode AND PLNCode = @plnCode
        AND AGECode = @ageCode AND (@payCode IS NULL OR PAYCode = @payCode)) THEN 1 ELSE 0 END AS bit)',
        N'@insCode nvarchar(2), @hltCode nvarchar(2), @plnCode nvarchar(4), @ageCode nvarchar(5), @payCode nvarchar(2)',
        @insCode = @insCode, @hltCode = @hltCode, @plnCode = @plnCode, @ageCode = @ageCode, @payCode = @payCode;
END
GO

CREATE OR ALTER PROCEDURE cis.PlanPremium_Delete @prmCode int AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.PRM WHERE PRMCode = @prmCode;
    SELECT @@ROWCOUNT;
END
GO
