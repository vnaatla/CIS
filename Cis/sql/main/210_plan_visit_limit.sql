-- Maintenance > Plan (frmMaintenancePlan) tab Visit Limit: CISDB VisitLimit, VIEW_VISIT_PAY_PLN. Combos: cis.PlanAnnual_Setup (same lists).

-- SearchVisit → VisitListing.
CREATE OR ALTER PROCEDURE cis.PlanVisit_Search
    @plnCode nvarchar(4) = NULL, @hltCode nvarchar(4) = NULL, @insCode nvarchar(4) = NULL, @payCode nvarchar(4) = NULL, @grpCompany nvarchar(4) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT HLTCode, VisitIndex, PLNCode, INSCode, VisitLimit, SuppVisitLimit, VisitEffDate, PAYCode, GRPCompany, VisitLimitVersion, SuppVisitLimitStatus
    FROM dbo.VIEW_VISIT_PAY_PLN
    WHERE (@plnCode IS NULL OR PLNCode = @plnCode) AND (@hltCode IS NULL OR HLTCode = @hltCode) AND (@insCode IS NULL OR INSCode = @insCode)
      AND (@payCode IS NULL OR PAYCode = @payCode) AND (@grpCompany IS NULL OR GRPCompany LIKE @grpCompany + N'%')
    ORDER BY VisitIndex;
END
GO

-- SaveVisit duplicate check (Add only): insured + health + plan + version + effective date.
CREATE OR ALTER PROCEDURE cis.PlanVisit_Exists @insCode nvarchar(2), @hltCode nvarchar(2), @plnCode nvarchar(4), @version int, @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.VisitLimit WHERE INSCode = @insCode AND HLTCode = @hltCode AND PLNCode = @plnCode
                                  AND VisitLimitVersion = @version AND VisitEffDate = @effDate) THEN 1 ELSE 0 END AS bit);
END
GO

-- SaveVisit: Add inserts; Edit updates the selected VisitIndex. Status A leaves SuppVisitLimit as it is (NULL on Add).
CREATE OR ALTER PROCEDURE cis.PlanVisit_Save
    @index int = NULL, @plnCode nvarchar(4), @insCode nvarchar(2), @hltCode nvarchar(2), @visitLimit money, @status varchar(1),
    @suppLimit money = NULL, @effDate datetime, @version int, @user nvarchar(50), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF @index IS NULL
        INSERT INTO dbo.VisitLimit (PLNCode, INSCode, HLTCode, VisitLimit, SuppVisitLimitStatus, SuppVisitLimit, VisitEffDate, VisitLimitVersion,
                                    VisitLastUpdateUser, VisitLastUpdateDate)
        VALUES (@plnCode, @insCode, @hltCode, @visitLimit, @status, CASE WHEN @status = 'A' THEN NULL ELSE @suppLimit END, @effDate, @version, @user, @today);
    ELSE
        UPDATE dbo.VisitLimit
        SET PLNCode = @plnCode, INSCode = @insCode, HLTCode = @hltCode, VisitLimit = @visitLimit, SuppVisitLimitStatus = @status,
            SuppVisitLimit = CASE WHEN @status = 'A' THEN SuppVisitLimit ELSE @suppLimit END, VisitEffDate = @effDate, VisitLimitVersion = @version,
            VisitLastUpdateUser = @user, VisitLastUpdateDate = @today
        WHERE VisitIndex = @index;
    SELECT @@ROWCOUNT;
END
GO

CREATE OR ALTER PROCEDURE cis.PlanVisit_Delete @index int AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.VisitLimit WHERE VisitIndex = @index;
    SELECT @@ROWCOUNT;
END
GO
