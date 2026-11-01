-- Maintenance > Plan (frmMaintenancePlan) tab MCO Fee: CISDB MCO (one row per insured type I / F / G / H).

-- MCOComboListing: health types and plan codes.
CREATE OR ALTER PROCEDURE cis.PlanMco_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(HLTCode) AS Code, ISNULL(HLTDesc, N'') AS Name FROM dbo.HLT ORDER BY HLTCode;
    SELECT RTRIM(PLNCode) AS Code, RTRIM(PLNCode) AS Name FROM dbo.PLN ORDER BY PLNCode;
END
GO

-- SearchMCO → MCOListing: health type (first character) =, plan =, effective date =.
CREATE OR ALTER PROCEDURE cis.PlanMco_Search @hltCode nchar(1) = NULL, @plnCode nvarchar(4) = NULL, @effDate datetime = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT MCOIndex, RTRIM(HLTCode) AS HLTCode, RTRIM(INSCode) AS INSCode, RTRIM(PLNCode) AS PLNCode, SuppMCOStatus, MCOAmtAdult, MCOAmtChild,
           MCOSuppAdultAmt, MCOSuppChildAmt, MCOEffDate
    FROM dbo.MCO
    WHERE (@hltCode IS NULL OR HLTCode = @hltCode) AND (@plnCode IS NULL OR PLNCode = @plnCode) AND (@effDate IS NULL OR MCOEffDate = @effDate)
    ORDER BY MCOIndex;
END
GO

-- SaveMCO1 duplicate check: plan + health + date + insured type.
CREATE OR ALTER PROCEDURE cis.PlanMco_Exists @plnCode nvarchar(4), @hltCode nvarchar(2), @effDate datetime, @insCode nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.MCO WHERE PLNCode = @plnCode AND HLTCode = @hltCode AND MCOEffDate = @effDate AND INSCode = @insCode)
                     THEN 1 ELSE 0 END AS bit);
END
GO

-- SaveMCO1: one insured type's fees.
CREATE OR ALTER PROCEDURE cis.PlanMco_Insert
    @plnCode nvarchar(4), @hltCode nvarchar(2), @insCode nvarchar(2), @suppStatus nvarchar(1), @effDate datetime, @adult money, @child money,
    @suppAdult money, @suppChild money, @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.MCO (PLNCode, HLTCode, INSCode, SuppMCOStatus, MCOEffDate, MCOAmtAdult, MCOAmtChild, MCOSuppAdultAmt, MCOSuppChildAmt,
                         MCOLastUpdateUser, MCOLastUpdateDate)
    VALUES (@plnCode, @hltCode, @insCode, @suppStatus, @effDate, @adult, @child, @suppAdult, @suppChild, @user, @today);
END
GO

-- DeleteMCO: the selected row (the desktop's DELETE named a column MCO does not have, MCOcode).
CREATE OR ALTER PROCEDURE cis.PlanMco_Delete @index int AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.MCO WHERE MCOIndex = @index;
    SELECT @@ROWCOUNT;
END
GO
