-- Maintenance > Benefit > Dental (frmMaintenanceBNFDental) tab Benefit Entry: DentalBNF grid over DentalBNFCategory. Same rules as the
-- Member tab (130_benefit_member.sql). Tooth Category / Tooth Reference / Benefit Category run on the code-table engine.

-- BNFcombolisting: INS, PLN codes, grid rows (DentalBNFCategory).
CREATE OR ALTER PROCEDURE cis.BenefitDental_EntrySetup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(INSCode) AS Code, ISNULL(INSDesc, N'') AS Name FROM dbo.INS ORDER BY INSCode;
    SELECT RTRIM(PLNCode) AS Code, RTRIM(PLNCode) AS Name FROM dbo.PLN ORDER BY PLNCode;
    SELECT RTRIM(DentalBNFCATCode) AS Code, ISNULL(DentalBNFCATDesc, N'') AS Name FROM dbo.DentalBNFCategory ORDER BY DentalBNFCATCode;
END
GO

-- searchtBNF → tBNFlisting.
CREATE OR ALTER PROCEDURE cis.BenefitDental_EntrySearch @plnCode nvarchar(10), @insCode nvarchar(2), @hltCode nvarchar(2), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(BNFCode) AS BnfCode, BNFCalculation AS Calculation, CAST(BNFNoOfDays AS nvarchar(10)) AS NoOfDays,
           CAST(BNFNoOfAnn AS nvarchar(10)) AS NoOfAnn, BNFRMPerVisit AS RmPerVisit, BNFPercentage AS Percentage
    FROM dbo.DentalBNF
    WHERE PLNCode = @plnCode AND INSCode = @insCode AND HLTCode = @hltCode AND BNFEffDate = @effDate
    ORDER BY BNFCode;
END
GO

-- SaveBNF duplicate check: plan + insured type + effective date.
CREATE OR ALTER PROCEDURE cis.BenefitDental_EntryExists @plnCode nvarchar(10), @insCode nvarchar(2), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.DentalBNF WHERE PLNCode = @plnCode AND INSCode = @insCode AND BNFEffDate = @effDate)
                     THEN 1 ELSE 0 END AS bit);
END
GO

-- SaveBNF: one DentalBNF row per qualifying grid row.
CREATE OR ALTER PROCEDURE cis.BenefitDental_EntryInsert
    @plnCode nvarchar(10), @insCode nvarchar(2), @hltCode nvarchar(2), @effDate datetime, @bnfCode nvarchar(10),
    @calculation nvarchar(10) = NULL, @noOfDays smallint = NULL, @noOfAnn smallint = NULL, @rmPerVisit nvarchar(15) = NULL,
    @percentage nvarchar(15) = NULL, @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.DentalBNF (BNFCode, HLTCode, INSCode, PLNCode, BNFCalculation, BNFNoOfDays, BNFNoOfAnn, BNFRMPerVisit, BNFPercentage,
                               BNFEffDate, BNFLastUpdateUser, BNFLastUpdateDate)
    VALUES (@bnfCode, @hltCode, @insCode, @plnCode, @calculation, @noOfDays, @noOfAnn, @rmPerVisit, @percentage, @effDate, @user, @today);
END
GO
