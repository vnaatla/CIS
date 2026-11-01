-- Maintenance > Benefit > Member (frmMaintenanceBenefits) tabs Benefit Entry (BNFNew grid) and Benefit Listing (VIEW_CISBNF_PLN_SubCat).
-- The Category tab runs on the code-table engine (BNFCategoryNew).

-- BNFcombolisting (Benefit Entry): INS, PLN codes, grid rows. Health types are fixed in the form (S - Sihat, N - Non-Sihat).
CREATE OR ALTER PROCEDURE cis.BenefitMember_EntrySetup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(INSCode) AS Code, ISNULL(INSDesc, N'') AS Name FROM dbo.INS ORDER BY INSCode;
    SELECT RTRIM(PLNCode) AS Code, RTRIM(PLNCode) AS Name FROM dbo.PLN ORDER BY PLNCode;
    SELECT RTRIM(BNFCATCode) AS Code, ISNULL(BNFCATDescription, N'') AS Name FROM dbo.BNFCategoryNew ORDER BY BNFCATCode;
END
GO

-- BNFCombolisting2 (Benefit Listing): HLT, INS, PLN (CISDB), PAY and distinct PLN.GRPCompany of HISDB (medixmasterconn), categories.
CREATE OR ALTER PROCEDURE cis.BenefitMember_ListingSetup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(HLTCode) AS Code, ISNULL(HLTDesc, N'') AS Name FROM dbo.HLT ORDER BY HLTCode;
    SELECT RTRIM(INSCode) AS Code, ISNULL(INSDesc, N'') AS Name FROM dbo.INS ORDER BY INSCode;
    SELECT RTRIM(PLNCode) AS Code, RTRIM(PLNCode) AS Name FROM dbo.PLN ORDER BY PLNCode;
    SELECT RTRIM(PAYCode) AS Code, ISNULL(PAYCompanyName, N'') AS Name FROM cis.Hdb_PAY ORDER BY PAYCode;
    SELECT DISTINCT RTRIM(GRPCompany) AS Code, RTRIM(GRPCompany) AS Name FROM cis.Hdb_PLN WHERE LEN(GRPCompany) > 0 ORDER BY Code;
    SELECT RTRIM(BNFCATCode) AS Code, ISNULL(BNFCATDescription, N'') AS Name FROM dbo.BNFCategoryNew ORDER BY BNFCATCode;
END
GO

-- searchtBNF → tBNFlisting: the saved grid values of plan + insured type + health type + effective date.
CREATE OR ALTER PROCEDURE cis.BenefitMember_EntrySearch @plnCode nvarchar(10), @insCode nvarchar(2), @hltCode nvarchar(2), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(BNFCode) AS BnfCode, BNFCalculation AS Calculation, CAST(BNFNoOfDays AS nvarchar(10)) AS NoOfDays,
           CAST(BNFNoOfAnn AS nvarchar(10)) AS NoOfAnn,
           BNFRMPerVisit AS RmPerVisit, BNFPercentage AS Percentage
    FROM dbo.BNFNew
    WHERE PLNCode = @plnCode AND INSCode = @insCode AND HLTCode = @hltCode AND BNFEffDate = @effDate
    ORDER BY BNFCode;
END
GO

-- SaveBNF duplicate check: plan + insured type + effective date (the health type is not part of it).
CREATE OR ALTER PROCEDURE cis.BenefitMember_EntryExists @plnCode nvarchar(10), @insCode nvarchar(2), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.BNFNew WHERE PLNCode = @plnCode AND INSCode = @insCode AND BNFEffDate = @effDate)
                     THEN 1 ELSE 0 END AS bit);
END
GO

-- SaveBNF: one BNFNew row per grid row that qualifies (see BenefitMemberService.Qualifies).
CREATE OR ALTER PROCEDURE cis.BenefitMember_EntryInsert
    @plnCode nvarchar(10), @insCode nvarchar(2), @hltCode nvarchar(2), @effDate datetime, @bnfCode nvarchar(10),
    @calculation nvarchar(10) = NULL, @noOfDays smallint = NULL, @noOfAnn smallint = NULL, @rmPerVisit nvarchar(15) = NULL,
    @percentage nvarchar(15) = NULL, @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.BNFNew (BNFCode, HLTCode, INSCode, PLNCode, BNFCalculation, BNFNoOfDays, BNFNoOfAnn, BNFRMPerVisit, BNFPercentage,
                            BNFEffDate, BNFLastUpdateUser, BNFLastUpdateDate)
    VALUES (@bnfCode, @hltCode, @insCode, @plnCode, @calculation, @noOfDays, @noOfAnn, @rmPerVisit, @percentage, @effDate, @user, @today);
END
GO

-- searchBNF → BNFListing: every filter optional, = except GRPCompany (LIKE prefix%).
CREATE OR ALTER PROCEDURE cis.BenefitMember_Listing
    @bnfCode nvarchar(10) = NULL, @hltCode nvarchar(2) = NULL, @insCode nvarchar(2) = NULL, @plnCode nvarchar(10) = NULL,
    @effDate datetime = NULL, @payCode nvarchar(10) = NULL, @grpCompany nvarchar(50) = NULL, @calculation nvarchar(10) = NULL,
    @noOfDays nvarchar(10) = NULL, @rmPerVisit nvarchar(15) = NULL, @noOfAnn nvarchar(10) = NULL, @percentage nvarchar(15) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT BNFCode, BNFCATDescription, HLTCode, INSCode, PLNCode, BNFCalculation, BNFNoOfDays, BNFNoOfAnn, BNFRMPerVisit, BNFPercentage,
           BNFEffDate, PAYCode, GRPCompany
    FROM dbo.VIEW_CISBNF_PLN_SubCat
    WHERE (@bnfCode IS NULL OR BNFCode = @bnfCode) AND (@hltCode IS NULL OR HLTCode = @hltCode) AND (@insCode IS NULL OR INSCode = @insCode)
      AND (@plnCode IS NULL OR PLNCode = @plnCode) AND (@effDate IS NULL OR BNFEffDate = @effDate) AND (@payCode IS NULL OR PAYCode = @payCode)
      AND (@grpCompany IS NULL OR GRPCompany LIKE @grpCompany + N'%') AND (@calculation IS NULL OR BNFCalculation = @calculation)
      AND (@noOfDays IS NULL OR BNFNoOfDays = TRY_CAST(@noOfDays AS smallint)) AND (@rmPerVisit IS NULL OR BNFRMPerVisit = @rmPerVisit)
      AND (@noOfAnn IS NULL OR BNFNoOfAnn = TRY_CAST(@noOfAnn AS smallint)) AND (@percentage IS NULL OR BNFPercentage = @percentage);
END
GO
