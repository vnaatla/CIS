-- Maintenance > Benefit > Clinic (frmMaintenanceBenefitsCln) and its never-opened twin frmMaintenanceBenefitsClnNew: STDClinicBNFSpe.

-- ClinicBNFcombolisting: the grid rows (every BNFCategoryNew category).
CREATE OR ALTER PROCEDURE cis.BenefitClinic_Categories AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(BNFCATCode) AS Code, ISNULL(BNFCATDescription, N'') AS Name FROM dbo.BNFCategoryNew ORDER BY BNFCATCode;
END
GO

-- ClinicBNFlisting: Select * from STDClinicBNFSpe where CLNBNFType = @type And CLNBNFEffDate = @eff order by BNFCode
CREATE OR ALTER PROCEDURE cis.BenefitClinic_Search @cardType nvarchar(20), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(BNFCode) AS BnfCode, CLNBNFRMPerVisit AS RmPerVisit
    FROM dbo.STDClinicBNFSpe WHERE CLNBNFType = @cardType AND CLNBNFEffDate = @effDate ORDER BY BNFCode;
END
GO

-- SaveClinicBNF duplicate check: a card type + effective date can be saved once.
CREATE OR ALTER PROCEDURE cis.BenefitClinic_Exists @cardType nvarchar(20), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.STDClinicBNFSpe WHERE CLNBNFType = @cardType AND CLNBNFEffDate = @effDate) THEN 1 ELSE 0 END AS bit);
END
GO

-- SaveClinicBNF: one row per grid row (category); RM Payable Fee NULL when left empty.
CREATE OR ALTER PROCEDURE cis.BenefitClinic_Insert
    @cardType nvarchar(20), @effDate datetime, @bnfCode nvarchar(10), @rmPerVisit nvarchar(50) = NULL, @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.STDClinicBNFSpe (BNFCode, CLNBNFRMPerVisit, CLNBNFEffDate, CLNBNFType, CLNBNFLastUpdateUser, CLNBNFLastUpdateDate)
    VALUES (@bnfCode, @rmPerVisit, @effDate, @cardType, @user, @today);
END
GO
