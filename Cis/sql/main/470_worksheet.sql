-- Worksheet calculation (frmWorksheet): the values it looks up. The calculation itself is in the API (WorksheetEngine).
-- Result sets: 1 BNFCategoryNew rows, 2 the clinic (CLN), 3 STDClinicBNF, 4 STDClinicBNFSPE of the card, 5 PLN.BNFStatus, 6 BNFNew of the plan,
-- 7 CLNRedAmt (first row). Text amounts that are not numbers come back as NULL (the desktop's IsNumeric → 0).
CREATE OR ALTER PROCEDURE cis.Worksheet_Lookups
    @clinicCode nvarchar(10), @card nvarchar(50) = NULL, @planCode nvarchar(20) = NULL, @insCode nvarchar(20) = NULL, @hltCode nvarchar(20) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(BNFCATCode) AS Code, RTRIM(BNFCATDescription) AS Description FROM dbo.BNFCategoryNew ORDER BY BNFCATCode;
    SELECT TOP (1) CAST(CLNBNFType AS varchar(10)) AS BenefitType, RTRIM(CLNRedStatus) AS ReductionStatus,
           COALESCE(CLNMCOProcfeePerc, 0) AS McoProcFeePercent
    FROM cis.Hdb_CLN WHERE CLNCode = @clinicCode;
    SELECT RTRIM(BNFCode) AS Code, TRY_CONVERT(money, CLNBNFRMPerVisit) AS Rate FROM dbo.STDClinicBNF ORDER BY BNFCode;
    SELECT RTRIM(BNFCode) AS Code, TRY_CONVERT(money, CLNBNFRMPerVisit) AS Rate FROM dbo.STDClinicBNFSPE WHERE CLNBNFType = @card ORDER BY BNFCode;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.PLN WHERE PLNCode = @planCode AND BNFStatus = 'Y') THEN 1 ELSE 0 END AS bit) AS PlanBenefitStatus;
    SELECT RTRIM(BNFCode) AS Code, TRY_CONVERT(money, BNFPercentage) AS Percentage, TRY_CONVERT(money, BNFRMPerVisit) AS RmPerVisit,
           RTRIM(BNFCalculation) AS Calculation
    FROM dbo.BNFNew WHERE PLNCode = @planCode AND INSCode = @insCode AND HLTCode = @hltCode ORDER BY BNFCode;
    SELECT TOP (1) COALESCE(CLNAmout, 0) AS Amount, COALESCE(CLNPerc, 0) AS [Percent] FROM dbo.CLNRedAmt;
END
GO
