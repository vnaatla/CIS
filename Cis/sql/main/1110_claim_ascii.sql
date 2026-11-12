-- Claim ASCII exports (frmClaimDetailsASCII, frmClaimASCIIToshiba, frmClaimASCIISinger, frmClaimASCIITIMS, frmClaimASCIICIMB, frmClaimDataOMNI). The Toshiba,
-- Singer, TIMS, CIMB and OMNI screens call the existing CISWorksheet procedures ClaimDetails_ASCII* (already parameterized); these are the remaining queries.
SET QUOTED_IDENTIFIER ON;
GO

-- frmClaimDetailsASCII ClaimASCII: View_ClaimDetails_ASCII for a bordx range and a clinic.
CREATE OR ALTER PROCEDURE cis.ClaimAscii_Bintulu @bordxFrom nvarchar(20), @bordxTo nvarchar(20), @clinic nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM CISWorksheet.dbo.View_ClaimDetails_ASCII WHERE BordxRefNo >= @bordxFrom AND BordxRefNo <= @bordxTo AND ClinicCode = @clinic;
END
GO

-- Form_Load clinic combos: SearchGrpCLN (Bintulu: the clinics of group company B003) or SearchCLN (panel clinics).
CREATE OR ALTER PROCEDURE cis.ClaimAscii_Clinics @bintulu bit AS
BEGIN
    SET NOCOUNT ON;
    IF @bintulu = 1
        SELECT RTRIM(g.GrpComClnCode) + ' - ' + RTRIM(c.CLNName) FROM dbo.GRPCOMClinic g JOIN cis.Hdb_CLN c ON g.GrpComClnCode = c.CLNCode WHERE g.GrpComCode = 'B003';
    ELSE
        SELECT RTRIM(CLNCode) + ' - ' + RTRIM(CLNName) FROM cis.Hdb_CLN WHERE CLNPanel = 'C' OR CLNPanel = 'Y';
END
GO

-- frmClaimASCIITIMS: the OP plan deductible of a plan and marital status.
CREATE OR ALTER PROCEDURE cis.ClaimAscii_Deductible @plan nvarchar(20), @marital nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 Amount FROM HISDB.dbo.OPPlanAmount WHERE PLNCode = @plan AND MaritalStatus = @marital;
END
GO

-- frmClaimDataOMNI CLNFile: the panel clinics not yet sent to OMNI, and marking one as sent.
CREATE OR ALTER PROCEDURE cis.ClaimAscii_OmniProviders AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM HISDB.dbo.CLN WHERE (CLNPanel = 'Y' OR CLNPanel = 'C') AND CLNOMNISent IS NULL;
END
GO

CREATE OR ALTER PROCEDURE cis.ClaimAscii_OmniSent @clinic nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE HISDB.dbo.CLN SET CLNOMNISent = 'Y' WHERE CLNCode = @clinic;
END
GO

-- frmClaimASCIIUtusan: ClaimDetails_ASCII_Utusan case @case (one view per claim type) with the bordx range as parameters.
CREATE OR ALTER PROCEDURE cis.ClaimAscii_Utusan @case int, @bordxFrom nvarchar(20), @bordxTo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @view sysname = CASE @case WHEN 1 THEN 'View_ClaimASCII_UtusanDelEO01A' WHEN 2 THEN 'View_ClaimASCII_UtusanNorDelS01B'
        WHEN 3 THEN 'View_ClaimASCII_UtusanCSecS01C' WHEN 4 THEN 'View_ClaimASCII_UtusanGPEO02A' WHEN 5 THEN 'View_ClaimASCII_UtusanGPDep02B'
        WHEN 6 THEN 'View_ClaimASCII_UtusanWadFam03A' WHEN 7 THEN 'View_ClaimASCII_UtusanWadFamNE03B' WHEN 8 THEN 'View_ClaimASCII_UtusanSPEO04A'
        WHEN 9 THEN 'View_ClaimASCII_UtusanSPDep04B' WHEN 10 THEN 'View_ClaimASCII_UtusanDental09' WHEN 11 THEN 'View_ClaimASCII_UtusanDepNE10A'
        WHEN 12 THEN 'View_ClaimASCII_UtusanPPEO11A' WHEN 13 THEN 'View_ClaimASCII_UtusanPPDep11B' END;
    IF @view IS NULL RETURN;
    DECLARE @sql nvarchar(400) = N'SELECT * FROM CISWorksheet.dbo.' + QUOTENAME(@view) + N' WHERE BordxRefNo >= @f AND BordxRefNo <= @t';
    EXEC sp_executesql @sql, N'@f nvarchar(20), @t nvarchar(20)', @f = @bordxFrom, @t = @bordxTo;
END
GO
