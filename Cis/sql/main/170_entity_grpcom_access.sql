-- Maintenance > Entity (frmMaintenanceEntity) tabs Group Company Clinic (CISDB GRPCOMClinic, VIEW_GrpCom_Clinic) and Clinic Access
-- (HISDB CLNAccessControl), plus frmSearchCLN (the clinic look-up of Clinic Access).

-- GRPCOMCLNComboListing (Check1): group companies (CISDB GRPCompany) and clinics (HISDB CLN).
CREATE OR ALTER PROCEDURE cis.EntityGrpCom_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoCode) AS Code, ISNULL(RTRIM(GRPCoName), N'') AS Name FROM dbo.GRPCompany ORDER BY GRPCoCode;
    SELECT RTRIM(CLNCode) AS Code, ISNULL(RTRIM(CLNName), N'') AS Name FROM cis.Hdb_CLN ORDER BY CLNCode;
END
GO

-- SearchGrpComCLN → GRPComCLNListing: = group company / = clinic; at most @max rows (the view has ~1.7 million) and the total.
CREATE OR ALTER PROCEDURE cis.EntityGrpCom_Search @grpCom nvarchar(10) = NULL, @cln nvarchar(10) = NULL, @max int = 1000 AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (@max) RTRIM(GrpComCode) AS GrpComCode, RTRIM(GRPCoName) AS GRPCoName, RTRIM(GrpComClnCode) AS GrpComClnCode, RTRIM(CLNName) AS CLNName,
           RTRIM(ISNULL(CLNAdd1, N'') + ISNULL(CLNAdd2, N'') + ISNULL(CLNAdd3, N'')) AS CLNAddress, RTRIM(CLNPostCode) AS CLNPostCode,
           RTRIM(CLNState) AS CLNState, RTRIM(CLNCity) AS CLNCity, RTRIM(CLNLocation) AS CLNLocation, RTRIM(CLN24Hrs) AS CLN24Hrs,
           RTRIM(CLNPhone) AS CLNPhone, RTRIM(CLNFax) AS CLNFax
    FROM dbo.VIEW_GrpCom_Clinic
    WHERE (@grpCom IS NULL OR GrpComCode = @grpCom) AND (@cln IS NULL OR GrpComClnCode = @cln)
    ORDER BY GrpComCode, GrpComClnCode;
    SELECT COUNT_BIG(*) FROM dbo.VIEW_GrpCom_Clinic WHERE (@grpCom IS NULL OR GrpComCode = @grpCom) AND (@cln IS NULL OR GrpComClnCode = @cln);
END
GO

-- SaveGrpComClinic duplicate check (Add and Edit): the pair already exists.
CREATE OR ALTER PROCEDURE cis.EntityGrpCom_Exists @grpCom nvarchar(10), @cln nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.GRPCOMClinic WHERE GrpComCode = @grpCom AND GrpComClnCode = @cln) THEN 1 ELSE 0 END AS bit);
END
GO

-- SaveGrpComClinic: Add inserts; Edit rewrites the selected pair (@oldGrpCom / @oldCln) with the new codes.
CREATE OR ALTER PROCEDURE cis.EntityGrpCom_Save
    @insert bit, @oldGrpCom nvarchar(10) = NULL, @oldCln nvarchar(10) = NULL, @grpCom nvarchar(10), @cln nvarchar(10), @user nvarchar(50), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF @insert = 1
        INSERT INTO dbo.GRPCOMClinic (GrpComCode, GrpComClnCode, GrpComClnLastUIpdateUser, GrpComClnLastUIpdateDate) VALUES (@grpCom, @cln, @user, @today);
    ELSE
        UPDATE dbo.GRPCOMClinic SET GrpComCode = @grpCom, GrpComClnCode = @cln, GrpComClnLastUIpdateUser = @user, GrpComClnLastUIpdateDate = @today
        WHERE GrpComCode = @oldGrpCom AND GrpComClnCode = @oldCln;
    SELECT @@ROWCOUNT;
END
GO

CREATE OR ALTER PROCEDURE cis.EntityGrpCom_Delete @grpCom nvarchar(10), @cln nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.GRPCOMClinic WHERE GrpComCode = @grpCom AND GrpComClnCode = @cln;
    SELECT @@ROWCOUNT;
END
GO

-- ClinicAccessListing: the latest 100 CLNAccessControl rows with the clinic name.
CREATE OR ALTER PROCEDURE cis.EntityAccess_List AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 100 RTRIM(a.ClinicCode) AS ClinicCode, c.CLNName, a.EffDateFrom, a.EffDateTo, a.CreatedBy, a.CreatedDate
    FROM cis.Hdb_CLNAccessControl a INNER JOIN cis.Hdb_CLN c ON c.CLNCode = a.ClinicCode
    ORDER BY a.CreatedDate DESC;
END
GO

-- SaveClnAccess: the clinic must exist.
CREATE OR ALTER PROCEDURE cis.EntityAccess_ClinicName @code nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CLNName FROM cis.Hdb_CLN WHERE CLNCode = @code;
END
GO

CREATE OR ALTER PROCEDURE cis.EntityAccess_Insert @code nvarchar(10), @from smalldatetime, @to smalldatetime, @user nvarchar(50), @now datetime AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO cis.Hdb_CLNAccessControl (ClinicCode, EffDateFrom, EffDateTo, CreatedBy, CreatedDate) VALUES (@code, @from, @to, @user, @now);
END
GO

-- frmSearchCLN.SearchClinic: CLNName LIKE %x%.
CREATE OR ALTER PROCEDURE cis.EntityAccess_FindClinics @name nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(CLNCode) AS Code, CLNName AS Name, COALESCE(CLNCity, N'') AS Location, COALESCE(CLNDocName, N'') AS DocName
    FROM cis.Hdb_CLN WHERE CLNName LIKE N'%' + @name + N'%' ORDER BY CLNCode;
END
GO
