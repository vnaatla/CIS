-- GENERATED FILE: built from sql/main/*.sql. Edit those files, then rebuild this one with
--     $env:CIS_WRITE_SQL='1'; dotnet test --filter Combined_script_is_up_to_date
-- How to run: sqlcmd -S <server> -E -d CISDB -i cis-main.sql   (or open it in SSMS on CISDB and press F5).
-- Check the database names at the top of the first section first.
-- Safe to run again: it only creates or replaces cis.* procedures and synonyms; tables and data are not changed.

-- ===== main/000_setup.sql =====
-- Setup for the CIS web stored procedures. Run on the CIS clinical database (the one that has dbo.CISCas and dbo.PAY).
-- Creates the cis schema and the synonyms the procedures use to reach the other legacy databases on the same server.
-- If a database has another name on your server, change the matching variable below before running.
SET NOEXEC OFF;
SET NOCOUNT ON;
DECLARE @HisDb sysname = N'HISDB';
DECLARE @MaintenanceDb sysname = N'HISMaintenance';
DECLARE @CisWorksheetDb sysname = N'CISWorksheet';
DECLARE @CisPaymentDb sysname = N'CISPayment';
DECLARE @HisWorksheetDb sysname = N'HISWorkSheet';
DECLARE @HisPaymentDb sysname = N'HISPayment';
DECLARE @ArchiveDb sysname = N'HISDMArchive';

DECLARE @major int = ISNULL(CAST(SERVERPROPERTY('ProductMajorVersion') AS int), 0);
DECLARE @build int = ISNULL(CAST(SERVERPROPERTY('ProductBuild') AS int), 0);
DECLARE @version nvarchar(128) = CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(128));
IF @major < 13 OR (@major = 13 AND @build < 4001)
BEGIN
    RAISERROR(N'STOPPED: this script needs SQL Server 2016 SP1 or later, but this server is version %s.', 16, 1, @version);
    SET NOEXEC ON;
END
ELSE IF OBJECT_ID(N'dbo.CISCas', N'U') IS NULL OR OBJECT_ID(N'dbo.PAY', N'U') IS NULL
BEGIN
    DECLARE @current sysname = DB_NAME();
    RAISERROR(N'STOPPED: database "%s" is not the CIS database (dbo.CISCas / dbo.PAY not found). Pick CISDB and run this script again.', 16, 1, @current);
    SET NOEXEC ON;
END
ELSE IF DB_ID(@MaintenanceDb) IS NULL OR OBJECT_ID(QUOTENAME(@MaintenanceDb) + N'.dbo.USR', N'U') IS NULL
BEGIN
    RAISERROR(N'STOPPED: maintenance database "%s" was not found (or it has no dbo.USR). Change @MaintenanceDb at the top of this script.', 16, 1, @MaintenanceDb);
    SET NOEXEC ON;
END
ELSE IF DB_ID(@HisDb) IS NULL OR OBJECT_ID(QUOTENAME(@HisDb) + N'.dbo.tblActionLog', N'U') IS NULL
BEGIN
    RAISERROR(N'STOPPED: HIS database "%s" was not found (or it has no dbo.tblActionLog). Change @HisDb at the top of this script.', 16, 1, @HisDb);
    SET NOEXEC ON;
END

IF SCHEMA_ID(N'cis') IS NULL EXEC (N'CREATE SCHEMA cis AUTHORIZATION dbo');

-- Synonym name prefix = the database it points to: Maint_ (HISMaintenance), Hdb_ (HISDB), Ws_ (CISWorksheet), Pay_ (CISPayment),
-- HWs_ (HISWorkSheet), HPay_ (HISPayment), Arch_ (HISDMArchive). Each phase adds the tables its screens use.
DECLARE @sql nvarchar(max) = N'';
SELECT @sql += N'IF OBJECT_ID(N''cis.' + s + N''', N''SN'') IS NOT NULL DROP SYNONYM cis.' + s + N';
CREATE SYNONYM cis.' + s + N' FOR ' + QUOTENAME(d) + N'.dbo.' + t + N';
'
FROM (VALUES
    (N'Maint_USR', @MaintenanceDb, N'USR'),
    (N'Maint_USRTrack', @MaintenanceDb, N'Tbl_USRTrack'),
    (N'Maint_ErrorLog', @MaintenanceDb, N'Tbl_ErrorLog'),
    (N'Hdb_ActionLog', @HisDb, N'tblActionLog'),
    (N'Hdb_CLN', @HisDb, N'CLN')) v (s, d, t)
WHERE DB_ID(d) IS NOT NULL;
EXEC (@sql);
GO

-- ===== main/010_auth.sql =====
-- Login / logout / change password (frmLogin, MDIMain.Logout_Click, mnuCloseConnection_Click) and the legacy audit tables.

-- frmLogin.cmdOK_Click: Select * from USR Where USRCode = '...' And USRStatus = 'A'
CREATE OR ALTER PROCEDURE cis.Login_GetActiveUser
    @userCode nvarchar(50)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT USRCode AS UsrCode, USRName AS UsrName, USRPassword AS UsrPassword, USRAccess AS UsrAccess
    FROM cis.Maint_USR
    WHERE USRCode = @userCode AND USRStatus = 'A';
END
GO

-- frmLogin.cmdOK_Click, fourth attempt after three wrong passwords: UPDATE USR SET USRStatus = 'S' WHERE (USRCode = '...')
CREATE OR ALTER PROCEDURE cis.Login_Suspend
    @userCode nvarchar(50)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Maint_USR SET USRStatus = 'S' WHERE USRCode = @userCode;
END
GO

-- frmLogin.cmdOK_Click, success: CIScnt + 1 and the Tbl_USRTrack row (text time and date, SYSName 'CIS').
-- The desktop also wrote the plain password to USR.actpass; that is deliberately not reproduced.
CREATE OR ALTER PROCEDURE cis.Login_Record
    @userCode nvarchar(50), @loginTime nvarchar(20), @loginDate nvarchar(20)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Maint_USR SET CIScnt = ISNULL(CIScnt, 0) + 1 WHERE USRCode = @userCode AND USRStatus = 'A';
    INSERT INTO cis.Maint_USRTrack (USRCode, SYSName, LoginTime, LoginDate) VALUES (LEFT(@userCode, 10), N'CIS', @loginTime, @loginDate);
END
GO

-- MDIMain.Logout_Click / mnuCloseConnection_Click: Update USR set CIScnt = CIScnt - 1 Where USRCode = '...' And USRStatus = 'A'
CREATE OR ALTER PROCEDURE cis.Login_Logout
    @userCode nvarchar(50)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Maint_USR SET CIScnt = CIScnt - 1 WHERE USRCode = @userCode AND USRStatus = 'A';
END
GO

-- frmLogin.cmdSave_Click. @password is already scrambled by the API (LegacyCrypt, upper case), so the user can log in with it.
CREATE OR ALTER PROCEDURE cis.Login_ChangePassword
    @userCode nvarchar(50), @password nvarchar(10)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Maint_USR SET USRPassword = @password WHERE USRCode = @userCode AND USRStatus = 'A';
END
GO

-- Medix.ErrorLog: Insert into Tbl_ErrorLog (SysName, FormName, MethodName, ErrDesc, ErrTime) values ('CIS', ..., GETDATE())
CREATE OR ALTER PROCEDURE cis.ErrorLog_Write
    @formName nvarchar(50), @methodName nvarchar(50), @description nvarchar(max)
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO cis.Maint_ErrorLog (SysName, FormName, MethodName, ErrDesc, ErrTime) VALUES (N'CIS', @formName, @methodName, @description, GETDATE());
END
GO

-- Medix.ActionLogUpdate: Insert into hisdb..tblactionlog values (user, actionId, '', '', '', getdate(), 'CIS', moduleName)
CREATE OR ALTER PROCEDURE cis.ActionLog_Insert
    @userId nvarchar(50), @actionId int, @notes nvarchar(max)
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO cis.Hdb_ActionLog (UserId, ActionId, MBMNumber, CoverID, CASNumber, DateTimeStamp, ProjType, Notes)
    VALUES (@userId, @actionId, N'', N'', N'', GETDATE(), N'CIS', @notes);
END
GO

-- ===== main/020_lookups.sql =====
-- Shared drop-down lists (the desktop's combo loaders, e.g. "select PAYCode, PAYCompanyName from PAY").

CREATE OR ALTER PROCEDURE cis.Lookup_Payors
    @healthType nvarchar(10) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PAYCode) AS Code, ISNULL(RTRIM(PAYCompanyName), N'') AS Name
    FROM dbo.PAY
    WHERE @healthType IS NULL OR HLTCode = @healthType
    ORDER BY PAYCode;
END
GO

CREATE OR ALTER PROCEDURE cis.Lookup_HealthTypes
AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(HLTCode) AS Code, ISNULL(RTRIM(HLTDesc), N'') AS Name FROM dbo.HLT ORDER BY HLTCode;
END
GO

CREATE OR ALTER PROCEDURE cis.Lookup_InsuredTypes
AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(INSCode) AS Code, ISNULL(RTRIM(INSDesc), N'') AS Name FROM dbo.INS ORDER BY INSCode;
END
GO

CREATE OR ALTER PROCEDURE cis.Lookup_GroupCompanies
    @payor nvarchar(10) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoCode) AS Code, ISNULL(RTRIM(GRPCoName), N'') AS Name
    FROM dbo.GRPCompany
    WHERE @payor IS NULL OR GRPCoPaycode = @payor
    ORDER BY GRPCoName;
END
GO

-- "Select CLNCode, CLNName from CLN WHERE CLNPanel = 'Y' OR CLNPanel = 'C'" (HISDB)
CREATE OR ALTER PROCEDURE cis.Lookup_Clinics
    @panelOnly bit = 0, @search nvarchar(200) = NULL, @max int = 500
AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (@max) RTRIM(CLNCode) AS Code, ISNULL(RTRIM(CLNName), N'') AS Name
    FROM cis.Hdb_CLN
    WHERE (@panelOnly = 0 OR CLNPanel IN ('Y', 'C'))
      AND (@search IS NULL OR CLNCode LIKE @search OR CLNName LIKE @search)
    ORDER BY CLNName;
END
GO

-- ===== main/100_list_reports.sql =====
-- Maintenance list print-outs: the 26 DataReports of the desktop (Rpt*.Dsr / rpt*.Dsr). Their data came from DataEnvironment1
-- (CISDB), which is not part of medixcis.vbp, so the desktop's Print buttons could not run; the commands below are the
-- DataEnvironment1 command texts, sorted as each report's DataReport_Initialize did. @code = the screen's search code (the desktop
-- filtered the recordset with the form's <Name>Criteria string); NULL = all rows.

CREATE OR ALTER PROCEDURE cis.ListReport_CheckItems @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT CatCode, CatName, Amount FROM dbo.CheckItems WHERE @code IS NULL OR CatCode = @code ORDER BY CatCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_CoPayment @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT COPCode, COPDesc, COPRemark FROM dbo.CoPayment WHERE @code IS NULL OR COPCode = @code ORDER BY COPCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_Fee @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT FeeCode, FeeDesc, FeeRemark FROM dbo.FEE WHERE @code IS NULL OR FeeCode = @code ORDER BY FeeCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_Reason @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT ReasonCode, ReasonDesc FROM dbo.TreatmentReason WHERE @code IS NULL OR ReasonCode = @code ORDER BY ReasonCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_HealthType @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT HLTCode, HLTDesc FROM dbo.HLT WHERE @code IS NULL OR HLTCode = @code ORDER BY HLTCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_InsuredType @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT INSCode, INSDesc FROM dbo.INS WHERE @code IS NULL OR INSCode = @code ORDER BY INSCode;
END
GO

-- RptMCOList had no sort and no filter.
CREATE OR ALTER PROCEDURE cis.ListReport_Mco @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT HLTCode, PLNCode, INSCode, MCOAmtAdult, MCOAmtChild, MCOEffDate FROM dbo.MCO WHERE @code IS NULL OR PLNCode = @code;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_Premium @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT PLNCode, PRMAmount, HLTCode, INSCode, PAYCode, SuppPRMAmt, AGECode, PRMEffDate, SuppPRMStatus, PRMVersion
    FROM dbo.PRM WHERE @code IS NULL OR PLNCode = @code ORDER BY PRMCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_VisitLimit @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT PLNCode, VisitLimit, HLTCode, INSCode, PAYCode, SuppVisitLimit, GRPCompany, VisitEffDate, SuppVisitLimitStatus, VisitLimitVersion
    FROM dbo.VIEW_VISIT_PAY_PLN WHERE @code IS NULL OR PLNCode = @code ORDER BY VisitIndex;
END
GO

-- RptAnnualList sorted on VisitIndex, which VIEW_ANN_PAY_PLN does not have (the print failed); its own index is AnnualIndex.
CREATE OR ALTER PROCEDURE cis.ListReport_AnnualLimit @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT PLNCode, AnnLimit, HLTCode, INSCode, PAYCode, SuppLimit, GRPCompany, AnnEffDate, SuppLimitStatus, AnnLimitVersion
    FROM dbo.VIEW_ANN_PAY_PLN WHERE @code IS NULL OR PLNCode = @code ORDER BY AnnualIndex;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_Procedures @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT ProcedureCode, ProcedureDesc FROM dbo.Procedures WHERE @code IS NULL OR ProcedureCode = @code ORDER BY ProcedureCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_Drug @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT DrugCode, DrugUnit, DrugValue, DrugDesc FROM dbo.Drug WHERE @code IS NULL OR DrugCode = @code ORDER BY DrugCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_Repudiation @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT RepCode, RepDesc FROM dbo.Repudiation WHERE @code IS NULL OR RepCode = @code ORDER BY RepCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_DiagnosisNotPayable @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT PLNCode, DiagCode FROM dbo.DiagNotPayable WHERE @code IS NULL OR PLNCode = @code ORDER BY PLNCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_GroupClinic @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT GRPClnCode, GRPClnName FROM dbo.GRPClinic WHERE @code IS NULL OR GRPClnCode = @code ORDER BY GRPClnCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_Doctor @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT DocCode, DocName FROM dbo.Doctor WHERE @code IS NULL OR DocCode = @code ORDER BY DocCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_GroupCompany @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT GRPCoCode, GRPCoName, GRPCoPaycode FROM dbo.GRPCompany WHERE @code IS NULL OR GRPCoCode = @code ORDER BY GRPCoCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_ToothCategory @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT ToothRefCatCode, ToothRefCatDesc FROM dbo.ToothRefCategory WHERE @code IS NULL OR ToothRefCatCode = @code ORDER BY ToothRefCatCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_ToothReference @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT ToothRefCode, ToothRefDesc, ToothRefCatDesc FROM dbo.VIEW_ToothRef_Category WHERE @code IS NULL OR ToothRefCode = @code ORDER BY ToothRefCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_DentalBenefitCategory @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT DentalBNFCATCode, DentalBNFCATDesc FROM dbo.DentalBNFCategory WHERE @code IS NULL OR DentalBNFCATCode = @code ORDER BY DentalBNFCATCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_BenefitCategory @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT BNFCATCode, BNFCATDescription FROM dbo.BNFCategoryNew WHERE @code IS NULL OR BNFCATCode = @code ORDER BY BNFCATCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_DentalBenefitEntry @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT BNFCode, DentalBNFCATDesc, BNFCalculation, BNFNoOfDays, BNFNoOfAnn, BNFRMPerVisit, BNFPercentage
    FROM dbo.VIEW_DentalBNF_Entry WHERE @code IS NULL OR BNFCode = @code ORDER BY BNFCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_CoPaymentMember @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT CoAmount, HLTType, PLNCode FROM dbo.CoPaymentMBM WHERE @code IS NULL OR PLNCode = @code ORDER BY HLTType;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_ClinicBenefit @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT BNFCode, BNFCATDescription, CLNBNFType, CLNBNFRMPerVisit, CLNBNFEffDate
    FROM dbo.VIEW_ClinicBNF_CardType WHERE @code IS NULL OR BNFCode = @code ORDER BY BNFCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_GroupCompanyClinic @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT GrpComCode, GRPCoName, GrpComClnCode, CLNName FROM dbo.VIEW_GrpCom_Clinic WHERE @code IS NULL OR GrpComCode = @code ORDER BY GrpComCode;
END
GO

CREATE OR ALTER PROCEDURE cis.ListReport_Plan @code nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT PLNCode, PLNDesc, HLTCode, PAYCode, GRPCompany, HISPAYCode FROM dbo.PLN WHERE @code IS NULL OR PLNCode = @code ORDER BY PLNCode;
END
GO

-- ===== main/110_lg_template_binding.sql =====
-- Maintenance > LG Maintenance > GL Template Binding (frmLGMaintenance tab 0, Tbl_OPGLMaster).

-- BindGroupCompany / BindGLType / BindGLTemplateName / BindCLNLGType
CREATE OR ALTER PROCEDURE cis.LgBinding_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT DISTINCT RTRIM(GRPCoCode) AS Code, ISNULL(RTRIM(GRPCoName), N'') AS Name FROM dbo.GRPCompany ORDER BY Code;
    SELECT CAST(GLID AS nvarchar(20)) AS Code, ISNULL(GLType, N'') AS Name FROM dbo.Tbl_GLType ORDER BY GLID;
    SELECT GLTemplateName AS Code, ISNULL(TmpRemarks, N'') AS Name FROM dbo.Tbl_GLTemplates ORDER BY GLTemplateID;
    SELECT RTRIM(CLNGLFormat) AS Code, RTRIM(CLNGLFormat) AS Name FROM cis.Hdb_CLN WHERE LEN(CLNGLFormat) > 0 GROUP BY CLNGLFormat ORDER BY Code;
END
GO

-- BindPlanCode / BindPayCode: PLN of the group company (first 4 characters of the combo = GRPCoCode).
CREATE OR ALTER PROCEDURE cis.LgBinding_PlansOfGroup @groupCode nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PLNCode) AS Code, ISNULL(RTRIM(PLNDesc), N'') AS Name FROM dbo.PLN WHERE GRPCompany = @groupCode ORDER BY PLNCode;
    SELECT DISTINCT RTRIM(PAYCode) AS Code, RTRIM(PAYCode) AS Name FROM dbo.PLN WHERE GRPCompany = @groupCode AND PAYCode IS NOT NULL ORDER BY Code;
END
GO

-- SearchGLTemplate: GRPCompnay like '%<group name>%' [and Paycode like @pay] [and PLNCode like @plan]
CREATE OR ALTER PROCEDURE cis.LgBinding_Search @groupName nvarchar(200), @payCode nvarchar(20) = NULL, @planCode nvarchar(10) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT ID, GRPCompnay AS GroupCompany, Paycode AS PayCode, PLNCode AS PlanCode, GLType, GLTemplate, RTRIM(LGFormat) AS LgFormat,
           RTRIM(VIPInd) AS VipInd, DepInd, DepTemplate, TmpRemarks
    FROM dbo.Tbl_OPGLMaster
    WHERE GRPCompnay LIKE @groupName AND (@payCode IS NULL OR Paycode LIKE @payCode) AND (@planCode IS NULL OR PLNCode LIKE @planCode)
    ORDER BY ID;
END
GO

-- SaveGLTemplate (Add): insert into Tbl_OPGLMaster.
CREATE OR ALTER PROCEDURE cis.LgBinding_Insert
    @groupName nvarchar(200), @planCode nvarchar(10), @glType nvarchar(50), @template nvarchar(200), @payCode nvarchar(20), @lgFormat nvarchar(5),
    @vip nchar(1), @depInd nchar(1), @depTemplate nvarchar(200), @remarks nvarchar(2000)
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.Tbl_OPGLMaster (GRPCompnay, PLNCode, GLType, GLTemplate, Paycode, LGFormat, VIPInd, DepInd, DepTemplate, TmpRemarks)
    VALUES (@groupName, @planCode, @glType, @template, @payCode, @lgFormat, @vip, @depInd, @depTemplate, @remarks);
END
GO

-- SaveGLTemplate (Edit): the desktop updates EVERY row of the plan code (… From Tbl_OPGLMaster where PLNCode = ...). Kept as is.
CREATE OR ALTER PROCEDURE cis.LgBinding_UpdatePlan
    @groupName nvarchar(200), @planCode nvarchar(10), @glType nvarchar(50), @template nvarchar(200), @payCode nvarchar(20), @lgFormat nvarchar(5),
    @vip nchar(1), @depInd nchar(1), @depTemplate nvarchar(200), @remarks nvarchar(2000)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.Tbl_OPGLMaster
    SET GRPCompnay = @groupName, GLType = @glType, GLTemplate = @template, Paycode = @payCode, LGFormat = @lgFormat, VIPInd = @vip,
        DepInd = @depInd, DepTemplate = @depTemplate, TmpRemarks = @remarks
    WHERE PLNCode = @planCode;
    SELECT @@ROWCOUNT;
END
GO

PRINT N'Done.';
GO
SET NOEXEC OFF;
GO
