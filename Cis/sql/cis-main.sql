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
    (N'Hdb_CLN', @HisDb, N'CLN'),
    (N'Hdb_PAY', @HisDb, N'PAY'),
    (N'Hdb_PLN', @HisDb, N'PLN'),
    (N'Hdb_ClinicGrp', @HisDb, N'ClinicGrp'),
    (N'Hdb_STA', @HisDb, N'STA'),
    (N'Maint_CTY', @MaintenanceDb, N'CTY'),
    (N'Maint_CLNSeq', @MaintenanceDb, N'CLNSeq')) v (s, d, t)
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

-- ===== main/120_benefit_clinic.sql =====
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

-- ===== main/130_benefit_member.sql =====
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

-- ===== main/140_benefit_dental.sql =====
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

-- ===== main/150_entity_group_doctor.sql =====
-- Maintenance > Entity (frmMaintenanceEntity) tabs Clinic Group (GRPClinic) and Doctor (Doctor).

-- SearchClinicGrp → GrpClinicListing: GRPClnCode / GRPClnName LIKE %x% (upper-cased by the service).
CREATE OR ALTER PROCEDURE cis.EntityGroup_Search @code nvarchar(10) = NULL, @name nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT GRPClnCode, GRPClnName FROM dbo.GRPClinic
    WHERE (@code IS NULL OR GRPClnCode LIKE N'%' + @code + N'%') AND (@name IS NULL OR GRPClnName LIKE N'%' + @name + N'%')
    ORDER BY GRPClnCode;
END
GO

-- lstGrpClinicList_Click.
CREATE OR ALTER PROCEDURE cis.EntityGroup_Get @code nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT GRPClnCode AS Code, GRPClnName AS Name, GRPClnInitial AS Initial, GRPClnAdd1 AS Add1, GRPClnAdd2 AS Add2, GRPClnAdd3 AS Add3,
           GRPClnContact AS Contact, GRPClnAdmContact AS AdmContact, GRPClnCEO AS Ceo, GRPClnPhone AS Phone, GRPClnFax AS Fax,
           GRPClnEmail AS Email, GRPClnHomepage AS Homepage, CONVERT(nvarchar(10), GRPClnEffDate, 103) AS EffDate,
           RTRIM(GRPClnTax) AS Tax, RTRIM(GRPClnConversion) AS Conversion, RTRIM(GRPClnPrePost) AS PrePost
    FROM dbo.GRPClinic WHERE GRPClnCode = @code;
END
GO

-- GenerateGRPClnSeq (Medix.bas): GRPClnSeq row of the letter, +1, "0000"; letter + seq. Skips numbers already used in GRPClinic
-- (the desktop's insert then failed on PK_GRPClinic).
CREATE OR ALTER PROCEDURE cis.EntityGroup_NextCode @letter nchar(1) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seq int = (SELECT TRY_CAST(GRPClnSeq AS int) FROM dbo.GRPClnSeq WITH (UPDLOCK, HOLDLOCK) WHERE GRPClnCode = @letter);
    DECLARE @new bit = CASE WHEN @seq IS NULL THEN 1 ELSE 0 END;
    SET @seq = ISNULL(@seq, 0) + 1;
    WHILE EXISTS (SELECT 1 FROM dbo.GRPClinic WHERE GRPClnCode = @letter + RIGHT(N'0000' + CAST(@seq AS nvarchar(10)), 4)) SET @seq += 1;
    DECLARE @text nvarchar(5) = RIGHT(N'0000' + CAST(@seq AS nvarchar(10)), 4);
    IF @new = 1 INSERT INTO dbo.GRPClnSeq (GRPClnCode, GRPClnSeq) VALUES (@letter, @text);
    ELSE UPDATE dbo.GRPClnSeq SET GRPClnSeq = @text WHERE GRPClnCode = @letter;
    SELECT @letter + @text;
END
GO

-- SaveClinicGrp (Add and Edit write the same columns).
CREATE OR ALTER PROCEDURE cis.EntityGroup_Save
    @insert bit, @code nvarchar(5), @name nvarchar(50), @initial nvarchar(20), @add1 nvarchar(50) = NULL, @add2 nvarchar(50) = NULL,
    @add3 nvarchar(50) = NULL, @contact nvarchar(50) = NULL, @admContact nvarchar(50) = NULL, @ceo nvarchar(50) = NULL,
    @phone nvarchar(50) = NULL, @fax nvarchar(50) = NULL, @email nvarchar(50) = NULL, @homepage nvarchar(50), @effDate datetime = NULL,
    @tax nvarchar(5), @conversion nvarchar(5), @prePost nvarchar(5), @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF @insert = 1
        INSERT INTO dbo.GRPClinic (GRPClnCode, GRPClnName, GRPClnInitial, GRPClnAdd1, GRPClnAdd2, GRPClnAdd3, GRPClnContact, GRPClnAdmContact,
                                   GRPClnCEO, GRPClnPhone, GRPClnFax, GRPClnEmail, GRPClnHomepage, GRPClnEffDate, GRPClnTax, GRPClnConversion,
                                   GRPClnPrePost, GRPClnLastUpdateUser, GRPClnLastUpdateDate)
        VALUES (@code, @name, @initial, @add1, @add2, @add3, @contact, @admContact, @ceo, @phone, @fax, @email, @homepage, @effDate, @tax,
                @conversion, @prePost, @user, @today);
    ELSE
        UPDATE dbo.GRPClinic
        SET GRPClnName = @name, GRPClnInitial = @initial, GRPClnAdd1 = @add1, GRPClnAdd2 = @add2, GRPClnAdd3 = @add3, GRPClnContact = @contact,
            GRPClnAdmContact = @admContact, GRPClnCEO = @ceo, GRPClnPhone = @phone, GRPClnFax = @fax, GRPClnEmail = @email,
            GRPClnHomepage = @homepage, GRPClnEffDate = @effDate, GRPClnTax = @tax, GRPClnConversion = @conversion, GRPClnPrePost = @prePost,
            GRPClnLastUpdateUser = @user, GRPClnLastUpdateDate = @today
        WHERE GRPClnCode = @code;
    SELECT @@ROWCOUNT;
END
GO

-- DeleteClinicGrp: refused while a clinic (HISDB CLN.ClinicGrpCode) uses the group.
CREATE OR ALTER PROCEDURE cis.EntityGroup_InUse @code nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM cis.Hdb_CLN WHERE ClinicGrpCode = @code) THEN 1 ELSE 0 END AS bit);
END
GO

CREATE OR ALTER PROCEDURE cis.EntityGroup_Delete @code nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.GRPClinic WHERE GRPClnCode = @code;
    SELECT @@ROWCOUNT;
END
GO

-- SearchDoctor → DoctorListing.
CREATE OR ALTER PROCEDURE cis.EntityDoctor_Search @code nvarchar(10) = NULL, @name nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(DocCode) AS DocCode, DocName FROM dbo.Doctor
    WHERE (@code IS NULL OR DocCode LIKE N'%' + @code + N'%') AND (@name IS NULL OR DocName LIKE N'%' + @name + N'%')
    ORDER BY DocCode;
END
GO

-- SaveDoctor duplicate check on Add: the name is already used.
CREATE OR ALTER PROCEDURE cis.EntityDoctor_NameExists @name nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.Doctor WHERE DocName = @name) THEN 1 ELSE 0 END AS bit);
END
GO

-- GenerateDocSeq (Medix.bas): DOCSeq row of the letter, +1, "00000"; letter + seq; numbers already used in Doctor are skipped.
CREATE OR ALTER PROCEDURE cis.EntityDoctor_NextCode @letter nchar(1) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seq int = (SELECT TRY_CAST(DocSeq AS int) FROM dbo.DOCSeq WITH (UPDLOCK, HOLDLOCK) WHERE DocCode = @letter);
    DECLARE @new bit = CASE WHEN @seq IS NULL THEN 1 ELSE 0 END;
    SET @seq = ISNULL(@seq, 0) + 1;
    WHILE EXISTS (SELECT 1 FROM dbo.Doctor WHERE DocCode = @letter + RIGHT(N'00000' + CAST(@seq AS nvarchar(10)), 5)) SET @seq += 1;
    DECLARE @text nvarchar(5) = RIGHT(N'00000' + CAST(@seq AS nvarchar(10)), 5);
    IF @new = 1 INSERT INTO dbo.DOCSeq (DocCode, DocSeq) VALUES (@letter, @text);
    ELSE UPDATE dbo.DOCSeq SET DocSeq = @text WHERE DocCode = @letter;
    SELECT @letter + @text;
END
GO

-- SaveDoctor: Add inserts; Edit rewrites the name (the box is locked on Edit, so it is unchanged) and the audit columns.
CREATE OR ALTER PROCEDURE cis.EntityDoctor_Save @insert bit, @code nvarchar(10), @name nvarchar(50), @user nvarchar(10), @today datetime AS
BEGIN
    SET NOCOUNT ON;
    IF @insert = 1
        INSERT INTO dbo.Doctor (DocCode, DocName, DocLastUpdateUser, DocLastUpdateDate) VALUES (@code, @name, @user, @today);
    ELSE
        UPDATE dbo.Doctor SET DocName = @name, DocLastUpdateUser = @user, DocLastUpdateDate = @today WHERE DocCode = @code;
    SELECT @@ROWCOUNT;
END
GO

CREATE OR ALTER PROCEDURE cis.EntityDoctor_Delete @code nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.Doctor WHERE DocCode = @code;
    SELECT @@ROWCOUNT;
END
GO

-- ===== main/160_entity_clinic.sql =====
-- Maintenance > Entity (frmMaintenanceEntity) tab Clinic Info: HISDB CLN (medixmasterconn), ClinicGrp / STA (HISDB), CTY / CLNSeq
-- (HISMaintenance), cases in CISDB CISCAS.

-- CLNComboListing: clinic groups, cities, states.
CREATE OR ALTER PROCEDURE cis.EntityClinic_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(ClinicGrpCode) AS Code, ISNULL(ClinicGrpName, N'') AS Name FROM cis.Hdb_ClinicGrp ORDER BY ClinicGrpCode;
    SELECT DISTINCT RTRIM(CTYDescription) AS Code, RTRIM(CTYDescription) AS Name FROM cis.Maint_CTY WHERE LEN(CTYDescription) > 0 ORDER BY Code;
    SELECT DISTINCT RTRIM(STADescription) AS Code, RTRIM(STADescription) AS Name FROM cis.Hdb_STA WHERE LEN(STADescription) > 0 ORDER BY Code;
END
GO

-- SearchCLN → CLNListing (Select * from CLN where …). Values are already reduced by the service (upper-case, first letters, group prefix).
CREATE OR ALTER PROCEDURE cis.EntityClinic_Search
    @code nvarchar(5) = NULL, @initial nvarchar(50) = NULL, @name nvarchar(100) = NULL, @state nvarchar(30) = NULL, @city nvarchar(30) = NULL,
    @postCode nvarchar(10) = NULL, @panel nchar(1) = NULL, @groupPrefix nvarchar(10) = NULL, @type nchar(1) = NULL, @hrs24 nchar(1) = NULL,
    @ffs nchar(1) = NULL, @online nchar(1) = NULL, @glFormat nchar(1) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT CLNCode, CLNName, RTRIM(ClinicGrpCode) AS ClinicGrpCode, CLNInitial, CLNType, CLNBNFType, CLNAdd1, CLNAdd2, CLNAdd3, RTRIM(CLNPostCode) AS CLNPostCode,
           CLNCity, CLNState, CLNReceivedDate, CLNPanel, CLNXRay, CLNForRef, CLNNoofDoc, CLNPhone, CLNFax, CLNEmail, CLNHomepage, CLNBankAcc,
           CLNBankName, CLNBankBrn, CLNAccPerson, CLN24Hrs, CLNRemark, CLNFFSStatus, CLNOperatingHr, CLNPName, CLNPAdd1, CLNPAdd2, CLNPAdd3,
           RTRIM(CLNPPostCode) AS CLNPPostCode, CLNPCity, CLNPState, CLNOnlineStatus, CLNRegNo, CLNLocation, CLNMCOProcfeePerc, CLNGLFormat
    FROM cis.Hdb_CLN
    WHERE (@code IS NULL OR CLNCode = @code) AND (@initial IS NULL OR CLNInitial LIKE N'%' + @initial + N'%')
      AND (@name IS NULL OR CLNName LIKE N'%' + @name + N'%') AND (@state IS NULL OR CLNState LIKE N'%' + @state + N'%')
      AND (@city IS NULL OR CLNCity LIKE N'%' + @city + N'%') AND (@postCode IS NULL OR CLNPostCode LIKE N'%' + @postCode + N'%')
      AND (@panel IS NULL OR CLNPanel = @panel) AND (@groupPrefix IS NULL OR ClinicGrpCode LIKE @groupPrefix + N'%')
      AND (@type IS NULL OR CLNType = @type) AND (@hrs24 IS NULL OR CLN24Hrs = @hrs24) AND (@ffs IS NULL OR CLNFFSStatus = @ffs)
      AND (@online IS NULL OR (@online = N'Y' AND CLNOnlineStatus = 'Y') OR (@online <> N'Y' AND (CLNOnlineStatus = '' OR CLNOnlineStatus IS NULL OR CLNOnlineStatus = 'N')))
      AND (@glFormat IS NULL OR CLNGLFormat = CASE WHEN @glFormat = N'M' THEN 'M' ELSE 'H' END)
    ORDER BY CLNCode;
END
GO

-- SaveClinic duplicate check on Add: Address 1 already used (exact text).
CREATE OR ALTER PROCEDURE cis.EntityClinic_AddressExists @add1 nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM cis.Hdb_CLN WHERE CLNAdd1 = @add1) THEN 1 ELSE 0 END AS bit);
END
GO

-- GenerateCLNSeq (Medix.bas) on HISMaintenance CLNSeq: an existing letter → seq + 1 formatted "000", code = letter + seq; a new letter →
-- row "0001" and code = letter + letter + "001" (the desktop's doubling, kept). Codes already in CLN are skipped (PK_CLNNew).
CREATE OR ALTER PROCEDURE cis.EntityClinic_NextCode @letter nchar(1) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seq int = (SELECT TOP (1) TRY_CAST(CLNSeq AS int) FROM cis.Maint_CLNSeq WITH (UPDLOCK, HOLDLOCK) WHERE CLNCode = @letter);
    DECLARE @code nvarchar(10);
    IF @seq IS NULL
    BEGIN
        INSERT INTO cis.Maint_CLNSeq (CLNCode, CLNSeq) VALUES (@letter, N'0001');
        SET @code = @letter + @letter + N'001';
        SET @seq = 1;
        WHILE EXISTS (SELECT 1 FROM cis.Hdb_CLN WHERE CLNCode = @code)
        BEGIN
            SET @seq += 1;
            SET @code = @letter + RIGHT(N'000' + CAST(@seq AS nvarchar(10)), CASE WHEN @seq > 999 THEN LEN(CAST(@seq AS nvarchar(10))) ELSE 3 END);
        END
        IF @seq > 1 UPDATE cis.Maint_CLNSeq SET CLNSeq = RIGHT(N'000' + CAST(@seq AS nvarchar(10)), 3) WHERE CLNCode = @letter;
    END
    ELSE
    BEGIN
        SET @seq += 1;
        WHILE EXISTS (SELECT 1 FROM cis.Hdb_CLN WHERE CLNCode = @letter + CASE WHEN @seq > 999 THEN CAST(@seq AS nvarchar(10)) ELSE RIGHT(N'000' + CAST(@seq AS nvarchar(10)), 3) END)
            SET @seq += 1;
        DECLARE @text nvarchar(5) = CASE WHEN @seq > 999 THEN CAST(@seq AS nvarchar(10)) ELSE RIGHT(N'000' + CAST(@seq AS nvarchar(10)), 3) END;
        UPDATE cis.Maint_CLNSeq SET CLNSeq = @text WHERE CLNCode = @letter;
        SET @code = @letter + @text;
    END
    SELECT @code;
END
GO

-- SaveClinic: Add inserts, Edit updates the same columns (CLNPPostCode / CLNPCity / CLNPState are not written, as in the desktop).
CREATE OR ALTER PROCEDURE cis.EntityClinic_Save
    @insert bit, @code nvarchar(5), @name nvarchar(100), @groupCode nvarchar(6), @phone nvarchar(15), @fax nvarchar(15), @add1 nvarchar(200),
    @add2 nvarchar(200), @add3 nvarchar(200), @email nvarchar(1000), @homepage nvarchar(50), @representative nvarchar(200), @docName nvarchar(200),
    @bnfType int = NULL, @forRef nvarchar(50), @noOfDoc int = NULL, @type nchar(1), @xray nchar(1), @panel nchar(1), @hrs24 nchar(1), @ffs nchar(1),
    @ffsPerc numeric(18, 4) = NULL, @state nvarchar(30), @city nvarchar(30), @postCode nvarchar(5), @initial nvarchar(50), @receivedDate datetime,
    @bankAcc nvarchar(25), @bankName nvarchar(50), @bankBrn nvarchar(50), @accPerson nvarchar(50), @remark nvarchar(1500), @location nvarchar(100),
    @operatingHr nvarchar(100), @pName nvarchar(200), @pAdd1 nvarchar(50), @pAdd2 nvarchar(50), @pAdd3 nvarchar(50), @online nchar(1),
    @glFormat nchar(1), @regNo nvarchar(50), @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF @insert = 1
        INSERT INTO cis.Hdb_CLN (CLNCode, CLNName, ClinicGrpCode, CLNPhone, CLNFax, CLNAdd1, CLNAdd2, CLNAdd3, CLNEmail, CLNHomepage, CLNRepresentative,
                                 CLNDocName, CLNBNFType, CLNForRef, CLNNoofDoc, CLNType, CLNXRay, CLNPanel, CLN24Hrs, CLNFFSStatus, CLNMCOProcfeePerc,
                                 CLNState, CLNCity, CLNPostCode, CLNInitial, CLNReceivedDate, CLNBankAcc, CLNBankName, CLNBankBrn, CLNAccPerson,
                                 CLNRemark, CLNLocation, CLNLastUpdateUser, CLNLastUpdateDate, CLNOperatingHr, CLNPName, CLNPAdd1, CLNPAdd2, CLNPAdd3,
                                 CLNOnlineStatus, CLNGLFormat, CLNRegNo)
        VALUES (@code, @name, @groupCode, @phone, @fax, @add1, @add2, @add3, @email, @homepage, @representative, @docName, @bnfType, @forRef, @noOfDoc,
                @type, @xray, @panel, @hrs24, @ffs, @ffsPerc, @state, @city, @postCode, @initial, @receivedDate, @bankAcc, @bankName, @bankBrn,
                @accPerson, @remark, @location, @user, @today, @operatingHr, @pName, @pAdd1, @pAdd2, @pAdd3, @online, @glFormat, @regNo);
    ELSE
        UPDATE cis.Hdb_CLN
        SET CLNName = @name, ClinicGrpCode = @groupCode, CLNPhone = @phone, CLNFax = @fax, CLNAdd1 = @add1, CLNAdd2 = @add2, CLNAdd3 = @add3,
            CLNEmail = @email, CLNHomepage = @homepage, CLNRepresentative = @representative, CLNDocName = @docName, CLNBNFType = @bnfType,
            CLNForRef = @forRef, CLNNoofDoc = CASE WHEN @noOfDoc IS NULL THEN CLNNoofDoc ELSE @noOfDoc END, CLNType = @type, CLNXRay = @xray,
            CLNPanel = @panel, CLN24Hrs = @hrs24, CLNFFSStatus = @ffs, CLNMCOProcfeePerc = @ffsPerc, CLNState = @state, CLNCity = @city,
            CLNPostCode = @postCode, CLNInitial = @initial, CLNReceivedDate = @receivedDate, CLNBankAcc = @bankAcc, CLNBankName = @bankName,
            CLNBankBrn = @bankBrn, CLNAccPerson = @accPerson, CLNRemark = @remark, CLNLocation = @location, CLNLastUpdateUser = @user,
            CLNLastUpdateDate = @today, CLNOperatingHr = @operatingHr, CLNPName = @pName, CLNPAdd1 = @pAdd1, CLNPAdd2 = @pAdd2, CLNPAdd3 = @pAdd3,
            CLNOnlineStatus = @online, CLNGLFormat = @glFormat, CLNRegNo = @regNo
        WHERE CLNCode = @code;
    SELECT @@ROWCOUNT;
END
GO

-- DeleteClinic: refused while CISCAS has cases of the clinic.
CREATE OR ALTER PROCEDURE cis.EntityClinic_HasCases @code nvarchar(5) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.CISCAS WHERE ClinicCode = @code) THEN 1 ELSE 0 END AS bit);
END
GO

CREATE OR ALTER PROCEDURE cis.EntityClinic_Delete @code nvarchar(5) AS
BEGIN
    SET NOCOUNT ON;
    DELETE cis.Hdb_CLN WHERE CLNCode = @code;
    SELECT @@ROWCOUNT;
END
GO

PRINT N'Done.';
GO
SET NOEXEC OFF;
GO
