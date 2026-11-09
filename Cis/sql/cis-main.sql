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
    (N'Maint_CLNSeq', @MaintenanceDb, N'CLNSeq'),
    (N'Hdb_CLNAccessControl', @HisDb, N'CLNAccessControl'),
    (N'Hdb_MBMCLNOthers', @HisDb, N'MBMCLNOthers'),
    (N'Maint_MBMCrossReference', @MaintenanceDb, N'MBMCrossReference'),
    (N'Hdb_MBM', @HisDb, N'MBM'),
    (N'Hdb_MBMOthers', @HisDb, N'MBMOthers'),
    (N'Hdb_MBMCoveredPersons', @HisDb, N'MBMCoveredPersons'),
    (N'Hdb_Search_CISMBM_05', @HisDb, N'Search_CISMBM_05'),
    (N'Hdb_Search_CISMBMSupp05', @HisDb, N'Search_CISMBMSupp05'),
    (N'Hdb_SearchCISMemberProc', @HisDb, N'SearchCISMemberProc'),
    (N'Hdb_AnnualLimitType', @HisDb, N'AnnualLimitType'),
    (N'Hdb_MBMCoveredPersonsCLN', @HisDb, N'MBMCoveredPersonsCLN'),
    (N'Hdb_SP_SearchCISMBM_IC', @HisDb, N'SP_SearchCISMBM_IC'),
    (N'Hdb_SearchRace', @HisDb, N'SearchRace'),
    (N'Maint_MBMSEQ', @MaintenanceDb, N'MBMSEQ'),
    (N'Hdb_MBMCheckDigits', @HisDb, N'MBMCheckDigits'),
    (N'Hdb_AGE', @HisDb, N'AGE'),
    (N'Hdb_MBMTwo', @HisDb, N'MBMTwo'),
    (N'Hdb_MBMOthersTwo', @HisDb, N'MBMOthersTwo'),
    (N'Hdb_REL', @HisDb, N'REL'),
    (N'Hdb_SearchCISMBMHistory', @HisDb, N'SearchCISMBMHistory'),
    (N'Hdb_SearchCISMemberProc_test', @HisDb, N'SearchCISMemberProc_test'),
    (N'Hdb_ASOMemberDetails', @HisDb, N'Tbl_ASOMemberDetails'),
    (N'Ws_Claim', @CisWorksheetDb, N'Claim'),
    (N'Ws_ClaimInvoiceActualNew', @CisWorksheetDb, N'ClaimInvoiceActualNew'),
    (N'Ws_VIEW_Invoice_Tracking_2016', @CisWorksheetDb, N'VIEW_Invoice_Tracking_2016'),
    (N'Ws_ClaimInvoice', @CisWorksheetDb, N'ClaimInvoice'),
    (N'Pay_Payment2CLN', @CisPaymentDb, N'Payment2CLN'),
    (N'Pay_Payment2MBM', @CisPaymentDb, N'Payment2MBM'),
    (N'Maint_VIEW_MMA_Schedule', @MaintenanceDb, N'VIEW_MMA_Schedule'),
    (N'Ws_VIEW_ClaimPass2Claim', @CisWorksheetDb, N'VIEW_ClaimPass2Claim'),
    (N'Ws_ClaimPass2ClaimUpdate', @CisWorksheetDb, N'ClaimPass2ClaimUpdate'),
    (N'Hdb_SearchCASToken', @HisDb, N'SearchCASToken'),
    (N'Ws_Bordx', @CisWorksheetDb, N'Bordx'),
    (N'Hdb_VIEW_MBM_Clinical', @HisDb, N'VIEW_MBM_Clinical'),
    (N'Ws_Save_Invoice', @CisWorksheetDb, N'Save_Invoice'),
    (N'Ws_Save_ClaimActual', @CisWorksheetDb, N'Save_ClaimActual'),
    (N'Ws_SaveActualOthers', @CisWorksheetDb, N'SaveActualOthers'),
    (N'Ws_SaveINVActual', @CisWorksheetDb, N'SaveINVActual'),
    (N'Ws_SaveINVCLNContract', @CisWorksheetDb, N'SaveINVCLNContract'),
    (N'Ws_SaveINVMBMApp', @CisWorksheetDb, N'SaveINVMBMApp'),
    (N'Ws_SaveINVMBMLmt', @CisWorksheetDb, N'SaveINVMBMLmt'),
    (N'Ws_SaveINVMBMExcess', @CisWorksheetDb, N'SaveINVMBMExcess'),
    (N'Ws_SaveINVCLNExcess', @CisWorksheetDb, N'SaveINVCLNExcess'),
    (N'Ws_SaveINVCLNApp', @CisWorksheetDb, N'SaveINVCLNApp'),
    (N'Ws_SaveINVClaim_2015', @CisWorksheetDb, N'SaveINVClaim_2015'),
    (N'Ws_Save_Charges', @CisWorksheetDb, N'Save_Charges'),
    (N'Hdb_sp_ASOClaimsUtilization', @HisDb, N'sp_ASOClaimsUtilization'),
    (N'Ws_ClaimInvoiceActualOthers', @CisWorksheetDb, N'ClaimInvoiceActualOthers'),
    (N'Ws_CLMAppNew', @CisWorksheetDb, N'CLMAppNew'),
    (N'Ws_SP_INV_CLAIM', @CisWorksheetDb, N'SP_INV_CLAIM'),
    (N'Ws_SP_Claim_Bordx', @CisWorksheetDb, N'SP_Claim_Bordx'),
    (N'Pay_ReceiptFrMBM', @CisPaymentDb, N'ReceiptFrMBM'),
    (N'Ws_CLMActualNew', @CisWorksheetDb, N'CLMActualNew'),
    (N'Ws_CLMClinicLmtNew', @CisWorksheetDb, N'CLMClinicLmtNew'),
    (N'Ws_CLMLimitNew', @CisWorksheetDb, N'CLMLimitNew'),
    (N'Ws_CLMMemberNew', @CisWorksheetDb, N'CLMMemberNew'),
    (N'Ws_CLMClinicNew', @CisWorksheetDb, N'CLMClinicNew'),
    (N'Ws_CLMClinicAppNew', @CisWorksheetDb, N'CLMClinicAppNew'),
    (N'Ws_SearchCASOP', @CisWorksheetDb, N'SearchCASOP'),
    (N'Ws_INVAdj_SaveInvoice', @CisWorksheetDb, N'INVAdj_SaveInvoice'),
    (N'Ws_INVAdj_saveActual', @CisWorksheetDb, N'INVAdj_saveActual'),
    (N'Ws_INVAdj_saveActualOthers', @CisWorksheetDb, N'INVAdj_saveActualOthers'),
    (N'Ws_INVAdj_SaveINVClaim_2015', @CisWorksheetDb, N'INVAdj_SaveINVClaim_2015'),
    (N'Ws_INVAdj_SaveINVActual', @CisWorksheetDb, N'INVAdj_SaveINVActual'),
    (N'Ws_INVAdj_SaveINVCLNApp', @CisWorksheetDb, N'INVAdj_SaveINVCLNApp'),
    (N'Ws_INVAdj_SaveINVCLNContract', @CisWorksheetDb, N'INVAdj_SaveINVCLNContract'),
    (N'Ws_INVAdj_SaveINVCLNExcess', @CisWorksheetDb, N'INVAdj_SaveINVCLNExcess'),
    (N'Ws_INVAdj_SaveINVMBMApp', @CisWorksheetDb, N'INVAdj_SaveINVMBMApp'),
    (N'Ws_INVAdj_SaveINVMBMExcess', @CisWorksheetDb, N'INVAdj_SaveINVMBMExcess'),
    (N'Ws_INVAdj_SaveINVMBMLmt', @CisWorksheetDb, N'INVAdj_SaveINVMBMLmt'),
    (N'Ws_CLMMedCharges', @CisWorksheetDb, N'CLMMedCharges'),
    (N'Ws_Claim_BordxSEQ', @CisWorksheetDb, N'Claim_BordxSEQ'),
    (N'Ws_SUNPosting', @CisWorksheetDb, N'SUNPosting'),
    (N'Ws_sp_CreateSubCasNumClaim', @CisWorksheetDb, N'sp_CreateSubCasNumClaim'),
    (N'Hdb_CLNLG1', @HisDb, N'CLNLG1'),
    (N'Hdb_SP_ValidateLG1', @HisDb, N'SP_ValidateLG1'),
    (N'Hdb_MessageComboList', @HisDb, N'MessageComboList'),
    (N'Ws_GetLGAnnLimitSpecLimit', @CisWorksheetDb, N'GetLGAnnLimitSpecLimit'),
    (N'Maint_CISEmailMaster', @MaintenanceDb, N'Tbl_CISEmailMaster'),
    (N'Maint_CLNMCOCRNoteSeq', @MaintenanceDb, N'CLNMCOCRNoteSeq'),
    (N'Hdb_Subsidary', @HisDb, N'Tbl_Subsidary'),
    (N'Hdb_CompanyGRP', @HisDb, N'CompanyGRP'),
    (N'Maint_Banks', @MaintenanceDb, N'Banks')) v (s, d, t)
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

-- ===== main/1000_sun_posting.sql =====
-- A/C Export: the posting batch numbers (Medix.bas GenSUNPostingBatchNo / GeneratePostingINVNo) and Invoice Posting > Clinic (frmPostingInvoiceCln):
-- CISWorksheet SUN_CIS_Posting 'InvCln' (VIEW_SUN_Posting_InvCln) with the pasted criteria replaced by parameters, and the SUNPosting row of a posted invoice.
SET QUOTED_IDENTIFIER ON;
GO

-- GenSUNPostingBatchNo: SunPostBatch holds the last batch no. ('1001' when empty); the row is locked until the caller's transaction ends.
CREATE OR ALTER PROCEDURE cis.SunPost_NextBatch @batchNo nvarchar(20) OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @last nvarchar(20) = (SELECT TOP 1 BatchNo FROM CISWorksheet.dbo.SunPostBatch WITH (UPDLOCK, HOLDLOCK) ORDER BY BatchIndex DESC);
    IF @last IS NULL
    BEGIN
        SET @batchNo = '1001';
        INSERT CISWorksheet.dbo.SunPostBatch (BatchNo) VALUES (@batchNo);
    END
    ELSE
    BEGIN
        SET @batchNo = CAST(CAST(@last AS int) + 1 AS nvarchar(20));
        UPDATE CISWorksheet.dbo.SunPostBatch SET BatchNo = @batchNo;
    END
END
GO

-- GeneratePostingINVNo: PostingINVSEQ holds a letter and 6 digits ("A000001" when empty); 999999 moves to the next letter.
CREATE OR ALTER PROCEDURE cis.SunPost_NextInvNo @seqNo nvarchar(20) OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @last nvarchar(20) = (SELECT TOP 1 SEQNumber FROM CISWorksheet.dbo.PostingINVSEQ WITH (UPDLOCK, HOLDLOCK));
    IF @last IS NULL
    BEGIN
        SET @seqNo = 'A000001';
        INSERT CISWorksheet.dbo.PostingINVSEQ (SEQNumber) VALUES (@seqNo);
        RETURN;
    END
    DECLARE @ch char(1) = LEFT(@last, 1), @n int = CAST(SUBSTRING(@last, 2, 6) AS int);
    IF @n = 999999 SELECT @ch = CHAR(ASCII(@ch) + 1), @n = 1 ELSE SET @n = @n + 1;
    SET @seqNo = @ch + RIGHT('000000' + CAST(@n AS varchar(6)), 6);
    UPDATE CISWorksheet.dbo.PostingINVSEQ SET SEQNumber = @seqNo;
END
GO

-- SearchRecord: @posted 0 / 1 as cboPostStatus. The bordx month and year are compared separately, as the desktop. @company is not applied: the view has no
-- group company column (the desktop's filter on MBMGroupCompany failed).
CREATE OR ALTER PROCEDURE cis.SunPost_InvCln_Search @batchNo nvarchar(20) = NULL, @clinic nvarchar(10) = NULL, @company nvarchar(200) = NULL,
    @invDateFrom datetime = NULL, @invDateTo datetime = NULL, @invFrom nvarchar(100) = NULL, @invTo nvarchar(100) = NULL, @posted char(1) = '0',
    @postDateFrom datetime = NULL, @postDateTo datetime = NULL, @batchFrom nvarchar(20) = NULL, @batchTo nvarchar(20) = NULL, @bordxMonthFrom nvarchar(2) = NULL,
    @bordxMonthTo nvarchar(2) = NULL, @bordxYearFrom nvarchar(4) = NULL, @bordxYearTo nvarchar(4) = NULL, @accMonth nvarchar(2) = NULL, @accYear nvarchar(4) = NULL,
    @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT ClnName, InvoiceDate, InvoiceNo, CLNAccCode, BordxMonth, BordxYear, PostStatus, PostType, AccMonth, AccYear, InvCLNStatus, PAYCode, INVPayTo, INVAmount,
        ClinicCode, BatchNo, PostDate, CLMPayableByMedix2C, CLMTotalActual, CLMCoPayColl, SUNInvoiceNo, CLMMCOProcFee, CLMFFSNoGST, CLMFFSGST, PVCLNDate, CLNInvoiceBatch
    FROM CISWorksheet.dbo.VIEW_SUN_Posting_InvCln
    WHERE (@batchNo IS NULL OR BatchNo = @batchNo)
      AND (@batchNo IS NOT NULL OR (
            (@clinic IS NULL OR ClinicCode = @clinic)
        AND (@invDateFrom IS NULL OR InvoiceDate >= @invDateFrom) AND (@invDateTo IS NULL OR InvoiceDate <= @invDateTo)
        AND (@invFrom IS NULL OR InvoiceNo >= @invFrom) AND (@invTo IS NULL OR InvoiceNo <= @invTo)
        AND (@posted <> '0' OR ((InvCLNStatus IS NULL OR InvCLNStatus = '0') AND (PostStatus IS NULL OR PostStatus = '0')))
        AND (@posted <> '1' OR (InvCLNStatus = '1' AND PostStatus = '1'))
        AND (@postDateFrom IS NULL OR PostDate >= @postDateFrom) AND (@postDateTo IS NULL OR PostDate <= @postDateTo)
        AND (@batchFrom IS NULL OR BatchNo >= @batchFrom) AND (@batchTo IS NULL OR BatchNo <= @batchTo)
        AND (@bordxMonthFrom IS NULL OR BordxMonth >= @bordxMonthFrom) AND (@bordxMonthTo IS NULL OR BordxMonth <= @bordxMonthTo)
        AND (@bordxYearFrom IS NULL OR BordxYear >= @bordxYearFrom) AND (@bordxYearTo IS NULL OR BordxYear <= @bordxYearTo)
        AND (@accMonth IS NULL OR AccMonth = @accMonth) AND (@accYear IS NULL OR AccYear = @accYear)
        AND (@pvDateFrom IS NULL OR PVCLNDate >= @pvDateFrom) AND (@pvDateTo IS NULL OR PVCLNDate <= @pvDateTo)))
    OPTION (RECOMPILE);
END
GO

-- CmdPost_Click: the SUNPosting row of one posted clinic invoice.
CREATE OR ALTER PROCEDURE cis.SunPost_InvCln_Insert @invoiceNo nvarchar(100), @clinic nvarchar(10), @accMonth nvarchar(2), @accYear nvarchar(4), @batchNo nvarchar(20),
    @invDate datetime = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISWorksheet.dbo.SUNPosting (InvoiceNo, TransRef, InvCLNStatus, PostStatus, PostType, PostDate, PostBy, AccMonth, AccYear, BatchNo, CLNInvDate)
    VALUES (@invoiceNo, @clinic, '1', '1', 'InvCln', CAST(GETDATE() AS date), @user, @accMonth, @accYear, @batchNo, @invDate);
END
GO

-- The accounting periods (HISMaintenance AccPeriod, "MM/YYYY").
CREATE OR ALTER PROCEDURE cis.SunPost_Periods AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(AccMonth) + '/' + RTRIM(AccYear) AS Period FROM HISMaintenance.dbo.AccPeriod ORDER BY AccYear DESC, AccMonth DESC;
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

-- ===== main/170_entity_grpcom_access.sql =====
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

-- SearchGrpComCLN → GRPComCLNListing: = group company / = clinic; one page (sorted by a listed column) and the total (the view has ~1.7 million rows).
CREATE OR ALTER PROCEDURE cis.EntityGrpCom_Search
    @grpCom nvarchar(10) = NULL, @cln nvarchar(10) = NULL, @page int = 0, @pageSize int = 50, @sortKey nvarchar(30) = NULL, @sortDesc bit = 0
AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GrpComCode) AS GrpComCode, RTRIM(GRPCoName) AS GRPCoName, RTRIM(GrpComClnCode) AS GrpComClnCode, RTRIM(CLNName) AS CLNName,
           RTRIM(ISNULL(CLNAdd1, N'') + ISNULL(CLNAdd2, N'') + ISNULL(CLNAdd3, N'')) AS CLNAddress, RTRIM(CLNPostCode) AS CLNPostCode,
           RTRIM(CLNState) AS CLNState, RTRIM(CLNCity) AS CLNCity, RTRIM(CLNLocation) AS CLNLocation, RTRIM(CLN24Hrs) AS CLN24Hrs,
           RTRIM(CLNPhone) AS CLNPhone, RTRIM(CLNFax) AS CLNFax
    FROM dbo.VIEW_GrpCom_Clinic
    WHERE (@grpCom IS NULL OR GrpComCode = @grpCom) AND (@cln IS NULL OR GrpComClnCode = @cln)
    ORDER BY
        CASE WHEN @sortDesc = 0 THEN CASE @sortKey WHEN N'grpCoName' THEN GRPCoName WHEN N'grpComClnCode' THEN GrpComClnCode WHEN N'clnName' THEN CLNName
            WHEN N'clnState' THEN CLNState WHEN N'clnCity' THEN CLNCity WHEN N'clnPostCode' THEN CLNPostCode ELSE GrpComCode END END,
        CASE WHEN @sortDesc = 1 THEN CASE @sortKey WHEN N'grpCoName' THEN GRPCoName WHEN N'grpComClnCode' THEN GrpComClnCode WHEN N'clnName' THEN CLNName
            WHEN N'clnState' THEN CLNState WHEN N'clnCity' THEN CLNCity WHEN N'clnPostCode' THEN CLNPostCode ELSE GrpComCode END END DESC,
        GrpComCode, GrpComClnCode
    OFFSET @page * @pageSize ROWS FETCH NEXT @pageSize ROWS ONLY;
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

-- ===== main/180_plan_code.sql =====
-- Maintenance > Plan (frmMaintenancePlan) tab Plan Code: CISDB PLN, VIEW_PLNListing, PLNSeq.

-- PLNComboListing / Check2_Click: health types, CIS and HIS payors, clinic contracts (CheckItems CSMD), group companies (not ZD01).
CREATE OR ALTER PROCEDURE cis.PlanCode_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(HLTCode) AS Code, ISNULL(HLTDesc, N'') AS Name FROM dbo.HLT ORDER BY HLTCode;
    SELECT RTRIM(PAYCode) AS Code, ISNULL(PAYCompanyName, N'') AS Name FROM dbo.PAY ORDER BY PAYCode;
    SELECT RTRIM(PAYCode) AS Code, ISNULL(PAYCompanyName, N'') AS Name FROM cis.Hdb_PAY ORDER BY PAYCode;
    SELECT RTRIM(CatName) AS Code, RTRIM(CatName) AS Name FROM dbo.CheckItems WHERE CatCode = 'CSMD' ORDER BY CatName;
    SELECT RTRIM(GRPCoCode) AS Code, RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoCode <> 'ZD01'
    GROUP BY GRPCoCode, GRPCoName ORDER BY GRPCoCode, GRPCoName;
END
GO

-- cboPLNHLTCode_Click: CIS payors of the health type.
CREATE OR ALTER PROCEDURE cis.PlanCode_Payors @hltCode nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    SELECT UPPER(RTRIM(PAYCode)) AS Code, UPPER(ISNULL(PAYCompanyName, N'')) AS Name FROM dbo.PAY WHERE HLTCode = @hltCode ORDER BY PAYCode;
END
GO

-- SearchPLN → PLNListing (VIEW_PLNListing); the indicator filters the view lacks are read from PLN.
CREATE OR ALTER PROCEDURE cis.PlanCode_Search
    @plnCode nvarchar(10) = NULL, @payCode nvarchar(2) = NULL, @hisPayCode nvarchar(2) = NULL, @hltCode nchar(1) = NULL, @grpCode nvarchar(4) = NULL,
    @clnContract nvarchar(50) = NULL, @annLmt nchar(1) = NULL, @visLmt nchar(1) = NULL, @prm nchar(1) = NULL, @mco nchar(1) = NULL,
    @pDentalStatus nchar(1) = NULL, @pDentalLimit nchar(1) = NULL, @pSpecStatus nchar(1) = NULL, @pSpecLimit nchar(1) = NULL,
    @pMepStatus nchar(1) = NULL, @sDentalStatus nchar(1) = NULL, @sDentalLimit nchar(1) = NULL, @sSpecStatus nchar(1) = NULL,
    @sSpecLimit nchar(1) = NULL, @sMepStatus nchar(1) = NULL, @sAnnLmt nchar(1) = NULL, @sVisLmt nchar(1) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT v.HLTCode, v.PLNCode, v.PLNDesc, RTRIM(v.PAYCode) AS PAYCode, RTRIM(v.HISPAYCode) AS HISPAYCode, v.CoGRPCode, v.CoGRPName, v.ClnContract,
           v.Insured, v.HRApproval, v.PrinSPLimit, v.PrinGPLimit, v.DepSPLimit, v.DepGPLimit, v.OPFamilyStatus, v.PLNSpecialRemark, v.isopdependent
    FROM dbo.VIEW_PLNListing v
    INNER JOIN dbo.PLN p ON p.PLNCode = v.PLNCode
    WHERE (@plnCode IS NULL OR v.PLNCode LIKE N'%' + @plnCode + N'%') AND (@payCode IS NULL OR v.PAYCode = @payCode)
      AND (@hisPayCode IS NULL OR v.HISPAYCode = @hisPayCode) AND (@hltCode IS NULL OR v.HLTCode = @hltCode)
      AND (@grpCode IS NULL OR v.CoGRPCode = @grpCode) AND (@clnContract IS NULL OR v.ClnContract = @clnContract)
      AND (@annLmt IS NULL OR p.ANNLmtInd = @annLmt) AND (@visLmt IS NULL OR p.VISLmtInd = @visLmt) AND (@prm IS NULL OR p.PRMInd = @prm)
      AND (@mco IS NULL OR p.MCOInd = @mco) AND (@pDentalStatus IS NULL OR p.PDentalStatus = @pDentalStatus)
      AND (@pDentalLimit IS NULL OR p.PDentalLimit = @pDentalLimit) AND (@pSpecStatus IS NULL OR p.PSpecStatus = @pSpecStatus)
      AND (@pSpecLimit IS NULL OR p.PSpecLimit = @pSpecLimit) AND (@pMepStatus IS NULL OR p.PMEPStatus = @pMepStatus)
      AND (@sDentalStatus IS NULL OR p.SDentalStatus = @sDentalStatus) AND (@sDentalLimit IS NULL OR p.SDentalLimit = @sDentalLimit)
      AND (@sSpecStatus IS NULL OR p.SSpecStatus = @sSpecStatus) AND (@sSpecLimit IS NULL OR p.SSpecLimit = @sSpecLimit)
      AND (@sMepStatus IS NULL OR p.SMEPStatus = @sMepStatus) AND (@sAnnLmt IS NULL OR p.ANNLmtIndSupp = @sAnnLmt)
      AND (@sVisLmt IS NULL OR p.VISLmtIndSupp = @sVisLmt)
    ORDER BY v.PLNCode;
END
GO

-- lstPlanList_Click: the plan's columns.
CREATE OR ALTER PROCEDURE cis.PlanCode_Get @code nvarchar(4) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PLNCode) AS PlnCode, PLNDesc AS PlnDesc, RTRIM(HLTCode) AS HltCode, RTRIM(PAYCode) AS PayCode, RTRIM(HISPAYCode) AS HisPayCode,
           RTRIM(GRPCompany) AS GrpCompany, ClnContract, RTRIM(ANNLmtInd) AS AnnLmt, RTRIM(VISLmtInd) AS VisLmt, RTRIM(PRMInd) AS Prm,
           RTRIM(MCOInd) AS Mco, RTRIM(DCStatus) AS DcStatus, RTRIM(PDentalStatus) AS PDentalStatus, RTRIM(PDentalLimit) AS PDentalLimit,
           RTRIM(PSpecStatus) AS PSpecStatus, RTRIM(PSpecLimit) AS PSpecLimit, RTRIM(PMEPStatus) AS PMepStatus, RTRIM(SDentalStatus) AS SDentalStatus,
           RTRIM(SDentalLimit) AS SDentalLimit, RTRIM(SSpecStatus) AS SSpecStatus, RTRIM(SSpecLimit) AS SSpecLimit, RTRIM(SMEPStatus) AS SMepStatus,
           RTRIM(ANNLmtIndSupp) AS SAnnLmt, RTRIM(VISLmtIndSupp) AS SVisLmt, RTRIM(Insured) AS Insured, RTRIM(HRApproval) AS HrApproval,
           OPFamilyStatus AS OpFamilyStatus, CAST(PrinSPLimit AS nvarchar(30)) AS PrinSpLimit, CAST(PrinGPLimit AS nvarchar(30)) AS PrinGpLimit,
           CAST(DepSPLimit AS nvarchar(30)) AS DepSpLimit, CAST(DepGPLimit AS nvarchar(30)) AS DepGpLimit, PLNSpecialRemark AS SpecialRemark,
           RTRIM(isOPDependent) AS OpDependent
    FROM dbo.PLN WHERE PLNCode = @code;
END
GO

-- SavePLN (N): "Record Duplicated!" when the Grp Comp. text is a group company name.
CREATE OR ALTER PROCEDURE cis.PlanCode_GroupNameExists @name nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.GRPCompany WHERE GRPCoName = @name) THEN 1 ELSE 0 END AS bit);
END
GO

-- GeneratePLNSeq (Medix.bas): PLNSeq row of the letter, +1 "000" (a new letter starts at 001); codes already in PLN are skipped.
CREATE OR ALTER PROCEDURE cis.PlanCode_NextCode @letter nchar(1) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seq int = (SELECT TOP (1) TRY_CAST(PLNNoSeq AS int) FROM dbo.PLNSeq WITH (UPDLOCK, HOLDLOCK) WHERE PLNSeqCode = @letter);
    DECLARE @new bit = CASE WHEN @seq IS NULL THEN 1 ELSE 0 END;
    SET @seq = ISNULL(@seq, 0) + 1;
    WHILE EXISTS (SELECT 1 FROM dbo.PLN WHERE PLNCode = @letter + RIGHT(N'000' + CAST(@seq AS nvarchar(10)), 3)) SET @seq += 1;
    IF @seq > 999 THROW 51000, N'No plan codes left for this letter.', 1;
    DECLARE @text nvarchar(3) = RIGHT(N'000' + CAST(@seq AS nvarchar(10)), 3);
    IF @new = 1 INSERT INTO dbo.PLNSeq (PLNSeqCode, PLNNoSeq) VALUES (@letter, @text);
    ELSE UPDATE dbo.PLNSeq SET PLNNoSeq = @text WHERE PLNSeqCode = @letter;
    SELECT @letter + @text;
END
GO

-- SavePLN (N): one PLN row per plan; dental / spec / MEP indicators always N, as the desktop wrote them.
CREATE OR ALTER PROCEDURE cis.PlanCode_Insert
    @code nvarchar(4), @desc nvarchar(200), @hltCode nvarchar(2), @payCode nvarchar(2), @hisPayCode nvarchar(2), @grpCompany nvarchar(200),
    @annLmt nchar(1), @visLmt nchar(1), @sAnnLmt nchar(1), @sVisLmt nchar(1), @prm nchar(1), @mco nchar(1), @clnContract nvarchar(50),
    @dcStatus nchar(1), @insured nchar(1), @hrApproval nchar(1), @opFamilyStatus nvarchar(50), @prinSpLimit float, @prinGpLimit float,
    @depSpLimit float, @depGpLimit float, @specialRemark varchar(1500), @opDependent nchar(1), @user nvarchar(100), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.PLN (PLNCode, PLNDesc, HLTCode, PAYCode, HISPAYCode, GRPCompany, ANNLmtInd, VISLmtInd, ANNLmtIndSupp, VISLmtIndSupp, PRMInd, MCOInd,
                         ClnContract, DCStatus, PDentalStatus, PSpecStatus, PMEPStatus, PDentalLimit, PSpecLimit, SDentalStatus, SSpecStatus, SMEPStatus,
                         SDentalLimit, SSpecLimit, Insured, HRApproval, OPFamilyStatus, PrinSPLimit, PrinGPLimit, DepSPLimit, DepGPLimit, PLNSpecialRemark,
                         isOPDependent, PLNLastUpdateUser, PLNLastUpdateDate)
    VALUES (@code, @desc, @hltCode, @payCode, @hisPayCode, @grpCompany, @annLmt, @visLmt, @sAnnLmt, @sVisLmt, @prm, @mco, @clnContract, @dcStatus,
            'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', @insured, @hrApproval, @opFamilyStatus, @prinSpLimit, @prinGpLimit, @depSpLimit,
            @depGpLimit, @specialRemark, @opDependent, @user, @today);
END
GO

-- DeletePLN: refused while membership (HISDB MBMCLNOthers.PLNType) uses the plan.
CREATE OR ALTER PROCEDURE cis.PlanCode_InUse @code nvarchar(4) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM cis.Hdb_MBMCLNOthers WHERE PLNType = @code) THEN 1 ELSE 0 END AS bit);
END
GO

CREATE OR ALTER PROCEDURE cis.PlanCode_Delete @code nvarchar(4) AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.PLN WHERE PLNCode = @code;
    SELECT @@ROWCOUNT;
END
GO

-- ===== main/190_plan_premium.sql =====
-- Maintenance > Plan (frmMaintenancePlan) tab Premium: CISDB PRM (one row per age band 01–07).

-- PRMComboListing: insured types, health types, payors, age codes, plans.
CREATE OR ALTER PROCEDURE cis.PlanPremium_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(INSCode) AS Code, ISNULL(INSDesc, N'') AS Name FROM dbo.INS ORDER BY INSCode;
    SELECT RTRIM(HLTCode) AS Code, ISNULL(HLTDesc, N'') AS Name FROM dbo.HLT ORDER BY HLTCode;
    SELECT RTRIM(PAYCode) AS Code, ISNULL(PAYCompanyName, N'') AS Name FROM dbo.PAY ORDER BY PAYCode;
    SELECT RTRIM(AgeCode) AS Code, ISNULL(AGEDesc, N'') AS Name FROM dbo.AGE ORDER BY AgeCode;
    SELECT RTRIM(PLNCode) AS Code, ISNULL(PLNDesc, N'') AS Name FROM dbo.PLN ORDER BY PLNCode;
END
GO

-- SearchPRM → PRMListing.
CREATE OR ALTER PROCEDURE cis.PlanPremium_Search
    @insCode nvarchar(4) = NULL, @payCode nvarchar(4) = NULL, @hltCode nvarchar(4) = NULL, @ageCode nvarchar(4) = NULL, @plnCode nvarchar(4) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT PRMCode, INSCode, PAYCode, HLTCode, AGECode, PLNCode, PRMAmount, SuppPRMAmt, PRMEffDate, PRMVersion, SuppPRMStatus
    FROM dbo.PRM
    WHERE (@insCode IS NULL OR INSCode = @insCode) AND (@payCode IS NULL OR PAYCode = @payCode) AND (@hltCode IS NULL OR HLTCode = @hltCode)
      AND (@ageCode IS NULL OR AGECode = @ageCode) AND (@plnCode IS NULL OR PLNCode = @plnCode)
    ORDER BY PRMCode;
END
GO

-- lstPMPRMList_Click: the age-band amounts of insured type + plan (every effective date; the last row of an age code wins, as in the desktop).
CREATE OR ALTER PROCEDURE cis.PlanPremium_Bands @insCode nvarchar(2), @plnCode nvarchar(4) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(AGECode) AS AgeCode, PRMAmount AS Amount, SuppPRMAmt AS SuppAmount FROM dbo.PRM WHERE INSCode = @insCode AND PLNCode = @plnCode ORDER BY PRMCode;
END
GO

-- SavePRM duplicate check: insured + health + plan + age + effective date (+ payor unless health type S).
CREATE OR ALTER PROCEDURE cis.PlanPremium_Exists
    @insCode nvarchar(2), @hltCode nvarchar(2), @plnCode nvarchar(4), @ageCode nvarchar(5), @effDate datetime, @payCode nvarchar(2) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.PRM WHERE INSCode = @insCode AND HLTCode = @hltCode AND PLNCode = @plnCode AND AGECode = @ageCode
                                  AND PRMEffDate = @effDate AND (@payCode IS NULL OR PAYCode = @payCode)) THEN 1 ELSE 0 END AS bit);
END
GO

-- SavePRM: one row (called seven times).
CREATE OR ALTER PROCEDURE cis.PlanPremium_Insert
    @insCode nvarchar(2), @payCode nvarchar(2), @hltCode nvarchar(2), @ageCode nvarchar(5), @plnCode nvarchar(4), @amount money = NULL,
    @suppAmount money = NULL, @suppStatus varchar(1), @effDate datetime, @version int = NULL, @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.PRM (INSCode, PAYCode, HLTCode, AGECode, PLNCode, PRMAmount, SuppPRMAmt, SuppPRMStatus, PRMEffDate, PRMVersion, PRMLastUpdateUser, PRMLastUpdateDate)
    VALUES (@insCode, @payCode, @hltCode, @ageCode, @plnCode, @amount, @suppAmount, @suppStatus, @effDate, @version, @user, @today);
END
GO

-- DeletePRM: refused while MBM (members) uses the premium's keys. CISDB has no MBM locally, so the check only runs where it exists.
CREATE OR ALTER PROCEDURE cis.PlanPremium_InUse
    @insCode nvarchar(2), @hltCode nvarchar(2), @plnCode nvarchar(4), @ageCode nvarchar(5), @payCode nvarchar(2) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    IF OBJECT_ID(N'dbo.MBM') IS NULL BEGIN SELECT CAST(0 AS bit); RETURN; END
    EXEC sp_executesql N'SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.MBM WHERE INSCode = @insCode AND HLTCode = @hltCode AND PLNCode = @plnCode
        AND AGECode = @ageCode AND (@payCode IS NULL OR PAYCode = @payCode)) THEN 1 ELSE 0 END AS bit)',
        N'@insCode nvarchar(2), @hltCode nvarchar(2), @plnCode nvarchar(4), @ageCode nvarchar(5), @payCode nvarchar(2)',
        @insCode = @insCode, @hltCode = @hltCode, @plnCode = @plnCode, @ageCode = @ageCode, @payCode = @payCode;
END
GO

CREATE OR ALTER PROCEDURE cis.PlanPremium_Delete @prmCode int AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.PRM WHERE PRMCode = @prmCode;
    SELECT @@ROWCOUNT;
END
GO

-- ===== main/200_plan_annual_limit.sql =====
-- Maintenance > Plan (frmMaintenancePlan) tab Annual Limit: CISDB AnnualLimit, VIEW_ANN_PAY_PLN.

-- AnnualComboListing: plans, insured types, health types, payors, distinct PLN.GRPCompany.
CREATE OR ALTER PROCEDURE cis.PlanAnnual_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PLNCode) AS Code, ISNULL(PLNDesc, N'') AS Name FROM dbo.PLN ORDER BY PLNCode;
    SELECT RTRIM(INSCode) AS Code, ISNULL(INSDesc, N'') AS Name FROM dbo.INS ORDER BY INSCode;
    SELECT RTRIM(HLTCode) AS Code, ISNULL(HLTDesc, N'') AS Name FROM dbo.HLT ORDER BY HLTCode;
    SELECT RTRIM(PAYCode) AS Code, ISNULL(PAYCompanyName, N'') AS Name FROM dbo.PAY ORDER BY PAYCode;
    SELECT DISTINCT RTRIM(GRPCompany) AS Code, RTRIM(GRPCompany) AS Name FROM dbo.PLN WHERE LEN(GRPCompany) > 0 ORDER BY Code;
END
GO

-- SearchANN → AnnualListing (VIEW_ANN_PAY_PLN).
CREATE OR ALTER PROCEDURE cis.PlanAnnual_Search
    @plnCode nvarchar(4) = NULL, @hltCode nvarchar(4) = NULL, @insCode nvarchar(4) = NULL, @payCode nvarchar(4) = NULL, @grpCompany nvarchar(4) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT HLTCode, AnnualIndex, PLNCode, INSCode, AnnLimit, PDentalAnnLimit, PSpecAnnLimit, PMEPAnnLimit, SuppLimit, SDentalAnnLimit, SSpecAnnLimit,
           SMEPAnnLimit, AnnEffDate, AnnLimitVersion, PAYCode, GRPCompany, SuppLimitStatus
    FROM dbo.VIEW_ANN_PAY_PLN
    WHERE (@plnCode IS NULL OR PLNCode = @plnCode) AND (@hltCode IS NULL OR HLTCode = @hltCode) AND (@insCode IS NULL OR INSCode = @insCode)
      AND (@payCode IS NULL OR PAYCode = @payCode) AND (@grpCompany IS NULL OR GRPCompany LIKE @grpCompany + N'%')
    ORDER BY AnnualIndex;
END
GO

-- SaveANN duplicate check (Add only): insured + health + plan + version + effective date.
CREATE OR ALTER PROCEDURE cis.PlanAnnual_Exists @insCode nvarchar(2), @hltCode nvarchar(2), @plnCode nvarchar(4), @version int, @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.AnnualLimit WHERE INSCode = @insCode AND HLTCode = @hltCode AND PLNCode = @plnCode
                                  AND AnnLimitVersion = @version AND AnnEffDate = @effDate) THEN 1 ELSE 0 END AS bit);
END
GO

-- SaveANN: Add inserts; Edit updates the selected AnnualIndex. Status A leaves the supplementary columns as they are (NULL on Add).
CREATE OR ALTER PROCEDURE cis.PlanAnnual_Save
    @index int = NULL, @plnCode nvarchar(4), @insCode nvarchar(2), @hltCode nvarchar(2), @annLimit money, @pDental money, @pSpec money, @pMep money,
    @status varchar(1), @suppLimit money = NULL, @sDental money = NULL, @sSpec money = NULL, @sMep money = NULL, @effDate datetime,
    @version int, @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF @index IS NULL
        INSERT INTO dbo.AnnualLimit (PLNCode, INSCode, HLTCode, AnnLimit, PDentalAnnLimit, PSpecAnnLimit, PMEPAnnLimit, SuppLimitStatus, SuppLimit,
                                     SDentalAnnLimit, SSpecAnnLimit, SMEPAnnLimit, AnnEffDate, AnnLimitVersion, AnnLastUpdateUser, AnnLastUpdateDate)
        VALUES (@plnCode, @insCode, @hltCode, @annLimit, @pDental, @pSpec, @pMep, @status,
                CASE WHEN @status = 'A' THEN NULL ELSE @suppLimit END, CASE WHEN @status = 'A' THEN NULL ELSE @sDental END,
                CASE WHEN @status = 'A' THEN NULL ELSE @sSpec END, CASE WHEN @status = 'A' THEN NULL ELSE @sMep END, @effDate, @version, @user, @today);
    ELSE
        UPDATE dbo.AnnualLimit
        SET PLNCode = @plnCode, INSCode = @insCode, HLTCode = @hltCode, AnnLimit = @annLimit, PDentalAnnLimit = @pDental, PSpecAnnLimit = @pSpec,
            PMEPAnnLimit = @pMep, SuppLimitStatus = @status,
            SuppLimit = CASE WHEN @status = 'A' THEN SuppLimit ELSE @suppLimit END, SDentalAnnLimit = CASE WHEN @status = 'A' THEN SDentalAnnLimit ELSE @sDental END,
            SSpecAnnLimit = CASE WHEN @status = 'A' THEN SSpecAnnLimit ELSE @sSpec END, SMEPAnnLimit = CASE WHEN @status = 'A' THEN SMEPAnnLimit ELSE @sMep END,
            AnnEffDate = @effDate, AnnLimitVersion = @version, AnnLastUpdateUser = @user, AnnLastUpdateDate = @today
        WHERE AnnualIndex = @index;
    SELECT @@ROWCOUNT;
END
GO

-- DeleteANN: refused while MBM (members) uses the plan; runs only where CISDB has MBM.
CREATE OR ALTER PROCEDURE cis.PlanAnnual_InUse @plnCode nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    IF OBJECT_ID(N'dbo.MBM') IS NULL BEGIN SELECT CAST(0 AS bit); RETURN; END
    EXEC sp_executesql N'SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.MBM WHERE PLNCode = @plnCode) THEN 1 ELSE 0 END AS bit)',
        N'@plnCode nvarchar(10)', @plnCode = @plnCode;
END
GO

CREATE OR ALTER PROCEDURE cis.PlanAnnual_Delete @index int AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.AnnualLimit WHERE AnnualIndex = @index;
    SELECT @@ROWCOUNT;
END
GO

-- ===== main/210_plan_visit_limit.sql =====
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

-- ===== main/220_plan_mco.sql =====
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

-- ===== main/300_member_search.sql =====
-- Member search pop-ups (frmSearchMBM, frmSearchMBMIC, frmSearchMBM05), used by MBM Adjustment / Enquiry / Topup and the invoices.
-- The desktop passed the filter as SQL text to HISDB Search_MBMCLN2013 / Search_MBMCovCLN2013 (dynamic SQL); these are the same
-- joins and columns with parameters. @by: MemberNo, PolicyNo, Name, IC (12+ characters: with and without dashes), IcExact, EmployeeNo,
-- PolicySubNo. At most @max rows each (the name search is LIKE '%text%' over ~1 million members).

-- Search_MBMCLN2013: principals (MBMCLNOthers + MBMCrossReference + MBMOthers).
CREATE OR ALTER PROCEDURE cis.MemberSearch_Principals @by varchar(20), @text nvarchar(100), @max int = 5000 AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @ic2 nvarchar(100) = CASE
        WHEN LEN(@text) < 12 THEN @text
        WHEN CHARINDEX(N'-', @text) > 0 THEN REPLACE(@text, N'-', N'')
        ELSE SUBSTRING(@text, 1, 6) + N'-' + SUBSTRING(@text, 7, 2) + N'-' + SUBSTRING(@text, 9, 4) END;
    -- One branch per search type so each can use its index.
    IF @by = 'MemberNo'
        SELECT TOP (@max) c.MBMNumber, x.MBMCOVID AS CoverId, RTRIM(x.MBMName) AS Name, RTRIM(x.MBMIcBcPp) AS Ic,
               c.MBMCLNStatus AS PolicyStatus, c.MBMCLNDataStatus AS DataStatus, c.MBMPolEffDate AS EffDate, c.MBMPolExpDate AS ExpDate,
               o.MBMGroupCompany AS GroupCompany, o.MBMDepartment AS Department, CAST(0 AS bit) AS Supp
        FROM cis.Hdb_MBMCLNOthers c
        INNER JOIN cis.Maint_MBMCrossReference x ON x.MBMNumber = c.MBMNumber
        LEFT JOIN cis.Hdb_MBMOthers o WITH (NOLOCK) ON o.MBMNumber = c.MBMNumber
        WHERE c.MBMNumber = @text;
    ELSE IF @by = 'PolicyNo'
        SELECT TOP (@max) c.MBMNumber, x.MBMCOVID AS CoverId, RTRIM(x.MBMName) AS Name, RTRIM(x.MBMIcBcPp) AS Ic,
               c.MBMCLNStatus AS PolicyStatus, c.MBMCLNDataStatus AS DataStatus, c.MBMPolEffDate AS EffDate, c.MBMPolExpDate AS ExpDate,
               o.MBMGroupCompany AS GroupCompany, o.MBMDepartment AS Department, CAST(0 AS bit) AS Supp
        FROM cis.Hdb_MBMCLNOthers c
        INNER JOIN cis.Maint_MBMCrossReference x ON x.MBMNumber = c.MBMNumber
        LEFT JOIN cis.Hdb_MBMOthers o WITH (NOLOCK) ON o.MBMNumber = c.MBMNumber
        WHERE x.MBMPolicyNo = @text;
    ELSE IF @by = 'Name'
        SELECT TOP (@max) c.MBMNumber, x.MBMCOVID AS CoverId, RTRIM(x.MBMName) AS Name, RTRIM(x.MBMIcBcPp) AS Ic,
               c.MBMCLNStatus AS PolicyStatus, c.MBMCLNDataStatus AS DataStatus, c.MBMPolEffDate AS EffDate, c.MBMPolExpDate AS ExpDate,
               o.MBMGroupCompany AS GroupCompany, o.MBMDepartment AS Department, CAST(0 AS bit) AS Supp
        FROM cis.Hdb_MBMCLNOthers c
        INNER JOIN cis.Maint_MBMCrossReference x ON x.MBMNumber = c.MBMNumber
        LEFT JOIN cis.Hdb_MBMOthers o WITH (NOLOCK) ON o.MBMNumber = c.MBMNumber
        WHERE x.MBMName LIKE N'%' + @text + N'%';
    ELSE IF @by = 'IC'
        SELECT TOP (@max) c.MBMNumber, x.MBMCOVID AS CoverId, RTRIM(x.MBMName) AS Name, RTRIM(x.MBMIcBcPp) AS Ic,
               c.MBMCLNStatus AS PolicyStatus, c.MBMCLNDataStatus AS DataStatus, c.MBMPolEffDate AS EffDate, c.MBMPolExpDate AS ExpDate,
               o.MBMGroupCompany AS GroupCompany, o.MBMDepartment AS Department, CAST(0 AS bit) AS Supp
        FROM cis.Hdb_MBMCLNOthers c
        INNER JOIN cis.Maint_MBMCrossReference x ON x.MBMNumber = c.MBMNumber
        LEFT JOIN cis.Hdb_MBMOthers o WITH (NOLOCK) ON o.MBMNumber = c.MBMNumber
        WHERE (x.MBMIcBcPp = @text OR x.MBMIcBcPp = @ic2);
    ELSE IF @by = 'IcExact'
        SELECT TOP (@max) c.MBMNumber, x.MBMCOVID AS CoverId, RTRIM(x.MBMName) AS Name, RTRIM(x.MBMIcBcPp) AS Ic,
               c.MBMCLNStatus AS PolicyStatus, c.MBMCLNDataStatus AS DataStatus, c.MBMPolEffDate AS EffDate, c.MBMPolExpDate AS ExpDate,
               o.MBMGroupCompany AS GroupCompany, o.MBMDepartment AS Department, CAST(0 AS bit) AS Supp
        FROM cis.Hdb_MBMCLNOthers c
        INNER JOIN cis.Maint_MBMCrossReference x ON x.MBMNumber = c.MBMNumber
        LEFT JOIN cis.Hdb_MBMOthers o WITH (NOLOCK) ON o.MBMNumber = c.MBMNumber
        WHERE x.MBMIcBcPp = @text;
    ELSE IF @by = 'EmployeeNo'
        SELECT TOP (@max) c.MBMNumber, x.MBMCOVID AS CoverId, RTRIM(x.MBMName) AS Name, RTRIM(x.MBMIcBcPp) AS Ic,
               c.MBMCLNStatus AS PolicyStatus, c.MBMCLNDataStatus AS DataStatus, c.MBMPolEffDate AS EffDate, c.MBMPolExpDate AS ExpDate,
               o.MBMGroupCompany AS GroupCompany, o.MBMDepartment AS Department, CAST(0 AS bit) AS Supp
        FROM cis.Hdb_MBMCLNOthers c
        INNER JOIN cis.Maint_MBMCrossReference x ON x.MBMNumber = c.MBMNumber
        LEFT JOIN cis.Hdb_MBMOthers o WITH (NOLOCK) ON o.MBMNumber = c.MBMNumber
        WHERE o.MBMEmployeeNo = @text;
    ELSE IF @by = 'PolicySubNo'
        SELECT TOP (@max) c.MBMNumber, x.MBMCOVID AS CoverId, RTRIM(x.MBMName) AS Name, RTRIM(x.MBMIcBcPp) AS Ic,
               c.MBMCLNStatus AS PolicyStatus, c.MBMCLNDataStatus AS DataStatus, c.MBMPolEffDate AS EffDate, c.MBMPolExpDate AS ExpDate,
               o.MBMGroupCompany AS GroupCompany, o.MBMDepartment AS Department, CAST(0 AS bit) AS Supp
        FROM cis.Hdb_MBMCLNOthers c
        INNER JOIN cis.Maint_MBMCrossReference x ON x.MBMNumber = c.MBMNumber
        LEFT JOIN cis.Hdb_MBMOthers o WITH (NOLOCK) ON o.MBMNumber = c.MBMNumber
        WHERE o.MBMPOLSubNo1 = @text;
END
GO

-- Search_MBMCovCLN2013: supplementary members (MBMCoveredPersons); the desktop ran it for member number, name and IC only.
CREATE OR ALTER PROCEDURE cis.MemberSearch_Supps @by varchar(20), @text nvarchar(100), @max int = 5000 AS
BEGIN
    SET NOCOUNT ON;
    IF @by = 'MemberNo'
        SELECT TOP (@max) p.MBMCNumber AS MBMNumber, p.MBMCCoverID AS CoverId, RTRIM(p.MBMCName) AS Name, RTRIM(p.MBMCICBCPPNo) AS Ic,
               c.MBMCLNStatus AS PolicyStatus, c.MBMCLNDataStatus AS DataStatus, c.MBMPolEffDate AS EffDate, c.MBMPolExpDate AS ExpDate,
               CAST(NULL AS nvarchar(200)) AS GroupCompany, CAST(NULL AS nvarchar(200)) AS Department, CAST(1 AS bit) AS Supp
        FROM cis.Hdb_MBMCLNOthers c
        INNER JOIN cis.Hdb_MBMCoveredPersons p WITH (NOLOCK) ON p.MBMCNumber = c.MBMNumber
        WHERE p.MBMCNumber = @text;
    ELSE IF @by = 'Name'
        SELECT TOP (@max) p.MBMCNumber AS MBMNumber, p.MBMCCoverID AS CoverId, RTRIM(p.MBMCName) AS Name, RTRIM(p.MBMCICBCPPNo) AS Ic,
               c.MBMCLNStatus AS PolicyStatus, c.MBMCLNDataStatus AS DataStatus, c.MBMPolEffDate AS EffDate, c.MBMPolExpDate AS ExpDate,
               CAST(NULL AS nvarchar(200)) AS GroupCompany, CAST(NULL AS nvarchar(200)) AS Department, CAST(1 AS bit) AS Supp
        FROM cis.Hdb_MBMCLNOthers c
        INNER JOIN cis.Hdb_MBMCoveredPersons p WITH (NOLOCK) ON p.MBMCNumber = c.MBMNumber
        WHERE p.MBMCName LIKE N'%' + @text + N'%';
    ELSE IF @by IN ('IC', 'IcExact')
        SELECT TOP (@max) p.MBMCNumber AS MBMNumber, p.MBMCCoverID AS CoverId, RTRIM(p.MBMCName) AS Name, RTRIM(p.MBMCICBCPPNo) AS Ic,
               c.MBMCLNStatus AS PolicyStatus, c.MBMCLNDataStatus AS DataStatus, c.MBMPolEffDate AS EffDate, c.MBMPolExpDate AS ExpDate,
               CAST(NULL AS nvarchar(200)) AS GroupCompany, CAST(NULL AS nvarchar(200)) AS Department, CAST(1 AS bit) AS Supp
        FROM cis.Hdb_MBMCLNOthers c
        INNER JOIN cis.Hdb_MBMCoveredPersons p WITH (NOLOCK) ON p.MBMCNumber = c.MBMNumber
        WHERE p.MBMCICBCPPNo = @text;
END
GO

-- frmSearchMBM05: member numbers starting with the text (HISDB Search_CISMBM_05 / Search_CISMBMSupp05, already parameterized),
-- group company from CISDB PLN (PlnDesc).
CREATE OR ALTER PROCEDURE cis.MemberSearch_ByPrefix @prefix nvarchar(19) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @like nvarchar(20) = @prefix + N'%';
    DECLARE @p TABLE (MBMNumber nvarchar(400), MBMIcBcPp nvarchar(400), MBMName nvarchar(400), MBMEmployeeNo nvarchar(400), MBMPOLSubNo1 nvarchar(400),
        MBMCOVID nvarchar(400), MBMPolicyNo nvarchar(400), MBMPolEffDate datetime, MBMPolExpDate datetime, PLNType nvarchar(400), MBMDepartment nvarchar(400),
        MBMCLNDataStatus nvarchar(400), MBMCLNStatus nvarchar(400), MBMGroupCompany nvarchar(400), MBMDepartment2 nvarchar(400));
    INSERT INTO @p EXEC cis.Hdb_Search_CISMBM_05 @StrSql = @like;
    DECLARE @s TABLE (MBMCNumber nvarchar(400), MBMCName nvarchar(400), MBMCICBCPPNo nvarchar(400), MBMCCoverID nvarchar(400), MBMPolEffDate datetime,
        MBMPolExpDate datetime, MBMCLNDataStatus nvarchar(400), MBMCLNStatus nvarchar(400), PLNType nvarchar(400), MBMPolicyNo nvarchar(400));
    INSERT INTO @s EXEC cis.Hdb_Search_CISMBMSupp05 @StrSql = @like;
    SELECT p.MBMNumber, p.MBMCOVID AS CoverId, RTRIM(p.MBMName) AS Name, RTRIM(p.MBMIcBcPp) AS Ic, RTRIM(p.MBMPolicyNo) AS PolicyNo,
           RTRIM(p.MBMCLNStatus) AS PolicyStatus, p.MBMPolEffDate AS EffDate, p.MBMPolExpDate AS ExpDate,
           RTRIM(pl.GRPCompany) AS GroupCompany, RTRIM(p.MBMDepartment) AS Department, CAST(0 AS bit) AS Supp
    FROM @p p LEFT JOIN dbo.PLN pl ON pl.PLNCode = RTRIM(p.PLNType)
    UNION ALL
    SELECT s.MBMCNumber, s.MBMCCoverID, RTRIM(s.MBMCName), RTRIM(s.MBMCICBCPPNo), RTRIM(s.MBMPolicyNo), RTRIM(s.MBMCLNStatus), s.MBMPolEffDate,
           s.MBMPolExpDate, RTRIM(pl.GRPCompany), NULL, CAST(1 AS bit)
    FROM @s s LEFT JOIN dbo.PLN pl ON pl.PLNCode = RTRIM(s.PLNType);
END
GO

-- ===== main/310_mbm_topup.sql =====
-- Membership > MBM Topup Adj (FrmMBMTopupLmt): member (HISDB SearchCISMemberProc, MBMOthers), balance (CISDB spUtilisationOP_test1),
-- supplementary limit type (CISDB AnnualLimit + HISDB AnnualLimitType), history (CISDB MBMCISTopupHistory) and the limit / IC updates
-- (HISDB MBMCLNOthers, MBMCoveredPersonsCLN, MBM, MBMCoveredPersons; HISMaintenance MBMCrossReference).

-- FindMBM: group company and department (MBMOthers).
CREATE OR ALTER PROCEDURE cis.MbmTopup_Others @mbmNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) RTRIM(MBMGroupCompany) AS GroupCompany, RTRIM(MBMDepartment) AS Department FROM cis.Hdb_MBMOthers WHERE MBMNumber = @mbmNumber;
END
GO

-- GetSptLmtInd: SuppLimitStatus of the latest AnnualLimit on or before the effective date (blank = A) and its description.
CREATE OR ALTER PROCEDURE cis.MbmTopup_LimitType @plnCode nvarchar(10), @insCode nvarchar(5), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) ISNULL(NULLIF(RTRIM(a.SuppLimitStatus), ''), 'A') AS LimitCode,
           (SELECT TOP (1) RTRIM(t.LimitDesc) FROM cis.Hdb_AnnualLimitType t WHERE t.LimitCode = ISNULL(NULLIF(RTRIM(a.SuppLimitStatus), ''), 'A')) AS LimitDesc
    FROM dbo.AnnualLimit a
    WHERE a.PLNCode = @plnCode AND a.INSCode = @insCode AND a.AnnEffDate <= @effDate
    ORDER BY a.AnnEffDate DESC;
END
GO

-- CmdNext_Click / CmdPreview_Click: the next / previous clinic member number (HSCLNInd <> 'H').
CREATE OR ALTER PROCEDURE cis.MbmTopup_Step @mbmNumber nvarchar(30), @next bit AS
BEGIN
    SET NOCOUNT ON;
    IF @next = 1
        SELECT TOP (1) MBMNumber FROM cis.Hdb_MBM WHERE HSCLNInd <> 'H' AND MBMNumber > @mbmNumber ORDER BY MBMNumber;
    ELSE
        SELECT TOP (1) MBMNumber FROM cis.Hdb_MBM WHERE HSCLNInd <> 'H' AND MBMNumber < @mbmNumber ORDER BY MBMNumber DESC;
END
GO

-- CmdSave_Click: history row, IC change (GP only), new balance limit; one transaction (the caller's).
CREATE OR ALTER PROCEDURE cis.MbmTopup_Save
    @mbmNumber nvarchar(15), @covId char(2), @gp bit, @icNo nvarchar(30), @prevIcNo nvarchar(30), @prevLimit money, @topupAmt money,
    @newLimit money, @limitCode varchar(1), @user nvarchar(50), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF @gp = 1
        INSERT INTO dbo.MBMCISTopupHistory (MBMNumber, MBMCovID, MBMICNo, PrevICNo, PrevLimit, TopupAmt, NewLimit, UpdateDate, UpdateBy)
        VALUES (@mbmNumber, @covId, @icNo, @prevIcNo, @prevLimit, @topupAmt, @newLimit, @today, @user);
    ELSE
        INSERT INTO dbo.MBMCISTopupHistory (MBMNumber, MBMCovID, PrevLimit, SP_TopupAmt, NewLimit, UpdateDate, UpdateBy)
        VALUES (@mbmNumber, @covId, @prevLimit, @topupAmt, @newLimit, @today, @user);

    IF @gp = 1 AND ISNULL(RTRIM(@prevIcNo), N'') <> ISNULL(RTRIM(@icNo), N'')
    BEGIN
        IF @covId = '00'
        BEGIN
            UPDATE cis.Hdb_MBM SET MBMIcBcPp = @icNo WHERE MBMNumber = @mbmNumber;
            UPDATE cis.Maint_MBMCrossReference SET MBMIcBcPp = @icNo WHERE MBMNumber = @mbmNumber;
        END
        ELSE
            UPDATE cis.Hdb_MBMCoveredPersons SET MBMCICBCPPNo = @icNo WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @covId;
    END

    IF @newLimit <> 0
    BEGIN
        IF @limitCode = 'A' OR @covId = '00'
            UPDATE cis.Hdb_MBMCLNOthers SET MBMCLNBalLmt = @newLimit WHERE MBMNumber = @mbmNumber;
        ELSE IF @limitCode <> 'N'
            UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCCLNBalLmt = @newLimit
            WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = CASE WHEN @limitCode = 'B' THEN '01' ELSE @covId END;
    END
END
GO

-- ===== main/320_case_transfer.sql =====
-- Invoice > Other > Case Transfer (frmMembershipTransfer): move a case (CISDB CISCAS) to another member / cover.

-- ListCases: the case's current member (VIEW_CISMBMTransfer).
CREATE OR ALTER PROCEDURE cis.CaseTransfer_Case @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) CASNumber, RTRIM(MBMNumber) AS MbmNumber, RTRIM(CoverID) AS CoverId, RTRIM(MBMName) AS MbmName, RTRIM(MBMCName) AS MbmcName,
           DateVisit
    FROM dbo.VIEW_CISMBMTransfer WHERE CASNumber = @casNumber;
END
GO

-- cmdSearch_Click: the transfer member's name (cover 00: HISDB MBM; otherwise MBMCoveredPersons).
CREATE OR ALTER PROCEDURE cis.CaseTransfer_MemberName @mbmNumber nvarchar(50), @covId nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    IF @covId = '00'
        SELECT TOP (1) RTRIM(MBMName) AS Name FROM cis.Hdb_MBM WHERE MBMNumber = @mbmNumber;
    ELSE
        SELECT TOP (1) RTRIM(MBMCName) AS Name FROM cis.Hdb_MBMCoveredPersons WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @covId;
END
GO

-- cmdSave_Click: the case's member number and cover id.
CREATE OR ALTER PROCEDURE cis.CaseTransfer_Save @casNumber nvarchar(20), @mbmNumber nvarchar(50), @covId nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET MBMnumber = @mbmNumber, CoverID = @covId WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO

-- ===== main/330_member_registration.sql =====
-- Clinic member registration pop-up (frmMemberRegistrationCLN, opened from the clinic invoice): look-ups, limits, age band and the new
-- DC member (HISDB MBM, MBMTwo, MBMOthers, MBMOthersTwo, MBMCLNOthers; HISMaintenance MBMCrossReference, MBMSEQ).

-- ComboListing: races (HISDB SearchRace), staff categories (CISDB WebCategory), group company names (CISDB GRPCompany).
CREATE OR ALTER PROCEDURE cis.MbmReg_Setup AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Hdb_SearchRace;
    SELECT RTRIM(Category) AS Category, RTRIM(INSType) AS InsType, RTRIM(PlanCode) AS PlanCode FROM dbo.WebCategory;
    SELECT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL;
END
GO

-- cmdSearch_Click: departments of group company G001 containing the text.
CREATE OR ALTER PROCEDURE cis.MbmReg_Departments @text nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GrpCompDept) AS Department FROM dbo.GroupCompanyDept WHERE GrpCompCode = 'G001' AND GrpCompDept LIKE N'%' + @text + N'%';
END
GO

-- CallCLNLimit: annual limit (PLN.ANNLmtInd N → 0, else GetCLNAnnualLimit) and visit limit (PLN.VISLmtInd N → 0, else GetVisitLimit:
-- the DC plan row with VISLmtInd N = "FULL" (NULL here), else the latest VisitLimit), health type N, principal.
CREATE OR ALTER PROCEDURE cis.MbmReg_Limits @plnCode nvarchar(4), @insCode nvarchar(2), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @ann money = 0, @visit money = 0, @full bit = 0;
    IF ISNULL((SELECT TOP (1) RTRIM(ANNLmtInd) FROM dbo.PLN WHERE PLNCode = @plnCode), '') <> 'N'
        SET @ann = ISNULL((SELECT TOP (1) AnnLimit FROM dbo.AnnualLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND HLTCode = 'N'
                           AND AnnEffDate <= @effDate ORDER BY AnnEffDate DESC), 0);
    IF EXISTS (SELECT 1 FROM dbo.PLN WHERE PLNCode = @plnCode) AND ISNULL((SELECT TOP (1) RTRIM(VISLmtInd) FROM dbo.PLN WHERE PLNCode = @plnCode), '') <> 'N'
    BEGIN
        IF EXISTS (SELECT 1 FROM dbo.PLN WHERE PLNCode = @plnCode AND PAYCode = 'DC' AND RTRIM(VISLmtInd) = 'N') SET @full = 1;
        ELSE SET @visit = ISNULL((SELECT TOP (1) VisitLimit FROM dbo.VisitLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND HLTCode = 'N'
                                  AND VisitEffDate <= @effDate ORDER BY VisitEffDate DESC), 0);
    END
    SELECT @ann AS AnnualLimit, @visit AS VisitLimit, @full AS VisitFull;
END
GO

-- GetAgeBand: years against AGEFrom / AGETo ("…Y"), days against AGEFrom ("…D"), HISDB AGE.
CREATE OR ALTER PROCEDURE cis.MbmReg_AgeBand @age int, @days bit AS
BEGIN
    SET NOCOUNT ON;
    IF @days = 0
        SELECT TOP (1) RTRIM(AGECode) AS AgeCode, RTRIM(AGEDescription) AS AgeDescription FROM cis.Hdb_AGE
        WHERE TRY_CAST(SUBSTRING(AGEFrom, 1, LEN(AGEFrom) - 1) AS int) <= @age AND TRY_CAST(SUBSTRING(AGETo, 1, LEN(AGETo) - 1) AS int) >= @age
        ORDER BY AGECode;
    ELSE
        SELECT TOP (1) RTRIM(AGECode) AS AgeCode, RTRIM(AGEDescription) AS AgeDescription FROM cis.Hdb_AGE
        WHERE TRY_CAST(SUBSTRING(AGEFrom, 1, LEN(AGEFrom) - 1) AS int) <= @age ORDER BY AGEFrom;
END
GO

-- cmdReg_Click: GenerateMembershipNo("DC", insured type) + GenerateChkDigits, then the six inserts; one transaction (the caller's).
CREATE OR ALTER PROCEDURE cis.MbmReg_Register
    @ic nvarchar(30), @name nvarchar(200), @sex nvarchar(1), @adultChild nvarchar(1), @ageCode nvarchar(2), @dob datetime, @race nvarchar(5),
    @groupCompany nvarchar(200), @employeeNo nvarchar(50), @department nvarchar(200), @plnCode nvarchar(4), @insCode nvarchar(2),
    @effDate datetime, @expDate datetime, @visitLimit money, @annualLimit money, @user nvarchar(50), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seqCode nvarchar(10) = N'DC' + @insCode, @seq nvarchar(10), @ch nchar(1), @n int, @num nvarchar(10);
    SELECT @seq = RTRIM(SEQNumber) FROM cis.Maint_MBMSEQ WITH (UPDLOCK, HOLDLOCK) WHERE SEQCode = @seqCode;
    IF @seq IS NULL
    BEGIN
        SET @num = N'A00001';
        INSERT INTO cis.Maint_MBMSEQ (SEQCode, PAYCode, INSCode, SEQNumber) VALUES (@seqCode, N'DC', @insCode, @num);
    END
    ELSE
    BEGIN
        SELECT @ch = LEFT(@seq, 1), @n = TRY_CAST(SUBSTRING(@seq, 2, 5) AS int);
        IF @n < 99999 SET @n += 1;
        ELSE SELECT @ch = NCHAR(UNICODE(@ch) + 1), @n = 1;
        SET @num = @ch + RIGHT(N'00000' + CAST(@n AS nvarchar(10)), 5);
        UPDATE cis.Maint_MBMSEQ SET SEQNumber = @num WHERE SEQCode = @seqCode;
    END

    DECLARE @d nvarchar(2) = (SELECT TOP (1) RTRIM(MBMChkDigits) FROM cis.Hdb_MBMCheckDigits WITH (UPDLOCK, HOLDLOCK));
    SET @d = CASE WHEN @d = N'99' OR TRY_CAST(@d AS int) IS NULL THEN N'01' ELSE RIGHT(N'00' + CAST(CAST(@d AS int) + 1 AS nvarchar(3)), 2) END;
    UPDATE cis.Hdb_MBMCheckDigits SET MBMChkDigits = @d;

    DECLARE @mbm nvarchar(30) = N'DC' + @insCode + @num + @d;
    INSERT INTO cis.Hdb_MBM (MBMNumber, MBMIcBcPp, MBMName, MBMSex, MBMAdultChild, AGECode, MBMValueDate, MBMBordxDate, MBMMemberType, HLTCode,
                             MBMBatchNo, MBMCOVID, HSCLNInd, MBMExpiredStatus, MBMDOB, RACCode)
    VALUES (@mbm, @ic, @name, @sex, @adultChild, @ageCode, @today, @today, 'P', 'N', '1', '00', 'C', 'N', @dob, @race);
    INSERT INTO cis.Hdb_MBMTwo (MBMNumber, MBMSalutation, MBMTelnoH, MBMTelNoM, MBMTelNoO, MBMEmail, NATCode, SRVCodeList, LGNCode, MBMRegUser,
                                MBMLastUpdateUser, MBMUploadStatus)
    VALUES (@mbm, '', '', '', '', '', '', 'ME', '', @user, @user, 'M');
    INSERT INTO cis.Hdb_MBMOthers (MBMNumber, MBMGroupCompany, MBMEmployeeNo, MBMDepartment, RELCODE) VALUES (@mbm, @groupCompany, @employeeNo, @department, 'P');
    INSERT INTO cis.Hdb_MBMOthersTwo (MBMNumber) VALUES (@mbm);
    INSERT INTO cis.Maint_MBMCrossReference (MBMNumber, MBMIcBcPp, MBMName, MBMAdultChild, AGECode, MBMBordxDate, MBMMemberType, HLTCode, MBMBatchNo,
                                             MBMCOVID, MBMStatus)
    VALUES (@mbm, @ic, @name, @adultChild, @ageCode, @today, 'P', 'N', '1', '00', 'A');
    INSERT INTO cis.Hdb_MBMCLNOthers (MBMNumber, PAYCode, PLNType, InsType, MBMPolEffDate, MBMPolExpDate, MBMCLNMemberType, MBMCLNExpiryStatus,
                                      MBMCLNDataStatus, MBMCLNStatus, MBMCLNVisitLmt, MBMCLNBalLmt)
    VALUES (@mbm, 'DC', @plnCode, @insCode, @effDate, @expDate, 'P', 'N', 'A', 'A', @visitLimit, @annualLimit);
    SELECT @mbm;
END
GO

-- ===== main/340_general_exclusions.sql =====
-- Invoice > General Exclusions (Enquiry) (frmGeneralBenefitExcEnq): CISDB GeneralExclusions per group company (read-only; the desktop's
-- Update button is hidden). frmGeneralBenefitExclusion (the editing form) is not in the source.

-- ComboList: group company names.
CREATE OR ALTER PROCEDURE cis.GeneralExclusions_GroupCompanies AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL;
END
GO

-- CallExclusions: the group company's exclusion flags and remarks.
CREATE OR ALTER PROCEDURE cis.GeneralExclusions_Get @groupCompany nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) Dermatitis, Gynaecological, Dental, Optical, MedCheckup, PreEmpCheckup, Hypertension, Diabetes, IHDCAD, Thyroid, Immunisation, HRT,
           PregnacyTest, Antenatal, Abortion, Infertility, Contraception, Genitourinary, Vitamins, BloodTest, XRAY, PapSmear, Typhoid, LTMedication, Remarks
    FROM dbo.GeneralExclusions WHERE GRPCompany = @groupCompany;
END
GO

-- ===== main/350_mbm_adjustment.sql =====
-- Membership > MBM Adjustment (FrmMBMCLNAdj): look-ups (Medix.bas GetPremium / GetMCO / GetAnnualLimit / GetVisitLimit), history tabs and
-- the endorsement writes (CISDB MBMCISAmendHistory, MBMCISRefund; HISDB MBMCLNOthers, MBMCoveredPersonsCLN, MBMTwo, MBMOthersTwo).
-- Member rows: HISDB SearchCISMemberProc (cis.Hdb_SearchCISMemberProc), group company: cis.MbmTopup_Others, age band: cis.MbmReg_AgeBand.

-- ComboListing: card types (CheckItems CSMD), relationships (HISDB REL).
CREATE OR ALTER PROCEDURE cis.MbmAdj_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(CatName) AS CardType FROM dbo.CheckItems WHERE CatCode = 'CSMD';
    SELECT RTRIM(RELCode) AS Code, RTRIM(RELDescription) AS Name FROM cis.Hdb_REL;
END
GO

-- showMBM: SuppLimitStatus of plan + insured type (any date, first row; blank = A; none = N).
CREATE OR ALTER PROCEDURE cis.MbmAdj_LimitStatus @plnCode nvarchar(10), @insCode nvarchar(5) AS
BEGIN
    SET NOCOUNT ON;
    SELECT ISNULL((SELECT TOP (1) ISNULL(NULLIF(RTRIM(SuppLimitStatus), ''), 'A') FROM dbo.AnnualLimit WHERE PLNCode = @plnCode AND INSCode = @insCode), 'N');
END
GO

-- showMBMCov: the latest SuppLimitStatus on or before the policy date (blank / none = A).
CREATE OR ALTER PROCEDURE cis.MbmAdj_CoverLimitStatus @plnCode nvarchar(10), @insCode nvarchar(5), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT ISNULL((SELECT TOP (1) NULLIF(RTRIM(SuppLimitStatus), '') FROM dbo.AnnualLimit WHERE PLNCode = @plnCode AND INSCode = @insCode
                   AND AnnEffDate <= @effDate ORDER BY AnnEffDate DESC), 'A');
END
GO

-- GetPremium: PLN.PRMInd must be Y for plan + payor; PRM row (first, as the desktop) on or before the date; P = PRMAmount, S = SuppPRMAmt.
-- Found = 0 → "Age Out of Premium Range!".
CREATE OR ALTER PROCEDURE cis.MbmAdj_Premium
    @mbmType char(1), @insCode nvarchar(5), @payCode nvarchar(5), @ageCode nvarchar(5), @plnCode nvarchar(10), @hltCode nvarchar(2), @effDate datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM dbo.PLN WHERE PLNCode = @plnCode AND PAYCode = @payCode AND RTRIM(PRMInd) = 'Y')
    BEGIN SELECT CAST(0 AS money) AS Amount, CAST(1 AS bit) AS Found; RETURN; END
    SELECT TOP (1) CAST(ISNULL(CASE WHEN @mbmType = 'S' THEN SuppPRMAmt ELSE PRMAmount END, 0) AS money) AS Amount, CAST(1 AS bit) AS Found
    FROM dbo.PRM
    WHERE INSCode = @insCode AND HLTCode = @hltCode AND PLNCode = @plnCode AND AGECode = @ageCode AND PRMEffDate <= @effDate
      AND (@hltCode = 'S' OR PAYCode = @payCode);
    IF @@ROWCOUNT = 0 SELECT CAST(0 AS money) AS Amount, CAST(0 AS bit) AS Found;
END
GO

-- GetMCO: PLN.MCOInd must be Y; the plan's latest MCO row, else the insured type's latest; column by adult / child and P / S. Found = 0 → message.
CREATE OR ALTER PROCEDURE cis.MbmAdj_Mco
    @mbmType char(1), @insCode nvarchar(5), @hltCode nvarchar(2), @adultChild char(1), @effDate datetime, @plnCode nvarchar(10), @payCode nvarchar(5)
AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM dbo.PLN WHERE PLNCode = @plnCode AND PAYCode = @payCode AND RTRIM(MCOInd) = 'Y')
    BEGIN SELECT CAST(0 AS money) AS Amount, CAST(1 AS bit) AS Found; RETURN; END
    DECLARE @amt money, @found bit = 0;
    SELECT TOP (1) @amt = CASE WHEN @adultChild = 'A' THEN CASE WHEN @mbmType = 'S' THEN MCOSuppAdultAmt ELSE MCOAmtAdult END
                               ELSE CASE WHEN @mbmType = 'S' THEN MCOSuppChildAmt ELSE MCOAmtChild END END, @found = 1
    FROM dbo.MCO WHERE INSCode = @insCode AND HLTCode = @hltCode AND PLNCode = @plnCode AND MCOEffDate <= @effDate ORDER BY MCOEffDate DESC;
    IF @found = 0
        SELECT TOP (1) @amt = CASE WHEN @adultChild = 'A' THEN CASE WHEN @mbmType = 'S' THEN MCOSuppAdultAmt ELSE MCOAmtAdult END
                                   ELSE CASE WHEN @mbmType = 'S' THEN MCOSuppChildAmt ELSE MCOAmtChild END END, @found = 1
        FROM dbo.MCO WHERE INSCode = @insCode AND HLTCode = @hltCode AND MCOEffDate <= @effDate ORDER BY MCOEffDate DESC;
    SELECT CAST(ISNULL(@amt, 0) AS money) AS Amount, @found AS Found;
END
GO

-- GetAnnualLimit / GetVisitLimit (health N): Full = 1 is the desktop's "FULL"; Status = the AnnualLimit SuppLimitStatus it returns.
CREATE OR ALTER PROCEDURE cis.MbmAdj_Limits
    @mbmType char(1), @insCode nvarchar(5), @plnCode nvarchar(10), @effDate datetime, @payCode nvarchar(5), @limitStatus varchar(1)
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @annInd varchar(1), @annSupp varchar(1), @shared bit, @visInd varchar(1);
    SELECT TOP (1) @annInd = RTRIM(ANNLmtInd), @annSupp = RTRIM(ANNLmtIndSupp), @shared = ISNULL(IsPrinDepGPShared, 0), @visInd = RTRIM(VISLmtInd)
    FROM dbo.PLN WHERE PLNCode = @plnCode AND PAYCode = @payCode;
    DECLARE @annFull bit = CASE WHEN (@mbmType = 'P' AND @annInd = 'N') OR (@mbmType = 'S' AND ISNULL(@limitStatus, '') <> 'B' AND @annSupp = 'N') THEN 1 ELSE 0 END;
    DECLARE @ann money = 0, @annOk bit = 1, @status varchar(1) = @limitStatus;
    IF @annFull = 0
    BEGIN
        DECLARE @annLimit money, @supp money, @aso money, @st varchar(1), @hasRow bit = 0;
        SELECT TOP (1) @annLimit = AnnLimit, @supp = SuppLimit, @aso = ISNULL(ASOLimit, 0), @st = RTRIM(SuppLimitStatus), @hasRow = 1
        FROM dbo.AnnualLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND HLTCode = 'N' AND AnnEffDate <= @effDate ORDER BY AnnEffDate DESC;
        IF @hasRow = 1
        BEGIN
            SET @status = @st;
            IF ISNULL(@st, '') = '' SELECT @ann = 0, @annOk = 0;
            ELSE IF @st = 'D' SET @annFull = 1;
            ELSE SET @ann = ISNULL(@annLimit, 0) + @aso;
            IF @mbmType = 'S'
                IF @shared = 1 SELECT @ann = ISNULL(@annLimit, 0) + @aso, @annFull = 0;
                ELSE IF ISNULL(@st, '') <> 'A' SELECT @ann = ISNULL(@supp, 0), @annFull = 0;
        END
    END
    DECLARE @visFull bit = CASE WHEN @visInd = 'N' THEN 1 ELSE 0 END, @visit money = 0;
    IF @visFull = 0
        SELECT TOP (1) @visit = CASE WHEN @mbmType = 'S' AND ISNULL(RTRIM(SuppVisitLimitStatus), '') <> 'A' THEN ISNULL(SuppVisitLimit, 0) ELSE ISNULL(VisitLimit, 0) END
        FROM dbo.VisitLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND HLTCode = 'N' AND VisitEffDate <= @effDate ORDER BY VisitEffDate DESC;
    SELECT @ann AS AnnualLimit, @annFull AS AnnualFull, @annOk AS AnnualOk, @visit AS VisitLimit, @visFull AS VisitFull, ISNULL(@status, '') AS Status;
END
GO

-- showMBMCov: refund figures of a cancelled membership.
CREATE OR ALTER PROCEDURE cis.MbmAdj_Refunds @mbmNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(MBMCoveredID) AS CoverId, MBMPremReduction AS PremReduction, MBMPremEffective AS PremEffective, MBMMCORefund AS McoRefund,
           MBMMCOEffective AS McoEffective
    FROM dbo.MBMCISRefund WHERE MBMNumber = @mbmNumber;
END
GO

-- CmdSave_Click (C): an outstanding case blocks the cancellation.
CREATE OR ALTER PROCEDURE cis.MbmAdj_OpenCase @mbmNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.CISCAS WHERE MBMnumber = @mbmNumber AND CASStatus = 'O') THEN 1 ELSE 0 END AS bit);
END
GO

-- ADJListing: CISDB MBMCISAmendHistory.
CREATE OR ALTER PROCEDURE cis.MbmAdj_AmendHistory @mbmNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT MBMAEndtBordxDate AS BordxDate, MBMAmendDate AS AmendDate, MBMAmendEdtType AS EndtType, MBMAEndtEffDate AS EndtEffDate, RTRIM(MBMAmendUser) AS AmendUser
    FROM dbo.MBMCISAmendHistory WHERE MBMNumber = @mbmNumber ORDER BY MBMAmendID;
END
GO

-- Write2MBMnHistory: history row from the current MBMCLNOthers, then the endorsement on MBMCLNOthers.
CREATE OR ALTER PROCEDURE cis.MbmAdj_Endorse
    @mbmNumber nvarchar(30), @endtType char(1), @status nvarchar(1), @bordxDate datetime, @endtEffDate datetime, @user nvarchar(50), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.MBMCISAmendHistory (MBMNumber, MBMPolicyNo, MBMEffDate, MBMExpiryDate, MBMAmendEdtType, MBMAEndtBordxDate, MBMAEndtEffDate, MBMAmendUser, MBMAmendDate)
    SELECT TOP (1) MBMNumber, MBMPolNumber, MBMPolEffDate, MBMPolExpDate, @endtType, @bordxDate, @endtEffDate, @user, @today
    FROM cis.Hdb_MBMCLNOthers WHERE MBMNumber = @mbmNumber;
    UPDATE cis.Hdb_MBMCLNOthers
    SET MBMCLNLastEndtStatus = @endtType, MBMCLNStatus = @status, MBMCLNEndtEffDate = @endtEffDate, MBMCancelDate = @endtEffDate, MBMCLNEndtBordxDate = @bordxDate
    WHERE MBMNumber = @mbmNumber;
END
GO

-- WriteAdjStatus2Cov: every cover's status, or one cover's (K, with today's cancel date).
CREATE OR ALTER PROCEDURE cis.MbmAdj_CoverStatus @mbmNumber nvarchar(30), @status nvarchar(1), @covId nvarchar(2) = NULL, @today datetime AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Hdb_MBMCoveredPersonsCLN
    SET MBMCCLNStatus = @status, MBMCCancelDate = CASE WHEN @covId IS NULL THEN MBMCCancelDate ELSE @today END
    WHERE MBMCNumber = @mbmNumber AND (@covId IS NULL OR MBMCCoverID = @covId);
END
GO

-- MemberMCORefund: replace the cover's MBMCISRefund row.
CREATE OR ALTER PROCEDURE cis.MbmAdj_Refund
    @mbmNumber nvarchar(30), @covId nvarchar(2), @premEffective money, @premReduction money, @mcoEffective money, @mcoRefund money
AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.MBMCISRefund WHERE MBMNumber = @mbmNumber AND MBMCoveredID = @covId;
    INSERT INTO dbo.MBMCISRefund (MBMNumber, MBMCoveredID, MBMPremEffective, MBMPremReduction, MBMMCOEffective, MBMMCORefund)
    VALUES (@mbmNumber, @covId, @premEffective, @premReduction, @mcoEffective, @mcoRefund);
END
GO

-- Amendment (A): SaveCLNOthers + SaveMBMTwo + SaveMBMOthersTwo.
CREATE OR ALTER PROCEDURE cis.MbmAdj_SavePrincipal
    @mbmNumber nvarchar(30), @policyNo nvarchar(100), @effDate datetime = NULL, @expDate datetime = NULL, @remarks nvarchar(max), @cardType nvarchar(50),
    @mcoFee money, @premium money, @dataStatus nvarchar(1), @status nvarchar(1), @warranty nvarchar(max), @plnCode nvarchar(10),
    @email nvarchar(200), @mobile nvarchar(50), @bankName nvarchar(200), @bankAccount nvarchar(100)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Hdb_MBMCLNOthers
    SET MBMPolNumber = @policyNo, MBMPolEffDate = ISNULL(@effDate, MBMPolEffDate), MBMPolExpDate = ISNULL(@expDate, MBMPolExpDate), MBMRemarks = @remarks,
        MBMCardType = @cardType, MBMCLNMCOFee = @mcoFee, MBMCLNPremium = @premium, MBMCLNDataStatus = @dataStatus, MBMCLNStatus = @status,
        MBMCLNWarranty = @warranty, PLNType = @plnCode
    WHERE MBMNumber = @mbmNumber;
    UPDATE cis.Hdb_MBMTwo SET MBMEmail = @email, MBMTelNoM = @mobile WHERE MBMNumber = @mbmNumber;
    UPDATE cis.Hdb_MBMOthersTwo SET MBMBankName = @bankName, MBMBankACNo = @bankAccount WHERE MBMNumber = @mbmNumber;
END
GO

-- SaveCisCovPersons: one cover.
CREATE OR ALTER PROCEDURE cis.MbmAdj_SaveCover
    @mbmNumber nvarchar(30), @covId nvarchar(2), @remarks nvarchar(max), @mco money, @premium money, @warranty nvarchar(max), @cardType nvarchar(50),
    @plnCode nvarchar(10), @effDate datetime = NULL, @expDate datetime = NULL
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Hdb_MBMCoveredPersonsCLN
    SET MBMCCLNRemarks = @remarks, MBMCCLNMCO = @mco, MBMCCLNPrem = @premium, MBMCCLNWarranty = @warranty, MBMCCardType = @cardType,
        MBMCCLNPLNCode = @plnCode, MBMCCLNEffDate = @effDate, MBMCCLNExpDate = @expDate
    WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @covId;
END
GO

-- ===== main/360_mbm_enquiry.sql =====
-- Membership > MBM Enquiry (FrmMBMCLNEnq): member (HISDB SearchCISMemberProc_test), plan look-ups, top-ups, utilisation, cases.
-- Limits / age band / histories reuse the MBM Adjustment procedures (cis.MbmAdj_*).

-- FindMBM: SearchCISMemberProc_test.
CREATE OR ALTER PROCEDURE cis.MbmEnq_Member @mbmNumber varchar(64) AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Hdb_SearchCISMemberProc_test @MBMNumber = @mbmNumber;
END
GO

-- showMBM / cmdShow_Click look-ups of the member's plan: AnnualLimit (first row of plan + insured type), maternity limit, specialist
-- limit status, CISDB PLN flags, the product caption (HISDB MBM + PLN) and the FlexCare check (tblFlexibleBenefitPlanLimit_OP).
CREATE OR ALTER PROCEDURE cis.MbmEnq_Plan @mbmNumber nvarchar(30), @plnCode nvarchar(10), @insCode nvarchar(5), @expDate datetime = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 SuppLimitStatus, PDentalAnnLimit, SDentalAnnLimit, PSpecAnnLimit, SSpecAnnLimit, AnnLimit
    FROM dbo.AnnualLimit WHERE PLNCode = @plnCode AND INSCode = @insCode;
    SELECT TOP 1 a.AnnMaternityLmt FROM dbo.AnnualLimit a INNER JOIN dbo.PLN p ON p.PLNCode = a.PLNCode
    WHERE p.PLNMaternityInd = 'Y' AND p.PLNCode = @plnCode;
    SELECT TOP 1 SuppLimitStatus FROM dbo.SpecialistLimit WHERE PLNCode = @plnCode AND INSCode = @insCode;
    SELECT TOP 1 PLNMaternityInd, PDentalStatus FROM dbo.PLN WHERE PLNCode = @plnCode;
    SELECT TOP 1 p.ProductName, p.PLNDescription FROM cis.Hdb_MBM m INNER JOIN cis.Hdb_PLN p ON p.PLNCode = m.PLNCode WHERE m.MBMNumber = @mbmNumber;
    SELECT COUNT(*) FROM cis.Hdb_MBMCLNOthers o INNER JOIN dbo.tblFlexibleBenefitPlanLimit_OP f ON f.CISPlnCode = o.PLNType
    WHERE o.MBMNumber = @mbmNumber AND @expDate BETWEEN f.FBEffDate AND f.FBExpDate;
END
GO

-- showMBM: top-up totals (MBMCISTopupHistory) of all covers (@covId NULL) or one cover.
CREATE OR ALTER PROCEDURE cis.MbmEnq_Topup @mbmNumber nvarchar(30), @covId nvarchar(2) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 SUM(COALESCE(TopupAmt, 0)) AS TopupAmt, SUM(COALESCE(SP_TopupAmt, 0)) AS SpTopupAmt
    FROM dbo.MBMCISTopupHistory WHERE MBMNumber = @mbmNumber AND (@covId IS NULL OR MBMCovID = @covId) GROUP BY MBMNumber;
END
GO

-- CallLimit (ASO members): Tbl_ASOMemberDetails.
CREATE OR ALTER PROCEDURE cis.MbmEnq_Aso @mbmNumber nvarchar(30), @covId nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    SELECT MBMASOLimit, MBMIPLimit, MBMOPLimit, MBMIPUtilized, MBMOPUtilized, IsIPOPCombined
    FROM cis.Hdb_ASOMemberDetails WHERE MBMNumber = @mbmNumber AND MBMCoverID = @covId;
END
GO

-- Medix.bas GetSpecLimit: PLN.PSpecLimit of plan + payor (last row), SpecialistLimit of insured type + plan, health type N.
CREATE OR ALTER PROCEDURE cis.MbmEnq_SpecLimit @plnCode nvarchar(10), @payCode nvarchar(10), @insCode nvarchar(5) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PSpecLimit) AS SpecLmtInd FROM dbo.PLN WHERE PLNCode = @plnCode AND PAYCode = @payCode;
    SELECT TOP 1 PrinLimit, SuppLimit, RTRIM(SuppLimitStatus) AS SuppLimitStatus FROM dbo.SpecialistLimit
    WHERE INSCode = @insCode AND PLNCode = @plnCode AND HLTCode = 'N';
END
GO

-- CIS_FFS2018_test with its dynamic filter replaced by parameters: the member's claims, all covers (@coverMode 0), cover 00 (1),
-- covers other than 00 (2) or one cover (3, @covId).
CREATE OR ALTER PROCEDURE cis.MbmEnq_Utilisation @mbmNumber nvarchar(30), @coverMode int = 0, @covId nvarchar(2) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT CASE
             WHEN '305' IN (c.CASTreatReason1, c.CASTreatReason2, c.CASTreatReason3, c.CASTreatReason4, c.CASTreatReason5) THEN 'OPTICAL'
             WHEN '303' IN (c.CASTreatReason1, c.CASTreatReason2, c.CASTreatReason3, c.CASTreatReason4, c.CASTreatReason5) THEN 'DENTAL'
             WHEN '302' IN (c.CASTreatReason1, c.CASTreatReason2, c.CASTreatReason3, c.CASTreatReason4, c.CASTreatReason5) THEN 'MATERNITY'
             WHEN '465' IN (c.CASTreatReason1, c.CASTreatReason2, c.CASTreatReason3, c.CASTreatReason4, c.CASTreatReason5) THEN 'DENTAL EXTRACTION (EXCLUDE IN DENTAL LIMIT)'
             WHEN '467' IN (c.CASTreatReason1, c.CASTreatReason2, c.CASTreatReason3, c.CASTreatReason4, c.CASTreatReason5) THEN 'DENTALEXTRACTION'
             WHEN cln.CLNType = 'G' THEN 'GP'
             WHEN cln.CLNType = 'S' THEN 'SP'
           END AS TreatmentType,
           cl.CLMMCOProcFee AS FFS, cln.CLNType, ia.CLMMaternity, c.MBMnumber, c.CoverID, c.CASNumber,
           CASE WHEN p.CoPayInd = 'Y' AND COALESCE(cl.CLMCoPayColl, 0) <> 0 THEN cl.CLMTotalActual - cl.CLMCoPayColl ELSE cl.CLMTotalActual END AS CLMBordxAmt,
           p.PLNGPSPShared,
           CASE WHEN tr1.ReasonCategory IN ('dental', 'optical') AND COALESCE(p.PDentalStatus, 'N') <> 'N' THEN 0
                ELSE COALESCE(ia.CLMOthersNoDed, 0.00) END AS CLMOthersNoDed,
           ia.CLMTyphiadJabNoDed,
           CASE WHEN 'dental' IN (tr1.ReasonCategory, tr2.ReasonCategory, tr3.ReasonCategory, tr4.ReasonCategory, tr5.ReasonCategory)
                     AND ia.CLMDental = 0 AND ia.CLMOthersNoDed = 0 THEN cl.CLMBordxAmt ELSE ia.CLMDental END AS CLMDental,
           CASE WHEN c.CASTreatReason1 IN ('303', '305', '467') THEN cl.CLMBordxAmt ELSE 0 END AS Utilization,
           CASE WHEN c.CASTreatReason1 NOT IN ('303', '305', '467') AND COALESCE(ia.CLMOthersNoDed, 0.00) = 0 THEN COALESCE(cl.CLMBTBGAmt, 0) ELSE 0 END AS CLMBTBGAmt,
           c.OPBatchNo, p.IsGPDentalShared, CLMGPFeeBasicMed, p.PLNCode, SDentalAnnLimit, SDentalStatus, p.PSpecLimit, p.PSpecStatus, p.SSpecLimit,
           p.SSpecStatus, PSpecAnnLimit, SSpecAnnLimit
    FROM dbo.CISCas c WITH (NOLOCK)
    INNER JOIN cis.Ws_Claim cl WITH (NOLOCK) ON c.CASNumber = cl.CASNumber
    INNER JOIN cis.Hdb_MBMCLNOthers o ON o.MBMNumber = c.MBMnumber
    INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    LEFT JOIN dbo.AnnualLimit al ON al.PLNCode = p.PLNCode AND al.INSCode IN ('H')
    INNER JOIN cis.Ws_ClaimInvoiceActualNew ia WITH (NOLOCK) ON cl.CASNumber = ia.CASNumber
    LEFT OUTER JOIN cis.Hdb_CLN cln WITH (NOLOCK) ON c.ClinicCode = cln.CLNCode
    LEFT JOIN dbo.TreatmentReason tr1 ON tr1.ReasonCode = c.CASTreatReason1
    LEFT JOIN dbo.TreatmentReason tr2 ON tr2.ReasonCode = c.CASTreatReason2
    LEFT JOIN dbo.TreatmentReason tr3 ON tr3.ReasonCode = c.CASTreatReason3
    LEFT JOIN dbo.TreatmentReason tr4 ON tr4.ReasonCode = c.CASTreatReason4
    LEFT JOIN dbo.TreatmentReason tr5 ON tr5.ReasonCode = c.CASTreatReason5
    WHERE c.CASStatus <> 'D' AND c.MBMnumber = @mbmNumber
      AND (@coverMode = 0 OR (@coverMode = 1 AND c.CoverID = '00') OR (@coverMode = 2 AND c.CoverID <> '00') OR (@coverMode = 3 AND c.CoverID = @covId));
END
GO

-- tmppridentallimit / Dentallimit_OneLimit: the existing CISDB dental procedures (parameterised).
CREATE OR ALTER PROCEDURE cis.MbmEnq_Dental @mbmNumber nvarchar(30), @covId nvarchar(4) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.tmpDlimit_local_test1 @MBMNumber = @mbmNumber, @MBMCCoverID = @covId;
END
GO

CREATE OR ALTER PROCEDURE cis.MbmEnq_DentalOneLimit @mbmNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.tmpDentallimit_OneLimit @MBMNumber = @mbmNumber;
END
GO

-- ListCases / ListSpecialistCases: VIEW_CISMBMInquiry_GPCASE / VIEW_CISMBMInquiry_Spe_CASE.
CREATE OR ALTER PROCEDURE cis.MbmEnq_Cases @mbmNumber nvarchar(30), @specialist bit AS
BEGIN
    SET NOCOUNT ON;
    IF @specialist = 0
        SELECT CASNumber, MBMCName, ReasonDesc, CoverID, CLMMCOProcFee, DateVisit, MBMCRELCode, CLMTotalActual, CLMBordxAmt, CASStatus, MBMName, MBMNumber, Claimability
        FROM dbo.VIEW_CISMBMInquiry_GPCASE WHERE MBMNumber = @mbmNumber;
    ELSE
        SELECT CASNumber, MBMCName, ReasonDesc, CoverID, CLMMCOProcFee, DateVisit, MBMCRELCode, CLMTotalActual, CLMBordxAmt, CASStatus, MBMName, MBMNumber, Claimability
        FROM dbo.VIEW_CISMBMInquiry_Spe_CASE WHERE MBMNumber = @mbmNumber;
END
GO

-- showMBMCov / lstCovAdd_Click: the cover plan's AnnualLimit and SpecialistLimit statuses on the policy date, in the desktop's read order
-- (latest first; the desktop keeps the last row it reads).
CREATE OR ALTER PROCEDURE cis.MbmEnq_CoverStatus @plnCode nvarchar(10), @insCode nvarchar(5), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(SuppLimitStatus) AS SuppLimitStatus FROM dbo.AnnualLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND AnnEffDate <= @effDate ORDER BY AnnEffDate DESC;
    SELECT RTRIM(SuppLimitStatus) AS SuppLimitStatus FROM dbo.SpecialistLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND EffDate <= @effDate ORDER BY EffDate DESC;
END
GO

-- showMBMCov, limit status F: PLN + AnnualLimit of the plan (first row).
CREATE OR ALTER PROCEDURE cis.MbmEnq_FLimit @plnCode nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 RTRIM(p.ANNLmtIndSupp) AS AnnLmtIndSupp, a.SuppLimit, a.ChildDepLimit, RTRIM(p.SSpecLimit) AS SSpecLimit, RTRIM(p.VISLmtInd) AS VisLmtInd
    FROM dbo.PLN p INNER JOIN dbo.AnnualLimit a ON a.PLNCode = p.PLNCode WHERE p.PLNCode = @plnCode;
END
GO

-- lstCovAdd_Click: FFS of CIS_FFS2013 (which also needs an HISDB MBMTwo row) for covers other than 00 (@coverMode 2) or one cover (3).
CREATE OR ALTER PROCEDURE cis.MbmEnq_CoverFfs @mbmNumber nvarchar(30), @coverMode int, @covId nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    SELECT cl.CLMMCOProcFee AS FFS
    FROM dbo.CISCas c WITH (NOLOCK)
    INNER JOIN cis.Ws_Claim cl WITH (NOLOCK) ON c.CASNumber = cl.CASNumber
    INNER JOIN cis.Hdb_MBMCLNOthers o ON o.MBMNumber = c.MBMnumber
    INNER JOIN cis.Hdb_MBMTwo t ON t.MBMNumber = c.MBMnumber
    INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    INNER JOIN cis.Ws_ClaimInvoiceActualNew ia WITH (NOLOCK) ON cl.CASNumber = ia.CASNumber
    WHERE c.CASStatus <> 'D' AND c.MBMnumber = @mbmNumber AND ((@coverMode = 2 AND c.CoverID <> '00') OR (@coverMode = 3 AND c.CoverID = @covId));
END
GO

-- MessageDisplay (UB plans with PlncategoryInd): CategoryAmountCal (home office limit) and Dental_Utilization_Limit_Sum.
CREATE OR ALTER PROCEDURE cis.MbmEnq_UbNotes @mbmNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.CategoryAmountCal @mbmnumber = @mbmNumber;
    EXEC dbo.Dental_Utilization_Limit_Sum @mbmnumber = @mbmNumber;
END
GO

-- ===== main/365_mbm_enquiry_flex.sql =====
-- Membership > MBM Enquiry FlexCare / Family Benefits: copies of CISDB spCalcFlexUOB and spGet_UOB_Flex_New_Cal (UOB FlexCare figures).
-- Differences: spGet_UOB_Flex_New_Cal read spCalcFlexUOB through OPENROWSET(SQLNCLI, Server=sqlinstance1) (an ad hoc distributed query to the
-- production server, disabled here); the copies share a temp table instead. spCalcFlexUOB's IC / employee filter is a parameter, not SQL text.

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER PROCEDURE cis.MbmEnq_CalcFlex -- '951009135440',null
	@icno nvarchar(20),
	@empno nvarchar(20)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	declare @icno_l nvarchar(20) = @icno
	declare @empno_l nvarchar(20) = @empno

	Declare @queryMBM AS NVARCHAR(MAX)
Declare @queryMBMcondition as nvarchar(max) =''

set @queryMBM=
	'
select 
row_number() over(order by mbmclnothers.MBMPolEffDate) as MBMCountNo,
mbm.MBMNumber,
cast(MBMCLNOthers.MBMPolEffDate as date) MBMPolEffDate,
cast(MBMCLNOthers.MBMPolExpDate as date) MBMPolExpDate,
cast(cast(ceiling(floor(cast(datediff(d,MBMCLNOthers.MBMPolEffDate,MBMCLNOthers.MBMPolExpDate) as decimal(18,2))/IIF(YEAR(MBMCLNOthers.MBMPolEffDate)%4=0,365,364) * tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as decimal(18,2)) as nvarchar) as AnnualLimit,
--sum(coalesce(claim.CLMBordxAmt,0)) FlexUtil,
coalesce(fConly.FlexUtilFinal,0) as FlexUtil,
coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0) * coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0) as LeavesSoldValue,
coalesce(tblUOBFlexCFOthersConfig.Reward,0) as Reward,

(cast(cast(ceiling(floor(cast(datediff(d,MBMCLNOthers.MBMPolEffDate,MBMCLNOthers.MBMPolExpDate) as decimal(18,2))/IIF(YEAR(MBMCLNOthers.MBMPolEffDate)%4=0,365,364) * tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as decimal(18,2)) as nvarchar)
+(coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0) * coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0))
+coalesce(tblUOBFlexCFOthersConfig.Reward,0)) as OverallLimit,

(cast(cast(ceiling(floor(cast(datediff(d,MBMCLNOthers.MBMPolEffDate,MBMCLNOthers.MBMPolExpDate) as decimal(18,2))/IIF(YEAR(MBMCLNOthers.MBMPolEffDate)%4=0,365,364) * tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as decimal(18,2)) as nvarchar)
+(coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0) * coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0))
+coalesce(tblUOBFlexCFOthersConfig.Reward,0))-coalesce(fConly.FlexUtilFinal,0) as OverallBal

From 
hisdb..mbm
inner join hisdb..mbmclnothers on mbmclnothers.mbmnumber=mbm.mbmnumber
inner join hisdb..mbmothers on mbmothers.mbmnumber=mbm.mbmnumber
inner join cisdb..tblFlexibleBenefitPlanLimit_OP on tblFlexibleBenefitPlanLimit_OP.CISPlnCode=MBMCLNOthers.PLNType and mbmclnothers.MBMPolEffDate between tblFlexibleBenefitPlanLimit_OP.FBEffDate and tblFlexibleBenefitPlanLimit_OP.FBExpDate
left join cisdb..ciscas on ciscas.MBMnumber=mbmclnothers.MBMNumber and ciscas.CoverID=''00'' and CASStatus =''O''
left join cisdb..tblFlexibleBenefitPlan_OP on tblFlexibleBenefitPlan_OP.FBCode=ciscas.CASTreatReason1 
left join CISWorksheet..claim on claim.CASNumber=ciscas.CASNumber
left join cisdb..tblUOBFlexCFOthersConfig on tblUOBFlexCFOthersConfig.MemberNo=mbm.MBMNumber and tblUOBFlexCFOthersConfig.CoverId=''00''
left join (select ciscas.MBMnumber,count(*) cnt,sum(coalesce(CLMBordxAmt,0)) as FlexUtilFinal from cisdb..ciscas inner join cisdb..tblFlexibleBenefitPlan_OP FBP_OP on FBP_OP.FBCode=ciscas.CASTreatReason1 inner join CISWorksheet..claim on claim.CASNumber=ciscas.CASNumber where ciscas.CoverID=''00'' and CASStatus =''O'' group by ciscas.MBMnumber) FConly on FConly.MBMnumber=mbm.MBMNumber
where  
(
	0=0 
) '

+ iif(@icno_l is not null,' and mbm.MBMIcBcPp=@icno','')
+ iif(@empno_l is not null,' and mbmothers.MBMEmployeeNo=@empno','')
+
' 
group by 
fconly.FlexUtilFinal,
mbm.MBMNumber,
MBMCLNOthers.MBMPolEffDate,
MBMCLNOthers.MBMPolExpDate,
tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit,
tblUOBFlexCFOthersConfig.LeavesCF,tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,tblUOBFlexCFOthersConfig.LeavesSold,
tblUOBFlexCFOthersConfig.Reward
order by MBMCLNOthers.MBMPolEffDate
'

--print @queryMBM
declare @F1 as table(
MBMCountNo int,
MBMNumber nvarchar(20),
MBMPolEffDate date,
MBMPolExpDate date,
AnnualLimit money,
FlexUtil money,
LeavesSoldValue money,
Reward money,
OverallLimit money,
OverallBal money
)

insert into @F1
exec sp_executesql @queryMBM, N'@icno nvarchar(20), @empno nvarchar(20)', @icno = @icno_l, @empno = @empno_l

;with cte as
(
select 
F1.*,
case when F2.OverallBal>0 then F2.OverallBal/2 else 0 end as CFfromLastYear
from @F1 F1
left join @F1 F2 on F2.MBMCountNo=F1.MBMCountNo-1
)
insert into #flex_calc
select 
cte.*,
((case when cteFinal.OverallBal >0 then coalesce(cteFinal.OverallBal,0) else 0 end) +coalesce(cteFinal.CFfromLastYear,0))/2 as FinalCF,
coalesce(cte.OverallBal,0)+((coalesce(cteFinal.OverallBal,0)+coalesce(cteFinal.CFfromLastYear,0))/2) as YearBalance
from cte
left join cte cteFinal on cteFinal.MBMCountNo=cte.MBMCountNo-1

	
END
GO

CREATE OR ALTER PROCEDURE cis.MbmEnq_Flex -- null,null,'UBHA0374880*05','2021'
	@icno nvarchar(20)=null,
	@empno nvarchar(20)=null,
	@mbmno nvarchar(20)=null,
	@polYear int=null
	--@res int output
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	if (@mbmno is not null) 
		Begin
			select @empno = PrincipleEmpNoActual from hisdb..vuMemberFamily_OP_short where vuMemberFamily_OP_short.MemberNo = @mbmno
			Set @mbmno='' 
		End	

	Declare @tmpFlexAccu table
	(
		YearId int,
		MBMCountNo int,
		MemberNo nvarchar(20),
		CoverId nvarchar(2),
		ICNo nvarchar(20),
		PrincipalICNo nvarchar(20),
		EffectiveDate date,
		ExpiryDate date,
		MbmStatus char(1),
		mbmemployeeno nvarchar(20),
		PLNType nvarchar(20),
		AccuAmt_CLMBordxAmt money,
		FCEffDate date,
		ALEffDate date
	);

	Declare @tmpFlexLimit table
	(
		YearId int,
		MBMCountNo int,
		MemberNo nvarchar(20),
		CoverId nvarchar(2),
		ICNo nvarchar(20),
		PrincipalICNo nvarchar(20),
		EffectiveDate date,
		ExpiryDate date,
		MbmStatus char(1),
		mbmemployeeno nvarchar(20),
		PLNType nvarchar(20),
		AccuAmt_CLMBordxAmt money,
		Limits nvarchar(20),
		ALEffDate date
	);


	Declare @queryMBM AS NVARCHAR(MAX)
	Declare @queryMBMcondition as nvarchar(max) =''

	set @queryMBM='
		with cte as 
	(
	
		select mbm.mbmnumber,''00'' as CoverId,mbmicbcpp as ICNo,mbmicbcpp PrincipalICNo,mbmclnothers.MBMPolEffDate EffectiveDate,mbmclnothers.MBMPolExpDate ExpiryDate,mbmstatus,
		mbmemployeeno,mbmclnothers.PLNType
		From 
		hisdb..mbm
		inner join hisdb..mbmclnothers on mbmclnothers.mbmnumber=mbm.mbmnumber
		inner join hisdb..mbmothers on mbmothers.mbmnumber=mbm.mbmnumber

		union 

		select 
		mbmcoveredpersons.mbmcnumber,mbmcoveredpersons.mbmccoverid as CoverId,mbmcoveredpersons.MBMCICBCPPNo as ICNo,mbmicbcpp PrincipalICNo ,MBMCCLNEffDate,MBMCCLNExpDate,mbmcstatus,
		mbmemployeeno,mbmclnothers.PLNType
		From hisdb..mbmcoveredpersons
		inner join hisdb..mbm on mbm.mbmnumber=mbmcoveredpersons.mbmcnumber
		inner join hisdb..mbmcoveredpersonscln on mbmcoveredpersonscln.mbmcnumber=mbmcoveredpersons.mbmcnumber and mbmcoveredpersonscln.mbmccoverid=mbmcoveredpersons.mbmccoverid
		inner join hisdb..mbmothers on mbmothers.mbmnumber=mbm.mbmnumber
		inner join hisdb..mbmclnothers on mbmclnothers.mbmnumber=mbm.mbmnumber
	)
	select dense_rank() over (order by cte.mbmnumber) as YearId, row_number() over(partition by cte.mbmnumber order by cte.mbmnumber,cte.CoverId)-1 as MBMCountNo, cte.*,
	--case 
	--			when ((row_number() over(partition by cte.mbmnumber order by cte.mbmnumber,cte.CoverId)-1)=0 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(ceiling(floor(cast((datediff(d,cte.EffectiveDate,cte.ExpiryDate)+1) as numeric)/365*tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as nvarchar) 
	--			when ((row_number() over(partition by cte.mbmnumber order by cte.mbmnumber,cte.CoverId)-1)=1 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(tblFlexibleBenefitPlanLimit_OP.FBFamilyLimit as nvarchar)
	--			when ((row_number() over(partition by cte.mbmnumber order by cte.mbmnumber,cte.CoverId)-1)>1 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then ''NA''
	--			end as Limits,
	sum(coalesce(claim.CLMBordxAmt,0)) as AccuAmt_CLMBordxAmt,tblFlexibleBenefitPlanLimit_OP.ClaimSubmissionLastDate as FCEffDate,tblFlexibleBenefitPlanLimit_OP.AnnualLeaveSellingLastDate as ALEffDate
		From cte
	inner join tblFlexibleBenefitPlanLimit_OP on tblFlexibleBenefitPlanLimit_OP.CISPlnCode=cte.PLNType and (cte.EffectiveDate between tblFlexibleBenefitPlanLimit_OP.FBEffDate and tblFlexibleBenefitPlanLimit_OP.FBExpDate)
	left join cisdb..ciscas on (ciscas.MBMnumber=cte.mbmnumber and ciscas.CoverID=cte.CoverID) and coalesce(ciscas.CASStatus,''O'')=''O''
	left join CISWorksheet..claim on claim.casnumber=ciscas.casnumber
	where 
	--tblFlexibleBenefitPlanLimit_OP.CISPlnCode in (''U206'',''U129'')
	--and mbmstatus=''a''
	--and 
	--year(ExpiryDate) in (year(getdate()),year(getdate())-1)
	year(ExpiryDate) in ('+cast(@polYear as nvarchar(4))+','+cast((@polYear-1) as nvarchar(4))+')'
	

	if @icno is not null
		set @queryMBMcondition=' and cte.PrincipalICNo= '''+@icno +''''
	else if @empno is not null
		set @queryMBMcondition=' and cte.mbmemployeeno= '''+@empno+''''

	set @queryMBM = @queryMBM+@queryMBMcondition+' group by cte.mbmnumber,cte.CoverID,ICNo,cte.PrincipalICNo,cte.mbmemployeeno,PLNType,EffectiveDate,ExpiryDate,cte.mbmstatus,CISPlnCode,FBEmployeeLimit,FBFamilyLimit,tblFlexibleBenefitPlanLimit_OP.ClaimSubmissionLastDate,tblFlexibleBenefitPlanLimit_OP.AnnualLeaveSellingLastDate
				order by year(cte.EffectiveDate),cte.mbmnumber,CoverID'

	print @queryMBM
	insert into @tmpFlexAccu
	exec sp_executesql @queryMBM

--	delete from @tmpFlexAccu
--	where MemberNo in
--	(
--'UBHA0030096*03',
--'UBHA0030096*04',
--'UBHA0091013*03',
--'UBHA0091013*04',
--'UBHA0212743*03',
--'UBHA0212743*04',
--'UBHA0213854*03',
--'UBHA0213854*04',
--'UBHA0299217*03',
--'UBHA0299217*04',
--'UBHA0301338*03',
--'UBHA0301338*04',
--'UBHA1112690*02',
--'UBHA1112690*03',
--'UBHA1132149*02',
--'UBHA1132149*03',
--'UBHA1148747*02',
--'UBHA1148747*03',
--'UBHA1168053*02',
--'UBHA1168053*03',
--'UBHA1178778*02',
--'UBHA1178778*03'
--)

	declare @FCEffDate as date = null
	if (select max(YearId) from @tmpFlexAccu where MBMCountNo=0 )<>1
	begin
		select @FCEffDate=FCEffDate from @tmpFlexAccu where MBMCountNo=0 and YearId=1
	end

insert into @tmpFlexLimit
select 
tFAcc.YearId,MBMCountNo,MemberNo,CoverId,ICNo,PrincipalICNo,EffectiveDate,ExpiryDate,MbmStatus,mbmemployeeno,PLNType,AccuAmt_CLMBordxAmt,
case
	--when (@polYear=2019 and MBMCountNo=0 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(ceiling(floor(cast((datediff(d,tFAcc.EffectiveDate,tFAcc.ExpiryDate)+1) as numeric)/365*tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as nvarchar) 
	--when (MBMCountNo=0 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(cast(ceiling(floor(cast((datediff(d,tFAcc.EffectiveDate,tFAcc.ExpiryDate)+ case when year(tFAcc.EffectiveDate)=2019 or (FORMAT (tFAcc.EffectiveDate, 'dd-MM')<>'01-01') then 1 else 0 end) as decimal(18,2))/365*tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as decimal(18,2)) as nvarchar) 
	when (MBMCountNo=0 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(cast(ceiling(floor(cast(datediff(d,tFAcc.EffectiveDate,tFAcc.ExpiryDate) as decimal(18,2))/IIF(YEAR(tFAcc.EffectiveDate)%4=0,365,364) * tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as decimal(18,2)) as nvarchar)
	when (MBMCountNo=0 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit as nvarchar) 
	when (MBMCountNo=1 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(tblFlexibleBenefitPlanLimit_OP.FBFamilyLimit as nvarchar)
	when (MBMCountNo>1 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then 'NA'
	end as Limits,tFAcc.ALEffDate
from @tmpFlexAccu tFAcc
inner join tblFlexibleBenefitPlanLimit_OP on tblFlexibleBenefitPlanLimit_OP.CISPlnCode=tFAcc.PLNType and (tFAcc.EffectiveDate between tblFlexibleBenefitPlanLimit_OP.FBEffDate and tblFlexibleBenefitPlanLimit_OP.FBExpDate)

	

--select * From @tmpFlexLimit

if (select count(*) from @tmpFlexLimit)>0
begin

select YearId,MBMCountNo,MemberNo,CoverId,ICNo,PLNType,EffectiveDate,ExpiryDate,Limits,AccuAmt_CLMBordxAmt,ALEffDate into #tmpFlexLimit from @tmpFlexLimit 

DECLARE 
	@cols AS NVARCHAR(MAX),
	@colsSum AS NVARCHAR(MAX),
	@query  AS NVARCHAR(MAX)

	select @cols = STUFF((SELECT ',' + QUOTENAME(FBCode) 
						from cisdb..tblFlexibleBenefitPlan_OP
						--where tblFlexibleBenefitPlan_OP.FBCode not in ('301','302')
						group by FBCode
						order by FBCode
				FOR XML PATH(''), TYPE
				).value('.', 'NVARCHAR(MAX)') 
			,1,1,'')


	set @colsSum=replace(@cols,',','+')
	set @colsSum=replace(@colsSum,'[','coalesce([')
	set @colsSum=replace(@colsSum,']','],0)')

	set @query = N'SELECT 
	dense_rank() over (order by MemberNo) as FamilyRowId, 
	mbmno,
	MBMCountNo,MemberNo,CoverID,ICNo, PLNType[PLNType],EffectiveDate,ExpiryDate,Limits,AccuAmt_CLMBordxAmt,
	coalesce([135],0) [135],[135CapLimit],case when [135CapLimit] is not null then [135CapLimit] - coalesce([135],0) end [135Bal],
	coalesce([266],0) [266],[266CapLimit],case when [266CapLimit] is not null then [266CapLimit] - coalesce([266],0) end [266Bal],
	coalesce([300],0)+coalesce([301],0)+coalesce([302],0)+coalesce([211],0) [300],[300CapLimit],case when [300CapLimit] is not null then [300CapLimit] - coalesce([300],0) end [300Bal],
	coalesce([303],0) [303],[303CapLimit],case when [303CapLimit] is not null then [303CapLimit] - coalesce([303],0) end [303Bal],
	coalesce([305],0) [305],[305CapLimit],case when [305CapLimit] is not null then [305CapLimit] - coalesce([305],0) end [305Bal],
	coalesce([306],0) [306],[306CapLimit],case when [306CapLimit] is not null then [306CapLimit] - coalesce([306],0) end [306Bal],
	(coalesce([466],0)+coalesce([662],0)) [466],[466CapLimit],case when [466CapLimit] is not null then [466CapLimit] - (coalesce([466],0) + coalesce([662],0)) end [466Bal],
	coalesce([468],0) [468],[468CapLimit],case when [468CapLimit] is not null then [468CapLimit] - coalesce([468],0) end [468Bal],
	coalesce([614],0) [614],[614CapLimit],case when [614CapLimit] is not null then [614CapLimit] - coalesce([614],0) end [614Bal],
	coalesce([615],0) [615],[615CapLimit],case when [615CapLimit] is not null then [615CapLimit] - coalesce([615],0) end [615Bal],
	coalesce([616],0) [616],[616CapLimit],case when [616CapLimit] is not null then [616CapLimit] - coalesce([616],0) end [616Bal],
	coalesce([617],0) [617],[617CapLimit],case when [617CapLimit] is not null then [617CapLimit] - coalesce([617],0) end [617Bal],
	coalesce([618],0) [618],[618CapLimit],case when [618CapLimit] is not null then [618CapLimit] - coalesce([618],0) end [618Bal],
	coalesce([619],0) [619],[619CapLimit],case when [619CapLimit] is not null then [619CapLimit] - coalesce([619],0) end [619Bal],
	coalesce([654],0) [654],[654CapLimit],case when [654CapLimit] is not null then [654CapLimit] - coalesce([654],0) end [654Bal],
	
	coalesce([696],0) [696],[696CapLimit],case when [696CapLimit] is not null then [696CapLimit] - coalesce([696],0) end [696Bal],
	coalesce([697],0) [697],[697CapLimit],case when [697CapLimit] is not null then [697CapLimit] - coalesce([697],0) end [697Bal],
	(coalesce([703],0)+coalesce([698],0)+coalesce([699],0)+coalesce([700],0)+coalesce([701],0)+coalesce([702],0)) [703],[703CapLimit],case when [703CapLimit] is not null then [703CapLimit] - (coalesce([703],0) + coalesce([703],0)+coalesce([698],0)+coalesce([699],0)+coalesce([700],0)+coalesce([701],0)+coalesce([702],0)) end [703Bal],
	'  
	 
	+ @colsSum 
	+' as SumFlexAmount, case when CoverID=''00'' then coalesce(AccuAmt_ClmBordxAmt,0)-(' +@colsSum+ ') else coalesce(AccuAmt_ClmBordxAmt,0) end as SumOP ' 
	+N'from 
				 (
					select 
					CASE WHEN CHARINDEX(''*'', #tmpFlexLimit.MemberNo) > 0 THEN LEFT(#tmpFlexLimit.MemberNo, CHARINDEX(''*'', #tmpFlexLimit.MemberNo) - 1) ELSE #tmpFlexLimit.MemberNo END AS MbmNo,
					MBMCountNo,#tmpFlexLimit.MemberNo,#tmpFlexLimit.CoverID,#tmpFlexLimit.ICNo,#tmpFlexLimit.PLNType,#tmpFlexLimit.EffectiveDate,#tmpFlexLimit.ExpiryDate,#tmpFlexLimit.Limits,
					#tmpFlexLimit.AccuAmt_CLMBordxAmt,
					--count(*) as CaseCount, 
					sum(coalesce(claim.CLMBordxAmt,0)) as AccuFlexSum,
					tblFlexibleBenefitPlan_OP.FBCode,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''135'',case when mbmcountno=0 then 1 else 0 end) end as [135CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''266'',case when mbmcountno=0 then 1 else 0 end) end as [266CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''300'',case when mbmcountno=0 then 1 else 0 end) end as [300CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''303'',case when mbmcountno=0 then 1 else 0 end) end as [303CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''305'',case when mbmcountno=0 then 1 else 0 end) end as [305CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''306'',case when mbmcountno=0 then 1 else 0 end) end as [306CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''466'',case when mbmcountno=0 then 1 else 0 end) end as [466CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''468'',case when mbmcountno=0 then 1 else 0 end) end as [468CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''614'',case when mbmcountno=0 then 1 else 0 end) end as [614CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''615'',case when mbmcountno=0 then 1 else 0 end) end as [615CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''616'',case when mbmcountno=0 then 1 else 0 end) end as [616CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''617'',case when mbmcountno=0 then 1 else 0 end) end as [617CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''618'',case when mbmcountno=0 then 1 else 0 end) end as [618CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''619'',case when mbmcountno=0 then 1 else 0 end) end as [619CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''654'',case when mbmcountno=0 then 1 else 0 end) end as [654CapLimit] ,

					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''696'',case when mbmcountno=0 then 1 else 0 end) end as [696CapLimit] , 
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''697'',case when mbmcountno=0 then 1 else 0 end) end as [697CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''703'',case when mbmcountno=0 then 1 else 0 end) end as [703CapLimit] 
					
					From #tmpFlexLimit
					left join cisdb..ciscas on ciscas.MBMnumber=#tmpFlexLimit.MemberNo and ciscas.CoverID=#tmpFlexLimit.CoverID and coalesce(ciscas.CASStatus,''O'')=''O''
					left join CISWorksheet..claim on claim.casnumber=ciscas.casnumber
					left join cisdb..TreatmentReason TR1 on tr1.ReasonCode=ciscas.CASTreatReason1
					left join cisdb..tblFlexibleBenefitPlan_OP on tblFlexibleBenefitPlan_OP.FBCode=tr1.ReasonCode and ((tblFlexibleBenefitPlan_OP.FBEmployee=1 and #tmpFlexLimit.CoverID=''00'') or (tblFlexibleBenefitPlan_OP.FBFamily=1 and #tmpFlexLimit.CoverID<>''00''))
					--where and coalesce(ciscas.CASStatus,''O'')=''O''
					group by 
					MBMCountNo,
					#tmpFlexLimit.MemberNo,
					#tmpFlexLimit.CoverID,
					#tmpFlexLimit.ICNo,
					#tmpFlexLimit.PLNType,
					#tmpFlexLimit.EffectiveDate,
					#tmpFlexLimit.ExpiryDate,
					#tmpFlexLimit.Limits,
					#tmpFlexLimit.AccuAmt_CLMBordxAmt,
					tblFlexibleBenefitPlan_OP.FBCode
				) x
				pivot 
				(
					sum(AccuFlexSum)
					for FBCode in (' + @cols + N')
				) p '


		--PRINT CAST(@query AS NTEXT)
		create table #tmpUtil (
		FamilyRowId int,
		mbmno nvarchar(20),
		MBMCountNo tinyint,
		MemberNo nvarchar(20),
		CoverID nvarchar(2),
		ICNO nvarchar(100),
		PLNType nvarchar(4),
		EffectiveDate Date,
		ExpiryDate Date,
		Limits nvarchar(20),
		AccuAmt_CLMBordxAmt money,
		[135] money,
		[135Cap] money,
		[135Bal] money,
		[266] money,
		[266Cap] money,
		[266Bal] money,
		[300] money,
		[300Cap] money,
		[300Bal] money,
		[303] money,
		[303Cap] money,
		[303Bal] money,
		[305] money,
		[305Cap] money,
		[305Bal] money,
		[306] money,
		[306Cap] money,
		[306Bal] money,
		[466] money,
		[466Cap] money,
		[466Bal] money,
		[468] money,
		[468Cap] money,
		[468Bal] money,
		[614] money,
		[614Cap] money,
		[614Bal] money,
		[615] money,
		[615Cap] money,
		[615Bal] money,
		[616] money,
		[616Cap] money,
		[616Bal] money,
		[617] money,
		[617Cap] money,
		[617Bal] money,
		[618] money,
		[618Cap] money,
		[618Bal] money,
		[619] money,
		[619Cap] money,
		[619Bal] money,
		[654] money,
		[654Cap] money,
		[654Bal] money,

		[696] money,
		[696Cap] money,
		[696Bal] money,
		[697] money,
		[697Cap] money,
		[697Bal] money,
		[703] money,
		[703Cap] money,
		[703Bal] money,
		SumFlexAmount money,
		SumOP money
		);

	insert into #tmpUtil
	exec sp_executesql @query;

	ALTER TABLE #tmpUtil ADD DepUtil money,	DepBal money, PriFlexBal money

	declare @DepUtil money, @DepLimits money
	declare @FamilyCount int = 0
	declare @FamilyMemberCount int = 0
	declare @CurrentFamilyRowId int = 1
	declare @CurrentFamilyMemberRowId int = 1

	select @FamilyCount=max(FamilyRowId) from #tmpUtil

	while @CurrentFamilyRowId<=@FamilyCount
	begin
		select @DepUtil=sum(SumOP) From #tmpUtil where mbmcountno>0 and FamilyRowId=@CurrentFamilyRowId
		select @DepLimits=Limits from #tmpUtil where MBMCountNo=1 and FamilyRowId=@CurrentFamilyRowId
	
		update #tmpUtil set DepUtil=@DepUtil,DepBal=@DepLimits-@DepUtil where MBMCountNo=1 and FamilyRowId=@CurrentFamilyRowId
		update #tmpUtil set PriFlexBal=Limits-SumFlexAmount where MBMCountNo=0 and FamilyRowId=@CurrentFamilyRowId

		set @CurrentFamilyRowId=@CurrentFamilyRowId+1
	end

	
	-- The desktop procedure loads spCalcFlexUOB through OPENROWSET(SQLNCLI, sqlinstance1); here a temp table shared with the callee.
	create table #flex_calc (MBMCountNo int, MBMNumber nvarchar(20), MBMPolEffDate date, MBMPolExpDate date, AnnualLimit money, FlexUtil money,
		LeavesSoldValue money, Reward money, OverallLimit money, OverallBal money, CFfromLastYear money, FinalCF money, YearBalance money)
	if @icno is not null exec cis.MbmEnq_CalcFlex @icno, null
	else if @empno is not null exec cis.MbmEnq_CalcFlex null, @empno

	declare @tblFlexCalc as table(
		MBMCountNo int,
		MBMNumber nvarchar(20),
		MBMPolEffDate date,
		MBMPolExpDate date,
		AnnualLimit money,
		FlexUtil money,
		LeavesSoldValue money,
		Reward money,
		OverallLimit money,
		OverallBal money,
		CFfromLastYear money,
		FinalCF money,
		YearBalance money

	)

	insert into @tblFlexCalc select * from #flex_calc


	select 
	dense_rank() over (order by #tmpUtil.mbmno) as MBMId,
	#tmpUtil.* ,
	--case when #tmpUtil.MBMCountNo='0' then iif(coalesce(#tmpUtil.Priflexbal,0)>0,coalesce(#tmpUtil.Priflexbal,0)/2,0) end as PriFlexBalCF,
	--case when #tmpUtil.MBMCountNo='0' then iif(coalesce(#tmpUtil.Priflexbal,0)>0,coalesce(#tmpUtil.Priflexbal,0)/2,0) + iif(((coalesce(tblUOBFlexCFOthersConfig.leavessold,0)*coalesce(leaveexchangevalueperday,0))+coalesce(Reward,0))>0, ((coalesce(tblUOBFlexCFOthersConfig.leavessold,0)*coalesce(leaveexchangevalueperday,0))+coalesce(Reward,0))/2,0) end as PriFlexBalCF,
	FlexCareCFcalc.FinalCF as PriFlexBalCF,
	case when #tmpUtil.CoverID<>'00' then (coalesce(#tmpUtil.SumOP,0)-coalesce(SumFlexAmount,0)) else null end as IPOPSPUtil,
	case when #tmpUtil.MBMCountNo='0' then coalesce(#tmpUtil.Priflexbal,0) when #tmpUtil.MBMCountNo='1' then coalesce(#tmpUtil.DepBal,0) else null end as FinalBalAll,
	vuMemberFamily_OP_short.MemberName,
	vuMemberFamily_OP_short.PrincipleEmpNoActual,
	ALEffDate
	into #tmpFlexCF
	From #tmpUtil 
	inner join #tmpFlexLimit on #tmpFlexLimit.MemberNo=#tmpUtil.MemberNo and #tmpFlexLimit.CoverID=#tmpUtil.CoverID 
	inner join hisdb..vuMemberFamily_OP_short on vuMemberFamily_OP_short.MemberNo=#tmpUtil.MemberNo and vuMemberFamily_OP_short.CoverID=#tmpUtil.CoverID 
	left join tblUOBFlexCFOthersConfig on tblUOBFlexCFOthersConfig.MemberNo=#tmpUtil.MemberNo and tblUOBFlexCFOthersConfig.CoverId=#tmpUtil.CoverId
	left join @tblFlexCalc FlexCareCFcalc on FlexCareCFcalc.MBMNumber=vuMemberFamily_OP_short.MemberNo and vuMemberFamily_OP_short.CoverID='00'
	order by FamilyRowId,#tmpUtil.MBMCountNo
	
	--select * from #tmpFlexCF

	declare @cnt int
	select @cnt=max(mbmid) from #tmpFlexCF
	
	--declare @currMBMId int =1
	--while @currMBMId<=@cnt
	--begin
	--		if (select count(mbmcountno) from #tmpFlexCF where mbmid=@currMBMId and mbmcountno=0)=2
	--		begin
	--		declare @CFAmt as nvarchar(20) =(select coalesce(PriFlexBalCF,0) from #tmpFlexCF where mbmid=@currMBMId and mbmcountno=0 and year(effectivedate)= (@polYear-1)) --year(getdate())-1)
	--			 update #tmpFlexCF set PriFlexBalCF=@CFAmt where mbmid=@currMBMId and mbmcountno=0 and year(effectivedate)=@polYear --year(getdate())
	--			 update #tmpFlexCF set PriFlexBalCF=null where mbmid=@currMBMId and mbmcountno=0 and year(effectivedate)=(@polYear-1) --year(getdate())-1
	--		end
	--		else
	--			begin
	--			 update #tmpFlexCF set PriFlexBalCF=0.00 where mbmid=@currMBMId and mbmcountno=0 
	--			end
	--	set @currMBMId=@currMBMId+1
	--end

	select MBMId,	FamilyRowId,	mbmno,	MBMCountNo,	#tmpFlexCF.MemberNo,	#tmpFlexCF.CoverID,	ICNO,	PLNType,	EffectiveDate,	ExpiryDate,	
	
	Limits,
	--iif(coalesce(Limits,0)>0,cast(Limits as decimal(18,2)),Limits) as Limits,
	
	AccuAmt_CLMBordxAmt,	
	#tmpFlexCF.[135],	#tmpFlexCF.[135Cap],	#tmpFlexCF.[135Bal],	
	#tmpFlexCF.[266],	#tmpFlexCF.[266Cap],	#tmpFlexCF.[266Bal],	
	#tmpFlexCF.[300],	#tmpFlexCF.[300Cap],	#tmpFlexCF.[300Bal],	
	#tmpFlexCF.[303],	#tmpFlexCF.[303Cap],	#tmpFlexCF.[303Bal],	
	#tmpFlexCF.[305],	#tmpFlexCF.[305Cap],	#tmpFlexCF.[305Bal],	
	#tmpFlexCF.[306],	#tmpFlexCF.[306Cap],	#tmpFlexCF.[306Bal],	
	#tmpFlexCF.[466],	#tmpFlexCF.[466Cap],	#tmpFlexCF.[466Bal],	
	#tmpFlexCF.[468],	#tmpFlexCF.[468Cap],	#tmpFlexCF.[468Bal],	
	#tmpFlexCF.[614],	#tmpFlexCF.[614Cap],	#tmpFlexCF.[614Bal],	
	#tmpFlexCF.[615],	#tmpFlexCF.[615Cap],	#tmpFlexCF.[615Bal],	
	#tmpFlexCF.[616],	#tmpFlexCF.[616Cap],	#tmpFlexCF.[616Bal],	
	#tmpFlexCF.[617],	#tmpFlexCF.[617Cap],	#tmpFlexCF.[617Bal],	
	#tmpFlexCF.[618],	#tmpFlexCF.[618Cap],	#tmpFlexCF.[618Bal],	
	#tmpFlexCF.[619],	#tmpFlexCF.[619Cap],	#tmpFlexCF.[619Bal],	
	#tmpFlexCF.[654],	#tmpFlexCF.[654Cap],	#tmpFlexCF.[654Bal],

	#tmpFlexCF.[696],	#tmpFlexCF.[696Cap],	#tmpFlexCF.[696Bal],
	#tmpFlexCF.[697],	#tmpFlexCF.[697Cap],	#tmpFlexCF.[697Bal],
	#tmpFlexCF.[703],	#tmpFlexCF.[703Cap],	#tmpFlexCF.[703Bal],
	
	SumFlexAmount,	SumOP,	#tmpFlexCF.DepUtil,	#tmpFlexCF.DepBal,	#tmpFlexCF.PriFlexBal,	PriFlexBalCF,	IPOPSPUtil,	FinalBalAll,	MemberName,	PrincipleEmpNoActual,	
	LeavesCF,	cast(LeavesSold as int) LeavesSold,	tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,Reward,
	case when #tmpFlexCF.coverid='00' then convert(decimal(18,2),coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0)*coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0)) else null end as LeavesSoldAmount,
	
	case when #tmpFlexCF.coverid='00' then 
	convert
	(
		decimal(18,2),coalesce(Limits,'0')+
		(
			iif(@FCEffDate<cast(getdate() as date),coalesce(PriFlexBalCF,0),0)
			+ iif(cast(getdate() as date)> #tmpFlexCF.ALEffDate, (coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0)*coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0)),0)
			+ iif(cast(getdate() as date)> #tmpFlexCF.ALEffDate, coalesce(tblUOBFlexCFOthersConfig.Reward,0),0)
		 )
	 )-convert(decimal(18,2),coalesce(SumFlexAmount,0))
	else null
	end as PriFlexCurrAvailBal,

	case when #tmpFlexCF.coverid='00' then 
	convert
	(
		decimal(18,2),coalesce(Limits,'0')+
		(
			iif(@FCEffDate<cast(getdate() as date),coalesce(PriFlexBalCF,0),0)
			+ iif(cast(getdate() as date)> #tmpFlexCF.ALEffDate, (coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0)*coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0)),0)
			+ iif(cast(getdate() as date)> #tmpFlexCF.ALEffDate, coalesce(tblUOBFlexCFOthersConfig.Reward,0),0)
		 )
	 )
	else null
	end as PriFlexAvailLimit,

	coalesce(@FCEffDate,'') as FCEffDate,
	coalesce(ALEffDate,'') as ALEffDate
	From #tmpFlexCF 
	left join tblUOBFlexCFOthersConfig on tblUOBFlexCFOthersConfig.MemberNo=#tmpFlexCF.MemberNo and tblUOBFlexCFOthersConfig.CoverId=#tmpFlexCF.CoverId
	where year(effectivedate)=@polYear --year(getdate())
	--order by MBMId,FamilyRowId,MBMCountNo
	order by MBMId,MBMCountNo,ExpiryDate Desc



	drop table #tmpFlexLimit
	drop table #tmpUtil
	drop table #tmpFlexCF
	--set @res=1
end

else
	begin
	print 'no data'
	--set @res=0
	end
END
GO

-- ===== main/400_batch_transfer.sql =====
-- Invoice > Clinic > Batch Transfer (FrmInvoiceBatchNoUpd): move the cases of one OP batch to another (CISDB CISCAS).

-- ShowInfo: the batch's cases per invoice / clinic (the desktop's SP_SeachCasByOPBatch, already parameterised).
CREATE OR ALTER PROCEDURE cis.BatchTransfer_Batch @batchNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.SP_SeachCasByOPBatch @tmpBatchNo = @batchNo;
END
GO

-- BatchTransProcess: batch 1's cases move to batch 2 with the transfer remark, then batch 2's invoice statement amount becomes the total.
CREATE OR ALTER PROCEDURE cis.BatchTransfer_Save @fromBatch nvarchar(30), @toBatch nvarchar(30), @remark nvarchar(400), @amount money AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET OPBatchNo = @toBatch, BatchTransRemarks = @remark WHERE OPBatchNo = @fromBatch;
    DECLARE @moved int = @@ROWCOUNT;
    UPDATE dbo.CISCAS SET InvoiceStmtAmt = @amount WHERE OPBatchNo = @toBatch;
    SELECT @moved;
END
GO

-- ===== main/410_invoice_update.sql =====
-- Invoice > Other > Invoice Update (frmUpdateCLNInvoice): list cases by invoice / clinic / payor / dates (CISWorksheet
-- VIEW_Invoice_Tracking_2016, the desktop's ClaimBordxModule case 1) and stamp the checked cases with an invoice number and date (CISCAS).

-- ChkClinic_Click / ChkGrpComp_Click: panel clinics (HISDB CLN, panel Y / C) and group companies (CISDB GRPCompany).
CREATE OR ALTER PROCEDURE cis.InvUpdate_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(CLNCode) AS Code, RTRIM(CLNName) AS Name FROM cis.Hdb_CLN WHERE (CLNPanel = 'Y' OR CLNPanel = 'C') AND LEN(CLNName) > 0;
    SELECT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany;
END
GO

-- SearchRecord → InvoiceListing: every filter optional; @noInvoice = ChkInvNo (BillNo empty).
CREATE OR ALTER PROCEDURE cis.InvUpdate_Search
    @casFrom nvarchar(20) = NULL, @casTo nvarchar(20) = NULL, @clinicCode nvarchar(10) = NULL, @groupCompany nvarchar(200) = NULL,
    @billNo nvarchar(100) = NULL, @visitFrom datetime = NULL, @visitTo datetime = NULL, @noInvoice bit = 0, @billFrom datetime = NULL,
    @billTo datetime = NULL, @amountFrom money = NULL, @amountTo money = NULL, @billDate datetime = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT CASNumber, MBMnumber AS MBMNumber, CoverID, PatientName, DateVisit, CLNName, BillNo, BillDate, INVEnteredDate, INVEnteredBy, INVAmount,
           CLMMCOProcFee, CLMFinalApprove, MBMGroupCompany
    FROM cis.Ws_VIEW_Invoice_Tracking_2016
    WHERE (@casFrom IS NULL OR CASNumber >= @casFrom) AND (@casTo IS NULL OR CASNumber <= @casTo)
      AND (@clinicCode IS NULL OR ClinicCode = @clinicCode) AND (@groupCompany IS NULL OR MBMGroupCompany = @groupCompany)
      AND (@billNo IS NULL OR BillNo = @billNo) AND (@visitFrom IS NULL OR DateVisit >= @visitFrom) AND (@visitTo IS NULL OR DateVisit <= @visitTo)
      AND (@noInvoice = 0 OR BillNo IS NULL OR BillNo = '') AND (@billFrom IS NULL OR BillDate >= @billFrom) AND (@billTo IS NULL OR BillDate <= @billTo)
      AND (@amountFrom IS NULL OR INVAmount >= @amountFrom) AND (@amountTo IS NULL OR INVAmount <= @amountTo)
      AND (@billDate IS NULL OR BillDate = @billDate);
END
GO

-- UpdateRecord: one checked case.
CREATE OR ALTER PROCEDURE cis.InvUpdate_Save @casNumber nvarchar(20), @invoiceNo nvarchar(100), @invoiceDate datetime, @today datetime, @userCode nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET InvoiceNo = @invoiceNo, InvoiceDate = @invoiceDate, RegisterDate = @today, RegisterBy = @userCode WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO

-- ===== main/420_op_batch.sql =====
-- Invoice > Other > OP Batch/Invoice (frmUpdateInvoiceBatch): list open cases (the desktop's ClaimBordxModule case 27, parameterised),
-- give the checked cases an OP batch number (GenerateOPSeq "OC") and / or an invoice number and date (CISCAS).

-- SearchRecord → InvoiceListing. Every filter optional; flags: 1 = the desktop's check box / combo choice.
CREATE OR ALTER PROCEDURE cis.OpBatch_Search
    @noBatch bit = 0, @noBordx bit = 0, @bordxFrom nvarchar(30) = NULL, @bordxTo nvarchar(30) = NULL, @noBtbgBordx bit = 0,
    @btbgFrom nvarchar(30) = NULL, @btbgTo nvarchar(30) = NULL, @casFrom nvarchar(20) = NULL, @casTo nvarchar(20) = NULL,
    @clinicCode nvarchar(10) = NULL, @groupCompany nvarchar(200) = NULL, @invoiceNo nvarchar(100) = NULL, @visitFrom datetime = NULL,
    @visitTo datetime = NULL, @noInvoice bit = 0, @invoiceFrom datetime = NULL, @invoiceTo datetime = NULL, @amountFrom money = NULL,
    @amountTo money = NULL, @batchFrom nvarchar(30) = NULL, @batchTo nvarchar(30) = NULL, @omni bit = 0, @regFrom datetime = NULL,
    @regTo datetime = NULL, @clinicType char(1) = NULL, @ffs char(1) = NULL, @hasBatch char(1) = NULL, @payTo nvarchar(4) = NULL,
    @referral char(1) = NULL, @entity int = 0
AS
BEGIN
    SET NOCOUNT ON;
    SELECT c.CASNumber, c.MBMnumber AS MBMNumber, c.CoverID,
           CASE c.CoverID WHEN '00' THEN m.MBMName ELSE cp.MBMCName END AS PatientName,
           c.DateVisit, o.MBMGroupCompany, ci.INVAmount, cl.CLMMCOProcFee,
           CASE WHEN p.PLNBTBG = 1 THEN COALESCE(cl.CLMFinalApprove, 0) - COALESCE(cl.CLMBTBGAmt, 0) ELSE cl.CLMFinalApprove END AS CLMFinalApprove,
           c.InvoiceNo AS BillNo, o.MBMDepartment, cln.CLNName, c.InvoiceDate AS BillDate,
           CASE WHEN LEN(ci.INVEnteredBy) = 0 THEN CONVERT(varchar(10), r.RecDate, 111) ELSE CONVERT(varchar(10), ci.INVEnteredDate, 111) END AS INVEnteredDate,
           CASE WHEN LEN(ci.INVEnteredBy) = 0 THEN r.Incharge ELSE ci.INVEnteredBy END AS INVEnteredBy,
           c.OPBatchNo,
           CASE WHEN ci.INVPayTo = 'M' THEN pm.PVNo ELSE pc.PVCLNNo END AS PNO,
           CASE WHEN ci.INVPayTo = 'M' THEN pm.PVDate ELSE pc.PVCLNDate END AS PDATE,
           CASE WHEN ci.INVPayTo = 'M' THEN pm.PVChequeNo ELSE pc.PVCLNChequeNo END AS PCNO,
           dbo.WebGetTreatmentDesc(c.CASTreatReason1) AS ReasonDesc1, dbo.WebGetTreatmentDesc(c.CASTreatReason2) AS ReasonDesc2,
           dbo.WebGetTreatmentDesc(c.CASTreatReason3) AS ReasonDesc3, dbo.WebGetTreatmentDesc(c.CASTreatReason4) AS ReasonDesc4,
           dbo.WebGetTreatmentDesc(c.CASTreatReason5) AS ReasonDesc5,
           c.CASDiagnosis5, c.MCDays, cl.CLMCoPayCLNColl,
           (CASE WHEN p.PLNBTBG = 1 THEN COALESCE(cl.CLMFinalApprove, 0) - COALESCE(cl.CLMBTBGAmt, 0) ELSE COALESCE(cl.CLMFinalApprove, 0) END)
               - COALESCE(ci.INVAmount, 0) AS CLMFinalApprove1,
           CASE c.CoverID WHEN '00' THEN m.MBMIcBcPp ELSE cp.MBMCICBCPPNo END AS ICNo,
           CASE WHEN LEN(c.OPBatchNo) = 0 THEN 'NO' ELSE 'YES' END AS OPNO,
           CASE WHEN ci.INVPayTo = 'M' THEN 'Reimbursement' ELSE 'Cash Less' END AS PType,
           c.OPBatchDate, c.LastUpdateDate, COALESCE(c.Pending, 'N') AS Pending, c.Claimability
    FROM dbo.CISCAS c WITH (NOLOCK)
    INNER JOIN cis.Hdb_MBMCLNOthers co WITH (NOLOCK) ON co.MBMNumber = c.MBMnumber
    INNER JOIN cis.Hdb_MBMOthers o WITH (NOLOCK) ON o.MBMNumber = co.MBMNumber
    INNER JOIN cis.Hdb_MBM m WITH (NOLOCK) ON m.MBMNumber = o.MBMNumber
    INNER JOIN cis.Ws_Claim cl WITH (NOLOCK) ON cl.CASNumber = c.CASNumber
    LEFT JOIN cis.Hdb_MBMCoveredPersons cp WITH (NOLOCK) ON cp.MBMCNumber = c.MBMnumber AND cp.MBMCCoverID = c.CoverID
    LEFT JOIN cis.Hdb_CLN cln WITH (NOLOCK) ON c.ClinicCode = cln.CLNCode
    LEFT JOIN dbo.GRPCompany g WITH (NOLOCK) ON g.GRPCoName = o.MBMGroupCompany AND (g.HISpaycode IS NULL OR g.HISpaycode = m.PAYCode)
    LEFT JOIN cis.Ws_ClaimInvoice ci WITH (NOLOCK) ON ci.CASNumber = cl.CASNumber
    LEFT JOIN dbo.tblRecord r WITH (NOLOCK) ON r.CaseNo = cl.CASNumber
    LEFT JOIN cis.Pay_Payment2CLN pc WITH (NOLOCK) ON pc.CASNumber = cl.CASNumber
    LEFT JOIN cis.Pay_Payment2MBM pm WITH (NOLOCK) ON pm.CASNumber = cl.CASNumber
    INNER JOIN dbo.PLN p WITH (NOLOCK) ON p.PLNCode = co.PLNType
    WHERE co.PAYCode NOT IN ('DZ', 'EM', 'ES') AND c.CASStatus = 'O' AND o.MBMGroupCompany <> 'DEMO COMPANY BERHAD'
      AND (@noBatch = 0 OR COALESCE(c.OPBatchNo, '') = '')
      AND (@noBordx = 0 OR COALESCE(cl.BordxRefNo, '') = '')
      AND (@bordxFrom IS NULL OR cl.BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR cl.BordxRefNo <= @bordxTo)
      AND (@noBtbgBordx = 0 OR COALESCE(cl.BTBGBordxRefNo, '') = '')
      AND (@btbgFrom IS NULL OR cl.BTBGBordxRefNo >= @btbgFrom) AND (@btbgTo IS NULL OR cl.BTBGBordxRefNo <= @btbgTo)
      AND (@casFrom IS NULL OR c.CASNumber >= @casFrom) AND (@casTo IS NULL OR c.CASNumber <= @casTo)
      AND (@clinicCode IS NULL OR c.ClinicCode = @clinicCode) AND (@groupCompany IS NULL OR o.MBMGroupCompany = @groupCompany)
      AND (@invoiceNo IS NULL OR c.InvoiceNo = @invoiceNo)
      AND (@visitFrom IS NULL OR c.DateVisit >= @visitFrom) AND (@visitTo IS NULL OR c.DateVisit <= @visitTo)
      AND (@noInvoice = 0 OR COALESCE(c.InvoiceNo, '') = '')
      AND (@invoiceFrom IS NULL OR c.InvoiceDate >= @invoiceFrom) AND (@invoiceTo IS NULL OR c.InvoiceDate <= @invoiceTo)
      AND (@amountFrom IS NULL OR ci.INVAmount >= @amountFrom) AND (@amountTo IS NULL OR ci.INVAmount <= @amountTo)
      AND (@batchFrom IS NULL OR c.OPBatchNo >= @batchFrom) AND (@batchTo IS NULL OR c.OPBatchNo <= @batchTo)
      AND (@omni = 0 OR g.GRPOMNICLM = 'Y')
      AND (@regFrom IS NULL OR c.RegisterDate >= @regFrom) AND (@regTo IS NULL OR c.RegisterDate <= @regTo)
      AND (@clinicType IS NULL OR (@clinicType = '0' AND c.CLNType = 'S') OR (@clinicType = '1' AND c.CLNType = 'G')
           OR (@clinicType = '2' AND COALESCE(c.CLNType, '') = ''))
      AND (@ffs IS NULL OR (@ffs = 'Y' AND cln.CLNFFSStatus = 'Y') OR (@ffs = 'N' AND (cln.CLNFFSStatus = 'N' OR COALESCE(cln.CLNFFSStatus, '') = '')))
      AND (@hasBatch IS NULL OR (@hasBatch = 'Y' AND COALESCE(c.OPBatchNo, '') <> '') OR (@hasBatch = 'N' AND COALESCE(c.OPBatchNo, '') = ''))
      AND (@payTo IS NULL OR ci.INVPayTo = @payTo)
      AND (@referral IS NULL OR (@referral = 'Y' AND cl.CLMRefLetter = 'Y') OR (@referral = 'N' AND (cl.CLMRefLetter = 'N' OR COALESCE(cl.CLMRefLetter, '') = '')))
      AND ((@entity = 0 AND SUBSTRING(c.CASNumber, 1, 2) <> 'SJ') OR (@entity = 1 AND SUBSTRING(c.CASNumber, 1, 2) = 'SJ') OR @entity NOT IN (0, 1));
END
GO

-- The invoice number is already used (CISCAS.InvoiceNo).
CREATE OR ALTER PROCEDURE cis.OpBatch_InvoiceUsed @invoiceNo nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.CISCAS WHERE InvoiceNo = @invoiceNo) THEN 1 ELSE 0 END AS bit);
END
GO

-- GenerateOPSeq("OC"): next OPBatch sequence of the code (letter + 7 digits; the letter moves on after 9999999), locked against other users.
CREATE OR ALTER PROCEDURE cis.OpBatch_Next @opCode nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seq nvarchar(20);
    SELECT @seq = RTRIM(OPBatchSeq) FROM dbo.OPBatch WITH (UPDLOCK, HOLDLOCK) WHERE OPCode = @opCode;
    IF @seq IS NULL
    BEGIN
        SET @seq = 'A0000001';
        INSERT INTO dbo.OPBatch (OPCode, OPBatchSeq) VALUES (@opCode, @seq);
    END
    ELSE
    BEGIN
        DECLARE @letter char(1) = LEFT(@seq, 1), @no int = TRY_CAST(SUBSTRING(@seq, 2, 8) AS int);
        IF @letter BETWEEN 'A' AND 'Y'
            SELECT @seq = CASE WHEN @no = 9999999 THEN CHAR(ASCII(@letter) + 1) + '0000001' ELSE @letter + RIGHT('0000000' + CAST(@no + 1 AS varchar(8)), 7) END;
        ELSE
            SET @seq = 'A' + RIGHT('0000000' + CAST(@no AS varchar(8)), 7);
        UPDATE dbo.OPBatch SET OPBatchSeq = @seq WHERE OPCode = @opCode;
    END
    SELECT @opCode + @seq;
END
GO

-- UpdateRecord: one checked case. @batchNo set = the case had no OP batch; @invoiceNo set = invoice number and date (with the month / year).
CREATE OR ALTER PROCEDURE cis.OpBatch_Save
    @casNumber nvarchar(20), @batchNo nvarchar(30) = NULL, @userCode nvarchar(20), @today date, @invoiceNo nvarchar(100) = NULL, @invoiceDate datetime = NULL
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET
        OPBatchNo = COALESCE(@batchNo, OPBatchNo),
        OPNoEnterUser = CASE WHEN @batchNo IS NULL THEN OPNoEnterUser ELSE @userCode END,
        OPBatchDate = CASE WHEN @batchNo IS NULL THEN OPBatchDate ELSE @today END,
        InvoiceNo = COALESCE(@invoiceNo, InvoiceNo),
        InvoiceDate = CASE WHEN @invoiceNo IS NULL THEN InvoiceDate ELSE @invoiceDate END,
        CASInvMonth = CASE WHEN @invoiceNo IS NULL THEN CASInvMonth ELSE RIGHT('0' + CAST(MONTH(@invoiceDate) AS varchar(2)), 2) END,
        CASInvYear = CASE WHEN @invoiceNo IS NULL THEN CASInvYear ELSE CAST(YEAR(@invoiceDate) AS varchar(4)) END
    WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO

-- ===== main/430_invoice_popups.sql =====
-- Invoice pop-ups shared by the invoice screens: ICD / treatment reason search (frmSearchICD) and case list (frmCasList).

-- frmSearchICD cmdSearch_Click: code starts with, description contains, and a second description joined with AND / OR.
-- The desktop appended "AND/OR ReasonDesc like ..." without brackets, so with OR the other filters only apply to the first part; kept.
CREATE OR ALTER PROCEDURE cis.Popup_Icd @code nvarchar(20) = NULL, @desc nvarchar(200) = NULL, @desc2 nvarchar(200) = NULL, @useOr bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(ReasonCode) AS ReasonCode, RTRIM(ReasonDesc) AS ReasonDesc
    FROM dbo.TreatmentReason
    WHERE ((@code IS NULL OR ReasonCode LIKE @code + '%') AND (@desc IS NULL OR ReasonDesc LIKE '%' + @desc + '%')
           AND (@desc2 IS NULL OR @useOr = 1 OR ReasonDesc LIKE '%' + @desc2 + '%'))
       OR (@desc2 IS NOT NULL AND @useOr = 1 AND ReasonDesc LIKE '%' + @desc2 + '%');
END
GO

-- frmCasList ListCases: cases of active members (VIEW_Cas_MBM), invoice (C) or dental (D) cases; @by: 1 member no, 2 policy no, 3 name, 4 IC,
-- 5 case no. The clinic name comes from panel clinics (C / Y). Each search option is its own statement with OPTION (RECOMPILE): one catch-all
-- query scanned the 8.5 million cases (about 9 s per search); the member number test (member OR cover number) is two index lookups.
CREATE OR ALTER PROCEDURE cis.Popup_CaseList @caseType char(1), @by int, @text nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @keys TABLE (CASNumber nvarchar(20) PRIMARY KEY);
    IF @by = 1
    BEGIN
        INSERT INTO @keys SELECT DISTINCT v.CASNumber FROM dbo.VIEW_Cas_MBM v
        WHERE v.MBMCLNDataStatus <> 'D' AND v.MBMCLNStatus <> 'D' AND v.HSCLNInd IN ('C', 'M', 'B') AND v.CaseType = @caseType AND v.MBMnumber LIKE @text + '%'
        OPTION (RECOMPILE);
        INSERT INTO @keys SELECT DISTINCT v.CASNumber FROM dbo.VIEW_Cas_MBM v
        WHERE v.MBMCLNDataStatus <> 'D' AND v.MBMCLNStatus <> 'D' AND v.HSCLNInd IN ('C', 'M', 'B') AND v.CaseType = @caseType AND v.MBMCNumber LIKE @text + '%'
          AND NOT EXISTS (SELECT 1 FROM @keys k WHERE k.CASNumber = v.CASNumber)
        OPTION (RECOMPILE);
    END
    ELSE
        INSERT INTO @keys SELECT DISTINCT v.CASNumber FROM dbo.VIEW_Cas_MBM v
        WHERE v.MBMCLNDataStatus <> 'D' AND v.MBMCLNStatus <> 'D' AND v.HSCLNInd IN ('C', 'M', 'B') AND v.CaseType = @caseType
          AND ((@by = 2 AND v.MBMPolicyNo LIKE @text + '%') OR (@by = 3 AND v.MBMName LIKE '%' + @text + '%')
               OR (@by = 4 AND v.MBMIcBcPp LIKE @text + '%') OR (@by = 5 AND v.CASNumber LIKE @text + '%'))
        OPTION (RECOMPILE);
    SELECT v.CASNumber, v.MBMnumber AS MBMNumber, v.CoverID, v.MBMName, v.MBMCName, v.MBMPolicyNo, v.DateVisit, v.CASStatus, v.ClinicCode,
           RTRIM(c.CLNName) AS CLNName, v.MBMIcBcPp
    FROM @keys k
    INNER JOIN dbo.VIEW_Cas_MBM v ON v.CASNumber = k.CASNumber
    LEFT JOIN cis.Hdb_CLN c ON c.CLNCode = v.ClinicCode AND (c.CLNPanel = 'C' OR c.CLNPanel = 'Y')
    WHERE v.MBMCLNDataStatus <> 'D' AND v.MBMCLNStatus <> 'D' AND v.HSCLNInd IN ('C', 'M', 'B') AND v.CaseType = @caseType
    OPTION (RECOMPILE);
END
GO

-- ===== main/440_mma_cpt.sql =====
-- Invoice > Other > MMA/CPT (frmMMASchedule) and its detail pop-up (frmShowInvestigation): HISMaintenance VIEW_MMA_Schedule.

-- CmdSearch_Click: every filter optional. The category filters use the text without its first character, as the desktop (Mid(x, 2, 8)).
CREATE OR ALTER PROCEDURE cis.Mma_Search
    @mainBp nvarchar(50) = NULL, @opcs nvarchar(50) = NULL, @procedure nvarchar(200) = NULL, @surgCat nvarchar(20) = NULL, @anaesCat nvarchar(20) = NULL,
    @part nvarchar(10) = NULL, @pageFrom nvarchar(20) = NULL, @pageTo nvarchar(20) = NULL, @cptCode nvarchar(50) = NULL, @cptDesc nvarchar(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT ProcMainBPCode, OPCSCode, OPCSNarrative, PartACost, CPTCode, CPTDesc, ProcSurgCat, ProcAnaesCat, ProcPageNo, ProcPart, SurgeonsFees,
           AnaesthetistsFees, Surgeon13thNew, Anesthetist13thNew
    FROM cis.Maint_VIEW_MMA_Schedule
    WHERE (@mainBp IS NULL OR ProcMainBPCode LIKE '%' + @mainBp + '%') AND (@opcs IS NULL OR OPCSCode LIKE '%' + @opcs + '%')
      AND (@procedure IS NULL OR OPCSNarrative LIKE '%' + @procedure + '%') AND (@surgCat IS NULL OR ProcSurgCat LIKE '%' + @surgCat + '%')
      AND (@anaesCat IS NULL OR ProcAnaesCat LIKE '%' + @anaesCat + '%') AND (@part IS NULL OR ProcPart = @part)
      AND (@pageFrom IS NULL OR ProcPageNo >= @pageFrom) AND (@pageTo IS NULL OR ProcPageNo <= @pageTo)
      AND (@cptCode IS NULL OR CPTCode = @cptCode) AND (@cptDesc IS NULL OR CPTDesc LIKE '%' + @cptDesc + '%');
END
GO

-- ===== main/450_pass_to_claim.sql =====
-- Claim > Pass to Claim (frmPass2Claim): list cases of CISWorksheet VIEW_ClaimPass2Claim (the desktop's ClaimPass2Claim case 1, parameterised)
-- and pass the checked ones (ClaimPass2ClaimUpdate → ClaimFileTransfer).

-- chkPass_Click: the users who passed cases (HISDB SearchCASToken).
CREATE OR ALTER PROCEDURE cis.Pass2Claim_Users AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Hdb_SearchCASToken;
END
GO

-- SearchRecord → DischargeListing. @caseNo = the scanned case (scan mode), else @casFrom / @casTo; @caseStatus 0 O, 1 D, 2 O or C;
-- @passed 'P' passed / 'N' not passed; @noVisit / @noRegister / @noPass = the date is empty; @casNumbers = the cases just passed
-- (comma-separated, the list shown after Pass).
CREATE OR ALTER PROCEDURE cis.Pass2Claim_Search
    @caseNo nvarchar(20) = NULL, @casFrom nvarchar(20) = NULL, @casTo nvarchar(20) = NULL, @caseStatus char(1) = NULL, @claimability char(1) = NULL,
    @passBy nvarchar(50) = NULL, @passed char(1) = NULL, @noVisit bit = 0, @noRegister bit = 0, @noPass bit = 0, @regFrom datetime = NULL,
    @regTo datetime = NULL, @visitFrom datetime = NULL, @visitTo datetime = NULL, @passFrom datetime = NULL, @passTo datetime = NULL,
    @casNumbers nvarchar(max) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT ClaimStatus, CASNumber, MBMnumber AS MBMNumber, PatientName, RegisterDate, DateVisit, ClaimPassDate, CLNName, OPBatchNo, ClaimPassBy,
           CASStatus, Claimability, CoverID, MBMCoverID, NoofDay
    FROM cis.Ws_VIEW_ClaimPass2Claim
    WHERE (@caseNo IS NULL OR CASNumber = @caseNo) AND (@casFrom IS NULL OR CASNumber >= @casFrom) AND (@casTo IS NULL OR CASNumber <= @casTo)
      AND (@caseStatus IS NULL OR (@caseStatus = '0' AND CASStatus = 'O') OR (@caseStatus = '1' AND CASStatus = 'D')
           OR (@caseStatus = '2' AND (CASStatus = 'O' OR CASStatus = 'C')))
      AND (@claimability IS NULL OR Claimability = @claimability) AND (@passBy IS NULL OR ClaimPassBy = @passBy)
      AND (@passed IS NULL OR (@passed = 'P' AND ClaimStatus = 'P') OR (@passed <> 'P' AND (ClaimStatus IS NULL OR ClaimStatus = '')))
      AND (@noVisit = 0 OR DateVisit IS NULL) AND (@noRegister = 0 OR RegisterDate IS NULL) AND (@noPass = 0 OR ClaimPassDate IS NULL)
      AND (@regFrom IS NULL OR RegisterDate >= @regFrom) AND (@regTo IS NULL OR RegisterDate <= @regTo)
      AND (@visitFrom IS NULL OR DateVisit >= @visitFrom) AND (@visitTo IS NULL OR DateVisit <= @visitTo)
      AND (@passFrom IS NULL OR ClaimPassDate >= @passFrom) AND (@passTo IS NULL OR ClaimPassDate <= @passTo)
      AND (@casNumbers IS NULL OR ',' + @casNumbers + ',' LIKE '%,' + CASNumber + ',%')
    GROUP BY ClaimStatus, CASNumber, MBMnumber, PatientName, RegisterDate, DateVisit, ClaimPassDate, CLNName, OPBatchNo, ClaimPassBy, CASStatus,
             Claimability, CoverID, MBMCoverID, NoofDay
    OPTION (RECOMPILE);
END
GO

-- UpdateRecord: one checked case (the desktop's ClaimPass2ClaimUpdate, status P).
CREATE OR ALTER PROCEDURE cis.Pass2Claim_Pass @casNumber nvarchar(10), @userCode nvarchar(10), @today datetime AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Ws_ClaimPass2ClaimUpdate @CASNumber = @casNumber, @ClaimStatus = 'P', @ClaimPassBy = @userCode, @ClaimPassDate = @today;
END
GO

-- ===== main/460_update_claims_data.sql =====
-- Claim Bordx > Update Claims Data (frmUpdateClaimsData): spreadsheet uploads into CISWorksheet Claim / Bordx.
-- The desktop's uploads could not work (see the screen doc); these are the working versions agreed for the web.

-- UpdateFFS: CLMMCOProcFee of a case.
CREATE OR ALTER PROCEDURE cis.ClaimData_Ffs @casNumber nvarchar(20), @amount money, @today datetime, @userCode nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Ws_Claim SET CLMMCOProcFee = @amount, ClaimLastUpdateDate = @today, ClaimLastUpdateUser = @userCode WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO

-- UpdateCopay: CLMCoPayColl of a case.
CREATE OR ALTER PROCEDURE cis.ClaimData_Copay @casNumber nvarchar(20), @amount money, @today datetime, @userCode nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Ws_Claim SET CLMCoPayColl = @amount, ClaimLastUpdateDate = @today, ClaimLastUpdateUser = @userCode WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO

-- UpdateSent: CLMBordxAmt of a case.
CREATE OR ALTER PROCEDURE cis.ClaimData_BordxAmt @casNumber nvarchar(20), @amount money, @today datetime, @userCode nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Ws_Claim SET CLMBordxAmt = @amount, ClaimLastUpdateDate = @today, ClaimLastUpdateUser = @userCode WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO

-- UpdateBordxDate: the bordereaux month / year of its claims and the Bordx date.
CREATE OR ALTER PROCEDURE cis.ClaimData_BordxDate @bordxRefNo nvarchar(30), @date datetime, @today datetime, @userCode nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Ws_Claim SET CLMBordxMonth = CAST(MONTH(@date) AS varchar(2)), CLMBordxYear = CAST(YEAR(@date) AS varchar(4)),
        ClaimLastUpdateDate = @today, ClaimLastUpdateUser = @userCode
    WHERE BordxRefNo = @bordxRefNo;
    DECLARE @n int = @@ROWCOUNT;
    UPDATE cis.Ws_Bordx SET BordxDate = @date WHERE BordxRefNo = @bordxRefNo;
    SELECT @n;
END
GO

-- ===== main/470_worksheet.sql =====
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

-- ===== main/480_member_supp.sql =====
-- Medix.bas MemberSuppInfo: the member, plan and limit details the invoice screens and the member info pop-up (FrmINVMBMinfo) show.
-- Result sets: 1 member (VIEW_MBM_Clinical), 2 group company, 3 MBMCLNOthers, 4 PLN, 5 CoPaymentMBM, 6 AnnualLimit.SuppLimitStatus,
-- 7 covered person, 8 MBMCoveredPersonsCLN of the cover and of cover 01 (the shared-limit cover). The rules are in MemberSuppService.
CREATE OR ALTER PROCEDURE cis.MemberSupp_Info @mbmNumber nvarchar(50), @covId nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @group nvarchar(200), @hlt nvarchar(20), @pln nvarchar(20), @ins nvarchar(20), @pay nvarchar(20);
    SELECT TOP (1) @group = MBMGroupCompany, @hlt = RTRIM(HLTCode) FROM cis.Hdb_VIEW_MBM_Clinical WHERE MBMNumber = @mbmNumber AND MBMCLNDataStatus <> 'D';
    SELECT TOP (1) @pln = RTRIM(PLNType), @ins = RTRIM(INSType), @pay = RTRIM(PAYCode) FROM cis.Hdb_MBMCLNOthers WHERE MBMNumber = @mbmNumber;

    SELECT TOP (1) MBMNumber, MBMName, MBMIcBcPp, PAYCode, MBMCLNStatus, MBMGroupCompany, MBMDOB, MBMCLNDataStatus, RTRIM(HLTCode) AS HLTCode,
           MBMDepartment, MBMCLNMCDaysBal
    FROM cis.Hdb_VIEW_MBM_Clinical WHERE MBMNumber = @mbmNumber AND MBMCLNDataStatus <> 'D';
    SELECT TOP (1) GRPMCOStatus, GRPMCStatus, GRPPostageFees, GRPBordxFees FROM dbo.GRPCompany WHERE GrpCoName = @group;
    SELECT TOP (1) MBMPolNumber, MBMPolEffDate, MBMPolExpDate, MBMCLNBalLmt, MBMCLNVisitLmt, MBMAccumAmt, MBMCardType, RTRIM(INSType) AS INSType,
           RTRIM(PLNType) AS PLNType, MBMCLNLTM, RTRIM(PAYCode) AS PAYCode, MBMMEPLmt, MBMDentalLmt, MBMSpeclmt, MBMDentalExtLmt, MBMCancelDate,
           MBMFirstSpecLmt
    FROM cis.Hdb_MBMCLNOthers WHERE MBMNumber = @mbmNumber;
    SELECT TOP (1) AnnLmtIND, VisLmtIND, ClnContract, PDentalStatus, SDentalStatus, PSpecStatus, SSpecStatus, PMEPStatus, SMEPStatus, BNFStatus,
           PDentalLimit, SDentalLimit, PSpecLimit, SSpecLimit, ANNLmtIndSupp, VISLmtIndSupp, GrossAmtStatus, PLNMMIndPerc, ASOInd
    FROM dbo.PLN WHERE PLNCode = @pln AND HLTCode = @hlt AND PAYCode = @pay;
    SELECT TOP (1) CoAmount FROM dbo.CoPaymentMBM WHERE HLTType = @hlt AND PlnCode = @pln;
    SELECT TOP (1) SuppLimitStatus FROM dbo.AnnualLimit WHERE PLNCode = @pln AND HLTCode = @hlt AND INSCode = @ins;
    SELECT TOP (1) MBMCName, MBMCCoverID, MBMCICBCPPNo, MBMCDOB, LTRIM(RTRIM(MBMCRELCode)) AS MBMCRELCode
    FROM cis.Hdb_MBMCoveredPersons WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @covId;
    SELECT MBMCCoverID, MBMCCLNBalLmt, MBMCCLNVisitLmt, MBMCCLNDentalLmt, MBMCCLNSpecLmt, MBMCCLNMEPLmt, MBMCCLNDentalExtLmt, MBMCCLNFirstSpecLmt,
           MBMCAccumAmt, MBMCCardType, MBMCCancelDate, MBMCCLNExpDate, MBMCCLNEffDate, MBMCCLNStatus
    FROM cis.Hdb_MBMCoveredPersonsCLN WHERE MBMCNumber = @mbmNumber AND MBMCCoverID IN (@covId, '01');
END
GO

-- ===== main/490_clinic_invoice.sql =====
-- Invoice > Clinic > Clinic (frmInvoice): look-ups, the member's claims (CIS_FFS2013) and the save. Limits / top-ups / ASO / specialist
-- limits reuse cis.MbmAdj_Limits, cis.MbmEnq_Topup, cis.MbmEnq_Aso and cis.MbmEnq_SpecLimit (Phase 2).

-- ComboListing: receipt modes and process months.
CREATE OR ALTER PROCEDURE cis.InvClinic_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(RecModeCode) AS Code, RTRIM(RecModeDesc) AS Name FROM dbo.InvRecMode;
    SELECT RTRIM(INVProcMth) + '/' + RTRIM(INVProcYR) AS Month FROM dbo.INVProcMonth;
END
GO

-- TxtClinicCode_LostFocus (CLN) and GetClinicLimit (Tbl_CLNGPLimits).
CREATE OR ALTER PROCEDURE cis.InvClinic_Clinic @clinicCode nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) RTRIM(CLNPanel) AS Panel, RTRIM(CLNCode) AS Code, RTRIM(CLNName) AS Name, RTRIM(CLNType) AS ClinicType, RTRIM(CLNFFSStatus) AS FfsStatus,
           CLNMCOProcfeePerc AS McoProcFeePercent
    FROM cis.Hdb_CLN WHERE CLNCode = @clinicCode;
END
GO

-- CmdGo_Click look-ups after MemberSuppInfo: 1 AnnualLimit + PLN visit limits, 2 PLN maternity / dental / ASO, 3 GP / SP limits,
-- 4 SpecialistLimit status, 5 MBMOthers + GRPCompany (claim cost centre etc.), 6 Tbl_ASOMemberDetails.
CREATE OR ALTER PROCEDURE cis.InvClinic_MemberExtra @mbmNumber nvarchar(50), @covId nvarchar(2), @planCode nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) RTRIM(p.HealthlimitIND) AS HealthLimitInd, a.HealthLimit, a.VisitLimit, a.prepostmatellimit AS PrePostMatLimit, p.PDentalStatus, a.PDentalAnnLimit, a.SPVisitLimit, a.MEVisitLimit, RTRIM(p.VISLmtInd) AS VisLmtInd,
           RTRIM(p.ChildLimitInd) AS ChildLimitInd, a.ChildDepLimit
    FROM dbo.AnnualLimit a INNER JOIN dbo.PLN p ON p.PLNCode = a.PLNCode WHERE a.PLNCode = @planCode;
    SELECT TOP (1) RTRIM(PLNMaternityInd) AS MaternityInd, RTRIM(PDentalStatus) AS DentalInd, RTRIM(ASOInd) AS AsoInd FROM dbo.PLN WHERE PLNCode = @planCode;
    IF @covId = '00'
        SELECT TOP (1) p.GRPCompany, p.PrinGPLimit AS GPLimit, p.PrinSPLimit AS SPLimit
        FROM cis.Hdb_MBMCLNOthers o INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType WHERE o.MBMNumber = @mbmNumber;
    ELSE
        SELECT TOP (1) p.GRPCompany, p.DepGPLimit AS GPLimit, p.DepSPLimit AS SPLimit
        FROM cis.Hdb_MBMCoveredPersonsCLN c INNER JOIN dbo.PLN p ON p.PLNCode = c.MBMCCLNPLNCode WHERE c.MBMCNumber = @mbmNumber AND c.MBMCCoverID = @covId;
    SELECT TOP (1) RTRIM(SuppLimitStatus) AS SuppLimitStatus FROM dbo.SpecialistLimit WHERE PLNCode = @planCode;
    SELECT TOP (1) LTRIM(RTRIM(COALESCE(o.MBMDepartment, ''))) AS MBMDepartment, o.MBMCostCenter, LTRIM(RTRIM(COALESCE(o.MBMDeptCode, ''))) AS MBMDeptCode,
           LTRIM(RTRIM(COALESCE(o.Subsidiary, ''))) AS Subsidiary, LTRIM(RTRIM(g.GRPCoCode)) AS GRPCoCode
    FROM cis.Hdb_MBMOthers o INNER JOIN dbo.GRPCompany g ON g.GRPCoName = o.MBMGroupCompany WHERE o.MBMNumber = @mbmNumber;
    SELECT MBMASOLimit, MBMOPLimit, MBMIPUtilized, MBMOPUtilized, RTRIM(IsIPOPCombined) AS IsIPOPCombined
    FROM cis.Hdb_ASOMemberDetails WHERE MBMNumber = @mbmNumber AND MBMCoverID = @covId;
END
GO

-- CIS_FFS2013 with the dynamic filter replaced: @mode 0 the member, 1 cover 00, 2 covers other than 00, 3 one cover.
CREATE OR ALTER PROCEDURE cis.InvClinic_Claims @mbmNumber nvarchar(50), @mode int, @covId nvarchar(2) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT cl.CLMMCOProcFee AS FFS, RTRIM(cln.CLNType) AS CLNType, COALESCE(ia.CLMMaternity, 0) AS CLMMaternity, c.CoverID,
           CASE WHEN p.CoPayInd = 'Y' AND COALESCE(cl.CLMCoPayColl, 0) <> 0 THEN cl.CLMTotalActual - cl.CLMCoPayColl ELSE cl.CLMTotalActual END AS CLMBordxAmt,
           RTRIM(p.PLNGPSPShared) AS PLNGPSPShared,
           CASE WHEN COALESCE(cl.CLMBTBGAmt, 0) <> 0 THEN COALESCE(cl.CLMBTBGAmt, 0) ELSE COALESCE(ia.CLMOthersNoDed, 0) END AS CLMOthersNoDed,
           ia.CLMTyphiadJabNoDed, COALESCE(ia.CLMDental, 0) AS CLMDental
    FROM dbo.CISCas c WITH (NOLOCK)
    INNER JOIN cis.Ws_Claim cl WITH (NOLOCK) ON c.CASNumber = cl.CASNumber
    INNER JOIN cis.Hdb_MBMCLNOthers o ON o.MBMNumber = c.MBMnumber
    INNER JOIN cis.Hdb_MBMTwo t ON t.MBMNumber = c.MBMnumber
    INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    INNER JOIN cis.Ws_ClaimInvoiceActualNew ia WITH (NOLOCK) ON cl.CASNumber = ia.CASNumber
    LEFT JOIN cis.Hdb_CLN cln WITH (NOLOCK) ON c.ClinicCode = cln.CLNCode
    WHERE c.CASStatus <> 'D' AND c.MBMnumber = @mbmNumber
      AND (@mode = 0 OR (@mode = 1 AND c.CoverID = '00') OR (@mode = 2 AND c.CoverID <> '00') OR (@mode = 3 AND c.CoverID = @covId));
END
GO

-- ChkData: cases of the member / cover on the visit date at the clinic (the duplicate check); cmdVerify: CheckItems CSMD and DiagNotPayable.
CREATE OR ALTER PROCEDURE cis.InvClinic_Checks @mbmNumber nvarchar(50), @covId nvarchar(2), @visitDate datetime, @clinicCode nvarchar(10), @card nvarchar(50),
    @planCode nvarchar(20), @diag1 nvarchar(20) = NULL, @diag2 nvarchar(20) = NULL, @diag3 nvarchar(20) = NULL, @diag4 nvarchar(20) = NULL, @diag5 nvarchar(20) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT CoverID, DateVisit, ClinicCode, InvoiceStmtAmt FROM dbo.CISCAS
    WHERE MBMnumber = @mbmNumber AND CoverID = @covId AND DateVisit = @visitDate AND ClinicCode = @clinicCode;
    SELECT Amount FROM dbo.CheckItems WHERE CatCode = 'CSMD' AND CatName = @card;
    SELECT RTRIM(DiagCode) AS DiagCode FROM dbo.DiagNotPayable WHERE PLNCode = @planCode AND DiagCode IN (@diag1, @diag2, @diag3, @diag4, @diag5);
END
GO

-- GenerateCISCaseNo: WebSearchCASNo_New2013test (pay code EA) → "EA" + 8 digits.
CREATE OR ALTER PROCEDURE cis.InvClinic_CaseNo AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @no nvarchar(2000);
    DECLARE @rows TABLE (v sql_variant);
    -- The procedure also returns the number as a result set; it is swallowed so only the formatted number comes back.
    INSERT INTO @rows EXEC dbo.WebSearchCASNo_New2013test @PayCode = 'EA', @CaseNo1 = @no OUTPUT;
    SELECT 'EA' + RIGHT('00000000' + CAST(CAST(@no AS bigint) AS varchar(20)), 8);
END
GO

-- SaveDup_ori: number the cases of the same member / cover / visit date / clinic when there is more than one.
CREATE OR ALTER PROCEDURE cis.InvClinic_Duplicates @mbmNumber nvarchar(50), @covId nvarchar(2), @visitDate datetime, @clinicCode nvarchar(10),
    @userCode nvarchar(50), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF (SELECT COUNT(*) FROM dbo.CISCas WHERE MBMnumber = @mbmNumber AND CoverID = @covId AND DateVisit = @visitDate AND ClinicCode = @clinicCode) <= 1 RETURN;
    DECLARE @seq nvarchar(20) = (SELECT TOP (1) RTRIM(DupSeqNo) FROM dbo.CISCas
        WHERE MBMnumber = @mbmNumber AND CoverID = @covId AND DateVisit = @visitDate AND ClinicCode = @clinicCode GROUP BY DupSeqNo ORDER BY DupSeqNo DESC);
    IF ISNULL(@seq, '') = ''
    BEGIN
        DECLARE @n int = (SELECT TOP (1) CAST(DupNo AS int) FROM dbo.DupClaimSeq WITH (UPDLOCK, HOLDLOCK) WHERE Code = 'DUP');
        IF @n IS NULL
        BEGIN
            SET @n = 1;
            INSERT INTO dbo.DupClaimSeq (Code, DupNo) VALUES ('DUP', '000001');
        END
        ELSE
        BEGIN
            SET @n += 1;
            UPDATE dbo.DupClaimSeq SET DupNo = RIGHT('000000' + CAST(@n AS varchar(10)), 6) WHERE Code = 'DUP';
        END
        SET @seq = 'DUP' + RIGHT('000000' + CAST(@n AS varchar(10)), 6);
    END
    UPDATE dbo.CISCas SET DupSeqNo = @seq, DupUpdateUser = @userCode, DupUpdateDate = @today, DupAppInd = 'N'
    WHERE MBMnumber = @mbmNumber AND CoverID = @covId AND DateVisit = @visitDate AND ClinicCode = @clinicCode;
END
GO

-- UpdateMBMBalance / UpdateMBMMCDays (HISDB): the clinical limit row of the principal, cover 01 (shared limit B / D) or the cover (C).
CREATE OR ALTER PROCEDURE cis.InvClinic_Balance @mbmNumber nvarchar(50), @limitCover nvarchar(2) = NULL, @adjAnnLmt money, @approved money,
    @deduct money, @full bit, @mcDays money
AS
BEGIN
    SET NOCOUNT ON;
    IF @limitCover IS NULL
        UPDATE cis.Hdb_MBMCLNOthers SET MBMAdjAnnLmt = COALESCE(TRY_CONVERT(money, MBMAdjAnnLmt), 0) + @adjAnnLmt,
            MBMAccumAmt = COALESCE(TRY_CONVERT(money, MBMAccumAmt), 0) + @approved,
            MBMCLNBalLmt = CASE WHEN @full = 1 THEN MBMCLNBalLmt
                                WHEN COALESCE(TRY_CONVERT(money, MBMCLNBalLmt), 0) - @deduct < 0 THEN 0 ELSE COALESCE(TRY_CONVERT(money, MBMCLNBalLmt), 0) - @deduct END
        WHERE MBMNumber = @mbmNumber;
    ELSE
        UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCAdjAnnLmt = COALESCE(TRY_CONVERT(money, MBMCAdjAnnLmt), 0) + @adjAnnLmt,
            MBMCAccumAmt = COALESCE(TRY_CONVERT(money, MBMCAccumAmt), 0) + @approved,
            MBMCCLNBalLmt = CASE WHEN @full = 1 THEN MBMCCLNBalLmt
                                 WHEN COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) - @deduct < 0 THEN 0 ELSE COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) - @deduct END
        WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @limitCover;
    UPDATE cis.Hdb_MBMCLNOthers SET MBMCLNMCDaysBal = CASE WHEN COALESCE(TRY_CONVERT(money, MBMCLNMCDaysBal), 0) > @mcDays
        THEN COALESCE(TRY_CONVERT(money, MBMCLNMCDaysBal), 0) - @mcDays ELSE 0 END
    WHERE MBMNumber = @mbmNumber;
END
GO

-- VerifyDiagnosis + TreatmentReason.ReasonOtherDiag (diagnosis 5 "others" indicator).
CREATE OR ALTER PROCEDURE cis.InvClinic_Diagnosis @code nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) RTRIM(ReasonCode) AS Code, RTRIM(ReasonDesc) AS Description, RTRIM(ReasonOtherDiag) AS OtherIndicator
    FROM dbo.TreatmentReason WHERE ReasonCode = @code;
END
GO

-- ===== main/500_member_invoice.sql =====
-- Invoice > Clinic > Member (frmInvoiceMBM): reimbursement invoices. Go reuses the clinic invoice procedures; the desktop saves with recordset
-- inserts (not the clinic's procedures), reproduced here as parameterised INSERTs.

-- ValidateIbmUser: sp_IBMMemberInformation → Number ("0" = not an IBM member).
CREATE OR ALTER PROCEDURE cis.InvMember_Ibm @mbmNumber nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_IBMMemberInformation @MemeberNo = @mbmNumber;
END
GO

-- SaveCase: CISCAS.
CREATE OR ALTER PROCEDURE cis.InvMember_Case
    @casNumber nvarchar(20), @mbmNumber nvarchar(50), @coverId nvarchar(2), @userCode nvarchar(50), @today datetime, @clinicCode nvarchar(10),
    @npClinic nvarchar(200), @invoiceNo nvarchar(50), @invoiceDate datetime, @invoiceStmtAmt money, @dateVisit datetime, @mcDays money,
    @visitTime nvarchar(10) = NULL, @opBatchNo nvarchar(30),
    @t1 nvarchar(200), @t2 nvarchar(200), @t3 nvarchar(200), @t4 nvarchar(200), @t5 nvarchar(200), @t6 nvarchar(200), @t7 nvarchar(200),
    @d1 nvarchar(10), @d2 nvarchar(10), @d3 nvarchar(10), @d4 nvarchar(10), @d5 nvarchar(10), @d6 nvarchar(10), @d7 nvarchar(10),
    @q1 nvarchar(10), @q2 nvarchar(10), @q3 nvarchar(10), @q4 nvarchar(10), @q5 nvarchar(10), @q6 nvarchar(10), @q7 nvarchar(10),
    @receiptDate datetime, @receiptMode nvarchar(10), @remark nvarchar(500), @r1 nvarchar(20), @r2 nvarchar(20), @r3 nvarchar(20), @r4 nvarchar(20),
    @r5 nvarchar(20), @diagnosis5 nvarchar(200), @mcStatus int, @noOfDiag int, @noOfMedication int, @clnType nvarchar(2) = NULL,
    @repReason nvarchar(500), @repAction nvarchar(500), @repAmt nvarchar(20), @isPharmacy nvarchar(2)
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.CISCAS (CASNumber, MBMNumber, CoverID, CASStatus, Claimability, CaseType, Pending, RegisterDate, RegisterBy, LastUpdateDate,
        LastUpdateBy, ClinicCode, NPClinic, InvoiceNo, InvoiceDate, InvoiceStmtAmt, DateVisit, MCDays, VisitTime, OPBatchNo,
        MedInjType1, MedInjType2, MedInjType3, MedInjType4, MedInjType5, MedInjType6, MedInjType7, MedInjDays1, MedInjDays2, MedInjDays3, MedInjDays4,
        MedInjDays5, MedInjDays6, MedInjDays7, MedInjQty1, MedInjQty2, MedInjQty3, MedInjQty4, MedInjQty5, MedInjQty6, MedInjQty7, InvoiceReceiptDate,
        InvoiceReceiptMode, CasRemark, CASTreatReason1, CASTreatReason2, CASTreatReason3, CASTreatReason4, CASTreatReason5, CASDiagnosis5, MCStatus,
        NoOfDiag, NoOfMedication, CLNType, RepudiationReason, CASRepAction, CASRepAmt, IsPharmacy)
    VALUES (@casNumber, @mbmNumber, @coverId, 'O', 'A', 'C', 'N', @today, @userCode, @today, @userCode, @clinicCode, @npClinic, @invoiceNo, @invoiceDate,
        @invoiceStmtAmt, @dateVisit, @mcDays, @visitTime, @opBatchNo, @t1, @t2, @t3, @t4, @t5, @t6, @t7, @d1, @d2, @d3, @d4, @d5, @d6, @d7,
        @q1, @q2, @q3, @q4, @q5, @q6, @q7, @receiptDate, @receiptMode, @remark, @r1, @r2, @r3, @r4, @r5, @diagnosis5, @mcStatus, @noOfDiag,
        @noOfMedication, @clnType, @repReason, @repAction, @repAmt, @isPharmacy);
END
GO

-- SaveInvoice: ClaimInvoice (amount less the co-pay), paid to the member.
CREATE OR ALTER PROCEDURE cis.InvMember_Invoice @casNumber nvarchar(20), @userCode nvarchar(50), @today datetime, @billNo nvarchar(50), @billDate datetime,
    @amount money
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO cis.Ws_ClaimInvoice (CASNumber, INVEnteredDate, INVEnteredBy, INVLastUpdateBy, INVLastUpdateDate, BillNo, BillDate, INVAmount, INVBillAmt,
        INVPayTo, INVPayee)
    VALUES (@casNumber, @today, @userCode, @userCode, @today, @billNo, @billDate, @amount, @amount, 'M', 'MEMBER');
END
GO

-- saveActual: ClaimInvoiceActualNew (the 28 items, co-payment, descriptions, vision, others without deduction).
CREATE OR ALTER PROCEDURE cis.InvMember_Actual @casNumber nvarchar(20),
    @a1 money, @a2 money, @a3 money, @a4 money, @a5 money, @a6 money, @a7 money, @a8 money, @a9 money, @a10 money, @a11 money, @a12 money, @a13 money,
    @a14 money, @a15 money, @a16 money, @a17 money, @a18 money, @a19 money, @a20 money, @a21 money, @a22 money, @a23 money, @a24 money, @a25 money,
    @a26 money, @a27 money, @a28 money, @coPayment money, @pOthDesc nvarchar(200), @lOthDesc nvarchar(200), @iOthDesc nvarchar(200), @vision money,
    @othersNoDed money
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO cis.Ws_ClaimInvoiceActualNew (CASNumber, CLMGPFeeBasicMed, CLMGPFeeLTMed, CLMGPFeeMedInj, CLMSPFeeBasicMed, CLMSPFeeLTMed, CLMSPFeeMedInj,
        CLMDental, CLMMaternity, CLMPRVCheckup, CLMXRay, CLMUltraSound, CLMECG, CLMDressing, CLMNebulizer, CLMID, CLMTSSTO, CLMMinorSurgery, CLMPOthers,
        CLMBloodDrawing, CLMBloodGlucose, CLMUrine, CLMPapSmear, CLMLOthers, CLMDTPPolio, CLMDTPPolioHIB, CLMHepatitisB, CLMMMR, CLMIOthers, CLMCoPayment,
        CLMPOthersDesc, CLMLOthersDesc, CLMIOthersDesc, CLMVision, CLMOthersNoDed)
    VALUES (@casNumber, @a1, @a2, @a3, @a4, @a5, @a6, @a7, @a8, @a9, @a10, @a11, @a12, @a13, @a14, @a15, @a16, @a17, @a18, @a19, @a20, @a21, @a22, @a23,
        @a24, @a25, @a26, @a27, @a28, @coPayment, @pOthDesc, @lOthDesc, @iOthDesc, @vision, @othersNoDed);
END
GO

-- saveActualOthers: ClaimInvoiceActualOthers (GP / SP fee parts).
CREATE OR ALTER PROCEDURE cis.InvMember_ActualOthers @casNumber nvarchar(20), @gpFee money, @gpLtMed money, @gpInj money, @gpMed money, @spFee money,
    @spLtMed money, @spInj money, @spMed money
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO cis.Ws_ClaimInvoiceActualOthers (CASNumber, GPFee, GPLTMed, GPInj, GPMed, SPFee, SPLTMed, SPInj, SPMed)
    VALUES (@casNumber, @gpFee, @gpLtMed, @gpInj, @gpMed, @spFee, @spLtMed, @spInj, @spMed);
END
GO

-- SaveINVClaim: Claim.
CREATE OR ALTER PROCEDURE cis.InvMember_Claim @casNumber nvarchar(20), @userCode nvarchar(50), @today datetime, @payableM money, @clnExcess money,
    @totalActual money, @finalApprove money, @balanceLimit money, @previousBalance money, @paidM money, @paidC money, @postageFee money, @adjAnnLmt money,
    @nonGpFee money, @bordxAmt money, @coPay money, @memo char(1), @refLetter nvarchar(2), @deptName nvarchar(200), @costCenter nvarchar(100),
    @deptCode nvarchar(100), @subsidiary nvarchar(100), @grpCoCode nvarchar(50)
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO cis.Ws_Claim (CASNumber, ClaimOpenDate, ClaimOpenBy, ClaimCloseDate, ClaimClosedBy, ClaimLastUpdateDate, ClaimLastUpdateUser,
        CLMPayableByMedix2M, CLMCLNExcess, CLMTotalActual, CLMFinalApprove, CLMBalanceLimit, CLMPreviousBalance, CLMPaidBYMedix2M, CLMPaidBYMedix2C,
        CLMPostageFee, CLMAdjAnnLmt, ClmNonGPFee, CLMBordxAmt, CLMCoPayColl, CLMCoPayCLNColl, CLMBordxMemo, CLMRefLetter, CLMDeptName, CLMCostCenter,
        CLMDeptCode, CLMSubsidiary, CLMGrpComCode)
    VALUES (@casNumber, @today, @userCode, @today, @userCode, @today, @userCode, @payableM, @clnExcess, @totalActual, @finalApprove, @balanceLimit,
        @previousBalance, @paidM, @paidC, @postageFee, @adjAnnLmt, @nonGpFee, @bordxAmt, @coPay, @coPay, @memo, @refLetter, @deptName, @costCenter,
        @deptCode, @subsidiary, @grpCoCode);
END
GO

-- UpdateMBMBalance (member invoice): adjusted annual limit, accumulated amount, balance and the health limit.
CREATE OR ALTER PROCEDURE cis.InvMember_Balance @mbmNumber nvarchar(50), @limitCover nvarchar(2) = NULL, @adjAnnLmt money, @approved money,
    @deduct money, @full bit, @health bit
AS
BEGIN
    SET NOCOUNT ON;
    IF @limitCover IS NULL
        UPDATE cis.Hdb_MBMCLNOthers SET MBMAdjAnnLmt = COALESCE(MBMAdjAnnLmt, 0) + @adjAnnLmt, MBMAccumAmt = COALESCE(MBMAccumAmt, 0) + @approved,
            MBMCLNBalLmt = CASE WHEN @full = 1 THEN MBMCLNBalLmt WHEN COALESCE(MBMCLNBalLmt, 0) - @deduct < 0 THEN 0 ELSE COALESCE(MBMCLNBalLmt, 0) - @deduct END,
            MBMHealthlimit = CASE WHEN @health = 1 THEN COALESCE(TRY_CONVERT(money, MBMHealthlimit), 0) - @approved ELSE MBMHealthlimit END
        WHERE MBMNumber = @mbmNumber;
    ELSE
        UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCAdjAnnLmt = COALESCE(TRY_CONVERT(money, MBMCAdjAnnLmt), 0) + @adjAnnLmt,
            MBMCAccumAmt = COALESCE(TRY_CONVERT(money, MBMCAccumAmt), 0) + @approved,
            MBMCCLNBalLmt = CASE WHEN @full = 1 THEN MBMCCLNBalLmt
                                 WHEN COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) - @deduct < 0 THEN 0 ELSE COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) - @deduct END,
            MBMCHealthlimit = CASE WHEN @health = 1 THEN COALESCE(TRY_CONVERT(money, MBMCHealthlimit), 0) - @approved ELSE MBMCHealthlimit END
        WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @limitCover;
END
GO

-- ===== main/510_invoice_enquiry.sql =====
-- Invoice > Other > Enquiry (FrmInvoiceEnq): one case read-only — the case (SP_INV_CLAIM), its stored worksheet, bordereaux, payments,
-- receipts and medication requests.

-- InvoiceINFO: SP_INV_CLAIM (already parameterised) and the GP / SP fee parts.
CREATE OR ALTER PROCEDURE cis.InvEnq_Case @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Ws_SP_INV_CLAIM @StrSql = @casNumber;
    SELECT TOP (1) GPFee, GPLTMed, GPInj, GPMed, SPFee, SPLTMed, SPInj, SPMed FROM cis.Ws_ClaimInvoiceActualOthers WHERE CASNumber = @casNumber;
END
GO

-- CmdNext_Click / CmdPrev_Click: the next / previous clinical case of an active member.
CREATE OR ALTER PROCEDURE cis.InvEnq_Step @casNumber nvarchar(20), @next bit AS
BEGIN
    SET NOCOUNT ON;
    IF @next = 1
        SELECT TOP (1) CASNumber FROM dbo.VIEW_Cas_MBM WHERE MBMCLNDataStatus <> 'D' AND MBMCLNStatus <> 'D' AND HSCLNInd IN ('C', 'M', 'B') AND CaseType = 'C'
            AND CASNumber > @casNumber ORDER BY CASNumber;
    ELSE
        SELECT TOP (1) CASNumber FROM dbo.VIEW_Cas_MBM WHERE MBMCLNDataStatus <> 'D' AND MBMCLNStatus <> 'D' AND HSCLNInd IN ('C', 'M', 'B') AND CaseType = 'C'
            AND CASNumber < @casNumber ORDER BY CASNumber DESC;
END
GO

-- WSInfo: Claim and the seven worksheet tables of the case (actual, contracted rate, approved, limit, member excess, clinic excess, clinic approved).
CREATE OR ALTER PROCEDURE cis.InvEnq_Worksheet @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) * FROM cis.Ws_Claim WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMActualNew WHERE CLMNo = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMClinicLmtNew WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMAppNew WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMLimitNew WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMMemberNew WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMClinicNew WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMClinicAppNew WHERE CASNumber = @casNumber;
END
GO

-- ClaimBordxInfo: SP_Claim_Bordx.
CREATE OR ALTER PROCEDURE cis.InvEnq_Bordx @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Ws_SP_Claim_Bordx @tmpCasNo = @casNumber;
END
GO

-- PaymentInfo / ReceiptInfo (CISPayment).
CREATE OR ALTER PROCEDURE cis.InvEnq_Payments @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT PVCLNNo AS PvNo, PVCLNDate AS PvDate, PVCLNChequeAmt AS ChequeAmt, 'CLINIC' AS PayTo, CLNPaidAmt AS PaidAmt, CLNRemark AS Remark,
           PVCLNChequeDate AS ChequeDate
    FROM cis.Pay_Payment2CLN WHERE CASNumber = @casNumber
    UNION ALL
    SELECT PVNo, PVDate, PVCLNChequeAmt, PVPayto, MBMPaidAmt, Remark, PVChequeDate FROM cis.Pay_Payment2MBM WHERE CASNumber = @casNumber;
    SELECT ReceiptNo, ReceiptDate, ReceiptAmt FROM cis.Pay_ReceiptFrMBM WHERE CASNumber = @casNumber;
END
GO

-- MedicationRequestInfo / MedicineDetails.
CREATE OR ALTER PROCEDURE cis.InvEnq_MedRequests @casNumber nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_CISMedRequestInfo @CaseNo = @casNumber;
    EXEC dbo.sp_CISMedReqDOMedicine @CaseNo = @casNumber;
END
GO

-- ===== main/520_invoice_adjustment.sql =====
-- Invoice > Clinic > Adjustment (frmInvoiceAdj, with frmCasePending): load a saved case, check and update it, audit review, void bordereaux.
-- The save calls the desktop's INVAdj_* procedures with parameters (the desktop pasted formatted amounts into SQL text).

-- InvoiceINFO + CmdGo_Click: SearchCASOP, the case's Claim row, payments, SUN posting, the latest pending reasons, the specialist co-pay,
-- same-visit cases and the plan flags.
CREATE OR ALTER PROCEDURE cis.InvAdj_Load @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Ws_SearchCASOP @CASNumber = @casNumber;
    SELECT TOP (1) CLMCoPayColl, CLMFinalApprove, ClmNonGPFee, CLMAdjAnnLmt, CLMRefLetter, CLMLimitExceeded, CLMPayableByMedix2C, CLMMCOProcStatus,
           CLMMCOProcFee, CLMProcMth, CLMProcYR, CLMCoPayCLNColl, CLMBordxMemo, CLMPayableByMedix2M
    FROM cis.Ws_Claim WHERE CASNumber = @casNumber;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM cis.Pay_Payment2CLN WHERE CASNumber = @casNumber)
                       OR EXISTS (SELECT 1 FROM cis.Pay_Payment2MBM WHERE CASNumber = @casNumber) THEN 1 ELSE 0 END AS bit) AS Paid,
           CAST(CASE WHEN EXISTS (SELECT 1 FROM cis.Ws_SUNPosting WHERE InvoiceNo = @casNumber) THEN 1 ELSE 0 END AS bit) AS Posted;
    DECLARE @ids TABLE (ReasonID nvarchar(400));
    INSERT INTO @ids EXEC dbo.spGetPendingReasonsByCase @casNo = @casNumber, @latest = 1;
    SELECT ReasonID FROM @ids;
    EXEC dbo.SP_SpecialistCopayamt @CASNumber = @casNumber;
    -- The member's e-mail (MBMTwo), the plan's special remark (shown on Edit) and BTBG flag.
    DECLARE @mbm nvarchar(50) = (SELECT TOP (1) MBMnumber FROM dbo.CISCAS WHERE CASNumber = @casNumber);
    SELECT TOP (1) RTRIM(MBMEmail) AS MBMEmail FROM cis.Hdb_MBMTwo WHERE MBMNumber = @mbm;
    SELECT TOP (1) RTRIM(p.PLNSpecialRemark) AS PLNSpecialRemark, p.PLNBTBG FROM cis.Hdb_MBMCLNOthers o INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    WHERE o.MBMNumber = @mbm;
END
GO

-- PopulatePendingReasons: the active reasons of the category (pay to C / M).
CREATE OR ALTER PROCEDURE cis.InvAdj_PendingReasons @category char(1) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.spGetPendingReasons_OP @Category = @category, @Active = 1;
END
GO

-- duplicatecasesIdentify (on Go) / duplicateinvoicesIdentify (UB members) / CheckDoubleEntry.
CREATE OR ALTER PROCEDURE cis.InvAdj_SameVisitCases @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.spMemberduplicatecases_Identify @CasNumber = @casNumber;
END
GO

CREATE OR ALTER PROCEDURE cis.InvAdj_DuplicateInvoices @casNumber nvarchar(20), @invoiceDate date, @billAmt nvarchar(30), @invoiceNo nvarchar(50),
    @patientName nvarchar(200), @patientIc nvarchar(50), @userCode nvarchar(50), @coverId nvarchar(2)
AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.spMemberduplicateInvoices_Identify @CaseNo = @casNumber, @Invreceiptdate = @invoiceDate, @InvBillamt = @billAmt, @Invno = @invoiceNo,
        @PatientName = @patientName, @PatientIC = @patientIc, @Userid = @userCode, @Coverid = @coverId;
END
GO

CREATE OR ALTER PROCEDURE cis.InvAdj_DoubleEntry @mbmNumber nvarchar(50), @coverId nvarchar(2), @clinicCode nvarchar(10), @invoiceNo nvarchar(200),
    @visitTime nvarchar(20), @dateVisit datetime, @totalActual money, @casNumber nvarchar(100)
AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_chkDoubleCaseEntry @MBMNumber = @mbmNumber, @CoverID = @coverId, @ClnCode = @clinicCode, @InvoiceNo = @invoiceNo, @VisitTime = @visitTime,
        @DateVisit = @dateVisit, @TotalActualAmt = @totalActual, @CasNumber = @casNumber;
END
GO

-- CheckExceedLimit: the member's claims (CIS_FFS2018 filter) without this case.
CREATE OR ALTER PROCEDURE cis.InvAdj_OtherClaims @mbmNumber nvarchar(50), @mode int, @covId nvarchar(2), @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @rows TABLE (FFS money, CLNType nvarchar(5), CLMMaternity money, CoverID nvarchar(2), CLMBordxAmt money, PLNGPSPShared nvarchar(2),
        CLMOthersNoDed money, CLMTyphiadJabNoDed money, CLMDental money, CASNumber nvarchar(20));
    INSERT INTO @rows
    SELECT cl.CLMMCOProcFee, RTRIM(cln.CLNType), COALESCE(ia.CLMMaternity, 0), c.CoverID,
           CASE WHEN p.CoPayInd = 'Y' AND COALESCE(cl.CLMCoPayColl, 0) <> 0 THEN cl.CLMTotalActual - cl.CLMCoPayColl ELSE cl.CLMTotalActual END,
           RTRIM(p.PLNGPSPShared), CASE WHEN COALESCE(cl.CLMBTBGAmt, 0) <> 0 THEN COALESCE(cl.CLMBTBGAmt, 0) ELSE COALESCE(ia.CLMOthersNoDed, 0) END,
           ia.CLMTyphiadJabNoDed, COALESCE(ia.CLMDental, 0), c.CASNumber
    FROM dbo.CISCas c WITH (NOLOCK)
    INNER JOIN cis.Ws_Claim cl WITH (NOLOCK) ON c.CASNumber = cl.CASNumber
    INNER JOIN cis.Hdb_MBMCLNOthers o ON o.MBMNumber = c.MBMnumber
    INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    INNER JOIN cis.Ws_ClaimInvoiceActualNew ia WITH (NOLOCK) ON cl.CASNumber = ia.CASNumber
    LEFT JOIN cis.Hdb_CLN cln WITH (NOLOCK) ON c.ClinicCode = cln.CLNCode
    WHERE c.CASStatus <> 'D' AND c.MBMnumber = @mbmNumber AND c.CASNumber <> @casNumber
      AND (@mode = 0 OR (@mode = 1 AND c.CoverID = '00') OR (@mode = 2 AND c.CoverID <> '00') OR (@mode = 3 AND c.CoverID = @covId));
    SELECT FFS, CLNType, CLMMaternity, CoverID, CLMBordxAmt, PLNGPSPShared, CLMOthersNoDed, CLMTyphiadJabNoDed, CLMDental FROM @rows;
END
GO

-- SaveRecord's CISDB / CISWorksheet writes besides the INVAdj procedures: QA, LG amount, medication charges, OP batch.
CREATE OR ALTER PROCEDURE cis.InvAdj_SaveExtras @casNumber nvarchar(20), @qa nvarchar(max), @glAmount money,
    @c1 money, @c2 money, @c3 money, @c4 money, @c5 money, @c6 money, @c7 money
AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.spUpdateQA @casNumber = @casNumber, @QA = @qa;
    UPDATE dbo.LG SET GLAmt = @glAmount WHERE LGNo = @casNumber;
    UPDATE cis.Ws_CLMMedCharges SET MedCharges1 = @c1, MedCharges2 = @c2, MedCharges3 = @c3, MedCharges4 = @c4, MedCharges5 = @c5, MedCharges6 = @c6,
        MedCharges7 = @c7 WHERE CASNumber = @casNumber;
END
GO

-- UpdateRecord / UpdateEmptyOPbatchno.
CREATE OR ALTER PROCEDURE cis.InvAdj_Batch @casNumber nvarchar(20), @opBatchNo nvarchar(30) = NULL, @userCode nvarchar(50), @today datetime AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCas SET OPBatchNo = @opBatchNo, OPNoEnterUser = CASE WHEN @opBatchNo IS NULL THEN NULL ELSE @userCode END,
        OPBatchDate = CASE WHEN @opBatchNo IS NULL THEN NULL ELSE @today END
    WHERE CASNumber = @casNumber;
END
GO

-- DiagnosisUpdate (a case posted to SUN: only the diagnoses and the remark).
CREATE OR ALTER PROCEDURE cis.InvAdj_Diagnosis @casNumber nvarchar(20), @d1 nvarchar(20), @d2 nvarchar(20), @d3 nvarchar(20), @d4 nvarchar(20),
    @d5 nvarchar(20), @other nvarchar(200), @remark nvarchar(500)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET CASTreatReason1 = @d1, CASTreatReason2 = @d2, CASTreatReason3 = @d3, CASTreatReason4 = @d4, CASTreatReason5 = @d5,
        CASDiagnosis5 = @other, CasRemark = @remark
    WHERE CASNumber = @casNumber;
END
GO

-- UpdateMBMBalance of the adjustment: undo the previous approval, apply the new one (row: principal, cover 01 for shared covers B / D, or the cover).
CREATE OR ALTER PROCEDURE cis.InvAdj_Balance @mbmNumber nvarchar(50), @limitCover nvarchar(2) = NULL, @adjDelta money, @accumDelta money,
    @balanceDelta money = NULL
AS
BEGIN
    SET NOCOUNT ON;
    IF @limitCover IS NULL
        UPDATE cis.Hdb_MBMCLNOthers SET MBMAdjAnnLmt = COALESCE(MBMAdjAnnLmt, 0) + @adjDelta, MBMAccumAmt = COALESCE(MBMAccumAmt, 0) + @accumDelta,
            MBMCLNBalLmt = CASE WHEN @balanceDelta IS NULL THEN MBMCLNBalLmt WHEN COALESCE(MBMCLNBalLmt, 0) + @balanceDelta < 0 THEN 0
                                ELSE COALESCE(MBMCLNBalLmt, 0) + @balanceDelta END
        WHERE MBMNumber = @mbmNumber;
    ELSE
        UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCAdjAnnLmt = COALESCE(TRY_CONVERT(money, MBMCAdjAnnLmt), 0) + @adjDelta,
            MBMCAccumAmt = COALESCE(TRY_CONVERT(money, MBMCAccumAmt), 0) + @accumDelta,
            MBMCCLNBalLmt = CASE WHEN @balanceDelta IS NULL THEN MBMCCLNBalLmt WHEN COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) + @balanceDelta < 0 THEN 0
                                 ELSE COALESCE(TRY_CONVERT(money, MBMCCLNBalLmt), 0) + @balanceDelta END
        WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @limitCover;
END
GO

-- cmdauditreviewupdate_Click.
CREATE OR ALTER PROCEDURE cis.InvAdj_Audit @casNumber nvarchar(20), @query nvarchar(max), @answer nvarchar(max), @discount nvarchar(50),
    @remarks nvarchar(max), @auditBy nvarchar(100), @auditDate datetime
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCas SET AuditQuery = @query, AuditAnswer = @answer, AuditDiscountReceived = @discount, AuditRemarks = @remarks, AuditBy = @auditBy,
        AuditDate = @auditDate
    WHERE CASNumber = @casNumber;
END
GO

-- cmdCLMBordxApp_Click.
CREATE OR ALTER PROCEDURE cis.InvAdj_ClaimsApp @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET CASClaimsApp = 'Y' WHERE CASNumber = @casNumber;
END
GO

-- cmdvoidbordx_Click: a "KIV" bordereaux number (Claim_BordxSEQ of pay code IV), the invoice and claim marked bordered, a BORDX row and a VOID
-- member payment.
CREATE OR ALTER PROCEDURE cis.InvAdj_VoidBordx @casNumber nvarchar(20), @groupCompany nvarchar(100), @userCode nvarchar(50), @today datetime AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @n int = (SELECT TOP (1) CAST(CLMBordxNo AS int) FROM cis.Ws_Claim_BordxSEQ WITH (UPDLOCK, HOLDLOCK) WHERE PAYCode = 'IV');
    IF @n IS NULL
    BEGIN
        SET @n = 1;
        INSERT INTO cis.Ws_Claim_BordxSEQ (HLTCode, PAYCode, CLMBordxNo) VALUES ('', 'IV', '00001');
    END
    ELSE
    BEGIN
        SET @n += 1;
        UPDATE cis.Ws_Claim_BordxSEQ SET CLMBordxNo = RIGHT('00000' + CAST(@n AS varchar(10)), 5) WHERE PAYCode = 'IV';
    END
    DECLARE @no nvarchar(20) = 'KIV' + RIGHT('00000' + CAST(@n AS varchar(10)), 5);
    UPDATE cis.Ws_ClaimInvoice SET INVBordxed = 1, INVBordxRefNo = @no WHERE CASNumber = @casNumber;
    UPDATE cis.Ws_Claim SET BordxRefNo = @no, ClaimBordxStatus = 1 WHERE CASNumber = @casNumber;
    INSERT INTO cis.Ws_Bordx (BordxRefNo, BordxPAYCode, BordxDate, BordxBy, GroupCompany) VALUES (@no, 'IV', @today, @userCode, @groupCompany);
    INSERT INTO cis.Pay_Payment2MBM (PVNo, PVDate, PVChequeDate, PVChequeNo, CASNumber, Pay2MBMLastUpdateDate, Pay2MBMLastUpdateUser, PAYAccPeriodMth, PAYAccPeriodYr)
    VALUES ('VOID', @today, @today, 'VOID', @casNumber, @today, @userCode, MONTH(@today), YEAR(@today));
    SELECT @no;
END
GO

-- CheckRejectionPending: the desktop's database e-mail notification (1 = sent).
CREATE OR ALTER PROCEDURE cis.InvAdj_Notify @casNumber nvarchar(20), @notes varchar(2000), @pendingReasons varchar(2000), @email varchar(2000) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_EmailNotifyPendingReject_LatestModified @Casnumber = @casNumber, @tempReason = @notes, @PendingReason = @pendingReasons, @tempemail = @email;
END
GO

-- ===== main/530_invoice_adjustment_after_post.sql =====
-- Invoice > Clinic > Adjustment (After Posting) (frmInvoiceAdjAfterPost): a case already posted to SUN is not changed; a sub-case is created
-- (sp_CreateSubCasNum / sp_CreateSubCasNumClaim with the master case number). Its GL amount follows SaveINVClaim.
CREATE OR ALTER PROCEDURE cis.InvAdj_LgAmount @casNumber nvarchar(20), @glAmount money AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.LG SET GLAmt = @glAmount WHERE LGNo = @casNumber;
END
GO

-- ===== main/600_audit_review.sql =====
-- LG & Audit > Audit review (frmAuditInvoice): one case and its audit details. The update reuses cis.InvAdj_Audit, Prev / Next reuse
-- cis.InvEnq_Step.

-- CmdGo_Click: SP_CISAuditdetails (already parameterised) and the diagnosis descriptions (DiagDesc).
CREATE OR ALTER PROCEDURE cis.Audit_Case @casNumber varchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.SP_CISAuditdetails @CASNumber = @casNumber;
END
GO

-- ===== main/610_audit_tracking.sql =====
-- LG & Audit > Audit Tracking (FrmAuditInvoiceTracking): Sp_Audittraccking_2018's report with the desktop's filters. The desktop pasted the
-- filter text into the procedure; here every filter is a parameter. The filter clauses are added as text (as the desktop did, so unqualified
-- columns resolve the same way) but all values go through sp_executesql parameters.
-- Reproduced quirks: the claimability filter and the FFS "N" filter add an unbracketed OR; cases without any medication are not listed; the
-- fixed condition CASStatus NOT IN ('D','C') means the C / D status filter finds nothing.
CREATE OR ALTER PROCEDURE cis.AuditTrack_Search
    @caseFrom nvarchar(20) = NULL, @caseTo nvarchar(20) = NULL, @caseBlank bit = 0,
    @payTo nvarchar(2) = NULL, @payToBlank bit = 0, @clinic nvarchar(20) = NULL, @clinicBlank bit = 0,
    @payor nvarchar(10) = NULL, @payorBlank bit = 0, @groupCompany nvarchar(200) = NULL, @groupBlank bit = 0,
    @patientName nvarchar(200) = NULL, @nameBlank bit = 0, @mbmFrom nvarchar(50) = NULL, @mbmTo nvarchar(50) = NULL, @mbmBlank bit = 0,
    @invFrom nvarchar(50) = NULL, @invTo nvarchar(50) = NULL, @invBlank bit = 0,
    @bordxFrom nvarchar(30) = NULL, @bordxTo nvarchar(30) = NULL, @bordxBlank bit = 0,
    @caseStatus char(1) = NULL, @claimability char(1) = NULL,
    @planFrom nvarchar(20) = NULL, @planTo nvarchar(20) = NULL, @planBlank bit = 0,
    @bordxDateFrom datetime = NULL, @bordxDateTo datetime = NULL,
    @icFrom nvarchar(50) = NULL, @icTo nvarchar(50) = NULL, @icBlank bit = 0,
    @visitFrom datetime = NULL, @visitTo datetime = NULL, @visitBlank bit = 0,
    @regFrom datetime = NULL, @regTo datetime = NULL, @regBlank bit = 0,
    @invDateFrom datetime = NULL, @invDateTo datetime = NULL, @invDateBlank bit = 0,
    @enteredFrom datetime = NULL, @enteredTo datetime = NULL, @enteredBlank bit = 0,
    @amtFrom money = NULL, @amtTo money = NULL, @amtBlank bit = 0,
    @monthFrom nvarchar(4) = NULL, @monthTo nvarchar(4) = NULL, @monthBlank bit = 0,
    @yearFrom nvarchar(4) = NULL, @yearTo nvarchar(4) = NULL, @yearBlank bit = 0,
    @clinicType char(1) = NULL, @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @opBlank bit = 0,
    @ffs char(1) = NULL, @omni bit = 0, @tims bit = 0, @subsidiary nvarchar(200) = NULL, @department nvarchar(200) = NULL,
    @omniRefFrom nvarchar(50) = NULL, @omniRefTo nvarchar(50) = NULL, @omniRefBlank bit = 0, @omniSent char(1) = NULL, @audited bit = 0
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @w nvarchar(max) = N'';
    IF @caseFrom IS NOT NULL SET @w += N' AND ciscas.CASNumber >= @caseFrom';
    IF @caseTo IS NOT NULL SET @w += N' AND ciscas.CASNumber <= @caseTo';
    IF @caseBlank = 1 SET @w += N' AND (ciscas.CASNumber Is null OR ciscas.CASNumber = '''')';
    IF @payTo IS NOT NULL SET @w += N' AND INVPayTo = @payTo';
    IF @payToBlank = 1 SET @w += N' AND (INVPayTo Is null OR INVPayTo = '''')';
    IF @clinic IS NOT NULL SET @w += N' AND INVPayee = @clinic';
    IF @clinicBlank = 1 SET @w += N' AND (INVPayee Is null OR INVPayee = '''')';
    IF @payor IS NOT NULL SET @w += N' AND MBMCLNOthers.PAYCode = @payor';
    IF @payorBlank = 1 SET @w += N' AND (MBMCLNOthers.PAYCode Is null OR MBMCLNOthers.PAYCode = '''')';
    IF @groupCompany IS NOT NULL SET @w += N' AND MBMGroupCompany = @groupCompany';
    IF @groupBlank = 1 SET @w += N' AND (MBMGroupCompany IS NULL OR MBMGroupCompany = '''')';
    IF @patientName IS NOT NULL SET @w += N' AND mbm.mbmname = @patientName';
    IF @nameBlank = 1 SET @w += N' AND (mbm.mbmname Is null OR mbm.mbmname = '''')';
    IF @mbmFrom IS NOT NULL SET @w += N' AND CISCas.MBMNumber >= @mbmFrom';
    IF @mbmTo IS NOT NULL SET @w += N' AND CISCas.MBMNumber <= @mbmTo';
    IF @mbmBlank = 1 SET @w += N' AND (CISCas.MBMNumber Is null OR CISCas.MBMNumber = '''')';
    IF @invFrom IS NOT NULL SET @w += N' AND CISCas.InvoiceNo >= @invFrom';
    IF @invTo IS NOT NULL SET @w += N' AND CISCas.InvoiceNo <= @invTo';
    IF @invBlank = 1 SET @w += N' AND (CISCas.InvoiceNo Is null OR CISCas.InvoiceNo = '''')';
    IF @bordxFrom IS NOT NULL SET @w += N' AND Claim.BordxRefNo >= @bordxFrom';
    IF @bordxTo IS NOT NULL SET @w += N' AND Claim.BordxRefNo <= @bordxTo';
    IF @bordxBlank = 1 SET @w += N' AND (Claim.BordxRefNo Is null OR Claim.BordxRefNo = '''')';
    IF @caseStatus IN ('O', 'C', 'D') SET @w += N' AND CISCas.CASStatus = @caseStatus';
    IF @claimability IN ('A', 'R') SET @w += N' Or CISCas.Claimability = @claimability';
    IF @planFrom IS NOT NULL SET @w += N' AND MBMCLNOthers.PLNType >= @planFrom';
    IF @planTo IS NOT NULL SET @w += N' AND MBMCLNOthers.PLNType <= @planTo';
    IF @planBlank = 1 SET @w += N' AND (MBMCLNOthers.PLNType Is null OR MBMCLNOthers.PLNType = '''')';
    IF @bordxDateFrom IS NOT NULL SET @w += N' AND Bordx.BordxDate >= @bordxDateFrom';
    IF @bordxDateTo IS NOT NULL SET @w += N' AND Bordx.BordxDate <= @bordxDateTo';
    IF @icFrom IS NOT NULL SET @w += N' AND mbm.MBMIcBcPp >= @icFrom';
    IF @icTo IS NOT NULL SET @w += N' AND mbm.MBMIcBcPp <= @icTo';
    IF @icBlank = 1 SET @w += N' AND (mbm.MBMIcBcPp Is null)';
    IF @visitFrom IS NOT NULL SET @w += N' AND CISCas.DateVisit >= @visitFrom';
    IF @visitTo IS NOT NULL SET @w += N' AND CISCas.DateVisit <= @visitTo';
    IF @visitBlank = 1 SET @w += N' AND (CISCas.DateVisit Is null)';
    IF @regFrom IS NOT NULL SET @w += N' AND CISCas.RegisterDate >= @regFrom';
    IF @regTo IS NOT NULL SET @w += N' AND CISCas.RegisterDate <= @regTo';
    IF @regBlank = 1 SET @w += N' AND (CISCas.RegisterDate Is null)';
    IF @invDateFrom IS NOT NULL SET @w += N' AND CISCas.InvoiceDate >= @invDateFrom';
    IF @invDateTo IS NOT NULL SET @w += N' AND CISCas.InvoiceDate <= @invDateTo';
    IF @invDateBlank = 1 SET @w += N' AND (CISCas.InvoiceDate Is null)';
    IF @enteredFrom IS NOT NULL SET @w += N' AND ClaimInvoice.INVEnteredDate >= @enteredFrom';
    IF @enteredTo IS NOT NULL SET @w += N' AND ClaimInvoice.INVEnteredDate <= @enteredTo';
    IF @enteredBlank = 1 SET @w += N' AND (ClaimInvoice.INVEnteredDate Is null)';
    IF @amtFrom IS NOT NULL SET @w += N' AND ClaimInvoice.INVAmount >= @amtFrom';
    IF @amtTo IS NOT NULL SET @w += N' AND ClaimInvoice.INVAmount <= @amtTo';
    IF @amtBlank = 1 SET @w += N' AND (ClaimInvoice.INVAmount Is null)';
    IF @monthFrom IS NOT NULL SET @w += N' AND AccPeriodMonth >= @monthFrom';
    IF @monthTo IS NOT NULL SET @w += N' AND AccPeriodMonth <= @monthTo';
    IF @monthBlank = 1 SET @w += N' AND AccPeriodMonth Is null';
    IF @yearFrom IS NOT NULL SET @w += N' AND AccPeriodYear >= @yearFrom';
    IF @yearTo IS NOT NULL SET @w += N' AND AccPeriodYear <= @yearTo';
    IF @yearBlank = 1 SET @w += N' AND AccPeriodYear Is null';
    IF @clinicType = '0' SET @w += N' AND CISCas.CLNType = ''S''';
    IF @clinicType = '1' SET @w += N' AND CISCas.CLNType = ''G''';
    IF @clinicType = '2' SET @w += N' AND (CISCas.CLNType = '''' OR CISCas.CLNType IS NULL)';
    IF @opFrom IS NOT NULL SET @w += N' AND OPBatchNo >= @opFrom';
    IF @opTo IS NOT NULL SET @w += N' AND OPBatchNo <= @opTo';
    IF @opBlank = 1 SET @w += N' AND (OPBatchNo Is null OR OPBatchNo = '''')';
    IF @ffs = 'Y' SET @w += N' AND CLNFFSStatus = ''Y''';
    IF @ffs = 'N' SET @w += N' AND CLNFFSStatus = ''N'' OR (CLNFFSStatus = '''' OR CLNFFSStatus IS NULL)';
    IF @omni = 1 SET @w += N' AND GRPOMNICLM = ''Y''';
    IF @tims = 1 SET @w += N' AND GRPTIMS = ''Y''';
    IF @subsidiary IS NOT NULL SET @w += N' AND MBMOthers.Subsidiary = @subsidiary';
    IF @department IS NOT NULL SET @w += N' AND MBMOthers.MBMDepartment = @department';
    IF @omniRefBlank = 1 SET @w += N' AND (OMNIRefNo Is null OR OMNIRefNo = '''')';
    IF @omniRefFrom IS NOT NULL SET @w += N' AND OMNIRefNo >= @omniRefFrom';
    IF @omniRefTo IS NOT NULL SET @w += N' AND OMNIRefNo <= @omniRefTo';
    IF @omniSent = 'Y' SET @w += N' AND CLMOMNISent = ''Y''';
    IF @omniSent = 'N' SET @w += N' AND CLMOMNISent IS NULL';
    IF @audited = 1 SET @w += N' AND ciscas.AuditQuery IS NOT NULL';

    DECLARE @sql nvarchar(max) = N'
SELECT CASE WHEN MBMCoveredPersons.MBMCICBCPPNo IS NULL THEN mbm.MBMIcBcPp ELSE MBMCoveredPersons.MBMCICBCPPNo END AS MBMIcBcPp,
    CASE WHEN MBMCoveredPersons.MBMCName IS NULL THEN mbm.MBMName ELSE MBMCoveredPersons.MBMCName END AS MBMName,
    ciscas.CASNumber, ciscas.DateVisit,
    CASE WHEN CISCas.CLNType = ''S'' THEN COALESCE(ClaimInvoiceActualOthers.SPFee, 0) ELSE COALESCE(ClaimInvoiceActualOthers.GPFee, 0) END AS SPFee,
    COALESCE(ClaimInvoice.INVAmount, 0) AS INVAmount, LTRIM(RTRIM(CISCas.CasRemark)) AS CasRemark, LTRIM(RTRIM(CISCas.AuditQuery)) AS AuditQuery,
    LTRIM(RTRIM(CISCas.AuditAnswer)) AS AuditAnswer, LTRIM(RTRIM(CISCas.AuditRemarks)) AS AuditRemarks, CISCas.AuditDiscountReceived,
    LTRIM(RTRIM(CISCas.AuditBy)) AS AuditBy, CISCas.AuditDate, CLN.CLNCode, LTRIM(RTRIM(CLN.CLNName)) AS CLNName,
    UPPER(LTRIM(RTRIM(MBMOthers.MBMGroupCompany))) AS MBMGroupCompany,
    CASE CASTreatReason1 WHEN '''' THEN '''' ELSE (SELECT ReasonDesc FROM dbo.TreatmentReason WHERE ReasonCode = CASTreatReason1) END AS ReasonDesc1,
    CASE CASTreatReason2 WHEN '''' THEN '''' ELSE (SELECT ReasonDesc FROM dbo.TreatmentReason WHERE ReasonCode = CASTreatReason2) END AS ReasonDesc2,
    CASE CASTreatReason3 WHEN '''' THEN '''' ELSE (SELECT ReasonDesc FROM dbo.TreatmentReason WHERE ReasonCode = CASTreatReason3) END AS ReasonDesc3,
    CASE CASTreatReason4 WHEN '''' THEN '''' ELSE (SELECT ReasonDesc FROM dbo.TreatmentReason WHERE ReasonCode = CASTreatReason4) END AS ReasonDesc4,
    CASE CASTreatReason5 WHEN '''' THEN '''' ELSE (SELECT ReasonDesc FROM dbo.TreatmentReason WHERE ReasonCode = CASTreatReason5) END AS ReasonDesc5,
    CASDiagnosis5,
    ciscas.MedInjType1, MedInjDays1, MedInjQty1, CLMMedCharges.MedCharges1, ciscas.MedInjType2, MedInjDays2, MedInjQty2, CLMMedCharges.MedCharges2,
    ciscas.MedInjType3, MedInjDays3, MedInjQty3, CLMMedCharges.MedCharges3, ciscas.MedInjType4, MedInjDays4, MedInjQty4, CLMMedCharges.MedCharges4,
    ciscas.MedInjType5, MedInjDays5, MedInjQty5, CLMMedCharges.MedCharges5, ciscas.MedInjType6, MedInjDays6, MedInjQty6, CLMMedCharges.MedCharges6,
    ciscas.MedInjType7, MedInjDays7, MedInjQty7, CLMMedCharges.MedCharges7,
    Bordx.BordxRefNo,
    CASE WHEN CISCas.CLNType = ''S'' THEN COALESCE(ClaimInvoiceActualOthers.SPInj, 0) ELSE COALESCE(ClaimInvoiceActualOthers.GPInj, 0) END AS Procedurefee
INTO #p
FROM dbo.CISCas
INNER JOIN cis.Hdb_MBMCLNOthers MBMCLNOthers WITH (NOLOCK) ON MBMCLNOthers.MBMNumber = CISCas.MBMnumber
INNER JOIN cis.Hdb_MBMOthers MBMOthers WITH (NOLOCK) ON MBMOthers.MBMNumber = MBMCLNOthers.MBMNumber
INNER JOIN cis.Hdb_MBM MBM WITH (NOLOCK) ON MBM.MBMNumber = MBMOthers.MBMNumber
INNER JOIN cis.Ws_Claim Claim WITH (NOLOCK) ON Claim.CASNumber = CISCas.CASNumber
LEFT OUTER JOIN cis.Ws_CLMMedCharges CLMMedCharges ON CLMMedCharges.CASNumber = ciscas.CASNumber
LEFT OUTER JOIN cis.Hdb_MBMCoveredPersons MBMCoveredPersons WITH (NOLOCK) ON MBMCoveredPersons.MBMCNumber = ciscas.MBMnumber AND MBMCoveredPersons.MBMCCoverID = ciscas.CoverID
LEFT OUTER JOIN cis.Hdb_CLN CLN WITH (NOLOCK) ON CISCas.ClinicCode = CLN.CLNCode
LEFT OUTER JOIN dbo.GRPCompany GRPCompany WITH (NOLOCK) ON GRPCompany.GRPCoName = MBMOthers.MBMGroupCompany AND (GRPCompany.hispaycode IS NULL OR GRPCompany.hispaycode = mbm.paycode)
LEFT OUTER JOIN cis.Ws_ClaimInvoice ClaimInvoice WITH (NOLOCK) ON ClaimInvoice.CASNumber = Claim.CASNumber
LEFT OUTER JOIN cis.Ws_ClaimInvoiceActualOthers ClaimInvoiceActualOthers WITH (NOLOCK) ON ClaimInvoiceActualOthers.CASNumber = Claim.CASNumber
LEFT OUTER JOIN cis.Ws_ClaimInvoiceActualNew ClaimInvoiceActualNew WITH (NOLOCK) ON ClaimInvoiceActualNew.CASNumber = Claim.CASNumber
LEFT OUTER JOIN cis.Ws_Bordx Bordx WITH (NOLOCK) ON Bordx.BordxRefNo = Claim.BordxRefNo
LEFT OUTER JOIN dbo.tblRecord tblRecord WITH (NOLOCK) ON tblRecord.CaseNo = Claim.CASNumber
LEFT OUTER JOIN cis.Pay_Payment2CLN Payment2CLN WITH (NOLOCK) ON Payment2CLN.CASNumber = Claim.CASNumber
LEFT JOIN dbo.LG LG WITH (NOLOCK) ON LG.LGNo = CISCAS.CASNumber
INNER JOIN dbo.PLN PLN WITH (NOLOCK) ON PLN.PLNCode = MBMCLNOthers.PLNType
WHERE (MBMCLNOthers.PAYCode <> N''DZ'') AND (MBMCLNOthers.PAYCode <> N''EM'') AND (MBMCLNOthers.PAYCode <> N''ES'') AND (CISCas.CASStatus NOT IN (''D'', ''C''))'
        + @w + N'
OPTION (RECOMPILE);
WITH cte AS (
    SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType1 AS MedInjType, MedInjDays1 AS MedInjDays, MedInjQty1 AS MedInjQty, MedCharges1 AS MedCharges, 1 AS MedNo,
        SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer, AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5,
        CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName, MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType1)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType2, MedInjDays2, MedInjQty2, MedCharges2, 2, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType2)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType3, MedInjDays3, MedInjQty3, MedCharges3, 3, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType3)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType4, MedInjDays4, MedInjQty4, MedCharges4, 4, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType4)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType5, MedInjDays5, MedInjQty5, MedCharges5, 5, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType5)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType6, MedInjDays6, MedInjQty6, MedCharges6, 6, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType6)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType7, MedInjDays7, MedInjQty7, MedCharges7, 7, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType7)), '''') <> ''''
)
SELECT COUNT(*) OVER (PARTITION BY CASNumber) AS MedCount, ROW_NUMBER() OVER (PARTITION BY CASNumber ORDER BY MedNo) AS RNo, *
FROM cte ORDER BY CASNumber, MedNo;
DROP TABLE #p;';

    EXEC sp_executesql @sql, N'@caseFrom nvarchar(20), @caseTo nvarchar(20), @payTo nvarchar(2), @clinic nvarchar(20), @payor nvarchar(10),
        @groupCompany nvarchar(200), @patientName nvarchar(200), @mbmFrom nvarchar(50), @mbmTo nvarchar(50), @invFrom nvarchar(50), @invTo nvarchar(50),
        @bordxFrom nvarchar(30), @bordxTo nvarchar(30), @caseStatus char(1), @claimability char(1), @planFrom nvarchar(20), @planTo nvarchar(20),
        @bordxDateFrom datetime, @bordxDateTo datetime, @icFrom nvarchar(50), @icTo nvarchar(50), @visitFrom datetime, @visitTo datetime,
        @regFrom datetime, @regTo datetime, @invDateFrom datetime, @invDateTo datetime, @enteredFrom datetime, @enteredTo datetime,
        @amtFrom money, @amtTo money, @monthFrom nvarchar(4), @monthTo nvarchar(4), @yearFrom nvarchar(4), @yearTo nvarchar(4),
        @opFrom nvarchar(30), @opTo nvarchar(30), @subsidiary nvarchar(200), @department nvarchar(200), @omniRefFrom nvarchar(50), @omniRefTo nvarchar(50)',
        @caseFrom, @caseTo, @payTo, @clinic, @payor, @groupCompany, @patientName, @mbmFrom, @mbmTo, @invFrom, @invTo, @bordxFrom, @bordxTo,
        @caseStatus, @claimability, @planFrom, @planTo, @bordxDateFrom, @bordxDateTo, @icFrom, @icTo, @visitFrom, @visitTo, @regFrom, @regTo,
        @invDateFrom, @invDateTo, @enteredFrom, @enteredTo, @amtFrom, @amtTo, @monthFrom, @monthTo, @yearFrom, @yearTo, @opFrom, @opTo,
        @subsidiary, @department, @omniRefFrom, @omniRefTo;
END
GO

-- CboGrpCom_LostFocus / ChkPayor2_Click: subsidiaries and departments of a group company, payors.
CREATE OR ALTER PROCEDURE cis.AuditTrack_Lookups @groupCompany nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT DISTINCT LTRIM(RTRIM(Subsidiary)) AS Subsidiary FROM cis.Hdb_MBMOthers
    WHERE @groupCompany IS NOT NULL AND MBMGroupCompany = @groupCompany AND COALESCE(LTRIM(RTRIM(Subsidiary)), '') <> '' ORDER BY 1;
    SELECT DISTINCT LTRIM(RTRIM(MBMDepartment)) AS Department FROM cis.Hdb_MBMOthers
    WHERE @groupCompany IS NOT NULL AND MBMGroupCompany = @groupCompany AND COALESCE(LTRIM(RTRIM(MBMDepartment)), '') <> '' ORDER BY 1;
END
GO

-- ===== main/620_lg_generation.sql =====
-- LG & Audit > Generate LG (frmLGGeneration): read side. Member search, limits, the member's LGs, one LG, the case tabs and the drop-downs.
SET QUOTED_IDENTIFIER ON;
GO

-- SearchRecords: CLNLG1 (one row per cover) and, for the listed companies, the member's warranty text.
CREATE OR ALTER PROCEDURE cis.Lg_Search @member nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Hdb_CLNLG1 @StrAppend = @member;
    SELECT DISTINCT TOP 1 MBMCLNWarranty FROM cis.Hdb_MBMCLNOthers WHERE MBMNumber LIKE @member + N'%';
END
GO

-- CheckMBMLimit / ValidateLG inputs: GetLGAnnLimitSpecLimit, SP_ValidateLG1, the specialist top-ups, AnnualLimit.SuppLimitStatus and the plan flags.
CREATE OR ALTER PROCEDURE cis.Lg_Member @member nvarchar(30), @cover nvarchar(2), @ic nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @lim TABLE (
        SchMBMNumber nvarchar(50), CoverId nvarchar(10), PolEffDate nvarchar(20), PolExpDate nvarchar(20), PLNCode nvarchar(20), HLTCode nvarchar(10),
        PAYCode nvarchar(10), INSType nvarchar(10), ANNLmtInd nvarchar(5), AnnLimit nvarchar(30), SuppLimit nvarchar(30), SuppLimitStatus nvarchar(5),
        SpecLmtInd nvarchar(5), SpecAnnLimit nvarchar(30), SpecSuppLimit nvarchar(30), SpecSuppLimitStatus nvarchar(5), TopupAmt nvarchar(30),
        GBordxAmt nvarchar(30), GFFS nvarchar(30), GNoLimitDed1 nvarchar(30), GNoLimitDed2 nvarchar(30), GNoLimitDedTotal nvarchar(30), GUtilized nvarchar(30),
        GBalance nvarchar(30), SBordxAmt nvarchar(30), SFFS nvarchar(30), SNoLimitDed1 nvarchar(30), SNoLimitDed1b nvarchar(30), SNoLimitDedTotal nvarchar(30),
        SUtilized nvarchar(30), SBalance nvarchar(30), Error nvarchar(500), OriSpecAnnLimit nvarchar(30), OriAnnLimit nvarchar(30), LGIssuedInd nvarchar(5),
        LGIssuedSuppInd nvarchar(5), LGIssuedAfterLimitExInd nvarchar(5), LGUsingLimit nvarchar(5), LGUsingLimitDesc nvarchar(50), LGHosTypeTInd nvarchar(5),
        PLNGLNo nvarchar(50), HRApproval nvarchar(5));
    INSERT @lim EXEC cis.Ws_GetLGAnnLimitSpecLimit @SchMBMNumber = @member, @CoverId = @cover;
    SELECT * FROM @lim;
    -- SP_ValidateLG1 returns no result set at all for some inputs (e.g. a blank IC): capture it so the result sets stay in order.
    DECLARE @v TABLE (PLNGPSPShared nvarchar(5), LGIssuedInd nvarchar(5), ANNLmtInd nvarchar(5), AnnLimit nvarchar(30), PlnCode nvarchar(20),
        MBMStatus nvarchar(5), LGIssuedAfterLimitExInd nvarchar(5));
    INSERT @v EXEC cis.Hdb_SP_ValidateLG1 @ICNo = @ic, @MBMNumber = @member;
    SELECT * FROM @v;
    SELECT SUM(COALESCE(SP_TopupAmt, 0)) AS SpTopup FROM dbo.MBMCISTopupHistory WHERE MBMNumber = @member;
    DECLARE @plan nvarchar(20) = (SELECT TOP 1 RTRIM(PLNCode) FROM @lim);
    SELECT TOP 1 SuppLimitStatus FROM dbo.AnnualLimit WHERE PLNCode = @plan;
    SELECT TOP 1 PLNMaternityInd, PDentalStatus, COALESCE(IsPrinDepGPShared, 0) AS IsPrinDepGPShared FROM dbo.PLN WHERE PLNCode = @plan;
    SELECT MBMASOLimit, MBMOPLimit, MBMIPUtilized, MBMOPUtilized, IsIPOPCombined FROM cis.Hdb_ASOMemberDetails WHERE MBMNumber = @member AND MBMCoverID = @cover;
END
GO

-- spMERCHANTRADE (plan M230 of MERCHANTRADE ASIA SDN BHD): this month's balance.
CREATE OR ALTER PROCEDURE cis.Lg_Merchantrade @member nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.spMERCHANTRADE @mbmno = @member;
END
GO

-- LGListing: LGlst2020 with its filter built from the parameters (the desktop pasted it).
CREATE OR ALTER PROCEDURE cis.Lg_List @member nvarchar(30), @cover nvarchar(2) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @f varchar(2000) = ' AND LG.MBMNumber = ''' + REPLACE(@member, '''', '''''') + ''''
        + CASE WHEN @cover IS NULL THEN '' ELSE ' AND LG.CoverID = ''' + REPLACE(@cover, '''', '''''') + '''' END;
    EXEC dbo.LGlst2020 @StrAppend = @f;
END
GO

-- lstLG_Click: View_PrintLG, pending, pharmacy (clinic type D), the master case of a follow-up, PEMC offline GL data.
CREATE OR ALTER PROCEDURE cis.Lg_Load @lgNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM dbo.View_PrintLG WHERE LGNo = @lgNo;
    SELECT Pending FROM dbo.CISCas WHERE CASNumber = @lgNo;
    SELECT COUNT(*) AS Pharmacy FROM dbo.LG INNER JOIN cis.Hdb_CLN CLN ON CLN.CLNCode = LG.CLNName WHERE LGNo = @lgNo AND CLN.CLNType = 'D';
    SELECT MainCaseNo FROM dbo.tblMedRequestFollowup WHERE FollowupCaseNo = @lgNo;
    SELECT PatientName, IcNo FROM dbo.tblOfflineGL WHERE CasNo = @lgNo;
END
GO

-- IsMedRequestInProgress.
CREATE OR ALTER PROCEDURE cis.Lg_MedRequestInProgress @caseNo nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @r bit = 0;
    EXEC dbo.sp_HCMedRequestInProgress @CaseNo = @caseNo, @RetOut = @r OUTPUT;
    SELECT @r AS InProgress;
END
GO

-- ValidateLG1: the clinic's GL format / address and Sp_Gettemplatedetails_2 (template, dependant template, remarks).
CREATE OR ALTER PROCEDURE cis.Lg_Template @clinic nvarchar(20), @company nvarchar(200), @plan nvarchar(20), @glType nvarchar(100), @ic nvarchar(50),
    @category nvarchar(20), @caseNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @format nvarchar(5) = (SELECT TOP 1 RTRIM(CLNGLFormat) FROM cis.Hdb_CLN WHERE CLNCode = @clinic);
    SELECT TOP 1 CLNGLFormat, COALESCE(CLNName, '') AS CLNName, CLNAdd1, COALESCE(CLNAdd2, '') + ',' + COALESCE(CLNAdd3, '') AS CLNAdd2,
        COALESCE(CLNState, '') + ',' + COALESCE(CLNPostCode, '') AS CLNState, CLNPanel
    FROM cis.Hdb_CLN WHERE CLNCode = @clinic;
    DECLARE @f varchar(10) = CASE WHEN @format = 'M' THEN 'M' ELSE 'H' END;
    EXEC dbo.Sp_Gettemplatedetails_2 @CompanyName = @company, @plncode = @plan, @GLFormat = @f, @GLType = @glType, @IC = @ic, @PEMC = @category, @caseno = @caseNo;
END
GO

-- ListCases / ListSpecialistCases.
CREATE OR ALTER PROCEDURE cis.Lg_Cases @member nvarchar(30), @cover nvarchar(2) = NULL, @specialist bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    IF @specialist = 0
        SELECT CASNumber, MBMCName, ReasonDesc, CoverID, CLMMCOProcFee, DateVisit, MBMCRELCode, CLMTotalActual, CLMBordxAmt, CASStatus, MBMName, MBMNumber, Claimability
        FROM dbo.VIEW_CISMBMInquiry_GPCASE WHERE MBMNumber = @member AND (@cover IS NULL OR CoverID = @cover);
    ELSE
        SELECT CASNumber, MBMCName, ReasonDesc, CoverID, CLMMCOProcFee, DateVisit, MBMCRELCode, CLMTotalActual, CLMBordxAmt, CASStatus, MBMName, MBMNumber, Claimability
        FROM dbo.VIEW_CISMBMInquiry_Spe_CASE WHERE MBMNumber = @member AND (@cover IS NULL OR CoverID = @cover);
END
GO

-- Form_Load: BindGLType, BindDecline, BindDelayreasons.
CREATE OR ALTER PROCEDURE cis.Lg_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT GLID, GLType, GLDesc FROM dbo.Tbl_GLType;
    SELECT ID, DeclineReason FROM dbo.Tbl_DeclineReason WHERE COALESCE(DeclineReason, '') <> '';
    SELECT MessageContent FROM cis.Hdb_MessageComboList
    WHERE MessageModule = 'LG Generation' AND MessageForm = 'Generate Guarantee Letter-CIS' AND MessageTag = 'Reason for delay';
END
GO

-- chkCLN_Click: the group company's clinics.
CREATE OR ALTER PROCEDURE cis.Lg_Clinics @company nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.SearchGrpComCLN @company;
END
GO

-- ===== main/630_lg_generation_save.sql =====
-- LG & Audit > Generate LG (frmLGGeneration): the checks of Edit / LG, issuing (SaveLG), declining (cmdDecline) and the LG e-mail recipients.
SET QUOTED_IDENTIFIER ON;
GO

-- CmdEdit_Click / cmdLG_Click checks: bordereaux / invoice, payments, spLGforSP_check, spclnforpanel_check (by the clinic text, as the desktop),
-- statusInfo (expiry within 30 days of today), pending, and the LG's case date / time.
CREATE OR ALTER PROCEDURE cis.Lg_Checks @lgNo nvarchar(20), @member nvarchar(30), @cover nvarchar(2), @clinicCode nvarchar(20), @clinicText nvarchar(300) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(COALESCE(c.ClaimBordxStatus, 0) AS bit) AS Bordx, LTRIM(RTRIM(COALESCE(cas.InvoiceNo, ''))) AS InvoiceNo,
        CAST(COALESCE(c.BTBGBordxStatus, 0) AS bit) AS BtbgBordx
    FROM cis.Ws_Claim c INNER JOIN dbo.CISCas cas ON cas.CASNumber = c.CASNumber WHERE c.CASNumber = @lgNo;
    SELECT (SELECT COUNT(*) FROM cis.Pay_Payment2CLN WHERE CASNumber = @lgNo) + (SELECT COUNT(*) FROM cis.Pay_Payment2MBM WHERE CASNumber = @lgNo) AS Payments;
    DECLARE @sp TABLE (Result nvarchar(20));
    INSERT @sp EXEC dbo.spLGforSP_check @mbmno = @member, @clncode = @clinicCode;
    SELECT TOP 1 Result FROM @sp;
    SELECT CLNPanel FROM cis.Hdb_CLN WHERE LTRIM(RTRIM(CLNName)) + '[' + CONVERT(varchar, LTRIM(RTRIM(CLNCode))) + ']' = @clinicText;
    IF @cover = '00'
        SELECT TOP 1 o.MBMPolExpDate AS ExpiryDate FROM dbo.CISCas cas INNER JOIN cis.Hdb_MBMCLNOthers o ON o.MBMNumber = cas.MBMNumber
        WHERE o.MBMPolExpDate >= GETDATE() - 30 AND cas.CASNumber = @lgNo;
    ELSE
        SELECT TOP 1 p.MBMCCLNExpDate AS ExpiryDate FROM dbo.CISCas cas
        INNER JOIN cis.Hdb_MBMCoveredPersonsCLN p ON p.MBMCNumber = cas.MBMNumber AND p.MBMCCoverID = cas.CoverID
        WHERE p.MBMCCLNExpDate >= GETDATE() - 30 AND cas.CASNumber = @lgNo AND cas.CoverID = @cover;
    SELECT Pending FROM dbo.CISCas WHERE CASNumber = @lgNo;
END
GO

-- SaveLG (the update of the selected LG): LG, Claim, CISCas, TblRecord, ClaimInvoice and ClaimInvoiceActualNew in one transaction.
CREATE OR ALTER PROCEDURE cis.Lg_Issue @lgNo nvarchar(20), @user nvarchar(50), @treatmentDate datetime, @clinicCode nvarchar(20), @remarks nvarchar(max),
    @diagnosis nvarchar(max), @lgAmount money, @glAmount money, @internalRemarks nvarchar(max), @doctor nvarchar(200), @glType nvarchar(100),
    @maternity money, @pending bit, @delivery bit, @delayReason nvarchar(500) AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRAN;
    UPDATE dbo.LG SET LGDate = @treatmentDate, LGIssuedBy = @user, CLNName = @clinicCode, LGRemarks = @remarks, Diagnosis = @diagnosis, GLAmt = @lgAmount,
        InternalRemarks = @internalRemarks, DoctorName = @doctor, GLType = @glType, LGIssueDate = GETDATE()
    WHERE LGNo = @lgNo;
    UPDATE cis.Ws_Claim SET ClaimOpenDate = CAST(GETDATE() AS date), ClaimOpenBy = @user, ClaimCloseDate = CAST(GETDATE() AS date), ClaimClosedBy = @user,
        CLMTotalActual = @lgAmount, CLMFinalApprove = @lgAmount, CLMPaidBYMedix2C = @lgAmount, CLMTotalApproved = @lgAmount, CLMBordxAmt = @lgAmount
    WHERE ClaimBordxStatus = 0 AND CASNumber = @lgNo;
    IF @pending = 1
        UPDATE dbo.CISCas SET CasDelayReason = @delayReason, Pending = 'N', CASTreatReason1 = CASE WHEN @delivery = 1 THEN '302' END, PendingReminderDate = NULL,
            CASRepAction = CASRepAction + '|' + CONVERT(nvarchar, GETDATE(), 103) + ' ' + CONVERT(varchar, GETDATE(), 108) + '-', DateVisit = @treatmentDate,
            Claimability = 'A'
        WHERE CASNumber = @lgNo;
    ELSE
        UPDATE dbo.CISCas SET CasDelayReason = @delayReason, DateVisit = @treatmentDate, CASTreatReason1 = CASE WHEN @delivery = 1 THEN '302' END, Claimability = 'A'
        WHERE CASNumber = @lgNo;
    UPDATE dbo.tblRecord SET IsLGissued = 1, IsRejected = 0, CaseProgress = 'LG Issued' WHERE CaseNo = @lgNo;
    UPDATE cis.Ws_ClaimInvoice SET INVAmount = @glAmount, INVBillAmt = @glAmount WHERE InvBordxed = 0 AND CASNumber = @lgNo;
    UPDATE cis.Ws_ClaimInvoiceActualNew SET CLMMaternity = @maternity WHERE CASNumber = @lgNo;
    COMMIT;
END
GO

-- cmdDecline_Click: UpdateLG2 (its remarks parameter gets the decline reason, as the desktop passed it), the pharmacy request, TblRecord.
CREATE OR ALTER PROCEDURE cis.Lg_Decline @lgNo nvarchar(20), @doctor nvarchar(200), @clinicCode nvarchar(20), @reason varchar(500), @user nvarchar(50),
    @diagnosis nvarchar(max), @pharmacy bit AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRAN;
    DECLARE @now datetime = CAST(GETDATE() AS date);
    EXEC dbo.UpdateLG2 @LGNo = @lgNo, @LGDate = @now, @DoctorName = @doctor, @CLNName = @clinicCode, @LGRemarks = @reason, @LGUpdateBy = @user,
        @Diagnosis = @diagnosis, @DeclineReason = @reason;
    IF @pharmacy = 1 EXEC dbo.sp_UpdateMedReqDeclineReason @CaseNo = @lgNo, @DeclineReason = @reason;
    UPDATE dbo.tblRecord SET IsLGissued = 0, IsRejected = 1, CaseProgress = 'LG Declined' WHERE CaseNo = @lgNo;
    COMMIT;
END
GO

-- SaveLG e-mail: the group company's mail type (first row), the member's subsidiary / cost centre / department, and the first matching recipient row.
CREATE OR ALTER PROCEDURE cis.Lg_EmailRecipients @company nvarchar(200), @member nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @cnt int, @type nvarchar(5);
    SELECT TOP 1 @cnt = COUNT(*), @type = Mailtype FROM cis.Maint_CISEmailMaster
    WHERE COALESCE(Mailtype, '') <> 'E' AND COALESCE(IsCIS, '') = 'Y' AND Grpcompany = @company GROUP BY Mailtype;
    IF COALESCE(@cnt, 0) < 1 RETURN;
    DECLARE @sub nvarchar(200), @cost nvarchar(200), @dept nvarchar(200);
    SELECT TOP 1 @sub = COALESCE(Subsidiary, ''), @cost = COALESCE(MBMCostCenter, ''), @dept = COALESCE(MBMDepartment, '') FROM cis.Hdb_MBMOthers WHERE MBMNumber = @member;
    SELECT TOP 1 EmailAddress, BCC, CC, @sub AS Subsidiary FROM cis.Maint_CISEmailMaster
    WHERE LGInd = 1 AND IsCIS = 'Y' AND Grpcompany = @company
      AND (LTRIM(RTRIM(@type)) = 'G' OR (LTRIM(RTRIM(@type)) = 'S' AND Subsidary = @sub) OR (LTRIM(RTRIM(@type)) = 'D' AND Department = @dept)
           OR (LTRIM(RTRIM(@type)) = 'C' AND Costcenter = @cost));
END
GO

-- ===== main/700_mco_fee_listing.sql =====
-- MCO Fees > Fee Listing (frmMCOFeeListNew): View_Clinic_MCO_ListingNew by bordereaux batch / range / invoice no, payor or group company,
-- and the MCO invoice it prints (numbered from CLNMCOINVQ, stamped on MBMCLNOthers / MBMCoveredPersonsCLN).
SET QUOTED_IDENTIFIER ON;
GO

-- MCOFeeListing (MCOFeeNew case 4 with the pasted criteria replaced by parameters). @mode 0 batch, 1 range, 2 invoice no.
-- The report (CrystalRptName) uses the same filters, with the payor of the selected row: @columns = 1 returns every view column.
CREATE OR ALTER PROCEDURE cis.McoList_Search @mode int, @bordxDate datetime = NULL, @batchNo int = NULL, @from datetime = NULL, @to datetime = NULL,
    @invoiceNo nvarchar(30) = NULL, @invoiced char(1) = 'N', @payor nvarchar(10) = NULL, @groupCode nvarchar(10) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM dbo.View_Clinic_MCO_ListingNew
    WHERE (@mode <> 0 OR (MBMBordxDate = @bordxDate AND MBMBatchNo = @batchNo))
      AND (@mode <> 1 OR (MBMBordxDate >= @from AND MBMBordxDate <= @to))
      AND (@mode <> 2 OR MBMCInvoiceNo = @invoiceNo)
      AND (@mode = 2 OR MBMCInvoiced = @invoiced)
      AND (@payor IS NULL OR PAYCode = @payor)
      AND (@groupCode IS NULL OR GRPCoCode = @groupCode)
    ORDER BY MBMNumber
    OPTION (RECOMPILE);
END
GO

-- Form_Load drop-downs and PrintPAYInfo (the desktop always reads GRPCompany, see the screen doc).
CREATE OR ALTER PROCEDURE cis.McoList_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM dbo.PAY;
    SELECT RTRIM(GRPCoCode) AS Code, RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany;
END
GO

CREATE OR ALTER PROCEDURE cis.McoList_Company @code nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 GRPCoName AS CompanyName, GRPCoAdd1 AS Add1, GRPCoAdd2 AS Add2, GRPCoAdd3 AS Add3, GRPCoContact AS Contact, GRPCoPhone AS TelNo,
        GRPCoFax AS FaxNo
    FROM dbo.GRPCompany WHERE GRPCoCode = @code;
END
GO

-- IncInvoiceNo: the next number of CLNMCOINVQ ("00001" when the table is empty), locked for the save.
CREATE OR ALTER PROCEDURE cis.McoList_NextInvoiceNo @invoiceNo nvarchar(10) OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM dbo.CLNMCOINVQ WITH (UPDLOCK, HOLDLOCK))
    BEGIN
        INSERT dbo.CLNMCOINVQ (MCOInvoice_No) VALUES ('00001');
        SET @invoiceNo = 'I00001';
        RETURN;
    END
    DECLARE @n nvarchar(10);
    UPDATE TOP (1) dbo.CLNMCOINVQ SET @n = MCOInvoice_No = RIGHT('00000' + CAST(CAST(MCOInvoice_No AS int) + 1 AS varchar(10)), 5);
    SET @invoiceNo = 'I' + @n;
END
GO

-- UpdateINVPrintingFile: cover 00 -> MBMCLNOthers (also MBMCInvoiced = 'Y'); other covers -> MBMCoveredPersonsCLN (the desktop does not set MBMCInvoiced there).
CREATE OR ALTER PROCEDURE cis.McoList_StampInvoice @member nvarchar(30), @cover nvarchar(2), @invoiceNo nvarchar(30), @user nvarchar(50), @mco money AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @today datetime = CAST(GETDATE() AS date);
    IF @cover = '00'
        UPDATE cis.Hdb_MBMCLNOthers SET MBMCInvoiceDate = @today, MBMCInvoiceNo = @invoiceNo, MBMCInvoiceBy = @user, MBMCInvoiced = 'Y', MBMCLNBilledMCO = @mco
        WHERE MBMNumber = @member;
    ELSE
        UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCInvoiceDate = @today, MBMCInvoiceNo = @invoiceNo, MBMCInvoiceBy = @user, MBMCCLNBilledMCO = @mco
        WHERE MBMCNumber = @member AND MBMCCoverID = @cover;
END
GO

-- ===== main/710_mco_refund_listing.sql =====
-- MCO Fees > Refund Listing (frmMCORefundListing): MCO refunds of endorsed members (MCOFeeModule cases 10 / 11 / 12 with the pasted filters
-- replaced by parameters) and the credit note (CLNMCOCRNoteSeq, stamped on MBMCLNOthers / MBMCoveredPersonsCLN).
SET QUOTED_IDENTIFIER ON;
GO

-- @byBordx 0: endorsement effective date, 1: endorsement bordereaux date. @noted N: without credit note, Y: with one (principals only, as the desktop).
CREATE OR ALTER PROCEDURE cis.McoRefund_Search @payor nvarchar(10) = NULL, @company nvarchar(200) = NULL, @department nvarchar(200) = NULL, @byBordx bit = 0,
    @from datetime = NULL, @to datetime = NULL, @noted char(1) = 'N' AS
BEGIN
    SET NOCOUNT ON;
    SELECT MBMNumber, MBMPolNumber, MBMName, MBMMCOUtilised, MBMCLNStatus, MBMCOVID, MBMCLNEndtEffDate, MBMCLNEndtBordxDate, MBMPolEffDate, MBMPolExpDate,
        MBMCLNMCOFee, MBMMCORefund, MBMBordxDate, DayInsured, MBMIcBcPp, DayUtilised, INSType, MBMCLNCRNote, MBMCLNCRNoteDate, MBMCInvoiceNo, MBMCInvoiceDate, TotalCase
    FROM dbo.VIEW_CLN_MBM_Refund_Search
    WHERE (@payor IS NULL OR PAYCode = @payor) AND (@company IS NULL OR MBMGroupCompany = @company) AND (@department IS NULL OR MBMDepartment = @department)
      AND (@from IS NULL OR (@byBordx = 0 AND MBMCLNEndtEffDate >= @from) OR (@byBordx = 1 AND MBMCLNEndtBordxDate >= @from))
      AND (@to IS NULL OR (@byBordx = 0 AND MBMCLNEndtEffDate <= @to) OR (@byBordx = 1 AND MBMCLNEndtBordxDate <= @to))
      AND ((@noted = 'N' AND (MBMCLNCRNote IS NULL OR MBMCLNCRNote = '')) OR (@noted = 'Y' AND MBMCLNCRNote <> ''))
    ORDER BY MBMNumber
    OPTION (RECOMPILE);
END
GO

-- MCORefundListingSupp: the covered persons of the listed members (the desktop applies the same filters but not the credit-note status).
CREATE OR ALTER PROCEDURE cis.McoRefund_Supp @payor nvarchar(10) = NULL, @company nvarchar(200) = NULL, @department nvarchar(200) = NULL, @byBordx bit = 0,
    @from datetime = NULL, @to datetime = NULL, @members nvarchar(max) AS
BEGIN
    SET NOCOUNT ON;
    SELECT MBMNumber, MBMPolNumber, MBMCName, MBMCLNStatus, MBMCCoverID, MBMCLNEndtEffDate, MBMCLNEndtBordxDate, MBMPolEffDate, MBMPolExpDate, TotalCase,
        MBMCCLNCRNote, MBMCCLNCRNoteDate, MBMCBordxDate, DayInsured, MBMCICBCPPNo, DayUtilised, MBMCInvoiceNo, MBMCInvoiceDate, MBMMCOUtilised, MBMCCLNMCO, MBMMCORefund
    FROM dbo.VIEW_CLN_Supp_Refund_Search
    WHERE ('|' + @members + '|') LIKE '%|' + MBMNumber + '|%'
      AND (@payor IS NULL OR PAYCode = @payor) AND (@company IS NULL OR MBMGroupCompany = @company) AND (@department IS NULL OR MBMDepartment = @department)
      AND (@from IS NULL OR (@byBordx = 0 AND MBMCLNEndtEffDate >= @from) OR (@byBordx = 1 AND MBMCLNEndtBordxDate >= @from))
      AND (@to IS NULL OR (@byBordx = 0 AND MBMCLNEndtEffDate <= @to) OR (@byBordx = 1 AND MBMCLNEndtBordxDate <= @to))
    OPTION (RECOMPILE);
END
GO

-- ComboListing / cboMCOGrpCom_Click / the credit note's company block.
CREATE OR ALTER PROCEDURE cis.McoRefund_Lookups @company nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM dbo.PAY;
    SELECT DISTINCT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL ORDER BY 1;
    SELECT DISTINCT LTRIM(RTRIM(MBMDepartment)) AS Department FROM cis.Hdb_MBMOthers
    WHERE @company IS NOT NULL AND MBMGroupCompany = @company AND COALESCE(LTRIM(RTRIM(MBMDepartment)), '') <> '' ORDER BY 1;
END
GO

CREATE OR ALTER PROCEDURE cis.McoRefund_Company @payor nvarchar(10) = NULL, @company nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF @payor IS NOT NULL
        SELECT TOP 1 PAYCompanyName AS Name, PAYAddress1 AS Add1, PAYAddress2 AS Add2, PAYAddress3 AS Add3, PAYTELno AS Tel, PAYMAINContact AS Contact
        FROM dbo.PAY WHERE PAYCode = @payor;
    ELSE
        SELECT TOP 1 GRPCoName AS Name, GRPCoAdd1 AS Add1, GRPCoAdd2 AS Add2, GRPCoAdd3 AS Add3, GRPCoPhone AS Tel, GRPCoContact AS Contact
        FROM dbo.GRPCompany WHERE GRPCoName = @company;
END
GO

-- GenerateCLNMCOCreditNo: "C" + the next number of HISMaintenance CLNMCOCRNoteSeq, locked for the save.
CREATE OR ALTER PROCEDURE cis.McoRefund_NextCreditNote @creditNote nvarchar(10) OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM cis.Maint_CLNMCOCRNoteSeq WITH (UPDLOCK, HOLDLOCK))
    BEGIN
        INSERT cis.Maint_CLNMCOCRNoteSeq (CLNCRNote) VALUES ('00001');
        SET @creditNote = 'C00001';
        RETURN;
    END
    DECLARE @n nvarchar(10);
    UPDATE TOP (1) cis.Maint_CLNMCOCRNoteSeq SET @n = CLNCRNote = RIGHT('00000' + CAST(CAST(CLNCRNote AS int) + 1 AS varchar(10)), 5);
    SET @creditNote = 'C' + @n;
END
GO

-- UpdateCreditNote: the member and all its covers.
CREATE OR ALTER PROCEDURE cis.McoRefund_StampCreditNote @member nvarchar(30), @creditNote nvarchar(10), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @today datetime = CAST(GETDATE() AS date);
    UPDATE cis.Hdb_MBMCLNOthers SET MBMCLNCRNoteDate = @today, MBMCLNCRNote = @creditNote, MBMCLNCRNoteby = @user WHERE MBMNumber = @member;
    UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCCLNCRNoteDate = @today, MBMCCLNCRNote = @creditNote, MBMCCLNCRNoteby = @user WHERE MBMCNumber = @member;
END
GO

-- ===== main/720_mco_invoice.sql =====
-- MCO Fees > MCO Invoice (frmMCOInvoiceNew): principals (VIEW_CisMCO_Princ), covered persons (VIEW_CisMCO_Supp) and the invoice summary
-- (VIEW_CisMCO_All) — MCOFeeNew cases 1–3 with the pasted criteria replaced by parameters. @mode 0 batch, 1 range, 2 invoice no;
-- @version 0 payor, 1 group company, 2 department; @reprint 0 print new (no invoice no), 1 reprint (the desktop's filter matches every row).
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.McoInv_Principals @mode int, @bordxDate datetime = NULL, @batchNo int = NULL, @from datetime = NULL, @to datetime = NULL,
    @invoiceNo nvarchar(30) = NULL, @version int = 0, @payor nvarchar(10) = NULL, @groupCode nvarchar(10) = NULL, @department nvarchar(200) = NULL, @reprint bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT PAYCode, MBMNumber, MBMMemberType, MBMBordxDate, MBMPolicyNo, MBMBatchNo, MBMName, MBMGroupCompany, MBMCInvoiceNo, MBMCInvoiced, INSCode
    FROM dbo.VIEW_CisMCO_Princ
    WHERE (@mode <> 2 OR MBMCInvoiceNo = @invoiceNo)
      AND (@mode <> 0 OR (MBMBordxDate = @bordxDate AND MBMBatchNo = @batchNo))
      AND (@mode <> 1 OR (MBMBordxDate >= @from AND MBMBordxDate <= @to))
      AND (@mode = 2 OR @version <> 0 OR PAYCode = @payor)
      AND (@mode = 2 OR @version = 0 OR GRPCoCode = @groupCode)
      AND (@mode = 2 OR @version <> 2 OR MBMDepartment = @department)
      AND (@mode = 2 OR @reprint = 1 OR MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')
    ORDER BY MBMNumber
    OPTION (RECOMPILE);
END
GO

-- InvSuppList: covered persons matching the same criteria (@member: one member's, else all).
CREATE OR ALTER PROCEDURE cis.McoInv_Supps @mode int, @bordxDate datetime = NULL, @batchNo int = NULL, @from datetime = NULL, @to datetime = NULL,
    @invoiceNo nvarchar(30) = NULL, @version int = 0, @payor nvarchar(10) = NULL, @groupCode nvarchar(10) = NULL, @department nvarchar(200) = NULL, @reprint bit = 0,
    @member nvarchar(30) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT PAYCode, MBMCNumber, MBMCCoverID, MBMCRELCode, MBMCName, MBMGroupCompany, MBMCBordxDate, MBMCBatchNo, MBMCInvoiceNo, MBMPolicyNo
    FROM dbo.VIEW_CisMCO_Supp
    WHERE (@member IS NULL OR MBMCNumber = @member)
      AND (@mode <> 2 OR MBMCInvoiceNo = @invoiceNo)
      AND (@mode <> 0 OR (MBMCBordxDate = @bordxDate AND MBMCBatchNo = @batchNo))
      AND (@mode <> 1 OR (MBMCBordxDate >= @from AND MBMCBordxDate <= @to))
      AND (@mode = 2 OR @version <> 0 OR PAYCode = @payor)
      AND (@mode = 2 OR @version = 0 OR GRPCoCode = @groupCode)
      AND (@mode = 2 OR @version <> 2 OR MBMDepartment = @department)
      AND (@mode = 2 OR @reprint = 1 OR MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')
    ORDER BY MBMCNumber
    OPTION (RECOMPILE);
END
GO

-- PrintORPreviewInvoice: the summary lines (insurer, bordx date, member type, adult / child, quantity, MCO fee).
CREATE OR ALTER PROCEDURE cis.McoInv_Summary @mode int, @bordxDate datetime = NULL, @batchNo int = NULL, @from datetime = NULL, @to datetime = NULL,
    @invoiceNo nvarchar(30) = NULL, @version int = 0, @payor nvarchar(10) = NULL, @groupCode nvarchar(10) = NULL, @department nvarchar(200) = NULL, @reprint bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT INSCode, MBMBordxDate, MBMMemberType, MBMCInvoiceNo, MBMCInvoiceDate, PAYCode, GRPCoCode, MBMAdultChild,
        CASE WHEN @version = 2 THEN MBMDepartment END AS MBMDepartment, SUM(MBMQty) AS MBMQty, SUM(MCOFee) AS MCOFee
    FROM dbo.VIEW_CisMCO_All
    WHERE (@mode <> 2 OR MBMCInvoiceNo = @invoiceNo)
      AND (@mode <> 0 OR (MBMBordxDate = @bordxDate AND MBMBatchNo = @batchNo))
      AND (@mode <> 1 OR (MBMBordxDate >= @from AND MBMBordxDate <= @to))
      AND (@mode = 2 OR @version <> 0 OR PAYCode = @payor)
      AND (@mode = 2 OR @version = 0 OR GRPCoCode = @groupCode)
      AND (@mode = 2 OR @version <> 2 OR MBMDepartment = @department)
      AND (@mode = 2 OR @reprint = 1 OR MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')
    GROUP BY INSCode, MBMBordxDate, MBMMemberType, MBMCInvoiceNo, MBMCInvoiceDate, PAYCode, GRPCoCode, MBMAdultChild, CASE WHEN @version = 2 THEN MBMDepartment END
    ORDER BY INSCode, MBMBordxDate, MBMAdultChild, MBMMemberType
    OPTION (RECOMPILE);
END
GO

-- PrintPAYInfo: PAY by code (by payor) or GRPCompany by code.
CREATE OR ALTER PROCEDURE cis.McoInv_Company @code nvarchar(10), @byPayor bit AS
BEGIN
    SET NOCOUNT ON;
    IF @byPayor = 1
        SELECT TOP 1 PAYCompanyName AS CompanyName, PAYAddress1 AS Add1, PAYAddress2 AS Add2, PAYAddress3 AS Add3, PAYMAINContact AS Contact, PAYTELno AS TelNo,
            PAYFAXNo AS FaxNo FROM dbo.PAY WHERE PAYCode = @code;
    ELSE
        EXEC cis.McoList_Company @code = @code;
END
GO

-- cboINVGrpCom_LostFocus: SearchDepartment of the company.
CREATE OR ALTER PROCEDURE cis.McoInv_Departments @company nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    SELECT DISTINCT LTRIM(RTRIM(MBMDepartment)) AS Department FROM cis.Hdb_MBMOthers
    WHERE MBMGroupCompany = @company AND COALESCE(LTRIM(RTRIM(MBMDepartment)), '') <> '' ORDER BY 1;
END
GO

-- UpdateINVPrintingFile (this screen does not write the billed MCO).
CREATE OR ALTER PROCEDURE cis.McoInv_Stamp @member nvarchar(30), @cover nvarchar(2), @invoiceNo nvarchar(30), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @today datetime = CAST(GETDATE() AS date);
    IF @cover = '00'
        UPDATE cis.Hdb_MBMCLNOthers SET MBMCInvoiceDate = @today, MBMCInvoiceNo = @invoiceNo, MBMCInvoiceBy = @user, MBMCInvoiced = 'Y' WHERE MBMNumber = @member;
    ELSE
        UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCInvoiceDate = @today, MBMCInvoiceNo = @invoiceNo, MBMCInvoiceBy = @user
        WHERE MBMCNumber = @member AND MBMCCoverID = @cover;
END
GO

-- ===== main/730_mco_invoice_gst.sql =====
-- MCO Fees > MCO Invoice GST (frmMCOFeeGST): SP_MCOGSTInvoiceOP (View_MCOGSTInvoiceOP, the member list and the detail sheet) and
-- SP_MCOSummaryGSTOP (View_OPMCOSummary, the invoice lines) with the pasted criteria replaced by parameters.
-- @mode 0 batch, 1 range (both without invoice no), 2 invoice no; group / subsidiary / payor apply when given (SelectionCriterial).
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.McoGst_List @mode int, @bordxDate datetime = NULL, @batchNo int = NULL, @from datetime = NULL, @to datetime = NULL,
    @invoiceNo nvarchar(30) = NULL, @groupCode nvarchar(10) = NULL, @subsidiary nvarchar(200) = NULL, @payor nvarchar(10) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM dbo.View_MCOGSTInvoiceOP
    WHERE (@mode <> 2 OR MBMCInvoiceNo = @invoiceNo)
      AND (@mode <> 0 OR (MBMBordxDate = @bordxDate AND MBMBatchNo = @batchNo AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')))
      AND (@mode <> 1 OR (MBMBordxDate >= @from AND MBMBordxDate <= @to AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')))
      AND (@groupCode IS NULL OR GRPCoCode = @groupCode)
      AND (@subsidiary IS NULL OR Subsidiary = @subsidiary)
      AND (@payor IS NULL OR PAYCode = @payor)
    OPTION (RECOMPILE);
END
GO

-- PrintInvoice / PrintPreview: the summary lines. @invoiceOnly: Print (by the invoice no only); else the preview filter (criteria + payor).
CREATE OR ALTER PROCEDURE cis.McoGst_Summary @mode int, @bordxDate datetime = NULL, @batchNo int = NULL, @from datetime = NULL, @to datetime = NULL,
    @invoiceNo nvarchar(30) = NULL, @payor nvarchar(10) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM dbo.View_OPMCOSummary
    WHERE (@mode <> 2 OR MBMCInvoiceNo = @invoiceNo)
      AND (@mode <> 0 OR (MBMBordxDate = @bordxDate AND MBMBatchNo = @batchNo))
      AND (@mode <> 1 OR (MBMBordxDate >= @from AND MBMBordxDate <= @to))
      AND (@payor IS NULL OR PAYCode = @payor)
    OPTION (RECOMPILE);
END
GO

-- BindSubsidary: SearchSubsidiary of the company.
CREATE OR ALTER PROCEDURE cis.McoGst_Subsidiaries @company nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    SELECT DISTINCT LTRIM(RTRIM(Subsidiary)) AS Subsidiary FROM cis.Hdb_MBMOthers
    WHERE MBMGroupCompany = @company AND COALESCE(LTRIM(RTRIM(Subsidiary)), '') <> '' ORDER BY 1;
END
GO

CREATE OR ALTER PROCEDURE cis.McoGst_Details @mode int, @bordxDate datetime = NULL, @batchNo int = NULL, @from datetime = NULL, @to datetime = NULL,
    @invoiceNo nvarchar(30) = NULL, @payor nvarchar(10) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM dbo.View_MCOGSTInvoiceOP
    WHERE (@mode <> 2 OR MBMCInvoiceNo = @invoiceNo)
      AND (@mode <> 0 OR (MBMBordxDate = @bordxDate AND MBMBatchNo = @batchNo))
      AND (@mode <> 1 OR (MBMBordxDate >= @from AND MBMBordxDate <= @to))
      AND (@payor IS NULL OR PAYCode = @payor)
    OPTION (RECOMPILE);
END
GO

-- ===== main/740_mco_gst_zeroised.sql =====
-- MCO Fees > MCO Invoice GST Zeroised (frmMCOFeeGSTZeroised): the GST member list with the policy no filter, the zeroised invoice rows
-- (SP_MCOGSTInvoiceOP by invoice no) and its company block (Tbl_Subsidary or CompanyGRP).
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.McoZero_List @mode int, @bordxDate datetime = NULL, @batchNo int = NULL, @from datetime = NULL, @to datetime = NULL,
    @invoiceNo nvarchar(30) = NULL, @groupCode nvarchar(10) = NULL, @subsidiary nvarchar(200) = NULL, @payor nvarchar(10) = NULL, @policyNo nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM dbo.View_MCOGSTInvoiceOP
    WHERE (@mode <> 2 OR MBMCInvoiceNo = @invoiceNo)
      AND (@mode <> 0 OR (MBMBordxDate = @bordxDate AND MBMBatchNo = @batchNo AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')))
      AND (@mode <> 1 OR (MBMBordxDate >= @from AND MBMBordxDate <= @to AND (MBMCInvoiceNo IS NULL OR MBMCInvoiceNo = '')))
      AND (@groupCode IS NULL OR GRPCoCode = @groupCode)
      AND (@subsidiary IS NULL OR Subsidiary = @subsidiary)
      AND (@payor IS NULL OR PAYCode = @payor)
      AND (@policyNo IS NULL OR POLNO = @policyNo)
    OPTION (RECOMPILE);
END
GO

-- PrintZeroisedInvoice's company block: the subsidiary (chosen or the first row's) from Tbl_Subsidary, else the group from CompanyGRP.
CREATE OR ALTER PROCEDURE cis.McoZero_Company @subsidiary nvarchar(200) = NULL, @subsidiary2 nvarchar(200) = NULL, @group nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF @subsidiary IS NOT NULL OR @subsidiary2 IS NOT NULL
        SELECT TOP 1 PIC AS Contact, SubsidaryName AS Name, SubAdd1 AS Add1, SubAdd2 AS Add2, SubAdd3 AS Add3 FROM cis.Hdb_Subsidary
        WHERE SubsidaryName = @subsidiary OR SubsidaryName = @subsidiary2;
    ELSE
        SELECT TOP 1 CoGRPContact AS Contact, CoGRPName AS Name, CoGRPAdd1 AS Add1, CoGRPAdd2 AS Add2, CoGRPAdd3 AS Add3 FROM cis.Hdb_CompanyGRP
        WHERE CoGRPName = @group;
END
GO

-- ===== main/800_bordx_insert.sql =====
-- Claim Bordx > Insert Worksheet (frmWSBordx): the worksheet lists, copied from CISWorksheet ClaimBordxModule_isPureASO (case 21) and
-- ClaimBordxModule (cases 23 UOB flexCare / 24 UOB outpatient) unchanged apart from the filter; the criteria the desktop pasted
-- into @strAppend are parameters here. Bordx numbering and the save are in 810_bordx_insert_save.sql.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxWs_Search @bordxType tinyint, @isPureAso tinyint = 1,
    @groupCode nvarchar(10) = NULL, @scan bit = 0, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL,
    @visitFrom datetime = NULL, @visitTo datetime = NULL, @noVisitDate bit = 0, @department nvarchar(200) = NULL, @clinicType int = 0,
    @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @subsidiary nvarchar(200) = NULL, @payTo char(1) = NULL, @treatment nvarchar(10) = NULL,
    @omni bit = 0, @regFrom datetime = NULL, @regTo datetime = NULL, @clinic nvarchar(20) = NULL, @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = N''
        + CASE WHEN @groupCode IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.CLMGrpComCode = @groupCode' ELSE N'' END
        + CASE WHEN @scan = 1 AND @caseFrom IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber = @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseFrom IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseTo IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber <= @caseTo' ELSE N'' END
        + CASE WHEN @visitFrom IS NOT NULL THEN N' And CISCas.DateVisit >= @visitFrom' ELSE N'' END
        + CASE WHEN @visitTo IS NOT NULL THEN N' And CISCas.DateVisit <= @visitTo' ELSE N'' END
        + CASE WHEN @noVisitDate = 1 THEN N' AND (CISCas.DateVisit Is null OR CISCas.DateVisit = '''')' ELSE N'' END
        + CASE WHEN @department IS NOT NULL THEN N' AND claim.CLMDeptName = @department' ELSE N'' END
        + CASE @clinicType WHEN 1 THEN N' AND HISDB.dbo.CLN.CLNType = ''S''' WHEN 2 THEN N' AND HISDB.dbo.CLN.CLNType = ''G'''
            WHEN 3 THEN N' AND (HISDB.dbo.CLN.CLNType <> ''G'') AND (HISDB.dbo.CLN.CLNType <> ''S'')' ELSE N'' END
        + CASE WHEN @opFrom IS NOT NULL THEN N' AND CISCas.OPBatchNo >= @opFrom' ELSE N'' END
        + CASE WHEN @opTo IS NOT NULL THEN N' AND CISCas.OPBatchNo <= @opTo' ELSE N'' END
        + CASE WHEN @subsidiary IS NOT NULL THEN N' AND claim.clmsubsidiary = @subsidiary' ELSE N'' END
        + CASE @payTo WHEN 'C' THEN N' AND CISWorksheet.dbo.ClaimInvoice.INVPayto = ''C''' WHEN 'M' THEN N' AND CISWorksheet.dbo.ClaimInvoice.INVPayto = ''M''' ELSE N'' END
        + CASE WHEN @treatment IS NOT NULL THEN N' AND CASTreatReason1 = @treatment' ELSE N'' END
        + CASE WHEN @omni = 1 THEN N' AND GRPCompany.GRPOMNICLM = ''Y''' ELSE N'' END
        + CASE WHEN @regFrom IS NOT NULL THEN N' And CISCas.RegisterDate >= @regFrom' ELSE N'' END
        + CASE WHEN @regTo IS NOT NULL THEN N' And CISCas.RegisterDate <= @regTo' ELSE N'' END
        + CASE WHEN @clinic IS NOT NULL THEN N' AND CISCas.ClinicCode = @clinic' ELSE N'' END
        + CASE @entity WHEN 0 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) <>''SJ'' ' WHEN 1 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) =''SJ'' ' ELSE N'' END;

			       declare @sql nvarchar(max)

		                    --set @l_sqlstr2 =  N'Select  top (100) percent * From CISDB..VIEW_CISWS_Pending where 0 = 0'
--						set @sql =  N'SELECT TOP (100) PERCENT CISDB.dbo.CISCas.CASNumber, CISDB.dbo.CISCas.MBMnumber, CISDB.dbo.CISCas.CoverID, CISCas.CASStatus, CISCas.DateVisit, 
--                         CISCas.RegisterDate, HISDB.dbo.MBM.MBMName, HISDB.dbo.MBMCLNOthers.MBMPolNumber AS MBMPolicyNo, 
--                         HISDB.dbo.MBMCoveredPersons.MBMCName, HISDB.dbo.CLN.CLNName, HISDB.dbo.CLN.CLNCode, HISDB.dbo.MBMOthers.MBMGroupCompany, 
--                         HISDB.dbo.MBMCoveredPersons.MBMCRELCode, HISDB.dbo.MBMOthers.RELCode, HISDB.dbo.MBM.HLTCode, HISDB.dbo.MBM.MBMMemberType, 
--                         CISCas.CASTreatReason1, CISCas.CASTreatReason2, CISCas.CASTreatReason3, CISCas.CASTreatReason4, CISCas.CASTreatReason5, 
--                         TreatmentReason1.ReasonDesc AS ReasonDesc1, TreatmentReason3.ReasonDesc AS ReasonDesc3, TreatmentReason2.ReasonDesc AS ReasonDesc2, 
--                         TreatmentReason.ReasonDesc AS ReasonDesc4, TreatmentReason_1.ReasonDesc AS ReasonDesc5, CISCas.ClinicCode, CISCas.Remarks, 
--                         HISDB.dbo.MBM.AGECode, HISDB.dbo.MBM.HSCLNInd, HISDB.dbo.MBMCLNOthers.PLNType AS PLNCode, HISDB.dbo.MBMCLNOthers.INSType AS INSCode, 
--                         CISCas.CASDiagnosis5, HISDB.dbo.MBMCLNOthers.MBMPolNumber, HISDB.dbo.MBMCLNOthers.PAYCode, HISDB.dbo.MBMOthers.MBMDepartment, 
--                         HISDB.dbo.MBMOthers.MBMEmployeeNo, CISCas.InvoiceDate, HISDB.dbo.MBMOthers.MBMCostCenter, GRPCompany.GRPBordxFormat, 
--                         GRPCompany.RptFormat, HISDB.dbo.CLN.CLNType, GRPCompany.TyphoidRpt, HISDB.dbo.MBMOthers.MBMBranch, CISCas.CaseType, 
--                         CISCas.MCDays, HISDB.dbo.MBMOthers.MBMPosition, CISCas.OPBatchNo, CISCas.InvoiceNo, HISDB.dbo.MBM.MBMPolicyNo AS MBMHISPolNo, 
--                         HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMGroupCategory, HISDB.dbo.CLN.CLNFFSStatus, GRPCompany.GRPMCOStatus, 
--                         CISCas.CASClaimsApp, HISDB.dbo.MBMCoveredPersons.MBMAXANumber, HISDB.dbo.MBMOthers.MBMPrevMemNo, CISCas.MedInjType1, 
--                         CISCas.MedInjType2, CISCas.MedInjType3, CISCas.MedInjType4, CISCas.MedInjType5, CISCas.MedInjType6, CISCas.MedInjType7, 
--                         CISWorksheet.dbo.ClaimInvoiceActualOthers.GPMed, CISWorksheet.dbo.ClaimInvoiceActualOthers.SPMed, HISDB.dbo.CLN.CLNLocation, HISDB.dbo.CLN.CLNPState,
--                          HISDB.dbo.MBM.PLNCode AS HISPLN, HISDB.dbo.CLN.CLNState, CISCas.NPClinic, CISWorksheet.dbo.Claim.CASNumber AS Expr1, 
--                         CISWorksheet.dbo.Claim.ClaimStatus, CISWorksheet.dbo.Claim.ClaimBordxStatus, CISWorksheet.dbo.Claim.CLMPayableByMedix2C, 
--                         CISWorksheet.dbo.Claim.CLMTotalActual, CISWorksheet.dbo.Claim.CLMBordxAmt, CISWorksheet.dbo.Claim.CLMNonGPFee, 
--                         CISWorksheet.dbo.Claim.CLMMCOProcFee, CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISWorksheet.dbo.ClaimInvoice.INVBillAmt, 
--                         CISWorksheet.dbo.Claim.CLMFinalApprove AS CLMTotalApproved, GRPCompany.GRPOMNICLM, CISWorksheet.dbo.Claim.CLMBordxMemo, 
--                         GRPCompany.ASO_INS
--FROM           CISDB.dbo.TreatmentReason AS TreatmentReason3 RIGHT OUTER JOIN
--                         CISWorksheet.dbo.Claim WITH (nolock) RIGHT OUTER JOIN
--                         CISWorksheet.dbo.ClaimInvoiceActualOthers WITH (nolock) RIGHT OUTER JOIN
--                         CISDB.dbo.CISCas WITH (nolock) ON CISWorksheet.dbo.ClaimInvoiceActualOthers.CASNumber = CISDB.dbo.CISCas.CASNumber ON 
--                         CISWorksheet.dbo.Claim.CASNumber = CISDB.dbo.CISCas.CASNumber LEFT OUTER JOIN
--                         CISWorksheet.dbo.ClaimInvoice WITH (nolock) ON CISWorksheet.dbo.Claim.CASNumber = CISWorksheet.dbo.ClaimInvoice.CASNumber LEFT OUTER JOIN
--                         HISDB.dbo.CLN ON CISDB.dbo.CISCas.ClinicCode = HISDB.dbo.CLN.CLNCode ON TreatmentReason3.ReasonCode = CISDB.dbo.CISCas.CASTreatReason3 LEFT OUTER JOIN
--                         CISDB.dbo.TreatmentReason AS TreatmentReason_1 WITH (nolock) ON CISDB.dbo.CISCas.CASTreatReason5 = TreatmentReason_1.ReasonCode LEFT OUTER JOIN
--                         CISDB.dbo.TreatmentReason WITH (nolock) ON CISDB.dbo.CISCas.CASTreatReason4 = CISDB.dbo.TreatmentReason.ReasonCode LEFT OUTER JOIN
--                         CISDB.dbo.TreatmentReason AS TreatmentReason2 WITH (nolock) ON CISDB.dbo.CISCas.CASTreatReason2 = TreatmentReason2.ReasonCode LEFT OUTER JOIN
--                         CISDB.dbo.TreatmentReason AS TreatmentReason1 WITH (nolock) ON CISDB.dbo.CISCas.CASTreatReason1 = TreatmentReason1.ReasonCode LEFT OUTER JOIN
--                         HISDB.dbo.MBMCoveredPersons WITH (nolock) ON CISDB.dbo.CISCas.MBMnumber = HISDB.dbo.MBMCoveredPersons.MBMCNumber AND 
--                         CISDB.dbo.CISCas.CoverID = HISDB.dbo.MBMCoveredPersons.MBMCCoverID LEFT OUTER JOIN
--                         CISDB.dbo.GRPCompany RIGHT OUTER JOIN
--                         HISDB.dbo.MBMOthers WITH (nolock) INNER JOIN
--                         HISDB.dbo.MBM WITH (nolock) INNER JOIN
--                         HISDB.dbo.MBMCLNOthers WITH (nolock) ON HISDB.dbo.MBM.MBMNumber = HISDB.dbo.MBMCLNOthers.MBMNumber ON 
--                         HISDB.dbo.MBMOthers.MBMNumber = HISDB.dbo.MBM.MBMNumber ON CISDB.dbo.GRPCompany.GRPCoName = HISDB.dbo.MBMOthers.MBMGroupCompany ON 
--                         CISDB.dbo.CISCas.MBMnumber = HISDB.dbo.MBM.MBMNumber
--WHERE (CISWorksheet.dbo.Claim.ClaimBordxStatus = 0) 
--and CISCas.CASStatus <>''D''
--AND (CISCas.OPBatchNo IS NOT NULL) AND (CISCas.OPBatchNo <> N'''')
--and  ( MBM.HSCLNInd <> ''H'') 
--AND  (CISWorksheet.dbo.Claim.CLMBordxMemo = ''Y'' OR  CISWorksheet.dbo.Claim.CLMBordxMemo = '''' OR CISWorksheet.dbo.Claim.CLMBordxMemo IS NULL)  
--AND  0=0 '



			set @sql =  N'SELECT TOP (100) PERCENT CISDB.dbo.CISCas.CASNumber, CISDB.dbo.CISCas.MBMnumber, CISDB.dbo.CISCas.CoverID, CISCas.CASStatus, CISCas.DateVisit, 
                         CISCas.RegisterDate, HISDB.dbo.MBM.MBMName, HISDB.dbo.MBMCLNOthers.MBMPolNumber AS MBMPolicyNo, 
                      HISDB.dbo.MBMCoveredPersons.MBMCName, HISDB.dbo.CLN.CLNName, HISDB.dbo.CLN.CLNCode, 
					  coalesce(CISDB.dbo.GRPCompany.GRPCoName,'''') as MBMGroupCompany, 
                 HISDB.dbo.MBMCoveredPersons.MBMCRELCode, HISDB.dbo.MBMOthers.RELCode, HISDB.dbo.MBM.HLTCode, HISDB.dbo.MBM.MBMMemberType, 
						CISCas.CASTreatReason1, CISCas.CASTreatReason2, CISCas.CASTreatReason3, CISCas.CASTreatReason4, CISCas.CASTreatReason5,                        
						cisdb.dbo.fnGetTreatmentDesc(CASTreatReason1) as ReasonDesc1,
						cisdb.dbo.fnGetTreatmentDesc(CASTreatReason2) as ReasonDesc2,
						cisdb.dbo.fnGetTreatmentDesc(CASTreatReason3) as ReasonDesc3,
						cisdb.dbo.fnGetTreatmentDesc(CASTreatReason4) as ReasonDesc4,
						cisdb.dbo.fnGetTreatmentDesc(CASTreatReason5) as ReasonDesc5,
						 CISCas.ClinicCode, CISCas.Remarks, 
                         HISDB.dbo.MBM.AGECode, HISDB.dbo.MBM.HSCLNInd, HISDB.dbo.MBMCLNOthers.PLNType AS PLNCode, HISDB.dbo.MBMCLNOthers.INSType AS INSCode, 
                         CISCas.CASDiagnosis5, HISDB.dbo.MBMCLNOthers.MBMPolNumber, HISDB.dbo.MBMCLNOthers.PAYCode, HISDB.dbo.MBMOthers.MBMDepartment, 
                         HISDB.dbo.MBMOthers.MBMEmployeeNo, CISCas.InvoiceDate, HISDB.dbo.MBMOthers.MBMCostCenter, GRPCompany.GRPBordxFormat, 
                         GRPCompany.RptFormat, HISDB.dbo.CLN.CLNType, GRPCompany.TyphoidRpt, HISDB.dbo.MBMOthers.MBMBranch, CISCas.CaseType, 
                         CISCas.MCDays, HISDB.dbo.MBMOthers.MBMPosition, CISCas.OPBatchNo, CISCas.InvoiceNo, HISDB.dbo.MBM.MBMPolicyNo AS MBMHISPolNo, 
                         HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMGroupCategory, HISDB.dbo.CLN.CLNFFSStatus, GRPCompany.GRPMCOStatus, 
                         CISCas.CASClaimsApp, HISDB.dbo.MBMCoveredPersons.MBMAXANumber, HISDB.dbo.MBMOthers.MBMPrevMemNo, CISCas.MedInjType1, 
                         CISCas.MedInjType2, CISCas.MedInjType3, CISCas.MedInjType4, CISCas.MedInjType5, CISCas.MedInjType6, CISCas.MedInjType7, 
                         CISWorksheet.dbo.ClaimInvoiceActualOthers.GPMed, CISWorksheet.dbo.ClaimInvoiceActualOthers.SPMed, HISDB.dbo.CLN.CLNLocation, HISDB.dbo.CLN.CLNPState,
                          HISDB.dbo.MBM.PLNCode AS HISPLN, HISDB.dbo.CLN.CLNState, CISCas.NPClinic, CISWorksheet.dbo.Claim.CASNumber AS Expr1, 
                         CISWorksheet.dbo.Claim.ClaimStatus, CISWorksheet.dbo.Claim.ClaimBordxStatus, CISWorksheet.dbo.Claim.CLMPayableByMedix2C, 
                         CISWorksheet.dbo.Claim.CLMTotalActual, CISWorksheet.dbo.Claim.CLMSubsidiary, CISWorksheet.dbo.Claim.CLMDeptName, '
						 
						 IF @bordxType=0
							set @sql =  @sql + ' CASE WHEN Pln.PLNBTBG = 1 THEN  CISWorksheet.dbo.Claim.CLMBordxAmt - coalesce(Claim.CLMBTBGAmt,0) else CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '							
						 ELSE
							set @sql =  @sql + ' CASE WHEN (Pln.PLNBTBG = 1 and coalesce(CISWorksheet.dbo.Claim.CLMBTBGAmt,0)<>0) THEN CISWorksheet.dbo.Claim.CLMBTBGAmt ELSE CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '
							
						 						 
			set @sql =  @sql + 'CISWorksheet.dbo.Claim.CLMNonGPFee, 
                         CISWorksheet.dbo.Claim.CLMMCOProcFee, CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISWorksheet.dbo.ClaimInvoice.INVBillAmt, 
                         CISWorksheet.dbo.Claim.CLMFinalApprove AS CLMTotalApproved, GRPCompany.GRPOMNICLM, CISWorksheet.dbo.Claim.CLMBordxMemo, 
                         GRPCompany.ASO_INS,lg.DoctorName,CISWorksheet.dbo.claim.CLMCoPayColl,ciscas.Claimability,claim.clmcostcenter
							From cisdb..ciscas
							inner join CISWorksheet.dbo.Claim on claim.CASNumber=ciscas.CASNumber
							inner join CISWorksheet..ClaimInvoice on ClaimInvoice.CASNumber=claim.CASNumber 
							inner join CISWorksheet.dbo.ClaimInvoiceActualOthers on ClaimInvoiceActualOthers.CASNumber = ClaimInvoice.CASNumber 
							inner join hisdb..mbm on mbm.MBMNumber=ciscas.MBMnumber
							left join HISDB.dbo.MBMCoveredPersons on MBMCoveredPersons.MBMCNumber=CISCas.MBMNumber and MBMCCoverID=ciscas.CoverID
							inner join HISDB.dbo.MBMCLNOthers on MBMCLNOthers.MBMNumber=mbm.MBMNumber 
							inner join hisdb..MBMOthers on mbmothers.MBMnumber=mbm.MBMNumber
							left join CISDB.dbo.GRPCompany on GRPCoCode=claim.clmgrpcomcode and (grpcompany.hispaycode is null or ((grpcompany.hispaycode=mbm.paycode or grpcompany.hispaycodePrevYear = mbm.paycode) or (grpcompany.GRPCoPaycode=''DC'' and mbm.MBMPayorEffDate is null)))							
							inner join hisdb..cln on cln.CLNCode=ciscas.ClinicCode
							LEFT JOIN CISDB..LG ON LG.LGNo=CISCAS.CASNumber 
							left join cisdb..pln on pln.PLNCode= MBMCLNOthers.PLNType
							WHERE '
							IF @bordxType=0
								set @sql =  @sql + '(((pln.PLNBTBG =1 and coalesce(CISWorksheet.dbo.Claim.BTBGBordxStatus,0) = 0 and coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0 and coalesce(Claim.CLMBordxAmt,0) <> 0 and coalesce(Claim.CLMTotalActual,0) <> coalesce(Claim.CLMBTBGAmt,0)) OR  ( (coalesce(Claim.CLMBordxAmt,0) <> coalesce(Claim.CLMBTBGAmt,0)) and  coalesce(pln.PureASO,0) =0 and coalesce(pln.PLNBTBG,0) =0 and coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0 and coalesce(Claim.CLMBordxAmt,0) <> 0))) '								
							ELSE if @bordxType=1														
								--set @sql =  @sql + '((pln.PLNBTBG =1 and coalesce(Claim.BTBGBordxStatus,0)=0 and coalesce(Claim.CLMBTBGAmt,0) <> 0) OR (coalesce(pln.PLNBTBG,0) =0 and coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0)) '

								if @isPureAso=1
								begin
									set @sql =  @sql + '
										(
											(coalesce(pln.PureASO,0) =1 and coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0)
											OR
											(GRPCompany.grpcocode in (''W058'',''W003'',''A139'',''C083'') and coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0)
										) 
										'
								end
								else if @isPureAso=2
								begin
									set @sql =  @sql + 'pln.ASOType=1 and coalesce(Claim.BTBGBordxStatus,0)=0 and coalesce(Claim.CLMBTBGAmt,0) <> 0'
								end
								else
								begin

								set @sql =  @sql + '
										(
											(pln.PLNBTBG =1 and coalesce(Claim.BTBGBordxStatus,0)=0 and coalesce(Claim.CLMBTBGAmt,0) <> 0) 
											OR 
											(GRPCompany.grpcocode in (''W058'',''W003'',''A139'',''C083'') and coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0)
										) 
										'
								end
							else IF @bordxType=2
								set @sql =  @sql + ' 
										(

										( pln.PLNBTBG=1

										and coalesce(Claim.CLMBTBGAmt,0) = coalesce(Claim.CLMTotalActual,0)

										and coalesce(BTBGBordxStatus,0)<>1  and coalesce(Claim.ClaimBordxStatus,0)=0 ) 
										Or 
										(
										coalesce(Claim.BTBGBordxStatus,0)=0 and coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0
										)
										) '
							else IF @bordxType=4
								set @sql =  @sql + ' coalesce(Claim.BTBGBordxStatus,0)=0 and (coalesce(CLMBordxAmt,0)<>0 and coalesce(claim.CLMBTBGAmt,0)=0) and (coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0)'										
							else 
							set @sql =  @sql + 'coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0'


			set @sql =  @sql + ' and CISCas.CASStatus NOT IN (''D'' ,''C'')
							AND (CISCas.OPBatchNo IS NOT NULL) AND (CISCas.OPBatchNo <> N'''')
							AND  (CISWorksheet.dbo.Claim.CLMBordxMemo = ''Y'' OR  CISWorksheet.dbo.Claim.CLMBordxMemo = '''' OR CISWorksheet.dbo.Claim.CLMBordxMemo IS NULL)  
							AND  ( HISDB.dbo.MBM.HSCLNInd <> ''H'') AND  (CISDB.dbo.CISCas.pending <> ''Y'') 
							AND CISCas.Claimability =''A'' 
							AND  0=0 '
		
		set @sql = @sql + @filter
				   --print @sql
		exec sp_executesql @sql, N'@groupCode nvarchar(10), @caseFrom nvarchar(30), @caseTo nvarchar(30), @visitFrom datetime, @visitTo datetime,
            @department nvarchar(200), @opFrom nvarchar(30), @opTo nvarchar(30), @subsidiary nvarchar(200), @treatment nvarchar(10), @regFrom datetime,
            @regTo datetime, @clinic nvarchar(20)',
            @groupCode, @caseFrom, @caseTo, @visitFrom, @visitTo, @department, @opFrom, @opTo, @subsidiary, @treatment, @regFrom, @regTo, @clinic;
END
GO

-- @flex 1: CboUOB flexCare (case 23), 0: outpatient (case 24); cash type C / M as @payTo (the UOB filter depends on both).
CREATE OR ALTER PROCEDURE cis.BordxWs_SearchUob @bordxType tinyint, @flex bit,
    @groupCode nvarchar(10) = NULL, @scan bit = 0, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL,
    @visitFrom datetime = NULL, @visitTo datetime = NULL, @noVisitDate bit = 0, @department nvarchar(200) = NULL, @clinicType int = 0,
    @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @subsidiary nvarchar(200) = NULL, @payTo char(1) = NULL, @treatment nvarchar(10) = NULL,
    @omni bit = 0, @regFrom datetime = NULL, @regTo datetime = NULL, @clinic nvarchar(20) = NULL, @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = N''
        + CASE WHEN @groupCode IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.CLMGrpComCode = @groupCode' ELSE N'' END
        + CASE WHEN @scan = 1 AND @caseFrom IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber = @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseFrom IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseTo IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber <= @caseTo' ELSE N'' END
        + CASE WHEN @visitFrom IS NOT NULL THEN N' And CISCas.DateVisit >= @visitFrom' ELSE N'' END
        + CASE WHEN @visitTo IS NOT NULL THEN N' And CISCas.DateVisit <= @visitTo' ELSE N'' END
        + CASE WHEN @noVisitDate = 1 THEN N' AND (CISCas.DateVisit Is null OR CISCas.DateVisit = '''')' ELSE N'' END
        + CASE WHEN @department IS NOT NULL THEN N' AND MBMOthers.MBMDepartment = @department' ELSE N'' END
        + CASE @clinicType WHEN 1 THEN N' AND CLN.CLNType = ''S''' WHEN 2 THEN N' AND CLN.CLNType = ''G'''
            WHEN 3 THEN N' AND (CLN.CLNType <> ''G'') AND (CLN.CLNType <> ''S'')' ELSE N'' END
        + CASE WHEN @opFrom IS NOT NULL THEN N' AND CISCas.OPBatchNo >= @opFrom' ELSE N'' END
        + CASE WHEN @opTo IS NOT NULL THEN N' AND CISCas.OPBatchNo <= @opTo' ELSE N'' END
        + CASE WHEN @subsidiary IS NOT NULL THEN N' AND MBMOthers.Subsidiary = @subsidiary' ELSE N'' END
        + CASE @payTo WHEN 'C' THEN N' AND CISWorksheet.dbo.ClaimInvoice.INVPayto = ''C''' WHEN 'M' THEN N' AND CISWorksheet.dbo.ClaimInvoice.INVPayto = ''M''' ELSE N'' END
        + CASE WHEN @treatment IS NOT NULL THEN N' AND CASTreatReason1 = @treatment' ELSE N'' END
        + CASE WHEN @omni = 1 THEN N' AND GRPCompany.GRPOMNICLM = ''Y''' ELSE N'' END
        + CASE WHEN @regFrom IS NOT NULL THEN N' And CISCas.RegisterDate >= @regFrom' ELSE N'' END
        + CASE WHEN @regTo IS NOT NULL THEN N' And CISCas.RegisterDate <= @regTo' ELSE N'' END
        + CASE WHEN @clinic IS NOT NULL THEN N' AND CISCas.ClinicCode = @clinic' ELSE N'' END
        + CASE @entity WHEN 0 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) <>''SJ'' ' WHEN 1 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) =''SJ'' ' ELSE N'' END
        + CASE WHEN @payTo = 'C' AND @flex = 0 THEN N' AND ((tblFlexibleBenefitPlanLimit_OP.CISPlnCode is NOT null and tblFlexibleBenefitPlan_OP.FBCode is null) OR (tblFlexibleBenefitPlanLimit_OP.CISPlnCode is null) or (tblFlexibleBenefitPlan_OP.FBFamily=1 and ciscas.CoverID<>''00'') or tblFlexibleBenefitPlan_OP.FBisOPFlexCombine = 1 )'
            WHEN @payTo = 'C' AND @flex = 1 THEN N' AND (tblFlexibleBenefitPlan_OP.FBEmployee is not null and CISCas.CoverID=''00'') AND (tblFlexibleBenefitPlan_OP.FBisOPFlexCombine <> 1) '
            WHEN @payTo = 'M' AND @flex = 0 THEN N' AND ((tblFlexibleBenefitPlanLimit_OP.CISPlnCode is NOT null and tblFlexibleBenefitPlan_OP.FBCode is null) OR (tblFlexibleBenefitPlanLimit_OP.CISPlnCode is null) or (tblFlexibleBenefitPlan_OP.FBFamily=1 and ciscas.CoverID<>''00''))'
            WHEN @payTo = 'M' AND @flex = 1 THEN N' AND (tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null and tblFlexibleBenefitPlan_OP.FBEmployee=1 and ciscas.CoverID=''00'')'
            ELSE N'' END;
    IF @flex = 1
    BEGIN
	--flexCARE
			       declare @sql nvarchar(max)

		                    set @sql =  N'SELECT TOP (100) PERCENT CISDB.dbo.CISCas.CASNumber, CISDB.dbo.CISCas.MBMnumber, CISDB.dbo.CISCas.CoverID, CISCas.CASStatus, CISCas.DateVisit, 
												CISCas.RegisterDate, HISDB.dbo.MBM.MBMName, HISDB.dbo.MBMCLNOthers.MBMPolNumber AS MBMPolicyNo, 
												HISDB.dbo.MBMCoveredPersons.MBMCName, HISDB.dbo.CLN.CLNName, HISDB.dbo.CLN.CLNCode, CISDB.dbo.GRPCompany.GRPCoName as MBMGroupCompany, 
												HISDB.dbo.MBMCoveredPersons.MBMCRELCode, HISDB.dbo.MBMOthers.RELCode, HISDB.dbo.MBM.HLTCode, HISDB.dbo.MBM.MBMMemberType, 
												CISCas.CASTreatReason1, CISCas.CASTreatReason2, CISCas.CASTreatReason3, CISCas.CASTreatReason4, CISCas.CASTreatReason5,                        
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason1) as ReasonDesc1,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason2) as ReasonDesc2,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason3) as ReasonDesc3,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason4) as ReasonDesc4,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason5) as ReasonDesc5,
												CISCas.ClinicCode, CISCas.Remarks, 
												HISDB.dbo.MBM.AGECode, HISDB.dbo.MBM.HSCLNInd, HISDB.dbo.MBMCLNOthers.PLNType AS PLNCode, HISDB.dbo.MBMCLNOthers.INSType AS INSCode, 
												CISCas.CASDiagnosis5, HISDB.dbo.MBMCLNOthers.MBMPolNumber, HISDB.dbo.MBMCLNOthers.PAYCode, HISDB.dbo.MBMOthers.MBMDepartment, 
												HISDB.dbo.MBMOthers.MBMEmployeeNo, CISCas.InvoiceDate, HISDB.dbo.MBMOthers.MBMCostCenter, GRPCompany.GRPBordxFormat, 
												GRPCompany.RptFormat, HISDB.dbo.CLN.CLNType, GRPCompany.TyphoidRpt, HISDB.dbo.MBMOthers.MBMBranch, CISCas.CaseType, 
												CISCas.MCDays, HISDB.dbo.MBMOthers.MBMPosition, CISCas.OPBatchNo, CISCas.InvoiceNo, HISDB.dbo.MBM.MBMPolicyNo AS MBMHISPolNo, 
												HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMGroupCategory, HISDB.dbo.CLN.CLNFFSStatus, GRPCompany.GRPMCOStatus, 
												CISCas.CASClaimsApp, HISDB.dbo.MBMCoveredPersons.MBMAXANumber, HISDB.dbo.MBMOthers.MBMPrevMemNo, CISCas.MedInjType1, 
												CISCas.MedInjType2, CISCas.MedInjType3, CISCas.MedInjType4, CISCas.MedInjType5, CISCas.MedInjType6, CISCas.MedInjType7, 
												CISWorksheet.dbo.ClaimInvoiceActualOthers.GPMed, CISWorksheet.dbo.ClaimInvoiceActualOthers.SPMed, HISDB.dbo.CLN.CLNLocation, HISDB.dbo.CLN.CLNPState,
												HISDB.dbo.MBM.PLNCode AS HISPLN, HISDB.dbo.CLN.CLNState, CISCas.NPClinic, CISWorksheet.dbo.Claim.CASNumber AS Expr1, 
												CISWorksheet.dbo.Claim.ClaimStatus, CISWorksheet.dbo.Claim.ClaimBordxStatus, CISWorksheet.dbo.Claim.CLMPayableByMedix2C, 
												CISWorksheet.dbo.Claim.CLMTotalActual, CISWorksheet.dbo.Claim.CLMSubsidiary, CISWorksheet.dbo.Claim.CLMDeptName,'
						  IF @bordxType=0
							set @sql =  @sql + ' CASE WHEN Pln.PLNBTBG = 1 THEN  CISWorksheet.dbo.Claim.CLMBordxAmt - coalesce(Claim.CLMBTBGAmt,0) else CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '							
						 ELSE
							set @sql =  @sql + ' CASE WHEN Pln.PLNBTBG = 1 THEN CISWorksheet.dbo.Claim.CLMBTBGAmt ELSE CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '
		
		  set @sql =  @sql + 'CISWorksheet.dbo.Claim.CLMNonGPFee, 
											CISWorksheet.dbo.Claim.CLMMCOProcFee, CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISWorksheet.dbo.ClaimInvoice.INVBillAmt, 
											CISWorksheet.dbo.Claim.CLMFinalApprove AS CLMTotalApproved, GRPCompany.GRPOMNICLM, CISWorksheet.dbo.Claim.CLMBordxMemo, 
											GRPCompany.ASO_INS,lg.DoctorName,CISWorksheet.dbo.claim.CLMCoPayColl,MBM.PAYCode AS MemberPayCode,claim.clmcostcenter
											From cisdb..ciscas
											INNER JOIN cisdb..TreatmentReason on TreatmentReason.ReasonCode = case when coalesce(ciscas.CASTreatReason1,'''')='''' then ciscas.CASTreatReason5 else ciscas.CASTreatReason1 end
											INNER JOIN HISDB..MBM ON MBM.MBMNumber = CISCas.MBMNumber
											INNER JOIN HISDB..MBMCLNOthers ON MBMCLNOthers.MBMNumber = MBM.MBMNUMBER
											INNER JOIN hisdb..MBMOthers on mbmothers.MBMnumber=mbm.MBMNumber
											LEFT JOIN HISDB.dbo.MBMCoveredPersons on MBMCoveredPersons.MBMCNumber=CISCas.MBMNumber and MBMCoveredPersons.MBMCCoverID = CISCas.CoverID
											LEFT JOIN cisdb..pln on pln.PLNCode= MBMCLNOthers.PLNType
											INNER JOIN hisdb..cln on cln.CLNCode=ciscas.ClinicCode
											INNER JOIN CISWorksheet..claim on claim.CASNumber=ciscas.CASNumber
											LEFT OUTER JOIN CISDB.dbo.GRPCompany on GRPCoCode=claim.clmgrpcomcode 
											INNER JOIN CISWorksheet..ClaimInvoice on ClaimInvoice.CASNumber=claim.CASNumber
											LEFT JOIN CISWorksheet.dbo.ClaimInvoiceActualOthers on ClaimInvoiceActualOthers.CASNumber = ClaimInvoice.CASNumber
											LEFT JOIN CISDB..LG ON LG.LGNo=CISCAS.CASNumber 
											LEFT JOIN cisdb..tblFlexibleBenefitPlan_OP on tblFlexibleBenefitPlan_OP.FBCode=TreatmentReason.ReasonCode
											INNER JOIN CISDB..tblFlexibleBenefitPlanLimit_OP on tblFlexibleBenefitPlanLimit_OP.CISPlnCode = MBMCLNOthers.PLNType  and (ciscas.DateVisit between tblFlexibleBenefitPlanLimit_OP.FBEffDate and tblFlexibleBenefitPlanLimit_OP.FBExpDate)
											where '

							IF @bordxType=0
								set @sql =  @sql + '(coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0) '								
							ELSE																
								set @sql =   @sql + '(coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0) '	

							set @sql =  @sql + 'AND CISCas.CASStatus <>''D''
							AND coalesce(CISCas.OPBatchNo,'''')<>''''
							AND coalesce(Claim.CLMBordxMemo,'''')<>''N''
							AND  ( CISDB.dbo.CISCas.pending <> ''Y'')  
							AND 0=0 '
							
		set @sql = @sql + @filter
		exec sp_executesql @sql, N'@groupCode nvarchar(10), @caseFrom nvarchar(30), @caseTo nvarchar(30), @visitFrom datetime, @visitTo datetime,
            @department nvarchar(200), @opFrom nvarchar(30), @opTo nvarchar(30), @subsidiary nvarchar(200), @treatment nvarchar(10), @regFrom datetime,
            @regTo datetime, @clinic nvarchar(20)',
            @groupCode, @caseFrom, @caseTo, @visitFrom, @visitTo, @department, @opFrom, @opTo, @subsidiary, @treatment, @regFrom, @regTo, @clinic;
    END
    ELSE
    BEGIN
	--op
			       SET @sql = N''

		                    set @sql =  N'SELECT TOP (100) PERCENT CISDB.dbo.CISCas.CASNumber, CISDB.dbo.CISCas.MBMnumber, CISDB.dbo.CISCas.CoverID, CISCas.CASStatus, CISCas.DateVisit, 
												CISCas.RegisterDate, HISDB.dbo.MBM.MBMName, HISDB.dbo.MBMCLNOthers.MBMPolNumber AS MBMPolicyNo, 
												HISDB.dbo.MBMCoveredPersons.MBMCName, HISDB.dbo.CLN.CLNName, HISDB.dbo.CLN.CLNCode, CISDB.dbo.GRPCompany.GRPCoName as MBMGroupCompany, 
												HISDB.dbo.MBMCoveredPersons.MBMCRELCode, HISDB.dbo.MBMOthers.RELCode, HISDB.dbo.MBM.HLTCode, HISDB.dbo.MBM.MBMMemberType, 
												CISCas.CASTreatReason1, CISCas.CASTreatReason2, CISCas.CASTreatReason3, CISCas.CASTreatReason4, CISCas.CASTreatReason5,                        
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason1) as ReasonDesc1,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason2) as ReasonDesc2,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason3) as ReasonDesc3,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason4) as ReasonDesc4,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason5) as ReasonDesc5,
												CISCas.ClinicCode, CISCas.Remarks, 
												HISDB.dbo.MBM.AGECode, HISDB.dbo.MBM.HSCLNInd, HISDB.dbo.MBMCLNOthers.PLNType AS PLNCode, HISDB.dbo.MBMCLNOthers.INSType AS INSCode, 
												CISCas.CASDiagnosis5, HISDB.dbo.MBMCLNOthers.MBMPolNumber, HISDB.dbo.MBMCLNOthers.PAYCode, HISDB.dbo.MBMOthers.MBMDepartment, 
												HISDB.dbo.MBMOthers.MBMEmployeeNo, CISCas.InvoiceDate, HISDB.dbo.MBMOthers.MBMCostCenter, GRPCompany.GRPBordxFormat, 
												GRPCompany.RptFormat, HISDB.dbo.CLN.CLNType, GRPCompany.TyphoidRpt, HISDB.dbo.MBMOthers.MBMBranch, CISCas.CaseType, 
												CISCas.MCDays, HISDB.dbo.MBMOthers.MBMPosition, CISCas.OPBatchNo, CISCas.InvoiceNo, HISDB.dbo.MBM.MBMPolicyNo AS MBMHISPolNo, 
												HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMGroupCategory, HISDB.dbo.CLN.CLNFFSStatus, GRPCompany.GRPMCOStatus, 
												CISCas.CASClaimsApp, HISDB.dbo.MBMCoveredPersons.MBMAXANumber, HISDB.dbo.MBMOthers.MBMPrevMemNo, CISCas.MedInjType1, 
												CISCas.MedInjType2, CISCas.MedInjType3, CISCas.MedInjType4, CISCas.MedInjType5, CISCas.MedInjType6, CISCas.MedInjType7, 
												CISWorksheet.dbo.ClaimInvoiceActualOthers.GPMed, CISWorksheet.dbo.ClaimInvoiceActualOthers.SPMed, HISDB.dbo.CLN.CLNLocation, HISDB.dbo.CLN.CLNPState,
												HISDB.dbo.MBM.PLNCode AS HISPLN, HISDB.dbo.CLN.CLNState, CISCas.NPClinic, CISWorksheet.dbo.Claim.CASNumber AS Expr1, 
												CISWorksheet.dbo.Claim.ClaimStatus, CISWorksheet.dbo.Claim.ClaimBordxStatus, CISWorksheet.dbo.Claim.CLMPayableByMedix2C, 
												CISWorksheet.dbo.Claim.CLMTotalActual, CISWorksheet.dbo.Claim.CLMSubsidiary, CISWorksheet.dbo.Claim.CLMDeptName,'
						  IF @bordxType=0
							set @sql =  @sql + ' CASE WHEN Pln.PLNBTBG = 1 THEN  CISWorksheet.dbo.Claim.CLMBordxAmt - coalesce(Claim.CLMBTBGAmt,0) else CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '							
						 ELSE
							set @sql =  @sql + ' CASE WHEN Pln.PLNBTBG = 1 THEN CISWorksheet.dbo.Claim.CLMBTBGAmt ELSE CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '
		
		  set @sql =  @sql + 'CISWorksheet.dbo.Claim.CLMNonGPFee, 
											CISWorksheet.dbo.Claim.CLMMCOProcFee, CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISWorksheet.dbo.ClaimInvoice.INVBillAmt, 
											CISWorksheet.dbo.Claim.CLMFinalApprove AS CLMTotalApproved, GRPCompany.GRPOMNICLM, CISWorksheet.dbo.Claim.CLMBordxMemo, 
											GRPCompany.ASO_INS,lg.DoctorName,CISWorksheet.dbo.claim.CLMCoPayColl,MBM.PAYCode AS MemberPayCode,claim.clmcostcenter
											From cisdb..ciscas
											LEFT JOIN cisdb..TreatmentReason on TreatmentReason.ReasonCode = case when coalesce(ciscas.CASTreatReason1,'''')='''' then ciscas.CASTreatReason5 else ciscas.CASTreatReason1 end
											INNER JOIN HISDB..MBM ON MBM.MBMNumber = CISCas.MBMNumber
											INNER JOIN HISDB..MBMCLNOthers ON MBMCLNOthers.MBMNumber = MBM.MBMNUMBER
											INNER JOIN hisdb..MBMOthers on mbmothers.MBMnumber=mbm.MBMNumber
											LEFT JOIN HISDB.dbo.MBMCoveredPersons on MBMCoveredPersons.MBMCNumber=CISCas.MBMNumber and MBMCoveredPersons.MBMCCoverID = CISCas.CoverID
											LEFT JOIN cisdb..pln on pln.PLNCode= MBMCLNOthers.PLNType
											INNER JOIN hisdb..cln on cln.CLNCode=ciscas.ClinicCode
											INNER JOIN CISWorksheet..claim on claim.CASNumber=ciscas.CASNumber
											LEFT OUTER JOIN CISDB.dbo.GRPCompany on GRPCoCode=claim.clmgrpcomcode 
											INNER JOIN CISWorksheet..ClaimInvoice on ClaimInvoice.CASNumber=claim.CASNumber
											LEFT JOIN CISWorksheet.dbo.ClaimInvoiceActualOthers on ClaimInvoiceActualOthers.CASNumber = ClaimInvoice.CASNumber
											LEFT JOIN CISDB..LG ON LG.LGNo=CISCAS.CASNumber 
											LEFT JOIN cisdb..tblFlexibleBenefitPlan_OP on tblFlexibleBenefitPlan_OP.FBCode=TreatmentReason.ReasonCode
											LEFT JOIN CISDB..tblFlexibleBenefitPlanLimit_OP on tblFlexibleBenefitPlanLimit_OP.CISPlnCode = MBMCLNOthers.PLNType and (ciscas.DateVisit between tblFlexibleBenefitPlanLimit_OP.FBEffDate and tblFlexibleBenefitPlanLimit_OP.FBExpDate)
											where '

							IF @bordxType=0
								set @sql =  @sql + '(coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0) '								
							ELSE																
								set @sql =   @sql + '(coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0) '	

							set @sql =  @sql + 'AND CISCas.CASStatus <>''D''
							AND coalesce(CISCas.OPBatchNo,'''')<>''''
							AND coalesce(Claim.CLMBordxMemo,'''')<>''N''
							AND  ( CISDB.dbo.CISCas.pending <> ''Y'')  
							AND 0=0 '
							
		set @sql = @sql + @filter
		exec sp_executesql @sql, N'@groupCode nvarchar(10), @caseFrom nvarchar(30), @caseTo nvarchar(30), @visitFrom datetime, @visitTo datetime,
            @department nvarchar(200), @opFrom nvarchar(30), @opTo nvarchar(30), @subsidiary nvarchar(200), @treatment nvarchar(10), @regFrom datetime,
            @regTo datetime, @clinic nvarchar(20)',
            @groupCode, @caseFrom, @caseTo, @visitFrom, @visitTo, @department, @opFrom, @opTo, @subsidiary, @treatment, @regFrom, @regTo, @clinic;
    END
END
GO

-- ===== main/810_bordx_insert_save.sql =====
-- Claim Bordx > Insert Worksheet (frmWSBordx): look-ups, the BTBG plan check and CmdBordx_Click's save.
SET QUOTED_IDENTIFIER ON;
GO

-- PayorComboListing (GRPCoName|GRPCoCode), the treatment codes and the panel clinics (ChkClinic).
CREATE OR ALTER PROCEDURE cis.BordxWs_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoCode) AS Code, RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany;
    SELECT RTRIM(ReasonCode) AS Code, RTRIM(ReasonDesc) AS Name FROM dbo.TreatmentReason;
    SELECT RTRIM(CLNCode) AS Code, RTRIM(CLNName) AS Name FROM cis.Hdb_CLN WHERE (CLNPanel = 'Y' OR CLNPanel = 'C') AND LEN(CLNName) > 0;
END
GO

-- CboPayor_LostFocus: SearchDepartment / SearchSubsidiary_bygrp of the group company name.
CREATE OR ALTER PROCEDURE cis.BordxWs_Company @company nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @d TABLE (MBMDepartment nvarchar(200));
    DECLARE @s TABLE (MBMGRPCompany nvarchar(200), MBMGRPSubsidiary nvarchar(200));
    INSERT @d EXEC HISDB.dbo.SearchDepartment @company;
    INSERT @s EXEC HISDB.dbo.SearchSubsidiary_bygrp @company;
    SELECT DISTINCT LTRIM(RTRIM(MBMDepartment)) AS Name FROM @d WHERE COALESCE(LTRIM(RTRIM(MBMDepartment)), '') <> '';
    SELECT DISTINCT LTRIM(RTRIM(MBMGRPSubsidiary)) AS Name FROM @s WHERE COALESCE(LTRIM(RTRIM(MBMGRPSubsidiary)), '') <> '';
END
GO

-- CheckIsBTBG: sp_GetPLNBTBGInd.
CREATE OR ALTER PROCEDURE cis.BordxWs_IsBtbg @plnCode nvarchar(8), @isBtbg bit OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    SET @isBtbg = 0;
    EXEC dbo.sp_GetPLNBTBGInd @PLNCode = @plnCode, @RetOut = @isBtbg OUTPUT;
END
GO

-- GenerateCLMBordxNo / GenerateCLMBordxNoVOID: K + payor + the next 5-digit number of CLAIM_BordxSEQ (KIV + number for void), locked for the save.
CREATE OR ALTER PROCEDURE cis.BordxWs_NextNo @payCode nvarchar(4), @bordxNo nvarchar(20) OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @n nvarchar(10);
    IF NOT EXISTS (SELECT 1 FROM CISWorksheet.dbo.CLAIM_BordxSEQ WITH (UPDLOCK, HOLDLOCK) WHERE PAYCode = @payCode)
    BEGIN
        INSERT CISWorksheet.dbo.CLAIM_BordxSEQ (HLTCode, PAYCode, CLMBordxNo)
        VALUES (CASE WHEN @payCode = 'IV' THEN NULL ELSE (SELECT TOP 1 HLTCode FROM dbo.PAY WHERE PAYCode = @payCode) END, @payCode, '00001');
        SET @n = '00001';
    END
    ELSE
        UPDATE TOP (1) CISWorksheet.dbo.CLAIM_BordxSEQ
        SET @n = CLMBordxNo = RIGHT('00000' + CAST(CAST(CLMBordxNo AS int) + 1 AS varchar(10)), CASE WHEN LEN(CAST(CAST(CLMBordxNo AS int) + 1 AS varchar(10))) > 5 THEN LEN(CAST(CAST(CLMBordxNo AS int) + 1 AS varchar(10))) ELSE 5 END)
        WHERE PAYCode = @payCode;
    SET @bordxNo = CASE WHEN @payCode = 'IV' THEN 'KIV' ELSE 'K' + @payCode END + @n;
END
GO

-- UpdateRecord / UpdateRecordvoid / UpdateBTBGRecord for one case; returns the amount the desktop adds to the bordx total.
-- @mode 0 normal (CLMFinalApprove), 1 void (+ Payment2MBM 'VOID' row for a member reimbursement), 2 BTBG (CLMBTBGAmt).
CREATE OR ALTER PROCEDURE cis.BordxWs_UpdateCase @casNumber nvarchar(30), @bordxNo nvarchar(20), @mode int, @user nvarchar(50), @amount money OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @today datetime = CAST(GETDATE() AS date);
    SELECT @amount = CASE WHEN @mode = 2 THEN COALESCE(CLMBTBGAmt, 0) ELSE CLMFinalApprove END FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber;
    IF @mode = 2
    BEGIN
        UPDATE CISWorksheet.dbo.Claim SET BTBGBordxRefNo = @bordxNo, BTBGBordxStatus = 1, CLMPending = 'N' WHERE CASNumber = @casNumber;
        UPDATE CISWorksheet.dbo.ClaimInvoice SET BTBGINVBordxRefNo = @bordxNo, BTBGINVBordxed = 1 WHERE CASNumber = @casNumber;
    END
    ELSE
    BEGIN
        UPDATE CISWorksheet.dbo.Claim SET BordxRefNo = @bordxNo, ClaimBordxStatus = 1, CLMPending = 'N', CLMBordxMonth = FORMAT(@today, 'MM'),
            CLMBordxYear = FORMAT(@today, 'yyyy') WHERE CASNumber = @casNumber;
        UPDATE CISWorksheet.dbo.ClaimInvoice SET INVBordxRefNo = @bordxNo, INVBordxed = 1 WHERE CASNumber = @casNumber;
    END
    UPDATE dbo.CISCas SET Pending = 'N' WHERE CASNumber = @casNumber;
    IF @mode = 1 AND EXISTS (SELECT 1 FROM CISWorksheet.dbo.ClaimInvoice WHERE CASNumber = @casNumber AND RTRIM(INVPayTo) = 'M')
        INSERT CISPayment.dbo.Payment2MBM (PVNo, PVDate, PVChequeDate, PVChequeNo, CASNumber, Pay2MBMLastUpdateDate, Pay2MBMLastUpdateUser, PAYAccPeriodMth,
            PAYAccPeriodYr, PVPayMBMStatus)
        VALUES ('VOID', @today, @today, 'VOID', @casNumber, @today, @user, MONTH(@today), YEAR(@today), 'F');
END
GO

-- SaveCLMBordx.
CREATE OR ALTER PROCEDURE cis.BordxWs_SaveBordx @bordxNo nvarchar(20), @payCode nvarchar(4), @user nvarchar(50), @amount money, @claims int,
    @company nvarchar(200), @rptType nvarchar(2), @subsidiary nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISWorksheet.dbo.Bordx (BordxPAYCode, BordxRefNo, BordxDate, BordxBy, BordxClaimAmt, BordxNoClaims, GroupCompany, RptTypeByDept, BordxSubsidary)
    VALUES (@payCode, @bordxNo, CAST(GETDATE() AS date), @user, @amount, @claims, @company, @rptType, @subsidiary);
END
GO

-- ===== main/820_bordx_remove.sql =====
-- Claim Bordx > Worksheet > Remove Worksheet (FrmBordxRevise): a bordereaux's worksheets (ClaimBordxModule case 3) and RemoveFromBordx.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxRemove_List @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 30000 CASNumber, ClaimOpendate, MBMNumber, PatientName, DateVisit, CLMTotalApproved, CASStatus, BordxPrintedDate, BordxDate, BordxClaimAmt
    FROM CISWorksheet.dbo.VIEW_CIS_BordxRevise WHERE BordxRefNo = @bordxNo;
END
GO

-- RemoveFromBordx for one case: the claim count comes off the case's own bordx, the amount off the bordx typed on the screen (as the desktop).
CREATE OR ALTER PROCEDURE cis.BordxRemove_Case @bordxNo nvarchar(20), @casNumber nvarchar(30), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @caseBordx nvarchar(20), @approved money, @bordxDate datetime;
    SELECT TOP 1 @caseBordx = BordxRefNo, @approved = CLMFinalApprove FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber;
    SELECT TOP 1 @bordxDate = BordxDate FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @bordxNo;
    UPDATE CISWorksheet.dbo.Bordx SET BordxNoClaims = BordxNoClaims - 1 WHERE BordxRefNo = @caseBordx;
    INSERT CISWorksheet.dbo.BordxWSArchive (CASNumber, BordxRefNo, BordxDate, LastUpdateDate, LastUpdateUser)
    VALUES (@casNumber, @bordxNo, @bordxDate, CAST(GETDATE() AS date), @user);
    UPDATE CISWorksheet.dbo.Bordx SET BordxClaimAmt = BordxClaimAmt - @approved WHERE BordxRefNo = @bordxNo;
    UPDATE CISWorksheet.dbo.ClaimInvoice SET INVBordxRefNo = '', INVBordxed = 0 WHERE INVBordxRefNo = @bordxNo AND CASNumber = @casNumber;
    UPDATE CISWorksheet.dbo.Claim SET ClaimBordxStatus = 0, BordxRefNo = '',
        CLMBordxWSStatus = CASE WHEN CLMBordxWSASCII = 'Y' THEN 'D' ELSE CLMBordxWSStatus END,
        CLMBordxWSASCII = CASE WHEN CLMBordxWSASCII = 'Y' THEN 'N' ELSE CLMBordxWSASCII END
    WHERE CASNumber = @casNumber;
END
GO

-- ===== main/830_bordx_date.sql =====
-- Claim Bordx > ClaimBordxDate (FrmBordxDate): BordxDateChange and UpdateBordxDate.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxDate_Find @bordxNo nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    EXEC CISWorksheet.dbo.BordxDateChange @BordxNo = @bordxNo;
END
GO

CREATE OR ALTER PROCEDURE cis.BordxDate_Update @bordxNo nvarchar(20), @bordxDate datetime, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISWorksheet.dbo.Bordx SET BordxBy = @user, BordxDate = @bordxDate WHERE BordxRefNo = @bordxNo;
    UPDATE CISWorksheet.dbo.Claim SET CLMBordxMonth = FORMAT(@bordxDate, 'MM'), CLMBordxYear = FORMAT(@bordxDate, 'yyyy') WHERE BordxRefNo = @bordxNo;
END
GO

-- ===== main/840_bordx_delivery.sql =====
-- Claim Bordx > Bordx Delivery: Send to Grp. Company (FrmSendtoPayor, ClaimBordxModule case 7), Returned Fr Grp. Company (FrmReturnPayor, case 8 / 9),
-- Send to A/C (FrmSendtoAC, case 10) and Send to MNRB (FrmSendtoMNRB, its own select). @kind: payor, returned, ac, mnrb.
-- The pasted criteria are parameters; amounts / periods are compared as the desktop did (text against the columns).
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxDelivery_Search @kind varchar(10), @company nvarchar(200) = NULL, @payCode nvarchar(10) = NULL, @scan bit = 0,
    @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL, @dateFrom datetime = NULL, @dateTo datetime = NULL, @amountFrom nvarchar(20) = NULL,
    @amountTo nvarchar(20) = NULL, @monthFrom nvarchar(4) = NULL, @monthTo nvarchar(4) = NULL, @yearFrom nvarchar(6) = NULL, @yearTo nvarchar(6) = NULL,
    @status char(1) = NULL, @sentFrom datetime = NULL, @sentTo datetime = NULL,
    @periodMonthFrom nvarchar(2) = NULL, @periodYearFrom nvarchar(4) = NULL, @periodMonthTo nvarchar(2) = NULL, @periodYearTo nvarchar(4) = NULL, @noPeriod bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    -- @kind acperiod (FrmACPeriod, case 11): the A/C period filter (the year only goes with months 10-12, as the desktop) instead of month / year ranges.
    SELECT TOP 30000 BordxRefNo, BordxPAYCode, BordxDate, BordxClaimAmt, GroupCompany, BordxPAYSentDate, BordxPAYToStatus, BordxPAYRecDocDate, BordxPAYRecDate,
        BordxPAYDOCFRStatus, BordxMNRBSentDate, BordxMNRBRecDocDate, BordxMNRBToStatus, BordxSendACDate, BordxSendToACStatus, PAYToACPeriodMth, PAYToACPeriodYR, ACPeriodStatus,
        CASE @kind WHEN 'payor' THEN COALESCE(CAST(BordxPAYToStatus AS int), 0) WHEN 'returned' THEN COALESCE(CAST(BordxPAYDOCFRStatus AS int), 0)
            WHEN 'ac' THEN COALESCE(CAST(BordxSendToACStatus AS int), 0) WHEN 'acperiod' THEN CASE WHEN ACPeriodStatus IN ('1', 'True') THEN 1 ELSE 0 END
            ELSE COALESCE(CAST(BordxMNRBToStatus AS int), 0) END AS Done
    FROM CISWorksheet.dbo.Bordx
    WHERE (@company IS NULL OR @kind = 'mnrb' OR GroupCompany = @company)
      AND (@payCode IS NULL OR @kind <> 'mnrb' OR BordxPAYCode = @payCode)
      AND (@scan = 0 OR BordxRefNo = @bordxFrom)
      AND (@scan = 1 OR @bordxFrom IS NULL OR BordxRefNo >= @bordxFrom)
      AND (@scan = 1 OR @bordxTo IS NULL OR BordxRefNo <= @bordxTo)
      AND (@dateFrom IS NULL OR BordxDate >= @dateFrom) AND (@dateTo IS NULL OR BordxDate <= @dateTo)
      AND (@amountFrom IS NULL OR BordxClaimAmt >= @amountFrom) AND (@amountTo IS NULL OR BordxClaimAmt <= @amountTo)
      AND (@kind IN ('ac', 'acperiod') OR @monthFrom IS NULL OR PAYToACPeriodMth >= @monthFrom) AND (@kind IN ('ac', 'acperiod') OR @monthTo IS NULL OR PAYToACPeriodMth <= @monthTo)
      AND (@kind IN ('ac', 'acperiod') OR @yearFrom IS NULL OR PAYToACPeriodYR >= @yearFrom) AND (@kind IN ('ac', 'acperiod') OR @yearTo IS NULL OR PAYToACPeriodYR <= @yearTo)
      AND (@kind <> 'acperiod' OR @periodMonthFrom IS NULL OR PAYToACPeriodMth >= @periodMonthFrom)
      AND (@kind <> 'acperiod' OR @periodYearFrom IS NULL OR PAYToACPeriodYR >= @periodYearFrom)
      AND (@kind <> 'acperiod' OR @periodMonthTo IS NULL OR PAYToACPeriodMth <= @periodMonthTo)
      AND (@kind <> 'acperiod' OR @periodYearTo IS NULL OR PAYToACPeriodYR <= @periodYearTo)
      AND (@kind <> 'acperiod' OR @noPeriod = 0 OR ((PAYToACPeriodMth IS NULL OR PAYToACPeriodMth = '') AND (PAYToACPeriodYR IS NULL OR PAYToACPeriodYR = '')))
      AND (@kind <> 'ac' OR @sentFrom IS NULL OR BordxSendACDate >= @sentFrom) AND (@kind <> 'ac' OR @sentTo IS NULL OR BordxSendACDate <= @sentTo)
      AND (@kind IN ('ac', 'acperiod') OR @status IS NULL OR
           (@kind = 'payor' AND BordxPAYToStatus = @status) OR (@kind = 'returned' AND BordxPAYDOCFRStatus = @status) OR (@kind = 'mnrb' AND BordxMNRBToStatus = @status))
    OPTION (RECOMPILE);
END
GO

-- UpdateSentPayor / UpdateReturnPayor / UpdateSentAC / UpdateSentMNRB for one bordereaux. @result 1 updated, 2 returned date earlier than the sent date.
CREATE OR ALTER PROCEDURE cis.BordxDelivery_Update @kind varchar(10), @bordxNo nvarchar(20), @date datetime = NULL, @month nvarchar(2) = NULL,
    @year nvarchar(4) = NULL, @result int OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    SET @result = 1;
    IF @kind = 'payor'
        UPDATE CISWorksheet.dbo.Bordx SET BordxPayToStatus = 1, BordxPAYSentDate = @date WHERE BordxRefNo = @bordxNo;
    ELSE IF @kind = 'returned'
    BEGIN
        IF EXISTS (SELECT 1 FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @bordxNo AND BordxPAYSentDate > @date)
        BEGIN
            SET @result = 2;
            RETURN;
        END
        UPDATE CISWorksheet.dbo.Bordx SET BordxPayDocFRStatus = 1, BordxPAYRecDocDate = @date WHERE BordxRefNo = @bordxNo;
    END
    ELSE IF @kind = 'ac'
        UPDATE CISWorksheet.dbo.Bordx SET BordxSendToACStatus = 1, BordxSendACDate = @date WHERE BordxRefNo = @bordxNo;
    ELSE IF @kind = 'acperiod'
        UPDATE CISWorksheet.dbo.Bordx SET PAYToACPeriodMth = @month, ACPeriodStatus = '1', PAYToACPeriodYR = @year WHERE BordxRefNo = @bordxNo;
    ELSE
        UPDATE CISWorksheet.dbo.Bordx SET BordxMNRBSentDate = @date, BordxMNRBToStatus = 1 WHERE BordxRefNo = @bordxNo;
END
GO

-- The drop-downs: group company names (DISTINCT GRPCoName) and, for MNRB, the payors.
CREATE OR ALTER PROCEDURE cis.BordxDelivery_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT DISTINCT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM dbo.PAY;
END
GO

-- ===== main/850_bordx_monitoring.sql =====
-- Claim Bordx > Monitoring: Bordx Control (FrmBordxControl: CISWorksheet ClaimBordxModule1 case 17 copied, the pasted @strAppend replaced by
-- parameters), Outstanding Bordx (FrmOSBordx, ClaimBordxModule case 14) and Bordx & Worksheet (FrmBordxWks, case 15).
SET QUOTED_IDENTIFIER ON;
GO

-- @sent / @returned / @paid: the desktop's check boxes (1 / 0 / F / P / U = unpaid; NULL = not ticked); @payTo C / M; @entity 0 Health Connect, 1 Sihat Johor.
CREATE OR ALTER PROCEDURE cis.BordxControl_Search @company nvarchar(200) = NULL, @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL,
    @dateFrom datetime = NULL, @dateTo datetime = NULL, @sent char(1) = NULL, @returned char(1) = NULL, @paid char(1) = NULL, @subsidiary nvarchar(200) = NULL,
    @payTo char(1) = NULL, @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = N''
        + CASE WHEN @company IS NOT NULL THEN N' AND Bordx.GroupCompany = @company' ELSE N'' END
        + CASE WHEN @bordxFrom IS NOT NULL THEN N' AND Bordx.BordxRefNo >= @bordxFrom' ELSE N'' END
        + CASE WHEN @bordxTo IS NOT NULL THEN N' AND Bordx.BordxRefNo <= @bordxTo' ELSE N'' END
        + CASE WHEN @dateFrom IS NOT NULL THEN N' And Bordx.BordxDate >= @dateFrom' ELSE N'' END
        + CASE WHEN @dateTo IS NOT NULL THEN N' And Bordx.BordxDate <= @dateTo' ELSE N'' END
        + CASE @sent WHEN '1' THEN N' AND BordxPAYToStatus = ''1''' WHEN '0' THEN N' AND BordxPAYToStatus = ''0''' ELSE N'' END
        + CASE @returned WHEN '1' THEN N' AND BordxPAYDOCFrStatus = ''1''' WHEN '0' THEN N' AND BordxPAYDOCFrStatus = ''0''' ELSE N'' END
        + CASE @paid WHEN 'F' THEN N' AND BordxPAYFrStatus = ''F''' WHEN 'P' THEN N' AND BordxPAYFrStatus = ''P'''
            WHEN 'U' THEN N' AND (BordxPAYFrStatus = '''' or BordxPAYFrStatus is null)' ELSE N'' END
        + CASE WHEN @subsidiary IS NOT NULL THEN N' AND BordxSubsidary = @subsidiary' ELSE N'' END
        + CASE @payTo WHEN 'C' THEN N' AND INVPayTo = ''C''' WHEN 'M' THEN N' AND INVPayTo = ''M'' ' ELSE N'' END
        + CASE @entity WHEN 0 THEN N' AND Substring(CISCas.CasNumber, 1, 2) <>''SJ'' ' WHEN 1 THEN N' AND Substring(CISCas.CasNumber, 1, 2) =''SJ'' ' ELSE N'' END;

			       declare @sql nvarchar(max)
				   declare @groupBy nvarchar(max)

		                    set @sql =  N' SELECT 
												Bordx.BordxRefNo, 
												Bordx.BordxDate, 
												Bordx.BordxPAYDOCFRStatus, 
												Bordx.BordxPAYToStatus, 
												Bordx.BordxPAYSentDate, 
												Bordx.BordxPAYRecDate, 
												Bordx.BordxClaimAmt, 
												COALESCE(Claim.CLMTotalActual,0) AS BillAmt, 
												Bordx.BordxPAYCode, 
												PAY.HLTCode, 
												Bordx.BordxPAYRecDocDate, 
												Bordx.BordxPAYFRStatus, 
												Bordx.GroupCompany, 
												COALESCE(Claim.CLMNonGPFee,0) AS CLMNonGpFee,
												Bordx.GroupCompany as MBMGroupCompany,
												ClaimInvoice.INVPayTo,
												CASE WHEN Pln.PLNBTBG = 1 THEN COALESCE(Claim.CLMBordxAmt,0) - COALESCE(Claim.CLMBTBGAmt,0) ELSE COALESCE(Claim.CLMFinalApprove,0) END as CLMFinalApprove,
												COALESCE(Claim.CLMMCOProcFee,0) AS CLMMCOProcFee, 
												COALESCE(Claim.CLMCoPayCLNColl,0) AS CLMCoPayCLNColl,
												CASE WHEN (PLN.PLNBTBG=1 AND substring(Bordx.BordxRefNo,1,3) IN (''KDC'',''KCN'')) then COALESCE(dbo.Claim.CLMBTBGAmt,0) ELSE 
												CASE WHEN PLN.PLNBTBG=1 THEN COALESCE(dbo.Claim.CLMBordxAmt,0)-COALESCE(Claim.CLMBTBGAmt,0) ELSE COALESCE(dbo.Claim.CLMBordxAmt,0) END END AS CLMBordxAmt,
												Bordx.BordxNoClaims as BordxNoClaims,
												Substring(CISCas.CasNumber, 1, 2) AS checkcase,Bordx.BordxSubsidary as Subsidiary
												into #tempcontrol
												FROM CISDB..CISCas 
												INNER JOIN CISWorksheet..Claim  ON CISDB.dbo.CISCas.CASNumber = dbo.Claim.CASNumber 
												INNER JOIN dbo.ClaimInvoice  ON dbo.Claim.CASNumber = dbo.ClaimInvoice.CASNumber                 
												INNER JOIN dbo.Bordx   ON ((Claim.BordxRefNo = Bordx.BordxRefNo) OR (Claim.BTBGBordxRefNo = Bordx.BordxRefNo))
												INNER JOIN HISDB..PAY  ON HISDB.dbo.PAY.PAYCode = dbo.Bordx.BordxPAYCode
												INNER JOIN hisdb..MBMCLNOthers  on MBMCLNOthers.MBMnumber=ciscas.MBMnumber
												INNER JOIN cisdb..pln  on pln.PLNCode=MBMCLNOthers.PLNType    
												INNER JOIN HISDB..MBMOthers ON CISCas.MBMnumber = MBMOthers.MBMNumber where 0=0  '
												
												set @groupBy = '  Select BordxRefNo, BordxDate, BordxPAYDOCFRStatus, BordxPAYToStatus, BordxPAYSentDate, 
												BordxPAYRecDate, BordxClaimAmt, SUM(BillAmt) as BillAmt, BordxPAYCode, HLTCode, 
												BordxPAYRecDocDate, BordxPAYFRStatus, GroupCompany, CLMNonGpFee,
												MBMGroupCompany,INVPayTo,SUM(CLMFinalApprove) as CLMFinalApprove,sum(CLMCoPayCLNColl) as CLMCoPayCLNColl,sum(CLMBordxAmt) as CLMBordxAmt ,
												BordxNoClaims,
												SUM(CLMMCOProcFee) as CLMMCOProcFee, DateDiff(d, BordxPAYSentDate, BordxPAYRecDocDate) AS NoofDay,DateDiff(d, BordxPAYSentDate, BordxPAYRecDate) AS NoofDay2
												,checkcase,Subsidiary from #tempcontrol
												GROUP BY BordxRefNo, BordxDate, BordxPAYDOCFRStatus, BordxPAYToStatus, BordxPAYSentDate, 
												BordxPAYRecDate, BordxClaimAmt,  BordxPAYCode, HLTCode, 
												BordxPAYRecDocDate, BordxPAYFRStatus, GroupCompany, CLMNonGpFee,
												MBMGroupCompany,INVPayTo,BordxNoClaims
												,checkcase,Subsidiary

												drop table #tempcontrol '
					--set @sql = @sql + @strAppend
		set @sql = @sql + @filter + @groupBy
					--print @sql
		exec CISWorksheet.sys.sp_executesql @sql, N'@company nvarchar(200), @bordxFrom nvarchar(20), @bordxTo nvarchar(20), @dateFrom datetime, @dateTo datetime, @subsidiary nvarchar(200)',
            @company, @bordxFrom, @bordxTo, @dateFrom, @dateTo, @subsidiary;
END
GO

CREATE OR ALTER PROCEDURE cis.BordxOs_Search @company nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 30000 BordxRefNo, BordxPAYCode, GroupCompany, BordxPAYSentDate, BordxClaimAmt, BordxPAYToStatus
    FROM CISWorksheet.dbo.VIEW_OS_CISBordx
    WHERE BordxPAYToStatus = 1 AND (BordxPAYFRStatus = 'P' OR BordxPAYFRStatus IS NULL) AND (@company IS NULL OR GroupCompany = @company);
END
GO

CREATE OR ALTER PROCEDURE cis.BordxWks_Search @company nvarchar(200) = NULL, @payCode nvarchar(10) = NULL, @bordxFrom nvarchar(20) = NULL,
    @bordxTo nvarchar(20) = NULL, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL, @receiptFrom nvarchar(30) = NULL, @receiptTo nvarchar(30) = NULL,
    @recFrom datetime = NULL, @recTo datetime = NULL, @noRecDate bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 30000 CASNumber, BordxPAYCode, BordxClaimAmt, BordxDate, BordxRefNo, CLMTotalApproved, ClaimOpendate, BordxPAYRecDate, ReceiptNoPAY, GroupCompany
    FROM CISWorksheet.dbo.VIEW_CISBordx_Wks
    WHERE (@company IS NULL OR GroupCompany = @company) AND (@payCode IS NULL OR BordxPAYCode = @payCode)
      AND (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
      AND (@caseFrom IS NULL OR CASNumber >= @caseFrom) AND (@caseTo IS NULL OR CASNumber <= @caseTo)
      AND (@receiptFrom IS NULL OR ReceiptNoPAY >= @receiptFrom) AND (@receiptTo IS NULL OR ReceiptNoPAY <= @receiptTo)
      AND (@recFrom IS NULL OR BordxPAYRecDate >= @recFrom) AND (@recTo IS NULL OR BordxPAYRecDate <= @recTo)
      AND (@noRecDate = 0 OR BordxPAYRecDate IS NULL OR BordxPAYRecDate = '')
    OPTION (RECOMPILE);
END
GO

-- FrmBordxControl's BindSubsidary (SearchSubsidiary of the group company name) and the drop-downs.
CREATE OR ALTER PROCEDURE cis.BordxMonitoring_Lookups @company nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT DISTINCT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM dbo.PAY;
    SELECT DISTINCT LTRIM(RTRIM(Subsidiary)) AS Name FROM cis.Hdb_MBMOthers
    WHERE @company IS NOT NULL AND MBMGroupCompany = @company AND COALESCE(LTRIM(RTRIM(Subsidiary)), '') <> '';
END
GO

-- ===== main/860_print_bordx.sql =====
-- Claim Bordx > Print Bordx (frmPrintBordxExcel and the client variants): the bordereaux header (GetBordxInfo), its pay-advice rows
-- (CISWorksheet CLM_PayAdviceCLNBSN case 1, called with the bordx no. quoted here), the group summary (VIEW_OP_GRPSummaryALL) and UpdateBordx.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.PrintBordx_Info @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 BordxRefNo, BordxDate, GroupCompany, RptTypeByDept, BordxPrintedStatus, BORDXPrintMonth, BORDXPrintYear, BordxClaimAmt, BordxNoClaims
    FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @bordxNo;
END
GO

CREATE OR ALTER PROCEDURE cis.PrintBordx_Rows @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @append nvarchar(200) = N' CISWorksheet.dbo.Claim.BordxRefNo = ''' + REPLACE(@bordxNo, '''', '''''') + N'''';
    EXEC CISWorksheet.dbo.CLM_PayAdviceCLNBSN @strAppend = @append, @strAppendCase = 1;
END
GO

-- @variant: the client screens' summary sources (Prasarana View_CISBordxSummary_SPNB, AXA Affin / Time Dotcom VIEW_OP_AXASummary, BSN View_OP_BSNSummary,
-- Berjaya Sompo / Crystal View_CIS_PrintBordx_Summary); standard and SEB / Indah Water VIEW_OP_GRPSummaryALL.
CREATE OR ALTER PROCEDURE cis.PrintBordx_Summary @bordxNo nvarchar(20), @variant varchar(20) = 'standard' AS
BEGIN
    SET NOCOUNT ON;
    IF @variant = 'prasarana' SELECT * FROM CISWorksheet.dbo.View_CISBordxSummary_SPNB WHERE BordxRefNo = @bordxNo;
    ELSE IF @variant IN ('axa-affin', 'time-dotcom') SELECT * FROM dbo.VIEW_OP_AXASummary WHERE BordxRefNo = @bordxNo;
    ELSE IF @variant = 'bsn' SELECT * FROM dbo.View_OP_BSNSummary WHERE BordxRefNo = @bordxNo;
    ELSE IF @variant IN ('berjaya-sompo', 'crystal') SELECT * FROM CISWorksheet.dbo.View_CIS_PrintBordx_Summary WHERE BordxRefNo = @bordxNo;
    ELSE SELECT * FROM dbo.VIEW_OP_GRPSummaryALL WHERE BordxRefNo = @bordxNo;
END
GO

-- UpdateBordx (Print, not Preview).
CREATE OR ALTER PROCEDURE cis.PrintBordx_Update @bordxNo nvarchar(20), @month nvarchar(20), @year nvarchar(4), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISWorksheet.dbo.Bordx SET BordxPrintedDate = CAST(GETDATE() AS date), BORDXPrintMonth = @month, BORDXPrintYear = @year, BordxPrintedStatus = '1',
        BordxPrintedBy = @user
    WHERE BordxRefNo = @bordxNo;
END
GO

-- ===== main/865_bordx_invoice.sql =====
-- Claim Bordx > Insert Invoice (frmWSBordx2): VIEW_CISINV_PendingBordx with the pasted criteria replaced by parameters, and the bordx of the ticked invoices.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxInv_Search @regFrom datetime = NULL, @regTo datetime = NULL, @payCode nvarchar(10) = NULL, @grpCode nvarchar(10) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM dbo.VIEW_CISINV_PendingBordx
    WHERE (@regFrom IS NULL OR RegisterDate >= @regFrom) AND (@regTo IS NULL OR RegisterDate <= @regTo)
      AND (@payCode IS NULL OR GRPCoPaycode = @payCode) AND (@grpCode IS NULL OR GRPCoCode = @grpCode)
    OPTION (RECOMPILE);
END
GO

-- The not-yet-bordxed claims behind one listed invoice row.
CREATE OR ALTER PROCEDURE cis.BordxInv_Cases @invoiceNo nvarchar(50), @invoiceDate datetime = NULL, @registerDate datetime = NULL, @grpCode nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(k.CASNumber) AS CASNumber, k.CLMFinalApprove
    FROM dbo.CISCas c
    JOIN CISWorksheet.dbo.Claim k ON k.CASNumber = c.CASNumber AND k.ClaimBordxStatus = 0
    JOIN cis.Hdb_MBMOthers o ON o.MBMNumber = c.MBMnumber
    JOIN dbo.GRPCompany g ON g.GRPCoName = o.MBMGroupCompany AND g.GRPCoCode = @grpCode
    WHERE c.InvoiceNo = @invoiceNo
      AND ((@invoiceDate IS NULL AND c.InvoiceDate IS NULL) OR c.InvoiceDate = @invoiceDate)
      AND ((@registerDate IS NULL AND c.RegisterDate IS NULL) OR c.RegisterDate = @registerDate);
END
GO

-- UpdateRecord for one claim.
CREATE OR ALTER PROCEDURE cis.BordxInv_UpdateCase @casNumber nvarchar(30), @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISWorksheet.dbo.Claim SET BordxRefNo = @bordxNo, ClaimBordxStatus = 1, CLMPending = 'N' WHERE CASNumber = @casNumber;
    UPDATE CISWorksheet.dbo.ClaimInvoice SET INVBordxRefNo = @bordxNo, INVBordxed = 1 WHERE CASNumber = @casNumber;
    UPDATE dbo.CISCas SET Pending = 'N' WHERE CASNumber = @casNumber;
END
GO

-- SaveCLMBordx (GroupCompany holds the combo text "code - name", as the desktop).
CREATE OR ALTER PROCEDURE cis.BordxInv_SaveBordx @bordxNo nvarchar(20), @payCode nvarchar(4), @company nvarchar(200), @user nvarchar(50), @amount money, @claims int AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISWorksheet.dbo.Bordx (BordxPAYCode, BordxRefNo, BordxDate, BordxBy, BordxClaimAmt, BordxNoClaims, GroupCompany)
    VALUES (@payCode, @bordxNo, CAST(GETDATE() AS date), @user, @amount, @claims, @company);
END
GO

-- Combos: group companies (with their payor) and payors.
CREATE OR ALTER PROCEDURE cis.BordxInv_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoCode) AS Code, RTRIM(GRPCoName) AS Name, RTRIM(GRPCoPaycode) AS PayCode FROM dbo.GRPCompany ORDER BY GRPCoCode;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM dbo.PAY ORDER BY PAYCode;
END
GO

-- ===== main/866_bordx_bsn_summary.sql =====
-- Claim Bordx > Print Bordx > BSN-PrintSummary (frmBSNPrintSummary): CISWorksheet SP_Bordxdetails case 1 (BANK SIMPANAN NASIONAL bordereaux) for one or more
-- bordx numbers (@bordxNos comma-separated).
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxBsn_Summary @bordxNos nvarchar(max) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @x xml = CAST('<c>' + REPLACE(REPLACE(REPLACE(REPLACE(@bordxNos, '&', ''), '<', ''), ' ', ''), ',', '</c><c>') + '</c>' AS xml);
    DECLARE @b TABLE (BordxRefNo nvarchar(20) PRIMARY KEY);
    INSERT @b SELECT DISTINCT t.c.value('.', 'nvarchar(20)') FROM @x.nodes('/c') t(c) WHERE t.c.value('.', 'nvarchar(20)') <> '';
    SELECT Bordx.BordxRefNo, COUNT(*) AS NOOFCLAIMS, SUM(Claim.CLMBordxAmt) AS CLMBordxAmt, Bordx.BordxDate
    FROM CISWorksheet.dbo.Bordx
    INNER JOIN CISWorksheet.dbo.Claim ON Claim.BordxRefNo = Bordx.BordxRefNo
    INNER JOIN dbo.CISCas ON CISCas.CASNumber = Claim.CASNumber
    INNER JOIN cis.Hdb_MBMOthers MBMOthers ON MBMOthers.MBMNumber = CISCas.MBMnumber
    WHERE MBMOthers.MBMGroupCompany = 'BANK SIMPANAN NASIONAL' AND Bordx.BordxRefNo IN (SELECT BordxRefNo FROM @b)
    GROUP BY Bordx.BordxRefNo, Bordx.BordxDate
    ORDER BY Bordx.BordxRefNo ASC, Bordx.BordxDate ASC;
END
GO

-- ===== main/900_payment_clinic.sql =====
-- Payment > To Clinic (frmPaymentClinic): the payment list (CISWorksheet ClaimBordxModule case 26 copied, the pasted criteria as parameters),
-- GenerateSPVNo (PVSeq2013) and UpdateRecord (Payment2CLN).
SET QUOTED_IDENTIFIER ON;
GO

-- @pvNo set: the list of one PV (after a save / update); otherwise the search. @payment 1 paid (F), 2 unpaid, 3 partial, 4 partial or paid;
-- @blank*: the "empty" check boxes; @entity 0 Health Connect, 1 Sihat Johor.
CREATE OR ALTER PROCEDURE cis.PayClinic_Search @pvNo nvarchar(30) = NULL, @company nvarchar(200) = NULL, @clinic nvarchar(20) = NULL, @payment int = 0,
    @scan bit = 0, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL, @invoiceNo nvarchar(50) = NULL, @blankInvoiceNo bit = 0,
    @invFrom datetime = NULL, @invTo datetime = NULL, @blankBatch bit = 0, @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @blankInvDate bit = 0,
    @amtFrom nvarchar(20) = NULL, @amtTo nvarchar(20) = NULL, @blankAmt bit = 0, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL, @blankPv bit = 0,
    @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL, @blankPvDate bit = 0, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @blankMth bit = 0,
    @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL, @blankYr bit = 0, @sunFrom nvarchar(30) = NULL, @sunTo nvarchar(30) = NULL, @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = CASE WHEN @pvNo IS NOT NULL THEN N' Where 0=0 and Payment2CLN.PVCLNNo = @pvNo' ELSE N' Where ClaimInvoice.INVPayTo = ''C'' AND CISCas.CASStatus=''O'' '
        + CASE WHEN @company IS NOT NULL THEN N' AND MBMOthers.MBMGroupCompany = @company' ELSE N'' END
        + CASE WHEN @clinic IS NOT NULL THEN N' AND CISCas.ClinicCode = @clinic' ELSE N'' END
        + CASE @payment WHEN 1 THEN N' and Payment2CLN.PVPayCLNStatus = ''F''' WHEN 2 THEN N' and COALESCE(Payment2CLN.PVPayCLNStatus , '''')='''' '
            WHEN 3 THEN N' and Payment2CLN.PVPayCLNStatus = ''P'' ' WHEN 4 THEN N' and (Payment2CLN.PVPayCLNStatus = ''P'' or Payment2CLN.PVPayCLNStatus = ''F'') ' ELSE N'' END
        + CASE WHEN @scan = 1 THEN N' AND CISCas.CASNumber = @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseFrom IS NOT NULL THEN N' AND CISCas.CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseTo IS NOT NULL THEN N' AND CISCas.CASNumber <= @caseTo' ELSE N'' END
        + CASE WHEN @invoiceNo IS NOT NULL THEN N' and CISCas.InvoiceNo = @invoiceNo' ELSE N'' END
        + CASE WHEN @blankInvoiceNo = 1 THEN N' AND COALESCE(CISCas.InvoiceNo, '''')='''' ' ELSE N'' END
        + CASE WHEN @invFrom IS NOT NULL THEN N' And CISCas.InvoiceDate >= @invFrom' ELSE N'' END
        + CASE WHEN @invTo IS NOT NULL THEN N' And CISCas.InvoiceDate <= @invTo' ELSE N'' END
        + CASE WHEN @blankBatch = 1 THEN N' AND COALESCE(CISCas.OPBatchNo, '''')='''' ' ELSE N'' END
        + CASE WHEN @opFrom IS NOT NULL THEN N' AND CISCas.OPBatchNo >= @opFrom' ELSE N'' END
        + CASE WHEN @opTo IS NOT NULL THEN N' AND CISCas.OPBatchNo <= @opTo' ELSE N'' END
        + CASE WHEN @blankInvDate = 1 THEN N' AND COALESCE(CISCas.InvoiceDate, '''')='''' ' ELSE N'' END
        + CASE WHEN @amtFrom IS NOT NULL THEN N' and CISCas.InvoiceStmtAmt >= @amtFrom' ELSE N'' END
        + CASE WHEN @amtTo IS NOT NULL THEN N' and CISCas.InvoiceStmtAmt <= @amtTo' ELSE N'' END
        + CASE WHEN @blankAmt = 1 THEN N' AND COALESCE(CISCas.InvoiceStmtAmt , '''')='''' ' ELSE N'' END
        + CASE WHEN @pvFrom IS NOT NULL THEN N' and Payment2CLN.PVCLNNo >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' and Payment2CLN.PVCLNNo <= @pvTo' ELSE N'' END
        + CASE WHEN @blankPv = 1 THEN N' AND COALESCE(Payment2CLN.PVCLNNo ,'''')='''' ' ELSE N'' END
        + CASE WHEN @pvDateFrom IS NOT NULL THEN N' And Payment2CLN.PVCLNDate >= @pvDateFrom' ELSE N'' END
        + CASE WHEN @pvDateTo IS NOT NULL THEN N' And Payment2CLN.PVCLNDate <= @pvDateTo' ELSE N'' END
        + CASE WHEN @blankPvDate = 1 THEN N' AND COALESCE(Payment2CLN.PVCLNDate, '''')='''' ' ELSE N'' END
        + CASE WHEN @mthFrom IS NOT NULL THEN N' and Payment2CLN.CLNAccPeriodMth >= @mthFrom' ELSE N'' END
        + CASE WHEN @mthTo IS NOT NULL THEN N' and Payment2CLN.CLNAccPeriodMth <= @mthTo' ELSE N'' END
        + CASE WHEN @blankMth = 1 THEN N' AND COALESCE(Payment2CLN.CLNAccPeriodMth,'''')='''' ' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL AND @yrFrom IS NOT NULL THEN N' and Payment2CLN.CLNAccPeriodYr >= @yrFrom' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL THEN N' and Payment2CLN.CLNAccPeriodYr <= @yrTo' ELSE N'' END
        + CASE WHEN @blankYr = 1 THEN N' AND COALESCE(Payment2CLN.CLNAccPeriodYr,'''')='''' ' ELSE N'' END
        + CASE WHEN @sunFrom IS NOT NULL THEN N' and SUNPosting.CLNInvoiceBatch >= @sunFrom' ELSE N'' END
        + CASE WHEN @sunTo IS NOT NULL THEN N' and SUNPosting.CLNInvoiceBatch <= @sunTo' ELSE N'' END
        + CASE @entity WHEN 0 THEN N' AND Substring(CISCas.CasNumber, 1, 2) <>''SJ'' ' WHEN 1 THEN N' AND Substring(CISCas.CasNumber, 1, 2) =''SJ'' ' ELSE N'' END
        END;

			       declare @sql nvarchar(max)
				   declare @groupBy nvarchar(max)
				   
						set @sql =  N'   SELECT  CISCas.CASNumber, CISCas.MBMnumber, CISCas.CoverID, CISCas.DateVisit, CISCas.CASStatus, CISCas.Claimability, CISCas.Pending, CISCas.ClinicCode, CISCas.Remarks, 
												Payment2CLN.PVCLNNo, Payment2CLN.PVCLNDate, Payment2CLN.PVCLNIndex, Payment2CLN.PVCLNChequeDate, Payment2CLN.PVCLNChequeNo, Payment2CLN.PVPayCLNStatus, Claim.CLMPayableByMedix2C, 
												Payment2CLN.CLNPaidAmt, Claim.BordxRefNo, Payment2CLN.CLNAccPeriodMth, Payment2CLN.CLNAccPeriodYr, CLN.CLNName, ClaimInvoice.INVPayTo, ClaimInvoice.INVEnteredBy, ClaimInvoice.BillNo, CLN.CLNAdd1, 
												CLN.CLNAdd2, CLN.CLNAdd3, ClaimInvoice.BillDate, CLN.CLNState, CLN.CLNCity, CLN.CLNPostCode, Bordx.GroupCompany, CISCas.InvoiceNo, CISCas.InvoiceDate, Payment2CLN.PVCLNBank, Payment2CLN.PVCLNChequeAmt, 
												Payment2CLN.PVCLNAccount, Payment2CLN.CLNRemark, CISCas.OPBatchNo, CISCas.InvoiceStmtAmt, CLN.ClinicGrpCode, Claim.CLMNonGPFee, Claim.CLMTotalActual, Claim.CLMMCOProcFee,MBMOthers.MBMGroupCompany, GRPCompany.GRPCoCode, 
												SUNPosting.PostDate AS SUNPostDate, SUNPosting.CLNInvoiceBatch AS SUNInvNo, Claim.CLMCoPayCLNColl, Claim.CLMBordxAmt, Claim.CLMFinalApprove
												,CASE CISCAS.CoverID WHEN ''00'' THEN mbm.MBMName ELSE MBMCoveredPersons.MBMCName END AS PatientName
												FROM CISDB..CISCas     
												INNER JOIN  CISWorksheet..Claim ON CISCas.CASNumber = Claim.CASNumber
												INNER JOIN CISWorksheet..ClaimInvoice  ON ClaimInvoice.CASNumber = Claim.CASNumber 
												INNER JOIN HISDB..CLN  ON CLN.CLNCode = CISCas.ClinicCode 
												INNER JOIN HISDB..MBMOthers ON CISCas.MBMnumber = MBMOthers.MBMNumber 
												INner JOin HISDB..mbm ON mbm.mbmnumber=CISCas.MBMNumber
												LEFT OUTER JOIN HISDB..MBMCoveredPersons ON MBMCoveredPersons.MBMCNumber = CISCas.MBMnumber AND MBMCoveredPersons.MBMCCoverID = CISCas.CoverID 
												LEFT JOIN CISWorksheet..SUNPosting ON CISCas.CASNumber = SUNPosting.InvoiceNo AND CISCas.ClinicCode = SUNPosting.TransRef 
												LEFT JOIN CISDB..GRPCompany ON MBMOthers.MBMGroupCompany = GRPCompany.GRPCoName 
												LEFT JOIN cispayment..Payment2CLN  ON Payment2CLN.CASNumber = Claim.CASNumber
												LEFT JOIN CISWorksheet..Bordx ON Bordx.BordxRefNo = Claim.BordxRefNo   '
		           
			       Set @groupBy = ' GROUP BY CISCas.CASNumber, CISCas.MBMnumber, CISCas.CoverID, CISCas.DateVisit, CISCas.CASStatus, CISCas.Claimability, CISCas.Pending, CISCas.ClinicCode, CISCas.Remarks, 
												Payment2CLN.PVCLNNo, Payment2CLN.PVCLNDate, Payment2CLN.PVCLNIndex, Payment2CLN.PVCLNChequeDate, Payment2CLN.PVCLNChequeNo, Payment2CLN.PVPayCLNStatus, Claim.CLMPayableByMedix2C, 
												Payment2CLN.CLNPaidAmt, Claim.BordxRefNo, Payment2CLN.CLNAccPeriodMth, Payment2CLN.CLNAccPeriodYr, CLN.CLNName, ClaimInvoice.INVPayTo, ClaimInvoice.INVEnteredBy, ClaimInvoice.BillNo,
												CLN.CLNAdd1, CLN.CLNAdd2, CLN.CLNAdd3, ClaimInvoice.BillDate, CLN.CLNState, CLN.CLNCity, CLN.CLNPostCode, Bordx.GroupCompany, CISCas.InvoiceNo, CISCas.InvoiceDate, Payment2CLN.PVCLNBank, 
												Payment2CLN.PVCLNChequeAmt, Payment2CLN.PVCLNAccount, Payment2CLN.CLNRemark, CISCas.OPBatchNo, CISCas.InvoiceStmtAmt, CLN.ClinicGrpCode, Claim.CLMNonGPFee, Claim.CLMTotalActual, Claim.CLMMCOProcFee, MBMOthers.MBMGroupCompany, 
												GRPCompany.GRPCoCode, SUNPosting.PostDate, SUNPosting.CLNInvoiceBatch, Claim.CLMCoPayCLNColl, Claim.CLMBordxAmt, Claim.CLMFinalApprove,mbm.MBMName,MBMCoveredPersons.MBMCName '
				   
		set @sql = @sql + @filter + @groupBy
				   			   
		exec CISWorksheet.sys.sp_executesql @sql, N'@company nvarchar(200), @clinic nvarchar(20), @caseFrom nvarchar(30), @caseTo nvarchar(30), @invoiceNo nvarchar(50),
            @invFrom datetime, @invTo datetime, @opFrom nvarchar(30), @opTo nvarchar(30), @amtFrom nvarchar(20), @amtTo nvarchar(20), @pvFrom nvarchar(30), @pvTo nvarchar(30),
            @pvDateFrom datetime, @pvDateTo datetime, @mthFrom nvarchar(2), @mthTo nvarchar(2), @yrFrom nvarchar(4), @yrTo nvarchar(4), @sunFrom nvarchar(30), @sunTo nvarchar(30),
            @pvNo nvarchar(30)',
            @company, @clinic, @caseFrom, @caseTo, @invoiceNo, @invFrom, @invTo, @opFrom, @opTo, @amtFrom, @amtTo, @pvFrom, @pvTo, @pvDateFrom, @pvDateTo,
            @mthFrom, @mthTo, @yrFrom, @yrTo, @sunFrom, @sunTo, @pvNo;
END
GO

CREATE OR ALTER PROCEDURE cis.PayClinic_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(BankCode) AS Code, RTRIM(BankDesc) AS Name, RTRIM(PVCode) AS PvCode FROM cis.Maint_Banks;
    SELECT DISTINCT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL;
    SELECT RTRIM(CLNCode) AS Code, RTRIM(CLNName) AS Name FROM cis.Hdb_CLN WHERE LEN(CLNName) > 0;
END
GO

-- GenerateSPVNo: PV code + the next 9-digit number of PVSeq2013 for the bank, locked for the save.
CREATE OR ALTER PROCEDURE cis.PayClinic_NextPv @bank nvarchar(10), @pvCode nvarchar(10), @pvNo nvarchar(30) OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @n char(10);
    IF NOT EXISTS (SELECT 1 FROM CISPayment.dbo.PVSeq2013 WITH (UPDLOCK, HOLDLOCK) WHERE BankCode = @bank)
    BEGIN
        INSERT CISPayment.dbo.PVSeq2013 (BankCode, PVCode, PVSEQNO) VALUES (@bank, @pvCode, '100000001');
        SET @n = '100000001';
    END
    ELSE
        UPDATE TOP (1) CISPayment.dbo.PVSeq2013 SET @n = PVSEQNO = RIGHT('000000000' + CAST(CAST(PVSEQNO AS bigint) + 1 AS varchar(12)), 9) WHERE BankCode = @bank;
    SET @pvNo = RTRIM(@pvCode) + RTRIM(@n);
END
GO

-- CmdPrintPV (@pvNo) / PreviewPV (@cases, comma list): View_Payment2CLN rows with the receipt no. and the HQ clinic (ClinicHQInfo, preview only).
CREATE OR ALTER PROCEDURE cis.PayClinic_PvRows @pvNo nvarchar(30) = NULL, @cases nvarchar(max) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @c TABLE (CASNumber nvarchar(30) PRIMARY KEY);
    IF @pvNo IS NOT NULL
        INSERT @c SELECT DISTINCT RTRIM(CASNumber) FROM CISPayment.dbo.Payment2CLN WHERE PVCLNNo = @pvNo;
    ELSE
    BEGIN
        DECLARE @x xml = CAST('<c>' + REPLACE(REPLACE(REPLACE(REPLACE(@cases, '&', ''), '<', ''), ' ', ''), ',', '</c><c>') + '</c>' AS xml);
        INSERT @c SELECT DISTINCT t.c.value('.', 'nvarchar(30)') FROM @x.nodes('/c') t(c) WHERE t.c.value('.', 'nvarchar(30)') <> '';
    END
    SELECT v.PVCLNIndex, v.CASNumber, v.BordxRefNo, v.BillNo, v.BillDate, v.PVCLNNo, v.PVCLNChequeNo, v.PVCLNDate, v.CLMTotalActual, v.CLMMCOProcFee,
        v.ClinicGrpCode, v.ClinicCode, v.CLNName, v.CLNAdd1, v.CLNAdd2, v.CLNAdd3, v.CLNPostCode, v.CLNCity, v.CLNState, v.CLNPaidAmt, v.CLMPayableByMedix2C,
        (SELECT TOP 1 r.ReceiptNo FROM CISPayment.dbo.REceiptFrPAY r WHERE r.CASNumber = v.CASNumber) AS ReceiptNo,
        hq.CLNCode AS HqCode, hq.CLNName AS HqName, hq.CLNAdd1 AS HqAdd1, hq.CLNAdd2 AS HqAdd2, hq.CLNAdd3 AS HqAdd3, hq.CLNPostCode AS HqPostCode,
        hq.CLNCity AS HqCity, hq.CLNState AS HqState, v.InvoiceNo, v.InvoiceDate, v.CLMFinalApprove, v.PVCLNChequeAmt, v.CLNAccPeriodMth, v.CLNAccPeriodYr,
        p.CLNPName, p.CLNPAdd1, p.CLNPAdd2, p.CLNPAdd3
    FROM CISPayment.dbo.View_Payment2CLN v
    LEFT JOIN cis.Hdb_CLN p ON p.CLNCode = v.ClinicCode
    LEFT JOIN cis.Hdb_CLN hq ON @pvNo IS NULL AND COALESCE(RTRIM(v.ClinicGrpCode), '') <> '' AND hq.CLNCode = RTRIM(v.ClinicGrpCode)
    WHERE v.CASNumber IN (SELECT CASNumber FROM @c) AND (@pvNo IS NULL OR v.PVCLNNo = @pvNo);
END
GO

-- UpdateRecord: the Payment2CLN row of @index is overwritten (edit, or not partial), or a new row is added (no row yet, or a partial payment's next instalment).
CREATE OR ALTER PROCEDURE cis.PayClinic_Save @index nvarchar(20) = NULL, @edit bit = 0, @casNumber nvarchar(30), @pvNo nvarchar(30), @pvDate datetime,
    @chequeDate datetime = NULL, @chequeNo nvarchar(30) = NULL, @chequeAmt money, @bank nvarchar(10), @account nvarchar(50) = NULL, @remark nvarchar(500) = NULL,
    @paidAmt money, @mth nvarchar(2), @yr nvarchar(4), @status char(1), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @old char(1), @found bit = 0;
    SELECT TOP 1 @old = RTRIM(PVPayCLNStatus), @found = 1 FROM CISPayment.dbo.Payment2CLN WHERE CAST(PVCLNIndex AS nvarchar(20)) = @index;
    IF @found = 1 AND @edit = 0 AND @old = 'P'
    BEGIN
        UPDATE CISPayment.dbo.Payment2CLN SET PVPayCLNStatus = @status WHERE CAST(PVCLNIndex AS nvarchar(20)) = @index;
        SET @found = 0;
    END
    IF @found = 1
        UPDATE CISPayment.dbo.Payment2CLN SET CASNumber = @casNumber, PVCLNNo = @pvNo, PVCLNDate = @pvDate, PVCLNChequeDate = @chequeDate,
            PVCLNChequeNo = COALESCE(@chequeNo, ''), PVCLNChequeAmt = @chequeAmt, PVCLNBank = @bank, PVCLNAccount = COALESCE(@account, ''),
            CLNRemark = COALESCE(@remark, ''), CLNPaidAmt = @paidAmt, CLNAccPeriodMth = @mth, CLNAccPeriodYr = @yr, Pay2CLNLastUpdateDate = CAST(GETDATE() AS date),
            Pay2CLNLastUpdateUser = @user, PVPayCLNStatus = @status
        WHERE CAST(PVCLNIndex AS nvarchar(20)) = @index;
    ELSE
        INSERT CISPayment.dbo.Payment2CLN (CASNumber, PVCLNNo, PVCLNDate, PVCLNChequeDate, PVCLNChequeNo, PVCLNChequeAmt, PVCLNBank, PVCLNAccount, CLNRemark,
            CLNPaidAmt, CLNAccPeriodMth, CLNAccPeriodYr, Pay2CLNLastUpdateDate, Pay2CLNLastUpdateUser, PVPayCLNStatus)
        VALUES (@casNumber, @pvNo, @pvDate, @chequeDate, COALESCE(@chequeNo, ''), @chequeAmt, @bank, COALESCE(@account, ''), COALESCE(@remark, ''), @paidAmt, @mth, @yr,
            CAST(GETDATE() AS date), @user, @status);
END
GO

-- ===== main/910_payment_clinic_invoice.sql =====
-- Payment > To Clinic - Invoice (frmPaymentCLNInvoice, frmPrintPVRange, frmPrintChequeRange, frmInvoiceDetPopup): CISPayment SP_CISPayment
-- 'PaymentCLNV2' and 'Pay2CLN' copied with the pasted criteria replaced by parameters; the save's case list, the cheques and their print status.
SET QUOTED_IDENTIFIER ON;
GO

-- The list (SearchRecord): one row per clinic / invoice month / year / PV. @pvNo..@pvLast: the PVs just saved / updated (Payment2CLNListing by PV).
-- @payment 1 paid, 2 unpaid; @blank*: the "empty" boxes; @entity 0 Health Connect, 1 Sihat Johor.
CREATE OR ALTER PROCEDURE cis.PayClnInv_Search @pvNo nvarchar(30) = NULL, @pvLast nvarchar(30) = NULL, @clinic nvarchar(20) = NULL, @yrFrom nvarchar(4) = NULL,
    @yrTo nvarchar(4) = NULL, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @payment int = 0, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL,
    @blankPv bit = 0, @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL, @blankPvDate bit = 0, @acMthFrom nvarchar(2) = NULL, @acMthTo nvarchar(2) = NULL,
    @blankAcMth bit = 0, @acYrFrom nvarchar(4) = NULL, @acYrTo nvarchar(4) = NULL, @blankAcYr bit = 0, @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = CASE WHEN @pvNo IS NOT NULL THEN N' PVCLNNo >= @pvNo and PVCLNNo <= COALESCE(@pvLast, @pvNo)' ELSE N''
        + N' SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> ''EM'' and SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> ''ES'' '
        + CASE WHEN @clinic IS NOT NULL THEN N' And ClinicCode = @clinic' ELSE N'' END
        + CASE WHEN @yrFrom IS NOT NULL THEN N' and CASInvYear >= @yrFrom' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL THEN N' and CASInvYear <= @yrTo' ELSE N'' END
        + CASE WHEN @mthFrom IS NOT NULL THEN N' and CASInvMonth >= @mthFrom' ELSE N'' END
        + CASE WHEN @mthTo IS NOT NULL THEN N' and CASInvMonth <= @mthTo' ELSE N'' END
        + CASE @payment WHEN 1 THEN N' and Payment2CLN.PVPayCLNStatus = ''F''' WHEN 2 THEN N' and (Payment2CLN.PVPayCLNStatus = '''' or Payment2CLN.PVPayCLNStatus is null) ' ELSE N'' END
        + N' and (INVPayTo  = ''C'')  and (CLMPayableByMedix2C  <> ''0'') '
        + CASE WHEN @pvFrom IS NOT NULL THEN N' and PVCLNNo >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' and PVCLNNo <= @pvTo' ELSE N'' END
        + CASE WHEN @blankPv = 1 THEN N'AND (PVCLNNo Is null OR PVCLNNo = '''') ' ELSE N'' END
        + CASE WHEN @pvDateFrom IS NOT NULL THEN N' And PVCLNDate >= @pvDateFrom' ELSE N'' END
        + CASE WHEN @pvDateTo IS NOT NULL THEN N' And PVCLNDate <= @pvDateTo' ELSE N'' END
        + CASE WHEN @blankPvDate = 1 THEN N'AND (PVCLNDate Is null OR PVCLNDate = '''') ' ELSE N'' END
        + CASE WHEN @acMthFrom IS NOT NULL THEN N' and CLNAccPeriodMth >= @acMthFrom' ELSE N'' END
        + CASE WHEN @acMthTo IS NOT NULL THEN N' and CLNAccPeriodMth <= @acMthTo' ELSE N'' END
        + CASE WHEN @blankAcMth = 1 THEN N' AND CLNAccPeriodMth Is null ' ELSE N'' END
        + CASE WHEN @acYrTo IS NOT NULL AND @acYrFrom IS NOT NULL THEN N' and CLNAccPeriodYr >= @acYrFrom' ELSE N'' END
        + CASE WHEN @acYrTo IS NOT NULL THEN N' and CLNAccPeriodYr <= @acYrTo' ELSE N'' END
        + CASE @entity WHEN 0 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) <>''SJ'' ' WHEN 1 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) =''SJ'' ' ELSE N'' END
        + CASE WHEN @blankAcYr = 1 THEN N' AND CLNAccPeriodYr Is null ' ELSE N'' END
        END;
			declare @sql nvarchar(max)
			declare @groupBy nvarchar(max)
			declare @tmpINVPayto as char(1)
			set @tmpINVPayto  = 'C'
			Set @sql ='SELECT TOP (100000) CISDB.dbo.CISCas.ClinicCode, dbo.Payment2CLN.PVPayCLNStatus, SUM(CISWorksheet.dbo.Claim.CLMPayableByMedix2C) AS CLMPayableByMedix2C, 
                      dbo.Payment2CLN.PVCLNNo, dbo.Payment2CLN.PVCLNDate, dbo.Payment2CLN.PVCLNChequeDate, dbo.Payment2CLN.PVCLNChequeNo, 
                      SUM(dbo.Payment2CLN.CLNPaidAmt) AS CLNPaidAmt, dbo.Payment2CLN.PVCLNChequeAmt, dbo.Payment2CLN.PVCLNBank, dbo.Payment2CLN.PVCLNAccount, 
                      dbo.Payment2CLN.CLNRemark, SUM(CISWorksheet.dbo.Claim.CLMNonGPFee) AS CLMNonGPFee, 
                      SUM(CISWorksheet.dbo.Claim.CLMTotalActual - CISWorksheet.dbo.Claim.CLMCoPayCLNColl) AS CLMFinalApprove, SUM(CISWorksheet.dbo.Claim.CLMMCOProcFee) 
                      AS CLMMCOProcFee, dbo.Payment2CLN.PVWordingOne, dbo.Payment2CLN.PVWordingTwo, SUM(CISWorksheet.dbo.Claim.CLMTotalActual) AS CLMTotalActual, 
                      CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISDB.dbo.CISCas.CASInvMonth, CISDB.dbo.CISCas.CASInvYear,SUM(CISWorksheet.dbo.Claim.CLMCoPayCLNColl) as CoPay
					  FROM         CISDB.dbo.CISCas 
					INNER JOIN CISWorksheet..ClaimInvoice ON CISCas.CASNumber = ClaimInvoice.CASNumber 
					INNER JOIN HISDB..MBM ON CISCas.MBMnumber = MBM.MBMNumber 
					Left JOIN CISWorksheet..Claim ON CISCas.CASNumber = Claim.CASNumber 
					left JOIN CISPayment..Payment2CLN ON Payment2CLN.CASNumber = Claim.CASNumber
					  Where coalesce(CISCas.InvoiceNo,'''')<>'''' AND '
				Set @sql = @sql + ' CISCas.CASStatus =''O'' and  CISCas.Claimability = ''A'' AND  ' 

			Set @groupBy = 'GROUP BY CISDB.dbo.CISCas.ClinicCode,Payment2CLN.PVPayCLNStatus,
					Payment2CLN.PVCLNNo,Payment2CLN.PVCLNDate,Payment2CLN.PVCLNChequeDate,Payment2CLN.PVCLNChequeNo,
					Payment2CLN.PVCLNChequeAmt,Payment2CLN.PVCLNBank,Payment2CLN.PVCLNAccount,Payment2CLN.CLNRemark, 
					Payment2CLN.PVWordingOne,Payment2CLN.PVWordingTwo,CISWorksheet.dbo.ClaimInvoice.INVPayTo,
					CISDB.dbo.CISCas.CASInvMonth , CISDB.dbo.CISCas.CASInvYear '
			set @sql = @sql + @filter + @groupBy
			exec CISPayment.sys.sp_executesql @sql, N'@clinic nvarchar(20), @yrFrom nvarchar(4), @yrTo nvarchar(4), @mthFrom nvarchar(2), @mthTo nvarchar(2), @pvFrom nvarchar(30), @pvTo nvarchar(30), @pvDateFrom datetime, @pvDateTo datetime, @acMthFrom nvarchar(2), @acMthTo nvarchar(2), @acYrFrom nvarchar(4), @acYrTo nvarchar(4), @pvNo nvarchar(30), @pvLast nvarchar(30)', @clinic, @yrFrom, @yrTo, @mthFrom, @mthTo, @pvFrom, @pvTo, @pvDateFrom, @pvDateTo, @acMthFrom, @acMthTo, @acYrFrom, @acYrTo, @pvNo, @pvLast;
END
GO

-- frmInvoiceDetPopup (SearchRecordTopup: the row's clinic / invoice month / year) and frmPrintPVRange (PVCLNListing: a PV range): the Pay2CLN branch.
CREATE OR ALTER PROCEDURE cis.PayClnInv_Invoices @clinic nvarchar(20) = NULL, @yr nvarchar(4) = NULL, @mth nvarchar(2) = NULL, @pvFrom nvarchar(30) = NULL,
    @pvTo nvarchar(30) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = CASE WHEN @pvFrom IS NOT NULL OR @pvTo IS NOT NULL THEN N' 0 = 0 '
        + CASE WHEN @pvFrom IS NOT NULL THEN N' and dbo.Payment2CLN.PVCLNNo >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' and dbo.Payment2CLN.PVCLNNo <= @pvTo' ELSE N'' END
        + N' and InvoiceNo <> '''' '
        ELSE N' SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> ''EM'' and SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> ''ES'' '
        + N' and CASInvYear >= @yr and CASInvYear <= @yr and CASInvMonth >= @mth and CASInvMonth <= @mth and (CLMPayableByMedix2C  <> ''0'') and (INVPayTo  = ''C'') '
        + CASE WHEN @clinic IS NOT NULL THEN N' and CISCas.ClinicCode = @clinic' ELSE N'' END
        END;
			declare @sql nvarchar(max)
			declare @groupBy nvarchar(max)

			
			set @sql =  '   SELECT CISCas.ClinicCode, Payment2CLN.PVPayCLNStatus, SUM(Claim.CLMPayableByMedix2C) AS CLMPayableByMedix2C, Payment2CLN.PVCLNNo, Payment2CLN.PVCLNDate, 
								Payment2CLN.PVCLNChequeDate, Payment2CLN.PVCLNChequeNo, SUM(Payment2CLN.CLNPaidAmt) AS CLNPaidAmt, CISCas.InvoiceDate, Payment2CLN.PVCLNChequeAmt, 
								Payment2CLN.PVCLNBank, Payment2CLN.PVCLNAccount, Payment2CLN.CLNRemark, SUM(Claim.CLMNonGPFee) AS CLMNonGPFee, 
								SUM(coalesce(Claim.CLMTotalActual,0) - coalesce(Claim.CLMCoPayColl,0)) AS CLMFinalApprove, SUM(Claim.CLMMCOProcFee) AS CLMMCOProcFee, 
								Payment2CLN.PVWordingOne, Payment2CLN.PVWordingTwo, SUM(Claim.CLMTotalActual) AS CLMTotalActual, ClaimInvoice.INVPayTo, CISCas.InvoiceNo
								,CLN.CLNName,CLN.CLNAdd1,CLN.CLNAdd2,CLN.CLNAdd3,CLN.CLNPostCode,CLN.CLNCity,CLN.CLNState,CLN.CLNRemark,CLN.CLNBankAcc,CLN.CLNPName,CLN.CLNPAdd1,CLN.CLNPAdd2,
								CLN.CLNPAdd3,CLN.CLNType,CLN.CLNFFSStatus,CLN.CLNPName,CLN.CLNRegNo,CLN.CLNMCOProcfeePerc,CLN.CLNBankName,CLN.CLNBankAcc
								FROM  CISDB..CISCas 
								INNER JOIN CISWorksheet..ClaimInvoice ON CISCas.CASNumber = ClaimInvoice.CASNumber 
								INNER JOIN HISDB..MBM ON CISCas.MBMnumber = MBM.MBMNumber 
								LEFT JOIN CISWorksheet..Claim  ON CISCas.CASNumber = Claim.CASNumber 
								LEFT JOIN CISPayment..Payment2CLN ON Payment2CLN.CASNumber = Claim.CASNumber
								Inner Join HISDB..cln ON cln.clncode=CISCas.ClinicCode
								WHERE CISCas.CASStatus =''O'' and   CISCas.Claimability =''A'' AND coalesce(CISCas.InvoiceNo,'''')<>'''' And  '
			set @groupBy =  '  GROUP BY CISCas.ClinicCode,Payment2CLN.PVPayCLNStatus, CISCas.InvoiceNo,Payment2CLN.PVCLNNo,Payment2CLN.PVCLNDate,
								Payment2CLN.PVCLNChequeDate,Payment2CLN.PVCLNChequeNo,CISCas.InvoiceDate,Payment2CLN.PVCLNChequeAmt, 
								Payment2CLN.PVCLNBank,Payment2CLN.PVCLNAccount,Payment2CLN.CLNRemark,Payment2CLN.PVWordingOne,
								Payment2CLN.PVWordingTwo,ClaimInvoice.INVPayTo 
								,CLN.CLNName,CLN.CLNAdd1,CLN.CLNAdd2,CLN.CLNAdd3,CLN.CLNPostCode,CLN.CLNCity,CLN.CLNState,CLN.CLNRemark,CLN.CLNBankAcc,CLN.CLNPName,CLN.CLNPAdd1,CLN.CLNPAdd2,
								CLN.CLNPAdd3,CLN.CLNType,CLN.CLNFFSStatus,CLN.CLNPName,CLN.CLNRegNo,CLN.CLNMCOProcfeePerc,CLN.CLNBankName,CLN.CLNBankAcc
								order by CISCas.ClinicCode  '
			set @sql = @sql + @filter + @groupBy
			exec CISPayment.sys.sp_executesql @sql, N'@clinic nvarchar(20), @yr nvarchar(4), @mth nvarchar(2), @pvFrom nvarchar(30), @pvTo nvarchar(30)', @clinic, @yr, @mth, @pvFrom, @pvTo;
END
GO

-- CmdSave_Click: the unpaid cases of a clinic for the year and the ticked invoice months (the desktop's StrInvSql).
CREATE OR ALTER PROCEDURE cis.PayClnInv_Cases @clinic nvarchar(20), @yr nvarchar(4), @months nvarchar(200), @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT c.CASNumber, b.BordxDate, p.PVCLNNo, cl.CLMPayableByMedix2C
    FROM HISDB.dbo.MBM m
    INNER JOIN CISDB.dbo.CISCas c ON m.MBMNumber = c.MBMnumber
    INNER JOIN CISWorksheet.dbo.Claim cl ON cl.CASNumber = c.CASNumber
    LEFT OUTER JOIN CISPayment.dbo.Payment2CLN p ON cl.CASNumber = p.CASNumber
    LEFT OUTER JOIN CISWorksheet.dbo.Bordx b ON cl.BordxRefNo = b.BordxRefNo
    WHERE c.ClinicCode = @clinic AND c.InvoiceNo IS NOT NULL AND c.InvoiceNo <> N'' AND c.CASINVYear >= @yr AND c.CASINVYear <= @yr
      AND p.PVCLNNo IS NULL AND SUBSTRING(m.MBMNumber, 1, 2) <> 'EM' AND SUBSTRING(m.MBMNumber, 1, 2) <> 'ES'
      AND ',' + REPLACE(@months, ' ', '') + ',' LIKE '%,' + RTRIM(c.CASINVMonth) + ',%'
      AND ((@entity = 0 AND SUBSTRING(c.CASNumber, 1, 2) <> 'SJ') OR (@entity = 1 AND SUBSTRING(c.CASNumber, 1, 2) = 'SJ'));
END
GO

-- Payment2CLNListing's per-row CLN look-up (name, payee and its address, bank, type, FFS, reg. no., MCO %).
CREATE OR ALTER PROCEDURE cis.PayClnInv_Clinics @codes nvarchar(max) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(CLNCode) AS CLNCode, CLNName, CLNPName, CLNPAdd1, CLNPAdd2, CLNPAdd3, CLNBankAcc, CLNBankName, CLNType, CLNFFSStatus, CLNRegNo, CLNMCOProcfeePerc,
        ClinicGrpCode
    FROM cis.Hdb_CLN WHERE ',' + REPLACE(@codes, ' ', '') + ',' LIKE '%,' + RTRIM(CLNCode) + ',%';
END
GO

-- The Payment2CLN rows of a PV (CmdUpdate: status F only).
CREATE OR ALTER PROCEDURE cis.PayClnInv_PvRows @pvNo nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(PVCLNIndex AS nvarchar(20)) AS PVCLNIndex, RTRIM(CASNumber) AS CASNumber, PVPayCLNStatus, CLNPaidAmt
    FROM CISPayment.dbo.Payment2CLN WHERE PVCLNNo = @pvNo AND PVPayCLNStatus = 'F';
END
GO

-- The Payment2CLN insert / overwrite of frmPaymentCLNInvoice UpdateRecord (with the amount wording).
CREATE OR ALTER PROCEDURE cis.PayClnInv_Save @index nvarchar(20) = NULL, @casNumber nvarchar(30), @pvNo nvarchar(30), @pvDate datetime, @chequeDate datetime = NULL,
    @chequeNo nvarchar(30), @chequeAmt money, @bank nvarchar(10), @remark nvarchar(500) = NULL, @paidAmt money, @mth nvarchar(2), @yr nvarchar(4), @status char(1),
    @wordingOne nvarchar(500), @wordingTwo nvarchar(500), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.Payment2CLN WHERE CAST(PVCLNIndex AS nvarchar(20)) = @index)
        UPDATE CISPayment.dbo.Payment2CLN SET CASNumber = @casNumber, PVCLNNo = @pvNo, PVCLNDate = @pvDate, PVCLNChequeDate = @chequeDate, PVCLNChequeNo = @chequeNo,
            PVCLNChequeAmt = @chequeAmt, PVCLNBank = @bank, CLNRemark = COALESCE(@remark, ''), CLNPaidAmt = @paidAmt, CLNAccPeriodMth = @mth, CLNAccPeriodYr = @yr,
            Pay2CLNLastUpdateDate = CAST(GETDATE() AS date), Pay2CLNLastUpdateUser = @user, PVPayCLNStatus = @status, PVWordingOne = @wordingOne, PVWordingTwo = @wordingTwo
        WHERE CAST(PVCLNIndex AS nvarchar(20)) = @index;
    ELSE
        INSERT CISPayment.dbo.Payment2CLN (CASNumber, PVCLNNo, PVCLNDate, PVCLNChequeDate, PVCLNChequeNo, PVCLNChequeAmt, PVCLNBank, CLNRemark, CLNPaidAmt,
            CLNAccPeriodMth, CLNAccPeriodYr, Pay2CLNLastUpdateDate, Pay2CLNLastUpdateUser, PVPayCLNStatus, PVWordingOne, PVWordingTwo)
        VALUES (@casNumber, @pvNo, @pvDate, @chequeDate, @chequeNo, @chequeAmt, @bank, COALESCE(@remark, ''), @paidAmt, @mth, @yr, CAST(GETDATE() AS date), @user,
            @status, @wordingOne, @wordingTwo);
END
GO

-- cmdPrintChq / frmPrintChequeRange: the cheques (VIEW_Payment2CLN_Chq) and their print status.
CREATE OR ALTER PROCEDURE cis.PayClnInv_Cheques @chequeFrom nvarchar(30), @chequeTo nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT PVCLNChequeNo, PVCLNChequeDate, PVCLNChequeAmt, PVWordingOne, PVWordingTwo, CLNPName
    FROM CISPayment.dbo.VIEW_Payment2CLN_Chq WHERE PVCLNChequeNo >= @chequeFrom AND PVCLNChequeNo <= @chequeTo ORDER BY PVCLNChequeNo;
END
GO

CREATE OR ALTER PROCEDURE cis.PayClnInv_ChequesPrinted @chequeFrom nvarchar(30), @chequeTo nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISPayment.dbo.Payment2CLN SET PVChequeStatus = 'Y' WHERE PVCLNChequeNo >= @chequeFrom AND PVCLNChequeNo <= @chequeTo;
END
GO

-- ===== main/920_payment_member.sql =====
-- Payment > To Member (frmPaymentMemberAuto08 with frmPrintPVRange / frmPrintChequeRange in member mode): CISPayment SP_CISPayment_TESTING 'Pay2MBM'
-- copied with the pasted criteria as parameters, Savecispayment2mbm, the PV wording and edit updates, the PV rows and the member cheques.
SET QUOTED_IDENTIFIER ON;
GO

-- @pvNo: the PV just saved / updated; @payment 0-6 as CboPayment; @bordx / @btbg 0 / 1 (NULL = not chosen); @blank*: the "empty" boxes.
CREATE OR ALTER PROCEDURE cis.PayMbm_Search @pvNo nvarchar(30) = NULL, @caseNo nvarchar(30) = NULL, @scan bit = 0, @caseFrom nvarchar(30) = NULL,
    @caseTo nvarchar(30) = NULL, @mbmNo nvarchar(30) = NULL, @mbmName nvarchar(200) = NULL, @company nvarchar(200) = NULL, @payment int = 0,
    @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL, @blankBatch bit = 0, @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @blankPv bit = 0,
    @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL, @blankPvDate bit = 0, @chqFrom nvarchar(50) = NULL, @chqTo nvarchar(50) = NULL, @blankChq bit = 0,
    @crFrom nvarchar(30) = NULL, @crTo nvarchar(30) = NULL, @blankCr bit = 0, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @blankMth bit = 0,
    @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL, @blankYr bit = 0, @bordx int = NULL, @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL,
    @btbg int = NULL, @btbgFrom nvarchar(20) = NULL, @btbgTo nvarchar(20) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = CASE WHEN @pvNo IS NOT NULL THEN N' and PVNo >= @pvNo and PVNo <= @pvNo'
        WHEN @caseNo IS NOT NULL THEN N' And Claim.CASNumber = @caseFrom' ELSE N''
        + CASE WHEN @scan = 1 THEN N' AND CISWorksheet.dbo.ClaimInvoice.CASNumber = @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseFrom IS NOT NULL THEN N' AND CISWorksheet.dbo.ClaimInvoice.CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseTo IS NOT NULL THEN N' AND CISWorksheet.dbo.ClaimInvoice.CASNumber <= @caseTo' ELSE N'' END
        + N' And CISWorksheet.dbo.ClaimInvoice.INVPayTo = ''M'''
        + CASE WHEN @mbmNo IS NOT NULL THEN N' And CISDB.dbo.CISCas.MBMnumber = @mbmNo' ELSE N'' END
        + CASE WHEN @mbmName IS NOT NULL THEN N' And MBMName like ''%'' + @mbmName + ''%''' ELSE N'' END
        + CASE WHEN @company IS NOT NULL THEN N' AND MBMGroupCompany = @company' ELSE N'' END
        + CASE @payment WHEN 1 THEN N' and PVPayMBMStatus = ''F''' WHEN 2 THEN N' and (PVPayMBMStatus = '''' or PVPayMBMStatus is null) '
            WHEN 3 THEN N' and PVPayMBMStatus = ''P''' WHEN 4 THEN N' and (PVPayMBMStatus = ''P'' or PVPayMBMStatus = ''F'') '
            WHEN 5 THEN N' and (PVPayMBMStatus = '''' or PVPayMBMStatus is null or PVPayMBMStatus = ''P'') '
            WHEN 6 THEN N' and (PVPayMBMStatus in (''F'',''P'') and PVNo=''VOID'' and PVChequeNo=''VOID'') ' ELSE N'' END
        + CASE WHEN @payment <> 6 THEN N' and (Payment2MBM.PVNo<>''VOID'' or Payment2MBM.PVNo is null) ' ELSE N'' END
        + CASE WHEN @pvFrom IS NOT NULL THEN N' and PVNo >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' and PVNo <= @pvTo' ELSE N'' END
        + CASE WHEN @blankBatch = 1 THEN N'AND (OPBatchNo Is null OR OPBatchNo = '''') ' ELSE N'' END
        + CASE WHEN @opFrom IS NOT NULL THEN N' AND OPBatchNo >= @opFrom' ELSE N'' END
        + CASE WHEN @opTo IS NOT NULL THEN N' AND OPBatchNo <= @opTo' ELSE N'' END
        + CASE WHEN @blankPv = 1 THEN N'AND (PVNo Is null OR PVNo = '''') ' ELSE N'' END
        + CASE WHEN @pvDateFrom IS NOT NULL THEN N' And PVDate >= @pvDateFrom' ELSE N'' END
        + CASE WHEN @pvDateTo IS NOT NULL THEN N' And PVDate <= @pvDateTo' ELSE N'' END
        + CASE WHEN @blankPvDate = 1 THEN N'AND (PVDate Is null OR PVDate = '''') ' ELSE N'' END
        + CASE WHEN @chqFrom IS NOT NULL THEN N' and PVChequeNo >= @chqFrom' ELSE N'' END
        + CASE WHEN @chqTo IS NOT NULL THEN N' and PVChequeNo <= @chqTo' ELSE N'' END
        + CASE WHEN @blankChq = 1 THEN N'AND (PVChequeNo Is null OR PVChequeNo = '''') ' ELSE N'' END
        + CASE WHEN @crFrom IS NOT NULL THEN N' and DebitCreditRefNo >= @crFrom' ELSE N'' END
        + CASE WHEN @crTo IS NOT NULL THEN N' and DebitCreditRefNo <= @crTo' ELSE N'' END
        + CASE WHEN @blankCr = 1 THEN N'AND (DebitCreditRefNo Is null OR DebitCreditRefNo = '''') ' ELSE N'' END
        + CASE WHEN @mthFrom IS NOT NULL THEN N' and PAYAccPeriodMth >= @mthFrom' ELSE N'' END
        + CASE WHEN @mthTo IS NOT NULL THEN N' and PAYAccPeriodMth <= @mthTo' ELSE N'' END
        + CASE WHEN @blankMth = 1 THEN N' AND PAYAccPeriodMth Is null ' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL AND @yrFrom IS NOT NULL THEN N' and PAYAccPeriodYr >= @yrFrom' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL THEN N' and PAYAccPeriodYr <= @yrTo' ELSE N'' END
        + CASE WHEN @blankYr = 1 THEN N' AND PAYAccPeriodYr Is null ' ELSE N'' END
        + CASE @bordx WHEN 0 THEN N' and ClaimBordxStatus = ''0''' WHEN 1 THEN N' and ClaimBordxStatus = 1' ELSE N'' END
        + CASE WHEN @bordxFrom IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.BordxRefNo >= @bordxFrom' ELSE N'' END
        + CASE WHEN @bordxTo IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.BordxRefNo <= @bordxTo' ELSE N'' END
        + CASE @btbg WHEN 0 THEN N' and BTBGBordxStatus = ''0''' WHEN 1 THEN N' and BTBGBordxStatus = 1' ELSE N'' END
        + CASE WHEN @btbgFrom IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.BTBGBordxRefNo >= @btbgFrom' ELSE N'' END
        + CASE WHEN @btbgTo IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.BTBGBordxRefNo <= @btbgTo' ELSE N'' END
        END;
    IF @caseNo IS NOT NULL SET @caseFrom = @caseNo;
			declare @sql nvarchar(max)
			declare @orderBy nvarchar(max)
			set @sql = 'SELECT  CISCas.MBMnumber,CISCas.CoverID,CISCas.DateVisit,CISCas.Remarks,CISCas.Pending,CISCas.Claimability, 
							CISCas.CASStatus,Payment2MBM.PVNo,Payment2MBM.PVMBMIndex,Payment2MBM.PVDate,Payment2MBM.PVChequeNo,Payment2MBM.PVChequeDate, 
							Payment2MBM.PAYAccPeriodMth,Payment2MBM.MBMPaidAmt,Payment2MBM.PAYAccPeriodYr,Payment2MBM.PVPayMBMStatus, 
							Payment2MBM.Pay2MBMLastUpdateDate,Payment2MBM.Pay2MBMLastUpdateUser,CLN.CLNName,CISCas.ClinicCode,Claim.DebitCreditRefNo,
							ClaimInvoice.INVPayTo,ClaimInvoice.INVEnteredBy,ClaimInvoice.BillNo,Claim.CASNumber,Claim.BordxRefNo, 
							Claim.CLMPayableByMedix2M,Claim.ClaimCloseDate,Payment2MBM.PVPayto,Bordx.GroupCompany,Payment2MBM.PVPayeeBank,Payment2MBM.PVPayeeACNo, 
							Payment2MBM.Remark,Payment2MBM.PVCLNChequeAmt,CISCas.OPBatchNo,Claim.CLMPostageFee,Claim.CLMNonGPFee,mbm.MBMName, 
							Claim.ClaimBordxStatus,Payment2MBM.PVWordingOne,Payment2MBM.PVWordingTwo,Claim.CLMFinalApprove,MBMOthersTwo.MBMBankACNo, 
							MBMOthersTwo.MBMBankName,Claim.CLMTotalActual,MBMOthers.MBMGroupCompany,CISCas.InvoiceNo,MBMOthers.MBMGroupCategory, 
							MBMOthers.MBMCostCenter,MBMOthers.MBMSubDivision,MBMOthers.MBMDivision,MBMOthers.Subsidiary, 
							MBMOthers.MBMEmployeeNo,GRPCompany.BusinessChanel,Tbl_BusinessChaneel.Description AS Descr,MBMOthers.MBMEmployeeNo , 
							cisdb.dbo.fnGetTreatmentDesc(CASTreatReason1) as CASTreatReason1,cisdb.dbo.fnGetTreatmentDesc(CASTreatReason2) as CASTreatReason2,
							Claim.BTBGBordxStatus,Claim.BTBGBordxRefNo,Claim.CLMPayableByMedix2M,Claim.CLMNonGPFee FROM  CISDB..CISCas 
							LEFT JOIN HISDB..MBMOthers ON MBMOthers.MBMNumber=CISCas.MBMnumber
							INNER JOIN HISDB..mbm ON MBMOthers.MBMNumber = mbm.MBMNumber 				
							LEFT JOIN HISDB..MBMOthersTwo ON mbm.MBMNumber = MBMOthersTwo.MBMNumber
							LEFT JOIN CISWorksheet..Claim ON CISCas.CASNumber = Claim.CASNumber 
							LEFT JOIN CISWorksheet..Bordx ON Claim.BordxRefNo = Bordx.BordxRefNo				
							LEFT JOIN CISWorksheet..ClaimInvoice ON Claim.CASNumber = ClaimInvoice.CASNumber	
							LEFT JOIN CISPayment..Payment2MBM ON Claim.CASNumber = Payment2MBM.CASNumber 
							LEFT JOIN HISDB..CLN ON CISCas.ClinicCode = CLN.CLNCode 
							LEFT JOIN CISDB..GRPCompany 
							on GRPCompany.GRPCoCode=claim.clmgrpcomcode and (grpcompany.hispaycode is null or ((grpcompany.hispaycode=mbm.paycode or grpcompany.hispaycodePrevYear = mbm.paycode) or (grpcompany.GRPCoPaycode=''DC'' and mbm.MBMPayorEffDate is null)))
							LEFT JOIN HISMaintenance..Tbl_BusinessChaneel ON GRPCompany.BusinessChanel = Tbl_BusinessChaneel.PreFix
							where				  
							(coalesce(Claim.CLMPayableByMedix2M,0) + coalesce(Claim.CLMNonGPFee,0) <> 0) AND CISCas.CASStatus = ''O'' and  CISCas.Claimability = ''A''  '

			set @orderBy=' order by Claim.CASNumber asc, CISCas.ClinicCode '
			set @sql = @sql + @filter + @orderBy
			exec CISPayment.sys.sp_executesql @sql, N'@caseFrom nvarchar(30), @caseTo nvarchar(30), @mbmNo nvarchar(30), @mbmName nvarchar(200), @company nvarchar(200), @pvFrom nvarchar(30), @pvTo nvarchar(30), @opFrom nvarchar(30), @opTo nvarchar(30), @pvDateFrom datetime, @pvDateTo datetime, @chqFrom nvarchar(50), @chqTo nvarchar(50), @crFrom nvarchar(30), @crTo nvarchar(30), @mthFrom nvarchar(2), @mthTo nvarchar(2), @yrFrom nvarchar(4), @yrTo nvarchar(4), @bordxFrom nvarchar(20), @bordxTo nvarchar(20), @btbgFrom nvarchar(20), @btbgTo nvarchar(20), @pvNo nvarchar(30)', @caseFrom, @caseTo, @mbmNo, @mbmName, @company, @pvFrom, @pvTo, @opFrom, @opTo, @pvDateFrom, @pvDateTo, @chqFrom, @chqTo, @crFrom, @crTo, @mthFrom, @mthTo, @yrFrom, @yrTo, @bordxFrom, @bordxTo, @btbgFrom, @btbgTo, @pvNo;
END
GO

-- Payment2MBMListing's patient name (cover 00: MBM.MBMName, else MBMCoveredPersons.MBMCName) and the partial-paid total per case.
CREATE OR ALTER PROCEDURE cis.PayMbm_Extra @cases nvarchar(max) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @x xml = CAST('<c>' + REPLACE(REPLACE(REPLACE(REPLACE(@cases, '&', ''), '<', ''), ' ', ''), ',', '</c><c>') + '</c>' AS xml);
    DECLARE @c TABLE (CASNumber nvarchar(30) PRIMARY KEY);
    INSERT @c SELECT DISTINCT t.c.value('.', 'nvarchar(30)') FROM @x.nodes('/c') t(c) WHERE t.c.value('.', 'nvarchar(30)') <> '';
    SELECT RTRIM(c.CASNumber) AS CASNumber,
        CASE WHEN c.CoverID = '00' THEN RTRIM(m.MBMName) ELSE RTRIM(cp.MBMCName) END AS PatientName,
        (SELECT SUM(p.MBMPaidAmt) FROM CISPayment.dbo.Payment2MBM p WHERE p.CASNumber = c.CASNumber) AS PaidTotal,
        (SELECT TOP 1 RTRIM(p.PVPayMBMStatus) FROM CISPayment.dbo.Payment2MBM p WHERE p.CASNumber = c.CASNumber) AS FirstStatus
    FROM dbo.CISCas c
    LEFT JOIN cis.Hdb_MBM m ON m.MBMNumber = c.MBMnumber
    LEFT JOIN cis.Hdb_MBMCoveredPersons cp ON cp.MBMCNumber = c.MBMnumber AND cp.MBMCCoverID = c.CoverID
    WHERE c.CASNumber IN (SELECT CASNumber FROM @c);
END
GO

-- The case's Payment2MBM rows (ProcessCheckedItem: status; Payment2MBMListing: the partial paid total).
CREATE OR ALTER PROCEDURE cis.PayMbm_CaseRows @casNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(PVMBMIndex AS nvarchar(20)) AS PVMBMIndex, RTRIM(PVNo) AS PVNo, RTRIM(PVPayMBMStatus) AS PVPayMBMStatus, MBMPaidAmt
    FROM CISPayment.dbo.Payment2MBM WHERE CASNumber = @casNumber;
END
GO

-- UpdateRecord: a partial row only gets the new status; otherwise Savecispayment2mbm's insert.
CREATE OR ALTER PROCEDURE cis.PayMbm_Save @index nvarchar(20) = NULL, @casNumber nvarchar(15), @pvNo nvarchar(15), @pvDate datetime, @chequeDate datetime,
    @chequeNo nvarchar(50), @payTo nvarchar(60), @bank nvarchar(100), @account nvarchar(30), @remark nvarchar(200), @paidAmt money, @user nvarchar(10),
    @mth nvarchar(2), @yr nvarchar(4), @status char(2) AS
BEGIN
    SET NOCOUNT ON;
    IF @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.Payment2MBM WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index)
    BEGIN
        UPDATE CISPayment.dbo.Payment2MBM SET PVPayMBMStatus = LEFT(@status, 1) WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index AND RTRIM(PVPayMBMStatus) = 'P';
        RETURN;
    END
    INSERT CISPayment.dbo.Payment2MBM (CasNumber, PVNo, PVDate, PVChequeDate, PVChequeNo, PVPayto, PVPayeeBank, PVPayeeACNo, Remark, MBMPaidAmt, Pay2MBMLastUpdateDate,
        Pay2MBMLastUpdateUser, PAYAccPeriodMth, PAYAccPeriodYr, PVPayMBMStatus)
    VALUES (@casNumber, @pvNo, @pvDate, @chequeDate, @chequeNo, @payTo, @bank, @account, COALESCE(@remark, ''), @paidAmt, CAST(GETDATE() AS date), @user, @mth, @yr, @status);
END
GO

-- CmdSave: the PV's cheque details and wording.
CREATE OR ALTER PROCEDURE cis.PayMbm_PvCheque @pvNo nvarchar(15), @chequeDate datetime, @chequeNo nvarchar(50), @chequeAmt money, @payTo nvarchar(60),
    @wordingOne nvarchar(500), @wordingTwo nvarchar(500) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISPayment.dbo.Payment2MBM SET PVChequeDate = @chequeDate, PVChequeNo = @chequeNo, PVCLNChequeAmt = @chequeAmt, PVPayto = @payTo, PVWordingOne = @wordingOne,
        PVWordingTwo = @wordingTwo
    WHERE PVNo = @pvNo;
END
GO

-- CmdUpdate: the rows of the PV (or the case's VOID rows when the PV has none).
CREATE OR ALTER PROCEDURE cis.PayMbm_PvRows @pvNo nvarchar(15), @casNumber nvarchar(15) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS (SELECT 1 FROM CISPayment.dbo.Payment2MBM WHERE PVNo <> 'VOID' AND PVNo = @pvNo)
        SELECT CAST(PVMBMIndex AS nvarchar(20)) AS PVMBMIndex, RTRIM(CASNumber) AS CASNumber, CAST(0 AS bit) AS Void
        FROM CISPayment.dbo.Payment2MBM WHERE PVNo <> 'VOID' AND PVNo = @pvNo;
    ELSE
        SELECT CAST(PVMBMIndex AS nvarchar(20)) AS PVMBMIndex, RTRIM(CASNumber) AS CASNumber, CAST(1 AS bit) AS Void
        FROM CISPayment.dbo.Payment2MBM WHERE PVNo = 'VOID' AND CASNumber = @casNumber;
END
GO

-- UpdatePVPaymentMBM.
CREATE OR ALTER PROCEDURE cis.PayMbm_Edit @index nvarchar(20), @casNumber nvarchar(15), @pvDate datetime, @chequeDate datetime, @chequeNo nvarchar(50),
    @chequeAmt money, @payTo nvarchar(60), @remark nvarchar(200), @wordingOne nvarchar(500), @wordingTwo nvarchar(500) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISPayment.dbo.Payment2MBM SET CASNumber = @casNumber, PVDate = @pvDate, PVChequeDate = @chequeDate, PVChequeNo = @chequeNo, PVCLNChequeAmt = @chequeAmt,
        PVPayto = @payTo, Remark = COALESCE(@remark, ''), PVWordingOne = @wordingOne, PVWordingTwo = @wordingTwo
    WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index;
END
GO

-- PreviewPV (@cases) / the member PV range (@pvFrom..@pvTo): View_Payment2MBM rows with the member's address.
CREATE OR ALTER PROCEDURE cis.PayMbm_PvPrint @cases nvarchar(max) = NULL, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @c TABLE (CASNumber nvarchar(30) PRIMARY KEY);
    IF @cases IS NOT NULL
    BEGIN
        DECLARE @x xml = CAST('<c>' + REPLACE(REPLACE(REPLACE(REPLACE(@cases, '&', ''), '<', ''), ' ', ''), ',', '</c><c>') + '</c>' AS xml);
        INSERT @c SELECT DISTINCT t.c.value('.', 'nvarchar(30)') FROM @x.nodes('/c') t(c) WHERE t.c.value('.', 'nvarchar(30)') <> '';
    END
    ELSE
        INSERT @c SELECT DISTINCT RTRIM(CASNumber) FROM CISPayment.dbo.Payment2MBM WHERE PVNo >= @pvFrom AND PVNo <= @pvTo;
    SELECT v.CASNumber, v.BordxRefNo, v.MBMPaidAmt, v.MBMnumber, v.CLMPayableByMedix2M, v.ClaimCloseDate, v.PVNo, v.PVChequeNo, v.PVPayto, v.PVDate, v.CLMPostageFee,
        v.PVCLNChequeAmt, v.CLMFinalApprove, v.MBMName, m.MBMAddress1, m.MBMAddress2, m.MBMAddress3, m.MBMPostCode, m.CTYCode, m.STACode
    FROM CISPayment.dbo.View_Payment2MBM v
    LEFT JOIN cis.Hdb_MBM m ON m.MBMNumber = v.MBMnumber
    WHERE v.CASNumber IN (SELECT CASNumber FROM @c) AND (@pvFrom IS NULL OR (v.PVNo >= @pvFrom AND v.PVNo <= @pvTo))
    ORDER BY v.PVNo, v.CASNumber;
END
GO

-- frmPrintChequeRange (member): the cheques (VIEW_Payment2MBM_Chq, or VIEW_Payment2MBM_Acc for Print Acc No) and their print status.
CREATE OR ALTER PROCEDURE cis.PayMbm_Cheques @chequeFrom nvarchar(50), @chequeTo nvarchar(50), @account bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    IF @account = 1
        SELECT PVChequeNo, PVChequeDate, PVCLNChequeAmt, PVWordingOne, PVWordingTwo, PVPayto, PVPayeeACNo, MBMGroupCompany
        FROM CISPayment.dbo.VIEW_Payment2MBM_Acc WHERE PVChequeNo >= @chequeFrom AND PVChequeNo <= @chequeTo ORDER BY PVChequeNo;
    ELSE
        SELECT PVChequeNo, PVChequeDate, PVCLNChequeAmt, PVWordingOne, PVWordingTwo, PVPayto, PVPayeeACNo, MBMGroupCompany
        FROM CISPayment.dbo.VIEW_Payment2MBM_Chq WHERE PVChequeNo >= @chequeFrom AND PVChequeNo <= @chequeTo ORDER BY PVChequeNo;
END
GO

CREATE OR ALTER PROCEDURE cis.PayMbm_ChequesPrinted @chequeFrom nvarchar(50), @chequeTo nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISPayment.dbo.Payment2MBM SET PVChequeStatus = 'Y' WHERE PVChequeNo >= @chequeFrom AND PVChequeNo <= @chequeTo;
END
GO

-- ===== main/925_payment_member_old.sql =====
-- Payment > To Member (old) (frmPaymentMember): View_Payment2MBM with the pasted criteria replaced by parameters, and UpdateRecord.
SET QUOTED_IDENTIFIER ON;
GO

-- @payment 0-5 as CboPayment; @bordx 0 / 1 (NULL: not chosen); @blank*: the "empty" boxes; @pvNo: the PV just saved. The PV no. boxes are each an "equals"
-- filter, as the desktop.
CREATE OR ALTER PROCEDURE cis.PayMbmOld_Search @pvNo nvarchar(30) = NULL, @scan bit = 0, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL,
    @mbmNo nvarchar(30) = NULL, @mbmName nvarchar(200) = NULL, @company nvarchar(200) = NULL, @payment int = 0, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL,
    @blankBatch bit = 0, @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @blankPv bit = 0, @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL,
    @blankPvDate bit = 0, @chqFrom nvarchar(50) = NULL, @chqTo nvarchar(50) = NULL, @blankChq bit = 0, @crFrom nvarchar(30) = NULL, @crTo nvarchar(30) = NULL,
    @blankCr bit = 0, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @blankMth bit = 0, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL,
    @blankYr bit = 0, @bordx int = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM CISPayment.dbo.View_Payment2MBM
    WHERE (@pvNo IS NULL OR PVNo = @pvNo)
      AND (@pvNo IS NOT NULL OR (
            (@scan = 0 OR CASNumber = @caseFrom)
        AND (@scan = 1 OR @caseFrom IS NULL OR CASNumber >= @caseFrom) AND (@scan = 1 OR @caseTo IS NULL OR CASNumber <= @caseTo)
        AND (@mbmNo IS NULL OR MBMnumber = @mbmNo) AND (@mbmName IS NULL OR MBMName LIKE '%' + @mbmName + '%') AND (@company IS NULL OR GroupCompany = @company)
        AND (@payment <> 1 OR PVPayMBMStatus = 'F') AND (@payment <> 2 OR PVPayMBMStatus = '' OR PVPayMBMStatus IS NULL)
        AND (@payment <> 3 OR PVPayMBMStatus = 'P') AND (@payment <> 4 OR PVPayMBMStatus IN ('P', 'F'))
        AND (@payment <> 5 OR PVPayMBMStatus = '' OR PVPayMBMStatus IS NULL OR PVPayMBMStatus = 'P')
        AND (@pvFrom IS NULL OR PVNo = @pvFrom) AND (@pvTo IS NULL OR PVNo = @pvTo)
        AND (@blankBatch = 0 OR OPBatchNo IS NULL OR OPBatchNo = '') AND (@opFrom IS NULL OR OPBatchNo >= @opFrom) AND (@opTo IS NULL OR OPBatchNo <= @opTo)
        AND (@blankPv = 0 OR PVNo IS NULL OR PVNo = '')
        AND (@pvDateFrom IS NULL OR PVDate >= @pvDateFrom) AND (@pvDateTo IS NULL OR PVDate <= @pvDateTo) AND (@blankPvDate = 0 OR PVDate IS NULL OR PVDate = '')
        AND (@chqFrom IS NULL OR PVChequeNo >= @chqFrom) AND (@chqTo IS NULL OR PVChequeNo <= @chqTo) AND (@blankChq = 0 OR PVChequeNo IS NULL OR PVChequeNo = '')
        AND (@crFrom IS NULL OR DebitCreditRefNo >= @crFrom) AND (@crTo IS NULL OR DebitCreditRefNo <= @crTo)
        AND (@blankCr = 0 OR DebitCreditRefNo IS NULL OR DebitCreditRefNo = '')
        AND (@mthFrom IS NULL OR PAYAccPeriodMth >= @mthFrom) AND (@mthTo IS NULL OR PAYAccPeriodMth <= @mthTo) AND (@blankMth = 0 OR PAYAccPeriodMth IS NULL)
        AND (@yrTo IS NULL OR @yrFrom IS NULL OR PAYAccPeriodYr >= @yrFrom) AND (@yrTo IS NULL OR PAYAccPeriodYr <= @yrTo) AND (@blankYr = 0 OR PAYAccPeriodYr IS NULL)
        AND (@bordx IS NULL OR ClaimBordxStatus = @bordx)))
    OPTION (RECOMPILE);
END
GO

-- UpdateRecord: Save (@edit 0): a partial row first gets the new status, then a new row is added (also when the case has none); Update (@edit 1): the row of
-- @index is overwritten.
CREATE OR ALTER PROCEDURE cis.PayMbmOld_Save @edit bit, @index nvarchar(20) = NULL, @casNumber nvarchar(15), @pvNo nvarchar(15), @pvDate datetime, @chequeDate datetime,
    @chequeNo nvarchar(50), @chequeAmt money, @payTo nvarchar(60), @bank nvarchar(100), @account nvarchar(30) = NULL, @remark nvarchar(200) = NULL, @paidAmt money,
    @user nvarchar(10), @mth nvarchar(2), @yr nvarchar(4), @status char(2), @wordingOne nvarchar(500), @wordingTwo nvarchar(500) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @found bit = CASE WHEN @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.Payment2MBM WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index) THEN 1 ELSE 0 END;
    IF @found = 1 AND @edit = 1
    BEGIN
        UPDATE CISPayment.dbo.Payment2MBM SET CASNumber = @casNumber, PVNo = @pvNo, PVDate = @pvDate, PVChequeDate = @chequeDate, PVChequeNo = @chequeNo,
            PVCLNChequeAmt = @chequeAmt, PVPayto = @payTo, PVPayeeBank = @bank, PVPayeeACNo = COALESCE(@account, ''), Remark = COALESCE(@remark, ''), MBMPaidAmt = @paidAmt,
            Pay2MBMLastUpdateDate = CAST(GETDATE() AS date), Pay2MBMLastUpdateUser = @user, PAYAccPeriodMth = @mth, PAYAccPeriodYr = @yr, PVPayMBMStatus = LEFT(@status, 1),
            PVWordingOne = @wordingOne, PVWordingTwo = @wordingTwo
        WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index;
        RETURN;
    END
    IF @found = 1
        UPDATE CISPayment.dbo.Payment2MBM SET PVPayMBMStatus = LEFT(@status, 1) WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index AND RTRIM(PVPayMBMStatus) = 'P';
    INSERT CISPayment.dbo.Payment2MBM (CASNumber, PVNo, PVDate, PVChequeDate, PVChequeNo, PVCLNChequeAmt, PVPayto, PVPayeeBank, PVPayeeACNo, Remark, MBMPaidAmt,
        Pay2MBMLastUpdateDate, Pay2MBMLastUpdateUser, PAYAccPeriodMth, PAYAccPeriodYr, PVPayMBMStatus, PVWordingOne, PVWordingTwo)
    VALUES (@casNumber, @pvNo, @pvDate, @chequeDate, @chequeNo, @chequeAmt, @payTo, @bank, COALESCE(@account, ''), COALESCE(@remark, ''), @paidAmt,
        CAST(GETDATE() AS date), @user, @mth, @yr, LEFT(@status, 1), @wordingOne, @wordingTwo);
END
GO

-- ===== main/930_payment_tracking.sql =====
-- Payment > Tracking / Tracking (Jupiter) (FrmPayment / frmPayment2): CISPayment PaymentReceiptModule cases 4-9 (Pay to Clinic, VIEW_PaymentTracking) and
-- 10-15 (Pay to Member, VIEW_PaymentMBMTracking), with the pasted criteria replaced by parameters.
SET QUOTED_IDENTIFIER ON;
GO

-- @member 0 clinic / 1 member; @by: general, bordx, claim, clinic, group, pv; @payment 1 paid (F), 2 partial (P).
CREATE OR ALTER PROCEDURE cis.PayTracking_Search @member bit, @by varchar(10), @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL,
    @company nvarchar(200) = NULL, @mbmNo nvarchar(30) = NULL, @invFrom nvarchar(50) = NULL, @invTo nvarchar(50) = NULL, @bordxFrom nvarchar(20) = NULL,
    @bordxTo nvarchar(20) = NULL, @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @clinic nvarchar(4) = NULL, @invDateFrom datetime = NULL,
    @invDateTo datetime = NULL, @payment int = 0, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL, @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL,
    @amtFrom money = NULL, @amtTo money = NULL, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @view nvarchar(50) = CASE WHEN @member = 1 THEN N'VIEW_PaymentMBMTracking' ELSE N'VIEW_PaymentTracking' END,
        @status nvarchar(30) = CASE WHEN @member = 1 THEN N'PVPayMBMStatus' ELSE N'PVPayCLNStatus' END,
        @pv nvarchar(30) = CASE WHEN @member = 1 THEN N'PVNo' ELSE N'PVCLNNo' END,
        @pvDate nvarchar(30) = CASE WHEN @member = 1 THEN N'PVDate' ELSE N'PVCLNDate' END,
        @payable nvarchar(30) = CASE WHEN @member = 1 THEN N'CLMPayableByMedix2M' ELSE N'CLMPayableByMedix2C' END,
        @mth nvarchar(30) = CASE WHEN @member = 1 THEN N'PAYAccPeriodMth' ELSE N'CLNAccPeriodMth' END,
        @yr nvarchar(30) = CASE WHEN @member = 1 THEN N'PAYAccPeriodYR' ELSE N'CLNAccPeriodYR' END,
        @paid nvarchar(30) = CASE WHEN @member = 1 THEN N'MBMPaidAmt' ELSE N'CLNPaidAmt' END;
    DECLARE @filter nvarchar(max) = N' WHERE 0 = 0'
        + CASE WHEN @caseFrom IS NOT NULL THEN N' And CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN @caseTo IS NOT NULL THEN N' And CASNumber <= @caseTo' ELSE N'' END
        + CASE WHEN @company IS NOT NULL THEN N' And GroupCompany = @company' ELSE N'' END
        + CASE WHEN @mbmNo IS NOT NULL THEN N' And MBMNumber = @mbmNo' ELSE N'' END
        + CASE WHEN @invFrom IS NOT NULL THEN N' AND InvoiceNo >= @invFrom' ELSE N'' END
        + CASE WHEN @invTo IS NOT NULL THEN N' AND InvoiceNo <= @invTo' ELSE N'' END
        + CASE WHEN @bordxFrom IS NOT NULL THEN N' And BordxRefNo >= @bordxFrom' ELSE N'' END
        + CASE WHEN @bordxTo IS NOT NULL THEN N' And BordxRefNo <= @bordxTo' ELSE N'' END
        + CASE WHEN @opFrom IS NOT NULL THEN N' And OPBatchNo >= @opFrom' ELSE N'' END
        + CASE WHEN @opTo IS NOT NULL THEN N' And OPBatchNo <= @opTo' ELSE N'' END
        + CASE WHEN @clinic IS NOT NULL THEN N' AND CLinicCode = @clinic' ELSE N'' END
        + CASE WHEN @invDateFrom IS NOT NULL THEN N' And InvoiceDate >= @invDateFrom' ELSE N'' END
        + CASE WHEN @invDateTo IS NOT NULL THEN N' And InvoiceDate <= @invDateTo' ELSE N'' END
        + CASE @payment WHEN 1 THEN N' and ' + @status + N' = ''F''' WHEN 2 THEN N' and ' + @status + N' = ''P''' ELSE N'' END
        + CASE WHEN @pvFrom IS NOT NULL THEN N' AND ' + @pv + N' >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' AND ' + @pv + N' <= @pvTo' ELSE N'' END
        + CASE WHEN @pvDateFrom IS NOT NULL THEN N' And ' + @pvDate + N' >= @pvDateFrom' ELSE N'' END
        + CASE WHEN @pvDateTo IS NOT NULL THEN N' And ' + @pvDate + N' <= @pvDateTo' ELSE N'' END
        + CASE WHEN @amtFrom IS NOT NULL THEN N' AND ' + @payable + N' >= @amtFrom' ELSE N'' END
        + CASE WHEN @amtTo IS NOT NULL THEN N' AND ' + @payable + N' <= @amtTo' ELSE N'' END
        + CASE WHEN @mthFrom IS NOT NULL THEN N' And ' + @mth + N' >= @mthFrom' ELSE N'' END
        + CASE WHEN @mthTo IS NOT NULL THEN N' And ' + @mth + N' <= @mthTo' ELSE N'' END
        + CASE WHEN @yrFrom IS NOT NULL THEN N' And ' + @yr + N' >= @yrFrom' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL THEN N' And ' + @yr + N' <= @yrTo' ELSE N'' END;
    DECLARE @agg nvarchar(200) = N', COUNT(CASNumber) AS NoofCase, SUM(' + @paid + N') AS ' + @paid;
    DECLARE @sql nvarchar(max) = CASE @by
        WHEN 'bordx' THEN N'Select BordxRefNo, BordxDate' + @agg + N' From ' + @view + @filter + N' GROUP BY BordxRefNo, BordxDate'
        WHEN 'claim' THEN N'Select CASNumber, ' + @paid + N' From ' + @view + @filter
        WHEN 'clinic' THEN N'Select ClinicCode, CLNName, ' + @pvDate + @agg + N' From ' + @view + @filter + N' GROUP BY ClinicCode, CLNName, ' + @pvDate
        WHEN 'group' THEN N'Select GroupCompany, ' + @pvDate + @agg + N' From ' + @view + @filter + N' GROUP BY GroupCompany, ' + @pvDate
        WHEN 'pv' THEN N'Select ' + @pv + N', ' + @pvDate + @agg + N' From ' + @view + @filter + N' GROUP BY ' + @pv + N', ' + @pvDate
        ELSE N'Select * From ' + @view + @filter END;
    EXEC CISPayment.sys.sp_executesql @sql,
        N'@caseFrom nvarchar(30), @caseTo nvarchar(30), @company nvarchar(200), @mbmNo nvarchar(30), @invFrom nvarchar(50), @invTo nvarchar(50), @bordxFrom nvarchar(20),
          @bordxTo nvarchar(20), @opFrom nvarchar(30), @opTo nvarchar(30), @clinic nvarchar(4), @invDateFrom datetime, @invDateTo datetime, @pvFrom nvarchar(30),
          @pvTo nvarchar(30), @pvDateFrom datetime, @pvDateTo datetime, @amtFrom money, @amtTo money, @mthFrom nvarchar(2), @mthTo nvarchar(2), @yrFrom nvarchar(4),
          @yrTo nvarchar(4)',
        @caseFrom, @caseTo, @company, @mbmNo, @invFrom, @invTo, @bordxFrom, @bordxTo, @opFrom, @opTo, @clinic, @invDateFrom, @invDateTo, @pvFrom, @pvTo, @pvDateFrom,
        @pvDateTo, @amtFrom, @amtTo, @mthFrom, @mthTo, @yrFrom, @yrTo;
END
GO

-- ===== main/935_payment_clinic_post.sql =====
-- Payment > Pay Clinic Post (frmPayClinicPost): CISPayment VIEW_PayClinic_Posting with the pasted criteria replaced by parameters, and the posting flag.
SET QUOTED_IDENTIFIER ON;
GO

-- @posted 0 / 1 as cboPostStatus (NULL: not chosen).
CREATE OR ALTER PROCEDURE cis.PayClnPost_Search @payCode nvarchar(10) = NULL, @invFrom nvarchar(50) = NULL, @invTo nvarchar(50) = NULL, @regFrom datetime = NULL,
    @regTo datetime = NULL, @posted char(1) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT PAYCode, PVCLNNo, CLNPaidAmt, InvoiceNo, MBMGroupCompany, CASPayCLNPostUser, CASPayCLNPostDate, InvoiceDate, BordxRefNo, ClinicCode, CLNName, RegisterDate,
        GRPCoCode, CASPayCLNPost, PVCLNDate, PVCLNChequeNo
    FROM CISPayment.dbo.VIEW_PayClinic_Posting
    WHERE (@payCode IS NULL OR PAYCode = @payCode) AND (@invFrom IS NULL OR InvoiceNo >= @invFrom) AND (@invTo IS NULL OR InvoiceNo <= @invTo)
      AND (@regFrom IS NULL OR RegisterDate >= @regFrom) AND (@regTo IS NULL OR RegisterDate <= @regTo) AND (@posted IS NULL OR CASPayCLNPost = @posted)
    OPTION (RECOMPILE);
END
GO

-- Post: the cases of an invoice and clinic not yet posted are marked as posted.
CREATE OR ALTER PROCEDURE cis.PayClnPost_Mark @invoiceNo nvarchar(50), @clinicCode nvarchar(10), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCas SET CASPayCLNPost = '1', CASPayCLNPostUser = @user, CASPayCLNPostDate = CAST(GETDATE() AS date)
    WHERE InvoiceNo = @invoiceNo AND ClinicCode = @clinicCode AND COALESCE(CASPayCLNPost, '0') = '0';
END
GO

-- ===== main/940_payment_ffs.sql =====
-- Payment > FFS Payment (FrmPaymentFFS): CISPayment SP_CISPayment 'PayFFS' (View_PaymentFFS) with the pasted criteria replaced by parameters, and UpdateRecord.
SET QUOTED_IDENTIFIER ON;
GO

-- @pvNo: the PV just saved / updated; @scan: one bordx no.; @payment 1 paid, 2 unpaid; @blank*: the "empty" boxes; @bordxDate: the preview row's bordx date.
CREATE OR ALTER PROCEDURE cis.PayFfs_Search @pvNo nvarchar(30) = NULL, @company nvarchar(200) = NULL, @payment int = 0, @scan bit = 0, @bordxFrom nvarchar(20) = NULL,
    @bordxTo nvarchar(20) = NULL, @blankBordxDate bit = 0, @bordxDateFrom datetime = NULL, @bordxDateTo datetime = NULL, @blankPv bit = 0, @pvFrom nvarchar(30) = NULL,
    @pvTo nvarchar(30) = NULL, @blankPvDate bit = 0, @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL, @blankMth bit = 0, @mthFrom nvarchar(2) = NULL,
    @mthTo nvarchar(2) = NULL, @blankYr bit = 0, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 100000 GroupCompany, BordxRefNo, CLMMCOProcFee, PaymentIndex, PVNo, PVDate, PVChqNo, PVChqDate, PVChqAmt, PVBank, PVAcc, Remark, FFSPaidAmt, FFSAccMth,
        FFSAccYr, PVPayStatus, GRPCoCode, BordxDate, GRPCoName, GRPCoAdd1, GRPCoAdd2, GRPCoAdd3, PVRecNo, PVRecDate, CLMBordxAmt
    FROM CISPayment.dbo.View_PaymentFFS
    WHERE (@pvNo IS NULL OR PVNo = @pvNo)
      AND (@pvNo IS NOT NULL OR (
            (@company IS NULL OR GroupCompany = @company)
        AND (@payment <> 1 OR PVPayStatus = 'F') AND (@payment <> 2 OR PVPayStatus = '' OR PVPayStatus IS NULL)
        AND (@scan = 0 OR BordxRefNo = @bordxFrom)
        AND (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
        AND (@blankBordxDate = 0 OR BordxDate IS NULL OR BordxDate = '')
        AND (@bordxDateFrom IS NULL OR BordxDate >= @bordxDateFrom) AND (@bordxDateTo IS NULL OR BordxDate <= @bordxDateTo)
        AND (@blankPv = 0 OR PVNo IS NULL OR PVNo = '')
        AND (@pvFrom IS NULL OR PVNo >= @pvFrom) AND (@pvTo IS NULL OR PVNo <= @pvTo)
        AND (@blankPvDate = 0 OR PVDate IS NULL OR PVDate = '')
        AND (@pvDateFrom IS NULL OR PVDate >= @pvDateFrom) AND (@pvDateTo IS NULL OR PVDate <= @pvDateTo)
        AND (@blankMth = 0 OR FFSAccMth IS NULL) AND (@mthFrom IS NULL OR FFSAccMth >= @mthFrom) AND (@mthTo IS NULL OR FFSAccMth <= @mthTo)
        AND (@blankYr = 0 OR FFSAccYr IS NULL) AND (@yrFrom IS NULL OR FFSAccYr >= @yrFrom) AND (@yrTo IS NULL OR FFSAccYr <= @yrTo)))
    OPTION (RECOMPILE);
END
GO

-- UpdateRecord: the PaymentFFS row of @index is overwritten, or a new row is added; GRPCoCode is the payee's code.
CREATE OR ALTER PROCEDURE cis.PayFfs_Save @index nvarchar(20) = NULL, @bordxNo nvarchar(20), @pvNo nvarchar(30), @pvDate datetime, @chequeDate datetime = NULL,
    @chequeNo nvarchar(30) = NULL, @chequeAmt money, @bank nvarchar(10), @account nvarchar(50) = NULL, @remark nvarchar(500) = NULL, @paidAmt money, @mth nvarchar(2),
    @yr nvarchar(4), @payee nvarchar(200), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @grp nvarchar(20) = (SELECT TOP 1 RTRIM(GRPCoCode) FROM dbo.GRPCompany WHERE GRPCoName = @payee);
    IF @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.PaymentFFS WHERE CAST(PaymentIndex AS nvarchar(20)) = @index)
        UPDATE CISPayment.dbo.PaymentFFS SET BordxRefNo = @bordxNo, PVNo = @pvNo, PVDate = @pvDate, PVChqDate = @chequeDate, PVChqNo = COALESCE(@chequeNo, ''),
            PVChqAmt = @chequeAmt, PVBank = @bank, PVAcc = COALESCE(@account, ''), Remark = COALESCE(@remark, ''), FFSPaidAmt = @paidAmt, FFSAccMth = @mth, FFSAccYr = @yr,
            PVLastUpdateDate = CAST(GETDATE() AS date), PVLastUpdateUser = @user, PVPayStatus = 'F', GRPCoCode = COALESCE(@grp, '')
        WHERE CAST(PaymentIndex AS nvarchar(20)) = @index;
    ELSE
        INSERT CISPayment.dbo.PaymentFFS (BordxRefNo, PVNo, PVDate, PVChqDate, PVChqNo, PVChqAmt, PVBank, PVAcc, Remark, FFSPaidAmt, FFSAccMth, FFSAccYr, PVLastUpdateDate,
            PVLastUpdateUser, PVPayStatus, GRPCoCode)
        VALUES (@bordxNo, @pvNo, @pvDate, @chequeDate, COALESCE(@chequeNo, ''), @chequeAmt, @bank, COALESCE(@account, ''), COALESCE(@remark, ''), @paidAmt, @mth, @yr,
            CAST(GETDATE() AS date), @user, 'F', COALESCE(@grp, ''));
END
GO

-- ===== main/950_receipt_member.sql =====
-- Receipt > Fr Member (FrmReceiptMember): CISPayment PaymentReceiptModule case 1 (VIEW_CISReceiptFrMBM, member debit notes) with the pasted criteria
-- replaced by parameters, UpdateRecord (the receipt of the ticked cases) and cmdUpdate (the edited receipt).
SET QUOTED_IDENTIFIER ON;
GO

-- @payment 0-5 as CboPayment; @blank*: the "empty" boxes. The receipt no. range uses the "from" box for both ends, as the desktop.
CREATE OR ALTER PROCEDURE cis.RcpMbm_Search @scan bit = 0, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL, @mbmNo nvarchar(30) = NULL,
    @mbmName nvarchar(200) = NULL, @payment int = 0, @receiptFrom nvarchar(30) = NULL, @blankReceipt bit = 0, @recDateFrom datetime = NULL, @recDateTo datetime = NULL,
    @drFrom nvarchar(30) = NULL, @drTo nvarchar(30) = NULL, @blankDr bit = 0, @amtFrom money = NULL, @amtTo money = NULL, @blankAmt bit = 0,
    @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @blankMth bit = 0, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL, @blankYr bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT CASNumber, MBMNumber, MBMName, DebitCreditRefNo, ReceiptAmt, ReceiptNo, ReceiptDate, CLMPayableByMedix2M, PVRecMBMStatus, PAYCode, ReceiptIndex,
        RECMBMAccPeriodMth, RECMBMAccPeriodYr,
        (SELECT SUM(r.ReceiptAmt) FROM CISPayment.dbo.ReceiptfrMBM r WHERE r.CASNumber = v.CASNumber) AS ReceivedTotal
    FROM CISPayment.dbo.VIEW_CISReceiptFrMBM v
    WHERE CLMPayableByMedix2M <= 0 AND SUBSTRING(DebitCreditRefNo, 1, 1) = 'D'
      AND (@scan = 0 OR CASNumber = @caseFrom)
      AND (@scan = 1 OR @caseFrom IS NULL OR CASNumber >= @caseFrom) AND (@scan = 1 OR @caseTo IS NULL OR CASNumber <= @caseTo)
      AND (@mbmNo IS NULL OR MBMNumber LIKE '%' + @mbmNo + '%') AND (@mbmName IS NULL OR MBMName LIKE '%' + @mbmName + '%')
      AND (@payment <> 1 OR PVRecMBMStatus = 'F') AND (@payment <> 2 OR PVRecMBMStatus = '' OR PVRecMBMStatus IS NULL)
      AND (@payment <> 3 OR PVRecMBMStatus = 'P') AND (@payment <> 4 OR PVRecMBMStatus IN ('P', 'F'))
      AND (@payment <> 5 OR PVRecMBMStatus = '' OR PVRecMBMStatus IS NULL OR PVRecMBMStatus = 'P')
      AND (@receiptFrom IS NULL OR (ReceiptNo >= @receiptFrom AND ReceiptNo <= @receiptFrom))
      AND (@blankReceipt = 0 OR ReceiptNo IS NULL OR ReceiptNo = '')
      AND (@recDateFrom IS NULL OR ReceiptDate >= @recDateFrom) AND (@recDateTo IS NULL OR ReceiptDate <= @recDateTo)
      AND (@drFrom IS NULL OR DebitCreditRefNo >= @drFrom) AND (@drTo IS NULL OR DebitCreditRefNo <= @drTo)
      AND (@blankDr = 0 OR DebitCreditRefNo IS NULL OR DebitCreditRefNo = '')
      AND (@amtFrom IS NULL OR CLMPayableByMedix2M >= @amtFrom) AND (@amtTo IS NULL OR CLMPayableByMedix2M <= @amtTo)
      AND (@blankAmt = 0 OR CLMPayableByMedix2M IS NULL)
      AND (@mthFrom IS NULL OR RECMBMAccPeriodMth >= @mthFrom) AND (@mthTo IS NULL OR RECMBMAccPeriodMth <= @mthTo) AND (@blankMth = 0 OR RECMBMAccPeriodMth IS NULL)
      AND (@yrFrom IS NULL OR RECMBMAccPeriodYr >= @yrFrom) AND (@yrTo IS NULL OR RECMBMAccPeriodYr <= @yrTo) AND (@blankYr = 0 OR RECMBMAccPeriodYr IS NULL)
    OPTION (RECOMPILE);
END
GO

-- UpdateRecord: the case's first ReceiptfrMBM row (a new one when none) gets the receipt details, the status and the amount added (Full: replaced when equal).
CREATE OR ALTER PROCEDURE cis.RcpMbm_Save @casNumber nvarchar(30), @receiptNo nvarchar(30), @receiptDate datetime, @chequeNo nvarchar(30), @remarks nvarchar(500) = NULL,
    @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @status nvarchar(10), @amount money, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @index int = (SELECT TOP 1 ReceiptIndex FROM CISPayment.dbo.ReceiptfrMBM WITH (UPDLOCK) WHERE CASNumber = @casNumber);
    IF @index IS NULL
    BEGIN
        INSERT CISPayment.dbo.ReceiptfrMBM (CASNumber, ReceiptAmt) VALUES (@casNumber, NULL);
        SET @index = SCOPE_IDENTITY();
    END
    UPDATE CISPayment.dbo.ReceiptfrMBM SET ReceiptNo = @receiptNo, ReceiptDate = @receiptDate, ReceiptChequeNo = @chequeNo, ReceiptLastUpdateDate = CAST(GETDATE() AS date),
        ReceiptLastUpdateUser = @user, CASNumber = @casNumber, ReceiptRemarks = COALESCE(@remarks, ''),
        RECMBMAccPeriodMth = COALESCE(@mth, RECMBMAccPeriodMth), RECMBMAccPeriodYr = COALESCE(@yr, RECMBMAccPeriodYr), PVRecMBMStatus = LEFT(@status, 1),
        ReceiptAmt = CASE WHEN @status = 'Full' THEN CASE WHEN ReceiptAmt IS NULL OR ReceiptAmt = @amount THEN @amount ELSE @amount + ReceiptAmt END
            ELSE @amount + COALESCE(ReceiptAmt, 0) END
    WHERE ReceiptIndex = @index;
END
GO

-- cmdUpdate_Click: the selected receipt row (a new row when there is none).
CREATE OR ALTER PROCEDURE cis.RcpMbm_Edit @index nvarchar(20) = NULL, @receiptNo nvarchar(30), @receiptDate datetime, @amount money, @chequeNo nvarchar(30),
    @remarks nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptfrMBM WHERE CAST(ReceiptIndex AS nvarchar(20)) = @index)
        UPDATE CISPayment.dbo.ReceiptfrMBM SET ReceiptNo = @receiptNo, ReceiptDate = @receiptDate, ReceiptAmt = @amount, ReceiptChequeNo = @chequeNo,
            ReceiptRemarks = COALESCE(@remarks, ''), ReceiptLastUpdateDate = CAST(GETDATE() AS date), ReceiptLastUpdateUser = @user,
            RECMBMAccPeriodMth = COALESCE(@mth, RECMBMAccPeriodMth), RECMBMAccPeriodYr = COALESCE(@yr, RECMBMAccPeriodYr)
        WHERE CAST(ReceiptIndex AS nvarchar(20)) = @index;
    ELSE
        INSERT CISPayment.dbo.ReceiptfrMBM (ReceiptNo, ReceiptDate, ReceiptAmt, ReceiptChequeNo, ReceiptRemarks, ReceiptLastUpdateDate, ReceiptLastUpdateUser,
            RECMBMAccPeriodMth, RECMBMAccPeriodYr)
        VALUES (@receiptNo, @receiptDate, @amount, @chequeNo, COALESCE(@remarks, ''), CAST(GETDATE() AS date), @user, @mth, @yr);
END
GO

-- lstReceiptMBM_Click (edit mode): the receipt row.
CREATE OR ALTER PROCEDURE cis.RcpMbm_Receipt @index nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 ReceiptIndex, ReceiptNo, ReceiptDate, ReceiptAmt, ReceiptChequeNo, ReceiptRemarks, RECMBMAccPeriodMth, RECMBMAccPeriodYr
    FROM CISPayment.dbo.ReceiptFrMBM WHERE CAST(ReceiptIndex AS nvarchar(20)) = @index;
END
GO

-- ===== main/960_receipt_payor.sql =====
-- Receipt > Fr Grp Company (FrmReceiptPayor): CISPayment PaymentReceiptModule cases 18 (by bordereaux) and 19 (by worksheet) copied with the pasted
-- criteria replaced by parameters, UpdateRecord / UpdateRecord2 / UpdateBordxStatus, the OP receipt summary and the edits.
SET QUOTED_IDENTIFIER ON;
GO

-- @worksheet 0 bordereaux list (case 18) / 1 worksheet list (case 19); @scan: one bordx no. (bordereaux) or one worksheet no. (worksheet);
-- @bordxStatus 0-5, @wsStatus 0-2, @insured '' / YES / NO as the combos; @receiptNo: the receipt just saved.
CREATE OR ALTER PROCEDURE cis.RcpPay_Search @worksheet bit, @scan bit = 0, @receiptNo nvarchar(30) = NULL, @company nvarchar(200) = NULL,
    @subsidiary nvarchar(200) = NULL, @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL,
    @wsStatus int = 0, @bordxStatus int = 0, @recFrom nvarchar(30) = NULL, @recTo nvarchar(30) = NULL, @recDateFrom datetime = NULL, @recDateTo datetime = NULL,
    @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL, @payor nvarchar(10) = NULL,
    @client nvarchar(200) = NULL, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL, @mrFrom nvarchar(30) = NULL, @mrTo nvarchar(30) = NULL,
    @insured varchar(3) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @params nvarchar(max) = N'@company nvarchar(200), @subsidiary nvarchar(200), @bordxFrom nvarchar(20), @bordxTo nvarchar(20), @caseFrom nvarchar(30), @caseTo nvarchar(30), @recFrom nvarchar(30), @recTo nvarchar(30), @recDateFrom datetime, @recDateTo datetime, @mthFrom nvarchar(2), @mthTo nvarchar(2), @yrFrom nvarchar(4), @yrTo nvarchar(4), @payor nvarchar(10), @client nvarchar(200), @pvFrom nvarchar(30), @pvTo nvarchar(30), @mrFrom nvarchar(30), @mrTo nvarchar(30), @receiptNo nvarchar(30), @bordxEnd nvarchar(20)';
    DECLARE @bordxEnd nvarchar(20) = CASE WHEN @scan = 1 AND @worksheet = 0 THEN @bordxFrom ELSE @bordxTo END;
    DECLARE @filter nvarchar(max) = CASE WHEN @receiptNo IS NOT NULL THEN N' AND ReceiptfrPAY.ReceiptNo = @receiptNo' ELSE N''
        + CASE WHEN @company IS NOT NULL THEN N' And Bordx.GroupCompany = @company' ELSE N'' END
        + CASE WHEN @worksheet = 1 AND @subsidiary IS NOT NULL THEN N' AND MBMOthers.Subsidiary = @subsidiary' ELSE N'' END
        + CASE WHEN @wsStatus = 2 THEN N' and COALESCE(Claim.PVRecPayorStatus , '''')='''' ' ELSE N'' END
        + CASE WHEN @wsStatus = 1 THEN N' and Claim.PVRecPayorStatus = ''F''' ELSE N'' END
        + CASE WHEN @bordxStatus = 1 THEN N' and COALESCE(ReceiptfrPAY.ReceiptNo, '''') <> ''''' ELSE N'' END
        + CASE WHEN @bordxStatus = 2 THEN N' and COALESCE(ReceiptfrPAY.ReceiptNo, '''') = ''''' ELSE N'' END
        + CASE WHEN @bordxStatus = 3 THEN N' and Bordx.BordxPAYFrStatus = ''P''' ELSE N'' END
        + CASE WHEN @bordxStatus = 4 THEN N' and (Bordx.BordxPAYFrStatus = ''P'' or Bordx.BordxPAYFrStatus = ''F'') ' ELSE N'' END
        + CASE WHEN @bordxStatus = 5 THEN N' and (COALESCE(Bordx.BordxPAYFrStatus , '''')='''' or BordxPAYFrStatus = ''P'') ' ELSE N'' END
        + CASE WHEN @recFrom IS NOT NULL THEN N' AND ReceiptfrPAY.ReceiptNo >= @recFrom' ELSE N'' END
        + CASE WHEN @recTo IS NOT NULL THEN N' AND ReceiptfrPAY.ReceiptNo <= @recTo' ELSE N'' END
        + CASE WHEN @recDateFrom IS NOT NULL THEN N' And ReceiptfrPAY.ReceiptDate >= @recDateFrom' ELSE N'' END
        + CASE WHEN @recDateTo IS NOT NULL THEN N' And ReceiptfrPAY.ReceiptDate <= @recDateTo' ELSE N'' END
        + CASE WHEN @mthFrom IS NOT NULL THEN N' AND ReceiptfrPAY.RECAccPeriodMth >= @mthFrom' ELSE N'' END
        + CASE WHEN @mthTo IS NOT NULL THEN N' AND ReceiptfrPAY.RECAccPeriodMth <= @mthTo' ELSE N'' END
        + CASE WHEN @yrFrom IS NOT NULL THEN N' AND ReceiptfrPAY.RECAccPeriodYr >= @yrFrom' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL THEN N' AND ReceiptfrPAY.RECAccPeriodYr <= @yrTo' ELSE N'' END
        + CASE WHEN @worksheet = 1 AND @scan = 1 AND @caseFrom IS NOT NULL THEN N' AND CISCas.CASNumber = @caseFrom' ELSE N'' END
        + CASE WHEN NOT (@worksheet = 1 AND @scan = 1) AND @caseFrom IS NOT NULL THEN N' AND CISCas.CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN NOT (@worksheet = 1 AND @scan = 1) AND @caseTo IS NOT NULL THEN N' AND CISCas.CASNumber <= @caseTo' ELSE N'' END
        + CASE WHEN @worksheet = 0 AND @scan = 1 AND @bordxFrom IS NOT NULL THEN N' and Bordx.BordxRefNo = @bordxFrom' ELSE N'' END
        + CASE WHEN NOT (@worksheet = 0 AND @scan = 1) AND @bordxFrom IS NOT NULL THEN N' and Bordx.BordxRefNo >= @bordxFrom' ELSE N'' END
        + CASE WHEN NOT (@worksheet = 0 AND @scan = 1) AND @bordxTo IS NOT NULL THEN N' and Bordx.BordxRefNo <= @bordxTo' ELSE N'' END
        + CASE WHEN @bordxFrom IS NOT NULL OR @bordxTo IS NOT NULL THEN N' AND Claim.CASNumber IN (SELECT k.CASNumber FROM CISWorksheet.dbo.Claim k WHERE (k.BordxRefNo >= COALESCE(@bordxFrom, k.BordxRefNo) AND k.BordxRefNo <= COALESCE(@bordxEnd, k.BordxRefNo)) OR (k.BTBGBordxRefNo >= COALESCE(@bordxFrom, k.BTBGBordxRefNo) AND k.BTBGBordxRefNo <= COALESCE(@bordxEnd, k.BTBGBordxRefNo)))' ELSE N'' END
        + CASE WHEN @payor IS NOT NULL THEN N' And  GRPCompany.ASOINSPaycode = @payor' ELSE N'' END
        + CASE WHEN @client IS NOT NULL THEN N' And Bordx.GroupCompany = @client' ELSE N'' END
        + CASE WHEN @pvFrom IS NOT NULL THEN N' AND ReceiptfrPAY.ClientPVNo >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' AND ReceiptfrPAY.ClientPVNo <= @pvTo' ELSE N'' END
        + CASE WHEN @mrFrom IS NOT NULL THEN N' AND ReceiptfrPAY.MedixReceiptNo >= @mrFrom' ELSE N'' END
        + CASE WHEN @mrTo IS NOT NULL THEN N' AND ReceiptfrPAY.MedixReceiptNo <= @mrTo' ELSE N'' END
        + CASE WHEN @insured = 'YES' THEN N' AND substring(Bordx.BordxRefNo, 1, 3) = ''KDC'' ' ELSE N'' END
        + CASE WHEN @insured = 'NO' THEN N' AND substring(Bordx.BordxRefNo, 1, 3) <> ''KDC'' ' ELSE N'' END
        END;
    IF @worksheet = 0
    BEGIN
	              	declare @l_sqlstr18 as nvarchar(MAX)
					declare @l_sqlstrGRP18 as nvarchar(MAX)

			 set @l_sqlstr18 = N'  SELECT  
									Bordx.BordxRefNo, 
									Bordx.BordxPAYCode, 
									Bordx.BordxMNRBFRStatus, 
									Bordx.GroupCompany, 
									CASE WHEN PLN.PLNBTBG = 1 THEN SUM(COALESCE(Claim.CLMBordxAmt,0) - COALESCE (Claim.CLMBTBGAmt, 0)) ELSE SUM(Claim.CLMBordxAmt) END AS BordxClaimAmt,
									Bordx.BordxDate, 
									COALESCE (MBMOthers.Subsidiary, N'''') AS Subsidiary, 
									GRPCompany.ASOINSPaycode, 
									ReceiptfrPAY.ReceiptNo, 
									ReceiptfrPAY.ReceiptDate, 
									ReceiptfrPAY.ReceiptAmt AS RecPayAmt, 
									ReceiptfrPAY.ChequeNo, 
									ReceiptfrPAY.Remark, 
									SUM(ReceiptfrPAY.CheckAmt) AS CheckAmt, 
									ReceiptfrPAY.RECAccPeriodYR, 
									ReceiptfrPAY.RECAccPeriodMth, 
									ReceiptfrPAY.ClientPVNo, 
									CASE WHEN (substring(Bordx.BordxRefNo, 1, 3) = ''KDC'') THEN ''NO'' ELSE ''YES'' END AS isInsured, 
									CASE WHEN ISNULL(Bordx.BordxPAYFRStatus, '''') <> '''' THEN Bordx.BordxPAYFRStatus WHEN ISNULL(ReceiptfrPAY.ReceiptNo, '''') <> '''' THEN ''F'' ELSE Bordx.BordxPAYFRStatus END AS BordxPAYFRStatus, 
									CISPayment..ReceiptfrPAY.MedixReceiptNo 
									FROM    CISWorksheet..Bordx WITH (NOLOCK) 
									INNER JOIN CISWorksheet..Claim WITH (NOLOCK) ON Bordx.BordxRefNo = Claim.BordxRefNo or Bordx.BordxRefNo = Claim.BTBGBordxRefNo
									INNER JOIN CISDB..CISCas WITH (NOLOCK) ON CISCas.CASNumber = Claim.CASNumber 
									INNER JOIN HISDB..MBMOthers WITH (NOLOCK) ON MBMOthers.MBMNumber = CISCas.MBMnumber 
									INNER JOIN HISDB..MBM WITH (NOLOCK) ON MBM.MBMNumber = MBMOthers.MBMNumber 
									LEFT OUTER JOIN CISDB..GRPCompany WITH (NOLOCK) ON GRPCompany.GRPCoName = MBMOthers.MBMGroupCompany AND (GRPCompany.HISpaycode IS NULL OR GRPCompany.HISpaycode = HISDB..MBM.PAYCode) 
									INNER JOIN HISDB..MBMCLNOthers WITH (NOLOCK) ON CISCas.MBMnumber = MBMCLNOthers.MBMNumber 
									INNER JOIN CISDB..PLN WITH (NOLOCK) ON PLN.plncode = MBMCLNOthers.PLNType 
									LEFT OUTER JOIN CISPayment..ReceiptfrPAY ON Claim.CASNumber = ReceiptfrPAY.CASNumber     WHere 0=0  '
				Set @l_sqlstrGRP18 = ' GROUP BY Bordx.BordxRefNo, Bordx.BordxPAYCode, Bordx.BordxMNRBFRStatus, Bordx.GroupCompany, Bordx.BordxDate, COALESCE (MBMOthers.Subsidiary, N''''), GRPCompany.ASOINSPaycode, 
									   PLN.PLNBTBG, ReceiptfrPAY.ReceiptNo, ReceiptfrPAY.ReceiptDate, ReceiptfrPAY.ReceiptAmt, ReceiptfrPAY.ChequeNo, 
									   ReceiptfrPAY.Remark, ReceiptfrPAY.RECAccPeriodYR, ReceiptfrPAY.RECAccPeriodMth, ReceiptfrPAY.ClientPVNo, 
									   CASE WHEN ISNULL(Bordx.BordxPAYFRStatus, '''') <> '''' THEN Bordx.BordxPAYFRStatus WHEN ISNULL(ReceiptfrPAY.ReceiptNo, '''') <> '''' THEN ''F'' ELSE Bordx.BordxPAYFRStatus END, ReceiptfrPAY.MedixReceiptNo '
				 set @l_sqlstr18 = @l_sqlstr18 + @filter + @l_sqlstrGRP18 + N' OPTION (RECOMPILE)'
			 exec CISPayment.sys.sp_executesql @l_sqlstr18, @params, @company, @subsidiary, @bordxFrom, @bordxTo, @caseFrom, @caseTo, @recFrom, @recTo, @recDateFrom, @recDateTo, @mthFrom, @mthTo, @yrFrom, @yrTo, @payor, @client, @pvFrom, @pvTo, @mrFrom, @mrTo, @receiptNo, @bordxEnd
    END
    ELSE
    BEGIN
	              	declare @l_sqlstr19 as nvarchar(MAX)
					declare @l_sqlstrGRP19 as nvarchar(MAX)

			 set @l_sqlstr19 = N' SELECT Claim.CASNumber, case when (substring(ReceiptfrPAY.BordxNo,1,3)=''KDC'' and Claim.clmbtbgamt>0)  then  ReceiptfrPAY.BordxNo else Claim.BordxRefNo end as BordxRefNo,
									Claim.CLMTotalApproved, Bordx.BordxPAYCode, ReceiptfrPAY.ReceiptNo, ReceiptfrPAY.ReceiptDate, ReceiptfrPAY.ReceiptAmt, ReceiptfrPAY.RECAccPeriodMth, ReceiptfrPAY.RECAccPeriodYR, 
									ReceiptfrPAY.RecPAYAmt, ReceiptfrPAY.CheckAmt, 
									CASE WHEN ISNULL(CISWorksheet.dbo.Claim.PVRecPayorStatus, '''') <> '''' THEN CISWorksheet.dbo.Claim.PVRecPayorStatus WHEN ISNULL(ReceiptfrPAY.ReceiptNo, '''') <> '''' THEN ''F'' ELSE Claim.PVRecPayorStatus END AS PVRecPayorStatus, 
									Bordx.GroupCompany, Claim.CLMFinalApprove, Claim.CLMTotalActual, Claim.CLMBordxAmt, Bordx.BordxDate, HISDB.dbo.MBMOthers.Subsidiary, Claim.ClaimOpenDate, MBM.PAYCode, GRPCompany.ASOINSPaycode,
									ReceiptfrPAY.ClientPVNo, ReceiptfrPAY.MedixReceiptNo, case when (substring(ReceiptfrPAY.BordxNo,1,3)=''KDC'')  then  ''NO'' else ''YES'' end as isInsured FROM  CISWorksheet.dbo.Claim 
									LEFT JOIN CISPayment..ReceiptfrPAY ON Claim.CASNumber = ReceiptfrPAY.CASNumber 
									INNER JOIN CISWorksheet..Bordx ----ON case when (substring(ReceiptfrPAY.BordxNo,1,3)=''KDC'' and Claim.clmbtbgamt>0)  then  ReceiptfrPAY.BordxNo else CISWorksheet.dbo.Claim.BordxRefNo  end  =CISWorksheet.dbo.Bordx.BordxRefNo 
									  ON Bordx.BordxRefNo = Claim.BordxRefNo or Bordx.BordxRefNo = Claim.BTBGBordxRefNo
									INNER JOIN CISDB..CISCas ON CISCas.CASNumber = Claim.CASNumber 
									INNER JOIN HISDB..MBMOthers ON MBMOthers.MBMNumber = CISCas.MBMnumber 
									INNER JOIN HISDB..MBM ON CISCas.MBMnumber = MBM.MBMNumber 
									LEFT JOIN CISDB..GRPCompany ON GRPCompany.GRPCoName = MBMOthers.MBMGroupCompany AND (GRPCompany.HISpaycode IS NULL OR GRPCompany.HISpaycode = MBM.PAYCode) 
									WHere 0=0 
									'
				Set @l_sqlstrGRP19 = ' GROUP BY Claim.CASNumber, Claim.BordxRefNo, Claim.CLMTotalApproved, Bordx.BordxPAYCode, ReceiptfrPAY.ReceiptNo, ReceiptfrPAY.ReceiptDate, ReceiptfrPAY.ReceiptAmt, ReceiptfrPAY.RECAccPeriodMth, ReceiptfrPAY.RECAccPeriodYR, 
										ReceiptfrPAY.RecPAYAmt, ReceiptfrPAY.CheckAmt, Claim.PVRecPayorStatus, Bordx.GroupCompany, Claim.CLMFinalApprove, Claim.CLMTotalActual, Claim.CLMBordxAmt, Bordx.BordxDate, MBMOthers.Subsidiary, Claim.ClaimOpenDate, MBM.PAYCode, 
										GRPCompany.ASOINSPaycode,ReceiptfrPAY.ClientPVNo,ReceiptfrPAY.MedixReceiptNo,ReceiptfrPAY.MedixReceiptNo,ReceiptfrPAY.BordxNo,CISWorksheet.dbo.Claim.clmbtbgamt HAVING Claim.CASNumber IS NOT NULL '
				 set @l_sqlstr19 = @l_sqlstr19 + @filter + @l_sqlstrGRP19 + N' OPTION (RECOMPILE)'
			 exec CISPayment.sys.sp_executesql @l_sqlstr19, @params, @company, @subsidiary, @bordxFrom, @bordxTo, @caseFrom, @caseTo, @recFrom, @recTo, @recDateFrom, @recDateTo, @mthFrom, @mthTo, @yrFrom, @yrTo, @payor, @client, @pvFrom, @pvTo, @mrFrom, @mrTo, @receiptNo, @bordxEnd
    END
END
GO

-- Combos: group companies, CIS payors (CISDB..PAY), the payor's insured clients and the group company's subsidiaries.
CREATE OR ALTER PROCEDURE cis.RcpPay_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE COALESCE(GRPCoName, '') <> '' ORDER BY GRPCoName;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM dbo.PAY ORDER BY PAYCode;
END
GO

CREATE OR ALTER PROCEDURE cis.RcpPay_Clients @payor nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE ASOINSPaycode = @payor ORDER BY GRPCoName;
END
GO

-- Bordereaux tick: its status and the cases it takes (partial: the cases not yet received).
CREATE OR ALTER PROCEDURE cis.RcpPay_BordxCases @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @status nvarchar(5) = (SELECT TOP 1 RTRIM(BordxPAYFRStatus) FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @bordxNo);
    SELECT @status AS BordxPAYFRStatus;
    IF @status = 'P'
        SELECT RTRIM(CASNumber) AS CASNumber FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo AND PVRecPayorStatus IS NULL;
    ELSE
        SELECT RTRIM(CASNumber) AS CASNumber FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo OR BTBGBordxRefNo = @bordxNo;
END
GO

-- Worksheet tick: the case's receipt status.
CREATE OR ALTER PROCEDURE cis.RcpPay_CaseStatus @casNumber nvarchar(30), @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 RTRIM(PVRecPayorStatus) AS PVRecPayorStatus, CLMBordxAmt FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo;
END
GO

-- UpdateRecord (bordereaux save): the claim is received in full, a ReceiptfrPAY row is added and the bordereaux is received (F).
CREATE OR ALTER PROCEDURE cis.RcpPay_SaveBordxCase @casNumber nvarchar(30), @bordxNo nvarchar(20), @receiptNo nvarchar(30), @receiptDate datetime,
    @receiptAmt money, @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @amt money;
    SELECT TOP 1 @amt = CASE WHEN BTBGBordxRefNo IS NOT NULL THEN CLMBordxAmt - CLMBTBGAmt ELSE CLMBordxAmt END
    FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND (BordxRefNo = @bordxNo OR BTBGBordxRefNo = @bordxNo);
    UPDATE CISWorksheet.dbo.Claim SET PVRecPayorStatus = 'F' WHERE CASNumber = @casNumber AND (BordxRefNo = @bordxNo OR BTBGBordxRefNo = @bordxNo);
    SET @amt = COALESCE(@amt, 0);
    INSERT CISPayment.dbo.ReceiptfrPAY (ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, BordxNo, CASNumber, Remark, PAYLastUpdateDate, PAYLastUpdateUser, RECAccPeriodMth,
        RECAccPeriodYR, RecPayAmt, CheckAmt)
    VALUES (@receiptNo, @receiptDate, @receiptAmt, @chequeNo, @bordxNo, @casNumber, COALESCE(@remark, ''), CAST(GETDATE() AS date), @user, COALESCE(@mth, ''),
        COALESCE(@yr, ''), @amt, @amt);
    UPDATE CISWorksheet.dbo.Bordx SET BordxPAYRecDate = @receiptDate, BordxPAYFRStatus = 'F' WHERE BordxRefNo = @bordxNo;
END
GO

-- UpdateRecord2 (worksheet save): a ReceiptfrPAY row when the case has none for the bordereaux, the claim's status, then the bordereaux is F when all
-- its claims are received, else P.
CREATE OR ALTER PROCEDURE cis.RcpPay_SaveWksCase @casNumber nvarchar(30), @bordxNo nvarchar(20), @status nvarchar(2), @receiptNo nvarchar(30),
    @receiptDate datetime, @receiptAmt money, @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL,
    @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @amt money = (SELECT TOP 1 CLMBordxAmt FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo AND CASNumber = @casNumber);
    IF NOT EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptfrPAY WHERE CASNumber = @casNumber AND BordxNo = @bordxNo)
        INSERT CISPayment.dbo.ReceiptfrPAY (CASNumber, ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, BordxNo, PAYLastUpdateDate, PAYLastUpdateUser, RecPayAmt,
            CheckAmt, RECAccPeriodMth, RECAccPeriodYR)
        VALUES (@casNumber, @receiptNo, @receiptDate, @receiptAmt, @chequeNo, COALESCE(@remark, ''), @bordxNo, CAST(GETDATE() AS date), @user, COALESCE(@amt, 0),
            COALESCE(@amt, 0), @mth, @yr);
    UPDATE CISWorksheet.dbo.Claim SET PVRecPayorStatus = @status WHERE BordxRefNo = @bordxNo AND CASNumber = @casNumber;
    DECLARE @all int = (SELECT COUNT(*) FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo),
        @full int = (SELECT COUNT(*) FROM CISWorksheet.dbo.Claim WHERE PVRecPayorStatus = 'F' AND BordxRefNo = @bordxNo);
    UPDATE CISWorksheet.dbo.Bordx SET BordxPAYRecDate = @receiptDate, BordxPAYFRStatus = CASE WHEN @all = @full THEN 'F' ELSE 'P' END WHERE BordxRefNo = @bordxNo;
END
GO

-- SaveBordx: the OP claim / OP fee lines of the receipt.
CREATE OR ALTER PROCEDURE cis.RcpPay_Summary @receiptNo nvarchar(30), @receiptDate datetime, @amount money, @type nvarchar(10), @desc nvarchar(500) = NULL,
    @chequeNo nvarchar(30), @payCode nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISPayment.dbo.ReceiptSummary (ReceiptNo, ReceiptDate, ReceiptAmt, ReceiptType, ReceiptDesc, RecChqNo, RecPayCode)
    VALUES (@receiptNo, @receiptDate, @amount, @type, COALESCE(@desc, ''), @chequeNo, @payCode);
END
GO

-- lstBordx_Click / lstWorksheet_Click in edit mode: the receipt (@casNumber NULL: the bordereaux's first receipt row).
CREATE OR ALTER PROCEDURE cis.RcpPay_Receipt @bordxNo nvarchar(20), @casNumber nvarchar(30) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, RECAccPeriodMth, RECAccPeriodYR
    FROM CISPayment.dbo.ReceiptfrPAY WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber);
END
GO

-- UpdateBordx (@casNumber NULL: every receipt row of the bordereaux, with the period) / UpdateWorksheet (the case's row); a row is added when there is none.
CREATE OR ALTER PROCEDURE cis.RcpPay_Edit @bordxNo nvarchar(20), @casNumber nvarchar(30) = NULL, @receiptNo nvarchar(30), @receiptDate datetime, @receiptAmt money,
    @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptfrPAY WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber))
        INSERT CISPayment.dbo.ReceiptfrPAY (ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, PAYLastUpdateDate, PAYLastUpdateUser, RECAccPeriodMth, RECAccPeriodYR)
        VALUES (@receiptNo, @receiptDate, @receiptAmt, @chequeNo, COALESCE(@remark, ''), CAST(GETDATE() AS date), @user, CASE WHEN @casNumber IS NULL THEN @mth END,
            CASE WHEN @casNumber IS NULL THEN @yr END);
    ELSE
        UPDATE CISPayment.dbo.ReceiptfrPAY SET ReceiptNo = @receiptNo, ReceiptDate = @receiptDate, ReceiptAmt = @receiptAmt, ChequeNo = @chequeNo,
            Remark = COALESCE(@remark, ''), PAYLastUpdateDate = CAST(GETDATE() AS date), PAYLastUpdateUser = @user,
            RECAccPeriodMth = CASE WHEN @casNumber IS NULL AND @mth IS NOT NULL THEN @mth ELSE RECAccPeriodMth END,
            RECAccPeriodYR = CASE WHEN @casNumber IS NULL AND @yr IS NOT NULL THEN @yr ELSE RECAccPeriodYR END
        WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber);
END
GO

-- ===== main/970_receipt_ffs.sql =====
-- Receipt > Fr FFS (FrmReceiptFFS): CISPayment PaymentReceiptModule case 17 (VIEW_CISReceiptFFS) with the pasted criteria replaced by parameters, and UpdateRecord.
SET QUOTED_IDENTIFIER ON;
GO

-- @scan: one bordx no.; @payment 1 received, 2 not received; @receiptNo: the receipt just saved.
CREATE OR ALTER PROCEDURE cis.RcpFfs_Search @receiptNo nvarchar(30) = NULL, @scan bit = 0, @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL,
    @payment int = 0, @recFrom nvarchar(30) = NULL, @recTo nvarchar(30) = NULL, @recDateFrom datetime = NULL, @recDateTo datetime = NULL, @mthFrom nvarchar(2) = NULL,
    @mthTo nvarchar(2) = NULL, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL, @bordxDateFrom datetime = NULL, @bordxDateTo datetime = NULL,
    @company nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF @receiptNo IS NOT NULL
    BEGIN
        -- The view aggregates every claim and only a bordx no. is pushed into it, so the receipt's bordereaux are read one at a time.
        SELECT TOP 0 * INTO #r FROM CISPayment.dbo.VIEW_CISReceiptFFS;
        DECLARE @b nvarchar(20);
        DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT DISTINCT BordxRefNo FROM CISPayment.dbo.ReceiptFrFFS WHERE ReceiptNo = @receiptNo AND BordxRefNo IS NOT NULL;
        OPEN c;
        FETCH NEXT FROM c INTO @b;
        WHILE @@FETCH_STATUS = 0
        BEGIN
            INSERT #r SELECT * FROM CISPayment.dbo.VIEW_CISReceiptFFS WHERE BordxRefNo = @b AND ReceiptNo = @receiptNo OPTION (RECOMPILE);
            FETCH NEXT FROM c INTO @b;
        END
        CLOSE c;
        SELECT * FROM #r ORDER BY BordxRefNo;
        RETURN;
    END
    SELECT * FROM CISPayment.dbo.VIEW_CISReceiptFFS
    WHERE (
            (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
        AND (@scan = 0 OR @bordxFrom IS NULL OR BordxRefNo = @bordxFrom)
        AND (@payment <> 1 OR ReceiptStatus = 'F') AND (@payment <> 2 OR ReceiptStatus = '' OR ReceiptStatus IS NULL)
        AND (@recFrom IS NULL OR ReceiptNo >= @recFrom) AND (@recTo IS NULL OR ReceiptNo <= @recTo)
        AND (@recDateFrom IS NULL OR ReceiptDate >= @recDateFrom) AND (@recDateTo IS NULL OR ReceiptDate <= @recDateTo)
        AND (@mthFrom IS NULL OR ReceiptAccMth >= @mthFrom) AND (@mthTo IS NULL OR ReceiptAccMth <= @mthTo)
        AND (@yrFrom IS NULL OR ReceiptAccYr >= @yrFrom) AND (@yrTo IS NULL OR ReceiptAccYr <= @yrTo)
        AND (@bordxDateFrom IS NULL OR BordxDate >= @bordxDateFrom) AND (@bordxDateTo IS NULL OR BordxDate <= @bordxDateTo)
        AND (@company IS NULL OR GroupCompany = @company))
    OPTION (RECOMPILE);
END
GO

-- UpdateRecord: a new ReceiptFrFFS row (Save) or the row of @index overwritten (Update).
CREATE OR ALTER PROCEDURE cis.RcpFfs_Save @index nvarchar(20) = NULL, @bordxNo nvarchar(20), @receiptNo nvarchar(30), @receiptDate datetime, @chequeNo nvarchar(30) = NULL,
    @amount money, @remarks nvarchar(500) = NULL, @mth nvarchar(2), @yr nvarchar(4), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptFrFFS WHERE CAST(ReceiptIndex AS nvarchar(20)) = @index)
        UPDATE CISPayment.dbo.ReceiptFrFFS SET BordxRefNo = @bordxNo, ReceiptNo = @receiptNo, ReceiptDate = @receiptDate, ReceiptChqNo = COALESCE(@chequeNo, ''),
            ReceiptAmt = @amount, ReceiptRemarks = COALESCE(@remarks, ''), ReceiptAccMth = @mth, ReceiptAccYr = @yr, ReceiptLastUpdDate = CAST(GETDATE() AS date),
            ReceiptLastUpdUser = @user, ReceiptStatus = 'F'
        WHERE CAST(ReceiptIndex AS nvarchar(20)) = @index;
    ELSE
        INSERT CISPayment.dbo.ReceiptFrFFS (BordxRefNo, ReceiptNo, ReceiptDate, ReceiptChqNo, ReceiptAmt, ReceiptRemarks, ReceiptAccMth, ReceiptAccYr, ReceiptLastUpdDate,
            ReceiptLastUpdUser, ReceiptStatus)
        VALUES (@bordxNo, @receiptNo, @receiptDate, COALESCE(@chequeNo, ''), @amount, COALESCE(@remarks, ''), @mth, @yr, CAST(GETDATE() AS date), @user, 'F');
END
GO

-- ===== main/980_receipt_upload.sql =====
-- Receipt > Upload Receipt / Upload Receipt (Multiple Receipts) (frmReceiptUpload / frmReceiptUpMultiple): the claims of the uploaded rows and SaveBordx.
SET QUOTED_IDENTIFIER ON;
GO

-- cmpUpload_Click: the claim of each uploaded claim no. (@cases comma-separated).
CREATE OR ALTER PROCEDURE cis.RcpUp_Claims @cases nvarchar(max) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @x xml = CAST('<c>' + REPLACE(REPLACE(REPLACE(REPLACE(@cases, '&', ''), '<', ''), ' ', ''), ',', '</c><c>') + '</c>' AS xml);
    DECLARE @c TABLE (CASNumber nvarchar(30) PRIMARY KEY);
    INSERT @c SELECT DISTINCT t.c.value('.', 'nvarchar(30)') FROM @x.nodes('/c') t(c) WHERE t.c.value('.', 'nvarchar(30)') <> '';
    SELECT RTRIM(k.CASNumber) AS CASNumber, RTRIM(k.BordxRefNo) AS BordxRefNo, k.CLMBordxAmt
    FROM CISWorksheet.dbo.Claim k WHERE k.CASNumber IN (SELECT CASNumber FROM @c);
END
GO

-- SaveBordx (one matched row): a ReceiptfrPAY row and the claim received in full.
CREATE OR ALTER PROCEDURE cis.RcpUp_Save @casNumber nvarchar(30), @bordxNo nvarchar(20) = NULL, @receiptNo nvarchar(30), @receiptDate datetime = NULL,
    @receiptAmt money, @chequeNo nvarchar(30) = NULL, @excelAmt money, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @clientPvNo nvarchar(50) = NULL,
    @medixReceiptNo nvarchar(30) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF @bordxNo IS NULL SET @bordxNo = (SELECT TOP 1 RTRIM(BordxRefNo) FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber);
    INSERT CISPayment.dbo.ReceiptfrPAY (CASNumber, ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, BordxNo, PAYLastUpdateDate, PAYLastUpdateUser, RecPayAmt, CheckAmt,
        RECAccPeriodMth, RECAccPeriodYR, ClientPVNo, MedixReceiptNo)
    VALUES (@casNumber, @receiptNo, @receiptDate, @receiptAmt, COALESCE(@chequeNo, ''), COALESCE(@bordxNo, ''), CAST(GETDATE() AS date), @user, @excelAmt, @excelAmt,
        @mth, @yr, @clientPvNo, @medixReceiptNo);
    UPDATE CISWorksheet.dbo.Claim SET PVRecPayorStatus = 'F' WHERE CASNumber = @casNumber;
END
GO

-- ===== main/990_receipt_mnrb.sql =====
-- Receipt > Fr MNRB (FrmReceiptMNRB): the bordereaux (VIEW_CISBORReceiptMNRB) and worksheet (VIEW_CISWKSReceiptMNRB) lists with the pasted criteria replaced by
-- parameters, UpdateRecord / UpdateRecord2 / UpdateBordxStatus and the edits.
SET QUOTED_IDENTIFIER ON;
GO

-- PayorComboListing: HISDB PAY with HLTCode S.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_Payors AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM cis.Hdb_PAY WHERE HLTCode = 'S' ORDER BY PAYCompanyName;
END
GO

-- @worksheet 0 bordereaux / 1 worksheet; @scan: one bordx no. (bordereaux) or one worksheet no. (worksheet); @bordxStatus 0-5, @wsStatus 0-2 as the combos.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_Search @worksheet bit, @scan bit = 0, @payor nvarchar(10) = NULL, @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL,
    @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL, @wsStatus int = 0, @bordxStatus int = 0, @recFrom nvarchar(30) = NULL, @recTo nvarchar(30) = NULL,
    @recDateFrom datetime = NULL, @recDateTo datetime = NULL, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @yrFrom nvarchar(4) = NULL,
    @yrTo nvarchar(4) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF @worksheet = 0
        SELECT BordxRefNo, BordxClaimAmt, BordxPAYCode, BordxMNRBFRStatus, SUM(CheckAmt) AS CheckAmt, RECAccPeriodMth, RECAccPeriodYR
        FROM CISWorksheet.dbo.VIEW_CISBORReceiptMNRB
        WHERE (@payor IS NULL OR BordxPAYCode = @payor)
          AND (@bordxFrom IS NULL OR (@scan = 1 AND BordxRefNo = @bordxFrom) OR (@scan = 0 AND BordxRefNo >= @bordxFrom))
          AND (@bordxTo IS NULL OR @scan = 1 OR BordxRefNo <= @bordxTo)
          AND (@bordxStatus <> 1 OR BordxMNRBFRStatus = 'F') AND (@bordxStatus <> 2 OR BordxMNRBFRStatus = '' OR BordxMNRBFRStatus IS NULL)
          AND (@bordxStatus <> 3 OR BordxMNRBFRStatus = 'P') AND (@bordxStatus <> 4 OR BordxMNRBFRStatus IN ('P', 'F'))
          AND (@bordxStatus <> 5 OR BordxMNRBFRStatus = '' OR BordxMNRBFRStatus IS NULL OR BordxMNRBFRStatus = 'P')
          AND (@recFrom IS NULL OR ReceiptNo >= @recFrom) AND (@recTo IS NULL OR ReceiptNo <= @recTo)
          AND (@recDateFrom IS NULL OR ReceiptDate >= @recDateFrom) AND (@recDateTo IS NULL OR ReceiptDate <= @recDateTo)
          AND (@mthFrom IS NULL OR RECAccPeriodMth >= @mthFrom) AND (@mthTo IS NULL OR RECAccPeriodMth <= @mthTo)
          AND (@yrFrom IS NULL OR RECAccPeriodYR >= @yrFrom) AND (@yrTo IS NULL OR RECAccPeriodYR <= @yrTo)
        GROUP BY BordxRefNo, BordxClaimAmt, BordxPAYCode, BordxMNRBFRStatus, RECAccPeriodMth, RECAccPeriodYR
        OPTION (RECOMPILE);
    ELSE
        SELECT CASNumber, BordxRefNo, CLMTotalApproved, ReceiptDate, ReceiptNo, RECAccPeriodMth, RECAccPeriodYR, BordxPAYCode, PVRecMNRBStatus, CheckAmt
        FROM CISWorksheet.dbo.VIEW_CISWKSReceiptMNRB
        WHERE (@payor IS NULL OR BordxPAYCode = @payor)
          AND (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
          AND (@caseFrom IS NULL OR (@scan = 1 AND CASNumber = @caseFrom) OR (@scan = 0 AND CASNumber >= @caseFrom))
          AND (@caseTo IS NULL OR @scan = 1 OR CASNumber <= @caseTo)
          AND (@wsStatus <> 2 OR PVRecMNRBStatus = '' OR PVRecMNRBStatus IS NULL) AND (@wsStatus <> 1 OR PVRecMNRBStatus = 'F')
          AND (@recFrom IS NULL OR ReceiptNo >= @recFrom) AND (@recTo IS NULL OR ReceiptNo <= @recTo)
          AND (@recDateFrom IS NULL OR ReceiptDate >= @recDateFrom) AND (@recDateTo IS NULL OR ReceiptDate <= @recDateTo)
          AND (@mthFrom IS NULL OR RECAccPeriodMth >= @mthFrom) AND (@mthTo IS NULL OR RECAccPeriodMth <= @mthTo)
          AND (@yrFrom IS NULL OR RECAccPeriodYR >= @yrFrom) AND (@yrTo IS NULL OR RECAccPeriodYR <= @yrTo)
        OPTION (RECOMPILE);
END
GO

-- lstBordx_ItemCheck: the bordereaux's MNRB status and the claims it takes (partial: the claims not yet received).
CREATE OR ALTER PROCEDURE cis.RcpMnrb_BordxCases @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @status nvarchar(5) = (SELECT TOP 1 RTRIM(BordxMNRBFRStatus) FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @bordxNo);
    SELECT @status AS BordxMNRBFRStatus;
    SELECT RTRIM(CASNumber) AS CASNumber FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo AND (@status <> 'P' OR @status IS NULL OR PVRecMNRBStatus IS NULL);
END
GO

-- ProcessCheckedItem: the claim's MNRB status.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_CaseStatus @casNumber nvarchar(30), @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 RTRIM(PVRecMNRBStatus) AS PVRecMNRBStatus FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo;
END
GO

-- UpdateRecord (bordereaux save, one claim): a ReceiptfrMNRB row, the claim received and the bordereaux received (F).
CREATE OR ALTER PROCEDURE cis.RcpMnrb_SaveBordxCase @casNumber nvarchar(30), @bordxNo nvarchar(20), @receiptNo nvarchar(30), @receiptDate datetime, @receiptAmt money,
    @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo) RETURN;
    DECLARE @amt money = (SELECT TOP 1 CLMTotalApproved FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo);
    INSERT CISPayment.dbo.ReceiptfrMNRB (ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, BordxNo, CASNumber, Remark, MNRBLastUpdatedate, MNRBLastUpdateUser, RECAccPeriodMth,
        RECAccPeriodYR, RecMNRBAmt, CheckAmt)
    VALUES (@receiptNo, @receiptDate, @receiptAmt, @chequeNo, @bordxNo, @casNumber, COALESCE(@remark, ''), CAST(GETDATE() AS date), @user, @mth, @yr, @amt, @amt);
    UPDATE CISWorksheet.dbo.Claim SET PVRecMNRBStatus = 'F' WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo;
    UPDATE CISWorksheet.dbo.Bordx SET BordxMNRBRecDate = @receiptDate, BordxMNRBFRStatus = 'F' WHERE BordxRefNo = @bordxNo;
END
GO

-- UpdateRecord2 (worksheet save): a ReceiptfrMNRB row when the claim has none for the bordereaux, the claim's status, then the bordereaux F when all its claims are
-- received, else P.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_SaveWksCase @casNumber nvarchar(30), @bordxNo nvarchar(20), @status nvarchar(2), @receiptNo nvarchar(30), @receiptDate datetime,
    @selectedAmt money, @receiptAmt money, @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @amt money = (SELECT TOP 1 CLMTotalApproved FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo);
    IF NOT EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptfrMNRB WHERE CASNumber = @casNumber AND BordxNo = @bordxNo)
        INSERT CISPayment.dbo.ReceiptfrMNRB (CASNumber, ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, BordxNo, MNRBLastUpdatedate, MNRBLastUpdateUser, RecMNRBAmt,
            CheckAmt, RECAccPeriodMth, RECAccPeriodYR)
        VALUES (@casNumber, @receiptNo, @receiptDate, @selectedAmt, @chequeNo, COALESCE(@remark, ''), @bordxNo, CAST(GETDATE() AS date), @user, @receiptAmt,
            COALESCE(@amt, 0), @mth, @yr);
    UPDATE CISWorksheet.dbo.Claim SET PVRecMNRBStatus = @status WHERE BordxRefNo = @bordxNo AND CASNumber = @casNumber;
    DECLARE @all int = (SELECT COUNT(*) FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo),
        @full int = (SELECT COUNT(*) FROM CISWorksheet.dbo.Claim WHERE PVRecMNRBStatus = 'F' AND BordxRefNo = @bordxNo);
    UPDATE CISWorksheet.dbo.Bordx SET BordxMNRBRecDate = @receiptDate, BordxMNRBFRStatus = CASE WHEN @all = @full THEN 'F' ELSE 'P' END WHERE BordxRefNo = @bordxNo;
END
GO

-- lstBordx_Click / lstWorksheet_Click in edit mode.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_Receipt @bordxNo nvarchar(20), @casNumber nvarchar(30) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, RECAccPeriodMth, RECAccPeriodYR
    FROM CISPayment.dbo.ReceiptfrMNRB WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber);
END
GO

-- UpdateBordx (@casNumber NULL: the bordereaux's receipt rows, with the period) / UpdateWorksheet (the claim's row); a row is added when there is none.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_Edit @bordxNo nvarchar(20), @casNumber nvarchar(30) = NULL, @receiptNo nvarchar(30), @receiptDate datetime, @receiptAmt money,
    @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptfrMNRB WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber))
        INSERT CISPayment.dbo.ReceiptfrMNRB (ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, BordxNo, CASNumber, MNRBLastUpdatedate, MNRBLastUpdateUser,
            RECAccPeriodMth, RECAccPeriodYR)
        VALUES (@receiptNo, @receiptDate, @receiptAmt, @chequeNo, COALESCE(@remark, ''), @bordxNo, @casNumber, CAST(GETDATE() AS date), @user,
            CASE WHEN @casNumber IS NULL THEN @mth END, CASE WHEN @casNumber IS NULL THEN @yr END);
    ELSE
        UPDATE CISPayment.dbo.ReceiptfrMNRB SET ReceiptNo = @receiptNo, ReceiptDate = @receiptDate, ReceiptAmt = @receiptAmt, ChequeNo = @chequeNo,
            Remark = COALESCE(@remark, ''), MNRBLastUpdatedate = CAST(GETDATE() AS date), MNRBLastUpdateUser = @user,
            RECAccPeriodMth = CASE WHEN @casNumber IS NULL AND @mth IS NOT NULL THEN @mth ELSE RECAccPeriodMth END,
            RECAccPeriodYR = CASE WHEN @casNumber IS NULL AND @yr IS NOT NULL THEN @yr ELSE RECAccPeriodYR END
        WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber);
END
GO

PRINT N'Done.';
GO
SET NOEXEC OFF;
GO
