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
    (N'Ws_sp_CreateSubCasNumClaim', @CisWorksheetDb, N'sp_CreateSubCasNumClaim')) v (s, d, t)
WHERE DB_ID(d) IS NOT NULL;
EXEC (@sql);
GO
