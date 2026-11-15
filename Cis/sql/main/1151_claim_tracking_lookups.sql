-- Claim / Bordx > Claim Tracking (frmClaimRpt) combos: policy statuses, states, clinic groups, panel clinics (SearchPolStatus / SearchState / SearchCLNGrp /
-- SearchCLN), payors of a health type, group companies of a payor, departments of a group company.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.ClmTrk_Lookup @kind nvarchar(20), @key nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF @kind = 'polstatus' SELECT RTRIM(pstatuscode) + ' - ' + RTRIM(pstatusDesc) FROM HISDB.dbo.PolicyStatus;
    ELSE IF @kind = 'states' SELECT RTRIM(STADescription) FROM HISDB.dbo.STA;
    ELSE IF @kind = 'clinicgroups' SELECT RTRIM(ClinicGRPName) FROM HISDB.dbo.ClinicGRP;
    ELSE IF @kind = 'clinics' SELECT RTRIM(CLNCode) + ' - ' + RTRIM(CLNName) FROM HISDB.dbo.CLN WHERE CLNPanel = 'C' OR CLNPanel = 'Y';
    ELSE IF @kind = 'payors' SELECT RTRIM(PAYCode) + ' - ' + RTRIM(PAYCompanyName) FROM HISDB.dbo.PAY WHERE HLTCode = LEFT(@key, 1);
    ELSE IF @kind = 'companies' SELECT RTRIM(GRPCoName) FROM dbo.GRPCompany WHERE GRPCoPAYCode = LEFT(@key, 2);
    ELSE IF @kind = 'departments' SELECT DISTINCT RTRIM(MBMDepartment) FROM HISDB.dbo.MBMOthers WHERE MBMGroupCompany = @key AND MBMDepartment IS NOT NULL;
END
GO
