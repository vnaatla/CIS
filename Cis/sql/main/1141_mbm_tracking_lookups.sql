-- Tracking > MBM Tracking (FrmNewMBMTracking) combos: states, insured types, payors of a health type, group companies of a payor, departments of a group company.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.MbmTrk_Lookup @kind nvarchar(20), @key nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF @kind = 'states' SELECT RTRIM(STADescription) FROM HISDB.dbo.STA;
    ELSE IF @kind = 'ins' SELECT RTRIM(INSCode) + ' - ' + RTRIM(INSDescription) FROM HISDB.dbo.INS;
    ELSE IF @kind = 'payors' SELECT RTRIM(PAYCode) + ' - ' + RTRIM(PAYCompanyName) FROM HISDB.dbo.PAY WHERE HLTCODE = CASE WHEN LEFT(@key, 1) = 'N' THEN 'N' ELSE 'S' END;
    ELSE IF @kind = 'companies' SELECT RTRIM(GRPCoName) FROM dbo.GRPCompany WHERE GRPCoPAYCode = LEFT(@key, 2);
    ELSE IF @kind = 'departments' SELECT DISTINCT RTRIM(MBMDepartment) FROM HISDB.dbo.MBMOthers WHERE MBMGroupCompany = @key AND MBMDepartment IS NOT NULL;
END
GO
