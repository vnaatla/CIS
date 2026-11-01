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
