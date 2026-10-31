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
