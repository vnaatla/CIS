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
