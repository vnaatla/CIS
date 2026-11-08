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
