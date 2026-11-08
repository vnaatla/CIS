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
