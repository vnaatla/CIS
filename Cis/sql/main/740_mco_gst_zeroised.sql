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
