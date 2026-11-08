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
