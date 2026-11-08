-- MCO Fees > Refund Listing (frmMCORefundListing): MCO refunds of endorsed members (MCOFeeModule cases 10 / 11 / 12 with the pasted filters
-- replaced by parameters) and the credit note (CLNMCOCRNoteSeq, stamped on MBMCLNOthers / MBMCoveredPersonsCLN).
SET QUOTED_IDENTIFIER ON;
GO

-- @byBordx 0: endorsement effective date, 1: endorsement bordereaux date. @noted N: without credit note, Y: with one (principals only, as the desktop).
CREATE OR ALTER PROCEDURE cis.McoRefund_Search @payor nvarchar(10) = NULL, @company nvarchar(200) = NULL, @department nvarchar(200) = NULL, @byBordx bit = 0,
    @from datetime = NULL, @to datetime = NULL, @noted char(1) = 'N' AS
BEGIN
    SET NOCOUNT ON;
    SELECT MBMNumber, MBMPolNumber, MBMName, MBMMCOUtilised, MBMCLNStatus, MBMCOVID, MBMCLNEndtEffDate, MBMCLNEndtBordxDate, MBMPolEffDate, MBMPolExpDate,
        MBMCLNMCOFee, MBMMCORefund, MBMBordxDate, DayInsured, MBMIcBcPp, DayUtilised, INSType, MBMCLNCRNote, MBMCLNCRNoteDate, MBMCInvoiceNo, MBMCInvoiceDate, TotalCase
    FROM dbo.VIEW_CLN_MBM_Refund_Search
    WHERE (@payor IS NULL OR PAYCode = @payor) AND (@company IS NULL OR MBMGroupCompany = @company) AND (@department IS NULL OR MBMDepartment = @department)
      AND (@from IS NULL OR (@byBordx = 0 AND MBMCLNEndtEffDate >= @from) OR (@byBordx = 1 AND MBMCLNEndtBordxDate >= @from))
      AND (@to IS NULL OR (@byBordx = 0 AND MBMCLNEndtEffDate <= @to) OR (@byBordx = 1 AND MBMCLNEndtBordxDate <= @to))
      AND ((@noted = 'N' AND (MBMCLNCRNote IS NULL OR MBMCLNCRNote = '')) OR (@noted = 'Y' AND MBMCLNCRNote <> ''))
    ORDER BY MBMNumber
    OPTION (RECOMPILE);
END
GO

-- MCORefundListingSupp: the covered persons of the listed members (the desktop applies the same filters but not the credit-note status).
CREATE OR ALTER PROCEDURE cis.McoRefund_Supp @payor nvarchar(10) = NULL, @company nvarchar(200) = NULL, @department nvarchar(200) = NULL, @byBordx bit = 0,
    @from datetime = NULL, @to datetime = NULL, @members nvarchar(max) AS
BEGIN
    SET NOCOUNT ON;
    SELECT MBMNumber, MBMPolNumber, MBMCName, MBMCLNStatus, MBMCCoverID, MBMCLNEndtEffDate, MBMCLNEndtBordxDate, MBMPolEffDate, MBMPolExpDate, TotalCase,
        MBMCCLNCRNote, MBMCCLNCRNoteDate, MBMCBordxDate, DayInsured, MBMCICBCPPNo, DayUtilised, MBMCInvoiceNo, MBMCInvoiceDate, MBMMCOUtilised, MBMCCLNMCO, MBMMCORefund
    FROM dbo.VIEW_CLN_Supp_Refund_Search
    WHERE ('|' + @members + '|') LIKE '%|' + MBMNumber + '|%'
      AND (@payor IS NULL OR PAYCode = @payor) AND (@company IS NULL OR MBMGroupCompany = @company) AND (@department IS NULL OR MBMDepartment = @department)
      AND (@from IS NULL OR (@byBordx = 0 AND MBMCLNEndtEffDate >= @from) OR (@byBordx = 1 AND MBMCLNEndtBordxDate >= @from))
      AND (@to IS NULL OR (@byBordx = 0 AND MBMCLNEndtEffDate <= @to) OR (@byBordx = 1 AND MBMCLNEndtBordxDate <= @to))
    OPTION (RECOMPILE);
END
GO

-- ComboListing / cboMCOGrpCom_Click / the credit note's company block.
CREATE OR ALTER PROCEDURE cis.McoRefund_Lookups @company nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM dbo.PAY;
    SELECT DISTINCT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL ORDER BY 1;
    SELECT DISTINCT LTRIM(RTRIM(MBMDepartment)) AS Department FROM cis.Hdb_MBMOthers
    WHERE @company IS NOT NULL AND MBMGroupCompany = @company AND COALESCE(LTRIM(RTRIM(MBMDepartment)), '') <> '' ORDER BY 1;
END
GO

CREATE OR ALTER PROCEDURE cis.McoRefund_Company @payor nvarchar(10) = NULL, @company nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF @payor IS NOT NULL
        SELECT TOP 1 PAYCompanyName AS Name, PAYAddress1 AS Add1, PAYAddress2 AS Add2, PAYAddress3 AS Add3, PAYTELno AS Tel, PAYMAINContact AS Contact
        FROM dbo.PAY WHERE PAYCode = @payor;
    ELSE
        SELECT TOP 1 GRPCoName AS Name, GRPCoAdd1 AS Add1, GRPCoAdd2 AS Add2, GRPCoAdd3 AS Add3, GRPCoPhone AS Tel, GRPCoContact AS Contact
        FROM dbo.GRPCompany WHERE GRPCoName = @company;
END
GO

-- GenerateCLNMCOCreditNo: "C" + the next number of HISMaintenance CLNMCOCRNoteSeq, locked for the save.
CREATE OR ALTER PROCEDURE cis.McoRefund_NextCreditNote @creditNote nvarchar(10) OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM cis.Maint_CLNMCOCRNoteSeq WITH (UPDLOCK, HOLDLOCK))
    BEGIN
        INSERT cis.Maint_CLNMCOCRNoteSeq (CLNCRNote) VALUES ('00001');
        SET @creditNote = 'C00001';
        RETURN;
    END
    DECLARE @n nvarchar(10);
    UPDATE TOP (1) cis.Maint_CLNMCOCRNoteSeq SET @n = CLNCRNote = RIGHT('00000' + CAST(CAST(CLNCRNote AS int) + 1 AS varchar(10)), 5);
    SET @creditNote = 'C' + @n;
END
GO

-- UpdateCreditNote: the member and all its covers.
CREATE OR ALTER PROCEDURE cis.McoRefund_StampCreditNote @member nvarchar(30), @creditNote nvarchar(10), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @today datetime = CAST(GETDATE() AS date);
    UPDATE cis.Hdb_MBMCLNOthers SET MBMCLNCRNoteDate = @today, MBMCLNCRNote = @creditNote, MBMCLNCRNoteby = @user WHERE MBMNumber = @member;
    UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCCLNCRNoteDate = @today, MBMCCLNCRNote = @creditNote, MBMCCLNCRNoteby = @user WHERE MBMCNumber = @member;
END
GO
