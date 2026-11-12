-- A/C Export > Invoice Posting (frmInvPosting): VIEW_InvSummary_Posting / VIEW_Inv_Posting (CISWorksheet) with the pasted criteria replaced by parameters, and the
-- invoice posting flag (CISCas.CASInvPost).
SET QUOTED_IDENTIFIER ON;
GO

-- INVPostListing: the summary rows, then the detail rows of each (invoice no. and the register day; the desktop compared the detail's full register date with
-- the day, so a registration with a time had no debit line). The register date "to" includes the whole day.
CREATE OR ALTER PROCEDURE cis.InvPost_Search @payCode nvarchar(10) = NULL, @invFrom nvarchar(100) = NULL, @invTo nvarchar(100) = NULL, @regFrom datetime = NULL,
    @regTo datetime = NULL, @posted char(1) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT PAYCode, CLMPayableByMedix2C, InvoiceNo, CASInvPostUser, CASInvPostDate, InvoiceDate, ClinicCode, RegisterDate, CASInvPost, GRPCoCode, BordxRefNo
    INTO #s
    FROM CISWorksheet.dbo.VIEW_InvSummary_Posting
    WHERE (@payCode IS NULL OR PAYCode = @payCode) AND (@invFrom IS NULL OR InvoiceNo >= @invFrom) AND (@invTo IS NULL OR InvoiceNo <= @invTo)
      AND (@regFrom IS NULL OR RegisterDate >= @regFrom) AND (@regTo IS NULL OR RegisterDate < DATEADD(day, 1, CAST(@regTo AS date))) AND (@posted IS NULL OR CASInvPost = @posted)
    OPTION (RECOMPILE);
    SELECT * FROM #s;
    SELECT d.InvoiceNo AS SumInvoiceNo, s.RegisterDate AS SumRegisterDate, d.PAYCode, d.CLMPayableByMedix2C, d.InvoiceNo, d.CASInvPostUser, d.CASInvPostDate, d.InvoiceDate,
        d.BordxRefNo, d.ClinicCode, d.CLNName, d.RegisterDate, d.GRPCoCode, d.CASInvPost
    FROM CISWorksheet.dbo.VIEW_Inv_Posting d
    JOIN #s s ON d.InvoiceNo = s.InvoiceNo AND d.RegisterDate >= CAST(CAST(s.RegisterDate AS date) AS datetime)
        AND d.RegisterDate < DATEADD(day, 1, CAST(CAST(s.RegisterDate AS date) AS datetime))
    WHERE d.RegisterDate >= (SELECT CAST(CAST(MIN(RegisterDate) AS date) AS datetime) FROM #s)
      AND d.RegisterDate < (SELECT DATEADD(day, 1, CAST(CAST(MAX(RegisterDate) AS date) AS datetime)) FROM #s)
    OPTION (RECOMPILE);
END
GO

-- INVPostExcel: the rows posted (VIEW_Inv_Posting with the status and the criteria).
CREATE OR ALTER PROCEDURE cis.InvPost_Rows @payCode nvarchar(10) = NULL, @invFrom nvarchar(100) = NULL, @invTo nvarchar(100) = NULL, @regFrom datetime = NULL,
    @regTo datetime = NULL, @posted char(1) = '0' AS
BEGIN
    SET NOCOUNT ON;
    SELECT PAYCode, CLMPayableByMedix2C, InvoiceNo, MBMGroupCompany, InvoiceDate, BordxRefNo, ClinicCode, CLNName, RegisterDate, GRPCoCode, CASInvPost
    FROM CISWorksheet.dbo.VIEW_Inv_Posting
    WHERE CASInvPost = @posted AND (@payCode IS NULL OR PAYCode = @payCode) AND (@invFrom IS NULL OR InvoiceNo >= @invFrom) AND (@invTo IS NULL OR InvoiceNo <= @invTo)
      AND (@regFrom IS NULL OR RegisterDate >= @regFrom) AND (@regTo IS NULL OR RegisterDate < DATEADD(day, 1, CAST(@regTo AS date)))
    OPTION (RECOMPILE);
END
GO

-- UpdateCase: the cases of an invoice and clinic not yet posted.
CREATE OR ALTER PROCEDURE cis.InvPost_Mark @invoiceNo nvarchar(100), @clinicCode nvarchar(10), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCas SET CASInvPost = '1', CASInvPostUser = @user, CASInvPostDate = CAST(GETDATE() AS date)
    WHERE InvoiceNo = @invoiceNo AND ClinicCode = @clinicCode AND CASInvPost = '0';
END
GO
