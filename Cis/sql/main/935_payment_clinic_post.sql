-- Payment > Pay Clinic Post (frmPayClinicPost): CISPayment VIEW_PayClinic_Posting with the pasted criteria replaced by parameters, and the posting flag.
SET QUOTED_IDENTIFIER ON;
GO

-- @posted 0 / 1 as cboPostStatus (NULL: not chosen).
CREATE OR ALTER PROCEDURE cis.PayClnPost_Search @payCode nvarchar(10) = NULL, @invFrom nvarchar(50) = NULL, @invTo nvarchar(50) = NULL, @regFrom datetime = NULL,
    @regTo datetime = NULL, @posted char(1) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT PAYCode, PVCLNNo, CLNPaidAmt, InvoiceNo, MBMGroupCompany, CASPayCLNPostUser, CASPayCLNPostDate, InvoiceDate, BordxRefNo, ClinicCode, CLNName, RegisterDate,
        GRPCoCode, CASPayCLNPost, PVCLNDate, PVCLNChequeNo
    FROM CISPayment.dbo.VIEW_PayClinic_Posting
    WHERE (@payCode IS NULL OR PAYCode = @payCode) AND (@invFrom IS NULL OR InvoiceNo >= @invFrom) AND (@invTo IS NULL OR InvoiceNo <= @invTo)
      AND (@regFrom IS NULL OR RegisterDate >= @regFrom) AND (@regTo IS NULL OR RegisterDate <= @regTo) AND (@posted IS NULL OR CASPayCLNPost = @posted)
    OPTION (RECOMPILE);
END
GO

-- Post: the cases of an invoice and clinic not yet posted are marked as posted.
CREATE OR ALTER PROCEDURE cis.PayClnPost_Mark @invoiceNo nvarchar(50), @clinicCode nvarchar(10), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCas SET CASPayCLNPost = '1', CASPayCLNPostUser = @user, CASPayCLNPostDate = CAST(GETDATE() AS date)
    WHERE InvoiceNo = @invoiceNo AND ClinicCode = @clinicCode AND COALESCE(CASPayCLNPost, '0') = '0';
END
GO
