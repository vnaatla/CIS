-- A/C Export: the posting batch numbers (Medix.bas GenSUNPostingBatchNo / GeneratePostingINVNo) and Invoice Posting > Clinic (frmPostingInvoiceCln):
-- CISWorksheet SUN_CIS_Posting 'InvCln' (VIEW_SUN_Posting_InvCln) with the pasted criteria replaced by parameters, and the SUNPosting row of a posted invoice.
SET QUOTED_IDENTIFIER ON;
GO

-- GenSUNPostingBatchNo: SunPostBatch holds the last batch no. ('1001' when empty); the row is locked until the caller's transaction ends.
CREATE OR ALTER PROCEDURE cis.SunPost_NextBatch @batchNo nvarchar(20) OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @last nvarchar(20) = (SELECT TOP 1 BatchNo FROM CISWorksheet.dbo.SunPostBatch WITH (UPDLOCK, HOLDLOCK) ORDER BY BatchIndex DESC);
    IF @last IS NULL
    BEGIN
        SET @batchNo = '1001';
        INSERT CISWorksheet.dbo.SunPostBatch (BatchNo) VALUES (@batchNo);
    END
    ELSE
    BEGIN
        SET @batchNo = CAST(CAST(@last AS int) + 1 AS nvarchar(20));
        UPDATE CISWorksheet.dbo.SunPostBatch SET BatchNo = @batchNo;
    END
END
GO

-- GeneratePostingINVNo: PostingINVSEQ holds a letter and 6 digits ("A000001" when empty); 999999 moves to the next letter.
CREATE OR ALTER PROCEDURE cis.SunPost_NextInvNo @seqNo nvarchar(20) OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @last nvarchar(20) = (SELECT TOP 1 SEQNumber FROM CISWorksheet.dbo.PostingINVSEQ WITH (UPDLOCK, HOLDLOCK));
    IF @last IS NULL
    BEGIN
        SET @seqNo = 'A000001';
        INSERT CISWorksheet.dbo.PostingINVSEQ (SEQNumber) VALUES (@seqNo);
        RETURN;
    END
    DECLARE @ch char(1) = LEFT(@last, 1), @n int = CAST(SUBSTRING(@last, 2, 6) AS int);
    IF @n = 999999 SELECT @ch = CHAR(ASCII(@ch) + 1), @n = 1 ELSE SET @n = @n + 1;
    SET @seqNo = @ch + RIGHT('000000' + CAST(@n AS varchar(6)), 6);
    UPDATE CISWorksheet.dbo.PostingINVSEQ SET SEQNumber = @seqNo;
END
GO

-- SearchRecord: @posted 0 / 1 as cboPostStatus. The bordx month and year are compared as one period (the desktop compared them separately, so 12/2020 – 01/2021
-- found nothing). @company is not applied: the view has no
-- group company column (the desktop's filter on MBMGroupCompany failed).
CREATE OR ALTER PROCEDURE cis.SunPost_InvCln_Search @batchNo nvarchar(20) = NULL, @clinic nvarchar(10) = NULL, @company nvarchar(200) = NULL,
    @invDateFrom datetime = NULL, @invDateTo datetime = NULL, @invFrom nvarchar(100) = NULL, @invTo nvarchar(100) = NULL, @posted char(1) = '0',
    @postDateFrom datetime = NULL, @postDateTo datetime = NULL, @batchFrom nvarchar(20) = NULL, @batchTo nvarchar(20) = NULL, @bordxMonthFrom nvarchar(2) = NULL,
    @bordxMonthTo nvarchar(2) = NULL, @bordxYearFrom nvarchar(4) = NULL, @bordxYearTo nvarchar(4) = NULL, @accMonth nvarchar(2) = NULL, @accYear nvarchar(4) = NULL,
    @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT ClnName, InvoiceDate, InvoiceNo, CLNAccCode, BordxMonth, BordxYear, PostStatus, PostType, AccMonth, AccYear, InvCLNStatus, PAYCode, INVPayTo, INVAmount,
        ClinicCode, BatchNo, PostDate, CLMPayableByMedix2C, CLMTotalActual, CLMCoPayColl, SUNInvoiceNo, CLMMCOProcFee, CLMFFSNoGST, CLMFFSGST, PVCLNDate, CLNInvoiceBatch
    FROM CISWorksheet.dbo.VIEW_SUN_Posting_InvCln
    WHERE (@batchNo IS NULL OR BatchNo = @batchNo)
      AND (@batchNo IS NOT NULL OR (
            (@clinic IS NULL OR ClinicCode = @clinic)
        AND (@invDateFrom IS NULL OR InvoiceDate >= @invDateFrom) AND (@invDateTo IS NULL OR InvoiceDate <= @invDateTo)
        AND (@invFrom IS NULL OR InvoiceNo >= @invFrom) AND (@invTo IS NULL OR InvoiceNo <= @invTo)
        AND (@posted <> '0' OR ((InvCLNStatus IS NULL OR InvCLNStatus = '0') AND (PostStatus IS NULL OR PostStatus = '0')))
        AND (@posted <> '1' OR (InvCLNStatus = '1' AND PostStatus = '1'))
        AND (@postDateFrom IS NULL OR PostDate >= @postDateFrom) AND (@postDateTo IS NULL OR PostDate <= @postDateTo)
        AND (@batchFrom IS NULL OR BatchNo >= @batchFrom) AND (@batchTo IS NULL OR BatchNo <= @batchTo)
        AND (@bordxYearFrom IS NULL OR TRY_CAST(BordxYear AS int) * 100 + TRY_CAST(BordxMonth AS int)
            >= TRY_CAST(@bordxYearFrom AS int) * 100 + TRY_CAST(@bordxMonthFrom AS int))
        AND (@bordxYearTo IS NULL OR TRY_CAST(BordxYear AS int) * 100 + TRY_CAST(BordxMonth AS int)
            <= TRY_CAST(@bordxYearTo AS int) * 100 + TRY_CAST(@bordxMonthTo AS int))
        AND (@accMonth IS NULL OR AccMonth = @accMonth) AND (@accYear IS NULL OR AccYear = @accYear)
        AND (@pvDateFrom IS NULL OR PVCLNDate >= @pvDateFrom) AND (@pvDateTo IS NULL OR PVCLNDate <= @pvDateTo)))
    OPTION (RECOMPILE);
END
GO

-- CmdPost_Click: the SUNPosting row of one posted clinic invoice.
CREATE OR ALTER PROCEDURE cis.SunPost_InvCln_Insert @invoiceNo nvarchar(100), @clinic nvarchar(10), @accMonth nvarchar(2), @accYear nvarchar(4), @batchNo nvarchar(20),
    @invDate datetime = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISWorksheet.dbo.SUNPosting (InvoiceNo, TransRef, InvCLNStatus, PostStatus, PostType, PostDate, PostBy, AccMonth, AccYear, BatchNo, CLNInvDate)
    VALUES (@invoiceNo, @clinic, '1', '1', 'InvCln', CAST(GETDATE() AS date), @user, @accMonth, @accYear, @batchNo, @invDate);
END
GO

-- The accounting periods (HISMaintenance AccPeriod, "MM/YYYY").
CREATE OR ALTER PROCEDURE cis.SunPost_Periods AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(AccMonth) + '/' + RTRIM(AccYear) AS Period FROM HISMaintenance.dbo.AccPeriod ORDER BY AccYear DESC, AccMonth DESC;
END
GO

-- Invoice Posting > Clinic (GST) (frmPostingInvoiceClnGST): SUN_CIS_Posting 'InvClnGST' (VIEW_SUN_Posting_InvClnGST). @batchNo lists a posted batch; @company is not
-- applied (the view has no group company column, so the desktop's filter failed).
CREATE OR ALTER PROCEDURE cis.SunPost_InvClnGST_Search @batchNo nvarchar(20) = NULL, @clinic nvarchar(10) = NULL, @invDateFrom datetime = NULL, @invDateTo datetime = NULL,
    @invFrom nvarchar(100) = NULL, @invTo nvarchar(100) = NULL, @posted char(1) = '0', @postDateFrom datetime = NULL, @postDateTo datetime = NULL,
    @batchFrom nvarchar(20) = NULL, @batchTo nvarchar(20) = NULL, @sunFrom nvarchar(20) = NULL, @sunTo nvarchar(20) = NULL, @accMonth nvarchar(2) = NULL,
    @accYear nvarchar(4) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT ClnName, InvoiceDate, InvoiceNo, CLNAccCode, PostStatus, PostType, AccMonth, AccYear, InvCLNStatus, PAYCode, INVPayTo, INVAmount, ClinicCode, BatchNo, PostDate,
        CLMPayableByMedix2C, CLMTotalActual, CLMCoPayColl, SUNInvoiceNo, CLMMCOProcFee, CLMFFSNoGST, CLMFFSGST, CLNInvoiceBatch
    FROM CISWorksheet.dbo.VIEW_SUN_Posting_InvClnGST
    WHERE (@batchNo IS NULL OR BatchNo = @batchNo)
      AND (@batchNo IS NOT NULL OR (
            (@clinic IS NULL OR ClinicCode = @clinic)
        AND (@invDateFrom IS NULL OR InvoiceDate >= @invDateFrom) AND (@invDateTo IS NULL OR InvoiceDate <= @invDateTo)
        AND (@invFrom IS NULL OR InvoiceNo >= @invFrom) AND (@invTo IS NULL OR InvoiceNo <= @invTo)
        AND (@posted <> '0' OR ((InvCLNStatus IS NULL OR InvCLNStatus = '0') AND (PostStatus IS NULL OR PostStatus = '0')))
        AND (@posted <> '1' OR (InvCLNStatus = '1' AND PostStatus = '1'))
        AND (@postDateFrom IS NULL OR PostDate >= @postDateFrom) AND (@postDateTo IS NULL OR PostDate <= @postDateTo)
        AND (@batchFrom IS NULL OR BatchNo >= @batchFrom) AND (@batchTo IS NULL OR BatchNo <= @batchTo)
        AND (@sunFrom IS NULL OR CLNInvoiceBatch >= @sunFrom) AND (@sunTo IS NULL OR CLNInvoiceBatch <= @sunTo)
        AND (@accMonth IS NULL OR AccMonth = @accMonth) AND (@accYear IS NULL OR AccYear = @accYear)))
    OPTION (RECOMPILE);
END
GO

-- CmdPost_Click: one SUNPosting row per posted invoice, with its own SUN invoice batch (I / T + PostingINVSEQ) under the post's batch no. (SunPostBatch).
CREATE OR ALTER PROCEDURE cis.SunPost_InvClnGST_Insert @invoiceNo nvarchar(100), @clinic nvarchar(10), @accMonth nvarchar(2), @accYear nvarchar(4),
    @invBatch nvarchar(20), @invDate datetime = NULL, @batchNo nvarchar(20), @user nvarchar(50), @postType nvarchar(20) = 'InvClnGST' AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISWorksheet.dbo.SUNPosting (InvoiceNo, TransRef, InvCLNStatus, PostStatus, PostType, PostDate, PostBy, AccMonth, AccYear, CLNInvoiceBatch, CLNInvDate, BatchNo)
    VALUES (@invoiceNo, @clinic, '1', '1', @postType, CAST(GETDATE() AS date), @user, @accMonth, @accYear, @invBatch, @invDate, @batchNo);
END
GO

-- Invoice Posting > Clinic (GST-2017) (frmPostingInvClnGST2017): SUN_CIS_Posting 'InvClnGST2017' (VIEW_SUN_Posting_InvClnGST_2017); posted = PostStatus only.
CREATE OR ALTER PROCEDURE cis.SunPost_InvClnGST2017_Search @batchNo nvarchar(20) = NULL, @invBatch nvarchar(20) = NULL, @clinic nvarchar(10) = NULL,
    @invDateFrom datetime = NULL, @invDateTo datetime = NULL, @invFrom nvarchar(100) = NULL, @invTo nvarchar(100) = NULL, @posted char(1) = '0',
    @postDateFrom datetime = NULL, @postDateTo datetime = NULL, @batchFrom nvarchar(20) = NULL, @batchTo nvarchar(20) = NULL, @sunFrom nvarchar(20) = NULL,
    @sunTo nvarchar(20) = NULL, @accMonth nvarchar(2) = NULL, @accYear nvarchar(4) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT ClnName, InvoiceDate, InvoiceNo, CLNAccCode, PostStatus, PostType, AccMonth, AccYear, InvCLNStatus, INVPayTo, INVAmount, ClinicCode, BatchNo, PostDate,
        CLMPayableByMedix2C, CLMTotalActual, CLMCoPayColl, CLMMCOProcFee, CLMFFSNoGST, CLMFFSGST, CLNInvoiceBatch, SUNPosting
    FROM CISWorksheet.dbo.VIEW_SUN_Posting_InvClnGST_2017
    WHERE (@batchNo IS NULL OR BatchNo = @batchNo) AND (@invBatch IS NULL OR CLNInvoiceBatch = @invBatch)
      AND (@batchNo IS NOT NULL OR @invBatch IS NOT NULL OR (
            (@clinic IS NULL OR ClinicCode = @clinic)
        AND (@invDateFrom IS NULL OR InvoiceDate >= @invDateFrom) AND (@invDateTo IS NULL OR InvoiceDate <= @invDateTo)
        AND (@invFrom IS NULL OR InvoiceNo >= @invFrom) AND (@invTo IS NULL OR InvoiceNo <= @invTo)
        AND (@posted <> '0' OR PostStatus IS NULL OR PostStatus = '0') AND (@posted <> '1' OR PostStatus = '1')
        AND (@postDateFrom IS NULL OR PostDate >= @postDateFrom) AND (@postDateTo IS NULL OR PostDate <= @postDateTo)
        AND (@batchFrom IS NULL OR BatchNo >= @batchFrom) AND (@batchTo IS NULL OR BatchNo <= @batchTo)
        AND (@sunFrom IS NULL OR CLNInvoiceBatch >= @sunFrom) AND (@sunTo IS NULL OR CLNInvoiceBatch <= @sunTo)
        AND (@accMonth IS NULL OR AccMonth = @accMonth) AND (@accYear IS NULL OR AccYear = @accYear)))
    OPTION (RECOMPILE);
END
GO

-- CmdPost_Click (GST-2017): the cases of an invoice; each gets its own SUNPosting row.
CREATE OR ALTER PROCEDURE cis.SunPost_InvClnGST2017_Cases @invoiceNo nvarchar(100), @clinic nvarchar(10), @invDate datetime = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(CASNumber) FROM dbo.CISCas WHERE InvoiceNo = @invoiceNo AND ClinicCode = @clinic AND InvoiceDate = @invDate;
END
GO

-- cmdgenerate_Click: the clinic block of the FFS tax invoice.
CREATE OR ALTER PROCEDURE cis.SunPost_TaxInvoiceClinic @clinic nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 CLNRepresentative, CLNName, CLNAdd1, CLNAdd2, CLNAdd3, CLNPhone, CLNCode FROM cis.Hdb_CLN WHERE CLNCode = @clinic;
END
GO
