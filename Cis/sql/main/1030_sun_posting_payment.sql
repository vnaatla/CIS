-- A/C Export > Payment Posting > Clinic (FrmPostingPay2Cln): SUN_CIS_Posting 'Pay2Cln' / 'Pay2ClnDone' (VIEW_SUN_Posting_Pay2Cln / _Done) with the pasted criteria
-- replaced by parameters, and the SUNPosting row of a posted PV.
SET QUOTED_IDENTIFIER ON;
GO

-- @done lists through VIEW_SUN_Posting_Pay2Cln_Done (the desktop's view after a post); @sj 0 Health Connect (payors other than SJ) / 1 Sihat Johor (SJ). @company is not
-- applied: the views have no group company column (the desktop's filter failed).
CREATE OR ALTER PROCEDURE cis.SunPost_Pay2Cln_Search @done bit = 0, @batchNo nvarchar(20) = NULL, @clinic nvarchar(10) = NULL, @pvDateFrom datetime = NULL,
    @pvDateTo datetime = NULL, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL, @posted char(1) = '0', @postDateFrom datetime = NULL, @postDateTo datetime = NULL,
    @batchFrom nvarchar(20) = NULL, @batchTo nvarchar(20) = NULL, @sj bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 0 PVCLNNo, PVCLNDate, PVCLNChequeNo, CLNPaidAmt, CLNAccCode, CLNName, Pay2ClnStatus, PostStatus, PostType, PostDate, PayCode, ClinicCode, AccMonth, AccYear,
        BatchNo, TransRef INTO #r FROM CISWorksheet.dbo.VIEW_SUN_Posting_Pay2Cln;
    IF @done = 1
        INSERT #r SELECT PVCLNNo, PVCLNDate, PVCLNChequeNo, CLNPaidAmt, CLNAccCode, CLNName, Pay2ClnStatus, PostStatus, PostType, PostDate, PayCode, ClinicCode, AccMonth,
            AccYear, BatchNo, TransRef
        FROM CISWorksheet.dbo.VIEW_SUN_Posting_Pay2Cln_Done
        WHERE (@batchNo IS NULL OR BatchNo = @batchNo) AND (@batchNo IS NOT NULL OR (
                (@clinic IS NULL OR ClinicCode = @clinic) AND (@pvDateFrom IS NULL OR PVCLNDate >= @pvDateFrom) AND (@pvDateTo IS NULL OR PVCLNDate <= @pvDateTo)
            AND (@pvFrom IS NULL OR PVCLNNo >= @pvFrom) AND (@pvTo IS NULL OR PVCLNNo <= @pvTo)
            AND (@posted <> '0' OR ((Pay2ClnStatus IS NULL OR Pay2ClnStatus = '0') AND (PostStatus IS NULL OR PostStatus = '0')))
            AND (@posted <> '1' OR (Pay2ClnStatus = '1' AND PostStatus = '1'))
            AND (@postDateFrom IS NULL OR PostDate >= @postDateFrom) AND (@postDateTo IS NULL OR PostDate <= @postDateTo)
            AND (@batchFrom IS NULL OR BatchNo >= @batchFrom) AND (@batchTo IS NULL OR BatchNo <= @batchTo)
            AND ((@sj = 0 AND SUBSTRING(PayCode, 1, 2) <> 'SJ') OR (@sj = 1 AND SUBSTRING(PayCode, 1, 2) = 'SJ'))))
        OPTION (RECOMPILE);
    ELSE
        INSERT #r SELECT PVCLNNo, PVCLNDate, PVCLNChequeNo, CLNPaidAmt, CLNAccCode, CLNName, Pay2ClnStatus, PostStatus, PostType, PostDate, PayCode, ClinicCode, AccMonth,
            AccYear, BatchNo, TransRef
        FROM CISWorksheet.dbo.VIEW_SUN_Posting_Pay2Cln
        WHERE (@batchNo IS NULL OR BatchNo = @batchNo) AND (@batchNo IS NOT NULL OR (
                (@clinic IS NULL OR ClinicCode = @clinic) AND (@pvDateFrom IS NULL OR PVCLNDate >= @pvDateFrom) AND (@pvDateTo IS NULL OR PVCLNDate <= @pvDateTo)
            AND (@pvFrom IS NULL OR PVCLNNo >= @pvFrom) AND (@pvTo IS NULL OR PVCLNNo <= @pvTo)
            AND (@posted <> '0' OR ((Pay2ClnStatus IS NULL OR Pay2ClnStatus = '0') AND (PostStatus IS NULL OR PostStatus = '0')))
            AND (@posted <> '1' OR (Pay2ClnStatus = '1' AND PostStatus = '1'))
            AND (@postDateFrom IS NULL OR PostDate >= @postDateFrom) AND (@postDateTo IS NULL OR PostDate <= @postDateTo)
            AND (@batchFrom IS NULL OR BatchNo >= @batchFrom) AND (@batchTo IS NULL OR BatchNo <= @batchTo)
            AND ((@sj = 0 AND SUBSTRING(PayCode, 1, 2) <> 'SJ') OR (@sj = 1 AND SUBSTRING(PayCode, 1, 2) = 'SJ'))))
        OPTION (RECOMPILE);
    SELECT * FROM #r;
END
GO

-- CmdPost_Click: InvoiceNo takes the line's analysis code ("4" + clinic) and TransRef the PV no., as the desktop.
CREATE OR ALTER PROCEDURE cis.SunPost_Pay2Cln_Insert @pvNo nvarchar(30), @invoiceNo nvarchar(30), @accMonth nvarchar(2), @accYear nvarchar(4), @batchNo nvarchar(20),
    @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISWorksheet.dbo.SUNPosting (InvoiceNo, TransRef, Pay2ClnStatus, PostStatus, PostType, PostDate, PostBy, AccMonth, AccYear, BatchNo)
    VALUES (@invoiceNo, @pvNo, '1', '1', 'Pay2Cln', CAST(GETDATE() AS date), @user, @accMonth, @accYear, @batchNo);
END
GO
