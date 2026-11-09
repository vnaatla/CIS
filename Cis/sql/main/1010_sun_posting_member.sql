-- A/C Export > Invoice Posting > Member (frmPostingInvoiceMember): SUN_CIS_Posting 'InvMBM' (VIEW_SUN_Posting_InvMBM) with the pasted criteria replaced by parameters,
-- the covered person's name for the line description, and the SUNPosting row of a posted case.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.SunPost_InvMbm_Search @batchNo nvarchar(20) = NULL, @company nvarchar(200) = NULL, @invDateFrom datetime = NULL, @invDateTo datetime = NULL,
    @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL, @posted char(1) = '0', @postDateFrom datetime = NULL, @postDateTo datetime = NULL,
    @batchFrom nvarchar(20) = NULL, @batchTo nvarchar(20) = NULL, @bordxDateFrom datetime = NULL, @bordxDateTo datetime = NULL, @bordxFrom nvarchar(20) = NULL,
    @bordxTo nvarchar(20) = NULL, @regMonth nvarchar(2) = NULL, @regYear nvarchar(4) = NULL AS
BEGIN
    SET NOCOUNT ON;
    -- The view joins SUNPosting, CISCas and ClaimInvoice without usable indexes for these filters (a 5-day search took over 4 minutes), so the cases are picked
    -- first through the most selective filter given; the full criteria below still apply.
    CREATE TABLE #c (CASNumber nvarchar(30) PRIMARY KEY);
    DECLARE @narrow bit = 1;
    IF @batchNo IS NOT NULL
        INSERT #c SELECT DISTINCT TransRef FROM CISWorksheet.dbo.SUNPosting WHERE BatchNo = @batchNo AND PostType = 'InvMBM' AND TransRef IS NOT NULL;
    ELSE IF @invDateFrom IS NOT NULL OR @invDateTo IS NOT NULL OR @caseFrom IS NOT NULL OR @caseTo IS NOT NULL
        INSERT #c SELECT CASNumber FROM dbo.CISCas
        WHERE (@invDateFrom IS NULL OR InvoiceDate >= @invDateFrom) AND (@invDateTo IS NULL OR InvoiceDate <= @invDateTo)
          AND (@caseFrom IS NULL OR CASNumber >= @caseFrom) AND (@caseTo IS NULL OR CASNumber <= @caseTo);
    ELSE IF @batchFrom IS NOT NULL OR @batchTo IS NOT NULL
        INSERT #c SELECT DISTINCT TransRef FROM CISWorksheet.dbo.SUNPosting
        WHERE PostType = 'InvMBM' AND TransRef IS NOT NULL AND (@batchFrom IS NULL OR BatchNo >= @batchFrom) AND (@batchTo IS NULL OR BatchNo <= @batchTo);
    ELSE IF @bordxFrom IS NOT NULL OR @bordxTo IS NOT NULL
        INSERT #c SELECT DISTINCT CASNumber FROM CISWorksheet.dbo.Claim
        WHERE CASNumber IS NOT NULL AND (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo);
    ELSE SET @narrow = 0;
    -- VIEW_SUN_Posting_InvMBM's joins and conditions, starting from the picked cases (the view itself is used unchanged by the desktop). SUNPosting has no index on
    -- TransRef, so its rows for the picked cases are read once first.
    SELECT TransRef, BatchNo, AccYear, AccMonth, PostType, PostStatus, PostDate, InvMBMStatus INTO #s FROM CISWorksheet.dbo.SUNPosting
    WHERE @narrow = 0 OR TransRef IN (SELECT CASNumber FROM #c);
    CREATE INDEX ix ON #s (TransRef);
    SELECT k.CASNumber, k.InvoiceNo, k.InvoiceDate, k.ClinicCode, mo.MBMGroupCompany, g.GRPCoCode, o.PAYCode, ci.INVPayTo,
        CASE WHEN COALESCE(ci.INVBordxRefNo, '') = '' THEN ci.BTBGINVBordxRefNo ELSE ci.INVBordxRefNo END AS INVBordxRefNo, ci.INVAmount, ci.INVEnteredDate,
        s.BatchNo, s.AccYear, s.AccMonth, s.PostType, s.PostStatus, s.TransRef, s.PostDate, b.BordxDate, k.CoverID, k.MBMnumber, x.MBMName, cl.CLMPostageFee,
        g.CoGRPAcccode, o.INSType, s.InvMBMStatus, cl.CLMCoPayColl, MONTH(ci.INVEnteredDate) AS RegMonth, YEAR(ci.INVEnteredDate) AS RegYr, cl.BordxRefNo,
        cl.CLMPayableByMedix2M AS CLMTotalActual
    INTO #v
    FROM dbo.CISCas k
    INNER JOIN HISDB.dbo.MBMCLNOthers o ON o.MBMNumber = k.MBMnumber
    INNER JOIN CISWorksheet.dbo.ClaimInvoice ci ON ci.CASNumber = k.CASNumber
    INNER JOIN CISWorksheet.dbo.Claim cl ON cl.CASNumber = ci.CASNumber
    LEFT JOIN CISWorksheet.dbo.Bordx b ON b.BordxRefNo = CASE WHEN COALESCE(ci.INVBordxRefNo, '') = '' THEN ci.BTBGINVBordxRefNo ELSE ci.INVBordxRefNo END
    LEFT JOIN #s s ON s.TransRef = k.CASNumber
    LEFT JOIN HISMaintenance.dbo.MBMCrossReference x ON x.MBMNumber = k.MBMnumber
    LEFT JOIN (HISDB.dbo.MBMOthers mo INNER JOIN dbo.GRPCompany g ON mo.MBMGroupCompany = g.GRPCoName) ON mo.MBMNumber = k.MBMnumber
    WHERE (@narrow = 0 OR k.CASNumber IN (SELECT CASNumber FROM #c)) AND ci.INVPayTo = 'M' AND cl.CLMPayableByMedix2M <> 0
    OPTION (RECOMPILE);
    SELECT v.MBMNumber, v.MBMName, v.CoverID, v.InvoiceDate, v.InvoiceNo, v.CASNumber, v.MBMGroupCompany, v.GRPCoCode, v.CLMPostageFee, v.PostStatus, v.PostType, v.AccMonth,
        v.AccYear, v.InvMBMStatus, v.INSType, v.PAYCode, v.INVAmount, v.BatchNo, v.INVBordxRefNo, v.BordxDate, v.PostDate, v.CoGRPAcccode, v.CLMTotalActual, v.CLMCoPayColl,
        v.RegMonth, v.RegYr, v.BordxRefNo,
        CASE WHEN RTRIM(v.CoverID) = '00' THEN v.MBMName ELSE
            (SELECT TOP 1 RTRIM(c.MBMCName) FROM cis.Hdb_MBMCoveredPersons c WHERE c.MBMCNumber = RTRIM(v.MBMnumber) AND c.MBMCCoverID = RTRIM(v.CoverID)) END AS PatientName
    FROM #v v
    WHERE (@batchNo IS NULL OR v.BatchNo = @batchNo)
      AND (@batchNo IS NOT NULL OR (
            (@company IS NULL OR v.MBMGroupCompany = @company)
        AND (@invDateFrom IS NULL OR v.InvoiceDate >= @invDateFrom) AND (@invDateTo IS NULL OR v.InvoiceDate <= @invDateTo)
        AND (@caseFrom IS NULL OR v.CASNumber >= @caseFrom) AND (@caseTo IS NULL OR v.CASNumber <= @caseTo)
        AND (@posted <> '0' OR ((v.InvMBMStatus IS NULL OR v.InvMBMStatus = '0') AND (v.PostStatus IS NULL OR v.PostStatus = '0')))
        AND (@posted <> '1' OR (v.InvMBMStatus = '1' AND v.PostStatus = '1'))
        AND (@postDateFrom IS NULL OR v.PostDate >= @postDateFrom) AND (@postDateTo IS NULL OR v.PostDate <= @postDateTo)
        AND (@batchFrom IS NULL OR v.BatchNo >= @batchFrom) AND (@batchTo IS NULL OR v.BatchNo <= @batchTo)
        AND (@bordxDateFrom IS NULL OR v.BordxDate >= @bordxDateFrom) AND (@bordxDateTo IS NULL OR v.BordxDate <= @bordxDateTo)
        AND (@bordxFrom IS NULL OR v.BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR v.BordxRefNo <= @bordxTo)
        AND (@regMonth IS NULL OR v.RegMonth = @regMonth) AND (@regYear IS NULL OR v.RegYr = @regYear)))
    OPTION (RECOMPILE);
END
GO

-- CmdPost_Click: InvoiceNo takes the listed row's batch no. (blank before posting) and TransRef the case no., as the desktop.
CREATE OR ALTER PROCEDURE cis.SunPost_InvMbm_Insert @casNumber nvarchar(30), @invoiceNo nvarchar(100) = NULL, @accMonth nvarchar(2), @accYear nvarchar(4),
    @batchNo nvarchar(20), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISWorksheet.dbo.SUNPosting (InvoiceNo, TransRef, InvMBMStatus, PostStatus, PostType, PostDate, PostBy, AccMonth, AccYear, BatchNo)
    VALUES (COALESCE(@invoiceNo, ''), @casNumber, '1', '1', 'InvMBM', CAST(GETDATE() AS date), @user, @accMonth, @accYear, @batchNo);
END
GO

-- The group companies (CboGrpCom).
CREATE OR ALTER PROCEDURE cis.SunPost_Companies AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoName) FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL ORDER BY GRPCoName;
END
GO
