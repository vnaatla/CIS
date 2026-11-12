-- A/C Export > Bordx Posting > Grp Company (frmPostingBordxtoGrpComp, hidden) and Grp Company (GST) (frmPostingBordxtoGrpGST): SUN_CIS_Posting 'BordxGrp'
-- (VIEW_Sun_Posting_BordxGrpComp) with the pasted criteria replaced by parameters, and the SUNPosting row of a posted bordereaux.
SET QUOTED_IDENTIFIER ON;
GO

-- The view takes minutes even for a few days, so the bordereaux are picked first (batch, bordx date, bordx no.) and the view's two branches (the bordereaux of the
-- claims, and the BTBG bordereaux with the BTBG amounts) are read for them with the view's joins and sums. @sj: NULL any payor, 0 not SJ, 1 SJ. The bordx no. range is
-- applied as given (the GST screen swaps its two boxes before calling).
CREATE OR ALTER PROCEDURE cis.SunPost_BordxGrp_Search @batchNo nvarchar(20) = NULL, @company nvarchar(200) = NULL, @bordxFrom nvarchar(20) = NULL,
    @bordxTo nvarchar(20) = NULL, @bordxDateFrom datetime = NULL, @bordxDateTo datetime = NULL, @postDateFrom datetime = NULL, @postDateTo datetime = NULL,
    @posted char(1) = '0', @batchFrom nvarchar(20) = NULL, @batchTo nvarchar(20) = NULL, @sj bit = NULL AS
BEGIN
    SET NOCOUNT ON;
    CREATE TABLE #b (BordxRefNo nvarchar(20) PRIMARY KEY);
    IF @batchNo IS NOT NULL
        INSERT #b SELECT DISTINCT TransRef FROM CISWorksheet.dbo.SUNPosting WHERE BatchNo = @batchNo AND PostType = 'BordxGrp' AND TransRef IS NOT NULL;
    ELSE IF @batchFrom IS NOT NULL OR @batchTo IS NOT NULL
        INSERT #b SELECT DISTINCT TransRef FROM CISWorksheet.dbo.SUNPosting
        WHERE PostType = 'BordxGrp' AND TransRef IS NOT NULL AND (@batchFrom IS NULL OR BatchNo >= @batchFrom) AND (@batchTo IS NULL OR BatchNo <= @batchTo);
    ELSE
        INSERT #b SELECT BordxRefNo FROM CISWorksheet.dbo.Bordx
        WHERE BordxRefNo IS NOT NULL AND (@bordxDateFrom IS NULL OR BordxDate >= @bordxDateFrom) AND (@bordxDateTo IS NULL OR BordxDate <= @bordxDateTo)
          AND (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
        GROUP BY BordxRefNo;

    SELECT CASNumber, BordxRefNo, BTBGBordxRefNo, CLMMCOProcFee, CLMTotalActual, CLMCoPayColl, CLMPayableByMedix2C, CLMPayableByMedix2M, CLMBordxAmt, CLMBTBGAmt,
        CASE WHEN BordxRefNo IN (SELECT BordxRefNo FROM #b) THEN 1 ELSE 0 END AS InMain,
        CASE WHEN BTBGBordxRefNo IN (SELECT BordxRefNo FROM #b) THEN 1 ELSE 0 END AS InBtbg
    INTO #c FROM CISWorksheet.dbo.Claim
    WHERE BordxRefNo IN (SELECT BordxRefNo FROM #b) OR BTBGBordxRefNo IN (SELECT BordxRefNo FROM #b);
    SELECT TransRef, PostStatus, PostType, PostDate, PostBy, AccMonth, AccYear, BatchNo, BordxGrpCompStatus INTO #s FROM CISWorksheet.dbo.SUNPosting
    WHERE TransRef IN (SELECT BordxRefNo FROM #c);
    CREATE INDEX ix ON #s (TransRef);

    SELECT o.PAYCode, s.PostStatus, s.PostType, s.PostDate, s.PostBy, s.AccMonth, s.AccYear, s.BatchNo, s.BordxGrpCompStatus, g.CoGRPAcccode, g.GRPCoCode, g.GRPMCOStatus,
        b.BordxRefNo, b.BordxDate, b.BORDXPrintMonth, b.BORDXPrintYear, ci.INVPayTo, SUM(c.CLMMCOProcFee) AS CLMMCOProcFee, SUM(c.CLMTotalActual) AS CLMTotalActual,
        SUM(c.CLMCoPayColl) AS CLMCoPayColl, SUM(c.CLMPayableByMedix2C) AS CLMPayableByMedix2C, SUM(c.CLMPayableByMedix2M) AS CLMPayableByMedix2M,
        MONTH(b.BordxDate) AS RegMonth, YEAR(b.BordxDate) AS RegYear,
        CASE WHEN p.PLNBTBG = 1 THEN SUM(c.CLMBordxAmt - COALESCE(c.CLMBTBGAmt, 0)) ELSE SUM(c.CLMBordxAmt) END AS CLMBordxAmt, g.GRPCoName AS MBMGroupCompany,
        g.GRPCoCode AS CoGRPCode, g.GRPCoPaycode AS CoGRPPaycode, h.PAYGRPCode, h.CISPAYAccCode AS PAYAccCode, b.BordxSubsidary
    INTO #v
    FROM #c c
    JOIN dbo.CISCas k ON k.CASNumber = c.CASNumber
    JOIN CISWorksheet.dbo.ClaimInvoice ci ON ci.CASNumber = c.CASNumber
    JOIN CISWorksheet.dbo.Bordx b ON b.BordxRefNo = c.BordxRefNo
    LEFT JOIN #s s ON s.TransRef = c.BordxRefNo
    JOIN HISDB.dbo.MBMCLNOthers o ON o.MBMNumber = k.MBMnumber
    JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    JOIN HISDB.dbo.MBMOthers mo ON mo.MBMNumber = k.MBMnumber
    JOIN HISDB.dbo.MBM m ON m.MBMNumber = mo.MBMNumber
    LEFT JOIN (HISDB.dbo.PAY h JOIN dbo.GRPCompany g ON h.PAYCode = g.ASOINSPaycode) ON g.GRPCoName = mo.MBMGroupCompany
    WHERE c.InMain = 1
    GROUP BY o.PAYCode, s.PostStatus, s.PostType, s.PostDate, s.PostBy, s.AccMonth, s.AccYear, s.BatchNo, g.CoGRPAcccode, g.GRPCoCode, b.BordxRefNo, b.BordxDate,
        b.BORDXPrintMonth, b.BORDXPrintYear, s.BordxGrpCompStatus, ci.INVPayTo, g.GRPMCOStatus, MONTH(b.BordxDate), YEAR(b.BordxDate), g.GRPCoName, g.GRPCoPaycode,
        h.PAYGRPCode, h.CISPAYAccCode, b.BordxSubsidary, p.PLNBTBG
    UNION
    SELECT o.PAYCode, s.PostStatus, s.PostType, s.PostDate, s.PostBy, s.AccMonth, s.AccYear, s.BatchNo, s.BordxGrpCompStatus, g.CoGRPAcccode, g.GRPCoCode, g.GRPMCOStatus,
        b.BordxRefNo, b.BordxDate, b.BORDXPrintMonth, b.BORDXPrintYear, ci.INVPayTo, SUM(c.CLMMCOProcFee), SUM(c.CLMTotalActual), SUM(c.CLMCoPayColl),
        SUM(c.CLMPayableByMedix2C), SUM(c.CLMPayableByMedix2M), MONTH(b.BordxDate), YEAR(b.BordxDate), SUM(c.CLMBTBGAmt), g.GRPCoName, g.GRPCoCode, g.GRPCoPaycode,
        h.PAYGRPCode, h.CISPAYAccCode, b.BordxSubsidary
    FROM #c c
    JOIN dbo.CISCas k ON k.CASNumber = c.CASNumber
    JOIN CISWorksheet.dbo.ClaimInvoice ci ON ci.CASNumber = c.CASNumber
    JOIN CISWorksheet.dbo.Bordx b ON b.BordxRefNo = c.BTBGBordxRefNo
    LEFT JOIN #s s ON s.TransRef = c.BordxRefNo
    JOIN HISDB.dbo.MBMCLNOthers o ON o.MBMNumber = k.MBMnumber
    JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    JOIN HISDB.dbo.MBMOthers mo ON mo.MBMNumber = k.MBMnumber
    JOIN HISDB.dbo.MBM m ON m.MBMNumber = mo.MBMNumber
    LEFT JOIN (HISDB.dbo.PAY h JOIN dbo.GRPCompany g ON h.PAYCode = g.ASOINSPaycode) ON g.GRPCoName = mo.MBMGroupCompany
    WHERE c.InBtbg = 1
    GROUP BY o.PAYCode, s.PostStatus, s.PostType, s.PostDate, s.PostBy, s.AccMonth, s.AccYear, s.BatchNo, g.CoGRPAcccode, g.GRPCoCode, b.BordxRefNo, b.BordxDate,
        b.BORDXPrintMonth, b.BORDXPrintYear, s.BordxGrpCompStatus, ci.INVPayTo, g.GRPMCOStatus, MONTH(b.BordxDate), YEAR(b.BordxDate), g.GRPCoName, g.GRPCoPaycode,
        h.PAYGRPCode, h.CISPAYAccCode, b.BordxSubsidary;

    SELECT * FROM #v
    WHERE (@batchNo IS NULL OR BatchNo = @batchNo) AND (@batchNo IS NOT NULL OR (
            (@company IS NULL OR MBMGroupCompany = @company)
        AND (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
        AND (@bordxDateFrom IS NULL OR BordxDate >= @bordxDateFrom) AND (@bordxDateTo IS NULL OR BordxDate <= @bordxDateTo)
        AND (@postDateFrom IS NULL OR PostDate >= @postDateFrom) AND (@postDateTo IS NULL OR PostDate <= @postDateTo)
        AND (@posted <> '0' OR ((BordxGrpCompStatus IS NULL OR BordxGrpCompStatus = '0') AND (PostStatus IS NULL OR PostStatus = '0')))
        AND (@posted <> '1' OR (BordxGrpCompStatus = '1' AND PostStatus = '1'))
        AND (@batchFrom IS NULL OR BatchNo >= @batchFrom) AND (@batchTo IS NULL OR BatchNo <= @batchTo)
        AND (@sj IS NULL OR (@sj = 0 AND PAYCode <> 'SJ') OR (@sj = 1 AND PAYCode = 'SJ'))))
    ORDER BY BordxRefNo
    OPTION (RECOMPILE);
END
GO

-- CmdExp_Click: the SUNPosting row of a posted bordereaux.
CREATE OR ALTER PROCEDURE cis.SunPost_BordxGrp_Insert @bordxNo nvarchar(20), @accMonth nvarchar(2), @accYear nvarchar(4), @batchNo nvarchar(20), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISWorksheet.dbo.SUNPosting (TransRef, BordxGrpCompStatus, PostStatus, PostType, PostDate, PostBy, AccMonth, AccYear, BatchNo)
    VALUES (@bordxNo, '1', '1', 'BordxGrp', CAST(GETDATE() AS date), @user, @accMonth, @accYear, @batchNo);
END
GO
