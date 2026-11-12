-- A/C Export > Receipt Posting > Grp Company (frmPostingRecGrpCompany): SUN_CIS_Posting 'RecGrp' (VIEW_Sun_Posting_RecFrGrpComp) with the pasted criteria replaced by
-- parameters, and the SUNPosting row of a posted receipt.
SET QUOTED_IDENTIFIER ON;
GO

-- The view takes minutes (114 s for 3 rows of two months), so the receipts are picked first (batch, receipt date, receipt no.) and the view's joins and grouping are
-- applied to them. @sj: 0 payors other than SJ, 1 SJ.
CREATE OR ALTER PROCEDURE cis.SunPost_RecGrp_Search @batchNo nvarchar(20) = NULL, @company nvarchar(200) = NULL, @recFrom nvarchar(30) = NULL, @recTo nvarchar(30) = NULL,
    @recDateFrom datetime = NULL, @recDateTo datetime = NULL, @postDateFrom datetime = NULL, @postDateTo datetime = NULL, @posted char(1) = '0',
    @batchFrom nvarchar(20) = NULL, @batchTo nvarchar(20) = NULL, @sj bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    CREATE TABLE #n (ReceiptNo nvarchar(30) PRIMARY KEY);
    IF @batchNo IS NOT NULL
        INSERT #n SELECT DISTINCT TransRef FROM CISWorksheet.dbo.SUNPosting WHERE BatchNo = @batchNo AND PostType = 'RecGrp' AND TransRef IS NOT NULL;
    ELSE IF @batchFrom IS NOT NULL OR @batchTo IS NOT NULL
        INSERT #n SELECT DISTINCT TransRef FROM CISWorksheet.dbo.SUNPosting
        WHERE PostType = 'RecGrp' AND TransRef IS NOT NULL AND (@batchFrom IS NULL OR BatchNo >= @batchFrom) AND (@batchTo IS NULL OR BatchNo <= @batchTo);
    ELSE
        INSERT #n SELECT DISTINCT ReceiptNo FROM CISPayment.dbo.ReceiptfrPAY
        WHERE ReceiptNo IS NOT NULL AND (@recDateFrom IS NULL OR ReceiptDate >= @recDateFrom) AND (@recDateTo IS NULL OR ReceiptDate <= @recDateTo)
          AND (@recFrom IS NULL OR ReceiptNo >= @recFrom) AND (@recTo IS NULL OR ReceiptNo <= @recTo);
    SELECT TransRef, RecGrpCompStatus, PostStatus, PostType, PostDate, PostBy, AccMonth, AccYear, BatchNo INTO #s FROM CISWorksheet.dbo.SUNPosting
    WHERE TransRef IN (SELECT ReceiptNo FROM #n);
    CREATE INDEX ix ON #s (TransRef);

    SELECT r.ReceiptNo, r.ReceiptDate, mo.MBMGroupCompany, o.PAYCode AS Expr2, s.RecGrpCompStatus, s.PostStatus, s.PostType, s.PostDate, s.PostBy, s.AccMonth, s.AccYear,
        s.BatchNo, r.ReceiptAmt AS CheckAmt, g.CoGRPAcccode, g.GRPCoCode, g.GRPCoPaycode, cg.CoGRPCode, cg.CoGRPPaycode, p.PAYCode, p.PAYAccCode, cl.BordxRefNo
    INTO #v
    FROM CISPayment.dbo.ReceiptfrPAY r
    JOIN CISWorksheet.dbo.Claim cl ON cl.CASNumber = r.CASNumber
    JOIN dbo.CISCas k ON k.CASNumber = cl.CASNumber
    JOIN HISDB.dbo.MBMCLNOthers o ON o.MBMNumber = k.MBMnumber
    JOIN HISDB.dbo.MBMOthers mo ON mo.MBMNumber = o.MBMNumber
    JOIN dbo.GRPCompany g ON g.GRPCoName = mo.MBMGroupCompany
    JOIN HISDB.dbo.CompanyGRP cg ON cg.CoGRPName = mo.MBMGroupCompany
    JOIN HISDB.dbo.PAY p ON p.PAYCode = cg.CoGRPPaycode
    LEFT JOIN #s s ON s.TransRef = r.ReceiptNo
    WHERE r.ReceiptNo IN (SELECT ReceiptNo FROM #n)
    GROUP BY r.ReceiptNo, r.ReceiptDate, mo.MBMGroupCompany, o.PAYCode, s.RecGrpCompStatus, s.PostStatus, s.PostType, s.PostDate, s.PostBy, s.AccMonth, s.AccYear,
        s.BatchNo, g.CoGRPAcccode, g.GRPCoCode, r.ReceiptAmt, g.GRPCoPaycode, cg.CoGRPCode, cg.CoGRPPaycode, p.PAYCode, p.PAYAccCode, cl.BordxRefNo;

    SELECT * FROM #v
    WHERE (@batchNo IS NULL OR BatchNo = @batchNo) AND (@batchNo IS NOT NULL OR (
            (@company IS NULL OR MBMGroupCompany = @company)
        AND (@recFrom IS NULL OR ReceiptNo >= @recFrom) AND (@recTo IS NULL OR ReceiptNo <= @recTo)
        AND (@recDateFrom IS NULL OR ReceiptDate >= @recDateFrom) AND (@recDateTo IS NULL OR ReceiptDate <= @recDateTo)
        AND (@postDateFrom IS NULL OR PostDate >= @postDateFrom) AND (@postDateTo IS NULL OR PostDate <= @postDateTo)
        AND (@posted <> '0' OR ((RecGrpCompStatus IS NULL OR RecGrpCompStatus = '0') AND (PostStatus IS NULL OR PostStatus = '0')))
        AND (@posted <> '1' OR (RecGrpCompStatus = '1' AND PostStatus = '1'))
        AND (@batchFrom IS NULL OR BatchNo >= @batchFrom) AND (@batchTo IS NULL OR BatchNo <= @batchTo)
        AND ((@sj = 0 AND PAYCode <> 'SJ') OR (@sj = 1 AND PAYCode = 'SJ'))))
    ORDER BY ReceiptNo
    OPTION (RECOMPILE);
END
GO

-- CmdExp_Click: the SUNPosting row of a posted receipt.
CREATE OR ALTER PROCEDURE cis.SunPost_RecGrp_Insert @receiptNo nvarchar(30), @accMonth nvarchar(2), @accYear nvarchar(4), @batchNo nvarchar(20), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISWorksheet.dbo.SUNPosting (TransRef, RecGrpCompStatus, PostStatus, PostType, PostDate, PostBy, AccMonth, AccYear, BatchNo)
    VALUES (@receiptNo, '1', '1', 'RecGrp', CAST(GETDATE() AS date), @user, @accMonth, @accYear, @batchNo);
END
GO
