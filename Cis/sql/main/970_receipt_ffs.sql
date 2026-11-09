-- Receipt > Fr FFS (FrmReceiptFFS): CISPayment PaymentReceiptModule case 17 (VIEW_CISReceiptFFS) with the pasted criteria replaced by parameters, and UpdateRecord.
SET QUOTED_IDENTIFIER ON;
GO

-- @scan: one bordx no.; @payment 1 received, 2 not received; @receiptNo: the receipt just saved.
CREATE OR ALTER PROCEDURE cis.RcpFfs_Search @receiptNo nvarchar(30) = NULL, @scan bit = 0, @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL,
    @payment int = 0, @recFrom nvarchar(30) = NULL, @recTo nvarchar(30) = NULL, @recDateFrom datetime = NULL, @recDateTo datetime = NULL, @mthFrom nvarchar(2) = NULL,
    @mthTo nvarchar(2) = NULL, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL, @bordxDateFrom datetime = NULL, @bordxDateTo datetime = NULL,
    @company nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF @receiptNo IS NOT NULL
    BEGIN
        -- The view aggregates every claim and only a bordx no. is pushed into it, so the receipt's bordereaux are read one at a time.
        SELECT TOP 0 * INTO #r FROM CISPayment.dbo.VIEW_CISReceiptFFS;
        DECLARE @b nvarchar(20);
        DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT DISTINCT BordxRefNo FROM CISPayment.dbo.ReceiptFrFFS WHERE ReceiptNo = @receiptNo AND BordxRefNo IS NOT NULL;
        OPEN c;
        FETCH NEXT FROM c INTO @b;
        WHILE @@FETCH_STATUS = 0
        BEGIN
            INSERT #r SELECT * FROM CISPayment.dbo.VIEW_CISReceiptFFS WHERE BordxRefNo = @b AND ReceiptNo = @receiptNo OPTION (RECOMPILE);
            FETCH NEXT FROM c INTO @b;
        END
        CLOSE c;
        SELECT * FROM #r ORDER BY BordxRefNo;
        RETURN;
    END
    SELECT * FROM CISPayment.dbo.VIEW_CISReceiptFFS
    WHERE (
            (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
        AND (@scan = 0 OR @bordxFrom IS NULL OR BordxRefNo = @bordxFrom)
        AND (@payment <> 1 OR ReceiptStatus = 'F') AND (@payment <> 2 OR ReceiptStatus = '' OR ReceiptStatus IS NULL)
        AND (@recFrom IS NULL OR ReceiptNo >= @recFrom) AND (@recTo IS NULL OR ReceiptNo <= @recTo)
        AND (@recDateFrom IS NULL OR ReceiptDate >= @recDateFrom) AND (@recDateTo IS NULL OR ReceiptDate <= @recDateTo)
        AND (@mthFrom IS NULL OR ReceiptAccMth >= @mthFrom) AND (@mthTo IS NULL OR ReceiptAccMth <= @mthTo)
        AND (@yrFrom IS NULL OR ReceiptAccYr >= @yrFrom) AND (@yrTo IS NULL OR ReceiptAccYr <= @yrTo)
        AND (@bordxDateFrom IS NULL OR BordxDate >= @bordxDateFrom) AND (@bordxDateTo IS NULL OR BordxDate <= @bordxDateTo)
        AND (@company IS NULL OR GroupCompany = @company))
    OPTION (RECOMPILE);
END
GO

-- UpdateRecord: a new ReceiptFrFFS row (Save) or the row of @index overwritten (Update).
CREATE OR ALTER PROCEDURE cis.RcpFfs_Save @index nvarchar(20) = NULL, @bordxNo nvarchar(20), @receiptNo nvarchar(30), @receiptDate datetime, @chequeNo nvarchar(30) = NULL,
    @amount money, @remarks nvarchar(500) = NULL, @mth nvarchar(2), @yr nvarchar(4), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptFrFFS WHERE CAST(ReceiptIndex AS nvarchar(20)) = @index)
        UPDATE CISPayment.dbo.ReceiptFrFFS SET BordxRefNo = @bordxNo, ReceiptNo = @receiptNo, ReceiptDate = @receiptDate, ReceiptChqNo = COALESCE(@chequeNo, ''),
            ReceiptAmt = @amount, ReceiptRemarks = COALESCE(@remarks, ''), ReceiptAccMth = @mth, ReceiptAccYr = @yr, ReceiptLastUpdDate = CAST(GETDATE() AS date),
            ReceiptLastUpdUser = @user, ReceiptStatus = 'F'
        WHERE CAST(ReceiptIndex AS nvarchar(20)) = @index;
    ELSE
        INSERT CISPayment.dbo.ReceiptFrFFS (BordxRefNo, ReceiptNo, ReceiptDate, ReceiptChqNo, ReceiptAmt, ReceiptRemarks, ReceiptAccMth, ReceiptAccYr, ReceiptLastUpdDate,
            ReceiptLastUpdUser, ReceiptStatus)
        VALUES (@bordxNo, @receiptNo, @receiptDate, COALESCE(@chequeNo, ''), @amount, COALESCE(@remarks, ''), @mth, @yr, CAST(GETDATE() AS date), @user, 'F');
END
GO
