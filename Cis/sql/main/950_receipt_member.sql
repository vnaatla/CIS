-- Receipt > Fr Member (FrmReceiptMember): CISPayment PaymentReceiptModule case 1 (VIEW_CISReceiptFrMBM, member debit notes) with the pasted criteria
-- replaced by parameters, UpdateRecord (the receipt of the ticked cases) and cmdUpdate (the edited receipt).
SET QUOTED_IDENTIFIER ON;
GO

-- @payment 0-5 as CboPayment; @blank*: the "empty" boxes. The receipt no. range uses the "from" box for both ends, as the desktop.
CREATE OR ALTER PROCEDURE cis.RcpMbm_Search @scan bit = 0, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL, @mbmNo nvarchar(30) = NULL,
    @mbmName nvarchar(200) = NULL, @payment int = 0, @receiptFrom nvarchar(30) = NULL, @blankReceipt bit = 0, @recDateFrom datetime = NULL, @recDateTo datetime = NULL,
    @drFrom nvarchar(30) = NULL, @drTo nvarchar(30) = NULL, @blankDr bit = 0, @amtFrom money = NULL, @amtTo money = NULL, @blankAmt bit = 0,
    @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @blankMth bit = 0, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL, @blankYr bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT CASNumber, MBMNumber, MBMName, DebitCreditRefNo, ReceiptAmt, ReceiptNo, ReceiptDate, CLMPayableByMedix2M, PVRecMBMStatus, PAYCode, ReceiptIndex,
        RECMBMAccPeriodMth, RECMBMAccPeriodYr,
        (SELECT SUM(r.ReceiptAmt) FROM CISPayment.dbo.ReceiptfrMBM r WHERE r.CASNumber = v.CASNumber) AS ReceivedTotal
    FROM CISPayment.dbo.VIEW_CISReceiptFrMBM v
    WHERE CLMPayableByMedix2M <= 0 AND SUBSTRING(DebitCreditRefNo, 1, 1) = 'D'
      AND (@scan = 0 OR CASNumber = @caseFrom)
      AND (@scan = 1 OR @caseFrom IS NULL OR CASNumber >= @caseFrom) AND (@scan = 1 OR @caseTo IS NULL OR CASNumber <= @caseTo)
      AND (@mbmNo IS NULL OR MBMNumber LIKE '%' + @mbmNo + '%') AND (@mbmName IS NULL OR MBMName LIKE '%' + @mbmName + '%')
      AND (@payment <> 1 OR PVRecMBMStatus = 'F') AND (@payment <> 2 OR PVRecMBMStatus = '' OR PVRecMBMStatus IS NULL)
      AND (@payment <> 3 OR PVRecMBMStatus = 'P') AND (@payment <> 4 OR PVRecMBMStatus IN ('P', 'F'))
      AND (@payment <> 5 OR PVRecMBMStatus = '' OR PVRecMBMStatus IS NULL OR PVRecMBMStatus = 'P')
      AND (@receiptFrom IS NULL OR (ReceiptNo >= @receiptFrom AND ReceiptNo <= @receiptFrom))
      AND (@blankReceipt = 0 OR ReceiptNo IS NULL OR ReceiptNo = '')
      AND (@recDateFrom IS NULL OR ReceiptDate >= @recDateFrom) AND (@recDateTo IS NULL OR ReceiptDate <= @recDateTo)
      AND (@drFrom IS NULL OR DebitCreditRefNo >= @drFrom) AND (@drTo IS NULL OR DebitCreditRefNo <= @drTo)
      AND (@blankDr = 0 OR DebitCreditRefNo IS NULL OR DebitCreditRefNo = '')
      AND (@amtFrom IS NULL OR CLMPayableByMedix2M >= @amtFrom) AND (@amtTo IS NULL OR CLMPayableByMedix2M <= @amtTo)
      AND (@blankAmt = 0 OR CLMPayableByMedix2M IS NULL)
      AND (@mthFrom IS NULL OR RECMBMAccPeriodMth >= @mthFrom) AND (@mthTo IS NULL OR RECMBMAccPeriodMth <= @mthTo) AND (@blankMth = 0 OR RECMBMAccPeriodMth IS NULL)
      AND (@yrFrom IS NULL OR RECMBMAccPeriodYr >= @yrFrom) AND (@yrTo IS NULL OR RECMBMAccPeriodYr <= @yrTo) AND (@blankYr = 0 OR RECMBMAccPeriodYr IS NULL)
    OPTION (RECOMPILE);
END
GO

-- UpdateRecord: the case's first ReceiptfrMBM row (a new one when none) gets the receipt details, the status and the amount added (Full: replaced when equal).
CREATE OR ALTER PROCEDURE cis.RcpMbm_Save @casNumber nvarchar(30), @receiptNo nvarchar(30), @receiptDate datetime, @chequeNo nvarchar(30), @remarks nvarchar(500) = NULL,
    @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @status nvarchar(10), @amount money, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @index int = (SELECT TOP 1 ReceiptIndex FROM CISPayment.dbo.ReceiptfrMBM WITH (UPDLOCK) WHERE CASNumber = @casNumber);
    IF @index IS NULL
    BEGIN
        INSERT CISPayment.dbo.ReceiptfrMBM (CASNumber, ReceiptAmt) VALUES (@casNumber, NULL);
        SET @index = SCOPE_IDENTITY();
    END
    UPDATE CISPayment.dbo.ReceiptfrMBM SET ReceiptNo = @receiptNo, ReceiptDate = @receiptDate, ReceiptChequeNo = @chequeNo, ReceiptLastUpdateDate = CAST(GETDATE() AS date),
        ReceiptLastUpdateUser = @user, CASNumber = @casNumber, ReceiptRemarks = COALESCE(@remarks, ''),
        RECMBMAccPeriodMth = COALESCE(@mth, RECMBMAccPeriodMth), RECMBMAccPeriodYr = COALESCE(@yr, RECMBMAccPeriodYr), PVRecMBMStatus = LEFT(@status, 1),
        ReceiptAmt = CASE WHEN @status = 'Full' THEN CASE WHEN ReceiptAmt IS NULL OR ReceiptAmt = @amount THEN @amount ELSE @amount + ReceiptAmt END
            ELSE @amount + COALESCE(ReceiptAmt, 0) END
    WHERE ReceiptIndex = @index;
END
GO

-- cmdUpdate_Click: the selected receipt row (a new row when there is none).
CREATE OR ALTER PROCEDURE cis.RcpMbm_Edit @index nvarchar(20) = NULL, @receiptNo nvarchar(30), @receiptDate datetime, @amount money, @chequeNo nvarchar(30),
    @remarks nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptfrMBM WHERE CAST(ReceiptIndex AS nvarchar(20)) = @index)
        UPDATE CISPayment.dbo.ReceiptfrMBM SET ReceiptNo = @receiptNo, ReceiptDate = @receiptDate, ReceiptAmt = @amount, ReceiptChequeNo = @chequeNo,
            ReceiptRemarks = COALESCE(@remarks, ''), ReceiptLastUpdateDate = CAST(GETDATE() AS date), ReceiptLastUpdateUser = @user,
            RECMBMAccPeriodMth = COALESCE(@mth, RECMBMAccPeriodMth), RECMBMAccPeriodYr = COALESCE(@yr, RECMBMAccPeriodYr)
        WHERE CAST(ReceiptIndex AS nvarchar(20)) = @index;
    ELSE
        INSERT CISPayment.dbo.ReceiptfrMBM (ReceiptNo, ReceiptDate, ReceiptAmt, ReceiptChequeNo, ReceiptRemarks, ReceiptLastUpdateDate, ReceiptLastUpdateUser,
            RECMBMAccPeriodMth, RECMBMAccPeriodYr)
        VALUES (@receiptNo, @receiptDate, @amount, @chequeNo, COALESCE(@remarks, ''), CAST(GETDATE() AS date), @user, @mth, @yr);
END
GO

-- lstReceiptMBM_Click (edit mode): the receipt row.
CREATE OR ALTER PROCEDURE cis.RcpMbm_Receipt @index nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 ReceiptIndex, ReceiptNo, ReceiptDate, ReceiptAmt, ReceiptChequeNo, ReceiptRemarks, RECMBMAccPeriodMth, RECMBMAccPeriodYr
    FROM CISPayment.dbo.ReceiptFrMBM WHERE CAST(ReceiptIndex AS nvarchar(20)) = @index;
END
GO
