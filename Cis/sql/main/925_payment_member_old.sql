-- Payment > To Member (old) (frmPaymentMember): View_Payment2MBM with the pasted criteria replaced by parameters, and UpdateRecord.
SET QUOTED_IDENTIFIER ON;
GO

-- @payment 0-5 as CboPayment; @bordx 0 / 1 (NULL: not chosen); @blank*: the "empty" boxes; @pvNo: the PV just saved. The PV no. boxes are each an "equals"
-- filter, as the desktop.
CREATE OR ALTER PROCEDURE cis.PayMbmOld_Search @pvNo nvarchar(30) = NULL, @scan bit = 0, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL,
    @mbmNo nvarchar(30) = NULL, @mbmName nvarchar(200) = NULL, @company nvarchar(200) = NULL, @payment int = 0, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL,
    @blankBatch bit = 0, @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @blankPv bit = 0, @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL,
    @blankPvDate bit = 0, @chqFrom nvarchar(50) = NULL, @chqTo nvarchar(50) = NULL, @blankChq bit = 0, @crFrom nvarchar(30) = NULL, @crTo nvarchar(30) = NULL,
    @blankCr bit = 0, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @blankMth bit = 0, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL,
    @blankYr bit = 0, @bordx int = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM CISPayment.dbo.View_Payment2MBM
    WHERE (@pvNo IS NULL OR PVNo = @pvNo)
      AND (@pvNo IS NOT NULL OR (
            (@scan = 0 OR CASNumber = @caseFrom)
        AND (@scan = 1 OR @caseFrom IS NULL OR CASNumber >= @caseFrom) AND (@scan = 1 OR @caseTo IS NULL OR CASNumber <= @caseTo)
        AND (@mbmNo IS NULL OR MBMnumber = @mbmNo) AND (@mbmName IS NULL OR MBMName LIKE '%' + @mbmName + '%') AND (@company IS NULL OR GroupCompany = @company)
        AND (@payment <> 1 OR PVPayMBMStatus = 'F') AND (@payment <> 2 OR PVPayMBMStatus = '' OR PVPayMBMStatus IS NULL)
        AND (@payment <> 3 OR PVPayMBMStatus = 'P') AND (@payment <> 4 OR PVPayMBMStatus IN ('P', 'F'))
        AND (@payment <> 5 OR PVPayMBMStatus = '' OR PVPayMBMStatus IS NULL OR PVPayMBMStatus = 'P')
        AND (@pvFrom IS NULL OR PVNo = @pvFrom) AND (@pvTo IS NULL OR PVNo = @pvTo)
        AND (@blankBatch = 0 OR OPBatchNo IS NULL OR OPBatchNo = '') AND (@opFrom IS NULL OR OPBatchNo >= @opFrom) AND (@opTo IS NULL OR OPBatchNo <= @opTo)
        AND (@blankPv = 0 OR PVNo IS NULL OR PVNo = '')
        AND (@pvDateFrom IS NULL OR PVDate >= @pvDateFrom) AND (@pvDateTo IS NULL OR PVDate <= @pvDateTo) AND (@blankPvDate = 0 OR PVDate IS NULL OR PVDate = '')
        AND (@chqFrom IS NULL OR PVChequeNo >= @chqFrom) AND (@chqTo IS NULL OR PVChequeNo <= @chqTo) AND (@blankChq = 0 OR PVChequeNo IS NULL OR PVChequeNo = '')
        AND (@crFrom IS NULL OR DebitCreditRefNo >= @crFrom) AND (@crTo IS NULL OR DebitCreditRefNo <= @crTo)
        AND (@blankCr = 0 OR DebitCreditRefNo IS NULL OR DebitCreditRefNo = '')
        AND (@mthFrom IS NULL OR PAYAccPeriodMth >= @mthFrom) AND (@mthTo IS NULL OR PAYAccPeriodMth <= @mthTo) AND (@blankMth = 0 OR PAYAccPeriodMth IS NULL)
        AND (@yrTo IS NULL OR @yrFrom IS NULL OR PAYAccPeriodYr >= @yrFrom) AND (@yrTo IS NULL OR PAYAccPeriodYr <= @yrTo) AND (@blankYr = 0 OR PAYAccPeriodYr IS NULL)
        AND (@bordx IS NULL OR ClaimBordxStatus = @bordx)))
    OPTION (RECOMPILE);
END
GO

-- UpdateRecord: Save (@edit 0): a partial row first gets the new status, then a new row is added (also when the case has none); Update (@edit 1): the row of
-- @index is overwritten.
CREATE OR ALTER PROCEDURE cis.PayMbmOld_Save @edit bit, @index nvarchar(20) = NULL, @casNumber nvarchar(15), @pvNo nvarchar(15), @pvDate datetime, @chequeDate datetime,
    @chequeNo nvarchar(50), @chequeAmt money, @payTo nvarchar(60), @bank nvarchar(100), @account nvarchar(30) = NULL, @remark nvarchar(200) = NULL, @paidAmt money,
    @user nvarchar(10), @mth nvarchar(2), @yr nvarchar(4), @status char(2), @wordingOne nvarchar(500), @wordingTwo nvarchar(500) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @found bit = CASE WHEN @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.Payment2MBM WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index) THEN 1 ELSE 0 END;
    IF @found = 1 AND @edit = 1
    BEGIN
        UPDATE CISPayment.dbo.Payment2MBM SET CASNumber = @casNumber, PVNo = @pvNo, PVDate = @pvDate, PVChequeDate = @chequeDate, PVChequeNo = @chequeNo,
            PVCLNChequeAmt = @chequeAmt, PVPayto = @payTo, PVPayeeBank = @bank, PVPayeeACNo = COALESCE(@account, ''), Remark = COALESCE(@remark, ''), MBMPaidAmt = @paidAmt,
            Pay2MBMLastUpdateDate = CAST(GETDATE() AS date), Pay2MBMLastUpdateUser = @user, PAYAccPeriodMth = @mth, PAYAccPeriodYr = @yr, PVPayMBMStatus = LEFT(@status, 1),
            PVWordingOne = @wordingOne, PVWordingTwo = @wordingTwo
        WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index;
        RETURN;
    END
    IF @found = 1
        UPDATE CISPayment.dbo.Payment2MBM SET PVPayMBMStatus = LEFT(@status, 1) WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index AND RTRIM(PVPayMBMStatus) = 'P';
    INSERT CISPayment.dbo.Payment2MBM (CASNumber, PVNo, PVDate, PVChequeDate, PVChequeNo, PVCLNChequeAmt, PVPayto, PVPayeeBank, PVPayeeACNo, Remark, MBMPaidAmt,
        Pay2MBMLastUpdateDate, Pay2MBMLastUpdateUser, PAYAccPeriodMth, PAYAccPeriodYr, PVPayMBMStatus, PVWordingOne, PVWordingTwo)
    VALUES (@casNumber, @pvNo, @pvDate, @chequeDate, @chequeNo, @chequeAmt, @payTo, @bank, COALESCE(@account, ''), COALESCE(@remark, ''), @paidAmt,
        CAST(GETDATE() AS date), @user, @mth, @yr, LEFT(@status, 1), @wordingOne, @wordingTwo);
END
GO
