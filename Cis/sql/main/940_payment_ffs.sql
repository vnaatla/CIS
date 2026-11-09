-- Payment > FFS Payment (FrmPaymentFFS): CISPayment SP_CISPayment 'PayFFS' (View_PaymentFFS) with the pasted criteria replaced by parameters, and UpdateRecord.
SET QUOTED_IDENTIFIER ON;
GO

-- @pvNo: the PV just saved / updated; @scan: one bordx no.; @payment 1 paid, 2 unpaid; @blank*: the "empty" boxes; @bordxDate: the preview row's bordx date.
CREATE OR ALTER PROCEDURE cis.PayFfs_Search @pvNo nvarchar(30) = NULL, @company nvarchar(200) = NULL, @payment int = 0, @scan bit = 0, @bordxFrom nvarchar(20) = NULL,
    @bordxTo nvarchar(20) = NULL, @blankBordxDate bit = 0, @bordxDateFrom datetime = NULL, @bordxDateTo datetime = NULL, @blankPv bit = 0, @pvFrom nvarchar(30) = NULL,
    @pvTo nvarchar(30) = NULL, @blankPvDate bit = 0, @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL, @blankMth bit = 0, @mthFrom nvarchar(2) = NULL,
    @mthTo nvarchar(2) = NULL, @blankYr bit = 0, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 100000 GroupCompany, BordxRefNo, CLMMCOProcFee, PaymentIndex, PVNo, PVDate, PVChqNo, PVChqDate, PVChqAmt, PVBank, PVAcc, Remark, FFSPaidAmt, FFSAccMth,
        FFSAccYr, PVPayStatus, GRPCoCode, BordxDate, GRPCoName, GRPCoAdd1, GRPCoAdd2, GRPCoAdd3, PVRecNo, PVRecDate, CLMBordxAmt
    FROM CISPayment.dbo.View_PaymentFFS
    WHERE (@pvNo IS NULL OR PVNo = @pvNo)
      AND (@pvNo IS NOT NULL OR (
            (@company IS NULL OR GroupCompany = @company)
        AND (@payment <> 1 OR PVPayStatus = 'F') AND (@payment <> 2 OR PVPayStatus = '' OR PVPayStatus IS NULL)
        AND (@scan = 0 OR BordxRefNo = @bordxFrom)
        AND (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
        AND (@blankBordxDate = 0 OR BordxDate IS NULL OR BordxDate = '')
        AND (@bordxDateFrom IS NULL OR BordxDate >= @bordxDateFrom) AND (@bordxDateTo IS NULL OR BordxDate <= @bordxDateTo)
        AND (@blankPv = 0 OR PVNo IS NULL OR PVNo = '')
        AND (@pvFrom IS NULL OR PVNo >= @pvFrom) AND (@pvTo IS NULL OR PVNo <= @pvTo)
        AND (@blankPvDate = 0 OR PVDate IS NULL OR PVDate = '')
        AND (@pvDateFrom IS NULL OR PVDate >= @pvDateFrom) AND (@pvDateTo IS NULL OR PVDate <= @pvDateTo)
        AND (@blankMth = 0 OR FFSAccMth IS NULL) AND (@mthFrom IS NULL OR FFSAccMth >= @mthFrom) AND (@mthTo IS NULL OR FFSAccMth <= @mthTo)
        AND (@blankYr = 0 OR FFSAccYr IS NULL) AND (@yrFrom IS NULL OR FFSAccYr >= @yrFrom) AND (@yrTo IS NULL OR FFSAccYr <= @yrTo)))
    OPTION (RECOMPILE);
END
GO

-- UpdateRecord: the PaymentFFS row of @index is overwritten, or a new row is added; GRPCoCode is the payee's code.
CREATE OR ALTER PROCEDURE cis.PayFfs_Save @index nvarchar(20) = NULL, @bordxNo nvarchar(20), @pvNo nvarchar(30), @pvDate datetime, @chequeDate datetime = NULL,
    @chequeNo nvarchar(30) = NULL, @chequeAmt money, @bank nvarchar(10), @account nvarchar(50) = NULL, @remark nvarchar(500) = NULL, @paidAmt money, @mth nvarchar(2),
    @yr nvarchar(4), @payee nvarchar(200), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @grp nvarchar(20) = (SELECT TOP 1 RTRIM(GRPCoCode) FROM dbo.GRPCompany WHERE GRPCoName = @payee);
    IF @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.PaymentFFS WHERE CAST(PaymentIndex AS nvarchar(20)) = @index)
        UPDATE CISPayment.dbo.PaymentFFS SET BordxRefNo = @bordxNo, PVNo = @pvNo, PVDate = @pvDate, PVChqDate = @chequeDate, PVChqNo = COALESCE(@chequeNo, ''),
            PVChqAmt = @chequeAmt, PVBank = @bank, PVAcc = COALESCE(@account, ''), Remark = COALESCE(@remark, ''), FFSPaidAmt = @paidAmt, FFSAccMth = @mth, FFSAccYr = @yr,
            PVLastUpdateDate = CAST(GETDATE() AS date), PVLastUpdateUser = @user, PVPayStatus = 'F', GRPCoCode = COALESCE(@grp, '')
        WHERE CAST(PaymentIndex AS nvarchar(20)) = @index;
    ELSE
        INSERT CISPayment.dbo.PaymentFFS (BordxRefNo, PVNo, PVDate, PVChqDate, PVChqNo, PVChqAmt, PVBank, PVAcc, Remark, FFSPaidAmt, FFSAccMth, FFSAccYr, PVLastUpdateDate,
            PVLastUpdateUser, PVPayStatus, GRPCoCode)
        VALUES (@bordxNo, @pvNo, @pvDate, @chequeDate, COALESCE(@chequeNo, ''), @chequeAmt, @bank, COALESCE(@account, ''), COALESCE(@remark, ''), @paidAmt, @mth, @yr,
            CAST(GETDATE() AS date), @user, 'F', COALESCE(@grp, ''));
END
GO
