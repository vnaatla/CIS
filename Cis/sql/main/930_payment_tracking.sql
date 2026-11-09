-- Payment > Tracking / Tracking (Jupiter) (FrmPayment / frmPayment2): CISPayment PaymentReceiptModule cases 4-9 (Pay to Clinic, VIEW_PaymentTracking) and
-- 10-15 (Pay to Member, VIEW_PaymentMBMTracking), with the pasted criteria replaced by parameters.
SET QUOTED_IDENTIFIER ON;
GO

-- @member 0 clinic / 1 member; @by: general, bordx, claim, clinic, group, pv; @payment 1 paid (F), 2 partial (P).
CREATE OR ALTER PROCEDURE cis.PayTracking_Search @member bit, @by varchar(10), @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL,
    @company nvarchar(200) = NULL, @mbmNo nvarchar(30) = NULL, @invFrom nvarchar(50) = NULL, @invTo nvarchar(50) = NULL, @bordxFrom nvarchar(20) = NULL,
    @bordxTo nvarchar(20) = NULL, @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @clinic nvarchar(4) = NULL, @invDateFrom datetime = NULL,
    @invDateTo datetime = NULL, @payment int = 0, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL, @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL,
    @amtFrom money = NULL, @amtTo money = NULL, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @view nvarchar(50) = CASE WHEN @member = 1 THEN N'VIEW_PaymentMBMTracking' ELSE N'VIEW_PaymentTracking' END,
        @status nvarchar(30) = CASE WHEN @member = 1 THEN N'PVPayMBMStatus' ELSE N'PVPayCLNStatus' END,
        @pv nvarchar(30) = CASE WHEN @member = 1 THEN N'PVNo' ELSE N'PVCLNNo' END,
        @pvDate nvarchar(30) = CASE WHEN @member = 1 THEN N'PVDate' ELSE N'PVCLNDate' END,
        @payable nvarchar(30) = CASE WHEN @member = 1 THEN N'CLMPayableByMedix2M' ELSE N'CLMPayableByMedix2C' END,
        @mth nvarchar(30) = CASE WHEN @member = 1 THEN N'PAYAccPeriodMth' ELSE N'CLNAccPeriodMth' END,
        @yr nvarchar(30) = CASE WHEN @member = 1 THEN N'PAYAccPeriodYR' ELSE N'CLNAccPeriodYR' END,
        @paid nvarchar(30) = CASE WHEN @member = 1 THEN N'MBMPaidAmt' ELSE N'CLNPaidAmt' END;
    DECLARE @filter nvarchar(max) = N' WHERE 0 = 0'
        + CASE WHEN @caseFrom IS NOT NULL THEN N' And CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN @caseTo IS NOT NULL THEN N' And CASNumber <= @caseTo' ELSE N'' END
        + CASE WHEN @company IS NOT NULL THEN N' And GroupCompany = @company' ELSE N'' END
        + CASE WHEN @mbmNo IS NOT NULL THEN N' And MBMNumber = @mbmNo' ELSE N'' END
        + CASE WHEN @invFrom IS NOT NULL THEN N' AND InvoiceNo >= @invFrom' ELSE N'' END
        + CASE WHEN @invTo IS NOT NULL THEN N' AND InvoiceNo <= @invTo' ELSE N'' END
        + CASE WHEN @bordxFrom IS NOT NULL THEN N' And BordxRefNo >= @bordxFrom' ELSE N'' END
        + CASE WHEN @bordxTo IS NOT NULL THEN N' And BordxRefNo <= @bordxTo' ELSE N'' END
        + CASE WHEN @opFrom IS NOT NULL THEN N' And OPBatchNo >= @opFrom' ELSE N'' END
        + CASE WHEN @opTo IS NOT NULL THEN N' And OPBatchNo <= @opTo' ELSE N'' END
        + CASE WHEN @clinic IS NOT NULL THEN N' AND CLinicCode = @clinic' ELSE N'' END
        + CASE WHEN @invDateFrom IS NOT NULL THEN N' And InvoiceDate >= @invDateFrom' ELSE N'' END
        + CASE WHEN @invDateTo IS NOT NULL THEN N' And InvoiceDate <= @invDateTo' ELSE N'' END
        + CASE @payment WHEN 1 THEN N' and ' + @status + N' = ''F''' WHEN 2 THEN N' and ' + @status + N' = ''P''' ELSE N'' END
        + CASE WHEN @pvFrom IS NOT NULL THEN N' AND ' + @pv + N' >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' AND ' + @pv + N' <= @pvTo' ELSE N'' END
        + CASE WHEN @pvDateFrom IS NOT NULL THEN N' And ' + @pvDate + N' >= @pvDateFrom' ELSE N'' END
        + CASE WHEN @pvDateTo IS NOT NULL THEN N' And ' + @pvDate + N' <= @pvDateTo' ELSE N'' END
        + CASE WHEN @amtFrom IS NOT NULL THEN N' AND ' + @payable + N' >= @amtFrom' ELSE N'' END
        + CASE WHEN @amtTo IS NOT NULL THEN N' AND ' + @payable + N' <= @amtTo' ELSE N'' END
        + CASE WHEN @mthFrom IS NOT NULL THEN N' And ' + @mth + N' >= @mthFrom' ELSE N'' END
        + CASE WHEN @mthTo IS NOT NULL THEN N' And ' + @mth + N' <= @mthTo' ELSE N'' END
        + CASE WHEN @yrFrom IS NOT NULL THEN N' And ' + @yr + N' >= @yrFrom' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL THEN N' And ' + @yr + N' <= @yrTo' ELSE N'' END;
    DECLARE @agg nvarchar(200) = N', COUNT(CASNumber) AS NoofCase, SUM(' + @paid + N') AS ' + @paid;
    DECLARE @sql nvarchar(max) = CASE @by
        WHEN 'bordx' THEN N'Select BordxRefNo, BordxDate' + @agg + N' From ' + @view + @filter + N' GROUP BY BordxRefNo, BordxDate'
        WHEN 'claim' THEN N'Select CASNumber, ' + @paid + N' From ' + @view + @filter
        WHEN 'clinic' THEN N'Select ClinicCode, CLNName, ' + @pvDate + @agg + N' From ' + @view + @filter + N' GROUP BY ClinicCode, CLNName, ' + @pvDate
        WHEN 'group' THEN N'Select GroupCompany, ' + @pvDate + @agg + N' From ' + @view + @filter + N' GROUP BY GroupCompany, ' + @pvDate
        WHEN 'pv' THEN N'Select ' + @pv + N', ' + @pvDate + @agg + N' From ' + @view + @filter + N' GROUP BY ' + @pv + N', ' + @pvDate
        ELSE N'Select * From ' + @view + @filter END;
    EXEC CISPayment.sys.sp_executesql @sql,
        N'@caseFrom nvarchar(30), @caseTo nvarchar(30), @company nvarchar(200), @mbmNo nvarchar(30), @invFrom nvarchar(50), @invTo nvarchar(50), @bordxFrom nvarchar(20),
          @bordxTo nvarchar(20), @opFrom nvarchar(30), @opTo nvarchar(30), @clinic nvarchar(4), @invDateFrom datetime, @invDateTo datetime, @pvFrom nvarchar(30),
          @pvTo nvarchar(30), @pvDateFrom datetime, @pvDateTo datetime, @amtFrom money, @amtTo money, @mthFrom nvarchar(2), @mthTo nvarchar(2), @yrFrom nvarchar(4),
          @yrTo nvarchar(4)',
        @caseFrom, @caseTo, @company, @mbmNo, @invFrom, @invTo, @bordxFrom, @bordxTo, @opFrom, @opTo, @clinic, @invDateFrom, @invDateTo, @pvFrom, @pvTo, @pvDateFrom,
        @pvDateTo, @amtFrom, @amtTo, @mthFrom, @mthTo, @yrFrom, @yrTo;
END
GO
