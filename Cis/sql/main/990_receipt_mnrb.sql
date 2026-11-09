-- Receipt > Fr MNRB (FrmReceiptMNRB): the bordereaux (VIEW_CISBORReceiptMNRB) and worksheet (VIEW_CISWKSReceiptMNRB) lists with the pasted criteria replaced by
-- parameters, UpdateRecord / UpdateRecord2 / UpdateBordxStatus and the edits.
SET QUOTED_IDENTIFIER ON;
GO

-- PayorComboListing: HISDB PAY with HLTCode S.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_Payors AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM cis.Hdb_PAY WHERE HLTCode = 'S' ORDER BY PAYCompanyName;
END
GO

-- @worksheet 0 bordereaux / 1 worksheet; @scan: one bordx no. (bordereaux) or one worksheet no. (worksheet); @bordxStatus 0-5, @wsStatus 0-2 as the combos.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_Search @worksheet bit, @scan bit = 0, @payor nvarchar(10) = NULL, @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL,
    @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL, @wsStatus int = 0, @bordxStatus int = 0, @recFrom nvarchar(30) = NULL, @recTo nvarchar(30) = NULL,
    @recDateFrom datetime = NULL, @recDateTo datetime = NULL, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @yrFrom nvarchar(4) = NULL,
    @yrTo nvarchar(4) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF @worksheet = 0
        SELECT BordxRefNo, BordxClaimAmt, BordxPAYCode, BordxMNRBFRStatus, SUM(CheckAmt) AS CheckAmt, RECAccPeriodMth, RECAccPeriodYR
        FROM CISWorksheet.dbo.VIEW_CISBORReceiptMNRB
        WHERE (@payor IS NULL OR BordxPAYCode = @payor)
          AND (@bordxFrom IS NULL OR (@scan = 1 AND BordxRefNo = @bordxFrom) OR (@scan = 0 AND BordxRefNo >= @bordxFrom))
          AND (@bordxTo IS NULL OR @scan = 1 OR BordxRefNo <= @bordxTo)
          AND (@bordxStatus <> 1 OR BordxMNRBFRStatus = 'F') AND (@bordxStatus <> 2 OR BordxMNRBFRStatus = '' OR BordxMNRBFRStatus IS NULL)
          AND (@bordxStatus <> 3 OR BordxMNRBFRStatus = 'P') AND (@bordxStatus <> 4 OR BordxMNRBFRStatus IN ('P', 'F'))
          AND (@bordxStatus <> 5 OR BordxMNRBFRStatus = '' OR BordxMNRBFRStatus IS NULL OR BordxMNRBFRStatus = 'P')
          AND (@recFrom IS NULL OR ReceiptNo >= @recFrom) AND (@recTo IS NULL OR ReceiptNo <= @recTo)
          AND (@recDateFrom IS NULL OR ReceiptDate >= @recDateFrom) AND (@recDateTo IS NULL OR ReceiptDate <= @recDateTo)
          AND (@mthFrom IS NULL OR RECAccPeriodMth >= @mthFrom) AND (@mthTo IS NULL OR RECAccPeriodMth <= @mthTo)
          AND (@yrFrom IS NULL OR RECAccPeriodYR >= @yrFrom) AND (@yrTo IS NULL OR RECAccPeriodYR <= @yrTo)
        GROUP BY BordxRefNo, BordxClaimAmt, BordxPAYCode, BordxMNRBFRStatus, RECAccPeriodMth, RECAccPeriodYR
        OPTION (RECOMPILE);
    ELSE
        SELECT CASNumber, BordxRefNo, CLMTotalApproved, ReceiptDate, ReceiptNo, RECAccPeriodMth, RECAccPeriodYR, BordxPAYCode, PVRecMNRBStatus, CheckAmt
        FROM CISWorksheet.dbo.VIEW_CISWKSReceiptMNRB
        WHERE (@payor IS NULL OR BordxPAYCode = @payor)
          AND (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
          AND (@caseFrom IS NULL OR (@scan = 1 AND CASNumber = @caseFrom) OR (@scan = 0 AND CASNumber >= @caseFrom))
          AND (@caseTo IS NULL OR @scan = 1 OR CASNumber <= @caseTo)
          AND (@wsStatus <> 2 OR PVRecMNRBStatus = '' OR PVRecMNRBStatus IS NULL) AND (@wsStatus <> 1 OR PVRecMNRBStatus = 'F')
          AND (@recFrom IS NULL OR ReceiptNo >= @recFrom) AND (@recTo IS NULL OR ReceiptNo <= @recTo)
          AND (@recDateFrom IS NULL OR ReceiptDate >= @recDateFrom) AND (@recDateTo IS NULL OR ReceiptDate <= @recDateTo)
          AND (@mthFrom IS NULL OR RECAccPeriodMth >= @mthFrom) AND (@mthTo IS NULL OR RECAccPeriodMth <= @mthTo)
          AND (@yrFrom IS NULL OR RECAccPeriodYR >= @yrFrom) AND (@yrTo IS NULL OR RECAccPeriodYR <= @yrTo)
        OPTION (RECOMPILE);
END
GO

-- lstBordx_ItemCheck: the bordereaux's MNRB status and the claims it takes (partial: the claims not yet received).
CREATE OR ALTER PROCEDURE cis.RcpMnrb_BordxCases @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @status nvarchar(5) = (SELECT TOP 1 RTRIM(BordxMNRBFRStatus) FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @bordxNo);
    SELECT @status AS BordxMNRBFRStatus;
    SELECT RTRIM(CASNumber) AS CASNumber FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo AND (@status <> 'P' OR @status IS NULL OR PVRecMNRBStatus IS NULL);
END
GO

-- ProcessCheckedItem: the claim's MNRB status.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_CaseStatus @casNumber nvarchar(30), @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 RTRIM(PVRecMNRBStatus) AS PVRecMNRBStatus FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo;
END
GO

-- UpdateRecord (bordereaux save, one claim): a ReceiptfrMNRB row, the claim received and the bordereaux received (F).
CREATE OR ALTER PROCEDURE cis.RcpMnrb_SaveBordxCase @casNumber nvarchar(30), @bordxNo nvarchar(20), @receiptNo nvarchar(30), @receiptDate datetime, @receiptAmt money,
    @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo) RETURN;
    DECLARE @amt money = (SELECT TOP 1 CLMTotalApproved FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo);
    INSERT CISPayment.dbo.ReceiptfrMNRB (ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, BordxNo, CASNumber, Remark, MNRBLastUpdatedate, MNRBLastUpdateUser, RECAccPeriodMth,
        RECAccPeriodYR, RecMNRBAmt, CheckAmt)
    VALUES (@receiptNo, @receiptDate, @receiptAmt, @chequeNo, @bordxNo, @casNumber, COALESCE(@remark, ''), CAST(GETDATE() AS date), @user, @mth, @yr, @amt, @amt);
    UPDATE CISWorksheet.dbo.Claim SET PVRecMNRBStatus = 'F' WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo;
    UPDATE CISWorksheet.dbo.Bordx SET BordxMNRBRecDate = @receiptDate, BordxMNRBFRStatus = 'F' WHERE BordxRefNo = @bordxNo;
END
GO

-- UpdateRecord2 (worksheet save): a ReceiptfrMNRB row when the claim has none for the bordereaux, the claim's status, then the bordereaux F when all its claims are
-- received, else P.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_SaveWksCase @casNumber nvarchar(30), @bordxNo nvarchar(20), @status nvarchar(2), @receiptNo nvarchar(30), @receiptDate datetime,
    @selectedAmt money, @receiptAmt money, @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @amt money = (SELECT TOP 1 CLMTotalApproved FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo);
    IF NOT EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptfrMNRB WHERE CASNumber = @casNumber AND BordxNo = @bordxNo)
        INSERT CISPayment.dbo.ReceiptfrMNRB (CASNumber, ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, BordxNo, MNRBLastUpdatedate, MNRBLastUpdateUser, RecMNRBAmt,
            CheckAmt, RECAccPeriodMth, RECAccPeriodYR)
        VALUES (@casNumber, @receiptNo, @receiptDate, @selectedAmt, @chequeNo, COALESCE(@remark, ''), @bordxNo, CAST(GETDATE() AS date), @user, @receiptAmt,
            COALESCE(@amt, 0), @mth, @yr);
    UPDATE CISWorksheet.dbo.Claim SET PVRecMNRBStatus = @status WHERE BordxRefNo = @bordxNo AND CASNumber = @casNumber;
    DECLARE @all int = (SELECT COUNT(*) FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo),
        @full int = (SELECT COUNT(*) FROM CISWorksheet.dbo.Claim WHERE PVRecMNRBStatus = 'F' AND BordxRefNo = @bordxNo);
    UPDATE CISWorksheet.dbo.Bordx SET BordxMNRBRecDate = @receiptDate, BordxMNRBFRStatus = CASE WHEN @all = @full THEN 'F' ELSE 'P' END WHERE BordxRefNo = @bordxNo;
END
GO

-- lstBordx_Click / lstWorksheet_Click in edit mode.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_Receipt @bordxNo nvarchar(20), @casNumber nvarchar(30) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, RECAccPeriodMth, RECAccPeriodYR
    FROM CISPayment.dbo.ReceiptfrMNRB WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber);
END
GO

-- UpdateBordx (@casNumber NULL: the bordereaux's receipt rows, with the period) / UpdateWorksheet (the claim's row); a row is added when there is none.
CREATE OR ALTER PROCEDURE cis.RcpMnrb_Edit @bordxNo nvarchar(20), @casNumber nvarchar(30) = NULL, @receiptNo nvarchar(30), @receiptDate datetime, @receiptAmt money,
    @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptfrMNRB WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber))
        INSERT CISPayment.dbo.ReceiptfrMNRB (ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, BordxNo, CASNumber, MNRBLastUpdatedate, MNRBLastUpdateUser,
            RECAccPeriodMth, RECAccPeriodYR)
        VALUES (@receiptNo, @receiptDate, @receiptAmt, @chequeNo, COALESCE(@remark, ''), @bordxNo, @casNumber, CAST(GETDATE() AS date), @user,
            CASE WHEN @casNumber IS NULL THEN @mth END, CASE WHEN @casNumber IS NULL THEN @yr END);
    ELSE
        UPDATE CISPayment.dbo.ReceiptfrMNRB SET ReceiptNo = @receiptNo, ReceiptDate = @receiptDate, ReceiptAmt = @receiptAmt, ChequeNo = @chequeNo,
            Remark = COALESCE(@remark, ''), MNRBLastUpdatedate = CAST(GETDATE() AS date), MNRBLastUpdateUser = @user,
            RECAccPeriodMth = CASE WHEN @casNumber IS NULL AND @mth IS NOT NULL THEN @mth ELSE RECAccPeriodMth END,
            RECAccPeriodYR = CASE WHEN @casNumber IS NULL AND @yr IS NOT NULL THEN @yr ELSE RECAccPeriodYR END
        WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber);
END
GO
