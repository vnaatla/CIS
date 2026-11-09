-- Claim Bordx > Bordx Delivery: Send to Grp. Company (FrmSendtoPayor, ClaimBordxModule case 7), Returned Fr Grp. Company (FrmReturnPayor, case 8 / 9),
-- Send to A/C (FrmSendtoAC, case 10) and Send to MNRB (FrmSendtoMNRB, its own select). @kind: payor, returned, ac, mnrb.
-- The pasted criteria are parameters; amounts / periods are compared as the desktop did (text against the columns).
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxDelivery_Search @kind varchar(10), @company nvarchar(200) = NULL, @payCode nvarchar(10) = NULL, @scan bit = 0,
    @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL, @dateFrom datetime = NULL, @dateTo datetime = NULL, @amountFrom nvarchar(20) = NULL,
    @amountTo nvarchar(20) = NULL, @monthFrom nvarchar(4) = NULL, @monthTo nvarchar(4) = NULL, @yearFrom nvarchar(6) = NULL, @yearTo nvarchar(6) = NULL,
    @status char(1) = NULL, @sentFrom datetime = NULL, @sentTo datetime = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 30000 BordxRefNo, BordxPAYCode, BordxDate, BordxClaimAmt, GroupCompany, BordxPAYSentDate, BordxPAYToStatus, BordxPAYRecDocDate,
        BordxPAYDOCFRStatus, BordxMNRBSentDate, BordxMNRBRecDocDate, BordxMNRBToStatus, BordxSendACDate, BordxSendToACStatus, PAYToACPeriodMth, PAYToACPeriodYR,
        CASE @kind WHEN 'payor' THEN COALESCE(CAST(BordxPAYToStatus AS int), 0) WHEN 'returned' THEN COALESCE(CAST(BordxPAYDOCFRStatus AS int), 0)
            WHEN 'ac' THEN COALESCE(CAST(BordxSendToACStatus AS int), 0) ELSE COALESCE(CAST(BordxMNRBToStatus AS int), 0) END AS Done
    FROM CISWorksheet.dbo.Bordx
    WHERE (@company IS NULL OR @kind = 'mnrb' OR GroupCompany = @company)
      AND (@payCode IS NULL OR @kind <> 'mnrb' OR BordxPAYCode = @payCode)
      AND (@scan = 0 OR BordxRefNo = @bordxFrom)
      AND (@scan = 1 OR @bordxFrom IS NULL OR BordxRefNo >= @bordxFrom)
      AND (@scan = 1 OR @bordxTo IS NULL OR BordxRefNo <= @bordxTo)
      AND (@dateFrom IS NULL OR BordxDate >= @dateFrom) AND (@dateTo IS NULL OR BordxDate <= @dateTo)
      AND (@amountFrom IS NULL OR BordxClaimAmt >= @amountFrom) AND (@amountTo IS NULL OR BordxClaimAmt <= @amountTo)
      AND (@kind = 'ac' OR @monthFrom IS NULL OR PAYToACPeriodMth >= @monthFrom) AND (@kind = 'ac' OR @monthTo IS NULL OR PAYToACPeriodMth <= @monthTo)
      AND (@kind = 'ac' OR @yearFrom IS NULL OR PAYToACPeriodYR >= @yearFrom) AND (@kind = 'ac' OR @yearTo IS NULL OR PAYToACPeriodYR <= @yearTo)
      AND (@kind <> 'ac' OR @sentFrom IS NULL OR BordxSendACDate >= @sentFrom) AND (@kind <> 'ac' OR @sentTo IS NULL OR BordxSendACDate <= @sentTo)
      AND (@kind = 'ac' OR @status IS NULL OR
           (@kind = 'payor' AND BordxPAYToStatus = @status) OR (@kind = 'returned' AND BordxPAYDOCFRStatus = @status) OR (@kind = 'mnrb' AND BordxMNRBToStatus = @status))
    OPTION (RECOMPILE);
END
GO

-- UpdateSentPayor / UpdateReturnPayor / UpdateSentAC / UpdateSentMNRB for one bordereaux. @result 1 updated, 2 returned date earlier than the sent date.
CREATE OR ALTER PROCEDURE cis.BordxDelivery_Update @kind varchar(10), @bordxNo nvarchar(20), @date datetime, @result int OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    SET @result = 1;
    IF @kind = 'payor'
        UPDATE CISWorksheet.dbo.Bordx SET BordxPayToStatus = 1, BordxPAYSentDate = @date WHERE BordxRefNo = @bordxNo;
    ELSE IF @kind = 'returned'
    BEGIN
        IF EXISTS (SELECT 1 FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @bordxNo AND BordxPAYSentDate > @date)
        BEGIN
            SET @result = 2;
            RETURN;
        END
        UPDATE CISWorksheet.dbo.Bordx SET BordxPayDocFRStatus = 1, BordxPAYRecDocDate = @date WHERE BordxRefNo = @bordxNo;
    END
    ELSE IF @kind = 'ac'
        UPDATE CISWorksheet.dbo.Bordx SET BordxSendToACStatus = 1, BordxSendACDate = @date WHERE BordxRefNo = @bordxNo;
    ELSE
        UPDATE CISWorksheet.dbo.Bordx SET BordxMNRBSentDate = @date, BordxMNRBToStatus = 1 WHERE BordxRefNo = @bordxNo;
END
GO

-- The drop-downs: group company names (DISTINCT GRPCoName) and, for MNRB, the payors.
CREATE OR ALTER PROCEDURE cis.BordxDelivery_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT DISTINCT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM dbo.PAY;
END
GO
