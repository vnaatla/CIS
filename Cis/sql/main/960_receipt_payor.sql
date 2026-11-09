-- Receipt > Fr Grp Company (FrmReceiptPayor): CISPayment PaymentReceiptModule cases 18 (by bordereaux) and 19 (by worksheet) copied with the pasted
-- criteria replaced by parameters, UpdateRecord / UpdateRecord2 / UpdateBordxStatus, the OP receipt summary and the edits.
SET QUOTED_IDENTIFIER ON;
GO

-- @worksheet 0 bordereaux list (case 18) / 1 worksheet list (case 19); @scan: one bordx no. (bordereaux) or one worksheet no. (worksheet);
-- @bordxStatus 0-5, @wsStatus 0-2, @insured '' / YES / NO as the combos; @receiptNo: the receipt just saved.
CREATE OR ALTER PROCEDURE cis.RcpPay_Search @worksheet bit, @scan bit = 0, @receiptNo nvarchar(30) = NULL, @company nvarchar(200) = NULL,
    @subsidiary nvarchar(200) = NULL, @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL,
    @wsStatus int = 0, @bordxStatus int = 0, @recFrom nvarchar(30) = NULL, @recTo nvarchar(30) = NULL, @recDateFrom datetime = NULL, @recDateTo datetime = NULL,
    @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL, @payor nvarchar(10) = NULL,
    @client nvarchar(200) = NULL, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL, @mrFrom nvarchar(30) = NULL, @mrTo nvarchar(30) = NULL,
    @insured varchar(3) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @params nvarchar(max) = N'@company nvarchar(200), @subsidiary nvarchar(200), @bordxFrom nvarchar(20), @bordxTo nvarchar(20), @caseFrom nvarchar(30), @caseTo nvarchar(30), @recFrom nvarchar(30), @recTo nvarchar(30), @recDateFrom datetime, @recDateTo datetime, @mthFrom nvarchar(2), @mthTo nvarchar(2), @yrFrom nvarchar(4), @yrTo nvarchar(4), @payor nvarchar(10), @client nvarchar(200), @pvFrom nvarchar(30), @pvTo nvarchar(30), @mrFrom nvarchar(30), @mrTo nvarchar(30), @receiptNo nvarchar(30), @bordxEnd nvarchar(20)';
    DECLARE @bordxEnd nvarchar(20) = CASE WHEN @scan = 1 AND @worksheet = 0 THEN @bordxFrom ELSE @bordxTo END;
    DECLARE @filter nvarchar(max) = CASE WHEN @receiptNo IS NOT NULL THEN N' AND ReceiptfrPAY.ReceiptNo = @receiptNo' ELSE N''
        + CASE WHEN @company IS NOT NULL THEN N' And Bordx.GroupCompany = @company' ELSE N'' END
        + CASE WHEN @worksheet = 1 AND @subsidiary IS NOT NULL THEN N' AND MBMOthers.Subsidiary = @subsidiary' ELSE N'' END
        + CASE WHEN @wsStatus = 2 THEN N' and COALESCE(Claim.PVRecPayorStatus , '''')='''' ' ELSE N'' END
        + CASE WHEN @wsStatus = 1 THEN N' and Claim.PVRecPayorStatus = ''F''' ELSE N'' END
        + CASE WHEN @bordxStatus = 1 THEN N' and COALESCE(ReceiptfrPAY.ReceiptNo, '''') <> ''''' ELSE N'' END
        + CASE WHEN @bordxStatus = 2 THEN N' and COALESCE(ReceiptfrPAY.ReceiptNo, '''') = ''''' ELSE N'' END
        + CASE WHEN @bordxStatus = 3 THEN N' and Bordx.BordxPAYFrStatus = ''P''' ELSE N'' END
        + CASE WHEN @bordxStatus = 4 THEN N' and (Bordx.BordxPAYFrStatus = ''P'' or Bordx.BordxPAYFrStatus = ''F'') ' ELSE N'' END
        + CASE WHEN @bordxStatus = 5 THEN N' and (COALESCE(Bordx.BordxPAYFrStatus , '''')='''' or BordxPAYFrStatus = ''P'') ' ELSE N'' END
        + CASE WHEN @recFrom IS NOT NULL THEN N' AND ReceiptfrPAY.ReceiptNo >= @recFrom' ELSE N'' END
        + CASE WHEN @recTo IS NOT NULL THEN N' AND ReceiptfrPAY.ReceiptNo <= @recTo' ELSE N'' END
        + CASE WHEN @recDateFrom IS NOT NULL THEN N' And ReceiptfrPAY.ReceiptDate >= @recDateFrom' ELSE N'' END
        + CASE WHEN @recDateTo IS NOT NULL THEN N' And ReceiptfrPAY.ReceiptDate <= @recDateTo' ELSE N'' END
        + CASE WHEN @mthFrom IS NOT NULL THEN N' AND ReceiptfrPAY.RECAccPeriodMth >= @mthFrom' ELSE N'' END
        + CASE WHEN @mthTo IS NOT NULL THEN N' AND ReceiptfrPAY.RECAccPeriodMth <= @mthTo' ELSE N'' END
        + CASE WHEN @yrFrom IS NOT NULL THEN N' AND ReceiptfrPAY.RECAccPeriodYr >= @yrFrom' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL THEN N' AND ReceiptfrPAY.RECAccPeriodYr <= @yrTo' ELSE N'' END
        + CASE WHEN @worksheet = 1 AND @scan = 1 AND @caseFrom IS NOT NULL THEN N' AND CISCas.CASNumber = @caseFrom' ELSE N'' END
        + CASE WHEN NOT (@worksheet = 1 AND @scan = 1) AND @caseFrom IS NOT NULL THEN N' AND CISCas.CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN NOT (@worksheet = 1 AND @scan = 1) AND @caseTo IS NOT NULL THEN N' AND CISCas.CASNumber <= @caseTo' ELSE N'' END
        + CASE WHEN @worksheet = 0 AND @scan = 1 AND @bordxFrom IS NOT NULL THEN N' and Bordx.BordxRefNo = @bordxFrom' ELSE N'' END
        + CASE WHEN NOT (@worksheet = 0 AND @scan = 1) AND @bordxFrom IS NOT NULL THEN N' and Bordx.BordxRefNo >= @bordxFrom' ELSE N'' END
        + CASE WHEN NOT (@worksheet = 0 AND @scan = 1) AND @bordxTo IS NOT NULL THEN N' and Bordx.BordxRefNo <= @bordxTo' ELSE N'' END
        + CASE WHEN @bordxFrom IS NOT NULL OR @bordxTo IS NOT NULL THEN N' AND Claim.CASNumber IN (SELECT k.CASNumber FROM CISWorksheet.dbo.Claim k WHERE (k.BordxRefNo >= COALESCE(@bordxFrom, k.BordxRefNo) AND k.BordxRefNo <= COALESCE(@bordxEnd, k.BordxRefNo)) OR (k.BTBGBordxRefNo >= COALESCE(@bordxFrom, k.BTBGBordxRefNo) AND k.BTBGBordxRefNo <= COALESCE(@bordxEnd, k.BTBGBordxRefNo)))' ELSE N'' END
        + CASE WHEN @payor IS NOT NULL THEN N' And  GRPCompany.ASOINSPaycode = @payor' ELSE N'' END
        + CASE WHEN @client IS NOT NULL THEN N' And Bordx.GroupCompany = @client' ELSE N'' END
        + CASE WHEN @pvFrom IS NOT NULL THEN N' AND ReceiptfrPAY.ClientPVNo >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' AND ReceiptfrPAY.ClientPVNo <= @pvTo' ELSE N'' END
        + CASE WHEN @mrFrom IS NOT NULL THEN N' AND ReceiptfrPAY.MedixReceiptNo >= @mrFrom' ELSE N'' END
        + CASE WHEN @mrTo IS NOT NULL THEN N' AND ReceiptfrPAY.MedixReceiptNo <= @mrTo' ELSE N'' END
        + CASE WHEN @insured = 'YES' THEN N' AND substring(Bordx.BordxRefNo, 1, 3) = ''KDC'' ' ELSE N'' END
        + CASE WHEN @insured = 'NO' THEN N' AND substring(Bordx.BordxRefNo, 1, 3) <> ''KDC'' ' ELSE N'' END
        END;
    IF @worksheet = 0
    BEGIN
	              	declare @l_sqlstr18 as nvarchar(MAX)
					declare @l_sqlstrGRP18 as nvarchar(MAX)

			 set @l_sqlstr18 = N'  SELECT  
									Bordx.BordxRefNo, 
									Bordx.BordxPAYCode, 
									Bordx.BordxMNRBFRStatus, 
									Bordx.GroupCompany, 
									CASE WHEN PLN.PLNBTBG = 1 THEN SUM(COALESCE(Claim.CLMBordxAmt,0) - COALESCE (Claim.CLMBTBGAmt, 0)) ELSE SUM(Claim.CLMBordxAmt) END AS BordxClaimAmt,
									Bordx.BordxDate, 
									COALESCE (MBMOthers.Subsidiary, N'''') AS Subsidiary, 
									GRPCompany.ASOINSPaycode, 
									ReceiptfrPAY.ReceiptNo, 
									ReceiptfrPAY.ReceiptDate, 
									ReceiptfrPAY.ReceiptAmt AS RecPayAmt, 
									ReceiptfrPAY.ChequeNo, 
									ReceiptfrPAY.Remark, 
									SUM(ReceiptfrPAY.CheckAmt) AS CheckAmt, 
									ReceiptfrPAY.RECAccPeriodYR, 
									ReceiptfrPAY.RECAccPeriodMth, 
									ReceiptfrPAY.ClientPVNo, 
									CASE WHEN (substring(Bordx.BordxRefNo, 1, 3) = ''KDC'') THEN ''NO'' ELSE ''YES'' END AS isInsured, 
									CASE WHEN ISNULL(Bordx.BordxPAYFRStatus, '''') <> '''' THEN Bordx.BordxPAYFRStatus WHEN ISNULL(ReceiptfrPAY.ReceiptNo, '''') <> '''' THEN ''F'' ELSE Bordx.BordxPAYFRStatus END AS BordxPAYFRStatus, 
									CISPayment..ReceiptfrPAY.MedixReceiptNo 
									FROM    CISWorksheet..Bordx WITH (NOLOCK) 
									INNER JOIN CISWorksheet..Claim WITH (NOLOCK) ON Bordx.BordxRefNo = Claim.BordxRefNo or Bordx.BordxRefNo = Claim.BTBGBordxRefNo
									INNER JOIN CISDB..CISCas WITH (NOLOCK) ON CISCas.CASNumber = Claim.CASNumber 
									INNER JOIN HISDB..MBMOthers WITH (NOLOCK) ON MBMOthers.MBMNumber = CISCas.MBMnumber 
									INNER JOIN HISDB..MBM WITH (NOLOCK) ON MBM.MBMNumber = MBMOthers.MBMNumber 
									LEFT OUTER JOIN CISDB..GRPCompany WITH (NOLOCK) ON GRPCompany.GRPCoName = MBMOthers.MBMGroupCompany AND (GRPCompany.HISpaycode IS NULL OR GRPCompany.HISpaycode = HISDB..MBM.PAYCode) 
									INNER JOIN HISDB..MBMCLNOthers WITH (NOLOCK) ON CISCas.MBMnumber = MBMCLNOthers.MBMNumber 
									INNER JOIN CISDB..PLN WITH (NOLOCK) ON PLN.plncode = MBMCLNOthers.PLNType 
									LEFT OUTER JOIN CISPayment..ReceiptfrPAY ON Claim.CASNumber = ReceiptfrPAY.CASNumber     WHere 0=0  '
				Set @l_sqlstrGRP18 = ' GROUP BY Bordx.BordxRefNo, Bordx.BordxPAYCode, Bordx.BordxMNRBFRStatus, Bordx.GroupCompany, Bordx.BordxDate, COALESCE (MBMOthers.Subsidiary, N''''), GRPCompany.ASOINSPaycode, 
									   PLN.PLNBTBG, ReceiptfrPAY.ReceiptNo, ReceiptfrPAY.ReceiptDate, ReceiptfrPAY.ReceiptAmt, ReceiptfrPAY.ChequeNo, 
									   ReceiptfrPAY.Remark, ReceiptfrPAY.RECAccPeriodYR, ReceiptfrPAY.RECAccPeriodMth, ReceiptfrPAY.ClientPVNo, 
									   CASE WHEN ISNULL(Bordx.BordxPAYFRStatus, '''') <> '''' THEN Bordx.BordxPAYFRStatus WHEN ISNULL(ReceiptfrPAY.ReceiptNo, '''') <> '''' THEN ''F'' ELSE Bordx.BordxPAYFRStatus END, ReceiptfrPAY.MedixReceiptNo '
				 set @l_sqlstr18 = @l_sqlstr18 + @filter + @l_sqlstrGRP18 + N' OPTION (RECOMPILE)'
			 exec CISPayment.sys.sp_executesql @l_sqlstr18, @params, @company, @subsidiary, @bordxFrom, @bordxTo, @caseFrom, @caseTo, @recFrom, @recTo, @recDateFrom, @recDateTo, @mthFrom, @mthTo, @yrFrom, @yrTo, @payor, @client, @pvFrom, @pvTo, @mrFrom, @mrTo, @receiptNo, @bordxEnd
    END
    ELSE
    BEGIN
	              	declare @l_sqlstr19 as nvarchar(MAX)
					declare @l_sqlstrGRP19 as nvarchar(MAX)

			 set @l_sqlstr19 = N' SELECT Claim.CASNumber, case when (substring(ReceiptfrPAY.BordxNo,1,3)=''KDC'' and Claim.clmbtbgamt>0)  then  ReceiptfrPAY.BordxNo else Claim.BordxRefNo end as BordxRefNo,
									Claim.CLMTotalApproved, Bordx.BordxPAYCode, ReceiptfrPAY.ReceiptNo, ReceiptfrPAY.ReceiptDate, ReceiptfrPAY.ReceiptAmt, ReceiptfrPAY.RECAccPeriodMth, ReceiptfrPAY.RECAccPeriodYR, 
									ReceiptfrPAY.RecPAYAmt, ReceiptfrPAY.CheckAmt, 
									CASE WHEN ISNULL(CISWorksheet.dbo.Claim.PVRecPayorStatus, '''') <> '''' THEN CISWorksheet.dbo.Claim.PVRecPayorStatus WHEN ISNULL(ReceiptfrPAY.ReceiptNo, '''') <> '''' THEN ''F'' ELSE Claim.PVRecPayorStatus END AS PVRecPayorStatus, 
									Bordx.GroupCompany, Claim.CLMFinalApprove, Claim.CLMTotalActual, Claim.CLMBordxAmt, Bordx.BordxDate, HISDB.dbo.MBMOthers.Subsidiary, Claim.ClaimOpenDate, MBM.PAYCode, GRPCompany.ASOINSPaycode,
									ReceiptfrPAY.ClientPVNo, ReceiptfrPAY.MedixReceiptNo, case when (substring(ReceiptfrPAY.BordxNo,1,3)=''KDC'')  then  ''NO'' else ''YES'' end as isInsured FROM  CISWorksheet.dbo.Claim 
									LEFT JOIN CISPayment..ReceiptfrPAY ON Claim.CASNumber = ReceiptfrPAY.CASNumber 
									INNER JOIN CISWorksheet..Bordx ----ON case when (substring(ReceiptfrPAY.BordxNo,1,3)=''KDC'' and Claim.clmbtbgamt>0)  then  ReceiptfrPAY.BordxNo else CISWorksheet.dbo.Claim.BordxRefNo  end  =CISWorksheet.dbo.Bordx.BordxRefNo 
									  ON Bordx.BordxRefNo = Claim.BordxRefNo or Bordx.BordxRefNo = Claim.BTBGBordxRefNo
									INNER JOIN CISDB..CISCas ON CISCas.CASNumber = Claim.CASNumber 
									INNER JOIN HISDB..MBMOthers ON MBMOthers.MBMNumber = CISCas.MBMnumber 
									INNER JOIN HISDB..MBM ON CISCas.MBMnumber = MBM.MBMNumber 
									LEFT JOIN CISDB..GRPCompany ON GRPCompany.GRPCoName = MBMOthers.MBMGroupCompany AND (GRPCompany.HISpaycode IS NULL OR GRPCompany.HISpaycode = MBM.PAYCode) 
									WHere 0=0 
									'
				Set @l_sqlstrGRP19 = ' GROUP BY Claim.CASNumber, Claim.BordxRefNo, Claim.CLMTotalApproved, Bordx.BordxPAYCode, ReceiptfrPAY.ReceiptNo, ReceiptfrPAY.ReceiptDate, ReceiptfrPAY.ReceiptAmt, ReceiptfrPAY.RECAccPeriodMth, ReceiptfrPAY.RECAccPeriodYR, 
										ReceiptfrPAY.RecPAYAmt, ReceiptfrPAY.CheckAmt, Claim.PVRecPayorStatus, Bordx.GroupCompany, Claim.CLMFinalApprove, Claim.CLMTotalActual, Claim.CLMBordxAmt, Bordx.BordxDate, MBMOthers.Subsidiary, Claim.ClaimOpenDate, MBM.PAYCode, 
										GRPCompany.ASOINSPaycode,ReceiptfrPAY.ClientPVNo,ReceiptfrPAY.MedixReceiptNo,ReceiptfrPAY.MedixReceiptNo,ReceiptfrPAY.BordxNo,CISWorksheet.dbo.Claim.clmbtbgamt HAVING Claim.CASNumber IS NOT NULL '
				 set @l_sqlstr19 = @l_sqlstr19 + @filter + @l_sqlstrGRP19 + N' OPTION (RECOMPILE)'
			 exec CISPayment.sys.sp_executesql @l_sqlstr19, @params, @company, @subsidiary, @bordxFrom, @bordxTo, @caseFrom, @caseTo, @recFrom, @recTo, @recDateFrom, @recDateTo, @mthFrom, @mthTo, @yrFrom, @yrTo, @payor, @client, @pvFrom, @pvTo, @mrFrom, @mrTo, @receiptNo, @bordxEnd
    END
END
GO

-- Combos: group companies, CIS payors (CISDB..PAY), the payor's insured clients and the group company's subsidiaries.
CREATE OR ALTER PROCEDURE cis.RcpPay_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE COALESCE(GRPCoName, '') <> '' ORDER BY GRPCoName;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM dbo.PAY ORDER BY PAYCode;
END
GO

CREATE OR ALTER PROCEDURE cis.RcpPay_Clients @payor nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE ASOINSPaycode = @payor ORDER BY GRPCoName;
END
GO

-- Bordereaux tick: its status and the cases it takes (partial: the cases not yet received).
CREATE OR ALTER PROCEDURE cis.RcpPay_BordxCases @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @status nvarchar(5) = (SELECT TOP 1 RTRIM(BordxPAYFRStatus) FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @bordxNo);
    SELECT @status AS BordxPAYFRStatus;
    IF @status = 'P'
        SELECT RTRIM(CASNumber) AS CASNumber FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo AND PVRecPayorStatus IS NULL;
    ELSE
        SELECT RTRIM(CASNumber) AS CASNumber FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo OR BTBGBordxRefNo = @bordxNo;
END
GO

-- Worksheet tick: the case's receipt status.
CREATE OR ALTER PROCEDURE cis.RcpPay_CaseStatus @casNumber nvarchar(30), @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 RTRIM(PVRecPayorStatus) AS PVRecPayorStatus, CLMBordxAmt FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND BordxRefNo = @bordxNo;
END
GO

-- UpdateRecord (bordereaux save): the claim is received in full, a ReceiptfrPAY row is added and the bordereaux is received (F).
CREATE OR ALTER PROCEDURE cis.RcpPay_SaveBordxCase @casNumber nvarchar(30), @bordxNo nvarchar(20), @receiptNo nvarchar(30), @receiptDate datetime,
    @receiptAmt money, @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @amt money;
    SELECT TOP 1 @amt = CASE WHEN BTBGBordxRefNo IS NOT NULL THEN CLMBordxAmt - CLMBTBGAmt ELSE CLMBordxAmt END
    FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber AND (BordxRefNo = @bordxNo OR BTBGBordxRefNo = @bordxNo);
    UPDATE CISWorksheet.dbo.Claim SET PVRecPayorStatus = 'F' WHERE CASNumber = @casNumber AND (BordxRefNo = @bordxNo OR BTBGBordxRefNo = @bordxNo);
    SET @amt = COALESCE(@amt, 0);
    INSERT CISPayment.dbo.ReceiptfrPAY (ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, BordxNo, CASNumber, Remark, PAYLastUpdateDate, PAYLastUpdateUser, RECAccPeriodMth,
        RECAccPeriodYR, RecPayAmt, CheckAmt)
    VALUES (@receiptNo, @receiptDate, @receiptAmt, @chequeNo, @bordxNo, @casNumber, COALESCE(@remark, ''), CAST(GETDATE() AS date), @user, COALESCE(@mth, ''),
        COALESCE(@yr, ''), @amt, @amt);
    UPDATE CISWorksheet.dbo.Bordx SET BordxPAYRecDate = @receiptDate, BordxPAYFRStatus = 'F' WHERE BordxRefNo = @bordxNo;
END
GO

-- UpdateRecord2 (worksheet save): a ReceiptfrPAY row when the case has none for the bordereaux, the claim's status, then the bordereaux is F when all
-- its claims are received, else P.
CREATE OR ALTER PROCEDURE cis.RcpPay_SaveWksCase @casNumber nvarchar(30), @bordxNo nvarchar(20), @status nvarchar(2), @receiptNo nvarchar(30),
    @receiptDate datetime, @receiptAmt money, @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL,
    @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @amt money = (SELECT TOP 1 CLMBordxAmt FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo AND CASNumber = @casNumber);
    IF NOT EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptfrPAY WHERE CASNumber = @casNumber AND BordxNo = @bordxNo)
        INSERT CISPayment.dbo.ReceiptfrPAY (CASNumber, ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, BordxNo, PAYLastUpdateDate, PAYLastUpdateUser, RecPayAmt,
            CheckAmt, RECAccPeriodMth, RECAccPeriodYR)
        VALUES (@casNumber, @receiptNo, @receiptDate, @receiptAmt, @chequeNo, COALESCE(@remark, ''), @bordxNo, CAST(GETDATE() AS date), @user, COALESCE(@amt, 0),
            COALESCE(@amt, 0), @mth, @yr);
    UPDATE CISWorksheet.dbo.Claim SET PVRecPayorStatus = @status WHERE BordxRefNo = @bordxNo AND CASNumber = @casNumber;
    DECLARE @all int = (SELECT COUNT(*) FROM CISWorksheet.dbo.Claim WHERE BordxRefNo = @bordxNo),
        @full int = (SELECT COUNT(*) FROM CISWorksheet.dbo.Claim WHERE PVRecPayorStatus = 'F' AND BordxRefNo = @bordxNo);
    UPDATE CISWorksheet.dbo.Bordx SET BordxPAYRecDate = @receiptDate, BordxPAYFRStatus = CASE WHEN @all = @full THEN 'F' ELSE 'P' END WHERE BordxRefNo = @bordxNo;
END
GO

-- SaveBordx: the OP claim / OP fee lines of the receipt.
CREATE OR ALTER PROCEDURE cis.RcpPay_Summary @receiptNo nvarchar(30), @receiptDate datetime, @amount money, @type nvarchar(10), @desc nvarchar(500) = NULL,
    @chequeNo nvarchar(30), @payCode nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISPayment.dbo.ReceiptSummary (ReceiptNo, ReceiptDate, ReceiptAmt, ReceiptType, ReceiptDesc, RecChqNo, RecPayCode)
    VALUES (@receiptNo, @receiptDate, @amount, @type, COALESCE(@desc, ''), @chequeNo, @payCode);
END
GO

-- lstBordx_Click / lstWorksheet_Click in edit mode: the receipt (@casNumber NULL: the bordereaux's first receipt row).
CREATE OR ALTER PROCEDURE cis.RcpPay_Receipt @bordxNo nvarchar(20), @casNumber nvarchar(30) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, RECAccPeriodMth, RECAccPeriodYR
    FROM CISPayment.dbo.ReceiptfrPAY WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber);
END
GO

-- UpdateBordx (@casNumber NULL: every receipt row of the bordereaux, with the period) / UpdateWorksheet (the case's row); a row is added when there is none.
CREATE OR ALTER PROCEDURE cis.RcpPay_Edit @bordxNo nvarchar(20), @casNumber nvarchar(30) = NULL, @receiptNo nvarchar(30), @receiptDate datetime, @receiptAmt money,
    @chequeNo nvarchar(30), @remark nvarchar(500) = NULL, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF NOT EXISTS (SELECT 1 FROM CISPayment.dbo.ReceiptfrPAY WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber))
        INSERT CISPayment.dbo.ReceiptfrPAY (ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, Remark, PAYLastUpdateDate, PAYLastUpdateUser, RECAccPeriodMth, RECAccPeriodYR)
        VALUES (@receiptNo, @receiptDate, @receiptAmt, @chequeNo, COALESCE(@remark, ''), CAST(GETDATE() AS date), @user, CASE WHEN @casNumber IS NULL THEN @mth END,
            CASE WHEN @casNumber IS NULL THEN @yr END);
    ELSE
        UPDATE CISPayment.dbo.ReceiptfrPAY SET ReceiptNo = @receiptNo, ReceiptDate = @receiptDate, ReceiptAmt = @receiptAmt, ChequeNo = @chequeNo,
            Remark = COALESCE(@remark, ''), PAYLastUpdateDate = CAST(GETDATE() AS date), PAYLastUpdateUser = @user,
            RECAccPeriodMth = CASE WHEN @casNumber IS NULL AND @mth IS NOT NULL THEN @mth ELSE RECAccPeriodMth END,
            RECAccPeriodYR = CASE WHEN @casNumber IS NULL AND @yr IS NOT NULL THEN @yr ELSE RECAccPeriodYR END
        WHERE BordxNo = @bordxNo AND (@casNumber IS NULL OR CASNumber = @casNumber);
END
GO
