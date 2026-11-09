-- Claim Bordx > Monitoring: Bordx Control (FrmBordxControl: CISWorksheet ClaimBordxModule1 case 17 copied, the pasted @strAppend replaced by
-- parameters), Outstanding Bordx (FrmOSBordx, ClaimBordxModule case 14) and Bordx & Worksheet (FrmBordxWks, case 15).
SET QUOTED_IDENTIFIER ON;
GO

-- @sent / @returned / @paid: the desktop's check boxes (1 / 0 / F / P / U = unpaid; NULL = not ticked); @payTo C / M; @entity 0 Health Connect, 1 Sihat Johor.
CREATE OR ALTER PROCEDURE cis.BordxControl_Search @company nvarchar(200) = NULL, @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL,
    @dateFrom datetime = NULL, @dateTo datetime = NULL, @sent char(1) = NULL, @returned char(1) = NULL, @paid char(1) = NULL, @subsidiary nvarchar(200) = NULL,
    @payTo char(1) = NULL, @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = N''
        + CASE WHEN @company IS NOT NULL THEN N' AND Bordx.GroupCompany = @company' ELSE N'' END
        + CASE WHEN @bordxFrom IS NOT NULL THEN N' AND Bordx.BordxRefNo >= @bordxFrom' ELSE N'' END
        + CASE WHEN @bordxTo IS NOT NULL THEN N' AND Bordx.BordxRefNo <= @bordxTo' ELSE N'' END
        + CASE WHEN @dateFrom IS NOT NULL THEN N' And Bordx.BordxDate >= @dateFrom' ELSE N'' END
        + CASE WHEN @dateTo IS NOT NULL THEN N' And Bordx.BordxDate <= @dateTo' ELSE N'' END
        + CASE @sent WHEN '1' THEN N' AND BordxPAYToStatus = ''1''' WHEN '0' THEN N' AND BordxPAYToStatus = ''0''' ELSE N'' END
        + CASE @returned WHEN '1' THEN N' AND BordxPAYDOCFrStatus = ''1''' WHEN '0' THEN N' AND BordxPAYDOCFrStatus = ''0''' ELSE N'' END
        + CASE @paid WHEN 'F' THEN N' AND BordxPAYFrStatus = ''F''' WHEN 'P' THEN N' AND BordxPAYFrStatus = ''P'''
            WHEN 'U' THEN N' AND (BordxPAYFrStatus = '''' or BordxPAYFrStatus is null)' ELSE N'' END
        + CASE WHEN @subsidiary IS NOT NULL THEN N' AND BordxSubsidary = @subsidiary' ELSE N'' END
        + CASE @payTo WHEN 'C' THEN N' AND INVPayTo = ''C''' WHEN 'M' THEN N' AND INVPayTo = ''M'' ' ELSE N'' END
        + CASE @entity WHEN 0 THEN N' AND Substring(CISCas.CasNumber, 1, 2) <>''SJ'' ' WHEN 1 THEN N' AND Substring(CISCas.CasNumber, 1, 2) =''SJ'' ' ELSE N'' END;

			       declare @sql nvarchar(max)
				   declare @groupBy nvarchar(max)

		                    set @sql =  N' SELECT 
												Bordx.BordxRefNo, 
												Bordx.BordxDate, 
												Bordx.BordxPAYDOCFRStatus, 
												Bordx.BordxPAYToStatus, 
												Bordx.BordxPAYSentDate, 
												Bordx.BordxPAYRecDate, 
												Bordx.BordxClaimAmt, 
												COALESCE(Claim.CLMTotalActual,0) AS BillAmt, 
												Bordx.BordxPAYCode, 
												PAY.HLTCode, 
												Bordx.BordxPAYRecDocDate, 
												Bordx.BordxPAYFRStatus, 
												Bordx.GroupCompany, 
												COALESCE(Claim.CLMNonGPFee,0) AS CLMNonGpFee,
												Bordx.GroupCompany as MBMGroupCompany,
												ClaimInvoice.INVPayTo,
												CASE WHEN Pln.PLNBTBG = 1 THEN COALESCE(Claim.CLMBordxAmt,0) - COALESCE(Claim.CLMBTBGAmt,0) ELSE COALESCE(Claim.CLMFinalApprove,0) END as CLMFinalApprove,
												COALESCE(Claim.CLMMCOProcFee,0) AS CLMMCOProcFee, 
												COALESCE(Claim.CLMCoPayCLNColl,0) AS CLMCoPayCLNColl,
												CASE WHEN (PLN.PLNBTBG=1 AND substring(Bordx.BordxRefNo,1,3) IN (''KDC'',''KCN'')) then COALESCE(dbo.Claim.CLMBTBGAmt,0) ELSE 
												CASE WHEN PLN.PLNBTBG=1 THEN COALESCE(dbo.Claim.CLMBordxAmt,0)-COALESCE(Claim.CLMBTBGAmt,0) ELSE COALESCE(dbo.Claim.CLMBordxAmt,0) END END AS CLMBordxAmt,
												Bordx.BordxNoClaims as BordxNoClaims,
												Substring(CISCas.CasNumber, 1, 2) AS checkcase,Bordx.BordxSubsidary as Subsidiary
												into #tempcontrol
												FROM CISDB..CISCas 
												INNER JOIN CISWorksheet..Claim  ON CISDB.dbo.CISCas.CASNumber = dbo.Claim.CASNumber 
												INNER JOIN dbo.ClaimInvoice  ON dbo.Claim.CASNumber = dbo.ClaimInvoice.CASNumber                 
												INNER JOIN dbo.Bordx   ON ((Claim.BordxRefNo = Bordx.BordxRefNo) OR (Claim.BTBGBordxRefNo = Bordx.BordxRefNo))
												INNER JOIN HISDB..PAY  ON HISDB.dbo.PAY.PAYCode = dbo.Bordx.BordxPAYCode
												INNER JOIN hisdb..MBMCLNOthers  on MBMCLNOthers.MBMnumber=ciscas.MBMnumber
												INNER JOIN cisdb..pln  on pln.PLNCode=MBMCLNOthers.PLNType    
												INNER JOIN HISDB..MBMOthers ON CISCas.MBMnumber = MBMOthers.MBMNumber where 0=0  '
												
												set @groupBy = '  Select BordxRefNo, BordxDate, BordxPAYDOCFRStatus, BordxPAYToStatus, BordxPAYSentDate, 
												BordxPAYRecDate, BordxClaimAmt, SUM(BillAmt) as BillAmt, BordxPAYCode, HLTCode, 
												BordxPAYRecDocDate, BordxPAYFRStatus, GroupCompany, CLMNonGpFee,
												MBMGroupCompany,INVPayTo,SUM(CLMFinalApprove) as CLMFinalApprove,sum(CLMCoPayCLNColl) as CLMCoPayCLNColl,sum(CLMBordxAmt) as CLMBordxAmt ,
												BordxNoClaims,
												SUM(CLMMCOProcFee) as CLMMCOProcFee, DateDiff(d, BordxPAYSentDate, BordxPAYRecDocDate) AS NoofDay,DateDiff(d, BordxPAYSentDate, BordxPAYRecDate) AS NoofDay2
												,checkcase,Subsidiary from #tempcontrol
												GROUP BY BordxRefNo, BordxDate, BordxPAYDOCFRStatus, BordxPAYToStatus, BordxPAYSentDate, 
												BordxPAYRecDate, BordxClaimAmt,  BordxPAYCode, HLTCode, 
												BordxPAYRecDocDate, BordxPAYFRStatus, GroupCompany, CLMNonGpFee,
												MBMGroupCompany,INVPayTo,BordxNoClaims
												,checkcase,Subsidiary

												drop table #tempcontrol '
					--set @sql = @sql + @strAppend
		set @sql = @sql + @filter + @groupBy
					--print @sql
		exec CISWorksheet.sys.sp_executesql @sql, N'@company nvarchar(200), @bordxFrom nvarchar(20), @bordxTo nvarchar(20), @dateFrom datetime, @dateTo datetime, @subsidiary nvarchar(200)',
            @company, @bordxFrom, @bordxTo, @dateFrom, @dateTo, @subsidiary;
END
GO

CREATE OR ALTER PROCEDURE cis.BordxOs_Search @company nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 30000 BordxRefNo, BordxPAYCode, GroupCompany, BordxPAYSentDate, BordxClaimAmt, BordxPAYToStatus
    FROM CISWorksheet.dbo.VIEW_OS_CISBordx
    WHERE BordxPAYToStatus = 1 AND (BordxPAYFRStatus = 'P' OR BordxPAYFRStatus IS NULL) AND (@company IS NULL OR GroupCompany = @company);
END
GO

CREATE OR ALTER PROCEDURE cis.BordxWks_Search @company nvarchar(200) = NULL, @payCode nvarchar(10) = NULL, @bordxFrom nvarchar(20) = NULL,
    @bordxTo nvarchar(20) = NULL, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL, @receiptFrom nvarchar(30) = NULL, @receiptTo nvarchar(30) = NULL,
    @recFrom datetime = NULL, @recTo datetime = NULL, @noRecDate bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 30000 CASNumber, BordxPAYCode, BordxClaimAmt, BordxDate, BordxRefNo, CLMTotalApproved, ClaimOpendate, BordxPAYRecDate, ReceiptNoPAY, GroupCompany
    FROM CISWorksheet.dbo.VIEW_CISBordx_Wks
    WHERE (@company IS NULL OR GroupCompany = @company) AND (@payCode IS NULL OR BordxPAYCode = @payCode)
      AND (@bordxFrom IS NULL OR BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR BordxRefNo <= @bordxTo)
      AND (@caseFrom IS NULL OR CASNumber >= @caseFrom) AND (@caseTo IS NULL OR CASNumber <= @caseTo)
      AND (@receiptFrom IS NULL OR ReceiptNoPAY >= @receiptFrom) AND (@receiptTo IS NULL OR ReceiptNoPAY <= @receiptTo)
      AND (@recFrom IS NULL OR BordxPAYRecDate >= @recFrom) AND (@recTo IS NULL OR BordxPAYRecDate <= @recTo)
      AND (@noRecDate = 0 OR BordxPAYRecDate IS NULL OR BordxPAYRecDate = '')
    OPTION (RECOMPILE);
END
GO

-- FrmBordxControl's BindSubsidary (SearchSubsidiary of the group company name) and the drop-downs.
CREATE OR ALTER PROCEDURE cis.BordxMonitoring_Lookups @company nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT DISTINCT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM dbo.PAY;
    SELECT DISTINCT LTRIM(RTRIM(Subsidiary)) AS Name FROM cis.Hdb_MBMOthers
    WHERE @company IS NOT NULL AND MBMGroupCompany = @company AND COALESCE(LTRIM(RTRIM(Subsidiary)), '') <> '';
END
GO
