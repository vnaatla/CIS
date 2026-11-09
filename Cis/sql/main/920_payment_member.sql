-- Payment > To Member (frmPaymentMemberAuto08 with frmPrintPVRange / frmPrintChequeRange in member mode): CISPayment SP_CISPayment_TESTING 'Pay2MBM'
-- copied with the pasted criteria as parameters, Savecispayment2mbm, the PV wording and edit updates, the PV rows and the member cheques.
SET QUOTED_IDENTIFIER ON;
GO

-- @pvNo: the PV just saved / updated; @payment 0-6 as CboPayment; @bordx / @btbg 0 / 1 (NULL = not chosen); @blank*: the "empty" boxes.
CREATE OR ALTER PROCEDURE cis.PayMbm_Search @pvNo nvarchar(30) = NULL, @caseNo nvarchar(30) = NULL, @scan bit = 0, @caseFrom nvarchar(30) = NULL,
    @caseTo nvarchar(30) = NULL, @mbmNo nvarchar(30) = NULL, @mbmName nvarchar(200) = NULL, @company nvarchar(200) = NULL, @payment int = 0,
    @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL, @blankBatch bit = 0, @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @blankPv bit = 0,
    @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL, @blankPvDate bit = 0, @chqFrom nvarchar(50) = NULL, @chqTo nvarchar(50) = NULL, @blankChq bit = 0,
    @crFrom nvarchar(30) = NULL, @crTo nvarchar(30) = NULL, @blankCr bit = 0, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @blankMth bit = 0,
    @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL, @blankYr bit = 0, @bordx int = NULL, @bordxFrom nvarchar(20) = NULL, @bordxTo nvarchar(20) = NULL,
    @btbg int = NULL, @btbgFrom nvarchar(20) = NULL, @btbgTo nvarchar(20) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = CASE WHEN @pvNo IS NOT NULL THEN N' and PVNo >= @pvNo and PVNo <= @pvNo'
        WHEN @caseNo IS NOT NULL THEN N' And Claim.CASNumber = @caseFrom' ELSE N''
        + CASE WHEN @scan = 1 THEN N' AND CISWorksheet.dbo.ClaimInvoice.CASNumber = @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseFrom IS NOT NULL THEN N' AND CISWorksheet.dbo.ClaimInvoice.CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseTo IS NOT NULL THEN N' AND CISWorksheet.dbo.ClaimInvoice.CASNumber <= @caseTo' ELSE N'' END
        + N' And CISWorksheet.dbo.ClaimInvoice.INVPayTo = ''M'''
        + CASE WHEN @mbmNo IS NOT NULL THEN N' And CISDB.dbo.CISCas.MBMnumber = @mbmNo' ELSE N'' END
        + CASE WHEN @mbmName IS NOT NULL THEN N' And MBMName like ''%'' + @mbmName + ''%''' ELSE N'' END
        + CASE WHEN @company IS NOT NULL THEN N' AND MBMGroupCompany = @company' ELSE N'' END
        + CASE @payment WHEN 1 THEN N' and PVPayMBMStatus = ''F''' WHEN 2 THEN N' and (PVPayMBMStatus = '''' or PVPayMBMStatus is null) '
            WHEN 3 THEN N' and PVPayMBMStatus = ''P''' WHEN 4 THEN N' and (PVPayMBMStatus = ''P'' or PVPayMBMStatus = ''F'') '
            WHEN 5 THEN N' and (PVPayMBMStatus = '''' or PVPayMBMStatus is null or PVPayMBMStatus = ''P'') '
            WHEN 6 THEN N' and (PVPayMBMStatus in (''F'',''P'') and PVNo=''VOID'' and PVChequeNo=''VOID'') ' ELSE N'' END
        + CASE WHEN @payment <> 6 THEN N' and (Payment2MBM.PVNo<>''VOID'' or Payment2MBM.PVNo is null) ' ELSE N'' END
        + CASE WHEN @pvFrom IS NOT NULL THEN N' and PVNo >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' and PVNo <= @pvTo' ELSE N'' END
        + CASE WHEN @blankBatch = 1 THEN N'AND (OPBatchNo Is null OR OPBatchNo = '''') ' ELSE N'' END
        + CASE WHEN @opFrom IS NOT NULL THEN N' AND OPBatchNo >= @opFrom' ELSE N'' END
        + CASE WHEN @opTo IS NOT NULL THEN N' AND OPBatchNo <= @opTo' ELSE N'' END
        + CASE WHEN @blankPv = 1 THEN N'AND (PVNo Is null OR PVNo = '''') ' ELSE N'' END
        + CASE WHEN @pvDateFrom IS NOT NULL THEN N' And PVDate >= @pvDateFrom' ELSE N'' END
        + CASE WHEN @pvDateTo IS NOT NULL THEN N' And PVDate <= @pvDateTo' ELSE N'' END
        + CASE WHEN @blankPvDate = 1 THEN N'AND (PVDate Is null OR PVDate = '''') ' ELSE N'' END
        + CASE WHEN @chqFrom IS NOT NULL THEN N' and PVChequeNo >= @chqFrom' ELSE N'' END
        + CASE WHEN @chqTo IS NOT NULL THEN N' and PVChequeNo <= @chqTo' ELSE N'' END
        + CASE WHEN @blankChq = 1 THEN N'AND (PVChequeNo Is null OR PVChequeNo = '''') ' ELSE N'' END
        + CASE WHEN @crFrom IS NOT NULL THEN N' and DebitCreditRefNo >= @crFrom' ELSE N'' END
        + CASE WHEN @crTo IS NOT NULL THEN N' and DebitCreditRefNo <= @crTo' ELSE N'' END
        + CASE WHEN @blankCr = 1 THEN N'AND (DebitCreditRefNo Is null OR DebitCreditRefNo = '''') ' ELSE N'' END
        + CASE WHEN @mthFrom IS NOT NULL THEN N' and PAYAccPeriodMth >= @mthFrom' ELSE N'' END
        + CASE WHEN @mthTo IS NOT NULL THEN N' and PAYAccPeriodMth <= @mthTo' ELSE N'' END
        + CASE WHEN @blankMth = 1 THEN N' AND PAYAccPeriodMth Is null ' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL AND @yrFrom IS NOT NULL THEN N' and PAYAccPeriodYr >= @yrFrom' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL THEN N' and PAYAccPeriodYr <= @yrTo' ELSE N'' END
        + CASE WHEN @blankYr = 1 THEN N' AND PAYAccPeriodYr Is null ' ELSE N'' END
        + CASE @bordx WHEN 0 THEN N' and ClaimBordxStatus = ''0''' WHEN 1 THEN N' and ClaimBordxStatus = 1' ELSE N'' END
        + CASE WHEN @bordxFrom IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.BordxRefNo >= @bordxFrom' ELSE N'' END
        + CASE WHEN @bordxTo IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.BordxRefNo <= @bordxTo' ELSE N'' END
        + CASE @btbg WHEN 0 THEN N' and BTBGBordxStatus = ''0''' WHEN 1 THEN N' and BTBGBordxStatus = 1' ELSE N'' END
        + CASE WHEN @btbgFrom IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.BTBGBordxRefNo >= @btbgFrom' ELSE N'' END
        + CASE WHEN @btbgTo IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.BTBGBordxRefNo <= @btbgTo' ELSE N'' END
        END;
    IF @caseNo IS NOT NULL SET @caseFrom = @caseNo;
			declare @sql nvarchar(max)
			declare @orderBy nvarchar(max)
			set @sql = 'SELECT  CISCas.MBMnumber,CISCas.CoverID,CISCas.DateVisit,CISCas.Remarks,CISCas.Pending,CISCas.Claimability, 
							CISCas.CASStatus,Payment2MBM.PVNo,Payment2MBM.PVMBMIndex,Payment2MBM.PVDate,Payment2MBM.PVChequeNo,Payment2MBM.PVChequeDate, 
							Payment2MBM.PAYAccPeriodMth,Payment2MBM.MBMPaidAmt,Payment2MBM.PAYAccPeriodYr,Payment2MBM.PVPayMBMStatus, 
							Payment2MBM.Pay2MBMLastUpdateDate,Payment2MBM.Pay2MBMLastUpdateUser,CLN.CLNName,CISCas.ClinicCode,Claim.DebitCreditRefNo,
							ClaimInvoice.INVPayTo,ClaimInvoice.INVEnteredBy,ClaimInvoice.BillNo,Claim.CASNumber,Claim.BordxRefNo, 
							Claim.CLMPayableByMedix2M,Claim.ClaimCloseDate,Payment2MBM.PVPayto,Bordx.GroupCompany,Payment2MBM.PVPayeeBank,Payment2MBM.PVPayeeACNo, 
							Payment2MBM.Remark,Payment2MBM.PVCLNChequeAmt,CISCas.OPBatchNo,Claim.CLMPostageFee,Claim.CLMNonGPFee,mbm.MBMName, 
							Claim.ClaimBordxStatus,Payment2MBM.PVWordingOne,Payment2MBM.PVWordingTwo,Claim.CLMFinalApprove,MBMOthersTwo.MBMBankACNo, 
							MBMOthersTwo.MBMBankName,Claim.CLMTotalActual,MBMOthers.MBMGroupCompany,CISCas.InvoiceNo,MBMOthers.MBMGroupCategory, 
							MBMOthers.MBMCostCenter,MBMOthers.MBMSubDivision,MBMOthers.MBMDivision,MBMOthers.Subsidiary, 
							MBMOthers.MBMEmployeeNo,GRPCompany.BusinessChanel,Tbl_BusinessChaneel.Description AS Descr,MBMOthers.MBMEmployeeNo , 
							cisdb.dbo.fnGetTreatmentDesc(CASTreatReason1) as CASTreatReason1,cisdb.dbo.fnGetTreatmentDesc(CASTreatReason2) as CASTreatReason2,
							Claim.BTBGBordxStatus,Claim.BTBGBordxRefNo,Claim.CLMPayableByMedix2M,Claim.CLMNonGPFee FROM  CISDB..CISCas 
							LEFT JOIN HISDB..MBMOthers ON MBMOthers.MBMNumber=CISCas.MBMnumber
							INNER JOIN HISDB..mbm ON MBMOthers.MBMNumber = mbm.MBMNumber 				
							LEFT JOIN HISDB..MBMOthersTwo ON mbm.MBMNumber = MBMOthersTwo.MBMNumber
							LEFT JOIN CISWorksheet..Claim ON CISCas.CASNumber = Claim.CASNumber 
							LEFT JOIN CISWorksheet..Bordx ON Claim.BordxRefNo = Bordx.BordxRefNo				
							LEFT JOIN CISWorksheet..ClaimInvoice ON Claim.CASNumber = ClaimInvoice.CASNumber	
							LEFT JOIN CISPayment..Payment2MBM ON Claim.CASNumber = Payment2MBM.CASNumber 
							LEFT JOIN HISDB..CLN ON CISCas.ClinicCode = CLN.CLNCode 
							LEFT JOIN CISDB..GRPCompany 
							on GRPCompany.GRPCoCode=claim.clmgrpcomcode and (grpcompany.hispaycode is null or ((grpcompany.hispaycode=mbm.paycode or grpcompany.hispaycodePrevYear = mbm.paycode) or (grpcompany.GRPCoPaycode=''DC'' and mbm.MBMPayorEffDate is null)))
							LEFT JOIN HISMaintenance..Tbl_BusinessChaneel ON GRPCompany.BusinessChanel = Tbl_BusinessChaneel.PreFix
							where				  
							(coalesce(Claim.CLMPayableByMedix2M,0) + coalesce(Claim.CLMNonGPFee,0) <> 0) AND CISCas.CASStatus = ''O'' and  CISCas.Claimability = ''A''  '

			set @orderBy=' order by Claim.CASNumber asc, CISCas.ClinicCode '
			set @sql = @sql + @filter + @orderBy
			exec CISPayment.sys.sp_executesql @sql, N'@caseFrom nvarchar(30), @caseTo nvarchar(30), @mbmNo nvarchar(30), @mbmName nvarchar(200), @company nvarchar(200), @pvFrom nvarchar(30), @pvTo nvarchar(30), @opFrom nvarchar(30), @opTo nvarchar(30), @pvDateFrom datetime, @pvDateTo datetime, @chqFrom nvarchar(50), @chqTo nvarchar(50), @crFrom nvarchar(30), @crTo nvarchar(30), @mthFrom nvarchar(2), @mthTo nvarchar(2), @yrFrom nvarchar(4), @yrTo nvarchar(4), @bordxFrom nvarchar(20), @bordxTo nvarchar(20), @btbgFrom nvarchar(20), @btbgTo nvarchar(20), @pvNo nvarchar(30)', @caseFrom, @caseTo, @mbmNo, @mbmName, @company, @pvFrom, @pvTo, @opFrom, @opTo, @pvDateFrom, @pvDateTo, @chqFrom, @chqTo, @crFrom, @crTo, @mthFrom, @mthTo, @yrFrom, @yrTo, @bordxFrom, @bordxTo, @btbgFrom, @btbgTo, @pvNo;
END
GO

-- Payment2MBMListing's patient name (cover 00: MBM.MBMName, else MBMCoveredPersons.MBMCName) and the partial-paid total per case.
CREATE OR ALTER PROCEDURE cis.PayMbm_Extra @cases nvarchar(max) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @x xml = CAST('<c>' + REPLACE(REPLACE(REPLACE(REPLACE(@cases, '&', ''), '<', ''), ' ', ''), ',', '</c><c>') + '</c>' AS xml);
    DECLARE @c TABLE (CASNumber nvarchar(30) PRIMARY KEY);
    INSERT @c SELECT DISTINCT t.c.value('.', 'nvarchar(30)') FROM @x.nodes('/c') t(c) WHERE t.c.value('.', 'nvarchar(30)') <> '';
    SELECT RTRIM(c.CASNumber) AS CASNumber,
        CASE WHEN c.CoverID = '00' THEN RTRIM(m.MBMName) ELSE RTRIM(cp.MBMCName) END AS PatientName,
        (SELECT SUM(p.MBMPaidAmt) FROM CISPayment.dbo.Payment2MBM p WHERE p.CASNumber = c.CASNumber) AS PaidTotal,
        (SELECT TOP 1 RTRIM(p.PVPayMBMStatus) FROM CISPayment.dbo.Payment2MBM p WHERE p.CASNumber = c.CASNumber) AS FirstStatus
    FROM dbo.CISCas c
    LEFT JOIN cis.Hdb_MBM m ON m.MBMNumber = c.MBMnumber
    LEFT JOIN cis.Hdb_MBMCoveredPersons cp ON cp.MBMCNumber = c.MBMnumber AND cp.MBMCCoverID = c.CoverID
    WHERE c.CASNumber IN (SELECT CASNumber FROM @c);
END
GO

-- The case's Payment2MBM rows (ProcessCheckedItem: status; Payment2MBMListing: the partial paid total).
CREATE OR ALTER PROCEDURE cis.PayMbm_CaseRows @casNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(PVMBMIndex AS nvarchar(20)) AS PVMBMIndex, RTRIM(PVNo) AS PVNo, RTRIM(PVPayMBMStatus) AS PVPayMBMStatus, MBMPaidAmt
    FROM CISPayment.dbo.Payment2MBM WHERE CASNumber = @casNumber;
END
GO

-- UpdateRecord: a partial row only gets the new status; otherwise Savecispayment2mbm's insert.
CREATE OR ALTER PROCEDURE cis.PayMbm_Save @index nvarchar(20) = NULL, @casNumber nvarchar(15), @pvNo nvarchar(15), @pvDate datetime, @chequeDate datetime,
    @chequeNo nvarchar(50), @payTo nvarchar(60), @bank nvarchar(100), @account nvarchar(30), @remark nvarchar(200), @paidAmt money, @user nvarchar(10),
    @mth nvarchar(2), @yr nvarchar(4), @status char(2) AS
BEGIN
    SET NOCOUNT ON;
    IF @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.Payment2MBM WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index)
    BEGIN
        UPDATE CISPayment.dbo.Payment2MBM SET PVPayMBMStatus = LEFT(@status, 1) WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index AND RTRIM(PVPayMBMStatus) = 'P';
        RETURN;
    END
    INSERT CISPayment.dbo.Payment2MBM (CasNumber, PVNo, PVDate, PVChequeDate, PVChequeNo, PVPayto, PVPayeeBank, PVPayeeACNo, Remark, MBMPaidAmt, Pay2MBMLastUpdateDate,
        Pay2MBMLastUpdateUser, PAYAccPeriodMth, PAYAccPeriodYr, PVPayMBMStatus)
    VALUES (@casNumber, @pvNo, @pvDate, @chequeDate, @chequeNo, @payTo, @bank, @account, COALESCE(@remark, ''), @paidAmt, CAST(GETDATE() AS date), @user, @mth, @yr, @status);
END
GO

-- CmdSave: the PV's cheque details and wording.
CREATE OR ALTER PROCEDURE cis.PayMbm_PvCheque @pvNo nvarchar(15), @chequeDate datetime, @chequeNo nvarchar(50), @chequeAmt money, @payTo nvarchar(60),
    @wordingOne nvarchar(500), @wordingTwo nvarchar(500) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISPayment.dbo.Payment2MBM SET PVChequeDate = @chequeDate, PVChequeNo = @chequeNo, PVCLNChequeAmt = @chequeAmt, PVPayto = @payTo, PVWordingOne = @wordingOne,
        PVWordingTwo = @wordingTwo
    WHERE PVNo = @pvNo;
END
GO

-- CmdUpdate: the rows of the PV (or the case's VOID rows when the PV has none).
CREATE OR ALTER PROCEDURE cis.PayMbm_PvRows @pvNo nvarchar(15), @casNumber nvarchar(15) = NULL AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS (SELECT 1 FROM CISPayment.dbo.Payment2MBM WHERE PVNo <> 'VOID' AND PVNo = @pvNo)
        SELECT CAST(PVMBMIndex AS nvarchar(20)) AS PVMBMIndex, RTRIM(CASNumber) AS CASNumber, CAST(0 AS bit) AS Void
        FROM CISPayment.dbo.Payment2MBM WHERE PVNo <> 'VOID' AND PVNo = @pvNo;
    ELSE
        SELECT CAST(PVMBMIndex AS nvarchar(20)) AS PVMBMIndex, RTRIM(CASNumber) AS CASNumber, CAST(1 AS bit) AS Void
        FROM CISPayment.dbo.Payment2MBM WHERE PVNo = 'VOID' AND CASNumber = @casNumber;
END
GO

-- UpdatePVPaymentMBM.
CREATE OR ALTER PROCEDURE cis.PayMbm_Edit @index nvarchar(20), @casNumber nvarchar(15), @pvDate datetime, @chequeDate datetime, @chequeNo nvarchar(50),
    @chequeAmt money, @payTo nvarchar(60), @remark nvarchar(200), @wordingOne nvarchar(500), @wordingTwo nvarchar(500) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISPayment.dbo.Payment2MBM SET CASNumber = @casNumber, PVDate = @pvDate, PVChequeDate = @chequeDate, PVChequeNo = @chequeNo, PVCLNChequeAmt = @chequeAmt,
        PVPayto = @payTo, Remark = COALESCE(@remark, ''), PVWordingOne = @wordingOne, PVWordingTwo = @wordingTwo
    WHERE CAST(PVMBMIndex AS nvarchar(20)) = @index;
END
GO

-- PreviewPV (@cases) / the member PV range (@pvFrom..@pvTo): View_Payment2MBM rows with the member's address.
CREATE OR ALTER PROCEDURE cis.PayMbm_PvPrint @cases nvarchar(max) = NULL, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @c TABLE (CASNumber nvarchar(30) PRIMARY KEY);
    IF @cases IS NOT NULL
    BEGIN
        DECLARE @x xml = CAST('<c>' + REPLACE(REPLACE(REPLACE(REPLACE(@cases, '&', ''), '<', ''), ' ', ''), ',', '</c><c>') + '</c>' AS xml);
        INSERT @c SELECT DISTINCT t.c.value('.', 'nvarchar(30)') FROM @x.nodes('/c') t(c) WHERE t.c.value('.', 'nvarchar(30)') <> '';
    END
    ELSE
        INSERT @c SELECT DISTINCT RTRIM(CASNumber) FROM CISPayment.dbo.Payment2MBM WHERE PVNo >= @pvFrom AND PVNo <= @pvTo;
    SELECT v.CASNumber, v.BordxRefNo, v.MBMPaidAmt, v.MBMnumber, v.CLMPayableByMedix2M, v.ClaimCloseDate, v.PVNo, v.PVChequeNo, v.PVPayto, v.PVDate, v.CLMPostageFee,
        v.PVCLNChequeAmt, v.CLMFinalApprove, v.MBMName, m.MBMAddress1, m.MBMAddress2, m.MBMAddress3, m.MBMPostCode, m.CTYCode, m.STACode
    FROM CISPayment.dbo.View_Payment2MBM v
    LEFT JOIN cis.Hdb_MBM m ON m.MBMNumber = v.MBMnumber
    WHERE v.CASNumber IN (SELECT CASNumber FROM @c) AND (@pvFrom IS NULL OR (v.PVNo >= @pvFrom AND v.PVNo <= @pvTo))
    ORDER BY v.PVNo, v.CASNumber;
END
GO

-- frmPrintChequeRange (member): the cheques (VIEW_Payment2MBM_Chq, or VIEW_Payment2MBM_Acc for Print Acc No) and their print status.
CREATE OR ALTER PROCEDURE cis.PayMbm_Cheques @chequeFrom nvarchar(50), @chequeTo nvarchar(50), @account bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    IF @account = 1
        SELECT PVChequeNo, PVChequeDate, PVCLNChequeAmt, PVWordingOne, PVWordingTwo, PVPayto, PVPayeeACNo, MBMGroupCompany
        FROM CISPayment.dbo.VIEW_Payment2MBM_Acc WHERE PVChequeNo >= @chequeFrom AND PVChequeNo <= @chequeTo ORDER BY PVChequeNo;
    ELSE
        SELECT PVChequeNo, PVChequeDate, PVCLNChequeAmt, PVWordingOne, PVWordingTwo, PVPayto, PVPayeeACNo, MBMGroupCompany
        FROM CISPayment.dbo.VIEW_Payment2MBM_Chq WHERE PVChequeNo >= @chequeFrom AND PVChequeNo <= @chequeTo ORDER BY PVChequeNo;
END
GO

CREATE OR ALTER PROCEDURE cis.PayMbm_ChequesPrinted @chequeFrom nvarchar(50), @chequeTo nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISPayment.dbo.Payment2MBM SET PVChequeStatus = 'Y' WHERE PVChequeNo >= @chequeFrom AND PVChequeNo <= @chequeTo;
END
GO
