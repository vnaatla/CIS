-- Payment > To Clinic - Invoice (frmPaymentCLNInvoice, frmPrintPVRange, frmPrintChequeRange, frmInvoiceDetPopup): CISPayment SP_CISPayment
-- 'PaymentCLNV2' and 'Pay2CLN' copied with the pasted criteria replaced by parameters; the save's case list, the cheques and their print status.
SET QUOTED_IDENTIFIER ON;
GO

-- The list (SearchRecord): one row per clinic / invoice month / year / PV. @pvNo..@pvLast: the PVs just saved / updated (Payment2CLNListing by PV).
-- @payment 1 paid, 2 unpaid; @blank*: the "empty" boxes; @entity 0 Health Connect, 1 Sihat Johor.
CREATE OR ALTER PROCEDURE cis.PayClnInv_Search @pvNo nvarchar(30) = NULL, @pvLast nvarchar(30) = NULL, @clinic nvarchar(20) = NULL, @yrFrom nvarchar(4) = NULL,
    @yrTo nvarchar(4) = NULL, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @payment int = 0, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL,
    @blankPv bit = 0, @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL, @blankPvDate bit = 0, @acMthFrom nvarchar(2) = NULL, @acMthTo nvarchar(2) = NULL,
    @blankAcMth bit = 0, @acYrFrom nvarchar(4) = NULL, @acYrTo nvarchar(4) = NULL, @blankAcYr bit = 0, @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = CASE WHEN @pvNo IS NOT NULL THEN N' PVCLNNo >= @pvNo and PVCLNNo <= COALESCE(@pvLast, @pvNo)' ELSE N''
        + N' SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> ''EM'' and SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> ''ES'' '
        + CASE WHEN @clinic IS NOT NULL THEN N' And ClinicCode = @clinic' ELSE N'' END
        + CASE WHEN @yrFrom IS NOT NULL THEN N' and CASInvYear >= @yrFrom' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL THEN N' and CASInvYear <= @yrTo' ELSE N'' END
        + CASE WHEN @mthFrom IS NOT NULL THEN N' and CASInvMonth >= @mthFrom' ELSE N'' END
        + CASE WHEN @mthTo IS NOT NULL THEN N' and CASInvMonth <= @mthTo' ELSE N'' END
        + CASE @payment WHEN 1 THEN N' and Payment2CLN.PVPayCLNStatus = ''F''' WHEN 2 THEN N' and (Payment2CLN.PVPayCLNStatus = '''' or Payment2CLN.PVPayCLNStatus is null) ' ELSE N'' END
        + N' and (INVPayTo  = ''C'')  and (CLMPayableByMedix2C  <> ''0'') '
        + CASE WHEN @pvFrom IS NOT NULL THEN N' and PVCLNNo >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' and PVCLNNo <= @pvTo' ELSE N'' END
        + CASE WHEN @blankPv = 1 THEN N'AND (PVCLNNo Is null OR PVCLNNo = '''') ' ELSE N'' END
        + CASE WHEN @pvDateFrom IS NOT NULL THEN N' And PVCLNDate >= @pvDateFrom' ELSE N'' END
        + CASE WHEN @pvDateTo IS NOT NULL THEN N' And PVCLNDate <= @pvDateTo' ELSE N'' END
        + CASE WHEN @blankPvDate = 1 THEN N'AND (PVCLNDate Is null OR PVCLNDate = '''') ' ELSE N'' END
        + CASE WHEN @acMthFrom IS NOT NULL THEN N' and CLNAccPeriodMth >= @acMthFrom' ELSE N'' END
        + CASE WHEN @acMthTo IS NOT NULL THEN N' and CLNAccPeriodMth <= @acMthTo' ELSE N'' END
        + CASE WHEN @blankAcMth = 1 THEN N' AND CLNAccPeriodMth Is null ' ELSE N'' END
        + CASE WHEN @acYrTo IS NOT NULL AND @acYrFrom IS NOT NULL THEN N' and CLNAccPeriodYr >= @acYrFrom' ELSE N'' END
        + CASE WHEN @acYrTo IS NOT NULL THEN N' and CLNAccPeriodYr <= @acYrTo' ELSE N'' END
        + CASE @entity WHEN 0 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) <>''SJ'' ' WHEN 1 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) =''SJ'' ' ELSE N'' END
        + CASE WHEN @blankAcYr = 1 THEN N' AND CLNAccPeriodYr Is null ' ELSE N'' END
        END;
			declare @sql nvarchar(max)
			declare @groupBy nvarchar(max)
			declare @tmpINVPayto as char(1)
			set @tmpINVPayto  = 'C'
			Set @sql ='SELECT TOP (100000) CISDB.dbo.CISCas.ClinicCode, dbo.Payment2CLN.PVPayCLNStatus, SUM(CISWorksheet.dbo.Claim.CLMPayableByMedix2C) AS CLMPayableByMedix2C, 
                      dbo.Payment2CLN.PVCLNNo, dbo.Payment2CLN.PVCLNDate, dbo.Payment2CLN.PVCLNChequeDate, dbo.Payment2CLN.PVCLNChequeNo, 
                      SUM(dbo.Payment2CLN.CLNPaidAmt) AS CLNPaidAmt, dbo.Payment2CLN.PVCLNChequeAmt, dbo.Payment2CLN.PVCLNBank, dbo.Payment2CLN.PVCLNAccount, 
                      dbo.Payment2CLN.CLNRemark, SUM(CISWorksheet.dbo.Claim.CLMNonGPFee) AS CLMNonGPFee, 
                      SUM(CISWorksheet.dbo.Claim.CLMTotalActual - CISWorksheet.dbo.Claim.CLMCoPayCLNColl) AS CLMFinalApprove, SUM(CISWorksheet.dbo.Claim.CLMMCOProcFee) 
                      AS CLMMCOProcFee, dbo.Payment2CLN.PVWordingOne, dbo.Payment2CLN.PVWordingTwo, SUM(CISWorksheet.dbo.Claim.CLMTotalActual) AS CLMTotalActual, 
                      CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISDB.dbo.CISCas.CASInvMonth, CISDB.dbo.CISCas.CASInvYear,SUM(CISWorksheet.dbo.Claim.CLMCoPayCLNColl) as CoPay
					  FROM         CISDB.dbo.CISCas 
					INNER JOIN CISWorksheet..ClaimInvoice ON CISCas.CASNumber = ClaimInvoice.CASNumber 
					INNER JOIN HISDB..MBM ON CISCas.MBMnumber = MBM.MBMNumber 
					Left JOIN CISWorksheet..Claim ON CISCas.CASNumber = Claim.CASNumber 
					left JOIN CISPayment..Payment2CLN ON Payment2CLN.CASNumber = Claim.CASNumber
					  Where coalesce(CISCas.InvoiceNo,'''')<>'''' AND '
				Set @sql = @sql + ' CISCas.CASStatus =''O'' and  CISCas.Claimability = ''A'' AND  ' 

			Set @groupBy = 'GROUP BY CISDB.dbo.CISCas.ClinicCode,Payment2CLN.PVPayCLNStatus,
					Payment2CLN.PVCLNNo,Payment2CLN.PVCLNDate,Payment2CLN.PVCLNChequeDate,Payment2CLN.PVCLNChequeNo,
					Payment2CLN.PVCLNChequeAmt,Payment2CLN.PVCLNBank,Payment2CLN.PVCLNAccount,Payment2CLN.CLNRemark, 
					Payment2CLN.PVWordingOne,Payment2CLN.PVWordingTwo,CISWorksheet.dbo.ClaimInvoice.INVPayTo,
					CISDB.dbo.CISCas.CASInvMonth , CISDB.dbo.CISCas.CASInvYear '
			set @sql = @sql + @filter + @groupBy
			exec CISPayment.sys.sp_executesql @sql, N'@clinic nvarchar(20), @yrFrom nvarchar(4), @yrTo nvarchar(4), @mthFrom nvarchar(2), @mthTo nvarchar(2), @pvFrom nvarchar(30), @pvTo nvarchar(30), @pvDateFrom datetime, @pvDateTo datetime, @acMthFrom nvarchar(2), @acMthTo nvarchar(2), @acYrFrom nvarchar(4), @acYrTo nvarchar(4), @pvNo nvarchar(30), @pvLast nvarchar(30)', @clinic, @yrFrom, @yrTo, @mthFrom, @mthTo, @pvFrom, @pvTo, @pvDateFrom, @pvDateTo, @acMthFrom, @acMthTo, @acYrFrom, @acYrTo, @pvNo, @pvLast;
END
GO

-- frmInvoiceDetPopup (SearchRecordTopup: the row's clinic / invoice month / year) and frmPrintPVRange (PVCLNListing: a PV range): the Pay2CLN branch.
CREATE OR ALTER PROCEDURE cis.PayClnInv_Invoices @clinic nvarchar(20) = NULL, @yr nvarchar(4) = NULL, @mth nvarchar(2) = NULL, @pvFrom nvarchar(30) = NULL,
    @pvTo nvarchar(30) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = CASE WHEN @pvFrom IS NOT NULL OR @pvTo IS NOT NULL THEN N' 0 = 0 '
        + CASE WHEN @pvFrom IS NOT NULL THEN N' and dbo.Payment2CLN.PVCLNNo >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' and dbo.Payment2CLN.PVCLNNo <= @pvTo' ELSE N'' END
        + N' and InvoiceNo <> '''' '
        ELSE N' SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> ''EM'' and SUBSTRING(HISDB.dbo.MBM.MBMNumber, 1, 2)  <> ''ES'' '
        + N' and CASInvYear >= @yr and CASInvYear <= @yr and CASInvMonth >= @mth and CASInvMonth <= @mth and (CLMPayableByMedix2C  <> ''0'') and (INVPayTo  = ''C'') '
        + CASE WHEN @clinic IS NOT NULL THEN N' and CISCas.ClinicCode = @clinic' ELSE N'' END
        END;
			declare @sql nvarchar(max)
			declare @groupBy nvarchar(max)

			
			set @sql =  '   SELECT CISCas.ClinicCode, Payment2CLN.PVPayCLNStatus, SUM(Claim.CLMPayableByMedix2C) AS CLMPayableByMedix2C, Payment2CLN.PVCLNNo, Payment2CLN.PVCLNDate, 
								Payment2CLN.PVCLNChequeDate, Payment2CLN.PVCLNChequeNo, SUM(Payment2CLN.CLNPaidAmt) AS CLNPaidAmt, CISCas.InvoiceDate, Payment2CLN.PVCLNChequeAmt, 
								Payment2CLN.PVCLNBank, Payment2CLN.PVCLNAccount, Payment2CLN.CLNRemark, SUM(Claim.CLMNonGPFee) AS CLMNonGPFee, 
								SUM(coalesce(Claim.CLMTotalActual,0) - coalesce(Claim.CLMCoPayColl,0)) AS CLMFinalApprove, SUM(Claim.CLMMCOProcFee) AS CLMMCOProcFee, 
								Payment2CLN.PVWordingOne, Payment2CLN.PVWordingTwo, SUM(Claim.CLMTotalActual) AS CLMTotalActual, ClaimInvoice.INVPayTo, CISCas.InvoiceNo
								,CLN.CLNName,CLN.CLNAdd1,CLN.CLNAdd2,CLN.CLNAdd3,CLN.CLNPostCode,CLN.CLNCity,CLN.CLNState,CLN.CLNRemark,CLN.CLNBankAcc,CLN.CLNPName,CLN.CLNPAdd1,CLN.CLNPAdd2,
								CLN.CLNPAdd3,CLN.CLNType,CLN.CLNFFSStatus,CLN.CLNPName,CLN.CLNRegNo,CLN.CLNMCOProcfeePerc,CLN.CLNBankName,CLN.CLNBankAcc
								FROM  CISDB..CISCas 
								INNER JOIN CISWorksheet..ClaimInvoice ON CISCas.CASNumber = ClaimInvoice.CASNumber 
								INNER JOIN HISDB..MBM ON CISCas.MBMnumber = MBM.MBMNumber 
								LEFT JOIN CISWorksheet..Claim  ON CISCas.CASNumber = Claim.CASNumber 
								LEFT JOIN CISPayment..Payment2CLN ON Payment2CLN.CASNumber = Claim.CASNumber
								Inner Join HISDB..cln ON cln.clncode=CISCas.ClinicCode
								WHERE CISCas.CASStatus =''O'' and   CISCas.Claimability =''A'' AND coalesce(CISCas.InvoiceNo,'''')<>'''' And  '
			set @groupBy =  '  GROUP BY CISCas.ClinicCode,Payment2CLN.PVPayCLNStatus, CISCas.InvoiceNo,Payment2CLN.PVCLNNo,Payment2CLN.PVCLNDate,
								Payment2CLN.PVCLNChequeDate,Payment2CLN.PVCLNChequeNo,CISCas.InvoiceDate,Payment2CLN.PVCLNChequeAmt, 
								Payment2CLN.PVCLNBank,Payment2CLN.PVCLNAccount,Payment2CLN.CLNRemark,Payment2CLN.PVWordingOne,
								Payment2CLN.PVWordingTwo,ClaimInvoice.INVPayTo 
								,CLN.CLNName,CLN.CLNAdd1,CLN.CLNAdd2,CLN.CLNAdd3,CLN.CLNPostCode,CLN.CLNCity,CLN.CLNState,CLN.CLNRemark,CLN.CLNBankAcc,CLN.CLNPName,CLN.CLNPAdd1,CLN.CLNPAdd2,
								CLN.CLNPAdd3,CLN.CLNType,CLN.CLNFFSStatus,CLN.CLNPName,CLN.CLNRegNo,CLN.CLNMCOProcfeePerc,CLN.CLNBankName,CLN.CLNBankAcc
								order by CISCas.ClinicCode  '
			set @sql = @sql + @filter + @groupBy
			exec CISPayment.sys.sp_executesql @sql, N'@clinic nvarchar(20), @yr nvarchar(4), @mth nvarchar(2), @pvFrom nvarchar(30), @pvTo nvarchar(30)', @clinic, @yr, @mth, @pvFrom, @pvTo;
END
GO

-- CmdSave_Click: the unpaid cases of a clinic for the year and the ticked invoice months (the desktop's StrInvSql).
CREATE OR ALTER PROCEDURE cis.PayClnInv_Cases @clinic nvarchar(20), @yr nvarchar(4), @months nvarchar(200), @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT c.CASNumber, b.BordxDate, p.PVCLNNo, cl.CLMPayableByMedix2C
    FROM HISDB.dbo.MBM m
    INNER JOIN CISDB.dbo.CISCas c ON m.MBMNumber = c.MBMnumber
    INNER JOIN CISWorksheet.dbo.Claim cl ON cl.CASNumber = c.CASNumber
    LEFT OUTER JOIN CISPayment.dbo.Payment2CLN p ON cl.CASNumber = p.CASNumber
    LEFT OUTER JOIN CISWorksheet.dbo.Bordx b ON cl.BordxRefNo = b.BordxRefNo
    WHERE c.ClinicCode = @clinic AND c.InvoiceNo IS NOT NULL AND c.InvoiceNo <> N'' AND c.CASINVYear >= @yr AND c.CASINVYear <= @yr
      AND p.PVCLNNo IS NULL AND SUBSTRING(m.MBMNumber, 1, 2) <> 'EM' AND SUBSTRING(m.MBMNumber, 1, 2) <> 'ES'
      AND ',' + REPLACE(@months, ' ', '') + ',' LIKE '%,' + RTRIM(c.CASINVMonth) + ',%'
      AND ((@entity = 0 AND SUBSTRING(c.CASNumber, 1, 2) <> 'SJ') OR (@entity = 1 AND SUBSTRING(c.CASNumber, 1, 2) = 'SJ'));
END
GO

-- Payment2CLNListing's per-row CLN look-up (name, payee and its address, bank, type, FFS, reg. no., MCO %).
CREATE OR ALTER PROCEDURE cis.PayClnInv_Clinics @codes nvarchar(max) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(CLNCode) AS CLNCode, CLNName, CLNPName, CLNPAdd1, CLNPAdd2, CLNPAdd3, CLNBankAcc, CLNBankName, CLNType, CLNFFSStatus, CLNRegNo, CLNMCOProcfeePerc,
        ClinicGrpCode
    FROM cis.Hdb_CLN WHERE ',' + REPLACE(@codes, ' ', '') + ',' LIKE '%,' + RTRIM(CLNCode) + ',%';
END
GO

-- The Payment2CLN rows of a PV (CmdUpdate: status F only).
CREATE OR ALTER PROCEDURE cis.PayClnInv_PvRows @pvNo nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(PVCLNIndex AS nvarchar(20)) AS PVCLNIndex, RTRIM(CASNumber) AS CASNumber, PVPayCLNStatus, CLNPaidAmt
    FROM CISPayment.dbo.Payment2CLN WHERE PVCLNNo = @pvNo AND PVPayCLNStatus = 'F';
END
GO

-- The Payment2CLN insert / overwrite of frmPaymentCLNInvoice UpdateRecord (with the amount wording).
CREATE OR ALTER PROCEDURE cis.PayClnInv_Save @index nvarchar(20) = NULL, @casNumber nvarchar(30), @pvNo nvarchar(30), @pvDate datetime, @chequeDate datetime = NULL,
    @chequeNo nvarchar(30), @chequeAmt money, @bank nvarchar(10), @remark nvarchar(500) = NULL, @paidAmt money, @mth nvarchar(2), @yr nvarchar(4), @status char(1),
    @wordingOne nvarchar(500), @wordingTwo nvarchar(500), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF @index IS NOT NULL AND EXISTS (SELECT 1 FROM CISPayment.dbo.Payment2CLN WHERE CAST(PVCLNIndex AS nvarchar(20)) = @index)
        UPDATE CISPayment.dbo.Payment2CLN SET CASNumber = @casNumber, PVCLNNo = @pvNo, PVCLNDate = @pvDate, PVCLNChequeDate = @chequeDate, PVCLNChequeNo = @chequeNo,
            PVCLNChequeAmt = @chequeAmt, PVCLNBank = @bank, CLNRemark = COALESCE(@remark, ''), CLNPaidAmt = @paidAmt, CLNAccPeriodMth = @mth, CLNAccPeriodYr = @yr,
            Pay2CLNLastUpdateDate = CAST(GETDATE() AS date), Pay2CLNLastUpdateUser = @user, PVPayCLNStatus = @status, PVWordingOne = @wordingOne, PVWordingTwo = @wordingTwo
        WHERE CAST(PVCLNIndex AS nvarchar(20)) = @index;
    ELSE
        INSERT CISPayment.dbo.Payment2CLN (CASNumber, PVCLNNo, PVCLNDate, PVCLNChequeDate, PVCLNChequeNo, PVCLNChequeAmt, PVCLNBank, CLNRemark, CLNPaidAmt,
            CLNAccPeriodMth, CLNAccPeriodYr, Pay2CLNLastUpdateDate, Pay2CLNLastUpdateUser, PVPayCLNStatus, PVWordingOne, PVWordingTwo)
        VALUES (@casNumber, @pvNo, @pvDate, @chequeDate, @chequeNo, @chequeAmt, @bank, COALESCE(@remark, ''), @paidAmt, @mth, @yr, CAST(GETDATE() AS date), @user,
            @status, @wordingOne, @wordingTwo);
END
GO

-- cmdPrintChq / frmPrintChequeRange: the cheques (VIEW_Payment2CLN_Chq) and their print status.
CREATE OR ALTER PROCEDURE cis.PayClnInv_Cheques @chequeFrom nvarchar(30), @chequeTo nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT PVCLNChequeNo, PVCLNChequeDate, PVCLNChequeAmt, PVWordingOne, PVWordingTwo, CLNPName
    FROM CISPayment.dbo.VIEW_Payment2CLN_Chq WHERE PVCLNChequeNo >= @chequeFrom AND PVCLNChequeNo <= @chequeTo ORDER BY PVCLNChequeNo;
END
GO

CREATE OR ALTER PROCEDURE cis.PayClnInv_ChequesPrinted @chequeFrom nvarchar(30), @chequeTo nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISPayment.dbo.Payment2CLN SET PVChequeStatus = 'Y' WHERE PVCLNChequeNo >= @chequeFrom AND PVCLNChequeNo <= @chequeTo;
END
GO
