-- Payment > To Clinic (frmPaymentClinic): the payment list (CISWorksheet ClaimBordxModule case 26 copied, the pasted criteria as parameters),
-- GenerateSPVNo (PVSeq2013) and UpdateRecord (Payment2CLN).
SET QUOTED_IDENTIFIER ON;
GO

-- @pvNo set: the list of one PV (after a save / update); otherwise the search. @payment 1 paid (F), 2 unpaid, 3 partial, 4 partial or paid;
-- @blank*: the "empty" check boxes; @entity 0 Health Connect, 1 Sihat Johor.
CREATE OR ALTER PROCEDURE cis.PayClinic_Search @pvNo nvarchar(30) = NULL, @company nvarchar(200) = NULL, @clinic nvarchar(20) = NULL, @payment int = 0,
    @scan bit = 0, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL, @invoiceNo nvarchar(50) = NULL, @blankInvoiceNo bit = 0,
    @invFrom datetime = NULL, @invTo datetime = NULL, @blankBatch bit = 0, @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @blankInvDate bit = 0,
    @amtFrom nvarchar(20) = NULL, @amtTo nvarchar(20) = NULL, @blankAmt bit = 0, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL, @blankPv bit = 0,
    @pvDateFrom datetime = NULL, @pvDateTo datetime = NULL, @blankPvDate bit = 0, @mthFrom nvarchar(2) = NULL, @mthTo nvarchar(2) = NULL, @blankMth bit = 0,
    @yrFrom nvarchar(4) = NULL, @yrTo nvarchar(4) = NULL, @blankYr bit = 0, @sunFrom nvarchar(30) = NULL, @sunTo nvarchar(30) = NULL, @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = CASE WHEN @pvNo IS NOT NULL THEN N' Where 0=0 and Payment2CLN.PVCLNNo = @pvNo' ELSE N' Where ClaimInvoice.INVPayTo = ''C'' AND CISCas.CASStatus=''O'' '
        + CASE WHEN @company IS NOT NULL THEN N' AND MBMOthers.MBMGroupCompany = @company' ELSE N'' END
        + CASE WHEN @clinic IS NOT NULL THEN N' AND CISCas.ClinicCode = @clinic' ELSE N'' END
        + CASE @payment WHEN 1 THEN N' and Payment2CLN.PVPayCLNStatus = ''F''' WHEN 2 THEN N' and COALESCE(Payment2CLN.PVPayCLNStatus , '''')='''' '
            WHEN 3 THEN N' and Payment2CLN.PVPayCLNStatus = ''P'' ' WHEN 4 THEN N' and (Payment2CLN.PVPayCLNStatus = ''P'' or Payment2CLN.PVPayCLNStatus = ''F'') ' ELSE N'' END
        + CASE WHEN @scan = 1 THEN N' AND CISCas.CASNumber = @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseFrom IS NOT NULL THEN N' AND CISCas.CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseTo IS NOT NULL THEN N' AND CISCas.CASNumber <= @caseTo' ELSE N'' END
        + CASE WHEN @invoiceNo IS NOT NULL THEN N' and CISCas.InvoiceNo = @invoiceNo' ELSE N'' END
        + CASE WHEN @blankInvoiceNo = 1 THEN N' AND COALESCE(CISCas.InvoiceNo, '''')='''' ' ELSE N'' END
        + CASE WHEN @invFrom IS NOT NULL THEN N' And CISCas.InvoiceDate >= @invFrom' ELSE N'' END
        + CASE WHEN @invTo IS NOT NULL THEN N' And CISCas.InvoiceDate <= @invTo' ELSE N'' END
        + CASE WHEN @blankBatch = 1 THEN N' AND COALESCE(CISCas.OPBatchNo, '''')='''' ' ELSE N'' END
        + CASE WHEN @opFrom IS NOT NULL THEN N' AND CISCas.OPBatchNo >= @opFrom' ELSE N'' END
        + CASE WHEN @opTo IS NOT NULL THEN N' AND CISCas.OPBatchNo <= @opTo' ELSE N'' END
        + CASE WHEN @blankInvDate = 1 THEN N' AND COALESCE(CISCas.InvoiceDate, '''')='''' ' ELSE N'' END
        + CASE WHEN @amtFrom IS NOT NULL THEN N' and CISCas.InvoiceStmtAmt >= @amtFrom' ELSE N'' END
        + CASE WHEN @amtTo IS NOT NULL THEN N' and CISCas.InvoiceStmtAmt <= @amtTo' ELSE N'' END
        + CASE WHEN @blankAmt = 1 THEN N' AND COALESCE(CISCas.InvoiceStmtAmt , '''')='''' ' ELSE N'' END
        + CASE WHEN @pvFrom IS NOT NULL THEN N' and Payment2CLN.PVCLNNo >= @pvFrom' ELSE N'' END
        + CASE WHEN @pvTo IS NOT NULL THEN N' and Payment2CLN.PVCLNNo <= @pvTo' ELSE N'' END
        + CASE WHEN @blankPv = 1 THEN N' AND COALESCE(Payment2CLN.PVCLNNo ,'''')='''' ' ELSE N'' END
        + CASE WHEN @pvDateFrom IS NOT NULL THEN N' And Payment2CLN.PVCLNDate >= @pvDateFrom' ELSE N'' END
        + CASE WHEN @pvDateTo IS NOT NULL THEN N' And Payment2CLN.PVCLNDate <= @pvDateTo' ELSE N'' END
        + CASE WHEN @blankPvDate = 1 THEN N' AND COALESCE(Payment2CLN.PVCLNDate, '''')='''' ' ELSE N'' END
        + CASE WHEN @mthFrom IS NOT NULL THEN N' and Payment2CLN.CLNAccPeriodMth >= @mthFrom' ELSE N'' END
        + CASE WHEN @mthTo IS NOT NULL THEN N' and Payment2CLN.CLNAccPeriodMth <= @mthTo' ELSE N'' END
        + CASE WHEN @blankMth = 1 THEN N' AND COALESCE(Payment2CLN.CLNAccPeriodMth,'''')='''' ' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL AND @yrFrom IS NOT NULL THEN N' and Payment2CLN.CLNAccPeriodYr >= @yrFrom' ELSE N'' END
        + CASE WHEN @yrTo IS NOT NULL THEN N' and Payment2CLN.CLNAccPeriodYr <= @yrTo' ELSE N'' END
        + CASE WHEN @blankYr = 1 THEN N' AND COALESCE(Payment2CLN.CLNAccPeriodYr,'''')='''' ' ELSE N'' END
        + CASE WHEN @sunFrom IS NOT NULL THEN N' and SUNPosting.CLNInvoiceBatch >= @sunFrom' ELSE N'' END
        + CASE WHEN @sunTo IS NOT NULL THEN N' and SUNPosting.CLNInvoiceBatch <= @sunTo' ELSE N'' END
        + CASE @entity WHEN 0 THEN N' AND Substring(CISCas.CasNumber, 1, 2) <>''SJ'' ' WHEN 1 THEN N' AND Substring(CISCas.CasNumber, 1, 2) =''SJ'' ' ELSE N'' END
        END;

			       declare @sql nvarchar(max)
				   declare @groupBy nvarchar(max)
				   
						set @sql =  N'   SELECT  CISCas.CASNumber, CISCas.MBMnumber, CISCas.CoverID, CISCas.DateVisit, CISCas.CASStatus, CISCas.Claimability, CISCas.Pending, CISCas.ClinicCode, CISCas.Remarks, 
												Payment2CLN.PVCLNNo, Payment2CLN.PVCLNDate, Payment2CLN.PVCLNIndex, Payment2CLN.PVCLNChequeDate, Payment2CLN.PVCLNChequeNo, Payment2CLN.PVPayCLNStatus, Claim.CLMPayableByMedix2C, 
												Payment2CLN.CLNPaidAmt, Claim.BordxRefNo, Payment2CLN.CLNAccPeriodMth, Payment2CLN.CLNAccPeriodYr, CLN.CLNName, ClaimInvoice.INVPayTo, ClaimInvoice.INVEnteredBy, ClaimInvoice.BillNo, CLN.CLNAdd1, 
												CLN.CLNAdd2, CLN.CLNAdd3, ClaimInvoice.BillDate, CLN.CLNState, CLN.CLNCity, CLN.CLNPostCode, Bordx.GroupCompany, CISCas.InvoiceNo, CISCas.InvoiceDate, Payment2CLN.PVCLNBank, Payment2CLN.PVCLNChequeAmt, 
												Payment2CLN.PVCLNAccount, Payment2CLN.CLNRemark, CISCas.OPBatchNo, CISCas.InvoiceStmtAmt, CLN.ClinicGrpCode, Claim.CLMNonGPFee, Claim.CLMTotalActual, Claim.CLMMCOProcFee,MBMOthers.MBMGroupCompany, GRPCompany.GRPCoCode, 
												SUNPosting.PostDate AS SUNPostDate, SUNPosting.CLNInvoiceBatch AS SUNInvNo, Claim.CLMCoPayCLNColl, Claim.CLMBordxAmt, Claim.CLMFinalApprove
												,CASE CISCAS.CoverID WHEN ''00'' THEN mbm.MBMName ELSE MBMCoveredPersons.MBMCName END AS PatientName
												FROM CISDB..CISCas     
												INNER JOIN  CISWorksheet..Claim ON CISCas.CASNumber = Claim.CASNumber
												INNER JOIN CISWorksheet..ClaimInvoice  ON ClaimInvoice.CASNumber = Claim.CASNumber 
												INNER JOIN HISDB..CLN  ON CLN.CLNCode = CISCas.ClinicCode 
												INNER JOIN HISDB..MBMOthers ON CISCas.MBMnumber = MBMOthers.MBMNumber 
												INner JOin HISDB..mbm ON mbm.mbmnumber=CISCas.MBMNumber
												LEFT OUTER JOIN HISDB..MBMCoveredPersons ON MBMCoveredPersons.MBMCNumber = CISCas.MBMnumber AND MBMCoveredPersons.MBMCCoverID = CISCas.CoverID 
												LEFT JOIN CISWorksheet..SUNPosting ON CISCas.CASNumber = SUNPosting.InvoiceNo AND CISCas.ClinicCode = SUNPosting.TransRef 
												LEFT JOIN CISDB..GRPCompany ON MBMOthers.MBMGroupCompany = GRPCompany.GRPCoName 
												LEFT JOIN cispayment..Payment2CLN  ON Payment2CLN.CASNumber = Claim.CASNumber
												LEFT JOIN CISWorksheet..Bordx ON Bordx.BordxRefNo = Claim.BordxRefNo   '
		           
			       Set @groupBy = ' GROUP BY CISCas.CASNumber, CISCas.MBMnumber, CISCas.CoverID, CISCas.DateVisit, CISCas.CASStatus, CISCas.Claimability, CISCas.Pending, CISCas.ClinicCode, CISCas.Remarks, 
												Payment2CLN.PVCLNNo, Payment2CLN.PVCLNDate, Payment2CLN.PVCLNIndex, Payment2CLN.PVCLNChequeDate, Payment2CLN.PVCLNChequeNo, Payment2CLN.PVPayCLNStatus, Claim.CLMPayableByMedix2C, 
												Payment2CLN.CLNPaidAmt, Claim.BordxRefNo, Payment2CLN.CLNAccPeriodMth, Payment2CLN.CLNAccPeriodYr, CLN.CLNName, ClaimInvoice.INVPayTo, ClaimInvoice.INVEnteredBy, ClaimInvoice.BillNo,
												CLN.CLNAdd1, CLN.CLNAdd2, CLN.CLNAdd3, ClaimInvoice.BillDate, CLN.CLNState, CLN.CLNCity, CLN.CLNPostCode, Bordx.GroupCompany, CISCas.InvoiceNo, CISCas.InvoiceDate, Payment2CLN.PVCLNBank, 
												Payment2CLN.PVCLNChequeAmt, Payment2CLN.PVCLNAccount, Payment2CLN.CLNRemark, CISCas.OPBatchNo, CISCas.InvoiceStmtAmt, CLN.ClinicGrpCode, Claim.CLMNonGPFee, Claim.CLMTotalActual, Claim.CLMMCOProcFee, MBMOthers.MBMGroupCompany, 
												GRPCompany.GRPCoCode, SUNPosting.PostDate, SUNPosting.CLNInvoiceBatch, Claim.CLMCoPayCLNColl, Claim.CLMBordxAmt, Claim.CLMFinalApprove,mbm.MBMName,MBMCoveredPersons.MBMCName '
				   
		set @sql = @sql + @filter + @groupBy
				   			   
		exec CISWorksheet.sys.sp_executesql @sql, N'@company nvarchar(200), @clinic nvarchar(20), @caseFrom nvarchar(30), @caseTo nvarchar(30), @invoiceNo nvarchar(50),
            @invFrom datetime, @invTo datetime, @opFrom nvarchar(30), @opTo nvarchar(30), @amtFrom nvarchar(20), @amtTo nvarchar(20), @pvFrom nvarchar(30), @pvTo nvarchar(30),
            @pvDateFrom datetime, @pvDateTo datetime, @mthFrom nvarchar(2), @mthTo nvarchar(2), @yrFrom nvarchar(4), @yrTo nvarchar(4), @sunFrom nvarchar(30), @sunTo nvarchar(30),
            @pvNo nvarchar(30)',
            @company, @clinic, @caseFrom, @caseTo, @invoiceNo, @invFrom, @invTo, @opFrom, @opTo, @amtFrom, @amtTo, @pvFrom, @pvTo, @pvDateFrom, @pvDateTo,
            @mthFrom, @mthTo, @yrFrom, @yrTo, @sunFrom, @sunTo, @pvNo;
END
GO

CREATE OR ALTER PROCEDURE cis.PayClinic_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(BankCode) AS Code, RTRIM(BankDesc) AS Name, RTRIM(PVCode) AS PvCode FROM cis.Maint_Banks;
    SELECT DISTINCT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL;
    SELECT RTRIM(CLNCode) AS Code, RTRIM(CLNName) AS Name FROM cis.Hdb_CLN WHERE LEN(CLNName) > 0;
END
GO

-- GenerateSPVNo: PV code + the next 9-digit number of PVSeq2013 for the bank, locked for the save.
CREATE OR ALTER PROCEDURE cis.PayClinic_NextPv @bank nvarchar(10), @pvCode nvarchar(10), @pvNo nvarchar(30) OUTPUT AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @n char(10);
    IF NOT EXISTS (SELECT 1 FROM CISPayment.dbo.PVSeq2013 WITH (UPDLOCK, HOLDLOCK) WHERE BankCode = @bank)
    BEGIN
        INSERT CISPayment.dbo.PVSeq2013 (BankCode, PVCode, PVSEQNO) VALUES (@bank, @pvCode, '100000001');
        SET @n = '100000001';
    END
    ELSE
        UPDATE TOP (1) CISPayment.dbo.PVSeq2013 SET @n = PVSEQNO = RIGHT('000000000' + CAST(CAST(PVSEQNO AS bigint) + 1 AS varchar(12)), 9) WHERE BankCode = @bank;
    SET @pvNo = RTRIM(@pvCode) + RTRIM(@n);
END
GO

-- CmdPrintPV (@pvNo) / PreviewPV (@cases, comma list): View_Payment2CLN rows with the receipt no. and the HQ clinic (ClinicHQInfo, preview only).
CREATE OR ALTER PROCEDURE cis.PayClinic_PvRows @pvNo nvarchar(30) = NULL, @cases nvarchar(max) = NULL AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @c TABLE (CASNumber nvarchar(30) PRIMARY KEY);
    IF @pvNo IS NOT NULL
        INSERT @c SELECT DISTINCT RTRIM(CASNumber) FROM CISPayment.dbo.Payment2CLN WHERE PVCLNNo = @pvNo;
    ELSE
    BEGIN
        DECLARE @x xml = CAST('<c>' + REPLACE(REPLACE(REPLACE(REPLACE(@cases, '&', ''), '<', ''), ' ', ''), ',', '</c><c>') + '</c>' AS xml);
        INSERT @c SELECT DISTINCT t.c.value('.', 'nvarchar(30)') FROM @x.nodes('/c') t(c) WHERE t.c.value('.', 'nvarchar(30)') <> '';
    END
    SELECT v.PVCLNIndex, v.CASNumber, v.BordxRefNo, v.BillNo, v.BillDate, v.PVCLNNo, v.PVCLNChequeNo, v.PVCLNDate, v.CLMTotalActual, v.CLMMCOProcFee,
        v.ClinicGrpCode, v.ClinicCode, v.CLNName, v.CLNAdd1, v.CLNAdd2, v.CLNAdd3, v.CLNPostCode, v.CLNCity, v.CLNState, v.CLNPaidAmt, v.CLMPayableByMedix2C,
        (SELECT TOP 1 r.ReceiptNo FROM CISPayment.dbo.REceiptFrPAY r WHERE r.CASNumber = v.CASNumber) AS ReceiptNo,
        hq.CLNCode AS HqCode, hq.CLNName AS HqName, hq.CLNAdd1 AS HqAdd1, hq.CLNAdd2 AS HqAdd2, hq.CLNAdd3 AS HqAdd3, hq.CLNPostCode AS HqPostCode,
        hq.CLNCity AS HqCity, hq.CLNState AS HqState, v.InvoiceNo, v.InvoiceDate, v.CLMFinalApprove, v.PVCLNChequeAmt, v.CLNAccPeriodMth, v.CLNAccPeriodYr,
        p.CLNPName, p.CLNPAdd1, p.CLNPAdd2, p.CLNPAdd3
    FROM CISPayment.dbo.View_Payment2CLN v
    LEFT JOIN cis.Hdb_CLN p ON p.CLNCode = v.ClinicCode
    LEFT JOIN cis.Hdb_CLN hq ON @pvNo IS NULL AND COALESCE(RTRIM(v.ClinicGrpCode), '') <> '' AND hq.CLNCode = RTRIM(v.ClinicGrpCode)
    WHERE v.CASNumber IN (SELECT CASNumber FROM @c) AND (@pvNo IS NULL OR v.PVCLNNo = @pvNo);
END
GO

-- UpdateRecord: the Payment2CLN row of @index is overwritten (edit, or not partial), or a new row is added (no row yet, or a partial payment's next instalment).
CREATE OR ALTER PROCEDURE cis.PayClinic_Save @index nvarchar(20) = NULL, @edit bit = 0, @casNumber nvarchar(30), @pvNo nvarchar(30), @pvDate datetime,
    @chequeDate datetime = NULL, @chequeNo nvarchar(30) = NULL, @chequeAmt money, @bank nvarchar(10), @account nvarchar(50) = NULL, @remark nvarchar(500) = NULL,
    @paidAmt money, @mth nvarchar(2), @yr nvarchar(4), @status char(1), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @old char(1), @found bit = 0;
    SELECT TOP 1 @old = RTRIM(PVPayCLNStatus), @found = 1 FROM CISPayment.dbo.Payment2CLN WHERE CAST(PVCLNIndex AS nvarchar(20)) = @index;
    IF @found = 1 AND @edit = 0 AND @old = 'P'
    BEGIN
        UPDATE CISPayment.dbo.Payment2CLN SET PVPayCLNStatus = @status WHERE CAST(PVCLNIndex AS nvarchar(20)) = @index;
        SET @found = 0;
    END
    IF @found = 1
        UPDATE CISPayment.dbo.Payment2CLN SET CASNumber = @casNumber, PVCLNNo = @pvNo, PVCLNDate = @pvDate, PVCLNChequeDate = @chequeDate,
            PVCLNChequeNo = COALESCE(@chequeNo, ''), PVCLNChequeAmt = @chequeAmt, PVCLNBank = @bank, PVCLNAccount = COALESCE(@account, ''),
            CLNRemark = COALESCE(@remark, ''), CLNPaidAmt = @paidAmt, CLNAccPeriodMth = @mth, CLNAccPeriodYr = @yr, Pay2CLNLastUpdateDate = CAST(GETDATE() AS date),
            Pay2CLNLastUpdateUser = @user, PVPayCLNStatus = @status
        WHERE CAST(PVCLNIndex AS nvarchar(20)) = @index;
    ELSE
        INSERT CISPayment.dbo.Payment2CLN (CASNumber, PVCLNNo, PVCLNDate, PVCLNChequeDate, PVCLNChequeNo, PVCLNChequeAmt, PVCLNBank, PVCLNAccount, CLNRemark,
            CLNPaidAmt, CLNAccPeriodMth, CLNAccPeriodYr, Pay2CLNLastUpdateDate, Pay2CLNLastUpdateUser, PVPayCLNStatus)
        VALUES (@casNumber, @pvNo, @pvDate, @chequeDate, COALESCE(@chequeNo, ''), @chequeAmt, @bank, COALESCE(@account, ''), COALESCE(@remark, ''), @paidAmt, @mth, @yr,
            CAST(GETDATE() AS date), @user, @status);
END
GO
