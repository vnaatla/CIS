-- Invoice > Other > OP Batch/Invoice (frmUpdateInvoiceBatch): list open cases (the desktop's ClaimBordxModule case 27, parameterised),
-- give the checked cases an OP batch number (GenerateOPSeq "OC") and / or an invoice number and date (CISCAS).

-- SearchRecord → InvoiceListing. Every filter optional; flags: 1 = the desktop's check box / combo choice.
CREATE OR ALTER PROCEDURE cis.OpBatch_Search
    @noBatch bit = 0, @noBordx bit = 0, @bordxFrom nvarchar(30) = NULL, @bordxTo nvarchar(30) = NULL, @noBtbgBordx bit = 0,
    @btbgFrom nvarchar(30) = NULL, @btbgTo nvarchar(30) = NULL, @casFrom nvarchar(20) = NULL, @casTo nvarchar(20) = NULL,
    @clinicCode nvarchar(10) = NULL, @groupCompany nvarchar(200) = NULL, @invoiceNo nvarchar(100) = NULL, @visitFrom datetime = NULL,
    @visitTo datetime = NULL, @noInvoice bit = 0, @invoiceFrom datetime = NULL, @invoiceTo datetime = NULL, @amountFrom money = NULL,
    @amountTo money = NULL, @batchFrom nvarchar(30) = NULL, @batchTo nvarchar(30) = NULL, @omni bit = 0, @regFrom datetime = NULL,
    @regTo datetime = NULL, @clinicType char(1) = NULL, @ffs char(1) = NULL, @hasBatch char(1) = NULL, @payTo nvarchar(4) = NULL,
    @referral char(1) = NULL, @entity int = 0
AS
BEGIN
    SET NOCOUNT ON;
    SELECT c.CASNumber, c.MBMnumber AS MBMNumber, c.CoverID,
           CASE c.CoverID WHEN '00' THEN m.MBMName ELSE cp.MBMCName END AS PatientName,
           c.DateVisit, o.MBMGroupCompany, ci.INVAmount, cl.CLMMCOProcFee,
           CASE WHEN p.PLNBTBG = 1 THEN COALESCE(cl.CLMFinalApprove, 0) - COALESCE(cl.CLMBTBGAmt, 0) ELSE cl.CLMFinalApprove END AS CLMFinalApprove,
           c.InvoiceNo AS BillNo, o.MBMDepartment, cln.CLNName, c.InvoiceDate AS BillDate,
           CASE WHEN LEN(ci.INVEnteredBy) = 0 THEN CONVERT(varchar(10), r.RecDate, 111) ELSE CONVERT(varchar(10), ci.INVEnteredDate, 111) END AS INVEnteredDate,
           CASE WHEN LEN(ci.INVEnteredBy) = 0 THEN r.Incharge ELSE ci.INVEnteredBy END AS INVEnteredBy,
           c.OPBatchNo,
           CASE WHEN ci.INVPayTo = 'M' THEN pm.PVNo ELSE pc.PVCLNNo END AS PNO,
           CASE WHEN ci.INVPayTo = 'M' THEN pm.PVDate ELSE pc.PVCLNDate END AS PDATE,
           CASE WHEN ci.INVPayTo = 'M' THEN pm.PVChequeNo ELSE pc.PVCLNChequeNo END AS PCNO,
           dbo.WebGetTreatmentDesc(c.CASTreatReason1) AS ReasonDesc1, dbo.WebGetTreatmentDesc(c.CASTreatReason2) AS ReasonDesc2,
           dbo.WebGetTreatmentDesc(c.CASTreatReason3) AS ReasonDesc3, dbo.WebGetTreatmentDesc(c.CASTreatReason4) AS ReasonDesc4,
           dbo.WebGetTreatmentDesc(c.CASTreatReason5) AS ReasonDesc5,
           c.CASDiagnosis5, c.MCDays, cl.CLMCoPayCLNColl,
           (CASE WHEN p.PLNBTBG = 1 THEN COALESCE(cl.CLMFinalApprove, 0) - COALESCE(cl.CLMBTBGAmt, 0) ELSE COALESCE(cl.CLMFinalApprove, 0) END)
               - COALESCE(ci.INVAmount, 0) AS CLMFinalApprove1,
           CASE c.CoverID WHEN '00' THEN m.MBMIcBcPp ELSE cp.MBMCICBCPPNo END AS ICNo,
           CASE WHEN LEN(c.OPBatchNo) = 0 THEN 'NO' ELSE 'YES' END AS OPNO,
           CASE WHEN ci.INVPayTo = 'M' THEN 'Reimbursement' ELSE 'Cash Less' END AS PType,
           c.OPBatchDate, c.LastUpdateDate, COALESCE(c.Pending, 'N') AS Pending, c.Claimability
    FROM dbo.CISCAS c WITH (NOLOCK)
    INNER JOIN cis.Hdb_MBMCLNOthers co WITH (NOLOCK) ON co.MBMNumber = c.MBMnumber
    INNER JOIN cis.Hdb_MBMOthers o WITH (NOLOCK) ON o.MBMNumber = co.MBMNumber
    INNER JOIN cis.Hdb_MBM m WITH (NOLOCK) ON m.MBMNumber = o.MBMNumber
    INNER JOIN cis.Ws_Claim cl WITH (NOLOCK) ON cl.CASNumber = c.CASNumber
    LEFT JOIN cis.Hdb_MBMCoveredPersons cp WITH (NOLOCK) ON cp.MBMCNumber = c.MBMnumber AND cp.MBMCCoverID = c.CoverID
    LEFT JOIN cis.Hdb_CLN cln WITH (NOLOCK) ON c.ClinicCode = cln.CLNCode
    LEFT JOIN dbo.GRPCompany g WITH (NOLOCK) ON g.GRPCoName = o.MBMGroupCompany AND (g.HISpaycode IS NULL OR g.HISpaycode = m.PAYCode)
    LEFT JOIN cis.Ws_ClaimInvoice ci WITH (NOLOCK) ON ci.CASNumber = cl.CASNumber
    LEFT JOIN dbo.tblRecord r WITH (NOLOCK) ON r.CaseNo = cl.CASNumber
    LEFT JOIN cis.Pay_Payment2CLN pc WITH (NOLOCK) ON pc.CASNumber = cl.CASNumber
    LEFT JOIN cis.Pay_Payment2MBM pm WITH (NOLOCK) ON pm.CASNumber = cl.CASNumber
    INNER JOIN dbo.PLN p WITH (NOLOCK) ON p.PLNCode = co.PLNType
    WHERE co.PAYCode NOT IN ('DZ', 'EM', 'ES') AND c.CASStatus = 'O' AND o.MBMGroupCompany <> 'DEMO COMPANY BERHAD'
      AND (@noBatch = 0 OR COALESCE(c.OPBatchNo, '') = '')
      AND (@noBordx = 0 OR COALESCE(cl.BordxRefNo, '') = '')
      AND (@bordxFrom IS NULL OR cl.BordxRefNo >= @bordxFrom) AND (@bordxTo IS NULL OR cl.BordxRefNo <= @bordxTo)
      AND (@noBtbgBordx = 0 OR COALESCE(cl.BTBGBordxRefNo, '') = '')
      AND (@btbgFrom IS NULL OR cl.BTBGBordxRefNo >= @btbgFrom) AND (@btbgTo IS NULL OR cl.BTBGBordxRefNo <= @btbgTo)
      AND (@casFrom IS NULL OR c.CASNumber >= @casFrom) AND (@casTo IS NULL OR c.CASNumber <= @casTo)
      AND (@clinicCode IS NULL OR c.ClinicCode = @clinicCode) AND (@groupCompany IS NULL OR o.MBMGroupCompany = @groupCompany)
      AND (@invoiceNo IS NULL OR c.InvoiceNo = @invoiceNo)
      AND (@visitFrom IS NULL OR c.DateVisit >= @visitFrom) AND (@visitTo IS NULL OR c.DateVisit <= @visitTo)
      AND (@noInvoice = 0 OR COALESCE(c.InvoiceNo, '') = '')
      AND (@invoiceFrom IS NULL OR c.InvoiceDate >= @invoiceFrom) AND (@invoiceTo IS NULL OR c.InvoiceDate <= @invoiceTo)
      AND (@amountFrom IS NULL OR ci.INVAmount >= @amountFrom) AND (@amountTo IS NULL OR ci.INVAmount <= @amountTo)
      AND (@batchFrom IS NULL OR c.OPBatchNo >= @batchFrom) AND (@batchTo IS NULL OR c.OPBatchNo <= @batchTo)
      AND (@omni = 0 OR g.GRPOMNICLM = 'Y')
      AND (@regFrom IS NULL OR c.RegisterDate >= @regFrom) AND (@regTo IS NULL OR c.RegisterDate <= @regTo)
      AND (@clinicType IS NULL OR (@clinicType = '0' AND c.CLNType = 'S') OR (@clinicType = '1' AND c.CLNType = 'G')
           OR (@clinicType = '2' AND COALESCE(c.CLNType, '') = ''))
      AND (@ffs IS NULL OR (@ffs = 'Y' AND cln.CLNFFSStatus = 'Y') OR (@ffs = 'N' AND (cln.CLNFFSStatus = 'N' OR COALESCE(cln.CLNFFSStatus, '') = '')))
      AND (@hasBatch IS NULL OR (@hasBatch = 'Y' AND COALESCE(c.OPBatchNo, '') <> '') OR (@hasBatch = 'N' AND COALESCE(c.OPBatchNo, '') = ''))
      AND (@payTo IS NULL OR ci.INVPayTo = @payTo)
      AND (@referral IS NULL OR (@referral = 'Y' AND cl.CLMRefLetter = 'Y') OR (@referral = 'N' AND (cl.CLMRefLetter = 'N' OR COALESCE(cl.CLMRefLetter, '') = '')))
      AND ((@entity = 0 AND SUBSTRING(c.CASNumber, 1, 2) <> 'SJ') OR (@entity = 1 AND SUBSTRING(c.CASNumber, 1, 2) = 'SJ') OR @entity NOT IN (0, 1));
END
GO

-- The invoice number is already used (CISCAS.InvoiceNo).
CREATE OR ALTER PROCEDURE cis.OpBatch_InvoiceUsed @invoiceNo nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.CISCAS WHERE InvoiceNo = @invoiceNo) THEN 1 ELSE 0 END AS bit);
END
GO

-- GenerateOPSeq("OC"): next OPBatch sequence of the code (letter + 7 digits; the letter moves on after 9999999), locked against other users.
CREATE OR ALTER PROCEDURE cis.OpBatch_Next @opCode nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seq nvarchar(20);
    SELECT @seq = RTRIM(OPBatchSeq) FROM dbo.OPBatch WITH (UPDLOCK, HOLDLOCK) WHERE OPCode = @opCode;
    IF @seq IS NULL
    BEGIN
        SET @seq = 'A0000001';
        INSERT INTO dbo.OPBatch (OPCode, OPBatchSeq) VALUES (@opCode, @seq);
    END
    ELSE
    BEGIN
        DECLARE @letter char(1) = LEFT(@seq, 1), @no int = TRY_CAST(SUBSTRING(@seq, 2, 8) AS int);
        IF @letter BETWEEN 'A' AND 'Y'
            SELECT @seq = CASE WHEN @no = 9999999 THEN CHAR(ASCII(@letter) + 1) + '0000001' ELSE @letter + RIGHT('0000000' + CAST(@no + 1 AS varchar(8)), 7) END;
        ELSE
            SET @seq = 'A' + RIGHT('0000000' + CAST(@no AS varchar(8)), 7);
        UPDATE dbo.OPBatch SET OPBatchSeq = @seq WHERE OPCode = @opCode;
    END
    SELECT @opCode + @seq;
END
GO

-- UpdateRecord: one checked case. @batchNo set = the case had no OP batch; @invoiceNo set = invoice number and date (with the month / year).
CREATE OR ALTER PROCEDURE cis.OpBatch_Save
    @casNumber nvarchar(20), @batchNo nvarchar(30) = NULL, @userCode nvarchar(20), @today date, @invoiceNo nvarchar(100) = NULL, @invoiceDate datetime = NULL
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET
        OPBatchNo = COALESCE(@batchNo, OPBatchNo),
        OPNoEnterUser = CASE WHEN @batchNo IS NULL THEN OPNoEnterUser ELSE @userCode END,
        OPBatchDate = CASE WHEN @batchNo IS NULL THEN OPBatchDate ELSE @today END,
        InvoiceNo = COALESCE(@invoiceNo, InvoiceNo),
        InvoiceDate = CASE WHEN @invoiceNo IS NULL THEN InvoiceDate ELSE @invoiceDate END,
        CASInvMonth = CASE WHEN @invoiceNo IS NULL THEN CASInvMonth ELSE RIGHT('0' + CAST(MONTH(@invoiceDate) AS varchar(2)), 2) END,
        CASInvYear = CASE WHEN @invoiceNo IS NULL THEN CASInvYear ELSE CAST(YEAR(@invoiceDate) AS varchar(4)) END
    WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO
