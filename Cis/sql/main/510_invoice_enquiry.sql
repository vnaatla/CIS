-- Invoice > Other > Enquiry (FrmInvoiceEnq): one case read-only — the case (SP_INV_CLAIM), its stored worksheet, bordereaux, payments,
-- receipts and medication requests.

-- InvoiceINFO: SP_INV_CLAIM (already parameterised) and the GP / SP fee parts.
CREATE OR ALTER PROCEDURE cis.InvEnq_Case @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Ws_SP_INV_CLAIM @StrSql = @casNumber;
    SELECT TOP (1) GPFee, GPLTMed, GPInj, GPMed, SPFee, SPLTMed, SPInj, SPMed FROM cis.Ws_ClaimInvoiceActualOthers WHERE CASNumber = @casNumber;
END
GO

-- CmdNext_Click / CmdPrev_Click: the next / previous clinical case of an active member.
CREATE OR ALTER PROCEDURE cis.InvEnq_Step @casNumber nvarchar(20), @next bit AS
BEGIN
    SET NOCOUNT ON;
    IF @next = 1
        SELECT TOP (1) CASNumber FROM dbo.VIEW_Cas_MBM WHERE MBMCLNDataStatus <> 'D' AND MBMCLNStatus <> 'D' AND HSCLNInd IN ('C', 'M', 'B') AND CaseType = 'C'
            AND CASNumber > @casNumber ORDER BY CASNumber;
    ELSE
        SELECT TOP (1) CASNumber FROM dbo.VIEW_Cas_MBM WHERE MBMCLNDataStatus <> 'D' AND MBMCLNStatus <> 'D' AND HSCLNInd IN ('C', 'M', 'B') AND CaseType = 'C'
            AND CASNumber < @casNumber ORDER BY CASNumber DESC;
END
GO

-- WSInfo: Claim and the seven worksheet tables of the case (actual, contracted rate, approved, limit, member excess, clinic excess, clinic approved).
CREATE OR ALTER PROCEDURE cis.InvEnq_Worksheet @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) * FROM cis.Ws_Claim WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMActualNew WHERE CLMNo = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMClinicLmtNew WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMAppNew WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMLimitNew WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMMemberNew WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMClinicNew WHERE CASNumber = @casNumber;
    SELECT TOP (1) * FROM cis.Ws_CLMClinicAppNew WHERE CASNumber = @casNumber;
END
GO

-- ClaimBordxInfo: SP_Claim_Bordx.
CREATE OR ALTER PROCEDURE cis.InvEnq_Bordx @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Ws_SP_Claim_Bordx @tmpCasNo = @casNumber;
END
GO

-- PaymentInfo / ReceiptInfo (CISPayment).
CREATE OR ALTER PROCEDURE cis.InvEnq_Payments @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT PVCLNNo AS PvNo, PVCLNDate AS PvDate, PVCLNChequeAmt AS ChequeAmt, 'CLINIC' AS PayTo, CLNPaidAmt AS PaidAmt, CLNRemark AS Remark,
           PVCLNChequeDate AS ChequeDate
    FROM cis.Pay_Payment2CLN WHERE CASNumber = @casNumber
    UNION ALL
    SELECT PVNo, PVDate, PVCLNChequeAmt, PVPayto, MBMPaidAmt, Remark, PVChequeDate FROM cis.Pay_Payment2MBM WHERE CASNumber = @casNumber;
    SELECT ReceiptNo, ReceiptDate, ReceiptAmt FROM cis.Pay_ReceiptFrMBM WHERE CASNumber = @casNumber;
END
GO

-- MedicationRequestInfo / MedicineDetails.
CREATE OR ALTER PROCEDURE cis.InvEnq_MedRequests @casNumber nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_CISMedRequestInfo @CaseNo = @casNumber;
    EXEC dbo.sp_CISMedReqDOMedicine @CaseNo = @casNumber;
END
GO
