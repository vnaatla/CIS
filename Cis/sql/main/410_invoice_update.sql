-- Invoice > Other > Invoice Update (frmUpdateCLNInvoice): list cases by invoice / clinic / payor / dates (CISWorksheet
-- VIEW_Invoice_Tracking_2016, the desktop's ClaimBordxModule case 1) and stamp the checked cases with an invoice number and date (CISCAS).

-- ChkClinic_Click / ChkGrpComp_Click: panel clinics (HISDB CLN, panel Y / C) and group companies (CISDB GRPCompany).
CREATE OR ALTER PROCEDURE cis.InvUpdate_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(CLNCode) AS Code, RTRIM(CLNName) AS Name FROM cis.Hdb_CLN WHERE (CLNPanel = 'Y' OR CLNPanel = 'C') AND LEN(CLNName) > 0;
    SELECT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany;
END
GO

-- SearchRecord → InvoiceListing: every filter optional; @noInvoice = ChkInvNo (BillNo empty).
CREATE OR ALTER PROCEDURE cis.InvUpdate_Search
    @casFrom nvarchar(20) = NULL, @casTo nvarchar(20) = NULL, @clinicCode nvarchar(10) = NULL, @groupCompany nvarchar(200) = NULL,
    @billNo nvarchar(100) = NULL, @visitFrom datetime = NULL, @visitTo datetime = NULL, @noInvoice bit = 0, @billFrom datetime = NULL,
    @billTo datetime = NULL, @amountFrom money = NULL, @amountTo money = NULL, @billDate datetime = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT CASNumber, MBMnumber AS MBMNumber, CoverID, PatientName, DateVisit, CLNName, BillNo, BillDate, INVEnteredDate, INVEnteredBy, INVAmount,
           CLMMCOProcFee, CLMFinalApprove, MBMGroupCompany
    FROM cis.Ws_VIEW_Invoice_Tracking_2016
    WHERE (@casFrom IS NULL OR CASNumber >= @casFrom) AND (@casTo IS NULL OR CASNumber <= @casTo)
      AND (@clinicCode IS NULL OR ClinicCode = @clinicCode) AND (@groupCompany IS NULL OR MBMGroupCompany = @groupCompany)
      AND (@billNo IS NULL OR BillNo = @billNo) AND (@visitFrom IS NULL OR DateVisit >= @visitFrom) AND (@visitTo IS NULL OR DateVisit <= @visitTo)
      AND (@noInvoice = 0 OR BillNo IS NULL OR BillNo = '') AND (@billFrom IS NULL OR BillDate >= @billFrom) AND (@billTo IS NULL OR BillDate <= @billTo)
      AND (@amountFrom IS NULL OR INVAmount >= @amountFrom) AND (@amountTo IS NULL OR INVAmount <= @amountTo)
      AND (@billDate IS NULL OR BillDate = @billDate);
END
GO

-- UpdateRecord: one checked case.
CREATE OR ALTER PROCEDURE cis.InvUpdate_Save @casNumber nvarchar(20), @invoiceNo nvarchar(100), @invoiceDate datetime, @today datetime, @userCode nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET InvoiceNo = @invoiceNo, InvoiceDate = @invoiceDate, RegisterDate = @today, RegisterBy = @userCode WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO
