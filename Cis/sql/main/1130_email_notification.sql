-- E-mail notification > Payment (frmEmailNotification): View_Email_notification / View_Email_Details (CISPayment) with the pasted criteria replaced by parameters,
-- the claimant's name (the covered person for dependants) and the sent flag.
SET QUOTED_IDENTIFIER ON;
GO

-- cmdSearch_Click / PaymentListing. @resend 1 = "R - Resend" (already e-mailed), 0 = "N - New"; @cs 1 = Sihat Johor (PV no. CS...), 0 = the others.
CREATE OR ALTER PROCEDURE cis.EmlPay_Search @pvDateFrom datetime, @pvDateTo datetime, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL, @resend bit = 0,
    @company nvarchar(200) = NULL, @cs bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT v.*, COALESCE(CASE WHEN RTRIM(v.CoverID) <> '00' THEN c.MBMCName END, v.MBMName) AS PatientName
    FROM CISPayment.dbo.View_Email_notification v
    OUTER APPLY (SELECT TOP 1 MBMCName FROM HISDB.dbo.MBMCoveredPersons p WHERE p.MBMCNumber = RTRIM(v.MBMnumber) AND p.MBMCCoverID = v.CoverID) c
    WHERE v.PVDate >= @pvDateFrom AND v.PVDate <= @pvDateTo AND (@pvFrom IS NULL OR v.PVNo >= @pvFrom) AND (@pvTo IS NULL OR v.PVNo <= @pvTo)
      AND ((@resend = 1 AND v.PVEmailStatus = 'Y') OR (@resend = 0 AND (v.PVEmailStatus = '' OR v.PVEmailStatus IS NULL)))
      AND (@company IS NULL OR v.MBMGroupCompany = @company)
      AND ((@cs = 0 AND SUBSTRING(v.PVNo, 1, 2) <> 'CS') OR (@cs = 1 AND SUBSTRING(v.PVNo, 1, 2) = 'CS'))
    ORDER BY v.MBMEmail DESC
    OPTION (RECOMPILE);
END
GO

-- CmdSend_Click: the claims of a PV for the e-mail table.
CREATE OR ALTER PROCEDURE cis.EmlPay_Details @pvNo nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT d.MBMName, d.DateVisit, d.CLMTotalActual, d.CLMBordxAmt, d.CoverID, d.MBMNumber, d.CasNumber, d.MBMGroupCompany,
        COALESCE(CASE WHEN d.CoverID <> '00' THEN c.MBMCName END, d.MBMName) AS PatientName
    FROM CISPayment.dbo.View_Email_Details d
    OUTER APPLY (SELECT TOP 1 MBMCName FROM HISDB.dbo.MBMCoveredPersons p WHERE p.MBMCNumber = RTRIM(d.MBMNumber) AND p.MBMCCoverID = d.CoverID) c
    WHERE d.PVNo = @pvNo;
END
GO

-- CmdSend_Click: the PV is marked as e-mailed.
CREATE OR ALTER PROCEDURE cis.EmlPay_Mark @pvNo nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISPayment.dbo.Payment2MBM SET PVEmailStatus = 'Y', PVEmailDate = CAST(GETDATE() AS date) WHERE PVNo = @pvNo;
END
GO

-- PayorComboListing: the group companies.
CREATE OR ALTER PROCEDURE cis.EmlPay_Companies AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoName) FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL;
END
GO

-- E-mail notification > Providers (frmEmailNotificationProvider): View_email_Provider with parameters (clinic PVs); the invoices of a PV come from the existing
-- CISPayment sp_ProviderInv (already parameterized).
CREATE OR ALTER PROCEDURE cis.EmlPrv_Search @pvDateFrom datetime, @pvDateTo datetime, @pvFrom nvarchar(30) = NULL, @pvTo nvarchar(30) = NULL, @resend bit = 0,
    @cs bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM CISPayment.dbo.View_email_Provider
    WHERE PVCLNDate >= @pvDateFrom AND PVCLNDate <= @pvDateTo AND (@pvFrom IS NULL OR PVCLNNo >= @pvFrom) AND (@pvTo IS NULL OR PVCLNNo <= @pvTo)
      AND ((@resend = 1 AND PVEmailStatus = 'Y') OR (@resend = 0 AND (PVEmailStatus = '' OR PVEmailStatus IS NULL)))
      AND ((@cs = 0 AND SUBSTRING(PVCLNNo, 1, 2) <> 'CS') OR (@cs = 1 AND SUBSTRING(PVCLNNo, 1, 2) = 'CS'))
    ORDER BY CLNEMAIL DESC
    OPTION (RECOMPILE);
END
GO

-- CmdSend_Click: the clinic PV is marked as e-mailed (the desktop sets no date here).
CREATE OR ALTER PROCEDURE cis.EmlPrv_Mark @pvNo nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISPayment.dbo.Payment2CLN SET PVEmailStatus = 'Y' WHERE PVCLNNo = @pvNo;
END
GO
