-- Claim Bordx > Insert Invoice (frmWSBordx2): VIEW_CISINV_PendingBordx with the pasted criteria replaced by parameters, and the bordx of the ticked invoices.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxInv_Search @regFrom datetime = NULL, @regTo datetime = NULL, @payCode nvarchar(10) = NULL, @grpCode nvarchar(10) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM dbo.VIEW_CISINV_PendingBordx
    WHERE (@regFrom IS NULL OR RegisterDate >= @regFrom) AND (@regTo IS NULL OR RegisterDate <= @regTo)
      AND (@payCode IS NULL OR GRPCoPaycode = @payCode) AND (@grpCode IS NULL OR GRPCoCode = @grpCode)
    OPTION (RECOMPILE);
END
GO

-- The not-yet-bordxed claims behind one listed invoice row.
CREATE OR ALTER PROCEDURE cis.BordxInv_Cases @invoiceNo nvarchar(50), @invoiceDate datetime = NULL, @registerDate datetime = NULL, @grpCode nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(k.CASNumber) AS CASNumber, k.CLMFinalApprove
    FROM dbo.CISCas c
    JOIN CISWorksheet.dbo.Claim k ON k.CASNumber = c.CASNumber AND k.ClaimBordxStatus = 0
    JOIN cis.Hdb_MBMOthers o ON o.MBMNumber = c.MBMnumber
    JOIN dbo.GRPCompany g ON g.GRPCoName = o.MBMGroupCompany AND g.GRPCoCode = @grpCode
    WHERE c.InvoiceNo = @invoiceNo
      AND ((@invoiceDate IS NULL AND c.InvoiceDate IS NULL) OR c.InvoiceDate = @invoiceDate)
      AND ((@registerDate IS NULL AND c.RegisterDate IS NULL) OR c.RegisterDate = @registerDate);
END
GO

-- UpdateRecord for one claim.
CREATE OR ALTER PROCEDURE cis.BordxInv_UpdateCase @casNumber nvarchar(30), @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISWorksheet.dbo.Claim SET BordxRefNo = @bordxNo, ClaimBordxStatus = 1, CLMPending = 'N' WHERE CASNumber = @casNumber;
    UPDATE CISWorksheet.dbo.ClaimInvoice SET INVBordxRefNo = @bordxNo, INVBordxed = 1 WHERE CASNumber = @casNumber;
    UPDATE dbo.CISCas SET Pending = 'N' WHERE CASNumber = @casNumber;
END
GO

-- SaveCLMBordx (GroupCompany holds the combo text "code - name", as the desktop).
CREATE OR ALTER PROCEDURE cis.BordxInv_SaveBordx @bordxNo nvarchar(20), @payCode nvarchar(4), @company nvarchar(200), @user nvarchar(50), @amount money, @claims int AS
BEGIN
    SET NOCOUNT ON;
    INSERT CISWorksheet.dbo.Bordx (BordxPAYCode, BordxRefNo, BordxDate, BordxBy, BordxClaimAmt, BordxNoClaims, GroupCompany)
    VALUES (@payCode, @bordxNo, CAST(GETDATE() AS date), @user, @amount, @claims, @company);
END
GO

-- Combos: group companies (with their payor) and payors.
CREATE OR ALTER PROCEDURE cis.BordxInv_Lookups AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GRPCoCode) AS Code, RTRIM(GRPCoName) AS Name, RTRIM(GRPCoPaycode) AS PayCode FROM dbo.GRPCompany ORDER BY GRPCoCode;
    SELECT RTRIM(PAYCode) AS Code, RTRIM(PAYCompanyName) AS Name FROM dbo.PAY ORDER BY PAYCode;
END
GO
