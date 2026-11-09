-- Claim Bordx > Print Bordx (frmPrintBordxExcel and the client variants): the bordereaux header (GetBordxInfo), its pay-advice rows
-- (CISWorksheet CLM_PayAdviceCLNBSN case 1, called with the bordx no. quoted here), the group summary (VIEW_OP_GRPSummaryALL) and UpdateBordx.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.PrintBordx_Info @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 BordxRefNo, BordxDate, GroupCompany, RptTypeByDept, BordxPrintedStatus, BORDXPrintMonth, BORDXPrintYear, BordxClaimAmt, BordxNoClaims
    FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @bordxNo;
END
GO

CREATE OR ALTER PROCEDURE cis.PrintBordx_Rows @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @append nvarchar(200) = N' CISWorksheet.dbo.Claim.BordxRefNo = ''' + REPLACE(@bordxNo, '''', '''''') + N'''';
    EXEC CISWorksheet.dbo.CLM_PayAdviceCLNBSN @strAppend = @append, @strAppendCase = 1;
END
GO

CREATE OR ALTER PROCEDURE cis.PrintBordx_Summary @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM dbo.VIEW_OP_GRPSummaryALL WHERE BordxRefNo = @bordxNo;
END
GO

-- UpdateBordx (Print, not Preview).
CREATE OR ALTER PROCEDURE cis.PrintBordx_Update @bordxNo nvarchar(20), @month nvarchar(20), @year nvarchar(4), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISWorksheet.dbo.Bordx SET BordxPrintedDate = CAST(GETDATE() AS date), BORDXPrintMonth = @month, BORDXPrintYear = @year, BordxPrintedStatus = '1',
        BordxPrintedBy = @user
    WHERE BordxRefNo = @bordxNo;
END
GO
