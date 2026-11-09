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

-- @variant: the client screens' summary sources (Prasarana View_CISBordxSummary_SPNB, AXA Affin / Time Dotcom VIEW_OP_AXASummary, BSN View_OP_BSNSummary,
-- Berjaya Sompo / Crystal View_CIS_PrintBordx_Summary); standard and SEB / Indah Water VIEW_OP_GRPSummaryALL.
CREATE OR ALTER PROCEDURE cis.PrintBordx_Summary @bordxNo nvarchar(20), @variant varchar(20) = 'standard' AS
BEGIN
    SET NOCOUNT ON;
    IF @variant = 'prasarana' SELECT * FROM CISWorksheet.dbo.View_CISBordxSummary_SPNB WHERE BordxRefNo = @bordxNo;
    ELSE IF @variant IN ('axa-affin', 'time-dotcom') SELECT * FROM dbo.VIEW_OP_AXASummary WHERE BordxRefNo = @bordxNo;
    ELSE IF @variant = 'bsn' SELECT * FROM dbo.View_OP_BSNSummary WHERE BordxRefNo = @bordxNo;
    ELSE IF @variant IN ('berjaya-sompo', 'crystal') SELECT * FROM CISWorksheet.dbo.View_CIS_PrintBordx_Summary WHERE BordxRefNo = @bordxNo;
    ELSE SELECT * FROM dbo.VIEW_OP_GRPSummaryALL WHERE BordxRefNo = @bordxNo;
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
