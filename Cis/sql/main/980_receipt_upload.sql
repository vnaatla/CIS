-- Receipt > Upload Receipt / Upload Receipt (Multiple Receipts) (frmReceiptUpload / frmReceiptUpMultiple): the claims of the uploaded rows and SaveBordx.
SET QUOTED_IDENTIFIER ON;
GO

-- cmpUpload_Click: the claim of each uploaded claim no. (@cases comma-separated).
CREATE OR ALTER PROCEDURE cis.RcpUp_Claims @cases nvarchar(max) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @x xml = CAST('<c>' + REPLACE(REPLACE(REPLACE(REPLACE(@cases, '&', ''), '<', ''), ' ', ''), ',', '</c><c>') + '</c>' AS xml);
    DECLARE @c TABLE (CASNumber nvarchar(30) PRIMARY KEY);
    INSERT @c SELECT DISTINCT t.c.value('.', 'nvarchar(30)') FROM @x.nodes('/c') t(c) WHERE t.c.value('.', 'nvarchar(30)') <> '';
    SELECT RTRIM(k.CASNumber) AS CASNumber, RTRIM(k.BordxRefNo) AS BordxRefNo, k.CLMBordxAmt
    FROM CISWorksheet.dbo.Claim k WHERE k.CASNumber IN (SELECT CASNumber FROM @c);
END
GO

-- SaveBordx (one matched row): a ReceiptfrPAY row and the claim received in full.
CREATE OR ALTER PROCEDURE cis.RcpUp_Save @casNumber nvarchar(30), @bordxNo nvarchar(20) = NULL, @receiptNo nvarchar(30), @receiptDate datetime = NULL,
    @receiptAmt money, @chequeNo nvarchar(30) = NULL, @excelAmt money, @mth nvarchar(2) = NULL, @yr nvarchar(4) = NULL, @clientPvNo nvarchar(50) = NULL,
    @medixReceiptNo nvarchar(30) = NULL, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    IF @bordxNo IS NULL SET @bordxNo = (SELECT TOP 1 RTRIM(BordxRefNo) FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber);
    INSERT CISPayment.dbo.ReceiptfrPAY (CASNumber, ReceiptNo, ReceiptDate, ReceiptAmt, ChequeNo, BordxNo, PAYLastUpdateDate, PAYLastUpdateUser, RecPayAmt, CheckAmt,
        RECAccPeriodMth, RECAccPeriodYR, ClientPVNo, MedixReceiptNo)
    VALUES (@casNumber, @receiptNo, @receiptDate, @receiptAmt, COALESCE(@chequeNo, ''), COALESCE(@bordxNo, ''), CAST(GETDATE() AS date), @user, @excelAmt, @excelAmt,
        @mth, @yr, @clientPvNo, @medixReceiptNo);
    UPDATE CISWorksheet.dbo.Claim SET PVRecPayorStatus = 'F' WHERE CASNumber = @casNumber;
END
GO
