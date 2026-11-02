-- Invoice > Clinic > Batch Transfer (FrmInvoiceBatchNoUpd): move the cases of one OP batch to another (CISDB CISCAS).

-- ShowInfo: the batch's cases per invoice / clinic (the desktop's SP_SeachCasByOPBatch, already parameterised).
CREATE OR ALTER PROCEDURE cis.BatchTransfer_Batch @batchNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.SP_SeachCasByOPBatch @tmpBatchNo = @batchNo;
END
GO

-- BatchTransProcess: batch 1's cases move to batch 2 with the transfer remark, then batch 2's invoice statement amount becomes the total.
CREATE OR ALTER PROCEDURE cis.BatchTransfer_Save @fromBatch nvarchar(30), @toBatch nvarchar(30), @remark nvarchar(400), @amount money AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET OPBatchNo = @toBatch, BatchTransRemarks = @remark WHERE OPBatchNo = @fromBatch;
    DECLARE @moved int = @@ROWCOUNT;
    UPDATE dbo.CISCAS SET InvoiceStmtAmt = @amount WHERE OPBatchNo = @toBatch;
    SELECT @moved;
END
GO
