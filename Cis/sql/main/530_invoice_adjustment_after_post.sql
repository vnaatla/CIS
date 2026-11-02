-- Invoice > Clinic > Adjustment (After Posting) (frmInvoiceAdjAfterPost): a case already posted to SUN is not changed; a sub-case is created
-- (sp_CreateSubCasNum / sp_CreateSubCasNumClaim with the master case number). Its GL amount follows SaveINVClaim.
CREATE OR ALTER PROCEDURE cis.InvAdj_LgAmount @casNumber nvarchar(20), @glAmount money AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.LG SET GLAmt = @glAmount WHERE LGNo = @casNumber;
END
GO
