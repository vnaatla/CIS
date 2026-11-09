-- Claim Bordx > Worksheet > Remove Worksheet (FrmBordxRevise): a bordereaux's worksheets (ClaimBordxModule case 3) and RemoveFromBordx.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxRemove_List @bordxNo nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 30000 CASNumber, ClaimOpendate, MBMNumber, PatientName, DateVisit, CLMTotalApproved, CASStatus, BordxPrintedDate, BordxDate, BordxClaimAmt
    FROM CISWorksheet.dbo.VIEW_CIS_BordxRevise WHERE BordxRefNo = @bordxNo;
END
GO

-- RemoveFromBordx for one case: the claim count comes off the case's own bordx, the amount off the bordx typed on the screen (as the desktop).
CREATE OR ALTER PROCEDURE cis.BordxRemove_Case @bordxNo nvarchar(20), @casNumber nvarchar(30), @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @caseBordx nvarchar(20), @approved money, @bordxDate datetime;
    SELECT TOP 1 @caseBordx = BordxRefNo, @approved = CLMFinalApprove FROM CISWorksheet.dbo.Claim WHERE CASNumber = @casNumber;
    SELECT TOP 1 @bordxDate = BordxDate FROM CISWorksheet.dbo.Bordx WHERE BordxRefNo = @bordxNo;
    UPDATE CISWorksheet.dbo.Bordx SET BordxNoClaims = BordxNoClaims - 1 WHERE BordxRefNo = @caseBordx;
    INSERT CISWorksheet.dbo.BordxWSArchive (CASNumber, BordxRefNo, BordxDate, LastUpdateDate, LastUpdateUser)
    VALUES (@casNumber, @bordxNo, @bordxDate, CAST(GETDATE() AS date), @user);
    UPDATE CISWorksheet.dbo.Bordx SET BordxClaimAmt = BordxClaimAmt - @approved WHERE BordxRefNo = @bordxNo;
    UPDATE CISWorksheet.dbo.ClaimInvoice SET INVBordxRefNo = '', INVBordxed = 0 WHERE INVBordxRefNo = @bordxNo AND CASNumber = @casNumber;
    UPDATE CISWorksheet.dbo.Claim SET ClaimBordxStatus = 0, BordxRefNo = '',
        CLMBordxWSStatus = CASE WHEN CLMBordxWSASCII = 'Y' THEN 'D' ELSE CLMBordxWSStatus END,
        CLMBordxWSASCII = CASE WHEN CLMBordxWSASCII = 'Y' THEN 'N' ELSE CLMBordxWSASCII END
    WHERE CASNumber = @casNumber;
END
GO
