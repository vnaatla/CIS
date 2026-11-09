-- Claim Bordx > ClaimBordxDate (FrmBordxDate): BordxDateChange and UpdateBordxDate.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxDate_Find @bordxNo nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    EXEC CISWorksheet.dbo.BordxDateChange @BordxNo = @bordxNo;
END
GO

CREATE OR ALTER PROCEDURE cis.BordxDate_Update @bordxNo nvarchar(20), @bordxDate datetime, @user nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE CISWorksheet.dbo.Bordx SET BordxBy = @user, BordxDate = @bordxDate WHERE BordxRefNo = @bordxNo;
    UPDATE CISWorksheet.dbo.Claim SET CLMBordxMonth = FORMAT(@bordxDate, 'MM'), CLMBordxYear = FORMAT(@bordxDate, 'yyyy') WHERE BordxRefNo = @bordxNo;
END
GO
