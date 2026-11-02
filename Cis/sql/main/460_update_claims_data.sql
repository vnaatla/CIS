-- Claim Bordx > Update Claims Data (frmUpdateClaimsData): spreadsheet uploads into CISWorksheet Claim / Bordx.
-- The desktop's uploads could not work (see the screen doc); these are the working versions agreed for the web.

-- UpdateFFS: CLMMCOProcFee of a case.
CREATE OR ALTER PROCEDURE cis.ClaimData_Ffs @casNumber nvarchar(20), @amount money, @today datetime, @userCode nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Ws_Claim SET CLMMCOProcFee = @amount, ClaimLastUpdateDate = @today, ClaimLastUpdateUser = @userCode WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO

-- UpdateCopay: CLMCoPayColl of a case.
CREATE OR ALTER PROCEDURE cis.ClaimData_Copay @casNumber nvarchar(20), @amount money, @today datetime, @userCode nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Ws_Claim SET CLMCoPayColl = @amount, ClaimLastUpdateDate = @today, ClaimLastUpdateUser = @userCode WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO

-- UpdateSent: CLMBordxAmt of a case.
CREATE OR ALTER PROCEDURE cis.ClaimData_BordxAmt @casNumber nvarchar(20), @amount money, @today datetime, @userCode nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Ws_Claim SET CLMBordxAmt = @amount, ClaimLastUpdateDate = @today, ClaimLastUpdateUser = @userCode WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO

-- UpdateBordxDate: the bordereaux month / year of its claims and the Bordx date.
CREATE OR ALTER PROCEDURE cis.ClaimData_BordxDate @bordxRefNo nvarchar(30), @date datetime, @today datetime, @userCode nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE cis.Ws_Claim SET CLMBordxMonth = CAST(MONTH(@date) AS varchar(2)), CLMBordxYear = CAST(YEAR(@date) AS varchar(4)),
        ClaimLastUpdateDate = @today, ClaimLastUpdateUser = @userCode
    WHERE BordxRefNo = @bordxRefNo;
    DECLARE @n int = @@ROWCOUNT;
    UPDATE cis.Ws_Bordx SET BordxDate = @date WHERE BordxRefNo = @bordxRefNo;
    SELECT @n;
END
GO
