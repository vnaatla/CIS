-- Maintenance > Entity (frmMaintenanceEntity) tabs Clinic Group (GRPClinic) and Doctor (Doctor).

-- SearchClinicGrp → GrpClinicListing: GRPClnCode / GRPClnName LIKE %x% (upper-cased by the service).
CREATE OR ALTER PROCEDURE cis.EntityGroup_Search @code nvarchar(10) = NULL, @name nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT GRPClnCode, GRPClnName FROM dbo.GRPClinic
    WHERE (@code IS NULL OR GRPClnCode LIKE N'%' + @code + N'%') AND (@name IS NULL OR GRPClnName LIKE N'%' + @name + N'%')
    ORDER BY GRPClnCode;
END
GO

-- lstGrpClinicList_Click.
CREATE OR ALTER PROCEDURE cis.EntityGroup_Get @code nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT GRPClnCode AS Code, GRPClnName AS Name, GRPClnInitial AS Initial, GRPClnAdd1 AS Add1, GRPClnAdd2 AS Add2, GRPClnAdd3 AS Add3,
           GRPClnContact AS Contact, GRPClnAdmContact AS AdmContact, GRPClnCEO AS Ceo, GRPClnPhone AS Phone, GRPClnFax AS Fax,
           GRPClnEmail AS Email, GRPClnHomepage AS Homepage, CONVERT(nvarchar(10), GRPClnEffDate, 103) AS EffDate,
           RTRIM(GRPClnTax) AS Tax, RTRIM(GRPClnConversion) AS Conversion, RTRIM(GRPClnPrePost) AS PrePost
    FROM dbo.GRPClinic WHERE GRPClnCode = @code;
END
GO

-- GenerateGRPClnSeq (Medix.bas): GRPClnSeq row of the letter, +1, "0000"; letter + seq. Skips numbers already used in GRPClinic
-- (the desktop's insert then failed on PK_GRPClinic).
CREATE OR ALTER PROCEDURE cis.EntityGroup_NextCode @letter nchar(1) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seq int = (SELECT TRY_CAST(GRPClnSeq AS int) FROM dbo.GRPClnSeq WITH (UPDLOCK, HOLDLOCK) WHERE GRPClnCode = @letter);
    DECLARE @new bit = CASE WHEN @seq IS NULL THEN 1 ELSE 0 END;
    SET @seq = ISNULL(@seq, 0) + 1;
    WHILE EXISTS (SELECT 1 FROM dbo.GRPClinic WHERE GRPClnCode = @letter + RIGHT(N'0000' + CAST(@seq AS nvarchar(10)), 4)) SET @seq += 1;
    DECLARE @text nvarchar(5) = RIGHT(N'0000' + CAST(@seq AS nvarchar(10)), 4);
    IF @new = 1 INSERT INTO dbo.GRPClnSeq (GRPClnCode, GRPClnSeq) VALUES (@letter, @text);
    ELSE UPDATE dbo.GRPClnSeq SET GRPClnSeq = @text WHERE GRPClnCode = @letter;
    SELECT @letter + @text;
END
GO

-- SaveClinicGrp (Add and Edit write the same columns).
CREATE OR ALTER PROCEDURE cis.EntityGroup_Save
    @insert bit, @code nvarchar(5), @name nvarchar(50), @initial nvarchar(20), @add1 nvarchar(50) = NULL, @add2 nvarchar(50) = NULL,
    @add3 nvarchar(50) = NULL, @contact nvarchar(50) = NULL, @admContact nvarchar(50) = NULL, @ceo nvarchar(50) = NULL,
    @phone nvarchar(50) = NULL, @fax nvarchar(50) = NULL, @email nvarchar(50) = NULL, @homepage nvarchar(50), @effDate datetime = NULL,
    @tax nvarchar(5), @conversion nvarchar(5), @prePost nvarchar(5), @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF @insert = 1
        INSERT INTO dbo.GRPClinic (GRPClnCode, GRPClnName, GRPClnInitial, GRPClnAdd1, GRPClnAdd2, GRPClnAdd3, GRPClnContact, GRPClnAdmContact,
                                   GRPClnCEO, GRPClnPhone, GRPClnFax, GRPClnEmail, GRPClnHomepage, GRPClnEffDate, GRPClnTax, GRPClnConversion,
                                   GRPClnPrePost, GRPClnLastUpdateUser, GRPClnLastUpdateDate)
        VALUES (@code, @name, @initial, @add1, @add2, @add3, @contact, @admContact, @ceo, @phone, @fax, @email, @homepage, @effDate, @tax,
                @conversion, @prePost, @user, @today);
    ELSE
        UPDATE dbo.GRPClinic
        SET GRPClnName = @name, GRPClnInitial = @initial, GRPClnAdd1 = @add1, GRPClnAdd2 = @add2, GRPClnAdd3 = @add3, GRPClnContact = @contact,
            GRPClnAdmContact = @admContact, GRPClnCEO = @ceo, GRPClnPhone = @phone, GRPClnFax = @fax, GRPClnEmail = @email,
            GRPClnHomepage = @homepage, GRPClnEffDate = @effDate, GRPClnTax = @tax, GRPClnConversion = @conversion, GRPClnPrePost = @prePost,
            GRPClnLastUpdateUser = @user, GRPClnLastUpdateDate = @today
        WHERE GRPClnCode = @code;
    SELECT @@ROWCOUNT;
END
GO

-- DeleteClinicGrp: refused while a clinic (HISDB CLN.ClinicGrpCode) uses the group.
CREATE OR ALTER PROCEDURE cis.EntityGroup_InUse @code nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM cis.Hdb_CLN WHERE ClinicGrpCode = @code) THEN 1 ELSE 0 END AS bit);
END
GO

CREATE OR ALTER PROCEDURE cis.EntityGroup_Delete @code nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.GRPClinic WHERE GRPClnCode = @code;
    SELECT @@ROWCOUNT;
END
GO

-- SearchDoctor → DoctorListing.
CREATE OR ALTER PROCEDURE cis.EntityDoctor_Search @code nvarchar(10) = NULL, @name nvarchar(50) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(DocCode) AS DocCode, DocName FROM dbo.Doctor
    WHERE (@code IS NULL OR DocCode LIKE N'%' + @code + N'%') AND (@name IS NULL OR DocName LIKE N'%' + @name + N'%')
    ORDER BY DocCode;
END
GO

-- SaveDoctor duplicate check on Add: the name is already used.
CREATE OR ALTER PROCEDURE cis.EntityDoctor_NameExists @name nvarchar(50) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.Doctor WHERE DocName = @name) THEN 1 ELSE 0 END AS bit);
END
GO

-- GenerateDocSeq (Medix.bas): DOCSeq row of the letter, +1, "00000"; letter + seq; numbers already used in Doctor are skipped.
CREATE OR ALTER PROCEDURE cis.EntityDoctor_NextCode @letter nchar(1) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seq int = (SELECT TRY_CAST(DocSeq AS int) FROM dbo.DOCSeq WITH (UPDLOCK, HOLDLOCK) WHERE DocCode = @letter);
    DECLARE @new bit = CASE WHEN @seq IS NULL THEN 1 ELSE 0 END;
    SET @seq = ISNULL(@seq, 0) + 1;
    WHILE EXISTS (SELECT 1 FROM dbo.Doctor WHERE DocCode = @letter + RIGHT(N'00000' + CAST(@seq AS nvarchar(10)), 5)) SET @seq += 1;
    DECLARE @text nvarchar(5) = RIGHT(N'00000' + CAST(@seq AS nvarchar(10)), 5);
    IF @new = 1 INSERT INTO dbo.DOCSeq (DocCode, DocSeq) VALUES (@letter, @text);
    ELSE UPDATE dbo.DOCSeq SET DocSeq = @text WHERE DocCode = @letter;
    SELECT @letter + @text;
END
GO

-- SaveDoctor: Add inserts; Edit rewrites the name (the box is locked on Edit, so it is unchanged) and the audit columns.
CREATE OR ALTER PROCEDURE cis.EntityDoctor_Save @insert bit, @code nvarchar(10), @name nvarchar(50), @user nvarchar(10), @today datetime AS
BEGIN
    SET NOCOUNT ON;
    IF @insert = 1
        INSERT INTO dbo.Doctor (DocCode, DocName, DocLastUpdateUser, DocLastUpdateDate) VALUES (@code, @name, @user, @today);
    ELSE
        UPDATE dbo.Doctor SET DocName = @name, DocLastUpdateUser = @user, DocLastUpdateDate = @today WHERE DocCode = @code;
    SELECT @@ROWCOUNT;
END
GO

CREATE OR ALTER PROCEDURE cis.EntityDoctor_Delete @code nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.Doctor WHERE DocCode = @code;
    SELECT @@ROWCOUNT;
END
GO
