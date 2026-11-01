-- Maintenance > Entity (frmMaintenanceEntity) tab Clinic Info: HISDB CLN (medixmasterconn), ClinicGrp / STA (HISDB), CTY / CLNSeq
-- (HISMaintenance), cases in CISDB CISCAS.

-- CLNComboListing: clinic groups, cities, states.
CREATE OR ALTER PROCEDURE cis.EntityClinic_Setup AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(ClinicGrpCode) AS Code, ISNULL(ClinicGrpName, N'') AS Name FROM cis.Hdb_ClinicGrp ORDER BY ClinicGrpCode;
    SELECT DISTINCT RTRIM(CTYDescription) AS Code, RTRIM(CTYDescription) AS Name FROM cis.Maint_CTY WHERE LEN(CTYDescription) > 0 ORDER BY Code;
    SELECT DISTINCT RTRIM(STADescription) AS Code, RTRIM(STADescription) AS Name FROM cis.Hdb_STA WHERE LEN(STADescription) > 0 ORDER BY Code;
END
GO

-- SearchCLN → CLNListing (Select * from CLN where …). Values are already reduced by the service (upper-case, first letters, group prefix).
CREATE OR ALTER PROCEDURE cis.EntityClinic_Search
    @code nvarchar(5) = NULL, @initial nvarchar(50) = NULL, @name nvarchar(100) = NULL, @state nvarchar(30) = NULL, @city nvarchar(30) = NULL,
    @postCode nvarchar(10) = NULL, @panel nchar(1) = NULL, @groupPrefix nvarchar(10) = NULL, @type nchar(1) = NULL, @hrs24 nchar(1) = NULL,
    @ffs nchar(1) = NULL, @online nchar(1) = NULL, @glFormat nchar(1) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT CLNCode, CLNName, RTRIM(ClinicGrpCode) AS ClinicGrpCode, CLNInitial, CLNType, CLNBNFType, CLNAdd1, CLNAdd2, CLNAdd3, RTRIM(CLNPostCode) AS CLNPostCode,
           CLNCity, CLNState, CLNReceivedDate, CLNPanel, CLNXRay, CLNForRef, CLNNoofDoc, CLNPhone, CLNFax, CLNEmail, CLNHomepage, CLNBankAcc,
           CLNBankName, CLNBankBrn, CLNAccPerson, CLN24Hrs, CLNRemark, CLNFFSStatus, CLNOperatingHr, CLNPName, CLNPAdd1, CLNPAdd2, CLNPAdd3,
           RTRIM(CLNPPostCode) AS CLNPPostCode, CLNPCity, CLNPState, CLNOnlineStatus, CLNRegNo, CLNLocation, CLNMCOProcfeePerc, CLNGLFormat
    FROM cis.Hdb_CLN
    WHERE (@code IS NULL OR CLNCode = @code) AND (@initial IS NULL OR CLNInitial LIKE N'%' + @initial + N'%')
      AND (@name IS NULL OR CLNName LIKE N'%' + @name + N'%') AND (@state IS NULL OR CLNState LIKE N'%' + @state + N'%')
      AND (@city IS NULL OR CLNCity LIKE N'%' + @city + N'%') AND (@postCode IS NULL OR CLNPostCode LIKE N'%' + @postCode + N'%')
      AND (@panel IS NULL OR CLNPanel = @panel) AND (@groupPrefix IS NULL OR ClinicGrpCode LIKE @groupPrefix + N'%')
      AND (@type IS NULL OR CLNType = @type) AND (@hrs24 IS NULL OR CLN24Hrs = @hrs24) AND (@ffs IS NULL OR CLNFFSStatus = @ffs)
      AND (@online IS NULL OR (@online = N'Y' AND CLNOnlineStatus = 'Y') OR (@online <> N'Y' AND (CLNOnlineStatus = '' OR CLNOnlineStatus IS NULL OR CLNOnlineStatus = 'N')))
      AND (@glFormat IS NULL OR CLNGLFormat = CASE WHEN @glFormat = N'M' THEN 'M' ELSE 'H' END)
    ORDER BY CLNCode;
END
GO

-- SaveClinic duplicate check on Add: Address 1 already used (exact text).
CREATE OR ALTER PROCEDURE cis.EntityClinic_AddressExists @add1 nvarchar(200) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM cis.Hdb_CLN WHERE CLNAdd1 = @add1) THEN 1 ELSE 0 END AS bit);
END
GO

-- GenerateCLNSeq (Medix.bas) on HISMaintenance CLNSeq: an existing letter → seq + 1 formatted "000", code = letter + seq; a new letter →
-- row "0001" and code = letter + letter + "001" (the desktop's doubling, kept). Codes already in CLN are skipped (PK_CLNNew).
CREATE OR ALTER PROCEDURE cis.EntityClinic_NextCode @letter nchar(1) AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seq int = (SELECT TOP (1) TRY_CAST(CLNSeq AS int) FROM cis.Maint_CLNSeq WITH (UPDLOCK, HOLDLOCK) WHERE CLNCode = @letter);
    DECLARE @code nvarchar(10);
    IF @seq IS NULL
    BEGIN
        INSERT INTO cis.Maint_CLNSeq (CLNCode, CLNSeq) VALUES (@letter, N'0001');
        SET @code = @letter + @letter + N'001';
        SET @seq = 1;
        WHILE EXISTS (SELECT 1 FROM cis.Hdb_CLN WHERE CLNCode = @code)
        BEGIN
            SET @seq += 1;
            SET @code = @letter + RIGHT(N'000' + CAST(@seq AS nvarchar(10)), CASE WHEN @seq > 999 THEN LEN(CAST(@seq AS nvarchar(10))) ELSE 3 END);
        END
        IF @seq > 1 UPDATE cis.Maint_CLNSeq SET CLNSeq = RIGHT(N'000' + CAST(@seq AS nvarchar(10)), 3) WHERE CLNCode = @letter;
    END
    ELSE
    BEGIN
        SET @seq += 1;
        WHILE EXISTS (SELECT 1 FROM cis.Hdb_CLN WHERE CLNCode = @letter + CASE WHEN @seq > 999 THEN CAST(@seq AS nvarchar(10)) ELSE RIGHT(N'000' + CAST(@seq AS nvarchar(10)), 3) END)
            SET @seq += 1;
        DECLARE @text nvarchar(5) = CASE WHEN @seq > 999 THEN CAST(@seq AS nvarchar(10)) ELSE RIGHT(N'000' + CAST(@seq AS nvarchar(10)), 3) END;
        UPDATE cis.Maint_CLNSeq SET CLNSeq = @text WHERE CLNCode = @letter;
        SET @code = @letter + @text;
    END
    SELECT @code;
END
GO

-- SaveClinic: Add inserts, Edit updates the same columns (CLNPPostCode / CLNPCity / CLNPState are not written, as in the desktop).
CREATE OR ALTER PROCEDURE cis.EntityClinic_Save
    @insert bit, @code nvarchar(5), @name nvarchar(100), @groupCode nvarchar(6), @phone nvarchar(15), @fax nvarchar(15), @add1 nvarchar(200),
    @add2 nvarchar(200), @add3 nvarchar(200), @email nvarchar(1000), @homepage nvarchar(50), @representative nvarchar(200), @docName nvarchar(200),
    @bnfType int = NULL, @forRef nvarchar(50), @noOfDoc int = NULL, @type nchar(1), @xray nchar(1), @panel nchar(1), @hrs24 nchar(1), @ffs nchar(1),
    @ffsPerc numeric(18, 4) = NULL, @state nvarchar(30), @city nvarchar(30), @postCode nvarchar(5), @initial nvarchar(50), @receivedDate datetime,
    @bankAcc nvarchar(25), @bankName nvarchar(50), @bankBrn nvarchar(50), @accPerson nvarchar(50), @remark nvarchar(1500), @location nvarchar(100),
    @operatingHr nvarchar(100), @pName nvarchar(200), @pAdd1 nvarchar(50), @pAdd2 nvarchar(50), @pAdd3 nvarchar(50), @online nchar(1),
    @glFormat nchar(1), @regNo nvarchar(50), @user nvarchar(10), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF @insert = 1
        INSERT INTO cis.Hdb_CLN (CLNCode, CLNName, ClinicGrpCode, CLNPhone, CLNFax, CLNAdd1, CLNAdd2, CLNAdd3, CLNEmail, CLNHomepage, CLNRepresentative,
                                 CLNDocName, CLNBNFType, CLNForRef, CLNNoofDoc, CLNType, CLNXRay, CLNPanel, CLN24Hrs, CLNFFSStatus, CLNMCOProcfeePerc,
                                 CLNState, CLNCity, CLNPostCode, CLNInitial, CLNReceivedDate, CLNBankAcc, CLNBankName, CLNBankBrn, CLNAccPerson,
                                 CLNRemark, CLNLocation, CLNLastUpdateUser, CLNLastUpdateDate, CLNOperatingHr, CLNPName, CLNPAdd1, CLNPAdd2, CLNPAdd3,
                                 CLNOnlineStatus, CLNGLFormat, CLNRegNo)
        VALUES (@code, @name, @groupCode, @phone, @fax, @add1, @add2, @add3, @email, @homepage, @representative, @docName, @bnfType, @forRef, @noOfDoc,
                @type, @xray, @panel, @hrs24, @ffs, @ffsPerc, @state, @city, @postCode, @initial, @receivedDate, @bankAcc, @bankName, @bankBrn,
                @accPerson, @remark, @location, @user, @today, @operatingHr, @pName, @pAdd1, @pAdd2, @pAdd3, @online, @glFormat, @regNo);
    ELSE
        UPDATE cis.Hdb_CLN
        SET CLNName = @name, ClinicGrpCode = @groupCode, CLNPhone = @phone, CLNFax = @fax, CLNAdd1 = @add1, CLNAdd2 = @add2, CLNAdd3 = @add3,
            CLNEmail = @email, CLNHomepage = @homepage, CLNRepresentative = @representative, CLNDocName = @docName, CLNBNFType = @bnfType,
            CLNForRef = @forRef, CLNNoofDoc = CASE WHEN @noOfDoc IS NULL THEN CLNNoofDoc ELSE @noOfDoc END, CLNType = @type, CLNXRay = @xray,
            CLNPanel = @panel, CLN24Hrs = @hrs24, CLNFFSStatus = @ffs, CLNMCOProcfeePerc = @ffsPerc, CLNState = @state, CLNCity = @city,
            CLNPostCode = @postCode, CLNInitial = @initial, CLNReceivedDate = @receivedDate, CLNBankAcc = @bankAcc, CLNBankName = @bankName,
            CLNBankBrn = @bankBrn, CLNAccPerson = @accPerson, CLNRemark = @remark, CLNLocation = @location, CLNLastUpdateUser = @user,
            CLNLastUpdateDate = @today, CLNOperatingHr = @operatingHr, CLNPName = @pName, CLNPAdd1 = @pAdd1, CLNPAdd2 = @pAdd2, CLNPAdd3 = @pAdd3,
            CLNOnlineStatus = @online, CLNGLFormat = @glFormat, CLNRegNo = @regNo
        WHERE CLNCode = @code;
    SELECT @@ROWCOUNT;
END
GO

-- DeleteClinic: refused while CISCAS has cases of the clinic.
CREATE OR ALTER PROCEDURE cis.EntityClinic_HasCases @code nvarchar(5) AS
BEGIN
    SET NOCOUNT ON;
    SELECT CAST(CASE WHEN EXISTS (SELECT 1 FROM dbo.CISCAS WHERE ClinicCode = @code) THEN 1 ELSE 0 END AS bit);
END
GO

CREATE OR ALTER PROCEDURE cis.EntityClinic_Delete @code nvarchar(5) AS
BEGIN
    SET NOCOUNT ON;
    DELETE cis.Hdb_CLN WHERE CLNCode = @code;
    SELECT @@ROWCOUNT;
END
GO
