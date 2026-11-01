-- Clinic member registration pop-up (frmMemberRegistrationCLN, opened from the clinic invoice): look-ups, limits, age band and the new
-- DC member (HISDB MBM, MBMTwo, MBMOthers, MBMOthersTwo, MBMCLNOthers; HISMaintenance MBMCrossReference, MBMSEQ).

-- ComboListing: races (HISDB SearchRace), staff categories (CISDB WebCategory), group company names (CISDB GRPCompany).
CREATE OR ALTER PROCEDURE cis.MbmReg_Setup AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Hdb_SearchRace;
    SELECT RTRIM(Category) AS Category, RTRIM(INSType) AS InsType, RTRIM(PlanCode) AS PlanCode FROM dbo.WebCategory;
    SELECT RTRIM(GRPCoName) AS Name FROM dbo.GRPCompany WHERE GRPCoName IS NOT NULL;
END
GO

-- cmdSearch_Click: departments of group company G001 containing the text.
CREATE OR ALTER PROCEDURE cis.MbmReg_Departments @text nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(GrpCompDept) AS Department FROM dbo.GroupCompanyDept WHERE GrpCompCode = 'G001' AND GrpCompDept LIKE N'%' + @text + N'%';
END
GO

-- CallCLNLimit: annual limit (PLN.ANNLmtInd N → 0, else GetCLNAnnualLimit) and visit limit (PLN.VISLmtInd N → 0, else GetVisitLimit:
-- the DC plan row with VISLmtInd N = "FULL" (NULL here), else the latest VisitLimit), health type N, principal.
CREATE OR ALTER PROCEDURE cis.MbmReg_Limits @plnCode nvarchar(4), @insCode nvarchar(2), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @ann money = 0, @visit money = 0, @full bit = 0;
    IF ISNULL((SELECT TOP (1) RTRIM(ANNLmtInd) FROM dbo.PLN WHERE PLNCode = @plnCode), '') <> 'N'
        SET @ann = ISNULL((SELECT TOP (1) AnnLimit FROM dbo.AnnualLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND HLTCode = 'N'
                           AND AnnEffDate <= @effDate ORDER BY AnnEffDate DESC), 0);
    IF EXISTS (SELECT 1 FROM dbo.PLN WHERE PLNCode = @plnCode) AND ISNULL((SELECT TOP (1) RTRIM(VISLmtInd) FROM dbo.PLN WHERE PLNCode = @plnCode), '') <> 'N'
    BEGIN
        IF EXISTS (SELECT 1 FROM dbo.PLN WHERE PLNCode = @plnCode AND PAYCode = 'DC' AND RTRIM(VISLmtInd) = 'N') SET @full = 1;
        ELSE SET @visit = ISNULL((SELECT TOP (1) VisitLimit FROM dbo.VisitLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND HLTCode = 'N'
                                  AND VisitEffDate <= @effDate ORDER BY VisitEffDate DESC), 0);
    END
    SELECT @ann AS AnnualLimit, @visit AS VisitLimit, @full AS VisitFull;
END
GO

-- GetAgeBand: years against AGEFrom / AGETo ("…Y"), days against AGEFrom ("…D"), HISDB AGE.
CREATE OR ALTER PROCEDURE cis.MbmReg_AgeBand @age int, @days bit AS
BEGIN
    SET NOCOUNT ON;
    IF @days = 0
        SELECT TOP (1) RTRIM(AGECode) AS AgeCode, RTRIM(AGEDescription) AS AgeDescription FROM cis.Hdb_AGE
        WHERE TRY_CAST(SUBSTRING(AGEFrom, 1, LEN(AGEFrom) - 1) AS int) <= @age AND TRY_CAST(SUBSTRING(AGETo, 1, LEN(AGETo) - 1) AS int) >= @age
        ORDER BY AGECode;
    ELSE
        SELECT TOP (1) RTRIM(AGECode) AS AgeCode, RTRIM(AGEDescription) AS AgeDescription FROM cis.Hdb_AGE
        WHERE TRY_CAST(SUBSTRING(AGEFrom, 1, LEN(AGEFrom) - 1) AS int) <= @age ORDER BY AGEFrom;
END
GO

-- cmdReg_Click: GenerateMembershipNo("DC", insured type) + GenerateChkDigits, then the six inserts; one transaction (the caller's).
CREATE OR ALTER PROCEDURE cis.MbmReg_Register
    @ic nvarchar(30), @name nvarchar(200), @sex nvarchar(1), @adultChild nvarchar(1), @ageCode nvarchar(2), @dob datetime, @race nvarchar(5),
    @groupCompany nvarchar(200), @employeeNo nvarchar(50), @department nvarchar(200), @plnCode nvarchar(4), @insCode nvarchar(2),
    @effDate datetime, @expDate datetime, @visitLimit money, @annualLimit money, @user nvarchar(50), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @seqCode nvarchar(10) = N'DC' + @insCode, @seq nvarchar(10), @ch nchar(1), @n int, @num nvarchar(10);
    SELECT @seq = RTRIM(SEQNumber) FROM cis.Maint_MBMSEQ WITH (UPDLOCK, HOLDLOCK) WHERE SEQCode = @seqCode;
    IF @seq IS NULL
    BEGIN
        SET @num = N'A00001';
        INSERT INTO cis.Maint_MBMSEQ (SEQCode, PAYCode, INSCode, SEQNumber) VALUES (@seqCode, N'DC', @insCode, @num);
    END
    ELSE
    BEGIN
        SELECT @ch = LEFT(@seq, 1), @n = TRY_CAST(SUBSTRING(@seq, 2, 5) AS int);
        IF @n < 99999 SET @n += 1;
        ELSE SELECT @ch = NCHAR(UNICODE(@ch) + 1), @n = 1;
        SET @num = @ch + RIGHT(N'00000' + CAST(@n AS nvarchar(10)), 5);
        UPDATE cis.Maint_MBMSEQ SET SEQNumber = @num WHERE SEQCode = @seqCode;
    END

    DECLARE @d nvarchar(2) = (SELECT TOP (1) RTRIM(MBMChkDigits) FROM cis.Hdb_MBMCheckDigits WITH (UPDLOCK, HOLDLOCK));
    SET @d = CASE WHEN @d = N'99' OR TRY_CAST(@d AS int) IS NULL THEN N'01' ELSE RIGHT(N'00' + CAST(CAST(@d AS int) + 1 AS nvarchar(3)), 2) END;
    UPDATE cis.Hdb_MBMCheckDigits SET MBMChkDigits = @d;

    DECLARE @mbm nvarchar(30) = N'DC' + @insCode + @num + @d;
    INSERT INTO cis.Hdb_MBM (MBMNumber, MBMIcBcPp, MBMName, MBMSex, MBMAdultChild, AGECode, MBMValueDate, MBMBordxDate, MBMMemberType, HLTCode,
                             MBMBatchNo, MBMCOVID, HSCLNInd, MBMExpiredStatus, MBMDOB, RACCode)
    VALUES (@mbm, @ic, @name, @sex, @adultChild, @ageCode, @today, @today, 'P', 'N', '1', '00', 'C', 'N', @dob, @race);
    INSERT INTO cis.Hdb_MBMTwo (MBMNumber, MBMSalutation, MBMTelnoH, MBMTelNoM, MBMTelNoO, MBMEmail, NATCode, SRVCodeList, LGNCode, MBMRegUser,
                                MBMLastUpdateUser, MBMUploadStatus)
    VALUES (@mbm, '', '', '', '', '', '', 'ME', '', @user, @user, 'M');
    INSERT INTO cis.Hdb_MBMOthers (MBMNumber, MBMGroupCompany, MBMEmployeeNo, MBMDepartment, RELCODE) VALUES (@mbm, @groupCompany, @employeeNo, @department, 'P');
    INSERT INTO cis.Hdb_MBMOthersTwo (MBMNumber) VALUES (@mbm);
    INSERT INTO cis.Maint_MBMCrossReference (MBMNumber, MBMIcBcPp, MBMName, MBMAdultChild, AGECode, MBMBordxDate, MBMMemberType, HLTCode, MBMBatchNo,
                                             MBMCOVID, MBMStatus)
    VALUES (@mbm, @ic, @name, @adultChild, @ageCode, @today, 'P', 'N', '1', '00', 'A');
    INSERT INTO cis.Hdb_MBMCLNOthers (MBMNumber, PAYCode, PLNType, InsType, MBMPolEffDate, MBMPolExpDate, MBMCLNMemberType, MBMCLNExpiryStatus,
                                      MBMCLNDataStatus, MBMCLNStatus, MBMCLNVisitLmt, MBMCLNBalLmt)
    VALUES (@mbm, 'DC', @plnCode, @insCode, @effDate, @expDate, 'P', 'N', 'A', 'A', @visitLimit, @annualLimit);
    SELECT @mbm;
END
GO
