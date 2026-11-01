-- Membership > MBM Topup Adj (FrmMBMTopupLmt): member (HISDB SearchCISMemberProc, MBMOthers), balance (CISDB spUtilisationOP_test1),
-- supplementary limit type (CISDB AnnualLimit + HISDB AnnualLimitType), history (CISDB MBMCISTopupHistory) and the limit / IC updates
-- (HISDB MBMCLNOthers, MBMCoveredPersonsCLN, MBM, MBMCoveredPersons; HISMaintenance MBMCrossReference).

-- FindMBM: group company and department (MBMOthers).
CREATE OR ALTER PROCEDURE cis.MbmTopup_Others @mbmNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) RTRIM(MBMGroupCompany) AS GroupCompany, RTRIM(MBMDepartment) AS Department FROM cis.Hdb_MBMOthers WHERE MBMNumber = @mbmNumber;
END
GO

-- GetSptLmtInd: SuppLimitStatus of the latest AnnualLimit on or before the effective date (blank = A) and its description.
CREATE OR ALTER PROCEDURE cis.MbmTopup_LimitType @plnCode nvarchar(10), @insCode nvarchar(5), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) ISNULL(NULLIF(RTRIM(a.SuppLimitStatus), ''), 'A') AS LimitCode,
           (SELECT TOP (1) RTRIM(t.LimitDesc) FROM cis.Hdb_AnnualLimitType t WHERE t.LimitCode = ISNULL(NULLIF(RTRIM(a.SuppLimitStatus), ''), 'A')) AS LimitDesc
    FROM dbo.AnnualLimit a
    WHERE a.PLNCode = @plnCode AND a.INSCode = @insCode AND a.AnnEffDate <= @effDate
    ORDER BY a.AnnEffDate DESC;
END
GO

-- CmdNext_Click / CmdPreview_Click: the next / previous clinic member number (HSCLNInd <> 'H').
CREATE OR ALTER PROCEDURE cis.MbmTopup_Step @mbmNumber nvarchar(30), @next bit AS
BEGIN
    SET NOCOUNT ON;
    IF @next = 1
        SELECT TOP (1) MBMNumber FROM cis.Hdb_MBM WHERE HSCLNInd <> 'H' AND MBMNumber > @mbmNumber ORDER BY MBMNumber;
    ELSE
        SELECT TOP (1) MBMNumber FROM cis.Hdb_MBM WHERE HSCLNInd <> 'H' AND MBMNumber < @mbmNumber ORDER BY MBMNumber DESC;
END
GO

-- CmdSave_Click: history row, IC change (GP only), new balance limit; one transaction (the caller's).
CREATE OR ALTER PROCEDURE cis.MbmTopup_Save
    @mbmNumber nvarchar(15), @covId char(2), @gp bit, @icNo nvarchar(30), @prevIcNo nvarchar(30), @prevLimit money, @topupAmt money,
    @newLimit money, @limitCode varchar(1), @user nvarchar(50), @today datetime
AS
BEGIN
    SET NOCOUNT ON;
    IF @gp = 1
        INSERT INTO dbo.MBMCISTopupHistory (MBMNumber, MBMCovID, MBMICNo, PrevICNo, PrevLimit, TopupAmt, NewLimit, UpdateDate, UpdateBy)
        VALUES (@mbmNumber, @covId, @icNo, @prevIcNo, @prevLimit, @topupAmt, @newLimit, @today, @user);
    ELSE
        INSERT INTO dbo.MBMCISTopupHistory (MBMNumber, MBMCovID, PrevLimit, SP_TopupAmt, NewLimit, UpdateDate, UpdateBy)
        VALUES (@mbmNumber, @covId, @prevLimit, @topupAmt, @newLimit, @today, @user);

    IF @gp = 1 AND ISNULL(RTRIM(@prevIcNo), N'') <> ISNULL(RTRIM(@icNo), N'')
    BEGIN
        IF @covId = '00'
        BEGIN
            UPDATE cis.Hdb_MBM SET MBMIcBcPp = @icNo WHERE MBMNumber = @mbmNumber;
            UPDATE cis.Maint_MBMCrossReference SET MBMIcBcPp = @icNo WHERE MBMNumber = @mbmNumber;
        END
        ELSE
            UPDATE cis.Hdb_MBMCoveredPersons SET MBMCICBCPPNo = @icNo WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @covId;
    END

    IF @newLimit <> 0
    BEGIN
        IF @limitCode = 'A' OR @covId = '00'
            UPDATE cis.Hdb_MBMCLNOthers SET MBMCLNBalLmt = @newLimit WHERE MBMNumber = @mbmNumber;
        ELSE IF @limitCode <> 'N'
            UPDATE cis.Hdb_MBMCoveredPersonsCLN SET MBMCCLNBalLmt = @newLimit
            WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = CASE WHEN @limitCode = 'B' THEN '01' ELSE @covId END;
    END
END
GO
