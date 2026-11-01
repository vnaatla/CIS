-- Invoice > Other > Case Transfer (frmMembershipTransfer): move a case (CISDB CISCAS) to another member / cover.

-- ListCases: the case's current member (VIEW_CISMBMTransfer).
CREATE OR ALTER PROCEDURE cis.CaseTransfer_Case @casNumber nvarchar(20) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP (1) CASNumber, RTRIM(MBMNumber) AS MbmNumber, RTRIM(CoverID) AS CoverId, RTRIM(MBMName) AS MbmName, RTRIM(MBMCName) AS MbmcName,
           DateVisit
    FROM dbo.VIEW_CISMBMTransfer WHERE CASNumber = @casNumber;
END
GO

-- cmdSearch_Click: the transfer member's name (cover 00: HISDB MBM; otherwise MBMCoveredPersons).
CREATE OR ALTER PROCEDURE cis.CaseTransfer_MemberName @mbmNumber nvarchar(50), @covId nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    IF @covId = '00'
        SELECT TOP (1) RTRIM(MBMName) AS Name FROM cis.Hdb_MBM WHERE MBMNumber = @mbmNumber;
    ELSE
        SELECT TOP (1) RTRIM(MBMCName) AS Name FROM cis.Hdb_MBMCoveredPersons WHERE MBMCNumber = @mbmNumber AND MBMCCoverID = @covId;
END
GO

-- cmdSave_Click: the case's member number and cover id.
CREATE OR ALTER PROCEDURE cis.CaseTransfer_Save @casNumber nvarchar(20), @mbmNumber nvarchar(50), @covId nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.CISCAS SET MBMnumber = @mbmNumber, CoverID = @covId WHERE CASNumber = @casNumber;
    SELECT @@ROWCOUNT;
END
GO
