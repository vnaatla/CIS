-- Claim > Pass to Claim (frmPass2Claim): list cases of CISWorksheet VIEW_ClaimPass2Claim (the desktop's ClaimPass2Claim case 1, parameterised)
-- and pass the checked ones (ClaimPass2ClaimUpdate → ClaimFileTransfer).

-- chkPass_Click: the users who passed cases (HISDB SearchCASToken).
CREATE OR ALTER PROCEDURE cis.Pass2Claim_Users AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Hdb_SearchCASToken;
END
GO

-- SearchRecord → DischargeListing. @caseNo = the scanned case (scan mode), else @casFrom / @casTo; @caseStatus 0 O, 1 D, 2 O or C;
-- @passed 'P' passed / 'N' not passed; @noVisit / @noRegister / @noPass = the date is empty; @casNumbers = the cases just passed
-- (comma-separated, the list shown after Pass).
CREATE OR ALTER PROCEDURE cis.Pass2Claim_Search
    @caseNo nvarchar(20) = NULL, @casFrom nvarchar(20) = NULL, @casTo nvarchar(20) = NULL, @caseStatus char(1) = NULL, @claimability char(1) = NULL,
    @passBy nvarchar(50) = NULL, @passed char(1) = NULL, @noVisit bit = 0, @noRegister bit = 0, @noPass bit = 0, @regFrom datetime = NULL,
    @regTo datetime = NULL, @visitFrom datetime = NULL, @visitTo datetime = NULL, @passFrom datetime = NULL, @passTo datetime = NULL,
    @casNumbers nvarchar(max) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SELECT ClaimStatus, CASNumber, MBMnumber AS MBMNumber, PatientName, RegisterDate, DateVisit, ClaimPassDate, CLNName, OPBatchNo, ClaimPassBy,
           CASStatus, Claimability, CoverID, MBMCoverID, NoofDay
    FROM cis.Ws_VIEW_ClaimPass2Claim
    WHERE (@caseNo IS NULL OR CASNumber = @caseNo) AND (@casFrom IS NULL OR CASNumber >= @casFrom) AND (@casTo IS NULL OR CASNumber <= @casTo)
      AND (@caseStatus IS NULL OR (@caseStatus = '0' AND CASStatus = 'O') OR (@caseStatus = '1' AND CASStatus = 'D')
           OR (@caseStatus = '2' AND (CASStatus = 'O' OR CASStatus = 'C')))
      AND (@claimability IS NULL OR Claimability = @claimability) AND (@passBy IS NULL OR ClaimPassBy = @passBy)
      AND (@passed IS NULL OR (@passed = 'P' AND ClaimStatus = 'P') OR (@passed <> 'P' AND (ClaimStatus IS NULL OR ClaimStatus = '')))
      AND (@noVisit = 0 OR DateVisit IS NULL) AND (@noRegister = 0 OR RegisterDate IS NULL) AND (@noPass = 0 OR ClaimPassDate IS NULL)
      AND (@regFrom IS NULL OR RegisterDate >= @regFrom) AND (@regTo IS NULL OR RegisterDate <= @regTo)
      AND (@visitFrom IS NULL OR DateVisit >= @visitFrom) AND (@visitTo IS NULL OR DateVisit <= @visitTo)
      AND (@passFrom IS NULL OR ClaimPassDate >= @passFrom) AND (@passTo IS NULL OR ClaimPassDate <= @passTo)
      AND (@casNumbers IS NULL OR ',' + @casNumbers + ',' LIKE '%,' + CASNumber + ',%')
    GROUP BY ClaimStatus, CASNumber, MBMnumber, PatientName, RegisterDate, DateVisit, ClaimPassDate, CLNName, OPBatchNo, ClaimPassBy, CASStatus,
             Claimability, CoverID, MBMCoverID, NoofDay
    OPTION (RECOMPILE);
END
GO

-- UpdateRecord: one checked case (the desktop's ClaimPass2ClaimUpdate, status P).
CREATE OR ALTER PROCEDURE cis.Pass2Claim_Pass @casNumber nvarchar(10), @userCode nvarchar(10), @today datetime AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Ws_ClaimPass2ClaimUpdate @CASNumber = @casNumber, @ClaimStatus = 'P', @ClaimPassBy = @userCode, @ClaimPassDate = @today;
END
GO
