-- Invoice pop-ups shared by the invoice screens: ICD / treatment reason search (frmSearchICD) and case list (frmCasList).

-- frmSearchICD cmdSearch_Click: code starts with, description contains, and a second description joined with AND / OR.
-- The desktop appended "AND/OR ReasonDesc like ..." without brackets, so with OR the other filters only apply to the first part; kept.
CREATE OR ALTER PROCEDURE cis.Popup_Icd @code nvarchar(20) = NULL, @desc nvarchar(200) = NULL, @desc2 nvarchar(200) = NULL, @useOr bit = 0 AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(ReasonCode) AS ReasonCode, RTRIM(ReasonDesc) AS ReasonDesc
    FROM dbo.TreatmentReason
    WHERE ((@code IS NULL OR ReasonCode LIKE @code + '%') AND (@desc IS NULL OR ReasonDesc LIKE '%' + @desc + '%')
           AND (@desc2 IS NULL OR @useOr = 1 OR ReasonDesc LIKE '%' + @desc2 + '%'))
       OR (@desc2 IS NOT NULL AND @useOr = 1 AND ReasonDesc LIKE '%' + @desc2 + '%');
END
GO

-- frmCasList ListCases: cases of active members (VIEW_Cas_MBM), invoice (C) or dental (D) cases; @by: 1 member no, 2 policy no, 3 name, 4 IC,
-- 5 case no. The clinic name comes from panel clinics (C / Y).
CREATE OR ALTER PROCEDURE cis.Popup_CaseList @caseType char(1), @by int, @text nvarchar(100) AS
BEGIN
    SET NOCOUNT ON;
    SELECT v.CASNumber, v.MBMnumber AS MBMNumber, v.CoverID, v.MBMName, v.MBMCName, v.MBMPolicyNo, v.DateVisit, v.CASStatus, v.ClinicCode,
           RTRIM(c.CLNName) AS CLNName, v.MBMIcBcPp
    FROM dbo.VIEW_Cas_MBM v
    LEFT JOIN cis.Hdb_CLN c ON c.CLNCode = v.ClinicCode AND (c.CLNPanel = 'C' OR c.CLNPanel = 'Y')
    WHERE v.MBMCLNDataStatus <> 'D' AND v.MBMCLNStatus <> 'D' AND v.HSCLNInd IN ('C', 'M', 'B') AND v.CaseType = @caseType
      AND ((@by = 1 AND (v.MBMnumber LIKE @text + '%' OR v.MBMCNumber LIKE @text + '%'))
           OR (@by = 2 AND v.MBMPolicyNo LIKE @text + '%')
           OR (@by = 3 AND v.MBMName LIKE '%' + @text + '%')
           OR (@by = 4 AND v.MBMIcBcPp LIKE @text + '%')
           OR (@by = 5 AND v.CASNumber LIKE @text + '%'));
END
GO
