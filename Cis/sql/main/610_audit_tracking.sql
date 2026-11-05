-- LG & Audit > Audit Tracking (FrmAuditInvoiceTracking): Sp_Audittraccking_2018's report with the desktop's filters. The desktop pasted the
-- filter text into the procedure; here every filter is a parameter. The filter clauses are added as text (as the desktop did, so unqualified
-- columns resolve the same way) but all values go through sp_executesql parameters.
-- Reproduced quirks: the claimability filter and the FFS "N" filter add an unbracketed OR; cases without any medication are not listed; the
-- fixed condition CASStatus NOT IN ('D','C') means the C / D status filter finds nothing.
CREATE OR ALTER PROCEDURE cis.AuditTrack_Search
    @caseFrom nvarchar(20) = NULL, @caseTo nvarchar(20) = NULL, @caseBlank bit = 0,
    @payTo nvarchar(2) = NULL, @payToBlank bit = 0, @clinic nvarchar(20) = NULL, @clinicBlank bit = 0,
    @payor nvarchar(10) = NULL, @payorBlank bit = 0, @groupCompany nvarchar(200) = NULL, @groupBlank bit = 0,
    @patientName nvarchar(200) = NULL, @nameBlank bit = 0, @mbmFrom nvarchar(50) = NULL, @mbmTo nvarchar(50) = NULL, @mbmBlank bit = 0,
    @invFrom nvarchar(50) = NULL, @invTo nvarchar(50) = NULL, @invBlank bit = 0,
    @bordxFrom nvarchar(30) = NULL, @bordxTo nvarchar(30) = NULL, @bordxBlank bit = 0,
    @caseStatus char(1) = NULL, @claimability char(1) = NULL,
    @planFrom nvarchar(20) = NULL, @planTo nvarchar(20) = NULL, @planBlank bit = 0,
    @bordxDateFrom datetime = NULL, @bordxDateTo datetime = NULL,
    @icFrom nvarchar(50) = NULL, @icTo nvarchar(50) = NULL, @icBlank bit = 0,
    @visitFrom datetime = NULL, @visitTo datetime = NULL, @visitBlank bit = 0,
    @regFrom datetime = NULL, @regTo datetime = NULL, @regBlank bit = 0,
    @invDateFrom datetime = NULL, @invDateTo datetime = NULL, @invDateBlank bit = 0,
    @enteredFrom datetime = NULL, @enteredTo datetime = NULL, @enteredBlank bit = 0,
    @amtFrom money = NULL, @amtTo money = NULL, @amtBlank bit = 0,
    @monthFrom nvarchar(4) = NULL, @monthTo nvarchar(4) = NULL, @monthBlank bit = 0,
    @yearFrom nvarchar(4) = NULL, @yearTo nvarchar(4) = NULL, @yearBlank bit = 0,
    @clinicType char(1) = NULL, @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @opBlank bit = 0,
    @ffs char(1) = NULL, @omni bit = 0, @tims bit = 0, @subsidiary nvarchar(200) = NULL, @department nvarchar(200) = NULL,
    @omniRefFrom nvarchar(50) = NULL, @omniRefTo nvarchar(50) = NULL, @omniRefBlank bit = 0, @omniSent char(1) = NULL, @audited bit = 0
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @w nvarchar(max) = N'';
    IF @caseFrom IS NOT NULL SET @w += N' AND ciscas.CASNumber >= @caseFrom';
    IF @caseTo IS NOT NULL SET @w += N' AND ciscas.CASNumber <= @caseTo';
    IF @caseBlank = 1 SET @w += N' AND (ciscas.CASNumber Is null OR ciscas.CASNumber = '''')';
    IF @payTo IS NOT NULL SET @w += N' AND INVPayTo = @payTo';
    IF @payToBlank = 1 SET @w += N' AND (INVPayTo Is null OR INVPayTo = '''')';
    IF @clinic IS NOT NULL SET @w += N' AND INVPayee = @clinic';
    IF @clinicBlank = 1 SET @w += N' AND (INVPayee Is null OR INVPayee = '''')';
    IF @payor IS NOT NULL SET @w += N' AND MBMCLNOthers.PAYCode = @payor';
    IF @payorBlank = 1 SET @w += N' AND (MBMCLNOthers.PAYCode Is null OR MBMCLNOthers.PAYCode = '''')';
    IF @groupCompany IS NOT NULL SET @w += N' AND MBMGroupCompany = @groupCompany';
    IF @groupBlank = 1 SET @w += N' AND (MBMGroupCompany IS NULL OR MBMGroupCompany = '''')';
    IF @patientName IS NOT NULL SET @w += N' AND mbm.mbmname = @patientName';
    IF @nameBlank = 1 SET @w += N' AND (mbm.mbmname Is null OR mbm.mbmname = '''')';
    IF @mbmFrom IS NOT NULL SET @w += N' AND CISCas.MBMNumber >= @mbmFrom';
    IF @mbmTo IS NOT NULL SET @w += N' AND CISCas.MBMNumber <= @mbmTo';
    IF @mbmBlank = 1 SET @w += N' AND (CISCas.MBMNumber Is null OR CISCas.MBMNumber = '''')';
    IF @invFrom IS NOT NULL SET @w += N' AND CISCas.InvoiceNo >= @invFrom';
    IF @invTo IS NOT NULL SET @w += N' AND CISCas.InvoiceNo <= @invTo';
    IF @invBlank = 1 SET @w += N' AND (CISCas.InvoiceNo Is null OR CISCas.InvoiceNo = '''')';
    IF @bordxFrom IS NOT NULL SET @w += N' AND Claim.BordxRefNo >= @bordxFrom';
    IF @bordxTo IS NOT NULL SET @w += N' AND Claim.BordxRefNo <= @bordxTo';
    IF @bordxBlank = 1 SET @w += N' AND (Claim.BordxRefNo Is null OR Claim.BordxRefNo = '''')';
    IF @caseStatus IN ('O', 'C', 'D') SET @w += N' AND CISCas.CASStatus = @caseStatus';
    IF @claimability IN ('A', 'R') SET @w += N' Or CISCas.Claimability = @claimability';
    IF @planFrom IS NOT NULL SET @w += N' AND MBMCLNOthers.PLNType >= @planFrom';
    IF @planTo IS NOT NULL SET @w += N' AND MBMCLNOthers.PLNType <= @planTo';
    IF @planBlank = 1 SET @w += N' AND (MBMCLNOthers.PLNType Is null OR MBMCLNOthers.PLNType = '''')';
    IF @bordxDateFrom IS NOT NULL SET @w += N' AND Bordx.BordxDate >= @bordxDateFrom';
    IF @bordxDateTo IS NOT NULL SET @w += N' AND Bordx.BordxDate <= @bordxDateTo';
    IF @icFrom IS NOT NULL SET @w += N' AND mbm.MBMIcBcPp >= @icFrom';
    IF @icTo IS NOT NULL SET @w += N' AND mbm.MBMIcBcPp <= @icTo';
    IF @icBlank = 1 SET @w += N' AND (mbm.MBMIcBcPp Is null)';
    IF @visitFrom IS NOT NULL SET @w += N' AND CISCas.DateVisit >= @visitFrom';
    IF @visitTo IS NOT NULL SET @w += N' AND CISCas.DateVisit <= @visitTo';
    IF @visitBlank = 1 SET @w += N' AND (CISCas.DateVisit Is null)';
    IF @regFrom IS NOT NULL SET @w += N' AND CISCas.RegisterDate >= @regFrom';
    IF @regTo IS NOT NULL SET @w += N' AND CISCas.RegisterDate <= @regTo';
    IF @regBlank = 1 SET @w += N' AND (CISCas.RegisterDate Is null)';
    IF @invDateFrom IS NOT NULL SET @w += N' AND CISCas.InvoiceDate >= @invDateFrom';
    IF @invDateTo IS NOT NULL SET @w += N' AND CISCas.InvoiceDate <= @invDateTo';
    IF @invDateBlank = 1 SET @w += N' AND (CISCas.InvoiceDate Is null)';
    IF @enteredFrom IS NOT NULL SET @w += N' AND ClaimInvoice.INVEnteredDate >= @enteredFrom';
    IF @enteredTo IS NOT NULL SET @w += N' AND ClaimInvoice.INVEnteredDate <= @enteredTo';
    IF @enteredBlank = 1 SET @w += N' AND (ClaimInvoice.INVEnteredDate Is null)';
    IF @amtFrom IS NOT NULL SET @w += N' AND ClaimInvoice.INVAmount >= @amtFrom';
    IF @amtTo IS NOT NULL SET @w += N' AND ClaimInvoice.INVAmount <= @amtTo';
    IF @amtBlank = 1 SET @w += N' AND (ClaimInvoice.INVAmount Is null)';
    IF @monthFrom IS NOT NULL SET @w += N' AND AccPeriodMonth >= @monthFrom';
    IF @monthTo IS NOT NULL SET @w += N' AND AccPeriodMonth <= @monthTo';
    IF @monthBlank = 1 SET @w += N' AND AccPeriodMonth Is null';
    IF @yearFrom IS NOT NULL SET @w += N' AND AccPeriodYear >= @yearFrom';
    IF @yearTo IS NOT NULL SET @w += N' AND AccPeriodYear <= @yearTo';
    IF @yearBlank = 1 SET @w += N' AND AccPeriodYear Is null';
    IF @clinicType = '0' SET @w += N' AND CISCas.CLNType = ''S''';
    IF @clinicType = '1' SET @w += N' AND CISCas.CLNType = ''G''';
    IF @clinicType = '2' SET @w += N' AND (CISCas.CLNType = '''' OR CISCas.CLNType IS NULL)';
    IF @opFrom IS NOT NULL SET @w += N' AND OPBatchNo >= @opFrom';
    IF @opTo IS NOT NULL SET @w += N' AND OPBatchNo <= @opTo';
    IF @opBlank = 1 SET @w += N' AND (OPBatchNo Is null OR OPBatchNo = '''')';
    IF @ffs = 'Y' SET @w += N' AND CLNFFSStatus = ''Y''';
    IF @ffs = 'N' SET @w += N' AND CLNFFSStatus = ''N'' OR (CLNFFSStatus = '''' OR CLNFFSStatus IS NULL)';
    IF @omni = 1 SET @w += N' AND GRPOMNICLM = ''Y''';
    IF @tims = 1 SET @w += N' AND GRPTIMS = ''Y''';
    IF @subsidiary IS NOT NULL SET @w += N' AND MBMOthers.Subsidiary = @subsidiary';
    IF @department IS NOT NULL SET @w += N' AND MBMOthers.MBMDepartment = @department';
    IF @omniRefBlank = 1 SET @w += N' AND (OMNIRefNo Is null OR OMNIRefNo = '''')';
    IF @omniRefFrom IS NOT NULL SET @w += N' AND OMNIRefNo >= @omniRefFrom';
    IF @omniRefTo IS NOT NULL SET @w += N' AND OMNIRefNo <= @omniRefTo';
    IF @omniSent = 'Y' SET @w += N' AND CLMOMNISent = ''Y''';
    IF @omniSent = 'N' SET @w += N' AND CLMOMNISent IS NULL';
    IF @audited = 1 SET @w += N' AND ciscas.AuditQuery IS NOT NULL';

    DECLARE @sql nvarchar(max) = N'
SELECT CASE WHEN MBMCoveredPersons.MBMCICBCPPNo IS NULL THEN mbm.MBMIcBcPp ELSE MBMCoveredPersons.MBMCICBCPPNo END AS MBMIcBcPp,
    CASE WHEN MBMCoveredPersons.MBMCName IS NULL THEN mbm.MBMName ELSE MBMCoveredPersons.MBMCName END AS MBMName,
    ciscas.CASNumber, ciscas.DateVisit,
    CASE WHEN CISCas.CLNType = ''S'' THEN COALESCE(ClaimInvoiceActualOthers.SPFee, 0) ELSE COALESCE(ClaimInvoiceActualOthers.GPFee, 0) END AS SPFee,
    COALESCE(ClaimInvoice.INVAmount, 0) AS INVAmount, LTRIM(RTRIM(CISCas.CasRemark)) AS CasRemark, LTRIM(RTRIM(CISCas.AuditQuery)) AS AuditQuery,
    LTRIM(RTRIM(CISCas.AuditAnswer)) AS AuditAnswer, LTRIM(RTRIM(CISCas.AuditRemarks)) AS AuditRemarks, CISCas.AuditDiscountReceived,
    LTRIM(RTRIM(CISCas.AuditBy)) AS AuditBy, CISCas.AuditDate, CLN.CLNCode, LTRIM(RTRIM(CLN.CLNName)) AS CLNName,
    UPPER(LTRIM(RTRIM(MBMOthers.MBMGroupCompany))) AS MBMGroupCompany,
    CASE CASTreatReason1 WHEN '''' THEN '''' ELSE (SELECT ReasonDesc FROM dbo.TreatmentReason WHERE ReasonCode = CASTreatReason1) END AS ReasonDesc1,
    CASE CASTreatReason2 WHEN '''' THEN '''' ELSE (SELECT ReasonDesc FROM dbo.TreatmentReason WHERE ReasonCode = CASTreatReason2) END AS ReasonDesc2,
    CASE CASTreatReason3 WHEN '''' THEN '''' ELSE (SELECT ReasonDesc FROM dbo.TreatmentReason WHERE ReasonCode = CASTreatReason3) END AS ReasonDesc3,
    CASE CASTreatReason4 WHEN '''' THEN '''' ELSE (SELECT ReasonDesc FROM dbo.TreatmentReason WHERE ReasonCode = CASTreatReason4) END AS ReasonDesc4,
    CASE CASTreatReason5 WHEN '''' THEN '''' ELSE (SELECT ReasonDesc FROM dbo.TreatmentReason WHERE ReasonCode = CASTreatReason5) END AS ReasonDesc5,
    CASDiagnosis5,
    ciscas.MedInjType1, MedInjDays1, MedInjQty1, CLMMedCharges.MedCharges1, ciscas.MedInjType2, MedInjDays2, MedInjQty2, CLMMedCharges.MedCharges2,
    ciscas.MedInjType3, MedInjDays3, MedInjQty3, CLMMedCharges.MedCharges3, ciscas.MedInjType4, MedInjDays4, MedInjQty4, CLMMedCharges.MedCharges4,
    ciscas.MedInjType5, MedInjDays5, MedInjQty5, CLMMedCharges.MedCharges5, ciscas.MedInjType6, MedInjDays6, MedInjQty6, CLMMedCharges.MedCharges6,
    ciscas.MedInjType7, MedInjDays7, MedInjQty7, CLMMedCharges.MedCharges7,
    Bordx.BordxRefNo,
    CASE WHEN CISCas.CLNType = ''S'' THEN COALESCE(ClaimInvoiceActualOthers.SPInj, 0) ELSE COALESCE(ClaimInvoiceActualOthers.GPInj, 0) END AS Procedurefee
INTO #p
FROM dbo.CISCas
INNER JOIN cis.Hdb_MBMCLNOthers MBMCLNOthers WITH (NOLOCK) ON MBMCLNOthers.MBMNumber = CISCas.MBMnumber
INNER JOIN cis.Hdb_MBMOthers MBMOthers WITH (NOLOCK) ON MBMOthers.MBMNumber = MBMCLNOthers.MBMNumber
INNER JOIN cis.Hdb_MBM MBM WITH (NOLOCK) ON MBM.MBMNumber = MBMOthers.MBMNumber
INNER JOIN cis.Ws_Claim Claim WITH (NOLOCK) ON Claim.CASNumber = CISCas.CASNumber
LEFT OUTER JOIN cis.Ws_CLMMedCharges CLMMedCharges ON CLMMedCharges.CASNumber = ciscas.CASNumber
LEFT OUTER JOIN cis.Hdb_MBMCoveredPersons MBMCoveredPersons WITH (NOLOCK) ON MBMCoveredPersons.MBMCNumber = ciscas.MBMnumber AND MBMCoveredPersons.MBMCCoverID = ciscas.CoverID
LEFT OUTER JOIN cis.Hdb_CLN CLN WITH (NOLOCK) ON CISCas.ClinicCode = CLN.CLNCode
LEFT OUTER JOIN dbo.GRPCompany GRPCompany WITH (NOLOCK) ON GRPCompany.GRPCoName = MBMOthers.MBMGroupCompany AND (GRPCompany.hispaycode IS NULL OR GRPCompany.hispaycode = mbm.paycode)
LEFT OUTER JOIN cis.Ws_ClaimInvoice ClaimInvoice WITH (NOLOCK) ON ClaimInvoice.CASNumber = Claim.CASNumber
LEFT OUTER JOIN cis.Ws_ClaimInvoiceActualOthers ClaimInvoiceActualOthers WITH (NOLOCK) ON ClaimInvoiceActualOthers.CASNumber = Claim.CASNumber
LEFT OUTER JOIN cis.Ws_ClaimInvoiceActualNew ClaimInvoiceActualNew WITH (NOLOCK) ON ClaimInvoiceActualNew.CASNumber = Claim.CASNumber
LEFT OUTER JOIN cis.Ws_Bordx Bordx WITH (NOLOCK) ON Bordx.BordxRefNo = Claim.BordxRefNo
LEFT OUTER JOIN dbo.tblRecord tblRecord WITH (NOLOCK) ON tblRecord.CaseNo = Claim.CASNumber
LEFT OUTER JOIN cis.Pay_Payment2CLN Payment2CLN WITH (NOLOCK) ON Payment2CLN.CASNumber = Claim.CASNumber
LEFT JOIN dbo.LG LG WITH (NOLOCK) ON LG.LGNo = CISCAS.CASNumber
INNER JOIN dbo.PLN PLN WITH (NOLOCK) ON PLN.PLNCode = MBMCLNOthers.PLNType
WHERE (MBMCLNOthers.PAYCode <> N''DZ'') AND (MBMCLNOthers.PAYCode <> N''EM'') AND (MBMCLNOthers.PAYCode <> N''ES'') AND (CISCas.CASStatus NOT IN (''D'', ''C''))'
        + @w + N'
OPTION (RECOMPILE);
WITH cte AS (
    SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType1 AS MedInjType, MedInjDays1 AS MedInjDays, MedInjQty1 AS MedInjQty, MedCharges1 AS MedCharges, 1 AS MedNo,
        SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer, AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5,
        CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName, MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType1)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType2, MedInjDays2, MedInjQty2, MedCharges2, 2, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType2)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType3, MedInjDays3, MedInjQty3, MedCharges3, 3, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType3)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType4, MedInjDays4, MedInjQty4, MedCharges4, 4, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType4)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType5, MedInjDays5, MedInjQty5, MedCharges5, 5, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType5)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType6, MedInjDays6, MedInjQty6, MedCharges6, 6, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType6)), '''') <> ''''
    UNION ALL SELECT MBMIcBcPp, MBMName, CASNumber, DateVisit, MedInjType7, MedInjDays7, MedInjQty7, MedCharges7, 7, SPFee, INVAmount, CasRemark, AuditQuery, AuditAnswer,
        AuditRemarks, AuditDiscountReceived, ReasonDesc1, ReasonDesc2, ReasonDesc3, ReasonDesc4, ReasonDesc5, CASDiagnosis5, AuditBy, AuditDate, CLNCode, CLNName,
        MBMGroupCompany, BordxRefNo, Procedurefee FROM #p WHERE COALESCE(LTRIM(RTRIM(MedInjType7)), '''') <> ''''
)
SELECT COUNT(*) OVER (PARTITION BY CASNumber) AS MedCount, ROW_NUMBER() OVER (PARTITION BY CASNumber ORDER BY MedNo) AS RNo, *
FROM cte ORDER BY CASNumber, MedNo;
DROP TABLE #p;';

    EXEC sp_executesql @sql, N'@caseFrom nvarchar(20), @caseTo nvarchar(20), @payTo nvarchar(2), @clinic nvarchar(20), @payor nvarchar(10),
        @groupCompany nvarchar(200), @patientName nvarchar(200), @mbmFrom nvarchar(50), @mbmTo nvarchar(50), @invFrom nvarchar(50), @invTo nvarchar(50),
        @bordxFrom nvarchar(30), @bordxTo nvarchar(30), @caseStatus char(1), @claimability char(1), @planFrom nvarchar(20), @planTo nvarchar(20),
        @bordxDateFrom datetime, @bordxDateTo datetime, @icFrom nvarchar(50), @icTo nvarchar(50), @visitFrom datetime, @visitTo datetime,
        @regFrom datetime, @regTo datetime, @invDateFrom datetime, @invDateTo datetime, @enteredFrom datetime, @enteredTo datetime,
        @amtFrom money, @amtTo money, @monthFrom nvarchar(4), @monthTo nvarchar(4), @yearFrom nvarchar(4), @yearTo nvarchar(4),
        @opFrom nvarchar(30), @opTo nvarchar(30), @subsidiary nvarchar(200), @department nvarchar(200), @omniRefFrom nvarchar(50), @omniRefTo nvarchar(50)',
        @caseFrom, @caseTo, @payTo, @clinic, @payor, @groupCompany, @patientName, @mbmFrom, @mbmTo, @invFrom, @invTo, @bordxFrom, @bordxTo,
        @caseStatus, @claimability, @planFrom, @planTo, @bordxDateFrom, @bordxDateTo, @icFrom, @icTo, @visitFrom, @visitTo, @regFrom, @regTo,
        @invDateFrom, @invDateTo, @enteredFrom, @enteredTo, @amtFrom, @amtTo, @monthFrom, @monthTo, @yearFrom, @yearTo, @opFrom, @opTo,
        @subsidiary, @department, @omniRefFrom, @omniRefTo;
END
GO

-- CboGrpCom_LostFocus / ChkPayor2_Click: subsidiaries and departments of a group company, payors.
CREATE OR ALTER PROCEDURE cis.AuditTrack_Lookups @groupCompany nvarchar(200) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT DISTINCT LTRIM(RTRIM(Subsidiary)) AS Subsidiary FROM cis.Hdb_MBMOthers
    WHERE @groupCompany IS NOT NULL AND MBMGroupCompany = @groupCompany AND COALESCE(LTRIM(RTRIM(Subsidiary)), '') <> '' ORDER BY 1;
    SELECT DISTINCT LTRIM(RTRIM(MBMDepartment)) AS Department FROM cis.Hdb_MBMOthers
    WHERE @groupCompany IS NOT NULL AND MBMGroupCompany = @groupCompany AND COALESCE(LTRIM(RTRIM(MBMDepartment)), '') <> '' ORDER BY 1;
END
GO
