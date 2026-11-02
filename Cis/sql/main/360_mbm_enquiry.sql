-- Membership > MBM Enquiry (FrmMBMCLNEnq): member (HISDB SearchCISMemberProc_test), plan look-ups, top-ups, utilisation, cases.
-- Limits / age band / histories reuse the MBM Adjustment procedures (cis.MbmAdj_*).

-- FindMBM: SearchCISMemberProc_test.
CREATE OR ALTER PROCEDURE cis.MbmEnq_Member @mbmNumber varchar(64) AS
BEGIN
    SET NOCOUNT ON;
    EXEC cis.Hdb_SearchCISMemberProc_test @MBMNumber = @mbmNumber;
END
GO

-- showMBM / cmdShow_Click look-ups of the member's plan: AnnualLimit (first row of plan + insured type), maternity limit, specialist
-- limit status, CISDB PLN flags, the product caption (HISDB MBM + PLN) and the FlexCare check (tblFlexibleBenefitPlanLimit_OP).
CREATE OR ALTER PROCEDURE cis.MbmEnq_Plan @mbmNumber nvarchar(30), @plnCode nvarchar(10), @insCode nvarchar(5), @expDate datetime = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 SuppLimitStatus, PDentalAnnLimit, SDentalAnnLimit, PSpecAnnLimit, SSpecAnnLimit, AnnLimit
    FROM dbo.AnnualLimit WHERE PLNCode = @plnCode AND INSCode = @insCode;
    SELECT TOP 1 a.AnnMaternityLmt FROM dbo.AnnualLimit a INNER JOIN dbo.PLN p ON p.PLNCode = a.PLNCode
    WHERE p.PLNMaternityInd = 'Y' AND p.PLNCode = @plnCode;
    SELECT TOP 1 SuppLimitStatus FROM dbo.SpecialistLimit WHERE PLNCode = @plnCode AND INSCode = @insCode;
    SELECT TOP 1 PLNMaternityInd, PDentalStatus FROM dbo.PLN WHERE PLNCode = @plnCode;
    SELECT TOP 1 p.ProductName, p.PLNDescription FROM cis.Hdb_MBM m INNER JOIN cis.Hdb_PLN p ON p.PLNCode = m.PLNCode WHERE m.MBMNumber = @mbmNumber;
    SELECT COUNT(*) FROM cis.Hdb_MBMCLNOthers o INNER JOIN dbo.tblFlexibleBenefitPlanLimit_OP f ON f.CISPlnCode = o.PLNType
    WHERE o.MBMNumber = @mbmNumber AND @expDate BETWEEN f.FBEffDate AND f.FBExpDate;
END
GO

-- showMBM: top-up totals (MBMCISTopupHistory) of all covers (@covId NULL) or one cover.
CREATE OR ALTER PROCEDURE cis.MbmEnq_Topup @mbmNumber nvarchar(30), @covId nvarchar(2) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 SUM(COALESCE(TopupAmt, 0)) AS TopupAmt, SUM(COALESCE(SP_TopupAmt, 0)) AS SpTopupAmt
    FROM dbo.MBMCISTopupHistory WHERE MBMNumber = @mbmNumber AND (@covId IS NULL OR MBMCovID = @covId) GROUP BY MBMNumber;
END
GO

-- CallLimit (ASO members): Tbl_ASOMemberDetails.
CREATE OR ALTER PROCEDURE cis.MbmEnq_Aso @mbmNumber nvarchar(30), @covId nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    SELECT MBMASOLimit, MBMIPLimit, MBMOPLimit, MBMIPUtilized, MBMOPUtilized, IsIPOPCombined
    FROM cis.Hdb_ASOMemberDetails WHERE MBMNumber = @mbmNumber AND MBMCoverID = @covId;
END
GO

-- Medix.bas GetSpecLimit: PLN.PSpecLimit of plan + payor (last row), SpecialistLimit of insured type + plan, health type N.
CREATE OR ALTER PROCEDURE cis.MbmEnq_SpecLimit @plnCode nvarchar(10), @payCode nvarchar(10), @insCode nvarchar(5) AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(PSpecLimit) AS SpecLmtInd FROM dbo.PLN WHERE PLNCode = @plnCode AND PAYCode = @payCode;
    SELECT TOP 1 PrinLimit, SuppLimit, RTRIM(SuppLimitStatus) AS SuppLimitStatus FROM dbo.SpecialistLimit
    WHERE INSCode = @insCode AND PLNCode = @plnCode AND HLTCode = 'N';
END
GO

-- CIS_FFS2018_test with its dynamic filter replaced by parameters: the member's claims, all covers (@coverMode 0), cover 00 (1),
-- covers other than 00 (2) or one cover (3, @covId).
CREATE OR ALTER PROCEDURE cis.MbmEnq_Utilisation @mbmNumber nvarchar(30), @coverMode int = 0, @covId nvarchar(2) = NULL AS
BEGIN
    SET NOCOUNT ON;
    SELECT CASE
             WHEN '305' IN (c.CASTreatReason1, c.CASTreatReason2, c.CASTreatReason3, c.CASTreatReason4, c.CASTreatReason5) THEN 'OPTICAL'
             WHEN '303' IN (c.CASTreatReason1, c.CASTreatReason2, c.CASTreatReason3, c.CASTreatReason4, c.CASTreatReason5) THEN 'DENTAL'
             WHEN '302' IN (c.CASTreatReason1, c.CASTreatReason2, c.CASTreatReason3, c.CASTreatReason4, c.CASTreatReason5) THEN 'MATERNITY'
             WHEN '465' IN (c.CASTreatReason1, c.CASTreatReason2, c.CASTreatReason3, c.CASTreatReason4, c.CASTreatReason5) THEN 'DENTAL EXTRACTION (EXCLUDE IN DENTAL LIMIT)'
             WHEN '467' IN (c.CASTreatReason1, c.CASTreatReason2, c.CASTreatReason3, c.CASTreatReason4, c.CASTreatReason5) THEN 'DENTALEXTRACTION'
             WHEN cln.CLNType = 'G' THEN 'GP'
             WHEN cln.CLNType = 'S' THEN 'SP'
           END AS TreatmentType,
           cl.CLMMCOProcFee AS FFS, cln.CLNType, ia.CLMMaternity, c.MBMnumber, c.CoverID, c.CASNumber,
           CASE WHEN p.CoPayInd = 'Y' AND COALESCE(cl.CLMCoPayColl, 0) <> 0 THEN cl.CLMTotalActual - cl.CLMCoPayColl ELSE cl.CLMTotalActual END AS CLMBordxAmt,
           p.PLNGPSPShared,
           CASE WHEN tr1.ReasonCategory IN ('dental', 'optical') AND COALESCE(p.PDentalStatus, 'N') <> 'N' THEN 0
                ELSE COALESCE(ia.CLMOthersNoDed, 0.00) END AS CLMOthersNoDed,
           ia.CLMTyphiadJabNoDed,
           CASE WHEN 'dental' IN (tr1.ReasonCategory, tr2.ReasonCategory, tr3.ReasonCategory, tr4.ReasonCategory, tr5.ReasonCategory)
                     AND ia.CLMDental = 0 AND ia.CLMOthersNoDed = 0 THEN cl.CLMBordxAmt ELSE ia.CLMDental END AS CLMDental,
           CASE WHEN c.CASTreatReason1 IN ('303', '305', '467') THEN cl.CLMBordxAmt ELSE 0 END AS Utilization,
           CASE WHEN c.CASTreatReason1 NOT IN ('303', '305', '467') AND COALESCE(ia.CLMOthersNoDed, 0.00) = 0 THEN COALESCE(cl.CLMBTBGAmt, 0) ELSE 0 END AS CLMBTBGAmt,
           c.OPBatchNo, p.IsGPDentalShared, CLMGPFeeBasicMed, p.PLNCode, SDentalAnnLimit, SDentalStatus, p.PSpecLimit, p.PSpecStatus, p.SSpecLimit,
           p.SSpecStatus, PSpecAnnLimit, SSpecAnnLimit
    FROM dbo.CISCas c WITH (NOLOCK)
    INNER JOIN cis.Ws_Claim cl WITH (NOLOCK) ON c.CASNumber = cl.CASNumber
    INNER JOIN cis.Hdb_MBMCLNOthers o ON o.MBMNumber = c.MBMnumber
    INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    LEFT JOIN dbo.AnnualLimit al ON al.PLNCode = p.PLNCode AND al.INSCode IN ('H')
    INNER JOIN cis.Ws_ClaimInvoiceActualNew ia WITH (NOLOCK) ON cl.CASNumber = ia.CASNumber
    LEFT OUTER JOIN cis.Hdb_CLN cln WITH (NOLOCK) ON c.ClinicCode = cln.CLNCode
    LEFT JOIN dbo.TreatmentReason tr1 ON tr1.ReasonCode = c.CASTreatReason1
    LEFT JOIN dbo.TreatmentReason tr2 ON tr2.ReasonCode = c.CASTreatReason2
    LEFT JOIN dbo.TreatmentReason tr3 ON tr3.ReasonCode = c.CASTreatReason3
    LEFT JOIN dbo.TreatmentReason tr4 ON tr4.ReasonCode = c.CASTreatReason4
    LEFT JOIN dbo.TreatmentReason tr5 ON tr5.ReasonCode = c.CASTreatReason5
    WHERE c.CASStatus <> 'D' AND c.MBMnumber = @mbmNumber
      AND (@coverMode = 0 OR (@coverMode = 1 AND c.CoverID = '00') OR (@coverMode = 2 AND c.CoverID <> '00') OR (@coverMode = 3 AND c.CoverID = @covId));
END
GO

-- tmppridentallimit / Dentallimit_OneLimit: the existing CISDB dental procedures (parameterised).
CREATE OR ALTER PROCEDURE cis.MbmEnq_Dental @mbmNumber nvarchar(30), @covId nvarchar(4) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.tmpDlimit_local_test1 @MBMNumber = @mbmNumber, @MBMCCoverID = @covId;
END
GO

CREATE OR ALTER PROCEDURE cis.MbmEnq_DentalOneLimit @mbmNumber nvarchar(30) AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.tmpDentallimit_OneLimit @MBMNumber = @mbmNumber;
END
GO

-- ListCases / ListSpecialistCases: VIEW_CISMBMInquiry_GPCASE / VIEW_CISMBMInquiry_Spe_CASE.
CREATE OR ALTER PROCEDURE cis.MbmEnq_Cases @mbmNumber nvarchar(30), @specialist bit AS
BEGIN
    SET NOCOUNT ON;
    IF @specialist = 0
        SELECT CASNumber, MBMCName, ReasonDesc, CoverID, CLMMCOProcFee, DateVisit, MBMCRELCode, CLMTotalActual, CLMBordxAmt, CASStatus, MBMName, MBMNumber, Claimability
        FROM dbo.VIEW_CISMBMInquiry_GPCASE WHERE MBMNumber = @mbmNumber;
    ELSE
        SELECT CASNumber, MBMCName, ReasonDesc, CoverID, CLMMCOProcFee, DateVisit, MBMCRELCode, CLMTotalActual, CLMBordxAmt, CASStatus, MBMName, MBMNumber, Claimability
        FROM dbo.VIEW_CISMBMInquiry_Spe_CASE WHERE MBMNumber = @mbmNumber;
END
GO

-- showMBMCov / lstCovAdd_Click: the cover plan's AnnualLimit and SpecialistLimit statuses on the policy date, in the desktop's read order
-- (latest first; the desktop keeps the last row it reads).
CREATE OR ALTER PROCEDURE cis.MbmEnq_CoverStatus @plnCode nvarchar(10), @insCode nvarchar(5), @effDate datetime AS
BEGIN
    SET NOCOUNT ON;
    SELECT RTRIM(SuppLimitStatus) AS SuppLimitStatus FROM dbo.AnnualLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND AnnEffDate <= @effDate ORDER BY AnnEffDate DESC;
    SELECT RTRIM(SuppLimitStatus) AS SuppLimitStatus FROM dbo.SpecialistLimit WHERE INSCode = @insCode AND PLNCode = @plnCode AND EffDate <= @effDate ORDER BY EffDate DESC;
END
GO

-- showMBMCov, limit status F: PLN + AnnualLimit of the plan (first row).
CREATE OR ALTER PROCEDURE cis.MbmEnq_FLimit @plnCode nvarchar(10) AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 RTRIM(p.ANNLmtIndSupp) AS AnnLmtIndSupp, a.SuppLimit, a.ChildDepLimit, RTRIM(p.SSpecLimit) AS SSpecLimit, RTRIM(p.VISLmtInd) AS VisLmtInd
    FROM dbo.PLN p INNER JOIN dbo.AnnualLimit a ON a.PLNCode = p.PLNCode WHERE p.PLNCode = @plnCode;
END
GO

-- lstCovAdd_Click: FFS of CIS_FFS2013 (which also needs an HISDB MBMTwo row) for covers other than 00 (@coverMode 2) or one cover (3).
CREATE OR ALTER PROCEDURE cis.MbmEnq_CoverFfs @mbmNumber nvarchar(30), @coverMode int, @covId nvarchar(2) AS
BEGIN
    SET NOCOUNT ON;
    SELECT cl.CLMMCOProcFee AS FFS
    FROM dbo.CISCas c WITH (NOLOCK)
    INNER JOIN cis.Ws_Claim cl WITH (NOLOCK) ON c.CASNumber = cl.CASNumber
    INNER JOIN cis.Hdb_MBMCLNOthers o ON o.MBMNumber = c.MBMnumber
    INNER JOIN cis.Hdb_MBMTwo t ON t.MBMNumber = c.MBMnumber
    INNER JOIN dbo.PLN p ON p.PLNCode = o.PLNType
    INNER JOIN cis.Ws_ClaimInvoiceActualNew ia WITH (NOLOCK) ON cl.CASNumber = ia.CASNumber
    WHERE c.CASStatus <> 'D' AND c.MBMnumber = @mbmNumber AND ((@coverMode = 2 AND c.CoverID <> '00') OR (@coverMode = 3 AND c.CoverID = @covId));
END
GO
