-- Claim / Bordx > Claim Tracking (Simple) (frmClaimRptSimple and its Jupiter copy): the 107 cases of CISWorksheet CISClaimTracking_Simple (generated from the procedure by
-- gen-trk.js). Each case's query text is kept as the desktop's; the pasted criteria are replaced by a WHERE fragment built only from the filters that are set,
-- with every value passed as a parameter.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.ClmTrkSimple_Run @case int,
    @hlt nvarchar(1) = NULL,
    @hltBlank bit = 0,
    @ins nvarchar(1) = NULL,
    @insBlank bit = 0,
    @memType nvarchar(10) = NULL,
    @race nvarchar(1) = NULL,
    @raceBlank bit = 0,
    @sex nvarchar(1) = NULL,
    @sexBlank bit = 0,
    @relation nvarchar(1) = NULL,
    @relationBlank bit = 0,
    @dataStatus nvarchar(1) = NULL,
    @dataStatusBlank bit = 0,
    @claimability nvarchar(1) = NULL,
    @claimabilityBlank bit = 0,
    @polStatus nvarchar(1) = NULL,
    @polStatusBlank bit = 0,
    @expired nvarchar(1) = NULL,
    @expiredBlank bit = 0,
    @premMode nvarchar(1) = NULL,
    @premModeBlank bit = 0,
    @takeover nvarchar(1) = NULL,
    @takeoverBlank bit = 0,
    @caseStatus nvarchar(10) = NULL,
    @caseStatusBlank bit = 0,
    @state nvarchar(100) = NULL,
    @stateBlank bit = 0,
    @departmentBlank bit = 0,
    @renewal nvarchar(10) = NULL,
    @renewalBlank bit = 0,
    @agent nvarchar(200) = NULL,
    @agentBlank bit = 0,
    @policy nvarchar(200) = NULL,
    @policyBlank bit = 0,
    @patient nvarchar(200) = NULL,
    @agebandFrom nvarchar(20) = NULL,
    @agebandTo nvarchar(20) = NULL,
    @polYearFrom nvarchar(20) = NULL,
    @polYearTo nvarchar(20) = NULL,
    @polYearBlank bit = 0,
    @mcDaysFrom nvarchar(20) = NULL,
    @mcDaysTo nvarchar(20) = NULL,
    @mcDaysBlank bit = 0,
    @payor nvarchar(10) = NULL,
    @payorBlank bit = 0,
    @company nvarchar(200) = NULL,
    @companyBlank bit = 0,
    @clinic nvarchar(20) = NULL,
    @clinicBlank bit = 0,
    @clinicGroup nvarchar(200) = NULL,
    @branch nvarchar(20) = NULL,
    @dovFrom datetime = NULL,
    @dovTo datetime = NULL,
    @dovBlank bit = 0,
    @dorFrom datetime = NULL,
    @dorTo datetime = NULL,
    @dorBlank bit = 0,
    @dowFrom datetime = NULL,
    @dowTo datetime = NULL,
    @dowBlank bit = 0,
    @invoiceDateFrom datetime = NULL,
    @invoiceDateTo datetime = NULL,
    @invoiceDateBlank bit = 0,
    @treatmentFrom nvarchar(20) = NULL,
    @treatmentTo nvarchar(20) = NULL,
    @treatmentBlank bit = 0,
    @planFrom nvarchar(20) = NULL,
    @planTo nvarchar(20) = NULL,
    @planBlank bit = 0,
    @effFrom datetime = NULL,
    @effTo datetime = NULL,
    @effBlank bit = 0,
    @expFrom datetime = NULL,
    @expTo datetime = NULL,
    @expBlank bit = 0,
    @joinedFrom datetime = NULL,
    @joinedTo datetime = NULL,
    @joinedBlank bit = 0,
    @dobFrom datetime = NULL,
    @dobTo datetime = NULL,
    @dobBlank bit = 0,
    @mbmNoFrom nvarchar(30) = NULL,
    @mbmNoTo nvarchar(30) = NULL,
    @caseNoFrom nvarchar(20) = NULL,
    @caseNoTo nvarchar(20) = NULL,
    @bordxType nvarchar(10) = NULL,
    @bordxFrom nvarchar(20) = NULL,
    @bordxTo nvarchar(20) = NULL,
    @bordxBlank bit = 0,
    @bordxDateFrom datetime = NULL,
    @bordxDateTo datetime = NULL,
    @bordxDateBlank bit = 0,
    @icFrom nvarchar(30) = NULL,
    @icTo nvarchar(30) = NULL,
    @icBlank bit = 0,
    @mbmIcFrom nvarchar(30) = NULL,
    @mbmIcTo nvarchar(30) = NULL,
    @mbmIcBlank bit = 0,
    @actualFrom money = NULL,
    @actualTo money = NULL,
    @approvedFrom money = NULL,
    @approvedTo money = NULL,
    @osUtilizedFrom money = NULL,
    @osUtilizedTo money = NULL,
    @panel nvarchar(10) = NULL,
    @nonPanel bit = 0,
    @pending nvarchar(10) = NULL,
    @year nvarchar(4) = NULL,
    @invoiceFrom nvarchar(50) = NULL,
    @invoiceTo nvarchar(50) = NULL,
    @invoiceBlank bit = 0,
    @submissionFrom nvarchar(50) = NULL,
    @submissionTo nvarchar(50) = NULL,
    @submissionBlank bit = 0,
    @opBatchFrom nvarchar(30) = NULL,
    @opBatchTo nvarchar(30) = NULL,
    @clinicCodeFrom nvarchar(20) = NULL,
    @clinicCodeTo nvarchar(20) = NULL,
    @clinicType nvarchar(10) = NULL,
    @payTo nvarchar(1) = NULL,
    @vsAmt bit = 0,
    @deptSelection nvarchar(200) = NULL,
    @ieInvoice bit = 0
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @b nvarchar(max), @a nvarchar(max) = N'';
    IF @case = 1 SELECT @b = N'SELECT MBMDepartment, CASNumber, BordxRefNo, BordxDate, MBMNumber, CoverID,
            						PatientName, ClinicCode, CLNName, INVPayTo, CASTreatReason1, CLNType,
            						CASDiagnosis5, ReasonDesc, CLMTotalActual, CLMNonGpFee, DateVisit, MCDays
            						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'';
    IF @case = 2 SELECT @b = N'SELECT ClinicCode, CLNName, CLNType, COUNT(CASNumber) as NoofVisit,
            						SUM(MCStatus) as MCTimes, SUM(MCDays) as MCDays, SUM(CLMTotalActual) as BillAmt
            						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY ClinicCode, CLNName, CLNType';
    IF @case = 3 SELECT @b = N'SELECT MBMDepartment, BordxRefNo, BordxDate, 
            						SUM(CLMTotalActual) as CLMTotalActual, SUM(CLMNonGPFee) as CLMNonGPFee
            						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMDepartment, BordxRefNo, BordxDate';
    IF @case = 4 SELECT @b = N'SELECT MBMDepartment, MBMNumber, CoverID, PatientName, Count(CASNumber) AS TotalVisit,
          						SUM(MCStatus) as MCTimes, SUM(MCDays) AS MCDays,  
          						SUM(CLMTotalActual) as CLMTotalActual, SUM(CLMNonGPFee) as CLMNonGPFee 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMDepartment, MBMNumber, CoverID, PatientName';
    IF @case = 5 SELECT @b = N'SELECT MBMDepartment, MBMNumber, MBMName, CoverID, AnnLimit, AnnLmtIND, SUM(CLMTotalActual) AS ApprovedAmt,
          						Count(CASNumber) AS TotalClaim, SUM(MCStatus) AS MCTimes, 
          						SUM(MCDays) AS MCDays, MBMCLNBalLmt,Sum(TopupAmt) AS TopupAmt From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMDepartment, MBMNumber, MBMName, CoverID, AnnLimit, AnnLmtIND, MBMCLNBalLmt
							ORDER BY MBMNumber';
    IF @case = 6 SELECT @b = N'SELECT CASTreatReason1, Sum(CLMTotalActual) As CLMTotalActual,
          						SUM(CLMNonGpFee) AS CLMNonGpFee From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY CASTreatReason1';
    IF @case = 7 SELECT @b = N'SELECT Distinct BordxRefNo, BordxDate, BordxNoClaims, SUM(CLMTotalActual) 
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, 
          						BordxClaimAmt From VIEW_CLMReport_Simple Where BordxRefNo IS NOT NULL', @a = N'GROUP BY BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt';
    IF @case = 8 SELECT @b = N'SELECT Distinct MBMGroupCompany, BordxRefNo, BordxDate, BordxNoClaims, SUM(CLMTotalActual) 
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						BordxClaimAmt From VIEW_CLMReport_Simple Where BordxRefNo IS NOT NULL', @a = N'GROUP BY MBMGroupCompany, BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt';
    IF @case = 9 SELECT @b = N'SELECT Distinct MBMDepartment, MBMGroupCompany, BordxRefNo, BordxDate, BordxNoClaims, SUM(CLMTotalActual)
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						BordxClaimAmt From VIEW_CLMReport_Simple Where BordxRefNo IS NOT NULL', @a = N'GROUP BY MBMDepartment, MBMGroupCompany, BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt 
							ORDER BY BordxRefNo';
    IF @case = 11 SELECT @b = N'SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved, Sum(CLMNonGPFee) As CLMNonGPFee
                           				From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY CASTreatReason1';
    IF @case = 12 SELECT @b = N'SELECT PLNCode, SUM(CLMTotalApproved) AS TotalPanelAmt
                  					From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PLNCode';
    IF @case = 13 SELECT @b = N'SELECT Sex, Ageband, COUNT(CASNumber) AS TotalCases,
          						SUM(CLMTotalActual) AS TotalActualAmt,
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY Sex, Ageband';
    IF @case = 14 SELECT @b = N'SELECT Ageband, COUNT(CASNumber) AS TotalCases, 
          						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY Ageband';
    IF @case = 15 SELECT @b = N'SELECT PAYCode, INSCode, Ageband, COUNT(CASNumber) AS TotalCases,
          						SUM(CLMTotalActual) AS TotalActualAmt,
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, INSCode, Ageband';
    IF @case = 16 SELECT @b = N'SELECT PAYCode, MBMAgentCode, CASDoctor, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode, CASDoctor';
    IF @case = 17 SELECT @b = N'SELECT PAYCode, MBMAgentCode, CASNumber, MBMNumber,
          						CoverID, PatientName, CLMTotalActual,
          						DateVisit, CLMTotalApproved, CLMRecoveryPaidbyM,
          						CASTreatReason1, InsCode, PLNCode, CASStatus, MCDays 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'';
    IF @case = 18 SELECT @b = N'SELECT PAYCode, MBMAgentCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) 
              					AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
              					From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode';
    IF @case = 19 SELECT @b = N'SELECT Top 50 PAYCode, MBMAgentCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
              					AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
              					From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode 
							ORDER BY TotalCases Desc';
    IF @case = 20 SELECT @b = N'SELECT INSCode, PLNCode, Ageband, COUNT(CASNumber) AS TotalCases,
          						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY INSCode, PLNCode, Ageband';
    IF @case = 21 SELECT @b = N'SELECT RACCode, Ageband, COUNT(CASNumber) AS TotalCases,
          						SUM(CLMTotalActual) AS TotalActualAmt,
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY RACCode, Ageband';
    IF @case = 22 SELECT @b = N'SELECT ClinicGRPName, COUNT(CASNumber) AS TotalCases, 
          						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY ClinicGRPName ORDER BY ClinicGRPName';
    IF @case = 23 SELECT @b = N'SELECT COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved,
          						Sum(CLMTotalActual) As TotalActual, ClinicCode, CLNName
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY CASTreatReason1, ClinicCode, CLNName';
    IF @case = 24 SELECT @b = N'SELECT COUNT(CASNumber) AS Total, CLNType, Sum(CLMTotalApproved) As TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY CLNType';
    IF @case = 25 SELECT @b = N'SELECT PAYCode, ClinicCode, CLNName, CASNumber, MBMNumber, CoverID,
          						PatientName, CLMTotalActual, CLMTotalApproved,
          						CLMRecoveryPaidbyM, DateVisit, 
          						CASStatus, Claimability, CASTreatReason1, MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'';
    IF @case = 26 SELECT @b = N'SELECT PAYCode, ClinicCode, CLNName, COUNT(CLNName) AS Total, 
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, ClinicCode, CLNName';
    IF @case = 27 SELECT @b = N'SELECT ClinicCode, CLNName, CLNType, CLNLocation, SUM(MCDays) AS MCDays, COUNT(CLNName) AS Total,
          						SUM(MCStatus) AS MCTimes, SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY ClinicCode, CLNName, CLNType, CLNLocation';
    IF @case = 28 SELECT @b = N'SELECT MBMGroupCompany, ClinicCode, CLNName, COUNT(CLNName) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMGroupCompany, ClinicCode, CLNName';
    IF @case = 29 SELECT @b = N'SELECT TOP 20 PAYCode, ClinicCode, CLNName, COUNT(CLNName) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where  0 = 0', @a = N'GROUP BY PAYCode, ClinicCode, CLNName ORDER BY Total DESC';
    IF @case = 30 SELECT @b = N'SELECT TOP 20 PAYCode, ClinicCode, CLNName, COUNT(CLNName) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where  0 = 0', @a = N'GROUP BY PAYCode, ClinicCode, CLNName ORDER BY TotalActual DESC';
    IF @case = 31 SELECT @b = N'SELECT CLNType, Count(CASNumber) AS TotalClaimP, SUM(CLMTotalApproved) AS TotalPanelAmt
                           				From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY CLNType';
    IF @case = 32 SELECT @b = N'SELECT CLNType, Count(CASNumber) AS TotalClaimNP, SUM(CLMTotalApproved) AS TotalNonPanelAmt
                          					From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY CLNType';
    IF @case = 33 SELECT @b = N'SELECT PLNCode, SUM(CLMTotalApproved) AS TotalNonPanelAmt
                  					From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PLNCode';
    IF @case = 34 SELECT @b = N'SELECT CASNumber, MBMNumber, MBMPolicyNo, 
          						PatientName, DateVisit, ClinicCode, CLNName, CASTreatReason1, MCDays 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'';
    IF @case = 35 SELECT @b = N'SELECT Distinct PAYCode From VIEW_CLMReport_Simple Where 0 = 0', @a = N'ORDER BY PAYCode';
    IF @case = 36 SELECT @b = N'SELECT Distinct PAYCode From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Order by PAYCode';
    IF @case = 37 SELECT @b = N'SELECT MBMDepartment, CLNName, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment, CLNName Order By CLNName, MBMDepartment';
    IF @case = 38 SELECT @b = N'SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, 
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment Order By MBMDepartment';
    IF @case = 39 SELECT @b = N'SELECT MBMDepartment, MBMNumber, MBMName, MBMCName, CoverID, 
       						MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMDepartment, MBMNumber, MBMName, MBMCName, CoverID, 
        						MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit ORDER BY MBMNumber, MBMDepartment';
    IF @case = 40 SELECT @b = N'SELECT MBMDepartment, ReasonDesc, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,

      						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
         						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment, ReasonDesc ORDER BY ReasonDesc, MBMDepartment';
    IF @case = 41 SELECT @b = N'SELECT MBMDepartment, Count(CASNumber) AS NoofMC, MCDays,
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group by MBMDepartment, MCDays';
    IF @case = 42 SELECT @b = N'SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, 
        						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment Order By MCDays Desc';
    IF @case = 43 SELECT @b = N'SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
        						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment Order By MCTimes Desc';
    IF @case = 44 SELECT @b = N'SELECT MBMDepartment, MBMNumber, CoverID, PatientName, AnnLimit, AnnLmtIND, Count(CASNumber) AS TotalClaim,
          						SUM(MCStatus) as MCTimes, SUM(MCDays) AS MCDays,  
          						SUM(CLMTotalActual) as ApprovedAmt, MBMCLNBalLmt ,Sum(TopupAmt) AS TopupAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMDepartment, MBMNumber, CoverID, PatientName, AnnLimit, AnnLmtIND,  MBMCLNBalLmt
							ORDER BY MBMNumber';
    IF @case = 45 SELECT @b = N'SELECT MBMDepartment, COUNT(CASNumber) as TotalCases, SUM(CLMTotalActual) AS totalActual,
      						SUM(CLMTotalApproved) As TotalApproved From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment';
    IF @case = 46 SELECT @b = N'SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
        						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
      						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment Order By Visit Desc';
    IF @case = 47 SELECT @b = N'SELECT PAYCode, MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, COUNT(CASNumber) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, 
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate) ';
    IF @case = 48 SELECT @b = N'SELECT MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, COUNT(CASNumber) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)';
    IF @case = 49 SELECT @b = N'SELECT PAYCode, MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, COUNT(CASNumber) AS Total, 
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)';
    IF @case = 50 SELECT @b = N'SELECT MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, COUNT(CASNumber) AS Total, 
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, 
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)';
    IF @case = 51 SELECT @b = N'SELECT PAYCode, MONTH(DateVisit) AS MONTH, YEAR(DateVisit) AS YEAR, COUNT(CASNumber) AS Total, 
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, MONTH(DateVisit), YEAR(DateVisit)';
    IF @case = 52 SELECT @b = N'SELECT MONTH(DateVisit) AS MONTH, YEAR(DateVisit) AS YEAR, COUNT(CASNumber) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MONTH(DateVisit), YEAR(DateVisit)';
    IF @case = 53 SELECT @b = N'SELECT PAYCode, MONTH(RegisterDate) AS MONTH, YEAR(RegisterDate) AS YEAR, COUNT(CASNumber) AS Total,
         						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, MONTH(RegisterDate), YEAR(RegisterDate)';
    IF @case = 54 SELECT @b = N'SELECT MONTH(RegisterDate) AS MONTH, YEAR(RegisterDate) AS YEAR, COUNT(CASNumber) AS Total, 
        						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, 
       						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MONTH(RegisterDate), YEAR(RegisterDate)';
    IF @case = 55 SELECT @b = N'SELECT MBMDepartment, MBMGroupCompany, CLNName, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment, CLNName, MBMGroupCompany 
							ORDER BY CLNName, MBMGroupCompany';
    IF @case = 56 SELECT @b = N'SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, 
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment, MBMGroupCompany 
							ORDER BY MBMDepartment, MBMGroupCompany';
    IF @case = 57 SELECT @b = N'SELECT PAYCode, MBMGroupCompany, CASNumber, MBMNumber, CoverID,
          						PatientName, CLMTotalActual, CLMTotalApproved, 
          						CLMRecoveryPaidbyM, DateVisit,
          						CASStatus, Claimability, CASTreatReason1, MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'';
    IF @case = 58 SELECT @b = N'SELECT MBMDepartment, MBMGroupCompany, MBMNumber, MBMName, MBMCName, CoverID,
        						MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMDepartment, MBMGroupCompany, MBMNumber, MBMName, MBMCName, CoverID,
        							MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit ORDER BY MBMNumber, MBMGroupCompany';
    IF @case = 59 SELECT @b = N'SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment, MBMGroupCompany Order By MBMGroupCompany, MBMDepartment';
    IF @case = 60 SELECT @b = N'SELECT MBMDepartment, MBMGroupCompany, ReasonDesc, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
       						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment, MBMGroupCompany, ReasonDesc ORDER BY ReasonDesc, MBMGroupCompany';
    IF @case = 61 SELECT @b = N'SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment, MBMGroupCompany Order By MCDays Desc';
    IF @case = 62 SELECT @b = N'SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By MBMDepartment, MBMGroupCompany Order By MCTimes Desc';
    IF @case = 63 SELECT @b = N'SELECT PAYCode, MBMGroupCompany, COUNT(MBMGroupCompany) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, MBMGroupCompany';
    IF @case = 64 SELECT @b = N'SELECT MBMGroupCompany, COUNT(MBMGroupCompany) AS Total,
         						 SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMGroupCompany';
    IF @case = 65 SELECT @b = N'SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
        						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
         						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMDepartment, MBMGroupCompany ORDER BY Visit Desc';
    IF @case = 66 SELECT @b = N'SELECT ICNo, Count(CASNumber) AS TotalClaim, PatientName, 
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt, SUM(CONVERT(Money,CLMPayableByMedix2M)) AS OSAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY ICNo, PatientName';
    IF @case = 67 SELECT @b = N'SELECT PAYCode, INSCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
         						SUM(clmtotalapproved) As TotalApproved  
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By PAYCode, INSCode';
    IF @case = 68 SELECT @b = N'SELECT PAYCode, PLNCode, INSCode, COUNT(CASNumber) AS TotalCases,
         						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, PLNCode, INSCode';
    IF @case = 69 SELECT @b = N'SELECT PAYCode, INSCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By PAYCode, INSCode, Sex';
    IF @case = 70 SELECT @b = N'SELECT PAYCode, INSCode, RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
        						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By PAYCode, INSCode, RACCode';
    IF @case = 71 SELECT @b = N'SELECT Count(CASNumber) AS NoofMC, MCDays,
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group by MCDays';
    IF @case = 72 SELECT @b = N'SELECT ClinicCode, CLNName, Count(CASNumber) AS NoofMC, MCDays,
         						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims
          						From VIEW_CLMReport_Simple Where 0 = 0 ', @a = N'Group by ClinicCode, CLNName, MCDays';
    IF @case = 73 SELECT @b = N'SELECT MBMGroupCompany, Count(CASNumber) AS NoofMC, MCDays, 
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group by MBMGroupCompany, MCDays';
    IF @case = 74 SELECT @b = N'SELECT CASTreatReason1, Count(CASNumber) AS NoofMC, MCDays,
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group by CASTreatReason1, MCDays';
    IF @case = 75 SELECT @b = N'SELECT MBMNumber, COUNT(CASNumber) AS TotalClaim, 
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt, 
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMNumber';
    IF @case = 76 SELECT @b = N'SELECT MBMGroupCompany, MBMNumber, Count(CASNumber) AS TotalClaim,
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMGroupCompany, MBMNumber';
    IF @case = 77 SELECT @b = N'SELECT PAYCode, GRPCoCode, MBMNumber, PatientName, PLNCode, Count(CASNumber) AS TotalClaim,
    						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt,
       						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, GRPCoCode, MBMNumber, PatientName, PLNCode';
    IF @case = 78 SELECT @b = N'SELECT MBMDepartment, MBMNumber, PatientName, CoverID, SUM(CLMTotalActual) AS ApprovedAmt, 
       						Count(CASNumber) AS TotalClaim, SUM(MCStatus) AS MCTimes, 
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0 ', @a = N'GROUP BY MBMDepartment, MBMNumber, PatientName, CoverID';
    IF @case = 79 SELECT @b = N'SELECT MBMNumber, MBMName, COUNT(CASNumber) AS TotalClaim,
      						SUM(CONVERT(Numeric, MCStatus)) AS MCStatus, SUM(MCDays) AS MCDays, AnnLimit, 
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMNumber, MBMName, AnnLimit';
    IF @case = 80 SELECT @b = N'SELECT PAYCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
          						SUM(clmtotalapproved) As TotalApproved From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By PAYCode';
    IF @case = 81 SELECT @b = N'SELECT PAYCode, PLNCode, Ageband, COUNT(CASNumber) AS TotalCases, 



          						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, PLNCode, Ageband';
    IF @case = 82 SELECT @b = N'SELECT PLNCode, SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PLNCode';
    IF @case = 83 SELECT @b = N'SELECT PAYCode, PLNCode, INSCode, Ageband, COUNT(CASNumber) AS TotalCases, 
          						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, PLNCode, INSCode, Ageband';
    IF @case = 84 SELECT @b = N'SELECT PAYCode, PLNCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
       						SUM(clmtotalapproved) As TotalApproved
        						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By PAYCode, PLNCode';
    IF @case = 85 SELECT @b = N'SELECT PAYCode, PLNCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
          						SUM(clmtotalapproved) As TotalApproved, 
          						SUM(MBMCLNBasicPrem) AS BasicPremium, 
          						SUM(MBMCLNPremium) AS Premium From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By PAYCode, PLNCode';
    IF @case = 86 SELECT @b = N'SELECT Top 10 PAYCode, PLNCode, CASTreatReason1, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, 
          						SUM(clmtotalapproved) As TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By PAYCode, PLNCode, CASTreatReason1 ORDER BY TotalCases Desc';
    IF @case = 87 SELECT @b = N'SELECT Top 10 PAYCode, PLNCode, ClinicCode, CLNName, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
      						SUM(clmtotalapproved) As TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By PAYCode, PLNCode, ClinicCode, CLNName ORDER BY TotalCases Desc';
    IF @case = 88 SELECT @b = N'SELECT PAYCode, MBMPolicyNo, MBMGroupCompany,
          						COUNT(CASNumber) AS TotalCases, 
          						SUM(CLMTotalActual) AS TotalActualAmt,
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, MBMPolicyNo, MBMGroupCompany';
    IF @case = 89 SELECT @b = N'SELECT RACCode, Ageband, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
         						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By RACCode, Ageband';
    IF @case = 90 SELECT @b = N'SELECT RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
         						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved 
      						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By RACCode';
    IF @case = 91 SELECT @b = N'SELECT INSCode, PLNCode, RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
      						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By INSCode, PLNCode, RACCode';
    IF @case = 92 SELECT @b = N'SELECT RACCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By RACCode, Sex';
    IF @case = 93 SELECT @b = N'SELECT PAYCode, CASNumber, MBMNumber, MBMPolicyNo, 
          						CoverID, MBMName, MBMDOB, DOB, PatientName, MBMSex, AgeCode, Sex, Ageband, 
         						INSCode, PLNCode, ClinicCode, CLNName, CASTreatReason1, 
          						CLMTotalActual, CLMTotalApproved, MCDays, ReasonDesc,
          						CLMPayablebyMedix2M, BordxRefNo, BordxDate, ClaimOpenDate, 
          						Claimability, CASStatus, MBMDataStatus, MBMStatus, DateVisit,
          						MBMPayorEffDate, MBMPayorExpDate, MBMCLNPremium, 
          						RegisterDate, MBMIcBcPp, ICNo, CASDiagnosis5, MBMDepartment, INVPayTo, 
          						CLMNonGpFee, CLNType From VIEW_CLMReport_Simple Where 0 = 0', @a = N'';
    IF @case = 94 SELECT @b = N'SELECT PAYCode, CASNumber, MBMNumber, MBMPolicyNo,
         						CoverID, MBMName, MBMDOB, DOB, PatientName, MBMSex, AgeCode, Sex, Ageband, 
          						INSCode, PLNCode, ClinicCode, CLNName, CASTreatReason1, ReasonDesc,
          						CLMTotalActual, CLMTotalApproved, Remarks, MCDays,
          						CLMPayablebyMedix2M, BordxRefNo, BordxDate, ClaimOpenDate,  
          						Claimability, CASStatus, MBMDataStatus, MBMStatus, DateVisit, 
        						MBMPayorEffDate, MBMPayorExpDate, MBMCLNPremium,
          						RegisterDate, MBMIcBcPp, ICNo, CASDiagnosis5 From VIEW_CLMReport_Simple Where 0 = 0', @a = N'';
    IF @case = 95 SELECT @b = N'SELECT Count(CASNumber) AS NoofVisits, MBMNumber, 
          						PatientName, PLNCode, SUM(CLMTotalApproved) AS AmtClaims,
          						ICNo From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group by MBMNumber, ICNo, PatientName, PLNCode';
    IF @case = 96 SELECT @b = N'SELECT PAYCode, DateVisit, Count(CASNumber) AS Total From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, DateVisit ORDER BY DateVisit';
    IF @case = 97 SELECT @b = N'SELECT PAYCode, Month(DateVisit) As Month,
          						COUNT(CASNumber) AS TotalCases
          						FROM VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, MONTH(DateVisit), YEAR(DateVisit) ORDER BY MONTH(DateVisit)';
    IF @case = 98 SELECT @b = N'SELECT PAYCode, STACode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
        						SUM(clmtotalapproved) As TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY PAYCode, STACode';
    IF @case = 99 SELECT @b = N'SELECT SEX, Ageband, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) 
         						 AS totalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By SEX, Ageband';
    IF @case = 100 SELECT @b = N'SELECT SEX, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By SEX';
    IF @case = 101 SELECT @b = N'SELECT INSCode, PLNCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						 AS totalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By INSCode, PLNCode, Sex';
    IF @case = 102 SELECT @b = N'SELECT RACCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						 AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By RACCode, Sex';
    IF @case = 103 SELECT @b = N'SELECT COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved,
          						Sum(CLMTotalActual) As TotalActual, ClinicCode, CLNName, 
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY CASTreatReason1, ClinicCode, CLNName';
    IF @case = 104 SELECT @b = N'SELECT Top 20 COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved,
              					Sum(CLMTotalActual) As TotalActual, SUM(MCDays) AS MCDays
              					From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY CASTreatReason1 ORDER BY Total DESC';
    IF @case = 105 SELECT @b = N'SELECT Top 20 Sum(CLMTotalApproved) As TotalApproved, COUNT(CASTreatReason1) AS Total, CASTreatReason1, 
           						Sum(CLMTotalActual) As TotalActual, SUM(MCDays) AS MCDays 
             						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'Group By CASTreatReason1 Order By TotalApproved Desc';
    IF @case = 106 SELECT @b = N'SELECT COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved, 
            						Sum(CLMTotalActual) As TotalActual, SUM(MCDays) AS MCDays
              					From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY CASTreatReason1 ORDER BY Total DESC';
    IF @case = 107 SELECT @b = N'SELECT MBMDepartment, CLNType,BordxRefNo,
            						SUM(CLMTotalActual) as CLMTotalActual, SUM(CLMNonGPFee) as CLMNonGPFee
            						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMDepartment, CLNType,BordxRefNo,BordxDate';
    IF @case = 108 SELECT @b = N'SELECT MBMDepartment, MBMNumber, CoverID, PatientName, CLNType, 
          						SUM(CLMTotalActual) as CLMTotalActual, SUM(CLMNonGPFee) as CLMNonGPFee 
          						From VIEW_CLMReport_Simple Where 0 = 0', @a = N'GROUP BY MBMDepartment, MBMNumber, CoverID, CLNType, PatientName';
    IF @b IS NULL RETURN;
    DECLARE @w nvarchar(max) = N'';
    IF @hlt IS NOT NULL SET @w += N' AND HLTCode = @hlt';
    IF @hltBlank = 1 SET @w += N' AND (HLTCode IS NULL OR HLTCode = '''')';
    IF @ins IS NOT NULL SET @w += N' AND INSCode = @ins';
    IF @insBlank = 1 SET @w += N' AND (INSCode IS NULL OR INSCode = '''')';
    IF @memType IS NOT NULL SET @w += CASE @memType WHEN 'P|Both' THEN N' AND (MBMMemberType = ''P'' Or MBMMemberType = ''Q'') And CoverID = ''00'' ' WHEN 'P|C' THEN N' AND MBMMemberType = ''P''  And CoverID = ''00''' WHEN 'P|S' THEN N' AND MBMMemberType = ''Q'' And CoverID = ''00''' WHEN 'S|Both' THEN N' AND (MBMCMemberType = ''S'' OR MBMCMemberType = ''C'')' WHEN 'S|C' THEN N' AND MBMCMemberType = ''C''' WHEN 'S|S' THEN N' AND MBMCMemberType = ''S''' WHEN 'Both|C' THEN N' AND (MBMMemberType = ''P'' OR MBMCMemberType = ''C'')' WHEN 'Both|S' THEN N' AND (MBMMemberType = ''Q'' OR MBMCMemberType = ''S'')' ELSE N'' END;
    IF @race IS NOT NULL SET @w += N' AND RACCode = @race';
    IF @raceBlank = 1 SET @w += N' AND (RACCode IS NULL OR RACCode = '''')';
    IF @sex IS NOT NULL SET @w += N' AND Sex = @sex';
    IF @sexBlank = 1 SET @w += N' AND (Sex IS NULL OR Sex = '''')';
    IF @relation IS NOT NULL SET @w += N' AND Relation = @relation';
    IF @relationBlank = 1 SET @w += N' AND (Relation IS NULL OR Relation = '''')';
    IF @dataStatus IS NOT NULL SET @w += N' AND MBMDataStatus = @dataStatus';
    IF @dataStatusBlank = 1 SET @w += N' AND (MBMDataStatus IS NULL OR MBMDataStatus = '''')';
    IF @claimability IS NOT NULL SET @w += N' AND Claimability = @claimability';
    IF @claimabilityBlank = 1 SET @w += N' AND (Claimability IS NULL OR Claimability = '''')';
    IF @polStatus IS NOT NULL SET @w += N' AND MBMStatus = @polStatus';
    IF @polStatusBlank = 1 SET @w += N' AND (MBMStatus IS NULL OR MBMStatus = '''')';
    IF @expired IS NOT NULL SET @w += N' AND MBMExpiredStatus = @expired';
    IF @expiredBlank = 1 SET @w += N' AND (MBMExpiredStatus IS NULL OR MBMExpiredStatus = '''')';
    IF @premMode IS NOT NULL SET @w += N' AND MBMInstallment = @premMode';
    IF @premModeBlank = 1 SET @w += N' AND (MBMInstallment IS NULL OR MBMInstallment = '''')';
    IF @takeover IS NOT NULL SET @w += N' AND MBMTakeOver = @takeover';
    IF @takeoverBlank = 1 SET @w += N' AND (MBMTakeOver IS NULL OR MBMTakeOver = '''')';
    IF @caseStatus IS NOT NULL SET @w += CASE @caseStatus WHEN '0' THEN N' AND CASStatus = ''O''' WHEN '1' THEN N' AND (CASStatus = ''O'' OR CASStatus = ''C'')' WHEN '2' THEN N' AND CASStatus = ''D''' ELSE N'' END;
    IF @caseStatusBlank = 1 SET @w += N' AND (CASStatus IS NULL OR CASStatus = '''')';
    IF @state IS NOT NULL SET @w += N' AND STACode = @state';
    IF @stateBlank = 1 SET @w += N' AND (STACode IS NULL OR STACode = '''')';
    IF @departmentBlank = 1 SET @w += N' AND (MBMDepartment IS NULL OR MBMDepartment = '''')';
    IF @renewal IS NOT NULL SET @w += CASE @renewal WHEN 'P' THEN N' AND (MBMRenewFlag = ''P'')' WHEN 'N' THEN N' AND (MBMRenewFlag Is null OR MBMRenewFlag = '''' OR MBMRenewFlag = ''N'') ' WHEN 'R' THEN N' AND (MBMRenewFlag = ''R'')' WHEN 'T' THEN N' AND (MBMRenewFlag = ''T'')' ELSE N'' END;
    IF @renewalBlank = 1 SET @w += N' AND (MBMRenewFlag IS NULL OR MBMRenewFlag = '''')';
    IF @agent IS NOT NULL SET @w += N' AND MBMAgentCode LIKE ''%'' + @agent + ''%''';
    IF @agentBlank = 1 SET @w += N' AND (MBMAgentCode IS NULL OR MBMAgentCode = '''')';
    IF @policy IS NOT NULL SET @w += N' AND MBMPolicyNo LIKE ''%'' + @policy + ''%''';
    IF @policyBlank = 1 SET @w += N' AND (MBMPolicyNo IS NULL OR MBMPolicyNo = '''')';
    IF @patient IS NOT NULL SET @w += N' AND PatientName LIKE ''%'' + @patient + ''%''';
    IF @agebandFrom IS NOT NULL SET @w += N' AND AGEBand >= @agebandFrom';
    IF @agebandTo IS NOT NULL SET @w += N' AND AGEBand <= @agebandTo';
    IF @polYearFrom IS NOT NULL SET @w += N' AND MBMPolicyYear >= @polYearFrom';
    IF @polYearTo IS NOT NULL SET @w += N' AND MBMPolicyYear <= @polYearTo';
    IF @polYearBlank = 1 SET @w += N' AND (MBMPolicyYear IS NULL OR MBMPolicyYear = '''')';
    IF @mcDaysFrom IS NOT NULL SET @w += N' AND MCDays >= @mcDaysFrom';
    IF @mcDaysTo IS NOT NULL SET @w += N' AND MCDays <= @mcDaysTo';
    IF @mcDaysBlank = 1 SET @w += N' AND (MCDays IS NULL OR MCDays = '''')';
    IF @payor IS NOT NULL SET @w += N' AND PAYCode = @payor';
    IF @payorBlank = 1 SET @w += N' AND (PAYCode IS NULL OR PAYCode = '''')';
    IF @company IS NOT NULL SET @w += N' AND MBMGroupCompany = @company';
    IF @companyBlank = 1 SET @w += N' AND (MBMGroupCompany IS NULL OR MBMGroupCompany = '''')';
    IF @clinic IS NOT NULL SET @w += N' AND ClinicCode = @clinic';
    IF @clinicBlank = 1 SET @w += N' AND (CLNName IS NULL OR CLNName = '''')';
    IF @clinicGroup IS NOT NULL SET @w += N' AND ClinicGRPName = @clinicGroup';
    IF @branch IS NOT NULL SET @w += N' AND BRNCode = @branch';
    IF @dovFrom IS NOT NULL SET @w += N' AND DateVisit >= @dovFrom';
    IF @dovTo IS NOT NULL SET @w += N' AND DateVisit <= @dovTo';
    IF @dovBlank = 1 SET @w += N' AND (DateVisit IS NULL OR DateVisit = '''')';
    IF @dorFrom IS NOT NULL SET @w += N' AND RegisterDate >= @dorFrom';
    IF @dorTo IS NOT NULL SET @w += N' AND RegisterDate <= @dorTo';
    IF @dorBlank = 1 SET @w += N' AND (RegisterDate IS NULL OR RegisterDate = '''')';
    IF @dowFrom IS NOT NULL SET @w += N' AND ClaimOpenDate >= @dowFrom';
    IF @dowTo IS NOT NULL SET @w += N' AND ClaimOpenDate <= @dowTo';
    IF @dowBlank = 1 SET @w += N' AND (ClaimOpenDate IS NULL OR ClaimOpenDate = '''')';
    IF @invoiceDateFrom IS NOT NULL SET @w += N' AND InvoiceDate >= @invoiceDateFrom';
    IF @invoiceDateTo IS NOT NULL SET @w += N' AND InvoiceDate <= @invoiceDateTo';
    IF @invoiceDateBlank = 1 SET @w += N' AND (InvoiceDate IS NULL OR InvoiceDate = '''')';
    IF @treatmentFrom IS NOT NULL SET @w += N' AND CASTreatReason1 >= @treatmentFrom';
    IF @treatmentTo IS NOT NULL SET @w += N' AND CASTreatReason1 <= @treatmentTo';
    IF @treatmentBlank = 1 SET @w += N' AND (CASTreatReason1 IS NULL OR CASTreatReason1 = '''')';
    IF @planFrom IS NOT NULL SET @w += N' AND PLNCode >= @planFrom';
    IF @planTo IS NOT NULL SET @w += N' AND PLNCode <= @planTo';
    IF @planBlank = 1 SET @w += N' AND (PLNCode IS NULL OR PLNCode = '''')';
    IF @effFrom IS NOT NULL SET @w += N' AND MBMPayorEffDate >= @effFrom';
    IF @effTo IS NOT NULL SET @w += N' AND MBMPayorEffDate <= @effTo';
    IF @effBlank = 1 SET @w += N' AND (MBMPayorEffDate IS NULL OR MBMPayorEffDate = '''')';
    IF @expFrom IS NOT NULL SET @w += N' AND MBMPayorExpDate >= @expFrom';
    IF @expTo IS NOT NULL SET @w += N' AND MBMPayorExpDate <= @expTo';
    IF @expBlank = 1 SET @w += N' AND (MBMPayorExpDate IS NULL OR MBMPayorExpDate = '''')';
    IF @joinedFrom IS NOT NULL SET @w += N' AND MBMFirstJoinedDate >= @joinedFrom';
    IF @joinedTo IS NOT NULL SET @w += N' AND MBMFirstJoinedDate <= @joinedTo';
    IF @joinedBlank = 1 SET @w += N' AND (MBMFirstJoinedDate IS NULL OR MBMFirstJoinedDate = '''')';
    IF @dobFrom IS NOT NULL SET @w += N' AND DOB >= @dobFrom';
    IF @dobTo IS NOT NULL SET @w += N' AND DOB <= @dobTo';
    IF @dobBlank = 1 SET @w += N' AND (DOB IS NULL OR DOB = '''')';
    IF @mbmNoFrom IS NOT NULL SET @w += N' AND MBMNumber >= @mbmNoFrom';
    IF @mbmNoTo IS NOT NULL SET @w += N' AND MBMNumber <= @mbmNoTo';
    IF @caseNoFrom IS NOT NULL SET @w += N' AND CASNumber >= @caseNoFrom';
    IF @caseNoTo IS NOT NULL SET @w += N' AND CASNumber <= @caseNoTo';
    IF @bordxFrom IS NOT NULL AND @bordxType = 'INSURER' SET @w += N' AND BordxRefNo >= @bordxFrom'; IF @bordxFrom IS NOT NULL AND @bordxType = 'ASO' SET @w += N' AND BTBGBordxRefNo >= @bordxFrom';
    IF @bordxTo IS NOT NULL AND @bordxType = 'INSURER' SET @w += N' AND BordxRefNo <= @bordxTo'; IF @bordxTo IS NOT NULL AND @bordxType = 'ASO' SET @w += N' AND BTBGBordxRefNo <= @bordxTo';
    IF @bordxBlank = 1 AND @bordxType = 'INSURER' SET @w += N' AND (BordxRefNo IS NULL OR BordxRefNo = '''')'; IF @bordxBlank = 1 AND @bordxType = 'ASO' SET @w += N' AND (BTBGBordxRefNo IS NULL OR BTBGBordxRefNo = '''')';
    IF @bordxDateFrom IS NOT NULL SET @w += N' AND BordxDate >= @bordxDateFrom';
    IF @bordxDateTo IS NOT NULL SET @w += N' AND BordxDate <= @bordxDateTo';
    IF @bordxDateBlank = 1 SET @w += N' AND (BordxDate IS NULL OR BordxDate = '''')';
    IF @icFrom IS NOT NULL SET @w += N' AND ICNo >= @icFrom';
    IF @icTo IS NOT NULL SET @w += N' AND ICNo <= @icTo';
    IF @icBlank = 1 SET @w += N' AND (ICNo IS NULL OR ICNo = '''')';
    IF @mbmIcFrom IS NOT NULL SET @w += N' AND MBMIcBcPp >= @mbmIcFrom';
    IF @mbmIcTo IS NOT NULL SET @w += N' AND MBMIcBcPp <= @mbmIcTo';
    IF @mbmIcBlank = 1 SET @w += N' AND (MBMIcBcPp IS NULL OR MBMIcBcPp = '''')';
    IF @actualFrom IS NOT NULL SET @w += N' AND CLMTotalActual >= @actualFrom';
    IF @actualTo IS NOT NULL SET @w += N' AND CLMTotalActual <= @actualTo';
    IF @approvedFrom IS NOT NULL SET @w += N' AND CLMTotalApproved >= @approvedFrom';
    IF @approvedTo IS NOT NULL SET @w += N' AND CLMTotalApproved <= @approvedTo';
    IF @osUtilizedFrom IS NOT NULL SET @w += N' AND OSUtilized >= @osUtilizedFrom';
    IF @osUtilizedTo IS NOT NULL SET @w += N' AND OSUtilized <= @osUtilizedTo';
    IF @panel IS NOT NULL SET @w += CASE @panel WHEN 'C' THEN N' AND CLNPanel = ''C''' WHEN 'Y' THEN N' AND CLNPanel = ''Y''' ELSE N'' END;
    IF @nonPanel = 1 SET @w += N' AND (CLNPanel <> ''C'' OR CLNPanel IS NULL OR CLNPanel = '''')';
    IF @pending IS NOT NULL SET @w += CASE @pending WHEN 'N' THEN N' AND (CLMPending = ''N'' OR CLMPending = '''' OR CLMPending IS NULL)' WHEN 'Y' THEN N' AND CLMPending = ''Y''' ELSE N'' END;
    IF @year IS NOT NULL SET @w += N' AND DOVYear = @year';
    IF @invoiceFrom IS NOT NULL SET @w += N' AND InvoiceNo >= @invoiceFrom';
    IF @invoiceTo IS NOT NULL SET @w += N' AND InvoiceNo <= @invoiceTo';
    IF @invoiceBlank = 1 SET @w += N' AND (InvoiceNo IS NULL OR InvoiceNo = '''')';
    IF @submissionFrom IS NOT NULL SET @w += N' AND BillNo >= @submissionFrom';
    IF @submissionTo IS NOT NULL SET @w += N' AND BillNo <= @submissionTo';
    IF @submissionBlank = 1 SET @w += N' AND (BillNo IS NULL OR BillNo = '''')';
    IF @opBatchFrom IS NOT NULL SET @w += N' AND OPBatchNo >= @opBatchFrom';
    IF @opBatchTo IS NOT NULL SET @w += N' AND OPBatchNo <= @opBatchTo';
    IF @clinicCodeFrom IS NOT NULL SET @w += N' AND ClinicCode >= @clinicCodeFrom';
    IF @clinicCodeTo IS NOT NULL SET @w += N' AND ClinicCode <= @clinicCodeTo';
    IF @clinicType IS NOT NULL SET @w += CASE @clinicType WHEN '0' THEN N' AND CLNType = ''S''' WHEN '1' THEN N' AND CLNType = ''G''' WHEN '2' THEN N' AND (CLNType = '''' OR CLNType IS NULL)' ELSE N'' END;
    IF @payTo IS NOT NULL SET @w += N' AND INVPayTo = @payTo';
    IF @vsAmt = 1 SET @w += N' AND VSAmt <> 0';
    IF @deptSelection IS NOT NULL SET @w += N' AND MBMDepartment = @deptSelection';
    IF @ieInvoice = 1 SET @w += N' AND substring(InvoiceNo,1,2)=''IE'' ';
    DECLARE @sql nvarchar(max) = CAST(@b AS nvarchar(max)) + @w + N' ' + @a;
    EXEC CISWorksheet.sys.sp_executesql @sql, N'@hlt nvarchar(1), @hltBlank bit, @ins nvarchar(1), @insBlank bit, @memType nvarchar(10), @race nvarchar(1), @raceBlank bit, @sex nvarchar(1), @sexBlank bit, @relation nvarchar(1), @relationBlank bit, @dataStatus nvarchar(1), @dataStatusBlank bit, @claimability nvarchar(1), @claimabilityBlank bit, @polStatus nvarchar(1), @polStatusBlank bit, @expired nvarchar(1), @expiredBlank bit, @premMode nvarchar(1), @premModeBlank bit, @takeover nvarchar(1), @takeoverBlank bit, @caseStatus nvarchar(10), @caseStatusBlank bit, @state nvarchar(100), @stateBlank bit, @departmentBlank bit, @renewal nvarchar(10), @renewalBlank bit, @agent nvarchar(200), @agentBlank bit, @policy nvarchar(200), @policyBlank bit, @patient nvarchar(200), @agebandFrom nvarchar(20), @agebandTo nvarchar(20), @polYearFrom nvarchar(20), @polYearTo nvarchar(20), @polYearBlank bit, @mcDaysFrom nvarchar(20), @mcDaysTo nvarchar(20), @mcDaysBlank bit, @payor nvarchar(10), @payorBlank bit, @company nvarchar(200), @companyBlank bit, @clinic nvarchar(20), @clinicBlank bit, @clinicGroup nvarchar(200), @branch nvarchar(20), @dovFrom datetime, @dovTo datetime, @dovBlank bit, @dorFrom datetime, @dorTo datetime, @dorBlank bit, @dowFrom datetime, @dowTo datetime, @dowBlank bit, @invoiceDateFrom datetime, @invoiceDateTo datetime, @invoiceDateBlank bit, @treatmentFrom nvarchar(20), @treatmentTo nvarchar(20), @treatmentBlank bit, @planFrom nvarchar(20), @planTo nvarchar(20), @planBlank bit, @effFrom datetime, @effTo datetime, @effBlank bit, @expFrom datetime, @expTo datetime, @expBlank bit, @joinedFrom datetime, @joinedTo datetime, @joinedBlank bit, @dobFrom datetime, @dobTo datetime, @dobBlank bit, @mbmNoFrom nvarchar(30), @mbmNoTo nvarchar(30), @caseNoFrom nvarchar(20), @caseNoTo nvarchar(20), @bordxType nvarchar(10), @bordxFrom nvarchar(20), @bordxTo nvarchar(20), @bordxBlank bit, @bordxDateFrom datetime, @bordxDateTo datetime, @bordxDateBlank bit, @icFrom nvarchar(30), @icTo nvarchar(30), @icBlank bit, @mbmIcFrom nvarchar(30), @mbmIcTo nvarchar(30), @mbmIcBlank bit, @actualFrom money, @actualTo money, @approvedFrom money, @approvedTo money, @osUtilizedFrom money, @osUtilizedTo money, @panel nvarchar(10), @nonPanel bit, @pending nvarchar(10), @year nvarchar(4), @invoiceFrom nvarchar(50), @invoiceTo nvarchar(50), @invoiceBlank bit, @submissionFrom nvarchar(50), @submissionTo nvarchar(50), @submissionBlank bit, @opBatchFrom nvarchar(30), @opBatchTo nvarchar(30), @clinicCodeFrom nvarchar(20), @clinicCodeTo nvarchar(20), @clinicType nvarchar(10), @payTo nvarchar(1), @vsAmt bit, @deptSelection nvarchar(200), @ieInvoice bit',
        @hlt, @hltBlank, @ins, @insBlank, @memType, @race, @raceBlank, @sex, @sexBlank, @relation, @relationBlank, @dataStatus, @dataStatusBlank, @claimability, @claimabilityBlank, @polStatus, @polStatusBlank, @expired, @expiredBlank, @premMode, @premModeBlank, @takeover, @takeoverBlank, @caseStatus, @caseStatusBlank, @state, @stateBlank, @departmentBlank, @renewal, @renewalBlank, @agent, @agentBlank, @policy, @policyBlank, @patient, @agebandFrom, @agebandTo, @polYearFrom, @polYearTo, @polYearBlank, @mcDaysFrom, @mcDaysTo, @mcDaysBlank, @payor, @payorBlank, @company, @companyBlank, @clinic, @clinicBlank, @clinicGroup, @branch, @dovFrom, @dovTo, @dovBlank, @dorFrom, @dorTo, @dorBlank, @dowFrom, @dowTo, @dowBlank, @invoiceDateFrom, @invoiceDateTo, @invoiceDateBlank, @treatmentFrom, @treatmentTo, @treatmentBlank, @planFrom, @planTo, @planBlank, @effFrom, @effTo, @effBlank, @expFrom, @expTo, @expBlank, @joinedFrom, @joinedTo, @joinedBlank, @dobFrom, @dobTo, @dobBlank, @mbmNoFrom, @mbmNoTo, @caseNoFrom, @caseNoTo, @bordxType, @bordxFrom, @bordxTo, @bordxBlank, @bordxDateFrom, @bordxDateTo, @bordxDateBlank, @icFrom, @icTo, @icBlank, @mbmIcFrom, @mbmIcTo, @mbmIcBlank, @actualFrom, @actualTo, @approvedFrom, @approvedTo, @osUtilizedFrom, @osUtilizedTo, @panel, @nonPanel, @pending, @year, @invoiceFrom, @invoiceTo, @invoiceBlank, @submissionFrom, @submissionTo, @submissionBlank, @opBatchFrom, @opBatchTo, @clinicCodeFrom, @clinicCodeTo, @clinicType, @payTo, @vsAmt, @deptSelection, @ieInvoice;
END
GO
