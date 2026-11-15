-- Claim / Bordx > Claim Tracking 2 (FrmClaimRpt2 and its Jupiter copy): the 39 cases of CISWorksheet CISClaimTracking2 (generated from the procedure by
-- gen-trk.js). Each case's query text is kept as the desktop's; the pasted criteria are replaced by a WHERE fragment built only from the filters that are set,
-- with every value passed as a parameter.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.ClmTrk2_Run @case int,
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
    @caseStatus nvarchar(1) = NULL,
    @caseStatusBlank bit = 0,
    @expired nvarchar(1) = NULL,
    @expiredBlank bit = 0,
    @premMode nvarchar(1) = NULL,
    @premModeBlank bit = 0,
    @takeover nvarchar(1) = NULL,
    @takeoverBlank bit = 0,
    @state nvarchar(100) = NULL,
    @stateBlank bit = 0,
    @renewal nvarchar(10) = NULL,
    @renewalBlank bit = 0,
    @agent nvarchar(200) = NULL,
    @agentBlank bit = 0,
    @policy nvarchar(200) = NULL,
    @policyBlank bit = 0,
    @patient nvarchar(200) = NULL,
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
    @clinicName nvarchar(200) = NULL,
    @clinicNameBlank bit = 0,
    @clinicGroup nvarchar(200) = NULL,
    @clinicGroupBlank bit = 0,
    @branch nvarchar(20) = NULL,
    @branchBlank bit = 0,
    @dovFrom datetime = NULL,
    @dovTo datetime = NULL,
    @dovBlank bit = 0,
    @dorFrom datetime = NULL,
    @dorTo datetime = NULL,
    @dorBlank bit = 0,
    @dowFrom datetime = NULL,
    @dowTo datetime = NULL,
    @dowBlank bit = 0,
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
    @invoiceFrom nvarchar(50) = NULL,
    @invoiceTo nvarchar(50) = NULL,
    @invoiceBlank bit = 0,
    @submissionFrom nvarchar(50) = NULL,
    @submissionTo nvarchar(50) = NULL,
    @submissionBlank bit = 0
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @b nvarchar(max), @a nvarchar(max) = N'';
    IF @case = 1 SELECT @b = N'SELECT AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by AGEBand';
    IF @case = 3 SELECT @b = N'SELECT AGEBand, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by AGEBand, Sex';
    IF @case = 4 SELECT @b = N'SELECT BordxRefNo, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by BordxRefNo Order by BordxRefNo';
    IF @case = 5 SELECT @b = N'SELECT CASNumber, MBMNumber, PatientName, DateVisit,
          						CLMGPFeeBasicMed, CLMGPFeeMedInj, CLMGPFeeLTMed, CLMSPFeeBasicMed,
          						CLMSPFeeMedInj, CLMSPFeeLTMed, CLMXRay, CLMUltraSound, CLMECG, CLMDressing,
          						CLMNebulizer, CLMID, CLMTSSTO, CLMMinorSurgery, CLMPOthers,
          						CLMBloodDrawing, CLMBloodGlucose, CLMUrine, CLMPapSmear, CLMLOthers,
          						CLMDTPPolio, CLMDTPPolioHIB, CLMHepatitisB, CLMMMR, CLMIOthers,
          						CLMDental, CLMMaternity, CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'';
    IF @case = 6 SELECT @b = N'SELECT ClaimNatureCode, SUM(CONVERT(Money,CLMBoardAmt)) AS CLMBoardAmt, SUM(CONVERT(Money,CLMICUAmt)) AS CLMICUAmt,
          						SUM(CONVERT(Money,CLMServices)) AS CLMServices, SUM(CONVERT(Money,CLMTheatre)) AS CLMTheatre, SUM(CONVERT(Money,CLMPreHSAmt)) AS CLMPreHSAmt,
          						SUM(CONVERT(Money,CLMSurgicalAmt)) AS CLMSurgicalAmt, SUM(CONVERT(Money,CLMAna)) AS CLMAna, SUM(CONVERT(Money,CLMNSPreHos)) AS CLMNSPreHos,
          						SUM(CONVERT(Money,CLMPhysician)) AS CLMPhysician, SUM(CONVERT(Money,CLMPost)) AS CLMPost, SUM(CONVERT(Money,CLMAmbulance)) AS CLMAmbulance,
          						SUM(CONVERT(Money,CLMOutPatient)) AS CLMOutPatient, SUM(CONVERT(Money,CLMMonthly)) AS CLMMonthly, SUM(CONVERT(Money,CLMOrgan)) AS CLMOrgan,
          						SUM(CONVERT(Money,CLMMedical)) AS CLMMedical, SUM(CONVERT(Money,CLMPhone)) AS CLMPhone, SUM(CONVERT(Money,CLMGuardianAmt)) AS CLMGuardianAmt,
          						SUM(CONVERT(Money,CLMAdmin)) AS CLMAdmin, SUM(CONVERT(Money,CLMTax)) AS CLMTax, SUM(CONVERT(Money,CLMOthers2)) AS CLMOthers2, SUM(CONVERT(Money,CLMCashAmt)) AS CLMCashAmt
          						From VIEW_CLMReport_Act Where 0 = 0', @a = N'Group by ClaimNatureCode Order by ClaimNatureCode';
    IF @case = 7 SELECT @b = N'SELECT ClinicCode, CLNName, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
        						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by ClinicCode, CLNName Order by ClinicCode, CLNName';
    IF @case = 8 SELECT @b = N'SELECT CASTreatReason1, ReasonDesc, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by CASTreatReason1, ReasonDesc Order by CASTreatReason1, ReasonDesc';
    IF @case = 9 SELECT @b = N'SELECT MBMGroupCompany, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by MBMGroupCompany Order by MBMGroupCompany';
    IF @case = 10 SELECT @b = N'SELECT INSCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by INSCode Order by INSCode';
    IF @case = 11 SELECT @b = N'SELECT INSCode, AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
        						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by INSCode, AGEBand Order by INSCode, Ageband';
    IF @case = 12 SELECT @b = N'SELECT INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
         						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, 
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by INSCode, PLNCode Order by INSCode, PLNCode';
    IF @case = 13 SELECT @b = N'SELECT INSCode, Relation, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, 
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by INSCode, Relation Order by INSCode, Relation';
    IF @case = 14 SELECT @b = N'SELECT MBMPolicyYear, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, 
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by MBMPolicyYear Order by MBMPolicyYear';
    IF @case = 15 SELECT @b = N'SELECT INSCode, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by INSCode, Sex Order by INSCode, Sex';
    IF @case = 16 SELECT @b = N'SELECT PAYCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by PAYCode Order by PAYCode ';
    IF @case = 17 SELECT @b = N'SELECT PAYCode, INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, 
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, 
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by PAYCode, INSCode, PLNCode Order by PAYCode, INSCode, PLNCode';
    IF @case = 18 SELECT @b = N'SELECT Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
         						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by Sex Order by Sex';
    IF @case = 19 SELECT @b = N'SELECT MBMTakeOver, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by MBMTakeOver Order by MBMTakeOver';
    IF @case = 20 SELECT @b = N'SELECT AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by AGEBand Order by AGEBand';
    IF @case = 21 SELECT @b = N'SELECT AGEBand, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by AGEBand, Sex Order by Ageband, Sex';
    IF @case = 22 SELECT @b = N'SELECT BordxRefNo, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by BordxRefNo Order by BordxRefNo';
    IF @case = 23 SELECT @b = N'SELECT CASNumber, MBMNumber, PatientName, DateVisit,
          						CLMGPFeeBasicMed, CLMGPFeeMedInj, CLMGPFeeLTMed, CLMSPFeeBasicMed,
          						CLMSPFeeMedInj, CLMSPFeeLTMed, CLMXRay, CLMUltraSound, CLMECG, CLMDressing,
          						CLMNebulizer, CLMID, CLMTSSTO, CLMMinorSurgery, CLMPOthers,
          						CLMBloodDrawing, CLMBloodGlucose, CLMUrine, CLMPapSmear, CLMLOthers,
          						CLMDTPPolio, CLMDTPPolioHIB, CLMHepatitisB, CLMMMR, CLMIOthers,
          						CLMDental, CLMMaternity, CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'';
    IF @case = 24 SELECT @b = N'SELECT ClinicCode, CLNName, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, 
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by ClinicCode, CLNName Order by ClinicCode, CLNName';
    IF @case = 25 SELECT @b = N'SELECT CASTreatReason1, ReasonDesc, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by CASTreatReason1, ReasonDesc Order by CASTreatReason1, ReasonDesc';
    IF @case = 26 SELECT @b = N'SELECT MBMGroupCompany, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, 
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup 
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by MBMGroupCompany Order by MBMGroupCompany';
    IF @case = 27 SELECT @b = N'SELECT INSCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by INSCode Order by INSCode';
    IF @case = 28 SELECT @b = N'SELECT INSCode, AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by INSCode, AGEBand Order by INSCode, Ageband';
    IF @case = 29 SELECT @b = N'SELECT INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by INSCode, PLNCode Order by INSCode, PLNCode';
    IF @case = 30 SELECT @b = N'SELECT INSCode, Relation, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by INSCode, Relation Order by INSCode, Relation';
    IF @case = 31 SELECT @b = N'SELECT INSCode, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by INSCode, Sex Order by INSCode, Sex';
    IF @case = 32 SELECT @b = N'SELECT PAYCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, 
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by PAYCode Order by PAYCode';
    IF @case = 33 SELECT @b = N'SELECT PAYCode, INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, 
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by PAYCode, INSCode, PLNCode Order by PAYCode, INSCode, PLNCode';
    IF @case = 34 SELECT @b = N'SELECT MBMPolicyYear, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, 
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by MBMPolicyYear Order by MBMPolicyYear';
    IF @case = 35 SELECT @b = N'SELECT Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, 
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by Sex Order by Sex';
    IF @case = 36 SELECT @b = N'SELECT MBMTakeOver, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup 
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'Group by MBMTakeOver Order by MBMTakeOver';
    IF @case = 37 SELECT @b = N'SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'';
    IF @case = 38 SELECT @b = N'SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup 
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'';
    IF @case = 39 SELECT @b = N'SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, 
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, 
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup 
          						From VIEW_CLMReport_Clinic Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'';
    IF @case = 40 SELECT @b = N'SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_MBM Where SUBSTRING(PAYCode,1,1) <> ''Z''', @a = N'';
    IF @b IS NULL RETURN;
    DECLARE @w nvarchar(max) = N'';
    IF @hlt IS NOT NULL SET @w += N' AND HLTCode = @hlt';
    IF @hltBlank = 1 SET @w += N' AND (HLTCode IS NULL OR HLTCode = '''')';
    IF @ins IS NOT NULL SET @w += N' AND INSCode = @ins';
    IF @insBlank = 1 SET @w += N' AND (INSCode IS NULL OR INSCode = '''')';
    IF @memType IS NOT NULL SET @w += CASE @memType WHEN 'P|Both' THEN N' AND (MBMMemberType = ''P'' Or MBMMemberType = ''Q'') And MBMPatientCovID = ''00'' ' WHEN 'P|C' THEN N' AND MBMMemberType = ''P''  And MBMPatientCovID = ''00''' WHEN 'P|S' THEN N' AND MBMMemberType = ''Q'' And MBMPatientCovID = ''00''' WHEN 'S|Both' THEN N' AND (MBMCMemberType = ''S'' OR MBMCMemberType = ''C'')' WHEN 'S|C' THEN N' AND MBMCMemberType = ''C''' WHEN 'S|S' THEN N' AND MBMCMemberType = ''S''' WHEN 'Both|C' THEN N' AND (MBMMemberType = ''P'' OR MBMCMemberType = ''C'')' WHEN 'Both|S' THEN N' AND (MBMMemberType = ''Q'' OR MBMCMemberType = ''S'')' ELSE N'' END;
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
    IF @caseStatus IS NOT NULL SET @w += N' AND CASStatus = @caseStatus';
    IF @caseStatusBlank = 1 SET @w += N' AND (CASStatus IS NULL OR CASStatus = '''')';
    IF @expired IS NOT NULL SET @w += N' AND MBMExpiredStatus = @expired';
    IF @expiredBlank = 1 SET @w += N' AND (MBMExpiredStatus IS NULL OR MBMExpiredStatus = '''')';
    IF @premMode IS NOT NULL SET @w += N' AND MBMInstallment = @premMode';
    IF @premModeBlank = 1 SET @w += N' AND (MBMInstallment IS NULL OR MBMInstallment = '''')';
    IF @takeover IS NOT NULL SET @w += N' AND MBMTakeOver = @takeover';
    IF @takeoverBlank = 1 SET @w += N' AND (MBMTakeOver IS NULL OR MBMTakeOver = '''')';
    IF @state IS NOT NULL SET @w += N' AND STACode = @state';
    IF @stateBlank = 1 SET @w += N' AND (STACode IS NULL OR STACode = '''')';
    IF @renewal IS NOT NULL SET @w += CASE @renewal WHEN 'P' THEN N' AND (MBMRenewFlag = ''P'')' WHEN 'N' THEN N' AND (MBMRenewFlag Is null OR MBMRenewFlag = '''' OR MBMRenewFlag = ''N'') ' WHEN 'R' THEN N' AND (MBMRenewFlag = ''R'')' WHEN 'T' THEN N' AND (MBMRenewFlag = ''T'')' ELSE N'' END;
    IF @renewalBlank = 1 SET @w += N' AND (MBMRenewFlag IS NULL OR MBMRenewFlag = '''')';
    IF @agent IS NOT NULL SET @w += N' AND MBMAgentCode LIKE ''%'' + @agent + ''%''';
    IF @agentBlank = 1 SET @w += N' AND (MBMAgentCode IS NULL OR MBMAgentCode = '''')';
    IF @policy IS NOT NULL SET @w += N' AND MBMPolicyNo LIKE ''%'' + @policy + ''%''';
    IF @policyBlank = 1 SET @w += N' AND (MBMPolicyNo IS NULL OR MBMPolicyNo = '''')';
    IF @patient IS NOT NULL SET @w += N' AND PatientName LIKE ''%'' + @patient + ''%''';
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
    IF @clinicName IS NOT NULL SET @w += N' AND HOSName = @clinicName';
    IF @clinicNameBlank = 1 SET @w += N' AND (HOSName IS NULL OR HOSName = '''')';
    IF @clinicGroup IS NOT NULL SET @w += N' AND HOSGRPName = @clinicGroup';
    IF @clinicGroupBlank = 1 SET @w += N' AND (HOSGRPName IS NULL OR HOSGRPName = '''')';
    IF @branch IS NOT NULL SET @w += N' AND BRNCode = @branch';
    IF @branchBlank = 1 SET @w += N' AND (BRNCode IS NULL OR BRNCode = '''')';
    IF @dovFrom IS NOT NULL SET @w += N' AND DateVisit >= @dovFrom';
    IF @dovTo IS NOT NULL SET @w += N' AND DateVisit <= @dovTo';
    IF @dovBlank = 1 SET @w += N' AND (DateVisit IS NULL OR DateVisit = '''')';
    IF @dorFrom IS NOT NULL SET @w += N' AND RegisterDate >= @dorFrom';
    IF @dorTo IS NOT NULL SET @w += N' AND RegisterDate <= @dorTo';
    IF @dorBlank = 1 SET @w += N' AND (RegisterDate IS NULL OR RegisterDate = '''')';
    IF @dowFrom IS NOT NULL SET @w += N' AND ClaimOpenDate >= @dowFrom';
    IF @dowTo IS NOT NULL SET @w += N' AND ClaimOpenDate <= @dowTo';
    IF @dowBlank = 1 SET @w += N' AND (ClaimOpenDate IS NULL OR ClaimOpenDate = '''')';
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
    IF @bordxFrom IS NOT NULL SET @w += N' AND BordxRefNo >= @bordxFrom';
    IF @bordxTo IS NOT NULL SET @w += N' AND BordxRefNo <= @bordxTo';
    IF @bordxBlank = 1 SET @w += N' AND (BordxRefNo IS NULL OR BordxRefNo = '''')';
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
    IF @invoiceFrom IS NOT NULL SET @w += N' AND InvoiceNo >= @invoiceFrom';
    IF @invoiceTo IS NOT NULL SET @w += N' AND InvoiceNo <= @invoiceTo';
    IF @invoiceBlank = 1 SET @w += N' AND (InvoiceNo IS NULL OR InvoiceNo = '''')';
    IF @submissionFrom IS NOT NULL SET @w += N' AND BillNo >= @submissionFrom';
    IF @submissionTo IS NOT NULL SET @w += N' AND BillNo <= @submissionTo';
    IF @submissionBlank = 1 SET @w += N' AND (BillNo IS NULL OR BillNo = '''')';
    DECLARE @sql nvarchar(max) = CAST(@b AS nvarchar(max)) + @w + N' ' + @a;
    EXEC CISWorksheet.sys.sp_executesql @sql, N'@hlt nvarchar(1), @hltBlank bit, @ins nvarchar(1), @insBlank bit, @memType nvarchar(10), @race nvarchar(1), @raceBlank bit, @sex nvarchar(1), @sexBlank bit, @relation nvarchar(1), @relationBlank bit, @dataStatus nvarchar(1), @dataStatusBlank bit, @claimability nvarchar(1), @claimabilityBlank bit, @polStatus nvarchar(1), @polStatusBlank bit, @caseStatus nvarchar(1), @caseStatusBlank bit, @expired nvarchar(1), @expiredBlank bit, @premMode nvarchar(1), @premModeBlank bit, @takeover nvarchar(1), @takeoverBlank bit, @state nvarchar(100), @stateBlank bit, @renewal nvarchar(10), @renewalBlank bit, @agent nvarchar(200), @agentBlank bit, @policy nvarchar(200), @policyBlank bit, @patient nvarchar(200), @polYearFrom nvarchar(20), @polYearTo nvarchar(20), @polYearBlank bit, @mcDaysFrom nvarchar(20), @mcDaysTo nvarchar(20), @mcDaysBlank bit, @payor nvarchar(10), @payorBlank bit, @company nvarchar(200), @companyBlank bit, @clinicName nvarchar(200), @clinicNameBlank bit, @clinicGroup nvarchar(200), @clinicGroupBlank bit, @branch nvarchar(20), @branchBlank bit, @dovFrom datetime, @dovTo datetime, @dovBlank bit, @dorFrom datetime, @dorTo datetime, @dorBlank bit, @dowFrom datetime, @dowTo datetime, @dowBlank bit, @treatmentFrom nvarchar(20), @treatmentTo nvarchar(20), @treatmentBlank bit, @planFrom nvarchar(20), @planTo nvarchar(20), @planBlank bit, @effFrom datetime, @effTo datetime, @effBlank bit, @expFrom datetime, @expTo datetime, @expBlank bit, @joinedFrom datetime, @joinedTo datetime, @joinedBlank bit, @dobFrom datetime, @dobTo datetime, @dobBlank bit, @mbmNoFrom nvarchar(30), @mbmNoTo nvarchar(30), @caseNoFrom nvarchar(20), @caseNoTo nvarchar(20), @bordxFrom nvarchar(20), @bordxTo nvarchar(20), @bordxBlank bit, @bordxDateFrom datetime, @bordxDateTo datetime, @bordxDateBlank bit, @icFrom nvarchar(30), @icTo nvarchar(30), @icBlank bit, @mbmIcFrom nvarchar(30), @mbmIcTo nvarchar(30), @mbmIcBlank bit, @actualFrom money, @actualTo money, @approvedFrom money, @approvedTo money, @invoiceFrom nvarchar(50), @invoiceTo nvarchar(50), @invoiceBlank bit, @submissionFrom nvarchar(50), @submissionTo nvarchar(50), @submissionBlank bit',
        @hlt, @hltBlank, @ins, @insBlank, @memType, @race, @raceBlank, @sex, @sexBlank, @relation, @relationBlank, @dataStatus, @dataStatusBlank, @claimability, @claimabilityBlank, @polStatus, @polStatusBlank, @caseStatus, @caseStatusBlank, @expired, @expiredBlank, @premMode, @premModeBlank, @takeover, @takeoverBlank, @state, @stateBlank, @renewal, @renewalBlank, @agent, @agentBlank, @policy, @policyBlank, @patient, @polYearFrom, @polYearTo, @polYearBlank, @mcDaysFrom, @mcDaysTo, @mcDaysBlank, @payor, @payorBlank, @company, @companyBlank, @clinicName, @clinicNameBlank, @clinicGroup, @clinicGroupBlank, @branch, @branchBlank, @dovFrom, @dovTo, @dovBlank, @dorFrom, @dorTo, @dorBlank, @dowFrom, @dowTo, @dowBlank, @treatmentFrom, @treatmentTo, @treatmentBlank, @planFrom, @planTo, @planBlank, @effFrom, @effTo, @effBlank, @expFrom, @expTo, @expBlank, @joinedFrom, @joinedTo, @joinedBlank, @dobFrom, @dobTo, @dobBlank, @mbmNoFrom, @mbmNoTo, @caseNoFrom, @caseNoTo, @bordxFrom, @bordxTo, @bordxBlank, @bordxDateFrom, @bordxDateTo, @bordxDateBlank, @icFrom, @icTo, @icBlank, @mbmIcFrom, @mbmIcTo, @mbmIcBlank, @actualFrom, @actualTo, @approvedFrom, @approvedTo, @invoiceFrom, @invoiceTo, @invoiceBlank, @submissionFrom, @submissionTo, @submissionBlank;
END
GO
