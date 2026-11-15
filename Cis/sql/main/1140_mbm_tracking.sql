-- Tracking > MBM Tracking (FrmNewMBMTracking and its variants): the 178 cases of HISDB CISMBMTracking (generated from the procedure by
-- gen-mbmtrk.js). Each case's query text is kept as the desktop's; the pasted criteria are replaced by a WHERE fragment built only from the filters that are set,
-- with every value passed as a parameter.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.MbmTrk_Run @case int,
    @hlt nvarchar(1) = NULL,
    @hltBlank bit = 0,
    @payorBlank bit = 0,
    @ins nvarchar(1) = NULL,
    @insBlank bit = 0,
    @prinType nvarchar(10) = NULL,
    @cancelFrom datetime = NULL,
    @cancelTo datetime = NULL,
    @cancelBlank bit = 0,
    @dayInsuredFrom nvarchar(20) = NULL,
    @dayInsuredTo nvarchar(20) = NULL,
    @dayInsuredBlank bit = 0,
    @ageDaysFrom nvarchar(20) = NULL,
    @ageDaysTo nvarchar(20) = NULL,
    @ageDaysBlank bit = 0,
    @delayBordxFrom nvarchar(20) = NULL,
    @delayBordxTo nvarchar(20) = NULL,
    @delayBordxBlank bit = 0,
    @delayInvBordxFrom nvarchar(20) = NULL,
    @delayInvBordxTo nvarchar(20) = NULL,
    @delayInvBordxBlank bit = 0,
    @delayRegFrom nvarchar(20) = NULL,
    @delayRegTo nvarchar(20) = NULL,
    @delayRegBlank bit = 0,
    @delayInvDFrom nvarchar(20) = NULL,
    @delayInvDTo nvarchar(20) = NULL,
    @delayInvDBlank bit = 0,
    @delayInvEFrom nvarchar(20) = NULL,
    @delayInvETo nvarchar(20) = NULL,
    @delayInvEBlank bit = 0,
    @premiumFrom money = NULL,
    @premiumTo money = NULL,
    @premiumBlank bit = 0,
    @mcoFrom money = NULL,
    @mcoTo money = NULL,
    @mcoBlank bit = 0,
    @employeeFrom nvarchar(50) = NULL,
    @employeeTo nvarchar(50) = NULL,
    @employeeBlank bit = 0,
    @memberTypeBlank bit = 0,
    @adultChild nvarchar(1) = NULL,
    @adultChildBlank bit = 0,
    @race nvarchar(1) = NULL,
    @raceBlank bit = 0,
    @sex nvarchar(1) = NULL,
    @sexBlank bit = 0,
    @rel nvarchar(1) = NULL,
    @relBlank bit = 0,
    @dataStatus nvarchar(10) = NULL,
    @dataStatusBlank bit = 0,
    @status nvarchar(1) = NULL,
    @statusBlank bit = 0,
    @expired nvarchar(1) = NULL,
    @expiredBlank bit = 0,
    @renewal nvarchar(1) = NULL,
    @renewalBlank bit = 0,
    @premMode nvarchar(1) = NULL,
    @premModeBlank bit = 0,
    @state nvarchar(100) = NULL,
    @stateBlank bit = 0,
    @takeover nvarchar(1) = NULL,
    @takeoverBlank bit = 0,
    @agent nvarchar(30) = NULL,
    @agentBlank bit = 0,
    @mbmNo nvarchar(30) = NULL,
    @mbmNoBlank bit = 0,
    @policy nvarchar(60) = NULL,
    @policyBlank bit = 0,
    @payor nvarchar(10) = NULL,
    @company nvarchar(200) = NULL,
    @companyBlank bit = 0,
    @planFrom nvarchar(20) = NULL,
    @planTo nvarchar(20) = NULL,
    @planBlank bit = 0,
    @ageFrom nvarchar(20) = NULL,
    @ageTo nvarchar(20) = NULL,
    @ageBlank bit = 0,
    @effFrom datetime = NULL,
    @effTo datetime = NULL,
    @effBlank bit = 0,
    @expFrom datetime = NULL,
    @expTo datetime = NULL,
    @expBlank bit = 0,
    @joinedFrom datetime = NULL,
    @joinedTo datetime = NULL,
    @joinedBlank bit = 0,
    @dorFrom datetime = NULL,
    @dorTo datetime = NULL,
    @dorBlank bit = 0,
    @dobFrom datetime = NULL,
    @dobTo datetime = NULL,
    @dobBlank bit = 0,
    @bordxFrom datetime = NULL,
    @bordxTo datetime = NULL,
    @bordxBlank bit = 0,
    @batchFrom nvarchar(50) = NULL,
    @batchTo nvarchar(50) = NULL,
    @batchBlank bit = 0,
    @postcodeFrom nvarchar(50) = NULL,
    @postcodeTo nvarchar(50) = NULL,
    @postcodeBlank bit = 0,
    @polYearFrom nvarchar(50) = NULL,
    @polYearTo nvarchar(50) = NULL,
    @polYearBlank bit = 0,
    @invoiceFrom nvarchar(50) = NULL,
    @invoiceTo nvarchar(50) = NULL,
    @invoiceBlank bit = 0,
    @invoiceDateFrom datetime = NULL,
    @invoiceDateTo datetime = NULL,
    @invoiceDateBlank bit = 0,
    @icFrom nvarchar(50) = NULL,
    @icTo nvarchar(50) = NULL,
    @icBlank bit = 0,
    @label nvarchar(1) = NULL,
    @labelBlank bit = 0,
    @department nvarchar(200) = NULL,
    @departmentBlank bit = 0,
    @cardFrom datetime = NULL,
    @cardTo datetime = NULL,
    @cardBlank bit = 0,
    @endtFrom datetime = NULL,
    @endtTo datetime = NULL,
    @endtBlank bit = 0,
    @suppBordxFrom datetime = NULL,
    @suppBordxTo datetime = NULL,
    @detailPrinType nvarchar(1) = NULL,
    @suppRegFrom datetime = NULL,
    @suppRegTo datetime = NULL,
    @suppRegBlank bit = 0
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @b nvarchar(max), @a nvarchar(max) = N'';
    IF @case = 1 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee, 
                          			                  SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0 ', @a = N'GROUP BY PAYCode, HLTCode, INSCode, AGECode';
    IF @case = 2 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalS, SUM(MBMMCOFee) AS MCOGross,  
                                                 			SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0 ', @a = N'GROUP BY PAYCode, HLTCode, INSCode, AGECode';
    IF @case = 3 SELECT @b = N'SELECT PAYCode, HLTCode, AGECode, COUNT(MBMNumber) AS TotalP,SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee, 
			          			 SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, AGECode';
    IF @case = 4 SELECT @b = N'SELECT PAYCode, HLTCode, PrinAGECode, COUNT(MBMNumber) AS TotalS,SUM(MBMMCOFee) AS MCOGross, 
			          			 SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, PrinAGECode';
    IF @case = 5 SELECT @b = N'SELECT PAYCode, HLTCode, AGECode, COUNT(MBMNumber) AS TotalP,SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee, 
          			 			SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, AGECode';
    IF @case = 6 SELECT @b = N'SELECT PAYCode, HLTCode, AGECode, COUNT(MBMNumber) AS TotalS, SUM(MBMMCOFee) AS MCOGross, 
						SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, AGECode';
    IF @case = 7 SELECT @b = N'SELECT HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, AGECode';
    IF @case = 8 SELECT @b = N'SELECT HLTCode, INSCode, PrinAGECode, COUNT(MBMNumber) AS TotalS, SUM(MBMMCOFee) AS MCOGross, 
						SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, PrinAGECode';
    IF @case = 9 SELECT @b = N'SELECT HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee, 
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, AGECode';
    IF @case = 10 SELECT @b = N'SELECT HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalS, SUM(MBMMCOFee) AS MCOGross, 
						SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, AGECode';
    IF @case = 11 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, AGECode';
    IF @case = 12 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, PrinAGECode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, PrinAGECode';
    IF @case = 13 SELECT @b = N'SELECT PAYCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, INSCode, AGECode';
    IF @case = 14 SELECT @b = N'SELECT PAYCode, INSCode, PrinAGECode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, INSCode, PrinAGECode';
    IF @case = 15 SELECT @b = N'SELECT PAYCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, INSCode, AGECode';
    IF @case = 16 SELECT @b = N'SELECT PAYCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, INSCode, AGECode';
    IF @case = 17 SELECT @b = N'SELECT PAYCode, MBMAgentCode, HLTCode, INSCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode, HLTCode, INSCode';
    IF @case = 18 SELECT @b = N'SELECT PAYCode, MBMAgentCode, HLTCode, INSCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode, HLTCode, INSCode';
    IF @case = 19 SELECT @b = N'SELECT MBMAgentCode, PAYCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY MBMAgentCode, PAYCode';
    IF @case = 20 SELECT @b = N'SELECT MBMAgentCode, PAYCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY MBMAgentCode, PAYCode';
    IF @case = 21 SELECT @b = N'SELECT PAYCode, MBMAgentCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode, INSCode, PLNCode';
    IF @case = 22 SELECT @b = N'SELECT PAYCode, MBMAgentCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode, INSCode, PLNCode';
    IF @case = 23 SELECT @b = N'SELECT PAYCode, MBMAgentCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode, INSCode, PLNCode';
    IF @case = 24 SELECT @b = N'SELECT PAYCode, MBMAgentCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode, INSCode, PLNCode';
    IF @case = 25 SELECT @b = N'SELECT PAYCode, MBMAgentCode, RACCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross,
						SUM(MBMAccessFee) AS EAFee From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode, RACCode';
    IF @case = 26 SELECT @b = N'SELECT PAYCode, MBMAgentCode, RACCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode, RACCode';
    IF @case = 27 SELECT @b = N'SELECT PAYCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode';
    IF @case = 28 SELECT @b = N'SELECT PAYCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode';
    IF @case = 29 SELECT @b = N'SELECT TOP 20 PAYCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode';
    IF @case = 30 SELECT @b = N'SELECT TOP 20 PAYCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMAgentCode ORDER BY PremiumGross Desc';
    IF @case = 31 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, MBMMemberType, PLNCode, MBMCCoverID, HLTCode, MBMAgentCode, 
						MBMStatus, MBMAccessFee, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMPremium, DayInsured, PAYCode
						From VIEW_CISMBM_Search Where 0 = 0', @a = N'';
    IF @case = 32 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, MBMMemberType, PLNCode, MBMCCoverID, HLTCode, MBMAgentCode, 
						MBMStatus, MBMAccessFee, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMPremium, DayInsured, PAYCode,
						RElCode, MBMCCoverID, MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 33 SELECT @b = N'SELECT MBMValueDate, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalP,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY MBMValueDate, MBMBordxDate, MBMBatchNo';
    IF @case = 34 SELECT @b = N'SELECT MBMValueDate, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalS,
						SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY MBMValueDate, MBMBordxDate, MBMBatchNo';
    IF @case = 35 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMBordxDate, MBMBatchNo, MBMPremium,
						MBMMemberType, MBMStatus, MBMCCoverID, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMAccessFee
						From VIEW_CISMBM_Search Where 0 = 0', @a = N'';
    IF @case = 36 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMBordxDate, MBMBatchNo, MBMPremium, MBMCCoverID,
						MBMMemberType, MBMStatus, MBMCCoverID, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMAccessFee, MBMCEffDate,MBMCExpDate
						From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 37 SELECT @b = N'SELECT PAYCode, MBMBordxDate From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 38 SELECT @b = N'SELECT PAYCode, MBMBordxDate From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 39 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalIC,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = ''I'' AND MBMAdultChild = ''C''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 40 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalIA,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = ''I'' AND MBMAdultChild = ''A''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 41 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalF,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = ''F''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 42 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalGC,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = ''G'' AND MBMAdultChild = ''C''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 43 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalGA,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = ''G'' AND MBMAdultChild = ''A''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 44 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalH,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = ''H''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 45 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalFS,
						SUM(MBMMCOFee) AS MCOFee From View_CISMBM_Supplementary Where INSCode = ''F''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 46 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalHS,
						SUM(MBMMCOFee) AS MCOFee From View_CISMBM_Supplementary Where INSCode = ''H''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 47 SELECT @b = N'SELECT PAYCode, MBMBordxDate From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 48 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalIC,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = ''I'' AND MBMAdultChild = ''C''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 49 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalIA,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = ''I'' AND MBMAdultChild = ''A''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 50 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalF,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = ''F''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 51 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalFS,
						       SUM(MBMPremium) AS Premium From View_CISMBM_Supplementary Where INSCode = ''F''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 52 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalGC,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = ''G'' AND MBMAdultChild = ''C''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 53 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalGA,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = ''G'' AND MBMAdultChild = ''A''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 54 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalH,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = ''H''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 55 SELECT @b = N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalHS,
						       SUM(MBMPremium) AS Premium From View_CISMBM_Supplementary Where INSCode = ''H''', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 56 SELECT @b = N'SELECT PAYCode, MBMBordxDate From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMBordxDate';
    IF @case = 57 SELECT @b = N'SELECT PAYCode, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMBordxDate, MBMBatchNo';
    IF @case = 58 SELECT @b = N'SELECT PAYCode, MBMBordxDate, MBMBatchNo, COUNT(MBMCNumber) AS TotalS,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMBordxDate, MBMBatchNo';
    IF @case = 59 SELECT @b = N'SELECT HLTCode, BRNCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, BRNCode';
    IF @case = 60 SELECT @b = N'SELECT HLTCode, BRNCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, BRNCode';
    IF @case = 61 SELECT @b = N'SELECT BRNCode, MBMAgentCode From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 62 SELECT @b = N'SELECT BRNCode, MBMAgentCode From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 63 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0
						       AND (BRNCode IS NULL OR BRNCode = '') AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 64 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0
						       AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 65 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0
						       AND (BRNCode = '' OR BRNCode IS NULL)', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 66 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 67 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0
						       AND (BRNCode IS NULL OR BRNCode = '') AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 68 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0
						       AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 69 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0
						       AND (BRNCode IS NULL OR BRNCode = '')', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 70 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 71 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(CASNumber) AS TotalClaim,
						       SUM(CLMTotalApproved) AS AppAmt From VIEW_MBMClaim_Search Where 0 = 0
						       AND (BRNCode IS NULL OR BRNCode = '') AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 72 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(CASNumber) AS TotalClaim,
						       SUM(CLMTotalApproved) AS AppAmt From VIEW_MBMClaim_Search Where 0 = 0
						       AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 73 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(CASNumber) AS TotalClaim,
						       SUM(CLMTotalApproved) AS AppAmt From VIEW_MBMClaim_Search Where 0 = 0
						       AND (BRNCode IS NULL OR BRNCode = '')', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 74 SELECT @b = N'SELECT BRNCode, MBMAgentCode, COUNT(CASNumber) AS TotalClaim,
						       SUM(CLMTotalApproved) AS AppAmt From VIEW_MBMClaim_Search Where 0 = 0', @a = N'GROUP BY BRNCode, MBMAgentCode';
    IF @case = 75 SELECT @b = N'SELECT HLTCode, PAYCode, BRNCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode, BRNCode';
    IF @case = 76 SELECT @b = N'SELECT HLTCode, PAYCode, BRNCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode, BRNCode';
    IF @case = 77 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMDOB, MBMBordxDate,  MBMValueDate, MBMInvoiceDate, 
						      MBMMemberType, MBMStatus, MBMPayorEffDate, MBMPayorExpDate, MBMRenewFlag, PAYCode, DayInsured,
						      AgeDays, DelayBordx, DelayReg, DelayInvD, DelayInvE, DelayInvBordx, RELCode, MBMCCoverID From VIEW_CISMBM_Search Where 0 = 0', @a = N'';
    IF @case = 78 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMDOB, MBMBordxDate,  MBMValueDate, MBMInvoiceDate, 
						      MBMMemberType, MBMStatus, MBMPayorEffDate, MBMPayorExpDate, MBMRenewFlag, PAYCode, DayInsured,	
						      AgeDays, DelayBordx, DelayReg, DelayInvD, DelayInvE, DelayInvBordx, RELCode, MBMCCoverID, MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 79 SELECT @b = N'SELECT MBMNumber, MBMPolicyNo, MBMName, MBMPayorEffDate, MBMPayorExpDate, DayInsured,
						      INSCode, PLNCode From VIEW_CISMBM_Search Where 0 = 0', @a = N'';
    IF @case = 80 SELECT @b = N'SELECT COUNT(MBMNumber) AS TotalP, DayInsured From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY DayInsured';
    IF @case = 81 SELECT @b = N'SELECT DayInsured, COUNT(MBMNumber) AS TotalS From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY DayInsured';
    IF @case = 82 SELECT @b = N'SELECT HLTCode, INSCode, MBMSex, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, MBMSex';
    IF @case = 83 SELECT @b = N'SELECT HLTCode, INSCode, PrinSex, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						       From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, PrinSex';
    IF @case = 84 SELECT @b = N'SELECT HLTCode, INSCode, MBMSex, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						       From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, MBMSex';
    IF @case = 85 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, MBMSex, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, MBMSex';
    IF @case = 86 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, PrinSex, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						       From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, PrinSex';
    IF @case = 87 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, MBMSex, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						       From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, MBMSex';
    IF @case = 88 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, AGECode, MBMSex, MBMExpiredStatus,
						     PLNCode, MBMGroupCompany, MBMMemberType, case When cast(MBMPayorExpDate as date)<cast(getdate() as date) then ''E'' else MBMStatus end as MBMStatus, MBMDataStatus, RElCode, MBMCCoverID,
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee, MBMValueDate,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMIcBcPp, 
						     MBMEmployeeNo, MBMDepartment, MBMCancelDate,Subsidiary,GRPCoPaycode,mbmremarks,REMARK,MBMCostCenter,MBMDepartment,MBMBranch,MBMDivision,Principalname 
							 ,MBMInvoiceDate,MBMInvoiceNo
							 From VIEW_CISMBM_Search Where 0 = 0', @a = N'';
    IF @case = 89 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, AGECode, MBMSex, MBMExpiredStatus,
						     PLNCode, MBMGroupCompany, MBMMemberType, case When cast(MBMPayorExpDate as date)<cast(getdate() as date) then ''E'' else MBMStatus end as MBMStatus, MBMDataStatus, 
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee, MBMValueDate,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMIcBcPp, 
						     MBMEmployeeNo, MBMDepartment, RELCode, MBMCCoverID, MBMCancelDate,GRPCoPaycode, MBMCEffDate,MBMCExpDate,Subsidiary,mbmremarks,REMARK,MBMCostCenter,MBMDepartment,
							 MBMBranch,MBMDivision,Principalname ,MBMInvoiceDate,MBMInvoiceNo
							 From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 90 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, AGECode, MBMSex, MBMExpiredStatus,
						     PLNCode, MBMGroupCompany, MBMMemberType, MBMStatus, MBMDataStatus,
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee, MBMValueDate,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMIcBcPp, 
						     MBMEmployeeNo, MBMDepartment, RELCode, MBMCCoverID, MBMCRegDate, MBMCancelDate, MBMBordxDate,MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 91 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, AGECode, MBMSex, MBMEmployeeNo,
						     PLNCode, MBMGroupCompany, MBMMemberType, MBMStatus, RELCode, MBMCCoverID,
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee, MBMCancelDate,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMDepartment
						     From VIEW_CISMBM_Search Where 0 = 0', @a = N'';
    IF @case = 92 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName,  MBMCName, INSCode, AGECode, MBMSex, MBMEmployeeNo,
						     PLNCode, MBMGroupCompany, MBMMemberType, MBMStatus, 
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMDepartment,
						     MBMCSex, MBMCAgeband, MBMCDOB, RELCode, MBMCCoverID, MBMCancelDate, MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary2 Where 0 = 0', @a = N'';
    IF @case = 93 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMCName, INSCode, AGECode, MBMSex, MBMEmployeeNo,
						     PLNCode, MBMGroupCompany, MBMMemberType, MBMStatus, RELCode, MBMCCoverID,
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMDepartment,
						     MBMCSex, MBMCAgeband, MBMCDOB, MBMCRegDate, MBMCancelDate, MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary2 Where 0 = 0', @a = N'';
    IF @case = 94 SELECT @b = N'Select MBMNumber, MBMName, RElCode, MBMCCoverID, MBMStatus,
						     MBMMemberType, MBMPremium, MBMBasicMCO, MBMMCOFee, MBMBasicPrem,
						     MBMPayorMCONet, MBMPayorMCOGross, MBMPayorPremNet, MBMPayorPremGross
						     From VIEW_CISMBM_Search Where 0 = 0', @a = N'';
    IF @case = 95 SELECT @b = N'Select MBMNumber, MBMName, RElCode, MBMCCoverID, MBMStatus,
						     MBMMemberType, MBMPremium, MBMBasicMCO, MBMMCOFee, MBMBasicPrem,
						     MBMPayorMCONet, MBMPayorMCOGross, MBMPayorPremNet, MBMPayorPremGross
						     From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 96 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMAddress1, CTYCode,
						     MBMGroupCompany, MBMMemberType, MBMStatus, MBMCCoverID, MBMPostCode,
						     STACode, MBMRenewFlag, PAYCode, MBMAddress2, MBMAddress3
						     From VIEW_CISMBM_Search Where 0 = 0', @a = N'';
    IF @case = 97 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMAddress1, CTYCode,
						     MBMGroupCompany, MBMMemberType, MBMStatus, MBMCCoverID, MBMPostCode,
						     STACode, MBMRenewFlag, PAYCode, MBMAddress2, MBMAddress3, MBMCCoverID, MBMCRegDate
						     From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 98 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMAddress1, CTYCode,
						     MBMGroupCompany, MBMMemberType, MBMStatus, MBMCCoverID, MBMPostCode,
						     STACode, MBMRenewFlag, PAYCode, MBMAddress2, MBMAddress3, MBMCCoverID
						     From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 99 SELECT @b = N'SELECT PAYCode, MBMExpiredStatus, INSCode, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMExpiredStatus, INSCode';
    IF @case = 100 SELECT @b = N'SELECT PAYCode, MBMExpiredStatus, INSCode, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMExpiredStatus, INSCode';
    IF @case = 101 SELECT @b = N'Select MBMNumber, MBMName, MBMIcBcPp, MBMInstallment,
						     MBMMemberType, MBMStatus, MBMCCoverID, MBMTakeover,
						     MBMPayorEffDate, MBMPayorExpDate, MBMPolicyYear,
						     MBMRenewFlag, MBMFirstJoinedDate From VIEW_CISMBM_Search Where 0 = 0', @a = N'';
    IF @case = 102 SELECT @b = N'Select MBMNumber, MBMName, MBMIcBcPp, MBMInstallment,
						     MBMMemberType, MBMStatus, MBMCCoverID, MBMTakeover,
						     MBMPayorEffDate, MBMPayorExpDate, MBMPolicyYear,
						     MBMRenewFlag, MBMFirstJoinedDate, MBMCCoverID, MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 103 SELECT @b = N'SELECT MBMPolicyYear, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY MBMPolicyYear';
    IF @case = 104 SELECT @b = N'SELECT MBMPolicyYear, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY MBMPolicyYear';
    IF @case = 105 SELECT @b = N'SELECT PAYCode, MBMInstallment, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMInstallment';
    IF @case = 106 SELECT @b = N'SELECT PAYCode, MBMInstallment, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMInstallment';
    IF @case = 107 SELECT @b = N'SELECT PAYCode, MBMRenewFlag, INSCode, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMRenewFlag, INSCode';
    IF @case = 108 SELECT @b = N'SELECT PAYCode, MBMRenewFlag, INSCode, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMRenewFlag, INSCode';
    IF @case = 109 SELECT @b = N'SELECT PAYCode, MBMTakeover, INSCode, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMTakeover, INSCode';
    IF @case = 110 SELECT @b = N'SELECT PAYCode, MBMTakeover, INSCode, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMTakeover, INSCode';
    IF @case = 111 SELECT @b = N'SELECT HLTCode, INSCode, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode';
    IF @case = 112 SELECT @b = N'SELECT HLTCode, INSCode, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode';
    IF @case = 113 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode';
    IF @case = 114 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode';
    IF @case = 115 SELECT @b = N'SELECT MONTH(MBMPayorEffDate) AS Month,
						     YEAR(MBMPayorEffDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)';
    IF @case = 116 SELECT @b = N'SELECT MONTH(MBMPayorEffDate) AS Month,
						     YEAR(MBMPayorEffDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)';
    IF @case = 117 SELECT @b = N'SELECT PAYCode, MONTH(MBMPayorEffDate) AS Month,
						     YEAR(MBMPayorEffDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)';
    IF @case = 118 SELECT @b = N'SELECT PAYCode, MONTH(MBMPayorEffDate) AS Month,
						     YEAR(MBMPayorEffDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)';
    IF @case = 119 SELECT @b = N'SELECT MONTH(MBMPayorExpDate) AS Month,
						     YEAR(MBMPayorExpDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)';
    IF @case = 120 SELECT @b = N'SELECT MONTH(MBMPayorExpDate) AS Month,
						     YEAR(MBMPayorExpDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)';
    IF @case = 121 SELECT @b = N'SELECT PAYCode, MONTH(MBMPayorExpDate) AS Month,
						     YEAR(MBMPayorExpDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)';
    IF @case = 122 SELECT @b = N'SELECT PAYCode, MONTH(MBMPayorExpDate) AS Month,
						     YEAR(MBMPayorExpDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)';
    IF @case = 123 SELECT @b = N'SELECT MONTH(MBMValueDate) AS Month,
						     YEAR(MBMValueDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY MONTH(MBMValueDate), YEAR(MBMValueDate)';
    IF @case = 124 SELECT @b = N'SELECT MONTH(MBMValueDate) AS Month,
						     YEAR(MBMValueDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY MONTH(MBMValueDate), YEAR(MBMValueDate)';
    IF @case = 125 SELECT @b = N'SELECT PAYCode, MONTH(MBMValueDate) AS Month,
						     YEAR(MBMValueDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MONTH(MBMValueDate), YEAR(MBMValueDate)';
    IF @case = 126 SELECT @b = N'SELECT PAYCode, MONTH(MBMValueDate) AS Month,
						     YEAR(MBMValueDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MONTH(MBMValueDate), YEAR(MBMValueDate)';
    IF @case = 127 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMMemberType, MBMBordxDate, MBMBatchNo,
						     MBMCCoverID, MBMStatus, MBMInvoiceNo, MBMInvoiceDate, MBMMCOFee, MBMValueDate
						     From VIEW_CISMBM_Search Where 0 = 0', @a = N'';
    IF @case = 128 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMMemberType, MBMBordxDate, MBMBatchNo,
						     MBMCCoverID, MBMStatus, MBMInvoiceNo, MBMInvoiceDate, MBMMCOFee, MBMValueDate,
						     RELCode From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 129 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMMemberType, MBMBordxDate, MBMBatchNo,
						     MBMCCoverID, MBMStatus, MBMInvoiceNo, MBMInvoiceDate, MBMMCOFee, MBMValueDate,
						     RELCode From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 130 SELECT @b = N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate';
    IF @case = 131 SELECT @b = N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate';
    IF @case = 132 SELECT @b = N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate, COUNT(MBMNumber) AS TotalMBM,
						        SUM(MBMMCOFEE) AS MCOFEE,MBMGroupCompany,Subsidiary From VIEW_MBM_Search_MCO_V1 Where 0 = 0', @a = N'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate,MBMGroupCompany,Subsidiary';
    IF @case = 133 SELECT @b = N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate, COUNT(MBMNumber) AS TotalMBMS,
						        SUM(MBMMCOFEE) AS MCOFEE From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate';
    IF @case = 134 SELECT @b = N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate, COUNT(MBMNumber) AS TotalMBM,
						        SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate';
    IF @case = 135 SELECT @b = N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate, COUNT(MBMNumber) AS TotalMBMS,
						        SUM(MBMPremium) AS Premium From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate';
    IF @case = 136 SELECT @b = N'SELECT HLTCode, PAYCode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode';
    IF @case = 137 SELECT @b = N'SELECT HLTCode, PAYCode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode';
    IF @case = 138 SELECT @b = N'SELECT PAYCode, SUM(MBMMCOFee) AS MBMMCOFee, SUM(MBMBasicMCO) AS MBMBasicMCO,
						        SUM(MBMPayorMCOGross) AS MBMPayorMCOGross, SUM(MBMPayorMCONet) AS MBMPayorMCONet,
						        SUM(MBMPremium) AS MBMPremium, SUM(MBMBasicPrem) AS MBMBasicPrem,
						        SUM(MBMPayorPremGross) AS MBMPayorPremGross, SUM(MBMPayorPremNet) AS MBMPayorPremNet
						        From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode';
    IF @case = 139 SELECT @b = N'SELECT PAYCode, SUM(MBMMCOFee) AS MBMMCOFee, SUM(MBMBasicMCO) AS MBMBasicMCO,
						        SUM(MBMPayorMCOGross) AS MBMPayorMCOGross, SUM(MBMPayorMCONet) AS MBMPayorMCONet,
						        SUM(MBMPremium) AS MBMPremium, SUM(MBMBasicPrem) AS MBMBasicPrem,
						        SUM(MBMPayorPremGross) AS MBMPayorPremGross, SUM(MBMPayorPremNet) AS MBMPayorPremNet
						        From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode';
    IF @case = 140 SELECT @b = N'SELECT PAYCompanyName, MBMGroupCompany, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMPremium) AS PremiumGross, SUM(CLMTotalApproved) AS CLMTotalApproved
						        From VIEW_MBM_Search4 Where (INSCode = ''G'' OR INSCode = ''H'')', @a = N'GROUP BY PAYCompanyName, MBMGroupCompany';
    IF @case = 141 SELECT @b = N'SELECT PAYCompanyName, MBMGroupCompany, COUNT(MBMNumber) AS TotalS
						        From View_MBM_Supplementary4 Where (INSCode = ''G'' OR INSCode = ''H'')', @a = N'GROUP BY PAYCompanyName, MBMGroupCompany';
    IF @case = 142 SELECT @b = N'SELECT HLTCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, PLNCode';
    IF @case = 143 SELECT @b = N'SELECT HLTCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, PLNCode';
    IF @case = 144 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, PLNCode';
    IF @case = 145 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, PLNCode';
    IF @case = 146 SELECT @b = N'SELECT HLTCode, INSCode, PLNCode, AGECode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, PLNCode, AGECode';
    IF @case = 147 SELECT @b = N'SELECT HLTCode, INSCode, PLNCode, PrinAGECode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, PLNCode, PrinAGECode';
    IF @case = 148 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, PLNCode, AGECode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, PLNCode, AGECode';
    IF @case = 149 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, PLNCode, PrinAGECode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, PLNCode, PrinAGECode';
    IF @case = 150 SELECT @b = N'SELECT HLTCode, INSCode, PLNCode, AGECode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, PLNCode, AGECode';
    IF @case = 151 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, PLNCode, AGECode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, PLNCode, AGECode';
    IF @case = 152 SELECT @b = N'SELECT PAYCode, MBMPolicyNo, MBMGroupCompany, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, MBMPolicyNo, MBMGroupCompany';
    IF @case = 153 SELECT @b = N'SELECT PAYCode, MBMPolicyNo, MBMGroupCompany, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, MBMPolicyNo, MBMGroupCompany';
    IF @case = 154 SELECT @b = N'SELECT HLTCode, INSCode, RACCode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, RACCode';
    IF @case = 155 SELECT @b = N'SELECT HLTCode, INSCode, RACCode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, INSCode, RACCode';
    IF @case = 156 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, RACCode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, RACCode';
    IF @case = 157 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, RACCode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, RACCode';
    IF @case = 158 SELECT @b = N'SELECT PAYCode, HLTCode, RELCode, AGECode, COUNT(DISTINCT MBMNumber) AS TotalP,
						        COUNT(MBMNumber) AS TotalS From VIEW_MBM_Search3 Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, RELCode, AGECode';
    IF @case = 159 SELECT @b = N'SELECT PAYCode, HLTCode, RELCode, AGECode, COUNT(DISTINCT MBMNumber) AS TotalP,
						        COUNT(MBMNumber) AS TotalS From View_CISMBM_Supplementary Where RELCode <> ''P''', @a = N'GROUP BY PAYCode, HLTCode, RELCode, AGECode';
    IF @case = 160 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, RELCode, COUNT(DISTINCT MBMNumber) AS TotalP,
						        COUNT(MBMNumber) AS TotalS From VIEW_MBM_Search3 Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, INSCode, RELCode';
    IF @case = 161 SELECT @b = N'SELECT PAYCode, HLTCode, INSCode, RELCode, COUNT(DISTINCT MBMNumber) AS TotalP,
						        COUNT(MBMNumber) AS TotalS From View_CISMBM_Supplementary Where RELCode <> ''P''', @a = N'GROUP BY PAYCode, HLTCode, INSCode, RELCode';
    IF @case = 162 SELECT @b = N'SELECT HLTCode, STACode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, STACode';
    IF @case = 163 SELECT @b = N'SELECT HLTCode, STACode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, STACode';
    IF @case = 164 SELECT @b = N'SELECT PAYCode, HLTCode, STACode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, STACode';
    IF @case = 165 SELECT @b = N'SELECT PAYCode, HLTCode, STACode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY PAYCode, HLTCode, STACode';
    IF @case = 166 SELECT @b = N'SELECT HLTCode, PAYCode, MBMGroupCompany, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode, MBMGroupCompany';
    IF @case = 167 SELECT @b = N'SELECT HLTCode, PAYCode, MBMGroupCompany, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode, MBMGroupCompany';
    IF @case = 168 SELECT @b = N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMDepartment';
    IF @case = 169 SELECT @b = N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMDepartment';
    IF @case = 170 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMCCoverID, MBMBordxDate, MBMInvoiceDate,
						     MBMCLNCardSentDate, MBMPayorEffDate, BRNCode, MBMAgentCode, MBMStatus, MBMMemberType
						     From VIEW_CISMBM_Search Where 0 = 0', @a = N'';
    IF @case = 171 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMCCoverID, MBMBordxDate, MBMInvoiceDate,
						     MBMCLNCardSentDate, MBMPayorEffDate, BRNCode, MBMAgentCode, MBMStatus, MBMMemberType
						     From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 172 SELECT @b = N'Select MBMNumber, MBMPolicyNo, MBMName, MBMCCoverID, MBMBordxDate, MBMInvoiceDate,
						     MBMCLNCardSentDate, MBMPayorEffDate, BRNCode, MBMAgentCode, MBMStatus, MBMMemberType
						     From View_CISMBM_Supplementary Where 0 = 0', @a = N'';
    IF @case = 173 SELECT @b = N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMBordxDate, MBMBatchNo';
    IF @case = 174 SELECT @b = N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMBordxDate, MBMBatchNo';
    IF @case = 175 SELECT @b = N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, MBMBordxDate, MBMBatchNo';
    IF @case = 176 SELECT @b = N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0', @a = N'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, MBMBordxDate, MBMBatchNo';
    IF @case = 177 SELECT @b = N'Select * from HISDB..View_CISInvoice_Summary Where 0 = 0', @a = N'';
    IF @case = 178 SELECT @b = N' SELECT case when MBMCCoverID=''00'' then ''Principal'' else ''Dependent'' end as cc,
										PAYCode, MBMInvoiceNo, MBMInvoiceDate, COUNT(MBMNumber) AS TotalMBM,SUM(MBMMCOFEE) AS MCOFEE 
										,MBMGroupCompany,coalesce(Subsidiary,'''') as Subsidiary,OPAccCode
										into #temp 
										From VIEW_MBM_Search_MCO_V1 Where 0 = 0  ', @a = N' GROUP BY case when MBMCCoverID=''00'' then ''Principal'' else ''Dependent'' end,PAYCode, MBMInvoiceNo, MBMInvoiceDate,MBMGroupCompany,Subsidiary,OPAccCode 
										Select [PAYCode],Sum(coalesce(Principal_amt,0)) as Principal_amt,
										Sum(coalesce(Dependent_amt,0)) as Dependent_amt, 
										Sum(coalesce(Principal_mbmcnt,0)) as Principal_mbmcnt,
										Sum(coalesce(Dependent_mbmcnt,0)) as Dependent_mbmcnt,
										MBMInvoiceNo, MBMInvoiceDate,MBMGroupCompany,Subsidiary,OPAccCode
										from (
										Select 
										[PAYCode],MBMInvoiceNo, MBMInvoiceDate
										,Sum(MCOFEE) as MCOFEE
										,Sum(TotalMBM) as TotalMBM 
										,[MCOFEE_amt]=[CC]+''_amt''
										,[TotalMBM_cnt]=[CC]+''_mbmcnt''
										,MBMGroupCompany,Subsidiary,OPAccCode
										from #temp 
										group by PAYCode,[CC],MBMInvoiceNo, MBMInvoiceDate,MBMGroupCompany,Subsidiary,OPAccCode
										)NN 
										PIVOT (Sum(MCOFEE) for [MCOFEE_amt] IN (Principal_amt,Dependent_amt)) As PV1
										PIVOT (Sum(TotalMBM) for [TotalMBM_cnt] IN (Principal_mbmcnt,Dependent_mbmcnt)) As PV2
										group by [PAYCode],MBMInvoiceNo, MBMInvoiceDate,MBMGroupCompany,Subsidiary,OPAccCode	
										order by MBMInvoiceNo asc										
										Drop table #temp   
									';
    IF @b IS NULL RETURN;
    DECLARE @w nvarchar(max) = N'';
    IF @hlt IS NOT NULL SET @w += N' AND HLTCode = @hlt';
    IF @hltBlank = 1 SET @w += N' AND (HLTCode IS NULL OR HLTCode = '''')';
    IF @payorBlank = 1 SET @w += N' AND (PAYCode IS NULL OR PAYCode = '''')';
    IF @ins IS NOT NULL SET @w += N' AND INSCode = @ins';
    IF @insBlank = 1 SET @w += N' AND (INSCode IS NULL OR INSCode = '''')';
    IF @prinType IS NOT NULL SET @w += CASE @prinType WHEN 'P|Both' THEN N' AND (MBMMemberType = ''P'' or  MBMMemberType = ''Q'') AND MBMCCoverID = ''00''' WHEN 'P|C' THEN N' AND MBMMemberType = ''P'' AND MBMCCoverID = ''00''' WHEN 'P|S' THEN N' AND MBMMemberType = ''Q'' AND MBMCCoverID = ''00''' WHEN 'S|Both' THEN N' AND (MBMCMemberType = ''S'' OR MBMCMemberType = ''C'')' WHEN 'S|C' THEN N' AND MBMCMemberType = ''C''' WHEN 'S|S' THEN N' AND MBMCMemberType = ''S''' WHEN 'Both|C' THEN N' AND (MBMMemberType = ''P'' OR MBMCMemberType = ''C'')' WHEN 'Both|S' THEN N' AND (MBMMemberType = ''Q'' OR MBMCMemberType = ''S'')' WHEN 'Both|Both' THEN N' AND (MBMMemberType = ''P'' OR MBMMemberType = ''Q'')' ELSE N'' END;
    IF @cancelFrom IS NOT NULL SET @w += N' AND MBMCancelDate >= @cancelFrom';
    IF @cancelTo IS NOT NULL SET @w += N' AND MBMCancelDate <= @cancelTo';
    IF @cancelBlank = 1 SET @w += N' AND (MBMCancelDate IS NULL)';
    IF @dayInsuredFrom IS NOT NULL SET @w += N' AND DayInsured >= @dayInsuredFrom';
    IF @dayInsuredTo IS NOT NULL SET @w += N' AND DayInsured <= @dayInsuredTo';
    IF @dayInsuredBlank = 1 SET @w += N' AND (DayInsured IS NULL)';
    IF @ageDaysFrom IS NOT NULL SET @w += N' AND AgeDays >= @ageDaysFrom';
    IF @ageDaysTo IS NOT NULL SET @w += N' AND AgeDays <= @ageDaysTo';
    IF @ageDaysBlank = 1 SET @w += N' AND (AgeDays IS NULL)';
    IF @delayBordxFrom IS NOT NULL SET @w += N' AND DelayBordx >= @delayBordxFrom';
    IF @delayBordxTo IS NOT NULL SET @w += N' AND DelayBordx <= @delayBordxTo';
    IF @delayBordxBlank = 1 SET @w += N' AND (DelayBordx IS NULL)';
    IF @delayInvBordxFrom IS NOT NULL SET @w += N' AND DelayInvBordx >= @delayInvBordxFrom';
    IF @delayInvBordxTo IS NOT NULL SET @w += N' AND DelayInvBordx <= @delayInvBordxTo';
    IF @delayInvBordxBlank = 1 SET @w += N' AND (DelayInvBordx IS NULL)';
    IF @delayRegFrom IS NOT NULL SET @w += N' AND DelayReg >= @delayRegFrom';
    IF @delayRegTo IS NOT NULL SET @w += N' AND DelayReg <= @delayRegTo';
    IF @delayRegBlank = 1 SET @w += N' AND (DelayReg IS NULL)';
    IF @delayInvDFrom IS NOT NULL SET @w += N' AND DelayInvD >= @delayInvDFrom';
    IF @delayInvDTo IS NOT NULL SET @w += N' AND DelayInvD <= @delayInvDTo';
    IF @delayInvDBlank = 1 SET @w += N' AND (DelayInvD IS NULL)';
    IF @delayInvEFrom IS NOT NULL SET @w += N' AND DelayInvE >= @delayInvEFrom';
    IF @delayInvETo IS NOT NULL SET @w += N' AND DelayInvE <= @delayInvETo';
    IF @delayInvEBlank = 1 SET @w += N' AND (DelayInvE IS NULL)';
    IF @premiumFrom IS NOT NULL SET @w += N' AND MBMPremium >= @premiumFrom';
    IF @premiumTo IS NOT NULL SET @w += N' AND MBMPremium <= @premiumTo';
    IF @premiumBlank = 1 SET @w += N' AND (MBMPremium IS NULL)';
    IF @mcoFrom IS NOT NULL SET @w += N' AND MBMMCOFee >= @mcoFrom';
    IF @mcoTo IS NOT NULL SET @w += N' AND MBMMCOFee <= @mcoTo';
    IF @mcoBlank = 1 SET @w += N' AND (MBMMCOFee IS NULL)';
    IF @employeeFrom IS NOT NULL SET @w += N' AND MBMEmployeeNo >= @employeeFrom';
    IF @employeeTo IS NOT NULL SET @w += N' AND MBMEmployeeNo <= @employeeTo';
    IF @employeeBlank = 1 SET @w += N' AND (MBMEmployeeNo IS NULL OR MBMEmployeeNo = '''')';
    IF @memberTypeBlank = 1 SET @w += N' AND (MBMMemberType IS NULL OR MBMMemberType = '''')';
    IF @adultChild IS NOT NULL SET @w += N' AND MBMAdultChild = @adultChild';
    IF @adultChildBlank = 1 SET @w += N' AND (MBMAdultChild IS NULL OR MBMAdultChild = '''')';
    IF @race IS NOT NULL SET @w += N' AND RACCode = @race';
    IF @raceBlank = 1 SET @w += N' AND (RACCode IS NULL OR RACCode = '''')';
    IF @sex IS NOT NULL SET @w += N' AND MBMSex = @sex';
    IF @sexBlank = 1 SET @w += N' AND (MBMSex IS NULL OR MBMSex = '''')';
    IF @rel IS NOT NULL SET @w += N' AND RELCode = @rel';
    IF @relBlank = 1 SET @w += N' AND (RELCode IS NULL OR RELCode = '''')';
    IF @dataStatus IS NOT NULL SET @w += N' AND MBMDataStatus = @dataStatus';
    IF @dataStatusBlank = 1 SET @w += N' AND (MBMDataStatus IS NULL OR MBMDataStatus = '''')';
    IF @status IS NOT NULL SET @w += CASE @status WHEN '0' THEN N' AND MBMStatus = ''A''' WHEN '1' THEN N' AND MBMStatus = ''C''' WHEN '2' THEN N' AND MBMStatus = ''D''' WHEN '3' THEN N' AND (MBMStatus = ''A'' OR MBMStatus = ''C'')' WHEN '4' THEN N' AND (MBMStatus = ''A'' OR MBMStatus = ''S'')' WHEN '5' THEN N' AND (MBMStatus = ''A'' OR MBMStatus = ''C'' OR MBMStatus = ''S'')' WHEN '6' THEN N' AND MBMStatus = ''S''' ELSE N'' END;
    IF @statusBlank = 1 SET @w += N' AND (MBMStatus IS NULL OR MBMStatus = '''')';
    IF @expired IS NOT NULL SET @w += N' AND MBMExpiredStatus = @expired';
    IF @expiredBlank = 1 SET @w += N' AND (MBMExpiredStatus IS NULL OR MBMExpiredStatus = '''')';
    IF @renewal IS NOT NULL SET @w += CASE @renewal WHEN 'P' THEN N' AND (MBMRenewFlag = ''P'')' WHEN 'N' THEN N' AND (MBMRenewFlag IS NULL OR MBMRenewFlag = '''' OR MBMRenewFlag = ''N'')' WHEN 'R' THEN N' AND (MBMRenewFlag = ''R'')' WHEN 'T' THEN N' AND (MBMRenewFlag = ''T'')' ELSE N'' END;
    IF @renewalBlank = 1 SET @w += N' AND (MBMRenewFlag IS NULL OR MBMRenewFlag = '''')';
    IF @premMode IS NOT NULL SET @w += N' AND MBMInstallment = @premMode';
    IF @premModeBlank = 1 SET @w += N' AND (MBMInstallment IS NULL OR MBMInstallment = '''')';
    IF @state IS NOT NULL SET @w += N' AND STACode = @state';
    IF @stateBlank = 1 SET @w += N' AND (STACode IS NULL OR STACode = '''')';
    IF @takeover IS NOT NULL SET @w += N' AND MBMTakeover = @takeover';
    IF @takeoverBlank = 1 SET @w += N' AND (MBMTakeover IS NULL OR MBMTakeover = '''')';
    IF @agent IS NOT NULL SET @w += N' AND MBMAgentCode = @agent';
    IF @agentBlank = 1 SET @w += N' AND (MBMAgentCode IS NULL OR MBMAgentCode = '''')';
    IF @mbmNo IS NOT NULL SET @w += N' AND MBMNumber = @mbmNo';
    IF @mbmNoBlank = 1 SET @w += N' AND (MBMNumber IS NULL OR MBMNumber = '''')';
    IF @policy IS NOT NULL SET @w += N' AND MBMPolicyNo = @policy';
    IF @policyBlank = 1 SET @w += N' AND (MBMPolicyNo IS NULL OR MBMPolicyNo = '''')';
    IF @payor IS NOT NULL SET @w += N' AND PAYCode = @payor';
    IF @company IS NOT NULL SET @w += N' AND MBMGroupCompany = @company';
    IF @companyBlank = 1 SET @w += N' AND (MBMGroupCompany IS NULL OR MBMGroupCompany = '''')';
    IF @planFrom IS NOT NULL SET @w += N' AND PLNCode >= @planFrom';
    IF @planTo IS NOT NULL SET @w += N' AND PLNCode <= @planTo';
    IF @planBlank = 1 SET @w += N' AND (PLNCode IS NULL OR PLNCode = '''')';
    IF @ageFrom IS NOT NULL SET @w += N' AND AGECode >= @ageFrom';
    IF @ageTo IS NOT NULL SET @w += N' AND AGECode <= @ageTo';
    IF @ageBlank = 1 SET @w += N' AND (AGECode IS NULL OR AGECode = '''')';
    IF @effFrom IS NOT NULL SET @w += N' AND MBMPayorEffDate >= @effFrom';
    IF @effTo IS NOT NULL SET @w += N' AND MBMPayorEffDate <= @effTo';
    IF @effBlank = 1 SET @w += N' AND (MBMPayorEffDate IS NULL OR MBMPayorEffDate = '''')';
    IF @expFrom IS NOT NULL SET @w += N' AND MBMPayorExpDate >= @expFrom';
    IF @expTo IS NOT NULL SET @w += N' AND MBMPayorExpDate <= @expTo';
    IF @expBlank = 1 SET @w += N' AND (MBMPayorExpDate IS NULL OR MBMPayorExpDate = '''')';
    IF @joinedFrom IS NOT NULL SET @w += N' AND MBMFirstJoinedDate >= @joinedFrom';
    IF @joinedTo IS NOT NULL SET @w += N' AND MBMFirstJoinedDate <= @joinedTo';
    IF @joinedBlank = 1 SET @w += N' AND (MBMFirstJoinedDate IS NULL OR MBMFirstJoinedDate = '''')';
    IF @dorFrom IS NOT NULL SET @w += N' AND MBMValueDate >= @dorFrom';
    IF @dorTo IS NOT NULL SET @w += N' AND MBMValueDate <= @dorTo';
    IF @dorBlank = 1 SET @w += N' AND (MBMValueDate IS NULL OR MBMValueDate = '''')';
    IF @dobFrom IS NOT NULL SET @w += N' AND MBMDOB >= @dobFrom';
    IF @dobTo IS NOT NULL SET @w += N' AND MBMDOB <= @dobTo';
    IF @dobBlank = 1 SET @w += N' AND (MBMDOB IS NULL OR MBMDOB = '''')';
    IF @bordxFrom IS NOT NULL SET @w += N' AND MBMBordxDate >= @bordxFrom';
    IF @bordxTo IS NOT NULL SET @w += N' AND MBMBordxDate <= @bordxTo';
    IF @bordxBlank = 1 SET @w += N' AND (MBMBordxDate IS NULL OR MBMBordxDate = '''')';
    IF @batchFrom IS NOT NULL SET @w += N' AND MBMBatchNo >= @batchFrom';
    IF @batchTo IS NOT NULL SET @w += N' AND MBMBatchNo <= @batchTo';
    IF @batchBlank = 1 SET @w += N' AND (MBMBatchNo IS NULL OR MBMBatchNo = '''')';
    IF @postcodeFrom IS NOT NULL SET @w += N' AND MBMPostCode >= @postcodeFrom';
    IF @postcodeTo IS NOT NULL SET @w += N' AND MBMPostCode <= @postcodeTo';
    IF @postcodeBlank = 1 SET @w += N' AND (MBMPostCode IS NULL OR MBMPostCode = '''')';
    IF @polYearFrom IS NOT NULL SET @w += N' AND MBMPolicyYear >= @polYearFrom';
    IF @polYearTo IS NOT NULL SET @w += N' AND MBMPolicyYear <= @polYearTo';
    IF @polYearBlank = 1 SET @w += N' AND (MBMPolicyYear IS NULL OR MBMPolicyYear = '''')';
    IF @invoiceFrom IS NOT NULL SET @w += N' AND MBMInvoiceNo >= @invoiceFrom';
    IF @invoiceTo IS NOT NULL SET @w += N' AND MBMInvoiceNo <= @invoiceTo';
    IF @invoiceBlank = 1 SET @w += N' AND (MBMInvoiceNo IS NULL OR MBMInvoiceNo = '''')';
    IF @invoiceDateFrom IS NOT NULL SET @w += N' AND MBMInvoiceDate >= @invoiceDateFrom';
    IF @invoiceDateTo IS NOT NULL SET @w += N' AND MBMInvoiceDate <= @invoiceDateTo';
    IF @invoiceDateBlank = 1 SET @w += N' AND (MBMInvoiceDate IS NULL OR MBMInvoiceDate = '''')';
    IF @icFrom IS NOT NULL SET @w += N' AND MBMIcBcPp >= @icFrom';
    IF @icTo IS NOT NULL SET @w += N' AND MBMIcBcPp <= @icTo';
    IF @icBlank = 1 SET @w += N' AND (MBMIcBcPp IS NULL OR MBMIcBcPp = '''')';
    IF @label IS NOT NULL SET @w += N' AND LabelStatus = @label';
    IF @labelBlank = 1 SET @w += N' AND (LabelStatus IS NULL OR LabelStatus = '''')';
    IF @department IS NOT NULL SET @w += N' AND MBMDepartment = @department';
    IF @departmentBlank = 1 SET @w += N' AND (MBMDepartment IS NULL OR MBMDepartment = '''')';
    IF @cardFrom IS NOT NULL SET @w += N' AND MBMCLNCardSentDate >= @cardFrom';
    IF @cardTo IS NOT NULL SET @w += N' AND MBMCLNCardSentDate <= @cardTo';
    IF @cardBlank = 1 SET @w += N' AND (MBMCLNCardSentDate IS NULL OR MBMCLNCardSentDate = '''')';
    IF @endtFrom IS NOT NULL SET @w += N' AND MBMCLNEndtEffDate >= @endtFrom';
    IF @endtTo IS NOT NULL SET @w += N' AND MBMCLNEndtEffDate <= @endtTo';
    IF @endtBlank = 1 SET @w += N' AND (MBMCLNEndtEffDate IS NULL OR MBMCLNEndtEffDate = '''')';
    IF @suppBordxFrom IS NOT NULL SET @w += N' AND MBMBordxDate >= @suppBordxFrom';
    IF @suppBordxTo IS NOT NULL SET @w += N' AND MBMBordxDate <= @suppBordxTo';
    IF @detailPrinType = 'C' SET @w += N' AND (MBMMemberType = ''P'' OR MBMMemberType = ''C'')'; IF @detailPrinType = 'S' SET @w += N' AND (MBMMemberType = ''Q'')';
    IF @suppRegFrom IS NOT NULL SET @w += N' AND MBMCRegDate >= @suppRegFrom';
    IF @suppRegTo IS NOT NULL SET @w += N' AND MBMCRegDate <= @suppRegTo';
    IF @suppRegBlank = 1 SET @w += N' AND (MBMCRegDate IS NULL OR MBMCRegDate = '''')';
    DECLARE @sql nvarchar(max) = CAST(@b AS nvarchar(max)) + @w + N' ' + @a;
    EXEC HISDB.sys.sp_executesql @sql, N'@hlt nvarchar(1), @hltBlank bit, @payorBlank bit, @ins nvarchar(1), @insBlank bit, @prinType nvarchar(10), @cancelFrom datetime, @cancelTo datetime, @cancelBlank bit, @dayInsuredFrom nvarchar(20), @dayInsuredTo nvarchar(20), @dayInsuredBlank bit, @ageDaysFrom nvarchar(20), @ageDaysTo nvarchar(20), @ageDaysBlank bit, @delayBordxFrom nvarchar(20), @delayBordxTo nvarchar(20), @delayBordxBlank bit, @delayInvBordxFrom nvarchar(20), @delayInvBordxTo nvarchar(20), @delayInvBordxBlank bit, @delayRegFrom nvarchar(20), @delayRegTo nvarchar(20), @delayRegBlank bit, @delayInvDFrom nvarchar(20), @delayInvDTo nvarchar(20), @delayInvDBlank bit, @delayInvEFrom nvarchar(20), @delayInvETo nvarchar(20), @delayInvEBlank bit, @premiumFrom money, @premiumTo money, @premiumBlank bit, @mcoFrom money, @mcoTo money, @mcoBlank bit, @employeeFrom nvarchar(50), @employeeTo nvarchar(50), @employeeBlank bit, @memberTypeBlank bit, @adultChild nvarchar(1), @adultChildBlank bit, @race nvarchar(1), @raceBlank bit, @sex nvarchar(1), @sexBlank bit, @rel nvarchar(1), @relBlank bit, @dataStatus nvarchar(10), @dataStatusBlank bit, @status nvarchar(1), @statusBlank bit, @expired nvarchar(1), @expiredBlank bit, @renewal nvarchar(1), @renewalBlank bit, @premMode nvarchar(1), @premModeBlank bit, @state nvarchar(100), @stateBlank bit, @takeover nvarchar(1), @takeoverBlank bit, @agent nvarchar(30), @agentBlank bit, @mbmNo nvarchar(30), @mbmNoBlank bit, @policy nvarchar(60), @policyBlank bit, @payor nvarchar(10), @company nvarchar(200), @companyBlank bit, @planFrom nvarchar(20), @planTo nvarchar(20), @planBlank bit, @ageFrom nvarchar(20), @ageTo nvarchar(20), @ageBlank bit, @effFrom datetime, @effTo datetime, @effBlank bit, @expFrom datetime, @expTo datetime, @expBlank bit, @joinedFrom datetime, @joinedTo datetime, @joinedBlank bit, @dorFrom datetime, @dorTo datetime, @dorBlank bit, @dobFrom datetime, @dobTo datetime, @dobBlank bit, @bordxFrom datetime, @bordxTo datetime, @bordxBlank bit, @batchFrom nvarchar(50), @batchTo nvarchar(50), @batchBlank bit, @postcodeFrom nvarchar(50), @postcodeTo nvarchar(50), @postcodeBlank bit, @polYearFrom nvarchar(50), @polYearTo nvarchar(50), @polYearBlank bit, @invoiceFrom nvarchar(50), @invoiceTo nvarchar(50), @invoiceBlank bit, @invoiceDateFrom datetime, @invoiceDateTo datetime, @invoiceDateBlank bit, @icFrom nvarchar(50), @icTo nvarchar(50), @icBlank bit, @label nvarchar(1), @labelBlank bit, @department nvarchar(200), @departmentBlank bit, @cardFrom datetime, @cardTo datetime, @cardBlank bit, @endtFrom datetime, @endtTo datetime, @endtBlank bit, @suppBordxFrom datetime, @suppBordxTo datetime, @detailPrinType nvarchar(1), @suppRegFrom datetime, @suppRegTo datetime, @suppRegBlank bit',
        @hlt, @hltBlank, @payorBlank, @ins, @insBlank, @prinType, @cancelFrom, @cancelTo, @cancelBlank, @dayInsuredFrom, @dayInsuredTo, @dayInsuredBlank, @ageDaysFrom, @ageDaysTo, @ageDaysBlank, @delayBordxFrom, @delayBordxTo, @delayBordxBlank, @delayInvBordxFrom, @delayInvBordxTo, @delayInvBordxBlank, @delayRegFrom, @delayRegTo, @delayRegBlank, @delayInvDFrom, @delayInvDTo, @delayInvDBlank, @delayInvEFrom, @delayInvETo, @delayInvEBlank, @premiumFrom, @premiumTo, @premiumBlank, @mcoFrom, @mcoTo, @mcoBlank, @employeeFrom, @employeeTo, @employeeBlank, @memberTypeBlank, @adultChild, @adultChildBlank, @race, @raceBlank, @sex, @sexBlank, @rel, @relBlank, @dataStatus, @dataStatusBlank, @status, @statusBlank, @expired, @expiredBlank, @renewal, @renewalBlank, @premMode, @premModeBlank, @state, @stateBlank, @takeover, @takeoverBlank, @agent, @agentBlank, @mbmNo, @mbmNoBlank, @policy, @policyBlank, @payor, @company, @companyBlank, @planFrom, @planTo, @planBlank, @ageFrom, @ageTo, @ageBlank, @effFrom, @effTo, @effBlank, @expFrom, @expTo, @expBlank, @joinedFrom, @joinedTo, @joinedBlank, @dorFrom, @dorTo, @dorBlank, @dobFrom, @dobTo, @dobBlank, @bordxFrom, @bordxTo, @bordxBlank, @batchFrom, @batchTo, @batchBlank, @postcodeFrom, @postcodeTo, @postcodeBlank, @polYearFrom, @polYearTo, @polYearBlank, @invoiceFrom, @invoiceTo, @invoiceBlank, @invoiceDateFrom, @invoiceDateTo, @invoiceDateBlank, @icFrom, @icTo, @icBlank, @label, @labelBlank, @department, @departmentBlank, @cardFrom, @cardTo, @cardBlank, @endtFrom, @endtTo, @endtBlank, @suppBordxFrom, @suppBordxTo, @detailPrinType, @suppRegFrom, @suppRegTo, @suppRegBlank;
END
GO
