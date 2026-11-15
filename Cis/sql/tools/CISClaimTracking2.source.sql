CREATE procedure CISClaimTracking2

	@strAppend as nvarchar (2000), @strAppendCase as int

              
as
	if @strAppendCase = 1

		begin
         	      		set nocount on
			       declare @l_sqlstr as nvarchar(2000)
		                    declare @l_sqlgroupby as nvarchar(500)

		                    set  @l_sqlstr =  N'SELECT AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby = 'Group by AGEBand'
			       set @l_sqlstr = @l_sqlstr + @strAppend + @l_sqlgroupby
			       exec  sp_executesql  @l_sqlstr
		end

	if @strAppendCase = 3

		begin
         	      		set nocount on
			       declare @l_sqlstr3 as nvarchar(2000)
		                    declare @l_sqlgroupby3 as nvarchar(500)

		                    set  @l_sqlstr3 =  N'SELECT AGEBand, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby3 = 'Group by AGEBand, Sex'
			       set @l_sqlstr3 = @l_sqlstr3 + @strAppend + @l_sqlgroupby3
			       exec  sp_executesql  @l_sqlstr3
		end

	if @strAppendCase = 4

		begin
         	      		set nocount on
			       declare @l_sqlstr4 as nvarchar(2000)
		                    declare @l_sqlgroupby4 as nvarchar(500)

		                    set  @l_sqlstr4 =  N'SELECT BordxRefNo, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby4 = 'Group by BordxRefNo Order by BordxRefNo'
			       set @l_sqlstr4 = @l_sqlstr4 + @strAppend + @l_sqlgroupby4
			       exec  sp_executesql  @l_sqlstr4
		end

	if @strAppendCase = 5

		begin
         	      		set nocount on
			       declare @l_sqlstr5 as nvarchar(2000)

		                    set  @l_sqlstr5 =  N'SELECT CASNumber, MBMNumber, PatientName, DateVisit,
          						CLMGPFeeBasicMed, CLMGPFeeMedInj, CLMGPFeeLTMed, CLMSPFeeBasicMed,
          						CLMSPFeeMedInj, CLMSPFeeLTMed, CLMXRay, CLMUltraSound, CLMECG, CLMDressing,
          						CLMNebulizer, CLMID, CLMTSSTO, CLMMinorSurgery, CLMPOthers,
          						CLMBloodDrawing, CLMBloodGlucose, CLMUrine, CLMPapSmear, CLMLOthers,
          						CLMDTPPolio, CLMDTPPolioHIB, CLMHepatitisB, CLMMMR, CLMIOthers,
          						CLMDental, CLMMaternity, CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlstr5 = @l_sqlstr5 + @strAppend
			       exec  sp_executesql  @l_sqlstr5
		end

	if @strAppendCase = 6

		begin
         	      		set nocount on
			       declare @l_sqlstr6 as nvarchar(2000)
		                    declare @l_sqlgroupby6 as nvarchar(500)

		                    set  @l_sqlstr6 =  N'SELECT ClaimNatureCode, SUM(CONVERT(Money,CLMBoardAmt)) AS CLMBoardAmt, SUM(CONVERT(Money,CLMICUAmt)) AS CLMICUAmt,
          						SUM(CONVERT(Money,CLMServices)) AS CLMServices, SUM(CONVERT(Money,CLMTheatre)) AS CLMTheatre, SUM(CONVERT(Money,CLMPreHSAmt)) AS CLMPreHSAmt,
          						SUM(CONVERT(Money,CLMSurgicalAmt)) AS CLMSurgicalAmt, SUM(CONVERT(Money,CLMAna)) AS CLMAna, SUM(CONVERT(Money,CLMNSPreHos)) AS CLMNSPreHos,
          						SUM(CONVERT(Money,CLMPhysician)) AS CLMPhysician, SUM(CONVERT(Money,CLMPost)) AS CLMPost, SUM(CONVERT(Money,CLMAmbulance)) AS CLMAmbulance,
          						SUM(CONVERT(Money,CLMOutPatient)) AS CLMOutPatient, SUM(CONVERT(Money,CLMMonthly)) AS CLMMonthly, SUM(CONVERT(Money,CLMOrgan)) AS CLMOrgan,
          						SUM(CONVERT(Money,CLMMedical)) AS CLMMedical, SUM(CONVERT(Money,CLMPhone)) AS CLMPhone, SUM(CONVERT(Money,CLMGuardianAmt)) AS CLMGuardianAmt,
          						SUM(CONVERT(Money,CLMAdmin)) AS CLMAdmin, SUM(CONVERT(Money,CLMTax)) AS CLMTax, SUM(CONVERT(Money,CLMOthers2)) AS CLMOthers2, SUM(CONVERT(Money,CLMCashAmt)) AS CLMCashAmt
          						From VIEW_CLMReport_Act Where 0 = 0'
			       set @l_sqlgroupby6 = 'Group by ClaimNatureCode Order by ClaimNatureCode'
			       set @l_sqlstr6 = @l_sqlstr6 + @strAppend + @l_sqlgroupby6
			       exec  sp_executesql  @l_sqlstr6
		end

	if @strAppendCase = 7

		begin
         	      		set nocount on
			       declare @l_sqlstr7 as nvarchar(2000)
		                    declare @l_sqlgroupby7 as nvarchar(500)

		                    set  @l_sqlstr7 =  N'SELECT ClinicCode, CLNName, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
        						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby7 = 'Group by ClinicCode, CLNName Order by ClinicCode, CLNName'
			       set @l_sqlstr7 = @l_sqlstr7 + @strAppend + @l_sqlgroupby7
			       exec  sp_executesql  @l_sqlstr7
		end
	if @strAppendCase = 8

		begin
         	      		set nocount on
			       declare @l_sqlstr8 as nvarchar(2000)
		                    declare @l_sqlgroupby8 as nvarchar(500)

		                    set  @l_sqlstr8 =  N'SELECT CASTreatReason1, ReasonDesc, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby8 = 'Group by CASTreatReason1, ReasonDesc Order by CASTreatReason1, ReasonDesc'
			       set @l_sqlstr8 = @l_sqlstr8 + @strAppend + @l_sqlgroupby8
			       exec  sp_executesql  @l_sqlstr8
		end

	if @strAppendCase = 9

		begin
         	      		set nocount on
			       declare @l_sqlstr9 as nvarchar(2000)
		                    declare @l_sqlgroupby9 as nvarchar(500)

		                    set  @l_sqlstr9 =  N'SELECT MBMGroupCompany, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby9 = 'Group by MBMGroupCompany Order by MBMGroupCompany'
			       set @l_sqlstr9 = @l_sqlstr9 + @strAppend + @l_sqlgroupby9
			       exec  sp_executesql  @l_sqlstr9
		end

	if @strAppendCase = 10

		begin
         	      		set nocount on
			       declare @l_sqlstr10 as nvarchar(2000)
		                    declare @l_sqlgroupby10 as nvarchar(500)

		                    set  @l_sqlstr10 =  N'SELECT INSCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby10 = 'Group by INSCode Order by INSCode'
			       set @l_sqlstr10 = @l_sqlstr10 + @strAppend + @l_sqlgroupby10
			       exec  sp_executesql  @l_sqlstr10
		end

	if @strAppendCase = 11

		begin
         	      		set nocount on
			       declare @l_sqlstr11 as nvarchar(2000)
		                    declare @l_sqlgroupby11 as nvarchar(500)

		                    set  @l_sqlstr11 =  N'SELECT INSCode, AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
        						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby11 = 'Group by INSCode, AGEBand Order by INSCode, Ageband'
			       set @l_sqlstr11 = @l_sqlstr11 + @strAppend + @l_sqlgroupby11
			       exec  sp_executesql  @l_sqlstr11
		end
			
	if @strAppendCase = 12

		begin
         	      		set nocount on
			       declare @l_sqlstr12 as nvarchar(2000)
		                    declare @l_sqlgroupby12 as nvarchar(500)

		                    set  @l_sqlstr12 =  N'SELECT INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
         						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, 
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby12 = 'Group by INSCode, PLNCode Order by INSCode, PLNCode'
			       set @l_sqlstr12 = @l_sqlstr12 + @strAppend + @l_sqlgroupby12
			       exec  sp_executesql  @l_sqlstr12
			
		end

	if @strAppendCase = 13

		begin
         	      		set nocount on
			       declare @l_sqlstr13 as nvarchar(2000)
		                    declare @l_sqlgroupby13 as nvarchar(500)

		                    set  @l_sqlstr13 =  N'SELECT INSCode, Relation, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, 
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby13 = 'Group by INSCode, Relation Order by INSCode, Relation'
			       set @l_sqlstr13 = @l_sqlstr13 + @strAppend + @l_sqlgroupby13
			       exec  sp_executesql  @l_sqlstr13
		end

	if @strAppendCase = 14

		begin
         	      		set nocount on
			       declare @l_sqlstr14 as nvarchar(2000)
		                    declare @l_sqlgroupby14 as nvarchar(500)

		                    set  @l_sqlstr14 =  N'SELECT MBMPolicyYear, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, 
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby14 = 'Group by MBMPolicyYear Order by MBMPolicyYear'
			       set @l_sqlstr14 = @l_sqlstr14 + @strAppend + @l_sqlgroupby14
			       exec  sp_executesql  @l_sqlstr14
		end

	if @strAppendCase = 15

		begin
         	      		set nocount on
			       declare @l_sqlstr15 as nvarchar(2000)
		                    declare @l_sqlgroupby15 as nvarchar(500)

		                    set  @l_sqlstr15 =  N'SELECT INSCode, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby15 = 'Group by INSCode, Sex Order by INSCode, Sex'
			       set @l_sqlstr15 = @l_sqlstr15 + @strAppend + @l_sqlgroupby15
			       exec  sp_executesql  @l_sqlstr15
		end

	if @strAppendCase = 16

		begin
         	      		set nocount on
			       declare @l_sqlstr16 as nvarchar(2000)
		                    declare @l_sqlgroupby16 as nvarchar(500)

		                    set @l_sqlstr16 =  N'SELECT PAYCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby16 = 'Group by PAYCode Order by PAYCode '
			       set @l_sqlstr16 = @l_sqlstr16 + @strAppend + @l_sqlgroupby16
			       exec  sp_executesql  @l_sqlstr16
		end

	if @strAppendCase = 17

		begin
         	      		set nocount on
			       declare @l_sqlstr17 as nvarchar(2000)
		                    declare @l_sqlgroupby17 as nvarchar(500)

		                    set  @l_sqlstr17 =  N'SELECT PAYCode, INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, 
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, 
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby17 = 'Group by PAYCode, INSCode, PLNCode Order by PAYCode, INSCode, PLNCode'
			       set @l_sqlstr17 = @l_sqlstr17 + @strAppend + @l_sqlgroupby17
			       exec  sp_executesql  @l_sqlstr17
		end

	if @strAppendCase = 18

		begin
         	      		set nocount on
			       declare @l_sqlstr18 as nvarchar(2000)
		                    declare @l_sqlgroupby18 as nvarchar(500)

		                    set  @l_sqlstr18 =  N'SELECT Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
         						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby18 = 'Group by Sex Order by Sex'
			       set @l_sqlstr18 = @l_sqlstr18 + @strAppend + @l_sqlgroupby18
			       exec  sp_executesql  @l_sqlstr18
		end

	if @strAppendCase = 19

		begin
         	      		set nocount on
			       declare @l_sqlstr19 as nvarchar(2000)
		                    declare @l_sqlgroupby19 as nvarchar(500)

		                    set  @l_sqlstr19 =  N'SELECT MBMTakeOver, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby19 = 'Group by MBMTakeOver Order by MBMTakeOver'
			       set @l_sqlstr19 = @l_sqlstr19 + @strAppend + @l_sqlgroupby19
			       exec  sp_executesql  @l_sqlstr19
		end

	if @strAppendCase = 20

		begin
         	      		set nocount on
			       declare @l_sqlstr20 as nvarchar(2000)
		                    declare @l_sqlgroupby20 as nvarchar(500)

		                    set  @l_sqlstr20 =  N'SELECT AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby20 = 'Group by AGEBand Order by AGEBand'
			       set @l_sqlstr20 = @l_sqlstr20 + @strAppend + @l_sqlgroupby20
			       exec  sp_executesql  @l_sqlstr20
		end

	if @strAppendCase = 21

		begin
         	      		set nocount on
			       declare @l_sqlstr21 as nvarchar(2000)
		                    declare @l_sqlgroupby21 as nvarchar(500)

		                    set  @l_sqlstr21 =  N'SELECT AGEBand, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby21 = 'Group by AGEBand, Sex Order by Ageband, Sex'
			       set @l_sqlstr21 = @l_sqlstr21 + @strAppend + @l_sqlgroupby21
			       exec  sp_executesql  @l_sqlstr21
		end

	if @strAppendCase = 22

		begin
         	      		set nocount on
			       declare @l_sqlstr22 as nvarchar(2000)
		                    declare @l_sqlgroupby22 as nvarchar(500)

		                    set  @l_sqlstr22 =  N'SELECT BordxRefNo, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby22 = 'Group by BordxRefNo Order by BordxRefNo'
			       set @l_sqlstr22 = @l_sqlstr22 + @strAppend + @l_sqlgroupby22
			       exec  sp_executesql  @l_sqlstr22
		end

	if @strAppendCase = 23

		begin
         	      		set nocount on
			       declare @l_sqlstr23 as nvarchar(2000)

		                    set  @l_sqlstr23 =  N'SELECT CASNumber, MBMNumber, PatientName, DateVisit,
          						CLMGPFeeBasicMed, CLMGPFeeMedInj, CLMGPFeeLTMed, CLMSPFeeBasicMed,
          						CLMSPFeeMedInj, CLMSPFeeLTMed, CLMXRay, CLMUltraSound, CLMECG, CLMDressing,
          						CLMNebulizer, CLMID, CLMTSSTO, CLMMinorSurgery, CLMPOthers,
          						CLMBloodDrawing, CLMBloodGlucose, CLMUrine, CLMPapSmear, CLMLOthers,
          						CLMDTPPolio, CLMDTPPolioHIB, CLMHepatitisB, CLMMMR, CLMIOthers,
          						CLMDental, CLMMaternity, CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlstr23 = @l_sqlstr23 + @strAppend
			       exec  sp_executesql  @l_sqlstr23
		end

	if @strAppendCase = 24

		begin
         	      		set nocount on
			       declare @l_sqlstr24 as nvarchar(2000)
		                    declare @l_sqlgroupby24 as nvarchar(500)

		                    set  @l_sqlstr24 =  N'SELECT ClinicCode, CLNName, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj, 
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby24 = 'Group by ClinicCode, CLNName Order by ClinicCode, CLNName'
			       set @l_sqlstr24 = @l_sqlstr24 + @strAppend + @l_sqlgroupby24
			       exec  sp_executesql  @l_sqlstr24
		end

	if @strAppendCase = 25

		begin
         	      		set nocount on
			       declare @l_sqlstr25 as nvarchar(2000)
		                    declare @l_sqlgroupby25 as nvarchar(500)

		                    set  @l_sqlstr25 =  N'SELECT CASTreatReason1, ReasonDesc, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby25 = 'Group by CASTreatReason1, ReasonDesc Order by CASTreatReason1, ReasonDesc'
			       set @l_sqlstr25 = @l_sqlstr25 + @strAppend + @l_sqlgroupby25
			       exec  sp_executesql  @l_sqlstr25
		end

	if @strAppendCase = 26

		begin
         	      		set nocount on
			       declare @l_sqlstr26 as nvarchar(2000)
		                    declare @l_sqlgroupby26 as nvarchar(500)

		                    set  @l_sqlstr26 =  N'SELECT MBMGroupCompany, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID, 
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup 
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby26 = 'Group by MBMGroupCompany Order by MBMGroupCompany'
			       set @l_sqlstr26 = @l_sqlstr26 + @strAppend + @l_sqlgroupby26
			       exec  sp_executesql  @l_sqlstr26
		end

	if @strAppendCase = 27

		begin
         	      		set nocount on
			       declare @l_sqlstr27 as nvarchar(2000)
		                    declare @l_sqlgroupby27 as nvarchar(500)

		                    set  @l_sqlstr27 =  N'SELECT INSCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby27 = 'Group by INSCode Order by INSCode'
			       set @l_sqlstr27 = @l_sqlstr27 + @strAppend + @l_sqlgroupby27
			       exec  sp_executesql  @l_sqlstr27
		end

	if @strAppendCase = 28

		begin
         	      		set nocount on
			       declare @l_sqlstr28 as nvarchar(2000)
		                    declare @l_sqlgroupby28 as nvarchar(500)

		                    set  @l_sqlstr28 =  N'SELECT INSCode, AGEBand, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby28 = 'Group by INSCode, AGEBand Order by INSCode, Ageband'
			       set @l_sqlstr28 = @l_sqlstr28 + @strAppend + @l_sqlgroupby28
			       exec  sp_executesql  @l_sqlstr28
		end

	if @strAppendCase = 29

		begin
         	      		set nocount on
			       declare @l_sqlstr29 as nvarchar(2000)
		                    declare @l_sqlgroupby29 as nvarchar(500)

		                    set  @l_sqlstr29 =  N'SELECT INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby29 = 'Group by INSCode, PLNCode Order by INSCode, PLNCode'
			       set @l_sqlstr29 = @l_sqlstr29 + @strAppend + @l_sqlgroupby29
			       exec  sp_executesql  @l_sqlstr29
		end

	if @strAppendCase = 30

		begin
         	      		set nocount on
			       declare @l_sqlstr30 as nvarchar(2000)
		                    declare @l_sqlgroupby30 as nvarchar(500)

		                    set  @l_sqlstr30 =  N'SELECT INSCode, Relation, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby30 = 'Group by INSCode, Relation Order by INSCode, Relation'
			       set @l_sqlstr30 = @l_sqlstr30 + @strAppend + @l_sqlgroupby30
			       exec  sp_executesql  @l_sqlstr30
		end

	if @strAppendCase = 31

		begin
         	      		set nocount on
			       declare @l_sqlstr31 as nvarchar(2000)
		                    declare @l_sqlgroupby31 as nvarchar(500)

		                    set  @l_sqlstr31 =  N'SELECT INSCode, Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby31 = 'Group by INSCode, Sex Order by INSCode, Sex'
			       set @l_sqlstr31 = @l_sqlstr31 + @strAppend + @l_sqlgroupby31
			       exec  sp_executesql  @l_sqlstr31
		end

	if @strAppendCase = 32

		begin
         	      		set nocount on
			       declare @l_sqlstr32 as nvarchar(2000)
		                    declare @l_sqlgroupby32 as nvarchar(500)

		                    set  @l_sqlstr32 =  N'SELECT PAYCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, 
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby32 = 'Group by PAYCode Order by PAYCode'
			       set @l_sqlstr32 = @l_sqlstr32 + @strAppend + @l_sqlgroupby32
			       exec  sp_executesql  @l_sqlstr32
		end

	if @strAppendCase = 33

		begin
         	      		set nocount on
			       declare @l_sqlstr33 as nvarchar(2000)
		                    declare @l_sqlgroupby33 as nvarchar(500)

		                    set  @l_sqlstr33 =  N'SELECT PAYCode, INSCode, PLNCode, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound, 
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby33 = 'Group by PAYCode, INSCode, PLNCode Order by PAYCode, INSCode, PLNCode'
			       set @l_sqlstr33 = @l_sqlstr33 + @strAppend + @l_sqlgroupby33
			       exec  sp_executesql  @l_sqlstr33
		end

	if @strAppendCase = 34

		begin
         	      		set nocount on
			       declare @l_sqlstr34 as nvarchar(2000)
		                    declare @l_sqlgroupby34 as nvarchar(500)

		                    set  @l_sqlstr34 =  N'SELECT MBMPolicyYear, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
         						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, 
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby34 = 'Group by MBMPolicyYear Order by MBMPolicyYear'
			       set @l_sqlstr34 = @l_sqlstr34 + @strAppend +  @l_sqlgroupby34
			       exec  sp_executesql  @l_sqlstr34
		end

	if @strAppendCase = 35

		begin
         	      		set nocount on
			       declare @l_sqlstr35 as nvarchar(2000)
		                    declare @l_sqlgroupby35 as nvarchar(500)

		                    set  @l_sqlstr35 =  N'SELECT Sex, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers, 
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby35 = 'Group by Sex Order by Sex'
			       set @l_sqlstr35 = @l_sqlstr35 + @strAppend + @l_sqlgroupby35
			       exec  sp_executesql  @l_sqlstr35
		end

	if @strAppendCase = 36

		begin
         	      		set nocount on
			       declare @l_sqlstr36 as nvarchar(2000)
		                    declare @l_sqlgroupby36 as nvarchar(500)

		                    set  @l_sqlstr36 =  N'SELECT MBMTakeOver, SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup 
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlgroupby36 = 'Group by MBMTakeOver Order by MBMTakeOver'
			       set @l_sqlstr36 = @l_sqlstr36 + @strAppend + @l_sqlgroupby36
			       exec  sp_executesql  @l_sqlstr36
		end

	if @strAppendCase = 37

		begin
         	      		set nocount on
			       declare @l_sqlstr37 as nvarchar(2000)

		                    set  @l_sqlstr37 =  N'SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_Actual Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlstr37 = @l_sqlstr37 + @strAppend
			       exec  sp_executesql  @l_sqlstr37
		end

	if @strAppendCase = 38

		begin
         	      		set nocount on
			       declare @l_sqlstr38 as nvarchar(2000)

		                    set  @l_sqlstr38 =  N'SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup 
          						From VIEW_CLMReport_Approved Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlstr38 = @l_sqlstr38 + @strAppend
			       exec  sp_executesql  @l_sqlstr38
		end

	if @strAppendCase = 39

		begin
         	      		set nocount on
			       declare @l_sqlstr39 as nvarchar(2000)

		                    set  @l_sqlstr39 =  N'SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine, 
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental, 
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup 
          						From VIEW_CLMReport_Clinic Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlstr39 = @l_sqlstr39 + @strAppend
			       exec  sp_executesql  @l_sqlstr39
		end

	if @strAppendCase = 40

		begin
         	      		set nocount on
			       declare @l_sqlstr40 as nvarchar(2000)

		                    set  @l_sqlstr40 =  N'SELECT SUM(CLMGPFeeBasicMed) AS CLMGPFeeBasicMed, SUM(CLMGPFeeMedInj) AS CLMGPFeeMedInj,
          						SUM(CLMGPFeeLTMed) AS CLMGPFeeLTMed, SUM(CLMSPFeeBasicMed) AS CLMSPFeeBasicMed, SUM(CLMSPFeeMedInj) AS CLMSPFeeMedInj,
          						SUM(CLMSPFeeLTMed) AS CLMSPFeeLTMed, SUM(CLMXRay) AS CLMXRay, SUM(CLMUltraSound) AS CLMUltraSound,
          						SUM(CLMECG) AS CLMECG, SUM(CLMDressing) AS CLMDressing, SUM(CLMNebulizer) AS CLMNebulizer, SUM(CLMID) AS CLMID,
          						SUM(CLMTSSTO) AS CLMTSSTO, SUM(CLMMinorSurgery) AS CLMMinorSurgery, SUM(CLMPOthers) AS CLMPOthers,
          						SUM(CLMBloodDrawing) AS CLMBloodDrawing, SUM(CLMBloodGlucose) AS CLMBloodGlucose, SUM(CLMUrine) AS CLMUrine,
          						SUM(CLMPapSmear) AS CLMPapSmear, SUM(CLMLOthers) AS CLMLOthers, SUM(CLMDTPPolio) AS CLMDTPPolio , SUM(CLMDTPPolioHIB) AS CLMDTPPolioHIB,
          						SUM(CLMHepatitisB) AS CLMHepatitisB, SUM(CLMMMR) AS CLMMMR, SUM(CLMIOthers) AS CLMIOthers, SUM(CLMDental) AS CLMDental,
          						SUM(CLMMaternity) AS CLMMaternity, SUM(CLMPRVCheckup) AS CLMPRVCheckup
          						From VIEW_CLMReport_MBM Where SUBSTRING(PAYCode,1,1) <> "Z"'
			       set @l_sqlstr40 = @l_sqlstr40 + @strAppend
			       exec  sp_executesql  @l_sqlstr40
		end

	if @strAppendCase = 41

		begin
         	      		set nocount on
			       declare @l_sqlstr41 as nvarchar(2000)

		                    set  @l_sqlstr41 =  N'SELECT PAYCode, PAYCompanyName from CISDB.dbo.PAY Where HLTCODE = "N"'
			       exec  sp_executesql  @l_sqlstr41
		end

	if @strAppendCase = 42

		begin
         	      		set nocount on
			       declare @l_sqlstr42 as nvarchar(2000)

		                    set  @l_sqlstr42 =  N'SELECT PAYCode, PAYCompanyName from CISDB.dbo.PAY Where HLTCODE = "S"'
			       exec  sp_executesql  @l_sqlstr42
		end

return


