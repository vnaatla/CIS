




CREATE procedure [dbo].[CISMBMTracking]

	@strAppend as nvarchar (2000), @strAppendCase as int

              
as
	if @strAppendCase = 1  	

		begin
         	      		set nocount on
			       declare @l_sqlstr as nvarchar(2000)
		                    declare @l_sqlgroupby as nvarchar(500)

		                    set  @l_sqlstr =  N'SELECT PAYCode, HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee, 
                          			                  SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0 '
		                    set @l_sqlgroupby = 'GROUP BY PAYCode, HLTCode, INSCode, AGECode'
			       set @l_sqlstr = @l_sqlstr +  @strAppend + @l_sqlgroupby
			       exec  sp_executesql  @l_sqlstr 

		end

	if @strAppendCase = 2

		begin
         			      set nocount on
				declare @l_sqlstr2 as nvarchar(2000)
				declare @l_sqlgroupby2 as nvarchar(500)
     
				set  @l_sqlstr2 =  N'SELECT PAYCode, HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalS, SUM(MBMMCOFee) AS MCOGross,  
                                                 			SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0 '
				set @l_sqlgroupby2 = 'GROUP BY PAYCode, HLTCode, INSCode, AGECode'
				set @l_sqlstr2 = @l_sqlstr2 +  @strAppend + @l_sqlgroupby2
				exec  sp_executesql  @l_sqlstr2 

		end

	if @strAppendCase = 3

		begin
			set nocount on
				declare @l_sqlstr3 as nvarchar(2000)
				declare @l_sqlgroupby3 as nvarchar(500)

				set  @l_sqlstr3 =  N'SELECT PAYCode, HLTCode, AGECode, COUNT(MBMNumber) AS TotalP,SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee, 
			          			 SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby3 = 'GROUP BY PAYCode, HLTCode, AGECode'
				set @l_sqlstr3 = @l_sqlstr3 +  @strAppend + @l_sqlgroupby3
				exec  sp_executesql  @l_sqlstr3 

		end

	if @strAppendCase = 4

		begin
			set nocount on
				declare @l_sqlstr4 as nvarchar(2000)
				declare @l_sqlgroupby4 as nvarchar(500)

				set  @l_sqlstr4 =  N'SELECT PAYCode, HLTCode, PrinAGECode, COUNT(MBMNumber) AS TotalS,SUM(MBMMCOFee) AS MCOGross, 
			          			 SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby4 = 'GROUP BY PAYCode, HLTCode, PrinAGECode'
				set @l_sqlstr4 = @l_sqlstr4 +  @strAppend + @l_sqlgroupby4
				exec  sp_executesql  @l_sqlstr4 

		end

	if @strAppendCase = 5

		begin
			set nocount on
				declare @l_sqlstr5 as nvarchar(2000)
				declare @l_sqlgroupby5 as nvarchar(500)

				set  @l_sqlstr5 =  N'SELECT PAYCode, HLTCode, AGECode, COUNT(MBMNumber) AS TotalP,SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee, 
          			 			SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby5 = 'GROUP BY PAYCode, HLTCode, AGECode'
				set @l_sqlstr5 = @l_sqlstr5 +  @strAppend + @l_sqlgroupby5
				exec  sp_executesql  @l_sqlstr5 

		end

	if @strAppendCase = 6

		begin
			set nocount on
				declare @l_sqlstr6 as nvarchar(2000)
				declare @l_sqlgroupby6 as nvarchar(500)

				set  @l_sqlstr6 =  N'SELECT PAYCode, HLTCode, AGECode, COUNT(MBMNumber) AS TotalS, SUM(MBMMCOFee) AS MCOGross, 
						SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby6 = 'GROUP BY PAYCode, HLTCode, AGECode'
				set @l_sqlstr6 = @l_sqlstr6 +  @strAppend + @l_sqlgroupby6
				exec  sp_executesql  @l_sqlstr6

		end

	if @strAppendCase = 7

		begin
			set nocount on
				declare @l_sqlstr7 as nvarchar(2000)
				declare @l_sqlgroupby7 as nvarchar(500)

				set  @l_sqlstr7 =  N'SELECT HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby7 = 'GROUP BY HLTCode, INSCode, AGECode'
				set @l_sqlstr7 = @l_sqlstr7 +  @strAppend + @l_sqlgroupby7
				exec  sp_executesql  @l_sqlstr7

		end

	if @strAppendCase = 8

		begin
			set nocount on
				declare @l_sqlstr8 as nvarchar(2000)
				declare @l_sqlgroupby8 as nvarchar(500)

				set  @l_sqlstr8 =  N'SELECT HLTCode, INSCode, PrinAGECode, COUNT(MBMNumber) AS TotalS, SUM(MBMMCOFee) AS MCOGross, 
						SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby8 = 'GROUP BY HLTCode, INSCode, PrinAGECode'
				set @l_sqlstr8 = @l_sqlstr8 +  @strAppend + @l_sqlgroupby8
				exec  sp_executesql  @l_sqlstr8

		end

	if @strAppendCase = 9

		begin
			set nocount on
				declare @l_sqlstr9 as nvarchar(2000)
				declare @l_sqlgroupby9 as nvarchar(500)

				set  @l_sqlstr9 =  N'SELECT HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee, 
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby9 = 'GROUP BY HLTCode, INSCode, AGECode'
				set @l_sqlstr9 = @l_sqlstr9 +  @strAppend + @l_sqlgroupby9
				exec  sp_executesql  @l_sqlstr9

		end

	if @strAppendCase = 10

		begin
			set nocount on
				declare @l_sqlstr10 as nvarchar(2000)
				declare @l_sqlgroupby10 as nvarchar(500)

				set  @l_sqlstr10 =  N'SELECT HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalS, SUM(MBMMCOFee) AS MCOGross, 
						SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby10 = 'GROUP BY HLTCode, INSCode, AGECode'
				set @l_sqlstr10 = @l_sqlstr10 +  @strAppend + @l_sqlgroupby10
				exec  sp_executesql  @l_sqlstr10

		end

	if @strAppendCase = 11

		begin
			set nocount on
				declare @l_sqlstr11 as nvarchar(2000)
				declare @l_sqlgroupby11 as nvarchar(500)

				set  @l_sqlstr11 =  N'SELECT PAYCode, HLTCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby11 = 'GROUP BY PAYCode, HLTCode, INSCode, AGECode'
				set @l_sqlstr11 = @l_sqlstr11 +  @strAppend + @l_sqlgroupby11
				exec  sp_executesql  @l_sqlstr11

		end

	if @strAppendCase = 12

		begin
			set nocount on
				declare @l_sqlstr12 as nvarchar(2000)
				declare @l_sqlgroupby12 as nvarchar(500)

				set  @l_sqlstr12 =  N'SELECT PAYCode, HLTCode, INSCode, PrinAGECode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby12 = 'GROUP BY PAYCode, HLTCode, INSCode, PrinAGECode'
				set @l_sqlstr12 = @l_sqlstr12 +  @strAppend + @l_sqlgroupby12
				exec  sp_executesql  @l_sqlstr12

		end

	if @strAppendCase = 13

		begin
			set nocount on
				declare @l_sqlstr13 as nvarchar(2000)
				declare @l_sqlgroupby13 as nvarchar(500)

				set  @l_sqlstr13 =  N'SELECT PAYCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby13 = 'GROUP BY PAYCode, INSCode, AGECode'
				set @l_sqlstr13 = @l_sqlstr13 +  @strAppend + @l_sqlgroupby13
				exec  sp_executesql  @l_sqlstr13

		end

	if @strAppendCase = 14

		begin
			set nocount on
				declare @l_sqlstr14 as nvarchar(2000)
				declare @l_sqlgroupby14 as nvarchar(500)

				set  @l_sqlstr14 =  N'SELECT PAYCode, INSCode, PrinAGECode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby14 = 'GROUP BY PAYCode, INSCode, PrinAGECode'
				set @l_sqlstr14 = @l_sqlstr14 +  @strAppend + @l_sqlgroupby14
				exec  sp_executesql  @l_sqlstr14

		end

	if @strAppendCase = 15

		begin
			set nocount on
				declare @l_sqlstr15 as nvarchar(2000)
				declare @l_sqlgroupby15 as nvarchar(500)

				set  @l_sqlstr15 =  N'SELECT PAYCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby15 = 'GROUP BY PAYCode, INSCode, AGECode'
				set @l_sqlstr15 = @l_sqlstr15 +  @strAppend + @l_sqlgroupby15
				exec  sp_executesql  @l_sqlstr15

		end

	if @strAppendCase = 16

		begin
			set nocount on
				declare @l_sqlstr16 as nvarchar(2000)
				declare @l_sqlgroupby16 as nvarchar(500)

				set  @l_sqlstr16 =  N'SELECT PAYCode, INSCode, AGECode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby16 = 'GROUP BY PAYCode, INSCode, AGECode'
				set @l_sqlstr16 = @l_sqlstr16 +  @strAppend + @l_sqlgroupby16
				exec  sp_executesql  @l_sqlstr16

		end

	if @strAppendCase = 17

		begin
			set nocount on
				declare @l_sqlstr17 as nvarchar(2000)
				declare @l_sqlgroupby17 as nvarchar(500)

				set  @l_sqlstr17 =  N'SELECT PAYCode, MBMAgentCode, HLTCode, INSCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby17 = 'GROUP BY PAYCode, MBMAgentCode, HLTCode, INSCode'
				set @l_sqlstr17 = @l_sqlstr17 +  @strAppend + @l_sqlgroupby17
				exec  sp_executesql  @l_sqlstr17

		end

	if @strAppendCase = 18

		begin
			set nocount on
				declare @l_sqlstr18 as nvarchar(2000)
				declare @l_sqlgroupby18 as nvarchar(500)

				set  @l_sqlstr18 =  N'SELECT PAYCode, MBMAgentCode, HLTCode, INSCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby18 = 'GROUP BY PAYCode, MBMAgentCode, HLTCode, INSCode'
				set @l_sqlstr18 = @l_sqlstr18 +  @strAppend + @l_sqlgroupby18
				exec  sp_executesql  @l_sqlstr18

		end

	if @strAppendCase = 19

		begin
			set nocount on
				declare @l_sqlstr19 as nvarchar(2000)
				declare @l_sqlgroupby19 as nvarchar(500)

				set  @l_sqlstr19 =  N'SELECT MBMAgentCode, PAYCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby19 = 'GROUP BY MBMAgentCode, PAYCode'
				set @l_sqlstr19 = @l_sqlstr19 +  @strAppend + @l_sqlgroupby19
				exec  sp_executesql  @l_sqlstr19

		end

	if @strAppendCase = 20

		begin
			set nocount on
				declare @l_sqlstr20 as nvarchar(2000)
				declare @l_sqlgroupby20 as nvarchar(500)

				set  @l_sqlstr20 =  N'SELECT MBMAgentCode, PAYCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby20 = 'GROUP BY MBMAgentCode, PAYCode'
				set @l_sqlstr20 = @l_sqlstr20 +  @strAppend + @l_sqlgroupby20
				exec  sp_executesql  @l_sqlstr20

		end

	if @strAppendCase = 21

		begin
			set nocount on
				declare @l_sqlstr21 as nvarchar(2000)
				declare @l_sqlgroupby21 as nvarchar(500)

				set  @l_sqlstr21 =  N'SELECT PAYCode, MBMAgentCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby21 = 'GROUP BY PAYCode, MBMAgentCode, INSCode, PLNCode'
				set @l_sqlstr21 = @l_sqlstr21 +  @strAppend + @l_sqlgroupby21
				exec  sp_executesql  @l_sqlstr21

		end

	if @strAppendCase = 22

		begin
			set nocount on
				declare @l_sqlstr22 as nvarchar(2000)
				declare @l_sqlgroupby22 as nvarchar(500)

				set  @l_sqlstr22 =  N'SELECT PAYCode, MBMAgentCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby22 = 'GROUP BY PAYCode, MBMAgentCode, INSCode, PLNCode'
				set @l_sqlstr22 = @l_sqlstr22 +  @strAppend + @l_sqlgroupby22
				exec  sp_executesql  @l_sqlstr22

		end

	if @strAppendCase = 23

		begin
			set nocount on
				declare @l_sqlstr23 as nvarchar(2000)
				declare @l_sqlgroupby23 as nvarchar(500)

				set  @l_sqlstr23 =  N'SELECT PAYCode, MBMAgentCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby23 = 'GROUP BY PAYCode, MBMAgentCode, INSCode, PLNCode'
				set @l_sqlstr23 = @l_sqlstr23 +  @strAppend + @l_sqlgroupby23
				exec  sp_executesql  @l_sqlstr23

		end

	if @strAppendCase = 24

		begin
			set nocount on
				declare @l_sqlstr24 as nvarchar(2000)
				declare @l_sqlgroupby24 as nvarchar(500)

				set  @l_sqlstr24 =  N'SELECT PAYCode, MBMAgentCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby24 = 'GROUP BY PAYCode, MBMAgentCode, INSCode, PLNCode'
				set @l_sqlstr24 = @l_sqlstr24 +  @strAppend + @l_sqlgroupby24
				exec  sp_executesql  @l_sqlstr24

		end

	if @strAppendCase = 25

		begin
			set nocount on
				declare @l_sqlstr25 as nvarchar(2000)
				declare @l_sqlgroupby25 as nvarchar(500)

				set  @l_sqlstr25 =  N'SELECT PAYCode, MBMAgentCode, RACCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross,
						SUM(MBMAccessFee) AS EAFee From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby25 = 'GROUP BY PAYCode, MBMAgentCode, RACCode'
				set @l_sqlstr25 = @l_sqlstr25 +  @strAppend + @l_sqlgroupby25
				exec  sp_executesql  @l_sqlstr25

		end

	if @strAppendCase = 26

		begin
			set nocount on
				declare @l_sqlstr26 as nvarchar(2000)
				declare @l_sqlgroupby26 as nvarchar(500)

				set  @l_sqlstr26 =  N'SELECT PAYCode, MBMAgentCode, RACCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby26 = 'GROUP BY PAYCode, MBMAgentCode, RACCode'
				set @l_sqlstr26 = @l_sqlstr26 +  @strAppend + @l_sqlgroupby26
				exec  sp_executesql  @l_sqlstr26

		end

	if @strAppendCase = 27

		begin
			set nocount on
				declare @l_sqlstr27 as nvarchar(2000)
				declare @l_sqlgroupby27 as nvarchar(500)

				set  @l_sqlstr27 =  N'SELECT PAYCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby27 = 'GROUP BY PAYCode, MBMAgentCode'
				set @l_sqlstr27 = @l_sqlstr27 +  @strAppend + @l_sqlgroupby27
				exec  sp_executesql  @l_sqlstr27

		end

	if @strAppendCase = 28

		begin
			set nocount on
				declare @l_sqlstr28 as nvarchar(2000)
				declare @l_sqlgroupby28 as nvarchar(500)

				set  @l_sqlstr28 =  N'SELECT PAYCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby28 = 'GROUP BY PAYCode, MBMAgentCode'
				set @l_sqlstr28 = @l_sqlstr28 +  @strAppend + @l_sqlgroupby28
				exec  sp_executesql  @l_sqlstr28

		end

	if @strAppendCase = 29

		begin
			set nocount on
				declare @l_sqlstr29 as nvarchar(2000)
				declare @l_sqlgroupby29 as nvarchar(500)

				set  @l_sqlstr29 =  N'SELECT TOP 20 PAYCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP, SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby29 = 'GROUP BY PAYCode, MBMAgentCode'
				set @l_sqlstr29 = @l_sqlstr29 +  @strAppend + @l_sqlgroupby29
				exec  sp_executesql  @l_sqlstr29

		end

	if @strAppendCase = 30

		begin
			set nocount on
				declare @l_sqlstr30 as nvarchar(2000)
				declare @l_sqlgroupby30 as nvarchar(500)

				set  @l_sqlstr30 =  N'SELECT TOP 20 PAYCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby30 = 'GROUP BY PAYCode, MBMAgentCode ORDER BY PremiumGross Desc'
				set @l_sqlstr30 = @l_sqlstr30 +  @strAppend + @l_sqlgroupby30
				exec  sp_executesql  @l_sqlstr30

		end

	if @strAppendCase = 31

		begin
			set nocount on
				declare @l_sqlstr31 as nvarchar(2000)

				set  @l_sqlstr31 =  N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, MBMMemberType, PLNCode, MBMCCoverID, HLTCode, MBMAgentCode, 
						MBMStatus, MBMAccessFee, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMPremium, DayInsured, PAYCode
						From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlstr31 = @l_sqlstr31 +  @strAppend
				exec  sp_executesql  @l_sqlstr31

		end

	if @strAppendCase = 32
	--Add Dependent MBMCEffDate,MBMCExpDate field for coverperson records- By CHEng 6Dec2016
		begin
			set nocount on
				declare @l_sqlstr32 as nvarchar(2000)

				set  @l_sqlstr32 =  N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, MBMMemberType, PLNCode, MBMCCoverID, HLTCode, MBMAgentCode, 
						MBMStatus, MBMAccessFee, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMPremium, DayInsured, PAYCode,
						RElCode, MBMCCoverID, MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr32 = @l_sqlstr32 +  @strAppend
				exec  sp_executesql  @l_sqlstr32

		end

	if @strAppendCase = 33

		begin
			set nocount on
				declare @l_sqlstr33 as nvarchar(2000)
				declare @l_sqlgroupby33 as nvarchar(500)

				set  @l_sqlstr33 =  N'SELECT MBMValueDate, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalP,
						SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby33 = 'GROUP BY MBMValueDate, MBMBordxDate, MBMBatchNo'
				set @l_sqlstr33 = @l_sqlstr33 +  @strAppend + @l_sqlgroupby33
				exec  sp_executesql  @l_sqlstr33

		end

	if @strAppendCase = 34

		begin
			set nocount on
				declare @l_sqlstr34 as nvarchar(2000)
				declare @l_sqlgroupby34 as nvarchar(500)

				set  @l_sqlstr34 =  N'SELECT MBMValueDate, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalS,
						SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby34 = 'GROUP BY MBMValueDate, MBMBordxDate, MBMBatchNo'
				set @l_sqlstr34 = @l_sqlstr34 +  @strAppend + @l_sqlgroupby34
				exec  sp_executesql  @l_sqlstr34

		end

	if @strAppendCase = 35

		begin
			set nocount on
				declare @l_sqlstr35 as nvarchar(2000)

				set  @l_sqlstr35 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMBordxDate, MBMBatchNo, MBMPremium,
						MBMMemberType, MBMStatus, MBMCCoverID, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMAccessFee
						From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlstr35 = @l_sqlstr35 +  @strAppend
				exec  sp_executesql  @l_sqlstr35

		end

	if @strAppendCase = 36
	--Add Dependent MBMCEffDate,MBMCExpDate field for coverperson records- By CHEng 6Dec2016
		begin
			set nocount on
				declare @l_sqlstr36 as nvarchar(2000)

				set  @l_sqlstr36 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMBordxDate, MBMBatchNo, MBMPremium, MBMCCoverID,
						MBMMemberType, MBMStatus, MBMCCoverID, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMAccessFee, MBMCEffDate,MBMCExpDate
						From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr36 = @l_sqlstr36 +  @strAppend
				exec  sp_executesql  @l_sqlstr36

		end

	if @strAppendCase = 37

		begin
			set nocount on
				declare @l_sqlstr37 as nvarchar(2000)
				declare @l_sqlgroupby37 as nvarchar(500)

				set  @l_sqlstr37 =  N'SELECT PAYCode, MBMBordxDate From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby37 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr37 = @l_sqlstr37 +  @strAppend + @l_sqlgroupby37
				exec  sp_executesql  @l_sqlstr37

		end

	if @strAppendCase = 38

		begin
			set nocount on
				declare @l_sqlstr38 as nvarchar(2000)
				declare @l_sqlgroupby38 as nvarchar(500)

				set  @l_sqlstr38 =  N'SELECT PAYCode, MBMBordxDate From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby38 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr38 = @l_sqlstr38 +  @strAppend + @l_sqlgroupby38
				exec  sp_executesql  @l_sqlstr38

		end

	if @strAppendCase = 39

		begin
			set nocount on
				declare @l_sqlstr39 as nvarchar(2000)
				declare @l_sqlgroupby39 as nvarchar(500)

				set  @l_sqlstr39 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalIC,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = "I" AND MBMAdultChild = "C"'
				set @l_sqlgroupby39 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr39 = @l_sqlstr39 +  @strAppend + @l_sqlgroupby39
				exec  sp_executesql  @l_sqlstr39

		end

	if @strAppendCase = 40

		begin
			set nocount on
				declare @l_sqlstr40 as nvarchar(2000)
				declare @l_sqlgroupby40 as nvarchar(500)

				set  @l_sqlstr40 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalIA,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = "I" AND MBMAdultChild = "A"'
				set @l_sqlgroupby40 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr40 = @l_sqlstr40 +  @strAppend + @l_sqlgroupby40
				exec  sp_executesql  @l_sqlstr40

		end

	if @strAppendCase = 41

		begin
			set nocount on
				declare @l_sqlstr41 as nvarchar(2000)
				declare @l_sqlgroupby41 as nvarchar(500)

				set  @l_sqlstr41 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalF,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = "F"'
				set @l_sqlgroupby41 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr41 = @l_sqlstr41 +  @strAppend + @l_sqlgroupby41
				exec  sp_executesql  @l_sqlstr41

		end

	if @strAppendCase = 42

		begin
			set nocount on
				declare @l_sqlstr42 as nvarchar(2000)
				declare @l_sqlgroupby42 as nvarchar(500)

				set  @l_sqlstr42 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalGC,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = "G" AND MBMAdultChild = "C"'
				set @l_sqlgroupby42 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr42 = @l_sqlstr42 +  @strAppend + @l_sqlgroupby42
				exec  sp_executesql  @l_sqlstr42

		end

	if @strAppendCase = 43

		begin
			set nocount on
				declare @l_sqlstr43 as nvarchar(2000)
				declare @l_sqlgroupby43 as nvarchar(500)

				set  @l_sqlstr43 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalGA,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = "G" AND MBMAdultChild = "A"'
				set @l_sqlgroupby43 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr43 = @l_sqlstr43 +  @strAppend + @l_sqlgroupby43
				exec  sp_executesql  @l_sqlstr43

		end

	if @strAppendCase = 44

		begin
			set nocount on
				declare @l_sqlstr44 as nvarchar(2000)
				declare @l_sqlgroupby44 as nvarchar(500)

				set  @l_sqlstr44 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalH,
						SUM(MBMMCOFEE) AS MCOFEE From VIEW_CISMBM_Search Where INSCode = "H"'
				set @l_sqlgroupby44 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr44 = @l_sqlstr44 +  @strAppend + @l_sqlgroupby44
				exec  sp_executesql  @l_sqlstr44

		end

	if @strAppendCase = 45

		begin
			set nocount on
				declare @l_sqlstr45 as nvarchar(2000)
				declare @l_sqlgroupby45 as nvarchar(500)

				set  @l_sqlstr45 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalFS,
						SUM(MBMMCOFee) AS MCOFee From View_CISMBM_Supplementary Where INSCode = "F"'
				set @l_sqlgroupby45 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr45 = @l_sqlstr45 +  @strAppend + @l_sqlgroupby45
				exec  sp_executesql  @l_sqlstr45

		end

	if @strAppendCase = 46

		begin
			set nocount on
				declare @l_sqlstr46 as nvarchar(2000)
				declare @l_sqlgroupby46 as nvarchar(500)


				set  @l_sqlstr46 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalHS,
						SUM(MBMMCOFee) AS MCOFee From View_CISMBM_Supplementary Where INSCode = "H"'
				set @l_sqlgroupby46 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr46 = @l_sqlstr46 +  @strAppend + @l_sqlgroupby46
				exec  sp_executesql  @l_sqlstr46

		end

	if @strAppendCase = 47

		begin
			set nocount on
				declare @l_sqlstr47 as nvarchar(2000)
				declare @l_sqlgroupby47 as nvarchar(500)

				set  @l_sqlstr47 =  N'SELECT PAYCode, MBMBordxDate From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby47 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr47 = @l_sqlstr47 +  @strAppend + @l_sqlgroupby47
				exec  sp_executesql  @l_sqlstr47

		end

	if @strAppendCase = 48

		begin
			set nocount on
				declare @l_sqlstr48 as nvarchar(2000)
				declare @l_sqlgroupby48 as nvarchar(500)

				set  @l_sqlstr48 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalIC,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = "I" AND MBMAdultChild = "C"'	
				set @l_sqlgroupby48 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr48 = @l_sqlstr48 +  @strAppend + @l_sqlgroupby48
				exec  sp_executesql  @l_sqlstr48

		end

	if @strAppendCase = 49

		begin
			set nocount on
				declare @l_sqlstr49 as nvarchar(2000)
				declare @l_sqlgroupby49 as nvarchar(500)

				set  @l_sqlstr49 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalIA,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = "I" AND MBMAdultChild = "A"'	
				set @l_sqlgroupby49 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr49 = @l_sqlstr49 +  @strAppend + @l_sqlgroupby49
				exec  sp_executesql  @l_sqlstr49

		end

	if @strAppendCase = 50

		begin
			set nocount on
				declare @l_sqlstr50 as nvarchar(2000)
				declare @l_sqlgroupby50 as nvarchar(500)

				set  @l_sqlstr50 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalF,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = "F"'
				set @l_sqlgroupby50 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr50 = @l_sqlstr50 +  @strAppend + @l_sqlgroupby50
				exec  sp_executesql  @l_sqlstr50

		end

	if @strAppendCase = 51

		begin
			set nocount on
				declare @l_sqlstr51 as nvarchar(2000)
				declare @l_sqlgroupby51 as nvarchar(500)

				set  @l_sqlstr51 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalFS,
						       SUM(MBMPremium) AS Premium From View_CISMBM_Supplementary Where INSCode = "F"'
				set @l_sqlgroupby51 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr51 = @l_sqlstr51 +  @strAppend + @l_sqlgroupby51
				exec  sp_executesql  @l_sqlstr51

		end


	if @strAppendCase = 52

		begin
			set nocount on
				declare @l_sqlstr52 as nvarchar(2000)
				declare @l_sqlgroupby52 as nvarchar(500)

				set  @l_sqlstr52 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalGC,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = "G" AND MBMAdultChild = "C"'
				set @l_sqlgroupby52 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr52 = @l_sqlstr52 +  @strAppend + @l_sqlgroupby52
				exec  sp_executesql  @l_sqlstr52

		end

	if @strAppendCase = 53

		begin
			set nocount on
				declare @l_sqlstr53 as nvarchar(2000)
				declare @l_sqlgroupby53 as nvarchar(500)

				set  @l_sqlstr53 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalGA,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = "G" AND MBMAdultChild = "A"'
				set @l_sqlgroupby53 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr53 = @l_sqlstr53 +  @strAppend + @l_sqlgroupby53
				exec  sp_executesql  @l_sqlstr53

		end

	if @strAppendCase = 54

		begin
			set nocount on
				declare @l_sqlstr54 as nvarchar(2000)
				declare @l_sqlgroupby54 as nvarchar(500)

				set  @l_sqlstr54 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalH,
						       SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where INSCode = "H"'
				set @l_sqlgroupby54 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr54 = @l_sqlstr54 +  @strAppend + @l_sqlgroupby54
				exec  sp_executesql  @l_sqlstr54

		end

	if @strAppendCase = 55

		begin
			set nocount on
				declare @l_sqlstr55 as nvarchar(2000)
				declare @l_sqlgroupby55 as nvarchar(500)

				set  @l_sqlstr55 =  N'SELECT PAYCode, MBMBordxDate, COUNT(MBMNumber) AS TotalHS,
						       SUM(MBMPremium) AS Premium From View_CISMBM_Supplementary Where INSCode = "H"'
				set @l_sqlgroupby55 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr55 = @l_sqlstr55 +  @strAppend + @l_sqlgroupby55
				exec  sp_executesql  @l_sqlstr55

		end

	if @strAppendCase = 56

		begin
			set nocount on
				declare @l_sqlstr56 as nvarchar(2000)
				declare @l_sqlgroupby56 as nvarchar(500)

				set  @l_sqlstr56 =  N'SELECT PAYCode, MBMBordxDate From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby56 = 'GROUP BY PAYCode, MBMBordxDate'
				set @l_sqlstr56 = @l_sqlstr56 +  @strAppend + @l_sqlgroupby56
				exec  sp_executesql  @l_sqlstr56

		end

	if @strAppendCase = 57

		begin
			set nocount on
				declare @l_sqlstr57 as nvarchar(2000)
				declare @l_sqlgroupby57 as nvarchar(500)

				set  @l_sqlstr57 =  N'SELECT PAYCode, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'	
				set @l_sqlgroupby57 = 'GROUP BY PAYCode, MBMBordxDate, MBMBatchNo'
				set @l_sqlstr57 = @l_sqlstr57 +  @strAppend + @l_sqlgroupby57
				exec  sp_executesql  @l_sqlstr57

		end

	if @strAppendCase = 58

		begin
			set nocount on
				declare @l_sqlstr58 as nvarchar(2000)
				declare @l_sqlgroupby58 as nvarchar(500)

				set  @l_sqlstr58 =  N'SELECT PAYCode, MBMBordxDate, MBMBatchNo, COUNT(MBMCNumber) AS TotalS,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby58 = 'GROUP BY PAYCode, MBMBordxDate, MBMBatchNo'
				set @l_sqlstr58 = @l_sqlstr58 +  @strAppend + @l_sqlgroupby58
				exec  sp_executesql  @l_sqlstr58

		end

	if @strAppendCase = 59

		begin
			set nocount on
				declare @l_sqlstr59 as nvarchar(2000)
				declare @l_sqlgroupby59 as nvarchar(500)

				set  @l_sqlstr59 =  N'SELECT HLTCode, BRNCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby59 = 'GROUP BY HLTCode, BRNCode'
				set @l_sqlstr59 = @l_sqlstr59 +  @strAppend + @l_sqlgroupby59
				exec  sp_executesql  @l_sqlstr59

		end

	if @strAppendCase = 60

		begin
			set nocount on
				declare @l_sqlstr60 as nvarchar(2000)
				declare @l_sqlgroupby60 as nvarchar(500)

				set  @l_sqlstr60 =  N'SELECT HLTCode, BRNCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby60 = 'GROUP BY HLTCode, BRNCode'
				set @l_sqlstr60 = @l_sqlstr60 +  @strAppend + @l_sqlgroupby60
				exec  sp_executesql  @l_sqlstr60

		end

	if @strAppendCase = 61

		begin
			set nocount on
				declare @l_sqlstr61 as nvarchar(2000)
				declare @l_sqlgroupby61 as nvarchar(500)

				set  @l_sqlstr61 =  N'SELECT BRNCode, MBMAgentCode From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby61 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr61 = @l_sqlstr61 +  @strAppend + @l_sqlgroupby61
				exec  sp_executesql  @l_sqlstr61

		end

	if @strAppendCase = 62

		begin
			set nocount on
				declare @l_sqlstr62 as nvarchar(2000)
				declare @l_sqlgroupby62 as nvarchar(500)

				set  @l_sqlstr62 =  N'SELECT BRNCode, MBMAgentCode From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby62 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr62 = @l_sqlstr62 +  @strAppend + @l_sqlgroupby62
				exec  sp_executesql  @l_sqlstr62

		end

	if @strAppendCase = 63

		begin
			set nocount on
				declare @l_sqlstr63 as nvarchar(2000)
				declare @l_sqlgroupby63 as nvarchar(500)

				set  @l_sqlstr63 =  N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0
						       AND (BRNCode IS NULL OR BRNCode = '') AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)'	
				set @l_sqlgroupby63 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr63 = @l_sqlstr63 +  @strAppend + @l_sqlgroupby63
				exec  sp_executesql  @l_sqlstr63

		end

	if @strAppendCase = 64

		begin
			set nocount on
				declare @l_sqlstr64 as nvarchar(2000)
				declare @l_sqlgroupby64 as nvarchar(500)

				set  @l_sqlstr64 =  N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0
						       AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)'
				set @l_sqlgroupby64 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr64 = @l_sqlstr64 +  @strAppend + @l_sqlgroupby64
				exec  sp_executesql  @l_sqlstr64

		end

	if @strAppendCase = 65

		begin
			set nocount on
				declare @l_sqlstr65 as nvarchar(2000)
				declare @l_sqlgroupby65 as nvarchar(500)

				set  @l_sqlstr65 =  N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0
						       AND (BRNCode = '' OR BRNCode IS NULL)'
				set @l_sqlgroupby65 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr65 = @l_sqlstr65 + @strAppend + @l_sqlgroupby65
				exec  sp_executesql  @l_sqlstr65

		end

	if @strAppendCase = 66

		begin
			set nocount on
				declare @l_sqlstr66 as nvarchar(2000)
				declare @l_sqlgroupby66 as nvarchar(500)

				set  @l_sqlstr66 =  N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby66 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr66 = @l_sqlstr66 +  @strAppend + @l_sqlgroupby66
				exec  sp_executesql  @l_sqlstr66

		end


	if @strAppendCase = 67

		begin
			set nocount on
				declare @l_sqlstr67 as nvarchar(2000)
				declare @l_sqlgroupby67 as nvarchar(500)

				set  @l_sqlstr67 =  N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0
						       AND (BRNCode IS NULL OR BRNCode = '') AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)'	
				set @l_sqlgroupby67 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr67 = @l_sqlstr67 +  @strAppend + @l_sqlgroupby67
				exec  sp_executesql  @l_sqlstr67

		end

	if @strAppendCase = 68

		begin
			set nocount on
				declare @l_sqlstr68 as nvarchar(2000)
				declare @l_sqlgroupby68 as nvarchar(500)

				set  @l_sqlstr68 =  N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0
						       AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)'
				set @l_sqlgroupby68 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr68 = @l_sqlstr68 +  @strAppend + @l_sqlgroupby68
				exec  sp_executesql  @l_sqlstr68

		end

	if @strAppendCase = 69

		begin
			set nocount on
				declare @l_sqlstr69 as nvarchar(2000)
				declare @l_sqlgroupby69 as nvarchar(500)

				set  @l_sqlstr69 =  N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0
						       AND (BRNCode IS NULL OR BRNCode = '')'
				set @l_sqlgroupby69 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr69 = @l_sqlstr69 +  @strAppend + @l_sqlgroupby69
				exec  sp_executesql  @l_sqlstr69

		end

	if @strAppendCase = 70

		begin
			set nocount on
				declare @l_sqlstr70 as nvarchar(2000)
				declare @l_sqlgroupby70 as nvarchar(500)

				set  @l_sqlstr70 =  N'SELECT BRNCode, MBMAgentCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby70 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr70 = @l_sqlstr70 +  @strAppend + @l_sqlgroupby70
				exec  sp_executesql  @l_sqlstr70

		end

	if @strAppendCase = 71

		begin
			set nocount on
				declare @l_sqlstr71 as nvarchar(2000)
				declare @l_sqlgroupby71 as nvarchar(500)

				set  @l_sqlstr71 =  N'SELECT BRNCode, MBMAgentCode, COUNT(CASNumber) AS TotalClaim,
						       SUM(CLMTotalApproved) AS AppAmt From VIEW_MBMClaim_Search Where 0 = 0
						       AND (BRNCode IS NULL OR BRNCode = '') AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)'	
				set @l_sqlgroupby71 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr71 = @l_sqlstr71 +  @strAppend + @l_sqlgroupby71
				exec  sp_executesql  @l_sqlstr71

		end

	if @strAppendCase = 72

		begin
			set nocount on
				declare @l_sqlstr72 as nvarchar(2000)
				declare @l_sqlgroupby72 as nvarchar(500)

				set  @l_sqlstr72 =  N'SELECT BRNCode, MBMAgentCode, COUNT(CASNumber) AS TotalClaim,
						       SUM(CLMTotalApproved) AS AppAmt From VIEW_MBMClaim_Search Where 0 = 0
						       AND (MBMAgentCode = '' OR MBMAgentCode IS NULL)'
				set @l_sqlgroupby72 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr72 = @l_sqlstr72 +  @strAppend + @l_sqlgroupby72
				exec  sp_executesql  @l_sqlstr72

		end

	if @strAppendCase = 73

		begin
			set nocount on
				declare @l_sqlstr73 as nvarchar(2000)
				declare @l_sqlgroupby73 as nvarchar(500)

				set  @l_sqlstr73 =  N'SELECT BRNCode, MBMAgentCode, COUNT(CASNumber) AS TotalClaim,
						       SUM(CLMTotalApproved) AS AppAmt From VIEW_MBMClaim_Search Where 0 = 0
						       AND (BRNCode IS NULL OR BRNCode = '')'
				set @l_sqlgroupby73 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr73 = @l_sqlstr73 +  @strAppend + @l_sqlgroupby73
				exec  sp_executesql  @l_sqlstr73

		end

	if @strAppendCase = 74

		begin
			set nocount on
				declare @l_sqlstr74 as nvarchar(2000)
				declare @l_sqlgroupby74 as nvarchar(500)

				set  @l_sqlstr74 =  N'SELECT BRNCode, MBMAgentCode, COUNT(CASNumber) AS TotalClaim,
						       SUM(CLMTotalApproved) AS AppAmt From VIEW_MBMClaim_Search Where 0 = 0'
				set @l_sqlgroupby74 = 'GROUP BY BRNCode, MBMAgentCode'
				set @l_sqlstr74 = @l_sqlstr74 +  @strAppend + @l_sqlgroupby74
				exec  sp_executesql  @l_sqlstr74

		end

	if @strAppendCase = 75

		begin
			set nocount on
				declare @l_sqlstr75 as nvarchar(2000)
				declare @l_sqlgroupby75 as nvarchar(500)

				set  @l_sqlstr75 =  N'SELECT HLTCode, PAYCode, BRNCode, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby75 = 'GROUP BY HLTCode, PAYCode, BRNCode'
				set @l_sqlstr75 = @l_sqlstr75 +  @strAppend + @l_sqlgroupby75
				exec  sp_executesql  @l_sqlstr75

		end

	if @strAppendCase = 76

		begin
			set nocount on
				declare @l_sqlstr76 as nvarchar(2000)
				declare @l_sqlgroupby76 as nvarchar(500)

				set  @l_sqlstr76 =  N'SELECT HLTCode, PAYCode, BRNCode, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby76 = 'GROUP BY HLTCode, PAYCode, BRNCode'
				set @l_sqlstr76 = @l_sqlstr76 +  @strAppend + @l_sqlgroupby76
				exec  sp_executesql  @l_sqlstr76

		end

	if @strAppendCase = 77

		begin
			set nocount on
				declare @l_sqlstr77 as nvarchar(2000)

				set  @l_sqlstr77 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMDOB, MBMBordxDate,  MBMValueDate, MBMInvoiceDate, 
						      MBMMemberType, MBMStatus, MBMPayorEffDate, MBMPayorExpDate, MBMRenewFlag, PAYCode, DayInsured,
						      AgeDays, DelayBordx, DelayReg, DelayInvD, DelayInvE, DelayInvBordx, RELCode, MBMCCoverID From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlstr77 = @l_sqlstr77 +  @strAppend
				exec  sp_executesql  @l_sqlstr77

		end

	if @strAppendCase = 78
	--Add Dependent MBMCEffDate,MBMCExpDate field for coverperson records- By CHEng 6Dec2016
		begin
			set nocount on
				declare @l_sqlstr78 as nvarchar(2000)

				set  @l_sqlstr78 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMDOB, MBMBordxDate,  MBMValueDate, MBMInvoiceDate, 
						      MBMMemberType, MBMStatus, MBMPayorEffDate, MBMPayorExpDate, MBMRenewFlag, PAYCode, DayInsured,	
						      AgeDays, DelayBordx, DelayReg, DelayInvD, DelayInvE, DelayInvBordx, RELCode, MBMCCoverID, MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr78 = @l_sqlstr78 +  @strAppend
				exec  sp_executesql  @l_sqlstr78

		end

	if @strAppendCase = 79

		begin
			set nocount on
				declare @l_sqlstr79 as nvarchar(2000)

				set  @l_sqlstr79 =  N'SELECT MBMNumber, MBMPolicyNo, MBMName, MBMPayorEffDate, MBMPayorExpDate, DayInsured,
						      INSCode, PLNCode From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlstr79 = @l_sqlstr79 +  @strAppend
				exec  sp_executesql  @l_sqlstr79

		end

	if @strAppendCase = 80

		begin
			set nocount on
				declare @l_sqlstr80 as nvarchar(2000)
				declare @l_sqlgroupby80 as nvarchar(500)

				set  @l_sqlstr80 =  N'SELECT COUNT(MBMNumber) AS TotalP, DayInsured From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby80 = 'GROUP BY DayInsured'
				set @l_sqlstr80 = @l_sqlstr80 +  @strAppend + @l_sqlgroupby80
				exec  sp_executesql  @l_sqlstr80

		end

	if @strAppendCase = 81

		begin
			set nocount on
				declare @l_sqlstr81 as nvarchar(2000)
				declare @l_sqlgroupby81 as nvarchar(500)

				set  @l_sqlstr81 =  N'SELECT DayInsured, COUNT(MBMNumber) AS TotalS From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby81 = 'GROUP BY DayInsured'
				set @l_sqlstr81 = @l_sqlstr81 +  @strAppend + @l_sqlgroupby81
				exec  sp_executesql  @l_sqlstr81

		end

	if @strAppendCase = 82

		begin
			set nocount on
				declare @l_sqlstr82 as nvarchar(2000)
				declare @l_sqlgroupby82 as nvarchar(500)

				set  @l_sqlstr82 =  N'SELECT HLTCode, INSCode, MBMSex, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby82 = 'GROUP BY HLTCode, INSCode, MBMSex'
				set @l_sqlstr82 = @l_sqlstr82 +  @strAppend + @l_sqlgroupby82
				exec  sp_executesql  @l_sqlstr82

		end

	if @strAppendCase = 83

		begin
			set nocount on
				declare @l_sqlstr83 as nvarchar(2000)
				declare @l_sqlgroupby83 as nvarchar(500)

				set  @l_sqlstr83 =  N'SELECT HLTCode, INSCode, PrinSex, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						       From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby83 = 'GROUP BY HLTCode, INSCode, PrinSex'
				set @l_sqlstr83 = @l_sqlstr83 +  @strAppend + @l_sqlgroupby83
				exec  sp_executesql  @l_sqlstr83

		end

	if @strAppendCase = 84

		begin
			set nocount on
				declare @l_sqlstr84 as nvarchar(2000)
				declare @l_sqlgroupby84 as nvarchar(500)

				set  @l_sqlstr84 =  N'SELECT HLTCode, INSCode, MBMSex, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						       From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby84 = 'GROUP BY HLTCode, INSCode, MBMSex'
				set @l_sqlstr84 = @l_sqlstr84 +  @strAppend + @l_sqlgroupby84
				exec  sp_executesql  @l_sqlstr84

		end

	if @strAppendCase = 85

		begin
			set nocount on
				declare @l_sqlstr85 as nvarchar(2000)
				declare @l_sqlgroupby85 as nvarchar(500)

				set  @l_sqlstr85 =  N'SELECT PAYCode, HLTCode, INSCode, MBMSex, COUNT(MBMNumber) AS TotalP,
						       SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						       SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby85 = 'GROUP BY PAYCode, HLTCode, INSCode, MBMSex'
				set @l_sqlstr85 = @l_sqlstr85 +  @strAppend + @l_sqlgroupby85
				exec  sp_executesql  @l_sqlstr85

		end

	if @strAppendCase = 86

		begin
			set nocount on
				declare @l_sqlstr86 as nvarchar(2000)
				declare @l_sqlgroupby86 as nvarchar(500)

				set  @l_sqlstr86 =  N'SELECT PAYCode, HLTCode, INSCode, PrinSex, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						       From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby86 = 'GROUP BY PAYCode, HLTCode, INSCode, PrinSex'
				set @l_sqlstr86 = @l_sqlstr86 +  @strAppend + @l_sqlgroupby86
				exec  sp_executesql  @l_sqlstr86

		end

	if @strAppendCase = 87

		begin
			set nocount on
				declare @l_sqlstr87 as nvarchar(2000)
				declare @l_sqlgroupby87 as nvarchar(500)

				set  @l_sqlstr87 =  N'SELECT PAYCode, HLTCode, INSCode, MBMSex, COUNT(MBMNumber) AS TotalS,
						       SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						       From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby87 = 'GROUP BY PAYCode, HLTCode, INSCode, MBMSex'
				set @l_sqlstr87 = @l_sqlstr87 +  @strAppend + @l_sqlgroupby87
				exec  sp_executesql  @l_sqlstr87

		end

	if @strAppendCase = 88

		begin
			set nocount on
				declare @l_sqlstr88 as nvarchar(2000)

				set  @l_sqlstr88 =  N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, AGECode, MBMSex, MBMExpiredStatus,
						     PLNCode, MBMGroupCompany, MBMMemberType, case When cast(MBMPayorExpDate as date)<cast(getdate() as date) then ''E'' else MBMStatus end as MBMStatus, MBMDataStatus, RElCode, MBMCCoverID,
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee, MBMValueDate,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMIcBcPp, 
						     MBMEmployeeNo, MBMDepartment, MBMCancelDate,Subsidiary,GRPCoPaycode,mbmremarks,REMARK,MBMCostCenter,MBMDepartment,MBMBranch,MBMDivision,Principalname 
							 ,MBMInvoiceDate,MBMInvoiceNo
							 From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlstr88 = @l_sqlstr88 +  @strAppend
				exec  sp_executesql  @l_sqlstr88

		end

	if @strAppendCase = 89
	-- Add Dependent MBMCEffDate,MBMCExpDate field for coverperson records- By CHEng 6Dec2016
		begin
			set nocount on
				declare @l_sqlstr89 as nvarchar(2000)

				set  @l_sqlstr89 =  N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, AGECode, MBMSex, MBMExpiredStatus,
						     PLNCode, MBMGroupCompany, MBMMemberType, case When cast(MBMPayorExpDate as date)<cast(getdate() as date) then ''E'' else MBMStatus end as MBMStatus, MBMDataStatus, 
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee, MBMValueDate,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMIcBcPp, 
						     MBMEmployeeNo, MBMDepartment, RELCode, MBMCCoverID, MBMCancelDate,GRPCoPaycode, MBMCEffDate,MBMCExpDate,Subsidiary,mbmremarks,REMARK,MBMCostCenter,MBMDepartment,
							 MBMBranch,MBMDivision,Principalname ,MBMInvoiceDate,MBMInvoiceNo
							 From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr89 = @l_sqlstr89 +  @strAppend
				exec  sp_executesql  @l_sqlstr89

		end

	if @strAppendCase = 90
	-- Add Dependent MBMCEffDate,MBMCExpDate field for coverperson records By CHEng 6Dec2016
		begin
			set nocount on
				declare @l_sqlstr90 as nvarchar(2000)

				set  @l_sqlstr90 =  N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, AGECode, MBMSex, MBMExpiredStatus,
						     PLNCode, MBMGroupCompany, MBMMemberType, MBMStatus, MBMDataStatus,
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee, MBMValueDate,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMIcBcPp, 
						     MBMEmployeeNo, MBMDepartment, RELCode, MBMCCoverID, MBMCRegDate, MBMCancelDate, MBMBordxDate,MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr90 = @l_sqlstr90 +  @strAppend
				exec  sp_executesql  @l_sqlstr90

		end


	if @strAppendCase = 91

		begin
			set nocount on
				declare @l_sqlstr91 as nvarchar(2000)

				set  @l_sqlstr91 =  N'Select MBMNumber, MBMPolicyNo, MBMName, INSCode, AGECode, MBMSex, MBMEmployeeNo,
						     PLNCode, MBMGroupCompany, MBMMemberType, MBMStatus, RELCode, MBMCCoverID,
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee, MBMCancelDate,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMDepartment
						     From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlstr91 = @l_sqlstr91 +  @strAppend
				exec  sp_executesql  @l_sqlstr91

		end

	if @strAppendCase = 92
	-- Add Dependent MBMCEffDate,MBMCExpDate field for coverperson records By CHEng 6Dec2016
		begin
			set nocount on
				declare @l_sqlstr92 as nvarchar(2000)

				set  @l_sqlstr92 =  N'Select MBMNumber, MBMPolicyNo, MBMName,  MBMCName, INSCode, AGECode, MBMSex, MBMEmployeeNo,
						     PLNCode, MBMGroupCompany, MBMMemberType, MBMStatus, 
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMDepartment,
						     MBMCSex, MBMCAgeband, MBMCDOB, RELCode, MBMCCoverID, MBMCancelDate, MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary2 Where 0 = 0'
				set @l_sqlstr92 = @l_sqlstr92 +  @strAppend
				exec  sp_executesql  @l_sqlstr92

		end

	if @strAppendCase = 93
	-- Add Dependent MBMCEffDate,MBMCExpDate field for coverperson records By CHEng 6Dec2016
		begin
			set nocount on
				declare @l_sqlstr93 as nvarchar(2000)

				set  @l_sqlstr93 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMCName, INSCode, AGECode, MBMSex, MBMEmployeeNo,
						     PLNCode, MBMGroupCompany, MBMMemberType, MBMStatus, RELCode, MBMCCoverID,
						     MBMDOB, MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMBasicMCO, MBMAccessFee,
						     MBMPremium, MBMBasicPrem, STACode, MBMBordxDate, RACCode, DayInsured, MBMRenewFlag, PAYCode, MBMDepartment,
						     MBMCSex, MBMCAgeband, MBMCDOB, MBMCRegDate, MBMCancelDate, MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary2 Where 0 = 0'
				set @l_sqlstr93 = @l_sqlstr93 +  @strAppend
				exec  sp_executesql  @l_sqlstr93

		end

	if @strAppendCase = 94

		begin
			set nocount on
				declare @l_sqlstr94 as nvarchar(2000)

				set  @l_sqlstr94 =  N'Select MBMNumber, MBMName, RElCode, MBMCCoverID, MBMStatus,
						     MBMMemberType, MBMPremium, MBMBasicMCO, MBMMCOFee, MBMBasicPrem,
						     MBMPayorMCONet, MBMPayorMCOGross, MBMPayorPremNet, MBMPayorPremGross
						     From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlstr94 = @l_sqlstr94 +  @strAppend
				exec  sp_executesql  @l_sqlstr94

		end

	if @strAppendCase = 95

		begin
			set nocount on
				declare @l_sqlstr95 as nvarchar(2000)

				set  @l_sqlstr95 =  N'Select MBMNumber, MBMName, RElCode, MBMCCoverID, MBMStatus,
						     MBMMemberType, MBMPremium, MBMBasicMCO, MBMMCOFee, MBMBasicPrem,
						     MBMPayorMCONet, MBMPayorMCOGross, MBMPayorPremNet, MBMPayorPremGross
						     From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr95 = @l_sqlstr95 +  @strAppend
				exec  sp_executesql  @l_sqlstr95

		end

	if @strAppendCase = 96

		begin
			set nocount on
				declare @l_sqlstr96 as nvarchar(2000)

				set  @l_sqlstr96 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMAddress1, CTYCode,
						     MBMGroupCompany, MBMMemberType, MBMStatus, MBMCCoverID, MBMPostCode,
						     STACode, MBMRenewFlag, PAYCode, MBMAddress2, MBMAddress3
						     From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlstr96 = @l_sqlstr96 +  @strAppend
				exec  sp_executesql  @l_sqlstr96

		end

	if @strAppendCase = 97

		begin
			set nocount on
				declare @l_sqlstr97 as nvarchar(2000)

				set  @l_sqlstr97 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMAddress1, CTYCode,
						     MBMGroupCompany, MBMMemberType, MBMStatus, MBMCCoverID, MBMPostCode,
						     STACode, MBMRenewFlag, PAYCode, MBMAddress2, MBMAddress3, MBMCCoverID, MBMCRegDate
						     From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr97 = @l_sqlstr97 +  @strAppend
				exec  sp_executesql  @l_sqlstr97

		end

	if @strAppendCase = 98

		begin
			set nocount on
				declare @l_sqlstr98 as nvarchar(2000)

				set  @l_sqlstr98 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMAddress1, CTYCode,
						     MBMGroupCompany, MBMMemberType, MBMStatus, MBMCCoverID, MBMPostCode,
						     STACode, MBMRenewFlag, PAYCode, MBMAddress2, MBMAddress3, MBMCCoverID
						     From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr98 = @l_sqlstr98 +  @strAppend
				exec  sp_executesql  @l_sqlstr98

		end

	if @strAppendCase = 99

		begin
			set nocount on
				declare @l_sqlstr99 as nvarchar(2000)
				declare @l_sqlgroupby99 as nvarchar(500)

				set  @l_sqlstr99 =  N'SELECT PAYCode, MBMExpiredStatus, INSCode, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby99 = 'GROUP BY PAYCode, MBMExpiredStatus, INSCode'
				set @l_sqlstr99 = @l_sqlstr99 +  @strAppend + @l_sqlgroupby99
				exec  sp_executesql  @l_sqlstr99

		end

	if @strAppendCase = 100

		begin
			set nocount on
				declare @l_sqlstr100 as nvarchar(2000)
				declare @l_sqlgroupby100 as nvarchar(500)

				set  @l_sqlstr100 =  N'SELECT PAYCode, MBMExpiredStatus, INSCode, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby100 = 'GROUP BY PAYCode, MBMExpiredStatus, INSCode'
				set @l_sqlstr100 = @l_sqlstr100 +  @strAppend + @l_sqlgroupby100
				exec  sp_executesql  @l_sqlstr100

		end

	if @strAppendCase = 101

		begin
			set nocount on
				declare @l_sqlstr101 as nvarchar(2000)

				set  @l_sqlstr101 =  N'Select MBMNumber, MBMName, MBMIcBcPp, MBMInstallment,
						     MBMMemberType, MBMStatus, MBMCCoverID, MBMTakeover,
						     MBMPayorEffDate, MBMPayorExpDate, MBMPolicyYear,
						     MBMRenewFlag, MBMFirstJoinedDate From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlstr101 = @l_sqlstr101 +  @strAppend
				exec  sp_executesql  @l_sqlstr101

		end

	if @strAppendCase = 102
	--Add Dependent MBMCEffDate,MBMCExpDate field for coverperson records- By CHEng 6Dec2016
		begin
			set nocount on
				declare @l_sqlstr102 as nvarchar(2000)

				set  @l_sqlstr102 =  N'Select MBMNumber, MBMName, MBMIcBcPp, MBMInstallment,
						     MBMMemberType, MBMStatus, MBMCCoverID, MBMTakeover,
						     MBMPayorEffDate, MBMPayorExpDate, MBMPolicyYear,
						     MBMRenewFlag, MBMFirstJoinedDate, MBMCCoverID, MBMCEffDate,MBMCExpDate From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr102 = @l_sqlstr102 +  @strAppend
				exec  sp_executesql  @l_sqlstr102

		end

	if @strAppendCase = 103

		begin
			set nocount on
				declare @l_sqlstr103 as nvarchar(2000)
				declare @l_sqlgroupby103 as nvarchar(500)

				set  @l_sqlstr103 =  N'SELECT MBMPolicyYear, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby103 = 'GROUP BY MBMPolicyYear'
				set @l_sqlstr103 = @l_sqlstr103 +  @strAppend + @l_sqlgroupby103
				exec  sp_executesql  @l_sqlstr103

		end

	if @strAppendCase = 104

		begin
			set nocount on
				declare @l_sqlstr104 as nvarchar(2000)
				declare @l_sqlgroupby104 as nvarchar(500)

				set  @l_sqlstr104 =  N'SELECT MBMPolicyYear, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby104 = 'GROUP BY MBMPolicyYear'
				set @l_sqlstr104 = @l_sqlstr104 +  @strAppend + @l_sqlgroupby104
				exec  sp_executesql  @l_sqlstr104

		end

	if @strAppendCase = 105

		begin
			set nocount on
				declare @l_sqlstr105 as nvarchar(2000)
				declare @l_sqlgroupby105 as nvarchar(500)

				set  @l_sqlstr105 =  N'SELECT PAYCode, MBMInstallment, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby105 = 'GROUP BY PAYCode, MBMInstallment'
				set @l_sqlstr105 = @l_sqlstr105 +  @strAppend + @l_sqlgroupby105
				exec  sp_executesql  @l_sqlstr105

		end

	if @strAppendCase = 106

		begin
			set nocount on
				declare @l_sqlstr106 as nvarchar(2000)
				declare @l_sqlgroupby106 as nvarchar(500)

				set  @l_sqlstr106 =  N'SELECT PAYCode, MBMInstallment, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby106 = 'GROUP BY PAYCode, MBMInstallment'
				set @l_sqlstr106 = @l_sqlstr106 +  @strAppend + @l_sqlgroupby106
				exec  sp_executesql  @l_sqlstr106

		end

	if @strAppendCase = 107

		begin
			set nocount on
				declare @l_sqlstr107 as nvarchar(2000)
				declare @l_sqlgroupby107 as nvarchar(500)

				set  @l_sqlstr107 =  N'SELECT PAYCode, MBMRenewFlag, INSCode, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby107 = 'GROUP BY PAYCode, MBMRenewFlag, INSCode'
				set @l_sqlstr107 = @l_sqlstr107 +  @strAppend + @l_sqlgroupby107
				exec  sp_executesql  @l_sqlstr107

		end

	if @strAppendCase = 108

		begin
			set nocount on
				declare @l_sqlstr108 as nvarchar(2000)
				declare @l_sqlgroupby108 as nvarchar(500)

				set  @l_sqlstr108 =  N'SELECT PAYCode, MBMRenewFlag, INSCode, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby108 = 'GROUP BY PAYCode, MBMRenewFlag, INSCode'
				set @l_sqlstr108 = @l_sqlstr108 +  @strAppend + @l_sqlgroupby108
				exec  sp_executesql  @l_sqlstr108

		end

	if @strAppendCase = 109

		begin
			set nocount on
				declare @l_sqlstr109 as nvarchar(2000)
				declare @l_sqlgroupby109 as nvarchar(500)

				set  @l_sqlstr109 =  N'SELECT PAYCode, MBMTakeover, INSCode, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby109 = 'GROUP BY PAYCode, MBMTakeover, INSCode'
				set @l_sqlstr109 = @l_sqlstr109 +  @strAppend + @l_sqlgroupby109
				exec  sp_executesql  @l_sqlstr109

		end

	if @strAppendCase = 110

		begin
			set nocount on
				declare @l_sqlstr110 as nvarchar(2000)
				declare @l_sqlgroupby110 as nvarchar(500)

				set  @l_sqlstr110 =  N'SELECT PAYCode, MBMTakeover, INSCode, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby110 = 'GROUP BY PAYCode, MBMTakeover, INSCode'
				set @l_sqlstr110 = @l_sqlstr110 +  @strAppend + @l_sqlgroupby110
				exec  sp_executesql  @l_sqlstr110

		end

	if @strAppendCase = 111

		begin
			set nocount on
				declare @l_sqlstr111 as nvarchar(2000)
				declare @l_sqlgroupby111 as nvarchar(500)

				set  @l_sqlstr111 =  N'SELECT HLTCode, INSCode, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby111 = 'GROUP BY HLTCode, INSCode'
				set @l_sqlstr111 = @l_sqlstr111 +  @strAppend + @l_sqlgroupby111
				exec  sp_executesql  @l_sqlstr111

		end

	if @strAppendCase = 112

		begin
			set nocount on
				declare @l_sqlstr112 as nvarchar(2000)
				declare @l_sqlgroupby112 as nvarchar(500)

				set  @l_sqlstr112 =  N'SELECT HLTCode, INSCode, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby112 = 'GROUP BY HLTCode, INSCode'
				set @l_sqlstr112 = @l_sqlstr112 +  @strAppend + @l_sqlgroupby112
				exec  sp_executesql  @l_sqlstr112

		end

	if @strAppendCase = 113

		begin
			set nocount on
				declare @l_sqlstr113 as nvarchar(2000)
				declare @l_sqlgroupby113 as nvarchar(500)

				set  @l_sqlstr113 =  N'SELECT PAYCode, HLTCode, INSCode, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						     SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby113 = 'GROUP BY PAYCode, HLTCode, INSCode'
				set @l_sqlstr113 = @l_sqlstr113 +  @strAppend + @l_sqlgroupby113
				exec  sp_executesql  @l_sqlstr113

		end

	if @strAppendCase = 114

		begin
			set nocount on
				declare @l_sqlstr114 as nvarchar(2000)
				declare @l_sqlgroupby114 as nvarchar(500)

				set  @l_sqlstr114 =  N'SELECT PAYCode, HLTCode, INSCode, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby114 = 'GROUP BY PAYCode, HLTCode, INSCode'
				set @l_sqlstr114 = @l_sqlstr114 +  @strAppend + @l_sqlgroupby114
				exec  sp_executesql  @l_sqlstr114

		end

	if @strAppendCase = 115

		begin
			set nocount on
				declare @l_sqlstr115 as nvarchar(2000)
				declare @l_sqlgroupby115 as nvarchar(500)

				set  @l_sqlstr115 =  N'SELECT MONTH(MBMPayorEffDate) AS Month,
						     YEAR(MBMPayorEffDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby115 = 'GROUP BY MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)'
				set @l_sqlstr115 = @l_sqlstr115 +  @strAppend + @l_sqlgroupby115
				exec  sp_executesql  @l_sqlstr115

		end

	if @strAppendCase = 116

		begin
			set nocount on
				declare @l_sqlstr116 as nvarchar(2000)
				declare @l_sqlgroupby116 as nvarchar(500)

				set  @l_sqlstr116 =  N'SELECT MONTH(MBMPayorEffDate) AS Month,
						     YEAR(MBMPayorEffDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby116 = 'GROUP BY MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)'
				set @l_sqlstr116 = @l_sqlstr116 +  @strAppend + @l_sqlgroupby116
				exec  sp_executesql  @l_sqlstr116

		end

	if @strAppendCase = 117

		begin
			set nocount on
				declare @l_sqlstr117 as nvarchar(2000)
				declare @l_sqlgroupby117 as nvarchar(500)

				set  @l_sqlstr117 =  N'SELECT PAYCode, MONTH(MBMPayorEffDate) AS Month,
						     YEAR(MBMPayorEffDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby117 = 'GROUP BY PAYCode, MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)'
				set @l_sqlstr117 = @l_sqlstr117 +  @strAppend + @l_sqlgroupby117
				exec  sp_executesql  @l_sqlstr117

		end

	if @strAppendCase = 118

		begin
			set nocount on
				declare @l_sqlstr118 as nvarchar(2000)
				declare @l_sqlgroupby118 as nvarchar(500)

				set  @l_sqlstr118 =  N'SELECT PAYCode, MONTH(MBMPayorEffDate) AS Month,
						     YEAR(MBMPayorEffDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby118 = 'GROUP BY PAYCode, MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)'
				set @l_sqlstr118 = @l_sqlstr118 +  @strAppend + @l_sqlgroupby118
				exec  sp_executesql  @l_sqlstr118

		end

	if @strAppendCase = 119

		begin

			set nocount on

				declare @l_sqlstr119 as nvarchar(2000)
				declare @l_sqlgroupby119 as nvarchar(500)

				set  @l_sqlstr119 =  N'SELECT MONTH(MBMPayorExpDate) AS Month,
						     YEAR(MBMPayorExpDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby119 = 'GROUP BY MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)'
				set @l_sqlstr119 = @l_sqlstr119 +  @strAppend + @l_sqlgroupby119
				exec  sp_executesql  @l_sqlstr119

		end

	if @strAppendCase = 120

		begin
			set nocount on
				declare @l_sqlstr120 as nvarchar(2000)
				declare @l_sqlgroupby120 as nvarchar(500)

				set  @l_sqlstr120 =  N'SELECT MONTH(MBMPayorExpDate) AS Month,
						     YEAR(MBMPayorExpDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby120 = 'GROUP BY MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)'
				set @l_sqlstr120 = @l_sqlstr120 +  @strAppend + @l_sqlgroupby120
				exec  sp_executesql  @l_sqlstr120

		end

	if @strAppendCase = 121

		begin
			set nocount on
				declare @l_sqlstr121 as nvarchar(2000)
				declare @l_sqlgroupby121 as nvarchar(500)

				set  @l_sqlstr121 =  N'SELECT PAYCode, MONTH(MBMPayorExpDate) AS Month,
						     YEAR(MBMPayorExpDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby121 = 'GROUP BY PAYCode, MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)'
				set @l_sqlstr121 = @l_sqlstr121 +  @strAppend + @l_sqlgroupby121
				exec  sp_executesql  @l_sqlstr121

		end

	if @strAppendCase = 122

		begin
			set nocount on
				declare @l_sqlstr122 as nvarchar(2000)
				declare @l_sqlgroupby122 as nvarchar(500)

				set  @l_sqlstr122 =  N'SELECT PAYCode, MONTH(MBMPayorExpDate) AS Month,
						     YEAR(MBMPayorExpDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby122 = 'GROUP BY PAYCode, MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)'
				set @l_sqlstr122 = @l_sqlstr122 +  @strAppend + @l_sqlgroupby122
				exec  sp_executesql  @l_sqlstr122

		end

	if @strAppendCase = 123

		begin
			set nocount on
				declare @l_sqlstr123 as nvarchar(2000)
				declare @l_sqlgroupby123 as nvarchar(500)

				set  @l_sqlstr123 =  N'SELECT MONTH(MBMValueDate) AS Month,
						     YEAR(MBMValueDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby123 = 'GROUP BY MONTH(MBMValueDate), YEAR(MBMValueDate)'
				set @l_sqlstr123 = @l_sqlstr123 +  @strAppend + @l_sqlgroupby123
				exec  sp_executesql  @l_sqlstr123

		end

	if @strAppendCase = 124

		begin
			set nocount on
				declare @l_sqlstr124 as nvarchar(2000)
				declare @l_sqlgroupby124 as nvarchar(500)

				set  @l_sqlstr124 =  N'SELECT MONTH(MBMValueDate) AS Month,
						     YEAR(MBMValueDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby124 = 'GROUP BY MONTH(MBMValueDate), YEAR(MBMValueDate)'
				set @l_sqlstr124 = @l_sqlstr124 +  @strAppend + @l_sqlgroupby124
				exec  sp_executesql  @l_sqlstr124

		end

	if @strAppendCase = 125

		begin
			set nocount on
				declare @l_sqlstr125 as nvarchar(2000)
				declare @l_sqlgroupby125 as nvarchar(500)

				set  @l_sqlstr125 =  N'SELECT PAYCode, MONTH(MBMValueDate) AS Month,
						     YEAR(MBMValueDate) AS Year, COUNT(MBMNumber) AS TotalP,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross,
						     SUM(MBMAccessFee) AS EAFee FROM VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby125 = 'GROUP BY PAYCode, MONTH(MBMValueDate), YEAR(MBMValueDate)'
				set @l_sqlstr125 = @l_sqlstr125 +  @strAppend + @l_sqlgroupby125
				exec  sp_executesql  @l_sqlstr125

		end

	if @strAppendCase = 126

		begin
			set nocount on
				declare @l_sqlstr126 as nvarchar(2000)
				declare @l_sqlgroupby126 as nvarchar(500)

				set  @l_sqlstr126 =  N'SELECT PAYCode, MONTH(MBMValueDate) AS Month,
						     YEAR(MBMValueDate) AS Year, COUNT(MBMNumber) AS TotalS,
						     SUM(MBMPremium) AS PremiumGross, SUM(MBMMCOFee) AS MCOGross
						     FROM View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby126 = 'GROUP BY PAYCode, MONTH(MBMValueDate), YEAR(MBMValueDate)'
				set @l_sqlstr126 = @l_sqlstr126 +  @strAppend + @l_sqlgroupby126
				exec  sp_executesql  @l_sqlstr126

		end

	if @strAppendCase = 127

		begin
			set nocount on
				declare @l_sqlstr127 as nvarchar(2000)

				set  @l_sqlstr127 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMMemberType, MBMBordxDate, MBMBatchNo,
						     MBMCCoverID, MBMStatus, MBMInvoiceNo, MBMInvoiceDate, MBMMCOFee, MBMValueDate
						     From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlstr127 = @l_sqlstr127 +  @strAppend
				exec  sp_executesql  @l_sqlstr127

		end

	if @strAppendCase = 128

		begin
			set nocount on
				declare @l_sqlstr128 as nvarchar(2000)

				set  @l_sqlstr128 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMMemberType, MBMBordxDate, MBMBatchNo,
						     MBMCCoverID, MBMStatus, MBMInvoiceNo, MBMInvoiceDate, MBMMCOFee, MBMValueDate,
						     RELCode From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr128 = @l_sqlstr128 +  @strAppend
				exec  sp_executesql  @l_sqlstr128

		end

	if @strAppendCase = 129

		begin
			set nocount on
				declare @l_sqlstr129 as nvarchar(2000)

				set  @l_sqlstr129 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMMemberType, MBMBordxDate, MBMBatchNo,
						     MBMCCoverID, MBMStatus, MBMInvoiceNo, MBMInvoiceDate, MBMMCOFee, MBMValueDate,
						     RELCode From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr129 = @l_sqlstr129 +  @strAppend
				exec  sp_executesql  @l_sqlstr129

		end

	if @strAppendCase = 130

		begin
			set nocount on
				declare @l_sqlstr130 as nvarchar(2000)
				declare @l_sqlgroupby130 as nvarchar(500)

				set  @l_sqlstr130 =  N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby130 = 'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate'
				set @l_sqlstr130 = @l_sqlstr130 +  @strAppend + @l_sqlgroupby130
				exec  sp_executesql  @l_sqlstr130

		end

	if @strAppendCase = 131

		begin
			set nocount on
				declare @l_sqlstr131 as nvarchar(2000)
				declare @l_sqlgroupby131 as nvarchar(500)

				set  @l_sqlstr131 =  N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby131 = 'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate'
				set @l_sqlstr131 = @l_sqlstr131 +  @strAppend + @l_sqlgroupby131
				exec  sp_executesql  @l_sqlstr131

		end

	if @strAppendCase = 132
	-----VIEW_CISMBM_Search
		begin
			set nocount on
				declare @l_sqlstr132 as nvarchar(2000)
				declare @l_sqlgroupby132 as nvarchar(500)

				set  @l_sqlstr132 =  N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate, COUNT(MBMNumber) AS TotalMBM,
						        SUM(MBMMCOFEE) AS MCOFEE,MBMGroupCompany,Subsidiary From VIEW_MBM_Search_MCO_V1 Where 0 = 0'
				set @l_sqlgroupby132 = 'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate,MBMGroupCompany,Subsidiary'
				set @l_sqlstr132 = @l_sqlstr132 +  @strAppend + @l_sqlgroupby132
				exec  sp_executesql  @l_sqlstr132

		end

	if @strAppendCase = 133

		begin
			set nocount on
				declare @l_sqlstr133 as nvarchar(2000)
				declare @l_sqlgroupby133 as nvarchar(500)



				set  @l_sqlstr133 =  N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate, COUNT(MBMNumber) AS TotalMBMS,
						        SUM(MBMMCOFEE) AS MCOFEE From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby133 = 'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate'
				set @l_sqlstr133 = @l_sqlstr133 +  @strAppend + @l_sqlgroupby133
				exec  sp_executesql  @l_sqlstr133

		end

	if @strAppendCase = 134

		begin
			set nocount on
				declare @l_sqlstr134 as nvarchar(2000)
				declare @l_sqlgroupby134 as nvarchar(500)

				set  @l_sqlstr134 =  N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate, COUNT(MBMNumber) AS TotalMBM,
						        SUM(MBMPremium) AS Premium From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby134 = 'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate'
				set @l_sqlstr134 = @l_sqlstr134 +  @strAppend + @l_sqlgroupby134
				exec  sp_executesql  @l_sqlstr134

		end

	if @strAppendCase = 135

		begin
			set nocount on
				declare @l_sqlstr135 as nvarchar(2000)
				declare @l_sqlgroupby135 as nvarchar(500)

				set  @l_sqlstr135 =  N'SELECT PAYCode, MBMInvoiceNo, MBMInvoiceDate, COUNT(MBMNumber) AS TotalMBMS,
						        SUM(MBMPremium) AS Premium From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby135 = 'GROUP BY PAYCode, MBMInvoiceNo, MBMInvoiceDate'
				set @l_sqlstr135 = @l_sqlstr135 +  @strAppend + @l_sqlgroupby135
				exec  sp_executesql  @l_sqlstr135

		end

	if @strAppendCase = 136

		begin
			set nocount on
				declare @l_sqlstr136 as nvarchar(2000)
				declare @l_sqlgroupby136 as nvarchar(500)

				set  @l_sqlstr136 =  N'SELECT HLTCode, PAYCode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby136 = 'GROUP BY HLTCode, PAYCode'
				set @l_sqlstr136 = @l_sqlstr136 +  @strAppend + @l_sqlgroupby136
				exec  sp_executesql  @l_sqlstr136

		end

	if @strAppendCase = 137

		begin
			set nocount on
				declare @l_sqlstr137 as nvarchar(2000)
				declare @l_sqlgroupby137 as nvarchar(500)

				set  @l_sqlstr137 =  N'SELECT HLTCode, PAYCode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby137 = 'GROUP BY HLTCode, PAYCode'
				set @l_sqlstr137 = @l_sqlstr137 +  @strAppend + @l_sqlgroupby137
				exec  sp_executesql  @l_sqlstr137

		end

	if @strAppendCase = 138

		begin
			set nocount on
				declare @l_sqlstr138 as nvarchar(2000)
				declare @l_sqlgroupby138 as nvarchar(500)

				set  @l_sqlstr138 =  N'SELECT PAYCode, SUM(MBMMCOFee) AS MBMMCOFee, SUM(MBMBasicMCO) AS MBMBasicMCO,
						        SUM(MBMPayorMCOGross) AS MBMPayorMCOGross, SUM(MBMPayorMCONet) AS MBMPayorMCONet,
						        SUM(MBMPremium) AS MBMPremium, SUM(MBMBasicPrem) AS MBMBasicPrem,
						        SUM(MBMPayorPremGross) AS MBMPayorPremGross, SUM(MBMPayorPremNet) AS MBMPayorPremNet
						        From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby138 = 'GROUP BY PAYCode'
				set @l_sqlstr138 = @l_sqlstr138 +  @strAppend + @l_sqlgroupby138
				exec  sp_executesql  @l_sqlstr138

		end

	if @strAppendCase = 139

		begin
			set nocount on
				declare @l_sqlstr139 as nvarchar(2000)
				declare @l_sqlgroupby139 as nvarchar(500)

				set  @l_sqlstr139 =  N'SELECT PAYCode, SUM(MBMMCOFee) AS MBMMCOFee, SUM(MBMBasicMCO) AS MBMBasicMCO,
						        SUM(MBMPayorMCOGross) AS MBMPayorMCOGross, SUM(MBMPayorMCONet) AS MBMPayorMCONet,
						        SUM(MBMPremium) AS MBMPremium, SUM(MBMBasicPrem) AS MBMBasicPrem,
						        SUM(MBMPayorPremGross) AS MBMPayorPremGross, SUM(MBMPayorPremNet) AS MBMPayorPremNet
						        From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby139 = 'GROUP BY PAYCode'
				set @l_sqlstr139 = @l_sqlstr139 +  @strAppend + @l_sqlgroupby139
				exec  sp_executesql  @l_sqlstr139

		end

	if @strAppendCase = 140

		begin
			set nocount on
				declare @l_sqlstr140 as nvarchar(2000)
				declare @l_sqlgroupby140 as nvarchar(500)

				set  @l_sqlstr140 =  N'SELECT PAYCompanyName, MBMGroupCompany, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMPremium) AS PremiumGross, SUM(CLMTotalApproved) AS CLMTotalApproved
						        From VIEW_MBM_Search4 Where (INSCode = "G" OR INSCode = "H")'
				set @l_sqlgroupby140 = 'GROUP BY PAYCompanyName, MBMGroupCompany'
				set @l_sqlstr140 = @l_sqlstr140 +  @strAppend + @l_sqlgroupby140
				exec  sp_executesql  @l_sqlstr140

		end

	if @strAppendCase = 141

		begin
			set nocount on
				declare @l_sqlstr141 as nvarchar(2000)
				declare @l_sqlgroupby141 as nvarchar(500)

				set  @l_sqlstr141 =  N'SELECT PAYCompanyName, MBMGroupCompany, COUNT(MBMNumber) AS TotalS
						        From View_MBM_Supplementary4 Where (INSCode = "G" OR INSCode = "H")'
				set @l_sqlgroupby141 = 'GROUP BY PAYCompanyName, MBMGroupCompany'
				set @l_sqlstr141 = @l_sqlstr141 +  @strAppend + @l_sqlgroupby141
				exec  sp_executesql  @l_sqlstr141

		end

	if @strAppendCase = 142

		begin
			set nocount on
				declare @l_sqlstr142 as nvarchar(2000)
				declare @l_sqlgroupby142 as nvarchar(500)

				set  @l_sqlstr142 =  N'SELECT HLTCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby142 = 'GROUP BY HLTCode, INSCode, PLNCode'
				set @l_sqlstr142 = @l_sqlstr142 +  @strAppend + @l_sqlgroupby142
				exec  sp_executesql  @l_sqlstr142

		end

	if @strAppendCase = 143

		begin
			set nocount on
				declare @l_sqlstr143 as nvarchar(2000)
				declare @l_sqlgroupby143 as nvarchar(500)

				set  @l_sqlstr143 =  N'SELECT HLTCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby143 = 'GROUP BY HLTCode, INSCode, PLNCode'
				set @l_sqlstr143 = @l_sqlstr143 +  @strAppend + @l_sqlgroupby143
				exec  sp_executesql  @l_sqlstr143

		end


	if @strAppendCase = 144

		begin
			set nocount on
				declare @l_sqlstr144 as nvarchar(2000)
				declare @l_sqlgroupby144 as nvarchar(500)

				set  @l_sqlstr144 =  N'SELECT PAYCode, HLTCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby144 = 'GROUP BY PAYCode, HLTCode, INSCode, PLNCode'
				set @l_sqlstr144 = @l_sqlstr144 +  @strAppend + @l_sqlgroupby144
				exec  sp_executesql  @l_sqlstr144

		end

	if @strAppendCase = 145

		begin
			set nocount on
				declare @l_sqlstr145 as nvarchar(2000)
				declare @l_sqlgroupby145 as nvarchar(500)

				set  @l_sqlstr145 =  N'SELECT PAYCode, HLTCode, INSCode, PLNCode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby145 = 'GROUP BY PAYCode, HLTCode, INSCode, PLNCode'
				set @l_sqlstr145 = @l_sqlstr145 +  @strAppend + @l_sqlgroupby145
				exec  sp_executesql  @l_sqlstr145

		end

	if @strAppendCase = 146

		begin
			set nocount on
				declare @l_sqlstr146 as nvarchar(2000)
				declare @l_sqlgroupby146 as nvarchar(500)

				set  @l_sqlstr146 =  N'SELECT HLTCode, INSCode, PLNCode, AGECode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby146 = 'GROUP BY HLTCode, INSCode, PLNCode, AGECode'
				set @l_sqlstr146 = @l_sqlstr146 +  @strAppend + @l_sqlgroupby146
				exec  sp_executesql  @l_sqlstr146

		end

	if @strAppendCase = 147

		begin
			set nocount on
				declare @l_sqlstr147 as nvarchar(2000)
				declare @l_sqlgroupby147 as nvarchar(500)

				set  @l_sqlstr147 =  N'SELECT HLTCode, INSCode, PLNCode, PrinAGECode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby147 = 'GROUP BY HLTCode, INSCode, PLNCode, PrinAGECode'
				set @l_sqlstr147 = @l_sqlstr147 +  @strAppend + @l_sqlgroupby147
				exec  sp_executesql  @l_sqlstr147

		end

	if @strAppendCase = 148

		begin
			set nocount on
				declare @l_sqlstr148 as nvarchar(2000)
				declare @l_sqlgroupby148 as nvarchar(500)

				set  @l_sqlstr148 =  N'SELECT PAYCode, HLTCode, INSCode, PLNCode, AGECode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby148 = 'GROUP BY PAYCode, HLTCode, INSCode, PLNCode, AGECode'
				set @l_sqlstr148 = @l_sqlstr148 +  @strAppend + @l_sqlgroupby148
				exec  sp_executesql  @l_sqlstr148

		end

	if @strAppendCase = 149

		begin
			set nocount on
				declare @l_sqlstr149 as nvarchar(2000)
				declare @l_sqlgroupby149 as nvarchar(500)

				set  @l_sqlstr149 =  N'SELECT PAYCode, HLTCode, INSCode, PLNCode, PrinAGECode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby149 = 'GROUP BY PAYCode, HLTCode, INSCode, PLNCode, PrinAGECode'
				set @l_sqlstr149 = @l_sqlstr149 +  @strAppend + @l_sqlgroupby149
				exec  sp_executesql  @l_sqlstr149

		end


	if @strAppendCase = 150

		begin
			set nocount on
				declare @l_sqlstr150 as nvarchar(2000)
				declare @l_sqlgroupby150 as nvarchar(500)

				set  @l_sqlstr150 =  N'SELECT HLTCode, INSCode, PLNCode, AGECode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby150 = 'GROUP BY HLTCode, INSCode, PLNCode, AGECode'
				set @l_sqlstr150 = @l_sqlstr150 +  @strAppend + @l_sqlgroupby150
				exec  sp_executesql  @l_sqlstr150

		end

	if @strAppendCase = 151

		begin
			set nocount on
				declare @l_sqlstr151 as nvarchar(2000)
				declare @l_sqlgroupby151 as nvarchar(500)

				set  @l_sqlstr151 =  N'SELECT PAYCode, HLTCode, INSCode, PLNCode, AGECode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFee) AS MCOGross, SUM(MBMPremium) AS PremiumGross
						        From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby151 = 'GROUP BY PAYCode, HLTCode, INSCode, PLNCode, AGECode'
				set @l_sqlstr151 = @l_sqlstr151 +  @strAppend + @l_sqlgroupby151
				exec  sp_executesql  @l_sqlstr151

		end

	if @strAppendCase = 152

		begin
			set nocount on
				declare @l_sqlstr152 as nvarchar(2000)
				declare @l_sqlgroupby152 as nvarchar(500)

				set  @l_sqlstr152 =  N'SELECT PAYCode, MBMPolicyNo, MBMGroupCompany, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby152 = 'GROUP BY PAYCode, MBMPolicyNo, MBMGroupCompany'
				set @l_sqlstr152 = @l_sqlstr152 +  @strAppend + @l_sqlgroupby152
				exec  sp_executesql  @l_sqlstr152

		end

	if @strAppendCase = 153

		begin
			set nocount on
				declare @l_sqlstr153 as nvarchar(2000)
				declare @l_sqlgroupby153 as nvarchar(500)

				set  @l_sqlstr153 =  N'SELECT PAYCode, MBMPolicyNo, MBMGroupCompany, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby153 = 'GROUP BY PAYCode, MBMPolicyNo, MBMGroupCompany'
				set @l_sqlstr153 = @l_sqlstr153 +  @strAppend + @l_sqlgroupby153
				exec  sp_executesql  @l_sqlstr153

		end

	if @strAppendCase = 154

		begin
			set nocount on
				declare @l_sqlstr154 as nvarchar(2000)
				declare @l_sqlgroupby154 as nvarchar(500)

				set  @l_sqlstr154 =  N'SELECT HLTCode, INSCode, RACCode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby154 = 'GROUP BY HLTCode, INSCode, RACCode'
				set @l_sqlstr154 = @l_sqlstr154 +  @strAppend + @l_sqlgroupby154
				exec  sp_executesql  @l_sqlstr154

		end

	if @strAppendCase = 155

		begin
			set nocount on
				declare @l_sqlstr155 as nvarchar(2000)
				declare @l_sqlgroupby155 as nvarchar(500)

				set  @l_sqlstr155 =  N'SELECT HLTCode, INSCode, RACCode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby155 = 'GROUP BY HLTCode, INSCode, RACCode'
				set @l_sqlstr155 = @l_sqlstr155 +  @strAppend + @l_sqlgroupby155
				exec  sp_executesql  @l_sqlstr155

		end

	if @strAppendCase = 156

		begin
			set nocount on
				declare @l_sqlstr156 as nvarchar(2000)
				declare @l_sqlgroupby156 as nvarchar(500)

				set  @l_sqlstr156 =  N'SELECT PAYCode, HLTCode, INSCode, RACCode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlgroupby156 = 'GROUP BY PAYCode, HLTCode, INSCode, RACCode'
				set @l_sqlstr156 = @l_sqlstr156 +  @strAppend + @l_sqlgroupby156
				exec  sp_executesql  @l_sqlstr156

		end

	if @strAppendCase = 157

		begin
			set nocount on
				declare @l_sqlstr157 as nvarchar(2000)
				declare @l_sqlgroupby157 as nvarchar(500)

				set  @l_sqlstr157 =  N'SELECT PAYCode, HLTCode, INSCode, RACCode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby157 = 'GROUP BY PAYCode, HLTCode, INSCode, RACCode'
				set @l_sqlstr157 = @l_sqlstr157 +  @strAppend + @l_sqlgroupby157
				exec  sp_executesql  @l_sqlstr157

		end

	if @strAppendCase = 158

		begin
			set nocount on
				declare @l_sqlstr158 as nvarchar(2000)
				declare @l_sqlgroupby158 as nvarchar(500)

				set  @l_sqlstr158 =  N'SELECT PAYCode, HLTCode, RELCode, AGECode, COUNT(DISTINCT MBMNumber) AS TotalP,
						        COUNT(MBMNumber) AS TotalS From VIEW_MBM_Search3 Where 0 = 0'
				set @l_sqlgroupby158 = 'GROUP BY PAYCode, HLTCode, RELCode, AGECode'
				set @l_sqlstr158 = @l_sqlstr158 +  @strAppend + @l_sqlgroupby158
				exec  sp_executesql  @l_sqlstr158

		end

	if @strAppendCase = 159

		begin
			set nocount on
				declare @l_sqlstr159 as nvarchar(2000)
				declare @l_sqlgroupby159 as nvarchar(500)

				set  @l_sqlstr159 =  N'SELECT PAYCode, HLTCode, RELCode, AGECode, COUNT(DISTINCT MBMNumber) AS TotalP,
						        COUNT(MBMNumber) AS TotalS From View_CISMBM_Supplementary Where RELCode <> "P"'
				set @l_sqlgroupby159 = 'GROUP BY PAYCode, HLTCode, RELCode, AGECode'
				set @l_sqlstr159 = @l_sqlstr159 +  @strAppend + @l_sqlgroupby159
				exec  sp_executesql  @l_sqlstr159

		end

	if @strAppendCase = 160

		begin
			set nocount on
				declare @l_sqlstr160 as nvarchar(2000)
				declare @l_sqlgroupby160 as nvarchar(500)

				set  @l_sqlstr160 =  N'SELECT PAYCode, HLTCode, INSCode, RELCode, COUNT(DISTINCT MBMNumber) AS TotalP,
						        COUNT(MBMNumber) AS TotalS From VIEW_MBM_Search3 Where 0 = 0'
				set @l_sqlgroupby160 = 'GROUP BY PAYCode, HLTCode, INSCode, RELCode'
				set @l_sqlstr160 = @l_sqlstr160 +  @strAppend + @l_sqlgroupby160
				exec  sp_executesql  @l_sqlstr160




		end

	if @strAppendCase = 161

		begin
			set nocount on
				declare @l_sqlstr161 as nvarchar(2000)
				declare @l_sqlgroupby161 as nvarchar(500)

				set  @l_sqlstr161 =  N'SELECT PAYCode, HLTCode, INSCode, RELCode, COUNT(DISTINCT MBMNumber) AS TotalP,
						        COUNT(MBMNumber) AS TotalS From View_CISMBM_Supplementary Where RELCode <> "P"'
				set @l_sqlgroupby161 = 'GROUP BY PAYCode, HLTCode, INSCode, RELCode'
				set @l_sqlstr161 = @l_sqlstr161 +  @strAppend + @l_sqlgroupby161
				exec  sp_executesql  @l_sqlstr161

		end

	if @strAppendCase = 162

		begin
			set nocount on
				declare @l_sqlstr162 as nvarchar(2000)
				declare @l_sqlgroupby162 as nvarchar(500)

				set  @l_sqlstr162 =  N'SELECT HLTCode, STACode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'	
				set @l_sqlgroupby162 = 'GROUP BY HLTCode, STACode'
				set @l_sqlstr162 = @l_sqlstr162 +  @strAppend + @l_sqlgroupby162
				exec  sp_executesql  @l_sqlstr162

		end

	if @strAppendCase = 163

		begin
			set nocount on
				declare @l_sqlstr163 as nvarchar(2000)
				declare @l_sqlgroupby163 as nvarchar(500)

				set  @l_sqlstr163 =  N'SELECT HLTCode, STACode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby163 = 'GROUP BY HLTCode, STACode'
				set @l_sqlstr163 = @l_sqlstr163 +  @strAppend + @l_sqlgroupby163
				exec  sp_executesql  @l_sqlstr163

		end

	if @strAppendCase = 164

		begin
			set nocount on
				declare @l_sqlstr164 as nvarchar(2000)
				declare @l_sqlgroupby164 as nvarchar(500)

				set  @l_sqlstr164 =  N'SELECT PAYCode, HLTCode, STACode, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'	
				set @l_sqlgroupby164 = 'GROUP BY PAYCode, HLTCode, STACode'
				set @l_sqlstr164 = @l_sqlstr164 +  @strAppend + @l_sqlgroupby164
				exec  sp_executesql  @l_sqlstr164

		end

	if @strAppendCase = 165

		begin
			set nocount on
				declare @l_sqlstr165 as nvarchar(2000)
				declare @l_sqlgroupby165 as nvarchar(500)

				set  @l_sqlstr165 =  N'SELECT PAYCode, HLTCode, STACode, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby165 = 'GROUP BY PAYCode, HLTCode, STACode'
				set @l_sqlstr165 = @l_sqlstr165 +  @strAppend + @l_sqlgroupby165
				exec  sp_executesql  @l_sqlstr165

		end

	if @strAppendCase = 166

		begin
			set nocount on
				declare @l_sqlstr166 as nvarchar(2000)
				declare @l_sqlgroupby166 as nvarchar(500)

				set  @l_sqlstr166 =  N'SELECT HLTCode, PAYCode, MBMGroupCompany, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'	
				set @l_sqlgroupby166 = 'GROUP BY HLTCode, PAYCode, MBMGroupCompany'
				set @l_sqlstr166 = @l_sqlstr166 +  @strAppend + @l_sqlgroupby166
				exec  sp_executesql  @l_sqlstr166

		end

	if @strAppendCase = 167

		begin
			set nocount on
				declare @l_sqlstr167 as nvarchar(2000)
				declare @l_sqlgroupby167 as nvarchar(500)

				set  @l_sqlstr167 =  N'SELECT HLTCode, PAYCode, MBMGroupCompany, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby167 = 'GROUP BY HLTCode, PAYCode, MBMGroupCompany'
				set @l_sqlstr167 = @l_sqlstr167 +  @strAppend + @l_sqlgroupby167
				exec  sp_executesql  @l_sqlstr167

		end

	if @strAppendCase = 168

		begin
			set nocount on
				declare @l_sqlstr168 as nvarchar(2000)
				declare @l_sqlgroupby168 as nvarchar(500)

				set  @l_sqlstr168 =  N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'	
				set @l_sqlgroupby168 = 'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMDepartment'
				set @l_sqlstr168 = @l_sqlstr168 +  @strAppend + @l_sqlgroupby168
				exec  sp_executesql  @l_sqlstr168

		end

	if @strAppendCase = 169

		begin
			set nocount on
				declare @l_sqlstr169 as nvarchar(2000)
				declare @l_sqlgroupby169 as nvarchar(500)

				set  @l_sqlstr169 =  N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby169 = 'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMDepartment'
				set @l_sqlstr169 = @l_sqlstr169 +  @strAppend + @l_sqlgroupby169
				exec  sp_executesql  @l_sqlstr169

		end

	if @strAppendCase = 170

		begin
			set nocount on
				declare @l_sqlstr170 as nvarchar(2000)

				set  @l_sqlstr170 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMCCoverID, MBMBordxDate, MBMInvoiceDate,
						     MBMCLNCardSentDate, MBMPayorEffDate, BRNCode, MBMAgentCode, MBMStatus, MBMMemberType
						     From VIEW_CISMBM_Search Where 0 = 0'
				set @l_sqlstr170 = @l_sqlstr170 +  @strAppend
				exec  sp_executesql  @l_sqlstr170

		end

	if @strAppendCase = 171

		begin
			set nocount on
				declare @l_sqlstr171 as nvarchar(2000)

				set  @l_sqlstr171 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMCCoverID, MBMBordxDate, MBMInvoiceDate,
						     MBMCLNCardSentDate, MBMPayorEffDate, BRNCode, MBMAgentCode, MBMStatus, MBMMemberType
						     From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr171 = @l_sqlstr171 +  @strAppend
				exec  sp_executesql  @l_sqlstr171

		end

	if @strAppendCase = 172

		begin
			set nocount on
				declare @l_sqlstr172 as nvarchar(2000)

				set  @l_sqlstr172 =  N'Select MBMNumber, MBMPolicyNo, MBMName, MBMCCoverID, MBMBordxDate, MBMInvoiceDate,
						     MBMCLNCardSentDate, MBMPayorEffDate, BRNCode, MBMAgentCode, MBMStatus, MBMMemberType
						     From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlstr172 = @l_sqlstr172 +  @strAppend
				exec  sp_executesql  @l_sqlstr172

		end

	if @strAppendCase = 173

		begin
			set nocount on
				declare @l_sqlstr173 as nvarchar(2000)
				declare @l_sqlgroupby173 as nvarchar(500)

				set  @l_sqlstr173 =  N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'	
				set @l_sqlgroupby173 = 'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMBordxDate, MBMBatchNo'
				set @l_sqlstr173 = @l_sqlstr173 +  @strAppend + @l_sqlgroupby173
				exec  sp_executesql  @l_sqlstr173

		end

	if @strAppendCase = 174

		begin
			set nocount on
				declare @l_sqlstr174 as nvarchar(2000)
				declare @l_sqlgroupby174 as nvarchar(500)

				set  @l_sqlstr174 =  N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby174 = 'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMBordxDate, MBMBatchNo'
				set @l_sqlstr174 = @l_sqlstr174 +  @strAppend + @l_sqlgroupby174
				exec  sp_executesql  @l_sqlstr174

		end

	if @strAppendCase = 175

		begin
			set nocount on
				declare @l_sqlstr175 as nvarchar(2000)
				declare @l_sqlgroupby175 as nvarchar(500)

				set  @l_sqlstr175 =  N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalP,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMAccessFee) AS EAFee,
						        SUM(MBMPremium) AS PremiumGross From VIEW_CISMBM_Search Where 0 = 0'	
				set @l_sqlgroupby175 = 'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, MBMBordxDate, MBMBatchNo'
				set @l_sqlstr175 = @l_sqlstr175 +  @strAppend + @l_sqlgroupby175
				exec  sp_executesql  @l_sqlstr175

		end

	if @strAppendCase = 176

		begin
			set nocount on
				declare @l_sqlstr176 as nvarchar(2000)
				declare @l_sqlgroupby176 as nvarchar(500)

				set  @l_sqlstr176 =  N'SELECT HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, MBMBordxDate, MBMBatchNo, COUNT(MBMNumber) AS TotalS,
						        SUM(MBMMCOFEE) AS MCOGross, SUM(MBMPremium) AS PremiumGross From View_CISMBM_Supplementary Where 0 = 0'
				set @l_sqlgroupby176 = 'GROUP BY HLTCode, PAYCode, MBMGroupCompany, MBMDepartment, MBMBordxDate, MBMBatchNo'
				set @l_sqlstr176 = @l_sqlstr176 +  @strAppend + @l_sqlgroupby176
				exec  sp_executesql  @l_sqlstr176

		end


		if @strAppendCase = 177

		begin
			set nocount on
				declare @l_sqlstr177 as nvarchar(MAX)
				declare @l_sqlgroupby177 as nvarchar(MAX)

				set  @l_sqlstr177 =  N'Select * from HISDB..View_CISInvoice_Summary Where 0 = 0'
				set @l_sqlgroupby177 = ''
				set @l_sqlstr177 = @l_sqlstr177 +  @strAppend + @l_sqlgroupby177
				exec  sp_executesql  @l_sqlstr177

		end


		if @strAppendCase = 178

		begin
			set nocount on
			declare @l_sqlstr178 as nvarchar(MAX)
				declare @l_sqlgroupby178 as nvarchar(MAX)
				declare @l_sqlpivot178 as nvarchar(MAX)

				set  @l_sqlstr178 =  N' SELECT case when MBMCCoverID=''00'' then ''Principal'' else ''Dependent'' end as cc,
										PAYCode, MBMInvoiceNo, MBMInvoiceDate, COUNT(MBMNumber) AS TotalMBM,SUM(MBMMCOFEE) AS MCOFEE 
										,MBMGroupCompany,coalesce(Subsidiary,'''') as Subsidiary,OPAccCode
										into #temp 
										From VIEW_MBM_Search_MCO_V1 Where 0 = 0  '
				set @l_sqlgroupby178 = ' GROUP BY case when MBMCCoverID=''00'' then ''Principal'' else ''Dependent'' end,PAYCode, MBMInvoiceNo, MBMInvoiceDate,MBMGroupCompany,Subsidiary,OPAccCode '

				Set @l_sqlpivot178= '
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
									'

					set @l_sqlstr178 = @l_sqlstr178 +  @strAppend + @l_sqlgroupby178
					set @l_sqlstr178=@l_sqlstr178+@l_sqlpivot178
					exec  sp_executesql  @l_sqlstr178
		end 


	return












