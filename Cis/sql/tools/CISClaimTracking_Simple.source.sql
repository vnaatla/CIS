CREATE procedure CISClaimTracking_Simple

	@strAppend as nvarchar (2000), @strAppendCase as int

              
as
	if @strAppendCase = 1

		begin
         	      		set nocount on
			       declare @l_sqlstr as nvarchar(2000)

		                    set  @l_sqlstr =  N'SELECT MBMDepartment, CASNumber, BordxRefNo, BordxDate, MBMNumber, CoverID,
            						PatientName, ClinicCode, CLNName, INVPayTo, CASTreatReason1, CLNType,
            						CASDiagnosis5, ReasonDesc, CLMTotalActual, CLMNonGpFee, DateVisit, MCDays
            						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlstr = @l_sqlstr + @strAppend
			       exec  sp_executesql  @l_sqlstr
		end

	if @strAppendCase = 2

		begin
         	      		set nocount on
			       declare @l_sqlstr2 as nvarchar(2000)
		                    declare @l_sqlgroupby2 as nvarchar(500)

		                    set  @l_sqlstr2 =  N'SELECT ClinicCode, CLNName, CLNType, COUNT(CASNumber) as NoofVisit,
            						SUM(MCStatus) as MCTimes, SUM(MCDays) as MCDays, SUM(CLMTotalActual) as BillAmt
            						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby2 = 'GROUP BY ClinicCode, CLNName, CLNType'
			       set @l_sqlstr2 = @l_sqlstr2 + @strAppend + @l_sqlgroupby2
			       exec  sp_executesql  @l_sqlstr2
		end

	if @strAppendCase = 3

		begin
         	      		set nocount on
			       declare @l_sqlstr3 as nvarchar(2000)
		                    declare @l_sqlgroupby3 as nvarchar(500)

		                    set  @l_sqlstr3 =  N'SELECT MBMDepartment, BordxRefNo, BordxDate, 
            						SUM(CLMTotalActual) as CLMTotalActual, SUM(CLMNonGPFee) as CLMNonGPFee
            						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby3 = 'GROUP BY MBMDepartment, BordxRefNo, BordxDate'
			       set @l_sqlstr3 = @l_sqlstr3 + @strAppend + @l_sqlgroupby3
			       exec  sp_executesql  @l_sqlstr3
		end

	if @strAppendCase = 4

		begin
         	      		set nocount on
			       declare @l_sqlstr4 as nvarchar(2000)
		                    declare @l_sqlgroupby4 as nvarchar(500)

		                    set  @l_sqlstr4 =  N'SELECT MBMDepartment, MBMNumber, CoverID, PatientName, Count(CASNumber) AS TotalVisit,
          						SUM(MCStatus) as MCTimes, SUM(MCDays) AS MCDays,  
          						SUM(CLMTotalActual) as CLMTotalActual, SUM(CLMNonGPFee) as CLMNonGPFee 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby4 = 'GROUP BY MBMDepartment, MBMNumber, CoverID, PatientName'
			       set @l_sqlstr4 = @l_sqlstr4 + @strAppend + @l_sqlgroupby4
			       exec  sp_executesql  @l_sqlstr4
		end

	if @strAppendCase = 5

		begin
         	      		set nocount on
			       declare @l_sqlstr5 as nvarchar(2000)
		                    declare @l_sqlgroupby5 as nvarchar(500)

		                    set  @l_sqlstr5 =  N'SELECT MBMDepartment, MBMNumber, MBMName, CoverID, AnnLimit, AnnLmtIND, SUM(CLMTotalActual) AS ApprovedAmt,
          						Count(CASNumber) AS TotalClaim, SUM(MCStatus) AS MCTimes, 
          						SUM(MCDays) AS MCDays, MBMCLNBalLmt,Sum(TopupAmt) AS TopupAmt From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby5 = 'GROUP BY MBMDepartment, MBMNumber, MBMName, CoverID, AnnLimit, AnnLmtIND, MBMCLNBalLmt
							ORDER BY MBMNumber'
			       set @l_sqlstr5 = @l_sqlstr5 + @strAppend + @l_sqlgroupby5
			       exec  sp_executesql  @l_sqlstr5
		end

	if @strAppendCase = 6

		begin
         	      		set nocount on
			       declare @l_sqlstr6 as nvarchar(2000)
		                    declare @l_sqlgroupby6 as nvarchar(500)

		                    set  @l_sqlstr6 =  N'SELECT CASTreatReason1, Sum(CLMTotalActual) As CLMTotalActual,
          						SUM(CLMNonGpFee) AS CLMNonGpFee From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby6 = 'GROUP BY CASTreatReason1'
			       set @l_sqlstr6 = @l_sqlstr6 + @strAppend + @l_sqlgroupby6
			       exec  sp_executesql  @l_sqlstr6
		end

	if @strAppendCase = 7

		begin
         	      		set nocount on
			       declare @l_sqlstr7 as nvarchar(2000)
		                    declare @l_sqlgroupby7 as nvarchar(500)

		                    set  @l_sqlstr7 =  N'SELECT Distinct BordxRefNo, BordxDate, BordxNoClaims, SUM(CLMTotalActual) 
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, 
          						BordxClaimAmt From VIEW_CLMReport_Simple Where BordxRefNo IS NOT NULL'
			       set @l_sqlgroupby7 = 'GROUP BY BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt'
			       set @l_sqlstr7 = @l_sqlstr7 + @strAppend + @l_sqlgroupby7
			       exec  sp_executesql  @l_sqlstr7
		end
	if @strAppendCase = 8

		begin
         	      		set nocount on
			       declare @l_sqlstr8 as nvarchar(2000)
		                    declare @l_sqlgroupby8 as nvarchar(500)

		                    set  @l_sqlstr8 =  N'SELECT Distinct MBMGroupCompany, BordxRefNo, BordxDate, BordxNoClaims, SUM(CLMTotalActual) 
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						BordxClaimAmt From VIEW_CLMReport_Simple Where BordxRefNo IS NOT NULL'
			       set @l_sqlgroupby8 = 'GROUP BY MBMGroupCompany, BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt'
			       set @l_sqlstr8 = @l_sqlstr8 + @strAppend + @l_sqlgroupby8
			       exec  sp_executesql  @l_sqlstr8
		end

	if @strAppendCase = 9

		begin
         	      		set nocount on
			       declare @l_sqlstr9 as nvarchar(2000)
		                    declare @l_sqlgroupby9 as nvarchar(500)

		                    set  @l_sqlstr9 =  N'SELECT Distinct MBMDepartment, MBMGroupCompany, BordxRefNo, BordxDate, BordxNoClaims, SUM(CLMTotalActual)
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						BordxClaimAmt From VIEW_CLMReport_Simple Where BordxRefNo IS NOT NULL'
			       set @l_sqlgroupby9 = 'GROUP BY MBMDepartment, MBMGroupCompany, BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt 
							ORDER BY BordxRefNo'
			       set @l_sqlstr9 = @l_sqlstr9 + @strAppend + @l_sqlgroupby9
			       exec  sp_executesql  @l_sqlstr9
		end

	if @strAppendCase = 10

		begin
         	      		set nocount on
			SELECT ReasonCode, ReasonDesc 
			FROM CISDB.dbo.TreatmentReason 
			WHERE ReasonCode = @strAppend
		end

	if @strAppendCase = 11

		begin
         	      		set nocount on
			       declare @l_sqlstr11 as nvarchar(2000)
		                    declare @l_sqlgroupby11 as nvarchar(500)

		                    set  @l_sqlstr11 =  N'SELECT COUNT(CASNumber) AS Total, Sum(CLMTotalActual) As TotalApproved, Sum(CLMNonGPFee) As CLMNonGPFee
                           				From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby11 = 'GROUP BY CASTreatReason1'
			       set @l_sqlstr11 = @l_sqlstr11 + @strAppend + @l_sqlgroupby11
			       exec  sp_executesql  @l_sqlstr11
		end
			
	if @strAppendCase = 12

		begin
         	      		set nocount on
			       declare @l_sqlstr12 as nvarchar(2000)
		                    declare @l_sqlgroupby12 as nvarchar(500)

		                    set  @l_sqlstr12 =  N'SELECT PLNCode, SUM(CLMTotalApproved) AS TotalPanelAmt
                  					From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby12 = 'GROUP BY PLNCode'
			       set @l_sqlstr12 = @l_sqlstr12 + @strAppend + @l_sqlgroupby12
			       exec  sp_executesql  @l_sqlstr12
			
		end

	if @strAppendCase = 13

		begin
         	      		set nocount on
			       declare @l_sqlstr13 as nvarchar(2000)
		                    declare @l_sqlgroupby13 as nvarchar(500)

		                    set  @l_sqlstr13 =  N'SELECT Sex, Ageband, COUNT(CASNumber) AS TotalCases,
          						SUM(CLMTotalActual) AS TotalActualAmt,
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby13 = 'GROUP BY Sex, Ageband'
			       set @l_sqlstr13 = @l_sqlstr13 + @strAppend + @l_sqlgroupby13
			       exec  sp_executesql  @l_sqlstr13
		end

	if @strAppendCase = 14

		begin
         	      		set nocount on
			       declare @l_sqlstr14 as nvarchar(2000)
		                    declare @l_sqlgroupby14 as nvarchar(500)

		                    set  @l_sqlstr14 =  N'SELECT Ageband, COUNT(CASNumber) AS TotalCases, 
          						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby14 = 'GROUP BY Ageband'
			       set @l_sqlstr14 = @l_sqlstr14 + @strAppend + @l_sqlgroupby14
			       exec  sp_executesql  @l_sqlstr14
		end

	if @strAppendCase = 15

		begin
         	      		set nocount on
			       declare @l_sqlstr15 as nvarchar(2000)
		                    declare @l_sqlgroupby15 as nvarchar(500)

		                    set  @l_sqlstr15 =  N'SELECT PAYCode, INSCode, Ageband, COUNT(CASNumber) AS TotalCases,
          						SUM(CLMTotalActual) AS TotalActualAmt,
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby15 = 'GROUP BY PAYCode, INSCode, Ageband'
			       set @l_sqlstr15 = @l_sqlstr15 + @strAppend + @l_sqlgroupby15
			       exec  sp_executesql  @l_sqlstr15
		end

	if @strAppendCase = 16

		begin
         	      		set nocount on
			       declare @l_sqlstr16 as nvarchar(2000)
		                    declare @l_sqlgroupby16 as nvarchar(500)

		                    set @l_sqlstr16 =  N'SELECT PAYCode, MBMAgentCode, CASDoctor, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby16 = 'GROUP BY PAYCode, MBMAgentCode, CASDoctor'
			       set @l_sqlstr16 = @l_sqlstr16 + @strAppend + @l_sqlgroupby16
			       exec  sp_executesql  @l_sqlstr16
		end

	if @strAppendCase = 17

		begin
         	      		set nocount on
			       declare @l_sqlstr17 as nvarchar(2000)

		                    set  @l_sqlstr17 =  N'SELECT PAYCode, MBMAgentCode, CASNumber, MBMNumber,
          						CoverID, PatientName, CLMTotalActual,
          						DateVisit, CLMTotalApproved, CLMRecoveryPaidbyM,
          						CASTreatReason1, InsCode, PLNCode, CASStatus, MCDays 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlstr17 = @l_sqlstr17 + @strAppend
			       exec  sp_executesql  @l_sqlstr17
		end

	if @strAppendCase = 18

		begin
         	      		set nocount on
			       declare @l_sqlstr18 as nvarchar(2000)
		                    declare @l_sqlgroupby18 as nvarchar(500)

		                    set  @l_sqlstr18 =  N'SELECT PAYCode, MBMAgentCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) 
              					AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
              					From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby18 = 'GROUP BY PAYCode, MBMAgentCode'
			       set @l_sqlstr18 = @l_sqlstr18 + @strAppend + @l_sqlgroupby18
			       exec  sp_executesql  @l_sqlstr18
		end

	if @strAppendCase = 19

		begin
         	      		set nocount on
			       declare @l_sqlstr19 as nvarchar(2000)
		                    declare @l_sqlgroupby19 as nvarchar(500)

		                    set  @l_sqlstr19 =  N'SELECT Top 50 PAYCode, MBMAgentCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
              					AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
              					From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby19 = 'GROUP BY PAYCode, MBMAgentCode 
							ORDER BY TotalCases Desc'
			       set @l_sqlstr19 = @l_sqlstr19 + @strAppend + @l_sqlgroupby19
			       exec  sp_executesql  @l_sqlstr19
		end

	if @strAppendCase = 20

		begin
         	      		set nocount on
			       declare @l_sqlstr20 as nvarchar(2000)
		                    declare @l_sqlgroupby20 as nvarchar(500)

		                    set  @l_sqlstr20 =  N'SELECT INSCode, PLNCode, Ageband, COUNT(CASNumber) AS TotalCases,
          						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby20 = 'GROUP BY INSCode, PLNCode, Ageband'
			       set @l_sqlstr20 = @l_sqlstr20 + @strAppend + @l_sqlgroupby20
			       exec  sp_executesql  @l_sqlstr20
		end

	if @strAppendCase = 21

		begin
         	      		set nocount on
			       declare @l_sqlstr21 as nvarchar(2000)
		                    declare @l_sqlgroupby21 as nvarchar(500)

		                    set  @l_sqlstr21 =  N'SELECT RACCode, Ageband, COUNT(CASNumber) AS TotalCases,
          						SUM(CLMTotalActual) AS TotalActualAmt,
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby21 = 'GROUP BY RACCode, Ageband'
			       set @l_sqlstr21 = @l_sqlstr21 + @strAppend + @l_sqlgroupby21
			       exec  sp_executesql  @l_sqlstr21
		end

	if @strAppendCase = 22

		begin
         	      		set nocount on
			       declare @l_sqlstr22 as nvarchar(2000)
		                    declare @l_sqlgroupby22 as nvarchar(500)

		                    set  @l_sqlstr22 =  N'SELECT ClinicGRPName, COUNT(CASNumber) AS TotalCases, 
          						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby22 = 'GROUP BY ClinicGRPName ORDER BY ClinicGRPName'
			       set @l_sqlstr22 = @l_sqlstr22 + @strAppend + @l_sqlgroupby22
			       exec  sp_executesql  @l_sqlstr22
		end

	if @strAppendCase = 23

		begin
         	      		set nocount on
			       declare @l_sqlstr23 as nvarchar(2000)
		                    declare @l_sqlgroupby23 as nvarchar(500)

		                    set  @l_sqlstr23 =  N'SELECT COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved,
          						Sum(CLMTotalActual) As TotalActual, ClinicCode, CLNName
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby23 = 'GROUP BY CASTreatReason1, ClinicCode, CLNName'
			       set @l_sqlstr23 = @l_sqlstr23 + @strAppend + @l_sqlgroupby23
			       exec  sp_executesql  @l_sqlstr23
		end

	if @strAppendCase = 24

		begin
         	      		set nocount on
			       declare @l_sqlstr24 as nvarchar(2000)
		                    declare @l_sqlgroupby24 as nvarchar(500)

		                    set  @l_sqlstr24 =  N'SELECT COUNT(CASNumber) AS Total, CLNType, Sum(CLMTotalApproved) As TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby24 = 'GROUP BY CLNType'
			       set @l_sqlstr24 = @l_sqlstr24 + @strAppend + @l_sqlgroupby24
			       exec  sp_executesql  @l_sqlstr24
		end

	if @strAppendCase = 25

		begin
         	      		set nocount on
			       declare @l_sqlstr25 as nvarchar(2000)

		                    set  @l_sqlstr25 =  N'SELECT PAYCode, ClinicCode, CLNName, CASNumber, MBMNumber, CoverID,
          						PatientName, CLMTotalActual, CLMTotalApproved,
          						CLMRecoveryPaidbyM, DateVisit, 
          						CASStatus, Claimability, CASTreatReason1, MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlstr25 = @l_sqlstr25 + @strAppend
			       exec  sp_executesql  @l_sqlstr25
		end

	if @strAppendCase = 26

		begin
         	      		set nocount on
			       declare @l_sqlstr26 as nvarchar(2000)
		                    declare @l_sqlgroupby26 as nvarchar(500)

		                    set  @l_sqlstr26 =  N'SELECT PAYCode, ClinicCode, CLNName, COUNT(CLNName) AS Total, 
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby26 = 'GROUP BY PAYCode, ClinicCode, CLNName'
			       set @l_sqlstr26 = @l_sqlstr26 + @strAppend + @l_sqlgroupby26
			       exec  sp_executesql  @l_sqlstr26
		end

	if @strAppendCase = 27

		begin
         	      		set nocount on
			       declare @l_sqlstr27 as nvarchar(2000)
		                    declare @l_sqlgroupby27 as nvarchar(500)

		                    set  @l_sqlstr27 =  N'SELECT ClinicCode, CLNName, CLNType, CLNLocation, SUM(MCDays) AS MCDays, COUNT(CLNName) AS Total,
          						SUM(MCStatus) AS MCTimes, SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby27 = 'GROUP BY ClinicCode, CLNName, CLNType, CLNLocation'
			       set @l_sqlstr27 = @l_sqlstr27 + @strAppend + @l_sqlgroupby27
			       exec  sp_executesql  @l_sqlstr27
		end

	if @strAppendCase = 28

		begin
         	      		set nocount on
			       declare @l_sqlstr28 as nvarchar(2000)
		                    declare @l_sqlgroupby28 as nvarchar(500)

		                    set  @l_sqlstr28 =  N'SELECT MBMGroupCompany, ClinicCode, CLNName, COUNT(CLNName) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby28 = 'GROUP BY MBMGroupCompany, ClinicCode, CLNName'
			       set @l_sqlstr28 = @l_sqlstr28 + @strAppend + @l_sqlgroupby28
			       exec  sp_executesql  @l_sqlstr28
		end

	if @strAppendCase = 29

		begin
         	      		set nocount on
			       declare @l_sqlstr29 as nvarchar(2000)
		                    declare @l_sqlgroupby29 as nvarchar(500)

		                    set  @l_sqlstr29 =  N'SELECT TOP 20 PAYCode, ClinicCode, CLNName, COUNT(CLNName) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where  0 = 0'
			       set @l_sqlgroupby29 = 'GROUP BY PAYCode, ClinicCode, CLNName ORDER BY Total DESC'
			       set @l_sqlstr29 = @l_sqlstr29 + @strAppend + @l_sqlgroupby29
			       exec  sp_executesql  @l_sqlstr29
		end

	if @strAppendCase = 30

		begin
         	      		set nocount on
			       declare @l_sqlstr30 as nvarchar(2000)
		                    declare @l_sqlgroupby30 as nvarchar(500)

		                    set  @l_sqlstr30 =  N'SELECT TOP 20 PAYCode, ClinicCode, CLNName, COUNT(CLNName) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where  0 = 0'
			       set @l_sqlgroupby30 = 'GROUP BY PAYCode, ClinicCode, CLNName ORDER BY TotalActual DESC'
			       set @l_sqlstr30 = @l_sqlstr30 + @strAppend + @l_sqlgroupby30
			       exec  sp_executesql  @l_sqlstr30
		end

	if @strAppendCase = 31

		begin
         	      		set nocount on
			       declare @l_sqlstr31 as nvarchar(2000)
		                    declare @l_sqlgroupby31 as nvarchar(500)

		                    set  @l_sqlstr31 =  N'SELECT CLNType, Count(CASNumber) AS TotalClaimP, SUM(CLMTotalApproved) AS TotalPanelAmt
                           				From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby31 = 'GROUP BY CLNType'
			       set @l_sqlstr31 = @l_sqlstr31 + @strAppend + @l_sqlgroupby31
			       exec  sp_executesql  @l_sqlstr31
		end

	if @strAppendCase = 32

		begin
         	      		set nocount on
			       declare @l_sqlstr32 as nvarchar(2000)
		                    declare @l_sqlgroupby32 as nvarchar(500)

		                    set  @l_sqlstr32 =  N'SELECT CLNType, Count(CASNumber) AS TotalClaimNP, SUM(CLMTotalApproved) AS TotalNonPanelAmt
                          					From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby32 = 'GROUP BY CLNType'
			       set @l_sqlstr32 = @l_sqlstr32 + @strAppend + @l_sqlgroupby32
			       exec  sp_executesql  @l_sqlstr32
		end

	if @strAppendCase = 33

		begin
         	      		set nocount on
			       declare @l_sqlstr33 as nvarchar(2000)
		                    declare @l_sqlgroupby33 as nvarchar(500)

		                    set  @l_sqlstr33 =  N'SELECT PLNCode, SUM(CLMTotalApproved) AS TotalNonPanelAmt
                  					From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby33 = 'GROUP BY PLNCode'
			       set @l_sqlstr33 = @l_sqlstr33 + @strAppend + @l_sqlgroupby33
			       exec  sp_executesql  @l_sqlstr33
		end

	if @strAppendCase = 34

		begin
         	      		set nocount on
			       declare @l_sqlstr34 as nvarchar(2000)

		                    set  @l_sqlstr34 =  N'SELECT CASNumber, MBMNumber, MBMPolicyNo, 
          						PatientName, DateVisit, ClinicCode, CLNName, CASTreatReason1, MCDays 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlstr34 = @l_sqlstr34 + @strAppend
			       exec  sp_executesql  @l_sqlstr34
		end

	if @strAppendCase = 35

		begin
         	      		set nocount on
			       declare @l_sqlstr35 as nvarchar(2000)
		                    declare @l_sqlgroupby35 as nvarchar(500)

		                    set  @l_sqlstr35 =  N'SELECT Distinct PAYCode From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby35 = 'ORDER BY PAYCode'
			       set @l_sqlstr35 = @l_sqlstr35 + @strAppend + @l_sqlgroupby35
			       exec  sp_executesql  @l_sqlstr35
		end

	if @strAppendCase = 36

		begin
         	      		set nocount on
			       declare @l_sqlstr36 as nvarchar(2000)
		                    declare @l_sqlgroupby36 as nvarchar(500)

		                    set  @l_sqlstr36 =  N'SELECT Distinct PAYCode From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby36 = 'Order by PAYCode'
			       set @l_sqlstr36 = @l_sqlstr36 + @strAppend + @l_sqlgroupby36
			       exec  sp_executesql  @l_sqlstr36
		end

	if @strAppendCase = 37

		begin
         	      		set nocount on
			       declare @l_sqlstr37 as nvarchar(2000)
		                    declare @l_sqlgroupby37 as nvarchar(500)

		                    set  @l_sqlstr37 =  N'SELECT MBMDepartment, CLNName, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby37 = 'Group By MBMDepartment, CLNName Order By CLNName, MBMDepartment'
			       set @l_sqlstr37 = @l_sqlstr37 + @strAppend + @l_sqlgroupby37
			       exec  sp_executesql  @l_sqlstr37
		end

	if @strAppendCase = 38

		begin
         	      		set nocount on
			       declare @l_sqlstr38 as nvarchar(2000)
		                    declare @l_sqlgroupby38 as nvarchar(500)

		                    set  @l_sqlstr38 =  N'SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, 
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby38 = 'Group By MBMDepartment Order By MBMDepartment'
			       set @l_sqlstr38 = @l_sqlstr38 + @strAppend + @l_sqlgroupby38
			       exec  sp_executesql  @l_sqlstr38
		end

	if @strAppendCase = 39

		begin
         	      		set nocount on
			       declare @l_sqlstr39 as nvarchar(2000)
		                    declare @l_sqlgroupby39 as nvarchar(500)

		                    set  @l_sqlstr39 =  N'SELECT MBMDepartment, MBMNumber, MBMName, MBMCName, CoverID, 
       						MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby39 = 'GROUP BY MBMDepartment, MBMNumber, MBMName, MBMCName, CoverID, 
        						MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit ORDER BY MBMNumber, MBMDepartment'
			       set @l_sqlstr39 = @l_sqlstr39 + @strAppend + @l_sqlgroupby39
			       exec  sp_executesql  @l_sqlstr39
		end

	if @strAppendCase = 40

		begin
         	      		set nocount on
			       declare @l_sqlstr40 as nvarchar(2000)
		                    declare @l_sqlgroupby40 as nvarchar(500)

		                    set  @l_sqlstr40 =  N'SELECT MBMDepartment, ReasonDesc, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
      						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
         						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby40 = 'Group By MBMDepartment, ReasonDesc ORDER BY ReasonDesc, MBMDepartment'
			       set @l_sqlstr40 = @l_sqlstr40 + @strAppend + @l_sqlgroupby40
			       exec  sp_executesql  @l_sqlstr40
		end

	if @strAppendCase = 41

		begin
         	      		set nocount on
			       declare @l_sqlstr41 as nvarchar(2000)
		                    declare @l_sqlgroupby41 as nvarchar(500)

		                    set  @l_sqlstr41 =  N'SELECT MBMDepartment, Count(CASNumber) AS NoofMC, MCDays,
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby41 = 'Group by MBMDepartment, MCDays'
			       set @l_sqlstr41 = @l_sqlstr41 + @strAppend + @l_sqlgroupby41
			       exec  sp_executesql  @l_sqlstr41
		end

	if @strAppendCase = 42

		begin
         	      		set nocount on
			       declare @l_sqlstr42 as nvarchar(2000)
		                    declare @l_sqlgroupby42 as nvarchar(500)

		                    set  @l_sqlstr42 =  N'SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, 
        						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby42 = 'Group By MBMDepartment Order By MCDays Desc'
			       set @l_sqlstr42 = @l_sqlstr42 + @strAppend + @l_sqlgroupby42
			       exec  sp_executesql  @l_sqlstr42
		end

	if @strAppendCase = 43

		begin
         	      		set nocount on
			       declare @l_sqlstr43 as nvarchar(2000)
		                    declare @l_sqlgroupby43 as nvarchar(500)

		                    set  @l_sqlstr43 =  N'SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
        						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby43 = 'Group By MBMDepartment Order By MCTimes Desc'
			       set @l_sqlstr43 = @l_sqlstr43 + @strAppend + @l_sqlgroupby43
			       exec  sp_executesql  @l_sqlstr43
		end

	if @strAppendCase = 44

		begin
         	      		set nocount on
			       declare @l_sqlstr44 as nvarchar(2000)
		                    declare @l_sqlgroupby44 as nvarchar(500)

			        set  @l_sqlstr44 =  N'SELECT MBMDepartment, MBMNumber, CoverID, PatientName, AnnLimit, AnnLmtIND, Count(CASNumber) AS TotalClaim,
          						SUM(MCStatus) as MCTimes, SUM(MCDays) AS MCDays,  
          						SUM(CLMTotalActual) as ApprovedAmt, MBMCLNBalLmt ,Sum(TopupAmt) AS TopupAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby44 = 'GROUP BY MBMDepartment, MBMNumber, CoverID, PatientName, AnnLimit, AnnLmtIND,  MBMCLNBalLmt
							ORDER BY MBMNumber'
			       set @l_sqlstr44 = @l_sqlstr44 + @strAppend + @l_sqlgroupby44
			       exec  sp_executesql  @l_sqlstr44
		end

	if @strAppendCase = 45

		begin
         	      		set nocount on
			       declare @l_sqlstr45 as nvarchar(2000)
		                    declare @l_sqlgroupby45 as nvarchar(500)

		                    set  @l_sqlstr45 =  N'SELECT MBMDepartment, COUNT(CASNumber) as TotalCases, SUM(CLMTotalActual) AS totalActual,
      						SUM(CLMTotalApproved) As TotalApproved From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby45 = 'Group By MBMDepartment'
			       set @l_sqlstr45 = @l_sqlstr45 + @strAppend + @l_sqlgroupby45
			       exec  sp_executesql  @l_sqlstr45
		end

	if @strAppendCase = 46

		begin
         	      		set nocount on
			       declare @l_sqlstr46 as nvarchar(2000)
		                    declare @l_sqlgroupby46 as nvarchar(500)

		                    set  @l_sqlstr46 =  N'SELECT MBMDepartment, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
        						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
      						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby46 = 'Group By MBMDepartment Order By Visit Desc'
			       set @l_sqlstr46 = @l_sqlstr46 + @strAppend + @l_sqlgroupby46
			       exec  sp_executesql  @l_sqlstr46
		end

	if @strAppendCase = 47

		begin
         	      		set nocount on
			       declare @l_sqlstr47 as nvarchar(2000)
		                    declare @l_sqlgroupby47 as nvarchar(500)

		                    set  @l_sqlstr47 =  N'SELECT PAYCode, MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, COUNT(CASNumber) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, 
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby47 = 'GROUP BY PAYCode, MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate) '
			       set @l_sqlstr47 = @l_sqlstr47 + @strAppend + @l_sqlgroupby47
			       exec  sp_executesql  @l_sqlstr47
		end

	if @strAppendCase = 48

		begin
         	      		set nocount on
			       declare @l_sqlstr48 as nvarchar(2000)
		                    declare @l_sqlgroupby48 as nvarchar(500)

		                    set  @l_sqlstr48 =  N'SELECT MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, COUNT(CASNumber) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby48 = 'GROUP BY MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)'
			       set @l_sqlstr48 = @l_sqlstr48 + @strAppend + @l_sqlgroupby48
			       exec  sp_executesql  @l_sqlstr48
		end

	if @strAppendCase = 49

		begin
         	      		set nocount on
			       declare @l_sqlstr49 as nvarchar(2000)
		                    declare @l_sqlgroupby49 as nvarchar(500)

		                    set  @l_sqlstr49 =  N'SELECT PAYCode, MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, COUNT(CASNumber) AS Total, 
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby49 = 'GROUP BY PAYCode, MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)'
			       set @l_sqlstr49 = @l_sqlstr49 + @strAppend + @l_sqlgroupby49
			       exec  sp_executesql  @l_sqlstr49
		end

	if @strAppendCase = 50

		begin
         	      		set nocount on
			       declare @l_sqlstr50 as nvarchar(2000)
		                    declare @l_sqlgroupby50 as nvarchar(500)

		                    set  @l_sqlstr50 =  N'SELECT MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, COUNT(CASNumber) AS Total, 
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, 
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby50 = 'GROUP BY MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)'
			       set @l_sqlstr50 = @l_sqlstr50 + @strAppend + @l_sqlgroupby50
			       exec  sp_executesql  @l_sqlstr50
		end

	if @strAppendCase = 51

		begin
         	      		set nocount on
			       declare @l_sqlstr51 as nvarchar(2000)
		                    declare @l_sqlgroupby51 as nvarchar(500)

		                    set  @l_sqlstr51 =  N'SELECT PAYCode, MONTH(DateVisit) AS MONTH, YEAR(DateVisit) AS YEAR, COUNT(CASNumber) AS Total, 
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby51 = 'GROUP BY PAYCode, MONTH(DateVisit), YEAR(DateVisit)'
			       set @l_sqlstr51 = @l_sqlstr51 + @strAppend + @l_sqlgroupby51
			       exec  sp_executesql  @l_sqlstr51
		end

	if @strAppendCase = 52

		begin
         	      		set nocount on
			       declare @l_sqlstr52 as nvarchar(2000)
		                    declare @l_sqlgroupby52 as nvarchar(500)

		                    set  @l_sqlstr52 =  N'SELECT MONTH(DateVisit) AS MONTH, YEAR(DateVisit) AS YEAR, COUNT(CASNumber) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby52 = 'GROUP BY MONTH(DateVisit), YEAR(DateVisit)'
			       set @l_sqlstr52 = @l_sqlstr52 + @strAppend + @l_sqlgroupby52
			       exec  sp_executesql  @l_sqlstr52
		end

	if @strAppendCase = 53

		begin
         	      		set nocount on
			       declare @l_sqlstr53 as nvarchar(2000)
		                    declare @l_sqlgroupby53 as nvarchar(500)

		                    set  @l_sqlstr53 =  N'SELECT PAYCode, MONTH(RegisterDate) AS MONTH, YEAR(RegisterDate) AS YEAR, COUNT(CASNumber) AS Total,
         						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby53 = 'GROUP BY PAYCode, MONTH(RegisterDate), YEAR(RegisterDate)'
			       set @l_sqlstr53 = @l_sqlstr53 + @strAppend + @l_sqlgroupby53
			       exec  sp_executesql  @l_sqlstr53
		end

	if @strAppendCase = 54

		begin
         	      		set nocount on
			       declare @l_sqlstr54 as nvarchar(2000)
		                    declare @l_sqlgroupby54 as nvarchar(500)

		                    set  @l_sqlstr54 =  N'SELECT MONTH(RegisterDate) AS MONTH, YEAR(RegisterDate) AS YEAR, COUNT(CASNumber) AS Total, 
        						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved, 
       						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby54 = 'GROUP BY MONTH(RegisterDate), YEAR(RegisterDate)'
			       set @l_sqlstr54 = @l_sqlstr54 + @strAppend + @l_sqlgroupby54
			       exec  sp_executesql  @l_sqlstr54
		end

	if @strAppendCase = 55

		begin
         	      		set nocount on
			       declare @l_sqlstr55 as nvarchar(2000)
		                    declare @l_sqlgroupby55 as nvarchar(500)

		                    set  @l_sqlstr55 =  N'SELECT MBMDepartment, MBMGroupCompany, CLNName, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby55 = 'Group By MBMDepartment, CLNName, MBMGroupCompany 
							ORDER BY CLNName, MBMGroupCompany'
			       set @l_sqlstr55 = @l_sqlstr55 + @strAppend + @l_sqlgroupby55
			       exec  sp_executesql  @l_sqlstr55
		end

	if @strAppendCase = 56

		begin
         	      		set nocount on
			       declare @l_sqlstr56 as nvarchar(2000)
		                    declare @l_sqlgroupby56 as nvarchar(500)

		                    set  @l_sqlstr56 =  N'SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes, 
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby56 = 'Group By MBMDepartment, MBMGroupCompany 
							ORDER BY MBMDepartment, MBMGroupCompany'
			       set @l_sqlstr56 = @l_sqlstr56 + @strAppend + @l_sqlgroupby56
			       exec  sp_executesql  @l_sqlstr56
		end

	if @strAppendCase = 57

		begin
         	      		set nocount on
			       declare @l_sqlstr57 as nvarchar(2000)

		                    set  @l_sqlstr57 =  N'SELECT PAYCode, MBMGroupCompany, CASNumber, MBMNumber, CoverID,
          						PatientName, CLMTotalActual, CLMTotalApproved, 
          						CLMRecoveryPaidbyM, DateVisit,
          						CASStatus, Claimability, CASTreatReason1, MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlstr57 = @l_sqlstr57 + @strAppend
			       exec  sp_executesql  @l_sqlstr57
		end

	if @strAppendCase = 58

		begin
         	      		set nocount on
			       declare @l_sqlstr58 as nvarchar(2000)
		                    declare @l_sqlgroupby58 as nvarchar(500)

		                    set  @l_sqlstr58 =  N'SELECT MBMDepartment, MBMGroupCompany, MBMNumber, MBMName, MBMCName, CoverID,
        						MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby58 = 'GROUP BY MBMDepartment, MBMGroupCompany, MBMNumber, MBMName, MBMCName, CoverID,
        							MBMEmployeeNo, AnnLimit, CLMTotalApproved, CLMBalanceLimit ORDER BY MBMNumber, MBMGroupCompany'
			       set @l_sqlstr58 = @l_sqlstr58 + @strAppend + @l_sqlgroupby58
			       exec  sp_executesql  @l_sqlstr58
		end

	if @strAppendCase = 59

		begin
         	      		set nocount on
			       declare @l_sqlstr59 as nvarchar(2000)
		                    declare @l_sqlgroupby59 as nvarchar(500)

		                    set  @l_sqlstr59 =  N'SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby59 = 'Group By MBMDepartment, MBMGroupCompany Order By MBMGroupCompany, MBMDepartment'
			       set @l_sqlstr59 = @l_sqlstr59 + @strAppend + @l_sqlgroupby59
			       exec  sp_executesql  @l_sqlstr59
		end

	if @strAppendCase = 60

		begin
         	      		set nocount on
			       declare @l_sqlstr60 as nvarchar(2000)
		                    declare @l_sqlgroupby60 as nvarchar(500)

		                    set  @l_sqlstr60 =  N'SELECT MBMDepartment, MBMGroupCompany, ReasonDesc, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
       						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby60 = 'Group By MBMDepartment, MBMGroupCompany, ReasonDesc ORDER BY ReasonDesc, MBMGroupCompany'
			       set @l_sqlstr60 = @l_sqlstr60 + @strAppend + @l_sqlgroupby60
			       exec  sp_executesql  @l_sqlstr60
		end

	if @strAppendCase = 61

		begin
         	      		set nocount on
			       declare @l_sqlstr61 as nvarchar(2000)
		                    declare @l_sqlgroupby61 as nvarchar(500)

		                    set  @l_sqlstr61 =  N'SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby61 = 'Group By MBMDepartment, MBMGroupCompany Order By MCDays Desc'
			       set @l_sqlstr61 = @l_sqlstr61 + @strAppend + @l_sqlgroupby61
			       exec  sp_executesql  @l_sqlstr61
		end

	if @strAppendCase = 62

		begin
         	      		set nocount on
			       declare @l_sqlstr62 as nvarchar(2000)
		                    declare @l_sqlgroupby62 as nvarchar(500)

		                    set  @l_sqlstr62 =  N'SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
          						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby62 = 'Group By MBMDepartment, MBMGroupCompany Order By MCTimes Desc'
			       set @l_sqlstr62 = @l_sqlstr62 + @strAppend + @l_sqlgroupby62
			       exec  sp_executesql  @l_sqlstr62
		end

	if @strAppendCase = 63

		begin
         	      		set nocount on
			       declare @l_sqlstr63 as nvarchar(2000)
		                    declare @l_sqlgroupby63 as nvarchar(500)

		                    set  @l_sqlstr63 =  N'SELECT PAYCode, MBMGroupCompany, COUNT(MBMGroupCompany) AS Total,
          						SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby63 = 'GROUP BY PAYCode, MBMGroupCompany'
			       set @l_sqlstr63 = @l_sqlstr63 + @strAppend + @l_sqlgroupby63
			       exec  sp_executesql  @l_sqlstr63
		end

	if @strAppendCase = 64

		begin
         	      		set nocount on
			       declare @l_sqlstr64 as nvarchar(2000)
		                    declare @l_sqlgroupby64 as nvarchar(500)

		                    set  @l_sqlstr64 =  N'SELECT MBMGroupCompany, COUNT(MBMGroupCompany) AS Total,
         						 SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby64 = 'GROUP BY MBMGroupCompany'
			       set @l_sqlstr64 = @l_sqlstr64 + @strAppend + @l_sqlgroupby64
			       exec  sp_executesql  @l_sqlstr64
		end

	if @strAppendCase = 65

		begin
         	      		set nocount on
			       declare @l_sqlstr65 as nvarchar(2000)
		                    declare @l_sqlgroupby65 as nvarchar(500)

		                    set  @l_sqlstr65 =  N'SELECT MBMDepartment, MBMGroupCompany, COUNT(CASNumber) as Visit, SUM(CONVERT(Numeric, MCStatus)) AS MCTimes,
        						SUM(MCDays) AS MCDays, SUM(CLMTotalActual) AS BillAmt, SUM(CLMTotalApproved) AS AppAmt
         						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby65 = 'GROUP BY MBMDepartment, MBMGroupCompany ORDER BY Visit Desc'
			       set @l_sqlstr65 = @l_sqlstr65 + @strAppend + @l_sqlgroupby65
			       exec  sp_executesql  @l_sqlstr65
		end

	if @strAppendCase = 66

		begin
         	      		set nocount on
			       declare @l_sqlstr66 as nvarchar(2000)
		                    declare @l_sqlgroupby66 as nvarchar(500)

		                    set  @l_sqlstr66 =  N'SELECT ICNo, Count(CASNumber) AS TotalClaim, PatientName, 
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt, SUM(CONVERT(Money,CLMPayableByMedix2M)) AS OSAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby66 = 'GROUP BY ICNo, PatientName'
			       set @l_sqlstr66 = @l_sqlstr66 + @strAppend + @l_sqlgroupby66
			       exec  sp_executesql  @l_sqlstr66
		end

	if @strAppendCase = 67

		begin
         	      		set nocount on
			       declare @l_sqlstr67 as nvarchar(2000)
		                    declare @l_sqlgroupby67 as nvarchar(500)

		                    set  @l_sqlstr67 =  N'SELECT PAYCode, INSCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
         						SUM(clmtotalapproved) As TotalApproved  
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby67 = 'Group By PAYCode, INSCode'
			       set @l_sqlstr67 = @l_sqlstr67 + @strAppend + @l_sqlgroupby67
			       exec  sp_executesql  @l_sqlstr67
		end

	if @strAppendCase = 68

		begin
         	      		set nocount on
			       declare @l_sqlstr68 as nvarchar(2000)
		                    declare @l_sqlgroupby68 as nvarchar(500)

		                    set  @l_sqlstr68 =  N'SELECT PAYCode, PLNCode, INSCode, COUNT(CASNumber) AS TotalCases,
         						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby68 = 'GROUP BY PAYCode, PLNCode, INSCode'
			       set @l_sqlstr68 = @l_sqlstr68 + @strAppend + @l_sqlgroupby68
			       exec  sp_executesql  @l_sqlstr68
		end

	if @strAppendCase = 69

		begin
         	      		set nocount on
			       declare @l_sqlstr69 as nvarchar(2000)
		                    declare @l_sqlgroupby69 as nvarchar(500)

		                    set  @l_sqlstr69 =  N'SELECT PAYCode, INSCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby69 = 'Group By PAYCode, INSCode, Sex'
			       set @l_sqlstr69 = @l_sqlstr69 + @strAppend + @l_sqlgroupby69
			       exec  sp_executesql  @l_sqlstr69
		end

	if @strAppendCase = 70

		begin
         	      		set nocount on
			       declare @l_sqlstr70 as nvarchar(2000)
		                    declare @l_sqlgroupby70 as nvarchar(500)

		                    set  @l_sqlstr70 =  N'SELECT PAYCode, INSCode, RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
        						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby70 = 'Group By PAYCode, INSCode, RACCode'
			       set @l_sqlstr70 = @l_sqlstr70 + @strAppend + @l_sqlgroupby70
			       exec  sp_executesql  @l_sqlstr70
		end

	if @strAppendCase = 71

		begin
         	      		set nocount on
			       declare @l_sqlstr71 as nvarchar(2000)
		                    declare @l_sqlgroupby71 as nvarchar(500)

		                    set  @l_sqlstr71 =  N'SELECT Count(CASNumber) AS NoofMC, MCDays,
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby71 = 'Group by MCDays'
			       set @l_sqlstr71 = @l_sqlstr71 + @strAppend + @l_sqlgroupby71
			       exec  sp_executesql  @l_sqlstr71
		end

	if @strAppendCase = 72

		begin
         	      		set nocount on
			       declare @l_sqlstr72 as nvarchar(2000)
		                    declare @l_sqlgroupby72 as nvarchar(500)

		                    set  @l_sqlstr72 =  N'SELECT ClinicCode, CLNName, Count(CASNumber) AS NoofMC, MCDays,
         						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims
          						From VIEW_CLMReport_Simple Where 0 = 0 '
			       set @l_sqlgroupby72 = 'Group by ClinicCode, CLNName, MCDays'
			       set @l_sqlstr72 = @l_sqlstr72 + @strAppend + @l_sqlgroupby72
			       exec  sp_executesql  @l_sqlstr72
		end

	if @strAppendCase = 73

		begin
         	      		set nocount on
			       declare @l_sqlstr73 as nvarchar(2000)
		                    declare @l_sqlgroupby73 as nvarchar(500)

		                    set  @l_sqlstr73 =  N'SELECT MBMGroupCompany, Count(CASNumber) AS NoofMC, MCDays, 
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby73 = 'Group by MBMGroupCompany, MCDays'
			       set @l_sqlstr73 = @l_sqlstr73 + @strAppend + @l_sqlgroupby73
			       exec  sp_executesql  @l_sqlstr73
		end

	if @strAppendCase = 74

		begin
         	      		set nocount on
			       declare @l_sqlstr74 as nvarchar(2000)
		                    declare @l_sqlgroupby74 as nvarchar(500)

		                    set  @l_sqlstr74 =  N'SELECT CASTreatReason1, Count(CASNumber) AS NoofMC, MCDays,
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS AmtClaims
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby74 = 'Group by CASTreatReason1, MCDays'
			       set @l_sqlstr74 = @l_sqlstr74 + @strAppend + @l_sqlgroupby74
			       exec  sp_executesql  @l_sqlstr74
		end

	if @strAppendCase = 75

		begin
         	      		set nocount on
			       declare @l_sqlstr75 as nvarchar(2000)
		                    declare @l_sqlgroupby75 as nvarchar(500)

		                    set  @l_sqlstr75 =  N'SELECT MBMNumber, COUNT(CASNumber) AS TotalClaim, 
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt, 
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby75 = 'GROUP BY MBMNumber'
			       set @l_sqlstr75 = @l_sqlstr75 + @strAppend + @l_sqlgroupby75
			       exec  sp_executesql  @l_sqlstr75
		end

	if @strAppendCase = 76

		begin
         	      		set nocount on
			       declare @l_sqlstr76 as nvarchar(2000)
		                    declare @l_sqlgroupby76 as nvarchar(500)

		                    set  @l_sqlstr76 =  N'SELECT MBMGroupCompany, MBMNumber, Count(CASNumber) AS TotalClaim,
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt,
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby76 = 'GROUP BY MBMGroupCompany, MBMNumber'
			       set @l_sqlstr76 = @l_sqlstr76 + @strAppend + @l_sqlgroupby76
			       exec  sp_executesql  @l_sqlstr76
		end

	if @strAppendCase = 77

		begin
         	      		set nocount on
			       declare @l_sqlstr77 as nvarchar(2000)
		                    declare @l_sqlgroupby77 as nvarchar(500)

		                    set  @l_sqlstr77 =  N'SELECT PAYCode, GRPCoCode, MBMNumber, PatientName, PLNCode, Count(CASNumber) AS TotalClaim,
    						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt,
       						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby77 = 'GROUP BY PAYCode, GRPCoCode, MBMNumber, PatientName, PLNCode'
			       set @l_sqlstr77 = @l_sqlstr77 + @strAppend + @l_sqlgroupby77
			       exec  sp_executesql  @l_sqlstr77
		end

	if @strAppendCase = 78

		begin
         	      		set nocount on
			       declare @l_sqlstr78 as nvarchar(2000)
		                    declare @l_sqlgroupby78 as nvarchar(500)

		                    set  @l_sqlstr78 =  N'SELECT MBMDepartment, MBMNumber, PatientName, CoverID, SUM(CLMTotalActual) AS ApprovedAmt, 
       						Count(CASNumber) AS TotalClaim, SUM(MCStatus) AS MCTimes, 
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0 '
			       set @l_sqlgroupby78 = 'GROUP BY MBMDepartment, MBMNumber, PatientName, CoverID'
			       set @l_sqlstr78 = @l_sqlstr78 + @strAppend + @l_sqlgroupby78
			       exec  sp_executesql  @l_sqlstr78
		end

	if @strAppendCase = 79

		begin
         	      		set nocount on
			       declare @l_sqlstr79 as nvarchar(2000)
		                    declare @l_sqlgroupby79 as nvarchar(500)

		                    set  @l_sqlstr79 =  N'SELECT MBMNumber, MBMName, COUNT(CASNumber) AS TotalClaim,
      						SUM(CONVERT(Numeric, MCStatus)) AS MCStatus, SUM(MCDays) AS MCDays, AnnLimit, 
          						SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby79 = 'GROUP BY MBMNumber, MBMName, AnnLimit'
			       set @l_sqlstr79 = @l_sqlstr79 + @strAppend + @l_sqlgroupby79
			       exec  sp_executesql  @l_sqlstr79
		end

	if @strAppendCase = 80

		begin
         	      		set nocount on
			       declare @l_sqlstr80 as nvarchar(2000)
		                    declare @l_sqlgroupby80 as nvarchar(500)

		                    set  @l_sqlstr80 =  N'SELECT PAYCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
          						SUM(clmtotalapproved) As TotalApproved From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby80 = 'Group By PAYCode'
			       set @l_sqlstr80 = @l_sqlstr80 + @strAppend + @l_sqlgroupby80
			       exec  sp_executesql  @l_sqlstr80
		end

	if @strAppendCase = 81

		begin
         	      		set nocount on
			       declare @l_sqlstr81 as nvarchar(2000)
		                    declare @l_sqlgroupby81 as nvarchar(500)

		                    set  @l_sqlstr81 =  N'SELECT PAYCode, PLNCode, Ageband, COUNT(CASNumber) AS TotalCases, 
          						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby81 = 'GROUP BY PAYCode, PLNCode, Ageband'
			       set @l_sqlstr81 = @l_sqlstr81 + @strAppend + @l_sqlgroupby81
			       exec  sp_executesql  @l_sqlstr81
		end

	if @strAppendCase = 82

		begin
         	      		set nocount on
			       declare @l_sqlstr82 as nvarchar(2000)
		                    declare @l_sqlgroupby82 as nvarchar(500)

		                    set  @l_sqlstr82 =  N'SELECT PLNCode, SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby82 = 'GROUP BY PLNCode'
			       set @l_sqlstr82 = @l_sqlstr82 + @strAppend + @l_sqlgroupby82
			       exec  sp_executesql  @l_sqlstr82
		end

	if @strAppendCase = 83

		begin
         	      		set nocount on
			       declare @l_sqlstr83 as nvarchar(2000)
		                    declare @l_sqlgroupby83 as nvarchar(500)

		                    set  @l_sqlstr83 =  N'SELECT PAYCode, PLNCode, INSCode, Ageband, COUNT(CASNumber) AS TotalCases, 
          						SUM(CLMTotalActual) AS TotalActualAmt, 
          						SUM(CLMTotalApproved) AS TotalApprovedAmt 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby83 = 'GROUP BY PAYCode, PLNCode, INSCode, Ageband'
			       set @l_sqlstr83 = @l_sqlstr83 + @strAppend + @l_sqlgroupby83
			       exec  sp_executesql  @l_sqlstr83
		end

	if @strAppendCase = 84

		begin
         	      		set nocount on
			       declare @l_sqlstr84 as nvarchar(2000)
		                    declare @l_sqlgroupby84 as nvarchar(500)

		                    set  @l_sqlstr84 =  N'SELECT PAYCode, PLNCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
       						SUM(clmtotalapproved) As TotalApproved
        						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby84 = 'Group By PAYCode, PLNCode'
			       set @l_sqlstr84 = @l_sqlstr84 + @strAppend + @l_sqlgroupby84
			       exec  sp_executesql  @l_sqlstr84
		end

	if @strAppendCase = 85

		begin
         	      		set nocount on
			       declare @l_sqlstr85 as nvarchar(2000)
		                    declare @l_sqlgroupby85 as nvarchar(500)

		                    set  @l_sqlstr85 =  N'SELECT PAYCode, PLNCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
          						SUM(clmtotalapproved) As TotalApproved, 
          						SUM(MBMCLNBasicPrem) AS BasicPremium, 
          						SUM(MBMCLNPremium) AS Premium From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby85 = 'Group By PAYCode, PLNCode'
			       set @l_sqlstr85 = @l_sqlstr85 + @strAppend + @l_sqlgroupby85
			       exec  sp_executesql  @l_sqlstr85
		end

	if @strAppendCase = 86

		begin
         	      		set nocount on
			       declare @l_sqlstr86 as nvarchar(2000)
		                    declare @l_sqlgroupby86 as nvarchar(500)

		                    set  @l_sqlstr86 =  N'SELECT Top 10 PAYCode, PLNCode, CASTreatReason1, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, 
          						SUM(clmtotalapproved) As TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby86 = 'Group By PAYCode, PLNCode, CASTreatReason1 ORDER BY TotalCases Desc'
			       set @l_sqlstr86 = @l_sqlstr86 + @strAppend + @l_sqlgroupby86
			       exec  sp_executesql  @l_sqlstr86
		end

	if @strAppendCase = 87

		begin
         	      		set nocount on
			       declare @l_sqlstr87 as nvarchar(2000)
		                    declare @l_sqlgroupby87 as nvarchar(500)

		                    set  @l_sqlstr87 =  N'SELECT Top 10 PAYCode, PLNCode, ClinicCode, CLNName, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
      						SUM(clmtotalapproved) As TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby87 = 'Group By PAYCode, PLNCode, ClinicCode, CLNName ORDER BY TotalCases Desc'
			       set @l_sqlstr87 = @l_sqlstr87 + @strAppend + @l_sqlgroupby87
			       exec  sp_executesql  @l_sqlstr87
		end

	if @strAppendCase = 88

		begin
         	      		set nocount on
			       declare @l_sqlstr88 as nvarchar(2000)
		                    declare @l_sqlgroupby88 as nvarchar(500)

		                    set  @l_sqlstr88 =  N'SELECT PAYCode, MBMPolicyNo, MBMGroupCompany,
          						COUNT(CASNumber) AS TotalCases, 
          						SUM(CLMTotalActual) AS TotalActualAmt,
          						SUM(CLMTotalApproved) AS TotalApprovedAmt
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby88 = 'GROUP BY PAYCode, MBMPolicyNo, MBMGroupCompany'
			       set @l_sqlstr88 = @l_sqlstr88 + @strAppend + @l_sqlgroupby88
			       exec  sp_executesql  @l_sqlstr88
		end

	if @strAppendCase = 89

		begin
         	      		set nocount on
			       declare @l_sqlstr89 as nvarchar(2000)
		                    declare @l_sqlgroupby89 as nvarchar(500)

		                    set  @l_sqlstr89 =  N'SELECT RACCode, Ageband, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
         						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby89 = 'Group By RACCode, Ageband'
			       set @l_sqlstr89 = @l_sqlstr89 + @strAppend + @l_sqlgroupby89
			       exec  sp_executesql  @l_sqlstr89
		end

	if @strAppendCase = 90

		begin
         	      		set nocount on
			       declare @l_sqlstr90 as nvarchar(2000)
		                    declare @l_sqlgroupby90 as nvarchar(500)

		                    set  @l_sqlstr90 =  N'SELECT RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
         						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved 
      						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby90 = 'Group By RACCode'
			       set @l_sqlstr90 = @l_sqlstr90 + @strAppend + @l_sqlgroupby90
			       exec  sp_executesql  @l_sqlstr90
		end

	if @strAppendCase = 91

		begin
         	      		set nocount on
			       declare @l_sqlstr91 as nvarchar(2000)
		                    declare @l_sqlgroupby91 as nvarchar(500)

		                    set  @l_sqlstr91 =  N'SELECT INSCode, PLNCode, RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
      						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby91 = 'Group By INSCode, PLNCode, RACCode'
			       set @l_sqlstr91 = @l_sqlstr91 + @strAppend + @l_sqlgroupby91
			       exec  sp_executesql  @l_sqlstr91

		end

	if @strAppendCase = 92

		begin
         	      		set nocount on
			       declare @l_sqlstr92 as nvarchar(2000)
		                    declare @l_sqlgroupby92 as nvarchar(500)

		                    set  @l_sqlstr92 =  N'SELECT RACCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby92 = 'Group By RACCode, Sex'
			       set @l_sqlstr92 = @l_sqlstr92 + @strAppend + @l_sqlgroupby92
			       exec  sp_executesql  @l_sqlstr92

		end

	if @strAppendCase = 93

		begin
         	      		set nocount on
			       declare @l_sqlstr93 as nvarchar(2000)

		                    set  @l_sqlstr93 =  N'SELECT PAYCode, CASNumber, MBMNumber, MBMPolicyNo, 
          						CoverID, MBMName, MBMDOB, DOB, PatientName, MBMSex, AgeCode, Sex, Ageband, 
         						INSCode, PLNCode, ClinicCode, CLNName, CASTreatReason1, 
          						CLMTotalActual, CLMTotalApproved, MCDays, ReasonDesc,
          						CLMPayablebyMedix2M, BordxRefNo, BordxDate, ClaimOpenDate, 
          						Claimability, CASStatus, MBMDataStatus, MBMStatus, DateVisit,
          						MBMPayorEffDate, MBMPayorExpDate, MBMCLNPremium, 
          						RegisterDate, MBMIcBcPp, ICNo, CASDiagnosis5, MBMDepartment, INVPayTo, 
          						CLMNonGpFee, CLNType From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlstr93 = @l_sqlstr93 + @strAppend
			       exec  sp_executesql  @l_sqlstr93
		end

	if @strAppendCase = 94

		begin
         	      		set nocount on
			       declare @l_sqlstr94 as nvarchar(2000)

		                    set  @l_sqlstr94 =  N'SELECT PAYCode, CASNumber, MBMNumber, MBMPolicyNo,
         						CoverID, MBMName, MBMDOB, DOB, PatientName, MBMSex, AgeCode, Sex, Ageband, 
          						INSCode, PLNCode, ClinicCode, CLNName, CASTreatReason1, ReasonDesc,
          						CLMTotalActual, CLMTotalApproved, Remarks, MCDays,
          						CLMPayablebyMedix2M, BordxRefNo, BordxDate, ClaimOpenDate,  
          						Claimability, CASStatus, MBMDataStatus, MBMStatus, DateVisit, 
        						MBMPayorEffDate, MBMPayorExpDate, MBMCLNPremium,
          						RegisterDate, MBMIcBcPp, ICNo, CASDiagnosis5 From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlstr94 = @l_sqlstr94 + @strAppend
			       exec  sp_executesql  @l_sqlstr94
		end

	if @strAppendCase = 95

		begin
         	      		set nocount on
			       declare @l_sqlstr95 as nvarchar(2000)
		                    declare @l_sqlgroupby95 as nvarchar(500)

		                    set  @l_sqlstr95 =  N'SELECT Count(CASNumber) AS NoofVisits, MBMNumber, 
          						PatientName, PLNCode, SUM(CLMTotalApproved) AS AmtClaims,
          						ICNo From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby95 = 'Group by MBMNumber, ICNo, PatientName, PLNCode'
			       set @l_sqlstr95 = @l_sqlstr95 + @strAppend + @l_sqlgroupby95
			       exec  sp_executesql  @l_sqlstr95
		end

	if @strAppendCase = 96

		begin
         	      		set nocount on
			       declare @l_sqlstr96 as nvarchar(2000)
		                    declare @l_sqlgroupby96 as nvarchar(500)

		                    set  @l_sqlstr96 =  N'SELECT PAYCode, DateVisit, Count(CASNumber) AS Total From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby96 = 'GROUP BY PAYCode, DateVisit ORDER BY DateVisit'
			       set @l_sqlstr96 = @l_sqlstr96 + @strAppend + @l_sqlgroupby96
			       exec  sp_executesql  @l_sqlstr96
		end

	if @strAppendCase = 97

		begin
         	      		set nocount on
			       declare @l_sqlstr97 as nvarchar(2000)
		                    declare @l_sqlgroupby97 as nvarchar(500)

		                    set  @l_sqlstr97 =  N'SELECT PAYCode, Month(DateVisit) As Month,
          						COUNT(CASNumber) AS TotalCases
          						FROM VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby97 = 'GROUP BY PAYCode, MONTH(DateVisit), YEAR(DateVisit) ORDER BY MONTH(DateVisit)'
			       set @l_sqlstr97 = @l_sqlstr97 + @strAppend + @l_sqlgroupby97
			       exec  sp_executesql  @l_sqlstr97
		end

	if @strAppendCase = 98

		begin
         	      		set nocount on
			       declare @l_sqlstr98 as nvarchar(2000)
		                    declare @l_sqlgroupby98 as nvarchar(500)

		                    set  @l_sqlstr98 =  N'SELECT PAYCode, STACode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual,
        						SUM(clmtotalapproved) As TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby98 = 'GROUP BY PAYCode, STACode'
			       set @l_sqlstr98 = @l_sqlstr98 + @strAppend + @l_sqlgroupby98
			       exec  sp_executesql  @l_sqlstr98
		end

	if @strAppendCase = 99

		begin
         	      		set nocount on
			       declare @l_sqlstr99 as nvarchar(2000)
		                    declare @l_sqlgroupby99 as nvarchar(500)

		                    set  @l_sqlstr99 =  N'SELECT SEX, Ageband, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) 
         						 AS totalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby99 = 'Group By SEX, Ageband'
			       set @l_sqlstr99 = @l_sqlstr99 + @strAppend + @l_sqlgroupby99
			       exec  sp_executesql  @l_sqlstr99
		end

	if @strAppendCase = 100

		begin
         	      		set nocount on
			       declare @l_sqlstr100 as nvarchar(2000)
		                    declare @l_sqlgroupby100 as nvarchar(500)

		                    set  @l_sqlstr100 =  N'SELECT SEX, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby100 = 'Group By SEX'
			       set @l_sqlstr100 = @l_sqlstr100 + @strAppend + @l_sqlgroupby100
			       exec  sp_executesql  @l_sqlstr100
		end

	if @strAppendCase = 101

		begin
         	      		set nocount on
			       declare @l_sqlstr101 as nvarchar(2000)
		                    declare @l_sqlgroupby101 as nvarchar(500)

		                    set  @l_sqlstr101 =  N'SELECT INSCode, PLNCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						 AS totalActual, SUM(CLMTotalApproved) AS TotalApproved 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby101 = 'Group By INSCode, PLNCode, Sex'
			       set @l_sqlstr101 = @l_sqlstr101 + @strAppend + @l_sqlgroupby101
			       exec  sp_executesql  @l_sqlstr101
		end

	if @strAppendCase = 102

		begin
         	      		set nocount on
			       declare @l_sqlstr102 as nvarchar(2000)
		                    declare @l_sqlgroupby102 as nvarchar(500)

		                    set  @l_sqlstr102 =  N'SELECT RACCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)
          						 AS totalActual, SUM(CLMTotalApproved) AS TotalApproved
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby102 = 'Group By RACCode, Sex'
			       set @l_sqlstr102 = @l_sqlstr102 + @strAppend + @l_sqlgroupby102
			       exec  sp_executesql  @l_sqlstr102
		end

	if @strAppendCase = 103

		begin
         	      		set nocount on
			       declare @l_sqlstr103 as nvarchar(2000)
		                    declare @l_sqlgroupby103 as nvarchar(500)

		                    set  @l_sqlstr103 =  N'SELECT COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved,
          						Sum(CLMTotalActual) As TotalActual, ClinicCode, CLNName, 
          						SUM(MCDays) AS MCDays From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby103 = 'GROUP BY CASTreatReason1, ClinicCode, CLNName'
			       set @l_sqlstr103 = @l_sqlstr103 + @strAppend + @l_sqlgroupby103
			       exec  sp_executesql  @l_sqlstr103
		end

	if @strAppendCase = 104

		begin
         	      		set nocount on
			       declare @l_sqlstr104 as nvarchar(2000)
		                    declare @l_sqlgroupby104 as nvarchar(500)

		                    set  @l_sqlstr104 =  N'SELECT Top 20 COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved,
              					Sum(CLMTotalActual) As TotalActual, SUM(MCDays) AS MCDays
              					From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby104 = 'GROUP BY CASTreatReason1 ORDER BY Total DESC'
			       set @l_sqlstr104 = @l_sqlstr104 + @strAppend + @l_sqlgroupby104
			       exec  sp_executesql  @l_sqlstr104
		end

	if @strAppendCase = 105

		begin
         	      		set nocount on
			       declare @l_sqlstr105 as nvarchar(2000)
		                    declare @l_sqlgroupby105 as nvarchar(500)

		                    set  @l_sqlstr105 =  N'SELECT Top 20 Sum(CLMTotalApproved) As TotalApproved, COUNT(CASTreatReason1) AS Total, CASTreatReason1, 
           						Sum(CLMTotalActual) As TotalActual, SUM(MCDays) AS MCDays 
             						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby105 = 'Group By CASTreatReason1 Order By TotalApproved Desc'
			       set @l_sqlstr105 = @l_sqlstr105 + @strAppend + @l_sqlgroupby105
			       exec  sp_executesql  @l_sqlstr105
		end

	if @strAppendCase = 106

		begin
         	      		set nocount on
			       declare @l_sqlstr106 as nvarchar(2000)
		                    declare @l_sqlgroupby106 as nvarchar(500)

		                    set  @l_sqlstr106 =  N'SELECT COUNT(CASTreatReason1) AS Total, CASTreatReason1, Sum(CLMTotalApproved) As TotalApproved, 
            						Sum(CLMTotalActual) As TotalActual, SUM(MCDays) AS MCDays
              					From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby106 = 'GROUP BY CASTreatReason1 ORDER BY Total DESC'
			       set @l_sqlstr106 = @l_sqlstr106 + @strAppend + @l_sqlgroupby106
			       exec  sp_executesql  @l_sqlstr106
		end

	if @strAppendCase = 107

		begin
         	      		set nocount on
			       declare @l_sqlstr107 as nvarchar(2000)
		                    declare @l_sqlgroupby107 as nvarchar(500)

		                   set  @l_sqlstr107 =  N'SELECT MBMDepartment, CLNType,BordxRefNo,
            						SUM(CLMTotalActual) as CLMTotalActual, SUM(CLMNonGPFee) as CLMNonGPFee
            						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby107 = 'GROUP BY MBMDepartment, CLNType,BordxRefNo,BordxDate'
			       set @l_sqlstr107 = @l_sqlstr107 + @strAppend + @l_sqlgroupby107 
			       exec  sp_executesql  @l_sqlstr107
		end

	if @strAppendCase = 108

		begin
         	      		set nocount on
			       declare @l_sqlstr108 as nvarchar(2000)
		                    declare @l_sqlgroupby108 as nvarchar(500)

		                    set  @l_sqlstr108 =  N'SELECT MBMDepartment, MBMNumber, CoverID, PatientName, CLNType, 
          						SUM(CLMTotalActual) as CLMTotalActual, SUM(CLMNonGPFee) as CLMNonGPFee 
          						From VIEW_CLMReport_Simple Where 0 = 0'
			       set @l_sqlgroupby108 = 'GROUP BY MBMDepartment, MBMNumber, CoverID, CLNType, PatientName'
			       set @l_sqlstr108 = @l_sqlstr108 + @strAppend + @l_sqlgroupby108
			       exec  sp_executesql  @l_sqlstr108
		end

return





