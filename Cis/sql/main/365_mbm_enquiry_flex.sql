-- Membership > MBM Enquiry FlexCare / Family Benefits: copies of CISDB spCalcFlexUOB and spGet_UOB_Flex_New_Cal (UOB FlexCare figures).
-- Differences: spGet_UOB_Flex_New_Cal read spCalcFlexUOB through OPENROWSET(SQLNCLI, Server=sqlinstance1) (an ad hoc distributed query to the
-- production server, disabled here); the copies share a temp table instead. spCalcFlexUOB's IC / employee filter is a parameter, not SQL text.

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER PROCEDURE cis.MbmEnq_CalcFlex -- '951009135440',null
	@icno nvarchar(20),
	@empno nvarchar(20)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	declare @icno_l nvarchar(20) = @icno
	declare @empno_l nvarchar(20) = @empno

	Declare @queryMBM AS NVARCHAR(MAX)
Declare @queryMBMcondition as nvarchar(max) =''

set @queryMBM=
	'
select 
row_number() over(order by mbmclnothers.MBMPolEffDate) as MBMCountNo,
mbm.MBMNumber,
cast(MBMCLNOthers.MBMPolEffDate as date) MBMPolEffDate,
cast(MBMCLNOthers.MBMPolExpDate as date) MBMPolExpDate,
cast(cast(ceiling(floor(cast(datediff(d,MBMCLNOthers.MBMPolEffDate,MBMCLNOthers.MBMPolExpDate) as decimal(18,2))/IIF(YEAR(MBMCLNOthers.MBMPolEffDate)%4=0,365,364) * tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as decimal(18,2)) as nvarchar) as AnnualLimit,
--sum(coalesce(claim.CLMBordxAmt,0)) FlexUtil,
coalesce(fConly.FlexUtilFinal,0) as FlexUtil,
coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0) * coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0) as LeavesSoldValue,
coalesce(tblUOBFlexCFOthersConfig.Reward,0) as Reward,

(cast(cast(ceiling(floor(cast(datediff(d,MBMCLNOthers.MBMPolEffDate,MBMCLNOthers.MBMPolExpDate) as decimal(18,2))/IIF(YEAR(MBMCLNOthers.MBMPolEffDate)%4=0,365,364) * tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as decimal(18,2)) as nvarchar)
+(coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0) * coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0))
+coalesce(tblUOBFlexCFOthersConfig.Reward,0)) as OverallLimit,

(cast(cast(ceiling(floor(cast(datediff(d,MBMCLNOthers.MBMPolEffDate,MBMCLNOthers.MBMPolExpDate) as decimal(18,2))/IIF(YEAR(MBMCLNOthers.MBMPolEffDate)%4=0,365,364) * tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as decimal(18,2)) as nvarchar)
+(coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0) * coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0))
+coalesce(tblUOBFlexCFOthersConfig.Reward,0))-coalesce(fConly.FlexUtilFinal,0) as OverallBal

From 
hisdb..mbm
inner join hisdb..mbmclnothers on mbmclnothers.mbmnumber=mbm.mbmnumber
inner join hisdb..mbmothers on mbmothers.mbmnumber=mbm.mbmnumber
inner join cisdb..tblFlexibleBenefitPlanLimit_OP on tblFlexibleBenefitPlanLimit_OP.CISPlnCode=MBMCLNOthers.PLNType and mbmclnothers.MBMPolEffDate between tblFlexibleBenefitPlanLimit_OP.FBEffDate and tblFlexibleBenefitPlanLimit_OP.FBExpDate
left join cisdb..ciscas on ciscas.MBMnumber=mbmclnothers.MBMNumber and ciscas.CoverID=''00'' and CASStatus =''O''
left join cisdb..tblFlexibleBenefitPlan_OP on tblFlexibleBenefitPlan_OP.FBCode=ciscas.CASTreatReason1 
left join CISWorksheet..claim on claim.CASNumber=ciscas.CASNumber
left join cisdb..tblUOBFlexCFOthersConfig on tblUOBFlexCFOthersConfig.MemberNo=mbm.MBMNumber and tblUOBFlexCFOthersConfig.CoverId=''00''
left join (select ciscas.MBMnumber,count(*) cnt,sum(coalesce(CLMBordxAmt,0)) as FlexUtilFinal from cisdb..ciscas inner join cisdb..tblFlexibleBenefitPlan_OP FBP_OP on FBP_OP.FBCode=ciscas.CASTreatReason1 inner join CISWorksheet..claim on claim.CASNumber=ciscas.CASNumber where ciscas.CoverID=''00'' and CASStatus =''O'' group by ciscas.MBMnumber) FConly on FConly.MBMnumber=mbm.MBMNumber
where  
(
	0=0 
) '

+ iif(@icno_l is not null,' and mbm.MBMIcBcPp=@icno','')
+ iif(@empno_l is not null,' and mbmothers.MBMEmployeeNo=@empno','')
+
' 
group by 
fconly.FlexUtilFinal,
mbm.MBMNumber,
MBMCLNOthers.MBMPolEffDate,
MBMCLNOthers.MBMPolExpDate,
tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit,
tblUOBFlexCFOthersConfig.LeavesCF,tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,tblUOBFlexCFOthersConfig.LeavesSold,
tblUOBFlexCFOthersConfig.Reward
order by MBMCLNOthers.MBMPolEffDate
'

--print @queryMBM
declare @F1 as table(
MBMCountNo int,
MBMNumber nvarchar(20),
MBMPolEffDate date,
MBMPolExpDate date,
AnnualLimit money,
FlexUtil money,
LeavesSoldValue money,
Reward money,
OverallLimit money,
OverallBal money
)

insert into @F1
exec sp_executesql @queryMBM, N'@icno nvarchar(20), @empno nvarchar(20)', @icno = @icno_l, @empno = @empno_l

;with cte as
(
select 
F1.*,
case when F2.OverallBal>0 then F2.OverallBal/2 else 0 end as CFfromLastYear
from @F1 F1
left join @F1 F2 on F2.MBMCountNo=F1.MBMCountNo-1
)
insert into #flex_calc
select 
cte.*,
((case when cteFinal.OverallBal >0 then coalesce(cteFinal.OverallBal,0) else 0 end) +coalesce(cteFinal.CFfromLastYear,0))/2 as FinalCF,
coalesce(cte.OverallBal,0)+((coalesce(cteFinal.OverallBal,0)+coalesce(cteFinal.CFfromLastYear,0))/2) as YearBalance
from cte
left join cte cteFinal on cteFinal.MBMCountNo=cte.MBMCountNo-1

	
END
GO

CREATE OR ALTER PROCEDURE cis.MbmEnq_Flex -- null,null,'UBHA0374880*05','2021'
	@icno nvarchar(20)=null,
	@empno nvarchar(20)=null,
	@mbmno nvarchar(20)=null,
	@polYear int=null
	--@res int output
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	if (@mbmno is not null) 
		Begin
			select @empno = PrincipleEmpNoActual from hisdb..vuMemberFamily_OP_short where vuMemberFamily_OP_short.MemberNo = @mbmno
			Set @mbmno='' 
		End	

	Declare @tmpFlexAccu table
	(
		YearId int,
		MBMCountNo int,
		MemberNo nvarchar(20),
		CoverId nvarchar(2),
		ICNo nvarchar(20),
		PrincipalICNo nvarchar(20),
		EffectiveDate date,
		ExpiryDate date,
		MbmStatus char(1),
		mbmemployeeno nvarchar(20),
		PLNType nvarchar(20),
		AccuAmt_CLMBordxAmt money,
		FCEffDate date,
		ALEffDate date
	);

	Declare @tmpFlexLimit table
	(
		YearId int,
		MBMCountNo int,
		MemberNo nvarchar(20),
		CoverId nvarchar(2),
		ICNo nvarchar(20),
		PrincipalICNo nvarchar(20),
		EffectiveDate date,
		ExpiryDate date,
		MbmStatus char(1),
		mbmemployeeno nvarchar(20),
		PLNType nvarchar(20),
		AccuAmt_CLMBordxAmt money,
		Limits nvarchar(20),
		ALEffDate date
	);


	Declare @queryMBM AS NVARCHAR(MAX)
	Declare @queryMBMcondition as nvarchar(max) =''

	set @queryMBM='
		with cte as 
	(
	
		select mbm.mbmnumber,''00'' as CoverId,mbmicbcpp as ICNo,mbmicbcpp PrincipalICNo,mbmclnothers.MBMPolEffDate EffectiveDate,mbmclnothers.MBMPolExpDate ExpiryDate,mbmstatus,
		mbmemployeeno,mbmclnothers.PLNType
		From 
		hisdb..mbm
		inner join hisdb..mbmclnothers on mbmclnothers.mbmnumber=mbm.mbmnumber
		inner join hisdb..mbmothers on mbmothers.mbmnumber=mbm.mbmnumber

		union 

		select 
		mbmcoveredpersons.mbmcnumber,mbmcoveredpersons.mbmccoverid as CoverId,mbmcoveredpersons.MBMCICBCPPNo as ICNo,mbmicbcpp PrincipalICNo ,MBMCCLNEffDate,MBMCCLNExpDate,mbmcstatus,
		mbmemployeeno,mbmclnothers.PLNType
		From hisdb..mbmcoveredpersons
		inner join hisdb..mbm on mbm.mbmnumber=mbmcoveredpersons.mbmcnumber
		inner join hisdb..mbmcoveredpersonscln on mbmcoveredpersonscln.mbmcnumber=mbmcoveredpersons.mbmcnumber and mbmcoveredpersonscln.mbmccoverid=mbmcoveredpersons.mbmccoverid
		inner join hisdb..mbmothers on mbmothers.mbmnumber=mbm.mbmnumber
		inner join hisdb..mbmclnothers on mbmclnothers.mbmnumber=mbm.mbmnumber
	)
	select dense_rank() over (order by cte.mbmnumber) as YearId, row_number() over(partition by cte.mbmnumber order by cte.mbmnumber,cte.CoverId)-1 as MBMCountNo, cte.*,
	--case 
	--			when ((row_number() over(partition by cte.mbmnumber order by cte.mbmnumber,cte.CoverId)-1)=0 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(ceiling(floor(cast((datediff(d,cte.EffectiveDate,cte.ExpiryDate)+1) as numeric)/365*tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as nvarchar) 
	--			when ((row_number() over(partition by cte.mbmnumber order by cte.mbmnumber,cte.CoverId)-1)=1 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(tblFlexibleBenefitPlanLimit_OP.FBFamilyLimit as nvarchar)
	--			when ((row_number() over(partition by cte.mbmnumber order by cte.mbmnumber,cte.CoverId)-1)>1 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then ''NA''
	--			end as Limits,
	sum(coalesce(claim.CLMBordxAmt,0)) as AccuAmt_CLMBordxAmt,tblFlexibleBenefitPlanLimit_OP.ClaimSubmissionLastDate as FCEffDate,tblFlexibleBenefitPlanLimit_OP.AnnualLeaveSellingLastDate as ALEffDate
		From cte
	inner join tblFlexibleBenefitPlanLimit_OP on tblFlexibleBenefitPlanLimit_OP.CISPlnCode=cte.PLNType and (cte.EffectiveDate between tblFlexibleBenefitPlanLimit_OP.FBEffDate and tblFlexibleBenefitPlanLimit_OP.FBExpDate)
	left join cisdb..ciscas on (ciscas.MBMnumber=cte.mbmnumber and ciscas.CoverID=cte.CoverID) and coalesce(ciscas.CASStatus,''O'')=''O''
	left join CISWorksheet..claim on claim.casnumber=ciscas.casnumber
	where 
	--tblFlexibleBenefitPlanLimit_OP.CISPlnCode in (''U206'',''U129'')
	--and mbmstatus=''a''
	--and 
	--year(ExpiryDate) in (year(getdate()),year(getdate())-1)
	year(ExpiryDate) in ('+cast(@polYear as nvarchar(4))+','+cast((@polYear-1) as nvarchar(4))+')'
	

	if @icno is not null
		set @queryMBMcondition=' and cte.PrincipalICNo= '''+@icno +''''
	else if @empno is not null
		set @queryMBMcondition=' and cte.mbmemployeeno= '''+@empno+''''

	set @queryMBM = @queryMBM+@queryMBMcondition+' group by cte.mbmnumber,cte.CoverID,ICNo,cte.PrincipalICNo,cte.mbmemployeeno,PLNType,EffectiveDate,ExpiryDate,cte.mbmstatus,CISPlnCode,FBEmployeeLimit,FBFamilyLimit,tblFlexibleBenefitPlanLimit_OP.ClaimSubmissionLastDate,tblFlexibleBenefitPlanLimit_OP.AnnualLeaveSellingLastDate
				order by year(cte.EffectiveDate),cte.mbmnumber,CoverID'

	print @queryMBM
	insert into @tmpFlexAccu
	exec sp_executesql @queryMBM

--	delete from @tmpFlexAccu
--	where MemberNo in
--	(
--'UBHA0030096*03',
--'UBHA0030096*04',
--'UBHA0091013*03',
--'UBHA0091013*04',
--'UBHA0212743*03',
--'UBHA0212743*04',
--'UBHA0213854*03',
--'UBHA0213854*04',
--'UBHA0299217*03',
--'UBHA0299217*04',
--'UBHA0301338*03',
--'UBHA0301338*04',
--'UBHA1112690*02',
--'UBHA1112690*03',
--'UBHA1132149*02',
--'UBHA1132149*03',
--'UBHA1148747*02',
--'UBHA1148747*03',
--'UBHA1168053*02',
--'UBHA1168053*03',
--'UBHA1178778*02',
--'UBHA1178778*03'
--)

	declare @FCEffDate as date = null
	if (select max(YearId) from @tmpFlexAccu where MBMCountNo=0 )<>1
	begin
		select @FCEffDate=FCEffDate from @tmpFlexAccu where MBMCountNo=0 and YearId=1
	end

insert into @tmpFlexLimit
select 
tFAcc.YearId,MBMCountNo,MemberNo,CoverId,ICNo,PrincipalICNo,EffectiveDate,ExpiryDate,MbmStatus,mbmemployeeno,PLNType,AccuAmt_CLMBordxAmt,
case
	--when (@polYear=2019 and MBMCountNo=0 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(ceiling(floor(cast((datediff(d,tFAcc.EffectiveDate,tFAcc.ExpiryDate)+1) as numeric)/365*tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as nvarchar) 
	--when (MBMCountNo=0 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(cast(ceiling(floor(cast((datediff(d,tFAcc.EffectiveDate,tFAcc.ExpiryDate)+ case when year(tFAcc.EffectiveDate)=2019 or (FORMAT (tFAcc.EffectiveDate, 'dd-MM')<>'01-01') then 1 else 0 end) as decimal(18,2))/365*tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as decimal(18,2)) as nvarchar) 
	when (MBMCountNo=0 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(cast(ceiling(floor(cast(datediff(d,tFAcc.EffectiveDate,tFAcc.ExpiryDate) as decimal(18,2))/IIF(YEAR(tFAcc.EffectiveDate)%4=0,365,364) * tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit)/10)*10 as decimal(18,2)) as nvarchar)
	when (MBMCountNo=0 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(tblFlexibleBenefitPlanLimit_OP.FBEmployeeLimit as nvarchar) 
	when (MBMCountNo=1 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then cast(tblFlexibleBenefitPlanLimit_OP.FBFamilyLimit as nvarchar)
	when (MBMCountNo>1 and tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null) then 'NA'
	end as Limits,tFAcc.ALEffDate
from @tmpFlexAccu tFAcc
inner join tblFlexibleBenefitPlanLimit_OP on tblFlexibleBenefitPlanLimit_OP.CISPlnCode=tFAcc.PLNType and (tFAcc.EffectiveDate between tblFlexibleBenefitPlanLimit_OP.FBEffDate and tblFlexibleBenefitPlanLimit_OP.FBExpDate)

	

--select * From @tmpFlexLimit

if (select count(*) from @tmpFlexLimit)>0
begin

select YearId,MBMCountNo,MemberNo,CoverId,ICNo,PLNType,EffectiveDate,ExpiryDate,Limits,AccuAmt_CLMBordxAmt,ALEffDate into #tmpFlexLimit from @tmpFlexLimit 

DECLARE 
	@cols AS NVARCHAR(MAX),
	@colsSum AS NVARCHAR(MAX),
	@query  AS NVARCHAR(MAX)

	select @cols = STUFF((SELECT ',' + QUOTENAME(FBCode) 
						from cisdb..tblFlexibleBenefitPlan_OP
						--where tblFlexibleBenefitPlan_OP.FBCode not in ('301','302')
						group by FBCode
						order by FBCode
				FOR XML PATH(''), TYPE
				).value('.', 'NVARCHAR(MAX)') 
			,1,1,'')


	set @colsSum=replace(@cols,',','+')
	set @colsSum=replace(@colsSum,'[','coalesce([')
	set @colsSum=replace(@colsSum,']','],0)')

	set @query = N'SELECT 
	dense_rank() over (order by MemberNo) as FamilyRowId, 
	mbmno,
	MBMCountNo,MemberNo,CoverID,ICNo, PLNType[PLNType],EffectiveDate,ExpiryDate,Limits,AccuAmt_CLMBordxAmt,
	coalesce([135],0) [135],[135CapLimit],case when [135CapLimit] is not null then [135CapLimit] - coalesce([135],0) end [135Bal],
	coalesce([266],0) [266],[266CapLimit],case when [266CapLimit] is not null then [266CapLimit] - coalesce([266],0) end [266Bal],
	coalesce([300],0)+coalesce([301],0)+coalesce([302],0)+coalesce([211],0) [300],[300CapLimit],case when [300CapLimit] is not null then [300CapLimit] - coalesce([300],0) end [300Bal],
	coalesce([303],0) [303],[303CapLimit],case when [303CapLimit] is not null then [303CapLimit] - coalesce([303],0) end [303Bal],
	coalesce([305],0) [305],[305CapLimit],case when [305CapLimit] is not null then [305CapLimit] - coalesce([305],0) end [305Bal],
	coalesce([306],0) [306],[306CapLimit],case when [306CapLimit] is not null then [306CapLimit] - coalesce([306],0) end [306Bal],
	(coalesce([466],0)+coalesce([662],0)) [466],[466CapLimit],case when [466CapLimit] is not null then [466CapLimit] - (coalesce([466],0) + coalesce([662],0)) end [466Bal],
	coalesce([468],0) [468],[468CapLimit],case when [468CapLimit] is not null then [468CapLimit] - coalesce([468],0) end [468Bal],
	coalesce([614],0) [614],[614CapLimit],case when [614CapLimit] is not null then [614CapLimit] - coalesce([614],0) end [614Bal],
	coalesce([615],0) [615],[615CapLimit],case when [615CapLimit] is not null then [615CapLimit] - coalesce([615],0) end [615Bal],
	coalesce([616],0) [616],[616CapLimit],case when [616CapLimit] is not null then [616CapLimit] - coalesce([616],0) end [616Bal],
	coalesce([617],0) [617],[617CapLimit],case when [617CapLimit] is not null then [617CapLimit] - coalesce([617],0) end [617Bal],
	coalesce([618],0) [618],[618CapLimit],case when [618CapLimit] is not null then [618CapLimit] - coalesce([618],0) end [618Bal],
	coalesce([619],0) [619],[619CapLimit],case when [619CapLimit] is not null then [619CapLimit] - coalesce([619],0) end [619Bal],
	coalesce([654],0) [654],[654CapLimit],case when [654CapLimit] is not null then [654CapLimit] - coalesce([654],0) end [654Bal],
	
	coalesce([696],0) [696],[696CapLimit],case when [696CapLimit] is not null then [696CapLimit] - coalesce([696],0) end [696Bal],
	coalesce([697],0) [697],[697CapLimit],case when [697CapLimit] is not null then [697CapLimit] - coalesce([697],0) end [697Bal],
	(coalesce([703],0)+coalesce([698],0)+coalesce([699],0)+coalesce([700],0)+coalesce([701],0)+coalesce([702],0)) [703],[703CapLimit],case when [703CapLimit] is not null then [703CapLimit] - (coalesce([703],0) + coalesce([703],0)+coalesce([698],0)+coalesce([699],0)+coalesce([700],0)+coalesce([701],0)+coalesce([702],0)) end [703Bal],
	'  
	 
	+ @colsSum 
	+' as SumFlexAmount, case when CoverID=''00'' then coalesce(AccuAmt_ClmBordxAmt,0)-(' +@colsSum+ ') else coalesce(AccuAmt_ClmBordxAmt,0) end as SumOP ' 
	+N'from 
				 (
					select 
					CASE WHEN CHARINDEX(''*'', #tmpFlexLimit.MemberNo) > 0 THEN LEFT(#tmpFlexLimit.MemberNo, CHARINDEX(''*'', #tmpFlexLimit.MemberNo) - 1) ELSE #tmpFlexLimit.MemberNo END AS MbmNo,
					MBMCountNo,#tmpFlexLimit.MemberNo,#tmpFlexLimit.CoverID,#tmpFlexLimit.ICNo,#tmpFlexLimit.PLNType,#tmpFlexLimit.EffectiveDate,#tmpFlexLimit.ExpiryDate,#tmpFlexLimit.Limits,
					#tmpFlexLimit.AccuAmt_CLMBordxAmt,
					--count(*) as CaseCount, 
					sum(coalesce(claim.CLMBordxAmt,0)) as AccuFlexSum,
					tblFlexibleBenefitPlan_OP.FBCode,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''135'',case when mbmcountno=0 then 1 else 0 end) end as [135CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''266'',case when mbmcountno=0 then 1 else 0 end) end as [266CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''300'',case when mbmcountno=0 then 1 else 0 end) end as [300CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''303'',case when mbmcountno=0 then 1 else 0 end) end as [303CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''305'',case when mbmcountno=0 then 1 else 0 end) end as [305CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''306'',case when mbmcountno=0 then 1 else 0 end) end as [306CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''466'',case when mbmcountno=0 then 1 else 0 end) end as [466CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''468'',case when mbmcountno=0 then 1 else 0 end) end as [468CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''614'',case when mbmcountno=0 then 1 else 0 end) end as [614CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''615'',case when mbmcountno=0 then 1 else 0 end) end as [615CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''616'',case when mbmcountno=0 then 1 else 0 end) end as [616CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''617'',case when mbmcountno=0 then 1 else 0 end) end as [617CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''618'',case when mbmcountno=0 then 1 else 0 end) end as [618CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''619'',case when mbmcountno=0 then 1 else 0 end) end as [619CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''654'',case when mbmcountno=0 then 1 else 0 end) end as [654CapLimit] ,

					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''696'',case when mbmcountno=0 then 1 else 0 end) end as [696CapLimit] , 
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''697'',case when mbmcountno=0 then 1 else 0 end) end as [697CapLimit] ,
					case when mbmcountno in (0,1) then dbo.fnGetFlexConfig(1,''703'',case when mbmcountno=0 then 1 else 0 end) end as [703CapLimit] 
					
					From #tmpFlexLimit
					left join cisdb..ciscas on ciscas.MBMnumber=#tmpFlexLimit.MemberNo and ciscas.CoverID=#tmpFlexLimit.CoverID and coalesce(ciscas.CASStatus,''O'')=''O''
					left join CISWorksheet..claim on claim.casnumber=ciscas.casnumber
					left join cisdb..TreatmentReason TR1 on tr1.ReasonCode=ciscas.CASTreatReason1
					left join cisdb..tblFlexibleBenefitPlan_OP on tblFlexibleBenefitPlan_OP.FBCode=tr1.ReasonCode and ((tblFlexibleBenefitPlan_OP.FBEmployee=1 and #tmpFlexLimit.CoverID=''00'') or (tblFlexibleBenefitPlan_OP.FBFamily=1 and #tmpFlexLimit.CoverID<>''00''))
					--where and coalesce(ciscas.CASStatus,''O'')=''O''
					group by 
					MBMCountNo,
					#tmpFlexLimit.MemberNo,
					#tmpFlexLimit.CoverID,
					#tmpFlexLimit.ICNo,
					#tmpFlexLimit.PLNType,
					#tmpFlexLimit.EffectiveDate,
					#tmpFlexLimit.ExpiryDate,
					#tmpFlexLimit.Limits,
					#tmpFlexLimit.AccuAmt_CLMBordxAmt,
					tblFlexibleBenefitPlan_OP.FBCode
				) x
				pivot 
				(
					sum(AccuFlexSum)
					for FBCode in (' + @cols + N')
				) p '


		--PRINT CAST(@query AS NTEXT)
		create table #tmpUtil (
		FamilyRowId int,
		mbmno nvarchar(20),
		MBMCountNo tinyint,
		MemberNo nvarchar(20),
		CoverID nvarchar(2),
		ICNO nvarchar(100),
		PLNType nvarchar(4),
		EffectiveDate Date,
		ExpiryDate Date,
		Limits nvarchar(20),
		AccuAmt_CLMBordxAmt money,
		[135] money,
		[135Cap] money,
		[135Bal] money,
		[266] money,
		[266Cap] money,
		[266Bal] money,
		[300] money,
		[300Cap] money,
		[300Bal] money,
		[303] money,
		[303Cap] money,
		[303Bal] money,
		[305] money,
		[305Cap] money,
		[305Bal] money,
		[306] money,
		[306Cap] money,
		[306Bal] money,
		[466] money,
		[466Cap] money,
		[466Bal] money,
		[468] money,
		[468Cap] money,
		[468Bal] money,
		[614] money,
		[614Cap] money,
		[614Bal] money,
		[615] money,
		[615Cap] money,
		[615Bal] money,
		[616] money,
		[616Cap] money,
		[616Bal] money,
		[617] money,
		[617Cap] money,
		[617Bal] money,
		[618] money,
		[618Cap] money,
		[618Bal] money,
		[619] money,
		[619Cap] money,
		[619Bal] money,
		[654] money,
		[654Cap] money,
		[654Bal] money,

		[696] money,
		[696Cap] money,
		[696Bal] money,
		[697] money,
		[697Cap] money,
		[697Bal] money,
		[703] money,
		[703Cap] money,
		[703Bal] money,
		SumFlexAmount money,
		SumOP money
		);

	insert into #tmpUtil
	exec sp_executesql @query;

	ALTER TABLE #tmpUtil ADD DepUtil money,	DepBal money, PriFlexBal money

	declare @DepUtil money, @DepLimits money
	declare @FamilyCount int = 0
	declare @FamilyMemberCount int = 0
	declare @CurrentFamilyRowId int = 1
	declare @CurrentFamilyMemberRowId int = 1

	select @FamilyCount=max(FamilyRowId) from #tmpUtil

	while @CurrentFamilyRowId<=@FamilyCount
	begin
		select @DepUtil=sum(SumOP) From #tmpUtil where mbmcountno>0 and FamilyRowId=@CurrentFamilyRowId
		select @DepLimits=Limits from #tmpUtil where MBMCountNo=1 and FamilyRowId=@CurrentFamilyRowId
	
		update #tmpUtil set DepUtil=@DepUtil,DepBal=@DepLimits-@DepUtil where MBMCountNo=1 and FamilyRowId=@CurrentFamilyRowId
		update #tmpUtil set PriFlexBal=Limits-SumFlexAmount where MBMCountNo=0 and FamilyRowId=@CurrentFamilyRowId

		set @CurrentFamilyRowId=@CurrentFamilyRowId+1
	end

	
	-- The desktop procedure loads spCalcFlexUOB through OPENROWSET(SQLNCLI, sqlinstance1); here a temp table shared with the callee.
	create table #flex_calc (MBMCountNo int, MBMNumber nvarchar(20), MBMPolEffDate date, MBMPolExpDate date, AnnualLimit money, FlexUtil money,
		LeavesSoldValue money, Reward money, OverallLimit money, OverallBal money, CFfromLastYear money, FinalCF money, YearBalance money)
	if @icno is not null exec cis.MbmEnq_CalcFlex @icno, null
	else if @empno is not null exec cis.MbmEnq_CalcFlex null, @empno

	declare @tblFlexCalc as table(
		MBMCountNo int,
		MBMNumber nvarchar(20),
		MBMPolEffDate date,
		MBMPolExpDate date,
		AnnualLimit money,
		FlexUtil money,
		LeavesSoldValue money,
		Reward money,
		OverallLimit money,
		OverallBal money,
		CFfromLastYear money,
		FinalCF money,
		YearBalance money

	)

	insert into @tblFlexCalc select * from #flex_calc


	select 
	dense_rank() over (order by #tmpUtil.mbmno) as MBMId,
	#tmpUtil.* ,
	--case when #tmpUtil.MBMCountNo='0' then iif(coalesce(#tmpUtil.Priflexbal,0)>0,coalesce(#tmpUtil.Priflexbal,0)/2,0) end as PriFlexBalCF,
	--case when #tmpUtil.MBMCountNo='0' then iif(coalesce(#tmpUtil.Priflexbal,0)>0,coalesce(#tmpUtil.Priflexbal,0)/2,0) + iif(((coalesce(tblUOBFlexCFOthersConfig.leavessold,0)*coalesce(leaveexchangevalueperday,0))+coalesce(Reward,0))>0, ((coalesce(tblUOBFlexCFOthersConfig.leavessold,0)*coalesce(leaveexchangevalueperday,0))+coalesce(Reward,0))/2,0) end as PriFlexBalCF,
	FlexCareCFcalc.FinalCF as PriFlexBalCF,
	case when #tmpUtil.CoverID<>'00' then (coalesce(#tmpUtil.SumOP,0)-coalesce(SumFlexAmount,0)) else null end as IPOPSPUtil,
	case when #tmpUtil.MBMCountNo='0' then coalesce(#tmpUtil.Priflexbal,0) when #tmpUtil.MBMCountNo='1' then coalesce(#tmpUtil.DepBal,0) else null end as FinalBalAll,
	vuMemberFamily_OP_short.MemberName,
	vuMemberFamily_OP_short.PrincipleEmpNoActual,
	ALEffDate
	into #tmpFlexCF
	From #tmpUtil 
	inner join #tmpFlexLimit on #tmpFlexLimit.MemberNo=#tmpUtil.MemberNo and #tmpFlexLimit.CoverID=#tmpUtil.CoverID 
	inner join hisdb..vuMemberFamily_OP_short on vuMemberFamily_OP_short.MemberNo=#tmpUtil.MemberNo and vuMemberFamily_OP_short.CoverID=#tmpUtil.CoverID 
	left join tblUOBFlexCFOthersConfig on tblUOBFlexCFOthersConfig.MemberNo=#tmpUtil.MemberNo and tblUOBFlexCFOthersConfig.CoverId=#tmpUtil.CoverId
	left join @tblFlexCalc FlexCareCFcalc on FlexCareCFcalc.MBMNumber=vuMemberFamily_OP_short.MemberNo and vuMemberFamily_OP_short.CoverID='00'
	order by FamilyRowId,#tmpUtil.MBMCountNo
	
	--select * from #tmpFlexCF

	declare @cnt int
	select @cnt=max(mbmid) from #tmpFlexCF
	
	--declare @currMBMId int =1
	--while @currMBMId<=@cnt
	--begin
	--		if (select count(mbmcountno) from #tmpFlexCF where mbmid=@currMBMId and mbmcountno=0)=2
	--		begin
	--		declare @CFAmt as nvarchar(20) =(select coalesce(PriFlexBalCF,0) from #tmpFlexCF where mbmid=@currMBMId and mbmcountno=0 and year(effectivedate)= (@polYear-1)) --year(getdate())-1)
	--			 update #tmpFlexCF set PriFlexBalCF=@CFAmt where mbmid=@currMBMId and mbmcountno=0 and year(effectivedate)=@polYear --year(getdate())
	--			 update #tmpFlexCF set PriFlexBalCF=null where mbmid=@currMBMId and mbmcountno=0 and year(effectivedate)=(@polYear-1) --year(getdate())-1
	--		end
	--		else
	--			begin
	--			 update #tmpFlexCF set PriFlexBalCF=0.00 where mbmid=@currMBMId and mbmcountno=0 
	--			end
	--	set @currMBMId=@currMBMId+1
	--end

	select MBMId,	FamilyRowId,	mbmno,	MBMCountNo,	#tmpFlexCF.MemberNo,	#tmpFlexCF.CoverID,	ICNO,	PLNType,	EffectiveDate,	ExpiryDate,	
	
	Limits,
	--iif(coalesce(Limits,0)>0,cast(Limits as decimal(18,2)),Limits) as Limits,
	
	AccuAmt_CLMBordxAmt,	
	#tmpFlexCF.[135],	#tmpFlexCF.[135Cap],	#tmpFlexCF.[135Bal],	
	#tmpFlexCF.[266],	#tmpFlexCF.[266Cap],	#tmpFlexCF.[266Bal],	
	#tmpFlexCF.[300],	#tmpFlexCF.[300Cap],	#tmpFlexCF.[300Bal],	
	#tmpFlexCF.[303],	#tmpFlexCF.[303Cap],	#tmpFlexCF.[303Bal],	
	#tmpFlexCF.[305],	#tmpFlexCF.[305Cap],	#tmpFlexCF.[305Bal],	
	#tmpFlexCF.[306],	#tmpFlexCF.[306Cap],	#tmpFlexCF.[306Bal],	
	#tmpFlexCF.[466],	#tmpFlexCF.[466Cap],	#tmpFlexCF.[466Bal],	
	#tmpFlexCF.[468],	#tmpFlexCF.[468Cap],	#tmpFlexCF.[468Bal],	
	#tmpFlexCF.[614],	#tmpFlexCF.[614Cap],	#tmpFlexCF.[614Bal],	
	#tmpFlexCF.[615],	#tmpFlexCF.[615Cap],	#tmpFlexCF.[615Bal],	
	#tmpFlexCF.[616],	#tmpFlexCF.[616Cap],	#tmpFlexCF.[616Bal],	
	#tmpFlexCF.[617],	#tmpFlexCF.[617Cap],	#tmpFlexCF.[617Bal],	
	#tmpFlexCF.[618],	#tmpFlexCF.[618Cap],	#tmpFlexCF.[618Bal],	
	#tmpFlexCF.[619],	#tmpFlexCF.[619Cap],	#tmpFlexCF.[619Bal],	
	#tmpFlexCF.[654],	#tmpFlexCF.[654Cap],	#tmpFlexCF.[654Bal],

	#tmpFlexCF.[696],	#tmpFlexCF.[696Cap],	#tmpFlexCF.[696Bal],
	#tmpFlexCF.[697],	#tmpFlexCF.[697Cap],	#tmpFlexCF.[697Bal],
	#tmpFlexCF.[703],	#tmpFlexCF.[703Cap],	#tmpFlexCF.[703Bal],
	
	SumFlexAmount,	SumOP,	#tmpFlexCF.DepUtil,	#tmpFlexCF.DepBal,	#tmpFlexCF.PriFlexBal,	PriFlexBalCF,	IPOPSPUtil,	FinalBalAll,	MemberName,	PrincipleEmpNoActual,	
	LeavesCF,	cast(LeavesSold as int) LeavesSold,	tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,Reward,
	case when #tmpFlexCF.coverid='00' then convert(decimal(18,2),coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0)*coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0)) else null end as LeavesSoldAmount,
	
	case when #tmpFlexCF.coverid='00' then 
	convert
	(
		decimal(18,2),coalesce(Limits,'0')+
		(
			iif(@FCEffDate<cast(getdate() as date),coalesce(PriFlexBalCF,0),0)
			+ iif(cast(getdate() as date)> #tmpFlexCF.ALEffDate, (coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0)*coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0)),0)
			+ iif(cast(getdate() as date)> #tmpFlexCF.ALEffDate, coalesce(tblUOBFlexCFOthersConfig.Reward,0),0)
		 )
	 )-convert(decimal(18,2),coalesce(SumFlexAmount,0))
	else null
	end as PriFlexCurrAvailBal,

	case when #tmpFlexCF.coverid='00' then 
	convert
	(
		decimal(18,2),coalesce(Limits,'0')+
		(
			iif(@FCEffDate<cast(getdate() as date),coalesce(PriFlexBalCF,0),0)
			+ iif(cast(getdate() as date)> #tmpFlexCF.ALEffDate, (coalesce(tblUOBFlexCFOthersConfig.LeavesSold,0)*coalesce(tblUOBFlexCFOthersConfig.LeaveExchangeValuePerDay,0)),0)
			+ iif(cast(getdate() as date)> #tmpFlexCF.ALEffDate, coalesce(tblUOBFlexCFOthersConfig.Reward,0),0)
		 )
	 )
	else null
	end as PriFlexAvailLimit,

	coalesce(@FCEffDate,'') as FCEffDate,
	coalesce(ALEffDate,'') as ALEffDate
	From #tmpFlexCF 
	left join tblUOBFlexCFOthersConfig on tblUOBFlexCFOthersConfig.MemberNo=#tmpFlexCF.MemberNo and tblUOBFlexCFOthersConfig.CoverId=#tmpFlexCF.CoverId
	where year(effectivedate)=@polYear --year(getdate())
	--order by MBMId,FamilyRowId,MBMCountNo
	order by MBMId,MBMCountNo,ExpiryDate Desc



	drop table #tmpFlexLimit
	drop table #tmpUtil
	drop table #tmpFlexCF
	--set @res=1
end

else
	begin
	print 'no data'
	--set @res=0
	end
END
GO
