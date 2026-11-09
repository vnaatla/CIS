-- Claim Bordx > Insert Worksheet (frmWSBordx): the worksheet lists, copied from CISWorksheet ClaimBordxModule_isPureASO (case 21) and
-- ClaimBordxModule (cases 23 UOB flexCare / 24 UOB outpatient) unchanged apart from the filter; the criteria the desktop pasted
-- into @strAppend are parameters here. Bordx numbering and the save are in 810_bordx_insert_save.sql.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.BordxWs_Search @bordxType tinyint, @isPureAso tinyint = 1,
    @groupCode nvarchar(10) = NULL, @scan bit = 0, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL,
    @visitFrom datetime = NULL, @visitTo datetime = NULL, @noVisitDate bit = 0, @department nvarchar(200) = NULL, @clinicType int = 0,
    @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @subsidiary nvarchar(200) = NULL, @payTo char(1) = NULL, @treatment nvarchar(10) = NULL,
    @omni bit = 0, @regFrom datetime = NULL, @regTo datetime = NULL, @clinic nvarchar(20) = NULL, @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = N''
        + CASE WHEN @groupCode IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.CLMGrpComCode = @groupCode' ELSE N'' END
        + CASE WHEN @scan = 1 AND @caseFrom IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber = @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseFrom IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseTo IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber <= @caseTo' ELSE N'' END
        + CASE WHEN @visitFrom IS NOT NULL THEN N' And CISCas.DateVisit >= @visitFrom' ELSE N'' END
        + CASE WHEN @visitTo IS NOT NULL THEN N' And CISCas.DateVisit <= @visitTo' ELSE N'' END
        + CASE WHEN @noVisitDate = 1 THEN N' AND (CISCas.DateVisit Is null OR CISCas.DateVisit = '''')' ELSE N'' END
        + CASE WHEN @department IS NOT NULL THEN N' AND claim.CLMDeptName = @department' ELSE N'' END
        + CASE @clinicType WHEN 1 THEN N' AND HISDB.dbo.CLN.CLNType = ''S''' WHEN 2 THEN N' AND HISDB.dbo.CLN.CLNType = ''G'''
            WHEN 3 THEN N' AND (HISDB.dbo.CLN.CLNType <> ''G'') AND (HISDB.dbo.CLN.CLNType <> ''S'')' ELSE N'' END
        + CASE WHEN @opFrom IS NOT NULL THEN N' AND CISCas.OPBatchNo >= @opFrom' ELSE N'' END
        + CASE WHEN @opTo IS NOT NULL THEN N' AND CISCas.OPBatchNo <= @opTo' ELSE N'' END
        + CASE WHEN @subsidiary IS NOT NULL THEN N' AND claim.clmsubsidiary = @subsidiary' ELSE N'' END
        + CASE @payTo WHEN 'C' THEN N' AND CISWorksheet.dbo.ClaimInvoice.INVPayto = ''C''' WHEN 'M' THEN N' AND CISWorksheet.dbo.ClaimInvoice.INVPayto = ''M''' ELSE N'' END
        + CASE WHEN @treatment IS NOT NULL THEN N' AND CASTreatReason1 = @treatment' ELSE N'' END
        + CASE WHEN @omni = 1 THEN N' AND GRPCompany.GRPOMNICLM = ''Y''' ELSE N'' END
        + CASE WHEN @regFrom IS NOT NULL THEN N' And CISCas.RegisterDate >= @regFrom' ELSE N'' END
        + CASE WHEN @regTo IS NOT NULL THEN N' And CISCas.RegisterDate <= @regTo' ELSE N'' END
        + CASE WHEN @clinic IS NOT NULL THEN N' AND CISCas.ClinicCode = @clinic' ELSE N'' END
        + CASE @entity WHEN 0 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) <>''SJ'' ' WHEN 1 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) =''SJ'' ' ELSE N'' END;

			       declare @sql nvarchar(max)

		                    --set @l_sqlstr2 =  N'Select  top (100) percent * From CISDB..VIEW_CISWS_Pending where 0 = 0'
--						set @sql =  N'SELECT TOP (100) PERCENT CISDB.dbo.CISCas.CASNumber, CISDB.dbo.CISCas.MBMnumber, CISDB.dbo.CISCas.CoverID, CISCas.CASStatus, CISCas.DateVisit, 
--                         CISCas.RegisterDate, HISDB.dbo.MBM.MBMName, HISDB.dbo.MBMCLNOthers.MBMPolNumber AS MBMPolicyNo, 
--                         HISDB.dbo.MBMCoveredPersons.MBMCName, HISDB.dbo.CLN.CLNName, HISDB.dbo.CLN.CLNCode, HISDB.dbo.MBMOthers.MBMGroupCompany, 
--                         HISDB.dbo.MBMCoveredPersons.MBMCRELCode, HISDB.dbo.MBMOthers.RELCode, HISDB.dbo.MBM.HLTCode, HISDB.dbo.MBM.MBMMemberType, 
--                         CISCas.CASTreatReason1, CISCas.CASTreatReason2, CISCas.CASTreatReason3, CISCas.CASTreatReason4, CISCas.CASTreatReason5, 
--                         TreatmentReason1.ReasonDesc AS ReasonDesc1, TreatmentReason3.ReasonDesc AS ReasonDesc3, TreatmentReason2.ReasonDesc AS ReasonDesc2, 
--                         TreatmentReason.ReasonDesc AS ReasonDesc4, TreatmentReason_1.ReasonDesc AS ReasonDesc5, CISCas.ClinicCode, CISCas.Remarks, 
--                         HISDB.dbo.MBM.AGECode, HISDB.dbo.MBM.HSCLNInd, HISDB.dbo.MBMCLNOthers.PLNType AS PLNCode, HISDB.dbo.MBMCLNOthers.INSType AS INSCode, 
--                         CISCas.CASDiagnosis5, HISDB.dbo.MBMCLNOthers.MBMPolNumber, HISDB.dbo.MBMCLNOthers.PAYCode, HISDB.dbo.MBMOthers.MBMDepartment, 
--                         HISDB.dbo.MBMOthers.MBMEmployeeNo, CISCas.InvoiceDate, HISDB.dbo.MBMOthers.MBMCostCenter, GRPCompany.GRPBordxFormat, 
--                         GRPCompany.RptFormat, HISDB.dbo.CLN.CLNType, GRPCompany.TyphoidRpt, HISDB.dbo.MBMOthers.MBMBranch, CISCas.CaseType, 
--                         CISCas.MCDays, HISDB.dbo.MBMOthers.MBMPosition, CISCas.OPBatchNo, CISCas.InvoiceNo, HISDB.dbo.MBM.MBMPolicyNo AS MBMHISPolNo, 
--                         HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMGroupCategory, HISDB.dbo.CLN.CLNFFSStatus, GRPCompany.GRPMCOStatus, 
--                         CISCas.CASClaimsApp, HISDB.dbo.MBMCoveredPersons.MBMAXANumber, HISDB.dbo.MBMOthers.MBMPrevMemNo, CISCas.MedInjType1, 
--                         CISCas.MedInjType2, CISCas.MedInjType3, CISCas.MedInjType4, CISCas.MedInjType5, CISCas.MedInjType6, CISCas.MedInjType7, 
--                         CISWorksheet.dbo.ClaimInvoiceActualOthers.GPMed, CISWorksheet.dbo.ClaimInvoiceActualOthers.SPMed, HISDB.dbo.CLN.CLNLocation, HISDB.dbo.CLN.CLNPState,
--                          HISDB.dbo.MBM.PLNCode AS HISPLN, HISDB.dbo.CLN.CLNState, CISCas.NPClinic, CISWorksheet.dbo.Claim.CASNumber AS Expr1, 
--                         CISWorksheet.dbo.Claim.ClaimStatus, CISWorksheet.dbo.Claim.ClaimBordxStatus, CISWorksheet.dbo.Claim.CLMPayableByMedix2C, 
--                         CISWorksheet.dbo.Claim.CLMTotalActual, CISWorksheet.dbo.Claim.CLMBordxAmt, CISWorksheet.dbo.Claim.CLMNonGPFee, 
--                         CISWorksheet.dbo.Claim.CLMMCOProcFee, CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISWorksheet.dbo.ClaimInvoice.INVBillAmt, 
--                         CISWorksheet.dbo.Claim.CLMFinalApprove AS CLMTotalApproved, GRPCompany.GRPOMNICLM, CISWorksheet.dbo.Claim.CLMBordxMemo, 
--                         GRPCompany.ASO_INS
--FROM           CISDB.dbo.TreatmentReason AS TreatmentReason3 RIGHT OUTER JOIN
--                         CISWorksheet.dbo.Claim WITH (nolock) RIGHT OUTER JOIN
--                         CISWorksheet.dbo.ClaimInvoiceActualOthers WITH (nolock) RIGHT OUTER JOIN
--                         CISDB.dbo.CISCas WITH (nolock) ON CISWorksheet.dbo.ClaimInvoiceActualOthers.CASNumber = CISDB.dbo.CISCas.CASNumber ON 
--                         CISWorksheet.dbo.Claim.CASNumber = CISDB.dbo.CISCas.CASNumber LEFT OUTER JOIN
--                         CISWorksheet.dbo.ClaimInvoice WITH (nolock) ON CISWorksheet.dbo.Claim.CASNumber = CISWorksheet.dbo.ClaimInvoice.CASNumber LEFT OUTER JOIN
--                         HISDB.dbo.CLN ON CISDB.dbo.CISCas.ClinicCode = HISDB.dbo.CLN.CLNCode ON TreatmentReason3.ReasonCode = CISDB.dbo.CISCas.CASTreatReason3 LEFT OUTER JOIN
--                         CISDB.dbo.TreatmentReason AS TreatmentReason_1 WITH (nolock) ON CISDB.dbo.CISCas.CASTreatReason5 = TreatmentReason_1.ReasonCode LEFT OUTER JOIN
--                         CISDB.dbo.TreatmentReason WITH (nolock) ON CISDB.dbo.CISCas.CASTreatReason4 = CISDB.dbo.TreatmentReason.ReasonCode LEFT OUTER JOIN
--                         CISDB.dbo.TreatmentReason AS TreatmentReason2 WITH (nolock) ON CISDB.dbo.CISCas.CASTreatReason2 = TreatmentReason2.ReasonCode LEFT OUTER JOIN
--                         CISDB.dbo.TreatmentReason AS TreatmentReason1 WITH (nolock) ON CISDB.dbo.CISCas.CASTreatReason1 = TreatmentReason1.ReasonCode LEFT OUTER JOIN
--                         HISDB.dbo.MBMCoveredPersons WITH (nolock) ON CISDB.dbo.CISCas.MBMnumber = HISDB.dbo.MBMCoveredPersons.MBMCNumber AND 
--                         CISDB.dbo.CISCas.CoverID = HISDB.dbo.MBMCoveredPersons.MBMCCoverID LEFT OUTER JOIN
--                         CISDB.dbo.GRPCompany RIGHT OUTER JOIN
--                         HISDB.dbo.MBMOthers WITH (nolock) INNER JOIN
--                         HISDB.dbo.MBM WITH (nolock) INNER JOIN
--                         HISDB.dbo.MBMCLNOthers WITH (nolock) ON HISDB.dbo.MBM.MBMNumber = HISDB.dbo.MBMCLNOthers.MBMNumber ON 
--                         HISDB.dbo.MBMOthers.MBMNumber = HISDB.dbo.MBM.MBMNumber ON CISDB.dbo.GRPCompany.GRPCoName = HISDB.dbo.MBMOthers.MBMGroupCompany ON 
--                         CISDB.dbo.CISCas.MBMnumber = HISDB.dbo.MBM.MBMNumber
--WHERE (CISWorksheet.dbo.Claim.ClaimBordxStatus = 0) 
--and CISCas.CASStatus <>''D''
--AND (CISCas.OPBatchNo IS NOT NULL) AND (CISCas.OPBatchNo <> N'''')
--and  ( MBM.HSCLNInd <> ''H'') 
--AND  (CISWorksheet.dbo.Claim.CLMBordxMemo = ''Y'' OR  CISWorksheet.dbo.Claim.CLMBordxMemo = '''' OR CISWorksheet.dbo.Claim.CLMBordxMemo IS NULL)  
--AND  0=0 '



			set @sql =  N'SELECT TOP (100) PERCENT CISDB.dbo.CISCas.CASNumber, CISDB.dbo.CISCas.MBMnumber, CISDB.dbo.CISCas.CoverID, CISCas.CASStatus, CISCas.DateVisit, 
                         CISCas.RegisterDate, HISDB.dbo.MBM.MBMName, HISDB.dbo.MBMCLNOthers.MBMPolNumber AS MBMPolicyNo, 
                      HISDB.dbo.MBMCoveredPersons.MBMCName, HISDB.dbo.CLN.CLNName, HISDB.dbo.CLN.CLNCode, 
					  coalesce(CISDB.dbo.GRPCompany.GRPCoName,'''') as MBMGroupCompany, 
                 HISDB.dbo.MBMCoveredPersons.MBMCRELCode, HISDB.dbo.MBMOthers.RELCode, HISDB.dbo.MBM.HLTCode, HISDB.dbo.MBM.MBMMemberType, 
						CISCas.CASTreatReason1, CISCas.CASTreatReason2, CISCas.CASTreatReason3, CISCas.CASTreatReason4, CISCas.CASTreatReason5,                        
						cisdb.dbo.fnGetTreatmentDesc(CASTreatReason1) as ReasonDesc1,
						cisdb.dbo.fnGetTreatmentDesc(CASTreatReason2) as ReasonDesc2,
						cisdb.dbo.fnGetTreatmentDesc(CASTreatReason3) as ReasonDesc3,
						cisdb.dbo.fnGetTreatmentDesc(CASTreatReason4) as ReasonDesc4,
						cisdb.dbo.fnGetTreatmentDesc(CASTreatReason5) as ReasonDesc5,
						 CISCas.ClinicCode, CISCas.Remarks, 
                         HISDB.dbo.MBM.AGECode, HISDB.dbo.MBM.HSCLNInd, HISDB.dbo.MBMCLNOthers.PLNType AS PLNCode, HISDB.dbo.MBMCLNOthers.INSType AS INSCode, 
                         CISCas.CASDiagnosis5, HISDB.dbo.MBMCLNOthers.MBMPolNumber, HISDB.dbo.MBMCLNOthers.PAYCode, HISDB.dbo.MBMOthers.MBMDepartment, 
                         HISDB.dbo.MBMOthers.MBMEmployeeNo, CISCas.InvoiceDate, HISDB.dbo.MBMOthers.MBMCostCenter, GRPCompany.GRPBordxFormat, 
                         GRPCompany.RptFormat, HISDB.dbo.CLN.CLNType, GRPCompany.TyphoidRpt, HISDB.dbo.MBMOthers.MBMBranch, CISCas.CaseType, 
                         CISCas.MCDays, HISDB.dbo.MBMOthers.MBMPosition, CISCas.OPBatchNo, CISCas.InvoiceNo, HISDB.dbo.MBM.MBMPolicyNo AS MBMHISPolNo, 
                         HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMGroupCategory, HISDB.dbo.CLN.CLNFFSStatus, GRPCompany.GRPMCOStatus, 
                         CISCas.CASClaimsApp, HISDB.dbo.MBMCoveredPersons.MBMAXANumber, HISDB.dbo.MBMOthers.MBMPrevMemNo, CISCas.MedInjType1, 
                         CISCas.MedInjType2, CISCas.MedInjType3, CISCas.MedInjType4, CISCas.MedInjType5, CISCas.MedInjType6, CISCas.MedInjType7, 
                         CISWorksheet.dbo.ClaimInvoiceActualOthers.GPMed, CISWorksheet.dbo.ClaimInvoiceActualOthers.SPMed, HISDB.dbo.CLN.CLNLocation, HISDB.dbo.CLN.CLNPState,
                          HISDB.dbo.MBM.PLNCode AS HISPLN, HISDB.dbo.CLN.CLNState, CISCas.NPClinic, CISWorksheet.dbo.Claim.CASNumber AS Expr1, 
                         CISWorksheet.dbo.Claim.ClaimStatus, CISWorksheet.dbo.Claim.ClaimBordxStatus, CISWorksheet.dbo.Claim.CLMPayableByMedix2C, 
                         CISWorksheet.dbo.Claim.CLMTotalActual, CISWorksheet.dbo.Claim.CLMSubsidiary, CISWorksheet.dbo.Claim.CLMDeptName, '
						 
						 IF @bordxType=0
							set @sql =  @sql + ' CASE WHEN Pln.PLNBTBG = 1 THEN  CISWorksheet.dbo.Claim.CLMBordxAmt - coalesce(Claim.CLMBTBGAmt,0) else CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '							
						 ELSE
							set @sql =  @sql + ' CASE WHEN (Pln.PLNBTBG = 1 and coalesce(CISWorksheet.dbo.Claim.CLMBTBGAmt,0)<>0) THEN CISWorksheet.dbo.Claim.CLMBTBGAmt ELSE CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '
							
						 						 
			set @sql =  @sql + 'CISWorksheet.dbo.Claim.CLMNonGPFee, 
                         CISWorksheet.dbo.Claim.CLMMCOProcFee, CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISWorksheet.dbo.ClaimInvoice.INVBillAmt, 
                         CISWorksheet.dbo.Claim.CLMFinalApprove AS CLMTotalApproved, GRPCompany.GRPOMNICLM, CISWorksheet.dbo.Claim.CLMBordxMemo, 
                         GRPCompany.ASO_INS,lg.DoctorName,CISWorksheet.dbo.claim.CLMCoPayColl,ciscas.Claimability,claim.clmcostcenter
							From cisdb..ciscas
							inner join CISWorksheet.dbo.Claim on claim.CASNumber=ciscas.CASNumber
							inner join CISWorksheet..ClaimInvoice on ClaimInvoice.CASNumber=claim.CASNumber 
							inner join CISWorksheet.dbo.ClaimInvoiceActualOthers on ClaimInvoiceActualOthers.CASNumber = ClaimInvoice.CASNumber 
							inner join hisdb..mbm on mbm.MBMNumber=ciscas.MBMnumber
							left join HISDB.dbo.MBMCoveredPersons on MBMCoveredPersons.MBMCNumber=CISCas.MBMNumber and MBMCCoverID=ciscas.CoverID
							inner join HISDB.dbo.MBMCLNOthers on MBMCLNOthers.MBMNumber=mbm.MBMNumber 
							inner join hisdb..MBMOthers on mbmothers.MBMnumber=mbm.MBMNumber
							left join CISDB.dbo.GRPCompany on GRPCoCode=claim.clmgrpcomcode and (grpcompany.hispaycode is null or ((grpcompany.hispaycode=mbm.paycode or grpcompany.hispaycodePrevYear = mbm.paycode) or (grpcompany.GRPCoPaycode=''DC'' and mbm.MBMPayorEffDate is null)))							
							inner join hisdb..cln on cln.CLNCode=ciscas.ClinicCode
							LEFT JOIN CISDB..LG ON LG.LGNo=CISCAS.CASNumber 
							left join cisdb..pln on pln.PLNCode= MBMCLNOthers.PLNType
							WHERE '
							IF @bordxType=0
								set @sql =  @sql + '(((pln.PLNBTBG =1 and coalesce(CISWorksheet.dbo.Claim.BTBGBordxStatus,0) = 0 and coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0 and coalesce(Claim.CLMBordxAmt,0) <> 0 and coalesce(Claim.CLMTotalActual,0) <> coalesce(Claim.CLMBTBGAmt,0)) OR  ( (coalesce(Claim.CLMBordxAmt,0) <> coalesce(Claim.CLMBTBGAmt,0)) and  coalesce(pln.PureASO,0) =0 and coalesce(pln.PLNBTBG,0) =0 and coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0 and coalesce(Claim.CLMBordxAmt,0) <> 0))) '								
							ELSE if @bordxType=1														
								--set @sql =  @sql + '((pln.PLNBTBG =1 and coalesce(Claim.BTBGBordxStatus,0)=0 and coalesce(Claim.CLMBTBGAmt,0) <> 0) OR (coalesce(pln.PLNBTBG,0) =0 and coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0)) '

								if @isPureAso=1
								begin
									set @sql =  @sql + '
										(
											(coalesce(pln.PureASO,0) =1 and coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0)
											OR
											(GRPCompany.grpcocode in (''W058'',''W003'',''A139'',''C083'') and coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0)
										) 
										'
								end
								else if @isPureAso=2
								begin
									set @sql =  @sql + 'pln.ASOType=1 and coalesce(Claim.BTBGBordxStatus,0)=0 and coalesce(Claim.CLMBTBGAmt,0) <> 0'
								end
								else
								begin

								set @sql =  @sql + '
										(
											(pln.PLNBTBG =1 and coalesce(Claim.BTBGBordxStatus,0)=0 and coalesce(Claim.CLMBTBGAmt,0) <> 0) 
											OR 
											(GRPCompany.grpcocode in (''W058'',''W003'',''A139'',''C083'') and coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0)
										) 
										'
								end
							else IF @bordxType=2
								set @sql =  @sql + ' 
										(

										( pln.PLNBTBG=1

										and coalesce(Claim.CLMBTBGAmt,0) = coalesce(Claim.CLMTotalActual,0)

										and coalesce(BTBGBordxStatus,0)<>1  and coalesce(Claim.ClaimBordxStatus,0)=0 ) 
										Or 
										(
										coalesce(Claim.BTBGBordxStatus,0)=0 and coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0
										)
										) '
							else IF @bordxType=4
								set @sql =  @sql + ' coalesce(Claim.BTBGBordxStatus,0)=0 and (coalesce(CLMBordxAmt,0)<>0 and coalesce(claim.CLMBTBGAmt,0)=0) and (coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0)'										
							else 
							set @sql =  @sql + 'coalesce(Claim.ClaimBordxStatus,0)=0 and coalesce(Claim.CLMBordxAmt,0) <> 0'


			set @sql =  @sql + ' and CISCas.CASStatus NOT IN (''D'' ,''C'')
							AND (CISCas.OPBatchNo IS NOT NULL) AND (CISCas.OPBatchNo <> N'''')
							AND  (CISWorksheet.dbo.Claim.CLMBordxMemo = ''Y'' OR  CISWorksheet.dbo.Claim.CLMBordxMemo = '''' OR CISWorksheet.dbo.Claim.CLMBordxMemo IS NULL)  
							AND  ( HISDB.dbo.MBM.HSCLNInd <> ''H'') AND  (CISDB.dbo.CISCas.pending <> ''Y'') 
							AND CISCas.Claimability =''A'' 
							AND  0=0 '
		
		set @sql = @sql + @filter
				   --print @sql
		exec sp_executesql @sql, N'@groupCode nvarchar(10), @caseFrom nvarchar(30), @caseTo nvarchar(30), @visitFrom datetime, @visitTo datetime,
            @department nvarchar(200), @opFrom nvarchar(30), @opTo nvarchar(30), @subsidiary nvarchar(200), @treatment nvarchar(10), @regFrom datetime,
            @regTo datetime, @clinic nvarchar(20)',
            @groupCode, @caseFrom, @caseTo, @visitFrom, @visitTo, @department, @opFrom, @opTo, @subsidiary, @treatment, @regFrom, @regTo, @clinic;
END
GO

-- @flex 1: CboUOB flexCare (case 23), 0: outpatient (case 24); cash type C / M as @payTo (the UOB filter depends on both).
CREATE OR ALTER PROCEDURE cis.BordxWs_SearchUob @bordxType tinyint, @flex bit,
    @groupCode nvarchar(10) = NULL, @scan bit = 0, @caseFrom nvarchar(30) = NULL, @caseTo nvarchar(30) = NULL,
    @visitFrom datetime = NULL, @visitTo datetime = NULL, @noVisitDate bit = 0, @department nvarchar(200) = NULL, @clinicType int = 0,
    @opFrom nvarchar(30) = NULL, @opTo nvarchar(30) = NULL, @subsidiary nvarchar(200) = NULL, @payTo char(1) = NULL, @treatment nvarchar(10) = NULL,
    @omni bit = 0, @regFrom datetime = NULL, @regTo datetime = NULL, @clinic nvarchar(20) = NULL, @entity int = 0 AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @filter nvarchar(max) = N''
        + CASE WHEN @groupCode IS NOT NULL THEN N' AND CISWorksheet.dbo.Claim.CLMGrpComCode = @groupCode' ELSE N'' END
        + CASE WHEN @scan = 1 AND @caseFrom IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber = @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseFrom IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber >= @caseFrom' ELSE N'' END
        + CASE WHEN @scan = 0 AND @caseTo IS NOT NULL THEN N' AND CISDB.dbo.CISCas.CASNumber <= @caseTo' ELSE N'' END
        + CASE WHEN @visitFrom IS NOT NULL THEN N' And CISCas.DateVisit >= @visitFrom' ELSE N'' END
        + CASE WHEN @visitTo IS NOT NULL THEN N' And CISCas.DateVisit <= @visitTo' ELSE N'' END
        + CASE WHEN @noVisitDate = 1 THEN N' AND (CISCas.DateVisit Is null OR CISCas.DateVisit = '''')' ELSE N'' END
        + CASE WHEN @department IS NOT NULL THEN N' AND MBMOthers.MBMDepartment = @department' ELSE N'' END
        + CASE @clinicType WHEN 1 THEN N' AND CLN.CLNType = ''S''' WHEN 2 THEN N' AND CLN.CLNType = ''G'''
            WHEN 3 THEN N' AND (CLN.CLNType <> ''G'') AND (CLN.CLNType <> ''S'')' ELSE N'' END
        + CASE WHEN @opFrom IS NOT NULL THEN N' AND CISCas.OPBatchNo >= @opFrom' ELSE N'' END
        + CASE WHEN @opTo IS NOT NULL THEN N' AND CISCas.OPBatchNo <= @opTo' ELSE N'' END
        + CASE WHEN @subsidiary IS NOT NULL THEN N' AND MBMOthers.Subsidiary = @subsidiary' ELSE N'' END
        + CASE @payTo WHEN 'C' THEN N' AND CISWorksheet.dbo.ClaimInvoice.INVPayto = ''C''' WHEN 'M' THEN N' AND CISWorksheet.dbo.ClaimInvoice.INVPayto = ''M''' ELSE N'' END
        + CASE WHEN @treatment IS NOT NULL THEN N' AND CASTreatReason1 = @treatment' ELSE N'' END
        + CASE WHEN @omni = 1 THEN N' AND GRPCompany.GRPOMNICLM = ''Y''' ELSE N'' END
        + CASE WHEN @regFrom IS NOT NULL THEN N' And CISCas.RegisterDate >= @regFrom' ELSE N'' END
        + CASE WHEN @regTo IS NOT NULL THEN N' And CISCas.RegisterDate <= @regTo' ELSE N'' END
        + CASE WHEN @clinic IS NOT NULL THEN N' AND CISCas.ClinicCode = @clinic' ELSE N'' END
        + CASE @entity WHEN 0 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) <>''SJ'' ' WHEN 1 THEN N' AND Substring(CISCAS.CasNumber, 1, 2) =''SJ'' ' ELSE N'' END
        + CASE WHEN @payTo = 'C' AND @flex = 0 THEN N' AND ((tblFlexibleBenefitPlanLimit_OP.CISPlnCode is NOT null and tblFlexibleBenefitPlan_OP.FBCode is null) OR (tblFlexibleBenefitPlanLimit_OP.CISPlnCode is null) or (tblFlexibleBenefitPlan_OP.FBFamily=1 and ciscas.CoverID<>''00'') or tblFlexibleBenefitPlan_OP.FBisOPFlexCombine = 1 )'
            WHEN @payTo = 'C' AND @flex = 1 THEN N' AND (tblFlexibleBenefitPlan_OP.FBEmployee is not null and CISCas.CoverID=''00'') AND (tblFlexibleBenefitPlan_OP.FBisOPFlexCombine <> 1) '
            WHEN @payTo = 'M' AND @flex = 0 THEN N' AND ((tblFlexibleBenefitPlanLimit_OP.CISPlnCode is NOT null and tblFlexibleBenefitPlan_OP.FBCode is null) OR (tblFlexibleBenefitPlanLimit_OP.CISPlnCode is null) or (tblFlexibleBenefitPlan_OP.FBFamily=1 and ciscas.CoverID<>''00''))'
            WHEN @payTo = 'M' AND @flex = 1 THEN N' AND (tblFlexibleBenefitPlanLimit_OP.CISPlnCode is not null and tblFlexibleBenefitPlan_OP.FBEmployee=1 and ciscas.CoverID=''00'')'
            ELSE N'' END;
    IF @flex = 1
    BEGIN
	--flexCARE
			       declare @sql nvarchar(max)

		                    set @sql =  N'SELECT TOP (100) PERCENT CISDB.dbo.CISCas.CASNumber, CISDB.dbo.CISCas.MBMnumber, CISDB.dbo.CISCas.CoverID, CISCas.CASStatus, CISCas.DateVisit, 
												CISCas.RegisterDate, HISDB.dbo.MBM.MBMName, HISDB.dbo.MBMCLNOthers.MBMPolNumber AS MBMPolicyNo, 
												HISDB.dbo.MBMCoveredPersons.MBMCName, HISDB.dbo.CLN.CLNName, HISDB.dbo.CLN.CLNCode, CISDB.dbo.GRPCompany.GRPCoName as MBMGroupCompany, 
												HISDB.dbo.MBMCoveredPersons.MBMCRELCode, HISDB.dbo.MBMOthers.RELCode, HISDB.dbo.MBM.HLTCode, HISDB.dbo.MBM.MBMMemberType, 
												CISCas.CASTreatReason1, CISCas.CASTreatReason2, CISCas.CASTreatReason3, CISCas.CASTreatReason4, CISCas.CASTreatReason5,                        
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason1) as ReasonDesc1,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason2) as ReasonDesc2,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason3) as ReasonDesc3,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason4) as ReasonDesc4,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason5) as ReasonDesc5,
												CISCas.ClinicCode, CISCas.Remarks, 
												HISDB.dbo.MBM.AGECode, HISDB.dbo.MBM.HSCLNInd, HISDB.dbo.MBMCLNOthers.PLNType AS PLNCode, HISDB.dbo.MBMCLNOthers.INSType AS INSCode, 
												CISCas.CASDiagnosis5, HISDB.dbo.MBMCLNOthers.MBMPolNumber, HISDB.dbo.MBMCLNOthers.PAYCode, HISDB.dbo.MBMOthers.MBMDepartment, 
												HISDB.dbo.MBMOthers.MBMEmployeeNo, CISCas.InvoiceDate, HISDB.dbo.MBMOthers.MBMCostCenter, GRPCompany.GRPBordxFormat, 
												GRPCompany.RptFormat, HISDB.dbo.CLN.CLNType, GRPCompany.TyphoidRpt, HISDB.dbo.MBMOthers.MBMBranch, CISCas.CaseType, 
												CISCas.MCDays, HISDB.dbo.MBMOthers.MBMPosition, CISCas.OPBatchNo, CISCas.InvoiceNo, HISDB.dbo.MBM.MBMPolicyNo AS MBMHISPolNo, 
												HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMGroupCategory, HISDB.dbo.CLN.CLNFFSStatus, GRPCompany.GRPMCOStatus, 
												CISCas.CASClaimsApp, HISDB.dbo.MBMCoveredPersons.MBMAXANumber, HISDB.dbo.MBMOthers.MBMPrevMemNo, CISCas.MedInjType1, 
												CISCas.MedInjType2, CISCas.MedInjType3, CISCas.MedInjType4, CISCas.MedInjType5, CISCas.MedInjType6, CISCas.MedInjType7, 
												CISWorksheet.dbo.ClaimInvoiceActualOthers.GPMed, CISWorksheet.dbo.ClaimInvoiceActualOthers.SPMed, HISDB.dbo.CLN.CLNLocation, HISDB.dbo.CLN.CLNPState,
												HISDB.dbo.MBM.PLNCode AS HISPLN, HISDB.dbo.CLN.CLNState, CISCas.NPClinic, CISWorksheet.dbo.Claim.CASNumber AS Expr1, 
												CISWorksheet.dbo.Claim.ClaimStatus, CISWorksheet.dbo.Claim.ClaimBordxStatus, CISWorksheet.dbo.Claim.CLMPayableByMedix2C, 
												CISWorksheet.dbo.Claim.CLMTotalActual, CISWorksheet.dbo.Claim.CLMSubsidiary, CISWorksheet.dbo.Claim.CLMDeptName,'
						  IF @bordxType=0
							set @sql =  @sql + ' CASE WHEN Pln.PLNBTBG = 1 THEN  CISWorksheet.dbo.Claim.CLMBordxAmt - coalesce(Claim.CLMBTBGAmt,0) else CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '							
						 ELSE
							set @sql =  @sql + ' CASE WHEN Pln.PLNBTBG = 1 THEN CISWorksheet.dbo.Claim.CLMBTBGAmt ELSE CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '
		
		  set @sql =  @sql + 'CISWorksheet.dbo.Claim.CLMNonGPFee, 
											CISWorksheet.dbo.Claim.CLMMCOProcFee, CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISWorksheet.dbo.ClaimInvoice.INVBillAmt, 
											CISWorksheet.dbo.Claim.CLMFinalApprove AS CLMTotalApproved, GRPCompany.GRPOMNICLM, CISWorksheet.dbo.Claim.CLMBordxMemo, 
											GRPCompany.ASO_INS,lg.DoctorName,CISWorksheet.dbo.claim.CLMCoPayColl,MBM.PAYCode AS MemberPayCode,claim.clmcostcenter
											From cisdb..ciscas
											INNER JOIN cisdb..TreatmentReason on TreatmentReason.ReasonCode = case when coalesce(ciscas.CASTreatReason1,'''')='''' then ciscas.CASTreatReason5 else ciscas.CASTreatReason1 end
											INNER JOIN HISDB..MBM ON MBM.MBMNumber = CISCas.MBMNumber
											INNER JOIN HISDB..MBMCLNOthers ON MBMCLNOthers.MBMNumber = MBM.MBMNUMBER
											INNER JOIN hisdb..MBMOthers on mbmothers.MBMnumber=mbm.MBMNumber
											LEFT JOIN HISDB.dbo.MBMCoveredPersons on MBMCoveredPersons.MBMCNumber=CISCas.MBMNumber and MBMCoveredPersons.MBMCCoverID = CISCas.CoverID
											LEFT JOIN cisdb..pln on pln.PLNCode= MBMCLNOthers.PLNType
											INNER JOIN hisdb..cln on cln.CLNCode=ciscas.ClinicCode
											INNER JOIN CISWorksheet..claim on claim.CASNumber=ciscas.CASNumber
											LEFT OUTER JOIN CISDB.dbo.GRPCompany on GRPCoCode=claim.clmgrpcomcode 
											INNER JOIN CISWorksheet..ClaimInvoice on ClaimInvoice.CASNumber=claim.CASNumber
											LEFT JOIN CISWorksheet.dbo.ClaimInvoiceActualOthers on ClaimInvoiceActualOthers.CASNumber = ClaimInvoice.CASNumber
											LEFT JOIN CISDB..LG ON LG.LGNo=CISCAS.CASNumber 
											LEFT JOIN cisdb..tblFlexibleBenefitPlan_OP on tblFlexibleBenefitPlan_OP.FBCode=TreatmentReason.ReasonCode
											INNER JOIN CISDB..tblFlexibleBenefitPlanLimit_OP on tblFlexibleBenefitPlanLimit_OP.CISPlnCode = MBMCLNOthers.PLNType  and (ciscas.DateVisit between tblFlexibleBenefitPlanLimit_OP.FBEffDate and tblFlexibleBenefitPlanLimit_OP.FBExpDate)
											where '

							IF @bordxType=0
								set @sql =  @sql + '(coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0) '								
							ELSE																
								set @sql =   @sql + '(coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0) '	

							set @sql =  @sql + 'AND CISCas.CASStatus <>''D''
							AND coalesce(CISCas.OPBatchNo,'''')<>''''
							AND coalesce(Claim.CLMBordxMemo,'''')<>''N''
							AND  ( CISDB.dbo.CISCas.pending <> ''Y'')  
							AND 0=0 '
							
		set @sql = @sql + @filter
		exec sp_executesql @sql, N'@groupCode nvarchar(10), @caseFrom nvarchar(30), @caseTo nvarchar(30), @visitFrom datetime, @visitTo datetime,
            @department nvarchar(200), @opFrom nvarchar(30), @opTo nvarchar(30), @subsidiary nvarchar(200), @treatment nvarchar(10), @regFrom datetime,
            @regTo datetime, @clinic nvarchar(20)',
            @groupCode, @caseFrom, @caseTo, @visitFrom, @visitTo, @department, @opFrom, @opTo, @subsidiary, @treatment, @regFrom, @regTo, @clinic;
    END
    ELSE
    BEGIN
	--op
			       SET @sql = N''

		                    set @sql =  N'SELECT TOP (100) PERCENT CISDB.dbo.CISCas.CASNumber, CISDB.dbo.CISCas.MBMnumber, CISDB.dbo.CISCas.CoverID, CISCas.CASStatus, CISCas.DateVisit, 
												CISCas.RegisterDate, HISDB.dbo.MBM.MBMName, HISDB.dbo.MBMCLNOthers.MBMPolNumber AS MBMPolicyNo, 
												HISDB.dbo.MBMCoveredPersons.MBMCName, HISDB.dbo.CLN.CLNName, HISDB.dbo.CLN.CLNCode, CISDB.dbo.GRPCompany.GRPCoName as MBMGroupCompany, 
												HISDB.dbo.MBMCoveredPersons.MBMCRELCode, HISDB.dbo.MBMOthers.RELCode, HISDB.dbo.MBM.HLTCode, HISDB.dbo.MBM.MBMMemberType, 
												CISCas.CASTreatReason1, CISCas.CASTreatReason2, CISCas.CASTreatReason3, CISCas.CASTreatReason4, CISCas.CASTreatReason5,                        
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason1) as ReasonDesc1,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason2) as ReasonDesc2,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason3) as ReasonDesc3,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason4) as ReasonDesc4,
												cisdb.dbo.fnGetTreatmentDesc(CASTreatReason5) as ReasonDesc5,
												CISCas.ClinicCode, CISCas.Remarks, 
												HISDB.dbo.MBM.AGECode, HISDB.dbo.MBM.HSCLNInd, HISDB.dbo.MBMCLNOthers.PLNType AS PLNCode, HISDB.dbo.MBMCLNOthers.INSType AS INSCode, 
												CISCas.CASDiagnosis5, HISDB.dbo.MBMCLNOthers.MBMPolNumber, HISDB.dbo.MBMCLNOthers.PAYCode, HISDB.dbo.MBMOthers.MBMDepartment, 
												HISDB.dbo.MBMOthers.MBMEmployeeNo, CISCas.InvoiceDate, HISDB.dbo.MBMOthers.MBMCostCenter, GRPCompany.GRPBordxFormat, 
												GRPCompany.RptFormat, HISDB.dbo.CLN.CLNType, GRPCompany.TyphoidRpt, HISDB.dbo.MBMOthers.MBMBranch, CISCas.CaseType, 
												CISCas.MCDays, HISDB.dbo.MBMOthers.MBMPosition, CISCas.OPBatchNo, CISCas.InvoiceNo, HISDB.dbo.MBM.MBMPolicyNo AS MBMHISPolNo, 
												HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMGroupCategory, HISDB.dbo.CLN.CLNFFSStatus, GRPCompany.GRPMCOStatus, 
												CISCas.CASClaimsApp, HISDB.dbo.MBMCoveredPersons.MBMAXANumber, HISDB.dbo.MBMOthers.MBMPrevMemNo, CISCas.MedInjType1, 
												CISCas.MedInjType2, CISCas.MedInjType3, CISCas.MedInjType4, CISCas.MedInjType5, CISCas.MedInjType6, CISCas.MedInjType7, 
												CISWorksheet.dbo.ClaimInvoiceActualOthers.GPMed, CISWorksheet.dbo.ClaimInvoiceActualOthers.SPMed, HISDB.dbo.CLN.CLNLocation, HISDB.dbo.CLN.CLNPState,
												HISDB.dbo.MBM.PLNCode AS HISPLN, HISDB.dbo.CLN.CLNState, CISCas.NPClinic, CISWorksheet.dbo.Claim.CASNumber AS Expr1, 
												CISWorksheet.dbo.Claim.ClaimStatus, CISWorksheet.dbo.Claim.ClaimBordxStatus, CISWorksheet.dbo.Claim.CLMPayableByMedix2C, 
												CISWorksheet.dbo.Claim.CLMTotalActual, CISWorksheet.dbo.Claim.CLMSubsidiary, CISWorksheet.dbo.Claim.CLMDeptName,'
						  IF @bordxType=0
							set @sql =  @sql + ' CASE WHEN Pln.PLNBTBG = 1 THEN  CISWorksheet.dbo.Claim.CLMBordxAmt - coalesce(Claim.CLMBTBGAmt,0) else CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '							
						 ELSE
							set @sql =  @sql + ' CASE WHEN Pln.PLNBTBG = 1 THEN CISWorksheet.dbo.Claim.CLMBTBGAmt ELSE CISWorksheet.dbo.Claim.CLMBordxAmt END AS CLMBordxAmt, '
		
		  set @sql =  @sql + 'CISWorksheet.dbo.Claim.CLMNonGPFee, 
											CISWorksheet.dbo.Claim.CLMMCOProcFee, CISWorksheet.dbo.ClaimInvoice.INVPayTo, CISWorksheet.dbo.ClaimInvoice.INVBillAmt, 
											CISWorksheet.dbo.Claim.CLMFinalApprove AS CLMTotalApproved, GRPCompany.GRPOMNICLM, CISWorksheet.dbo.Claim.CLMBordxMemo, 
											GRPCompany.ASO_INS,lg.DoctorName,CISWorksheet.dbo.claim.CLMCoPayColl,MBM.PAYCode AS MemberPayCode,claim.clmcostcenter
											From cisdb..ciscas
											LEFT JOIN cisdb..TreatmentReason on TreatmentReason.ReasonCode = case when coalesce(ciscas.CASTreatReason1,'''')='''' then ciscas.CASTreatReason5 else ciscas.CASTreatReason1 end
											INNER JOIN HISDB..MBM ON MBM.MBMNumber = CISCas.MBMNumber
											INNER JOIN HISDB..MBMCLNOthers ON MBMCLNOthers.MBMNumber = MBM.MBMNUMBER
											INNER JOIN hisdb..MBMOthers on mbmothers.MBMnumber=mbm.MBMNumber
											LEFT JOIN HISDB.dbo.MBMCoveredPersons on MBMCoveredPersons.MBMCNumber=CISCas.MBMNumber and MBMCoveredPersons.MBMCCoverID = CISCas.CoverID
											LEFT JOIN cisdb..pln on pln.PLNCode= MBMCLNOthers.PLNType
											INNER JOIN hisdb..cln on cln.CLNCode=ciscas.ClinicCode
											INNER JOIN CISWorksheet..claim on claim.CASNumber=ciscas.CASNumber
											LEFT OUTER JOIN CISDB.dbo.GRPCompany on GRPCoCode=claim.clmgrpcomcode 
											INNER JOIN CISWorksheet..ClaimInvoice on ClaimInvoice.CASNumber=claim.CASNumber
											LEFT JOIN CISWorksheet.dbo.ClaimInvoiceActualOthers on ClaimInvoiceActualOthers.CASNumber = ClaimInvoice.CASNumber
											LEFT JOIN CISDB..LG ON LG.LGNo=CISCAS.CASNumber 
											LEFT JOIN cisdb..tblFlexibleBenefitPlan_OP on tblFlexibleBenefitPlan_OP.FBCode=TreatmentReason.ReasonCode
											LEFT JOIN CISDB..tblFlexibleBenefitPlanLimit_OP on tblFlexibleBenefitPlanLimit_OP.CISPlnCode = MBMCLNOthers.PLNType and (ciscas.DateVisit between tblFlexibleBenefitPlanLimit_OP.FBEffDate and tblFlexibleBenefitPlanLimit_OP.FBExpDate)
											where '

							IF @bordxType=0
								set @sql =  @sql + '(coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0) '								
							ELSE																
								set @sql =   @sql + '(coalesce(CISWorksheet.dbo.Claim.ClaimBordxStatus,0) = 0) '	

							set @sql =  @sql + 'AND CISCas.CASStatus <>''D''
							AND coalesce(CISCas.OPBatchNo,'''')<>''''
							AND coalesce(Claim.CLMBordxMemo,'''')<>''N''
							AND  ( CISDB.dbo.CISCas.pending <> ''Y'')  
							AND 0=0 '
							
		set @sql = @sql + @filter
		exec sp_executesql @sql, N'@groupCode nvarchar(10), @caseFrom nvarchar(30), @caseTo nvarchar(30), @visitFrom datetime, @visitTo datetime,
            @department nvarchar(200), @opFrom nvarchar(30), @opTo nvarchar(30), @subsidiary nvarchar(200), @treatment nvarchar(10), @regFrom datetime,
            @regTo datetime, @clinic nvarchar(20)',
            @groupCode, @caseFrom, @caseTo, @visitFrom, @visitTo, @department, @opFrom, @opTo, @subsidiary, @treatment, @regFrom, @regTo, @clinic;
    END
END
GO
