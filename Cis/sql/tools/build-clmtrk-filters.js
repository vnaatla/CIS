// Builds clmtrk-filters.sql (--PARAM / --WHERE lines) for gen-trk.js: frmClaimRpt SearchRecords' criteria as parameters.
const fs = require('fs')
const P = []; const W = []
const add = (p, w) => { P.push(p); if (w) W.push(w) }
const blank = (name, col, alsoEmpty = true) => add(`@${name}Blank bit = 0`, `IF @${name}Blank = 1 SET @w += N' AND (${col} IS NULL${alsoEmpty ? ` OR ${col} = ''''` : ''})';`)
const eq = (name, col, type = 'nvarchar(200)') => add(`@${name} ${type} = NULL`, `IF @${name} IS NOT NULL SET @w += N' AND ${col} = @${name}';`)
const like = (name, col) => add(`@${name} nvarchar(200) = NULL`, `IF @${name} IS NOT NULL SET @w += N' AND ${col} LIKE ''%'' + @${name} + ''%''';`)
const range = (name, col, type = 'nvarchar(50)') => {
  add(`@${name}From ${type} = NULL`, `IF @${name}From IS NOT NULL SET @w += N' AND ${col} >= @${name}From';`)
  add(`@${name}To ${type} = NULL`, `IF @${name}To IS NOT NULL SET @w += N' AND ${col} <= @${name}To';`)
}
const choice = (name, cases) => add(`@${name} nvarchar(10) = NULL`, `IF @${name} IS NOT NULL SET @w += CASE @${name} ${cases.map(([v, f]) => `WHEN '${v}' THEN N'${f.replace(/'/g, "''")}'`).join(' ')} ELSE N'' END;`)
eq('hlt', 'HLTCode', 'nvarchar(1)'); blank('hlt', 'HLTCode')
eq('ins', 'INSCode', 'nvarchar(1)'); blank('ins', 'INSCode')
// Member type x supplementary type (CboMemType / cboSuppType), code '<P|S|Both>|<Both|C|S>'.
choice('memType', [['P|Both', " AND (MBMMemberType = 'P' Or MBMMemberType = 'Q') And CoverID = '00' "], ['P|C', " AND MBMMemberType = 'P'  And CoverID = '00'"],
  ['P|S', " AND MBMMemberType = 'Q' And CoverID = '00'"], ['S|Both', " AND (MBMCMemberType = 'S' OR MBMCMemberType = 'C')"], ['S|C', " AND MBMCMemberType = 'C'"],
  ['S|S', " AND MBMCMemberType = 'S'"], ['Both|C', " AND (MBMMemberType = 'P' OR MBMCMemberType = 'C')"], ['Both|S', " AND (MBMMemberType = 'Q' OR MBMCMemberType = 'S')"]])
for (const [n, c] of [['race', 'RACCode'], ['sex', 'Sex'], ['relation', 'Relation'], ['dataStatus', 'MBMDataStatus'], ['claimability', 'Claimability'],
  ['polStatus', 'MBMStatus'], ['expired', 'MBMExpiredStatus'], ['premMode', 'MBMInstallment'], ['takeover', 'MBMTakeOver']]) { eq(n, c, 'nvarchar(1)'); blank(n, c) }
choice('caseStatus', [['0', " AND CASStatus = 'O'"], ['1', " AND (CASStatus = 'O' OR CASStatus = 'C')"], ['2', " AND CASStatus = 'D'"]]); blank('caseStatus', 'CASStatus')
eq('state', 'STACode', 'nvarchar(100)'); blank('state', 'STACode')
blank('department', 'MBMDepartment')
choice('renewal', [['P', " AND (MBMRenewFlag = 'P')"], ['N', " AND (MBMRenewFlag Is null OR MBMRenewFlag = '' OR MBMRenewFlag = 'N') "], ['R', " AND (MBMRenewFlag = 'R')"],
  ['T', " AND (MBMRenewFlag = 'T')"]]); blank('renewal', 'MBMRenewFlag')
like('agent', 'MBMAgentCode'); blank('agent', 'MBMAgentCode')
like('policy', 'MBMPolicyNo'); blank('policy', 'MBMPolicyNo')
like('patient', 'PatientName')
range('ageband', 'AGEBand', 'nvarchar(20)')
range('polYear', 'MBMPolicyYear', 'nvarchar(20)'); blank('polYear', 'MBMPolicyYear')
range('mcDays', 'MCDays', 'nvarchar(20)'); blank('mcDays', 'MCDays')
// The desktop procedure turns PAYCode = 'DC' into PAYCode in ('DC','EA').
add('@payor nvarchar(10) = NULL', "IF @payor = 'DC' SET @w += N' AND PAYCode in (''DC'',''EA'')'; ELSE IF @payor IS NOT NULL SET @w += N' AND PAYCode = @payor';"); blank('payor', 'PAYCode')
eq('company', 'MBMGroupCompany'); blank('company', 'MBMGroupCompany')
eq('clinic', 'ClinicCode', 'nvarchar(20)'); blank('clinic', 'CLNName')
eq('clinicGroup', 'ClinicGRPName')
eq('branch', 'BRNCode', 'nvarchar(20)')
for (const [n, c] of [['dov', 'DateVisit'], ['dor', 'RegisterDate'], ['dow', 'ClaimOpenDate'], ['invoiceDate', 'InvoiceDate']]) { range(n, c, 'datetime'); blank(n, c) }
range('treatment', 'CASTreatReason1', 'nvarchar(20)'); blank('treatment', 'CASTreatReason1')
range('plan', 'PLNCode', 'nvarchar(20)'); blank('plan', 'PLNCode')
for (const [n, c] of [['eff', 'MBMPayorEffDate'], ['exp', 'MBMPayorExpDate'], ['joined', 'MBMFirstJoinedDate'], ['dob', 'DOB']]) { range(n, c, 'datetime'); blank(n, c) }
range('mbmNo', 'MBMNumber', 'nvarchar(30)')
range('caseNo', 'CASNumber', 'nvarchar(20)')
// Bordx no. boxes apply to the insurer bordx (BordxRefNo) or the ASO one (BTBGBordxRefNo) by the bordx type; without a type they are ignored, as the desktop.
add('@bordxType nvarchar(10) = NULL', null)
add('@bordxFrom nvarchar(20) = NULL', `IF @bordxFrom IS NOT NULL AND @bordxType = 'INSURER' SET @w += N' AND BordxRefNo >= @bordxFrom'; IF @bordxFrom IS NOT NULL AND @bordxType = 'ASO' SET @w += N' AND BTBGBordxRefNo >= @bordxFrom';`)
add('@bordxTo nvarchar(20) = NULL', `IF @bordxTo IS NOT NULL AND @bordxType = 'INSURER' SET @w += N' AND BordxRefNo <= @bordxTo'; IF @bordxTo IS NOT NULL AND @bordxType = 'ASO' SET @w += N' AND BTBGBordxRefNo <= @bordxTo';`)
add('@bordxBlank bit = 0', `IF @bordxBlank = 1 AND @bordxType = 'INSURER' SET @w += N' AND (BordxRefNo IS NULL OR BordxRefNo = '''')'; IF @bordxBlank = 1 AND @bordxType = 'ASO' SET @w += N' AND (BTBGBordxRefNo IS NULL OR BTBGBordxRefNo = '''')';`)
range('bordxDate', 'BordxDate', 'datetime'); blank('bordxDate', 'BordxDate')
range('ic', 'ICNo', 'nvarchar(30)'); blank('ic', 'ICNo')
range('mbmIc', 'MBMIcBcPp', 'nvarchar(30)'); blank('mbmIc', 'MBMIcBcPp')
range('actual', 'CLMTotalActual', 'money'); range('approved', 'CLMTotalApproved', 'money'); range('osUtilized', 'OSUtilized', 'money')
choice('panel', [['C', " AND CLNPanel = 'C'"], ['Y', " AND CLNPanel = 'Y'"]])
// The clinic type / plan code by panel reports' non-panel figures (CLNByTypePanel, PlanByPanel).
add('@nonPanel bit = 0', "IF @nonPanel = 1 SET @w += N' AND (CLNPanel <> ''C'' OR CLNPanel IS NULL OR CLNPanel = '''')';")
choice('pending', [['N', " AND (CLMPending = 'N' OR CLMPending = '' OR CLMPending IS NULL)"], ['Y', " AND CLMPending = 'Y'"]])
eq('year', 'DOVYear', 'nvarchar(4)')
range('invoice', 'InvoiceNo', 'nvarchar(50)'); blank('invoice', 'InvoiceNo')
range('submission', 'BillNo', 'nvarchar(50)'); blank('submission', 'BillNo')
range('opBatch', 'OPBatchNo', 'nvarchar(30)')
range('clinicCode', 'ClinicCode', 'nvarchar(20)')
choice('clinicType', [['0', " AND CLNType = 'S'"], ['1', " AND CLNType = 'G'"], ['2', " AND (CLNType = '' OR CLNType IS NULL)"]])
eq('payTo', 'INVPayTo', 'nvarchar(1)')
add('@vsAmt bit = 0', `IF @vsAmt = 1 SET @w += N' AND VSAmt <> 0';`)
eq('deptSelection', 'MBMDepartment')
add('@ieInvoice bit = 0', `IF @ieInvoice = 1 SET @w += N' AND substring(InvoiceNo,1,2)=''IE'' ';`)
// The Simple procedure (CISClaimTracking_Simple) has no DC -> DC / EA rewrite: its payor filter is a plain comparison.
const simple = [...P.map((p) => '--PARAM ' + p), ...W.map((w) => '--WHERE ' + (w.startsWith("IF @payor = 'DC'") ? "IF @payor IS NOT NULL SET @w += N' AND PAYCode = @payor';" : w))]
fs.writeFileSync('C:/Users/Naresh/AppData/Local/Temp/clmtrk-simple-filters.sql', simple.join('\n') + '\n')
fs.writeFileSync('C:/Users/Naresh/AppData/Local/Temp/clmtrk-filters.sql', [...P.map((p) => '--PARAM ' + p), ...W.map((w) => '--WHERE ' + w)].join('\n') + '\n')
console.log(P.length, 'params', W.length, 'where')
