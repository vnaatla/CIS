// Builds ct2-filters.sql (--PARAM / --WHERE lines) for gen-trk.js claim2: FrmClaimRpt2 SearchRecords' criteria as parameters.
const fs = require('fs')
const P = []; const W = []
const add = (p, w) => { P.push(p); if (w) W.push(w) }
const blank = (name, col) => add(`@${name}Blank bit = 0`, `IF @${name}Blank = 1 SET @w += N' AND (${col} IS NULL OR ${col} = '''')';`)
const eq = (name, col, type = 'nvarchar(200)') => add(`@${name} ${type} = NULL`, `IF @${name} IS NOT NULL SET @w += N' AND ${col} = @${name}';`)
const like = (name, col) => add(`@${name} nvarchar(200) = NULL`, `IF @${name} IS NOT NULL SET @w += N' AND ${col} LIKE ''%'' + @${name} + ''%''';`)
const range = (name, col, type = 'nvarchar(50)') => {
  add(`@${name}From ${type} = NULL`, `IF @${name}From IS NOT NULL SET @w += N' AND ${col} >= @${name}From';`)
  add(`@${name}To ${type} = NULL`, `IF @${name}To IS NOT NULL SET @w += N' AND ${col} <= @${name}To';`)
}
const choice = (name, cases) => add(`@${name} nvarchar(10) = NULL`, `IF @${name} IS NOT NULL SET @w += CASE @${name} ${cases.map(([v, f]) => `WHEN '${v}' THEN N'${f.replace(/'/g, "''")}'`).join(' ')} ELSE N'' END;`)
eq('hlt', 'HLTCode', 'nvarchar(1)'); blank('hlt', 'HLTCode')
eq('ins', 'INSCode', 'nvarchar(1)'); blank('ins', 'INSCode')
// Member type x supplementary type: this form uses MBMPatientCovID (frmClaimRpt: CoverID).
choice('memType', [['P|Both', " AND (MBMMemberType = 'P' Or MBMMemberType = 'Q') And MBMPatientCovID = '00' "], ['P|C', " AND MBMMemberType = 'P'  And MBMPatientCovID = '00'"],
  ['P|S', " AND MBMMemberType = 'Q' And MBMPatientCovID = '00'"], ['S|Both', " AND (MBMCMemberType = 'S' OR MBMCMemberType = 'C')"], ['S|C', " AND MBMCMemberType = 'C'"],
  ['S|S', " AND MBMCMemberType = 'S'"], ['Both|C', " AND (MBMMemberType = 'P' OR MBMCMemberType = 'C')"], ['Both|S', " AND (MBMMemberType = 'Q' OR MBMCMemberType = 'S')"]])
for (const [n, c] of [['race', 'RACCode'], ['sex', 'Sex'], ['relation', 'Relation'], ['dataStatus', 'MBMDataStatus'], ['claimability', 'Claimability'],
  ['polStatus', 'MBMStatus'], ['caseStatus', 'CASStatus'], ['expired', 'MBMExpiredStatus'], ['premMode', 'MBMInstallment'], ['takeover', 'MBMTakeOver']]) { eq(n, c, 'nvarchar(1)'); blank(n, c) }
eq('state', 'STACode', 'nvarchar(100)'); blank('state', 'STACode')
choice('renewal', [['P', " AND (MBMRenewFlag = 'P')"], ['N', " AND (MBMRenewFlag Is null OR MBMRenewFlag = '' OR MBMRenewFlag = 'N') "], ['R', " AND (MBMRenewFlag = 'R')"],
  ['T', " AND (MBMRenewFlag = 'T')"]]); blank('renewal', 'MBMRenewFlag')
like('agent', 'MBMAgentCode'); blank('agent', 'MBMAgentCode')
like('policy', 'MBMPolicyNo'); blank('policy', 'MBMPolicyNo')
like('patient', 'PatientName')
range('polYear', 'MBMPolicyYear', 'nvarchar(20)'); blank('polYear', 'MBMPolicyYear')
range('mcDays', 'MCDays', 'nvarchar(20)'); blank('mcDays', 'MCDays')
eq('payor', 'PAYCode', 'nvarchar(10)'); blank('payor', 'PAYCode')
eq('company', 'MBMGroupCompany'); blank('company', 'MBMGroupCompany')
eq('clinicName', 'HOSName'); blank('clinicName', 'HOSName')
eq('clinicGroup', 'HOSGRPName'); blank('clinicGroup', 'HOSGRPName')
eq('branch', 'BRNCode', 'nvarchar(20)'); blank('branch', 'BRNCode')
for (const [n, c] of [['dov', 'DateVisit'], ['dor', 'RegisterDate'], ['dow', 'ClaimOpenDate']]) { range(n, c, 'datetime'); blank(n, c) }
range('treatment', 'CASTreatReason1', 'nvarchar(20)'); blank('treatment', 'CASTreatReason1')
range('plan', 'PLNCode', 'nvarchar(20)'); blank('plan', 'PLNCode')
for (const [n, c] of [['eff', 'MBMPayorEffDate'], ['exp', 'MBMPayorExpDate'], ['joined', 'MBMFirstJoinedDate'], ['dob', 'DOB']]) { range(n, c, 'datetime'); blank(n, c) }
range('mbmNo', 'MBMNumber', 'nvarchar(30)')
range('caseNo', 'CASNumber', 'nvarchar(20)')
range('bordx', 'BordxRefNo', 'nvarchar(20)'); blank('bordx', 'BordxRefNo')
range('bordxDate', 'BordxDate', 'datetime'); blank('bordxDate', 'BordxDate')
range('ic', 'ICNo', 'nvarchar(30)'); blank('ic', 'ICNo')
range('mbmIc', 'MBMIcBcPp', 'nvarchar(30)'); blank('mbmIc', 'MBMIcBcPp')
range('actual', 'CLMTotalActual', 'money'); range('approved', 'CLMTotalApproved', 'money')
range('invoice', 'InvoiceNo', 'nvarchar(50)'); blank('invoice', 'InvoiceNo')
range('submission', 'BillNo', 'nvarchar(50)'); blank('submission', 'BillNo')
fs.writeFileSync('C:/Users/Naresh/AppData/Local/Temp/ct2-filters.sql', [...P.map((p) => '--PARAM ' + p), ...W.map((w) => '--WHERE ' + w)].join('\n') + '\n')
console.log(P.length, 'params', W.length, 'where')
