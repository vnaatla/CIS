// Builds mbmtrk-filters.sql (--PARAM / --WHERE lines) for gen-mbmtrk.js: FrmNewMBMTracking SearchRecord's criteria as parameters.
const fs = require('fs')
const P = []; const W = []
const add = (p, w) => { P.push(p); if (w) W.push(w) }
const blank = (name, col, alsoEmpty = true) => add(`@${name}Blank bit = 0`, `IF @${name}Blank = 1 SET @w += N' AND (${col} IS NULL${alsoEmpty ? ` OR ${col} = ''''` : ''})';`)
const eq = (name, col, type = 'nvarchar(200)') => add(`@${name} ${type} = NULL`, `IF @${name} IS NOT NULL SET @w += N' AND ${col} = @${name}';`)
const range = (name, col, type = 'nvarchar(50)') => {
  add(`@${name}From ${type} = NULL`, `IF @${name}From IS NOT NULL SET @w += N' AND ${col} >= @${name}From';`)
  add(`@${name}To ${type} = NULL`, `IF @${name}To IS NOT NULL SET @w += N' AND ${col} <= @${name}To';`)
}
eq('hlt', 'HLTCode', 'nvarchar(1)'); blank('hlt', 'HLTCode'); blank('payor', 'PAYCode')
eq('ins', 'INSCode', 'nvarchar(1)'); blank('ins', 'INSCode')
// General Listing (Prin. Name): member type x principal type.
// @prinType = '<member type>|<principal type>' (P / S / Both x Both / C / S), mapped to the desktop's fixed fragments.
add('@prinType nvarchar(10) = NULL', `IF @prinType IS NOT NULL SET @w += CASE @prinType WHEN 'P|Both' THEN N' AND (MBMMemberType = ''P'' or  MBMMemberType = ''Q'') AND MBMCCoverID = ''00''' WHEN 'P|C' THEN N' AND MBMMemberType = ''P'' AND MBMCCoverID = ''00''' WHEN 'P|S' THEN N' AND MBMMemberType = ''Q'' AND MBMCCoverID = ''00''' WHEN 'S|Both' THEN N' AND (MBMCMemberType = ''S'' OR MBMCMemberType = ''C'')' WHEN 'S|C' THEN N' AND MBMCMemberType = ''C''' WHEN 'S|S' THEN N' AND MBMCMemberType = ''S''' WHEN 'Both|C' THEN N' AND (MBMMemberType = ''P'' OR MBMCMemberType = ''C'')' WHEN 'Both|S' THEN N' AND (MBMMemberType = ''Q'' OR MBMCMemberType = ''S'')' WHEN 'Both|Both' THEN N' AND (MBMMemberType = ''P'' OR MBMMemberType = ''Q'')' ELSE N'' END;`)
range('cancel', 'MBMCancelDate', 'datetime'); blank('cancel', 'MBMCancelDate', false)
for (const [n, c] of [['dayInsured', 'DayInsured'], ['ageDays', 'AgeDays'], ['delayBordx', 'DelayBordx'], ['delayInvBordx', 'DelayInvBordx'], ['delayReg', 'DelayReg'],
  ['delayInvD', 'DelayInvD'], ['delayInvE', 'DelayInvE']]) { range(n, c, 'nvarchar(20)'); blank(n, c, false) }
range('premium', 'MBMPremium', 'money'); blank('premium', 'MBMPremium', false)
range('mco', 'MBMMCOFee', 'money'); blank('mco', 'MBMMCOFee', false)
range('employee', 'MBMEmployeeNo'); blank('employee', 'MBMEmployeeNo')
blank('memberType', 'MBMMemberType')
eq('adultChild', 'MBMAdultChild', 'nvarchar(1)'); blank('adultChild', 'MBMAdultChild')
eq('race', 'RACCode', 'nvarchar(1)'); blank('race', 'RACCode')
eq('sex', 'MBMSex', 'nvarchar(1)'); blank('sex', 'MBMSex')
eq('rel', 'RELCode', 'nvarchar(1)'); blank('rel', 'RELCode')
eq('dataStatus', 'MBMDataStatus', 'nvarchar(10)'); blank('dataStatus', 'MBMDataStatus')
add('@status nvarchar(1) = NULL', `IF @status IS NOT NULL SET @w += CASE @status WHEN '0' THEN N' AND MBMStatus = ''A''' WHEN '1' THEN N' AND MBMStatus = ''C''' WHEN '2' THEN N' AND MBMStatus = ''D''' WHEN '3' THEN N' AND (MBMStatus = ''A'' OR MBMStatus = ''C'')' WHEN '4' THEN N' AND (MBMStatus = ''A'' OR MBMStatus = ''S'')' WHEN '5' THEN N' AND (MBMStatus = ''A'' OR MBMStatus = ''C'' OR MBMStatus = ''S'')' WHEN '6' THEN N' AND MBMStatus = ''S''' ELSE N'' END;`)
blank('status', 'MBMStatus')
eq('expired', 'MBMExpiredStatus', 'nvarchar(1)'); blank('expired', 'MBMExpiredStatus')
add('@renewal nvarchar(1) = NULL', `IF @renewal IS NOT NULL SET @w += CASE @renewal WHEN 'P' THEN N' AND (MBMRenewFlag = ''P'')' WHEN 'N' THEN N' AND (MBMRenewFlag IS NULL OR MBMRenewFlag = '''' OR MBMRenewFlag = ''N'')' WHEN 'R' THEN N' AND (MBMRenewFlag = ''R'')' WHEN 'T' THEN N' AND (MBMRenewFlag = ''T'')' ELSE N'' END;`)
blank('renewal', 'MBMRenewFlag')
eq('premMode', 'MBMInstallment', 'nvarchar(1)'); blank('premMode', 'MBMInstallment')
eq('state', 'STACode', 'nvarchar(100)'); blank('state', 'STACode')
eq('takeover', 'MBMTakeover', 'nvarchar(1)'); blank('takeover', 'MBMTakeover')
eq('agent', 'MBMAgentCode', 'nvarchar(30)'); blank('agent', 'MBMAgentCode')
eq('mbmNo', 'MBMNumber', 'nvarchar(30)'); blank('mbmNo', 'MBMNumber')
eq('policy', 'MBMPolicyNo', 'nvarchar(60)'); blank('policy', 'MBMPolicyNo')
eq('payor', 'PAYCode', 'nvarchar(10)')
eq('company', 'MBMGroupCompany'); blank('company', 'MBMGroupCompany')
range('plan', 'PLNCode', 'nvarchar(20)'); blank('plan', 'PLNCode')
range('age', 'AGECode', 'nvarchar(20)'); blank('age', 'AGECode')
for (const [n, c] of [['eff', 'MBMPayorEffDate'], ['exp', 'MBMPayorExpDate'], ['joined', 'MBMFirstJoinedDate'], ['dor', 'MBMValueDate'], ['dob', 'MBMDOB'],
  ['bordx', 'MBMBordxDate']]) { range(n, c, 'datetime'); blank(n, c) }
for (const [n, c] of [['batch', 'MBMBatchNo'], ['postcode', 'MBMPostCode'], ['polYear', 'MBMPolicyYear'], ['invoice', 'MBMInvoiceNo']]) { range(n, c); blank(n, c) }
range('invoiceDate', 'MBMInvoiceDate', 'datetime'); blank('invoiceDate', 'MBMInvoiceDate')
range('ic', 'MBMIcBcPp'); blank('ic', 'MBMIcBcPp')
eq('label', 'LabelStatus', 'nvarchar(1)'); blank('label', 'LabelStatus')
eq('department', 'MBMDepartment'); blank('department', 'MBMDepartment')
range('card', 'MBMCLNCardSentDate', 'datetime'); blank('card', 'MBMCLNCardSentDate')
range('endt', 'MBMCLNEndtEffDate', 'datetime'); blank('endt', 'MBMCLNEndtEffDate')
range('suppBordx', 'MBMBordxDate', 'datetime')
// Detail listings only: principal type (all their queries) and the supplementary registration date (their supplementary queries).
add('@detailPrinType nvarchar(1) = NULL', "IF @detailPrinType = 'C' SET @w += N' AND (MBMMemberType = ''P'' OR MBMMemberType = ''C'')'; IF @detailPrinType = 'S' SET @w += N' AND (MBMMemberType = ''Q'')';")
range('suppReg', 'MBMCRegDate', 'datetime'); blank('suppReg', 'MBMCRegDate')
fs.writeFileSync('C:/Users/Naresh/AppData/Local/Temp/mbmtrk-filters.sql', [...P.map((p) => '--PARAM ' + p), ...W.map((w) => '--WHERE ' + w)].join('\n') + '\n')
console.log(P.length, 'params', W.length, 'where')
