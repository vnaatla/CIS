// Claim Tracking 2 catalogue from FrmClaimRpt2 (or its Jupiter copy: TRK_FORM / TRK_CAT): amount basis (Option1 Actual / Option2 Approved) x report (Option4..20) ->
// routine, CISClaimTracking2 case, columns (f field, fe field or "Empty", d date, m0 money or "0", sum = total of the 28 fee columns) and headers (Option<n>_Click).
const fs = require('fs')
const FORM = process.env.TRK_FORM || 'FrmClaimRpt2'
const CAT = process.env.TRK_CAT || 'ClmTracking2.catalogue.json'
const lines = fs.readFileSync('C:/Users/Naresh/AppData/Local/Temp/p10/' + FORM + '.txt', 'utf8').split(/\r?\n/).map((l) => l.trim())
const body = {}
let cur = null
for (const l of lines) {
  const m = /^(?:Private|Public) (?:Sub|Function) (\w+)/.exec(l)
  if (m) { cur = m[1]; body[cur] = []; continue }
  if (cur) body[cur].push(l)
}
const dispatch = []
for (const l of body.SearchRecords ?? []) {
  const m = /Option([12])\.Value = True And Option(\d+)\.Value = True Then/.exec(l)
  if (m) dispatch.push({ basis: Number(m[1]), report: Number(m[2]) })
  const c = /^Call (\w+)\(strAppend\)/.exec(l)
  if (c && dispatch.length && !dispatch.at(-1).routine) dispatch.at(-1).routine = c[1]
}
const names = { 4: 'Claim', 5: 'Insured Type & Plan', 6: 'Insured Type & Ageband', 7: 'Insured Type & Gender', 8: 'Ageband', 9: 'Gender', 10: 'Ageband & Gender',
  11: 'Final Diagnosis', 12: 'Clinic', 13: 'Payor', 14: 'Payor, Insured Type & Plan', 15: 'Group Company', 16: 'Insured Type & Relation', 17: 'Insured Type',
  18: 'Bordx No', 19: 'Policy Year', 20: 'Takeover' }
const parse = (e) => {
  e = e.replace(/rst\d?!/g, '!').trim()
  let m
  if ((m = /^Format\(!(\w+), "dd\/mm\/yyyy"\)$/i.exec(e))) return { k: 'd', n: m[1] }
  if ((m = /^Format\(!(\w+), "#,##0\.00"\)$/.exec(e))) return { k: 'm0', n: m[1] }
  if ((m = /^(?:Trim\()?!(\w+)\)?$/.exec(e))) return { k: 'f', n: m[1] }
  if (/^Format\(total, "#,##0\.00"\)$/i.test(e)) return { k: 'sum' }
  return null
}
const headers = (report, basis) => {
  const out = []; let branch = 0
  for (const l of body[`Option${report}_Click`] ?? []) {
    const b = /^(?:Else)?If Option([12])\.Value = True Then/.exec(l)
    if (b) { branch = Number(b[1]); continue }
    const h = /ColumnHeaders\((\d+)\)\.Text = "([^"]*)"/.exec(l)
    if (h && (branch === 0 || branch === basis)) out[Number(h[1]) - 1] = h[2]
  }
  return [...out].map((x) => x ?? '')
}
const bad = []
const catalogue = dispatch.filter((d) => d.routine).map((d) => {
  const b = (body[d.routine] ?? []).map((l) => l.replace(/rst\d?!/g, '!'))
  const caseNo = Number(b.map((l) => /strAppendCase = "(\d+)"/.exec(l)?.[1]).find(Boolean))
  const alts = {}
  for (const l of b) {
    const a = /ListItems\.Add\(, , (.*)\)$/.exec(l)
    if (a) (alts[0] ??= []).push(a[1])
    const s = /\.SubItems\((\d+)\) = (.*)$/.exec(l)
    if (s) (alts[s[1]] ??= []).push(s[2])
  }
  const n = Math.max(...Object.keys(alts).map(Number))
  const cols = []
  for (let i = 0; i <= n; i++) {
    const list = alts[i] ?? []
    let c = list.length ? parse(list[0]) : { k: 'l' }
    if (c && c.k === 'f' && list.some((x) => x.trim() === '"Empty"')) c = { k: 'fe', n: c.n }
    if (!c) bad.push(`${d.routine}[${i}] ${list[0]}`)
    cols.push(c ?? { k: 'l' })
  }
  return { basis: d.basis, report: d.report, name: names[d.report], routine: d.routine, c: caseNo, cols, headers: Array.from({ length: n + 1 }, (_, i) => headers(d.report, d.basis)[i] ?? '') }
})
fs.writeFileSync('D:/CIS/Cis/src/Cis.Application/Reports/' + CAT, JSON.stringify(catalogue, null, 1))
console.log(catalogue.length, 'reports; unparsed:', bad.length, bad.slice(0, 8).join(' | '))
console.log('header count mismatches:', catalogue.filter((e) => e.headers.length !== e.cols.length).map((e) => `${e.routine} ${e.headers.length}/${e.cols.length}`).join(', ') || '-')
