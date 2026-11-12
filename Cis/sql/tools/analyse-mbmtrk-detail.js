// Detail listings of FrmNewMBMTracking: per routine, the cases (P / S / B) and the SubItems of each recordset loop.
const fs = require('fs')
const lines = fs.readFileSync('C:/Users/Naresh/AppData/Local/Temp/p10/FrmNewMBMTracking.txt', 'utf8').split(/\r?\n/)
const names = ['GeneralListing', 'GeneralListing2', 'PrintLabel', 'CardSentListing', 'GeneralListing4', 'AgentGeneralListing', 'MCOInvoice', 'CheckDate', 'ICListing', 'BordxDateListing']
const out = {}
for (const name of names) {
  const start = lines.findIndex((l) => new RegExp(`^Private Sub ${name}\\(`).test(l))
  let end = lines.findIndex((l, i) => i > start && /^Private (Sub|Function) /.test(l))
  if (end < 0) end = lines.length
  const body = lines.slice(start, end)
  const cases = body.map((l, i) => [i, /strAppendCase = "(\d+)"/.exec(l)?.[1]]).filter((x) => x[1])
  const loops = []
  let cur = null
  for (const l of body) {
    const lp = /^Do Until (\w+)\.EOF/.exec(l)
    if (lp) { cur = { rs: lp[1], first: null, cols: {} }; loops.push(cur); continue }
    if (!cur) continue
    const f = /ListItems\.Add\(, , (.*)\)$/.exec(l)
    if (f && !cur.first) cur.first = f[1]
    const s = /\.SubItems\((\d+)\) = (.*)$/.exec(l)
    if (s) (cur.cols[s[1]] ??= []).push(s[2].slice(0, 70))
  }
  const crit = body.findIndex((l) => /strCriteria = ""/.test(l)) >= 0
  out[name] = { cases: cases.map((c) => c[1]), ownCriteria: crit, extra: body.filter((l) => /strCriteria2 = |PrincipalType/.test(l)).slice(0, 6), loops }
}
fs.writeFileSync('C:/Users/Naresh/AppData/Local/Temp/mbmtrk-detail.json', JSON.stringify(out, null, 1))
for (const [k, v] of Object.entries(out)) console.log(k, 'cases', v.cases.join('/'), 'loops', v.loops.map((l) => l.rs + ':' + Object.keys(l.cols).length).join(' '))
