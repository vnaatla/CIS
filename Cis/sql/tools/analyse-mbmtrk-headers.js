// (report type, selection) -> column headers, from FrmNewMBMTracking's ListViewHeader1..18. A header routine without a selection test applies to its whole type
// (ListViewHeaderN = type N-1); ListViewHeader18 tests ListIndex = 18 although States is 17 (desktop bug: its headers were never set) and is mapped to 17.
const fs = require('fs')
const lines = fs.readFileSync('C:/Users/Naresh/AppData/Local/Temp/p10/FrmNewMBMTracking.txt', 'utf8').split(/\r?\n/)
const out = {}
let type = null, sel = null
for (const l of lines) {
  const hs = /^Private Sub ListViewHeader(\d+)\(\)/.exec(l)
  if (hs) { type = Number(hs[1]) - 1; sel = ''; continue }
  if (/^(Private|Public) (Sub|Function) /.test(l)) { type = null; sel = null; continue }
  if (type === null) continue
  const c = /CboReportType\.ListIndex = (\d+)(?: And CboRptSelection\.Text = "([^"]*)")?/.exec(l)
  if (c && /^(Else)?If /.test(l)) { const t = Number(c[1]); type = t === 18 ? 17 : t; sel = c[2] ?? ''; continue }
  const h = /ColumnHeaders\((\d+)\)\.Text = "([^"]*)"/.exec(l)
  if (h) (out[`${type}|${sel}`] ??= [])[Number(h[1]) - 1] = h[2]
}
for (const k of Object.keys(out)) { const a = out[k]; while (a.length && !a[a.length - 1]) a.pop(); out[k] = Array.from(a, (x) => x ?? '') }
fs.writeFileSync('C:/Users/Naresh/AppData/Local/Temp/mbmtrk-headers.json', JSON.stringify(out, null, 1))
console.log(Object.keys(out).length, 'header sets')
