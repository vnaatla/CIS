// frmClaimRpt -> Cis/src/Cis.Application/Reports/ClmTracking.catalogue.json: report type / choice -> listing routine -> case(s), columns, headers.
// Same column kinds as the member tracking catalogue (gen-mbmtrk-spec.js); routines that do not fit are marked special.
const fs = require('fs')
const lines = fs.readFileSync('C:/Users/Naresh/AppData/Local/Temp/p10/frmClaimRpt.txt', 'utf8').split(/\r?\n/)
const body = {}
let cur = null
for (const l of lines) {
  const m = /^(?:Private|Public) (?:Sub|Function) (\w+)/.exec(l)
  if (m) { cur = m[1]; body[cur] = []; continue }
  if (cur) body[cur].push(l)
}
// Header routine per type (CboSelection_Click), headers per selection inside it.
const headerOf = {}
let t = null
for (const l of body.CboSelection_Click) {
  const c = /CboReportType\.ListIndex = (\d+)/.exec(l)
  if (c) t = Number(c[1])
  const h = /^Call (ListViewHeader\d+)/.exec(l.trim())
  if (h && t !== null) headerOf[t] = h[1]
}
const headers = (routine, sel) => {
  const out = {}; let s = ''
  for (const l of body[routine] ?? []) {
    const c = /CboSelection\.Text = "([^"]*)"/.exec(l)
    if (c && /^(Else)?If /.test(l)) s = c[1]
    const h = /ColumnHeaders\((\d+)\)\.Text = "([^"]*)"/.exec(l)
    if (h) (out[s] ??= [])[Number(h[1]) - 1] = h[2]
  }
  const a = out[sel] ?? out[''] ?? []
  while (a.length && !a[a.length - 1]) a.pop()
  return Array.from(a, (x) => x ?? '')
}
// Dispatch (SearchRecords).
const dispatch = []
let cond = null
for (const l of body.SearchRecords) {
  const c = /CboReportType\.ListIndex = (\d+)(?: And CboSelection\.Text = "([^"]*)")?\s+Then/.exec(l)
  if (c && /^(Else)?If /.test(l)) { cond = { type: Number(c[1]), selection: c[2] ?? '' }; continue }
  const call = /^Call (\w+)\(strAppend\)$/.exec(l.trim())
  if (call && cond) { if (!dispatch.some((d) => d.type === cond.type && d.selection === cond.selection)) dispatch.push({ ...cond, routine: call[1] }); cond = null }
}
const parse = (e) => {
  e = (e || '').trim().replace(/\b\w+!/g, '!')
  let m
  if ((m = /^!(\w+)$/.exec(e)) || (m = /^Trim\(!(\w+)\)$/.exec(e))) return { k: 'f', n: m[1] }
  if ((m = /^Format\(!(\w+), "#,##0\.00"\)$/.exec(e))) return { k: 'm', n: m[1] }
  if ((m = /^Format\(!(\w+), "dd\/mm\/yyyy"\)$/.exec(e))) return { k: 'd', n: m[1] }
  if ((m = /^"([^"]*)"$/.exec(e))) return { k: 'l', v: m[1] }
  if ((m = /^Format\(!(\w+), "#,##0\.00;\(#,##0\.00\)"\)$/.exec(e))) return { k: 'mp', n: m[1] }
  if ((m = /^Format\(!(\w+), "##,#0\.00"\)$/.exec(e))) return { k: 'm', n: m[1] }
  if ((m = /^Format\(Round\(!(\w+) \/ !(\w+), 2\), "#,##0\.00"\)$/.exec(e))) return { k: 'avg', n: m[1], v: m[2] }
  if (/^!Year & "\/" & Month$/.test(e)) return { k: 'ym' }
  if ((m = /^IIf\(Len\(!(\w+)\), !(\w+), ""\)$/.exec(e)) && m[1] === m[2]) return { k: 'f', n: m[1] }
  if ((m = /^!(\w+) & "([^"]*)" & !(\w+)$/.exec(e))) return { k: 'cat', n: m[1] + '|' + m[3], v: m[2] }
  if ((m = /^Round\(\(!(\w+) \/ !(\w+)\) \* 100, 0\)$/.exec(e))) return { k: 'pct', n: m[1], v: m[2] }
  return null
}
const out = dispatch.map((d) => {
  const b = body[d.routine] ?? []
  const cases = [...new Set(b.map((l) => /strAppendCase = "(\d+)"/.exec(l)?.[1]).filter(Boolean).map(Number))]
  const calls = b.map((l) => /^Call (\w+)/.exec(l.trim())?.[1]).filter(Boolean)
  let first = null; const cols = {}; const firstAlts = []; const alts = {}
  for (const l of b) {
    const f = /ListItems\.Add\(, , (.*)\)$/.exec(l)
    if (f) { if (first === null) first = f[1]; firstAlts.push(f[1]) }
    const s = /\.SubItems\((\d+)\) = (.*)$/.exec(l)
    if (s) { if (!(s[1] in cols)) cols[s[1]] = s[2]; (alts[s[1]] ??= []).push(s[2]) }
  }
  // A field column whose other branch writes "Empty" (blank values) is shown as "Empty" when blank, as the desktop.
  const withEmpty = (c, list) => c && c.k === 'f' && (list ?? []).some((x) => x.trim() === '"Empty"') ? { k: 'fe', n: c.n } : c
  const n = Math.max(0, ...Object.keys(cols).map(Number))
  const spec = [withEmpty(parse(first), firstAlts)]
  for (let i = 1; i <= n; i++) spec.push(cols[i] === undefined ? { k: 'l', v: '' } : withEmpty(parse(cols[i]), alts[i]))
  const special = cases.length !== 1 || calls.length > 0 || spec.some((c) => c === null)
  return { type: d.type, selection: d.selection, routine: d.routine, headers: headers(headerOf[d.type], d.selection), special,
    p: special ? { c: cases[0] ?? 0, special: true } : { c: cases[0], cols: spec }, s: null, helper: null,
    why: special ? { cases, calls: [...new Set(calls)], bad: [first, ...Object.values(cols)].filter((x) => parse(x) === null).slice(0, 4) } : undefined }
})
fs.writeFileSync('D:/CIS/Cis/src/Cis.Application/Reports/ClmTracking.catalogue.json', JSON.stringify(out, null, 1))
console.log(out.length, 'reports;', out.filter((r) => !r.special).length, 'generic; no headers:', out.filter((r) => !r.headers.length).map((r) => `${r.type}|${r.selection}`).join(', '))
for (const r of out.filter((r) => r.special)) console.log(' ', r.type, r.selection, r.routine, JSON.stringify(r.why).slice(0, 160))
