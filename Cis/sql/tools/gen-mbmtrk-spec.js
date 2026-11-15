// mbmtrk-catalogue.json -> Cis/src/Cis.Application/Reports/MbmTracking.catalogue.json: per report the principal / supplementary cases and parsed column specs.
// Column spec: { k: 'f', n: field } | { k: 'm', n: field } (money) | { k: 'l', v: literal } | { k: 'sc' } supp count | { k: 'sm' } supp MCO | { k: 'sp' } supp premium
// | { k: 'fl', n: field, v: suffix } (field + literal suffix). Reports whose columns use other accumulators are marked special (ported by hand).
const fs = require('fs')
const T = 'C:/Users/Naresh/AppData/Local/Temp/'
const cat = JSON.parse(fs.readFileSync(T + 'mbmtrk-catalogue.json', 'utf8'))
const parse = (e) => {
  e = (e || '').trim().replace(/\brst\d!/g, '!')
  let m
  if ((m = /^!(\w+)$/.exec(e)) || (m = /^Trim\(!(\w+)\)$/.exec(e))) return { k: 'f', n: m[1] }
  if ((m = /^Format\(!(\w+), "#,##0\.00"\)$/.exec(e))) return { k: 'm', n: m[1] }
  if ((m = /^"([^"]*)"$/.exec(e))) return { k: 'l', v: m[1] }
  if (e === 'TotalS') return { k: 'sc' }
  if (e === 'TotalSMCO') return { k: 'sm' }
  if (e === 'TotalSPrem') return { k: 'sp' }
  if ((m = /^Format\((\w+), "#,##0\.00"\)$/.exec(e)) && !/^Total/.test(m[1])) return { k: 'm', n: m[1] }
  if ((m = /^Format\(!(\w+), "dd\/mm\/yyyy"\)$/.exec(e))) return { k: 'd', n: m[1] }
  if (/^!Year & "\/" & Month$/.test(e)) return { k: 'ym' }
  if (/^Format\(TotalSMCO, "#,##0\.00"\)$/.test(e)) return { k: 'sm' }
  if (/^Format\(TotalSPrem, "#,##0\.00"\)$/.test(e)) return { k: 'sp' }
  if ((m = /^!(\w+) & "([^"]*)" & "([^"]*)"$/.exec(e))) return { k: 'fl', n: m[1], v: m[2] + m[3] }
  return null
}
const block = (b) => {
  if (!b) return null
  const n = Math.max(0, ...Object.keys(b.cols).map(Number))
  const cols = [parse(b.first)]
  for (let i = 1; i <= n; i++) cols.push(b.cols[i] === undefined ? { k: 'l', v: '' } : parse(b.cols[i]))
  // A field column whose other branch writes "Empty" shows "Empty" when blank, as the desktop.
  for (const i of b.empty ?? []) if (cols[i] && cols[i].k === 'f') cols[i] = { k: 'fe', n: cols[i].n }
  return cols.some((c) => c === null) || !b.first ? { c: b.c, special: true } : { c: b.c, cols }
}
const out = cat.map((r) => {
  const p = block(r.blocks[0]); const s = block(r.blocks[1])
  const helper = r.helpers.find((h) => h.blocks.length)
  const h = helper ? block(helper.blocks[0]) : null
  const special = !p || p.special || (s && s.special) || (h && h.special && helper.blocks[0].first) || r.blocks.length > 2
  return { type: r.type, selection: r.selection, routine: r.routine, headers: r.headers, special: !!special, p, s,
    helper: h ? { c: h.c, append: !!(helper.blocks[0].first), cols: h.cols ?? null } : null }
})
fs.writeFileSync('D:/CIS/Cis/src/Cis.Application/Reports/MbmTracking.catalogue.json', JSON.stringify(out, null, 1))
console.log(out.length, 'reports;', out.filter((r) => !r.special).length, 'generic;', 'special:', out.filter((r) => r.special).map((r) => r.routine).join(' '))
