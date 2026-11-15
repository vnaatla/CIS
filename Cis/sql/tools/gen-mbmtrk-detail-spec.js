// Adds the detail listings (one row per member) to MbmTracking.catalogue.json: cases [P, S, B principal, B supplementary] and per loop the column specs.
// Kinds: f field, d date, m money, cv cover id (literal fallback), nm name + " (Suspended / Cancelled / Deleted)", l literal.
const fs = require('fs')
const FORM = process.env.TRK_FORM || 'FrmNewMBMTracking'
const SUF = FORM === 'FrmNewMBMTracking' ? '' : '-' + FORM
const CAT = process.env.TRK_CAT || 'MbmTracking.catalogue.json'
const detail = JSON.parse(fs.readFileSync('C:/Users/Naresh/AppData/Local/Temp/mbmtrk-detail' + SUF + '.json', 'utf8'))
const file = 'D:/CIS/Cis/src/Cis.Application/Reports/' + CAT
const cat = JSON.parse(fs.readFileSync(file, 'utf8'))
const one = (alts) => {
  const xs = alts.map((a) => a.trim().replace(/\b\w+!/g, '!'))
  if (xs.includes('strName')) return { k: 'nm' }
  if (xs.some((x) => /CoverID/i.test(x))) return { k: 'cv', n: /!(\w*CoverID)/i.exec(xs.find((x) => /!\w*CoverID/i.test(x)))[1], v: (/^"([^"]*)"$/.exec(xs.find((x) => /^"/.test(x)) ?? '"00"') ?? [])[1] ?? '00' }
  let m
  for (const x of xs) {
    if ((m = /^Format\(!(\w+), "dd\/mm\/yyyy"\)$/.exec(x))) return { k: 'd', n: m[1] }
    if ((m = /^Format\(!(\w+), "#,##0\.00"\)$/.exec(x))) return { k: 'm', n: m[1] }
    if ((m = /^(?:Trim\()?!(\w+)\)?$/.exec(x))) return { k: 'f', n: m[1] }
  }
  if ((m = /^"([^"]*)"$/.exec(xs[0]))) return { k: 'l', v: m[1] }
  return { k: 'l', v: '' }
}
const cols = (loop) => {
  const n = Math.max(0, ...Object.keys(loop.cols).map(Number))
  const out = [{ k: 'f', n: /!(\w+)/.exec(loop.first.replace(/\b\w+!/g, '!'))?.[1] ?? 'MBMNumber' }]
  for (let i = 1; i <= n; i++) out.push(loop.cols[i] ? one(loop.cols[i]) : { k: 'l', v: '' })
  return out
}
for (const e of cat) {
  const d = detail[e.routine]
  if (!d) continue
  const [p, s, bp, bs] = d.cases.map(Number)
  e.special = false
  e.detail = { p, s, bp, bs, supp: cols(d.loops[0]), prin: cols(d.loops[1]) }
}
fs.writeFileSync(file, JSON.stringify(cat, null, 1))
console.log(cat.filter((e) => e.detail).length, 'detail listings;', cat.filter((e) => e.special).map((e) => e.routine).join(' '))
console.log(JSON.stringify(cat.find((e) => e.routine === 'GeneralListing4').detail.prin))
