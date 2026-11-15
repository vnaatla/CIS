// Summarises FrmNewMBMTracking's listing routines: cases called, first column, SubItems columns, per branch (P/B vs S).
const fs = require('fs')
const lines = fs.readFileSync('C:/Users/Naresh/AppData/Local/Temp/p10/FrmNewMBMTracking.txt', 'utf8').split(/\r?\n/)
const subs = []
let cur = null
for (const l of lines) {
  const m = /^(?:Private|Public) (?:Sub|Function) (\w+)/.exec(l)
  if (m) { cur = { name: m[1], body: [] }; subs.push(cur); continue }
  if (cur) cur.body.push(l)
}
const out = {}
for (const s of subs) {
  const cases = [...s.body.join('\n').matchAll(/strAppendCase = "(\d+)"/g)].map((m) => Number(m[1]))
  if (!cases.length) continue
  const blocks = []
  let b = null
  for (const l of s.body) {
    const c = /strAppendCase = "(\d+)"/.exec(l)
    if (c) { b = { c: Number(c[1]), first: null, cols: {}, empty: [] }; blocks.push(b); continue }
    if (!b) continue
    const f = /ListItems\.Add\(, , (.*)\)$/.exec(l)
    if (f && !b.first) b.first = f[1]
    if (f && f[1].trim() === '"Empty"' && !b.empty.includes(0)) b.empty.push(0)
    const si = /Record\.SubItems\((\d+)\) = (.*)$/.exec(l)
    if (si && si[2].trim() === '"Empty"' && !b.empty.includes(Number(si[1]))) b.empty.push(Number(si[1]))
    if (si && !(si[1] in b.cols)) b.cols[si[1]] = si[2].replace(/rst2!/g, '!').slice(0, 60)
  }
  out[s.name] = blocks.map((x) => ({ c: x.c, first: (x.first || '').replace(/rst2!/g, '!'), cols: x.cols, empty: x.empty }))
}
fs.writeFileSync('C:/Users/Naresh/AppData/Local/Temp/mbmtrk-listings.json', JSON.stringify(out, null, 1))
console.log(Object.keys(out).length, 'routines with cases')
for (const [k, v] of Object.entries(out).slice(0, 6)) console.log(k, JSON.stringify(v).slice(0, 300))
