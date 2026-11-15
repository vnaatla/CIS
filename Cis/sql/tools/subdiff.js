// Per-routine differences between two VB form extracts (case / connection / whitespace normalised).
const fs = require('fs')
const [a, b] = process.argv.slice(2)
const norm = (l) => l.trim().toLowerCase().replace(/medixneptuneconn|medixmasterconn|mediconncisws|medineptuneconncisws|mediconncisdb|medineptuneconncisdb|mediconncispayment|medineptuneconncispayment/g, 'conn').replace(/adopenkeyset, adlockoptimistic|adopenforwardonly, adlockreadonly/g, 'mode').replace(/\s+/g, ' ')
const subs = (f) => {
  const out = {}; let cur = '(top)'
  for (const l of fs.readFileSync(f, 'utf8').split(/\r?\n/)) {
    const m = /^(?:private|public) (?:sub|function) (\w+)/i.exec(l.trim())
    if (m) { cur = m[1].toLowerCase(); out[cur] = []; continue }
    ;(out[cur] ??= []).push(norm(l))
  }
  return out
}
const A = subs(a), B = subs(b)
const ignore = /^(lstmbminquirylist|listview2)\.(left|height|top|width)|^lblcurrent = ""$|^(if|do until|do while not) \w+\.(recordcount|eof)( then)?$|^call actionlogupdate|^$/
const clean = (x) => (x ?? []).filter((l) => !ignore.test(l))
const onlyA = Object.keys(A).filter((k) => !(k in B)), onlyB = Object.keys(B).filter((k) => !(k in A))
console.log('only in first:', onlyA.join(' ') || '-')
console.log('only in second:', onlyB.join(' ') || '-')
for (const k of Object.keys(A).filter((k) => k in B)) {
  const x = clean(A[k]), y = clean(B[k])
  if (x.join('\n') === y.join('\n')) continue
  const sx = new Set(x), sy = new Set(y)
  const minus = x.filter((l) => !sy.has(l)), plus = y.filter((l) => !sx.has(l))
  if (!minus.length && !plus.length) { console.log(`~ ${k}: same lines, different order`); continue }
  console.log(`~ ${k}: -${minus.length} +${plus.length} | - ${minus.slice(0, 2).join(' ¦ ').slice(0, 160)} | + ${plus.slice(0, 2).join(' ¦ ').slice(0, 160)}`)
}
