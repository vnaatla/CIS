// Joins FrmNewMBMTracking's dispatch (report type + selection -> listing routine), the routines' cases / columns and the headers into one catalogue JSON.
const fs = require('fs')
const T = 'C:/Users/Naresh/AppData/Local/Temp/'
const lines = fs.readFileSync(T + 'p10/FrmNewMBMTracking.txt', 'utf8').split(/\r?\n/)
const listings = JSON.parse(fs.readFileSync(T + 'mbmtrk-listings.json', 'utf8'))
const headers = JSON.parse(fs.readFileSync(T + 'mbmtrk-headers.json', 'utf8'))
const dispatch = {}
let cond = null
for (const l of lines) {
  const c = /CboReportType\.ListIndex = (\d+)(?: And CboRptSelection\.Text = "([^"]*)")?\s+Then/.exec(l)
  if (c && /^(Else)?If /.test(l)) { cond = `${c[1]}|${c[2] ?? ''}`; continue }
  const call = /^Call (\w+)(?:\(strAppend\))?$/.exec(l.trim())
  if (call && cond && !/^(SearchRecord|SearchForm|ListViewHeader\d+|GeneralListing\w*|LoadPayor|LoadInsuredType)$/.test('') ) {
    if (!dispatch[cond] && listings[call[1]] !== undefined || ['GeneralListing', 'GeneralListing2', 'GeneralListing4', 'PrintLabel', 'CardSentListing', 'AgentGeneralListing',
      'MCOInvoice', 'CheckDate', 'ICListing', 'BordxDateListing'].includes(call[1])) dispatch[cond] ??= call[1]
    cond = null
  }
}
// Helper routines each listing calls (CountSupp...), by name.
const body = {}
let cur = null
for (const l of lines) {
  const m = /^(?:Private|Public) (?:Sub|Function) (\w+)/.exec(l)
  if (m) { cur = m[1]; body[cur] = []; continue }
  if (cur) body[cur].push(l)
}
const catalogue = Object.entries(dispatch).filter(([k]) => Number(k.split('|')[0]) <= 17).map(([key, routine]) => {
  const [type, selection] = key.split('|')
  const calls = [...new Set((body[routine] || []).map((l) => /^Call (\w+)\(strAppend\)/.exec(l.trim())?.[1]).filter(Boolean))]
  return { type: Number(type), selection, routine, blocks: listings[routine] ?? [], helpers: calls.map((c) => ({ name: c, blocks: listings[c] ?? [] })),
    headers: headers[key] ?? headers[`${type}|`] ?? [] }
})
fs.writeFileSync(T + 'mbmtrk-catalogue.json', JSON.stringify(catalogue, null, 1))
console.log(catalogue.length, 'reports;', 'no blocks:', catalogue.filter((c) => !c.blocks.length).map((c) => `${c.type}|${c.selection}=${c.routine}`).join(', '))
console.log('no headers:', catalogue.filter((c) => !c.headers.length).map((c) => `${c.type}|${c.selection}`).join(', '))
