// Run from CisUI: node scripts/test-guide-coverage.cjs. Lists the migrated menu items (menu.ts migrated: true) that no test guide mentions (by form name, route or label).
const fs = require('fs')
const path = require('path')
const menu = fs.readFileSync(path.join(__dirname, '../src/app/layout/menu.ts'), 'utf8')
const dir = path.join(__dirname, '../../Cis/docs/testing/')
const guides = Object.fromEntries(fs.readdirSync(dir).map((f) => [f, fs.readFileSync(dir + f, 'utf8')]))
const all = Object.values(guides).join('\n').toLowerCase()
const items = [...menu.matchAll(/\{ label: '([^']*)', path: '([^']*)'[^}]*?\}/g)].filter((m) => /migrated: true/.test(m[0]))
const missing = []
for (const m of items) {
  const forms = (/form: '([^']*)'/.exec(m[0])?.[1] ?? '').split(/,\s*/).filter(Boolean)
  const path = m[2]
  const hit = forms.some((f) => all.includes(f.toLowerCase())) || all.includes(path.toLowerCase()) || all.includes('`' + path.toLowerCase())
  if (!hit) missing.push({ label: m[1], path, forms: forms.join(', '), phase: /phase: (\d+)/.exec(m[0])?.[1] })
}
console.log('migrated', items.length, 'not in any test guide', missing.length)
for (const x of missing) console.log(`${x.phase}\t${x.label}\t${x.path}\t${x.forms}`)
