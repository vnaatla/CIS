// Browser smoke test of the running API + UI (see e2e/README.md). Usage:
//   $env:E2E_PASSWORD='...'; npm run e2e
// Writes what a real login writes on the local databases: USR.CIScnt +1 / -1, a Tbl_USRTrack row, a LoginCIS.txt line, tblActionLog rows.
import puppeteer from 'puppeteer-core'
import { existsSync } from 'node:fs'

const base = process.env.E2E_URL ?? 'http://localhost:4200'
const user = process.env.E2E_USER ?? 'DEVADMIN'
const password = process.env.E2E_PASSWORD
const browserPath =
  process.env.E2E_BROWSER ??
  [
    'C:/Program Files/Google/Chrome/Application/chrome.exe',
    'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
    'C:/Program Files/Microsoft/Edge/Application/msedge.exe',
    '/usr/bin/google-chrome',
    '/usr/bin/chromium',
  ].find(existsSync)

if (!password) {
  console.error('E2E_PASSWORD is not set (local dev password of E2E_USER, see AGENTS.md).')
  process.exit(2)
}
if (!browserPath) {
  console.error('No Chrome / Edge found. Set E2E_BROWSER to the browser executable.')
  process.exit(2)
}

let failed = 0
const check = (ok, text) => {
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${text}`)
  if (!ok) failed++
}

const browser = await puppeteer.launch({ executablePath: browserPath, headless: true, args: ['--no-first-run'] })
try {
  const page = await browser.newPage()
  await page.setViewport({ width: 1400, height: 900 })

  // 1. Not signed in: any page goes to the login.
  await page.goto(`${base}/membership`, { waitUntil: 'networkidle0' })
  check(page.url().endsWith('/login'), 'signed-out users are sent to /login')

  // 2. Wrong password: the desktop's message (only once, the account is suspended after three).
  const [userBox, passwordBox] = await page.$$('form.card input')
  await userBox.type(user)
  await passwordBox.type('wrong-password')
  await page.click('form.card button[type=submit]')
  await page.waitForSelector('form.card .error')
  check((await page.$eval('form.card .error', (e) => e.textContent)).includes('Invalid Password!! Please enter again..'), 'wrong password shows the desktop message')

  // 3. Login.
  await passwordBox.type(password)
  await page.click('form.card button[type=submit]')
  await page.waitForSelector('header.topbar', { timeout: 15000 })
  check((await page.$eval('header.topbar', (e) => e.textContent)).includes(user), 'topbar shows the signed-in user')
  check((await page.$$('nav.rail a')).length === 12, 'rail lists Home + the 11 desktop modules')
  check((await page.$$('.home-modules a')).length === 11, 'home page has one card per module')

  // 4. Module overview + sidebar.
  await page.click('nav.rail a[aria-label="Membership"]')
  await page.waitForSelector('.module-items table')
  check((await page.$$eval('.module-items tbody tr', (rows) => rows.length)) === 8, 'Membership lists its 8 desktop screens')
  check((await page.$eval('nav.sidebar', (e) => e.textContent)).includes('MBM Adjustment'), 'sidebar shows the Membership items')
  await page.goto(`${base}/maintenance/no-such-screen`, { waitUntil: 'networkidle0' })
  check(page.url().endsWith('/maintenance'), 'unknown screens inside a module show the module overview')

  // 5. Phase 1: Maintenance > Check Item (search all rows) and Case (five tabs).
  await page.goto(`${base}/maintenance/check-item`, { waitUntil: 'networkidle0' })
  await page.waitForSelector('app-code-table-panel fieldset input')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-data-grid tbody tr:not(.empty-row)').length > 0, { timeout: 30000 })
  check((await page.$$eval('app-data-grid tbody tr:not(.empty-row)', (r) => r.length)) > 0, 'Check Item search lists the CheckItems rows')
  await page.goto(`${base}/maintenance/case`, { waitUntil: 'networkidle0' })
  await page.waitForSelector('.tabs button')
  check((await page.$$('.tabs button')).length === 5, 'Case shows the desktop\'s five tabs')

  // 6. LG Maintenance > GL Template Binding: the drop-downs load, a group company fills its plans, Search lists its bindings.
  await page.goto(`${base}/maintenance/lg`, { waitUntil: 'domcontentloaded' })
  await page.waitForFunction(() => document.querySelectorAll('app-lg-binding-panel select[name=groupCode] option').length > 1, { timeout: 60000 })
  const counts = await page.$$eval('app-lg-binding-panel select', (s) => s.map((e) => `${e.name}:${e.options.length - 1}`))
  check(['groupCode', 'glType', 'template', 'lgFormat'].every((n) => !counts.includes(`${n}:0`)), `LG drop-downs load (${counts.join(' ')})`)
  const group = await page.evaluate(async () => {
    const t = JSON.parse(sessionStorage.getItem('cis.session') ?? localStorage.getItem('cis.session') ?? '{}').token
    const rows = await (await fetch('/api/maintenance/lg-binding?groupName=%25', { headers: { Authorization: `Bearer ${t}` } })).json()
    const names = Array.isArray(rows) ? rows.map((r) => (r.groupCompany ?? '').trim()) : []
    const opt = [...document.querySelectorAll('app-lg-binding-panel select[name=groupCode] option')].find((o) => o.value && names.includes((o.textContent.split('-')[1] ?? '').trim()))
    return opt?.value ?? ''
  })
  if (group) {
    await page.select('app-lg-binding-panel select[name=groupCode]', group)
    await page.waitForFunction(() => document.querySelectorAll('app-lg-binding-panel select[name=planCode] option').length > 1, { timeout: 60000 })
    await page.click('app-lg-binding-panel app-action-bar button.primary')
    await page.waitForFunction(() => document.querySelectorAll('app-lg-binding-panel tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  }
  check(!!group && (await page.$$eval('app-lg-binding-panel tbody tr:not(.empty-row)', (r) => r.length)) > 0, `LG group company ${group} fills its plans and Search lists its bindings`)

  // 7. Benefit > Clinic: every category is a grid row; Gold + 01/01/2002 (saved in the local data) fills RM Payable Fees.
  await page.goto(`${base}/maintenance/benefit/clinic`, { waitUntil: 'domcontentloaded' })
  await page.waitForFunction(() => document.querySelectorAll('app-benefit-clinic-page tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.select('app-benefit-clinic-page select[name=cardType]', 'Gold')
  await page.type('app-benefit-clinic-page app-date-input input:not([type=date])', '01012002')
  await page.click('app-benefit-clinic-page app-action-bar button.primary')
  await page.waitForFunction(() => [...document.querySelectorAll('app-benefit-clinic-page tbody input')].some((i) => i.value !== ''), { timeout: 60000 })
  check(true, 'Benefit Clinic lists the categories and Search fills the saved Gold fees')

  // 8. Benefit > Member: Benefit Entry search (A001 / H / N / 01/01/2004 has 29 saved rows locally) and Benefit Listing search by plan.
  await page.goto(`${base}/maintenance/benefit/member`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('.tabs button')
  const memberTabs = await page.$$('.tabs button')
  check(memberTabs.length === 3, 'Benefit Member shows Category, Benefit Entry and Benefit Listing')
  await memberTabs[1].click()
  await page.waitForFunction(() => document.querySelectorAll('app-benefit-entry-panel tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.select('app-benefit-entry-panel select[name=health]', 'N')
  await page.select('app-benefit-entry-panel select[name=insured]', 'H')
  await page.type('app-benefit-entry-panel input[name=plan]', 'A001')
  await page.type('app-benefit-entry-panel app-date-input input:not([type=date])', '01012004')
  await page.click('app-benefit-entry-panel app-action-bar button.primary')
  await page.waitForFunction(() => [...document.querySelectorAll('app-benefit-entry-panel tbody input')].some((i) => i.value !== ''), { timeout: 60000 })
  check(true, 'Benefit Entry search fills the saved A001 grid')
  await (await page.$$('.tabs button'))[2].click()
  await page.waitForSelector('app-benefit-listing-panel input[name=planCode]')
  await page.type('app-benefit-listing-panel input[name=planCode]', 'A001')
  await page.click('app-benefit-listing-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-benefit-listing-panel tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  const firstListed = await page.$eval('app-benefit-listing-panel tbody tr td:nth-child(5)', (td) => td.textContent.trim())
  check(firstListed === 'A001', `Benefit Listing search by plan lists A001 rows (${firstListed})`)

  // 9. Logout.
  await page.click('header.topbar button.link:last-of-type')
  await page.waitForSelector('form.card')
  check(page.url().endsWith('/login'), 'logout returns to the login page')
} finally {
  await browser.close()
}

console.log(failed === 0 ? '\nAll checks passed.' : `\n${failed} check(s) failed.`)
process.exit(failed === 0 ? 0 : 1)
