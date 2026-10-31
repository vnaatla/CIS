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

  // 6. Logout.
  await page.click('header.topbar button.link:last-of-type')
  await page.waitForSelector('form.card')
  check(page.url().endsWith('/login'), 'logout returns to the login page')
} finally {
  await browser.close()
}

console.log(failed === 0 ? '\nAll checks passed.' : `\n${failed} check(s) failed.`)
process.exit(failed === 0 ? 0 : 1)
