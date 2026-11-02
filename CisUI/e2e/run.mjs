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

  // 9. Benefit > Dental: four tabs; Benefit Entry search (01 / I / S / 24/09/2003 has 3 saved rows locally).
  await page.goto(`${base}/maintenance/benefit/dental`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('.tabs button')
  const dentalTabs = await page.$$('.tabs button')
  check(dentalTabs.length === 4, 'Benefit Dental shows the desktop\'s four tabs')
  await dentalTabs[3].click()
  await page.waitForFunction(() => document.querySelectorAll('app-benefit-entry-panel tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.select('app-benefit-entry-panel select[name=health]', 'S')
  await page.select('app-benefit-entry-panel select[name=insured]', 'I')
  await page.type('app-benefit-entry-panel input[name=plan]', '01')
  await page.type('app-benefit-entry-panel app-date-input input:not([type=date])', '24092003')
  await page.click('app-benefit-entry-panel app-action-bar button.primary')
  await page.waitForFunction(() => [...document.querySelectorAll('app-benefit-entry-panel tbody input')].some((i) => i.value !== ''), { timeout: 60000 })
  check(true, 'Dental Benefit Entry search fills the saved grid')

  // 10. Entity: five tabs; Clinic Group search by name lists GRPClinic rows and a row opens read-only.
  await page.goto(`${base}/maintenance/entity`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('app-entity-group-panel input[name=name]')
  check((await page.$$('.tabs button')).length === 5, 'Entity shows the desktop\'s five tabs')
  await page.type('app-entity-group-panel input[name=name]', 'KLINIK')
  await page.click('app-entity-group-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-entity-group-panel tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.click('app-entity-group-panel tbody tr:not(.empty-row)')
  await page.waitForFunction(() => { const i = document.querySelector('app-entity-group-panel input[name=name]'); return i?.disabled && i.value.includes('KLINIK') }, { timeout: 60000 })
  check(true, 'Clinic Group search lists groups and a row opens read-only')
  await (await page.$$('.tabs button'))[2].click()
  await page.waitForSelector('app-entity-clinic-panel input[name=name]')
  await page.type('app-entity-clinic-panel input[name=name]', 'KLINIK')
  await page.click('app-entity-clinic-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-entity-clinic-panel tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.click('app-entity-clinic-panel tbody tr:not(.empty-row)')
  await page.waitForFunction(() => { const i = document.querySelector('app-entity-clinic-panel input[name=name]'); return i?.disabled && i.value.includes('KLINIK') }, { timeout: 60000 })
  check(true, 'Clinic Info search by name lists HISDB clinics and a row opens read-only')
  await (await page.$$('.tabs button'))[3].click()
  await page.waitForSelector('app-entity-grpcom-panel input[name=groupCompany]')
  await page.type('app-entity-grpcom-panel input[name=groupCompany]', 'B006')
  await page.click('app-entity-grpcom-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-entity-grpcom-panel tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  check(true, 'Group Company Clinic search by group company lists its clinics')
  await (await page.$$('.tabs button'))[4].click()
  await page.waitForSelector('app-entity-access-panel app-action-bar button.primary')
  await page.click('app-entity-access-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-entity-access-panel tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.click('app-entity-access-panel app-action-bar button:nth-of-type(2)')
  await page.click('app-entity-access-panel fieldset button')
  await page.waitForSelector('app-modal input[name=lookupName]')
  await page.type('app-modal input[name=lookupName]', 'KLINIK')
  await page.click('app-modal app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-modal tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.click('app-modal tbody tr:not(.empty-row)')
  await page.waitForFunction(() => document.querySelector('app-entity-access-panel input[name=clinicCode]')?.value !== '', { timeout: 30000 })
  check(true, 'Clinic Access lists the latest rows and the clinic look-up fills the code')

  // 11. Plan: eight tabs; HealthType lists HLT; Plan Code search by code lists A001 and a row opens read-only.
  await page.goto(`${base}/maintenance/plan`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('app-code-table-panel app-action-bar button.primary')
  check((await page.$$('.tabs button')).length === 8, 'Plan shows the desktop\'s eight tabs')
  await page.click('app-code-table-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-code-table-panel tbody tr:not(.empty-row)').length >= 2, { timeout: 60000 })
  check(true, 'HealthType search lists the health types')
  await (await page.$$('.tabs button'))[2].click()
  await page.waitForSelector('app-plan-code-panel input[name=plan0]')
  await page.type('app-plan-code-panel input[name=plan0]', 'A001')
  await page.click('app-plan-code-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-plan-code-panel app-data-grid tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.click('app-plan-code-panel app-data-grid tbody tr:not(.empty-row)')
  await page.waitForFunction(() => { const i = document.querySelector('app-plan-code-panel input[name=plan0]'); return i?.disabled && i.value === 'A001' }, { timeout: 60000 })
  check(true, 'Plan Code search lists plan A001 and a row opens read-only')
  await page.click('app-plan-code-panel app-action-bar button.ghost')
  await page.click('app-plan-code-panel app-action-bar button.primary')
  await page.waitForSelector('app-plan-code-panel .grid-pager', { timeout: 60000 })
  const firstPlan = () => page.$eval('app-plan-code-panel app-data-grid tbody tr td:nth-child(2)', (td) => td.textContent.trim())
  const before = await firstPlan()
  const pagerText = await page.$eval('app-plan-code-panel .grid-pager', (n) => n.textContent)
  check(/Rows 1–50 of \d{4}/.test(pagerText), 'Plan Code: an unfiltered search pages on the server (Rows 1–50 of the full total)')
  await page.evaluate(() => [...document.querySelectorAll('app-plan-code-panel .grid-pager .page-btn')].find((b) => b.textContent.trim() === 'Next').click())
  await page.waitForFunction((b) => document.querySelector('app-plan-code-panel app-data-grid tbody tr td:nth-child(2)')?.textContent.trim() !== b, { timeout: 60000 }, before)
  check((await page.$eval('app-plan-code-panel .grid-pager', (n) => n.textContent)).includes('Rows 51–100'), 'Plan Code: Next loads page 2 from the API')
  await (await page.$$('.tabs button'))[3].click()
  await page.waitForSelector('app-plan-premium-panel select[name=insCode] option[value=H]')
  await page.select('app-plan-premium-panel select[name=insCode]', 'H')
  await page.click('app-plan-premium-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-plan-premium-panel app-data-grid tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.click('app-plan-premium-panel app-data-grid tbody tr:not(.empty-row)')
  await page.waitForFunction(() => document.querySelector('app-plan-premium-panel input[name=plnCode]')?.disabled, { timeout: 60000 })
  check(true, 'Premium search by insured type lists PRM rows and a row opens its age bands')
  await (await page.$$('.tabs button'))[4].click()
  await page.waitForSelector('app-plan-annual-panel input[name=plnCode]')
  await page.type('app-plan-annual-panel input[name=plnCode]', 'A001')
  await page.click('app-plan-annual-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-plan-annual-panel app-data-grid tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.click('app-plan-annual-panel app-data-grid tbody tr:not(.empty-row)')
  await page.waitForFunction(() => document.querySelector('app-plan-annual-panel input[name=annLimit]')?.disabled, { timeout: 60000 })
  check(true, 'Annual Limit search by plan lists its limits and a row opens read-only')
  await (await page.$$('.tabs button'))[5].click()
  await page.waitForSelector('app-plan-visit-panel input[name=visitLimit]')
  await page.click('app-plan-visit-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-plan-visit-panel app-data-grid tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.click('app-plan-visit-panel app-data-grid tbody tr:not(.empty-row)')
  await page.waitForFunction(() => document.querySelector('app-plan-visit-panel input[name=visitLimit]')?.disabled, { timeout: 60000 })
  check(true, 'Visit Limit search lists the limits and a row opens read-only')
  await (await page.$$('.tabs button'))[6].click()
  await page.waitForSelector('app-plan-mco-panel input[name=plnCode]')
  await page.type('app-plan-mco-panel input[name=plnCode]', 'A001')
  await page.click('app-plan-mco-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-plan-mco-panel app-data-grid tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  await page.click('app-plan-mco-panel app-data-grid tbody tr:not(.empty-row)')
  await page.waitForFunction(() => document.querySelector('app-plan-mco-panel input[name=plnCode]')?.disabled, { timeout: 60000 })
  check(true, 'MCO Fee search by plan lists its fees and a row opens read-only')
  await (await page.$$('.tabs button'))[7].click()
  await page.waitForFunction(() => document.querySelector('app-code-table-panel legend')?.textContent.includes('Co-Payment'), { timeout: 60000 })
  await page.click('app-code-table-panel app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-code-table-panel tbody tr:not(.empty-row)').length > 0, { timeout: 60000 })
  check(true, 'Co-Payment (MBM) search lists the plan co-payments')

  // 12. Membership > MBM Topup Adj: Go shows the member and the GP balance; the member search pop-up finds by name and loads the pick.
  await page.goto(`${base}/membership/topup`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=member]')
  await page.type('input[name=member]', 'DCHC9086608')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelector('input[name=name]')?.value.length > 0, { timeout: 60000 })
  const balance = await page.$eval('input[name=balance]', (i) => i.value)
  check(balance === '3,129.00', `Topup Go loads member DCHC9086608 with GP balance ${balance}`)
  await page.evaluate(() => [...document.querySelectorAll('app-action-bar button')].find((b) => b.textContent.trim() === 'Search').click())
  await page.waitForSelector('app-member-search-modal input[name=memberSearchText]')
  await page.type('app-member-search-modal input[name=memberSearchText]', 'NUR RABIZAH BINTI ADENI')
  await page.click('app-member-search-modal app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelectorAll('app-member-search-modal tbody tr:not(.empty-row)').length > 0, { timeout: 90000 })
  await page.click('app-member-search-modal tbody tr:not(.empty-row)')
  await page.waitForFunction(() => !document.querySelector('app-member-search-modal') && document.querySelector('input[name=name]')?.value.length > 0, { timeout: 60000 })
  check((await page.$eval('input[name=member]', (i) => i.value)).startsWith('DCHC9086608-'), 'Member search by name lists the member and a click loads it into Topup')

  // 13. Invoice > Case Transfer: Go shows the case's member; Search finds the transfer member and enables Transfer (not pressed).
  await page.goto(`${base}/invoice/case-transfer`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=casNumber]')
  await page.type('input[name=casNumber]', 'K001596118')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelector('input[name=currentMember]')?.value === 'DCGB3518731', { timeout: 60000 })
  await page.type('input[name=mbmNumber]', 'DCHC9086608')
  await page.type('input[name=covId]', '00')
  await page.evaluate(() => [...document.querySelectorAll('app-action-bar button')].find((b) => b.textContent.trim() === 'Search').click())
  await page.waitForFunction(() => document.querySelector('input[name=transferPatient]')?.value === 'NUR RABIZAH BINTI ADENI', { timeout: 60000 })
  const transferEnabled = await page.evaluate(() => ![...document.querySelectorAll('app-action-bar button')].find((b) => b.textContent.trim() === 'Transfer').disabled)
  check(transferEnabled, 'Case Transfer: Go shows case K001596118 (DCGB3518731) and Search enables Transfer for DCHC9086608-00')

  // 14. Invoice > General Exclusions (Enquiry): a company's flags and remarks.
  await page.goto(`${base}/invoice/general-exclusions/enquiry`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=groupCompany]')
  await page.type('input[name=groupCompany]', '1MILLENNIUM INTERNATIONAL SDN BHD')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelector('textarea[name=remarks]')?.value.startsWith('PAYOR : ST'), { timeout: 60000 })
  const dental = await page.$eval('input[name=Dental]', (i) => i.checked)
  check(dental, 'General Exclusions Enquiry shows 1MILLENNIUM INTERNATIONAL SDN BHD (Dental excluded, remarks)')

  // 15. Membership > MBM Adjustment: Show loads the member; the Supplementary tab lists its three covers.
  await page.goto(`${base}/membership/adjustment`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=member]')
  await page.type('input[name=member]', 'DCHC9086608')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelector('input[name=name]')?.value === 'NUR RABIZAH BINTI ADENI', { timeout: 90000 })
  await page.evaluate(() => [...document.querySelectorAll('.tabs button')].find((b) => b.textContent.trim() === 'Supplementary').click())
  await page.waitForFunction(() => document.querySelectorAll('app-data-grid tbody tr:not(.empty-row)').length === 3, { timeout: 60000 })
  check(true, 'MBM Adjustment shows DCHC9086608 and its 3 covered persons')

  // 15b. Membership > MBM Enquiry: Show computes the limits (7,500 - 4,371 top-ups = 3,129 balance); the Case tab lists the cases of a member with claims.
  await page.goto(`${base}/membership/enquiry`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=member]')
  await page.type('input[name=member]', 'DCHC9086608')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelector('input[name="Balance Limit"]')?.value === '3,129.00', { timeout: 90000 })
  const gpRow = await page.$eval('.limits-grid tbody tr:nth-child(3)', (r) => [...r.cells].map((c) => c.textContent.trim()).join(' '))
  check(gpRow.startsWith('Balance 3,129.00'), `MBM Enquiry shows DCHC9086608's limits grid (${gpRow})`)
  await page.$eval('input[name=member]', (i) => { i.value = '' })
  await page.type('input[name=member]', 'DCHC4245383')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelector('input[name="Accum. Amt"]')?.value === '1,985.000', { timeout: 90000 })
  await page.evaluate(() => [...document.querySelectorAll('.tabs button')].find((b) => b.textContent.trim() === 'Case').click())
  await page.waitForFunction(() => document.querySelector('[aria-label="Specialist totals"]')?.textContent.includes('Total Visits: 10'), { timeout: 60000 })
  check(true, 'MBM Enquiry: DCHC4245383 utilised 1,985.000 and 10 specialist cases')
  await page.$eval('input[name=member]', (i) => { i.value = '' })
  await page.type('input[name=member]', 'DCHC8922156')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelector('input[name="Accum. Amt"]')?.value === '908.000', { timeout: 90000 })
  await page.evaluate(() => [...document.querySelectorAll('.tabs button')].find((b) => b.textContent.trim() === 'Supplementary').click())
  await page.waitForSelector('app-data-grid tbody tr:not(.empty-row)')
  await page.click('app-data-grid tbody tr:not(.empty-row)')
  await page.waitForSelector('.cover-limits')
  const covRow = await page.$eval('.cover-limits tbody tr:nth-child(3)', (r) => [...r.cells].map((c) => c.textContent.trim()).join(' '))
  check(covRow.startsWith('Balance 1,098.48'), `MBM Enquiry: a click on cover 01 of DCHC8922156 shows its shared limit (${covRow})`)
  await page.$eval('input[name=member]', (i) => { i.value = '' })
  await page.type('input[name=member]', 'UBHA0000156*05')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => [...document.querySelectorAll('.tabs button')].some((b) => b.textContent.trim() === 'FlexCare'), { timeout: 90000 })
  await page.evaluate(() => [...document.querySelectorAll('.tabs button')].find((b) => b.textContent.trim() === 'FlexCare').click())
  await page.waitForFunction(() => document.querySelector('[aria-label="FlexCare totals"]')?.textContent.includes('4,711.42'), { timeout: 90000 })
  await page.evaluate(() => [...document.querySelectorAll('.tabs button')].find((b) => b.textContent.trim() === 'Supplementary').click())
  await page.waitForSelector('app-data-grid tbody tr:not(.empty-row)')
  await page.click('app-data-grid tbody tr:not(.empty-row)')
  await page.evaluate(() => [...document.querySelectorAll('.tabs button')].find((b) => b.textContent.trim() === 'Family Benefits').click())
  await page.waitForFunction(() => document.querySelector('[aria-label="Family benefits"]')?.textContent.includes('4,722.00'), { timeout: 90000 })
  check(true, 'MBM Enquiry: UBHA0000156*05 FlexCare balance 4,711.42 and family balance 4,722.00 (cover 01)')

  // 16. Invoice > Batch Transfer: both batches show their totals and the final totals (nothing is transferred).
  await page.goto(`${base}/invoice/batch-transfer`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=batch1]')
  await page.type('input[name=batch1]', 'opy9978')
  await page.keyboard.press('Enter')
  await page.waitForFunction(() => document.querySelector('input[name=clinic1]')?.value === 'KLINIK DAN SURGERI HILLTOP', { timeout: 60000 })
  await page.type('input[name=batch2]', 'OPY9987')
  await page.keyboard.press('Enter')
  await page.waitForFunction(() => document.querySelector('input[name=finalAmt]')?.value === '283.00', { timeout: 60000 })
  const finalCases = await page.$eval('input[name=finalCases]', (i) => i.value)
  check(finalCases === '6', `Batch Transfer shows OPY9978 (4 cases, 184.00) + OPY9987 (2 cases, 99.00) = ${finalCases} cases, 283.00`)

  // 17. Invoice > Invoice Update: Search by invoice number lists its cases (nothing is updated).
  await page.goto(`${base}/invoice/update`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=invoiceNo]')
  await page.type('input[name=invoiceNo]', 'I00496-ERMA-M813')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => [...document.querySelectorAll('app-data-grid tbody td')].some((td) => td.textContent.trim() === 'EA12582360'), { timeout: 120000 })
  check(true, 'Invoice Update finds the cases of invoice I00496-ERMA-M813 (EA12582360)')

  // 18. Invoice > OP Batch/Invoice: clinic T301, August 2021 lists its 3 open cases (nothing is updated).
  await page.goto(`${base}/invoice/op-batch`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=clinic]')
  await page.type('input[name=clinic]', 'T301 - X')
  await page.type('input[aria-label="Visit Date From"]', '01/08/2021')
  await page.type('input[aria-label="Visit Date To"]', '31/08/2021')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => document.querySelector('[aria-label="Totals"]')?.textContent.includes('Records: 3'), { timeout: 120000 })
  check(true, 'OP Batch/Invoice lists the 3 open cases of clinic T301 in August 2021')

  // 19. Invoice > MMA/CPT: search BRAIN / A01, a row opens its details.
  await page.goto(`${base}/invoice/mma-cpt`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=mainBp]')
  await page.type('input[name=mainBp]', 'brain')
  await page.type('input[name=opcs]', 'A01')
  await page.click('app-action-bar button.primary')
  await page.waitForFunction(() => [...document.querySelectorAll('app-data-grid tbody td')].some((td) => td.textContent.trim() === 'A0110'), { timeout: 60000 })
  await page.evaluate(() => [...document.querySelectorAll('app-data-grid tbody tr')].find((tr) => tr.textContent.includes('A0110')).click())
  await page.waitForFunction(() => document.querySelector('input[name=dSurgFee]')?.value === '7030', { timeout: 30000 })
  check(true, 'MMA/CPT finds A0110 (BRAIN) and its details show the surgeon fee 7030')
  await page.click('button[aria-label="Close"]')
  await page.waitForFunction(() => !document.querySelector('input[name=dSurgFee]'), { timeout: 10000 })

  // 16. Logout.
  await page.click('header.topbar button.link:last-of-type')
  await page.waitForSelector('form.card')
  check(page.url().endsWith('/login'), 'logout returns to the login page')
} finally {
  await browser.close()
}

console.log(failed === 0 ? '\nAll checks passed.' : `\n${failed} check(s) failed.`)
process.exit(failed === 0 ? 0 : 1)
