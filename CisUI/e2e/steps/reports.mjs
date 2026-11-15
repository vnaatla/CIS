// Phase 10 (tracking and reports) browser steps. Run by run.mjs (full suite) and by only.mjs (E2E_STEPS=reports).
const click = (page, text) => page.evaluate((t) => [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === t).click(), text)

export async function reports(page, base, check) {
  // Tracking > Clinic: clinics named like BALA include B020.
  await page.goto(`${base}/tracking/clinic`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=name]')
  await page.type('input[name=name]', 'BALA')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Clinics"] tbody')?.textContent.includes('B020'), { timeout: 120000 })
  check(true, 'Clinic Tracking: clinics named like BALA include B020')

  // Tracking > Grp. Com. Clinic: group companies per clinic for clinic code B324.
  await page.goto(`${base}/tracking/grp-company-clinic`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=clinicCode]')
  await page.evaluate(() => [...document.querySelectorAll('input[name=mode]')][3].click())
  await page.type('input[name=clinicCode]', 'B324')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Group company clinics"] tbody')?.textContent.includes('KLINIK BALA'), { timeout: 120000 })
  check(true, 'Grp. Com. Clinic Tracking: B324 lists its number of group companies')

  // Claim ASCII (CIMB): bordx KIS01540, cashless -> 571 visits (the file is downloaded).
  await page.goto(`${base}/invoice/claim-ascii/cimb`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=bordxFrom]')
  await page.type('input[name=bordxFrom]', 'KIS01540')
  await click(page, 'Generate')
  await page.waitForFunction(() => document.querySelector('[aria-label="Totals"]')?.textContent.includes('571'), { timeout: 120000 })
  check(true, 'Claim ASCII (CIMB): KIS01540 generates 571 visits')

  // Invoice > Tracking: cases EA12538600 - EA12538699.
  await page.goto(`${base}/invoice/tracking`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=caseFrom]')
  await page.type('input[name=caseFrom]', 'EA12538600')
  await page.type('input[name=caseTo]', 'EA12538699')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Invoices"] tbody')?.textContent.includes('EA125386'), { timeout: 120000 })
  check(true, 'Invoice Tracking: cases EA12538600 - EA12538699 are listed')

  // Invoice > Tracking (Jupiter): the same cases, with the MS No column.
  await page.goto(`${base}/invoice/tracking-jupiter`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('select[name=status]')
  await page.type('input[name=caseFrom]', 'EA12538600')
  await page.type('input[name=caseTo]', 'EA12538699')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Invoices"] tbody')?.textContent.includes('EA125386'), { timeout: 120000 })
  check(true, 'Invoice Tracking (Jupiter): cases EA12538600 - EA12538699 are listed')

  // Email Notification (Payment): new PV CP100471624 of 03/08/2021 is listed (nothing is sent).
  await page.goto(`${base}/invoice/email/payment`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('select[name=status]')
  await page.type('input[aria-label="PV Date From"]', '03/08/2021')
  await page.type('input[aria-label="PV Date To"]', '03/08/2021')
  await page.type('input[name=pvFrom]', 'CP100471624')
  await page.type('input[name=pvTo]', 'CP100471624')
  await page.select('select[name=status]', 'N - New')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Payments to e-mail"] tbody')?.textContent.includes('CP100471624'), { timeout: 120000 })
  check(true, 'Email Notification (Payment): new PV CP100471624 is listed')

  // Email Notification (Providers): new clinic PV CP100471757 of 03/08/2021 is listed (nothing is sent).
  await page.goto(`${base}/invoice/email/providers`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('select[name=status]')
  await page.type('input[aria-label="PV Date From"]', '03/08/2021')
  await page.type('input[aria-label="PV Date To"]', '03/08/2021')
  await page.type('input[name=pvFrom]', 'CP100471757')
  await page.type('input[name=pvTo]', 'CP100471757')
  await page.select('select[name=status]', 'N - New')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Clinic payments to e-mail"] tbody')?.textContent.includes('CP100471757'), { timeout: 120000 })
  check(true, 'Email Notification (Providers): new clinic PV CP100471757 is listed')

  // Email Notification (Rejection): rejected claim EA12514425 registered 01/07/2021 is listed (nothing is sent).
  await page.goto(`${base}/invoice/email/rejection`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('select[name=status]')
  await page.type('input[aria-label="Registration Date From"]', '01/07/2021')
  await page.type('input[aria-label="Registration Date To"]', '02/07/2021')
  await page.type('input[name=caseFrom]', 'EA12514425')
  await page.type('input[name=caseTo]', 'EA12514425')
  await page.select('select[name=status]', 'N - New')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Rejected claims"] tbody')?.textContent.includes('EA12514425'), { timeout: 120000 })
  check(true, 'Email Notification (Rejection): rejected claim EA12514425 is listed')

  // MBM Tracking: INS Type & MBM Type by payor, principals of payor DC (active).
  await page.goto(`${base}/membership/tracking`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('select[name=reportType] option:nth-child(12)')
  await page.evaluate(() => {
    const s = document.querySelector('select[name=reportType]')
    s.value = [...s.options].find((o) => o.textContent.trim() === 'INS Type & MBM Type').value
    s.dispatchEvent(new Event('change'))
  })
  await page.waitForFunction(() => [...document.querySelectorAll('select[name=selection] option')].some((o) => o.textContent.includes('By Payor')))
  await page.select('select[name=status]', '0 - Active')
  await page.type('input[name=payor]', 'DC - x')
  const searched = page.waitForResponse((r) => r.url().includes('/api/tracking/member/search'), { timeout: 240000 })
  await click(page, 'Search')
  const res = await searched
  if (!res.ok()) console.log('MBM Tracking search:', res.status(), (await res.text()).slice(0, 300), res.request().postData())
  await page.waitForFunction(() => document.querySelector('table[aria-label="Member tracking"] tbody')?.textContent.includes('DC'), { timeout: 30000 })
  check(true, 'MBM Tracking: INS Type & MBM Type by payor lists payor DC')

  // Claim Tracking: Ageband summary for payor DC, visits of 01/07/2021.
  await page.goto(`${base}/claim-bordx/claim-tracking/standard`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('select[name=reportType] option:nth-child(20)')
  await page.evaluate(() => {
    const s = document.querySelector('select[name=reportType]')
    s.value = [...s.options].find((o) => o.textContent.trim() === 'Ageband').value
    s.dispatchEvent(new Event('change'))
  })
  await page.waitForFunction(() => [...document.querySelectorAll('select[name=selection] option')].some((o) => o.textContent.includes('Summary')))
  await page.type('input[name=payor]', 'DC - x')
  await page.type('input[aria-label="DOV From"]', '01/07/2021')
  await page.type('input[aria-label="DOV To"]', '01/07/2021')
  const claimSearch = page.waitForResponse((r) => r.url().includes('/api/tracking/claim/search'), { timeout: 240000 })
  await click(page, 'Search')
  const claimRes = await claimSearch
  if (!claimRes.ok()) console.log('Claim Tracking search:', claimRes.status(), (await claimRes.text()).slice(0, 300))
  await page.waitForFunction(() => Number(document.querySelector('p[aria-label=Totals] strong')?.textContent) > 0, { timeout: 30000 })
  check(true, 'Claim Tracking: Ageband summary for payor DC, 01/07/2021 lists age bands')

  // MBM Tracking (Jupiter): its own choices (no Card Sent Date Listing).
  await page.goto(`${base}/membership/tracking-jupiter`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('select[name=reportType] option:nth-child(12)')
  await page.evaluate(() => {
    const s = document.querySelector('select[name=reportType]')
    s.value = [...s.options].find((o) => o.textContent.trim() === 'General Listing').value
    s.dispatchEvent(new Event('change'))
  })
  await page.waitForFunction(() => [...document.querySelectorAll('select[name=selection] option')].some((o) => o.textContent.includes('Print Label')))
  const options = await page.evaluate(() => [...document.querySelectorAll('select[name=selection] option')].map((x) => x.textContent))
  check(!options.some((o) => o.includes('Card Sent')), 'MBM Tracking (Jupiter): General Listing offers its own choices')

  // MBM Tracking (Simple): Head Count of July 2021.
  await page.goto(`${base}/membership/tracking-simple`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('select[name=reportType] option:nth-child(20)')
  await page.evaluate(() => {
    const s = document.querySelector('select[name=reportType]')
    s.value = [...s.options].find((o) => o.textContent.trim() === 'Head Count').value
    s.dispatchEvent(new Event('change'))
  })
  await page.waitForSelector('input[aria-label="Head Count From"]')
  await page.type('input[aria-label="Head Count From"]', '01/07/2021')
  await page.type('input[aria-label="Head Count To"]', '31/07/2021')
  const hcSearch = page.waitForResponse((r) => r.url().includes('/api/tracking/member-simple/search'), { timeout: 240000 })
  await click(page, 'Search')
  const hcRes = await hcSearch
  if (!hcRes.ok()) console.log('Head Count search:', hcRes.status(), (await hcRes.text()).slice(0, 300))
  await page.waitForFunction(() => document.querySelector('table[aria-label="Member tracking"] tbody')?.textContent.includes('July'), { timeout: 30000 })
  check(true, 'MBM Tracking (Simple): Head Count lists July 2021')
}
