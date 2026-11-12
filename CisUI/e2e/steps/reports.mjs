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
}
