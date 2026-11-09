// Phase 6 (Claim Bordx) browser steps. Run by run.mjs (full suite) and by only.mjs (E2E_STEPS=claim-bordx).
const pick = (page, name, text) =>
  page.evaluate(
    (n, t) => {
      const sel = document.querySelector(`select[name=${n}]`)
      sel.value = [...sel.options].find((o) => o.textContent.trim() === t).value
      sel.dispatchEvent(new Event('change'))
    },
    name,
    text,
  )
const click = (page, text) => page.evaluate((t) => [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === t).click(), text)

export async function claimBordx(page, base, check) {
  // Insert Worksheet: CREDIT NOTE-CN for group company M005 lists EA12573444; ticking it is allowed (nothing is saved).
  await page.goto(`${base}/claim-bordx/insert-worksheet`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('select[name=bordxType]')
  await pick(page, 'bordxType', 'CREDIT NOTE-CN')
  await page.type('input[name=groupCompany]', 'X|M005')
  await click(page, 'Search')
  await page.waitForFunction(() => [...document.querySelectorAll('table[aria-label="Worksheets"] tbody tr')].some((tr) => tr.textContent.includes('EA12573444')), {
    timeout: 120000,
  })
  await click(page, 'Check All')
  await page.waitForFunction(() => document.querySelector('[aria-label=Totals]')?.textContent.includes('Selected: 1'), { timeout: 15000 })
  check(true, 'Insert Worksheet: CREDIT NOTE-CN for M005 lists EA12573444 and Check All ticks it')

  // Remove Worksheet: bordx KIV03270 lists EA12476058 (nothing is removed).
  await page.goto(`${base}/claim-bordx/remove-worksheet`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=bordxNo]')
  await page.type('input[name=bordxNo]', 'KIV03270')
  await click(page, 'Go')
  await page.waitForFunction(() => [...document.querySelectorAll('table[aria-label="Bordereaux worksheets"] tbody tr')].some((tr) => tr.textContent.includes('EA12476058')), {
    timeout: 60000,
  })
  check(true, 'Remove Worksheet: bordx KIV03270 lists EA12476058')

  // ClaimBordxDate: KIV03270 shows its line (nothing is updated).
  await page.goto(`${base}/claim-bordx/bordx-date`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=bordxNo]')
  await page.type('input[name=bordxNo]', 'KIV03270')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Bordx lines"] tbody')?.textContent.includes('KIV03270'), { timeout: 60000 })
  check(true, 'ClaimBordxDate: KIV03270 shows its bordx date and posting status')

  // Bordx Delivery: Send to Grp. Company and Send to A/C for DAYANG ENTERPRISE SDN BHD from 01/01/2021 (nothing is updated).
  for (const [path, label] of [['delivery/send-to-payor', 'Send to Grp. Company'], ['delivery/send-to-ac', 'Send to A/C'], ['ac-period', 'Bor Acc Period']]) {
    await page.goto(`${base}/claim-bordx/${path}`, { waitUntil: 'domcontentloaded' })
    await page.waitForSelector('input[name=company]')
    await page.type('input[name=company]', 'DAYANG ENTERPRISE SDN BHD')
    await page.type('input[aria-label="Bordx Date From"]', '01/01/2021')
    await click(page, 'Search')
    await page.waitForFunction(() => document.querySelectorAll('table[aria-label="Bordereaux"] tbody tr:not(.empty-row)').length >= 17, { timeout: 60000 })
    check(true, `${label}: DAYANG ENTERPRISE SDN BHD lists its bordereaux from 2021`)
  }

  // Monitoring: Bordx Control (DAYANG from 2021), Outstanding Bordx (ALAM TEKNOKRAT), Bordx & Worksheet (DAYANG KDC55049).
  const monitor = async (kind, company, fill, rows, label) => {
    await page.goto(`${base}/claim-bordx/monitoring/${kind}`, { waitUntil: 'domcontentloaded' })
    await page.waitForSelector('input[name=company]')
    await page.type('input[name=company]', company)
    await fill()
    await click(page, 'Search')
    await page.waitForFunction((n) => document.querySelectorAll('app-data-grid tbody tr:not(.empty-row)').length >= n, { timeout: 120000 }, rows)
    check(true, label)
  }
  await monitor('control', 'DAYANG ENTERPRISE SDN BHD', () => page.type('input[aria-label="Bordx Date From"]', '01/01/2021'), 17, 'Bordx Control: DAYANG ENTERPRISE SDN BHD from 2021 lists 17 bordereaux')
  await monitor('outstanding', 'ALAM TEKNOKRAT SDN BHD', async () => {}, 4, 'Outstanding Bordx: ALAM TEKNOKRAT SDN BHD lists 4 outstanding bordereaux')
  await monitor('bordx-worksheet', 'DAYANG ENTERPRISE SDN BHD', async () => {
    await page.type('input[name=bordxFrom]', 'KDC55049')
    await page.type('input[name=bordxTo]', 'KDC55049')
  }, 1, 'Bordx & Worksheet: DAYANG ENTERPRISE SDN BHD bordx KDC55049 lists its worksheets')

  // Print Bordx: KDC55049 lists its worksheet (nothing is printed).
  await page.goto(`${base}/claim-bordx/print/standard`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=bordxNo]')
  await page.type('input[name=bordxNo]', 'KDC55049')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Bordereaux worksheets"] tbody')?.textContent.includes('EA12409034'), { timeout: 120000 })
  check(true, 'Print Bordx: KDC55049 lists EA12409034')

  // A client variant (BSN) uses the same page without company / bill type.
  await page.goto(`${base}/claim-bordx/print/bsn`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=bordxNo]')
  await page.type('input[name=bordxNo]', 'KDC55049')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Bordereaux worksheets"] tbody')?.textContent.includes('EA12409034'), { timeout: 120000 })
  check(!(await page.$('select[name=billType]')), 'Print Bordx (BSN): KDC55049 lists EA12409034 without the bill type')

  // Insert Invoice: A242, registered 1-4 Aug 2021, lists invoice 7486 (nothing is saved).
  await page.goto(`${base}/claim-bordx/insert-invoice`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=groupCompany]')
  await page.type('input[name=groupCompany]', 'A242 - ASIAFLEX PRODUCTS SDN. BHD.')
  await page.type('input[aria-label="Register Date From"]', '01/08/2021')
  await page.type('input[aria-label="Register Date To"]', '04/08/2021')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Invoices pending bordx"] tbody')?.textContent.includes('7486'), { timeout: 120000 })
  check(true, 'Insert Invoice: A242 registered 1-4 Aug 2021 lists invoice 7486')

  // BSN-PrintSummary: KDC55481 has 772 claims.
  await page.goto(`${base}/claim-bordx/print/bsn-summary`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=bordx0]')
  await page.type('input[name=bordx0]', 'KDC55481')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="BSN bordereaux"] tbody')?.textContent.includes('772'), { timeout: 120000 })
  check(true, 'BSN-PrintSummary: KDC55481 has 772 claims')
}
