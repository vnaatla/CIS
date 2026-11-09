// Phase 7 (Payment) browser steps. Run by run.mjs (full suite) and by only.mjs (E2E_STEPS=payment).
const click = (page, text) => page.evaluate((t) => [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === t).click(), text)

export async function payment(page, base, check) {
  // Payment to Clinic: unpaid cases invoiced 01/08/2021 - 05/08/2021 include EA12565564 (nothing is saved).
  await page.goto(`${base}/payment/clinic`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=scan]')
  await page.click('input[name=scan]')
  await page.evaluate(() => {
    const sel = document.querySelector('select[name=payment]')
    sel.value = [...sel.options].find((o) => o.textContent.includes('Unpaid')).value
    sel.dispatchEvent(new Event('change'))
  })
  await page.type('input[aria-label="Invoice Date From"]', '01/08/2021')
  await page.type('input[aria-label="Invoice Date To"]', '05/08/2021')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Clinic payments"] tbody')?.textContent.includes('EA12565564'), { timeout: 120000 })
  check(true, 'Payment to Clinic: unpaid cases invoiced 1-5 Aug 2021 include EA12565564')

  // Payment to Clinic - Invoice: G029, May 2021 lists its paid PV CP100468765 and the unpaid month (nothing is saved).
  await page.goto(`${base}/payment/clinic-invoice`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=clinic]')
  await page.type('input[name=clinic]', 'G029 - GLENEAGLES')
  await page.type('input[name=yearFrom]', '2021')
  await page.type('input[name=yearTo]', '2021')
  for (const n of ['monthFrom', 'monthTo'])
    await page.evaluate((name) => {
      const sel = document.querySelector(`select[name=${name}]`)
      sel.value = [...sel.options].find((o) => o.textContent.trim() === '05').value
      sel.dispatchEvent(new Event('change'))
    }, n)
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Clinic invoice payments"] tbody')?.textContent.includes('CP100468765'), { timeout: 180000 })
  check(true, 'Payment to Clinic - Invoice: G029 May 2021 lists PV CP100468765')

  // Payment to Member: scanning SLK004743875 lists the unpaid case with the patient's name (nothing is saved).
  await page.goto(`${base}/payment/member`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=caseFrom]')
  await page.type('input[name=caseFrom]', 'SLK004743875')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Member payments"] tbody')?.textContent.includes('CHOOK SHENG CAI'), { timeout: 120000 })
  check(true, 'Payment to Member: SLK004743875 lists CHOOK SHENG CAI')

  // Payment Tracking: Pay to Clinic, By PV No for CP100472166 shows its 21 cases.
  await page.goto(`${base}/payment/tracking`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('select[name=reportType]')
  await page.evaluate(() => {
    const sel = document.querySelector('select[name=reportType]')
    sel.value = [...sel.options].find((o) => o.textContent.trim() === 'Pay to Clinic').value
    sel.dispatchEvent(new Event('change'))
    const by = document.querySelector('select[name=selection]')
    by.value = 'By PV No'
    by.dispatchEvent(new Event('change'))
  })
  await page.type('input[name=pvFrom]', 'CP100472166')
  await page.type('input[name=pvTo]', 'CP100472166')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Payment tracking"] tbody')?.textContent.includes('CP100472166'), { timeout: 120000 })
  check(await page.evaluate(() => document.querySelector('table[aria-label="Payment tracking"] tbody').textContent.includes('21')), 'Payment Tracking: PV CP100472166 has 21 cases')
}
