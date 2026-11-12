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
}
