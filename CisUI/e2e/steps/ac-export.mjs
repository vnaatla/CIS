// Phase 9 (A/C Export) browser steps. Run by run.mjs (full suite) and by only.mjs (E2E_STEPS=ac-export).
const click = (page, text) => page.evaluate((t) => [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === t).click(), text)

export async function acExport(page, base, check) {
  // Invoice Posting > Clinic: unposted invoices dated 01/07/2021 list 0210152590 as a credit and a debit line (nothing is posted).
  await page.goto(`${base}/ac-export/invoice/clinic`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('select[name=posted]')
  await page.type('input[aria-label="Invoice Date From"]', '01/07/2021')
  await page.type('input[aria-label="Invoice Date To"]', '01/07/2021')
  await page.type('input[name=invoiceFrom]', '0210152590')
  await page.type('input[name=invoiceTo]', '0210152590')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Clinic invoice postings"] tbody')?.textContent.includes('0210152590'), { timeout: 120000 })
  check(await page.evaluate(() => document.querySelector('table[aria-label="Clinic invoice postings"] tbody').textContent.includes('84P320')),
    'Invoice Posting - Clinic: 01/07/2021 lists 0210152590 with its 84P320 credit line')
}
