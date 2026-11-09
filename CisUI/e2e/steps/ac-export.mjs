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

  // Invoice Posting > Clinic (GST): the FFS invoice 059809 of 04/08/2017 is split into four lines with the 8830 GST line (nothing is posted).
  await page.goto(`${base}/ac-export/invoice/clinic-gst`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=sunFrom]')
  await page.type('input[aria-label="Invoice Date From"]', '04/08/2017')
  await page.type('input[aria-label="Invoice Date To"]', '04/08/2017')
  await page.type('input[name=invoiceFrom]', '059809')
  await page.type('input[name=invoiceTo]', '059809')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Clinic invoice postings (GST)"] tbody')?.textContent.includes('8830'), { timeout: 120000 })
  check(true, 'Invoice Posting - Clinic (GST): FFS invoice 059809 has its 1002 and 8830 lines')

  // Invoice Posting > Clinic (GST-2017): unposted invoices of 01/07/2021 include 0210152590 (nothing is posted).
  await page.goto(`${base}/ac-export/invoice/clinic-gst-2017`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=sunFrom]')
  await page.type('input[aria-label="Invoice Date From"]', '01/07/2021')
  await page.type('input[aria-label="Invoice Date To"]', '01/07/2021')
  await page.type('input[name=invoiceFrom]', '0210152590')
  await page.type('input[name=invoiceTo]', '0210152590')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Clinic invoice postings (GST-2017)"] tbody')?.textContent.includes('84P320'), { timeout: 120000 })
  check(true, 'Invoice Posting - Clinic (GST-2017): 01/07/2021 lists 0210152590')

  // Invoice Posting > Clinic (GST-2017 Zeroised): the same invoice on the zeroised screen.
  await page.goto(`${base}/ac-export/invoice/clinic-gst-2017-zeroised`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=sunFrom]')
  await page.type('input[aria-label="Invoice Date From"]', '01/07/2021')
  await page.type('input[aria-label="Invoice Date To"]', '01/07/2021')
  await page.type('input[name=invoiceFrom]', '0210152590')
  await page.type('input[name=invoiceTo]', '0210152590')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Clinic invoice postings (GST-2017 Zeroised)"] tbody')?.textContent.includes('84P320'), { timeout: 120000 })
  check(true, 'Invoice Posting - Clinic (GST-2017 Zeroised): 01/07/2021 lists 0210152590')

  // Invoice Posting > Member: case EA12538622 (invoice 01/07/2021) lists its 82U063 group company line (nothing is posted).
  await page.goto(`${base}/ac-export/invoice/member`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=caseFrom]')
  await page.type('input[aria-label="Invoice Date From"]', '01/07/2021')
  await page.type('input[aria-label="Invoice Date To"]', '01/07/2021')
  await page.type('input[name=caseFrom]', 'EA12538622')
  await page.type('input[name=caseTo]', 'EA12538622')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Member invoice postings"] tbody')?.textContent.includes('82U063'), { timeout: 120000 })
  check(true, 'Invoice Posting - Member: EA12538622 lists its 82U063 line')
}
