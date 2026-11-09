// Phase 8 (Receipt) browser steps. Run by run.mjs (full suite) and by only.mjs (E2E_STEPS=receipt).
const click = (page, text) => page.evaluate((t) => [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === t).click(), text)
const toast = (page, text) => page.waitForFunction((t) => document.body.innerText.includes(t), { timeout: 60000 }, text)

export async function receipt(page, base, check) {
  // Receipt from Member: the restored data has no member debit notes, so only the scan check is exercised (nothing is saved).
  await page.goto(`${base}/receipt/member`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=caseFrom]')
  await click(page, 'Search')
  await toast(page, 'Please Enter Worksheet No')
  check(true, 'Receipt from Member: Search without a worksheet no. says "Please Enter Worksheet No"')

  // Receipt from Group Company: bordx KIT00036 is listed as received (RP37929), and ticking it is refused (nothing is saved).
  await page.goto(`${base}/receipt/payor`, { waitUntil: 'domcontentloaded' })
  await page.waitForSelector('input[name=bordxFrom]')
  await page.type('input[name=bordxFrom]', 'KIT00036')
  await click(page, 'Search')
  await page.waitForFunction(() => document.querySelector('table[aria-label="Bordx receipts"] tbody')?.textContent.includes('RP37929'), { timeout: 120000 })
  await page.click('table[aria-label="Bordx receipts"] tbody input[type=checkbox]')
  await toast(page, 'Record has been paid !')
  check(true, 'Receipt from Group Company: KIT00036 lists receipt RP37929 and cannot be ticked')
}
