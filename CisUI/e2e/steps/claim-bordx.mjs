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
}
