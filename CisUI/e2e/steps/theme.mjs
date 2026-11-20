// Colour theme browser steps (top-bar palette). Run by run.mjs (full suite) and by only.mjs (E2E_STEPS=theme).
// Switches to Navy blue, reloads, then back to Purple, so the browser is left on the default theme.
const state = (page) =>
  page.evaluate(() => ({
    attr: document.documentElement.dataset.theme ?? '',
    accent: getComputedStyle(document.documentElement).getPropertyValue('--accent').trim(),
    saved: localStorage.getItem('cis.theme'),
  }))

const pick = async (page, label) => {
  await page.click('button[aria-label=Theme]')
  await page.waitForSelector('.theme-menu')
  await page.evaluate((t) => [...document.querySelectorAll('.theme-menu .theme-option')].find((b) => b.textContent.trim() === t).click(), label)
  await page.waitForFunction(() => !document.querySelector('.theme-menu'))
}

export async function theme(page, base, check) {
  await page.goto(`${base}/membership/adjustment`, { waitUntil: 'networkidle0' })
  await pick(page, 'Navy blue')
  let s = await state(page)
  check(s.attr === 'navy' && s.saved === 'navy' && s.accent === '#1f4e8c', `Theme: Navy blue applies (accent #1f4e8c) and is remembered (${JSON.stringify(s)})`)

  await page.reload({ waitUntil: 'networkidle0' })
  s = await state(page)
  check(s.attr === 'navy', `Theme: Navy blue is still applied after a reload (${s.attr})`)

  await pick(page, 'Purple')
  s = await state(page)
  check(s.attr === '' && s.saved === null && s.accent === '#9c2b8e', `Theme: back to Purple clears the theme (${JSON.stringify(s)})`)
}
