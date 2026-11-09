// Runs only the given step modules after one login (no wrong-password attempt), then logs out. Usage:
//   $env:E2E_PASSWORD='...'; $env:E2E_STEPS='claim-bordx'; node e2e/only.mjs
// Keeps the run short and the load on the local SQL Server low while a phase is being built; run.mjs runs everything.
import puppeteer from 'puppeteer-core'
import { existsSync } from 'node:fs'
import { claimBordx } from './steps/claim-bordx.mjs'
import { payment } from './steps/payment.mjs'
import { receipt } from './steps/receipt.mjs'

const modules = { 'claim-bordx': claimBordx, payment, receipt }
const base = process.env.E2E_URL ?? 'http://localhost:4200'
const user = process.env.E2E_USER ?? 'DEVADMIN'
const password = process.env.E2E_PASSWORD
const steps = (process.env.E2E_STEPS ?? '').split(',').map((s) => s.trim()).filter(Boolean)
const browserPath =
  process.env.E2E_BROWSER ??
  ['C:/Program Files/Google/Chrome/Application/chrome.exe', 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', 'C:/Program Files/Microsoft/Edge/Application/msedge.exe']
    .find(existsSync)
if (!password || !steps.length || steps.some((s) => !modules[s])) {
  console.error(`Set E2E_PASSWORD and E2E_STEPS (one or more of: ${Object.keys(modules).join(', ')}).`)
  process.exit(2)
}

let failed = 0
const check = (ok, text) => {
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${text}`)
  if (!ok) failed++
}
const browser = await puppeteer.launch({ executablePath: browserPath, headless: true, args: ['--no-first-run'] })
try {
  const page = await browser.newPage()
  await page.setViewport({ width: 1400, height: 900 })
  await page.goto(`${base}/login`, { waitUntil: 'networkidle0' })
  const [userBox, passwordBox] = await page.$$('form.card input')
  await userBox.type(user)
  await passwordBox.type(password)
  await page.click('form.card button[type=submit]')
  await page.waitForSelector('header.topbar', { timeout: 60000 })
  for (const s of steps) await modules[s](page, base, check)
  await page.click('header.topbar button.link:last-of-type')
  await page.waitForSelector('form.card')
} finally {
  await browser.close()
}
console.log(failed === 0 ? '\nAll checks passed.' : `\n${failed} check(s) failed.`)
process.exit(failed === 0 ? 0 : 1)
