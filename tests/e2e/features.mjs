// Browser test for Phase 2–3 features (sorting, filters, validation, autocomplete, tables, charts, pivots, printing).
//   npm run build && node tests/e2e/serve.mjs dist 4173 &
//   PLAYWRIGHT_MODULE=playwright CHROME_PATH=/path/to/chrome node tests/e2e/features.mjs
import { createRequire } from 'node:module'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { execFileSync } from 'node:child_process'

const require = createRequire(import.meta.url)
const { chromium } = require(process.env.PLAYWRIGHT_MODULE ?? 'playwright')
const BASE = process.env.BASE_URL ?? 'http://localhost:4173/'
const SHOTS = process.env.SHOTS ?? path.join(os.tmpdir(), 'excelpro-shots')
fs.mkdirSync(SHOTS, { recursive: true })

const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? `  (${detail})` : ''}`)
}

const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH })
const ctx = await browser.newContext({ viewport: { width: 1500, height: 900 }, acceptDownloads: true })
const page = await ctx.newPage()
const errors = []
page.on('pageerror', (e) => errors.push(e.message))
page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))

const sc = page.locator('[data-testid=grid-scroller]')
let box
const cellXY = (r, c) => ({ x: box.x + c * 88 + 44, y: box.y + r * 22 + 11 })
const clickCell = (r, c) => page.mouse.click(cellXY(r, c).x, cellXY(r, c).y)
const press = (k) => page.keyboard.press(k)
const type = (s) => page.keyboard.type(s)
const bar = async () => (await page.waitForTimeout(50), page.locator('[data-testid=formula-bar]').inputValue())
const goto = async (ref) => {
  await page.locator('[data-testid=name-box]').fill(ref)
  await page.locator('[data-testid=name-box]').press('Enter')
  await page.waitForTimeout(50)
}
const valueAt = async (ref) => (await goto(ref), bar())
const ribbonTab = (name) => page.getByRole('tab', { name, exact: true }).click()
const visibleRowTexts = () =>
  page.evaluate(() => {
    const out = {}
    for (const el of document.querySelectorAll('.xp-cell')) {
      const b = el.getBoundingClientRect()
      if (b.left < 120 && b.left > 30) out[Math.round(b.top)] = el.textContent
    }
    return Object.keys(out).sort((a, b) => a - b).map((k) => out[k])
  })
const paste = (text) =>
  page.evaluate((t) => {
    const dt = new DataTransfer()
    dt.setData('text/plain', t)
    document.querySelector('[data-testid=grid-keytrap]').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }))
  }, text)

await page.goto(BASE)
await page.waitForSelector('[data-testid=new-workbook]')
await page.click('[data-testid=new-workbook]')
await page.waitForSelector('[data-testid=grid-scroller]')
box = await sc.boundingBox()

// ---------------------------------------------------------------- data
await clickCell(0, 0)
await paste('Region\tProduct\tUnits\tPrice\nEast\tPens\t10\t2\nWest\tPads\t25\t4\nEast\tPads\t5\t4\nNorth\tPens\t40\t2\nWest\tPens\t15\t2\nNorth\tInk\t8\t9\n')
await page.waitForTimeout(150)

// ---------------------------------------------------------------- filter + sort
await clickCell(1, 0)
await press('Control+Shift+L')
await page.waitForSelector('[data-testid=filter-C]')
check('Ctrl+Shift+L adds filter buttons', (await page.locator('[data-testid^=filter-]').count()) === 4)
await page.click('[data-testid=filter-C]')
await page.waitForSelector('[data-testid=filter-menu]')
await page.getByRole('button', { name: 'Sort Z to A' }).click()
await page.waitForTimeout(100)
check('filter menu sorts Z→A', (await valueAt('C2')) === '40' && (await valueAt('A2')) === 'North', `${await valueAt('C2')} ${await valueAt('A2')}`)

await page.click('[data-testid=filter-A]')
await page.waitForSelector('[data-testid=filter-menu]')
await page.locator('[data-testid=filter-menu]').getByText('East', { exact: true }).click()
await page.locator('[data-testid=filter-menu]').getByRole('button', { name: 'OK' }).click()
await page.waitForTimeout(150)
let rows = await visibleRowTexts()
check('value filter hides rows', !rows.includes('East') && rows.includes('North') && rows.includes('West'), rows.join(','))
await page.click('[data-testid=filter-A]')
await page.getByRole('button', { name: /Clear filter from/ }).click()
await page.waitForTimeout(100)
rows = await visibleRowTexts()
check('clear filter restores rows', rows.filter((x) => x === 'East').length === 2, rows.join(','))

// condition filter
await page.click('[data-testid=filter-C]')
await page.getByRole('button', { name: 'By condition' }).click()
await page.getByLabel('Condition', { exact: true }).selectOption('gt')
await page.getByLabel('Condition value').fill('12')
await page.locator('[data-testid=filter-menu]').getByRole('button', { name: 'OK' }).click()
await page.waitForTimeout(100)
rows = await visibleRowTexts()
check('condition filter (Units > 12)', rows.length === 4, rows.join(','))
await ribbonTab('Data')
await page.getByRole('button', { name: 'Clear all filters on this sheet' }).click()
await page.getByRole('button', { name: /^Filter \(/ }).click()
await page.waitForTimeout(100)
check('filter removed', (await page.locator('[data-testid^=filter-]').count()) === 0)

// ---------------------------------------------------------------- autocomplete
await goto('F1')
await type('=SUMI')
await page.waitForSelector('[data-testid=formula-assist]')
const sugg = await page.locator('[data-testid=formula-assist] [role=option]').allInnerTexts()
check('autocomplete suggests functions', sugg.some((s) => s.includes('SUMIF')), sugg.slice(0, 3).join(' | '))
await press('Tab')
await page.waitForTimeout(50)
const edited = await page.locator('.xp-editor').inputValue()
check('Tab inserts function with (', edited === '=SUMIF(', edited)
await type('A2:A7,"East",C2:C7')
await page.waitForSelector('[data-testid=arg-hint]')
check('argument hint shows syntax', /SUMIF\(range, criteria/.test(await page.locator('[data-testid=arg-hint]').innerText()))
await press('Enter')
await page.waitForTimeout(150)
const sumifCell = await page.evaluate(() => [...document.querySelectorAll('.xp-cell')].map((e) => e.textContent))
check('SUMIF result', (await valueAt('F1')).startsWith('=SUMIF(') && sumifCell.includes('15'), sumifCell.join(','))

// ---------------------------------------------------------------- data validation
await goto('H1')
await ribbonTab('Data')
await page.getByRole('button', { name: 'Data validation' }).click()
await page.getByRole('menuitem', { name: /Data validation/ }).click()
await page.getByLabel('Source').fill('Open, Closed')
await page.getByRole('dialog').getByRole('button', { name: 'OK' }).click()
await page.waitForSelector('[data-testid=validation-dropdown]')
await page.click('[data-testid=validation-dropdown]')
await page.getByRole('option', { name: 'Closed' }).click()
check('validation dropdown picks value', (await valueAt('H1')) === 'Closed')
await type('Maybe')
await press('Enter')
await page.waitForSelector('[data-testid=validation-alert]')
check('invalid entry is stopped', (await page.locator('[data-testid=validation-alert]').count()) === 1)
await page.getByRole('dialog').getByRole('button', { name: 'Cancel' }).click()
check('invalid value not stored', (await valueAt('H1')) === 'Closed')

// ---------------------------------------------------------------- conditional formatting
await goto('C2:C7')
await ribbonTab('Home')
await page.getByRole('button', { name: /^Conditional formatting/ }).click()
await page.getByRole('menuitem', { name: 'Data bars (blue)' }).click()
await page.waitForTimeout(100)
check('data bars render', (await page.locator('[data-testid=databar]').count()) >= 5, String(await page.locator('[data-testid=databar]').count()))

// ---------------------------------------------------------------- table
await goto('A1')
await press('Control+t')
await page.getByRole('dialog').getByRole('button', { name: 'OK' }).click()
await page.waitForTimeout(100)
check('Ctrl+T creates table (Table Design tab)', (await page.getByRole('tab', { name: 'Table Design' }).count()) === 1)
await ribbonTab('Table Design')
await page.getByText('Total row').click()
await page.waitForTimeout(100)
check('table total row adds SUBTOTAL', /SUBTOTAL\(109/.test(await valueAt('D8')) || /SUBTOTAL\(109/.test(await valueAt('C8')), await valueAt('D8'))

// ---------------------------------------------------------------- chart
await goto('B1:C7')
await ribbonTab('Insert')
await page.getByRole('button', { name: 'Insert column chart' }).click()
await page.waitForSelector('[data-testid=chart]')
check('column chart inserted', (await page.locator('[data-testid=chart] svg rect').count()) > 3)
await page.getByLabel('Chart title').fill('Units by product')
await page.getByLabel('Chart title').press('Enter')
await page.waitForTimeout(100)
check('chart title edits', (await page.locator('[data-testid=chart]').innerText()).includes('Units by product'))
await page.getByLabel('Chart type').selectOption('pie')
await page.waitForTimeout(100)
check('chart type changes to pie', (await page.locator('[data-testid=chart] svg path').count()) >= 5)
await page.screenshot({ path: path.join(SHOTS, 'chart.png') })
await page.getByLabel('Close chart panel').click()

// ---------------------------------------------------------------- pivot
await goto('A2')
await ribbonTab('Insert')
await page.getByRole('button', { name: /Create a PivotTable/ }).click()
await page.getByLabel('Area for Region').selectOption('rows')
await page.getByLabel('Area for Units').selectOption('values')
await page.getByRole('dialog').getByRole('button', { name: 'Create' }).click()
await page.waitForSelector('[data-testid=pivot-panel]')
await page.waitForTimeout(150)
check('pivot created on new sheet', (await page.getByRole('tab', { name: /Pivot/ }).count()) >= 1)
const pivotCells = await page.evaluate(() => [...document.querySelectorAll('.xp-cell')].map((e) => e.textContent))
check('pivot totals', pivotCells.includes('Grand Total') && pivotCells.includes('103') && pivotCells.includes('48') && !pivotCells.includes('Total'), pivotCells.slice(0, 14).join(','))
await page.getByLabel('Area for Product').selectOption('cols')
await page.waitForTimeout(150)
const pivot2 = await page.evaluate(() => [...document.querySelectorAll('.xp-cell')].map((e) => e.textContent))
check('pivot field list updates layout', pivot2.includes('Pens') && pivot2.includes('Ink'), pivot2.slice(0, 14).join(','))
await page.screenshot({ path: path.join(SHOTS, 'pivot.png') })
await page.getByRole('tab', { name: 'Sheet1' }).click()

// ---------------------------------------------------------------- note, link, freeze
await goto('E2')
await press('Shift+F2')
await page.getByRole('textbox', { name: 'Note' }).fill('Check this price')
await page.getByRole('dialog').getByRole('button', { name: 'Save' }).click()
await ribbonTab('View')
await page.getByRole('button', { name: 'Freeze panes' }).click()
await page.getByRole('menuitem', { name: 'Freeze top row' }).click()

// ---------------------------------------------------------------- print preview + pdf
await press('Control+p')
await page.waitForSelector('[data-testid=preview-page]', { timeout: 10000 })
check('print preview renders a page', /Page 1 of \d/.test(await page.locator('[data-testid=page-count]').innerText()))
const [pdf] = await Promise.all([page.waitForEvent('download', { timeout: 20000 }), page.getByRole('button', { name: 'Download PDF' }).click()])
const pdfPath = path.join(SHOTS, 'export.pdf')
await pdf.saveAs(pdfPath)
const head = fs.readFileSync(pdfPath).subarray(0, 5).toString()
check('PDF downloaded', head === '%PDF-' && fs.statSync(pdfPath).size > 5000, `${head} ${fs.statSync(pdfPath).size} bytes`)
await page.getByRole('button', { name: 'Back' }).click()

// ---------------------------------------------------------------- xlsx export of phase 2/3 parts
await page.getByRole('button', { name: /Download/ }).first().click()
const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('dialog').getByRole('button', { name: /Download/ }).click()])
const xfile = path.join(SHOTS, 'features.xlsx')
await dl.saveAs(xfile)
let py = ''
try {
  py = execFileSync('python3', ['-c', `
import openpyxl
wb = openpyxl.load_workbook(${JSON.stringify(xfile)})
ws = wb['Sheet1']
print(len(ws.tables), len(ws.data_validations.dataValidation), len(ws.conditional_formatting), ws.freeze_panes, ws['E2'].comment is not None, len(wb.sheetnames))
`]).toString().trim()
} catch (e) {
  py = String(e)
}
check('XLSX keeps table, validation, CF, freeze, note, pivot sheet', py === '1 1 1 A2 True 2', py)

check('no page errors', errors.length === 0, errors.slice(0, 3).join(' | '))
await browser.close()
const failed = results.filter((r) => !r.ok)
console.log(`\n${results.length - failed.length}/${results.length} checks passed`)
process.exit(failed.length ? 1 : 0)
