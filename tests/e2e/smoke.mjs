// Browser smoke test for the built app (Playwright + Chromium).
//   npm run build && node tests/e2e/serve.mjs dist 4173 &
//   PLAYWRIGHT_MODULE=playwright CHROME_PATH=/path/to/chrome node tests/e2e/smoke.mjs
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
const ctx = await browser.newContext({ viewport: { width: 1400, height: 860 }, acceptDownloads: true })
await ctx.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: BASE })
const page = await ctx.newPage()
const errors = []
page.on('pageerror', (e) => errors.push(e.message))
page.on('console', (m) => m.type() === 'error' && errors.push(m.text()))

const sc = page.locator('[data-testid=grid-scroller]')
let box
const cellXY = (r, c) => ({ x: box.x + c * 88 + 44, y: box.y + r * 22 + 11 })
const clickCell = async (r, c, opts) => {
  const { x, y } = cellXY(r, c)
  await page.mouse.click(x, y, opts)
}
const shiftClick = async (r, c) => {
  await page.keyboard.down('Shift')
  await clickCell(r, c)
  await page.keyboard.up('Shift')
}
const bar = async () => {
  await page.waitForTimeout(40)
  return page.locator('[data-testid=formula-bar]').inputValue()
}
const nameBox = async () => {
  await page.waitForTimeout(40)
  return page.locator('[data-testid=name-box]').inputValue()
}
const cellText = async (r, c) => {
  const { x, y } = cellXY(r, c)
  return page.evaluate(
    ([x, y]) => {
      for (const el of document.querySelectorAll('.xp-cell')) {
        const b = el.getBoundingClientRect()
        if (x >= b.left && x <= b.left + Math.min(b.width, 88) && y >= b.top && y <= b.bottom) return el.textContent
      }
      return ''
    },
    [x, y],
  )
}
const type = async (s) => {
  await page.keyboard.type(s)
}
const press = (k) => page.keyboard.press(k)

// ---------------------------------------------------------------- start
await page.goto(BASE)
await page.waitForSelector('[data-testid=new-workbook]')
await page.click('[data-testid=new-workbook]')
await page.waitForSelector('[data-testid=grid-scroller]')
box = await sc.boundingBox()
check('blank workbook opens', (await nameBox()) === 'A1')

// typing values and formulas
await clickCell(0, 0)
await type('10')
await press('Enter')
await type('20')
await press('Enter')
await type('=A1+A2')
await press('Enter')
check('formula calculates', (await cellText(2, 0)) === '30', await cellText(2, 0))
await clickCell(2, 0)
check('formula bar shows formula', (await bar()) === '=A1+A2', await bar())
check('name box shows address', (await nameBox()) === 'A3')

await clickCell(0, 1)
await type('=sum(a1:a3')
await press('Enter')
check('function auto-uppercased & parens closed', (await cellText(0, 1)) === '60')
await clickCell(0, 1)
check('normalized formula text', (await bar()) === '=SUM(A1:A3)', await bar())

// dependent recalculation
await clickCell(0, 0)
await type('15')
await press('Enter')
check('dependents recalculate', (await cellText(2, 0)) === '35' && (await cellText(0, 1)) === '70', `${await cellText(2, 0)} ${await cellText(0, 1)}`)

// point mode: click to insert reference
await clickCell(4, 3)
await type('=')
await clickCell(0, 0)
await type('*2')
await press('Enter')
check('click-to-insert reference in formula', (await cellText(4, 3)) === '30', await cellText(4, 3))

// formatting + undo/redo
await clickCell(0, 1)
await press('Control+b')
const weight = () =>
  page.evaluate(([x, y]) => {
    const el = document.elementsFromPoint(x, y).find((e) => e.classList.contains('xp-cell'))
    return el ? getComputedStyle(el).fontWeight : ''
  }, [cellXY(0, 1).x, cellXY(0, 1).y])
check('Ctrl+B bolds', (await weight()) === '700')
await press('Control+z')
check('undo removes bold', (await weight()) !== '700')
await press('Control+y')
check('redo restores bold', (await weight()) === '700')

// fill handle series
await clickCell(0, 0)
await shiftClick(1, 0)
await clickCell(0, 0)
await type('1')
await press('Enter')
await type('2')
await press('Enter')
await clickCell(0, 0)
await shiftClick(1, 0)
const fh = { x: box.x + 88, y: box.y + 44 }
await page.mouse.move(fh.x - 1, fh.y - 1)
await page.mouse.down()
await page.mouse.move(fh.x - 1, box.y + 22 * 6 - 5, { steps: 6 })
await page.mouse.up()
check('fill handle extends series', (await cellText(5, 0)) === '6', await cellText(5, 0))

// copy / paste with relative reference adjustment
await clickCell(0, 1)
await press('Control+c')
await clickCell(1, 2)
await press('Control+v')
await page.waitForTimeout(100)
await clickCell(1, 2)
check('copy/paste adjusts references', (await bar()) === '=SUM(B2:B4)', await bar())

// external paste (TSV like Excel puts on the clipboard)
await clickCell(8, 0)
await page.evaluate(() => {
  const dt = new DataTransfer()
  dt.setData('text/plain', 'Item\tQty\tPrice\nApple\t3\t$1.25\nPear\t5\t$0.80\n')
  document.querySelector('[data-testid=grid-keytrap]').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }))
})
await page.waitForTimeout(100)
check('external TSV paste', (await cellText(9, 2)) === '$1.25' && (await cellText(10, 0)) === 'Pear', `${await cellText(9, 2)} ${await cellText(10, 0)}`)

// keyboard navigation
await clickCell(9, 0)
await press('Control+ArrowRight')
check('Ctrl+Arrow jumps to data edge', (await nameBox()) === 'C10', await nameBox())
await press('Tab')
check('Tab moves right', (await nameBox()) === 'D10', await nameBox())
await press('Control+Home')
check('Ctrl+Home goes to A1', (await nameBox()) === 'A1')

// in-cell edit with F2 and Escape cancel
await clickCell(0, 0)
await press('F2')
await type('99')
await press('Escape')
check('Escape cancels edit', (await cellText(0, 0)) === '1')

// insert row via context menu shifts formulas
await page.mouse.click(box.x - 20, cellXY(0, 0).y, { button: 'right' })
await page.getByRole('menuitem', { name: /Insert row above/ }).click()
await clickCell(1, 1)
check('insert row shifts formulas', (await bar()) === '=SUM(A2:A4)', (await bar()) || 'empty:' + (await nameBox()))
await press('Control+z')

// resize a column by dragging the header border
const hdr = await page.locator('[data-testid=col-headers]').boundingBox()
await page.mouse.move(box.x + 88 * 4 - 1, hdr.y + 10)
await page.mouse.down()
await page.mouse.move(box.x + 88 * 4 + 40, hdr.y + 10, { steps: 5 })
await page.mouse.up()
const eWidth = await page.evaluate(() => [...document.querySelectorAll('[data-testid=col-headers] .xp-hdr')].find((e) => e.textContent === 'D')?.style.width)
check('column resize by dragging', eWidth === '129px', eWidth)
await press('Control+z')

// sheets
await page.getByRole('button', { name: 'Add sheet' }).click()
check('add sheet', (await page.getByRole('tab', { name: 'Sheet2' }).count()) === 1)
box = await sc.boundingBox()
await clickCell(0, 0)
await type('=Sheet1!A1*100')
await press('Enter')
check('cross-sheet reference', (await cellText(0, 0)) === '100', await cellText(0, 0))
await page.getByRole('tab', { name: 'Sheet2' }).dblclick()
await page.locator('input[aria-label="Sheet name"]').fill('Totals Q1')
await press('Enter')
await page.getByRole('tab', { name: 'Sheet1' }).click()
await clickCell(6, 0)
await type("='Totals Q1'!A1+1")
await press('Enter')
check('reference to sheet with space', (await cellText(6, 0)) === '101', (await cellText(6, 0)) + ' tabs=' + (await page.getByRole('tab').allInnerTexts()).join('|'))
await page.screenshot({ path: path.join(SHOTS, 'after-sheets.png') })

// find
await press('Control+f')
await page.getByLabel('Find what').fill('Pear')
await page.getByRole('button', { name: 'Find next' }).click()
check('find selects match', (await nameBox()) === 'A11', await nameBox())
await press('Escape')

// save + reload
await page.waitForFunction(() => document.querySelector('[data-testid=save-status]')?.textContent?.includes('Saved'), null, { timeout: 5000 })
await page.screenshot({ path: path.join(SHOTS, 'edited.png') })
await page.reload()
await page.waitForSelector('[data-testid=grid-scroller]')
box = await sc.boundingBox()
check('workbook survives reload', (await cellText(6, 0)) === '101' && (await cellText(0, 1)) === '6', `${await cellText(6, 0)} / ${await cellText(0, 1)}`)
check('sheet rename survives reload', (await page.getByRole('tab', { name: 'Totals Q1' }).count()) === 1)

// export xlsx and verify in openpyxl
await page.getByRole('button', { name: /Download/ }).first().click()
const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('dialog').getByRole('button', { name: /Download/ }).click()])
const xfile = path.join(SHOTS, 'export.xlsx')
await dl.saveAs(xfile)
let py = ''
try {
  py = execFileSync('python3', ['-c', `import openpyxl;wb=openpyxl.load_workbook(${JSON.stringify(xfile)});ws=wb['Sheet1'];print(wb.sheetnames, ws['B1'].value, ws['A7'].value, ws['B1'].font.b)`]).toString().trim()
} catch (e) {
  py = String(e)
}
check('XLSX export keeps formulas & bold', py.includes("'Totals Q1'") && py.includes('=SUM(A1:A3)') && py.includes("='Totals Q1'!A1+1") && py.endsWith('True'), py)

// import a real xlsx
await page.goto(BASE + '#/')
await page.waitForSelector('[data-testid=new-workbook]')
const fixture = path.resolve('tests/fixtures/openpyxl-sample.xlsx')
await page.locator('input[type=file]').setInputFiles(fixture)
await page.waitForSelector('text=Import report', { timeout: 10000 })
const reportText = await page.getByRole('dialog').innerText()
check('import report lists unsupported features', /TEXTSPLIT/.test(reportText))
await page.getByRole('button', { name: 'OK' }).click()
box = await sc.boundingBox()
await page.screenshot({ path: path.join(SHOTS, 'imported.png') })
await page.locator('[data-testid=name-box]').fill('D7')
await page.locator('[data-testid=name-box]').press('Enter')
check('imported formula value', (await bar()) === '=SUM(D2:D5)')

// large CSV performance
const rows = 100000
let csv = 'id,name,amount,date\n'
for (let i = 1; i <= rows; i++) csv += `${i},Item ${i},${(i * 1.5).toFixed(2)},2024-01-${String((i % 28) + 1).padStart(2, '0')}\n`
const big = path.join(SHOTS, 'big.csv')
fs.writeFileSync(big, csv)
await page.goto(BASE + '#/')
await page.waitForSelector('[data-testid=new-workbook]')
const t0 = Date.now()
await page.locator('input[type=file]').setInputFiles(big)
await page.waitForSelector('[data-testid=grid-scroller]', { timeout: 60000 })
await page.waitForFunction(() => document.querySelectorAll('.xp-cell').length > 10, null, { timeout: 60000 })
const tImport = Date.now() - t0
box = await sc.boundingBox()
const t1 = Date.now()
await page.locator('[data-testid=name-box]').fill('A100001')
await page.locator('[data-testid=name-box]').press('Enter')
await page.waitForFunction(() => [...document.querySelectorAll('.xp-cell')].some((e) => e.textContent === '100000'), null, { timeout: 5000 }).catch(() => {})
const tJump = Date.now() - t1
const lastText = await page.evaluate(() => [...document.querySelectorAll('.xp-cell')].map((e) => e.textContent).includes('100000'))
await clickCell(0, 4)
const t2 = Date.now()
await page.locator('[data-testid=name-box]').fill('E1')
await page.locator('[data-testid=name-box]').press('Enter')
await type('=SUM(C2:C100001)')
await press('Enter')
await page.waitForFunction(() => [...document.querySelectorAll('.xp-cell')].some((e) => /^7\.50*E\+09$|^7500075000$/.test(e.textContent)), null, { timeout: 5000 }).catch(() => {})
const tSum = Date.now() - t2
const sumOk = await page.evaluate(() => [...document.querySelectorAll('.xp-cell')].some((e) => /^7\.50*E\+09$|^7500075000$/.test(e.textContent)))
check(`100k-row CSV import (${tImport} ms), jump to end (${tJump} ms), SUM over 100k (${tSum} ms)`, lastText && sumOk)
// scroll performance: average frame time while scrolling
const scrollMs = await page.evaluate(async () => {
  const s = document.querySelector('[data-testid=grid-scroller]')
  const times = []
  let last = performance.now()
  for (let i = 0; i < 60; i++) {
    s.scrollTop += 400
    await new Promise((r) => requestAnimationFrame(r))
    const now = performance.now()
    times.push(now - last)
    last = now
  }
  return times.reduce((a, b) => a + b, 0) / times.length
})
check(`scrolling 100k rows: avg frame ${scrollMs.toFixed(1)} ms`, scrollMs < 50)
await page.screenshot({ path: path.join(SHOTS, 'big.png') })

check('no page errors', errors.length === 0, errors.slice(0, 3).join(' | '))
await browser.close()
const failed = results.filter((r) => !r.ok)
console.log(`\n${results.length - failed.length}/${results.length} checks passed`)
process.exit(failed.length ? 1 : 0)
