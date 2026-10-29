// Engine benchmark: build, edit and serialize workbooks of increasing size.
//   node --import tsx scripts/bench.ts
import { Workbook } from '../src/engine/workbook'
import { exportXlsx, importXlsx } from '../src/files/xlsx'

const ms = (t: number) => `${(performance.now() - t).toFixed(0)} ms`
for (const rows of [10_000, 50_000, 100_000]) {
  const cells: (string | number)[][] = [['id', 'qty', 'price', 'total', 'running']]
  for (let i = 1; i <= rows; i++) cells.push([i, (i % 17) + 1, +(i * 0.37).toFixed(2), `=B${i + 1}*C${i + 1}`, i === 1 ? `=D2` : `=E${i}+D${i + 1}`])
  const doc = Workbook.blank('bench').toDoc()
  doc.sheets[0].cells = cells
  let t = performance.now()
  const wb = Workbook.fromDoc(doc)
  const tBuild = ms(t)
  const s = wb.order[0]
  t = performance.now()
  wb.setInputs(s, [{ r: 1, c: 1, input: '100' }]) // changes every running total below
  const tEdit = ms(t)
  t = performance.now()
  wb.setInputs(s, [{ r: rows, c: 1, input: '5' }]) // changes only the last rows
  const tEditEnd = ms(t)
  t = performance.now()
  wb.undo()
  const tUndo = ms(t)
  t = performance.now()
  wb.insertRows(s, 10, 1)
  const tInsert = ms(t)
  t = performance.now()
  const ser = JSON.stringify(wb.serializeSheet(s))
  const tSer = ms(t)
  t = performance.now()
  const x = await exportXlsx(wb)
  const tExp = ms(t)
  t = performance.now()
  await importXlsx(x, 'b.xlsx')
  const tImp = ms(t)
  console.log(
    `| ${rows.toLocaleString()} rows × 5 cols (${(rows * 2).toLocaleString()} formulas) | ${tBuild} | ${tEdit} | ${tEditEnd} | ${tUndo} | ${tInsert} | ${tSer} (${(ser.length / 1e6).toFixed(1)} MB) | ${tExp} (${(x.length / 1e6).toFixed(1)} MB) | ${tImp} |`,
  )
}
