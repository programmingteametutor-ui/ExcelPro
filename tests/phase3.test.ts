import { test } from 'node:test'
import assert from 'node:assert/strict'
import { Workbook } from '../src/engine/workbook'
import { convertToRange, createTable, maybeExtendTable, resolveStructuredRefs, updateTable } from '../src/engine/features/tables'
import { computePivot, createPivot, readPivotSource, refreshPivot, removePivot, updatePivot } from '../src/engine/features/pivot'
import { flashFill, goalSeek, inferFlashFill, removeDuplicates, textToColumns, transformText } from '../src/engine/features/clean'
import { setPageSetup } from '../src/engine/features/data'
import { layoutPages, pageBreaks } from '../src/print/paginate'
import { addChart, extractChartData, guessChartLayout, histogram, newChart, removeChart, updateChart } from '../src/charts/data'

const R = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 })

function sales() {
  const wb = Workbook.blank('T')
  const s = wb.order[0]
  const rows = [
    ['Region', 'Product', 'Date', 'Units', 'Revenue'],
    ['East', 'Pens', '2024-01-15', '10', '25'],
    ['West', 'Pens', '2024-02-03', '4', '10'],
    ['East', 'Books', '2024-02-20', '2', '30'],
    ['West', 'Books', '2024-03-11', '5', '75'],
    ['East', 'Pens', '2024-03-30', '6', '15'],
  ]
  rows.forEach((row, r) => row.forEach((v, c) => wb.setInputs(s, [{ r, c, input: v }])))
  return { wb, s }
}

test('tables: create, structured references, totals, auto-extend, rename, convert', () => {
  const { wb, s } = sales()
  const { id, error } = createTable(wb, s, R(0, 0, 5, 4), { hasHeader: true })
  assert.equal(error, undefined)
  const t = wb.meta(s).tables![0]
  assert.equal(t.name, 'Table1')
  assert.deepEqual(wb.meta(s).autoFilter!.range, R(0, 0, 5, 4))
  wb.setInputs(s, [{ r: 8, c: 0, input: '=SUM(Table1[Revenue])' }])
  assert.equal(wb.getRaw(s, 8, 0), '=SUM($E$2:$E$6)')
  assert.equal(wb.getValue(s, 8, 0), 155)
  assert.equal(resolveStructuredRefs(wb, '=ROWS(Table1[#All])+COUNTA(Table1[#Headers])', s, 0), '=ROWS($A$1:$E$6)+COUNTA($A$1:$E$1)')
  assert.equal(resolveStructuredRefs(wb, '=Table1[@Units]*2', s, 3), '=$D$4*2')
  assert.equal(resolveStructuredRefs(wb, '="Table1[Revenue]"', s, 0), '="Table1[Revenue]"')
  assert.equal(updateTable(wb, s, id!, { totalRow: true }), null)
  assert.deepEqual(wb.meta(s).tables![0].range, R(0, 0, 6, 4))
  assert.equal(wb.getRaw(s, 6, 4), '=SUBTOTAL(109,E2:E6)')
  assert.equal(wb.getValue(s, 6, 0), 'Total')
  updateTable(wb, s, id!, { totals: { 3: 'average', 4: 'sum' } })
  assert.equal(wb.getValue(s, 6, 3), 5.4)
  updateTable(wb, s, id!, { totalRow: false })
  assert.equal(wb.getRaw(s, 6, 4), null)
  assert.equal(maybeExtendTable(wb, s, 6, 1), true)
  assert.deepEqual(wb.meta(s).tables![0].range, R(0, 0, 6, 4))
  assert.match(updateTable(wb, s, id!, { name: 'A1' }) ?? '', /must start/)
  assert.equal(updateTable(wb, s, id!, { name: 'Sales' }), null)
  assert.match(createTable(wb, s, R(2, 2, 8, 6), { hasHeader: true }).error ?? '', /overlap/)
  convertToRange(wb, s, id!)
  assert.equal(wb.meta(s).tables!.length, 0)
  assert.equal(wb.getStyle(s, 0, 0)?.bold, true) // table header formatting kept
})

test('table without headers gets Column1..n header row', () => {
  const wb = Workbook.blank('T')
  const s = wb.order[0]
  wb.setInputs(s, [{ r: 0, c: 0, input: '1' }, { r: 0, c: 1, input: '2' }, { r: 1, c: 0, input: '3' }])
  createTable(wb, s, R(0, 0, 1, 1), { hasHeader: false })
  assert.equal(wb.getValue(s, 0, 0), 'Column1')
  assert.equal(wb.getValue(s, 1, 0), 1)
  assert.deepEqual(wb.meta(s).tables![0].range, R(0, 0, 2, 1))
})

test('pivot computation: rows, columns, multiple values, filters, date grouping, sort, totals', () => {
  const { wb, s } = sales()
  const src = readPivotSource(wb, s, R(0, 0, 5, 4))
  assert.deepEqual(src.fields, ['Region', 'Product', 'Date', 'Units', 'Revenue'])
  const base = { id: 'p', name: 'P', source: { sheetId: s, range: R(0, 0, 5, 4) }, target: { r: 0, c: 8 }, filters: [], sort: 'asc' as const, grandTotals: true, autoRefresh: false }
  let out = computePivot(src, { ...base, rows: ['Region'], cols: [], values: [{ field: 'Revenue', agg: 'sum' }] })
  assert.deepEqual(out.cells, [['Region', 'Sum of Revenue'], ['East', 70], ['West', 85], ['Grand Total', 155]])
  out = computePivot(src, { ...base, rows: ['Region'], cols: ['Product'], values: [{ field: 'Units', agg: 'sum' }] })
  assert.deepEqual(out.cells[0], ['Sum of Units', 'Books', 'Pens', 'Grand Total'])
  assert.deepEqual(out.cells[2], ['East', 2, 16, 18])
  assert.deepEqual(out.cells.at(-1), ['Grand Total', 7, 20, 27])
  out = computePivot(src, { ...base, rows: ['Date'], cols: [], values: [{ field: 'Revenue', agg: 'sum' }, { field: 'Units', agg: 'count' }], dateGroup: { Date: 'month' }, sort: 'desc' })
  assert.deepEqual(out.cells.slice(0, 2), [['Date', 'Sum of Revenue', 'Count of Units'], ['2024-03 Mar', 90, 2]])
  out = computePivot(src, { ...base, rows: ['Product'], cols: [], values: [{ field: 'Revenue', agg: 'average' }], filters: [{ field: 'Region', values: ['East'] }] })
  assert.deepEqual(out.cells, [['Product', 'Average of Revenue'], ['Books', 30], ['Pens', 20], ['Grand Total', (25 + 30 + 15) / 3]])
})

test('pivot tables: create, refresh after source change, edit, refuse overlap, delete, undo', () => {
  const { wb, s } = sales()
  const res = createPivot(wb, s, { source: { sheetId: s, range: R(0, 0, 5, 4) }, target: { r: 0, c: 7 }, rows: ['Product'], cols: [], values: [{ field: 'Revenue', agg: 'sum' }], filters: [], sort: 'asc', grandTotals: true, autoRefresh: false })
  assert.equal(res.error, undefined)
  assert.equal(wb.getValue(s, 1, 8), 105) // Books
  assert.equal(wb.getStyle(s, 0, 7)?.bold, true)
  wb.setInputs(s, [{ r: 1, c: 4, input: '125' }])
  assert.equal(wb.getValue(s, 2, 8), 50) // Pens not yet refreshed
  refreshPivot(wb, s, res.id!)
  assert.equal(wb.getValue(s, 2, 8), 150)
  updatePivot(wb, s, res.id!, { rows: ['Region', 'Product'] })
  assert.equal(wb.getValue(s, 0, 8), 'Product')
  assert.equal(wb.getValue(s, 1, 7), 'East')
  wb.undo()
  assert.equal(wb.getValue(s, 0, 8), 'Sum of Revenue')
  assert.match(createPivot(wb, s, { source: { sheetId: s, range: R(0, 0, 5, 4) }, target: { r: 2, c: 2 }, rows: ['Product'], cols: [], values: [{ field: 'Units', agg: 'sum' }], filters: [], sort: 'asc', grandTotals: true, autoRefresh: false }).error ?? '', /overlap|data/)
  removePivot(wb, s, res.id!)
  assert.equal(wb.getRaw(s, 1, 8), null)
  assert.equal(wb.meta(s).pivots!.length, 0)
})

test('remove duplicates, trim/case/convert, text to columns', () => {
  const wb = Workbook.blank('T')
  const s = wb.order[0]
  const rows = [['Name', 'City'], ['ann', 'Pune'], ['Bob', 'Delhi'], ['ANN', 'pune'], ['Cy', 'Goa'], ['Bob', 'Delhi']]
  rows.forEach((row, r) => row.forEach((v, c) => wb.setInputs(s, [{ r, c, input: v }])))
  assert.equal(removeDuplicates(wb, s, R(0, 0, 5, 1), [], true), 2)
  assert.deepEqual([1, 2, 3, 4].map((r) => wb.getValue(s, r, 0)), ['ann', 'Bob', 'Cy', null])
  wb.undo()
  assert.equal(removeDuplicates(wb, s, R(0, 0, 5, 1), [1], true), 2)
  wb.undo()
  wb.setInputs(s, [{ r: 8, c: 0, input: '  extra   spaces  ' }, { r: 9, c: 0, input: "'00123" }, { r: 10, c: 0, input: "' 45 " }])
  transformText(wb, s, [R(8, 0, 8, 0)], 'trim')
  assert.equal(wb.getValue(s, 8, 0), 'extra spaces')
  transformText(wb, s, [R(1, 0, 5, 0)], 'proper')
  assert.equal(wb.getValue(s, 3, 0), 'Ann')
  transformText(wb, s, [R(9, 0, 10, 0)], 'toNumber')
  assert.equal(wb.getValue(s, 9, 0), 123)
  assert.equal(wb.getValue(s, 10, 0), 45)
  wb.setInputs(s, [{ r: 12, c: 3, input: 'Ravi,Kumar,Pune' }, { r: 13, c: 3, input: 'Anu,Shah,Goa' }])
  assert.equal(textToColumns(wb, s, R(12, 3, 13, 3), ','), null)
  assert.deepEqual([wb.getValue(s, 12, 3), wb.getValue(s, 12, 4), wb.getValue(s, 13, 5)], ['Ravi', 'Kumar', 'Goa'])
  wb.setInputs(s, [{ r: 12, c: 2, input: 'x,y' }])
  assert.match(textToColumns(wb, s, R(12, 2, 12, 2), ',') ?? '', /already data/)
})

test('flash fill infers patterns from examples', () => {
  const fn = inferFlashFill([{ source: ['Asha Rao'], result: 'Asha' }])!
  assert.equal(fn(['Ben Iyer']), 'Ben')
  const em = inferFlashFill([{ source: ['Asha Rao'], result: 'asha.rao@acme.com' }])!
  assert.equal(em(['Ben Iyer']), 'ben.iyer@acme.com')
  const ini = inferFlashFill([{ source: ['Asha Rao'], result: 'AR' }, { source: ['Ben K Iyer'], result: 'BKI' }])!
  assert.equal(ini(['Cy Dee']), 'CD')
  const wb = Workbook.blank('T')
  const s = wb.order[0]
  ;['Name', 'Asha Rao', 'Ben Iyer', 'Chen Li'].forEach((v, r) => wb.setInputs(s, [{ r, c: 0, input: v }]))
  wb.setInputs(s, [{ r: 0, c: 1, input: 'Last' }, { r: 1, c: 1, input: 'Rao' }])
  const res = flashFill(wb, s, 1, 1, 3)
  assert.equal(res.filled, 2)
  assert.deepEqual([wb.getValue(s, 2, 1), wb.getValue(s, 3, 1)], ['Iyer', 'Li'])
})

test('goal seek finds the input and stays undoable', () => {
  const wb = Workbook.blank('T')
  const s = wb.order[0]
  wb.setInputs(s, [{ r: 0, c: 0, input: '0.05' }, { r: 1, c: 0, input: '100000' }, { r: 2, c: 0, input: '=PMT(A1/12,60,-A2)' }])
  const before = wb.getValue(s, 2, 0) as number
  const res = goalSeek(wb, s, { r: 2, c: 0 }, 2000, { r: 1, c: 0 })
  assert.equal(res.ok, true, res.message)
  assert.ok(Math.abs((wb.getValue(s, 2, 0) as number) - 2000) < 1e-6)
  wb.undo()
  assert.equal(wb.getValue(s, 1, 0), 100000)
  assert.equal(wb.getValue(s, 2, 0), before)
  wb.redo()
  assert.ok(Math.abs((wb.getValue(s, 2, 0) as number) - 2000) < 1e-6)
  assert.equal(goalSeek(wb, s, { r: 0, c: 0 }, 1, { r: 1, c: 0 }).ok, false)
})

test('page layout: pages split by size, fit to width, manual breaks, print area', () => {
  const wb = Workbook.blank('T')
  const s = wb.order[0]
  const entries = []
  for (let r = 0; r < 120; r++) for (let c = 0; c < 12; c++) entries.push({ r, c, input: String(r * c) })
  wb.setInputs(s, entries)
  let lay = layoutPages(wb, s)!
  // A4 portrait: 12 cols * 66pt = 792pt > ~494pt printable -> 2 column pages; rows 16.5pt -> ~40 per page
  assert.equal(new Set(lay.pages.map((p) => p.cols[0])).size, 2)
  assert.equal(lay.pages.length, 2 * Math.ceil(120 / Math.floor(lay.printableH / 16.5)))
  setPageSetup(wb, s, { fitToWidth: true, orientation: 'landscape' })
  lay = layoutPages(wb, s)!
  assert.equal(new Set(lay.pages.map((p) => p.cols[0])).size, 1)
  assert.ok(lay.scale <= 1)
  setPageSetup(wb, s, { rowBreaks: [10], printArea: R(0, 0, 19, 3), fitToWidth: false })
  lay = layoutPages(wb, s)!
  assert.equal(lay.pages.length, 2)
  assert.deepEqual(lay.pages[1].rows[0], 10)
  assert.deepEqual(pageBreaks(lay).rows, [10])
})

test('charts: layout guess, series extraction, histogram, CRUD with undo', () => {
  const { wb, s } = sales()
  const lay = guessChartLayout(wb, s, R(0, 1, 5, 4))
  assert.deepEqual(lay, { seriesIn: 'cols', firstRowHeader: true, firstColLabels: true })
  const ch = newChart(wb, s, { sheetId: s, range: R(0, 3, 5, 4) }, 'column', { x: 10, y: 10 })
  const data = extractChartData(wb, ch)
  assert.equal(data.series.length, 2)
  assert.deepEqual(data.series[1], { name: 'Revenue', values: [25, 10, 30, 75, 15] })
  const byRows = extractChartData(wb, { ...ch, seriesIn: 'rows', firstRowHeader: true, firstColLabels: false, source: { sheetId: s, range: R(0, 3, 2, 4) } })
  assert.deepEqual(byRows.categories, ['Units', 'Revenue'])
  assert.deepEqual(byRows.series.map((x) => x.values), [[10, 25], [4, 10]])
  assert.deepEqual(histogram([1, 2, 2, 3, 10], 3).map((b) => b.count), [4, 0, 1])
  addChart(wb, s, ch)
  updateChart(wb, s, ch.id, { title: 'Units vs revenue', type: 'line' })
  assert.equal(wb.meta(s).charts![0].type, 'line')
  wb.undo()
  assert.equal(wb.meta(s).charts![0].type, 'column')
  removeChart(wb, s, ch.id)
  assert.equal(wb.meta(s).charts!.length, 0)
  wb.undo()
  const wb2 = Workbook.fromDoc(JSON.parse(JSON.stringify(wb.toDoc())))
  assert.equal(wb2.meta(wb2.order[0]).charts![0].title, 'Chart title')
})
