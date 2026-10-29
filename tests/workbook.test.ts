import { test } from 'node:test'
import assert from 'node:assert/strict'
import { detectSeries, parseTsv, Workbook } from '../src/engine/workbook'

const R = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 })

function setup() {
  const wb = Workbook.blank('T')
  const s = wb.sheets[0].id
  return { wb, s }
}

test('edit, undo, redo', () => {
  const { wb, s } = setup()
  wb.setInputs(s, [{ r: 0, c: 0, input: '5' }])
  wb.setInputs(s, [{ r: 0, c: 1, input: '=A1*2' }])
  assert.equal(wb.getValue(s, 0, 1), 10)
  wb.setInputs(s, [{ r: 0, c: 0, input: '7' }])
  assert.equal(wb.getValue(s, 0, 1), 14)
  wb.undo()
  assert.equal(wb.getValue(s, 0, 1), 10)
  wb.undo()
  assert.equal(wb.getRaw(s, 0, 1), null)
  wb.redo()
  assert.equal(wb.getValue(s, 0, 1), 10)
})

test('auto number formats are applied on entry and undone together', () => {
  const { wb, s } = setup()
  wb.setInputs(s, [{ r: 0, c: 0, input: '2024-05-06' }])
  assert.equal(wb.getStyle(s, 0, 0)?.numFmt, 'yyyy-mm-dd')
  assert.equal(wb.getEditText(s, 0, 0), '2024-05-06')
  wb.undo()
  assert.equal(wb.getStyle(s, 0, 0), undefined)
  assert.equal(wb.getRaw(s, 0, 0), null)
  wb.setInputs(s, [{ r: 1, c: 0, input: '15%' }])
  assert.equal(wb.getDisplay(s, 1, 0).text, '15%')
  assert.equal(wb.getEditText(s, 1, 0), '15%')
})

test('text that looks like a number round-trips through the editor', () => {
  const { wb, s } = setup()
  wb.setInputs(s, [{ r: 0, c: 0, input: "'00123" }])
  assert.equal(wb.getValue(s, 0, 0), '00123')
  assert.equal(wb.getEditText(s, 0, 0), "'00123")
})

test('formatting with undo', () => {
  const { wb, s } = setup()
  wb.applyStyle(s, [R(0, 0, 1, 1)], { bold: true, fill: '#ffff00' })
  assert.equal(wb.getStyle(s, 1, 1)?.bold, true)
  wb.applyStyle(s, [R(0, 0)], (st) => ({ ...st, bold: !st?.bold }))
  assert.equal(wb.getStyle(s, 0, 0)?.bold, undefined)
  wb.undo()
  assert.equal(wb.getStyle(s, 0, 0)?.bold, true)
  wb.clearFormats(s, [R(0, 0, 5, 5)])
  assert.equal(wb.getStyle(s, 1, 1), undefined)
  wb.undo()
  assert.equal(wb.getStyle(s, 1, 1)?.fill, '#ffff00')
})

test('insert/delete rows and columns shift formulas, styles and merges; undo restores', () => {
  const { wb, s } = setup()
  wb.setInputs(s, [
    { r: 0, c: 0, input: '1' },
    { r: 1, c: 0, input: '2' },
    { r: 2, c: 0, input: '=SUM(A1:A2)' },
  ])
  wb.applyStyle(s, [R(1, 0)], { bold: true })
  wb.merge(s, R(1, 1, 1, 2))
  assert.equal(wb.insertRows(s, 1, 2), null)
  assert.equal(wb.getRaw(s, 4, 0), '=SUM(A1:A4)')
  assert.equal(wb.getStyle(s, 3, 0)?.bold, true)
  assert.deepEqual(wb.meta(s).merges, [R(3, 1, 3, 2)])
  wb.undo()
  assert.equal(wb.getRaw(s, 2, 0), '=SUM(A1:A2)')
  assert.equal(wb.getStyle(s, 1, 0)?.bold, true)
  assert.deepEqual(wb.meta(s).merges, [R(1, 1, 1, 2)])
  assert.equal(wb.deleteRows(s, 0, 1), null)
  assert.equal(wb.getRaw(s, 1, 0), '=SUM(A1:A1)')
  assert.equal(wb.getValue(s, 1, 0), 2)
  wb.undo()
  assert.equal(wb.insertCols(s, 0, 1), null)
  assert.equal(wb.getRaw(s, 2, 1), '=SUM(B1:B2)')
  assert.equal(wb.getStyle(s, 1, 1)?.bold, true)
  wb.undo()
  wb.deleteRows(s, 0, 2)
  assert.equal(wb.getDisplay(s, 0, 0).text, '#REF!')
})

test('worksheet add, rename, duplicate, move, hide, delete with undo', () => {
  const { wb, s } = setup()
  const s2 = wb.addSheet()
  assert.equal(wb.meta(s2).name, 'Sheet2')
  wb.setInputs(s2, [{ r: 0, c: 0, input: '9' }])
  wb.setInputs(s, [{ r: 0, c: 0, input: '=Sheet2!A1+1' }])
  assert.equal(wb.getValue(s, 0, 0), 10)
  const dup = wb.duplicateSheet(s2)
  assert.equal(wb.meta(dup).name, 'Sheet2 (2)')
  assert.equal(wb.getValue(dup, 0, 0), 9)
  wb.moveSheet(dup, 0)
  assert.equal(wb.order[0], dup)
  wb.undo()
  assert.equal(wb.order[2], dup)
  assert.match(wb.renameSheet(s2, 'sheet1') ?? '', /already exists/)
  assert.match(wb.renameSheet(s2, 'Bad:Name') ?? '', /cannot contain/)
  assert.equal(wb.removeSheet(s2), null)
  assert.equal(wb.getDisplay(s, 0, 0).text, '#REF!')
  wb.undo()
  assert.equal(wb.getValue(s, 0, 0), 10)
  assert.ok(wb.order.includes(s2))
  assert.equal(wb.setSheetHidden(s2, true), null)
  assert.equal(wb.meta(s2).hidden, true)
  assert.equal(wb.removeSheet(dup), null)
  assert.match(wb.removeSheet(s) ?? '', /at least one visible/)
})

test('internal copy/paste adjusts relative references and tiles', () => {
  const { wb, s } = setup()
  wb.setInputs(s, [
    { r: 0, c: 0, input: '1' },
    { r: 1, c: 0, input: '2' },
    { r: 0, c: 1, input: '=A1*10' },
  ])
  wb.applyStyle(s, [R(0, 1)], { italic: true })
  const clip = wb.copy(s, R(0, 1))
  wb.paste(clip, s, R(1, 1))
  assert.equal(wb.getRaw(s, 1, 1), '=A2*10')
  assert.equal(wb.getValue(s, 1, 1), 20)
  assert.equal(wb.getStyle(s, 1, 1)?.italic, true)
  const dest = wb.paste(clip, s, R(2, 3, 5, 3))
  assert.deepEqual(dest, R(2, 3, 5, 3))
  assert.equal(wb.getRaw(s, 5, 3), '=C6*10')
  wb.paste(wb.copy(s, R(1, 1)), s, R(0, 5), 'values')
  assert.equal(wb.getRaw(s, 0, 5), 20)
})

test('cut/paste moves cells and keeps references pointing at them', () => {
  const { wb, s } = setup()
  wb.setInputs(s, [
    { r: 0, c: 0, input: '3' },
    { r: 0, c: 1, input: '=A1+1' },
  ])
  wb.applyStyle(s, [R(0, 0)], { bold: true })
  const clip = wb.copy(s, R(0, 0), true)
  wb.paste(clip, s, R(4, 4))
  assert.equal(wb.getRaw(s, 0, 1), '=E5+1')
  assert.equal(wb.getValue(s, 0, 1), 4)
  assert.equal(wb.getStyle(s, 4, 4)?.bold, true)
  assert.equal(wb.getStyle(s, 0, 0), undefined)
  wb.undo()
  assert.equal(wb.getRaw(s, 0, 1), '=A1+1')
  assert.equal(wb.getStyle(s, 0, 0)?.bold, true)
})

test('external paste (TSV from Excel) parses values', () => {
  const { wb, s } = setup()
  const text = 'Name\tAmount\tDate\r\nAlice\t1,200.50\t2024-01-02\r\n"Bob ""B"""\t15%\t\r\n'
  assert.deepEqual(parseTsv(text)[2], ['Bob "B"', '15%', ''])
  wb.pasteText(s, { r: 0, c: 0 }, text)
  assert.equal(wb.getValue(s, 1, 1), 1200.5)
  assert.equal(wb.getDisplay(s, 1, 1).text, '1,200.50')
  assert.equal(wb.getDisplay(s, 1, 2).text, '2024-01-02')
  assert.equal(wb.getValue(s, 2, 0), 'Bob "B"')
  assert.equal(wb.getValue(s, 2, 1), 0.15)
})

test('fill handle: series, text numbers, formulas, months', () => {
  const { wb, s } = setup()
  wb.setInputs(s, [
    { r: 0, c: 0, input: '1' },
    { r: 1, c: 0, input: '3' },
    { r: 0, c: 1, input: 'Item 1' },
    { r: 0, c: 2, input: '=A1*2' },
    { r: 0, c: 3, input: 'Jan' },
    { r: 0, c: 4, input: 'x' },
  ])
  wb.fill(s, R(0, 0, 1, 0), R(0, 0, 4, 0))
  assert.deepEqual([2, 3, 4].map((r) => wb.getValue(s, r, 0)), [5, 7, 9])
  wb.fill(s, R(0, 1), R(0, 1, 2, 1))
  assert.equal(wb.getValue(s, 2, 1), 'Item 3')
  wb.fill(s, R(0, 2), R(0, 2, 3, 2))
  assert.equal(wb.getRaw(s, 3, 2), '=A4*2')
  assert.equal(wb.getValue(s, 3, 2), 14)
  wb.fill(s, R(0, 3), R(0, 3, 0, 3 + 0))
  wb.fill(s, R(0, 3), R(0, 3, 13, 3))
  assert.equal(wb.getValue(s, 12, 3), 'Jan')
  wb.fill(s, R(0, 4), R(0, 4, 2, 4))
  assert.equal(wb.getValue(s, 2, 4), 'x')
  wb.undo()
  assert.equal(wb.getRaw(s, 2, 4), null)
  assert.equal(detectSeries([45000], ['yyyy-mm-dd'])!(3), 45003)
})

test('merge clears covered cells; unmerge; merge & center', () => {
  const { wb, s } = setup()
  wb.setInputs(s, [
    { r: 0, c: 0, input: 'keep' },
    { r: 0, c: 1, input: 'drop' },
  ])
  wb.merge(s, R(0, 0, 0, 2), true)
  assert.equal(wb.getRaw(s, 0, 1), null)
  assert.equal(wb.getStyle(s, 0, 0)?.hAlign, 'center')
  assert.deepEqual(wb.mergeAt(s, 0, 2), R(0, 0, 0, 2))
  wb.unmerge(s, R(0, 1))
  assert.equal(wb.mergeAt(s, 0, 2), undefined)
  wb.undo()
  wb.undo()
  assert.equal(wb.getRaw(s, 0, 1), 'drop')
})

test('find and replace across sheets', () => {
  const { wb, s } = setup()
  const s2 = wb.addSheet()
  wb.setInputs(s, [{ r: 0, c: 0, input: 'apple pie' }, { r: 3, c: 2, input: '=UPPER("apple")' }])
  wb.setInputs(s2, [{ r: 1, c: 1, input: 'Apple' }])
  assert.equal(wb.find('apple', { sheetIds: [s], lookIn: 'values' }).length, 2)
  assert.equal(wb.find('apple', { sheetIds: [s, s2], lookIn: 'values', matchCase: true }).length, 1)
  assert.equal(wb.find('apple', { sheetIds: [s, s2], lookIn: 'values', entireCell: true }).length, 2)
  const m = wb.find('apple', { sheetIds: [s, s2], lookIn: 'formulas' })
  assert.equal(wb.replace(m, 'apple', 'pear'), 3)
  assert.equal(wb.getValue(s, 0, 0), 'pear pie')
  assert.equal(wb.getRaw(s, 3, 2), '=UPPER("pear")')
  assert.equal(wb.getValue(s2, 1, 1), 'pear')
})

test('Ctrl+Arrow data-region navigation', () => {
  const { wb, s } = setup()
  for (let r = 0; r < 5; r++) wb.setInputs(s, [{ r, c: 0, input: String(r) }])
  wb.setInputs(s, [{ r: 10, c: 0, input: 'x' }])
  assert.deepEqual(wb.dataEdge(s, { r: 0, c: 0 }, 1, 0, 1000, 100), { r: 4, c: 0 })
  assert.deepEqual(wb.dataEdge(s, { r: 4, c: 0 }, 1, 0, 1000, 100), { r: 10, c: 0 })
  assert.deepEqual(wb.dataEdge(s, { r: 10, c: 0 }, 1, 0, 1000, 100), { r: 999, c: 0 })
  assert.deepEqual(wb.dataEdge(s, { r: 10, c: 0 }, -1, 0, 1000, 100), { r: 4, c: 0 })
})

test('serialize -> fromDoc round trip preserves contents, styles, layout, names', () => {
  const { wb, s } = setup()
  const s2 = wb.addSheet('Data Sheet')
  wb.setInputs(s2, [{ r: 0, c: 0, input: '4' }])
  wb.setInputs(s, [{ r: 0, c: 0, input: "='Data Sheet'!A1*2" }, { r: 1, c: 0, input: '2024-02-03' }])
  wb.applyStyle(s, [{ r1: 0, c1: 0, r2: 0, c2: 0 }], { bold: true, color: '#ff0000' })
  wb.setColWidths(s, [0], 140)
  wb.setRowHeights(s, [1], 40)
  wb.merge(s, R(3, 0, 3, 3))
  wb.setTabColor(s2, '#00aa00')
  assert.equal(wb.addNamedRange('Rate', "='Data Sheet'!$A$1"), null)
  wb.setInputs(s, [{ r: 5, c: 0, input: '=Rate*10' }])
  const doc = JSON.parse(JSON.stringify(wb.toDoc()))
  const wb2 = Workbook.fromDoc(doc)
  const t = wb2.sheets[0].id
  assert.equal(wb2.getValue(t, 0, 0), 8)
  assert.equal(wb2.getDisplay(t, 1, 0).text, '2024-02-03')
  assert.equal(wb2.getStyle(t, 0, 0)?.color, '#ff0000')
  assert.equal(wb2.meta(t).colWidths[0], 140)
  assert.equal(wb2.meta(t).rowHeights[1], 40)
  assert.deepEqual(wb2.meta(t).merges, [R(3, 0, 3, 3)])
  assert.equal(wb2.sheets[1].tabColor, '#00aa00')
  assert.equal(wb2.getValue(t, 5, 0), 40)
  assert.equal(wb2.canUndo(), false)
})
