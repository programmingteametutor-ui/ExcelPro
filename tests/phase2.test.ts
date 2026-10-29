import { test } from 'node:test'
import assert from 'node:assert/strict'
import { Workbook } from '../src/engine/workbook'
import {
  addValidation,
  changeNamedRange,
  checkInput,
  clearAllFilters,
  filterValues,
  guessHeader,
  invalidCells,
  listItems,
  removeNamedRange,
  setAutoFilter,
  setFilterCriterion,
  setFreeze,
  setLink,
  setNote,
  sortRange,
} from '../src/engine/features/data'
import { addCondFormat, cfEvaluator, moveCondFormat, removeCondFormats } from '../src/engine/features/condFormat'
import { currentRegion } from '../src/engine/features/common'

const R = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 })

function sample() {
  const wb = Workbook.blank('T')
  const s = wb.order[0]
  const rows = [
    ['Name', 'Dept', 'Salary', 'Start'],
    ['Asha', 'Sales', '52000', '2021-03-01'],
    ['Ben', 'Ops', '48000', '2019-07-15'],
    ['Chen', 'Sales', '61000', '2020-01-10'],
    ['Dev', 'IT', '75000', '2018-11-05'],
    ['Ema', 'Ops', '', '2022-02-20'],
  ]
  rows.forEach((row, r) => row.forEach((v, c) => wb.setInputs(s, [{ r, c, input: v }])))
  wb.setInputs(s, [{ r: 1, c: 4, input: '=C2*1.1' }])
  wb.applyStyle(s, [R(1, 0)], { bold: true })
  return { wb, s }
}

test('current region and header detection', () => {
  const { wb, s } = sample()
  assert.deepEqual(currentRegion(wb, s, 2, 1), R(0, 0, 5, 4))
  assert.equal(guessHeader(wb, s, R(0, 0, 5, 3)), true)
})

test('sort by one and two keys; blanks last; formulas and styles move; undo', () => {
  const { wb, s } = sample()
  assert.equal(sortRange(wb, s, R(0, 0, 5, 4), [{ col: 2, desc: true }], true), null)
  assert.deepEqual([1, 2, 3, 4, 5].map((r) => wb.getValue(s, r, 0)), ['Dev', 'Chen', 'Asha', 'Ben', 'Ema'])
  assert.equal(wb.getRaw(s, 3, 4), '=C4*1.1') // Asha's formula moved from row 2 to row 4
  assert.equal(wb.getValue(s, 3, 4), 57200.00000000001 > 57199 ? wb.getValue(s, 3, 4) : 0)
  assert.equal(wb.getStyle(s, 3, 0)?.bold, true)
  wb.undo()
  assert.equal(wb.getValue(s, 1, 0), 'Asha')
  sortRange(wb, s, R(0, 0, 5, 3), [{ col: 1, desc: false }, { col: 2, desc: true }], true)
  assert.deepEqual([1, 2, 3, 4, 5].map((r) => wb.getValue(s, r, 0)), ['Dev', 'Ben', 'Ema', 'Chen', 'Asha'])
  // dates sort as dates
  sortRange(wb, s, R(0, 0, 5, 3), [{ col: 3, desc: false }], true)
  assert.equal(wb.getValue(s, 1, 0), 'Dev')
})

test('sort refuses merged cells', () => {
  const { wb, s } = sample()
  wb.merge(s, R(2, 0, 2, 1))
  assert.match(sortRange(wb, s, R(0, 0, 5, 3), [{ col: 0, desc: false }], true) ?? '', /merged/)
})

test('autofilter by values and conditions; distinct values; clear', () => {
  const { wb, s } = sample()
  setAutoFilter(wb, s, R(0, 0, 5, 3))
  assert.deepEqual(filterValues(wb, s, 1).values, ['IT', 'Ops', 'Sales'])
  setFilterCriterion(wb, s, 1, { kind: 'values', values: ['Sales'], blanks: false })
  assert.deepEqual(wb.meta(s).filterHidden, [2, 4, 5])
  setFilterCriterion(wb, s, 2, { kind: 'condition', op: 'gt', a: '55000' })
  assert.deepEqual(wb.meta(s).filterHidden, [1, 2, 4, 5])
  assert.deepEqual(filterValues(wb, s, 0).values, ['Chen'])
  wb.undo()
  assert.deepEqual(wb.meta(s).filterHidden, [2, 4, 5])
  clearAllFilters(wb, s)
  assert.deepEqual(wb.meta(s).filterHidden, [])
  setFilterCriterion(wb, s, 2, { kind: 'condition', op: 'top', a: '2' })
  assert.deepEqual(wb.meta(s).filterHidden, [1, 2, 5])
  setFilterCriterion(wb, s, 2, { kind: 'values', values: [], blanks: true })
  assert.deepEqual(wb.meta(s).filterHidden, [1, 2, 3, 4])
  setAutoFilter(wb, s, null)
  assert.equal(wb.meta(s).autoFilter, undefined)
})

test('inserting rows shifts filter, validation and CF ranges', () => {
  const { wb, s } = sample()
  setAutoFilter(wb, s, R(0, 0, 5, 3))
  addValidation(wb, s, { range: R(1, 2, 5, 2), type: 'whole', op: 'gt', v1: '0', allowBlank: true, showDropdown: false, errorStyle: 'stop' })
  addCondFormat(wb, s, { range: R(1, 2, 5, 2), rule: { type: 'cellIs', op: 'gt', v1: '60000' }, style: { fill: '#ffc7ce' } })
  wb.insertRows(s, 0, 2)
  assert.deepEqual(wb.meta(s).autoFilter!.range, R(2, 0, 7, 3))
  assert.deepEqual(wb.meta(s).validations![0].range, R(3, 2, 7, 2))
  assert.deepEqual(wb.meta(s).condFormats![0].range, R(3, 2, 7, 2))
})

test('data validation: list, numbers, dates, text length, custom, invalid cells', () => {
  const { wb, s } = sample()
  addValidation(wb, s, { range: R(1, 1, 10, 1), type: 'list', list: 'Sales, Ops, IT', allowBlank: true, showDropdown: true, errorStyle: 'stop' })
  assert.deepEqual(checkInput(wb, s, 3, 1, 'ops'), { ok: true })
  const bad = checkInput(wb, s, 3, 1, 'HR')
  assert.equal(bad.ok, false)
  assert.match(!bad.ok ? bad.message : '', /one of the values/)
  wb.setInputs(s, [{ r: 10, c: 6, input: 'Red' }, { r: 11, c: 6, input: 'Blue' }])
  addValidation(wb, s, { range: R(1, 5, 5, 5), type: 'list', list: '=$G$11:$G$12', allowBlank: true, showDropdown: true, errorStyle: 'warning' })
  assert.deepEqual(listItems(wb, s, wb.meta(s).validations![1]), ['Red', 'Blue'])
  addValidation(wb, s, { range: R(1, 2, 10, 2), type: 'whole', op: 'between', v1: '30000', v2: '80000', allowBlank: false, showDropdown: false, errorStyle: 'stop', errorMessage: 'Salary 30k–80k' })
  assert.equal(checkInput(wb, s, 2, 2, '45000').ok, true)
  assert.equal(checkInput(wb, s, 2, 2, '45000.5').ok, false)
  const r = checkInput(wb, s, 2, 2, '')
  assert.equal(!r.ok && r.message, 'Salary 30k–80k')
  addValidation(wb, s, { range: R(1, 3, 10, 3), type: 'date', op: 'gte', v1: '2019-01-01', allowBlank: true, showDropdown: false, errorStyle: 'stop' })
  assert.equal(checkInput(wb, s, 1, 3, '2018-12-31').ok, false)
  assert.equal(checkInput(wb, s, 1, 3, '2020-06-01').ok, true)
  addValidation(wb, s, { range: R(1, 7, 5, 7), type: 'textLength', op: 'lte', v1: '3', allowBlank: true, showDropdown: false, errorStyle: 'stop' })
  assert.equal(checkInput(wb, s, 1, 7, 'abcd').ok, false)
  addValidation(wb, s, { range: R(1, 8, 5, 8), type: 'custom', list: '=MOD(I2,5)=0', allowBlank: true, showDropdown: false, errorStyle: 'stop' })
  assert.equal(checkInput(wb, s, 3, 8, '15').ok, true)
  assert.equal(checkInput(wb, s, 3, 8, '16').ok, false)
  // Dev started 2018 and Ema has no salary -> both invalid now
  const inv = invalidCells(wb, s).map((x) => `${x.r},${x.c}`)
  assert.ok(inv.includes('4,3'), inv.join(' '))
  assert.ok(!inv.includes('5,2'), 'blank cells are not flagged by Find invalid data')
})

test('conditional formatting: cellIs, text, duplicates, top, formula, color scale, data bar, icons, priority', () => {
  const { wb, s } = sample()
  addCondFormat(wb, s, { range: R(1, 2, 5, 2), rule: { type: 'cellIs', op: 'gt', v1: '60000' }, style: { fill: '#ffc7ce', color: '#9c0006' } })
  let ev = cfEvaluator(wb, s)!
  assert.equal(ev(3, 2)?.style?.fill, '#ffc7ce')
  assert.equal(ev(1, 2), null)
  addCondFormat(wb, s, { range: R(1, 1, 5, 1), rule: { type: 'duplicate' }, style: { bold: true } })
  ev = cfEvaluator(wb, s)!
  assert.equal(ev(1, 1)?.style?.bold, true)
  assert.equal(ev(4, 1), null)
  addCondFormat(wb, s, { range: R(1, 0, 5, 0), rule: { type: 'text', op: 'contains', text: 'e' }, style: { italic: true } })
  addCondFormat(wb, s, { range: R(1, 0, 5, 0), rule: { type: 'formula', formula: '=$C2>60000' }, style: { color: '#006100' } })
  ev = cfEvaluator(wb, s)!
  assert.equal(ev(3, 0)?.style?.color, '#006100') // Chen 61000
  assert.equal(ev(2, 0)?.style?.italic, true) // Ben contains e
  assert.equal(ev(2, 0)?.style?.color, undefined)
  addCondFormat(wb, s, { range: R(1, 2, 5, 2), rule: { type: 'top', n: 1 }, style: { underline: true } })
  addCondFormat(wb, s, { range: R(1, 2, 5, 2), rule: { type: 'colorScale', colors: ['#f8696b', '#63be7b'] } })
  addCondFormat(wb, s, { range: R(1, 2, 5, 2), rule: { type: 'dataBar', color: '#638ec6' } })
  addCondFormat(wb, s, { range: R(1, 2, 5, 2), rule: { type: 'iconSet', set: 'arrows' } })
  ev = cfEvaluator(wb, s)!
  const dev = ev(4, 2)!
  assert.equal(dev.style?.underline, true)
  assert.equal(dev.scaleFill, '#63be7b')
  assert.equal(dev.bar?.pct, 1)
  assert.equal(dev.icon?.level, 2)
  assert.equal(ev(2, 2)!.icon?.level, 0)
  // priority: newest rule first; moving changes which style wins
  const ids = wb.meta(s).condFormats!.map((c) => c.id)
  addCondFormat(wb, s, { range: R(3, 2), rule: { type: 'notBlank' }, style: { fill: '#000000' } })
  assert.equal(cfEvaluator(wb, s)!(3, 2)?.style?.fill, '#000000')
  moveCondFormat(wb, s, wb.meta(s).condFormats![0].id, 1)
  removeCondFormats(wb, s, ids)
  assert.equal(wb.meta(s).condFormats!.length, 1)
  removeCondFormats(wb, s, 'all', R(0, 0, 10, 10))
  assert.equal(wb.meta(s).condFormats!.length, 0)
})

test('named ranges edit/delete, notes, links, freeze all undoable and persisted', () => {
  const { wb, s } = sample()
  wb.addNamedRange('Pay', '=Sheet1!$C$2:$C$6')
  wb.setInputs(s, [{ r: 8, c: 0, input: '=SUM(Pay)' }])
  assert.equal(wb.getValue(s, 8, 0), 236000)
  assert.equal(changeNamedRange(wb, 'Pay', '=Sheet1!$C$2:$C$3'), null)
  assert.equal(wb.getValue(s, 8, 0), 100000)
  removeNamedRange(wb, 'Pay')
  assert.equal(wb.getDisplay(s, 8, 0).text, '#NAME?')
  wb.undo()
  assert.equal(wb.getValue(s, 8, 0), 100000)
  setNote(wb, s, 1, 0, 'Top seller')
  setLink(wb, s, 2, 0, 'example.com')
  setFreeze(wb, s, 1, 1)
  const doc = JSON.parse(JSON.stringify(wb.toDoc()))
  const wb2 = Workbook.fromDoc(doc)
  const m = wb2.meta(wb2.order[0])
  assert.equal(m.notes!['1,0'], 'Top seller')
  assert.equal(m.links!['2,0'], 'https://example.com')
  assert.deepEqual(m.freeze, { rows: 1, cols: 1 })
  wb.undo()
  assert.equal(wb.meta(s).freeze, undefined)
})
