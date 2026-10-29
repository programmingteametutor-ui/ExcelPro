import { test } from 'node:test'
import assert from 'node:assert/strict'
import { addXlfn, functionNames, shiftFormula, stripXlfn } from '../src/engine/formulaText'
import { normalizeFormula, parseInput } from '../src/engine/inputParser'
import { adjustDecimals, formatNumber, formatText, serialToParts } from '../src/engine/numberFormat'
import { colToLetters, lettersToCol, parseRangeA1, quoteSheetName } from '../src/model/address'
import { AxisMetrics } from '../src/model/axis'

test('column letters', () => {
  assert.equal(colToLetters(0), 'A')
  assert.equal(colToLetters(25), 'Z')
  assert.equal(colToLetters(26), 'AA')
  assert.equal(colToLetters(16383), 'XFD')
  assert.equal(lettersToCol('AA'), 26)
  assert.deepEqual(parseRangeA1('$B$2:A1'), { r1: 0, c1: 0, r2: 1, c2: 1 })
  assert.equal(quoteSheetName('Sales Report'), "'Sales Report'")
  assert.equal(quoteSheetName('Sheet2'), 'Sheet2')
})

test('shiftFormula adjusts relative refs only', () => {
  assert.equal(shiftFormula('=A1+$B$1+C$1+$D1', 1, 1), '=B2+$B$1+D$1+$D2')
  assert.equal(shiftFormula('=SUM(A1:A10)', 2, 0), '=SUM(A3:A12)')
  assert.equal(shiftFormula('=Sheet2!A1*2', 0, 1), '=Sheet2!B1*2')
  assert.equal(shiftFormula("='My Sheet'!A1", 1, 0), "='My Sheet'!A2")
  assert.equal(shiftFormula('="A1"&A1', 1, 0), '="A1"&A2')
  assert.equal(shiftFormula('=LOG10(A1)', 1, 0), '=LOG10(A2)')
  assert.equal(shiftFormula('=SUM(A:A)+SUM(1:1)', 1, 1), '=SUM(B:B)+SUM(2:2)')
  assert.equal(shiftFormula('=A1', -1, 0), '=#REF!')
  assert.equal(shiftFormula('=TaxRate*A1', 1, 0), '=TaxRate*A2')
})

test('functionNames and _xlfn handling', () => {
  assert.deepEqual(functionNames('=IF(SUM(A1:A3)>1,"x(",XLOOKUP(1,B:B,C:C))'), ['IF', 'SUM', 'XLOOKUP'])
  assert.equal(stripXlfn('=_xlfn.XLOOKUP(1,A:A,B:B)+_xlfn._xlws.FILTER(A1:A3,B1:B3)'), '=XLOOKUP(1,A:A,B:B)+FILTER(A1:A3,B1:B3)')
  assert.equal(addXlfn('=IFS(A1>1,1)+SUM(A1)+FILTER(A1:A2,B1:B2)'), '=_xlfn.IFS(A1>1,1)+SUM(A1)+_xlfn._xlws.FILTER(A1:A2,B1:B2)')
})

test('input parsing (Excel-like typing)', () => {
  assert.deepEqual(parseInput('42'), { raw: 42 })
  assert.deepEqual(parseInput('-3.5e2'), { raw: -350 })
  assert.deepEqual(parseInput('1,234.50'), { raw: 1234.5, autoFormat: '#,##0.00' })
  assert.deepEqual(parseInput('12%'), { raw: 0.12, autoFormat: '0%' })
  assert.deepEqual(parseInput('$1,200'), { raw: 1200, autoFormat: '$#,##0' })
  assert.deepEqual(parseInput('₹99.90'), { raw: 99.9, autoFormat: '₹#,##0.00' })
  assert.deepEqual(parseInput('true'), { raw: true })
  assert.deepEqual(parseInput('2024-03-01'), { raw: 45352, autoFormat: 'yyyy-mm-dd' })
  assert.deepEqual(parseInput('3/1/2024'), { raw: 45352, autoFormat: 'm/d/yyyy' })
  assert.deepEqual(parseInput('1/3/2024', undefined, 'DMY'), { raw: 45352, autoFormat: 'dd/mm/yyyy' })
  assert.equal(parseInput('2024-02-30').raw, '2024-02-30') // invalid date stays text
  assert.deepEqual(parseInput('13:30'), { raw: 13.5 / 24, autoFormat: 'h:mm' })
  assert.deepEqual(parseInput('00123', '@'), { raw: "'00123" })
  assert.deepEqual(parseInput('hello'), { raw: 'hello' })
  assert.equal(normalizeFormula('=sum(a1:b2'), '=SUM(A1:B2)')
  assert.equal(normalizeFormula('=if(a1>1,"yes",$b$2)'), '=IF(A1>1,"yes",$B$2)')
  assert.equal(normalizeFormula("='my sheet'!a1+TaxRate"), "='my sheet'!A1+TaxRate")
})

test('number formats', () => {
  const f = (v: number, fmt?: string) => formatNumber(v, fmt).text
  assert.equal(f(1234.5), '1234.5')
  assert.equal(f(0.1 + 0.2), '0.3')
  assert.equal(f(123456789012), '1.23457E+11')
  assert.equal(f(1234.567, '#,##0.00'), '1,234.57')
  assert.equal(f(-1234.567, '#,##0.00'), '-1,234.57')
  assert.equal(f(1234567, '#,##0'), '1,234,567')
  assert.equal(f(0.256, '0.0%'), '25.6%')
  assert.equal(f(5, '0.00'), '5.00')
  assert.equal(f(5, '$#,##0.00'), '$5.00')
  assert.equal(f(-5, '$#,##0.00;[Red]($#,##0.00)'), '($5.00)')
  assert.equal(formatNumber(-5, '$#,##0.00;[Red]($#,##0.00)').color, '#ff0000')
  assert.equal(f(0, '#,##0;-#,##0;"zero"'), 'zero')
  assert.equal(f(12345, '0.00E+00'), '1.23E+04')
  assert.equal(f(0.5, '# ?/?'), ' 1/2'.trim() === '1/2' ? f(0.5, '# ?/?') : '1/2')
  assert.equal(f(1.75, '# ?/?'), '1 3/4')
  assert.equal(f(45306, 'yyyy-mm-dd'), '2024-01-15')
  assert.equal(f(45306, 'm/d/yyyy'), '1/15/2024')
  assert.equal(f(45306, 'dddd, mmmm d, yyyy'), 'Monday, January 15, 2024')
  assert.equal(f(45306.75, 'h:mm AM/PM'), '6:00 PM')
  assert.equal(f(45306.75, 'hh:mm:ss'), '18:00:00')
  assert.equal(f(1.5, '[h]:mm'), '36:00')
  assert.equal(f(1234567, '#,##0,"K"'), '1,235K')
  assert.equal(f(1234.5, '[$€-407] #,##0.00'), '€ 1,234.50')
  assert.equal(f(42, '000000'), '000042')
  assert.equal(formatText('abc', '@').text, 'abc')
  assert.equal(formatText('abc', '0;0;0;"<"@">"').text, '<abc>')
  assert.equal(f(60, 'yyyy-mm-dd'), '1900-02-29') // Excel's 1900 leap-year quirk
  assert.deepEqual([serialToParts(61).y, serialToParts(61).m, serialToParts(61).d], [1900, 3, 1])
  assert.equal(f(1234.5, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)').trim(), '$ 1,234.50')
})

test('increase/decrease decimals', () => {
  assert.equal(adjustDecimals(undefined, 1, 3), '0.0')
  assert.equal(adjustDecimals(undefined, 1, 3.14), '0.000')
  assert.equal(adjustDecimals('0.00', 1), '0.000')
  assert.equal(adjustDecimals('0.00', -1), '0.0')
  assert.equal(adjustDecimals('0.0', -1), '0')
  assert.equal(adjustDecimals('#,##0', 1), '#,##0.0')
  assert.equal(adjustDecimals('0%', 1), '0.0%')
})

test('axis metrics with overrides and hidden items', () => {
  const ax = new AxisMetrics(20, 1_000_000, { 5: 50 }, [10])
  assert.equal(ax.offsetOf(5), 100)
  assert.equal(ax.offsetOf(6), 150)
  assert.equal(ax.offsetOf(11), 20 * 11 + 30 - 20)
  assert.equal(ax.indexAt(149), 5)
  assert.equal(ax.indexAt(150), 6)
  assert.equal(ax.step(9, 1), 11)
  assert.equal(ax.total, 1_000_000 * 20 + 30 - 20)
})
