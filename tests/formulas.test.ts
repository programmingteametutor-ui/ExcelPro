import { test } from 'node:test'
import assert from 'node:assert/strict'
import { Workbook } from '../src/engine/workbook'

function wbWith(rows: (string | number | boolean | null)[][], extra?: Record<string, (string | number | null)[][]>) {
  const wb = Workbook.blank('T', ['Sheet1', ...Object.keys(extra ?? {})])
  const s1 = wb.sheets[0].id
  rows.forEach((row, r) => row.forEach((v, c) => v !== null && wb.setInputs(s1, [{ r, c, input: String(v) }])))
  for (const [name, data] of Object.entries(extra ?? {})) {
    const sid = wb.sheetByName(name)!.id
    data.forEach((row, r) => row.forEach((v, c) => v !== null && wb.setInputs(sid, [{ r, c, input: String(v) }])))
  }
  return { wb, s1 }
}

const val = (wb: Workbook, sid: string, a1: string) => {
  const c = a1.charCodeAt(0) - 65
  const r = parseInt(a1.slice(1), 10) - 1
  return wb.getValue(sid, r, c)
}

test('spec formula list (§24)', () => {
  const { wb, s1 } = wbWith(
    [
      [10, 1, 'Sales', 100, 'K2'],
      [20, 2, 'Ops', 200, 'K3'],
      [30, 3, 'Sales', 300],
      [40, 4, 'Pass', 400],
      [150, 5, 'Pass', 500],
      [null, 6],
      [null, 7],
      [null, 8],
      [null, 9],
      [null, 10],
    ],
    { Sheet2: [[5]] },
  )
  const put = (a1: string, f: string) => {
    const c = a1.charCodeAt(0) - 65
    const r = parseInt(a1.slice(1), 10) - 1
    wb.setInputs(s1, [{ r, c, input: f }])
    return wb.getValue(s1, r, c)
  }
  assert.equal(put('H1', '=SUM(A1:A5)'), 250)
  assert.equal(put('H2', '=AVERAGE(B1:B10)'), 5.5)
  assert.equal(put('H3', '=IF(A1>100,"High","Low")'), 'Low')
  assert.equal(put('H4', '=IF(A5>100,"High","Low")'), 'High')
  assert.equal(put('H5', '=COUNTIF(C1:C20,"Pass")'), 2)
  assert.equal(put('H6', '=SUMIF(C1:C20,"Sales",D1:D20)'), 400)
  assert.equal(put('H7', '=IFERROR(A1/F1,0)'), 0)
  assert.equal(put('H8', '=Sheet2!A1+10'), 15)
  // lookups
  wb.setInputs(s1, [
    { r: 0, c: 9, input: 'K1' }, { r: 0, c: 10, input: 'Apple' }, { r: 0, c: 11, input: '1.5' },
    { r: 1, c: 9, input: 'K2' }, { r: 1, c: 10, input: 'Pear' }, { r: 1, c: 11, input: '2.25' },
    { r: 2, c: 9, input: 'K3' }, { r: 2, c: 10, input: 'Fig' }, { r: 2, c: 11, input: '3' },
  ])
  assert.equal(put('H9', '=VLOOKUP(E1,J1:L20,3,FALSE)'), 2.25)
  assert.equal(put('H10', '=INDEX(K1:K20,MATCH(E2,J1:J20,0))'), 'Fig')
  assert.equal(put('H11', '=XLOOKUP("K3",J1:J3,K1:K3)'), 'Fig')
})

test('absolute, relative and mixed references; precedence and concatenation', () => {
  const { wb, s1 } = wbWith([[2, 3], [4, 5]])
  wb.setInputs(s1, [{ r: 0, c: 2, input: '=$A$1*B1+A$2-$B2' }])
  assert.equal(val(wb, s1, 'C1'), 2 * 3 + 4 - 5)
  wb.setInputs(s1, [{ r: 0, c: 3, input: '=2+3*4^2' }])
  assert.equal(val(wb, s1, 'D1'), 50)
  wb.setInputs(s1, [{ r: 0, c: 4, input: '="Total: "&A1+B1' }])
  assert.equal(val(wb, s1, 'E1'), 'Total: 5')
  wb.setInputs(s1, [{ r: 0, c: 5, input: '=A1>=2' }])
  assert.equal(val(wb, s1, 'F1'), true)
})

test('dependent formulas recalculate after edits', () => {
  const { wb, s1 } = wbWith([[1], [2], ['=A1+A2'], ['=A3*10']])
  assert.equal(val(wb, s1, 'A4'), 30)
  wb.setInputs(s1, [{ r: 0, c: 0, input: '5' }])
  assert.equal(val(wb, s1, 'A3'), 7)
  assert.equal(val(wb, s1, 'A4'), 70)
})

test('cross-sheet references with quoted names survive renames', () => {
  const { wb, s1 } = wbWith([["='Sales Report'!B5*2"]], { 'Sales Report': [[], [], [], [], [null, 21]] })
  assert.equal(val(wb, s1, 'A1'), 42)
  const sr = wb.sheetByName('Sales Report')!.id
  assert.equal(wb.renameSheet(sr, 'Q1 Sales'), null)
  assert.equal(wb.getRaw(s1, 0, 0), "='Q1 Sales'!B5*2")
  assert.equal(val(wb, s1, 'A1'), 42)
  wb.undo()
  assert.equal(wb.getRaw(s1, 0, 0), "='Sales Report'!B5*2")
})

test('error values', () => {
  const { wb, s1 } = wbWith([[1, 0, '=A1/B1', '=NOSUCHFN(1)', '=A1+"x"', '=C5', '=VLOOKUP(99,A1:B1,2,FALSE)', '=SQRT(-1)']])
  const err = (c: number) => wb.getDisplay(s1, 0, c).text
  assert.equal(err(2), '#DIV/0!')
  assert.equal(err(3), '#NAME?')
  assert.match(wb.explainError(s1, 0, 3)!, /NOSUCHFN is not supported/)
  assert.equal(err(4), '#VALUE!')
  assert.equal(wb.getDisplay(s1, 0, 5).text, '0') // empty cell reference -> 0 like Excel
  assert.equal(err(6), '#N/A')
  assert.equal(err(7), '#NUM!')
})

test('circular references are detected', () => {
  const { wb, s1 } = wbWith([['=B1'], []])
  wb.setInputs(s1, [{ r: 0, c: 1, input: '=A1+1' }])
  assert.equal(wb.getDisplay(s1, 0, 0).text, '#CYCLE!')
  assert.match(wb.explainError(s1, 0, 0)!, /Circular/)
})

test('empty cells, text and booleans in aggregates', () => {
  const { wb, s1 } = wbWith([[1, 'x'], [null, true], [3, ''], ['=COUNT(A1:B3)', '=COUNTA(A1:B3)'], ['=COUNTBLANK(A1:B3)', '=SUM(A1:B3)']])
  assert.equal(val(wb, s1, 'A4'), 2)
  assert.equal(val(wb, s1, 'B4'), 4)
  assert.equal(val(wb, s1, 'A5'), 2)
  assert.equal(val(wb, s1, 'B5'), 4)
})

test('dates are serial numbers and date functions work', () => {
  const { wb, s1 } = wbWith([['2024-01-15', '=A1+30', '=YEAR(A1)', '=DATE(2024,1,1)', '=TEXT(A1,"dd mmm yyyy")', '=EOMONTH(A1,1)']])
  assert.equal(val(wb, s1, 'A1'), 45306)
  assert.equal(wb.getDisplay(s1, 0, 0).text, '2024-01-15')
  assert.equal(val(wb, s1, 'C1'), 2024)
  assert.equal(val(wb, s1, 'D1'), 45292)
  assert.equal(val(wb, s1, 'E1'), '15 Jan 2024')
  assert.equal(val(wb, s1, 'F1'), 45351)
})

test('ExcelPro plugin functions (TEXT override, CONCAT, RANK.EQ, MODE.SNGL, TYPE, ERROR.TYPE, NUMBERVALUE)', () => {
  const { wb, s1 } = wbWith([[3, 1, 2, 3, 5]])
  const put = (c: number, f: string) => {
    wb.setInputs(s1, [{ r: 1, c, input: f }])
    return wb.getValue(s1, 1, c)
  }
  assert.equal(put(0, '=TEXT(0.256,"0.0%")'), '25.6%')
  assert.equal(put(1, '=TEXT(1234.5,"#,##0.00")'), '1,234.50')
  assert.equal(put(2, '=CONCAT("a",B1:C1,"z")'), 'a12z')
  assert.equal(put(3, '=RANK.EQ(3,A1:E1)'), 2)
  assert.equal(put(4, '=RANK(5,A1:E1,1)'), 5)
  assert.equal(put(5, '=MODE.SNGL(A1:E1)'), 3)
  assert.equal(put(6, '=TYPE("x")'), 2)
  assert.equal(put(7, '=ERROR.TYPE(1/0)'), 2)
  assert.equal(put(8, '=NUMBERVALUE("1.234,5",",",".")'), 1234.5)
})

test('statistical, logical, text and financial samples', () => {
  const { wb, s1 } = wbWith([[2, 4, 4, 4, 5, 5, 7, 9]])
  const put = (f: string) => {
    wb.setInputs(s1, [{ r: 5, c: 0, input: f }])
    return wb.getValue(s1, 5, 0)
  }
  assert.equal(put('=STDEV.P(A1:H1)'), 2)
  assert.equal(put('=MEDIAN(A1:H1)'), 4.5)
  assert.equal(put('=LARGE(A1:H1,2)'), 7)
  assert.equal(put('=IFS(A1>5,"big",A1>1,"mid")'), 'mid')
  assert.equal(put('=SWITCH(B1,4,"four","other")'), 'four')
  assert.equal(put('=PROPER("hello wORLD")'), 'Hello World')
  assert.equal(put('=TEXTJOIN("-",TRUE,"a","","b")'), 'a-b')
  assert.equal(put('=SUBSTITUTE("a-b-c","-","+")'), 'a+b+c')
  assert.equal(Math.round((put('=PMT(0.05/12,360,200000)') as number) * 100) / 100, -1073.64)
})

test('dynamic arrays spill and are not persisted as constants', () => {
  const { wb, s1 } = wbWith([['=SEQUENCE(3)']])
  assert.equal(val(wb, s1, 'A3'), 3)
  const doc = wb.serializeSheet(s1)
  assert.deepEqual(doc.cells, [['=SEQUENCE(3)']])
})
