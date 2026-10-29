import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { execFileSync } from 'node:child_process'
import JSZip from 'jszip'
import { Workbook } from '../src/engine/workbook'
import { exportXlsx, importXlsx } from '../src/files/xlsx'
import { exportDelimited, importDelimited } from '../src/files/csv'

const fixture = (name: string) => fs.readFileSync(path.join(import.meta.dirname, 'fixtures', name))

async function load(name: string) {
  const res = await importXlsx(fixture(name), name)
  return { ...res, wb: Workbook.fromDoc(res.doc) }
}

for (const name of ['openpyxl-sample.xlsx', 'libreoffice-sample.xlsx']) {
  test(`XLSX import: ${name}`, async () => {
    const { wb, report } = await load(name)
    assert.deepEqual(wb.sheets.map((s) => s.name), ['Sales Report', 'Data', 'Hidden'])
    const s = wb.sheetByName('Sales Report')!.id
    const d = wb.sheetByName('Data')!.id
    assert.equal(wb.sheetByName('Hidden')!.hidden, true)
    // values & formulas
    assert.equal(wb.getRaw(s, 1, 3), '=SUM(B2:C2)')
    assert.equal(wb.getValue(s, 6, 3), 8750.75)
    assert.equal(wb.getDisplay(s, 1, 4).text, '28.6%')
    assert.equal(wb.getValue(s, 9, 1), 2900)
    assert.equal(wb.getEditText(s, 9, 1), '=VLOOKUP("East",A2:D5,4,FALSE)')
    assert.equal(wb.getValue(s, 13, 1), 1600) // _xlfn.XLOOKUP stripped and evaluated
    assert.equal(wb.getDisplay(s, 8, 1).text, '2024-03-15')
    assert.equal(wb.getValue(s, 11, 1), '00501') // text stays text
    assert.equal(wb.getValue(s, 12, 1), true)
    assert.equal(wb.getDisplay(s, 10, 1).text, '#NAME?') // unsupported function kept, not blanked
    assert.match(String(wb.getRaw(s, 10, 1)), /TEXTSPLIT/i)
    // cross-sheet + named range
    assert.equal(wb.getValue(d, 1, 0), 17501.5)
    assert.equal(wb.getValue(d, 0, 1), 1.8)
    // formatting
    const h = wb.getStyle(s, 0, 0)!
    assert.equal(h.bold, true)
    assert.equal(h.color, '#ffffff')
    assert.equal(h.fill, '#1f4e78')
    assert.equal(h.hAlign, 'center')
    assert.equal(h.wrap, true)
    assert.equal(wb.getStyle(s, 6, 0)?.fontFamily, 'Arial')
    assert.equal(wb.getStyle(s, 6, 0)?.fontSize, 14)
    assert.equal(wb.getStyle(s, 6, 3)?.borders?.bottom?.style, 'double')
    assert.equal(wb.getStyle(s, 6, 3)?.borders?.bottom?.color, '#ff0000')
    // layout
    const m = wb.meta(s)
    assert.deepEqual(m.merges, [{ r1: 15, c1: 0, r2: 15, c2: 2 }])
    assert.ok(Math.abs(m.colWidths[0] - 159) <= 3, `col A width ${m.colWidths[0]}`)
    assert.equal(m.rowHeights[0], 40)
    assert.ok(m.hiddenCols.includes(5))
    assert.equal(m.tabColor, '#00b050')
    // report
    assert.ok(report.unsupportedFunctions.some((f) => f.name === 'TEXTSPLIT'))
    assert.equal(wb.meta(s).condFormats?.length, 1) // conditional formatting is imported since Phase 2
    assert.equal(wb.meta(s).condFormats![0].style?.fill, '#ffc7ce')
  })
}

test('XLSX shared formulas (Excel-style) are expanded with relative references', async () => {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<Types/>')
  zip.file('xl/workbook.xml', '<workbook xmlns:r="r"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>')
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet><sheetData>' +
      '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>' +
      '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>' +
      '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c><c r="C3" t="inlineStr"><is><t>inline</t></is></c></row>' +
      '</sheetData></worksheet>',
  )
  const buf = await zip.generateAsync({ type: 'uint8array' })
  const { doc } = await importXlsx(buf, 'shared.xlsx')
  const wb = Workbook.fromDoc(doc)
  const s = wb.sheets[0].id
  assert.equal(wb.getRaw(s, 2, 1), '=A3*2')
  assert.equal(wb.getValue(s, 2, 1), 6)
  assert.equal(wb.getValue(s, 2, 2), 'inline')
})

test('XLSX import rejects corrupt files with a clear message', async () => {
  await assert.rejects(importXlsx(new Uint8Array([1, 2, 3]), 'bad.xlsx'), /not a valid \.xlsx/)
})

function buildSample() {
  const wb = Workbook.blank('Round Trip', ['Main', 'Second Sheet'])
  const [a, b] = wb.order
  wb.setInputs(a, [
    { r: 0, c: 0, input: 'Item' }, { r: 0, c: 1, input: 'Price' }, { r: 0, c: 2, input: 'Date' },
    { r: 1, c: 0, input: 'Pen' }, { r: 1, c: 1, input: '$1.50' }, { r: 1, c: 2, input: '2024-12-31' },
    { r: 2, c: 0, input: 'Book' }, { r: 2, c: 1, input: '12.25' }, { r: 2, c: 2, input: '2025-01-01' },
    { r: 3, c: 0, input: 'Total' }, { r: 3, c: 1, input: '=SUM(B2:B3)' }, { r: 3, c: 2, input: '=IFS(B4>10,"big",TRUE,"small")' },
    { r: 4, c: 0, input: "'007" }, { r: 4, c: 1, input: '=XLOOKUP("Book",A2:A3,B2:B3)' }, { r: 4, c: 2, input: '=B4>5' },
    { r: 5, c: 0, input: '=SEQUENCE(2)' },
  ])
  wb.setInputs(b, [{ r: 0, c: 0, input: "=Main!B4*2" }])
  wb.applyStyle(a, [{ r1: 0, c1: 0, r2: 0, c2: 2 }], { bold: true, fill: '#ddebf7', borders: { bottom: { style: 'medium', color: '#000000' } } })
  wb.applyStyle(a, [{ r1: 3, c1: 1, r2: 3, c2: 1 }], { numFmt: '#,##0.00', italic: true, color: '#c00000', hAlign: 'right', wrap: true })
  wb.merge(a, { r1: 8, c1: 0, r2: 8, c2: 3 }, true)
  wb.setColWidths(a, [0], 150)
  wb.setRowHeights(a, [0], 32)
  wb.setHidden(a, 'row', [7], true)
  wb.setTabColor(b, '#ff9900')
  wb.addNamedRange('TaxRate', '=Main!$B$2')
  return wb
}

test('XLSX export -> import round trip preserves formulas, styles and layout', async () => {
  const wb = buildSample()
  const bytes = await exportXlsx(wb)
  const { doc } = await importXlsx(bytes, 'rt.xlsx')
  const wb2 = Workbook.fromDoc(doc)
  const [a, b] = wb2.order
  assert.deepEqual(wb2.sheets.map((s) => s.name), ['Main', 'Second Sheet'])
  assert.equal(wb2.getRaw(a, 3, 1), '=SUM(B2:B3)')
  assert.equal(wb2.getEditText(a, 3, 2), '=IFS(B4>10,"big",TRUE,"small")')
  assert.equal(wb2.getValue(a, 3, 2), 'big')
  assert.equal(wb2.getValue(a, 4, 1), 12.25)
  assert.equal(wb2.getValue(a, 4, 0), '007')
  assert.equal(wb2.getValue(a, 6, 0), 2) // spilled
  assert.equal(wb2.getDisplay(a, 1, 1).text, '$1.50')
  assert.equal(wb2.getDisplay(a, 1, 2).text, '2024-12-31')
  assert.equal(wb2.getValue(b, 0, 0), 27.5)
  assert.equal(wb2.getStyle(a, 0, 1)?.bold, true)
  assert.equal(wb2.getStyle(a, 0, 1)?.fill, '#ddebf7')
  assert.equal(wb2.getStyle(a, 0, 1)?.borders?.bottom?.style, 'medium')
  const t = wb2.getStyle(a, 3, 1)!
  assert.deepEqual([t.numFmt, t.italic, t.color, t.hAlign, t.wrap], ['#,##0.00', true, '#c00000', 'right', true])
  assert.deepEqual(wb2.meta(a).merges, [{ r1: 8, c1: 0, r2: 8, c2: 3 }])
  assert.ok(Math.abs(wb2.meta(a).colWidths[0] - 150) <= 1)
  assert.equal(wb2.meta(a).rowHeights[0], 32)
  assert.ok(wb2.meta(a).hiddenRows.includes(7))
  assert.equal(wb2.sheets[1].tabColor, '#ff9900')
  assert.deepEqual(wb2.namedRanges(), [{ name: 'TaxRate', expression: '=Main!$B$2' }])
})

function hasPython(mod: string) {
  try {
    execFileSync('python3', ['-c', `import ${mod}`], { stdio: 'ignore' })
    return true
  } catch {
    return false
  }
}

test('exported XLSX opens in openpyxl with formulas as formulas', { skip: !hasPython('openpyxl') }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xp-'))
  const file = path.join(dir, 'out.xlsx')
  fs.writeFileSync(file, await exportXlsx(buildSample()))
  const out = execFileSync('python3', ['-c', `
import openpyxl, json
wb = openpyxl.load_workbook(${JSON.stringify(file)})
ws = wb["Main"]
print(json.dumps({"sheets": wb.sheetnames, "b4": ws["B4"].value, "c4": ws["C4"].value, "b5": ws["B5"].value, "bold": ws["A1"].font.b,
  "fill": ws["A1"].fill.fgColor.rgb, "fmt": ws["B4"].number_format, "merged": [str(r) for r in ws.merged_cells.ranges],
  "a5": ws["A5"].value, "names": list(wb.defined_names.keys())}))
`]).toString()
  const j = JSON.parse(out)
  assert.deepEqual(j.sheets, ['Main', 'Second Sheet'])
  assert.equal(j.b4, '=SUM(B2:B3)')
  assert.equal(j.c4, '=_xlfn.IFS(B4>10,"big",TRUE,"small")')
  assert.equal(j.b5, '=_xlfn.XLOOKUP("Book",A2:A3,B2:B3)')
  assert.equal(j.bold, true)
  assert.equal(j.fill, 'FFDDEBF7')
  assert.equal(j.fmt, '#,##0.00')
  assert.deepEqual(j.merged, ['A9:D9'])
  assert.equal(j.a5, '007')
  assert.deepEqual(j.names, ['TaxRate'])
})

test('exported XLSX recalculates correctly in LibreOffice', { skip: !fs.existsSync('/usr/bin/soffice') }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xp-'))
  const file = path.join(dir, 'calc.xlsx')
  fs.writeFileSync(file, await exportXlsx(buildSample()))
  execFileSync('soffice', ['--headless', '--convert-to', 'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,true', '--outdir', dir, file], { stdio: 'ignore', timeout: 120000 })
  const csv = fs.readFileSync(path.join(dir, 'calc.csv'), 'utf8').split(/\r?\n/)
  assert.equal(csv[3], 'Total,13.75,big')
  assert.match(csv[1], /^Pen,\$1\.50,/)
})

test('CSV import parses values like Excel and export writes displayed values', () => {
  const csv = '﻿Name,Amount,When,Note\r\n"Smith, J",1200.5,2024-02-01,"He said ""hi"""\r\nLee,15%,,=1+1\r\n'
  const { doc, report } = importDelimited(csv, 'people.csv')
  const wb = Workbook.fromDoc(doc)
  const s = wb.sheets[0].id
  assert.equal(wb.sheets[0].name, 'people')
  assert.equal(wb.getValue(s, 1, 0), 'Smith, J')
  assert.equal(wb.getValue(s, 1, 1), 1200.5)
  assert.equal(wb.getDisplay(s, 1, 2).text, '2024-02-01')
  assert.equal(wb.getValue(s, 1, 3), 'He said "hi"')
  assert.equal(wb.getValue(s, 2, 1), 0.15)
  assert.equal(wb.getValue(s, 2, 3), 2)
  assert.equal(report.stats.cells, 11)
  const out = exportDelimited(wb, s, ',')
  assert.equal(out.split('\r\n')[1], '"Smith, J",1200.5,2024-02-01,"He said ""hi"""')
  assert.equal(out.split('\r\n')[2], 'Lee,15%,,2')
  const tsv = exportDelimited(wb, s, '\t')
  assert.equal(tsv.split('\r\n')[1], 'Smith, J\t1200.5\t2024-02-01\t"He said ""hi"""')
  const semi = importDelimited('a;b\n1;2\n', 'x.csv', { delimiter: ';' })
  assert.equal(Workbook.fromDoc(semi.doc).getValue(Workbook.fromDoc(semi.doc).order[0], 1, 1), 2)
  const t = importDelimited('a\tb\n3\t4\n', 'x.tsv')
  const wt = Workbook.fromDoc(t.doc)
  assert.equal(wt.getValue(wt.order[0], 1, 1), 4)
})

// ------------------------------------------------------------------ Phase 2/3 file features
import { addValidation, setAutoFilter, setFilterCriterion, setFreeze, setLink, setNote, setPageSetup } from '../src/engine/features/data'
import { addCondFormat } from '../src/engine/features/condFormat'
import { createTable, updateTable } from '../src/engine/features/tables'

function richSample() {
  const wb = Workbook.blank('Rich', ['Data', 'Notes'])
  const [a] = wb.order
  const rows = [['Item', 'Qty', 'Price', 'Due'], ['Pen', '3', '1.5', '2024-05-01'], ['Book', '1', '12', '2024-06-10'], ['Bag', '2', '30', '2024-07-15']]
  rows.forEach((row, r) => row.forEach((v, c) => wb.setInputs(a, [{ r, c, input: v }])))
  createTable(wb, a, { r1: 0, c1: 0, r2: 3, c2: 3 }, { hasHeader: true, style: 'blue' })
  updateTable(wb, a, wb.meta(a).tables![0].id, { totalRow: true, totals: { 1: 'sum', 2: 'average' } })
  wb.setInputs(a, [{ r: 7, c: 0, input: '=SUM(Table1[Qty])' }])
  addValidation(wb, a, { range: { r1: 1, c1: 1, r2: 20, c2: 1 }, type: 'whole', op: 'between', v1: '0', v2: '100', allowBlank: true, showDropdown: false, errorStyle: 'stop', errorMessage: 'Qty 0-100' })
  addValidation(wb, a, { range: { r1: 1, c1: 5, r2: 20, c2: 5 }, type: 'list', list: 'Open,Closed', allowBlank: true, showDropdown: true, errorStyle: 'warning' })
  addCondFormat(wb, a, { range: { r1: 1, c1: 2, r2: 3, c2: 2 }, rule: { type: 'cellIs', op: 'gt', v1: '10' }, style: { fill: '#ffc7ce', color: '#9c0006' } })
  addCondFormat(wb, a, { range: { r1: 1, c1: 1, r2: 3, c2: 1 }, rule: { type: 'colorScale', colors: ['#f8696b', '#ffeb84', '#63be7b'] } })
  addCondFormat(wb, a, { range: { r1: 1, c1: 0, r2: 3, c2: 0 }, rule: { type: 'formula', formula: '=$B2>1' }, style: { bold: true } })
  setFreeze(wb, a, 1, 1)
  setNote(wb, a, 1, 0, 'Blue ink only')
  setLink(wb, a, 2, 0, 'https://example.com/book')
  setPageSetup(wb, a, { orientation: 'landscape', paper: 'Letter', fitToWidth: true, gridlines: true, printArea: { r1: 0, c1: 0, r2: 4, c2: 3 }, rowBreaks: [3], footer: 'Page &P of &N' })
  const b = wb.order[1]
  ;[['Name', 'Score'], ['A', '5'], ['B', '9']].forEach((row, r) => row.forEach((v, c) => wb.setInputs(b, [{ r, c, input: v }])))
  setAutoFilter(wb, b, { r1: 0, c1: 0, r2: 2, c2: 1 })
  setFilterCriterion(wb, b, 1, { kind: 'condition', op: 'gt', a: '6' })
  return wb
}

test('XLSX round trip of tables, validation, conditional formatting, freeze, notes, links, page setup, filters', async () => {
  const wb = richSample()
  const { doc, report } = await importXlsx(await exportXlsx(wb), 'rich.xlsx')
  const wb2 = Workbook.fromDoc(doc)
  const [a, b] = wb2.order
  const m = wb2.meta(a)
  assert.equal(m.tables?.[0].name, 'Table1')
  assert.equal(m.tables?.[0].totalRow, true)
  assert.equal(m.tables?.[0].totals[1], 'sum')
  assert.equal(wb2.getValue(a, 4, 1), 6)
  assert.equal(wb2.getValue(a, 7, 0), 6)
  assert.equal(m.validations?.length, 2)
  assert.deepEqual([m.validations![0].type, m.validations![0].op, m.validations![0].v1, m.validations![0].v2, m.validations![0].errorMessage], ['whole', 'between', '0', '100', 'Qty 0-100'])
  assert.equal(m.validations![1].list, 'Open,Closed')
  assert.equal(m.validations![1].showDropdown, true)
  assert.equal(m.condFormats?.length, 3)
  assert.equal(m.condFormats![0].rule.type, 'formula')
  assert.equal(m.condFormats![2].style?.fill, '#ffc7ce')
  assert.deepEqual((m.condFormats![1].rule as { colors: string[] }).colors, ['#f8696b', '#ffeb84', '#63be7b'])
  assert.deepEqual(m.freeze, { rows: 1, cols: 1 })
  assert.equal(m.notes?.['1,0'], 'Blue ink only')
  assert.equal(m.links?.['2,0'], 'https://example.com/book')
  assert.equal(m.pageSetup?.orientation, 'landscape')
  assert.equal(m.pageSetup?.paper, 'Letter')
  assert.equal(m.pageSetup?.fitToWidth, true)
  assert.equal(m.pageSetup?.gridlines, true)
  assert.deepEqual(m.pageSetup?.printArea, { r1: 0, c1: 0, r2: 4, c2: 3 })
  assert.deepEqual(m.pageSetup?.rowBreaks, [3])
  const mb = wb2.meta(b)
  assert.deepEqual(mb.autoFilter?.range, { r1: 0, c1: 0, r2: 2, c2: 1 })
  assert.ok(mb.hiddenRows.includes(1), 'filtered-out row exported as hidden')
  assert.ok(!report.warnings.some((w) => /Conditional|validation|Hyperlink|Comments|tables/i.test(w)), report.warnings.join('\n'))
})

test('Phase 2/3 parts open in openpyxl', { skip: !hasPython('openpyxl') }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xp-'))
  const file = path.join(dir, 'rich.xlsx')
  fs.writeFileSync(file, await exportXlsx(richSample()))
  const out = execFileSync('python3', ['-c', `
import openpyxl, json, warnings
warnings.simplefilter('ignore')
wb = openpyxl.load_workbook(${JSON.stringify(file)})
ws = wb['Data']
print(json.dumps({
 "tables": list(ws.tables.keys()),
 "dv": [d.type for d in ws.data_validations.dataValidation],
 "cf": len(ws.conditional_formatting),
 "freeze": ws.freeze_panes,
 "note": ws['A2'].comment.text if ws['A2'].comment else None,
 "link": ws['A3'].hyperlink.target if ws['A3'].hyperlink else None,
 "orient": ws.page_setup.orientation,
 "area": ws.print_area,
 "filter": wb['Notes'].auto_filter.ref,
}))
`]).toString()
  const j = JSON.parse(out)
  assert.deepEqual(j.tables, ['Table1'])
  assert.deepEqual(j.dv.sort(), ['list', 'whole'])
  assert.equal(j.cf, 3)
  assert.equal(j.freeze, 'B2')
  assert.equal(j.note, 'Blue ink only')
  assert.equal(j.link, 'https://example.com/book')
  assert.equal(j.orient, 'landscape')
  assert.match(String(j.area), /\$A\$1:\$D\$5/)
  assert.equal(j.filter, 'A1:B3')
})

test('Phase 2/3 workbook opens and recalculates in LibreOffice', { skip: !fs.existsSync('/usr/bin/soffice') }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xp-'))
  const file = path.join(dir, 'rich.xlsx')
  fs.writeFileSync(file, await exportXlsx(richSample()))
  execFileSync('soffice', ['--headless', '--convert-to', 'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,true', '--outdir', dir, file], { stdio: 'ignore', timeout: 120000 })
  const csv = fs.readFileSync(path.join(dir, 'rich.csv'), 'utf8').split(/\r?\n/)
  assert.match(csv[4], /^Total,6,14\.5/)
  assert.match(csv[7], /^6/)
})
