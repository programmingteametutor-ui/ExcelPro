/** CSV/TSV import & export (PapaParse for parsing). */
import Papa from 'papaparse'
import type { SheetDoc, WorkbookDoc } from '../model/types'
import { parseInput, type DateOrder, type RawContent } from '../engine/inputParser'
import { formatForCsv, newId, type Workbook } from '../engine/workbook'
import type { ImportResult } from './xlsx'

export type Delimiter = ',' | ';' | '\t' | '|'

export function importDelimited(text: string, fileName: string, opts: { delimiter?: Delimiter; dateOrder?: DateOrder } = {}): ImportResult {
  const isTsv = /\.tsv$/i.test(fileName) || /\.tab$/i.test(fileName)
  const clean = text.replace(/^﻿/, '')
  const res = Papa.parse<string[]>(clean, {
    delimiter: opts.delimiter ?? (isTsv ? '\t' : ''),
    skipEmptyLines: false,
  })
  const rows = res.data
  while (rows.length && rows[rows.length - 1].every((v) => v === '')) rows.pop()
  const warnings: string[] = []
  const realErrors = res.errors.filter((e) => e.code !== 'UndetectableDelimiter')
  if (realErrors.length) warnings.push(`Some lines could not be parsed cleanly (${realErrors.length}); check quotes in the source file.`)
  const sheet: SheetDoc = {
    id: newId(),
    name: (fileName.replace(/\.[^.]+$/, '') || 'Sheet1').replace(/[\\/?*[\]:]/g, '_').slice(0, 31),
    colWidths: {},
    rowHeights: {},
    hiddenRows: [],
    hiddenCols: [],
    styles: {},
    merges: [],
    cells: [],
  }
  let cellCount = 0
  const cells: RawContent[][] = rows.map((row, r) =>
    row.map((v, c) => {
      if (v === '') return null
      cellCount++
      // CSV files never contain formulas as intent from ExcelPro's side unless they start with "=",
      // which Excel also evaluates. Keep that behaviour, but parse values like typed input.
      const p = parseInput(v, undefined, opts.dateOrder ?? 'MDY')
      if (p.autoFormat) sheet.styles[`${r},${c}`] = { numFmt: p.autoFormat }
      return p.raw
    }),
  )
  sheet.cells = cells
  const now = new Date().toISOString()
  const doc: WorkbookDoc = {
    schemaVersion: 1,
    id: newId('wb'),
    name: fileName.replace(/\.[^.]+$/, '') || 'Imported',
    activeSheetId: sheet.id,
    sheetOrder: [sheet.id],
    sheets: [sheet],
    namedRanges: [],
    createdAt: now,
    updatedAt: now,
  }
  const delim = res.meta.delimiter === '\t' ? 'tab' : `"${res.meta.delimiter}"`
  warnings.push(`Read ${rows.length} rows using ${delim} as the delimiter. CSV files hold values only: formatting, formulas from other cells and extra sheets are not part of the format.`)
  return { doc, report: { warnings, unsupportedFunctions: [], stats: { sheets: 1, cells: cellCount, formulas: 0 } } }
}

/** Exports one sheet's displayed values (like Excel's "Save as CSV"). */
export function exportDelimited(wb: Workbook, sheetId: string, delimiter: Delimiter = ',', opts: { bom?: boolean } = {}): string {
  const hfId = wb.hfId(sheetId)
  const dims = wb.hf.getSheetDimensions(hfId)
  const lines: string[] = []
  const needsQuote = new RegExp(`[${delimiter === '\t' ? '\\t' : delimiter.replace(/[|]/g, '\\|')}"\\r\\n]|^\\s|\\s$`)
  for (let r = 0; r < dims.height; r++) {
    const fields: string[] = []
    for (let c = 0; c < dims.width; c++) {
      const t = formatForCsv(wb, sheetId, r, c)
      fields.push(needsQuote.test(t) ? `"${t.replace(/"/g, '""')}"` : t)
    }
    while (fields.length && fields[fields.length - 1] === '') fields.pop()
    lines.push(fields.join(delimiter))
  }
  return (opts.bom ? '﻿' : '') + lines.join('\r\n') + (lines.length ? '\r\n' : '')
}
