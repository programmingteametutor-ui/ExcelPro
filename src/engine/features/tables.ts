/**
 * Spreadsheet tables: a named range with a header row, banded styling, optional total row,
 * per-column filtering and structured references (Table1[Column], Table1[#All], …).
 *
 * Structured references are translated to ordinary absolute ranges when a formula is entered,
 * so they work with the formula engine; they do not re-expand automatically if the table grows.
 */
import { colToLetters, quoteSheetName, rangesIntersect } from '../../model/address'
import type { CellStyle, Range, TableDef, TotalFn } from '../../model/types'
import type { Workbook } from '../workbook'
import { uid } from './common'

export const TABLE_STYLES: Record<TableDef['style'], { header: string; headerText: string; band: string; total: string }> = {
  teal: { header: '#0f766e', headerText: '#ffffff', band: '#e3f4f1', total: '#c9e9e3' },
  blue: { header: '#2f5597', headerText: '#ffffff', band: '#dde7f5', total: '#c4d4ec' },
  orange: { header: '#c55a11', headerText: '#ffffff', band: '#fbe5d6', total: '#f6cdb0' },
  gray: { header: '#595959', headerText: '#ffffff', band: '#ededed', total: '#d9d9d9' },
  green: { header: '#548235', headerText: '#ffffff', band: '#e2efd9', total: '#c9e1b8' },
  purple: { header: '#7030a0', headerText: '#ffffff', band: '#ece2f4', total: '#d9c6e9' },
}

const TOTAL_CODE: Record<Exclude<TotalFn, 'none'>, number> = { average: 101, count: 103, countNumbers: 102, max: 104, min: 105, sum: 109 }

export function tableAt(wb: Workbook, sheetId: string, r: number, c: number): TableDef | undefined {
  return wb.meta(sheetId).tables?.find((t) => r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2)
}

export function allTables(wb: Workbook): { sheetId: string; table: TableDef }[] {
  return wb.sheets.flatMap((s) => (s.tables ?? []).map((t) => ({ sheetId: s.id, table: t })))
}

export function validTableName(wb: Workbook, name: string, exceptId?: string): string | null {
  if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name) || /^[RC]$/i.test(name)) return 'Table names must start with a letter and contain only letters, numbers, periods and underscores.'
  const lower = name.toLowerCase()
  if (allTables(wb).some((t) => t.table.name.toLowerCase() === lower && t.table.id !== exceptId)) return `A table named ${name} already exists.`
  if (wb.namedRanges().some((n) => n.name.toLowerCase() === lower)) return `${name} is already a defined name.`
  return null
}

export function headerNames(wb: Workbook, sheetId: string, t: TableDef): string[] {
  const out: string[] = []
  for (let c = t.range.c1; c <= t.range.c2; c++) {
    const d = wb.getDisplay(sheetId, t.range.r1, c).text.trim()
    out.push(d || `Column${c - t.range.c1 + 1}`)
  }
  return out
}

/** Data rows (excluding header and total rows). */
export function dataRange(t: TableDef): Range {
  return { ...t.range, r1: t.range.r1 + 1, r2: t.totalRow ? t.range.r2 - 1 : t.range.r2 }
}

/** Converts a range into a table. Missing headers are filled in as Column1, Column2… */
export function createTable(wb: Workbook, sheetId: string, range: Range, opts: { hasHeader: boolean; style?: TableDef['style']; name?: string }): { id?: string; error?: string } {
  const meta = wb.meta(sheetId)
  if ((meta.tables ?? []).some((t) => rangesIntersect(t.range, range))) return { error: 'Tables cannot overlap another table.' }
  if (meta.merges.some((m) => rangesIntersect(m, range))) return { error: 'Tables cannot contain merged cells.' }
  let n = allTables(wb).length + 1
  let name = opts.name ?? `Table${n}`
  while (!opts.name && validTableName(wb, name)) name = `Table${++n}`
  const err = validTableName(wb, name)
  if (err) return { error: err }
  let rg = { ...range }
  const hfId = wb.hfId(sheetId)
  const headerWrites: { c: number; text: string }[] = []
  if (!opts.hasHeader) {
    // insert a header row above the data by shifting the block down one row
    rg = { ...rg, r2: rg.r2 + 1 }
  }
  for (let c = rg.c1; c <= rg.c2; c++) {
    const text = opts.hasHeader ? wb.getDisplay(sheetId, rg.r1, c).text.trim() : ''
    if (!text) headerWrites.push({ c, text: `Column${c - rg.c1 + 1}` })
  }
  const id = uid('tbl')
  const block = !opts.hasHeader
  wb.run({
    label: 'Create table',
    dirty: 'all',
    metaScope: [sheetId],
    hfOps:
      headerWrites.length || block
        ? () => {
            if (block) {
              const hf = wb.hf
              hf.moveCells({ start: { sheet: hfId, row: range.r1, col: range.c1 }, end: { sheet: hfId, row: range.r2, col: range.c2 } }, { sheet: hfId, row: range.r1 + 1, col: range.c1 })
            }
            for (const h of headerWrites) wb.hf.setCellContents({ sheet: hfId, row: rg.r1, col: h.c }, h.text)
          }
        : undefined,
    metaOps: () => {
      const m = wb.meta(sheetId)
      m.tables = [...(m.tables ?? []), { id, name, range: rg, style: opts.style ?? 'teal', banded: true, totalRow: false, totals: {}, filter: true }]
      if (!m.autoFilter) {
        m.autoFilter = { range: rg, criteria: {} }
        m.filterHidden = []
      }
    },
  })
  return { id }
}

export function updateTable(wb: Workbook, sheetId: string, id: string, patch: Partial<Omit<TableDef, 'id' | 'range'>>): string | null {
  const t = wb.meta(sheetId).tables?.find((x) => x.id === id)
  if (!t) return 'Table not found.'
  if (patch.name && patch.name !== t.name) {
    const err = validTableName(wb, patch.name, id)
    if (err) return err
  }
  const next: TableDef = { ...t, ...patch }
  const hfId = wb.hfId(sheetId)
  const ops: (() => void)[] = []
  let range = t.range
  if (patch.totalRow !== undefined && patch.totalRow !== t.totalRow) {
    if (patch.totalRow) {
      range = { ...t.range, r2: t.range.r2 + 1 }
      const totals: Record<number, TotalFn> = { ...next.totals }
      const lastOff = t.range.c2 - t.range.c1
      if (!Object.keys(totals).length) totals[lastOff] = 'sum'
      next.totals = totals
    } else {
      range = { ...t.range, r2: t.range.r2 - 1 }
      ops.push(() => {
        for (let c = t.range.c1; c <= t.range.c2; c++) wb.hf.setCellContents({ sheet: hfId, row: t.range.r2, col: c }, null)
      })
    }
  }
  next.range = range
  if (next.totalRow) {
    const totalR = range.r2
    const dr = dataRange(next)
    ops.push(() => {
      for (let c = range.c1; c <= range.c2; c++) {
        const off = c - range.c1
        const fn = next.totals[off] ?? 'none'
        let content: string | null = null
        if (off === 0 && fn === 'none') content = 'Total'
        else if (fn !== 'none') content = `=SUBTOTAL(${TOTAL_CODE[fn]},${colToLetters(c)}${dr.r1 + 1}:${colToLetters(c)}${dr.r2 + 1})`
        wb.hf.setCellContents({ sheet: hfId, row: totalR, col: c }, content)
      }
    })
  }
  wb.run({
    label: 'Table design',
    dirty: 'all',
    metaScope: [sheetId],
    hfOps: ops.length ? () => ops.forEach((f) => f()) : undefined,
    metaOps: () => {
      const m = wb.meta(sheetId)
      m.tables = (m.tables ?? []).map((x) => (x.id === id ? next : x))
    },
  })
  return null
}

/** Removes the table object, keeping the cell contents (Convert to Range). */
export function convertToRange(wb: Workbook, sheetId: string, id: string, keepFormatting = true) {
  const t = wb.meta(sheetId).tables?.find((x) => x.id === id)
  if (!t) return
  const st = TABLE_STYLES[t.style]
  wb.run({
    label: 'Convert to range',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      if (keepFormatting) {
        for (let r = t.range.r1; r <= t.range.r2; r++)
          for (let c = t.range.c1; c <= t.range.c2; c++) {
            const deco = tableCellStyle(t, r, c, st)
            if (deco) m.styles[`${r},${c}`] = { ...deco, ...m.styles[`${r},${c}`] }
          }
      }
      m.tables = (m.tables ?? []).filter((x) => x.id !== id)
      if (m.autoFilter && m.autoFilter.range.r1 === t.range.r1 && m.autoFilter.range.c1 === t.range.c1) m.autoFilter = undefined
    },
  })
}

/** Style a table contributes to one cell (cell formatting still wins over it). */
export function tableCellStyle(t: TableDef, r: number, c: number, st = TABLE_STYLES[t.style]): CellStyle | undefined {
  if (r === t.range.r1) return { fill: st.header, color: st.headerText, bold: true }
  if (t.totalRow && r === t.range.r2) return { fill: st.total, bold: true, borders: { top: { style: 'double', color: st.header } } }
  if (t.banded && (r - t.range.r1) % 2 === 0) return { fill: st.band }
  void c
  return undefined
}

/** Extends a table by one row when the user types directly below it. Returns true when extended. */
export function maybeExtendTable(wb: Workbook, sheetId: string, r: number, c: number): boolean {
  const t = wb.meta(sheetId).tables?.find((x) => !x.totalRow && r === x.range.r2 + 1 && c >= x.range.c1 && c <= x.range.c2)
  if (!t) return false
  wb.run({
    label: 'Extend table',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      m.tables = (m.tables ?? []).map((x) => (x.id === t.id ? { ...x, range: { ...x.range, r2: x.range.r2 + 1 } } : x))
    },
  })
  return true
}

export interface TableInfo {
  name: string
  sheetName: string
  range: Range
  headers: string[]
  totalRow: boolean
}

/**
 * Translates structured references to absolute ranges:
 * Table1[Column], Table1[[#Headers],[Col]], Table1[#All], Table1[#Data], Table1[#Headers],
 * Table1[#Totals], Table1[@Column] (same row), and Table1 alone (data rows).
 */
export function resolveStructuredRefsWith(formula: string, tables: TableInfo[], fromSheetName: string, row: number): string {
  if (!formula.startsWith('=') || !tables.length) return formula
  const byName = new Map(tables.map((t) => [t.name.toLowerCase(), t]))
  const abs = (t: TableInfo, rg: Range) => {
    const ref = rg.r1 === rg.r2 && rg.c1 === rg.c2 ? `$${colToLetters(rg.c1)}$${rg.r1 + 1}` : `$${colToLetters(rg.c1)}$${rg.r1 + 1}:$${colToLetters(rg.c2)}$${rg.r2 + 1}`
    return t.sheetName === fromSheetName ? ref : `${quoteSheetName(t.sheetName)}!${ref}`
  }
  const parts = formula.split(/("(?:[^"]|"")*")/)
  return parts
    .map((part, i) => {
      if (i % 2 === 1) return part
      return part.replace(/(?<![A-Za-z0-9_.!$'])([A-Za-z_\\][A-Za-z0-9_.]*)(\[(?:[^[\]]|\[[^\]]*\])*\])?(?![A-Za-z0-9_(])/g, (m, name: string, spec: string | undefined) => {
        const t = byName.get(name.toLowerCase())
        if (!t) return m
        const heads = t.headers.map((h) => h.toLowerCase())
        const data: Range = { ...t.range, r1: t.range.r1 + 1, r2: t.totalRow ? t.range.r2 - 1 : t.range.r2 }
        if (!spec) return abs(t, data)
        const inner = spec.slice(1, -1).trim()
        const items = inner.startsWith('[') ? [...inner.matchAll(/\[([^\]]*)\]/g)].map((x) => x[1].trim()) : [inner]
        let rows: Range = data
        let cols: number[] = []
        let thisRow = false
        for (const it of items) {
          const low = it.toLowerCase()
          if (low === '#all') rows = t.range
          else if (low === '#data') rows = data
          else if (low === '#headers') rows = { ...t.range, r2: t.range.r1 }
          else if (low === '#totals') rows = t.totalRow ? { ...t.range, r1: t.range.r2 } : data
          else if (low === '#this row' || low === '@') thisRow = true
          else {
            let colName = low
            if (colName.startsWith('@')) {
              thisRow = true
              colName = colName.slice(1).trim()
            }
            const idx = heads.indexOf(colName.replace(/^'|'$/g, ''))
            if (idx < 0) return m
            cols.push(t.range.c1 + idx)
          }
        }
        if (thisRow) rows = { ...rows, r1: row, r2: row }
        if (!cols.length) cols = [t.range.c1, t.range.c2]
        return abs(t, { r1: rows.r1, r2: rows.r2, c1: Math.min(...cols), c2: Math.max(...cols) })
      })
    })
    .join('')
}

export function resolveStructuredRefs(wb: Workbook, formula: string, fromSheetId: string, row: number): string {
  const tables = allTables(wb)
  if (!tables.length || !formula.startsWith('=')) return formula
  const infos: TableInfo[] = tables.map(({ sheetId, table }) => ({ name: table.name, sheetName: wb.meta(sheetId).name, range: table.range, headers: headerNames(wb, sheetId, table), totalRow: table.totalRow }))
  return resolveStructuredRefsWith(formula, infos, wb.meta(fromSheetId).name, row)
}
