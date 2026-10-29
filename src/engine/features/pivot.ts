/**
 * Pivot tables: group the source rows by row/column fields, aggregate value fields, and write the
 * result into the sheet. The definition is stored, so the pivot can be refreshed when its source
 * changes (manually, or automatically when autoRefresh is on).
 */
import type { RawCellContent } from 'hyperformula'
import { cellKey, rangesIntersect } from '../../model/address'
import type { AggFn, CellStyle, PivotDef, Range } from '../../model/types'
import { isDateFormat, serialToParts } from '../numberFormat'
import type { Workbook } from '../workbook'
import { clampToUsed, compareValues, toScalar, uid, type Scalar } from './common'

export interface PivotSource {
  fields: string[]
  rows: Scalar[][]
  dateFields: Set<string>
  formats: Record<string, string | undefined>
}

export function readPivotSource(wb: Workbook, sheetId: string, range: Range): PivotSource {
  const rg = clampToUsed(wb, sheetId, range)
  const fields: string[] = []
  const seen = new Map<string, number>()
  for (let c = rg.c1; c <= rg.c2; c++) {
    let name = wb.getDisplay(sheetId, rg.r1, c).text.trim() || `Column${c - rg.c1 + 1}`
    const n = seen.get(name.toLowerCase()) ?? 0
    seen.set(name.toLowerCase(), n + 1)
    if (n) name = `${name}${n + 1}`
    fields.push(name)
  }
  const rows: Scalar[][] = []
  for (let r = rg.r1 + 1; r <= rg.r2; r++) {
    const row: Scalar[] = []
    let any = false
    for (let c = rg.c1; c <= rg.c2; c++) {
      const v = toScalar(wb.getValue(sheetId, r, c))
      if (v !== null && v !== '') any = true
      row.push(v)
    }
    if (any) rows.push(row)
  }
  const dateFields = new Set<string>()
  const formats: Record<string, string | undefined> = {}
  fields.forEach((f, i) => {
    const fmt = wb.getStyle(sheetId, rg.r1 + 1, rg.c1 + i)?.numFmt
    formats[f] = fmt
    if (isDateFormat(fmt)) dateFields.add(f)
  })
  return { fields, rows, dateFields, formats }
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

function label(v: Scalar, field: string, src: PivotSource, def: PivotDef): string {
  if (v === null || v === '') return '(blank)'
  if (typeof v === 'object') return v.error
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'number' && src.dateFields.has(field)) {
    const p = serialToParts(v)
    const g = def.dateGroup?.[field]
    if (g === 'year') return String(p.y)
    if (g === 'quarter') return `${p.y} Q${Math.ceil(p.m / 3)}`
    if (g === 'month') return `${p.y}-${String(p.m).padStart(2, '0')} ${MONTHS[p.m - 1]}`
    return `${p.y}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}`
  }
  return String(v)
}

function aggregate(fn: AggFn, vals: Scalar[]): number | null {
  const nums = vals.filter((v): v is number => typeof v === 'number')
  switch (fn) {
    case 'sum':
      return nums.reduce((a, b) => a + b, 0)
    case 'count':
      return vals.filter((v) => v !== null && v !== '').length
    case 'average':
      return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : null
    case 'min':
      return nums.length ? Math.min(...nums) : null
    case 'max':
      return nums.length ? Math.max(...nums) : null
    case 'countDistinct':
      return new Set(vals.filter((v) => v !== null && v !== '').map((v) => (typeof v === 'object' && v ? v.error : String(v)))).size
  }
}

export const AGG_LABEL: Record<AggFn, string> = { sum: 'Sum', count: 'Count', average: 'Average', min: 'Min', max: 'Max', countDistinct: 'Distinct count' }

export interface PivotOutput {
  cells: (string | number | null)[][]
  headerRows: number
  rowFieldCols: number
  totalRow: number | null
  totalCol: number | null
  valueCols: { col: number; valueIndex: number }[]
}

/** Pure pivot computation (unit-tested without a workbook). */
export function computePivot(src: PivotSource, def: PivotDef): PivotOutput {
  const idx = (f: string) => src.fields.indexOf(f)
  const rowsF = def.rows.filter((f) => idx(f) >= 0)
  const colsF = def.cols.filter((f) => idx(f) >= 0)
  const vals = def.values.filter((v) => idx(v.field) >= 0)
  const filtered = src.rows.filter((row) =>
    def.filters.every((flt) => idx(flt.field) < 0 || flt.values.length === 0 || flt.values.includes(label(row[idx(flt.field)], flt.field, src, def))),
  )
  const SEP = '\u0001'
  const keyOf = (row: Scalar[], fields: string[]) => fields.map((f) => label(row[idx(f)], f, src, def)).join(SEP)
  const sortKeys = (keys: string[], fields: string[]) => {
    // sort by the underlying values of each level (numbers numerically)
    const sample = new Map<string, Scalar[]>()
    for (const row of filtered) {
      const k = keyOf(row, fields)
      if (!sample.has(k)) sample.set(k, fields.map((f) => row[idx(f)]))
    }
    const dir = def.sort === 'desc' ? -1 : 1
    return keys.sort((a, b) => {
      const va = sample.get(a) ?? []
      const vb = sample.get(b) ?? []
      for (let i = 0; i < fields.length; i++) {
        const la = a.split(SEP)[i]
        const lb = b.split(SEP)[i]
        const c = def.dateGroup?.[fields[i]] ? la.localeCompare(lb) : compareValues(va[i] ?? null, vb[i] ?? null) || la.localeCompare(lb)
        if (c) return c * dir
      }
      return 0
    })
  }
  const rowKeys = sortKeys([...new Set(filtered.map((r) => keyOf(r, rowsF)))], rowsF)
  const colKeys = colsF.length ? sortKeys([...new Set(filtered.map((r) => keyOf(r, colsF)))], colsF) : ['']
  const groups = new Map<string, Scalar[][]>()
  for (const row of filtered) {
    const k = keyOf(row, rowsF) + '\u0002' + (colsF.length ? keyOf(row, colsF) : '')
    let g = groups.get(k)
    if (!g) groups.set(k, (g = []))
    g.push(row)
  }
  const cellVal = (rowsSel: Scalar[][], vi: number) => aggregate(vals[vi].agg, rowsSel.map((r) => r[idx(vals[vi].field)]))
  const valueLabel = (vi: number) => `${AGG_LABEL[vals[vi].agg]} of ${vals[vi].field}`
  const out: (string | number | null)[][] = []
  const nRowF = Math.max(1, rowsF.length)
  const valueCols: { col: number; valueIndex: number }[] = []
  // header rows
  const headerRows = colsF.length ? 2 : 1
  if (colsF.length) {
    const h0: (string | number | null)[] = new Array(nRowF).fill(null)
    h0[0] = vals.length === 1 ? valueLabel(0) : 'Values'
    for (const ck of colKeys) for (let vi = 0; vi < Math.max(1, vals.length); vi++) h0.push(vi === 0 ? ck.split(SEP).join(' / ') : null)
    if (def.grandTotals) for (let vi = 0; vi < Math.max(1, vals.length); vi++) h0.push(vi === 0 ? 'Grand Total' : null)
    out.push(h0)
  }
  const h1: (string | number | null)[] = rowsF.length ? [...rowsF] : ['']
  for (let k = 0; k < colKeys.length; k++) for (let vi = 0; vi < vals.length; vi++) h1.push(colsF.length && vals.length === 1 ? colsF.join(' / ') === '' ? valueLabel(vi) : valueLabel(vi) : valueLabel(vi))
  if (colsF.length && def.grandTotals) for (let vi = 0; vi < vals.length; vi++) h1.push(valueLabel(vi))
  out.push(h1)
  // body
  for (const rk of rowKeys) {
    const line: (string | number | null)[] = rowsF.length ? rk.split(SEP) : ['Total']
    for (const ck of colKeys) {
      const g = groups.get(rk + '\u0002' + ck) ?? []
      for (let vi = 0; vi < vals.length; vi++) line.push(g.length ? cellVal(g, vi) : null)
    }
    if (colsF.length && def.grandTotals) {
      const all = filtered.filter((r) => keyOf(r, rowsF) === rk)
      for (let vi = 0; vi < vals.length; vi++) line.push(cellVal(all, vi))
    }
    out.push(line)
  }
  let totalRow: number | null = null
  if (def.grandTotals && rowsF.length) {
    const line: (string | number | null)[] = new Array(nRowF).fill(null)
    line[0] = 'Grand Total'
    for (const ck of colKeys) {
      const all = colsF.length ? filtered.filter((r) => keyOf(r, colsF) === ck) : filtered
      for (let vi = 0; vi < vals.length; vi++) line.push(cellVal(all, vi))
    }
    if (colsF.length) for (let vi = 0; vi < vals.length; vi++) line.push(cellVal(filtered, vi))
    totalRow = out.length
    out.push(line)
  }
  const width = Math.max(...out.map((r) => r.length))
  for (const r of out) while (r.length < width) r.push(null)
  for (let c = nRowF; c < width; c++) valueCols.push({ col: c, valueIndex: vals.length ? (c - nRowF) % vals.length : 0 })
  return { cells: out, headerRows, rowFieldCols: nRowF, totalRow, totalCol: colsF.length && def.grandTotals ? width - vals.length : null, valueCols }
}

function writePivot(wb: Workbook, sheetId: string, def: PivotDef, label: string, extraMeta?: () => void): string | null {
  const srcSheet = wb.sheets.find((s) => s.id === def.source.sheetId)
  if (!srcSheet) return 'The pivot table’s source sheet no longer exists.'
  const src = readPivotSource(wb, def.source.sheetId, def.source.range)
  const out = computePivot(src, def)
  const h = out.cells.length
  const w = out.cells[0]?.length ?? 0
  const target: Range = { r1: def.target.r, c1: def.target.c, r2: def.target.r + h - 1, c2: def.target.c + w - 1 }
  if (def.source.sheetId === sheetId && rangesIntersect(target, def.source.range)) return 'The pivot table would overlap its own source data. Choose another location.'
  const hfId = wb.hfId(sheetId)
  const old = def.lastOutput
  // refuse to overwrite unrelated content
  for (let r = target.r1; r <= target.r2; r++)
    for (let c = target.c1; c <= target.c2; c++) {
      const inOld = old && r >= old.r1 && r <= old.r2 && c >= old.c1 && c <= old.c2
      if (!inOld && !wb.isCellEmpty(sheetId, r, c)) return 'There is data where the pivot table would go. Choose an empty location.'
    }
  const styleUpdates: [string, CellStyle | undefined][] = []
  if (old) for (let r = old.r1; r <= old.r2; r++) for (let c = old.c1; c <= old.c2; c++) styleUpdates.push([cellKey(r, c), undefined])
  for (let i = 0; i < h; i++)
    for (let j = 0; j < w; j++) {
      const isHeader = i < out.headerRows
      const isTotal = i === out.totalRow || (out.totalCol !== null && j >= out.totalCol)
      const vc = out.valueCols.find((v) => v.col === j)
      const vdef = vc ? def.values[vc.valueIndex] : undefined
      const fmt = vdef && !isHeader && (vdef.agg === 'sum' || vdef.agg === 'min' || vdef.agg === 'max' || vdef.agg === 'average') ? src.formats[vdef.field] ?? (vdef.agg === 'average' ? '#,##0.00' : undefined) : undefined
      const st: CellStyle = {}
      if (isHeader) Object.assign(st, { bold: true, fill: '#d9ebe8', borders: i === out.headerRows - 1 ? { bottom: { style: 'thin', color: '#0f766e' } } : undefined })
      if (isTotal && !isHeader) Object.assign(st, { bold: true, fill: '#eef6f4' })
      if (i === out.totalRow) st.borders = { top: { style: 'thin', color: '#0f766e' } }
      if (fmt && !isDateFormat(fmt)) st.numFmt = fmt
      styleUpdates.push([cellKey(target.r1 + i, target.c1 + j), Object.keys(st).length ? st : undefined])
    }
  wb.run({
    label,
    dirty: [sheetId],
    metaScope: [sheetId],
    hfOps: () => {
      if (old) {
        const blank = Array.from({ length: old.r2 - old.r1 + 1 }, () => new Array(old.c2 - old.c1 + 1).fill(null))
        wb.hf.setCellContents({ sheet: hfId, row: old.r1, col: old.c1 }, blank)
      }
      wb.hf.setCellContents({ sheet: hfId, row: target.r1, col: target.c1 }, out.cells as RawCellContent[][])
    },
    metaOps: () => {
      const m = wb.meta(sheetId)
      for (const [k, v] of styleUpdates) {
        if (v) m.styles[k] = v
        else delete m.styles[k]
      }
      const next = { ...def, lastOutput: target }
      m.pivots = (m.pivots ?? []).some((p) => p.id === def.id) ? (m.pivots ?? []).map((p) => (p.id === def.id ? next : p)) : [...(m.pivots ?? []), next]
      extraMeta?.()
    },
  })
  return null
}

export function createPivot(wb: Workbook, sheetId: string, def: Omit<PivotDef, 'id' | 'name'> & { name?: string }): { id?: string; error?: string } {
  const id = uid('pvt')
  const n = wb.sheets.reduce((a, s) => a + (s.pivots?.length ?? 0), 0) + 1
  const full: PivotDef = { ...def, id, name: def.name ?? `PivotTable${n}` }
  if (!full.values.length) return { error: 'Add at least one value field.' }
  const err = writePivot(wb, sheetId, full, 'Insert PivotTable')
  return err ? { error: err } : { id }
}

export function updatePivot(wb: Workbook, sheetId: string, id: string, patch: Partial<PivotDef>): string | null {
  const def = wb.meta(sheetId).pivots?.find((p) => p.id === id)
  if (!def) return 'Pivot table not found.'
  return writePivot(wb, sheetId, { ...def, ...patch }, patch.rows || patch.cols || patch.values || patch.filters ? 'Edit PivotTable' : 'Refresh PivotTable')
}

export function refreshPivot(wb: Workbook, sheetId: string, id: string): string | null {
  return updatePivot(wb, sheetId, id, {})
}

export function refreshAllPivots(wb: Workbook, onlySource?: string): number {
  let n = 0
  for (const s of wb.sheets)
    for (const p of s.pivots ?? []) {
      if (onlySource && p.source.sheetId !== onlySource) continue
      if (!refreshPivot(wb, s.id, p.id)) n++
    }
  return n
}

export function removePivot(wb: Workbook, sheetId: string, id: string) {
  const def = wb.meta(sheetId).pivots?.find((p) => p.id === id)
  if (!def) return
  const old = def.lastOutput
  const hfId = wb.hfId(sheetId)
  wb.run({
    label: 'Delete PivotTable',
    dirty: [sheetId],
    metaScope: [sheetId],
    hfOps: old
      ? () => wb.hf.setCellContents({ sheet: hfId, row: old.r1, col: old.c1 }, Array.from({ length: old.r2 - old.r1 + 1 }, () => new Array(old.c2 - old.c1 + 1).fill(null)))
      : undefined,
    metaOps: () => {
      const m = wb.meta(sheetId)
      if (old) for (let r = old.r1; r <= old.r2; r++) for (let c = old.c1; c <= old.c2; c++) delete m.styles[cellKey(r, c)]
      m.pivots = (m.pivots ?? []).filter((p) => p.id !== id)
    },
  })
}

export function pivotAt(wb: Workbook, sheetId: string, r: number, c: number): PivotDef | undefined {
  return wb.meta(sheetId).pivots?.find((p) => p.lastOutput && r >= p.lastOutput.r1 && r <= p.lastOutput.r2 && c >= p.lastOutput.c1 && c <= p.lastOutput.c2)
}

/** Distinct labels of a field, for the filter picker. */
export function fieldValues(wb: Workbook, def: Pick<PivotDef, 'source' | 'dateGroup'>, field: string): string[] {
  const src = readPivotSource(wb, def.source.sheetId, def.source.range)
  const i = src.fields.indexOf(field)
  if (i < 0) return []
  const full = { ...def, rows: [], cols: [], values: [], filters: [], sort: 'asc' as const, grandTotals: false, autoRefresh: false, id: '', name: '', target: { r: 0, c: 0 } }
  return [...new Set(src.rows.map((r) => label(r[i], field, src, full)))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
}
