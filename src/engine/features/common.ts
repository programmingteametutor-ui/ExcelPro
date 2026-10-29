import { DetailedCellError } from 'hyperformula'
import type { CellValue } from 'hyperformula'
import type { Range } from '../../model/types'
import type { Workbook } from '../workbook'

export type Scalar = number | string | boolean | null | { error: string }

export function toScalar(v: CellValue): Scalar {
  if (v instanceof DetailedCellError) return { error: v.value }
  if (v === undefined) return null
  return v as Scalar
}

/** Excel sort order: numbers < text < booleans < errors; blanks are always last. */
export function compareValues(a: Scalar, b: Scalar): number {
  const rank = (v: Scalar) => (v === null || v === '' ? 4 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3)
  const ra = rank(a)
  const rb = rank(b)
  if (ra !== rb) return ra - rb
  if (typeof a === 'number' && typeof b === 'number') return a - b
  if (typeof a === 'string' && typeof b === 'string') return a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true })
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b)
  return 0
}

export const isBlank = (v: Scalar) => v === null || v === ''

/** Expands from a cell to the contiguous block of non-empty cells around it (Excel's "current region"). */
export function currentRegion(wb: Workbook, sheetId: string, r: number, c: number): Range {
  const hfId = wb.hfId(sheetId)
  const dims = wb.hf.getSheetDimensions(hfId)
  const filled = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr < dims.height && cc < dims.width && !wb.isCellEmpty(sheetId, rr, cc)
  let rg: Range = { r1: r, c1: c, r2: r, c2: c }
  let changed = true
  let guard = 0
  while (changed && guard++ < 10_000) {
    changed = false
    // grow each edge while the adjacent row/column (including diagonals) has content
    const rowHas = (rr: number) => {
      for (let cc = rg.c1 - 1; cc <= rg.c2 + 1; cc++) if (filled(rr, cc)) return true
      return false
    }
    const colHas = (cc: number) => {
      for (let rr = rg.r1 - 1; rr <= rg.r2 + 1; rr++) if (filled(rr, cc)) return true
      return false
    }
    if (rg.r1 > 0 && rowHas(rg.r1 - 1)) (rg = { ...rg, r1: rg.r1 - 1 }), (changed = true)
    if (rowHas(rg.r2 + 1)) (rg = { ...rg, r2: rg.r2 + 1 }), (changed = true)
    if (rg.c1 > 0 && colHas(rg.c1 - 1)) (rg = { ...rg, c1: rg.c1 - 1 }), (changed = true)
    if (colHas(rg.c2 + 1)) (rg = { ...rg, c2: rg.c2 + 1 }), (changed = true)
  }
  return rg
}

/** Reads a range as display texts and values (clamped to the used area). */
export function readRange(wb: Workbook, sheetId: string, rg: Range): { values: Scalar[][]; texts: string[][] } {
  const values: Scalar[][] = []
  const texts: string[][] = []
  for (let r = rg.r1; r <= rg.r2; r++) {
    const vr: Scalar[] = []
    const tr: string[] = []
    for (let c = rg.c1; c <= rg.c2; c++) {
      vr.push(toScalar(wb.getValue(sheetId, r, c)))
      tr.push(wb.getDisplay(sheetId, r, c).text)
    }
    values.push(vr)
    texts.push(tr)
  }
  return { values, texts }
}

export function compareOp(op: string, v: number, a: number, b?: number): boolean {
  switch (op) {
    case 'between':
      return v >= Math.min(a, b ?? a) && v <= Math.max(a, b ?? a)
    case 'notBetween':
      return v < Math.min(a, b ?? a) || v > Math.max(a, b ?? a)
    case 'eq':
      return v === a
    case 'ne':
      return v !== a
    case 'gt':
      return v > a
    case 'lt':
      return v < a
    case 'gte':
      return v >= a
    case 'lte':
      return v <= a
  }
  return false
}

export function clampToUsed(wb: Workbook, sheetId: string, rg: Range): Range {
  const d = wb.hf.getSheetDimensions(wb.hfId(sheetId))
  return { r1: rg.r1, c1: rg.c1, r2: Math.max(rg.r1, Math.min(rg.r2, d.height - 1)), c2: Math.max(rg.c1, Math.min(rg.c2, d.width - 1)) }
}

let counter = 0
export const uid = (p: string) => `${p}${Date.now().toString(36)}${(counter++).toString(36)}`
