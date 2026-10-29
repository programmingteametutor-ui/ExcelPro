import { normRange, rangesIntersect } from '../model/address'
import type { CellAddr, Range } from '../model/types'

export interface Selection {
  anchor: CellAddr // active cell
  focus: CellAddr // opposite corner of the primary range
  extra: Range[] // additional ranges (Ctrl+click)
}

export interface EditState {
  text: string
  mode: 'enter' | 'edit' // enter: arrows commit & move; edit: arrows move the caret
  origin: 'cell' | 'bar'
}

export const primaryRange = (s: Selection): Range => normRange(s.anchor, s.focus)
export const allRanges = (s: Selection): Range[] => [...s.extra, primaryRange(s)]
export const single = (a: CellAddr): Selection => ({ anchor: a, focus: a, extra: [] })

/** Grows a range until it fully contains every merged block it touches. */
export function expandForMerges(rg: Range, merges: Range[]): Range {
  let out = { ...rg }
  let changed = true
  let guard = 0
  while (changed && guard++ < 50) {
    changed = false
    for (const m of merges) {
      if (!rangesIntersect(m, out)) continue
      const n = { r1: Math.min(out.r1, m.r1), c1: Math.min(out.c1, m.c1), r2: Math.max(out.r2, m.r2), c2: Math.max(out.c2, m.c2) }
      if (n.r1 !== out.r1 || n.c1 !== out.c1 || n.r2 !== out.r2 || n.c2 !== out.c2) {
        out = n
        changed = true
      }
    }
  }
  return out
}

/** Whether the caret position in a formula accepts a clicked cell reference. */
export function canInsertRef(text: string, caret: number): boolean {
  if (!text.startsWith('=')) return false
  const before = text.slice(0, caret).trimEnd()
  if (before === '=') return true
  return /[=(,+\-*/^&<>:;]$/.test(before)
}

/** If the text before the caret ends with a reference we inserted, return its start index. */
export function trailingRefStart(text: string, caret: number): number | null {
  const m = /(?:(?:'[^']+'|[A-Za-z0-9_.]+)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/.exec(text.slice(0, caret))
  if (!m) return null
  const start = caret - m[0].length
  return canInsertRef(text, start) ? start : null
}
