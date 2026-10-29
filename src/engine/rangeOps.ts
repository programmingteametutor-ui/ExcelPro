import type { Range } from '../model/types'
import { cellKey, parseKey } from '../model/address'

/**
 * Adjusts a range for inserted (delta > 0) or deleted (delta < 0) rows/columns starting at `at`.
 * Returns null when the whole range was deleted.
 */
export function shiftRange(g: Range, axis: 'row' | 'col', at: number, delta: number): Range | null {
  const isRow = axis === 'row'
  const a1 = isRow ? g.r1 : g.c1
  const a2 = isRow ? g.r2 : g.c2
  let n1: number
  let n2: number
  if (delta > 0) {
    n1 = a1 >= at ? a1 + delta : a1
    n2 = a2 >= at ? a2 + delta : a2
  } else {
    const end = at - delta // exclusive
    const map = (i: number, isStart: boolean) => (i < at ? i : i >= end ? i + delta : isStart ? at : at - 1)
    n1 = map(a1, true)
    n2 = map(a2, false)
    if (n2 < n1) return null
  }
  return isRow ? { ...g, r1: n1, r2: n2 } : { ...g, c1: n1, c2: n2 }
}

export function shiftIndex(i: number, at: number, delta: number): number | null {
  if (i < at) return i
  if (delta < 0 && i < at - delta) return null
  return i + delta
}

export function shiftKeyed<T>(map: Record<string, T> | undefined, axis: 'row' | 'col', at: number, delta: number): Record<string, T> | undefined {
  if (!map) return map
  const out: Record<string, T> = {}
  for (const [k, v] of Object.entries(map)) {
    const { r, c } = parseKey(k)
    const n = shiftIndex(axis === 'row' ? r : c, at, delta)
    if (n === null) continue
    out[axis === 'row' ? cellKey(n, c) : cellKey(r, n)] = v
  }
  return out
}

export function rangeCells(rg: Range): number {
  return (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1)
}
