import type { CellAddr, Range } from './types'

export function colToLetters(c: number): string {
  let s = ''
  let n = c + 1
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

export function lettersToCol(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

export function addrToA1({ r, c }: CellAddr): string {
  return `${colToLetters(c)}${r + 1}`
}

export function parseA1(ref: string): CellAddr | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(ref.trim())
  if (!m) return null
  const c = lettersToCol(m[1])
  const r = parseInt(m[2], 10) - 1
  if (r < 0 || c < 0) return null
  return { r, c }
}

export function normRange(a: CellAddr, b: CellAddr): Range {
  return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) }
}

export function rangeToA1(rg: Range): string {
  const a = addrToA1({ r: rg.r1, c: rg.c1 })
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return a
  return `${a}:${addrToA1({ r: rg.r2, c: rg.c2 })}`
}

/** Parses "A1", "A1:C5", "$A$1:$C$5". Returns null if invalid. */
export function parseRangeA1(text: string): Range | null {
  const parts = text.trim().split(':')
  if (parts.length === 1) {
    const a = parseA1(parts[0])
    return a ? { r1: a.r, c1: a.c, r2: a.r, c2: a.c } : null
  }
  if (parts.length !== 2) return null
  const a = parseA1(parts[0])
  const b = parseA1(parts[1])
  if (!a || !b) return null
  return normRange(a, b)
}

export function inRange(rg: Range, r: number, c: number): boolean {
  return r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2
}

export function rangesIntersect(a: Range, b: Range): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2
}

export function rangeSize(rg: Range): number {
  return (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1)
}

export const cellKey = (r: number, c: number) => `${r},${c}`
export function parseKey(k: string): CellAddr {
  const i = k.indexOf(',')
  return { r: +k.slice(0, i), c: +k.slice(i + 1) }
}

/** Quote a sheet name for use in formulas when it needs it. */
export function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`
}
