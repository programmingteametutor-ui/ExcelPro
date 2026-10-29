/**
 * Data cleaning and what-if tools: remove duplicates, trim, change case, text to columns,
 * convert text to numbers, Flash Fill (pattern inference from examples), Goal Seek.
 */
import type { RawCellContent } from 'hyperformula'
import { cellKey } from '../../model/address'
import type { CellStyle, Range } from '../../model/types'
import { parseInput, wouldConvert, type RawContent } from '../inputParser'
import type { Workbook } from '../workbook'
import { clampToUsed, toScalar } from './common'

/** Removes rows whose values in `cols` repeat an earlier row. Returns the number removed. */
export function removeDuplicates(wb: Workbook, sheetId: string, range: Range, cols: number[], hasHeader: boolean): number {
  const rg = clampToUsed(wb, sheetId, range)
  const top = rg.r1 + (hasHeader ? 1 : 0)
  const keyCols = cols.length ? cols : Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i)
  const seen = new Set<string>()
  const keep: { raw: RawContent[]; styles: (CellStyle | undefined)[] }[] = []
  const meta = wb.meta(sheetId)
  let removed = 0
  for (let r = top; r <= rg.r2; r++) {
    const k = keyCols
      .map((c) => {
        const v = toScalar(wb.getValue(sheetId, r, c))
        return typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + JSON.stringify(v)
      })
      .join('\u0001')
    if (seen.has(k)) {
      removed++
      continue
    }
    seen.add(k)
    const raw: RawContent[] = []
    const styles: (CellStyle | undefined)[] = []
    for (let c = rg.c1; c <= rg.c2; c++) {
      raw.push(wb.getRaw(sheetId, r, c))
      styles.push(meta.styles[cellKey(r, c)])
    }
    keep.push({ raw, styles })
  }
  if (!removed) return 0
  const width = rg.c2 - rg.c1 + 1
  const block: RawContent[][] = []
  const styleChanges = new Map<string, CellStyle | undefined>()
  for (let i = 0; i <= rg.r2 - top; i++) {
    const row = keep[i]
    block.push(row ? row.raw : new Array(width).fill(null))
    for (let j = 0; j < width; j++) styleChanges.set(cellKey(top + i, rg.c1 + j), row?.styles[j])
  }
  const hfId = wb.hfId(sheetId)
  wb.runStyles('Remove duplicates', sheetId, styleChanges, () => wb.hf.setCellContents({ sheet: hfId, row: top, col: rg.c1 }, block as RawCellContent[][]), { sheetId, range: { ...rg, r2: top + keep.length - 1 } })
  return removed
}

export type TextTransform = 'trim' | 'upper' | 'lower' | 'proper' | 'toNumber' | 'clean'

const proper = (s: string) => s.toLowerCase().replace(/(^|[^\p{L}'])(\p{L})/gu, (_m, a: string, b: string) => a + b.toUpperCase())

/** Applies a text transform to constant text cells in the ranges (formulas are left alone). */
export function transformText(wb: Workbook, sheetId: string, ranges: Range[], kind: TextTransform): number {
  const entries: { r: number; c: number; input: string }[] = []
  for (const rg0 of ranges) {
    const rg = clampToUsed(wb, sheetId, rg0)
    for (let r = rg.r1; r <= rg.r2; r++)
      for (let c = rg.c1; c <= rg.c2; c++) {
        const raw = wb.getRaw(sheetId, r, c)
        if (typeof raw !== 'string' || raw.startsWith('=')) continue
        const text = raw.startsWith("'") ? raw.slice(1) : raw
        let next: string
        switch (kind) {
          case 'trim':
            next = text.replace(/[\s ]+/g, ' ').trim()
            break
          case 'clean':
            // eslint-disable-next-line no-control-regex
            next = text.replace(/[\u0000-\u001f\u007f]/g, '')
            break
          case 'upper':
            next = text.toUpperCase()
            break
          case 'lower':
            next = text.toLowerCase()
            break
          case 'proper':
            next = proper(text)
            break
          case 'toNumber': {
            const t = text.trim().replace(/^[ \s]+|[ \s]+$/g, '')
            const p = parseInput(t)
            if (typeof p.raw !== 'number') continue
            entries.push({ r, c, input: t })
            continue
          }
        }
        if (next === text) continue
        entries.push({ r, c, input: wouldConvert(next) ? `'${next}` : next })
      }
  }
  if (entries.length) wb.setInputs(sheetId, entries, { trim: 'Trim spaces', clean: 'Clean text', upper: 'UPPERCASE', lower: 'lowercase', proper: 'Proper Case', toNumber: 'Convert to numbers' }[kind])
  return entries.length
}

/** Splits one column of text into several columns. */
export function textToColumns(wb: Workbook, sheetId: string, range: Range, delimiter: string | { widths: number[] }): string | null {
  const rg = clampToUsed(wb, sheetId, range)
  if (rg.c1 !== rg.c2) return 'Select a single column to split.'
  const rows: string[][] = []
  let width = 1
  for (let r = rg.r1; r <= rg.r2; r++) {
    const text = wb.getDisplay(sheetId, r, rg.c1).text
    let parts: string[]
    if (typeof delimiter === 'string') {
      const d = delimiter === '\\t' ? '\t' : delimiter
      parts = d === ' ' ? text.trim().split(/\s+/) : text.split(d)
    } else {
      parts = []
      let pos = 0
      for (const w of delimiter.widths) {
        parts.push(text.slice(pos, pos + w).trim())
        pos += w
      }
      parts.push(text.slice(pos).trim())
    }
    rows.push(parts)
    width = Math.max(width, parts.length)
  }
  // refuse to overwrite data to the right
  for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1 + 1; c < rg.c1 + width; c++) if (!wb.isCellEmpty(sheetId, r, c)) return 'There is already data in the columns to the right. Insert empty columns first.'
  const entries: { r: number; c: number; input: string }[] = []
  rows.forEach((parts, i) => {
    for (let j = 0; j < width; j++) entries.push({ r: rg.r1 + i, c: rg.c1 + j, input: (parts[j] ?? '').trim() })
  })
  wb.setInputs(sheetId, entries, 'Text to columns')
  return null
}

// ----------------------------------------------------------------- Flash Fill
type Transform = (s: string) => string | null

const WORD = /[\p{L}\p{N}]+/gu

function candidates(): { name: string; fn: Transform }[] {
  const out: { name: string; fn: Transform }[] = []
  const words = (s: string) => s.match(WORD) ?? []
  for (let i = 0; i < 4; i++) {
    out.push({ name: `word ${i + 1}`, fn: (s) => words(s)[i] ?? null })
    out.push({ name: `word -${i + 1}`, fn: (s) => words(s).at(-1 - i) ?? null })
  }
  for (const d of [' ', ',', ';', '-', '/', '@', '.', '_', '|']) {
    for (let i = 0; i < 4; i++) {
      out.push({ name: `split ${d} ${i}`, fn: (s) => { const p = s.split(d); return p.length > i ? p[i].trim() : null } })
      out.push({ name: `split ${d} -${i}`, fn: (s) => { const p = s.split(d); return p.length > i ? p[p.length - 1 - i].trim() : null } })
    }
  }
  out.push({ name: 'initials', fn: (s) => words(s).map((w) => w[0].toUpperCase()).join('') })
  out.push({ name: 'initials.', fn: (s) => words(s).map((w) => w[0].toUpperCase() + '.').join('') })
  out.push({ name: 'digits', fn: (s) => s.replace(/\D/g, '') || null })
  out.push({ name: 'letters', fn: (s) => s.replace(/[^\p{L}]/gu, '') || null })
  out.push({ name: 'last, first', fn: (s) => { const w = words(s); return w.length >= 2 ? `${w.at(-1)}, ${w[0]}` : null } })
  out.push({ name: 'first last', fn: (s) => { const p = s.split(',').map((x) => x.trim()); return p.length === 2 ? `${p[1]} ${p[0]}` : null } })
  out.push({ name: 'first initial + last', fn: (s) => { const w = words(s); return w.length >= 2 ? `${(w[0] ?? "")[0]}. ${w.at(-1)}` : null } })
  for (let n = 1; n <= 6; n++) {
    out.push({ name: `left ${n}`, fn: (s) => (s.length >= n ? s.slice(0, n) : null) })
    out.push({ name: `right ${n}`, fn: (s) => (s.length >= n ? s.slice(-n) : null) })
  }
  return out
}

const CASES: { name: string; fn: (s: string) => string }[] = [
  { name: 'as is', fn: (s) => s },
  { name: 'upper', fn: (s) => s.toUpperCase() },
  { name: 'lower', fn: (s) => s.toLowerCase() },
  { name: 'proper', fn: proper },
]

/** Finds a transform that reproduces every example (source text -> typed result). */
export function inferFlashFill(examples: { source: string[]; result: string }[]): ((source: string[]) => string | null) | null {
  if (!examples.length) return null
  const cands = candidates()
  const nSrc = examples[0].source.length
  // single-source transforms
  for (let si = 0; si < nSrc; si++)
    for (const cand of cands)
      for (const cs of CASES) {
        if (examples.every((ex) => { const v = cand.fn(ex.source[si] ?? ''); return v !== null && cs.fn(v) === ex.result })) {
          return (src) => { const v = cand.fn(src[si] ?? ''); return v === null ? null : cs.fn(v) }
        }
      }
  // concatenation of two transforms with a literal separator (e.g. "first.last@company.com" style)
  const first = examples[0]
  const pieces: { si: number; cand: (typeof cands)[number]; cs: (typeof CASES)[number]; value: string }[] = []
  for (let si = 0; si < nSrc; si++)
    for (const cand of cands)
      for (const cs of CASES) {
        const v = cand.fn(first.source[si] ?? '')
        if (v && v.length >= 1 && first.result.includes(cs.fn(v))) pieces.push({ si, cand, cs, value: cs.fn(v) })
      }
  for (const a of pieces)
    for (const b of pieces) {
      const ia = first.result.indexOf(a.value)
      if (ia < 0) continue
      const rest = first.result.slice(ia + a.value.length)
      const ib = rest.indexOf(b.value)
      if (ib < 0) continue
      const prefix = first.result.slice(0, ia)
      const mid = rest.slice(0, ib)
      const suffix = rest.slice(ib + b.value.length)
      const fn = (src: string[]) => {
        const va = a.cand.fn(src[a.si] ?? '')
        const vb = b.cand.fn(src[b.si] ?? '')
        return va === null || vb === null ? null : prefix + a.cs.fn(va) + mid + b.cs.fn(vb) + suffix
      }
      if (examples.every((ex) => fn(ex.source) === ex.result)) return fn
    }
  return null
}

/**
 * Flash Fill for column `col`: uses the filled cells in that column as examples, with the
 * columns to the left (within the current region) as sources, and fills the empty cells below.
 */
export function flashFill(wb: Workbook, sheetId: string, col: number, r1: number, r2: number): { filled: number; error?: string } {
  const srcCols: number[] = []
  for (let c = col - 1; c >= 0 && srcCols.length < 4; c--) {
    let any = false
    for (let r = r1; r <= Math.min(r2, r1 + 20); r++) if (!wb.isCellEmpty(sheetId, r, c)) any = true
    if (!any) break
    srcCols.unshift(c)
  }
  if (!srcCols.length) return { filled: 0, error: 'Flash Fill needs data in the column to the left.' }
  const examples: { source: string[]; result: string }[] = []
  const targets: number[] = []
  for (let r = r1; r <= r2; r++) {
    const source = srcCols.map((c) => wb.getDisplay(sheetId, r, c).text)
    if (source.every((s) => !s)) continue
    if (!wb.isCellEmpty(sheetId, r, col)) examples.push({ source, result: wb.getDisplay(sheetId, r, col).text })
    else targets.push(r)
  }
  if (!examples.length) return { filled: 0, error: 'Type one or two examples of the result first, then run Flash Fill.' }
  const fn = inferFlashFill(examples)
  if (!fn) return { filled: 0, error: 'Flash Fill couldn’t find a pattern. Type another example and try again.' }
  const entries = targets.flatMap((r) => {
    const v = fn(srcCols.map((c) => wb.getDisplay(sheetId, r, c).text))
    return v === null ? [] : [{ r, c: col, input: wouldConvert(v) ? `'${v}` : v }]
  })
  if (entries.length) wb.setInputs(sheetId, entries, 'Flash Fill')
  return { filled: entries.length }
}

// ----------------------------------------------------------------- Goal Seek
/**
 * Changes `changing` (a constant cell) until the formula in `setCell` equals `target`.
 * Uses the secant method with bisection fallback; the final value is written as one undoable edit.
 */
export function goalSeek(
  wb: Workbook,
  sheetId: string,
  setCell: { r: number; c: number },
  target: number,
  changing: { r: number; c: number },
): { ok: boolean; value?: number; result?: number; message: string } {
  if (!wb.isFormula(sheetId, setCell.r, setCell.c)) return { ok: false, message: 'The “Set cell” must contain a formula.' }
  if (wb.isFormula(sheetId, changing.r, changing.c)) return { ok: false, message: 'The “By changing” cell must contain a value, not a formula.' }
  const hf = wb.hf
  const hfId = wb.hfId(sheetId)
  const original = wb.getRaw(sheetId, changing.r, changing.c)
  const x0 = typeof original === 'number' ? original : 0
  let writes = 0
  const f = (x: number): number => {
    hf.setCellContents({ sheet: hfId, row: changing.r, col: changing.c }, x)
    writes++
    const v = hf.getCellValue({ sheet: hfId, row: setCell.r, col: setCell.c })
    return typeof v === 'number' ? v - target : NaN
  }
  let a = x0
  let b = x0 === 0 ? 1 : x0 * 1.1
  let fa = f(a)
  let fb = f(b)
  let best = Math.abs(fa) < Math.abs(fb) ? a : b
  let bestErr = Math.min(Math.abs(fa), Math.abs(fb))
  if (!isFinite(bestErr)) bestErr = Infinity
  const tol = Math.max(1e-9, Math.abs(target) * 1e-9)
  for (let i = 0; i < 100 && bestErr > tol; i++) {
    let x: number
    if (isFinite(fa) && isFinite(fb) && fb !== fa) x = b - (fb * (b - a)) / (fb - fa)
    else x = b + (b - a || 1) * 2
    if (!isFinite(x)) x = (a + b) / 2
    const fx = f(x)
    if (isFinite(fx) && Math.abs(fx) < bestErr) {
      bestErr = Math.abs(fx)
      best = x
    }
    a = b
    fa = fb
    b = x
    fb = fx
  }
  // Each probe was recorded as an engine undo entry; undoing them all restores the original
  // value and keeps the engine's undo stack aligned with the workbook's command history.
  for (let i = 0; i < writes; i++) hf.undo()
  void original
  if (!(bestErr <= Math.max(1e-6, Math.abs(target) * 1e-6))) return { ok: false, message: 'Goal Seek may not have found a solution.' }
  const value = +best.toPrecision(15)
  wb.setInputs(sheetId, [{ r: changing.r, c: changing.c, input: String(value) }], 'Goal Seek')
  const result = wb.getValue(sheetId, setCell.r, setCell.c) as number
  return { ok: true, value, result, message: `Goal Seek found a solution: ${value}.` }
}
