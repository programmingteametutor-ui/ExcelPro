/**
 * Phase 2 data features: sorting, AutoFilter, data validation, named ranges, notes, hyperlinks,
 * freeze panes. Every change is an undoable Workbook command.
 */
import type { RawCellContent } from 'hyperformula'
import { cellKey, colToLetters, inRange, parseRangeA1, rangesIntersect } from '../../model/address'
import type { AutoFilter, CellStyle, FilterCriterion, Range, Validation } from '../../model/types'
import { shiftFormula } from '../formulaText'
import { parseInput, type RawContent } from '../inputParser'
import { isDateFormat } from '../numberFormat'
import type { Workbook } from '../workbook'
import { clampToUsed, compareOp, compareValues, isBlank, toScalar, uid, type Scalar } from './common'

// ===================================================================== sort
export interface SortKey {
  col: number // absolute column index
  desc: boolean
}

export function sortRange(wb: Workbook, sheetId: string, range: Range, keys: SortKey[], hasHeader: boolean): string | null {
  const rg = clampToUsed(wb, sheetId, range)
  const top = rg.r1 + (hasHeader ? 1 : 0)
  if (top >= rg.r2) return null
  if (!keys.length) return 'Choose at least one column to sort by.'
  const meta = wb.meta(sheetId)
  if (meta.merges.some((m) => rangesIntersect(m, { ...rg, r1: top }))) return 'Ranges containing merged cells cannot be sorted. Unmerge them first.'
  const rows: { idx: number; raw: RawContent[]; styles: (CellStyle | undefined)[]; keys: Scalar[] }[] = []
  for (let r = top; r <= rg.r2; r++) {
    const raw: RawContent[] = []
    const styles: (CellStyle | undefined)[] = []
    for (let c = rg.c1; c <= rg.c2; c++) {
      raw.push(wb.getRaw(sheetId, r, c))
      styles.push(meta.styles[cellKey(r, c)])
    }
    rows.push({ idx: r, raw, styles, keys: keys.map((k) => toScalar(wb.getValue(sheetId, r, k.col))) })
  }
  const sorted = [...rows].sort((a, b) => {
    for (let i = 0; i < keys.length; i++) {
      const va = a.keys[i]
      const vb = b.keys[i]
      // blanks stay last regardless of direction
      if (isBlank(va) !== isBlank(vb)) return isBlank(va) ? 1 : -1
      const c = compareValues(va, vb)
      if (c) return keys[i].desc ? -c : c
    }
    return a.idx - b.idx
  })
  const block: RawContent[][] = []
  const styleChanges = new Map<string, CellStyle | undefined>()
  sorted.forEach((row, j) => {
    const target = top + j
    block.push(row.raw.map((v) => (typeof v === 'string' && v.startsWith('=') ? shiftFormula(v, target - row.idx, 0) : v)))
    row.styles.forEach((s, k) => styleChanges.set(cellKey(target, rg.c1 + k), s))
  })
  const hfId = wb.hfId(sheetId)
  wb.runStyles(keys.length > 1 ? 'Custom sort' : keys[0].desc ? 'Sort Z to A' : 'Sort A to Z', sheetId, styleChanges, () => {
    wb.hf.setCellContents({ sheet: hfId, row: top, col: rg.c1 }, block as RawCellContent[][])
  }, { sheetId, range: rg })
  if (meta.autoFilter && rangesIntersect(meta.autoFilter.range, rg)) reapplyFilter(wb, sheetId)
  return null
}

/** Whether the first row looks like a header (text over numbers/dates), like Excel's guess. */
export function guessHeader(wb: Workbook, sheetId: string, rg: Range): boolean {
  if (rg.r2 <= rg.r1) return false
  let headerText = 0
  let bodyNonText = 0
  for (let c = rg.c1; c <= rg.c2; c++) {
    const h = wb.getValue(sheetId, rg.r1, c)
    const b = wb.getValue(sheetId, rg.r1 + 1, c)
    if (typeof h === 'string' && h !== '') headerText++
    if (typeof b !== 'string') bodyNonText++
    else if (typeof h === 'string' && typeof b === 'string' && wb.getStyle(sheetId, rg.r1, c)?.bold && !wb.getStyle(sheetId, rg.r1 + 1, c)?.bold) bodyNonText++
  }
  const width = rg.c2 - rg.c1 + 1
  return headerText === width && bodyNonText > 0
}

// ===================================================================== filter
function matchCondition(crit: Extract<FilterCriterion, { kind: 'condition' }>, value: Scalar, text: string, stats: { avg: number; threshold: number }): boolean {
  const t = text.toLowerCase()
  const a = (crit.a ?? '').toLowerCase()
  const num = typeof value === 'number' ? value : null
  const parseNum = (s: string | undefined) => {
    if (s === undefined) return NaN
    const p = parseInput(s)
    return typeof p.raw === 'number' ? p.raw : NaN
  }
  switch (crit.op) {
    case 'contains':
      return t.includes(a)
    case 'notContains':
      return !t.includes(a)
    case 'equals':
      return num !== null && !isNaN(parseNum(crit.a)) ? num === parseNum(crit.a) : t === a
    case 'notEquals':
      return num !== null && !isNaN(parseNum(crit.a)) ? num !== parseNum(crit.a) : t !== a
    case 'begins':
      return t.startsWith(a)
    case 'ends':
      return t.endsWith(a)
    case 'gt':
      return num !== null && num > parseNum(crit.a)
    case 'gte':
      return num !== null && num >= parseNum(crit.a)
    case 'lt':
      return num !== null && num < parseNum(crit.a)
    case 'lte':
      return num !== null && num <= parseNum(crit.a)
    case 'between':
      return num !== null && compareOp('between', num, parseNum(crit.a), parseNum(crit.b))
    case 'top':
      return num !== null && num >= stats.threshold
    case 'bottom':
      return num !== null && num <= stats.threshold
    case 'aboveAverage':
      return num !== null && num > stats.avg
    case 'belowAverage':
      return num !== null && num < stats.avg
  }
  return true
}

export function computeFilterHidden(wb: Workbook, sheetId: string, af: AutoFilter): number[] {
  const hidden: number[] = []
  const cols = Object.keys(af.criteria).map(Number)
  if (!cols.length) return hidden
  const r1 = af.range.r1 + 1
  const r2 = clampToUsed(wb, sheetId, af.range).r2
  const stats: Record<number, { avg: number; threshold: number }> = {}
  for (const c of cols) {
    const crit = af.criteria[c]
    if (crit.kind !== 'condition' || !['top', 'bottom', 'aboveAverage', 'belowAverage'].includes(crit.op)) continue
    const nums: number[] = []
    for (let r = r1; r <= r2; r++) {
      const v = wb.getValue(sheetId, r, c)
      if (typeof v === 'number') nums.push(v)
    }
    nums.sort((x, y) => x - y)
    const n = Math.max(1, Math.round(Number(crit.a) || 10))
    stats[c] = {
      avg: nums.reduce((s, x) => s + x, 0) / (nums.length || 1),
      threshold: crit.op === 'top' ? nums[Math.max(0, nums.length - n)] ?? Infinity : nums[Math.min(nums.length - 1, n - 1)] ?? -Infinity,
    }
  }
  for (let r = r1; r <= r2; r++) {
    let keep = true
    for (const c of cols) {
      const crit = af.criteria[c]
      const d = wb.getDisplay(sheetId, r, c)
      if (crit.kind === 'values') {
        keep = d.kind === 'empty' || d.text === '' ? crit.blanks : crit.values.includes(d.text)
      } else keep = matchCondition(crit, toScalar(wb.getValue(sheetId, r, c)), d.text, stats[c] ?? { avg: 0, threshold: 0 })
      if (!keep) break
    }
    if (!keep) hidden.push(r)
  }
  return hidden
}

/** Turns AutoFilter on for a range (header row first) or off when range is null. */
export function setAutoFilter(wb: Workbook, sheetId: string, range: Range | null) {
  wb.run({
    label: range ? 'Filter' : 'Remove filter',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      m.autoFilter = range ? { range, criteria: {} } : undefined
      m.filterHidden = []
    },
  })
}

export function setFilterCriterion(wb: Workbook, sheetId: string, col: number, crit: FilterCriterion | null) {
  const m0 = wb.meta(sheetId)
  if (!m0.autoFilter) return
  const af: AutoFilter = { range: m0.autoFilter.range, criteria: { ...m0.autoFilter.criteria } }
  if (crit) af.criteria[col] = crit
  else delete af.criteria[col]
  // grow the range to include data typed below it since the filter was created
  af.range = { ...af.range, r2: Math.max(af.range.r2, clampToUsed(wb, sheetId, { ...af.range, r2: 1_048_575 }).r2) }
  const hidden = computeFilterHidden(wb, sheetId, af)
  wb.run({
    label: crit ? 'Filter' : 'Clear filter',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      m.autoFilter = af
      m.filterHidden = hidden
    },
  })
}

export function clearAllFilters(wb: Workbook, sheetId: string) {
  const m0 = wb.meta(sheetId)
  if (!m0.autoFilter) return
  wb.run({
    label: 'Clear filters',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      m.autoFilter = { range: m0.autoFilter!.range, criteria: {} }
      m.filterHidden = []
    },
  })
}

export function reapplyFilter(wb: Workbook, sheetId: string) {
  const m0 = wb.meta(sheetId)
  if (!m0.autoFilter) return
  const hidden = computeFilterHidden(wb, sheetId, m0.autoFilter)
  wb.run({ label: 'Reapply filter', dirty: [sheetId], metaScope: [sheetId], metaOps: () => (wb.meta(sheetId).filterHidden = hidden) })
}

/** Distinct display values of a filter column, among rows visible under the OTHER columns' filters. */
export function filterValues(wb: Workbook, sheetId: string, col: number): { values: string[]; hasBlanks: boolean } {
  const af = wb.meta(sheetId).autoFilter
  if (!af) return { values: [], hasBlanks: false }
  const others: AutoFilter = { range: af.range, criteria: { ...af.criteria } }
  delete others.criteria[col]
  const hidden = new Set(computeFilterHidden(wb, sheetId, others))
  const r2 = Math.max(af.range.r2, clampToUsed(wb, sheetId, { ...af.range, r2: 1_048_575 }).r2)
  const seen = new Map<string, Scalar>()
  let hasBlanks = false
  for (let r = af.range.r1 + 1; r <= r2; r++) {
    if (hidden.has(r)) continue
    const d = wb.getDisplay(sheetId, r, col)
    if (d.kind === 'empty' || d.text === '') hasBlanks = true
    else if (!seen.has(d.text)) seen.set(d.text, toScalar(wb.getValue(sheetId, r, col)))
  }
  const values = [...seen.entries()].sort((a, b) => compareValues(a[1], b[1])).map(([t]) => t)
  return { values, hasBlanks }
}

// ===================================================================== data validation
export function validationAt(wb: Workbook, sheetId: string, r: number, c: number): Validation | undefined {
  const list = wb.meta(sheetId).validations
  if (!list) return undefined
  for (let i = list.length - 1; i >= 0; i--) if (inRange(list[i].range, r, c)) return list[i]
  return undefined
}

export function addValidation(wb: Workbook, sheetId: string, v: Omit<Validation, 'id'> & { id?: string }) {
  const id = v.id ?? uid('dv')
  wb.run({
    label: 'Data validation',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      // a new rule replaces older rules on exactly the same range
      const rest = (m.validations ?? []).filter((x) => x.id !== id && !(x.range.r1 === v.range.r1 && x.range.c1 === v.range.c1 && x.range.r2 === v.range.r2 && x.range.c2 === v.range.c2))
      m.validations = [...rest, { ...v, id } as Validation]
    },
  })
  return id
}

export function removeValidations(wb: Workbook, sheetId: string, range: Range) {
  if (!(wb.meta(sheetId).validations ?? []).some((v) => rangesIntersect(v.range, range))) return
  wb.run({
    label: 'Clear validation',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      m.validations = (m.validations ?? []).filter((v) => !rangesIntersect(v.range, range))
    },
  })
}

function evalFormulaValues(wb: Workbook, sheetId: string, formula: string): Scalar[] {
  try {
    const res = wb.hf.calculateFormula(formula, wb.hfId(sheetId))
    if (Array.isArray(res)) return (res as unknown[][]).flat().map((v) => toScalar(v as never))
    return [toScalar(res as never)]
  } catch {
    return []
  }
}

/** Items for a list validation (from literal text or a range/formula). */
export function listItems(wb: Workbook, sheetId: string, v: Validation): string[] {
  const src = (v.list ?? '').trim()
  if (!src) return []
  if (src.startsWith('=')) {
    return evalFormulaValues(wb, sheetId, src)
      .filter((x) => !isBlank(x) && !(typeof x === 'object' && x !== null))
      .map((x) => (typeof x === 'boolean' ? (x ? 'TRUE' : 'FALSE') : String(x)))
  }
  return src.split(/[,;]/).map((s) => s.trim()).filter(Boolean)
}

function bound(wb: Workbook, sheetId: string, s: string | undefined): number {
  if (s === undefined || s === '') return NaN
  if (s.startsWith('=')) {
    const v = evalFormulaValues(wb, sheetId, s)[0]
    return typeof v === 'number' ? v : NaN
  }
  const p = parseInput(s)
  return typeof p.raw === 'number' ? p.raw : NaN
}

export type ValidationResult = { ok: true } | { ok: false; style: Validation['errorStyle']; title: string; message: string }

/** Checks a typed value against the cell's validation rule (if any). */
export function checkInput(wb: Workbook, sheetId: string, r: number, c: number, input: string): ValidationResult {
  const v = validationAt(wb, sheetId, r, c)
  if (!v) return { ok: true }
  return checkValue(wb, sheetId, r, c, v, input)
}

function describe(v: Validation): string {
  const ops: Record<string, string> = { between: 'between', notBetween: 'not between', eq: 'equal to', ne: 'not equal to', gt: 'greater than', lt: 'less than', gte: 'at least', lte: 'at most' }
  const what = { whole: 'a whole number', decimal: 'a number', date: 'a date', time: 'a time', textLength: 'text whose length is', list: '', custom: '' }[v.type]
  if (v.type === 'list') return 'one of the values in the list'
  if (v.type === 'custom') return 'a value that satisfies the rule'
  const range = v.op === 'between' || v.op === 'notBetween' ? `${v.v1} and ${v.v2}` : `${v.v1}`
  return `${what} ${ops[v.op ?? 'between']} ${range}`
}

export function checkValue(wb: Workbook, sheetId: string, r: number, c: number, v: Validation, input: string): ValidationResult {
  const fail = (): ValidationResult => ({
    ok: false,
    style: v.errorStyle,
    title: v.errorTitle || 'Invalid value',
    message: v.errorMessage || `This cell only accepts ${describe(v)}.`,
  })
  if (input === '' || input === null) return v.allowBlank ? { ok: true } : fail()
  const parsed = parseInput(input)
  const raw = parsed.raw
  switch (v.type) {
    case 'list': {
      const items = listItems(wb, sheetId, v).map((s) => s.toLowerCase())
      return items.includes(input.trim().toLowerCase()) ? { ok: true } : fail()
    }
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time': {
      if (typeof raw !== 'number') return fail()
      if (v.type === 'whole' && !Number.isInteger(raw)) return fail()
      const a = bound(wb, sheetId, v.v1)
      const b = bound(wb, sheetId, v.v2)
      return compareOp(v.op ?? 'between', raw, a, b) ? { ok: true } : fail()
    }
    case 'textLength': {
      const len = input.length
      return compareOp(v.op ?? 'between', len, bound(wb, sheetId, v.v1), bound(wb, sheetId, v.v2)) ? { ok: true } : fail()
    }
    case 'custom': {
      const f = (v.list ?? v.v1 ?? '').trim()
      if (!f) return { ok: true }
      const shifted = shiftFormula(f.startsWith('=') ? f : '=' + f, r - v.range.r1, c - v.range.c1)
      // substitute the cell itself with the proposed value
      const self = `${colToLetters(c)}${r + 1}`
      const literal = typeof raw === 'number' || typeof raw === 'boolean' ? String(raw).toUpperCase() : `"${String(raw).replace(/"/g, '""')}"`
      const expr = shifted.replace(new RegExp(`(?<![A-Za-z0-9_$!:])\\$?${colToLetters(c)}\\$?${r + 1}(?![0-9:])`, 'g'), literal)
      void self
      const res = evalFormulaValues(wb, sheetId, expr)[0]
      return res === true || (typeof res === 'number' && res !== 0) ? { ok: true } : fail()
    }
  }
  return { ok: true }
}

/** Cells whose current content breaks their validation rule (Review > Find invalid data). */
export function invalidCells(wb: Workbook, sheetId: string, limit = 5000): { r: number; c: number }[] {
  const out: { r: number; c: number }[] = []
  for (const v of wb.meta(sheetId).validations ?? []) {
    const rg = clampToUsed(wb, sheetId, v.range)
    for (let r = rg.r1; r <= rg.r2; r++)
      for (let c = rg.c1; c <= rg.c2; c++) {
        if (out.length >= limit) return out
        if (wb.isCellEmpty(sheetId, r, c)) continue
        const text = wb.getEditText(sheetId, r, c).replace(/^'/, '')
        const isFormula = wb.isFormula(sheetId, r, c)
        const valText = isFormula ? wb.getDisplay(sheetId, r, c).text : text
        if (!checkValue(wb, sheetId, r, c, v, valText).ok) out.push({ r, c })
      }
  }
  return out
}

// ===================================================================== names, notes, links, freeze
export function removeNamedRange(wb: Workbook, name: string) {
  wb.run({ label: 'Delete name', dirty: 'all', hfOps: () => wb.hf.removeNamedExpression(name) })
}

export function changeNamedRange(wb: Workbook, name: string, expression: string): string | null {
  if (!wb.hf.isItPossibleToChangeNamedExpression(name, expression)) return 'That reference is not valid.'
  wb.run({ label: 'Edit name', dirty: 'all', hfOps: () => wb.hf.changeNamedExpression(name, expression) })
  return null
}

/** Range a defined name points to, when it is a plain reference. */
export function nameTarget(wb: Workbook, expression: string): { sheetId: string | null; range: Range } | null {
  const ex = expression.replace(/^=/, '')
  const bang = ex.lastIndexOf('!')
  const rg = parseRangeA1(bang > 0 ? ex.slice(bang + 1) : ex)
  if (!rg) return null
  const sheet = bang > 0 ? wb.sheetByName(ex.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'")) : undefined
  return { sheetId: sheet?.id ?? null, range: rg }
}

export function setNote(wb: Workbook, sheetId: string, r: number, c: number, text: string | null) {
  wb.run({
    label: text ? 'Edit note' : 'Delete note',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      const notes = { ...(m.notes ?? {}) }
      if (text) notes[cellKey(r, c)] = text
      else delete notes[cellKey(r, c)]
      m.notes = notes
    },
  })
}

export function setLink(wb: Workbook, sheetId: string, r: number, c: number, url: string | null, displayText?: string) {
  const safe = url && /^(https?:|mailto:|#)/i.test(url) ? url : url ? `https://${url}` : null
  const needsText = displayText !== undefined && displayText !== wb.getEditText(sheetId, r, c)
  wb.run({
    label: url ? 'Insert link' : 'Remove link',
    dirty: [sheetId],
    metaScope: [sheetId],
    hfOps: needsText || (safe && wb.isCellEmpty(sheetId, r, c)) ? () => wb.hf.setCellContents({ sheet: wb.hfId(sheetId), row: r, col: c }, (needsText ? displayText : safe) as string) : undefined,
    metaOps: () => {
      const m = wb.meta(sheetId)
      const links = { ...(m.links ?? {}) }
      if (safe) links[cellKey(r, c)] = safe
      else delete links[cellKey(r, c)]
      m.links = links
    },
  })
}

export function setFreeze(wb: Workbook, sheetId: string, rows: number, cols: number) {
  wb.run({
    label: rows || cols ? 'Freeze panes' : 'Unfreeze panes',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      wb.meta(sheetId).freeze = rows || cols ? { rows, cols } : undefined
    },
  })
}

export { isDateFormat }

// ===================================================================== page setup
import { DEFAULT_PAGE_SETUP, type PageSetup } from '../../model/types'
export function setPageSetup(wb: Workbook, sheetId: string, patch: Partial<PageSetup>, label = 'Page setup') {
  wb.run({
    label,
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      m.pageSetup = { ...DEFAULT_PAGE_SETUP, ...(m.pageSetup ?? {}), ...patch }
    },
  })
}
