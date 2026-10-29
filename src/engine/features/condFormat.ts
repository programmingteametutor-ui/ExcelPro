/**
 * Conditional formatting: rule storage and evaluation. Evaluation is cached per workbook version,
 * so rendering a screen of cells costs one pass of range statistics plus per-cell checks.
 */
import { inRange } from '../../model/address'
import type { CFRule, CFStyle, CondFormat, Range } from '../../model/types'
import { shiftFormula } from '../formulaText'
import { parseInput } from '../inputParser'
import type { Workbook } from '../workbook'
import { clampToUsed, compareOp, toScalar, uid, type Scalar } from './common'

export interface CFResult {
  style?: CFStyle
  scaleFill?: string
  bar?: { pct: number; color: string; negative: boolean }
  icon?: { set: 'arrows' | 'traffic' | 'flags'; level: 0 | 1 | 2 } // 0 = low, 2 = high
}

export function addCondFormat(wb: Workbook, sheetId: string, cf: Omit<CondFormat, 'id'>): string {
  const id = uid('cf')
  wb.run({
    label: 'Conditional formatting',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      m.condFormats = [{ ...cf, id }, ...(m.condFormats ?? [])] // newest rule has highest priority
    },
  })
  return id
}

export function updateCondFormat(wb: Workbook, sheetId: string, id: string, patch: Partial<CondFormat>) {
  wb.run({
    label: 'Edit rule',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      m.condFormats = (m.condFormats ?? []).map((c) => (c.id === id ? { ...c, ...patch } : c))
    },
  })
}

export function removeCondFormats(wb: Workbook, sheetId: string, ids: string[] | 'all', range?: Range) {
  wb.run({
    label: 'Clear rules',
    dirty: [sheetId],
    metaScope: [sheetId],
    metaOps: () => {
      const m = wb.meta(sheetId)
      m.condFormats = (m.condFormats ?? []).filter((c) => {
        if (ids !== 'all') return !ids.includes(c.id)
        if (!range) return false
        return !(c.range.r1 <= range.r2 && range.r1 <= c.range.r2 && c.range.c1 <= range.c2 && range.c1 <= c.range.c2)
      })
    },
  })
}

export function moveCondFormat(wb: Workbook, sheetId: string, id: string, dir: -1 | 1) {
  const list = [...(wb.meta(sheetId).condFormats ?? [])]
  const i = list.findIndex((c) => c.id === id)
  const j = i + dir
  if (i < 0 || j < 0 || j >= list.length) return
  ;[list[i], list[j]] = [list[j], list[i]]
  wb.run({ label: 'Reorder rules', dirty: [sheetId], metaScope: [sheetId], metaOps: () => (wb.meta(sheetId).condFormats = list) })
}

// ----------------------------------------------------------------- evaluation
interface RuleStats {
  nums: number[]
  min: number
  max: number
  mid: number
  avg: number
  counts: Map<string, number>
  threshold: number
}

function key(v: Scalar): string {
  if (v === null) return ''
  if (typeof v === 'object') return 'E' + v.error
  return typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + String(v)
}

function stats(wb: Workbook, sheetId: string, cf: CondFormat): RuleStats {
  const rg = clampToUsed(wb, sheetId, cf.range)
  const nums: number[] = []
  const counts = new Map<string, number>()
  let n = 0
  for (let r = rg.r1; r <= rg.r2 && n < 200_000; r++)
    for (let c = rg.c1; c <= rg.c2 && n < 200_000; c++, n++) {
      const v = toScalar(wb.getValue(sheetId, r, c))
      if (typeof v === 'number') nums.push(v)
      if (v !== null && v !== '') counts.set(key(v), (counts.get(key(v)) ?? 0) + 1)
    }
  const sorted = [...nums].sort((a, b) => a - b)
  const min = sorted[0] ?? 0
  const max = sorted[sorted.length - 1] ?? 0
  const pct = (p: number) => {
    if (!sorted.length) return 0
    const idx = (sorted.length - 1) * p
    const lo = Math.floor(idx)
    const hi = Math.ceil(idx)
    return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo)
  }
  let threshold = 0
  const rule = cf.rule
  if (rule.type === 'top') {
    const k = rule.percent ? Math.max(1, Math.floor((sorted.length * rule.n) / 100)) : Math.max(1, rule.n)
    threshold = rule.bottom ? sorted[Math.min(sorted.length - 1, k - 1)] ?? -Infinity : sorted[Math.max(0, sorted.length - k)] ?? Infinity
  }
  return { nums, min, max, mid: pct(0.5), avg: nums.reduce((a, b) => a + b, 0) / (nums.length || 1), counts, threshold }
}

function hexToRgb(h: string): [number, number, number] {
  const x = h.replace('#', '')
  return [parseInt(x.slice(0, 2), 16), parseInt(x.slice(2, 4), 16), parseInt(x.slice(4, 6), 16)]
}
function mix(a: string, b: string, t: number): string {
  const [r1, g1, b1] = hexToRgb(a)
  const [r2, g2, b2] = hexToRgb(b)
  const f = (x: number, y: number) => Math.round(x + (y - x) * t).toString(16).padStart(2, '0')
  return `#${f(r1, r2)}${f(g1, g2)}${f(b1, b2)}`
}

function num(s: string, wb: Workbook, sheetId: string): number {
  if (s.startsWith('=')) {
    try {
      const v = wb.hf.calculateFormula(s, wb.hfId(sheetId))
      return typeof v === 'number' ? v : NaN
    } catch {
      return NaN
    }
  }
  const p = parseInput(s)
  return typeof p.raw === 'number' ? p.raw : NaN
}

function ruleMatches(wb: Workbook, sheetId: string, cf: CondFormat, st: RuleStats, r: number, c: number): boolean {
  const rule: CFRule = cf.rule
  const v = toScalar(wb.getValue(sheetId, r, c))
  switch (rule.type) {
    case 'cellIs': {
      if (typeof v === 'number') return compareOp(rule.op, v, num(rule.v1, wb, sheetId), rule.v2 !== undefined ? num(rule.v2, wb, sheetId) : undefined)
      if (typeof v === 'string' && (rule.op === 'eq' || rule.op === 'ne')) {
        const same = v.toLowerCase() === rule.v1.replace(/^"|"$/g, '').toLowerCase()
        return rule.op === 'eq' ? same : !same
      }
      return false
    }
    case 'text': {
      const t = wb.getDisplay(sheetId, r, c).text.toLowerCase()
      const q = rule.text.toLowerCase()
      return rule.op === 'contains' ? t.includes(q) : rule.op === 'notContains' ? !t.includes(q) : rule.op === 'begins' ? t.startsWith(q) : t.endsWith(q)
    }
    case 'duplicate':
      return v !== null && v !== '' && (st.counts.get(key(v)) ?? 0) > 1
    case 'unique':
      return v !== null && v !== '' && (st.counts.get(key(v)) ?? 0) === 1
    case 'blank':
      return v === null || v === ''
    case 'notBlank':
      return !(v === null || v === '')
    case 'errors':
      return typeof v === 'object' && v !== null
    case 'top':
      return typeof v === 'number' && (rule.bottom ? v <= st.threshold : v >= st.threshold)
    case 'average':
      return typeof v === 'number' && (rule.above ? v > st.avg : v < st.avg)
    case 'formula': {
      const f = rule.formula.startsWith('=') ? rule.formula : '=' + rule.formula
      try {
        const res = wb.hf.calculateFormula(shiftFormula(f, r - cf.range.r1, c - cf.range.c1), wb.hfId(sheetId))
        return res === true || (typeof res === 'number' && res !== 0)
      } catch {
        return false
      }
    }
    default:
      return false
  }
}

const cache = new WeakMap<Workbook, { version: number; sheetId: string; stats: Map<string, RuleStats>; cells: Map<string, CFResult | null> }>()

/** Returns an evaluator for one sheet, valid for the current workbook version. */
export function cfEvaluator(wb: Workbook, sheetId: string): ((r: number, c: number) => CFResult | null) | null {
  const rules = wb.meta(sheetId).condFormats
  if (!rules?.length) return null
  let entry = cache.get(wb)
  if (!entry || entry.version !== wb.version || entry.sheetId !== sheetId) {
    entry = { version: wb.version, sheetId, stats: new Map(), cells: new Map() }
    cache.set(wb, entry)
  }
  const e = entry
  const getStats = (cf: CondFormat) => {
    let s = e.stats.get(cf.id)
    if (!s) e.stats.set(cf.id, (s = stats(wb, sheetId, cf)))
    return s
  }
  return (r: number, c: number) => {
    const k = `${r},${c}`
    const hit = e.cells.get(k)
    if (hit !== undefined) return hit
    let out: CFResult = {}
    // rules[0] is highest priority: apply from lowest to highest so higher ones win
    const applicable = rules.filter((cf) => inRange(cf.range, r, c))
    let stopAt = applicable.length
    for (let i = 0; i < applicable.length; i++) {
      const cf = applicable[i]
      if (cf.stopIfTrue && !['colorScale', 'dataBar', 'iconSet'].includes(cf.rule.type) && ruleMatches(wb, sheetId, cf, getStats(cf), r, c)) {
        stopAt = i + 1
        break
      }
    }
    for (let i = stopAt - 1; i >= 0; i--) {
      const cf = applicable[i]
      const rule = cf.rule
      const st = getStats(cf)
      const v = toScalar(wb.getValue(sheetId, r, c))
      if (rule.type === 'colorScale') {
        if (typeof v !== 'number' || st.max === st.min) {
          if (typeof v === 'number') out = { ...out, scaleFill: rule.colors[0] }
          continue
        }
        const t = (v - st.min) / (st.max - st.min)
        let fill: string
        if (rule.colors.length >= 3) {
          const midT = (st.mid - st.min) / (st.max - st.min) || 0.5
          fill = t <= midT ? mix(rule.colors[0], rule.colors[1], midT ? t / midT : 0) : mix(rule.colors[1], rule.colors[2], (t - midT) / (1 - midT || 1))
        } else fill = mix(rule.colors[0], rule.colors[1], t)
        out = { ...out, scaleFill: fill }
      } else if (rule.type === 'dataBar') {
        if (typeof v !== 'number') continue
        const lo = Math.min(0, st.min)
        const hi = Math.max(0, st.max)
        const pct = hi === lo ? 0 : Math.abs(v) / (hi - lo)
        out = { ...out, bar: { pct: Math.min(1, pct), color: rule.color, negative: v < 0 } }
      } else if (rule.type === 'iconSet') {
        if (typeof v !== 'number') continue
        const span = st.max - st.min
        const t = span ? (v - st.min) / span : 0.5
        out = { ...out, icon: { set: rule.set, level: t >= 2 / 3 ? 2 : t >= 1 / 3 ? 1 : 0 } }
      } else if (ruleMatches(wb, sheetId, cf, st, r, c)) {
        out = { ...out, style: { ...out?.style, ...cf.style } }
      }
    }
    const res = Object.keys(out).length ? out : null
    e.cells.set(k, res)
    return res
  }
}

export const CF_PRESET_STYLES: { label: string; style: CFStyle }[] = [
  { label: 'Light red fill with dark red text', style: { fill: '#ffc7ce', color: '#9c0006' } },
  { label: 'Yellow fill with dark yellow text', style: { fill: '#ffeb9c', color: '#9c5700' } },
  { label: 'Green fill with dark green text', style: { fill: '#c6efce', color: '#006100' } },
  { label: 'Light red fill', style: { fill: '#ffc7ce' } },
  { label: 'Red text', style: { color: '#c00000' } },
  { label: 'Bold', style: { bold: true } },
]
