/* Floating panels anchored to a screen rectangle: AutoFilter menu and validation list. */
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import type { Workbook } from '../engine/workbook'
import type { FilterCriterion, FilterOp } from '../model/types'
import { filterValues } from '../engine/features/data'
import { Button, Check, inputClass } from './components/ui'
import { Icon } from './icons'
import { cn } from './lib/utils'

export function FloatPanel({ rect, onClose, children, width = 260, testId }: { rect: DOMRect; onClose: () => void; children: ReactNode; width?: number; testId?: string }) {
  const ref = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState({ left: rect.left, top: rect.bottom + 2 })
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    let top = rect.bottom + 2
    if (top + el.offsetHeight > window.innerHeight - 4) top = Math.max(4, rect.top - el.offsetHeight - 2)
    setPos({ left: Math.max(4, Math.min(rect.left, window.innerWidth - el.offsetWidth - 4)), top })
  }, [rect])
  useEffect(() => {
    const h = (e: PointerEvent) => !ref.current?.contains(e.target as Node) && onClose()
    const k = (e: KeyboardEvent) => e.key === 'Escape' && (e.stopPropagation(), onClose())
    const t = setTimeout(() => document.addEventListener('pointerdown', h, true), 0)
    document.addEventListener('keydown', k, true)
    return () => {
      clearTimeout(t)
      document.removeEventListener('pointerdown', h, true)
      document.removeEventListener('keydown', k, true)
    }
  }, [onClose])
  return createPortal(
    <div ref={ref} data-testid={testId} className="fixed z-[96] rounded-lg border border-[var(--border)] bg-[var(--bg)] p-1.5 text-[12.5px] text-[var(--text)] shadow-xl" style={{ ...pos, width }}>
      {children}
    </div>,
    document.body,
  )
}

const CONDITION_OPS: { op: FilterOp; label: string; two?: boolean; none?: boolean }[] = [
  { op: 'equals', label: 'Equals' },
  { op: 'notEquals', label: 'Does not equal' },
  { op: 'contains', label: 'Contains' },
  { op: 'notContains', label: 'Does not contain' },
  { op: 'begins', label: 'Begins with' },
  { op: 'ends', label: 'Ends with' },
  { op: 'gt', label: 'Greater than' },
  { op: 'gte', label: 'Greater than or equal to' },
  { op: 'lt', label: 'Less than' },
  { op: 'lte', label: 'Less than or equal to' },
  { op: 'between', label: 'Between', two: true },
  { op: 'top', label: 'Top N items' },
  { op: 'bottom', label: 'Bottom N items' },
  { op: 'aboveAverage', label: 'Above average', none: true },
  { op: 'belowAverage', label: 'Below average', none: true },
]

export function FilterMenu({
  wb,
  sheetId,
  col,
  rect,
  colName,
  onClose,
  onSort,
  onApply,
}: {
  wb: Workbook
  sheetId: string
  col: number
  rect: DOMRect
  colName: string
  onClose: () => void
  onSort: (desc: boolean) => void
  onApply: (c: FilterCriterion | null) => void
}) {
  const current = wb.meta(sheetId).autoFilter?.criteria[col]
  const { values, hasBlanks } = useMemo(() => filterValues(wb, sheetId, col), [wb, sheetId, col])
  const [mode, setMode] = useState<'values' | 'condition'>(current?.kind === 'condition' ? 'condition' : 'values')
  const [q, setQ] = useState('')
  const [checked, setChecked] = useState<Set<string>>(() => new Set(current?.kind === 'values' ? current.values.filter((v) => values.includes(v)) : values))
  const [blanks, setBlanks] = useState(current?.kind === 'values' ? current.blanks : true)
  const [op, setOp] = useState<FilterOp>(current?.kind === 'condition' ? current.op : 'contains')
  const [a, setA] = useState(current?.kind === 'condition' ? current.a : '')
  const [b, setB] = useState(current?.kind === 'condition' ? current.b ?? '' : '')
  const shown = values.filter((v) => v.toLowerCase().includes(q.toLowerCase())).slice(0, 1000)
  const allOn = shown.every((v) => checked.has(v)) && (!hasBlanks || blanks)
  const opDef = CONDITION_OPS.find((o) => o.op === op)!
  const apply = () => {
    if (mode === 'values') {
      // a search narrows the selection to the matching items, like Excel
      const keep = q ? shown.filter((v) => checked.has(v)) : [...checked]
      const everything = !q && keep.length === values.length && (!hasBlanks || blanks)
      onApply(everything ? null : { kind: 'values', values: keep, blanks: q ? false : blanks })
    } else onApply({ kind: 'condition', op, a, b: opDef.two ? b : undefined })
  }
  return (
    <FloatPanel rect={rect} onClose={onClose} width={270} testId="filter-menu">
      <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1.5 text-left hover:bg-[var(--hover)]" onClick={() => onSort(false)}>
        <Icon name="sortAsc" size={15} /> Sort A to Z
      </button>
      <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1.5 text-left hover:bg-[var(--hover)]" onClick={() => onSort(true)}>
        <Icon name="sortDesc" size={15} /> Sort Z to A
      </button>
      <button type="button" disabled={!current} className="flex w-full items-center gap-2 rounded px-2 py-1.5 text-left hover:bg-[var(--hover)] disabled:opacity-40" onClick={() => onApply(null)}>
        <Icon name="eraser" size={15} /> Clear filter from “{colName}”
      </button>
      <div className="my-1 h-px bg-[var(--border)]" />
      <div className="mb-1.5 flex gap-1">
        {(['values', 'condition'] as const).map((m) => (
          <button key={m} type="button" onClick={() => setMode(m)} className={cn('flex-1 rounded px-2 py-1 text-[12px]', mode === m ? 'bg-[var(--active)] font-semibold text-[var(--accent)]' : 'hover:bg-[var(--hover)]')}>
            {m === 'values' ? 'By values' : 'By condition'}
          </button>
        ))}
      </div>
      {mode === 'values' ? (
        <>
          <input className={cn(inputClass, 'mb-1 h-7')} placeholder="Search" value={q} onChange={(e) => setQ(e.target.value)} aria-label="Search values" autoFocus />
          <div className="max-h-52 overflow-auto rounded border border-[var(--border)] px-1 py-0.5">
            <Check
              label="(Select all)"
              checked={allOn}
              onChange={(v) => {
                const next = new Set(checked)
                for (const x of shown) v ? next.add(x) : next.delete(x)
                setChecked(next)
                if (hasBlanks && !q) setBlanks(v)
              }}
            />
            {shown.map((v) => (
              <Check key={v} label={v} checked={checked.has(v)} onChange={(on) => setChecked((s) => {
                const n = new Set(s)
                on ? n.add(v) : n.delete(v)
                return n
              })} />
            ))}
            {hasBlanks && !q && <Check label="(Blanks)" checked={blanks} onChange={setBlanks} />}
            {values.length > 1000 && <p className="px-1 text-[11px] text-[var(--muted)]">Showing the first 1,000 items. Search to narrow the list.</p>}
          </div>
        </>
      ) : (
        <div className="space-y-1.5">
          <select className={cn(inputClass, 'h-7')} value={op} onChange={(e) => setOp(e.target.value as FilterOp)} aria-label="Condition">
            {CONDITION_OPS.map((o) => <option key={o.op} value={o.op}>{o.label}</option>)}
          </select>
          {!opDef.none && <input className={cn(inputClass, 'h-7')} value={a} onChange={(e) => setA(e.target.value)} placeholder={op === 'top' || op === 'bottom' ? 'N, e.g. 10' : 'Value (* and ? wildcards work)'} aria-label="Condition value" />}
          {opDef.two && <input className={cn(inputClass, 'h-7')} value={b} onChange={(e) => setB(e.target.value)} placeholder="and" aria-label="Second value" />}
        </div>
      )}
      <div className="mt-2 flex justify-end gap-1.5">
        <Button size="sm" variant="outline" onClick={onClose}>Cancel</Button>
        <Button size="sm" variant="default" onClick={apply} disabled={mode === 'values' && checked.size === 0 && !blanks}>OK</Button>
      </div>
    </FloatPanel>
  )
}

export function ListPicker({ items, rect, current, onPick, onClose }: { items: string[]; rect: DOMRect; current: string; onPick: (v: string) => void; onClose: () => void }) {
  return (
    <FloatPanel rect={rect} onClose={onClose} width={Math.max(160, rect.width + 120)} testId="validation-list">
      {items.length === 0 ? (
        <p className="px-2 py-1 text-[var(--muted)]">The list is empty.</p>
      ) : (
        <div className="max-h-60 overflow-auto" role="listbox">
          {items.map((it, i) => (
            <button key={`${it}${i}`} role="option" aria-selected={it === current} type="button" className={cn('block w-full truncate rounded px-2 py-1 text-left hover:bg-[var(--hover)]', it === current && 'bg-[var(--active)]')} onClick={() => onPick(it)}>
              {it || ' '}
            </button>
          ))}
        </div>
      )}
    </FloatPanel>
  )
}
