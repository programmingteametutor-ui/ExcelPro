/* PivotTable: create dialog and the field-list side panel. */
import { useMemo, useState } from 'react'
import type { Workbook } from '../engine/workbook'
import type { AggFn, PivotDef, Range } from '../model/types'
import { addrToA1, parseRangeA1, quoteSheetName, rangeToA1 } from '../model/address'
import { AGG_LABEL, fieldValues, readPivotSource } from '../engine/features/pivot'
import { Button, Check, Dialog, Field, inputClass } from './components/ui'
import { Icon } from './icons'
import { cn } from './lib/utils'

export type PivotFields = Pick<PivotDef, 'rows' | 'cols' | 'values' | 'filters' | 'dateGroup' | 'sort' | 'grandTotals' | 'autoRefresh'>
type Area = 'none' | 'rows' | 'cols' | 'values' | 'filters'

const sel = cn(inputClass, 'h-7 text-[12px]')

function areaOf(f: string, d: PivotFields): Area {
  if (d.rows.includes(f)) return 'rows'
  if (d.cols.includes(f)) return 'cols'
  if (d.values.some((v) => v.field === f)) return 'values'
  if (d.filters.some((v) => v.field === f)) return 'filters'
  return 'none'
}

function moveField(d: PivotFields, f: string, to: Area, numeric: boolean): PivotFields {
  const n: PivotFields = {
    ...d,
    rows: d.rows.filter((x) => x !== f),
    cols: d.cols.filter((x) => x !== f),
    values: d.values.filter((x) => x.field !== f),
    filters: d.filters.filter((x) => x.field !== f),
  }
  if (to === 'rows') n.rows.push(f)
  if (to === 'cols') n.cols.push(f)
  if (to === 'values') n.values.push({ field: f, agg: numeric ? 'sum' : 'count' })
  if (to === 'filters') n.filters.push({ field: f, values: [] })
  return n
}

/** The field list shared by the create dialog and the side panel. */
export function PivotFieldList({ wb, source, value, onChange }: { wb: Workbook; source: PivotDef['source']; value: PivotFields; onChange: (v: PivotFields) => void }) {
  const src = useMemo(() => readPivotSource(wb, source.sheetId, source.range), [wb, source])
  const numeric = (f: string) => {
    const i = src.fields.indexOf(f)
    const sample = src.rows.slice(0, 50).map((r) => r[i]).filter((v) => v !== null && v !== '')
    return sample.length > 0 && sample.every((v) => typeof v === 'number')
  }
  const [openFilter, setOpenFilter] = useState<string | null>(null)
  if (!src.fields.length) return <p className="text-[var(--muted)]">The source range is empty.</p>
  return (
    <div className="space-y-1">
      {src.fields.map((f) => {
        const area = areaOf(f, value)
        const v = value.values.find((x) => x.field === f)
        const flt = value.filters.find((x) => x.field === f)
        return (
          <div key={f} className="rounded-md border border-[var(--border)] bg-[var(--bg)] px-2 py-1">
            <div className="flex items-center gap-1.5">
              <span className="min-w-0 flex-1 truncate font-medium" title={f}>{f}</span>
              <select className={cn(sel, 'w-[92px]')} value={area} aria-label={`Area for ${f}`} onChange={(e) => onChange(moveField(value, f, e.target.value as Area, numeric(f)))}>
                <option value="none">—</option>
                <option value="rows">Rows</option>
                <option value="cols">Columns</option>
                <option value="values">Values</option>
                <option value="filters">Filter</option>
              </select>
            </div>
            {v && (
              <div className="mt-1 flex items-center gap-1.5 text-[11.5px] text-[var(--muted)]">
                Summarize by
                <select className={cn(sel, 'flex-1')} value={v.agg} aria-label={`Summarize ${f} by`} onChange={(e) => onChange({ ...value, values: value.values.map((x) => (x.field === f ? { ...x, agg: e.target.value as AggFn } : x)) })}>
                  {(Object.keys(AGG_LABEL) as AggFn[]).map((a) => <option key={a} value={a}>{AGG_LABEL[a]}</option>)}
                </select>
              </div>
            )}
            {(area === 'rows' || area === 'cols') && src.dateFields.has(f) && (
              <div className="mt-1 flex items-center gap-1.5 text-[11.5px] text-[var(--muted)]">
                Group dates by
                <select className={cn(sel, 'flex-1')} value={value.dateGroup?.[f] ?? ''} onChange={(e) => {
                  const g = { ...value.dateGroup }
                  if (e.target.value) g[f] = e.target.value as 'year' | 'quarter' | 'month'
                  else delete g[f]
                  onChange({ ...value, dateGroup: g })
                }}>
                  <option value="">Day (no grouping)</option>
                  <option value="month">Month</option>
                  <option value="quarter">Quarter</option>
                  <option value="year">Year</option>
                </select>
              </div>
            )}
            {flt && (
              <div className="mt-1 text-[11.5px]">
                <button type="button" className="text-[var(--accent)] hover:underline" onClick={() => setOpenFilter(openFilter === f ? null : f)}>
                  {flt.values.length ? `${flt.values.length} selected` : 'All items'} — choose…
                </button>
                {openFilter === f && (
                  <FieldFilter wb={wb} source={source} dateGroup={value.dateGroup} field={f} selected={flt.values} onChange={(vals) => onChange({ ...value, filters: value.filters.map((x) => (x.field === f ? { ...x, values: vals } : x)) })} />
                )}
              </div>
            )}
          </div>
        )
      })}
      <div className="flex flex-wrap items-center gap-x-3 pt-1">
        <label className="flex items-center gap-1.5 text-[12px]">
          Sort
          <select className={sel} value={value.sort} onChange={(e) => onChange({ ...value, sort: e.target.value as 'asc' | 'desc' })}>
            <option value="asc">Ascending</option>
            <option value="desc">Descending</option>
          </select>
        </label>
        <Check label="Grand totals" checked={value.grandTotals} onChange={(g) => onChange({ ...value, grandTotals: g })} />
        <Check label="Refresh when data changes" checked={value.autoRefresh} onChange={(g) => onChange({ ...value, autoRefresh: g })} />
      </div>
    </div>
  )
}

function FieldFilter({ wb, source, dateGroup, field, selected, onChange }: { wb: Workbook; source: PivotDef['source']; dateGroup: PivotDef['dateGroup']; field: string; selected: string[]; onChange: (v: string[]) => void }) {
  const all = useMemo(() => fieldValues(wb, { source, dateGroup }, field), [wb, source, dateGroup, field])
  const cur = new Set(selected.length ? selected : all)
  return (
    <div className="mt-1 max-h-40 overflow-auto rounded border border-[var(--border)] px-1">
      <Check label="(All)" checked={cur.size === all.length} onChange={(v) => onChange(v ? [] : [all[0]].filter(Boolean))} />
      {all.slice(0, 500).map((v) => (
        <Check key={v} label={v || '(blank)'} checked={cur.has(v)} onChange={(on) => {
          const n = new Set(cur)
          on ? n.add(v) : n.delete(v)
          onChange(n.size === all.length ? [] : [...n])
        }} />
      ))}
    </div>
  )
}

export function CreatePivotDialog({ wb, sheetId, initialRange, anchorFree, onClose, onCreate }: {
  wb: Workbook
  sheetId: string
  initialRange: Range
  anchorFree: { r: number; c: number }
  onClose: () => void
  onCreate: (o: { source: PivotDef['source']; dest: 'new' | { sheetId: string; r: number; c: number }; fields: PivotFields }) => void
}) {
  const [srcText, setSrcText] = useState(rangeToA1(initialRange))
  const [destMode, setDestMode] = useState<'new' | 'existing'>('new')
  const [destText, setDestText] = useState(`${quoteSheetName(wb.meta(sheetId).name)}!${addrToA1(anchorFree)}`)
  const [fields, setFields] = useState<PivotFields>({ rows: [], cols: [], values: [], filters: [], sort: 'asc', grandTotals: true, autoRefresh: false })
  const rg = parseRangeA1(srcText.replace(/\$/g, '').replace(/^.*!/, ''))
  const source = useMemo(() => (rg ? { sheetId, range: rg } : null), [srcText]) // eslint-disable-line react-hooks/exhaustive-deps
  const parseDest = () => {
    const bang = destText.lastIndexOf('!')
    const s = bang > 0 ? wb.sheetByName(destText.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'")) : wb.meta(sheetId)
    const a = parseRangeA1(destText.slice(bang + 1).replace(/\$/g, ''))
    return s && a ? { sheetId: s.id, r: a.r1, c: a.c1 } : null
  }
  const ok = source && fields.values.length > 0 && (destMode === 'new' || parseDest())
  return (
    <Dialog
      title="Create PivotTable"
      onClose={onClose}
      width={560}
      footer={
        <>
          <span className="mr-auto text-[12px] text-[var(--muted)]">{fields.values.length ? '' : 'Put at least one field in Values.'}</span>
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" disabled={!ok} onClick={() => source && onCreate({ source, dest: destMode === 'new' ? 'new' : parseDest()!, fields })}>Create</Button>
        </>
      }
    >
      <Field label="Source data (with a header row)">
        <input className={cn(inputClass, 'font-mono')} value={srcText} onChange={(e) => setSrcText(e.target.value)} aria-label="Source range" />
      </Field>
      <div className="mb-3 flex flex-wrap items-center gap-3 text-[12.5px]">
        <label className="flex items-center gap-1.5"><input type="radio" className="accent-[var(--accent)]" checked={destMode === 'new'} onChange={() => setDestMode('new')} />New worksheet</label>
        <label className="flex items-center gap-1.5"><input type="radio" className="accent-[var(--accent)]" checked={destMode === 'existing'} onChange={() => setDestMode('existing')} />Existing worksheet</label>
        {destMode === 'existing' && <input className={cn(inputClass, 'h-7 w-40 font-mono')} value={destText} onChange={(e) => setDestText(e.target.value)} aria-label="Destination cell" />}
      </div>
      <div className="mb-1 text-[12px] font-medium text-[var(--muted)]">Fields</div>
      <div className="max-h-[42vh] overflow-auto rounded-md bg-[var(--panel)] p-2">
        {source ? <PivotFieldList wb={wb} source={source} value={fields} onChange={setFields} /> : <p className="text-[var(--danger)]">Enter a valid range.</p>}
      </div>
    </Dialog>
  )
}

export function PivotPanel({ wb, pivot, error, onChange, onRefresh, onDelete, onClose }: { wb: Workbook; pivot: PivotDef; error: string | null; onChange: (f: PivotFields) => void; onRefresh: () => void; onDelete: () => void; onClose: () => void }) {
  const srcSheet = wb.sheets.find((s) => s.id === pivot.source.sheetId)
  return (
    <aside className="flex h-full w-[300px] shrink-0 flex-col border-l border-[var(--border)] bg-[var(--panel)]" aria-label="PivotTable fields" data-testid="pivot-panel">
      <div className="flex h-10 shrink-0 items-center gap-2 border-b border-[var(--border)] px-3">
        <Icon name="pivot" size={16} />
        <span className="truncate font-semibold">{pivot.name}</span>
        <Button size="icon" className="ml-auto" aria-label="Close PivotTable panel" onClick={onClose}><Icon name="x" size={15} /></Button>
      </div>
      <div className="min-h-0 flex-1 overflow-auto p-3 text-[12.5px]">
        <p className="mb-2 text-[11.5px] text-[var(--muted)]">Source: <span className="font-mono">{srcSheet ? `${quoteSheetName(srcSheet.name)}!${rangeToA1(pivot.source.range)}` : 'missing sheet'}</span></p>
        {error && <p className="mb-2 rounded bg-[var(--warn-bg)] px-2 py-1 text-[var(--warn-text)]" role="alert">{error}</p>}
        <PivotFieldList wb={wb} source={pivot.source} value={pivot} onChange={onChange} />
      </div>
      <div className="flex shrink-0 gap-2 border-t border-[var(--border)] p-2">
        <Button size="sm" variant="outline" onClick={onRefresh}><Icon name="refresh" size={14} /> Refresh</Button>
        <Button size="sm" variant="ghost" className="ml-auto text-[var(--danger)]" onClick={onDelete}><Icon name="trash" size={14} /> Delete</Button>
      </div>
    </aside>
  )
}
