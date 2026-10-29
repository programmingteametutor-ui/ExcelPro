/* Charts drawn over the grid (move, resize, select) and the chart editing side panel. */
import { useMemo, useRef, useState } from 'react'
import type { Workbook } from '../engine/workbook'
import type { ChartDef, ChartType } from '../model/types'
import { parseRangeA1, quoteSheetName, rangeToA1 } from '../model/address'
import type { ObjectsApi } from '../grid/Grid'
import { ChartView, DARK_CHART, LIGHT_CHART } from '../charts/ChartView'
import { CHART_COLORS, extractChartData, updateChart } from '../charts/data'
import { Button, Check, Field, inputClass } from './components/ui'
import { Icon } from './icons'
import { cn } from './lib/utils'

export const CHART_TYPES: { type: ChartType; label: string; icon: string }[] = [
  { type: 'column', label: 'Column', icon: 'chart' },
  { type: 'bar', label: 'Bar', icon: 'chart' },
  { type: 'line', label: 'Line', icon: 'lineChart' },
  { type: 'area', label: 'Area', icon: 'lineChart' },
  { type: 'pie', label: 'Pie', icon: 'pie' },
  { type: 'doughnut', label: 'Doughnut', icon: 'pie' },
  { type: 'scatter', label: 'Scatter (XY)', icon: 'target' },
  { type: 'histogram', label: 'Histogram', icon: 'chart' },
]

interface Drag {
  id: string
  mode: 'move' | 'resize'
  sx: number
  sy: number
  dx: number
  dy: number
}

export function ChartsLayer({ wb, sheetId, version, api, selected, onSelect, onOpenEditor, dark }: { wb: Workbook; sheetId: string; version: number; api: ObjectsApi; selected: string | null; onSelect: (id: string | null) => void; onOpenEditor: (id: string) => void; dark: boolean }) {
  const charts = wb.meta(sheetId).charts ?? []
  const [drag, setDrag] = useState<Drag | null>(null)
  const dragRef = useRef<Drag | null>(null)
  const data = useMemo(() => new Map(charts.map((c) => [c.id, extractChartData(wb, c)])), [charts, version, wb]) // eslint-disable-line react-hooks/exhaustive-deps
  if (!charts.length) return null
  const z = api.zoom
  const start = (e: React.PointerEvent, c: ChartDef, mode: Drag['mode']) => {
    if (e.button !== 0) return
    e.stopPropagation()
    e.preventDefault()
    onSelect(c.id)
    ;(e.currentTarget as HTMLElement).setPointerCapture(e.pointerId)
    const d = { id: c.id, mode, sx: e.clientX, sy: e.clientY, dx: 0, dy: 0 }
    dragRef.current = d
    setDrag(d)
  }
  const moveDrag = (e: React.PointerEvent) => {
    const d = dragRef.current
    if (!d) return
    const n = { ...d, dx: e.clientX - d.sx, dy: e.clientY - d.sy }
    dragRef.current = n
    setDrag(n)
  }
  const end = (c: ChartDef) => {
    const d = dragRef.current
    dragRef.current = null
    setDrag(null)
    if (!d || (Math.abs(d.dx) < 2 && Math.abs(d.dy) < 2)) return
    if (d.mode === 'move') updateChart(wb, sheetId, c.id, { x: Math.max(0, Math.round(c.x + d.dx / z)), y: Math.max(0, Math.round(c.y + d.dy / z)) }, 'Move chart')
    else updateChart(wb, sheetId, c.id, { w: Math.max(160, Math.round(c.w + d.dx / z)), h: Math.max(110, Math.round(c.h + d.dy / z)) }, 'Resize chart')
  }
  return (
    <>
      {charts.map((c) => {
        const d = drag?.id === c.id ? drag : null
        const p = api.toScreen((c.x + (d?.mode === 'move' ? d.dx / z : 0)) * z, (c.y + (d?.mode === 'move' ? d.dy / z : 0)) * z)
        const w = Math.max(160, c.w + (d?.mode === 'resize' ? d.dx / z : 0)) * z
        const h = Math.max(110, c.h + (d?.mode === 'resize' ? d.dy / z : 0)) * z
        if (p.x > api.viewport.w || p.y > api.viewport.h || p.x + w < 0 || p.y + h < 0) return null
        const sel = selected === c.id
        return (
          <div
            key={c.id}
            data-testid="chart"
            role="img"
            aria-label={`${c.title || c.type} chart`}
            className={cn('pointer-events-auto absolute overflow-hidden rounded-[3px] bg-[var(--bg)] shadow-sm', sel ? 'outline outline-2 outline-[var(--accent)]' : 'outline outline-1 outline-[var(--border)]')}
            style={{ left: p.x, top: p.y, width: w, height: h, zIndex: sel ? 12 : 11, cursor: d?.mode === 'move' ? 'grabbing' : 'move' }}
            onPointerDown={(e) => start(e, c, 'move')}
            onPointerMove={moveDrag}
            onPointerUp={() => end(c)}
            onDoubleClick={(e) => (e.stopPropagation(), onOpenEditor(c.id))}
            onContextMenu={(e) => (e.preventDefault(), e.stopPropagation(), onSelect(c.id), onOpenEditor(c.id))}
          >
            <div style={{ width: w / z, height: h / z, transform: `scale(${z})`, transformOrigin: '0 0' }}>
              <ChartView def={c} data={data.get(c.id) ?? { categories: [], series: [] }} width={w / z} height={h / z} theme={dark ? DARK_CHART : LIGHT_CHART} />
            </div>
            {sel && (
              <div
                aria-label="Resize chart"
                className="absolute bottom-0 right-0 h-3.5 w-3.5 cursor-nwse-resize rounded-tl bg-[var(--accent)]"
                onPointerDown={(e) => start(e, c, 'resize')}
                onPointerMove={moveDrag}
                onPointerUp={() => end(c)}
              />
            )}
          </div>
        )
      })}
    </>
  )
}

export function ChartPanel({ wb, sheetId, chart, onClose, onDelete }: { wb: Workbook; sheetId: string; chart: ChartDef; onClose: () => void; onDelete: () => void }) {
  const srcSheet = wb.sheets.find((s) => s.id === chart.source.sheetId)
  const [srcText, setSrcText] = useState(`${srcSheet && srcSheet.id !== sheetId ? quoteSheetName(srcSheet.name) + '!' : ''}${rangeToA1(chart.source.range)}`)
  const [srcErr, setSrcErr] = useState<string | null>(null)
  const up = (patch: Partial<ChartDef>, label?: string) => updateChart(wb, sheetId, chart.id, patch, label)
  const data = extractChartData(wb, chart)
  const pieLike = chart.type === 'pie' || chart.type === 'doughnut'
  const applySource = () => {
    const bang = srcText.lastIndexOf('!')
    let sid = chart.source.sheetId
    if (bang > 0) {
      const s = wb.sheetByName(srcText.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'"))
      if (!s) return setSrcErr('There is no sheet with that name.')
      sid = s.id
    } else sid = sheetId
    const rg = parseRangeA1(srcText.slice(bang + 1).replace(/\$/g, ''))
    if (!rg) return setSrcErr('Enter a range such as A1:C10.')
    setSrcErr(null)
    up({ source: { sheetId: sid, range: rg } }, 'Change chart data')
  }
  const sel = cn(inputClass, 'h-8')
  return (
    <aside className="flex h-full w-[290px] shrink-0 flex-col border-l border-[var(--border)] bg-[var(--panel)]" aria-label="Chart settings" data-testid="chart-panel">
      <div className="flex h-10 shrink-0 items-center gap-2 border-b border-[var(--border)] px-3">
        <Icon name="chart" size={16} />
        <span className="font-semibold">Chart</span>
        <Button size="icon" className="ml-auto" aria-label="Close chart panel" onClick={onClose}><Icon name="x" size={15} /></Button>
      </div>
      <div className="min-h-0 flex-1 overflow-auto p-3 text-[12.5px]">
        <Field label="Chart type">
          <select className={sel} value={chart.type} onChange={(e) => up({ type: e.target.value as ChartType }, 'Change chart type')} aria-label="Chart type">
            {CHART_TYPES.map((t) => <option key={t.type} value={t.type}>{t.label}</option>)}
          </select>
        </Field>
        <Field label="Data range" hint={srcErr ?? 'Press Enter to apply.'}>
          <input className={cn(inputClass, 'font-mono', srcErr && 'border-[var(--danger)]')} value={srcText} onChange={(e) => setSrcText(e.target.value)} onBlur={applySource} onKeyDown={(e) => e.key === 'Enter' && applySource()} aria-label="Data range" />
        </Field>
        <div className="mb-3 grid grid-cols-2 gap-x-2">
          <Field label="Series in">
            <select className={sel} value={chart.seriesIn} onChange={(e) => up({ seriesIn: e.target.value as ChartDef['seriesIn'] })}>
              <option value="cols">Columns</option>
              <option value="rows">Rows</option>
            </select>
          </Field>
          <Field label="Legend">
            <select className={sel} value={chart.legend} onChange={(e) => up({ legend: e.target.value as ChartDef['legend'] })} aria-label="Legend position">
              <option value="bottom">Bottom</option>
              <option value="top">Top</option>
              <option value="right">Right</option>
              <option value="none">None</option>
            </select>
          </Field>
        </div>
        <Check label={chart.seriesIn === 'cols' ? 'First row is series names' : 'First column is series names'} checked={chart.firstRowHeader} onChange={(v) => up({ firstRowHeader: v })} />
        <Check label={chart.seriesIn === 'cols' ? 'First column is categories' : 'First row is categories'} checked={chart.firstColLabels} onChange={(v) => up({ firstColLabels: v })} />
        <div className="my-3 h-px bg-[var(--border)]" />
        <Field label="Title">
          <input className={inputClass} defaultValue={chart.title} key={chart.id + 't'} onBlur={(e) => e.target.value !== chart.title && up({ title: e.target.value }, 'Chart title')} onKeyDown={(e) => e.key === 'Enter' && (e.target as HTMLInputElement).blur()} aria-label="Chart title" />
        </Field>
        {!pieLike && (
          <div className="grid grid-cols-2 gap-x-2">
            <Field label="Horizontal axis title">
              <input className={inputClass} defaultValue={chart.xTitle} key={chart.id + 'x'} onBlur={(e) => e.target.value !== chart.xTitle && up({ xTitle: e.target.value }, 'Axis title')} />
            </Field>
            <Field label="Vertical axis title">
              <input className={inputClass} defaultValue={chart.yTitle} key={chart.id + 'y'} onBlur={(e) => e.target.value !== chart.yTitle && up({ yTitle: e.target.value }, 'Axis title')} />
            </Field>
          </div>
        )}
        <Check label="Data labels" checked={chart.dataLabels} onChange={(v) => up({ dataLabels: v })} />
        {!pieLike && <Check label="Gridlines" checked={chart.gridlines} onChange={(v) => up({ gridlines: v })} />}
        {['column', 'bar', 'area'].includes(chart.type) && <Check label="Stacked" checked={chart.stacked} onChange={(v) => up({ stacked: v })} />}
        {chart.type === 'histogram' && (
          <Field label="Number of bins">
            <input className={inputClass} type="number" min={1} max={100} value={chart.bins ?? 10} onChange={(e) => up({ bins: Math.max(1, Math.min(100, +e.target.value || 10)) }, 'Histogram bins')} />
          </Field>
        )}
        {!pieLike && chart.type !== 'histogram' && data.series.length > 0 && (
          <>
            <div className="my-3 h-px bg-[var(--border)]" />
            <div className="mb-1 font-medium">Series</div>
            {data.series.map((s, i) => (
              <div key={i} className="flex items-center gap-2 py-0.5">
                <input type="color" value={chart.colors[i] ?? CHART_COLORS[i % CHART_COLORS.length]} onChange={(e) => {
                  const colors = [...chart.colors]
                  while (colors.length <= i) colors.push(CHART_COLORS[colors.length % CHART_COLORS.length])
                  colors[i] = e.target.value
                  up({ colors }, 'Series color')
                }} className="h-6 w-8 rounded border border-[var(--border)]" aria-label={`Color of ${s.name}`} />
                <span className="min-w-0 flex-1 truncate">{s.name}</span>
                <Check label="Show" checked={!(chart.hiddenSeries ?? []).includes(i)} onChange={(v) => up({ hiddenSeries: v ? (chart.hiddenSeries ?? []).filter((x) => x !== i) : [...(chart.hiddenSeries ?? []), i] })} />
              </div>
            ))}
          </>
        )}
      </div>
      <div className="flex shrink-0 gap-2 border-t border-[var(--border)] p-2">
        <Button size="sm" variant="ghost" className="text-[var(--danger)]" onClick={onDelete}><Icon name="trash" size={14} /> Delete chart</Button>
      </div>
    </aside>
  )
}
