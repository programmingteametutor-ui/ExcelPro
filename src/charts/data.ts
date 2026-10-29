/** Chart model helpers: series extraction from a sheet range and chart CRUD (pure data). */
import type { ChartDef, ChartType, Range } from '../model/types'
import type { Workbook } from '../engine/workbook'
import { uid } from '../engine/features/common'
import { isDateFormat } from '../engine/numberFormat'

export const CHART_COLORS = ['#0f766e', '#e07a1f', '#3b6fd8', '#c2410c', '#7c3aed', '#16a34a', '#db2777', '#64748b']

export interface Series {
  name: string
  values: (number | null)[]
  xs?: (number | null)[] // scatter
}

export interface ChartData {
  categories: string[]
  series: Series[]
  numberFormat?: string
}

export function extractChartData(wb: Workbook, def: Pick<ChartDef, 'source' | 'seriesIn' | 'firstRowHeader' | 'firstColLabels' | 'type'>): ChartData {
  const { sheetId, range } = def.source
  if (!wb.sheets.some((s) => s.id === sheetId)) return { categories: [], series: [] }
  const d = wb.hf.getSheetDimensions(wb.hfId(sheetId))
  const rg: Range = { ...range, r2: Math.min(range.r2, Math.max(range.r1, d.height - 1)), c2: Math.min(range.c2, Math.max(range.c1, d.width - 1)) }
  const val = (r: number, c: number) => {
    const v = wb.getValue(sheetId, r, c)
    return typeof v === 'number' ? v : null
  }
  const txt = (r: number, c: number) => wb.getDisplay(sheetId, r, c).text
  const byCols = def.seriesIn === 'cols'
  const outer = byCols ? [rg.c1, rg.c2] : [rg.r1, rg.r2]
  const inner = byCols ? [rg.r1, rg.r2] : [rg.c1, rg.c2]
  const hasHead = byCols ? def.firstRowHeader : def.firstColLabels
  const hasCat = byCols ? def.firstColLabels : def.firstRowHeader
  const catIdx = outer[0]
  const firstSeries = hasCat ? outer[0] + 1 : outer[0]
  const firstPoint = hasHead ? inner[0] + 1 : inner[0]
  const at = (o: number, i: number) => (byCols ? [i, o] : [o, i]) as [number, number]
  const categories: string[] = []
  for (let i = firstPoint; i <= inner[1]; i++) categories.push(hasCat ? txt(...at(catIdx, i)) : String(i - firstPoint + 1))
  const series: Series[] = []
  if (def.type === 'scatter') {
    // first data column = X values, others = Y series
    const xo = hasCat ? catIdx : firstSeries
    const xs: (number | null)[] = []
    for (let i = firstPoint; i <= inner[1]; i++) xs.push(val(...at(xo, i)))
    for (let o = xo + 1; o <= outer[1]; o++) {
      const values: (number | null)[] = []
      for (let i = firstPoint; i <= inner[1]; i++) values.push(val(...at(o, i)))
      series.push({ name: hasHead ? txt(...at(o, inner[0])) : `Series ${series.length + 1}`, values, xs })
    }
  } else {
    for (let o = firstSeries; o <= outer[1]; o++) {
      const values: (number | null)[] = []
      for (let i = firstPoint; i <= inner[1]; i++) values.push(val(...at(o, i)))
      series.push({ name: hasHead ? txt(...at(o, inner[0])) || `Series ${series.length + 1}` : `Series ${series.length + 1}`, values })
    }
  }
  const fmtCell = at(firstSeries, firstPoint)
  const nf = wb.getStyle(sheetId, fmtCell[0], fmtCell[1])?.numFmt
  return { categories, series, numberFormat: nf && !isDateFormat(nf) ? nf : undefined }
}

/** Guesses series orientation and headers the way Excel does for a selected block. */
export function guessChartLayout(wb: Workbook, sheetId: string, range: Range): Pick<ChartDef, 'seriesIn' | 'firstRowHeader' | 'firstColLabels'> {
  const isText = (r: number, c: number) => typeof wb.getValue(sheetId, r, c) === 'string' || isDateFormat(wb.getStyle(sheetId, r, c)?.numFmt)
  const rows = range.r2 - range.r1 + 1
  const cols = range.c2 - range.c1 + 1
  let firstRowHeader = false
  let firstColLabels = false
  for (let c = range.c1 + (cols > 1 ? 1 : 0); c <= range.c2; c++) if (typeof wb.getValue(sheetId, range.r1, c) === 'string') firstRowHeader = true
  for (let r = range.r1 + (rows > 1 ? 1 : 0); r <= range.r2; r++) if (isText(r, range.c1)) firstColLabels = true
  if (cols === 1) firstColLabels = false
  const dataRows = rows - (firstRowHeader ? 1 : 0)
  const dataCols = cols - (firstColLabels ? 1 : 0)
  return { seriesIn: dataRows >= dataCols ? 'cols' : 'rows', firstRowHeader, firstColLabels }
}

export function newChart(wb: Workbook, ownerSheetId: string, source: { sheetId: string; range: Range }, type: ChartType, place: { x: number; y: number }): ChartDef {
  const layout = guessChartLayout(wb, source.sheetId, source.range)
  const titleCell = layout.firstRowHeader && layout.seriesIn === 'cols' && source.range.c2 - source.range.c1 === (layout.firstColLabels ? 1 : 0)
    ? wb.getDisplay(source.sheetId, source.range.r1, source.range.c2).text
    : ''
  void ownerSheetId
  return {
    id: uid('ch'),
    type,
    source,
    ...layout,
    title: titleCell || 'Chart title',
    xTitle: '',
    yTitle: '',
    legend: type === 'pie' || type === 'doughnut' ? 'right' : 'bottom',
    dataLabels: false,
    gridlines: true,
    stacked: false,
    colors: [...CHART_COLORS],
    bins: 10,
    x: place.x,
    y: place.y,
    w: 480,
    h: 290,
  }
}

export function addChart(wb: Workbook, sheetId: string, chart: ChartDef) {
  wb.run({ label: 'Insert chart', dirty: [sheetId], metaScope: [sheetId], metaOps: () => {
    const m = wb.meta(sheetId)
    m.charts = [...(m.charts ?? []), chart]
  } })
}

export function updateChart(wb: Workbook, sheetId: string, id: string, patch: Partial<ChartDef>, label = 'Edit chart') {
  wb.run({ label, dirty: [sheetId], metaScope: [sheetId], metaOps: () => {
    const m = wb.meta(sheetId)
    m.charts = (m.charts ?? []).map((c) => (c.id === id ? { ...c, ...patch } : c))
  } })
}

export function removeChart(wb: Workbook, sheetId: string, id: string) {
  wb.run({ label: 'Delete chart', dirty: [sheetId], metaScope: [sheetId], metaOps: () => {
    const m = wb.meta(sheetId)
    m.charts = (m.charts ?? []).filter((c) => c.id !== id)
  } })
}

/** Histogram bins for the first series. */
export function histogram(values: (number | null)[], bins: number): { label: string; count: number }[] {
  const nums = values.filter((v): v is number => v !== null)
  if (!nums.length) return []
  const min = Math.min(...nums)
  const max = Math.max(...nums)
  const n = Math.max(1, Math.min(50, Math.round(bins)))
  const width = max === min ? 1 : (max - min) / n
  const counts = new Array(n).fill(0)
  for (const v of nums) counts[Math.min(n - 1, Math.floor((v - min) / width))]++
  const fmt = (x: number) => String(+x.toPrecision(4))
  return counts.map((count, i) => ({ label: `[${fmt(min + i * width)}, ${fmt(min + (i + 1) * width)}${i === n - 1 ? ']' : ')'}`, count }))
}
