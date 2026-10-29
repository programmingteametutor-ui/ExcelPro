/**
 * Chart renderer: pure SVG built with d3-scale / d3-shape. Works in the grid overlay, in print
 * (rendered to static markup) and in tests. Colors come from the chart definition; axis/text
 * colors come from props so the chart reads in light and dark themes.
 */
import { scaleBand, scaleLinear } from 'd3-scale'
import { arc as d3arc, area as d3area, line as d3line, pie as d3pie } from 'd3-shape'
import type { ChartDef } from '../model/types'
import { formatNumber } from '../engine/numberFormat'
import { histogram, type ChartData } from './data'

export interface ChartTheme {
  text: string
  muted: string
  grid: string
  bg: string
}

export const LIGHT_CHART: ChartTheme = { text: '#1f2328', muted: '#5d6670', grid: '#e3e6e9', bg: '#ffffff' }
export const DARK_CHART: ChartTheme = { text: '#e7e9ec', muted: '#a0a7b0', grid: '#3a3f46', bg: '#1e2024' }

function nice(v: number, fmt?: string): string {
  if (fmt) return formatNumber(v, fmt).text
  const a = Math.abs(v)
  if (a >= 1e9) return `${+(v / 1e9).toFixed(1)}B`
  if (a >= 1e6) return `${+(v / 1e6).toFixed(1)}M`
  if (a >= 1e4) return `${+(v / 1e3).toFixed(1)}K`
  return String(+v.toPrecision(4))
}

export function ChartView({ def, data, width, height, theme = LIGHT_CHART }: { def: ChartDef; data: ChartData; width: number; height: number; theme?: ChartTheme }) {
  const hidden = new Set(def.hiddenSeries ?? [])
  const series = data.series.map((s, i) => ({ ...s, i, color: def.colors[i % def.colors.length] })).filter((s) => !hidden.has(s.i))
  const titleH = def.title ? 26 : 8
  const isPie = def.type === 'pie' || def.type === 'doughnut'
  const legendItems = isPie ? data.categories.map((c, i) => ({ name: c, color: def.colors[i % def.colors.length] })) : series.map((s) => ({ name: s.name, color: s.color }))
  const showLegend = def.legend !== 'none' && legendItems.length > (isPie ? 0 : 1) && def.type !== 'histogram'
  const legendH = showLegend && (def.legend === 'top' || def.legend === 'bottom') ? 22 : 0
  const legendW = showLegend && def.legend === 'right' ? Math.min(160, Math.max(70, ...legendItems.map((l) => l.name.length * 6.5 + 22))) : 0
  const yTitleW = def.yTitle ? 18 : 0
  const xTitleH = def.xTitle ? 18 : 0
  const fs = 11
  const plot = {
    x: 10 + yTitleW + (isPie ? 0 : 46),
    y: titleH + (def.legend === 'top' ? legendH : 0) + 6,
    w: 0,
    h: 0,
  }
  plot.w = Math.max(20, width - plot.x - 14 - legendW)
  plot.h = Math.max(20, height - plot.y - (isPie ? 10 : 30) - xTitleH - (def.legend === 'bottom' ? legendH : 0))
  const nodes: React.ReactNode[] = []
  const fmt = data.numberFormat

  if (!series.length && !isPie) {
    nodes.push(<text key="empty" x={width / 2} y={height / 2} textAnchor="middle" fontSize={12} fill={theme.muted}>No numeric data in the selected range</text>)
  } else if (isPie) {
    const s = series[0]
    const vals = (s?.values ?? []).map((v, i) => ({ v: Math.max(0, v ?? 0), i, label: data.categories[i] ?? '' })).filter((x) => x.v > 0)
    const total = vals.reduce((a, b) => a + b.v, 0)
    const r = Math.max(10, Math.min(plot.w, plot.h) / 2)
    const cx = plot.x + plot.w / 2
    const cy = plot.y + plot.h / 2
    const arcs = d3pie<(typeof vals)[number]>().value((d) => d.v).sort(null)(vals)
    const gen = d3arc<(typeof arcs)[number]>().innerRadius(def.type === 'doughnut' ? r * 0.55 : 0).outerRadius(r)
    arcs.forEach((a) => {
      nodes.push(<path key={`a${a.data.i}`} d={gen(a) ?? ''} transform={`translate(${cx},${cy})`} fill={def.colors[a.data.i % def.colors.length]} stroke={theme.bg} strokeWidth={1.5} />)
      if (def.dataLabels && total) {
        const [lx, ly] = gen.centroid(a)
        nodes.push(<text key={`al${a.data.i}`} x={cx + lx} y={cy + ly} textAnchor="middle" dominantBaseline="middle" fontSize={fs} fill="#fff" fontWeight={600}>{`${Math.round((a.data.v / total) * 100)}%`}</text>)
      }
    })
  } else if (def.type === 'scatter') {
    const xs = series.flatMap((s) => s.xs ?? []).filter((v): v is number => v !== null)
    const ys = series.flatMap((s) => s.values).filter((v): v is number => v !== null)
    const x = scaleLinear().domain([Math.min(0, ...xs), Math.max(1, ...xs)]).nice().range([plot.x, plot.x + plot.w])
    const y = scaleLinear().domain([Math.min(0, ...ys), Math.max(1, ...ys)]).nice().range([plot.y + plot.h, plot.y])
    x.ticks(6).forEach((t, i) => {
      if (def.gridlines) nodes.push(<line key={`gx${i}`} x1={x(t)} x2={x(t)} y1={plot.y} y2={plot.y + plot.h} stroke={theme.grid} />)
      nodes.push(<text key={`tx${i}`} x={x(t)} y={plot.y + plot.h + 15} textAnchor="middle" fontSize={fs} fill={theme.muted}>{nice(t)}</text>)
    })
    y.ticks(5).forEach((t, i) => {
      if (def.gridlines) nodes.push(<line key={`gy${i}`} x1={plot.x} x2={plot.x + plot.w} y1={y(t)} y2={y(t)} stroke={theme.grid} />)
      nodes.push(<text key={`ty${i}`} x={plot.x - 6} y={y(t)} textAnchor="end" dominantBaseline="middle" fontSize={fs} fill={theme.muted}>{nice(t, fmt)}</text>)
    })
    series.forEach((s) =>
      s.values.forEach((v, i) => {
        const xv = s.xs?.[i]
        if (v === null || xv === null || xv === undefined) return
        nodes.push(<circle key={`p${s.i}-${i}`} cx={x(xv)} cy={y(v)} r={4} fill={s.color} fillOpacity={0.85} stroke={theme.bg} strokeWidth={1} />)
      }),
    )
  } else if (def.type === 'histogram') {
    const bins = histogram(series[0]?.values ?? [], def.bins ?? 10)
    const x = scaleBand<string>().domain(bins.map((b) => b.label)).range([plot.x, plot.x + plot.w]).paddingInner(0.04)
    const y = scaleLinear().domain([0, Math.max(1, ...bins.map((b) => b.count))]).nice().range([plot.y + plot.h, plot.y])
    y.ticks(5).forEach((t, i) => {
      if (def.gridlines) nodes.push(<line key={`gy${i}`} x1={plot.x} x2={plot.x + plot.w} y1={y(t)} y2={y(t)} stroke={theme.grid} />)
      nodes.push(<text key={`ty${i}`} x={plot.x - 6} y={y(t)} textAnchor="end" dominantBaseline="middle" fontSize={fs} fill={theme.muted}>{nice(t)}</text>)
    })
    const every = Math.max(1, Math.ceil((bins.length * 70) / plot.w))
    bins.forEach((b, i) => {
      nodes.push(<rect key={`h${i}`} x={x(b.label)} y={y(b.count)} width={x.bandwidth()} height={plot.y + plot.h - y(b.count)} fill={def.colors[0]} />)
      if (def.dataLabels && b.count) nodes.push(<text key={`hl${i}`} x={(x(b.label) ?? 0) + x.bandwidth() / 2} y={y(b.count) - 4} textAnchor="middle" fontSize={10} fill={theme.text}>{b.count}</text>)
      if (i % every === 0) nodes.push(<text key={`hx${i}`} x={(x(b.label) ?? 0) + x.bandwidth() / 2} y={plot.y + plot.h + 15} textAnchor="middle" fontSize={9.5} fill={theme.muted}>{b.label}</text>)
    })
  } else {
    const horizontal = def.type === 'bar'
    const cats = data.categories
    const stacked = def.stacked && (def.type === 'column' || def.type === 'bar' || def.type === 'area')
    const totals = cats.map((_, i) => series.reduce((a, s) => a + Math.max(0, s.values[i] ?? 0), 0))
    const negTotals = cats.map((_, i) => series.reduce((a, s) => a + Math.min(0, s.values[i] ?? 0), 0))
    const all = series.flatMap((s) => s.values).filter((v): v is number => v !== null)
    const lo = stacked ? Math.min(0, ...negTotals) : Math.min(0, ...all)
    const hi = stacked ? Math.max(0, ...totals) : Math.max(0, ...all, lo === 0 ? 1 : 0)
    const band = scaleBand<string>().domain(cats.map((c, i) => `${i}\u0001${c}`)).range(horizontal ? [plot.y, plot.y + plot.h] : [plot.x, plot.x + plot.w]).paddingInner(def.type === 'line' || def.type === 'area' ? 0 : 0.25).paddingOuter(def.type === 'line' || def.type === 'area' ? 0 : 0.15)
    const val = scaleLinear().domain([lo, hi]).nice().range(horizontal ? [plot.x, plot.x + plot.w] : [plot.y + plot.h, plot.y])
    const key = (i: number) => `${i}\u0001${cats[i]}`
    val.ticks(5).forEach((t, i) => {
      if (horizontal) {
        if (def.gridlines) nodes.push(<line key={`g${i}`} x1={val(t)} x2={val(t)} y1={plot.y} y2={plot.y + plot.h} stroke={theme.grid} />)
        nodes.push(<text key={`t${i}`} x={val(t)} y={plot.y + plot.h + 15} textAnchor="middle" fontSize={fs} fill={theme.muted}>{nice(t, fmt)}</text>)
      } else {
        if (def.gridlines) nodes.push(<line key={`g${i}`} x1={plot.x} x2={plot.x + plot.w} y1={val(t)} y2={val(t)} stroke={theme.grid} />)
        nodes.push(<text key={`t${i}`} x={plot.x - 6} y={val(t)} textAnchor="end" dominantBaseline="middle" fontSize={fs} fill={theme.muted}>{nice(t, fmt)}</text>)
      }
    })
    const every = Math.max(1, Math.ceil((cats.length * (horizontal ? 16 : 56)) / (horizontal ? plot.h : plot.w)))
    cats.forEach((c, i) => {
      if (i % every) return
      const pos = (band(key(i)) ?? 0) + band.bandwidth() / 2
      const label = c.length > 14 ? c.slice(0, 13) + '…' : c
      if (horizontal) nodes.push(<text key={`c${i}`} x={plot.x - 6} y={pos} textAnchor="end" dominantBaseline="middle" fontSize={fs} fill={theme.muted}>{label}</text>)
      else nodes.push(<text key={`c${i}`} x={def.type === 'line' || def.type === 'area' ? pos : pos} y={plot.y + plot.h + 15} textAnchor="middle" fontSize={fs} fill={theme.muted}>{label}</text>)
    })
    if (def.type === 'column' || def.type === 'bar') {
      const inner = scaleBand<number>().domain(series.map((s) => s.i)).range([0, band.bandwidth()]).paddingInner(0.08)
      const base = cats.map(() => ({ pos: 0, neg: 0 }))
      series.forEach((s) =>
        s.values.forEach((v, i) => {
          if (v === null) return
          const b0 = band(key(i)) ?? 0
          let a: number
          let b: number
          if (stacked) {
            const acc = base[i]
            if (v >= 0) {
              a = val(acc.pos)
              b = val(acc.pos + v)
              acc.pos += v
            } else {
              a = val(acc.neg)
              b = val(acc.neg + v)
              acc.neg += v
            }
          } else {
            a = val(0)
            b = val(v)
          }
          const off = stacked ? 0 : inner(s.i) ?? 0
          const thick = stacked ? band.bandwidth() : inner.bandwidth()
          const rect = horizontal ? { x: Math.min(a, b), y: b0 + off, width: Math.abs(b - a), height: thick } : { x: b0 + off, y: Math.min(a, b), width: thick, height: Math.abs(b - a) }
          nodes.push(<rect key={`r${s.i}-${i}`} {...rect} fill={s.color} rx={1.5} />)
          if (def.dataLabels) {
            const tx = horizontal ? Math.max(a, b) + 4 : rect.x + rect.width / 2
            const ty = horizontal ? rect.y + rect.height / 2 : Math.min(a, b) - 4
            nodes.push(<text key={`l${s.i}-${i}`} x={tx} y={ty} textAnchor={horizontal ? 'start' : 'middle'} dominantBaseline={horizontal ? 'middle' : 'auto'} fontSize={10} fill={theme.text}>{nice(v, fmt)}</text>)
          }
        }),
      )
    } else {
      const pos = (i: number) => (band(key(i)) ?? 0) + band.bandwidth() / 2
      const acc = cats.map(() => 0)
      series.forEach((s) => {
        const pts = s.values.map((v, i) => {
          const y0 = stacked ? acc[i] : 0
          const y1 = stacked ? acc[i] + (v ?? 0) : v
          if (stacked) acc[i] += v ?? 0
          return { i, v, y0, y1 }
        })
        if (def.type === 'area') {
          const gen = d3area<(typeof pts)[number]>().defined((d) => d.v !== null).x((d) => pos(d.i)).y0((d) => val(d.y0)).y1((d) => val(d.y1 ?? 0))
          nodes.push(<path key={`ar${s.i}`} d={gen(pts) ?? ''} fill={s.color} fillOpacity={stacked ? 0.85 : 0.35} stroke={s.color} strokeWidth={1.5} />)
        } else {
          const gen = d3line<(typeof pts)[number]>().defined((d) => d.v !== null).x((d) => pos(d.i)).y((d) => val(d.y1 ?? 0))
          nodes.push(<path key={`ln${s.i}`} d={gen(pts) ?? ''} fill="none" stroke={s.color} strokeWidth={2.2} strokeLinejoin="round" />)
          pts.forEach((pt) => pt.v !== null && nodes.push(<circle key={`pt${s.i}-${pt.i}`} cx={pos(pt.i)} cy={val(pt.y1 ?? 0)} r={cats.length > 40 ? 0 : 3} fill={s.color} />))
        }
        if (def.dataLabels) pts.forEach((pt) => pt.v !== null && nodes.push(<text key={`dl${s.i}-${pt.i}`} x={pos(pt.i)} y={val(pt.y1 ?? 0) - 7} textAnchor="middle" fontSize={10} fill={theme.text}>{nice(pt.v, fmt)}</text>))
      })
    }
    // axis line
    if (horizontal) nodes.push(<line key="ax" x1={val(0)} x2={val(0)} y1={plot.y} y2={plot.y + plot.h} stroke={theme.muted} strokeOpacity={0.6} />)
    else nodes.push(<line key="ax" x1={plot.x} x2={plot.x + plot.w} y1={val(0)} y2={val(0)} stroke={theme.muted} strokeOpacity={0.6} />)
  }

  // legend
  if (showLegend) {
    if (def.legend === 'right') {
      legendItems.forEach((l, i) => {
        const y = plot.y + i * 18
        if (y > height - 10) return
        nodes.push(<rect key={`lg${i}`} x={width - legendW - 4} y={y} width={10} height={10} rx={2} fill={l.color} />)
        nodes.push(<text key={`lt${i}`} x={width - legendW + 12} y={y + 9} fontSize={fs} fill={theme.text}>{l.name.length > 20 ? l.name.slice(0, 19) + '…' : l.name}</text>)
      })
    } else {
      const widths = legendItems.map((l) => Math.min(140, l.name.length * 6.3 + 22))
      const total = widths.reduce((a, b) => a + b, 0)
      let x = Math.max(6, (width - total) / 2)
      const y = def.legend === 'top' ? titleH + 2 : height - legendH + 4
      legendItems.forEach((l, i) => {
        if (x > width - 30) return
        nodes.push(<rect key={`lg${i}`} x={x} y={y + 2} width={10} height={10} rx={2} fill={l.color} />)
        nodes.push(<text key={`lt${i}`} x={x + 14} y={y + 11} fontSize={fs} fill={theme.text}>{l.name.length > 18 ? l.name.slice(0, 17) + '…' : l.name}</text>)
        x += widths[i]
      })
    }
  }
  return (
    <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} xmlns="http://www.w3.org/2000/svg" role="img" aria-label={def.title || 'Chart'} style={{ fontFamily: 'Segoe UI, Inter, system-ui, sans-serif', display: 'block' }}>
      <rect width={width} height={height} fill={theme.bg} />
      {def.title && <text x={width / 2} y={18} textAnchor="middle" fontSize={14} fontWeight={600} fill={theme.text}>{def.title}</text>}
      {nodes}
      {def.yTitle && !isPie && <text transform={`translate(${12},${plot.y + plot.h / 2}) rotate(-90)`} textAnchor="middle" fontSize={fs} fill={theme.muted}>{def.yTitle}</text>}
      {def.xTitle && !isPie && <text x={plot.x + plot.w / 2} y={plot.y + plot.h + 32} textAnchor="middle" fontSize={fs} fill={theme.muted}>{def.xTitle}</text>}
    </svg>
  )
}
