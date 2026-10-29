import { useMemo } from 'react'
import type { Workbook } from '../engine/workbook'
import type { Range } from '../model/types'
import { formatGeneral } from '../engine/numberFormat'
import { Icon } from './icons'

export function StatusBar({
  wb,
  sheetId,
  ranges,
  version,
  message,
  zoom,
  setZoom,
  mode,
}: {
  wb: Workbook
  sheetId: string
  ranges: Range[]
  version: number
  message: string | null
  zoom: number
  setZoom: (z: number) => void
  mode: string
}) {
  const stats = useMemo(() => {
    const hfId = wb.hfId(sheetId)
    const dims = wb.hf.getSheetDimensions(hfId)
    let sum = 0
    let count = 0
    let numCount = 0
    let min = Infinity
    let max = -Infinity
    let visited = 0
    let truncated = false
    const many = ranges.some((rg) => rg.r1 !== rg.r2 || rg.c1 !== rg.c2) || ranges.length > 1
    if (!many) return null
    outer: for (const rg of ranges) {
      for (let r = rg.r1; r <= Math.min(rg.r2, dims.height - 1); r++) {
        for (let c = rg.c1; c <= Math.min(rg.c2, dims.width - 1); c++) {
          if (++visited > 200_000) {
            truncated = true
            break outer
          }
          const v = wb.hf.getCellValue({ sheet: hfId, row: r, col: c })
          if (v === null || v === undefined || v === '') continue
          count++
          if (typeof v === 'number') {
            numCount++
            sum += v
            if (v < min) min = v
            if (v > max) max = v
          }
        }
      }
    }
    return { sum, count, numCount, min, max, truncated }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [wb, sheetId, ranges, version])

  const f = (n: number) => formatGeneral(+n.toPrecision(12))
  return (
    <div className="flex h-6 shrink-0 items-center gap-3 border-t border-[var(--border)] bg-[var(--panel)] px-2 text-[11.5px] text-[var(--muted)]" data-testid="status-bar">
      <span className="font-medium text-[var(--text)]">{mode}</span>
      {message && (
        <span className="flex min-w-0 items-center gap-1 truncate text-[var(--text)]" title={message}>
          <Icon name="info" size={13} className="shrink-0 text-[var(--accent)]" />
          <span className="truncate">{message}</span>
        </span>
      )}
      <span className="ml-auto flex items-center gap-3 whitespace-nowrap">
        {stats && stats.numCount > 0 && (
          <>
            <span className="hidden md:inline">Average: {f(stats.sum / stats.numCount)}</span>
            <span>Count: {stats.count}</span>
            <span className="hidden lg:inline">Min: {f(stats.min)}</span>
            <span className="hidden lg:inline">Max: {f(stats.max)}</span>
            <span>Sum: {f(stats.sum)}</span>
          </>
        )}
        {stats && stats.numCount === 0 && stats.count > 0 && <span>Count: {stats.count}</span>}
        {stats?.truncated && <span title="Statistics cover the first 200,000 cells">(partial)</span>}
      </span>
      <span className="flex items-center gap-1">
        <button type="button" aria-label="Zoom out" className="rounded px-1 hover:bg-[var(--hover)]" onClick={() => setZoom(Math.max(0.5, +(zoom - 0.1).toFixed(2)))}>
          <Icon name="minus" size={12} />
        </button>
        <input
          type="range"
          aria-label="Zoom"
          min={50}
          max={200}
          step={10}
          value={Math.round(zoom * 100)}
          onChange={(e) => setZoom(+e.target.value / 100)}
          className="hidden w-24 accent-[var(--accent)] sm:block"
        />
        <button type="button" aria-label="Zoom in" className="rounded px-1 hover:bg-[var(--hover)]" onClick={() => setZoom(Math.min(2, +(zoom + 0.1).toFixed(2)))}>
          <Icon name="plus" size={12} />
        </button>
        <button type="button" className="w-10 rounded text-right hover:bg-[var(--hover)]" onClick={() => setZoom(1)} title="Reset zoom">
          {Math.round(zoom * 100)}%
        </button>
      </span>
    </div>
  )
}
