/**
 * Virtualized spreadsheet grid with frozen panes.
 *
 * Rendering: only the visible window is in the DOM. Gridlines are one element per visible row
 * and column; cell elements exist only for cells with content, fill, borders or decorations.
 * Sizes come from AxisMetrics (sparse overrides, O(log n) lookups).
 *
 * Frozen panes: the cell area is split into up to four panes (corner, top, left, main). Each pane
 * renders the same world-coordinate content clipped to its rows/columns and scrolled on the
 * axes that are not frozen.
 *
 * Interaction: a transparent native scroller sits above the panes and receives pointer events
 * (hit-testing uses the axis metrics). Interactive overlays (filter buttons, dropdown arrows,
 * charts) sit above the scroller. Keyboard input goes to a hidden textarea so copy/cut/paste
 * events and IME work; the parent component decides what keys do.
 */
import { forwardRef, useCallback, useEffect, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import { cellKey, colToLetters, inRange, normRange, rangesIntersect } from '../model/address'
import { AxisMetrics } from '../model/axis'
import type { CellAddr, CellStyle, Range, TableDef } from '../model/types'
import { DEFAULT_COL_WIDTH, DEFAULT_FONT, DEFAULT_FONT_SIZE, DEFAULT_ROW_HEIGHT, MAX_COLS, MAX_ROWS } from '../model/types'
import type { Workbook } from '../engine/workbook'
import { cfEvaluator, type CFResult } from '../engine/features/condFormat'
import { tableCellStyle } from '../engine/features/tables'
import { validationAt } from '../engine/features/data'
import { expandForMerges, primaryRange, type EditState, type Selection } from './selection'

export interface GridHandle {
  focus: () => void
  ensureVisible: (a: CellAddr) => void
  pageRows: () => number
  counts: () => { rows: number; cols: number }
  /** Client-space rectangle of a cell (for anchoring popovers). */
  cellClientRect: (r: number, c: number) => DOMRect | null
  /** World pixel position (unzoomed) of the top-left of the visible main area. */
  viewOrigin: () => { x: number; y: number }
}

export interface ObjectsApi {
  zoom: number
  /** world (zoomed) px -> overlay px */
  toScreen: (x: number, y: number) => { x: number; y: number }
  /** Screen rectangle of a cell (merged area included) relative to the objects layer. */
  cellRect: (r: number, c: number) => { x: number; y: number; w: number; h: number }
  viewport: { w: number; h: number }
}

export interface GridProps {
  wb: Workbook
  sheetId: string
  version: number
  selection: Selection
  onSelect: (sel: Selection) => void
  editing: EditState | null
  onEditText: (text: string, caret?: number) => void
  onEditKeyDown: (e: React.KeyboardEvent<HTMLTextAreaElement>) => void
  onStartEdit: (initial: string | null, mode: 'enter' | 'edit') => void
  onInsertRef: (ref: string, replaceFrom: number | null) => void
  pointRef: { start: number; end: number } | null
  canPointRef: () => boolean
  zoom: number
  showGrid: boolean
  showHeadings: boolean
  showFormulas: boolean
  copyRange: Range | null
  onFill: (src: Range, dest: Range) => void
  onContextMenu: (pos: { x: number; y: number }, kind: 'cell' | 'row' | 'col') => void
  onResize: (axis: 'row' | 'col', indices: number[], px: number | null) => void
  onAutoFit: (axis: 'row' | 'col', indices: number[]) => void
  onKeyDown: (e: React.KeyboardEvent<HTMLTextAreaElement>) => void
  onCopy: (e: React.ClipboardEvent) => void
  onCut: (e: React.ClipboardEvent) => void
  onPaste: (e: React.ClipboardEvent) => void
  editorRef: React.RefObject<HTMLTextAreaElement | null>
  // Phase 2/3
  invalidCells?: Set<string> | null
  pageBreaks?: { rows: number[]; cols: number[] } | null
  onFilterButton?: (col: number, rect: DOMRect) => void
  onDropdown?: (r: number, c: number, rect: DOMRect) => void
  onOpenLink?: (url: string) => void
  renderObjects?: (api: ObjectsApi) => ReactNode
  editorOverlay?: ReactNode // e.g. formula autocomplete, positioned under the cell editor
  onEditorRect?: (rect: { x: number; y: number; w: number; h: number } | null) => void
}

const BORDER_CSS: Record<string, string> = {
  thin: '1px solid',
  medium: '2px solid',
  thick: '3px solid',
  dashed: '1px dashed',
  dotted: '1px dotted',
  double: '3px double',
}

const PT_TO_PX = 4 / 3

let measureCtx: CanvasRenderingContext2D | null | undefined
const widthCache = new Map<string, number>()
function measure(text: string, font: string): number {
  const key = font + '|' + text
  const hit = widthCache.get(key)
  if (hit !== undefined) return hit
  if (measureCtx === undefined) measureCtx = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null
  let w = text.length * 7
  if (measureCtx) {
    measureCtx.font = font
    w = measureCtx.measureText(text).width
  }
  if (widthCache.size > 20000) widthCache.clear()
  widthCache.set(key, w)
  return w
}

const ICON_COLORS = ['#c0392b', '#e5a50a', '#2e9d4f']
function CfIcon({ set, level, size }: { set: 'arrows' | 'traffic' | 'flags'; level: 0 | 1 | 2; size: number }) {
  const col = ICON_COLORS[level]
  if (set === 'traffic') return <svg width={size} height={size} viewBox="0 0 10 10" className="shrink-0"><circle cx="5" cy="5" r="4" fill={col} /></svg>
  if (set === 'flags')
    return (
      <svg width={size} height={size} viewBox="0 0 10 10" className="shrink-0">
        <path d="M2 1v8" stroke="#555" strokeWidth="1" />
        <path d="M2.5 1.5h6l-1.5 2 1.5 2h-6z" fill={col} />
      </svg>
    )
  const d = level === 2 ? 'M5 1l4 4H6.5v4h-3V5H1z' : level === 0 ? 'M5 9L1 5h2.5V1h3v4H9z' : 'M9 5L5 1v2.5H1v3h4V9z'
  return <svg width={size} height={size} viewBox="0 0 10 10" className="shrink-0"><path d={d} fill={col} /></svg>
}

type Drag =
  | { kind: 'cells'; extra: boolean }
  | { kind: 'fill'; src: Range; dest: Range }
  | { kind: 'ref'; start: CellAddr; replaceFrom: number }
  | { kind: 'rows'; anchor: number }
  | { kind: 'cols'; anchor: number }
  | { kind: 'resize'; axis: 'row' | 'col'; index: number; startPos: number; startSize: number; size: number }

export const Grid = forwardRef<GridHandle, GridProps>(function Grid(p, ref) {
  const { wb, sheetId, version, selection, zoom } = p
  const meta = wb.meta(sheetId)
  const rootRef = useRef<HTMLDivElement>(null)
  const scrollerRef = useRef<HTMLDivElement>(null)
  const keyRef = useRef<HTMLTextAreaElement>(null)
  const [size, setSize] = useState({ w: 800, h: 500 })
  const [scroll, setScroll] = useState({ x: 0, y: 0 })
  const [extraRows, setExtraRows] = useState(0)
  const [extraCols, setExtraCols] = useState(0)
  const [drag, setDrag] = useState<Drag | null>(null)
  const dragRef = useRef<Drag | null>(null)
  dragRef.current = drag
  const [hoverCursor, setHoverCursor] = useState<string>('cell')
  const [hoverCell, setHoverCell] = useState<CellAddr | null>(null)

  // ---------------------------------------------------------------- dimensions
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const used = useMemo(() => wb.usedSize(sheetId), [wb, sheetId, version])
  const realR = (r: number) => (r >= MAX_ROWS - 1 ? 0 : r)
  const realC = (c: number) => (c >= MAX_COLS - 1 ? 0 : c)
  const selMax = Math.max(realR(selection.anchor.r), realR(selection.focus.r))
  const selMaxC = Math.max(realC(selection.anchor.c), realC(selection.focus.c))
  const rowCount = Math.min(MAX_ROWS, Math.max(1000, used.rows + 200, selMax + 60) + extraRows)
  const colCount = Math.min(MAX_COLS, Math.max(52, used.cols + 26, selMaxC + 10) + extraCols)

  const hiddenRows = useMemo(() => [...meta.hiddenRows, ...(meta.filterHidden ?? [])], [meta.hiddenRows, meta.filterHidden])
  const rowAxis = useMemo(() => {
    const o: Record<number, number> = {}
    for (const [k, v] of Object.entries(meta.rowHeights)) o[+k] = v * zoom
    return new AxisMetrics(Math.round(DEFAULT_ROW_HEIGHT * zoom), rowCount, o, hiddenRows)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [meta.rowHeights, hiddenRows, rowCount, zoom, version])
  const colAxis = useMemo(() => {
    const o: Record<number, number> = {}
    for (const [k, v] of Object.entries(meta.colWidths)) o[+k] = v * zoom
    return new AxisMetrics(Math.round(DEFAULT_COL_WIDTH * zoom), colCount, o, meta.hiddenCols)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [meta.colWidths, meta.hiddenCols, colCount, zoom, version])

  const hw = p.showHeadings ? Math.round(Math.max(40, String(rowCount).length * 8 + 16) * Math.min(1.4, Math.max(0.8, zoom))) : 0
  const hh = p.showHeadings ? Math.round(22 * Math.min(1.4, Math.max(0.8, zoom))) : 0
  const vw = Math.max(0, size.w - hw)
  const vh = Math.max(0, size.h - hh)

  const fr = Math.min(meta.freeze?.rows ?? 0, rowCount - 1)
  const fc = Math.min(meta.freeze?.cols ?? 0, colCount - 1)
  const fW = colAxis.offsetOf(fc)
  const fH = rowAxis.offsetOf(fr)
  const sx = scroll.x
  const sy = scroll.y
  const screenX = useCallback((wx: number) => (wx < fW ? wx : wx - sx), [fW, sx])
  const screenY = useCallback((wy: number) => (wy < fH ? wy : wy - sy), [fH, sy])

  useLayoutEffect(() => {
    const el = rootRef.current
    if (!el) return
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }))
    ro.observe(el)
    setSize({ w: el.clientWidth, h: el.clientHeight })
    return () => ro.disconnect()
  }, [])

  useEffect(() => {
    const s = scrollerRef.current
    if (s) {
      s.scrollLeft = 0
      s.scrollTop = 0
    }
    setScroll({ x: 0, y: 0 })
    setExtraRows(0)
    setExtraCols(0)
  }, [sheetId])

  const raf = useRef(0)
  const onScroll = () => {
    const s = scrollerRef.current
    if (!s) return
    cancelAnimationFrame(raf.current)
    raf.current = requestAnimationFrame(() => {
      setScroll({ x: s.scrollLeft, y: s.scrollTop })
      setHoverCell(null)
      if (s.scrollTop + s.clientHeight > rowAxis.total - 300 && rowCount < MAX_ROWS) setExtraRows((n) => n + 500)
      if (s.scrollLeft + s.clientWidth > colAxis.total - 300 && colCount < MAX_COLS) setExtraCols((n) => n + 26)
    })
  }

  // ---------------------------------------------------------------- geometry helpers
  const rectOf = useCallback(
    (rg: Range) => {
      const x = colAxis.offsetOf(rg.c1)
      const y = rowAxis.offsetOf(rg.r1)
      return { x, y, w: colAxis.offsetOf(Math.min(rg.c2 + 1, colAxis.count)) - x, h: rowAxis.offsetOf(Math.min(rg.r2 + 1, rowAxis.count)) - y }
    },
    [colAxis, rowAxis],
  )
  const mergeOf = useCallback((r: number, c: number) => meta.merges.find((m) => inRange(m, r, c)), [meta.merges])

  const worldFromClient = useCallback(
    (clientX: number, clientY: number, clamp = true) => {
      const s = scrollerRef.current!
      const b = s.getBoundingClientRect()
      let x = clientX - b.left
      let y = clientY - b.top
      if (clamp) {
        x = Math.max(0, x)
        y = Math.max(0, y)
      }
      return { x: x < fW ? x : x + s.scrollLeft, y: y < fH ? y : y + s.scrollTop }
    },
    [fW, fH],
  )

  const cellAt = useCallback(
    (clientX: number, clientY: number): CellAddr => {
      const w = worldFromClient(clientX, clientY)
      return { r: rowAxis.indexAt(w.y), c: colAxis.indexAt(w.x) }
    },
    [rowAxis, colAxis, worldFromClient],
  )

  const ensureVisible = useCallback(
    (a: CellAddr, axes: { x: boolean; y: boolean } = { x: true, y: true }) => {
      const s = scrollerRef.current
      if (!s) return
      a = { r: Math.min(a.r, rowAxis.count - 1), c: Math.min(a.c, colAxis.count - 1) }
      const m = mergeOf(a.r, a.c)
      const rc = rectOf(m ?? { r1: a.r, c1: a.c, r2: a.r, c2: a.c })
      const viewW = s.clientWidth
      const viewH = s.clientHeight
      if (axes.x && a.c >= fc) {
        if (rc.x < s.scrollLeft + fW) s.scrollLeft = Math.max(0, rc.x - fW)
        else if (rc.x + rc.w > s.scrollLeft + viewW) s.scrollLeft = Math.min(rc.x - fW, rc.x + rc.w - viewW)
      }
      if (axes.y && a.r >= fr) {
        if (rc.y < s.scrollTop + fH) s.scrollTop = Math.max(0, rc.y - fH)
        else if (rc.y + rc.h > s.scrollTop + viewH) s.scrollTop = Math.min(rc.y - fH, rc.y + rc.h - viewH)
      }
    },
    [mergeOf, rectOf, rowAxis, colAxis, fc, fr, fW, fH],
  )

  useImperativeHandle(
    ref,
    () => ({
      focus: () => keyRef.current?.focus({ preventScroll: true }),
      ensureVisible,
      pageRows: () => Math.max(1, Math.floor(vh / (DEFAULT_ROW_HEIGHT * zoom)) - 1),
      counts: () => ({ rows: rowCount, cols: colCount }),
      cellClientRect: (r, c) => {
        const s = scrollerRef.current
        if (!s) return null
        const b = s.getBoundingClientRect()
        const m = mergeOf(r, c)
        const rc = rectOf(m ?? { r1: r, c1: c, r2: r, c2: c })
        return new DOMRect(b.left + screenX(rc.x), b.top + screenY(rc.y), rc.w, rc.h)
      },
      viewOrigin: () => ({ x: (fW + sx) / zoom, y: (fH + sy) / zoom }),
    }),
    [ensureVisible, vh, zoom, rowCount, colCount, mergeOf, rectOf, screenX, screenY, fW, fH, sx, sy],
  )

  useEffect(() => {
    if (dragRef.current && dragRef.current.kind !== 'cells') return
    const f = selection.focus
    const fullRow = f.c >= MAX_COLS - 1
    const fullCol = f.r >= MAX_ROWS - 1
    if (fullRow && fullCol) return
    ensureVisible(f, { x: !fullRow, y: !fullCol })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selection.focus.r, selection.focus.c])

  // ---------------------------------------------------------------- visible windows
  const mr0 = Math.max(fr, rowAxis.indexAt(fH + sy) - 1)
  const mr1 = Math.min(rowCount - 1, rowAxis.indexAt(sy + vh) + 2)
  const mc0 = Math.max(fc, colAxis.indexAt(fW + sx) - 1)
  const mc1 = Math.min(colCount - 1, colAxis.indexAt(sx + vw) + 1)

  // ---------------------------------------------------------------- pointer: cells
  const autoScroll = useRef<{ x: number; y: number; timer: number } | null>(null)
  const lastPointer = useRef({ x: 0, y: 0 })

  const updateDrag = useCallback(
    (clientX: number, clientY: number) => {
      const d = dragRef.current
      if (!d) return
      const a = cellAt(clientX, clientY)
      if (d.kind === 'cells') {
        p.onSelect({ ...selection, focus: a })
      } else if (d.kind === 'fill') {
        const src = d.src
        const below = a.r - src.r2
        const above = src.r1 - a.r
        const right = a.c - src.c2
        const left = src.c1 - a.c
        const best = Math.max(below, above, right, left)
        let dest = src
        if (best > 0) {
          if (best === below) dest = { ...src, r2: a.r }
          else if (best === above) dest = { ...src, r1: a.r }
          else if (best === right) dest = { ...src, c2: a.c }
          else dest = { ...src, c1: a.c }
        }
        setDrag({ ...d, dest })
      } else if (d.kind === 'ref') {
        const rg = normRange(d.start, a)
        const text = rg.r1 === rg.r2 && rg.c1 === rg.c2 ? `${colToLetters(rg.c1)}${rg.r1 + 1}` : `${colToLetters(rg.c1)}${rg.r1 + 1}:${colToLetters(rg.c2)}${rg.r2 + 1}`
        p.onInsertRef(text, d.replaceFrom)
      }
    },
    [cellAt, p, selection],
  )

  const stopAutoScroll = () => {
    if (autoScroll.current) {
      clearInterval(autoScroll.current.timer)
      autoScroll.current = null
    }
  }

  const handleAutoScroll = (clientX: number, clientY: number) => {
    const s = scrollerRef.current!
    const b = s.getBoundingClientRect()
    const dx = clientX < b.left ? -1 : clientX > b.right - 14 ? 1 : 0
    const dy = clientY < b.top ? -1 : clientY > b.bottom - 14 ? 1 : 0
    if (!dx && !dy) {
      stopAutoScroll()
      return
    }
    if (!autoScroll.current) {
      autoScroll.current = {
        x: dx,
        y: dy,
        timer: window.setInterval(() => {
          const as = autoScroll.current
          if (!as) return
          s.scrollLeft += as.x * 24
          s.scrollTop += as.y * 24
          updateDrag(lastPointer.current.x, lastPointer.current.y)
        }, 30),
      }
    } else {
      autoScroll.current.x = dx
      autoScroll.current.y = dy
    }
  }

  const prim = expandForMerges(primaryRange(selection), meta.merges)
  const nearFillHandle = (clientX: number, clientY: number) => {
    const w = worldFromClient(clientX, clientY, false)
    const rc = rectOf(prim)
    return Math.abs(w.x - (rc.x + rc.w)) <= 5 && Math.abs(w.y - (rc.y + rc.h)) <= 5
  }

  const onCellPointerDown = (e: React.PointerEvent) => {
    if (e.button !== 0) return
    const s = scrollerRef.current!
    const b = s.getBoundingClientRect()
    if (e.clientX > b.left + s.clientWidth || e.clientY > b.top + s.clientHeight) return // scrollbar
    e.preventDefault()
    const a = cellAt(e.clientX, e.clientY)
    lastPointer.current = { x: e.clientX, y: e.clientY }
    if (p.editing && p.canPointRef()) {
      const replaceFrom = p.pointRef ? p.pointRef.start : null
      p.onInsertRef(`${colToLetters(a.c)}${a.r + 1}`, replaceFrom)
      const start = replaceFrom ?? (p.editorRef.current?.selectionStart ?? p.editing.text.length)
      setDrag({ kind: 'ref', start: a, replaceFrom: start })
      s.setPointerCapture(e.pointerId)
      return
    }
    if (p.editing) p.onEditKeyDown({ key: '__commit__', preventDefault() {}, stopPropagation() {} } as unknown as React.KeyboardEvent<HTMLTextAreaElement>)
    if (!e.shiftKey && nearFillHandle(e.clientX, e.clientY)) {
      setDrag({ kind: 'fill', src: prim, dest: prim })
      s.setPointerCapture(e.pointerId)
      return
    }
    // Ctrl/Cmd+click follows a hyperlink
    const link = meta.links?.[cellKey(a.r, a.c)]
    if (link && (e.ctrlKey || e.metaKey) && p.onOpenLink) {
      p.onOpenLink(link)
      return
    }
    const m = mergeOf(a.r, a.c)
    const anchorCell = m ? { r: m.r1, c: m.c1 } : a
    if (e.shiftKey) p.onSelect({ ...selection, focus: a })
    else if (e.ctrlKey || e.metaKey) p.onSelect({ anchor: anchorCell, focus: m ? { r: m.r2, c: m.c2 } : a, extra: [...selection.extra, primaryRange(selection)] })
    else p.onSelect({ anchor: anchorCell, focus: m ? { r: m.r2, c: m.c2 } : a, extra: [] })
    setDrag({ kind: 'cells', extra: e.ctrlKey || e.metaKey })
    s.setPointerCapture(e.pointerId)
    keyRef.current?.focus({ preventScroll: true })
  }

  const onCellPointerMove = (e: React.PointerEvent) => {
    lastPointer.current = { x: e.clientX, y: e.clientY }
    if (!dragRef.current) {
      const c = nearFillHandle(e.clientX, e.clientY) ? 'crosshair' : 'cell'
      if (c !== hoverCursor) setHoverCursor(c)
      const a = cellAt(e.clientX, e.clientY)
      const k = cellKey(a.r, a.c)
      const want = meta.notes?.[k] || meta.links?.[k] ? a : null
      if ((want?.r ?? -1) !== (hoverCell?.r ?? -1) || (want?.c ?? -1) !== (hoverCell?.c ?? -1)) setHoverCell(want)
      return
    }
    updateDrag(e.clientX, e.clientY)
    handleAutoScroll(e.clientX, e.clientY)
  }

  const onCellPointerUp = () => {
    stopAutoScroll()
    const d = dragRef.current
    setDrag(null)
    if (d?.kind === 'fill' && (d.dest.r1 !== d.src.r1 || d.dest.r2 !== d.src.r2 || d.dest.c1 !== d.src.c1 || d.dest.c2 !== d.src.c2)) p.onFill(d.src, d.dest)
    if (d?.kind === 'ref') p.editorRef.current?.focus()
    else if (!p.editing) keyRef.current?.focus({ preventScroll: true })
  }

  const onDoubleClick = (e: React.MouseEvent) => {
    const a = cellAt(e.clientX, e.clientY)
    if (nearFillHandle(e.clientX, e.clientY)) return
    const m = mergeOf(a.r, a.c)
    const anchorCell = m ? { r: m.r1, c: m.c1 } : a
    p.onSelect({ anchor: anchorCell, focus: anchorCell, extra: [] })
    p.onStartEdit(null, 'edit')
  }

  const onContextMenu = (e: React.MouseEvent) => {
    e.preventDefault()
    const a = cellAt(e.clientX, e.clientY)
    const inSel = [...selection.extra, primaryRange(selection)].some((rg) => inRange(rg, a.r, a.c))
    if (!inSel) p.onSelect({ anchor: a, focus: a, extra: [] })
    p.onContextMenu({ x: e.clientX, y: e.clientY }, 'cell')
  }

  // ---------------------------------------------------------------- pointer: headers
  const RESIZE_ZONE = 4
  const headerHit = (axis: 'row' | 'col', clientPos: number, rectStart: number) => {
    const ax = axis === 'col' ? colAxis : rowAxis
    const local = clientPos - rectStart
    const frozenEdge = axis === 'col' ? fW : fH
    const world = local < frozenEdge ? local : local + (axis === 'col' ? sx : sy)
    const idx = ax.indexAt(world)
    const start = ax.offsetOf(idx)
    const end = start + ax.sizeOf(idx)
    if (end - world <= RESIZE_ZONE) return { idx, resize: idx }
    if (world - start <= RESIZE_ZONE && idx > 0) {
      let prev = idx - 1
      while (prev > 0 && ax.sizeOf(prev) === 0) prev--
      return { idx, resize: prev }
    }
    return { idx, resize: null as number | null }
  }

  const selectedFull = (axis: 'row' | 'col'): number[] => {
    const out: number[] = []
    for (const rg of [...selection.extra, primaryRange(selection)]) {
      const full = axis === 'col' ? rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1 : rg.c1 === 0 && rg.c2 >= MAX_COLS - 1
      if (!full) continue
      const [a, b] = axis === 'col' ? [rg.c1, rg.c2] : [rg.r1, rg.r2]
      for (let i = a; i <= b && out.length < 5000; i++) out.push(i)
    }
    return out
  }

  const onHeaderPointerDown = (axis: 'row' | 'col') => (e: React.PointerEvent<HTMLDivElement>) => {
    if (e.button !== 0) return
    e.preventDefault()
    const el = e.currentTarget
    const b = el.getBoundingClientRect()
    const hit = headerHit(axis, axis === 'col' ? e.clientX : e.clientY, axis === 'col' ? b.left : b.top)
    el.setPointerCapture(e.pointerId)
    if (p.editing) p.onEditKeyDown({ key: '__commit__', preventDefault() {}, stopPropagation() {} } as unknown as React.KeyboardEvent<HTMLTextAreaElement>)
    if (hit.resize !== null) {
      const ax = axis === 'col' ? colAxis : rowAxis
      const sz = ax.sizeOf(hit.resize)
      setDrag({ kind: 'resize', axis, index: hit.resize, startPos: axis === 'col' ? e.clientX : e.clientY, startSize: sz, size: sz })
      return
    }
    const idx = hit.idx
    const full = (i: number) => (axis === 'col' ? { r: 0, c: i } : { r: i, c: 0 })
    const fullEnd = (i: number) => (axis === 'col' ? { r: MAX_ROWS - 1, c: i } : { r: i, c: MAX_COLS - 1 })
    if (e.shiftKey) {
      const anchorIdx = axis === 'col' ? selection.anchor.c : selection.anchor.r
      p.onSelect({ anchor: { ...selection.anchor }, focus: fullEnd(idx), extra: selection.extra })
      setDrag(axis === 'col' ? { kind: 'cols', anchor: anchorIdx } : { kind: 'rows', anchor: anchorIdx })
    } else {
      const extra = e.ctrlKey || e.metaKey ? [...selection.extra, primaryRange(selection)] : []
      p.onSelect({ anchor: full(idx), focus: fullEnd(idx), extra })
      setDrag(axis === 'col' ? { kind: 'cols', anchor: idx } : { kind: 'rows', anchor: idx })
    }
    keyRef.current?.focus({ preventScroll: true })
  }

  const onHeaderPointerMove = (axis: 'row' | 'col') => (e: React.PointerEvent<HTMLDivElement>) => {
    const el = e.currentTarget
    const b = el.getBoundingClientRect()
    const d = dragRef.current
    if (!d) {
      const hit = headerHit(axis, axis === 'col' ? e.clientX : e.clientY, axis === 'col' ? b.left : b.top)
      el.style.cursor = hit.resize !== null ? (axis === 'col' ? 'col-resize' : 'row-resize') : axis === 'col' ? 's-resize' : 'e-resize'
      return
    }
    if (d.kind === 'resize') {
      const delta = (axis === 'col' ? e.clientX : e.clientY) - d.startPos
      setDrag({ ...d, size: Math.max(2, d.startSize + delta) })
      return
    }
    if (d.kind === 'cols' || d.kind === 'rows') {
      const hit = headerHit(axis, axis === 'col' ? e.clientX : e.clientY, axis === 'col' ? b.left : b.top)
      const anchor = d.anchor
      if (axis === 'col') p.onSelect({ anchor: { r: 0, c: anchor }, focus: { r: MAX_ROWS - 1, c: hit.idx }, extra: selection.extra })
      else p.onSelect({ anchor: { r: anchor, c: 0 }, focus: { r: hit.idx, c: MAX_COLS - 1 }, extra: selection.extra })
    }
  }

  const onHeaderPointerUp = (axis: 'row' | 'col') => () => {
    const d = dragRef.current
    setDrag(null)
    if (d?.kind === 'resize') {
      const sel = selectedFull(axis)
      const targets = sel.includes(d.index) ? sel : [d.index]
      const px = Math.round(d.size / zoom)
      if (Math.abs(d.size - d.startSize) >= 1) p.onResize(axis, targets, px)
    }
  }

  const onHeaderDoubleClick = (axis: 'row' | 'col') => (e: React.MouseEvent<HTMLDivElement>) => {
    const b = e.currentTarget.getBoundingClientRect()
    const hit = headerHit(axis, axis === 'col' ? e.clientX : e.clientY, axis === 'col' ? b.left : b.top)
    if (hit.resize === null) return
    const sel = selectedFull(axis)
    p.onAutoFit(axis, sel.includes(hit.resize) ? sel : [hit.resize])
  }

  const onHeaderContextMenu = (axis: 'row' | 'col') => (e: React.MouseEvent<HTMLDivElement>) => {
    e.preventDefault()
    const b = e.currentTarget.getBoundingClientRect()
    const hit = headerHit(axis, axis === 'col' ? e.clientX : e.clientY, axis === 'col' ? b.left : b.top)
    const sel = selectedFull(axis)
    if (!sel.includes(hit.idx)) {
      if (axis === 'col') p.onSelect({ anchor: { r: 0, c: hit.idx }, focus: { r: MAX_ROWS - 1, c: hit.idx }, extra: [] })
      else p.onSelect({ anchor: { r: hit.idx, c: 0 }, focus: { r: hit.idx, c: MAX_COLS - 1 }, extra: [] })
    }
    p.onContextMenu({ x: e.clientX, y: e.clientY }, axis)
  }

  // ---------------------------------------------------------------- decorations (per render)
  const cf = useMemo(() => cfEvaluator(wb, sheetId), [wb, sheetId, version]) // eslint-disable-line react-hooks/exhaustive-deps
  const tables = meta.tables ?? []
  const tableFor = (r: number, c: number): TableDef | undefined => (tables.length ? tables.find((t) => r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2) : undefined)
  const fontScale = zoom * PT_TO_PX
  const showFormulas = p.showFormulas
  const hfId = wb.hfId(sheetId)
  const dims = wb.hf.getSheetDimensions(hfId)
  const notes = meta.notes ?? {}
  const links = meta.links ?? {}

  // ---------------------------------------------------------------- pane renderer
  const renderPane = (pk: string, rA: number, rB: number, cA: number, cB: number): ReactNode[] => {
    const out: ReactNode[] = []
    const overlays: ReactNode[] = []
    if (rB < rA || cB < cA) return out
    const paneRange: Range = { r1: rA, c1: cA, r2: rB, c2: cB }
    const merges = meta.merges.filter((m) => rangesIntersect(m, paneRange))
    const covered = new Set<string>()
    for (const m of merges) for (let r = Math.max(m.r1, rA); r <= Math.min(m.r2, rB); r++) for (let c = Math.max(m.c1, cA); c <= Math.min(m.c2, cB); c++) covered.add(cellKey(r, c))

    const renderCell = (r: number, c: number, rect: { x: number; y: number; w: number; h: number }, merged: boolean) => {
      const k = cellKey(r, c)
      const own: CellStyle | undefined = meta.styles[k]
      const tbl = tableFor(r, c)
      const tstyle = tbl ? tableCellStyle(tbl, r, c) : undefined
      let st: CellStyle | undefined = tstyle ? { ...tstyle, ...own, borders: own?.borders ?? tstyle.borders } : own
      const cfr: CFResult | null = cf ? cf(r, c) : null
      if (cfr?.style || cfr?.scaleFill) st = { ...st, ...cfr.style, ...(cfr.scaleFill ? { fill: cfr.scaleFill } : {}) }
      const d = wb.getDisplay(sheetId, r, c, showFormulas)
      const hasBorder = !!st?.borders && Object.keys(st.borders).length > 0
      const note = notes[k]
      const link = links[k]
      if (d.kind === 'empty' && !st?.fill && !hasBorder && !merged && !note && !cfr?.bar) return
      let width = rect.w
      const wrap = !!st?.wrap
      let hAlign = st?.hAlign ?? (d.kind === 'number' && !showFormulas ? 'right' : d.kind === 'bool' || d.kind === 'error' ? 'center' : 'left')
      if (d.kind === 'number' && showFormulas && !d.isFormula) hAlign = st?.hAlign ?? 'right'
      let overflow = false
      if (!merged && !wrap && d.kind === 'text' && hAlign === 'left' && d.text.length * 6 * zoom > rect.w) {
        let cc = c + 1
        while (cc < colCount && cc <= cB + 2 && width < 4000 && wb.isCellEmpty(sheetId, r, cc) && !mergeOf(r, cc) && !covered.has(cellKey(r, cc))) {
          width += colAxis.sizeOf(cc)
          cc++
        }
        overflow = width > rect.w
      }
      let text = d.text
      const fontPx = (st?.fontSize ?? DEFAULT_FONT_SIZE) * fontScale
      if (d.kind === 'number' && !showFormulas && !wrap && text) {
        const font = `${st?.italic ? 'italic ' : ''}${st?.bold ? '700 ' : ''}${fontPx}px ${st?.fontFamily ? `"${st.fontFamily}", ` : ''}${DEFAULT_FONT}`
        const avail = rect.w - 7 - (cfr?.icon ? fontPx : 0)
        if (measure(text, font) > avail) {
          let fitted: string | null = null
          const v = wb.getValue(sheetId, r, c)
          if (!st?.numFmt && typeof v === 'number') {
            const cands: string[] = []
            if (!Number.isInteger(v) && Math.abs(v) < 1e11) for (let dp = 9; dp >= 0; dp--) cands.push(String(+v.toFixed(dp)))
            for (let kx = 5; kx >= 0; kx--) cands.push(v.toExponential(kx).replace('e', 'E').replace(/E([+-])(\d)$/, 'E$10$2'))
            fitted = cands.find((t) => measure(t, font) <= avail) ?? null
          }
          text = fitted ?? '#'.repeat(Math.max(1, Math.floor(avail / Math.max(1, measure('#', font)))))
        }
      }
      const style: CSSProperties = {
        left: rect.x,
        top: rect.y,
        width: overflow ? width : rect.w,
        height: rect.h,
        fontFamily: st?.fontFamily ? `${JSON.stringify(st.fontFamily)}, ${DEFAULT_FONT}` : DEFAULT_FONT,
        fontSize: fontPx,
        fontWeight: st?.bold ? 700 : undefined,
        fontStyle: st?.italic ? 'italic' : undefined,
        textDecoration: [(st?.underline || (link && !own?.color)) && 'underline', st?.strike && 'line-through'].filter(Boolean).join(' ') || undefined,
        color: d.color ?? st?.color ?? (link ? '#0563c1' : undefined),
        background: st?.fill ?? (overflow || merged ? 'var(--grid-bg)' : undefined),
        justifyContent: hAlign === 'center' ? 'center' : hAlign === 'right' ? 'flex-end' : 'flex-start',
        textAlign: hAlign,
        alignItems: st?.vAlign === 'top' ? 'flex-start' : st?.vAlign === 'middle' ? 'center' : 'flex-end',
        paddingLeft: st?.indent ? 4 + st.indent * 9 * zoom : undefined,
        zIndex: overflow || merged ? 2 : 1,
        paddingBottom: st?.vAlign === 'top' || st?.vAlign === 'middle' ? undefined : 1 * zoom,
        borderRight: merged && p.showGrid ? '1px solid var(--grid-line)' : undefined,
        borderBottom: merged && p.showGrid ? '1px solid var(--grid-line)' : undefined,
        gap: cfr?.icon ? 3 : undefined,
      }
      out.push(
        <div key={`${pk}c${r},${c}`} className={wrap ? 'xp-cell wrap' : 'xp-cell'} style={style} title={d.kind === 'error' && d.error?.message ? d.error.message : undefined}>
          {cfr?.bar && (
            <span
              data-testid="databar" className="pointer-events-none absolute left-[2px]"
              style={{ top: '15%', height: '70%', width: Math.max(1, (rect.w - 4) * cfr.bar.pct), background: `linear-gradient(90deg, ${cfr.bar.negative ? '#e06666' : cfr.bar.color}, ${cfr.bar.negative ? '#f4cccc' : cfr.bar.color}44)`, opacity: 0.9 }}
            />
          )}
          {cfr?.icon && <CfIcon set={cfr.icon.set} level={cfr.icon.level} size={Math.round(fontPx * 0.9)} />}
          <span className="relative">{text}</span>
          {note && <span className="pointer-events-none absolute right-0 top-0" style={{ width: 0, height: 0, borderTop: `${6 * zoom}px solid #d93025`, borderLeft: `${6 * zoom}px solid transparent` }} />}
        </div>,
      )
      if (hasBorder) {
        const b = st!.borders!
        overlays.push(
          <div
            key={`${pk}b${r},${c}`}
            className="pointer-events-none absolute box-border"
            style={{
              left: rect.x - 1,
              top: rect.y - 1,
              width: rect.w + 1,
              height: rect.h + 1,
              zIndex: 3,
              borderTop: b.top ? `${BORDER_CSS[b.top.style]} ${b.top.color}` : undefined,
              borderBottom: b.bottom ? `${BORDER_CSS[b.bottom.style]} ${b.bottom.color}` : undefined,
              borderLeft: b.left ? `${BORDER_CSS[b.left.style]} ${b.left.color}` : undefined,
              borderRight: b.right ? `${BORDER_CSS[b.right.style]} ${b.right.color}` : undefined,
            }}
          />,
        )
      }
    }

    // gridlines
    if (p.showGrid) {
      const left = colAxis.offsetOf(cA)
      const right = colAxis.offsetOf(cB + 1)
      const top = rowAxis.offsetOf(rA)
      const bottom = rowAxis.offsetOf(rB + 1)
      for (let r = rA; r <= rB; r++) {
        if (rowAxis.sizeOf(r) === 0) continue
        out.push(<div key={`${pk}h${r}`} className="absolute" style={{ left, top: rowAxis.offsetOf(r + 1) - 1, width: right - left, height: 1, background: 'var(--grid-line)' }} />)
      }
      for (let c = cA; c <= cB; c++) {
        if (colAxis.sizeOf(c) === 0) continue
        out.push(<div key={`${pk}v${c}`} className="absolute" style={{ left: colAxis.offsetOf(c + 1) - 1, top, width: 1, height: bottom - top, background: 'var(--grid-line)' }} />)
      }
    }
    // cells: anything with content, own style, table area or conditional formats
    const hasCf = !!cf
    for (let r = rA; r <= rB; r++) {
      if (rowAxis.sizeOf(r) === 0) continue
      const y = rowAxis.offsetOf(r)
      const h = rowAxis.sizeOf(r)
      const rowHasData = r < dims.height
      for (let c = cA; c <= cB; c++) {
        const k = cellKey(r, c)
        if (covered.has(k)) continue
        if (!(rowHasData && c < dims.width) && !meta.styles[k] && !notes[k] && !(tables.length && tableFor(r, c)) && !hasCf) continue
        if (colAxis.sizeOf(c) === 0) continue
        renderCell(r, c, { x: colAxis.offsetOf(c), y, w: colAxis.sizeOf(c), h }, false)
      }
    }
    for (const m of merges) renderCell(m.r1, m.c1, rectOf(m), true)
    out.push(...overlays)
    // page breaks
    if (p.pageBreaks) {
      for (const r of p.pageBreaks.rows) if (r >= rA && r <= rB + 1) out.push(<div key={`${pk}pr${r}`} className="pointer-events-none absolute" style={{ left: colAxis.offsetOf(cA), width: colAxis.offsetOf(cB + 1) - colAxis.offsetOf(cA), top: rowAxis.offsetOf(r) - 1, height: 0, borderTop: '2px dashed #3b6fd8', zIndex: 4 }} />)
      for (const c of p.pageBreaks.cols) if (c >= cA && c <= cB + 1) out.push(<div key={`${pk}pc${c}`} className="pointer-events-none absolute" style={{ top: rowAxis.offsetOf(rA), height: rowAxis.offsetOf(rB + 1) - rowAxis.offsetOf(rA), left: colAxis.offsetOf(c) - 1, width: 0, borderLeft: '2px dashed #3b6fd8', zIndex: 4 }} />)
    }
    // invalid data circles
    if (p.invalidCells) {
      for (const k of p.invalidCells) {
        const [r, c] = k.split(',').map(Number)
        if (r < rA || r > rB || c < cA || c > cB) continue
        const rc = rectOf({ r1: r, c1: c, r2: r, c2: c })
        out.push(<div key={`${pk}iv${k}`} className="pointer-events-none absolute rounded-[50%]" style={{ left: rc.x - 3, top: rc.y - 2, width: rc.w + 6, height: rc.h + 4, border: '2px solid #d93025', zIndex: 7 }} />)
      }
    }
    out.push(...selectionNodes(pk))
    return out
  }

  // selection (rendered into every pane, clipped by it)
  const clampR = (rg: Range): Range => ({ r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, rowCount - 1), c2: Math.min(rg.c2, colCount - 1) })
  const am = mergeOf(selection.anchor.r, selection.anchor.c)
  const pr = rectOf(clampR(prim))
  const ar = rectOf(am ?? { r1: selection.anchor.r, c1: selection.anchor.c, r2: selection.anchor.r, c2: selection.anchor.c })
  const selectionNodes = (pk: string): ReactNode[] => {
    const nodes: ReactNode[] = []
    for (const [i, rg0] of selection.extra.entries()) {
      const rc = rectOf(clampR(expandForMerges(rg0, meta.merges)))
      nodes.push(<div key={`${pk}x${i}`} className="pointer-events-none absolute" style={{ left: rc.x, top: rc.y, width: rc.w, height: rc.h, background: 'var(--sel-fill)', zIndex: 5 }} />)
    }
    if (pr.w > 0 && pr.h > 0 && (prim.r1 !== prim.r2 || prim.c1 !== prim.c2) && !(am && am.r1 === prim.r1 && am.c1 === prim.c1 && am.r2 === prim.r2 && am.c2 === prim.c2)) {
      const bands = [
        { x: pr.x, y: pr.y, w: pr.w, h: ar.y - pr.y },
        { x: pr.x, y: ar.y + ar.h, w: pr.w, h: pr.y + pr.h - (ar.y + ar.h) },
        { x: pr.x, y: ar.y, w: ar.x - pr.x, h: ar.h },
        { x: ar.x + ar.w, y: ar.y, w: pr.x + pr.w - (ar.x + ar.w), h: ar.h },
      ]
      bands.forEach((b, i) => b.w > 0 && b.h > 0 && nodes.push(<div key={`${pk}f${i}`} className="pointer-events-none absolute" style={{ left: b.x, top: b.y, width: b.w, height: b.h, background: 'var(--sel-fill)', zIndex: 5 }} />))
    }
    nodes.push(<div key={`${pk}pb`} className="pointer-events-none absolute box-border" style={{ left: pr.x - 1, top: pr.y - 1, width: pr.w + 1, height: pr.h + 1, border: '2px solid var(--sel-border)', zIndex: 6 }} />)
    if (!p.editing && !drag) nodes.push(<div key={`${pk}fh`} className="pointer-events-none absolute box-border" style={{ left: pr.x + pr.w - 4, top: pr.y + pr.h - 4, width: 7, height: 7, background: 'var(--sel-border)', border: '1px solid var(--grid-bg)', zIndex: 7 }} />)
    if (drag?.kind === 'fill') {
      const r = rectOf(clampR(drag.dest))
      nodes.push(<div key={`${pk}fp`} className="pointer-events-none absolute box-border" style={{ left: r.x - 1, top: r.y - 1, width: r.w + 1, height: r.h + 1, border: '2px dashed var(--muted)', zIndex: 7 }} />)
    }
    if (p.copyRange) {
      const r = rectOf(clampR(p.copyRange))
      nodes.push(<div key={`${pk}cp`} className="xp-marquee" style={{ left: r.x - 1, top: r.y - 1, width: r.w + 1, height: r.h + 1, zIndex: 8 }} />)
    }
    return nodes
  }

  // ---------------------------------------------------------------- panes
  const pane = (key: string, clip: { x: number; y: number; w: number; h: number }, tx: number, ty: number, rA: number, rB: number, cA: number, cB: number) =>
    clip.w > 0 && clip.h > 0 ? (
      <div key={key} className="absolute overflow-hidden" style={{ left: clip.x, top: clip.y, width: clip.w, height: clip.h }}>
        <div className="xp-layer" style={{ transform: `translate(${-tx - clip.x}px, ${-ty - clip.y}px)` }}>
          {renderPane(key, rA, rB, cA, cB)}
        </div>
      </div>
    ) : null
  const panes: ReactNode[] = [
    pane('m', { x: fW, y: fH, w: vw - fW, h: vh - fH }, sx, sy, mr0, mr1, mc0, mc1),
    fr > 0 ? pane('t', { x: fW, y: 0, w: vw - fW, h: fH }, sx, 0, 0, fr - 1, mc0, mc1) : null,
    fc > 0 ? pane('l', { x: 0, y: fH, w: fW, h: vh - fH }, 0, sy, mr0, mr1, 0, fc - 1) : null,
    fr > 0 && fc > 0 ? pane('k', { x: 0, y: 0, w: fW, h: fH }, 0, 0, 0, fr - 1, 0, fc - 1) : null,
  ]

  // ---------------------------------------------------------------- headers
  const colHeaders: ReactNode[] = []
  const rowHeaders: ReactNode[] = []
  if (p.showHeadings) {
    const ranges = [...selection.extra, primaryRange(selection)]
    const colSel = (c: number) => ranges.some((rg) => c >= rg.c1 && c <= rg.c2)
    const colFull = (c: number) => ranges.some((rg) => c >= rg.c1 && c <= rg.c2 && rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1)
    const rowSel = (r: number) => ranges.some((rg) => r >= rg.r1 && r <= rg.r2)
    const rowFull = (r: number) => ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && rg.c1 === 0 && rg.c2 >= MAX_COLS - 1)
    const fs = Math.max(10, Math.min(15, 12 * zoom))
    const colHeader = (c: number, frozen: boolean) => {
      const w = colAxis.sizeOf(c)
      if (w === 0) return
      const hiddenBefore = c > 0 && colAxis.sizeOf(c - 1) === 0
      colHeaders.push(
        <div key={c} className={`xp-hdr${colFull(c) ? ' full' : colSel(c) ? ' sel' : ''}`} style={{ left: screenX(colAxis.offsetOf(c)), top: 0, width: w, height: hh, fontSize: fs, zIndex: frozen ? 2 : 1, borderLeft: hiddenBefore ? '2px solid var(--border-strong)' : undefined }}>
          {colToLetters(c)}
        </div>,
      )
    }
    const rowHeader = (r: number, frozen: boolean) => {
      const h = rowAxis.sizeOf(r)
      if (h === 0) return
      const hiddenBefore = r > 0 && rowAxis.sizeOf(r - 1) === 0
      const filtered = meta.filterHidden?.length && hiddenBefore
      rowHeaders.push(
        <div key={r} className={`xp-hdr${rowFull(r) ? ' full' : rowSel(r) ? ' sel' : ''}`} style={{ left: 0, top: screenY(rowAxis.offsetOf(r)), width: hw, height: h, fontSize: fs, zIndex: frozen ? 2 : 1, borderTop: hiddenBefore ? '2px solid var(--border-strong)' : undefined, color: filtered ? '#2f5597' : undefined }}>
          {r + 1}
        </div>,
      )
    }
    for (let c = mc0; c <= mc1; c++) colHeader(c, false)
    for (let c = 0; c < fc; c++) colHeader(c, true)
    for (let r = mr0; r <= mr1; r++) rowHeader(r, false)
    for (let r = 0; r < fr; r++) rowHeader(r, true)
  }

  let guide: ReactNode = null
  if (drag?.kind === 'resize') {
    if (drag.axis === 'col') guide = <div className="pointer-events-none absolute top-0 bottom-0 z-40 w-px bg-[var(--sel-border)]" style={{ left: screenX(colAxis.offsetOf(drag.index)) + drag.size + hw }} />
    else guide = <div className="pointer-events-none absolute left-0 right-0 z-40 h-px bg-[var(--sel-border)]" style={{ top: screenY(rowAxis.offsetOf(drag.index)) + drag.size + hh }} />
  }

  // ---------------------------------------------------------------- editor
  let editor: ReactNode = null
  let editorBox: { x: number; y: number; w: number; h: number } | null = null
  if (p.editing && p.editing.origin === 'cell') {
    const a = selection.anchor
    const m = mergeOf(a.r, a.c)
    const rc = rectOf(m ?? { r1: a.r, c1: a.c, r2: a.r, c2: a.c })
    const st = meta.styles[cellKey(a.r, a.c)]
    const left = screenX(rc.x) - 1 + hw
    const top = screenY(rc.y) - 1 + hh
    const width = Math.max(rc.w + 2, Math.min(600, p.editing.text.split('\n').reduce((mx, l) => Math.max(mx, l.length), 0) * 7.5 * zoom + 16))
    const height = Math.max(rc.h + 2, p.editing.text.split('\n').length * 1.25 * (st?.fontSize ?? DEFAULT_FONT_SIZE) * fontScale + 6)
    editorBox = { x: left, y: top, w: width, h: height }
    editor = (
      <textarea
        ref={p.editorRef}
        className="xp-editor"
        spellCheck={false}
        value={p.editing.text}
        onChange={(e) => p.onEditText(e.target.value, e.target.selectionStart)}
        onKeyDown={p.onEditKeyDown}
        onSelect={(e) => p.onEditText(e.currentTarget.value, e.currentTarget.selectionStart)}
        style={{
          left,
          top,
          minWidth: rc.w + 2,
          width,
          height,
          fontFamily: st?.fontFamily ? `${JSON.stringify(st.fontFamily)}, ${DEFAULT_FONT}` : DEFAULT_FONT,
          fontSize: (st?.fontSize ?? DEFAULT_FONT_SIZE) * fontScale,
          fontWeight: st?.bold ? 700 : undefined,
          fontStyle: st?.italic ? 'italic' : undefined,
        }}
      />
    )
  }
  const onEditorRect = p.onEditorRect
  const eb = editorBox ? `${editorBox.x},${editorBox.y},${editorBox.w},${editorBox.h}` : ''
  useEffect(() => {
    onEditorRect?.(editorBox)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [eb])

  // ---------------------------------------------------------------- interactive overlays
  const controls: ReactNode[] = []
  const af = meta.autoFilter
  const btn = Math.round(16 * Math.min(1.3, Math.max(0.85, zoom)))
  if (af && p.onFilterButton) {
    const r = af.range.r1
    if (rowAxis.sizeOf(r) > 0) {
      for (let c = af.range.c1; c <= af.range.c2; c++) {
        if (colAxis.sizeOf(c) === 0) continue
        if (!(c < fc || (c >= mc0 && c <= mc1))) continue
        if (!(r < fr || (r >= mr0 && r <= mr1))) continue
        const x = screenX(colAxis.offsetOf(c + 1)) - btn - 2
        const y = screenY(rowAxis.offsetOf(r + 1)) - btn - 2
        const active = !!af.criteria[c]
        controls.push(
          <button
            key={`fb${c}`}
            type="button"
            aria-label={`Filter ${colToLetters(c)}`}
            data-testid={`filter-${colToLetters(c)}`}
            className={`pointer-events-auto absolute flex items-center justify-center rounded-[3px] border text-[10px] shadow-sm ${active ? 'border-[var(--accent)] bg-[var(--active)] text-[var(--accent)]' : 'border-[var(--border-strong)] bg-[var(--bg)] text-[var(--muted)]'}`}
            style={{ left: x, top: y, width: btn, height: btn }}
            onPointerDown={(e) => e.stopPropagation()}
            onClick={(e) => p.onFilterButton!(c, (e.currentTarget as HTMLElement).getBoundingClientRect())}
          >
            {active ? (
              <svg width="10" height="10" viewBox="0 0 10 10"><path d="M1 1h8L6 5v4L4 8V5z" fill="currentColor" /></svg>
            ) : (
              <svg width="9" height="9" viewBox="0 0 10 10"><path d="M1 3l4 4 4-4" stroke="currentColor" strokeWidth="1.6" fill="none" /></svg>
            )}
          </button>,
        )
      }
    }
  }
  if (!p.editing && p.onDropdown) {
    const a = selection.anchor
    const v = validationAt(wb, sheetId, a.r, a.c)
    if (v?.type === 'list' && v.showDropdown) {
      const m = mergeOf(a.r, a.c)
      const rc = rectOf(m ?? { r1: a.r, c1: a.c, r2: a.r, c2: a.c })
      controls.push(
        <button
          key="dd"
          type="button"
          aria-label="Show list"
          data-testid="validation-dropdown"
          className="pointer-events-auto absolute flex items-center justify-center rounded-[3px] border border-[var(--border-strong)] bg-[var(--bg)] text-[var(--muted)] shadow-sm"
          style={{ left: screenX(rc.x + rc.w) + 2, top: screenY(rc.y) + rc.h - btn - 1, width: btn, height: btn }}
          onPointerDown={(e) => e.stopPropagation()}
          onClick={(e) => p.onDropdown!(a.r, a.c, (e.currentTarget as HTMLElement).getBoundingClientRect())}
        >
          <svg width="9" height="9" viewBox="0 0 10 10"><path d="M1 3l4 4 4-4" stroke="currentColor" strokeWidth="1.6" fill="none" /></svg>
        </button>,
      )
    }
  }
  if (hoverCell && !drag) {
    const k = cellKey(hoverCell.r, hoverCell.c)
    const rc = rectOf({ r1: hoverCell.r, c1: hoverCell.c, r2: hoverCell.r, c2: hoverCell.c })
    const note = notes[k]
    const link = links[k]
    controls.push(
      <div key="hover" role="tooltip" className="pointer-events-none absolute z-50 max-w-64 rounded-md border border-amber-300 bg-[#fffbe6] px-2 py-1.5 text-[12px] text-[#3d3200] shadow-lg" style={{ left: screenX(rc.x + rc.w) + 6, top: screenY(rc.y) }}>
        {note && <div className="whitespace-pre-wrap">{note}</div>}
        {link && <div className={note ? 'mt-1 border-t border-amber-200 pt-1' : ''}><span className="break-all text-[#0563c1]">{link}</span><div className="text-[11px] text-[#7a6a2a]">Ctrl+click to open</div></div>}
      </div>,
    )
  }
  const objects = p.renderObjects?.({
    zoom,
    toScreen: (x, y) => ({ x: screenX(x), y: screenY(y) }),
    cellRect: (r, c) => {
      const rc = rectOf(mergeOf(r, c) ?? { r1: r, c1: c, r2: r, c2: c })
      return { x: screenX(rc.x), y: screenY(rc.y), w: rc.w, h: rc.h }
    },
    viewport: { w: vw, h: vh },
  })

  const cursor = drag?.kind === 'fill' || hoverCursor === 'crosshair' ? 'crosshair' : 'cell'

  return (
    <div ref={rootRef} className="xp-grid h-full w-full" onPointerDown={() => !p.editing && keyRef.current?.focus({ preventScroll: true })}>
      <div className="absolute overflow-hidden" style={{ left: hw, top: hh, right: 0, bottom: 0 }}>
        {panes}
        {fr > 0 && <div className="pointer-events-none absolute left-0 right-0 z-10 h-px bg-[var(--border-strong)]" style={{ top: fH - 1 }} />}
        {fc > 0 && <div className="pointer-events-none absolute top-0 bottom-0 z-10 w-px bg-[var(--border-strong)]" style={{ left: fW - 1 }} />}
      </div>
      <div
        ref={scrollerRef}
        className="xp-scroller"
        style={{ left: hw, top: hh, right: 0, bottom: 0, cursor }}
        onScroll={onScroll}
        onPointerDown={onCellPointerDown}
        onPointerMove={onCellPointerMove}
        onPointerUp={onCellPointerUp}
        onPointerCancel={onCellPointerUp}
        onPointerLeave={() => setHoverCell(null)}
        onMouseDown={(e) => p.editing && e.preventDefault()}
        onDoubleClick={onDoubleClick}
        onContextMenu={onContextMenu}
        data-testid="grid-scroller"
      >
        <div style={{ width: colAxis.total, height: rowAxis.total }} />
      </div>
      <div className="pointer-events-none absolute z-20 overflow-hidden" style={{ left: hw, top: hh, right: 14, bottom: 14 }}>
        {controls}
        {objects}
      </div>
      {p.showHeadings && (
        <>
          <div
            className="absolute z-20 overflow-hidden"
            style={{ left: hw, top: 0, right: 0, height: hh }}
            onPointerDown={onHeaderPointerDown('col')}
            onPointerMove={onHeaderPointerMove('col')}
            onPointerUp={onHeaderPointerUp('col')}
            onDoubleClick={onHeaderDoubleClick('col')}
            onContextMenu={onHeaderContextMenu('col')}
            data-testid="col-headers"
          >
            {colHeaders}
          </div>
          <div
            className="absolute z-20 overflow-hidden"
            style={{ left: 0, top: hh, width: hw, bottom: 0 }}
            onPointerDown={onHeaderPointerDown('row')}
            onPointerMove={onHeaderPointerMove('row')}
            onPointerUp={onHeaderPointerUp('row')}
            onDoubleClick={onHeaderDoubleClick('row')}
            onContextMenu={onHeaderContextMenu('row')}
            data-testid="row-headers"
          >
            {rowHeaders}
          </div>
          <div
            className="xp-hdr z-20"
            style={{ left: 0, top: 0, width: hw, height: hh }}
            title="Select all"
            onPointerDown={(e) => {
              e.preventDefault()
              p.onSelect({ anchor: { r: 0, c: 0 }, focus: { r: MAX_ROWS - 1, c: MAX_COLS - 1 }, extra: [] })
              keyRef.current?.focus({ preventScroll: true })
            }}
          >
            <svg width="10" height="10" viewBox="0 0 10 10" className="ml-auto mr-1 mt-auto mb-1 opacity-50">
              <path d="M10 0v10H0z" fill="currentColor" />
            </svg>
          </div>
        </>
      )}
      <div className="pointer-events-none absolute inset-0 z-30">
        <div className="pointer-events-auto">{editor}</div>
        {editor && p.editorOverlay}
      </div>
      {guide}
      <textarea
        ref={keyRef}
        className="xp-keytrap"
        aria-label="Spreadsheet grid"
        data-testid="grid-keytrap"
        onKeyDown={p.onKeyDown}
        onCopy={p.onCopy}
        onCut={p.onCut}
        onPaste={p.onPaste}
        onInput={(e) => {
          const v = e.currentTarget.value
          e.currentTarget.value = ''
          if (v && v.trim() && !p.editing) p.onStartEdit(v, 'enter')
        }}
        defaultValue=""
      />
    </div>
  )
})
