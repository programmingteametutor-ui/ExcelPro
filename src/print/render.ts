/**
 * Draws printed pages onto a canvas (used by Print Preview and PDF export). Uses the same
 * effective formatting as the grid: cell styles, table styles and conditional formatting.
 */
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { cellKey, colToLetters, inRange } from '../model/address'
import type { CellStyle, ChartDef } from '../model/types'
import { DEFAULT_FONT_SIZE } from '../model/types'
import type { Workbook } from '../engine/workbook'
import { cfEvaluator } from '../engine/features/condFormat'
import { tableCellStyle } from '../engine/features/tables'
import { ChartView, LIGHT_CHART } from '../charts/ChartView'
import { extractChartData } from '../charts/data'
import { HEADING_H_PT, HEADING_W_PT, headerFooterText, type Layout, type PageSlice } from './paginate'

const FONT = 'Calibri, Carlito, "Segoe UI", Arial, sans-serif'
const BORDER_W: Record<string, number> = { thin: 0.75, medium: 1.5, thick: 2.25, dashed: 0.75, dotted: 0.75, double: 2.25 }

export function effectiveStyle(wb: Workbook, sheetId: string, r: number, c: number, cf: ReturnType<typeof cfEvaluator>): CellStyle | undefined {
  const meta = wb.meta(sheetId)
  const own = meta.styles[cellKey(r, c)]
  const t = meta.tables?.find((x) => inRange(x.range, r, c))
  const ts = t ? tableCellStyle(t, r, c) : undefined
  let st = ts ? { ...ts, ...own, borders: own?.borders ?? ts.borders } : own
  const res = cf?.(r, c)
  if (res?.style || res?.scaleFill) st = { ...st, ...res.style, ...(res.scaleFill ? { fill: res.scaleFill } : {}) }
  return st
}

/** Renders charts to images so they can be drawn on printed pages. */
export async function chartImages(wb: Workbook, sheetId: string): Promise<Map<string, HTMLImageElement>> {
  const out = new Map<string, HTMLImageElement>()
  for (const ch of wb.meta(sheetId).charts ?? []) {
    const svg = renderToStaticMarkup(createElement(ChartView, { def: ch, data: extractChartData(wb, ch), width: ch.w * 2, height: ch.h * 2, theme: LIGHT_CHART }))
      .replace('<svg ', '<svg transform="scale(1)" ')
    const scaled = svg.replace(/viewBox="0 0 [\d.]+ [\d.]+"/, `viewBox="0 0 ${ch.w * 2} ${ch.h * 2}"`)
    const img = new Image()
    const url = URL.createObjectURL(new Blob([scaled], { type: 'image/svg+xml' }))
    await new Promise<void>((resolve) => {
      img.onload = () => resolve()
      img.onerror = () => resolve()
      img.src = url
    })
    out.set(ch.id, img)
  }
  return out
}

export function renderPage(
  wb: Workbook,
  sheetId: string,
  layout: Layout,
  page: PageSlice,
  opts: { pxPerPt: number; fileName: string; images?: Map<string, HTMLImageElement> },
): HTMLCanvasElement {
  const k = opts.pxPerPt
  const canvas = document.createElement('canvas')
  canvas.width = Math.round(layout.pageW * k)
  canvas.height = Math.round(layout.pageH * k)
  const ctx = canvas.getContext('2d')!
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, canvas.width, canvas.height)
  ctx.scale(k, k)
  const s = layout.scale
  const { setup } = layout
  const meta = wb.meta(sheetId)
  const cf = cfEvaluator(wb, sheetId)
  const left0 = setup.margins.left * 72 + (setup.headings ? HEADING_W_PT : 0)
  const top0 = setup.margins.top * 72 + (setup.header ? 18 : 0) + (setup.headings ? HEADING_H_PT : 0)
  const xs: Record<number, number> = {}
  const ys: Record<number, number> = {}
  let x = left0
  for (const c of page.cols) {
    xs[c] = x
    x += layout.colW(c) * s
  }
  const right = x
  let y = top0
  for (const r of page.rows) {
    ys[r] = y
    y += layout.rowH(r) * s
  }
  const bottom = y
  const cw = (c: number) => layout.colW(c) * s
  const rh = (r: number) => layout.rowH(r) * s
  const colSet = new Set(page.cols)
  const rowSet = new Set(page.rows)
  const covered = new Set<string>()
  const merges = meta.merges.filter((m) => page.rows.some((r) => r >= m.r1 && r <= m.r2) && page.cols.some((c) => c >= m.c1 && c <= m.c2))
  for (const m of merges) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (!(r === m.r1 && c === m.c1)) covered.add(cellKey(r, c))
  const box = (r: number, c: number) => {
    const m = merges.find((mm) => mm.r1 === r && mm.c1 === c)
    if (!m) return { x: xs[c], y: ys[r], w: cw(c), h: rh(r) }
    let w = 0
    let h = 0
    for (let cc = m.c1; cc <= m.c2; cc++) if (colSet.has(cc)) w += cw(cc)
    for (let rr = m.r1; rr <= m.r2; rr++) if (rowSet.has(rr)) h += rh(rr)
    return { x: xs[c], y: ys[r], w, h }
  }
  // fills
  for (const r of page.rows)
    for (const c of page.cols) {
      if (covered.has(cellKey(r, c))) continue
      const st = effectiveStyle(wb, sheetId, r, c, cf)
      if (st?.fill) {
        const b = box(r, c)
        ctx.fillStyle = st.fill
        ctx.fillRect(b.x, b.y, b.w, b.h)
      }
    }
  // gridlines
  if (setup.gridlines) {
    ctx.strokeStyle = '#c8c8c8'
    ctx.lineWidth = 0.5
    ctx.beginPath()
    for (const r of [...page.rows, page.rows[page.rows.length - 1] + 1]) {
      const yy = ys[r] ?? bottom
      ctx.moveTo(left0, yy)
      ctx.lineTo(right, yy)
    }
    for (const c of [...page.cols, page.cols[page.cols.length - 1] + 1]) {
      const xx = xs[c] ?? right
      ctx.moveTo(xx, top0)
      ctx.lineTo(xx, bottom)
    }
    ctx.stroke()
  }
  // text + borders
  for (const r of page.rows)
    for (const c of page.cols) {
      if (covered.has(cellKey(r, c))) continue
      const st = effectiveStyle(wb, sheetId, r, c, cf)
      const b = box(r, c)
      const d = wb.getDisplay(sheetId, r, c)
      if (d.text) {
        const size = (st?.fontSize ?? DEFAULT_FONT_SIZE) * s
        ctx.font = `${st?.italic ? 'italic ' : ''}${st?.bold ? '700 ' : ''}${size}px ${st?.fontFamily ? `"${st.fontFamily}", ` : ''}${FONT}`
        ctx.fillStyle = d.color ?? st?.color ?? (meta.links?.[cellKey(r, c)] ? '#0563c1' : '#000000')
        const h = st?.hAlign ?? (d.kind === 'number' ? 'right' : d.kind === 'bool' || d.kind === 'error' ? 'center' : 'left')
        const pad = 2 * s
        const lines = st?.wrap ? wrapText(ctx, d.text, b.w - pad * 2) : d.text.split('\n')
        const lh = size * 1.2
        const blockH = lines.length * lh
        let ty = st?.vAlign === 'top' ? b.y + pad : st?.vAlign === 'middle' ? b.y + (b.h - blockH) / 2 : b.y + b.h - blockH - pad * 0.5
        ctx.save()
        // allow left-aligned text to flow into empty neighbours on the same page
        let clipW = b.w
        if (!st?.wrap && h === 'left' && d.kind === 'text') {
          let cc = c + 1
          while (colSet.has(cc) && wb.isCellEmpty(sheetId, r, cc) && clipW < 2000) {
            clipW += cw(cc)
            cc++
          }
        }
        ctx.beginPath()
        ctx.rect(b.x, b.y, clipW, b.h)
        ctx.clip()
        ctx.textBaseline = 'top'
        for (const line of lines) {
          const tw = ctx.measureText(line).width
          const tx = h === 'right' ? b.x + b.w - pad - tw : h === 'center' ? b.x + (b.w - tw) / 2 : b.x + pad + (st?.indent ?? 0) * 6 * s
          ctx.fillText(line, tx, ty + size * 0.1)
          if (st?.underline || meta.links?.[cellKey(r, c)]) ctx.fillRect(tx, ty + size * 1.05, tw, Math.max(0.5, size / 14))
          if (st?.strike) ctx.fillRect(tx, ty + size * 0.6, tw, Math.max(0.5, size / 14))
          ty += lh
        }
        ctx.restore()
      }
      const brd = st?.borders
      if (brd) {
        for (const side of ['top', 'right', 'bottom', 'left'] as const) {
          const e = brd[side]
          if (!e) continue
          ctx.strokeStyle = e.color
          ctx.lineWidth = BORDER_W[e.style] ?? 0.75
          ctx.setLineDash(e.style === 'dashed' ? [3, 2] : e.style === 'dotted' ? [1, 1.5] : [])
          ctx.beginPath()
          if (side === 'top') (ctx.moveTo(b.x, b.y), ctx.lineTo(b.x + b.w, b.y))
          if (side === 'bottom') (ctx.moveTo(b.x, b.y + b.h), ctx.lineTo(b.x + b.w, b.y + b.h))
          if (side === 'left') (ctx.moveTo(b.x, b.y), ctx.lineTo(b.x, b.y + b.h))
          if (side === 'right') (ctx.moveTo(b.x + b.w, b.y), ctx.lineTo(b.x + b.w, b.y + b.h))
          ctx.stroke()
          ctx.setLineDash([])
        }
      }
    }
  // headings
  if (setup.headings) {
    ctx.font = `${8 * Math.min(1, s + 0.2)}px ${FONT}`
    ctx.fillStyle = '#f0f0f0'
    ctx.fillRect(left0, top0 - HEADING_H_PT, right - left0, HEADING_H_PT)
    ctx.fillRect(left0 - HEADING_W_PT, top0, HEADING_W_PT, bottom - top0)
    ctx.fillStyle = '#555'
    ctx.textBaseline = 'middle'
    ctx.textAlign = 'center'
    for (const c of page.cols) ctx.fillText(colToLetters(c), xs[c] + cw(c) / 2, top0 - HEADING_H_PT / 2)
    for (const r of page.rows) ctx.fillText(String(r + 1), left0 - HEADING_W_PT / 2, ys[r] + rh(r) / 2)
    ctx.textAlign = 'start'
  }
  // charts that start on this page
  if (opts.images) {
    const colStart = (c: number) => {
      let acc = 0
      for (let i = 0; i < c; i++) acc += layout.colW(i)
      return acc
    }
    const rowStart = (r: number) => {
      let acc = 0
      for (let i = 0; i < r; i++) acc += layout.rowH(i)
      return acc
    }
    const ox = colStart(page.cols[0])
    const oy = rowStart(page.rows[0])
    const ex = ox + page.cols.reduce((a, c) => a + layout.colW(c), 0)
    const ey = oy + page.rows.reduce((a, r) => a + layout.rowH(r), 0)
    for (const ch of (meta.charts ?? []) as ChartDef[]) {
      const cx = ch.x * 0.75
      const cy = ch.y * 0.75
      if (cx < ox || cx >= ex || cy < oy || cy >= ey) continue
      const img = opts.images.get(ch.id)
      if (!img || !img.width) continue
      ctx.drawImage(img, left0 + (cx - ox) * s, top0 + (cy - oy) * s, ch.w * 0.75 * s, ch.h * 0.75 * s)
    }
  }
  // header / footer
  ctx.font = `9px ${FONT}`
  ctx.fillStyle = '#333'
  ctx.textAlign = 'center'
  ctx.textBaseline = 'middle'
  const total = layout.pages.length
  if (setup.header) ctx.fillText(headerFooterText(setup.header, page.index + 1, total, opts.fileName, meta.name), layout.pageW / 2, setup.margins.top * 72 - 2)
  if (setup.footer) ctx.fillText(headerFooterText(setup.footer, page.index + 1, total, opts.fileName, meta.name), layout.pageW / 2, layout.pageH - setup.margins.bottom * 72 + 12)
  return canvas
}

function wrapText(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const out: string[] = []
  for (const para of text.split('\n')) {
    let line = ''
    for (const word of para.split(/(\s+)/)) {
      const test = line + word
      if (ctx.measureText(test).width > width && line.trim()) {
        out.push(line.trimEnd())
        line = word.trimStart()
      } else line = test
    }
    out.push(line)
  }
  return out
}
