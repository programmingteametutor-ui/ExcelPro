/** Page layout for printing / PDF: splits the print area into pages (pure, unit-tested). */
import type { PageSetup, Range } from '../model/types'
import { DEFAULT_COL_WIDTH, DEFAULT_PAGE_SETUP, DEFAULT_ROW_HEIGHT } from '../model/types'
import type { Workbook } from '../engine/workbook'

export const PAPER_PT: Record<PageSetup['paper'], [number, number]> = { A4: [595.28, 841.89], Letter: [612, 792], Legal: [612, 1008] }
const PX_TO_PT = 0.75
export const HEADING_W_PT = 30
export const HEADING_H_PT = 15
const HEADER_FOOTER_PT = 18

export interface PageSlice {
  rows: number[] // visible row indices on this page
  cols: number[]
  index: number
}

export interface Layout {
  setup: PageSetup
  pageW: number // points
  pageH: number
  printableW: number
  printableH: number
  scale: number
  area: Range
  pages: PageSlice[]
  colW: (c: number) => number // points (unscaled)
  rowH: (r: number) => number
}

export function printArea(wb: Workbook, sheetId: string, setup: PageSetup): Range | null {
  if (setup.printArea) return setup.printArea
  const u = wb.usedSize(sheetId)
  if (!u.rows || !u.cols) return null
  // include charts' footprint is not needed: charts are drawn over cells inside the area only
  return { r1: 0, c1: 0, r2: u.rows - 1, c2: u.cols - 1 }
}

export function layoutPages(wb: Workbook, sheetId: string, setupIn?: PageSetup): Layout | null {
  const setup = { ...DEFAULT_PAGE_SETUP, ...(setupIn ?? wb.meta(sheetId).pageSetup ?? {}) }
  const area = printArea(wb, sheetId, setup)
  if (!area) return null
  const meta = wb.meta(sheetId)
  const hiddenR = new Set([...meta.hiddenRows, ...(meta.filterHidden ?? [])])
  const hiddenC = new Set(meta.hiddenCols)
  const colW = (c: number) => (hiddenC.has(c) ? 0 : (meta.colWidths[c] ?? DEFAULT_COL_WIDTH) * PX_TO_PT)
  const rowH = (r: number) => (hiddenR.has(r) ? 0 : (meta.rowHeights[r] ?? DEFAULT_ROW_HEIGHT) * PX_TO_PT)
  let [pw, ph] = PAPER_PT[setup.paper]
  if (setup.orientation === 'landscape') [pw, ph] = [ph, pw]
  const m = setup.margins
  const printableW = pw - (m.left + m.right) * 72 - (setup.headings ? HEADING_W_PT : 0)
  const printableH = ph - (m.top + m.bottom) * 72 - (setup.header ? HEADER_FOOTER_PT : 0) - (setup.footer ? HEADER_FOOTER_PT : 0) - (setup.headings ? HEADING_H_PT : 0)
  let scale = Math.max(0.1, Math.min(4, setup.scale / 100))
  if (setup.fitToWidth) {
    let total = 0
    for (let c = area.c1; c <= area.c2; c++) total += colW(c)
    scale = Math.min(1, printableW / Math.max(1, total))
  }
  const split = (from: number, to: number, size: (i: number) => number, avail: number, breaks: number[]) => {
    const pages: number[][] = []
    let cur: number[] = []
    let used = 0
    const br = new Set(breaks)
    for (let i = from; i <= to; i++) {
      const s = size(i) * scale
      if (s === 0) continue
      if (cur.length && (used + s > avail + 0.01 || br.has(i))) {
        pages.push(cur)
        cur = []
        used = 0
      }
      cur.push(i)
      used += s
    }
    if (cur.length) pages.push(cur)
    return pages
  }
  const colPages = split(area.c1, area.c2, colW, printableW, setup.colBreaks)
  const rowPages = split(area.r1, area.r2, rowH, printableH, setup.rowBreaks)
  const pages: PageSlice[] = []
  // Excel default order: down, then over
  for (const cols of colPages) for (const rows of rowPages) pages.push({ rows, cols, index: pages.length })
  return { setup, pageW: pw, pageH: ph, printableW, printableH, scale, area, pages, colW, rowH }
}

/** Row and column indices where automatic or manual page breaks fall (for Page Break Preview). */
export function pageBreaks(layout: Layout): { rows: number[]; cols: number[] } {
  const rows = new Set<number>()
  const cols = new Set<number>()
  for (const p of layout.pages) {
    rows.add(p.rows[0])
    cols.add(p.cols[0])
  }
  rows.delete(layout.area.r1)
  cols.delete(layout.area.c1)
  return { rows: [...rows].sort((a, b) => a - b), cols: [...cols].sort((a, b) => a - b) }
}

export function headerFooterText(t: string | undefined, page: number, pages: number, file: string, sheet: string): string {
  if (!t) return ''
  return t
    .replace(/&P/g, String(page))
    .replace(/&N/g, String(pages))
    .replace(/&D/g, new Date().toLocaleDateString())
    .replace(/&F/g, file)
    .replace(/&A/g, sheet)
}
