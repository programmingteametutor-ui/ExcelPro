/**
 * XLSX import/export (Office Open XML SpreadsheetML) using JSZip for the container.
 *
 * Preserved both ways: sheets (names, order, hidden state, tab color, gridline visibility),
 * values, formulas (with cached results on export), shared formulas, defined names, fonts,
 * fills, borders, alignment, wrap, number formats, merged cells, column widths, row heights,
 * hidden rows/columns.
 *
 * Detected and reported but not imported: charts, images, shapes, tables, pivot tables,
 * comments/notes, conditional formatting, data validation, hyperlinks, filters, protection,
 * macros, external links, chart sheets, sparklines.
 */
import JSZip from 'jszip'
import { DetailedCellError } from 'hyperformula'
import { cellKey, colToLetters, lettersToCol, parseRangeA1, quoteSheetName, rangesIntersect, rangeToA1 } from '../model/address'
import type { CellBorders, CellStyle, CFRule, CFStyle, CompareOp, CondFormat, NamedRangeDoc, PageSetup, Range, SheetDoc, TableDef, TotalFn, Validation, WorkbookDoc } from '../model/types'
import { DEFAULT_PAGE_SETUP } from '../model/types'
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from '../model/types'
import { addXlfn, functionNames, shiftFormula, stripXlfn, toExcelFormula } from '../engine/formulaText'
import { normalizeFormula, parseInput, wouldConvert, type RawContent } from '../engine/inputParser'
import { resolveStructuredRefsWith, type TableInfo } from '../engine/features/tables'
import { partsToSerial } from '../engine/numberFormat'
import { newId, type Workbook } from '../engine/workbook'
import { child, children, decodeOoxmlEscapes, deepText, escapeXml, parseXml, type XmlEl } from './xml'
import { compatOf } from '../engine/compat'

export interface ImportReport {
  warnings: string[]
  unsupportedFunctions: { name: string; count: number }[]
  stats: { sheets: number; cells: number; formulas: number }
}

export interface ImportResult {
  doc: WorkbookDoc
  report: ImportReport
}

export const BUILTIN_NUMFMTS: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 5: '$#,##0_);($#,##0)', 6: '$#,##0_);[Red]($#,##0)',
  7: '$#,##0.00_);($#,##0.00)', 8: '$#,##0.00_);[Red]($#,##0.00)', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?',
  13: '# ??/??', 14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM',
  20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm', 37: '#,##0_);(#,##0)', 38: '#,##0_);[Red](#,##0)',
  39: '#,##0.00_);(#,##0.00)', 40: '#,##0.00_);[Red](#,##0.00)', 41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)', 43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
}

const INDEXED = (
  '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
  '800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
  '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
  '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333'
).split(' ')

const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']

function applyTint(hex: string, tint: number): string {
  if (!tint) return hex
  const r = parseInt(hex.slice(0, 2), 16) / 255
  const g = parseInt(hex.slice(2, 4), 16) / 255
  const b = parseInt(hex.slice(4, 6), 16) / 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  let h = 0
  let s = 0
  let l = (max + min) / 2
  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
    h /= 6
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  let rr: number, gg: number, bb: number
  if (s === 0) rr = gg = bb = l
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s
    const p = 2 * l - q
    rr = hue(p, q, h + 1 / 3)
    gg = hue(p, q, h)
    bb = hue(p, q, h - 1 / 3)
  }
  const to = (x: number) => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, '0')
  return (to(rr) + to(gg) + to(bb)).toUpperCase()
}

function resolveColor(el: XmlEl | undefined, theme: string[]): string | undefined {
  if (!el) return undefined
  const a = el.attrs
  let hex: string | undefined
  if (a.rgb) hex = a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb
  else if (a.theme !== undefined) {
    const t = +a.theme
    hex = theme[t] ?? DEFAULT_THEME[t]
  } else if (a.indexed !== undefined) {
    const i = +a.indexed
    if (i === 64 || i === 65) return undefined // system foreground/background
    hex = INDEXED[i]
  } else if (a.auto === '1') return undefined
  if (!hex) return undefined
  if (a.tint) hex = applyTint(hex, parseFloat(a.tint))
  return '#' + hex.toLowerCase()
}

function parseTheme(xml: string | undefined): string[] {
  if (!xml) return DEFAULT_THEME
  const root = parseXml(xml)
  const find = (el: XmlEl, name: string): XmlEl | undefined => {
    if (el.name === name) return el
    for (const c of el.children) {
      const f = find(c, name)
      if (f) return f
    }
    return undefined
  }
  const scheme = find(root, 'clrScheme')
  if (!scheme) return DEFAULT_THEME
  const names = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
  // Indexed by Excel's theme color number: 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4-9=accents (clrScheme stores dk/lt swapped)
  return names.map((n, i) => {
    const el = child(scheme, n)
    const c = el?.children[0]
    if (!c) return DEFAULT_THEME[i]
    return (c.attrs.val && c.name === 'srgbClr' ? c.attrs.val : c.attrs.lastClr) ?? DEFAULT_THEME[i]
  })
}

function mapBorderStyle(s: string | undefined): NonNullable<CellBorders['top']>['style'] | undefined {
  switch (s) {
    case undefined:
    case 'none':
      return undefined
    case 'medium':
    case 'mediumDashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
    case 'slantDashDot':
      return s === 'medium' ? 'medium' : 'dashed'
    case 'thick':
      return 'thick'
    case 'dashed':
    case 'dashDot':
    case 'dashDotDot':
      return 'dashed'
    case 'dotted':
    case 'hair':
      return 'dotted'
    case 'double':
      return 'double'
    default:
      return 'thin'
  }
}

const DEFAULT_FONT_NAME = 'Calibri'
const DEFAULT_FONT_SIZE = 11

function excelWidthToPx(w: number): number {
  return Math.max(0, Math.round(w * 7 + 5))
}
function pxToExcelWidth(px: number): number {
  return Math.max(0, Math.round(((px - 5) / 7) * 100) / 100)
}

function readRaw(zip: JSZip, path: string): Promise<string | undefined> {
  const f = zip.file(path) ?? zip.file(path.replace(/^\//, ''))
  return f ? f.async('string') : Promise.resolve(undefined)
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = base.split('/').slice(0, -1)
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop()
    else if (seg !== '.') parts.push(seg)
  }
  return parts.join('/')
}

async function readRels(zip: JSZip, partPath: string): Promise<Record<string, { target: string; type: string }>> {
  const dir = partPath.split('/').slice(0, -1).join('/')
  const file = partPath.split('/').pop()
  const xml = await readRaw(zip, `${dir}/_rels/${file}.rels`)
  if (!xml) return {}
  const out: Record<string, { target: string; type: string }> = {}
  for (const r of children(parseXml(xml), 'Relationship')) {
    out[r.attrs.Id] = {
      target: r.attrs.TargetMode === 'External' ? r.attrs.Target : resolvePath(partPath, r.attrs.Target),
      type: r.attrs.Type ?? '',
    }
  }
  return out
}

function parseCellRef(ref: string): { r: number; c: number } | null {
  const m = /^([A-Z]+)(\d+)$/i.exec(ref)
  if (!m) return null
  return { r: parseInt(m[2], 10) - 1, c: lettersToCol(m[1]) }
}

/** Reads an .xlsx file into a WorkbookDoc plus an import report. */
export async function importXlsx(data: ArrayBuffer | Uint8Array, fileName = 'Workbook.xlsx'): Promise<ImportResult> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(data)
  } catch {
    throw new Error('This file is not a valid .xlsx workbook (it could not be opened as a zip package). It may be corrupt, password-protected, or a different format.')
  }
  const warnings: string[] = []
  const notImported = new Map<string, number>()
  const note = (what: string, n = 1) => notImported.set(what, (notImported.get(what) ?? 0) + n)

  const wbPath = 'xl/workbook.xml'
  const wbXml = await readRaw(zip, wbPath)
  if (!wbXml) {
    if (zip.file('EncryptedPackage') || zip.file('EncryptionInfo')) throw new Error('This workbook is password-protected (encrypted). Remove the password in Excel and try again.')
    throw new Error('This file does not contain a spreadsheet workbook (xl/workbook.xml is missing).')
  }
  if (zip.file('xl/vbaProject.bin')) note('Macros (VBA)')
  if (Object.keys(zip.files).some((f) => f.startsWith('xl/pivotCache'))) note('Pivot tables')
  if (Object.keys(zip.files).some((f) => f.startsWith('xl/externalLinks'))) note('Links to external workbooks')
  const wbRoot = parseXml(wbXml)
  if (child(wbRoot, 'workbookPr')?.attrs.date1904 === '1' || child(wbRoot, 'workbookPr')?.attrs.date1904 === 'true') {
    warnings.push('This workbook uses the 1904 date system. ExcelPro uses the 1900 system, so dates were shifted to keep the same calendar dates.')
  }
  const date1904 = ['1', 'true'].includes(child(wbRoot, 'workbookPr')?.attrs.date1904 ?? '')
  const rels = await readRels(zip, wbPath)
  const theme = parseTheme(await readRaw(zip, Object.values(rels).find((r) => r.type.endsWith('/theme'))?.target ?? 'xl/theme/theme1.xml'))

  // shared strings
  const sstPath = Object.values(rels).find((r) => r.type.endsWith('/sharedStrings'))?.target ?? 'xl/sharedStrings.xml'
  const sstXml = await readRaw(zip, sstPath)
  const sst: string[] = sstXml ? children(parseXml(sstXml), 'si').map((si) => decodeOoxmlEscapes(deepText(si))) : []

  // styles
  const stylesPath = Object.values(rels).find((r) => r.type.endsWith('/styles'))?.target ?? 'xl/styles.xml'
  const stylesXml = await readRaw(zip, stylesPath)
  const xfStyles: (CellStyle | undefined)[] = []
  const dxfStyles: CFStyle[] = []
  if (stylesXml) {
    const st = parseXml(stylesXml)
    const numFmts: Record<number, string> = { ...BUILTIN_NUMFMTS }
    for (const nf of children(child(st, 'numFmts'), 'numFmt')) numFmts[+nf.attrs.numFmtId] = nf.attrs.formatCode
    const fonts = children(child(st, 'fonts'), 'font').map((f) => {
      const s: CellStyle = {}
      const name = child(f, 'name')?.attrs.val
      const sz = child(f, 'sz')?.attrs.val
      if (name && name !== DEFAULT_FONT_NAME) s.fontFamily = name
      if (sz && +sz !== DEFAULT_FONT_SIZE) s.fontSize = +sz
      const flag = (n: string) => {
        const e = child(f, n)
        return !!e && e.attrs.val !== '0' && e.attrs.val !== 'false' && e.attrs.val !== 'none'
      }
      if (flag('b')) s.bold = true
      if (flag('i')) s.italic = true
      if (flag('u')) s.underline = true
      if (flag('strike')) s.strike = true
      const color = resolveColor(child(f, 'color'), theme)
      if (color && color !== '#000000') s.color = color
      return s
    })
    const fills = children(child(st, 'fills'), 'fill').map((f) => {
      const p = child(f, 'patternFill')
      if (!p || !p.attrs.patternType || p.attrs.patternType === 'none') return undefined
      if (p.attrs.patternType === 'solid') return resolveColor(child(p, 'fgColor'), theme) ?? resolveColor(child(p, 'bgColor'), theme)
      return resolveColor(child(p, 'fgColor'), theme) // patterned fills approximated as solid
    })
    const borders = children(child(st, 'borders'), 'border').map((b) => {
      const out: CellBorders = {}
      for (const side of ['top', 'right', 'bottom', 'left'] as const) {
        const e = child(b, side)
        const style = mapBorderStyle(e?.attrs.style)
        if (style) out[side] = { style, color: resolveColor(child(e, 'color'), theme) ?? '#000000' }
      }
      return Object.keys(out).length ? out : undefined
    })
    for (const dxf of children(child(st, 'dxfs'), 'dxf')) {
      const out: CFStyle = {}
      const f = child(dxf, 'font')
      if (f) {
        const on = (n: string) => !!child(f, n) && !['0', 'false', 'none'].includes(child(f, n)!.attrs.val ?? '1')
        if (on('b')) out.bold = true
        if (on('i')) out.italic = true
        if (on('u')) out.underline = true
        if (on('strike')) out.strike = true
        const c = resolveColor(child(f, 'color'), theme)
        if (c) out.color = c
      }
      const pf = child(child(dxf, 'fill'), 'patternFill')
      if (pf) {
        const c = resolveColor(child(pf, 'bgColor'), theme) ?? resolveColor(child(pf, 'fgColor'), theme)
        if (c) out.fill = c
      }
      dxfStyles.push(out)
    }
    for (const xf of children(child(st, 'cellXfs'), 'xf')) {
      const s: CellStyle = { ...(fonts[+(xf.attrs.fontId ?? 0)] ?? {}) }
      const fill = fills[+(xf.attrs.fillId ?? 0)]
      if (fill) s.fill = fill
      const border = borders[+(xf.attrs.borderId ?? 0)]
      if (border) s.borders = border
      const nfId = +(xf.attrs.numFmtId ?? 0)
      const code = numFmts[nfId]
      if (nfId !== 0 && code && code !== 'General') s.numFmt = code
      const al = child(xf, 'alignment')
      if (al) {
        const h = al.attrs.horizontal
        if (h === 'left' || h === 'right') s.hAlign = h
        else if (h === 'center' || h === 'centerContinuous') s.hAlign = 'center'
        const v = al.attrs.vertical
        if (v === 'top') s.vAlign = 'top'
        else if (v === 'center') s.vAlign = 'middle'
        if (al.attrs.wrapText === '1' || al.attrs.wrapText === 'true') s.wrap = true
        if (al.attrs.indent) s.indent = +al.attrs.indent
        if (al.attrs.textRotation && al.attrs.textRotation !== '0') note('Rotated text (shown horizontally)')
      }
      xfStyles.push(Object.keys(s).length ? s : undefined)
    }
  }

  // sheets
  const sheetDocs: SheetDoc[] = []
  const sheetIndexToDoc = new Map<number, SheetDoc>()
  let activeIdx = +(child(child(wbRoot, 'bookViews'), 'workbookView')?.attrs.activeTab ?? 0)
  let totalCells = 0
  let totalFormulas = 0
  const fnCounts = new Map<string, number>()
  const sheetEls = children(child(wbRoot, 'sheets'), 'sheet')
  let skippedBefore = 0
  for (const [idx, sh] of sheetEls.entries()) {
    const rel = rels[sh.attrs['r:id'] ?? sh.attrs.id]
    const name = sh.attrs.name
    if (!rel || rel.type.endsWith('/chartsheet') || rel.type.endsWith('/dialogsheet') || rel.type.endsWith('/macrosheet')) {
      note(`Chart/dialog sheet "${name}"`)
      if (idx < activeIdx) skippedBefore++
      continue
    }
    const xml = await readRaw(zip, rel.target)
    if (!xml) {
      warnings.push(`Sheet "${name}" could not be read and was skipped.`)
      continue
    }
    const root = parseXml(xml)
    const sheetRels = await readRels(zip, rel.target)
    for (const r of Object.values(sheetRels)) {
      if (r.type.endsWith('/drawing')) note('Charts, images and shapes')
      else if (r.type.endsWith('/threadedComment')) note('Threaded comments (imported as notes)')
      else if (r.type.endsWith('/pivotTable')) note('Pivot table definitions (values kept as cells)')
    }
    if (child(root, 'sheetProtection')) note('Sheet protection')
    if (child(root, 'extLst') && xml.includes('sparklineGroup')) note('Sparklines')

    const meta: SheetDoc = {
      id: newId(),
      name,
      hidden: sh.attrs.state === 'hidden' || sh.attrs.state === 'veryHidden',
      colWidths: {},
      rowHeights: {},
      hiddenRows: [],
      hiddenCols: [],
      styles: {},
      merges: [],
      cells: [],
    }
    const tab = resolveColor(child(child(root, 'sheetPr'), 'tabColor'), theme)
    if (tab) meta.tabColor = tab
    const view = child(child(root, 'sheetViews'), 'sheetView')
    if (view?.attrs.showGridLines === '0' || view?.attrs.showGridLines === 'false') meta.showGridLines = false
    const fmtPr = child(root, 'sheetFormatPr')
    const defaultColPx = fmtPr?.attrs.defaultColWidth ? excelWidthToPx(+fmtPr.attrs.defaultColWidth) : undefined

    for (const col of children(child(root, 'cols'), 'col')) {
      const min = +col.attrs.min - 1
      const max = Math.min(+col.attrs.max - 1, min + 16383)
      const px = col.attrs.width ? excelWidthToPx(+col.attrs.width) : undefined
      const hidden = col.attrs.hidden === '1' || col.attrs.hidden === 'true'
      const styleIdx = col.attrs.style ? +col.attrs.style : 0
      for (let c = min; c <= max; c++) {
        if (px !== undefined && Math.abs(px - DEFAULT_COL_WIDTH) > 2 && max - min < 1000) meta.colWidths[c] = px
        if (hidden && max - min < 1000) meta.hiddenCols.push(c)
      }
      if (styleIdx && max - min > 50) note('Whole-column formatting (applied to used cells only)')
    }
    if (defaultColPx && Math.abs(defaultColPx - DEFAULT_COL_WIDTH) > 2) {
      // apply the sheet default to the first 200 columns that don't have explicit widths
      for (let c = 0; c < 200; c++) if (meta.colWidths[c] === undefined) meta.colWidths[c] = defaultColPx
    }

    const shared = new Map<string, { formula: string; r: number; c: number }>()
    const cells: RawContent[][] = []
    let rowCursor = -1
    let unsupportedArray = 0
    for (const row of children(child(root, 'sheetData'), 'row')) {
      const r = row.attrs.r ? +row.attrs.r - 1 : rowCursor + 1
      rowCursor = r
      if (row.attrs.ht && (row.attrs.customHeight === '1' || row.attrs.customHeight === 'true' || Math.abs(+row.attrs.ht * (4 / 3) - DEFAULT_ROW_HEIGHT) > 3)) {
        meta.rowHeights[r] = Math.round(+row.attrs.ht * (4 / 3))
      }
      if (row.attrs.hidden === '1' || row.attrs.hidden === 'true') meta.hiddenRows.push(r)
      let colCursor = -1
      for (const c of children(row, 'c')) {
        const ref = c.attrs.r ? parseCellRef(c.attrs.r) : { r, c: colCursor + 1 }
        if (!ref) continue
        colCursor = ref.c
        const style = c.attrs.s ? xfStyles[+c.attrs.s] : undefined
        if (style) meta.styles[cellKey(ref.r, ref.c)] = style
        const t = c.attrs.t ?? 'n'
        const fEl = child(c, 'f')
        const vEl = child(c, 'v')
        let raw: RawContent = null
        if (fEl) {
          const ft = fEl.attrs.t
          let formula = fEl.text.trim()
          if (ft === 'shared' && fEl.attrs.si !== undefined) {
            if (formula) shared.set(fEl.attrs.si, { formula, r: ref.r, c: ref.c })
            else {
              const m = shared.get(fEl.attrs.si)
              if (m) formula = shiftFormula('=' + m.formula, ref.r - m.r, ref.c - m.c).slice(1)
            }
          }
          if (ft === 'dataTable') {
            note('What-if data tables (values kept)')
            formula = ''
          }
          if (ft === 'array' && fEl.attrs.ref && fEl.attrs.ref.includes(':')) unsupportedArray++
          if (formula) {
            const f = normalizeFormula('=' + stripXlfn(decodeOoxmlEscapes(formula)))
            raw = f
            totalFormulas++
            for (const fn of functionNames(f)) {
              if (compatOf(fn).status === 'unsupported') fnCounts.set(fn, (fnCounts.get(fn) ?? 0) + 1)
            }
          }
        }
        if (raw === null && (vEl || t === 'inlineStr')) {
          const v = vEl?.text ?? ''
          switch (t) {
            case 's': {
              const s = sst[+v] ?? ''
              raw = s !== '' && (wouldConvert(s) || s.startsWith("'")) ? `'${s}` : s
              break
            }
            case 'inlineStr': {
              const is = child(c, 'is')
              const s = is ? decodeOoxmlEscapes(deepText(is)) : ''
              raw = s !== '' && wouldConvert(s) ? `'${s}` : s
              break
            }
            case 'str': {
              raw = wouldConvert(v) ? `'${v}` : v
              break
            }
            case 'b':
              raw = v === '1' || v === 'true'
              break
            case 'e':
              raw = v // engine parses error literals like #N/A
              break
            case 'd': {
              const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2}))?/.exec(v)
              raw = m ? partsToSerial(+m[1], +m[2], +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0)) : v
              break
            }
            default: {
              const n = Number(v)
              raw = isNaN(n) ? v : date1904 && style?.numFmt ? n + 1462 : n
            }
          }
        }
        if (raw === '') raw = null
        if (raw !== null) {
          totalCells++
          while (cells.length <= ref.r) cells.push([])
          const rowArr = cells[ref.r]
          while (rowArr.length < ref.c) rowArr.push(null)
          rowArr[ref.c] = raw
        }
      }
    }
    if (unsupportedArray) warnings.push(`Sheet "${name}": ${unsupportedArray} legacy array formula(s) (Ctrl+Shift+Enter) were imported as dynamic-array formulas; results may spill differently.`)
    for (const mc of children(child(root, 'mergeCells'), 'mergeCell')) {
      const rg = parseRangeA1(mc.attrs.ref ?? '')
      if (rg && !(rg.r1 === rg.r2 && rg.c1 === rg.c2)) meta.merges.push(rg)
    }
    meta.cells = cells
    await importSheetExtras(zip, root, sheetRels, meta, dxfStyles, note)
    sheetIndexToDoc.set(idx, meta)
    sheetDocs.push(meta)
  }
  activeIdx = Math.max(0, activeIdx - skippedBefore)
  if (!sheetDocs.length) throw new Error('This workbook has no worksheets that ExcelPro can open.')

  // names
  const namedRanges: NamedRangeDoc[] = []
  for (const dn of children(child(wbRoot, 'definedNames'), 'definedName')) {
    const n = dn.attrs.name
    if (!n || n.startsWith('_xlnm.') || n.startsWith('_xlfn.')) {
      if (n === '_xlnm.Print_Area' && dn.attrs.localSheetId !== undefined) {
        const doc = sheetIndexToDoc.get(+dn.attrs.localSheetId)
        const ref = dn.text.split(',')[0].trim()
        const rg = parseRangeA1(ref.slice(ref.lastIndexOf('!') + 1).replace(/\$/g, ''))
        if (doc && rg) doc.pageSetup = { ...DEFAULT_PAGE_SETUP, ...doc.pageSetup, printArea: rg }
      } else if (n === '_xlnm.Print_Titles') note('Print titles')
      continue
    }
    if (dn.attrs.localSheetId !== undefined) {
      note('Sheet-scoped names')
      continue
    }
    if (dn.attrs.hidden === '1') continue
    namedRanges.push({ name: n, expression: normalizeFormula('=' + stripXlfn(dn.text.trim())) })
  }

  resolveImportedStructuredRefs(sheetDocs)

  const unsupportedFunctions = [...fnCounts.entries()].map(([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count)
  if (unsupportedFunctions.length) {
    warnings.push(
      `Formulas using functions ExcelPro doesn't support yet were kept unchanged and show #NAME?: ${unsupportedFunctions.map((f) => `${f.name} (${f.count})`).join(', ')}.`,
    )
  }
  for (const [what, n] of notImported) warnings.push(`Not imported: ${what}${n > 1 ? ` (${n})` : ''}.`)

  const docName = fileName.replace(/\.(xlsx|xlsm|xltx|xltm)$/i, '')
  const now = new Date().toISOString()
  const visible = sheetDocs.filter((s) => !s.hidden)
  const active = sheetDocs[activeIdx] && !sheetDocs[activeIdx].hidden ? sheetDocs[activeIdx] : visible[0] ?? sheetDocs[0]
  if (!visible.length) sheetDocs[0].hidden = false
  return {
    doc: {
      schemaVersion: 1,
      id: newId('wb'),
      name: docName,
      activeSheetId: active.id,
      sheetOrder: sheetDocs.map((s) => s.id),
      sheets: sheetDocs,
      namedRanges,
      createdAt: now,
      updatedAt: now,
    },
    report: { warnings, unsupportedFunctions, stats: { sheets: sheetDocs.length, cells: totalCells, formulas: totalFormulas } },
  }
}

// ====================================================================== export

const BUILTIN_BY_CODE = new Map<string, number>([
  ['0', 1], ['0.00', 2], ['#,##0', 3], ['#,##0.00', 4], ['0%', 9], ['0.00%', 10], ['0.00E+00', 11], ['# ?/?', 12], ['# ??/??', 13], ['@', 49],
])

const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const hex8 = (c: string) => 'FF' + c.replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6)

/** Quotes bare currency symbols ("$" -> "\"$\"") the way Excel writes them, for other apps' sake. */
export function quoteCurrency(code: string): string {
  let out = ''
  let inQ = false
  let inB = false
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '\\' && !inQ) {
      out += ch + (code[i + 1] ?? '')
      i++
      continue
    }
    if (ch === '"') inQ = !inQ
    else if (!inQ && ch === '[') inB = true
    else if (!inQ && ch === ']') inB = false
    if (!inQ && !inB && /[$₹€£¥]/.test(ch)) out += `"${ch}"`
    else out += ch
  }
  return out
}

class StyleTable {
  fonts: string[] = [`<font><sz val="${DEFAULT_FONT_SIZE}"/><color theme="1"/><name val="${DEFAULT_FONT_NAME}"/><family val="2"/><scheme val="minor"/></font>`]
  fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>']
  borders: string[] = ['<border><left/><right/><top/><bottom/><diagonal/></border>']
  numFmts: { id: number; code: string }[] = []
  xfs: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
  private xfIndex = new Map<string, number>()
  dxfs: string[] = []

  dxf(st: CFStyle): number {
    let font = ''
    if (st.bold) font += '<b/>'
    if (st.italic) font += '<i/>'
    if (st.strike) font += '<strike/>'
    if (st.underline) font += '<u/>'
    if (st.color) font += `<color rgb="${hex8(st.color)}"/>`
    const xml = `<dxf>${font ? `<font>${font}</font>` : ''}${st.fill ? `<fill><patternFill patternType="solid"><bgColor rgb="${hex8(st.fill)}"/></patternFill></fill>` : ''}</dxf>`
    const i = this.dxfs.indexOf(xml)
    if (i >= 0) return i
    this.dxfs.push(xml)
    return this.dxfs.length - 1
  }

  private idx(list: string[], xml: string): number {
    const i = list.indexOf(xml)
    if (i >= 0) return i
    list.push(xml)
    return list.length - 1
  }

  numFmtId(code: string | undefined): number {
    if (!code || code.toLowerCase() === 'general') return 0
    code = quoteCurrency(code)
    const b = BUILTIN_BY_CODE.get(code)
    if (b !== undefined) return b
    const ex = this.numFmts.find((n) => n.code === code)
    if (ex) return ex.id
    const id = 164 + this.numFmts.length
    this.numFmts.push({ id, code })
    return id
  }

  xf(s: CellStyle | undefined): number {
    if (!s) return 0
    const key = JSON.stringify(s)
    const hit = this.xfIndex.get(key)
    if (hit !== undefined) return hit
    let font = '<font>'
    if (s.bold) font += '<b/>'
    if (s.italic) font += '<i/>'
    if (s.strike) font += '<strike/>'
    if (s.underline) font += '<u/>'
    font += `<sz val="${s.fontSize ?? DEFAULT_FONT_SIZE}"/>`
    font += s.color ? `<color rgb="${hex8(s.color)}"/>` : '<color theme="1"/>'
    font += `<name val="${escapeXml(s.fontFamily ?? DEFAULT_FONT_NAME)}"/><family val="2"/></font>`
    const fontId = this.idx(this.fonts, font)
    const fillId = s.fill ? this.idx(this.fills, `<fill><patternFill patternType="solid"><fgColor rgb="${hex8(s.fill)}"/><bgColor indexed="64"/></patternFill></fill>`) : 0
    let borderId = 0
    if (s.borders) {
      const side = (n: 'left' | 'right' | 'top' | 'bottom') => {
        const b = s.borders![n]
        return b ? `<${n} style="${b.style}"><color rgb="${hex8(b.color)}"/></${n}>` : `<${n}/>`
      }
      borderId = this.idx(this.borders, `<border>${side('left')}${side('right')}${side('top')}${side('bottom')}<diagonal/></border>`)
    }
    const numFmtId = this.numFmtId(s.numFmt)
    let align = ''
    if (s.hAlign || s.vAlign || s.wrap || s.indent) {
      align = '<alignment'
      if (s.hAlign) align += ` horizontal="${s.hAlign}"`
      if (s.vAlign) align += ` vertical="${s.vAlign === 'middle' ? 'center' : s.vAlign}"`
      if (s.wrap) align += ' wrapText="1"'
      if (s.indent) align += ` indent="${s.indent}"`
      align += '/>'
    }
    const attrs = [
      `numFmtId="${numFmtId}"`, `fontId="${fontId}"`, `fillId="${fillId}"`, `borderId="${borderId}"`, 'xfId="0"',
      numFmtId ? 'applyNumberFormat="1"' : '', fontId ? 'applyFont="1"' : '', fillId ? 'applyFill="1"' : '',
      borderId ? 'applyBorder="1"' : '', align ? 'applyAlignment="1"' : '',
    ].filter(Boolean).join(' ')
    const xml = align ? `<xf ${attrs}>${align}</xf>` : `<xf ${attrs}/>`
    this.xfs.push(xml)
    const i = this.xfs.length - 1
    this.xfIndex.set(key, i)
    return i
  }

  toXml(): string {
    const nf = this.numFmts.length
      ? `<numFmts count="${this.numFmts.length}">${this.numFmts.map((n) => `<numFmt numFmtId="${n.id}" formatCode="${escapeXml(n.code)}"/>`).join('')}</numFmts>`
      : ''
    return (
      xmlHeader +
      `<styleSheet xmlns="${NS_MAIN}">${nf}` +
      `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join('')}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.join('')}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      `<dxfs count="${this.dxfs.length}">${this.dxfs.join('')}</dxfs>` +
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
      '</styleSheet>'
    )
  }
}

function cachedValue(v: unknown): { t?: string; v: string } | null {
  if (v === null || v === undefined) return null
  if (v instanceof DetailedCellError) return { t: 'e', v: v.value === '#CYCLE!' || v.value === '#SPILL!' || v.value === '#LIC!' || v.value === '#ERROR!' ? '#VALUE!' : v.value }
  if (typeof v === 'number') return { v: String(v) }
  if (typeof v === 'boolean') return { t: 'b', v: v ? '1' : '0' }
  return { t: 'str', v: String(v) }
}

/** Writes the workbook to an .xlsx file (formulas exported as formulas with cached results). */
export async function exportXlsx(wb: Workbook): Promise<Uint8Array> {
  const zip = new JSZip()
  const styles = new StyleTable()
  const sst: string[] = []
  const sstIndex = new Map<string, number>()
  let sstCount = 0
  const str = (s: string) => {
    sstCount++
    let i = sstIndex.get(s)
    if (i === undefined) {
      i = sst.length
      sst.push(s)
      sstIndex.set(s, i)
    }
    return i
  }

  const sheets = wb.sheets
  const activeIdx = Math.max(0, wb.order.indexOf(wb.activeSheetId))
  const extraOverrides: string[] = []
  const printAreas: string[] = []
  let needsVml = false
  tableCounter = 0
  commentCounter = 0
  sheets.forEach((meta, idx) => {
    const hfId = wb.hfId(meta.id)
    const ser = wb.hf.getSheetSerialized(hfId)
    const dims = wb.hf.getSheetDimensions(hfId)
    // rows that need writing: content, styles, custom heights, hidden
    const rowsWith = new Map<number, number[]>()
    const addCell = (r: number, c: number) => {
      let list = rowsWith.get(r)
      if (!list) rowsWith.set(r, (list = []))
      list.push(c)
    }
    for (let r = 0; r < ser.length; r++) for (let c = 0; c < ser[r].length; c++) if (ser[r][c] !== null && ser[r][c] !== undefined) addCell(r, c)
    for (const k of Object.keys(meta.styles)) {
      const [r, c] = k.split(',').map(Number)
      if (!(ser[r]?.[c] !== null && ser[r]?.[c] !== undefined)) addCell(r, c)
    }
    const hiddenRowSet = new Set([...meta.hiddenRows, ...(meta.filterHidden ?? [])])
    const extraRows = new Set<number>([...Object.keys(meta.rowHeights).map(Number), ...hiddenRowSet])
    const allRows = [...new Set([...rowsWith.keys(), ...extraRows])].sort((a, b) => a - b)
    let maxC = 0
    const rowXml: string[] = []
    // spilled ranges of dynamic array formulas -> written as array formulas so values show in Excel
    for (const r of allRows) {
      const cols = [...new Set(rowsWith.get(r) ?? [])].sort((a, b) => a - b)
      let attrs = `r="${r + 1}"`
      const h = meta.rowHeights[r]
      if (h !== undefined) attrs += ` ht="${Math.round(h * 0.75 * 100) / 100}" customHeight="1"`
      if (hiddenRowSet.has(r)) attrs += ' hidden="1"'
      let cellsXml = ''
      for (const c of cols) {
        maxC = Math.max(maxC, c)
        const ref = `${colToLetters(c)}${r + 1}`
        const s = styles.xf(meta.styles[cellKey(r, c)])
        const sAttr = s ? ` s="${s}"` : ''
        const raw = ser[r]?.[c]
        if (raw === null || raw === undefined) {
          const addr = { sheet: hfId, row: r, col: c }
          if (r < dims.height && c < dims.width && wb.hf.isCellPartOfArray(addr)) {
            const cv = cachedValue(wb.hf.getCellValue(addr))
            if (cv) {
              cellsXml += cv.t === 'str' ? `<c r="${ref}"${sAttr} t="str"><v>${escapeXml(cv.v)}</v></c>` : `<c r="${ref}"${sAttr}${cv.t ? ` t="${cv.t}"` : ''}><v>${escapeXml(cv.v)}</v></c>`
              continue
            }
          }
          cellsXml += `<c r="${ref}"${sAttr}/>`
          continue
        }
        if (typeof raw === 'string' && raw.startsWith('=')) {
          const addr = { sheet: hfId, row: r, col: c }
          const f = addXlfn(toExcelFormula(raw)).slice(1)
          const cv = cachedValue(wb.hf.getCellValue(addr))
          // detect spill size
          let fAttrs = ''
          if (wb.hf.isCellPartOfArray(addr)) {
            let r2 = r
            let c2 = c
            while (r2 + 1 < dims.height && wb.hf.isCellPartOfArray({ sheet: hfId, row: r2 + 1, col: c }) && !wb.hf.doesCellHaveFormula({ sheet: hfId, row: r2 + 1, col: c })) r2++
            while (c2 + 1 < dims.width && wb.hf.isCellPartOfArray({ sheet: hfId, row: r, col: c2 + 1 }) && !wb.hf.doesCellHaveFormula({ sheet: hfId, row: r, col: c2 + 1 })) c2++
            if (r2 > r || c2 > c) fAttrs = ` t="array" ref="${ref}:${colToLetters(c2)}${r2 + 1}"`
          }
          const tAttr = cv?.t ? ` t="${cv.t}"` : ''
          cellsXml += `<c r="${ref}"${sAttr}${tAttr}><f${fAttrs}>${escapeXml(f)}</f>${cv ? `<v>${escapeXml(cv.v)}</v>` : ''}</c>`
        } else if (typeof raw === 'number') {
          cellsXml += `<c r="${ref}"${sAttr}><v>${raw}</v></c>`
        } else if (typeof raw === 'boolean') {
          cellsXml += `<c r="${ref}"${sAttr} t="b"><v>${raw ? 1 : 0}</v></c>`
        } else {
          let text = String(raw)
          if (/^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!)$/.test(text)) {
            cellsXml += `<c r="${ref}"${sAttr} t="e"><v>${text}</v></c>`
            continue
          }
          if (text.startsWith("'")) text = text.slice(1)
          cellsXml += `<c r="${ref}"${sAttr} t="s"><v>${str(text)}</v></c>`
        }
      }
      rowXml.push(cellsXml ? `<row ${attrs}>${cellsXml}</row>` : `<row ${attrs}/>`)
    }
    const maxR = allRows.length ? allRows[allRows.length - 1] : 0
    const dim = `A1:${colToLetters(Math.max(maxC, 0))}${maxR + 1}`
    // columns
    const colIdx = [...new Set([...Object.keys(meta.colWidths).map(Number), ...meta.hiddenCols])].sort((a, b) => a - b)
    const colsXml = colIdx.length
      ? `<cols>${colIdx
          .map((c) => {
            const w = meta.colWidths[c] ?? DEFAULT_COL_WIDTH
            return `<col min="${c + 1}" max="${c + 1}" width="${pxToExcelWidth(w)}" customWidth="1"${meta.hiddenCols.includes(c) ? ' hidden="1"' : ''}/>`
          })
          .join('')}</cols>`
      : ''
    const extras = buildSheetExtras(wb, meta.id, idx, styles)
    const sheetPrInner = (meta.tabColor ? `<tabColor rgb="${hex8(meta.tabColor)}"/>` : '') + extras.sheetPrInner
    const sheetPr = sheetPrInner ? `<sheetPr>${sheetPrInner}</sheetPr>` : ''
    const viewAttrs = [idx === activeIdx ? 'tabSelected="1"' : '', meta.showGridLines === false ? 'showGridLines="0"' : '', 'workbookViewId="0"'].filter(Boolean).join(' ')
    const merges = meta.merges.length ? `<mergeCells count="${meta.merges.length}">${meta.merges.map((m: Range) => `<mergeCell ref="${rangeToA1(m)}"/>`).join('')}</mergeCells>` : ''
    const xml =
      xmlHeader +
      `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">${sheetPr}<dimension ref="${dim}"/>` +
      `<sheetViews><sheetView ${viewAttrs}>${extras.pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>${colsXml}` +
      `<sheetData>${rowXml.join('')}</sheetData>${extras.autoFilter}${merges}` +
      extras.afterMerge +
      extras.tail +
      '</worksheet>'
    if (extras.rels.length) {
      zip.file(`xl/worksheets/_rels/sheet${idx + 1}.xml.rels`, xmlHeader + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${extras.rels.join('')}</Relationships>`)
    }
    for (const [path, content] of extras.files) zip.file(path, content)
    extraOverrides.push(...extras.overrides)
    if (extras.needsVml) needsVml = true
    if (extras.printArea) printAreas.push(`<definedName name="_xlnm.Print_Area" localSheetId="${idx}">${escapeXml(extras.printArea)}</definedName>`)
    zip.file(`xl/worksheets/sheet${idx + 1}.xml`, xml)
  })

  const names = wb.namedRanges()
  const definedNames = names.length || printAreas.length
    ? `<definedNames>${printAreas.join('')}${names.map((n) => `<definedName name="${escapeXml(n.name)}">${escapeXml(addXlfn(toExcelFormula(n.expression)).replace(/^=/, ''))}</definedName>`).join('')}</definedNames>`
    : ''
  zip.file(
    'xl/workbook.xml',
    xmlHeader +
      `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><workbookPr defaultThemeVersion="164011"/>` +
      `<bookViews><workbookView activeTab="${activeIdx}"/></bookViews><sheets>` +
      sheets.map((s, i) => `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}"${s.hidden ? ' state="hidden"' : ''} r:id="rId${i + 1}"/>`).join('') +
      `</sheets>${definedNames}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`,
  )
  const n = sheets.length
  zip.file(
    'xl/_rels/workbook.xml.rels',
    xmlHeader +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map((_s, i) => `<Relationship Id="rId${i + 1}" Type="${NS_R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
      `<Relationship Id="rId${n + 1}" Type="${NS_R}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId${n + 2}" Type="${NS_R}/sharedStrings" Target="sharedStrings.xml"/>` +
      '</Relationships>',
  )
  zip.file('xl/styles.xml', styles.toXml())
  zip.file(
    'xl/sharedStrings.xml',
    xmlHeader +
      `<sst xmlns="${NS_MAIN}" count="${sstCount}" uniqueCount="${sst.length}">` +
      sst.map((s) => `<si><t${/^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : ''}>${escapeXml(s)}</t></si>`).join('') +
      '</sst>',
  )
  zip.file(
    '[Content_Types].xml',
    xmlHeader +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      (needsVml ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' : '') +
      extraOverrides.join('') +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheets.map((_s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    xmlHeader +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/>` +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      `<Relationship Id="rId3" Type="${NS_R}/extended-properties" Target="docProps/app.xml"/>` +
      '</Relationships>',
  )
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z')
  zip.file(
    'docProps/core.xml',
    xmlHeader +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>${escapeXml(wb.name)}</dc:title><dc:creator>ExcelPro</dc:creator>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
      '</cp:coreProperties>',
  )
  zip.file(
    'docProps/app.xml',
    xmlHeader +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>ExcelPro</Application></Properties>',
  )
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } })
}

export { quoteSheetName }

// ====================================================================== Phase 2/3 parts

const OP_FROM_XML: Record<string, CompareOp> = {
  between: 'between', notBetween: 'notBetween', equal: 'eq', notEqual: 'ne', greaterThan: 'gt', lessThan: 'lt', greaterThanOrEqual: 'gte', lessThanOrEqual: 'lte',
}
const OP_TO_XML: Record<CompareOp, string> = Object.fromEntries(Object.entries(OP_FROM_XML).map(([k, v]) => [v, k])) as Record<CompareOp, string>

function sqrefRanges(sqref: string | undefined): Range[] {
  return (sqref ?? '').split(/\s+/).map((x) => parseRangeA1(x.replace(/\$/g, ''))).filter((x): x is Range => !!x)
}

const fromXmlFormula = (f: string | undefined) => (f === undefined ? undefined : normalizeFormula('=' + stripXlfn(decodeOoxmlEscapes(f.trim()))))

/** Turns an imported formula/value into the string form our rules store (numbers stay plain). */
function ruleValue(f: string | undefined): string | undefined {
  if (f === undefined) return undefined
  const t = f.trim()
  if (/^-?\d+(\.\d+)?(E[+-]?\d+)?$/i.test(t)) return t
  if (/^".*"$/.test(t)) return t.slice(1, -1).replace(/""/g, '"')
  return fromXmlFormula(t)
}

async function importSheetExtras(
  zip: JSZip,
  root: XmlEl,
  sheetRels: Record<string, { target: string; type: string }>,
  meta: SheetDoc,
  dxfs: CFStyle[],
  note: (what: string, n?: number) => void,
) {
  // freeze panes
  const pane = child(child(child(root, 'sheetViews'), 'sheetView'), 'pane')
  if (pane?.attrs.state?.startsWith('frozen')) {
    const rows = Math.round(+(pane.attrs.ySplit ?? 0))
    const cols = Math.round(+(pane.attrs.xSplit ?? 0))
    if (rows || cols) meta.freeze = { rows, cols }
  }
  // autofilter
  const af = child(root, 'autoFilter')
  if (af) {
    const rg = parseRangeA1((af.attrs.ref ?? '').replace(/\$/g, ''))
    if (rg) meta.autoFilter = { range: rg, criteria: {} }
    if (children(af, 'filterColumn').length) note('Saved filter criteria (the filter buttons are kept; reapply the filter)')
  }
  // data validation
  const dvs: Validation[] = []
  for (const dv of children(child(root, 'dataValidations'), 'dataValidation')) {
    const type = (dv.attrs.type ?? 'none') as string
    if (type === 'none') continue
    const f1 = child(dv, 'formula1')?.text
    const f2 = child(dv, 'formula2')?.text
    for (const range of sqrefRanges(dv.attrs.sqref)) {
      const base: Validation = {
        id: `dv${dvs.length}${Math.random().toString(36).slice(2, 6)}`,
        range,
        type: (['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'].includes(type) ? type : 'custom') as Validation['type'],
        op: OP_FROM_XML[dv.attrs.operator ?? 'between'] ?? 'between',
        allowBlank: dv.attrs.allowBlank === '1' || dv.attrs.allowBlank === 'true',
        showDropdown: !(dv.attrs.showDropDown === '1' || dv.attrs.showDropDown === 'true'),
        errorStyle: (dv.attrs.errorStyle as Validation['errorStyle']) ?? 'stop',
        errorTitle: dv.attrs.errorTitle,
        errorMessage: dv.attrs.error,
        inputTitle: dv.attrs.promptTitle,
        inputMessage: dv.attrs.prompt,
      }
      if (base.type === 'list') base.list = f1 && /^".*"$/s.test(f1.trim()) ? f1.trim().slice(1, -1) : fromXmlFormula(f1)
      else if (base.type === 'custom') base.list = fromXmlFormula(f1)
      else {
        base.v1 = ruleValue(f1)
        base.v2 = ruleValue(f2)
      }
      dvs.push(base)
    }
  }
  if (dvs.length) meta.validations = dvs
  // conditional formatting
  const cfs: { prio: number; cf: CondFormat }[] = []
  for (const block of children(root, 'conditionalFormatting')) {
    const ranges = sqrefRanges(block.attrs.sqref)
    for (const rule of children(block, 'cfRule')) {
      const t = rule.attrs.type
      const formulas = children(rule, 'formula').map((f) => f.text)
      const style = rule.attrs.dxfId !== undefined ? dxfs[+rule.attrs.dxfId] : undefined
      let r: CFRule | null = null
      if (t === 'cellIs') r = { type: 'cellIs', op: OP_FROM_XML[rule.attrs.operator ?? 'equal'] ?? 'eq', v1: ruleValue(formulas[0]) ?? '', v2: ruleValue(formulas[1]) }
      else if (t === 'containsText') r = { type: 'text', op: 'contains', text: rule.attrs.text ?? '' }
      else if (t === 'notContainsText') r = { type: 'text', op: 'notContains', text: rule.attrs.text ?? '' }
      else if (t === 'beginsWith') r = { type: 'text', op: 'begins', text: rule.attrs.text ?? '' }
      else if (t === 'endsWith') r = { type: 'text', op: 'ends', text: rule.attrs.text ?? '' }
      else if (t === 'duplicateValues') r = { type: 'duplicate' }
      else if (t === 'uniqueValues') r = { type: 'unique' }
      else if (t === 'containsBlanks') r = { type: 'blank' }
      else if (t === 'notContainsBlanks') r = { type: 'notBlank' }
      else if (t === 'containsErrors') r = { type: 'errors' }
      else if (t === 'expression') r = { type: 'formula', formula: fromXmlFormula(formulas[0]) ?? '=FALSE()' }
      else if (t === 'top10') r = { type: 'top', n: +(rule.attrs.rank ?? 10), percent: rule.attrs.percent === '1', bottom: rule.attrs.bottom === '1' }
      else if (t === 'aboveAverage') r = { type: 'average', above: rule.attrs.aboveAverage !== '0' }
      else if (t === 'colorScale') r = { type: 'colorScale', colors: children(child(rule, 'colorScale'), 'color').map((c) => resolveColor(c, DEFAULT_THEME) ?? '#ffffff') }
      else if (t === 'dataBar') r = { type: 'dataBar', color: resolveColor(child(child(rule, 'dataBar'), 'color'), DEFAULT_THEME) ?? '#638ec6' }
      else if (t === 'iconSet') {
        const set = child(rule, 'iconSet')?.attrs.iconSet ?? '3TrafficLights1'
        r = { type: 'iconSet', set: /Arrow/.test(set) ? 'arrows' : /Flag/.test(set) ? 'flags' : 'traffic' }
      }
      if (!r) {
        note(`Conditional formatting rule type "${t}"`)
        continue
      }
      for (const range of ranges) {
        cfs.push({ prio: +(rule.attrs.priority ?? 999), cf: { id: `cf${cfs.length}${Math.random().toString(36).slice(2, 6)}`, range, rule: r, style, stopIfTrue: rule.attrs.stopIfTrue === '1' } })
      }
    }
  }
  if (cfs.length) meta.condFormats = cfs.sort((a, b) => a.prio - b.prio).map((x) => x.cf)
  // hyperlinks
  const links: Record<string, string> = {}
  for (const h of children(child(root, 'hyperlinks'), 'hyperlink')) {
    const rg = parseRangeA1(h.attrs.ref ?? '')
    if (!rg) continue
    const rid = h.attrs['r:id'] ?? h.attrs.id
    const url = rid && sheetRels[rid] ? sheetRels[rid].target : h.attrs.location ? `#${h.attrs.location}` : undefined
    if (url) links[cellKey(rg.r1, rg.c1)] = url
  }
  if (Object.keys(links).length) meta.links = links
  // notes (comments)
  const notes: Record<string, string> = {}
  for (const rel of Object.values(sheetRels)) {
    if (rel.type.endsWith('/comments')) {
      const x = await readRaw(zip, rel.target)
      if (!x) continue
      for (const cm of children(child(parseXml(x), 'commentList'), 'comment')) {
        const rg = parseRangeA1(cm.attrs.ref ?? '')
        const t = child(cm, 'text')
        if (rg && t) notes[cellKey(rg.r1, rg.c1)] = decodeOoxmlEscapes(deepText(t)).trim()
      }
    }
  }
  if (Object.keys(notes).length) meta.notes = notes
  // tables
  const tables: TableDef[] = []
  for (const tp of children(child(root, 'tableParts'), 'tablePart')) {
    const rel = sheetRels[tp.attrs['r:id'] ?? tp.attrs.id]
    if (!rel) continue
    const x = await readRaw(zip, rel.target)
    if (!x) continue
    const t = parseXml(x)
    const rg = parseRangeA1((t.attrs.ref ?? '').replace(/\$/g, ''))
    if (!rg) continue
    if (t.attrs.headerRowCount === '0') {
      note('Tables without a header row (kept as plain cells)')
      continue
    }
    const totals: Record<number, TotalFn> = {}
    const map: Record<string, TotalFn> = { sum: 'sum', average: 'average', count: 'count', countNums: 'countNumbers', min: 'min', max: 'max' }
    children(child(t, 'tableColumns'), 'tableColumn').forEach((col, i) => {
      const f = col.attrs.totalsRowFunction
      if (f && map[f]) totals[i] = map[f]
    })
    const styleName = child(t, 'tableStyleInfo')?.attrs.name ?? ''
    const num = +(styleName.match(/(\d+)$/)?.[1] ?? 2)
    const palette: TableDef['style'][] = ['gray', 'blue', 'orange', 'gray', 'orange', 'blue', 'green']
    tables.push({
      id: `tbl${tables.length}${Math.random().toString(36).slice(2, 6)}`,
      name: t.attrs.displayName ?? t.attrs.name ?? `Table${tables.length + 1}`,
      range: rg,
      style: palette[num % 7] ?? 'teal',
      banded: child(t, 'tableStyleInfo')?.attrs.showRowStripes !== '0',
      totalRow: +(t.attrs.totalsRowCount ?? 0) > 0,
      totals,
      filter: !!child(t, 'autoFilter'),
    })
  }
  if (tables.length) {
    meta.tables = tables
    if (!meta.autoFilter) {
      const t0 = tables.find((t) => t.filter)
      if (t0) meta.autoFilter = { range: { ...t0.range, r2: t0.totalRow ? t0.range.r2 - 1 : t0.range.r2 }, criteria: {} }
    }
  }
  // page setup
  const ps: Partial<PageSetup> = {}
  const pm = child(root, 'pageMargins')
  if (pm) ps.margins = { top: +(pm.attrs.top ?? 0.75), bottom: +(pm.attrs.bottom ?? 0.75), left: +(pm.attrs.left ?? 0.7), right: +(pm.attrs.right ?? 0.7) }
  const pset = child(root, 'pageSetup')
  if (pset) {
    if (pset.attrs.orientation === 'landscape') ps.orientation = 'landscape'
    const paper = +(pset.attrs.paperSize ?? 9)
    ps.paper = paper === 1 ? 'Letter' : paper === 5 ? 'Legal' : 'A4'
    if (pset.attrs.scale) ps.scale = +pset.attrs.scale
    const fit = child(child(root, 'sheetPr'), 'pageSetUpPr')?.attrs.fitToPage === '1'
    if (fit && (pset.attrs.fitToWidth ?? '1') === '1') ps.fitToWidth = true
  }
  const po = child(root, 'printOptions')
  if (po) {
    ps.gridlines = po.attrs.gridLines === '1' || po.attrs.gridLines === 'true'
    ps.headings = po.attrs.headings === '1' || po.attrs.headings === 'true'
  }
  const hf = child(root, 'headerFooter')
  const cleanHF = (t: string | undefined) => t?.replace(/&[LCR]/g, ' ').replace(/&"[^"]*"/g, '').replace(/&\d+/g, '').trim() || undefined
  if (hf) {
    ps.header = cleanHF(child(hf, 'oddHeader')?.text)
    ps.footer = cleanHF(child(hf, 'oddFooter')?.text)
  }
  const rb = children(child(root, 'rowBreaks'), 'brk').map((b) => +(b.attrs.id ?? 0)).filter((x) => x > 0)
  const cb = children(child(root, 'colBreaks'), 'brk').map((b) => +(b.attrs.id ?? 0)).filter((x) => x > 0)
  if (rb.length) ps.rowBreaks = rb
  if (cb.length) ps.colBreaks = cb
  if (Object.keys(ps).length) meta.pageSetup = { ...DEFAULT_PAGE_SETUP, ...(hf ? { footer: undefined } : {}), ...ps }
}

/** After import, turns structured references (Table1[Col]) in formulas into plain ranges. */
function resolveImportedStructuredRefs(docs: SheetDoc[]) {
  const infos: TableInfo[] = []
  for (const d of docs)
    for (const t of d.tables ?? []) {
      const headers: string[] = []
      for (let c = t.range.c1; c <= t.range.c2; c++) {
        const v = d.cells[t.range.r1]?.[c]
        headers.push(v === null || v === undefined ? `Column${c - t.range.c1 + 1}` : String(v).replace(/^'/, ''))
      }
      infos.push({ name: t.name, sheetName: d.name, range: t.range, headers, totalRow: t.totalRow })
    }
  if (!infos.length) return
  for (const d of docs)
    d.cells.forEach((row, r) =>
      row.forEach((v, c) => {
        if (typeof v === 'string' && v.startsWith('=') && (v.includes('[') || infos.some((t) => v.toLowerCase().includes(t.name.toLowerCase())))) {
          row[c] = resolveStructuredRefsWith(v, infos, d.name, r)
        }
      }),
    )
}

// ---------------------------------------------------------------------- export of extras

export interface SheetExtrasXml {
  pane: string
  sheetPrInner: string
  autoFilter: string
  afterMerge: string
  tail: string
  rels: string[]
  files: [string, string][]
  overrides: string[]
  printArea?: string
  needsVml: boolean
}

let tableCounter = 0
let commentCounter = 0

function numOrFormula(v: string | undefined): string {
  if (v === undefined || v === '') return ''
  if (v.startsWith('=')) return escapeXml(addXlfn(toExcelFormula(v)).slice(1))
  const p = parseInput(v)
  if (typeof p.raw === 'number') return String(p.raw)
  return escapeXml(`"${v.replace(/"/g, '""')}"`)
}

function buildSheetExtras(wb: Workbook, sheetId: string, sheetIdx: number, styles: StyleTable): SheetExtrasXml {
  const meta = wb.meta(sheetId)
  const out: SheetExtrasXml = { pane: '', sheetPrInner: '', autoFilter: '', afterMerge: '', tail: '', rels: [], files: [], overrides: [], needsVml: false }
  let rid = 1
  const ref = (rg: Range) => rangeToA1(rg)
  // freeze
  if (meta.freeze && (meta.freeze.rows || meta.freeze.cols)) {
    const { rows, cols } = meta.freeze
    const pane = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight'
    out.pane = `<pane${cols ? ` xSplit="${cols}"` : ''}${rows ? ` ySplit="${rows}"` : ''} topLeftCell="${colToLetters(cols)}${rows + 1}" activePane="${pane}" state="frozen"/><selection pane="${pane}"/>`
  }
  // tables
  const tableRanges: Range[] = []
  const tableParts: string[] = []
  for (const t of meta.tables ?? []) {
    const id = ++tableCounter
    tableRanges.push(t.range)
    const cols: string[] = []
    const fnMap: Record<string, string> = { sum: 'sum', average: 'average', count: 'count', countNumbers: 'countNums', min: 'min', max: 'max' }
    const used = new Set<string>()
    for (let c = t.range.c1; c <= t.range.c2; c++) {
      let name = wb.getDisplay(sheetId, t.range.r1, c).text.trim() || `Column${c - t.range.c1 + 1}`
      while (used.has(name.toLowerCase())) name += '_'
      used.add(name.toLowerCase())
      const fn = t.totalRow ? t.totals[c - t.range.c1] : undefined
      const label = t.totalRow && c === t.range.c1 && (!fn || fn === 'none') ? ' totalsRowLabel="Total"' : ''
      cols.push(`<tableColumn id="${c - t.range.c1 + 1}" name="${escapeXml(name)}"${fn && fn !== 'none' ? ` totalsRowFunction="${fnMap[fn]}"` : ''}${label}/>`)
    }
    const dataRef = ref({ ...t.range, r2: t.totalRow ? t.range.r2 - 1 : t.range.r2 })
    const styleName = { teal: 'TableStyleMedium7', blue: 'TableStyleMedium2', orange: 'TableStyleMedium3', gray: 'TableStyleMedium1', green: 'TableStyleMedium6', purple: 'TableStyleMedium5' }[t.style]
    out.files.push([
      `xl/tables/table${id}.xml`,
      xmlHeader +
        `<table xmlns="${NS_MAIN}" id="${id}" name="${escapeXml(t.name)}" displayName="${escapeXml(t.name)}" ref="${ref(t.range)}"${t.totalRow ? ' totalsRowCount="1"' : ' totalsRowShown="0"'}>` +
        (t.filter ? `<autoFilter ref="${dataRef}"/>` : '') +
        `<tableColumns count="${cols.length}">${cols.join('')}</tableColumns>` +
        `<tableStyleInfo name="${styleName}" showFirstColumn="0" showLastColumn="0" showRowStripes="${t.banded ? 1 : 0}" showColumnStripes="0"/></table>`,
    ])
    out.overrides.push(`<Override PartName="/xl/tables/table${id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`)
    const r = `rId${rid++}`
    out.rels.push(`<Relationship Id="${r}" Type="${NS_R}/table" Target="../tables/table${id}.xml"/>`)
    tableParts.push(`<tablePart r:id="${r}"/>`)
  }
  // autofilter (skipped when it coincides with a table, which carries its own)
  const af = meta.autoFilter
  if (af && !tableRanges.some((t) => rangesIntersect(t, af.range))) {
    const colsXml: string[] = []
    for (const [k, crit] of Object.entries(af.criteria)) {
      const colId = +k - af.range.c1
      if (crit.kind === 'values') {
        colsXml.push(`<filterColumn colId="${colId}"><filters${crit.blanks ? ' blank="1"' : ''}>${crit.values.map((v) => `<filter val="${escapeXml(v)}"/>`).join('')}</filters></filterColumn>`)
      } else {
        const opx: Record<string, [string, (a: string) => string]> = {
          contains: ['equal', (a) => `*${a}*`], notContains: ['notEqual', (a) => `*${a}*`], equals: ['equal', (a) => a], notEquals: ['notEqual', (a) => a],
          begins: ['equal', (a) => `${a}*`], ends: ['equal', (a) => `*${a}`], gt: ['greaterThan', (a) => a], gte: ['greaterThanOrEqual', (a) => a], lt: ['lessThan', (a) => a], lte: ['lessThanOrEqual', (a) => a],
        }
        if (crit.op === 'top' || crit.op === 'bottom') colsXml.push(`<filterColumn colId="${colId}"><top10${crit.op === 'bottom' ? ' top="0"' : ''} val="${+crit.a || 10}"/></filterColumn>`)
        else if (crit.op === 'aboveAverage' || crit.op === 'belowAverage') colsXml.push(`<filterColumn colId="${colId}"><dynamicFilter type="${crit.op}"/></filterColumn>`)
        else if (crit.op === 'between') colsXml.push(`<filterColumn colId="${colId}"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="${escapeXml(crit.a)}"/><customFilter operator="lessThanOrEqual" val="${escapeXml(crit.b ?? '')}"/></customFilters></filterColumn>`)
        else if (opx[crit.op]) colsXml.push(`<filterColumn colId="${colId}"><customFilters><customFilter operator="${opx[crit.op][0]}" val="${escapeXml(opx[crit.op][1](crit.a))}"/></customFilters></filterColumn>`)
      }
    }
    out.autoFilter = `<autoFilter ref="${ref(af.range)}">${colsXml.join('')}</autoFilter>`
  }
  // conditional formatting
  let prio = 1
  const cfXml: string[] = []
  for (const cf of meta.condFormats ?? []) {
    const tl = `${colToLetters(cf.range.c1)}${cf.range.r1 + 1}`
    const r = cf.rule
    const dxf = cf.style && Object.keys(cf.style).length ? ` dxfId="${styles.dxf(cf.style)}"` : ''
    const stop = cf.stopIfTrue ? ' stopIfTrue="1"' : ''
    const esc = (t: string) => escapeXml(t.replace(/"/g, '""'))
    let rule = ''
    switch (r.type) {
      case 'cellIs':
        rule = `<cfRule type="cellIs"${dxf} priority="${prio}"${stop} operator="${OP_TO_XML[r.op]}"><formula>${numOrFormula(r.v1)}</formula>${r.v2 !== undefined && (r.op === 'between' || r.op === 'notBetween') ? `<formula>${numOrFormula(r.v2)}</formula>` : ''}</cfRule>`
        break
      case 'text': {
        const m = { contains: ['containsText', 'containsText', `NOT(ISERROR(SEARCH("${esc(r.text)}",${tl})))`], notContains: ['notContainsText', 'notContains', `ISERROR(SEARCH("${esc(r.text)}",${tl}))`], begins: ['beginsWith', 'beginsWith', `LEFT(${tl},LEN("${esc(r.text)}"))="${esc(r.text)}"`], ends: ['endsWith', 'endsWith', `RIGHT(${tl},LEN("${esc(r.text)}"))="${esc(r.text)}"`] }[r.op]
        rule = `<cfRule type="${m[0]}"${dxf} priority="${prio}"${stop} operator="${m[1]}" text="${escapeXml(r.text)}"><formula>${m[2]}</formula></cfRule>`
        break
      }
      case 'duplicate':
        rule = `<cfRule type="duplicateValues"${dxf} priority="${prio}"${stop}/>`
        break
      case 'unique':
        rule = `<cfRule type="uniqueValues"${dxf} priority="${prio}"${stop}/>`
        break
      case 'blank':
        rule = `<cfRule type="containsBlanks"${dxf} priority="${prio}"${stop}><formula>LEN(TRIM(${tl}))=0</formula></cfRule>`
        break
      case 'notBlank':
        rule = `<cfRule type="notContainsBlanks"${dxf} priority="${prio}"${stop}><formula>LEN(TRIM(${tl}))&gt;0</formula></cfRule>`
        break
      case 'errors':
        rule = `<cfRule type="containsErrors"${dxf} priority="${prio}"${stop}><formula>ISERROR(${tl})</formula></cfRule>`
        break
      case 'formula':
        rule = `<cfRule type="expression"${dxf} priority="${prio}"${stop}><formula>${escapeXml(addXlfn(toExcelFormula(r.formula.startsWith('=') ? r.formula : '=' + r.formula)).slice(1))}</formula></cfRule>`
        break
      case 'top':
        rule = `<cfRule type="top10"${dxf} priority="${prio}"${stop} rank="${r.n}"${r.percent ? ' percent="1"' : ''}${r.bottom ? ' bottom="1"' : ''}/>`
        break
      case 'average':
        rule = `<cfRule type="aboveAverage"${dxf} priority="${prio}"${stop}${r.above ? '' : ' aboveAverage="0"'}/>`
        break
      case 'colorScale':
        rule = `<cfRule type="colorScale" priority="${prio}"><colorScale><cfvo type="min"/>${r.colors.length >= 3 ? '<cfvo type="percentile" val="50"/>' : ''}<cfvo type="max"/>${r.colors.map((c) => `<color rgb="${hex8(c)}"/>`).join('')}</colorScale></cfRule>`
        break
      case 'dataBar':
        rule = `<cfRule type="dataBar" priority="${prio}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${hex8(r.color)}"/></dataBar></cfRule>`
        break
      case 'iconSet':
        rule = `<cfRule type="iconSet" priority="${prio}"><iconSet iconSet="${{ arrows: '3Arrows', traffic: '3TrafficLights1', flags: '3Flags' }[r.set]}"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>`
        break
    }
    prio++
    cfXml.push(`<conditionalFormatting sqref="${ref(cf.range)}">${rule}</conditionalFormatting>`)
  }
  // data validation
  const dvXml = (meta.validations ?? []).map((v) => {
    const attrs = [
      `type="${v.type}"`,
      v.type !== 'list' && v.type !== 'custom' ? `operator="${OP_TO_XML[v.op ?? 'between']}"` : '',
      v.allowBlank ? 'allowBlank="1"' : '',
      v.type === 'list' && !v.showDropdown ? 'showDropDown="1"' : '',
      'showInputMessage="1" showErrorMessage="1"',
      v.errorStyle !== 'stop' ? `errorStyle="${v.errorStyle}"` : '',
      v.errorTitle ? `errorTitle="${escapeXml(v.errorTitle)}"` : '',
      v.errorMessage ? `error="${escapeXml(v.errorMessage)}"` : '',
      v.inputTitle ? `promptTitle="${escapeXml(v.inputTitle)}"` : '',
      v.inputMessage ? `prompt="${escapeXml(v.inputMessage)}"` : '',
      `sqref="${ref(v.range)}"`,
    ].filter(Boolean).join(' ')
    let f1 = ''
    let f2 = ''
    if (v.type === 'list') f1 = (v.list ?? '').startsWith('=') ? numOrFormula(v.list) : escapeXml(`"${v.list ?? ''}"`)
    else if (v.type === 'custom') f1 = numOrFormula((v.list ?? v.v1 ?? '').startsWith('=') ? v.list ?? v.v1 : '=' + (v.list ?? v.v1 ?? 'TRUE()'))
    else {
      f1 = numOrFormula(v.v1)
      f2 = v.op === 'between' || v.op === 'notBetween' ? numOrFormula(v.v2) : ''
    }
    return `<dataValidation ${attrs}><formula1>${f1}</formula1>${f2 ? `<formula2>${f2}</formula2>` : ''}</dataValidation>`
  })
  // hyperlinks
  const hl: string[] = []
  for (const [k, url] of Object.entries(meta.links ?? {})) {
    const [r, c] = k.split(',').map(Number)
    const cell = `${colToLetters(c)}${r + 1}`
    if (url.startsWith('#')) hl.push(`<hyperlink ref="${cell}" location="${escapeXml(url.slice(1))}" display="${escapeXml(wb.getDisplay(sheetId, r, c).text)}"/>`)
    else {
      const id = `rId${rid++}`
      out.rels.push(`<Relationship Id="${id}" Type="${NS_R}/hyperlink" Target="${escapeXml(url)}" TargetMode="External"/>`)
      hl.push(`<hyperlink ref="${cell}" r:id="${id}"/>`)
    }
  }
  // page setup
  const ps = meta.pageSetup ?? DEFAULT_PAGE_SETUP
  const po = ps.gridlines || ps.headings ? `<printOptions${ps.gridlines ? ' gridLines="1"' : ''}${ps.headings ? ' headings="1"' : ''}/>` : ''
  const margins = `<pageMargins left="${ps.margins.left}" right="${ps.margins.right}" top="${ps.margins.top}" bottom="${ps.margins.bottom}" header="0.3" footer="0.3"/>`
  const paper = { A4: 9, Letter: 1, Legal: 5 }[ps.paper]
  const pset = `<pageSetup paperSize="${paper}" orientation="${ps.orientation}"${ps.fitToWidth ? ' fitToWidth="1" fitToHeight="0"' : ps.scale !== 100 ? ` scale="${Math.round(ps.scale)}"` : ''}/>`
  if (ps.fitToWidth) out.sheetPrInner += '<pageSetUpPr fitToPage="1"/>'
  const hfText = (t: string | undefined) => escapeXml(`&C${t ?? ''}`)
  const hfx = ps.header || ps.footer ? `<headerFooter>${ps.header ? `<oddHeader>${hfText(ps.header)}</oddHeader>` : ''}${ps.footer ? `<oddFooter>${hfText(ps.footer)}</oddFooter>` : ''}</headerFooter>` : ''
  const brk = (list: number[], max: number, tag: string) => (list.length ? `<${tag} count="${list.length}" manualBreakCount="${list.length}">${list.map((b) => `<brk id="${b}" max="${max}" man="1"/>`).join('')}</${tag}>` : '')
  if (ps.printArea) out.printArea = `${quoteSheetName(meta.name)}!$${colToLetters(ps.printArea.c1)}$${ps.printArea.r1 + 1}:$${colToLetters(ps.printArea.c2)}$${ps.printArea.r2 + 1}`
  // notes -> comments + legacy VML drawing
  let legacy = ''
  const notes = Object.entries(meta.notes ?? {})
  if (notes.length) {
    const n = ++commentCounter
    const list = notes
      .map(([k, text]) => {
        const [r, c] = k.split(',').map(Number)
        return `<comment ref="${colToLetters(c)}${r + 1}" authorId="0"><text><r><t xml:space="preserve">${escapeXml(text)}</t></r></text></comment>`
      })
      .join('')
    out.files.push([`xl/comments${n}.xml`, xmlHeader + `<comments xmlns="${NS_MAIN}"><authors><author>ExcelPro</author></authors><commentList>${list}</commentList></comments>`])
    const shapes = notes
      .map(([k], i) => {
        const [r, c] = k.split(',').map(Number)
        return `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${r}, 10, ${c + 3}, 15, ${r + 4}, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`
      })
      .join('')
    out.files.push([
      `xl/drawings/vmlDrawing${n}.vml`,
      `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${n}"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${shapes}</xml>`,
    ])
    out.overrides.push(`<Override PartName="/xl/comments${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`)
    const c1 = `rId${rid++}`
    const v1 = `rId${rid++}`
    out.rels.push(`<Relationship Id="${c1}" Type="${NS_R}/comments" Target="../comments${n}.xml"/>`)
    out.rels.push(`<Relationship Id="${v1}" Type="${NS_R}/vmlDrawing" Target="../drawings/vmlDrawing${n}.vml"/>`)
    legacy = `<legacyDrawing r:id="${v1}"/>`
    out.needsVml = true
  }
  void sheetIdx
  out.afterMerge =
    cfXml.join('') +
    (dvXml.length ? `<dataValidations count="${dvXml.length}">${dvXml.join('')}</dataValidations>` : '') +
    (hl.length ? `<hyperlinks>${hl.join('')}</hyperlinks>` : '') +
    po +
    margins +
    pset +
    hfx +
    brk(ps.rowBreaks, 16383, 'rowBreaks') +
    brk(ps.colBreaks, 1048575, 'colBreaks')
  out.tail = legacy + (tableParts.length ? `<tableParts count="${tableParts.length}">${tableParts.join('')}</tableParts>` : '')
  return out
}
