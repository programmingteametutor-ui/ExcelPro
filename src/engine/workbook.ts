/**
 * Workbook: the spreadsheet model.
 *
 * - Cell contents, formulas, dependency tracking and recalculation live in HyperFormula.
 * - Formatting and layout (styles, merges, sizes, tab colors…) live in SheetMeta objects.
 * - Every mutation is a Command on one undo stack. Commands that change cell contents run as a
 *   single HyperFormula batch (one HF undo entry), so undo = hf.undo() + restore meta snapshot.
 *
 * No React or DOM code here; the UI subscribes to change events.
 */
import { DetailedCellError, HyperFormula } from 'hyperformula'
import type { CellValue, ConfigParams, RawCellContent, SimpleCellAddress } from 'hyperformula'
import { cellKey, colToLetters, inRange, parseKey, quoteSheetName, rangesIntersect } from '../model/address'
import type { CellAddr, CellStyle, NamedRangeDoc, Range, SheetDoc, SheetMeta, WorkbookDoc } from '../model/types'
import { MAX_COLS, MAX_ROWS } from '../model/types'
import { functionNames, shiftFormula, toExcelFormula } from './formulaText'
import { shiftIndex, shiftKeyed, shiftRange } from './rangeOps'
import { resolveStructuredRefs } from './features/tables'
import { registerExcelProFunctions } from './functions'
import { parseInput, wouldConvert, type DateOrder, type RawContent } from './inputParser'
import { formatGeneral, formatNumber, formatText, isDateFormat } from './numberFormat'

registerExcelProFunctions()

export const UNDO_LIMIT = 200

export const HF_CONFIG: Partial<ConfigParams> = {
  licenseKey: 'gpl-v3',
  undoLimit: UNDO_LIMIT,
  maxRows: MAX_ROWS,
  maxColumns: MAX_COLS,
  evaluateNullToZero: true,
  leapYear1900: true,
  nullDate: { year: 1899, month: 12, day: 31 },
  dateFormats: ['YYYY-MM-DD', 'YYYY/MM/DD', 'MM/DD/YYYY', 'MM/DD/YY'],
  timeFormats: ['hh:mm', 'hh:mm:ss.sss'],
  currencySymbol: ['$', '₹', '€', '£'],
  useColumnIndex: false,
  smartRounding: true,
}

export type CellKind = 'empty' | 'number' | 'text' | 'bool' | 'error'

export interface DisplayCell {
  text: string
  kind: CellKind
  color?: string
  error?: { code: string; message: string }
  isFormula: boolean
}

export interface ClipboardData {
  sourceSheetId: string
  range: Range
  raw: RawContent[][]
  styles: (CellStyle | undefined)[][]
  cut: boolean
  text: string
}

export type PasteMode = 'all' | 'values' | 'formats' | 'formulas'

interface MetaSnapshot {
  order: string[]
  activeSheetId: string
  metas: Record<string, SheetMeta>
}

interface StyleDiff {
  sheetId: string
  keys: string[]
  before: (CellStyle | undefined)[]
  after: (CellStyle | undefined)[]
}

interface Command {
  label: string
  hf: boolean
  meta?: { before: MetaSnapshot; after: MetaSnapshot }
  styles?: StyleDiff
  dirty: string[] | 'all'
  selection?: { sheetId: string; range: Range }
}

export interface ChangeEvent {
  dirty: string[] | 'all'
  reason: 'edit' | 'undo' | 'redo' | 'meta' | 'active' | 'load'
}

let idCounter = 0
export const newId = (prefix = 's') =>
  `${prefix}_${Date.now().toString(36)}${(idCounter++).toString(36)}${Math.random().toString(36).slice(2, 6)}`

const PHASE_KEYS = ['autoFilter', 'filterHidden', 'validations', 'condFormats', 'freeze', 'notes', 'links', 'tables', 'pivots', 'charts', 'pageSetup'] as const

export const cloneMeta = (m: SheetMeta): SheetMeta => {
  const out: SheetMeta = {
    ...m,
    colWidths: { ...m.colWidths },
    rowHeights: { ...m.rowHeights },
    hiddenRows: [...m.hiddenRows],
    hiddenCols: [...m.hiddenCols],
    styles: { ...m.styles },
    merges: m.merges.map((r) => ({ ...r })),
  }
  for (const k of PHASE_KEYS) if (m[k] !== undefined) (out as unknown as Record<string, unknown>)[k] = structuredClone(m[k])
  return out
}

export function emptyMeta(id: string, name: string): SheetMeta {
  return { id, name, colWidths: {}, rowHeights: {}, hiddenRows: [], hiddenCols: [], styles: {}, merges: [] }
}

export function validateSheetName(name: string, existing: string[]): string | null {
  const n = name.trim()
  if (!n) return 'Sheet names cannot be blank.'
  if (n.length > 31) return 'Sheet names can be at most 31 characters.'
  if (/[\\/?*[\]:]/.test(n)) return 'Sheet names cannot contain \\ / ? * [ ] :'
  if (n.startsWith("'") || n.endsWith("'")) return "Sheet names cannot start or end with an apostrophe."
  if (existing.some((e) => e.toLowerCase() === n.toLowerCase())) return `A sheet named "${n}" already exists.`
  return null
}

function styleEmpty(s: CellStyle | undefined): boolean {
  if (!s) return true
  return Object.values(s).every((v) => v === undefined || v === false)
}

function cleanStyle(s: CellStyle): CellStyle | undefined {
  const out: CellStyle = {}
  for (const [k, v] of Object.entries(s) as [keyof CellStyle, unknown][]) {
    if (v === undefined || v === false || v === null) continue
    if (k === 'borders') {
      const b = v as NonNullable<CellStyle['borders']>
      const bb = Object.fromEntries(Object.entries(b).filter(([, x]) => x))
      if (Object.keys(bb).length) out.borders = bb
      continue
    }
    ;(out as Record<string, unknown>)[k] = v
  }
  return Object.keys(out).length ? out : undefined
}

export class Workbook {
  hf: HyperFormula
  id: string
  name: string
  createdAt: string
  activeSheetId: string
  order: string[]
  dateOrder: DateOrder = 'MDY'
  /** Bumped on every change; UI uses it to re-render. */
  version = 0

  private metas = new Map<string, SheetMeta>()
  private undoStack: Command[] = []
  private redoStack: Command[] = []
  private listeners = new Set<(e: ChangeEvent) => void>()
  private dirtySheets = new Set<string>()
  private dirtyAll = true

  private constructor(hf: HyperFormula, metas: SheetMeta[], opts: { id: string; name: string; createdAt: string; activeSheetId?: string }) {
    this.hf = hf
    this.id = opts.id
    this.name = opts.name
    this.createdAt = opts.createdAt
    for (const m of metas) this.metas.set(m.id, m)
    this.order = metas.map((m) => m.id)
    this.activeSheetId = opts.activeSheetId && this.metas.has(opts.activeSheetId) ? opts.activeSheetId : this.order[0]
    hf.clearUndoStack()
    hf.clearRedoStack()
  }

  // ------------------------------------------------------------------ construction
  static blank(name = 'Book1', sheetNames = ['Sheet1']): Workbook {
    const metas = sheetNames.map((n) => emptyMeta(newId(), n))
    const hf = HyperFormula.buildFromSheets(Object.fromEntries(sheetNames.map((n) => [n, []])), HF_CONFIG)
    return new Workbook(hf, metas, { id: newId('wb'), name, createdAt: new Date().toISOString() })
  }

  /** Builds a workbook from persisted/imported data. */
  static fromDoc(doc: WorkbookDoc): Workbook {
    const sheetsById = new Map(doc.sheets.map((s) => [s.id, s]))
    const ordered = doc.sheetOrder.map((id) => sheetsById.get(id)).filter((s): s is SheetDoc => !!s)
    for (const s of doc.sheets) if (!doc.sheetOrder.includes(s.id)) ordered.push(s)
    if (ordered.length === 0) ordered.push({ ...emptyMeta(newId(), 'Sheet1'), cells: [] })
    const sheets: Record<string, RawCellContent[][]> = {}
    for (const s of ordered) sheets[s.name] = s.cells as RawCellContent[][]
    const named = (doc.namedRanges ?? []).map((n) => ({ name: n.name, expression: n.expression }))
    let hf: HyperFormula
    try {
      hf = HyperFormula.buildFromSheets(sheets, HF_CONFIG, named)
    } catch {
      // A broken named range must not make the whole workbook unopenable.
      hf = HyperFormula.buildFromSheets(sheets, HF_CONFIG)
    }
    const metas = ordered.map((s) => {
      const { cells: _cells, ...meta } = s
      void _cells
      return { ...emptyMeta(s.id, s.name), ...meta } as SheetMeta
    })
    return new Workbook(hf, metas, { id: doc.id, name: doc.name, createdAt: doc.createdAt, activeSheetId: doc.activeSheetId })
  }

  toDoc(): WorkbookDoc {
    return {
      schemaVersion: 1,
      id: this.id,
      name: this.name,
      activeSheetId: this.activeSheetId,
      sheetOrder: [...this.order],
      sheets: this.order.map((id) => this.serializeSheet(id)),
      namedRanges: this.namedRanges(),
      createdAt: this.createdAt,
      updatedAt: new Date().toISOString(),
    }
  }

  serializeSheet(id: string): SheetDoc {
    const meta = this.meta(id)
    const hfId = this.hfId(id)
    const cells = this.hf.getSheetSerialized(hfId) as RawContent[][]
    // Spilled array results are outputs, not inputs: don't persist them as constants.
    for (let r = 0; r < cells.length; r++) {
      const row = cells[r]
      for (let c = 0; c < row.length; c++) {
        const v = row[c]
        if (v === null || v === undefined || (typeof v === 'string' && v.startsWith('='))) continue
        const addr = { sheet: hfId, row: r, col: c }
        if (this.hf.isCellPartOfArray(addr) && !this.hf.doesCellHaveFormula(addr)) row[c] = null
      }
      while (row.length && (row[row.length - 1] === null || row[row.length - 1] === undefined)) row.pop()
    }
    while (cells.length && cells[cells.length - 1].length === 0) cells.pop()
    const m = cloneMeta(meta)
    return { ...m, cells }
  }

  namedRanges(): NamedRangeDoc[] {
    return this.hf
      .getAllNamedExpressionsSerialized()
      .filter((n) => n.scope === undefined)
      .map((n) => ({ name: n.name, expression: String(n.expression) }))
  }

  // ------------------------------------------------------------------ events & persistence hooks
  subscribe(fn: (e: ChangeEvent) => void): () => void {
    this.listeners.add(fn)
    return () => this.listeners.delete(fn)
  }

  private emit(e: ChangeEvent) {
    this.version++
    if (e.reason !== 'active') {
      if (e.dirty === 'all') this.dirtyAll = true
      else for (const id of e.dirty) this.dirtySheets.add(id)
    }
    for (const l of this.listeners) l(e)
  }

  /** Returns and clears the set of sheets that changed since the last call. */
  takeDirty(): { all: boolean; sheets: string[] } {
    const out = { all: this.dirtyAll, sheets: [...this.dirtySheets] }
    this.dirtyAll = false
    this.dirtySheets.clear()
    return out
  }

  markAllDirty() {
    this.dirtyAll = true
  }

  setName(name: string) {
    this.name = name.trim() || this.name
    this.emit({ dirty: 'all', reason: 'meta' })
  }

  // ------------------------------------------------------------------ sheet queries
  get sheets(): SheetMeta[] {
    return this.order.map((id) => this.metas.get(id)!)
  }

  meta(id: string): SheetMeta {
    const m = this.metas.get(id)
    if (!m) throw new Error(`Unknown sheet ${id}`)
    return m
  }

  sheetByName(name: string): SheetMeta | undefined {
    const lower = name.toLowerCase()
    return this.sheets.find((s) => s.name.toLowerCase() === lower)
  }

  hfId(sheetId: string): number {
    const id = this.hf.getSheetId(this.meta(sheetId).name)
    if (id === undefined) throw new Error(`Sheet ${this.meta(sheetId).name} is missing from the engine`)
    return id
  }

  private addr(sheetId: string, r: number, c: number): SimpleCellAddress {
    return { sheet: this.hfId(sheetId), row: r, col: c }
  }

  setActiveSheet(id: string) {
    if (!this.metas.has(id) || this.activeSheetId === id) return
    this.activeSheetId = id
    this.emit({ dirty: [], reason: 'active' })
  }

  usedSize(sheetId: string): { rows: number; cols: number } {
    const d = this.hf.getSheetDimensions(this.hfId(sheetId))
    const meta = this.meta(sheetId)
    let rows = d.height
    let cols = d.width
    for (const k of Object.keys(meta.styles)) {
      const { r, c } = parseKey(k)
      if (r + 1 > rows) rows = r + 1
      if (c + 1 > cols) cols = c + 1
    }
    return { rows, cols }
  }

  // ------------------------------------------------------------------ cell queries
  getRaw(sheetId: string, r: number, c: number): RawContent {
    const v = this.hf.getCellSerialized(this.addr(sheetId, r, c)) as RawContent | undefined
    return v === undefined ? null : v
  }

  isFormula(sheetId: string, r: number, c: number): boolean {
    return this.hf.doesCellHaveFormula(this.addr(sheetId, r, c))
  }

  getValue(sheetId: string, r: number, c: number): CellValue {
    return this.hf.getCellValue(this.addr(sheetId, r, c))
  }

  getStyle(sheetId: string, r: number, c: number): CellStyle | undefined {
    return this.meta(sheetId).styles[cellKey(r, c)]
  }

  /** The text shown in the formula bar / cell editor. */
  getEditText(sheetId: string, r: number, c: number): string {
    const raw = this.getRaw(sheetId, r, c)
    if (raw === null) return ''
    if (typeof raw === 'boolean') return raw ? 'TRUE' : 'FALSE'
    const fmt = this.getStyle(sheetId, r, c)?.numFmt
    if (typeof raw === 'number') {
      if (isDateFormat(fmt)) {
        const hasTime = raw % 1 !== 0
        const hasDate = raw >= 1
        const f = hasDate && hasTime ? 'yyyy-mm-dd h:mm:ss' : hasDate ? 'yyyy-mm-dd' : 'h:mm:ss'
        return formatNumber(raw, f).text
      }
      if (fmt && /%/.test(fmt.split(';')[0].replace(/"[^"]*"/g, ''))) return `${+(raw * 100).toPrecision(15)}%`
      return String(+raw.toPrecision(15))
    }
    if (raw.startsWith('=')) return toExcelFormula(raw)
    if (raw.startsWith("'")) return raw
    if (fmt === '@') return raw
    return wouldConvert(raw, this.dateOrder) ? `'${raw}` : raw
  }

  getDisplay(sheetId: string, r: number, c: number, showFormulas = false): DisplayCell {
    const addr = this.addr(sheetId, r, c)
    const isFormula = this.hf.doesCellHaveFormula(addr)
    if (showFormulas && isFormula) {
      return { text: toExcelFormula(String(this.hf.getCellFormula(addr) ?? '')), kind: 'text', isFormula }
    }
    const v = this.hf.getCellValue(addr)
    const fmt = this.meta(sheetId).styles[cellKey(r, c)]?.numFmt
    if (v === null || v === undefined) return { text: '', kind: 'empty', isFormula }
    if (v instanceof DetailedCellError) {
      return { text: v.value, kind: 'error', error: { code: v.value, message: v.message }, isFormula }
    }
    if (typeof v === 'number') {
      const f = formatNumber(v, fmt)
      return { text: f.text, kind: 'number', color: f.color, isFormula }
    }
    if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', kind: 'bool', isFormula }
    const f = formatText(String(v), fmt)
    return { text: f.text, kind: 'text', color: f.color, isFormula }
  }

  mergeAt(sheetId: string, r: number, c: number): Range | undefined {
    return this.meta(sheetId).merges.find((m) => inRange(m, r, c))
  }

  /** Explains an error cell in plain language, including unsupported-function hints. */
  explainError(sheetId: string, r: number, c: number): string | null {
    const d = this.getDisplay(sheetId, r, c)
    if (d.kind !== 'error' || !d.error) return null
    const formula = this.isFormula(sheetId, r, c) ? String(this.getRaw(sheetId, r, c)) : ''
    if (d.error.code === '#NAME?' && formula) {
      const known = new Set(HyperFormula.getRegisteredFunctionNames('enGB'))
      const unknown = functionNames(formula).filter((n) => !known.has(n) && n !== 'OFFSET')
      if (unknown.length) {
        return `${unknown.join(', ')} ${unknown.length > 1 ? 'are' : 'is'} not supported by ExcelPro's formula engine yet. The formula is kept unchanged so it still works if you open the file in Excel.`
      }
      return d.error.message || 'The formula uses a name or function that is not defined.'
    }
    const generic: Record<string, string> = {
      '#DIV/0!': 'Division by zero.',
      '#N/A': 'A value is not available (often a lookup that found no match).',
      '#REF!': 'The formula refers to a cell that no longer exists.',
      '#VALUE!': 'A value has the wrong type for this operation.',
      '#NUM!': 'A number is invalid or out of range for this calculation.',
      '#NULL!': 'Two ranges that were expected to intersect do not.',
      '#CYCLE!': 'Circular reference: this formula depends on its own result.',
      '#ERROR!': 'The formula could not be parsed. Check parentheses, commas and quotes.',
      '#SPILL!': 'The result needs more cells than are free next to the formula.',
      '#LIC!': 'Function not available.',
    }
    const base = generic[d.error.code] ?? 'Formula error.'
    return d.error.message && d.error.message !== base ? `${base} ${d.error.message}` : base
  }

  // ------------------------------------------------------------------ command machinery
  private snapshot(ids: string[] | 'all'): MetaSnapshot {
    const list = ids === 'all' ? [...this.metas.keys()] : ids
    return {
      order: [...this.order],
      activeSheetId: this.activeSheetId,
      metas: Object.fromEntries(list.filter((id) => this.metas.has(id)).map((id) => [id, cloneMeta(this.metas.get(id)!)])),
    }
  }

  private restore(s: MetaSnapshot) {
    for (const [id, m] of Object.entries(s.metas)) this.metas.set(id, cloneMeta(m))
    this.order = [...s.order]
    for (const id of [...this.metas.keys()]) if (!this.order.includes(id) && !s.metas[id]) this.metas.delete(id)
    this.activeSheetId = this.metas.has(s.activeSheetId) ? s.activeSheetId : this.order[0]
  }

  private pushCommand(cmd: Command) {
    this.undoStack.push(cmd)
    if (this.undoStack.length > UNDO_LIMIT) this.undoStack.shift()
    this.redoStack = []
    this.emit({ dirty: cmd.dirty, reason: 'edit' })
  }

  /**
   * Runs one undoable command.
   * @param hfOps operations against HyperFormula (run inside one batch). Must not read evaluated data.
   * @param metaOps mutations of SheetMeta objects (run after hfOps).
   */
  /** @internal used by feature modules */
  run(opts: {
    label: string
    dirty: string[] | 'all'
    metaScope?: string[] | 'all'
    hfOps?: () => void
    metaOps?: () => void
    selection?: Command['selection']
  }): boolean {
    const before = opts.metaScope ? this.snapshot(opts.metaScope) : undefined
    let hfRan = false
    if (opts.hfOps) {
      this.hf.batch(opts.hfOps)
      hfRan = true
    }
    if (opts.metaOps) opts.metaOps()
    const after = opts.metaScope ? this.snapshot(opts.metaScope === 'all' ? 'all' : opts.metaScope) : undefined
    this.pushCommand({
      label: opts.label,
      hf: hfRan,
      meta: before && after ? { before, after } : undefined,
      dirty: opts.dirty,
      selection: opts.selection,
    })
    return true
  }

  canUndo() {
    return this.undoStack.length > 0
  }
  canRedo() {
    return this.redoStack.length > 0
  }
  undoLabel() {
    return this.undoStack[this.undoStack.length - 1]?.label
  }
  redoLabel() {
    return this.redoStack[this.redoStack.length - 1]?.label
  }

  undo(): Command['selection'] | null {
    const cmd = this.undoStack.pop()
    if (!cmd) return null
    if (cmd.hf) this.hf.undo()
    if (cmd.meta) this.restore(cmd.meta.before)
    if (cmd.styles) this.applyStyleValues(cmd.styles.sheetId, cmd.styles.keys, cmd.styles.before)
    this.redoStack.push(cmd)
    this.emit({ dirty: cmd.dirty, reason: 'undo' })
    return cmd.selection ?? null
  }

  redo(): Command['selection'] | null {
    const cmd = this.redoStack.pop()
    if (!cmd) return null
    if (cmd.hf) this.hf.redo()
    if (cmd.meta) this.restore(cmd.meta.after)
    if (cmd.styles) this.applyStyleValues(cmd.styles.sheetId, cmd.styles.keys, cmd.styles.after)
    this.undoStack.push(cmd)
    this.emit({ dirty: cmd.dirty, reason: 'redo' })
    return cmd.selection ?? null
  }

  private applyStyleValues(sheetId: string, keys: string[], values: (CellStyle | undefined)[]) {
    const styles = this.meta(sheetId).styles
    keys.forEach((k, i) => {
      const v = values[i]
      if (v) styles[k] = v
      else delete styles[k]
    })
  }

  /** Records a style-only change as a diff (cheap even for big sheets). */
  /** @internal used by feature modules */
  runStyles(label: string, sheetId: string, changes: Map<string, CellStyle | undefined>, hfOps?: () => void, selection?: Command['selection']) {
    const styles = this.meta(sheetId).styles
    const keys = [...changes.keys()]
    const before = keys.map((k) => styles[k])
    const after = keys.map((k) => changes.get(k))
    if (hfOps) this.hf.batch(hfOps)
    this.applyStyleValues(sheetId, keys, after)
    this.undoStack.push({ label, hf: !!hfOps, styles: { sheetId, keys, before, after }, dirty: [sheetId], selection })
    if (this.undoStack.length > UNDO_LIMIT) this.undoStack.shift()
    this.redoStack = []
    this.emit({ dirty: [sheetId], reason: 'edit' })
  }

  // ------------------------------------------------------------------ range helpers
  /** Clamp whole-row/column selections to the used area so formatting a column stays cheap. */
  clampToUsed(sheetId: string, rg: Range): Range {
    const used = this.usedSize(sheetId)
    const out = { ...rg }
    const maxR = Math.max(used.rows + 100, 1000)
    const maxC = Math.max(used.cols + 26, 52)
    if (out.r2 - out.r1 > 100_000 || out.r2 >= MAX_ROWS - 1) out.r2 = Math.min(out.r2, Math.max(out.r1, maxR))
    if (out.c2 - out.c1 > 1_000 || out.c2 >= MAX_COLS - 1) out.c2 = Math.min(out.c2, Math.max(out.c1, maxC))
    return out
  }

  private forEachCell(ranges: Range[], fn: (r: number, c: number) => void) {
    const seen = new Set<string>()
    for (const rg of ranges) {
      for (let r = rg.r1; r <= rg.r2; r++) {
        for (let c = rg.c1; c <= rg.c2; c++) {
          if (ranges.length > 1) {
            const k = cellKey(r, c)
            if (seen.has(k)) continue
            seen.add(k)
          }
          fn(r, c)
        }
      }
    }
  }

  // ------------------------------------------------------------------ content mutations
  /** Typed input for one or more cells (formula bar, in-cell editor, Ctrl+Enter fill). */
  setInputs(sheetId: string, entries: { r: number; c: number; input: string }[], label = 'Edit cell') {
    if (!entries.length) return
    const hfId = this.hfId(sheetId)
    const styles = this.meta(sheetId).styles
    const parsed = entries.map((e) => {
      const p = parseInput(e.input, styles[cellKey(e.r, e.c)]?.numFmt, this.dateOrder)
      if (typeof p.raw === 'string' && p.raw.startsWith('=') && this.sheets.some((s) => s.tables?.length)) {
        p.raw = resolveStructuredRefs(this, p.raw as string, sheetId, e.r)
      }
      return { ...e, p }
    })
    const styleChanges = new Map<string, CellStyle | undefined>()
    for (const e of parsed) {
      const k = cellKey(e.r, e.c)
      const cur = styles[k]
      if (e.p.autoFormat && (!cur?.numFmt || cur.numFmt.toLowerCase() === 'general')) {
        styleChanges.set(k, { ...cur, numFmt: e.p.autoFormat })
      }
    }
    const hfOps = () => {
      for (const e of parsed) this.hf.setCellContents({ sheet: hfId, row: e.r, col: e.c }, e.p.raw as RawCellContent)
    }
    const r1 = Math.min(...entries.map((e) => e.r))
    const c1 = Math.min(...entries.map((e) => e.c))
    const selection = { sheetId, range: { r1, c1, r2: Math.max(...entries.map((e) => e.r)), c2: Math.max(...entries.map((e) => e.c)) } }
    this.runStyles(label, sheetId, styleChanges, hfOps, selection)
  }

  /** Writes raw engine content (already parsed) into a block starting at (r, c). */
  setRawBlock(sheetId: string, r: number, c: number, block: RawContent[][], label = 'Paste') {
    const hfId = this.hfId(sheetId)
    const width = Math.max(0, ...block.map((row) => row.length))
    const rect = block.map((row) => {
      const out = row.slice()
      while (out.length < width) out.push(null)
      return out
    })
    this.runStyles(label, sheetId, new Map(), () => {
      this.hf.setCellContents({ sheet: hfId, row: r, col: c }, rect as RawCellContent[][])
    }, { sheetId, range: { r1: r, c1: c, r2: r + rect.length - 1, c2: c + width - 1 } })
  }

  clearContents(sheetId: string, ranges: Range[]) {
    const hfId = this.hfId(sheetId)
    const clamped = ranges.map((rg) => this.clampToUsed(sheetId, rg))
    const used = this.hf.getSheetDimensions(hfId)
    this.runStyles('Clear contents', sheetId, new Map(), () => {
      for (const rg of clamped) {
        const r2 = Math.min(rg.r2, used.height - 1)
        const c2 = Math.min(rg.c2, used.width - 1)
        if (r2 < rg.r1 || c2 < rg.c1) continue
        const block = Array.from({ length: r2 - rg.r1 + 1 }, () => new Array(c2 - rg.c1 + 1).fill(null))
        this.hf.setCellContents({ sheet: hfId, row: rg.r1, col: rg.c1 }, block)
      }
    })
  }

  clearFormats(sheetId: string, ranges: Range[]) {
    const styles = this.meta(sheetId).styles
    const changes = new Map<string, CellStyle | undefined>()
    for (const k of Object.keys(styles)) {
      const { r, c } = parseKey(k)
      if (ranges.some((rg) => inRange(rg, r, c))) changes.set(k, undefined)
    }
    this.runStyles('Clear formats', sheetId, changes)
  }

  clearAll(sheetId: string, ranges: Range[]) {
    const styles = this.meta(sheetId).styles
    const hfId = this.hfId(sheetId)
    const used = this.hf.getSheetDimensions(hfId)
    const changes = new Map<string, CellStyle | undefined>()
    for (const k of Object.keys(styles)) {
      const { r, c } = parseKey(k)
      if (ranges.some((rg) => inRange(rg, r, c))) changes.set(k, undefined)
    }
    this.runStyles('Clear all', sheetId, changes, () => {
      for (const rg of ranges) {
        const r2 = Math.min(rg.r2, used.height - 1)
        const c2 = Math.min(rg.c2, used.width - 1)
        if (r2 < rg.r1 || c2 < rg.c1) continue
        const block = Array.from({ length: r2 - rg.r1 + 1 }, () => new Array(c2 - rg.c1 + 1).fill(null))
        this.hf.setCellContents({ sheet: hfId, row: rg.r1, col: rg.c1 }, block)
      }
    })
  }

  // ------------------------------------------------------------------ formatting
  applyStyle(sheetId: string, ranges: Range[], patch: Partial<CellStyle> | ((s: CellStyle | undefined, r: number, c: number) => CellStyle | undefined), label = 'Format cells') {
    const styles = this.meta(sheetId).styles
    const changes = new Map<string, CellStyle | undefined>()
    const clamped = ranges.map((rg) => this.clampToUsed(sheetId, rg))
    this.forEachCell(clamped, (r, c) => {
      const k = cellKey(r, c)
      const cur = styles[k]
      const next = typeof patch === 'function' ? patch(cur, r, c) : { ...cur, ...patch }
      const cleaned = next ? cleanStyle(next) : undefined
      if (styleEmpty(cleaned) && !cur) return
      changes.set(k, styleEmpty(cleaned) ? undefined : cleaned)
    })
    this.runStyles(label, sheetId, changes)
  }

  /** Border presets for the Borders control. */
  applyBorders(sheetId: string, rg: Range, preset: 'all' | 'outside' | 'none' | 'top' | 'bottom' | 'left' | 'right' | 'thickOutside', color = '#000000') {
    const thin = { style: 'thin' as const, color }
    const thick = { style: 'thick' as const, color }
    const r = this.clampToUsed(sheetId, rg)
    this.applyStyle(
      sheetId,
      [r],
      (s, row, col) => {
        const b = { ...(s?.borders ?? {}) }
        const edge = preset === 'thickOutside' ? thick : thin
        switch (preset) {
          case 'none':
            return { ...s, borders: undefined }
          case 'all':
            return { ...s, borders: { top: thin, bottom: thin, left: thin, right: thin } }
          case 'outside':
          case 'thickOutside':
            if (row === r.r1) b.top = edge
            if (row === r.r2) b.bottom = edge
            if (col === r.c1) b.left = edge
            if (col === r.c2) b.right = edge
            return { ...s, borders: b }
          case 'top':
            if (row === r.r1) b.top = thin
            return { ...s, borders: b }
          case 'bottom':
            if (row === r.r2) b.bottom = thin
            return { ...s, borders: b }
          case 'left':
            if (col === r.c1) b.left = thin
            return { ...s, borders: b }
          case 'right':
            if (col === r.c2) b.right = thin
            return { ...s, borders: b }
        }
      },
      'Borders',
    )
  }

  // ------------------------------------------------------------------ merges
  merge(sheetId: string, rg: Range, center = false): string | null {
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return 'Select more than one cell to merge.'
    if (this.usedSize(sheetId) && (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1) > 100_000) return 'That range is too large to merge.'
    const hfId = this.hfId(sheetId)
    const used = this.hf.getSheetDimensions(hfId)
    const clears: SimpleCellAddress[] = []
    for (let r = rg.r1; r <= Math.min(rg.r2, used.height - 1); r++)
      for (let c = rg.c1; c <= Math.min(rg.c2, used.width - 1); c++) {
        if (r === rg.r1 && c === rg.c1) continue
        if (this.hf.getCellSerialized({ sheet: hfId, row: r, col: c }) !== null) clears.push({ sheet: hfId, row: r, col: c })
      }
    this.run({
      label: center ? 'Merge & Center' : 'Merge cells',
      dirty: [sheetId],
      metaScope: [sheetId],
      hfOps: clears.length ? () => clears.forEach((a) => this.hf.setCellContents(a, null)) : undefined,
      metaOps: () => {
        const m = this.meta(sheetId)
        m.merges = m.merges.filter((x) => !rangesIntersect(x, rg)).concat([{ ...rg }])
        if (center) {
          const k = cellKey(rg.r1, rg.c1)
          m.styles[k] = { ...m.styles[k], hAlign: 'center' }
        }
      },
    })
    return null
  }

  unmerge(sheetId: string, rg: Range) {
    const m = this.meta(sheetId)
    if (!m.merges.some((x) => rangesIntersect(x, rg))) return
    this.run({
      label: 'Unmerge cells',
      dirty: [sheetId],
      metaScope: [sheetId],
      metaOps: () => {
        const mm = this.meta(sheetId)
        mm.merges = mm.merges.filter((x) => !rangesIntersect(x, rg))
      },
    })
  }

  // ------------------------------------------------------------------ sizes & visibility
  setColWidths(sheetId: string, cols: number[], px: number | null, label = 'Column width') {
    this.run({
      label,
      dirty: [sheetId],
      metaScope: [sheetId],
      metaOps: () => {
        const m = this.meta(sheetId)
        const next = { ...m.colWidths }
        for (const c of cols) {
          if (px === null) delete next[c]
          else next[c] = Math.max(2, Math.round(px))
        }
        m.colWidths = next
      },
    })
  }

  setRowHeights(sheetId: string, rows: number[], px: number | null, label = 'Row height') {
    this.run({
      label,
      dirty: [sheetId],
      metaScope: [sheetId],
      metaOps: () => {
        const m = this.meta(sheetId)
        const next = { ...m.rowHeights }
        for (const r of rows) {
          if (px === null) delete next[r]
          else next[r] = Math.max(2, Math.round(px))
        }
        m.rowHeights = next
      },
    })
  }

  setHidden(sheetId: string, axis: 'row' | 'col', indices: number[], hidden: boolean) {
    this.run({
      label: hidden ? `Hide ${axis === 'row' ? 'rows' : 'columns'}` : `Unhide ${axis === 'row' ? 'rows' : 'columns'}`,
      dirty: [sheetId],
      metaScope: [sheetId],
      metaOps: () => {
        const m = this.meta(sheetId)
        const key = axis === 'row' ? 'hiddenRows' : 'hiddenCols'
        const set = new Set(m[key])
        for (const i of indices) (hidden ? set.add(i) : set.delete(i))
        m[key] = [...set].sort((a, b) => a - b)
      },
    })
  }

  // ------------------------------------------------------------------ structure
  private shiftMeta(m: SheetMeta, axis: 'row' | 'col', at: number, delta: number) {
    const isRow = axis === 'row'
    const si = (i: number) => shiftIndex(i, at, delta)
    m.styles = shiftKeyed(m.styles, axis, at, delta)!
    const sizes = isRow ? m.rowHeights : m.colWidths
    const nextSizes: Record<number, number> = {}
    for (const [k, v] of Object.entries(sizes)) {
      const n = si(+k)
      if (n !== null) nextSizes[n] = v
    }
    if (isRow) m.rowHeights = nextSizes
    else m.colWidths = nextSizes
    const hiddenKey = isRow ? 'hiddenRows' : 'hiddenCols'
    m[hiddenKey] = m[hiddenKey].map(si).filter((x): x is number => x !== null)
    if (isRow && m.filterHidden) m.filterHidden = m.filterHidden.map(si).filter((x): x is number => x !== null)
    m.merges = m.merges
      .map((g) => shiftRange(g, axis, at, delta))
      .filter((g): g is Range => g !== null && !(g.r1 === g.r2 && g.c1 === g.c2))
    m.notes = shiftKeyed(m.notes, axis, at, delta)
    m.links = shiftKeyed(m.links, axis, at, delta)
    if (m.validations) m.validations = m.validations.flatMap((v) => {
      const r = shiftRange(v.range, axis, at, delta)
      return r ? [{ ...v, range: r }] : []
    })
    if (m.condFormats) m.condFormats = m.condFormats.flatMap((v) => {
      const r = shiftRange(v.range, axis, at, delta)
      return r ? [{ ...v, range: r }] : []
    })
    if (m.tables) m.tables = m.tables.flatMap((t) => {
      const r = shiftRange(t.range, axis, at, delta)
      return r && r.r2 > r.r1 ? [{ ...t, range: r }] : []
    })
    if (m.autoFilter) {
      const r = shiftRange(m.autoFilter.range, axis, at, delta)
      if (!r) m.autoFilter = undefined
      else {
        const crit: typeof m.autoFilter.criteria = {}
        for (const [k, v] of Object.entries(m.autoFilter.criteria)) {
          const n = isRow ? +k : si(+k)
          if (n !== null) crit[n] = v
        }
        m.autoFilter = { range: r, criteria: crit }
      }
    }
    if (m.pivots) m.pivots = m.pivots.map((p) => {
      const t = { ...p.target }
      const n = si(isRow ? t.r : t.c)
      if (n !== null) {
        if (isRow) t.r = n
        else t.c = n
      }
      return { ...p, target: t, lastOutput: p.lastOutput ? shiftRange(p.lastOutput, axis, at, delta) ?? undefined : undefined }
    })
    if (m.pageSetup?.printArea) {
      const r = shiftRange(m.pageSetup.printArea, axis, at, delta)
      m.pageSetup = { ...m.pageSetup, printArea: r ?? undefined }
    }
  }

  insertRows(sheetId: string, at: number, count: number) {
    const hfId = this.hfId(sheetId)
    if (!this.hf.isItPossibleToAddRows(hfId, [at, count])) return 'Rows cannot be inserted here (the sheet would exceed its size limit).'
    this.run({
      label: count > 1 ? `Insert ${count} rows` : 'Insert row',
      dirty: 'all',
      metaScope: [sheetId],
      hfOps: () => this.hf.addRows(hfId, [at, count]),
      metaOps: () => this.shiftMeta(this.meta(sheetId), 'row', at, count),
    })
    return null
  }

  deleteRows(sheetId: string, at: number, count: number) {
    const hfId = this.hfId(sheetId)
    if (!this.hf.isItPossibleToRemoveRows(hfId, [at, count])) return 'These rows cannot be deleted.'
    this.run({
      label: count > 1 ? `Delete ${count} rows` : 'Delete row',
      dirty: 'all',
      metaScope: [sheetId],
      hfOps: () => this.hf.removeRows(hfId, [at, count]),
      metaOps: () => this.shiftMeta(this.meta(sheetId), 'row', at, -count),
    })
    return null
  }

  insertCols(sheetId: string, at: number, count: number) {
    const hfId = this.hfId(sheetId)
    if (!this.hf.isItPossibleToAddColumns(hfId, [at, count])) return 'Columns cannot be inserted here (the sheet would exceed its size limit).'
    this.run({
      label: count > 1 ? `Insert ${count} columns` : 'Insert column',
      dirty: 'all',
      metaScope: [sheetId],
      hfOps: () => this.hf.addColumns(hfId, [at, count]),
      metaOps: () => this.shiftMeta(this.meta(sheetId), 'col', at, count),
    })
    return null
  }

  deleteCols(sheetId: string, at: number, count: number) {
    const hfId = this.hfId(sheetId)
    if (!this.hf.isItPossibleToRemoveColumns(hfId, [at, count])) return 'These columns cannot be deleted.'
    this.run({
      label: count > 1 ? `Delete ${count} columns` : 'Delete column',
      dirty: 'all',
      metaScope: [sheetId],
      hfOps: () => this.hf.removeColumns(hfId, [at, count]),
      metaOps: () => this.shiftMeta(this.meta(sheetId), 'col', at, -count),
    })
    return null
  }

  // ------------------------------------------------------------------ worksheets
  nextSheetName(base = 'Sheet'): string {
    const names = new Set(this.sheets.map((s) => s.name.toLowerCase()))
    let i = this.order.length + 1
    while (names.has(`${base}${i}`.toLowerCase())) i++
    return `${base}${i}`
  }

  addSheet(name?: string, index?: number): string {
    const n = name ?? this.nextSheetName()
    const err = validateSheetName(n, this.sheets.map((s) => s.name))
    if (err) throw new Error(err)
    const id = newId()
    this.run({
      label: 'Add sheet',
      dirty: 'all',
      metaScope: 'all',
      hfOps: () => {
        this.hf.addSheet(n)
      },
      metaOps: () => {
        this.metas.set(id, emptyMeta(id, n))
        const at = index ?? this.order.indexOf(this.activeSheetId) + 1
        this.order.splice(Math.max(0, Math.min(at, this.order.length)), 0, id)
        this.activeSheetId = id
      },
    })
    return id
  }

  removeSheet(id: string): string | null {
    const visible = this.sheets.filter((s) => !s.hidden)
    if (this.order.length <= 1 || (visible.length <= 1 && !this.meta(id).hidden)) return 'A workbook must contain at least one visible sheet.'
    const hfId = this.hfId(id)
    this.run({
      label: 'Delete sheet',
      dirty: 'all',
      metaScope: 'all',
      hfOps: () => this.hf.removeSheet(hfId),
      metaOps: () => {
        const idx = this.order.indexOf(id)
        this.order.splice(idx, 1)
        this.metas.delete(id)
        if (this.activeSheetId === id) {
          const vis = this.order.filter((x) => !this.metas.get(x)!.hidden)
          this.activeSheetId = vis[Math.max(0, Math.min(idx, vis.length) - 1)] ?? vis[0] ?? this.order[0]
        }
      },
    })
    return null
  }

  renameSheet(id: string, newName: string): string | null {
    const n = newName.trim()
    const m = this.meta(id)
    if (n === m.name) return null
    const err = validateSheetName(n, this.sheets.filter((s) => s.id !== id).map((s) => s.name))
    if (err) return err
    const hfId = this.hfId(id)
    this.run({
      label: 'Rename sheet',
      dirty: 'all',
      metaScope: [id],
      hfOps: () => this.hf.renameSheet(hfId, n),
      metaOps: () => {
        this.meta(id).name = n
      },
    })
    return null
  }

  duplicateSheet(id: string): string {
    const src = this.meta(id)
    let i = 2
    const names = this.sheets.map((s) => s.name.toLowerCase())
    const base = src.name.replace(/ \(\d+\)$/, '').slice(0, 25)
    while (names.includes(`${base} (${i})`.toLowerCase())) i++
    const name = `${base} (${i})`
    const content = this.hf.getSheetSerialized(this.hfId(id))
    const newSheetId = newId()
    this.run({
      label: 'Duplicate sheet',
      dirty: 'all',
      metaScope: 'all',
      hfOps: () => {
        this.hf.addSheet(name)
        const hid = this.hf.getSheetId(name)!
        if (content.length) this.hf.setSheetContent(hid, content)
      },
      metaOps: () => {
        this.metas.set(newSheetId, { ...cloneMeta(src), id: newSheetId, name, hidden: false })
        this.order.splice(this.order.indexOf(id) + 1, 0, newSheetId)
        this.activeSheetId = newSheetId
      },
    })
    return newSheetId
  }

  moveSheet(id: string, toIndex: number) {
    const from = this.order.indexOf(id)
    if (from < 0 || from === toIndex) return
    this.run({
      label: 'Move sheet',
      dirty: 'all',
      metaScope: [],
      metaOps: () => {
        this.order.splice(from, 1)
        this.order.splice(Math.max(0, Math.min(toIndex, this.order.length)), 0, id)
      },
    })
  }

  setTabColor(id: string, color: string | undefined) {
    this.run({ label: 'Tab color', dirty: [id], metaScope: [id], metaOps: () => (this.meta(id).tabColor = color) })
  }

  setSheetHidden(id: string, hidden: boolean): string | null {
    if (hidden && this.sheets.filter((s) => !s.hidden).length <= 1) return 'A workbook must contain at least one visible sheet.'
    this.run({
      label: hidden ? 'Hide sheet' : 'Unhide sheet',
      dirty: [id],
      metaScope: [id],
      metaOps: () => {
        this.meta(id).hidden = hidden
        if (hidden && this.activeSheetId === id) this.activeSheetId = this.order.find((x) => !this.metas.get(x)!.hidden)!
        if (!hidden) this.activeSheetId = id
      },
    })
    return null
  }

  setGridLines(id: string, show: boolean) {
    this.run({ label: show ? 'Show gridlines' : 'Hide gridlines', dirty: [id], metaScope: [id], metaOps: () => (this.meta(id).showGridLines = show) })
  }

  // ------------------------------------------------------------------ clipboard
  copy(sheetId: string, rg: Range, cut = false): ClipboardData {
    const raw: RawContent[][] = []
    const styles: (CellStyle | undefined)[][] = []
    const lines: string[] = []
    const meta = this.meta(sheetId)
    for (let r = rg.r1; r <= rg.r2; r++) {
      const rowRaw: RawContent[] = []
      const rowStyle: (CellStyle | undefined)[] = []
      const texts: string[] = []
      for (let c = rg.c1; c <= rg.c2; c++) {
        rowRaw.push(this.getRaw(sheetId, r, c))
        rowStyle.push(meta.styles[cellKey(r, c)])
        const t = this.getDisplay(sheetId, r, c).text
        texts.push(/[\t\n"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t)
      }
      raw.push(rowRaw)
      styles.push(rowStyle)
      lines.push(texts.join('\t'))
    }
    return { sourceSheetId: sheetId, range: { ...rg }, raw, styles, cut, text: lines.join('\n') }
  }

  /**
   * Pastes internal clipboard data. Relative references in formulas are adjusted.
   * If the target selection is a multiple of the copied size, the block is tiled (like Excel).
   */
  paste(clip: ClipboardData, sheetId: string, target: Range, mode: PasteMode = 'all'): Range {
    const h = clip.range.r2 - clip.range.r1 + 1
    const w = clip.range.c2 - clip.range.c1 + 1
    const th = target.r2 - target.r1 + 1
    const tw = target.c2 - target.c1 + 1
    const tilesR = th % h === 0 && th >= h ? th / h : 1
    const tilesC = tw % w === 0 && tw >= w ? tw / w : 1
    const dest: Range = { r1: target.r1, c1: target.c1, r2: target.r1 + h * tilesR - 1, c2: target.c1 + w * tilesC - 1 }
    if (dest.r2 >= MAX_ROWS || dest.c2 >= MAX_COLS) throw new Error('The paste area extends beyond the sheet.')

    if (clip.cut && mode === 'all' && this.metas.has(clip.sourceSheetId)) {
      return this.moveRange(clip.sourceSheetId, clip.range, sheetId, { r: target.r1, c: target.c1 })
    }
    const hfId = this.hfId(sheetId)
    const block: RawContent[][] = []
    const styleChanges = new Map<string, CellStyle | undefined>()
    for (let tr = 0; tr < tilesR; tr++)
      for (let i = 0; i < h; i++) {
        const row: RawContent[] = []
        for (let tc = 0; tc < tilesC; tc++)
          for (let j = 0; j < w; j++) {
            const r = dest.r1 + tr * h + i
            const c = dest.c1 + tc * w + j
            let v = clip.raw[i][j]
            if (mode === 'values') {
              v = this.valueOf(clip, i, j)
            } else if (typeof v === 'string' && v.startsWith('=')) {
              v = shiftFormula(v, r - (clip.range.r1 + i), c - (clip.range.c1 + j))
            }
            row.push(v)
            if (mode === 'all' || mode === 'formats') styleChanges.set(cellKey(r, c), clip.styles[i][j])
            if (mode === 'values' || mode === 'formulas') {
              // keep destination formatting, but carry number format so dates stay dates
              const nf = clip.styles[i][j]?.numFmt
              if (nf) {
                const cur = this.meta(sheetId).styles[cellKey(r, c)]
                styleChanges.set(cellKey(r, c), { ...cur, numFmt: nf })
              }
            }
          }
        block.push(row)
      }
    const ops = mode === 'formats' ? undefined : () => this.hf.setCellContents({ sheet: hfId, row: dest.r1, col: dest.c1 }, block as RawCellContent[][])
    const label = mode === 'all' ? 'Paste' : `Paste ${mode}`
    this.runStyles(label, sheetId, styleChanges, ops, { sheetId, range: dest })
    return dest
  }

  private valueOf(clip: ClipboardData, i: number, j: number): RawContent {
    const raw = clip.raw[i][j]
    if (!(typeof raw === 'string' && raw.startsWith('='))) return raw
    if (!this.metas.has(clip.sourceSheetId)) return null
    const v = this.getValue(clip.sourceSheetId, clip.range.r1 + i, clip.range.c1 + j)
    if (v instanceof DetailedCellError) return v.value
    if (typeof v === 'string') return v.length && wouldConvert(v) ? `'${v}` : v
    return v as RawContent
  }

  /** Pastes external text (TSV from Excel/Sheets/other apps). Each value is parsed like typed input. */
  pasteText(sheetId: string, at: CellAddr, text: string): Range | null {
    const rows = parseTsv(text)
    if (!rows.length) return null
    const width = Math.max(...rows.map((r) => r.length))
    const entries: { r: number; c: number; input: string }[] = []
    rows.forEach((row, i) => {
      for (let j = 0; j < width; j++) entries.push({ r: at.r + i, c: at.c + j, input: row[j] ?? '' })
    })
    if (at.r + rows.length > MAX_ROWS || at.c + width > MAX_COLS) throw new Error('The paste area extends beyond the sheet.')
    this.setInputs(sheetId, entries, 'Paste')
    return { r1: at.r, c1: at.c, r2: at.r + rows.length - 1, c2: at.c + width - 1 }
  }

  /** Cut + paste: moves contents and formats, keeping references to the moved cells pointing at them. */
  moveRange(fromSheet: string, src: Range, toSheet: string, at: CellAddr): Range {
    const h = src.r2 - src.r1 + 1
    const w = src.c2 - src.c1 + 1
    const dest: Range = { r1: at.r, c1: at.c, r2: at.r + h - 1, c2: at.c + w - 1 }
    if (fromSheet === toSheet && dest.r1 === src.r1 && dest.c1 === src.c1) return dest
    const fromHf = this.hfId(fromSheet)
    const toHf = this.hfId(toSheet)
    const ok = this.hf.isItPossibleToMoveCells(
      { start: { sheet: fromHf, row: src.r1, col: src.c1 }, end: { sheet: fromHf, row: src.r2, col: src.c2 } },
      { sheet: toHf, row: at.r, col: at.c },
    )
    if (!ok) throw new Error('These cells cannot be moved there (overlapping array formulas or sheet limits).')
    const scope = fromSheet === toSheet ? [fromSheet] : [fromSheet, toSheet]
    this.run({
      label: 'Move cells',
      dirty: 'all',
      metaScope: scope,
      selection: { sheetId: toSheet, range: dest },
      hfOps: () =>
        this.hf.moveCells(
          { start: { sheet: fromHf, row: src.r1, col: src.c1 }, end: { sheet: fromHf, row: src.r2, col: src.c2 } },
          { sheet: toHf, row: at.r, col: at.c },
        ),
      metaOps: () => {
        const fm = this.meta(fromSheet)
        const tm = this.meta(toSheet)
        const moved: [number, number, CellStyle | undefined][] = []
        for (let r = src.r1; r <= src.r2; r++)
          for (let c = src.c1; c <= src.c2; c++) {
            const k = cellKey(r, c)
            moved.push([r - src.r1, c - src.c1, fm.styles[k]])
            delete fm.styles[k]
          }
        for (let r = dest.r1; r <= dest.r2; r++) for (let c = dest.c1; c <= dest.c2; c++) delete tm.styles[cellKey(r, c)]
        for (const [i, j, s] of moved) if (s) tm.styles[cellKey(dest.r1 + i, dest.c1 + j)] = s
        // merges fully inside the source move with it
        const inside = fm.merges.filter((m) => m.r1 >= src.r1 && m.r2 <= src.r2 && m.c1 >= src.c1 && m.c2 <= src.c2)
        fm.merges = fm.merges.filter((m) => !inside.includes(m))
        tm.merges = tm.merges.filter((m) => !rangesIntersect(m, dest))
        for (const m of inside)
          tm.merges.push({ r1: m.r1 - src.r1 + dest.r1, r2: m.r2 - src.r1 + dest.r1, c1: m.c1 - src.c1 + dest.c1, c2: m.c2 - src.c1 + dest.c1 })
      },
    })
    return dest
  }

  // ------------------------------------------------------------------ fill handle
  /** Fills `dest` (which contains `src` along one edge) from the source pattern. */
  fill(sheetId: string, src: Range, dest: Range) {
    const vertical = dest.c1 === src.c1 && dest.c2 === src.c2
    const forward = vertical ? dest.r2 > src.r2 : dest.c2 > src.c2
    const styles = this.meta(sheetId).styles
    const n = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1
    const lines = vertical ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1
    const fillCount = vertical ? dest.r2 - dest.r1 + 1 - n : dest.c2 - dest.c1 + 1 - n
    if (fillCount <= 0) return
    const hfId = this.hfId(sheetId)
    const writes: { r: number; c: number; v: RawContent }[] = []
    const styleChanges = new Map<string, CellStyle | undefined>()
    for (let line = 0; line < lines; line++) {
      const srcCells: { r: number; c: number; raw: RawContent; style?: CellStyle }[] = []
      for (let k = 0; k < n; k++) {
        const r = vertical ? src.r1 + k : src.r1 + line
        const c = vertical ? src.c1 + line : src.c1 + k
        srcCells.push({ r, c, raw: this.getRaw(sheetId, r, c), style: styles[cellKey(r, c)] })
      }
      const series = detectSeries(srcCells.map((s) => s.raw), srcCells.map((s) => s.style?.numFmt))
      for (let k = 1; k <= fillCount; k++) {
        const r = vertical ? (forward ? src.r2 + k : src.r1 - k) : src.r1 + line
        const c = vertical ? src.c1 + line : forward ? src.c2 + k : src.c1 - k
        // which source cell does this position copy (tiling)?
        const idx = forward ? (k - 1) % n : n - 1 - ((k - 1) % n)
        const s = srcCells[idx]
        let v: RawContent
        if (series) v = series(forward ? n - 1 + k : -k)
        else if (typeof s.raw === 'string' && s.raw.startsWith('=')) v = shiftFormula(s.raw, r - s.r, c - s.c)
        else v = s.raw
        writes.push({ r, c, v })
        styleChanges.set(cellKey(r, c), s.style)
      }
    }
    this.runStyles('Fill', sheetId, styleChanges, () => {
      for (const w of writes) this.hf.setCellContents({ sheet: hfId, row: w.r, col: w.c }, w.v as RawCellContent)
    }, { sheetId, range: dest })
  }

  // ------------------------------------------------------------------ named ranges
  addNamedRange(name: string, expression: string): string | null {
    if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name)) return 'Invalid name.'
    if (!this.hf.isItPossibleToAddNamedExpression(name, expression)) return 'That name already exists or the reference is invalid.'
    this.run({ label: 'Define name', dirty: 'all', hfOps: () => this.hf.addNamedExpression(name, expression) })
    return null
  }

  // ------------------------------------------------------------------ navigation helpers
  isCellEmpty(sheetId: string, r: number, c: number): boolean {
    return this.hf.getCellSerialized(this.addr(sheetId, r, c)) === null && !this.hf.isCellPartOfArray(this.addr(sheetId, r, c))
  }

  /** Ctrl+Arrow target: edge of the current data region (Excel semantics). */
  dataEdge(sheetId: string, from: CellAddr, dr: number, dc: number, maxR: number, maxC: number): CellAddr {
    const used = this.hf.getSheetDimensions(this.hfId(sheetId))
    const limitR = Math.min(maxR - 1, MAX_ROWS - 1)
    const limitC = Math.min(maxC - 1, MAX_COLS - 1)
    const inb = (r: number, c: number) => r >= 0 && c >= 0 && r <= limitR && c <= limitC
    let { r, c } = from
    const empty = (rr: number, cc: number) => rr >= used.height || cc >= used.width || this.isCellEmpty(sheetId, rr, cc)
    if (!inb(r + dr, c + dc)) return { r, c }
    const curEmpty = empty(r, c)
    const nextEmpty = empty(r + dr, c + dc)
    if (!curEmpty && !nextEmpty) {
      while (inb(r + dr, c + dc) && !empty(r + dr, c + dc)) {
        r += dr
        c += dc
      }
      return { r, c }
    }
    // jump to next non-empty cell (or the sheet edge)
    r += dr
    c += dc
    while (inb(r, c) && empty(r, c)) {
      if ((dr > 0 && r >= used.height) || (dc > 0 && c >= used.width)) {
        return dr > 0 ? { r: limitR, c } : { r, c: limitC }
      }
      if (!inb(r + dr, c + dc)) break
      r += dr
      c += dc
    }
    return { r, c }
  }

  // ------------------------------------------------------------------ find & replace
  find(query: string, opts: { sheetIds: string[]; matchCase?: boolean; entireCell?: boolean; lookIn: 'values' | 'formulas' }): { sheetId: string; r: number; c: number }[] {
    if (!query) return []
    const q = opts.matchCase ? query : query.toLowerCase()
    const out: { sheetId: string; r: number; c: number }[] = []
    for (const sid of opts.sheetIds) {
      const hfId = this.hfId(sid)
      const ser = opts.lookIn === 'formulas' ? this.hf.getSheetSerialized(hfId) : null
      const dims = this.hf.getSheetDimensions(hfId)
      for (let r = 0; r < dims.height; r++)
        for (let c = 0; c < dims.width; c++) {
          let text: string
          if (ser) {
            const v = ser[r]?.[c]
            if (v === null || v === undefined) continue
            text = String(v)
          } else {
            const d = this.getDisplay(sid, r, c)
            if (d.kind === 'empty') continue
            text = d.text
          }
          const t = opts.matchCase ? text : text.toLowerCase()
          if (opts.entireCell ? t === q : t.includes(q)) out.push({ sheetId: sid, r, c })
        }
    }
    return out
  }

  replace(matches: { sheetId: string; r: number; c: number }[], query: string, replacement: string, matchCase = false): number {
    const bySheet = new Map<string, { r: number; c: number; input: string }[]>()
    const re = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi')
    for (const m of matches) {
      const raw = this.getRaw(m.sheetId, m.r, m.c)
      if (raw === null) continue
      const src = typeof raw === 'string' && raw.startsWith('=') ? raw : this.getEditText(m.sheetId, m.r, m.c)
      const next = src.replace(re, () => replacement)
      if (next === src) continue
      const list = bySheet.get(m.sheetId) ?? []
      list.push({ r: m.r, c: m.c, input: next })
      bySheet.set(m.sheetId, list)
    }
    let count = 0
    for (const [sid, entries] of bySheet) {
      this.setInputs(sid, entries, entries.length > 1 ? 'Replace all' : 'Replace')
      count += entries.length
    }
    return count
  }
}

// ------------------------------------------------------------------ helpers (exported for tests)

/** Parses tab-separated text with quoted fields (as produced by Excel/Sheets copy). */
export function parseTsv(text: string): string[][] {
  const t = text.replace(/\r\n?/g, '\n').replace(/\n$/, '')
  if (!t) return []
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let i = 0
  let quoted = false
  while (i < t.length) {
    const ch = t[i]
    if (quoted) {
      if (ch === '"') {
        if (t[i + 1] === '"') {
          field += '"'
          i += 2
          continue
        }
        quoted = false
        i++
        continue
      }
      field += ch
      i++
      continue
    }
    if (ch === '"' && field === '') {
      quoted = true
      i++
    } else if (ch === '\t') {
      row.push(field)
      field = ''
      i++
    } else if (ch === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
      i++
    } else {
      field += ch
      i++
    }
  }
  row.push(field)
  rows.push(row)
  return rows
}

const MONTH_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
const MONTH_LONG = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAY_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
const DAY_LONG = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
const LISTS = [MONTH_SHORT, MONTH_LONG, DAY_SHORT, DAY_LONG]

/**
 * Detects an Excel-style fill series in the source cells. Returns a generator mapping a position
 * (0 = first source cell, negative = before it) to a value, or null to tile/copy instead.
 */
export function detectSeries(raws: RawContent[], fmts: (string | undefined)[]): ((pos: number) => RawContent) | null {
  if (raws.some((v) => typeof v === 'string' && v.startsWith('='))) return null
  const n = raws.length
  // numbers: linear trend for 2+, dates step by day for 1
  if (raws.every((v) => typeof v === 'number')) {
    const nums = raws as number[]
    if (n === 1) {
      if (isDateFormat(fmts[0])) return (pos) => nums[0] + pos
      return null
    }
    // least-squares linear fit (Excel's AutoFill trend)
    const xs = nums.map((_, i) => i)
    const mx = (n - 1) / 2
    const my = nums.reduce((a, b) => a + b, 0) / n
    let sxy = 0
    let sxx = 0
    xs.forEach((x, i) => {
      sxy += (x - mx) * (nums[i] - my)
      sxx += (x - mx) * (x - mx)
    })
    const slope = sxx ? sxy / sxx : 0
    const icpt = my - slope * mx
    return (pos) => +(icpt + slope * pos).toPrecision(15)
  }
  if (raws.every((v) => typeof v === 'string' && v !== '')) {
    const strs = raws as string[]
    // month / weekday names
    for (const list of LISTS) {
      const idx = strs.map((s) => list.findIndex((x) => x.toLowerCase() === s.toLowerCase()))
      if (idx.every((x) => x >= 0)) {
        const step = n > 1 ? (((idx[1] - idx[0]) % list.length) + list.length) % list.length || 1 : 1
        const upper = strs[0] === strs[0].toUpperCase()
        return (pos) => {
          const v = list[(((idx[0] + step * pos) % list.length) + list.length) % list.length]
          return upper ? v.toUpperCase() : v
        }
      }
    }
    // text with trailing number: "Item 1" -> "Item 2"
    const parts = strs.map((s) => /^(.*?)(\d+)$/.exec(s))
    if (parts.every((p) => p) && parts.every((p) => p![1] === parts[0]![1])) {
      const nums = parts.map((p) => parseInt(p![2], 10))
      const width = parts[0]![2].length
      const step = n > 1 ? nums[1] - nums[0] : 1
      const prefix = parts[0]![1]
      return (pos) => {
        const v = nums[0] + step * pos
        return prefix + (v < 0 ? String(v) : String(v).padStart(parts[0]![2].startsWith('0') ? width : 1, '0'))
      }
    }
  }
  return null
}

export function formatForCsv(wb: Workbook, sheetId: string, r: number, c: number): string {
  const d = wb.getDisplay(sheetId, r, c)
  if (d.kind === 'number') {
    const fmt = wb.getStyle(sheetId, r, c)?.numFmt
    if (!fmt) return formatGeneral(wb.getValue(sheetId, r, c) as number)
  }
  return d.text
}

export { colToLetters, quoteSheetName }
