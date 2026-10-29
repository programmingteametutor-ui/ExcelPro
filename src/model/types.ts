// Core workbook data types. Pure data — no UI, engine or persistence code here.

export interface CellAddr {
  r: number // 0-based row
  c: number // 0-based column
}

export interface Range {
  r1: number
  c1: number
  r2: number
  c2: number
}

export type BorderStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double'

export interface BorderSide {
  style: BorderStyle
  color: string
}

export interface CellBorders {
  top?: BorderSide
  right?: BorderSide
  bottom?: BorderSide
  left?: BorderSide
}

export type HAlign = 'left' | 'center' | 'right'
export type VAlign = 'top' | 'middle' | 'bottom'

export interface CellStyle {
  fontFamily?: string
  fontSize?: number // points
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  color?: string // #rrggbb
  fill?: string // #rrggbb
  hAlign?: HAlign
  vAlign?: VAlign
  wrap?: boolean
  indent?: number
  numFmt?: string // Excel-style number format code
  borders?: CellBorders
}

/** Formatting + layout for one sheet. Cell *contents* live in the formula engine. */
export interface SheetMeta {
  id: string
  name: string
  tabColor?: string
  hidden?: boolean
  showGridLines?: boolean
  colWidths: Record<number, number> // px
  rowHeights: Record<number, number> // px
  hiddenRows: number[]
  hiddenCols: number[]
  styles: Record<string, CellStyle> // key "r,c"
  merges: Range[]
  // ---- Phase 2
  autoFilter?: AutoFilter
  /** rows hidden by the filter (kept apart from rows the user hid) */
  filterHidden?: number[]
  validations?: Validation[]
  condFormats?: CondFormat[]
  freeze?: { rows: number; cols: number }
  notes?: Record<string, string> // key "r,c"
  links?: Record<string, string> // key "r,c" -> URL
  // ---- Phase 3
  tables?: TableDef[]
  pivots?: PivotDef[]
  charts?: ChartDef[]
  pageSetup?: PageSetup
}

/** Persisted form of one sheet: its meta plus raw cell contents. */
export interface SheetDoc extends SheetMeta {
  /** Raw cell contents, row-major, formulas as "=..." strings, trailing empties trimmed. */
  cells: (string | number | boolean | null)[][]
}

// ------------------------------------------------------------------ Phase 2 types
export type FilterOp = 'contains' | 'notContains' | 'equals' | 'notEquals' | 'begins' | 'ends' | 'gt' | 'gte' | 'lt' | 'lte' | 'between' | 'top' | 'bottom' | 'aboveAverage' | 'belowAverage'

export type FilterCriterion =
  | { kind: 'values'; values: string[]; blanks: boolean } // displayed texts to KEEP
  | { kind: 'condition'; op: FilterOp; a: string; b?: string }

export interface AutoFilter {
  range: Range // includes the header row
  criteria: Record<number, FilterCriterion> // absolute column index -> criterion
}

export type CompareOp = 'between' | 'notBetween' | 'eq' | 'ne' | 'gt' | 'lt' | 'gte' | 'lte'

export interface Validation {
  id: string
  range: Range
  type: 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom'
  op?: CompareOp
  v1?: string
  v2?: string
  /** list items separated by commas, or a reference/formula such as "=$F$1:$F$5" */
  list?: string
  allowBlank: boolean
  showDropdown: boolean
  inputTitle?: string
  inputMessage?: string
  errorStyle: 'stop' | 'warning' | 'information'
  errorTitle?: string
  errorMessage?: string
}

export interface CFStyle {
  fill?: string
  color?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
}

export type CFRule =
  | { type: 'cellIs'; op: CompareOp; v1: string; v2?: string }
  | { type: 'text'; op: 'contains' | 'notContains' | 'begins' | 'ends'; text: string }
  | { type: 'duplicate' }
  | { type: 'unique' }
  | { type: 'blank' }
  | { type: 'notBlank' }
  | { type: 'errors' }
  | { type: 'formula'; formula: string } // relative to the range's top-left cell
  | { type: 'top'; n: number; percent?: boolean; bottom?: boolean }
  | { type: 'average'; above: boolean }
  | { type: 'colorScale'; colors: string[] } // 2 or 3 colors: min, (mid), max
  | { type: 'dataBar'; color: string }
  | { type: 'iconSet'; set: 'arrows' | 'traffic' | 'flags' }

export interface CondFormat {
  id: string
  range: Range
  rule: CFRule
  style?: CFStyle
  stopIfTrue?: boolean
}

// ------------------------------------------------------------------ Phase 3 types
export type TotalFn = 'none' | 'sum' | 'average' | 'count' | 'countNumbers' | 'min' | 'max'

export interface TableDef {
  id: string
  name: string
  range: Range // header row + data (+ total row if totalRow)
  style: 'teal' | 'blue' | 'orange' | 'gray' | 'green' | 'purple'
  banded: boolean
  totalRow: boolean
  totals: Record<number, TotalFn> // column offset within the table -> function
  filter: boolean
}

export type AggFn = 'sum' | 'count' | 'average' | 'min' | 'max' | 'countDistinct'

export interface PivotDef {
  id: string
  name: string
  source: { sheetId: string; range: Range } // includes header row
  target: { r: number; c: number } // on the sheet that owns the pivot
  rows: string[]
  cols: string[]
  values: { field: string; agg: AggFn }[]
  filters: { field: string; values: string[] }[] // values to KEEP
  dateGroup?: Record<string, 'year' | 'quarter' | 'month'>
  sort: 'asc' | 'desc'
  grandTotals: boolean
  autoRefresh: boolean
  lastOutput?: Range
}

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'histogram'

export interface ChartDef {
  id: string
  type: ChartType
  source: { sheetId: string; range: Range }
  seriesIn: 'cols' | 'rows'
  firstRowHeader: boolean
  firstColLabels: boolean
  title: string
  xTitle: string
  yTitle: string
  legend: 'top' | 'bottom' | 'right' | 'none'
  dataLabels: boolean
  gridlines: boolean
  stacked: boolean
  colors: string[]
  bins?: number
  hiddenSeries?: number[]
  // placement in sheet pixels (unzoomed), anchored to the sheet that owns the chart
  x: number
  y: number
  w: number
  h: number
}

export interface PageSetup {
  orientation: 'portrait' | 'landscape'
  paper: 'A4' | 'Letter' | 'Legal'
  margins: { top: number; bottom: number; left: number; right: number } // inches
  scale: number // percent
  fitToWidth: boolean
  printArea?: Range
  gridlines: boolean
  headings: boolean
  header?: string
  footer?: string // supports &P (page), &N (pages), &D (date), &F (file), &A (sheet)
  rowBreaks: number[] // manual page break BEFORE these rows
  colBreaks: number[]
}

export const DEFAULT_PAGE_SETUP: PageSetup = {
  orientation: 'portrait',
  paper: 'A4',
  margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 },
  scale: 100,
  fitToWidth: false,
  gridlines: false,
  headings: false,
  footer: 'Page &P of &N',
  rowBreaks: [],
  colBreaks: [],
}

export interface NamedRangeDoc {
  name: string
  expression: string // e.g. "=Sheet1!$A$1:$B$10"
}

export interface WorkbookDoc {
  schemaVersion: 1
  id: string
  name: string
  activeSheetId: string
  sheetOrder: string[]
  sheets: SheetDoc[]
  namedRanges: NamedRangeDoc[]
  createdAt: string
  updatedAt: string
}

export const DEFAULT_COL_WIDTH = 88
export const DEFAULT_ROW_HEIGHT = 22
export const MAX_ROWS = 1_048_576
export const MAX_COLS = 16_384
export const DEFAULT_FONT = 'Calibri, Carlito, "Segoe UI", Arial, sans-serif'
export const DEFAULT_FONT_SIZE = 11
