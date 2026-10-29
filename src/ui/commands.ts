import type { BorderStyle, CellStyle, HAlign, VAlign } from '../model/types'
import type { PasteMode } from '../engine/workbook'
import type { FeatureAction } from './useFeatures'

/** Everything the ribbon/menus can ask the editor to do. Implemented in Editor.tsx. */
export interface Commands {
  undo(): void
  redo(): void
  save(): void
  cut(): void
  copy(): void
  paste(mode?: PasteMode): void
  formatPainter(): void
  toggle(prop: 'bold' | 'italic' | 'underline' | 'strike' | 'wrap'): void
  setStyle(patch: Partial<CellStyle>, label?: string): void
  setAlign(h?: HAlign, v?: VAlign): void
  borders(preset: 'all' | 'outside' | 'none' | 'top' | 'bottom' | 'left' | 'right' | 'thickOutside', style?: BorderStyle): void
  merge(kind: 'center' | 'merge' | 'unmerge'): void
  numFmt(fmt: string | undefined): void
  decimals(delta: 1 | -1): void
  insert(kind: 'rowsAbove' | 'rowsBelow' | 'colsLeft' | 'colsRight' | 'sheet'): void
  remove(kind: 'rows' | 'cols' | 'sheet'): void
  sizeDialog(axis: 'row' | 'col'): void
  autoFit(axis: 'row' | 'col'): void
  hide(axis: 'row' | 'col', hidden: boolean): void
  autoSum(fn?: 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN'): void
  clear(kind: 'all' | 'contents' | 'formats'): void
  find(replace?: boolean): void
  insertFunction(): void
  insertFunctionNamed(name: string): void
  defineName(): void
  showCompat(): void
  showShortcuts(): void
  showAbout(): void
  importFile(): void
  exportDialog(): void
  fillDown(): void
  fillRight(): void
  /** Phase 2–3 features (sorting, filters, tables, charts, printing…). */
  feature(x: FeatureAction): void
}

export interface ViewState {
  zoom: number
  showGrid: boolean
  showHeadings: boolean
  showFormulaBar: boolean
  showFormulas: boolean
  dark: boolean
}
