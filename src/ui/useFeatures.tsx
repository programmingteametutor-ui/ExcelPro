/* Phase 2–3 feature wiring for the editor: commands, dialogs, popovers, side panels and grid overlays. */
import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import type { Workbook } from '../engine/workbook'
import type { CellStyle, CFRule, ChartType, CondFormat, Range, TableDef, TotalFn } from '../model/types'
import { DEFAULT_COL_WIDTH, DEFAULT_PAGE_SETUP, DEFAULT_ROW_HEIGHT, MAX_COLS, MAX_ROWS } from '../model/types'
import { cellKey, colToLetters, parseRangeA1, rangeToA1 } from '../model/address'
import { AxisMetrics } from '../model/axis'
import type { GridProps, ObjectsApi } from '../grid/Grid'
import type { Selection } from '../grid/selection'
import { clampToUsed, currentRegion } from '../engine/features/common'
import {
  addValidation,
  checkInput,
  clearAllFilters,
  guessHeader,
  invalidCells,
  listItems,
  reapplyFilter,
  removeValidations,
  setAutoFilter,
  setFilterCriterion,
  setFreeze,
  setLink,
  setNote,
  setPageSetup,
  sortRange,
  validationAt,
  type SortKey,
} from '../engine/features/data'
import { addCondFormat, moveCondFormat, removeCondFormats, updateCondFormat } from '../engine/features/condFormat'
import { convertToRange, createTable, tableAt, updateTable } from '../engine/features/tables'
import { createPivot, pivotAt, refreshAllPivots, refreshPivot, removePivot, updatePivot, readPivotSource } from '../engine/features/pivot'
import { flashFill, goalSeek, removeDuplicates, textToColumns, transformText, type TextTransform } from '../engine/features/clean'
import { addChart, newChart, removeChart } from '../charts/data'
import { layoutPages, pageBreaks } from '../print/paginate'
import { exportPdf } from '../print/pdf'
import { Button, Dialog } from './components/ui'
import {
  CFManagerDialog,
  CFRuleDialog,
  CreateTableDialog,
  FormatCellsDialog,
  GoalSeekDialog,
  HyperlinkDialog,
  NameManagerDialog,
  NoteDialog,
  PageSetupDialog,
  RemoveDuplicatesDialog,
  SortDialog,
  TextToColumnsDialog,
  ValidationDialog,
} from './dialogs2'
import { FilterMenu, ListPicker } from './popovers'
import { ChartPanel, ChartsLayer } from './charts'
import { CreatePivotDialog, PivotPanel, type PivotFields } from './pivot'
import { PrintPreviewDialog } from './PrintPreview'
import { download, safeFileName } from './lib/utils'

export type FeatureAction =
  | { a: 'formatCells' }
  | { a: 'indent'; delta: 1 | -1 }
  | { a: 'cfNew'; rule?: CFRule }
  | { a: 'cfQuick'; rule: CFRule }
  | { a: 'cfManage' }
  | { a: 'cfClear'; scope: 'selection' | 'sheet' }
  | { a: 'table' }
  | { a: 'tableUpdate'; patch: Partial<Omit<TableDef, 'id' | 'range'>> }
  | { a: 'tableTotal'; fn: TotalFn }
  | { a: 'tableConvert' }
  | { a: 'sort'; desc: boolean }
  | { a: 'sortCustom' }
  | { a: 'filterToggle' }
  | { a: 'filterClear' }
  | { a: 'filterReapply' }
  | { a: 'validation' }
  | { a: 'circleInvalid'; on: boolean }
  | { a: 'textToColumns' }
  | { a: 'flashFill' }
  | { a: 'removeDuplicates' }
  | { a: 'clean'; kind: TextTransform }
  | { a: 'goalSeek' }
  | { a: 'pivot' }
  | { a: 'refreshPivots' }
  | { a: 'chart'; type: ChartType }
  | { a: 'hyperlink' }
  | { a: 'note' }
  | { a: 'deleteNote' }
  | { a: 'nameManager' }
  | { a: 'freeze'; mode: 'panes' | 'top' | 'first' | 'none' }
  | { a: 'pageSetup' }
  | { a: 'printPreview' }
  | { a: 'pdf' }
  | { a: 'printArea'; set: boolean }
  | { a: 'pageBreak'; axis: 'row' | 'col' | 'reset' }
  | { a: 'pageBreakPreview' }
  | { a: 'page'; patch: Partial<typeof DEFAULT_PAGE_SETUP> }

export interface RibbonContext {
  table?: TableDef
  tableHeaders?: string[]
  anchorCol: number
  hasFilter: boolean
  filterActive: boolean
  freeze: { rows: number; cols: number } | null
  showInvalid: boolean
  pageBreakPreview: boolean
  pageSetup: typeof DEFAULT_PAGE_SETUP
  hasNote: boolean
  pivotCount: number
}

type FDialog =
  | null
  | { kind: 'formatCells' }
  | { kind: 'cfRule'; initial?: CondFormat }
  | { kind: 'cfManager' }
  | { kind: 'table'; range: Range; header: boolean }
  | { kind: 'sort'; range: Range; header: boolean }
  | { kind: 'validation' }
  | { kind: 'textToColumns'; range: Range }
  | { kind: 'removeDuplicates'; range: Range; header: boolean }
  | { kind: 'goalSeek' }
  | { kind: 'pivot'; range: Range }
  | { kind: 'hyperlink' }
  | { kind: 'note' }
  | { kind: 'names' }
  | { kind: 'pageSetup'; back?: 'preview' }
  | { kind: 'preview' }
  | { kind: 'validationAlert'; style: 'stop' | 'warning' | 'information'; title: string; message: string; onYes: () => void; onRetry: () => void }

export interface FeatureDeps {
  wb: Workbook
  sheetId: string
  version: number
  selection: Selection
  setSelection: (s: Selection) => void
  ranges: Range[]
  prim: Range
  anchorStyle: CellStyle | undefined
  notify: (m: string) => void
  dark: boolean
  editing: boolean
  focusGrid: () => void
  navigate: (target: string) => string | null
  openDefineName: () => void
  sortKeysFor?: never
}

const isWhole = (rg: Range) => rg.r2 >= MAX_ROWS - 1 || rg.c2 >= MAX_COLS - 1

export function useFeatures(d: FeatureDeps) {
  const { wb, sheetId, version, selection, prim, ranges, notify } = d
  const [dialog, setDialog] = useState<FDialog>(null)
  const [filterMenu, setFilterMenu] = useState<{ col: number; rect: DOMRect } | null>(null)
  const [listMenu, setListMenu] = useState<{ r: number; c: number; rect: DOMRect } | null>(null)
  const [showInvalid, setShowInvalid] = useState(false)
  const [breakPreview, setBreakPreview] = useState(false)
  const [chartSel, setChartSel] = useState<string | null>(null)
  const [chartPanel, setChartPanel] = useState<string | null>(null)
  const [pivotDismissed, setPivotDismissed] = useState<string | null>(null)
  const [pivotErr, setPivotErr] = useState<string | null>(null)
  const meta = wb.meta(sheetId)
  const a = selection.anchor

  const close = useCallback(() => {
    setDialog(null)
    setTimeout(d.focusGrid, 0)
  }, [d.focusGrid])

  const guard = (fn: () => string | null | void) => {
    try {
      const r = fn()
      if (typeof r === 'string' && r) notify(r)
    } catch (e) {
      notify(e instanceof Error ? e.message : String(e))
    }
  }

  /** Selection when it is a real block, otherwise the contiguous data around the active cell. */
  const dataRange = (): Range | null => {
    const multi = prim.r1 !== prim.r2 || prim.c1 !== prim.c2
    if (multi) {
      const rg = clampToUsed(wb, sheetId, prim)
      return rg.r2 >= rg.r1 && rg.c2 >= rg.c1 ? rg : null
    }
    const rg = currentRegion(wb, sheetId, a.r, a.c)
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && wb.isCellEmpty(sheetId, a.r, a.c)) return null
    return rg
  }

  // chart deselection when the cell selection changes
  const lastSel = useRef(selection)
  useEffect(() => {
    if (lastSel.current !== selection) {
      lastSel.current = selection
      setChartSel(null)
    }
  }, [selection])
  useEffect(() => {
    setChartSel(null)
    setChartPanel(null)
    setFilterMenu(null)
    setListMenu(null)
  }, [sheetId])

  // Delete key removes the selected chart
  useEffect(() => {
    if (!chartSel) return
    const k = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement
      if (t.tagName === 'INPUT' || t.tagName === 'SELECT' || (t.tagName === 'TEXTAREA' && !t.classList.contains('xp-keytrap'))) return
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault()
        e.stopPropagation()
        removeChart(wb, sheetId, chartSel)
        setChartSel(null)
        setChartPanel(null)
      } else if (e.key === 'Escape') setChartSel(null)
    }
    window.addEventListener('keydown', k, true)
    return () => window.removeEventListener('keydown', k, true)
  }, [chartSel, wb, sheetId])

  // Ctrl+P opens print preview from anywhere in the editor (instead of the browser's print dialog)
  const runRef = useRef<(x: FeatureAction) => void>(() => {})
  useEffect(() => {
    const k = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && e.key.toLowerCase() === 'p') {
        e.preventDefault()
        e.stopPropagation()
        runRef.current({ a: 'printPreview' })
      }
    }
    window.addEventListener('keydown', k, true)
    return () => window.removeEventListener('keydown', k, true)
  }, [])

  // pivot tables that refresh when their source changes
  const pivotSigs = useRef(new Map<string, string>())
  useEffect(() => {
    const auto = wb.sheets.flatMap((s) => (s.pivots ?? []).filter((p) => p.autoRefresh).map((p) => ({ sheetId: s.id, p })))
    if (!auto.length) return
    const t = setTimeout(() => {
      for (const { sheetId: sid, p } of auto) {
        if (!wb.sheets.some((s) => s.id === p.source.sheetId)) continue
        const sig = JSON.stringify(readPivotSource(wb, p.source.sheetId, p.source.range).rows)
        const prev = pivotSigs.current.get(p.id)
        pivotSigs.current.set(p.id, sig)
        if (prev !== undefined && prev !== sig) refreshPivot(wb, sid, p.id)
      }
    }, 350)
    return () => clearTimeout(t)
  }, [version, wb])

  // ------------------------------------------------------------ commands
  const sortBy = (desc: boolean, col = a.c) =>
    guard(() => {
      const af = meta.autoFilter
      const t = tableAt(wb, sheetId, a.r, a.c)
      let rg: Range | null
      let header: boolean
      if (af && col >= af.range.c1 && col <= af.range.c2 && a.r >= af.range.r1 && (a.r <= af.range.r2 + 1 || filterMenu)) {
        rg = { ...af.range, r2: Math.max(af.range.r2, clampToUsed(wb, sheetId, { ...af.range, r2: MAX_ROWS - 1 }).r2) }
        header = true
      } else if (t) {
        rg = t.totalRow ? { ...t.range, r2: t.range.r2 - 1 } : t.range
        header = true
      } else {
        rg = dataRange()
        if (!rg) return 'Select a cell in the data you want to sort.'
        header = guessHeader(wb, sheetId, rg)
      }
      return sortRange(wb, sheetId, rg, [{ col, desc }], header)
    })

  const run = (x: FeatureAction) => {
    switch (x.a) {
      case 'formatCells':
        return setDialog({ kind: 'formatCells' })
      case 'indent':
        return guard(() => wb.applyStyle(sheetId, ranges, (s) => {
          const n = Math.max(0, Math.min(15, (s?.indent ?? 0) + x.delta))
          return { ...s, indent: n || undefined, hAlign: s?.hAlign ?? (n ? 'left' : undefined) }
        }, x.delta > 0 ? 'Increase indent' : 'Decrease indent'))
      case 'cfNew':
        return setDialog({ kind: 'cfRule', initial: x.rule ? { id: '', range: prim, rule: x.rule } : undefined })
      case 'cfQuick':
        return guard(() => void addCondFormat(wb, sheetId, { range: isWhole(prim) ? clampToUsed(wb, sheetId, prim) : prim, rule: x.rule }))
      case 'cfManage':
        return setDialog({ kind: 'cfManager' })
      case 'cfClear':
        return guard(() => removeCondFormats(wb, sheetId, 'all', x.scope === 'selection' ? prim : undefined))
      case 'table': {
        if (tableAt(wb, sheetId, a.r, a.c)) return notify('The selection is already part of a table.')
        const rg = dataRange()
        if (!rg) return notify('Select the data for the table first.')
        return setDialog({ kind: 'table', range: rg, header: guessHeader(wb, sheetId, rg) })
      }
      case 'tableUpdate': {
        const t = tableAt(wb, sheetId, a.r, a.c)
        if (t) guard(() => updateTable(wb, sheetId, t.id, x.patch))
        return
      }
      case 'tableTotal': {
        const t = tableAt(wb, sheetId, a.r, a.c)
        if (!t) return
        return guard(() => updateTable(wb, sheetId, t.id, { totalRow: true, totals: { ...t.totals, [a.c - t.range.c1]: x.fn } }))
      }
      case 'tableConvert': {
        const t = tableAt(wb, sheetId, a.r, a.c)
        if (t) guard(() => convertToRange(wb, sheetId, t.id))
        return
      }
      case 'sort':
        return sortBy(x.desc)
      case 'sortCustom': {
        const t = tableAt(wb, sheetId, a.r, a.c)
        const af = meta.autoFilter
        const rg = t ? (t.totalRow ? { ...t.range, r2: t.range.r2 - 1 } : t.range) : af && a.c >= af.range.c1 && a.c <= af.range.c2 && a.r >= af.range.r1 ? { ...af.range, r2: clampToUsed(wb, sheetId, { ...af.range, r2: MAX_ROWS - 1 }).r2 } : dataRange()
        if (!rg) return notify('Select a cell in the data you want to sort.')
        return setDialog({ kind: 'sort', range: rg, header: t || af ? true : guessHeader(wb, sheetId, rg) })
      }
      case 'filterToggle':
        return guard(() => {
          if (meta.autoFilter) return setAutoFilter(wb, sheetId, null)
          const rg = dataRange()
          if (!rg) return 'Select a cell in your data to filter it.'
          setAutoFilter(wb, sheetId, rg)
        })
      case 'filterClear':
        return guard(() => clearAllFilters(wb, sheetId))
      case 'filterReapply':
        return guard(() => reapplyFilter(wb, sheetId))
      case 'validation':
        return setDialog({ kind: 'validation' })
      case 'circleInvalid':
        if (x.on) {
          const n = invalidCells(wb, sheetId).length
          notify(n ? `${n} cell${n === 1 ? '' : 's'} with invalid data circled.` : 'No invalid data found on this sheet.')
        }
        return setShowInvalid(x.on)
      case 'textToColumns': {
        const rg = clampToUsed(wb, sheetId, prim)
        if (rg.c1 !== rg.c2) return notify('Select cells in one column to split them.')
        return setDialog({ kind: 'textToColumns', range: rg })
      }
      case 'flashFill':
        return guard(() => {
          const reg = currentRegion(wb, sheetId, a.r, Math.max(0, a.c - 1))
          const r1 = guessHeader(wb, sheetId, reg) ? reg.r1 + 1 : reg.r1
          const res = flashFill(wb, sheetId, a.c, Math.min(r1, a.r), Math.max(reg.r2, a.r))
          return res.error ?? `Flash Fill filled ${res.filled} cell${res.filled === 1 ? '' : 's'}.`
        })
      case 'removeDuplicates': {
        const rg = dataRange()
        if (!rg) return notify('Select the data to check for duplicates.')
        return setDialog({ kind: 'removeDuplicates', range: rg, header: guessHeader(wb, sheetId, rg) })
      }
      case 'clean':
        return guard(() => {
          const n = transformText(wb, sheetId, ranges.map((r) => clampToUsed(wb, sheetId, r)), x.kind)
          return n ? `Changed ${n} cell${n === 1 ? '' : 's'}.` : 'Nothing to change in the selection.'
        })
      case 'goalSeek':
        return setDialog({ kind: 'goalSeek' })
      case 'pivot': {
        const tb = tableAt(wb, sheetId, a.r, a.c)
        const rg = tb ? (tb.totalRow ? { ...tb.range, r2: tb.range.r2 - 1 } : tb.range) : dataRange()
        if (!rg || rg.r2 === rg.r1) return notify('Select a range with a header row and at least one row of data.')
        return setDialog({ kind: 'pivot', range: rg })
      }
      case 'refreshPivots':
        return guard(() => {
          const n = refreshAllPivots(wb)
          return n ? `Refreshed ${n} PivotTable${n === 1 ? '' : 's'}.` : 'This workbook has no PivotTables.'
        })
      case 'chart':
        return guard(() => {
          const rg = dataRange()
          if (!rg) return 'Select the data to chart first, including any header row and labels.'
          const rows = new AxisMetrics(DEFAULT_ROW_HEIGHT, MAX_ROWS, meta.rowHeights, [...meta.hiddenRows, ...(meta.filterHidden ?? [])])
          const cols = new AxisMetrics(DEFAULT_COL_WIDTH, MAX_COLS, meta.colWidths, meta.hiddenCols)
          const ch = newChart(wb, sheetId, { sheetId, range: rg }, x.type, { x: cols.offsetOf(rg.c2 + 1) + 24, y: rows.offsetOf(rg.r1) })
          addChart(wb, sheetId, ch)
          setChartSel(ch.id)
          setChartPanel(ch.id)
        })
      case 'hyperlink':
        return setDialog({ kind: 'hyperlink' })
      case 'note':
        return setDialog({ kind: 'note' })
      case 'deleteNote':
        return guard(() => setNote(wb, sheetId, a.r, a.c, null))
      case 'nameManager':
        return setDialog({ kind: 'names' })
      case 'freeze':
        return guard(() => {
          if (x.mode === 'none') return setFreeze(wb, sheetId, 0, 0)
          if (x.mode === 'top') return setFreeze(wb, sheetId, 1, 0)
          if (x.mode === 'first') return setFreeze(wb, sheetId, 0, 1)
          if (a.r === 0 && a.c === 0) return 'Select the cell below and to the right of where the panes should freeze.'
          setFreeze(wb, sheetId, a.r, a.c)
        })
      case 'pageSetup':
        return setDialog({ kind: 'pageSetup' })
      case 'printPreview':
        return setDialog({ kind: 'preview' })
      case 'pdf':
        return void doPdf(false)
      case 'printArea':
        return guard(() => setPageSetup(wb, sheetId, { printArea: x.set ? clampToUsed(wb, sheetId, prim) : undefined }, x.set ? 'Set print area' : 'Clear print area'))
      case 'pageBreak':
        return guard(() => {
          const ps = { ...DEFAULT_PAGE_SETUP, ...meta.pageSetup }
          if (x.axis === 'reset') return setPageSetup(wb, sheetId, { rowBreaks: [], colBreaks: [] }, 'Reset page breaks')
          if (x.axis === 'row') {
            if (a.r === 0) return 'Select a cell below row 1 to insert a page break above it.'
            return setPageSetup(wb, sheetId, { rowBreaks: [...new Set([...ps.rowBreaks, a.r])].sort((p, q) => p - q) }, 'Insert page break')
          }
          if (a.c === 0) return 'Select a cell right of column A to insert a page break left of it.'
          setPageSetup(wb, sheetId, { colBreaks: [...new Set([...ps.colBreaks, a.c])].sort((p, q) => p - q) }, 'Insert page break')
        })
      case 'pageBreakPreview':
        return setBreakPreview((v) => !v)
      case 'page':
        return guard(() => setPageSetup(wb, sheetId, x.patch))
    }
  }

  runRef.current = run

  const doPdf = async (all: boolean) => {
    try {
      const ids = all ? wb.order.filter((id) => !wb.meta(id).hidden) : [sheetId]
      const bytes = await exportPdf(wb, ids, wb.name)
      const name = safeFileName(`${wb.name}.pdf`)
      await download(bytes, name, 'application/pdf')
      notify(`Downloaded ${name}`)
    } catch (e) {
      notify(e instanceof Error ? e.message : String(e))
    }
  }

  // ------------------------------------------------------------ validation on commit
  /** Returns true when the edit may be committed now; otherwise shows the validation alert. */
  const guardCommit = (target: string, r: number, c: number, text: string, commit: () => void, retry: () => void): boolean => {
    const res = checkInput(wb, target, r, c, text)
    if (res.ok) return true
    setDialog({ kind: 'validationAlert', style: res.style, title: res.title, message: res.message, onYes: commit, onRetry: retry })
    return false
  }

  // ------------------------------------------------------------ grid props
  const invalid = useMemo(() => (showInvalid ? new Set(invalidCells(wb, sheetId).map((x) => `${x.r},${x.c}`)) : null), [showInvalid, wb, sheetId, version])
  const breaks = useMemo(() => {
    if (!breakPreview) return null
    const l = layoutPages(wb, sheetId)
    return l ? pageBreaks(l) : null
  }, [breakPreview, wb, sheetId, version])
  const inputMsg = !d.editing ? validationAt(wb, sheetId, a.r, a.c) : undefined

  const renderObjects = (api: ObjectsApi): ReactNode => (
    <>
      <ChartsLayer wb={wb} sheetId={sheetId} version={version} api={api} selected={chartSel} onSelect={setChartSel} onOpenEditor={(id) => (setChartSel(id), setChartPanel(id))} dark={d.dark} />
      {inputMsg && (inputMsg.inputTitle || inputMsg.inputMessage) && (() => {
        const rc = api.cellRect(a.r, a.c)
        return (
          <div className="pointer-events-none absolute z-40 max-w-60 rounded-md border border-[var(--border)] bg-[var(--bg)] px-2 py-1.5 text-[12px] text-[var(--text)] shadow-lg" style={{ left: rc.x + 8, top: rc.y + rc.h + 4 }} data-testid="input-message">
            {inputMsg.inputTitle && <div className="font-semibold">{inputMsg.inputTitle}</div>}
            {inputMsg.inputMessage && <div className="whitespace-pre-wrap">{inputMsg.inputMessage}</div>}
          </div>
        )
      })()}
    </>
  )

  const gridProps: Partial<GridProps> = {
    invalidCells: invalid,
    pageBreaks: breaks,
    onFilterButton: (col, rect) => setFilterMenu({ col, rect }),
    onDropdown: (r, c, rect) => setListMenu({ r, c, rect }),
    onOpenLink: (url) => {
      if (url.startsWith('#')) {
        const err = d.navigate(url.slice(1))
        if (err) notify(err)
      } else window.open(url, '_blank', 'noopener,noreferrer')
    },
    renderObjects,
  }

  // ------------------------------------------------------------ context for the ribbon
  const t = tableAt(wb, sheetId, a.r, a.c)
  const ctx: RibbonContext = {
    table: t,
    anchorCol: a.c,
    hasFilter: !!meta.autoFilter,
    filterActive: !!meta.autoFilter && Object.keys(meta.autoFilter.criteria).length > 0,
    freeze: meta.freeze ?? null,
    showInvalid,
    pageBreakPreview: breakPreview,
    pageSetup: { ...DEFAULT_PAGE_SETUP, ...meta.pageSetup },
    hasNote: !!meta.notes?.[cellKey(a.r, a.c)],
    pivotCount: wb.sheets.reduce((n, s) => n + (s.pivots?.length ?? 0), 0),
  }

  // ------------------------------------------------------------ side panel
  const chart = chartPanel ? meta.charts?.find((c) => c.id === chartPanel) : undefined
  const pivot = !chart ? pivotAt(wb, sheetId, a.r, a.c) : undefined
  let side: ReactNode = null
  if (chart)
    side = <ChartPanel key={chart.id} wb={wb} sheetId={sheetId} chart={chart} onClose={() => (setChartPanel(null), setChartSel(null))} onDelete={() => (removeChart(wb, sheetId, chart.id), setChartPanel(null), setChartSel(null))} />
  else if (pivot && pivotDismissed !== pivot.id)
    side = (
      <PivotPanel
        key={pivot.id}
        wb={wb}
        pivot={pivot}
        error={pivotErr}
        onChange={(f: PivotFields) => {
          if (!f.values.length) return setPivotErr('Put at least one field in Values.')
          const err = updatePivot(wb, sheetId, pivot.id, f)
          setPivotErr(err)
        }}
        onRefresh={() => setPivotErr(refreshPivot(wb, sheetId, pivot.id))}
        onDelete={() => (removePivot(wb, sheetId, pivot.id), setPivotErr(null))}
        onClose={() => setPivotDismissed(pivot.id)}
      />
    )

  // ------------------------------------------------------------ dialogs and popovers
  const colName = (c: number) => (meta.autoFilter ? wb.getDisplay(sheetId, meta.autoFilter.range.r1, c).text || colToLetters(c) : colToLetters(c))
  const sample = wb.getValue(sheetId, a.r, a.c)
  const ui: ReactNode = (
    <>
      {filterMenu && (
        <FilterMenu
          key={filterMenu.col}
          wb={wb}
          sheetId={sheetId}
          col={filterMenu.col}
          rect={filterMenu.rect}
          colName={colName(filterMenu.col)}
          onClose={() => (setFilterMenu(null), d.focusGrid())}
          onSort={(desc) => {
            sortBy(desc, filterMenu.col)
            setFilterMenu(null)
          }}
          onApply={(crit) => {
            guard(() => setFilterCriterion(wb, sheetId, filterMenu.col, crit))
            setFilterMenu(null)
            d.focusGrid()
          }}
        />
      )}
      {listMenu && (() => {
        const v = validationAt(wb, sheetId, listMenu.r, listMenu.c)
        if (!v) return null
        return (
          <ListPicker
            items={listItems(wb, sheetId, v)}
            rect={listMenu.rect}
            current={wb.getDisplay(sheetId, listMenu.r, listMenu.c).text}
            onClose={() => (setListMenu(null), d.focusGrid())}
            onPick={(item) => {
              guard(() => wb.setInputs(sheetId, [{ r: listMenu.r, c: listMenu.c, input: item }], 'Pick from list'))
              setListMenu(null)
              d.focusGrid()
            }}
          />
        )
      })()}
      {dialog?.kind === 'formatCells' && (
        <FormatCellsDialog
          initial={d.anchorStyle}
          sample={sample}
          onClose={close}
          onApply={(s) => {
            const before = (d.anchorStyle ?? {}) as Record<string, unknown>
            const after = s as Record<string, unknown>
            const patch: Record<string, unknown> = {}
            for (const k of new Set([...Object.keys(before), ...Object.keys(after)])) if (JSON.stringify(before[k]) !== JSON.stringify(after[k])) patch[k] = after[k]
            if (Object.keys(patch).length) guard(() => wb.applyStyle(sheetId, ranges, patch as Partial<CellStyle>, 'Format cells'))
            close()
          }}
        />
      )}
      {dialog?.kind === 'cfRule' && (
        <CFRuleDialog
          initial={dialog.initial}
          range={isWhole(prim) ? clampToUsed(wb, sheetId, prim) : prim}
          onClose={() => setDialog(dialog.initial?.id ? { kind: 'cfManager' } : null)}
          onSave={(cf) => {
            guard(() => {
              if (dialog.initial?.id) updateCondFormat(wb, sheetId, dialog.initial.id, cf)
              else addCondFormat(wb, sheetId, cf)
            })
            setDialog(dialog.initial?.id ? { kind: 'cfManager' } : null)
          }}
        />
      )}
      {dialog?.kind === 'cfManager' && (
        <CFManagerDialog
          rules={meta.condFormats ?? []}
          onClose={close}
          onNew={() => setDialog({ kind: 'cfRule' })}
          onEdit={(cf) => setDialog({ kind: 'cfRule', initial: cf })}
          onDelete={(id) => guard(() => removeCondFormats(wb, sheetId, [id]))}
          onMove={(id, dir) => guard(() => moveCondFormat(wb, sheetId, id, dir))}
        />
      )}
      {dialog?.kind === 'table' && (
        <CreateTableDialog
          range={dialog.range}
          headerGuess={dialog.header}
          onClose={close}
          onCreate={(range, hasHeader, style) => {
            const res = createTable(wb, sheetId, range, { hasHeader, style })
            if (res.error) return notify(res.error)
            close()
          }}
        />
      )}
      {dialog?.kind === 'sort' && (
        <SortDialog
          wb={wb}
          sheetId={sheetId}
          range={dialog.range}
          headerGuess={dialog.header}
          onClose={close}
          onSort={(keys: SortKey[], hasHeader) => {
            guard(() => sortRange(wb, sheetId, dialog.range, keys, hasHeader))
            close()
          }}
        />
      )}
      {dialog?.kind === 'validation' && (
        <ValidationDialog
          initial={validationAt(wb, sheetId, a.r, a.c)}
          range={isWhole(prim) ? { ...prim } : prim}
          onClose={close}
          onClear={() => {
            guard(() => removeValidations(wb, sheetId, prim))
            close()
          }}
          onSave={(v) => {
            const cur = validationAt(wb, sheetId, a.r, a.c)
            guard(() => void addValidation(wb, sheetId, { ...v, id: cur?.id }))
            close()
          }}
        />
      )}
      {dialog?.kind === 'textToColumns' && (
        <TextToColumnsDialog
          onClose={close}
          onApply={(delim) => {
            guard(() => textToColumns(wb, sheetId, dialog.range, delim))
            close()
          }}
        />
      )}
      {dialog?.kind === 'removeDuplicates' && (
        <RemoveDuplicatesDialog
          wb={wb}
          sheetId={sheetId}
          range={dialog.range}
          headerGuess={dialog.header}
          onClose={close}
          onApply={(cols, hasHeader) => {
            guard(() => {
              const total = dialog.range.r2 - dialog.range.r1 + 1 - (hasHeader ? 1 : 0)
              const n = removeDuplicates(wb, sheetId, dialog.range, cols, hasHeader)
              return n ? `Removed ${n} duplicate row${n === 1 ? '' : 's'}; ${total - n} unique row${total - n === 1 ? '' : 's'} remain.` : 'No duplicate rows found.'
            })
            close()
          }}
        />
      )}
      {dialog?.kind === 'goalSeek' && (
        <GoalSeekDialog
          initialSet={wb.isFormula(sheetId, a.r, a.c) ? `${colToLetters(a.c)}${a.r + 1}` : ''}
          onClose={close}
          onRun={(setCell, target, changing) => {
            const s = parseRangeA1(setCell.replace(/\$/g, ''))
            const c = parseRangeA1(changing.replace(/\$/g, ''))
            if (!s || !c) return 'Enter cell references such as B5.'
            if (!Number.isFinite(target)) return 'The target value must be a number.'
            const res = goalSeek(wb, sheetId, { r: s.r1, c: s.c1 }, target, { r: c.r1, c: c.c1 })
            return res.message
          }}
        />
      )}
      {dialog?.kind === 'pivot' && (
        <CreatePivotDialog
          wb={wb}
          sheetId={sheetId}
          initialRange={dialog.range}
          anchorFree={{ r: dialog.range.r1, c: dialog.range.c2 + 2 }}
          onClose={close}
          onCreate={({ source, dest, fields }) => {
            let target = dest === 'new' ? null : dest
            if (!target) {
              const id = wb.addSheet(wb.nextSheetName('Pivot'))
              target = { sheetId: id, r: 2, c: 0 }
            }
            const res = createPivot(wb, target.sheetId, { ...fields, source, target: { r: target.r, c: target.c } })
            if (res.error) return notify(res.error)
            if (wb.activeSheetId !== target.sheetId) wb.setActiveSheet(target.sheetId)
            d.setSelection({ anchor: { r: target.r, c: target.c }, focus: { r: target.r, c: target.c }, extra: [] })
            setPivotDismissed(null)
            setDialog(null)
          }}
        />
      )}
      {dialog?.kind === 'hyperlink' && (
        <HyperlinkDialog
          initialUrl={meta.links?.[cellKey(a.r, a.c)] ?? ''}
          initialText={wb.getDisplay(sheetId, a.r, a.c).text}
          onClose={close}
          onRemove={meta.links?.[cellKey(a.r, a.c)] ? () => (guard(() => setLink(wb, sheetId, a.r, a.c, null)), close()) : undefined}
          onSave={(url, text) => {
            guard(() => setLink(wb, sheetId, a.r, a.c, url, wb.isFormula(sheetId, a.r, a.c) ? undefined : text || undefined))
            close()
          }}
        />
      )}
      {dialog?.kind === 'note' && (
        <NoteDialog
          initial={meta.notes?.[cellKey(a.r, a.c)] ?? ''}
          onClose={close}
          onDelete={meta.notes?.[cellKey(a.r, a.c)] ? () => (guard(() => setNote(wb, sheetId, a.r, a.c, null)), close()) : undefined}
          onSave={(text) => {
            guard(() => setNote(wb, sheetId, a.r, a.c, text))
            close()
          }}
        />
      )}
      {dialog?.kind === 'names' && <NameManagerDialog wb={wb} onClose={close} onNew={() => (setDialog(null), d.openDefineName())} onChanged={() => setDialog({ kind: 'names' })} />}
      {dialog?.kind === 'pageSetup' && (
        <PageSetupDialog
          initial={meta.pageSetup}
          selection={clampToUsed(wb, sheetId, prim)}
          sheetName={meta.name}
          onClose={() => (dialog.back ? setDialog({ kind: 'preview' }) : close())}
          onSave={(p) => {
            guard(() => setPageSetup(wb, sheetId, p))
            if (dialog.back) setDialog({ kind: 'preview' })
            else close()
          }}
        />
      )}
      {dialog?.kind === 'preview' && <PrintPreviewDialog wb={wb} sheetId={sheetId} version={version} onClose={close} onPageSetup={() => setDialog({ kind: 'pageSetup', back: 'preview' })} onPdf={doPdf} />}
      {dialog?.kind === 'validationAlert' && (
        <Dialog
          title={dialog.title}
          onClose={() => setDialog(null)}
          width={400}
          footer={
            dialog.style === 'stop' ? (
              <>
                <Button variant="outline" onClick={() => setDialog(null)}>Cancel</Button>
                <Button variant="default" onClick={() => (setDialog(null), dialog.onRetry())}>Retry</Button>
              </>
            ) : dialog.style === 'warning' ? (
              <>
                <Button variant="outline" onClick={() => setDialog(null)}>Cancel</Button>
                <Button variant="outline" onClick={() => (setDialog(null), dialog.onRetry())}>No</Button>
                <Button variant="default" onClick={() => (setDialog(null), dialog.onYes())}>Yes</Button>
              </>
            ) : (
              <>
                <Button variant="outline" onClick={() => setDialog(null)}>Cancel</Button>
                <Button variant="default" onClick={() => (setDialog(null), dialog.onYes())}>OK</Button>
              </>
            )
          }
        >
          <div className="flex gap-3" data-testid="validation-alert">
            <span className={dialog.style === 'stop' ? 'text-[var(--danger)]' : 'text-[var(--warn-text)]'}>●</span>
            <p className="whitespace-pre-wrap text-[13px]">
              {dialog.message}
              {dialog.style === 'warning' && '\n\nContinue?'}
            </p>
          </div>
        </Dialog>
      )}
    </>
  )

  return { run, ui, side, gridProps, ctx, guardCommit, chartSelected: chartSel, rangeLabel: rangeToA1(prim) }
}
