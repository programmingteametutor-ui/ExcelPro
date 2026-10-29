import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { ClipboardData, PasteMode, Workbook } from '../engine/workbook'
import { parseTsv } from '../engine/workbook'
import { adjustDecimals, formatGeneral } from '../engine/numberFormat'
import type { DateOrder } from '../engine/inputParser'
import { exportXlsx } from '../files/xlsx'
import { exportDelimited } from '../files/csv'
import { addrToA1, cellKey, colToLetters, parseRangeA1, quoteSheetName, rangeToA1 } from '../model/address'
import type { CellAddr, CellStyle, Range } from '../model/types'
import { DEFAULT_COL_WIDTH, DEFAULT_FONT, DEFAULT_FONT_SIZE, DEFAULT_ROW_HEIGHT, MAX_COLS, MAX_ROWS } from '../model/types'
import type { Autosave, SaveState } from '../persistence/autosave'
import { Grid, type GridHandle } from '../grid/Grid'
import { allRanges, canInsertRef, primaryRange, single, type EditState, type Selection } from '../grid/selection'
import type { Commands, ViewState } from './commands'
import { Button, ContextMenu, MenuItem, MenuLabel, MenuSep, Popover, Tooltip } from './components/ui'
import {
  AboutDialog,
  CompatDialog,
  ConfirmDialog,
  DefineNameDialog,
  ExportDialog,
  FindReplaceDialog,
  ImportReportDialog,
  InsertFunctionDialog,
  ShortcutsDialog,
  SizeDialog,
  type ExportChoice,
} from './dialogs'
import { FormulaBar } from './FormulaBar'
import { Icon, Logo } from './icons'
import { cn, download, modKey, safeFileName, shortcut } from './lib/utils'
import { rememberFn, Ribbon } from './Ribbon'
import { SheetTabs } from './SheetTabs'
import { StatusBar } from './StatusBar'
import type { ImportReport } from '../files/xlsx'
import { useFeatures } from './useFeatures'
import { acceptSuggestion, analyzeFormula, FormulaAssistPanel } from './FormulaAssist'
import { maybeExtendTable } from '../engine/features/tables'

type DialogState =
  | null
  | { kind: 'find'; replace: boolean }
  | { kind: 'export' }
  | { kind: 'function' }
  | { kind: 'size'; axis: 'row' | 'col' }
  | { kind: 'name' }
  | { kind: 'compat' }
  | { kind: 'shortcuts' }
  | { kind: 'about' }
  | { kind: 'report'; report: ImportReport; fileName: string }
  | { kind: 'confirm'; title: string; message: string; confirm: string; danger?: boolean; action: () => void }

interface EditSession extends EditState {
  sheetId: string
  original: string
}

const isFullCol = (rg: Range) => rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1
const isFullRow = (rg: Range) => rg.c1 === 0 && rg.c2 >= MAX_COLS - 1

let measureCtx: CanvasRenderingContext2D | null = null
function textWidth(text: string, font: string) {
  if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d')
  if (!measureCtx) return text.length * 7
  measureCtx.font = font
  return measureCtx.measureText(text).width
}

export function Editor({
  wb,
  autosave,
  view,
  setView,
  dateOrder,
  setDateOrder,
  onHome,
  onImportFile,
  initialReport,
  onReloadFromStore,
}: {
  wb: Workbook
  autosave: Autosave
  view: ViewState
  setView: (v: Partial<ViewState>) => void
  dateOrder: DateOrder
  setDateOrder: (d: DateOrder) => void
  onHome: () => void
  onImportFile: (f: File) => void
  initialReport?: { report: ImportReport; fileName: string } | null
  onReloadFromStore: () => void
}) {
  const [version, setVersion] = useState(wb.version)
  const [saveState, setSaveState] = useState<SaveState>(autosave.state)
  const sheetId = wb.activeSheetId
  const selections = useRef(new Map<string, Selection>())
  const [selection, setSelectionState] = useState<Selection>(() => single({ r: 0, c: 0 }))
  const [editing, setEditing] = useState<EditSession | null>(null)
  const [caret, setCaret] = useState(0)
  const [pointRef, setPointRef] = useState<{ start: number; end: number } | null>(null)
  const pendingCaret = useRef<number | null>(null)
  const clipRef = useRef<ClipboardData | null>(null)
  const [copyRange, setCopyRange] = useState<{ sheetId: string; range: Range } | null>(null)
  const [painter, setPainter] = useState<(CellStyle | undefined)[][] | null>(null)
  const [dialog, setDialog] = useState<DialogState>(initialReport ? { kind: 'report', ...initialReport } : null)
  const [ctx, setCtx] = useState<{ x: number; y: number; kind: 'cell' | 'row' | 'col' } | null>(null)
  const [toast, setToast] = useState<string | null>(null)
  const [nameDraft, setNameDraft] = useState(wb.name)
  const gridRef = useRef<GridHandle>(null)
  const editorRef = useRef<HTMLTextAreaElement>(null)
  const barRef = useRef<HTMLTextAreaElement>(null)
  const fileRef = useRef<HTMLInputElement>(null)
  const toastTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const featuresRef = useRef<ReturnType<typeof useFeatures> | null>(null)
  const [assistIdx, setAssistIdx] = useState(0)
  const [assistClosed, setAssistClosed] = useState<string | null>(null)

  wb.dateOrder = dateOrder

  const notify = useCallback((msg: string) => {
    setToast(msg)
    if (toastTimer.current) clearTimeout(toastTimer.current)
    toastTimer.current = setTimeout(() => setToast(null), 4500)
  }, [])

  // ---------------------------------------------------------------- subscriptions
  useEffect(() => {
    const off = wb.subscribe(() => setVersion(wb.version))
    return () => void off()
  }, [wb])
  useEffect(() => {
    const off = autosave.onChange(setSaveState)
    return () => void off()
  }, [autosave])
  useEffect(() => {
    const h = (e: BeforeUnloadEvent) => {
      if (autosave.hasUnsavedChanges) {
        void autosave.flush()
        e.preventDefault()
        e.returnValue = ''
      }
    }
    window.addEventListener('beforeunload', h)
    return () => window.removeEventListener('beforeunload', h)
  }, [autosave])
  useEffect(() => {
    setTimeout(focusGrid, 0)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // per-sheet selection memory
  const setSelection = useCallback(
    (s: Selection) => {
      selections.current.set(wb.activeSheetId, s)
      setSelectionState(s)
    },
    [wb],
  )
  const lastSheet = useRef(sheetId)
  useEffect(() => {
    if (lastSheet.current !== sheetId) {
      lastSheet.current = sheetId
      setSelectionState(selections.current.get(sheetId) ?? single({ r: 0, c: 0 }))
    }
  }, [sheetId])

  // ---------------------------------------------------------------- derived
  const anchorStyle = wb.getStyle(sheetId, selection.anchor.r, selection.anchor.c)
  const prim = primaryRange(selection)
  const ranges = allRanges(selection)
  const address = useMemo(() => {
    if (isFullCol(prim) && isFullRow(prim)) return 'All'
    if (isFullCol(prim)) return prim.c1 === prim.c2 ? `${colToLetters(prim.c1)}:${colToLetters(prim.c1)}` : `${colToLetters(prim.c1)}:${colToLetters(prim.c2)}`
    if (isFullRow(prim)) return `${prim.r1 + 1}:${prim.r2 + 1}`
    const m = wb.mergeAt(sheetId, selection.anchor.r, selection.anchor.c)
    if (m && m.r1 === prim.r1 && m.c1 === prim.c1 && m.r2 >= prim.r2 && m.c2 >= prim.c2) return addrToA1(selection.anchor)
    return selection.extra.length || prim.r1 !== prim.r2 || prim.c1 !== prim.c2 ? rangeToA1(prim) : addrToA1(selection.anchor)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [prim.r1, prim.c1, prim.r2, prim.c2, selection, sheetId, version])

  const barText = editing ? editing.text : wb.getEditText(sheetId, selection.anchor.r, selection.anchor.c)
  const errorMsg = !editing ? wb.explainError(sheetId, selection.anchor.r, selection.anchor.c) : null

  const used = () => wb.usedSize(sheetId)
  /** Clamp a (possibly whole-row/column) range to the used area for content operations. */
  const clampContent = (rg: Range): Range => {
    const u = used()
    return { r1: rg.r1, c1: rg.c1, r2: Math.max(rg.r1, Math.min(rg.r2, u.rows - 1)), c2: Math.max(rg.c1, Math.min(rg.c2, u.cols - 1)) }
  }

  const act = useCallback(
    (fn: () => void | string | null) => {
      try {
        const r = fn()
        if (typeof r === 'string' && r) notify(r)
      } catch (e) {
        notify(e instanceof Error ? e.message : String(e))
      }
    },
    [notify],
  )

  // ---------------------------------------------------------------- editing
  const editingRef = useRef<EditSession | null>(null)
  editingRef.current = editing
  const focusGrid = useCallback(() => {
    if (!editingRef.current) gridRef.current?.focus()
  }, [])

  const startEdit = useCallback(
    (initial: string | null, mode: 'enter' | 'edit', origin: 'cell' | 'bar' = 'cell') => {
      const { r, c } = selection.anchor
      const original = wb.getEditText(sheetId, r, c)
      const text = initial ?? original
      setEditing({ text, mode, origin, sheetId, original })
      setCaret(text.length)
      pendingCaret.current = text.length
      setPointRef(null)
    },
    [selection.anchor, sheetId, wb],
  )

  useLayoutEffect(() => {
    if (!editing) return
    const el = editing.origin === 'cell' ? editorRef.current : barRef.current
    if (!el) return
    if (document.activeElement !== el) el.focus()
    if (pendingCaret.current !== null) {
      el.setSelectionRange(pendingCaret.current, pendingCaret.current)
      pendingCaret.current = null
    }
  }, [editing])

  const cancelEdit = useCallback(() => {
    setEditing(null)
    setPointRef(null)
    gridRef.current?.focus()
  }, [])

  const move = useCallback(
    (dr: number, dc: number, extend = false, jump = false, base?: Selection) => {
      const sel = base ?? selection
      const counts = gridRef.current?.counts() ?? { rows: 1000, cols: 52 }
      const meta = wb.meta(sheetId)
      let from = extend ? sel.focus : sel.anchor
      const m = wb.mergeAt(sheetId, from.r, from.c)
      if (m && !extend) {
        if (dr > 0) from = { r: m.r2, c: from.c }
        if (dc > 0) from = { r: from.r, c: m.c2 }
      }
      let to: CellAddr
      if (jump) to = wb.dataEdge(sheetId, from, Math.sign(dr), Math.sign(dc), counts.rows, counts.cols)
      else {
        let r = from.r + dr
        let c = from.c + dc
        const hr = new Set(meta.hiddenRows)
        const hc = new Set(meta.hiddenCols)
        while (hr.has(r) && r > 0 && r < MAX_ROWS - 1) r += Math.sign(dr) || 1
        while (hc.has(c) && c > 0 && c < MAX_COLS - 1) c += Math.sign(dc) || 1
        to = { r: Math.max(0, Math.min(MAX_ROWS - 1, r)), c: Math.max(0, Math.min(MAX_COLS - 1, c)) }
      }
      if (extend) setSelection({ ...sel, focus: to })
      else {
        const tm = wb.mergeAt(sheetId, to.r, to.c)
        const a = tm ? { r: tm.r1, c: tm.c1 } : to
        setSelection({ anchor: a, focus: tm ? { r: tm.r2, c: tm.c2 } : a, extra: [] })
      }
    },
    [selection, sheetId, wb, setSelection],
  )

  const commitEdit = useCallback(
    (dr: number, dc: number, allSelected = false) => {
      if (!editing) return
      const target = editing.sheetId
      const text = editing.text
      const { r, c } = selection.anchor
      act(() => {
        if (allSelected) {
          const rg = clampContent(prim)
          const n = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1)
          if (n > 50_000) return 'Select fewer than 50,000 cells to fill them at once.'
          const entries: { r: number; c: number; input: string }[] = []
          for (let rr = rg.r1; rr <= rg.r2; rr++)
            for (let cc = rg.c1; cc <= rg.c2; cc++) {
              const shifted = text.startsWith('=') ? text : text
              entries.push({ r: rr, c: cc, input: shifted })
            }
          // formulas get relative-reference adjustment via fill semantics
          if (text.startsWith('=')) {
            wb.setInputs(target, [{ r, c, input: text }], 'Edit cell')
            wb.fill(target, { r1: r, c1: c, r2: r, c2: c }, { r1: r, c1: c, r2: rg.r2, c2: c })
            if (rg.c2 > c) wb.fill(target, { r1: r, c1: c, r2: rg.r2, c2: c }, { r1: r, c1: c, r2: rg.r2, c2: rg.c2 })
          } else wb.setInputs(target, entries, 'Fill selection')
        } else if (text !== editing.original) {
          const doCommit = () => {
            wb.setInputs(target, [{ r, c, input: text }])
            maybeExtendTable(wb, target, r, c)
          }
          const f = featuresRef.current
          const ok = !f || f.guardCommit(
            target,
            r,
            c,
            text,
            () => act(() => {
              doCommit()
              if (dr || dc) move(dr, dc)
            }),
            () => {
              setEditing({ text, mode: 'edit', origin: 'cell', sheetId: target, original: editing.original })
              pendingCaret.current = text.length
              setCaret(text.length)
            },
          )
          if (!ok) {
            dr = 0
            dc = 0
            return
          }
          doCommit()
        }
      })
      setEditing(null)
      setPointRef(null)
      if (dr || dc) move(dr, dc)
      gridRef.current?.focus()
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [editing, selection, wb, act, move, prim],
  )

  const onEditText = useCallback(
    (text: string, pos?: number) => {
      setEditing((e) => {
        if (!e) return e
        if (e.text !== text) setPointRef(null)
        return e.text === text ? e : { ...e, text }
      })
      if (pos !== undefined) setCaret(pos)
      setAssistIdx(0)
    },
    [],
  )

  const onInsertRef = useCallback(
    (ref: string, replaceFrom: number | null) => {
      setEditing((e) => {
        if (!e) return e
        const prefix = e.sheetId !== wb.activeSheetId ? `${quoteSheetName(wb.meta(wb.activeSheetId).name)}!` : ''
        const full = prefix + ref
        let start: number
        let end: number
        if (replaceFrom !== null && pointRef && pointRef.start === replaceFrom) {
          start = pointRef.start
          end = pointRef.end
        } else if (replaceFrom !== null) {
          start = replaceFrom
          end = Math.max(replaceFrom, caret)
        } else {
          start = caret
          end = caret
        }
        const text = e.text.slice(0, start) + full + e.text.slice(end)
        setPointRef({ start, end: start + full.length })
        setCaret(start + full.length)
        pendingCaret.current = start + full.length
        return { ...e, text }
      })
    },
    [caret, pointRef, wb],
  )

  const canPointRef = useCallback(() => {
    if (!editing) return false
    if (pointRef && caret === pointRef.end) return true
    return canInsertRef(editing.text, caret)
  }, [editing, caret, pointRef])

  const assist = editing && assistClosed !== editing.text ? analyzeFormula(wb, editing.text, caret) : null
  const assistList = assist && assist.suggestions.length ? assist : null
  const acceptAssist = (i: number) => {
    if (!editing || !assistList) return
    const s = assistList.suggestions[i]
    if (!s) return
    const next = acceptSuggestion(editing.text, caret, assistList, s)
    pendingCaret.current = next.caret
    setCaret(next.caret)
    setAssistIdx(0)
    setEditing({ ...editing, text: next.text })
  }

  const onEditKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if (!editing) return
      const k = e.key
      if (assistList) {
        if (k === 'ArrowDown' || k === 'ArrowUp') {
          e.preventDefault()
          e.stopPropagation()
          const n = assistList.suggestions.length
          setAssistIdx((i) => (i + (k === 'ArrowDown' ? 1 : n - 1)) % n)
          return
        }
        if (k === 'Tab' && !e.shiftKey) {
          e.preventDefault()
          e.stopPropagation()
          return acceptAssist(Math.min(assistIdx, assistList.suggestions.length - 1))
        }
        if (k === 'Escape') {
          e.preventDefault()
          e.stopPropagation()
          setAssistClosed(editing.text)
          return
        }
      }
      if (k === '__commit__') return commitEdit(0, 0)
      if (k === 'Enter' && e.altKey) {
        e.preventDefault()
        const el = e.currentTarget
        const s = el.selectionStart
        const t = editing.text.slice(0, s) + '\n' + editing.text.slice(el.selectionEnd)
        pendingCaret.current = s + 1
        setEditing({ ...editing, text: t })
        return
      }
      if (k === 'Enter') {
        e.preventDefault()
        if (modKey(e)) return commitEdit(0, 0, true)
        return commitEdit(e.shiftKey ? -1 : 1, 0)
      }
      if (k === 'Tab') {
        e.preventDefault()
        return commitEdit(0, e.shiftKey ? -1 : 1)
      }
      if (k === 'Escape') {
        e.preventDefault()
        return cancelEdit()
      }
      if (k === 'F2') {
        e.preventDefault()
        setEditing({ ...editing, mode: editing.mode === 'enter' ? 'edit' : 'enter' })
        return
      }
      if (editing.mode === 'enter' && editing.origin === 'cell' && k.startsWith('Arrow') && !e.shiftKey && !modKey(e)) {
        e.preventDefault()
        const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k] as [number, number]
        return commitEdit(d[0], d[1])
      }
      e.stopPropagation()
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [editing, commitEdit, cancelEdit, assistList, assistIdx, caret],
  )

  // commit pending edit when the user switches sheets
  useEffect(() => {
    if (editing && editing.sheetId !== sheetId && !canInsertRef(editing.text, caret)) {
      wb.setInputs(editing.sheetId, [{ r: selections.current.get(editing.sheetId)?.anchor.r ?? 0, c: selections.current.get(editing.sheetId)?.anchor.c ?? 0, input: editing.text }])
      setEditing(null)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheetId])

  // ---------------------------------------------------------------- clipboard
  const htmlOf = (clip: ClipboardData) => {
    const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;')
    const rows = clip.text.split('\n').map((line, i) => {
      const cells = parseTsv(line)[0] ?? ['']
      return '<tr>' + cells.map((t, j) => {
        const st = clip.styles[i]?.[j]
        const css = [st?.bold && 'font-weight:bold', st?.italic && 'font-style:italic', st?.color && `color:${st.color}`, st?.fill && `background:${st.fill}`].filter(Boolean).join(';')
        return `<td${css ? ` style="${css}"` : ''}>${esc(t)}</td>`
      }).join('') + '</tr>'
    })
    return `<table>${rows.join('')}</table>`
  }

  const doCopy = (cut: boolean): ClipboardData | null => {
    if (selection.extra.length) {
      notify("This action won't work on multiple selections.")
      return null
    }
    const rg = clampContent(prim)
    if ((rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1) > 2_000_000) {
      notify('That selection is too large to copy.')
      return null
    }
    const clip = wb.copy(sheetId, rg, cut)
    clipRef.current = clip
    setCopyRange({ sheetId, range: rg })
    return clip
  }

  const onCopyEvent = (e: React.ClipboardEvent) => {
    if (editing) return
    e.preventDefault()
    const clip = doCopy(false)
    if (clip) {
      e.clipboardData.setData('text/plain', clip.text)
      e.clipboardData.setData('text/html', htmlOf(clip))
    }
  }
  const onCutEvent = (e: React.ClipboardEvent) => {
    if (editing) return
    e.preventDefault()
    const clip = doCopy(true)
    if (clip) e.clipboardData.setData('text/plain', clip.text)
  }

  const doPaste = (text: string | null, mode: PasteMode = 'all') => {
    const clip = clipRef.current
    const norm = (s: string) => s.replace(/\r\n?/g, '\n').replace(/\n$/, '')
    act(() => {
      let dest: Range | null = null
      if (clip && (text === null || norm(text) === norm(clip.text))) {
        dest = wb.paste(clip, sheetId, clampContent(prim).r2 - prim.r1 > 100000 ? { ...prim, r2: prim.r1, c2: prim.c1 } : prim, mode)
        if (clip.cut) {
          clipRef.current = null
          setCopyRange(null)
        }
      } else if (text) {
        const rows = parseTsv(text)
        if (rows.length === 1 && rows[0].length === 1 && (prim.r1 !== prim.r2 || prim.c1 !== prim.c2)) {
          const rg = clampContent(prim)
          const entries: { r: number; c: number; input: string }[] = []
          for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) entries.push({ r, c, input: rows[0][0] })
          wb.setInputs(sheetId, entries, 'Paste')
          dest = rg
        } else dest = wb.pasteText(sheetId, selection.anchor, text)
        clipRef.current = null
        setCopyRange(null)
      }
      if (dest) setSelection({ anchor: { r: dest.r1, c: dest.c1 }, focus: { r: dest.r2, c: dest.c2 }, extra: [] })
    })
  }

  const onPasteEvent = (e: React.ClipboardEvent) => {
    if (editing) return
    e.preventDefault()
    doPaste(e.clipboardData.getData('text/plain'))
  }

  // ---------------------------------------------------------------- commands
  const styleRanges = () => ranges

  const autoFit = (axis: 'row' | 'col', indices?: number[]) => {
    const meta = wb.meta(sheetId)
    const u = used()
    if (axis === 'col') {
      const cols = indices ?? range(prim.c1, Math.min(prim.c2, Math.max(prim.c1, u.cols - 1)))
      for (const c of cols.slice(0, 200)) {
        let w = 0
        for (let r = 0; r < Math.min(u.rows, 5000); r++) {
          if (wb.mergeAt(sheetId, r, c)) continue
          const d = wb.getDisplay(sheetId, r, c, view.showFormulas)
          if (!d.text) continue
          const st = meta.styles[cellKey(r, c)]
          const font = `${st?.italic ? 'italic ' : ''}${st?.bold ? '700 ' : ''}${((st?.fontSize ?? DEFAULT_FONT_SIZE) * 4) / 3}px ${st?.fontFamily ? `"${st.fontFamily}", ` : ''}${DEFAULT_FONT}`
          const longest = d.text.split('\n').reduce((a, b) => (b.length > a.length ? b : a), '')
          w = Math.max(w, textWidth(longest, font))
        }
        wb.setColWidths(sheetId, [c], w ? Math.ceil(w + 12) : null, 'AutoFit column')
      }
    } else {
      const rows = indices ?? range(prim.r1, Math.min(prim.r2, Math.max(prim.r1, u.rows - 1)))
      for (const r of rows.slice(0, 500)) {
        let h = 0
        for (let c = 0; c < Math.min(u.cols, 200); c++) {
          const st = meta.styles[cellKey(r, c)]
          const d = wb.getDisplay(sheetId, r, c)
          if (!d.text && !st?.fontSize) continue
          const px = ((st?.fontSize ?? DEFAULT_FONT_SIZE) * 4) / 3
          let lines = d.text.split('\n').length
          if (st?.wrap) {
            const font = `${st?.bold ? '700 ' : ''}${px}px ${DEFAULT_FONT}`
            const colW = meta.colWidths[c] ?? DEFAULT_COL_WIDTH
            lines = d.text.split('\n').reduce((n, l) => n + Math.max(1, Math.ceil(textWidth(l, font) / Math.max(10, colW - 8))), 0)
          }
          h = Math.max(h, lines * px * 1.25 + 6)
        }
        wb.setRowHeights(sheetId, [r], h > DEFAULT_ROW_HEIGHT ? Math.ceil(h) : null, 'AutoFit row')
      }
    }
  }

  const insertFunctionNamed = (name: string) => {
    if (editing) {
      const t = editing.text
      const ins = (t.startsWith('=') ? '' : '=') + `${name}(`
      const text = t.slice(0, caret) + ins + t.slice(caret)
      pendingCaret.current = caret + ins.length
      setEditing({ ...editing, text })
    } else startEdit(`=${name}(`, 'edit')
  }

  const cmd: Commands = {
    undo: () =>
      act(() => {
        if (editing) return cancelEdit()
        const s = wb.undo()
        if (s) {
          if (s.sheetId !== wb.activeSheetId && wb.order.includes(s.sheetId)) wb.setActiveSheet(s.sheetId)
          setSelection({ anchor: { r: s.range.r1, c: s.range.c1 }, focus: { r: s.range.r2, c: s.range.c2 }, extra: [] })
        }
      }),
    redo: () =>
      act(() => {
        const s = wb.redo()
        if (s) setSelection({ anchor: { r: s.range.r1, c: s.range.c1 }, focus: { r: s.range.r2, c: s.range.c2 }, extra: [] })
      }),
    save: () => {
      void autosave.flush().then(() => autosave.state.status === 'saved' && notify('Saved to this browser.'))
    },
    cut: () => {
      const clip = doCopy(true)
      if (clip) void navigator.clipboard?.writeText(clip.text).catch(() => {})
    },
    copy: () => {
      const clip = doCopy(false)
      if (clip) void navigator.clipboard?.writeText(clip.text).catch(() => {})
    },
    paste: (mode = 'all') => {
      if (clipRef.current) return doPaste(null, mode)
      if (mode !== 'all') return notify('Copy cells in ExcelPro first to use paste options.')
      if (!navigator.clipboard?.readText) return notify(`Your browser doesn't allow toolbar paste. Use ${shortcut('Ctrl+V')}.`)
      navigator.clipboard
        .readText()
        .then((t) => (t ? doPaste(t) : notify('The clipboard is empty.')))
        .catch(() => notify(`Your browser blocked clipboard access. Use ${shortcut('Ctrl+V')} to paste.`))
    },
    formatPainter: () => {
      if (painter) return setPainter(null)
      const rg = clampContent(prim)
      if ((rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1) > 10_000) return notify('Select a smaller range to copy formatting from.')
      const out: (CellStyle | undefined)[][] = []
      for (let r = rg.r1; r <= rg.r2; r++) {
        const row: (CellStyle | undefined)[] = []
        for (let c = rg.c1; c <= rg.c2; c++) row.push(wb.getStyle(sheetId, r, c))
        out.push(row)
      }
      setPainter(out)
      notify('Format Painter: select the cells to format. Press Esc to cancel.')
    },
    toggle: (prop) =>
      act(() => {
        const on = !anchorStyle?.[prop]
        wb.applyStyle(sheetId, styleRanges(), { [prop]: on || undefined }, prop === 'wrap' ? 'Wrap text' : `${prop[0].toUpperCase()}${prop.slice(1)}`)
      }),
    setStyle: (patch, label) => act(() => wb.applyStyle(sheetId, styleRanges(), patch, label)),
    setAlign: (h, v) =>
      act(() => {
        const patch: Partial<CellStyle> = {}
        if (h) patch.hAlign = anchorStyle?.hAlign === h ? undefined : h
        if (v) patch.vAlign = v === 'bottom' ? undefined : v
        wb.applyStyle(sheetId, styleRanges(), patch, 'Alignment')
      }),
    borders: (preset) => act(() => (selection.extra.length ? notify("Borders can't be applied to multiple selections.") : wb.applyBorders(sheetId, prim, preset))),
    merge: (kind) =>
      act(() => {
        if (kind === 'unmerge') return wb.unmerge(sheetId, prim)
        const rg = prim
        if ((rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1) > 1 && isFullCol(rg)) return 'Select a smaller range to merge.'
        let lost = 0
        for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 500); r++) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 100); c++) if ((r !== rg.r1 || c !== rg.c1) && !wb.isCellEmpty(sheetId, r, c)) lost++
        const run = () => {
          const err = wb.merge(sheetId, rg, kind === 'center')
          if (err) notify(err)
          else setSelection({ anchor: { r: rg.r1, c: rg.c1 }, focus: { r: rg.r2, c: rg.c2 }, extra: [] })
        }
        if (lost) setDialog({ kind: 'confirm', title: 'Merge cells', message: 'Merging keeps only the upper-left value and discards the other values. You can undo this.', confirm: 'Merge', action: run })
        else run()
      }),
    numFmt: (fmt) => act(() => wb.applyStyle(sheetId, styleRanges(), { numFmt: fmt }, 'Number format')),
    decimals: (delta) =>
      act(() => {
        const v = wb.getValue(sheetId, selection.anchor.r, selection.anchor.c)
        const next = adjustDecimals(anchorStyle?.numFmt, delta, typeof v === 'number' ? v : undefined)
        wb.applyStyle(sheetId, styleRanges(), (s, r, c) => {
          const vv = wb.getValue(sheetId, r, c)
          return { ...s, numFmt: s?.numFmt === anchorStyle?.numFmt ? next : adjustDecimals(s?.numFmt, delta, typeof vv === 'number' ? vv : undefined) }
        }, delta > 0 ? 'Increase decimal' : 'Decrease decimal')
      }),
    insert: (kind) =>
      act(() => {
        const rows = Math.min(prim.r2 - prim.r1 + 1, 10_000)
        const cols = Math.min(prim.c2 - prim.c1 + 1, 1_000)
        if (kind === 'rowsAbove') return wb.insertRows(sheetId, prim.r1, isFullCol(prim) ? 1 : rows)
        if (kind === 'rowsBelow') return wb.insertRows(sheetId, prim.r2 + 1, isFullCol(prim) ? 1 : rows)
        if (kind === 'colsLeft') return wb.insertCols(sheetId, prim.c1, isFullRow(prim) ? 1 : cols)
        if (kind === 'colsRight') return wb.insertCols(sheetId, prim.c2 + 1, isFullRow(prim) ? 1 : cols)
        wb.addSheet()
      }),
    remove: (kind) =>
      act(() => {
        if (kind === 'sheet') {
          const d = wb.hf.getSheetDimensions(wb.hfId(sheetId))
          const run = () => {
            const err = wb.removeSheet(sheetId)
            if (err) notify(err)
          }
          if (d.width * d.height > 0) setDialog({ kind: 'confirm', title: 'Delete sheet', message: `Delete “${wb.meta(sheetId).name}”? Formulas that refer to it will show #REF!. You can undo this.`, confirm: 'Delete', danger: true, action: run })
          else run()
          return
        }
        const u = used()
        if (kind === 'rows') {
          const r2 = Math.min(prim.r2, Math.max(prim.r1, u.rows))
          return wb.deleteRows(sheetId, prim.r1, r2 - prim.r1 + 1)
        }
        const c2 = Math.min(prim.c2, Math.max(prim.c1, u.cols))
        return wb.deleteCols(sheetId, prim.c1, c2 - prim.c1 + 1)
      }),
    sizeDialog: (axis) => setDialog({ kind: 'size', axis }),
    autoFit: (axis) => act(() => autoFit(axis)),
    hide: (axis, hidden) =>
      act(() => {
        const [a, b] = axis === 'row' ? [prim.r1, Math.min(prim.r2, Math.max(prim.r1, used().rows + 50))] : [prim.c1, Math.min(prim.c2, Math.max(prim.c1, used().cols + 20))]
        if (hidden) return wb.setHidden(sheetId, axis, range(a, b), true)
        const meta = wb.meta(sheetId)
        const list = (axis === 'row' ? meta.hiddenRows : meta.hiddenCols).filter((i) => i >= a - 1 && i <= b + 1 || (axis === 'row' ? isFullCol(prim) : isFullRow(prim)))
        if (!list.length) return 'No hidden ' + (axis === 'row' ? 'rows' : 'columns') + ' in the selection.'
        wb.setHidden(sheetId, axis, list, false)
      }),
    autoSum: (fn = 'SUM') =>
      act(() => {
        const u = wb.hf.getSheetDimensions(wb.hfId(sheetId))
        const isNum = (r: number, c: number) => r >= 0 && c >= 0 && typeof wb.getValue(sheetId, r, c) === 'number' && !wb.isCellEmpty(sheetId, r, c)
        if (prim.r1 === prim.r2 && prim.c1 === prim.c2) {
          const { r, c } = selection.anchor
          let top = r - 1
          while (top >= 0 && isNum(top, c)) top--
          if (top < r - 1) return startEdit(`=${fn}(${colToLetters(c)}${top + 2}:${colToLetters(c)}${r})`, 'edit')
          let left = c - 1
          while (left >= 0 && isNum(r, left)) left--
          if (left < c - 1) return startEdit(`=${fn}(${colToLetters(left + 1)}${r + 1}:${colToLetters(c - 1)}${r + 1})`, 'edit')
          return startEdit(`=${fn}()`, 'edit')
        }
        const rg = clampContent(prim)
        if (rg.r2 >= u.height + 1) return 'Select the numbers to total.'
        const entries = []
        for (let c = rg.c1; c <= rg.c2; c++) entries.push({ r: rg.r2 + 1, c, input: `=${fn}(${colToLetters(c)}${rg.r1 + 1}:${colToLetters(c)}${rg.r2 + 1})` })
        wb.setInputs(sheetId, entries, 'AutoSum')
      }),
    clear: (kind) =>
      act(() => {
        const rs = ranges
        if (kind === 'contents') return wb.clearContents(sheetId, rs)
        if (kind === 'formats') return wb.clearFormats(sheetId, rs)
        wb.clearAll(sheetId, rs.map(clampContent))
      }),
    find: (replace = false) => setDialog({ kind: 'find', replace }),
    insertFunction: () => setDialog({ kind: 'function' }),
    insertFunctionNamed,
    defineName: () => setDialog({ kind: 'name' }),
    showCompat: () => setDialog({ kind: 'compat' }),
    showShortcuts: () => setDialog({ kind: 'shortcuts' }),
    showAbout: () => setDialog({ kind: 'about' }),
    importFile: () => fileRef.current?.click(),
    exportDialog: () => setDialog({ kind: 'export' }),
    fillDown: () =>
      act(() => {
        const rg = clampContent(prim)
        if (rg.r1 === rg.r2) {
          if (rg.r1 === 0) return
          return wb.fill(sheetId, { ...rg, r1: rg.r1 - 1, r2: rg.r1 - 1 }, { ...rg, r1: rg.r1 - 1 })
        }
        // Ctrl+D copies the first row down (no series)
        const clip = wb.copy(sheetId, { ...rg, r2: rg.r1 })
        wb.paste(clip, sheetId, { ...rg, r1: rg.r1 + 1 })
      }),
    fillRight: () =>
      act(() => {
        const rg = clampContent(prim)
        if (rg.c1 === rg.c2) {
          if (rg.c1 === 0) return
          return wb.fill(sheetId, { ...rg, c1: rg.c1 - 1, c2: rg.c1 - 1 }, { ...rg, c1: rg.c1 - 1 })
        }
        const clip = wb.copy(sheetId, { ...rg, c2: rg.c1 })
        wb.paste(clip, sheetId, { ...rg, c1: rg.c1 + 1 })
      }),
    feature: (x) => featuresRef.current?.run(x),
  }

  // ---------------------------------------------------------------- grid keyboard
  const onGridKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (editing) {
      // keystrokes that arrive before focus reaches the editor are not lost
      if (e.key.length === 1 && !modKey(e) && !e.altKey) {
        e.preventDefault()
        const t = editing.text + e.key
        pendingCaret.current = t.length
        setCaret(t.length)
        setEditing({ ...editing, text: t })
      }
      return
    }
    const mod = modKey(e)
    const k = e.key
    const prevent = () => {
      e.preventDefault()
      e.stopPropagation()
    }
    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }
    if (arrows[k]) {
      prevent()
      const [dr, dc] = arrows[k]
      return move(dr, dc, e.shiftKey, mod)
    }
    if (mod && !e.altKey) {
      const lower = k.toLowerCase()
      if (['c', 'x', 'v'].includes(lower)) return // native clipboard events
      switch (lower) {
        case 'z':
          prevent()
          return e.shiftKey ? cmd.redo() : cmd.undo()
        case 'y':
          prevent()
          return cmd.redo()
        case 'b':
        case 'i':
        case 'u':
          prevent()
          return cmd.toggle(lower === 'b' ? 'bold' : lower === 'i' ? 'italic' : 'underline')
        case '5':
          prevent()
          return cmd.toggle('strike')
        case 'a':
          prevent()
          return setSelection({ anchor: { r: 0, c: 0 }, focus: { r: MAX_ROWS - 1, c: MAX_COLS - 1 }, extra: [] })
        case 's':
          prevent()
          return cmd.save()
        case 'f':
          prevent()
          return cmd.find(false)
        case 'h':
          prevent()
          return cmd.find(true)
        case 'd':
          prevent()
          return cmd.fillDown()
        case 'r':
          prevent()
          return cmd.fillRight()
        case '`':
          prevent()
          return setView({ showFormulas: !view.showFormulas })
        case ' ':
          prevent()
          return setSelection({ anchor: { r: 0, c: selection.anchor.c }, focus: { r: MAX_ROWS - 1, c: selection.focus.c }, extra: [] })
        case 'home':
          prevent()
          return setSelection(e.shiftKey ? { ...selection, focus: { r: 0, c: 0 } } : single({ r: 0, c: 0 }))
        case 'end': {
          prevent()
          const u = used()
          const last = { r: Math.max(0, u.rows - 1), c: Math.max(0, u.cols - 1) }
          return setSelection(e.shiftKey ? { ...selection, focus: last } : single(last))
        }
        case 'pageup':
        case 'pagedown': {
          prevent()
          const vis = wb.sheets.filter((s) => !s.hidden)
          const i = vis.findIndex((s) => s.id === sheetId)
          const next = vis[(i + (lower === 'pageup' ? -1 : 1) + vis.length) % vis.length]
          return wb.setActiveSheet(next.id)
        }
        case 'enter':
          prevent()
          return startEdit(null, 'edit')
        case '1':
          prevent()
          return cmd.feature({ a: 'formatCells' })
        case 'k':
          prevent()
          return cmd.feature({ a: 'hyperlink' })
        case 'e':
          prevent()
          return cmd.feature({ a: 'flashFill' })
        case 'l':
          if (e.shiftKey) {
            prevent()
            return cmd.feature({ a: 'filterToggle' })
          }
          return
        case 't':
          prevent()
          return cmd.feature({ a: 'table' })
        case 'p':
          prevent()
          return cmd.feature({ a: 'printPreview' })
      }
      return
    }
    if (e.altKey && k === 'ArrowDown') {
      prevent()
      const rect = gridRef.current?.cellClientRect(selection.anchor.r, selection.anchor.c)
      if (rect) features.gridProps.onDropdown?.(selection.anchor.r, selection.anchor.c, rect)
      return
    }
    if (e.altKey && (k === '=' || k === '≠')) {
      prevent()
      return cmd.autoSum('SUM')
    }
    switch (k) {
      case 'Tab':
        prevent()
        return move(0, e.shiftKey ? -1 : 1)
      case 'Enter':
        prevent()
        return move(e.shiftKey ? -1 : 1, 0)
      case 'Home':
        prevent()
        return e.shiftKey ? setSelection({ ...selection, focus: { r: selection.focus.r, c: 0 } }) : setSelection(single({ r: selection.anchor.r, c: 0 }))
      case 'PageDown':
      case 'PageUp': {
        prevent()
        const n = gridRef.current?.pageRows() ?? 20
        return move(k === 'PageDown' ? n : -n, 0, e.shiftKey)
      }
      case 'F2':
        prevent()
        if (e.shiftKey) return cmd.feature({ a: 'note' })
        return startEdit(null, 'edit')
      case 'F11':
        if (e.shiftKey) {
          prevent()
          return cmd.insert('sheet')
        }
        return
      case 'Delete':
        prevent()
        return cmd.clear('contents')
      case 'Backspace':
        prevent()
        return startEdit('', 'enter')
      case 'Escape':
        prevent()
        setPainter(null)
        clipRef.current = null
        return setCopyRange(null)
      case ' ':
        if (e.shiftKey) {
          prevent()
          return setSelection({ anchor: { r: selection.anchor.r, c: 0 }, focus: { r: selection.focus.r, c: MAX_COLS - 1 }, extra: [] })
        }
    }
    if (k.length === 1 && !mod && !e.altKey) {
      prevent()
      startEdit(k, 'enter')
    }
  }

  // ---------------------------------------------------------------- navigation from name box
  const navigate = (target: string): string | null => {
    if (!target) return null
    let sheet = sheetId
    let ref = target
    const bang = target.lastIndexOf('!')
    if (bang > 0) {
      const name = target.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'")
      const s = wb.sheetByName(name)
      if (!s) return `There is no sheet named “${name}”.`
      sheet = s.id
      ref = target.slice(bang + 1)
    }
    let rg = parseRangeA1(ref)
    if (!rg) {
      const named = wb.namedRanges().find((n) => n.name.toLowerCase() === target.toLowerCase())
      if (named) {
        const ex = named.expression.replace(/^=/, '')
        const b2 = ex.lastIndexOf('!')
        if (b2 > 0) {
          const s = wb.sheetByName(ex.slice(0, b2).replace(/^'(.*)'$/, '$1'))
          if (s) sheet = s.id
        }
        rg = parseRangeA1(b2 > 0 ? ex.slice(b2 + 1) : ex)
      }
    }
    if (!rg) return `“${target}” is not a cell reference, range or defined name.`
    if (sheet !== wb.activeSheetId) wb.setActiveSheet(sheet)
    const sel = { anchor: { r: rg.r1, c: rg.c1 }, focus: { r: rg.r2, c: rg.c2 }, extra: [] }
    selections.current.set(sheet, sel)
    setSelectionState(sel)
    focusGrid()
    return null
  }

  const features = useFeatures({
    wb,
    sheetId,
    version,
    selection,
    setSelection,
    ranges,
    prim,
    anchorStyle,
    notify,
    dark: view.dark,
    editing: !!editing,
    focusGrid,
    navigate,
    openDefineName: () => setDialog({ kind: 'name' }),
  })
  featuresRef.current = features

  // ---------------------------------------------------------------- export
  const doExport = async (c: ExportChoice, fileName: string) => {
    setDialog(null)
    try {
      if (c.format === 'xlsx') {
        const bytes = await exportXlsx(wb)
        await download(bytes, safeFileName(fileName), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
      } else {
        const text = exportDelimited(wb, c.sheetId, c.delimiter, { bom: c.bom })
        await download(text, safeFileName(fileName), c.delimiter === '\t' ? 'text/tab-separated-values' : 'text/csv')
      }
      notify(`Downloaded ${fileName}`)
    } catch (e) {
      notify(`Export failed: ${e instanceof Error ? e.message : String(e)}`)
    }
  }

  // ---------------------------------------------------------------- format painter apply on selection
  const applyPainter = () => {
    if (!painter) return
    const h = painter.length
    const w = painter[0]?.length ?? 1
    act(() =>
      wb.applyStyle(sheetId, [clampContent(prim).r2 - prim.r1 < h ? { ...prim, r2: prim.r1 + h - 1, c2: Math.max(prim.c2, prim.c1 + w - 1) } : prim], (_s, r, c) => painter[(r - prim.r1) % h][(c - prim.c1) % w], 'Format Painter'),
    )
    setPainter(null)
  }

  // ---------------------------------------------------------------- render
  const status = saveState.status
  const statusEl = (
    <span className={cn('flex items-center gap-1 text-[12px]', status === 'error' || status === 'conflict' ? 'text-[var(--danger)]' : 'text-[var(--muted)]')} data-testid="save-status" title={saveState.error ?? (saveState.lastSavedAt ? `Saved ${saveState.lastSavedAt.toLocaleTimeString()} to this browser` : undefined)}>
      <Icon name={status === 'saved' ? 'check' : status === 'saving' ? 'save' : status === 'error' || status === 'conflict' ? 'alert' : 'edit'} size={13} />
      {status === 'saved' ? 'Saved' : status === 'saving' ? 'Saving…' : status === 'unsaved' ? 'Unsaved changes' : status === 'conflict' ? 'Save conflict' : "Couldn't save"}
    </span>
  )

  const mode = editing ? (canPointRef() ? 'Point' : editing.mode === 'enter' ? 'Enter' : 'Edit') : painter ? 'Format Painter' : 'Ready'

  return (
    <div className="flex h-full flex-col bg-[var(--bg)]" onPointerUp={() => painter && setTimeout(applyPainter, 0)}>
      {/* header */}
      <header className="flex h-11 shrink-0 items-center gap-2 border-b border-[var(--border)] bg-[var(--panel)] px-2">
        <Tooltip text="All workbooks">
          <button type="button" onClick={onHome} className="flex items-center gap-1.5 rounded-md px-1.5 py-1 hover:bg-[var(--hover)]" aria-label="Go to file list">
            <Logo />
            <span className="hidden text-[14px] font-semibold tracking-tight sm:inline">ExcelPro</span>
          </button>
        </Tooltip>
        <div className="mx-1 hidden h-5 w-px bg-[var(--border)] sm:block" />
        <input
          aria-label="Workbook name"
          data-testid="workbook-name"
          value={nameDraft}
          onChange={(e) => setNameDraft(e.target.value)}
          onBlur={() => (nameDraft.trim() ? wb.setName(nameDraft) : setNameDraft(wb.name))}
          onKeyDown={(e) => {
            if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
            if (e.key === 'Escape') {
              setNameDraft(wb.name)
              ;(e.target as HTMLInputElement).blur()
            }
          }}
          className="h-7 w-32 min-w-0 rounded-md border border-transparent bg-transparent px-1.5 text-[14px] font-medium outline-none hover:border-[var(--border)] focus:border-[var(--accent)] focus:bg-[var(--bg)] sm:w-56"
        />
        {statusEl}
        <div className="ml-auto flex items-center gap-0.5">
          <Button size="icon" tip={`Undo${wb.undoLabel() ? ` ${wb.undoLabel()}` : ''} (${shortcut('Ctrl+Z')})`} disabled={!wb.canUndo()} onClick={cmd.undo} aria-label="Undo">
            <Icon name="undo" />
          </Button>
          <Button size="icon" tip={`Redo${wb.redoLabel() ? ` ${wb.redoLabel()}` : ''} (${shortcut('Ctrl+Y')})`} disabled={!wb.canRedo()} onClick={cmd.redo} aria-label="Redo">
            <Icon name="redo" />
          </Button>
          <Button size="icon" tip={`Find (${shortcut('Ctrl+F')})`} onClick={() => cmd.find(false)} aria-label="Search">
            <Icon name="search" />
          </Button>
          <Button size="sm" variant="outline" tip="Download as .xlsx, .csv or .tsv" onClick={cmd.exportDialog} className="ml-1">
            <Icon name="download" />
            <span className="hidden md:inline">Download</span>
          </Button>
          <Popover
            align="end"
            trigger={({ toggle }) => (
              <Button size="icon" tip="Settings" onClick={toggle} aria-label="Settings">
                <Icon name="settings" />
              </Button>
            )}
          >
            <MenuLabel>Appearance</MenuLabel>
            <MenuItem checked={!view.dark} onSelect={() => setView({ dark: false })}>
              Light theme
            </MenuItem>
            <MenuItem checked={view.dark} onSelect={() => setView({ dark: true })}>
              Dark theme
            </MenuItem>
            <MenuSep />
            <MenuLabel>Typing dates like 03/04/2025</MenuLabel>
            <MenuItem checked={dateOrder === 'MDY'} onSelect={() => setDateOrder('MDY')}>
              Month/Day/Year (US)
            </MenuItem>
            <MenuItem checked={dateOrder === 'DMY'} onSelect={() => setDateOrder('DMY')}>
              Day/Month/Year
            </MenuItem>
            <MenuSep />
            <MenuItem icon="keyboard" onSelect={cmd.showShortcuts}>
              Keyboard shortcuts
            </MenuItem>
            <MenuItem icon="help" onSelect={cmd.showAbout}>
              About & help
            </MenuItem>
          </Popover>
          <Popover
            align="end"
            trigger={({ toggle }) => (
              <button type="button" onClick={toggle} aria-label="Profile" className="ml-1 flex h-7 w-7 items-center justify-center rounded-full bg-[var(--accent)] text-[12px] font-semibold text-[var(--accent-contrast)]">
                <Icon name="user" size={15} />
              </button>
            )}
          >
            <div className="max-w-64 p-2 text-[12.5px]">
              <div className="mb-1 flex items-center gap-2 font-semibold">
                <Icon name="cloudOff" size={15} /> Local profile
              </div>
              <p className="text-[var(--muted)]">Workbooks are saved in this browser on this device. Accounts, cloud storage and sharing arrive in Phase 4.</p>
            </div>
          </Popover>
        </div>
      </header>

      {saveState.status === 'conflict' && (
        <div className="flex shrink-0 flex-wrap items-center gap-2 bg-[var(--warn-bg)] px-3 py-1.5 text-[12.5px] text-[var(--warn-text)]" role="alert">
          <Icon name="alert" size={15} />
          <span>This workbook was changed in another tab or window. Your latest edits here are not saved yet.</span>
          <span className="ml-auto flex gap-2">
            <Button size="sm" variant="outline" onClick={onReloadFromStore}>
              Load the other version
            </Button>
            <Button size="sm" variant="default" onClick={() => void autosave.overwrite(saveState.conflictRevision ?? 0)}>
              Keep this version
            </Button>
          </span>
        </div>
      )}
      {saveState.status === 'error' && saveState.error && (
        <div className="flex shrink-0 items-center gap-2 bg-[var(--warn-bg)] px-3 py-1.5 text-[12.5px] text-[var(--warn-text)]" role="alert">
          <Icon name="alert" size={15} />
          {saveState.error}
          <Button size="sm" variant="outline" className="ml-auto" onClick={() => void autosave.flush()}>
            Retry now
          </Button>
        </div>
      )}

      <Ribbon cmd={cmd} style={anchorStyle} view={view} setView={setView} painterActive={!!painter} hasClip={!!clipRef.current} ctx={features.ctx} />

      {view.showFormulaBar && (
        <FormulaBar
          address={address}
          text={barText}
          editing={!!editing}
          inputRef={barRef}
          onNavigate={navigate}
          onFocusInput={() => {
            if (!editing) startEdit(null, 'edit', 'bar')
            else if (editing.origin !== 'bar') setEditing({ ...editing, origin: 'bar' })
          }}
          onChange={(t, pos) => onEditText(t, pos)}
          onKeyDown={onEditKeyDown}
          onCancel={cancelEdit}
          onCommit={() => commitEdit(0, 0)}
          onInsertFunction={cmd.insertFunction}
        />
      )}

      <div className="flex min-h-0 flex-1">
      <div className="relative min-h-0 min-w-0 flex-1" onMouseDown={(e) => editing && (e.target as HTMLElement).closest('.xp-scroller') && e.preventDefault()}>
        <Grid
          {...features.gridProps}
          ref={gridRef}
          wb={wb}
          sheetId={sheetId}
          version={version}
          selection={selection}
          onSelect={setSelection}
          editing={editing}
          onEditText={onEditText}
          onEditKeyDown={onEditKeyDown}
          onStartEdit={(init, mode) => startEdit(init, mode)}
          onInsertRef={onInsertRef}
          pointRef={pointRef}
          canPointRef={canPointRef}
          zoom={view.zoom}
          showGrid={view.showGrid && wb.meta(sheetId).showGridLines !== false}
          showHeadings={view.showHeadings}
          showFormulas={view.showFormulas}
          copyRange={copyRange && copyRange.sheetId === sheetId ? copyRange.range : null}
          onFill={(src, dest) => act(() => {
            wb.fill(sheetId, src, dest)
            setSelection({ anchor: { r: dest.r1, c: dest.c1 }, focus: { r: dest.r2, c: dest.c2 }, extra: [] })
          })}
          onContextMenu={(pos, kind) => setCtx({ ...pos, kind })}
          onResize={(axis, idx, px) => act(() => (axis === 'col' ? wb.setColWidths(sheetId, idx, px) : wb.setRowHeights(sheetId, idx, px)))}
          onAutoFit={(axis, idx) => act(() => autoFit(axis, idx))}
          onKeyDown={onGridKeyDown}
          onCopy={onCopyEvent}
          onCut={onCutEvent}
          onPaste={onPasteEvent}
          editorRef={editorRef}
        />
      </div>
      {features.side}
      </div>

      <SheetTabs wb={wb} onError={notify} onChange={() => setTimeout(focusGrid, 0)} />
      <StatusBar
        wb={wb}
        sheetId={sheetId}
        ranges={ranges}
        version={version}
        message={errorMsg ?? (painter ? 'Select cells to apply the copied formatting.' : null)}
        zoom={view.zoom}
        setZoom={(z) => setView({ zoom: z })}
        mode={mode}
      />

      <input
        ref={fileRef}
        type="file"
        className="hidden"
        accept=".xlsx,.xlsm,.csv,.tsv,.txt,.xls,.ods"
        onChange={(e) => {
          const f = e.target.files?.[0]
          e.target.value = ''
          if (f) onImportFile(f)
        }}
      />

      {ctx && (
        <ContextMenu x={ctx.x} y={ctx.y} onClose={() => setCtx(null)}>
          <MenuItem icon="cut" shortcut={shortcut('Ctrl+X')} onSelect={cmd.cut}>Cut</MenuItem>
          <MenuItem icon="copy" shortcut={shortcut('Ctrl+C')} onSelect={cmd.copy}>Copy</MenuItem>
          <MenuItem icon="paste" shortcut={shortcut('Ctrl+V')} onSelect={() => cmd.paste()}>Paste</MenuItem>
          <MenuItem disabled={!clipRef.current} onSelect={() => cmd.paste('values')}>Paste values</MenuItem>
          <MenuSep />
          {ctx.kind !== 'col' && (
            <>
              <MenuItem icon="rowsInsert" onSelect={() => cmd.insert('rowsAbove')}>Insert {Math.min(prim.r2 - prim.r1 + 1, 10000) > 1 && !isFullCol(prim) ? `${prim.r2 - prim.r1 + 1} rows` : 'row'} above</MenuItem>
              <MenuItem onSelect={() => cmd.remove('rows')}>Delete {ctx.kind === 'row' ? 'rows' : 'row(s)'}</MenuItem>
            </>
          )}
          {ctx.kind !== 'row' && (
            <>
              <MenuItem icon="colsInsert" onSelect={() => cmd.insert('colsLeft')}>Insert column{prim.c2 > prim.c1 && !isFullRow(prim) ? 's' : ''} left</MenuItem>
              <MenuItem onSelect={() => cmd.remove('cols')}>Delete column(s)</MenuItem>
            </>
          )}
          <MenuSep />
          {ctx.kind === 'col' && (
            <>
              <MenuItem onSelect={() => cmd.sizeDialog('col')}>Column width…</MenuItem>
              <MenuItem onSelect={() => cmd.autoFit('col')}>AutoFit width</MenuItem>
              <MenuItem onSelect={() => cmd.hide('col', true)}>Hide</MenuItem>
              <MenuItem onSelect={() => cmd.hide('col', false)}>Unhide</MenuItem>
              <MenuSep />
            </>
          )}
          {ctx.kind === 'row' && (
            <>
              <MenuItem onSelect={() => cmd.sizeDialog('row')}>Row height…</MenuItem>
              <MenuItem onSelect={() => cmd.autoFit('row')}>AutoFit height</MenuItem>
              <MenuItem onSelect={() => cmd.hide('row', true)}>Hide</MenuItem>
              <MenuItem onSelect={() => cmd.hide('row', false)}>Unhide</MenuItem>
              <MenuSep />
            </>
          )}
          {ctx.kind === 'cell' && (
            <>
              <MenuItem icon="sortAsc" onSelect={() => cmd.feature({ a: 'sort', desc: false })}>Sort A to Z</MenuItem>
              <MenuItem icon="sortDesc" onSelect={() => cmd.feature({ a: 'sort', desc: true })}>Sort Z to A</MenuItem>
              <MenuItem icon="filter" shortcut={shortcut('Ctrl+Shift+L')} onSelect={() => cmd.feature({ a: 'filterToggle' })}>{features.ctx.hasFilter ? 'Remove filter' : 'Filter'}</MenuItem>
              <MenuSep />
              <MenuItem icon="note" shortcut="Shift+F2" onSelect={() => cmd.feature({ a: 'note' })}>{features.ctx.hasNote ? 'Edit note' : 'New note'}</MenuItem>
              {features.ctx.hasNote && <MenuItem onSelect={() => cmd.feature({ a: 'deleteNote' })}>Delete note</MenuItem>}
              <MenuItem icon="link" shortcut={shortcut('Ctrl+K')} onSelect={() => cmd.feature({ a: 'hyperlink' })}>Link…</MenuItem>
              <MenuItem icon="palette" shortcut={shortcut('Ctrl+1')} onSelect={() => cmd.feature({ a: 'formatCells' })}>Format cells…</MenuItem>
              <MenuSep />
            </>
          )}
          <MenuItem icon="eraser" shortcut="Del" onSelect={() => cmd.clear('contents')}>Clear contents</MenuItem>
          <MenuItem onSelect={() => cmd.clear('formats')}>Clear formats</MenuItem>
          {ctx.kind === 'cell' && (
            <>
              <MenuSep />
              <MenuItem icon="merge" onSelect={() => cmd.merge('center')}>Merge &amp; Center</MenuItem>
              <MenuItem onSelect={() => cmd.merge('unmerge')}>Unmerge</MenuItem>
              <MenuItem icon="tag" onSelect={cmd.defineName}>Define name…</MenuItem>
            </>
          )}
        </ContextMenu>
      )}

      {dialog?.kind === 'find' && (
        <FindReplaceDialog
          wb={wb}
          initialReplace={dialog.replace}
          activeSheetId={sheetId}
          activeCell={selection.anchor}
          onClose={() => (setDialog(null), gridRef.current?.focus())}
          onChanged={() => setVersion(wb.version)}
          onGoto={(m) => {
            if (m.sheetId !== wb.activeSheetId) wb.setActiveSheet(m.sheetId)
            const sel = single({ r: m.r, c: m.c })
            selections.current.set(m.sheetId, sel)
            setSelectionState(sel)
          }}
        />
      )}
      {dialog?.kind === 'export' && <ExportDialog wb={wb} onClose={() => setDialog(null)} onExport={(c, n) => void doExport(c, n)} />}
      {dialog?.kind === 'function' && (
        <InsertFunctionDialog
          onClose={() => setDialog(null)}
          onInsert={(name) => {
            setDialog(null)
            rememberFn(name)
            insertFunctionNamed(name)
          }}
        />
      )}
      {dialog?.kind === 'size' && (
        <SizeDialog
          axis={dialog.axis}
          initial={dialog.axis === 'row' ? wb.meta(sheetId).rowHeights[selection.anchor.r] ?? DEFAULT_ROW_HEIGHT : wb.meta(sheetId).colWidths[selection.anchor.c] ?? DEFAULT_COL_WIDTH}
          onClose={() => setDialog(null)}
          onApply={(px) => {
            setDialog(null)
            const u = used()
            act(() =>
              dialog.axis === 'row'
                ? wb.setRowHeights(sheetId, range(prim.r1, Math.min(prim.r2, Math.max(prim.r1, u.rows + 100))), px)
                : wb.setColWidths(sheetId, range(prim.c1, Math.min(prim.c2, Math.max(prim.c1, u.cols + 26))), px),
            )
          }}
        />
      )}
      {dialog?.kind === 'name' && (
        <DefineNameDialog
          wb={wb}
          refersTo={`=${quoteSheetName(wb.meta(sheetId).name)}!${rangeToA1(clampContent(prim)).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`}
          onClose={() => setDialog(null)}
          onDone={(m) => (setDialog(null), notify(m))}
        />
      )}
      {dialog?.kind === 'compat' && <CompatDialog onClose={() => setDialog(null)} />}
      {dialog?.kind === 'shortcuts' && <ShortcutsDialog onClose={() => setDialog(null)} />}
      {dialog?.kind === 'about' && <AboutDialog onClose={() => setDialog(null)} />}
      {dialog?.kind === 'report' && <ImportReportDialog report={dialog.report} fileName={dialog.fileName} onClose={() => (setDialog(null), gridRef.current?.focus())} />}
      {dialog?.kind === 'confirm' && (
        <ConfirmDialog
          title={dialog.title}
          message={dialog.message}
          confirm={dialog.confirm}
          danger={dialog.danger}
          onClose={() => setDialog(null)}
          onConfirm={() => {
            setDialog(null)
            act(dialog.action)
          }}
        />
      )}

      {features.ui}
      {editing && assist && (() => {
        const el = editing.origin === 'cell' ? editorRef.current : barRef.current
        const rect = el?.getBoundingClientRect()
        return rect ? <FormulaAssistPanel anchor={rect} state={assist} active={Math.min(assistIdx, Math.max(0, assist.suggestions.length - 1))} onHover={setAssistIdx} onPick={(sg) => acceptAssist(assist.suggestions.indexOf(sg))} /> : null
      })()}

      {toast && (
        <div role="status" className="fixed bottom-10 left-1/2 z-[70] max-w-[90vw] -translate-x-1/2 rounded-lg bg-[#1f2328] px-4 py-2 text-[12.5px] text-white shadow-xl dark:bg-[#e7e9ec] dark:text-[#1f2328]">
          {toast}
        </div>
      )}
    </div>
  )
}

function range(a: number, b: number): number[] {
  const out: number[] = []
  for (let i = a; i <= b; i++) out.push(i)
  return out
}

export { formatGeneral }
