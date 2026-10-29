import { useMemo, useState } from 'react'
import type { Workbook } from '../engine/workbook'
import { compatibilityMatrix, compatOf, engineFunctions, FUNCTION_HELP, SPEC_FUNCTIONS } from '../engine/compat'
import type { ImportReport } from '../files/xlsx'
import { addrToA1 } from '../model/address'
import { Button, Check, Dialog, Field, inputClass } from './components/ui'
import { Icon } from './icons'
import { cn, shortcut } from './lib/utils'

// ------------------------------------------------------------------ Find & Replace
export interface FindMatch {
  sheetId: string
  r: number
  c: number
}

export function FindReplaceDialog({
  wb,
  initialReplace,
  onClose,
  onGoto,
  activeSheetId,
  activeCell,
  onChanged,
}: {
  wb: Workbook
  initialReplace: boolean
  onClose: () => void
  onGoto: (m: FindMatch) => void
  activeSheetId: string
  activeCell: { r: number; c: number }
  onChanged: () => void
}) {
  const [q, setQ] = useState('')
  const [rep, setRep] = useState('')
  const [showReplace, setShowReplace] = useState(initialReplace)
  const [scope, setScope] = useState<'sheet' | 'workbook'>('sheet')
  const [lookIn, setLookIn] = useState<'values' | 'formulas'>('values')
  const [matchCase, setMatchCase] = useState(false)
  const [entire, setEntire] = useState(false)
  const [results, setResults] = useState<FindMatch[] | null>(null)
  const [info, setInfo] = useState<string | null>(null)

  const search = () => {
    const ids = scope === 'sheet' ? [activeSheetId] : wb.order
    return wb.find(q, { sheetIds: ids, matchCase, entireCell: entire, lookIn: showReplace ? 'formulas' : lookIn })
  }
  const next = () => {
    const all = search()
    setResults(all)
    if (!all.length) return setInfo(`No matches for “${q}”.`)
    const order = (m: FindMatch) => [wb.order.indexOf(m.sheetId), m.r, m.c]
    const cur = [wb.order.indexOf(activeSheetId), activeCell.r, activeCell.c]
    const after = all.find((m) => {
      const o = order(m)
      return o[0] > cur[0] || (o[0] === cur[0] && (o[1] > cur[1] || (o[1] === cur[1] && o[2] > cur[2])))
    })
    const target = after ?? all[0]
    setInfo(`${all.indexOf(target) + 1} of ${all.length}`)
    onGoto(target)
  }
  const replaceOne = () => {
    const all = search()
    const here = all.find((m) => m.sheetId === activeSheetId && m.r === activeCell.r && m.c === activeCell.c)
    if (here) {
      wb.replace([here], q, rep, matchCase)
      onChanged()
    }
    next()
  }
  const replaceAll = () => {
    const all = search()
    const n = wb.replace(all, q, rep, matchCase)
    setInfo(n ? `Replaced ${n} cell${n === 1 ? '' : 's'}. Undo with ${shortcut('Ctrl+Z')}.` : `No matches for “${q}”.`)
    setResults([])
    onChanged()
  }

  return (
    <Dialog title={showReplace ? 'Find and replace' : 'Find'} onClose={onClose} width={520}>
      <div className="mb-2 flex gap-1">
        <Button size="sm" variant={!showReplace ? 'outline' : 'ghost'} onClick={() => setShowReplace(false)}>
          Find
        </Button>
        <Button size="sm" variant={showReplace ? 'outline' : 'ghost'} onClick={() => setShowReplace(true)}>
          Replace
        </Button>
      </div>
      <Field label="Find what">
        <input className={inputClass} value={q} onChange={(e) => setQ(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && q && next()} aria-label="Find what" />
      </Field>
      {showReplace && (
        <Field label="Replace with">
          <input className={inputClass} value={rep} onChange={(e) => setRep(e.target.value)} aria-label="Replace with" />
        </Field>
      )}
      <div className="mb-2 flex flex-wrap items-center gap-x-3 gap-y-1">
        <label className="flex items-center gap-1.5 text-[12.5px]">
          Within
          <select className={cn(inputClass, 'h-7')} value={scope} onChange={(e) => setScope(e.target.value as 'sheet' | 'workbook')}>
            <option value="sheet">Sheet</option>
            <option value="workbook">Workbook</option>
          </select>
        </label>
        {!showReplace && (
          <label className="flex items-center gap-1.5 text-[12.5px]">
            Look in
            <select className={cn(inputClass, 'h-7')} value={lookIn} onChange={(e) => setLookIn(e.target.value as 'values' | 'formulas')}>
              <option value="values">Values</option>
              <option value="formulas">Formulas</option>
            </select>
          </label>
        )}
        <Check label="Match case" checked={matchCase} onChange={setMatchCase} />
        <Check label="Entire cell" checked={entire} onChange={setEntire} />
      </div>
      <div className="flex flex-wrap gap-2">
        <Button variant="outline" size="sm" disabled={!q} onClick={() => setResults(search())}>
          Find all
        </Button>
        <Button variant="default" size="sm" disabled={!q} onClick={next}>
          Find next
        </Button>
        {showReplace && (
          <>
            <Button variant="outline" size="sm" disabled={!q} onClick={replaceOne}>
              Replace
            </Button>
            <Button variant="outline" size="sm" disabled={!q} onClick={replaceAll}>
              Replace all
            </Button>
          </>
        )}
        {info && <span className="self-center text-[12px] text-[var(--muted)]">{info}</span>}
      </div>
      {results && results.length > 0 && (
        <div className="mt-3 max-h-56 overflow-auto rounded-md border border-[var(--border)]">
          <table className="w-full text-[12px]">
            <thead className="sticky top-0 bg-[var(--panel)] text-left text-[var(--muted)]">
              <tr>
                <th className="px-2 py-1 font-medium">Sheet</th>
                <th className="px-2 py-1 font-medium">Cell</th>
                <th className="px-2 py-1 font-medium">Value</th>
              </tr>
            </thead>
            <tbody>
              {results.slice(0, 500).map((m) => (
                <tr key={`${m.sheetId}${m.r},${m.c}`} className="cursor-default hover:bg-[var(--hover)]" onClick={() => onGoto(m)}>
                  <td className="px-2 py-1">{wb.meta(m.sheetId).name}</td>
                  <td className="px-2 py-1 font-mono">{addrToA1(m)}</td>
                  <td className="truncate px-2 py-1">{wb.getDisplay(m.sheetId, m.r, m.c).text}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {results.length > 500 && <div className="p-2 text-[11.5px] text-[var(--muted)]">Showing 500 of {results.length} matches.</div>}
        </div>
      )}
      {results && results.length === 0 && !info && <p className="mt-3 text-[12.5px] text-[var(--muted)]">No matches.</p>}
    </Dialog>
  )
}

// ------------------------------------------------------------------ Export
export type ExportChoice = { format: 'xlsx' } | { format: 'csv'; sheetId: string; delimiter: ',' | ';' | '\t' | '|'; bom: boolean }

export function ExportDialog({ wb, onClose, onExport }: { wb: Workbook; onClose: () => void; onExport: (c: ExportChoice, fileName: string) => void }) {
  const [format, setFormat] = useState<'xlsx' | 'csv' | 'tsv'>('xlsx')
  const [sheetId, setSheetId] = useState(wb.activeSheetId)
  const [delimiter, setDelimiter] = useState<',' | ';' | '|'>(',')
  const [bom, setBom] = useState(true)
  const [name, setName] = useState(wb.name)
  const ext = format
  return (
    <Dialog
      title="Download a copy"
      onClose={onClose}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="default"
            onClick={() => onExport(format === 'xlsx' ? { format: 'xlsx' } : { format: 'csv', sheetId, delimiter: format === 'tsv' ? '\t' : delimiter, bom }, `${name || 'Workbook'}.${ext}`)}
          >
            <Icon name="download" /> Download
          </Button>
        </>
      }
    >
      <Field label="File name">
        <div className="flex items-center gap-1">
          <input className={cn(inputClass, 'flex-1')} value={name} onChange={(e) => setName(e.target.value)} aria-label="File name" />
          <span className="text-[12.5px] text-[var(--muted)]">.{ext}</span>
        </div>
      </Field>
      <div className="mb-3 grid gap-2">
        {(
          [
            ['xlsx', 'Excel workbook (.xlsx)', 'All sheets, formulas, formatting, merges, sizes and named ranges.'],
            ['csv', 'CSV (.csv)', 'One sheet, displayed values only.'],
            ['tsv', 'Tab-separated (.tsv)', 'One sheet, displayed values only.'],
          ] as const
        ).map(([id, label, hint]) => (
          <label key={id} className={cn('flex cursor-default gap-2 rounded-lg border p-2.5', format === id ? 'border-[var(--accent)] bg-[var(--active)]/40' : 'border-[var(--border)] hover:bg-[var(--hover)]')}>
            <input type="radio" name="fmt" checked={format === id} onChange={() => setFormat(id)} className="mt-0.5 accent-[var(--accent)]" />
            <span>
              <span className="block text-[13px] font-medium">{label}</span>
              <span className="block text-[12px] text-[var(--muted)]">{hint}</span>
            </span>
          </label>
        ))}
      </div>
      {format !== 'xlsx' && (
        <div className="grid grid-cols-2 gap-3">
          <Field label="Sheet">
            <select className={inputClass} value={sheetId} onChange={(e) => setSheetId(e.target.value)}>
              {wb.sheets.map((s) => (
                <option key={s.id} value={s.id}>
                  {s.name}
                </option>
              ))}
            </select>
          </Field>
          {format === 'csv' && (
            <Field label="Delimiter">
              <select className={inputClass} value={delimiter} onChange={(e) => setDelimiter(e.target.value as ',' | ';' | '|')}>
                <option value=",">Comma ( , )</option>
                <option value=";">Semicolon ( ; )</option>
                <option value="|">Pipe ( | )</option>
              </select>
            </Field>
          )}
          <div className="col-span-2">
            <Check label="Add UTF-8 byte-order mark (helps Excel open non-English text)" checked={bom} onChange={setBom} />
          </div>
        </div>
      )}
      <p className="mt-1 text-[11.5px] text-[var(--muted)]">PDF export and print layouts arrive in Phase 3.</p>
    </Dialog>
  )
}

// ------------------------------------------------------------------ Insert function
export function InsertFunctionDialog({ onClose, onInsert }: { onClose: () => void; onInsert: (name: string) => void }) {
  const [q, setQ] = useState('')
  const [cat, setCat] = useState('Common')
  const all = useMemo(() => [...engineFunctions()].sort(), [])
  const categories = ['Common', ...Object.keys(SPEC_FUNCTIONS), 'All supported']
  const list = useMemo(() => {
    const ql = q.trim().toLowerCase()
    let base: string[]
    if (ql) {
      const words = ql.split(/\s+/).filter((w) => w.length > 2 && !['the', 'and', 'for', 'find', 'calculate'].includes(w))
      const score = (n: string) => {
        const h = FUNCTION_HELP[n]
        const hay = `${n} ${h?.description ?? ''} ${h?.category ?? ''}`.toLowerCase()
        let sc = n.toLowerCase().startsWith(ql) ? 100 : n.toLowerCase().includes(ql) ? 50 : 0
        for (const w of words) if (hay.includes(w)) sc += 10
        return sc
      }
      const scored = all.map((n) => [n, score(n)] as const).filter(([, sc]) => sc > 0)
      scored.sort((a, b) => b[1] - a[1] || Number(!FUNCTION_HELP[a[0]]) - Number(!FUNCTION_HELP[b[0]]) || a[0].localeCompare(b[0]))
      base = scored.map(([n]) => n)
    } else if (cat === 'Common') base = Object.keys(FUNCTION_HELP)
    else if (cat === 'All supported') base = all
    else base = SPEC_FUNCTIONS[cat] ?? []
    return base
  }, [q, cat, all])
  const [sel, setSel] = useState<string | null>(null)
  const current = sel ?? list[0] ?? null
  const help = current ? FUNCTION_HELP[current] : undefined
  const compat = current ? compatOf(current) : undefined
  return (
    <Dialog
      title="Insert function"
      onClose={onClose}
      width={640}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button variant="default" disabled={!current || compat?.status === 'unsupported'} onClick={() => current && onInsert(current)}>
            Insert
          </Button>
        </>
      }
    >
      <div className="mb-2 flex gap-2">
        <input
          className={cn(inputClass, 'flex-1')}
          placeholder='Search by name or purpose, e.g. "average" or "matching employee ID"'
          value={q}
          onChange={(e) => {
            setQ(e.target.value)
            setSel(null)
          }}
          aria-label="Search functions"
        />
        <select className={inputClass} value={cat} onChange={(e) => (setCat(e.target.value), setSel(null), setQ(''))} aria-label="Category">
          {categories.map((c) => (
            <option key={c}>{c}</option>
          ))}
        </select>
      </div>
      <div className="grid h-72 grid-cols-[180px_1fr] gap-3">
        <div className="overflow-auto rounded-md border border-[var(--border)]" role="listbox" aria-label="Functions">
          {list.map((n) => {
            const st = compatOf(n).status
            return (
              <button
                key={n}
                type="button"
                role="option"
                aria-selected={current === n}
                onClick={() => setSel(n)}
                onDoubleClick={() => st !== 'unsupported' && onInsert(n)}
                className={cn('flex w-full items-center justify-between px-2 py-1 text-left font-mono text-[12px] hover:bg-[var(--hover)]', current === n && 'bg-[var(--active)]')}
              >
                {n}
                {st !== 'supported' && <span className={cn('text-[10px]', st === 'partial' ? 'text-amber-600' : 'text-[var(--danger)]')}>{st === 'partial' ? 'partial' : 'n/a'}</span>}
              </button>
            )
          })}
          {list.length === 0 && <div className="p-2 text-[12px] text-[var(--muted)]">No matching functions.</div>}
        </div>
        <div className="overflow-auto text-[13px]">
          {current && (
            <>
              <div className="font-mono text-[13px] font-semibold">{help?.syntax ?? `${current}(…)`}</div>
              <p className="mt-1 text-[var(--muted)]">{help?.description ?? 'Supported by the formula engine.'}</p>
              {compat && compat.status !== 'supported' && (
                <p className={cn('mt-2 rounded-md p-2 text-[12px]', compat.status === 'partial' ? 'bg-[var(--warn-bg)] text-[var(--warn-text)]' : 'bg-red-500/10 text-[var(--danger)]')}>
                  {compat.status === 'partial' ? 'Partial support: ' : 'Not supported: '}
                  {compat.note}
                </p>
              )}
              {compat?.status === 'supported' && compat.note && <p className="mt-2 text-[12px] text-[var(--muted)]">{compat.note}</p>}
            </>
          )}
        </div>
      </div>
    </Dialog>
  )
}

// ------------------------------------------------------------------ size
export function SizeDialog({ axis, initial, onClose, onApply }: { axis: 'row' | 'col'; initial: number; onClose: () => void; onApply: (px: number) => void }) {
  const [v, setV] = useState(String(initial))
  const n = Number(v)
  const ok = Number.isFinite(n) && n >= 0 && n <= 2000
  return (
    <Dialog
      title={axis === 'row' ? 'Row height' : 'Column width'}
      onClose={onClose}
      width={320}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button variant="default" disabled={!ok} onClick={() => onApply(n)}>
            OK
          </Button>
        </>
      }
    >
      <Field label={`${axis === 'row' ? 'Row height' : 'Column width'} (pixels)`} hint={axis === 'row' ? 'Default is 22 px.' : 'Default is 88 px.'}>
        <input className={inputClass} value={v} onChange={(e) => setV(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && ok && onApply(n)} inputMode="numeric" />
      </Field>
    </Dialog>
  )
}

// ------------------------------------------------------------------ define name
export function DefineNameDialog({ wb, refersTo, onClose, onDone }: { wb: Workbook; refersTo: string; onClose: () => void; onDone: (msg: string) => void }) {
  const [name, setName] = useState('')
  const [ref, setRef] = useState(refersTo)
  const [err, setErr] = useState<string | null>(null)
  const existing = wb.namedRanges()
  const apply = () => {
    const e = wb.addNamedRange(name.trim(), ref.startsWith('=') ? ref : `=${ref}`)
    if (e) setErr(e)
    else onDone(`Defined name “${name.trim()}”. Use it in formulas, e.g. =SUM(${name.trim()}).`)
  }
  return (
    <Dialog
      title="Define name"
      onClose={onClose}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button variant="default" disabled={!name.trim()} onClick={apply}>
            OK
          </Button>
        </>
      }
    >
      <Field label="Name" hint="Letters, numbers, underscores and periods; must not look like a cell address.">
        <input className={inputClass} value={name} onChange={(e) => (setName(e.target.value), setErr(null))} />
      </Field>
      <Field label="Refers to">
        <input className={cn(inputClass, 'font-mono')} value={ref} onChange={(e) => (setRef(e.target.value), setErr(null))} />
      </Field>
      {err && <p className="text-[12.5px] text-[var(--danger)]">{err}</p>}
      {existing.length > 0 && (
        <div className="mt-2">
          <div className="mb-1 text-[12px] font-medium text-[var(--muted)]">Names in this workbook</div>
          <ul className="max-h-32 overflow-auto rounded-md border border-[var(--border)] text-[12px]">
            {existing.map((n) => (
              <li key={n.name} className="flex justify-between gap-2 px-2 py-1">
                <span className="font-medium">{n.name}</span>
                <span className="truncate font-mono text-[var(--muted)]">{n.expression}</span>
              </li>
            ))}
          </ul>
        </div>
      )}
    </Dialog>
  )
}

// ------------------------------------------------------------------ compatibility
export function CompatDialog({ onClose }: { onClose: () => void }) {
  const rows = useMemo(() => compatibilityMatrix(), [])
  const [filter, setFilter] = useState<'all' | 'partial' | 'unsupported'>('all')
  const counts = { supported: rows.filter((r) => r.status === 'supported').length, partial: rows.filter((r) => r.status === 'partial').length, unsupported: rows.filter((r) => r.status === 'unsupported').length }
  const shown = rows.filter((r) => filter === 'all' || r.status === filter)
  return (
    <Dialog title="Function compatibility" onClose={onClose} width={720}>
      <p className="mb-2 text-[12.5px] text-[var(--muted)]">
        Of the {rows.length} Excel functions in the product spec, {counts.supported} are supported, {counts.partial} partially, and {counts.unsupported} not yet. The engine provides{' '}
        {engineFunctions().size} functions in total. Formulas with unsupported functions are kept unchanged and show #NAME?.
      </p>
      <div className="mb-2 flex gap-1">
        {(['all', 'partial', 'unsupported'] as const).map((f) => (
          <Button key={f} size="sm" variant={filter === f ? 'outline' : 'ghost'} onClick={() => setFilter(f)}>
            {f === 'all' ? 'All' : f === 'partial' ? `Partial (${counts.partial})` : `Unsupported (${counts.unsupported})`}
          </Button>
        ))}
      </div>
      <div className="max-h-[55vh] overflow-auto rounded-md border border-[var(--border)]">
        <table className="w-full text-[12px]">
          <thead className="sticky top-0 bg-[var(--panel)] text-left text-[var(--muted)]">
            <tr>
              <th className="px-2 py-1 font-medium">Function</th>
              <th className="px-2 py-1 font-medium">Category</th>
              <th className="px-2 py-1 font-medium">Status</th>
              <th className="px-2 py-1 font-medium">Notes</th>
            </tr>
          </thead>
          <tbody>
            {shown.map((r) => (
              <tr key={r.name + r.category} className="border-t border-[var(--border)] align-top">
                <td className="px-2 py-1 font-mono">{r.name}</td>
                <td className="px-2 py-1 text-[var(--muted)]">{r.category}</td>
                <td className={cn('px-2 py-1 whitespace-nowrap', r.status === 'supported' ? 'text-[var(--accent)]' : r.status === 'partial' ? 'text-amber-600' : 'text-[var(--danger)]')}>
                  {r.status === 'supported' ? 'Supported' : r.status === 'partial' ? 'Partial' : 'Unsupported'}
                </td>
                <td className="px-2 py-1 text-[var(--muted)]">{r.status === 'unsupported' ? '' : r.note}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Dialog>
  )
}

// ------------------------------------------------------------------ shortcuts
const SHORTCUTS: [string, string][] = [
  ['Ctrl+C / Ctrl+X / Ctrl+V', 'Copy / cut / paste (also from Excel & Google Sheets)'],
  ['Ctrl+Z / Ctrl+Y', 'Undo / redo (Ctrl+Shift+Z also redoes)'],
  ['Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+5', 'Bold / italic / underline / strikethrough'],
  ['Ctrl+F / Ctrl+H', 'Find / find and replace'],
  ['Ctrl+S', 'Save now'],
  ['Ctrl+A', 'Select all'],
  ['Ctrl+D / Ctrl+R', 'Fill down / fill right'],
  ['Ctrl+`', 'Show formulas'],
  ['F2', 'Edit active cell'],
  ['Enter / Shift+Enter', 'Commit and move down / up'],
  ['Tab / Shift+Tab', 'Commit and move right / left'],
  ['Ctrl+Enter', 'Enter the same value in all selected cells'],
  ['Alt+Enter', 'New line inside a cell'],
  ['Esc', 'Cancel editing / clear copy outline'],
  ['Arrow keys', 'Move (Shift extends the selection)'],
  ['Ctrl+Arrow', 'Jump to the edge of the data region'],
  ['Home / Ctrl+Home / Ctrl+End', 'Row start / A1 / last used cell'],
  ['Page Up / Page Down', 'Move one screen'],
  ['Ctrl+Space / Shift+Space', 'Select column / row'],
  ['Delete / Backspace', 'Clear contents / clear and edit'],
  ['Alt+=', 'AutoSum'],
  ['Shift+F11', 'Insert sheet'],
  ['Ctrl+PageUp / Ctrl+PageDown', 'Previous / next sheet'],
]

export function ShortcutsDialog({ onClose }: { onClose: () => void }) {
  return (
    <Dialog title="Keyboard shortcuts" onClose={onClose} width={560}>
      <table className="w-full text-[12.5px]">
        <tbody>
          {SHORTCUTS.map(([k, d]) => (
            <tr key={k} className="border-t border-[var(--border)] first:border-t-0">
              <td className="whitespace-nowrap py-1.5 pr-3 font-mono text-[12px]">{shortcut(k)}</td>
              <td className="py-1.5 text-[var(--muted)]">{d}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </Dialog>
  )
}

// ------------------------------------------------------------------ about
export function AboutDialog({ onClose }: { onClose: () => void }) {
  return (
    <Dialog title="About ExcelPro (Phase 1)" onClose={onClose} width={620}>
      <div className="space-y-3 text-[13px] leading-5">
        <p>
          ExcelPro is a browser spreadsheet with a real formula engine (HyperFormula), multi-sheet workbooks, formatting, undo/redo, XLSX/CSV import & export and autosave. It is an
          independent product and is not affiliated with Microsoft.
        </p>
        <div>
          <h3 className="mb-1 font-semibold">Working in this release</h3>
          <p className="text-[var(--muted)]">
            Editing, formulas with dependency tracking and recalculation, cross-sheet references and named ranges, fill handle with series, copy/paste (internal and from other apps),
            fonts/colors/borders/alignment/wrap/merge/number formats, row & column insert/delete/resize/hide, worksheet management, find & replace, autosave in this browser with conflict
            detection, XLSX/CSV/TSV import & export.
          </p>
        </div>
        <div>
          <h3 className="mb-1 font-semibold">Where files are stored</h3>
          <p className="text-[var(--muted)]">
            Workbooks are saved in this browser’s storage (IndexedDB) on this device. Clearing site data deletes them; download an .xlsx copy to keep a backup. Cloud accounts, sharing and
            version history arrive in Phase 4.
          </p>
        </div>
        <div>
          <h3 className="mb-1 font-semibold">Import limits</h3>
          <p className="text-[var(--muted)]">
            Charts, images, tables, pivot tables, comments, conditional formatting, data validation, hyperlinks, filters and macros in .xlsx files are not imported yet; you get a report
            listing what was skipped. Legacy .xls and .ods files: open them in a spreadsheet app and save as .xlsx first.
          </p>
        </div>
        <div>
          <h3 className="mb-1 font-semibold">Troubleshooting</h3>
          <ul className="list-disc pl-5 text-[var(--muted)]">
            <li>
              <b>#NAME?</b> — the formula uses a function ExcelPro doesn’t support yet, or a misspelled name. The status bar names the function.
            </li>
            <li>
              <b>#CYCLE!</b> — a circular reference; the formula depends on its own result.
            </li>
            <li>
              <b>Toolbar Paste does nothing</b> — some browsers block toolbar access to the clipboard. Use {shortcut('Ctrl+V')}.
            </li>
            <li>
              <b>“Changed in another tab”</b> — the same workbook is open twice. Pick which version to keep in the banner.
            </li>
          </ul>
        </div>
      </div>
    </Dialog>
  )
}

// ------------------------------------------------------------------ import report
export function ImportReportDialog({ report, fileName, onClose }: { report: ImportReport; fileName: string; onClose: () => void }) {
  return (
    <Dialog
      title="Import report"
      onClose={onClose}
      width={560}
      footer={
        <Button variant="default" onClick={onClose} data-autofocus>
          OK
        </Button>
      }
    >
      <p className="mb-2 text-[13px]">
        Opened <b>{fileName}</b>: {report.stats.sheets} sheet{report.stats.sheets === 1 ? '' : 's'}, {report.stats.cells.toLocaleString()} cells, {report.stats.formulas.toLocaleString()} formulas.
      </p>
      {report.warnings.length > 0 ? (
        <ul className="space-y-1.5">
          {report.warnings.map((w, i) => (
            <li key={i} className="flex gap-2 rounded-md bg-[var(--warn-bg)] p-2 text-[12.5px] text-[var(--warn-text)]">
              <Icon name="alert" size={14} className="mt-0.5 shrink-0" />
              <span>{w}</span>
            </li>
          ))}
        </ul>
      ) : (
        <p className="text-[12.5px] text-[var(--muted)]">Everything in the file was imported.</p>
      )}
    </Dialog>
  )
}

export function ConfirmDialog({ title, message, confirm, danger, onConfirm, onClose }: { title: string; message: string; confirm: string; danger?: boolean; onConfirm: () => void; onClose: () => void }) {
  return (
    <Dialog
      title={title}
      onClose={onClose}
      width={400}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button variant={danger ? 'danger' : 'default'} onClick={onConfirm} data-autofocus>
            {confirm}
          </Button>
        </>
      }
    >
      <p className="text-[13px]">{message}</p>
    </Dialog>
  )
}
