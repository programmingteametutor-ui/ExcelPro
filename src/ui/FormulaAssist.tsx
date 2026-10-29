/* Formula autocomplete (function and name suggestions) and argument hints. */
import { createPortal } from 'react-dom'
import type { Workbook } from '../engine/workbook'
import { engineFunctions, FUNCTION_HELP, helpFor } from '../engine/compat'

export interface Suggestion {
  name: string
  kind: 'function' | 'name' | 'table'
  detail: string
}

export interface AssistState {
  /** start of the token being completed */
  from: number
  suggestions: Suggestion[]
  /** innermost function call around the caret */
  hint: { name: string; syntax: string; arg: number } | null
}

let allFns: string[] | null = null
function functionNames(): string[] {
  if (!allFns) allFns = [...new Set([...engineFunctions(), ...Object.keys(FUNCTION_HELP)])].filter((n) => /^[A-Z][A-Z0-9.]*$/.test(n)).sort()
  return allFns
}

/** Scans the formula up to the caret, skipping strings, to find the token and the open call. */
export function analyzeFormula(wb: Workbook, text: string, caret: number): AssistState | null {
  if (!text.startsWith('=')) return null
  const before = text.slice(0, caret)
  // inside a string literal? count quotes
  let inStr = false
  const stack: { name: string; arg: number }[] = []
  let tokenStart = -1
  for (let i = 1; i < before.length; i++) {
    const ch = before[i]
    if (ch === '"') {
      inStr = !inStr
      continue
    }
    if (inStr) continue
    if (ch === '(') {
      const m = /([A-Za-z_][A-Za-z0-9._]*)$/.exec(before.slice(0, i))
      stack.push({ name: m ? m[1].toUpperCase() : '', arg: 0 })
    } else if (ch === ')') stack.pop()
    else if ((ch === ',' || ch === ';') && stack.length) stack[stack.length - 1].arg++
  }
  if (inStr) return null
  const m = /(^|[=(,;+\-*/^&<>:\s])([A-Za-z_][A-Za-z0-9._]*)$/.exec(before)
  let suggestions: Suggestion[] = []
  if (m && !/^[A-Za-z]{1,3}\d+$/.test(m[2])) {
    const prefix = m[2].toUpperCase()
    tokenStart = before.length - m[2].length
    // a sheet-qualified ref or a cell ref like A1 is not a function
    if (!/!$/.test(before.slice(0, tokenStart))) {
      const fns = functionNames().filter((n) => n.startsWith(prefix))
      const names = wb.namedRanges().filter((n) => n.name.toUpperCase().startsWith(prefix)).map((n) => ({ name: n.name, kind: 'name' as const, detail: n.expression }))
      const tables = wb.sheets.flatMap((s) => s.tables ?? []).filter((t) => t.name.toUpperCase().startsWith(prefix)).map((t) => ({ name: t.name, kind: 'table' as const, detail: 'Table' }))
      suggestions = [
        ...names,
        ...tables,
        ...fns.map((n) => ({ name: n, kind: 'function' as const, detail: helpFor(n)?.description ?? '' })),
      ].slice(0, 9)
      if (suggestions.length === 1 && suggestions[0].name.toUpperCase() === prefix && suggestions[0].kind !== 'function') suggestions = []
    }
  }
  const call = [...stack].reverse().find((s) => s.name && helpFor(s.name))
  const hint = call ? { name: call.name, syntax: helpFor(call.name)!.syntax, arg: call.arg } : null
  if (!suggestions.length && !hint) return null
  return { from: tokenStart, suggestions, hint }
}

/** Inserts the chosen suggestion; functions get an opening parenthesis. */
export function acceptSuggestion(text: string, caret: number, st: AssistState, s: Suggestion): { text: string; caret: number } {
  const after = text.slice(caret)
  const tail = /^[A-Za-z0-9._]*/.exec(after)![0]
  const ins = s.kind === 'function' ? `${s.name}(` : s.kind === 'table' ? `${s.name}[` : s.name
  const skipParen = s.kind === 'function' && after.slice(tail.length).startsWith('(')
  const next = text.slice(0, st.from) + (skipParen ? s.name : ins) + after.slice(tail.length)
  return { text: next, caret: st.from + (skipParen ? s.name.length + 1 : ins.length) }
}

function SyntaxHint({ syntax, arg }: { syntax: string; arg: number }) {
  const open = syntax.indexOf('(')
  if (open < 0) return <span>{syntax}</span>
  const name = syntax.slice(0, open)
  const args = syntax.slice(open + 1, syntax.lastIndexOf(')')).split(',').map((s) => s.trim())
  const repeating = args[args.length - 1] === '...'
  const idx = arg < args.length ? arg : repeating ? Math.max(0, args.length - 2) : args.length - 1
  return (
    <span>
      <b>{name}</b>(
      {args.map((a, i) => (
        <span key={i}>
          {i > 0 && ', '}
          <span className={i === idx ? 'font-semibold underline' : ''}>{a}</span>
        </span>
      ))}
      )
    </span>
  )
}

export function FormulaAssistPanel({ anchor, state, active, onPick, onHover }: { anchor: DOMRect; state: AssistState; active: number; onPick: (s: Suggestion) => void; onHover: (i: number) => void }) {
  const top = anchor.bottom + 2
  const left = Math.max(4, Math.min(anchor.left, window.innerWidth - 384))
  return createPortal(
    <div className="fixed z-[97] w-[380px] max-w-[94vw] text-[12.5px]" style={{ left, top }} onMouseDown={(e) => e.preventDefault()} data-testid="formula-assist">
      {state.hint && !state.suggestions.length && (
        <div className="inline-block rounded-md border border-[var(--border)] bg-[var(--bg)] px-2 py-1 font-mono text-[12px] text-[var(--text)] shadow-lg" data-testid="arg-hint">
          <SyntaxHint syntax={state.hint.syntax} arg={state.hint.arg} />
        </div>
      )}
      {state.suggestions.length > 0 && (
        <div className="overflow-hidden rounded-md border border-[var(--border)] bg-[var(--bg)] text-[var(--text)] shadow-xl" role="listbox" aria-label="Formula suggestions">
          {state.suggestions.map((s, i) => (
            <div
              key={s.kind + s.name}
              role="option"
              aria-selected={i === active}
              className={`flex cursor-pointer items-baseline gap-2 px-2 py-1 ${i === active ? 'bg-[var(--active)]' : ''}`}
              onMouseEnter={() => onHover(i)}
              onClick={() => onPick(s)}
            >
              <span className="w-4 shrink-0 text-center text-[10.5px] text-[var(--muted)]">{s.kind === 'function' ? 'fx' : s.kind === 'table' ? '▦' : '#'}</span>
              <span className="shrink-0 font-mono font-medium">{s.name}</span>
              <span className="min-w-0 truncate text-[11.5px] text-[var(--muted)]">{s.detail}</span>
            </div>
          ))}
          {state.suggestions[active]?.kind === 'function' && helpFor(state.suggestions[active].name) && (
            <div className="border-t border-[var(--border)] bg-[var(--panel)] px-2 py-1 font-mono text-[11.5px] text-[var(--muted)]">{helpFor(state.suggestions[active].name)!.syntax}</div>
          )}
          <div className="border-t border-[var(--border)] px-2 py-0.5 text-[10.5px] text-[var(--muted)]">Tab to insert · ↑↓ to choose · Esc to close</div>
        </div>
      )}
    </div>,
    document.body,
  )
}
