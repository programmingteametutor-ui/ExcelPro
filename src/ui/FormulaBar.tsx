import { useEffect, useRef, useState } from 'react'
import { Tooltip } from './components/ui'
import { Icon } from './icons'

export function FormulaBar({
  address,
  text,
  editing,
  onNavigate,
  onFocusInput,
  onChange,
  onKeyDown,
  onCancel,
  onCommit,
  onInsertFunction,
  inputRef,
}: {
  address: string
  text: string
  editing: boolean
  onNavigate: (target: string) => string | null
  onFocusInput: () => void
  onChange: (text: string, caret: number) => void
  onKeyDown: (e: React.KeyboardEvent<HTMLTextAreaElement>) => void
  onCancel: () => void
  onCommit: () => void
  onInsertFunction: () => void
  inputRef: React.RefObject<HTMLTextAreaElement | null>
}) {
  const [box, setBox] = useState(address)
  const [boxErr, setBoxErr] = useState<string | null>(null)
  const [expanded, setExpanded] = useState(false)
  const boxRef = useRef<HTMLInputElement>(null)
  useEffect(() => {
    if (document.activeElement !== boxRef.current) setBox(address)
  }, [address])
  const lines = text.split('\n').length

  return (
    <div className="flex shrink-0 items-start gap-1 border-b border-[var(--border)] bg-[var(--bg)] px-1.5 py-1">
      <div className="relative">
        <input
          ref={boxRef}
          aria-label="Name box"
          data-testid="name-box"
          className="h-7 w-[92px] rounded-[5px] border border-[var(--border)] bg-[var(--bg)] px-1.5 font-mono text-[12.5px] outline-none focus:border-[var(--accent)] sm:w-[120px]"
          value={box}
          onChange={(e) => {
            setBox(e.target.value)
            setBoxErr(null)
          }}
          onFocus={(e) => e.target.select()}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault()
              const err = onNavigate(box.trim())
              setBoxErr(err)
              if (!err) (e.target as HTMLInputElement).blur()
            } else if (e.key === 'Escape') {
              setBox(address)
              setBoxErr(null)
              ;(e.target as HTMLInputElement).blur()
            }
          }}
          onBlur={() => {
            setBox(address)
          }}
        />
        {boxErr && <div className="absolute left-0 top-8 z-50 w-60 rounded-md border border-[var(--border)] bg-[var(--bg)] p-2 text-[12px] text-[var(--danger)] shadow-lg">{boxErr}</div>}
      </div>
      <div className="flex h-7 items-center">
        <Tooltip text="Cancel (Esc)">
          <button type="button" aria-label="Cancel edit" disabled={!editing} onMouseDown={(e) => e.preventDefault()} onClick={onCancel} className="flex h-7 w-6 items-center justify-center rounded text-[var(--muted)] hover:bg-[var(--hover)] disabled:opacity-30">
            <Icon name="x" size={14} />
          </button>
        </Tooltip>
        <Tooltip text="Enter">
          <button type="button" aria-label="Confirm edit" disabled={!editing} onMouseDown={(e) => e.preventDefault()} onClick={onCommit} className="flex h-7 w-6 items-center justify-center rounded text-[var(--muted)] hover:bg-[var(--hover)] disabled:opacity-30">
            <Icon name="check" size={14} />
          </button>
        </Tooltip>
        <Tooltip text="Insert function">
          <button type="button" aria-label="Insert function" onMouseDown={(e) => e.preventDefault()} onClick={onInsertFunction} className="flex h-7 w-7 items-center justify-center rounded text-[var(--muted)] hover:bg-[var(--hover)]">
            <span className="font-serif text-[14px] italic">fx</span>
          </button>
        </Tooltip>
      </div>
      <textarea
        ref={inputRef}
        aria-label="Formula bar"
        data-testid="formula-bar"
        spellCheck={false}
        className="min-h-7 flex-1 resize-none rounded-[5px] border border-[var(--border)] bg-[var(--bg)] px-2 py-[5px] font-mono text-[12.5px] leading-[16px] outline-none focus:border-[var(--accent)]"
        rows={expanded ? Math.min(8, Math.max(3, lines)) : 1}
        value={text}
        onFocus={onFocusInput}
        onChange={(e) => onChange(e.target.value, e.target.selectionStart)}
        onSelect={(e) => onChange(e.currentTarget.value, e.currentTarget.selectionStart)}
        onKeyDown={onKeyDown}
      />
      <Tooltip text={expanded ? 'Collapse formula bar' : 'Expand formula bar'}>
        <button type="button" aria-label="Toggle formula bar height" onClick={() => setExpanded((x) => !x)} className="flex h-7 w-6 items-center justify-center rounded text-[var(--muted)] hover:bg-[var(--hover)]">
          <Icon name="chevronDown" size={14} className={expanded ? 'rotate-180' : ''} />
        </button>
      </Tooltip>
    </div>
  )
}
