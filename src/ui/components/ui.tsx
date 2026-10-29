/* shadcn/ui-style primitives built on Tailwind + class-variance-authority (no Radix dependency). */
import { cva, type VariantProps } from 'class-variance-authority'
import {
  createContext,
  forwardRef,
  useCallback,
  useContext,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type ButtonHTMLAttributes,
  type ReactNode,
} from 'react'
import { createPortal } from 'react-dom'
import { cn } from '../lib/utils'
import { Icon, type IconName } from '../icons'

// ------------------------------------------------------------------ Button
export const buttonVariants = cva(
  'inline-flex items-center justify-center gap-1.5 whitespace-nowrap rounded-md text-[13px] transition-colors disabled:opacity-40 disabled:pointer-events-none focus-visible:outline-2 focus-visible:outline-offset-1 focus-visible:outline-[var(--accent)]',
  {
    variants: {
      variant: {
        default: 'bg-[var(--accent)] text-[var(--accent-contrast)] hover:opacity-90',
        outline: 'border border-[var(--border)] bg-[var(--bg)] hover:bg-[var(--hover)]',
        ghost: 'hover:bg-[var(--hover)]',
        danger: 'bg-[var(--danger)] text-white hover:opacity-90',
      },
      size: { sm: 'h-7 px-2', md: 'h-8 px-3', icon: 'h-7 w-7', lg: 'h-9 px-4' },
    },
    defaultVariants: { variant: 'ghost', size: 'md' },
  },
)

export interface ButtonProps extends ButtonHTMLAttributes<HTMLButtonElement>, VariantProps<typeof buttonVariants> {
  tip?: string
}

export const Button = forwardRef<HTMLButtonElement, ButtonProps>(({ className, variant, size, tip, ...props }, ref) => {
  const btn = <button ref={ref} type="button" className={cn(buttonVariants({ variant, size }), className)} {...props} />
  return tip ? <Tooltip text={tip}>{btn}</Tooltip> : btn
})
Button.displayName = 'Button'

/** Toolbar icon button with tooltip and pressed state. */
export function ToolButton({
  icon,
  label,
  tip,
  active,
  onClick,
  disabled,
  children,
  className,
  wide,
}: {
  icon?: IconName
  label?: string
  tip: string
  active?: boolean
  onClick?: () => void
  disabled?: boolean
  children?: ReactNode
  className?: string
  wide?: boolean
}) {
  return (
    <Tooltip text={tip}>
      <button
        type="button"
        aria-label={tip}
        aria-pressed={active}
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={onClick}
        className={cn(
          'inline-flex h-7 items-center justify-center gap-1 rounded-[5px] text-[12.5px] text-[var(--text)] hover:bg-[var(--hover)] disabled:pointer-events-none disabled:opacity-35',
          wide ? 'px-2' : 'min-w-7 px-1',
          active && 'bg-[var(--active)] text-[var(--accent)] hover:bg-[var(--active)]',
          className,
        )}
      >
        {icon && <Icon name={icon} />}
        {label && <span>{label}</span>}
        {children}
      </button>
    </Tooltip>
  )
}

// ------------------------------------------------------------------ Tooltip
export function Tooltip({ text, children }: { text: string; children: ReactNode }) {
  const [show, setShow] = useState<DOMRect | null>(null)
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const wrap = useRef<HTMLSpanElement>(null)
  const enter = () => {
    timer.current = setTimeout(() => {
      const el = wrap.current?.firstElementChild as HTMLElement | null
      if (el) setShow(el.getBoundingClientRect())
    }, 450)
  }
  const leave = () => {
    if (timer.current) clearTimeout(timer.current)
    setShow(null)
  }
  return (
    <span ref={wrap} className="contents" onMouseEnter={enter} onMouseLeave={leave} onMouseDown={leave}>
      {children}
      {show &&
        createPortal(
          <div
            role="tooltip"
            className="pointer-events-none fixed z-[100] max-w-64 rounded-md bg-[#1f2328] px-2 py-1 text-[11.5px] text-white shadow-lg dark:bg-[#e7e9ec] dark:text-[#1f2328]"
            style={{ left: Math.min(Math.max(4, show.left + show.width / 2 - 60), window.innerWidth - 260), top: show.bottom + 6 }}
          >
            {text}
          </div>,
          document.body,
        )}
    </span>
  )
}

// ------------------------------------------------------------------ Popover / Menu
interface PopoverCtx {
  close: () => void
}
const PopCtx = createContext<PopoverCtx>({ close: () => {} })
export const usePopover = () => useContext(PopCtx)

function useOutside(ref: React.RefObject<HTMLElement | null>, extra: React.RefObject<HTMLElement | null>, onOut: () => void, active: boolean) {
  useEffect(() => {
    if (!active) return
    const h = (e: PointerEvent) => {
      const t = e.target as Node
      if (ref.current?.contains(t) || extra.current?.contains(t)) return
      onOut()
    }
    const k = (e: KeyboardEvent) => e.key === 'Escape' && onOut()
    document.addEventListener('pointerdown', h, true)
    document.addEventListener('keydown', k, true)
    return () => {
      document.removeEventListener('pointerdown', h, true)
      document.removeEventListener('keydown', k, true)
    }
  }, [ref, extra, onOut, active])
}

export function Popover({
  trigger,
  children,
  align = 'start',
  className,
}: {
  trigger: (props: { open: boolean; toggle: () => void }) => ReactNode
  children: ReactNode | ((close: () => void) => ReactNode)
  align?: 'start' | 'end'
  className?: string
}) {
  const [open, setOpen] = useState(false)
  const anchor = useRef<HTMLSpanElement>(null)
  const panel = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null)
  const close = useCallback(() => setOpen(false), [])
  useOutside(panel, anchor, close, open)
  useLayoutEffect(() => {
    if (!open) return
    const a = (anchor.current?.firstElementChild ?? anchor.current) as HTMLElement
    const r = a.getBoundingClientRect()
    const pw = panel.current?.offsetWidth ?? 200
    const ph = panel.current?.offsetHeight ?? 200
    let left = align === 'end' ? r.right - pw : r.left
    left = Math.max(4, Math.min(left, window.innerWidth - pw - 4))
    let top = r.bottom + 4
    if (top + ph > window.innerHeight - 4) top = Math.max(4, r.top - ph - 4)
    setPos({ left, top })
  }, [open, align])
  return (
    <>
      <span ref={anchor} className="contents">
        {trigger({ open, toggle: () => setOpen((o) => !o) })}
      </span>
      {open &&
        createPortal(
          <PopCtx.Provider value={{ close }}>
            <div
              ref={panel}
              className={cn(
                'fixed z-[90] min-w-44 rounded-lg border border-[var(--border)] bg-[var(--bg)] p-1 text-[var(--text)] shadow-xl',
                className,
              )}
              style={{ left: pos?.left ?? -9999, top: pos?.top ?? -9999 }}
              onMouseDown={(e) => {
                if (!(e.target as HTMLElement).closest('input,select,textarea')) e.preventDefault()
              }}
            >
              {typeof children === 'function' ? children(close) : children}
            </div>
          </PopCtx.Provider>,
          document.body,
        )}
    </>
  )
}

export function MenuItem({
  icon,
  children,
  onSelect,
  shortcut,
  disabled,
  danger,
  checked,
}: {
  icon?: IconName
  children: ReactNode
  onSelect?: () => void
  shortcut?: string
  disabled?: boolean
  danger?: boolean
  checked?: boolean
}) {
  const { close } = usePopover()
  return (
    <button
      type="button"
      role="menuitem"
      disabled={disabled}
      onClick={() => {
        close()
        onSelect?.()
      }}
      className={cn(
        'flex h-8 w-full items-center gap-2 rounded-md px-2 text-left text-[13px] hover:bg-[var(--hover)] disabled:opacity-40',
        danger && 'text-[var(--danger)]',
      )}
    >
      <span className="flex w-4 justify-center">{checked !== undefined ? checked && <Icon name="check" size={14} /> : icon && <Icon name={icon} size={15} />}</span>
      <span className="flex-1">{children}</span>
      {shortcut && <span className="text-[11px] text-[var(--muted)]">{shortcut}</span>}
    </button>
  )
}

export const MenuSep = () => <div className="my-1 h-px bg-[var(--border)]" />
export const MenuLabel = ({ children }: { children: ReactNode }) => (
  <div className="px-2 pb-1 pt-1.5 text-[11px] font-semibold uppercase tracking-wide text-[var(--muted)]">{children}</div>
)

/** Context menu rendered at a point. */
export function ContextMenu({ x, y, onClose, children }: { x: number; y: number; onClose: () => void; children: ReactNode }) {
  const ref = useRef<HTMLDivElement>(null)
  const none = useRef<HTMLElement>(null)
  useOutside(ref, none, onClose, true)
  const [pos, setPos] = useState({ left: x, top: y })
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    setPos({ left: Math.min(x, window.innerWidth - el.offsetWidth - 4), top: Math.min(y, window.innerHeight - el.offsetHeight - 4) })
  }, [x, y])
  return createPortal(
    <PopCtx.Provider value={{ close: onClose }}>
      <div
        ref={ref}
        role="menu"
        className="fixed z-[95] min-w-52 rounded-lg border border-[var(--border)] bg-[var(--bg)] p-1 text-[var(--text)] shadow-xl"
        style={pos}
        onMouseDown={(e) => e.preventDefault()}
        onContextMenu={(e) => e.preventDefault()}
      >
        {children}
      </div>
    </PopCtx.Provider>,
    document.body,
  )
}

// ------------------------------------------------------------------ Dialog
export function Dialog({
  title,
  onClose,
  children,
  footer,
  width = 440,
}: {
  title: string
  onClose: () => void
  children: ReactNode
  footer?: ReactNode
  width?: number
}) {
  const ref = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const k = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        onClose()
      }
    }
    document.addEventListener('keydown', k, true)
    const first = ref.current?.querySelector<HTMLElement>('input,select,textarea,button[data-autofocus]')
    first?.focus()
    return () => document.removeEventListener('keydown', k, true)
  }, [onClose])
  return createPortal(
    <div className="fixed inset-0 z-[80] flex items-center justify-center bg-black/30 p-4" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div
        ref={ref}
        role="dialog"
        aria-modal="true"
        aria-label={title}
        className="flex max-h-[90vh] w-full flex-col rounded-xl border border-[var(--border)] bg-[var(--bg)] text-[var(--text)] shadow-2xl"
        style={{ maxWidth: width }}
      >
        <div className="flex items-center justify-between border-b border-[var(--border)] px-4 py-3">
          <h2 className="text-[15px] font-semibold">{title}</h2>
          <Button size="icon" onClick={onClose} aria-label="Close">
            <Icon name="x" />
          </Button>
        </div>
        <div className="overflow-auto px-4 py-3">{children}</div>
        {footer && <div className="flex justify-end gap-2 border-t border-[var(--border)] px-4 py-3">{footer}</div>}
      </div>
    </div>,
    document.body,
  )
}

// ------------------------------------------------------------------ Inputs
export const inputClass =
  'h-8 rounded-md border border-[var(--border)] bg-[var(--bg)] px-2 text-[13px] text-[var(--text)] outline-none focus:border-[var(--accent)] focus:ring-2 focus:ring-[var(--accent)]/20'

export function Field({ label, children, hint }: { label: string; children: ReactNode; hint?: string }) {
  return (
    <label className="mb-3 flex flex-col gap-1">
      <span className="text-[12px] font-medium text-[var(--muted)]">{label}</span>
      {children}
      {hint && <span className="text-[11.5px] text-[var(--muted)]">{hint}</span>}
    </label>
  )
}

export function Check({ checked, onChange, label }: { checked: boolean; onChange: (v: boolean) => void; label: string }) {
  return (
    <label className="flex h-7 cursor-default items-center gap-2 rounded px-1.5 text-[12.5px] hover:bg-[var(--hover)]">
      <input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} className="accent-[var(--accent)]" />
      {label}
    </label>
  )
}

// ------------------------------------------------------------------ Color picker
const PALETTE = [
  ['#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff'],
  ['#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff'],
  ['#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc'],
  ['#dd7e6b', '#ea9999', '#f9cb9c', '#ffe599', '#b6d7a8', '#a2c4c9', '#a4c2f4', '#9fc5e8', '#b4a7d6', '#d5a6bd'],
  ['#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0'],
  ['#a61c00', '#cc0000', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#3d85c6', '#674ea7', '#a64d79'],
  ['#5b0f00', '#660000', '#783f04', '#7f6000', '#274e13', '#0c343d', '#1c4587', '#073763', '#20124d', '#4c1130'],
]

export function ColorGrid({ onPick, noneLabel }: { onPick: (c: string | undefined) => void; noneLabel: string }) {
  const { close } = usePopover()
  const [custom, setCustom] = useState('#0f766e')
  return (
    <div className="p-1">
      <button
        type="button"
        className="mb-1.5 flex h-7 w-full items-center gap-2 rounded px-1.5 text-[12.5px] hover:bg-[var(--hover)]"
        onClick={() => {
          onPick(undefined)
          close()
        }}
      >
        <span className="h-4 w-4 rounded border border-[var(--border)] bg-[linear-gradient(135deg,transparent_45%,#d00_45%,#d00_55%,transparent_55%)]" />
        {noneLabel}
      </button>
      <div className="grid grid-cols-10 gap-[3px]">
        {PALETTE.flat().map((c) => (
          <button
            key={c}
            type="button"
            aria-label={c}
            title={c}
            className="h-[18px] w-[18px] rounded-[3px] border border-black/10 hover:scale-110"
            style={{ background: c }}
            onClick={() => {
              onPick(c)
              close()
            }}
          />
        ))}
      </div>
      <div className="mt-2 flex items-center gap-2 px-0.5">
        <input type="color" value={custom} onChange={(e) => setCustom(e.target.value)} className="h-7 w-9 cursor-pointer rounded border border-[var(--border)] bg-transparent" aria-label="Custom color" />
        <Button size="sm" variant="outline" onClick={() => (onPick(custom), close())}>
          Use custom
        </Button>
      </div>
    </div>
  )
}
