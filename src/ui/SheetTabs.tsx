import { useState } from 'react'
import type { Workbook } from '../engine/workbook'
import { ColorGrid, ContextMenu, MenuItem, MenuLabel, MenuSep, Popover, Tooltip } from './components/ui'
import { Icon } from './icons'
import { cn } from './lib/utils'

export function SheetTabs({ wb, onError, onChange }: { wb: Workbook; onError: (msg: string) => void; onChange: () => void }) {
  const [renaming, setRenaming] = useState<string | null>(null)
  const [renameText, setRenameText] = useState('')
  const [menu, setMenu] = useState<{ x: number; y: number; id: string } | null>(null)
  const [dragId, setDragId] = useState<string | null>(null)
  const [dropIdx, setDropIdx] = useState<number | null>(null)
  const sheets = wb.sheets
  const hidden = sheets.filter((s) => s.hidden)

  const commitRename = () => {
    if (!renaming) return
    const err = wb.renameSheet(renaming, renameText)
    if (err) onError(err)
    setRenaming(null)
    onChange()
  }

  const startRename = (id: string) => {
    setRenaming(id)
    setRenameText(wb.meta(id).name)
  }

  return (
    <div className="flex h-8 shrink-0 items-stretch border-t border-[var(--border)] bg-[var(--panel)]" data-testid="sheet-tabs">
      <Tooltip text="New sheet (Shift+F11)">
        <button
          type="button"
          aria-label="Add sheet"
          className="flex w-8 items-center justify-center text-[var(--muted)] hover:bg-[var(--hover)]"
          onClick={() => {
            wb.addSheet()
            onChange()
          }}
        >
          <Icon name="plus" />
        </button>
      </Tooltip>
      <Popover
        trigger={({ toggle }) => (
          <Tooltip text="All sheets">
            <button type="button" aria-label="All sheets" className="flex w-7 items-center justify-center text-[var(--muted)] hover:bg-[var(--hover)]" onClick={toggle}>
              <Icon name="more" />
            </button>
          </Tooltip>
        )}
      >
        <MenuLabel>Sheets</MenuLabel>
        {sheets.map((s) => (
          <MenuItem key={s.id} checked={s.id === wb.activeSheetId} onSelect={() => (s.hidden ? wb.setSheetHidden(s.id, false) : wb.setActiveSheet(s.id), onChange())}>
            {s.name}
            {s.hidden ? ' (hidden)' : ''}
          </MenuItem>
        ))}
      </Popover>
      <div className="xp-scroll-thin flex flex-1 items-stretch overflow-x-auto" role="tablist" aria-label="Worksheets">
        {sheets.map((s, idx) =>
          s.hidden ? null : (
            <div
              key={s.id}
              role="tab"
              aria-selected={s.id === wb.activeSheetId}
              draggable={renaming !== s.id}
              onDragStart={(e) => {
                setDragId(s.id)
                e.dataTransfer.effectAllowed = 'move'
                e.dataTransfer.setData('text/plain', s.name)
              }}
              onDragOver={(e) => {
                if (!dragId) return
                e.preventDefault()
                const r = e.currentTarget.getBoundingClientRect()
                setDropIdx(e.clientX < r.left + r.width / 2 ? idx : idx + 1)
              }}
              onDrop={(e) => {
                e.preventDefault()
                if (dragId !== null && dropIdx !== null) {
                  const from = wb.order.indexOf(dragId)
                  wb.moveSheet(dragId, dropIdx > from ? dropIdx - 1 : dropIdx)
                  onChange()
                }
                setDragId(null)
                setDropIdx(null)
              }}
              onDragEnd={() => {
                setDragId(null)
                setDropIdx(null)
              }}
              onClick={() => {
                if (s.id === wb.activeSheetId) return
                wb.setActiveSheet(s.id)
                onChange()
              }}
              onDoubleClick={() => startRename(s.id)}
              onContextMenu={(e) => {
                e.preventDefault()
                setMenu({ x: e.clientX, y: e.clientY, id: s.id })
              }}
              className={cn(
                'relative flex min-w-[72px] max-w-[200px] shrink-0 items-center border-r border-[var(--border)] px-3 text-[12.5px] hover:bg-[var(--hover)]',
                s.id === wb.activeSheetId && 'bg-[var(--bg)] font-semibold text-[var(--accent)] hover:bg-[var(--bg)]',
                dropIdx === idx && 'shadow-[inset_2px_0_0_var(--accent)]',
                dropIdx === idx + 1 && idx === sheets.length - 1 && 'shadow-[inset_-2px_0_0_var(--accent)]',
              )}
            >
              {s.tabColor && <span className="absolute inset-x-1 bottom-0 h-[3px] rounded-t" style={{ background: s.tabColor }} />}
              {s.id === wb.activeSheetId && <span className="absolute inset-x-0 top-0 h-[2px] bg-[var(--accent)]" />}
              {renaming === s.id ? (
                <input
                  autoFocus
                  aria-label="Sheet name"
                  className="h-6 w-28 rounded border border-[var(--accent)] bg-[var(--bg)] px-1 text-[12.5px] outline-none"
                  value={renameText}
                  onChange={(e) => setRenameText(e.target.value)}
                  onBlur={commitRename}
                  onKeyDown={(e) => {
                    e.stopPropagation()
                    if (e.key === 'Enter') commitRename()
                    if (e.key === 'Escape') setRenaming(null)
                  }}
                  onFocus={(e) => e.target.select()}
                />
              ) : (
                <span className="truncate">{s.name}</span>
              )}
            </div>
          ),
        )}
      </div>
      {menu && (
        <ContextMenu x={menu.x} y={menu.y} onClose={() => setMenu(null)}>
          <MenuItem icon="plus" onSelect={() => (wb.addSheet(undefined, wb.order.indexOf(menu.id)), onChange())}>Insert sheet</MenuItem>
          <MenuItem icon="edit" onSelect={() => startRename(menu.id)}>Rename</MenuItem>
          <MenuItem icon="dup" onSelect={() => (wb.duplicateSheet(menu.id), onChange())}>Duplicate</MenuItem>
          <MenuItem icon="chevronLeft" disabled={wb.order.indexOf(menu.id) === 0} onSelect={() => (wb.moveSheet(menu.id, wb.order.indexOf(menu.id) - 1), onChange())}>
            Move left
          </MenuItem>
          <MenuItem icon="chevronRight" disabled={wb.order.indexOf(menu.id) === wb.order.length - 1} onSelect={() => (wb.moveSheet(menu.id, wb.order.indexOf(menu.id) + 1), onChange())}>
            Move right
          </MenuItem>
          <MenuSep />
          <MenuLabel>Tab color</MenuLabel>
          <ColorGrid noneLabel="No color" onPick={(c) => (wb.setTabColor(menu.id, c), onChange())} />
          <MenuSep />
          <MenuItem
            icon="eye"
            onSelect={() => {
              const err = wb.setSheetHidden(menu.id, true)
              if (err) onError(err)
              onChange()
            }}
          >
            Hide
          </MenuItem>
          {hidden.length > 0 && (
            <>
              <MenuLabel>Unhide</MenuLabel>
              {hidden.map((h) => (
                <MenuItem key={h.id} onSelect={() => (wb.setSheetHidden(h.id, false), onChange())}>
                  {h.name}
                </MenuItem>
              ))}
            </>
          )}
          <MenuSep />
          <MenuItem
            icon="trash"
            danger
            onSelect={() => {
              const err = wb.removeSheet(menu.id)
              if (err) onError(err)
              onChange()
            }}
          >
            Delete
          </MenuItem>
        </ContextMenu>
      )}
    </div>
  )
}
