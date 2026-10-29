import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { WorkbookHeader, WorkbookStore } from '../persistence/store'
import { Button, ContextMenu, Dialog, Field, inputClass, MenuItem, MenuSep } from './components/ui'
import { Icon, Logo } from './icons'
import { cn, relativeTime } from './lib/utils'

export function Home({
  store,
  storageNote,
  onOpen,
  onNew,
  onImport,
  onDuplicate,
  onDownload,
  dark,
  toggleDark,
}: {
  store: WorkbookStore
  storageNote: string | null
  onOpen: (id: string) => void
  onNew: () => void
  onImport: (f: File) => void
  onDuplicate: (id: string) => Promise<void>
  onDownload: (id: string) => Promise<void>
  dark: boolean
  toggleDark: () => void
}) {
  const [items, setItems] = useState<WorkbookHeader[] | null>(null)
  const [q, setQ] = useState('')
  const [sort, setSort] = useState<'modified' | 'name'>('modified')
  const [trash, setTrash] = useState(false)
  const [menu, setMenu] = useState<{ x: number; y: number; h: WorkbookHeader } | null>(null)
  const [rename, setRename] = useState<WorkbookHeader | null>(null)
  const [renameText, setRenameText] = useState('')
  const [dragOver, setDragOver] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const fileRef = useRef<HTMLInputElement>(null)

  const refresh = useCallback(() => {
    store
      .list(true)
      .then(setItems)
      .catch((e) => setError(e instanceof Error ? e.message : String(e)))
  }, [store])
  useEffect(refresh, [refresh])

  const shown = useMemo(() => {
    const list = (items ?? []).filter((h) => (trash ? !!h.deletedAt : !h.deletedAt)).filter((h) => h.name.toLowerCase().includes(q.trim().toLowerCase()))
    return list.sort((a, b) => (sort === 'name' ? a.name.localeCompare(b.name) : b.updatedAt.localeCompare(a.updatedAt)))
  }, [items, q, sort, trash])
  const trashCount = (items ?? []).filter((h) => h.deletedAt).length

  const run = (p: Promise<unknown>) => p.then(refresh).catch((e) => setError(e instanceof Error ? e.message : String(e)))

  return (
    <div
      className="flex h-full flex-col overflow-auto bg-[var(--bg)]"
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes('Files')) {
          e.preventDefault()
          setDragOver(true)
        }
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault()
        setDragOver(false)
        const f = e.dataTransfer.files?.[0]
        if (f) onImport(f)
      }}
    >
      <header className="flex h-12 shrink-0 items-center gap-2 border-b border-[var(--border)] bg-[var(--panel)] px-4">
        <Logo size={24} />
        <span className="text-[15px] font-semibold tracking-tight">ExcelPro</span>
        <div className="ml-auto flex items-center gap-1">
          <Button size="icon" tip={dark ? 'Light theme' : 'Dark theme'} onClick={toggleDark} aria-label="Toggle theme">
            <Icon name={dark ? 'sun' : 'moon'} />
          </Button>
        </div>
      </header>
      <main className="mx-auto w-full max-w-5xl flex-1 px-4 py-6">
        <div className="mb-6 flex flex-wrap items-center gap-3">
          <button
            type="button"
            onClick={onNew}
            data-testid="new-workbook"
            className="group flex h-24 w-40 flex-col items-start justify-between rounded-xl border border-[var(--border)] bg-[var(--panel)] p-3 text-left hover:border-[var(--accent)]"
          >
            <span className="flex h-8 w-8 items-center justify-center rounded-lg bg-[var(--accent)] text-[var(--accent-contrast)]">
              <Icon name="plus" />
            </span>
            <span className="text-[13px] font-semibold">Blank workbook</span>
          </button>
          <button
            type="button"
            onClick={() => fileRef.current?.click()}
            className={cn('flex h-24 w-56 flex-col items-start justify-between rounded-xl border border-dashed p-3 text-left hover:border-[var(--accent)]', dragOver ? 'border-[var(--accent)] bg-[var(--active)]' : 'border-[var(--border-strong)]')}
          >
            <span className="flex h-8 w-8 items-center justify-center rounded-lg bg-[var(--panel-2)]">
              <Icon name="upload" />
            </span>
            <span>
              <span className="block text-[13px] font-semibold">Open a file</span>
              <span className="block text-[11.5px] text-[var(--muted)]">.xlsx, .csv or .tsv — or drop it here</span>
            </span>
          </button>
          <div className="flex h-24 max-w-sm flex-1 items-center rounded-xl bg-[var(--panel)] p-3 text-[12px] leading-5 text-[var(--muted)]">
            <span>
              <Icon name="info" size={14} className="mr-1 inline align-[-2px]" />
              Starter templates (budget, invoice, trackers…) are planned for Phase 2. Workbooks are stored in this browser; download .xlsx copies for backups.
            </span>
          </div>
        </div>
        {storageNote && (
          <div className="mb-4 flex items-center gap-2 rounded-lg bg-[var(--warn-bg)] px-3 py-2 text-[12.5px] text-[var(--warn-text)]" role="alert">
            <Icon name="alert" size={15} /> {storageNote}
          </div>
        )}
        {error && (
          <div className="mb-4 rounded-lg bg-[var(--warn-bg)] px-3 py-2 text-[12.5px] text-[var(--warn-text)]" role="alert">
            {error}
          </div>
        )}
        <div className="mb-2 flex flex-wrap items-center gap-2">
          <h1 className="mr-auto text-[16px] font-semibold">{trash ? 'Trash' : 'Recent workbooks'}</h1>
          <div className="relative">
            <Icon name="search" size={14} className="absolute left-2 top-2 text-[var(--muted)]" />
            <input className={cn(inputClass, 'w-52 pl-7')} placeholder="Search workbooks" value={q} onChange={(e) => setQ(e.target.value)} aria-label="Search workbooks" />
          </div>
          <select className={inputClass} value={sort} onChange={(e) => setSort(e.target.value as 'modified' | 'name')} aria-label="Sort">
            <option value="modified">Last modified</option>
            <option value="name">Name</option>
          </select>
          <Button variant="outline" size="md" onClick={() => setTrash((t) => !t)}>
            <Icon name={trash ? 'file' : 'trash'} /> {trash ? 'Back to files' : `Trash${trashCount ? ` (${trashCount})` : ''}`}
          </Button>
        </div>
        <div className="overflow-hidden rounded-xl border border-[var(--border)]">
          <table className="w-full text-[13px]">
            <thead className="bg-[var(--panel)] text-left text-[12px] text-[var(--muted)]">
              <tr>
                <th className="px-3 py-2 font-medium">Name</th>
                <th className="hidden px-3 py-2 font-medium sm:table-cell">Owner</th>
                <th className="px-3 py-2 font-medium">Modified</th>
                <th className="w-10" />
              </tr>
            </thead>
            <tbody>
              {items === null && (
                <tr>
                  <td colSpan={4} className="px-3 py-6 text-center text-[var(--muted)]">
                    Loading…
                  </td>
                </tr>
              )}
              {items !== null && shown.length === 0 && (
                <tr>
                  <td colSpan={4} className="px-3 py-8 text-center text-[var(--muted)]">
                    {trash ? 'Trash is empty.' : q ? 'No workbooks match your search.' : 'No workbooks yet. Create a blank workbook or open a file.'}
                  </td>
                </tr>
              )}
              {shown.map((h) => (
                <tr
                  key={h.id}
                  className="cursor-default border-t border-[var(--border)] hover:bg-[var(--hover)]"
                  onClick={() => !trash && onOpen(h.id)}
                  onContextMenu={(e) => {
                    e.preventDefault()
                    setMenu({ x: e.clientX, y: e.clientY, h })
                  }}
                  data-testid="workbook-row"
                >
                  <td className="px-3 py-2">
                    <span className="flex items-center gap-2">
                      <Icon name="sheet" className="shrink-0 text-[var(--accent)]" />
                      <span className="truncate font-medium">{h.name}</span>
                    </span>
                  </td>
                  <td className="hidden px-3 py-2 text-[var(--muted)] sm:table-cell">{h.ownerName}</td>
                  <td className="px-3 py-2 text-[var(--muted)]" title={new Date(h.updatedAt).toLocaleString()}>
                    {relativeTime(h.updatedAt)}
                  </td>
                  <td className="px-1">
                    <Button
                      size="icon"
                      aria-label={`Actions for ${h.name}`}
                      onClick={(e) => {
                        e.stopPropagation()
                        const r = (e.currentTarget as HTMLElement).getBoundingClientRect()
                        setMenu({ x: r.left - 160, y: r.bottom, h })
                      }}
                    >
                      <Icon name="more" />
                    </Button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </main>
      <input
        ref={fileRef}
        type="file"
        className="hidden"
        accept=".xlsx,.xlsm,.csv,.tsv,.txt,.xls,.ods"
        onChange={(e) => {
          const f = e.target.files?.[0]
          e.target.value = ''
          if (f) onImport(f)
        }}
      />
      {menu && (
        <ContextMenu x={menu.x} y={menu.y} onClose={() => setMenu(null)}>
          {menu.h.deletedAt ? (
            <>
              <MenuItem icon="restore" onSelect={() => run(store.restore(menu.h.id))}>
                Restore
              </MenuItem>
              <MenuItem icon="trash" danger onSelect={() => run(store.purge(menu.h.id))}>
                Delete forever
              </MenuItem>
            </>
          ) : (
            <>
              <MenuItem icon="file" onSelect={() => onOpen(menu.h.id)}>
                Open
              </MenuItem>
              <MenuItem
                icon="edit"
                onSelect={() => {
                  setRename(menu.h)
                  setRenameText(menu.h.name)
                }}
              >
                Rename
              </MenuItem>
              <MenuItem icon="dup" onSelect={() => run(onDuplicate(menu.h.id))}>
                Make a copy
              </MenuItem>
              <MenuItem icon="download" onSelect={() => void onDownload(menu.h.id).catch((e) => setError(String(e)))}>
                Download .xlsx
              </MenuItem>
              <MenuSep />
              <MenuItem icon="trash" danger onSelect={() => run(store.trash(menu.h.id))}>
                Move to trash
              </MenuItem>
            </>
          )}
        </ContextMenu>
      )}
      {rename && (
        <Dialog
          title="Rename workbook"
          onClose={() => setRename(null)}
          width={380}
          footer={
            <>
              <Button variant="outline" onClick={() => setRename(null)}>
                Cancel
              </Button>
              <Button
                variant="default"
                disabled={!renameText.trim()}
                onClick={() => {
                  run(store.rename(rename.id, renameText.trim()))
                  setRename(null)
                }}
              >
                Rename
              </Button>
            </>
          }
        >
          <Field label="Name">
            <input className={inputClass} value={renameText} onChange={(e) => setRenameText(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && renameText.trim() && (run(store.rename(rename.id, renameText.trim())), setRename(null))} />
          </Field>
        </Dialog>
      )}
    </div>
  )
}
