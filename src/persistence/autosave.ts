/**
 * Autosave controller: debounces workbook changes, writes only dirty sheets, reports honest
 * status ("Saved" only after the store confirms), retries transient failures with backoff and
 * surfaces conflicts instead of overwriting newer data.
 */
import type { Workbook } from '../engine/workbook'
import { ConflictError, type WorkbookStore } from './store'

export type SaveStatus = 'saved' | 'unsaved' | 'saving' | 'error' | 'conflict'

export interface SaveState {
  status: SaveStatus
  lastSavedAt?: Date
  error?: string
  conflictRevision?: number
}

export const LOCAL_OWNER = { id: 'local', name: 'You (this browser)' }

export class Autosave {
  private timer: ReturnType<typeof setTimeout> | null = null
  private inFlight: Promise<void> | null = null
  private pending = false
  private retryDelay = 1000
  private knownSheetIds: Set<string>
  private unsub: () => void
  state: SaveState
  revision: number
  private listeners = new Set<(s: SaveState) => void>()

  constructor(
    private wb: Workbook,
    private store: WorkbookStore,
    opts: { revision: number; debounceMs?: number; persisted: boolean },
    private debounceMs = opts.debounceMs ?? 800,
  ) {
    this.revision = opts.revision
    this.knownSheetIds = new Set(opts.persisted ? wb.order : [])
    this.state = { status: opts.persisted ? 'saved' : 'unsaved' }
    if (!opts.persisted) wb.markAllDirty()
    else wb.takeDirty()
    this.unsub = wb.subscribe((e) => {
      if (e.reason === 'active') {
        this.schedule() // remember the active sheet too, but don't flag as unsaved
        return
      }
      this.setState({ status: this.state.status === 'conflict' ? 'conflict' : 'unsaved' })
      this.schedule()
    })
    if (!opts.persisted) this.schedule(0)
  }

  onChange(fn: (s: SaveState) => void) {
    this.listeners.add(fn)
    return () => this.listeners.delete(fn)
  }

  private setState(s: Partial<SaveState>) {
    this.state = { ...this.state, ...s }
    for (const l of this.listeners) l(this.state)
  }

  get hasUnsavedChanges() {
    return this.state.status !== 'saved'
  }

  schedule(delay = this.debounceMs) {
    if (this.state.status === 'conflict') return
    if (this.timer) clearTimeout(this.timer)
    this.timer = setTimeout(() => void this.flush(), delay)
  }

  /** Saves now (Ctrl+S). Resolves when the save finished (or failed). */
  async flush(): Promise<void> {
    if (this.timer) {
      clearTimeout(this.timer)
      this.timer = null
    }
    if (this.inFlight) {
      this.pending = true
      return this.inFlight
    }
    if (this.state.status === 'conflict') return
    this.inFlight = this.doSave().finally(() => {
      this.inFlight = null
      if (this.pending) {
        this.pending = false
        void this.flush()
      }
    })
    return this.inFlight
  }

  private async doSave() {
    const dirty = this.wb.takeDirty()
    const ids = this.wb.order
    const toWrite = dirty.all ? ids : dirty.sheets.filter((id) => ids.includes(id))
    const removed = [...this.knownSheetIds].filter((id) => !ids.includes(id))
    this.setState({ status: 'saving', error: undefined })
    try {
      const sheets = toWrite.map((id) => this.wb.serializeSheet(id))
      const sizeHint = ids.reduce((n, id) => {
        const d = this.wb.hf.getSheetDimensions(this.wb.hfId(id))
        return n + d.width * d.height
      }, 0)
      this.revision = await this.store.save({
        header: {
          id: this.wb.id,
          name: this.wb.name,
          ownerId: LOCAL_OWNER.id,
          ownerName: LOCAL_OWNER.name,
          sheetOrder: [...ids],
          activeSheetId: this.wb.activeSheetId,
          namedRanges: this.wb.namedRanges(),
          createdAt: this.wb.createdAt,
          deletedAt: null,
          sizeHint,
        },
        sheets,
        removedSheetIds: removed,
        baseRevision: this.revision,
      })
      this.knownSheetIds = new Set(ids)
      this.retryDelay = 1000
      // a change may have arrived while saving
      const stillDirty = this.wb.takeDirty()
      if (stillDirty.all || stillDirty.sheets.length) {
        // put them back and save again
        if (stillDirty.all) this.wb.markAllDirty()
        else this.requeue(stillDirty.sheets)
        this.setState({ status: 'unsaved' })
        this.schedule(200)
      } else {
        this.setState({ status: 'saved', lastSavedAt: new Date() })
      }
    } catch (e) {
      // put the unsaved sheets back so the next attempt includes them
      if (dirty.all) this.wb.markAllDirty()
      else this.requeue(dirty.sheets)
      if (e instanceof ConflictError) {
        this.setState({ status: 'conflict', error: e.message, conflictRevision: e.storedRevision })
        return
      }
      const msg = e instanceof Error ? e.message : String(e)
      const quota = /quota/i.test(msg)
      this.setState({
        status: 'error',
        error: quota ? 'Browser storage is full. Export your workbook to keep a copy, then free space.' : `Couldn't save: ${msg}. Retrying…`,
      })
      if (!quota) {
        const d = this.retryDelay
        this.retryDelay = Math.min(this.retryDelay * 2, 30000)
        this.schedule(d)
      }
    }
  }

  private requeue(ids: string[]) {
    // Workbook exposes only "mark all"; marking all is safe (just writes more).
    if (ids.length) this.wb.markAllDirty()
  }

  /** After a conflict: overwrite the stored copy with this tab's state. */
  async overwrite(storedRevision: number) {
    this.revision = storedRevision
    this.wb.markAllDirty()
    this.setState({ status: 'unsaved', error: undefined })
    await this.flush()
  }

  dispose() {
    if (this.timer) clearTimeout(this.timer)
    this.unsub()
    this.listeners.clear()
  }
}
