/**
 * Persistence interface + IndexedDB implementation.
 *
 * Workbooks are stored as one header record plus one record per sheet, so an edit to one sheet
 * rewrites only that sheet. Every header carries a revision number; a save is rejected with a
 * ConflictError if the stored revision is newer than the one this tab loaded (another tab or
 * device saved in between), so stale state never silently overwrites newer work.
 *
 * The Supabase adapter (Phase 4) implements the same WorkbookStore interface; its schema and
 * Row Level Security policies are in supabase/migrations.
 */
import type { NamedRangeDoc, SheetDoc, WorkbookDoc } from '../model/types'

export interface WorkbookHeader {
  id: string
  name: string
  ownerId: string
  ownerName: string
  revision: number
  sheetOrder: string[]
  activeSheetId: string
  namedRanges: NamedRangeDoc[]
  createdAt: string
  updatedAt: string
  deletedAt: string | null
  sizeHint: number // number of cells, for the file list
}

export interface SaveRequest {
  header: Omit<WorkbookHeader, 'revision' | 'updatedAt'>
  /** sheets to (re)write; omitted sheets are kept as stored */
  sheets: SheetDoc[]
  /** sheet ids that no longer exist */
  removedSheetIds: string[]
  baseRevision: number
}

export class ConflictError extends Error {
  constructor(public storedRevision: number) {
    super('This workbook was changed in another tab or window since you opened it.')
    this.name = 'ConflictError'
  }
}

export interface WorkbookStore {
  list(includeDeleted?: boolean): Promise<WorkbookHeader[]>
  load(id: string): Promise<{ doc: WorkbookDoc; revision: number } | null>
  save(req: SaveRequest): Promise<number> // returns new revision
  rename(id: string, name: string): Promise<void>
  trash(id: string): Promise<void>
  restore(id: string): Promise<void>
  purge(id: string): Promise<void>
}

const DB_NAME = 'excelpro'
const DB_VERSION = 1

function req<T>(r: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    r.onsuccess = () => resolve(r.result)
    r.onerror = () => reject(r.error)
  })
}

function txDone(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error ?? new Error('Transaction failed'))
    tx.onabort = () => reject(tx.error ?? new Error('Transaction aborted'))
  })
}

export class IndexedDbStore implements WorkbookStore {
  private dbp: Promise<IDBDatabase> | null = null

  constructor(private factory: IDBFactory = indexedDB) {}

  private db(): Promise<IDBDatabase> {
    if (!this.dbp) {
      this.dbp = new Promise((resolve, reject) => {
        const open = this.factory.open(DB_NAME, DB_VERSION)
        open.onupgradeneeded = () => {
          const db = open.result
          if (!db.objectStoreNames.contains('workbooks')) db.createObjectStore('workbooks', { keyPath: 'id' })
          if (!db.objectStoreNames.contains('sheets')) {
            const s = db.createObjectStore('sheets', { keyPath: ['workbookId', 'id'] })
            s.createIndex('byWorkbook', 'workbookId')
          }
        }
        open.onsuccess = () => resolve(open.result)
        open.onerror = () => reject(open.error ?? new Error('Could not open browser storage'))
        open.onblocked = () => reject(new Error('Browser storage is blocked by another open ExcelPro tab being upgraded.'))
      })
      this.dbp.catch(() => (this.dbp = null))
    }
    return this.dbp
  }

  async list(includeDeleted = false): Promise<WorkbookHeader[]> {
    const db = await this.db()
    const all = await req(db.transaction('workbooks').objectStore('workbooks').getAll() as IDBRequest<WorkbookHeader[]>)
    return all.filter((h) => includeDeleted || !h.deletedAt).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))
  }

  async load(id: string) {
    const db = await this.db()
    const tx = db.transaction(['workbooks', 'sheets'])
    const header = await req(tx.objectStore('workbooks').get(id) as IDBRequest<WorkbookHeader | undefined>)
    if (!header) return null
    const rows = await req(tx.objectStore('sheets').index('byWorkbook').getAll(id) as IDBRequest<(SheetDoc & { workbookId: string })[]>)
    const sheets = rows.map(({ workbookId: _w, ...s }) => {
      void _w
      return s as SheetDoc
    })
    const doc: WorkbookDoc = {
      schemaVersion: 1,
      id: header.id,
      name: header.name,
      activeSheetId: header.activeSheetId,
      sheetOrder: header.sheetOrder,
      sheets,
      namedRanges: header.namedRanges ?? [],
      createdAt: header.createdAt,
      updatedAt: header.updatedAt,
    }
    return { doc, revision: header.revision }
  }

  async save(r: SaveRequest): Promise<number> {
    const db = await this.db()
    const tx = db.transaction(['workbooks', 'sheets'], 'readwrite')
    const wbs = tx.objectStore('workbooks')
    const current = await req(wbs.get(r.header.id) as IDBRequest<WorkbookHeader | undefined>)
    const stored = current?.revision ?? 0
    if (stored !== r.baseRevision) {
      tx.abort()
      throw new ConflictError(stored)
    }
    const revision = stored + 1
    const header: WorkbookHeader = { ...r.header, revision, updatedAt: new Date().toISOString(), deletedAt: current?.deletedAt ?? null }
    wbs.put(header)
    const sheetStore = tx.objectStore('sheets')
    for (const s of r.sheets) sheetStore.put({ ...s, workbookId: r.header.id })
    for (const id of r.removedSheetIds) sheetStore.delete([r.header.id, id])
    await txDone(tx)
    return revision
  }

  private async patch(id: string, fn: (h: WorkbookHeader) => WorkbookHeader) {
    const db = await this.db()
    const tx = db.transaction('workbooks', 'readwrite')
    const s = tx.objectStore('workbooks')
    const h = await req(s.get(id) as IDBRequest<WorkbookHeader | undefined>)
    if (h) s.put(fn(h))
    await txDone(tx)
  }

  rename(id: string, name: string) {
    return this.patch(id, (h) => ({ ...h, name, revision: h.revision + 1, updatedAt: new Date().toISOString() }))
  }
  trash(id: string) {
    return this.patch(id, (h) => ({ ...h, deletedAt: new Date().toISOString() }))
  }
  restore(id: string) {
    return this.patch(id, (h) => ({ ...h, deletedAt: null }))
  }
  async purge(id: string) {
    const db = await this.db()
    const tx = db.transaction(['workbooks', 'sheets'], 'readwrite')
    tx.objectStore('workbooks').delete(id)
    const keys = await req(tx.objectStore('sheets').index('byWorkbook').getAllKeys(id))
    for (const k of keys) tx.objectStore('sheets').delete(k)
    await txDone(tx)
  }
}

/** In-memory store (tests, or when the browser blocks IndexedDB). */
export class MemoryStore implements WorkbookStore {
  private headers = new Map<string, WorkbookHeader>()
  private sheets = new Map<string, Map<string, SheetDoc>>()
  failNext = 0

  async list(includeDeleted = false) {
    return [...this.headers.values()].filter((h) => includeDeleted || !h.deletedAt).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))
  }
  async load(id: string) {
    const h = this.headers.get(id)
    if (!h) return null
    const sheets = [...(this.sheets.get(id)?.values() ?? [])].map((s) => structuredClone(s))
    return {
      doc: { schemaVersion: 1 as const, id, name: h.name, activeSheetId: h.activeSheetId, sheetOrder: [...h.sheetOrder], sheets, namedRanges: h.namedRanges, createdAt: h.createdAt, updatedAt: h.updatedAt },
      revision: h.revision,
    }
  }
  async save(r: SaveRequest) {
    if (this.failNext > 0) {
      this.failNext--
      throw new Error('Simulated storage failure')
    }
    const cur = this.headers.get(r.header.id)
    const stored = cur?.revision ?? 0
    if (stored !== r.baseRevision) throw new ConflictError(stored)
    const revision = stored + 1
    this.headers.set(r.header.id, { ...r.header, revision, updatedAt: new Date().toISOString(), deletedAt: cur?.deletedAt ?? null })
    const m = this.sheets.get(r.header.id) ?? new Map<string, SheetDoc>()
    for (const s of r.sheets) m.set(s.id, structuredClone(s))
    for (const id of r.removedSheetIds) m.delete(id)
    this.sheets.set(r.header.id, m)
    return revision
  }
  async rename(id: string, name: string) {
    const h = this.headers.get(id)
    if (h) this.headers.set(id, { ...h, name, revision: h.revision + 1 })
  }
  async trash(id: string) {
    const h = this.headers.get(id)
    if (h) this.headers.set(id, { ...h, deletedAt: new Date().toISOString() })
  }
  async restore(id: string) {
    const h = this.headers.get(id)
    if (h) this.headers.set(id, { ...h, deletedAt: null })
  }
  async purge(id: string) {
    this.headers.delete(id)
    this.sheets.delete(id)
  }
}
