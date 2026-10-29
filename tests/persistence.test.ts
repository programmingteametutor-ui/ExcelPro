import { test } from 'node:test'
import assert from 'node:assert/strict'
import { Workbook } from '../src/engine/workbook'
import { Autosave } from '../src/persistence/autosave'
import { ConflictError, MemoryStore } from '../src/persistence/store'

const tick = (ms: number) => new Promise((r) => setTimeout(r, ms))

test('autosave writes after debounce; reload restores contents, formats and sheet order', async () => {
  const store = new MemoryStore()
  const wb = Workbook.blank('Budget')
  const auto = new Autosave(wb, store, { revision: 0, persisted: false, debounceMs: 20 })
  await auto.flush()
  assert.equal(auto.state.status, 'saved')
  const s = wb.order[0]
  wb.setInputs(s, [{ r: 0, c: 0, input: '5' }, { r: 0, c: 1, input: '=A1*3' }])
  assert.equal(auto.state.status, 'unsaved')
  wb.applyStyle(s, [{ r1: 0, c1: 0, r2: 0, c2: 0 }], { bold: true })
  const s2 = wb.addSheet('Notes')
  wb.moveSheet(s2, 0)
  await tick(80)
  assert.equal(auto.state.status, 'saved')
  const loaded = await store.load(wb.id)
  assert.ok(loaded)
  const wb2 = Workbook.fromDoc(loaded!.doc)
  assert.deepEqual(wb2.sheets.map((x) => x.name), ['Notes', 'Sheet1'])
  assert.equal(wb2.getValue(s, 0, 1), 15)
  assert.equal(wb2.getStyle(s, 0, 0)?.bold, true)
  auto.dispose()
})

test('only dirty sheets are rewritten', async () => {
  const store = new MemoryStore()
  const wb = Workbook.blank('B', ['A', 'B'])
  const auto = new Autosave(wb, store, { revision: 0, persisted: false, debounceMs: 5 })
  await auto.flush()
  let written: string[] = []
  const orig = store.save.bind(store)
  store.save = async (r) => {
    written = r.sheets.map((x) => x.name)
    return orig(r)
  }
  wb.setInputs(wb.order[1], [{ r: 0, c: 0, input: 'x' }])
  await auto.flush()
  assert.deepEqual(written, ['B'])
  auto.dispose()
})

test('"Saved" is only shown after persistence succeeds; failures retry', async () => {
  const store = new MemoryStore()
  const wb = Workbook.blank('B')
  const auto = new Autosave(wb, store, { revision: 0, persisted: false, debounceMs: 5 })
  await auto.flush()
  store.failNext = 1
  wb.setInputs(wb.order[0], [{ r: 0, c: 0, input: '1' }])
  await auto.flush()
  assert.equal(auto.state.status, 'error')
  await tick(1200) // backoff retry
  assert.equal(auto.state.status, 'saved')
  const loaded = await store.load(wb.id)
  assert.equal(Workbook.fromDoc(loaded!.doc).getValue(wb.order[0], 0, 0), 1)
  auto.dispose()
})

test('stale tabs do not overwrite newer saves (conflict detection)', async () => {
  const store = new MemoryStore()
  const wb = Workbook.blank('B')
  const a1 = new Autosave(wb, store, { revision: 0, persisted: false, debounceMs: 5 })
  await a1.flush()
  // second "tab" opens the same workbook and saves a change
  const loaded = await store.load(wb.id)
  const tab2 = Workbook.fromDoc(loaded!.doc)
  const a2 = new Autosave(tab2, store, { revision: loaded!.revision, persisted: true, debounceMs: 5 })
  tab2.setInputs(tab2.order[0], [{ r: 0, c: 0, input: 'from tab 2' }])
  await a2.flush()
  // first tab edits based on the old revision
  wb.setInputs(wb.order[0], [{ r: 0, c: 0, input: 'from tab 1' }])
  await a1.flush()
  assert.equal(a1.state.status, 'conflict')
  const now = await store.load(wb.id)
  assert.equal(Workbook.fromDoc(now!.doc).getValue(wb.order[0], 0, 0), 'from tab 2')
  // user chooses to keep tab 1's version
  await a1.overwrite(a1.state.conflictRevision!)
  assert.equal(a1.state.status, 'saved')
  const after = await store.load(wb.id)
  assert.equal(Workbook.fromDoc(after!.doc).getValue(wb.order[0], 0, 0), 'from tab 1')
  await assert.rejects(store.save({ header: { ...(await store.list())[0] }, sheets: [], removedSheetIds: [], baseRevision: 0 }), ConflictError)
  a1.dispose()
  a2.dispose()
})

test('deleted sheets are removed from storage', async () => {
  const store = new MemoryStore()
  const wb = Workbook.blank('B', ['A', 'B'])
  const auto = new Autosave(wb, store, { revision: 0, persisted: false, debounceMs: 5 })
  await auto.flush()
  wb.removeSheet(wb.order[1])
  await auto.flush()
  const loaded = await store.load(wb.id)
  assert.equal(loaded!.doc.sheets.length, 1)
  auto.dispose()
})
