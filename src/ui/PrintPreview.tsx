/* Print preview with page navigation, PDF download and browser printing. */
import { useEffect, useMemo, useState } from 'react'
import type { Workbook } from '../engine/workbook'
import { layoutPages } from '../print/paginate'
import { chartImages, renderPage } from '../print/render'
import { Button, Check } from './components/ui'
import { Icon } from './icons'

export function PrintPreviewDialog({ wb, sheetId, version, onClose, onPageSetup, onPdf }: { wb: Workbook; sheetId: string; version: number; onClose: () => void; onPageSetup: () => void; onPdf: (allSheets: boolean) => Promise<void> }) {
  const layout = useMemo(() => layoutPages(wb, sheetId), [wb, sheetId, version])
  const [page, setPage] = useState(0)
  const [img, setImg] = useState<string | null>(null)
  const [images, setImages] = useState<Map<string, HTMLImageElement> | null>(null)
  const [busy, setBusy] = useState<string | null>(null)
  const [all, setAll] = useState(false)
  const n = layout?.pages.length ?? 0
  useEffect(() => {
    let live = true
    void chartImages(wb, sheetId).then((m) => live && setImages(m))
    return () => void (live = false)
  }, [wb, sheetId, version])
  useEffect(() => {
    if (!layout || !images || !layout.pages[page]) return setImg(null)
    const c = renderPage(wb, sheetId, layout, layout.pages[Math.min(page, n - 1)], { pxPerPt: 1.6, fileName: wb.name, images })
    setImg(c.toDataURL('image/png'))
  }, [layout, page, images, wb, sheetId, n])
  useEffect(() => {
    if (page >= n && n) setPage(n - 1)
  }, [n, page])
  useEffect(() => {
    const k = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
      if (e.key === 'ArrowRight' || e.key === 'PageDown') setPage((p) => Math.min(n - 1, p + 1))
      if (e.key === 'ArrowLeft' || e.key === 'PageUp') setPage((p) => Math.max(0, p - 1))
    }
    window.addEventListener('keydown', k)
    return () => window.removeEventListener('keydown', k)
  }, [n, onClose])

  const print = async () => {
    if (!layout || !images) return
    setBusy('Preparing pages…')
    try {
      const urls: string[] = []
      for (const p of layout.pages) {
        urls.push(renderPage(wb, sheetId, layout, p, { pxPerPt: 2.2, fileName: wb.name, images }).toDataURL('image/png'))
        await new Promise((r) => setTimeout(r, 0))
      }
      const frame = document.createElement('iframe')
      frame.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0'
      document.body.appendChild(frame)
      const d = frame.contentDocument!
      const size = `${layout.pageW}pt ${layout.pageH}pt`
      d.open()
      d.write(`<!doctype html><html><head><title>${wb.name.replace(/</g, '')}</title><style>@page{size:${size};margin:0}html,body{margin:0}img{display:block;width:${layout.pageW}pt;height:${layout.pageH}pt;page-break-after:always}</style></head><body>${urls.map((u) => `<img src="${u}">`).join('')}</body></html>`)
      d.close()
      await new Promise((r) => setTimeout(r, 300))
      try {
        frame.contentWindow!.focus()
        frame.contentWindow!.print()
      } finally {
        setTimeout(() => frame.remove(), 2000)
      }
    } catch {
      setBusy('This browser blocked printing here. Download the PDF and print it instead.')
      return
    }
    setBusy(null)
  }

  return (
    <div className="fixed inset-0 z-[90] flex flex-col bg-[var(--panel)]" role="dialog" aria-modal="true" aria-label="Print preview" data-testid="print-preview">
      <div className="flex h-12 shrink-0 flex-wrap items-center gap-2 border-b border-[var(--border)] bg-[var(--bg)] px-3">
        <Button size="sm" variant="ghost" onClick={onClose}><Icon name="chevronLeft" size={15} /> Back</Button>
        <span className="font-semibold">Print preview</span>
        <span className="text-[12px] text-[var(--muted)]">{wb.meta(sheetId).name}</span>
        <div className="ml-auto flex items-center gap-2">
          <Button size="sm" variant="outline" onClick={onPageSetup}><Icon name="page" size={14} /> Page setup</Button>
          <Check label="All sheets in PDF" checked={all} onChange={setAll} />
          <Button size="sm" variant="outline" disabled={!n || !!busy} onClick={() => {
            setBusy('Creating PDF…')
            void onPdf(all).finally(() => setBusy(null))
          }}><Icon name="download" size={14} /> Download PDF</Button>
          <Button size="sm" variant="default" disabled={!n || !!busy} onClick={() => void print()}><Icon name="print" size={14} /> Print</Button>
        </div>
      </div>
      <div className="flex min-h-0 flex-1 items-center justify-center overflow-auto p-6">
        {!layout || !n ? (
          <p className="text-[var(--muted)]">There is nothing to print on this sheet. Enter some data or set a print area.</p>
        ) : img ? (
          <img src={img} alt={`Page ${page + 1}`} className="max-h-full max-w-full bg-white shadow-xl" style={{ aspectRatio: `${layout.pageW} / ${layout.pageH}` }} data-testid="preview-page" />
        ) : (
          <p className="text-[var(--muted)]">Rendering…</p>
        )}
      </div>
      <div className="flex h-10 shrink-0 items-center justify-center gap-3 border-t border-[var(--border)] bg-[var(--bg)] text-[12.5px]">
        <Button size="icon" aria-label="Previous page" disabled={page === 0} onClick={() => setPage(page - 1)}><Icon name="chevronLeft" size={15} /></Button>
        <span data-testid="page-count">Page {n ? page + 1 : 0} of {n}</span>
        <Button size="icon" aria-label="Next page" disabled={page >= n - 1} onClick={() => setPage(page + 1)}><Icon name="chevronRight" size={15} /></Button>
        {layout && <span className="text-[var(--muted)]">{layout.setup.paper} · {layout.setup.orientation} · {Math.round(layout.scale * 100)}%</span>}
        {busy && <span role="status" className="text-[var(--muted)]">{busy}</span>}
      </div>
    </div>
  )
}
