/** PDF export: each page is rendered at print resolution and placed on a pdf-lib page. */
import { PDFDocument } from 'pdf-lib'
import type { Workbook } from '../engine/workbook'
import { layoutPages, type Layout } from './paginate'
import { chartImages, renderPage } from './render'

export async function exportPdf(wb: Workbook, sheetIds: string[], fileName: string, onProgress?: (done: number, total: number) => void): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  doc.setTitle(fileName)
  doc.setCreator('ExcelPro')
  const layouts: { id: string; layout: Layout }[] = []
  for (const id of sheetIds) {
    const l = layoutPages(wb, id)
    if (l) layouts.push({ id, layout: l })
  }
  const total = layouts.reduce((a, l) => a + l.layout.pages.length, 0)
  if (!total) throw new Error('There is nothing to print on the selected sheet.')
  let done = 0
  for (const { id, layout } of layouts) {
    const images = await chartImages(wb, id)
    for (const page of layout.pages) {
      const canvas = renderPage(wb, id, layout, page, { pxPerPt: 2.2, fileName, images })
      const blob: Blob = await new Promise((resolve, reject) => canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('Rendering failed'))), 'image/png'))
      const png = await doc.embedPng(new Uint8Array(await blob.arrayBuffer()))
      const p = doc.addPage([layout.pageW, layout.pageH])
      p.drawImage(png, { x: 0, y: 0, width: layout.pageW, height: layout.pageH })
      onProgress?.(++done, total)
      // keep the UI responsive between pages
      await new Promise((r) => setTimeout(r, 0))
    }
  }
  return doc.save()
}
