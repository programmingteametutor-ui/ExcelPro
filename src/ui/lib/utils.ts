import { clsx, type ClassValue } from 'clsx'
import { twMerge } from 'tailwind-merge'

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs))
}

export const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent)

/** Primary modifier: Cmd on macOS, Ctrl elsewhere. */
export const modKey = (e: { ctrlKey: boolean; metaKey: boolean }) => (isMac ? e.metaKey : e.ctrlKey)

export const shortcut = (s: string) => (isMac ? s.replace(/Ctrl\+/g, '⌘').replace(/Alt\+/g, '⌥').replace(/Shift\+/g, '⇧') : s)

interface HostDownloads {
  save(req: { filename: string; data: Blob | string }): Promise<unknown>
}

/**
 * Saves a generated file. Inside a claude.ai artifact frame (where plain downloads are blocked) it
 * uses the host's `downloads` capability; everywhere else it uses a normal browser download.
 */
export async function download(data: Uint8Array | string, fileName: string, mime: string): Promise<void> {
  const blob = new Blob([data as BlobPart], { type: mime })
  const host = (window as unknown as { claude?: { use?: (n: string) => Promise<HostDownloads | null> } }).claude
  if (host?.use) {
    const dl = await host.use('downloads').catch(() => null)
    if (dl) {
      try {
        await dl.save({ filename: fileName, data: blob })
        return
      } catch (e) {
        const code = (e as { code?: string })?.code
        if (code === 'declined') throw new Error('Download cancelled.')
        if (code === 'rejected_extension') throw new Error('This file type cannot be downloaded here; choose .xlsx or .csv.')
        throw new Error((e as { message?: string })?.message ?? 'The download was blocked.')
      }
    }
  }
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 2000)
}

export function safeFileName(name: string) {
  return name.replace(/[\\/:*?"<>|]+/g, '_').trim() || 'Workbook'
}

export function relativeTime(iso: string): string {
  const d = new Date(iso)
  const diff = (Date.now() - d.getTime()) / 1000
  if (diff < 60) return 'just now'
  if (diff < 3600) return `${Math.floor(diff / 60)} min ago`
  if (diff < 86400) return `${Math.floor(diff / 3600)} h ago`
  if (diff < 86400 * 7) return `${Math.floor(diff / 86400)} d ago`
  return d.toLocaleDateString()
}

export function storageGet(key: string): string | null {
  try {
    return localStorage.getItem(key)
  } catch {
    return null
  }
}
export function storageSet(key: string, value: string) {
  try {
    localStorage.setItem(key, value)
  } catch {
    /* storage unavailable */
  }
}
