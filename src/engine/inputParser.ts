/**
 * Turns what a user types into a cell into engine content, Excel-style:
 * "1,234.5" -> 1234.5 (+ number format), "12%" -> 0.12 (+ % format), "$5" -> 5 (+ currency),
 * "2024-03-01" -> date serial (+ date format), "true" -> TRUE, "=sum(a1:a3" -> "=SUM(A1:A3)".
 */
import { toEngineFormula } from './formulaText'
import { partsToSerial } from './numberFormat'

export type RawContent = string | number | boolean | null

export type DateOrder = 'MDY' | 'DMY'

export interface ParsedInput {
  raw: RawContent
  /** Number format Excel would apply automatically, if the cell has none yet. */
  autoFormat?: string
}

const MONTHS: Record<string, number> = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12,
  january: 1, february: 2, march: 3, april: 4, june: 6, july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
}

function validDate(y: number, m: number, d: number): boolean {
  if (m < 1 || m > 12 || d < 1 || y < 1900 || y > 9999) return false
  const dim = new Date(Date.UTC(y, m, 0)).getUTCDate()
  return d <= dim
}

function fullYear(y: string): number {
  const n = parseInt(y, 10)
  if (y.length <= 2) return n < 30 ? 2000 + n : 1900 + n
  return n
}

function parseTime(t: string): number | null {
  const m = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*(am|pm|a|p)?$/i.exec(t.trim())
  if (!m) return null
  let h = parseInt(m[1], 10)
  const mi = parseInt(m[2], 10)
  const s = m[3] ? parseFloat(m[3]) : 0
  if (m[4]) {
    if (h < 1 || h > 12) return null
    const pm = m[4].toLowerCase().startsWith('p')
    h = (h % 12) + (pm ? 12 : 0)
  }
  if (h > 23 || mi > 59 || s >= 60) return null
  return (h * 3600 + mi * 60 + s) / 86400
}

export function parseDateTime(text: string, order: DateOrder = 'MDY'): { serial: number; format: string } | null {
  const t = text.trim()
  if (/^\d{1,2}:\d{2}/.test(t)) {
    const tv = parseTime(t)
    if (tv === null) return null
    return { serial: tv, format: /:\d{2}:\d{2}/.test(t) ? 'h:mm:ss' : /am|pm|\b[ap]$/i.test(t) ? 'h:mm AM/PM' : 'h:mm' }
  }
  // split date and time parts
  const dm = /^(.*?)(?:[ T]+(\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?\s*(?:am|pm|a|p)?))?$/i.exec(t)
  if (!dm) return null
  const datePart = dm[1]
  const timePart = dm[2]
  let ymd: [number, number, number] | null = null
  let fmt = ''
  let m: RegExpExecArray | null
  if ((m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(datePart))) {
    ymd = [+m[1], +m[2], +m[3]]
    fmt = 'yyyy-mm-dd'
  } else if ((m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/.exec(datePart))) {
    const a = +m[1]
    const b = +m[2]
    const y = fullYear(m[3])
    ymd = order === 'MDY' ? [y, a, b] : [y, b, a]
    fmt = order === 'MDY' ? 'm/d/yyyy' : 'dd/mm/yyyy'
  } else if ((m = /^(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s,]*(\d{2,4})$/.exec(datePart)) && MONTHS[m[2].toLowerCase()]) {
    ymd = [fullYear(m[3]), MONTHS[m[2].toLowerCase()], +m[1]]
    fmt = 'd-mmm-yyyy'
  } else if ((m = /^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})$/.exec(datePart)) && MONTHS[m[1].toLowerCase()]) {
    ymd = [+m[3], MONTHS[m[1].toLowerCase()], +m[2]]
    fmt = 'mmmm d, yyyy'
  } else if (datePart === '' && timePart) {
    const tv = parseTime(timePart)
    if (tv === null) return null
    return { serial: tv, format: /:\d{2}:\d{2}/.test(timePart) ? 'h:mm:ss' : /am|pm/i.test(timePart) ? 'h:mm AM/PM' : 'h:mm' }
  } else return null
  if (!ymd || !validDate(ymd[0], ymd[1], ymd[2])) return null
  let serial = partsToSerial(ymd[0], ymd[1], ymd[2])
  if (timePart) {
    const tv = parseTime(timePart)
    if (tv === null) return null
    serial += tv
    fmt += ' h:mm'
  }
  return { serial, format: fmt }
}

/** Upper-cases function names and cell references and closes unbalanced parentheses. */
export function normalizeFormula(f: string): string {
  let out = ''
  let i = 1
  let depth = 0
  out = '='
  while (i < f.length) {
    const ch = f[i]
    if (ch === '"' || ch === "'") {
      let j = i + 1
      while (j < f.length) {
        if (f[j] === ch) {
          if (f[j + 1] === ch) j += 2
          else break
        } else j++
      }
      out += f.slice(i, j + 1)
      i = j + 1
      continue
    }
    const m = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(f.slice(i))
    if (m) {
      const word = m[0]
      const after = f.slice(i + word.length)
      const prev = out[out.length - 1] ?? ''
      const isFunc = /^\s*\(/.test(after)
      const isRef = /^\$?[A-Za-z]{1,3}\$?\d+$/.test(word) || /^(true|false)$/i.test(word)
      const isSheetPrefix = after.startsWith('!')
      if (!isSheetPrefix && !/[A-Za-z0-9_]/.test(prev) && (isFunc || isRef)) out += word.toUpperCase()
      else out += word
      i += word.length
      continue
    }
    if (ch === '(') depth++
    if (ch === ')') depth--
    out += ch
    i++
  }
  // a "$" inside a token was skipped by the word regex; handle "$a$1" style refs
  out = out.replace(/(^|[^A-Za-z0-9_])\$([a-z]{1,3})\$?(\d+)/g, (s) => s.toUpperCase())
  while (depth-- > 0) out += ')'
  return toEngineFormula(out)
}

const CURRENCY = /^([+-])?\s*([$₹€£¥])\s*([+-])?\s*(\d{1,3}(?:,\d{3})+|\d+)?(\.\d+)?$/
const PLAIN_NUMBER = /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i
const THOUSANDS = /^([+-])?(\d{1,3}(?:,\d{3})+)(\.\d+)?$/
const PERCENT = /^([+-]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?)\s*%$/

/**
 * @param input text as typed
 * @param currentFormat the cell's current number format (affects "@" text cells)
 */
export function parseInput(input: string, currentFormat?: string, order: DateOrder = 'MDY'): ParsedInput {
  if (input === '') return { raw: null }
  if (currentFormat === '@') return { raw: input.startsWith('=') ? `'${input}` : `'${input}` }
  if (input.startsWith('=') && input.length > 1) return { raw: normalizeFormula(input) }
  if (input.startsWith("'")) return { raw: input } // explicit text
  const t = input.trim()
  if (PLAIN_NUMBER.test(t)) return { raw: Number(t) }
  let m: RegExpExecArray | null
  if ((m = THOUSANDS.exec(t))) {
    const v = Number((m[1] ?? '') + m[2].replace(/,/g, '') + (m[3] ?? ''))
    return { raw: v, autoFormat: m[3] ? '#,##0.00' : '#,##0' }
  }
  if ((m = PERCENT.exec(t)) && m[1] !== '' && m[1] !== '+' && m[1] !== '-') {
    const n = Number(m[1].replace(/,/g, ''))
    if (!isNaN(n)) return { raw: n / 100, autoFormat: /\./.test(m[1]) ? '0.00%' : '0%' }
  }
  if ((m = CURRENCY.exec(t)) && (m[4] || m[5])) {
    const sign = m[1] === '-' || m[3] === '-' ? -1 : 1
    const v = sign * Number((m[4] ?? '0').replace(/,/g, '') + (m[5] ?? ''))
    const sym = m[2]
    return { raw: v, autoFormat: m[5] ? `${sym}#,##0.00` : `${sym}#,##0` }
  }
  if (/^(true|false)$/i.test(t)) return { raw: t.toUpperCase() === 'TRUE' }
  const dt = parseDateTime(t, order)
  if (dt) return { raw: dt.serial, autoFormat: dt.format }
  return { raw: input }
}

/** Whether a string, if typed back in, would be converted to something other than text. */
export function wouldConvert(s: string, order: DateOrder = 'MDY'): boolean {
  if (s.startsWith('=') && s.length > 1) return true
  if (s.startsWith("'")) return true
  const p = parseInput(s, undefined, order)
  return typeof p.raw !== 'string'
}
