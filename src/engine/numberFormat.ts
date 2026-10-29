/**
 * Excel-style number format rendering (a practical subset of the format-code language):
 * sections (pos;neg;zero;text), General, 0 # ? placeholders, thousands separators and scaling,
 * decimals, %, scientific (E+), fractions (# ?/?), literals ("..", \x, _x), [Color], [$€-409]
 * currency, dates/times (y m d h s, AM/PM, elapsed [h]) and @ text.
 */

export interface Formatted {
  text: string
  color?: string
}

type Tok =
  | { t: 'lit'; v: string }
  | { t: 'dig'; v: '0' | '#' | '?' }
  | { t: 'dot' }
  | { t: 'comma' }
  | { t: 'pct' }
  | { t: 'exp'; v: '+' | '-' }
  | { t: 'slash' }
  | { t: 'date'; v: string }
  | { t: 'ampm'; v: string }
  | { t: 'elapsed'; v: string }
  | { t: 'subsec'; v: number }
  | { t: 'text' }
  | { t: 'general' }

interface Section {
  toks: Tok[]
  color?: string
  isDate: boolean
  cond?: { op: string; val: number }
}

const COLORS: Record<string, string> = {
  black: '#000000', blue: '#0000ff', cyan: '#00ffff', green: '#008000', magenta: '#ff00ff',
  red: '#ff0000', white: '#ffffff', yellow: '#ffff00',
}

const cache = new Map<string, Section[]>()

function splitSections(fmt: string): string[] {
  const out: string[] = []
  let cur = ''
  let inQ = false
  let inB = false
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i]
    if (ch === '\\' && !inQ) {
      cur += ch + (fmt[i + 1] ?? '')
      i++
      continue
    }
    if (ch === '"') inQ = !inQ
    else if (!inQ && ch === '[') inB = true
    else if (!inQ && ch === ']') inB = false
    if (ch === ';' && !inQ && !inB) {
      out.push(cur)
      cur = ''
    } else cur += ch
  }
  out.push(cur)
  return out
}

function parseSection(src: string): Section {
  const toks: Tok[] = []
  let color: string | undefined
  let cond: Section['cond']
  let i = 0
  const lower = src.toLowerCase()
  while (i < src.length) {
    const ch = src[i]
    const lc = lower[i]
    if (ch === '"') {
      const j = src.indexOf('"', i + 1)
      toks.push({ t: 'lit', v: src.slice(i + 1, j < 0 ? undefined : j) })
      i = j < 0 ? src.length : j + 1
    } else if (ch === '\\') {
      toks.push({ t: 'lit', v: src[i + 1] ?? '' })
      i += 2
    } else if (ch === '_') {
      toks.push({ t: 'lit', v: ' ' })
      i += 2
    } else if (ch === '*') {
      // fill character: Excel repeats it to the column width; we render it once
      toks.push({ t: 'lit', v: src[i + 1] ?? '' })
      i += 2
    } else if (ch === '[') {
      const j = src.indexOf(']', i)
      const inner = src.slice(i + 1, j < 0 ? undefined : j)
      const il = inner.toLowerCase()
      if (COLORS[il]) color = COLORS[il]
      else if (/^color\d+$/.test(il)) color = undefined
      else if (inner.startsWith('$')) {
        const sym = inner.slice(1).split('-')[0]
        if (sym) toks.push({ t: 'lit', v: sym })
      } else if (/^(h+|m+|s+)$/.test(il)) toks.push({ t: 'elapsed', v: il })
      else {
        const m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)$/.exec(inner)
        if (m) cond = { op: m[1], val: parseFloat(m[2]) }
      }
      i = j < 0 ? src.length : j + 1
    } else if (lower.startsWith('general', i)) {
      toks.push({ t: 'general' })
      i += 7
    } else if (lower.startsWith('am/pm', i)) {
      toks.push({ t: 'ampm', v: src.slice(i, i + 5) })
      i += 5
    } else if (lower.startsWith('a/p', i)) {
      toks.push({ t: 'ampm', v: src.slice(i, i + 3) })
      i += 3
    } else if (ch === '0' || ch === '#' || ch === '?') {
      toks.push({ t: 'dig', v: ch })
      i++
    } else if (ch === '.') {
      // fractional seconds: ".0", ".00" right after a seconds token
      const prev = [...toks].reverse().find((t) => t.t !== 'lit')
      if (prev && prev.t === 'date' && prev.v.startsWith('s') && src[i + 1] === '0') {
        let n = 0
        while (src[i + 1 + n] === '0') n++
        toks.push({ t: 'subsec', v: n })
        i += 1 + n
      } else {
        toks.push({ t: 'dot' })
        i++
      }
    } else if (ch === ',') {
      toks.push({ t: 'comma' })
      i++
    } else if (ch === '%') {
      toks.push({ t: 'pct' })
      i++
    } else if ((ch === 'E' || ch === 'e') && (src[i + 1] === '+' || src[i + 1] === '-')) {
      toks.push({ t: 'exp', v: src[i + 1] as '+' | '-' })
      i += 2
    } else if (ch === '/') {
      toks.push({ t: 'slash' })
      i++
    } else if (ch === '@') {
      toks.push({ t: 'text' })
      i++
    } else if ('ymdhs'.includes(lc)) {
      let j = i
      while (j < src.length && lower[j] === lc) j++
      toks.push({ t: 'date', v: lower.slice(i, j) })
      i = j
    } else if (lc === 'b' && lower[i + 1] === '2') {
      i += 2 // buddhist calendar flag; ignore
    } else {
      toks.push({ t: 'lit', v: ch })
      i++
    }
  }
  const isDate = toks.some((t) => t.t === 'date' || t.t === 'elapsed' || t.t === 'ampm')
  if (isDate) {
    // resolve m -> minutes when adjacent to h or s
    const sig = toks.map((t, idx) => ({ t, idx })).filter((x) => x.t.t === 'date' || x.t.t === 'elapsed')
    sig.forEach((x, k) => {
      const tk = x.t as { t: 'date'; v: string }
      if (tk.t !== 'date' || tk.v[0] !== 'm' || tk.v.length > 2) return
      const prev = sig[k - 1]?.t as { v?: string } | undefined
      const next = sig[k + 1]?.t as { v?: string } | undefined
      if ((prev?.v && prev.v[0] === 'h') || (next?.v && next.v[0] === 's')) tk.v = tk.v === 'm' ? 'n' : 'nn'
    })
    // slashes in dates are literals
    for (let k = 0; k < toks.length; k++) if (toks[k].t === 'slash') toks[k] = { t: 'lit', v: '/' }
    for (let k = 0; k < toks.length; k++) if (toks[k].t === 'dot') toks[k] = { t: 'lit', v: '.' }
    for (let k = 0; k < toks.length; k++) if (toks[k].t === 'comma') toks[k] = { t: 'lit', v: ',' }
  }
  return { toks, color, isDate, cond }
}

function getSections(fmt: string): Section[] {
  let s = cache.get(fmt)
  if (!s) {
    s = splitSections(fmt).map(parseSection)
    if (cache.size > 500) cache.clear()
    cache.set(fmt, s)
  }
  return s
}

export function isDateFormat(fmt: string | undefined): boolean {
  if (!fmt) return false
  return getSections(fmt)[0]?.isDate ?? false
}

export function isTextFormat(fmt: string | undefined): boolean {
  return fmt === '@'
}

// ---------- General ----------
export function formatGeneral(v: number): string {
  if (!isFinite(v)) return '#NUM!'
  if (v === 0) return '0'
  const a = Math.abs(v)
  if (a >= 1e11 || a < 1e-9) {
    return v.toExponential(5).replace(/\.?0+e/, 'E').replace('E+', 'E+').replace(/e/, 'E').replace(/E([+-])(\d)$/, 'E$10$2')
  }
  if (Number.isInteger(v)) return String(v)
  let s = String(Number(v.toPrecision(10)))
  if (s.includes('e')) s = v.toFixed(10).replace(/0+$/, '').replace(/\.$/, '')
  return s
}

// ---------- Dates ----------
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

export interface DateParts {
  y: number
  m: number // 1-12
  d: number
  wd: number // 0 = Sunday
  hh: number
  mi: number
  ss: number
  ms: number
}

/** Excel 1900 date system serial -> parts (includes the 1900 leap-year quirk). */
export function serialToParts(serial: number): DateParts {
  let days = Math.floor(serial)
  let frac = serial - days
  let msTotal = Math.round(frac * 86400000)
  if (msTotal >= 86400000) {
    days += 1
    msTotal -= 86400000
  }
  let y: number, m: number, d: number, wd: number
  if (days === 60) {
    y = 1900
    m = 2
    d = 29
    wd = 3
  } else {
    const base = Date.UTC(1899, 11, days < 60 ? 31 : 30)
    const dt = new Date(base + days * 86400000)
    y = dt.getUTCFullYear()
    m = dt.getUTCMonth() + 1
    d = dt.getUTCDate()
    wd = dt.getUTCDay()
  }
  if (days === 0) {
    y = 1900
    m = 1
    d = 0
  }
  const hh = Math.floor(msTotal / 3600000)
  const mi = Math.floor((msTotal % 3600000) / 60000)
  const ss = Math.floor((msTotal % 60000) / 1000)
  frac = msTotal % 1000
  return { y, m, d, wd, hh, mi, ss, ms: frac }
}

export function partsToSerial(y: number, m: number, d: number, hh = 0, mi = 0, ss = 0): number {
  const ms = Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)
  let serial = ms / 86400000
  if (serial < 61) serial -= 1 // before 1900-03-01 (Excel's fake 1900-02-29)
  return serial + (hh * 3600 + mi * 60 + ss) / 86400
}

const pad = (n: number, w = 2) => String(n).padStart(w, '0')

function formatDate(serial: number, sec: Section): string {
  if (serial < 0) return '#'.repeat(10)
  const hasSub = sec.toks.find((t) => t.t === 'subsec') as { v: number } | undefined
  const p = serialToParts(hasSub ? serial : Math.round(serial * 86400) / 86400)
  const hasAmPm = sec.toks.some((t) => t.t === 'ampm')
  let out = ''
  for (const t of sec.toks) {
    switch (t.t) {
      case 'lit':
        out += t.v
        break
      case 'date': {
        const v = t.v
        if (v[0] === 'y') out += v.length <= 2 ? pad(p.y % 100) : String(p.y)
        else if (v[0] === 'm') {
          if (v.length === 1) out += p.m
          else if (v.length === 2) out += pad(p.m)
          else if (v.length === 3) out += MONTHS[p.m - 1].slice(0, 3)
          else if (v.length === 5) out += MONTHS[p.m - 1][0]
          else out += MONTHS[p.m - 1]
        } else if (v[0] === 'n') out += v.length === 1 ? p.mi : pad(p.mi)
        else if (v[0] === 'd') {
          if (v.length === 1) out += p.d
          else if (v.length === 2) out += pad(p.d)
          else if (v.length === 3) out += DAYS[p.wd].slice(0, 3)
          else out += DAYS[p.wd]
        } else if (v[0] === 'h') {
          let h = p.hh
          if (hasAmPm) h = h % 12 === 0 ? 12 : h % 12
          out += v.length === 1 ? h : pad(h)
        } else if (v[0] === 's') out += v.length === 1 ? p.ss : pad(p.ss)
        break
      }
      case 'subsec':
        out += '.' + pad(Math.floor(p.ms / Math.pow(10, 3 - t.v)), t.v)
        break
      case 'ampm': {
        const pm = p.hh >= 12
        out += t.v.length > 3 ? (pm ? 'PM' : 'AM') : pm ? 'P' : 'A'
        break
      }
      case 'elapsed': {
        const totalSec = Math.round(serial * 86400)
        if (t.v[0] === 'h') out += pad(Math.floor(totalSec / 3600), t.v.length)
        else if (t.v[0] === 'm') out += pad(Math.floor(totalSec / 60), t.v.length)
        else out += pad(totalSec, t.v.length)
        break
      }
      case 'general':
        out += formatGeneral(serial)
        break
      default:
        break
    }
  }
  return out
}

// ---------- Numbers ----------
function fillInteger(digits: string, pattern: Tok[], thousands: boolean): string {
  // pattern: integer-part tokens (dig + lit). Fill right-to-left.
  const placeholders = pattern.filter((t) => t.t === 'dig') as { t: 'dig'; v: string }[]
  let ds = digits === '0' ? '' : digits
  if (thousands && ds.length > 3) {
    // thousands separators are inserted as part of the digit stream
    ds = ds.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  }
  const out: string[] = []
  let di = ds.length - 1
  let pi = placeholders.length
  for (let k = pattern.length - 1; k >= 0; k--) {
    const t = pattern[k]
    if (t.t === 'lit') {
      out.push(t.v)
      continue
    }
    if (t.t !== 'dig') continue
    pi--
    if (pi === 0) {
      // leftmost placeholder takes all remaining digits
      const rest = di >= 0 ? ds.slice(0, di + 1) : ''
      di = -1
      if (rest) out.push(rest)
      else out.push(t.v === '0' ? '0' : t.v === '?' ? ' ' : '')
    } else if (di >= 0) {
      let chunk = ds[di--]
      if (chunk === ',' && di >= 0) chunk = ds[di--] + ','
      out.push(chunk)
    } else out.push(t.v === '0' ? '0' : t.v === '?' ? ' ' : '')
  }
  return out.reverse().join('')
}

function fillDecimal(digits: string, pattern: Tok[]): string {
  let out = ''
  let di = 0
  // trailing optional placeholders drop trailing zeros
  const placeholders = pattern.filter((t) => t.t === 'dig') as { v: string }[]
  let lastSig = digits.length - 1
  while (lastSig >= 0 && digits[lastSig] === '0') lastSig--
  let pIdx = 0
  for (const t of pattern) {
    if (t.t === 'lit') {
      out += t.v
      continue
    }
    if (t.t !== 'dig') continue
    const ph = placeholders[pIdx++].v
    const d = digits[di++] ?? '0'
    if (ph === '0' || di - 1 <= lastSig) out += d
    else if (ph === '?') out += ' '
  }
  return out
}

function gcdFrac(x: number, maxDen: number): [number, number] {
  // best rational approximation with denominator <= maxDen (Stern–Brocot / continued fractions)
  let [a, b, c, d] = [0, 1, 1, 0]
  let bestN = 0
  let bestD = 1
  let bestErr = Math.abs(x)
  for (let i = 0; i < 64; i++) {
    const mn = a + c
    const md = b + d
    if (md > maxDen) break
    const err = Math.abs(x - mn / md)
    if (err < bestErr) {
      bestErr = err
      bestN = mn
      bestD = md
    }
    if (mn / md < x) {
      a = mn
      b = md
    } else if (mn / md > x) {
      c = mn
      d = md
    } else break
  }
  if (Math.abs(x - 1) < bestErr) return [1, 1]
  return [bestN, bestD]
}

function formatNumberSection(v: number, sec: Section, signed: boolean): string {
  const toks = sec.toks
  if (toks.some((t) => t.t === 'general')) {
    return toks.map((t) => (t.t === 'general' ? formatGeneral(v) : t.t === 'lit' ? t.v : '')).join('')
  }
  let value = v
  const neg = value < 0
  value = Math.abs(value)
  const pctCount = toks.filter((t) => t.t === 'pct').length
  value *= Math.pow(100, pctCount)

  // thousands scaling: commas immediately after the last digit placeholder (before dot or end)
  let lastDig = -1
  toks.forEach((t, k) => {
    if (t.t === 'dig') lastDig = k
  })
  const dotIdx = toks.findIndex((t) => t.t === 'dot')
  let scaleEnd = dotIdx >= 0 ? dotIdx : toks.length
  let k = (dotIdx >= 0 ? dotIdx : lastDig + 1) - 1
  if (dotIdx < 0) scaleEnd = lastDig + 1
  let scale = 0
  for (let j = scaleEnd; j < toks.length && toks[j].t === 'comma'; j++) scale++
  if (dotIdx >= 0) {
    // commas right before the dot
    while (k >= 0 && toks[k].t === 'comma') {
      scale++
      k--
    }
  }
  value /= Math.pow(1000, scale)
  const firstDig = toks.findIndex((t) => t.t === 'dig')
  const thousands = toks.some((t, idx) => t.t === 'comma' && idx > firstDig && idx < (dotIdx >= 0 ? dotIdx : lastDig) && firstDig >= 0)

  const prefix: Tok[] = []
  const suffix: Tok[] = []
  const expIdx = toks.findIndex((t) => t.t === 'exp')
  const slashIdx = toks.findIndex((t) => t.t === 'slash')

  // Literal text before the first / after the last placeholder
  const core: Tok[] = []
  toks.forEach((t, idx) => {
    if (t.t === 'pct') {
      ;(idx < firstDig ? prefix : idx > lastDig ? suffix : core).push({ t: 'lit', v: '%' })
      return
    }
    if (t.t === 'comma' && (idx <= firstDig || idx > lastDig)) return
    if (firstDig < 0 || idx < firstDig) {
      if (t.t === 'lit') prefix.push(t)
    } else if (idx > lastDig && t.t !== 'exp') {
      if (t.t === 'lit') suffix.push(t)
    } else core.push(t)
  })
  const lit = (arr: Tok[]) => arr.map((t) => (t.t === 'lit' ? t.v : '')).join('')
  let body = ''

  if (firstDig < 0) {
    body = ''
  } else if (slashIdx >= 0) {
    // fraction: [int part] num/den
    const before = core.slice(0, core.findIndex((t) => t.t === 'slash'))
    const after = core.slice(core.findIndex((t) => t.t === 'slash') + 1)
    const denDigits = after.filter((t) => t.t === 'dig').length
    const fixedDen = parseInt(lit(after).replace(/\D/g, '') || '0', 10)
    // split "before" into integer placeholders and numerator placeholders (separated by a literal space)
    const spaceIdx = before.findIndex((t) => t.t === 'lit' && t.v === ' ')
    const hasInt = spaceIdx > 0
    const intPart = hasInt ? Math.floor(value) : 0
    const fracVal = value - intPart
    const [n, d] = fixedDen > 0 ? [Math.round(fracVal * fixedDen), fixedDen] : gcdFrac(fracVal, Math.pow(10, denDigits) - 1)
    let ip = intPart
    let nn = n
    if (d > 0 && nn === d && hasInt) {
      ip += 1
      nn = 0
    }
    if (hasInt) {
      body = (ip ? String(ip) : '') + (nn ? ` ${nn}/${d}` : ip ? '' : '0')
    } else body = `${nn}/${d}`
  } else if (expIdx >= 0) {
    const mant = core.slice(0, core.findIndex((t) => t.t === 'exp'))
    const expDigits = core.slice(core.findIndex((t) => t.t === 'exp') + 1).filter((t) => t.t === 'dig').length
    const md = mant.findIndex((t) => t.t === 'dot')
    const intPh = (md >= 0 ? mant.slice(0, md) : mant).filter((t) => t.t === 'dig').length
    const decPh = md >= 0 ? mant.slice(md + 1).filter((t) => t.t === 'dig').length : 0
    let e = value === 0 ? 0 : Math.floor(Math.log10(value))
    // engineering-style grouping when several integer placeholders (e.g. ##0.0E+0)
    if (intPh > 1) e = Math.floor(e / intPh) * intPh
    let m = value / Math.pow(10, e)
    m = parseFloat(m.toFixed(decPh))
    if (m >= Math.pow(10, Math.max(1, intPh))) {
      m /= 10
      e += 1
    }
    const [mi, mdc = ''] = m.toFixed(decPh).split('.')
    body = fillInteger(mi, (md >= 0 ? mant.slice(0, md) : mant).filter((t) => t.t !== 'comma'), false)
    if (decPh) body += '.' + fillDecimal(mdc, mant.slice(md + 1))
    const expTok = core.find((t) => t.t === 'exp') as { v: string }
    body += 'E' + (e < 0 ? '-' : expTok.v === '+' ? '+' : '') + String(Math.abs(e)).padStart(expDigits, '0')
  } else {
    const cd = core.findIndex((t) => t.t === 'dot')
    const intToks = (cd >= 0 ? core.slice(0, cd) : core).filter((t) => t.t !== 'comma')
    const decToks = cd >= 0 ? core.slice(cd + 1).filter((t) => t.t !== 'comma') : []
    const decPh = decToks.filter((t) => t.t === 'dig').length
    const fixed = value.toFixed(decPh)
    const [ip, dp = ''] = fixed.split('.')
    body = fillInteger(ip, intToks, thousands)
    if (cd >= 0) {
      const dec = fillDecimal(dp, decToks)
      body += dec.length || decToks.some((t) => t.t === 'dig' && t.v === '0') ? '.' + dec : '.'
    }
  }
  const isZeroDisplay = !/[1-9]/.test(body)
  const minus = neg && signed && !isZeroDisplay ? '-' : ''
  return minus + lit(prefix) + body + lit(suffix)
}

function condMatch(c: Section['cond'], v: number): boolean {
  if (!c) return false
  switch (c.op) {
    case '<': return v < c.val
    case '<=': return v <= c.val
    case '>': return v > c.val
    case '>=': return v >= c.val
    case '=': return v === c.val
    case '<>': return v !== c.val
  }
  return false
}

export function formatNumber(v: number, fmt: string | undefined): Formatted {
  if (!fmt || fmt.toLowerCase() === 'general') return { text: formatGeneral(v) }
  if (fmt === '@') return { text: formatGeneral(v) }
  const secs = getSections(fmt).filter((s) => !(s.toks.length === 1 && s.toks[0].t === 'text') || true)
  let sec: Section
  let signed = true
  if (secs.some((s) => s.cond)) {
    sec = secs.find((s) => condMatch(s.cond, v)) ?? secs.find((s) => !s.cond) ?? secs[0]
    signed = !sec.cond || sec.cond.val >= 0
  } else if (secs.length === 1 || (secs.length >= 2 && v > 0) || (secs.length === 2 && v === 0)) sec = secs[0]
  else if (v < 0) {
    sec = secs[1]
    signed = false
  } else sec = secs.length >= 3 ? secs[2] : secs[0]
  if (sec.toks.length === 0) return { text: '' }
  if (sec.isDate) return { text: formatDate(v, sec), color: sec.color }
  return { text: formatNumberSection(v, sec, signed), color: sec.color }
}

export function formatText(s: string, fmt: string | undefined): Formatted {
  if (!fmt) return { text: s }
  const secs = getSections(fmt)
  const textSec = secs.length >= 4 ? secs[3] : secs.find((x) => x.toks.some((t) => t.t === 'text'))
  if (!textSec) return { text: s }
  return {
    text: textSec.toks.map((t) => (t.t === 'text' ? s : t.t === 'lit' ? t.v : '')).join(''),
    color: textSec.color,
  }
}

/** Number-format presets offered in the UI. */
export const FORMAT_PRESETS: { id: string; label: string; fmt: string | undefined }[] = [
  { id: 'general', label: 'General', fmt: undefined },
  { id: 'number', label: 'Number', fmt: '#,##0.00' },
  { id: 'currency', label: 'Currency', fmt: '$#,##0.00' },
  { id: 'accounting', label: 'Accounting', fmt: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' },
  { id: 'shortDate', label: 'Short Date', fmt: 'm/d/yyyy' },
  { id: 'isoDate', label: 'ISO Date', fmt: 'yyyy-mm-dd' },
  { id: 'longDate', label: 'Long Date', fmt: 'dddd, mmmm d, yyyy' },
  { id: 'time', label: 'Time', fmt: 'h:mm:ss AM/PM' },
  { id: 'percent', label: 'Percentage', fmt: '0.00%' },
  { id: 'fraction', label: 'Fraction', fmt: '# ?/?' },
  { id: 'scientific', label: 'Scientific', fmt: '0.00E+00' },
  { id: 'text', label: 'Text', fmt: '@' },
]

/** Increase/decrease the number of decimals in a format code (first section drives it). */
export function adjustDecimals(fmt: string | undefined, delta: 1 | -1, sampleValue?: number): string {
  if (!fmt || fmt.toLowerCase() === 'general') {
    let decimals = 0
    if (sampleValue !== undefined && !Number.isInteger(sampleValue)) {
      const s = formatGeneral(sampleValue)
      decimals = s.includes('.') ? s.split('.')[1].replace(/E.*/, '').length : 0
    }
    decimals = Math.max(0, decimals + delta)
    return decimals ? '0.' + '0'.repeat(decimals) : '0'
  }
  return splitSections(fmt)
    .map((sec) => {
      if (delta > 0) {
        if (/0\.0*|#\.#*/.test(sec)) return sec.replace(/([0#])(\.)?([0#]*)(?![0#.])/, (m, a, dot, ds) => (dot ? `${a}.${ds}0` : m))
        return sec.replace(/0(?![0.,])(?![^"]*"(?:[^"]*"[^"]*")*[^"]*$)/, '0.0')
      }
      return sec.replace(/\.([0#]+)/, (_m, ds: string) => (ds.length > 1 ? '.' + ds.slice(0, -1) : ''))
    })
    .join(';')
}
