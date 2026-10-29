/** Minimal, dependency-free XML parser/writer for OOXML parts (works in browsers and Node). */

export interface XmlEl {
  name: string // local name (namespace prefix removed)
  qname: string // qualified name as written
  attrs: Record<string, string> // keys as written (e.g. "r:id") plus local-name aliases
  children: XmlEl[]
  text: string // concatenated direct text content
}

const ENT: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }

export function decodeEntities(s: string): string {
  if (!s.includes('&')) return s
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (m, e: string) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)
      return isNaN(code) ? m : String.fromCodePoint(code)
    }
    return ENT[e] ?? m
  })
}

export function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, (c) => `_x${c.charCodeAt(0).toString(16).padStart(4, '0').toUpperCase()}_`)
}

/** OOXML escapes control characters as _xHHHH_. */
export function decodeOoxmlEscapes(s: string): string {
  return s.includes('_x') ? s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16))) : s
}

const local = (q: string) => {
  const i = q.indexOf(':')
  return i < 0 ? q : q.slice(i + 1)
}

export function parseXml(xml: string): XmlEl {
  const root: XmlEl = { name: '#root', qname: '#root', attrs: {}, children: [], text: '' }
  const stack: XmlEl[] = [root]
  let i = 0
  const n = xml.length
  const attrRe = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g
  while (i < n) {
    const lt = xml.indexOf('<', i)
    if (lt < 0) break
    if (lt > i) {
      const top = stack[stack.length - 1]
      top.text += decodeEntities(xml.slice(i, lt))
    }
    if (xml.startsWith('<!--', lt)) {
      const end = xml.indexOf('-->', lt + 4)
      i = end < 0 ? n : end + 3
      continue
    }
    if (xml.startsWith('<![CDATA[', lt)) {
      const end = xml.indexOf(']]>', lt + 9)
      stack[stack.length - 1].text += xml.slice(lt + 9, end < 0 ? n : end)
      i = end < 0 ? n : end + 3
      continue
    }
    if (xml[lt + 1] === '?' || xml[lt + 1] === '!') {
      const end = xml.indexOf('>', lt)
      i = end < 0 ? n : end + 1
      continue
    }
    const gt = xml.indexOf('>', lt)
    if (gt < 0) break
    let tag = xml.slice(lt + 1, gt)
    if (tag[0] === '/') {
      if (stack.length > 1) stack.pop()
      i = gt + 1
      continue
    }
    const selfClose = tag.endsWith('/')
    if (selfClose) tag = tag.slice(0, -1)
    const sp = tag.search(/\s/)
    const qname = sp < 0 ? tag : tag.slice(0, sp)
    const el: XmlEl = { name: local(qname), qname, attrs: {}, children: [], text: '' }
    if (sp >= 0) {
      const rest = tag.slice(sp)
      attrRe.lastIndex = 0
      let m: RegExpExecArray | null
      while ((m = attrRe.exec(rest))) {
        const v = decodeEntities(m[3] ?? m[4] ?? '')
        el.attrs[m[1]] = v
        const l = local(m[1])
        if (!(l in el.attrs)) el.attrs[l] = v
      }
    }
    stack[stack.length - 1].children.push(el)
    if (!selfClose) stack.push(el)
    i = gt + 1
  }
  return root.children[0] ?? root
}

export function child(el: XmlEl | undefined, name: string): XmlEl | undefined {
  return el?.children.find((c) => c.name === name)
}

export function children(el: XmlEl | undefined, name: string): XmlEl[] {
  return el ? el.children.filter((c) => c.name === name) : []
}

/** All text inside an element (recursively), used for rich-text runs. */
export function deepText(el: XmlEl): string {
  if (el.name === 't') return el.text
  if (el.name === 'rPh') return '' // phonetic hints are not cell text
  return el.children.map(deepText).join('')
}
