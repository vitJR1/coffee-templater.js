/**
 * A tiny, lossless XML tokenizer. Joining `raw` of every token reproduces the input exactly.
 * It is intentionally forgiving: OOXML produced by Office is well-formed, and we never need validation.
 */

export type XmlToken =
  | { type: 'open'; name: string; raw: string; selfClosing: boolean }
  | { type: 'close'; name: string; raw: string }
  | { type: 'text'; raw: string }
  | { type: 'other'; raw: string } // <?xml?>, comments, CDATA, doctype

export function tokenizeXml(xml: string): XmlToken[] {
  const tokens: XmlToken[] = []
  const len = xml.length
  let i = 0
  while (i < len) {
    const lt = xml.indexOf('<', i)
    if (lt === -1) {
      tokens.push({ type: 'text', raw: xml.slice(i) })
      break
    }
    if (lt > i) tokens.push({ type: 'text', raw: xml.slice(i, lt) })

    if (xml.startsWith('<!--', lt)) {
      const end = indexOrThrow(xml, '-->', lt)
      tokens.push({ type: 'other', raw: xml.slice(lt, end + 3) })
      i = end + 3
      continue
    }
    if (xml.startsWith('<![CDATA[', lt)) {
      const end = indexOrThrow(xml, ']]>', lt)
      tokens.push({ type: 'other', raw: xml.slice(lt, end + 3) })
      i = end + 3
      continue
    }
    if (xml.startsWith('<?', lt)) {
      const end = indexOrThrow(xml, '?>', lt)
      tokens.push({ type: 'other', raw: xml.slice(lt, end + 2) })
      i = end + 2
      continue
    }
    if (xml.startsWith('<!', lt)) {
      const end = indexOrThrow(xml, '>', lt)
      tokens.push({ type: 'other', raw: xml.slice(lt, end + 1) })
      i = end + 1
      continue
    }

    // Element tag: scan to the closing '>' while respecting quoted attribute values.
    let j = lt + 1
    let quote = ''
    while (j < len) {
      const ch = xml[j]
      if (quote) {
        if (ch === quote) quote = ''
      } else if (ch === '"' || ch === "'") {
        quote = ch
      } else if (ch === '>') {
        break
      }
      j++
    }
    if (j >= len) throw new Error(`Malformed XML: unterminated tag at offset ${lt}`)
    const raw = xml.slice(lt, j + 1)
    if (raw[1] === '/') {
      tokens.push({ type: 'close', name: raw.slice(2, -1).trim(), raw })
    } else {
      const selfClosing = raw[raw.length - 2] === '/'
      const m = /^<([^\s/>]+)/.exec(raw)
      tokens.push({ type: 'open', name: m ? m[1] : '', raw, selfClosing })
    }
    i = j + 1
  }
  return tokens
}

function indexOrThrow(xml: string, needle: string, from: number): number {
  const idx = xml.indexOf(needle, from)
  if (idx === -1) throw new Error(`Malformed XML: missing "${needle}" after offset ${from}`)
  return idx
}

/** Parses attributes from a raw open tag such as `<w:t xml:space="preserve">`. */
export function parseAttributes(raw: string): Record<string, string> {
  const attrs: Record<string, string> = {}
  const re = /([^\s=/<>]+)\s*=\s*("([^"]*)"|'([^']*)')/g
  const start = raw.search(/\s/)
  if (start === -1) return attrs
  const body = raw.slice(start)
  let m: RegExpExecArray | null
  while ((m = re.exec(body))) attrs[m[1]] = m[3] ?? m[4] ?? ''
  return attrs
}

/** Returns a copy of a raw open tag with attribute `name` set (escaped value must be provided). */
export function setRawAttribute(raw: string, name: string, escapedValue: string): string {
  const re = new RegExp(`(\\s${name.replace(/[.:]/g, '\\$&')}\\s*=\\s*)("[^"]*"|'[^']*')`)
  if (re.test(raw)) return raw.replace(re, `$1"${escapedValue}"`)
  const selfClosing = raw.endsWith('/>')
  const head = raw.slice(0, selfClosing ? -2 : -1).replace(/\s+$/, '')
  return `${head} ${name}="${escapedValue}"${selfClosing ? '/>' : '>'}`
}

/** Removes attribute `name` from a raw open tag. */
export function removeRawAttribute(raw: string, name: string): string {
  const re = new RegExp(`\\s${name.replace(/[.:]/g, '\\$&')}\\s*=\\s*("[^"]*"|'[^']*')`)
  return raw.replace(re, '')
}
