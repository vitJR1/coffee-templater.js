import { decodeXml, escapeAttr, escapeXml } from './escape'
import { parseAttributes, tokenizeXml } from './tokenizer'

/** Minimal mutable XML tree used by the document model (parse/build). */
export interface XmlElement {
  type: 'element'
  name: string
  attrs: Record<string, string>
  children: XmlNode[]
}
export interface XmlText {
  type: 'text'
  text: string
}
export interface XmlOther {
  type: 'other'
  raw: string
}
export type XmlNode = XmlElement | XmlText | XmlOther

export interface XmlDocument {
  prolog: XmlNode[]
  root: XmlElement
}

export function parseXml(xml: string): XmlDocument {
  const tokens = tokenizeXml(xml)
  const prolog: XmlNode[] = []
  const stack: XmlElement[] = []
  let root: XmlElement | undefined

  for (const tok of tokens) {
    const parent = stack[stack.length - 1]
    if (tok.type === 'open') {
      const el: XmlElement = { type: 'element', name: tok.name, attrs: decodeAttrs(parseAttributes(tok.raw)), children: [] }
      if (parent) parent.children.push(el)
      else if (!root) root = el
      if (!tok.selfClosing) stack.push(el)
    } else if (tok.type === 'close') {
      const top = stack.pop()
      if (!top || top.name !== tok.name) throw new Error(`Malformed XML: unexpected </${tok.name}>`)
    } else if (tok.type === 'text') {
      if (parent) parent.children.push({ type: 'text', text: decodeXml(tok.raw) })
    } else if (parent) {
      if (tok.raw.startsWith('<![CDATA[')) parent.children.push({ type: 'text', text: tok.raw.slice(9, -3) })
      else parent.children.push({ type: 'other', raw: tok.raw })
    } else if (!root) {
      prolog.push({ type: 'other', raw: tok.raw })
    }
  }
  if (!root) throw new Error('Malformed XML: no root element')
  return { prolog, root }
}

function decodeAttrs(attrs: Record<string, string>): Record<string, string> {
  for (const k of Object.keys(attrs)) attrs[k] = decodeXml(attrs[k])
  return attrs
}

export function serializeXml(doc: XmlDocument): string {
  let out = ''
  for (const n of doc.prolog) out += serializeNode(n)
  return out + serializeNode(doc.root)
}

export function serializeNode(node: XmlNode): string {
  if (node.type === 'text') return escapeXml(node.text)
  if (node.type === 'other') return node.raw
  let out = `<${node.name}`
  for (const [k, v] of Object.entries(node.attrs)) out += ` ${k}="${escapeAttr(v)}"`
  if (node.children.length === 0) return out + '/>'
  out += '>'
  for (const c of node.children) out += serializeNode(c)
  return out + `</${node.name}>`
}

export function serializeChildren(el: XmlElement): string {
  return el.children.map(serializeNode).join('')
}

/** Parses an XML fragment (possibly several sibling elements). */
export function parseFragment(xml: string): XmlNode[] {
  if (!xml) return []
  return parseXml(`<__root>${xml}</__root>`).root.children
}

export function el(name: string, attrs: Record<string, string | undefined> = {}, children: XmlNode[] = []): XmlElement {
  const clean: Record<string, string> = {}
  for (const [k, v] of Object.entries(attrs)) if (v !== undefined) clean[k] = v
  return { type: 'element', name, attrs: clean, children }
}

export function text(value: string): XmlText {
  return { type: 'text', text: value }
}

export function childElements(node: XmlElement, name?: string): XmlElement[] {
  return node.children.filter((c): c is XmlElement => c.type === 'element' && (name === undefined || c.name === name))
}

export function firstChild(node: XmlElement, name: string): XmlElement | undefined {
  return node.children.find((c): c is XmlElement => c.type === 'element' && c.name === name)
}

/** Depth-first search for the first descendant with the given name. */
export function findDescendant(node: XmlElement, name: string): XmlElement | undefined {
  for (const c of node.children) {
    if (c.type !== 'element') continue
    if (c.name === name) return c
    const found = findDescendant(c, name)
    if (found) return found
  }
  return undefined
}

export function findDescendants(node: XmlElement, name: string, out: XmlElement[] = []): XmlElement[] {
  for (const c of node.children) {
    if (c.type !== 'element') continue
    if (c.name === name) out.push(c)
    findDescendants(c, name, out)
  }
  return out
}

/** Concatenated text of all descendant text nodes. */
export function textContent(node: XmlNode): string {
  if (node.type === 'text') return node.text
  if (node.type === 'other') return ''
  return node.children.map(textContent).join('')
}

/** Sorts element children according to a schema order list; unknown names keep relative order at the end. */
export function sortChildren(children: XmlNode[], order: readonly string[]): XmlNode[] {
  const rank = (n: XmlNode): number => {
    if (n.type !== 'element') return order.length
    const idx = order.indexOf(n.name)
    return idx === -1 ? order.length : idx
  }
  return children
    .map((n, i) => ({ n, i, r: rank(n) }))
    .sort((a, b) => a.r - b.r || a.i - b.i)
    .map((x) => x.n)
}
