import { childElements, el, parseFragment, serializeNode, sortChildren, type XmlElement } from '../xml/dom'

/**
 * Properties elements (w:rPr, a:pPr, ...) are kept verbatim in the model as `props` XML, and the
 * convenient fields (bold, size, ...) are read from them. When building, a field is written back only
 * if it differs from the value in `props` — so everything we don't model survives a round-trip.
 */
export interface ChildField<T = unknown> {
  /** Child element names owned by this field. The first one is used for reading. */
  names: string[]
  read(child: XmlElement): T | undefined
  write(value: T): XmlElement[]
}

export interface AttrField<T = unknown> {
  attr: string
  read(value: string): T | undefined
  write(value: T): string
}

export type FieldMap = Record<string, ChildField<any> | AttrField<any>>

const isAttr = (f: ChildField<any> | AttrField<any>): f is AttrField<any> => 'attr' in f

export function readFields(propsEl: XmlElement | undefined, fields: FieldMap): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  if (!propsEl) return out
  for (const [key, f] of Object.entries(fields)) {
    let v: unknown
    if (isAttr(f)) {
      const raw = propsEl.attrs[f.attr]
      v = raw === undefined ? undefined : f.read(raw)
    } else {
      const child = childElements(propsEl).find((c) => c.name === f.names[0])
      v = child ? f.read(child) : undefined
    }
    if (v !== undefined) out[key] = v
  }
  return out
}

/** Builds the properties element from original `props` XML + model fields. Returns '' when empty. */
export function writeProps(
  name: string,
  propsXml: string | undefined,
  model: Record<string, unknown>,
  fields: FieldMap,
  order: readonly string[],
  keepEmpty = false,
): string {
  const parsed = propsXml ? parseFragment(propsXml).find((n): n is XmlElement => n.type === 'element') : undefined
  const propsEl: XmlElement = parsed ? { ...parsed, name, attrs: { ...parsed.attrs }, children: [...parsed.children] } : el(name)
  const original = readFields(propsEl, fields)

  for (const [key, f] of Object.entries(fields)) {
    const value = model[key]
    if (same(value, original[key])) continue
    if (isAttr(f)) {
      if (value === undefined || value === null) delete propsEl.attrs[f.attr]
      else propsEl.attrs[f.attr] = f.write(value)
    } else {
      propsEl.children = propsEl.children.filter((c) => !(c.type === 'element' && f.names.includes(c.name)))
      if (value !== undefined && value !== null) propsEl.children.push(...f.write(value))
    }
  }
  propsEl.children = sortChildren(propsEl.children, order)
  if (!keepEmpty && propsEl.children.length === 0 && Object.keys(propsEl.attrs).length === 0) return ''
  return serializeNode(propsEl)
}

function same(a: unknown, b: unknown): boolean {
  if (a === b) return true
  if (a && b && typeof a === 'object' && typeof b === 'object') return JSON.stringify(a) === JSON.stringify(b)
  return false
}

// Small field builders -----------------------------------------------------------------------

const OFF = new Set(['0', 'false', 'off', 'none'])

/** `<w:b/>` = true, `<w:b w:val="0"/>` = false */
export function onOff(name: string, valAttr = 'w:val'): ChildField<boolean> {
  return {
    names: [name],
    read: (c) => !OFF.has(c.attrs[valAttr] ?? 'true'),
    write: (v) => [v ? el(name) : el(name, { [valAttr]: '0' })],
  }
}

export function valueOf(name: string, valAttr = 'w:val'): ChildField<string> {
  return {
    names: [name],
    read: (c) => c.attrs[valAttr],
    write: (v) => [el(name, { [valAttr]: String(v) })],
  }
}

export function serializeEl(e: XmlElement | undefined): string | undefined {
  return e ? serializeNode(e) : undefined
}
