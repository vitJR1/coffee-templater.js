import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate'
import { escapeAttr } from './xml/escape'
import { childElements, parseXml } from './xml/dom'

export type BinaryInput = Uint8Array | ArrayBuffer | ArrayBufferView

export function toUint8Array(input: BinaryInput | string): Uint8Array {
  if (typeof input === 'string') return base64ToBytes(input)
  if (input instanceof Uint8Array) return input
  if (input instanceof ArrayBuffer) return new Uint8Array(input)
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength)
  throw new TypeError('Expected Uint8Array, ArrayBuffer, Buffer or base64 string')
}

function base64ToBytes(b64: string): Uint8Array {
  const clean = b64.replace(/^data:[^;]+;base64,/, '').replace(/\s+/g, '')
  if (typeof atob === 'function') {
    const bin = atob(clean)
    const out = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
    return out
  }
  // Node < 16 fallback
  const B = (globalThis as { Buffer?: { from(s: string, e: string): Uint8Array } }).Buffer
  if (!B) throw new Error('No base64 decoder available')
  return new Uint8Array(B.from(clean, 'base64'))
}

export interface Relationship {
  id: string
  type: string
  target: string
  targetMode?: string
}

export const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  footnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  worksheet: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
  sharedStrings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  calcChain: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
} as const

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tiff: 'image/tiff',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
}

/** An Open Packaging Convention container (the ZIP inside every .docx/.xlsx/.pptx). */
export class OpcPackage {
  readonly files: Map<string, Uint8Array>
  private textCache = new Map<string, string>()

  constructor(files: Map<string, Uint8Array>) {
    this.files = files
  }

  static load(input: BinaryInput): OpcPackage {
    let entries: Record<string, Uint8Array>
    try {
      entries = unzipSync(toUint8Array(input))
    } catch (e) {
      throw new Error(`Input is not a valid OOXML (zip) file: ${(e as Error).message}`)
    }
    const files = new Map<string, Uint8Array>()
    for (const [name, data] of Object.entries(entries)) files.set(name.replace(/^\//, ''), data)
    if (!files.has('[Content_Types].xml')) throw new Error('Input is not an OOXML package: [Content_Types].xml is missing')
    return new OpcPackage(files)
  }

  clone(): OpcPackage {
    const pkg = new OpcPackage(new Map(this.files))
    for (const [k, v] of this.textCache) pkg.textCache.set(k, v)
    return pkg
  }

  has(path: string): boolean {
    return this.files.has(path) || this.textCache.has(path)
  }

  getText(path: string): string | undefined {
    const cached = this.textCache.get(path)
    if (cached !== undefined) return cached
    const data = this.files.get(path)
    if (!data) return undefined
    const str = strFromU8(data)
    this.textCache.set(path, str)
    return str
  }

  setText(path: string, value: string): void {
    this.textCache.set(path, value)
    this.files.set(path, strToU8(value))
  }

  setBinary(path: string, data: Uint8Array): void {
    this.textCache.delete(path)
    this.files.set(path, data)
  }

  remove(path: string): void {
    this.files.delete(path)
    this.textCache.delete(path)
  }

  toBytes(): Uint8Array {
    const zippable: Zippable = {}
    // [Content_Types].xml first, as Office itself writes it.
    const names = [...this.files.keys()].sort((a, b) =>
      a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : 0,
    )
    for (const name of names) zippable[name] = [this.files.get(name)!, { level: 6 }]
    return zipSync(zippable)
  }

  // ---------- relationships ----------

  static relsPathFor(partPath: string): string {
    const slash = partPath.lastIndexOf('/')
    const dir = slash === -1 ? '' : partPath.slice(0, slash + 1)
    const file = partPath.slice(slash + 1)
    return `${dir}_rels/${file}.rels`
  }

  /** Relationships of a part (use '' for package-level rels). */
  getRelationships(partPath: string): Relationship[] {
    const relsPath = partPath === '' ? '_rels/.rels' : OpcPackage.relsPathFor(partPath)
    const xml = this.getText(relsPath)
    if (!xml) return []
    return childElements(parseXml(xml).root, 'Relationship').map((r) => ({
      id: r.attrs.Id,
      type: r.attrs.Type,
      target: r.attrs.Target,
      targetMode: r.attrs.TargetMode,
    }))
  }

  /** Resolves a relationship target to an absolute part path inside the package. */
  static resolveTarget(partPath: string, target: string): string {
    if (target.startsWith('/')) return target.slice(1)
    const base = partPath.includes('/') ? partPath.slice(0, partPath.lastIndexOf('/')).split('/') : []
    for (const seg of target.split('/')) {
      if (seg === '..') base.pop()
      else if (seg !== '.' && seg !== '') base.push(seg)
    }
    return base.join('/')
  }

  /** Returns [relationship, absolute target path] pairs of the given type. */
  getRelatedParts(partPath: string, type: string): Array<{ rel: Relationship; path: string }> {
    return this.getRelationships(partPath)
      .filter((r) => r.type === type && r.targetMode !== 'External')
      .map((rel) => ({ rel, path: OpcPackage.resolveTarget(partPath, rel.target) }))
  }

  mainDocumentPath(): string | undefined {
    return this.getRelatedParts('', REL.officeDocument)[0]?.path
  }

  addRelationship(partPath: string, type: string, target: string, external = false): string {
    const relsPath = OpcPackage.relsPathFor(partPath)
    let xml = this.getText(relsPath)
    if (!xml) {
      xml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    }
    const used = new Set([...xml.matchAll(/\bId="([^"]+)"/g)].map((m) => m[1]))
    let n = used.size + 1
    while (used.has(`rId${n}`)) n++
    const id = `rId${n}`
    const mode = external ? ' TargetMode="External"' : ''
    const rel = `<Relationship Id="${id}" Type="${escapeAttr(type)}" Target="${escapeAttr(target)}"${mode}/>`
    xml = xml.includes('</Relationships>')
      ? xml.replace('</Relationships>', `${rel}</Relationships>`)
      : xml.replace(/<Relationships([^>]*)\/>/, `<Relationships$1>${rel}</Relationships>`)
    this.setText(relsPath, xml)
    return id
  }

  removeRelationship(partPath: string, predicate: (r: Relationship) => boolean): Relationship[] {
    const relsPath = OpcPackage.relsPathFor(partPath)
    const xml = this.getText(relsPath)
    if (!xml) return []
    const removed: Relationship[] = []
    const next = xml.replace(/<Relationship\b[^>]*\/>/g, (raw) => {
      const id = /\bId="([^"]*)"/.exec(raw)?.[1] ?? ''
      const type = /\bType="([^"]*)"/.exec(raw)?.[1] ?? ''
      const target = /\bTarget="([^"]*)"/.exec(raw)?.[1] ?? ''
      const rel = { id, type, target }
      if (predicate(rel)) {
        removed.push(rel)
        return ''
      }
      return raw
    })
    this.setText(relsPath, next)
    return removed
  }

  // ---------- content types ----------

  ensureDefaultContentType(extension: string, contentType: string): void {
    const path = '[Content_Types].xml'
    const xml = this.getText(path)!
    const ext = extension.toLowerCase()
    if (new RegExp(`<Default\\b[^>]*Extension="${ext}"`, 'i').test(xml)) return
    this.setText(path, xml.replace('</Types>', `<Default Extension="${ext}" ContentType="${escapeAttr(contentType)}"/></Types>`))
  }

  ensureOverride(partPath: string, contentType: string): void {
    const path = '[Content_Types].xml'
    const xml = this.getText(path)!
    if (xml.includes(`PartName="/${partPath}"`)) return
    this.setText(path, xml.replace('</Types>', `<Override PartName="/${escapeAttr(partPath)}" ContentType="${escapeAttr(contentType)}"/></Types>`))
  }

  removeOverride(partPath: string): void {
    const path = '[Content_Types].xml'
    const xml = this.getText(path)!
    this.setText(path, xml.replace(new RegExp(`<Override\\b[^>]*PartName="/${partPath.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), ''))
  }

  /** Stores an image in `<dir>/media/` and links it from `partPath`. Returns the relationship id. */
  addImage(partPath: string, mediaDir: string, data: Uint8Array, extension: string): string {
    const ext = extension.toLowerCase() === 'jpg' ? 'jpeg' : extension.toLowerCase()
    let n = 1
    while (this.has(`${mediaDir}/image_ct${n}.${ext}`)) n++
    const mediaPath = `${mediaDir}/image_ct${n}.${ext}`
    this.setBinary(mediaPath, data)
    this.ensureDefaultContentType(ext, IMAGE_MIME[ext] ?? `image/${ext}`)
    const partDir = partPath.slice(0, partPath.lastIndexOf('/'))
    const target = relativePath(partDir, mediaPath)
    return this.addRelationship(partPath, REL.image, target)
  }
}

function relativePath(fromDir: string, to: string): string {
  const from = fromDir ? fromDir.split('/') : []
  const parts = to.split('/')
  let i = 0
  while (i < from.length && i < parts.length - 1 && from[i] === parts[i]) i++
  return [...Array(from.length - i).fill('..'), ...parts.slice(i)].join('/')
}
