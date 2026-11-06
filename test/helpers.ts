import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

export const fixture = (name: string) => readFileSync(join(__dirname, 'fixtures', name))

export function readPart(file: Uint8Array, part: string): string {
  const entry = unzipSync(file)[part]
  if (!entry) throw new Error(`no part ${part}`)
  return strFromU8(entry)
}

export function partNames(file: Uint8Array): string[] {
  return Object.keys(unzipSync(file))
}

/** Plain text of every <w:t>/<a:t> in a part, paragraphs separated by \n. */
export function textOf(xml: string, para = 'w:p', t = 'w:t'): string {
  const pieces = xml.split(`</${para}>`)
  pieces.pop() // content after the last paragraph
  return pieces
    .map((p) => [...p.matchAll(new RegExp(`<${t}(?:\\s[^>]*)?>([^<]*)</${t}>`, 'g'))].map((m) => m[1]).join(''))
    .join('\n')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

/** Minimal docx with the given body XML. */
export function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    'word/document.xml': strToU8(`<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`),
  })
}

/** `p('Hello ', '{{na', 'me}}')` → paragraph with one run per argument. */
export const p = (...runs: string[]) => `<w:p>${runs.map((r) => `<w:r><w:t xml:space="preserve">${r}</w:t></w:r>`).join('')}</w:p>`

export const body = (file: Uint8Array) => textOf(readPart(file, 'word/document.xml'))
