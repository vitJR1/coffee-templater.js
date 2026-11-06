import { detectFormat } from '../formats/adapter'
import { OpcPackage, type BinaryInput } from '../package'
import { blankDocx, blankXlsx } from './blank'
import { buildDocx, docxText, parseDocx, type DocxDocument } from './docx'
import { buildPptx, parsePptx, pptxText, type PptxDocument } from './pptx'
import { buildXlsx, parseXlsx, xlsxText, type XlsxDocument } from './xlsx'

export type OfficeDocument = DocxDocument | XlsxDocument | PptxDocument

/**
 * Parses a .docx/.xlsx/.pptx file into a plain JSON-serializable model.
 * Edit it and pass it to `build()` (together with the original file) to get a document back.
 */
export function parse(input: BinaryInput): OfficeDocument {
  const pkg = OpcPackage.load(input)
  switch (detectFormat(pkg)) {
    case 'docx':
      return parseDocx(pkg)
    case 'xlsx':
      return parseXlsx(pkg)
    case 'pptx':
      return parsePptx(pkg)
  }
}

export function parseDocxFile(input: BinaryInput): DocxDocument {
  return expect(parse(input), 'docx')
}
export function parseXlsxFile(input: BinaryInput): XlsxDocument {
  return expect(parse(input), 'xlsx')
}
export function parsePptxFile(input: BinaryInput): PptxDocument {
  return expect(parse(input), 'pptx')
}

function expect<F extends OfficeDocument['format']>(doc: OfficeDocument, format: F): Extract<OfficeDocument, { format: F }> {
  if (doc.format !== format) throw new Error(`Expected a ${format} file, got ${doc.format}`)
  return doc as Extract<OfficeDocument, { format: F }>
}

export interface BuildOptions {
  /**
   * The original file the model was parsed from. Styles, media, settings and everything else not in
   * the model are taken from it. Optional for docx/xlsx (a blank document is used), required for pptx.
   */
  base?: BinaryInput
}

/** Builds a document from a model. */
export function build(model: OfficeDocument, options: BuildOptions = {}): Uint8Array {
  let pkg: OpcPackage
  if (options.base) {
    pkg = OpcPackage.load(options.base)
    const format = detectFormat(pkg)
    if (format !== model.format) throw new Error(`Base document is ${format}, model is ${model.format}`)
  } else if (model.format === 'docx') pkg = blankDocx()
  else if (model.format === 'xlsx') pkg = blankXlsx()
  else throw new Error('Building a pptx requires `base` (the original presentation)')

  if (model.format === 'docx') {
    if (!options.base) {
      // blank document: only the body exists
      model = { ...model, parts: model.parts.filter((p) => p.kind === 'body').map((p) => ({ ...p, path: 'word/document.xml' })) }
    }
    buildDocx(model as DocxDocument, pkg)
  } else if (model.format === 'xlsx') {
    if (!options.base) model = { ...model, sheets: model.sheets.map((s) => ({ ...s, path: undefined })) }
    buildXlsx(model as XlsxDocument, pkg)
  } else buildPptx(model, pkg)
  return pkg.toBytes()
}

/** Plain text of a document: paragraphs (docx), tab-separated cells (xlsx), slide texts (pptx). */
export function extractText(input: BinaryInput | OfficeDocument): string {
  const doc = isModel(input) ? input : parse(input)
  if (doc.format === 'docx') return docxText(doc)
  if (doc.format === 'xlsx') return xlsxText(doc)
  return pptxText(doc)
}

function isModel(x: unknown): x is OfficeDocument {
  return typeof x === 'object' && x !== null && 'format' in x && !(x instanceof Uint8Array)
}
