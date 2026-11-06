import { request } from '../engine/core'
import { toBytes, type BinaryInput } from '../engine/values'
import type { DocumentFormat, Limits } from '../types'
import { modelText } from './text'
import type { DocxDocument, OfficeDocument, PptxDocument, XlsxDocument } from './types'

export interface ParseOptions {
  /** Safety limits for the input file. */
  limits?: Limits
}

/**
 * Parses a .docx/.xlsx/.pptx file into a plain JSON-serializable model.
 * Edit it and pass it to `build()` (together with the original file) to get a document back.
 */
export function parse(input: BinaryInput, options: ParseOptions = {}): OfficeDocument {
  const limits = options.limits ? `,"limits":${JSON.stringify(options.limits)}` : ''
  return request<{ model: OfficeDocument }>(`{"op":"parse","document":{"$blob":0}${limits}}`, [toBytes(input)]).result.model
}

export function parseDocxFile(input: BinaryInput, options?: ParseOptions): DocxDocument {
  return expectFormat(parse(input, options), 'docx')
}
export function parseXlsxFile(input: BinaryInput, options?: ParseOptions): XlsxDocument {
  return expectFormat(parse(input, options), 'xlsx')
}
export function parsePptxFile(input: BinaryInput, options?: ParseOptions): PptxDocument {
  return expectFormat(parse(input, options), 'pptx')
}

function expectFormat<F extends OfficeDocument['format']>(doc: OfficeDocument, format: F): Extract<OfficeDocument, { format: F }> {
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
  const blobs = options.base ? [toBytes(options.base)] : []
  const base = options.base ? ',"base":{"$blob":0}' : ''
  const json = `{"op":"build","model":${JSON.stringify(model, modelReplacer)}${base}}`
  return request(json, blobs).blobs[0]
}

/** `Date` cell values (e.g. from `sheetFromArrays`) travel as ISO strings with `type: 'date'`. */
function modelReplacer(_key: string, value: unknown): unknown {
  // Date#toJSON already yields an ISO string; the cell just needs to be marked as a date.
  if (value && typeof value === 'object' && !Array.isArray(value) && (value as { value?: unknown }).value instanceof Date) {
    return { ...value, type: 'date' }
  }
  return value
}

/** Plain text of a document: paragraphs (docx), tab-separated cells (xlsx), slide texts (pptx). */
export function extractText(input: BinaryInput | OfficeDocument): string {
  if (isModel(input)) return modelText(input)
  return request<{ text: string }>('{"op":"extractText","document":{"$blob":0}}', [toBytes(input)]).result.text
}

/** Detects the format of an Office file. */
export function detectFormat(input: BinaryInput): DocumentFormat {
  return request<{ format: DocumentFormat }>('{"op":"detect","document":{"$blob":0}}', [toBytes(input)]).result.format
}

function isModel(x: unknown): x is OfficeDocument {
  return typeof x === 'object' && x !== null && 'format' in x && !ArrayBuffer.isView(x) && !(x instanceof ArrayBuffer)
}
