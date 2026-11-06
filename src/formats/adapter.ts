import type { OpcPackage } from '../package'
import type { FormatSpec, RenderContext } from '../template/engine'

export type DocumentFormat = 'docx' | 'xlsx' | 'pptx'

/** Glue between the generic template engine and a concrete Office format. */
export interface FormatAdapter {
  format: DocumentFormat
  spec: FormatSpec
  /** Parts (XML files) that contain templatable text, in document order. */
  templateParts(pkg: OpcPackage): string[]
  /** Optional rewrite of part XML before compilation. May store data in `state`. */
  preprocess?(pkg: OpcPackage, part: string, xml: string, state: Map<string, unknown>): string
  /** Optional rewrite of rendered part XML. */
  postprocess?(pkg: OpcPackage, part: string, xml: string, state: Map<string, unknown>, ctx: RenderContext): string
  /** Called once after all parts were rendered. */
  finalize?(pkg: OpcPackage, ctx: RenderContext): void
}

export function detectFormat(pkg: OpcPackage): DocumentFormat {
  const main = pkg.mainDocumentPath() ?? ''
  const ct = pkg.getText('[Content_Types].xml') ?? ''
  if (main.startsWith('word/') || ct.includes('wordprocessingml.document')) return 'docx'
  if (main.startsWith('xl/') || ct.includes('spreadsheetml.sheet')) return 'xlsx'
  if (main.startsWith('ppt/') || ct.includes('presentationml.presentation')) return 'pptx'
  throw new Error('Unsupported document: expected a .docx, .xlsx or .pptx file')
}
