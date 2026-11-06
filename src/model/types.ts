/**
 * Document model exchanged with the engine (`parse` / `build`). Plain JSON-serializable data.
 * The engine is the source of truth for the schema; `modelVersion` tracks breaking changes.
 */

// ---------------------------------------------------------------------------------------------
// Word
// ---------------------------------------------------------------------------------------------

export interface DocxDocument {
  format: 'docx'
  /** Model schema version set by the engine on `parse`. */
  modelVersion?: number
  parts: DocxPart[]
}

export interface DocxPart {
  /** Package path, e.g. `word/document.xml`, `word/header1.xml`. */
  path: string
  kind: 'body' | 'header' | 'footer'
  blocks: DocxBlock[]
  /** Final section properties of the body (page size, margins, header references). */
  sectionProps?: string
}

export type DocxBlock = DocxParagraph | DocxTable | DocxRaw

export interface DocxParagraph {
  type: 'paragraph'
  /** Paragraph style id, e.g. `Heading1`, `ListParagraph`. */
  style?: string
  align?: 'left' | 'center' | 'right' | 'justify'
  /** Numbering (bullets / numbered lists). */
  list?: { id: number; level: number }
  children: DocxInline[]
  /** Original `<w:pPr>` XML. Fields above override it when changed. */
  props?: string
}

export interface DocxRunStyle {
  /** Character style id. */
  style?: string
  bold?: boolean
  italic?: boolean
  /** `true` = single underline, or an underline type (`double`, `wave`, ...). */
  underline?: boolean | string
  strike?: boolean
  /** Hex color without `#`, e.g. `FF0000`. */
  color?: string
  /** Highlight color name, e.g. `yellow`. */
  highlight?: string
  /** Font size in points. */
  size?: number
  font?: string
  vertAlign?: 'superscript' | 'subscript' | 'baseline'
}

export interface DocxRun extends DocxRunStyle {
  type: 'run'
  /** Text; `\n` is a line break and `\t` a tab. */
  text: string
  /** Original `<w:rPr>` XML. Fields override it when changed. */
  props?: string
}

export interface DocxBreak {
  type: 'break'
  kind: 'page' | 'column'
}

export interface DocxHyperlink {
  type: 'hyperlink'
  /** External URL. */
  url?: string
  /** Internal bookmark name. */
  anchor?: string
  children: DocxInline[]
}

/** Anything not modelled (images, fields, bookmarks, content controls...), kept as XML. */
export interface DocxRaw {
  type: 'raw'
  xml: string
}

export type DocxInline = DocxRun | DocxBreak | DocxHyperlink | DocxRaw

export interface DocxTable {
  type: 'table'
  rows: DocxTableRow[]
  /** Original `<w:tblPr>` XML. */
  props?: string
  /** Original `<w:tblGrid>` XML. */
  grid?: string
}

export interface DocxTableRow {
  cells: DocxTableCell[]
  /** Original `<w:trPr>` XML. */
  props?: string
}

export interface DocxTableCell {
  blocks: DocxBlock[]
  colSpan?: number
  vMerge?: 'restart' | 'continue'
  /** Original `<w:tcPr>` XML. */
  props?: string
}

// ---------------------------------------------------------------------------------------------
// Excel
// ---------------------------------------------------------------------------------------------

export interface XlsxDocument {
  format: 'xlsx'
  /** Model schema version set by the engine on `parse`. */
  modelVersion?: number
  sheets: XlsxSheet[]
}

export interface XlsxSheet {
  name: string
  /** Package path of an existing sheet. Omit to create a new sheet on build. */
  path?: string
  rows: XlsxRow[]
  /** Merged ranges, e.g. `A1:C1`. */
  merges?: string[]
}

export interface XlsxRow {
  /** 1-based row number. */
  index: number
  /** Custom height in points. */
  height?: number
  hidden?: boolean
  cells: XlsxCell[]
}

export type XlsxValue = string | number | boolean | null

export interface XlsxCell {
  /** A1-style reference. When omitted, the cell goes to the column after the previous one. */
  ref?: string
  /** Strings, numbers and booleans map to cell types; dates are ISO strings with `type: 'date'`. */
  value: XlsxValue | Date
  /** `date` for date-formatted numbers (value is an ISO string), `error` for `#N/A` & co. */
  type?: 'date' | 'error'
  /** Formula without the leading `=`. `value` holds the cached result. */
  formula?: string
  /** Index into the workbook cell formats (`cellXfs`). */
  style?: number
  /** Original rich-text XML (`<r>` runs). Used when `value` still equals its plain text. */
  richText?: string
}

// ---------------------------------------------------------------------------------------------
// PowerPoint
// ---------------------------------------------------------------------------------------------

export interface PptxDocument {
  format: 'pptx'
  /** Model schema version set by the engine on `parse`. */
  modelVersion?: number
  /** Slides in presentation order. Reorder or remove items to reorder/remove slides. */
  slides: PptxSlide[]
}

export interface PptxSlide {
  /** Package path of the slide, e.g. `ppt/slides/slide1.xml`. */
  path: string
  /** Shapes in z-order. Shapes missing from the list are removed on build. */
  shapes: PptxShape[]
  /** Speaker notes as plain text. */
  notes?: string
}

export interface ShapeBase {
  /** Shape id (`p:cNvPr/@id`), used to match shapes on build. */
  id: number
  name: string
}

export interface PptxTextShape extends ShapeBase {
  type: 'text'
  /** Placeholder type when the shape is a layout placeholder: `title`, `body`, `subTitle`, ... */
  placeholder?: string
  paragraphs: PptxParagraph[]
}

export interface PptxTableShape extends ShapeBase {
  type: 'table'
  rows: Array<{ cells: Array<{ paragraphs: PptxParagraph[] }> }>
}

export interface PptxOtherShape extends ShapeBase {
  type: 'picture' | 'group' | 'connector' | 'graphic' | 'other'
  /** Raw XML. Only needed to insert a new shape; existing shapes are taken from the base file. */
  xml?: string
}

export type PptxShape = PptxTextShape | PptxTableShape | PptxOtherShape

export interface PptxParagraph {
  align?: 'left' | 'center' | 'right' | 'justify'
  /** Outline (bullet) level, 0-based. */
  level?: number
  children: PptxInline[]
  /** Original `<a:pPr>` XML. */
  props?: string
  /** Original `<a:endParaRPr>` XML. */
  endProps?: string
}

export interface PptxRun {
  type: 'run'
  /** Text; `\n` is a line break. */
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean | string
  strike?: boolean
  /** Font size in points. */
  size?: number
  /** Hex RGB color without `#`. */
  color?: string
  font?: string
  /** Original `<a:rPr>` XML. */
  props?: string
}

/** Fields (slide number, date) and other inline content, kept as XML. */
export interface PptxRaw {
  type: 'raw'
  xml: string
}

export type PptxInline = PptxRun | PptxRaw

export type OfficeDocument = DocxDocument | XlsxDocument | PptxDocument
