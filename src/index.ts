// Templating
export { Templater, Template, createTemplater, render } from './templater'
export type { InspectResult, TemplatePartInfo } from './templater'
export type { TemplaterOptions, TagInfo, TagSyntax, FilterFn } from './template/options'
export type { TemplateTag } from './template/engine'
export type { SectionMode } from './template/tags'
export { builtinFilters, formatDate, formatNumber } from './template/filters'
export { TemplateError } from './errors'
export type { TemplateErrorCode } from './errors'
export type { ImageValue } from './image'

// Document model (parse → JSON → build)
export { parse, parseDocxFile, parseXlsxFile, parsePptxFile, build, extractText } from './model'
export type { OfficeDocument, BuildOptions } from './model'
export type {
  DocxDocument, DocxPart, DocxBlock, DocxParagraph, DocxRun, DocxRunStyle, DocxBreak, DocxHyperlink, DocxRaw,
  DocxInline, DocxTable, DocxTableRow, DocxTableCell,
} from './model/docx'
export type { XlsxDocument, XlsxSheet, XlsxRow, XlsxCell, XlsxValue } from './model/xlsx'
export {
  sheetToArrays, sheetToObjects, sheetFromArrays, columnToIndex, indexToColumn, parseRef, serialToDate, dateToSerial,
} from './model/xlsx'
export type {
  PptxDocument, PptxSlide, PptxShape, PptxTextShape, PptxTableShape, PptxOtherShape, PptxParagraph, PptxRun, PptxRaw, PptxInline,
} from './model/pptx'

export type { DocumentFormat } from './formats/adapter'
export type { BinaryInput } from './package'
