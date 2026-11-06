// Templating
export { Templater, Template, createTemplater, render } from './templater'
export type {
  TemplaterOptions, TagInfo, TagSyntax, FilterFn, Limits, InspectResult, TemplatePartInfo, TemplateTag, SectionMode,
  ImageValue, DocumentFormat, BinaryInput,
} from './types'
export { TemplateError } from './errors'
export type { TemplateErrorCode, TemplateErrorInfo } from './errors'

// Document model (parse → JSON → build)
export { parse, parseDocxFile, parseXlsxFile, parsePptxFile, build, extractText, detectFormat } from './model'
export type { ParseOptions, BuildOptions } from './model'
export type {
  OfficeDocument,
  DocxDocument, DocxPart, DocxBlock, DocxParagraph, DocxRun, DocxRunStyle, DocxBreak, DocxHyperlink, DocxRaw,
  DocxInline, DocxTable, DocxTableRow, DocxTableCell,
  XlsxDocument, XlsxSheet, XlsxRow, XlsxCell, XlsxValue,
  PptxDocument, PptxSlide, PptxShape, PptxTextShape, PptxTableShape, PptxOtherShape, PptxParagraph, PptxRun, PptxRaw,
  PptxInline,
} from './model/types'
export {
  sheetToArrays, sheetToObjects, sheetFromArrays, columnToIndex, indexToColumn, parseRef, serialToDate, dateToSerial,
} from './model/sheets'

// Engine
export { initEngine, resetEngine, engineInfo, ENGINE_VERSION } from './engine/core'
export type { EngineInfo } from './engine/core'
