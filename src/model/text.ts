import { sheetToArrays } from './sheets'
import type { DocxBlock, DocxInline, OfficeDocument, PptxParagraph } from './types'

export function inlineText(children: DocxInline[]): string {
  return children
    .map((c) => (c.type === 'run' ? c.text : c.type === 'hyperlink' ? inlineText(c.children) : c.type === 'break' ? '\n' : ''))
    .join('')
}

export function blocksText(blocks: DocxBlock[]): string {
  return blocks
    .map((b) => {
      if (b.type === 'paragraph') return inlineText(b.children)
      if (b.type === 'table') return b.rows.map((r) => r.cells.map((c) => blocksText(c.blocks).replace(/\n/g, ' ')).join('\t')).join('\n')
      return ''
    })
    .join('\n')
}

function xlsxText(model: Extract<OfficeDocument, { format: 'xlsx' }>): string {
  return model.sheets
    .map((s) => `${s.name}\n` + sheetToArrays(s).map((r) => r.map((v) => (v === null ? '' : String(v))).join('\t')).join('\n'))
    .join('\n\n')
}

export function paragraphText(p: PptxParagraph): string {
  return p.children.map((c) => (c.type === 'run' ? c.text : '')).join('')
}

function pptxText(model: Extract<OfficeDocument, { format: 'pptx' }>): string {
  return model.slides
    .map((s, i) => {
      const lines: string[] = [`--- Slide ${i + 1} ---`]
      for (const sh of s.shapes) {
        if (sh.type === 'text') lines.push(sh.paragraphs.map(paragraphText).join('\n'))
        if (sh.type === 'table') lines.push(sh.rows.map((r) => r.cells.map((c) => c.paragraphs.map(paragraphText).join(' ')).join('\t')).join('\n'))
      }
      if (s.notes) lines.push(`Notes: ${s.notes}`)
      return lines.filter(Boolean).join('\n')
    })
    .join('\n\n')
}

/** Plain text of a parsed model (same output as the engine's `extractText`). */
export function modelText(doc: OfficeDocument): string {
  if (doc.format === 'docx') {
    const body = doc.parts.find((p) => p.kind === 'body')
    return body ? blocksText(body.blocks) : ''
  }
  if (doc.format === 'xlsx') return xlsxText(doc)
  return pptxText(doc)
}
