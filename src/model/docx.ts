import { docxMainPart } from '../formats/docx/adapter'
import { REL, type OpcPackage } from '../package'
import { childElements, el, firstChild, parseFragment, parseXml, serializeNode, serializeXml, type XmlElement, type XmlNode } from '../xml/dom'
import { escapeXml } from '../xml/escape'
import { onOff, readFields, serializeEl, valueOf, writeProps, type FieldMap } from './props'

// ---------------------------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------------------------

export interface DocxDocument {
  format: 'docx'
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
// Schema (element order matters to Word)
// ---------------------------------------------------------------------------------------------

const RPR_ORDER = [
  'w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline',
  'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing',
  'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText',
  'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath', 'w:rPrChange',
]
const PPR_ORDER = [
  'w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr', 'w:widowControl', 'w:numPr',
  'w:suppressLineNumbers', 'w:pBdr', 'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
  'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd', 'w:snapToGrid',
  'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection',
  'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange',
]
const TCPR_ORDER = ['w:cnfStyle', 'w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge', 'w:tcBorders', 'w:shd', 'w:noWrap', 'w:tcMar', 'w:textDirection', 'w:tcFitText', 'w:vAlign', 'w:hideMark']

const RUN_FIELDS: FieldMap = {
  style: valueOf('w:rStyle'),
  bold: onOff('w:b'),
  italic: onOff('w:i'),
  strike: onOff('w:strike'),
  underline: {
    names: ['w:u'],
    read: (c: XmlElement) => {
      const v = c.attrs['w:val'] ?? 'single'
      return v === 'none' ? false : v === 'single' ? true : v
    },
    write: (v: boolean | string) => [el('w:u', { 'w:val': v === true ? 'single' : v === false ? 'none' : v })],
  },
  color: valueOf('w:color'),
  highlight: valueOf('w:highlight'),
  size: {
    names: ['w:sz', 'w:szCs'],
    read: (c: XmlElement) => (c.attrs['w:val'] ? Number(c.attrs['w:val']) / 2 : undefined),
    write: (v: number) => [el('w:sz', { 'w:val': String(Math.round(v * 2)) }), el('w:szCs', { 'w:val': String(Math.round(v * 2)) })],
  },
  font: {
    names: ['w:rFonts'],
    read: (c: XmlElement) => c.attrs['w:ascii'] ?? c.attrs['w:hAnsi'],
    write: (v: string) => [el('w:rFonts', { 'w:ascii': v, 'w:hAnsi': v, 'w:cs': v, 'w:eastAsia': v })],
  },
  vertAlign: valueOf('w:vertAlign'),
}

const ALIGN_IN: Record<string, DocxParagraph['align']> = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }
const ALIGN_OUT: Record<string, string> = { left: 'left', center: 'center', right: 'right', justify: 'both' }

const PARA_FIELDS: FieldMap = {
  style: valueOf('w:pStyle'),
  align: { names: ['w:jc'], read: (c: XmlElement) => ALIGN_IN[c.attrs['w:val']], write: (v: string) => [el('w:jc', { 'w:val': ALIGN_OUT[v] ?? v })] },
  list: {
    names: ['w:numPr'],
    read: (c: XmlElement) => {
      const id = firstChild(c, 'w:numId')?.attrs['w:val']
      if (id === undefined) return undefined
      return { id: Number(id), level: Number(firstChild(c, 'w:ilvl')?.attrs['w:val'] ?? 0) }
    },
    write: (v: { id: number; level?: number }) => [
      el('w:numPr', {}, [el('w:ilvl', { 'w:val': String(v.level ?? 0) }), el('w:numId', { 'w:val': String(v.id) })]),
    ],
  },
}

const CELL_FIELDS: FieldMap = {
  colSpan: { names: ['w:gridSpan'], read: (c: XmlElement) => Number(c.attrs['w:val']), write: (v: number) => (v > 1 ? [el('w:gridSpan', { 'w:val': String(v) })] : []) },
  vMerge: {
    names: ['w:vMerge'],
    read: (c: XmlElement) => (c.attrs['w:val'] === 'restart' ? 'restart' : 'continue'),
    write: (v: string) => [v === 'restart' ? el('w:vMerge', { 'w:val': 'restart' }) : el('w:vMerge')],
  },
}

// ---------------------------------------------------------------------------------------------
// Parse
// ---------------------------------------------------------------------------------------------

interface ParseCtx {
  rels: Map<string, string>
}

export function parseDocx(pkg: OpcPackage): DocxDocument {
  const main = docxMainPart(pkg)
  const parts: DocxPart[] = []
  const add = (path: string, kind: DocxPart['kind']) => {
    const xml = pkg.getText(path)
    if (!xml) return
    const ctx: ParseCtx = {
      rels: new Map(pkg.getRelationships(path).filter((r) => r.type === REL.hyperlink).map((r) => [r.id, r.target])),
    }
    const root = parseXml(xml).root
    const container = kind === 'body' ? firstChild(root, 'w:body') : root
    if (!container) return
    const part: DocxPart = { path, kind, blocks: [] }
    for (const child of container.children) {
      if (child.type !== 'element') continue
      if (kind === 'body' && child.name === 'w:sectPr') part.sectionProps = serializeNode(child)
      else part.blocks.push(parseBlock(child, ctx))
    }
    parts.push(part)
  }
  add(main, 'body')
  for (const { path } of pkg.getRelatedParts(main, REL.header)) add(path, 'header')
  for (const { path } of pkg.getRelatedParts(main, REL.footer)) add(path, 'footer')
  return { format: 'docx', parts }
}

function parseBlock(node: XmlElement, ctx: ParseCtx): DocxBlock {
  if (node.name === 'w:p') return parseParagraph(node, ctx)
  if (node.name === 'w:tbl') return parseTable(node, ctx)
  return { type: 'raw', xml: serializeNode(node) }
}

function parseParagraph(p: XmlElement, ctx: ParseCtx): DocxParagraph {
  const pPr = firstChild(p, 'w:pPr')
  const para: DocxParagraph = { type: 'paragraph', ...(readFields(pPr, PARA_FIELDS) as Partial<DocxParagraph>), children: [] }
  if (pPr) para.props = serializeNode(pPr)
  para.children = parseInlines(p.children, ctx)
  return para
}

const SKIP = new Set(['w:pPr', 'w:proofErr', 'w:lastRenderedPageBreak'])

function parseInlines(nodes: XmlNode[], ctx: ParseCtx): DocxInline[] {
  const out: DocxInline[] = []
  for (const child of nodes) {
    if (child.type !== 'element' || SKIP.has(child.name)) continue
    if (child.name === 'w:r') out.push(...parseRun(child))
    else if (child.name === 'w:hyperlink') {
      const link: DocxHyperlink = { type: 'hyperlink', children: parseInlines(child.children, ctx) }
      const rId = child.attrs['r:id']
      if (rId && ctx.rels.has(rId)) link.url = ctx.rels.get(rId)
      if (child.attrs['w:anchor']) link.anchor = child.attrs['w:anchor']
      out.push(link)
    } else out.push({ type: 'raw', xml: serializeNode(child) })
  }
  return mergeRuns(out)
}

function parseRun(r: XmlElement): DocxInline[] {
  const rPr = firstChild(r, 'w:rPr')
  const style = readFields(rPr, RUN_FIELDS) as DocxRunStyle
  const propsXml = serializeEl(rPr)
  const out: DocxInline[] = []
  let text: string | undefined
  const flush = () => {
    if (text === undefined) return
    const run: DocxRun = { type: 'run', text, ...style }
    if (propsXml) run.props = propsXml
    out.push(run)
    text = undefined
  }
  const append = (s: string) => (text = (text ?? '') + s)
  for (const c of r.children) {
    if (c.type !== 'element' || c.name === 'w:rPr' || c.name === 'w:lastRenderedPageBreak') continue
    switch (c.name) {
      case 'w:t':
        append(c.children.map((n) => (n.type === 'text' ? n.text : '')).join(''))
        break
      case 'w:tab':
        append('\t')
        break
      case 'w:cr':
        append('\n')
        break
      case 'w:noBreakHyphen':
        append('‑')
        break
      case 'w:softHyphen':
        append('­')
        break
      case 'w:br': {
        const kind = c.attrs['w:type']
        if (kind === 'page' || kind === 'column') {
          flush()
          out.push({ type: 'break', kind })
        } else append('\n')
        break
      }
      default:
        flush()
        out.push({ type: 'raw', xml: `<w:r>${propsXml ?? ''}${serializeNode(c)}</w:r>` })
    }
  }
  flush()
  return out
}

/** Merges neighbouring runs with identical formatting (Word splits them for spell-check/rsid noise). */
function mergeRuns(list: DocxInline[]): DocxInline[] {
  const out: DocxInline[] = []
  for (const item of list) {
    const prev = out[out.length - 1]
    if (item.type === 'run' && prev?.type === 'run' && sameStyle(prev, item)) prev.text += item.text
    else out.push(item)
  }
  return out
}

function sameStyle(a: DocxRun, b: DocxRun): boolean {
  const strip = (r: DocxRun) => normalizeRunProps(r.props)
  return JSON.stringify({ ...a, text: '', props: strip(a) }) === JSON.stringify({ ...b, text: '', props: strip(b) })
}

/** rsid attributes differ between otherwise identical runs; ignore them when comparing. */
function normalizeRunProps(xml?: string): string {
  return (xml ?? '').replace(/\sw:rsid\w*="[^"]*"/g, '')
}

function parseTable(tbl: XmlElement, ctx: ParseCtx): DocxTable {
  const table: DocxTable = { type: 'table', rows: [] }
  const tblPr = firstChild(tbl, 'w:tblPr')
  const grid = firstChild(tbl, 'w:tblGrid')
  if (tblPr) table.props = serializeNode(tblPr)
  if (grid) table.grid = serializeNode(grid)
  for (const tr of childElements(tbl, 'w:tr')) {
    const trPr = firstChild(tr, 'w:trPr')
    const row: DocxTableRow = { cells: [] }
    if (trPr) row.props = serializeNode(trPr)
    for (const tc of childElements(tr, 'w:tc')) {
      const tcPr = firstChild(tc, 'w:tcPr')
      const cell: DocxTableCell = { ...(readFields(tcPr, CELL_FIELDS) as Partial<DocxTableCell>), blocks: [] }
      if (tcPr) cell.props = serializeNode(tcPr)
      for (const b of childElements(tc)) if (b.name !== 'w:tcPr') cell.blocks.push(parseBlock(b, ctx))
      row.cells.push(cell)
    }
    table.rows.push(row)
  }
  return table
}

// ---------------------------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------------------------

interface BuildCtx {
  pkg: OpcPackage
  part: string
}

export function buildDocx(model: DocxDocument, pkg: OpcPackage): void {
  for (const part of model.parts) {
    const xml = pkg.getText(part.path)
    if (!xml) throw new Error(`Part ${part.path} does not exist in the base document`)
    const doc = parseXml(xml)
    const container = part.kind === 'body' ? firstChild(doc.root, 'w:body') : doc.root
    if (!container) throw new Error(`Part ${part.path} has no body`)
    const ctx: BuildCtx = { pkg, part: part.path }
    let inner = part.blocks.map((b) => buildBlock(b, ctx)).join('')
    if (part.kind === 'body') {
      const sect = part.sectionProps ?? serializeEl(firstChild(container, 'w:sectPr'))
      if (sect) inner += sect
    }
    container.children = parseFragment(inner)
    if (inner.includes(' r:id=')) doc.root.attrs['xmlns:r'] ??= 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg.setText(part.path, serializeXml(doc))
  }
}

function buildBlock(b: DocxBlock, ctx: BuildCtx): string {
  if (b.type === 'raw') return b.xml
  if (b.type === 'table') return buildTable(b, ctx)
  return buildParagraph(b, ctx)
}

function buildParagraph(p: DocxParagraph, ctx: BuildCtx): string {
  const pPr = writeProps('w:pPr', p.props, p as unknown as Record<string, unknown>, PARA_FIELDS, PPR_ORDER)
  return `<w:p>${pPr}${(p.children ?? []).map((c) => buildInline(c, ctx)).join('')}</w:p>`
}

function buildInline(i: DocxInline, ctx: BuildCtx): string {
  switch (i.type) {
    case 'raw':
      return i.xml
    case 'break':
      return `<w:r><w:br w:type="${i.kind}"/></w:r>`
    case 'hyperlink': {
      const attrs: string[] = []
      if (i.url) {
        const existing = ctx.pkg.getRelationships(ctx.part).find((r) => r.type === REL.hyperlink && r.target === i.url)
        const id = existing?.id ?? ctx.pkg.addRelationship(ctx.part, REL.hyperlink, i.url, true)
        attrs.push(`r:id="${id}"`)
      }
      if (i.anchor) attrs.push(`w:anchor="${escapeXml(i.anchor).replace(/"/g, '&quot;')}"`)
      attrs.push('w:history="1"')
      return `<w:hyperlink ${attrs.join(' ')}>${i.children.map((c) => buildInline(c, ctx)).join('')}</w:hyperlink>`
    }
    case 'run': {
      const rPr = writeProps('w:rPr', i.props, i as unknown as Record<string, unknown>, RUN_FIELDS, RPR_ORDER)
      let content = ''
      for (const piece of (i.text ?? '').split(/(\r\n|\n|\r|\t|‑|­)/)) {
        if (!piece) continue
        if (piece === '\t') content += '<w:tab/>'
        else if (piece === '‑') content += '<w:noBreakHyphen/>'
        else if (piece === '­') content += '<w:softHyphen/>'
        else if (/^[\r\n]+$/.test(piece)) content += '<w:br/>'
        else content += `<w:t xml:space="preserve">${escapeXml(piece)}</w:t>`
      }
      return `<w:r>${rPr}${content}</w:r>`
    }
  }
}

function buildTable(t: DocxTable, ctx: BuildCtx): string {
  const cols = Math.max(1, ...t.rows.map((r) => r.cells.reduce((n, c) => n + (c.colSpan ?? 1), 0)))
  const tblPr = t.props ?? '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
  const grid = t.grid ?? `<w:tblGrid>${'<w:gridCol/>'.repeat(cols)}</w:tblGrid>`
  const rows = t.rows
    .map((r) => {
      const cells = r.cells
        .map((c) => {
          const tcPr = writeProps('w:tcPr', c.props, c as unknown as Record<string, unknown>, CELL_FIELDS, TCPR_ORDER)
          let blocks = (c.blocks ?? []).map((b) => buildBlock(b, ctx)).join('')
          // A cell must end with a paragraph.
          const last = c.blocks?.[c.blocks.length - 1]
          if (!last || last.type !== 'paragraph') blocks += '<w:p/>'
          return `<w:tc>${tcPr}${blocks}</w:tc>`
        })
        .join('')
      return `<w:tr>${r.props ?? ''}${cells}</w:tr>`
    })
    .join('')
  return `<w:tbl>${tblPr}${grid}${rows}</w:tbl>`
}

// ---------------------------------------------------------------------------------------------
// Text helpers
// ---------------------------------------------------------------------------------------------

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

export function docxText(model: DocxDocument): string {
  const body = model.parts.find((p) => p.kind === 'body')
  return body ? blocksText(body.blocks) : ''
}

