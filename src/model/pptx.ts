import { presentationPath, slideParts } from '../formats/pptx/adapter'
import { REL, type OpcPackage } from '../package'
import { childElements, el, findDescendant, findDescendants, firstChild, parseFragment, parseXml, serializeNode, serializeXml, type XmlElement, type XmlNode } from '../xml/dom'
import { escapeXml } from '../xml/escape'
import { readFields, writeProps, type FieldMap } from './props'

// ---------------------------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------------------------

export interface PptxDocument {
  format: 'pptx'
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

interface ShapeBase {
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

// ---------------------------------------------------------------------------------------------
// Schema
// ---------------------------------------------------------------------------------------------

const FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill']
const RPR_ORDER = ['a:ln', ...FILLS, 'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill', 'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst']

const bool = (attr: string) => ({ attr, read: (v: string) => v === '1' || v === 'true', write: (v: boolean) => (v ? '1' : '0') })

const RUN_FIELDS: FieldMap = {
  bold: bool('b'),
  italic: bool('i'),
  underline: {
    attr: 'u',
    read: (v: string) => (v === 'none' ? false : v === 'sng' ? true : v),
    write: (v: boolean | string) => (v === true ? 'sng' : v === false ? 'none' : v),
  },
  strike: { attr: 'strike', read: (v: string) => v !== 'noStrike', write: (v: boolean) => (v ? 'sngStrike' : 'noStrike') },
  size: { attr: 'sz', read: (v: string) => Number(v) / 100, write: (v: number) => String(Math.round(v * 100)) },
  color: {
    names: ['a:solidFill', ...FILLS.filter((f) => f !== 'a:solidFill')],
    read: (c: XmlElement) => firstChild(c, 'a:srgbClr')?.attrs.val,
    write: (v: string) => [el('a:solidFill', {}, [el('a:srgbClr', { val: v.replace(/^#/, '').toUpperCase() })])],
  },
  font: { names: ['a:latin'], read: (c: XmlElement) => c.attrs.typeface, write: (v: string) => [el('a:latin', { typeface: v })] },
}

const ALIGN_IN: Record<string, PptxParagraph['align']> = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }
const ALIGN_OUT: Record<string, string> = { left: 'l', center: 'ctr', right: 'r', justify: 'just' }

const PARA_FIELDS: FieldMap = {
  align: { attr: 'algn', read: (v: string) => ALIGN_IN[v], write: (v: string) => ALIGN_OUT[v] ?? v },
  level: { attr: 'lvl', read: (v: string) => Number(v), write: (v: number) => String(v) },
}

// ---------------------------------------------------------------------------------------------
// Parse
// ---------------------------------------------------------------------------------------------

export function parsePptx(pkg: OpcPackage): PptxDocument {
  return {
    format: 'pptx',
    slides: slideParts(pkg).map(({ path }) => {
      const root = parseXml(pkg.getText(path)!).root
      const tree = findDescendant(root, 'p:spTree')
      const slide: PptxSlide = { path, shapes: tree ? shapeChildren(tree).map(parseShape) : [] }
      const notes = notesText(pkg, path)
      if (notes !== undefined) slide.notes = notes
      return slide
    }),
  }
}

const SHAPE_TAGS = new Set(['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp', 'mc:AlternateContent', 'p:contentPart'])

function shapeChildren(tree: XmlElement): XmlElement[] {
  return childElements(tree).filter((c) => SHAPE_TAGS.has(c.name))
}

function shapeId(node: XmlElement): { id: number; name: string } {
  const nv = findDescendant(node, 'p:cNvPr')
  return { id: Number(nv?.attrs.id ?? 0), name: nv?.attrs.name ?? '' }
}

function parseShape(node: XmlElement): PptxShape {
  const base = shapeId(node)
  if (node.name === 'p:sp') {
    const body = firstChild(node, 'p:txBody')
    if (body) {
      const shape: PptxTextShape = { type: 'text', ...base, paragraphs: childElements(body, 'a:p').map(parseParagraph) }
      const ph = findDescendant(node, 'p:ph')
      if (ph) shape.placeholder = ph.attrs.type ?? 'body'
      return shape
    }
    return { type: 'other', ...base }
  }
  if (node.name === 'p:graphicFrame') {
    const tbl = findDescendant(node, 'a:tbl')
    if (tbl) {
      return {
        type: 'table',
        ...base,
        rows: childElements(tbl, 'a:tr').map((tr) => ({
          cells: childElements(tr, 'a:tc').map((tc) => {
            const body = firstChild(tc, 'a:txBody')
            return { paragraphs: body ? childElements(body, 'a:p').map(parseParagraph) : [] }
          }),
        })),
      }
    }
    return { type: 'graphic', ...base }
  }
  const kind = node.name === 'p:pic' ? 'picture' : node.name === 'p:grpSp' ? 'group' : node.name === 'p:cxnSp' ? 'connector' : 'other'
  return { type: kind, ...base }
}

function parseParagraph(p: XmlElement): PptxParagraph {
  const pPr = firstChild(p, 'a:pPr')
  const para: PptxParagraph = { ...(readFields(pPr, PARA_FIELDS) as Partial<PptxParagraph>), children: [] }
  if (pPr) para.props = serializeNode(pPr)
  const end = firstChild(p, 'a:endParaRPr')
  if (end) para.endProps = serializeNode(end)
  for (const c of childElements(p)) {
    if (c.name === 'a:r') {
      const rPr = firstChild(c, 'a:rPr')
      const t = firstChild(c, 'a:t')
      const run: PptxRun = { type: 'run', text: t ? t.children.map((n) => (n.type === 'text' ? n.text : '')).join('') : '', ...(readFields(rPr, RUN_FIELDS) as Partial<PptxRun>) }
      if (rPr) run.props = serializeNode(rPr)
      para.children.push(run)
    } else if (c.name === 'a:br') {
      const prev = para.children[para.children.length - 1]
      if (prev?.type === 'run') prev.text += '\n'
      else {
        const rPr = firstChild(c, 'a:rPr')
        const run: PptxRun = { type: 'run', text: '\n', ...(readFields(rPr, RUN_FIELDS) as Partial<PptxRun>) }
        if (rPr) run.props = serializeNode(rPr)
        para.children.push(run)
      }
    } else if (c.name !== 'a:pPr' && c.name !== 'a:endParaRPr') {
      para.children.push({ type: 'raw', xml: serializeNode(c) })
    }
  }
  // merge runs with identical formatting
  const merged: PptxInline[] = []
  for (const item of para.children) {
    const prev = merged[merged.length - 1]
    if (item.type === 'run' && prev?.type === 'run' && JSON.stringify({ ...prev, text: '' }) === JSON.stringify({ ...item, text: '' })) prev.text += item.text
    else merged.push(item)
  }
  para.children = merged
  return para
}

function notesPart(pkg: OpcPackage, slide: string): string | undefined {
  const p = pkg.getRelatedParts(slide, REL.notesSlide)[0]?.path
  return p && pkg.has(p) ? p : undefined
}

function notesBody(root: XmlElement): XmlElement | undefined {
  for (const sp of findDescendants(root, 'p:sp')) {
    const ph = findDescendant(sp, 'p:ph')
    if (ph && ph.attrs.type === 'body') return firstChild(sp, 'p:txBody')
  }
  return undefined
}

function notesText(pkg: OpcPackage, slide: string): string | undefined {
  const path = notesPart(pkg, slide)
  if (!path) return undefined
  const body = notesBody(parseXml(pkg.getText(path)!).root)
  if (!body) return undefined
  return childElements(body, 'a:p').map((p) => paragraphText(parseParagraph(p))).join('\n')
}

export function paragraphText(p: PptxParagraph): string {
  return p.children.map((c) => (c.type === 'run' ? c.text : '')).join('')
}

// ---------------------------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------------------------

export function buildPptx(model: PptxDocument, pkg: OpcPackage): void {
  const pres = presentationPath(pkg)
  const existing = slideParts(pkg)
  const seen = new Set<string>()
  for (const s of model.slides) {
    if (!existing.some((e) => e.path === s.path)) throw new Error(`Slide ${s.path} does not exist in the base presentation`)
    if (seen.has(s.path)) throw new Error(`Slide ${s.path} is listed twice; duplicating slides is not supported`)
    seen.add(s.path)
  }

  // slide order / removal
  const presDoc = parseXml(pkg.getText(pres)!)
  const list = findDescendant(presDoc.root, 'p:sldIdLst')
  if (list) {
    const byRid = new Map(childElements(list, 'p:sldId').map((s) => [s.attrs['r:id'], s]))
    list.children = model.slides.map((s) => byRid.get(existing.find((e) => e.path === s.path)!.rId)!)
    pkg.setText(pres, serializeXml(presDoc))
  }
  for (const e of existing) {
    if (seen.has(e.path)) continue
    const notes = notesPart(pkg, e.path)
    if (notes) {
      pkg.remove(notes)
      pkg.remove(notes.replace(/([^/]+)$/, '_rels/$1.rels'))
      pkg.removeOverride(notes)
    }
    pkg.removeRelationship(pres, (r) => r.id === e.rId)
    pkg.remove(e.path)
    pkg.remove(e.path.replace(/([^/]+)$/, '_rels/$1.rels'))
    pkg.removeOverride(e.path)
  }

  for (const slide of model.slides) {
    const doc = parseXml(pkg.getText(slide.path)!)
    const tree = findDescendant(doc.root, 'p:spTree')
    if (tree) {
      const byId = new Map(shapeChildren(tree).map((n) => [shapeId(n).id, n]))
      const head = childElements(tree).filter((c) => c.name === 'p:nvGrpSpPr' || c.name === 'p:grpSpPr')
      const tail = childElements(tree).filter((c) => c.name === 'p:extLst')
      const shapes: XmlNode[] = []
      for (const shape of slide.shapes) {
        if ('xml' in shape && shape.xml) {
          shapes.push(...parseFragment(shape.xml))
          continue
        }
        const node = byId.get(shape.id)
        if (!node) throw new Error(`Shape #${shape.id} not found on ${slide.path}; pass \`xml\` to insert a new shape`)
        if (shape.type === 'text') writeTextShape(node, shape)
        else if (shape.type === 'table') writeTable(node, shape)
        shapes.push(node)
      }
      tree.children = [...head, ...shapes, ...tail]
    }
    pkg.setText(slide.path, serializeXml(doc))

    if (slide.notes !== undefined) {
      const path = notesPart(pkg, slide.path)
      if (path && notesText(pkg, slide.path) !== slide.notes) {
        const ndoc = parseXml(pkg.getText(path)!)
        const body = notesBody(ndoc.root)
        if (body) {
          body.children = [
            ...childElements(body).filter((c) => c.name !== 'a:p'),
            ...slide.notes.split('\n').map((line) => el('a:p', {}, line ? [el('a:r', {}, [el('a:t', {}, [{ type: 'text', text: line }])])] : [])),
          ]
          pkg.setText(path, serializeXml(ndoc))
        }
      }
    }
  }
}

function writeTextShape(node: XmlElement, shape: PptxTextShape) {
  let body = firstChild(node, 'p:txBody')
  if (!body) {
    body = el('p:txBody', {}, [el('a:bodyPr'), el('a:lstStyle')])
    node.children.push(body)
  }
  setParagraphs(body, shape.paragraphs)
}

function setParagraphs(body: XmlElement, paragraphs: PptxParagraph[]) {
  const keep = childElements(body).filter((c) => c.name === 'a:bodyPr' || c.name === 'a:lstStyle')
  const ps = paragraphs.length ? paragraphs.map(buildParagraph).join('') : '<a:p/>'
  body.children = [...keep, ...parseFragment(ps)]
}

function buildParagraph(p: PptxParagraph): string {
  const pPr = writeProps('a:pPr', p.props, p as unknown as Record<string, unknown>, PARA_FIELDS, [])
  let inner = ''
  for (const c of p.children ?? []) {
    if (c.type === 'raw') {
      inner += c.xml
      continue
    }
    const rPr = writeProps('a:rPr', c.props, c as unknown as Record<string, unknown>, RUN_FIELDS, RPR_ORDER)
    const lines = (c.text ?? '').split(/\r\n|\r|\n/)
    lines.forEach((line, i) => {
      if (i > 0) inner += rPr ? `<a:br>${rPr}</a:br>` : '<a:br/>'
      if (line) inner += `<a:r>${rPr}<a:t>${escapeXml(line)}</a:t></a:r>`
    })
  }
  return `<a:p>${pPr}${inner}${p.endProps ?? ''}</a:p>`
}

function writeTable(node: XmlElement, shape: PptxTableShape) {
  const tbl = findDescendant(node, 'a:tbl')
  if (!tbl) return
  const rows = childElements(tbl, 'a:tr')
  const cols = childElements(firstChild(tbl, 'a:tblGrid') ?? el('a:tblGrid'), 'a:gridCol').length
  if (!rows.length) return
  const newRows: XmlElement[] = shape.rows.map((r, i) => {
    if (cols && r.cells.length !== cols) throw new Error(`Table #${shape.id}: row ${i + 1} has ${r.cells.length} cells, the table has ${cols} columns`)
    const tr = rows[i] ?? (parseFragment(serializeNode(rows[rows.length - 1]))[0] as XmlElement)
    const tcs = childElements(tr, 'a:tc')
    r.cells.forEach((cell, j) => {
      const tc = tcs[j]
      if (!tc) return
      let body = firstChild(tc, 'a:txBody')
      if (!body) {
        body = el('a:txBody', {}, [el('a:bodyPr'), el('a:lstStyle')])
        tc.children.unshift(body)
      }
      setParagraphs(body, cell.paragraphs)
    })
    return tr
  })
  tbl.children = [...childElements(tbl).filter((c) => c.name !== 'a:tr'), ...newRows]
}

export function pptxText(model: PptxDocument): string {
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

