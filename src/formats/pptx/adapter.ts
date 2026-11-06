import { EMU_PER_PX } from '../../image'
import { REL, type OpcPackage } from '../../package'
import { escapeAttr } from '../../xml/escape'
import { childElements, parseXml } from '../../xml/dom'
import type { Tok } from '../../template/engine'
import type { FormatAdapter } from '../adapter'

const REL_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'

export function presentationPath(pkg: OpcPackage): string {
  return pkg.mainDocumentPath() ?? 'ppt/presentation.xml'
}

/** Slides in presentation order. */
export function slideParts(pkg: OpcPackage): Array<{ path: string; rId: string; id: string }> {
  const pres = presentationPath(pkg)
  const xml = pkg.getText(pres)
  if (!xml) return []
  const rels = new Map(pkg.getRelatedParts(pres, REL.slide).map((r) => [r.rel.id, r.path]))
  const list = parseXml(xml).root.children.find((c) => c.type === 'element' && c.name === 'p:sldIdLst')
  if (!list || list.type !== 'element') return []
  return childElements(list, 'p:sldId')
    .map((s) => ({ id: s.attrs.id, rId: s.attrs['r:id'], path: rels.get(s.attrs['r:id']) ?? '' }))
    .filter((s) => s.path && pkg.has(s.path))
}

function runProps(tokens: readonly Tok[], runIdx: number, closeOf: ReadonlyMap<number, number>): string {
  const next = tokens[runIdx + 1]
  if (!next || next.t !== 'xml' || next.name !== 'a:rPr') return ''
  if (next.kind === 'self') return next.raw
  const end = closeOf.get(runIdx + 1)
  if (end === undefined) return ''
  let xml = ''
  for (let i = runIdx + 1; i <= end; i++) {
    const t = tokens[i]
    xml += t.t === 'xml' ? t.raw : ''
  }
  return xml
}

const num = (xml: string, re: RegExp) => {
  const m = re.exec(xml)
  return m ? Number(m[1]) : undefined
}

interface Box {
  x: number
  y: number
  cx: number
  cy: number
}

function xfrmOf(xml: string): Box | undefined {
  const spPr = /<p:spPr\b[^>]*\/>|<p:spPr\b[^>]*>[\s\S]*?<\/p:spPr>/.exec(xml)?.[0] ?? ''
  const x = num(spPr, /<a:off\b[^>]*\sx="(-?\d+)"/)
  const y = num(spPr, /<a:off\b[^>]*\sy="(-?\d+)"/)
  const cx = num(spPr, /<a:ext\b[^>]*\scx="(\d+)"/)
  const cy = num(spPr, /<a:ext\b[^>]*\scy="(\d+)"/)
  return x !== undefined && y !== undefined && cx !== undefined && cy !== undefined ? { x, y, cx, cy } : undefined
}

/** Position of a placeholder shape inherited from the slide layout. */
function layoutBox(pkg: OpcPackage, slidePart: string, shapeXml: string): Box | undefined {
  const ph = /<p:ph\b[^>]*>/.exec(shapeXml)?.[0]
  if (!ph) return undefined
  const type = /\stype="([^"]*)"/.exec(ph)?.[1]
  const idx = /\sidx="([^"]*)"/.exec(ph)?.[1]
  const layout = pkg.getRelatedParts(slidePart, REL_LAYOUT)[0]?.path
  const layoutXml = layout && pkg.getText(layout)
  if (!layoutXml) return undefined
  for (const m of layoutXml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)) {
    const lph = /<p:ph\b[^>]*>/.exec(m[0])?.[0]
    if (!lph) continue
    const ltype = /\stype="([^"]*)"/.exec(lph)?.[1]
    const lidx = /\sidx="([^"]*)"/.exec(lph)?.[1]
    if ((idx !== undefined && lidx === idx) || (idx === undefined && type !== undefined && ltype === type)) return xfrmOf(m[0])
  }
  return undefined
}

export const pptxAdapter: FormatAdapter = {
  format: 'pptx',
  spec: {
    textTags: new Set(['a:t']),
    groupTags: new Set(['a:p']),
    paragraphTag: 'a:p',
    rowTags: new Set(['a:tr']),
    contentTags: new Set(['a:br', 'a:fld', 'a14:m', 'mc:AlternateContent']),
    imageMode: 'shape',
    shapeTag: 'p:sp',
    breaks({ tokens, stack, closeOf }) {
      const run = stack[stack.length - 2]
      const tok = run !== undefined ? tokens[run] : undefined
      if (!tok || tok.t !== 'xml' || tok.name !== 'a:r') return {}
      const rPr = runProps(tokens, run, closeOf)
      const br = rPr ? `<a:br>${rPr}</a:br>` : '<a:br/>'
      return { br: `</a:t></a:r>${br}<a:r>${rPr}<a:t>` }
    },
    renderImage(img, ctx) {
      const shape = ctx.shapeXml ?? ''
      const id = /<p:cNvPr\b[^>]*\sid="(\d+)"/.exec(shape)?.[1] ?? String((ctx.state.shapeId = (ctx.state.shapeId ?? 0x7000) + 1))
      const name = /<p:cNvPr\b[^>]*\sname="([^"]*)"/.exec(shape)?.[1] ?? `Picture ${id}`
      const box = xfrmOf(shape) ?? layoutBox(ctx.pkg, ctx.part, shape)
      let cx = Math.round(img.width * EMU_PER_PX)
      let cy = Math.round(img.height * EMU_PER_PX)
      let x = 0
      let y = 0
      if (box) {
        // fit ("contain") the picture into the text box, centered
        const scale = Math.min(box.cx / cx, box.cy / cy)
        cx = Math.round(cx * scale)
        cy = Math.round(cy * scale)
        x = box.x + Math.round((box.cx - cx) / 2)
        y = box.y + Math.round((box.cy - cy) / 2)
      }
      const rId = ctx.pkg.addImage(ctx.part, 'ppt/media', img.data, img.extension)
      return (
        `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}" descr="${escapeAttr(img.alt)}"/>` +
        `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
      )
    },
  },
  templateParts(pkg) {
    const parts: string[] = []
    for (const s of slideParts(pkg)) {
      parts.push(s.path)
      for (const { path } of pkg.getRelatedParts(s.path, REL.notesSlide)) if (pkg.has(path)) parts.push(path)
    }
    return parts
  },
}

