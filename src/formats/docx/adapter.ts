import { EMU_PER_PX } from '../../image'
import { OpcPackage, REL } from '../../package'
import { escapeAttr } from '../../xml/escape'
import type { FormatAdapter } from '../adapter'

const PRESERVE_T = '<w:t xml:space="preserve">'

export function docxMainPart(pkg: OpcPackage): string {
  return pkg.mainDocumentPath() ?? 'word/document.xml'
}

export const docxAdapter: FormatAdapter = {
  format: 'docx',
  spec: {
    textTags: new Set(['w:t']),
    groupTags: new Set(['w:p']),
    paragraphTag: 'w:p',
    rowTags: new Set(['w:tr']),
    contentTags: new Set(['w:drawing', 'w:pict', 'w:object', 'w:tab', 'w:br', 'w:sym', 'w:fldChar', 'w:fldSimple', 'm:oMath', 'm:oMathPara']),
    preserveSpace: true,
    imageMode: 'inline',
    breaks: () => ({ br: `</w:t><w:br/>${PRESERVE_T}`, tab: `</w:t><w:tab/>${PRESERVE_T}` }),
    renderImage(img, ctx) {
      const rId = ctx.pkg.addImage(ctx.part, 'word/media', img.data, img.extension)
      const id = (ctx.state.docPr = (ctx.state.docPr ?? 0x7000) + 1)
      const cx = Math.round(img.width * EMU_PER_PX)
      const cy = Math.round(img.height * EMU_PER_PX)
      const alt = escapeAttr(img.alt)
      const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
      const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
      return (
        `</w:t><w:drawing>` +
        `<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
        `<wp:docPr id="${id}" name="Picture ${id}" descr="${alt}"/>` +
        `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A}" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
        `<a:graphic xmlns:a="${A}"><a:graphicData uri="${PIC}"><pic:pic xmlns:pic="${PIC}">` +
        `<pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}.${img.extension}" descr="${alt}"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
        `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>${PRESERVE_T}`
      )
    },
  },
  templateParts(pkg) {
    const main = docxMainPart(pkg)
    const parts = [main]
    for (const type of [REL.header, REL.footer, REL.footnotes, REL.endnotes]) {
      for (const { path } of pkg.getRelatedParts(main, type)) if (pkg.has(path) && !parts.includes(path)) parts.push(path)
    }
    return parts
  },
}
