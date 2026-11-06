import { REL, type OpcPackage } from '../../package'
import { decodeXml, escapeXml } from '../../xml/escape'
import { childElements, parseXml, serializeChildren, textContent } from '../../xml/dom'
import { removeRawAttribute, setRawAttribute } from '../../xml/tokenizer'
import type { FormatAdapter } from '../adapter'

const GAP_ATTR = 'data-ct-gap'
const MARK = '\u0000'
const MARK_SEP = '\u0001'

export interface SharedString {
  /** Inner XML of `<si>`. */
  xml: string
  text: string
}

export function sharedStringsPath(pkg: OpcPackage, workbook: string): string | undefined {
  return pkg.getRelatedParts(workbook, REL.sharedStrings)[0]?.path
}

export function readSharedStrings(pkg: OpcPackage, workbook: string): SharedString[] {
  const path = sharedStringsPath(pkg, workbook)
  const xml = path && pkg.getText(path)
  if (!xml) return []
  return childElements(parseXml(xml).root, 'si').map((si) => ({
    xml: serializeChildren(si),
    // phonetic runs (rPh) are not part of the visible text
    text: si.children.filter((c) => !(c.type === 'element' && c.name === 'rPh')).map(textContent).join(''),
  }))
}

export function workbookPath(pkg: OpcPackage): string {
  return pkg.mainDocumentPath() ?? 'xl/workbook.xml'
}

export function worksheetParts(pkg: OpcPackage): Array<{ name: string; path: string; rId: string }> {
  const wb = workbookPath(pkg)
  const xml = pkg.getText(wb)
  if (!xml) return []
  const rels = new Map(pkg.getRelatedParts(wb, REL.worksheet).map((r) => [r.rel.id, r.path]))
  const sheets = parseXml(xml).root.children.find((c) => c.type === 'element' && c.name === 'sheets')
  if (!sheets || sheets.type !== 'element') return []
  return childElements(sheets, 'sheet')
    .map((s) => ({ name: s.attrs.name, rId: s.attrs['r:id'], path: rels.get(s.attrs['r:id']) ?? '' }))
    .filter((s) => s.path && pkg.has(s.path))
}

const CELL_RE = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g
const ROW_RE = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g
const attr = (attrs: string, name: string) => new RegExp(`\\s${name}="([^"]*)"`).exec(attrs)?.[1]

interface SheetState {
  templateRows: Set<number>
}

export const xlsxAdapter: FormatAdapter = {
  format: 'xlsx',
  spec: {
    textTags: new Set(['t']),
    groupTags: new Set(['c']),
    rowTags: new Set(['row']),
    contentTags: new Set(),
    preserveSpace: true,
    imageMode: 'none',
    renderScalar(value, text) {
      if ((typeof value === 'number' && Number.isFinite(value)) || typeof value === 'boolean') {
        const kind = typeof value === 'number' ? 'N' : 'B'
        const v = typeof value === 'number' ? String(value) : value ? '1' : '0'
        return `${MARK}${kind}${v}${MARK_SEP}${escapeXml(text)}${MARK}`
      }
      return undefined
    },
  },

  templateParts(pkg) {
    return worksheetParts(pkg).map((s) => s.path)
  },

  preprocess(pkg, part, xml, state) {
    const start = (state.get('delimiterStart') as string) ?? '{{'
    let shared = state.get('sharedStrings') as SharedString[] | undefined
    if (!shared) {
      shared = readSharedStrings(pkg, workbookPath(pkg))
      state.set('sharedStrings', shared)
    }
    // 1. Inline shared strings that contain tags, so each cell owns its text.
    xml = xml.replace(CELL_RE, (whole, attrs: string, inner?: string) => {
      if (attr(attrs, 't') !== 's' || !inner) return whole
      const idx = /<v>\s*(\d+)\s*<\/v>/.exec(inner)?.[1]
      const si = idx !== undefined ? shared![Number(idx)] : undefined
      if (!si || !si.text.includes(start)) return whole
      const head = setRawAttribute(`<c${attrs}>`, 't', 'inlineStr')
      return `${head}${inner.replace(/<v>[\s\S]*?<\/v>/, '')}<is>${si.xml}</is></c>`
    })
    // 2. Remember the gap between consecutive rows so we can renumber after loops.
    const templateRows = new Set<number>()
    let prev = 0
    xml = xml.replace(/<row\b[^>]*>/g, (tag) => {
      const r = Number(attr(tag, 'r') ?? prev + 1)
      templateRows.add(r)
      const gap = r - prev
      prev = r
      return setRawAttribute(tag, GAP_ATTR, String(gap))
    })
    state.set(`sheet:${part}`, { templateRows } satisfies SheetState)
    return xml
  },

  postprocess(_pkg, part, xml, state) {
    xml = resolveScalarMarkers(xml)
    const sheet = state.get(`sheet:${part}`) as SheetState | undefined
    return renumberRows(xml, sheet?.templateRows ?? new Set())
  },

  finalize(pkg) {
    // Cached calculation order is invalid after rows move: drop it and ask Excel to recalculate.
    const wb = workbookPath(pkg)
    for (const { path } of pkg.getRelatedParts(wb, REL.calcChain)) {
      pkg.remove(path)
      pkg.removeOverride(path)
    }
    pkg.removeRelationship(wb, (r) => r.type === REL.calcChain)
    const xml = pkg.getText(wb)
    if (!xml) return
    let next: string
    if (/<calcPr\b/.test(xml)) next = xml.replace(/<calcPr\b[^>]*?\/?>/, (t) => setRawAttribute(t, 'fullCalcOnLoad', '1'))
    else if (xml.includes('</workbook>')) {
      // calcPr must precede these elements
      const m = /<(oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b/.exec(xml)
      next = m ? xml.slice(0, m.index) + '<calcPr fullCalcOnLoad="1"/>' + xml.slice(m.index) : xml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
    } else next = xml
    pkg.setText(wb, next)
  },
}

/** Turns cells that consist of a single numeric/boolean value into real typed cells. */
function resolveScalarMarkers(xml: string): string {
  if (!xml.includes(MARK)) return xml
  const cellRe = new RegExp(
    `<c\\b([^>]*?)>(\\s*<f\\b[^>]*>[^<]*<\\/f>)?\\s*<is>\\s*<t\\b[^>]*>${MARK}([NB])([^${MARK_SEP}]*)${MARK_SEP}[^${MARK}]*${MARK}<\\/t>\\s*<\\/is>\\s*<\\/c>`,
    'g',
  )
  xml = xml.replace(cellRe, (_w, attrs: string, _f, kind: string, v: string) => {
    let head = removeRawAttribute(`<c${attrs}>`, 't')
    if (kind === 'B') head = setRawAttribute(head, 't', 'b')
    return `${head}<v>${v}</v></c>`
  })
  // Values mixed with other text stay strings.
  return xml.replace(new RegExp(`${MARK}[NB][^${MARK_SEP}]*${MARK_SEP}([^${MARK}]*)${MARK}`, 'g'), '$1')
}

// ------------------------------------------------------------------------------------------
// Row renumbering
// ------------------------------------------------------------------------------------------

const REF_RE = /(\$?)([A-Z]{1,3})(\$?)(\d+)/g

function renumberRows(xml: string, templateRows: Set<number>): string {
  const sdStart = xml.indexOf('<sheetData')
  if (sdStart === -1) return xml
  const sdEnd = xml.indexOf('</sheetData>')
  if (sdEnd === -1) return xml.replace(new RegExp(`\\s${GAP_ATTR}="\\d+"`, 'g'), '')

  // pass 1: assign new row numbers
  const rows: Array<{ orig: number; next: number }> = []
  let cur = 0
  let prevOrig = 0
  const body = xml.slice(sdStart, sdEnd)
  for (const m of body.matchAll(ROW_RE)) {
    const gap = Number(attr(m[1], GAP_ATTR) ?? 1)
    const orig = Number(attr(m[1], 'r') ?? prevOrig + 1)
    cur += gap
    rows.push({ orig, next: cur })
    prevOrig = orig
  }
  const map = new Map<number, number[]>()
  for (const r of rows) {
    const list = map.get(r.orig)
    if (list) list.push(r.next)
    else map.set(r.orig, [r.next])
  }
  const changed = rows.some((r) => r.orig !== r.next) || [...templateRows].some((r) => !map.has(r)) || rows.length !== map.size
  if (!changed) return xml.replace(new RegExp(`\\s${GAP_ATTR}="\\d+"`, 'g'), '')

  const sortedOrig = [...map.keys()].sort((a, b) => a - b)
  /** Maps an original row reference to its new number. `end` = last copy (for range ends). */
  const mapRow = (r: number, end: boolean, selfOrig?: number, selfNew?: number): number | undefined => {
    if (selfOrig !== undefined && r === selfOrig) return selfNew
    const list = map.get(r)
    if (list) return end ? list[list.length - 1] : list[0]
    if (templateRows.has(r)) return undefined // row was removed by an empty loop
    // Row not in the template (empty row): shift by the delta of the closest row above.
    let base: number | undefined
    for (const o of sortedOrig) {
      if (o < r) base = o
      else break
    }
    if (base === undefined) return r
    const l = map.get(base)!
    return r + (l[l.length - 1] - base)
  }
  const mapRef = (ref: string, end: boolean, selfOrig?: number, selfNew?: number) =>
    ref.replace(REF_RE, (w, d1: string, col: string, d2: string, row: string) => {
      const n = mapRow(Number(row), end, selfOrig, selfNew)
      return n === undefined ? w : `${d1}${col}${d2}${n}`
    })
  const mapRange = (range: string, selfOrig?: number, selfNew?: number) => {
    const [a, b] = range.split(':')
    return b === undefined ? mapRef(a, false, selfOrig, selfNew) : `${mapRef(a, false, selfOrig, selfNew)}:${mapRef(b, true, selfOrig, selfNew)}`
  }
  const mapFormula = (f: string, selfOrig: number, selfNew: number) =>
    f
      // keep string literals and references to other sheets ('My sheet'!A1) untouched
      .split(/("(?:[^"]|"")*"|'(?:[^']|'')*'!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/)
      .map((seg, i) => {
        if (i % 2 === 1) return seg
        // ranges first, then single refs; skip refs to other sheets (Sheet2!A1)
        return seg.replace(/(^|[^A-Za-z0-9_!.])(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?(?![A-Za-z0-9_(])/g, (w, pre: string, a: string, b?: string) =>
          pre + (b ? `${mapRef(a, false, selfOrig, selfNew)}:${mapRef(b, true, selfOrig, selfNew)}` : mapRef(a, false, selfOrig, selfNew)),
        )
      })
      .join('')

  // pass 2: rewrite rows, cells and formulas
  let n = 0
  const newBody = body.replace(ROW_RE, (_w, attrs: string, close: string, inner?: string) => {
    const { orig, next } = rows[n++]
    let head = removeRawAttribute(`<row${attrs}${close === '/>' ? '/>' : '>'}`, GAP_ATTR)
    head = setRawAttribute(head, 'r', String(next))
    if (close === '/>') return head
    const content = (inner ?? '')
      .replace(/(<c\b[^>]*?\sr=")([A-Z]+)\d+"/g, `$1$2${next}"`)
      .replace(/(<f\b[^>]*>)([\s\S]*?)(<\/f>)/g, (_m, o: string, f: string, c: string) => {
        const openTag = /\sref="([^"]*)"/.test(o) ? o.replace(/\sref="([^"]*)"/, (_x, r: string) => ` ref="${mapRange(r, orig, next)}"`) : o
        return openTag + escapeXml(mapFormula(decodeXml(f), orig, next)) + c
      })
    return `${head}${content}</row>`
  })

  let out = xml.slice(0, sdStart) + newBody + xml.slice(sdEnd)

  // merged cells: duplicate merges inside repeated rows, shift the others
  out = out.replace(/<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/, (_w, inner: string) => {
    const merges: string[] = []
    for (const m of inner.matchAll(/<mergeCell\b[^>]*\sref="([^"]*)"[^>]*\/>/g)) {
      const [a, b] = m[1].split(':')
      const ra = Number(/\d+/.exec(a)?.[0])
      const rb = Number(/\d+/.exec(b ?? a)?.[0])
      const ca = a.replace(/\$?(\d+)/, '').replace('$', '')
      const cb = (b ?? a).replace(/\$?(\d+)/, '').replace('$', '')
      if (ra === rb && map.has(ra)) {
        for (const r of map.get(ra)!) merges.push(`${ca}${r}:${cb}${r}`)
        continue
      }
      const na = mapRow(ra, false)
      const nb = mapRow(rb, true)
      if (na === undefined || nb === undefined) continue
      merges.push(`${ca}${na}:${cb}${nb}`)
    }
    if (!merges.length) return ''
    return `<mergeCells count="${merges.length}">${merges.map((r) => `<mergeCell ref="${r}"/>`).join('')}</mergeCells>`
  })
  // ranges in conditional formatting, data validation and hyperlinks
  out = out.replace(/(<(?:conditionalFormatting|dataValidation)\b[^>]*\ssqref=")([^"]*)"/g, (_w, p: string, list: string) => `${p}${list.split(/\s+/).map((r) => mapRange(r)).join(' ')}"`)
  out = out.replace(/(<hyperlink\b[^>]*\sref=")([^"]*)"/g, (_w, p: string, r: string) => `${p}${mapRange(r)}"`)
  // the used-range hint is optional; drop it rather than keep a wrong one
  out = out.replace(/<dimension\b[^>]*\/>/, '')
  return out
}
