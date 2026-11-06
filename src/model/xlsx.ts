import { readSharedStrings, sharedStringsPath, workbookPath, worksheetParts, type SharedString } from '../formats/xlsx/adapter'
import { REL, type OpcPackage } from '../package'
import { childElements, el, firstChild, parseFragment, parseXml, serializeChildren, serializeXml, textContent, type XmlElement } from '../xml/dom'
import { escapeXml } from '../xml/escape'

// ---------------------------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------------------------

export interface XlsxDocument {
  format: 'xlsx'
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
// References
// ---------------------------------------------------------------------------------------------

export function columnToIndex(col: string): number {
  let n = 0
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}

export function indexToColumn(n: number): string {
  let s = ''
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

export function parseRef(ref: string): { col: number; row: number } {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(ref)
  if (!m) throw new Error(`Invalid cell reference "${ref}"`)
  return { col: columnToIndex(m[1]), row: Number(m[2]) }
}

// ---------------------------------------------------------------------------------------------
// Dates
// ---------------------------------------------------------------------------------------------

const DAY_MS = 86400000
const EPOCH_1900 = Date.UTC(1899, 11, 30)
const EPOCH_1904 = Date.UTC(1904, 0, 1)

export function serialToDate(serial: number, date1904 = false): Date {
  return new Date((date1904 ? EPOCH_1904 : EPOCH_1900) + Math.round(serial * DAY_MS))
}

export function dateToSerial(date: Date, date1904 = false): number {
  return (date.getTime() - (date1904 ? EPOCH_1904 : EPOCH_1900)) / DAY_MS
}

function toIsoDate(d: Date): string {
  const iso = d.toISOString()
  return iso.endsWith('T00:00:00.000Z') ? iso.slice(0, 10) : iso.replace('.000Z', 'Z')
}

const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58])

function isDateFormat(code: string): boolean {
  const cleaned = code.replace(/"[^"]*"|\\.|\[[^\]]*]|_.|\*./g, '')
  return /[dmyhs]/i.test(cleaned) && !/^general$/i.test(cleaned.trim())
}

interface StyleInfo {
  dateStyles: Set<number>
  date1904: boolean
}

function readStyles(pkg: OpcPackage, wb: string): StyleInfo {
  const date1904 = /<workbookPr\b[^>]*\sdate1904="(1|true)"/.test(pkg.getText(wb) ?? '')
  const dateStyles = new Set<number>()
  const path = pkg.getRelatedParts(wb, REL.styles)[0]?.path
  const xml = path && pkg.getText(path)
  if (!xml) return { dateStyles, date1904 }
  const root = parseXml(xml).root
  const custom = new Map<number, string>()
  const numFmts = firstChild(root, 'numFmts')
  if (numFmts) for (const f of childElements(numFmts, 'numFmt')) custom.set(Number(f.attrs.numFmtId), f.attrs.formatCode ?? '')
  const xfs = firstChild(root, 'cellXfs')
  if (xfs) {
    childElements(xfs, 'xf').forEach((xf, i) => {
      const id = Number(xf.attrs.numFmtId ?? 0)
      if (BUILTIN_DATE_FORMATS.has(id) || (custom.has(id) && isDateFormat(custom.get(id)!))) dateStyles.add(i)
    })
  }
  return { dateStyles, date1904 }
}

// ---------------------------------------------------------------------------------------------
// Parse
// ---------------------------------------------------------------------------------------------

export function parseXlsx(pkg: OpcPackage): XlsxDocument {
  const wb = workbookPath(pkg)
  const shared = readSharedStrings(pkg, wb)
  const styles = readStyles(pkg, wb)
  const sheets = worksheetParts(pkg).map((s) => parseSheet(pkg, s.name, s.path, shared, styles))
  return { format: 'xlsx', sheets }
}

function parseSheet(pkg: OpcPackage, name: string, path: string, shared: SharedString[], styles: StyleInfo): XlsxSheet {
  const root = parseXml(pkg.getText(path)!).root
  const sheet: XlsxSheet = { name, path, rows: [] }
  const data = firstChild(root, 'sheetData')
  const sharedFormulas = new Map<string, { formula: string; col: number; row: number }>()

  let prevRow = 0
  for (const r of data ? childElements(data, 'row') : []) {
    const index = r.attrs.r ? Number(r.attrs.r) : prevRow + 1
    prevRow = index
    const row: XlsxRow = { index, cells: [] }
    if (r.attrs.customHeight === '1' && r.attrs.ht) row.height = Number(r.attrs.ht)
    if (r.attrs.hidden === '1' || r.attrs.hidden === 'true') row.hidden = true
    let prevCol = 0
    for (const c of childElements(r, 'c')) {
      const ref = c.attrs.r ?? `${indexToColumn(prevCol + 1)}${index}`
      prevCol = parseRef(ref).col
      const cell = parseCell(c, ref, shared, styles, sharedFormulas)
      if (cell) row.cells.push(cell)
    }
    if (row.cells.length || row.height !== undefined || row.hidden) sheet.rows.push(row)
  }
  const merges = firstChild(root, 'mergeCells')
  if (merges) sheet.merges = childElements(merges, 'mergeCell').map((m) => m.attrs.ref)
  return sheet
}

function parseCell(
  c: XmlElement,
  ref: string,
  shared: SharedString[],
  styles: StyleInfo,
  sharedFormulas: Map<string, { formula: string; col: number; row: number }>,
): XlsxCell | undefined {
  const t = c.attrs.t ?? 'n'
  const v = firstChild(c, 'v')
  const vText = v ? textContent(v) : undefined
  const f = firstChild(c, 'f')
  const style = c.attrs.s !== undefined ? Number(c.attrs.s) : undefined
  const cell: XlsxCell = { ref, value: null }
  if (style) cell.style = style

  if (f) {
    const text = textContent(f)
    if (f.attrs.t === 'shared' && f.attrs.si !== undefined) {
      const pos = parseRef(ref)
      if (text) sharedFormulas.set(f.attrs.si, { formula: text, ...pos })
      const master = sharedFormulas.get(f.attrs.si)
      if (master) cell.formula = text || shiftFormula(master.formula, pos.row - master.row, pos.col - master.col)
    } else if (text) cell.formula = text
  }

  switch (t) {
    case 's': {
      const si = vText !== undefined ? shared[Number(vText)] : undefined
      cell.value = si?.text ?? ''
      if (si && /<r[\s>]/.test(si.xml)) cell.richText = si.xml
      break
    }
    case 'inlineStr': {
      const is = firstChild(c, 'is')
      cell.value = is ? childElements(is).filter((e) => e.name !== 'rPh').map(textContent).join('') : ''
      if (is && firstChild(is, 'r')) cell.richText = serializeChildren(is)
      break
    }
    case 'str':
      cell.value = vText ?? ''
      break
    case 'b':
      cell.value = vText === '1' || vText === 'true'
      break
    case 'e':
      cell.value = vText ?? '#N/A'
      cell.type = 'error'
      break
    case 'd':
      cell.value = vText ?? null
      cell.type = 'date'
      break
    default:
      if (vText === undefined || vText === '') cell.value = null
      else if (style !== undefined && styles.dateStyles.has(style)) {
        cell.value = toIsoDate(serialToDate(Number(vText), styles.date1904))
        cell.type = 'date'
      } else cell.value = Number(vText)
  }
  if (cell.value === null && !cell.formula && cell.style === undefined) return undefined
  return cell
}

/** Shifts relative references of a formula (used to expand shared formulas). */
export function shiftFormula(formula: string, dRow: number, dCol: number): string {
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((seg, i) =>
      i % 2 === 1
        ? seg
        : seg.replace(/(^|[^A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/g, (_w, pre: string, dc: string, col: string, dr: string, row: string) => {
            const c = dc ? col : indexToColumn(columnToIndex(col) + dCol)
            const r = dr ? row : String(Number(row) + dRow)
            return `${pre}${dc}${c}${dr}${r}`
          }),
    )
    .join('')
}

// ---------------------------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------------------------

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const CT_SHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
const CT_SST = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
const MERGE_BEFORE = [
  'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
  'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing',
  'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst',
]

class StringTable {
  readonly items: string[] = []
  private index = new Map<string, number>()
  add(xml: string): number {
    let i = this.index.get(xml)
    if (i === undefined) {
      i = this.items.length
      this.items.push(xml)
      this.index.set(xml, i)
    }
    return i
  }
}

export function buildXlsx(model: XlsxDocument, pkg: OpcPackage): void {
  const wb = workbookPath(pkg)
  const wbDoc = parseXml(pkg.getText(wb)!)
  const styles = readStyles(pkg, wb)
  const sst = new StringTable()
  const existing = worksheetParts(pkg)
  const keep = new Set(model.sheets.map((s) => s.path).filter(Boolean))
  let dateStyle: number | undefined
  const getDateStyle = () => (dateStyle ??= ensureDateStyle(pkg, wb))

  // remove sheets that are not in the model
  for (const s of existing) {
    if (keep.has(s.path)) continue
    pkg.remove(s.path)
    pkg.removeOverride(s.path)
    pkg.removeRelationship(wb, (r) => r.id === s.rId)
  }

  const sheetsEl = firstChild(wbDoc.root, 'sheets') ?? el('sheets')
  const oldSheetEls = new Map(childElements(sheetsEl, 'sheet').map((s) => [s.attrs['r:id'], s]))
  let maxId = Math.max(0, ...childElements(sheetsEl, 'sheet').map((s) => Number(s.attrs.sheetId) || 0))
  const newSheetEls: XmlElement[] = []

  for (const sheet of model.sheets) {
    let path = sheet.path
    let rId: string
    const found = path ? existing.find((s) => s.path === path) : undefined
    if (found) rId = found.rId
    else {
      let n = 1
      while (pkg.has(`xl/worksheets/sheet${n}.xml`)) n++
      path = `xl/worksheets/sheet${n}.xml`
      pkg.setText(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>')
      pkg.ensureOverride(path, CT_SHEET)
      rId = pkg.addRelationship(wb, REL.worksheet, path.replace(/^xl\//, ''))
    }
    const sheetEl = oldSheetEls.get(rId) ?? el('sheet', { name: sheet.name, sheetId: String(++maxId), 'r:id': rId })
    sheetEl.attrs.name = sheet.name
    newSheetEls.push(sheetEl)
    writeSheet(pkg, path!, sheet, sst, styles, getDateStyle)
  }
  sheetsEl.children = newSheetEls
  wbDoc.root.attrs['xmlns:r'] ??= R_NS
  pkg.setText(wb, serializeXml(wbDoc))

  // shared strings
  let sstPath = sharedStringsPath(pkg, wb)
  if (!sstPath) {
    sstPath = 'xl/sharedStrings.xml'
    pkg.ensureOverride(sstPath, CT_SST)
    pkg.addRelationship(wb, REL.sharedStrings, 'sharedStrings.xml')
  }
  pkg.setText(
    sstPath,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sst.items.length}" uniqueCount="${sst.items.length}">` +
      sst.items.map((x) => `<si>${x}</si>`).join('') +
      '</sst>',
  )
}

function writeSheet(
  pkg: OpcPackage,
  path: string,
  sheet: XlsxSheet,
  sst: StringTable,
  styles: StyleInfo,
  getDateStyle: () => number,
) {
  const doc = parseXml(pkg.getText(path)!)
  const root = doc.root
  let rowsXml = ''
  const rows = [...sheet.rows].sort((a, b) => a.index - b.index)
  for (const row of rows) {
    let attrs = ` r="${row.index}"`
    if (row.height !== undefined) attrs += ` ht="${row.height}" customHeight="1"`
    if (row.hidden) attrs += ' hidden="1"'
    let cells = ''
    let col = 0
    for (const cell of row.cells) {
      const ref = cell.ref ?? `${indexToColumn(col + 1)}${row.index}`
      col = parseRef(ref).col
      cells += writeCell(cell, ref.replace(/\d+$/, String(row.index)), sst, styles, getDateStyle)
    }
    rowsXml += `<row${attrs}>${cells}</row>`
  }

  const sheetData = firstChild(root, 'sheetData')
  const dataNodes = parseFragment(`<sheetData>${rowsXml}</sheetData>`)[0] as XmlElement
  if (sheetData) sheetData.children = dataNodes.children
  else root.children.push(dataNodes)

  // merged cells
  root.children = root.children.filter((c) => !(c.type === 'element' && c.name === 'mergeCells'))
  if (sheet.merges?.length) {
    const merge = el('mergeCells', { count: String(sheet.merges.length) }, sheet.merges.map((ref) => el('mergeCell', { ref })))
    const at = root.children.findIndex((c) => c.type === 'element' && MERGE_BEFORE.includes(c.name))
    if (at === -1) root.children.push(merge)
    else root.children.splice(at, 0, merge)
  }
  root.children = root.children.filter((c) => !(c.type === 'element' && c.name === 'dimension'))
  pkg.setText(path, serializeXml(doc))
}

function writeCell(cell: XlsxCell, ref: string, sst: StringTable, styles: StyleInfo, getDateStyle: () => number): string {
  let style = cell.style
  let value = cell.value
  let t = ''
  let v: string | undefined

  if (value instanceof Date || (cell.type === 'date' && typeof value === 'string')) {
    const d = value instanceof Date ? value : new Date(value as string)
    if (!isNaN(d.getTime())) {
      value = dateToSerial(d, styles.date1904)
      if (style === undefined || !styles.dateStyles.has(style)) style = getDateStyle()
    }
  }

  if (value === null || value === undefined) {
    v = undefined
  } else if (cell.type === 'error') {
    t = 'e'
    v = String(value)
  } else if (typeof value === 'number') {
    v = Number.isFinite(value) ? String(value) : undefined
  } else if (typeof value === 'boolean') {
    t = 'b'
    v = value ? '1' : '0'
  } else if (cell.formula) {
    t = 'str'
    v = escapeXml(String(value))
  } else {
    const text = String(value)
    const rich = cell.richText && richPlainText(cell.richText) === text ? cell.richText : undefined
    const xml = rich ?? `<t xml:space="preserve">${escapeXml(text)}</t>`
    t = 's'
    v = String(sst.add(xml))
  }

  const s = style ? ` s="${style}"` : ''
  const tAttr = t ? ` t="${t}"` : ''
  const f = cell.formula ? `<f>${escapeXml(cell.formula.replace(/^=/, ''))}</f>` : ''
  if (!f && v === undefined) return style ? `<c r="${ref}"${s}/>` : ''
  return `<c r="${ref}"${s}${tAttr}>${f}${v !== undefined ? `<v>${v}</v>` : ''}</c>`
}

function richPlainText(xml: string): string {
  return parseFragment(xml)
    .filter((n) => !(n.type === 'element' && n.name === 'rPh'))
    .map(textContent)
    .join('')
}

/** Finds or adds a cell format with the built-in date format (numFmtId 14). */
function ensureDateStyle(pkg: OpcPackage, wb: string): number {
  const path = pkg.getRelatedParts(wb, REL.styles)[0]?.path
  const xml = path && pkg.getText(path)
  if (!path || !xml) return 0
  const doc = parseXml(xml)
  let xfs = firstChild(doc.root, 'cellXfs')
  if (!xfs) {
    xfs = el('cellXfs', { count: '1' }, [el('xf', { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0', xfId: '0' })])
    doc.root.children.push(xfs)
  }
  const list = childElements(xfs, 'xf')
  const found = list.findIndex((x) => x.attrs.numFmtId === '14' && !x.attrs.fontId?.match(/[1-9]/) && !x.attrs.fillId?.match(/[1-9]/))
  if (found !== -1) return found
  xfs.children.push(el('xf', { numFmtId: '14', fontId: '0', fillId: '0', borderId: '0', xfId: '0', applyNumberFormat: '1' }))
  xfs.attrs.count = String(list.length + 1)
  pkg.setText(path, serializeXml(doc))
  return list.length
}

// ---------------------------------------------------------------------------------------------
// Convenience helpers
// ---------------------------------------------------------------------------------------------

/** Sheet → 2D array of values (`null` for empty cells). */
export function sheetToArrays(sheet: XlsxSheet): Array<Array<XlsxValue | Date>> {
  const out: Array<Array<XlsxValue | Date>> = []
  const maxRow = Math.max(0, ...sheet.rows.map((r) => r.index))
  let maxCol = 0
  for (const r of sheet.rows) for (const c of r.cells) if (c.ref) maxCol = Math.max(maxCol, parseRef(c.ref).col)
  for (let i = 0; i < maxRow; i++) out.push(new Array(maxCol).fill(null))
  for (const r of sheet.rows) {
    let col = 0
    for (const c of r.cells) {
      col = c.ref ? parseRef(c.ref).col : col + 1
      out[r.index - 1][col - 1] = c.value
    }
  }
  return out
}

/**
 * Sheet → array of objects using a header row (1-based, default 1). Empty rows are skipped.
 * @example sheetToObjects(sheet) // [{ Name: 'Tea', Qty: 2 }, ...]
 */
export function sheetToObjects(sheet: XlsxSheet, headerRow = 1): Array<Record<string, XlsxValue | Date>> {
  const rows = sheetToArrays(sheet)
  const header = (rows[headerRow - 1] ?? []).map((h, i) => (h === null || h === '' ? `column${i + 1}` : String(h)))
  return rows
    .slice(headerRow)
    .filter((r) => r.some((v) => v !== null && v !== ''))
    .map((r) => Object.fromEntries(header.map((h, i) => [h, r[i] ?? null])))
}

/** 2D array → sheet model. Dates become date cells. */
export function sheetFromArrays(name: string, rows: Array<Array<XlsxValue | Date | undefined>>): XlsxSheet {
  return {
    name,
    rows: rows.map((r, i) => ({
      index: i + 1,
      cells: r
        .map((value, j) => ({ ref: `${indexToColumn(j + 1)}${i + 1}`, value: value ?? null }) as XlsxCell)
        .filter((c) => c.value !== null),
    })),
  }
}

export function xlsxText(model: XlsxDocument): string {
  return model.sheets
    .map((s) => `${s.name}\n` + sheetToArrays(s).map((r) => r.map((v) => (v === null ? '' : String(v))).join('\t')).join('\n'))
    .join('\n\n')
}

