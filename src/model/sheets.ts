import type { XlsxCell, XlsxSheet, XlsxValue } from './types'

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

const DAY_MS = 86400000
const EPOCH_1900 = Date.UTC(1899, 11, 30)
const EPOCH_1904 = Date.UTC(1904, 0, 1)

export function serialToDate(serial: number, date1904 = false): Date {
  return new Date((date1904 ? EPOCH_1904 : EPOCH_1900) + Math.round(serial * DAY_MS))
}

export function dateToSerial(date: Date, date1904 = false): number {
  return (date.getTime() - (date1904 ? EPOCH_1904 : EPOCH_1900)) / DAY_MS
}

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
