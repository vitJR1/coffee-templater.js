import { describe, expect, it } from 'vitest'
import { createTemplater, parse, type PptxDocument, type XlsxDocument } from '../src'
import { fixture, partNames, readPart } from './helpers'

const t = createTemplater()
const items = [
  { name: 'Coffee', qty: 2, price: 10.5 },
  { name: 'Tea', qty: 1, price: 4 },
  { name: 'Milk', qty: 3, price: 1.25 },
]

const cells = (doc: XlsxDocument, sheet = 0) =>
  Object.fromEntries(doc.sheets[sheet].rows.flatMap((r) => r.cells.map((c) => [c.ref, c.formula ? `=${c.formula}` : c.value])))

describe('xlsx templates', () => {
  it('repeats rows, keeps types and shifts rows below', () => {
    const out = t.render(fixture('report.xlsx'), { title: 'Q1', items, author: 'me' })
    const c = cells(parse(out) as XlsxDocument)
    expect(c.A1).toBe('Report: Q1')
    expect([c.A3, c.B3, c.C3]).toEqual(['Coffee', 2, 10.5])
    expect([c.A5, c.B5, c.C5]).toEqual(['Milk', 3, 1.25])
    expect(c.D4).toBe('=B4*C4')
    expect(c.A7).toBe('Total')
    expect(c.D7).toBe('=SUM(D3:D5)')
    expect(c.A8).toBe('Escaped {{x}}')
  })

  it('removes loop rows for empty arrays and shifts rows up', () => {
    const out = t.render(fixture('report.xlsx'), { title: 'Q1', items: [], author: 'me' })
    const c = cells(parse(out) as XlsxDocument)
    expect(c.A2).toBe('Name')
    expect(c.A3).toBeUndefined()
    expect(c.A4).toBe('Total')
  })

  it('keeps merged cells and renders other sheets', () => {
    const out = t.render(fixture('report.xlsx'), { title: 'Q1', items, author: 'me' })
    const doc = parse(out) as XlsxDocument
    expect(doc.sheets[0].merges).toEqual(['A1:D1'])
    expect(cells(doc, 1).B1).toBe('me')
  })

  it('drops calcChain and forces recalculation', () => {
    const out = t.render(fixture('report.xlsx'), { title: 'Q1', items, author: 'me' })
    expect(readPart(out, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"')
    expect(partNames(out)).not.toContain('xl/calcChain.xml')
  })

  it('keeps numbers as text when mixed with other text', () => {
    const out = t.render(fixture('report.xlsx'), { title: 2024, items: [], author: true })
    const c = cells(parse(out) as XlsxDocument)
    expect(c.A1).toBe('Report: 2024')
    expect(cells(parse(out) as XlsxDocument, 1).B1).toBe(true)
  })

  it('inspects variables', () => {
    expect(t.inspect(fixture('report.xlsx')).variables).toEqual(['title', 'items[].name', 'items[].qty', 'items[].price', 'author'])
  })
})

describe('pptx templates', () => {
  const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR4nGP4z8DwnwEIGBgAHvAD/c5k+8YAAAAASUVORK5CYII=', 'base64')
  const data = {
    title: 'Hello',
    subtitle: 'line 1\nline 2',
    points: ['a', 'b', 'c'],
    logo: { data: png },
    rows: [{ name: 'x', value: 1 }, { name: 'y', value: 2 }],
    speaker: 'Bob',
  }

  it('renders text, paragraph loops, tables and notes', () => {
    const doc = parse(t.render(fixture('deck.pptx'), data)) as PptxDocument
    const [s1, s2] = doc.slides
    const text = (sh: PptxDocument['slides'][0]['shapes'][0]) =>
      sh.type === 'text' ? sh.paragraphs.map((p) => p.children.map((c) => (c.type === 'run' ? c.text : '')).join('')).join('|') : ''
    expect(text(s1.shapes[0])).toBe('Hello')
    expect(text(s1.shapes[1])).toBe('line 1\nline 2')
    expect(text(s2.shapes[0])).toBe('• a|• b|• c')
    const table = s2.shapes.find((s) => s.type === 'table')!
    expect(table.type === 'table' && table.rows.map((r) => r.cells.map((c) => c.paragraphs[0].children.map((x) => (x.type === 'run' ? x.text : '')).join('')))).toEqual([
      ['Name', 'Value'],
      ['x', '1'],
      ['y', '2'],
    ])
    expect(s2.notes).toBe('Speaker: Bob')
  })

  it('uses a:br with run properties for line breaks', () => {
    const xml = readPart(t.render(fixture('deck.pptx'), data), 'ppt/slides/slide1.xml')
    expect(xml).toMatch(/line 1<\/a:t><\/a:r><a:br\b/)
  })

  it('replaces the image text box with a picture at the same place', () => {
    const out = t.render(fixture('deck.pptx'), data)
    const xml = readPart(out, 'ppt/slides/slide2.xml')
    expect(xml).toContain('<p:pic>')
    expect(xml).not.toContain('{{%logo}}')
    // text box is 3in x 2in at (6in, 1in); a 2:1 image fits as 3in x 1.5in, vertically centered
    const inch = 914400
    expect(xml).toContain(`<a:off x="${6 * inch}" y="${1 * inch + 0.25 * inch}"/><a:ext cx="${3 * inch}" cy="${1.5 * inch}"/>`)
    expect(partNames(out).some((n) => n.startsWith('ppt/media/image_ct'))).toBe(true)
  })

  it('removes the image shape when the value is empty', () => {
    const xml = readPart(t.render(fixture('deck.pptx'), { ...data, logo: undefined }), 'ppt/slides/slide2.xml')
    expect(xml).not.toContain('<p:pic>')
    expect(xml).not.toContain('{{%logo}}')
  })
})
