import { describe, expect, it } from 'vitest'
import {
  build,
  extractText,
  parse,
  parseDocxFile,
  parsePptxFile,
  parseXlsxFile,
  sheetFromArrays,
  sheetToArrays,
  sheetToObjects,
  type DocxDocument,
  type DocxParagraph,
  type XlsxDocument,
} from '../src'
import { fixture, readPart } from './helpers'

const json = <T>(x: T): T => JSON.parse(JSON.stringify(x))

describe('docx model', () => {
  const doc = parseDocxFile(fixture('invoice.docx'))
  const bodyPart = doc.parts.find((p) => p.kind === 'body')!

  it('parses paragraphs, styles, runs and tables', () => {
    const [heading, client] = bodyPart.blocks as DocxParagraph[]
    expect(heading).toMatchObject({ type: 'paragraph', style: 'Heading1', children: [{ type: 'run', text: 'Invoice {{number}}' }] })
    expect(client.children).toMatchObject([
      { type: 'run', text: 'Client: ' },
      { type: 'run', text: '{{cli', bold: true },
      { type: 'run', text: 'ent.name}}' },
    ])
    const table = bodyPart.blocks.find((b) => b.type === 'table')!
    expect(table.type === 'table' && table.rows[0].cells.map((c) => (c.blocks[0] as DocxParagraph).children[0])).toMatchObject([
      { text: 'Name' },
      { text: 'Qty' },
      { text: 'Price' },
    ])
    expect(bodyPart.sectionProps).toContain('<w:pgSz')
    expect(doc.parts.some((p) => p.kind === 'header')).toBe(true)
  })

  it('is JSON-serializable and round-trips without changes', () => {
    const rebuilt = build(json(doc), { base: fixture('invoice.docx') })
    expect(parse(rebuilt)).toEqual(doc)
  })

  it('applies edits and keeps original formatting', () => {
    const edited = json(doc) as DocxDocument
    const body = edited.parts.find((p) => p.kind === 'body')!
    const total = body.blocks.find((b) => b.type === 'paragraph' && b.children.some((c) => c.type === 'run' && c.text.startsWith('Total')))! as DocxParagraph
    const run = total.children[0]
    if (run.type !== 'run') throw new Error()
    expect(run.size).toBe(14)
    run.text = 'Total: 100'
    run.bold = true
    run.color = 'FF0000'
    body.blocks.push({ type: 'paragraph', align: 'center', children: [{ type: 'run', text: 'Line 1\nLine 2\tTabbed', italic: true }] })
    body.blocks.push({
      type: 'paragraph',
      children: [{ type: 'hyperlink', url: 'https://example.com', children: [{ type: 'run', text: 'link', underline: true }] }],
    })

    const out = build(edited, { base: fixture('invoice.docx') })
    const xml = readPart(out, 'word/document.xml')
    // fields are written in schema order, the original size is kept
    expect(xml).toContain('<w:rPr><w:b/><w:color w:val="FF0000"/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">Total: 100</w:t>')
    expect(xml).toContain('<w:pPr><w:jc w:val="center"/></w:pPr>')
    expect(xml).toContain('<w:t xml:space="preserve">Line 1</w:t><w:br/><w:t xml:space="preserve">Line 2</w:t><w:tab/>')
    expect(readPart(out, 'word/_rels/document.xml.rels')).toMatch(/Target="https:\/\/example\.com" TargetMode="External"/)

    const reparsed = parseDocxFile(out).parts.find((p) => p.kind === 'body')!
    const last = reparsed.blocks[reparsed.blocks.length - 1] as DocxParagraph
    expect(last.children[0]).toMatchObject({ type: 'hyperlink', url: 'https://example.com', children: [{ text: 'link', underline: true }] })
  })

  it('removes a field when it is deleted from the model', () => {
    const edited = json(doc) as DocxDocument
    const client = edited.parts[0].blocks[1] as DocxParagraph
    const run = client.children[1]
    if (run.type !== 'run') throw new Error()
    delete run.bold
    const out = build(edited, { base: fixture('invoice.docx') })
    const again = parseDocxFile(out).parts[0].blocks[1] as DocxParagraph
    expect(again.children).toEqual([{ type: 'run', text: 'Client: {{client.name}}' }])
  })

  it('builds a new document without a base file', () => {
    const out = build({
      format: 'docx',
      parts: [
        {
          path: 'word/document.xml',
          kind: 'body',
          blocks: [
            { type: 'paragraph', style: 'Heading1', children: [{ type: 'run', text: 'Report' }] },
            {
              type: 'table',
              rows: [
                { cells: [{ blocks: [{ type: 'paragraph', children: [{ type: 'run', text: 'a', bold: true }] }] }, { blocks: [] }] },
              ],
            },
          ],
        },
      ],
    })
    expect(extractText(out)).toBe('Report\na\t')
    expect(readPart(out, 'word/styles.xml')).toContain('w:styleId="Heading1"')
  })

  it('extracts plain text', () => {
    const text = extractText(fixture('invoice.docx'))
    expect(text).toContain('Invoice {{number}}')
    expect(text).toContain('Name\tQty\tPrice')
  })
})

describe('xlsx model', () => {
  const doc = parseXlsxFile(fixture('report.xlsx'))

  it('parses sheets, typed cells, formulas and merges', () => {
    expect(doc.sheets.map((s) => s.name)).toEqual(['Report', 'Info'])
    const sheet = doc.sheets[0]
    expect(sheet.merges).toEqual(['A1:D1'])
    const d5 = sheet.rows.find((r) => r.index === 5)!.cells.find((c) => c.ref === 'D5')
    expect(d5).toMatchObject({ formula: 'SUM(D3:D3)' })
    expect(sheetToArrays(sheet)[1]).toEqual(['Name', 'Qty', 'Price', 'Sum', null])
  })

  it('round-trips without changes', () => {
    expect(parse(build(json(doc), { base: fixture('report.xlsx') }))).toEqual(doc)
  })

  it('edits values, renames and removes sheets', () => {
    const edited = json(doc) as XlsxDocument
    edited.sheets[0].name = 'Отчёт'
    edited.sheets[0].rows.push({ index: 10, cells: [{ ref: 'A10', value: 42 }, { ref: 'B10', value: true }, { ref: 'C10', value: 'text' }] })
    edited.sheets.pop()
    const out = parseXlsxFile(build(edited, { base: fixture('report.xlsx') }))
    expect(out.sheets.map((s) => s.name)).toEqual(['Отчёт'])
    expect(out.sheets[0].rows.find((r) => r.index === 10)!.cells.map((c) => c.value)).toEqual([42, true, 'text'])
  })

  it('creates a workbook from arrays, with dates', () => {
    const sheet = sheetFromArrays('Data', [
      ['Name', 'Born', 'Score'],
      ['Ann', new Date(Date.UTC(1990, 4, 17)), 9.5],
      ['Bob', new Date(Date.UTC(2001, 0, 1)), 7],
    ])
    const out = build({ format: 'xlsx', sheets: [sheet, { name: 'Empty', rows: [] }] })
    const parsed = parseXlsxFile(out)
    expect(parsed.sheets.map((s) => s.name)).toEqual(['Data', 'Empty'])
    expect(sheetToObjects(parsed.sheets[0])).toEqual([
      { Name: 'Ann', Born: '1990-05-17', Score: 9.5 },
      { Name: 'Bob', Born: '2001-01-01', Score: 7 },
    ])
    expect(parsed.sheets[0].rows[1].cells[1].type).toBe('date')
  })
})

describe('pptx model', () => {
  const doc = parsePptxFile(fixture('deck.pptx'))

  it('parses slides, shapes, tables and notes', () => {
    expect(doc.slides).toHaveLength(2)
    expect(doc.slides[0].shapes.map((s) => s.type === 'text' && s.placeholder)).toEqual(['ctrTitle', 'subTitle'])
    const table = doc.slides[1].shapes.find((s) => s.type === 'table')
    expect(table?.type === 'table' && table.rows.length).toBe(2)
    expect(doc.slides[1].notes).toBe('Speaker: {{speaker}}')
  })

  it('round-trips without changes', () => {
    expect(parse(build(json(doc), { base: fixture('deck.pptx') }))).toEqual(doc)
  })

  it('edits text, table rows and notes; reorders and removes slides', () => {
    const edited = json(doc)
    const title = edited.slides[0].shapes[0]
    if (title.type !== 'text') throw new Error()
    title.paragraphs = [{ align: 'center', children: [{ type: 'run', text: 'New title', bold: true, size: 40, color: '#00ff00' }] }]
    const table = edited.slides[1].shapes.find((s) => s.type === 'table')!
    if (table.type !== 'table') throw new Error()
    table.rows.push({ cells: [{ paragraphs: [{ children: [{ type: 'run', text: 'r3' }] }] }, { paragraphs: [] }] })
    edited.slides[1].notes = 'New notes'
    edited.slides.reverse()

    const out = build(edited, { base: fixture('deck.pptx') })
    const again = parsePptxFile(out)
    expect(again.slides.map((s) => s.path)).toEqual(['ppt/slides/slide2.xml', 'ppt/slides/slide1.xml'])
    const t = again.slides[1].shapes[0]
    expect(t.type === 'text' && t.paragraphs[0]).toMatchObject({ align: 'center', children: [{ text: 'New title', bold: true, size: 40, color: '00FF00' }] })
    const tb = again.slides[0].shapes.find((s) => s.type === 'table')!
    expect(tb.type === 'table' && tb.rows.length).toBe(3)
    expect(again.slides[0].notes).toBe('New notes')

    const removed = build({ ...edited, slides: [edited.slides[0]] }, { base: fixture('deck.pptx') })
    expect(parsePptxFile(removed).slides).toHaveLength(1)
  })

  it('requires a base file', () => {
    expect(() => build(doc)).toThrowError(/requires `base`/)
  })
})
