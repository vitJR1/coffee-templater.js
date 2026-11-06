import { describe, expect, it } from 'vitest'
import { createTemplater, detectFormat, engineInfo, ENGINE_VERSION, initEngine, resetEngine, TemplateError } from '../src'
import { serialize, wallClock } from '../src/engine/values'
import { body, docx, fixture, p } from './helpers'

const t = createTemplater()

describe('value serialization', () => {
  it('encodes dates, binary data and bigints', () => {
    const blobs: Uint8Array[] = []
    const bytes = new Uint8Array([1, 2, 3])
    const json = serialize({ d: new Date('2024-05-17T10:00:00Z'), b: bytes, ab: bytes.buffer, n: 12345678901234567890n }, blobs)
    expect(json).toBe('{"d":{"$date":"2024-05-17T10:00:00.000Z"},"b":{"$blob":0},"ab":{"$blob":1},"n":12345678901234567890}')
    expect(blobs).toHaveLength(2)
  })

  it('follows JSON.stringify semantics', () => {
    const json = serialize({ f: () => 1, u: undefined, arr: [undefined, () => 1], nan: NaN, j: { toJSON: () => 'x' }, m: new Map([['k', 1]]), s: new Set([1, 2]) }, [])
    expect(json).toBe('{"arr":[null,null],"nan":null,"j":"x","m":{"k":1},"s":[1,2]}')
  })

  it('escapes user objects that look like tagged values', () => {
    expect(serialize({ $date: 'not a date' }, [])).toBe('{"$literal":{"$date":"not a date"}}')
    expect(serialize({ $date: 'x', other: 1 }, [])).toBe('{"$date":"x","other":1}')
  })

  it('rejects circular data', () => {
    const a: Record<string, unknown> = {}
    a.self = a
    expect(() => serialize(a, [])).toThrow(/circular/)
  })

  it('prints wall-clock time in a time zone', () => {
    expect(wallClock(new Date('2024-07-01T12:00:00Z'), 'Europe/Berlin')).toBe('2024-07-01T14:00:00.000')
  })

  it('keeps big integers exact in output', () => {
    expect(body(t.render(docx(p('{{id}}')), { id: 12345678901234567890n }))).toBe('12345678901234567890')
  })

  it('prints data that looks like a tag literally', () => {
    expect(body(t.render(docx(p('{{x.$date}}')), { x: { $date: 'plain text' } }))).toBe('plain text')
  })
})

describe('host callbacks', () => {
  it('passes dates and bytes to custom filters as JS values', () => {
    const seen: unknown[] = []
    const tt = createTemplater({
      filters: {
        inspect: (v) => {
          seen.push(v)
          return v instanceof Date ? 'date' : v instanceof Uint8Array ? `bytes:${v.length}` : typeof v
        },
      },
    })
    const out = tt.render(docx(p('{{d | inspect}} {{b | inspect}} {{n | inspect}}')), { d: new Date(0), b: new Uint8Array(4), n: 1 })
    expect(body(out)).toBe('date bytes:4 number')
    expect(seen[0]).toEqual(new Date(0))
  })

  it('lets filters return structured values used by sections', () => {
    const tt = createTemplater({ filters: { evens: (list: number[]) => list.filter((x) => x % 2 === 0) } })
    expect(body(tt.render(docx(p('{{#each nums | evens}}{{.}};{{/each}}')), { nums: [1, 2, 3, 4] }))).toBe('2;4;')
  })

  it('overrides built-in filters', () => {
    const tt = createTemplater({ filters: { upper: (v) => `UP(${v})` } })
    expect(body(tt.render(docx(p('{{a | upper}}')), { a: 'x' }))).toBe('UP(x)')
  })

  it('reports filter exceptions with the host message', () => {
    const tt = createTemplater({ filters: { boom: () => { throw new Error('bad input') } } })
    try {
      tt.render(docx(p('{{x | boom}}')), {})
      expect.unreachable()
    } catch (e) {
      expect(e).toBeInstanceOf(TemplateError)
      expect((e as TemplateError).code).toBe('filter_failed')
      expect((e as TemplateError).message).toContain('bad input')
    }
  })

  it('refuses re-entrant calls from callbacks and keeps working afterwards', () => {
    const tt = createTemplater({ filters: { nested: () => t.render(docx(p('x')), {}) } })
    expect(() => tt.render(docx(p('{{a | nested}}')), {})).toThrow(/busy/)
    expect(body(t.render(docx(p('{{a}}')), { a: 'still fine' }))).toBe('still fine')
  })

  it('uses nullGetter and stringify', () => {
    const tt = createTemplater({ nullGetter: (tag) => `?${tag.expression}`, stringify: (v) => `<${String(v)}>` })
    expect(body(tt.render(docx(p('{{missing}} {{n}} {{s}}')), { n: 5, s: 'text' }))).toBe('?missing <5> text')
  })
})

describe('engine lifecycle', () => {
  it('reports the bundled engine version', () => {
    const info = engineInfo()
    expect(info.coreVersion).toBe(ENGINE_VERSION)
    expect(info.abiVersion).toBe(1)
    expect(info.formats).toEqual(['docx', 'xlsx', 'pptx'])
    expect(info.sha256).toMatch(/^[0-9a-f]{64}$/)
  })

  it('initializes asynchronously (browsers)', async () => {
    await expect(initEngine()).resolves.toBeUndefined()
  })

  it('recompiles templates transparently after the engine is reset', () => {
    const tpl = t.compile(docx(p('Hi {{name}}')))
    expect(body(tpl.render({ name: 'a' }))).toBe('Hi a')
    resetEngine()
    expect(body(tpl.render({ name: 'b' }))).toBe('Hi b')
  })

  it('cannot be used after dispose', () => {
    const tpl = t.compile(docx(p('{{a}}')))
    tpl.dispose()
    tpl.dispose()
    expect(() => tpl.render({})).toThrow(/disposed/)
  })

  it('exposes format and parts of a compiled template', () => {
    const tpl = t.compile(fixture('invoice.docx'))
    expect(tpl.format).toBe('docx')
    expect(tpl.parts).toContain('word/document.xml')
  })

  it('detects formats', () => {
    expect(detectFormat(fixture('report.xlsx'))).toBe('xlsx')
    expect(detectFormat(fixture('deck.pptx'))).toBe('pptx')
  })
})

describe('engine errors', () => {
  it('collects every tag mistake at once', () => {
    try {
      t.compile(docx(p('{{a | nope}}') + p('{{b c}}')))
      expect.unreachable()
    } catch (e) {
      const err = e as TemplateError
      expect(err.errors.map((x) => [x.code, x.tag])).toEqual([
        ['unknown_filter', '{{a | nope}}'],
        ['invalid_expression', '{{b c}}'],
      ])
      expect(err.code).toBe('unknown_filter')
      expect(err.message).toMatch(/\+1 more/)
    }
  })

  it('rejects files that are not Office documents', () => {
    expect(() => t.compile(new Uint8Array([1, 2, 3]))).toThrow(expect.objectContaining({ code: 'invalid_document' }))
  })

  it('enforces limits', () => {
    const tt = createTemplater({ limits: { maxEntries: 2 } })
    expect(() => tt.compile(fixture('invoice.docx'))).toThrow(expect.objectContaining({ code: 'limit_exceeded' }))
  })

  it('validates options', () => {
    expect(() => createTemplater({ delimiters: ['{{', '{{'] }).compile(docx(p('x')))).toThrow(expect.objectContaining({ code: 'invalid_options' }))
  })
})
