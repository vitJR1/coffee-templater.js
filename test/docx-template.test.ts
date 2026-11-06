import { describe, expect, it } from 'vitest'
import { createTemplater, render, TemplateError } from '../src'
import { body, docx, fixture, p, partNames, readPart } from './helpers'

const t = createTemplater()

describe('variables', () => {
  it('replaces a tag split across runs', () => {
    const out = t.render(docx(p('Hello, ', '{{na', 'me}}', '!')), { name: 'World' })
    expect(body(out)).toBe('Hello, World!')
  })

  it('escapes XML special characters in values', () => {
    const out = t.render(docx(p('{{v}}')), { v: '<b> & "q"' })
    expect(readPart(out, 'word/document.xml')).toContain('&lt;b&gt; &amp; "q"')
  })

  it('supports nested paths, indexes and scope lookup', () => {
    const out = t.render(docx(p('{{a.b.c}} {{list[1]}} {{list[-1]}} {{obj["key with space"]}}')), {
      a: { b: { c: 1 } },
      list: ['x', 'y', 'z'],
      obj: { 'key with space': 'ok' },
    })
    expect(body(out)).toBe('1 y z ok')
  })

  it('does not expose prototype properties', () => {
    const out = t.render(docx(p('[{{constructor}}][{{a.__proto__}}][{{s.length}}]')), { a: {}, s: 'abc' })
    expect(body(out)).toBe('[][][3]')
  })

  it('prints missing values via nullGetter', () => {
    expect(body(t.render(docx(p('[{{missing}}]')), {}))).toBe('[]')
    const custom = createTemplater({ nullGetter: (tag) => `<${tag.expression}?>` })
    expect(body(custom.render(docx(p('{{missing}}')), {}))).toBe('<missing?>')
  })

  it('throws on missing values in strict mode', () => {
    const strict = createTemplater({ strict: true })
    expect(() => strict.render(docx(p('{{missing}}')), {})).toThrowError(/Missing value/)
  })

  it('converts newlines and tabs into Word breaks', () => {
    const xml = readPart(t.render(docx(p('{{v}}')), { v: 'a\nb\tc' }), 'word/document.xml')
    expect(xml).toContain('a</w:t><w:br/><w:t xml:space="preserve">b</w:t><w:tab/><w:t xml:space="preserve">c')
    const off = createTemplater({ linebreaks: false })
    expect(readPart(off.render(docx(p('{{v}}')), { v: 'a\nb' }), 'word/document.xml')).not.toContain('<w:br/>')
  })

  it('adds xml:space="preserve" to text that receives values', () => {
    const tpl = docx('<w:p><w:r><w:t>{{v}}</w:t></w:r></w:p>')
    expect(readPart(t.render(tpl, { v: '  padded  ' }), 'word/document.xml')).toContain('<w:t xml:space="preserve">  padded  </w:t>')
  })

  it('stringifies numbers, booleans, dates and arrays', () => {
    const out = t.render(docx(p('{{n}}|{{b}}|{{d}}|{{arr}}')), { n: 1.5, b: false, d: new Date('2024-01-02T03:04:05Z'), arr: [1, 2] })
    expect(body(out)).toBe('1.5|false|2024-01-02T03:04:05.000Z|1, 2')
  })
})

describe('escaping', () => {
  it('prints escaped delimiters literally', () => {
    expect(body(t.render(docx(p('\\{{name}} = {{name}}')), { name: 'x' }))).toBe('{{name}} = x')
  })

  it('keeps a literal backslash before a real tag with a double escape', () => {
    expect(body(t.render(docx(p('C:\\\\{{dir}}')), { dir: 'tmp' }))).toBe('C:\\tmp')
  })

  it('leaves other backslashes untouched', () => {
    expect(body(t.render(docx(p('a\\b {{x}}')), { x: 1 }))).toBe('a\\b 1')
  })

  it('supports a custom escape character or disabling it', () => {
    expect(body(createTemplater({ escapeChar: '!' }).render(docx(p('!{{a}} {{a}}')), { a: 1 }))).toBe('{{a}} 1')
    expect(body(createTemplater({ escapeChar: false }).render(docx(p('\\{{a}}')), { a: 1 }))).toBe('\\1')
  })
})

describe('custom delimiters and syntax', () => {
  it('uses ${ } delimiters', () => {
    const tt = createTemplater({ delimiters: ['${', '}'] })
    expect(body(tt.render(docx(p('Hi ${name}, {{untouched}}')), { name: 'Ann' }))).toBe('Hi Ann, {{untouched}}')
  })

  it('uses delimiters with XML-special characters', () => {
    const tt = createTemplater({ delimiters: { start: '<<', end: '>>' } })
    expect(body(tt.render(docx(p('Hi &lt;&lt;na', 'me&gt;&gt;')), { name: 'Ann' }))).toBe('Hi Ann')
  })

  it('uses single-brace delimiters', () => {
    const tt = createTemplater({ delimiters: ['{', '}'] })
    expect(body(tt.render(docx(p('{#items}{.};{/items}')), { items: [1, 2] }))).toBe('1;2;')
  })

  it('allows overriding sigils and keywords', () => {
    const tt = createTemplater({ syntax: { section: '@', close: 'end ', if: 'когда', else: 'иначе' } })
    expect(body(tt.render(docx(p('{{@когда ok}}да{{иначе}}нет{{end когда}}')), { ok: false }))).toBe('нет')
  })
})

describe('sections', () => {
  it('repeats paragraphs when section tags are alone in their paragraphs', () => {
    const tpl = docx(p('Before') + p('{{#items}}') + p('- {{name}}') + p('{{/items}}') + p('After'))
    expect(body(t.render(tpl, { items: [{ name: 'a' }, { name: 'b' }] }))).toBe('Before\n- a\n- b\nAfter')
    expect(body(t.render(tpl, { items: [] }))).toBe('Before\nAfter')
  })

  it('repeats inline content within a paragraph', () => {
    const out = t.render(docx(p('Tags: {{#tags}}', '[{{.}}]', '{{/tags}}.')), { tags: ['x', 'y'] })
    expect(body(out)).toBe('Tags: [x][y].')
  })

  it('repeats table rows', () => {
    const cell = (s: string) => `<w:tc><w:p><w:r><w:t>${s}</w:t></w:r></w:p></w:tc>`
    const tpl = docx(`<w:tbl><w:tr>${cell('Name')}${cell('Qty')}</w:tr><w:tr>${cell('{{#rows}}{{name}}')}${cell('{{qty}}{{/rows}}')}</w:tr></w:tbl>`)
    const xml = readPart(t.render(tpl, { rows: [{ name: 'a', qty: 1 }, { name: 'b', qty: 2 }] }), 'word/document.xml')
    expect(xml.match(/<w:tr>/g)).toHaveLength(3)
    expect(xml).not.toContain('{{')
    expect(body(t.render(tpl, { rows: [] }))).toBe('Name\nQty')
  })

  it('repeats multiple rows', () => {
    const row = (s: string) => `<w:tr><w:tc><w:p><w:r><w:t>${s}</w:t></w:r></w:p></w:tc></w:tr>`
    const tpl = docx(`<w:tbl>${row('{{#items}}{{name}}')}${row('{{price}}{{/items}}')}</w:tbl>`)
    expect(body(t.render(tpl, { items: [{ name: 'a', price: 1 }, { name: 'b', price: 2 }] }))).toBe('a\n1\nb\n2')
  })

  it('provides loop locals and parent scope access', () => {
    const tpl = docx(p('{{#items}}{{@number}}/{{@length}} {{name}} of {{../title}}{{#@last}}.{{/@last}}{{^@last}}, {{/@last}}{{/items}}'))
    expect(body(t.render(tpl, { title: 'T', items: [{ name: 'a' }, { name: 'b' }] }))).toBe('1/2 a of T, 2/2 b of T.')
  })

  it('looks up outer variables from inside loops', () => {
    expect(body(t.render(docx(p('{{#items}}{{currency}}{{.}} {{/items}}')), { currency: '$', items: [1, 2] }))).toBe('$1 $2 ')
  })

  it('nests loops', () => {
    const tpl = docx(p('{{#groups}}{{name}}:{{#items}}{{.}}{{/items}};{{/groups}}'))
    expect(body(t.render(tpl, { groups: [{ name: 'a', items: [1, 2] }, { name: 'b', items: [3] }] }))).toBe('a:12;b:3;')
  })

  it('uses an object as context', () => {
    expect(body(t.render(docx(p('{{#client}}{{name}} ({{city}}){{/client}}')), { client: { name: 'Ann', city: 'Oslo' } }))).toBe('Ann (Oslo)')
  })

  it('iterates objects with #each and @key', () => {
    expect(body(t.render(docx(p('{{#each prices}}{{@key}}={{.}};{{/each}}')), { prices: { a: 1, b: 2 } }))).toBe('a=1;b=2;')
  })

  it('supports if / else / unless / inverted', () => {
    const tpl = docx(p('{{#if paid}}paid{{else}}due{{/if}}|{{#unless paid}}!{{/unless}}|{{^items}}none{{/items}}'))
    expect(body(t.render(tpl, { paid: true, items: [1] }))).toBe('paid||')
    expect(body(t.render(tpl, { paid: false, items: [] }))).toBe('due|!|none')
  })

  it('supports else in paragraph sections', () => {
    const tpl = docx(p('{{#if vip}}') + p('VIP') + p('{{else}}') + p('Regular') + p('{{/if}}'))
    expect(body(t.render(tpl, { vip: true }))).toBe('VIP')
    expect(body(t.render(tpl, { vip: false }))).toBe('Regular')
  })

  it('accepts filters in section expressions', () => {
    const tpl = docx(p('{{#each items | sort: "n", "desc"}}{{n}}{{/each}}|{{#if total | gt: 100}}big{{else}}small{{/if}}'))
    expect(body(t.render(tpl, { items: [{ n: 1 }, { n: 3 }, { n: 2 }], total: 150 }))).toBe('321|big')
  })

  it('closes sections with a generic {{/}}', () => {
    expect(body(t.render(docx(p('{{#items}}{{.}}{{/}}')), { items: [1, 2] }))).toBe('12')
  })
})

describe('filters', () => {
  it('applies built-in filters', () => {
    const tpl = docx(p('{{name | upper}}|{{missing | default: "—"}}|{{n | number: 2}}|{{d | date: "DD.MM.YYYY"}}|{{list | join: " / "}}|{{text | truncate: 5}}'))
    const out = t.render(tpl, { name: 'ann', n: 1234567.891, d: new Date(Date.UTC(2024, 0, 31)), list: ['a', 'b'], text: 'abcdefgh' })
    expect(body(out)).toBe('ANN|—|1 234 567,89|31.01.2024|a / b|abcd…')
  })

  it('formats dates in UTC by default and in a configured time zone', () => {
    const tpl = docx(p('{{d | date: "DD.MM.YYYY HH:mm"}}'))
    const d = new Date('2024-01-30T21:30:00Z')
    expect(body(t.render(tpl, { d }))).toBe('30.01.2024 21:30')
    expect(body(createTemplater({ timeZone: 'Europe/Moscow' }).render(tpl, { d }))).toBe('31.01.2024 00:30')
    expect(body(createTemplater({ timeZone: 'America/New_York' }).render(tpl, { d }))).toBe('30.01.2024 16:30')
    expect(() => createTemplater({ timeZone: 'Nowhere/City' })).toThrow(RangeError)
  })

  it('accepts typographic quotes inserted by Word autocorrect', () => {
    expect(body(t.render(docx(p('{{x | default: “none”}} {{x | default: «нет»}}')), {}))).toBe('none нет')
  })

  it('supports custom filters with arguments', () => {
    const tt = createTemplater({ filters: { money: (v, cur: string) => `${Number(v).toFixed(2)} ${cur}` } })
    expect(body(tt.render(docx(p('{{price | money: "RUB"}}')), { price: 5 }))).toBe('5.00 RUB')
  })

  it('passes variables as filter arguments', () => {
    expect(body(t.render(docx(p('{{a | default: b}}')), { b: 'fallback' }))).toBe('fallback')
  })

  it('rejects unknown filters at compile time', () => {
    expect(() => t.compile(docx(p('{{x | nope}}')))).toThrowError(/Unknown filter "nope"/)
  })

  it('wraps filter exceptions', () => {
    const tt = createTemplater({ filters: { boom: () => { throw new Error('bad') } } })
    expect(() => tt.render(docx(p('{{x | boom}}')), {})).toThrowError(/Filter failed: bad/)
  })
})

describe('errors', () => {
  const code = (fn: () => unknown) => {
    try {
      fn()
    } catch (e) {
      expect(e).toBeInstanceOf(TemplateError)
      return (e as TemplateError).code
    }
    throw new Error('did not throw')
  }

  it('reports template mistakes with codes', () => {
    expect(code(() => t.compile(docx(p('{{name')))) ).toBe('unclosed_tag')
    expect(code(() => t.compile(docx(p('{{#items}}x'))))).toBe('unclosed_section')
    expect(code(() => t.compile(docx(p('x{{/items}}'))))).toBe('unmatched_close')
    expect(code(() => t.compile(docx(p('{{#a}}{{/b}}'))))).toBe('unmatched_close')
    expect(code(() => t.compile(docx(p('{{a b}}'))))).toBe('invalid_expression')
    expect(code(() => t.compile(docx(p('{{}}'))))).toBe('invalid_expression')
  })

  it('includes the part and tag in the message', () => {
    expect(() => t.compile(docx(p('{{#items}}')))).toThrowError(/\{\{#items\}\}.*word\/document\.xml/)
  })

  it('does not let a tag span paragraphs', () => {
    expect(code(() => t.compile(docx(p('{{na') + p('me}}'))))).toBe('unclosed_tag')
  })

  it('rejects sections crossing structure boundaries', () => {
    const tpl = docx(p('{{#a}} text') + `<w:tbl><w:tr><w:tc>${p('{{/a}}')}</w:tc></w:tr></w:tbl>`)
    expect(code(() => t.compile(tpl))).toBe('invalid_section')
  })
})

describe('fixture: invoice.docx', () => {
  const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR4nGP4z8DwnwEIGBgAHvAD/c5k+8YAAAAASUVORK5CYII=', 'base64')
  const data = {
    company: 'ACME',
    number: 42,
    client: { name: 'Иван' },
    items: [{ name: 'Coffee', qty: 2, price: 10.5 }, { name: 'Tea', qty: 1, price: 4 }],
    paid: true,
    notes: 'line 1\nline 2',
    logo: { data: png, width: 120 },
    total: 25,
  }

  it('renders body and header', () => {
    const out = render(fixture('invoice.docx'), data)
    const text = body(out)
    expect(text).toContain('Invoice 42')
    expect(text).toContain('Client: Иван')
    expect(text).toContain('- Coffee: 10,50\n- Tea: 4,00')
    expect(text).toContain('Status: paid')
    expect(text).toContain('Escaped: {{literal}}')
    expect(text).not.toMatch(/\{\{(?!literal)/)
    const header = partNames(out).find((n) => /word\/header\d*\.xml/.test(n))!
    expect(readPart(out, header)).toContain('ACME — confidential')
  })

  it('embeds images with relationships and content types', () => {
    const out = render(fixture('invoice.docx'), data)
    expect(partNames(out).some((n) => /^word\/media\/image_ct\d+\.png$/.test(n))).toBe(true)
    expect(readPart(out, '[Content_Types].xml')).toMatch(/Extension="png"/)
    const rels = readPart(out, 'word/_rels/document.xml.rels')
    const rId = /Id="(rId\d+)"[^>]*Target="media\/image_ct1\.png"/.exec(rels)?.[1] ?? /Target="media\/image_ct1\.png"[^>]*Id="(rId\d+)"/.exec(rels)?.[1]
    expect(rId).toBeTruthy()
    const doc = readPart(out, 'word/document.xml')
    expect(doc).toContain(`r:embed="${rId}"`)
    // 120px wide, aspect ratio 2:1 from the PNG header → 60px high
    expect(doc).toContain(`cx="${120 * 9525}" cy="${60 * 9525}"`)
  })

  it('removes the image tag when no image is given', () => {
    const out = render(fixture('invoice.docx'), { ...data, logo: null })
    expect(readPart(out, 'word/document.xml')).not.toContain('<w:drawing>')
  })

  it('renders a compiled template many times independently', () => {
    const tpl = t.compile(fixture('invoice.docx'))
    const a = body(tpl.render({ ...data, number: 1 }))
    const b = body(tpl.render({ ...data, number: 2 }))
    expect(a).toContain('Invoice 1')
    expect(b).toContain('Invoice 2')
    expect(partNames(tpl.render({ ...data, logo: null })).filter((n) => n.includes('media'))).toHaveLength(0)
  })

  it('inspects variables and the expected data shape', () => {
    const info = t.inspect(fixture('invoice.docx'))
    expect(info.format).toBe('docx')
    expect(info.variables).toEqual(
      expect.arrayContaining(['company', 'number', 'client.name', 'items[].name', 'items[].price', 'items[].qty', 'paid', 'notes', 'logo', 'total']),
    )
    expect(info.schema).toMatchObject({
      company: '',
      client: { name: '' },
      items: [{ name: '', price: '', qty: '' }],
      paid: true,
      logo: null,
    })
  })
})
