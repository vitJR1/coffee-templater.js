# coffee-templater

Open-source templating engine for **Word (.docx)**, **Excel (.xlsx)** and **PowerPoint (.pptx)** — plus a
JSON document model to parse Office files on the backend and build them back.

Powered by [coffee-templater-engine](https://github.com/vitJR1/coffee-templater-engine) — a Rust core
compiled to WebAssembly and bundled inside this package. The same engine is meant to back Go and Python
libraries, so documents render identically everywhere.

- Variables, loops, conditions, `else`, filters, images — **all free and MIT-licensed**
- Configurable tag delimiters (`{{x}}`, `${x}`, `[[x]]`, `<<x>>`, …) and syntax sigils
- Escaping: `\{{not a tag}}` prints the delimiters literally; values are always XML-escaped
- Tags split by Word into several runs (`{{na` + `me}}`) are handled transparently
- Loops repeat inline text, whole paragraphs, table rows (Word, PowerPoint) or sheet rows (Excel)
- Excel: numbers stay numbers, rows below a loop shift down, formulas / merged cells are adjusted
- `inspect()` lists template variables and returns a sample data shape — build forms or validate input
- `parse()` → plain JSON model → `build()`: edit documents programmatically without losing formatting
- Safe for untrusted uploads: limits against zip bombs, deep XML and runaway loops
- TypeScript-first, no native dependencies (the engine is WebAssembly), works in Node.js ≥ 18 and
  browsers (ESM + CJS)

```bash
npm install coffee-templater

# or straight from GitHub (a tag, branch or commit):
npm install github:vitJR1/coffee-templater.js#v0.1.1
# or the prebuilt package attached to a GitHub Release:
npm install https://github.com/vitJR1/coffee-templater.js/releases/download/v0.1.1/coffee-templater-0.1.1.tgz
```

## Quick start

```ts
import { readFileSync, writeFileSync } from 'node:fs'
import { createTemplater } from 'coffee-templater'

const templater = createTemplater() // options below

const output = templater.render(readFileSync('invoice.docx'), {
  number: 42,
  client: { name: 'Ivan Petrov' },
  items: [
    { name: 'Coffee', qty: 2, price: 10.5 },
    { name: 'Tea', qty: 1, price: 4 },
  ],
  paid: false,
})

writeFileSync('invoice-42.docx', output)
```

The same call works for `.xlsx` and `.pptx` — the format is detected from the file.

Compile once, render many times (much faster for batch generation):

```ts
const template = templater.compile(readFileSync('invoice.docx'))
for (const invoice of invoices) save(template.render(invoice))
```

## Template syntax

| Syntax                                   | Meaning                                                     |
| ---------------------------------------- | ----------------------------------------------------------- |
| `{{name}}`                               | variable                                                    |
| `{{client.address.city}}`, `{{items[0].name}}`, `{{map["key with spaces"]}}` | paths                    |
| `{{price \| number: 2}}`                 | filter with arguments                                       |
| `{{#items}}…{{/items}}`                  | loop over an array / context object / truthy check         |
| `{{#each items}}…{{/each}}`              | explicit loop (arrays and objects, `@key` for objects)      |
| `{{#if paid}}…{{else}}…{{/if}}`          | condition                                                   |
| `{{#unless paid}}…{{/unless}}`           | negative condition                                          |
| `{{^items}}…{{/items}}`                  | inverted section: rendered when value is empty/falsy        |
| `{{/}}`                                  | closes the innermost section                                |
| `{{%logo}}`                              | image                                                       |
| `\{{text}}`                              | literal `{{text}}` (escaped delimiter)                      |

Inside loops:

| Expression              | Value                                       |
| ----------------------- | ------------------------------------------- |
| `{{.}}` / `{{this}}`    | current item                                |
| `{{name}}`              | property of the current item, then outer scopes |
| `{{../title}}`          | value from the parent scope                 |
| `{{@index}}` / `{{@number}}` | 0-based / 1-based position             |
| `{{@first}}`, `{{@last}}`, `{{@length}}`, `{{@key}}` | loop info       |

String literals accept `"…"`, `'…'` and the typographic quotes Word inserts automatically (`“…”`, `«…»`).

### How loops expand

The engine decides what to repeat from where the opening and closing tags are:

| Tags placement                                       | Repeated content                      |
| ---------------------------------------------------- | ------------------------------------- |
| in the same paragraph / cell                         | the text between the tags             |
| each alone in its own paragraph                      | the paragraphs between them (tag paragraphs are removed) |
| in different cells of one table row (Word/PowerPoint)| the whole table row                   |
| in different cells of one row (Excel)                | the whole sheet row                   |
| in different rows                                    | the rows from the first to the last   |

```
Word paragraph loop          Word / PowerPoint table row          Excel row
──────────────────           ─────────────────────────────        ─────────────────────────────────────
{{#items}}                   | {{#items}}{{name}} | {{qty}}{{/items}} |    A3: {{#items}}{{name}}  B3: {{qty}}  C3: {{/items}}
• {{name}} — {{price}}
{{/items}}
```

`{{else}}` in a paragraph/row loop must be alone in its own paragraph/row.

### Filters

Built-in: `upper`, `lower`, `capitalize`, `trim`, `replace`, `truncate`, `prepend`, `append`, `default`,
`number` (`{{n | number: 2, " ", ","}}` → `1 234,50`), `fixed`, `round`, `currency`
(`{{n | currency: "EUR", "de-DE"}}`), `percent`, `date` (`{{d | date: "DD.MM.YYYY HH:mm"}}`), `join`, `length`,
`first`, `last`, `reverse`, `sort` (`{{#each items | sort: "price", "desc"}}`), `where`, `slice`, `json`,
comparison/logic for conditions: `eq`, `ne`, `gt`, `gte`, `lt`, `lte`, `not`, `and`, `or`, `contains`, `empty`,
and `yesno` (`{{paid | yesno: "Paid", "Due"}}`).

Custom filters:

```ts
const templater = createTemplater({
  filters: {
    money: (value, currency = 'RUB') => `${Number(value).toFixed(2)} ${currency}`,
  },
})
// {{ total | money: "USD" }}
```

Unknown filters are reported when the template is compiled.

### Images

```ts
templater.render(template, {
  logo: { data: pngBytes, width: 120 }, // Uint8Array | ArrayBuffer | Buffer | base64 | data URL
})
```

`width`/`height` are in pixels; if only one is given the other follows the aspect ratio, if none — the
natural size. PNG, JPEG, GIF, BMP and SVG are detected automatically (`extension` overrides).

- **Word** — the picture is inserted inline where the tag is.
- **PowerPoint** — put `{{%logo}}` in a text box: the box is replaced by the picture, fitted and centered
  in the box's area (placeholders inherit the area from the layout).
- **Excel** — not supported yet.

A falsy value removes the tag (and in PowerPoint the text box).

### Excel specifics

- A cell that contains only one tag gets the value's type: numbers and booleans become real numeric/boolean
  cells; text around the tag keeps the cell a string.
- When loops add or remove rows, everything below is shifted: cell references, formulas (`=SUM(D3:D3)`
  becomes `=SUM(D3:D12)`, per-row formulas like `=B3*C3` follow their row), merged cells, conditional
  formatting, data validation and hyperlinks. `calcChain.xml` is dropped and Excel recalculates on open.

## Options

```ts
createTemplater({
  delimiters: ['{{', '}}'],   // or { start: '${', end: '}' }
  escapeChar: '\\',           // `\{{x}}` → literal; `false` disables escaping
  syntax: {                   // all sigils/keywords are configurable
    section: '#', inverted: '^', close: '/', image: '%',
    else: 'else', if: 'if', unless: 'unless', each: 'each',
  },
  filters: {},                // merged with built-ins
  nullGetter: (tag) => '',    // value for undefined/null (tag.raw, tag.expression, tag.part)
  strict: false,              // throw TemplateError 'missing_value' instead of nullGetter
  linebreaks: true,           // '\n' → line break, '\t' → tab (Word)
  stringify: (value) => '…',  // how non-strings are printed (default: numbers, ISO dates, JSON)
  timeZone: 'Europe/Moscow',  // zone for JS Dates in data: 'local' or IANA name (default: UTC)
  limits: { maxEntries: 10_000, maxLoopIterations: 1_000_000 /* … */ }, // safety limits, 0 = off
})
```

**Dates.** A JS `Date` is an instant and is formatted in **UTC** by default, so a template renders the
same on every server. Set `timeZone` to print local time. Date strings without a zone
(`"2024-05-17"`) are printed as written.

**Browsers.** Optionally call `await initEngine()` once at startup: it compiles the WebAssembly engine
asynchronously instead of on the first call.

## Inspecting templates

```ts
const info = templater.inspect(readFileSync('invoice.docx'))

info.variables
// ['number', 'client.name', 'items[].name', 'items[].qty', 'items[].price', 'paid', 'logo']

info.schema
// { number: '', client: { name: '' }, items: [{ name: '', qty: '', price: '' }], paid: true, logo: null }

info.tags // full tag tree per part (sections with children, filters, raw tag text)
```

## Errors

Template mistakes throw a `TemplateError` with a `code`, the `tag` and the package `part`:

```ts
import { TemplateError } from 'coffee-templater'

try {
  templater.compile(file)
} catch (e) {
  if (e instanceof TemplateError) console.log(e.code, e.tag, e.part)
  // 'unclosed_section' '{{#items}}' 'word/document.xml'
}
```

Template codes: `unclosed_tag`, `unclosed_section`, `unmatched_close`, `invalid_expression`, `unknown_filter`,
`filter_failed`, `missing_value`, `invalid_section`, `invalid_image`, `unsupported`.
Engine codes: `invalid_document`, `unsupported_format`, `invalid_model`, `invalid_options`,
`limit_exceeded`, `host_error`, `internal`.

All mistakes in tags are reported at once: `e.errors` lists every `{ code, message, tag, part }`,
and the error itself describes the first one.

## Document model: parse → JSON → build

```ts
import { parse, build, extractText } from 'coffee-templater'

const original = readFileSync('contract.docx')
const doc = parse(original) // DocxDocument | XlsxDocument | PptxDocument (plain JSON)

// store it, send it to a frontend, edit it...
const body = doc.format === 'docx' && doc.parts.find((p) => p.kind === 'body')!
body.blocks.push({
  type: 'paragraph',
  style: 'Heading2',
  children: [{ type: 'run', text: 'Appendix', bold: true }],
})

const updated = build(doc, { base: original })
```

Formatting is kept: every paragraph/run/cell carries its original properties XML in `props`, and the
convenient fields (`bold`, `size`, `color`, `align`, …) are applied on top only when you change them.
Anything the model does not cover (images, fields, bookmarks, content controls) is kept as `raw` XML.

### Word (`DocxDocument`)

```ts
{
  format: 'docx',
  parts: [{
    path: 'word/document.xml', kind: 'body',           // also 'header' / 'footer' parts
    blocks: [
      { type: 'paragraph', style: 'Heading1', align: 'center', list: { id: 1, level: 0 },
        children: [
          { type: 'run', text: 'Hello\nworld', bold: true, italic: true, underline: true, strike: false,
            color: 'FF0000', highlight: 'yellow', size: 14, font: 'Arial', vertAlign: 'superscript' },
          { type: 'hyperlink', url: 'https://example.com', children: [{ type: 'run', text: 'link' }] },
          { type: 'break', kind: 'page' },
          { type: 'raw', xml: '<w:r><w:drawing>…</w:drawing></w:r>' },
        ] },
      { type: 'table', rows: [{ cells: [{ blocks: [/* paragraphs */], colSpan: 2, vMerge: 'restart' }] }] },
    ],
  }],
}
```

`\n` and `\t` in run text are line breaks and tabs. `build()` without `base` creates a new document with
default styles (`Normal`, `Title`, `Heading1-3`, `ListParagraph`, `TableGrid`).

### Excel (`XlsxDocument`)

```ts
{
  format: 'xlsx',
  sheets: [{
    name: 'Report', path: 'xl/worksheets/sheet1.xml',
    merges: ['A1:D1'],
    rows: [{ index: 1, height: 20, cells: [
      { ref: 'A1', value: 'Name' },
      { ref: 'B1', value: 42, style: 3 },
      { ref: 'C1', value: '2024-05-17', type: 'date' },   // date-formatted cells become ISO strings
      { ref: 'D1', value: 84, formula: 'B1*2' },
      { ref: 'E1', value: true },
    ] }],
  }],
}
```

Helpers for backend imports/exports:

```ts
import { parseXlsxFile, sheetToObjects, sheetToArrays, sheetFromArrays, build } from 'coffee-templater'

const rows = sheetToObjects(parseXlsxFile(upload).sheets[0]) // [{ Name: 'Tea', Qty: 2 }, ...]

const file = build({
  format: 'xlsx',
  sheets: [sheetFromArrays('Export', [['Name', 'Date'], ['Ann', new Date()]])],
}) // no base needed
```

Renaming, adding and removing sheets is supported. Shared formulas are expanded into regular formulas.

### PowerPoint (`PptxDocument`)

```ts
{
  format: 'pptx',
  slides: [{
    path: 'ppt/slides/slide1.xml',
    notes: 'Speaker notes',
    shapes: [
      { type: 'text', id: 2, name: 'Title 1', placeholder: 'title',
        paragraphs: [{ align: 'center', level: 0,
          children: [{ type: 'run', text: 'Hello', bold: true, size: 40, color: '1F4E79', font: 'Arial' }] }] },
      { type: 'table', id: 4, name: 'Table 3', rows: [{ cells: [{ paragraphs: [/* … */] }] }] },
      { type: 'picture', id: 5, name: 'Logo' },  // other shapes are kept from the base file by id
    ],
  }],
}
```

Reorder or remove slides/shapes by reordering/removing them in the arrays. Table rows can be added (the last
row is cloned). Building a pptx requires `base`.

### Plain text

```ts
extractText(file) // docx: paragraphs; xlsx: tab-separated rows per sheet; pptx: text per slide
```

## Limitations / roadmap

- Images in Excel templates
- Repeating whole slides in PowerPoint (`{{#slides}}`)
- Excel: shared formulas inside repeated rows and table objects (`ListObject`) ranges are not adjusted
- Tracked changes and comments are preserved as raw XML, not modelled

Contributions are welcome!

## Development

```bash
npm install        # requires Node.js 22+ for the dev tooling
npm test           # vitest
npm run typecheck
npm run build      # tsup → dist (ESM + CJS + d.ts)
```

The engine is not vendored: `npm run engine` (run automatically before test/typecheck/build) downloads
`coffee_core.wasm` from the engine's GitHub Release pinned in `package.json` → `coffeeEngine`, verifies its
sha256 and embeds it. To develop against a local engine build:

```bash
COFFEE_ENGINE_WASM=../coffee-templater-engine/dist/coffee_core.wasm npm test
```

Test fixtures are generated with real Office libraries: see `test/fixtures/generate.py`.
See [DOCS.md](DOCS.md) (Russian) for the full guide and internals.

## License

MIT

## Support

If this project helps you, consider buying me a coffee ☕

- PayPal ypi-2-22@mail.ru
