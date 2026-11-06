import type { BinaryInput } from './engine/values'

export type { BinaryInput }

export type DocumentFormat = 'docx' | 'xlsx' | 'pptx'

/** A filter receives the piped value and the filter arguments: `{{ value | name: arg1, arg2 }}`. */
export type FilterFn = (value: any, ...args: any[]) => unknown

/** Information about a tag, passed to `nullGetter`. */
export interface TagInfo {
  /** Full tag text including delimiters, e.g. `{{client.name}}`. */
  raw: string
  /** Expression inside the tag, e.g. `client.name | upper`. */
  expression: string
  /** Package part where the tag lives, e.g. `word/document.xml`. */
  part: string
}

export interface TagSyntax {
  /** Opens a section (loop / condition / context). Default `#` → `{{#items}}`. */
  section: string
  /** Opens an inverted section (rendered when value is falsy/empty). Default `^` → `{{^items}}`. */
  inverted: string
  /** Closes a section. Default `/` → `{{/items}}`. */
  close: string
  /** Image tag. Default `%` → `{{%logo}}`. */
  image: string
  /** Else branch keyword. Default `else` → `{{else}}`. */
  else: string
  /** Keyword after the section sigil for conditions. Default `if` → `{{#if paid}}`. */
  if: string
  /** Keyword for negative conditions. Default `unless` → `{{#unless paid}}`. */
  unless: string
  /** Keyword for explicit loops. Default `each` → `{{#each items}}`. */
  each: string
}

/** Safety limits for untrusted input. `0` disables a limit. */
export interface Limits {
  /** Size of the input file. Default 100 MiB. */
  maxInputBytes?: number
  /** Number of ZIP entries. Default 10 000. */
  maxEntries?: number
  /** Total uncompressed size. Default 500 MiB. */
  maxUncompressedBytes?: number
  /** Uncompressed / compressed ratio (zip bombs). Default 200. */
  maxCompressionRatio?: number
  /** XML nesting depth. Default 256. */
  maxXmlDepth?: number
  /** Size of the result before compression. Default 500 MiB. */
  maxOutputBytes?: number
  /** Total loop iterations per render. Default 1 000 000. */
  maxLoopIterations?: number
  /** Calls to filters / nullGetter / stringify per operation. Default 1 000 000. */
  maxHostCalls?: number
}

export interface TemplaterOptions {
  /**
   * Tag delimiters. Default `['{{', '}}']`.
   * @example { delimiters: ['${', '}'] }   // ${name}
   * @example { delimiters: { start: '[[', end: ']]' } }
   */
  delimiters?: [string, string] | { start: string; end: string }
  /**
   * Character that escapes a delimiter so it is printed literally: `\{{not a tag}}` → `{{not a tag}}`.
   * Set to `false` to disable. Default `\`.
   */
  escapeChar?: string | false
  /** Override sigils/keywords of the tag syntax. */
  syntax?: Partial<TagSyntax>
  /** Custom filters; they override built-in filters with the same name: `{{ price | money: "RUB" }}`. */
  filters?: Record<string, FilterFn>
  /** Value printed when a tag resolves to `undefined`/`null`. Default: empty string. */
  nullGetter?: (tag: TagInfo) => unknown
  /** Throw a `TemplateError` (code `missing_value`) on missing values instead of calling `nullGetter`. */
  strict?: boolean
  /** Convert `\n` in values into real line breaks (and `\t` into tabs in DOCX). Default `true`. */
  linebreaks?: boolean
  /** Converts a non-string value into text. Default: numbers as-is, Dates as ISO strings, objects as JSON. */
  stringify?: (value: unknown) => string
  /**
   * Time zone used for JavaScript `Date` values in data: `'local'` (the process time zone) or an IANA
   * name such as `'Europe/Moscow'`. By default dates are absolute instants printed in UTC — the same
   * result on every server. With a time zone, `{{ d | date: "DD.MM.YYYY" }}` shows the local day.
   */
  timeZone?: string
  /** Safety limits for templates and rendered output. */
  limits?: Limits
}

export type SectionMode = 'section' | 'inverted' | 'if' | 'unless' | 'each'

export type TemplateTag =
  | { type: 'variable'; path: string; expression: string; filters: string[]; raw: string }
  | { type: 'image'; path: string; expression: string; raw: string }
  | {
      type: 'section'
      /** `section` = mustache-style `{{#x}}`, `each` = loop, `if`/`unless`/`inverted` = condition. */
      mode: SectionMode
      path: string
      expression: string
      raw: string
      children: TemplateTag[]
      elseChildren?: TemplateTag[]
    }

export interface TemplatePartInfo {
  part: string
  tags: TemplateTag[]
}

export interface InspectResult {
  format: DocumentFormat
  /** Tag tree per package part. */
  parts: TemplatePartInfo[]
  /** Tag tree of the whole document (all parts merged, in order). */
  tags: TemplateTag[]
  /** Every data path used, loops marked with `[]`: `client.name`, `items[].price`. */
  variables: string[]
  /**
   * A sample data object with the shape the template expects
   * (strings for variables, arrays for loops, booleans for conditions, null for images).
   */
  schema: Record<string, unknown>
}

/** Value accepted by image tags (`{{%logo}}` by default). */
export interface ImageValue {
  /** Raw bytes or a base64 / data-URL string. */
  data: BinaryInput | string
  /** png | jpeg | gif | bmp | svg | ... Detected from bytes when omitted. */
  extension?: string
  /** Width in pixels (96 dpi). Detected from the image when omitted. */
  width?: number
  /** Height in pixels (96 dpi). Detected from the image when omitted. */
  height?: number
  /** Alternative text. */
  alt?: string
}
