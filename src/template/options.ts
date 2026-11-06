import { builtinFilters, type FilterFn } from './filters'

export type { FilterFn }

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
  /** Custom filters, merged with the built-in ones: `{{ price | money: "RUB" }}`. */
  filters?: Record<string, FilterFn>
  /** Value printed when a tag resolves to `undefined`/`null`. Default: empty string. */
  nullGetter?: (tag: TagInfo) => unknown
  /** Throw a `TemplateError` (code `missing_value`) on missing values instead of calling `nullGetter`. */
  strict?: boolean
  /** Convert `\n` in values into real line breaks (and `\t` into tabs in DOCX). Default `true`. */
  linebreaks?: boolean
  /** Converts a non-string value into text. Default handles numbers, booleans, Dates and objects (JSON). */
  stringify?: (value: unknown) => string
}

export interface ResolvedOptions {
  start: string
  end: string
  escapeChar: string | false
  syntax: TagSyntax
  filters: Record<string, FilterFn>
  nullGetter: (tag: TagInfo) => unknown
  strict: boolean
  linebreaks: boolean
  stringify: (value: unknown) => string
}

export const DEFAULT_SYNTAX: TagSyntax = {
  section: '#',
  inverted: '^',
  close: '/',
  image: '%',
  else: 'else',
  if: 'if',
  unless: 'unless',
  each: 'each',
}

export function defaultStringify(value: unknown): string {
  if (typeof value === 'string') return value
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return isNaN(value.getTime()) ? '' : value.toISOString()
  if (typeof value === 'object') {
    if (Array.isArray(value)) return value.map(defaultStringify).join(', ')
    const ts = (value as { toString?: () => string }).toString
    if (ts && ts !== Object.prototype.toString) return String(value)
    return JSON.stringify(value)
  }
  return String(value)
}

export function resolveOptions(options: TemplaterOptions = {}): ResolvedOptions {
  let start = '{{'
  let end = '}}'
  if (Array.isArray(options.delimiters)) [start, end] = options.delimiters
  else if (options.delimiters) ({ start, end } = options.delimiters)
  if (!start || !end) throw new Error('Delimiters must be non-empty strings')
  if (start === end) throw new Error('Start and end delimiters must differ')

  const escapeChar = options.escapeChar === undefined ? '\\' : options.escapeChar
  if (escapeChar !== false && escapeChar.length === 0) throw new Error('escapeChar must be a non-empty string or false')

  return {
    start,
    end,
    escapeChar,
    syntax: { ...DEFAULT_SYNTAX, ...options.syntax },
    filters: { ...builtinFilters, ...options.filters },
    nullGetter: options.nullGetter ?? (() => ''),
    strict: options.strict ?? false,
    linebreaks: options.linebreaks ?? true,
    stringify: options.stringify ?? defaultStringify,
  }
}
