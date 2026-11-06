import { TemplateError } from '../errors'
import { compileExpression, ExpressionError, type Expression } from './expression'
import type { ResolvedOptions } from './options'

export type SectionMode = 'section' | 'inverted' | 'if' | 'unless' | 'each'

export type ParsedTag =
  | { kind: 'var'; raw: string; expr: Expression }
  | { kind: 'image'; raw: string; expr: Expression }
  | { kind: 'open'; raw: string; mode: SectionMode; expr: Expression; closeName: string }
  | { kind: 'close'; raw: string; name: string }
  | { kind: 'else'; raw: string }

export function parseTag(inner: string, raw: string, opts: ResolvedOptions, part: string): ParsedTag {
  const s = opts.syntax
  const body = inner.trim()
  const compile = (src: string): Expression => {
    try {
      const expr = compileExpression(src)
      for (const f of expr.filters) {
        if (!opts.filters[f]) throw new TemplateError(`Unknown filter "${f}"`, { code: 'unknown_filter', part, tag: raw })
      }
      return expr
    } catch (e) {
      if (e instanceof TemplateError) throw e
      const msg = e instanceof ExpressionError ? e.message : String(e)
      throw new TemplateError(`Invalid expression: ${msg}`, { code: 'invalid_expression', part, tag: raw })
    }
  }

  if (body === s.else) return { kind: 'else', raw }
  if (body.startsWith(s.close)) return { kind: 'close', raw, name: body.slice(s.close.length).trim() }
  if (body.startsWith(s.image)) return { kind: 'image', raw, expr: compile(body.slice(s.image.length)) }
  if (body.startsWith(s.inverted)) {
    const src = body.slice(s.inverted.length).trim()
    return { kind: 'open', raw, mode: 'inverted', expr: compile(src), closeName: closeNameOf(src) }
  }
  if (body.startsWith(s.section)) {
    const src = body.slice(s.section.length).trim()
    for (const mode of ['if', 'unless', 'each'] as const) {
      const kw = s[mode]
      if (src === kw || src.startsWith(kw + ' ')) {
        const rest = src.slice(kw.length).trim()
        if (!rest) throw new TemplateError(`"${kw}" needs an expression`, { code: 'invalid_expression', part, tag: raw })
        return { kind: 'open', raw, mode, expr: compile(rest), closeName: kw }
      }
    }
    return { kind: 'open', raw, mode: 'section', expr: compile(src), closeName: closeNameOf(src) }
  }
  if (!body) throw new TemplateError('Empty tag', { code: 'invalid_expression', part, tag: raw })
  return { kind: 'var', raw, expr: compile(body) }
}

/** `items | sort: "x"` closes with `{{/items}}`. */
function closeNameOf(src: string): string {
  return src.split('|')[0].trim()
}

export function closes(open: Extract<ParsedTag, { kind: 'open' }>, close: Extract<ParsedTag, { kind: 'close' }>): boolean {
  return close.name === '' || close.name === open.closeName
}
