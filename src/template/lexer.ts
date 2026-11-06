import { TemplateError } from '../errors'
import type { ResolvedOptions } from './options'

export type LexPart = { type: 'text'; value: string } | { type: 'tag'; raw: string; inner: string }

/**
 * Finds tags in a sequence of text segments that belong to one paragraph/cell.
 * Office splits text into runs arbitrarily (spell-check, formatting, editing history), so `{{name}}`
 * may arrive as `{{na` + `me}}`. The concatenated text is scanned, and every tag is moved entirely into
 * the segment where it starts — it inherits that run's formatting. Escaped delimiters are unescaped.
 */
export function lexSegments(segments: string[], opts: ResolvedOptions, part: string): LexPart[][] {
  const full = segments.join('')
  const { start, end, escapeChar } = opts

  type Edit = { from: number; to: number; tag?: { raw: string; inner: string } }
  const edits: Edit[] = []

  let i = 0
  while (i < full.length) {
    const next = full.indexOf(start, i)
    if (next === -1) break
    if (escapeChar && next >= escapeChar.length && full.startsWith(escapeChar, next - escapeChar.length)) {
      const escStart = next - escapeChar.length
      const doubled = escStart >= escapeChar.length && full.startsWith(escapeChar, escStart - escapeChar.length)
      if (!doubled) {
        // `\{{` → literal `{{`: drop the escape character
        edits.push({ from: escStart, to: next })
        i = next + start.length
        continue
      }
      // `\\{{tag}}` → literal `\` followed by a real tag: drop one escape character
      edits.push({ from: escStart, to: next })
    }
    const close = full.indexOf(end, next + start.length)
    if (close === -1) {
      throw new TemplateError(`Unclosed tag: missing "${end}"`, {
        code: 'unclosed_tag',
        part,
        tag: full.slice(next, next + 40),
      })
    }
    const inner = full.slice(next + start.length, close)
    const nested = inner.indexOf(start)
    if (nested !== -1) {
      throw new TemplateError(`Unclosed tag: missing "${end}"`, { code: 'unclosed_tag', part, tag: full.slice(next, next + start.length + nested) })
    }
    const to = close + end.length
    edits.push({ from: next, to, tag: { raw: full.slice(next, to), inner } })
    i = to
  }

  if (edits.length === 0) return segments.map((s) => [{ type: 'text', value: s }])

  const result: LexPart[][] = []
  let offset = 0
  let e = 0
  for (const seg of segments) {
    const segStart = offset
    const segEnd = offset + seg.length
    const parts: LexPart[] = []
    let cursor = segStart
    // skip edits that ended before this segment
    while (e < edits.length && edits[e].to <= segStart && !(edits[e].tag && edits[e].from >= segStart)) e++
    let k = e
    while (k < edits.length && edits[k].from < segEnd) {
      const ed = edits[k]
      if (ed.from > cursor) pushText(parts, full.slice(cursor, ed.from))
      if (ed.tag && ed.from >= segStart) parts.push({ type: 'tag', ...ed.tag })
      cursor = Math.max(cursor, Math.min(segEnd, ed.to))
      k++
    }
    if (cursor < segEnd) pushText(parts, full.slice(cursor, segEnd))
    result.push(parts)
    offset = segEnd
  }
  return result
}

function pushText(parts: LexPart[], value: string) {
  if (!value) return
  const last = parts[parts.length - 1]
  if (last && last.type === 'text') last.value += value
  else parts.push({ type: 'text', value })
}
