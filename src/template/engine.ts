import { TemplateError } from '../errors'
import { isImageValue, resolveImage, type ResolvedImage } from '../image'
import type { OpcPackage } from '../package'
import { decodeXml, escapeXml } from '../xml/escape'
import { setRawAttribute, tokenizeXml } from '../xml/tokenizer'
import { ExpressionError, Scope, type Expression } from './expression'
import { lexSegments } from './lexer'
import type { ResolvedOptions, TagInfo } from './options'
import { closes, parseTag, type ParsedTag, type SectionMode } from './tags'

// ---------------------------------------------------------------------------------------------
// Format specification: tells the engine how a given XML dialect stores text and structure.
// ---------------------------------------------------------------------------------------------

export interface FormatSpec {
  /** Elements whose character data is document text (`w:t`, `a:t`, `t`). */
  textTags: ReadonlySet<string>
  /** Elements that bound a tag: a tag can't span two of them (`w:p`, `a:p`, `c`). */
  groupTags: ReadonlySet<string>
  /** Paragraph element; a section whose tags are alone in their paragraphs repeats whole paragraphs. */
  paragraphTag?: string
  /** Table-row elements; a section spanning cells of one row repeats the row. */
  rowTags: ReadonlySet<string>
  /** Elements that count as content when deciding whether a paragraph holds only a tag. */
  contentTags: ReadonlySet<string>
  /** Add `xml:space="preserve"` to text elements that receive values (DOCX). */
  preserveSpace?: boolean
  /** How image tags are rendered. */
  imageMode: 'inline' | 'shape' | 'none'
  /** Element replaced by a picture in `shape` mode (`p:sp`). */
  shapeTag?: string
  /** XML inserted for `\n` / `\t` in values, computed from the tag position. */
  breaks?(ctx: TokenContext): { br?: string; tab?: string }
  /** Custom rendering for a value (e.g. numeric cells in XLSX). Return undefined for default handling. */
  renderScalar?(value: unknown, text: string, ctx: RenderContext): string | undefined
  renderImage?(image: ResolvedImage, ctx: ImageContext): string
}

export interface TokenContext {
  tokens: readonly Tok[]
  /** Indices of open elements enclosing the tag, outermost first. */
  stack: readonly number[]
  closeOf: ReadonlyMap<number, number>
}

export interface RenderContext {
  opts: ResolvedOptions
  part: string
  pkg: OpcPackage
  /** Shared mutable counters (unique ids for drawings, etc.). */
  state: Record<string, number>
}

export interface ImageContext extends RenderContext {
  /** Raw XML of the element being replaced (shape mode). */
  shapeXml?: string
}

// ---------------------------------------------------------------------------------------------
// Tokens and nodes
// ---------------------------------------------------------------------------------------------

export type Tok =
  | { t: 'xml'; raw: string; name?: string; kind?: 'open' | 'close' | 'self' }
  | { t: 'text'; value: string }
  | { t: 'tag'; tag: ParsedTag }

type OpenTag = Extract<ParsedTag, { kind: 'open' }>

export type Node =
  | { type: 'raw'; xml: string }
  | { type: 'text'; value: string }
  | { type: 'var'; tag: Extract<ParsedTag, { kind: 'var' }>; br?: string; tab?: string }
  | { type: 'image'; tag: Extract<ParsedTag, { kind: 'image' }>; shapeXml?: string }
  | { type: 'section'; tag: OpenTag; body: Node[]; elseBody?: Node[] }

export interface CompiledPart {
  part: string
  nodes: Node[]
  hasTags: boolean
}

interface Ranged {
  kind: 'section' | 'image'
  from: number
  to: number
  order: number
  // section
  open?: number
  close?: number
  elseAt?: number
  bodies?: Array<[number, number]>
  // image
  tagIndex?: number
}

// ---------------------------------------------------------------------------------------------
// Compilation
// ---------------------------------------------------------------------------------------------

export function compilePart(xml: string, spec: FormatSpec, opts: ResolvedOptions, part: string): CompiledPart {
  if (!xml.includes(escapeXml(opts.start.charAt(0)))) {
    return { part, nodes: [{ type: 'raw', xml }], hasTags: false }
  }
  const tokens = prepareTokens(xml, spec, opts, part)
  if (!tokens.some((t) => t.t === 'tag')) {
    return { part, nodes: [{ type: 'raw', xml: tokens.map(tokText).join('') }], hasTags: false }
  }
  const { closeOf, stacks } = computeStructure(tokens, part)
  const items = computeRanges(tokens, stacks, closeOf, spec, part)
  const nodes = buildNodes(tokens, stacks, closeOf, spec, items, part)
  return { part, nodes, hasTags: true }
}

function tokText(t: Tok): string {
  return t.t === 'xml' ? t.raw : t.t === 'text' ? escapeXml(t.value) : ''
}

/** Tokenizes XML, finds tags inside document text and splits text tokens around them. */
function prepareTokens(xml: string, spec: FormatSpec, opts: ResolvedOptions, part: string): Tok[] {
  const xt = tokenizeXml(xml)

  // 1. Collect text-container character data grouped by paragraph/cell.
  const groupSegments = new Map<number, number[]>()
  const textOpenOf = new Map<number, number>()
  const groupStack: number[] = []
  let openText = -1
  for (let i = 0; i < xt.length; i++) {
    const tok = xt[i]
    if (tok.type === 'open') {
      if (tok.selfClosing) continue
      if (spec.groupTags.has(tok.name)) groupStack.push(i)
      if (spec.textTags.has(tok.name)) openText = i
    } else if (tok.type === 'close') {
      if (spec.textTags.has(tok.name)) openText = -1
      if (spec.groupTags.has(tok.name)) groupStack.pop()
    } else if (tok.type === 'text' && openText !== -1) {
      const g = groupStack.length ? groupStack[groupStack.length - 1] : -1
      let list = groupSegments.get(g)
      if (!list) groupSegments.set(g, (list = []))
      list.push(i)
      textOpenOf.set(i, openText)
    }
  }

  // 2. Lex each group.
  const partsOf = new Map<number, ReturnType<typeof lexSegments>[number]>()
  for (const indices of groupSegments.values()) {
    const segs = indices.map((i) => decodeXml(xt[i].raw))
    const lexed = lexSegments(segs, opts, part)
    indices.forEach((idx, n) => partsOf.set(idx, lexed[n]))
  }

  // 3. Mark text elements that receive tags (xml:space="preserve").
  const preserve = new Set<number>()
  if (spec.preserveSpace) {
    for (const [idx, parts] of partsOf) if (parts.some((p) => p.type === 'tag')) preserve.add(textOpenOf.get(idx)!)
  }

  // 4. Emit engine tokens.
  const out: Tok[] = []
  for (let i = 0; i < xt.length; i++) {
    const tok = xt[i]
    const parts = partsOf.get(i)
    if (parts) {
      for (const p of parts) {
        if (p.type === 'text') out.push({ t: 'text', value: p.value })
        else out.push({ t: 'tag', tag: parseTag(p.inner, p.raw, opts, part) })
      }
      continue
    }
    if (tok.type === 'open') {
      const raw = preserve.has(i) ? setRawAttribute(tok.raw, 'xml:space', 'preserve') : tok.raw
      out.push({ t: 'xml', raw, name: tok.name, kind: tok.selfClosing ? 'self' : 'open' })
    } else if (tok.type === 'close') {
      out.push({ t: 'xml', raw: tok.raw, name: tok.name, kind: 'close' })
    } else {
      out.push({ t: 'xml', raw: tok.raw })
    }
  }
  return out
}

function computeStructure(tokens: Tok[], part: string) {
  const closeOf = new Map<number, number>()
  const stacks = new Map<number, number[]>()
  const stack: number[] = []
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]
    if (t.t === 'xml') {
      if (t.kind === 'open') stack.push(i)
      else if (t.kind === 'close') {
        const open = stack.pop()
        if (open === undefined) throw new TemplateError(`Malformed XML: unexpected </${t.name}>`, { code: 'unsupported', part })
        closeOf.set(open, i)
      }
    } else if (t.t === 'tag') {
      stacks.set(i, stack.slice())
    }
  }
  return { closeOf, stacks }
}

function computeRanges(
  tokens: Tok[],
  stacks: Map<number, number[]>,
  closeOf: Map<number, number>,
  spec: FormatSpec,
  part: string,
): Ranged[] {
  const nameAt = (i: number) => (tokens[i] as { name?: string }).name ?? ''
  const tagAt = (i: number) => (tokens[i] as { tag: ParsedTag }).tag

  /** True when element `el` contains no content except whitespace and the tag at `tagIdx`. */
  const alone = (el: number, tagIdx: number): boolean => {
    const end = closeOf.get(el)!
    for (let i = el + 1; i < end; i++) {
      const t = tokens[i]
      if (t.t === 'tag' && i !== tagIdx) return false
      if (t.t === 'text' && t.value.trim() !== '') return false
      if (t.t === 'xml' && t.name && spec.contentTags.has(t.name)) return false
    }
    return true
  }
  const sameNames = (a: readonly number[], b: readonly number[]) =>
    a.length === b.length && a.every((x, i) => nameAt(x) === nameAt(b[i]))

  // Pair open/else/close tags.
  type Pair = { open: number; close: number; elseAt?: number }
  const pairs: Pair[] = []
  const open: Array<{ idx: number; elseAt?: number }> = []
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]
    if (t.t !== 'tag') continue
    const tag = t.tag
    if (tag.kind === 'open') open.push({ idx: i })
    else if (tag.kind === 'else') {
      const top = open[open.length - 1]
      if (!top) throw new TemplateError('"else" outside of a section', { code: 'unmatched_close', part, tag: tag.raw })
      if (top.elseAt !== undefined) throw new TemplateError('Duplicate "else"', { code: 'invalid_section', part, tag: tag.raw })
      top.elseAt = i
    } else if (tag.kind === 'close') {
      const top = open.pop()
      if (!top) throw new TemplateError('Closing tag without an opening tag', { code: 'unmatched_close', part, tag: tag.raw })
      const openTag = tagAt(top.idx) as OpenTag
      if (!closes(openTag, tag)) {
        throw new TemplateError(`Closing tag does not match ${openTag.raw}`, { code: 'unmatched_close', part, tag: tag.raw })
      }
      pairs.push({ open: top.idx, close: i, elseAt: top.elseAt })
    }
  }
  if (open.length) {
    throw new TemplateError('Section is never closed', { code: 'unclosed_section', part, tag: tagAt(open[open.length - 1].idx).raw })
  }

  const items: Ranged[] = []

  for (const p of pairs) {
    const S = stacks.get(p.open)!
    const E = stacks.get(p.close)!
    const raw = tagAt(p.open).raw
    let k = 0
    while (k < S.length && k < E.length && S[k] === E[k]) k++
    const lca = k > 0 ? S[k - 1] : -1
    const inGroup = S.slice(0, k).some((i) => spec.groupTags.has(nameAt(i)))

    const inline = (): Ranged => {
      if (!sameNames(S.slice(k), E.slice(k))) {
        throw new TemplateError(
          'Opening and closing tags must be in the same paragraph, the same table row, or each alone in its own paragraph',
          { code: 'invalid_section', part, tag: raw },
        )
      }
      const item: Ranged = { kind: 'section', from: p.open, to: p.close, order: p.open, open: p.open, close: p.close }
      if (p.elseAt === undefined) item.bodies = [[p.open + 1, p.close - 1]]
      else {
        if (!sameNames(stacks.get(p.elseAt)!, S)) {
          throw new TemplateError('"else" must be at the same level as its section tags', { code: 'invalid_section', part, tag: raw })
        }
        item.bodies = [
          [p.open + 1, p.elseAt - 1],
          [p.elseAt + 1, p.close - 1],
        ]
      }
      return item
    }

    const blocks = (from: number, to: number, bodyFrom: number, bodyTo: number, depth: number, what: string): Ranged => {
      const item: Ranged = { kind: 'section', from, to, order: p.open, open: p.open, close: p.close }
      if (p.elseAt === undefined) item.bodies = [[bodyFrom, bodyTo]]
      else {
        const el = stacks.get(p.elseAt)![depth]
        if (el === undefined || !alone(el, p.elseAt) || el < from || el > to) {
          throw new TemplateError(`"else" must be alone in its own ${what}`, { code: 'invalid_section', part, tag: raw })
        }
        item.bodies = [
          [bodyFrom, el - 1],
          [closeOf.get(el)! + 1, bodyTo],
        ]
      }
      return item
    }

    let item: Ranged
    if (inGroup || lca === -1) {
      item = inline()
    } else if (spec.rowTags.has(nameAt(lca))) {
      if (p.elseAt !== undefined) {
        throw new TemplateError('"else" is not supported in a single-row section; put it in its own row', { code: 'invalid_section', part, tag: raw })
      }
      const end = closeOf.get(lca)!
      item = { kind: 'section', from: lca, to: end, order: p.open, open: p.open, close: p.close, bodies: [[lca, end]] }
    } else {
      const A = S[k]
      const B = E[k]
      if (A !== undefined && B !== undefined && spec.rowTags.has(nameAt(A)) && spec.rowTags.has(nameAt(B))) {
        const end = closeOf.get(B)!
        item = blocks(A, end, A, end, k, 'row')
      } else if (
        A !== undefined &&
        B !== undefined &&
        spec.paragraphTag &&
        nameAt(A) === spec.paragraphTag &&
        nameAt(B) === spec.paragraphTag &&
        alone(A, p.open) &&
        alone(B, p.close)
      ) {
        item = blocks(A, closeOf.get(B)!, closeOf.get(A)! + 1, B - 1, k, 'paragraph')
      } else {
        item = inline()
      }
    }
    items.push(item)
  }

  if (spec.imageMode === 'shape' && spec.shapeTag) {
    for (const [idx, stack] of stacks) {
      const tag = tagAt(idx)
      if (tag.kind !== 'image') continue
      let sp = -1
      for (let d = stack.length - 1; d >= 0; d--) if (nameAt(stack[d]) === spec.shapeTag) {
        sp = stack[d]
        break
      }
      if (sp === -1) throw new TemplateError('Image tag must be placed inside a text box/shape', { code: 'invalid_image', part, tag: tag.raw })
      items.push({ kind: 'image', from: sp, to: closeOf.get(sp)!, order: idx, tagIndex: idx })
    }
  }

  items.sort((a, b) => a.from - b.from || b.to - a.to || a.order - b.order)
  return items
}

function buildNodes(
  tokens: Tok[],
  stacks: Map<number, number[]>,
  closeOf: Map<number, number>,
  spec: FormatSpec,
  items: Ranged[],
  part: string,
): Node[] {
  const push = (nodes: Node[], node: Node) => {
    const last = nodes[nodes.length - 1]
    if (node.type === 'raw' && last?.type === 'raw') last.xml += node.xml
    else nodes.push(node)
  }

  const overlapError = (it: Ranged): never => {
    const idx = it.open ?? it.tagIndex!
    throw new TemplateError('Sections overlap: a section must be fully inside another one', {
      code: 'invalid_section',
      part,
      tag: (tokens[idx] as { tag: ParsedTag }).tag.raw,
    })
  }

  const build = (from: number, to: number, list: Ranged[]): Node[] => {
    const nodes: Node[] = []
    let idx = 0
    let i = from
    while (i <= to) {
      const item = list[idx]
      if (item && item.from < i) overlapError(item)
      if (item && item.from === i) {
        let j = idx + 1
        while (j < list.length && list[j].from <= item.to) {
          if (list[j].to > item.to) overlapError(list[j])
          j++
        }
        push(nodes, makeNode(item, list.slice(idx + 1, j)))
        idx = j
        i = item.to + 1
        continue
      }
      const t = tokens[i]
      if (t.t === 'xml') push(nodes, { type: 'raw', xml: t.raw })
      else if (t.t === 'text') push(nodes, { type: 'text', value: t.value })
      else if (t.tag.kind === 'var') {
        const b = spec.breaks?.({ tokens, stack: stacks.get(i)!, closeOf }) ?? {}
        nodes.push({ type: 'var', tag: t.tag, br: b.br, tab: b.tab })
      } else if (t.tag.kind === 'image') {
        if (spec.imageMode === 'none') {
          throw new TemplateError('Image tags are not supported in this format yet', { code: 'unsupported', part, tag: t.tag.raw })
        }
        nodes.push({ type: 'image', tag: t.tag })
      }
      // open/close/else tags render as nothing
      i++
    }
    if (idx < list.length) overlapError(list[idx])
    return nodes
  }

  const makeNode = (item: Ranged, children: Ranged[]): Node => {
    if (item.kind === 'image') {
      const tag = (tokens[item.tagIndex!] as { tag: Extract<ParsedTag, { kind: 'image' }> }).tag
      let xml = ''
      for (let i = item.from; i <= item.to; i++) {
        const t = tokens[i]
        xml += t.t === 'xml' ? t.raw : t.t === 'text' ? escapeXml(t.value) : ''
      }
      return { type: 'image', tag, shapeXml: xml }
    }
    const tag = (tokens[item.open!] as { tag: OpenTag }).tag
    const bodies = item.bodies!.map(([f, t]) => {
      const inside = children.filter((c) => c.from >= f && c.to <= t)
      for (const c of children) if (!inside.includes(c) && c.from <= t && c.to >= f) overlapError(c)
      return build(f, t, inside)
    })
    return { type: 'section', tag, body: bodies[0], elseBody: bodies[1] }
  }

  return build(0, tokens.length - 1, items)
}

// ---------------------------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------------------------

const truthy = (v: unknown) => (Array.isArray(v) ? v.length > 0 : Boolean(v))

export function renderPart(compiled: CompiledPart, data: unknown, spec: FormatSpec, ctx: RenderContext): string {
  if (!compiled.hasTags) return (compiled.nodes[0] as { xml: string }).xml
  const out: string[] = []
  renderNodes(compiled.nodes, data instanceof Scope ? data : new Scope(data), spec, ctx, out)
  return out.join('')
}

function evaluate(tag: { raw: string; expr: Expression }, scope: Scope, ctx: RenderContext): unknown {
  try {
    return tag.expr.evaluate(scope, ctx.opts.filters)
  } catch (e) {
    if (e instanceof TemplateError) throw e
    const msg = e instanceof ExpressionError ? e.message : `Filter failed: ${(e as Error).message ?? e}`
    throw new TemplateError(msg, { code: 'filter_failed', part: ctx.part, tag: tag.raw })
  }
}

function tagInfo(tag: { raw: string; expr: Expression }, ctx: RenderContext): TagInfo {
  return { raw: tag.raw, expression: tag.expr.source, part: ctx.part }
}

function renderNodes(nodes: Node[], scope: Scope, spec: FormatSpec, ctx: RenderContext, out: string[]): void {
  for (const node of nodes) {
    switch (node.type) {
      case 'raw':
        out.push(node.xml)
        break
      case 'text':
        out.push(escapeXml(node.value))
        break
      case 'var': {
        let value = evaluate(node.tag, scope, ctx)
        if (value === undefined || value === null) {
          if (ctx.opts.strict) {
            throw new TemplateError(`Missing value for "${node.tag.expr.source}"`, { code: 'missing_value', part: ctx.part, tag: node.tag.raw })
          }
          value = ctx.opts.nullGetter(tagInfo(node.tag, ctx))
        }
        const text = value === undefined || value === null ? '' : ctx.opts.stringify(value)
        const custom = spec.renderScalar?.(value, text, ctx)
        if (custom !== undefined) {
          out.push(custom)
          break
        }
        let xml = escapeXml(text)
        if (ctx.opts.linebreaks) {
          if (node.br) xml = xml.replace(/\r\n|\r|\n/g, node.br)
          if (node.tab) xml = xml.replace(/\t/g, node.tab)
        }
        out.push(xml)
        break
      }
      case 'image': {
        const value = evaluate(node.tag, scope, ctx)
        if (value === undefined || value === null || value === false || value === '') {
          if (ctx.opts.strict && (value === undefined || value === null)) {
            throw new TemplateError(`Missing image for "${node.tag.expr.source}"`, { code: 'missing_value', part: ctx.part, tag: node.tag.raw })
          }
          break
        }
        if (!isImageValue(value)) {
          throw new TemplateError('Image tag expects an object like { data, width?, height?, extension? }', {
            code: 'invalid_image',
            part: ctx.part,
            tag: node.tag.raw,
          })
        }
        let img: ResolvedImage
        try {
          img = resolveImage(value)
        } catch (e) {
          throw new TemplateError((e as Error).message, { code: 'invalid_image', part: ctx.part, tag: node.tag.raw })
        }
        out.push(spec.renderImage!(img, { ...ctx, shapeXml: node.shapeXml }))
        break
      }
      case 'section':
        renderSection(node, scope, spec, ctx, out)
        break
    }
  }
}

function renderSection(node: Extract<Node, { type: 'section' }>, scope: Scope, spec: FormatSpec, ctx: RenderContext, out: string[]) {
  const value = evaluate(node.tag, scope, ctx)
  const mode: SectionMode = node.tag.mode
  const renderElse = () => node.elseBody && renderNodes(node.elseBody, scope, spec, ctx, out)

  const loop = (list: unknown[], keys?: string[]) => {
    list.forEach((item, index) => {
      const locals: Record<string, unknown> = {
        index,
        number: index + 1,
        first: index === 0,
        last: index === list.length - 1,
        length: list.length,
      }
      if (keys) locals.key = keys[index]
      renderNodes(node.body, scope.child(item, locals), spec, ctx, out)
    })
  }

  switch (mode) {
    case 'if':
      if (truthy(value)) renderNodes(node.body, scope, spec, ctx, out)
      else renderElse()
      return
    case 'unless':
    case 'inverted':
      if (!truthy(value)) renderNodes(node.body, scope, spec, ctx, out)
      else renderElse()
      return
    case 'each':
      if (Array.isArray(value) && value.length) loop(value)
      else if (value && typeof value === 'object' && !Array.isArray(value) && Object.keys(value).length) {
        const keys = Object.keys(value)
        loop(keys.map((k) => (value as Record<string, unknown>)[k]), keys)
      } else renderElse()
      return
    case 'section':
      if (Array.isArray(value)) {
        if (value.length) loop(value)
        else renderElse()
      } else if (value && typeof value === 'object') {
        renderNodes(node.body, scope.child(value), spec, ctx, out)
      } else if (value) {
        renderNodes(node.body, scope, spec, ctx, out)
      } else renderElse()
      return
  }
}

// ---------------------------------------------------------------------------------------------
// Inspection
// ---------------------------------------------------------------------------------------------

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

export function inspectNodes(nodes: Node[]): TemplateTag[] {
  const out: TemplateTag[] = []
  for (const n of nodes) {
    if (n.type === 'var') {
      out.push({ type: 'variable', path: n.tag.expr.path, expression: n.tag.expr.source, filters: n.tag.expr.filters, raw: n.tag.raw })
    } else if (n.type === 'image') {
      out.push({ type: 'image', path: n.tag.expr.path, expression: n.tag.expr.source, raw: n.tag.raw })
    } else if (n.type === 'section') {
      const tag: TemplateTag = {
        type: 'section',
        mode: n.tag.mode,
        path: n.tag.expr.path,
        expression: n.tag.expr.source,
        raw: n.tag.raw,
        children: inspectNodes(n.body),
      }
      if (n.elseBody) tag.elseChildren = inspectNodes(n.elseBody)
      out.push(tag)
    }
  }
  return out
}
