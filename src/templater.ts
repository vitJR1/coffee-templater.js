import { detectFormat, type DocumentFormat, type FormatAdapter } from './formats/adapter'
import { docxAdapter } from './formats/docx/adapter'
import { pptxAdapter } from './formats/pptx/adapter'
import { xlsxAdapter } from './formats/xlsx/adapter'
import { OpcPackage, type BinaryInput } from './package'
import { compilePart, inspectNodes, renderPart, type CompiledPart, type RenderContext, type TemplateTag } from './template/engine'
import { resolveOptions, type ResolvedOptions, type TemplaterOptions } from './template/options'

const ADAPTERS: Record<DocumentFormat, FormatAdapter> = {
  docx: docxAdapter,
  xlsx: xlsxAdapter,
  pptx: pptxAdapter,
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

/** A parsed template, ready to be rendered many times. */
export class Template {
  readonly format: DocumentFormat
  private readonly pkg: OpcPackage
  private readonly compiled: CompiledPart[]
  private readonly state: Map<string, unknown>
  private readonly adapter: FormatAdapter

  constructor(
    input: BinaryInput,
    private readonly opts: ResolvedOptions,
  ) {
    this.pkg = OpcPackage.load(input)
    this.format = detectFormat(this.pkg)
    this.adapter = ADAPTERS[this.format]
    this.state = new Map<string, unknown>([['delimiterStart', opts.start]])
    this.compiled = this.adapter.templateParts(this.pkg).map((part) => {
      let xml = this.pkg.getText(part) ?? ''
      if (this.adapter.preprocess) xml = this.adapter.preprocess(this.pkg, part, xml, this.state)
      return compilePart(xml, this.adapter.spec, opts, part)
    })
  }

  /** Renders the template with `data` and returns the bytes of the resulting document. */
  render(data: unknown = {}): Uint8Array {
    const pkg = this.pkg.clone()
    const ctx: RenderContext = { opts: this.opts, part: '', pkg, state: {} }
    for (const compiled of this.compiled) {
      ctx.part = compiled.part
      let xml = renderPart(compiled, data, this.adapter.spec, ctx)
      if (this.adapter.postprocess) xml = this.adapter.postprocess(pkg, compiled.part, xml, this.state, ctx)
      pkg.setText(compiled.part, xml)
    }
    this.adapter.finalize?.(pkg, ctx)
    return pkg.toBytes()
  }

  /** Describes the tags of the template: tree, variable paths and a sample data shape. */
  inspect(): InspectResult {
    const parts = this.compiled.map((c) => ({ part: c.part, tags: inspectNodes(c.nodes) }))
    const tags = parts.flatMap((p) => p.tags)
    const variables = new Set<string>()
    const schema: Record<string, unknown> = {}
    collect(tags, '', schema, variables)
    return { format: this.format, parts, tags, variables: [...variables], schema }
  }
}

/**
 * Walks the tag tree and fills `schema` (sample data) and `out` (dotted paths).
 * Returns true when any tag inside references data (including `{{.}}`).
 */
function collect(tags: TemplateTag[], prefix: string, schema: Record<string, unknown>, out: Set<string>): boolean {
  let referenced = false
  for (const t of tags) {
    const isSelf = t.path === '.' || t.path === 'this'
    const isLocal = t.path.startsWith('@')
    const clean = t.path.replace(/^(\.\.\/)+/, '').replace(/^this\./, '')
    const full = prefix + clean
    if (t.path && !isLocal) referenced = true

    if (t.type !== 'section') {
      if (!t.path || isLocal || isSelf) continue
      out.add(full)
      if (getPath(schema, clean) === undefined) setPath(schema, clean, t.type === 'image' ? null : '')
      continue
    }

    if (!t.path || isLocal) {
      referenced = collect(t.children, prefix, schema, out) || referenced
    } else {
      const childSchema: Record<string, unknown> = {}
      const childVars = new Set<string>()
      const childRefs = collect(t.children, '', childSchema, childVars)
      // {{#each x}} is always a loop; {{#x}} is treated as a loop when its body uses data.
      const isLoop = t.mode === 'each' || (t.mode === 'section' && childRefs)
      if (isLoop) {
        if (!isSelf) setPath(schema, clean, [Object.keys(childSchema).length ? childSchema : ''])
        for (const v of childVars) out.add(`${full}[].${v}`)
        if (!childVars.size) out.add(`${full}[]`)
      } else {
        if (!isSelf) {
          out.add(full)
          if (getPath(schema, clean) === undefined) setPath(schema, clean, true)
        }
        // condition bodies share the outer scope
        collect(t.children, prefix, schema, out)
      }
    }
    if (t.elseChildren) referenced = collect(t.elseChildren, prefix, schema, out) || referenced
  }
  return referenced
}

/** Merges two sample shapes, e.g. two loops over the same array using different fields. */
function mergeShape(a: unknown, b: unknown): unknown {
  if (a === undefined) return b
  if (Array.isArray(a) && Array.isArray(b)) return [mergeShape(a[0], b[0])]
  if (a && b && typeof a === 'object' && typeof b === 'object' && !Array.isArray(a) && !Array.isArray(b)) {
    const out: Record<string, unknown> = { ...(a as Record<string, unknown>) }
    for (const [k, v] of Object.entries(b as Record<string, unknown>)) out[k] = mergeShape(out[k], v)
    return out
  }
  if (a && typeof a === 'object') return a
  return b && typeof b === 'object' ? b : a
}

function splitPath(path: string): string[] {
  return path.split(/\.|\[(?:\d+|"[^"]*"|'[^']*')\]/).filter(Boolean)
}
function getPath(obj: Record<string, unknown>, path: string): unknown {
  let cur: unknown = obj
  for (const seg of splitPath(path)) cur = cur && typeof cur === 'object' ? (cur as Record<string, unknown>)[seg] : undefined
  return cur
}
function setPath(obj: Record<string, unknown>, path: string, value: unknown) {
  const segs = splitPath(path)
  let cur = obj
  segs.forEach((seg, i) => {
    if (i === segs.length - 1) {
      cur[seg] = mergeShape(cur[seg], value)
    } else {
      if (!cur[seg] || typeof cur[seg] !== 'object' || Array.isArray(cur[seg])) cur[seg] = {}
      cur = cur[seg] as Record<string, unknown>
    }
  })
}
/** A configured templater. Create one per configuration and reuse it. */
export class Templater {
  readonly options: ResolvedOptions

  constructor(options: TemplaterOptions = {}) {
    this.options = resolveOptions(options)
  }

  /** Parses a template once; call `.render(data)` as many times as needed. */
  compile(template: BinaryInput): Template {
    return new Template(template, this.options)
  }

  /** One-shot: compile + render. */
  render(template: BinaryInput, data: unknown = {}): Uint8Array {
    return this.compile(template).render(data)
  }

  /** Lists the tags/variables used in a template. */
  inspect(template: BinaryInput): InspectResult {
    return this.compile(template).inspect()
  }
}

export function createTemplater(options?: TemplaterOptions): Templater {
  return new Templater(options)
}

/** Renders a template with default options. */
export function render(template: BinaryInput, data: unknown = {}, options?: TemplaterOptions): Uint8Array {
  return new Templater(options).render(template, data)
}
