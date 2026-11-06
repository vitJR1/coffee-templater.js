import { engineGeneration, request, type HostHandler } from './engine/core'
import { checkTimeZone, deserialize, serialize, toBytes, type BinaryInput, type SerializeOptions } from './engine/values'
import type { DocumentFormat, InspectResult, TagInfo, TemplaterOptions } from './types'

const registry =
  typeof FinalizationRegistry === 'function'
    ? new FinalizationRegistry<{ handle: number; generation: number }>(({ handle, generation }) => disposeHandle(handle, generation))
    : undefined

function disposeHandle(handle: number, generation: number) {
  try {
    if (engineGeneration() === generation) request(`{"op":"dispose","template":${handle}}`)
  } catch {
    // the engine may be busy or gone; the handle dies with the instance anyway
  }
}

/** Engine options (core spec §6.2) from the public options. */
function wireOptions(o: TemplaterOptions): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  if (o.delimiters) out.delimiters = Array.isArray(o.delimiters) ? { start: o.delimiters[0], end: o.delimiters[1] } : o.delimiters
  if (o.escapeChar !== undefined) out.escapeChar = o.escapeChar === false ? null : o.escapeChar
  if (o.syntax) out.syntax = o.syntax
  if (o.strict !== undefined) out.strict = o.strict
  if (o.linebreaks !== undefined) out.linebreaks = o.linebreaks
  if (o.filters && Object.keys(o.filters).length) out.hostFilters = Object.keys(o.filters)
  if (o.nullGetter) out.nullValue = { mode: 'host' }
  if (o.stringify) out.stringify = { mode: 'host' }
  if (o.limits) out.limits = o.limits
  return out
}

/** Runs user callbacks requested by the engine during a render. */
function hostHandler(o: TemplaterOptions): HostHandler | undefined {
  if (!o.filters && !o.nullGetter && !o.stringify) return undefined
  const ser: SerializeOptions = { timeZone: o.timeZone }
  const ok = (value: unknown) => {
    const blobs: Uint8Array[] = []
    return { json: `{"ok":true,"value":${serialize(value, blobs, ser)}}`, blobs }
  }
  return (req, blobs) => {
    try {
      switch (req.call) {
        case 'filter': {
          const fn = o.filters?.[req.name]
          if (!fn) return { json: JSON.stringify({ ok: false, message: `Unknown filter "${req.name}"` }) }
          const args = (req.args ?? []).map((a: unknown) => deserialize(a, blobs))
          return ok(fn(deserialize(req.value, blobs), ...args))
        }
        case 'nullValue':
          return ok(o.nullGetter!(req.tag as TagInfo))
        case 'stringify':
          return ok(String(o.stringify!(deserialize(req.value, blobs))))
        default:
          return { json: JSON.stringify({ ok: false, message: `Unsupported host call "${req.call}"` }) }
      }
    } catch (e) {
      return { json: JSON.stringify({ ok: false, message: e instanceof Error ? e.message : String(e) }) }
    }
  }
}

/** A parsed template, ready to be rendered many times. */
export class Template {
  readonly format: DocumentFormat
  /** Package parts that contain template text, e.g. `word/document.xml`, `word/header1.xml`. */
  readonly parts: string[]
  private readonly source: Uint8Array
  private readonly options: Record<string, unknown>
  private readonly host: HostHandler | undefined
  private readonly serializeOptions: SerializeOptions
  private handle = 0
  private generation = -1
  private inspected?: InspectResult
  private disposed = false

  constructor(input: BinaryInput, options: TemplaterOptions = {}) {
    // keep a private copy: the template is recompiled if the engine instance is re-created
    this.source = toBytes(input).slice()
    this.options = wireOptions(options)
    this.host = hostHandler(options)
    this.serializeOptions = { timeZone: options.timeZone }
    const res = this.compile()
    this.format = res.format
    this.parts = res.parts
  }

  private compile(): { format: DocumentFormat; parts: string[] } {
    const { result } = request<{ template: number; format: DocumentFormat; parts: string[] }>(
      `{"op":"compile","template":{"$blob":0},"options":${JSON.stringify(this.options)}}`,
      [this.source],
    )
    this.handle = result.template
    this.generation = engineGeneration()
    registry?.unregister(this)
    registry?.register(this, { handle: this.handle, generation: this.generation }, this)
    return result
  }

  private ensure(): number {
    if (this.disposed) throw new Error('coffee-templater: template was disposed')
    if (this.generation !== engineGeneration()) this.compile()
    return this.handle
  }

  /** Renders the template with `data` and returns the bytes of the resulting document. */
  render(data: unknown = {}): Uint8Array {
    const handle = this.ensure()
    const blobs: Uint8Array[] = []
    const json = `{"op":"render","template":${handle},"data":${serialize(data, blobs, this.serializeOptions)}}`
    return request(json, blobs, this.host).blobs[0]
  }

  /** Describes the tags of the template: tree, variable paths and a sample data shape. */
  inspect(): InspectResult {
    this.inspected ??= request<InspectResult>(`{"op":"inspect","template":${this.ensure()}}`).result
    return JSON.parse(JSON.stringify(this.inspected))
  }

  /** Frees the compiled template inside the engine. Optional: unused templates are freed by GC. */
  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    registry?.unregister(this)
    disposeHandle(this.handle, this.generation)
  }
}

/** A configured templater. Create one per configuration and reuse it. */
export class Templater {
  readonly options: Readonly<TemplaterOptions>

  constructor(options: TemplaterOptions = {}) {
    if (options.timeZone) checkTimeZone(options.timeZone)
    this.options = { ...options }
  }

  /** Parses a template once; call `.render(data)` as many times as needed. */
  compile(template: BinaryInput): Template {
    return new Template(template, this.options)
  }

  /** One-shot: compile + render. */
  render(template: BinaryInput, data: unknown = {}): Uint8Array {
    const tpl = this.compile(template)
    try {
      return tpl.render(data)
    } finally {
      tpl.dispose()
    }
  }

  /** Lists the tags/variables used in a template. */
  inspect(template: BinaryInput): InspectResult {
    const tpl = this.compile(template)
    try {
      return tpl.inspect()
    } finally {
      tpl.dispose()
    }
  }
}

export function createTemplater(options?: TemplaterOptions): Templater {
  return new Templater(options)
}

/** Renders a template with default options. */
export function render(template: BinaryInput, data: unknown = {}, options?: TemplaterOptions): Uint8Array {
  return new Templater(options).render(template, data)
}
