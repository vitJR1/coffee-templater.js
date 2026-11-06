/**
 * Conversion between JavaScript values and the engine's JSON notation (core spec §5):
 *
 *   Date                         ↔ {"$date": "<ISO 8601>"}
 *   Uint8Array / ArrayBuffer     ↔ {"$blob": <index into the frame's blobs>}
 *   bigint                       → number lexeme (no precision loss)
 *   {"$date": ...} in user data  → {"$literal": ...} (so it is not mistaken for a tag)
 */

export type BinaryInput = Uint8Array | ArrayBuffer | ArrayBufferView

const TAGS = new Set(['$date', '$decimal', '$blob', '$base64', '$literal'])

export function toBytes(input: BinaryInput): Uint8Array {
  if (input instanceof Uint8Array) return input
  if (input instanceof ArrayBuffer) return new Uint8Array(input)
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength)
  throw new TypeError('Expected Uint8Array, ArrayBuffer, Buffer or another ArrayBufferView')
}

function isBinary(v: unknown): v is BinaryInput {
  return v instanceof ArrayBuffer || ArrayBuffer.isView(v)
}

export interface SerializeOptions {
  /** Send `Date`s as naive wall-clock time in this zone (`'local'` or an IANA name) instead of UTC. */
  timeZone?: string
}

const pad = (n: number, len = 2) => String(n).padStart(len, '0')
const formatters = new Map<string, Intl.DateTimeFormat>()

/** Validates a time zone option; throws RangeError for unknown zones. */
export function checkTimeZone(timeZone: string): void {
  if (timeZone !== 'local') formatter(timeZone)
}

function formatter(timeZone: string): Intl.DateTimeFormat {
  let f = formatters.get(timeZone)
  if (!f) {
    f = new Intl.DateTimeFormat('en-US', {
      timeZone,
      year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
    })
    formatters.set(timeZone, f)
  }
  return f
}

/** `2024-01-31T00:00:00.000` — the wall-clock time of `d` in `timeZone`, without an offset. */
export function wallClock(d: Date, timeZone: string): string {
  let y: number, mo: number, da: number, h: number, mi: number, s: number
  if (timeZone === 'local') {
    ;[y, mo, da, h, mi, s] = [d.getFullYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()]
  } else {
    const parts = Object.fromEntries(formatter(timeZone).formatToParts(d).map((p) => [p.type, p.value]))
    ;[y, mo, da, h, mi, s] = [parts.year, parts.month, parts.day, parts.hour, parts.minute, parts.second].map(Number)
  }
  return `${pad(y, 4)}-${pad(mo)}-${pad(da)}T${pad(h)}:${pad(mi)}:${pad(s)}.${pad(d.getMilliseconds(), 3)}`
}

/**
 * Serializes a value to engine JSON text, collecting binary values into `blobs`.
 * Follows `JSON.stringify` semantics (toJSON, dropped functions/undefined) plus the tags above.
 */
export function serialize(value: unknown, blobs: Uint8Array[], options: SerializeOptions = {}): string {
  const stack = new Set<object>()
  const date = (d: Date) => (options.timeZone ? wallClock(d, options.timeZone) : d.toISOString())

  const write = (v: unknown, inArray: boolean): string | undefined => {
    if (v === null) return 'null'
    if (v instanceof Date) return isNaN(v.getTime()) ? 'null' : `{"$date":"${date(v)}"}`
    if (isBinary(v)) {
      blobs.push(toBytes(v))
      return `{"$blob":${blobs.length - 1}}`
    }
    if (typeof v === 'object' && typeof (v as { toJSON?: unknown }).toJSON === 'function') {
      v = (v as { toJSON(): unknown }).toJSON()
      if (v === null || typeof v !== 'object') return write(v, inArray)
    }
    switch (typeof v) {
      case 'string':
        return JSON.stringify(v)
      case 'number':
        return Number.isFinite(v) ? JSON.stringify(v) : 'null'
      case 'boolean':
        return v ? 'true' : 'false'
      case 'bigint':
        return v.toString()
      case 'undefined':
      case 'function':
      case 'symbol':
        return inArray ? 'null' : undefined
    }
    const obj = v as object
    if (stack.has(obj)) throw new TypeError('coffee-templater: data contains a circular reference')
    stack.add(obj)
    try {
      if (Array.isArray(obj) || obj instanceof Set) {
        return `[${[...(obj as Iterable<unknown>)].map((x) => write(x, true)).join(',')}]`
      }
      const entries: Array<[string, unknown]> =
        obj instanceof Map ? [...obj.entries()].map(([k, x]) => [String(k), x]) : Object.entries(obj)
      const parts: string[] = []
      const written: string[] = []
      for (const [k, x] of entries) {
        const s = write(x, false)
        if (s === undefined) continue
        parts.push(`${JSON.stringify(k)}:${s}`)
        written.push(k)
      }
      const json = `{${parts.join(',')}}`
      // A user object that looks like a tag must be passed literally.
      return written.length === 1 && TAGS.has(written[0]) ? `{"$literal":${json}}` : json
    } finally {
      stack.delete(obj)
    }
  }

  return write(value, false) ?? 'null'
}

/** Converts engine JSON (already parsed) back into JavaScript values. */
export function deserialize(value: unknown, blobs: readonly Uint8Array[]): unknown {
  if (Array.isArray(value)) return value.map((v) => deserialize(v, blobs))
  if (!value || typeof value !== 'object') return value
  const obj = value as Record<string, unknown>
  const keys = Object.keys(obj)
  if (keys.length === 1 && TAGS.has(keys[0])) {
    const tag = keys[0]
    const v = obj[tag]
    switch (tag) {
      case '$date':
        return new Date(String(v))
      case '$blob':
        return blobs[Number(v)]
      case '$base64':
        return base64ToBytes(String(v))
      case '$decimal':
        return String(v)
      case '$literal':
        return v
    }
  }
  const out: Record<string, unknown> = {}
  for (const k of keys) out[k] = deserialize(obj[k], blobs)
  return out
}

export function base64ToBytes(b64: string): Uint8Array {
  const clean = b64.replace(/^data:[^;]+;base64,/, '').replace(/\s+/g, '')
  const B = (globalThis as { Buffer?: { from(s: string, e: string): Uint8Array } }).Buffer
  if (B) {
    const buf = B.from(clean, 'base64')
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength)
  }
  const bin = atob(clean)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}
