/**
 * Expression language used inside tags:
 *
 *   path            client.name, items[0].title, this, ., @index, ../title
 *   literal         "text", 'text', «text», “text”, 42, -1.5, true, false, null
 *   filters         value | upper | default: "—" | date: "DD.MM.YYYY"
 */
import type { FilterFn } from './filters'

export class Scope {
  constructor(
    readonly data: unknown,
    readonly parent?: Scope,
    readonly locals: Record<string, unknown> = {},
  ) {}

  child(data: unknown, locals: Record<string, unknown> = {}): Scope {
    return new Scope(data, this, locals)
  }
}

type Segment = string | number

type Operand =
  | { kind: 'literal'; value: unknown }
  | { kind: 'path'; up: number; self: boolean; local?: string; segments: Segment[]; source: string }

interface FilterCall {
  name: string
  args: Operand[]
}

export interface Expression {
  source: string
  /** Main operand path, e.g. `client.name` for `client.name | upper`. Empty for literals. */
  path: string
  /** Every path referenced by the expression (operand and filter arguments). */
  paths: string[]
  filters: string[]
  evaluate(scope: Scope, filters: Record<string, FilterFn>): unknown
}

export class ExpressionError extends Error {}

type Tok = { t: 'str'; v: string } | { t: 'num'; v: number } | { t: 'word'; v: string } | { t: 'punct'; v: string }

const QUOTES: Record<string, string> = { '"': '"', "'": "'", '“': '”', '„': '“', '«': '»', '‘': '’', '”': '”' }
const WORD_CHAR = /[\p{L}\p{N}_$@.[\]\/\-]/u

function lex(src: string): Tok[] {
  const toks: Tok[] = []
  let i = 0
  while (i < src.length) {
    const ch = src[i]
    if (/\s/.test(ch)) {
      i++
      continue
    }
    if (ch === '|' || ch === ':' || ch === ',') {
      toks.push({ t: 'punct', v: ch })
      i++
      continue
    }
    if (QUOTES[ch]) {
      const close = QUOTES[ch]
      let j = i + 1
      let value = ''
      while (j < src.length && src[j] !== close && !(close === '”' && src[j] === '"')) {
        if (src[j] === '\\' && j + 1 < src.length) {
          value += src[j + 1]
          j += 2
          continue
        }
        value += src[j++]
      }
      if (j >= src.length) throw new ExpressionError(`Unterminated string in "${src}"`)
      toks.push({ t: 'str', v: value })
      i = j + 1
      continue
    }
    const num = /^-?\d+(\.\d+)?(?![\p{L}_$])/u.exec(src.slice(i))
    if (num) {
      toks.push({ t: 'num', v: parseFloat(num[0]) })
      i += num[0].length
      continue
    }
    // A path word; brackets may contain quoted keys with spaces: a["some key"]
    let j = i
    let bracket = false
    let quote = ''
    while (j < src.length) {
      const c = src[j]
      if (quote) {
        if (c === quote) quote = ''
      } else if (bracket && (c === '"' || c === "'")) {
        quote = c
      } else if (c === '[') {
        bracket = true
      } else if (c === ']') {
        bracket = false
      } else if (!bracket && !WORD_CHAR.test(c)) {
        break
      }
      j++
    }
    if (j === i) throw new ExpressionError(`Unexpected character "${ch}" in "${src}"`)
    toks.push({ t: 'word', v: src.slice(i, j) })
    i = j
  }
  return toks
}

function parsePath(word: string): Operand {
  if (word === 'true') return { kind: 'literal', value: true }
  if (word === 'false') return { kind: 'literal', value: false }
  if (word === 'null') return { kind: 'literal', value: null }
  if (word === 'undefined') return { kind: 'literal', value: undefined }

  let rest = word
  let up = 0
  while (rest.startsWith('../')) {
    up++
    rest = rest.slice(3)
  }
  if (rest === '..') return { kind: 'path', up: up + 1, self: true, segments: [], source: word }
  if (rest === '.' || rest === 'this') return { kind: 'path', up, self: true, segments: [], source: word }

  const segments: Segment[] = []
  let local: string | undefined
  let self = false
  if (rest.startsWith('this.')) {
    self = true
    rest = rest.slice(5)
  }
  const re = /([^.[\]]+)|\[\s*(?:(-?\d+)|"([^"]*)"|'([^']*)')\s*\]|(\.)/g
  let m: RegExpExecArray | null
  let pos = 0
  while ((m = re.exec(rest))) {
    if (m.index !== pos) throw new ExpressionError(`Invalid path "${word}"`)
    pos = re.lastIndex
    if (m[5]) continue
    if (m[1] !== undefined) segments.push(m[1])
    else if (m[2] !== undefined) segments.push(parseInt(m[2], 10))
    else segments.push(m[3] ?? m[4] ?? '')
  }
  if (pos !== rest.length || segments.length === 0) throw new ExpressionError(`Invalid path "${word}"`)
  if (typeof segments[0] === 'string' && segments[0].startsWith('@')) {
    local = segments.shift() as string
    local = local.slice(1)
  }
  return { kind: 'path', up, self, local, segments, source: word }
}

function parseOperand(tok: Tok | undefined, src: string): Operand {
  if (!tok) throw new ExpressionError(`Expected a value in "${src}"`)
  if (tok.t === 'str' || tok.t === 'num') return { kind: 'literal', value: tok.v }
  if (tok.t === 'word') return parsePath(tok.v)
  throw new ExpressionError(`Unexpected "${tok.v}" in "${src}"`)
}

function get(obj: unknown, key: Segment): unknown {
  if (obj === null || obj === undefined) return undefined
  if (typeof key === 'number' && Array.isArray(obj) && key < 0) return obj[obj.length + key]
  if (key === 'length' && (Array.isArray(obj) || typeof obj === 'string')) return obj.length
  if (typeof obj !== 'object' && typeof obj !== 'function') return undefined
  // Only own properties — never walk into prototypes (`constructor`, `__proto__`, ...).
  return Object.prototype.hasOwnProperty.call(obj, key) ? (obj as Record<string, unknown>)[key as string] : undefined
}

function hasOwn(obj: unknown, key: Segment): boolean {
  return obj !== null && typeof obj === 'object' && Object.prototype.hasOwnProperty.call(obj, key)
}

function resolveOperand(op: Operand, scope: Scope): unknown {
  if (op.kind === 'literal') return op.value
  let s: Scope | undefined = scope
  for (let i = 0; i < op.up && s; i++) s = s.parent
  if (!s) return undefined

  if (op.local !== undefined) {
    let cur: Scope | undefined = s
    while (cur && !(op.local in cur.locals)) cur = cur.parent
    let value = cur?.locals[op.local]
    for (const seg of op.segments) value = get(value, seg)
    return value
  }
  if (op.self) {
    let value: unknown = s.data
    for (const seg of op.segments) value = get(value, seg)
    return value
  }
  // Lookup the first segment through the scope chain (like Mustache), then walk.
  const [first, ...rest] = op.segments
  let cur: Scope | undefined = s
  while (cur && !hasOwn(cur.data, first)) cur = cur.parent
  if (!cur) return undefined
  let value = get(cur.data, first)
  for (const seg of rest) value = get(value, seg)
  return value
}

const cache = new Map<string, Expression>()

export function compileExpression(source: string): Expression {
  const cached = cache.get(source)
  if (cached) return cached

  const toks = lex(source)
  if (toks.length === 0) throw new ExpressionError('Empty expression')
  let i = 0
  const operand = parseOperand(toks[i++], source)
  const calls: FilterCall[] = []
  while (i < toks.length) {
    const pipe = toks[i++]
    if (pipe.t !== 'punct' || pipe.v !== '|') throw new ExpressionError(`Expected "|" in "${source}"`)
    const name = toks[i++]
    if (!name || name.t !== 'word') throw new ExpressionError(`Expected filter name after "|" in "${source}"`)
    const call: FilterCall = { name: name.v, args: [] }
    if (toks[i]?.t === 'punct' && toks[i].v === ':') {
      i++
      call.args.push(parseOperand(toks[i++], source))
      while (toks[i]?.t === 'punct' && toks[i].v === ',') {
        i++
        call.args.push(parseOperand(toks[i++], source))
      }
    }
    calls.push(call)
  }

  const pathOf = (op: Operand) => (op.kind === 'path' ? op.source : '')
  const expr: Expression = {
    source,
    path: pathOf(operand),
    paths: [operand, ...calls.flatMap((c) => c.args)].map(pathOf).filter(Boolean),
    filters: calls.map((c) => c.name),
    evaluate(scope, filters) {
      let value = resolveOperand(operand, scope)
      for (const call of calls) {
        const fn = filters[call.name]
        if (!fn) throw new ExpressionError(`Unknown filter "${call.name}"`)
        value = fn(value, ...call.args.map((a) => resolveOperand(a, scope)))
      }
      return value
    },
  }
  if (cache.size > 5000) cache.clear()
  cache.set(source, expr)
  return expr
}
