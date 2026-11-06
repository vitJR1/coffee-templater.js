/** A filter receives the piped value and the filter arguments: `{{ value | name: arg1, arg2 }}`. */
export type FilterFn = (value: unknown, ...args: any[]) => unknown

const isEmpty = (v: unknown) =>
  v === undefined || v === null || v === '' || (Array.isArray(v) && v.length === 0)

const toNumber = (v: unknown): number => (typeof v === 'number' ? v : typeof v === 'string' ? parseFloat(v.replace(',', '.')) : Number(v))

const toDate = (v: unknown): Date | undefined => {
  if (v instanceof Date) return v
  if (typeof v === 'number' || typeof v === 'string') {
    const d = new Date(v)
    return isNaN(d.getTime()) ? undefined : d
  }
  return undefined
}

const pad = (n: number, len = 2) => String(n).padStart(len, '0')

/**
 * Formats a date with tokens: YYYY YY MM M DD D HH H mm m ss s.
 * Text in [brackets] is printed literally. Uses local time unless `utc` is true.
 */
export function formatDate(date: Date, format: string, utc = false): string {
  const parts = {
    Y: utc ? date.getUTCFullYear() : date.getFullYear(),
    M: (utc ? date.getUTCMonth() : date.getMonth()) + 1,
    D: utc ? date.getUTCDate() : date.getDate(),
    H: utc ? date.getUTCHours() : date.getHours(),
    m: utc ? date.getUTCMinutes() : date.getMinutes(),
    s: utc ? date.getUTCSeconds() : date.getSeconds(),
  }
  return format.replace(/\[([^\]]*)]|YYYY|YY|MM|M|DD|D|HH|H|mm|m|ss|s/g, (tok, literal?: string) => {
    if (literal !== undefined) return literal
    switch (tok) {
      case 'YYYY': return String(parts.Y)
      case 'YY': return pad(parts.Y % 100)
      case 'MM': return pad(parts.M)
      case 'M': return String(parts.M)
      case 'DD': return pad(parts.D)
      case 'D': return String(parts.D)
      case 'HH': return pad(parts.H)
      case 'H': return String(parts.H)
      case 'mm': return pad(parts.m)
      case 'm': return String(parts.m)
      case 'ss': return pad(parts.s)
      default: return String(parts.s)
    }
  })
}

/**
 * Formats a number: fixed decimals, thousands separator and decimal separator.
 * `{{ 1234567.891 | number: 2, " ", "," }}` → `1 234 567,89`
 */
export function formatNumber(value: number, decimals = 0, thousands = ' ', decimal = ','): string {
  if (!Number.isFinite(value)) return ''
  const fixed = Math.abs(value).toFixed(decimals)
  const [int, frac] = fixed.split('.')
  const grouped = int.replace(/\B(?=(\d{3})+(?!\d))/g, thousands)
  return (value < 0 && Number(fixed) !== 0 ? '-' : '') + grouped + (frac ? decimal + frac : '')
}

export const builtinFilters: Record<string, FilterFn> = {
  // strings
  upper: (v) => (v == null ? v : String(v).toUpperCase()),
  lower: (v) => (v == null ? v : String(v).toLowerCase()),
  capitalize: (v) => (v == null ? v : String(v).charAt(0).toUpperCase() + String(v).slice(1)),
  trim: (v) => (v == null ? v : String(v).trim()),
  replace: (v, search: string, replacement = '') => (v == null ? v : String(v).split(String(search)).join(String(replacement))),
  truncate: (v, length = 50, suffix = '…') => {
    if (v == null) return v
    const s = String(v)
    return s.length > length ? s.slice(0, Math.max(0, length - String(suffix).length)) + suffix : s
  },
  prepend: (v, prefix: string) => (isEmpty(v) ? v : String(prefix) + String(v)),
  append: (v, suffix: string) => (isEmpty(v) ? v : String(v) + String(suffix)),

  // fallbacks
  default: (v, fallback = '') => (isEmpty(v) ? fallback : v),

  // numbers / dates
  number: (v, decimals = 0, thousands = ' ', decimal = ',') => (isEmpty(v) ? v : formatNumber(toNumber(v), decimals, thousands, decimal)),
  fixed: (v, decimals = 2) => (isEmpty(v) ? v : toNumber(v).toFixed(decimals)),
  round: (v, decimals = 0) => (isEmpty(v) ? v : Math.round(toNumber(v) * 10 ** decimals) / 10 ** decimals),
  currency: (v, currency = 'USD', locale?: string, decimals?: number) =>
    isEmpty(v)
      ? v
      : new Intl.NumberFormat(locale, {
          style: 'currency',
          currency,
          ...(decimals !== undefined ? { minimumFractionDigits: decimals, maximumFractionDigits: decimals } : {}),
        }).format(toNumber(v)),
  percent: (v, decimals = 0) => (isEmpty(v) ? v : `${(toNumber(v) * 100).toFixed(decimals)}%`),
  date: (v, format = 'DD.MM.YYYY', utc = false) => {
    const d = toDate(v)
    return d ? formatDate(d, String(format), Boolean(utc)) : v
  },

  // collections
  join: (v, sep = ', ') => (Array.isArray(v) ? v.join(String(sep)) : v),
  length: (v) => (v == null ? 0 : Array.isArray(v) || typeof v === 'string' ? v.length : typeof v === 'object' ? Object.keys(v).length : 0),
  first: (v) => (Array.isArray(v) ? v[0] : v),
  last: (v) => (Array.isArray(v) ? v[v.length - 1] : v),
  reverse: (v) => (Array.isArray(v) ? [...v].reverse() : typeof v === 'string' ? [...v].reverse().join('') : v),
  sort: (v, key?: string, dir = 'asc') => {
    if (!Array.isArray(v)) return v
    const sign = String(dir).toLowerCase() === 'desc' ? -1 : 1
    const pick = (x: unknown) => (key ? (x as Record<string, unknown>)?.[key] : x) as never
    return [...v].sort((a, b) => (pick(a) > pick(b) ? sign : pick(a) < pick(b) ? -sign : 0))
  },
  /** `{{#each items | where: "active"}}` or `{{#each items | where: "status", "paid"}}` */
  where: (v, key: string, ...value: unknown[]) =>
    Array.isArray(v) ? v.filter((x) => (value.length ? x?.[key] === value[0] : Boolean(x?.[key]))) : v,
  slice: (v, start = 0, end?: number) => (Array.isArray(v) || typeof v === 'string' ? v.slice(start, end) : v),
  json: (v, indent = 0) => JSON.stringify(v, null, indent),

  // comparison / logic — handy in conditions: {{#if total | gt: 1000}}
  eq: (v, other) => v === other || (v != null && other != null && String(v) === String(other)),
  ne: (v, other) => !(v === other || (v != null && other != null && String(v) === String(other))),
  gt: (v, other) => toNumber(v) > toNumber(other),
  gte: (v, other) => toNumber(v) >= toNumber(other),
  lt: (v, other) => toNumber(v) < toNumber(other),
  lte: (v, other) => toNumber(v) <= toNumber(other),
  not: (v) => isEmpty(v) || !v,
  and: (v, ...others) => [v, ...others].every((x) => !isEmpty(x) && Boolean(x)),
  or: (v, ...others) => [v, ...others].find((x) => !isEmpty(x) && Boolean(x)) ?? false,
  contains: (v, item) => (Array.isArray(v) ? v.includes(item) : typeof v === 'string' ? v.includes(String(item)) : false),
  empty: (v) => isEmpty(v),
  /** `{{ paid | yesno: "Оплачено", "Не оплачено" }}` */
  yesno: (v, yes = 'yes', no = 'no') => (isEmpty(v) || !v ? no : yes),
}
