const NAMED: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }

/** Decodes XML entities (named + numeric). */
export function decodeXml(text: string): string {
  if (text.indexOf('&') === -1) return text
  return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (whole, body: string) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10)
      return Number.isFinite(code) ? String.fromCodePoint(code) : whole
    }
    return NAMED[body] ?? whole
  })
}

// Characters that are not allowed in XML 1.0 documents at all.
// eslint-disable-next-line no-control-regex
const INVALID_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g

/** Escapes text for use as XML character data. Strips characters forbidden in XML 1.0. */
export function escapeXml(text: string): string {
  return text
    .replace(INVALID_XML_CHARS, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

/** Escapes text for use inside a double-quoted XML attribute. */
export function escapeAttr(text: string): string {
  return escapeXml(text).replace(/"/g, '&quot;')
}
