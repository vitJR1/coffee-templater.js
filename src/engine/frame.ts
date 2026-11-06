/**
 * Wire frames of the engine ABI (core spec §4):
 *
 *   "CTF1" | u32 json_len | JSON (UTF-8) | u32 blob_count | (u32 blob_len | bytes)*
 *
 * All integers are little-endian u32.
 */

const MAGIC = [0x43, 0x54, 0x46, 0x31] // "CTF1"
const encoder = new TextEncoder()
const decoder = new TextDecoder()

export interface Frame<T = any> {
  json: T
  blobs: Uint8Array[]
}

export function encodeFrame(json: string, blobs: readonly Uint8Array[] = []): Uint8Array {
  const body = encoder.encode(json)
  let size = 12 + body.length
  for (const b of blobs) size += 4 + b.length
  const out = new Uint8Array(size)
  const view = new DataView(out.buffer)
  out.set(MAGIC, 0)
  view.setUint32(4, body.length, true)
  out.set(body, 8)
  let p = 8 + body.length
  view.setUint32(p, blobs.length, true)
  p += 4
  for (const b of blobs) {
    view.setUint32(p, b.length, true)
    out.set(b, p + 4)
    p += 4 + b.length
  }
  return out
}

/** Decodes a frame. Blobs are copied out of `buf` (which may be a view into WASM memory). */
export function decodeFrame(buf: Uint8Array): Frame {
  if (buf.length < 12 || MAGIC.some((m, i) => buf[i] !== m)) throw new Error('coffee-engine: invalid frame')
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength)
  const jsonLen = view.getUint32(4, true)
  const json = JSON.parse(decoder.decode(buf.subarray(8, 8 + jsonLen)))
  let p = 8 + jsonLen
  const count = view.getUint32(p, true)
  p += 4
  const blobs: Uint8Array[] = []
  for (let i = 0; i < count; i++) {
    const len = view.getUint32(p, true)
    blobs.push(buf.slice(p + 4, p + 4 + len))
    p += 4 + len
  }
  return { json, blobs }
}
