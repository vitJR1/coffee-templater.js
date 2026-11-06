import { toUint8Array, type BinaryInput } from './package'

/** Value accepted by image tags (`{{%logo}}` by default). */
export interface ImageValue {
  /** Raw bytes or a base64 / data-URL string. */
  data: BinaryInput | string
  /** png | jpeg | gif | bmp | svg | ... Detected from bytes when omitted. */
  extension?: string
  /** Width in pixels (96 dpi). Detected from the image when omitted. */
  width?: number
  /** Height in pixels (96 dpi). Detected from the image when omitted. */
  height?: number
  /** Alternative text. */
  alt?: string
}

export interface ResolvedImage {
  data: Uint8Array
  extension: string
  width: number
  height: number
  alt: string
}

export const EMU_PER_PX = 9525

export function isImageValue(v: unknown): v is ImageValue {
  return typeof v === 'object' && v !== null && 'data' in v
}

export function resolveImage(value: ImageValue): ResolvedImage {
  const data = toUint8Array(value.data)
  const info = sniffImage(data)
  const extension = (value.extension ?? info?.extension ?? '').replace(/^\./, '').toLowerCase()
  if (!extension) throw new Error('Cannot detect image format; pass `extension`')
  const natural = info && info.width > 0 && info.height > 0 ? info : undefined
  let width = value.width
  let height = value.height
  if (width === undefined && height === undefined) {
    width = natural?.width ?? 100
    height = natural?.height ?? 100
  } else if (width === undefined) {
    width = natural ? (height! * natural.width) / natural.height : height!
  } else if (height === undefined) {
    height = natural ? (width * natural.height) / natural.width : width
  }
  return { data, extension, width: width!, height: height!, alt: value.alt ?? '' }
}

/** Detects format and pixel size of PNG, JPEG, GIF and BMP images. */
export function sniffImage(d: Uint8Array): { extension: string; width: number; height: number } | undefined {
  const u32 = (o: number) => ((d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3]) >>> 0
  const u16be = (o: number) => (d[o] << 8) | d[o + 1]
  const u16le = (o: number) => d[o] | (d[o + 1] << 8)
  if (d.length > 24 && d[0] === 0x89 && d[1] === 0x50 && d[2] === 0x4e && d[3] === 0x47) {
    return { extension: 'png', width: u32(16), height: u32(20) }
  }
  if (d.length > 10 && d[0] === 0x47 && d[1] === 0x49 && d[2] === 0x46) {
    return { extension: 'gif', width: u16le(6), height: u16le(8) }
  }
  if (d.length > 26 && d[0] === 0x42 && d[1] === 0x4d) {
    const w = d[18] | (d[19] << 8) | (d[20] << 16) | (d[21] << 24)
    const h = d[22] | (d[23] << 8) | (d[24] << 16) | (d[25] << 24)
    return { extension: 'bmp', width: Math.abs(w), height: Math.abs(h) }
  }
  if (d.length > 4 && d[0] === 0xff && d[1] === 0xd8) {
    let o = 2
    while (o + 9 < d.length) {
      if (d[o] !== 0xff) {
        o++
        continue
      }
      const marker = d[o + 1]
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { extension: 'jpeg', width: u16be(o + 7), height: u16be(o + 5) }
      }
      o += 2 + u16be(o + 2)
    }
    return { extension: 'jpeg', width: 100, height: 100 }
  }
  const head = new TextDecoder().decode(d.subarray(0, Math.min(d.length, 512)))
  if (/<svg[\s>]/i.test(head)) {
    const w = /\bwidth="([\d.]+)/.exec(head)
    const h = /\bheight="([\d.]+)/.exec(head)
    return { extension: 'svg', width: w ? parseFloat(w[1]) : 100, height: h ? parseFloat(h[1]) : 100 }
  }
  return undefined
}
