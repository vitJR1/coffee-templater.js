// Visual preview of a generated document, entirely in the browser.
//   Word        → docx-preview (loaded on demand)
//   Excel       → our JSON model (parse) rendered as a sheet grid
//   PowerPoint  → our JSON model for text/tables + positions and pictures read from the slide XML
import { indexToColumn, parse, parseRef } from './lib/coffee-templater.js'

const MAX_ROWS = 300
const MAX_COLS = 40
const EMU_PER_PT = 12700

let zipLib
const zip = async () => (zipLib ??= await import('./lib/zip.js'))
let urls = []

/** Renders `bytes` (docx/xlsx/pptx) into `container`. */
export async function renderPreview(bytes, format, container, lang) {
  for (const u of urls) URL.revokeObjectURL(u)
  urls = []
  container.replaceChildren()
  if (format === 'docx') return renderDocx(bytes, container)
  if (format === 'xlsx') return renderXlsx(bytes, container, lang)
  if (format === 'pptx') return renderPptx(bytes, container, lang)
  container.textContent = 'Preview is not available for this format'
}

const h = (tag, attrs = {}, ...children) => {
  const el = document.createElement(tag)
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue
    if (k === 'style' && typeof v === 'object') {
      for (const [prop, val] of Object.entries(v)) if (val !== undefined && val !== null && val !== '') el.style[prop] = val
    }
    else if (k === 'class') el.className = v
    else el.setAttribute(k, v === true ? '' : v)
  }
  for (const c of children.flat()) if (c !== null && c !== undefined && c !== false) el.append(c)
  return el
}

// ------------------------------------------------------------------------------------- Word

async function renderDocx(bytes, container) {
  const { renderAsync } = await import('./lib/docx-preview.js')
  const host = h('div', { class: 'docx-host' })
  container.append(host)
  await renderAsync(bytes, host, undefined, {
    className: 'docx',
    inWrapper: true,
    breakPages: true,
    ignoreLastRenderedPageBreak: true,
    experimental: true,
  })
  // pages are rendered at real size (e.g. 816px for Letter): shrink them to the container
  const fit = () => {
    const page = host.querySelector('section.docx')
    if (!page) return
    host.style.zoom = ''
    const zoom = Math.min(1, (container.clientWidth - 8) / host.scrollWidth)
    host.style.zoom = String(zoom)
  }
  fit()
  new ResizeObserver(fit).observe(container)
}

// ------------------------------------------------------------------------------------- Excel

function renderXlsx(bytes, container, lang) {
  const doc = parse(bytes)
  const tabs = h('div', { class: 'sheet-tabs', role: 'tablist' })
  const view = h('div', { class: 'xlsx-view' })
  const note = h('p', { class: 'preview-note' })
  container.append(view, tabs, note)

  const show = (index) => {
    tabs.querySelectorAll('button').forEach((b, i) => b.classList.toggle('active', i === index))
    const { grid, notes } = sheetTable(doc.sheets[index], lang)
    view.replaceChildren(grid)
    note.textContent = notes.join(' · ')
    note.hidden = !notes.length
  }
  for (const s of doc.sheets) tabs.append(h('button', { type: 'button', role: 'tab' }, s.name))
  tabs.querySelectorAll('button').forEach((b, i) => b.addEventListener('click', () => show(i)))
  if (doc.sheets.length) show(0)
}

function displayValue(cell) {
  const v = cell.value
  if (v === null || v === undefined) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return String(v)
}

function sheetTable(sheet, lang) {
  const cells = new Map()
  let maxRow = 0
  let maxCol = 0
  for (const row of sheet.rows) {
    let col = 0
    for (const cell of row.cells) {
      col = cell.ref ? parseRef(cell.ref).col : col + 1
      cells.set(`${row.index}:${col}`, cell)
      maxRow = Math.max(maxRow, row.index)
      maxCol = Math.max(maxCol, col)
    }
  }
  // merged ranges: the top-left cell spans, the rest are skipped
  const spans = new Map()
  const covered = new Set()
  for (const range of sheet.merges ?? []) {
    const [a, b = a] = range.split(':')
    const s = parseRef(a)
    const e = parseRef(b)
    spans.set(`${s.row}:${s.col}`, { rows: e.row - s.row + 1, cols: e.col - s.col + 1 })
    for (let r = s.row; r <= e.row; r++) for (let c = s.col; c <= e.col; c++) if (r !== s.row || c !== s.col) covered.add(`${r}:${c}`)
    maxRow = Math.max(maxRow, e.row)
    maxCol = Math.max(maxCol, e.col)
  }
  const rows = Math.min(Math.max(maxRow, 1), MAX_ROWS)
  const cols = Math.min(Math.max(maxCol, 1), MAX_COLS)

  const head = h('tr', {}, h('th', { class: 'corner' }))
  for (let c = 1; c <= cols; c++) head.append(h('th', {}, indexToColumn(c)))
  const body = []
  for (let r = 1; r <= rows; r++) {
    const tr = h('tr', {}, h('th', {}, String(r)))
    for (let c = 1; c <= cols; c++) {
      const key = `${r}:${c}`
      if (covered.has(key)) continue
      const cell = cells.get(key)
      const span = spans.get(key)
      const cls = []
      let text = ''
      let title
      if (cell) {
        text = displayValue(cell)
        if (cell.formula) {
          title = `=${cell.formula}`
          if (text === '') {
            text = `=${cell.formula}`
            cls.push('formula')
          }
        }
        if (typeof cell.value === 'number') cls.push('num')
        if (typeof cell.value === 'boolean') cls.push('bool')
        if (cell.type === 'error') cls.push('err')
      }
      tr.append(h('td', { class: cls.join(' ') || undefined, title, rowspan: span?.rows > 1 ? span.rows : undefined, colspan: span?.cols > 1 ? span.cols : undefined }, text))
    }
    body.push(tr)
  }
  const wrap = h('div', { class: 'xlsx-grid' }, h('table', {}, h('thead', {}, head), h('tbody', {}, body)))
  const notes = []
  if (maxRow > MAX_ROWS || maxCol > MAX_COLS) {
    notes.push(lang === 'ru' ? `Показано ${rows} × ${cols} из ${maxRow} × ${maxCol}` : `Showing ${rows} × ${cols} of ${maxRow} × ${maxCol}`)
  }
  if ([...cells.values()].some((c) => c.formula && displayValue(c) === '')) {
    notes.push(
      lang === 'ru'
        ? 'Формулы показаны текстом: Excel вычислит их при открытии файла.'
        : 'Formulas are shown as text: Excel calculates them when the file is opened.',
    )
  }
  return { grid: wrap, notes }
}

// ------------------------------------------------------------------------------------- PowerPoint

const attr = (xml, name) => new RegExp(`\\s${name}="([^"]*)"`).exec(xml)?.[1]

function resolvePath(from, target) {
  if (target.startsWith('/')) return target.slice(1)
  const parts = from.split('/').slice(0, -1)
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop()
    else if (seg !== '.') parts.push(seg)
  }
  return parts.join('/')
}

function relsOf(text, part) {
  const relsPath = part.replace(/([^/]+)$/, '_rels/$1.rels')
  const map = new Map()
  for (const m of text(relsPath).matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attr(m[0], 'Id')
    const target = attr(m[0], 'Target')
    if (id && target && attr(m[0], 'TargetMode') !== 'External') map.set(id, { type: attr(m[0], 'Type') ?? '', path: resolvePath(part, target) })
  }
  return map
}

/** Top-level shape XML blocks of a slide/layout, keyed by shape id. */
function shapeBlocks(xml) {
  const tree = /<p:spTree\b[\s\S]*<\/p:spTree>/.exec(xml)?.[0] ?? ''
  const blocks = []
  const re = /<(p:sp|p:pic|p:graphicFrame|p:grpSp|p:cxnSp)\b[^>]*>/g
  let m
  while ((m = re.exec(tree))) {
    // find the matching close tag (shapes of the same kind can be nested in groups)
    const name = m[1]
    let depth = 0
    const tagRe = new RegExp(`<(/?)${name}\\b[^>]*?(/?)>`, 'g')
    tagRe.lastIndex = m.index
    let t
    let end = tree.length
    while ((t = tagRe.exec(tree))) {
      if (t[2]) continue
      depth += t[1] ? -1 : 1
      if (depth === 0) {
        end = tagRe.lastIndex
        break
      }
    }
    blocks.push(tree.slice(m.index, end))
    re.lastIndex = end
  }
  return blocks
}

function boxOf(xml) {
  const xfrm = /<(?:a|p):xfrm\b[\s\S]*?<\/(?:a|p):xfrm>/.exec(xml)?.[0]
  if (!xfrm) return undefined
  const off = /<a:off\b[^>]*>/.exec(xfrm)?.[0]
  const ext = /<a:ext\b[^>]*>/.exec(xfrm)?.[0]
  if (!off || !ext) return undefined
  return { x: +attr(off, 'x'), y: +attr(off, 'y'), cx: +attr(ext, 'cx'), cy: +attr(ext, 'cy') }
}

function placeholderOf(xml) {
  const ph = /<p:ph\b[^>]*>/.exec(xml)?.[0]
  return ph ? { type: attr(ph, 'type'), idx: attr(ph, 'idx') } : undefined
}

/** Box and text anchor of a placeholder, inherited from the layout, then the master. */
function inherited(ph, sources) {
  for (const xml of sources) {
    for (const block of shapeBlocks(xml)) {
      const lph = placeholderOf(block)
      if (!lph) continue
      const sameIdx = ph.idx !== undefined && lph.idx === ph.idx
      const sameType = (ph.type ?? 'body') === (lph.type ?? 'body')
      if (sameIdx || (ph.idx === undefined && sameType) || (sameType && ph.type)) {
        const box = boxOf(block)
        if (box) return { box, anchor: attr(/<a:bodyPr\b[^>]*>/.exec(block)?.[0] ?? '', 'anchor') }
      }
    }
  }
  return {}
}

const DEFAULT_SIZE = { title: 44, ctrTitle: 44, subTitle: 32, body: 28 }

async function renderPptx(bytes, container, lang) {
  const doc = parse(bytes)
  const { unzipSync, strFromU8 } = await zip()
  const files = unzipSync(bytes)
  const cache = new Map()
  const text = (p) => {
    if (!cache.has(p)) cache.set(p, files[p] ? strFromU8(files[p]) : '')
    return cache.get(p)
  }
  const sz = /<p:sldSz\b[^>]*>/.exec(text('ppt/presentation.xml'))?.[0] ?? ''
  const W = +(attr(sz, 'cx') ?? 12192000)
  const H = +(attr(sz, 'cy') ?? 6858000)
  const pct = (v, total) => `${(v / total) * 100}%`
  const fontSize = (pt) => `${((pt * EMU_PER_PT) / W) * 100}cqw`

  const list = h('div', { class: 'slides' })
  container.append(list)

  doc.slides.forEach((slide, n) => {
    const xml = text(slide.path)
    const rels = relsOf(text, slide.path)
    const layout = [...rels.values()].find((r) => r.type.endsWith('/slideLayout'))?.path
    const master = layout ? [...relsOf(text, layout).values()].find((r) => r.type.endsWith('/slideMaster'))?.path : undefined
    const sources = [layout, master].filter(Boolean).map(text)
    const byId = new Map(shapeBlocks(xml).map((b) => [Number(attr(/<p:cNvPr\b[^>]*>/.exec(b)?.[0] ?? '', 'id')), b]))

    const canvas = h('div', { class: 'slide', style: { aspectRatio: `${W} / ${H}` } })
    for (const shape of slide.shapes) {
      const block = byId.get(shape.id) ?? ''
      const ph = placeholderOf(block)
      const own = boxOf(block)
      const inh = ph && (!own || !attr(/<a:bodyPr\b[^>]*>/.exec(block)?.[0] ?? '', 'anchor')) ? inherited(ph, sources) : {}
      const box = own ?? inh.box
      if (!box) continue
      const anchor = attr(/<a:bodyPr\b[^>]*>/.exec(block)?.[0] ?? '', 'anchor') ?? inh.anchor ?? (ph?.type === 'ctrTitle' || ph?.type === 'title' ? 'ctr' : 't')
      const style = { left: pct(box.x, W), top: pct(box.y, H), width: pct(box.cx, W), height: pct(box.cy, H) }
      const kind = ph?.type ?? (ph ? 'body' : undefined)

      if (shape.type === 'text') {
        const defaultPt = DEFAULT_SIZE[kind] ?? 18
        const centered = kind === 'ctrTitle' || kind === 'subTitle'
        const el = h('div', { class: `shape text anchor-${anchor}`, style })
        for (const p of shape.paragraphs) {
          const align = p.align ?? (centered ? 'center' : 'left')
          const firstSize = p.children.find((c) => c.type === 'run' && c.size)?.size ?? defaultPt
          const para = h('p', { style: { textAlign: align, fontSize: fontSize(firstSize), marginLeft: p.level ? `${p.level * 4}cqw` : undefined } })
          if (kind === 'body' && p.children.some((c) => c.type === 'run' && c.text.trim())) para.append('• ')
          for (const run of p.children) {
            if (run.type !== 'run') continue
            const span = h('span', {
              style: {
                fontWeight: run.bold ? '700' : undefined,
                fontStyle: run.italic ? 'italic' : undefined,
                textDecoration: [run.underline ? 'underline' : '', run.strike ? 'line-through' : ''].join(' ').trim() || undefined,
                color: run.color ? `#${run.color}` : undefined,
                fontSize: run.size ? fontSize(run.size) : undefined,
                fontFamily: run.font ? `'${run.font}', sans-serif` : undefined,
              },
            })
            run.text.split('\n').forEach((line, i) => {
              if (i) span.append(h('br'))
              span.append(line)
            })
            para.append(span)
          }
          if (!para.textContent.trim()) para.append(' ')
          el.append(para)
        }
        canvas.append(el)
      } else if (shape.type === 'table') {
        const table = h('table', { style: { fontSize: fontSize(18) } })
        shape.rows.forEach((row, ri) => {
          const tr = h('tr', { class: ri === 0 ? 'head' : undefined })
          for (const cell of row.cells) tr.append(h('td', {}, cell.paragraphs.map((p) => p.children.map((c) => (c.type === 'run' ? c.text : '')).join('')).join('\n')))
          table.append(tr)
        })
        canvas.append(h('div', { class: 'shape table', style }, table))
      } else if (shape.type === 'picture') {
        const embed = attr(/<a:blip\b[^>]*>/.exec(block)?.[0] ?? '', 'r:embed')
        const media = embed ? rels.get(embed)?.path : undefined
        const data = media ? files[media] : undefined
        if (data) {
          const url = URL.createObjectURL(new Blob([data]))
          urls.push(url)
          canvas.append(h('img', { class: 'shape picture', src: url, alt: shape.name, style }))
        } else canvas.append(h('div', { class: 'shape other', style }, shape.name))
      } else {
        canvas.append(h('div', { class: 'shape other', style }, shape.name || shape.type))
      }
    }
    const label = `${lang === 'ru' ? 'Слайд' : 'Slide'} ${n + 1}`
    list.append(
      h('figure', { class: 'slide-figure' }, canvas, h('figcaption', {}, label, slide.notes ? h('span', { class: 'slide-notes' }, ` · ${slide.notes}`) : null)),
    )
  })
}
