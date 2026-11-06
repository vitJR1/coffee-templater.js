// Demo + small page behaviours. Runs the real library (lib/coffee-templater.js, built by
// scripts/build-site.mjs) entirely in the browser.
import { createTemplater, engineInfo, extractText, initEngine, TemplateError } from './lib/coffee-templater.js'

// ------------------------------------------------------------------------------ language & copy

const html = document.documentElement
document.querySelectorAll('[data-set-lang]').forEach((btn) =>
  btn.addEventListener('click', () => {
    const lang = btn.getAttribute('data-set-lang')
    html.setAttribute('data-lang', lang)
    html.setAttribute('lang', lang)
    localStorage.setItem('lang', lang)
  }),
)
// theme: explicit choice is saved; otherwise follow the system setting live
document.getElementById('theme-toggle').addEventListener('click', () => {
  const next = html.getAttribute('data-theme') === 'dark' ? 'light' : 'dark'
  html.setAttribute('data-theme', next)
  localStorage.setItem('theme', next)
})
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
  if (!localStorage.getItem('theme')) html.setAttribute('data-theme', e.matches ? 'dark' : 'light')
})

const t = (en, ru) => (html.getAttribute('data-lang') === 'ru' ? ru : en)

document.querySelectorAll('[data-copy]').forEach((btn) =>
  btn.addEventListener('click', async () => {
    const text = document.querySelector(btn.getAttribute('data-copy')).textContent
    try {
      await navigator.clipboard.writeText(text)
      const old = btn.innerHTML
      btn.textContent = t('Copied', 'Скопировано')
      setTimeout(() => (btn.innerHTML = old), 1400)
    } catch {
      /* clipboard not available */
    }
  }),
)

// ------------------------------------------------------------------------------ demo

const SAMPLES = {
  invoice: {
    file: 'samples/invoice.docx',
    data: {
      company: 'Coffee & Co',
      number: 42,
      client: { name: 'Ivan Petrov' },
      items: [
        { name: 'Espresso beans, 1 kg', qty: 2, price: 24.5 },
        { name: 'Grinder', qty: 1, price: 129 },
        { name: 'Paper filters', qty: 3, price: 4.2 },
      ],
      paid: false,
      notes: 'Delivery on Monday.\nCall before arrival.',
      total: 190.6,
    },
  },
  report: {
    file: 'samples/report.xlsx',
    data: {
      title: 'Q3 2026',
      items: [
        { name: 'Espresso', qty: 120, price: 2.5 },
        { name: 'Cappuccino', qty: 95, price: 3.2 },
        { name: 'Latte', qty: 140, price: 3.5 },
        { name: 'Flat white', qty: 60, price: 3.4 },
      ],
      author: 'Analytics team',
    },
  },
  deck: {
    file: 'samples/deck.pptx',
    data: {
      title: 'coffee-templater',
      subtitle: 'Fill Office templates from data\nFree and open source',
      points: ['Word, Excel, PowerPoint', 'Loops, conditions, filters', 'Images and JSON model'],
      rows: [
        { name: 'Formats', value: 3 },
        { name: 'Built-in filters', value: 40 },
        { name: 'Price', value: 'Free' },
      ],
      speaker: 'Your name',
    },
  },
}

const MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
}

const $ = (id) => document.getElementById(id)
const els = {
  badge: $('engine-badge'),
  name: $('template-name'),
  file: $('file'),
  start: $('d-start'),
  end: $('d-end'),
  vars: $('vars'),
  data: $('data'),
  render: $('render'),
  download: $('download'),
  error: $('error'),
  preview: $('preview'),
  previewText: $('preview-text'),
}

const state = { bytes: null, name: '', images: [], logo: null, url: null }

function templater() {
  return createTemplater({ delimiters: [els.start.value || '{{', els.end.value || '}}'] })
}

function showError(e) {
  els.error.hidden = false
  if (e instanceof TemplateError) {
    const items = e.errors
      .map((x) => `<li><code>${esc(x.code)}</code> ${esc(x.message)}${x.tag ? ` — <code>${esc(x.tag)}</code>` : ''}${x.part ? ` <small>(${esc(x.part)})</small>` : ''}</li>`)
      .join('')
    els.error.innerHTML = `<strong>${t('Template error', 'Ошибка шаблона')}</strong><ul>${items}</ul>`
  } else {
    els.error.innerHTML = `<strong>${t('Error', 'Ошибка')}:</strong> ${esc(e?.message ?? String(e))}`
  }
}

function clearResult() {
  els.error.hidden = true
  els.preview.hidden = true
  els.download.hidden = true
  if (state.url) URL.revokeObjectURL(state.url)
  state.url = null
}

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c])

/** Loads a template: inspects its tags, lists variables and fills the JSON box. */
function loadTemplate(bytes, name, sampleData) {
  clearResult()
  state.bytes = bytes
  state.name = name
  els.name.textContent = name
  els.vars.innerHTML = ''
  els.render.disabled = true
  try {
    const info = templater().inspect(bytes)
    state.images = collectImages(info.tags)
    els.vars.innerHTML = info.variables.map((v) => `<code>${esc(v)}</code>`).join('')
    const data = sampleData ? structuredClone(sampleData) : info.schema
    // images cannot be typed as JSON: offer a generated sample picture instead
    for (const path of state.images) if (path in data && data[path] === null) delete data[path]
    els.data.value = JSON.stringify(data, null, 2)
    renderLogoOption()
    els.render.disabled = false
  } catch (e) {
    showError(e)
  }
}

function collectImages(tags, out = []) {
  for (const tag of tags) {
    if (tag.type === 'image' && tag.path && !out.includes(tag.path)) out.push(tag.path)
    if (tag.type === 'section') {
      collectImages(tag.children, out)
      if (tag.elseChildren) collectImages(tag.elseChildren, out)
    }
  }
  return out
}

function renderLogoOption() {
  document.getElementById('logo-option')?.remove()
  if (!state.images.length) return
  const label = document.createElement('label')
  label.id = 'logo-option'
  label.className = 'hint logo-option'
  label.innerHTML =
    '<input type="checkbox" id="logo-check" checked> ' +
    '<span data-l="en">Insert a sample picture into</span><span data-l="ru">Подставить пример картинки в</span> ' +
    state.images.map((p) => `<code>{{%${esc(p)}}}</code>`).join(', ')
  els.data.insertAdjacentElement('afterend', label)
}

/** A generated PNG logo (so the demo can show image tags). */
async function sampleLogo() {
  if (state.logo) return state.logo
  const c = document.createElement('canvas')
  c.width = 360
  c.height = 120
  const g = c.getContext('2d')
  g.fillStyle = '#b5652a'
  g.beginPath()
  g.roundRect(0, 0, 360, 120, 24)
  g.fill()
  g.font = '64px serif'
  g.fillText('☕', 22, 84)
  g.fillStyle = '#fff'
  g.font = 'bold 34px system-ui, sans-serif'
  g.fillText('coffee', 112, 58)
  g.font = '26px system-ui, sans-serif'
  g.fillText('templater', 112, 92)
  const blob = await new Promise((r) => c.toBlob(r, 'image/png'))
  state.logo = new Uint8Array(await blob.arrayBuffer())
  return state.logo
}

async function render() {
  clearResult()
  let data
  try {
    data = JSON.parse(els.data.value || '{}')
  } catch (e) {
    showError(new Error(t('Data is not valid JSON: ', 'Данные — не корректный JSON: ') + e.message))
    return
  }
  try {
    if (document.getElementById('logo-check')?.checked) {
      const logo = await sampleLogo()
      for (const path of state.images) setPath(data, path, { data: logo, width: 180 })
    }
    const out = templater().render(state.bytes, data)
    const ext = state.name.split('.').pop().toLowerCase()
    state.url = URL.createObjectURL(new Blob([out], { type: MIME[ext] ?? 'application/octet-stream' }))
    const fileName = state.name.replace(/(\.[^.]+)$/, '-filled$1')
    els.download.href = state.url
    els.download.download = fileName
    els.download.textContent = `⬇ ${t('Download', 'Скачать')} ${fileName}`
    els.download.hidden = false
    els.previewText.textContent = extractText(out)
    els.preview.hidden = false
  } catch (e) {
    showError(e)
  }
}

function setPath(obj, path, value) {
  const segs = path.split('.')
  let cur = obj
  segs.slice(0, -1).forEach((s) => (cur = cur[s] && typeof cur[s] === 'object' ? cur[s] : (cur[s] = {})))
  cur[segs[segs.length - 1]] = value
}

async function loadSample(key) {
  document.querySelectorAll('.sample[data-sample]').forEach((b) => b.classList.toggle('active', b.dataset.sample === key))
  const sample = SAMPLES[key]
  const bytes = new Uint8Array(await (await fetch(sample.file)).arrayBuffer())
  loadTemplate(bytes, sample.file.split('/').pop(), sample.data)
}

document.querySelectorAll('.sample[data-sample]').forEach((b) => b.addEventListener('click', () => loadSample(b.dataset.sample)))
els.file.addEventListener('change', async () => {
  const f = els.file.files?.[0]
  if (!f) return
  document.querySelectorAll('.sample[data-sample]').forEach((b) => b.classList.remove('active'))
  loadTemplate(new Uint8Array(await f.arrayBuffer()), f.name)
  els.file.value = ''
})
for (const input of [els.start, els.end]) {
  input.addEventListener('change', () => state.bytes && loadTemplate(state.bytes, state.name, safeJson(els.data.value)))
}
els.render.addEventListener('click', render)

function safeJson(text) {
  try {
    return JSON.parse(text)
  } catch {
    return undefined
  }
}

// start: compile the engine without blocking, then open the first sample
initEngine()
  .then(() => {
    els.badge.textContent = `engine ${engineInfo().coreVersion} · wasm`
    return loadSample('invoice')
  })
  .catch((e) => showError(e))
