#!/usr/bin/env node
// Provides the coffee-templater-engine WebAssembly binary to this package.
//
// The engine lives in its own repository (https://github.com/vitJR1/coffee-templater-engine) and is
// NOT vendored here. This script takes the prebuilt `coffee_core.wasm` and generates
// `src/engine/wasm-binary.ts`, which embeds it (deflated) into the bundle — so users install a single package.
//
// Source, in order:
//   1. COFFEE_ENGINE_WASM=/path/to/coffee_core.wasm   — a local engine build (engine development)
//   2. a GitHub Release of the engine                   — package.json → "coffeeEngine"
//
// "coffeeEngine": { "repository", "version", "release"?, "sha256"? }
//   release  release tag to download from; default `v<version>`. A rolling tag such as `nightly` tracks
//            the engine's main branch: it is never cached and not checked against sha256.
//   sha256   expected checksum of coffee_core.wasm for a fixed release (required to publish to npm).
//
// Fixed releases are cached in node_modules/.cache/coffee-engine and verified against the pinned sha256.

import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { deflateRawSync } from 'node:zlib'

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const OUT = join(ROOT, 'src', 'engine', 'wasm-binary.ts')
const pkg = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8'))
const pin = pkg.coffeeEngine
if (!pin?.version || !pin?.repository) fail('package.json must define "coffeeEngine": { "repository", "version", "release"?, "sha256"? }')
const tag = pin.release ?? `v${pin.version}`
/** Version tags are immutable; anything else (e.g. `nightly`) moves with the engine's main branch. */
const rolling = !/^v\d/.test(tag)

/** Bump when the layout of the generated module changes. */
const FORMAT = 2
const sha256 = (buf) => createHash('sha256').update(buf).digest('hex')

function fail(message) {
  console.error(`\n[coffee-engine] ${message}\n`)
  process.exit(1)
}

function checkWasm(buf, where) {
  if (buf.length < 8 || buf[0] !== 0x00 || buf[1] !== 0x61 || buf[2] !== 0x73 || buf[3] !== 0x6d) fail(`${where} is not a WebAssembly binary`)
}

/** Delays before the 2nd, 3rd and 4th attempt. GitHub answers 5xx/429 under load, briefly. */
const RETRY_DELAYS_MS = [2000, 5000, 10000]
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

/**
 * Downloads a release asset, retrying network errors, 429 and 5xx. With GITHUB_TOKEN (set in CI) the
 * request is authenticated: GitHub throttles anonymous downloads first. The token goes to github.com
 * only; fetch drops it on the redirect to the asset storage.
 */
async function fetchWithRetry(url) {
  const token = process.env.GITHUB_TOKEN
  let auth = Boolean(token)
  let last = ''
  for (let attempt = 0; attempt <= RETRY_DELAYS_MS.length; attempt++) {
    if (attempt > 0) {
      const delay = RETRY_DELAYS_MS[attempt - 1]
      console.warn(`[coffee-engine] ${last}; retrying in ${delay / 1000} s (attempt ${attempt + 1} of ${RETRY_DELAYS_MS.length + 1})`)
      await sleep(delay)
    }
    try {
      const res = await fetch(url, { redirect: 'follow', headers: auth ? { Authorization: `Bearer ${token}` } : {} })
      if (res.ok) return res
      if (auth && (res.status === 401 || res.status === 403)) {
        // a token without access to this repository: try anonymously
        auth = false
        last = `HTTP ${res.status} with GITHUB_TOKEN`
        continue
      }
      if (res.status !== 429 && res.status < 500) {
        fail(
          `Cannot download the engine: ${url} → HTTP ${res.status}.\n` +
            `Make sure release ${tag} of ${pin.repository} exists, or set COFFEE_ENGINE_WASM=/path/to/coffee_core.wasm.`,
        )
      }
      last = `HTTP ${res.status}`
    } catch (e) {
      last = e.message
    }
  }
  fail(
    `Cannot download the engine from ${url} after ${RETRY_DELAYS_MS.length + 1} attempts: ${last}.\n` +
      'GitHub may be having trouble; try again later, or set COFFEE_ENGINE_WASM to a local coffee_core.wasm to build offline.',
  )
}

async function download() {
  const cacheDir = join(ROOT, 'node_modules', '.cache', 'coffee-engine')
  const cached = join(cacheDir, `coffee_core-${tag}.wasm`)
  if (!rolling && existsSync(cached)) {
    const buf = readFileSync(cached)
    if (!pin.sha256 || sha256(buf) === pin.sha256) return { buf, source: `cache (${tag})` }
  }
  const url = `https://github.com/${pin.repository}/releases/download/${tag}/coffee_core.wasm`
  const res = await fetchWithRetry(url)
  const buf = Buffer.from(await res.arrayBuffer())
  checkWasm(buf, url)
  const actual = sha256(buf)
  if (rolling) {
    console.warn(`[coffee-engine] using the rolling "${tag}" engine build (sha256 ${actual}); pin a version release before publishing`)
    return { buf, source: url }
  }
  if (pin.sha256 && actual !== pin.sha256) {
    fail(`sha256 mismatch for ${url}\n  expected ${pin.sha256}\n  actual   ${actual}`)
  }
  if (!pin.sha256) console.warn(`[coffee-engine] warning: no sha256 pinned in package.json; downloaded binary has sha256 ${actual}`)
  mkdirSync(cacheDir, { recursive: true })
  writeFileSync(cached, buf)
  return { buf, source: url }
}

async function main() {
  let buf
  let source
  const override = process.env.COFFEE_ENGINE_WASM
  if (override) {
    const path = resolve(override)
    if (!existsSync(path)) fail(`COFFEE_ENGINE_WASM points to a missing file: ${path}`)
    buf = readFileSync(path)
    checkWasm(buf, path)
    source = path
    if (pin.sha256 && sha256(buf) !== pin.sha256) {
      console.warn(`[coffee-engine] note: local engine differs from the pinned ${tag} (sha256 ${sha256(buf)})`)
    }
  } else {
    // already generated for the pinned release (e.g. by an earlier CI step): nothing to download
    if (!rolling && pin.sha256 && existsSync(OUT) && readFileSync(OUT, 'utf8').includes(`// format ${FORMAT} · sha256 ${pin.sha256}`)) return
    ;({ buf, source } = await download())
  }

  const hash = sha256(buf)
  const stamp = `// format ${FORMAT} · sha256 ${hash}`
  if (existsSync(OUT) && readFileSync(OUT, 'utf8').includes(stamp)) return
  mkdirSync(dirname(OUT), { recursive: true })
  writeFileSync(
    OUT,
    '// Generated by scripts/fetch-engine.mjs — do not edit, do not commit.\n' +
      `${stamp}\n` +
      `// Source: ${override ? 'local build (COFFEE_ENGINE_WASM)' : source}\n` +
      `export const ENGINE_VERSION = '${pin.version}'\n` +
      `export const ENGINE_SHA256 = '${hash}'\n` +
      `export const WASM_SIZE = ${buf.length}\n` +
      '/** coffee_core.wasm, raw-deflated and base64-encoded. */\n' +
      `export const WASM_DEFLATED_BASE64 = '${deflateRawSync(buf, { level: 9 }).toString('base64')}'\n`,
  )
  console.log(`[coffee-engine] embedded coffee_core.wasm (${(buf.length / 1024).toFixed(0)} KiB, sha256 ${hash.slice(0, 12)}…) from ${source}`)
}

main()
