#!/usr/bin/env node
// Builds the GitHub Pages site in ./site:
//   site/lib/coffee-templater.js  — browser bundle of the library (engine + fflate included)
//   site/lib/docx-preview.js      — Word preview (docx-preview, Apache-2.0), loaded on demand
//   site/lib/zip.js               — ZIP reader used by the PowerPoint/Excel preview
//   site/samples/*                — sample templates from test/fixtures
// Both are generated (git-ignored). Run `npm run engine` first (npm run build:site does it).

import { build } from 'esbuild'
import { copyFileSync, mkdirSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const SITE = join(ROOT, 'site')

await build({
  entryPoints: [join(ROOT, 'src', 'index.ts')],
  outfile: join(SITE, 'lib', 'coffee-templater.js'),
  bundle: true,
  format: 'esm',
  platform: 'browser',
  target: 'es2020',
  minify: true,
  legalComments: 'none',
})

// Helpers for the preview. Third-party license comments are kept at the end of the bundle.
const helper = (name, contents) =>
  build({
    stdin: { contents, resolveDir: ROOT, loader: 'js' },
    outfile: join(SITE, 'lib', `${name}.js`),
    bundle: true,
    format: 'esm',
    platform: 'browser',
    target: 'es2020',
    minify: true,
    legalComments: 'eof',
  })
await helper('docx-preview', "export { renderAsync } from 'docx-preview'")
await helper('zip', "export { unzipSync, strFromU8 } from 'fflate'")

mkdirSync(join(SITE, 'samples'), { recursive: true })
for (const f of ['invoice.docx', 'report.xlsx', 'deck.pptx']) {
  copyFileSync(join(ROOT, 'test', 'fixtures', f), join(SITE, 'samples', f))
}
console.log('site built → site/')
