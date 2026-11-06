import { defineConfig } from 'tsup'

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['esm', 'cjs'],
  dts: true,
  // no sourcemaps: sources are not published, and the maps would duplicate the embedded engine
  sourcemap: false,
  clean: true,
  target: 'es2020',
})
