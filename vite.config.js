import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import path from 'node:path'
import { pathToFileURL } from 'node:url'

// After the app is built, write the static /steel-sections pages (build-pages.mjs).
// Any problem there is logged and skipped: it can never stop the site from deploying.
function sectionPages() {
  let root = process.cwd()
  let outDir = 'dist'
  return {
    name: 'steel-section-pages',
    apply: 'build',
    configResolved(config) {
      root = config.root
      outDir = path.resolve(config.root, config.build.outDir)
    },
    async closeBundle() {
      try {
        const mod = await import(pathToFileURL(path.join(root, 'build-pages.mjs')).href)
        await mod.buildSectionPages({ root, outDir })
      } catch (e) {
        console.warn('[section-pages] skipped:', e && e.message)
      }
    },
  }
}

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), sectionPages()],
})
