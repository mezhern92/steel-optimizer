import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

// Paid-launch switch (release o): launch.json at the repo root.
//   {"paid": false}  the app stays free for visitors (owner tests with ?paytest=1)
//   {"paid": true}   sign-in + payment for everyone, live checkout, paid home page
// Anything missing or unreadable counts as false, so a typo can never lock visitors out.
function readLaunch() {
  const dirs = []
  try { dirs.push(path.dirname(fileURLToPath(import.meta.url))) } catch { /* ignore */ }
  dirs.push(process.cwd())
  for (const dir of dirs) {
    try {
      const j = JSON.parse(fs.readFileSync(path.join(dir, 'launch.json'), 'utf8').replace(/^\uFEFF/, ''))
      return !!(j && j.paid === true)
    } catch { /* try the next place */ }
  }
  return false
}
const LAUNCHED = readLaunch()
console.log('[launch] launch.json -> ' + (LAUNCHED ? 'PAID (sign-in and payment for everyone)' : 'free (visitors use the app as before)'))

// After the app is built, write the static /steel-sections pages (build-pages.mjs) and, once
// launched, take "Free" off the home page. Any problem there is logged and skipped: it can
// never stop the site from deploying.
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
      let mod = null
      try {
        mod = await import(pathToFileURL(path.join(root, 'build-pages.mjs')).href)
      } catch (e) {
        console.warn('[section-pages] skipped:', e && e.message)
        return
      }
      try {
        await mod.buildSectionPages({ root, outDir })
      } catch (e) {
        console.warn('[section-pages] skipped:', e && e.message)
      }
      try {
        if (typeof mod.patchLaunch === 'function') await mod.patchLaunch({ root, outDir })
      } catch (e) {
        console.warn('[launch] home page not changed:', e && e.message)
      }
    },
  }
}

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), sectionPages()],
  define: { __SO_LAUNCHED__: JSON.stringify(LAUNCHED) },
})
