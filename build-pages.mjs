/* Steel Optimizer - static section pages, built from the app's own section library.
   ----------------------------------------------------------------------------
   Runs at the end of `vite build` (see vite.config.js). It NEVER stops a build:
   any problem is logged and the site deploys exactly as before, without the pages.

   First batch (2026-10-02): every library section that has full verified
   dimensions (SECTION_DIMS), in English.
     /steel-sections/<family>/<section>   one page per section
     /steel-sections                      hub page listing them all
   The numbers come from the app itself: STEEL_DB, SECTION_DIMS and the
   @@SHARED-CORE block of src/SteelCutOptimizer.jsx are read and evaluated
   here, so the pages and the app can never disagree.

   The in-app links to these pages stay OFF on purpose: the app shows them only
   when /steel-sections/s.css exists, and these pages inline their CSS instead.

   IndexNow: on a Vercel production build only, new or changed pages are sent to
   api.indexnow.org (Bing and the other IndexNow engines). The list of what is
   live is kept in /steel-sections/manifest.json and compared on the next build.

   Paid launch (release o): patchLaunch() reads launch.json at the repo root. When it
   says {"paid": true}, the built home page (dist/index.html) loses the word "Free"
   (title, description, link previews, noscript text) and its Schema price becomes
   29 USD a month, tax included; Bing is told once that the home page changed.
   {"paid": false}, a missing file or bad JSON: the home page is left exactly as it is.
*/
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import crypto from "node:crypto";

const SITE = "https://steeloptimizer.com";
const INDEXNOW_ENDPOINT = "https://api.indexnow.org/indexnow";
const BUILD_DATE = new Date().toISOString().slice(0, 10);

const FAMILY = {
  IPE: { title: "IPE beams", std: "EN 10365", kind: "an I-beam with parallel flanges" },
  HEA: { title: "HEA wide-flange beams", std: "EN 10365", kind: "a light H-section (HE A)" },
  HEB: { title: "HEB wide-flange beams", std: "EN 10365", kind: "an H-section (HE B)" },
  HEM: { title: "HEM wide-flange beams", std: "EN 10365", kind: "a heavy H-section (HE M)" },
  UPN: { title: "UPN channels", std: "EN 10365", kind: "a channel with tapered flanges" },
  UPE: { title: "UPE channels", std: "EN 10365", kind: "a channel with parallel flanges" },
  UB: { title: "Universal Beams (UB)", std: "BS 4-1", kind: "a universal beam" },
  UC: { title: "Universal Columns (UC)", std: "BS 4-1", kind: "a universal column" },
  W: { title: "W shapes (wide-flange)", std: "AISC Shapes Database v16.0", kind: "a wide-flange shape", aisc: true },
  "C-AMER": { title: "C channels (American Standard)", std: "AISC Shapes Database v16.0", kind: "an American Standard channel", aisc: true },
  MC: { title: "MC channels (miscellaneous)", std: "AISC Shapes Database v16.0", kind: "a miscellaneous channel", aisc: true },
  "JIS-HW": { title: "HW H-sections (wide flange)", std: null, kind: "a wide-flange H-section (HW series)" },
  "JIS-HM": { title: "HM H-sections (medium flange)", std: null, kind: "a medium-flange H-section (HM series)" },
  "JIS-HN": { title: "HN H-sections (narrow flange)", std: null, kind: "a narrow-flange H-section (HN series)" },
  "JIS-I": { title: "I-sections (JIS)", std: "JIS G 3192", kind: "an I-section" },
  "JIS-C": { title: "Channels (JIS)", std: "JIS G 3192", kind: "a channel" },
};
const FAMILY_ORDER = ["IPE", "HEA", "HEB", "HEM", "UPN", "UPE", "UB", "UC", "W", "C-AMER", "MC", "JIS-HW", "JIS-HM", "JIS-HN", "JIS-I", "JIS-C"];

// link-preview image: the launch version drops "No signup" from the footer (public/preview-pro.png)
let ogImage = "/preview.png";

const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const num = (v, d) => {
  const f = Number(v).toFixed(d);
  const [i, dec] = f.split(".");
  return i.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (dec ? "." + dec : "");
};
const trim0 = v => String(Math.round(v * 100) / 100);            // 8.50 -> 8.5, 14 -> 14
const kgFmt = v => { const [i, d] = String(Math.round(v * 100) / 100).split("."); return i.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (d ? "." + d : ""); };   // library mass as stored: 9.36, 88.3, 117
const upTo = (v, d) => Math.ceil(v * 10 ** d - 1e-9) / 10 ** d;   // conservative rounding (paint area)

function slice(src, start, end, label) {
  const a = src.indexOf(start);
  if (a < 0) throw new Error(label + ": start marker not found");
  const b = src.indexOf(end, a);
  if (b < 0) throw new Error(label + ": end marker not found");
  return src.slice(a, b + end.length);
}

function loadLibrary(jsx) {
  const db = slice(jsx, "const STEEL_DB = [", "].map(([name, type, kgm]) => ({ name, type, kgm }));", "STEEL_DB");
  const dims = slice(jsx, "const SECTION_DIMS = {", "\n};", "SECTION_DIMS");
  const core = slice(jsx, "/*@@SHARED-CORE:BEGIN@@*/", "/*@@SHARED-CORE:END@@*/", "SHARED-CORE");
  const code = db + "\n" + dims + "\n" + core +
    "\n;({ STEEL_DB, SECTION_DIMS, sectionDims, sectionPaintM2pm, sectionPagePath, sectionDataIssue, sectionAliasOf, fmtM2pm });";
  return vm.runInNewContext(code, {}, { timeout: 10000, filename: "section-core.js" });
}

const CSS = `:root{--ink:#0b0f17;--mut:#475569;--acc:#b45309;--line:#e2e8f0;--soft:#f8fafc}
*{box-sizing:border-box}
body{margin:0;font:17px/1.7 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;color:var(--ink);background:#fff}
.wrap{max-width:760px;margin:0 auto;padding:28px 20px 64px}
nav.crumbs{font-size:14px;color:var(--mut);margin-bottom:18px}
nav.crumbs a{color:var(--mut)}
h1{font-size:32px;line-height:1.2;margin:0 0 4px}
.sub{color:var(--mut);font-size:16px;margin:0 0 18px}
h2{font-size:20px;margin:30px 0 10px}
p{margin:0 0 12px}
a{color:var(--acc)}
table{width:100%;border-collapse:collapse;margin:8px 0 14px;font-size:16px}
th,td{text-align:left;padding:9px 10px;border-bottom:1px solid var(--line);vertical-align:top}
th{font-weight:500;color:var(--mut)}
td{font-weight:600;text-align:right;white-space:nowrap}
.note{font-size:14px;color:var(--mut)}
.cta{margin:28px 0 8px;padding:20px;background:var(--ink);border-radius:10px;color:#e2e8f0}
.cta p{margin:0 0 12px}
.cta a.btn{display:inline-block;background:#f59e0b;color:#1a1206;font-weight:700;text-decoration:none;padding:11px 16px;border-radius:8px}
.nb{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;margin:18px 0 0;font-size:15px}
.links{columns:2;column-gap:24px;font-size:15px;padding-left:18px}
.links li{margin:0 0 4px;break-inside:avoid}
footer{margin-top:40px;padding-top:14px;border-top:1px solid var(--line);font-size:14px;color:var(--mut)}
footer a{color:var(--mut)}
@media (min-width:640px){.links{columns:3}}`;

function page({ title, desc, canonical, body, jsonld }) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(title)}</title>
<meta name="description" content="${esc(desc)}">
<link rel="canonical" href="${esc(canonical)}">
<meta name="robots" content="index, follow">
<link rel="icon" type="image/svg+xml" href="/favicon.svg">
<meta property="og:type" content="website">
<meta property="og:title" content="${esc(title)}">
<meta property="og:description" content="${esc(desc)}">
<meta property="og:url" content="${esc(canonical)}">
<meta property="og:image" content="${SITE}${ogImage}">
<script type="application/ld+json">${JSON.stringify(jsonld).replace(/</g, "\\u003c")}</script>
<style>${CSS}</style>
</head>
<body>
<div class="wrap">
${body}
<footer><a href="/">Steel Optimizer</a> · <a href="/steel-sections">Steel section tables</a> · <a href="/privacy.html">Privacy Policy</a> · <a href="mailto:support@steeloptimizer.com">support@steeloptimizer.com</a></footer>
</div>
</body>
</html>
`;
}

function sectionPage(p, prev, next) {
  const f = FAMILY[p.type];
  const [h, b, tw, tf] = p.dims;
  const famUrl = `/steel-sections#${p.famSlug}`;
  const lbft = f.aisc ? p.kgm / 1.48816 : null;
  const areaCm2 = p.kgm / 0.785;                              // A [cm2] = kg/m / 0.785 at 7850 kg/m3
  const paint = p.paint ? upTo(p.paint, 3) : null;
  const paintPerT = p.paint ? upTo(p.paint / (p.kgm / 1000), 1) : null;
  const w6 = p.kgm * 6, w12 = p.kgm * 12;
  const title = `${p.name} dimensions and weight: ${kgFmt(p.kgm)} kg/m | Steel Optimizer`;
  const desc = `${p.name}${f.std ? " (" + f.std + ")" : ""}: depth ${trim0(h)} mm, width ${trim0(b)} mm, web ${trim0(tw)} mm, flange ${trim0(tf)} mm, ${kgFmt(p.kgm)} kg/m.` +
    (paint ? ` Paint area ${paint.toFixed(3)} m²/m.` : "") + " Weights of 6 m and 12 m lengths.";
  const canonical = SITE + p.url;
  const twins = p.twins.length ? `<p class="note">Also listed as ${p.twins.map(esc).join(", ")}.</p>` : "";
  const rows = (pairs) => pairs.filter(Boolean).map(([k, v]) => `<tr><th>${k}</th><td>${v}</td></tr>`).join("");
  const body = `<nav class="crumbs"><a href="/">Steel Optimizer</a> › <a href="/steel-sections">Steel sections</a> › <a href="${famUrl}">${esc(f.title)}</a></nav>
<h1>${esc(p.name)}</h1>
<p class="sub">Dimensions, weight per metre and paint area</p>
<p>${esc(p.name)} is ${esc(f.kind)} weighing <strong>${kgFmt(p.kgm)} kg/m</strong>${lbft ? ` (${num(lbft, 1)} lb/ft)` : ""}: a 12 m length weighs ${num(w12, 1)} kg. It is ${trim0(h)} mm deep and ${trim0(b)} mm wide; the web is ${trim0(tw)} mm and the flanges ${trim0(tf)} mm thick.</p>
${twins}
<h2>Dimensions</h2>
<table>${rows([["Depth h", trim0(h) + " mm"], ["Flange width b", trim0(b) + " mm"], ["Web thickness t<sub>w</sub>", trim0(tw) + " mm"], ["Flange thickness t<sub>f</sub>", trim0(tf) + " mm"]])}</table>
<h2>Weight and area</h2>
<table>${rows([
    ["Weight per metre", kgFmt(p.kgm) + " kg/m"],
    lbft ? ["Weight per foot", num(lbft, 1) + " lb/ft"] : null,
    ["Cross-section area", num(areaCm2, 1) + " cm²"],
    paint ? ["Paint area per metre", paint.toFixed(3) + " m²/m"] : null,
    paint ? ["Paint area per tonne", paintPerT.toFixed(1) + " m²/t"] : null,
  ])}</table>
<h2>Weight by length</h2>
<table>${rows([["1 m", kgFmt(p.kgm) + " kg"], ["6 m", num(w6, 1) + " kg"], ["12 m", num(w12, 1) + " kg"]])}</table>
<p class="note">${f.std ? "Dimensions and mass per " + esc(f.std) + ", from Steel Optimizer's verified section library. " : "From Steel Optimizer's verified section library. "}Area is worked out from the mass at 7,850 kg/m³. Paint area counts the outside perimeter with square corners and ignores root radii, so it is on the safe side: typically 2–5 % above catalogue surface areas.</p>
<div class="cta"><p><strong>Cutting a list of ${esc(p.name)}?</strong> Steel Optimizer reads your Excel, Tekla or Advance Steel list as it is and gives the cutting plan, the bars to buy and the waste in seconds.</p><a class="btn" href="/?section=${encodeURIComponent(p.name)}">Optimize a cutting list with ${esc(p.name)} →</a></div>
<div class="nb">${prev ? `<a href="${prev.url}">← ${esc(prev.name)}</a>` : "<span></span>"}<a href="${famUrl}">All ${esc(f.title)}</a>${next ? `<a href="${next.url}">${esc(next.name)} →</a>` : "<span></span>"}</div>`;
  const jsonld = {
    "@context": "https://schema.org",
    "@type": "BreadcrumbList",
    itemListElement: [
      { "@type": "ListItem", position: 1, name: "Steel Optimizer", item: SITE + "/" },
      { "@type": "ListItem", position: 2, name: "Steel sections", item: SITE + "/steel-sections" },
      { "@type": "ListItem", position: 3, name: f.title, item: SITE + famUrl },
      { "@type": "ListItem", position: 4, name: p.name, item: canonical },
    ],
  };
  return page({ title, desc, canonical, body, jsonld });
}

function hubPage(groups, total) {
  const canonical = SITE + "/steel-sections";
  const title = "Steel section tables: IPE, HEA, HEB, UB, UC and W shapes | Steel Optimizer";
  const desc = `Dimensions, weight per metre and paint area for ${total} steel sections: IPE, HEA, HEB, HEM, UPN, UPE, UB, UC, W shapes, C and MC channels, and Asian H-sections.`;
  const fam = groups.map(([type, list]) => `<h2 id="${list[0].famSlug}">${esc(FAMILY[type].title)}${FAMILY[type].std ? ` <span class="note">(${esc(FAMILY[type].std)})</span>` : ""}</h2>
<ul class="links">${list.map(p => `<li><a href="${p.url}">${esc(p.name)}</a> <span class="note">${kgFmt(p.kgm)} kg/m</span></li>`).join("")}</ul>`).join("\n");
  const body = `<nav class="crumbs"><a href="/">Steel Optimizer</a> › Steel sections</nav>
<h1>Steel section tables</h1>
<p class="sub">Dimensions, weight per metre and paint area for ${total} sections</p>
<p>Pick a section for its depth, width, web and flange thickness, weight per metre and per 6 m and 12 m length, and a conservative paint area. Families: ${groups.map(([t, l]) => `<a href="#${l[0].famSlug}">${esc(FAMILY[t].title)}</a>`).join(", ")}.</p>
${fam}
<div class="cta"><p><strong>Have a material list to cut?</strong> Steel Optimizer reads Excel, Tekla and Advance Steel lists as they are and returns the cutting plan, the bars and sheets to buy and the waste in seconds.</p><a class="btn" href="/">Open Steel Optimizer →</a></div>`;
  return page({ title, desc, canonical, body, jsonld: { "@context": "https://schema.org", "@type": "CollectionPage", name: "Steel section tables", url: canonical } });
}

function writeFile(file, text) { fs.mkdirSync(path.dirname(file), { recursive: true }); fs.writeFileSync(file, text); }

function updateSitemap(outDir, urls, log) {
  const file = path.join(outDir, "sitemap.xml");
  if (!fs.existsSync(file)) { log("[section-pages] no sitemap.xml in the build, sitemap not updated"); return; }
  let xml = fs.readFileSync(file, "utf8");
  const have = new Set([...xml.matchAll(/<loc>([^<]+)<\/loc>/g)].map(m => m[1].trim()));
  const add = urls.filter(u => !have.has(u)).map(u => `  <url>\n    <loc>${u}</loc>\n    <lastmod>${BUILD_DATE}</lastmod>\n    <changefreq>monthly</changefreq>\n    <priority>0.6</priority>\n  </url>\n`).join("");
  if (!add || !xml.includes("</urlset>")) return;
  xml = xml.replace("</urlset>", add + "\n</urlset>");
  fs.writeFileSync(file, xml);
}

async function fetchJson(url, ms) {
  const ctl = new AbortController(); const tm = setTimeout(() => ctl.abort(), ms);
  try { const r = await fetch(url, { signal: ctl.signal, headers: { "cache-control": "no-cache" } }); return r.ok ? await r.json() : null; }
  catch { return null; } finally { clearTimeout(tm); }
}

async function fetchText(url, ms) {
  const ctl = new AbortController(); const tm = setTimeout(() => ctl.abort(), ms);
  try { const r = await fetch(url, { signal: ctl.signal, headers: { "cache-control": "no-cache" } }); return r.ok ? await r.text() : null; }
  catch { return null; } finally { clearTimeout(tm); }
}

function indexNowKey(root) {
  const pub = path.join(root, "public");
  return fs.existsSync(pub) ? fs.readdirSync(pub).find(f => /^[0-9a-f]{32}\.txt$/.test(f)) || null : null;
}

async function indexNowSend(keyFile, urls, tag, log) {
  const ctl = new AbortController(); const tm = setTimeout(() => ctl.abort(), 15000);
  try {
    const r = await fetch(INDEXNOW_ENDPOINT, {
      method: "POST", signal: ctl.signal,
      headers: { "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({ host: new URL(SITE).host, key: keyFile.slice(0, 32), keyLocation: `${SITE}/${keyFile}`, urlList: urls }),
    });
    log(`${tag} IndexNow: sent ${urls.length} URL(s), HTTP ${r.status}`);
  } catch (e) { log(`${tag} IndexNow: not sent (` + (e && e.message) + ")"); }
  finally { clearTimeout(tm); }
}

async function pingIndexNow(root, manifest, log) {
  if (process.env.VERCEL_ENV !== "production" || typeof fetch !== "function") return;
  const keyFile = indexNowKey(root);
  if (!keyFile) { log("[section-pages] IndexNow: no key file in public/, skipped"); return; }
  const live = (await fetchJson(SITE + "/steel-sections/manifest.json", 8000)) || {};
  const changed = Object.keys(manifest).filter(u => live[u] !== manifest[u]).map(u => SITE + u).slice(0, 10000);
  if (!changed.length) { log("[section-pages] IndexNow: nothing new or changed"); return; }
  await indexNowSend(keyFile, changed, "[section-pages]", log);
}

/* ---------- paid launch: launch.json → home page ---------- */
export function readLaunchFlag(root = process.cwd()) {
  try { const j = JSON.parse(fs.readFileSync(path.join(root, "launch.json"), "utf8").replace(/^\uFEFF/, "")); return !!(j && j.paid === true); }
  catch { return false; }
}

const OLD_TITLE = "Steel Cutting Optimizer — Free Bar, Section &amp; Plate Nesting Calculator";
// [as written in index.html, once launched]; each edit is applied only where its text is found
const LAUNCH_EDITS = [
  [`<title>${OLD_TITLE}</title>`, "<title>Steel Cutting Optimizer — Bar, Section &amp; Plate Nesting Calculator</title>"],
  [`content="Free online steel cutting optimizer. `, `content="Online steel cutting optimizer. `],
  [`content="Steel Cutting Optimizer — Free Bar, Section &amp; Plate Nesting"`, `content="Steel Cutting Optimizer — Bar, Section &amp; Plate Nesting"`],
  [`in one click. No install, no licence."`, `in one click. No install."`],
  [`content="Steel Cutting Optimizer — Free"`, `content="Steel Cutting Optimizer"`],
  [`<!-- Schema.org — tells Google this is a free web application -->`, `<!-- Schema.org — web application, monthly subscription, price includes tax -->`],
  [`"offers": { "@type": "Offer", "price": "0", "priceCurrency": "USD" }`,
   `"offers": { "@type": "Offer", "price": "29", "priceCurrency": "USD", "priceSpecification": { "@type": "UnitPriceSpecification", "price": "29", "priceCurrency": "USD", "unitText": "MONTH", "valueAddedTaxIncluded": true } }`],
  [`<p>Free browser-based cutting and nesting optimizer for structural steel.`, `<p>Browser-based cutting and nesting optimizer for structural steel.`],
  [`Runs in any browser. No install, no licence, no training.`, `Runs in any browser. No install, no training.`],
  [`content="https://steeloptimizer.com/preview.png"`, `content="https://steeloptimizer.com/preview-pro.png"`],   // og:image + twitter:image
];

export async function patchLaunch({ root = process.cwd(), outDir = "dist", log = console.log } = {}) {
  if (!readLaunchFlag(root)) { log("[launch] free: home page left as it is"); return { paid: false, edits: 0 }; }
  const file = path.join(outDir, "index.html");
  let html = fs.readFileSync(file, "utf8");
  let edits = 0; const missing = [];
  for (const [from, to] of LAUNCH_EDITS) {
    if (html.includes(from)) { html = html.split(from).join(to); edits++; }
    else if (!html.includes(to)) missing.push(from.slice(0, 48));
  }
  const ld = html.match(/<script type="application\/ld\+json">([\s\S]*?)<\/script>/);
  if (ld) JSON.parse(ld[1]);                    // the Schema block must still be valid JSON, or nothing is written
  fs.writeFileSync(file, html);
  const leftFree = (html.match(/\bFree\b/g) || []).length;
  log(`[launch] PAID: home page updated (${edits}/${LAUNCH_EDITS.length} edits` + (missing.length ? `; not found: ${missing.join(" | ")}` : "") + `; "Free" left: ${leftFree})`);
  // Bing: send the home page once, the first time the paid version goes live (the live page still has the old title)
  if (process.env.VERCEL_ENV === "production" && typeof fetch === "function" && edits) {
    const keyFile = indexNowKey(root);
    const live = await fetchText(SITE + "/?launch-check=" + Date.now(), 8000);
    if (keyFile && live && live.includes(OLD_TITLE)) await indexNowSend(keyFile, [SITE + "/"], "[launch]", log);
  }
  return { paid: true, edits, missing };
}

export async function buildSectionPages({ root = process.cwd(), outDir = "dist", log = console.log } = {}) {
  const t0 = Date.now();
  ogImage = readLaunchFlag(root) && fs.existsSync(path.join(root, "public", "preview-pro.png")) ? "/preview-pro.png" : "/preview.png";
  const jsx = fs.readFileSync(path.join(root, "src", "SteelCutOptimizer.jsx"), "utf8");
  const L = loadLibrary(jsx);
  const byPath = new Map();
  for (const row of L.STEEL_DB) {
    if (!L.SECTION_DIMS[row.name] || !FAMILY[row.type]) continue;      // first batch: full verified dimensions only
    if (L.sectionDataIssue(row)) continue;
    const url = L.sectionPagePath(row, "en");
    if (!url || byPath.has(url)) continue;
    const dims = L.sectionDims(row.name);
    if (!dims || dims.length < 4 || !(dims[0] > 0 && dims[1] > 0 && dims[2] > 0 && dims[3] > 0) || !(row.kgm > 0)) continue;
    const pm = L.sectionPaintM2pm(row);
    byPath.set(url, {
      name: row.name, type: row.type, kgm: +row.kgm, dims, url,
      famSlug: url.split("/")[2], paint: pm && pm.m2pm > 0 ? pm.m2pm : null,
      twins: L.STEEL_DB.filter(s => s.name !== row.name && L.sectionAliasOf(s.name) === row.name).map(s => s.name),
    });
  }
  const pages = [...byPath.values()];
  if (pages.length < 100) throw new Error(`only ${pages.length} pages - library not read as expected`);
  const groups = FAMILY_ORDER.map(t => [t, pages.filter(p => p.type === t).sort((a, c) => a.dims[0] - c.dims[0] || a.kgm - c.kgm)]).filter(([, l]) => l.length);
  const manifest = {};
  const hash = s => crypto.createHash("sha1").update(s).digest("hex").slice(0, 12);
  for (const [, list] of groups) {
    list.forEach((p, i) => {
      const html = sectionPage(p, list[i - 1], list[i + 1]);
      writeFile(path.join(outDir, p.url, "index.html"), html);
      manifest[p.url] = hash(html);
    });
  }
  const hub = hubPage(groups, pages.length);
  writeFile(path.join(outDir, "steel-sections", "index.html"), hub);
  manifest["/steel-sections"] = hash(hub);
  writeFile(path.join(outDir, "steel-sections", "manifest.json"), JSON.stringify(manifest));
  updateSitemap(outDir, ["/steel-sections", ...groups.flatMap(([, l]) => l.map(p => p.url))].map(u => SITE + u), log);
  log(`[section-pages] ${pages.length} section pages + hub written in ${Date.now() - t0} ms`);
  await pingIndexNow(root, manifest, log);
  return { pages: pages.length, manifest };
}

// CLI: node build-pages.mjs [outDir]
if (import.meta.url === `file://${process.argv[1]}`) {
  buildSectionPages({ outDir: process.argv[2] || "dist" }).catch(e => { console.warn("[section-pages] skipped:", e && e.message); });
}
