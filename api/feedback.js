// Vercel serverless function - in-app feedback -> e-mail to the owner (Steel Optimizer, release o)
// The app posts {kind, rating, message, email, ctx, hp}. This checks it and e-mails it through Resend
// (free plan: 3,000 e-mails a month, 100 a day), so feedback never uses the Formspree quota.
// Vercel -> Project -> Settings -> Environment Variables (then Redeploy):
//   RESEND_API_KEY   a Resend API key with sending access (the key used for Supabase SMTP is fine)
//   FEEDBACK_TO      the inbox that receives feedback - kept here, never in the app file
//   FEEDBACK_FROM    optional, default "Steel Optimizer <feedback@steeloptimizer.com>" (domain verified in Resend)
// Missing RESEND_API_KEY or FEEDBACK_TO -> 503, and the app falls back to Formspree.
// Nothing about files is ever sent here: only the message, the optional e-mail, a rating and app context.

const MAX_BODY = 8000;
const hits = new Map();                                   // per-instance throttle: ip -> recent timestamps
const CTX_KEYS = ["mod", "lang", "ver", "signed", "plan", "launched", "visit", "screen", "parts", "sheets", "thicknesses", "profiles", "bars", "waste"];

function readRaw(req) {
  return new Promise((resolve, reject) => {
    let d = "", size = 0;
    req.setEncoding("utf8");
    req.on("data", c => { size += c.length; if (size > MAX_BODY) { reject(new Error("too large")); return; } d += c; });
    req.on("end", () => resolve(d));
    req.on("error", reject);
  });
}
function send(res, code, text) { res.statusCode = code; res.setHeader("Content-Type", "application/json"); res.end(JSON.stringify(text)); }
function clip(v, n) { return String(v == null ? "" : v).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "").slice(0, n); }
function oneLine(v, n) { return clip(v, n).replace(/[\r\n\t]+/g, " ").trim(); }

export default async function handler(req, res) {
  if (req.method !== "POST") { send(res, 405, { error: "POST only" }); return; }

  // only the site's own pages (browsers always send Origin on a POST from a page)
  const origin = String(req.headers.origin || req.headers.referer || "");
  const prod = process.env.VERCEL_ENV === "production";
  const okOrigin = /^https:\/\/(www\.)?steeloptimizer\.com(\/|$)/.test(origin) ||
    (!prod && /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?(\/|$)/.test(origin));
  if (!okOrigin) { send(res, 403, { error: "origin" }); return; }

  // throttle: 5 messages per 10 minutes per address (per server instance)
  const ip = String(req.headers["x-forwarded-for"] || req.socket && req.socket.remoteAddress || "?").split(",")[0].trim();
  const now = Date.now();
  const list = (hits.get(ip) || []).filter(t => now - t < 600000);
  if (list.length >= 5) { send(res, 429, { error: "slow down" }); return; }
  list.push(now); hits.set(ip, list);
  if (hits.size > 5000) hits.clear();

  let p;
  try { p = JSON.parse(await readRaw(req)); } catch (e) { send(res, 400, { error: "bad request" }); return; }
  if (!p || typeof p !== "object") { send(res, 400, { error: "bad request" }); return; }
  if (p.hp) { send(res, 200, { ok: true }); return; }     // honeypot filled in: a bot - say ok, send nothing

  const kind = p.kind === "result" ? "result" : "general";
  const rating = p.rating === "up" || p.rating === "down" ? p.rating : "";
  const message = clip(p.message, 2000).trim();
  const email = clip(p.email, 200).trim();
  if (!message && !rating) { send(res, 400, { error: "empty" }); return; }
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) { send(res, 400, { error: "email" }); return; }
  if ((message.match(/https?:\/\//gi) || []).length > 3) { send(res, 200, { ok: true }); return; }   // link spam

  const ctx = {};
  const c = p.ctx && typeof p.ctx === "object" ? p.ctx : {};
  for (const k of CTX_KEYS) if (c[k] !== undefined && c[k] !== null && c[k] !== "") ctx[k] = oneLine(c[k], 40);

  const key = process.env.RESEND_API_KEY, to = process.env.FEEDBACK_TO;
  if (!key || !to) { send(res, 503, { error: "not configured" }); return; }
  const from = process.env.FEEDBACK_FROM || "Steel Optimizer <feedback@steeloptimizer.com>";

  const verdict = rating === "up" ? "useful" : rating === "down" ? "NOT useful" : "";
  const where = kind === "result" ? "result" + (ctx.mod ? " / " + ctx.mod : "") : "footer";
  const subject = "Steel Optimizer feedback" + (verdict ? " - " + verdict : "") + " (" + where + (ctx.lang ? ", " + ctx.lang : "") + ")";
  const text = [
    "Rating: " + (verdict || "-"),
    "",
    message || "(no message)",
    "",
    "Reply to: " + (email || "not given"),
    "Where: " + where,
    "Context: " + Object.keys(ctx).map(k => k + "=" + ctx[k]).join(", "),
    "Country: " + oneLine(req.headers["x-vercel-ip-country"] || "?", 8),
    "Time: " + new Date(now).toISOString(),
  ].join("\n");

  const ctl = new AbortController();
  const tm = setTimeout(() => ctl.abort(), 8000);
  try {
    const r = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { Authorization: "Bearer " + key, "Content-Type": "application/json" },
      body: JSON.stringify({ from, to: [to], subject, text, ...(email ? { reply_to: email } : {}) }),
      signal: ctl.signal,
    });
    if (!r.ok) { send(res, 502, { error: "mail " + r.status }); return; }
    send(res, 200, { ok: true });
  } catch (e) {
    send(res, 502, { error: "mail not sent" });
  } finally {
    clearTimeout(tm);
  }
}
