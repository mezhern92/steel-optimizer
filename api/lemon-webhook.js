// Vercel serverless function — Lemon Squeezy → Supabase «memberships»  (Steel Optimizer)
// 1. GitHub: put this file at  api/lemon-webhook.js  (folder "api" at the top of the repo, next to "src" — not inside it)
// 2. Vercel → Project → Settings → Environment Variables (then Redeploy):
//      SUPABASE_URL                  https://cjeqlypqoxiwnleimwlx.supabase.co
//      SUPABASE_SERVICE_ROLE_KEY     the Supabase secret key (sb_secret_…) — only here, never in the app file
//      LEMONSQUEEZY_WEBHOOK_SECRET   any long random text; type the same text in Lemon Squeezy
// 3. Lemon Squeezy → Settings → Webhooks → +  URL: https://steeloptimizer.com/api/lemon-webhook
//      Signing secret: the same text · Events: every "subscription_…" event
// Checks the Lemon Squeezy signature, then saves the subscription (status, renewal date, customer portal)
// against the signed-in user id the app sends to the checkout. Other events are ignored.
import crypto from "node:crypto";

function readRaw(req) {
  return new Promise((resolve, reject) => {
    let d = "";
    req.setEncoding("utf8");
    req.on("data", c => { d += c; });
    req.on("end", () => resolve(d));
    req.on("error", reject);
  });
}
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export default async function handler(req, res) {
  if (req.method !== "POST") { res.statusCode = 405; res.end("POST only"); return; }
  const secret = process.env.LEMONSQUEEZY_WEBHOOK_SECRET;
  const sbUrl = (process.env.SUPABASE_URL || "").replace(/\/+$/, "");
  const sbKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!secret || !sbUrl || !sbKey) { res.statusCode = 500; res.end("not configured"); return; }

  const raw = await readRaw(req);                       // the exact bytes Lemon Squeezy signed
  const sig = String(req.headers["x-signature"] || "");
  const digest = crypto.createHmac("sha256", secret).update(raw).digest("hex");
  if (sig.length !== digest.length || !crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(digest))) {
    res.statusCode = 401; res.end("bad signature"); return;
  }

  let p;
  try { p = JSON.parse(raw); } catch { res.statusCode = 400; res.end("bad json"); return; }
  const data = p && p.data, a = data && data.attributes;
  if (!data || data.type !== "subscriptions" || !a) { res.statusCode = 200; res.end("ignored"); return; }

  const custom = (p.meta && p.meta.custom_data) || {};
  const row = {
    ls_subscription_id: String(data.id),
    email: String(a.user_email || "").toLowerCase(),
    status: String(a.status || ""),
    variant_id: a.variant_id != null ? String(a.variant_id) : null,
    renews_at: a.renews_at || null,
    ends_at: a.ends_at || null,
    trial_ends_at: a.trial_ends_at || null,
    portal_url: (a.urls && a.urls.customer_portal) || null,
    updated_at: new Date().toISOString(),
  };
  if (UUID.test(String(custom.user_id || ""))) row.user_id = String(custom.user_id);   // never overwrite with empty

  const r = await fetch(${sbUrl}/rest/v1/memberships?on_conflict=ls_subscription_id, {
    method: "POST",
    headers: { apikey: sbKey, ...(sbKey.split(".").length === 3 ? { Authorization: Bearer ${sbKey} } : {}),   // sb_secret_… keys: apikey header only
      "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" },
    body: JSON.stringify(row),
  });
  if (!r.ok) { res.statusCode = 500; res.end("db " + r.status + " " + (await r.text()).slice(0, 200)); return; }
  res.statusCode = 200; res.end("ok");
}
