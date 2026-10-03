// Vercel serverless function - keeps the free Supabase project awake (Steel Optimizer, release o)
// Supabase pauses a Free project after about a week with too little database activity. Before the
// paid launch visitors never touch Supabase, so Vercel Cron (vercel.json) calls this once a day:
// three tiny reads with the server key. Nothing is returned and nothing is written.
// Uses the same Vercel environment variables as api/lemon-webhook.js:
//   SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY   (+ optional CRON_SECRET, see below)

export default async function handler(req, res) {
  // Only Vercel's scheduler makes the database calls. With a CRON_SECRET environment variable set in Vercel
  // (optional), Vercel sends it as "Authorization: Bearer <secret>" and nothing else is accepted.
  const secret = process.env.CRON_SECRET;
  const fromCron = secret
    ? String(req.headers["authorization"] || "") === "Bearer " + secret
    : !!req.headers["x-vercel-cron-schedule"] || /vercel-cron/i.test(String(req.headers["user-agent"] || ""));
  if (!fromCron) { res.statusCode = 204; res.end(); return; }

  const sbUrl = (process.env.SUPABASE_URL || "").replace(/\/+$/, "");
  const sbKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!sbUrl || !sbKey) { res.statusCode = 500; res.end("not configured"); return; }

  const headers = { apikey: sbKey, ...(sbKey.split(".").length === 3 ? { Authorization: "Bearer " + sbKey } : {}) };
  let ok = 0, last = 0;
  for (let i = 0; i < 3; i++) {
    const ctl = new AbortController();
    const tm = setTimeout(() => ctl.abort(), 8000);
    try {
      const r = await fetch(sbUrl + "/rest/v1/memberships?select=ls_subscription_id&limit=1", { headers, signal: ctl.signal });
      last = r.status;
      if (r.ok) ok++;
    } catch (e) {
      last = 0;
    } finally {
      clearTimeout(tm);
    }
  }
  res.statusCode = ok ? 200 : 502;
  res.end(ok ? "ok " + ok + "/3" : "supabase not reached (" + last + ")");
}
