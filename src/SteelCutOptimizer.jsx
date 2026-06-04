import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── NESTING ENGINE ────────────────────────────────────────────────────────────
function nestParts(sheetW, sheetH, parts, kerf = 2, margin = 5, allowRotation = true) {
  const sheets = [];
  let queue = [];
  parts.forEach((p) => { for (let q = 0; q < p.qty; q++) queue.push({ w: p.width, h: p.length, id: p.id, label: p.label }); });
  queue.sort((a, b) => (b.w * b.h) - (a.w * a.h));
  const usableW = sheetW - margin * 2, usableH = sheetH - margin * 2;
  let guard = 0;
  while (queue.length > 0 && guard < 10000) {
    guard++;
    const placements = [];
    let shelf = { usedW: 0, rowH: 0 }, shelfY = margin;
    const remaining = [];
    queue.forEach((part) => {
      const tryPlace = (w, h, rot) => {
        if (shelf.usedW + w + kerf <= usableW && h <= (usableH - (shelfY - margin))) {
          placements.push({ x: margin + shelf.usedW, y: shelfY, w, h, rotated: rot, id: part.id, label: part.label });
          shelf.usedW += w + kerf; shelf.rowH = Math.max(shelf.rowH, h); return true;
        }
        return false;
      };
      let done = false;
      if (tryPlace(part.w, part.h, false)) done = true;
      else if (allowRotation && tryPlace(part.h, part.w, true)) done = true;
      else {
        shelfY += shelf.rowH + kerf; shelf = { usedW: 0, rowH: 0 };
        if (shelfY - margin < usableH) {
          if (tryPlace(part.w, part.h, false)) done = true;
          else if (allowRotation && tryPlace(part.h, part.w, true)) done = true;
        }
      }
      if (!done) remaining.push(part);
    });
    if (placements.length === 0 && remaining.length > 0) {
      const p = remaining.shift();
      sheets.push([{ x: margin, y: margin, w: Math.min(p.w, usableW), h: Math.min(p.h, usableH), rotated: false, id: p.id, label: p.label }]);
    } else sheets.push(placements);
    queue = remaining;
  }
  return sheets;
}

const PART_COLORS = ["#3B82F6","#10B981","#F59E0B","#EF4444","#8B5CF6","#06B6D4","#F97316","#84CC16","#EC4899","#14B8A6","#6366F1","#D97706","#059669","#DC2626","#7C3AED","#0EA5E9","#65A30D","#DB2777"];
const DENSITY = 7850;

// Hot-rolled section prefixes to IGNORE (angles, beams, channels, tubes, etc.)
const HOTROLLED = /^(L|HEA?|HEB|HEM|IPE|IPN|UB|UC|UBP|PFC|RSA|RSJ|RHS|SHS|CHS|UPN|UPE|HD|HP|W|S|C|MC|HSS|BR|RD|UA|EA|ANGLE|CHANNEL|BEAM|TUBE|PIPE|ROUND|SECTION)/i;

// Parse a Tekla profile string → plate dims, or null if not a plate
function parsePlateProfile(profileRaw) {
  const profile = String(profileRaw || "").trim().toUpperCase();
  if (!profile) return null;
  // plate prefixes: PL, PLT, PLATE, FL, P  e.g. PLT20*270, PL8*127.3, 12X150
  const m = profile.match(/^(?:PLT|PLATE|PL|FL|P)\s*[-_]?\s*(\d+(?:\.\d+)?)\s*[*x×X]\s*(\d+(?:\.\d+)?)/);
  if (m) { const a = +m[1], b = +m[2]; return { thickness: Math.min(a, b), width: Math.max(a, b) }; }
  // bare "20*270" or "12x150" with no hot-rolled prefix
  if (!HOTROLLED.test(profile)) {
    const bare = profile.match(/^(\d+(?:\.\d+)?)\s*[*x×X]\s*(\d+(?:\.\d+)?)/);
    if (bare) { const a = +bare[1], b = +bare[2]; return { thickness: Math.min(a, b), width: Math.max(a, b) }; }
  }
  return null;
}

// ─── BRIGHT TEKLA-STYLE 3D HERO ────────────────────────────────────────────────
function SteelHero({ onStart }) {
  const canvasRef = useRef(null);
  const animRef = useRef(null);
  const mouseRef = useRef({ x: 0, y: 0, tx: 0, ty: 0 });

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    let t = 0, dpr = Math.min(window.devicePixelRatio || 1, 2);
    const resize = () => { dpr = Math.min(window.devicePixelRatio || 1, 2); canvas.width = canvas.offsetWidth * dpr; canvas.height = canvas.offsetHeight * dpr; ctx.setTransform(dpr, 0, 0, dpr, 0, 0); };
    resize(); window.addEventListener("resize", resize);
    const onMove = (e) => { const r = canvas.getBoundingClientRect(); mouseRef.current.tx = (e.clientX - r.left) / r.width - 0.5; mouseRef.current.ty = (e.clientY - r.top) / r.height - 0.5; };
    window.addEventListener("mousemove", onMove);

    // brushed steel texture (lighter, brighter base for Tekla look)
    const tex = document.createElement("canvas"); const TS = 512; tex.width = TS; tex.height = TS;
    const tcx = tex.getContext("2d");
    tcx.fillStyle = "#c2ccd6"; tcx.fillRect(0, 0, TS, TS);
    for (let i = 0; i < 2600; i++) {
      const y = Math.random() * TS, x = Math.random() * TS, len = 40 + Math.random() * 220;
      tcx.strokeStyle = Math.random() > 0.5 ? `rgba(255,255,255,${0.05 + Math.random() * 0.08})` : `rgba(90,105,120,${0.05 + Math.random() * 0.08})`;
      tcx.lineWidth = 0.5 + Math.random();
      tcx.beginPath(); tcx.moveTo(x, y); tcx.lineTo(x + len, y + (Math.random() - 0.5) * 1.2); tcx.stroke();
    }
    const pattern = ctx.createPattern(tex, "repeat");

    const project = (x, y, z, cx, cy, rx, ry) => {
      const fov = 1200, cosx = Math.cos(rx), sinx = Math.sin(rx), cosy = Math.cos(ry), siny = Math.sin(ry);
      const y1 = y * cosx - z * sinx, z1 = y * sinx + z * cosx;
      const x2 = x * cosy + z1 * siny, z2 = -x * siny + z1 * cosy;
      const s = fov / (fov + z2 + 460);
      return { x: cx + x2 * s, y: cy + y1 * s };
    };

    const render = () => {
      const W = canvas.offsetWidth, H = canvas.offsetHeight;
      mouseRef.current.x += (mouseRef.current.tx - mouseRef.current.x) * 0.05;
      mouseRef.current.y += (mouseRef.current.ty - mouseRef.current.y) * 0.05;
      const mx = mouseRef.current.x, my = mouseRef.current.y;
      ctx.clearRect(0, 0, W, H);

      // BRIGHT Tekla-style gradient background (light grey-blue, like the Tekla model viewer)
      const bg = ctx.createLinearGradient(0, 0, 0, H);
      bg.addColorStop(0, "#dfe8f0"); bg.addColorStop(0.5, "#c3d2df"); bg.addColorStop(1, "#9fb3c4");
      ctx.fillStyle = bg; ctx.fillRect(0, 0, W, H);

      // subtle floor grid (Tekla workplane feel)
      ctx.save();
      ctx.strokeStyle = "rgba(90,120,150,0.18)"; ctx.lineWidth = 1;
      const horizon = H * 0.62;
      for (let i = 0; i <= 22; i++) {
        const gx = (i / 22) * W;
        ctx.beginPath(); ctx.moveTo(gx, horizon); ctx.lineTo(W / 2 + (gx - W / 2) * 3.2, H); ctx.stroke();
      }
      for (let i = 1; i <= 9; i++) {
        const f = i / 9, gy = horizon + (H - horizon) * f * f;
        ctx.globalAlpha = 1 - f * 0.6;
        ctx.beginPath(); ctx.moveTo(0, gy); ctx.lineTo(W, gy); ctx.stroke();
      }
      ctx.restore();

      const cx = W / 2, cy = H * 0.5;
      const rx = -0.34 + my * 0.12 + Math.sin(t * 0.25) * 0.018;
      const ry = mx * 0.16 + Math.sin(t * 0.2) * 0.025;
      const PW = Math.min(W * 0.74, 720), PH = PW * 0.66, depth = 24;

      // soft contact shadow
      ctx.save();
      const sh = ctx.createRadialGradient(cx, cy + PH * 0.42, 0, cx, cy + PH * 0.42, PW * 0.7);
      sh.addColorStop(0, "rgba(40,60,80,0.4)"); sh.addColorStop(1, "rgba(40,60,80,0)");
      ctx.fillStyle = sh; ctx.beginPath(); ctx.ellipse(cx, cy + PH * 0.46, PW * 0.62, PH * 0.16, 0, 0, Math.PI * 2); ctx.fill();
      ctx.restore();

      const top = [project(-PW/2,-PH/2,depth,cx,cy,rx,ry),project(PW/2,-PH/2,depth,cx,cy,rx,ry),project(PW/2,PH/2,depth,cx,cy,rx,ry),project(-PW/2,PH/2,depth,cx,cy,rx,ry)];
      const bot = [project(-PW/2,-PH/2,0,cx,cy,rx,ry),project(PW/2,-PH/2,0,cx,cy,rx,ry),project(PW/2,PH/2,0,cx,cy,rx,ry),project(-PW/2,PH/2,0,cx,cy,rx,ry)];

      // edges
      [[0,1],[1,2],[2,3],[3,0]].forEach(([a,b]) => {
        ctx.beginPath(); ctx.moveTo(bot[a].x,bot[a].y); ctx.lineTo(bot[b].x,bot[b].y); ctx.lineTo(top[b].x,top[b].y); ctx.lineTo(top[a].x,top[a].y); ctx.closePath();
        const eg = ctx.createLinearGradient(bot[a].x,bot[a].y,top[b].x,top[b].y);
        eg.addColorStop(0,"#5d6b78"); eg.addColorStop(0.5,"#8593a0"); eg.addColorStop(1,"#49555f");
        ctx.fillStyle = eg; ctx.fill(); ctx.strokeStyle = "rgba(40,55,70,.7)"; ctx.lineWidth = 1; ctx.stroke();
      });

      // top textured face
      ctx.save();
      ctx.beginPath(); top.forEach((p,i)=>i===0?ctx.moveTo(p.x,p.y):ctx.lineTo(p.x,p.y)); ctx.closePath(); ctx.clip();
      ctx.save();
      ctx.translate(top[0].x, top[0].y);
      ctx.rotate(Math.atan2(top[1].y - top[0].y, top[1].x - top[0].x));
      const wpx = Math.hypot(top[1].x-top[0].x,top[1].y-top[0].y), hpx = Math.hypot(top[3].x-top[0].x,top[3].y-top[0].y);
      ctx.fillStyle = pattern || "#c2ccd6"; ctx.fillRect(-50, -50, wpx + 100, hpx + 200);
      ctx.restore();
      // bright moving sheen
      const sweep = Math.sin(t * 0.5) * 0.5 + 0.5;
      const shine = ctx.createLinearGradient(top[0].x,top[0].y,top[2].x,top[2].y);
      shine.addColorStop(Math.max(0,sweep-.28),"rgba(120,140,160,.15)");
      shine.addColorStop(sweep,"rgba(255,255,255,.65)");
      shine.addColorStop(Math.min(1,sweep+.28),"rgba(120,140,160,.15)");
      ctx.fillStyle = shine; ctx.globalCompositeOperation = "soft-light"; ctx.fillRect(0,0,W,H); ctx.globalCompositeOperation = "source-over";
      ctx.restore();

      // Tekla-style part outlines on plate (each a distinct coloured nested part)
      const gx = (u) => -PW/2 + PW * u, gy = (v) => -PH/2 + PH * v;
      const toS = (x, y) => project(x, y, depth + .5, cx, cy, rx, ry);
      ctx.save();
      ctx.beginPath(); top.forEach((p,i)=>i===0?ctx.moveTo(p.x,p.y):ctx.lineTo(p.x,p.y)); ctx.closePath(); ctx.clip();
      ctx.strokeStyle = "rgba(70,100,130,.18)"; ctx.lineWidth = .7;
      for (let i=1;i<10;i++){const a=toS(gx(i/10),gy(0)),b=toS(gx(i/10),gy(1));ctx.beginPath();ctx.moveTo(a.x,a.y);ctx.lineTo(b.x,b.y);ctx.stroke();}
      for (let i=1;i<7;i++){const a=toS(gx(0),gy(i/7)),b=toS(gx(1),gy(i/7));ctx.beginPath();ctx.moveTo(a.x,a.y);ctx.lineTo(b.x,b.y);ctx.stroke();}
      [[0.04,0.05,0.30,0.40,"#2563EB"],[0.04,0.50,0.30,0.44,"#059669"],[0.37,0.05,0.26,0.55,"#D97706"],[0.37,0.63,0.26,0.31,"#DB2777"],[0.66,0.05,0.30,0.30,"#7C3AED"],[0.66,0.38,0.30,0.26,"#0891B2"]].forEach(([u,v,w,h,col])=>{
        const c=[toS(gx(u),gy(v)),toS(gx(u+w),gy(v)),toS(gx(u+w),gy(v+h)),toS(gx(u),gy(v+h))];
        ctx.beginPath(); c.forEach((p,i)=>i===0?ctx.moveTo(p.x,p.y):ctx.lineTo(p.x,p.y)); ctx.closePath();
        ctx.fillStyle=col+"33"; ctx.fill(); ctx.strokeStyle=col; ctx.lineWidth=1.6; ctx.stroke();
      });
      // animated laser cut
      const lp = Math.sin(t*.6)*.5+.5;
      const la=toS(gx(lp),gy(0)),lb=toS(gx(lp),gy(1));
      ctx.save(); ctx.strokeStyle="#ff4d2e"; ctx.lineWidth=1.8; ctx.shadowColor="#ff6a3d"; ctx.shadowBlur=16;
      ctx.beginPath(); ctx.moveTo(la.x,la.y); ctx.lineTo(lb.x,lb.y); ctx.stroke();
      const spark=toS(gx(lp),gy(Math.sin(t*2)*.5+.5)); ctx.fillStyle="#fff"; ctx.beginPath(); ctx.arc(spark.x,spark.y,2.6,0,Math.PI*2); ctx.fill();
      ctx.restore();
      ctx.restore();

      // crisp bright top edge
      ctx.beginPath(); ctx.moveTo(top[0].x,top[0].y); ctx.lineTo(top[1].x,top[1].y); ctx.strokeStyle="rgba(255,255,255,.8)"; ctx.lineWidth=1.6; ctx.stroke();
      ctx.beginPath(); top.forEach((p,i)=>i===0?ctx.moveTo(p.x,p.y):ctx.lineTo(p.x,p.y)); ctx.closePath(); ctx.strokeStyle="rgba(50,80,110,.5)"; ctx.lineWidth=1; ctx.stroke();

      t += 0.016;
      animRef.current = requestAnimationFrame(render);
    };
    render();
    return () => { cancelAnimationFrame(animRef.current); window.removeEventListener("resize", resize); window.removeEventListener("mousemove", onMove); };
  }, []);

  return (
    <div style={{ position: "relative", width: "100%", height: "100vh", minHeight: 580, overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <canvas ref={canvasRef} style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }} />
      <div style={{ position: "absolute", inset: 0, background: "radial-gradient(ellipse 46% 40% at 50% 50%, rgba(8,22,36,0.55) 0%, rgba(8,22,36,0.32) 45%, rgba(8,22,36,0) 72%)", zIndex: 1, pointerEvents: "none" }} />
      <div style={{ position: "relative", zIndex: 2, textAlign: "center", padding: "0 24px", pointerEvents: "none" }}>
        <div style={{ display: "inline-block", background: "rgba(0,200,255,.14)", border: "1px solid rgba(120,220,255,.45)", borderRadius: 30, padding: "5px 18px", marginBottom: 22, fontFamily: "'Courier New', monospace", fontSize: 11, letterSpacing: 3, color: "#bfeaff", textTransform: "uppercase", backdropFilter: "blur(4px)" }}>Steel Fabrication Intelligence</div>
        <h1 style={{ margin: "0 0 16px", lineHeight: 1.05, fontFamily: "'Georgia', serif", fontWeight: 700, fontSize: "clamp(34px,5.5vw,68px)", color: "#f4fbff", textShadow: "0 2px 30px rgba(0,30,55,.85), 0 0 60px rgba(0,40,70,.5)" }}>Steel Plate<br /><span style={{ color: "#33d6ff" }}>Cut Optimizer</span></h1>
        <p style={{ margin: "0 0 38px", color: "#d6ecf6", fontSize: "clamp(14px,2vw,18px)", fontFamily: "'Georgia', serif", fontWeight: 600, textShadow: "0 2px 16px rgba(0,25,45,.8)" }}>Smart nesting &nbsp;•&nbsp; Minimum waste &nbsp;•&nbsp; Maximum efficiency</p>
        <button onClick={onStart} style={{ pointerEvents: "auto", background: "linear-gradient(135deg,#0a90d0,#0560a0)", border: "1px solid rgba(255,255,255,.4)", color: "#fff", padding: "15px 48px", borderRadius: 8, fontSize: 16, fontWeight: 600, cursor: "pointer", letterSpacing: 1, fontFamily: "'Georgia', serif", boxShadow: "0 10px 34px rgba(10,120,180,.4)", transition: "all .25s" }}
          onMouseEnter={e => { e.target.style.transform = "translateY(-2px)"; e.target.style.boxShadow = "0 14px 48px rgba(10,144,208,.6)"; }}
          onMouseLeave={e => { e.target.style.transform = "translateY(0)"; e.target.style.boxShadow = "0 10px 34px rgba(10,120,180,.4)"; }}>Start Optimizing →</button>
      </div>
      <div style={{ position: "absolute", bottom: 28, left: "50%", transform: "translateX(-50%)", color: "rgba(30,80,115,.6)", fontSize: 11, letterSpacing: 2, fontFamily: "'Courier New', monospace", textTransform: "uppercase", animation: "bob 2s ease-in-out infinite", zIndex: 2 }}>↓ scroll to workspace</div>
      <style>{`@keyframes bob{0%,100%{opacity:.4;transform:translate(-50%,0)}50%{opacity:1;transform:translate(-50%,6px)}}`}</style>
    </div>
  );
}

// ─── NESTING CANVAS ───────────────────────────────────────────────────────────
function NestingCanvas({ sheets, sheetW, sheetH, colorMap, thickness }) {
  const [active, setActive] = useState(0);
  const [zoom, setZoom] = useState(1);
  const canvasRef = useRef(null);
  useEffect(() => { setActive(0); }, [sheets]);
  const sheet = sheets[active] || [];
  const SCALE = Math.min(520 / sheetW, 360 / sheetH) * zoom;

  useEffect(() => {
    const canvas = canvasRef.current; if (!canvas) return;
    const ctx = canvas.getContext("2d");
    const W = sheetW * SCALE, H = sheetH * SCALE;
    canvas.width = W + 44; canvas.height = H + 44;
    ctx.fillStyle = "#0a141d"; ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.save(); ctx.translate(22, 22);
    ctx.fillStyle = "#16242f"; ctx.fillRect(0, 0, W, H);
    ctx.strokeStyle = "#3a6080"; ctx.lineWidth = 1.5; ctx.strokeRect(0, 0, W, H);
    ctx.strokeStyle = "rgba(60,100,140,.25)"; ctx.lineWidth = .5;
    const step = Math.max(20, Math.round(100 * SCALE));
    for (let x = step; x < W; x += step) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, H); ctx.stroke(); }
    for (let y = step; y < H; y += step) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(W, y); ctx.stroke(); }
    sheet.forEach((p, i) => {
      const px = p.x * SCALE, py = p.y * SCALE, pw = p.w * SCALE, ph = p.h * SCALE;
      const color = colorMap[p.id] || PART_COLORS[i % PART_COLORS.length];
      ctx.fillStyle = color + "55"; ctx.fillRect(px, py, pw, ph);
      ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.strokeRect(px, py, pw, ph);
      const fs = Math.max(8, Math.min(11, pw / 6, ph / 3));
      ctx.fillStyle = "rgba(255,255,255,.92)"; ctx.font = `${fs}px 'Courier New', monospace`; ctx.textAlign = "center"; ctx.textBaseline = "middle";
      if (pw > 22 && ph > 13) {
        ctx.fillText(p.rotated ? `${p.label}↺` : p.label, px + pw/2, py + ph/2 - (ph > 26 ? 6 : 0));
        if (ph > 26) { ctx.font = `${Math.max(7, fs - 1)}px 'Courier New', monospace`; ctx.fillStyle = "rgba(190,225,245,.75)"; ctx.fillText(`${Math.round(p.w)}×${Math.round(p.h)}`, px + pw/2, py + ph/2 + 7); }
      }
    });
    ctx.fillStyle = "rgba(110,185,225,.75)"; ctx.font = "10px 'Courier New', monospace"; ctx.textAlign = "center";
    ctx.fillText(`${sheetW} mm`, W / 2, H + 15);
    ctx.save(); ctx.translate(-9, H / 2); ctx.rotate(-Math.PI / 2); ctx.textAlign = "center"; ctx.fillText(`${sheetH} mm`, 0, 0); ctx.restore();
    ctx.restore();
  }, [sheet, sheetW, sheetH, SCALE, colorMap]);

  const used = sheet.reduce((s, p) => s + p.w * p.h, 0), total = sheetW * sheetH, util = Math.round((used / total) * 100);
  return (
    <div style={{ fontFamily: "'Courier New', monospace" }}>
      <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
        {sheets.map((sh, i) => { const pct = Math.round((sh.reduce((s, p) => s + p.w * p.h, 0) / total) * 100);
          return <button key={i} onClick={() => setActive(i)} style={{ padding: "5px 12px", borderRadius: 4, border: "1px solid", borderColor: active === i ? "#22cdf0" : "#2a4a5a", background: active === i ? "rgba(34,205,240,.15)" : "rgba(12,28,42,.8)", color: active === i ? "#22cdf0" : "#6a9ab0", fontSize: 11, cursor: "pointer" }}>Sheet {i + 1} · {pct}%</button>; })}
        <div style={{ marginLeft: "auto", display: "flex", gap: 6 }}>
          <button onClick={() => setZoom(z => Math.min(z + .2, 3))} style={zoomBtn}>+</button>
          <button onClick={() => setZoom(z => Math.max(z - .2, .4))} style={zoomBtn}>−</button>
        </div>
      </div>
      <div style={{ overflow: "auto", borderRadius: 6, border: "1px solid #1a3a4a", maxHeight: 440 }}><canvas ref={canvasRef} style={{ display: "block" }} /></div>
      <div style={{ marginTop: 10, display: "flex", gap: 18, flexWrap: "wrap", padding: "10px 14px", background: "rgba(0,20,35,.6)", borderRadius: 6, border: "1px solid #1a3a4a" }}>
        <MiniStat label="Sheet" value={`${active + 1} / ${sheets.length}`} />
        <MiniStat label="Thickness" value={`${thickness} mm`} />
        <MiniStat label="Parts" value={sheet.length} />
        <MiniStat label="Utilization" value={`${util}%`} accent={util > 80 ? "#10B981" : util > 60 ? "#F59E0B" : "#EF4444"} />
        <MiniStat label="Waste Area" value={`${Math.round((total - used) / 1000)} dm²`} />
      </div>
    </div>
  );
}
const zoomBtn = { width: 28, height: 28, borderRadius: 4, border: "1px solid #2a4a5a", background: "rgba(12,28,42,.8)", color: "#6a9ab0", cursor: "pointer", fontSize: 16 };
function MiniStat({ label, value, accent }) { return (<div style={{ textAlign: "center" }}><div style={{ fontSize: 10, color: "#4a7a90", marginBottom: 2, letterSpacing: 1 }}>{label}</div><div style={{ fontSize: 14, color: accent || "#c0dde8", fontWeight: 700 }}>{value}</div></div>); }

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const workspaceRef = useRef(null);
  const [phase, setPhase] = useState("input");
  const [inputMode, setInputMode] = useState(null);
  // STOCK LIST: each row = one available thickness with its stock sheet size.
  // This drives everything: manual entry can only pick these thicknesses,
  // and Excel import looks up the matching sheet size by thickness.
  const DEFAULT_SHEET = { w: 1220, h: 2440 }; // fallback when no stock size given
  const [stock, setStock] = useState([
    { thickness: 8, w: 1220, h: 2440 },
    { thickness: 12, w: 1220, h: 2440 },
    { thickness: 20, w: 1500, h: 3000 },
  ]);
  const [parts, setParts] = useState([
    { id: "P1", label: "P1", length: 500, width: 300, thickness: 12, qty: 4 },
    { id: "P2", label: "P2", length: 380, width: 250, thickness: 12, qty: 6 },
    { id: "P3", label: "P3", length: 600, width: 400, thickness: 20, qty: 3 },
    { id: "P4", label: "P4", length: 280, width: 180, thickness: 8, qty: 8 },
  ]);
  const [material, setMaterial] = useState("S235JR");
  const kerf = 3; // hidden cutting-gap assumption, suitable for most fabrication shops
  const [margin, setMargin] = useState(10);
  const [allowRotation, setAllowRotation] = useState(true);
  const [results, setResults] = useState(null);
  const [colorMap, setColorMap] = useState({});
  const [excelMsg, setExcelMsg] = useState("");
  const [columnMap, setColumnMap] = useState(null);
  const [isOptimizing, setIsOptimizing] = useState(false);

  const scrollToWorkspace = () => workspaceRef.current?.scrollIntoView({ behavior: "smooth" });

  // stock helpers
  const stockThicknesses = [...new Set(stock.map(s => s.thickness).filter(t => t > 0))].sort((a, b) => a - b);
  const sheetForThickness = (thk) => {
    const hit = stock.find(s => Number(s.thickness) === Number(thk));
    return hit ? { w: hit.w, h: hit.h } : DEFAULT_SHEET;
  };
  const addStock = () => setStock(s => [...s, { thickness: 0, w: 1220, h: 2440 }]);
  const updateStock = (i, f, v) => setStock(s => s.map((r, ri) => ri === i ? { ...r, [f]: Number(v) || 0 } : r));
  const deleteStock = (i) => setStock(s => s.filter((_, ri) => ri !== i));

  const addPart = () => { const id = `P${parts.length + 1}`; const thk = stockThicknesses[0] || 10; setParts(p => [...p, { id, label: id, length: 400, width: 200, thickness: thk, qty: 1 }]); };
  const updatePart = (i, f, v) => setParts(p => p.map((r, ri) => ri === i ? { ...r, [f]: f === "label" ? v : Number(v) || 0 } : r));
  const deletePart = (i) => setParts(p => p.filter((_, ri) => ri !== i));

  const detectColumn = (headers, aliases) => {
    for (const h of headers) for (const a of aliases) if (h.toLowerCase().trim() === a.toLowerCase()) return h;
    for (const h of headers) for (const a of aliases) if (h.toLowerCase().includes(a.toLowerCase())) return h;
    return null;
  };

  const parseExcel = useCallback((file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        // find header row (looks for Profile / Length / No. etc.)
        let hr = -1;
        for (let i = 0; i < Math.min(raw.length, 25); i++) {
          const r = raw[i].map(c => String(c).trim().toLowerCase());
          if (r.some(c => /profile|length|width|qty|no\.?$|thick|partpos|part pos/i.test(c))) { hr = i; break; }
        }
        if (hr === -1) hr = 0;
        const headers = raw[hr].map(c => String(c).trim());
        const rows = raw.slice(hr + 1).filter(r => r.some(c => c !== ""));
        const map = {
          profile: detectColumn(headers, ["profile", "section", "section size", "size"]),
          qty: detectColumn(headers, ["no.", "no", "qty", "number", "count", "quantity", "pcs"]),
          length: detectColumn(headers, ["length (mm)", "length(mm)", "length", "l (mm)", "l", "height"]),
          width: detectColumn(headers, ["width (mm)", "width(mm)", "width", "w", "breadth"]),
          thickness: detectColumn(headers, ["thickness", "thk", "t (mm)"]),
        };
        const out = extractParts(headers, rows, map);
        if (out.length) { setParts(out); setExcelMsg(`✓ Imported ${out.length} plate part${out.length > 1 ? "s" : ""} (hot-rolled sections ignored)`); setColumnMap(null); }
        else { setColumnMap({ headers, map, rows }); setExcelMsg("⚠ Couldn't auto-detect. Tell me which columns to use."); }
      } catch { setExcelMsg("⚠ Could not read this file. Please check the format."); }
    };
    reader.readAsBinaryString(file);
  }, []);

  // Robust extraction: understands Tekla (Profile column) AND messy free-form sheets.
  function extractParts(headers, rows, map) {
    const out = [];
    const idx = (col) => col ? headers.indexOf(col) : -1;
    const pi = idx(map.profile), qi = idx(map.qty), li = idx(map.length), wi = idx(map.width), ti = idx(map.thickness);

    rows.forEach((row, ri) => {
      const cells = row.map(c => String(c || "").trim());
      const joined = cells.join(" ");
      // skip total / summary lines
      if (/total\s+for|members|grand total/i.test(joined)) return;

      let thickness = 0, width = 0, length = 0, qty = 0;

      // 1) Tekla profile column (e.g. PLT20*270) — also ignores hot-rolled
      const profileCell = pi >= 0 ? cells[pi] : (cells.find(c => /^(PLT?|PLATE|FL|L|HE|IPE|UB|UC|RHS|CHS|SHS)/i.test(c)) || "");
      if (profileCell) {
        if (/^(L|HE|IPE|IPN|UB|UC|PFC|RHS|SHS|CHS|UPN|UPE|BR|RD|HD|HP|UA|EA)/i.test(profileCell) && !/^PL/i.test(profileCell)) return; // hot-rolled → skip
        const plate = parsePlateProfile(profileCell);
        if (plate) { thickness = plate.thickness; width = plate.width; }
      }

      // 2) length from its column, else scan numbers
      if (li >= 0) length = parseFloat(cells[li]) || 0;
      if (qi >= 0) qty = parseInt(cells[qi]) || 0;
      if (ti >= 0 && !thickness) thickness = parseFloat(cells[ti]) || 0;
      if (wi >= 0 && !width) width = parseFloat(cells[wi]) || 0;

      // 3) FALLBACK PARSER — messy / badly entered sheet:
      // pull all plausible numbers from the row and infer.
      if (!(length > 0 && width > 0)) {
        const nums = cells.flatMap(c => {
          // catch "20*270", "12x150" embedded too
          const pair = c.match(/(\d+(?:\.\d+)?)\s*[*x×X]\s*(\d+(?:\.\d+)?)/);
          if (pair) return [+pair[1], +pair[2]];
          const n = parseFloat(c.replace(/[^\d.]/g, ""));
          return isFinite(n) && c.match(/\d/) ? [n] : [];
        }).filter(n => n > 0 && n < 100000);

        if (nums.length >= 2) {
          // heuristic: smallest plausible = thickness (3–60mm), then width & length are the two largest
          const sorted = [...nums].sort((a, b) => a - b);
          if (!thickness) {
            const thkCand = sorted.find(n => n >= 3 && n <= 80);
            thickness = thkCand || 10;
          }
          const big = sorted.filter(n => n !== thickness);
          if (big.length >= 2) { width = width || big[0]; length = length || big[big.length - 1]; }
          else if (big.length === 1) { length = length || big[0]; width = width || big[0]; }
        }
      }

      if (!qty) {
        // qty: a small integer-looking standalone cell
        const qCell = cells.find(c => /^\d{1,4}$/.test(c) && +c <= 999 && +c !== length && +c !== width && +c !== thickness);
        qty = qCell ? +qCell : 1;
      }
      if (!thickness) thickness = 10;

      if (length > 0 && width > 0) out.push({ id: `P${ri + 1}`, label: `P${ri + 1}`, length: Math.round(length * 10) / 10, width: Math.round(width * 10) / 10, thickness, qty: qty || 1 });
    });
    return out;
  }

  const onDrop = useCallback((e) => { e.preventDefault(); const file = e.dataTransfer?.files?.[0] || e.target?.files?.[0]; if (file) parseExcel(file); }, [parseExcel]);

  const optimize = () => {
    setIsOptimizing(true);
    setTimeout(() => {
      const valid = parts.filter(p => p.length > 0 && p.width > 0 && p.qty > 0);
      const cm = {}; valid.forEach((p, i) => { cm[p.id] = PART_COLORS[i % PART_COLORS.length]; }); setColorMap(cm);
      const byThk = {}; valid.forEach(p => { const k = p.thickness || 10; (byThk[k] = byThk[k] || []).push(p); });
      const groups = Object.keys(byThk).map(Number).sort((a, b) => b - a).map(thk => {
        const gp = byThk[thk];
        const { w: sw, h: sh } = sheetForThickness(thk); // each thickness uses its own stock sheet
        const gsheets = nestParts(sw, sh, gp, kerf, margin, allowRotation);
        const usedArea = gsheets.reduce((s, st) => s + st.reduce((a, p) => a + p.w * p.h, 0), 0);
        const totalArea = gsheets.length * sw * sh;
        const wastePct = Math.round(((totalArea - usedArea) / totalArea) * 100);
        const partVol = gp.reduce((s, p) => s + p.length * p.width * thk * p.qty, 0);
        const partWeight = (partVol / 1e9) * DENSITY;
        const sheetWeight = (gsheets.length * sw * sh * thk / 1e9) * DENSITY;
        return { thickness: thk, sw, sh, sheets: gsheets, sheetCount: gsheets.length, partCount: gp.reduce((s, p) => s + p.qty, 0), usedArea, totalArea, wastePct, utilPct: 100 - wastePct, partWeight, sheetWeight, wasteWeight: sheetWeight - partWeight };
      });
      const totals = groups.reduce((a, g) => ({ sheets: a.sheets + g.sheetCount, parts: a.parts + g.partCount, sheetWeight: a.sheetWeight + g.sheetWeight, partWeight: a.partWeight + g.partWeight, wasteWeight: a.wasteWeight + g.wasteWeight, totalArea: a.totalArea + g.totalArea, usedArea: a.usedArea + g.usedArea }), { sheets: 0, parts: 0, sheetWeight: 0, partWeight: 0, wasteWeight: 0, totalArea: 0, usedArea: 0 });
      totals.utilPct = Math.round((totals.usedArea / totals.totalArea) * 100);
      totals.wastePct = 100 - totals.utilPct;
      setResults({ groups, totals });
      setPhase("results"); setIsOptimizing(false);
      setTimeout(scrollToWorkspace, 100);
    }, 500);
  };

  const fmtTon = (kg) => (kg / 1000).toFixed(kg / 1000 >= 100 ? 0 : 2);

  return (
    <div style={{ background: "#070d13", minHeight: "100vh", color: "#c0dde8", fontFamily: "'Georgia', serif" }}>
      <SteelHero onStart={scrollToWorkspace} />
      <div ref={workspaceRef} style={{ maxWidth: 1180, margin: "0 auto", padding: "44px 20px 70px" }}>

        {phase === "input" && (<>
          <SectionTitle>Optimization Workspace</SectionTitle>

          {inputMode === null && (
            <div>
              <div style={{ textAlign: "center", color: "#7ab3c8", fontSize: 15, marginBottom: 18 }}>How would you like to enter your cutting list?</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
                <ChooserCard icon="✏️" title="Manual Entry" desc="Type parts directly into a simple table. Best for small lists." onClick={() => setInputMode("manual")} />
                <ChooserCard icon="📊" title="Upload Excel / CSV" desc="Drop a Tekla report or any cutting schedule. Reads plates, ignores beams/angles, even fixes messy sheets." onClick={() => setInputMode("excel")} />
              </div>
            </div>
          )}

          {inputMode !== null && (
          <Card title="📐 Stock & Cutting Settings">
            <Label>Available Stock — add each thickness with its sheet size</Label>
            <div style={{ overflowX: "auto", marginTop: 6 }}>
              <table style={{ borderCollapse: "collapse", fontSize: 13 }}>
                <thead><tr style={{ borderBottom: "1px solid #1a3a4a" }}>
                  {["Thickness (mm)", "Width (mm)", "Length (mm)", ""].map(h => <th key={h} style={{ textAlign: "left", padding: "6px 10px", color: "#4a8a9a", fontWeight: 600, fontSize: 11, letterSpacing: 1 }}>{h}</th>)}
                </tr></thead>
                <tbody>
                  {stock.map((s, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #0d2030" }}>
                      <td style={{ padding: "5px 10px" }}><input type="number" value={s.thickness} onChange={e => updateStock(i, "thickness", e.target.value)} style={{ ...cellInput, width: 90, borderColor: "#2a6a8a" }} /></td>
                      <td style={{ padding: "5px 10px" }}><input type="number" value={s.w} onChange={e => updateStock(i, "w", e.target.value)} style={{ ...cellInput, width: 90 }} /></td>
                      <td style={{ padding: "5px 10px" }}><input type="number" value={s.h} onChange={e => updateStock(i, "h", e.target.value)} style={{ ...cellInput, width: 90 }} /></td>
                      <td style={{ padding: "5px 10px" }}><button onClick={() => deleteStock(i)} style={{ background: "none", border: "none", color: "#4a6a7a", cursor: "pointer", fontSize: 14 }}>✕</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <button onClick={addStock} style={{ marginTop: 10, padding: "7px 16px", background: "rgba(0,100,150,.3)", border: "1px dashed #2a6a8a", color: "#6ab8cc", borderRadius: 4, cursor: "pointer", fontSize: 12 }}>+ Add Thickness</button>

            <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end", marginTop: 18, paddingTop: 16, borderTop: "1px solid #12303f" }}>
              <div><Label>Material Grade</Label>
                <select value={material} onChange={e => setMaterial(e.target.value)} style={selectStyle}>{["S235JR","S275","S355","S420","ST37","A36","A572"].map(m => <option key={m}>{m}</option>)}</select>
              </div>
              <div><Label>Edge Margin (mm)</Label><input type="number" value={margin} onChange={e => setMargin(+e.target.value)} style={{ ...inputStyle, width: 80 }} /></div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, paddingBottom: 6 }}>
                <input type="checkbox" id="rot" checked={allowRotation} onChange={e => setAllowRotation(e.target.checked)} style={{ width: 16, height: 16, accentColor: "#22cdf0" }} />
                <label htmlFor="rot" style={{ color: "#8ab8cc", fontSize: 13 }}>Allow 90° Rotation</label>
              </div>
            </div>
          </Card>
          )}

          {inputMode === "manual" && (
            <Card title="✏️ Manual Part Entry" style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} />}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead><tr style={{ borderBottom: "1px solid #1a3a4a" }}>{["ID", "Length (mm)", "Width (mm)", "Thickness (mm)", "Qty", ""].map(h => <th key={h} style={{ textAlign: "left", padding: "6px 8px", color: "#4a8a9a", fontWeight: 600, fontSize: 11, letterSpacing: 1 }}>{h}</th>)}</tr></thead>
                  <tbody>
                    {parts.map((p, i) => (
                      <tr key={i} style={{ borderBottom: "1px solid #0d2030" }}>
                        <td style={{ padding: "5px 4px" }}><input value={p.label} onChange={e => updatePart(i, "label", e.target.value)} style={{ ...cellInput, width: 50 }} /></td>
                        <td style={{ padding: "5px 4px" }}><input type="number" value={p.length} onChange={e => updatePart(i, "length", e.target.value)} style={{ ...cellInput, width: 80 }} /></td>
                        <td style={{ padding: "5px 4px" }}><input type="number" value={p.width} onChange={e => updatePart(i, "width", e.target.value)} style={{ ...cellInput, width: 80 }} /></td>
                        <td style={{ padding: "5px 4px" }}>
                          <select value={p.thickness} onChange={e => updatePart(i, "thickness", e.target.value)} style={{ ...cellInput, width: 90, borderColor: "#2a6a8a", cursor: "pointer" }}>
                            {!stockThicknesses.includes(p.thickness) && <option value={p.thickness}>{p.thickness} (not in stock)</option>}
                            {stockThicknesses.map(t => <option key={t} value={t}>{t} mm</option>)}
                          </select>
                        </td>
                        <td style={{ padding: "5px 4px" }}><input type="number" value={p.qty} onChange={e => updatePart(i, "qty", e.target.value)} style={{ ...cellInput, width: 55 }} /></td>
                        <td style={{ padding: "5px 4px" }}><button onClick={() => deletePart(i)} style={{ background: "none", border: "none", color: "#4a6a7a", cursor: "pointer", fontSize: 14 }}>✕</button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <button onClick={addPart} style={{ marginTop: 12, padding: "8px 18px", background: "rgba(0,100,150,.3)", border: "1px dashed #2a6a8a", color: "#6ab8cc", borderRadius: 4, cursor: "pointer", fontSize: 12, width: "100%" }}>+ Add Part</button>
            </Card>
          )}

          {inputMode === "excel" && (
            <Card title="📊 Upload Excel / CSV" style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} />}>
              <div style={{ marginBottom: 14, padding: "10px 14px", borderRadius: 6, background: "rgba(0,40,60,.4)", border: "1px solid rgba(34,205,240,.25)", color: "#9cc4d6", fontSize: 12, lineHeight: 1.6 }}>
                💡 <b>Tip:</b> add your thicknesses &amp; stock sheet sizes in the settings above. The <b>thickness always comes from the uploaded file</b>; the matching stock sheet size is taken from your list. Any thickness not listed falls back to <b>1220 × 2440 mm</b>.
                {stockThicknesses.length > 0 && <div style={{ marginTop: 4, color: "#6ab8cc" }}>Stock sizes on file: {stock.map(s => `${s.thickness}mm → ${s.w}×${s.h}`).join("  ·  ")}</div>}
              </div>
              <div onDrop={onDrop} onDragOver={e => e.preventDefault()} onDragEnter={e => e.currentTarget.style.borderColor = "#22cdf0"} onDragLeave={e => e.currentTarget.style.borderColor = "#2a5a7a"} onClick={() => document.getElementById("xlsxInput").click()} style={{ border: "2px dashed #2a5a7a", borderRadius: 8, padding: "48px 20px", textAlign: "center", background: "rgba(0,40,60,.3)", cursor: "pointer", transition: "border-color .2s" }}>
                <div style={{ fontSize: 44, marginBottom: 12 }}>📋</div>
                <div style={{ color: "#5a9ab0", fontSize: 15, marginBottom: 6 }}>Drop Tekla report / Excel / CSV here</div>
                <div style={{ color: "#3a6a7a", fontSize: 12 }}>or click to browse</div>
                <div style={{ marginTop: 14, color: "#2a5a6a", fontSize: 10, lineHeight: 1.7 }}>Reads <b>Profile</b> column (PLT20*270 → 20mm thick, 270 wide)<br />Hot-rolled sections (L, HEA, IPE, UB…) ignored automatically<br />Messy / hand-typed sheets are auto-cleaned</div>
              </div>
              <input id="xlsxInput" type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={onDrop} />
              {excelMsg && <div style={{ marginTop: 10, padding: "8px 12px", borderRadius: 4, background: excelMsg.startsWith("✓") ? "rgba(16,185,129,.15)" : "rgba(245,158,11,.15)", border: `1px solid ${excelMsg.startsWith("✓") ? "rgba(16,185,129,.4)" : "rgba(245,158,11,.4)"}`, color: excelMsg.startsWith("✓") ? "#10B981" : "#F59E0B", fontSize: 12 }}>{excelMsg}</div>}
              {columnMap && (
                <div style={{ marginTop: 12, padding: 12, background: "rgba(0,20,35,.8)", borderRadius: 6, border: "1px solid #2a4a5a" }}>
                  <div style={{ fontSize: 12, color: "#8ab8cc", marginBottom: 8 }}>Map your columns:</div>
                  {["profile", "qty", "length", "width", "thickness"].map(f => (
                    <div key={f} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                      <span style={{ fontSize: 11, color: "#4a7a90", width: 70 }}>{f}</span>
                      <select value={columnMap.map[f] || ""} onChange={e => setColumnMap(c => ({ ...c, map: { ...c.map, [f]: e.target.value } }))} style={{ ...selectStyle, flex: 1 }}><option value="">—</option>{columnMap.headers.map(h => <option key={h} value={h}>{h}</option>)}</select>
                    </div>
                  ))}
                  <button onClick={() => { const out = extractParts(columnMap.headers, columnMap.rows, columnMap.map); if (out.length) { setParts(out); setExcelMsg(`✓ Imported ${out.length} plate parts`); setColumnMap(null); } else setExcelMsg("⚠ Still couldn't extract. Check the mapping."); }} style={{ ...btnStyle, marginTop: 6 }}>Apply Mapping</button>
                </div>
              )}
              {parts.length > 0 && (
                <div style={{ marginTop: 14 }}>
                  <div style={{ fontSize: 11, color: "#3a6a7a", marginBottom: 6, letterSpacing: 1 }}>PARTS READY ({parts.length})</div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                    {parts.slice(0, 16).map((p, i) => <div key={i} style={{ padding: "3px 8px", borderRadius: 3, background: (colorMap[p.id] || PART_COLORS[i % PART_COLORS.length]) + "22", border: `1px solid ${(colorMap[p.id] || PART_COLORS[i % PART_COLORS.length])}55`, fontSize: 10, color: "#8ab8cc", fontFamily: "'Courier New', monospace" }}>{p.label}: {p.length}×{p.width}×{p.thickness} ×{p.qty}</div>)}
                    {parts.length > 16 && <div style={{ fontSize: 10, color: "#3a6a7a", alignSelf: "center" }}>+{parts.length - 16} more</div>}
                  </div>
                </div>
              )}
            </Card>
          )}

          {inputMode !== null && (
            <div style={{ textAlign: "center", marginTop: 32 }}>
              <button onClick={optimize} disabled={isOptimizing} style={{ padding: "16px 60px", fontSize: 16, fontWeight: 700, background: isOptimizing ? "rgba(0,80,120,.4)" : "linear-gradient(135deg,#00a8d6,#0072a8)", border: "1px solid rgba(34,205,240,.5)", color: "#fff", borderRadius: 8, cursor: isOptimizing ? "default" : "pointer", boxShadow: "0 8px 30px rgba(0,150,210,.4)", letterSpacing: 2, fontFamily: "'Georgia', serif" }}>{isOptimizing ? "⚙ Optimizing…" : "⚡ OPTIMIZE CUTTING LAYOUT"}</button>
            </div>
          )}
        </>)}

        {phase === "results" && results && (<>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24, flexWrap: "wrap", gap: 12 }}>
            <SectionTitle>Optimization Results</SectionTitle>
            <button onClick={() => setPhase("input")} style={{ ...btnStyle, background: "rgba(0,60,90,.4)" }}>← Back to Input</button>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))", gap: 12, marginBottom: 24 }}>
            {[
              { l: "Total Parts", v: results.totals.parts, i: "◼" },
              { l: "Total Sheets", v: results.totals.sheets, i: "📄", a: "#3B82F6" },
              { l: "Thickness Groups", v: results.groups.length, i: "≡", a: "#8B5CF6" },
              { l: "Utilization", v: `${results.totals.utilPct}%`, i: "📊", a: results.totals.utilPct > 80 ? "#10B981" : results.totals.utilPct > 60 ? "#F59E0B" : "#EF4444" },
              { l: "Total Purchase Weight", v: `${fmtTon(results.totals.sheetWeight)} t`, i: "🧾", a: "#22cdf0" },
              { l: "Net Parts Weight", v: `${fmtTon(results.totals.partWeight)} t`, i: "⚖" },
              { l: "Scrap / Offcut Weight", v: `${fmtTon(results.totals.wasteWeight)} t`, i: "🗑", a: "#EF4444" },
            ].map(c => (
              <div key={c.l} style={{ padding: "16px 14px", borderRadius: 8, background: "rgba(0,20,35,.8)", border: "1px solid #1a3a4a", textAlign: "center" }}>
                <div style={{ fontSize: 20, marginBottom: 4 }}>{c.i}</div>
                <div style={{ fontSize: 22, fontWeight: 700, color: c.a || "#c0dde8", fontFamily: "'Courier New', monospace" }}>{c.v}</div>
                <div style={{ fontSize: 10, color: "#4a7a90", marginTop: 4, letterSpacing: 1 }}>{c.l}</div>
              </div>
            ))}
          </div>

          {results.groups.map((g, gi) => (
            <Card key={gi} title={`🗂 Nesting Layout — ${g.thickness} mm Plate (${g.sheetCount} sheet${g.sheetCount > 1 ? "s" : ""})`} style={{ marginBottom: 20 }}>
              <NestingCanvas sheets={g.sheets} sheetW={g.sw} sheetH={g.sh} colorMap={colorMap} thickness={g.thickness} />
            </Card>
          ))}

          <Card title="💾 Export" style={{ marginBottom: 28 }}>
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
              <button onClick={() => exportPDF(results, material, fmtTon)} style={exportBtn}>📄 Export PDF Report</button>
              <button onClick={() => exportExcel(results, parts, material)} style={exportBtn}>📊 Export Excel Summary</button>
            </div>
          </Card>

          <div style={{ border: "2px solid rgba(34,205,240,.3)", borderRadius: 12, overflow: "hidden", boxShadow: "0 0 60px rgba(0,180,255,.08)" }}>
            <div style={{ background: "linear-gradient(135deg,rgba(0,100,150,.5),rgba(0,60,100,.5))", borderBottom: "1px solid rgba(34,205,240,.3)", padding: "20px 28px", display: "flex", alignItems: "center", gap: 14 }}>
              <span style={{ fontSize: 28 }}>📦</span>
              <div>
                <div style={{ fontSize: 11, letterSpacing: 3, color: "#5cc8e8", textTransform: "uppercase", marginBottom: 3, fontFamily: "'Courier New', monospace" }}>Final Decision-Ready Output</div>
                <div style={{ fontSize: 22, fontWeight: 700, color: "#eef6fa" }}>Material Procurement Summary</div>
              </div>
              <div style={{ marginLeft: "auto", textAlign: "right" }}>
                <div style={{ fontSize: 10, color: "#4a8a9a", letterSpacing: 1, fontFamily: "'Courier New', monospace" }}>MATERIAL GRADE</div>
                <div style={{ fontSize: 16, color: "#22cdf0", fontWeight: 700 }}>{material}</div>
              </div>
            </div>
            <div style={{ padding: "26px 28px", background: "rgba(0,12,24,.92)" }}>

              {/* PLAIN-LANGUAGE BUY LIST */}
              <div style={{ marginBottom: 22, padding: "18px 20px", background: "rgba(0,40,60,.4)", border: "1px solid rgba(34,205,240,.3)", borderRadius: 10 }}>
                <div style={{ fontSize: 11, letterSpacing: 2, color: "#5cc8e8", textTransform: "uppercase", marginBottom: 12, fontFamily: "'Courier New', monospace" }}>What you need to buy</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {results.groups.map((g, gi) => (
                    <div key={gi} style={{ display: "flex", alignItems: "center", gap: 12, fontSize: 17, color: "#eef6fa", fontFamily: "'Georgia', serif" }}>
                      <span style={{ color: "#22cdf0", fontSize: 18 }}>✓</span>
                      <span>You need <b style={{ color: "#7cb6f8", fontSize: 20 }}>{g.sheetCount}</b> sheet{g.sheetCount > 1 ? "s" : ""} of <b style={{ color: "#22cdf0" }}>{g.thickness} mm</b> &nbsp;<span style={{ color: "#8ab8cc", fontSize: 14, fontFamily: "'Courier New', monospace" }}>({g.sw} × {g.sh} mm, {material})</span></span>
                    </div>
                  ))}
                </div>
              </div>

              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 14, fontFamily: "'Courier New', monospace" }}>
                  <thead>
                    <tr style={{ background: "rgba(0,60,90,.45)", borderBottom: "2px solid rgba(34,205,240,.25)" }}>
                      {["THICKNESS", "SHEET SIZE", "SHEETS REQ.", "PURCHASE WT (full sheets)", "PARTS", "UTILIZATION", "NET PARTS WT", "SCRAP WT"].map((h, i) => <th key={h} style={{ padding: "13px 14px", textAlign: i < 2 ? "left" : i === 7 ? "right" : "center", color: "#5cc8e8", fontSize: 11, letterSpacing: 1, fontWeight: 700 }}>{h}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {results.groups.map((g, gi) => (
                      <tr key={gi} style={{ borderBottom: "1px solid rgba(0,60,90,.4)" }}>
                        <td style={{ padding: "15px 14px" }}><span style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "rgba(34,205,240,.1)", border: "1px solid rgba(34,205,240,.3)", borderRadius: 4, padding: "4px 12px" }}><span style={{ color: "#22cdf0" }}>✓</span><span style={{ color: "#dceef5", fontWeight: 700, fontSize: 15 }}>{g.thickness} mm</span></span></td>
                        <td style={{ padding: "15px 14px", color: "#8ab8cc" }}>{g.sw} × {g.sh} mm</td>
                        <td style={{ padding: "15px 14px", textAlign: "center" }}><span style={{ display: "inline-block", background: "linear-gradient(135deg,rgba(59,130,246,.22),rgba(59,130,246,.1))", border: "1px solid rgba(59,130,246,.4)", borderRadius: 6, padding: "6px 16px", color: "#7cb6f8", fontWeight: 800, fontSize: 18 }}>{g.sheetCount} sheet{g.sheetCount > 1 ? "s" : ""}</span></td>
                        <td style={{ padding: "15px 14px", textAlign: "center", color: "#22cdf0", fontWeight: 700 }}>{fmtTon(g.sheetWeight)} t</td>
                        <td style={{ padding: "15px 14px", textAlign: "center", color: "#9cc4d6" }}>{g.partCount}</td>
                        <td style={{ padding: "15px 14px", textAlign: "center", fontWeight: 700, color: g.utilPct > 80 ? "#10B981" : g.utilPct > 60 ? "#F59E0B" : "#EF4444" }}>{g.utilPct}%</td>
                        <td style={{ padding: "15px 14px", textAlign: "center", color: "#9cc4d6" }}>{fmtTon(g.partWeight)} t</td>
                        <td style={{ padding: "15px 14px", textAlign: "right", color: "#f08a8a" }}>{fmtTon(g.wasteWeight)} t</td>
                      </tr>
                    ))}
                    <tr style={{ background: "rgba(0,70,105,.4)", borderTop: "2px solid rgba(34,205,240,.4)" }}>
                      <td style={{ padding: "16px 14px", fontWeight: 800, color: "#eef6fa" }}>TOTAL TO BUY</td>
                      <td colSpan={2} style={{ padding: "16px 14px", color: "#dceef5", fontSize: 13, lineHeight: 1.6 }}>
                        {results.groups.map((g, gi) => (
                          <span key={gi} style={{ display: "inline-block", marginRight: 14, whiteSpace: "nowrap" }}>
                            <b style={{ color: "#7cb6f8", fontSize: 16 }}>{g.sheetCount}</b> × <b style={{ color: "#22cdf0" }}>{g.thickness}mm</b> <span style={{ color: "#6a9ab0" }}>({g.sw}×{g.sh})</span>{gi < results.groups.length - 1 ? "  +" : ""}
                          </span>
                        ))}
                        <div style={{ marginTop: 4, color: "#9cc4d6", fontWeight: 700 }}>= {results.totals.sheets} sheet{results.totals.sheets > 1 ? "s" : ""} total</div>
                      </td>
                      <td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 800, color: "#22cdf0" }}>{fmtTon(results.totals.sheetWeight)} t</td>
                      <td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 700, color: "#dceef5" }}>{results.totals.parts}</td>
                      <td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 800, color: results.totals.utilPct > 80 ? "#10B981" : results.totals.utilPct > 60 ? "#F59E0B" : "#EF4444" }}>{results.totals.utilPct}%</td>
                      <td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 800, color: "#dceef5" }}>{fmtTon(results.totals.partWeight)} t</td>
                      <td style={{ padding: "16px 14px", textAlign: "right", fontWeight: 800, color: "#f08a8a" }}>{fmtTon(results.totals.wasteWeight)} t</td>
                    </tr>
                  </tbody>
                </table>
              </div>

              <div style={{ marginTop: 12, padding: "10px 14px", background: "rgba(0,20,35,.5)", border: "1px solid #12303f", borderRadius: 6, fontSize: 11, color: "#7ab3c8", lineHeight: 1.7, fontFamily: "'Courier New', monospace" }}>
                <span style={{ color: "#22cdf0" }}>● Purchase Wt</span> = gross weight of the full sheets you buy &nbsp;·&nbsp;
                <span style={{ color: "#10B981" }}>● Net Parts Wt</span> = weight of the finished cut plates only &nbsp;·&nbsp;
                <span style={{ color: "#f08a8a" }}>● Scrap Wt</span> = offcut left after cutting (Purchase − Net)
              </div>

              {/* TOTAL PURCHASE TONNAGE — the headline procurement number */}
              <div style={{ marginTop: 22, display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 16 }}>
                <BigTon label="Total Purchase Weight" value={fmtTon(results.totals.sheetWeight)} note={`gross — all ${results.totals.sheets} full sheets you buy`} color="#22cdf0" big />
                <BigTon label="Net Parts Weight" value={fmtTon(results.totals.partWeight)} note="net — finished cut plates only" color="#10B981" />
                <BigTon label="Scrap / Offcut Weight" value={fmtTon(results.totals.wasteWeight)} note={`leftover — ${results.totals.wastePct}% of purchase`} color="#EF4444" />
              </div>

              <div style={{ marginTop: 18, textAlign: "right", fontFamily: "'Courier New', monospace", fontSize: 10, color: "#2a5a6a" }}>Generated by Steel Plate Cut Optimizer · {new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })} · {material}</div>
            </div>
          </div>
        </>)}
      </div>
    </div>
  );
}

function BigTon({ label, value, note, color, big }) {
  return (
    <div style={{ padding: "20px 22px", borderRadius: 10, background: big ? `linear-gradient(135deg, ${color}22, ${color}0a)` : "rgba(0,15,30,.8)", border: `1px solid ${color}55`, textAlign: "center" }}>
      <div style={{ fontSize: 10, color: "#5a8a9a", letterSpacing: 1.5, textTransform: "uppercase", marginBottom: 8, fontFamily: "'Courier New', monospace" }}>{label}</div>
      <div style={{ fontFamily: "'Courier New', monospace", fontWeight: 800, color, fontSize: big ? 44 : 32, lineHeight: 1 }}>{value}<span style={{ fontSize: big ? 18 : 14, marginLeft: 4 }}>Ton</span></div>
      <div style={{ fontSize: 11, color: "#4a7a90", marginTop: 8 }}>{note}</div>
    </div>
  );
}

function ChooserCard({ icon, title, desc, onClick }) {
  return (
    <button onClick={onClick} style={{ textAlign: "left", padding: "26px 24px", borderRadius: 10, background: "rgba(0,18,32,.8)", border: "1px solid #1a3a4a", cursor: "pointer", transition: "all .2s", fontFamily: "'Georgia', serif", color: "#c0dde8" }}
      onMouseEnter={e => { e.currentTarget.style.borderColor = "#22cdf0"; e.currentTarget.style.background = "rgba(0,40,65,.8)"; e.currentTarget.style.transform = "translateY(-3px)"; }}
      onMouseLeave={e => { e.currentTarget.style.borderColor = "#1a3a4a"; e.currentTarget.style.background = "rgba(0,18,32,.8)"; e.currentTarget.style.transform = "translateY(0)"; }}>
      <div style={{ fontSize: 38, marginBottom: 12 }}>{icon}</div>
      <div style={{ fontSize: 18, fontWeight: 700, color: "#eef6fa", marginBottom: 8 }}>{title}</div>
      <div style={{ fontSize: 13, color: "#7ab3c8", lineHeight: 1.5 }}>{desc}</div>
      <div style={{ marginTop: 14, color: "#22cdf0", fontSize: 13, fontWeight: 600 }}>Choose →</div>
    </button>
  );
}
function BackChip({ onClick }) { return <button onClick={onClick} style={{ padding: "4px 12px", borderRadius: 20, border: "1px solid #2a5a7a", background: "rgba(0,40,60,.5)", color: "#6ab8cc", cursor: "pointer", fontSize: 11 }}>↺ Change method</button>; }
function SectionTitle({ children }) { return <h2 style={{ fontSize: 22, fontWeight: 700, color: "#eef6fa", margin: "0 0 20px", fontFamily: "'Georgia', serif", borderLeft: "3px solid #22cdf0", paddingLeft: 14 }}>{children}</h2>; }
function Card({ title, children, style, action }) {
  return (
    <div style={{ background: "rgba(0,15,30,.8)", border: "1px solid #1a3a4a", borderRadius: 8, overflow: "hidden", ...style }}>
      <div style={{ padding: "12px 18px", borderBottom: "1px solid #1a3a4a", background: "rgba(0,30,50,.5)", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <span style={{ fontSize: 13, fontWeight: 600, color: "#8ab8cc", letterSpacing: .5 }}>{title}</span>{action}
      </div>
      <div style={{ padding: 18 }}>{children}</div>
    </div>
  );
}
function Label({ children }) { return <div style={{ fontSize: 10, color: "#4a8a9a", letterSpacing: 1, textTransform: "uppercase", marginBottom: 5, fontFamily: "'Courier New', monospace" }}>{children}</div>; }
const inputStyle = { background: "rgba(0,20,35,.8)", border: "1px solid #1a4a6a", color: "#c0dde8", padding: "7px 10px", borderRadius: 4, fontSize: 13, width: 110, fontFamily: "'Courier New', monospace", outline: "none" };
const selectStyle = { background: "rgba(0,20,35,.8)", border: "1px solid #1a4a6a", color: "#c0dde8", padding: "7px 10px", borderRadius: 4, fontSize: 13, fontFamily: "'Courier New', monospace", cursor: "pointer", outline: "none" };
const cellInput = { background: "rgba(0,20,35,.6)", border: "1px solid #1a3a4a", color: "#c0dde8", padding: "5px 7px", borderRadius: 3, fontSize: 12, fontFamily: "'Courier New', monospace", outline: "none" };
const btnStyle = { padding: "8px 18px", borderRadius: 4, border: "1px solid #2a5a7a", background: "rgba(0,60,90,.5)", color: "#6ab8cc", cursor: "pointer", fontSize: 12 };
const exportBtn = { padding: "10px 22px", borderRadius: 5, border: "1px solid #2a5a7a", background: "rgba(0,60,100,.4)", color: "#8ab8cc", cursor: "pointer", fontSize: 13, fontFamily: "'Georgia', serif" };

function exportExcel(results, parts, material) {
  const wb = XLSX.utils.book_new();
  const rows = [
    ["Steel Plate Cut Optimizer — Procurement Summary"], [],
    ["Material", material], [],
    ["Thickness (mm)", "Sheet Size", "Sheets Req.", "Purchase Wt - full sheets (t)", "Parts", "Utilization %", "Net Parts Wt (t)", "Scrap Wt (t)"],
    ...results.groups.map(g => [g.thickness, `${g.sw}×${g.sh}`, `${g.sheetCount} sheets`, (g.sheetWeight / 1000).toFixed(2), g.partCount, g.utilPct, (g.partWeight / 1000).toFixed(2), (g.wasteWeight / 1000).toFixed(2)]),
    ["TOTAL TO BUY", results.groups.map(g => `${g.sheetCount}x${g.thickness}mm`).join(" + "), `${results.totals.sheets} sheets`, (results.totals.sheetWeight / 1000).toFixed(2), results.totals.parts, results.totals.utilPct, (results.totals.partWeight / 1000).toFixed(2), (results.totals.wasteWeight / 1000).toFixed(2)],
    [], ["TOTAL PURCHASE WEIGHT - gross, full sheets (Ton)", (results.totals.sheetWeight / 1000).toFixed(2)],
    ["NET PARTS WEIGHT - finished plates only (Ton)", (results.totals.partWeight / 1000).toFixed(2)],
    ["SCRAP / OFFCUT WEIGHT (Ton)", (results.totals.wasteWeight / 1000).toFixed(2)],
    [], ["Parts List"], ["ID", "Length", "Width", "Thickness", "Qty"],
    ...parts.map(p => [p.label, p.length, p.width, p.thickness, p.qty]),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Summary");
  XLSX.writeFile(wb, "SteelCutOptimizer_Report.xlsx");
}
function exportPDF(results, material, fmtTon) {
  const w = window.open("", "_blank");
  w.document.write(`<html><head><title>Steel Plate Cut Optimizer Report</title><style>
  body{font-family:Georgia,serif;padding:32px;color:#0a1a26}
  h1{color:#0080a8;border-bottom:2px solid #cde;padding-bottom:10px}
  table{width:100%;border-collapse:collapse;margin-top:14px;font-family:monospace;font-size:13px}
  th{background:#e8f4f8;color:#0080a8;padding:10px;text-align:left;font-size:11px}
  td{padding:10px;border-bottom:1px solid #dde}
  tr.total td{background:#eef6fa;font-weight:bold;border-top:2px solid #0080a8}
  .purchase{margin-top:20px;padding:18px;background:#eef8fc;border:2px solid #0080a8;border-radius:8px;font-size:18px}
  .purchase b{font-size:30px;color:#0080a8}
  .buylist{margin:14px 0;padding:16px 18px;background:#f3fafd;border:1px solid #b6dcea;border-radius:8px}
  .buylist h3{margin:0 0 10px;color:#0080a8;font-size:13px;letter-spacing:1px;text-transform:uppercase}
  .buylist .row{font-size:16px;margin:6px 0}
  .buylist .n{font-weight:bold;color:#0a66a0;font-size:19px}
  </style></head><body>
  <h1>Steel Plate Cut Optimizer — Procurement Summary</h1>
  <p>Material: <b>${material}</b> &nbsp;·&nbsp; ${new Date().toLocaleDateString("en-GB")}</p>
  <div class="buylist"><h3>What you need to buy</h3>
  ${results.groups.map(g => `<div class="row">✓ You need <span class="n">${g.sheetCount}</span> sheet${g.sheetCount > 1 ? "s" : ""} of <b>${g.thickness} mm</b> (${g.sw}×${g.sh} mm, ${material})</div>`).join("")}
  </div>
  <table><tr><th>Thickness</th><th>Sheet Size</th><th>Sheets Req.</th><th>Purchase Wt (full sheets)</th><th>Parts</th><th>Utilization</th><th>Net Parts Wt</th><th>Scrap Wt</th></tr>
  ${results.groups.map(g => `<tr><td>${g.thickness} mm</td><td>${g.sw}×${g.sh}</td><td>${g.sheetCount} sheet${g.sheetCount > 1 ? "s" : ""}</td><td>${fmtTon(g.sheetWeight)} t</td><td>${g.partCount}</td><td>${g.utilPct}%</td><td>${fmtTon(g.partWeight)} t</td><td>${fmtTon(g.wasteWeight)} t</td></tr>`).join("")}
  <tr class="total"><td>TOTAL TO BUY</td><td>${results.groups.map(g => `${g.sheetCount}×${g.thickness}mm`).join(" + ")}</td><td>${results.totals.sheets} sheets</td><td>${fmtTon(results.totals.sheetWeight)} t</td><td>${results.totals.parts}</td><td>${results.totals.utilPct}%</td><td>${fmtTon(results.totals.partWeight)} t</td><td>${fmtTon(results.totals.wasteWeight)} t</td></tr>
  </table>
  <p style="font-size:12px;color:#456">Purchase Wt = gross weight of full sheets bought &nbsp;·&nbsp; Net Parts Wt = finished cut plates only &nbsp;·&nbsp; Scrap Wt = offcut (Purchase − Net)</p>
  <div class="purchase">Total Purchase Weight (gross): <b>${fmtTon(results.totals.sheetWeight)} Ton</b> &nbsp; (${results.totals.sheets} sheets) &nbsp;·&nbsp; Net parts: ${fmtTon(results.totals.partWeight)} Ton &nbsp;·&nbsp; Scrap: ${fmtTon(results.totals.wasteWeight)} Ton (${results.totals.wastePct}%)</div>
  <script>window.print()</script></body></html>`);
}
