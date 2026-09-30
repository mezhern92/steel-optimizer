import { Component, useState, useRef, useEffect, useCallback, useMemo, createContext, useContext } from "react";
import * as XLSX from "xlsx";

/* ============================================================================
   STEEL OPTIMIZER — bilingual layer (English / العربية), one-click RTL switch
   ----------------------------------------------------------------------------
   DROP-IN. Paste this whole block near the TOP of your single-file app, right
   after your `import` lines. Then make the 4 small edits flagged "◀ WIRE" in
   the integration notes. Arabic terms are the ones Gulf/Saudi contractors &
   fabricators actually use (صاج، كمر، حصر الكميات، الهدر، الهالك، بواقي، عود …)
   — verified against Saudi steel-supplier & contractor pages, not raw MT.
   Numerals stay Western (12000, 1220×2440, S355) — correct for the trade.
============================================================================ */

/* ╔══════════════════════════════════════════════════════════════════════════╗
   ║  ANALYTICS — the funnel. Paste ONE id below and every event starts.       ║
   ╚══════════════════════════════════════════════════════════════════════════╝

   Google Analytics 4 is used because custom events are free and unlimited
   there. Vercel's own custom events need a paid Pro plan, so the same events
   are ALSO sent to Vercel automatically if you ever upgrade — no code change.

   SETUP (5 minutes, free, once):
     1. analytics.google.com → Admin → Create property → name it Steel Optimizer
     2. Choose platform "Web", enter steeloptimizer.com
     3. Copy the Measurement ID — it looks like  G-XXXXXXXXXX
     4. Paste it below. Upload. Done.

   Leave it empty and nothing is loaded, nothing is sent, nothing breaks.

   WHAT YOU WILL SEE (GA4 → Reports → Engagement → Events):
     page_view              how many arrived
     upload_excel           how many actually tried it        ◄ the real number
     optimization_started   how many pressed Optimize
     optimization_completed whether the optimizer succeeded
     download_results       whether the result was worth keeping
         
   The gap between page_view and upload_excel is the only number that tells
   you whether the product is understood. Everything else is downstream.
──────────────────────────────────────────────────────────────────────────── */

const GA_MEASUREMENT_ID = "";        // ◄── paste G-XXXXXXXXXX here

let _gaLoaded = false;
function loadGA() {
  if (_gaLoaded || !GA_MEASUREMENT_ID || typeof window === "undefined") return;
  _gaLoaded = true;
  const sc = document.createElement("script");
  sc.async = true;
  sc.src = "https://www.googletagmanager.com/gtag/js?id=" + GA_MEASUREMENT_ID;
  document.head.appendChild(sc);
  window.dataLayer = window.dataLayer || [];
  window.gtag = function () { window.dataLayer.push(arguments); };
  window.gtag("js", new Date());
  // anonymize_ip keeps this clean under GDPR-style rules without a cookie banner
  window.gtag("config", GA_MEASUREMENT_ID, { anonymize_ip: true });
}

/* ── INSTANT ALERT ────────────────────────────────────────────────────────
   GA4 reports lag by hours and you have to go looking. This emails you the
   moment somebody actually uses the tool, which at your volume is the signal
   you are waiting for. Set to false once it becomes noise.

   PRIVACY — this is a promise the product makes on screen, so it is kept in
   code: NO file content and NO file name are ever sent. A file name can carry
   a client or project name, which is not yours to transmit. Only the fact that
   an upload happened, the file type, and the row count.                      */
const ALERT_ME_ON_USE = true;
const ALERT_FORM_ID   = "xqerwkze";        // your existing Formspree form

let _alertCount = 0;

/* ── MANUAL-USE ALERT ──────────────────────────────────────────────────────
   One email for every press of Optimize after MANUAL entry (typed rows, not
   an Excel/PDF upload), in both modules, so you can see whether people use
   the tool by hand. Same privacy rule as above: counts only — no sizes, no
   profiles, no names. "demo: true" means the plate list was the untouched
   sample list. The 20-per-session cap above still applies.                  */
const MANUAL_ALERT_ON = true;                 // ◄ false = stop the manual-use emails
let _soManualPress = 0;
function soNotifyManual(module, detail) {
  if (!MANUAL_ALERT_ON) return;
  try {
    _soManualPress++;
    alertMe(`MANUAL OPTIMIZE (${module}) #${_soManualPress}`, { ...detail, press: _soManualPress, lang: typeof SO_LANG === "string" ? SO_LANG : "" });
  } catch { /* never block the optimizer */ }
}
const SO_DEMO_PLATES = "500x300x12x4|380x250x12x6|600x400x20x3|280x180x8x8";
function soIsDemoPlates(parts) {
  try { return (parts || []).map(p => `${+p.length}x${+p.width}x${+p.thickness}x${+p.qty}`).join("|") === SO_DEMO_PLATES; } catch { return false; }
}
const SO_BUILD = "2026-09-30";   // sent with every usage alert → shows which version is live
function alertMe(what, detail) {
  if (!ALERT_ME_ON_USE || typeof window === "undefined") return;
  if (_alertCount >= 20) return;                       // cap per session
  _alertCount++;
  try {
    fetch("https://formspree.io/f/" + ALERT_FORM_ID, {
      method: "POST",
      headers: { "Content-Type": "application/json", Accept: "application/json" },
      body: JSON.stringify({
        email: "usage@steeloptimizer.com",             // so Formspree accepts it
        _subject: "Steel Optimizer — " + what,
        event: what,
        detail: detail || {},
        referrer: document.referrer || "direct",
        language: navigator.language || "",
        screen: window.innerWidth + "x" + window.innerHeight,
        at: new Date().toISOString(),
        ver: SO_BUILD,
      }),
    }).catch(() => {});
  } catch { /* never block the app */ }
}

/* One call site for every event. Never throws — analytics must never be able
   to break the optimizer. Fires to GA4 and, if present, Vercel Analytics. */
function track(event, props) {
  try {
    const p = props || {};
    if (typeof window !== "undefined") {
      if (window.gtag && GA_MEASUREMENT_ID) window.gtag("event", event, p);
      if (window.va) window.va("event", { name: event, data: p });   // Vercel, if on Pro
      if (!GA_MEASUREMENT_ID && window.console) console.debug("[track]", event, p);
    }
  } catch { /* analytics is never load-bearing */ }
}

const LANG_DICT = {
  // ---- global chrome ----
  appName:        { en: "Steel Cut & Nest Optimizer", ar: "برنامج توفيق القطعيات وتحسين قص الحديد" },
  langName:       { en: "English",                    ar: "العربية" },
  // toggle button shows the OTHER language's name (by design ar→"English")
  otherLang:      { en: "العربية",                    ar: "English" },

  // ---- hero ----
  heroBadge:      { en: "To Reduce Steel Waste", ar: "لتقليل هدر الحديد" },
  heroTitleA:     { en: "Steel Cut", ar: "برنامج توفيق القطعيات" },
  heroTitleB:     { en: "& Nest Optimizer", ar: "وتحسين قص الحديد" },
  heroSub:        { en: "Plates & Sections • Minimum waste • Instant tonnage",
                    ar: "ألواح ومقاطع • أقل هدر • وزن بالطن فوري" },
  heroCta:        { en: "START OPTIMIZING →",          ar: "ابدأ التحسين →" },
  heroScroll:     { en: "↓ scroll to start",           ar: "↓ مرّر للبدء" },

  // ---- live showcase ----
  showTitle:      { en: "Watch how plate cutting is optimized", ar: "شاهد كيف يُحسَّن تقطيع اللوح" },
  chipImport:     { en: "Tekla · Excel · CSV import",  ar: "استيراد Tekla · Excel · CSV" },
  chipDb:         { en: "1,100+ sections — EU · US · JIS · GOST · GB",
                    ar: "+1100 مقطع — أوروبي · أمريكي · ياباني · روسي · صيني" },
  chipOffcut:     { en: "L-shaped offcut register",    ar: "سجل قصاصات على شكل L" },
  chipTon:        { en: "Instant tonnage & buy list",  ar: "وزن بالطن وقائمة شراء فورية" },
  chipReports:    { en: "PDF & Excel reports",         ar: "تقارير PDF و Excel" },

  // ---- module chooser ----
  chooseWhat:     { en: "Choose what to optimize",     ar: "اختر ما تريد تحسينه" },

  // ---- contact footer ----
  ctTop:          { en: "File not reading correctly? Email it to us",
                    ar: "الملف لم يُقرأ بشكل صحيح؟ أرسله لنا" },
  ctTitle:        { en: "File not imported correctly?", ar: "الملف لم يُستورد بشكل صحيح؟" },
  ctBody:         { en: "Send us the Excel file and we will run it and send the cutting plan back. Free.",
                    ar: "أرسل لنا ملف الإكسل وسنشغّله ونعيد لك خطة القص. مجانًا." },
  ctEmail:        { en: "Email",                       ar: "البريد الإلكتروني" },
  ctWhats:        { en: "WhatsApp",                    ar: "واتساب" },
  ctWhatsHint:    { en: "Fastest — send a photo of the list if that is easier",
                    ar: "الأسرع — أرسل صورة للقائمة إن كان أسهل" },
  ctBy:           { en: "Built by a structural engineer, for people who cut steel.",
                    ar: "من إعداد مهندس إنشائي، لمن يقصّون الحديد." },
  ctBlocked:      { en: "Site blocked on your office network? Some corporate firewalls block new domains. Open it on mobile data, or message me and I will send the result.",
                    ar: "الموقع محجوب على شبكة شركتك؟ بعض جدران الحماية تحجب الدومينات الجديدة. افتحه من بيانات الجوال، أو راسلني وأرسل لك النتيجة." },
  preNoticeTtl:   { en: "Free while in early access",  ar: "مجاني خلال فترة الإطلاق المبكر" },
  preNoticeBody:  { en: "Steel Optimizer moves to a paid subscription in the next few months. Everything you use today stays free until then, and early users keep a founding price.",
                    ar: "سيتحول Steel Optimizer إلى اشتراك مدفوع خلال الأشهر القادمة. كل ما تستخدمه اليوم يبقى مجانيًا حتى ذلك الحين، ومستخدمو البداية يحتفظون بسعر التأسيس." },
  preNoticeCta:   { en: "Tell me when it launches",    ar: "أبلغني عند الإطلاق" },
  preNoticeSent:  { en: "✓ Noted — you'll hear from us first.", ar: "✓ تم — ستصلك أول رسالة عند الإطلاق." },
  preNoticeMail:  { en: "your@email.com",              ar: "بريدك الإلكتروني" },

  // ---- Pro subscription gate (Lemon Squeezy) ----
  proTitle:       { en: "Exports are a Pro feature",   ar: "التصدير ميزة Pro" },
  proBody:        { en: "Your nesting layout, waste percentage and offcut register stay free on screen. Pro adds the downloadable PDF and Excel cutting reports.",
                    ar: "توزيع القص ونسبة الهدر وسجل البواقي تبقى مجانية على الشاشة. Pro يضيف تقارير القص القابلة للتنزيل بصيغة PDF و Excel." },
  proMonthly:     { en: "Monthly",                     ar: "شهري" },
  proYearly:      { en: "Yearly",                      ar: "سنوي" },
  proPerMo:       { en: "/mo",                         ar: "/شهر" },
  proPerYr:       { en: "/yr",                         ar: "/سنة" },
  proTrial:       { en: "First month free",            ar: "الشهر الأول مجانًا" },
  proSave:        { en: "2 months free",               ar: "شهران مجانًا" },
  proHaveKey:     { en: "Already have a key",          ar: "لديك مفتاح بالفعل" },
  proKeyPh:       { en: "Paste your licence key",      ar: "الصق مفتاح الترخيص" },
  proActivate:    { en: "Activate",                    ar: "تفعيل" },
  proChecking:    { en: "Checking…",                   ar: "جارٍ التحقق…" },
  proPrivacy:     { en: "Your files never leave your browser. Nothing is uploaded to a server.",
                    ar: "ملفاتك لا تغادر متصفحك. لا يُرفع أي شيء إلى أي خادم." },
  proErrEmpty:    { en: "Enter your licence key.",     ar: "أدخل مفتاح الترخيص." },
  proErrNet:      { en: "Could not reach the licence server. Check your connection.",
                    ar: "تعذّر الوصول إلى خادم التراخيص. تحقّق من الاتصال." },
  proErrLimit:    { en: "This key is already active on the maximum number of devices.",
                    ar: "هذا المفتاح مُفعّل بالفعل على الحد الأقصى من الأجهزة." },
  proErrExpired:  { en: "This subscription has ended. Renew to restore exports.",
                    ar: "انتهى هذا الاشتراك. جدّد لاستعادة التصدير." },
  proErrBad:      { en: "That key was not recognised. Check for missing characters.",
                    ar: "لم يتم التعرّف على المفتاح. تأكد من عدم نقص أحرف." },
  proOk:          { en: "✓ Pro active — exports unlocked.", ar: "✓ Pro مُفعّل — تم فتح التصدير." },
  proPay:         { en: "Subscribe",                   ar: "اشترك" },
  proLoading:     { en: "Opening secure checkout…",    ar: "جارٍ فتح صفحة الدفع الآمنة…" },
  proCards:       { en: "Visa · Mastercard · Amex · Apple Pay · Google Pay · PayPal",
                    ar: "‏Visa · Mastercard · Amex · Apple Pay · Google Pay · PayPal" },
  proSecure:      { en: "Secure checkout by Lemon Squeezy. Card details never touch this site.",
                    ar: "دفع آمن عبر Lemon Squeezy. بيانات البطاقة لا تمر عبر هذا الموقع إطلاقًا." },
  proRestore:     { en: "Paid on another device?",     ar: "دفعت من جهاز آخر؟" },
  proThanks:      { en: "Payment received — exports are unlocked. Thank you.",
                    ar: "تم استلام الدفعة — التصدير مفتوح الآن. شكرًا لك." },
  modSections:    { en: "Steel Sections",              ar: "المقاطع الحديدية" },
  modSectionsDesc:{ en: "Beams, columns, angles, Z & C cold-formed sections, and built-up sections.",
                    ar: "كمرات وأعمدة وزوايا ومقاطع C و Z المشكّلة على البارد والمقاطع المركّبة." },
  modPlates:      { en: "Steel Plates & Sheets",       ar: "ألواح وصاج الصلب" },
  modPlatesDesc:  { en: "Plate parts nested onto stock sheets (2D). Reuse leftover offcuts first, then minimum new sheets, with total tonnage.",
                    ar: "توزيع قطع الألواح على ألواح المخزون (ثنائي الأبعاد). استخدام القصاصات المتبقية أولًا ثم أقل عدد ألواح جديدة، مع إجمالي الوزن بالطن." },
  chooseArrow:    { en: "Choose →",                    ar: "اختر →" },

  // ---- shared input method ----
  howEnter:       { en: "How would you like to enter your cutting list?",
                    ar: "كيف تريد إدخال قائمة القص؟" },
  manualEntry:    { en: "Manual Entry",                ar: "إدخال يدوي" },
  uploadTitle:    { en: "Upload your material list",   ar: "ارفع قائمة المواد" },
  scanTitle:      { en: "Scan a PDF or photo",         ar: "مسح ملف PDF أو صورة" },
  scanDescP:      { en: "Read plate sizes from a PDF list, or a photo of a printed cutting list.",
                    ar: "اقرأ مقاسات الألواح من قائمة PDF أو صورة لقائمة قص مطبوعة." },
  scanDescS:      { en: "Read profiles and lengths from a PDF list, or a photo of a printed list.",
                    ar: "اقرأ المقاطع والأطوال من قائمة PDF أو صورة لقائمة مطبوعة." },
  manualDescP:    { en: "Type parts into a simple table.", ar: "اكتب القطع في جدول بسيط." },
  manualDescS:    { en: "Type a profile (IPE, UB, HEA, SHS…) and pick from the live list; enter lengths.",
                    ar: "اكتب المقطع (IPE، UB، HEA، SHS…) واختر من القائمة الحيّة؛ ثم أدخل الأطوال." },
  uploadDescP:    { en: "Drop a material list. Plates read, sections ignored, messy sheets cleaned.",
                    ar: "أفلت قائمة المواد. تُقرأ الألواح وتُتجاهل المقاطع، وتُنظَّف الملفات غير المرتبة." },
  uploadDescS:    { en: "Drop a material list. Sections detected; plates ignored.",
                    ar: "أفلت قائمة المواد. تُكتشف المقاطع وتُتجاهل الألواح." },
  changeMethod:   { en: "↺ Change method",             ar: "↺ تغيير الطريقة" },
  backModules:    { en: "↩ Back", ar: "↩ رجوع" },
  backInput:      { en: "← Back to input",             ar: "→ العودة للإدخال" },

  // ---- generic field labels ----
  lblThickness:   { en: "Thickness (mm)",              ar: "السماكة (مم)" },
  lblWidth:       { en: "Width (mm)",                  ar: "العرض (مم)" },
  lblLength:      { en: "Length (mm)",                 ar: "الطول (مم)" },
  lblQty:         { en: "Qty",                         ar: "العدد" },
  lblQtyHow:      { en: "How many? (qty)",             ar: "كم العدد؟" },
  lblGrade:       { en: "Material Grade",              ar: "رتبة المعدن" },
  lblKerf:        { en: "Cutting Gap / Kerf (mm)",     ar: "عرض القص / سماحية المنشار (مم)" },
  lblMargin:      { en: "Edge Margin (mm)",            ar: "هامش الحافة (مم)" },
  lblReuseMin:    { en: "Reusable Offcut Min (mm)",    ar: "أقل قصاصة قابلة لإعادة الاستخدام (مم)" },
  lblRotate:      { en: "Allow 90° Rotation",          ar: "السماح بالتدوير 90°" },
  lblId:          { en: "ID",                          ar: "الرمز" },
  lblProfile:     { en: "Profile",                     ar: "المقطع" },
  lblGradeShort:  { en: "Grade",                       ar: "الرتبة" },

  // ---- plates workspace ----
  platesWS:       { en: "Steel Plates — Workspace",    ar: "ألواح الصلب — مساحة العمل" },
  stockSettings:  { en: "📐 Stock & Cutting Settings",  ar: "📐 إعدادات المخزون والقص" },
  stockSheets:    { en: "Standard stock sheets — set a sheet size per thickness",
                    ar: "ألواح المخزون القياسية — حدّد مقاس اللوح لكل سماكة" },
  addThickness:   { en: "+ Add Thickness",             ar: "+ إضافة سماكة" },
  stockHint:      { en: "If parts are longer than 2440 mm, set a bigger sheet here (e.g. 1500×3000 or 2000×6000) for that thickness — otherwise those parts can't be nested.",
                    ar: "إذا كانت القطع أطول من 2440 مم، حدّد لوحًا أكبر هنا (مثل 1500×3000 أو 2000×6000) لتلك السماكة — وإلا تعذّر توزيع تلك القطع." },
  spliceStrat:    { en: "Splice strategy (parts bigger than sheet)", ar: "أسلوب الوصل (القطع الأكبر من اللوح)" },
  spliceWelds:    { en: "Fewest welds (bigger pieces)", ar: "أقل لحامات (قطع أكبر)" },
  splicePack:     { en: "Best material use (more, smaller pieces)", ar: "أفضل استغلال للمادة (قطع أصغر وأكثر)" },

  // ---- reuse leftovers (shared) ----
  reusePlatesTtl: { en: "♻ Reuse leftover plates (optional)", ar: "♻ إعادة استخدام ألواح متبقية (اختياري)" },
  reuseBarsTtl:   { en: "♻ Reuse leftover bars (optional)",   ar: "♻ إعادة استخدام أعواد متبقية (اختياري)" },
  reusePlatesLead:{ en: "Got leftover plates from a past job — or from another site? Add them here (or import a previous Leftover file) and we'll cut from them first, so you only buy the minimum new sheets you still need.",
                    ar: "لديك ألواح متبقية من مشروع سابق — أو من موقع آخر؟ أضفها هنا (أو استورد ملف بواقٍ سابق) وسنقصّ منها أولًا، لتشتري أقل عدد من الألواح الجديدة التي تحتاجها فعلًا." },
  reuseBarsLead:  { en: "Got leftover bars from a past job — or from another site? Add them here (or import a previous Leftover file) and we'll cut from them first, so you only buy the minimum new bars you still need.",
                    ar: "لديك أعواد متبقية من مشروع سابق — أو من موقع آخر؟ أضفها هنا (أو استورد ملف بواقٍ سابق) وسنقصّ منها أولًا، لتشتري أقل عدد من الأعواد الجديدة التي تحتاجها فعلًا." },
  noPlatesYet:    { en: "No leftover plates added yet.", ar: "لم تُضف ألواح متبقية بعد." },
  noBarsYet:      { en: "No leftover bars added yet.",   ar: "لم تُضف أعواد متبقية بعد." },
  addLeftPlate:   { en: "+ Add a leftover plate",       ar: "+ إضافة لوح متبقٍ" },
  addLeftBar:     { en: "+ Add a leftover bar",         ar: "+ إضافة عود متبقٍ" },
  reuseImportBtn: { en: "⤓ Reuse a previous Leftover file (PDF / Excel)",
                    ar: "⤓ إعادة استخدام ملف بواقٍ سابق (PDF / Excel)" },
  reuseReading:   { en: "Reading",                      ar: "جارٍ القراءة" },
  reuseBusy:      { en: "⏳ Reading…",                  ar: "⏳ جارٍ القراءة…" },

  // ---- manual plate table ----
  manualPartEntry:{ en: "✏️ Manual Part Entry",         ar: "✏️ إدخال القطع يدويًا" },
  addPart:        { en: "+ Add Part",                  ar: "+ إضافة قطعة" },

  // ---- upload screen ----
  uploadTeklaTtl: { en: "📊 Excel", ar: "📊 Excel" },
  dropTekla:      { en: "Drop Tekla list / Excel / CSV / TXT", ar: "أفلت قائمة Tekla / Excel / CSV / TXT" },
  orBrowse:       { en: "or click to browse",          ar: "أو انقر للتصفّح" },

  // ---- scan screen (PDF / photo) ----
  scanCardTtl:    { en: "🔍 PDF / Photo",              ar: "🔍 PDF / صورة" },
  scanDrop:       { en: "Drop a PDF or a photo of your list",
                    ar: "أفلت ملف PDF أو صورة لقائمتك" },
  scanAccept:     { en: "Accepts .pdf · .png · .jpg · .webp — or click to browse",
                    ar: "يقبل ‎.pdf · .png · .jpg · .webp — أو انقر للتصفّح" },
  scanWarn:       { en: "⚠ Reading a PDF or a photo is a best-effort guess, not a measurement. Check every row below before you optimise — a wrong number here means wrongly cut steel.",
                    ar: "⚠ القراءة من PDF أو صورة تقديرية وليست قياسًا. راجع كل صف بالأسفل قبل التحسين — رقم خاطئ هنا يعني حديدًا مقطوعًا خطأ." },
  scanReading:    { en: "Reading the file…",           ar: "جارٍ قراءة الملف…" },
  scanOcr:        { en: "Reading the text… {pct}%",    ar: "جارٍ قراءة النص… {pct}%" },
  scanNoRows:     { en: "Couldn't read any rows from this file. A flatter, better-lit photo — or the original Excel — will work far better.",
                    ar: "تعذّرت قراءة أي صف من هذا الملف. صورة مستوية بإضاءة أفضل — أو ملف Excel الأصلي — ستعطي نتيجة أفضل بكثير." },
  scanFail:       { en: "Couldn't read this file: {msg}", ar: "تعذّرت قراءة هذا الملف: {msg}" },
  scanFound:      { en: "Read {n} row{s}. Review and correct them before use.",
                    ar: "تمت قراءة {n} صف{s}. راجعها وصحّحها قبل الاستخدام." },
  scanUse:        { en: "✓ Use these rows",            ar: "✓ استخدم هذه الصفوف" },
  scanAddRow:     { en: "+ Add row",                   ar: "+ إضافة صف" },
  scanRedo:       { en: "↺ Choose another file",       ar: "↺ اختر ملفًا آخر" },
  scanTip:        { en: "A flat, well-lit photo of a printed list reads far better than a handwritten sketch.",
                    ar: "الصورة المستوية بإضاءة جيدة لقائمة مطبوعة تُقرأ أفضل بكثير من رسم يدوي." },
  scanUsedOk:     { en: "✓ {n} row{s} taken from the scan — check them before optimising.",
                    ar: "✓ تم أخذ {n} صف{s} من المسح — تحقق منها قبل التحسين." },

  // ---- optimize buttons ----
  optPlates:      { en: "⚡ OPTIMIZE PLATES",           ar: "⚡ حسّن الألواح" },
  optSections:    { en: "⚡ OPTIMIZE SECTIONS",         ar: "⚡ حسّن المقاطع" },
  optimizing:     { en: "⚙ Optimizing…",               ar: "⚙ جارٍ التحسين…" },

  // ---- results: shared ----
  optResults:     { en: "Optimization Results",        ar: "نتائج التحسين" },
  summaryBuy:     { en: "✦ SUMMARY — WHAT YOU NEED TO BUY", ar: "✦ الخلاصة — ما تحتاج شراءه" },
  wasteAfterCut:  { en: "UNUSED MATERIAL AFTER CUTTING", ar: "المواد غير المستخدمة بعد القص" },
  utilization:    { en: "Utilization",                 ar: "نسبة الاستغلال" },

  // ---- plate results words ----
  sheet:          { en: "sheet",                       ar: "لوح" },
  sheets:         { en: "sheets",                      ar: "ألواح" },
  sheetOf:        { en: "of",                          ar: "من" },
  totalSheets:    { en: "TOTAL",                       ar: "الإجمالي" },
  scrapWeight:    { en: "SCRAP / OFFCUT WEIGHT",       ar: "وزن السكراب / القصاصات" },
  reusableShort:  { en: "reusable",                    ar: "قابل لإعادة الاستخدام" },
  totalParts:     { en: "Total Parts",                 ar: "إجمالي القطع" },
  totalSheetsCard:{ en: "Total Sheets",                ar: "إجمالي الألواح" },
  thkGroups:      { en: "Thickness Groups",            ar: "مجموعات السماكة" },
  purchaseWt:     { en: "Total Weight (to purchase)",  ar: "الوزن الإجمالي (للشراء)" },
  netWt:          { en: "Net Weight (used in project)",ar: "الوزن الصافي (المستخدم بالمشروع)" },
  reusableOffcuts:{ en: "Reusable Offcuts",            ar: "قصاصات قابلة لإعادة الاستخدام" },

  // ---- procurement block ----
  finalOutput:    { en: "Final Decision-Ready Output", ar: "مخرجات جاهزة لاتخاذ القرار" },
  procSummary:    { en: "Material Procurement Summary",ar: "ملخص توريد المواد" },
  material:       { en: "MATERIAL",                    ar: "المعدن" },
  whatToBuy:      { en: "What you need to buy",        ar: "ما تحتاج شراءه" },
  youNeed:        { en: "You need",                    ar: "تحتاج" },

  // ---- sections results ----
  bar:            { en: "bar",                         ar: "عود" },
  bars:           { en: "bars",                        ar: "أعواد" },
  whatToOrder:    { en: "⬡ WHAT TO ORDER",             ar: "⬡ ما يجب طلبه" },
  barsToBuy:      { en: "Bars to Buy",                 ar: "أعواد للشراء" },
  totalLength:    { en: "Total Length",                ar: "الطول الإجمالي" },
  offcut:         { en: "Offcut",                      ar: "قصاصة" },
  waste:          { en: "Waste",                       ar: "الهدر" },
  weight:         { en: "Weight",                      ar: "الوزن" },
  offcutLength:   { en: "OFFCUT LENGTH",               ar: "طول القصاصة" },

  // ---- downloads ----
  dlReports:      { en: "💾 Download Reports & Files",  ar: "💾 تنزيل التقارير والملفات" },
  dlProcPDF:      { en: "📄 Procurement PDF",           ar: "📄 توريد PDF" },
  dlProcXLS:      { en: "📊 Procurement Excel",         ar: "📊 توريد Excel" },
  dlLeftPDF:      { en: "♻ Leftover PDF",               ar: "♻ بواقٍ PDF" },
  dlLeftXLS:      { en: "♻ Leftover Excel",             ar: "♻ بواقٍ Excel" },

  // ---- email gate ----
  gateTitle:      { en: "Get your cutting report",     ar: "احصل على تقرير القص" },
  gateBody:       { en: "Enter your email once to download your procurement & cutting files. We use it only to contact you about this tool. No spam, no sharing.",
                    ar: "أدخل بريدك مرة واحدة لتنزيل ملفات التوريد والقص. نستخدمه فقط للتواصل معك بشأن هذه الأداة. بلا إزعاج وبلا مشاركة." },
  gateBtn:        { en: "Download report →",           ar: "تنزيل التقرير →" },
  gateSaving:     { en: "Saving…",                     ar: "جارٍ الحفظ…" },
  gateEmailErr:   { en: "Please enter a valid email address.", ar: "يرجى إدخال بريد إلكتروني صحيح." },

  // ---- units / small words ----
  mm:             { en: "mm",                          ar: "مم" },
  t_ton:          { en: "t",                           ar: "طن" },
  kg:             { en: "kg",                          ar: "كجم" },
};
/* ── i18n CORE ───────────────────────────────────────────────────────────── */
/* ── ARABIC COMPLETENESS helpers ─────────────────────────────────────────────
   SO_LANG mirrors the interface language for code that is not a React
   component (unit formatters, report builders, error messages). In English
   every helper returns exactly the text it returned before.                 */
let SO_LANG = "en";
const RT = (en, ar, x) => (SO_LANG === "ar" ? ar : SO_LANG === "en" ? en : soRtI18n(en, x));   // ru/zh/es: international add-on
const RT_HTML = () => (SO_LANG === "ar" ? ' dir="rtl" lang="ar"' : SO_LANG === "en" ? "" : soHtmlLangAttr());
const RT_CSS = () => (SO_LANG === "ar" ? "body{font-family:Tahoma,'Segoe UI',Arial,sans-serif}th{text-align:right}" : "");
function rtWorkbook(wb) { if (SO_LANG === "ar" && wb) { wb.Workbook = wb.Workbook || {}; wb.Workbook.Views = [{ RTL: true }]; } return wb; }
function soGrade(gr) { return gr === "Steel (grade not specified)" ? RT(gr, "حديد (الرتبة غير محددة)") : gr; }
function soCutLabel(l, t) { return l === "(spliced run)" ? t("splicedRun") : l; }
const GRADE_GROUP_AR = { "European (EN 10025)": "أوروبية (EN 10025)", "American (ASTM)": "أمريكية (ASTM)", "Chinese (GB)": "صينية (GB)", "Japanese (JIS)": "يابانية (JIS)", "Other (CSA / DIN)": "أخرى (CSA / DIN)" };
Object.assign(LANG_DICT, {
  thKgMu:      { en: "kg/m",   ar: "كجم/م" },
  m2pmU:       { en: "m²/m",   ar: "م²/م" },
  gU:          { en: "g",      ar: "جم" },
  computedTag: { en: " (computed)", ar: " (محسوب)" },
  splicedRun:  { en: "(spliced run)", ar: "(طول موصول)" },
  scSaved:     { en: "SAVED", ar: "تم توفيرها" },
  scYours:     { en: "yours for the next job — that's money kept", ar: "لمشروعك القادم — مال محفوظ بدل الهدر" },
  scCaption:   { en: "1220 × 2440 mm · 6 mm kerf · 8 parts", ar: "1220 × 2440 مم · عرض القص 6 مم · 8 قطع" },
  scKept:      { en: "L-offcut kept ≈ 56 kg", ar: "قصاصة L محفوظة ≈ 56 كجم" },
  scNesting:   { en: "nesting… {p}% of sheet", ar: "جارٍ التوزيع… {p}% من اللوح" },
});

/* ╔══════════════════════════════════════════════════════════════════════════╗
   ║  INTERNATIONAL INTERFACE — Русский · 中文 · Español (isolated add-on)      ║
   ╚══════════════════════════════════════════════════════════════════════════╝
   English stays the reference and Arabic stays exactly as it was: every
   helper below gives back the old English / Arabic value untouched and only
   adds an answer when the interface is Russian, Chinese or Spanish.

     I18N_UI[lang][key]   interface texts — the LANG_DICT keys, plus SL.*
                          (Section Library), FAM.* (families), MISC.* and DIMSRC
     RT_I18N_SRC / _TX    report, PDF and Excel texts, found by their English;
                          report lines with numbers carry their own 3rd argument
     no translation       → the English text is shown (never a blank)

   Switches: I18N_EXTRA_ENABLED = false → only English and Arabic are offered
             AUTO_LANG_DETECT  = false → a first visit always opens in English
──────────────────────────────────────────────────────────────────────────── */
const I18N_EXTRA_ENABLED = true;   // ◄ Russian, Chinese, Spanish on/off
const AUTO_LANG_DETECT = true;     // ◄ first visit from a Russian/Chinese/Spanish browser opens in that language
const SO_LANGS = I18N_EXTRA_ENABLED ? ["en", "ar", "ru", "zh", "es"] : ["en", "ar"];
const SO_LANG_NAMES = { en: "English", ar: "العربية", ru: "Русский", zh: "中文", es: "Español" };
const SO_LANG_TAG = { ru: "ru", zh: "zh-CN", es: "es" };
const SO_DATE_LOC = { ru: "ru-RU", zh: "zh-CN", es: "es-ES" };

const I18N_UI = {"ru":{"appName":"Оптимизатор раскроя металла","langName":"Русский","otherLang":"English","heroBadge":"Меньше отходов металла","heroTitleA":"Раскрой металла","heroTitleB":"и раскладка деталей","heroSub":"Листы и профили • Минимум отходов • Масса сразу","heroCta":"НАЧАТЬ РАСЧЁТ →","heroScroll":"↓ прокрутите, чтобы начать","showTitle":"Посмотрите, как оптимизируется раскрой листа","chipImport":"Импорт Tekla · Excel · CSV","chipDb":"1100+ профилей — EU · US · JIS · ГОСТ · GB","chipOffcut":"Реестр Г-образных остатков","chipTon":"Масса и ведомость закупки сразу","chipReports":"Отчёты PDF и Excel","chooseWhat":"Выберите, что рассчитать","ctTop":"Файл читается неправильно? Пришлите его нам","ctTitle":"Файл импортирован неправильно?","ctBody":"Пришлите нам файл Excel — мы выполним расчёт и вернём карту раскроя. Бесплатно.","ctEmail":"Эл. почта","ctWhats":"WhatsApp","ctWhatsHint":"Быстрее всего — можно просто прислать фото ведомости","ctBy":"Сделано инженером-конструктором для тех, кто режет металл.","ctBlocked":"Сайт заблокирован в сети вашего офиса? Некоторые корпоративные файрволы блокируют новые домены. Откройте его через мобильный интернет или напишите мне — я пришлю результат.","preNoticeTtl":"Бесплатно на этапе раннего доступа","preNoticeBody":"В ближайшие месяцы Steel Optimizer перейдёт на платную подписку. До этого всё, чем вы пользуетесь сегодня, остаётся бесплатным, а первые пользователи сохранят стартовую цену.","preNoticeCta":"Сообщить о запуске","preNoticeSent":"✓ Принято — вы узнаете первыми.","preNoticeMail":"ваш@email.ru","proTitle":"Экспорт — функция Pro","proBody":"Карта раскроя, процент отходов и реестр деловых остатков остаются бесплатными на экране. Pro добавляет отчёты по раскрою в PDF и Excel для скачивания.","proMonthly":"Помесячно","proYearly":"На год","proPerMo":"/мес.","proPerYr":"/год","proTrial":"Первый месяц бесплатно","proSave":"2 месяца бесплатно","proHaveKey":"У меня уже есть ключ","proKeyPh":"Вставьте лицензионный ключ","proActivate":"Активировать","proChecking":"Проверка…","proPrivacy":"Ваши файлы не покидают браузер. Ничего не загружается на сервер.","proErrEmpty":"Введите лицензионный ключ.","proErrNet":"Не удалось связаться с сервером лицензий. Проверьте подключение.","proErrLimit":"Этот ключ уже активирован на максимальном числе устройств.","proErrExpired":"Подписка закончилась. Продлите её, чтобы вернуть экспорт.","proErrBad":"Ключ не распознан. Проверьте, не пропущены ли символы.","proOk":"✓ Pro активен — экспорт открыт.","proPay":"Оформить подписку","proLoading":"Открываем защищённую оплату…","proCards":"Visa · Mastercard · Amex · Apple Pay · Google Pay · PayPal","proSecure":"Безопасная оплата через Lemon Squeezy. Данные карты не проходят через этот сайт.","proRestore":"Оплатили на другом устройстве?","proThanks":"Оплата получена — экспорт открыт. Спасибо.","modSections":"Стальные профили","modSectionsDesc":"Балки, колонны, уголки, гнутые Z- и C-профили и составные сечения.","modPlates":"Стальные листы","modPlatesDesc":"Раскладка деталей на стандартных листах (2D). Сначала деловые остатки, затем минимум новых листов, с общей массой.","chooseArrow":"Выбрать →","howEnter":"Как вы хотите ввести ведомость раскроя?","manualEntry":"Ввод вручную","uploadTitle":"Загрузить спецификацию","scanTitle":"Распознать PDF или фото","scanDescP":"Считать размеры листовых деталей из PDF или с фото распечатанной ведомости.","scanDescS":"Считать профили и длины из PDF или с фото распечатанной ведомости.","manualDescP":"Введите детали в простую таблицу.","manualDescS":"Введите профиль (IPE, UB, HEA, SHS, 20Б1…), выберите из списка и укажите длины.","uploadDescP":"Перетащите спецификацию. Листовые детали считываются, профили пропускаются, «грязные» таблицы очищаются.","uploadDescS":"Перетащите спецификацию. Профили распознаются, листовые детали пропускаются.","changeMethod":"↺ Сменить способ","backModules":"↩ Назад","backInput":"← Назад к вводу","lblThickness":"Толщина (мм)","lblWidth":"Ширина (мм)","lblLength":"Длина (мм)","lblQty":"Кол-во","lblQtyHow":"Сколько? (кол-во)","lblGrade":"Марка стали","lblKerf":"Ширина реза (мм)","lblMargin":"Отступ от кромки (мм)","lblReuseMin":"Мин. деловой остаток (мм)","lblRotate":"Разрешить поворот на 90°","lblId":"Поз.","lblProfile":"Профиль","lblGradeShort":"Марка","platesWS":"Стальные листы — рабочая область","stockSettings":"📐 Листы на складе и параметры реза","stockSheets":"Стандартные листы — задайте размер листа для каждой толщины","addThickness":"+ Добавить толщину","stockHint":"Если детали длиннее 2440 мм, задайте здесь лист большего размера (например, 1500×3000 или 2000×6000) для этой толщины — иначе такие детали не разместить.","spliceStrat":"Стыковка (детали больше листа)","spliceWelds":"Меньше сварных швов (крупные части)","splicePack":"Лучшее использование металла (больше мелких частей)","reusePlatesTtl":"♻ Использовать остатки листов (необязательно)","reuseBarsTtl":"♻ Использовать остатки профилей (необязательно)","reusePlatesLead":"Есть остатки листов с прошлого заказа или с другого объекта? Добавьте их здесь (или загрузите прежний файл остатков) — мы раскроим их в первую очередь, и вы закупите только необходимый минимум новых листов.","reuseBarsLead":"Есть остатки профилей с прошлого заказа или с другого объекта? Добавьте их здесь (или загрузите прежний файл остатков) — мы раскроим их в первую очередь, и вы закупите только необходимый минимум новых хлыстов.","noPlatesYet":"Остатки листов пока не добавлены.","noBarsYet":"Остатки профилей пока не добавлены.","addLeftPlate":"+ Добавить остаток листа","addLeftBar":"+ Добавить остаток профиля","reuseImportBtn":"⤓ Загрузить прежний файл остатков (PDF / Excel)","reuseReading":"Чтение","reuseBusy":"⏳ Чтение…","manualPartEntry":"✏️ Ввод деталей вручную","addPart":"+ Добавить деталь","uploadTeklaTtl":"📊 Excel","dropTekla":"Перетащите ведомость Tekla / Excel / CSV / TXT","orBrowse":"или нажмите, чтобы выбрать файл","scanCardTtl":"🔍 PDF / фото","scanDrop":"Перетащите PDF или фото вашей ведомости","scanAccept":"Форматы .pdf · .png · .jpg · .webp — или нажмите, чтобы выбрать файл","scanWarn":"⚠ Распознавание PDF или фото — это оценка, а не измерение. Проверьте каждую строку ниже до расчёта: неверное число здесь — неверно отрезанный металл.","scanReading":"Чтение файла…","scanOcr":"Распознавание текста… {pct}%","scanNoRows":"Не удалось считать ни одной строки. Ровное, хорошо освещённое фото — или исходный Excel — дадут намного лучший результат.","scanFail":"Не удалось прочитать файл: {msg}","scanFound":"Считано строк: {n}. Проверьте и исправьте их перед использованием.","scanUse":"✓ Использовать эти строки","scanAddRow":"+ Добавить строку","scanRedo":"↺ Выбрать другой файл","scanTip":"Ровное, хорошо освещённое фото печатной ведомости читается гораздо лучше, чем рукописный эскиз.","scanUsedOk":"✓ Строк взято из распознавания: {n} — проверьте их перед расчётом.","optPlates":"⚡ РАССЧИТАТЬ РАСКРОЙ ЛИСТОВ","optSections":"⚡ РАССЧИТАТЬ РАСКРОЙ ПРОФИЛЕЙ","optimizing":"⚙ Расчёт…","optResults":"Результаты расчёта","summaryBuy":"✦ ИТОГ — ЧТО НУЖНО ЗАКУПИТЬ","wasteAfterCut":"НЕИСПОЛЬЗОВАННЫЙ МАТЕРИАЛ ПОСЛЕ РАСКРОЯ","utilization":"Использование","sheet":"лист","sheets":"листа(ов)","sheetOf":"толщиной","totalSheets":"ИТОГО","scrapWeight":"МАССА ОТХОДОВ / ОСТАТКОВ","reusableShort":"деловых","totalParts":"Всего деталей","totalSheetsCard":"Всего листов","thkGroups":"Групп по толщине","purchaseWt":"Общая масса (к закупке)","netWt":"Масса нетто (в конструкции)","reusableOffcuts":"Деловые остатки","finalOutput":"Итог для принятия решения","procSummary":"Сводка закупки металла","material":"МАТЕРИАЛ","whatToBuy":"Что нужно закупить","youNeed":"Нужно","bar":"хлыст","bars":"хлыста(ов)","whatToOrder":"⬡ ЧТО ЗАКАЗАТЬ","barsToBuy":"Хлыстов к закупке","totalLength":"Общая длина","offcut":"Остаток","waste":"Отходы","weight":"Масса","offcutLength":"ДЛИНА ОСТАТКОВ","dlReports":"💾 Скачать отчёты и файлы","dlProcPDF":"📄 Закупка — PDF","dlProcXLS":"📊 Закупка — Excel","dlLeftPDF":"♻ Остатки — PDF","dlLeftXLS":"♻ Остатки — Excel","gateTitle":"Получите отчёт по раскрою","gateBody":"Один раз укажите e-mail, чтобы скачать файлы закупки и раскроя. Мы используем его только для связи по этому инструменту. Без спама и передачи третьим лицам.","gateBtn":"Скачать отчёт →","gateSaving":"Сохранение…","gateEmailErr":"Введите корректный адрес e-mail.","mm":"мм","t_ton":"т","kg":"кг","thKgMu":"кг/м","m2pmU":"м²/м","gU":"г","computedTag":" (расчётная)","splicedRun":"(со стыком)","scSaved":"СОХРАНЕНО","scYours":"на следующий заказ — это сбережённые деньги","scCaption":"1220 × 2440 мм · рез 6 мм · 8 деталей","scKept":"Г-образный остаток ≈ 56 кг","scNesting":"раскладка… {p}% листа","nameThisPlate":"Название листа","removeX":"✕ удалить","whatShape":"Какой формы остаток?","shapeRect":"▭ Прямоугольник","shapeL":"⌐ Г-образный (с вырезанным углом)","willNestInto":"✓ Детали будут размещены в области","biggestRectL":" (наибольший прямоугольник внутри Г-образного остатка)","pieceWord":"шт.","piecesWord":"шт.","addLeftMore":"+ Добавить остаток с другого заказа / объекта","notInStock":"(нет на складе)","upBannerP":"💡 Толщина берётся из файла, размер листа — из вашего списка выше. Если толщина неизвестна, используется 1220×2440.","upHintP":"Читает столбец профиля (PLT/FLT/PL → толщина × ширина). Прокатные профили пропускаются.","mapColumns":"Сопоставьте столбцы:","applyMapping":"Применить","partsReady":"ДЕТАЛИ ГОТОВЫ","moreWord":"ещё","msgImportedP":"✓ Импортировано типов листовых деталей: {n}. Прокатные профили и строки итогов пропущены.","msgMapCols":"⚠ Не удалось определить столбцы автоматически. Сопоставьте их ниже.","msgReadFail":"⚠ Не удалось прочитать этот файл.","msgImportedP2":"✓ Импортировано листовых деталей: {n}","msgExtractFail":"⚠ Всё ещё не удалось извлечь данные.","liReading":"Чтение {name}…","liWrongType":"⚠ Похоже, это файл остатков для раздела «{type}». Используйте его в разделе «{mod}».","liImported":"✓ Импортировано типов остатков: {n} ({m} шт.). Они будут раскроены в первую очередь.","liErr":"⚠ {msg}","modPlatesName":"Листы","modSectionsName":"Профили","sheetUnitOf":"{n} × {sw} × {sh} мм ({mat})","acrossThk":"ИТОГО: листов {n}, толщин {g}","spliceBannerHd":"⚙ Деталей больше листа: {n} — вырезаются по частям и свариваются","spliceBannerBd":"Каждая такая деталь раскладывается на несколько листов и соединяется сварным швом (лист большего размера не нужен). До изготовления подтвердите, что сварной стык для этих деталей допустим:","reusedBannerP":"♻ В первую очередь использованы деловые остатки прошлых проектов: {n}. Ниже указано минимальное число новых стандартных листов, которые ещё нужно закупить.","nestLayoutTtl":"🗂 Карта раскроя — лист {thk} мм (листов: {n})","offcutRegTtl":"♻ Реестр деловых остатков","offcutRegLead":"Деловых остатков: {n} ≈ {t} т — сохраните для будущих заказов (≥ {min} мм по стороне). Для Г-образных остатков «ПОЛЕЗНО» — наибольший прямоугольник, который из них можно вырезать.","thQty":"КОЛ-ВО","thThickness":"ТОЛЩИНА","thFromSheet":"С ЛИСТА","thShape":"ФОРМА","thOverall":"ГАБАРИТ","thUsableRect":"ПОЛЕЗНЫЙ ПРЯМОУГ.","thWeight":"МАССА","thStatus":"СТАТУС","shapeLshort":"⌐ Г-образный","shapeRectShort":"▭ Прямоуг.","keepStatus":"✓ сохранить","fromSheetN":"Лист {n}","dlTwoDeliv":"Два комплекта: файлы закупки (что купить + карта раскроя) и файлы остатков (деловые остатки для хранения).","miniSheet":"Лист","miniThickness":"Толщина","miniParts":"Детали","miniUtil":"Использование","reusableLeftover":"♻ Деловой остаток","pbFinalReady":"Итог для принятия решения","pbProcSummary":"Сводка закупки металла","thSheetSize":"РАЗМЕР ЛИСТА","thSheetsReq":"ЛИСТОВ","thPurchaseWt":"МАССА ЗАКУПКИ (целые листы)","thParts":"ДЕТАЛИ","thNetWt":"МАССА ДЕТАЛЕЙ","thScrapWt":"МАССА ОТХОДОВ","pbLegend":"● Общая масса = закупаемые целые листы · ● Масса нетто = металл, вошедший в проект · ● Отходы = остатки (общая − нетто)","pbTotalPurchase":"Общая масса закупки","pbGrossNote":"брутто — все целые листы ({n} шт.)","pbNetNote":"металл, который войдёт в конструкцию","pbScrapNote":"{p}% от закупки","pbGenerated":"Создано в Steel Optimizer — раскрой и раскладка","thicknessesWord":"толщ.","sectionsWS":"Стальные профили — рабочая область","manualCutList":"✏️ Ведомость раскроя вручную","thLengths":"Длины (мм)","thMarketLen":"Длина проката в продаже (мм)","addProfile":"+ ДОБАВИТЬ ПРОФИЛЬ","piecesCount":"{n} шт.","lengthsHint":"например 3000 4500 6000x3  →  одна 3 м, одна 4,5 м, три по 6 м","blank12m":"пусто = 12 м","secManualHelp":"Вводите IPE, HEA, HEB, UB, UC, JIS HW/HM/HN, RHS, SHS, CHS, гнутые C/Z, ГОСТ (20Б1, 20П, уголок 75х6) и GB. Сварная балка: 900x400x20x15. Длины разделяйте пробелом, например 3000 4500 6000x3 — запись 6000x3 означает три детали по 6 м.","computedSuffix":"(расчётная)","upTitleS":"Загрузите спецификацию","upBodyS":"Перетащите ведомость Tekla — или любую спецификацию с длинами и количеством профилей.","upAcceptS":"Форматы .xlsx · .xls · .csv · .txt — или нажмите, чтобы выбрать файл","upHintS":"Профили распознаются автоматически · листовые строки пропускаются (для них — раздел «Листы») · длина хлыста по умолчанию 12 м, её можно изменить.","secUpErr":"Профили не обнаружены. Если в файле только листы, используйте раздел «Листы».","secReadErr":"Не удалось прочитать: {msg}","secNoValid":"Нет корректных деталей.","secEnterLen":"Введите хотя бы одну длину.","uploadedOk":"✓ {name} — профилей: {rows}, деталей: {pieces}","platesIgnored":" · листовых строк пропущено: {n}","thKgM":"кг/м","andMore":"…и ещё {n}","loLeftoverLen":"Длина остатка (мм)","loBarLeft":"длина оставшегося куска","loHowManyLike":"сколько таких","loSameProfile":"Остаток идёт только на детали того же профиля и марки — оставьте марку пустой, чтобы сочетать с деталями без марки. Сначала из каждого остатка вырезаются наиболее подходящие детали, затем закупается минимум новых хлыстов на оставшееся.","whatToOrderTtl":"⬡ ЧТО ЗАКАЗАТЬ","youNeedWord":"НУЖНО","gradeWord":"марка","totalAcrossProf":"ИТОГО: хлыстов {n}, профилей {g}","leftoversUsedHd":"♻ СНАЧАЛА ИСПОЛЬЗОВАНЫ ОСТАТКИ","leftoversUsedBd":"Ваших остатков раскроено до закупки новых: {n}","steelNotBought":"МЕТАЛЛ БЕЗ ЗАКУПКИ","longMembersHd":"⚙ ДЛИННЫЕ ЭЛЕМЕНТЫ — СО СТЫКОМ ИЗ ХЛЫСТОВ {len}","eachFromBars":" — каждый из {n} хлыстов по {len}","cutFromLeftTtl":"♻ Вырезано из ваших остатков — без закупки","cutFromLeftLead":"Эти детали получены из остатков, которые у вас уже были, — в ведомость закупки выше они не входят.","leftRemaining":" После этого остаётся: {kg}.","visualCutPlan":"Наглядная карта раскроя","noPlan":"Нет раскроя.","allFromLeft":"Все детали вырезаны из ваших остатков — закупать ничего не нужно.","reusableBarTtl":"♻ Деловые остатки профилей — металл для следующего проекта","reusableBarLead":"Концевых остатков ≥ 1 м: {n} ≈ {kg} — сохраните их для следующего заказа вместо закупки нового металла.","thFromBar":"ИЗ ХЛЫСТА","thStockLength":"ДЛИНА ХЛЫСТА","thReusableLeft":"ДЕЛОВОЙ ОСТАТОК","barNum":"Хлыст {n}","leftoverNum":"Остаток {n}","stBarsBuy":"Хлыстов к закупке","stTotalLen":"Общая длина","stOffcut":"Остаток","stWaste":"Отходы","stUtil":"Использование","stWeight":"Масса","cbBar":"ХЛЫСТ №{n}","cbAvailable":"длина {len}","cbUsed":"использовано {p}%","canvLoffcut":"♻ Г-ОСТАТОК","canvOffcut":"♻ ОСТАТОК","canvUse":"полезно {w}×{h}","weldLbl":"ШОВ {w}×{h}","sheetTab":"Лист {n} · {p}%","reuseListLead":"♻ Деловые остатки: {list}","lShapeUse":"Г-образный, полезно ","notchWord":"вырез","liveNesting":"● РАСКЛАДКА В РЕАЛЬНОМ ВРЕМЕНИ — РЕАЛЬНЫЙ ПРИМЕР","scrapReuse":"отходы/остатки","usableWord":"ПОЛЕЗНО","widthArrow":"↔ ширина","lengthArrow":"↕ длина","cutoutWord":"вырез","lThickNote":"толщина {t} мм · зелёным — прямоугольник, в который размещаются детали","lpIntro":"Встаньте у одного угла остатка и примите его за 0, 0. Пройдите по контуру и вводите каждый угол — чертёж строится по ходу.","lpColX":"X · ширина →","lpColY":"Y · длина ↑","lpAxisX":"X · ширина","lpAxisY":"Y · длина","lpReset":"↺ Сбросить форму","lpOverall":"Габарит","lpUsable":"полезно","lpTip":"подсказка: у Г-образного остатка 6 углов, образующих ровную ступень","lpHint0":"начальный угол — оставьте 0, 0","lpHint1":"→ вдоль нижней кромки","lpHint2":"↑ вверх по стороне","lpHint3":"← внутрь выреза","lpHint4":"↑ вверх до верха","lpHint5":"← обратно к началу","priceCardTtl":"💰 Цена металла и стоимость (необязательно)","priceCardLead":"Укажите цену металла за тонну — стоимость появится в результатах и в отчётах PDF/Excel. Оставьте поле пустым, чтобы не считать стоимость.","priceCurrency":"Валюта","pricePerTonLbl":"Цена за тонну","costPurchase":"ОБЩАЯ СТОИМОСТЬ ЗАКУПКИ","costNet":"СТОИМОСТЬ МЕТАЛЛА В ДЕЛЕ","costScrap":"СТОИМОСТЬ ОТХОДОВ","costOffcut":"СТОИМОСТЬ ДЕЛОВЫХ ОСТАТКОВ","costBasedOn":"Расчёт стоимости: {p} за тонну (ваша цена).","paintHd":"🎨 ПЛОЩАДЬ ОКРАСКИ","paintNoteS":"Длины деталей нетто из вашей ведомости × наружная поверхность на метр. Углы прямые, радиусы сопряжения не учитываются → с запасом (обычно на 2–5% выше сортамента). Остатки и отходы не включены.","paintNoteP":"Только одна сторона · ваши детали (длина × ширина × кол-во) · остатки и отходы не включены.","paintMissing":"Не включено — в сортаменте нет размеров: {list}","stPaint":"Площадь окраски","paintOneFace":"Площадь окраски (1 сторона)","m2":"м²","siteTables":"Сортамент стальных профилей — масса и площадь окраски на метр","SL.btn":"Сортамент","SL.title":"Сортамент стальных профилей","SL.pickFam":"Выберите группу профилей","SL.search":"Поиск профиля…","SL.back":"Назад","SL.close":"Закрыть","SL.sections":"профилей","SL.designation":"Обозначение","SL.family":"Группа","SL.unitMass":"Масса 1 м","SL.area":"Площадь сечения","SL.perBar":"Масса хлыста","SL.perTonne":"Хлыстов на тонну","SL.barsOf":"хлыст. по","SL.note":"Площадь рассчитана по массе при плотности 7850 кг/м³. Для карты раскроя и массы используйте расчёт.","SL.useIt":"Рассчитать с этим профилем ↓","SL.paint":"Площадь окраски (с запасом)","SL.page":"Полная страница с данными ↗","FAM.IPE":"IPE — европейские двутавры","FAM.HEA":"HEA — широкополочные (лёгкие)","FAM.HEB":"HEB — широкополочные (средние)","FAM.HEM":"HEM — широкополочные (тяжёлые)","FAM.IPN":"IPN — двутавры с уклоном полок","FAM.UB":"UB — британские балки","FAM.UC":"UC — британские колонны","FAM.UBP":"UBP — британские сваи","FAM.PFC":"PFC — британские швеллеры","FAM.UPN":"UPN — швеллеры с уклоном полок","FAM.UPE":"UPE — швеллеры с параллельными полками","FAM.W":"W — американские широкополочные","FAM.C-AMER":"C — американские швеллеры","FAM.MC":"MC — американские швеллеры (прочие)","FAM.L":"L — уголки","FAM.SHS":"SHS — трубы квадратные","FAM.RHS":"RHS — трубы прямоугольные","FAM.CHS":"CHS — трубы круглые","FAM.C-COLD":"C — гнутые прогоны","FAM.Z-COLD":"Z — гнутые прогоны","FAM.JIS-HW":"HW — широкополочные H (GB/T 11263 · KS D 3502)","FAM.JIS-HM":"HM — среднеполочные H (GB/T 11263 · KS D 3502)","FAM.JIS-HN":"HN — узкополочные H (GB/T 11263 · KS D 3502)","FAM.JIS-I":"I — двутавры с уклоном полок (JIS G 3192)","FAM.JIS-C":"C — швеллеры с уклоном полок (JIS G 3192)","FAM.JIS-LIP":"LC — гнутые швеллеры с отгибами (JIS G 3350)","FAM.GOST-I":"Двутавры ГОСТ","FAM.GOST-C":"Швеллеры ГОСТ","FAM.GOST-B":"ГОСТ Б — нормальные двутавры","FAM.GOST-K":"ГОСТ К — колонные двутавры","FAM.GB-I":"Двутавры GB (Китай)","FAM.GB-C":"Швеллеры GB (Китай)","FAM.HSS":"HSS — американские трубы","FAM.HP":"HP — американские сваи","FAM.S-AMER":"S — американские стандартные двутавры","FAM.M-AMER":"M — американские двутавры (прочие)","DIMSRC":"Размеры: AISC Shapes DB v16.0 · EN 10365 · British Steel (BS EN 10365:2017) · JIS G 3192 / KS / GB. Каждое значение сверено с массой 1 м; формула проверена по 379 опубликованным площадям сечений (макс. погрешность 0,75%).","MISC.langMenu":"Язык","MISC.dimDepth":"Высота h","MISC.dimFlangeW":"Ширина полки b","MISC.dimWebT":"Толщина стенки tw","MISC.dimFlangeT":"Толщина полки tf","MISC.dimLegA":"Полка a","MISC.dimLegB":"Полка b","MISC.dimThick":"Толщина t","MISC.dimHeight":"Высота h","MISC.dimWidth":"Ширина b","MISC.dimWall":"Толщина стенки t","MISC.dimOD":"Наружный диаметр D","MISC.dimStd":"Стандарт","MISC.asianNote":"Примечание: h, b, tw, tf совпадают в JIS G 3192, KS D 3502 и GB/T 11263 — отличается только радиус сопряжения, что даёт разницу в массе до 2%. Здесь масса по KS/GB.","MISC.intlSrc":"Источник: база сечений SAP2000 ({file}) · {std}. Масса = площадь сечения × 7850 кг/м³. Строки, где площадь не согласуется с размерами, исключены.","MISC.gEU":"Европейские (EN 10025)","MISC.gUS":"Американские (ASTM)","MISC.gCN":"Китайские (GB)","MISC.gJP":"Японские (JIS)","MISC.gOT":"Другие (CSA / DIN)","MISC.tbSave":"Сохранить в PDF","MISC.tbPrint":"Печать","MISC.tbHint":"Обе кнопки открывают окно печати браузера. Для PDF выберите <b>Принтер</b> → «Сохранить как PDF».","MISC.errTitle":"Произошла непредвиденная ошибка — перезагрузите страницу.","MISC.errReload":"Перезагрузить","MISC.reportFail":"Не удалось сформировать отчёт. Попробуйте ещё раз.","unknownRows":"⚠ Не распознано как профиль — не рассчитано: {list}. Проверьте эти наименования или пришлите нам файл.","fromFile":"как в вашем файле: {name}","fromFileTag":" (масса из вашего файла)","fileKgmNote":"в файле: {v} кг/м ({d}%)","fileKgmWarn":"⚠ в файле: {v} кг/м ({d}%) — проверьте этот профиль","fileKgmCell":"в файле {v}","MISC.barOf":""},"zh":{"appName":"钢材下料与排版优化","langName":"中文","otherLang":"English","heroBadge":"减少钢材浪费","heroTitleA":"钢材下料","heroTitleB":"与排版优化","heroSub":"钢板与型钢 • 损耗最低 • 即时吨数","heroCta":"开始优化 →","heroScroll":"↓ 向下滚动开始","showTitle":"看看钢板下料如何优化","chipImport":"导入 Tekla · Excel · CSV","chipDb":"1100+ 种截面 — EU · US · JIS · GOST · GB","chipOffcut":"L 形余料清单","chipTon":"即时吨数与采购清单","chipReports":"PDF 与 Excel 报告","chooseWhat":"选择要优化的内容","ctTop":"文件读取不正确？发邮件给我们","ctTitle":"文件导入不正确？","ctBody":"把 Excel 文件发给我们，我们来计算并把下料方案发回给您。免费。","ctEmail":"邮箱","ctWhats":"WhatsApp","ctWhatsHint":"最快 — 也可以直接发清单照片","ctBy":"由结构工程师打造，服务每一位下料的人。","ctBlocked":"公司网络打不开本站？部分企业防火墙会拦截新域名。请用手机流量打开，或给我留言，我把结果发给您。","preNoticeTtl":"抢先体验期间免费","preNoticeBody":"Steel Optimizer 将在未来几个月转为付费订阅。在此之前，您现在使用的全部功能继续免费，早期用户可保留创始价格。","preNoticeCta":"上线时通知我","preNoticeSent":"✓ 已记录 — 我们会第一时间通知您。","preNoticeMail":"您的邮箱","proTitle":"导出为 Pro 功能","proBody":"排版图、损耗率和余料清单在屏幕上始终免费。Pro 增加可下载的 PDF 和 Excel 下料报告。","proMonthly":"按月","proYearly":"按年","proPerMo":"/月","proPerYr":"/年","proTrial":"首月免费","proSave":"赠送 2 个月","proHaveKey":"我已有激活码","proKeyPh":"粘贴您的许可证密钥","proActivate":"激活","proChecking":"正在验证…","proPrivacy":"您的文件不会离开浏览器，不会上传到任何服务器。","proErrEmpty":"请输入许可证密钥。","proErrNet":"无法连接许可证服务器，请检查网络。","proErrLimit":"该密钥已在允许的最多设备上激活。","proErrExpired":"订阅已到期。续订后即可恢复导出。","proErrBad":"无法识别该密钥，请检查是否漏了字符。","proOk":"✓ Pro 已激活 — 导出已解锁。","proPay":"订阅","proLoading":"正在打开安全支付页面…","proCards":"Visa · Mastercard · Amex · Apple Pay · Google Pay · PayPal","proSecure":"由 Lemon Squeezy 提供安全支付，银行卡信息不经过本站。","proRestore":"在其他设备上付款了？","proThanks":"付款成功 — 导出已解锁。谢谢！","modSections":"型钢","modSectionsDesc":"梁、柱、角钢、冷弯 Z 型与 C 型钢以及焊接组合截面。","modPlates":"钢板与薄板","modPlatesDesc":"在标准钢板上排版零件（二维）。先用余料，再用最少的新钢板，并给出总吨数。","chooseArrow":"选择 →","howEnter":"您想如何输入下料清单？","manualEntry":"手动输入","uploadTitle":"上传材料表","scanTitle":"识别 PDF 或照片","scanDescP":"从 PDF 清单或打印清单的照片中读取钢板尺寸。","scanDescS":"从 PDF 清单或打印清单的照片中读取截面和长度。","manualDescP":"在简单的表格中输入零件。","manualDescS":"输入截面（IPE、UB、HEA、SHS、HW200×200…），从实时列表中选择，再输入长度。","uploadDescP":"拖入材料表：读取钢板零件，忽略型钢，自动清理杂乱表格。","uploadDescS":"拖入材料表：自动识别型钢，忽略钢板。","changeMethod":"↺ 更换方式","backModules":"↩ 返回","backInput":"← 返回输入","lblThickness":"厚度 (mm)","lblWidth":"宽度 (mm)","lblLength":"长度 (mm)","lblQty":"数量","lblQtyHow":"数量？","lblGrade":"钢材牌号","lblKerf":"割缝宽度 (mm)","lblMargin":"边距 (mm)","lblReuseMin":"可再利用余料最小尺寸 (mm)","lblRotate":"允许旋转 90°","lblId":"编号","lblProfile":"截面","lblGradeShort":"牌号","platesWS":"钢板 — 工作区","stockSettings":"📐 原材料与切割设置","stockSheets":"标准钢板 — 为每种厚度设置板幅","addThickness":"+ 添加厚度","stockHint":"如果零件长于 2440 mm，请在此为该厚度设置更大的板幅（如 1500×3000 或 2000×6000），否则这些零件无法排版。","spliceStrat":"拼接方式（零件大于板幅时）","spliceWelds":"焊缝最少（分块较大）","splicePack":"材料利用率最高（分块更多、更小）","reusePlatesTtl":"♻ 使用余料钢板（可选）","reuseBarsTtl":"♻ 使用余料型钢（可选）","reusePlatesLead":"有以前项目或其他工地剩下的钢板？在这里添加（或导入以前的余料文件），我们会先用余料下料，您只需采购最少的新钢板。","reuseBarsLead":"有以前项目或其他工地剩下的型钢？在这里添加（或导入以前的余料文件），我们会先用余料下料，您只需采购最少的新原材。","noPlatesYet":"尚未添加余料钢板。","noBarsYet":"尚未添加余料型钢。","addLeftPlate":"+ 添加余料钢板","addLeftBar":"+ 添加余料型钢","reuseImportBtn":"⤓ 导入以前的余料文件（PDF / Excel）","reuseReading":"正在读取","reuseBusy":"⏳ 正在读取…","manualPartEntry":"✏️ 手动输入零件","addPart":"+ 添加零件","uploadTeklaTtl":"📊 Excel","dropTekla":"拖入 Tekla 清单 / Excel / CSV / TXT","orBrowse":"或点击选择文件","scanCardTtl":"🔍 PDF / 照片","scanDrop":"拖入清单的 PDF 或照片","scanAccept":"支持 .pdf · .png · .jpg · .webp — 或点击选择文件","scanWarn":"⚠ 从 PDF 或照片读取只是尽力识别，并非测量。优化前请逐行核对 — 这里错一个数字，钢材就会切错。","scanReading":"正在读取文件…","scanOcr":"正在识别文字… {pct}%","scanNoRows":"未能从该文件读取到任何行。更平整、光线更好的照片 — 或原始 Excel — 效果会好得多。","scanFail":"无法读取该文件：{msg}","scanFound":"已读取 {n} 行。使用前请检查并更正。","scanUse":"✓ 使用这些行","scanAddRow":"+ 添加行","scanRedo":"↺ 选择其他文件","scanTip":"平整、光线充足的打印清单照片，比手绘草图识别效果好得多。","scanUsedOk":"✓ 已从识别结果取用 {n} 行 — 优化前请核对。","optPlates":"⚡ 优化钢板排版","optSections":"⚡ 优化型钢下料","optimizing":"⚙ 正在优化…","optResults":"优化结果","summaryBuy":"✦ 汇总 — 需要采购的材料","wasteAfterCut":"切割后未利用材料","utilization":"利用率","sheet":"张","sheets":"张","sheetOf":"板厚","totalSheets":"合计","scrapWeight":"废料 / 余料重量","reusableShort":"块可再利用","totalParts":"零件总数","totalSheetsCard":"钢板总数","thkGroups":"厚度组数","purchaseWt":"总重量（采购）","netWt":"净重（用于工程）","reusableOffcuts":"可再利用余料","finalOutput":"可直接决策的最终结果","procSummary":"材料采购汇总","material":"材质","whatToBuy":"需要采购的材料","youNeed":"需要","bar":"根","bars":"根","whatToOrder":"⬡ 订货清单","barsToBuy":"需采购根数","totalLength":"总长度","offcut":"余料","waste":"损耗","weight":"重量","offcutLength":"余料长度","dlReports":"💾 下载报告与文件","dlProcPDF":"📄 采购 PDF","dlProcXLS":"📊 采购 Excel","dlLeftPDF":"♻ 余料 PDF","dlLeftXLS":"♻ 余料 Excel","gateTitle":"获取下料报告","gateBody":"输入一次邮箱即可下载采购与下料文件。我们仅用它就本工具与您联系，不发垃圾邮件，不向他人提供。","gateBtn":"下载报告 →","gateSaving":"正在保存…","gateEmailErr":"请输入有效的邮箱地址。","mm":"mm","t_ton":"t","kg":"kg","thKgMu":"kg/m","m2pmU":"m²/m","gU":"g","computedTag":"（计算值）","splicedRun":"（拼接段）","scSaved":"已保留","scYours":"留给下一个项目 — 这就是省下的钱","scCaption":"1220 × 2440 mm · 割缝 6 mm · 8 个零件","scKept":"保留 L 形余料 ≈ 56 kg","scNesting":"排版中… 占板 {p}%","nameThisPlate":"为这块钢板命名","removeX":"✕ 删除","whatShape":"余料是什么形状？","shapeRect":"▭ 矩形","shapeL":"⌐ L 形（缺一个角）","willNestInto":"✓ 零件将排入可用区域","biggestRectL":"（L 形内可容纳的最大矩形）","pieceWord":"块","piecesWord":"块","addLeftMore":"+ 添加其他项目 / 工地的余料","notInStock":"（无库存）","upBannerP":"💡 厚度取自文件，对应的板幅取自上方列表。厚度未知时按 1220×2440。","upHintP":"读取截面列（PLT/FLT/PL → 厚度 × 宽度），忽略热轧型钢。","mapColumns":"对应列：","applyMapping":"应用对应","partsReady":"零件已就绪","moreWord":"更多","msgImportedP":"✓ 已导入 {n} 种钢板零件。已忽略热轧型钢和小计行。","msgMapCols":"⚠ 无法自动识别列，请在下方指定。","msgReadFail":"⚠ 无法读取该文件。","msgImportedP2":"✓ 已导入 {n} 个钢板零件","msgExtractFail":"⚠ 仍无法提取数据。","liReading":"正在读取 {name}…","liWrongType":"⚠ 这似乎是「{type}」余料文件，请在「{mod}」模块中使用。","liImported":"✓ 已导入 {n} 种余料（共 {m} 件），将优先用于下料。","liErr":"⚠ {msg}","modPlatesName":"钢板","modSectionsName":"型钢","sheetUnitOf":"{n} × {sw} × {sh} mm（{mat}）","acrossThk":"合计：{n} 张钢板，{g} 种厚度","spliceBannerHd":"⚙ {n} 个零件大于板幅 — 分块切割后焊接","spliceBannerBd":"每个此类零件分布在多张钢板上，用焊缝连接（无需更大的钢板）。加工前请确认这些零件允许设置焊缝：","reusedBannerP":"♻ 已优先使用以往项目的 {n} 块余料。下方张数为仍需采购的最少新标准钢板。","nestLayoutTtl":"🗂 排版图 — {thk} mm 钢板（{n} 张）","offcutRegTtl":"♻ 可再利用余料清单","offcutRegLead":"{n} 块可再利用余料 ≈ {t} t，留待后续项目（每边 ≥ {min} mm）。L 形余料的「可用」为其中可切出的最大矩形。","thQty":"数量","thThickness":"厚度","thFromSheet":"来源钢板","thShape":"形状","thOverall":"外形尺寸","thUsableRect":"可用矩形","thWeight":"重量","thStatus":"状态","shapeLshort":"⌐ L 形","shapeRectShort":"▭ 矩形","keepStatus":"✓ 保留","fromSheetN":"第 {n} 张","dlTwoDeliv":"两类文件：采购文件（采购清单 + 下料方案）和余料文件（需保留的可再利用余料）。","miniSheet":"钢板","miniThickness":"厚度","miniParts":"零件","miniUtil":"利用率","reusableLeftover":"♻ 可再利用余料","pbFinalReady":"可直接决策的最终结果","pbProcSummary":"材料采购汇总","thSheetSize":"板幅","thSheetsReq":"所需张数","thPurchaseWt":"采购重量（整张）","thParts":"零件","thNetWt":"零件净重","thScrapWt":"废料重量","pbLegend":"● 总重量 = 采购的整张钢板 · ● 净重 = 用于工程的钢材 · ● 废料 = 余料（总重 − 净重）","pbTotalPurchase":"采购总重量","pbGrossNote":"毛重 — 全部 {n} 张整板","pbNetNote":"最终用于结构的钢材","pbScrapNote":"占采购的 {p}%","pbGenerated":"由钢材下料与排版优化工具生成","thicknessesWord":"种厚度","sectionsWS":"型钢 — 工作区","manualCutList":"✏️ 手动下料清单","thLengths":"长度 (mm)","thMarketLen":"市场供货长度 (mm)","addProfile":"+ 添加截面","piecesCount":"{n} 件","lengthsHint":"例如 3000 4500 6000x3  →  一根 3 m、一根 4.5 m、三根 6 m","blank12m":"留空 = 12 m","secManualHelp":"可输入 IPE、HEA、HEB、UB、UC、JIS HW/HM/HN、RHS、SHS、CHS、冷弯 C/Z、GOST 和 GB（如 HW200×200×8×12、I20a、[20a、∠75×6）。焊接组合梁：输入 900x400x20x15。长度之间用空格分隔，例如 3000 4500 6000x3 — 6000x3 表示三根 6 m。","computedSuffix":"（计算值）","upTitleS":"上传材料表","upBodyS":"拖入 Tekla 材料表 — 或任何带有钢材长度和数量的材料表。","upAcceptS":"支持 .xlsx · .xls · .csv · .txt — 或点击选择文件","upHintS":"自动识别型钢 · 忽略钢板行（请用钢板模块处理）· 原材长度默认 12 m，可修改。","secUpErr":"未识别到型钢。如果文件中只有钢板，请使用钢板模块。","secReadErr":"读取失败：{msg}","secNoValid":"没有有效的下料项。","secEnterLen":"请至少输入一个长度。","uploadedOk":"✓ {name} — {rows} 种截面，{pieces} 件","platesIgnored":" · 已忽略 {n} 行钢板","thKgM":"kg/m","andMore":"…还有 {n} 行","loLeftoverLen":"余料长度 (mm)","loBarLeft":"剩余型钢的长度","loHowManyLike":"同样的有几根","loSameProfile":"余料只用于相同截面和牌号的零件 — 牌号留空即可匹配无牌号的零件。先从每根余料中切出最合适的零件，剩余部分再采购最少的新原材。","whatToOrderTtl":"⬡ 订货清单","youNeedWord":"需要","gradeWord":"牌号","totalAcrossProf":"合计：{n} 根，{g} 种截面","leftoversUsedHd":"♻ 优先使用余料","leftoversUsedBd":"采购新料前已用掉 {n} 根余料","steelNotBought":"无需采购的钢材","longMembersHd":"⚙ 超长构件 — 由 {len} 原材拼接","eachFromBars":" — 每件由 {n} 根 {len} 原材拼接","cutFromLeftTtl":"♻ 由您的余料切出 — 无需采购","cutFromLeftLead":"这些零件来自您已有的余料，不计入上方的采购清单。","leftRemaining":" 此后余料剩余：{kg}。","visualCutPlan":"下料示意图","noPlan":"无方案。","allFromLeft":"所有零件均由您的余料切出 — 无需采购新料。","reusableBarTtl":"♻ 可再利用型钢余料 — 留给下一个项目","reusableBarLead":"{n} 根 ≥ 1 m 的料头 ≈ {kg} — 留到下一个项目下料，无需再买新料。","thFromBar":"来源原材","thStockLength":"原材长度","thReusableLeft":"可再利用余料","barNum":"第 {n} 根","leftoverNum":"余料 {n}","stBarsBuy":"需采购根数","stTotalLen":"总长度","stOffcut":"余料","stWaste":"损耗","stUtil":"利用率","stWeight":"重量","cbBar":"原材 #{n}","cbAvailable":"可用 {len}","cbUsed":"已用 {p}%","canvLoffcut":"♻ L 形余料","canvOffcut":"♻ 余料","canvUse":"可用 {w}×{h}","weldLbl":"焊缝 {w}×{h}","sheetTab":"第 {n} 张 · {p}%","reuseListLead":"♻ 可再利用余料：{list}","lShapeUse":"L 形，可用 ","notchWord":"缺角","liveNesting":"● 实时排版 — 真实案例","scrapReuse":"废料/余料","usableWord":"可用","widthArrow":"↔ 宽度","lengthArrow":"↕ 长度","cutoutWord":"缺口","lThickNote":"厚 {t} mm · 绿色 = 零件排入的矩形","lpIntro":"站在余料的一个角，把它记为 0, 0。沿轮廓走一圈，依次输入经过的每个角点 — 图形会随之生成。","lpColX":"X · 宽度 →","lpColY":"Y · 长度 ↑","lpAxisX":"X · 宽度","lpAxisY":"Y · 长度","lpReset":"↺ 重置形状","lpOverall":"外形","lpUsable":"可用","lpTip":"提示：L 形有 6 个角点，构成一个规整的台阶","lpHint0":"起始角 — 保持 0, 0","lpHint1":"→ 沿底边","lpHint2":"↑ 沿侧边向上","lpHint3":"← 进入缺角","lpHint4":"↑ 向上到顶","lpHint5":"← 返回起点","priceCardTtl":"💰 钢材价格与成本（可选）","priceCardLead":"输入每吨钢材价格 — 成本会显示在结果以及 PDF/Excel 报告中。留空则不计算成本。","priceCurrency":"币种","pricePerTonLbl":"每吨价格","costPurchase":"采购总成本","costNet":"实际用料价值","costScrap":"废料成本（损耗）","costOffcut":"可再利用余料价值","costBasedOn":"成本依据：每吨 {p}（您的价格）。","paintHd":"🎨 涂装面积","paintNoteS":"按清单中的零件净长 × 每米外表面积计算。按直角计、不计圆角 → 偏保守（通常比型钢表高 2–5%）。不含余料和废料。","paintNoteP":"仅计单面 · 您的零件（长 × 宽 × 数量）· 不含余料和废料。","paintMissing":"未计入 — 截面库中无尺寸：{list}","stPaint":"涂装面积","paintOneFace":"涂装面积（单面）","m2":"m²","siteTables":"型钢截面表 — 每米重量与涂装面积","SL.btn":"截面库","SL.title":"型钢截面库","SL.pickFam":"选择截面类别","SL.search":"搜索截面…","SL.back":"返回","SL.close":"关闭","SL.sections":"种截面","SL.designation":"规格","SL.family":"类别","SL.unitMass":"每米重量","SL.area":"截面面积","SL.perBar":"每根原材重量","SL.perTonne":"每吨根数","SL.barsOf":"根，每根","SL.note":"面积按 7850 kg/m³ 由每米重量反算。下料方案和吨数请使用优化功能。","SL.useIt":"用此截面进行优化 ↓","SL.paint":"涂装面积（保守值）","SL.page":"完整数据页 ↗","FAM.IPE":"IPE — 欧洲工字钢","FAM.HEA":"HEA — 宽翼缘（轻型）","FAM.HEB":"HEB — 宽翼缘（中型）","FAM.HEM":"HEM — 宽翼缘（重型）","FAM.IPN":"IPN — 斜翼缘工字钢","FAM.UB":"UB — 英国梁","FAM.UC":"UC — 英国柱","FAM.UBP":"UBP — 英国承载桩","FAM.PFC":"PFC — 英国槽钢","FAM.UPN":"UPN — 斜翼缘槽钢","FAM.UPE":"UPE — 平行翼缘槽钢","FAM.W":"W — 美国宽翼缘型钢","FAM.C-AMER":"C — 美国槽钢","FAM.MC":"MC — 美国杂项槽钢","FAM.L":"L — 角钢","FAM.SHS":"SHS — 方管","FAM.RHS":"RHS — 矩形管","FAM.CHS":"CHS — 圆管","FAM.C-COLD":"C — 冷弯檩条","FAM.Z-COLD":"Z — 冷弯檩条","FAM.JIS-HW":"HW — 宽翼缘 H 型钢（GB/T 11263 · KS D 3502）","FAM.JIS-HM":"HM — 中翼缘 H 型钢（GB/T 11263 · KS D 3502）","FAM.JIS-HN":"HN — 窄翼缘 H 型钢（GB/T 11263 · KS D 3502）","FAM.JIS-I":"I — 斜翼缘工字钢（JIS G 3192）","FAM.JIS-C":"C — 斜翼缘槽钢（JIS G 3192）","FAM.JIS-LIP":"LC — 冷弯卷边槽钢（JIS G 3350）","FAM.GOST-I":"GOST 工字钢","FAM.GOST-C":"GOST 槽钢","FAM.GOST-B":"GOST Б 宽翼缘","FAM.GOST-K":"GOST К 柱用","FAM.GB-I":"GB 工字钢（中国）","FAM.GB-C":"GB 槽钢（中国）","FAM.HSS":"HSS — 美国空心结构型钢","FAM.HP":"HP — 美国承载桩","FAM.S-AMER":"S — 美国标准工字钢","FAM.M-AMER":"M — 美国杂项工字钢","DIMSRC":"尺寸来源：AISC Shapes DB v16.0 · EN 10365 · British Steel (BS EN 10365:2017) · JIS G 3192 / KS / GB 系列。每个数值均与每米重量交叉核对；公式已用 379 个公开截面面积验证（最大误差 0.75%）。","MISC.langMenu":"语言","MISC.dimDepth":"高度 h","MISC.dimFlangeW":"翼缘宽度 b","MISC.dimWebT":"腹板厚度 tw","MISC.dimFlangeT":"翼缘厚度 tf","MISC.dimLegA":"边宽 a","MISC.dimLegB":"边宽 b","MISC.dimThick":"厚度 t","MISC.dimHeight":"高度 h","MISC.dimWidth":"宽度 b","MISC.dimWall":"壁厚 t","MISC.dimOD":"外径 D","MISC.dimStd":"标准","MISC.asianNote":"注：JIS G 3192、KS D 3502 和 GB/T 11263 中的 h、b、tw、tf 相同 — 仅圆角半径不同，重量差最多 2%。此处重量按 KS/GB。","MISC.intlSrc":"来源：SAP2000 截面库（{file}）· {std}。重量 = 截面面积 × 7850 kg/m³。面积与尺寸不一致的条目已剔除。","MISC.gEU":"欧洲（EN 10025）","MISC.gUS":"美国（ASTM）","MISC.gCN":"中国（GB）","MISC.gJP":"日本（JIS）","MISC.gOT":"其他（CSA / DIN）","MISC.tbSave":"保存为 PDF","MISC.tbPrint":"打印","MISC.tbHint":"两个按钮都会打开浏览器的打印窗口。如需 PDF，请把<b>目标打印机</b>设为“另存为 PDF”。","MISC.errTitle":"发生意外错误 — 请刷新页面。","MISC.errReload":"刷新","MISC.reportFail":"无法生成报告，请重试。","unknownRows":"⚠ 未识别为型钢 — 未参与优化：{list}。请检查这些名称，或把文件发给我们。","fromFile":"文件中写作：{name}","fromFileTag":"（重量取自您的文件）","fileKgmNote":"文件：{v} kg/m（{d}%）","fileKgmWarn":"⚠ 文件：{v} kg/m（{d}%）— 请核对该截面","fileKgmCell":"文件 {v}","MISC.barOf":""},"es":{"appName":"Optimizador de corte y anidado de acero","langName":"Español","otherLang":"English","heroBadge":"Para reducir el desperdicio de acero","heroTitleA":"Optimizador de corte","heroTitleB":"y anidado de acero","heroSub":"Chapas y perfiles • Mínimo desperdicio • Tonelaje al instante","heroCta":"EMPEZAR A OPTIMIZAR →","heroScroll":"↓ desplácese para empezar","showTitle":"Vea cómo se optimiza el corte de una chapa","chipImport":"Importación de Tekla · Excel · CSV","chipDb":"Más de 1100 perfiles — EU · US · JIS · GOST · GB","chipOffcut":"Registro de sobrantes en L","chipTon":"Tonelaje y lista de compra al instante","chipReports":"Informes en PDF y Excel","chooseWhat":"Elija qué desea optimizar","ctTop":"¿El archivo no se lee bien? Envíenoslo por correo","ctTitle":"¿El archivo no se importó correctamente?","ctBody":"Envíenos el archivo de Excel: lo procesamos y le devolvemos el plan de corte. Gratis.","ctEmail":"Correo","ctWhats":"WhatsApp","ctWhatsHint":"Lo más rápido: envíe una foto de la lista si le resulta más fácil","ctBy":"Hecho por un ingeniero estructural, para quienes cortan acero.","ctBlocked":"¿El sitio está bloqueado en la red de su oficina? Algunos cortafuegos corporativos bloquean dominios nuevos. Ábralo con datos móviles o escríbame y le envío el resultado.","preNoticeTtl":"Gratis durante el acceso anticipado","preNoticeBody":"Steel Optimizer pasará a una suscripción de pago en los próximos meses. Todo lo que usa hoy sigue siendo gratis hasta entonces, y los primeros usuarios conservarán un precio de fundador.","preNoticeCta":"Avíseme cuando se lance","preNoticeSent":"✓ Anotado: usted será de los primeros en saberlo.","preNoticeMail":"su@correo.com","proTitle":"La exportación es una función Pro","proBody":"El plano de anidado, el porcentaje de desperdicio y el registro de sobrantes siguen gratis en pantalla. Pro añade los informes de corte descargables en PDF y Excel.","proMonthly":"Mensual","proYearly":"Anual","proPerMo":"/mes","proPerYr":"/año","proTrial":"Primer mes gratis","proSave":"2 meses gratis","proHaveKey":"Ya tengo una clave","proKeyPh":"Pegue su clave de licencia","proActivate":"Activar","proChecking":"Comprobando…","proPrivacy":"Sus archivos nunca salen de su navegador. No se sube nada a ningún servidor.","proErrEmpty":"Introduzca su clave de licencia.","proErrNet":"No se pudo conectar con el servidor de licencias. Compruebe su conexión.","proErrLimit":"Esta clave ya está activa en el número máximo de dispositivos.","proErrExpired":"Esta suscripción ha finalizado. Renuévela para recuperar la exportación.","proErrBad":"No se reconoció la clave. Compruebe que no falten caracteres.","proOk":"✓ Pro activo: exportación desbloqueada.","proPay":"Suscribirse","proLoading":"Abriendo el pago seguro…","proCards":"Visa · Mastercard · Amex · Apple Pay · Google Pay · PayPal","proSecure":"Pago seguro con Lemon Squeezy. Los datos de la tarjeta nunca pasan por este sitio.","proRestore":"¿Pagó en otro dispositivo?","proThanks":"Pago recibido: la exportación está desbloqueada. Gracias.","modSections":"Perfiles de acero","modSectionsDesc":"Vigas, columnas, angulares, perfiles Z y C conformados en frío y secciones armadas.","modPlates":"Chapas y placas de acero","modPlatesDesc":"Anidado de piezas sobre chapas comerciales (2D). Primero los sobrantes, luego el mínimo de chapas nuevas, con el tonelaje total.","chooseArrow":"Elegir →","howEnter":"¿Cómo desea introducir su lista de corte?","manualEntry":"Entrada manual","uploadTitle":"Cargue su lista de materiales","scanTitle":"Escanear un PDF o una foto","scanDescP":"Lea las medidas de las piezas de chapa desde una lista en PDF o una foto de una lista impresa.","scanDescS":"Lea perfiles y longitudes desde una lista en PDF o una foto de una lista impresa.","manualDescP":"Escriba las piezas en una tabla sencilla.","manualDescS":"Escriba un perfil (IPE, UB, HEA, SHS…), elíjalo de la lista y anote las longitudes.","uploadDescP":"Suelte una lista de materiales. Se leen las chapas, se ignoran los perfiles y se limpian las hojas desordenadas.","uploadDescS":"Suelte una lista de materiales. Se detectan los perfiles; las chapas se ignoran.","changeMethod":"↺ Cambiar método","backModules":"↩ Volver","backInput":"← Volver a los datos","lblThickness":"Espesor (mm)","lblWidth":"Ancho (mm)","lblLength":"Longitud (mm)","lblQty":"Cant.","lblQtyHow":"¿Cuántas? (cant.)","lblGrade":"Grado de acero","lblKerf":"Ancho de corte / sangría (mm)","lblMargin":"Margen de borde (mm)","lblReuseMin":"Sobrante reutilizable mín. (mm)","lblRotate":"Permitir giro de 90°","lblId":"ID","lblProfile":"Perfil","lblGradeShort":"Grado","platesWS":"Chapas de acero — Área de trabajo","stockSettings":"📐 Chapas en stock y parámetros de corte","stockSheets":"Chapas comerciales — defina un tamaño de chapa por espesor","addThickness":"+ Añadir espesor","stockHint":"Si hay piezas de más de 2440 mm, defina aquí una chapa mayor (p. ej., 1500×3000 o 2000×6000) para ese espesor; de lo contrario no se podrán anidar.","spliceStrat":"Empalme (piezas mayores que la chapa)","spliceWelds":"Menos soldaduras (trozos más grandes)","splicePack":"Mejor aprovechamiento (más trozos, más pequeños)","reusePlatesTtl":"♻ Reutilizar chapas sobrantes (opcional)","reuseBarsTtl":"♻ Reutilizar barras sobrantes (opcional)","reusePlatesLead":"¿Tiene chapas sobrantes de un trabajo anterior o de otra obra? Añádalas aquí (o importe un archivo de sobrantes anterior) y cortaremos primero de ellas, para que solo compre el mínimo de chapas nuevas que aún necesita.","reuseBarsLead":"¿Tiene barras sobrantes de un trabajo anterior o de otra obra? Añádalas aquí (o importe un archivo de sobrantes anterior) y cortaremos primero de ellas, para que solo compre el mínimo de barras nuevas que aún necesita.","noPlatesYet":"Todavía no se han añadido chapas sobrantes.","noBarsYet":"Todavía no se han añadido barras sobrantes.","addLeftPlate":"+ Añadir una chapa sobrante","addLeftBar":"+ Añadir una barra sobrante","reuseImportBtn":"⤓ Reutilizar un archivo de sobrantes anterior (PDF / Excel)","reuseReading":"Leyendo","reuseBusy":"⏳ Leyendo…","manualPartEntry":"✏️ Entrada manual de piezas","addPart":"+ Añadir pieza","uploadTeklaTtl":"📊 Excel","dropTekla":"Suelte la lista de Tekla / Excel / CSV / TXT","orBrowse":"o haga clic para buscar","scanCardTtl":"🔍 PDF / Foto","scanDrop":"Suelte un PDF o una foto de su lista","scanAccept":"Admite .pdf · .png · .jpg · .webp — o haga clic para buscar","scanWarn":"⚠ Leer un PDF o una foto es una estimación, no una medición. Revise cada fila antes de optimizar: un número erróneo aquí significa acero mal cortado.","scanReading":"Leyendo el archivo…","scanOcr":"Leyendo el texto… {pct}%","scanNoRows":"No se pudo leer ninguna fila de este archivo. Una foto más plana y mejor iluminada — o el Excel original — funcionará mucho mejor.","scanFail":"No se pudo leer este archivo: {msg}","scanFound":"Filas leídas: {n}. Revíselas y corríjalas antes de usarlas.","scanUse":"✓ Usar estas filas","scanAddRow":"+ Añadir fila","scanRedo":"↺ Elegir otro archivo","scanTip":"Una foto plana y bien iluminada de una lista impresa se lee mucho mejor que un croquis a mano.","scanUsedOk":"✓ Filas tomadas del escaneo: {n}. Revíselas antes de optimizar.","optPlates":"⚡ OPTIMIZAR CHAPAS","optSections":"⚡ OPTIMIZAR PERFILES","optimizing":"⚙ Optimizando…","optResults":"Resultados de la optimización","summaryBuy":"✦ RESUMEN — LO QUE DEBE COMPRAR","wasteAfterCut":"MATERIAL SIN USAR TRAS EL CORTE","utilization":"Aprovechamiento","sheet":"chapa","sheets":"chapas","sheetOf":"de","totalSheets":"TOTAL","scrapWeight":"PESO DE CHATARRA / SOBRANTES","reusableShort":"reutilizables","totalParts":"Total de piezas","totalSheetsCard":"Total de chapas","thkGroups":"Grupos de espesor","purchaseWt":"Peso total (a comprar)","netWt":"Peso neto (usado en obra)","reusableOffcuts":"Sobrantes reutilizables","finalOutput":"Resultado final listo para decidir","procSummary":"Resumen de compra de material","material":"MATERIAL","whatToBuy":"Lo que debe comprar","youNeed":"Necesita","bar":"barra","bars":"barras","whatToOrder":"⬡ QUÉ PEDIR","barsToBuy":"Barras a comprar","totalLength":"Longitud total","offcut":"Sobrante","waste":"Desperdicio","weight":"Peso","offcutLength":"LONGITUD DE SOBRANTES","dlReports":"💾 Descargar informes y archivos","dlProcPDF":"📄 Compras en PDF","dlProcXLS":"📊 Compras en Excel","dlLeftPDF":"♻ Sobrantes en PDF","dlLeftXLS":"♻ Sobrantes en Excel","gateTitle":"Obtenga su informe de corte","gateBody":"Introduzca su correo una sola vez para descargar sus archivos de compras y de corte. Solo lo usaremos para contactarle sobre esta herramienta. Sin spam y sin compartirlo.","gateBtn":"Descargar informe →","gateSaving":"Guardando…","gateEmailErr":"Introduzca una dirección de correo válida.","mm":"mm","t_ton":"t","kg":"kg","thKgMu":"kg/m","m2pmU":"m²/m","gU":"g","computedTag":" (calculado)","splicedRun":"(tramo empalmado)","scSaved":"AHORRADO","scYours":"para el próximo trabajo: dinero que se queda en casa","scCaption":"1220 × 2440 mm · corte de 6 mm · 8 piezas","scKept":"Sobrante en L conservado ≈ 56 kg","scNesting":"anidando… {p}% de la chapa","nameThisPlate":"Nombre de esta chapa","removeX":"✕ quitar","whatShape":"¿Qué forma tiene el sobrante?","shapeRect":"▭ Rectángulo","shapeL":"⌐ En L (con una esquina recortada)","willNestInto":"✓ Anidaremos las piezas en un área de","biggestRectL":" (el mayor rectángulo que cabe en la L)","pieceWord":"pieza","piecesWord":"piezas","addLeftMore":"+ Añadir un sobrante de otro trabajo / obra","notInStock":"(sin stock)","upBannerP":"💡 El espesor se toma del archivo y el tamaño de chapa correspondiente, de su lista de arriba. Si el espesor es desconocido, se usa 1220×2440.","upHintP":"Lee la columna de perfil (PLT/FLT/PL → espesor × ancho). Se ignoran los perfiles laminados.","mapColumns":"Asigne las columnas:","applyMapping":"Aplicar asignación","partsReady":"PIEZAS LISTAS","moreWord":"más","msgImportedP":"✓ Tipos de pieza de chapa importados: {n}. Se ignoraron perfiles laminados y filas de subtotal.","msgMapCols":"⚠ No se pudieron detectar las columnas automáticamente. Asígnelas abajo.","msgReadFail":"⚠ No se pudo leer este archivo.","msgImportedP2":"✓ Piezas de chapa importadas: {n}","msgExtractFail":"⚠ Aún no se pudieron extraer los datos.","liReading":"Leyendo {name}…","liWrongType":"⚠ Parece un archivo de sobrantes de {type}. Úselo en el módulo {mod}.","liImported":"✓ Tipos de sobrante importados: {n} ({m} uds.). Se cortarán primero.","liErr":"⚠ {msg}","modPlatesName":"Chapas","modSectionsName":"Perfiles","sheetUnitOf":"{n} × {sw} × {sh} mm ({mat})","acrossThk":"TOTAL: {n} chapa{s} en {g} espesor{es}","spliceBannerHd":"⚙ Piezas mayores que la chapa: {n} — se cortan en trozos y se sueldan","spliceBannerBd":"Cada una se reparte entre varias chapas y se une con una costura soldada (no hace falta una chapa mayor). Antes de fabricar, confirme que se admite una soldadura en estas piezas:","reusedBannerP":"♻ Se usaron primero sobrantes de proyectos anteriores: {n}. Las cantidades de chapa de abajo son el mínimo de chapas comerciales nuevas que aún debe comprar.","nestLayoutTtl":"🗂 Plano de anidado — chapa de {thk} mm ({n} chapa{s})","offcutRegTtl":"♻ Registro de sobrantes reutilizables","offcutRegLead":"Sobrantes reutilizables: {n} ≈ {t} t para conservar en futuros trabajos (≥ {min} mm por lado). En los sobrantes en L, APROVECHABLE es el mayor rectángulo que se puede cortar.","thQty":"CANT.","thThickness":"ESPESOR","thFromSheet":"DE LA CHAPA","thShape":"FORMA","thOverall":"DIM. TOTALES","thUsableRect":"RECT. APROVECHABLE","thWeight":"PESO","thStatus":"ESTADO","shapeLshort":"⌐ En L","shapeRectShort":"▭ Rect.","keepStatus":"✓ conservar","fromSheetN":"Chapa {n}","dlTwoDeliv":"Dos entregables: los archivos de compras (qué comprar + plan de corte) y los de sobrantes (sobrantes reutilizables para conservar).","miniSheet":"Chapa","miniThickness":"Espesor","miniParts":"Piezas","miniUtil":"Aprovechamiento","reusableLeftover":"♻ Sobrante reutilizable","pbFinalReady":"Resultado final listo para decidir","pbProcSummary":"Resumen de compra de material","thSheetSize":"TAMAÑO DE CHAPA","thSheetsReq":"CHAPAS NEC.","thPurchaseWt":"PESO DE COMPRA (chapas enteras)","thParts":"PIEZAS","thNetWt":"PESO NETO PIEZAS","thScrapWt":"PESO CHATARRA","pbLegend":"● Peso total = chapas enteras que compra · ● Peso neto = acero usado en la obra · ● Chatarra = sobrante (total − neto)","pbTotalPurchase":"Peso total de compra","pbGrossNote":"bruto — las {n} chapas enteras","pbNetNote":"acero que queda en la estructura","pbScrapNote":"{p}% de la compra","pbGenerated":"Generado con el Optimizador de corte y anidado de acero","thicknessesWord":"espesor{es}","sectionsWS":"Perfiles de acero — Área de trabajo","manualCutList":"✏️ Lista de corte manual","thLengths":"Longitudes (mm)","thMarketLen":"Longitud comercial disponible (mm)","addProfile":"+ AÑADIR PERFIL","piecesCount":"{n} uds.","lengthsHint":"p. ej. 3000 4500 6000x3  →  una de 3 m, una de 4,5 m, tres de 6 m","blank12m":"vacío = 12 m","secManualHelp":"Escriba IPE, HEA, HEB, UB, UC, JIS HW/HM/HN, RHS, SHS, CHS, C/Z conformados en frío, GOST y GB. Viga armada: escriba 900x400x20x15. Longitudes: sepárelas con un espacio, p. ej. 3000 4500 6000x3; escriba 6000x3 para tres piezas de 6 m.","computedSuffix":"(calculado)","upTitleS":"Cargue su lista de materiales","upBodyS":"Suelte una lista de materiales de Tekla, o cualquier lista con longitudes y cantidades de acero.","upAcceptS":"Admite .xlsx · .xls · .csv · .txt — o haga clic para buscar","upHintS":"Los perfiles se detectan automáticamente · las filas de chapa se ignoran (use el módulo Chapas) · la longitud comercial es 12 m por defecto y se puede editar.","secUpErr":"No se detectaron perfiles estructurales. Si el archivo solo contiene chapas, use el módulo Chapas.","secReadErr":"Error de lectura: {msg}","secNoValid":"No hay cortes válidos.","secEnterLen":"Introduzca al menos una longitud.","uploadedOk":"✓ {name} — perfiles: {rows}, piezas: {pieces}","platesIgnored":" · filas de chapa ignoradas: {n}","thKgM":"kg/m","andMore":"…y {n} más","loLeftoverLen":"Longitud del sobrante (mm)","loBarLeft":"longitud de la barra que le queda","loHowManyLike":"cuántas iguales","loSameProfile":"Un sobrante solo sirve para piezas del mismo perfil y grado; deje el grado vacío para combinarlo con piezas sin grado. Primero se corta de cada sobrante la pieza que mejor encaja y luego se compra el mínimo de barras nuevas para el resto.","whatToOrderTtl":"⬡ QUÉ PEDIR","youNeedWord":"NECESITA","gradeWord":"grado","totalAcrossProf":"TOTAL: barras: {n} · perfiles: {g}","leftoversUsedHd":"♻ SOBRANTES USADOS PRIMERO","leftoversUsedBd":"Barras sobrantes suyas cortadas antes de comprar nuevas: {n}","steelNotBought":"ACERO NO COMPRADO","longMembersHd":"⚙ ELEMENTOS LARGOS — EMPALMADOS DE BARRAS DE {len}","eachFromBars":" — cada uno con {n} barras de {len}","cutFromLeftTtl":"♻ Cortado de sus sobrantes — sin compra","cutFromLeftLead":"Estas piezas salen de barras sobrantes que ya tenía; no están en la lista de compra de arriba.","leftRemaining":" Sobrante que aún queda después: {kg}.","visualCutPlan":"Plan de corte visual","noPlan":"Sin plan.","allFromLeft":"Todas las piezas se cortaron de sus sobrantes: no hay que comprar nada nuevo.","reusableBarTtl":"♻ Sobrantes de barra reutilizables — acero para el próximo proyecto","reusableBarLead":"Extremos de barra ≥ 1 m: {n} ≈ {kg}. Consérvelos para cortar en su próximo trabajo en lugar de comprar nuevos.","thFromBar":"DE LA BARRA","thStockLength":"LONGITUD COMERCIAL","thReusableLeft":"SOBRANTE REUTILIZABLE","barNum":"Barra {n}","leftoverNum":"Sobrante {n}","stBarsBuy":"Barras a comprar","stTotalLen":"Longitud total","stOffcut":"Sobrante","stWaste":"Desperdicio","stUtil":"Aprovechamiento","stWeight":"Peso","cbBar":"BARRA N.º {n}","cbAvailable":"disponible {len}","cbUsed":"{p}% usado","canvLoffcut":"♻ SOBRANTE EN L","canvOffcut":"♻ SOBRANTE","canvUse":"usar {w}×{h}","weldLbl":"SOLDADURA {w}×{h}","sheetTab":"Chapa {n} · {p}%","reuseListLead":"♻ Sobrantes reutilizables: {list}","lShapeUse":"En L, usar ","notchWord":"recorte","liveNesting":"● ANIDADO EN VIVO — CASO REAL","scrapReuse":"chatarra/reutilización","usableWord":"APROVECHABLE","widthArrow":"↔ ancho","lengthArrow":"↕ longitud","cutoutWord":"recorte","lThickNote":"{t} mm de espesor · verde = el rectángulo donde anidamos","lpIntro":"Sitúese en una esquina del sobrante y llámela 0, 0. Recorra el contorno y escriba cada esquina que encuentre: el dibujo se construye sobre la marcha.","lpColX":"X · ancho →","lpColY":"Y · longitud ↑","lpAxisX":"X · ancho","lpAxisY":"Y · longitud","lpReset":"↺ Restablecer forma","lpOverall":"Dim. totales","lpUsable":"aprovechable","lpTip":"consejo: una L tiene 6 esquinas que forman un escalón limpio","lpHint0":"esquina inicial: déjela en 0, 0","lpHint1":"→ a lo largo del borde inferior","lpHint2":"↑ subiendo por el lado","lpHint3":"← hacia el recorte","lpHint4":"↑ hasta arriba","lpHint5":"← de vuelta al inicio","priceCardTtl":"💰 Precio del acero y costo (opcional)","priceCardLead":"Introduzca el precio del acero por tonelada: los costos aparecerán en los resultados y en los informes PDF/Excel. Déjelo vacío para no calcular precios.","priceCurrency":"Moneda","pricePerTonLbl":"Precio por tonelada","costPurchase":"COSTO TOTAL DE COMPRA","costNet":"VALOR NETO UTILIZADO","costScrap":"COSTO DE CHATARRA (DESPERDICIO)","costOffcut":"VALOR DE SOBRANTES REUTILIZABLES","costBasedOn":"Base de costo: {p} por tonelada (su precio).","paintHd":"🎨 SUPERFICIE DE PINTURA","paintNoteS":"Longitudes netas de corte de su lista × superficie exterior por metro. Esquinas vivas, sin radios de acuerdo → conservador (normalmente 2–5% por encima del catálogo). No incluye sobrantes ni chatarra.","paintNoteP":"Solo una cara · sus piezas (longitud × ancho × cant.) · no incluye sobrantes ni chatarra.","paintMissing":"No incluido — sin dimensiones en la biblioteca: {list}","stPaint":"Superficie de pintura","paintOneFace":"Superficie de pintura (1 cara)","m2":"m²","siteTables":"Tablas de perfiles de acero — peso y superficie de pintura por metro","SL.btn":"Tabla de perfiles","SL.title":"Biblioteca de perfiles de acero","SL.pickFam":"Elija una familia de perfiles","SL.search":"Buscar perfiles…","SL.back":"Volver","SL.close":"Cerrar","SL.sections":"perfiles","SL.designation":"Designación","SL.family":"Familia","SL.unitMass":"Peso por metro","SL.area":"Área de la sección","SL.perBar":"Peso por barra comercial","SL.perTonne":"Barras por tonelada","SL.barsOf":"barras de","SL.note":"Área obtenida del peso por metro a 7850 kg/m³. Use el optimizador para planes de corte y tonelaje.","SL.useIt":"Optimizar con este perfil ↓","SL.paint":"Superficie de pintura (conservadora)","SL.page":"Página completa de datos ↗","FAM.IPE":"IPE — Vigas I europeas","FAM.HEA":"HEA — Ala ancha (ligera)","FAM.HEB":"HEB — Ala ancha (media)","FAM.HEM":"HEM — Ala ancha (pesada)","FAM.IPN":"IPN — Vigas I de alas inclinadas","FAM.UB":"UB — Vigas británicas","FAM.UC":"UC — Columnas británicas","FAM.UBP":"UBP — Pilotes británicos","FAM.PFC":"PFC — Canales británicos","FAM.UPN":"UPN — Canales de alas inclinadas","FAM.UPE":"UPE — Canales de alas paralelas","FAM.W":"W — Ala ancha americana","FAM.C-AMER":"C — Canales americanos","FAM.MC":"MC — Canales americanos varios","FAM.L":"L — Angulares","FAM.SHS":"SHS — Tubo cuadrado","FAM.RHS":"RHS — Tubo rectangular","FAM.CHS":"CHS — Tubo redondo","FAM.C-COLD":"C — Correas conformadas en frío","FAM.Z-COLD":"Z — Correas conformadas en frío","FAM.JIS-HW":"HW — H de ala ancha (GB/T 11263 · KS D 3502)","FAM.JIS-HM":"HM — H de ala media (GB/T 11263 · KS D 3502)","FAM.JIS-HN":"HN — H de ala estrecha (GB/T 11263 · KS D 3502)","FAM.JIS-I":"I — Vigas I de alas inclinadas (JIS G 3192)","FAM.JIS-C":"C — Canales de alas inclinadas (JIS G 3192)","FAM.JIS-LIP":"LC — Canales con labios conformados en frío (JIS G 3350)","FAM.GOST-I":"Vigas I GOST","FAM.GOST-C":"Canales GOST","FAM.GOST-B":"GOST Б — ala ancha","FAM.GOST-K":"GOST К — columnas","FAM.GB-I":"Vigas I GB (China)","FAM.GB-C":"Canales GB (China)","FAM.HSS":"HSS — Perfiles tubulares americanos","FAM.HP":"HP — Pilotes americanos","FAM.S-AMER":"S — Vigas I estándar americanas","FAM.M-AMER":"M — Vigas I americanas varias","DIMSRC":"Dimensiones: AISC Shapes DB v16.0 · EN 10365 · British Steel (BS EN 10365:2017) · series JIS G 3192 / KS / GB. Cada valor se contrastó con el peso por metro; la fórmula se verificó con 379 áreas de sección publicadas (error máx. 0,75%).","MISC.langMenu":"Idioma","MISC.dimDepth":"Canto h","MISC.dimFlangeW":"Ancho de ala b","MISC.dimWebT":"Espesor del alma tw","MISC.dimFlangeT":"Espesor del ala tf","MISC.dimLegA":"Lado a","MISC.dimLegB":"Lado b","MISC.dimThick":"Espesor t","MISC.dimHeight":"Altura h","MISC.dimWidth":"Ancho b","MISC.dimWall":"Espesor de pared t","MISC.dimOD":"Diámetro exterior D","MISC.dimStd":"Norma","MISC.asianNote":"Nota: h, b, tw, tf son idénticos en JIS G 3192, KS D 3502 y GB/T 11263; solo cambia el radio de acuerdo, lo que da hasta un 2% de diferencia de peso. Aquí los pesos siguen KS/GB.","MISC.intlSrc":"Fuente: base de secciones de SAP2000 ({file}) · {std}. Peso = área de la sección × 7850 kg/m³. Se excluyeron las filas cuya área no concuerda con sus dimensiones.","MISC.gEU":"Europeos (EN 10025)","MISC.gUS":"Americanos (ASTM)","MISC.gCN":"Chinos (GB)","MISC.gJP":"Japoneses (JIS)","MISC.gOT":"Otros (CSA / DIN)","MISC.tbSave":"Guardar como PDF","MISC.tbPrint":"Imprimir","MISC.tbHint":"Ambos abren el cuadro de impresión del navegador. Para un PDF, elija en <b>Destino</b> «Guardar como PDF».","MISC.errTitle":"Se produjo un error inesperado. Vuelva a cargar la página.","MISC.errReload":"Volver a cargar","MISC.reportFail":"No se pudo generar el informe. Inténtelo de nuevo.","unknownRows":"⚠ No reconocido como perfil — no optimizado: {list}. Revise estos nombres o envíenos el archivo.","fromFile":"tal como figura en su archivo: {name}","fromFileTag":" (peso tomado de su archivo)","fileKgmNote":"archivo: {v} kg/m ({d}%)","fileKgmWarn":"⚠ archivo: {v} kg/m ({d}%) — revise este perfil","fileKgmCell":"archivo {v}","MISC.barOf":"de"}};

const RT_I18N_SRC = ["Steel (grade not specified)","pdf reader unavailable","No Steel Optimizer data found in this PDF. Re-export the Leftover PDF and try again.","This HTML file has no Steel Optimizer leftover data.","Couldn't recognise this file as a leftover list.","No plate leftovers found in this file.","No bar leftovers found in this file.","text reader unavailable","Hi, I'm using Steel Optimizer and I need help with my material list.","Export failed: ","Steel Optimizer — Plate Procurement","Material","Thickness (mm)","Sheet Size","Sheets Req.","Total Wt (t)","Parts","Utilization %","Net Wt (t)","Scrap Wt (t)","Paint Area 1 face (m²)","TOTAL","TOTAL PURCHASE WEIGHT - gross full sheets (Ton)","NET PARTS WEIGHT - finished plates only (Ton)","SCRAP / OFFCUT WEIGHT (Ton)","REUSABLE OFFCUTS","pieces","t","PAINT AREA - one face of your parts, L x W x qty (m²)","offcuts and scrap not included","STEEL PRICE (per ton)","your entered price","TOTAL PURCHASE COST - gross full sheets","NET USED VALUE - finished plates only","SCRAP COST - money lost in waste","REUSABLE OFFCUT VALUE - recoverable","Summary","From Sheet","Offcut W (mm)","Offcut L (mm)","Weight (kg)","Status","keep for future jobs","No reusable offcuts — all leftovers below the minimum size.","Reusable Offcuts","CUTTING PLAN — part placements per sheet (origin = top-left corner, mm)","Sheet #","Part Label","X (mm)","Y (mm)","Width (mm)","Length (mm)","Rotated 90°","Spliced/Welded","YES","no","No placements to list.","Cutting Plan","Reusable Offcut Register","Minimum size kept (each side)","mm","Shape","Overall W (mm)","Overall L (mm)","Notch W (mm)","Notch L (mm)","Usable Rect W (mm)","Usable Rect L (mm)","L-shape","Rectangle","No reusable offcuts above the minimum size.","TOTAL REUSABLE OFFCUTS"," kg","Steel Optimizer re-import data — keep this sheet.","Drop this file into the “Reuse a previous Leftover file” button to load these leftovers again.","Sheet","notch"," g","usable","Steel Optimizer — Reusable Offcuts","minimum kept side","Leftover shapes","Register","Thickness","Overall","Usable Rect","Weight","Re-import code — drop this PDF back into the optimizer's <b>“Reuse a previous Leftover file”</b> button to load these leftovers again.","Total purchase cost:","Net used:","Scrap cost:","Reusable offcut value:","Based on","ton","Rect","kg","Steel Optimizer — Plate Report","Steel Optimizer — Plate Procurement Summary","What you need to buy","Sheets","Total Wt","Utilization","Net Wt","Scrap Wt","Paint (1 face)"," m&sup2;","thk","Total Wt = full sheets purchased &middot; Net Wt = steel used in project &middot; Scrap Wt = offcut (Total - Net) &middot; Paint = one face of your parts (L &times; W &times; qty), offcuts and scrap excluded","Total Purchase Weight (gross):","Ton","sheets","Net:","Scrap:","Paint area (one face):","m&sup2;","Nesting layouts","Amber dashed = welded/spliced piece &middot; <b>Green = reusable leftover</b> (L-shape or rectangle).","STEEL OPTIMIZER — SECTION OPTIMIZATION REPORT","KEY FIGURES","Total Weight to Purchase (t)","Net Weight used in Project (t)","Scrap / Offcut Weight (t)","Total Bars to Buy","Profile Groups","Paint Area (m²) — net cut lengths, outside faces, square corners (conservative)","PER-PROFILE BREAKDOWN","Profile","Grade","kg/m","Bars to Buy","Total Length (m)","Net Length (m)","Offcut (m)","Total Weight (t)","Paint m²/m","Paint Area (m²)","TOTAL PURCHASE COST - all raw material incl. waste","NET USED VALUE","SCRAP COST - money in waste","Total Weight = full bars purchased · Net Weight = steel used in project · Offcut = Total − Net","Paint Area = net cut lengths from your list × outside surface per metre (square corners, root radii ignored — typically 2–5% above catalogue). Offcuts and scrap not included.","WHAT TO ORDER","Market Length (mm)","Qty to Order","Weight (t)","TOTAL STEEL TO PROCURE (t)","Procurement","CUTTING PLAN — bar by bar","Cut Map key:  █ = cut piece   |  = saw cut   ▒ = reusable leftover (≥1 m)   · = scrap","Bar #","Available Length (mm)","Cuts (mm)","Pieces","Leftover (mm)","Leftover Reusable?","Cut Map (to scale)","YES (≥1 m)","REUSABLE BAR OFFCUTS — leftover ends ≥ 1 m, worth keeping for the next job","From Bar #","Reusable Leftover (mm)","No reusable bar offcuts ≥ 1 m — material was used efficiently.","TOTAL REUSABLE","LONG MEMBERS — SPLICED FROM SHORTER BARS","Member Length (mm)","Qty","Bars per Member","From Bar (mm)","Spliced Members"," keep","purchase cost covers all raw material bought, including the waste portion.","Bar","used","Steel Optimizer — Section Report","Steel Optimizer — Section Procurement & Cutting","kerf","Total Steel to Procure:","bars","waste","Paint area:","Procurement — What to Order","Market Length","Total Weight","Paint Area","m&sup2;/m","Net cut length","Paint area"," m","Net cut lengths from your list &times; outside surface per metre &middot; square corners, root radii ignored (typically 2&ndash;5% above catalogue) &middot; offcuts and scrap not included","Visual Cut Plan","Each bar drawn to scale &middot; coloured blocks = cut pieces (length in mm) &middot; <b>green dashed = reusable leftover &ge; 1 m</b> &middot; grey dashed = scrap &middot; amber = spliced run.","Cutting Plan — Bar by Bar","Stock Length","Cuts","Offcut","Util","STEEL OPTIMIZER — REUSABLE BAR OFFCUTS","Leftover bar ends ≥ 1 m — keep these to cut from on your next job instead of buying new.","Keep?","Drop this file into the “Reuse a previous Leftover file” button to load these offcuts again.","keep","No reusable bar offcuts &ge; 1 m — material was used efficiently.","Reusable Bar Offcuts — leftover steel for the next project","Reusable leftover:","m","From Bar","Available Length","Reusable Leftover","Re-import code — drop this PDF back into the optimizer's <b>“Reuse a previous Leftover file”</b> button to load these offcuts again."];
const RT_I18N_TX = {"ru":["Сталь (марка не указана)","модуль чтения PDF недоступен","В этом PDF нет данных Steel Optimizer. Экспортируйте PDF остатков заново и повторите попытку.","В этом HTML-файле нет данных об остатках Steel Optimizer.","Не удалось распознать файл как список остатков.","В файле не найдены остатки листов.","В файле не найдены остатки профилей.","модуль распознавания текста недоступен","Здравствуйте! Я пользуюсь Steel Optimizer, и мне нужна помощь со спецификацией.","Ошибка экспорта: ","Steel Optimizer — закупка листов","Материал","Толщина (мм)","Размер листа","Листов","Масса общая (т)","Детали","Использование, %","Масса нетто (т)","Отходы (т)","Площадь окраски, 1 сторона (м²)","ИТОГО","ОБЩАЯ МАССА ЗАКУПКИ — брутто, целые листы (т)","МАССА ДЕТАЛЕЙ НЕТТО — только готовые детали (т)","МАССА ОТХОДОВ / ОСТАТКОВ (т)","ДЕЛОВЫЕ ОСТАТКИ","шт.","т","ПЛОЩАДЬ ОКРАСКИ — одна сторона деталей, Д × Ш × кол-во (м²)","остатки и отходы не включены","ЦЕНА МЕТАЛЛА (за тонну)","указанная вами цена","ОБЩАЯ СТОИМОСТЬ ЗАКУПКИ — брутто, целые листы","СТОИМОСТЬ МЕТАЛЛА В ДЕЛЕ — только готовые детали","СТОИМОСТЬ ОТХОДОВ — потери на отходах","СТОИМОСТЬ ДЕЛОВЫХ ОСТАТКОВ — можно вернуть в работу","Сводка","С листа","Остаток Ш (мм)","Остаток Д (мм)","Масса (кг)","Статус","сохранить для будущих заказов","Деловых остатков нет — все остатки меньше минимального размера.","Деловые остатки","КАРТА РАСКРОЯ — размещение деталей на листах (начало координат = верхний левый угол, мм)","Лист №","Марка детали","X (мм)","Y (мм)","Ширина (мм)","Длина (мм)","Повёрнута на 90°","Со стыком / сварная","ДА","нет","Нет размещений для вывода.","Карта раскроя","Реестр деловых остатков","Минимальный сохраняемый размер (каждая сторона)","мм","Форма","Габарит Ш (мм)","Габарит Д (мм)","Вырез Ш (мм)","Вырез Д (мм)","Полезный прямоуг. Ш (мм)","Полезный прямоуг. Д (мм)","Г-образный","Прямоугольник","Нет деловых остатков больше минимального размера.","ИТОГО ДЕЛОВЫХ ОСТАТКОВ"," кг","Данные для повторного импорта Steel Optimizer — не удаляйте этот лист.","Перетащите этот файл на кнопку «Загрузить прежний файл остатков», чтобы снова загрузить эти остатки.","Лист","вырез"," г","полезно","Steel Optimizer — деловые остатки","минимальная сохраняемая сторона","Формы остатков","Реестр","Толщина","Габарит","Полезный прямоуг.","Масса","Код повторного импорта — перетащите этот PDF обратно на кнопку <b>«Загрузить прежний файл остатков»</b>, чтобы снова загрузить эти остатки.","Общая стоимость закупки:","В деле:","Стоимость отходов:","Стоимость деловых остатков:","Исходя из","т","Прямоуг.","кг","Steel Optimizer — отчёт по листам","Steel Optimizer — сводка закупки листов","Что нужно закупить","Листов","Масса общая","Использование","Масса нетто","Отходы","Окраска (1 сторона)"," м&sup2;","толщ.","Масса общая = закупаемые целые листы &middot; Масса нетто = металл, вошедший в проект &middot; Отходы = остатки (общая − нетто) &middot; Окраска = одна сторона деталей (Д &times; Ш &times; кол-во), без остатков и отходов","Общая масса закупки (брутто):","т","лист.","Нетто:","Отходы:","Площадь окраски (одна сторона):","м&sup2;","Карты раскроя","Жёлтый пунктир = деталь со сварным стыком &middot; <b>Зелёный = деловой остаток</b> (Г-образный или прямоугольный).","STEEL OPTIMIZER — ОТЧЁТ ПО РАСКРОЮ ПРОФИЛЕЙ","ОСНОВНЫЕ ПОКАЗАТЕЛИ","Масса к закупке (т)","Масса нетто в проекте (т)","Масса отходов / остатков (т)","Всего хлыстов к закупке","Групп профилей","Площадь окраски (м²) — длины деталей нетто, наружная поверхность, прямые углы (с запасом)","РАЗБИВКА ПО ПРОФИЛЯМ","Профиль","Марка","кг/м","Хлыстов к закупке","Общая длина (м)","Длина нетто (м)","Остаток (м)","Масса общая (т)","Окраска м²/м","Площадь окраски (м²)","ОБЩАЯ СТОИМОСТЬ ЗАКУПКИ — весь металл, включая отходы","СТОИМОСТЬ МЕТАЛЛА В ДЕЛЕ","СТОИМОСТЬ ОТХОДОВ — деньги в отходах","Масса общая = закупаемые целые хлысты · Масса нетто = металл, вошедший в проект · Остаток = общая − нетто","Площадь окраски = длины деталей нетто из вашей ведомости × наружная поверхность на метр (углы прямые, радиусы сопряжения не учитываются — обычно на 2–5% выше сортамента). Остатки и отходы не включены.","ЧТО ЗАКАЗАТЬ","Длина проката (мм)","Кол-во к заказу","Масса (т)","ВСЕГО МЕТАЛЛА К ЗАКУПКЕ (т)","Закупка","КАРТА РАСКРОЯ — по хлыстам","Условные обозначения:  █ = деталь   |  = рез   ▒ = деловой остаток (≥1 м)   · = отходы","Хлыст №","Длина хлыста (мм)","Резы (мм)","Деталей","Остаток (мм)","Деловой остаток?","Схема раскроя (в масштабе)","ДА (≥1 м)","ДЕЛОВЫЕ ОСТАТКИ ПРОФИЛЕЙ — концы ≥ 1 м, стоит сохранить для следующего заказа","Из хлыста №","Деловой остаток (мм)","Деловых остатков ≥ 1 м нет — металл использован эффективно.","ИТОГО ДЕЛОВЫХ","ДЛИННЫЕ ЭЛЕМЕНТЫ — СО СТЫКОМ ИЗ БОЛЕЕ КОРОТКИХ ХЛЫСТОВ","Длина элемента (мм)","Кол-во","Хлыстов на элемент","Из хлыста (мм)","Стыкуемые элементы"," сохранить","стоимость закупки включает весь купленный металл, в том числе отходы.","Хлыст","использовано","Steel Optimizer — отчёт по профилям","Steel Optimizer — закупка и раскрой профилей","рез","Всего металла к закупке:","хлыст.","отходы","Площадь окраски:","Закупка — что заказать","Длина проката","Масса общая","Площадь окраски","м&sup2;/м","Длина деталей нетто","Площадь окраски"," м","Длины деталей нетто из вашей ведомости &times; наружная поверхность на метр &middot; углы прямые, радиусы сопряжения не учитываются (обычно на 2&ndash;5% выше сортамента) &middot; остатки и отходы не включены","Наглядная карта раскроя","Каждый хлыст показан в масштабе &middot; цветные блоки = детали (длина в мм) &middot; <b>зелёный пунктир = деловой остаток &ge; 1 м</b> &middot; серый пунктир = отходы &middot; жёлтый = стык.","Карта раскроя — по хлыстам","Длина хлыста","Резы","Остаток","Исп.","STEEL OPTIMIZER — ДЕЛОВЫЕ ОСТАТКИ ПРОФИЛЕЙ","Концевые остатки ≥ 1 м — сохраните их для следующего заказа вместо закупки нового металла.","Сохранить?","Перетащите этот файл на кнопку «Загрузить прежний файл остатков», чтобы снова загрузить эти остатки.","сохранить","Деловых остатков &ge; 1 м нет — металл использован эффективно.","Деловые остатки профилей — металл для следующего проекта","Деловые остатки:","м","Из хлыста","Длина хлыста","Деловой остаток","Код повторного импорта — перетащите этот PDF обратно на кнопку <b>«Загрузить прежний файл остатков»</b>, чтобы снова загрузить эти остатки."],"zh":["钢材（牌号未指定）","PDF 读取组件不可用","此 PDF 中没有 Steel Optimizer 数据。请重新导出余料 PDF 后再试。","此 HTML 文件中没有 Steel Optimizer 余料数据。","无法将此文件识别为余料清单。","此文件中未找到钢板余料。","此文件中未找到型钢余料。","文字识别组件不可用","您好，我在使用 Steel Optimizer，材料表需要帮助。","导出失败：","Steel Optimizer — 钢板采购","材质","厚度 (mm)","板幅","所需张数","总重 (t)","零件","利用率 %","净重 (t)","废料重 (t)","涂装面积（单面）(m²)","合计","采购总重 — 整张钢板毛重 (t)","零件净重 — 仅成品零件 (t)","废料 / 余料重量 (t)","可再利用余料","块","t","涂装面积 — 零件单面，长 × 宽 × 数量 (m²)","不含余料和废料","钢材价格（每吨）","您输入的价格","采购总成本 — 整张钢板毛重","实际用料价值 — 仅成品零件","废料成本 — 损耗的金额","可再利用余料价值 — 可回收","汇总","来源钢板","余料宽 (mm)","余料长 (mm)","重量 (kg)","状态","保留供后续项目使用","无可再利用余料 — 所有余料均小于最小尺寸。","可再利用余料","下料方案 — 各钢板零件位置（原点 = 左上角，mm）","钢板编号","零件编号","X (mm)","Y (mm)","宽度 (mm)","长度 (mm)","旋转 90°","拼接/焊接","是","否","没有可列出的排版位置。","下料方案","可再利用余料清单","保留的最小尺寸（每边）","mm","形状","外形宽 (mm)","外形长 (mm)","缺角宽 (mm)","缺角长 (mm)","可用矩形宽 (mm)","可用矩形长 (mm)","L 形","矩形","没有大于最小尺寸的可再利用余料。","可再利用余料合计"," kg","Steel Optimizer 重新导入数据 — 请保留此工作表。","将此文件拖到“导入以前的余料文件”按钮，即可再次载入这些余料。","钢板","缺角"," g","可用","Steel Optimizer — 可再利用余料","保留的最小边长","余料形状","清单","厚度","外形","可用矩形","重量","重新导入代码 — 将此 PDF 拖回优化工具的<b>“导入以前的余料文件”</b>按钮，即可再次载入这些余料。","采购总成本：","实际用料：","废料成本：","可再利用余料价值：","依据","吨","矩形","kg","Steel Optimizer — 钢板报告","Steel Optimizer — 钢板采购汇总","需要采购的材料","张数","总重","利用率","净重","废料重","涂装（单面）"," m&sup2;","厚","总重 = 采购的整张钢板 &middot; 净重 = 用于工程的钢材 &middot; 废料重 = 余料（总重 - 净重）&middot; 涂装 = 零件单面（长 &times; 宽 &times; 数量），不含余料和废料","采购总重（毛重）：","t","张","净重：","废料：","涂装面积（单面）：","m&sup2;","排版图","琥珀色虚线 = 焊接/拼接件 &middot; <b>绿色 = 可再利用余料</b>（L 形或矩形）。","STEEL OPTIMIZER — 型钢下料优化报告","关键数据","采购总重 (t)","工程用净重 (t)","废料 / 余料重量 (t)","需采购原材总根数","截面组数","涂装面积 (m²) — 零件净长、外表面、按直角计（保守）","按截面分项","截面","牌号","kg/m","需采购根数","总长度 (m)","净长度 (m)","余料 (m)","总重 (t)","涂装 m²/m","涂装面积 (m²)","采购总成本 — 全部原材料（含损耗）","实际用料价值","废料成本 — 损耗的金额","总重 = 采购的整根原材 · 净重 = 用于工程的钢材 · 余料 = 总重 − 净重","涂装面积 = 清单中零件净长 × 每米外表面积（按直角计、不计圆角 — 通常比型钢表高 2–5%）。不含余料和废料。","订货清单","市场供货长度 (mm)","订货数量","重量 (t)","需采购钢材合计 (t)","采购","下料方案 — 逐根原材","图例：  █ = 零件   |  = 锯口   ▒ = 可再利用余料（≥1 m）   · = 废料","原材编号","可用长度 (mm)","切割长度 (mm)","件数","余料 (mm)","余料可再利用？","下料图（按比例）","是（≥1 m）","可再利用型钢余料 — ≥ 1 m 的料头，值得留给下一个项目","来源原材编号","可再利用余料 (mm)","没有 ≥ 1 m 的可再利用型钢余料 — 材料已充分利用。","可再利用合计","超长构件 — 由较短原材拼接","构件长度 (mm)","数量","每件所需原材数","所用原材 (mm)","拼接构件"," 保留","采购成本包含购入的全部原材料（含损耗部分）。","原材","已用","Steel Optimizer — 型钢报告","Steel Optimizer — 型钢采购与下料","割缝","需采购钢材合计：","根","损耗","涂装面积：","采购 — 订货清单","市场供货长度","总重","涂装面积","m&sup2;/m","零件净长","涂装面积"," m","清单中零件净长 &times; 每米外表面积 &middot; 按直角计、不计圆角（通常比型钢表高 2&ndash;5%）&middot; 不含余料和废料","下料示意图","每根原材按比例绘制 &middot; 彩色块 = 零件（长度，mm）&middot; <b>绿色虚线 = 可再利用余料 &ge; 1 m</b> &middot; 灰色虚线 = 废料 &middot; 琥珀色 = 拼接段。","下料方案 — 逐根原材","原材长度","切割","余料","利用率","STEEL OPTIMIZER — 可再利用型钢余料","≥ 1 m 的料头 — 留到下一个项目下料，无需再买新料。","保留？","将此文件拖到“导入以前的余料文件”按钮，即可再次载入这些余料。","保留","没有 &ge; 1 m 的可再利用型钢余料 — 材料已充分利用。","可再利用型钢余料 — 留给下一个项目的钢材","可再利用余料：","m","来源原材","可用长度","可再利用余料","重新导入代码 — 将此 PDF 拖回优化工具的<b>“导入以前的余料文件”</b>按钮，即可再次载入这些余料。"],"es":["Acero (grado no especificado)","lector de PDF no disponible","No se encontraron datos de Steel Optimizer en este PDF. Vuelva a exportar el PDF de sobrantes e inténtelo de nuevo.","Este archivo HTML no contiene datos de sobrantes de Steel Optimizer.","No se pudo reconocer este archivo como una lista de sobrantes.","No se encontraron chapas sobrantes en este archivo.","No se encontraron barras sobrantes en este archivo.","lector de texto no disponible","Hola, estoy usando Steel Optimizer y necesito ayuda con mi lista de materiales.","Error al exportar: ","Steel Optimizer — Compra de chapas","Material","Espesor (mm)","Tamaño de chapa","Chapas nec.","Peso total (t)","Piezas","Aprovechamiento %","Peso neto (t)","Peso chatarra (t)","Superficie de pintura, 1 cara (m²)","TOTAL","PESO TOTAL DE COMPRA - bruto, chapas enteras (t)","PESO NETO DE PIEZAS - solo piezas terminadas (t)","PESO DE CHATARRA / SOBRANTES (t)","SOBRANTES REUTILIZABLES","uds.","t","SUPERFICIE DE PINTURA - una cara de sus piezas, L x A x cant. (m²)","no incluye sobrantes ni chatarra","PRECIO DEL ACERO (por tonelada)","precio introducido por usted","COSTO TOTAL DE COMPRA - bruto, chapas enteras","VALOR NETO UTILIZADO - solo piezas terminadas","COSTO DE CHATARRA - dinero perdido en desperdicio","VALOR DE SOBRANTES REUTILIZABLES - recuperable","Resumen","De la chapa","Ancho sobrante (mm)","Largo sobrante (mm)","Peso (kg)","Estado","conservar para futuros trabajos","Sin sobrantes reutilizables: todos quedan por debajo del tamaño mínimo.","Sobrantes reutilizables","PLAN DE CORTE — posición de las piezas en cada chapa (origen = esquina superior izquierda, mm)","Chapa n.º","Marca de pieza","X (mm)","Y (mm)","Ancho (mm)","Longitud (mm)","Girada 90°","Empalmada/soldada","SÍ","no","No hay posiciones que listar.","Plan de corte","Registro de sobrantes reutilizables","Tamaño mínimo conservado (cada lado)","mm","Forma","Ancho total (mm)","Largo total (mm)","Ancho del recorte (mm)","Largo del recorte (mm)","Ancho rect. aprovechable (mm)","Largo rect. aprovechable (mm)","En L","Rectángulo","No hay sobrantes reutilizables por encima del tamaño mínimo.","TOTAL DE SOBRANTES REUTILIZABLES"," kg","Datos de reimportación de Steel Optimizer: conserve esta hoja.","Suelte este archivo en el botón «Reutilizar un archivo de sobrantes anterior» para volver a cargar estos sobrantes.","Chapa","recorte"," g","aprovechable","Steel Optimizer — Sobrantes reutilizables","lado mínimo conservado","Formas de los sobrantes","Registro","Espesor","Dim. totales","Rect. aprovechable","Peso","Código de reimportación: suelte este PDF de nuevo en el botón <b>«Reutilizar un archivo de sobrantes anterior»</b> del optimizador para volver a cargar estos sobrantes.","Costo total de compra:","Neto utilizado:","Costo de chatarra:","Valor de sobrantes reutilizables:","Según","tonelada","Rect.","kg","Steel Optimizer — Informe de chapas","Steel Optimizer — Resumen de compra de chapas","Lo que debe comprar","Chapas","Peso total","Aprovechamiento","Peso neto","Chatarra","Pintura (1 cara)"," m&sup2;","esp.","Peso total = chapas enteras compradas &middot; Peso neto = acero usado en la obra &middot; Chatarra = sobrante (total - neto) &middot; Pintura = una cara de sus piezas (L &times; A &times; cant.), sin sobrantes ni chatarra","Peso total de compra (bruto):","t","chapas","Neto:","Chatarra:","Superficie de pintura (una cara):","m&sup2;","Planos de anidado","Discontinuo ámbar = pieza soldada/empalmada &middot; <b>Verde = sobrante reutilizable</b> (en L o rectangular).","STEEL OPTIMIZER — INFORME DE OPTIMIZACIÓN DE PERFILES","CIFRAS CLAVE","Peso total a comprar (t)","Peso neto usado en la obra (t)","Peso de chatarra / sobrantes (t)","Total de barras a comprar","Grupos de perfiles","Superficie de pintura (m²) — longitudes netas de corte, caras exteriores, esquinas vivas (conservador)","DESGLOSE POR PERFIL","Perfil","Grado","kg/m","Barras a comprar","Longitud total (m)","Longitud neta (m)","Sobrante (m)","Peso total (t)","Pintura m²/m","Superficie de pintura (m²)","COSTO TOTAL DE COMPRA - toda la materia prima, incl. desperdicio","VALOR NETO UTILIZADO","COSTO DE CHATARRA - dinero en desperdicio","Peso total = barras enteras compradas · Peso neto = acero usado en la obra · Sobrante = total − neto","Superficie de pintura = longitudes netas de corte de su lista × superficie exterior por metro (esquinas vivas, sin radios de acuerdo; normalmente 2–5% por encima del catálogo). No incluye sobrantes ni chatarra.","QUÉ PEDIR","Longitud comercial (mm)","Cant. a pedir","Peso (t)","TOTAL DE ACERO A COMPRAR (t)","Compras","PLAN DE CORTE — barra por barra","Leyenda:  █ = pieza cortada   |  = corte de sierra   ▒ = sobrante reutilizable (≥1 m)   · = chatarra","Barra n.º","Longitud disponible (mm)","Cortes (mm)","Piezas","Sobrante (mm)","¿Sobrante reutilizable?","Mapa de corte (a escala)","SÍ (≥1 m)","SOBRANTES DE BARRA REUTILIZABLES — extremos ≥ 1 m que vale la pena conservar para el próximo trabajo","De la barra n.º","Sobrante reutilizable (mm)","No hay sobrantes de barra reutilizables ≥ 1 m: el material se aprovechó bien.","TOTAL REUTILIZABLE","ELEMENTOS LARGOS — EMPALMADOS DE BARRAS MÁS CORTAS","Longitud del elemento (mm)","Cant.","Barras por elemento","De barra (mm)","Elementos empalmados"," conservar","el costo de compra incluye toda la materia prima comprada, incluida la parte desperdiciada.","Barra","usado","Steel Optimizer — Informe de perfiles","Steel Optimizer — Compra y corte de perfiles","corte","Total de acero a comprar:","barras","desperdicio","Superficie de pintura:","Compras — Qué pedir","Longitud comercial","Peso total","Superficie de pintura","m&sup2;/m","Longitud neta de corte","Superficie de pintura"," m","Longitudes netas de corte de su lista &times; superficie exterior por metro &middot; esquinas vivas, sin radios de acuerdo (normalmente 2&ndash;5% por encima del catálogo) &middot; no incluye sobrantes ni chatarra","Plan de corte visual","Cada barra dibujada a escala &middot; bloques de color = piezas cortadas (longitud en mm) &middot; <b>discontinuo verde = sobrante reutilizable &ge; 1 m</b> &middot; discontinuo gris = chatarra &middot; ámbar = tramo empalmado.","Plan de corte — barra por barra","Longitud comercial","Cortes","Sobrante","Aprov.","STEEL OPTIMIZER — SOBRANTES DE BARRA REUTILIZABLES","Extremos de barra ≥ 1 m: consérvelos para cortar en su próximo trabajo en lugar de comprar nuevos.","¿Conservar?","Suelte este archivo en el botón «Reutilizar un archivo de sobrantes anterior» para volver a cargar estos sobrantes.","conservar","No hay sobrantes de barra reutilizables &ge; 1 m: el material se aprovechó bien.","Sobrantes de barra reutilizables — acero para el próximo proyecto","Sobrante reutilizable:","m","De la barra","Longitud disponible","Sobrante reutilizable","Código de reimportación: suelte este PDF de nuevo en el botón <b>«Reutilizar un archivo de sobrantes anterior»</b> del optimizador para volver a cargar estos sobrantes."]};

const soIsExtra = (l) => l === "ru" || l === "zh" || l === "es";
// interface text for ru/zh/es; `def` (the English) otherwise or when missing
function soUiTx(lang, key, def) {
  const d = soIsExtra(lang) ? I18N_UI[lang] : null;
  return d && d[key] != null ? d[key] : def;
}
// bilingual lines (English · Arabic) keep their Arabic half in en/ar; ru/zh/es show their own
function soLoc(key, arText) { return soUiTx(SO_LANG, key, arText); }
// one language only (report print bar, alerts, error screen): Arabic in ar, otherwise the UI language (English fallback)
function soTb(key, en, ar) { return SO_LANG === "ar" ? ar : soUiTx(SO_LANG, key, en); }
let _soRtMap = null;
function soRtI18n(en, x) {
  try {
    if (!soIsExtra(SO_LANG)) return en;
    if (typeof x === "function") { const v = x(SO_LANG); if (v != null) return v; }
    if (!_soRtMap) {
      _soRtMap = {};
      for (const L of ["ru", "zh", "es"]) { const m = new Map(), tx = RT_I18N_TX[L] || []; RT_I18N_SRC.forEach((e, i) => { if (tx[i] != null) m.set(e, tx[i]); }); _soRtMap[L] = m; }
    }
    const m = _soRtMap[SO_LANG];
    return m && m.has(en) ? m.get(en) : en;
  } catch { return en; }
}
// Russian plural: 1 лист · 2 листа · 5 листов
function soRuPlural(n, one, few, many) {
  const a = Math.abs(Math.round(+n || 0)) % 100, b = a % 10;
  return a > 10 && a < 20 ? many : b === 1 ? one : b >= 2 && b <= 4 ? few : many;
}
// «2 листа», «5 хлыстов» in Russian; every other language exactly as t(n > 1 ? many : one)
const SO_RU_FORMS = { sheet: ["лист", "листа", "листов"], bar: ["хлыст", "хлыста", "хлыстов"] };
function soCountWord(t, n, one, many) {
  if (SO_LANG === "ru" && SO_RU_FORMS[one]) return soRuPlural(n, ...SO_RU_FORMS[one]);
  return t(n > 1 ? many : one);
}
// first letter up (Russian/Spanish family chips read «Трубы квадратные», not «трубы квадратные»)
const soCap = (s, lang) => (soIsExtra(lang) && typeof s === "string" ? s.charAt(0).toUpperCase() + s.slice(1) : s);
// recognised-section badge for an international row: its designation and standard, not the internal family key
function soIntlBadge(sec) {
  const std = sec.std && !String(sec.name).includes(sec.std) ? ` · ${sec.std}` : "";
  return `${sec.name}${std}`;
}
const soDateLoc = (def) => (soIsExtra(SO_LANG) ? SO_DATE_LOC[SO_LANG] : def);
const soHtmlLangAttr = () => ` lang="${SO_LANG_TAG[SO_LANG] || "en"}"`;
function soBrowserLang() {
  if (!AUTO_LANG_DETECT || !I18N_EXTRA_ENABLED) return null;
  try {
    const list = (navigator.languages && navigator.languages.length ? navigator.languages : [navigator.language]) || [];
    for (const l of list) {                                    // the first language the browser prefers decides
      const c = String(l || "").toLowerCase().slice(0, 2);
      if (c === "ru" || c === "zh" || c === "es") return c;
      if (c) return null;                                      // English, Arabic or anything else → English as before
    }
  } catch { /* no navigator */ }
  return null;
}
const SO_GRADE_GROUP_KEY = { "European (EN 10025)": "MISC.gEU", "American (ASTM)": "MISC.gUS", "Chinese (GB)": "MISC.gCN", "Japanese (JIS)": "MISC.gJP", "Other (CSA / DIN)": "MISC.gOT" };
const SO_GRADE_RU = ["Russian (GOST)", ["С235", "С245", "С255", "С345", "С345К", "С355", "С355-1", "С355К", "С355П", "С390", "С390-1", "С440", "С550", "С590", "С690", "09Г2С", "10ХСНД", "Ст3сп5", "Ст3пс5"]];
function soGradeGroup(grp, lang) {
  if (grp === SO_GRADE_RU[0]) return lang === "ru" ? "Российские (ГОСТ)" : "Russian (GOST)";
  return SO_GRADE_GROUP_KEY[grp] ? soUiTx(lang, SO_GRADE_GROUP_KEY[grp], grp) : grp;
}
// grade picker: Russian interface → GOST grades first; Chinese → GB first; others unchanged
function soGradeGroups(groups) {
  if (SO_LANG === "ru") return [SO_GRADE_RU, ...groups];
  if (SO_LANG === "zh") { const cn = groups.filter(g => g[0] === "Chinese (GB)"); return [...cn, ...groups.filter(g => g[0] !== "Chinese (GB)")]; }
  return groups;
}
// Space Mono has no Cyrillic or Chinese glyphs: give those characters a proper fallback
function soLangFonts(lang) {
  try {
    const H = document.head; if (!H) return;
    const old = document.getElementById("so-i18n-font"); if (old) old.remove();
    const oldL = document.getElementById("so-i18n-font-link"); if (oldL) oldL.remove();
    if (lang === "ru") {
      const l = document.createElement("link"); l.id = "so-i18n-font-link"; l.rel = "stylesheet";
      l.href = "https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&display=swap"; H.appendChild(l);
    }
    const stack = lang === "ru" ? "'Space Mono','JetBrains Mono',Consolas,Menlo,'DejaVu Sans Mono',monospace"
      : lang === "zh" ? "'Space Mono','PingFang SC','Microsoft YaHei','Noto Sans SC','Source Han Sans SC',sans-serif" : null;
    if (!stack) return;
    const st = document.createElement("style"); st.id = "so-i18n-font";
    st.textContent = `[style*="Space Mono"]{font-family:${stack}!important}` + (lang === "zh"
      ? "body{font-family:'DM Sans','PingFang SC','Microsoft YaHei','Noto Sans SC',system-ui,sans-serif}[style*=\"Playfair Display\"]{font-family:'Playfair Display','PingFang SC','Microsoft YaHei','Noto Sans SC',serif!important}"
      : "");
    H.appendChild(st);
  } catch { /* no document */ }
}
/* ── language menu (replaces the two-language toggle) ── */
function SoLangMenu() {
  const { lang, setLang } = useLang();
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return undefined;
    const away = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    const esc = (e) => { if (e.key === "Escape") setOpen(false); };
    document.addEventListener("mousedown", away); document.addEventListener("touchstart", away); document.addEventListener("keydown", esc);
    return () => { document.removeEventListener("mousedown", away); document.removeEventListener("touchstart", away); document.removeEventListener("keydown", esc); };
  }, [open]);
  const side = "right";   // same corner in every language (Section Library stays top-left)
  const title = lang === "ar" ? "اللغة" : soUiTx(lang, "MISC.langMenu", "Language");
  return (
    <div ref={ref} style={{ position: "fixed", top: 16, [side]: 16, zIndex: 10000 }}>
      <button onClick={() => setOpen(o => !o)} title={title} aria-label={title} aria-haspopup="listbox" aria-expanded={open}
        style={{
          display: "inline-flex", alignItems: "center", gap: 8, padding: "8px 14px 8px 16px",
          borderRadius: 30, cursor: "pointer",
          background: "rgba(19,25,32,.82)", backdropFilter: "blur(10px)", WebkitBackdropFilter: "blur(10px)",
          border: "1px solid rgba(245,158,11,.45)", color: "#fbbf24",
          fontFamily: "'Space Mono', monospace", fontSize: 12, fontWeight: 700, letterSpacing: 1,
          boxShadow: "0 8px 24px -8px rgba(0,0,0,.6)",
        }}>
        <span aria-hidden>🌐</span>
        <span lang={SO_LANG_TAG[lang] || lang}>{SO_LANG_NAMES[lang] || lang}</span>
        <span aria-hidden style={{ fontSize: 9, opacity: 0.8, transform: open ? "rotate(180deg)" : "none", transition: "transform .15s" }}>▼</span>
      </button>
      {open && (
        <div role="listbox" aria-label={title} style={{
          position: "absolute", top: "calc(100% + 8px)", [side]: 0, minWidth: 176, padding: 6, borderRadius: 14,
          background: "rgba(13,20,32,.97)", backdropFilter: "blur(10px)", WebkitBackdropFilter: "blur(10px)",
          border: "1px solid rgba(245,158,11,.35)", boxShadow: "0 18px 44px -12px rgba(0,0,0,.85)",
        }}>
          {SO_LANGS.map(code => {
            const on = code === lang;
            return (
              <button key={code} role="option" aria-selected={on} lang={SO_LANG_TAG[code] || code}
                onClick={() => { setOpen(false); if (!on) setLang(code); }}
                style={{
                  display: "flex", alignItems: "center", justifyContent: "space-between", gap: 14, width: "100%",
                  padding: "9px 12px", borderRadius: 9, border: "none", cursor: "pointer", textAlign: "start",
                  background: on ? "rgba(245,158,11,.14)" : "transparent", color: on ? "#fbbf24" : "#cbd5e1",
                  fontSize: 16, fontWeight: on ? 700 : 500, fontFamily: "inherit",
                }}>
                <span>{SO_LANG_NAMES[code]}</span>
                <span aria-hidden style={{ fontSize: 12.5, color: on ? "#fbbf24" : "#475569", letterSpacing: 1, fontFamily: "system-ui, sans-serif" }}>{on ? "✓" : code.toUpperCase()}</span>
              </button>
            );
          })}
          <SoAccountBox onDone={() => setOpen(false)} />
        </div>
      )}
    </div>
  );
}
/* ╚══ end of INTERNATIONAL INTERFACE add-on ══════════════════════════════════╝ */

const LANG_KEY = "steelopt_lang_v2"; // v2: reset any stale saved choice so the app opens in English by default

function readSavedLang() {
  try { const q = new URLSearchParams(window.location.search).get("lang"); if (SO_LANGS.includes(q)) return q; } catch { /* no window */ }
  try { const v = localStorage.getItem(LANG_KEY); if (SO_LANGS.includes(v)) return v; } catch { /* private mode */ }
  return soBrowserLang() || "en"; // default English (a Russian/Chinese/Spanish browser opens in its language); change to "ar" if your audience opens in Arabic first
}

const LangCtx = createContext({ lang: "en", setLang: () => {}, t: (k) => k, dir: "ltr" });

function LangProvider({ children }) {
  const [lang, setLangState] = useState(readSavedLang);
  const dir = lang === "ar" ? "rtl" : "ltr";
  SO_LANG = lang;                                   // formatters, reports, messages follow the UI language
  const setLang = useCallback((l) => {
    setLangState(l);
    try { localStorage.setItem(LANG_KEY, l); } catch { /* ignore */ }
  }, []);
  // keep <html dir/lang> in sync so native widgets, scrollbars & selects flip too
  useEffect(() => {
    if (typeof document !== "undefined") {
      document.documentElement.setAttribute("dir", dir);
      document.documentElement.setAttribute("lang", lang === "zh" ? "zh-CN" : lang);
      soLangFonts(lang);   // Cyrillic / Chinese font fallbacks: international add-on
    }
  }, [dir, lang]);
  const t = useCallback((key, vars) => {
    const entry = LANG_DICT[key];
    let s = entry ? (entry[lang] != null ? entry[lang] : soUiTx(lang, key, entry.en)) : key;
    if (vars) for (const k in vars) s = s.replace(new RegExp("\\{" + k + "\\}", "g"), vars[k]);
    return s;
  }, [lang]);
  return <LangCtx.Provider value={{ lang, setLang, t, dir }}>{children}</LangCtx.Provider>;
}

// hooks
const useLang = () => useContext(LangCtx);
const useT = () => useContext(LangCtx).t;

// Western (ASCII) numerals — do NOT localize to ٠١٢٣ for engineering output.
const fmtNum = (n) => String(n);

/* ── ONE-CLICK LANGUAGE TOGGLE (floating, amber, RTL-aware) ──────────────── */
function LangToggle() {
  const { lang, setLang } = useLang();
  const isAr = lang === "ar";
  return (
    <button
      onClick={() => setLang(isAr ? "en" : "ar")}
      title={isAr ? "التبديل إلى الإنجليزية" : "التبديل إلى العربية"}
      style={{
        position: "fixed", top: 16, right: 16, zIndex: 10000,
        display: "inline-flex", alignItems: "center", gap: 8, padding: "8px 16px",
        borderRadius: 30, cursor: "pointer",
        background: "rgba(19,25,32,.82)", backdropFilter: "blur(10px)", WebkitBackdropFilter: "blur(10px)",
        border: "1px solid rgba(245,158,11,.45)", color: "#fbbf24",
        fontFamily: "'Space Mono', monospace", fontSize: 12, fontWeight: 700, letterSpacing: 1,
        boxShadow: "0 8px 24px -8px rgba(0,0,0,.6)",
      }}>
      <span aria-hidden>🌐</span>
      {/* shows the OTHER language so the action is obvious */}
      {LANG_DICT.otherLang[lang]}
    </button>
  );
}

/* ════════════════════════════════════════════════════════════════════════════
   TRANSLATED SHELL COMPONENTS — paste these OVER your existing versions.
   (Everything else in your file keeps working; wire the two modules + the
    export reports with the same t() keys — all keys are already in LANG_DICT.)
════════════════════════════════════════════════════════════════════════════ */

/* SectionTitle — unchanged API, RTL handled by the global dir */
function SectionTitle({ children }) {
  const { dir } = useLang();
  return (
    <h2 style={{
      fontSize: 27, fontWeight: 900, color: "#f8fafc", margin: 0,
      fontFamily: "'Playfair Display', serif",
      [dir === "rtl" ? "borderRight" : "borderLeft"]: "4px solid #f59e0b",
      [dir === "rtl" ? "paddingRight" : "paddingLeft"]: 14,
    }}>{children}</h2>
  );
}

/* BackChip — pass a translation KEY (e.g. "backModules") OR a literal label */
function BackChip({ onClick, labelKey, label }) {
  const t = useT();
  return (
    <button onClick={onClick} style={{
      padding: "5px 14px", borderRadius: 20, border: "1px solid #2d3748",
      background: "rgba(15,19,24,.6)", color: "#94a3b8", cursor: "pointer",
      fontSize: 13, fontFamily: "'Space Mono', monospace",
    }}>{labelKey ? t(labelKey) : label}</button>
  );
}

/* Chooser — title/desc now accept translation KEYS */
function Chooser({ icon, titleKey, descKey, title, desc, onClick }) {
  const t = useT();
  return (
    <button onClick={onClick} style={{
      textAlign: "start", padding: "28px 26px", borderRadius: 16,
      background: "rgba(19,25,32,.72)", backdropFilter: "blur(12px)", WebkitBackdropFilter: "blur(12px)",
      border: "1px solid rgba(148,163,184,.16)", cursor: "pointer", transition: "all .25s ease",
      fontFamily: "inherit", color: "#cbd5e1", minHeight: 200, display: "flex", flexDirection: "column",
      boxShadow: "0 12px 40px -12px rgba(0,0,0,.5)",
    }}
      onMouseEnter={e => { e.currentTarget.style.borderColor = "#f59e0b"; e.currentTarget.style.transform = "translateY(-4px)"; e.currentTarget.style.boxShadow = "0 18px 50px -10px rgba(245,158,11,.25)"; }}
      onMouseLeave={e => { e.currentTarget.style.borderColor = "rgba(148,163,184,.16)"; e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "0 12px 40px -12px rgba(0,0,0,.5)"; }}>
      <div style={{ fontSize: 44 }}>{icon}</div>
      <div style={{ fontSize: 25, fontWeight: 700, color: "#f8fafc", margin: "16px 0 10px", fontFamily: "'Playfair Display', serif" }}>{titleKey ? t(titleKey) : title}</div>
      <div style={{ fontSize: 16, color: "#94a3b8", lineHeight: 1.5, flex: 1 }}>{descKey ? t(descKey) : desc}</div>
      <span style={{ color: "#f59e0b", fontFamily: "'Space Mono', monospace", fontSize: 15, fontWeight: 700, marginTop: 12 }}>{t("chooseArrow")}</span>
    </button>
  );
}

/* ModuleChooser — fully bilingual */
/* Early-access notice: states plainly that pricing is coming, and captures
   interest from the people who care. That signal is the point — it is worth
   more right now than the sign-ups themselves. */
function EarlyAccessNotice() {
  const t = useT();
  const [email, setEmail] = useState("");
  const [sent, setSent] = useState(false);
  const [busy, setBusy] = useState(false);
  const send = async () => {
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email.trim())) return;
    setBusy(true);
    try {
      await fetch("https://formspree.io/f/xqerwkze", {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify({ email: email.trim(), source: "early-access-waitlist", at: new Date().toISOString() }),
      });
    } catch { /* never block the app on a failed capture */ }
    setBusy(false); setSent(true);
  };
  return (
    <div style={{ marginBottom: 22, padding: "14px 18px", borderRadius: 12,
      background: "rgba(245,158,11,.06)", border: "1px solid rgba(245,158,11,.22)" }}>
      <div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, letterSpacing: 2,
        textTransform: "uppercase", color: "#fbbf24", marginBottom: 7 }}>{t("preNoticeTtl")}</div>
      <div style={{ fontSize: 15.5, lineHeight: 1.65, color: "#94a3b8", marginBottom: sent ? 0 : 12 }}>{t("preNoticeBody")}</div>
      {sent ? (
        <div style={{ marginTop: 10, fontSize: 15, color: "#6ee7b7" }}>{t("preNoticeSent")}</div>
      ) : (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <input value={email} onChange={e => setEmail(e.target.value)}
            onKeyDown={e => e.key === "Enter" && !busy && send()}
            placeholder={t("preNoticeMail")} spellCheck={false}
            style={{ ...IN, width: 230, flex: "0 1 230px" }} />
          <button onClick={send} disabled={busy}
            style={{ padding: "7px 18px", borderRadius: 4, border: "1px solid #d97706",
              background: "rgba(245,158,11,.15)", color: "#fbbf24", cursor: busy ? "default" : "pointer",
              fontSize: 14, opacity: busy ? .6 : 1 }}>{t("preNoticeCta")}</button>
        </div>
      )}
    </div>
  );
}

function ModuleChooser({ onPick }) {
  const t = useT();
  return (
    <div>
      <EarlyAccessNotice />
      <SectionTitle>{t("chooseWhat")}</SectionTitle>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))", gap: 20, marginTop: 18 }}>
        <Chooser icon="🅸" titleKey="modSections" descKey="modSectionsDesc" onClick={() => onPick("sections")} />
        <Chooser icon="▭" titleKey="modPlates" descKey="modPlatesDesc" onClick={() => onPick("plates")} />
      </div>
    </div>
  );
}

/* EmailGate — bilingual copy (logic identical to your original) */
function EmailGate() {
  const t = useT();
  const [open, setOpen] = useState(false);
  const [email, setEmail] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  useEffect(() => { gateBus.open = () => { setErr(""); setOpen(true); }; return () => { gateBus.open = null; }; }, []);

  const finish = () => {
    const fn = gateBus.fn; gateBus.fn = null;
    setOpen(false); setBusy(false); setEmail("");
    if (fn) setTimeout(fn, 120);
  };
  const submit = async () => {
    const v = email.trim();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) { setErr(t("gateEmailErr")); return; }
    setBusy(true); setErr("");
    try {
      if (FORMSPREE_ID && FORMSPREE_ID !== "YOUR_FORM_ID") {
        await fetch(`https://formspree.io/f/${FORMSPREE_ID}`, {
          method: "POST",
          headers: { "Content-Type": "application/json", Accept: "application/json" },
          body: JSON.stringify({ email: v, source: "steel-optimizer", at: new Date().toISOString() }),
        });
      }
    } catch { /* never block on a network hiccup */ }
    try { localStorage.setItem(GATE_KEY, v); } catch { /* private mode */ }
    finish();
  };
  if (!open) return null;
  return (
    <div onClick={() => !busy && setOpen(false)}
      style={{ position: "fixed", inset: 0, zIndex: 9999, background: "rgba(4,7,12,.82)", backdropFilter: "blur(6px)", WebkitBackdropFilter: "blur(6px)", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div onClick={e => e.stopPropagation()}
        style={{ width: "100%", maxWidth: 440, background: "linear-gradient(160deg,#161b22,#0c1016)", border: "1px solid rgba(245,158,11,.4)", borderRadius: 16, padding: "32px 30px", boxShadow: "0 30px 90px -20px rgba(0,0,0,.85)", fontFamily: "'Space Mono', monospace" }}>
        <div style={{ fontSize: 38, marginBottom: 8 }}>📩</div>
        <div style={{ fontSize: 25, fontWeight: 800, color: "#f8fafc", fontFamily: "'Playfair Display', serif", marginBottom: 8 }}>{t("gateTitle")}</div>
        <div style={{ fontSize: 15, color: "#94a3b8", lineHeight: 1.65, marginBottom: 20 }}>{t("gateBody")}</div>
        <input type="email" value={email} autoFocus
          onChange={e => { setEmail(e.target.value); setErr(""); }}
          onKeyDown={e => e.key === "Enter" && !busy && submit()}
          placeholder="you@company.com"
          style={{ ...INL, marginBottom: 6 }} />
        {err && <div style={{ color: "#f87171", fontSize: 14, marginTop: 4 }}>{err}</div>}
        <button onClick={submit} disabled={busy}
          style={{ ...EXP, width: "100%", padding: "14px", fontSize: 17, marginTop: 14, opacity: busy ? .6 : 1, cursor: busy ? "wait" : "pointer" }}>
          {busy ? t("gateSaving") : t("gateBtn")}
        </button>
        <div style={{ fontSize: 12, color: "#64748b", textAlign: "center", marginTop: 14, lineHeight: 1.5 }}>🔒</div>
      </div>
    </div>
  );
}
/* end of drop-in block — integration steps are in the chat message / README */

/* ============================================================================
   STEEL OPTIMIZER — bilingual PART 1b : dictionary addendum
   Paste this right AFTER the LANG_DICT object from Part 1. It folds the rest of
   the strings (modules, results, exports) into the same dictionary.
   Status glyphs (✓ / ⚠ / ♻ / ⚙) are kept at the START of message strings on
   purpose — code elsewhere checks msg.startsWith("✓") etc., and both languages
   must keep that prefix.
============================================================================ */
Object.assign(LANG_DICT, {
  // ── plates: offcut card ──
  nameThisPlate:  { en: "Name this plate",             ar: "سمِّ هذا اللوح" },
  removeX:        { en: "✕ remove",                    ar: "✕ إزالة" },
  whatShape:      { en: "What shape is the leftover?", ar: "ما شكل القطعة المتبقية؟" },
  shapeRect:      { en: "▭ Rectangle",                 ar: "▭ مستطيل" },
  shapeL:         { en: "⌐ L-shape (has a cut-out corner)", ar: "⌐ شكل L (بزاوية مقصوصة)" },
  willNestInto:   { en: "✓ We'll nest parts into a",   ar: "✓ سنوزّع القطع داخل مساحة" },
  biggestRectL:   { en: " (the biggest rectangle that fits the L)", ar: " (أكبر مستطيل يدخل في شكل L)" },
  pieceWord:      { en: "piece",                       ar: "قطعة" },
  piecesWord:     { en: "pieces",                      ar: "قطع" },
  addLeftMore:    { en: "+ Add a leftover from another job / site", ar: "+ إضافة قطعة متبقية من مشروع/موقع آخر" },

  // ── plates: manual + select ──
  notInStock:     { en: "(not in stock)",              ar: "(غير متوفر بالمخزون)" },

  // ── plates: upload screen ──
  upBannerP:      { en: "💡 Thickness comes from the file; the matching stock sheet size is taken from your list above. Unknown thickness falls back to 1220×2440.",
                    ar: "💡 تؤخذ السماكة من الملف، ويُؤخذ مقاس اللوح المطابق من قائمتك أعلاه. عند جهل السماكة يُستخدم 1220×2440." },
  upHintP:        { en: "Reads Profile column (PLT/FLT/PL → thickness × width). Hot-rolled ignored.",
                    ar: "يقرأ عمود المقطع (PLT/FLT/PL → سماكة × عرض). تُتجاهل المقاطع المدرفلة على الساخن." },
  mapColumns:     { en: "Map columns:",                ar: "اربط الأعمدة:" },
  applyMapping:   { en: "Apply Mapping",               ar: "تطبيق الربط" },
  partsReady:     { en: "PARTS READY",                 ar: "القطع جاهزة" },
  moreWord:       { en: "more",                        ar: "أخرى" },

  // ── plates: excel messages (keep leading glyph) ──
  msgImportedP:   { en: "✓ Imported {n} plate part type{s}. Hot-rolled sections & subtotal rows ignored.",
                    ar: "✓ تم استيراد {n} نوع قطعة لوح{s}. تم تجاهل المقاطع المدرفلة وصفوف المجاميع." },
  msgMapCols:     { en: "⚠ Couldn't auto-detect columns. Map them below.",
                    ar: "⚠ تعذّر اكتشاف الأعمدة تلقائيًا. اربطها بالأسفل." },
  msgReadFail:    { en: "⚠ Could not read this file.", ar: "⚠ تعذّرت قراءة هذا الملف." },
  msgImportedP2:  { en: "✓ Imported {n} plate parts",  ar: "✓ تم استيراد {n} قطعة لوح" },
  msgExtractFail: { en: "⚠ Still couldn't extract.",   ar: "⚠ ما زال التعذّر في الاستخراج قائمًا." },

  // ── leftover importer messages ──
  liReading:      { en: "Reading {name}…",             ar: "جارٍ قراءة {name}…" },
  liWrongType:    { en: "⚠ That looks like a {type} leftover file. Use it in the {mod} module.",
                    ar: "⚠ يبدو أن هذا ملف بواقٍ من نوع {type}. استخدمه في قسم {mod}." },
  liImported:     { en: "✓ Imported {n} leftover type{s} ({m} piece{ms}). They'll be cut from first.",
                    ar: "✓ تم استيراد {n} نوع بواقٍ{s} ({m} قطعة{ms}). سيتم القص منها أولًا." },
  liErr:          { en: "⚠ {msg}",                     ar: "⚠ {msg}" },
  modPlatesName:  { en: "Plates",                      ar: "الألواح" },
  modSectionsName:{ en: "Sections",                    ar: "المقاطع" },

  // ── plate results: summary + banners ──
  sheetUnitOf:    { en: "{n} × {sw} × {sh} mm ({mat})", ar: "{n} × {sw} × {sh} مم ({mat})" },
  acrossThk:      { en: "TOTAL: {n} sheet{s} across {g} thickness{es}", ar: "الإجمالي: {n} لوح{s} عبر {g} سماكة{es}" },
  spliceBannerHd: { en: "⚙ {n} part{s} larger than the sheet — cut & welded from multiple pieces",
                    ar: "⚙ {n} قطعة{s} أكبر من اللوح — تُقص وتُلحم من عدة أجزاء" },
  spliceBannerBd: { en: "Each is split across sheets and joined with a welded seam (no bigger sheet needed). Confirm a welded seam is acceptable for these pieces before fabrication:",
                    ar: "تُقسَّم كل قطعة على عدة ألواح وتُوصل بلحام (دون الحاجة للوح أكبر). تأكّد من قبول وصلة اللحام لهذه القطع قبل التصنيع:" },
  reusedBannerP:  { en: "♻ {n} reused offcut{s} from previous projects used first. The sheet counts below are the minimum new standard sheets you still need to buy.",
                    ar: "♻ تم استخدام {n} قطعة متبقية{s} من مشاريع سابقة أولًا. أعداد الألواح بالأسفل هي أقل عدد ألواح جديدة قياسية تحتاج لشرائها." },

  // ── plate results: nesting + offcut register ──
  nestLayoutTtl:  { en: "🗂 Nesting Layout — {thk} mm Plate ({n} sheet{s})", ar: "🗂 مخطط التوزيع — لوح {thk} مم ({n} لوح{s})" },
  offcutRegTtl:   { en: "♻ Reusable Offcut Register",  ar: "♻ سجل القصاصات القابلة لإعادة الاستخدام" },
  offcutRegLead:  { en: "{n} reusable leftover{s} ≈ {t} t to keep for future jobs (≥ {min} mm/side). For L-shaped leftovers, USABLE is the largest rectangle you can cut from it.",
                    ar: "{n} قطعة متبقية{s} ≈ {t} طن للاحتفاظ بها لمشاريع قادمة (≥ {min} مم/ضلع). للقطع على شكل L، \"القابل للاستخدام\" هو أكبر مستطيل يمكن قصّه منها." },
  thQty:          { en: "QTY",                         ar: "العدد" },
  thThickness:    { en: "THICKNESS",                   ar: "السماكة" },
  thFromSheet:    { en: "FROM SHEET",                  ar: "من اللوح" },
  thShape:        { en: "SHAPE",                       ar: "الشكل" },
  thOverall:      { en: "OVERALL",                     ar: "الأبعاد الكلية" },
  thUsableRect:   { en: "USABLE RECT",                 ar: "المستطيل القابل للاستخدام" },
  thWeight:       { en: "WEIGHT",                      ar: "الوزن" },
  thStatus:       { en: "STATUS",                      ar: "الحالة" },
  shapeLshort:    { en: "⌐ L-shape",                   ar: "⌐ شكل L" },
  shapeRectShort: { en: "▭ Rect",                      ar: "▭ مستطيل" },
  keepStatus:     { en: "✓ keep",                      ar: "✓ احتفظ" },
  fromSheetN:     { en: "Sheet {n}",                   ar: "لوح {n}" },

  // ── downloads card ──
  dlTwoDeliv:     { en: "Two deliverables: the procurement files (what to buy + cutting plan) and the leftover files (reusable offcuts to keep).",
                    ar: "مخرجان: ملفات التوريد (ما يُشترى + خطة القص) وملفات البواقي (القصاصات القابلة لإعادة الاستخدام للاحتفاظ بها)." },

  // ── plate canvas mini labels ──
  miniSheet:      { en: "Sheet",                       ar: "اللوح" },
  miniThickness:  { en: "Thickness",                   ar: "السماكة" },
  miniParts:      { en: "Parts",                       ar: "القطع" },
  miniUtil:       { en: "Utilization",                 ar: "الاستغلال" },
  reusableLeftover:{ en: "♻ Reusable leftover",        ar: "♻ قطعة قابلة لإعادة الاستخدام" },

  // ── procurement block ──
  pbFinalReady:   { en: "Final Decision-Ready Output", ar: "مخرجات جاهزة لاتخاذ القرار" },
  pbProcSummary:  { en: "Material Procurement Summary",ar: "ملخص توريد المواد" },
  thSheetSize:    { en: "SHEET SIZE",                  ar: "مقاس اللوح" },
  thSheetsReq:    { en: "SHEETS REQ.",                 ar: "الألواح المطلوبة" },
  thPurchaseWt:   { en: "PURCHASE WT (full sheets)",   ar: "وزن الشراء (ألواح كاملة)" },
  thParts:        { en: "PARTS",                       ar: "القطع" },
  thNetWt:        { en: "NET PARTS WT",                ar: "الوزن الصافي للقطع" },
  thScrapWt:      { en: "SCRAP WT",                    ar: "وزن السكراب" },
  pbLegend:       { en: "● Total Weight = full sheets you purchase · ● Net Weight = steel used in the project · ● Scrap = offcut (Total − Net)",
                    ar: "● الوزن الإجمالي = الألواح الكاملة التي تشتريها · ● الوزن الصافي = الصلب المستخدم في المشروع · ● السكراب = القصاصة (الإجمالي − الصافي)" },
  pbTotalPurchase:{ en: "Total Purchase Weight",       ar: "إجمالي وزن الشراء" },
  pbGrossNote:    { en: "gross — all {n} full sheets", ar: "إجمالي — كل {n} لوح كامل" },
  pbNetNote:      { en: "steel that ends up in the structure", ar: "الصلب الذي يدخل في المنشأ" },
  pbScrapNote:    { en: "{p}% of purchase",            ar: "{p}% من الشراء" },
  pbGenerated:    { en: "Generated by Steel Cut & Nest Optimizer", ar: "أُنشئ بواسطة برنامج توفيق القطعيات وتحسين قص الحديد" },
  thicknessesWord:{ en: "thickness{es}",               ar: "سماكة{es}" },

  // ── sections module ──
  sectionsWS:     { en: "Steel Sections — Workspace",  ar: "المقاطع الحديدية — مساحة العمل" },
  manualCutList:  { en: "✏️ Manual Cutting List",       ar: "✏️ قائمة قص يدوية" },
  thLengths:      { en: "Lengths (mm)",                ar: "الأطوال (مم)" },
  thMarketLen:    { en: "Available length in market (mm)", ar: "الطول المتوفر بالسوق (مم)" },
  addProfile:     { en: "+ ADD PROFILE",               ar: "+ إضافة مقطع" },
  piecesCount:    { en: "{n} pieces",                  ar: "{n} قطعة" },
  lengthsHint:    { en: "e.g. 3000 4500 6000x3  →  one 3 m, one 4.5 m, three 6 m",
                    ar: "مثال: 3000 4500 6000x3 ← واحد 3م، واحد 4.5م، ثلاثة 6م" },
  blank12m:       { en: "blank = 12m",                 ar: "فارغ = 12م" },
  secManualHelp:  { en: "Type IPE, HEA, HEB, UB, UC, JIS HW/HM/HN, RHS, SHS, CHS, cold-formed C/Z, GOST & GB. Built-up girder: type 900x400x20x15. Lengths: separate each by a space, e.g. 3000 4500 6000x3 — write 6000x3 for three 6 m pieces.",
                    ar: "اكتب IPE، HEA، HEB، UB، UC، JIS HW/HM/HN، RHS، SHS، CHS، مقاطع C/Z المشكّلة على البارد، GOST و GB. كمرة مركّبة: اكتب 900x400x20x15. الأطوال: افصل بينها بمسافة، مثل 3000 4500 6000x3 — اكتب 6000x3 لثلاث قطع 6م." },
  computedSuffix: { en: "(computed)",                  ar: "(محسوب)" },

  // ── sections upload ──
  upTitleS:       { en: "Upload your material list",   ar: "ارفع قائمة المواد" },
  upBodyS:        { en: "Drop a Tekla material list — or any material list that has steel lengths and quantities.",
                    ar: "أفلت قائمة مواد Tekla — أو أي قائمة مواد تحتوي أطوال وكميات الصلب." },
  upAcceptS:      { en: "Accepts .xlsx · .xls · .csv · .txt — or click to browse",
                    ar: "يقبل ‎.xlsx · .xls · .csv · .txt — أو انقر للتصفّح" },
  upHintS:        { en: "Sections are detected automatically · plate rows ignored (use the Plates module for those) · stock length defaults to 12 m and is editable.",
                    ar: "تُكتشف المقاطع تلقائيًا · تُتجاهل صفوف الألواح (استخدم قسم الألواح لها) · الطول القياسي الافتراضي 12م وقابل للتعديل." },
  secUpErr:       { en: "No structural sections detected. If this file is only plates, use the Plates module.",
                    ar: "لم تُكتشف مقاطع إنشائية. إذا كان الملف ألواحًا فقط، استخدم قسم الألواح." },
  secReadErr:     { en: "Failed to read: {msg}",       ar: "تعذّرت القراءة: {msg}" },
  secNoValid:     { en: "No valid cuts.",              ar: "لا توجد قطوعات صالحة." },
  secEnterLen:    { en: "Enter at least one length.",  ar: "أدخل طولًا واحدًا على الأقل." },
  uploadedOk:     { en: "✓ {name} — {rows} sections, {pieces} pieces",
                    ar: "✓ {name} — {rows} مقطع، {pieces} قطعة" },
  platesIgnored:  { en: " · {n} plate row{s} ignored", ar: " · تم تجاهل {n} صف ألواح{s}" },
  thKgM:          { en: "kg/m",                        ar: "كجم/م" },
  andMore:        { en: "…and {n} more",               ar: "…و{n} أخرى" },

  // ── sections leftovers ──
  loLeftoverLen:  { en: "Leftover length (mm)",        ar: "طول القطعة المتبقية (مم)" },
  loBarLeft:      { en: "length of the bar you have left", ar: "طول العود المتبقي لديك" },
  loHowManyLike:  { en: "how many like this",          ar: "كم عدد المماثل" },
  loSameProfile:  { en: "A leftover only feeds parts of the same profile & grade — leave grade blank to match parts with no grade. We cut the tightest fit from each leftover first, then buy the minimum new bars for what's left.",
                    ar: "القطعة المتبقية تُغذّي فقط قطعًا من نفس المقطع والرتبة — اترك الرتبة فارغة لمطابقة القطع بلا رتبة. نقصّ الأنسب من كل قطعة متبقية أولًا، ثم نشتري أقل عدد أعواد جديدة للمتبقي." },

  // ── section results ──
  whatToOrderTtl: { en: "⬡ WHAT TO ORDER",             ar: "⬡ ما يجب طلبه" },
  youNeedWord:    { en: "YOU NEED",                    ar: "تحتاج" },
  gradeWord:      { en: "grade",                       ar: "رتبة" },
  totalAcrossProf:{ en: "TOTAL: {n} bars across {g} profile{s}", ar: "الإجمالي: {n} عود عبر {g} مقطع{s}" },
  leftoversUsedHd:{ en: "♻ LEFTOVERS USED FIRST",      ar: "♻ استُخدمت البواقي أولًا" },
  leftoversUsedBd:{ en: "{n} of your leftover bar{s} cut before buying new", ar: "تم قص {n} من أعوادك المتبقية{s} قبل شراء جديد" },
  steelNotBought: { en: "STEEL NOT PURCHASED",         ar: "صلب لم يُشترَ" },
  longMembersHd:  { en: "⚙ LONG MEMBERS — SPLICED FROM {len} BARS", ar: "⚙ أعضاء طويلة — موصولة من أعواد {len}" },
  eachFromBars:   { en: " — each from {n} × {len} bars", ar: " — كلٌّ منها من {n} × عود {len}" },
  cutFromLeftTtl: { en: "♻ Cut from your leftovers — no purchase", ar: "♻ مقصوص من بواقيك — بلا شراء" },
  cutFromLeftLead:{ en: "These pieces came out of leftover bars you already owned — they are not in the buy list above.",
                    ar: "هذه القطع خرجت من أعواد متبقية تملكها — وهي ليست ضمن قائمة الشراء أعلاه." },
  leftRemaining:  { en: " Leftover still remaining after this: {kg}.", ar: " المتبقي بعد ذلك: {kg}." },
  visualCutPlan:  { en: "Visual Cut Plan",             ar: "خطة القص المرئية" },
  noPlan:         { en: "No plan.",                    ar: "لا توجد خطة." },
  allFromLeft:    { en: "All pieces were cut from your leftovers — nothing new to buy.",
                    ar: "تم قص كل القطع من بواقيك — لا شيء جديد للشراء." },
  reusableBarTtl: { en: "♻ Reusable Bar Offcuts — leftover steel for the next project",
                    ar: "♻ قصاصات الأعواد القابلة لإعادة الاستخدام — صلب متبقٍ للمشروع القادم" },
  reusableBarLead:{ en: "{n} leftover bar end{s} ≥ 1 m ≈ {kg} — keep these to cut from on your next job instead of buying new.",
                    ar: "{n} نهاية عود متبقية{s} ≥ 1م ≈ {kg} — احتفظ بها للقص منها في مشروعك القادم بدل الشراء." },
  thFromBar:      { en: "FROM BAR",                    ar: "من العود" },
  thStockLength:  { en: "STOCK LENGTH",                ar: "الطول القياسي" },
  thReusableLeft: { en: "REUSABLE LEFTOVER",           ar: "البواقي القابلة لإعادة الاستخدام" },
  barNum:         { en: "Bar {n}",                     ar: "عود {n}" },
  leftoverNum:    { en: "Leftover {n}",                ar: "بواقٍ {n}" },

  // ── section per-group stat labels ──
  stBarsBuy:      { en: "Bars to Buy",                 ar: "أعواد للشراء" },
  stTotalLen:     { en: "Total Length",               ar: "الطول الإجمالي" },
  stOffcut:       { en: "Offcut",                      ar: "قصاصة" },
  stWaste:        { en: "Waste",                       ar: "الهدر" },
  stUtil:         { en: "Utilization",                ar: "الاستغلال" },
  stWeight:       { en: "Weight",                      ar: "الوزن" },

  // ── cut bar ──
  cbBar:          { en: "BAR #{n}",                    ar: "عود رقم {n}" },
  cbAvailable:    { en: "available {len}",             ar: "متوفر {len}" },
  cbUsed:         { en: "{p}% used",                   ar: "{p}% مُستخدم" },
});

// ── canvas micro-labels + a few composed strings ──
Object.assign(LANG_DICT, {
  canvLoffcut:    { en: "♻ L-OFFCUT",                  ar: "♻ قصاصة L" },
  canvOffcut:     { en: "♻ OFFCUT",                    ar: "♻ قصاصة" },
  canvUse:        { en: "use {w}×{h}",                 ar: "استخدم {w}×{h}" },
  weldLbl:        { en: "WELD {w}×{h}",                ar: "لحام {w}×{h}" },
  sheetTab:       { en: "Sheet {n} · {p}%",            ar: "لوح {n} · {p}%" },
  reuseListLead:  { en: "♻ Reusable leftovers: {list}", ar: "♻ قصاصات قابلة لإعادة الاستخدام: {list}" },
  lShapeUse:      { en: "L-shape, use ",               ar: "شكل L، استخدم " },
});
Object.assign(LANG_DICT, { notchWord: { en: "notch", ar: "ركن" } });
Object.assign(LANG_DICT, {
  liveNesting: { en: "● LIVE NESTING — REAL CASE", ar: "● توزيع حيّ — حالة حقيقية" },
  scrapReuse:  { en: "scrap/reuse",                 ar: "هدر/إعادة استخدام" },
});
Object.assign(LANG_DICT, {
  usableWord:  { en: "USABLE",        ar: "قابل للاستخدام" },
  widthArrow:  { en: "↔ width",       ar: "↔ العرض" },
  lengthArrow: { en: "↕ length",      ar: "↕ الطول" },
  cutoutWord:  { en: "cut-out",       ar: "ركن مقصوص" },
  lThickNote:  { en: "{t} mm thick · green = the rectangle we nest into", ar: "السماكة {t} مم · الأخضر = المستطيل الذي نوزّع عليه" },
  lpIntro:     { en: "Stand at one corner of your leftover and call it 0, 0. Walk the outline and type each corner you reach — the drawing builds as you go.", ar: "قف عند أحد أركان القطعة المتبقية واعتبره 0، 0. سِر على المحيط واكتب كل ركن تصل إليه — يُرسم الشكل تلقائيًا." },
  lpColX:      { en: "X · width →",   ar: "X · العرض ←" },
  lpColY:      { en: "Y · length ↑",  ar: "Y · الطول ↑" },
  lpAxisX:     { en: "X · width",     ar: "X · العرض" },
  lpAxisY:     { en: "Y · length",    ar: "Y · الطول" },
  lpReset:     { en: "↺ Reset shape", ar: "↺ إعادة ضبط الشكل" },
  lpOverall:   { en: "Overall",       ar: "الإجمالي" },
  lpUsable:    { en: "usable",        ar: "القابل للاستخدام" },
  lpTip:       { en: "tip: an L has 6 corners forming a clean step", ar: "ملاحظة: شكل L له 6 أركان تُكوّن درجة واضحة" },
  lpHint0:     { en: "start corner — leave at 0, 0", ar: "ركن البداية — اتركه عند 0، 0" },
  lpHint1:     { en: "→ go along the bottom edge",   ar: "← سِر على الحافة السفلية" },
  lpHint2:     { en: "↑ up the side",                ar: "↑ صعودًا على الجانب" },
  lpHint3:     { en: "← into the notch",             ar: "→ إلى داخل القَطع" },
  lpHint4:     { en: "↑ up to the top",              ar: "↑ إلى الأعلى" },
  lpHint5:     { en: "← back toward start",          ar: "→ عودة نحو البداية" },
});
Object.assign(LANG_DICT, {
  priceCardTtl:   { en: "💰 Steel Price & Cost (optional)",  ar: "💰 سعر الحديد والتكلفة (اختياري)" },
  priceCardLead:  { en: "Enter your steel price per ton — costs will appear in the results and in the PDF/Excel reports. Leave it empty to skip pricing.",
                    ar: "أدخل سعر الطن لديك — ستظهر التكاليف في النتائج وفي تقارير PDF وExcel. اتركه فارغًا لتجاوز التسعير." },
  priceCurrency:  { en: "Currency",                          ar: "العملة" },
  pricePerTonLbl: { en: "Price per ton",                     ar: "سعر الطن" },
  costPurchase:   { en: "TOTAL PURCHASE COST",               ar: "تكلفة الشراء الإجمالية" },
  costNet:        { en: "NET USED VALUE",                    ar: "قيمة الحديد المستخدم" },
  costScrap:      { en: "SCRAP COST (WASTE)",                ar: "تكلفة الهالك" },
  costOffcut:     { en: "REUSABLE OFFCUT VALUE",             ar: "قيمة القصاصات الصالحة" },
  costBasedOn:    { en: "Cost basis: {p} per ton (your price).",
                    ar: "أساس التكلفة: {p} للطن (سعرك)." },
});

/* ============================================================================
   STEEL PRICE & COST LAYER (optional, additive — engine untouched)
   ----------------------------------------------------------------------------
   One price source only: the user's own price per ton, entered in Step 1.
   Empty = pricing off. The choice persists in localStorage. Plates module
   only — sections are intentionally unpriced.
============================================================================ */
const PRICE_KEY = "steelopt_price_v1";

function readSavedPricing() {
  try {
    const v = JSON.parse(localStorage.getItem(PRICE_KEY));
    if (v && typeof v === "object") {
      const price = isFinite(+v.price) && +v.price > 0 ? +v.price : (isFinite(+v.custom) && +v.custom > 0 ? +v.custom : 0); // v.custom = legacy shape
      return { price, currency: v.currency === "USD" ? "USD" : "SAR" };
    }
  } catch { /* private mode */ }
  return { price: 0, currency: "SAR" };
}

function usePricing() {
  const [cfg, setCfg] = useState(readSavedPricing);
  const save = useCallback(p => { setCfg(c => { const n = { ...c, ...p }; try { localStorage.setItem(PRICE_KEY, JSON.stringify(n)); } catch { /* ignore */ } return n; }); }, []);
  return { ...cfg, save, pricePerTon: cfg.price > 0 ? cfg.price : 0, active: cfg.price > 0, srcNote: "your entered price" };
}

/* Compact price picker card — sits between stock settings and leftover reuse */
function PriceCard({ pricing }) {
  const { t } = useLang();
  return (
    <Card title={t("priceCardTtl")} style={{ marginTop: 20 }}>
      <div style={{ fontSize: 15, color: "#94a3b8", marginBottom: 14, lineHeight: 1.6 }}>{t("priceCardLead")}</div>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end" }}>
        <div><Label>{t("pricePerTonLbl")}</Label><input type="number" min="0" value={pricing.price || ""} onChange={e => pricing.save({ price: Math.max(0, +e.target.value || 0) })} placeholder="2900" style={{ ...IN, width: 140 }} /></div>
        <div><Label>{t("priceCurrency")}</Label><select value={pricing.currency} onChange={e => pricing.save({ currency: e.target.value })} style={SEL}><option value="SAR">SAR ر.س</option><option value="USD">USD $</option></select></div>
        {pricing.active && <div style={{ paddingBottom: 6, fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#6ee7b7" }}>≈ <b style={{ color: "#34d399", fontSize: 18 }}>{Math.round(pricing.pricePerTon).toLocaleString()}</b> {pricing.currency}/{t("t_ton")}</div>}
      </div>
    </Card>
  );
}

/* Cost summary strip — appears in results only when a price is active */
function CostBlock({ buyKg, netKg, wasteKg, offcutKg, pricing }) {
  const { t } = useLang();
  if (!pricing || !pricing.active) return null;
  const P = pricing.pricePerTon, C = pricing.currency;
  const M = v => `${Math.round(v).toLocaleString()} ${C}`;
  const cards = [
    { l: t("costPurchase"), v: M(buyKg / 1000 * P), i: "💰", a: "#f59e0b" },
    { l: t("costNet"), v: M(netKg / 1000 * P), i: "⚖", a: "#cbd5e1" },
    { l: t("costScrap"), v: M(wasteKg / 1000 * P), i: "🗑", a: "#EF4444" },
    { l: t("costOffcut"), v: M(offcutKg / 1000 * P), i: "♻", a: "#10B981" },
  ];
  return (
    <div style={{ marginBottom: 24 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(175px,1fr))", gap: 12 }}>{cards.map((c, ci) => <StatCard key={ci} {...c} />)}</div>
      <div style={{ marginTop: 8, fontSize: 13, color: "#64748b", fontFamily: "'Space Mono', monospace", textAlign: "end" }}>{t("costBasedOn", { p: `${Math.round(P).toLocaleString()} ${C}` })}</div>
    </div>
  );
}

/* ============================================================================
   STEEL OPTIMIZER — bilingual PART 2 : Plates vertical + shared UI (wired)
   ----------------------------------------------------------------------------
   Paste these OVER your existing CutBar, PlateCanvas, LeftoverImporter,
   PlatesModule, BigTon, PlateResults, ProcurementBlock.
   Requires PART 1 (LangProvider/useT) + PART 1b (dictionary addendum) pasted first.
   Your engines, STEEL_DB, helpers, Mini, StatCard, LShapeDiagram, LPointBuilder
   and the export functions stay exactly as they are.
   NOTE: PlateResults' download buttons now call the export fns WITH a trailing
   `t` argument, e.g. exportPlatePDF(results, material, reuseMin, t). When you wire
   Part 3 (exports), have each export fn accept that final `t` param.
   The Sections vertical (SectionsModule, SectionResults) is delivered separately.
============================================================================ */

/* ─── SECTION CUT BAR (wired) ────────────────────────────────────────────── */
function CutBar({ bin, index }) {
  const t = useT();
  const used = (((bin.stockLength - bin.remaining) / bin.stockLength) * 100).toFixed(1);
  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
        <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 14, color: "#cbd5e1", fontWeight: 700 }}>{t("cbBar", { n: String(index + 1).padStart(2, "0") })} <span style={{ color: "#64748b" }}>· {t("cbAvailable", { len: fmtMm(bin.stockLength) })}</span></span>
        <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 14, color: "#f59e0b", fontWeight: 700 }}>{t("cbUsed", { p: used })}</span>
      </div>
      <div style={{ display: "flex", height: 40, borderRadius: 6, overflow: "hidden", border: "1px solid #2d3748", background: "#0f1318" }}>
        {bin.cuts.map((cut, ci) => { const w = (cut.length / bin.stockLength) * 100; const col = PART_COLORS[ci % PART_COLORS.length]; return <div key={ci} title={`${cut.label ? soCutLabel(cut.label, t) + ": " : ""}${fmtMm(cut.length)}`} style={{ width: `${w}%`, background: `linear-gradient(180deg, ${col}, ${col}cc)`, borderRight: "1px solid rgba(0,0,0,.45)", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", overflow: "hidden", textShadow: "0 1px 2px rgba(0,0,0,.5)" }}>{w > 9 && <><span style={{ fontSize: 14, fontFamily: "'Space Mono', monospace", color: "#fff", fontWeight: 800, lineHeight: 1.1, whiteSpace: "nowrap" }}>{fmtMm(cut.length)}</span>{cut.label && w > 16 && <span style={{ fontSize: 11, fontFamily: "'Space Mono', monospace", color: "rgba(255,255,255,.85)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "95%" }}>{soCutLabel(cut.label, t)}</span>}</>}</div>; })}
        {bin.remaining > 0 && <div style={{ width: `${(bin.remaining / bin.stockLength) * 100}%`, background: "repeating-linear-gradient(45deg, #16202c, #16202c 5px, #1b2733 5px, #1b2733 10px)", borderLeft: "2px dashed #475569", display: "flex", alignItems: "center", justifyContent: "center" }}>{(bin.remaining / bin.stockLength) > 0.05 && <span style={{ fontSize: 13, fontFamily: "'Space Mono', monospace", color: bin.remaining >= 1000 ? "#34d399" : "#64748b", fontWeight: 700, whiteSpace: "nowrap" }}>{fmtMm(bin.remaining)}{bin.remaining >= 1000 ? " ♻" : ""}</span>}</div>}
      </div>
    </div>
  );
}

/* ─── PLATE NESTING CANVAS (wired) ───────────────────────────────────────── */
function PlateCanvas({ sheets, sheetW, sheetH, colorMap, thickness }) {
  const t = useT();
  const [active, setActive] = useState(0); const [zoom, setZoom] = useState(1); const cv = useRef(null);
  useEffect(() => { setActive(0); }, [sheets]);
  const sObj = sheets[active] || { placements: [], offcuts: [] };
  const sheet = sObj.placements || [], offcuts = sObj.offcuts || [];
  const SCALE = Math.min(520 / sheetW, 360 / sheetH) * zoom;
  useEffect(() => {
    const c = cv.current; if (!c) return; const ctx = c.getContext("2d");
    const W = sheetW * SCALE, H = sheetH * SCALE; c.width = W + 44; c.height = H + 44;
    ctx.fillStyle = "#0f1318"; ctx.fillRect(0, 0, c.width, c.height); ctx.save(); ctx.translate(22, 22);
    ctx.fillStyle = "#1b2530"; ctx.fillRect(0, 0, W, H); ctx.strokeStyle = "#f59e0b88"; ctx.lineWidth = 1.5; ctx.strokeRect(0, 0, W, H);
    ctx.strokeStyle = "rgba(120,130,140,.2)"; ctx.lineWidth = .5; const step = Math.max(20, Math.round(100 * SCALE));
    for (let x = step; x < W; x += step) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, H); ctx.stroke(); }
    for (let y = step; y < H; y += step) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(W, y); ctx.stroke(); }
    sheet.forEach((p, i) => { const px = p.x * SCALE, py = p.y * SCALE, pw = p.w * SCALE, ph = p.h * SCALE; const color = colorMap[p.id] || PART_COLORS[i % PART_COLORS.length]; ctx.fillStyle = color + "55"; ctx.fillRect(px, py, pw, ph); ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.strokeRect(px, py, pw, ph);
      if (p.spliced) { ctx.save(); ctx.beginPath(); ctx.rect(px, py, pw, ph); ctx.clip(); ctx.strokeStyle = "rgba(255,255,255,.35)"; ctx.lineWidth = 1; for (let d = -ph; d < pw; d += 8) { ctx.beginPath(); ctx.moveTo(px + d, py); ctx.lineTo(px + d + ph, py + ph); ctx.stroke(); } ctx.restore(); ctx.strokeStyle = "#fbbf24"; ctx.lineWidth = 2; ctx.setLineDash([5, 3]); ctx.strokeRect(px, py, pw, ph); ctx.setLineDash([]); }
      const fs = Math.max(8, Math.min(11, pw / 6, ph / 3)); ctx.fillStyle = "rgba(255,255,255,.95)"; ctx.font = `${fs}px 'Space Mono', monospace`; ctx.textAlign = "center"; ctx.textBaseline = "middle"; if (pw > 22 && ph > 13) { ctx.fillText(p.rotated ? `${p.label}↺` : p.label, px + pw / 2, py + ph / 2 - (ph > 26 ? 6 : 0)); if (ph > 26) { ctx.font = `${Math.max(7, fs - 1)}px 'Space Mono', monospace`; ctx.fillStyle = p.spliced ? "#fcd34d" : "rgba(230,210,170,.8)"; ctx.fillText(p.spliced ? t("weldLbl", { w: Math.round(p.w), h: Math.round(p.h) }) : `${Math.round(p.w)}×${Math.round(p.h)}`, px + pw / 2, py + ph / 2 + 7); } } });
    offcuts.forEach(o => {
      ctx.save();
      if (o.shape === "L") {
        const ox = o.x * SCALE, oy = o.y * SCALE, A = o.A * SCALE, B = o.B * SCALE, nW = o.notchW * SCALE, nH = o.notchH * SCALE;
        ctx.beginPath(); ctx.moveTo(ox + nW, oy); ctx.lineTo(ox + A, oy); ctx.lineTo(ox + A, oy + B); ctx.lineTo(ox, oy + B); ctx.lineTo(ox, oy + nH); ctx.lineTo(ox + nW, oy + nH); ctx.closePath();
        ctx.fillStyle = "rgba(16,185,129,.18)"; ctx.fill(); ctx.strokeStyle = "#10b981"; ctx.lineWidth = 1.4; ctx.stroke();
      } else {
        const ox = o.x * SCALE, oy = o.y * SCALE, A = o.A * SCALE, B = o.B * SCALE;
        ctx.fillStyle = "rgba(16,185,129,.18)"; ctx.fillRect(ox, oy, A, B); ctx.strokeStyle = "#10b981"; ctx.lineWidth = 1.4; ctx.strokeRect(ox, oy, A, B);
      }
      ctx.restore();
      if (o.A * SCALE > 30 && o.B * SCALE > 16) { ctx.fillStyle = "#6ee7b7"; ctx.font = `${Math.max(8, Math.min(10, o.A * SCALE / 10))}px 'Space Mono', monospace`; ctx.textAlign = "center"; ctx.textBaseline = "middle"; const cx = (o.x + (o.shape === "L" ? o.A : o.A) / 2) * SCALE; const cy = (o.y + o.B / 2) * SCALE; ctx.fillText(o.shape === "L" ? t("canvLoffcut") : t("canvOffcut"), cx, cy - 5); ctx.font = `${Math.max(7, Math.min(9, o.A * SCALE / 12))}px 'Space Mono', monospace`; ctx.fillText(t("canvUse", { w: o.w, h: o.h }), cx, cy + 6); }
    });
    ctx.fillStyle = "rgba(245,200,120,.8)"; ctx.font = "10px 'Space Mono', monospace"; ctx.textAlign = "center"; ctx.fillText(`${sheetW} ${t("mm")}`, W / 2, H + 15);
    ctx.save(); ctx.translate(-9, H / 2); ctx.rotate(-Math.PI / 2); ctx.textAlign = "center"; ctx.fillText(`${sheetH} ${t("mm")}`, 0, 0); ctx.restore(); ctx.restore();
  }, [sheet, offcuts, sheetW, sheetH, SCALE, colorMap, t]);
  const used = sheet.reduce((s, p) => s + p.w * p.h, 0), total = sheetW * sheetH, util = Math.round((used / total) * 100);
  return (
    <div style={{ fontFamily: "'Space Mono', monospace" }}>
      <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
        {sheets.map((sh, i) => { const pct = Math.round(((sh.placements || []).reduce((s, p) => s + p.w * p.h, 0) / total) * 100); return <button key={i} onClick={() => setActive(i)} style={{ padding: "5px 12px", borderRadius: 4, border: "1px solid", borderColor: active === i ? "#f59e0b" : "#2d3748", background: active === i ? "rgba(245,158,11,.15)" : "rgba(15,19,24,.8)", color: active === i ? "#f59e0b" : "#94a3b8", fontSize: 13, cursor: "pointer" }}>{t("sheetTab", { n: i + 1, p: pct })}</button>; })}
        <div style={{ marginLeft: "auto", display: "flex", gap: 6 }}><button onClick={() => setZoom(z => Math.min(z + .2, 3))} style={ZB}>+</button><button onClick={() => setZoom(z => Math.max(z - .2, .4))} style={ZB}>−</button></div>
      </div>
      <div style={{ overflow: "auto", borderRadius: 6, border: "1px solid #2d3748", maxHeight: 440 }}><canvas ref={cv} style={{ display: "block" }} /></div>
      {offcuts.length > 0 && <div style={{ marginTop: 8, padding: "8px 12px", background: "rgba(16,185,129,.08)", border: "1px solid rgba(16,185,129,.3)", borderRadius: 6, fontSize: 13, color: "#6ee7b7" }}>{t("reuseListLead", { list: offcuts.map(o => `${o.shape === "L" ? t("lShapeUse") : ""}${o.w}×${o.h}${SO_LANG === "ar" ? " مم" : SO_LANG === "ru" ? " мм" : "mm"}`).join(" · ") })}</div>}
      <div style={{ marginTop: 10, display: "flex", gap: 18, flexWrap: "wrap", padding: "10px 14px", background: "rgba(15,19,24,.6)", borderRadius: 6, border: "1px solid #2d3748" }}>
        <Mini label={t("miniSheet")} value={`${active + 1} / ${sheets.length}`} /><Mini label={t("miniThickness")} value={`${thickness} ${t("mm")}`} /><Mini label={t("miniParts")} value={sheet.length} /><Mini label={t("miniUtil")} value={`${util}%`} accent={util > 80 ? "#10B981" : util > 60 ? "#F59E0B" : "#EF4444"} />
      </div>
    </div>
  );
}

/* ─── REUSE-A-LEFTOVER-FILE BUTTON (wired) ───────────────────────────────── */
function LeftoverImporter({ expects, onImported }) {
  const t = useT();
  const ref = useRef(null);
  const [msg, setMsg] = useState("");
  const [busy, setBusy] = useState(false);
  const pick = async file => {
    if (!file) return;
    setBusy(true); setMsg(t("liReading", { name: file.name }));
    try {
      const res = await parseLeftoverFile(file);
      const modName = m => m === "plates" ? t("modPlatesName") : t("modSectionsName");
      if (expects && res.type !== expects) { setMsg(t("liWrongType", { type: modName(res.type), mod: modName(res.type) })); setBusy(false); return; }
      const n = res.items.reduce((s, it) => s + (it.qty || 1), 0);
      onImported(res.items);
      setMsg(t("liImported", { n: res.items.length, s: "", m: n, ms: "" }));
    } catch (e) { setMsg(t("liErr", { msg: (e.message || t("msgReadFail").replace(/^⚠ /, "")) })); }
    setBusy(false);
  };
  return (
    <div style={{ marginBottom: 12 }}>
      <button onClick={() => ref.current?.click()} disabled={busy} style={{ padding: "9px 18px", background: "rgba(59,130,246,.12)", border: "1px dashed #3b82f6", color: "#93c5fd", borderRadius: 8, cursor: busy ? "wait" : "pointer", fontSize: 15, fontWeight: 700, fontFamily: "'Space Mono', monospace" }}>
        {busy ? t("reuseBusy") : t("reuseImportBtn")}
      </button>
      <input ref={ref} type="file" accept=".pdf,.xlsx,.xls,.csv,.html,.htm" style={{ display: "none" }} onChange={e => { const f = e.target.files[0]; e.target.value = ""; pick(f); }} />
      {msg && <div style={{ marginTop: 8, fontSize: 14, color: msg.startsWith("✓") ? "#6ee7b7" : msg.startsWith("⚠") ? "#fcd34d" : "#94a3b8", fontFamily: "'Space Mono', monospace" }}>{msg}</div>}
    </div>
  );
}

Object.assign(LANG_DICT, {   // plates: default sheet for every thickness (section 13c)
  stockAllLbl:  { en: "Stock sheet size — used for every thickness, unless you add a different size for one thickness below", ar: "مقاس اللوح القياسي — يُستخدم لكل السماكات، إلا إذا أضفت أدناه مقاسًا مختلفًا لسماكة معيّنة", ru: "Размер листа — для всех толщин, если ниже не задан другой размер для конкретной толщины", zh: "标准板幅 — 适用于所有厚度；某一厚度规格不同时，可在下方单独添加", es: "Tamaño de chapa comercial — para todos los espesores, salvo que abajo defina otro tamaño para un espesor concreto" },
  anyThk:       { en: "All thicknesses", ar: "كل السماكات", ru: "Все толщины", zh: "所有厚度", es: "Todos los espesores" },
  thkInParts:   { en: "Thicknesses in your parts: {list} mm", ar: "السماكات في قطعك: {list} مم", ru: "Толщины ваших деталей: {list} мм", zh: "零件中的厚度：{list} mm", es: "Espesores de sus piezas: {list} mm" },
  plNoParts:    { en: "No plate parts to optimize yet — upload your list or type the parts in.", ar: "لا توجد قطع ألواح للتحسين بعد — ارفع قائمتك أو أدخل القطع يدويًا.", ru: "Пока нет деталей для раскроя — загрузите спецификацию или введите детали вручную.", zh: "还没有可优化的钢板零件 — 请上传清单或手动输入零件。", es: "Todavía no hay piezas de chapa para optimizar: cargue su lista o escriba las piezas." },
  upBannerP2:   { en: "💡 Thickness comes from the file. Every thickness uses the stock sheet size above, unless you add its own row.", ar: "💡 تؤخذ السماكة من الملف، وتُستخدم لكل سماكة مقاس اللوح أعلاه ما لم تضف لها صفًا خاصًا.", ru: "💡 Толщина берётся из файла. Для каждой толщины используется размер листа выше, если для неё не добавлена своя строка.", zh: "💡 厚度取自文件。每种厚度都使用上方的标准板幅，除非为其单独添加一行。", es: "💡 El espesor se toma del archivo. Cada espesor usa el tamaño de chapa de arriba, salvo que le añada su propia fila." },
});
const SO_STD_THK = [3, 4, 5, 6, 8, 10, 12, 14, 15, 16, 18, 20, 22, 25, 28, 30, 32, 35, 40, 45, 50, 60];   // suggestions only — any value can be typed

Object.assign(LANG_DICT, {   // sections: one stock length for every profile (section 13d)
  secStockLbl:  { en: "Stock bar length (mm)", ar: "طول العود القياسي (مم)", ru: "Длина хлыста (мм)", zh: "原材定尺长度 (mm)", es: "Longitud de barra comercial (mm)" },
  secStockHint: { en: "Used for every profile, unless a row in your list sets its own available length.", ar: "يُستخدم لكل المقاطع، إلا إذا حدّد سطر في قائمتك طولًا متاحًا خاصًا به.", ru: "Используется для всех профилей, если в строке ведомости не задана своя длина.", zh: "适用于所有截面；清单中某行单独填写可用长度时，以该行为准。", es: "Se usa para todos los perfiles, salvo que una fila de su lista indique su propia longitud disponible." },
  blankEq:      { en: "blank = {len}", ar: "فارغ = {len}", ru: "пусто = {len}", zh: "留空 = {len}", es: "vacío = {len}" },
});
/* ─── PLATES MODULE (wired) ──────────────────────────────────────────────── */
function PlatesModule({ onBack }) {
  const { t, lang } = useLang();
  const pricing = usePricing();
  const [inputMode, setInputMode] = useState(null);
  const [DEFAULT_SHEET, setDefaultSheet] = useState({ w: 1220, h: 2440 });   // every thickness, unless a row below says otherwise
  const [stock, setStock] = useState([]);
  const [offcutStock, setOffcutStock] = useState([]);
  const [parts, setParts] = useState([{ id: "P1", label: "P1", length: 500, width: 300, thickness: 12, qty: 4 }, { id: "P2", label: "P2", length: 380, width: 250, thickness: 12, qty: 6 }, { id: "P3", label: "P3", length: 600, width: 400, thickness: 20, qty: 3 }, { id: "P4", label: "P4", length: 280, width: 180, thickness: 8, qty: 8 }]);
  const [material, setMaterial] = useState("S235JR");
  const [kerf, setKerf] = useState(3); const [margin, setMargin] = useState(10); const [reuseMin, setReuseMin] = useState(300); const [allowRotation, setAllowRotation] = useState(true); const [splicePref, setSplicePref] = useState("welds");
  const [results, setResults] = useState(null); const [colorMap, setColorMap] = useState({}); const [excelMsg, setExcelMsg] = useState(""); const [columnMap, setColumnMap] = useState(null); const [isOpt, setIsOpt] = useState(false);
  const [noPartsMsg, setNoPartsMsg] = useState(false);
  const optBtnRef = useRef(null);
  const scrollToOpt = () => setTimeout(() => optBtnRef.current?.scrollIntoView({ behavior: "smooth", block: "center" }), 120);
  const pls = n => ((lang === "en" || lang === "es") && n > 1) ? "s" : "";

  const stockThk = [...new Set(stock.map(s => s.thickness).filter(t2 => t2 > 0))].sort((a, b) => a - b);
  const sheetFor = thk => { const h = stock.find(s => +s.thickness === +thk); return h ? { w: h.w, h: h.h } : DEFAULT_SHEET; };
  const addStock = () => setStock(s => [...s, { thickness: 0, w: 1220, h: 2440 }]);
  const upStock = (i, f, v) => setStock(s => s.map((r, ri) => ri === i ? { ...r, [f]: +v || 0 } : r));
  const delStock = i => setStock(s => s.filter((_, ri) => ri !== i));
  const addOffcut = () => setOffcutStock(o => [...o, { label: `Leftover ${o.length + 1}`, thickness: stockThk[0] || 8, shape: "rect", A: 600, B: 400, C: 0, D: 0, qty: 1, pts: defaultLPoints() }]);
  const upOffcut = (i, f, v) => setOffcutStock(o => o.map((r, ri) => ri === i ? { ...r, [f]: (f === "label" || f === "shape") ? v : +v || 0 } : r));
  const patchOffcut = (i, patch) => setOffcutStock(o => o.map((r, ri) => ri === i ? { ...r, ...patch } : r));
  const delOffcut = i => setOffcutStock(o => o.filter((_, ri) => ri !== i));
  const importOffcuts = items => setOffcutStock(o => [...o, ...items.map((it, k) => ({ label: it.label || `Imported ${o.length + k + 1}`, thickness: it.thickness || stockThk[0] || 8, shape: it.shape === "L" ? "L" : "rect", A: it.A || 0, B: it.B || 0, C: it.C || 0, D: it.D || 0, qty: Math.max(1, it.qty || 1), pts: defaultLPoints() }))]);
  const addPart = () => { const id = `P${parts.length + 1}`; setParts(p => [...p, { id, label: id, length: 400, width: 200, thickness: (p.length && +p[p.length - 1].thickness) || stockThk[0] || 10, qty: 1 }]); };
  const upPart = (i, f, v) => setParts(p => p.map((r, ri) => ri === i ? { ...r, [f]: f === "label" ? v : +v || 0 } : r));
  const delPart = i => setParts(p => p.filter((_, ri) => ri !== i));

  const PLATE_PFX = /^(PLATE|PLT|PL|FLT|FLAT|FL|FB|P)\s*[-_]?\s*\d/i;
  const HOT = /^(SHS|RHS|CHS|IPE|IPN|HEA|HEB|HEM|HE|UBP|UB|UC|PFC|UPN|UPE|JIS|HSS|HD|HP|UA|EA|RSA|RSJ|L|W|S|C|MC|BR|RD)/i;
  const isPlate = p => { const s = String(p || "").trim(); return PLATE_PFX.test(s) && !soTeklaNonPlate(s); };
  const isHot = p => { const s = String(p || "").trim(); if (soTeklaNonPlate(s)) return true; if (isPlate(s)) return false; return HOT.test(s); };
  const parsePlate = raw => { const p = String(raw || "").trim().toUpperCase(); if (!p) return null; const m = p.match(/^(?:PLATE|PLT|PL|FLT|FLAT|FL|FB|P)\s*[-_]?\s*(\d+(?:\.\d+)?)\s*[*xX×]\s*(\d+(?:\.\d+)?)/); if (m) { const a = +m[1], b = +m[2]; return { thickness: Math.min(a, b), width: Math.max(a, b) }; } if (!isHot(p)) { const b = p.match(/^(\d+(?:\.\d+)?)\s*[*xX×]\s*(\d+(?:\.\d+)?)/); if (b) { const x = +b[1], y = +b[2]; return { thickness: Math.min(x, y), width: Math.max(x, y) }; } } return null; };

  function extract(headers, rows, map, unit) {
    const out = []; const idx = c => c ? headers.indexOf(c) : -1; const pi = idx(map.profile), qi = idx(map.qty), li = idx(map.length), wi = idx(map.width), ti = idx(map.thickness);
    rows.forEach((row, ri) => {
      const cells = row.map(c => normHeader(c)); const joined = cells.join(" ");
      if (/^\s*(grand\s+)?total\b/i.test(joined) || /total\s+for|members/i.test(joined)) return;
      if (map.profile) { const pc = cells[headers.indexOf(map.profile)] || ""; if (/^total$/i.test(pc.trim())) return; }
      let thickness = 0, width = 0, length = 0, qty = 0;
      const pcell = pi >= 0 ? cells[pi] : (cells.find(c => isPlate(c) || isHot(c)) || "");
      if (pcell) { if (isHot(pcell)) return; const pl = parsePlate(pcell); if (pl) { thickness = pl.thickness; width = pl.width; } }
      if (li >= 0) length = parseLengthToMm(cells[li], unit) || 0; if (qi >= 0) qty = parseInt(cells[qi]) || 0; if (ti >= 0 && !thickness) thickness = parseLengthToMm(cells[ti], unit) || 0; if (wi >= 0 && !width) width = parseLengthToMm(cells[wi], unit) || 0;
      const havePlate = thickness > 0 && width > 0;
      if (!havePlate && !(length > 0 && width > 0)) {
        const nums = cells.flatMap(c => { const pr = c.match(/(\d+(?:\.\d+)?)\s*[*xX×]\s*(\d+(?:\.\d+)?)/); if (pr) return [+pr[1], +pr[2]]; const n = parseFloat(c.replace(/[^\d.]/g, "")); return isFinite(n) && c.match(/\d/) ? [n] : []; }).filter(n => n > 0 && n < 100000);
        if (nums.length >= 2) { const sorted = [...nums].sort((a, b) => a - b); if (!thickness) { const tc = sorted.find(n => n >= 3 && n <= 80); thickness = tc || 10; } const big = sorted.filter(n => n !== thickness); if (big.length >= 2) { width = width || big[0]; length = length || big[big.length - 1]; } else if (big.length === 1) { length = length || big[0]; width = width || big[0]; } }
      }
      if (!qty) { const qc = cells.find(c => /^\d{1,4}$/.test(c) && +c <= 999 && +c !== length && +c !== width && +c !== thickness); qty = qc ? +qc : 1; }
      if (!thickness) thickness = 10;
      if (length > 0 && width > 0) out.push({ id: `P${ri + 1}`, label: `P${ri + 1}`, length: Math.round(length * 10) / 10, width: Math.round(width * 10) / 10, thickness, qty: qty || 1 });
    });
    return out;
  }
  const parseExcel = useCallback(async file => {
    track("upload_excel", { module: "plates", ext: (file.name || "").split(".").pop() });
    alertMe("SOMEONE UPLOADED A FILE (plates)", { ext: (file.name || "").split(".").pop(), sizeKB: Math.round((file.size || 0) / 1024) });
    try {
      const buf = await file.arrayBuffer();
      {
        const wb = XLSX.read(buf, { type: "array" }); const raw = soPlateSheetRows(wb);   // first sheet with a header (Arabic sheets: isolated add-on)
        let hr = -1; for (let i = 0; i < Math.min(raw.length, 25); i++) { const r = raw[i].map(c => normHeader(c).toLowerCase()); if (r.some(c => /profile|length|width|qty|no\.?$|thick|partpos|part pos/i.test(c))) { hr = i; break; } } if (hr === -1) hr = 0;
        const headers = raw[hr].map(c => normHeader(c)); const rows = raw.slice(hr + 1).filter(r => r.some(c => c !== ""));
        const map = { profile: detectCol(headers, ["profile", "section", "size"]), qty: detectCol(headers, ["no.", "no", "qty", "number", "count", "quantity", "pcs"]), length: detectCol(headers, ["length (mm)", "length(mm)", "length", "l (mm)"]), width: detectCol(headers, ["width (mm)", "width(mm)", "width", "breadth"]), thickness: detectCol(headers, ["thickness", "thk", "t (mm)"]) };
        const below = (raw[hr + 1] || []).map(c => normHeader(c));
        const liH = headers.findIndex(h => /length/i.test(h));
        const unit = liH !== -1 ? (detectLenUnit(headers[liH]) || detectLenUnit(below[liH]) || null) : null;
        const out = extract(headers, rows, map, unit);
        if (out.length) { setParts(out); setExcelMsg(t("msgImportedP", { n: out.length, s: pls(out.length) })); setColumnMap(null); scrollToOpt(); }
        else { setColumnMap({ headers, map, rows }); setExcelMsg(t("msgMapCols")); }
      }
    } catch { setExcelMsg(t("msgReadFail")); }
  }, [t, lang]);
  const onDrop = useCallback(e => { e.preventDefault(); const f = e.dataTransfer?.files?.[0] || e.target?.files?.[0]; if (f) parseExcel(f); }, [parseExcel]);

  const optimize = () => {
    if (!parts.some(p => p.length > 0 && p.width > 0 && p.qty > 0)) { setNoPartsMsg(true); return; }   // nothing to nest: say so, no empty result
    setNoPartsMsg(false);
    track("optimization_started", { module: "plates", parts: parts.length });
    setIsOpt(true);
    setTimeout(() => {
      const valid = parts.filter(p => p.length > 0 && p.width > 0 && p.qty > 0);
      const cm = {}; valid.forEach((p, i) => cm[p.id] = PART_COLORS[i % PART_COLORS.length]); setColorMap(cm);
      const byThk = {}; valid.forEach(p => { const k = p.thickness || 10; (byThk[k] = byThk[k] || []).push(p); });
      const groups = Object.keys(byThk).map(Number).sort((a, b) => b - a).map(thk => {
        let gp = byThk[thk]; const { w: sw, h: sh } = sheetFor(thk);
        const myOffcuts = offcutStock.filter(o => +o.thickness === +thk && o.A > 0 && o.B > 0 && (o.qty || 1) > 0);
        const offcutPlans = [];
        myOffcuts.forEach(o => {
          const rect = o.shape === "L" ? lUsableRect(o.A, o.B, o.C, o.D) : { w: o.A, h: o.B };
          if (rect.w <= 0 || rect.h <= 0) return;
          const units = Math.max(1, Math.round(o.qty || 1));
          for (let u = 0; u < units; u++) {
            const res = nestIntoRect(rect.w, rect.h, gp, kerf, allowRotation);
            if (!res.placedAny) break;
            offcutPlans.push({ label: units > 1 ? `${o.label} (${u + 1}/${units})` : o.label, shape: o.shape, rectW: rect.w, rectH: rect.h, placements: res.placements });
            gp = res.leftover;
            if (!gp.length) break;
          }
        });
        const gsheets = nestPlates(sw, sh, gp, kerf, margin, allowRotation, reuseMin, splicePref);
        const usedArea = gsheets.reduce((s, st) => s + (st.placements || []).reduce((a, p) => a + p.w * p.h, 0), 0);
        const totalArea = gsheets.length * sw * sh; const wastePct = gsheets.length ? Math.round(((totalArea - usedArea) / totalArea) * 100) : 0;
        const allThkParts = byThk[thk];
        const partVol = allThkParts.reduce((s, p) => s + p.length * p.width * thk * p.qty, 0); const partWeight = (partVol / 1e9) * DENSITY; const sheetWeight = (gsheets.length * sw * sh * thk / 1e9) * DENSITY;
        const offcuts = []; gsheets.forEach((st, si) => (st.offcuts || []).forEach(o => { const area = o.shape === "L" ? (o.A * o.B - o.notchW * o.notchH) : (o.A * o.B); offcuts.push({ sheet: si + 1, ...o, area, weight: (area * thk / 1e9) * DENSITY }); }));
        const splices = (gsheets.splices || []).map(s => ({ label: s.label, W: s.W, L: s.L, count: s.count, qty: s.qty }));
        const reusedUsedCount = offcutPlans.length;
        const paintM2 = plateFacePaintM2(allThkParts);   // paint area: isolated add-on (one face, user's parts)
        return { thickness: thk, sw, sh, paintM2, sheets: gsheets, sheetCount: gsheets.length, partCount: allThkParts.reduce((s, p) => s + p.qty, 0), usedArea, totalArea, wastePct, utilPct: gsheets.length ? 100 - wastePct : 100, partWeight, sheetWeight, wasteWeight: sheetWeight - partWeight, offcuts, offcutWeight: offcuts.reduce((s, o) => s + o.weight, 0), splices, marginUsed: gsheets.marginUsed, offcutPlans, reusedUsedCount };
      });
      const totals = groups.reduce((a, g) => ({ sheets: a.sheets + g.sheetCount, parts: a.parts + g.partCount, sheetWeight: a.sheetWeight + g.sheetWeight, partWeight: a.partWeight + g.partWeight, wasteWeight: a.wasteWeight + g.wasteWeight, totalArea: a.totalArea + g.totalArea, usedArea: a.usedArea + g.usedArea, offcutCount: a.offcutCount + g.offcuts.length, offcutWeight: a.offcutWeight + g.offcutWeight, spliceCount: a.spliceCount + g.splices.reduce((ss, s) => ss + s.qty, 0), reusedUsed: a.reusedUsed + g.reusedUsedCount, paintM2: addPaint(a.paintM2, g.paintM2) }), { sheets: 0, parts: 0, sheetWeight: 0, partWeight: 0, wasteWeight: 0, totalArea: 0, usedArea: 0, offcutCount: 0, offcutWeight: 0, spliceCount: 0, reusedUsed: 0, paintM2: 0 });
      totals.utilPct = Math.round((totals.usedArea / totals.totalArea) * 100); totals.wastePct = 100 - totals.utilPct;
      track("optimization_completed", { module: "plates", sheets: totals?.totalSheets || 0, waste_pct: Math.round(parseFloat(totals?.wastePct || 0)) });
      if (inputMode === "manual") soNotifyManual("plates", { parts: valid.length, pieces: totals.parts, thicknesses: groups.length, sheets: totals.sheets, demo: soIsDemoPlates(parts) });
      setResults({ groups, totals }); setIsOpt(false);
    }, 350);
  };

  if (results) return <PlateResults results={results} colorMap={colorMap} material={material} reuseMin={reuseMin} pricing={pricing} onBack={() => setResults(null)} />;

  return (
    <>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}><SectionTitle>{t("platesWS")}</SectionTitle><BackChip onClick={onBack} labelKey="backModules" /></div>
      {inputMode === null && (
        <div>
          <div style={{ textAlign: "center", color: "#94a3b8", fontSize: 17, marginBottom: 18 }}>{t("howEnter")}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(260px,1fr))", gap: 20 }}>
            <Chooser icon="✏️" titleKey="manualEntry" descKey="manualDescP" onClick={() => setInputMode("manual")} />
            <Chooser icon="📊" titleKey="uploadTitle" descKey="uploadDescP" onClick={() => { if (soIsDemoPlates(parts)) setParts([]); setInputMode("excel"); }} />
            <Chooser icon="🔍" titleKey="scanTitle" descKey="scanDescP" onClick={() => { if (soIsDemoPlates(parts)) setParts([]); setInputMode("scan"); }} />
          </div>
        </div>
      )}
      {inputMode !== null && (
        <Card title={t("stockSettings")}>
          <Label>{t("stockAllLbl")}</Label>
          <div style={{ overflowX: "auto", marginTop: 6 }}>
            <table style={{ borderCollapse: "collapse", fontSize: 15 }}><thead><tr style={{ borderBottom: "1px solid #2d3748" }}>{[t("lblThickness"), t("lblWidth"), t("lblLength"), ""].map((h, hi) => <th key={hi} style={{ textAlign: "start", padding: "6px 10px", color: "#64748b", fontWeight: 600, fontSize: 13, letterSpacing: 1 }}>{h}</th>)}</tr></thead>
              <tbody><tr style={{ borderBottom: "1px solid #1a2230", background: "rgba(245,158,11,.06)" }}><td style={{ padding: "5px 10px", color: "#fbbf24", fontSize: 14, fontWeight: 700, whiteSpace: "nowrap" }}>{t("anyThk")}</td><td style={{ padding: "5px 10px" }}><input type="number" value={DEFAULT_SHEET.w} onChange={e => setDefaultSheet(d => ({ ...d, w: +e.target.value || 0 }))} style={{ ...CI, width: 90 }} /></td><td style={{ padding: "5px 10px" }}><input type="number" value={DEFAULT_SHEET.h} onChange={e => setDefaultSheet(d => ({ ...d, h: +e.target.value || 0 }))} style={{ ...CI, width: 90 }} /></td><td /></tr>{stock.map((s, i) => <tr key={i} style={{ borderBottom: "1px solid #1a2230" }}><td style={{ padding: "5px 10px" }}><input type="number" value={s.thickness} onChange={e => upStock(i, "thickness", e.target.value)} style={{ ...CI, width: 90, borderColor: "#d9770699" }} /></td><td style={{ padding: "5px 10px" }}><input type="number" value={s.w} onChange={e => upStock(i, "w", e.target.value)} style={{ ...CI, width: 90 }} /></td><td style={{ padding: "5px 10px" }}><input type="number" value={s.h} onChange={e => upStock(i, "h", e.target.value)} style={{ ...CI, width: 90 }} /></td><td style={{ padding: "5px 10px" }}><button onClick={() => delStock(i)} style={{ background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 16 }}>✕</button></td></tr>)}</tbody>
            </table>
          </div>
          <button onClick={addStock} style={{ marginTop: 10, padding: "7px 16px", background: "rgba(245,158,11,.15)", border: "1px dashed #d97706", color: "#fbbf24", borderRadius: 4, cursor: "pointer", fontSize: 14 }}>{t("addThickness")}</button>
          <div style={{ marginTop: 8, fontSize: 13, color: "#94a3b8", fontFamily: "'Space Mono', monospace" }}>{t("stockHint")}</div>
          {(() => { const th = [...new Set(parts.filter(p => p.length > 0 && p.width > 0 && p.qty > 0).map(p => +p.thickness || 10))].sort((a, b) => a - b); return th.length ? <div style={{ marginTop: 6, fontSize: 13, color: "#6ee7b7", fontFamily: "'Space Mono', monospace" }}>{t("thkInParts", { list: th.join(" · ") })}</div> : null; })()}
          <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end", marginTop: 18, paddingTop: 16, borderTop: "1px solid #1a2230" }}>
            <div><Label>{t("lblGrade")}</Label><select value={material} onChange={e => setMaterial(e.target.value)} style={SEL}>{STEEL_GRADES.map(m => <option key={m}>{m}</option>)}</select></div>
            <div><Label>{t("lblKerf")}</Label><input type="number" value={kerf} onChange={e => setKerf(+e.target.value)} style={{ ...IN, width: 90 }} /></div>
            <div><Label>{t("lblMargin")}</Label><input type="number" value={margin} onChange={e => setMargin(+e.target.value)} style={{ ...IN, width: 80 }} /></div>
            <div><Label>{t("lblReuseMin")}</Label><input type="number" value={reuseMin} onChange={e => setReuseMin(+e.target.value)} style={{ ...IN, width: 100 }} /></div>
            <div style={{ display: "flex", alignItems: "center", gap: 8, paddingBottom: 6 }}><input type="checkbox" id="rot" checked={allowRotation} onChange={e => setAllowRotation(e.target.checked)} style={{ width: 16, height: 16, accentColor: "#f59e0b" }} /><label htmlFor="rot" style={{ color: "#94a3b8", fontSize: 15 }}>{t("lblRotate")}</label></div>
            <div><Label>{t("spliceStrat")}</Label><select value={splicePref} onChange={e => setSplicePref(e.target.value)} style={{ ...SEL, minWidth: 220 }}><option value="welds">{t("spliceWelds")}</option><option value="pack">{t("splicePack")}</option></select></div>
          </div>
        </Card>
      )}
      {inputMode !== null && <PriceCard pricing={pricing} />}
      {inputMode !== null && (
        <Card title={t("reusePlatesTtl")} style={{ marginTop: 20 }}>
          <div style={{ fontSize: 15, color: "#6ee7b7", marginBottom: 14, lineHeight: 1.6 }}>{t("reusePlatesLead")}</div>
          <LeftoverImporter expects="plates" onImported={importOffcuts} />
          {offcutStock.length === 0 && (
            <div style={{ padding: "20px 18px", textAlign: "center", border: "1px dashed rgba(16,185,129,.35)", borderRadius: 12, background: "rgba(16,185,129,.04)", marginBottom: 12 }}>
              <div style={{ fontSize: 15, color: "#94a3b8", marginBottom: 12 }}>{t("noPlatesYet")}</div>
              <button onClick={addOffcut} style={{ padding: "10px 22px", background: "linear-gradient(135deg,#10b981,#059669)", border: "none", color: "#04140d", borderRadius: 8, cursor: "pointer", fontSize: 15, fontWeight: 700 }}>{t("addLeftPlate")}</button>
            </div>
          )}
          {offcutStock.map((o, i) => { const use = o.shape === "L" ? lUsableRect(o.A, o.B, o.C, o.D) : { w: o.A, h: o.B }; return (
            <div key={i} style={{ display: "flex", gap: 22, flexWrap: "wrap", alignItems: "flex-start", padding: 16, marginBottom: 14, borderRadius: 12, background: "rgba(10,14,20,.5)", border: "1px solid rgba(148,163,184,.12)" }}>
              <div style={{ flex: "1 1 340px", minWidth: 290 }}>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 16 }}>
                  <input value={o.label} onChange={e => upOffcut(i, "label", e.target.value)} placeholder={t("nameThisPlate")} style={{ ...CI, width: 170, fontWeight: 700 }} />
                  <button onClick={() => delOffcut(i)} style={{ marginInlineStart: "auto", background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 14 }}>{t("removeX")}</button>
                </div>
                <div style={{ marginBottom: 14 }}>
                  <Label>{t("whatShape")}</Label>
                  <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
                    {[["rect", t("shapeRect")], ["L", t("shapeL")]].map(([v, lbl]) => (
                      <button key={v} onClick={() => upOffcut(i, "shape", v)} style={{ flex: 1, padding: "9px 12px", borderRadius: 8, cursor: "pointer", fontSize: 14, fontWeight: 700, border: o.shape === v ? "1.5px solid #10b981" : "1px solid rgba(148,163,184,.2)", background: o.shape === v ? "rgba(16,185,129,.15)" : "transparent", color: o.shape === v ? "#6ee7b7" : "#94a3b8" }}>{lbl}</button>
                    ))}
                  </div>
                </div>
                {o.shape === "rect" ? (
                  <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
                    <div><Label>{t("lblWidth")}</Label><input type="number" value={o.A} onChange={e => upOffcut(i, "A", e.target.value)} style={{ ...CI, width: 100 }} /></div>
                    <div><Label>{t("lblLength")}</Label><input type="number" value={o.B} onChange={e => upOffcut(i, "B", e.target.value)} style={{ ...CI, width: 100 }} /></div>
                    <div><Label>{t("lblThickness")}</Label><input type="number" value={o.thickness} onChange={e => upOffcut(i, "thickness", e.target.value)} style={{ ...CI, width: 100, borderColor: "#d9770699" }} /></div>
                    <div><Label>{t("lblQtyHow")}</Label><input type="number" min="1" value={o.qty} onChange={e => upOffcut(i, "qty", e.target.value)} style={{ ...CI, width: 90, borderColor: "#10b98199" }} /></div>
                  </div>
                ) : (
                  <div>
                    <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
                      <div><Label>{t("lblThickness")}</Label><input type="number" value={o.thickness} onChange={e => upOffcut(i, "thickness", e.target.value)} style={{ ...CI, width: 100, borderColor: "#d9770699" }} /></div>
                      <div><Label>{t("lblQtyHow")}</Label><input type="number" min="1" value={o.qty} onChange={e => upOffcut(i, "qty", e.target.value)} style={{ ...CI, width: 90, borderColor: "#10b98199" }} /></div>
                    </div>
                    <LPointBuilder points={o.pts} onChange={(pts, info) => patchOffcut(i, { pts, A: info.A, B: info.B, C: info.C, D: info.D })} />
                  </div>
                )}
                <div style={{ marginTop: 14, padding: "10px 14px", borderRadius: 8, background: "rgba(16,185,129,.1)", border: "1px solid rgba(16,185,129,.3)", fontSize: 14, color: "#6ee7b7" }}>{t("willNestInto")} <b>{Math.round(use.w)} × {Math.round(use.h)} {t("mm")}</b>{lang === "en" ? " area" : ""}{o.shape === "L" ? t("biggestRectL") : ""}{(o.qty || 1) > 1 ? <> × <b>{Math.round(o.qty)}</b> {o.qty > 1 ? t("piecesWord") : t("pieceWord")}</> : ""}.</div>
              </div>
              {o.shape === "rect" && <div style={{ flex: "0 0 auto" }}><LShapeDiagram A={o.A} B={o.B} C={0} D={0} t={o.thickness} /></div>}
            </div>
          ); })}
          {offcutStock.length > 0 && <button onClick={addOffcut} style={{ marginTop: 4, padding: "9px 20px", background: "rgba(16,185,129,.12)", border: "1px dashed #10b981", color: "#6ee7b7", borderRadius: 8, cursor: "pointer", fontSize: 15, fontWeight: 700 }}>{t("addLeftMore")}</button>}
        </Card>
      )}
      {inputMode === "manual" && (
        <Card title={t("manualPartEntry")} style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} labelKey="changeMethod" />}>
          <div style={{ overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 15 }}><thead><tr style={{ borderBottom: "1px solid #2d3748" }}>{[t("lblId"), t("lblLength"), t("lblWidth"), t("lblThickness"), t("lblQty"), ""].map((h, hi) => <th key={hi} style={{ textAlign: "start", padding: "6px 8px", color: "#64748b", fontWeight: 600, fontSize: 13, letterSpacing: 1 }}>{h}</th>)}</tr></thead>
            <tbody>{parts.map((p, i) => <tr key={i} style={{ borderBottom: "1px solid #1a2230" }}><td style={{ padding: "5px 4px" }}><input value={p.label} onChange={e => upPart(i, "label", e.target.value)} style={{ ...CI, width: 50 }} /></td><td style={{ padding: "5px 4px" }}><input type="number" value={p.length} onChange={e => upPart(i, "length", e.target.value)} style={{ ...CI, width: 80 }} /></td><td style={{ padding: "5px 4px" }}><input type="number" value={p.width} onChange={e => upPart(i, "width", e.target.value)} style={{ ...CI, width: 80 }} /></td><td style={{ padding: "5px 4px" }}><input type="number" min="0" step="any" list="so-thk-list" value={p.thickness} onChange={e => upPart(i, "thickness", e.target.value)} style={{ ...CI, width: 90, borderColor: "#d9770699" }} />{i === 0 && <datalist id="so-thk-list">{[...new Set([...stockThk, ...SO_STD_THK])].sort((a, b) => a - b).map(tk => <option key={tk} value={tk} />)}</datalist>}</td><td style={{ padding: "5px 4px" }}><input type="number" value={p.qty} onChange={e => upPart(i, "qty", e.target.value)} style={{ ...CI, width: 55 }} /></td><td style={{ padding: "5px 4px" }}><button onClick={() => delPart(i)} style={{ background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 16 }}>✕</button></td></tr>)}</tbody>
          </table></div>
          <button onClick={addPart} style={{ marginTop: 12, padding: "8px 18px", background: "rgba(245,158,11,.15)", border: "1px dashed #d97706", color: "#fbbf24", borderRadius: 4, cursor: "pointer", fontSize: 14, width: "100%" }}>{t("addPart")}</button>
        </Card>
      )}
      {inputMode === "excel" && (
        <Card title={t("uploadTeklaTtl")} style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} labelKey="changeMethod" />}>
          <div style={{ marginBottom: 14, padding: "10px 14px", borderRadius: 6, background: "rgba(245,158,11,.08)", border: "1px solid rgba(245,158,11,.25)", color: "#fcd34d", fontSize: 14, lineHeight: 1.6 }}>{t("upBannerP2")}</div>
          <div onDrop={onDrop} onDragOver={e => e.preventDefault()} onClick={() => document.getElementById("pxlsx").click()} style={{ border: "2px dashed #2d3748", borderRadius: 8, padding: "48px 20px", textAlign: "center", background: "rgba(15,19,24,.5)", cursor: "pointer" }}>
            <div style={{ fontSize: 48, marginBottom: 12 }}>📋</div><div style={{ color: "#94a3b8", fontSize: 17, marginBottom: 6 }}>{t("dropTekla")}</div><div style={{ color: "#475569", fontSize: 14 }}>{t("orBrowse")}</div><div style={{ marginTop: 14, color: "#3a4452", fontSize: 12, lineHeight: 1.7 }}>{t("upHintP")}</div>
          </div>
          <input id="pxlsx" type="file" accept=".xlsx,.xls,.csv,.txt" style={{ display: "none" }} onChange={onDrop} />
          {excelMsg && <div style={{ marginTop: 10, padding: "8px 12px", borderRadius: 4, background: excelMsg.startsWith("✓") ? "rgba(16,185,129,.15)" : "rgba(245,158,11,.15)", border: `1px solid ${excelMsg.startsWith("✓") ? "rgba(16,185,129,.4)" : "rgba(245,158,11,.4)"}`, color: excelMsg.startsWith("✓") ? "#6ee7b7" : "#fcd34d", fontSize: 14 }}>{excelMsg}</div>}
          {columnMap && <div style={{ marginTop: 12, padding: 12, background: "rgba(15,19,24,.8)", borderRadius: 6, border: "1px solid #2d3748" }}><div style={{ fontSize: 14, color: "#94a3b8", marginBottom: 8 }}>{t("mapColumns")}</div>{["profile", "qty", "length", "width", "thickness"].map(f => <div key={f} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}><span style={{ fontSize: 13, color: "#64748b", width: 70 }}>{f}</span><select value={columnMap.map[f] || ""} onChange={e => setColumnMap(c => ({ ...c, map: { ...c.map, [f]: e.target.value } }))} style={{ ...SEL, flex: 1 }}><option value="">—</option>{columnMap.headers.map(h => <option key={h} value={h}>{h}</option>)}</select></div>)}<button onClick={() => { const out = extract(columnMap.headers, columnMap.rows, columnMap.map); if (out.length) { setParts(out); setExcelMsg(t("msgImportedP2", { n: out.length })); setColumnMap(null); scrollToOpt(); } else setExcelMsg(t("msgExtractFail")); }} style={BTN}>{t("applyMapping")}</button></div>}
          {parts.length > 0 && <div style={{ marginTop: 14 }}><div style={{ fontSize: 13, color: "#475569", marginBottom: 6, letterSpacing: 1 }}>{t("partsReady")} ({parts.length})</div><div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>{parts.slice(0, 16).map((p, i) => <div key={i} style={{ padding: "3px 8px", borderRadius: 3, background: (colorMap[p.id] || PART_COLORS[i % PART_COLORS.length]) + "22", border: `1px solid ${(colorMap[p.id] || PART_COLORS[i % PART_COLORS.length])}55`, fontSize: 12, color: "#94a3b8", fontFamily: "'Space Mono', monospace" }}>{p.label}: {p.length}×{p.width}×{p.thickness} ×{p.qty}</div>)}{parts.length > 16 && <div style={{ fontSize: 12, color: "#475569", alignSelf: "center" }}>+{parts.length - 16} {t("moreWord")}</div>}</div></div>}
        </Card>
      )}
      {inputMode === "scan" && (
        <Card title={t("scanCardTtl")} style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} labelKey="changeMethod" />}>
          <ScanImporter mode="plates" onUse={rs => {
            setParts(rs.map((r, i) => ({ id: "S" + (i + 1), label: r.label || ("P" + (i + 1)), length: r.length, width: r.width, thickness: r.thickness, qty: r.qty })));
            setExcelMsg(t("scanUsedOk", { n: rs.length, s: pls(rs.length) }));
            scrollToOpt();
          }} />
          {excelMsg && <div style={{ marginTop: 12, padding: "8px 12px", borderRadius: 4, background: "rgba(16,185,129,.15)", border: "1px solid rgba(16,185,129,.4)", color: "#6ee7b7", fontSize: 14 }}>{excelMsg}</div>}
        </Card>
      )}
      {inputMode !== null && <div ref={optBtnRef} style={{ textAlign: "center", marginTop: 32, scrollMarginTop: 80 }}><button onClick={() => (parts.some(p => p.length > 0 && p.width > 0 && p.qty > 0) ? soRequireMember(optimize, true) : optimize())} disabled={isOpt} style={OPT_BTN(isOpt)}>{isOpt ? t("optimizing") : t("optPlates")}</button>{noPartsMsg && !parts.some(p => p.length > 0 && p.width > 0 && p.qty > 0) && <div style={{ marginTop: 12, color: "#fcd34d", fontSize: 15 }}>⚠ {t("plNoParts")}</div>}</div>}
    </>
  );
}

/* ─── BigTon (wired — only the "Ton" suffix needs the language) ───────────── */
function BigTon({ label, value, note, color, big }) {
  const t = useT();
  return <div style={{ padding: "20px 22px", borderRadius: 10, background: big ? `${color}1a` : "#0f1318", border: `1px solid ${color}55`, textAlign: "center" }}><div style={{ fontSize: 12, color: "#64748b", letterSpacing: 1.5, textTransform: "uppercase", marginBottom: 8, fontFamily: "'Space Mono', monospace" }}>{label}</div><div style={{ fontFamily: "'Space Mono', monospace", fontWeight: 800, color, fontSize: big ? 48 : 36, lineHeight: 1 }}>{value}<span style={{ fontSize: big ? 20 : 16, marginLeft: 4 }}>{t("t_ton")}</span></div><div style={{ fontSize: 13, color: "#64748b", marginTop: 8, fontFamily: "'Space Mono', monospace" }}>{note}</div></div>;
}

/* ─── PLATE RESULTS (wired) ──────────────────────────────────────────────── */
function PlateResults({ results, colorMap, material, reuseMin, pricing, onBack }) {
  const { t, lang } = useLang();
  const pls = n => ((lang === "en" || lang === "es") && n > 1) ? "s" : "";
  const ples = n => ((lang === "en" || lang === "es") && n > 1) ? "es" : "";
  const topRef = useRef(null);
  useEffect(() => { const tm = setTimeout(() => topRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }), 60); return () => clearTimeout(tm); }, []);
  const wastePct = results.totals.wastePct;
  const wasteBg = wastePct <= 8 ? "linear-gradient(135deg,#16a34a,#15803d)" : wastePct <= 20 ? "linear-gradient(135deg,#f59e0b,#d97706)" : "linear-gradient(135deg,#dc2626,#991b1b)";
  return (
    <>
      <div ref={topRef} style={{ scrollMarginTop: 12 }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24, flexWrap: "wrap", gap: 12 }}><SectionTitle>{t("optResults")}</SectionTitle><BackChip onClick={onBack} labelKey="backInput" /></div>
      <div style={{ background: "#1c1600", border: "2px solid #f59e0b", borderRadius: 12, padding: "20px 24px", marginBottom: 16 }}>
        <div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#f59e0b", letterSpacing: ".15em", marginBottom: 14 }}>{t("summaryBuy")}</div>
        {results.groups.map((g, gi) => <div key={gi} style={{ display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap", padding: "8px 0", borderBottom: gi < results.groups.length - 1 ? "1px dashed #3a2e0a" : "none" }}><span style={{ fontFamily: "'Playfair Display', serif", fontWeight: 900, fontSize: 29, color: "#f59e0b", lineHeight: 1 }}>{g.sheetCount}</span><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#cbd5e1" }}>{soCountWord(t, g.sheetCount, "sheet", "sheets")} {t("sheetOf")}</span><span style={{ fontFamily: "'Playfair Display', serif", fontWeight: 700, fontSize: 22, color: "#f8fafc" }}>{g.thickness} {t("mm")}</span><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#64748b" }}>— {g.sw} × {g.sh} {t("mm")} ({material})</span><span style={{ marginInlineStart: "auto", fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#22c55e", fontWeight: 700 }}>{fmtTon(g.sheetWeight)} {t("t_ton")}</span></div>)}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginTop: 14, paddingTop: 12, borderTop: "1px solid #3a2e0a" }}><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#cbd5e1", fontWeight: 700 }}>{t("acrossThk", { n: results.totals.sheets, s: pls(results.totals.sheets), g: results.groups.length, es: ples(results.groups.length) })}</span><span style={{ fontFamily: "'Playfair Display', serif", fontSize: 25, fontWeight: 900, color: "#f59e0b" }}>{fmtTon(results.totals.sheetWeight)} {t("t_ton")}</span></div>
      </div>
      <div style={{ background: wasteBg, borderRadius: 12, padding: "20px 24px", marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}><div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#fff", opacity: .85, letterSpacing: ".1em" }}>{t("wasteAfterCut")}</div><div style={{ fontFamily: "'Playfair Display', serif", fontSize: 44, fontWeight: 900, color: "#fff" }}>{wastePct}%</div></div><div style={{ textAlign: "end" }}><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#fff", opacity: .85 }}>{t("scrapWeight")}</div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 20, color: "#fff", fontWeight: 700 }}>{fmtTon(results.totals.wasteWeight)} {t("t_ton")}</div>{results.totals.offcutCount > 0 && <div style={{ fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#fff", opacity: .9, marginTop: 4 }}>♻ {results.totals.offcutCount} {t("reusableShort")} ({fmtTon(results.totals.offcutWeight)} {t("t_ton")})</div>}</div></div>
      <CostBlock buyKg={results.totals.sheetWeight} netKg={results.totals.partWeight} wasteKg={results.totals.wasteWeight} offcutKg={results.totals.offcutWeight} pricing={pricing} />
      <PaintBlock total={results.totals.paintM2} note={t("paintNoteP")} missing={[]} count={results.groups.length} />
      {results.totals.spliceCount > 0 && (
        <div style={{ marginBottom: 20, padding: "16px 18px", background: "rgba(245,158,11,.1)", border: "1px solid rgba(245,158,11,.45)", borderRadius: 10 }}>
          <div style={{ fontSize: 15, fontWeight: 700, color: "#fbbf24", marginBottom: 8, fontFamily: "'Space Mono', monospace" }}>{t("spliceBannerHd", { n: results.totals.spliceCount, areis: lang === "en" ? (results.totals.spliceCount > 1 ? "s are" : " is") : "" })}</div>
          <div style={{ fontSize: 15, color: "#fcd34d", lineHeight: 1.6 }}>{t("spliceBannerBd")}</div>
          <div style={{ marginTop: 10, display: "flex", flexWrap: "wrap", gap: 6 }}>{results.groups.flatMap(g => g.splices.map((s, k) => <span key={`${g.thickness}-${k}`} style={{ padding: "4px 10px", borderRadius: 4, background: "rgba(245,158,11,.15)", border: "1px solid rgba(245,158,11,.4)", color: "#fcd34d", fontSize: 13, fontFamily: "'Space Mono', monospace" }}>{s.label}: {Math.round(s.W)}×{Math.round(s.L)} {t("mm")} @ {g.thickness} {t("mm")} → {s.count} {t("piecesWord")}{s.qty > 1 ? ` ×${s.qty}` : ""}</span>))}</div>
        </div>
      )}
      {results.totals.reusedUsed > 0 && (
        <div style={{ marginBottom: 20, padding: "14px 18px", background: "rgba(16,185,129,.1)", border: "1px solid rgba(16,185,129,.4)", borderRadius: 10, fontSize: 15, color: "#6ee7b7", lineHeight: 1.6 }}>
          {t("reusedBannerP", { n: results.totals.reusedUsed, s: pls(results.totals.reusedUsed) })}
        </div>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))", gap: 12, marginBottom: 24 }}>
        {[{ l: t("totalParts"), v: results.totals.parts, i: "◼" }, { l: t("totalSheetsCard"), v: results.totals.sheets, i: "📄", a: "#3B82F6" }, { l: t("thkGroups"), v: results.groups.length, i: "≡", a: "#8B5CF6" }, { l: t("utilization"), v: `${results.totals.utilPct}%`, i: "📊", a: results.totals.utilPct > 80 ? "#10B981" : results.totals.utilPct > 60 ? "#F59E0B" : "#EF4444" }, { l: t("purchaseWt"), v: `${fmtTon(results.totals.sheetWeight)} ${t("t_ton")}`, i: "🧾", a: "#f59e0b" }, { l: t("netWt"), v: `${fmtTon(results.totals.partWeight)} ${t("t_ton")}`, i: "⚖" }, { l: t("scrapWeight"), v: `${fmtTon(results.totals.wasteWeight)} ${t("t_ton")}`, i: "🗑", a: "#EF4444" }, { l: t("reusableOffcuts"), v: `${results.totals.offcutCount} (${fmtTon(results.totals.offcutWeight)} ${t("t_ton")})`, i: "♻", a: "#10B981" }, ...(results.totals.paintM2 != null ? [{ l: t("paintOneFace"), v: `${fmtM2(results.totals.paintM2)} ${t("m2")}`, i: "🎨", a: "#38bdf8" }] : [])].map((c, ci) => <StatCard key={ci} {...c} />)}
      </div>
      {results.groups.map((g, gi) => <Card key={gi} title={t("nestLayoutTtl", { thk: g.thickness, n: g.sheetCount, s: pls(g.sheetCount) })} style={{ marginBottom: 20 }}><PlateCanvas sheets={g.sheets} sheetW={g.sw} sheetH={g.sh} colorMap={colorMap} thickness={g.thickness} /></Card>)}
      {results.totals.offcutCount > 0 && (
        <Card title={t("offcutRegTtl")} style={{ marginBottom: 20 }}>
          <div style={{ fontSize: 14, color: "#6ee7b7", marginBottom: 12, background: "rgba(16,185,129,.08)", border: "1px solid rgba(16,185,129,.25)", borderRadius: 6, padding: "8px 12px" }}>{t("offcutRegLead", { n: results.totals.offcutCount, s: pls(results.totals.offcutCount), t: fmtTon(results.totals.offcutWeight), min: reuseMin })}</div>
          <div style={{ overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 15, fontFamily: "'Space Mono', monospace" }}><thead><tr style={{ background: "rgba(16,185,129,.12)", borderBottom: "1px solid rgba(16,185,129,.3)" }}>{[t("thThickness"), t("thFromSheet"), t("thShape"), t("thOverall"), t("thUsableRect"), t("thWeight"), t("thStatus")].map((h, i) => <th key={i} style={{ padding: "10px 12px", textAlign: i > 4 ? "end" : "start", color: "#6ee7b7", fontSize: 13, letterSpacing: 1, fontWeight: 700 }}>{h}</th>)}</tr></thead>
            <tbody>{results.groups.flatMap(g => g.offcuts.map((o, oi) => <tr key={`${g.thickness}-${oi}`} style={{ borderBottom: "1px solid #1a2230" }}><td style={{ padding: "9px 12px", color: "#cbd5e1" }}>{g.thickness} {t("mm")}</td><td style={{ padding: "9px 12px", color: "#94a3b8" }}>{t("fromSheetN", { n: o.sheet })}</td><td style={{ padding: "9px 12px" }}>{o.shape === "L" ? <span style={{ color: "#fbbf24", fontWeight: 700 }}>{t("shapeLshort")}</span> : <span style={{ color: "#94a3b8" }}>{t("shapeRectShort")}</span>}</td><td style={{ padding: "9px 12px", color: "#cbd5e1" }}>{o.shape === "L" ? `${o.A}×${o.B} − ${t("notchWord")} ${o.notchW}×${o.notchH}` : `${o.A} × ${o.B} ${t("mm")}`}</td><td style={{ padding: "9px 12px", color: "#6ee7b7", fontWeight: 700 }}>{o.w} × {o.h} {t("mm")}</td><td style={{ padding: "9px 12px", textAlign: "end", color: "#94a3b8" }}>{o.weight >= 1 ? `${o.weight.toFixed(1)} ${t("kg")}` : `${(o.weight * 1000).toFixed(0)} ${t("gU")}`}</td><td style={{ padding: "9px 12px", textAlign: "end", color: "#10B981", fontWeight: 700 }}>{t("keepStatus")}</td></tr>))}</tbody>
          </table></div>
        </Card>
      )}
      <Card title={t("dlReports")} style={{ marginBottom: 28 }}>
        <div style={{ fontSize: 14, color: "#94a3b8", marginBottom: 12, lineHeight: 1.5 }}>{t("dlTwoDeliv")}</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button onClick={() => requestDownload(() => exportPlatePDF(results, material, reuseMin, t, pricing))} style={EXP}>{t("dlProcPDF")}</button>
          <button onClick={() => requestDownload(() => exportPlateExcel(results, material, reuseMin, t, pricing))} style={EXP}>{t("dlProcXLS")}</button>
          {results.totals.offcutCount > 0 && <button onClick={() => requestDownload(() => exportOffcutPDF(results, material, reuseMin, t))} style={{ ...EXP, background: "linear-gradient(135deg,#10b981,#059669)", color: "#04140d" }}>{t("dlLeftPDF")}</button>}
          {results.totals.offcutCount > 0 && <button onClick={() => requestDownload(() => exportOffcutExcel(results, material, reuseMin, t))} style={{ ...EXP, background: "linear-gradient(135deg,#10b981,#059669)", color: "#04140d" }}>{t("dlLeftXLS")}</button>}
        </div>
      </Card>
      <ProcurementBlock results={results} material={material} />
    </>
  );
}

/* ─── PROCUREMENT BLOCK (wired) ──────────────────────────────────────────── */
function ProcurementBlock({ results, material }) {
  const { t, lang } = useLang();
  const pls = n => ((lang === "en" || lang === "es") && n > 1) ? "s" : "";
  const ples = n => ((lang === "en" || lang === "es") && n > 1) ? "es" : "";
  return (
    <div style={{ border: "2px solid rgba(245,158,11,.35)", borderRadius: 12, overflow: "hidden", boxShadow: "0 0 60px rgba(245,158,11,.08)" }}>
      <div style={{ background: "linear-gradient(135deg,rgba(245,158,11,.25),rgba(217,119,6,.2))", borderBottom: "1px solid rgba(245,158,11,.3)", padding: "20px 28px", display: "flex", alignItems: "center", gap: 14 }}><span style={{ fontSize: 32 }}>📦</span><div><div style={{ fontSize: 13, letterSpacing: 3, color: "#fbbf24", textTransform: "uppercase", marginBottom: 3, fontFamily: "'Space Mono', monospace" }}>{t("pbFinalReady")}</div><div style={{ fontSize: 25, fontWeight: 700, color: "#f8fafc", fontFamily: "'Playfair Display', serif" }}>{t("pbProcSummary")}</div></div><div style={{ marginInlineStart: "auto", textAlign: "end" }}><div style={{ fontSize: 12, color: "#94a3b8", letterSpacing: 1, fontFamily: "'Space Mono', monospace" }}>{t("material")}</div><div style={{ fontSize: 18, color: "#f59e0b", fontWeight: 700 }}>{material}</div></div></div>
      <div style={{ padding: "26px 28px", background: "rgba(10,13,18,.92)" }}>
        <div style={{ marginBottom: 22, padding: "18px 20px", background: "rgba(245,158,11,.08)", border: "1px solid rgba(245,158,11,.3)", borderRadius: 10 }}>
          <div style={{ fontSize: 13, letterSpacing: 2, color: "#fbbf24", textTransform: "uppercase", marginBottom: 12, fontFamily: "'Space Mono', monospace" }}>{t("whatToBuy")}</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>{results.groups.map((g, gi) => <div key={gi} style={{ display: "flex", alignItems: "center", gap: 12, fontSize: 19, color: "#f8fafc", fontFamily: "'Playfair Display', serif" }}><span style={{ color: "#f59e0b", fontSize: 20 }}>✓</span><span>{t("youNeed")} <b style={{ color: "#fbbf24", fontSize: 22 }}>{g.sheetCount}</b> {soCountWord(t, g.sheetCount, "sheet", "sheets")} {t("sheetOf")} <b style={{ color: "#f59e0b" }}>{g.thickness} {t("mm")}</b> <span style={{ color: "#94a3b8", fontSize: 16, fontFamily: "'Space Mono', monospace" }}>({g.sw} × {g.sh} {t("mm")}, {material})</span></span></div>)}</div>
        </div>
        <div style={{ overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 16, fontFamily: "'Space Mono', monospace" }}><thead><tr style={{ background: "rgba(245,158,11,.15)", borderBottom: "2px solid rgba(245,158,11,.3)" }}>{[t("thThickness"), t("thSheetSize"), t("thSheetsReq"), t("thPurchaseWt"), t("thParts"), t("utilization"), t("thNetWt"), t("thScrapWt")].map((h, i) => <th key={i} style={{ padding: "13px 14px", textAlign: i < 2 ? "start" : i === 7 ? "end" : "center", color: "#fbbf24", fontSize: 13, letterSpacing: 1, fontWeight: 700 }}>{h}</th>)}</tr></thead>
          <tbody>{results.groups.map((g, gi) => <tr key={gi} style={{ borderBottom: "1px solid #1a2230" }}><td style={{ padding: "15px 14px" }}><span style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "rgba(245,158,11,.12)", border: "1px solid rgba(245,158,11,.3)", borderRadius: 4, padding: "4px 12px" }}><span style={{ color: "#f59e0b" }}>✓</span><span style={{ color: "#f8fafc", fontWeight: 700, fontSize: 17 }}>{g.thickness} {t("mm")}</span></span></td><td style={{ padding: "15px 14px", color: "#94a3b8" }}>{g.sw} × {g.sh} {t("mm")}</td><td style={{ padding: "15px 14px", textAlign: "center" }}><span style={{ display: "inline-block", background: "rgba(59,130,246,.18)", border: "1px solid rgba(59,130,246,.4)", borderRadius: 6, padding: "6px 16px", color: "#93c5fd", fontWeight: 800, fontSize: 20 }}>{g.sheetCount} {soCountWord(t, g.sheetCount, "sheet", "sheets")}</span></td><td style={{ padding: "15px 14px", textAlign: "center", color: "#f59e0b", fontWeight: 700 }}>{fmtTon(g.sheetWeight)} {t("t_ton")}</td><td style={{ padding: "15px 14px", textAlign: "center", color: "#94a3b8" }}>{g.partCount}</td><td style={{ padding: "15px 14px", textAlign: "center", fontWeight: 700, color: g.utilPct > 80 ? "#10B981" : g.utilPct > 60 ? "#F59E0B" : "#EF4444" }}>{g.utilPct}%</td><td style={{ padding: "15px 14px", textAlign: "center", color: "#94a3b8" }}>{fmtTon(g.partWeight)} {t("t_ton")}</td><td style={{ padding: "15px 14px", textAlign: "end", color: "#f08a8a" }}>{fmtTon(g.wasteWeight)} {t("t_ton")}</td></tr>)}
            <tr style={{ background: "rgba(245,158,11,.12)", borderTop: "2px solid rgba(245,158,11,.4)" }}><td style={{ padding: "16px 14px", fontWeight: 800, color: "#f8fafc" }}>{t("totalSheets")}</td><td style={{ padding: "16px 14px", color: "#94a3b8" }}>{results.groups.length} {t("thicknessesWord", { es: ples(results.groups.length) })}</td><td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 800, fontSize: 20, color: "#93c5fd" }}>{results.totals.sheets} {soCountWord(t, results.totals.sheets, "sheet", "sheets")}</td><td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 800, color: "#f59e0b" }}>{fmtTon(results.totals.sheetWeight)} {t("t_ton")}</td><td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 700, color: "#f8fafc" }}>{results.totals.parts}</td><td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 800, color: results.totals.utilPct > 80 ? "#10B981" : results.totals.utilPct > 60 ? "#F59E0B" : "#EF4444" }}>{results.totals.utilPct}%</td><td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 800, color: "#f8fafc" }}>{fmtTon(results.totals.partWeight)} {t("t_ton")}</td><td style={{ padding: "16px 14px", textAlign: "end", fontWeight: 800, color: "#f08a8a" }}>{fmtTon(results.totals.wasteWeight)} {t("t_ton")}</td></tr>
          </tbody></table></div>
        <div style={{ marginTop: 12, padding: "10px 14px", background: "rgba(15,19,24,.5)", border: "1px solid #1a2230", borderRadius: 6, fontSize: 13, color: "#94a3b8", lineHeight: 1.7, fontFamily: "'Space Mono', monospace" }}>{t("pbLegend")}</div>
        <div style={{ marginTop: 22, display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 16 }}>
          <BigTon label={t("pbTotalPurchase")} value={fmtTon(results.totals.sheetWeight)} note={t("pbGrossNote", { n: results.totals.sheets })} color="#f59e0b" big />
          <BigTon label={t("netWt")} value={fmtTon(results.totals.partWeight)} note={t("pbNetNote")} color="#10B981" />
          <BigTon label={t("scrapWeight")} value={fmtTon(results.totals.wasteWeight)} note={t("pbScrapNote", { p: results.totals.wastePct })} color="#EF4444" />
        </div>
        <div style={{ marginTop: 18, textAlign: "end", fontFamily: "'Space Mono', monospace", fontSize: 12, color: "#475569" }}>{t("pbGenerated")} · {new Date().toLocaleDateString(lang === "ar" ? "ar-u-nu-latn" : soDateLoc("en-GB"), { day: "2-digit", month: "short", year: "numeric" })} · {material}</div>
      </div>
    </div>
  );
}


/* ============================================================================
   STEEL OPTIMIZER — combined Plates + Sections (Stage 1)
   Theme: amber / dark engineering.  Imports only xlsx.
============================================================================ */

const DENSITY = 7850; // kg/m3
/* Shared header/text helpers — used by both importers (hoisted to avoid
   three duplicate definitions across modules). */
function normHeader(s) { return String(s == null ? "" : s).replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim(); }
function detectCol(h, al) {
  const H = h.map(normHeader);
  for (let i = 0; i < H.length; i++) for (const a of al) if (H[i].toLowerCase() === a.toLowerCase()) return h[i];
  for (let i = 0; i < H.length; i++) for (const a of al) if (H[i].toLowerCase().includes(a.toLowerCase())) return h[i];
  return null;
}

/* Length units and parsing.
   Cut lists arrive in mm, m, or — from Advance Steel and most US fabricators —
   as imperial feet-inch-fraction strings like 5'-10 15/16". Everything is
   normalised to millimetres here so the engine only ever sees mm. */
function detectLenUnit(txt) {
  const s = String(txt == null ? "" : txt).toLowerCase();
  if (/\bmm\b|\(\s*mm\s*\)|\[\s*mm\s*\]|millimet/.test(s)) return "mm";
  if (/\bcm\b|\(\s*cm\s*\)|\[\s*cm\s*\]/.test(s)) return "cm";
  if (/\(\s*m\s*\)|\[\s*m\s*\]|\bmetre|\bmeter/.test(s)) return "m";
  if (/inch|\(\s*in\s*\)|\[\s*in\s*\]|\bin\.\b/.test(s)) return "in";
  if (/\bfeet\b|\bfoot\b|\(\s*ft\s*\)|\[\s*ft\s*\]/.test(s)) return "ft";
  return null;
}

function parseLengthToMm(v, unit) {
  if (v == null || v === "") return null;
  const conv = n => !isFinite(n) ? null
    : unit === "m" ? n * 1000 : unit === "cm" ? n * 10
    : unit === "in" ? n * 25.4 : unit === "ft" ? n * 304.8 : n;
  if (typeof v === "number") return conv(v);
  const s = String(v).trim().replace(/[\u2032\u2019]/g, "'").replace(/[\u2033\u201d]/g, '"');
  if (!s) return null;
  // Imperial when the cell carries a foot/inch mark or a fraction.
  if (/['"]/.test(s) || /\d\s*\/\s*\d/.test(s)) {
    let feet = 0, inches = 0, rest = s.replace(/,/g, "");
    const fm = rest.match(/^\s*(\d+(?:\.\d+)?)\s*'/);
    if (fm) { feet = parseFloat(fm[1]); rest = rest.slice(fm[0].length); }
    rest = rest.replace(/^[\s\-\u2013\u2014]+/, "").replace(/"/g, " ").trim();
    const whole = rest.match(/^(\d+(?:\.\d+)?)(?![\d.])(?!\s*\/)/);
    if (whole) { inches += parseFloat(whole[1]); rest = rest.slice(whole[0].length).trim(); }
    const frac = rest.match(/^(\d+)\s*\/\s*(\d+)/);
    if (frac && +frac[2] !== 0) inches += (+frac[1]) / (+frac[2]);
    const mm = feet * 304.8 + inches * 25.4;
    return mm > 0 ? mm : null;
  }
  // Metric — mirrors the existing thousands/decimal-comma handling exactly.
  let t = s.replace(/\s/g, "");
  if (t.includes(",") && t.includes(".")) {
    if (t.lastIndexOf(",") > t.lastIndexOf(".")) t = t.replace(/\./g, "").replace(",", ".");
    else t = t.replace(/,/g, "");
  } else if (t.includes(",")) {
    t = /,\d{1,2}$/.test(t) ? t.replace(",", ".") : t.replace(/,/g, "");
  }
  const n = parseFloat(t);
  return isFinite(n) ? conv(n) : null;
}

/* ╔══════════════════════════════════════════════════════════════════════════╗
   ║  ARABIC EXCEL INPUT — isolated add-on                                     ║
   ╚══════════════════════════════════════════════════════════════════════════╝
   Reads Arabic material lists (Saudi and Egyptian style) the same way as
   English ones, whatever interface language is selected. It runs once on the
   raw sheet rows BEFORE the existing Excel readers and only rewrites cells that
   contain Arabic script or Arabic digits:

     digits    ٠١٢٣٤٥٦٧٨٩ / ۰۱۲۳ · ٫ decimal · ٬ thousands   →  0-9 . ,
     headers   المقطع/القطاع/البروفيل → Profile · البيان/الوصف → Description
               الطول (مم/م/سم) → Length (mm/m/cm) · العدد/عدد القطع → Qty
               الكمية → Quantity (Weight if طن/كجم) · الرتبة/الدرجة → Grade
               العرض → Width · السماكة/السمك/التخانة → Thickness · الوزن → Weight
               الطول الكلي → Total run · طول السيخ/العود → Stock bar (never Length)
               الإجمالي/المجموع/الجملة → TOTAL
     names     زاوية 50×50×5 → L 50x50x5 · مربع/علبة 100×100×4 → SHS 100x100x4
               مستطيل 150×100×5 → RHS … · ماسورة 114.3×4 → CHS 114.3x4
               كمر IPE 30 / مجرى UPN 16 → IPE 300 / UPN 160 (Egyptian sizes in cm)
               كمر عادي 24 → IPN 240 · كمر انجليزي خفيف/ثقيل 20 → HEA/HEB 200
               300 IPE → IPE 300 (RTL typing) · بليت/صاج/لوح → PL … · شريحة/خوصة → FL …
     units     6 م · 6000 مم · 600 سم inside the length/width/thickness columns;
               a "الطول" column with no unit whose values are all ≤ 30 is read in metres
   A name is only rewritten when the result is a real library section (or a
   plate/flat); anything ambiguous — e.g. "كمر 30" with no type, "مجرى 10" —
   is left exactly as written. A sheet with no Arabic at all is returned as the
   very same array, so English files are read exactly as before. Any error →
   the original rows are used unchanged.
──────────────────────────────────────────────────────────────────────────── */
const ARABIC_INPUT_ENABLED = true;   // ◄ Arabic Excel reading on/off

const AR_SCRIPT = /[؀-ٟٮ-ۯۺ-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]/;
const AR_NUMERIC = /[٠-٩۰-۹٫٬]/;

function arDigits(s) {
  return String(s)
    .replace(/[٠-٩]/g, d => String(d.charCodeAt(0) - 0x0660))
    .replace(/[۰-۹]/g, d => String(d.charCodeAt(0) - 0x06F0))
    .replace(/٫/g, ".").replace(/[٬،]/g, ",")
    .replace(/[‎‏‪-‮⁦-⁩]/g, "");            // bidi marks
}
// Normalised Arabic for matching: no diacritics/tatweel, one alef, ى→ي, ة→ه.
function arKey(s) {
  return arDigits(s).replace(/[ً-ٰٟـ]/g, "")
    .replace(/[أإآٱ]/g, "ا").replace(/ى/g, "ي").replace(/ة/g, "ه").replace(/ؤ/g, "و").replace(/ئ/g, "ي")
    .toLowerCase().replace(/\s+/g, " ").trim();
}

const AR_UNIT_TOK = {
  "مم": "mm", "ملم": "mm", "مليمتر": "mm", "ملليمتر": "mm", "مللي": "mm", "ملي": "mm", "بالمم": "mm", "بالملم": "mm", "بالمليمتر": "mm", "mm": "mm",
  "سم": "cm", "سنتيمتر": "cm", "سنتي": "cm", "بالسم": "cm", "بالسنتيمتر": "cm", "cm": "cm",
  "م": "m", "م.": "m", "متر": "m", "امتار": "m", "بالمتر": "m", "م.ط": "m", "مط": "m", "م/ط": "m", "m": "m",
};
const AR_SKIP_TOK = new Set(["ط", "طولي", "طول.", "بال"]);
const AR_WEIGHT_TOK = /^(?:طن|كجم|كغ|كجم\/م|كيلو|كيلوجرام|كيلوغرام|kg|ton|t)$/;

const AR_HEAD = [
  [/^(?:ال)?(?:طول|اطوال) (?:ال)?(?:سيخ|اسياخ|عود|بار|خام|تجاري|قياسي|مخزون|توريد)$/, "Stock bar"],
  [/^(?:ال)?(?:طول|اطوال) (?:ال)?(?:كلي|اجمالي)$|^مجموع (?:ال)?اطوال$|^(?:ال)?اطوال (?:ال)?كليه$/, "Total run"],
  [/^(?:ال)?(?:طول|اطوال)(?: (?:ال)?(?:قطعه|قطع|عنصر|جزء|قص|صافي|مطلوب|فعلي))?$/, "Length"],
  [/^(?:ال)?عرض(?: (?:ال)?(?:قطعه|لوح|صاج))?$/, "Width"],
  [/^(?:ال)?(?:سماكه|سمك|سماك|تخانه|تخن)(?: (?:ال)?(?:لوح|صاج|قطعه|جدار))?$/, "Thickness"],
  [/^(?:ال)?(?:عدد|عدد (?:ال)?(?:قطع|قطعه|اسياخ|عناصر|وحدات|اعواد|كلي)|قطع|قطعه|ع)$/, "Qty", "count"],
  [/^(?:ال)?كميه(?: (?:ال)?(?:قطع|مطلوبه))?$/, "Quantity"],
  [/^(?:ال)?(?:مقطع|قطاع|بروفيل|بروفايل|مقاطع|قطاعات|مقاس|قياس)(?: (?:ال)?(?:حديد|معدني|مقطع|قطاع))?$|^(?:نوع|اسم) (?:ال)?(?:مقطع|قطاع|بروفيل)$/, "Profile"],
  [/^(?:ال)?(?:بيان|وصف|توصيف|صنف|مواصفات|تفاصيل)(?: (?:ال)?(?:اعمال|مواد|صنف|قطعه|عنصر))?$|^اسم (?:ال)?(?:صنف|ماده)$/, "Description"],
  [/^(?:ال)?(?:رتبه|درجه|خامه|ماده|جوده)(?: (?:ال)?(?:حديد|صلب|خامه))?$|^نوع (?:ال)?(?:صلب|خامه)$/, "Grade"],
  [/^(?:م|ن|مسلسل|تسلسل|رقم|(?:ال)?رقم(?: (?:ال)?(?:قطعه|عنصر|بند|جزء|رسم))?|(?:ال)?(?:علامه|ماركه|رمز|كود|بند))$/, "Mark"],
  [/^(?:ال)?وحده$/, "Unit"],
  [/^(?:ال)?ملاحظات?$/, "Notes"],
];

// Arabic header cell → { name, unit, count } or null (bilingual cells are left alone).
function arHeaderCell(c) {
  if (typeof c !== "string" || !AR_SCRIPT.test(c)) return null;
  const k = arKey(c).replace(/[():\[\]{}\\\-–—_|\/]/g, " ").replace(/\s+/g, " ").trim();
  if (!k || k.length > 40) return null;
  let unit = null, weight = false; const base = [];
  for (const t of k.split(" ")) {
    if (AR_WEIGHT_TOK.test(t)) { weight = true; continue; }
    if (base.length && AR_UNIT_TOK[t]) { unit = unit || AR_UNIT_TOK[t]; continue; }
    if (base.length && AR_SKIP_TOK.has(t)) continue;
    base.push(t);
  }
  const b = base.join(" ");
  if (!b || /[a-z]/.test(b)) return null;
  if (weight || /وزن/.test(b)) return { name: "Weight", unit: null };
  for (const [rx, name, flag] of AR_HEAD) if (rx.test(b)) return { name, unit, count: flag === "count" };
  return null;
}
function arUnitOnly(c) {
  if (typeof c !== "string" || !AR_SCRIPT.test(c)) return null;
  const k = arKey(c).replace(/[()\[\]\s]/g, "");
  return AR_UNIT_TOK[k] || null;
}
function arIsTotal(v) {
  const k = arKey(v).replace(/[:.\-–—]/g, " ").replace(/\s+/g, " ").trim();
  return /^(?:ال)?(?:اجمالي|مجموع|جمله)(?: (?:ال)?(?:كلي|عام|نهائي))?$/.test(k);
}
// "6 م" / "6000 مم" / "600 سم" in a length column → number in the header's unit.
function arUnitValue(v, headerUnit) {
  const m = arKey(v).match(/^(-?\d+(?:[.,]\d+)?)\s*(مم|ملم|مليمتر|سم|م|م\.|متر|امتار|م\.?ط|م\/ط)$/);
  if (!m) return null;
  const n = parseFloat(m[1].replace(",", ".")); if (!isFinite(n)) return null;
  const f = { mm: 1, cm: 10, m: 1000 }, mm = n * f[AR_UNIT_TOK[m[2]] || (m[2].startsWith("م") ? "m" : "mm")];
  const out = mm / f[headerUnit || "mm"];
  return Math.round(out * 1e6) / 1e6;
}

/* Arabic section / plate name → designation the library understands, or null. */
const AR_CM_FAMS = new Set(["IPE", "IPN", "HEA", "HEB", "HEM", "UPN", "UPE"]);
function arFamSize(fam, n) {                                    // Egyptian cm sizes: IPE 30 → IPE 300
  const a = `${fam} ${n}`;
  if (findSection(a)) return a;
  if (AR_CM_FAMS.has(fam) && Number.isInteger(n) && n > 0 && n < 110) { const b = `${fam} ${n * 10}`; if (findSection(b)) return b; }
  return null;
}
function arSectionName(raw) {
  const s0 = arKey(raw);
  if (!AR_SCRIPT.test(s0)) return null;
  const s = s0.replace(/[×✕✖*∗]/g, "x").replace(/(\d)\s*x\s*(?=\d)/g, "$1x")
    .replace(/(\d)\s*(?:مم|ملم|mm)(?=[\sx]|$)/g, "$1 ");
  const NUM = "(\\d+(?:\\.\\d+)?)";
  const FAM = { inp: "IPN", unp: "UPN" };
  const up = f => FAM[f.toLowerCase()] || f.toUpperCase();
  let m;
  // 1) a Latin designation inside Arabic text: "كمر IPE 300", "جسر HEA200", "254x254x73 UC"
  if ((m = s.match(/(\d+x\d+x\d+(?:\.\d+)?)\s*(ubp|ub|uc|pfc)(?![a-z])/i))) { const c = `${m[1]} ${m[2].toUpperCase()}`; if (findSection(c)) return c; }
  if ((m = s.match(/(?:^|[^a-z])(ipe|ipn|inp|hea|heb|hem|upn|unp|upe|pfc|ubp|ub|uc|hw|hm|hn|hp|mc|shs|rhs|chs|hss|pipe|w|c|l)\s*-?\s*(\d+(?:\.\d+)?(?:x\d+(?:\.\d+)?)*)/i))) {
    const fam = up(m[1]), dims = m[2];
    if (/^\d+$/.test(dims)) { const c = arFamSize(fam, +dims); if (c) return c; }
    const c = fam.length === 1 || fam === "MC" || fam === "HP" ? `${fam}${dims}` : `${fam} ${dims}`;
    if (findSection(c)) return c;
  }
  if ((m = s.match(/(?:^|[^\d.])(\d+)\s*(ipe|ipn|inp|hea|heb|hem|upn|unp|upe)(?![a-z])/i))) { const c = arFamSize(up(m[2]), +m[1]); if (c) return c; }
  // 2) Arabic words only
  const grp = (s.match(new RegExp(`${NUM}(?:x${NUM})+`)) || [""])[0];
  const d = grp ? grp.split("x").map(Number) : [];
  const beam = /(?:كمر|كمره|كمرات|جسر|جسور|عمود|اعمده)/.test(s);
  const one = s.match(/(?:^|\s)(\d+)(?:\s|$)/);
  if (beam && one && !grp) {
    const fam = /عادي|عاده/.test(s) ? "IPN" : /(?:انجليزي|عريض) خفيف/.test(s) ? "HEA" : /(?:انجليزي|عريض) ثقيل/.test(s) ? "HEB" : null;
    if (fam) return arFamSize(fam, +one[1]);
    return null;                                                  // "كمر 30" — type unknown, not guessed
  }
  if (/(?:شريحه|شرائح|خوصه|خوص|فلات|مبطط|مبسط)/.test(s)) return d.length === 2 ? `FL ${d[0]}x${d[1]}` : null;
  if (/(?:صاج|بليت|بليته|لوح|الواح|صفيحه|صفائح)/.test(s)) {
    if (d.length >= 2) return `PL ${d.join("x")}`;
    const t = s.match(/(?:تخانه|سمك|سماكه)\s*(\d+(?:\.\d+)?)/) || s.match(/(?:^|\s)(\d+(?:\.\d+)?)(?:\s|$)/);
    return t ? `PL ${t[1]}` : null;
  }
  const valid = c => (findSection(c) ? c : null);
  if (/مربع/.test(s)) {
    if (d.length === 3 && d[0] === d[1] && d[2] < d[0] / 2) return valid(`SHS ${d[0]}x${d[1]}x${d[2]}`);
    if (d.length === 2 && d[1] < d[0] / 2) return valid(`SHS ${d[0]}x${d[0]}x${d[1]}`);
    return null;
  }
  if (/مستطيل/.test(s)) return d.length === 3 && d[2] < Math.min(d[0], d[1]) / 2 ? valid(`RHS ${d[0]}x${d[1]}x${d[2]}`) : null;
  if (/(?:ماسوره|مواسير|انبوب|انابيب|بايب|دائري)/.test(s)) {
    if (/(?:بوصه|انش|انچ|")/.test(s)) return null;                 // nominal inch pipe — not guessed
    return d.length === 2 && d[1] < d[0] / 2 ? valid(`CHS ${d[0]}x${d[1]}`) : null;
  }
  if (/(?:علبه|علب|تيوب|مجوف)/.test(s)) return d.length === 3 && d[2] < Math.min(d[0], d[1]) / 2 ? valid(`${d[0] === d[1] ? "SHS" : "RHS"} ${d[0]}x${d[1]}x${d[2]}`) : null;
  if (/(?:زاويه|زوايا|زاوي|انجل)/.test(s)) {
    const [a, b, t] = d.length === 3 ? d : d.length === 2 ? [d[0], d[0], d[1]] : [];
    return a >= 15 && b >= 15 && t > 0 && t < Math.min(a, b) / 2 ? valid(`L ${a}x${b}x${t}`) : null;
  }
  return null;
}

// Runs on the raw rows of one sheet (array of arrays) before the Excel readers.
function arPrepSheet(aoa) {
  if (!ARABIC_INPUT_ENABLED || !Array.isArray(aoa)) return aoa;
  let has = false;
  outer: for (const r of aoa) if (Array.isArray(r)) for (const c of r) if (typeof c === "string" && (AR_SCRIPT.test(c) || AR_NUMERIC.test(c))) { has = true; break outer; }
  if (!has) return aoa;                                           // English sheet: untouched
  try { return arPrepSheetInner(aoa); }
  catch (e) { try { console.warn("[arabic-input add-on]", e); } catch { /* no console */ } return aoa; }
}
function arPrepSheetInner(aoa) {
  const out = aoa.map(r => (Array.isArray(r) ? r.slice() : r));
  const EN_HEAD = /profile|section|designation|length|qty|quantity|grade|width|thick|description|weight|mark/i;
  const cols = [];                                                // length/width/thickness columns
  const headerRows = new Set();
  const AR_PAREN_UNIT = { "مم": "mm", "ملم": "mm", "سم": "cm", "م": "m", "متر": "m", "م.ط": "m" };
  for (let i = 0; i < out.length; i++) {
    const r = out[i]; if (!Array.isArray(r)) continue;
    // bilingual/English header with an Arabic unit, e.g. "Length (م)" → "Length (m)"
    r.forEach((c, j) => { if (typeof c === "string" && EN_HEAD.test(c) && AR_SCRIPT.test(c)) r[j] = c.replace(/\(\s*(مم|ملم|سم|م|متر|م\.ط)\s*\)/g, (x, u) => `(${AR_PAREN_UNIT[u]})`); });
    const tr = r.map(arHeaderCell);
    const arHits = tr.filter(x => x && x.name !== "Mark").length;
    const enHits = r.filter(c => typeof c === "string" && !AR_SCRIPT.test(c) && EN_HEAD.test(c)).length;
    if (!arHits || arHits + enHits < 2) continue;
    headerRows.add(i);
    const hasCount = tr.some(x => x && x.count);
    tr.forEach((x, j) => {
      if (!x) return;
      const name = x.name === "Quantity" && hasCount ? "Amount" : x.name;
      out[i][j] = name + (x.unit ? ` (${x.unit})` : "");
      if (name === "Length" || name === "Width" || name === "Thickness") cols.push({ row: i, col: j, unit: x.unit || null });
    });
    const nx = out[i + 1];                                          // units on a second header row
    if (Array.isArray(nx)) cols.filter(h => h.row === i).forEach(h => { const u = arUnitOnly(nx[h.col]); if (u) { nx[h.col] = `(${u})`; if (!h.unit) h.unit = u; headerRows.add(i + 1); } });
  }
  // "الطول" with no unit and every value ≤ 30 → the list is in metres (no cut piece is 30 mm long)
  for (const h of cols) {
    if (h.unit || !/^Length/.test(String(out[h.row][h.col]))) continue;
    let n = 0, max = 0, ok = true;
    for (let i = h.row + 1; i < out.length && ok; i++) {
      if (headerRows.has(i)) { if (cols.some(x => x.row === i)) break; continue; }
      const c = Array.isArray(out[i]) ? out[i][h.col] : null;
      if (c == null || c === "") continue;
      const v = typeof c === "number" ? c : parseFloat(arDigits(c).replace(",", "."));
      if (!isFinite(v) || !/^\s*[\d.,٠-٩٫]+\s*$/.test(String(c))) continue;
      if (v <= 0) continue;
      n++; if (v > max) max = v;
    }
    if (n && max <= 30) { h.unit = "m"; out[h.row][h.col] = "Length (m)"; }
  }
  for (let i = 0; i < out.length; i++) {
    const r = out[i]; if (!Array.isArray(r) || headerRows.has(i)) continue;
    for (let j = 0; j < r.length; j++) {
      const c = r[j];
      if (typeof c === "string" && !AR_SCRIPT.test(c)) {                  // RTL typing: "300 IPE" → "IPE 300"
        const rv = c.match(/^\s*(\d+)\s*(ipe|ipn|hea|heb|hem|upn|upe)\s*$/i);
        if (rv) { const f = arFamSize(rv[2].toUpperCase(), +rv[1]); if (f) { r[j] = f; continue; } }
      }
      if (typeof c !== "string" || !(AR_SCRIPT.test(c) || AR_NUMERIC.test(c))) continue;
      const v = arDigits(c);
      if (arIsTotal(v)) { r[j] = "TOTAL"; continue; }
      let h = null; for (const x of cols) if (x.col === j && x.row < i && (!h || x.row > h.row)) h = x;
      if (h) { const n = arUnitValue(v, h.unit); if (n != null) { r[j] = n; continue; } }
      const nm = AR_SCRIPT.test(v) ? arSectionName(v) : null;
      r[j] = nm || v;
    }
  }
  return out;
}
/* ╚══ end of ARABIC EXCEL INPUT add-on ═══════════════════════════════════════╝ */

/* ╔══════════════════════════════════════════════════════════════════════════╗
   ║  RUSSIAN · CHINESE · SPANISH EXCEL INPUT — isolated add-on                ║
   ╚══════════════════════════════════════════════════════════════════════════╝
   Runs after the Arabic layer, on the raw rows of one sheet, before the
   existing Excel readers — whatever interface language is selected. It only
   acts on a sheet whose header row is written in Russian, Chinese or Spanish,
   or that contains Cyrillic / Chinese text; any other sheet is returned as the
   very same array, so English files are read exactly as before. Any error →
   the original rows are used unchanged.

     headers   Профиль/Сечение · Наименование · Длина, мм · Кол-во, шт · Марка стали ·
               Масса ед./общая, кг · Масса 1 м · Ширина · Толщина · Поз./Марка · Итого
               规格/截面/型号 · 名称 · 长度(mm) · 数量/件数 · 材质/牌号 · 单重/总重 · 米重 ·
               宽度 · 厚度 · 编号/零件号 · 合计
               Perfil/Sección · Descripción · Longitud/Largo · Cantidad/Uds · Calidad/Material ·
               Peso unitario/total · Peso por metro · Ancho · Espesor · Marca/Pos. · Total
               (two-row headers — «Масса, кг» over «ед. | общая» — are joined first)
             → the English names the readers already understand
     units     «6 м», «6000 мм», «600 см», «6米», «6 m» inside length, width and
               thickness columns; a length column with no unit whose values are all
               ≤ 30 is read in metres; «12 000», «6,5» and — Spanish — «6.000» → numbers
     plates    «—10х200», «Лист 10», «Полоса 10х100», «钢板-10×200», «-10×200», «扁钢»,
               «Chapa 10x200», «Pletina 100x10», «δ=10» → PL10x200 · FL100x10 · PL10
     sections  native names stay as written (the section library reads them);
               «20B1» in a Russian sheet → 20Б1, and a standard written in its own
               column («ГОСТ 26020-83») is joined to the profile when that gives a
               library row.
──────────────────────────────────────────────────────────────────────────── */
const INTL_INPUT_ENABLED = true;   // ◄ Russian / Chinese / Spanish Excel reading on/off

const INTL_UNIT = {
  "мм": "mm", "mm": "mm", "毫米": "mm", "м": "m", "m": "m", "米": "m", "см": "cm", "cm": "cm", "厘米": "cm",
  "кг": "kg", "kg": "kg", "千克": "kg", "公斤": "kg", "т": "t", "тн": "t", "t": "t", "ton": "t", "吨": "t",
  "кг/м": "kg/m", "кг/пм": "kg/m", "кг/п.м": "kg/m", "kg/m": "kg/m", "kg/ml": "kg/m",
  "шт": "pcs", "pcs": "pcs", "件": "pcs", "根": "pcs", "个": "pcs", "块": "pcs", "张": "pcs", "ud": "pcs", "uds": "pcs",
};
// Russian headers are matched word by word (\b does not work next to Cyrillic letters)
function intlRuHead(k) {
  const w = k.split(" "), w0 = w[0] || "", rest = w.slice(1).join(" "), has = re => w.slice(1).some(x => re.test(x));
  if (w0 === "масса" || w0 === "вес") {
    if (/(?:^|\s)1\s?(?:м|пм|п м|п\.м)(?:\s|$)/.test(rest) || has(/^(?:погонн|метр|пм$)/)) return "kgm";
    if (has(/^(?:ед|единиц|одн|детали$|штуки$|элемента$)/) || /(?:^|\s)1\s?(?:шт|дет)/.test(rest)) return "w1";
    if (has(/^(?:общ|всех$|итог|всего$|сумм)/)) return "wAll";
    return "w";
  }
  if (/^(?:общ|сумм|итог)/.test(w0) && (w[1] === "масса" || w[1] === "вес")) return "wAll";
  if (((w0 === "марка" || w0 === "класс") && w[1] === "стали") || w0 === "сталь" || w0 === "материал" || (w0 === "марка" && w[1] === "материала")) return "grade";
  if (["профиль", "сечение", "сортамент", "типоразмер", "прокат"].includes(w0) || (w0 === "вид" && w[1] === "проката") || (w0 === "наименование" && /^(?:профиля|проката|сечения)$/.test(w[1] || ""))) return "profile";
  if (["наименование", "описание", "название"].includes(w0)) return "desc";
  if ((/^(?:общая|суммарная)$/.test(w0) && w[1] === "длина") || (w0 === "длина" && /^(?:общая|всего|итого|суммарная)$/.test(w[1] || ""))) return "run";
  if ((w0 === "длина" && /^(?:хлыста|проката|поставки)$/.test(w[1] || "")) || (w0 === "мерная" && w[1] === "длина") || w0 === "хлыст") return "stock";
  if (w0 === "длина" || w0 === "дл" || k === "l") return "len";
  if (w0 === "ширина" || k === "b") return "wid";
  if (w0 === "толщина" || k === "t") return "thk";
  if (/^(?:кол-во|кол|количество|число|шт)$/.test(w0)) return /на (?:1 )?марку|в марке|на 1 шт/.test(rest) ? "qtyPer" : "qty";
  if (/^(?:марка|поз|позиция|№|№ п п|номер|обозначение|код|шифр|n)$/.test(k) || ((w0 === "марка" || w0 === "номер") && /^(?:детали|элемента|сборки|конструкции)$/.test(w[1] || ""))) return "mark";
  if (/^примечани/.test(w0)) return "notes";
  if (w0 === "ед" && w[1] === "изм") return "unit";
  return null;
}
const INTL_HEAD_RULES = {
  zh: [
    [/^(?:米重|每米重|每米重量|线重|理论米重)/, "kgm"],
    [/^(?:单重|单件重|单根重|单件重量|单重量|每件重)/, "w1"],
    [/^(?:总重|总重量|合计重量|共重|合重|重量合计|总计重量)/, "wAll"],
    [/^(?:重量|质量|理论重量|重)$/, "w"],
    [/^(?:材质|材料|钢号|牌号|钢材牌号|钢种|钢材材质)$/, "grade"],
    [/^(?:规格|截面|型号|规格型号|截面规格|型材规格|构件规格|截面尺寸|断面|型钢规格|规格尺寸|截面型号)$/, "profile"],
    [/^(?:名称|材料名称|构件名称|零件名称|品名|描述|零件描述)$/, "desc"],
    [/^(?:总长|总长度|合计长度|长度合计)$/, "run"],
    [/^(?:定尺|定尺长度|原材长度|原料长度)$/, "stock"],
    [/^(?:长度|长|单长|下料长度|零件长度|构件长度|单件长度)$/, "len"],
    [/^(?:宽度|宽|板宽)$/, "wid"],
    [/^(?:厚度|厚|板厚)$/, "thk"],
    [/^(?:单构件数量|每构件数量|单件数量)$/, "qtyPer"],
    [/^(?:数量|件数|根数|块数|数目|总数量|总数|张数)$/, "qty"],
    [/^(?:编号|零件号|构件号|序号|件号|零件编号|构件编号|图号)$/, "mark"],
    [/^备注$/, "notes"],
    [/^单位$/, "unit"],
  ],
  es: [
    [/^peso\b.*(?:por\s*metro|\/\s*m\b|lineal|metro)|^kg\s*\/\s*m\b/, "kgm"],
    [/^peso\b.*(?:unit|\bud\b|unidad|por\s*pieza|\bpieza\b)/, "w1"],
    [/^peso\b.*\btotal\b|^total\s*peso/, "wAll"],
    [/^peso\b/, "w"],
    [/^(?:calidad|grado|tipo\s*de\s*acero|acero|material)\b/, "grade"],
    [/^(?:perfil|secci[oó]n|designaci[oó]n|tipo\s*de\s*perfil)\b/, "profile"],
    [/^(?:descripci[oó]n|denominaci[oó]n|concepto)\b/, "desc"],
    [/^(?:longitud|largo|metros)\s*total|^total\s*(?:longitud|metros)/, "run"],
    [/^(?:longitud|largo)\s*(?:comercial|de\s*barra|de\s*suministro)|^barra\s*comercial/, "stock"],
    [/^(?:longitud|largo|long)\b/, "len"],
    [/^(?:ancho|anchura)\b/, "wid"],
    [/^(?:espesor|grosor)\b/, "thk"],
    [/^(?:cantidad|cant|uds|unidades|piezas|n[ºo°]?\s*(?:de\s*)?piezas)\b.*(?:por\s*conjunto|por\s*marca)/, "qtyPer"],
    [/^(?:cantidad|cant|uds|unidades|piezas|n[ºo°]?\s*(?:de\s*)?piezas|n[uú]mero\s*de\s*piezas)\b/, "qty"],
    [/^(?:marca|pos|posici[oó]n|ref|referencia|n[ºo°]|[ií]tem|c[oó]digo)$/, "mark"],
    [/^(?:observaciones|notas|comentarios)\b/, "notes"],
    [/^(?:unidad|ud)$/, "unit"],
  ],
};
const INTL_ES_ONLY = /^(?:perfil|secci|designaci|longitud|largo|cantidad|cant|uds|unidades|piezas|espesor|grosor|ancho|anchura|descripci|denominaci|peso|calidad|posici|observaciones|marca)/;
const INTL_EN_HEAD = /profile|section|designation|length|qty|quantity|grade|width|thick|description|weight|mark|material/i;
const INTL_TOTAL = /^(?:итого|всего|итого по .*|всего по .*|合计|总计|小计|共计|suma|suma total|total general)$/;

function intlHeadKey(c) {
  return String(c == null ? "" : c).normalize("NFKC").toLowerCase().replace(/ё/g, "е")
    .replace(/[‐‑–—−]/g, "-").replace(/\s+/g, " ").trim();
}
// header cell → { k: words without units, unit }
function intlHeadSplit(c) {
  let k = intlHeadKey(c), unit = null;
  const take = u => { const x = INTL_UNIT[String(u).replace(/\s+/g, "").replace(/\.$/, "")]; if (x && !unit) unit = x; return !!x; };
  k = k.replace(/[(\[【]\s*([^)\]】]{1,8}?)\s*[)\]】]/g, (m, u) => (take(u) ? " " : " " + u + " "));
  k = k.replace(/,\s*([^,]{1,8})$/, (m, u) => (take(u.trim()) ? "" : m));
  const tm = k.match(/^(.*?[^\s,])\s*(мм|см|кг\/м|кг|шт\.?|mm|cm|kg\/m|kg|米|毫米|厘米|千克|公斤|吨)$/);
  if (tm && take(tm[2])) k = tm[1];
  k = k.replace(/[.,:;·\\_|*\/]+/g, " ").replace(/\s+/g, " ").trim();
  return { k, unit };
}
function intlHeadClass(c, lang) {
  if (typeof c !== "string" || !c.trim() || c.length > 60) return null;
  if (lang !== "es" && /[a-z]{3,}/i.test(c) && INTL_EN_HEAD.test(c)) return null;       // English / bilingual header: left alone
  const { k, unit } = intlHeadSplit(c);
  if (!k) return null;
  if (lang === "ru") { const cls = intlRuHead(k); return cls ? { cls, unit, es: false } : null; }
  for (const [rx, cls] of INTL_HEAD_RULES[lang]) if (rx.test(k)) return { cls, unit, es: lang === "es" && INTL_ES_ONLY.test(k) };
  return null;
}
const INTL_HEAD_EN = { profile: "Profile", desc: "Description", len: "Length", run: "Total run", stock: "Stock bar", wid: "Width", thk: "Thickness",
  qty: "Qty", qtyPer: "Per assembly", grade: "Grade", w: "Weight", w1: "Unit weight", wAll: "Total weight", kgm: "Mass per m (kg/m)", mark: "Mark", notes: "Notes", unit: "Unit" };
function intlHeadName(h) {
  const base = INTL_HEAD_EN[h.cls];
  if (h.cls === "len" || h.cls === "wid" || h.cls === "thk") return h.unit && h.unit !== "pcs" && h.unit !== "kg" && h.unit !== "t" ? `${base} (${h.unit})` : base;
  if (h.cls === "w" || h.cls === "w1" || h.cls === "wAll") return h.unit === "t" ? `${base} (t)` : h.unit === "kg/m" ? INTL_HEAD_EN.kgm : `${base} (kg)`;
  return base;
}
const INTL_HEAD_CORE = new Set(["profile", "desc", "len", "wid", "thk", "qty"]);
function intlFindHeader(row, lang) {
  if (!Array.isArray(row)) return null;
  const cls = row.map(c => intlHeadClass(c, lang));
  const hits = cls.filter(Boolean);
  if (hits.length < 2 || !hits.some(h => INTL_HEAD_CORE.has(h.cls))) return null;
  if (lang === "es" && !hits.some(h => h.es)) return null;
  return cls;
}
const INTL_CYR_ANY = /[А-Яа-яЁё]/, INTL_CJK_ANY = /[㐀-鿿豈-﫿]/;
const INTL_ES_WORD = /\b(?:perfil|longitud|cantidad|espesor|ancho|descripci[oó]n|chapa|pletina|viga|tubo|angular)\b/i;

// numbers written with units / spaces / commas in a length, width or thickness column
function intlUnitValue(v, headUnit, lang) {
  if (typeof v !== "string") return null;
  const s = v.normalize("NFKC").replace(/\s+/g, " ").trim().toLowerCase();
  const m = s.match(/^(-?\d[\d\s.,]*?)\s*(мм|см|м|mm|cm|m|米|毫米|厘米)?\.?$/);
  if (!m) return null;
  const raw = m[1].replace(/\s/g, ""), u = m[2] ? INTL_UNIT[m[2]] : null;
  let t = raw;
  if (lang === "es" && /^\d{1,3}(?:\.\d{3})+$/.test(t) && (u || headUnit || "mm") !== "m") t = t.replace(/\./g, "");              // 6.000 = six thousand
  else if (t.includes(",") && t.includes(".")) t = t.lastIndexOf(",") > t.lastIndexOf(".") ? t.replace(/\./g, "").replace(",", ".") : t.replace(/,/g, "");
  else if (t.includes(",")) t = lang === "zh" && /^\d{1,3}(?:,\d{3})+$/.test(t) ? t.replace(/,/g, "") : t.replace(",", ".");   // ru/es: decimal comma
  if (!/^-?\d+(?:\.\d+)?$/.test(t)) return null;
  if (!u && t === raw && raw === s) return null;                  // a plain number: the reader reads it as before
  const n = parseFloat(t), f = { mm: 1, cm: 10, m: 1000 };
  const out = u ? n * f[u] / f[headUnit || "mm"] : n;
  return Math.round(out * 1e6) / 1e6;
}
/* plate / flat names in any of the three languages → PL t×w / FL a×b / PL t (or null) */
function intlPlateName(raw) {
  const s0 = String(raw == null ? "" : raw).normalize("NFKC").replace(/ /g, " ").trim();
  if (!s0) return null;
  const s = s0.toUpperCase().replace(/[хХxX×✕✖*∗]/g, "x").replace(/(\d),(\d)/g, "$1.$2")
    .replace(/(?:ГОСТ|GOST|ТУ|СТО|GB\/?T?|JIS|EN|UNE|ASTM)\s*[РP]?\s*[\d.\s-]*\d(?:-\d{2,4})?/g, " ")
    .replace(/(^|\s)(?:С\d{3}[А-ЯЁ]?(?:-\d)?|\d{2}Г2С|10ХСНД|СТ\s*\d\S*|Q\d{3}[A-E]?|S\d{3}[A-Z0-9]*|SS\d{3})(?=\s|$)/g, "$1")
    .replace(/\s+/g, " ").trim();
  let kind = null, body = s;
  let m;
  if ((m = s.match(/^(?:ЛИСТ(?:ОВОЙ ПРОКАТ)?|ПЛАСТИНА|ФАСОНКА|РЕБРО|НАКЛАДКА|ПРОКЛАДКА|ОПОРНАЯ ПЛИТА|ПЛИТА|КОСЫНКА|ДИАФРАГМА)(?![А-ЯЁ])\s*/))) { kind = "PL"; body = s.slice(m[0].length); }
  else if ((m = s.match(/^(?:ПОЛОСА(?: СТАЛЬНАЯ)?|ПОЛОСОВАЯ СТАЛЬ)(?![А-ЯЁ])\s*/))) { kind = "FL"; body = s.slice(m[0].length); }
  else if ((m = s.match(/^(?:钢板|花纹板|钢带|加劲板|加劲肋|连接板|节点板|垫板|肋板|板)\s*/))) { kind = "PL"; body = s.slice(m[0].length); }
  else if ((m = s.match(/^扁钢\s*/))) { kind = "FL"; body = s.slice(m[0].length); }
  else if ((m = s.match(/^(?:CHAPA|PLACA|PLANCHA|L[AÁ]MINA|PALASTRO)\b\s*/))) { kind = "PL"; body = s.slice(m[0].length); }
  else if ((m = s.match(/^(?:PLETINA|LLANTA|FLEJE|PLATINA|SOLERA)\b\s*/))) { kind = "FL"; body = s.slice(m[0].length); }
  else if (/^[-—–－]\s*\d/.test(s)) kind = "PL";
  else if (/^[Δδ]\s*=?\s*\d/.test(s0)) { kind = "PL"; body = s0.replace(/^[Δδ]\s*=?\s*/i, ""); }
  else return null;
  if (!/\d/.test(body)) return /^[\s.,:;-]*$/.test(body) ? { name: kind, t: null } : null;     // «Лист», «Chapa», «钢板» with the sizes in their own columns
  body = body.replace(/^(?:PL|FL)?\s*[-—–－]?\s*/, "").replace(/^(?:[TSΔδ]\s*=\s*)/i, "").replace(/\s*(?:MM|ММ)(?![A-ZА-ЯЁ])/g, "").trim();
  const d = (body.match(/^(\d+(?:\.\d+)?)(?:\s*x\s*(\d+(?:\.\d+)?))?(?:\s*x\s*(\d+(?:\.\d+)?))?/) || []).slice(1).filter(v => v != null).map(Number);
  if (!d.length || !(d[0] > 0)) return null;
  if (d.length === 1) return d[0] <= 200 ? { name: `${kind}${d[0]}`, t: d[0] } : null;
  const [a, b, c] = d;
  return { name: `${kind}${a}x${b}`, t: Math.min(a, b), len: c || null };
}
function intlStdCell(v) {
  const s = String(v == null ? "" : v).trim();
  return /^(?:ГОСТ|ТУ|СТО)\s*(?:Р\s*)?(?:АСЧМ\s*)?\d[\d.\-]*\d$/i.test(s) ? s.replace(/\s+/g, " ") : null;
}

function intlPrepSheet(aoa) {
  if (!INTL_INPUT_ENABLED || !Array.isArray(aoa)) return aoa;
  let script = null, esHint = false;
  outer: for (let i = 0; i < aoa.length; i++) {
    const r = aoa[i]; if (!Array.isArray(r)) continue;
    for (const c of r) {
      if (typeof c !== "string") continue;
      if (INTL_CYR_ANY.test(c)) { script = "ru"; break outer; }
      if (INTL_CJK_ANY.test(c)) { script = "zh"; break outer; }
      if (i < 40 && !esHint && INTL_ES_WORD.test(c)) esHint = true;
    }
  }
  if (!script && !esHint) return aoa;                             // English (or Arabic-prepared) sheet: untouched
  try { return intlPrepSheetInner(aoa, script); }
  catch (e) { try { console.warn("[intl-input add-on]", e); } catch { /* no console */ } return aoa; }
}
function intlPrepSheetInner(aoa, script) {
  const langs = script ? [script, "es"] : ["es"];
  let lang = null, hr = -1, cls = null;
  for (let i = 0; i < Math.min(aoa.length, 40) && hr < 0; i++) {
    for (const L of langs) { const h = intlFindHeader(aoa[i], L); if (h) { lang = L; hr = i; cls = h; break; } }
  }
  if (hr < 0) {                                                    // no header: act only on a real Russian / Chinese list
    if (!script) return aoa;
    let n = 0;
    for (const r of aoa) if (Array.isArray(r)) for (const c of r) if (typeof c === "string" && /\d/.test(c) && (INTL_CYR_ANY.test(c) || INTL_CJK_ANY.test(c))) n++;
    if (n < 2) return aoa;
    lang = script;
  }
  const out = aoa.map(r => (Array.isArray(r) ? r.slice() : r));
  const headerRows = new Set();
  const col = {};                                                  // class → first column with it
  let thkCol = null;
  if (hr >= 0) {
    headerRows.add(hr);
    // two-row header: «Масса, кг» over «ед. | общая», «Длина» over «мм», «Кол-во» over «на марку | всего»
    const top = out[hr], nx = out[hr + 1];
    if (Array.isArray(nx) && !nx.some(c => typeof c === "number") && nx.some(c => typeof c === "string" && c.trim())) {
      const plan = []; let ok = true;
      nx.forEach((c, j) => {
        if (!ok || typeof c !== "string" || !c.trim()) return;
        const sp = intlHeadSplit(c);
        let above = null; for (let k = j; k >= 0; k--) { const tc = top[k]; if (typeof tc === "string" && tc.trim()) { above = tc; break; } }
        if (!sp.k && sp.unit) { plan.push([j, cls[j] ? { ...cls[j], unit: cls[j].unit || sp.unit } : (above ? intlHeadClass(`${above} ${c}`, lang) : null)]); return; }
        const joined = above ? intlHeadClass(`${above} ${c}`, lang) : null, own = intlHeadClass(c, lang);
        if (joined || own) { plan.push([j, joined || own]); return; }
        if (/^(?:ед|общ|всего|итого|на марку|на 1 марку|单|总|unit|total)/.test(sp.k)) return;
        ok = false;                                                // a data row, not a second header row
      });
      if (ok && plan.length) { plan.forEach(([j, h]) => { if (h) cls[j] = h; }); headerRows.add(hr + 1); out[hr + 1] = nx.map(() => ""); }
    }
    if (!cls.some(h => h && h.cls === "qty")) { const j = cls.findIndex(h => h && h.cls === "qtyPer"); if (j >= 0) cls[j] = { ...cls[j], cls: "qty" }; }
    cls.forEach((h, j) => { if (!h) return; if (col[h.cls] == null) col[h.cls] = j; out[hr][j] = intlHeadName(h); });
    thkCol = col.thk != null ? col.thk : null;
  }
  // «Длина» with no unit and every value ≤ 30 → metres (no cut piece is 30 mm long)
  const units = {};
  for (const c of ["len", "wid", "thk"]) if (col[c] != null) units[c] = cls[col[c]].unit && /^(?:mm|cm|m)$/.test(cls[col[c]].unit) ? cls[col[c]].unit : null;
  if (col.len != null && !units.len) {
    let n = 0, max = 0;
    for (let i = hr + 1; i < out.length; i++) {
      if (headerRows.has(i) || !Array.isArray(out[i])) continue;
      const c = out[i][col.len];
      if (c == null || c === "" || (typeof c === "string" && /[a-zа-яё米]/i.test(c))) continue;
      const v = typeof c === "number" ? c : parseFloat(String(c).replace(/\s/g, "").replace(",", "."));
      if (!(v > 0)) continue;
      n++; if (v > max) max = v;
    }
    if (n && max <= 30) { units.len = "m"; out[hr][col.len] = "Length (m)"; }
  }
  const NOT_NAME = new Set(["len", "wid", "thk", "qty", "qtyPer", "w", "w1", "wAll", "kgm", "grade", "mark", "notes", "unit", "run", "stock"]);
  const nameCols = hr < 0 ? null
    : col.profile != null || col.desc != null ? [col.profile, col.desc].filter(j => j != null)
    : cls.map((h, j) => (h && NOT_NAME.has(h.cls) ? -1 : j)).filter(j => j >= 0);
  const isTotal = c => typeof c === "string" && INTL_TOTAL.test(intlHeadKey(c).replace(/[:.]+$/, ""));
  for (let i = hr + 1; i < out.length; i++) {
    const r = out[i]; if (!Array.isArray(r) || headerRows.has(i)) continue;
    if (r.some(isTotal)) { r.forEach((c, j) => { if (isTotal(c)) r[j] = "TOTAL"; }); continue; }
    for (const c of ["len", "wid", "thk"]) {                        // «6 м», «6000 мм», «12 000», «6,5»
      const j = col[c]; if (j == null) continue;
      const n = intlUnitValue(r[j], units[c], lang); if (n != null) r[j] = n;
    }
    if (lang === "es") for (const c of ["qty", "w", "w1", "wAll", "kgm"]) {   // «1.250» = one thousand two hundred and fifty
      const j = col[c]; if (j == null || typeof r[j] !== "string") continue;
      const t = r[j].trim(); if (/^\d{1,3}(?:\.\d{3})+$/.test(t)) r[j] = +t.replace(/\./g, "");
    }
    const js = nameCols || r.map((_, j) => j);
    let std = null;
    if (lang === "ru") for (let j = 0; j < r.length && !std; j++) if (!js.includes(j)) std = intlStdCell(r[j]);
    for (const j of js) {
      const v = r[j];
      if (typeof v !== "string" || !v.trim()) continue;
      // «Уголок» | «75х6», «角钢» | «75×6», «Tubo cuadrado» | «100x100x5»: the type word sits in its own column
      const dsc = col.desc != null && j === col.profile && typeof r[col.desc] === "string" ? r[col.desc].trim() : "";
      const typed = dsc && !/\d/.test(dsc) && dsc.length <= 40 ? `${dsc} ${v.trim()}` : null;
      const probe = typed || v;
      if (!/[A-Za-zА-Яа-яЁё㐀-鿿\[□口∠∟ØΦФδΔ—–－-]/.test(probe)) continue;
      const pl = intlPlateName(v) || (typed ? intlPlateName(typed) : null);
      if (pl) {
        r[j] = pl.name;
        if (pl.len && col.len != null && (r[col.len] === "" || r[col.len] == null)) r[col.len] = units.len === "m" ? pl.len / 1000 : pl.len;
        if (pl.t && !pl.name.includes("x") && hr >= 0) {             // «Лист 10»: the thickness goes to a thickness column
          if (thkCol == null) { thkCol = out[hr].length; out[hr][thkCol] = "Thickness (mm)"; }
          if (r[thkCol] === "" || r[thkCol] == null) r[thkCol] = pl.t;
        }
        continue;
      }
      if (!/\d/.test(probe)) continue;
      const own = findSection(v);
      const base = !own && typed && findSection(typed) ? typed : v.trim();
      if (std && !/(?:ГОСТ|ТУ|СТО)/i.test(base)) {                  // «20Б1» + «ГОСТ 26020-83» in its own column
        const b = findSection(`${base} ${std}`), a = base === v.trim() ? own : findSection(base);
        if (b && b.intl && (!a || a.name !== b.name)) { r[j] = `${base} ${std}`; continue; }
      }
      if (own) continue;                                             // the library already reads it as written
      if (base !== v.trim()) { r[j] = base; continue; }
      const c = soSafe(() => intlSectionName(v, lang));
      if (c && c.row) r[j] = c.row.name;
      else if (c && c.name && !c.plate && findSection(c.name)) r[j] = c.name;
      else if (typed) { const t2 = soSafe(() => intlSectionName(typed, lang)); if (t2 && t2.row) r[j] = t2.row.name; else if (t2 && t2.name && !t2.plate && findSection(t2.name)) r[j] = t2.name; }
    }
  }
  return out;
}
/* ╚══ end of RUSSIAN · CHINESE · SPANISH EXCEL INPUT add-on ═════════════════╝ */

const PART_COLORS = ["#F59E0B","#3B82F6","#10B981","#EF4444","#8B5CF6","#06B6D4","#F97316","#84CC16","#EC4899","#14B8A6","#6366F1","#D97706","#059669","#DC2626","#7C3AED","#0EA5E9","#65A30D","#DB2777"];
const fmtTon = (kg) => (kg / 1000).toFixed(kg / 1000 >= 100 ? 0 : 2);
const fmtMm = (mm) => (SO_LANG === "ar"
  ? (mm >= 1000 ? `${(mm / 1000).toFixed(3).replace(/\.?0+$/, "")} م` : `${Math.round(mm)} مم`)
  : SO_LANG === "ru" ? (mm >= 1000 ? `${(mm / 1000).toFixed(3).replace(/\.?0+$/, "")} м` : `${Math.round(mm)} мм`)
  : (mm >= 1000 ? `${(mm / 1000).toFixed(3).replace(/\.?0+$/, "")}m` : `${Math.round(mm)}mm`));
const fmtKg = (kg) => (SO_LANG === "ar"
  ? (kg >= 1000 ? `${(kg / 1000).toFixed(2)} طن` : `${kg.toFixed(1)} كجم`)
  : SO_LANG === "ru" ? (kg >= 1000 ? `${(kg / 1000).toFixed(2)} т` : `${kg.toFixed(1)} кг`)
  : (kg >= 1000 ? `${(kg / 1000).toFixed(2)} t` : `${kg.toFixed(1)} kg`));
const genId = () => Math.random().toString(36).slice(2, 8).toUpperCase();

/* ─── SECTION DATABASE (expanded in Stage 2) ─────────────────────────────── */
const STEEL_DB = [
  ["IPE 80","IPE",6.0],["IPE 100","IPE",8.1],["IPE 120","IPE",10.4],["IPE 140","IPE",12.9],["IPE 160","IPE",15.8],["IPE 180","IPE",18.8],["IPE 200","IPE",22.4],["IPE 220","IPE",26.2],["IPE 240","IPE",30.7],["IPE 270","IPE",36.1],["IPE 300","IPE",42.2],["IPE 330","IPE",49.1],["IPE 360","IPE",57.1],["IPE 400","IPE",66.3],["IPE 450","IPE",77.6],["IPE 500","IPE",90.7],["IPE 550","IPE",106.0],["IPE 600","IPE",122.0],
  ["HEA 100","HEA",16.7],["HEA 120","HEA",19.9],["HEA 140","HEA",24.7],["HEA 160","HEA",30.4],["HEA 180","HEA",35.5],["HEA 200","HEA",42.3],["HEA 220","HEA",50.5],["HEA 240","HEA",60.3],["HEA 260","HEA",68.2],["HEA 280","HEA",76.4],["HEA 300","HEA",88.3],["HEA 320","HEA",97.6],["HEA 340","HEA",105.0],["HEA 360","HEA",112.0],["HEA 400","HEA",125.0],["HEA 450","HEA",140.0],["HEA 500","HEA",155.0],["HEA 550","HEA",166.0],["HEA 600","HEA",178.0],["HEA 650","HEA",190.0],["HEA 700","HEA",204.0],["HEA 800","HEA",224.0],["HEA 900","HEA",252.0],["HEA 1000","HEA",272.0],
  ["HEB 100","HEB",20.4],["HEB 120","HEB",26.7],["HEB 140","HEB",33.7],["HEB 160","HEB",42.6],["HEB 180","HEB",51.2],["HEB 200","HEB",61.3],["HEB 220","HEB",71.5],["HEB 240","HEB",83.2],["HEB 260","HEB",93.0],["HEB 280","HEB",103.0],["HEB 300","HEB",117.0],["HEB 320","HEB",127.0],["HEB 340","HEB",134.0],["HEB 360","HEB",142.0],["HEB 400","HEB",155.0],["HEB 450","HEB",171.0],["HEB 500","HEB",187.0],["HEB 550","HEB",199.0],["HEB 600","HEB",212.0],["HEB 650","HEB",225.0],["HEB 700","HEB",241.0],["HEB 800","HEB",262.0],["HEB 900","HEB",291.0],["HEB 1000","HEB",314.0],
  ["127x76x13 UB","UB",13.0],["152x89x16 UB","UB",16.0],["178x102x19 UB","UB",19.0],["203x102x23 UB","UB",23.1],["203x133x25 UB","UB",25.1],["203x133x30 UB","UB",30.0],["254x102x22 UB","UB",22.0],["254x102x25 UB","UB",25.2],["254x102x28 UB","UB",28.3],["254x146x31 UB","UB",31.1],["254x146x37 UB","UB",37.0],["254x146x43 UB","UB",43.0],["305x102x25 UB","UB",24.8],["305x102x28 UB","UB",28.2],["305x102x33 UB","UB",32.8],["305x127x37 UB","UB",37.0],["305x127x42 UB","UB",41.9],["305x127x48 UB","UB",48.1],["305x165x40 UB","UB",40.3],["305x165x46 UB","UB",46.1],["305x165x54 UB","UB",54.0],["356x127x33 UB","UB",33.1],["356x127x39 UB","UB",39.1],["356x171x45 UB","UB",45.0],["356x171x51 UB","UB",51.0],["356x171x57 UB","UB",57.0],["356x171x67 UB","UB",67.1],["406x140x39 UB","UB",39.0],["406x140x46 UB","UB",46.0],["406x178x54 UB","UB",54.1],["406x178x60 UB","UB",60.1],["406x178x67 UB","UB",67.1],["406x178x74 UB","UB",74.2],["457x152x52 UB","UB",52.3],["457x152x60 UB","UB",59.8],["457x152x67 UB","UB",67.2],["457x152x74 UB","UB",74.2],["457x152x82 UB","UB",82.1],["457x191x67 UB","UB",67.1],["457x191x74 UB","UB",74.3],["457x191x82 UB","UB",82.0],["457x191x89 UB","UB",89.3],["457x191x98 UB","UB",98.3],["533x210x82 UB","UB",82.2],["533x210x92 UB","UB",92.1],["533x210x101 UB","UB",101.0],["533x210x109 UB","UB",109.0],["533x210x122 UB","UB",122.0],["610x229x101 UB","UB",101.0],["610x229x113 UB","UB",113.0],["610x229x125 UB","UB",125.0],["610x229x140 UB","UB",140.0],["610x305x149 UB","UB",149.0],["610x305x179 UB","UB",179.0],["610x305x238 UB","UB",238.0],["686x254x125 UB","UB",125.0],["686x254x140 UB","UB",140.0],["686x254x152 UB","UB",152.0],["686x254x170 UB","UB",170.0],["762x267x134 UB","UB",134.0],["762x267x147 UB","UB",147.0],["762x267x173 UB","UB",173.0],["762x267x197 UB","UB",197.0],["838x292x176 UB","UB",176.0],["838x292x194 UB","UB",194.0],["838x292x226 UB","UB",226.0],["914x305x201 UB","UB",201.0],["914x305x224 UB","UB",224.0],["914x305x253 UB","UB",253.0],["914x305x289 UB","UB",289.0],
  ["152x152x23 UC","UC",23.0],["152x152x30 UC","UC",30.0],["152x152x37 UC","UC",37.0],["152x152x44 UC","UC",44.0],["203x203x46 UC","UC",46.1],["203x203x52 UC","UC",52.0],["203x203x60 UC","UC",60.0],["203x203x71 UC","UC",71.0],["203x203x86 UC","UC",86.1],["203x203x100 UC","UC",100.0],["254x254x73 UC","UC",73.1],["254x254x89 UC","UC",88.9],["254x254x107 UC","UC",107.0],["254x254x132 UC","UC",132.0],["254x254x167 UC","UC",167.0],["305x305x97 UC","UC",96.9],["305x305x118 UC","UC",117.0],["305x305x137 UC","UC",137.0],["305x305x158 UC","UC",158.0],["305x305x198 UC","UC",198.0],["305x305x240 UC","UC",240.0],["305x305x283 UC","UC",283.0],["356x368x129 UC","UC",129.0],["356x368x153 UC","UC",153.0],["356x368x177 UC","UC",177.0],["356x368x202 UC","UC",202.0],["356x406x235 UC","UC",235.0],["356x406x287 UC","UC",287.0],["356x406x340 UC","UC",340.0],["356x406x393 UC","UC",393.0],["356x406x467 UC","UC",467.0],["356x406x551 UC","UC",551.0],["356x406x634 UC","UC",634.0],
  ["HW 100x100","JIS-HW",17.2],["HW 125x125","JIS-HW",23.8],["HW 150x150","JIS-HW",31.5],["HW 175x175","JIS-HW",40.4],["HW 200x200","JIS-HW",49.9],["HW 250x250","JIS-HW",72.4],["HW 300x300","JIS-HW",94.0],["HW 350x350","JIS-HW",137.0],["HW 400x400","JIS-HW",172.0],
  ["HM 150x100","JIS-HM",21.1],["HM 200x150","JIS-HM",30.6],["HM 250x175","JIS-HM",44.1],["HM 300x200","JIS-HM",56.8],["HM 350x250","JIS-HM",79.7],["HM 400x300","JIS-HM",107.0],["HM 450x300","JIS-HM",124.0],["HM 500x300","JIS-HM",128.0],
  ["HN 100x50","JIS-HN",9.3],["HN 125x60","JIS-HN",13.1],["HN 150x75","JIS-HN",14.0],["HN 175x90","JIS-HN",18.0],["HN 200x100","JIS-HN",20.9],["HN 250x125","JIS-HN",29.0],["HN 300x150","JIS-HN",36.7],["HN 350x175","JIS-HN",49.4],["HN 400x200","JIS-HN",65.4],["HN 450x200","JIS-HN",74.9],["HN 500x200","JIS-HN",88.5],["HN 600x200","JIS-HN",102.0],["HN 700x300","JIS-HN",166.0],["HN 800x300","JIS-HN",191.0],["HN 900x300","JIS-HN",213.0],
  ["RHS 50x30x3","RHS",3.45],["RHS 60x40x3","RHS",4.39],["RHS 80x40x4","RHS",6.71],["RHS 100x50x4","RHS",8.59],["RHS 100x60x5","RHS",11.30],["RHS 120x60x5","RHS",12.80],["RHS 120x80x5","RHS",14.40],["RHS 150x100x5","RHS",18.70],["RHS 160x80x6","RHS",20.70],["RHS 200x100x6","RHS",26.40],["RHS 200x120x8","RHS",37.30],["RHS 250x150x8","RHS",47.70],["RHS 300x200x8","RHS",60.30],
  ["SHS 40x40x3","SHS",3.41],["SHS 50x50x3","SHS",4.35],["SHS 50x50x4","SHS",5.64],["SHS 60x60x4","SHS",6.90],["SHS 70x70x4","SHS",8.13],["SHS 80x80x5","SHS",11.30],["SHS 90x90x5","SHS",12.80],["SHS 100x100x5","SHS",14.40],["SHS 100x100x6","SHS",17.00],["SHS 120x120x6","SHS",20.70],["SHS 150x150x8","SHS",34.20],["SHS 200x200x8","SHS",46.50],["SHS 250x250x10","SHS",72.70],
  ["CHS 33.7x3.2","CHS",2.41],["CHS 42.4x3.2","CHS",3.09],["CHS 48.3x3.2","CHS",3.56],["CHS 60.3x3.6","CHS",5.03],["CHS 76.1x4.0","CHS",7.11],["CHS 88.9x4.0","CHS",8.38],["CHS 114.3x5.0","CHS",13.50],["CHS 139.7x5.0","CHS",16.60],["CHS 168.3x6.0","CHS",24.00],["CHS 219.1x6.0","CHS",31.50],["CHS 273.0x8.0","CHS",52.30],
  ["C 100x50x2.0","C-COLD",3.13],["C 120x50x2.0","C-COLD",3.45],["C 140x60x2.0","C-COLD",4.08],["C 150x65x2.0","C-COLD",4.40],["C 150x65x2.5","C-COLD",5.46],["C 200x65x2.0","C-COLD",5.18],["C 200x75x2.5","C-COLD",6.87],["C 250x75x2.5","C-COLD",7.66],["C 300x90x3.0","C-COLD",11.10],
  ["Z 100x50x2.0","Z-COLD",3.13],["Z 120x50x2.0","Z-COLD",3.45],["Z 140x60x2.0","Z-COLD",4.08],["Z 150x65x2.0","Z-COLD",4.40],["Z 150x65x2.5","Z-COLD",5.46],["Z 200x65x2.5","Z-COLD",6.45],["Z 200x75x2.5","Z-COLD",6.87],["Z 250x75x3.0","Z-COLD",9.12],["Z 300x90x3.0","Z-COLD",11.10],
  // UPN (European channels, tapered flange)
  ["UPN 80","UPN",8.64],["UPN 100","UPN",10.6],["UPN 120","UPN",13.4],["UPN 140","UPN",16.0],["UPN 160","UPN",18.8],["UPN 180","UPN",22.0],["UPN 200","UPN",25.3],["UPN 220","UPN",29.4],["UPN 240","UPN",33.2],["UPN 260","UPN",37.9],["UPN 280","UPN",41.8],["UPN 300","UPN",46.2],["UPN 320","UPN",59.5],["UPN 350","UPN",60.6],["UPN 380","UPN",63.1],["UPN 400","UPN",71.8],
  // UPE (European channels, parallel flange)
  ["UPE 80","UPE",7.9],["UPE 100","UPE",9.82],["UPE 120","UPE",12.1],["UPE 140","UPE",14.5],["UPE 160","UPE",17.0],["UPE 180","UPE",19.7],["UPE 200","UPE",22.8],["UPE 220","UPE",26.6],["UPE 240","UPE",30.2],["UPE 270","UPE",35.2],["UPE 300","UPE",44.4],["UPE 330","UPE",53.2],["UPE 360","UPE",61.2],["UPE 400","UPE",72.2],
  // PFC (UK parallel flange channels)
  ["100x50x10 PFC","PFC",10.2],["125x65x15 PFC","PFC",14.8],["150x75x18 PFC","PFC",17.9],["150x90x24 PFC","PFC",23.9],["180x75x20 PFC","PFC",20.3],["180x90x26 PFC","PFC",26.1],["200x75x23 PFC","PFC",23.4],["200x90x30 PFC","PFC",29.7],["230x75x26 PFC","PFC",25.7],["230x90x32 PFC","PFC",32.2],["260x75x28 PFC","PFC",27.6],["260x90x35 PFC","PFC",34.8],["300x90x41 PFC","PFC",41.4],["300x100x46 PFC","PFC",45.5],["380x100x54 PFC","PFC",54.0],["430x100x64 PFC","PFC",64.4],
  // HEM European wide-flange (heavy)
  ["HEM 100","HEM",41.8],["HEM 120","HEM",52.1],["HEM 140","HEM",63.2],["HEM 160","HEM",76.2],["HEM 180","HEM",88.9],["HEM 200","HEM",103],["HEM 220","HEM",117],["HEM 240","HEM",157],["HEM 260","HEM",172],["HEM 280","HEM",189],["HEM 300","HEM",238],["HEM 320","HEM",245],["HEM 340","HEM",248],["HEM 360","HEM",250],["HEM 400","HEM",256],["HEM 450","HEM",263],["HEM 500","HEM",270],["HEM 550","HEM",278],["HEM 600","HEM",285],["HEM 650","HEM",293],["HEM 700","HEM",301],["HEM 800","HEM",317],["HEM 900","HEM",333],["HEM 1000","HEM",349],
  // IPN European I-beam (narrow, tapered flange)
  ["IPN 80","IPN",5.94],["IPN 100","IPN",8.34],["IPN 120","IPN",11.1],["IPN 140","IPN",14.3],["IPN 160","IPN",17.9],["IPN 180","IPN",21.9],["IPN 200","IPN",26.2],["IPN 220","IPN",31.1],["IPN 240","IPN",36.2],["IPN 260","IPN",41.9],["IPN 280","IPN",47.9],["IPN 300","IPN",54.2],["IPN 320","IPN",61.0],["IPN 340","IPN",68.0],["IPN 360","IPN",76.1],["IPN 380","IPN",84.0],["IPN 400","IPN",92.4],["IPN 450","IPN",115],["IPN 500","IPN",141],["IPN 550","IPN",166],["IPN 600","IPN",199],
  // W American wide-flange (kg/m from lb/ft)
  ["W6x9","W",13.4],["W6x12","W",17.9],["W6x15","W",22.3],["W6x16","W",23.8],["W6x20","W",29.8],["W6x25","W",37.2],["W8x10","W",14.9],["W8x13","W",19.3],["W8x15","W",22.3],["W8x18","W",26.8],["W8x21","W",31.3],["W8x24","W",35.7],
  ["W8x28","W",41.7],["W8x31","W",46.1],["W8x35","W",52.1],["W8x40","W",59.5],["W8x48","W",71.4],["W8x58","W",86.3],["W8x67","W",99.7],["W10x12","W",17.9],["W10x15","W",22.3],["W10x17","W",25.3],["W10x19","W",28.3],["W10x22","W",32.7],
  ["W10x26","W",38.7],["W10x30","W",44.6],["W10x33","W",49.1],["W10x39","W",58.0],["W10x45","W",67.0],["W10x49","W",72.9],["W10x54","W",80.4],["W10x60","W",89.3],["W10x68","W",101.2],["W10x77","W",114.6],["W10x88","W",131.0],["W10x100","W",148.8],
  ["W10x112","W",166.7],["W12x14","W",20.8],["W12x16","W",23.8],["W12x19","W",28.3],["W12x22","W",32.7],["W12x26","W",38.7],["W12x30","W",44.6],["W12x35","W",52.1],["W12x40","W",59.5],["W12x45","W",67.0],["W12x50","W",74.4],["W12x53","W",78.9],
  ["W12x58","W",86.3],["W12x65","W",96.7],["W12x72","W",107.1],["W12x79","W",117.6],["W12x87","W",129.5],["W12x96","W",142.9],["W12x106","W",157.7],["W12x120","W",178.6],["W12x136","W",202.4],["W12x152","W",226.2],["W12x170","W",253.0],["W12x190","W",282.8],
  ["W12x210","W",312.5],["W12x230","W",342.3],["W12x252","W",375.0],["W12x279","W",415.2],["W12x305","W",453.9],["W12x336","W",500.0],["W14x22","W",32.7],["W14x26","W",38.7],["W14x30","W",44.6],["W14x34","W",50.6],["W14x38","W",56.6],["W14x43","W",64.0],
  ["W14x48","W",71.4],["W14x53","W",78.9],["W14x61","W",90.8],["W14x68","W",101.2],["W14x74","W",110.1],["W14x82","W",122.0],["W14x90","W",133.9],["W14x99","W",147.3],["W14x109","W",162.2],["W14x120","W",178.6],["W14x132","W",196.4],["W14x145","W",215.8],
  ["W14x159","W",236.6],["W14x176","W",261.9],["W14x193","W",287.2],["W14x211","W",314.0],["W14x233","W",346.7],["W14x257","W",382.5],["W14x283","W",421.2],["W14x311","W",462.8],["W14x342","W",509.0],["W14x370","W",550.6],["W14x398","W",592.3],["W14x426","W",634.0],
  ["W14x455","W",677.1],["W14x500","W",744.1],["W14x550","W",818.5],["W14x605","W",900.3],["W14x665","W",989.6],["W14x730","W",1086.4],["W16x26","W",38.7],["W16x31","W",46.1],["W16x36","W",53.6],["W16x40","W",59.5],["W16x45","W",67.0],["W16x50","W",74.4],
  ["W16x57","W",84.8],["W16x67","W",99.7],["W16x77","W",114.6],["W16x89","W",132.4],["W16x100","W",148.8],["W18x35","W",52.1],["W18x40","W",59.5],["W18x46","W",68.5],["W18x50","W",74.4],["W18x55","W",81.8],["W18x60","W",89.3],["W18x65","W",96.7],
  ["W18x71","W",105.7],["W18x76","W",113.1],["W18x86","W",128.0],["W18x97","W",144.4],["W18x106","W",157.7],["W18x119","W",177.1],["W18x130","W",193.5],["W18x143","W",212.8],["W18x158","W",235.1],["W21x44","W",65.5],["W21x50","W",74.4],["W21x57","W",84.8],
  ["W21x62","W",92.3],["W21x68","W",101.2],["W21x73","W",108.6],["W21x83","W",123.5],["W21x93","W",138.4],["W21x101","W",150.3],["W21x111","W",165.2],["W21x122","W",181.6],["W21x132","W",196.4],["W21x147","W",218.8],["W24x55","W",81.8],["W24x62","W",92.3],
  ["W24x68","W",101.2],["W24x76","W",113.1],["W24x84","W",125.0],["W24x94","W",139.9],["W24x103","W",153.3],["W24x104","W",154.8],["W24x117","W",174.1],["W24x131","W",194.9],["W24x146","W",217.3],["W24x162","W",241.1],["W27x84","W",125.0],["W27x94","W",139.9],
  ["W27x102","W",151.8],["W27x114","W",169.7],["W27x129","W",192.0],["W27x146","W",217.3],["W27x161","W",239.6],["W27x178","W",264.9],["W30x90","W",133.9],["W30x99","W",147.3],["W30x108","W",160.7],["W30x116","W",172.6],["W30x124","W",184.5],["W30x132","W",196.4],
  ["W30x148","W",220.2],["W30x173","W",257.5],["W30x191","W",284.2],["W30x211","W",314.0],["W33x118","W",175.6],["W33x130","W",193.5],["W33x141","W",209.8],["W33x152","W",226.2],["W33x169","W",251.5],["W33x201","W",299.1],["W33x221","W",328.9],["W33x241","W",358.6],
  ["W36x135","W",200.9],["W36x150","W",223.2],["W36x160","W",238.1],["W36x170","W",253.0],["W36x182","W",270.8],["W36x194","W",288.7],["W36x210","W",312.5],["W36x231","W",343.8],["W36x232","W",345.3],["W36x247","W",367.6],["W36x256","W",381.0],["W36x262","W",389.9],
  ["W36x282","W",419.7],["W36x302","W",449.4],["W36x361","W",537.2],
  // C American standard channel (kg/m from lb/ft)
  ["C3x4.1","C-AMER",6.1],["C3x5","C-AMER",7.4],["C3x6","C-AMER",8.9],["C4x5.4","C-AMER",8.0],["C4x7.25","C-AMER",10.8],["C5x6.7","C-AMER",10.0],["C5x9","C-AMER",13.4],["C6x8.2","C-AMER",12.2],["C6x10.5","C-AMER",15.6],["C6x13","C-AMER",19.3],["C7x9.8","C-AMER",14.6],["C7x12.25","C-AMER",18.2],
  ["C7x14.75","C-AMER",22.0],["C8x11.5","C-AMER",17.1],["C8x13.75","C-AMER",20.5],["C8x18.75","C-AMER",27.9],["C9x13.4","C-AMER",19.9],["C9x15","C-AMER",22.3],["C9x20","C-AMER",29.8],["C10x15.3","C-AMER",22.8],["C10x20","C-AMER",29.8],["C10x25","C-AMER",37.2],["C10x30","C-AMER",44.6],["C12x20.7","C-AMER",30.8],
  ["C12x25","C-AMER",37.2],["C12x30","C-AMER",44.6],["C15x33.9","C-AMER",50.4],["C15x40","C-AMER",59.5],["C15x50","C-AMER",74.4],
  // MC American miscellaneous channel
  ["MC6x12","MC",17.9],["MC6x15.1","MC",22.5],["MC8x18.7","MC",27.8],["MC8x21.4","MC",31.8],["MC10x22","MC",32.7],["MC10x25","MC",37.2],["MC12x31","MC",46.1],["MC12x35","MC",52.1],["MC18x42.7","MC",63.5],["MC18x58","MC",86.3],
  // UBP UK bearing piles
  ["203x203x45 UBP","UBP",45],["254x254x63 UBP","UBP",63.4],["254x254x71 UBP","UBP",71],["254x254x85 UBP","UBP",85.1],["305x305x79 UBP","UBP",78.9],["305x305x88 UBP","UBP",88],["305x305x95 UBP","UBP",94.9],["305x305x110 UBP","UBP",110],["305x305x126 UBP","UBP",126],["305x305x149 UBP","UBP",149],["305x305x186 UBP","UBP",186],["305x305x223 UBP","UBP",223],["356x368x109 UBP","UBP",109],["356x368x133 UBP","UBP",133],["356x368x152 UBP","UBP",152],["356x368x174 UBP","UBP",174],
  // Equal angles (metric)
  ["L 20x20x3","L",0.87],["L 25x25x3","L",1.11],["L 25x25x4","L",1.44],["L 30x30x3","L",1.34],["L 30x30x4","L",1.76],["L 40x40x4","L",2.39],["L 40x40x5","L",2.94],["L 45x45x5","L",3.34],["L 50x50x5","L",3.73],["L 50x50x6","L",4.43],
  ["L 60x60x6","L",5.37],["L 60x60x8","L",7.03],["L 65x65x6","L",5.84],["L 70x70x6","L",6.31],["L 70x70x7","L",7.31],["L 75x75x6","L",6.78],["L 75x75x8","L",8.92],["L 80x80x8","L",9.55],["L 80x80x10","L",11.78],["L 90x90x8","L",10.8],
  ["L 90x90x10","L",13.34],["L 100x100x10","L",14.91],["L 100x100x12","L",17.71],["L 120x120x10","L",18.05],["L 120x120x12","L",21.48],["L 150x150x12","L",27.13],["L 150x150x15","L",33.56],["L 200x200x16","L",48.23],["L 200x200x20","L",59.66],["L 200x200x24","L",70.84],
  // Unequal angles (metric)
  ["L 75x50x6","L",5.6],["L 75x50x8","L",7.35],["L 100x65x7","L",8.68],["L 100x65x10","L",12.17],["L 100x75x8","L",10.49],["L 100x75x10","L",12.95],["L 125x75x8","L",12.06],["L 125x75x10","L",14.91],["L 150x75x9","L",15.26],
  ["L 150x75x10","L",16.88],["L 150x90x10","L",18.05],["L 150x90x12","L",21.48],["L 150x100x10","L",18.84],["L 150x100x12","L",22.42],["L 200x100x10","L",22.76],["L 200x100x12","L",27.13],["L 200x150x12","L",31.84],["L 200x150x15","L",39.45],
  // RHS rectangular hollow (expanded)
  ["RHS 40x20x2","RHS",1.65],["RHS 40x20x3","RHS",2.3],["RHS 50x25x3","RHS",3.01],["RHS 60x40x4","RHS",5.35],["RHS 70x50x3","RHS",5.13],["RHS 80x40x3","RHS",5.13],
  ["RHS 80x60x4","RHS",7.86],["RHS 90x50x3","RHS",6.07],["RHS 100x40x4","RHS",7.86],["RHS 100x50x3","RHS",6.54],["RHS 100x50x5","RHS",10.32],["RHS 100x60x4","RHS",9.11],["RHS 120x60x4","RHS",10.37],
  ["RHS 120x80x6","RHS",16.74],["RHS 140x80x5","RHS",15.81],["RHS 150x100x6","RHS",21.45],["RHS 200x100x5","RHS",22.09],
  ["RHS 200x100x8","RHS",33.95],["RHS 250x150x6","RHS",35.58],["RHS 250x150x10","RHS",56.96],["RHS 300x200x10","RHS",72.66],["RHS 400x200x10","RHS",88.36],["RHS 400x200x12","RHS",104.64],
  ["RHS 400x300x12","RHS",123.48],["RHS 500x300x12","RHS",142.32],["RHS 500x300x16","RHS",186.02],
  // SHS square hollow (expanded)
  ["SHS 20x20x2","SHS",1.02],["SHS 25x25x2","SHS",1.34],["SHS 25x25x3","SHS",1.83],["SHS 30x30x2","SHS",1.65],["SHS 30x30x3","SHS",2.3],["SHS 40x40x2","SHS",2.28],["SHS 40x40x4","SHS",4.09],
  ["SHS 50x50x5","SHS",6.39],["SHS 60x60x3","SHS",5.13],["SHS 60x60x5","SHS",7.96],["SHS 70x70x5","SHS",9.53],["SHS 70x70x6","SHS",11.09],["SHS 80x80x4","SHS",9.11],
  ["SHS 80x80x6","SHS",12.97],["SHS 90x90x6","SHS",14.86],["SHS 100x100x4","SHS",11.63],["SHS 100x100x8","SHS",21.39],["SHS 120x120x5","SHS",17.38],
  ["SHS 120x120x8","SHS",26.41],["SHS 140x140x6","SHS",24.28],["SHS 140x140x8","SHS",31.43],["SHS 150x150x6","SHS",26.16],["SHS 150x150x10","SHS",41.26],["SHS 160x160x8","SHS",36.46],["SHS 180x180x8","SHS",41.48],
  ["SHS 200x200x10","SHS",56.96],["SHS 200x200x12","SHS",66.96],["SHS 250x250x8","SHS",59.07],["SHS 250x250x12","SHS",85.8],["SHS 300x300x10","SHS",88.36],["SHS 300x300x12","SHS",104.64],["SHS 350x350x12","SHS",123.48],
  ["SHS 400x400x12","SHS",142.32],["SHS 400x400x16","SHS",186.02],
  // CHS circular hollow (expanded)
  ["CHS 21.3x2.6","CHS",1.2],["CHS 26.9x2.6","CHS",1.56],["CHS 26.9x3.2","CHS",1.87],["CHS 33.7x2.6","CHS",1.99],["CHS 42.4x2.6","CHS",2.55],["CHS 48.3x2.6","CHS",2.93],
  ["CHS 48.3x4","CHS",4.37],["CHS 60.3x3.2","CHS",4.51],["CHS 60.3x5","CHS",6.82],["CHS 76.1x3.2","CHS",5.75],["CHS 76.1x4","CHS",7.11],["CHS 76.1x5","CHS",8.77],
  ["CHS 88.9x4","CHS",8.38],["CHS 88.9x5","CHS",10.35],["CHS 114.3x5","CHS",13.48],["CHS 114.3x6.3","CHS",16.78],["CHS 139.7x5","CHS",16.61],["CHS 139.7x6.3","CHS",20.73],["CHS 168.3x5","CHS",20.14],["CHS 168.3x6","CHS",24.02],
  ["CHS 168.3x6.3","CHS",25.17],["CHS 193.7x6.3","CHS",29.12],["CHS 219.1x6","CHS",31.53],["CHS 219.1x6.3","CHS",33.06],["CHS 244.5x6.3","CHS",37.01],["CHS 273x6.3","CHS",41.44],["CHS 273x8","CHS",52.28],["CHS 323.9x8","CHS",62.32],
  ["CHS 355.6x8","CHS",68.58],["CHS 355.6x10","CHS",85.23],["CHS 406.4x10","CHS",97.76],["CHS 457x10","CHS",110.24],["CHS 508x10","CHS",122.81],
  // JIS HW wide-flange — extra (heavy) series
  ["HW 200x204","JIS-HW",56.2],["HW 250x255","JIS-HW",82.2],["HW 300x305","JIS-HW",106],["HW 350x357","JIS-HW",160],["HW 400x408","JIS-HW",197],["HW 414x405","JIS-HW",232],["HW 428x407","JIS-HW",283],["HW 458x417","JIS-HW",415],["HW 498x432","JIS-HW",605],
  // JIS HM medium-flange — deep series
  ["HM 550x300","JIS-HM",137],["HM 600x300","JIS-HM",151],["HM 600x302","JIS-HM",175],["HM 700x300","JIS-HM",185],["HM 800x300","JIS-HM",210],["HM 900x300","JIS-HM",243],
  // JIS HN narrow-flange — extra sizes
  ["HN 198x99","JIS-HN",18.2],["HN 248x124","JIS-HN",25.1],["HN 298x149","JIS-HN",32.0],["HN 346x174","JIS-HN",41.4],["HN 396x199","JIS-HN",56.6],["HN 446x199","JIS-HN",66.7],["HN 496x199","JIS-HN",79.5],["HN 596x199","JIS-HN",94.6],["HN 606x201","JIS-HN",120],["HN 692x300","JIS-HN",166],["HN 792x300","JIS-HN",191],["HN 892x299","JIS-HN",213],
  // JIS I-beams (G3192)
  ["I 100x75x5x8","JIS-I",9.3],["I 125x75x5.5x9.5","JIS-I",13.2],["I 150x75x5.5x9.5","JIS-I",14.0],["I 150x125x8.5x14","JIS-I",28.0],["I 180x100x6x10","JIS-I",21.4],["I 200x100x7x10","JIS-I",26.0],
  ["I 200x150x9x16","JIS-I",50.4],["I 250x125x7.5x12.5","JIS-I",38.3],["I 250x125x10x19","JIS-I",55.5],["I 300x150x8x13","JIS-I",48.3],["I 300x150x10x18.5","JIS-I",65.5],["I 300x150x11.5x22","JIS-I",76.8],
  ["I 350x150x9x15","JIS-I",58.5],["I 350x150x12x24","JIS-I",87.2],["I 400x150x10x18","JIS-I",72.0],["I 400x150x12.5x25","JIS-I",95.8],["I 450x175x11x20","JIS-I",91.7],["I 450x175x13x26","JIS-I",115],
  ["I 600x190x13x25","JIS-I",133],["I 600x190x16x35","JIS-I",176],
  // JIS channels 溝形鋼 (G3192)
  ["C 75x40x5x7","JIS-C",6.92],["C 100x50x5x7.5","JIS-C",9.36],["C 125x65x6x8","JIS-C",13.4],["C 150x75x6.5x10","JIS-C",18.6],["C 150x75x9x12.5","JIS-C",24.0],["C 180x75x7x10.5","JIS-C",21.4],
  ["C 200x80x7.5x11","JIS-C",24.6],["C 200x90x8x13.5","JIS-C",30.3],["C 250x90x9x13","JIS-C",34.6],["C 250x90x11x14.5","JIS-C",40.2],["C 300x90x9x13","JIS-C",38.1],["C 300x90x10x15.5","JIS-C",43.8],
  ["C 300x90x12x16","JIS-C",48.6],["C 380x100x10.5x16","JIS-C",54.5],["C 380x100x13x16.5","JIS-C",62.0],["C 380x100x13x20","JIS-C",67.3],
  // JIS equal angles 等辺山形鋼 (G3192)
  ["L 40x40x3","L",1.81],["L 45x45x4","L",2.7],["L 50x50x4","L",3.01],
  ["L 60x60x4","L",3.64],["L 60x60x5","L",4.51],["L 65x65x5","L",4.91],["L 65x65x8","L",7.66],["L 75x75x9","L",9.96],
  ["L 80x80x6","L",7.25],["L 90x90x7","L",9.51],["L 100x100x7","L",10.61],["L 100x100x13","L",19.08],["L 120x120x8","L",14.57],["L 130x130x9","L",17.73],
  ["L 130x130x12","L",23.36],["L 150x150x19","L",41.91],["L 175x175x12","L",31.84],["L 175x175x15","L",39.45],["L 200x200x15","L",45.33],
  ["L 200x200x25","L",73.59],["L 250x250x25","L",93.22],["L 250x250x35","L",127.76],
  // JIS unequal angles 不等辺山形鋼 (G3192)
  ["L 90x75x9","L",11.02],["L 100x75x7","L",9.23],["L 125x75x7","L",10.61],["L 125x90x10","L",16.09],["L 125x90x13","L",20.61],["L 150x90x9","L",16.32],
  ["L 150x100x9","L",17.03],["L 200x90x9","L",19.85],["L 200x90x14","L",30.33],["L 200x100x15","L",33.56],
  // JIS lipped (light gauge) channels リップ溝形鋼 (G3350)
  ["LC 60x30x10x1.6","JIS-LIP",1.45],["LC 75x45x15x1.6","JIS-LIP",2.06],["LC 75x45x15x2.0","JIS-LIP",2.56],["LC 100x50x20x1.6","JIS-LIP",2.69],["LC 100x50x20x2.0","JIS-LIP",3.34],["LC 100x50x20x2.3","JIS-LIP",3.8],
  ["LC 125x50x20x2.0","JIS-LIP",3.69],["LC 125x50x20x2.3","JIS-LIP",4.21],["LC 150x65x20x2.0","JIS-LIP",4.51],["LC 150x65x20x2.3","JIS-LIP",5.14],["LC 150x75x20x2.3","JIS-LIP",5.5],["LC 200x75x20x2.3","JIS-LIP",6.29],
  ["LC 200x75x25x3.2","JIS-LIP",8.51],["LC 250x75x25x3.2","JIS-LIP",9.55],["LC 250x75x25x4.5","JIS-LIP",13.1],["LC 300x90x30x3.2","JIS-LIP",11.8],
  // JIS Channels (mm) — type JIS-C
  ["C 75x40","JIS-C",6.92],["C 100x50","JIS-C",9.36],["C 125x65","JIS-C",13.4],["C 150x75","JIS-C",18.6],["C 150x75H","JIS-C",24.0],["C 180x75","JIS-C",21.4],["C 200x80","JIS-C",24.6],["C 200x90","JIS-C",30.3],["C 250x90","JIS-C",34.6],["C 250x90H","JIS-C",40.2],["C 300x90","JIS-C",38.1],["C 300x90M","JIS-C",43.8],["C 300x90H","JIS-C",48.6],["C 380x100","JIS-C",54.5],["C 380x100M","JIS-C",62.0],["C 380x100H","JIS-C",67.3],
  // JIS I-beams (mm) — type JIS-I
  ["I 100x75","JIS-I",9.3],["I 125x75","JIS-I",11.5],["I 150x75","JIS-I",14.0],["I 150x125","JIS-I",28.0],["I 180x100","JIS-I",18.4],["I 200x100","JIS-I",21.3],["I 200x150","JIS-I",38.3],["I 250x125","JIS-I",29.6],["I 250x125H","JIS-I",38.3],["I 300x150","JIS-I",36.7],["I 300x150H","JIS-I",48.3],["I 350x150","JIS-I",44.1],["I 350x150H","JIS-I",63.8],["I 400x150","JIS-I",56.6],["I 400x150H","JIS-I",72.0],["I 450x175","JIS-I",76.5],["I 450x175H","JIS-I",91.7],["I 600x190","JIS-I",133],["I 600x190H","JIS-I",176],
  // Russian GOST 8239 I-beams (Двутавр)
  ["GOST I10","GOST-I",9.46],["GOST I12","GOST-I",11.5],["GOST I14","GOST-I",13.7],["GOST I16","GOST-I",15.9],["GOST I18","GOST-I",18.4],["GOST I20","GOST-I",21.0],
  ["GOST I22","GOST-I",24.0],["GOST I24","GOST-I",27.3],["GOST I27","GOST-I",31.5],["GOST I30","GOST-I",36.5],["GOST I33","GOST-I",42.2],["GOST I36","GOST-I",48.6],
  ["GOST I40","GOST-I",57.0],["GOST I45","GOST-I",66.5],["GOST I50","GOST-I",78.5],["GOST I55","GOST-I",92.6],["GOST I60","GOST-I",108.0],
  // Russian GOST 8240 channels (Швеллер)
  ["GOST C5","GOST-C",4.84],["GOST C6.5","GOST-C",5.9],["GOST C8","GOST-C",7.05],["GOST C10","GOST-C",8.59],["GOST C12","GOST-C",10.4],["GOST C14","GOST-C",12.3],
  ["GOST C14a","GOST-C",13.3],["GOST C16","GOST-C",14.2],["GOST C16a","GOST-C",15.3],["GOST C18","GOST-C",16.3],["GOST C18a","GOST-C",17.4],["GOST C20","GOST-C",18.4],
  ["GOST C20a","GOST-C",19.8],["GOST C22","GOST-C",21.0],["GOST C22a","GOST-C",22.6],["GOST C24","GOST-C",24.0],["GOST C24a","GOST-C",25.8],["GOST C27","GOST-C",27.7],
  ["GOST C30","GOST-C",31.8],["GOST C33","GOST-C",36.5],["GOST C36","GOST-C",41.9],["GOST C40","GOST-C",48.3],
  // Russian GOST 26020 wide-flange beams (Б, parallel flange)
  ["GOST 20B1","GOST-B",21.3],["GOST 25B1","GOST-B",25.7],["GOST 25B2","GOST-B",29.6],["GOST 30B1","GOST-B",32.9],["GOST 30B2","GOST-B",36.6],["GOST 35B1","GOST-B",38.9],
  ["GOST 35B2","GOST-B",43.3],["GOST 40B1","GOST-B",48.1],["GOST 40B2","GOST-B",53.0],["GOST 45B1","GOST-B",59.8],["GOST 45B2","GOST-B",66.3],["GOST 50B1","GOST-B",73.0],
  ["GOST 50B2","GOST-B",80.1],["GOST 55B1","GOST-B",89.0],["GOST 55B2","GOST-B",96.8],["GOST 60B1","GOST-B",108.0],["GOST 60B2","GOST-B",117.0],["GOST 70B1","GOST-B",129.0],
  ["GOST 80B1","GOST-B",159.0],["GOST 90B1","GOST-B",183.0],["GOST 100B1","GOST-B",214.0],
  // Russian GOST 26020 column beams (К)
  ["GOST 20K1","GOST-K",41.4],["GOST 20K2","GOST-K",49.9],["GOST 25K1","GOST-K",56.5],["GOST 25K2","GOST-K",67.6],["GOST 30K1","GOST-K",73.8],["GOST 30K2","GOST-K",84.8],
  ["GOST 35K1","GOST-K",109.0],["GOST 35K2","GOST-K",127.0],["GOST 40K1","GOST-K",145.0],["GOST 40K2","GOST-K",163.0],
  // Chinese GB/T 706 I-beams (工字钢)
  ["GB I10","GB-I",11.3],["GB I12.6","GB-I",14.2],["GB I14","GB-I",16.9],["GB I16","GB-I",20.5],["GB I18","GB-I",24.1],["GB I20a","GB-I",27.9],
  ["GB I20b","GB-I",31.1],["GB I22a","GB-I",33.0],["GB I22b","GB-I",36.5],["GB I25a","GB-I",38.1],["GB I25b","GB-I",42.0],["GB I28a","GB-I",43.5],
  ["GB I28b","GB-I",47.9],["GB I32a","GB-I",52.7],["GB I32b","GB-I",57.7],["GB I32c","GB-I",62.8],["GB I36a","GB-I",60.0],["GB I36b","GB-I",65.7],
  ["GB I36c","GB-I",71.3],["GB I40a","GB-I",67.6],["GB I40b","GB-I",73.9],["GB I40c","GB-I",80.2],["GB I45a","GB-I",80.4],["GB I45b","GB-I",87.5],
  ["GB I45c","GB-I",94.6],["GB I50a","GB-I",93.7],["GB I50b","GB-I",101.5],["GB I50c","GB-I",109.4],["GB I56a","GB-I",106.3],["GB I56b","GB-I",115.1],
  ["GB I56c","GB-I",123.9],["GB I63a","GB-I",121.4],["GB I63b","GB-I",131.3],["GB I63c","GB-I",141.2],
  // Chinese GB/T 706 channels (槽钢)
  ["GB C5","GB-C",5.44],["GB C6.3","GB-C",6.63],["GB C8","GB-C",8.04],["GB C10","GB-C",10.0],["GB C12.6","GB-C",12.3],["GB C14a","GB-C",14.5],
  ["GB C14b","GB-C",16.7],["GB C16a","GB-C",17.2],["GB C16b","GB-C",19.8],["GB C18a","GB-C",20.2],["GB C18b","GB-C",23.0],["GB C20a","GB-C",22.6],
  ["GB C20b","GB-C",25.8],["GB C22a","GB-C",25.0],["GB C22b","GB-C",28.5],["GB C25a","GB-C",27.4],["GB C25b","GB-C",31.3],["GB C25c","GB-C",35.3],
  ["GB C28a","GB-C",31.4],["GB C28b","GB-C",35.8],["GB C28c","GB-C",40.2],["GB C32a","GB-C",38.1],["GB C32b","GB-C",43.1],["GB C32c","GB-C",48.1],
  ["GB C36a","GB-C",47.8],["GB C36b","GB-C",53.5],["GB C36c","GB-C",59.1],["GB C40a","GB-C",58.9],["GB C40b","GB-C",65.2],["GB C40c","GB-C",71.5],
  // Additional cold-formed C purlins
  ["C 250x75x2.0","C-COLD",6.13],["C 250x75x3.0","C-COLD",9.12],["C 300x75x3.0","C-COLD",10.3],["C 300x90x2.5","C-COLD",9.34],["C 350x100x3.0","C-COLD",12.5],["C 400x100x3.0","C-COLD",13.7],
  ["C 400x100x4.0","C-COLD",18.1],

  /* ── Hollow sections added from CSI / APL / AISC15M exports ──────────
     Labels are METRIC as published in those databases. Imperial HSS
     written with fractions (HSS8X8X1/2) is NOT matched by these names;
     that needs a fraction-aware parser, not more rows.
     Mass here is derived as A x 7850, which runs ~1-3% above mill
     catalogue figures. Conservative for buying, and the catalogue rows
     above are listed first so findSection() still prefers them. */
  /* HSS — 516 added from AISC15M export.
     MASS CORRECTION: AISC computes HSS section area on the DESIGN wall
     (0.93 x nominal for ASTM A500) but publishes purchase weight on the
     NOMINAL wall. Deriving mass from area therefore understates what you
     actually buy by ~6%. These rows are divided by 0.93 to recover the
     nominal-wall mass: HSS8X8X1/2 -> 73.5 vs AISC published 72.7 kg/m,
     i.e. ~1% conservative (safe for pricing) instead of 6% short. */
  ["HSS50.8X25.4X3.2","HSS",3.31],["HSS48.3X3.0","HSS",3.4],["HSS42.2X3.6","HSS",3.4],["HSS63.5X25.4X3.2","HSS",3.95],["HSS50.8X38.1X3.2","HSS",3.95],["HSS48.3X3.7","HSS",4.08],
  ["HSS60.3X3.2","HSS",4.48],["HSS76.2X25.4X3.2","HSS",4.57],["HSS63.5X38.1X3.2","HSS",4.57],["HSS50.8X50.8X3.2","HSS",4.57],["HSS50.8X25.4X4.8","HSS",4.6],["HSS63.5X3.2","HSS",4.73],
  ["HSS57.2X50.8X3.2","HSS",4.89],["HSS48.3X4.8","HSS",5.13],["HSS76.2X38.1X3.2","HSS",5.2],["HSS63.5X50.8X3.2","HSS",5.2],["HSS57.2X57.2X3.2","HSS",5.2],["HSS60.3X3.9","HSS",5.44],
  ["HSS73X3.2","HSS",5.51],["HSS63.5X25.4X4.8","HSS",5.56],["HSS50.8X38.1X4.8","HSS",5.56],["HSS76.2X3.2","HSS",5.71],["HSS88.9X38.1X3.2","HSS",5.83],["HSS76.2X50.8X3.2","HSS",5.83],
  ["HSS63.5X63.5X3.2","HSS",5.83],["HSS76.2X3.4","HSS",6.11],["HSS88.9X50.8X3.2","HSS",6.48],["HSS76.2X63.5X3.2","HSS",6.48],["HSS76.2X25.4X4.8","HSS",6.48],["HSS63.5X38.1X4.8","HSS",6.48],
  ["HSS50.8X50.8X4.8","HSS",6.48],["HSS60.3X4.8","HSS",6.54],["HSS88.9X3.2","HSS",6.7],["HSS76.2X3.9","HSS",6.91],["HSS63.5X4.8","HSS",6.91],["HSS57.2X50.8X4.8","HSS",6.97],
  ["HSS101.6X50.8X3.2","HSS",7.09],["HSS88.9X63.5X3.2","HSS",7.09],["HSS76.2X76.2X3.2","HSS",7.09],["HSS76.2X38.1X4.8","HSS",7.46],["HSS63.5X50.8X4.8","HSS",7.46],["HSS57.2X57.2X4.8","HSS",7.46],
  ["HSS60.3X5.5","HSS",7.57],["HSS101.6X63.5X3.2","HSS",7.73],["HSS101.6X3.2","HSS",7.73],["HSS73X4.8","HSS",8.06],["HSS63.5X38.1X6.4","HSS",8.23],["HSS50.8X50.8X6.4","HSS",8.23],
  ["HSS127X50.8X3.2","HSS",8.39],["HSS101.6X76.2X3.2","HSS",8.39],["HSS88.9X88.9X3.2","HSS",8.39],["HSS88.9X38.1X4.8","HSS",8.39],["HSS76.2X50.8X4.8","HSS",8.39],["HSS63.5X63.5X4.8","HSS",8.39],
  ["HSS76.2X4.8","HSS",8.39],["HSS60.3X6.4","HSS",8.53],["HSS114.3X3.2","HSS",8.7],["HSS73X5.2","HSS",8.7],["HSS127X63.5X3.2","HSS",8.95],["HSS63.5X6.4","HSS",9.03],
  ["HSS76.2X5.2","HSS",9.12],["HSS88.9X50.8X4.8","HSS",9.28],["HSS76.2X63.5X4.8","HSS",9.28],["HSS76.2X38.1X6.4","HSS",9.45],["HSS63.5X50.8X6.4","HSS",9.45],["HSS57.2X57.2X6.4","HSS",9.45],
  ["HSS152.4X50.8X3.2","HSS",9.62],["HSS127X76.2X3.2","HSS",9.62],["HSS101.6X101.6X3.2","HSS",9.62],["HSS76.2X5.5","HSS",9.62],["HSS127X3.2","HSS",9.71],["HSS88.9X4.8","HSS",9.87],
  ["HSS101.6X50.8X4.8","HSS",10.3],["HSS88.9X63.5X4.8","HSS",10.3],["HSS76.2X76.2X4.8","HSS",10.3],["HSS73X6.4","HSS",10.55],["HSS88.9X38.1X6.4","HSS",10.72],["HSS76.2X50.8X6.4","HSS",10.72],
  ["HSS63.5X63.5X6.4","HSS",10.72],["HSS88.9X5.2","HSS",10.72],["HSS177.8X50.8X3.2","HSS",10.89],["HSS152.4X76.2X3.2","HSS",10.89],["HSS127X101.6X3.2","HSS",10.89],["HSS114.3X114.3X3.2","HSS",10.89],
  ["HSS76.2X6.4","HSS",11.05],["HSS101.6X63.5X4.8","HSS",11.23],["HSS88.9X5.5","HSS",11.31],["HSS101.6X4.8","HSS",11.4],["HSS141.3X3.4","HSS",11.56],["HSS152.4X3.2","HSS",11.65],
  ["HSS88.9X50.8X6.4","HSS",12.08],["HSS76.2X63.5X6.4","HSS",12.08],["HSS203.2X50.8X3.2","HSS",12.15],["HSS177.8X76.2X3.2","HSS",12.15],["HSS152.4X101.6X3.2","HSS",12.15],["HSS127X127X3.2","HSS",12.15],
  ["HSS127X50.8X4.8","HSS",12.24],["HSS101.6X76.2X4.8","HSS",12.24],["HSS88.9X88.9X4.8","HSS",12.24],["HSS76.2X50.8X7.9","HSS",12.83],["HSS63.5X63.5X7.9","HSS",12.83],["HSS114.3X4.8","HSS",12.83],
  ["HSS168.3X3.2","HSS",12.91],["HSS88.9X6.4","HSS",13.0],["HSS127X63.5X4.8","HSS",13.09],["HSS101.6X50.8X6.4","HSS",13.25],["HSS88.9X63.5X6.4","HSS",13.25],["HSS76.2X76.2X6.4","HSS",13.25],
  ["HSS101.6X5.6","HSS",13.25],["HSS203.2X76.2X3.2","HSS",13.42],["HSS177.8X101.6X3.2","HSS",13.42],["HSS152.4X127X3.2","HSS",13.42],["HSS139.7X139.7X3.2","HSS",13.42],["HSS101.6X5.7","HSS",13.59],
  ["HSS177.8X3.2","HSS",13.68],["HSS152.4X50.8X4.8","HSS",14.01],["HSS127X76.2X4.8","HSS",14.01],["HSS101.6X101.6X4.8","HSS",14.01],["HSS101.6X6","HSS",14.18],["HSS76.2X63.5X7.9","HSS",14.34],
  ["HSS127X4.8","HSS",14.34],["HSS101.6X63.5X6.4","HSS",14.52],["HSS254X50.8X3.2","HSS",14.69],["HSS203.2X101.6X3.2","HSS",14.69],["HSS177.8X127X3.2","HSS",14.69],["HSS152.4X152.4X3.2","HSS",14.69],
  ["HSS101.6X6.4","HSS",15.02],["HSS88.9X7.6","HSS",15.37],["HSS127X50.8X6.4","HSS",15.87],["HSS101.6X76.2X6.4","HSS",15.87],["HSS88.9X88.9X6.4","HSS",15.87],["HSS254X76.2X3.2","HSS",15.96],
  ["HSS177.8X50.8X4.8","HSS",15.96],["HSS152.4X76.2X4.8","HSS",15.96],["HSS127X101.6X4.8","HSS",15.96],["HSS114.3X114.3X4.8","HSS",15.96],["HSS88.9X8","HSS",15.96],["HSS101.6X50.8X7.9","HSS",16.03],
  ["HSS88.9X63.5X7.9","HSS",16.03],["HSS76.2X76.2X7.9","HSS",16.03],["HSS141.3X4.8","HSS",16.03],["HSS114.3X6","HSS",16.12],["HSS254X88.9X3.2","HSS",16.55],["HSS127X63.5X6.4","HSS",17.14],
  ["HSS254X101.6X3.2","HSS",17.22],["HSS177.8X177.8X3.2","HSS",17.22],["HSS152.4X4.8","HSS",17.3],["HSS101.6X63.5X7.9","HSS",17.56],["HSS203.2X50.8X4.8","HSS",17.89],["HSS177.8X76.2X4.8","HSS",17.89],
  ["HSS152.4X101.6X4.8","HSS",17.89],["HSS127X127X4.8","HSS",17.89],["HSS152.4X50.8X6.4","HSS",18.31],["HSS127X76.2X6.4","HSS",18.31],["HSS101.6X101.6X6.4","HSS",18.31],["HSS101.6X50.8X9.5","HSS",18.48],
  ["HSS88.9X63.5X9.5","HSS",18.48],["HSS76.2X76.2X9.5","HSS",18.48],["HSS101.6X8","HSS",18.48],["HSS127X6.4","HSS",18.99],["HSS127X50.8X7.9","HSS",19.16],["HSS101.6X76.2X7.9","HSS",19.16],
  ["HSS88.9X88.9X7.9","HSS",19.16],["HSS168.3X4.8","HSS",19.25],["HSS127X6.6","HSS",19.58],["HSS203.2X203.2X3.2","HSS",19.75],["HSS203.2X76.2X4.8","HSS",19.75],["HSS177.8X101.6X4.8","HSS",19.75],
  ["HSS152.4X127X4.8","HSS",19.75],["HSS139.7X139.7X4.8","HSS",19.75],["HSS174.6X4.8","HSS",19.92],["HSS101.6X63.5X9.5","HSS",20.34],["HSS177.8X4.8","HSS",20.34],["HSS177.8X50.8X6.4","HSS",20.94],
  ["HSS152.4X76.2X6.4","HSS",20.94],["HSS127X101.6X6.4","HSS",20.94],["HSS114.3X114.3X6.4","HSS",20.94],["HSS139.7X6.6","HSS",21.61],["HSS254X50.8X4.8","HSS",21.69],["HSS228.6X76.2X4.8","HSS",21.69],
  ["HSS203.2X101.6X4.8","HSS",21.69],["HSS177.8X127X4.8","HSS",21.69],["HSS152.4X152.4X4.8","HSS",21.69],["HSS190.5X4.8","HSS",21.77],["HSS141.3X6.6","HSS",21.86],["HSS228.6X228.6X3.2","HSS",22.28],
  ["HSS127X50.8X9.5","HSS",22.28],["HSS101.6X76.2X9.5","HSS",22.28],["HSS88.9X88.9X9.5","HSS",22.28],["HSS152.4X50.8X7.9","HSS",22.37],["HSS127X76.2X7.9","HSS",22.37],["HSS101.6X101.6X7.9","HSS",22.37],
  ["HSS114.3X8.6","HSS",22.45],["HSS152.4X6.4","HSS",22.96],["HSS203.2X50.8X6.4","HSS",23.38],["HSS177.8X76.2X6.4","HSS",23.38],["HSS152.4X101.6X6.4","HSS",23.38],["HSS127X127X6.4","HSS",23.38],
  ["HSS127X7.9","HSS",23.38],["HSS254X76.2X4.8","HSS",23.55],["HSS254X88.9X4.8","HSS",24.47],["HSS114.3X9.5","HSS",24.82],["HSS219.1X4.8","HSS",25.15],["HSS304.8X50.8X4.8","HSS",25.41],
  ["HSS254X101.6X4.8","HSS",25.41],["HSS228.6X127X4.8","HSS",25.41],["HSS203.2X152.4X4.8","HSS",25.41],["HSS177.8X177.8X4.8","HSS",25.41],["HSS152.4X76.2X7.9","HSS",25.49],["HSS127X101.6X7.9","HSS",25.49],
  ["HSS114.3X114.3X7.9","HSS",25.49],["HSS168.3X6.4","HSS",25.49],["HSS152.4X7.1","HSS",25.58],["HSS203.2X76.2X6.4","HSS",26.0],["HSS177.8X101.6X6.4","HSS",26.0],["HSS152.4X127X6.4","HSS",26.0],
  ["HSS152.4X50.8X9.5","HSS",26.0],["HSS139.7X139.7X6.4","HSS",26.0],["HSS127X76.2X9.5","HSS",26.0],["HSS101.6X101.6X9.5","HSS",26.0],["HSS174.6X6.4","HSS",26.51],["HSS177.8X6.4","HSS",26.92],
  ["HSS304.8X76.2X4.8","HSS",27.34],["HSS254X127X4.8","HSS",27.34],["HSS127X9.5","HSS",27.77],["HSS244.5X4.8","HSS",28.19],["HSS168.3X7.1","HSS",28.28],["HSS152.4X7.9","HSS",28.44],
  ["HSS254X50.8X6.4","HSS",28.53],["HSS228.6X76.2X6.4","HSS",28.53],["HSS203.2X101.6X6.4","HSS",28.53],["HSS177.8X127X6.4","HSS",28.53],["HSS152.4X152.4X6.4","HSS",28.53],["HSS203.2X50.8X7.9","HSS",28.61],
  ["HSS177.8X76.2X7.9","HSS",28.61],["HSS152.4X101.6X7.9","HSS",28.61],["HSS127X127X7.9","HSS",28.61],["HSS190.5X6.4","HSS",28.96],["HSS304.8X101.6X4.8","HSS",29.2],["HSS254X152.4X4.8","HSS",29.2],
  ["HSS228.6X177.8X4.8","HSS",29.2],["HSS203.2X203.2X4.8","HSS",29.2],["HSS254X4.8","HSS",29.2],["HSS152.4X76.2X9.5","HSS",29.88],["HSS127X101.6X9.5","HSS",29.88],["HSS114.3X114.3X9.5","HSS",29.88],
  ["HSS139.7X9.5","HSS",30.81],["HSS254X76.2X6.4","HSS",31.06],["HSS141.3X9.5","HSS",31.15],["HSS168.3X7.9","HSS",31.57],["HSS203.2X76.2X7.9","HSS",31.82],["HSS177.8X101.6X7.9","HSS",31.82],
  ["HSS152.4X127X7.9","HSS",31.82],["HSS139.7X139.7X7.9","HSS",31.82],["HSS254X88.9X6.4","HSS",32.33],["HSS127X76.2X12.7","HSS",32.75],["HSS101.6X101.6X12.7","HSS",32.75],["HSS174.6X7.9","HSS",32.75],
  ["HSS355.6X101.6X4.8","HSS",33.0],["HSS304.8X152.4X4.8","HSS",33.0],["HSS254X203.2X4.8","HSS",33.0],["HSS228.6X228.6X74.8","HSS",33.0],["HSS177.8X7.9","HSS",33.34],["HSS219.1X6.4","HSS",33.43],
  ["HSS304.8X50.8X6.4","HSS",33.59],["HSS254X101.6X6.4","HSS",33.59],["HSS228.6X127X6.4","HSS",33.59],["HSS203.2X152.4X6.4","HSS",33.59],["HSS177.8X177.8X6.4","HSS",33.59],["HSS203.2X50.8X9.5","HSS",33.68],
  ["HSS177.8X76.2X9.5","HSS",33.68],["HSS152.4X101.6X9.5","HSS",33.68],["HSS127X127X9.5","HSS",33.68],["HSS152.4X9.5","HSS",33.76],["HSS254X50.8X7.9","HSS",35.03],["HSS228.6X76.2X7.9","HSS",35.03],
  ["HSS203.2X101.6X7.9","HSS",35.03],["HSS177.8X127X7.9","HSS",35.03],["HSS152.4X152.4X7.9","HSS",35.03],["HSS190.5X7.9","HSS",35.87],["HSS127X12.7","HSS",36.04],["HSS304.8X76.2X6.4","HSS",36.13],
  ["HSS254X127X6.4","HSS",36.13],["HSS406.4X101.6X4.8","HSS",36.81],["HSS355.6X152.4X4.8","HSS",36.81],["HSS304.8X203.2X4.8","HSS",36.81],["HSS254X254X4.8","HSS",36.81],["HSS244.5X6.4","HSS",37.4],
  ["HSS203.2X76.2X9.5","HSS",37.47],["HSS177.8X101.6X9.5","HSS",37.47],["HSS152.4X127X9.5","HSS",37.47],["HSS139.7X139.7X9.5","HSS",37.47],["HSS168.3X9.5","HSS",37.47],["HSS152.4X76.2X12.7","HSS",37.82],
  ["HSS127X101.6X12.7","HSS",37.82],["HSS114.3X114.3X12.7","HSS",37.82],["HSS254X76.2X7.9","HSS",38.15],["HSS193.7X8.3","HSS",38.15],["HSS304.8X101.6X6.4","HSS",38.66],["HSS254X152.4X6.4","HSS",38.66],
  ["HSS228.6X177.8X6.4","HSS",38.66],["HSS203.2X203.2X6.4","HSS",38.66],["HSS254X6.4","HSS",38.91],["HSS174.6X9.5","HSS",39.0],["HSS177.8X9.5","HSS",39.67],["HSS254X88.9X7.9","HSS",39.75],
  ["HSS139.7X12.7","HSS",40.1],["HSS141.3X12.7","HSS",40.6],["HSS254X50.8X9.5","HSS",41.28],["HSS228.6X76.2X9.5","HSS",41.28],["HSS203.2X101.6X9.5","HSS",41.28],["HSS177.8X127X9.5","HSS",41.28],
  ["HSS152.4X152.4X9.5","HSS",41.28],["HSS304.8X50.8X7.9","HSS",41.35],["HSS254X101.6X7.9","HSS",41.35],["HSS228.6X127X7.9","HSS",41.35],["HSS203.2X152.4X7.9","HSS",41.35],["HSS177.8X177.8X7.9","HSS",41.35],
  ["HSS273.1X6.4","HSS",41.95],["HSS219.1X8.2","HSS",42.71],["HSS190.5X9.5","HSS",42.71],["HSS168.3X11","HSS",42.8],["HSS177.8X76.2X12.7","HSS",42.88],["HSS152.4X101.6X12.7","HSS",42.88],
  ["HSS127X127X12.7","HSS",42.88],["HSS193.7X9.5","HSS",43.47],["HSS355.6X101.6X6.4","HSS",43.72],["HSS304.8X152.4X6.4","HSS",43.72],["HSS254X203.2X6.4","HSS",43.72],["HSS228.6X228.6X6.4","HSS",43.72],
  ["HSS152.4X12.7","HSS",44.06],["HSS304.8X304.8X4.8","HSS",44.4],["HSS304.8X76.2X7.9","HSS",44.48],["HSS254X127X7.9","HSS",44.48],["HSS254X76.2X9.5","HSS",45.08],["HSS304.8X88.9X7.9","HSS",46.09],
  ["HSS244.5X7.9","HSS",46.42],["HSS254X88.9X9.5","HSS",46.94],["HSS304.8X101.6X7.9","HSS",47.69],["HSS254X152.4X7.9","HSS",47.69],["HSS228.6X177.8X7.9","HSS",47.69],["HSS203.2X203.2X7.9","HSS",47.69],
  ["HSS203.2X76.2X12.7","HSS",47.95],["HSS177.8X101.6X12.7","HSS",47.95],["HSS152.4X127X12.7","HSS",47.95],["HSS254X7.9","HSS",48.37],["HSS406.4X101.6X6.4","HSS",48.78],["HSS355.6X152.4X6.4","HSS",48.78],
  ["HSS304.8X203.2X6.4","HSS",48.78],["HSS254X254X6.4","HSS",48.78],["HSS254X101.6X9.5","HSS",48.87],["HSS228.6X127X9.5","HSS",48.87],["HSS203.2X152.4X9.5","HSS",48.87],["HSS177.8X177.8X9.5","HSS",48.87],
  ["HSS168.3X12.7","HSS",49.04],["HSS219.1X9.5","HSS",49.38],["HSS323.9X6.4","HSS",49.88],["HSS174.6X12.7","HSS",50.98],["HSS177.8X12.7","HSS",52.0],["HSS254X127X9.5","HSS",52.67],
  ["HSS228.6X76.2X12.7","HSS",53.01],["HSS203.2X101.6X12.7","HSS",53.01],["HSS177.8X127X12.7","HSS",53.01],["HSS152.4X152.4X12.7","HSS",53.01],["HSS304.8X254X6.4","HSS",53.94],["HSS355.6X101.6X7.9","HSS",54.02],
  ["HSS304.8X152.4X7.9","HSS",54.02],["HSS254X203.2X7.9","HSS",54.02],["HSS228.6X228.6X7.9","HSS",54.02],["HSS304.8X88.9X9.5","HSS",54.44],["HSS355.6X6.4","HSS",55.03],["HSS244.5X9.5","HSS",55.54],
  ["HSS190.5X12.7","HSS",56.13],["HSS304.8X101.6X9.5","HSS",56.63],["HSS254X152.4X9.5","HSS",56.63],["HSS228.6X177.8X9.5","HSS",56.63],["HSS203.2X203.2X9.5","HSS",56.63],["HSS254X9.5","HSS",57.73],
  ["HSS508X101.6X6.4","HSS",58.83],["HSS457.2X152.4X6.4","HSS",58.83],["HSS406.4X203.2X6.4","HSS",58.83],["HSS355.6X254X6.4","HSS",58.83],["HSS304.8X304.8X6.4","HSS",58.83],["HSS406.4X101.6X7.9","HSS",60.44],
  ["HSS355.6X152.4X7.9","HSS",60.44],["HSS304.8X203.2X7.9","HSS",60.44],["HSS254X254X7.9","HSS",60.44],["HSS254X88.9X12.7","HSS",60.44],["HSS273.1X9.5","HSS",62.04],["HSS406.4X6.4","HSS",62.63],
  ["HSS254X101.6X12.7","HSS",63.14],["HSS228.6X127X12.7","HSS",63.14],["HSS203.2X152.4X12.7","HSS",63.14],["HSS177.8X177.8X12.7","HSS",63.14],["HSS203.2X101.6X15.9","HSS",63.73],["HSS152.4X152.4X15.9","HSS",63.73],
  ["HSS355.6X101.6X9.5","HSS",64.24],["HSS304.8X152.4X9.5","HSS",64.24],["HSS254X203.2X9.5","HSS",64.24],["HSS228.6X228.6X9.5","HSS",64.24],["HSS219.1X12.7","HSS",64.83],["HSS304.8X254X7.9","HSS",66.43],
  ["HSS355.6X7.9","HSS",68.03],["HSS406.4X101.6X9.5","HSS",71.91],["HSS355.6X152.4X9.5","HSS",71.91],["HSS304.8X203.2X9.5","HSS",71.91],["HSS254X254X9.5","HSS",71.91],["HSS508X101.6X7.9","HSS",73.01],
  ["HSS457.2X152.4X7.9","HSS",73.01],["HSS406.4X203.2X7.9","HSS",73.01],["HSS355.6X254X7.9","HSS",73.01],["HSS304.8X304.8X7.9","HSS",73.01],["HSS244.5X12.7","HSS",73.01],["HSS304.8X101.6X12.7","HSS",73.52],
  ["HSS254X152.4X12.7","HSS",73.52],["HSS228.6X177.8X12.7","HSS",73.52],["HSS203.2X203.2X12.7","HSS",73.52],["HSS323.9X9.5","HSS",74.02],["HSS254X12.7","HSS",75.71],["HSS254X101.6X15.9","HSS",76.23],
  ["HSS228.6X127X15.9","HSS",76.23],["HSS203.2X152.4X15.9","HSS",76.23],["HSS177.8X177.8X15.9","HSS",76.23],["HSS406.4X7.9","HSS",78.42],["HSS304.8X254X9.5","HSS",79.52],["HSS219.1X15.9","HSS",80.02],
  ["HSS355.6X9.5","HSS",81.71],["HSS273.1X12.7","HSS",81.71],["HSS355.6X101.6X12.7","HSS",83.31],["HSS304.8X152.4X12.7","HSS",83.31],["HSS254X203.2X12.7","HSS",83.31],["HSS228.6X228.6X12.7","HSS",83.31],
  ["HSS508X203.2X7.9","HSS",85.25],["HSS406.4X304.8X7.9","HSS",85.25],["HSS355.6X355.6X7.9","HSS",85.25],["HSS508X101.6X9.5","HSS",86.94],["HSS457.2X152.4X9.5","HSS",86.94],["HSS406.4X203.2X9.5","HSS",86.94],
  ["HSS355.6X254X9.5","HSS",86.94],["HSS304.8X304.8X9.5","HSS",86.94],["HSS304.8X101.6X15.9","HSS",89.47],["HSS254X152.4X15.9","HSS",89.47],["HSS228.6X177.8X15.9","HSS",89.47],["HSS203.2X203.2X15.9","HSS",89.47],
  ["HSS406.4X101.6X12.7","HSS",93.69],["HSS355.6X152.4X12.7","HSS",93.69],["HSS304.8X203.2X12.7","HSS",93.69],["HSS254X254X12.7","HSS",93.69],["HSS406.4X9.5","HSS",93.69],["HSS254X15.9","HSS",93.69],
  ["HSS323.9X12.7","HSS",97.06],["HSS508X304.8X7.9","HSS",98.75],["HSS406.4X406.4X7.9","HSS",98.75],["HSS508X203.2X9.5","HSS",102.13],["HSS406.4X304.8X9.5","HSS",102.13],["HSS355.6X355.6X9.5","HSS",102.13],
  ["HSS355.6X101.6X15.9","HSS",102.13],["HSS304.8X152.4X15.9","HSS",102.13],["HSS254X203.2X15.9","HSS",102.13],["HSS228.6X228.6X15.9","HSS",102.13],["HSS304.8X254X12.7","HSS",103.82],["HSS457.2X9.5","HSS",105.51],
  ["HSS406.4X11.1","HSS",108.04],["HSS355.6X12.7","HSS",108.04],["HSS508X101.6X12.7","HSS",113.95],["HSS457.2X152.4X12.7","HSS",113.95],["HSS406.4X203.2X12.7","HSS",113.95],["HSS406.4X101.6X15.9","HSS",113.95],
  ["HSS355.6X254X12.7","HSS",113.95],["HSS355.6X152.4X15.9","HSS",113.95],["HSS304.8X304.8X12.7","HSS",113.95],["HSS304.8X203.2X15.9","HSS",113.95],["HSS254X254X15.9","HSS",113.95],["HSS508X304.8X9.5","HSS",117.32],
  ["HSS406.4X406.4X9.5","HSS",117.32],["HSS508X9.5","HSS",117.32],["HSS406.4X12.7","HSS",123.24],["HSS355.6X15.9","HSS",133.37],["HSS508X203.2X12.7","HSS",134.2],["HSS406.4X304.8X12.7","HSS",134.2],
  ["HSS355.6X355.6X12.7","HSS",134.2],["HSS254X254X19","HSS",134.2],["HSS457.2X12.7","HSS",139.27],["HSS457.2X152.4X15.9","HSS",140.12],["HSS406.4X203.2X15.9","HSS",140.12],["HSS355.6X254X15.9","HSS",140.12],
  ["HSS304.8X304.8X15.9","HSS",140.12],["HSS406.4X15.9","HSS",152.77],["HSS508X304.8X12.7","HSS",154.47],["HSS406.4X406.4X12.7","HSS",154.47],["HSS508X12.7","HSS",155.31],["HSS508X203.2X15.9","HSS",164.59],
  ["HSS406.4X304.8X15.9","HSS",164.59],["HSS355.6X355.6X15.9","HSS",164.59],["HSS304.8X304.8X19","HSS",164.59],["HSS609.6X304.8X12.7","HSS",174.72],["HSS457.2X457.2X12.7","HSS",174.72],["HSS508X304.8X15.9","HSS",190.76],
  ["HSS406.4X406.4X15.9","HSS",190.76],["HSS508X508X12.7","HSS",194.98],["HSS406.4X304.8X19","HSS",195.83],["HSS355.6X355.6X19","HSS",195.83],["HSS609.6X304.8X15.9","HSS",215.24],["HSS457.2X457.2X15.9","HSS",215.24],
  ["HSS355.6X355.6X22.2","HSS",224.53],["HSS508X304.8X19","HSS",226.22],["HSS406.4X406.4X19.0","HSS",226.22],["HSS508X508X15.9","HSS",241.41],["HSS609.6X304.8X19","HSS",256.6],["HSS457.2X457.2X19","HSS",256.6],
  ["HSS406.4X406.4X22.2","HSS",259.98],["HSS508X508X19","HSS",286.14],["HSS457.2X457.2X22.2","HSS",295.43],["HSS558.8X558.8X19","HSS",317.38],["HSS508X508X22.2","HSS",330.88],["HSS558.8X558.8X22.2","HSS",366.33],
  /* SHS — 266 added from CSI/APL/AISC15M export (mass = A x 7850, ~1-3% above mill catalogue) */
  ["SHS12X12X1.6","SHS",0.47],["SHS12X12X1.8","SHS",0.51],["SHS12X12X2","SHS",0.55],["SHS15X15X1.6","SHS",0.62],["SHS15X15X1.8","SHS",0.68],["SHS15X15X2","SHS",0.74],
  ["SHS15X15X2.2","SHS",0.79],["SHS20X20X1.6","SHS",0.87],["SHS20X20X1.8","SHS",0.97],["SHS25X25X1.6","SHS",1.12],["SHS20X20X2.2","SHS",1.13],["SHS25X25X1.8","SHS",1.25],
  ["SHS20X20X2.6","SHS",1.29],["SHS30X30X1.6","SHS",1.37],["SHS25X25X2.2","SHS",1.48],["SHS32X32X1.6","SHS",1.48],["SHS30X30X1.8","SHS",1.53],["SHS32X32X1.8","SHS",1.64],
  ["SHS25X25X2.6","SHS",1.7],["SHS38X38X1.6","SHS",1.77],["SHS32X32X2","SHS",1.81],["SHS30X30X2.2","SHS",1.82],["SHS25X25X2.9","SHS",1.84],["SHS40X40X1.6","SHS",1.88],
  ["SHS32X32X2.2","SHS",1.96],["SHS30X30X2.6","SHS",2.1],["SHS45X45X1.6","SHS",2.13],["SHS38X38X2","SHS",2.18],["SHS32X32X2.6","SHS",2.26],["SHS30X30X2.9","SHS",2.3],
  ["SHS38X38X2.2","SHS",2.38],["SHS50X50X1.6","SHS",2.38],["SHS32X32X2.9","SHS",2.48],["SHS30X30X3.2","SHS",2.49],["SHS40X40X2.2","SHS",2.51],["SHS45X45X2","SHS",2.62],
  ["SHS50X50X1.8","SHS",2.66],["SHS32X32X3.2","SHS",2.68],["SHS30X30X3.6","SHS",2.72],["SHS38X38X2.6","SHS",2.76],["SHS45X45X2.2","SHS",2.86],["SHS40X40X2.6","SHS",2.92],
  ["SHS30X30X4","SHS",2.94],["SHS50X50X2","SHS",2.94],["SHS38X38X2.9","SHS",3.03],["SHS30X30X4.5","SHS",3.19],["SHS50X50X2.2","SHS",3.2],["SHS40X40X2.9","SHS",3.21],
  ["SHS45X45X2.5","SHS",3.21],["SHS38X38X3.2","SHS",3.29],["SHS30X30X5","SHS",3.42],["SHS40X40X3.2","SHS",3.49],["SHS45X45X2.8","SHS",3.55],["SHS60X60X2","SHS",3.56],
  ["SHS30X30X5.4","SHS",3.58],["SHS38X38X3.6","SHS",3.63],["SHS50X50X2.6","SHS",3.74],["SHS30X30X6","SHS",3.79],["SHS40X40X3.6","SHS",3.85],["SHS60X60X2.2","SHS",3.89],
  ["SHS45X45X3.2","SHS",4.0],["SHS50X50X2.9","SHS",4.12],["SHS72X72X2","SHS",4.32],["SHS45X45X3.6","SHS",4.42],["SHS50X50X3.2","SHS",4.5],["SHS60X60X2.6","SHS",4.55],
  ["SHS40X40X4.5","SHS",4.61],["SHS72X72X2.2","SHS",4.73],["SHS80X80X2","SHS",4.82],["SHS45X45X4","SHS",4.83],["SHS50X50X3.6","SHS",4.98],["SHS40X40X5","SHS",4.99],
  ["SHS60X60X2.9","SHS",5.03],["SHS40X40X5.4","SHS",5.28],["SHS80X80X2.2","SHS",5.28],["SHS60X60X3.2","SHS",5.5],["SHS72X72X2.6","SHS",5.53],["SHS91.5X91.5X2","SHS",5.54],
  ["SHS40X40X6","SHS",5.68],["SHS75X75X2.6","SHS",5.78],["SHS50X50X4.5","SHS",6.02],["SHS91.5X91.5X2.2","SHS",6.07],["SHS100X100X2","SHS",6.08],["SHS60X60X3.6","SHS",6.12],
  ["SHS72X72X2.9","SHS",6.12],["SHS80X80X2.6","SHS",6.19],["SHS75X75X2.9","SHS",6.4],["SHS100X100X2.2","SHS",6.66],["SHS72X72X3.2","SHS",6.7],["SHS40X40X8","SHS",6.74],
  ["SHS80X80X2.9","SHS",6.85],["SHS50X50X5.4","SHS",6.97],["SHS75X75X3.2","SHS",7.01],["SHS91.5X91.5X2.6","SHS",7.12],["SHS60X60X4.5","SHS",7.43],["SHS72X72X3.6","SHS",7.47],
  ["SHS80X80X3.2","SHS",7.51],["SHS50X50X6","SHS",7.56],["SHS75X75X3.6","SHS",7.81],["SHS100X100X2.6","SHS",7.82],["SHS91.5X91.5X2.9","SHS",7.9],["SHS72X72X4","SHS",8.22],
  ["SHS80X80X3.6","SHS",8.38],["SHS75X75X4","SHS",8.6],["SHS110X110X2.6","SHS",8.63],["SHS60X60X5.4","SHS",8.67],["SHS91.5X91.5X3.2","SHS",8.67],["SHS100X100X2.9","SHS",8.67],
  ["SHS72X72X4.5","SHS",9.13],["SHS50X50X8","SHS",9.26],["SHS60X60X6","SHS",9.44],["SHS120X120X2.6","SHS",9.45],["SHS100X100X3.2","SHS",9.52],["SHS75X75X4.5","SHS",9.55],
  ["SHS110X110X2.9","SHS",9.58],["SHS91.5X91.5X3.6","SHS",9.67],["SHS125X125X2.6","SHS",9.86],["SHS72X72X5","SHS",10.02],["SHS80X80X4.5","SHS",10.26],["SHS75X75X5","SHS",10.49],
  ["SHS120X120X2.9","SHS",10.5],["SHS110X110X3.2","SHS",10.53],["SHS100X100X3.6","SHS",10.64],["SHS91.5X91.5X4","SHS",10.67],["SHS72X72X5.4","SHS",10.7],["SHS125X125X2.9","SHS",10.95],
  ["SHS75X75X5.4","SHS",11.21],["SHS120X120X3.2","SHS",11.53],["SHS72X72X6","SHS",11.7],["SHS60X60X8","SHS",11.77],["SHS110X110X3.6","SHS",11.77],["SHS91.5X91.5X4.5","SHS",11.88],
  ["SHS125X125X3.2","SHS",12.03],["SHS80X80X5.4","SHS",12.06],["SHS75X75X6","SHS",12.27],["SHS140X140X2.9","SHS",12.32],["SHS120X120X3.6","SHS",12.9],["SHS110X110X4","SHS",12.99],
  ["SHS91.5X91.5X5","SHS",13.08],["SHS100X100X4.5","SHS",13.09],["SHS150X150X2.9","SHS",13.23],["SHS125X125X3.6","SHS",13.46],["SHS140X140X3.2","SHS",13.54],["SHS91.5X91.5X5.4","SHS",14.01],
  ["SHS120X120X4","SHS",14.25],["SHS110X110X4.5","SHS",14.5],["SHS150X150X3.2","SHS",14.55],["SHS72X72X8","SHS",14.78],["SHS125X125X4","SHS",14.88],["SHS140X140X3.6","SHS",15.16],
  ["SHS91.5X91.5X6","SHS",15.38],["SHS100X100X5.4","SHS",15.45],["SHS120X120X4.5","SHS",15.91],["SHS110X110X5","SHS",15.98],["SHS150X150X3.6","SHS",16.29],["SHS125X125X4.5","SHS",16.62],
  ["SHS140X140X4","SHS",16.76],["SHS110X110X5.4","SHS",17.14],["SHS150X150X4","SHS",18.02],["SHS125X125X5","SHS",18.34],["SHS140X140X4.5","SHS",18.74],["SHS120X120X5.4","SHS",18.84],
  ["SHS110X110X6","SHS",18.86],["SHS100X100X7","SHS",19.45],["SHS125X125X5.4","SHS",19.69],["SHS150X150X4.5","SHS",20.15],["SHS140X140X5","SHS",20.69],["SHS175X175X4","SHS",21.16],
  ["SHS110X110X7","SHS",21.65],["SHS125X125X6","SHS",21.69],["SHS180X180X4","SHS",21.78],["SHS140X140X5.4","SHS",22.23],["SHS150X150X5","SHS",22.26],["SHS175X175X4.5","SHS",23.68],
  ["SHS120X120X7","SHS",23.85],["SHS150X150X5.4","SHS",23.93],["SHS200X200X4","SHS",24.3],["SHS110X110X8","SHS",24.33],["SHS180X180X4.5","SHS",24.39],["SHS125X125X7","SHS",24.95],
  ["SHS175X175X5","SHS",26.19],["SHS180X180X5","SHS",26.97],["SHS200X200X4.5","SHS",27.22],["SHS125X125X8","SHS",28.1],["SHS175X175X5.4","SHS",28.17],["SHS140X140X7","SHS",28.24],
  ["SHS180X180X5.4","SHS",29.01],["SHS220X220X4.5","SHS",30.04],["SHS200X200X5","SHS",30.11],["SHS225X225X4.5","SHS",30.75],["SHS175X175X6","SHS",31.11],["SHS180X180X6","SHS",32.05],
  ["SHS200X200X5.4","SHS",32.4],["SHS220X220X5","SHS",33.25],["SHS225X225X5","SHS",34.04],["SHS220X220X5.4","SHS",35.8],["SHS200X200X6","SHS",35.82],["SHS225X225X5.4","SHS",36.64],
  ["SHS250X250X5","SHS",37.96],["SHS220X220X6","SHS",39.59],["SHS225X225X6","SHS",40.53],["SHS175X175X8","SHS",40.66],["SHS275X275X5","SHS",41.89],["SHS250X250X6","SHS",45.24],
  ["SHS175X175X10","SHS",49.78],["SHS275X275X6","SHS",49.95],["SHS180X180X10","SHS",51.35],["SHS220X220X8","SHS",51.96],["SHS225X225X8","SHS",53.22],["SHS300X300X6","SHS",54.66],
  ["SHS325X325X6","SHS",59.37],["SHS220X220X10","SHS",63.91],["SHS350X350X6","SHS",64.08],["SHS225X225X10","SHS",65.48],["SHS275X275X8","SHS",65.78],["SHS300X300X8","SHS",72.06],
  ["SHS220X220X12","SHS",75.46],["SHS225X225X12","SHS",77.35],["SHS200X200X14","SHS",77.8],["SHS325X325X8","SHS",78.34],["SHS275X275X10","SHS",81.18],["SHS350X350X8","SHS",84.62],
  ["SHS220X220X14","SHS",86.59],["SHS225X225X14","SHS",88.79],["SHS375X375X8","SHS",90.9],["SHS275X275X12","SHS",96.19],["SHS325X325X10","SHS",96.88],["SHS400X400X8","SHS",97.18],
  ["SHS250X250X14","SHS",99.78],["SHS350X350X10","SHS",104.73],["SHS275X275X14","SHS",110.77],["SHS375X375X10","SHS",112.58],["SHS325X325X12","SHS",115.03],["SHS400X400X10","SHS",120.43],
  ["SHS300X300X14","SHS",121.76],["SHS325X325X14","SHS",132.75],["SHS375X375X12","SHS",133.87],["SHS450X450X10","SHS",136.13],["SHS300X300X16","SHS",137.51],["SHS350X350X14","SHS",143.74],
  ["SHS325X325X16","SHS",150.07],["SHS500X500X10","SHS",151.83],["SHS300X300X18","SHS",152.84],["SHS375X375X14","SHS",154.73],["SHS450X450X12","SHS",162.13],["SHS350X350X16","SHS",162.63],
  ["SHS400X400X14","SHS",165.72],["SHS325X325X18","SHS",166.97],["SHS300X300X20","SHS",167.75],["SHS375X375X16","SHS",175.19],["SHS500X500X12","SHS",180.97],["SHS350X350X18","SHS",181.1],
  ["SHS325X325X20","SHS",183.45],["SHS450X450X14","SHS",187.7],["SHS375X375X18","SHS",195.23],["SHS350X350X20","SHS",199.15],["SHS400X400X18","SHS",209.36],["SHS500X500X14","SHS",209.68],
  ["SHS450X450X16","SHS",212.87],["SHS375X375X20","SHS",214.85],["SHS400X400X20","SHS",230.55],["SHS450X450X18","SHS",237.62],["SHS500X500X16","SHS",237.99],["SHS450X450X20","SHS",261.95],
  ["SHS500X500X18","SHS",265.88],["SHS500X500X20","SHS",293.35],
  /* RHS — 238 added from CSI/APL/AISC15M export (mass = A x 7850, ~1-3% above mill catalogue) */
  ["RHS26X13X1.6","RHS",0.85],["RHS26X13X1.8","RHS",0.93],["RHS26X13X2","RHS",1.02],["RHS30X20X1.6","RHS",1.12],["RHS40X10X1.6","RHS",1.12],["RHS30X20X1.8","RHS",1.25],
  ["RHS40X10X1.8","RHS",1.25],["RHS30X20X2","RHS",1.37],["RHS40X10X2","RHS",1.37],["RHS40X20X1.6","RHS",1.37],["RHS30X20X2.2","RHS",1.48],["RHS40X25X1.6","RHS",1.5],
  ["RHS40X20X1.8","RHS",1.53],["RHS30X20X2.5","RHS",1.64],["RHS40X25X1.8","RHS",1.67],["RHS50X25X1.6","RHS",1.75],["RHS40X20X2.2","RHS",1.82],["RHS40X25X2","RHS",1.84],
  ["RHS50X30X1.6","RHS",1.88],["RHS40X25X2.2","RHS",1.99],["RHS40X20X2.5","RHS",2.03],["RHS50X25X2","RHS",2.15],["RHS40X25X2.5","RHS",2.23],["RHS50X30X2","RHS",2.31],
  ["RHS50X25X2.2","RHS",2.34],["RHS66X33X1.6","RHS",2.35],["RHS60X40X1.6","RHS",2.38],["RHS75X25X1.6","RHS",2.38],["RHS50X30X2.2","RHS",2.51],["RHS40X25X2.9","RHS",2.53],
  ["RHS50X25X2.5","RHS",2.62],["RHS50X30X2.5","RHS",2.82],["RHS80X40X1.6","RHS",2.88],["RHS66X33X2","RHS",2.9],["RHS60X40X2","RHS",2.94],["RHS75X25X2","RHS",2.94],
  ["RHS50X25X2.9","RHS",2.98],["RHS75X50X1.6","RHS",3.01],["RHS66X33X2.2","RHS",3.17],["RHS60X40X2.2","RHS",3.2],["RHS75X25X2.2","RHS",3.2],["RHS50X30X2.9","RHS",3.21],
  ["RHS80X40X1.8","RHS",3.23],["RHS50X25X3.2","RHS",3.24],["RHS50X30X3.2","RHS",3.49],["RHS96X48X1.6","RHS",3.49],["RHS66X33X2.5","RHS",3.56],["RHS80X40X2","RHS",3.56],
  ["RHS60X40X2.5","RHS",3.6],["RHS75X25X2.5","RHS",3.6],["RHS100X50X1.6","RHS",3.63],["RHS75X50X2","RHS",3.72],["RHS75X25X2.6","RHS",3.74],["RHS50X30X3.6","RHS",3.85],
  ["RHS80X40X2.2","RHS",3.89],["RHS66X33X2.9","RHS",4.07],["RHS75X50X2.2","RHS",4.07],["RHS60X40X2.9","RHS",4.12],["RHS75X25X2.9","RHS",4.12],["RHS50X30X4","RHS",4.2],
  ["RHS70X30X3","RHS",4.25],["RHS150X25X1.6","RHS",4.26],["RHS96X48X2","RHS",4.32],["RHS80X40X2.5","RHS",4.39],["RHS120X60X1.6","RHS",4.39],["RHS150X30X1.6","RHS",4.39],
  ["RHS66X33X3.2","RHS",4.44],["RHS60X40X3.2","RHS",4.5],["RHS75X25X3.2","RHS",4.5],["RHS70X30X3.2","RHS",4.5],["RHS100X50X2","RHS",4.51],["RHS75X50X2.5","RHS",4.58],
  ["RHS50X30X4.5","RHS",4.61],["RHS96X48X2.2","RHS",4.73],["RHS150X50X1.6","RHS",4.89],["RHS150X30X1.8","RHS",4.92],["RHS100X50X2.2","RHS",4.93],["RHS60X40X3.6","RHS",4.98],
  ["RHS70X30X3.6","RHS",4.98],["RHS50X30X5","RHS",4.99],["RHS80X40X2.9","RHS",5.03],["RHS75X50X2.9","RHS",5.26],["RHS150X25X2","RHS",5.29],["RHS96X48X2.5","RHS",5.33],
  ["RHS120X60X2","RHS",5.45],["RHS150X30X2","RHS",5.45],["RHS70X30X4","RHS",5.46],["RHS70X50X3.2","RHS",5.5],["RHS80X40X3.2","RHS",5.5],["RHS100X50X2.5","RHS",5.57],
  ["RHS50X30X6","RHS",5.68],["RHS75X50X3.2","RHS",5.75],["RHS150X25X2.2","RHS",5.79],["RHS120X60X2.2","RHS",5.97],["RHS60X40X4.5","RHS",6.02],["RHS70X30X4.5","RHS",6.02],
  ["RHS150X50X2","RHS",6.08],["RHS70X50X3.6","RHS",6.12],["RHS80X40X3.6","RHS",6.12],["RHS96X48X2.9","RHS",6.12],["RHS75X50X3.6","RHS",6.4],["RHS100X50X2.9","RHS",6.4],
  ["RHS60X40X5","RHS",6.56],["RHS70X30X5","RHS",6.56],["RHS150X50X2.2","RHS",6.66],["RHS96X48X3.2","RHS",6.7],["RHS70X50X4","RHS",6.71],["RHS50X30X8","RHS",6.74],
  ["RHS120X60X2.5","RHS",6.74],["RHS150X30X2.5","RHS",6.74],["RHS150X25X2.6","RHS",6.8],["RHS70X30X5.4","RHS",6.97],["RHS100X50X3.2","RHS",7.01],["RHS75X50X4","RHS",7.03],
  ["RHS96X48X3.6","RHS",7.47],["RHS150X50X2.5","RHS",7.53],["RHS60X40X6","RHS",7.56],["RHS70X30X6","RHS",7.56],["RHS120X60X2.9","RHS",7.76],["RHS150X30X2.9","RHS",7.76],
  ["RHS100X50X3.6","RHS",7.81],["RHS70X50X5","RHS",8.13],["RHS80X40X5","RHS",8.13],["RHS96X48X4","RHS",8.22],["RHS150X50X2.8","RHS",8.38],["RHS150X75X2.5","RHS",8.51],
  ["RHS120X60X3.2","RHS",8.52],["RHS70X50X5.4","RHS",8.67],["RHS96X48X4.5","RHS",9.13],["RHS70X30X8","RHS",9.26],["RHS70X50X6","RHS",9.44],["RHS120X60X3.6","RHS",9.51],
  ["RHS150X50X3.2","RHS",9.52],["RHS150X75X2.9","RHS",9.81],["RHS96X48X5","RHS",10.02],["RHS150X50X3.6","RHS",10.64],["RHS150X75X3.2","RHS",10.78],["RHS150X100X2.9","RHS",10.95],
  ["RHS100X50X5.4","RHS",11.21],["RHS120X60X4.5","RHS",11.67],["RHS150X50X4","RHS",11.74],["RHS70X50X8","RHS",11.77],["RHS150X100X3.2","RHS",12.03],["RHS150X75X3.6","RHS",12.05],
  ["RHS100X50X6","RHS",12.27],["RHS150X50X4.5","RHS",13.09],["RHS200X100X2.9","RHS",13.23],["RHS150X75X4","RHS",13.31],["RHS150X100X3.6","RHS",13.46],["RHS150X50X5","RHS",14.41],
  ["RHS200X100X3.2","RHS",14.55],["RHS150X100X4","RHS",14.88],["RHS120X60X6","RHS",15.1],["RHS100X50X8","RHS",15.54],["RHS200X100X3.6","RHS",16.29],["RHS150X75X5","RHS",16.38],
  ["RHS200X100X4","RHS",18.02],["RHS150X75X6","RHS",19.33],["RHS250X100X4","RHS",21.16],["RHS240X120X4","RHS",21.78],["RHS280X100X4","RHS",23.04],["RHS250X150X4","RHS",24.3],
  ["RHS250X100X5","RHS",26.19],["RHS240X120X5","RHS",26.97],["RHS300X150X4","RHS",27.44],["RHS150X100X8","RHS",28.1],["RHS250X100X5.4","RHS",28.17],["RHS280X100X5","RHS",28.54],
  ["RHS240X120X5.4","RHS",29.01],["RHS250X150X5","RHS",30.11],["RHS300X200X4","RHS",30.58],["RHS280X100X5.4","RHS",30.71],["RHS250X100X6","RHS",31.11],["RHS240X120X6","RHS",32.05],
  ["RHS250X150X5.4","RHS",32.4],["RHS280X100X6","RHS",33.94],["RHS300X150X5","RHS",34.04],["RHS300X150X5.4","RHS",36.64],["RHS300X200X5","RHS",37.96],["RHS300X150X6","RHS",40.53],
  ["RHS250X100X8","RHS",40.66],["RHS300X200X5.4","RHS",40.88],["RHS240X120X8","RHS",41.91],["RHS280X100X8","RHS",44.42],["RHS300X200X6","RHS",45.24],["RHS350X250X5","RHS",45.81],
  ["RHS400X200X5","RHS",45.81],["RHS350X250X5.4","RHS",49.36],["RHS300X150X8","RHS",53.22],["RHS350X250X6","RHS",54.66],["RHS400X200X6","RHS",54.66],["RHS400X300X6","RHS",64.08],
  ["RHS500X200X6","RHS",64.08],["RHS300X150X10","RHS",65.48],["RHS350X250X8","RHS",72.06],["RHS400X200X8","RHS",72.06],["RHS600X200X6","RHS",73.5],["RHS400X300X8","RHS",84.62],
  ["RHS500X200X8","RHS",84.62],["RHS300X200X12","RHS",86.77],["RHS350X250X10","RHS",89.03],["RHS500X300X8","RHS",97.18],["RHS600X200X8","RHS",97.18],["RHS300X200X14","RHS",99.78],
  ["RHS400X300X10","RHS",104.73],["RHS500X200X10","RHS",104.73],["RHS350X250X12","RHS",105.61],["RHS600X300X8","RHS",109.74],["RHS500X300X10","RHS",120.43],["RHS600X200X10","RHS",120.43],
  ["RHS350X250X14","RHS",121.76],["RHS400X200X14","RHS",121.76],["RHS600X400X8","RHS",122.3],["RHS500X200X12","RHS",124.45],["RHS600X300X10","RHS",136.13],["RHS600X200X12","RHS",143.29],
  ["RHS400X300X14","RHS",143.74],["RHS500X200X14","RHS",143.74],["RHS600X400X10","RHS",151.83],["RHS600X300X12","RHS",162.13],["RHS400X300X16","RHS",162.63],["RHS500X200X16","RHS",162.63],
  ["RHS500X300X14","RHS",165.72],["RHS600X200X14","RHS",165.72],["RHS600X400X12","RHS",180.97],["RHS400X300X18","RHS",181.1],["RHS600X300X14","RHS",187.7],["RHS600X200X16","RHS",187.75],
  ["RHS400X300X20","RHS",199.15],["RHS500X300X18","RHS",209.36],["RHS600X400X14","RHS",209.68],["RHS600X300X16","RHS",212.87],["RHS500X300X20","RHS",230.55],["RHS600X300X18","RHS",237.62],
  ["RHS600X400X16","RHS",237.99],["RHS600X300X20","RHS",261.95],["RHS600X400X18","RHS",265.88],["RHS600X400X20","RHS",293.35],
  /* CHS — 247 added from CSI/APL/AISC15M export (mass = A x 7850, ~1-3% above mill catalogue) */
  ["CHS15X1.6","CHS",0.78],["CHS15X1.8","CHS",0.87],["CHS15X2","CHS",0.95],["CHS20X1.6","CHS",0.98],["CHS15X2.3","CHS",1.08],["CHS20X1.8","CHS",1.09],
  ["PIPE15STD","CHS",1.19],["CHS15X2.6","CHS",1.2],["CHS20X2","CHS",1.2],["CHS30X1.6","CHS",1.27],["CHS15X2.9","CHS",1.32],["CHS20X2.3","CHS",1.37],
  ["CHS30X1.8","CHS",1.42],["PIPE15XS","CHS",1.53],["CHS20X2.6","CHS",1.53],["CHS25X2","CHS",1.56],["PIPE20STD","CHS",1.58],["CHS35X1.6","CHS",1.61],
  ["CHS20X2.9","CHS",1.68],["CHS25X2.3","CHS",1.78],["CHS35X1.8","CHS",1.8],["CHS45X1.6","CHS",1.84],["CHS25X2.6","CHS",1.99],["CHS35X2","CHS",1.99],
  ["PIPE20XS","CHS",2.06],["CHS40X1.8","CHS",2.06],["CHS25X2.9","CHS",2.2],["CHS35X2.3","CHS",2.27],["CHS40X2","CHS",2.28],["CHS55X1.6","CHS",2.32],
  ["PIPE25STD","CHS",2.38],["CHS25X3.2","CHS",2.41],["CHS35X2.6","CHS",2.55],["CHS40X2.3","CHS",2.61],["CHS25X3.6","CHS",2.67],["CHS30X3.2","CHS",2.75],
  ["CHS35X2.9","CHS",2.82],["CHS55X2","CHS",2.88],["CHS40X2.6","CHS",2.93],["PIPE25XS","CHS",3.05],["CHS30X3.6","CHS",3.06],["CHS35X3.2","CHS",3.09],
  ["PIPE32STD","CHS",3.16],["CHS40X2.9","CHS",3.25],["CHS55X2.3","CHS",3.29],["CHS30X4","CHS",3.36],["CHS35X3.6","CHS",3.44],["CHS40X3.2","CHS",3.56],
  ["CHS70X2","CHS",3.66],["CHS55X2.6","CHS",3.7],["CHS25X4.5","CHS",3.73],["PIPE40STD","CHS",3.79],["CHS40X3.6","CHS",3.97],["CHS25X5","CHS",4.08],
  ["CHS50X2.9","CHS",4.11],["CHS70X2.3","CHS",4.19],["PIPE32XS","CHS",4.24],["CHS80X2","CHS",4.29],["CHS25X5.4","CHS",4.35],["CHS40X4","CHS",4.37],
  ["CHS50X3.2","CHS",4.51],["CHS70X2.6","CHS",4.72],["CHS25X6","CHS",4.75],["CHS35X4.5","CHS",4.86],["CHS80X2.3","CHS",4.91],["CHS95X2","CHS",4.91],
  ["CHS50X3.6","CHS",5.03],["PIPE40XS","CHS",5.06],["PIPE50STD","CHS",5.17],["CHS70X2.9","CHS",5.24],["CHS35X5","CHS",5.34],["CHS65X3.2","CHS",5.39],
  ["CHS80X2.6","CHS",5.53],["CHS110X2","CHS",5.54],["CHS50X4","CHS",5.55],["CHS95X2.3","CHS",5.63],["CHS35X5.4","CHS",5.71],["CHS20X8","CHS",5.94],
  ["CHS60X3.6","CHS",6.03],["CHS80X2.9","CHS",6.15],["CHS50X4.5","CHS",6.19],["CHS35X6","CHS",6.26],["CHS95X2.6","CHS",6.35],["CHS105X2.3","CHS",6.35],
  ["CHS65X3.6","CHS",6.45],["CHS60X4","CHS",6.66],["CHS80X3.2","CHS",6.76],["CHS50X5","CHS",6.82],["CHS95X2.9","CHS",7.06],["PIPE50XS","CHS",7.09],
  ["CHS65X4","CHS",7.12],["CHS105X2.6","CHS",7.16],["CHS45X5.4","CHS",7.31],["CHS60X4.5","CHS",7.44],["CHS80X3.6","CHS",7.57],["CHS95X3.2","CHS",7.77],
  ["CHS30X8","CHS",7.95],["CHS65X4.5","CHS",7.96],["CHS105X2.9","CHS",7.97],["CHS120X2.6","CHS",7.98],["CHS45X6","CHS",8.03],["PIPE65STD","CHS",8.16],
  ["CHS60X5","CHS",8.2],["CHS80X4","CHS",8.38],["CHS90X3.6","CHS",8.7],["CHS105X3.2","CHS",8.77],["CHS65X5","CHS",8.78],["CHS130X2.6","CHS",8.79],
  ["CHS60X5.4","CHS",8.8],["CHS120X2.9","CHS",8.88],["CHS75X4.5","CHS",9.37],["CHS65X5.4","CHS",9.43],["CHS90X4","CHS",9.63],["CHS55X6","CHS",9.69],
  ["CHS120X3.2","CHS",9.77],["CHS130X2.9","CHS",9.78],["CHS105X3.6","CHS",9.83],["CHS40X8","CHS",10.32],["CHS75X5","CHS",10.35],["CHS60X6","CHS",10.39],
  ["CHS155X2.6","CHS",10.42],["PIPE80STD","CHS",10.52],["PIPE65XS","CHS",10.6],["CHS160X2.6","CHS",10.62],["CHS130X3.2","CHS",10.77],["CHS90X4.5","CHS",10.78],
  ["CHS105X4","CHS",10.88],["CHS115X3.6","CHS",10.96],["CHS75X5.4","CHS",11.12],["CHS155X2.9","CHS",11.6],["CHS160X2.9","CHS",11.83],["CHS90X5","CHS",11.91],
  ["CHS130X3.6","CHS",12.08],["CHS115X4","CHS",12.13],["CHS105X4.5","CHS",12.19],["CHS185X2.6","CHS",12.25],["CHS75X6","CHS",12.27],["CHS55X8","CHS",12.53],
  ["PIPE90STD","CHS",12.64],["PIPE50XXS","CHS",12.72],["CHS155X3.2","CHS",12.78],["CHS160X3.2","CHS",13.03],["CHS130X4","CHS",13.39],["CHS60X8","CHS",13.46],
  ["CHS100X5","CHS",13.48],["CHS115X4.5","CHS",13.59],["CHS185X2.9","CHS",13.65],["CHS210X2.6","CHS",13.92],["CHS155X3.6","CHS",14.34],["PIPE80XS","CHS",14.37],
  ["CHS160X3.6","CHS",14.62],["PIPE100STD","CHS",14.99],["CHS130X4.5","CHS",15.0],["CHS185X3.2","CHS",15.03],["CHS115X5","CHS",15.04],["CHS210X2.9","CHS",15.51],
  ["CHS155X4","CHS",15.89],["CHS70X8","CHS",15.96],["CHS100X6","CHS",16.03],["CHS160X4","CHS",16.21],["CHS125X5","CHS",16.61],["CHS185X3.6","CHS",16.88],
  ["CHS210X3.2","CHS",17.09],["PIPE90XS","CHS",17.35],["CHS155X4.5","CHS",17.82],["CHS115X6","CHS",17.9],["CHS185X4","CHS",18.71],["CHS210X3.6","CHS",19.19],
  ["PIPE65XXS","CHS",19.39],["CHS65X10","CHS",19.46],["CHS155X5","CHS",19.74],["CHS125X6","CHS",19.78],["PIPE125STD","CHS",20.33],["PIPE100XS","CHS",20.96],
  ["CHS180X4.5","CHS",21.0],["CHS210X4","CHS",21.28],["CHS180X5","CHS",23.27],["CHS150X6","CHS",23.54],["CHS210X4.5","CHS",23.88],["CHS155X6","CHS",24.02],
  ["CHS245X4","CHS",24.66],["CHS120X8","CHS",25.98],["PIPE80XXS","CHS",26.22],["PIPE150STD","CHS",26.3],["CHS205X5","CHS",26.47],["CHS245X4.5","CHS",27.69],
  ["CHS180X6","CHS",27.77],["PIPE125XS","CHS",29.04],["CHS260X4.5","CHS",29.81],["CHS240X5","CHS",30.7],["CHS145X8","CHS",30.99],["CHS205X6","CHS",31.62],
  ["CHS150X8","CHS",31.63],["CHS260X5","CHS",33.06],["CHS310X4.5","CHS",35.45],["CHS175X8","CHS",36.64],["CHS240X6","CHS",36.7],["PIPE100XXS","CHS",38.78],
  ["CHS310X5","CHS",39.32],["CHS260X6","CHS",39.52],["PIPE150XS","CHS",39.64],["PIPE200STD","CHS",39.72],["CHS200X8","CHS",41.77],["CHS310X6","CHS",47.04],
  ["CHS235X8","CHS",48.53],["CHS195X10","CHS",51.72],["CHS340X6","CHS",51.73],["CHS255X8","CHS",52.3],["PIPE125XXS","CHS",54.16],["CHS365X6","CHS",54.9],
  ["PIPE250STD","CHS",58.25],["CHS390X6","CHS",59.25],["CHS230X10","CHS",60.17],["PIPE200XS","CHS",60.29],["CHS305X8","CHS",62.32],["CHS250X10","CHS",64.88],
  ["CHS335X8","CHS",68.58],["PIPE300STD","CHS",69.39],["CHS360X8","CHS",72.8],["PIPE150XXS","CHS",74.42],["PIPE350STD","CHS",75.99],["PIPE250XS","CHS",76.46],
  ["CHS245X12","CHS",77.27],["CHS300X10","CHS",77.41],["CHS390X8","CHS",78.6],["CHS335X10","CHS",85.23],["PIPE400STD","CHS",87.13],["PIPE300XS","CHS",88.7],
  ["CHS355X10","CHS",90.51],["CHS295X12","CHS",92.3],["PIPE350XS","CHS",97.34],["CHS385X10","CHS",97.76],["PIPE450STD","CHS",98.12],["PIPE200XXS","CHS",101.26],
  ["CHS330X12","CHS",101.68],["CHS350X12","CHS",108.02],["PIPE500STD","CHS",109.11],["PIPE400XS","CHS",111.47],["CHS380X12","CHS",116.72],["PIPE450XS","CHS",125.6],
  ["PIPE600STD","CHS",131.88],["PIPE500XS","CHS",139.73],["PIPE650STD","CHS",142.87],["PIPE250XXS","CHS",146.01],["PIPE600XS","CHS",168.77],["PIPE300XXS","CHS",178.98],
  ["PIPE650XS","CHS",182.9],

  /* ── Open shapes added from AISC15M / ArcelorMittal exports ──────────
     Metric W/UB/UC/HP/MC/C/L designations (W310X39, UB457X152X52 ...).
     The imperial rows above (W6x9, ...) are listed first, so a lookup
     still resolves to the original catalogue value where both exist. */
  /* W — 283 added (metric AISC15M / ArcelorMittal labels) */
  ["W150X13","W",12.8],["W150X13.5","W",13.58],["W200X15","W",14.99],["W250X17.9","W",17.9],["W150X18","W",17.98],["W100X19.3","W",19.39],
  ["W200X19.3","W",19.47],["W310X21","W",21.04],["W250X22.3","W",22.37],["W200X22.5","W",22.45],["W150X22.5","W",22.45],["W310X23.8","W",23.86],
  ["W130X23.8","W",23.86],["W150X24","W",24.02],["W250X25.3","W",25.28],["W200X26.6","W",26.61],["W310X28.3","W",28.18],["W130X28.1","W",28.18],
  ["W250X28.4","W",28.5],["W150X29.8","W",29.75],["W200X31.3","W",31.16],["W310X32.7","W",32.81],["W360X32.9","W",32.89],["W250X32.7","W",32.89],
  ["W200X35.9","W",35.87],["W150X37.1","W",37.21],["W250X38.5","W",38.54],["W310X38.7","W",38.78],["W410X38.8","W",38.86],["W360X39","W",38.94],
  ["W200X41.7","W",41.76],["W310X44.5","W",44.51],["W250X44.8","W",44.74],["W360X44","W",44.82],["W410X46.1","W",46.24],["W200X46.1","W",46.24],
  ["W250X49.1","W",49.14],["W360X51","W",50.63],["W460X52","W",52.2],["W310X52","W",52.2],["W200X52","W",52.2],["W410X53","W",53.69],
  ["W360X57.8","W",56.76],["W250X58","W",58.25],["W310X60","W",59.27],["W200X59","W",59.27],["W460X60","W",59.74],["W410X60","W",59.74],
  ["W360X64","W",63.82],["W530X66","W",65.86],["W310X67","W",66.33],["W410X67","W",67.35],["W250X67","W",67.35],["W460X68","W",68.37],
  ["W530X72","W",71.43],["W360X72","W",71.43],["W200X71","W",71.43],["W250X73","W",72.93],["W310X74","W",73.95],["W530X74","W",74.42],
  ["W460X74","W",74.42],["W410X75","W",74.42],["W360X79","W",79.28],["W310X79","W",79.28],["W250X80","W",80.07],["W610X82","W",82.42],
  ["W530X82","W",82.42],["W460X82","W",82.42],["W530X85","W",84.78],["W410X85","W",84.78],["W310X86","W",86.35],["W200X86","W",86.35],
  ["W460X89","W",89.49],["W250X89","W",89.49],["W360X91","W",90.27],["W610X92","W",91.84],["W530X92","W",92.63],["W460X97","W",96.55],
  ["W310X97","W",96.55],["W410X100","W",98.91],["W200X100","W",99.69],["W250X101","W",100.48],["W530X101","W",101.26],["W360X101","W",101.26],
  ["W610X101","W",102.05],["W460X106","W",105.97],["W310X107","W",106.76],["W530X109","W",109.11],["W360X110","W",110.68],["W460X113","W",113.04],
  ["W610X113","W",113.82],["W410X114","W",114.61],["W250X115","W",114.61],["W310X117","W",117.75],["W360X122","W",121.67],["W530X123","W",123.24],
  ["W690X125","W",124.81],["W610X125","W",124.81],["W460X128","W",127.95],["W310X129","W",129.52],["W250X131","W",131.88],["W410X132","W",132.66],
  ["W760X134","W",133.45],["W360X134","W",134.23],["W530X138","W",138.16],["W690X140","W",139.73],["W610X140","W",140.51],["W310X143","W",142.87],
  ["W460X144","W",144.44],["W760X147","W",146.79],["W360X147","W",147.58],["W250X149","W",148.36],["W410X149","W",149.15],["W530X150","W",150.72],
  ["W690X152","W",152.29],["W610X153","W",153.07],["W610X155","W",155.43],["W460X158","W",157.78],["W310X158","W",157.78],["W760X161","W",160.92],
  ["W360X162","W",161.71],["W530X165","W",164.85],["W250X167","W",166.42],["W690X170","W",170.34],["W760X173","W",173.48],["W610X174","W",174.27],
  ["W840X176","W",175.84],["W460X177","W",177.41],["W310X179","W",178.19],["W360X179","W",178.98],["W530X182","W",182.12],["W760X185","W",184.47],
  ["W690X192","W",191.54],["W840X193","W",193.89],["W460X193","W",193.89],["W610X195","W",195.46],["W760X196","W",196.25],["W530X196","W",196.25],
  ["W360X196","W",196.25],["W920X201","W",201.74],["W310X202","W",201.74],["W840X210","W",210.38],["W460X213","W",212.73],["W360X216","W",215.87],
  ["W610X217","W",217.44],["W690X217","W",219.01],["W530X219","W",219.01],["W760X220","W",220.58],["W1000X222","W",222.15],["W920X223","W",224.51],
  ["W310X226","W",226.08],["W840X226","W",227.65],["W460X235","W",234.71],["W360X237","W",236.28],["W920X238","W",237.85],["W690X240","W",240.99],
  ["W610X241","W",241.78],["W530X248","W",247.27],["W1000X249","W",249.63],["W840X251","W",250.41],["W920X253","W",253.55],["W310X253","W",253.55],
  ["W760X257","W",257.48],["W460X260","W",260.62],["W610X262","W",262.19],["W360X262","W",262.19],["W690X265","W",266.11],["W1000X272","W",270.04],
  ["W920X271","W",271.61],["W530X272","W",271.61],["W310X283","W",283.38],["W760X284","W",284.17],["W460X286","W",284.95],["W610X285","W",286.52],
  ["W360X287","W",287.31],["W920X289","W",288.88],["W690X289","W",288.88],["W1000X296","W",297.51],["W840X299","W",299.08],["W530X300","W",300.65],
  ["W610X307","W",307.72],["W920X313","W",313.21],["W310X313","W",313.21],["W360X314","W",314.0],["W1000X314","W",314.78],["W760X314","W",315.57],
  ["W460X315","W",315.57],["W1000X321","W",321.85],["W690X323","W",323.42],["W840X329","W",330.48],["W530X332","W",336.76],["W610X341","W",340.69],
  ["W1100X343","W",343.04],["W310X342","W",343.04],["W920X345","W",344.61],["W920X344","W",345.4],["W360X347","W",346.97],["W460X349","W",347.75],
  ["W1000X350","W",350.11],["W760X350","W",350.89],["W690X350","W",351.68],["W840X359","W",360.31],["W920X368","W",367.38],["W1000X371","W",372.09],
  ["W610X372","W",372.09],["W530X369","W",373.66],["W310X375","W",375.23],["W920X381","W",381.51],["W360X382","W",383.08],["W460X384","W",384.65],
  ["W690X384","W",385.43],["W760X389","W",390.14],["W1100X390","W",390.93],["W920X390","W",390.93],["W1000X393","W",391.71],["W840X392","W",391.71],
  ["W1000X412","W",412.91],["W610X415","W",414.48],["W530X409","W",414.48],["W310X415","W",414.48],["W1000X415","W",416.83],["W920X420","W",419.97],
  ["W690X419","W",420.76],["W460X421","W",421.54],["W360X421","W",421.54],["W1100X433","W",432.53],["W840X433","W",433.32],["W760X434","W",435.67],
  ["W1000X438","W",436.46],["W1000X443","W",441.95],["W920X449","W",450.59],["W310X454","W",452.94],["W610X455","W",454.51],["W690X457","W",456.87],
  ["W360X463","W",463.15],["W460X464","W",463.93],["W840X473","W",474.92],["W1000X483","W",482.77],["W1000X486","W",485.91],["W760X484","W",485.91],
  ["W920X491","W",490.62],["W1000X494","W",494.55],["W610X498","W",497.69],["W1100X499","W",498.47],["W310X500","W",500.83],["W690X500","W",502.4],
  ["W360X509","W",511.82],["W840X527","W",526.73],["W760X531","W",531.44],["W1000X539","W",536.94],["W920X537","W",536.94],["W690X548","W",551.85],
  ["W610X551","W",551.85],["W360X551","W",551.85],["W1000X554","W",557.35],["W840X576","W",576.97],["W760X582","W",582.47],["W1000X584","W",587.18],
  ["W920X588","W",587.18],["W1000X591","W",592.67],["W360X592","W",592.67],["W360X634","W",632.71],["W1000X642","W",642.91],["W920X656","W",658.61],
  ["W360X677","W",679.02],["W920X725","W",724.55],["W360X744","W",744.18],["W1000X748","W",749.67],["W920X787","W",792.85],["W690X802","W",808.55],
  ["W360X818","W",824.25],["W1000X883","W",879.2],["W360X900","W",902.75],["W920X970","W",973.4],["W1000X976","W",981.25],["W360X990","W",989.1],
  ["W920X1077","W",1075.45],["W360X1086","W",1091.15],["W920X1194","W",1193.2],["W360X1202","W",1208.9],["W920X1269","W",1271.7],["W360X1299","W",1303.1],
  ["W920X1377","W",1373.75],
  /* UB — 95 added (metric AISC15M / ArcelorMittal labels) */
  ["UB127X76X13","UB",12.97],["UB152X89X16","UB",15.95],["UB178X102X19","UB",19.04],["UB254X102X22","UB",22.0],["UB203X102X23","UB",23.08],["UB305X102X25","UB",24.81],
  ["UB203X133X25","UB",25.1],["UB254X102X25","UB",25.15],["UB305X102X28","UB",28.17],["UB254X102X28","UB",28.32],["UB203X133X30","UB",29.99],["UB254X146X31","UB",31.15],
  ["UB305X102X33","UB",32.84],["UB356X127X33","UB",33.07],["UB254X146X37","UB",37.03],["UB305X127X37","UB",37.04],["UB406X140X39","UB",38.98],["UB356X127X39","UB",39.07],
  ["UB305X165X40","UB",40.29],["UB305X127X42","UB",41.92],["UB254X146X43","UB",42.99],["UB356X171X45","UB",45.0],["UB406X140X46","UB",46.03],["UB305X165X46","UB",46.12],
  ["UB305X127X48","UB",48.07],["UB356X171X51","UB",50.95],["UB457X152X52","UB",52.31],["UB305X165X54","UB",53.98],["UB406X178X54","UB",54.13],["UB356X171X57","UB",56.96],
  ["UB457X152X60","UB",59.84],["UB406X178X60","UB",60.07],["UB356X171X67","UB",67.11],["UB457X191X67","UB",67.13],["UB406X178X67","UB",67.15],["UB457X152X67","UB",67.16],
  ["UB457X152X74","UB",74.17],["UB406X178X74","UB",74.19],["UB457X191X74","UB",74.28],["UB457X152X82","UB",82.03],["UB457X191X82","UB",82.03],["UB533X210X82","UB",82.19],
  ["UB457X191X89","UB",89.33],["UB533X210X92","UB",92.16],["UB457X191X98","UB",98.36],["UB533X210X101","UB",101.03],["UB610X229X101","UB",101.19],["UB533X210X109","UB",109.04],
  ["UB610X229X113","UB",112.96],["UB533X210X122","UB",121.99],["UB610X229X125","UB",125.05],["UB686X254X125","UB",125.21],["UB762X267X134","UB",133.92],["UB610X229X140","UB",139.89],
  ["UB686X254X140","UB",140.04],["UB762X267X147","UB",146.95],["UB610X305X149","UB",149.15],["UB686X254X152","UB",152.37],["UB686X254X170","UB",170.19],["UB762X267X173","UB",173.01],
  ["UB838X292X176","UB",175.84],["UB610X305X179","UB",179.06],["UB838X292X194","UB",193.74],["UB762X267X197","UB",196.72],["UB914X305X201","UB",200.88],["UB1016X305X222","UB",222.0],
  ["UB914X305X224","UB",224.2],["UB838X292X226","UB",226.55],["UB610X305X238","UB",238.09],["UB914X305X238","UB",238.25],["UB1016X305X249","UB",248.69],["UB914X305X253","UB",253.4],
  ["UB914X305X271","UB",271.69],["UB1016X305X272","UB",272.24],["UB914X305X289","UB",289.12],["UB914X305X313","UB",312.74],["UB1016X305X314","UB",314.31],["UB914X419X343","UB",343.28],
  ["UB914X305X345","UB",345.16],["UB1016X305X350","UB",349.4],["UB914X305X381","UB",381.43],["UB914X419X388","UB",387.95],["UB1016X305X393","UB",392.66],["UB1016X305X415","UB",415.03],
  ["UB914X305X425","UB",425.47],["UB1016X305X438","UB",436.7],["UB914X305X474","UB",474.06],["UB1016X305X494","UB",493.84],["UB914X305X521","UB",521.0],["UB914X305X576","UB",575.56],
  ["UB1016X305X584","UB",583.8],["UB1016X305X642","UB",641.82],["UB1016X305X748","UB",748.42],["UB1016X305X883","UB",883.36],["UB1016X305X976","UB",975.52],
  /* UC — 41 added (metric AISC15M / ArcelorMittal labels) */
  ["UC152X152X23","UC",22.96],["UC152X152X30","UC",30.03],["UC152X152X37","UC",36.98],["UC203X203X46","UC",46.1],["UC203X203X52","UC",52.03],["UC203X203X60","UC",59.95],
  ["UC203X203X71","UC",70.99],["UC254X254X73","UC",73.08],["UC203X203X86","UC",86.04],["UC254X254X89","UC",88.94],["UC305X305X97","UC",96.87],["UC254X254X107","UC",107.07],
  ["UC305X305X118","UC",117.91],["UC356X368X129","UC",128.98],["UC254X254X132","UC",131.96],["UC305X305X137","UC",136.9],["UC356X368X153","UC",152.92],["UC305X305X158","UC",158.1],
  ["UC254X254X167","UC",167.13],["UC356X368X177","UC",177.02],["UC305X305X198","UC",198.13],["UC356X368X202","UC",201.9],["UC356X406X235","UC",234.71],["UC305X305X240","UC",240.05],
  ["UC305X305X283","UC",282.91],["UC356X406X287","UC",287.07],["UC356X406X340","UC",339.9],["UC356X406X393","UC",392.97],["UC356X406X467","UC",467.0],["UC356X406X509","UC",509.46],
  ["UC356X406X551","UC",550.99],["UC356X406X592","UC",592.6],["UC356X406X634","UC",633.89],["UC356X406X677","UC",677.77],["UC356X406X744","UC",744.26],["UC356X406X818","UC",818.75],
  ["UC356X406X900","UC",901.96],["UC356X406X990","UC",990.67],["UC356X406X1086","UC",1088.01],["UC356X406X1202","UC",1201.05],["UC356X406X1299","UC",1295.25],
  /* UBP — 17 added (metric AISC15M / ArcelorMittal labels) */
  ["UBP203X203X45","UBP",44.93],["UBP203X203X54","UBP",53.95],["UBP254X254X63","UBP",62.97],["UBP254X254X71","UBP",70.96],["UBP305X305X79","UBP",78.89],["UBP254X254X85","UBP",85.09],
  ["UBP305X305X88","UBP",88.0],["UBP305X305X95","UBP",94.91],["UBP356X368X109","UBP",108.88],["UBP305X305X110","UBP",109.98],["UBP305X305X126","UBP",126.07],["UBP356X368X133","UBP",132.98],
  ["UBP305X305X149","UBP",149.07],["UBP356X368X152","UBP",152.05],["UBP356X368X174","UBP",173.88],["UBP305X305X186","UBP",185.97],["UBP305X305X223","UBP",222.94],
  /* HP — 22 added (metric AISC15M / ArcelorMittal labels) */
  ["HP200X53","HP",53.69],["HP250X62","HP",62.8],["HP310X79","HP",78.5],["HP250X85","HP",84.78],["HP310X93","HP",93.41],["HP360X108","HP",108.33],
  ["HP310X110","HP",110.68],["HP310X125","HP",124.81],["HP410X131","HP",130.31],["HP310X132","HP",131.09],["HP360X132","HP",131.88],["HP410X151","HP",151.5],
  ["HP360X152","HP",152.29],["HP360X174","HP",174.27],["HP410X181","HP",181.33],["HP460X202","HP",201.74],["HP410X211","HP",211.16],["HP460X234","HP",233.93],
  ["HP410X242","HP",241.78],["HP460X269","HP",269.25],["HP410X272","HP",273.96],["HP460X304","HP",304.58],
  /* S-AMER — 28 added (metric AISC15M / ArcelorMittal labels) */
  ["S75X8.5","S-AMER",8.4],["S75X11.2","S-AMER",11.15],["S100X11.5","S-AMER",11.46],["S100X14.1","S-AMER",14.13],["S130X15","S-AMER",14.84],["S150X18.6","S-AMER",18.53],
  ["S150X25.7","S-AMER",25.59],["S200X27.4","S-AMER",27.32],["S200X34","S-AMER",34.23],["S250X37.8","S-AMER",37.76],["S310X47.3","S-AMER",47.18],["S310X52","S-AMER",51.65],
  ["S250X52","S-AMER",52.2],["S310X60.7","S-AMER",60.29],["S380X64","S-AMER",63.82],["S380X74","S-AMER",74.42],["S310X74","S-AMER",74.42],["S460X81.4","S-AMER",80.85],
  ["S510X98.2","S-AMER",98.12],["S460X104","S-AMER",103.62],["S510X112","S-AMER",111.47],["S610X119","S-AMER",119.32],["S510X128","S-AMER",127.95],["S610X134","S-AMER",134.23],
  ["S510X143","S-AMER",142.87],["S610X149","S-AMER",148.36],["S610X158","S-AMER",157.78],["S610X180","S-AMER",179.76],
  /* M-AMER — 18 added (metric AISC15M / ArcelorMittal labels) */
  ["M100X4.3","M-AMER",4.63],["M100X5.1","M-AMER",5.12],["M100X4.8","M-AMER",5.12],["M150X5.5","M-AMER",5.52],["M100X6.1","M-AMER",6.43],["M150X6.6","M-AMER",6.53],
  ["M100X8.9","M-AMER",8.87],["M200X9.2","M-AMER",9.18],["M200X9.7","M-AMER",9.73],["M250X11.2","M-AMER",11.23],["M250X11.9","M-AMER",12.01],["M250X13.4","M-AMER",13.42],
  ["M310X14.9","M-AMER",14.91],["M310X16.1","M-AMER",16.09],["M318X17.3","M-AMER",17.19],["M310X17.6","M-AMER",17.58],["M318X18.5","M-AMER",18.37],["M130X28.1","M-AMER",28.18],
  /* MC — 40 added (metric AISC15M / ArcelorMittal labels) */
  ["MC250X9.7","MC",9.89],["MC150X9.7","MC",9.89],["MC150X10.4","MC",10.6],["MC75X10.6","MC",10.68],["MC250X12.5","MC",12.48],["MC200X12.6","MC",12.64],
  ["MC310X15.8","MC",15.7],["MC150X17.9","MC",17.9],["MC100X20.5","MC",20.41],["MC310X21.3","MC",21.19],["MC150X22.5","MC",22.45],["MC150X22.8","MC",22.76],
  ["MC150X24.3","MC",24.26],["MC150X26.8","MC",26.77],["MC200X27.8","MC",27.87],["MC180X28.4","MC",28.42],["MC200X29.8","MC",29.75],["MC200X31.8","MC",31.79],
  ["MC250X33","MC",32.66],["MC180X33.8","MC",33.75],["MC200X33.9","MC",33.91],["MC230X35.6","MC",35.56],["MC250X37","MC",37.21],["MC230X37.8","MC",37.84],
  ["MC250X42.4","MC",42.39],["MC310X46","MC",46.16],["MC330X47.3","MC",47.34],["MC250X50","MC",50.0],["MC330X52","MC",52.2],["MC310X52","MC",52.2],
  ["MC330X60","MC",59.27],["MC310X60","MC",59.74],["MC250X61.2","MC",61.31],["MC460X63.5","MC",63.82],["MC310X67","MC",66.88],["MC460X68.2","MC",68.37],
  ["MC330X74","MC",74.42],["MC310X74","MC",74.42],["MC460X77.2","MC",77.48],["MC460X86","MC",86.35],
  /* C-AMER — 32 added (metric AISC15M / ArcelorMittal labels) */
  ["C75X5.2","C-AMER",5.52],["C75X6.1","C-AMER",6.08],["C100X6.7","C-AMER",6.79],["C75X7.4","C-AMER",7.44],["C100X8","C-AMER",8.01],["C75X8.9","C-AMER",8.95],
  ["C100X9.3","C-AMER",9.34],["C130X10.4","C-AMER",9.97],["C100X10.8","C-AMER",10.75],["C150X12.2","C-AMER",12.09],["C130X13","C-AMER",13.34],["C180X14.6","C-AMER",14.52],
  ["C150X15.6","C-AMER",15.54],["C200X17.1","C-AMER",17.03],["C180X18.2","C-AMER",18.21],["C150X19.3","C-AMER",19.31],["C230X19.9","C-AMER",19.94],["C200X20.5","C-AMER",20.41],
  ["C180X22","C-AMER",21.9],["C230X22","C-AMER",22.29],["C250X22.8","C-AMER",22.69],["C200X27.9","C-AMER",27.87],["C250X30","C-AMER",29.75],["C230X30","C-AMER",29.75],
  ["C310X30.8","C-AMER",30.77],["C310X37","C-AMER",37.21],["C250X37","C-AMER",37.21],["C310X45","C-AMER",44.59],["C250X45","C-AMER",44.59],["C380X50.4","C-AMER",50.63],
  ["C380X60","C-AMER",59.74],["C380X74","C-AMER",74.42],
  /* L — 137 added (metric AISC15M / ArcelorMittal labels) */
  ["L51X51X3.2","L",2.49],["L51X51X4.8","L",3.66],["L64X38X4.8","L",3.67],["L64X51X4.8","L",4.14],["L64X64X4.8","L",4.56],["L76X51X4.8","L",4.65],
  ["L51X51X6.4","L",4.78],["L64X38X6.4","L",4.8],["L76X64X4.8","L",5.06],["L64X51X6.4","L",5.42],["L76X76X4.8","L",5.52],["L51X51X7.9","L",5.87],
  ["L64X64X6.4","L",6.03],["L76X51X6.4","L",6.08],["L76X64X6.4","L",6.69],["L64X51X7.9","L",6.69],["L51X51X9.5","L",6.94],["L76X76X6.4","L",7.29],
  ["L89X64X6.4","L",7.34],["L64X64X7.9","L",7.39],["L76X51X7.9","L",7.5],["L64X51X9.5","L",7.85],["L89X76X6.4","L",8.01],["L76X64X7.9","L",8.24],
  ["L102X76X6.4","L",8.56],["L89X89X6.4","L",8.63],["L64X64X9.5","L",8.79],["L76X51X9.5","L",8.87],["L89X64X7.9","L",9.03],["L76X76X7.9","L",9.03],
  ["L102X89X6.4","L",9.18],["L127X76X6.4","L",9.81],["L102X102X6.4","L",9.81],["L76X64X9.5","L",9.81],["L89X76X7.9","L",9.89],["L127X89X6.4","L",10.52],
  ["L102X76X7.9","L",10.6],["L89X89X7.9","L",10.6],["L76X76X9.5","L",10.68],["L89X64X9.5","L",10.75],["L76X64X11.1","L",11.23],["L102X89X7.9","L",11.38],
  ["L76X51X12.7","L",11.46],["L64X64X12.7","L",11.46],["L89X76X9.5","L",11.77],["L127X76X7.9","L",12.17],["L102X102X7.9","L",12.17],["L76X76X11.1","L",12.32],
  ["L102X76X9.5","L",12.64],["L89X89X9.5","L",12.64],["L76X64X12.7","L",12.64],["L127X89X7.9","L",12.95],["L89X76X11.1","L",13.5],["L102X89X9.5","L",13.58],
  ["L76X76X12.7","L",13.97],["L89X64X12.7","L",14.05],["L127X76X9.5","L",14.52],["L102X102X9.5","L",14.52],["L152X89X7.9","L",14.6],["L89X89X11.1","L",14.6],
  ["L152X102X7.9","L",15.31],["L89X76X12.7","L",15.31],["L127X89X9.5","L",15.46],["L127X127X7.9","L",15.54],["L102X76X12.7","L",16.48],["L89X89X12.7","L",16.48],
  ["L102X102X11.1","L",16.72],["L127X76X11.1","L",16.8],["L152X89X9.5","L",17.43],["L102X89X12.7","L",17.74],["L152X102X9.5","L",18.29],["L127X127X9.5","L",18.45],
  ["L152X152X7.9","L",18.6],["L127X76X12.7","L",19.0],["L102X102X12.7","L",19.0],["L102X76X15.9","L",20.17],["L178X102X9.5","L",20.25],["L127X89X12.7","L",20.25],
  ["L152X102X11.1","L",21.19],["L127X127X11.1","L",21.35],["L152X152X9.5","L",22.22],["L152X89X12.7","L",22.76],["L102X102X15.9","L",23.31],["L178X102X11.1","L",23.47],
  ["L152X102X12.7","L",24.02],["L127X127X12.7","L",24.26],["L127X89X15.9","L",24.96],["L152X152X11.1","L",25.75],["L203X102X11.1","L",25.9],["L178X102X12.7","L",26.61],
  ["L152X102X14.3","L",26.93],["L102X102X19","L",27.55],["L152X152X12.7","L",29.2],["L203X102X12.7","L",29.36],["L127X89X19","L",29.59],["L152X102X15.9","L",29.67],
  ["L127X127X15.9","L",29.91],["L203X152X11.1","L",30.3],["L152X152X14.3","L",32.66],["L203X102X14.3","L",32.89],["L178X102X15.9","L",32.89],["L203X152X12.7","L",34.46],
  ["L152X102X19","L",35.17],["L127X127X19","L",35.32],["L152X152X15.9","L",36.11],["L203X102X15.9","L",36.27],["L203X152X14.3","L",38.54],["L178X102X19","L",39.17],
  ["L203X203X12.7","L",39.72],["L152X102X22.2","L",40.51],["L127X127X22.2","L",40.51],["L203X152X15.9","L",42.63],["L152X152X19","L",42.86],["L203X102X19","L",43.02],
  ["L203X203X14.3","L",44.43],["L203X203X15.9","L",49.06],["L152X152X22.2","L",49.38],["L203X102X22.2","L",49.61],["L203X152X19","L",50.63],["L152X152X25.4","L",55.73],
  ["L203X102X25.4","L",56.21],["L203X203X19","L",58.25],["L203X152X22.2","L",58.25],["L203X152X25.4","L",66.33],["L203X203X22.2","L",67.35],["L254X254X19","L",73.48],
  ["L203X203X25.4","L",76.46],["L254X254X22.2","L",84.78],["L203X203X28.6","L",84.78],["L254X254X25.4","L",96.55],["L254X254X28.6","L",107.54],["L305X305X25.4","L",116.18],
  ["L254X254X31.8","L",118.53],["L254X254X34.9","L",129.52],["L305X305X28.6","L",130.31],["L305X305X31.8","L",143.66],["L305X305X34.9","L",157.78],
].map(([name, type, kgm]) => ({ name, type, kgm }));

const STEEL_GRADES = [
  // ── European structural (EN 10025-2) ──
  "S185", "S235JR", "S235J0", "S235J2", "S275JR", "S275J0", "S275J2", "S355JR", "S355J0", "S355J2", "S355K2", "S450J0",
  // European fine-grain (EN 10025-3/4)
  "S275N", "S275NL", "S355N", "S355NL", "S420N", "S420NL", "S460N", "S460NL", "S275M", "S355M", "S420M", "S460M",
  // European quenched & weathering (EN 10025-5/6)
  "S460Q", "S460QL", "S690Q", "S690QL", "S235J0W", "S235J2W", "S355J0W", "S355J2W", "S355K2W",
  // ── American (ASTM) ──
  "A36", "A53", "A242", "A500-B", "A500-C", "A501", "A514", "A529-50", "A572-42", "A572-50", "A572-55", "A572-60", "A572-65", "A588", "A618", "A653", "A709-36", "A709-50", "A709-50W", "A847", "A913-50", "A913-65", "A913-70", "A992",
  // ── Chinese (GB) ──
  "Q195", "Q215", "Q235B", "Q235C", "Q235D", "Q275", "Q345B", "Q345C", "Q345D", "Q355B", "Q355C", "Q355D", "Q390", "Q420", "Q460",
  // ── Japanese (JIS) ──
  "SS400", "SS490", "SS540", "SM400A", "SM400B", "SM490A", "SM490B", "SM490YA", "SM490YB", "SM520B", "SM520C", "SM570", "SN400A", "SN400B", "SN490B", "SMA400", "SMA490",
  // ── Canadian (CSA) & other common ──
  "350W", "350WT", "ST37-2", "ST52-3", "Fe360", "Fe430", "Fe510"
];
const GRADE_GROUPS = [
  ["European (EN 10025)", ["S185", "S235JR", "S235J0", "S235J2", "S275JR", "S275J0", "S275J2", "S355JR", "S355J0", "S355J2", "S355K2", "S450J0", "S275N", "S275NL", "S355N", "S355NL", "S420N", "S420NL", "S460N", "S460NL", "S275M", "S355M", "S420M", "S460M", "S460Q", "S460QL", "S690Q", "S690QL", "S235J0W", "S235J2W", "S355J0W", "S355J2W", "S355K2W"]],
  ["American (ASTM)", ["A36", "A53", "A242", "A500-B", "A500-C", "A501", "A514", "A529-50", "A572-42", "A572-50", "A572-55", "A572-60", "A572-65", "A588", "A618", "A653", "A709-36", "A709-50", "A709-50W", "A847", "A913-50", "A913-65", "A913-70", "A992"]],
  ["Chinese (GB)", ["Q195", "Q215", "Q235B", "Q235C", "Q235D", "Q275", "Q345B", "Q345C", "Q345D", "Q355B", "Q355C", "Q355D", "Q390", "Q420", "Q460"]],
  ["Japanese (JIS)", ["SS400", "SS490", "SS540", "SM400A", "SM400B", "SM490A", "SM490B", "SM490YA", "SM490YB", "SM520B", "SM520C", "SM570", "SN400A", "SN400B", "SN490B", "SMA400", "SMA490"]],
  ["Other (CSA / DIN)", ["350W", "350WT", "ST37-2", "ST52-3", "Fe360", "Fe430", "Fe510"]],
];
function gradeMatches(q) {
  const s = (q || "").toUpperCase().trim();
  return soGradeGroups(GRADE_GROUPS).map(([grp, list]) => [grp, s ? list.filter(g => g.toUpperCase().includes(s)) : list]).filter(([, list]) => list.length);
}

/* ─── SECTION HELPERS ────────────────────────────────────────────────────── */
function isExcludedProfile(name) {
  if (!name) return false;
  const n = String(name).trim().toUpperCase().replace(/\s+/g, "");
  return /^(PL|PLT|PLAT|PLATE|FLT|FL|FB|FLAT|SHEET|SHT|GUSSET|GPL|BPL|CHK|CHEQ|CLEAT|STIFF|CAPPL|BASEPL)(?=$|[\d*x×.\-_/])/i.test(n) || /\bPLATE\b/i.test(name);
}
function hollowKgm(type, d) {
  const t = d.t; if (!t || t <= 0) return null;
  if (type === "CHS") { const D = d.a; if (!D) return null; return Math.round(Math.PI * (D - t) * t * 1e-6 * DENSITY * 100) / 100; }
  const H = d.a, B = d.b || d.a; if (!H || !B) return null;
  const area = (H * B) - ((H - 2 * t) * (B - 2 * t)) - (4 - Math.PI) * (2 * t) * (2 * t);
  return Math.round(area * 1e-6 * DENSITY * 100) / 100;
}
function normalizeHollow(raw) {
  let n = String(raw).trim().toUpperCase().replace(/×/g, "X").replace(/\*/g, "X").replace(/\s+/g, "");
  n = n.replace(/^(CF|HF|CFC|HFC)(RHS|SHS|CHS)/, "$2");
  let m;
  if ((m = n.match(/^SHS(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)(?:X(\d+(?:\.\d+)?))?$/))) { const a = +m[1], t = m[3] !== undefined ? +m[3] : +m[2]; return { type: "SHS", canonical: `SHS ${a}x${a}x${t}`, kgm: hollowKgm("SHS", { a, b: a, t }) }; }
  if ((m = n.match(/^RHS(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)$/))) { const fd = { a: +m[1], b: +m[2], t: +m[3] }; return { type: "RHS", canonical: `RHS ${fd.a}x${fd.b}x${fd.t}`, kgm: hollowKgm("RHS", fd) }; }
  if ((m = n.match(/^(?:CHS|PIPE|DIA|Ø|O|D)(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)$/))) { const fd = { a: +m[1], t: +m[2] }; return { type: "CHS", canonical: `CHS ${fd.a}x${fd.t}`, kgm: hollowKgm("CHS", fd) }; }
  return null;
}
function normalizeAngle(raw) {
  const n = String(raw).trim().toUpperCase().replace(/×/g, "X").replace(/\*/g, "X").replace(/\s+/g, "");
  let m;
  if ((m = n.match(/^[LA](\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)(?:X(\d+(?:\.\d+)?))?$/))) {
    const a = +m[1], b = m[3] !== undefined ? +m[2] : a, t = m[3] !== undefined ? +m[3] : +m[2];
    if (!(a > 0 && b > 0 && t > 0) || t > Math.min(a, b)) return null;
    return { type: "L", canonical: `L ${a}x${b}x${t}`, kgm: Math.round(t * (a + b - t) * 1e-6 * DENSITY * 100) / 100 };
  }
  return null;
}
function aliasToCanonical(raw) {
  const n = String(raw).trim().toUpperCase().replace(/\s+/g, ""); let m;
  if ((m = n.match(/^HE(\d+)(A|B|M)$/))) return `HE${m[2]} ${m[1]}`;
  if ((m = n.match(/^(HEA|HEB|HEM)(\d+)$/))) return `${m[1]} ${m[2]}`;
  if ((m = n.match(/^(IPE|IPN|UB|UC|HW|HM|HN|UPN|UPE|PFC|UBP)(\d.*)$/))) return `${m[1]} ${m[2]}`;
  return null;
}
function parseBuiltUp(name) {
  if (!name) return null;
  const m = name.trim().match(/^(?:BU|PG)?\s*(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)$/i);
  if (!m) return null;
  const D = +m[1], Bf = +m[2], tf = +m[3], tw = +m[4];
  if (!(D > 0 && Bf > 0 && tf > 0 && tw > 0) || tf >= D / 2) return null;
  const kgm = ((2 * Bf * tf + (D - 2 * tf) * tw) * 1e-6) * DENSITY;
  return { kgm: Math.round(kgm * 10) / 10, D, Bf, tf, tw, builtUp: true };
}
/* ════════════════════════════════════════════════════════════════════════
   IMPERIAL / FRACTION PARSING
   ------------------------------------------------------------------------
   US and Canadian cut lists write sizes as inch fractions: HSS8X8X1/2,
   L4X4X3/8, HSS 10 x 6 x 5/16, sometimes with unicode glyphs (HSS8X8X½) or
   mixed numbers (PIPE 2 1/2 STD). None of that matches a metric label, so the
   row lands on the user to fix by hand — which is exactly the friction this
   product exists to remove.

   Strategy: expand the fraction, convert inches to millimetres, rebuild the
   metric label and look it up. Because the metric AISC15M rows are already in
   STEEL_DB, HSS8X8X1/2 resolves to the real catalogue value for
   HSS203.2X203.2X12.7 rather than to a computed approximation.

   Deciding inch vs mm is done by magnitude, not by guessing intent: hollow and
   angle sections below 60 units are inches (a 200 mm SHS exists, a 200 inch one
   does not; an 8 inch HSS exists, an 8 mm one does not). Anything at or above
   that threshold is left alone as metric.
════════════════════════════════════════════════════════════════════════ */
const UNI_FRAC = { "½": "1/2", "⅓": "1/3", "⅔": "2/3", "¼": "1/4", "¾": "3/4",
  "⅕": "1/5", "⅖": "2/5", "⅗": "3/5", "⅘": "4/5", "⅙": "1/6", "⅚": "5/6",
  "⅐": "1/7", "⅛": "1/8", "⅜": "3/8", "⅝": "5/8", "⅞": "7/8", "⅑": "1/9", "⅒": "1/10" };

/* "1/2" -> 0.5 · "2 1/2" -> 2.5 · "2-1/2" -> 2.5 · "3/8" -> 0.375 · "½" -> 0.5
   Decimals and plain integers pass through untouched. */
function expandFractions(raw) {
  let v = String(raw == null ? "" : raw);
  for (const g in UNI_FRAC) v = v.split(g).join(" " + UNI_FRAC[g]);   // ½ -> " 1/2"
  // mixed number first: 2 1/2 or 2-1/2  (the hyphen form must not be read as minus)
  v = v.replace(/(\d+)\s*[-\s]\s*(\d+)\s*\/\s*(\d+)/g, (m, w, a, b) => {
    const d = +b; if (!d) return m;
    const val = +w + (+a) / d;
    return String(Math.round(val * 1e6) / 1e6);
  });
  // bare fraction: 1/2, 5/16
  v = v.replace(/(?:(?<=^)|(?<=[^\d.]))(\d+)\s*\/\s*(\d+)/g, (m, a, b) => {
    const d = +b; if (!d) return m;
    return String(Math.round(((+a) / d) * 1e6) / 1e6);
  });
  return v.replace(/\s+/g, " ").trim();
}

const IN2MM = 25.4;
const trimNum = x => {
  const r = Math.round(x * 10) / 10;
  return (Math.abs(r - Math.round(r)) < 1e-9 ? String(Math.round(r)) : String(r));
};

/* ── TELLING INCHES FROM MILLIMETRES ────────────────────────────────────────
   This is the part that must not guess, because a wrong call is not a small
   error: read "SHS 45x45x3" as inches and it becomes 1143 mm, and the tonnage
   is out by a factor of 25.

   The overall size is useless as a discriminator — 45 could be 45 mm or 45 in.
   WALL THICKNESS is decisive:
       metric walls are quoted in mm  → 1.5 … 30
       imperial walls are quoted in in → 0.125 … 1.0   (1/8" to 1")
   A wall under 1.2 is therefore inches, and a wall of 2 or more is millimetres.
   The overlap (1.2 … 2.0, e.g. cold-formed SHS 25x25x1.5 vs a 1.5 in wall) is
   resolved by naming convention, not by magnitude:
       HSS / TS   → American designations, inches
       SHS / RHS  → metric designations, millimetres
   and either way an explicit fraction or inch mark settles it outright.      */
const RE_FRACTION = /\d\s*\/\s*\d|[½⅓⅔¼¾⅕⅖⅗⅘⅙⅚⅐⅛⅜⅝⅞⅑⅒]/;
const RE_INCHMARK = /["”″]|\b(?:IN|INCH|INCHES)\b/i;

function unitOf(prefix, dims, raw) {
  const t = dims[dims.length - 1];                 // wall thickness
  if (RE_FRACTION.test(raw)) return "in";          // 1/2, 5/16, ½ — never metric
  if (RE_INCHMARK.test(raw)) return "in";          // 8" x 8" x 0.5"
  if (/\bMM\b/i.test(raw)) return "mm";            // stated explicitly
  const american = prefix === "HSS" || prefix === "TS";
  if (t < 1.2) return american ? "in" : "in";      // no steel section has a 1 mm wall
  if (t >= 2) return "mm";                         // no imperial wall is 2 in
  return american ? "in" : "mm";                   // 1.2-2.0 overlap → by convention
}

/* Rebuild an imperial hollow/angle designation as its metric label.
   Returns null for metric input, so metric names are never touched. */
function imperialToMetricLabel(raw) {
  if (!raw) return null;
  const src = String(raw);
  const n = expandFractions(src).toUpperCase()
    .replace(/×/g, "X").replace(/[*∗]/g, "X").replace(/["”″]/g, "").replace(/\s+/g, "");

  let m = n.match(/^(HSS|TS|SHS|RHS)(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)(?:X(\d+(?:\.\d+)?))?$/);
  if (m) {
    const pre = m[1];
    const dims = [m[2], m[3], m[4]].filter(v => v !== undefined).map(Number);
    if (dims.some(v => !(v > 0))) return null;
    if (unitOf(pre, dims, src) !== "in") return null;
    if (Math.max.apply(null, dims) > 40) return null;   // >40 in exceeds any rolled HSS
    return "HSS" + dims.map(v => trimNum(v * IN2MM)).join("X");
  }
  m = n.match(/^L(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)$/);
  if (m) {
    const dims = [+m[1], +m[2], +m[3]];
    if (dims.some(v => !(v > 0))) return null;
    if (unitOf("L", dims, src) !== "in") return null;
    if (Math.max(dims[0], dims[1]) > 12) return null;   // >12 in exceeds any rolled angle
    return "L" + dims.map(v => trimNum(v * IN2MM)).join("X");
  }
  return null;
}

/* ── Tekla parametric profile names — isolated add-on ────────────────────────
   Tekla material lists name cold-formed and hollow profiles parametrically:
     P h*b*t      rectangular / square hollow   → SHS / RHS (same catalogue or computed value as "SHS h x b x t")
     PD d*t       circular hollow               → CHS d x t
     ZZ h-t-c-b   lipped Z purlin (…-b2 when the flanges differ)   A = t(h + b1 + b2 + 2c − 4t)
     C h*b*t      plain cold-formed channel, written with *        A = t(h + 2b − 2t)
   Paint area for C / ZZ = both faces of the wall, sharp corners (2h + 4b · 2(h + b1 + b2 + 2c)).
   When the file carries a weight column, the file's kg/m replaces these computed values.
   Asked only after the library search found nothing, so no existing name changes.
   TEKLA_PARAM_ENABLED = false → switched off. */
const TEKLA_PARAM_ENABLED = true;   // ◄ Tekla parametric names on/off
function teklaParamSection(raw) {
  if (!TEKLA_PARAM_ENABLED || raw == null) return null;
  try {
    const n = String(raw).trim().toUpperCase().replace(/\s+/g, "");
    const N = "(\\d+(?:\\.\\d+)?)";
    let m;
    if ((m = n.match(new RegExp(`^P${N}[*X×]${N}[*X×]${N}$`)))) {
      const h = +m[1], b = +m[2], t = +m[3];
      if (!(h > 0 && b > 0 && t > 0 && t < Math.min(h, b) / 2)) return null;
      return findSection(`${h === b ? "SHS" : "RHS"} ${h}x${b}x${t}`) || null;
    }
    if ((m = n.match(new RegExp(`^PD${N}[*X×]${N}$`)))) {
      const D = +m[1], t = +m[2];
      if (!(D > 0 && t > 0 && t < D / 2)) return null;
      return findSection(`CHS ${D}x${t}`) || null;
    }
    if ((m = n.match(new RegExp(`^ZZ${N}-${N}-${N}-${N}(?:-${N})?$`)))) {
      const h = +m[1], t = +m[2], c = +m[3], b1 = +m[4], b2 = m[5] != null ? +m[5] : b1;
      if (!(h > 0 && t > 0 && c >= 0 && b1 > 0 && b2 > 0 && t < Math.min(h, b1, b2) / 4)) return null;
      const A = t * (h + b1 + b2 + 2 * c - 4 * t);
      return { name: n, type: "Z-COLD", kgm: Math.round(A * 1e-6 * DENSITY * 100) / 100, computed: true, tekla: true, pm: 2 * (h + b1 + b2 + 2 * c), basis: "thin" };
    }
    if ((m = n.match(new RegExp(`^C${N}\\*${N}\\*${N}$`)))) {
      const h = +m[1], b = +m[2], t = +m[3];
      if (!(h > 0 && b > 0 && t > 0 && t < Math.min(h, b) / 3)) return null;
      const A = t * (h + 2 * b - 2 * t);
      return { name: n, type: "C-COLD", kgm: Math.round(A * 1e-6 * DENSITY * 100) / 100, computed: true, tekla: true, pm: 2 * h + 4 * b, basis: "thin" };
    }
  } catch (e) { try { console.warn("[tekla-param add-on]", e); } catch { /* no console */ } }
  return null;
}
// Plates reader: names that are NOT plates (hollow P h*b*t, PD, ZZ, round bar D, bolts/nuts/washers, bare a*b over 100 mm).
function soTeklaNonPlate(s) {
  if (!TEKLA_PARAM_ENABLED) return false;
  const n = String(s || "").trim().toUpperCase().replace(/\s+/g, "");
  const m = n.match(/^P(\d+(?:\.\d+)?)[*X×](\d+(?:\.\d+)?)[*X×](\d+(?:\.\d+)?)$/);
  if (m && +m[3] > 0 && +m[3] < Math.min(+m[1], +m[2]) / 2) return true;
  const bare = n.match(/^(\d+(?:\.\d+)?)[*X×](\d+(?:\.\d+)?)$/);        // "2000*700", "120*150": solid / concrete sections, not plates
  if (bare && Math.min(+bare[1], +bare[2]) > 100) return true;
  return /^PD\d/.test(n) || /^ZZ\d/.test(n) || /^D\d+(?:\.\d+)?$/.test(n) || /(?:^|[^A-Z])(?:NUT|BOLT|WASHER|SCREW|ANCHOR|STUD|RIVET)(?:[^A-Z]|$)/.test(n);
}
// Plates reader: the first sheet that has a header row (exports often start with a cover sheet).
function soPlateSheetRows(wb) {
  const HEAD = /profile|length|width|qty|no\.?$|thick|partpos|part pos/i;
  let first = null;
  for (const sn of wb.SheetNames) {
    const raw = intlPrepSheet(arPrepSheet(XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "" })));
    if (first === null) first = raw;
    for (let i = 0; i < Math.min(raw.length, 25); i++) { const r = (raw[i] || []).map(c => normHeader(c).toLowerCase()); if (r.some(c => HEAD.test(c))) return raw; }
  }
  return first || [];
}
// Sections reader: rows that could not be read are listed after the upload instead of vanishing.
function soUnknownAdd(map, name, qty, gradeCell, kg) {
  try {
    const n = String(name == null ? "" : name).trim(); if (!n) return;
    if (/(?:^|[^A-Z])(?:NUT|BOLT|WASHER|SCREW|ANCHOR|STUD|RIVET)(?:[^A-Z]|$)/i.test(n)) return;      // fasteners are not cut from bars
    if (/concrete|beton|бетон|混凝土|خرسان/i.test(String(gradeCell == null ? "" : gradeCell))) return;   // concrete members
    const q = Math.max(1, Math.round(+qty || 1)), e = map.get(n) || { qty: 0, kg: 0 };
    e.qty += q; if (kg > 0 && isFinite(kg)) e.kg += kg;
    map.set(n, e);
  } catch { /* never breaks the reader */ }
}
function soUnknownList(map) { return [...map.entries()].map(([name, e]) => ({ name, qty: e.qty, kg: e.kg || 0 })).sort((a, b) => b.qty - a.qty); }
/* Weight in the file (Tekla prints it). Only headers that say what they are are used:
   per piece ("…for one", "each", "unit weight"), per row ("…for all", "total weight") or kg/m.
   A bare "Weight" column is ambiguous (piece or row?) and is not guessed. */
function soWeightCols(cells) {
  const W = "w(?:ei|ie)ght|\\bwt\\b|\\bmass\\b";
  const find = re => cells.findIndex(c => re.test(c));
  const kgmCol = find(/kg\s*\/\s*m\b|kg per m\b|(?:mass|w(?:ei|ie)ght) per (?:m|metre|meter)\b|unit mass/);
  const w1 = find(new RegExp(`(?:${W}).*(?:for one|each|per (?:piece|pc|pcs|item|part)|single|/\\s*pc)|(?:unit|single|piece) (?:${W})`));
  const wAll = find(new RegExp(`(?:${W}).*(?:for all|total)|total (?:${W})`));
  const unit = i => (i >= 0 && /\(\s*t\s*\)|\bton|tonne/.test(cells[i]) ? 1000 : 1);
  return { kgmCol, w1, wAll, w1f: unit(w1), wAllf: unit(wAll) };
}
function soRowKg(r, cm, qty, num) {                             // kg of the whole row, or null
  try {
    if (cm.wAll >= 0) { const w = num(r[cm.wAll]); if (w > 0) return w * cm.wAllf; }
    if (cm.w1 >= 0) { const w = num(r[cm.w1]); if (w > 0) return w * cm.w1f * Math.max(1, qty || 1); }
  } catch { /* ignore */ }
  return null;
}
function soRowKgm(r, cm, length, qty, num) {                    // kg/m from the file, or null
  try {
    const L = length / 1000; if (!(L > 0)) return null;
    let v = null;
    if (cm.kgmCol >= 0) v = num(r[cm.kgmCol]);
    if (!(v > 0)) { const kg = soRowKg(r, cm, qty, num); v = kg > 0 ? kg / (L * Math.max(1, qty || 1)) : null; }
    return v > 0 && v < 5000 && isFinite(v) ? Math.round(v * 1000) / 1000 : null;
  } catch { return null; }
}
function soFileKgm(items) {                                     // length-weighted kg/m of a profile group
  let w = 0, l = 0;
  for (const c of items || []) { const f = +c.fkgm, L = (+c.length || 0) * (+c.qty || 0); if (f > 0 && L > 0) { w += f * L; l += L; } }
  return l > 0 ? Math.round(w / l * 100) / 100 : null;
}
/* Library value kept for catalogue sections (the file only cross-checks them);
   a computed/approximate value is replaced by the file's own kg/m. */
function soKgm(sec, fk) {
  if (sec && !sec.computed) return { kgm: sec.kgm, fromFile: false, diff: fk > 0 && sec.kgm > 0 ? (fk - sec.kgm) / sec.kgm : null };
  if (sec && fk > 0) return { kgm: fk, fromFile: true, diff: null };
  return { kgm: sec ? sec.kgm : null, fromFile: false, diff: null };
}
// Results: the name exactly as written in the file, when it differs from the library name (L3X3X1/4 → L76X76X6.4).
function soSrcName(raw, canon) {
  try {
    const tok = s => (String(s || "").toUpperCase().replace(/[×*]/g, "X").match(/[A-ZЀ-ӿ]+|\d+(?:\.\d+)?/g) || []).map(x => /^\d/.test(x) ? String(+x) : x).sort().join(" ");
    const a = tok(raw), b = tok(canon);
    return a && b && a !== b ? String(raw).trim() : null;
  } catch { return null; }
}
Object.assign(LANG_DICT, {
  unknownRows: { en: "⚠ Not recognised as a section — not optimised: {list}. Check these names, or email us the file.",
                 ar: "⚠ لم يُتعرَّف عليها كمقاطع — لم تدخل الحساب: {list}. راجع هذه الأسماء أو أرسل لنا الملف." },
  fromFile:    { en: "as written in your file: {name}", ar: "كما في ملفك: {name}" },
  fromFileTag: { en: " (weight from your file)", ar: " (الوزن من ملفك)" },
  fileKgmNote: { en: "file: {v} kg/m ({d}%)", ar: "في الملف: {v} كجم/م ({d}%)" },
  fileKgmWarn: { en: "⚠ file: {v} kg/m ({d}%) — check this section", ar: "⚠ في الملف: {v} كجم/م ({d}%) — راجع هذا المقطع" },
  fileKgmCell: { en: "file {v}", ar: "بالملف {v}" },
});
/* ╚══ end of Tekla parametric add-on ═══════════════════════════════════════╝ */

function findSection(name) {
  if (!name || isExcludedProfile(name)) return null;
  const n = name.trim().toUpperCase().replace(/\s+/g, " ");
  let hit = STEEL_DB.find(s => s.name.toUpperCase() === n);
  if (hit) return hit;
  const compact = n.replace(/\s+/g, "");
  hit = STEEL_DB.find(s => s.name.toUpperCase().replace(/\s+/g, "") === compact);
  if (hit) return hit;
  // separator-agnostic match: unify -, ×, * to x and drop spaces (handles Tekla "H-200x200x8x12", "I-300x150x10x18.5")
  const norm = v => v.toUpperCase().replace(/×/g, "X").replace(/[*∗]/g, "X").replace(/[-_]/g, "").replace(/\s+/g, "");
  const nn = norm(name);
  hit = STEEL_DB.find(s => norm(s.name) === nn);
  if (hit) return hit;
  // UK sections: DB stores "406x178x74 UB" but Tekla/Advance Steel write "UB 406x178x74" / "UB406x178x74".
  const uk = nn.match(/^(UB|UC|UBP|PFC)(\d+X\d+X[\d.]+)$/);
  if (uk) { const want = uk[2] + uk[1]; hit = STEEL_DB.find(s => norm(s.name) === want); if (hit) return hit; }
  // Tekla often prefixes JIS H-shapes with "H" instead of HW/HM/HN — match on leading HxB.
  const hm = nn.match(/^H(\d+X\d+)/);
  if (hm) { const lead = hm[1]; hit = STEEL_DB.find(s => { const sn = norm(s.name); return /^H[WMN]/.test(sn) && sn.slice(2).startsWith(lead) && (sn.slice(2) === lead || sn.slice(2 + lead.length).startsWith("X") || sn.slice(2) === lead); }); if (hit) return hit; }
  const alias = aliasToCanonical(name);
  if (alias) { const a = STEEL_DB.find(s => s.name.toUpperCase().replace(/\s+/g, "") === alias.replace(/\s+/g, "")); if (a) return a; }
  const h = normalizeHollow(name); if (h && h.kgm) return STEEL_DB.find(s => s.name.toUpperCase().replace(/\s+/g, "") === h.canonical.toUpperCase().replace(/\s+/g, "")) || { name: h.canonical, type: h.type, kgm: h.kgm, computed: true };
  const ang = normalizeAngle(name); if (ang && ang.kgm) return { name: ang.canonical, type: ang.type, kgm: ang.kgm, computed: true };
  const bu = parseBuiltUp(name); if (bu) return { name: name.trim().toUpperCase(), type: "BUILT-UP", kgm: bu.kgm, builtUp: true };

  // ── fraction / imperial fallbacks — tried last, so nothing above changes ──
  // 1) imperial designation rebuilt as its metric label, then looked up
  const impLabel = imperialToMetricLabel(name);
  if (impLabel) {
    const norm2 = v => v.toUpperCase().replace(/\s+/g, "");
    const exact = STEEL_DB.find(s => norm2(s.name) === impLabel);
    if (exact) return exact;
    // nearest catalogue row within rounding tolerance of the inch conversion
    const want = impLabel.replace(/^HSS|^L/, "").split("X").map(Number);
    const pre = impLabel.startsWith("HSS") ? "HSS" : "L";
    let best = null, bestErr = Infinity;
    for (const sec of STEEL_DB) {
      const sn = norm2(sec.name);
      if (!sn.startsWith(pre)) continue;
      const got = sn.slice(pre.length).split("X").map(Number);
      if (got.length !== want.length || got.some(isNaN)) continue;
      let err = 0;
      for (let i = 0; i < want.length; i++) err = Math.max(err, Math.abs(got[i] - want[i]));
      if (err < bestErr) { bestErr = err; best = sec; }
    }
    if (best && bestErr <= 0.8) return best;      // 0.8 mm covers inch rounding
  }
  // 2) plain fraction expansion, in case the rest of the name is already metric
  const expanded = expandFractions(name);
  if (expanded && expanded.toUpperCase() !== String(name).trim().toUpperCase()) {
    const h2 = normalizeHollow(expanded);
    if (h2 && h2.kgm) return { name: h2.canonical, type: h2.type, kgm: h2.kgm, computed: true };
    const a2 = normalizeAngle(expanded);
    if (a2 && a2.kgm) return { name: a2.canonical, type: a2.type, kgm: a2.kgm, computed: true };
    const b2 = parseBuiltUp(expanded);
    if (b2) return { name: expanded.toUpperCase(), type: "BUILT-UP", kgm: b2.kgm, builtUp: true };
  }
  const tk = teklaParamSection(name); if (tk) return tk;   // Tekla parametric names: isolated add-on
  return intlFind(name);   // international SAP2000 sections: isolated add-on, asked last
}
function searchSections(q, limit = 12) {
  if (!q) return [];
  const Q = q.trim().toUpperCase().replace(/\s+/g, " ");
  const starts = [], contains = [];
  for (const s of STEEL_DB) { const up = s.name.toUpperCase(); if (up.startsWith(Q)) starts.push(s); else if (up.replace(/\s/g, "").includes(Q.replace(/\s/g, ""))) contains.push(s); }
  const found = [...starts, ...contains];
  return (found.length < limit ? found.concat(intlSearch(q, limit - found.length)) : found).slice(0, limit);   // international rows: isolated add-on
}

/* ╔══════════════════════════════════════════════════════════════════════════╗
   ║  INTERNATIONAL SECTIONS — SAP2000 databases (isolated add-on)             ║
   ╚══════════════════════════════════════════════════════════════════════════╝
   Russia (Russian2020.pro), China (ChineseGB08.pro), Japan (JIS G 3192:2014)
   and India (Indian.pro), read from the SAP2000 section-property files:
   h, b, tw, tf / legs / wall and the cross-section AREA. Mass = area × 7850
   kg/m³ (how GOST, GB, JIS and IS tables compute theoretical mass).
   Rows whose own numbers disagree (area vs. dimensions) were dropped, not
   repaired; double angles/channels (2L, 2C) are assemblies of the singles
   and are not listed.

   How it is wired — nothing that worked before can change:
     • the rows live in their own table (not in STEEL_DB), so the existing
       lookup, the section pages and build-pages.mjs never see them;
     • findSection() asks this add-on ONLY after its own search found nothing;
     • names are written so that the old search never resolves them
       (20Б1 · 20П · L75×6 ГОСТ 8509-93 · □100×100×5 ГОСТ 32931-2015 · I20a ·
        [20a · ∠75×6 · HW200×200×8×12 · JIS H… · ISMB 200 …).
   Native designations are understood too: «Двутавр 20Б1 ГОСТ Р 57837-2017»,
   «Швеллер 20П», «Уголок 75х6», «Труба 100х100х5 ГОСТ 30245-2003»,
   工字钢I20a, 槽钢[20a, 角钢∠75×6, 方管□100×100×5, Φ108×4,
   «Perfil IPE 300», «Tubo cuadrado 100x100x5», «Angular 50x50x5» …

   Switch: INTL_SECTIONS_ENABLED = false → the groups leave the Section Library
   and these names are no longer recognised.
──────────────────────────────────────────────────────────────────────────── */
const INTL_SECTIONS_ENABLED = true;   // ◄ international SAP2000 sections on/off

const INTL_REGION = {"RU": {"en": "Russia — GOST · SAP2000 section database", "ar": "روسيا — GOST · قاعدة مقاطع SAP2000", "ru": "Россия — ГОСТ · база сечений SAP2000", "zh": "俄罗斯 — GOST · SAP2000 截面库", "es": "Rusia — GOST · base de secciones de SAP2000"}, "CN": {"en": "China — GB · SAP2000 section database", "ar": "الصين — GB · قاعدة مقاطع SAP2000", "ru": "Китай — GB · база сечений SAP2000", "zh": "中国 — GB · SAP2000 截面库", "es": "China — GB · base de secciones de SAP2000"}, "JP": {"en": "Japan — JIS · SAP2000 section database", "ar": "اليابان — JIS · قاعدة مقاطع SAP2000", "ru": "Япония — JIS · база сечений SAP2000", "zh": "日本 — JIS · SAP2000 截面库", "es": "Japón — JIS · base de secciones de SAP2000"}, "IN": {"en": "India — IS · SAP2000 section database", "ar": "الهند — IS · قاعدة مقاطع SAP2000", "ru": "Индия — IS · база сечений SAP2000", "zh": "印度 — IS · SAP2000 截面库", "es": "India — IS · base de secciones de SAP2000"}};
const INTL_FAMS = [
  ["RU-B", "RU", "I", "ГОСТ Р 57837-2017", "Б", {"en": "Normal I-beams Б (GOST R 57837-2017)", "ar": "كمرات I عادية Б (GOST R 57837-2017)", "ru": "Двутавры нормальные Б (ГОСТ Р 57837-2017)", "zh": "普通工字钢 Б（GOST R 57837-2017）", "es": "Vigas I normales Б (GOST R 57837-2017)"}, null, ".", "Russian2020.pro"],
  ["RU-SH", "RU", "I", "ГОСТ Р 57837-2017", "Ш", {"en": "Wide-flange I-beams Ш (GOST R 57837-2017)", "ar": "كمرات I عريضة الشفة Ш (GOST R 57837-2017)", "ru": "Двутавры широкополочные Ш (ГОСТ Р 57837-2017)", "zh": "宽翼缘工字钢 Ш（GOST R 57837-2017）", "es": "Vigas I de ala ancha Ш (GOST R 57837-2017)"}, null, ".", "Russian2020.pro"],
  ["RU-K", "RU", "I", "ГОСТ Р 57837-2017", "К", {"en": "Column I-beams К (GOST R 57837-2017)", "ar": "كمرات أعمدة К (GOST R 57837-2017)", "ru": "Двутавры колонные К (ГОСТ Р 57837-2017)", "zh": "柱用工字钢 К（GOST R 57837-2017）", "es": "Perfiles I para columnas К (GOST R 57837-2017)"}, null, ".", "Russian2020.pro"],
  ["RU-S", "RU", "I", "ГОСТ Р 57837-2017", "С", {"en": "Pile I-beams С (GOST R 57837-2017)", "ar": "كمرات خوازيق С (GOST R 57837-2017)", "ru": "Двутавры свайные С (ГОСТ Р 57837-2017)", "zh": "桩用工字钢 С（GOST R 57837-2017）", "es": "Perfiles I para pilotes С (GOST R 57837-2017)"}, null, ".", "Russian2020.pro"],
  ["RU-DB", "RU", "I", "ГОСТ Р 57837-2017", "ДБ", {"en": "Additional-series beams ДБ (GOST R 57837-2017)", "ar": "كمرات السلسلة الإضافية ДБ (GOST R 57837-2017)", "ru": "Двутавры дополнительной серии ДБ (ГОСТ Р 57837-2017)", "zh": "补充系列梁用工字钢 ДБ（GOST R 57837-2017）", "es": "Vigas de serie adicional ДБ (GOST R 57837-2017)"}, null, ".", "Russian2020.pro"],
  ["RU-DK", "RU", "I", "ГОСТ Р 57837-2017", "ДК", {"en": "Additional-series columns ДК (GOST R 57837-2017)", "ar": "أعمدة السلسلة الإضافية ДК (GOST R 57837-2017)", "ru": "Двутавры дополнительной серии ДК (ГОСТ Р 57837-2017)", "zh": "补充系列柱用工字钢 ДК（GOST R 57837-2017）", "es": "Columnas de serie adicional ДК (GOST R 57837-2017)"}, null, ".", "Russian2020.pro"],
  ["RU-8239", "RU", "I", "ГОСТ 8239-89", "I", {"en": "Tapered-flange I-beams (GOST 8239-89)", "ar": "كمرات I مائلة الشفة (GOST 8239-89)", "ru": "Двутавры с уклоном граней полок (ГОСТ 8239-89)", "zh": "斜翼缘工字钢（GOST 8239-89）", "es": "Vigas I de alas inclinadas (GOST 8239-89)"}, null, ".", "Russian2020.pro"],
  ["RU-STO93", "RU", "I", "СТО АСЧМ 20-93", "СТО", {"en": "I-beams STO ASChM 20-93 (Б, Ш, К, М)", "ar": "كمرات I وفق STO ASChM 20-93 (Б، Ш، К، М)", "ru": "Двутавры СТО АСЧМ 20-93 (Б, Ш, К, М)", "zh": "工字钢 STO ASChM 20-93（Б、Ш、К、М）", "es": "Vigas I STO ASChM 20-93 (Б, Ш, К, М)"}, null, ".", "Russian2020.pro"],
  ["RU-R4093", "RU", "I", "R40-93", "А", {"en": "I-beams, А series (SAP2000 group R40-93)", "ar": "كمرات I سلسلة А (مجموعة SAP2000 ‏R40-93)", "ru": "Двутавры серии А (группа SAP2000 R40-93)", "zh": "А 系列工字钢（SAP2000 分组 R40-93）", "es": "Vigas I serie А (grupo SAP2000 R40-93)"}, null, ".", "Russian2020.pro"],
  ["RU-26020", "RU", "I", "ГОСТ 26020-83", "26020", {"en": "I-beams GOST 26020-83 (Б, Ш, К, ДБ, ДШ)", "ar": "كمرات I وفق GOST 26020-83 (Б، Ш، К، ДБ، ДШ)", "ru": "Двутавры ГОСТ 26020-83 (Б, Ш, К, ДБ, ДШ)", "zh": "工字钢 GOST 26020-83（Б、Ш、К、ДБ、ДШ）", "es": "Vigas I GOST 26020-83 (Б, Ш, К, ДБ, ДШ)"}, null, ".", "Russian2020.pro"],
  ["RU-19425-I", "RU", "I", "ГОСТ 19425-74", "I 19425", {"en": "Special I-beams (GOST 19425-74)", "ar": "كمرات I خاصة (GOST 19425-74)", "ru": "Двутавры специальные (ГОСТ 19425-74)", "zh": "特殊工字钢（GOST 19425-74）", "es": "Vigas I especiales (GOST 19425-74)"}, null, ".", "Russian2020.pro"],
  ["RU-P", "RU", "C", "ГОСТ 8240-97", "П", {"en": "Parallel-flange channels П (GOST 8240-97)", "ar": "قنوات متوازية الشفة П (GOST 8240-97)", "ru": "Швеллеры с параллельными гранями полок П (ГОСТ 8240-97)", "zh": "平行翼缘槽钢 П（GOST 8240-97）", "es": "Canales de alas paralelas П (GOST 8240-97)"}, null, ".", "Russian2020.pro"],
  ["RU-U", "RU", "C", "ГОСТ 8240-97", "У", {"en": "Tapered-flange channels У (GOST 8240-97)", "ar": "قنوات مائلة الشفة У (GOST 8240-97)", "ru": "Швеллеры с уклоном граней полок У (ГОСТ 8240-97)", "zh": "斜翼缘槽钢 У（GOST 8240-97）", "es": "Canales de alas inclinadas У (GOST 8240-97)"}, null, ".", "Russian2020.pro"],
  ["RU-E", "RU", "C", "ГОСТ 8240-97", "Э", {"en": "Economy channels Э (GOST 8240-97)", "ar": "قنوات اقتصادية Э (GOST 8240-97)", "ru": "Швеллеры экономичные Э (ГОСТ 8240-97)", "zh": "经济型槽钢 Э（GOST 8240-97）", "es": "Canales económicos Э (GOST 8240-97)"}, null, ".", "Russian2020.pro"],
  ["RU-CS", "RU", "C", "ГОСТ 8240-97", "С", {"en": "Special channels С (GOST 8240-97)", "ar": "قنوات خاصة С (GOST 8240-97)", "ru": "Швеллеры специальные С (ГОСТ 8240-97)", "zh": "特殊槽钢 С（GOST 8240-97）", "es": "Canales especiales С (GOST 8240-97)"}, null, ".", "Russian2020.pro"],
  ["RU-CL", "RU", "C", "ГОСТ 8240-97", "Л", {"en": "Light channels Л (GOST 8240-97)", "ar": "قنوات خفيفة Л (GOST 8240-97)", "ru": "Швеллеры лёгкой серии Л (ГОСТ 8240-97)", "zh": "轻型槽钢 Л（GOST 8240-97）", "es": "Canales ligeros Л (GOST 8240-97)"}, null, ".", "Russian2020.pro"],
  ["RU-19425-C", "RU", "C", "ГОСТ 19425-74", "[ 19425", {"en": "Special channels (GOST 19425-74)", "ar": "قنوات خاصة (GOST 19425-74)", "ru": "Швеллеры специальные (ГОСТ 19425-74)", "zh": "特殊槽钢（GOST 19425-74）", "es": "Canales especiales (GOST 19425-74)"}, null, ".", "Russian2020.pro"],
  ["RU-5267", "RU", "C", "ГОСТ 5267.1-90", "В", {"en": "Channels for wagon building (GOST 5267.1-90)", "ar": "قنوات لصناعة عربات القطارات (GOST 5267.1-90)", "ru": "Швеллеры для вагоностроения (ГОСТ 5267.1-90)", "zh": "车辆制造用槽钢（GOST 5267.1-90）", "es": "Canales para vagones (GOST 5267.1-90)"}, null, ".", "Russian2020.pro"],
  ["RU-T", "RU", "T", "", "Т", {"en": "Tees КТ, ШТ, БТ (SAP2000 Russian2020)", "ar": "مقاطع T ‏(КТ، ШТ، БТ) (SAP2000 Russian2020)", "ru": "Тавры КТ, ШТ, БТ (SAP2000 Russian2020)", "zh": "T 型钢 КТ、ШТ、БТ（SAP2000 Russian2020）", "es": "Perfiles T КТ, ШТ, БТ (SAP2000 Russian2020)"}, null, ".", "Russian2020.pro"],
  ["RU-L", "RU", "L", "ГОСТ 8509-93", "L", {"en": "Equal angles (GOST 8509-93)", "ar": "زوايا متساوية الأضلاع (GOST 8509-93)", "ru": "Уголки равнополочные (ГОСТ 8509-93)", "zh": "等边角钢（GOST 8509-93）", "es": "Angulares de lados iguales (GOST 8509-93)"}, "L{a}×{t} ГОСТ 8509-93", ",", "Russian2020.pro"],
  ["RU-LN", "RU", "L", "ГОСТ 8510-86", "L", {"en": "Unequal angles (GOST 8510-86)", "ar": "زوايا غير متساوية الأضلاع (GOST 8510-86)", "ru": "Уголки неравнополочные (ГОСТ 8510-86)", "zh": "不等边角钢（GOST 8510-86）", "es": "Angulares de lados desiguales (GOST 8510-86)"}, "L{a}×{b}×{t} ГОСТ 8510-86", ",", "Russian2020.pro"],
  ["RU-TK", "RU", "B", "ГОСТ 32931-2015", "□", {"en": "Square hollow sections (GOST 32931-2015)", "ar": "مقاطع مجوفة مربعة (GOST 32931-2015)", "ru": "Трубы профильные квадратные (ГОСТ 32931-2015)", "zh": "方形钢管（GOST 32931-2015）", "es": "Tubos cuadrados (GOST 32931-2015)"}, "□{h}×{b}×{t} ГОСТ 32931-2015", ",", "Russian2020.pro"],
  ["RU-TP", "RU", "B", "ГОСТ 32931-2015", "□", {"en": "Rectangular hollow sections (GOST 32931-2015)", "ar": "مقاطع مجوفة مستطيلة (GOST 32931-2015)", "ru": "Трубы профильные прямоугольные (ГОСТ 32931-2015)", "zh": "矩形钢管（GOST 32931-2015）", "es": "Tubos rectangulares (GOST 32931-2015)"}, "□{h}×{b}×{t} ГОСТ 32931-2015", ",", "Russian2020.pro"],
  ["RU-TK80", "RU", "B", "ТУ 36-2287-80", "□", {"en": "Square hollow sections (TU 36-2287-80)", "ar": "مقاطع مجوفة مربعة (TU 36-2287-80)", "ru": "Профили замкнутые квадратные (ТУ 36-2287-80)", "zh": "方形钢管（TU 36-2287-80）", "es": "Tubos cuadrados (TU 36-2287-80)"}, "□{h}×{b}×{t} ТУ 36-2287-80", ",", "Russian2020.pro"],
  ["RU-TP80", "RU", "B", "ТУ 36-2287-80", "□", {"en": "Rectangular hollow sections (TU 36-2287-80)", "ar": "مقاطع مجوفة مستطيلة (TU 36-2287-80)", "ru": "Профили замкнутые прямоугольные (ТУ 36-2287-80)", "zh": "矩形钢管（TU 36-2287-80）", "es": "Tubos rectangulares (TU 36-2287-80)"}, "□{h}×{b}×{t} ТУ 36-2287-80", ",", "Russian2020.pro"],
  ["RU-TK10", "RU", "B", "ГОСТ Р 54157-2010", "□", {"en": "Square hollow sections (GOST R 54157-2010)", "ar": "مقاطع مجوفة مربعة (GOST R 54157-2010)", "ru": "Трубы профильные квадратные (ГОСТ Р 54157-2010)", "zh": "方形钢管（GOST R 54157-2010）", "es": "Tubos cuadrados (GOST R 54157-2010)"}, "□{h}×{b}×{t} ГОСТ Р 54157-2010", ",", "Russian2020.pro"],
  ["RU-TP10", "RU", "B", "ГОСТ Р 54157-2010", "□", {"en": "Rectangular hollow sections (GOST R 54157-2010)", "ar": "مقاطع مجوفة مستطيلة (GOST R 54157-2010)", "ru": "Трубы профильные прямоугольные (ГОСТ Р 54157-2010)", "zh": "矩形钢管（GOST R 54157-2010）", "es": "Tubos rectangulares (GOST R 54157-2010)"}, "□{h}×{b}×{t} ГОСТ Р 54157-2010", ",", "Russian2020.pro"],
  ["RU-TK03", "RU", "B", "ГОСТ 30245-2003", "□", {"en": "Square hollow sections (GOST 30245-2003)", "ar": "مقاطع مجوفة مربعة (GOST 30245-2003)", "ru": "Профили гнутые замкнутые квадратные (ГОСТ 30245-2003)", "zh": "冷弯方形钢管（GOST 30245-2003）", "es": "Tubos cuadrados conformados (GOST 30245-2003)"}, "□{h}×{b}×{t} ГОСТ 30245-2003", ",", "Russian2020.pro"],
  ["RU-TP03", "RU", "B", "ГОСТ 30245-2003", "□", {"en": "Rectangular hollow sections (GOST 30245-2003)", "ar": "مقاطع مجوفة مستطيلة (GOST 30245-2003)", "ru": "Профили гнутые замкнутые прямоугольные (ГОСТ 30245-2003)", "zh": "冷弯矩形钢管（GOST 30245-2003）", "es": "Tubos rectangulares conformados (GOST 30245-2003)"}, "□{h}×{b}×{t} ГОСТ 30245-2003", ",", "Russian2020.pro"],
  ["RU-O", "RU", "P", "ГОСТ 32931-2015", "Ø", {"en": "Round tubes (GOST 32931-2015)", "ar": "أنابيب دائرية (GOST 32931-2015)", "ru": "Трубы круглые (ГОСТ 32931-2015)", "zh": "圆形钢管（GOST 32931-2015）", "es": "Tubos redondos (GOST 32931-2015)"}, "Ø{D}×{t} ГОСТ 32931-2015", ",", "Russian2020.pro"],
  ["RU-O10", "RU", "P", "ГОСТ Р 54157-2010", "Ø", {"en": "Round tubes (GOST R 54157-2010)", "ar": "أنابيب دائرية (GOST R 54157-2010)", "ru": "Трубы круглые (ГОСТ Р 54157-2010)", "zh": "圆形钢管（GOST R 54157-2010）", "es": "Tubos redondos (GOST R 54157-2010)"}, "Ø{D}×{t} ГОСТ Р 54157-2010", ",", "Russian2020.pro"],
  ["CN-HW", "CN", "I", "GB/T 11263", "HW", {"en": "Wide-flange H-sections HW (GB/T 11263)", "ar": "مقاطع H عريضة الشفة HW (GB/T 11263)", "ru": "Двутавры HW широкополочные (GB/T 11263)", "zh": "宽翼缘 H 型钢 HW（GB/T 11263）", "es": "Perfiles H de ala ancha HW (GB/T 11263)"}, null, ".", "ChineseGB08.pro"],
  ["CN-HM", "CN", "I", "GB/T 11263", "HM", {"en": "Medium-flange H-sections HM (GB/T 11263)", "ar": "مقاطع H متوسطة الشفة HM (GB/T 11263)", "ru": "Двутавры HM среднеполочные (GB/T 11263)", "zh": "中翼缘 H 型钢 HM（GB/T 11263）", "es": "Perfiles H de ala media HM (GB/T 11263)"}, null, ".", "ChineseGB08.pro"],
  ["CN-HN", "CN", "I", "GB/T 11263", "HN", {"en": "Narrow-flange H-sections HN (GB/T 11263)", "ar": "مقاطع H ضيقة الشفة HN (GB/T 11263)", "ru": "Двутавры HN узкополочные (GB/T 11263)", "zh": "窄翼缘 H 型钢 HN（GB/T 11263）", "es": "Perfiles H de ala estrecha HN (GB/T 11263)"}, null, ".", "ChineseGB08.pro"],
  ["CN-HT", "CN", "I", "GB/T 11263", "HT", {"en": "Thin-wall H-sections HT (GB/T 11263)", "ar": "مقاطع H رقيقة الجدار HT (GB/T 11263)", "ru": "Двутавры HT тонкостенные (GB/T 11263)", "zh": "薄壁 H 型钢 HT（GB/T 11263）", "es": "Perfiles H de pared delgada HT (GB/T 11263)"}, null, ".", "ChineseGB08.pro"],
  ["CN-YB-H", "CN", "I", "YB (SAP2000)", "YB-H", {"en": "Welded H-sections YB-H (SAP2000 YB group)", "ar": "مقاطع H ملحومة YB-H (مجموعة SAP2000 ‏YB)", "ru": "Сварные двутавры YB-H (группа SAP2000 YB)", "zh": "焊接 H 型钢 YB-H（SAP2000 YB 分组）", "es": "Perfiles H soldados YB-H (grupo SAP2000 YB)"}, null, ".", "ChineseGB08.pro"],
  ["CN-YB-LWH", "CN", "I", "YB (SAP2000)", "YB-LWH", {"en": "Light welded H-sections YB-LWH (SAP2000 YB group)", "ar": "مقاطع H ملحومة خفيفة YB-LWH (مجموعة SAP2000 ‏YB)", "ru": "Лёгкие сварные двутавры YB-LWH (группа SAP2000 YB)", "zh": "轻型焊接 H 型钢 YB-LWH（SAP2000 YB 分组）", "es": "Perfiles H soldados ligeros YB-LWH (grupo SAP2000 YB)"}, null, ".", "ChineseGB08.pro"],
  ["CN-YB-WH", "CN", "I", "YB (SAP2000)", "YB-WH", {"en": "Welded H-sections YB-WH (SAP2000 YB group)", "ar": "مقاطع H ملحومة YB-WH (مجموعة SAP2000 ‏YB)", "ru": "Сварные двутавры YB-WH (группа SAP2000 YB)", "zh": "焊接 H 型钢 YB-WH（SAP2000 YB 分组）", "es": "Perfiles H soldados YB-WH (grupo SAP2000 YB)"}, null, ".", "ChineseGB08.pro"],
  ["CN-LH", "CN", "I", "JG/T 137", "LH", {"en": "High-frequency welded thin-wall H LH (JG/T 137)", "ar": "مقاطع H رقيقة ملحومة بالتردد العالي LH (JG/T 137)", "ru": "Тонкостенные двутавры ТВЧ-сварки LH (JG/T 137)", "zh": "高频焊接薄壁 H 型钢 LH（JG/T 137）", "es": "Perfiles H de pared delgada soldados por alta frecuencia LH (JG/T 137)"}, null, ".", "ChineseGB08.pro"],
  ["CN-I", "CN", "I", "GB/T 706-2008", "I", {"en": "Hot-rolled I-beams (GB/T 706-2008)", "ar": "كمرات I مدرفلة على الساخن (GB/T 706-2008)", "ru": "Двутавры горячекатаные (GB/T 706-2008)", "zh": "热轧工字钢（GB/T 706-2008）", "es": "Vigas I laminadas en caliente (GB/T 706-2008)"}, null, ".", "ChineseGB08.pro"],
  ["CN-YB-I", "CN", "I", "YB (SAP2000)", "YB-I", {"en": "I-beams YB-I (SAP2000 YB group)", "ar": "كمرات I ‏YB-I (مجموعة SAP2000 ‏YB)", "ru": "Двутавры YB-I (группа SAP2000 YB)", "zh": "工字钢 YB-I（SAP2000 YB 分组）", "es": "Vigas I YB-I (grupo SAP2000 YB)"}, null, ".", "ChineseGB08.pro"],
  ["CN-C", "CN", "C", "GB/T 706-2008", "[", {"en": "Hot-rolled channels (GB/T 706-2008)", "ar": "قنوات مدرفلة على الساخن (GB/T 706-2008)", "ru": "Швеллеры горячекатаные (GB/T 706-2008)", "zh": "热轧槽钢（GB/T 706-2008）", "es": "Canales laminados en caliente (GB/T 706-2008)"}, null, ".", "ChineseGB08.pro"],
  ["CN-YB-C", "CN", "C", "YB (SAP2000)", "YB-C", {"en": "Channels YB-C (SAP2000 YB group)", "ar": "قنوات YB-C (مجموعة SAP2000 ‏YB)", "ru": "Швеллеры YB-C (группа SAP2000 YB)", "zh": "槽钢 YB-C（SAP2000 YB 分组）", "es": "Canales YB-C (grupo SAP2000 YB)"}, null, ".", "ChineseGB08.pro"],
  ["CN-TW", "CN", "T", "GB/T 11263", "TW", {"en": "Split tees TW (GB/T 11263)", "ar": "مقاطع T مشقوقة TW (GB/T 11263)", "ru": "Тавры разрезные TW (GB/T 11263)", "zh": "剖分 T 型钢 TW（GB/T 11263）", "es": "Perfiles T cortados TW (GB/T 11263)"}, null, ".", "ChineseGB08.pro"],
  ["CN-TM", "CN", "T", "GB/T 11263", "TM", {"en": "Split tees TM (GB/T 11263)", "ar": "مقاطع T مشقوقة TM (GB/T 11263)", "ru": "Тавры разрезные TM (GB/T 11263)", "zh": "剖分 T 型钢 TM（GB/T 11263）", "es": "Perfiles T cortados TM (GB/T 11263)"}, null, ".", "ChineseGB08.pro"],
  ["CN-TN", "CN", "T", "GB/T 11263", "TN", {"en": "Split tees TN (GB/T 11263)", "ar": "مقاطع T مشقوقة TN (GB/T 11263)", "ru": "Тавры разрезные TN (GB/T 11263)", "zh": "剖分 T 型钢 TN（GB/T 11263）", "es": "Perfiles T cortados TN (GB/T 11263)"}, null, ".", "ChineseGB08.pro"],
  ["CN-LN", "CN", "L", "GB/T 706-2008", "∠", {"en": "Unequal angles (GB/T 706-2008)", "ar": "زوايا غير متساوية الأضلاع (GB/T 706-2008)", "ru": "Уголки неравнополочные (GB/T 706-2008)", "zh": "热轧不等边角钢（GB/T 706-2008）", "es": "Angulares de lados desiguales (GB/T 706-2008)"}, "∠{a}×{b}×{t}", ".", "ChineseGB08.pro"],
  ["CN-L", "CN", "L", "GB/T 706-2008", "∠", {"en": "Equal angles (GB/T 706-2008)", "ar": "زوايا متساوية الأضلاع (GB/T 706-2008)", "ru": "Уголки равнополочные (GB/T 706-2008)", "zh": "热轧等边角钢（GB/T 706-2008）", "es": "Angulares de lados iguales (GB/T 706-2008)"}, "∠{a}×{t}", ".", "ChineseGB08.pro"],
  ["CN-GB-SSP", "CN", "P", "GB-SSP", "Φ", {"en": "Seamless steel pipes (SAP2000 GB-SSP)", "ar": "أنابيب فولاذية غير ملحومة (SAP2000 ‏GB-SSP)", "ru": "Трубы бесшовные (SAP2000 GB-SSP)", "zh": "无缝钢管（SAP2000 GB-SSP）", "es": "Tubos sin costura (SAP2000 GB-SSP)"}, "Φ{D}×{t} GB-SSP", ".", "ChineseGB08.pro"],
  ["CN-GB-SPWSP", "CN", "P", "GB-SPWSP", "Φ", {"en": "Spiral-welded steel pipes (SAP2000 GB-SPWSP)", "ar": "أنابيب ملحومة حلزونيًا (SAP2000 ‏GB-SPWSP)", "ru": "Трубы спиральношовные (SAP2000 GB-SPWSP)", "zh": "螺旋焊管（SAP2000 GB-SPWSP）", "es": "Tubos con soldadura helicoidal (SAP2000 GB-SPWSP)"}, "Φ{D}×{t} GB-SPWSP", ".", "ChineseGB08.pro"],
  ["CN-YB-STWSP", "CN", "P", "YB-STWSP", "Φ", {"en": "Straight-seam welded pipes (SAP2000 YB-STWSP)", "ar": "أنابيب ملحومة طوليًا (SAP2000 ‏YB-STWSP)", "ru": "Трубы прямошовные (SAP2000 YB-STWSP)", "zh": "直缝焊管（SAP2000 YB-STWSP）", "es": "Tubos con soldadura longitudinal (SAP2000 YB-STWSP)"}, "Φ{D}×{t} YB-STWSP", ".", "ChineseGB08.pro"],
  ["JP-H", "JP", "I", "JIS G 3192:2014", "JIS H", {"en": "H-sections (JIS G 3192:2014)", "ar": "مقاطع H ‏(JIS G 3192:2014)", "ru": "Двутавры H (JIS G 3192:2014)", "zh": "H 型钢（JIS G 3192:2014）", "es": "Perfiles H (JIS G 3192:2014)"}, null, ".", "JIS-G-3192-2014.pro"],
  ["IN-ISJB", "IN", "I", "IS 808", "ISJB", {"en": "Junior beams (IS 808)", "ar": "كمرات خفيفة جدًا ISJB (IS 808)", "ru": "Балки ISJB (IS 808)", "zh": "ISJB 轻型工字钢（IS 808）", "es": "Vigas ISJB (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISLB", "IN", "I", "IS 808", "ISLB", {"en": "Light beams (IS 808)", "ar": "كمرات خفيفة ISLB (IS 808)", "ru": "Балки лёгкие ISLB (IS 808)", "zh": "ISLB 轻型工字钢（IS 808）", "es": "Vigas ligeras ISLB (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISMB", "IN", "I", "IS 808", "ISMB", {"en": "Medium beams (IS 808)", "ar": "كمرات متوسطة ISMB (IS 808)", "ru": "Балки средние ISMB (IS 808)", "zh": "ISMB 中型工字钢（IS 808）", "es": "Vigas medias ISMB (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISWB", "IN", "I", "IS 808", "ISWB", {"en": "Wide-flange beams (IS 808)", "ar": "كمرات عريضة الشفة ISWB (IS 808)", "ru": "Балки широкополочные ISWB (IS 808)", "zh": "ISWB 宽翼缘工字钢（IS 808）", "es": "Vigas de ala ancha ISWB (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISHB", "IN", "I", "IS 808", "ISHB", {"en": "H-columns (IS 808)", "ar": "أعمدة H ‏ISHB (IS 808)", "ru": "Колонные двутавры ISHB (IS 808)", "zh": "ISHB H 型柱（IS 808）", "es": "Columnas H ISHB (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISJC", "IN", "C", "IS 808", "ISJC", {"en": "Junior channels (IS 808)", "ar": "قنوات خفيفة جدًا ISJC (IS 808)", "ru": "Швеллеры ISJC (IS 808)", "zh": "ISJC 轻型槽钢（IS 808）", "es": "Canales ISJC (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISLC", "IN", "C", "IS 808", "ISLC", {"en": "Light channels (IS 808)", "ar": "قنوات خفيفة ISLC (IS 808)", "ru": "Швеллеры лёгкие ISLC (IS 808)", "zh": "ISLC 轻型槽钢（IS 808）", "es": "Canales ligeros ISLC (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISMC", "IN", "C", "IS 808", "ISMC", {"en": "Medium channels (IS 808)", "ar": "قنوات متوسطة ISMC (IS 808)", "ru": "Швеллеры средние ISMC (IS 808)", "zh": "ISMC 中型槽钢（IS 808）", "es": "Canales medios ISMC (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISNT", "IN", "T", "IS 808", "ISNT", {"en": "Normal tees (IS 808)", "ar": "مقاطع T عادية ISNT (IS 808)", "ru": "Тавры ISNT (IS 808)", "zh": "ISNT T 型钢（IS 808）", "es": "Perfiles T ISNT (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISHT", "IN", "T", "IS 808", "ISHT", {"en": "H-tees (IS 808)", "ar": "مقاطع T ‏ISHT (IS 808)", "ru": "Тавры ISHT (IS 808)", "zh": "ISHT T 型钢（IS 808）", "es": "Perfiles T ISHT (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISST", "IN", "T", "IS 808", "ISST", {"en": "Slit tees (IS 808)", "ar": "مقاطع T مشقوقة ISST (IS 808)", "ru": "Тавры ISST (IS 808)", "zh": "ISST T 型钢（IS 808）", "es": "Perfiles T ISST (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISLT", "IN", "T", "IS 808", "ISLT", {"en": "Light tees (IS 808)", "ar": "مقاطع T خفيفة ISLT (IS 808)", "ru": "Тавры лёгкие ISLT (IS 808)", "zh": "ISLT T 型钢（IS 808）", "es": "Perfiles T ligeros ISLT (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISJT", "IN", "T", "IS 808", "ISJT", {"en": "Junior tees (IS 808)", "ar": "مقاطع T خفيفة جدًا ISJT (IS 808)", "ru": "Тавры ISJT (IS 808)", "zh": "ISJT T 型钢（IS 808）", "es": "Perfiles T ISJT (IS 808)"}, null, ".", "Indian.pro"],
  ["IN-ISA", "IN", "L", "IS 808", "ISA", {"en": "Angles (IS 808)", "ar": "زوايا ISA ‏(IS 808)", "ru": "Уголки ISA (IS 808)", "zh": "ISA 角钢（IS 808）", "es": "Angulares ISA (IS 808)"}, "ISA {a}×{b}×{t}", ".", "Indian.pro"],
  ["IN-ISB", "IN", "B", "IS 4923", "ISB", {"en": "Hollow sections (IS 4923)", "ar": "مقاطع مجوفة ISB ‏(IS 4923)", "ru": "Трубы профильные ISB (IS 4923)", "zh": "ISB 空心型钢（IS 4923）", "es": "Perfiles huecos ISB (IS 4923)"}, "ISB {h}×{b}×{t}", ".", "Indian.pro"],
  ["IN-ISNB", "IN", "P", "IS 1161", "ISNB", {"en": "Pipes, nominal bore (IS 1161)", "ar": "أنابيب بقطر اسمي ISNB ‏(IS 1161)", "ru": "Трубы ISNB (IS 1161)", "zh": "ISNB 钢管（IS 1161）", "es": "Tubos ISNB (IS 1161)"}, null, ".", "Indian.pro"]
];
const INTL_PACK = {
  "RU-B": "10Б1|100|55|4.1|5.7|1032;12Б1|117.6|64|3.8|5.1|1103;12Б2|120|64|4.4|6.3|1321;14Б1|137.4|73|3.8|5.6|1339;14Б2|140|73|4.7|6.9|1643;16Б1|157|82|4|5.9|1618;16Б2|160|82|5|7.4|2009;18Б1|177|91|4.3|6.5|1958;18Б2|180|91|5.3|8|2395;20Б0|198|99|4.5|7|2318;20Б1|200|100|5.5|8|2716;20Б2|203|101|6.5|9.5|3219;20Б3|208|102|8|12|4024;25Б1|248|124|5|8|3268;25Б2|250|125|6|9|3766;25Б3|255|126|7.5|11.5|4762;25Б4|260|127|9|14|5768;30Б1|298|149|5.5|8|4080;30Б2|300|150|6.5|9|4678;30Б3|305|151|8|11.5|5874;30Б4|310|152|9.5|14|7080;35Б1|346|174|6|9|5268;35Б2|350|175|7|11|6314;35Б3|355|176|8.5|13.5|7708;35Б4|361|177|10|16.5|9289;40Б1|396|199|7|11|7216;40Б2|400|200|8|13|8412;40Б3|406|201|9.5|16|10205;40Б4|412|202|11|19|12010;45Б1|446|199|8|12|8430;45Б2|450|200|9|14|9676;45Б3|456|201|10.5|17|11543;45Б4|462|202|12|20|13422;50Б1|492|199|8.8|12|9238;50Б2|496|199|9|14|10127;50Б3|500|200|10|16|11423;50Б4|508|201|12|20|13999;50Б5|516|202|15|24|17059;55Б1|543|220|9.5|13.5|11336;55Б2|547|220|10|15.5|12474;55Б3|553|221|12|18.5|14863;55Б4|560|222|14|22|17486;60Б1|596|199|10|15|12045;60Б2|600|200|11|17|13441;60Б3|604|201|12.5|19|15128;60Б4|612|202|15|23|18197;70Б1|691|260|12|15.5|16474;70Б2|697|260|12.5|18.5|18364;70Б3|702|261|14.5|21|21026;70Б4|710|262|17|25|24814",
  "RU-SH": "20Ш0|190|149|5|7|3111;20Ш1|194|150|6|9|3901;20Ш2|199|151|7.5|11.5|4938;20Ш3|204|152|9|14|5985;20Ш4|211|155|11|17.5|7506;20Ш5|218|157|13|21|9027;20Ш6|228|159|16|26|11229;25Ш0|240|174|6|9|4684;25Ш1|244|175|7|11|5624;25Ш2|249|176|8.5|13.5|6859;25Ш3|256|177|10.5|17|8569;25Ш4|264|182|13|21|10750;25Ш5|274|184|16|26|13340;25Ш6|286|186|19|32|16342;30Ш0|290|199|7|10|6148;30Ш1|294|200|8|12|7238;30Ш2|300|201|9|15|8738;30Ш3|306|203|11|18|10556;30Ш4|314|206|13|22|12852;30Ш5|326|208|16|28|16246;30Ш6|342|210|20|36|20798;35Ш1|334|249|8|11|8317;35Ш2|340|250|9|14|10151;35Ш3|347|252|11|17.5|12595;35Ш4|354|254|13|21|15067;35Ш5|364|258|16|26|18751;35Ш6|376|260|19|32|22911;35Ш7|392|262|23|40|28479;40Ш1|383|299|9.5|12.5|11291;40Ш2|390|300|10|16|13595;40Ш3|397|302|12|19.5|16489;40Ш4|406|304|14.5|24|20198;40Ш5|418|309|17.5|30|25220;40Ш6|430|311|21|36|30325;40Ш7|446|313|25|44|36909;45Ш0|434|299|10|15|13504;45Ш1|440|300|11|18|15738;45Ш2|446|302|13|21|18430;45Ш3|452|304|15|24|21146;45Ш4|464|308|18|30|26246;45Ш5|476|310|21|36|31298;45Ш6|492|312|25|44|38050;50Ш1|482|300|11|15|14552;50Ш2|487|300|14.5|17.5|17634;50Ш3|493|300|15.5|20.5|19886;50Ш4|499|300|16.5|23.5|22138;50Ш5|508|302|19|28|26080;50Ш6|518|310|22|33|30984;50Ш7|532|312|26|40|37292;50Ш8|548|314|30|48|44284;60Ш1|582|300|12|17|17449;60Ш2|589|300|16|20.5|21741;60Ш3|597|300|18|24.5|25237;60Ш4|605|300|20|28.5|28733;60Ш5|616|302|23|34|33813;60Ш6|630|315|27|41|41299;60Ш7|644|317|31|48|48093;60Ш8|664|319|36|58|57405;70Ш1|692|300|13|20|21149;70Ш2|698|300|15|23|24253;70Ш3|707|300|18|27.5|28909;70Ш4|715|300|20.5|31.5|32939;70Ш5|725|300|23|36.5|37569;70Ш6|740|313|27|44|45821;70Ш7|758|315|32|53|54927;70Ш8|780|317|38|64|66025",
  "RU-K": "15К1|147|149|6|8.5|3417;15К2|150|150|7|10|4014;15К3|155|151|8.5|12.5|4984;15К4|160|152|10|15|5964;15К5|166|153|12|18|7172;20К1|196|199|6.5|10|5269;20К2|200|200|8|12|6353;20К3|204|201|9|14|7357;20К4|210|201|10.5|17|8827;20К5|214|202|12|19|9933;20К6|220|202|14|22|11497;20К7|226|203|16|25|13111;20К8|234|203|18|29|15087;25К1|246|249|8|12|7972;25К2|250|250|9|14|9218;25К3|253|251|10|15.5|10221;25К4|257|252|11|17.5|11482;25К5|262|253|12.5|20|13115;25К6|267|253|14|22.5|14713;25К7|274|258|16|26|17188;25К8|281|259|18|29.5|19497;25К9|288|260|20|33|21820;25К10|298|261|23|38|25162;30К1|298|299|9|14|11080;30К2|300|300|10|15|11978;30К3|300|305|15|15|13478;30К4|304|301|11|17|13482;30К5|308|301|12|19|14956;30К6|312|302|13|21|16472;30К7|316|302|14.5|23|18085;30К8|316|357|14.5|23|20615;30К9|322|358|16|26|23214;30К10|328|359|18|29|25960;30К11|334|360|20|32|28718;30К12|341|361|22|35.5|31849;30К13|350|362|24|40|35718;30К14|356|371|27|43|39474;30К15|364|372|30|47|43346;30К16|374|373|33|52|47980;30К17|384|374|36|57|52634;30К18|396|375|39|63|58058;30К19|408|385|43|69|65018;30К20|422|387|47|76|71792;30К21|440|389|52|85|80448;35К1|342|348|10|15|13903;35К1,5|346|349|11|17|15641;35К2|350|350|12|19|17387;35К3|355|351|13.5|21.5|19648;35К4|360|352|15|24|21919;35К5|365|353|16.5|26.5|24200;35К6|369|360|18|28.5|26479;35К7|376|361|20|32|29687;35К8|382|362|22|35|32547;35К9|389|363|24|38.5|35782;35К10|396|364|26.5|42|39187;35К11|404|374|29|46|43799;35К12|414|375|32|51|48577;35К13|424|376|35|56|53375;35К14|434|377|38|61|58193;35К15|446|378|42|67|64099;35К16|458|392|46|73|71927;35К17|472|393|50|80|78823;35К18|488|394|55|88|86847;35К19|506|395|60|97|95693;35К20|520|409|65|104|105695;35К21|540|411|71|114|116203;35К22|562|413|77|125|127617;35К23|580|426|84|134|140719;35К24|604|430|92|146|154607;40К1|394|398|11|18|18681;40К2|400|400|13|21|21869;40К3|406|403|16|24|25487;40К4|414|405|18|28|29539;40К4,5|420|403|20|31|32561;40К5|429|400|23|35.5|37049;40К6|438|370|25|40|38965;40К7|448|371|28|45|43829;40К8|458|372|31|50|48713;40К9|470|373|35|56|54721;40К10|484|374|39|63|61501;40К11|494|392|43|68|69121;40К12|510|393|48|76|77335;40К13|528|394|53|85|86369;40К14|548|395|59|95|96587;40К15|564|410|65|103|108145;40К16|588|412|72|115|120951;40К17|616|414|80|129|135867;40К18|638|430|87|140|151961;40К19|668|435|96|155|169633",
  "RU-S": "13С1|126.5|114|9|9|3152;20С1|200|204|12|12|7153;25С1|244|252|11|11|8206;25С2|250|255|14|14|10468;30С1|294|302|12|12|10766;30С2|300|305|15|15|13478;32С1|326.7|319.7|24.8|24.8|22928;32С2|337.9|325.7|30.3|30.4|28397;35С1|338|351|13|13|13525;35С2|344|354|16|16|16663;35С3|350|357|19|19|19837;40С1|388|402|15|15|17845;40С2|394|405|18|18|21439;40С3|400|408|21|21|25069",
  "RU-DB": "20ДБ1|207|133|5.8|8.4|3387;20ДБ2|210|134|6.4|10.2|3997;25ДБ1|251|146|6|8.6|3964;25ДБ2|256|146|6.3|10.9|4708;25ДБ3|260|147|7.2|12.7|5473;25ДБ4|258|146|6.1|9.1|4170;25ДБ5|262|147|6.6|11.2|4924;25ДБ6|266|148|7.6|13|5722;30ДБ1|309|102|6|8.9|3612;30ДБ2|313|102|6.6|10.8|4176;30ДБ3|310|165|5.8|9.7|4954;30ДБ4|313|166|6.6|11.2|5704;30ДБ5|317|167|7.6|13.2|6685;30ДБ6|303|165|6|10.2|5130;30ДБ7|307|166|6.7|11.8|5884;30ДБ8|310|167|7.9|13.7|6876;35ДБ1|349|127|5.8|8.5|4174;35ДБ2|353|128|6.5|10.7|4984;35ДБ3|352|171|6.9|9.8|5734;35ДБ4|355|171|7.2|11.6|6445;35ДБ5|358|172|7.9|13.1|7217;35ДБ6|363|173.2|9.1|15.7|8545;35ДБ7|353|254|9.5|16.4|11593;35ДБ8|357|255|10.5|18.3|12917;35ДБ9|360|256|11.4|19.9|14059;35ДБ10|363|257|13|21.7|15528;40ДБ1|399|140|6.4|8.8|4994;40ДБ2|403|140|7|11.2|5890;40ДБ3|403|177|7.5|10.9|6807;40ДБ4|407|178|7.7|12.8|7583;40ДБ5|410|179|8.8|14.4|8599;40ДБ6|413|180|9.7|16|9545;40ДБ7|417|181|10.9|18.2|10826;45ДБ1|450|152|7.6|10.8|6628;45ДБ2|455|153|8|13.3|7586;45ДБ3|459|154|9.1|15.4|8729;45ДБ4|462|154.4|9.6|17|9448;45ДБ5|466|155.3|10.5|18.9|10456;45ДБ6|453|189.9|8.5|12.7|8547;45ДБ7|457|190|9|14.5|9451;45ДБ8|460|191|9.9|16|10439;45ДБ9|463|192|10.5|17.7|11376;45ДБ10|466|193|11.4|19|12303;45ДБ11|469|194|12.6|20.6|13472;53ДБ3|533|209|10.2|15.6|11778;53ДБ4|537|210|10.9|17.4|12920;53ДБ5|539|211|11.6|18.8|13888;53ДБ6|544|212|13.1|21.2|15698;53ДБ7|549|214|14.7|23.6|17616;60ДБ1|599|178|10|12.8|10429;60ДБ2|603|179|10.9|15|11754;60ДБ3|603|228|10.5|14.9|12951;60ДБ4|608|228|11.2|17.3|14449;60ДБ5|612|229|11.9|19.6|15932;60ДБ6|617|230|13.1|22.2|17852",
  "RU-DK": "10ДК1|96|100|5|8|2124;10ДК2|100|100|6|10|2604;10ДК3|120|106|12|20|5324;12ДК1|114|120|5|8|2534;12ДК2|120|120|6.5|11|3401;12ДК3|140|126|12.5|21|6641;14ДК1|133|140|5.5|8.5|3142;14ДК2|140|140|7|12|4296;14ДК3|160|145|13|22|8012;15ДК1|152|152|5.8|6.6|2861;15ДК2|157|153|6.6|9.3|3809;15ДК3|162|154|8.1|11.6|4747;16ДК1|152|160|6|9|3877;16ДК2|160|160|8|13|5425;16ДК3|180|166|14|23|9705;18ДК1|171|180|6|9.5|4525;18ДК2|180|180|8.3|14|6495;18ДК3|200|186|14.5|24|11325;20ДК1|203|203|7.2|11|5859;20ДК2|206|204|7.9|12.6|6658;20ДК3|210|205|9.1|14.2|7564;20ДК4|216|206|10.2|17.4|9106;20ДК5|222|209|13|20.6|11051;20ДК6|229|210|14.5|23.7|12677;25ДК1|253|254|8.6|14.2|9284;25ДК2|256|255|9.4|15.6|10208;25ДК3|260|256|10.7|17.3|11408;25ДК4|264|257|11.9|19.6|12888",
  "RU-8239": "I10 ГОСТ 8239-89|100|55|4.5|7.2|1200;I12 ГОСТ 8239-89|120|64|4.8|7.3|1470;I14 ГОСТ 8239-89|140|73|4.9|7.5|1740;I16 ГОСТ 8239-89|160|81|5|7.8|2020;I18 ГОСТ 8239-89|180|90|5.1|8.1|2340;I20 ГОСТ 8239-89|200|100|5.2|8.4|2680;I22 ГОСТ 8239-89|220|110|5.4|8.7|3060;I24 ГОСТ 8239-89|240|115|5.6|9.5|3480;I27 ГОСТ 8239-89|270|125|6|9.8|4020;I30 ГОСТ 8239-89|300|135|6.5|10.2|4650;I33 ГОСТ 8239-89|330|140|7|11.2|5380;I36 ГОСТ 8239-89|360|145|7.5|12.3|6190;I40 ГОСТ 8239-89|400|155|8.3|13|7260;I45 ГОСТ 8239-89|450|160|9|14.2|8470;I50 ГОСТ 8239-89|500|170|10|15.2|10000;I55 ГОСТ 8239-89|550|180|11|16.5|11800;I60 ГОСТ 8239-89|600|190|12|17.8|13800",
  "RU-STO93": "10Б1 СТО АСЧМ 20-93|100|55|4.1|5.7|1032;12Б1 СТО АСЧМ 20-93|117.6|64|3.8|5.1|1103;12Б2 СТО АСЧМ 20-93|120|64|4.4|6.3|1321;14Б1 СТО АСЧМ 20-93|137.4|73|3.8|5.6|1339;14Б2 СТО АСЧМ 20-93|140|73|4.7|6.9|1643;16Б1 СТО АСЧМ 20-93|157|82|4|5.9|1618;16Б2 СТО АСЧМ 20-93|160|82|5|7.4|2009;18Б1 СТО АСЧМ 20-93|177|91|4.3|6.5|1958;18Б2 СТО АСЧМ 20-93|180|91|5.3|8|2395;20Б1 СТО АСЧМ 20-93|200|100|5.5|8|2716;25Б1 СТО АСЧМ 20-93|248|124|5|8|3268;25Б2 СТО АСЧМ 20-93|250|125|6|9|3766;30Б1 СТО АСЧМ 20-93|298|149|5.5|8|4060;30Б2 СТО АСЧМ 20-93|300|150|6.5|9|4678;35Б1 СТО АСЧМ 20-93|346|174|6|9|5268;35Б2 СТО АСЧМ 20-93|350|175|7|11|6314;40Б1 СТО АСЧМ 20-93|396|199|7|11|7216;40Б2 СТО АСЧМ 20-93|400|200|8|13|8412;45Б1 СТО АСЧМ 20-93|446|199|8|12|8430;45Б2 СТО АСЧМ 20-93|450|200|9|14|9676;50Б1 СТО АСЧМ 20-93|492|199|8.8|12|9238;50Б2 СТО АСЧМ 20-93|496|199|9|14|10127;50Б3 СТО АСЧМ 20-93|500|200|10|16|11423;55Б1 СТО АСЧМ 20-93|543|220|9.5|13.5|11336;55Б2 СТО АСЧМ 20-93|547|220|10|15.5|12475;60Б1 СТО АСЧМ 20-93|596|199|10|15|12045;60Б2 СТО АСЧМ 20-93|600|200|11|17|13445;70Б0 СТО АСЧМ 20-93|693|230|11.3|15.2|15305;70Б1 СТО АСЧМ 20-93|691|260|12|15.5|16474;70Б2 СТО АСЧМ 20-93|697|260|12.5|18.5|18364;20Ш1 СТО АСЧМ 20-93|194|150|6|9|3901;25Ш1 СТО АСЧМ 20-93|244|175|7|11|5624;30Ш1 СТО АСЧМ 20-93|294|200|8|12|7238;30Ш2 СТО АСЧМ 20-93|300|201|9|15|8738;35Ш1 СТО АСЧМ 20-93|334|249|8|11|8317;35Ш2 СТО АСЧМ 20-93|340|250|9|14|10151;40Ш1 СТО АСЧМ 20-93|383|299|9.5|12.5|11291;40Ш2 СТО АСЧМ 20-93|390|300|10|16|13595;45Ш1 СТО АСЧМ 20-93|440|300|11|18|15739;50Ш1 СТО АСЧМ 20-93|482|300|11|15|14552;50Ш2 СТО АСЧМ 20-93|487|300|14.5|17.5|17634;50Ш3 СТО АСЧМ 20-93|493|300|15.5|20.5|19886;50Ш4 СТО АСЧМ 20-93|499|300|16.5|23.5|22138;60Ш1 СТО АСЧМ 20-93|582|300|12|17|17449;60Ш2 СТО АСЧМ 20-93|589|300|16|20.5|21741;60Ш3 СТО АСЧМ 20-93|597|300|18|24.5|25237;60Ш4 СТО АСЧМ 20-93|605|300|20|28.5|28733;70Ш1 СТО АСЧМ 20-93|692|300|13|20|21149;70Ш2 СТО АСЧМ 20-93|698|300|15|23|24253;70Ш3 СТО АСЧМ 20-93|707|300|18|27.5|28909;70Ш4 СТО АСЧМ 20-93|715|300|20.5|31.5|32939;70Ш5 СТО АСЧМ 20-93|725|300|23|36.5|37569;80Ш1 СТО АСЧМ 20-93|782|300|13.5|17|20971;80Ш2 СТО АСЧМ 20-93|792|300|14|22|24345;90Ш1 СТО АСЧМ 20-93|881|299|15|18.5|24396;90Ш2 СТО АСЧМ 20-93|890|299|15|23|27087;100Ш1 СТО АСЧМ 20-93|990|320|16|21|29380;100Ш2 СТО АСЧМ 20-93|998|320|17|25|32888;100Ш3 СТО АСЧМ 20-93|1006|320|18|29|36396;100Ш4 СТО АСЧМ 20-93|1013|320|19.5|32.5|40058;20К1 СТО АСЧМ 20-93|196|199|6.5|10|5269;20К2 СТО АСЧМ 20-93|200|200|8|12|6353;25К1 СТО АСЧМ 20-93|246|249|8|12|7972;25К2 СТО АСЧМ 20-93|250|250|9|14|9218;25К3 СТО АСЧМ 20-93|253|251|10|15.5|10221;30К1 СТО АСЧМ 20-93|298|299|9|14|11080;30К2 СТО АСЧМ 20-93|300|300|10|15|11978;30К3 СТО АСЧМ 20-93|300|305|15|15|13479;30К4 СТО АСЧМ 20-93|304|301|11|17|13462;35К1 СТО АСЧМ 20-93|342|348|10|15|13903;35К2 СТО АСЧМ 20-93|350|350|12|19|17387;40К1 СТО АСЧМ 20-93|394|398|11|18|18681;40К2 СТО АСЧМ 20-93|400|400|13|21|21867;40К3 СТО АСЧМ 20-93|406|403|16|24|25487;40К4 СТО АСЧМ 20-93|414|405|18|28|29539;40К5 СТО АСЧМ 20-93|429|400|23|35.5|37049;24М СТО АСЧМ 20-93|240|110|8.2|14|4870;30М СТО АСЧМ 20-93|300|130|9|15|6400;36М СТО АСЧМ 20-93|360|130|9.5|16|7390;40ЕС СТО АСЧМ 20-93|392|165|7|8|5888;40ЕС1 СТО АСЧМ 20-93|392|165|7|9.5|5888;КХБ-515 СТО АСЧМ 20-93|515|180|15|25|16300;КХБ-526 СТО АСЧМ 20-93|526|300|22|35.5|31890",
  "RU-R4093": "20Д1А R40-93|207|133|5.8|8.4|3390;20Д2А R40-93|210|134|6.4|10.2|3970;25Д2А R40-93|258|146|6.1|9.1|4190;25Д3А R40-93|262|147|6.6|11.2|4910;36У1А R40-93|349|127|5.8|8.5|4190;36У2А R40-93|353|128|6.5|10.7|4960;15К2А R40-93|157|153|6.6|9.3|3790;15К3А R40-93|162|154|8.1|11.6|4740;20К2А R40-93|203|203|7.2|11|5890;20К3А R40-93|206|204|7.9|12.6|6660;20К4А R40-93|210|205|9.1|11.2|7990",
  "RU-26020": "20К1 ГОСТ 26020-83|195|200|6.5|10|5282;20К2 ГОСТ 26020-83|198|200|7|11.5|5970;23К1 ГОСТ 26020-83|227|240|7|10.5|6651;23К2 ГОСТ 26020-83|230|240|8|12|7577;26К1 ГОСТ 26020-83|255|260|8|12|8308;26К2 ГОСТ 26020-83|258|260|9|13.5|9319;26К3 ГОСТ 26020-83|262|260|10|15.5|10590;30К1 ГОСТ 26020-83|296|300|9|13.5|10800;30К2 ГОСТ 26020-83|300|300|10|15.5|12270;30К3 ГОСТ 26020-83|304|300|11.5|17.5|13872;35К1 ГОСТ 26020-83|343|350|10|15|13970;35К2 ГОСТ 26020-83|348|350|11|17.5|16040;35К3 ГОСТ 26020-83|353|350|13|20|18410;40К1 ГОСТ 26020-83|393|400|11|16.5|17580;40К2 ГОСТ 26020-83|400|400|13|20|21096;40К3 ГОСТ 26020-83|409|400|16|24.5|25780;40К4 ГОСТ 26020-83|419|400|19|29.5|30860;40К5 ГОСТ 26020-83|431|400|23|35.5|37100;24ДБ1 ГОСТ 26020-83|239|115|5.5|9.3|3545;27ДБ1 ГОСТ 26020-83|269|125|6|9.5|4068;36ДБ1 ГОСТ 26020-83|360|145|7.2|12.3|6260;35ДБ1 ГОСТ 26020-83|349|127|5.8|8.5|4278;40ДБ1 ГОСТ 26020-83|399|139|6.2|9|5058;45ДБ1 ГОСТ 26020-83|450|152|7.4|11|6705;45ДБ2 ГОСТ 26020-83|450|180|7.6|13.3|8280;30ДШ1 ГОСТ 26020-83|300.6|201.9|9.4|16|9260;40ДШ1 ГОСТ 26020-83|397.6|302|11.5|18.7|15900;50ДШ1 ГОСТ 26020-83|496.2|303.8|14.2|21|19800;10Б1 ГОСТ 26020-83|100|55|4.1|5.7|1032;12Б1 ГОСТ 26020-83|117.6|64|3.8|5.1|1103;12Б2 ГОСТ 26020-83|120|64|4.4|6.3|1321;14Б1 ГОСТ 26020-83|137.4|73|3.8|5.6|1339;14Б2 ГОСТ 26020-83|140|73|4.7|6.9|1643;16Б1 ГОСТ 26020-83|157|82|4|5.9|1618;16Б2 ГОСТ 26020-83|160|82|5|7.4|2009;18Б1 ГОСТ 26020-83|177|91|4.3|6.5|1958;18Б2 ГОСТ 26020-83|180|91|5.3|8|2395;20Б1 ГОСТ 26020-83|200|100|5.6|8.5|2849;23Б1 ГОСТ 26020-83|230|110|5.6|9|3291;26Б1 ГОСТ 26020-83|258|120|5.8|8.5|3562;26Б2 ГОСТ 26020-83|261|120|6|10|3970;30Б1 ГОСТ 26020-83|296|140|5.8|8.5|4192;30Б2 ГОСТ 26020-83|299|140|6|10|4667;35Б1 ГОСТ 26020-83|346|155|6.2|8.5|4953;35Б2 ГОСТ 26020-83|349|155|6.5|10|5517;40Б1 ГОСТ 26020-83|392|165|7|9.5|6125;40Б2 ГОСТ 26020-83|396|165|7.5|11.5|6972;45Б1 ГОСТ 26020-83|443|180|7.8|11|7623;45Б2 ГОСТ 26020-83|447|180|8.4|13|8596;50Б1 ГОСТ 26020-83|492|200|8.8|12|9298;50Б2 ГОСТ 26020-83|496|200|9.2|14|10280;55Б1 ГОСТ 26020-83|543|220|9.5|13.5|11337;55Б2 ГОСТ 26020-83|547|220|10|15.5|12475;60Б1 ГОСТ 26020-83|593|230|10.5|15.5|13526;60Б2 ГОСТ 26020-83|597|230|11|17.5|14730;70Б1 ГОСТ 26020-83|691|260|12|15.5|16470;70Б2 ГОСТ 26020-83|697|260|12.5|18.5|18360;80Б1 ГОСТ 26020-83|791|280|13.5|17|20320;80Б2 ГОСТ 26020-83|798|280|14|20.5|22660;90Б1 ГОСТ 26020-83|893|300|15|18.5|24710;90Б2 ГОСТ 26020-83|900|300|15.5|22|27240;100Б1 ГОСТ 26020-83|990|320|16|21|29382;100Б2 ГОСТ 26020-83|998|320|17|25|32890;100Б3 ГОСТ 26020-83|1006|320|18|29|36400;100Б4 ГОСТ 26020-83|1013|320|19.5|32.5|40060;20Ш1 ГОСТ 26020-83|193|150|6|9|3895;23Ш1 ГОСТ 26020-83|226|155|6.5|10|4608;26Ш1 ГОСТ 26020-83|251|180|7|10|5437;26Ш2 ГОСТ 26020-83|255|180|7.5|12|6273;30Ш1 ГОСТ 26020-83|291|200|8|11|6831;30Ш2 ГОСТ 26020-83|295|200|8.5|13|7765;30Ш3 ГОСТ 26020-83|299|200|9|15|8700;35Ш1 ГОСТ 26020-83|338|250|9.5|12.5|9567;35Ш2 ГОСТ 26020-83|341|250|10|14|10474;35Ш3 ГОСТ 26020-83|345|250|10.5|16|11630;40Ш1 ГОСТ 26020-83|388|300|9.5|14|12240;40Ш2 ГОСТ 26020-83|392|300|11.5|16|14160;40Ш3 ГОСТ 26020-83|396|300|12.5|18|15720;50Ш1 ГОСТ 26020-83|484|300|11|15|14570;50Ш2 ГОСТ 26020-83|489|300|14.5|17.5|17660;50Ш3 ГОСТ 26020-83|495|300|15.5|20.5|19920;50Ш4 ГОСТ 26020-83|501|300|16.5|23.5|22170;60Ш1 ГОСТ 26020-83|580|320|12|17|18110;60Ш2 ГОСТ 26020-83|587|320|16|20.5|22530;60Ш3 ГОСТ 26020-83|595|320|18|24.5|26180;60Ш4 ГОСТ 26020-83|603|320|20|28.5|29834;70Ш1 ГОСТ 26020-83|683|320|13.5|19|21640;70Ш2 ГОСТ 26020-83|691|320|15|23|25170;70Ш3 ГОСТ 26020-83|700|320|18|27.5|29980;70Ш4 ГОСТ 26020-83|708|320|20.5|31.5|34160;70Ш5 ГОСТ 26020-83|718|320|23|36.5|38970",
  "RU-19425-I": "I14С ГОСТ 19425-74|140|80|5.5|9.1|2150;I18М ГОСТ 19425-74|180|90|7|12|3290;I20С ГОСТ 19425-74|200|100|7|11.4|3560;I20Са ГОСТ 19425-74|200|102|9|11.4|3960;I22С ГОСТ 19425-74|220|110|7.5|12.3|4210;I24М ГОСТ 19425-74|240|110|8.2|14|4870;I27С ГОСТ 19425-74|270|122|8.5|13.7|5150;I27Са ГОСТ 19425-74|270|124|10.5|13.7|5990;I30М ГОСТ 19425-74|300|130|9|15|6400;I36М ГОСТ 19425-74|360|130|9.5|16|7380;I36С ГОСТ 19425-74|350|140|11|15.8|9090;I45М ГОСТ 19425-74|450|150|10.5|18|9880",
  "RU-P": "5П|50|32|4.4|7|616;6,5П|65|36|4.4|7.2|751;8П|80|40|4.5|7.4|898;10П|100|46|4.5|7.6|1090;12П|120|52|4.8|7.8|1330;14П|140|58|4.9|8.1|1560;16П|160|64|5|8.4|1810;16аП|160|68|5|9|1950;18П|180|70|5.1|8.7|2070;18аП|180|74|5.1|9.3|2220;20П|200|76|5.2|9|2340;22П|220|82|5.4|9.5|2670;24П|240|90|5.6|10|3060;27П|270|95|6|10.5|3520;30П|300|100|6.5|11|4050;33П|330|105|7|11.7|4650;36П|360|110|7.5|12.6|5340;40П|400|115|8|13.5|6150",
  "RU-U": "5У|50|32|4.4|7|616;6,5У|65|36|4.4|7.2|751;8У|80|40|4.5|7.4|898;10У|100|46|4.5|7.6|1090;12У|120|52|4.8|7.8|1330;14У|140|58|4.9|8.1|1560;16У|160|64|5|8.4|1810;16аУ|160|68|5|9|1950;18У|180|70|5.1|8.7|2070;18аУ|180|74|5.1|9.3|2220;20У|200|76|5.2|9|2340;22У|220|82|5.4|9.5|2670;24У|240|90|5.6|10|3060;27У|270|95|6|10.5|3520;30У|300|100|6.5|11|4050;33У|330|105|7|11.7|4650;36У|360|110|7.5|12.6|5340;40У|400|115|8|13.5|6150",
  "RU-E": "5Э|50|32|4.2|7|610;6,5Э|65|36|4.2|7.2|741;8Э|80|40|4.2|7.4|882;10Э|100|46|4.2|7.6|1079;12Э|120|52|4.5|7.8|1309;14Э|140|58|4.6|8.1|1541;16Э|160|64|4.7|8.4|1785;18Э|180|70|4.8|8.7|2040;20Э|200|76|4.9|9|2302;22Э|220|82|5.1|9.5|2636;24Э|240|90|5.3|10|3019;27Э|270|95|5.8|10.5|3487;30Э|300|100|6.3|11|3994;33Э|330|105|6.9|11.7|4615;36Э|360|110|7.4|12.6|5290;40Э|400|115|7.9|13.5|6111",
  "RU-CS": "8С|80|45|5.5|9|1180;14С|140|58|6|9.5|1851;14Са|140|60|8|9.5|2130;16С|160|63|6.5|10|2195;16Са|160|65|8.5|10|2515;18С|180|68|7|10.5|2570;18Са|180|70|9|10.5|2930;18Сб|180|100|8|10.5|3404;20С|200|73|7|11|2883;20Сб|200|100|8|11|3658;24С|240|85|9.5|14|4446;26С|260|65|10|16|4409;26Са|260|90|10|15|5060;30С|300|85|7.5|13.5|4388;30Са|300|87|9.5|13.5|4988;30Сб|300|89|11.5|13.5|5588",
  "RU-CL": "12Л|120|30|3|4.8|639;14Л|140|32|3.2|5.6|757;16Л|160|35|3.4|5.3|904;18Л|180|40|3.6|5.6|1081;20Л|200|45|3.8|6|1289;22Л|220|50|4|6.4|1511;24Л|240|55|4.2|6.8|1741;27Л|270|60|4.5|7.3|2077;30Л|300|65|4.8|7.8|2430",
  "RU-19425-C": "[18С ГОСТ 19425-74|180|68|7|10.5|2570;[18Са ГОСТ 19425-74|180|70|9|10.5|2930;[20С ГОСТ 19425-74|200|73|7|11|2880;[30С ГОСТ 19425-74|300|87|9.5|13.5|4960",
  "RU-5267": "8В ГОСТ 5267.1-90|80|45|5.5|9|1180;14В ГОСТ 5267.1-90|140|60|8|9.5|2130;18В ГОСТ 5267.1-90|180|100|8|10.5|3404;20В ГОСТ 5267.1-90|200|73|7|11|2883;20В-1 ГОСТ 5267.1-90|200|75|9|11|3283;20В-2 ГОСТ 5267.1-90|200|100|8|11|3658;26В ГОСТ 5267.1-90|260|90|10|15|5060;30В ГОСТ 5267.1-90|300|85|7.5|13.5|4388;30В-1 ГОСТ 5267.1-90|300|87|9.5|13.5|4988;30В-2 ГОСТ 5267.1-90|300|89|11.5|13.5|5388",
  "RU-T": "10КТ1|94|200|6.5|10|2619;10КТ2|95.5|200|7|11.5|2961;11,5КТ1|110|240|7|10.5|3301;11,5КТ2|111.5|240|8|12|3760;13КТ1|124|260|8|12|4126;13КТ2|125.5|260|9|13.5|4628;13КТ3|127.5|260|10|15.5|5260;15КТ1|144.5|300|9|13.5|5369;15КТ2|146.5|300|10|15.5|6099;15КТ3|148.5|300|11.5|17.5|6896;17,5КТ1|168|350|10|15|6952;17,5КТ2|170.5|350|11|17.5|7980;20КТ1|193|400|11|16.5|8749;20КТ2|196.5|400|13|20|10502;13ШТ1|122|180|7|10|2694;13ШТ2|124|180|7.5|12|3110;15ШТ1|142|200|8|11|3397;15ШТ2|144|200|8.5|13|3853;15ШТ3|146|200|9|15|4318;17,5ШТ1|165.5|250|9.5|12.5|4750;17,5ШТ2|167|250|10|14|5202;17,5ШТ3|169|250|10.5|16|5778;20ШТ1|190.5|300|9.5|14|6084;20ШТ2|192.5|300|11.5|16|7037;20ШТ3|194.5|300|12.5|18|7814;25ШТ1|238.5|300|11|15|7249;25ШТ2|241|300|14.5|17.5|8781;25ШТ3|244|300|15.5|20.5|9904;25ШТ4|247|300|16.5|23.5|11028;30ШТ1|286.5|320|12|17|9010;30ШТ2|290|320|16|20.5|11208;30ШТ3|294|320|18|24.5|13027;30ШТ4|298|320|20|28.5|14846",
  "RU-L": "20|20|3|113;20|20|4|146;25|25|3|143;25|25|4|186;25|25|5|227;28|28|3|162;30|30|3|174;30|30|4|227;30|30|5|278;32|32|3|186;32|32|4|243;35|35|3|204;35|35|4|267;35|35|5|328;40|40|3|235;40|40|4|308;40|40|5|379;40|40|6|448;45|45|3|265;45|45|4|348;45|45|5|429;45|45|6|508;50|50|3|296;50|50|4|389;50|50|5|480;50|50|6|569;50|50|7|656;50|50|8|741;56|56|4|438;56|56|5|541;60|60|4|472;60|60|5|583;60|60|6|692;60|60|8|904;60|60|10|1108;63|63|4|496;63|63|5|613;63|63|6|728;65|65|6|752;65|65|8|984;70|70|4.5|620;70|70|5|686;70|70|6|815;70|70|7|942;70|70|8|1067;70|70|10|1311;75|75|5|739;75|75|6|878;75|75|7|1015;75|75|8|1150;75|75|9|1283;80|80|5.5|863;80|80|6|938;80|80|7|1085;80|80|8|1230;80|80|10|1514;80|80|12|1790;90|90|6|1061;90|90|7|1228;90|90|8|1393;90|90|9|1560;90|90|10|1717;90|90|12|2033;100|100|6.5|1282;100|100|7|1375;100|100|8|1560;100|100|10|1924;100|100|12|2280;100|100|14|2628;100|100|15|2799;100|100|16|2968;110|110|7|1515;110|110|8|1720;120|120|8|1880;120|120|10|2324;120|120|12|2760;120|120|15|3399;125|125|8|1969;125|125|9|2200;125|125|10|2433;125|125|12|2889;125|125|14|3337;125|125|16|3777;140|140|9|2472;140|140|10|2733;140|140|12|3249;150|150|10|2933;150|150|12|3489;150|150|15|4308;150|150|18|5109;160|160|10|3143;160|160|11|3442;160|160|12|3739;160|160|14|4357;160|160|16|4907;160|160|18|5479;160|160|20|6040;180|180|11|3880;180|180|12|4219;180|180|15|5218;180|180|18|6199;180|180|20|6843;200|200|12|4710;200|200|13|5085;200|200|14|5460;200|200|16|6198;200|200|18|6930;200|200|20|7654;200|200|24|9078;200|200|25|9429;200|200|30|11154;220|220|14|6038;220|220|16|6858;250|250|16|7840;250|250|18|8772;250|250|20|9696;250|250|22|10612;250|250|25|11971;250|250|28|13312;250|250|30|14196;250|250|35|16371",
  "RU-LN": "25|16|3|116;30|20|3|143;30|20|4|186;32|20|3|149;32|20|4|194;40|25|3|189;40|25|4|247;40|25|5|303;40|30|4|267;40|30|5|328;45|28|3|214;45|28|4|280;50|32|3|242;50|32|4|317;56|36|4|358;56|36|5|441;63|40|4|404;63|40|5|498;63|40|6|590;63|40|8|768;65|50|5|556;65|50|6|660;65|50|7|762;65|50|8|862;70|45|5|559;75|50|5|611;75|50|6|725;75|50|7|837;75|50|8|947;80|50|5|636;80|50|6|755;80|60|6|815;80|60|7|942;80|60|8|1067;90|56|5.5|786;90|56|6|854;90|56|8|1118;100|63|6|958;100|63|7|1109;100|63|8|1257;100|63|10|1547;100|65|7|1123;100|65|8|1273;100|65|10|1567;110|70|6.5|1145;110|70|8|1393;125|80|7|1406;125|80|8|1598;125|80|10|1970;125|80|12|2336;140|90|8|1800;140|90|10|2224;160|100|9|2287;160|100|10|2528;160|100|12|3004;160|100|14|3472;180|110|10|2833;180|110|12|3369;200|125|11|3487;200|125|12|3789;200|125|14|4387;200|125|16|4977",
  "RU-TK": "40|40|2|294;40|40|2.5|359;40|40|3|421;40|40|3.5|479;40|40|4|535;50|50|2|374;50|50|2.5|459;50|50|3|541;50|50|3.5|619;50|50|4|695;50|50|4.5|767;50|50|5|836;50|50|5.5|901;50|50|6|963;60|60|2|454;60|60|2.5|559;60|60|3|661;60|60|3.5|759;60|60|4|855;60|60|4.5|947;60|60|5|1036;60|60|5.5|1121;60|60|6|1203;70|70|2|534;70|70|2.5|659;70|70|3|781;70|70|3.5|899;70|70|4|1015;70|70|4.5|1127;70|70|5|1236;70|70|5.5|1341;70|70|6|1443;70|70|6.5|1506;70|70|7|1596;80|80|3|901;80|80|3.5|1039;80|80|4|1175;80|80|4.5|1307;80|80|5|1436;80|80|5.5|1561;80|80|6|1683;80|80|6.5|1766;80|80|7|1876;80|80|7.5|1982;80|80|8|2084;90|90|3|1021;90|90|3.5|1179;90|90|4|1335;90|90|4.5|1487;90|90|5|1636;90|90|5.5|1781;90|90|6|1923;90|90|6.5|2026;90|90|7|2156;90|90|7.5|2282;90|90|8|2404;100|100|3|1141;100|100|3.5|1319;100|100|4|1495;100|100|4.5|1667;100|100|5|1836;100|100|5.5|2001;100|100|6|2163;100|100|6.5|2286;100|100|7|2436;100|100|7.5|2582;100|100|8|2724;120|120|3|1381;120|120|3.5|1599;120|120|4|1815;120|120|4.5|2027;120|120|5|2236;120|120|5.5|2441;120|120|6|2643;120|120|6.5|2806;120|120|7|2996;120|120|7.5|3182;120|120|8|3364;140|140|4|2135;140|140|4.5|2387;140|140|5|2636;140|140|5.5|2881;140|140|6|3123;140|140|6.5|3326;140|140|7|3556;140|140|7.5|3782;140|140|8|4004;150|150|4|2295;150|150|4.5|2567;150|150|5|2836;150|150|5.5|3101;150|150|6|3363;150|150|6.5|3586;150|150|7|3836;150|150|7.5|4082;150|150|8|4324;160|160|4|2455;160|160|4.5|2747;160|160|5|3036;160|160|5.5|3321;160|160|6|3603;160|160|6.5|3846;160|160|7|4116;160|160|7.5|4382;160|160|8|4644;180|180|5|3436;180|180|5.5|3761;180|180|6|4083;180|180|6.5|4366;180|180|7|4676;180|180|7.5|4982;180|180|8|5284;180|180|8.5|5583;180|180|9|5878;180|180|9.5|6169;180|180|10|6457;180|180|10.5|6645;180|180|11|6916;180|180|11.5|7182;180|180|12|7445;180|180|12.5|7703;180|180|13|7957;180|180|13.5|8207;180|180|14|8453;180|180|14.5|8695;180|180|15|8933;180|180|15.5|9166;180|180|16|9395;200|200|5|3836;200|200|5.5|4201;200|200|6|4563;200|200|6.5|4886;200|200|7|5236;200|200|7.5|5582;200|200|8|5924;200|200|8.5|6263;200|200|9|6598;200|200|9.5|6929;200|200|10|7257;200|200|10.5|7486;200|200|11|7797;200|200|11.5|8103;200|200|12|8406;250|250|6|5763;250|250|6.5|6186;250|250|7|6636;250|250|7.5|7082;250|250|8|7524;250|250|8.5|7963;250|250|9|8398;250|250|9.5|8829;250|250|10|9257;250|250|10.5|9586;250|250|11|9997;250|250|11.5|10400;250|250|12|10810;300|300|6|6963;300|300|6.5|7486;300|300|7|8036;300|300|7.5|8582;300|300|8|9124;300|300|8.5|9663;300|300|9|10200;300|300|9.5|10730;300|300|10|11260;300|300|10.5|11690;300|300|11|12200;300|300|11.5|12700;300|300|12|13210;300|300|12.5|13700;300|300|13|14200;300|300|13.5|14690;300|300|14|15170;300|300|14.5|15650;300|300|15|16130;300|300|15.5|16610;300|300|16|17080;300|300|16.5|17540;300|300|17|18000;300|300|17.5|18460;300|300|18|18910;300|300|18.5|19360;300|300|19|19800;300|300|20|20680;300|300|21|21540;300|300|22|22380;350|350|6|8163;350|350|6.5|8786;350|350|7|9435;350|350|7.5|10080;350|350|8|10720;350|350|8.5|11360;350|350|9|12000;350|350|9.5|12630;350|350|10|13260;350|350|10.5|13780;350|350|11|14400;350|350|11.5|15000;350|350|12|15600;350|350|12.5|16200;350|350|13|16800;350|350|13.5|17390;350|350|14.5|18550;350|350|15|19130;350|350|15.5|19710;350|350|16|20280;350|350|16.5|20840;350|350|17|21400;350|350|17.5|21960;350|350|18|22510;350|350|18.5|23060;350|350|19|23600;350|350|20|24680;350|350|21|25740;350|350|22|26780;400|400|7|10840;400|400|7.5|11580;400|400|8|12320;400|400|8.5|13060;400|400|9|13800;400|400|9.5|14530;400|400|10|15260;400|400|10.5|15880;400|400|11|16600;400|400|11.5|17300;400|400|12|18000;400|400|12.5|18700;400|400|13|19400;400|400|13.5|20090;400|400|14|20770;400|400|14.5|21450;400|400|15|22130;400|400|15.5|22810;400|400|16|23480;400|400|16.5|24140;400|400|17|24800;400|400|17.5|25460;400|400|18|26110;400|400|18.5|26760;400|400|19|27400;400|400|20|28680;400|400|21|29940;450|450|7|12240;450|450|7.5|13080;450|450|8|13920;450|450|8.5|14760;450|450|9|15600;450|450|9.5|16430;450|450|10|17260;450|450|10.5|17980;450|450|11|18800;450|450|11.5|19600;450|450|12|20400;450|450|12.5|21200;450|450|13|22000;450|450|13.5|22790;450|450|14|23570;450|450|14.5|24350;450|450|15|25130;450|450|15.5|25910;450|450|16|26680;450|450|16.5|27440;450|450|17|28200;450|450|17.5|28960;450|450|18|29710;450|450|18.5|30460;450|450|19|31200;450|450|20|32680;450|450|21|34140;450|450|22|35580;500|500|8|15520;500|500|8.5|16460;500|500|9|17400;500|500|9.5|18330;500|500|10|19260;500|500|10.5|20080;500|500|11|21000;500|500|11.5|21900;500|500|12|22800;500|500|12.5|23700;500|500|13|24600;500|500|13.5|25490;500|500|14|26370;500|500|14.5|27250;500|500|15|28130;500|500|15.5|29010;500|500|16|29880;500|500|16.5|30740;500|500|17|31600;500|500|17.5|32460;500|500|18|33310;500|500|18.5|34160;500|500|19|35000;500|500|20|36680;500|500|21|38340;500|500|22|39980",
  "RU-TP": "50|25|2|274;50|25|2.5|334;50|25|3|391;50|25|3.5|444;50|25|4|495;50|30|2|294;50|30|2.5|359;50|30|3|421;50|30|3.5|479;50|30|4|535;50|30|5|636;50|40|2|334;50|40|2.5|409;50|40|3|481;50|40|3.5|549;50|40|4|615;50|40|4.5|677;50|40|5|736;60|30|2.5|409;60|30|3|481;60|30|3.5|549;60|30|4|615;60|30|4.5|677;60|30|5|736;60|30|5.5|791;60|30|6|843;60|40|2|374;60|40|2.5|459;60|40|3|541;60|40|3.5|619;60|40|4|695;60|40|4.5|767;60|40|5|836;60|40|5.5|901;60|40|6|963;70|50|2|454;70|50|2.5|559;70|50|3|661;70|50|3.5|759;70|50|4|855;70|50|4.5|947;70|50|5|1036;70|50|5.5|1121;70|50|6|1203;80|40|2|454;80|40|2.5|559;80|40|3|661;80|40|3.5|759;80|40|4|855;80|40|4.5|947;80|40|5|1036;80|40|5.5|1121;80|40|6|1203;80|60|2|534;80|60|2.5|659;80|60|3|781;80|60|3.5|899;80|60|4|1015;80|60|4.5|1127;80|60|5|1236;80|60|5.5|1341;80|60|6|1443;80|60|6.5|1506;80|60|7|1596;80|70|3|841;80|70|3.5|969;80|70|4|1095;80|70|4.5|1217;80|70|5|1336;80|70|5.5|1451;80|70|6|1563;80|70|6.5|1636;80|70|7|1736;90|50|3|781;90|50|3.5|899;90|50|4|1015;90|50|4.5|1127;90|50|5|1236;90|50|5.5|1341;90|50|6|1443;90|50|6.5|1506;90|50|7|1596;90|60|3|841;90|60|3.5|969;90|60|4|1095;90|60|4.5|1217;90|60|5|1336;90|60|5.5|1451;90|60|6|1563;90|60|7|1736;100|40|3|781;100|40|3.5|899;100|40|4|1015;100|40|4.5|1127;100|40|5|1236;100|40|5.5|1341;100|40|6|1443;100|40|6.5|1506;100|40|7|1596;100|50|3|841;100|50|3.5|969;100|50|4|1095;100|50|4.5|1217;100|50|5|1336;100|50|5.5|1451;100|50|6|1563;100|50|6.5|1636;100|50|7|1736;100|60|3|901;100|60|3.5|1039;100|60|4|1175;100|60|4.5|1307;100|60|5|1436;100|60|5.5|1561;100|60|6|1683;100|60|6.5|1766;100|60|7|1876;120|40|3|901;120|40|3.5|1039;120|40|4|1175;120|40|4.5|1307;120|40|5|1436;120|40|5.5|1561;120|40|6|1683;120|40|6.5|1766;120|40|7|1876;120|60|3|1021;120|60|3.5|1179;120|60|4|1335;120|60|4.5|1487;120|60|5|1636;120|60|5.5|1781;120|60|6|1923;120|60|6.5|2026;120|60|7|2156;120|80|3|1141;120|80|3.5|1319;120|80|4|1495;120|80|4.5|1667;120|80|5|1836;120|80|5.5|2001;120|80|6|2163;120|80|6.5|2286;120|80|7|2436;140|60|3|1141;140|60|3.5|1319;140|60|4|1495;140|60|4.5|1667;140|60|5|1836;140|60|5.5|2001;140|60|6|2163;140|60|6.5|2286;140|60|7|2436;140|100|4|1815;140|100|4.5|2027;140|100|5|2236;140|100|5.5|2441;140|100|6|2643;140|100|6.5|2806;140|100|7|2996;140|110|4|1895;140|110|4.5|2117;140|110|5|2336;140|110|5.5|2551;140|110|6|2763;140|110|6.5|2936;140|110|7|3135;140|120|4|1975;140|120|4.5|2207;140|120|5|2436;140|120|5.5|2661;140|120|6|2883;140|120|6.5|3066;140|120|7|3276;140|120|7.5|3482;140|120|8|3684;150|50|3|1141;150|50|3.5|1319;150|50|4|1495;150|50|4.5|1667;150|50|5|1836;150|50|5.5|2001;150|50|6|2163;150|50|6.5|2286;150|50|7|2435;150|100|4|1895;150|100|4.5|2117;150|100|5|2336;150|100|5.5|2551;150|100|6|2763;150|100|6.5|2936;150|100|7|3136;150|130|4|2135;150|130|4.5|2387;150|130|5|2636;150|130|5.5|2881;150|130|6|3123;150|130|6.5|3326;150|130|7|3555;150|130|7.5|3782;150|130|8|4004;160|40|3|1141;160|40|3.5|1319;160|40|4|1495;160|40|4.5|1667;160|40|5|1836;160|40|5.5|2001;160|40|6|2163;160|40|6.5|2286;160|40|7|2436;160|80|4|1815;160|80|4.5|2027;160|80|5|2236;160|80|5.5|2441;160|80|6|2643;160|80|6.5|2806;160|80|7|2996;160|100|4|1975;160|100|4.5|2207;160|100|5|2436;160|100|5.5|2661;160|100|6|2883;160|100|6.5|3066;160|100|7|3276;160|100|7.5|3482;160|100|8|3684;160|120|4|2135;160|120|4.5|2387;160|120|5|2636;160|120|5.5|2881;160|120|6|3123;160|120|6.5|3326;160|120|7|3556;160|120|7.5|3782;160|120|8|4004;160|140|4|2295;160|140|4.5|2567;160|140|5|2836;160|140|5.5|3101;160|140|6|3363;160|140|6.5|3586;160|140|7|3836;160|140|7.5|4082;160|140|8|4324;180|60|4|1815;180|60|4.5|2027;180|60|5|2236;180|60|5.5|2441;180|60|6|2643;180|60|6.5|2806;180|60|7|2996;180|60|7.5|3182;180|60|8|3364;180|80|4|1975;180|80|4.5|2207;180|80|5|2436;180|80|5.5|2661;180|80|6|2883;180|80|6.5|3066;180|80|7|3276;180|80|7.5|3482;180|80|8|3684;180|100|4|2135;180|100|4.5|2387;180|100|5|2636;180|100|5.5|2881;180|100|6|3123;180|100|6.5|3326;180|100|7|3556;180|100|7.5|3782;180|100|8|4004;180|120|4|2295;180|120|4.5|2567;180|120|5|2836;180|120|5.5|3101;180|120|6|3363;180|120|6.5|3586;180|120|7|3835;180|120|7.5|4082;180|120|8|4324;180|140|4|2455;180|140|4.5|2747;180|140|5|3036;180|140|5.5|3321;180|140|6|3603;180|140|6.5|3846;180|140|7|4116;180|140|7.5|4382;180|140|8|4644;200|40|4|1815;200|40|4.5|2027;200|40|5|2236;200|40|5.5|2441;200|40|6|2643;200|40|6.5|2806;200|40|7|2996;200|80|4|2135;200|80|4.5|2387;200|80|5|2636;200|80|5.5|2881;200|80|6|3123;200|80|6.5|3326;200|80|7|3556;200|80|7.5|3782;200|80|8|4004;200|100|4|2295;200|100|4.5|2567;200|100|5|2836;200|100|5.5|3101;200|100|6|3363;200|100|6.5|3586;200|100|7|3836;200|100|7.5|4082;200|100|8|4324;200|120|4|2455;200|120|4.5|2747;200|120|5|3036;200|120|5.5|3321;200|120|6|3603;200|120|6.5|3846;200|120|7|4116;200|120|7.5|4382;200|120|8|4644;200|150|5|3336;200|150|5.5|3651;200|150|6|3963;200|150|6.5|4236;200|150|7|4535;200|160|5|3436;200|160|5.5|3761;200|160|6|4083;200|160|6.5|4366;200|160|7|4676;200|160|7.5|4982;200|160|8|5284;200|160|8.5|5583;200|160|9|5878;200|160|9.5|6169;200|160|10|6457;200|160|10.5|6654;200|160|11|6916;200|160|11.5|7182;200|160|12|7445;200|160|12.5|7703;200|160|13|7957;200|160|13.5|8207;200|160|14|8453;200|160|14.5|8695;200|160|15|8933;200|160|15.5|9166;200|160|16|9395;220|100|4|2455;220|100|4.5|2747;220|100|5|3036;220|100|5.5|3321;220|100|6|3603;220|100|6.5|3846;220|100|7|4116;220|100|7.5|4382;220|100|8|4644;220|140|5|3436;220|140|5.5|3761;220|140|6|4083;220|140|6.5|4366;220|140|7|4676;220|140|7.5|4982;220|140|8|5284;220|140|8.5|5582;220|140|9|5877;220|140|9.5|6169;220|140|10|6456;220|140|10.5|6645;220|140|11|6916;220|140|11.5|7182;220|140|12|7445;220|140|12.5|7703;240|120|5|3436;240|120|5.5|3761;240|120|6|4083;240|120|6.5|4366;240|120|7|4676;240|120|7.5|4982;240|120|8|5284;240|120|8.5|5582;240|120|9|5877;240|120|9.5|6169;240|120|10|6456;240|120|10.5|6645;240|120|11|6916;240|120|11.5|7182;240|120|12|7445;240|120|12.5|7703;240|120|13|7957;240|120|13.5|8207;240|120|14|8453;240|120|14.5|8695;240|120|15|8933;240|120|15.5|9166;240|120|16|9395;240|150|5|3736;240|150|5.5|4091;240|150|6|4443;240|150|6.5|4756;240|150|7|5095;240|150|7.5|5432;240|150|8|5764;240|160|5|3836;240|160|5.5|4201;240|160|6|4563;240|160|6.5|4886;240|160|7|5236;240|160|7.5|5582;240|160|8|5924;240|160|8.5|6263;240|160|9|6598;240|160|9.5|6929;240|160|10|7257;240|160|10.5|7486;240|160|11|7797;240|160|11.5|8103;240|160|12|8406;250|140|5|3736;250|140|5.5|4091;250|140|6|4443;250|140|6.5|4756;250|140|7|5095;250|140|7.5|5432;250|140|8|5764;250|150|5|3836;250|150|5.5|4201;250|150|6|4563;250|150|6.5|4886;250|150|7|5236;250|150|7.5|5582;250|150|8|5924;250|150|8.5|6262;250|150|9|6597;250|150|9.5|6929;250|150|10|7256;250|150|10.5|7485;250|150|11|7796;250|150|11.5|8102;250|150|12|8405;260|130|6|4443;260|130|6.5|4756;260|130|7|5096;260|130|7.5|5432;260|130|8|5764;260|130|8.5|6093;260|130|9|6418;260|130|9.5|6739;260|130|10|7057;260|130|10.5|7276;260|130|11|7577;260|130|11.5|7873;260|130|12|8166;260|140|5|3836;260|140|5.5|4201;260|140|6|4563;260|240|6|5763;260|240|6.5|6186;260|240|7|6635;260|240|7.5|7082;260|240|8|7524;260|240|8.5|7962;260|240|9|8397;260|240|9.5|8829;260|240|10|9170;260|240|10.5|9585;260|240|11|9996;260|240|11.5|10402;260|240|12|10805;300|100|6|4563;300|100|6.5|4886;300|100|7|5236;300|100|7.5|5582;300|100|8|5924;300|100|8.5|6263;300|100|9|6598;300|100|9.5|6929;300|100|10|7257;300|100|10.5|7485;300|100|11|7796;300|100|11.5|8102;300|100|12|8405;300|200|6|5763;300|200|6.5|6186;300|200|7|6636;300|200|7.5|7082;300|200|8|7524;300|200|8.5|7963;300|200|9|8398;300|200|9.5|8829;300|200|10|9257;300|200|10.5|9586;300|200|11|9997;300|200|11.5|10400;300|200|12|10810;320|180|6|5763;320|180|6.5|6186;320|180|7|6636;320|180|7.5|7082;320|180|8|7524;320|180|8.5|7963;320|180|9|8398;320|180|9.5|8829;320|180|10|9257;320|180|10.5|9586;320|180|11|9997;320|180|11.5|10400;320|180|12|10810;350|250|13.5|14690;350|250|14|15170;350|250|14.5|15650;350|250|15|16130;350|250|15.5|16610;350|250|16|17080;350|250|16.5|17540;350|250|18|18910;350|250|18.5|19360;350|250|19|19800;350|250|20|20680;350|250|21|21540;350|250|22|22380;350|300|6|7563;350|300|6.5|8136;350|300|7|8736;350|300|7.5|9332;350|300|8|9924;350|300|8.5|10510;350|300|9|11100;350|300|9.5|11680;350|300|10|12260;350|300|10.5|12740;350|300|11|13300;350|300|11.5|13850;350|300|12|14410;360|220|6.5|7486;360|220|7|8036;360|220|7.5|8582;360|220|8|9124;400|200|10|11260;400|200|10.5|11690;400|200|11|12200;400|200|11.5|12700;400|200|12|13210;400|200|12.5|13700;400|200|13|14200;400|200|13.5|14690;400|200|14|15170;400|200|14.5|15650;400|200|15|16130;400|200|15.5|16610;400|200|16|17080;400|300|6|8163;400|300|6.5|8786;400|300|7|9435;400|300|7.5|10080;400|300|8|10720;400|300|8.5|11360;400|300|9|12000;400|300|9.5|12630;400|300|10|13260;400|300|10.5|13780;400|300|11|14400;400|300|11.5|15000;400|300|12|15600;400|300|12.5|16200;400|300|13|16800;400|300|13.5|17390;400|300|14|17970;400|300|14.5|18550;400|300|15|19130;400|300|15.5|19710;400|300|16|20280;400|300|16.5|20840;400|300|17|21400;400|300|17.5|21960;400|300|18|22510;400|300|18.5|23060;400|300|19|23600;400|300|20|24680;400|300|21|25740;400|300|22|26780;400|350|17|24800;400|350|17.5|25460;400|350|18|26110;400|350|18.5|26760;400|350|19|27400;400|350|20|28680;400|350|21|29940;400|350|22|31180",
  "RU-TK80": "80|80|3|924;80|80|4|1216;80|80|5|1500;80|80|6|1776;100|100|3|1164;100|100|4|1536;100|100|5|1900;100|100|6|2256;120|120|3|1404;120|120|4|1856;120|120|5|2300;120|120|6|2736;140|140|4|2176;140|140|5|2700;140|140|6|3216;140|140|7|3724;140|140|8|4224;160|160|4|2496;160|160|5|3100;160|160|6|3696;160|160|7|4284;160|160|8|4864;180|180|5|3500;180|180|6|4176;180|180|7|4844;180|180|8|5504",
  "RU-TP80": "100|60|3|924;100|60|4|1216;100|60|5|1500;100|60|6|1776;120|80|3|1164;120|80|4|1536;120|80|5|1900;120|80|6|2256;140|60|4|1536;140|60|5|1900;140|60|6|2256;140|100|4|1856;140|100|5|2300;140|100|6|2736;140|100|7|3164;160|80|4|1856;160|80|5|2300;160|80|6|2736;160|80|7|3164;160|120|4|2176;160|120|5|2700;160|120|6|3216;160|120|7|3724;160|120|8|4224;180|60|5|2300;180|60|6|2736;180|60|7|3164;180|100|5|2700;180|100|6|3216;180|100|7|3724;180|100|8|4224;180|140|5|3100;180|140|6|3696;180|140|7|4284;180|140|8|4864;200|160|5|3500;200|160|6|4176;200|160|7|4844;200|160|8|5504",
  "RU-TK10": "10|10|0.8|28;10|10|0.9|31;10|10|1|34;10|10|1.2|39;10|10|1.4|44;15|15|0.8|44;15|15|0.9|49;15|15|1|54;15|15|1.2|64;15|15|1.4|73;15|15|1.5|77;20|20|0.8|60;20|20|0.9|67;20|20|1|74;20|20|1.2|89;20|20|1.4|101;20|20|1.5|107;20|20|2|137;25|25|0.8|76;25|25|0.9|85;25|25|1|94;25|25|1.2|112;25|25|1.4|129;25|25|1.5|137;25|25|2|177;25|25|2.5|214;25|25|3|241;30|30|0.8|92;30|30|0.9|103;30|30|1|114;30|30|1.2|136;30|30|1.3|146;30|30|1.4|157;30|30|1.5|167;30|30|2|217;30|30|2.5|264;30|30|3|301;30|30|3.5|339;30|30|4|375;35|35|0.8|108;35|35|0.9|121;35|35|1.4|185;35|35|1.5|197;35|35|2|257;35|35|2.5|314;35|35|3|361;35|35|3.5|409;35|35|4|455;35|35|4.5|497;35|35|5|536;40|40|1.4|213;40|40|1.5|227;40|40|2|297;40|40|2.5|364;40|40|3|421;40|40|3.5|479;40|40|4|535;40|40|4.5|587;40|40|5|636;40|40|5.5|681;40|40|6|723;42|42|3|445;42|42|3.5|507;42|42|4|567;42|42|4.5|623;42|42|5|676;42|42|5.5|725;42|42|6|771;45|45|2|337;45|45|3|481;45|45|3.5|549;45|45|4|615;45|45|4.5|677;45|45|5|736;45|45|5.5|791;45|45|6|843;45|45|6.5|856;45|45|7|896;45|45|7.5|932;45|45|8|964;50|50|2|377;50|50|2.5|464;50|50|3|541;50|50|3.5|619;50|50|4|695;50|50|4.5|767;50|50|5|836;50|50|5.5|901;50|50|6|963;50|50|6.5|986;50|50|7|1036;50|50|7.5|1082;50|50|8|1124;55|55|3|609;60|60|2|457;60|60|2.5|564;60|60|3|661;60|60|3.5|759;60|60|4|855;60|60|4.5|947;60|60|5|1036;60|60|5.5|1121;60|60|6|1203;60|60|6.5|1246;60|60|7|1316;60|60|7.5|1382;60|60|8|1444;70|70|2|537;70|70|2.5|664;70|70|3|781;70|70|3.5|899;70|70|4|1015;70|70|4.5|1127;70|70|5|1236;70|70|5.5|1341;70|70|6|1443;70|70|6.5|1506;70|70|7|1596;70|70|7.5|1682;70|70|8|1764;80|80|3|901;80|80|3.5|1039;80|80|4|1175;80|80|4.5|1307;80|80|5|1436;80|80|5.5|1561;80|80|6|1683;80|80|6.5|1766;80|80|7|1876;80|80|7.5|1982;80|80|8|2084;80|80|9|2297;80|80|10|2576;80|80|11|2721;90|90|3|1021;90|90|3.5|1179;90|90|4|1335;90|90|4.5|1467;90|90|5|1636;90|90|5.5|1781;90|90|6|1923;90|90|6.5|2026;90|90|7|2156;90|90|7.5|2282;90|90|8|2404;100|100|3|1141;100|100|3.5|1319;100|100|4|1495;100|100|4.5|1667;100|100|5|1836;100|100|5.5|2001;100|100|6|2163;100|100|6.5|2286;100|100|7|2436;100|100|7.5|2582;100|100|8|2724;100|100|9|3137;110|110|6|2403;110|110|6.5|2546;110|110|7|2716;110|110|7.5|2882;110|110|8|3044;110|110|8.5|3203;110|110|9|3358;120|120|3|1381;120|120|3.5|1599;120|120|4|1815;120|120|4.5|2027;120|120|5|2236;120|120|5.5|2441;120|120|6|2643;120|120|6.5|2806;120|120|7|2996;120|120|7.5|3182;120|120|8|3364;120|120|9|3857;140|140|4|2135;140|140|4.5|2387;140|140|5|2636;140|140|5.5|2881;140|140|6|3123;140|140|6.5|3326;140|140|7|3556;140|140|7.5|3782;140|140|8|4004;140|140|9|4577;150|150|4|2295;150|150|4.5|2567;150|150|5|2836;150|150|5.5|3101;150|150|6|3363;150|150|6.5|3586;150|150|7|3836;150|150|7.5|4082;150|150|8|4324;150|150|9|4937;150|150|10|5428;160|160|4|2455;160|160|4.5|2747;160|160|5|3036;160|160|5.5|3321;160|160|6|3603;160|160|6.5|3846;160|160|7|4116;160|160|7.5|4382;160|160|8|4644;180|180|3|2101;180|180|3.5|2439;180|180|4|2775;180|180|4.5|3107;180|180|5|3436;180|180|5.5|3761;180|180|6|4083;180|180|6.5|4366;180|180|7|4676;180|180|7.5|4982;180|180|8|5284;180|180|8.5|5583;180|180|9|5878;180|180|9.5|6169;180|180|10|6457;180|180|10.5|6646;180|180|11|6917;180|180|11.5|7183;180|180|12|7446;180|180|12.5|7704;180|180|13|7959;180|180|13.5|8209;180|180|14|8455;180|180|14.5|8697;180|180|15|8934;180|180|15.5|9168;180|180|16|9397;200|200|5|3836;200|200|5.5|4201;200|200|6|4563;200|200|6.5|4886;200|200|7|5236;200|200|7.5|5582;200|200|8|5924;200|200|8.5|6263;200|200|9|6598;200|200|9.5|6929;200|200|10|7257;200|200|10.5|7486;200|200|11|7797;200|200|11.5|8103;200|200|12|8406;250|250|6|5763;250|250|6.5|6186;250|250|7|6636;250|250|7.5|7082;250|250|8|7524;250|250|8.5|7963;250|250|9|8398;250|250|9.5|8829;250|250|10|9257;250|250|10.5|9586;250|250|11|9997;250|250|11.5|10403;250|250|12|10806;300|300|4|4695;300|300|4.5|5267;300|300|5|5836;300|300|5.5|6401;300|300|6|6963;300|300|6.5|7486;300|300|7|8036;300|300|7.5|8582;300|300|8|9124;300|300|8.5|9663;300|300|9|10198;300|300|9.5|10729;300|300|10|11257;300|300|10.5|11686;300|300|11|12197;300|300|11.5|12703;300|300|12|13206;300|300|12.5|13704;300|300|13|14199;300|300|13.5|14689;300|300|14|15175;300|300|14.5|15657;300|300|15|16134;300|300|15.5|16608;300|300|16|17077;300|300|16.5|17542;300|300|17|18004;300|300|17.5|18461;300|300|18|18913;300|300|18.5|19362;300|300|19|19807;300|300|20|20683;300|300|21|21543;300|300|22|22387;350|350|5|6836;350|350|5.5|7501;350|350|6|8163;350|350|6.5|8786;350|350|7|9436;350|350|7.5|10082;350|350|8|10724;350|350|8.5|11363;350|350|9|11998;350|350|9.5|12629;350|350|10|13257;350|350|10.5|13786;350|350|11|14397;350|350|11.5|15003;350|350|12|15606;350|350|12.5|16204;350|350|13|16799;350|350|13.5|17389;350|350|14|17975;350|350|14.5|18557;350|350|15|19134;350|350|15.5|19708;350|350|16|20277;350|350|16.5|20842;350|350|17|21404;350|350|17.5|21961;350|350|18|22513;350|350|18.5|23062;350|350|19|23607;350|350|20|24683;350|350|21|25743;350|350|22|26787;400|400|6|9363;400|400|6.5|10086;400|400|7|10836;400|400|7.5|11582;400|400|8|12324;400|400|8.5|13063;400|400|9|13798;400|400|9.5|14529;400|400|10|15257;400|400|10.5|15886;400|400|11|16597;400|400|11.5|17303;400|400|12|18006;400|400|12.5|18704;400|400|13|19399;400|400|13.5|20089;400|400|14|20775;400|400|14.5|21457;400|400|15|22134;400|400|15.5|22808;400|400|16|23477;400|400|16.5|24142;400|400|17|24804;400|400|17.5|25461;400|400|18|26113;400|400|18.5|26762;400|400|19|27407;400|400|20|28683;400|400|21|29943;400|400|22|31187;450|450|6|10563;450|450|6.5|11386;450|450|7|12236;450|450|7.5|13082;450|450|8|13924;450|450|8.5|14763;450|450|9|15598;450|450|9.5|16429;450|450|10|17257;450|450|10.5|17986;450|450|11|18797;450|450|11.5|19603;450|450|12|20406;450|450|12.5|21204;450|450|13|21999;450|450|13.5|22789;450|450|14|23575;450|450|14.5|24357;450|450|15|25134;450|450|15.5|25908;450|450|16|26677;450|450|16.5|27442;450|450|17|28204;450|450|17.5|28961;450|450|18|29713;450|450|18.5|30462;450|450|19|31207;450|450|20|32683;450|450|21|34143;450|450|22|35587;500|500|7|13636;500|500|7.5|14582;500|500|8|15524;500|500|8.5|16463;500|500|9|17398;500|500|9.5|18329;500|500|10|19257;500|500|10.5|20086;500|500|11|20997;500|500|11.5|21903;500|500|12|22806;500|500|12.5|23704;500|500|13|24599;500|500|13.5|25489;500|500|14|26375;500|500|14.5|27257;500|500|15|28134;500|500|15.5|29008;500|500|16|29877;500|500|16.5|30742;500|500|17|31604;500|500|17.5|32461;500|500|18|33313;500|500|18.5|34162;500|500|19|35007;500|500|20|36683;500|500|21|38343;500|500|22|39987",
  "RU-TP10": "15|10|1|44;15|10|1.5|62;15|10|2|77;20|10|1|54;20|10|1.2|63;20|10|1.5|77;20|10|2|97;20|15|1|64;20|15|1.2|75;20|15|1.5|92;20|15|2|117;20|15|2.5|139;25|10|1|64;25|10|1.5|92;25|10|2|117;25|10|2.5|139;25|15|0.8|60;25|15|0.9|67;25|15|1|74;25|15|1.2|87;25|15|1.5|107;25|15|2|137;25|15|2.5|164;28|25|0.8|81;28|25|0.9|90;28|25|1|99;28|25|1.2|119;30|10|1|74;30|10|1.5|107;30|10|2|137;30|10|2.5|164;30|10|3|181;30|15|0.8|68;30|15|0.9|76;30|15|1|84;30|15|1.2|99;30|15|1.5|122;30|15|2|157;30|15|2.5|189;30|15|3|211;30|20|0.8|76;30|20|0.9|85;30|20|1|94;30|20|1.2|112;30|20|1.5|137;30|20|2|177;30|20|2.5|214;30|20|3|241;35|15|0.8|76;35|15|0.9|85;35|15|1|94;35|15|1.2|112;35|15|1.5|137;35|15|2|177;35|15|2.5|214;35|15|3|241;35|15|3.5|269;35|20|0.8|84;35|20|0.9|94;35|20|1|104;35|20|1.2|124;35|20|1.5|152;35|20|2|197;35|20|2.5|239;35|20|3|271;35|20|3.5|304;35|25|1.5|167;35|25|2|217;35|25|2.5|264;35|25|3|301;35|25|3.5|339;35|30|0.8|100;35|30|0.9|112;35|30|1|124;35|30|1.2|148;35|30|1.5|182;35|30|2|237;40|15|2|197;40|15|2.5|239;40|15|3|271;40|15|3.5|304;40|15|4|335;40|20|0.8|92;40|20|0.9|94;40|20|1|104;40|20|1.2|124;40|20|1.5|152;40|20|2|217;40|20|2.5|264;40|20|3|301;40|20|3.5|339;40|20|4|375;40|25|1.5|182;40|25|2|237;40|25|2.5|289;40|25|3|331;40|25|3.5|374;40|25|4|415;40|30|1.5|197;40|30|2|257;40|30|2.5|314;40|30|3|361;40|30|3.5|409;40|30|4|455;45|20|2|237;45|20|2.5|289;45|20|3|331;45|20|3.5|374;45|20|4|415;45|30|2|277;45|30|2.5|339;45|30|3|391;45|30|3.5|444;45|30|4|495;50|25|1.5|212;50|25|2|277;50|25|2.5|339;50|25|3|391;50|25|3.5|444;50|25|4|495;50|30|1.5|227;50|30|2|297;50|30|2.5|364;50|30|3|421;50|30|3.5|479;50|30|4|535;50|30|4.5|587;50|30|5|636;50|35|1.5|242;50|35|2|317;50|35|2.2|346;50|35|2.5|389;50|35|3|458;50|35|3.5|525;50|35|4|588;50|40|2|337;50|40|2.5|414;50|40|3|481;50|40|3.5|549;50|40|4|615;50|40|4.5|677;50|40|5|736;60|20|2|297;60|25|2.5|389;60|25|3|451;60|25|3.5|514;60|25|4|575;60|25|4.5|632;60|25|5|686;60|30|1.5|257;60|30|2|337;60|30|2.5|414;60|30|3|481;60|30|3.5|549;60|30|4|615;60|30|4.5|677;60|30|5|736;60|30|5.5|791;60|30|6|843;60|40|1.5|287;60|40|2|377;60|40|2.5|464;60|40|3|541;60|40|3.5|619;60|40|4|695;60|40|4.5|767;60|40|5|836;60|40|5.5|901;60|40|6|963;70|30|3|541;70|30|3.5|619;70|30|4|695;70|30|4.5|767;70|30|5|836;70|30|5.5|901;70|30|6|963;70|40|3|601;70|40|3.5|689;70|40|4|775;70|40|4.5|857;70|40|5|936;70|40|5.5|1011;70|40|6|1083;70|50|2|457;70|50|2.5|564;70|50|3|661;70|50|3.5|759;70|50|4|855;70|50|4.5|947;70|50|5|1036;70|50|5.5|1121;70|50|6|1203;80|40|2|457;80|40|2.5|564;80|40|3|661;80|40|3.5|759;80|40|4|855;80|40|4.5|947;80|40|5|1036;80|40|5.5|1121;80|40|6|1203;80|40|6.5|1246;80|40|7|1316;80|50|3|728;80|50|3.5|840;80|50|4|948;80|60|2|537;80|60|2.5|664;80|60|3|781;80|60|3.5|899;80|60|4|1015;80|60|4.5|1127;80|60|5|1236;80|60|5.5|1341;80|60|6|1443;80|60|6.5|1506;80|60|7|1596;80|70|3|841;80|70|3.5|969;80|70|4|1095;80|70|4.5|1217;80|70|5|1336;80|70|5.5|1451;80|70|6|1563;80|70|6.5|1636;80|70|7|1736;90|40|3.5|840;90|40|4|948;90|40|5|1157;90|40|6|1354;90|40|7|1540;90|50|3|781;90|50|3.5|899;90|50|4|1015;90|50|4.5|1127;90|50|5|1236;90|50|5.5|1341;90|50|6|1443;90|50|6.5|1506;90|50|7|1596;90|60|3|841;90|60|3.5|969;90|60|4|1095;90|60|4.5|1217;90|60|5|1336;90|60|5.5|1451;90|60|6|1563;90|60|6.5|1636;90|60|7|1736;100|40|3|781;100|40|3.5|899;100|40|4|1015;100|40|4.5|1127;100|40|5|1236;100|40|5.5|1341;100|40|6|1443;100|40|6.5|1506;100|40|7|1596;100|50|3|841;100|50|3.5|969;100|50|4|1095;100|50|4.5|1217;100|50|5|1336;100|50|5.5|1451;100|50|6|1563;100|50|6.5|1636;100|50|7|1736;100|60|3|901;100|60|3.5|1039;100|60|4|1175;100|60|4.5|1307;100|60|5|1436;100|60|5.5|1561;100|60|6|1683;100|60|6.5|1766;100|60|7|1876;100|70|4|1268;100|70|5|1558;100|70|6|1834;100|70|7|2100;110|40|4|1095;110|40|4.5|1217;110|40|5|1336;110|40|5.5|1451;110|40|6|1563;110|40|6.5|1636;110|40|7|1736;110|50|4|1175;110|50|4.5|1307;110|50|5|1436;110|50|5.5|1561;110|50|6|1683;110|50|6.5|1766;110|50|7|1876;110|60|4|1255;110|60|4.5|1397;110|60|5|1536;110|60|5.5|1671;110|60|6|1803;110|60|6.5|1896;110|60|7|2016;120|40|3|901;120|40|3.5|1039;120|40|4|1175;120|40|4.5|1307;120|40|5|1436;120|40|5.5|1561;120|40|6|1683;120|40|6.5|1766;120|40|7|1876;120|40|8|1921;120|60|3|1021;120|60|3.5|1179;120|60|4|1335;120|60|4.5|1487;120|60|5|1636;120|60|5.5|1781;120|60|6|1923;120|60|6.5|2026;120|60|7|2156;120|60|8|2422;120|80|3|1141;120|80|3.5|1319;120|80|4|1495;120|80|4.5|1667;120|80|5|1836;120|80|5.5|2001;120|80|6|2163;120|80|6.5|2286;120|80|7|2436;120|80|8|2634;140|60|3|1141;140|60|3.5|1319;140|60|4|1495;140|60|4.5|1667;140|60|5|1836;140|60|5.5|2001;140|60|6|2163;140|60|6.5|2286;140|60|7|2436;140|60|8|2834;140|80|5|2057;140|80|6|2434;140|80|7|2800;140|80|8|3154;140|100|4|1815;140|100|4.5|2027;140|100|5|2236;140|100|5.5|2441;140|100|6|2643;140|100|6.5|2806;140|100|7|2996;140|120|4|1975;140|120|4.5|2207;140|120|5|2436;140|120|5.5|2661;140|120|6|2883;140|120|6.5|3066;140|120|7|3276;140|120|7.5|3482;140|120|8|3684;140|120|9|4217;150|50|3|1141;150|50|3.5|1319;150|50|4|1495;150|50|4.5|1667;150|50|5|1836;150|50|5.5|2001;150|50|6|2163;150|50|6.5|2286;150|50|7|2436;150|80|6|2523;150|80|6.5|2676;150|80|7|2856;150|80|7.5|3032;150|80|8|3204;150|80|8.5|3373;150|80|9|3538;150|80|9.5|3699;150|80|10|3857;150|100|4|1895;150|100|4.5|2117;150|100|5|2336;150|100|5.5|2551;150|100|6|2763;150|100|6.5|2936;150|100|7|3136;150|100|7.5|3332;150|100|8|3524;150|100|8.5|3713;150|100|9|3898;150|100|9.5|4079;150|100|10|4257;150|130|4|2135;150|130|4.5|2387;150|130|5|2636;150|130|6|3123;150|130|6.5|3326;150|130|7|3556;150|130|7.5|3782;150|130|8|4004;160|40|3|1141;160|40|3.5|1319;160|40|4|1495;160|40|4.5|1667;160|40|5|1836;160|40|5.5|2001;160|40|6|2163;160|40|6.5|2286;160|40|7|2436;160|80|4|1815;160|80|4.5|2027;160|80|5|2236;160|80|5.5|2441;160|80|6|2643;160|80|6.5|2806;160|80|7|2996;160|100|4|1975;160|100|4.5|2207;160|100|5|2436;160|100|5.5|2661;160|100|6|2883;160|100|6.5|3066;160|100|7|3276;160|100|7.5|3482;160|100|8|3684;160|120|4|2135;160|120|4.5|2387;160|120|5|2636;160|120|5.5|2881;160|120|6|3123;160|120|6.5|3326;160|120|7|3556;160|120|7.5|3782;160|120|8|4004;160|140|4|2295;160|140|4.5|2567;160|140|5|2836;160|140|5.5|3101;160|140|6|3363;160|140|6.5|3586;160|140|7|3836;160|140|7.5|4082;160|140|8|4324;180|60|4|1815;180|60|4.5|2027;180|60|5|2236;180|60|5.5|2441;180|60|6|2643;180|60|6.5|2806;180|60|7|2996;180|60|7.5|3182;180|60|8|3364;180|80|4|1975;180|80|4.5|2207;180|80|5|2436;180|80|5.5|2661;180|80|6|2883;180|80|6.5|3066;180|80|7|3276;180|80|7.5|3482;180|80|8|3684;180|80|8.5|3883;180|80|9|4078;180|80|9.5|4269;180|80|10|4457;180|80|10.5|4640;180|80|11|4821;180|80|11.5|4997;180|80|12|5170;180|100|4|2135;180|100|4.5|2387;180|100|5|2636;180|100|5.5|2881;180|100|6|3123;180|100|6.5|3326;180|100|7|3556;180|100|7.5|3782;180|100|8|4004;180|100|8.5|4223;180|100|9|4438;180|100|9.5|4649;180|100|10|4857;180|100|10.5|5060;180|100|11|5261;180|100|11.5|5457;180|100|12|5650;180|120|4|2295;180|120|4.5|2567;180|120|5|2836;180|120|5.5|3101;180|120|6|3363;180|120|6.5|3586;180|120|7|3836;180|120|7.5|4082;180|120|8|4324;180|140|4|2455;180|140|4.5|2747;180|140|5|3036;180|140|5.5|3321;180|140|6|3603;180|140|6.5|3846;180|140|7|4116;180|140|7.5|4382;180|140|8|4644;180|150|8|4804;180|150|8.5|5073;180|150|9|5338;180|150|9.5|5599;180|150|10|5857;180|150|10.5|6110;180|150|11|6361;180|150|11.5|6607;180|150|12|6850;200|40|4|1815;200|40|4.5|2027;200|40|5|2236;200|40|5.5|2441;200|40|6|2643;200|40|6.5|2806;200|40|7|2996;200|80|4|2135;200|80|4.5|2387;200|80|5|2636;200|80|5.5|2881;200|80|6|3123;200|80|6.5|3326;200|80|7|3556;200|80|7.5|3782;200|80|8|4004;200|100|4|2295;200|100|4.5|2567;200|100|5|2836;200|100|5.5|3101;200|100|6|3363;200|100|6.5|3586;200|100|7|3836;200|100|7.5|4082;200|100|8|4324;200|120|4|2455;200|120|4.5|2747;200|120|5|3036;200|120|5.5|3321;200|120|6|3603;200|120|6.5|3846;200|120|7|4116;200|120|7.5|4382;200|120|8|4644;200|150|5|3336;200|150|5.5|3651;200|150|6|3963;200|150|6.5|4236;200|150|7|4536;200|160|5|3436;200|160|5.5|3761;200|160|6|4083;200|160|6.5|4366;200|160|7|4676;200|160|7.5|4982;200|160|8|5284;200|160|8.5|5583;200|160|9|5878;200|160|9.5|6169;200|160|10|6457;220|100|4|2455;220|100|4.5|2747;220|100|5|3036;220|100|5.5|3321;220|100|6|3603;220|100|6.5|3846;220|100|7|4116;220|100|7.5|4382;220|100|8|4644;220|140|5|3436;220|140|5.5|3761;220|140|6|4083;220|140|6.5|4366;220|140|7|4676;220|140|7.5|4982;220|140|8|5284;240|120|3|2101;240|120|3.5|2439;240|120|4|2775;240|120|4.5|3107;240|120|5|3436;240|120|5.5|3761;240|120|6|4083;240|120|6.5|4366;240|120|7|4676;240|120|7.5|4982;240|120|8|5284;240|120|8.5|5583;240|120|9|5878;240|120|9.5|6169;240|120|10|6457;240|120|10.5|6740;240|120|11|7021;240|120|11.5|7297;240|120|12|7570;240|120|13|8104;240|120|14|8623;240|120|15|9127;240|120|16|9617;240|150|5|3736;240|150|5.5|4091;240|150|6|4443;240|150|6.5|4756;240|150|7|5096;240|150|7.5|5432;240|150|8|5764;240|160|5|3836;240|160|5.5|4201;240|160|6|4563;240|160|6.5|4886;240|160|7|5236;240|160|7.5|5582;240|160|8|5924;240|160|8.5|6263;240|160|9|6598;240|160|9.5|6929;240|160|10|7257;240|160|10.5|7486;240|160|11|7797;240|160|11.5|8103;240|160|12|8406;250|140|5|3736;250|140|5.5|4091;250|140|6|4443;250|140|6.5|4756;250|140|7|5096;250|140|7.5|5432;250|140|8|5764;250|150|5|3836;250|150|5.5|4201;250|150|6|4563;250|150|6.5|4886;250|150|7|5236;250|150|7.5|5582;250|150|8|5924;250|150|8.5|6263;250|150|9|6598;250|150|9.5|6929;250|150|10|7257;250|150|10.5|7486;250|150|11|7797;250|150|11.5|8103;250|150|12|8406;260|130|6|4443;260|130|6.5|4756;260|130|7|5096;260|130|7.5|5432;260|130|8|5764;260|130|8.5|6093;260|130|9|6418;260|130|9.5|6739;260|130|10|7057;260|130|10.5|7276;260|130|11|7577;260|130|11.5|7873;260|130|12|8166;260|140|5|3836;260|140|5.5|4201;260|140|6|4563;260|140|6.5|4886;260|140|7|5236;260|140|7.5|5582;260|140|8|5924;260|140|8.5|6263;260|140|9|6598;260|140|9.5|6929;260|140|10|7257;260|140|10.5|7486;260|140|11|7797;260|140|11.5|8103;260|140|12|8406;260|240|6|5763;260|240|6.5|6186;260|240|7|6636;260|240|7.5|7082;260|240|8|7524;260|240|8.5|7963;260|240|9|8398;260|240|9.5|8829;260|240|10|9257;260|240|10.5|9586;260|240|11|9997;260|240|11.5|10403;260|240|12|10806;300|100|6|4563;300|100|6.5|4886;300|100|7|5236;300|100|7.5|5582;300|100|8|5924;300|100|8.5|6263;300|100|9|6598;300|100|9.5|6929;300|100|10|7257;300|100|10.5|7486;300|100|11|7797;300|100|11.5|8103;300|100|12|8406;300|200|6|5763;300|200|6.5|6186;300|200|7|6636;300|200|7.5|7082;300|200|8|7524;300|200|8.5|7963;300|200|9|8398;300|200|9.5|8829;300|200|10|9257;300|200|10.5|9586;300|200|11|9997;300|200|11.5|10403;300|200|12|10806;320|180|6|5763;320|180|6.5|6186;320|180|7|6636;320|180|7.5|7082;320|180|8|7524;320|180|8.5|7963;320|180|9|8398;320|180|9.5|8829;320|180|10|9257;320|180|10.5|9586;320|180|11|9997;320|180|11.5|10403;320|180|12|10806;350|150|6|5763;350|150|6.5|6186;350|150|7|6636;350|150|7.5|7082;350|150|8|7524;350|150|8.5|7963;350|150|9|8398;350|150|9.5|8829;350|150|10|9257;350|150|10.5|9586;350|150|11|9997;350|150|11.5|10403;350|150|12|10806;350|250|4|4695;350|250|4.5|5267;350|250|5|5836;350|250|5.5|6401;350|250|6|6963;350|250|6.5|7486;350|250|7|8036;350|250|7.5|8582;350|250|8|9124;350|250|8.5|9663;350|250|9|10198;350|250|9.5|10729;350|250|10|11257;350|250|10.5|11686;350|250|11|12197;350|250|11.5|12703;350|250|12|13206;350|250|13|14199;350|250|14|15175;350|250|15|16134;350|250|16|17077;350|250|17|18004;350|250|18|18913;350|250|19|19807;350|250|20|20683;350|250|21|21543;350|250|22|22387;350|300|6|7563;350|300|6.5|8136;350|300|7|8736;350|300|7.5|9332;350|300|8|9924;350|300|8.5|10513;350|300|9|11098;350|300|9.5|11679;350|300|10|12257;350|300|10.5|12736;350|300|11|13297;350|300|11.5|13853;350|300|12|14406;380|220|6|6963;380|220|6.5|7486;380|220|7|8036;380|220|7.5|8582;380|220|8|9124;400|200|4|4695;400|200|4.5|5267;400|200|5|5836;400|200|5.5|6401;400|200|6|6963;400|200|6.5|7486;400|200|7|8036;400|200|7.5|8582;400|200|8|9124;400|200|8.5|9663;400|200|9|10198;400|200|9.5|10729",
  "RU-TK03": "40|40|2|294;40|40|2.5|359;40|40|3|421;40|40|3.5|479;40|40|4|535;50|50|2|374;50|50|2.5|459;50|50|3|541;50|50|3.5|619;50|50|4|695;50|50|4.5|767;50|50|5|836;50|50|5.5|901;50|50|6|963;60|60|2|454;60|60|2.5|559;60|60|3|661;60|60|3.5|759;60|60|4|855;60|60|4.5|947;60|60|5|1036;60|60|5.5|1121;60|60|6|1203;70|70|2|534;70|70|2.5|659;70|70|3|781;70|70|3.5|899;70|70|4|1015;70|70|4.5|1127;70|70|5|1236;70|70|5.5|1341;70|70|6|1443;70|70|6.5|1506;70|70|7|1596;80|80|3|901;80|80|3.5|1039;80|80|4|1175;80|80|4.5|1307;80|80|5|1436;80|80|5.5|1561;80|80|6|1683;80|80|6.5|1766;80|80|7|1876;80|80|7.5|1982;80|80|8|2084;90|90|3|1021;90|90|3.5|1179;90|90|4|1335;90|90|4.5|1487;90|90|5|1636;90|90|5.5|1781;90|90|6|1923;90|90|6.5|2026;90|90|7|2156;90|90|7.5|2282;90|90|8|2404;100|100|3|1141;100|100|3.5|1319;100|100|4|1495;100|100|4.5|1667;100|100|5|1836;100|100|5.5|2001;100|100|6|2163;100|100|6.5|2286;100|100|7|2436;100|100|7.5|2582;100|100|8|2724;120|120|3|1381;120|120|3.5|1599;120|120|4|1815;120|120|4.5|2027;120|120|5|2236;120|120|5.5|2441;120|120|6|2643;120|120|6.5|2806;120|120|7|2996;120|120|7.5|3182;120|120|8|3364;140|140|4|2135;140|140|4.5|2387;140|140|5|2636;140|140|5.5|2881;140|140|6|3123;140|140|6.5|3326;140|140|7|3556;140|140|7.5|3782;140|140|8|4004;150|150|4|2295;150|150|4.5|2567;150|150|5|2836;150|150|5.5|3101;150|150|6|3363;150|150|6.5|3586;150|150|7|3836;150|150|7.5|4082;150|150|8|4324;160|160|4|2455;160|160|4.5|2747;160|160|5|3036;160|160|5.5|3321;160|160|6|3603;160|160|6.5|3846;160|160|7|4116;160|160|7.5|4382;160|160|8|4644;180|180|5|3436;180|180|5.5|3761;180|180|6|4083;180|180|6.5|4366;180|180|7|4676;180|180|7.5|4982;180|180|8|5284;180|180|8.5|5583;180|180|9|5878;180|180|9.5|6169;180|180|10|6457;200|200|6|4563;200|200|6.5|4886;200|200|7|5236;200|200|7.5|5582;200|200|8|5924;200|200|8.5|6263;200|200|9|6598;200|200|9.5|6929;200|200|10|7257;200|200|10.5|7486;200|200|11|7797;200|200|11.5|8103;200|200|12|8406;250|250|6|5763;250|250|6.5|6186;250|250|7|6636;250|250|7.5|7082;250|250|8|7524;250|250|8.5|7963;250|250|9|8398;250|250|9.5|8829;250|250|10|9257;250|250|10.5|9586;250|250|11|9997;250|250|11.5|10400;250|250|12|10810;300|300|6|6963;300|300|6.5|7486;300|300|7|8036;300|300|7.5|8582;300|300|8|9124;300|300|8.5|9663;300|300|9|10200;300|300|9.5|10730;300|300|10|11260;300|300|10.5|11690;300|300|11|12200;300|300|11.5|12700;300|300|12|13210",
  "RU-TP03": "50|25|2|274;50|25|2.5|334;50|25|3|391;50|25|3.5|444;50|25|4|495;50|30|2|294;50|30|2.5|359;50|30|3|421;50|30|3.5|479;50|30|4|535;50|30|5|636;50|40|2|334;50|40|2.5|409;50|40|3|481;50|40|3.5|549;50|40|4|615;50|40|4.5|677;50|40|5|736;60|30|2|334;60|30|2.5|409;60|30|3|481;60|30|3.5|549;60|30|4|615;60|30|4.5|677;60|30|5|736;60|30|5.5|791;60|30|6|843;60|40|2|374;60|40|2.5|459;60|40|3|541;60|40|3.5|619;60|40|4|695;60|40|4.5|767;60|40|5|836;60|40|5.5|901;60|40|6|963;70|50|2|454;70|50|2.5|559;70|50|3|661;70|50|3.5|759;70|50|4|855;70|50|4.5|947;70|50|5|1036;70|50|5.5|1121;70|50|6|1203;80|40|2|454;80|40|2.5|559;80|40|3|661;80|40|3.5|759;80|40|4|855;80|40|4.5|947;80|40|5|1036;80|40|5.5|1121;80|40|6|1203;80|60|2|534;80|60|2.5|659;80|60|3|781;80|60|3.5|899;80|60|4|1015;80|60|4.5|1127;80|60|5|1236;80|60|5.5|1341;80|60|6|1443;80|60|6.5|1506;80|60|7|1596;80|70|3|841;80|70|3.5|969;80|70|4|1095;80|70|4.5|1217;80|70|5|1336;80|70|5.5|1451;80|70|6|1563;80|70|6.5|1636;80|70|7|1736;90|50|3|781;90|50|3.5|899;90|50|4|1015;90|50|4.5|1127;90|50|5|1236;90|50|5.5|1341;90|50|6|1443;90|50|6.5|1506;90|50|7|1596;90|60|3|841;90|60|3.5|969;90|60|4|1095;90|60|4.5|1217;90|60|5|1336;90|60|5.5|1451;90|60|6|1563;90|60|7|1736;100|40|3|781;100|40|3.5|899;100|40|4|1015;100|40|4.5|1127;100|40|5|1236;100|40|5.5|1341;100|40|6|1443;100|40|6.5|1506;100|40|7|1596;100|50|3|841;100|50|3.5|969;100|50|4|1095;100|50|4.5|1217;100|50|5|1336;100|50|5.5|1451;100|50|6|1563;100|50|6.5|1636;100|50|7|1736;100|60|3|901;100|60|3.5|1039;100|60|4|1175;100|60|4.5|1307;100|60|5|1436;100|60|5.5|1561;100|60|6|1683;100|60|6.5|1766;100|60|7|1876;120|40|3|901;120|40|3.5|1039;120|40|4|1175;120|40|4.5|1307;120|40|5|1436;120|40|5.5|1561;120|40|6|1683;120|40|6.5|1766;120|40|7|1876;120|60|3|1021;120|60|3.5|1179;120|60|4|1335;120|60|4.5|1487;120|60|5|1636;120|60|5.5|1781;120|60|6|1923;120|60|6.5|2026;120|60|7|2156;120|80|3|1141;120|80|3.5|1319;120|80|4|1495;120|80|4.5|1667;120|80|5|1836;120|80|5.5|2001;120|80|6|2163;120|80|6.5|2286;120|80|7|2436;140|60|3|1141;140|60|3.5|1319;140|60|4|1495;140|60|4.5|1667;140|60|5|1836;140|60|5.5|2001;140|60|6|2163;140|60|6.5|2286;140|60|7|2436;140|100|4|1815;140|100|4.5|2027;140|100|5|2236;140|100|5.5|2441;140|100|6|2643;140|100|6.5|2806;140|100|7|2996;140|120|4|1975;140|120|4.5|2207;140|120|5|2436;140|120|5.5|2661;140|120|6|2883;140|120|6.5|3066;140|120|7|3276;140|120|7.5|3482;140|120|8|3684;150|100|4|1895;150|100|4.5|2117;150|100|5|2336;150|100|5.5|2551;150|100|6|2763;150|100|6.5|2936;150|100|7|3136;160|40|3|1141;160|40|3.5|1319;160|40|4|1495;160|40|4.5|1667;160|40|5|1836;160|40|5.5|2001;160|40|6|2163;160|40|6.5|2286;160|40|7|2436;160|80|4|1815;160|80|4.5|2027;160|80|5|2236;160|80|5.5|2441;160|80|6|2643;160|80|6.5|2806;160|80|7|2996;160|100|4|1975;160|100|4.5|2207;160|100|5|2436;160|100|5.5|2661;160|100|6|2883;160|100|6.5|3066;160|100|7|3276;160|100|7.5|3482;160|100|8|3684;160|120|4|2135;160|120|4.5|2387;160|120|5|2636;160|120|5.5|2881;160|120|6|3123;160|120|6.5|3326;160|120|7|3556;160|120|7.5|3782;160|120|8|4004;160|140|5|2836;160|140|5.5|3101;160|140|6|3363;160|140|6.5|3586;160|140|7|3836;160|140|7.5|4082;160|140|8|4324;180|60|4|1815;180|60|4.5|2027;180|60|5|2236;180|60|5.5|2441;180|60|6|2643;180|60|6.5|2806;180|60|7|2996;180|60|7.5|3182;180|60|8|3364;180|80|4|1975;180|80|4.5|2207;180|80|5|2436;180|80|5.5|2661;180|80|6|2883;180|80|6.5|3066;180|80|7|3276;180|80|7.5|3482;180|80|8|3684;180|100|4|2135;180|100|4.5|2387;180|100|5|2636;180|100|5.5|2881;180|100|6|3123;180|100|6.5|3326;180|100|7|3556;180|100|7.5|3782;180|100|8|4004;180|140|4|2455;180|140|4.5|2747;180|140|5|3036;180|140|5.5|3321;180|140|6|3603;180|140|6.5|3846;180|140|7|4116;180|140|7.5|4382;180|140|8|4644;200|40|4|1815;200|40|4.5|2027;200|40|5|2236;200|40|5.5|2441;200|40|6|2643;200|40|6.5|2806;200|40|7|2996;200|80|4|2135;200|80|4.5|2387;200|80|5|2636;200|80|5.5|2881;200|80|6|3123;200|80|6.5|3326;200|80|7|3556;200|80|7.5|3782;200|80|8|4004;200|100|4|2295;200|100|4.5|2567;200|100|5|2836;200|100|5.5|3101;200|100|6|3363;200|100|6.5|3586;200|100|7|3836;200|100|7.5|4082;200|100|8|4324;200|120|4|2455;200|120|4.5|2747;200|120|5|3036;200|120|5.5|3321;200|120|6|3603;200|120|6.5|3846;200|120|7|4116;200|120|7.5|4382;200|120|8|4644;200|160|5|3436;200|160|5.5|3761;200|160|6|4083;200|160|6.5|4366;200|160|7|4676;200|160|7.5|4982;200|160|8|5284;200|160|8.5|5583;200|160|9|5878;200|160|9.5|6169;200|160|10|6457;220|100|4|2455;220|100|4.5|2747;220|100|5|3036;220|100|5.5|3321;220|100|6|3603;220|100|6.5|3846;220|100|7|4116;220|100|7.5|4382;220|100|8|4644;220|140|5|3436;220|140|5.5|3761;220|140|6|4083;220|140|6.5|4366;220|140|7|4676;220|140|7.5|4982;220|140|8|5284;240|120|5|3436;240|120|5.5|3761;240|120|6|4083;240|120|6.5|4366;240|120|7|4676;240|120|7.5|4982;240|120|8|5284;240|160|6|4563;240|160|6.5|4886;240|160|7|5236;240|160|7.5|5582;240|160|8|5924;240|160|8.5|6263;240|160|9|6598;240|160|9.5|6929;240|160|10|7257;240|160|10.5|7486;240|160|11|7797;240|160|11.5|8103;240|160|12|8406;250|150|6|4563;250|150|6.5|4886;250|150|7|5236;250|150|7.5|5582;250|150|8|5924;260|130|6|4443;260|130|6.5|4756;260|130|7|5096;260|130|7.5|5432;260|130|8|5764;260|130|8.5|6093;260|130|9|6418;260|130|9.5|6739;260|130|10|7057;260|130|10.5|7276;260|130|11|7577;260|130|11.5|7873;260|130|12|8166;300|100|6|4563;300|100|6.5|4886;300|100|7|5236;300|100|7.5|5582;300|100|8|5924;300|100|8.5|6263;300|100|9|6598;300|100|9.5|6929;300|100|10|7257;300|200|6|5763;300|200|6.5|6186;300|200|7|6636;300|200|7.5|7082;300|200|8|7524;300|200|8.5|7963;300|200|9|8398;300|200|9.5|8829;300|200|10|9257;300|200|10.5|9586;300|200|11|9997;300|200|11.5|10400;300|200|12|10810;320|180|6|5763;320|180|6.5|6186;320|180|7|6636;320|180|7.5|7082;320|180|8|7524;320|180|8.5|7963;320|180|9|8398;320|180|9.5|8829;320|180|10|9257;320|180|10.5|9586;320|180|11|9997;320|180|11.5|10400;320|180|12|10810;350|250|6|6963;350|250|6.5|7486;350|250|7|8036;350|250|7.5|8582;350|250|8|9124;350|250|8.5|9663;350|250|9|10200;350|250|9.5|10730;350|250|10|11260;350|250|10.5|11690;350|250|11|12200;350|250|11.5|12700;350|250|12|13210;350|300|6|7563;350|300|6.5|8136;350|300|7|8736;350|300|7.5|9332;350|300|8|9924;350|300|8.5|10510;350|300|9|11100;350|300|9.5|11680;350|300|10|12260;350|300|10.5|12740;350|300|11|13300;350|300|11.5|13850;350|300|12|14410;380|220|6|6963;380|220|6.5|7486;380|220|7|8036;380|220|7.5|8582;380|220|8|9124;380|220|10|11260;380|220|10.5|11690;400|200|11|12200;400|200|11.5|12700;400|200|12|13210",
  "RU-O": "42|3|367.6;42|3.5|423.3;42|4|477.5;48|3|424.1;48|3.5|489.3;48|4|552.9;54|3|480.7;54|3.5|555.3;54|4|628.3;57|3|508.9;57|3.5|588.3;57|4|666;60|3|537.2;60|3.5|621.2;60|4|703.7;70|3|631.5;70|3.5|731.2;70|4|829.4;73|3|659.7;73|3.5|764.2;73|4|867.1;73|5|1068.1;73|5.5|1166.3;76|3|688;76|3.5|797.2;76|4|904.8;76|5|1115.3;76|5.5|1218.2;76|6|1319.5;83|3|754;83|3.5|874.1;83|4|992.7;83|5|1225.2;83|5.5|1339.1;83|6|1451.4;89|3|810.5;89|3.5|940.1;89|4|1068.1;89|5|1319.5;89|5.5|1442.8;89|6|1564.5;95|3|867.1;95|3.5|1006.1;95|4|1143.5;95|5|1413.7;95|5.5|1546.4;95|6|1677.6;102|3|933.1;102|3.5|1083.1;102|4|1231.5;102|5|1523.7;102|5.5|1667.4;102|6|1809.6;108|3|989.6;108|3.5|1149;108|4|1306.9;108|5|1617.9;108|5.5|1771.1;108|6|1922.7;108|7|2221.1;108|8|2513.3;114|3|1046.2;114|3.5|1215;114|4|1382.3;114|5|1712.2;114|5.5|1874.7;114|6|2035.8;114|7|2353.1;114|8|2664.1;114|9|2968.8;114|10|3267.3;121|3|1112.1;121|3.5|1292;121|4|1470.3;121|5|1822.1;121|5.5|1995.7;121|6|2167.7;121|7|2507;121|8|2840;121|9|3166.7;121|10|3487.2;127|4|1545.7;127|5|1916.4;127|5.5|2099.4;127|6|2280.8;127|7|2638.9;127|8|2990.8;127|9|3336.4;127|10|3675.7;133|4|1621.1;133|5|2010.6;133|5.5|2203;133|6|2393.9;133|7|2770.9;133|8|3141.6;133|9|3506;133|10|3864.2;140|4|1709;140|5|2120.6;140|5.5|2324;140|6|2525.8;140|7|2924.8;140|8|3317.5;140|9|3703.9;140|10|4084.1;140|11|4457.9;140|12|4825.5;140|13|5186.8;146|4|1784.4;146|5|2214.8;146|5.5|2427.7;146|6|2638.9;146|7|3056.8;146|8|3468.3;146|9|3873.6;146|10|4272.6;146|11|4665.3;146|12|5051.7;146|13|5431.8;152|4|1859.8;152|5|2309.1;152|5.5|2531.3;152|6|2752;152|7|3188.7;152|8|3619.1;152|9|4043.2;152|10|4461.1;152|11|4872.6;152|12|5277.9;152|13|5676.9;159|4|1947.8;159|5|2419;159|5.5|2652.3;159|6|2884;159|7|3342.7;159|8|3795;159|9|4241.1;159|10|4681;159|11|5114.5;159|12|5541.8;159|13|5962.7;168|4|2060.9;168|5|2560.4;168|5.5|2807.8;168|6|3053.6;168|7|3540.6;168|8|4021.2;168|9|4495.6;168|10|4963.7;168|11|5425.5;168|12|5881.1;168|13|6330.3;178|4|2186.5;178|5|2717.5;178|5.5|2980.6;178|6|3242.1;178|7|3760.5;178|8|4272.6;178|9|4778.4;178|10|5277.9;178|11|5771.1;193.7|4|2383.8;193.7|5|2964.1;193.7|5.5|3251.9;193.7|6|3538.1;193.7|7|4105.7;193.7|8|4667.1;193.7|9|5222.3;193.7|10|5771.1;193.7|11|6313.7;193.7|12|6849.9;193.7|13|7379.9;219|4|2701.8;219|5|3361.5;219|5.5|3689;219|6|4015;219|7|4662.1;219|8|5303;219|9|5937.6;219|10|6565.9;219|11|7188;219|12|7803.7;219|13|8413.2;219|14|9016.4;219|15|9613.3;219|16|10203.9;245|4|3028.5;245|5|3769.9;245|5.5|4138.3;245|6|4505;245|7|5233.9;245|8|5956.5;245|9|6672.7;245|10|7382.7;245|11|8086.5;245|12|8783.9;245|13|9475;245|14|10159.9;245|15|10838.5;245|16|11510.8;273|4|3380.4;273|5|4209.7;273|5.5|4622.1;273|6|5032.8;273|7|5849.6;273|8|6660.2;273|9|7464.4;273|10|8262.4;273|11|9054.1;273|12|9839.5;273|13|10618.6;273|14|11391.4;273|15|12158;273|16|12918.2;273|17|13672.2;273|18|14419.9;273|19|15161.3;325|5|5026.5;325|5.5|5520.6;325|6|6013;325|7|6993.2;325|8|7967.1;325|9|8934.7;325|10|9896;325|11|10851.1;325|12|11799.8;325|13|12742.3;325|14|13678.5;325|15|14608.4;325|16|15532;325|17|16449.4;325|18|17360.4;325|19|18265.2;325|20|19163.7;325|21|20055.9;325|22|20941.9;356|6|6597.3;356|7|7674.9;356|8|8746.2;356|9|9811.2;356|10|10869.9;356|11|11922.3;356|12|12968.5;356|13|14008.4;356|14|15041.9;356|15|16069.2;356|16|17090.3;356|17|18105;356|18|19113.4;356|19|20115.6;356|20|21111.5;356|21|22101.1;356|22|23084.4;377|9|10405;377|10|11529.6;377|11|12648.1;377|12|13760.2;377|13|14866;377|14|15965.6;377|15|17058.8;377|16|18145.8;377|17|19226.5;377|18|20301;377|19|21369.1;377|20|22431;377|21|23486.5;377|22|24535.8;406.4|6|7547.4;406.4|7|8783.3;406.4|8|10012.9;406.4|9|11236.2;406.4|10|12453.3;406.4|11|13664;406.4|12|14868.5;406.4|13|16066.7;406.4|14|17258.7;406.4|15|18444.3;406.4|16|19623.6;406.4|17|20796.7;406.4|18|21963.5;406.4|19|23124;406.4|20|24278.2;406.4|21|25426.2;406.4|22|26567.8;426|7|9214.3;426|8|10505.5;426|9|11790.4;426|10|13069;426|11|14341.4;426|12|15607.4;426|13|16867.2;426|14|18120.7;426|15|19367.9;426|16|20608.8;426|17|21843.5;426|18|23071.9;426|19|24293.9;426|20|25509.7;426|21|26719.2;426|22|27922.5;457|7|9896;457|8|11284.6;457|9|12666.9;457|10|14042.9;457|11|15412.7;457|12|16776.1;457|13|18133.3;457|14|19484.2;457|15|20828.8;457|16|22167.1;457|17|23499.1;457|18|24824.9;457|19|26144.3;457|20|27457.5;457|21|28764.4;457|22|30065;508|8|12566.4;508|9|14108.9;508|10|15645.1;508|11|17175.1;508|12|18698.8;508|13|20216.1;508|14|21727.3;508|15|23232.1;508|16|24730.6;508|17|26222.9;508|18|27708.8;508|19|29188.5;508|20|30661.9;508|21|32129.1;508|22|33589.9;508|23|35044.5;508|24|36492.7;508|25|37934.7;508|26|39370.4;508|27|40799.9;508|28|42223;508|29|43639.9;508|30|45050.4;508|31|46454.7;530|8|13119.3;530|9|14730.9;530|10|16336.3;530|11|17935.4;530|12|19528.1;530|13|21114.6;530|14|22694.9;530|15|24268.8;530|16|25836.5;530|17|27397.8;530|18|28952.9;530|19|30501.7;530|20|32044.2;530|21|33580.5;530|22|35110.4;530|23|36634.1;530|24|38151.5;530|25|39662.6;530|26|41167.4;530|27|42666;530|28|44158.2;530|29|45644.2;530|30|47123.9;530|31|48597.3;630|9|17558.4;630|10|19477.9;630|11|21391.1;630|12|23298.1;630|13|25198.7;630|14|27093.1;630|15|28981.2;630|16|30863;630|17|32738.5;630|18|34607.8;630|19|36470.8;630|20|38327.4;630|21|40177.8;630|22|42021.9;630|23|43859.8;630|24|45691.3;630|25|47516.6;630|26|49335.6;630|27|51148.3;630|28|52954.7;630|29|54754.8;630|30|56548.7;630|31|58336.2;630|32|60117.5;630|33|61892.5;630|34|63661.2;630|35|65423.7;630|36|67179.8;630|37|68929.7;720|11|24501.3;720|12|26691;720|13|28874.4;720|14|31051.5;720|15|33222.3;720|16|35386.9;720|17|37545.2;720|18|39697.2;720|19|41842.9;720|20|43982.3;720|21|46115.4;720|22|48242.3;720|23|50362.9;720|24|52477.2;720|25|54585.2;720|26|56686.9;720|27|58782.3;720|28|60871.5;720|29|62954.4;720|30|65031;720|31|67101.3;720|32|69165.3;720|33|71223;720|34|73274.5;720|35|75319.7;720|36|77358.6;720|37|79391.2;720|38|81417.5;720|39|83437.6;720|40|85451.3;720|41|87458.8;820|12|30460.9;820|13|32958.4;820|14|35449.7;820|15|37934.7;820|16|40413.4;820|17|42885.9;820|18|45352;820|19|47811.9;820|20|50265.5;820|21|52712.8;820|22|55153.8;820|23|57588.5;820|24|60017;820|25|62439.2;820|26|64855;820|27|67264.6;820|28|69668;820|29|72065;820|30|74455.7;820|31|76840.2;820|32|79218.4;820|33|81590.3;820|34|83955.9;820|35|86315.3;1220|26|97527.6;1220|27|101193.8;1220|28|104853.8;1220|29|108507.5;1220|30|112154.9;1220|31|115796;1220|32|119430.8;1220|33|123059.3;1220|34|126681.6;1220|35|130297.6;1220|36|133907.2;1220|37|137510.7;1220|38|141107.8;1220|39|144698.6;1220|40|148283.2;1220|41|151861.5;1220|42|155433.4;1220|43|158999.1;1220|44|162558.6;1220|45|166111.7;1220|46|169658.6;1220|47|173199.1;1220|48|176733.4;1420|21|92296.9;1420|22|96622.8;1420|23|100942.5;1420|24|105255.9;1420|25|109563;1420|26|113863.9;1420|27|118158.4;1420|28|122446.7;1420|29|126728.7;1420|30|131004.4;1420|31|135273.8;1420|32|139537;1420|33|143793.8;1420|34|148044.4;1420|35|152288.7;1420|36|156526.7;1420|37|160758.4;1420|38|164983.9;1420|39|169203;1420|40|173415.9;1420|41|177622.5;1420|42|181822.8;1420|43|186016.8;1420|44|190204.6;1420|45|194386;1420|46|198561.2;1420|47|202730.1;1420|48|206892.7",
  "RU-O10": "57|3|508.9;57|3.2|540.9;57|3.5|588.3;60|3|537.2;60|3.2|571;60|3.5|621.2;60|3.8|670.9;70|3|631.5;70|3.2|671.5;70|3.5|731.2;70|3.8|790.3;70|4|829.4;73|3|659.7;73|3.2|701.7;73|3.5|764.2;73|3.8|826.1;73|4|867.1;76|3|688;76|3.2|731.9;76|3.5|797.2;76|3.8|861.9;76|4|904.8;76|4.5|1010.8;76|5|1115.3;76|5.5|1218.2;83|3|754;83|3.2|802.2;83|3.5|874.1;83|3.8|945.5;83|4|992.7;83|4.5|1109.8;83|5|1225.2;83|5.5|1339.1;89|3|810.5;89|3.2|862.6;89|3.5|940.1;89|3.8|1017.1;89|4|1068.1;89|4.5|1194.6;89|5|1319.5;89|5.5|1442.8;95|3.2|922.9;95|5|1413.7;102|3|933.1;102|3.2|993.2;102|3.5|1083.1;102|3.8|1172.3;102|4|1231.5;102|4.5|1378.4;102|5|1523.7;102|5.5|1667.4;108|3|989.6;108|3.2|1053.6;108|3.5|1149;108|3.8|1243.9;108|4|1306.9;108|4.5|1463.2;108|5|1617.9;108|5.5|1771.1;114|3|1046.2;114|3.2|1113.9;114|3.5|1215;114|3.8|1315.6;114|4|1382.3;114|4.5|1548;114|5|1712.2;114|5.5|1874.7;127|3|1168.7;127|3.2|1244.6;127|3.5|1358;127|3.8|1470.8;127|4|1545.7;127|4.5|1731.8;127|5|1916.4;127|5.5|2099.4;133|3|1225.2;133|3.2|1304.9;133|3.5|1423.9;133|3.8|1542.4;133|4|1621.1;133|4.5|1816.6;133|5|2010.6;133|5.5|2203;140|3|1291.2;140|3.2|1375.3;140|3.5|1500.9;140|3.8|1626;140|4|1709;140|4.5|1915.6;140|5|2120.6;140|5.5|2324;152|3|1404.3;152|3.2|1495.9;152|3.5|1632.8;152|3.8|1769.2;152|4|1859.8;152|4.5|2085.2;152|5|2309.1;152|5.5|2531.3;159|3|1470.3;159|3.2|1566.3;159|3.5|1709.8;159|3.8|1852.8;159|4|1947.8;159|4.5|2184.2;159|5|2419;159|5.5|2652.3;159|6|2884;159|7|3342.7;159|8|3795;168|3|1555.1;168|3.2|1656.8;168|3.5|1808.8;168|3.8|1960.2;168|4|2060.9;168|4.5|2311.4;168|5|2560.4;168|5.5|2807.8;168|6|3053.6;168|7|3540.6;168|8|4021.2;177.8|3|1647.5;177.8|3.2|1755.3;177.8|3.5|1916.5;177.8|3.8|2077.2;177.8|4|2184;177.8|4.5|2450;177.8|5|2714.3;177.8|5.5|2977.1;177.8|6|3238.4;177.8|7|3756.1;177.8|8|4267.5;180|4|2211.7;180|5|2748.9;193.7|3|1797.3;193.7|3.2|1915.1;193.7|3.5|2091.4;193.7|3.8|2267;193.7|4|2383.8;193.7|4.5|2674.8;193.7|5|2964.1;193.7|5.5|3251.9;193.7|6|3538.1;193.7|7|4105.7;193.7|8|4667.1;219|3|2035.8;219|3.2|2169.5;219|3.5|2369.5;219|3.8|2569.1;219|4|2701.8;219|4.5|3032.4;219|5|3361.5;219|5.5|3689;219|6|4015;219|7|4662.1;219|8|5303;219|9|5937.6;219|10|6565.9;219|11|7188;219|12|7803.7;219|13|8413.2;219|14|9016.4;219|16|10203.9;219|17|10788.2;219|17.5|11078;219|18|11366.3;219|19|11938.1;219|20|12503.5;244.5|3|2276.1;244.5|3.2|2425.8;244.5|3.5|2649.9;244.5|3.8|2873.5;244.5|4|3022.2;244.5|4.5|3392.9;244.5|5|3762.1;244.5|5.5|4129.6;244.5|6|4495.6;244.5|7|5222.9;244.5|8|5943.9;244.5|9|6658.6;273|3|2544.7;273|3.2|2712.3;273|3.5|2963.3;273|3.8|3213.7;273|4|3380.4;273|4.5|3795.8;273|5|4209.7;273|5.5|4622.1;273|6|5032.8;273|7|5849.6;273|8|6660.2;273|9|7464.4;273|10|8262.4;273|11|9054.1;273|12|9839.5;273|13|10618.6;273|14|11391.4;273|16|12918.2;273|17|13672.2;273|17.5|14046.8;273|18|14419.9;273|19|15161.3;273|20|15896.5;273|21|16625.3;273|22|17347.9;325|3.5|3535.1;325|3.8|3834.5;325|4|4033.8;325|4.5|4531;325|5|5026.5;325|5.5|5520.6;325|6|6013;325|7|6993.2;325|8|7967.1;325|9|8934.7;325|10|9896;325|11|10851.1;325|12|11799.8;325|13|12742.3;325|14|13678.5;325|16|15532;325|17|16449.4;325|17.5|16905.7;325|18|17360.4;325|19|18265.2;325|20|19163.7;325|21|20055.9;325|22|20941.9;355.6|4|4418.3;355.6|4.5|4963.6;355.6|5|5507.2;355.6|5.5|6049.3;355.6|6|6589.8;355.6|7|7666.1;355.6|8|8736.1;355.6|9|9799.9;355.6|10|10857.3;355.6|11|11908.5;355.6|12|12953.4;355.6|13|13992;355.6|14|15024.4;355.6|16|17070.2;355.6|17|18083.6;355.6|17.5|18588;355.6|18|19090.8;355.6|19|20091.7;355.6|20|21086.4;355.6|21|22074.7;355.6|22|23056.8;377|4|4687.3;377|4.5|5266.1;377|5|5843.4;377|5.5|6419.1;377|6|6993.2;377|7|8136.7;377|8|9274;377|9|10405;377|10|11529.6;377|11|12648.1;377|12|13760.2;377|13|14866;377|14|15965.6;377|16|18145.8;377|17|19226.5;377|17.5|19764.5;377|18|20301;377|19|21369.1;377|20|22431;377|21|23486.5;377|22|24535.8;406.4|4|5056.7;406.4|4.5|5681.7;406.4|5|6305.2;406.4|5.5|6927.1;406.4|6|7547.4;406.4|7|8783.3;406.4|8|10012.9;406.4|9|11236.2;406.4|10|12453.3;406.4|11|13664;406.4|12|14868.5;406.4|13|16066.7;406.4|14|17258.7;406.4|16|19623.6;406.4|17|20796.7;406.4|17.5|21380.9;406.4|18|21963.5;406.4|19|23124;406.4|20|24278.2;406.4|21|25426.2;406.4|22|26567.8;426|4|5303;426|4.5|5958.8;426|5|6613.1;426|5.5|7265.7;426|6|7916.8;426|7|9214.3;426|8|10505.5;426|9|11790.4;426|10|13069;426|11|14341.4;426|12|15607.4;426|13|16867.2;426|14|18120.7;426|16|20608.8;426|17|21843.5;426|17.5|22458.5;426|18|23071.9;426|19|24293.9;426|20|25509.7;426|21|26719.2;426|22|27922.5;530|5|8246.7;530|5.5|9062.7;530|6|9877.2;530|7|11501.4;530|8|13119.3;530|9|14730.9;530|10|16336.3;530|11|17935.4;530|12|19528.1;530|13|21114.6;530|14|22694.9;530|16|25836.5;530|17|27397.8;530|17.5|28176.2;530|18|28952.9;530|19|30501.7;530|20|32044.2;530|21|33580.5;530|22|35110.4;530|23|36634.1;530|24|38151.5;630|7|13700.5;630|8|15632.6;630|9|17558.4;630|10|19477.9;630|11|21391.1;630|12|23298.1;630|13|25198.7;630|14|27093.1;630|16|30863;630|17|32738.5;630|17.5|33673.9;630|18|34607.8;630|19|36470.8;630|20|38327.4;630|21|40177.8;630|22|42021.9",
  "CN-HW": "HW100×100×6×8|100|100|6|8|2159;HW125×125×6.5×9|125|125|6.5|9|3000;HW150×150×7×10|150|150|7|10|3965;HW175×175×7.5×11|175|175|7.5|11|5143;HW200×200×8×12|200|200|8|12|6353;HW200×200×12×12|200|204|12|12|7153;HW250×250×11×11|244|252|11|11|8131;HW250×250×9×14|250|250|9|14|9143;HW250×250×14×14|250|255|14|14|10393;HW300×300×12×12|294|302|12|12|10633;HW300×300×10×15|300|300|10|15|11845;HW300×300×15×15|300|305|15|15|13345;HW350×350×13×13|338|351|13|13|13327;HW350×350×10×16|344|348|10|16|14401;HW350×350×16×16|344|354|16|16|16465;HW350×350×12×19|350|350|12|19|17189;HW350×350×19×19|350|357|19|19|19639;HW400×400×15×15|388|402|15|15|17845;HW400×400×11×18|394|398|11|18|18681;HW400×400×18×18|394|405|18|18|21439;HW400×400×13×21|400|400|13|21|21869;HW400×400×21×21|400|408|21|21|25069;HW400×400×18×28|414|405|18|28|29539;HW400×400×20×35|428|407|20|35|36065;HW400×400×30×50|458|417|30|50|52855;HW400×400×45×70|498|432|45|70|77005;HW500×500×15×20|492|465|15|20|25795;HW500×500×15×25|502|465|15|25|30445;HW500×500×20×25|502|470|20|25|32955",
  "CN-HM": "HM150×100×6×9|148|100|6|9|2635;HM200×150×6×9|194|150|6|9|3811;HM250×175×7×11|244|175|7|11|5549;HM300×200×8×12|294|200|8|12|7105;HM350×250×9×14|340|250|9|14|9953;HM400×300×10×16|390|300|10|16|13325;HM450×300×11×18|440|300|11|18|15389;HM500×300×11×15|482|300|11|15|14117;HM500×300×11×18|488|300|11|18|15917;HM550×300×11×15|544|300|11|15|14799;HM550×300×11×18|550|300|11|18|16599;HM600×300×12×17|582|300|12|17|16921;HM600×300×12×20|588|300|12|20|18721;HM600×300×14×23|594|302|14|23|21709",
  "CN-HN": "HN100×50×5×7|100|50|5|7|1185;HN125×60×6×8|125|60|6|8|1669;HN150×75×5×7|150|75|5|7|1785;HN175×90×5×8|175|90|5|8|2290;HN200×100×4.5×7|198|99|4.5|7|2269;HN200×100×5.5×8|200|100|5.5|8|2667;HN250×125×5×8|248|124|5|8|3199;HN250×125×6×9|250|125|6|9|3697;HN300×150×5.5×8|298|149|5.5|8|4080;HN300×150×6.5×9|300|150|6.5|9|4678;HN350×175×6×9|346|174|6|9|5245;HN350×175×7×11|350|175|7|11|6291;HN400×150×8×13|400|150|8|13|7037;HN400×200×7×11|396|199|7|11|7141;HN400×200×8×13|400|200|8|13|8337;HN450×200×8×12|446|199|8|12|8297;HN450×200×9×14|450|200|9|14|9543;HN500×200×9×14|496|199|9|14|9929;HN500×200×10×16|500|200|10|16|11225;HN500×200×11×19|506|201|11|19|12931;HN550×200×9×14|546|199|9|14|10379;HN550×200×10×16|550|200|10|16|14925;HN600×200×10×15|596|199|10|15|11775;HN600×200×11×17|600|200|11|17|13171;HN600×200×12×20|606|201|12|20|14977;HN650×300×10×15|646|299|10|15|15275;HN650×300×11×17|650|300|11|17|17121;HN650×300×12×20|656|301|12|20|19577;HN700×300×13×20|692|300|13|20|20754;HN700×300×13×24|700|300|13|24|23154;HN750×300×12×16|734|299|12|16|18270;HN750×300×13×20|742|300|13|20|21404;HN750×300×13×24|750|300|13|24|23804;HN750×300×16×28|758|303|16|28|28478;HN800×300×14×22|792|300|14|22|23950;HN800×300×14×26|800|300|14|26|26350;HN850×300×14×19|834|298|14|19|22746;HN850×300×15×23|842|299|15|23|25972;HN850×300×16×27|850|300|16|27|29214;HN850×300×17×31|858|301|17|31|32472;HN900×300×15×23|890|299|15|23|26692;HN900×300×16×28|900|300|16|28|30582;HN900×300×18×34|912|302|18|34|36006;HN1000×300×16×21|970|297|16|21|27600;HN1000×300×17×26|980|298|17|26|31550;HN1000×300×17×31|990|298|17|31|34530;HN1000×300×19×36|1000|300|19|36|39510;HN1000×300×21×40|1008|302|21|40|43926;HN100×75×6×8|100|75|6|8|1790;HN126×75×6×8|126|75|6|8|1946;HN140×90×5×8|140|90|5|8|2146;HN160×90×5×8|160|90|5|8|2246;HN180×90×5×8|180|90|5|8|2346;HN220×125×6×9|220|125|6|9|3607;HN280×125×6×9|280|125|6|9|3967;HN320×150×6.5×9|320|150|6.5|9|4883;HN360×150×7×11|360|150|7|11|5886;HN560×175×11×17|560|175|11|17|12230;HN630×200×13×20|630|200|13|20|16340",
  "CN-HT": "HT100×50×3.2×4.5|95|48|3.2|4.5|762;HT100×40×4×5.5|97|49|4|5.5|938;HT100×100×4.5×6|96|99|4.5|6|1621;HT125×60×3.2×4.5|118|58|3.2|4.5|926;HT125×60×4×5.5|120|59|4|5.5|1140;HT125×125×4.5×6|119|123|4.5|6|2012;HT150×75×3.2×4.5|145|73|3.2|4.5|1147;HT150×75×4×5.5|147|74|4|5.5|1413;HT150×100×3.2×4.5|139|97|3.2|4.5|1344;HT150×100×4.5×6|142|99|4.5|6|1828;HT150×150×5×7|144|148|5|7|2777;HT150×150×6×8.5|147|149|6|8.5|3368;HT175×90×3.2×4.5|168|88|3.2|4.5|1356;HT175×90×4×6|171|89|4|6|1759;HT175×175×5×7|167|173|5|7|3332;HT175×175×6.5×9.5|172|175|6.5|9.5|4465;HT200×100×3.2×4.5|193|98|3.2|4.5|1526;HT200×100×4×6|196|99|4|6|1979;HT200×150×4.5×6|188|149|4.5|6|2635;HT200×200×6×8|192|198|6|8|4369;HT250×125×4.5×6|244|124|4.5|6|2587;HT250×175×4.5×8|238|173|4.5|8|3912;HT300×150×4.5×6|294|148|4.5|6|3190;HT300×200×6×8|286|198|6|8|4933;HT350×175×4.5×6|340|173|4.5|6|3697;HT400×150×6×8|390|148|6|8|4757;HT400×200×6×8|390|198|6|8|5557",
  "CN-YB-H": "YB-H300×200×6×10|300|200|6|10|5680;YB-H300×200×6×12|300|200|6|12|6460;YB-H300×200×8×14|300|200|8|14|7730;YB-H300×250×8×12|300|250|8|12|8210;YB-H300×250×10×14|300|250|10|14|9720;YB-H300×300×8×12|300|300|8|12|9410;YB-H300×300×10×16|300|300|10|16|12300;YB-H300×300×12×20|300|300|12|20|15100;YB-H350×175×6×10|350|175|6|10|5480;YB-H350×175×8×12|350|175|8|12|6810;YB-H350×200×6×8|350|200|6|8|5200;YB-H350×200×8×10|350|200|8|10|6640;YB-H350×200×8×12|350|200|8|12|7410;YB-H350×200×10×16|350|200|10|16|9580;YB-H350×250×8×10|350|250|8|10|7640;YB-H350×250×8×12|350|250|8|12|8610;YB-H350×250×10×16|350|250|10|16|11200;YB-H350×300×8×12|350|300|8|12|9810;YB-H350×300×10×16|350|300|10|16|12800;YB-H350×350×8×12|350|350|8|12|11000;YB-H350×350×10×16|350|350|10|16|14400;YB-H350×350×12×20|350|350|12|20|17700;YB-H400×200×6×10|400|200|6|10|6280;YB-H400×200×8×12|400|200|8|12|7810;YB-H400×200×8×16|400|200|8|16|9340;YB-H400×200×10×20|400|200|10|20|11600;YB-H400×250×6×10|400|250|6|10|7280;YB-H400×250×8×12|400|250|8|12|9010;YB-H400×250×8×16|400|250|8|16|10900;YB-H400×250×10×20|400|250|10|20|13600;YB-H400×300×8×12|400|300|8|12|10200;YB-H400×300×10×16|400|300|10|16|13300;YB-H400×300×12×20|400|300|12|20|16300;YB-H400×400×8×14|400|400|8|14|14200;YB-H400×400×10×16|400|400|10|16|16500;YB-H400×400×12×20|400|400|12|20|20300;YB-H400×400×16×25|400|400|16|25|25600;YB-H400×400×20×32|400|400|20|32|32300;YB-H400×400×25×40|400|400|25|40|40000;YB-H392×400×10×16|392|400|10|16|16400;YB-H410×400×16×25|410|400|16|25|25800;YB-H424×400×20×32|424|400|20|32|32800;YB-H440×400×25×40|440|400|25|40|41000;YB-H450×250×8×12|450|250|8|12|9410;YB-H450×250×10×16|450|250|10|16|12200;YB-H450×250×10×20|450|250|10|20|14100;YB-H450×300×8×12|450|300|8|12|10600;YB-H450×300×10×16|450|300|10|16|13800;YB-H450×300×12×20|450|300|12|20|16900;YB-H450×300×12×25|450|300|12|25|19800;YB-H450×400×10×16|450|400|10|16|17000;YB-H450×400×10×20|450|400|10|20|20100;YB-H450×400×12×25|450|400|12|25|24800;YB-H500×250×8×16|500|250|8|16|11700;YB-H500×250×10×20|500|250|10|20|14600;YB-H500×250×12×25|500|250|12|25|17900;YB-H500×300×8×16|500|300|8|16|13300;YB-H500×300×10×20|500|300|10|20|16600;YB-H500×300×12×25|500|300|12|25|20400;YB-H500×400×10×16|500|400|10|16|17500;YB-H500×400×10×20|500|400|10|20|20600;YB-H500×400×12×25|500|400|12|25|25400;YB-H600×300×10×16|600|300|10|16|15300;YB-H600×300×10×20|600|300|10|20|17600;YB-H600×300×12×25|600|300|12|25|21600;YB-H600×400×10×16|600|400|10|16|18500;YB-H600×400×10×20|600|400|10|20|21600;YB-H600×400×12×30|600|400|12|30|30500;YB-H700×300×10×20|700|300|10|20|18600;YB-H700×300×10×25|700|300|10|25|21500;YB-H700×300×12×30|700|300|12|30|25700;YB-H700×350×10×20|700|350|10|20|20600;YB-H700×350×10×25|700|350|10|25|24000;YB-H700×350×12×30|700|350|12|30|28700;YB-H700×400×10×20|700|400|10|20|22600;YB-H700×400×10×25|700|400|10|25|26500;YB-H700×400×12×30|700|400|12|30|31700;YB-H800×300×12×20|800|300|12|20|21100;YB-H800×300×12×25|800|300|12|25|24000;YB-H800×300×12×30|800|300|12|30|26900;YB-H800×350×12×20|800|350|12|20|23100;YB-H800×350×12×25|800|350|12|25|26500;YB-H800×350×12×30|800|350|12|30|29900;YB-H800×400×12×20|800|400|12|20|25100;YB-H800×400×12×25|800|400|12|25|29000;YB-H800×400×12×28|800|400|12|28|31300;YB-H800×400×14×32|800|400|14|32|35900;YB-H900×350×14×20|900|350|14|20|26000;YB-H900×350×14×25|900|350|14|25|29400;YB-H900×350×14×28|900|350|14|28|31400;YB-H900×350×14×23|900|350|14|23|34100;YB-H900×400×14×20|900|400|14|20|28000;YB-H900×400×14×25|900|400|14|25|31900;YB-H900×400×14×30|900|400|14|30|35800;YB-H900×400×14×36|900|400|14|36|40400;YB-H900×450×14×20|900|450|14|20|30000;YB-H900×450×14×25|900|450|14|25|34400;YB-H900×450×14×30|900|450|14|30|38800;YB-H900×450×16×36|900|450|16|36|45700;YB-H1000×400×14×20|1000|400|14|20|29400;YB-H1000×400×14×25|1000|400|14|25|33300;YB-H1000×400×14×30|1000|400|14|30|37200;YB-H1000×400×16×36|1000|400|16|36|43700;YB-H1000×450×14×20|1000|450|14|20|31400;YB-H1000×450×14×25|1000|450|14|25|35800;YB-H1000×450×14×30|1000|450|14|30|40200;YB-H1000×450×14×36|1000|450|14|36|45400;YB-H1000×500×14×20|1000|500|14|20|33400;YB-H1000×500×14×25|1000|500|14|25|38300;YB-H1000×500×16×30|1000|500|16|30|45000;YB-H1000×500×16×36|1000|500|16|36|50900;YB-H1100×400×16×25|1100|400|16|25|36800;YB-H1100×400×16×30|1100|400|16|30|40600;YB-H1100×400×20×36|1100|400|20|36|49400;YB-H1100×500×16×30|1100|500|16|30|46600;YB-H1100×500×20×36|1100|500|20|36|56600;YB-H1200×400×18×25|1200|400|18|25|40700;YB-H1200×400×18×30|1200|400|18|30|44500;YB-H1200×400×18×36|1200|400|18|36|49100;YB-H1200×400×20×40|1200|400|20|40|54400;YB-H1200×450×18×30|1200|450|18|30|47500;YB-H1200×450×20×36|1200|450|20|36|55000;YB-H1200×500×20×30|1200|500|20|30|52300;YB-H1200×500×20×36|1200|500|20|36|58600;YB-H1200×500×20×40|1200|500|20|40|62400;YB-H1200×600×14×25|1200|600|14|25|46100;YB-H1200×600×14×30|1200|600|14|30|52000;YB-H1200×600×16×36|1200|600|16|36|61300",
  "CN-YB-LWH": "YB-LWH100×50×3.5×5|100|50|3.5|5|815;YB-LWH102×75×4×6|102|75|4|6|1260;YB-LWH102×100×4×6|102|100|4|6|1560;YB-LWH125×75×4×6|125|75|4|6|1350;YB-LWH125×125×4×6|125|125|4|6|1950;YB-LWH150×75×5×7|150|75|5|7|1730;YB-LWH150×100×5×7|150|100|5|7|2080;YB-LWH152×150×6×8|152|150|6|8|3230;YB-LWH200×100×5×8|200|100|5|8|2520;YB-LWH200×150×5×8|200|150|5|8|3320;YB-LWH202×200×6×9|202|200|6|9|4700;YB-LWH250×125×5×8|250|125|5|8|3170;YB-LWH250×150×5×8|250|150|5|8|3570;YB-LWH250×200×5×8|250|200|5|8|4370;YB-LWH252×250×6×9|252|250|6|9|5900;YB-LWH300×150×5×8|300|150|5|8|3820;YB-LWH300×200×5×8|300|200|5|8|4620;YB-LWH302×250×6×9|302|250|6|9|6200;YB-LWH304×300×7×10|304|300|7|10|7990;YB-LWH350×200×7×10|350|200|7|10|6310;YB-LWH350×250×7×10|350|250|7|10|7310;YB-LWH350×300×7×10|350|300|7|10|8310;YB-LWH400×200×7×10|400|200|7|10|6660;YB-LWH400×250×7×10|400|250|7|10|7660;YB-LWH404×300×8×12|404|300|8|12|10200;YB-LWH450×200×7×10|450|200|7|10|7010;YB-LWH454×250×8×12|454|250|8|12|9440;YB-LWH454×300×9×12|454|300|9|12|11400",
  "CN-YB-WH": "YB-WH100×50×3.2×4.5|100|50|3.2|4.5|741;YB-WH100×50×4×5|100|50|4|5|860;YB-WH100×75×4×6|100|75|4|6|1250;YB-WH100×100×4×6|100|100|4|6|1550;YB-WH100×100×6×8|100|100|6|8|2100;YB-WH125×75×4×6|125|75|4|6|1350;YB-WH125×125×4×6|125|125|4|6|1950;YB-WH150×75×3.2×4.5|150|75|3.2|4.5|1120;YB-WH150×75×4×6|150|75|4|6|1450;YB-WH150×75×5×8|150|75|5|8|1870;YB-WH150×100×3.2×4.5|150|100|3.2|4.5|1350;YB-WH150×100×4×6|150|100|4|6|1750;YB-WH150×100×5×8|150|100|5|8|2270;YB-WH150×150×4×6|150|150|4|6|2350;YB-WH150×150×5×8|150|150|5|8|3070;YB-WH150×150×6×8|150|150|6|8|3200;YB-WH200×100×3.2×4.5|200|100|3.2|4.5|1510;YB-WH200×100×4×6|200|100|4|6|1950;YB-WH200×100×5×8|200|100|5|8|2520;YB-WH200×150×4×6|200|150|4|6|2550;YB-WH200×150×5×8|200|150|5|8|3320;YB-WH200×200×5×8|200|200|5|8|4120;YB-WH200×200×6×10|200|200|6|10|5080;YB-WH250×125×4×6|250|125|4|6|2450;YB-WH250×125×5×8|250|125|5|8|3170;YB-WH250×125×6×10|250|125|6|10|3880;YB-WH250×150×4×6|250|150|4|6|2750;YB-WH250×150×5×8|250|150|5|8|3570;YB-WH250×150×6×10|250|150|6|10|4380;YB-WH250×200×5×8|250|200|5|8|4370;YB-WH250×200×5×10|250|200|5|10|5150;YB-WH250×200×6×10|250|200|6|10|5380;YB-WH250×200×6×12|250|200|6|12|6150;YB-WH250×250×6×10|250|250|6|10|6380;YB-WH250×250×6×12|250|250|6|12|7350;YB-WH250×250×8×14|250|250|8|14|8770;YB-WH300×200×6×8|300|200|6|8|4900;YB-WH300×200×6×10|300|200|6|10|5680;YB-WH300×200×6×12|300|200|6|12|6450;YB-WH300×200×8×14|300|200|8|14|7770;YB-WH300×200×10×16|300|200|10|16|9080;YB-WH300×250×6×10|300|250|6|10|6680;YB-WH300×250×6×12|300|250|6|12|7650;YB-WH300×250×8×14|300|250|8|14|9170;YB-WH300×250×10×16|300|250|10|16|10600;YB-WH300×300×6×10|300|300|6|10|7680;YB-WH300×300×8×12|300|300|8|12|9400;YB-WH300×300×8×14|300|300|8|14|10500;YB-WH300×300×10×16|300|300|10|16|12200;YB-WH300×300×10×18|300|300|10|18|13400;YB-WH300×300×12×20|300|300|12|20|15100;YB-WH350×175×4.5×6|350|175|4.5|6|3620;YB-WH350×175×4.5×8|350|175|4.5|8|4300;YB-WH350×175×6×8|350|175|6|8|4800;YB-WH350×175×6×10|350|175|6|10|5480;YB-WH350×175×6×12|350|175|6|12|6150;YB-WH350×175×8×12|350|175|8|12|6800;YB-WH350×175×8×14|350|175|8|14|7470;YB-WH350×175×10×16|350|175|10|16|8780;YB-WH350×200×6×8|350|200|6|8|5200;YB-WH350×200×6×10|350|200|6|10|5980;YB-WH350×200×6×12|350|200|6|12|6750;YB-WH350×200×8×10|350|200|8|10|6640;YB-WH350×200×8×12|350|200|8|12|7400;YB-WH350×200×8×14|350|200|8|14|8170;YB-WH350×200×10×16|350|200|10|16|9580;YB-WH350×250×6×10|350|250|6|10|6980;YB-WH350×250×6×12|350|250|6|12|7950;YB-WH350×250×8×12|350|250|8|12|8600;YB-WH350×250×8×14|350|250|8|14|9570;YB-WH350×250×10×16|350|250|10|16|11100;YB-WH350×300×6×10|350|300|6|10|7980;YB-WH350×300×6×12|350|300|6|12|9150;YB-WH350×300×8×14|350|300|8|14|10900;YB-WH350×300×10×16|350|300|10|16|12700;YB-WH350×300×10×18|350|300|10|18|13900;YB-WH350×350×6×12|350|350|6|12|10300;YB-WH350×350×8×14|350|350|8|14|12300;YB-WH350×350×8×16|350|350|8|16|13700;YB-WH350×350×10×16|350|350|10|16|14300;YB-WH350×350×10×18|350|350|10|18|15700;YB-WH350×350×12×20|350|350|12|20|17700;YB-WH400×200×6×8|400|200|6|8|5500;YB-WH400×200×6×10|400|200|6|10|6280;YB-WH400×200×6×12|400|200|6|12|7050;YB-WH400×200×8×12|400|200|8|12|7800;YB-WH400×200×8×14|400|200|8|14|8570;YB-WH400×200×8×16|400|200|8|16|9340;YB-WH400×200×8×18|400|200|8|18|10100;YB-WH400×200×10×16|400|200|10|16|10000;YB-WH400×200×10×18|400|200|10|18|10800;YB-WH400×200×10×20|400|200|10|20|11600;YB-WH400×250×6×10|400|250|6|10|7280;YB-WH400×250×6×12|400|250|6|12|8250;YB-WH400×250×8×14|400|250|8|14|9970;YB-WH400×250×8×16|400|250|8|16|10900;YB-WH400×250×8×18|400|250|8|18|11900;YB-WH400×250×10×16|400|250|10|16|11600;YB-WH400×250×10×18|400|250|10|18|12600;YB-WH400×250×10×20|400|250|10|20|13600;YB-WH400×300×6×10|400|300|6|10|8280;YB-WH400×300×6×12|400|300|6|12|9450;YB-WH400×300×8×14|400|300|8|14|11300;YB-WH400×300×10×16|400|300|10|16|13200;YB-WH400×300×10×18|400|300|10|18|14400;YB-WH400×300×10×20|400|300|10|20|15600;YB-WH400×300×12×20|400|300|12|20|16300;YB-WH400×400×8×14|400|400|8|14|14100;YB-WH400×400×8×18|400|400|8|18|17300;YB-WH400×400×10×16|400|400|10|16|16400;YB-WH400×400×10×18|400|400|10|18|18000;YB-WH400×400×10×20|400|400|10|20|19600;YB-WH400×400×12×22|400|400|12|22|21800;YB-WH400×400×12×25|400|400|12|25|24200;YB-WH400×400×16×25|400|400|16|25|25600;YB-WH400×400×20×32|400|400|20|32|32300;YB-WH400×400×20×40|400|400|20|40|38400;YB-WH450×250×8×12|450|250|8|12|9400;YB-WH450×250×8×14|450|250|8|14|10300;YB-WH450×250×10×16|450|250|10|16|12100;YB-WH450×250×10×18|450|250|10|18|13100;YB-WH450×250×10×20|450|250|10|20|14100;YB-WH450×250×12×22|450|250|12|22|15800;YB-WH450×250×12×25|450|250|12|25|17300;YB-WH450×300×8×12|450|300|8|12|10600;YB-WH450×300×8×14|450|300|8|14|11700;YB-WH450×300×10×16|450|300|10|16|13700;YB-WH450×300×10×18|450|300|10|18|14900;YB-WH450×300×10×20|450|300|10|20|16100;YB-WH450×300×12×20|450|300|12|20|16900;YB-WH450×300×12×22|450|300|12|22|18000;YB-WH450×300×12×25|450|300|12|25|19800;YB-WH450×400×8×14|450|400|8|14|14500;YB-WH450×400×10×16|450|400|10|16|16900;YB-WH450×400×10×18|450|400|10|18|18500;YB-WH450×400×10×20|450|400|10|20|20100;YB-WH450×400×12×22|450|400|12|22|22400;YB-WH450×400×12×25|450|400|12|25|24800;YB-WH500×250×8×12|500|250|8|12|9800;YB-WH500×250×8×14|500|250|8|14|10700;YB-WH500×250×8×16|500|250|8|16|11700;YB-WH500×250×10×16|500|250|10|16|12600;YB-WH500×250×10×18|500|250|10|18|13600;YB-WH500×250×10×20|500|250|10|20|14600;YB-WH500×250×12×22|500|250|12|22|16400;YB-WH500×250×12×25|500|250|12|25|17900;YB-WH500×300×8×12|500|300|8|12|11000;YB-WH500×300×8×14|500|300|8|14|12100;YB-WH500×300×8×16|500|300|8|16|13300;YB-WH500×300×10×16|500|300|10|16|14200;YB-WH500×300×10×18|500|300|10|18|15400;YB-WH500×300×10×20|500|300|10|20|16600;YB-WH500×300×12×22|500|300|12|22|18600;YB-WH500×300×12×25|500|300|12|25|20400;YB-WH500×400×8×14|500|400|8|14|14900;YB-WH500×400×10×16|500|400|10|16|17400;YB-WH500×400×10×18|500|400|10|18|19000;YB-WH500×400×10×20|500|400|10|20|20600;YB-WH500×400×12×22|500|400|12|22|23000;YB-WH500×400×12×25|500|400|12|25|25400;YB-WH500×500×10×18|500|500|10|18|22600;YB-WH500×500×10×20|500|500|10|20|24600;YB-WH500×500×12×22|500|500|12|22|27400;YB-WH500×500×12×25|500|500|12|25|30400;YB-WH500×500×20×25|500|500|20|25|34000;YB-WH600×300×8×14|600|300|8|14|12900;YB-WH600×300×10×16|600|300|10|16|15200;YB-WH600×300×10×18|600|300|10|18|16400;YB-WH600×300×10×20|600|300|10|20|17600;YB-WH600×300×12×22|600|300|12|22|19800;YB-WH600×300×12×25|600|300|12|25|21600;YB-WH600×400×8×14|600|400|8|14|15700;YB-WH600×400×10×16|600|400|10|16|18400;YB-WH600×400×10×18|600|400|10|18|20000;YB-WH600×400×10×20|600|400|10|20|21600;YB-WH600×400×10×25|600|400|10|25|25500;YB-WH600×400×12×22|600|400|12|22|24200;YB-WH600×400×12×28|600|400|12|28|28900;YB-WH600×400×12×30|600|400|12|30|30400;YB-WH600×400×14×32|600|400|14|32|33100;YB-WH700×300×10×18|700|300|10|18|17400;YB-WH700×300×10×20|700|300|10|20|18600;YB-WH700×300×10×25|700|300|10|25|21500;YB-WH700×300×12×22|700|300|12|22|21000;YB-WH700×300×12×25|700|300|12|25|22800;YB-WH700×300×12×30|700|300|12|30|25600;YB-WH700×300×12×36|700|300|12|36|29100;YB-WH700×300×14×32|700|300|14|32|28100;YB-WH700×300×16×36|700|300|16|36|31600;YB-WH700×350×10×18|700|350|10|18|19200;YB-WH700×350×10×20|700|350|10|20|20600;YB-WH700×350×10×25|700|350|10|25|24000;YB-WH700×350×12×22|700|350|12|22|23200;YB-WH700×350×12×25|700|350|12|25|25300;YB-WH700×350×12×28|700|350|12|28|27300;YB-WH700×350×12×30|700|350|12|30|28600;YB-WH700×350×12×36|700|350|12|36|32700;YB-WH700×350×14×32|700|350|14|32|31300;YB-WH700×350×16×36|700|350|16|36|35200;YB-WH700×400×10×18|700|400|10|18|21000;YB-WH700×400×10×20|700|400|10|20|22600;YB-WH700×400×10×25|700|400|10|25|26500;YB-WH700×400×12×22|700|400|12|22|25400;YB-WH700×400×12×25|700|400|12|25|27800;YB-WH700×400×12×28|700|400|12|28|30100;YB-WH700×400×12×30|700|400|12|30|31600;YB-WH700×400×12×36|700|400|12|36|36300;YB-WH700×400×14×32|700|400|14|32|34500;YB-WH700×400×16×36|700|400|16|36|38800;YB-WH800×300×10×18|800|300|10|18|18400;YB-WH800×300×10×20|800|300|10|20|19600;YB-WH800×300×10×25|800|300|10|25|22500;YB-WH800×300×12×22|800|300|12|22|22200;YB-WH800×300×12×25|800|300|12|25|24000;YB-WH800×300×12×28|800|300|12|28|25700;YB-WH800×300×12×30|800|300|12|30|26800;YB-WH800×300×12×36|800|300|12|36|30300;YB-WH800×300×14×32|800|300|14|32|29500;YB-WH800×300×16×36|800|300|16|36|33200;YB-WH800×350×10×18|800|350|10|18|20200;YB-WH800×350×10×20|800|350|10|20|21600;YB-WH800×350×10×25|800|350|10|25|25000;YB-WH800×350×12×22|800|350|12|22|24400;YB-WH800×350×12×25|800|350|12|25|26500;YB-WH800×350×12×28|800|350|12|28|28500;YB-WH800×350×12×30|800|350|12|30|29800;YB-WH800×350×12×36|800|350|12|36|33900;YB-WH800×350×14×32|800|350|14|32|32700;YB-WH800×350×16×36|800|350|16|36|36800;YB-WH800×400×10×18|800|400|10|18|22000;YB-WH800×400×10×20|800|400|10|20|23600;YB-WH800×400×10×25|800|400|10|25|27500;YB-WH800×400×10×28|800|400|10|28|29800;YB-WH800×400×12×22|800|400|12|22|26600;YB-WH800×400×12×25|800|400|12|25|29000;YB-WH800×400×12×28|800|400|12|28|31300;YB-WH800×400×12×32|800|400|12|32|34400;YB-WH800×400×12×36|800|400|12|36|37500;YB-WH800×400×14×32|800|400|14|32|35900;YB-WH800×400×16×36|800|400|16|36|40400;YB-WH900×350×10×20|900|350|10|20|22600;YB-WH900×350×12×20|900|350|12|20|24300;YB-WH900×350×12×22|900|350|12|22|25600;YB-WH900×350×12×25|900|350|12|25|27700;YB-WH900×350×12×28|900|350|12|28|29700;YB-WH900×350×14×32|900|350|14|32|34100;YB-WH900×350×14×36|900|350|14|36|36700;YB-WH900×350×16×36|900|350|16|36|38400;YB-WH900×400×10×20|900|400|10|20|24600;YB-WH900×400×12×20|900|400|12|20|26300;YB-WH900×400×12×22|900|400|12|22|27800;YB-WH900×400×12×25|900|400|12|25|30200;YB-WH900×400×12×28|900|400|12|28|32500;YB-WH900×400×12×30|900|400|12|30|34000;YB-WH900×400×14×32|900|400|14|32|37300;YB-WH900×400×14×36|900|400|14|36|40300;YB-WH900×400×14×40|900|400|14|40|43400;YB-WH900×400×16×36|900|400|16|36|42000;YB-WH900×400×16×40|900|400|16|40|45100;YB-WH1100×400×12×20|1100|400|12|20|28700;YB-WH1100×400×12×22|1100|400|12|22|30200;YB-WH1100×400×12×25|1100|400|12|25|32600;YB-WH1100×400×12×28|1100|400|12|28|34900;YB-WH1100×400×14×30|1100|400|14|30|38500;YB-WH1100×400×14×32|1100|400|14|32|40100;YB-WH1100×400×14×36|1100|400|14|36|43100;YB-WH1100×400×16×40|1100|400|16|40|48300;YB-WH1100×500×12×20|1100|500|12|20|32700;YB-WH1100×500×12×22|1100|500|12|22|34600;YB-WH1100×500×12×25|1100|500|12|25|37600;YB-WH1100×500×12×28|1100|500|12|28|40500;YB-WH1100×500×14×30|1100|500|14|30|44500;YB-WH1100×500×14×32|1100|500|14|32|46500;YB-WH1100×500×14×36|1100|500|14|36|50300;YB-WH1100×500×16×40|1100|500|16|40|56300;YB-WH1200×400×14×20|1200|400|14|20|32200;YB-WH1200×400×14×22|1200|400|14|22|33700;YB-WH1200×400×14×25|1200|400|14|25|36100;YB-WH1200×400×14×28|1200|400|14|28|38400;YB-WH1200×400×14×30|1200|400|14|30|39900;YB-WH1200×400×14×32|1200|400|14|32|41500;YB-WH1200×400×14×36|1200|400|14|36|44500;YB-WH1200×400×16×40|1200|400|16|40|49900;YB-WH1200×450×14×20|1200|450|14|20|34200;YB-WH1200×450×14×22|1200|450|14|22|35900;YB-WH1200×450×14×25|1200|450|14|25|38600;YB-WH1200×450×14×28|1200|450|14|28|41200;YB-WH1200×450×14×30|1200|450|14|30|42900;YB-WH1200×450×14×32|1200|450|14|32|44700;YB-WH1200×450×14×36|1200|450|14|36|48100;YB-WH1200×450×16×36|1200|450|16|36|50400;YB-WH1200×450×16×40|1200|450|16|40|53900;YB-WH1200×500×14×20|1200|500|14|20|36200;YB-WH1200×500×14×22|1200|500|14|22|38100;YB-WH1200×500×14×25|1200|500|14|25|41100;YB-WH1200×500×14×28|1200|500|14|28|44000;YB-WH1200×500×14×32|1200|500|14|32|47900;YB-WH1200×500×14×36|1200|500|14|36|51700;YB-WH1200×500×16×36|1200|500|16|36|54000;YB-WH1200×500×16×40|1200|500|16|40|57900;YB-WH1200×500×16×45|1200|500|16|45|62700;YB-WH1200×600×14×30|1200|600|14|30|51900;YB-WH1200×600×16×36|1200|600|16|36|61200;YB-WH1200×600×16×40|1200|600|16|40|65900;YB-WH1200×600×16×45|1200|600|16|45|71700;YB-WH1300×450×16×25|1300|450|16|25|42500;YB-WH1300×450×16×30|1300|450|16|30|46800;YB-WH1300×450×16×36|1300|450|16|36|52000;YB-WH1300×450×18×40|1300|450|18|40|57900;YB-WH1300×450×18×45|1300|450|18|45|62200;YB-WH1300×500×16×25|1300|500|16|25|45000;YB-WH1300×500×16×30|1300|500|16|30|49800;YB-WH1300×500×16×36|1300|500|16|36|55600;YB-WH1300×500×18×40|1300|500|18|40|61900;YB-WH1300×500×18×45|1300|500|18|45|66700;YB-WH1300×600×16×30|1300|600|16|30|55800;YB-WH1300×600×16×36|1300|600|16|36|62800;YB-WH1300×600×18×40|1300|600|18|40|69900;YB-WH1300×600×18×45|1300|600|18|45|75700;YB-WH1300×600×20×50|1300|600|20|50|84000;YB-WH1400×450×16×25|1400|450|16|25|44100;YB-WH1400×450×16×30|1400|450|16|30|48400;YB-WH1400×450×18×36|1400|450|18|36|56300;YB-WH1400×450×18×40|1400|450|18|40|59700;YB-WH1400×450×18×45|1400|450|18|45|64000;YB-WH1400×500×16×25|1400|500|16|25|46600;YB-WH1400×500×16×30|1400|500|16|30|51400;YB-WH1400×500×18×36|1400|500|18|36|59900;YB-WH1400×500×18×40|1400|500|18|40|63700;YB-WH1400×500×18×45|1400|500|18|45|68500;YB-WH1400×600×16×30|1400|600|16|30|57400;YB-WH1400×600×16×36|1400|600|16|36|64400;YB-WH1400×600×18×40|1400|600|18|40|71700;YB-WH1400×600×18×45|1400|600|18|45|77500;YB-WH1400×600×18×50|1400|600|18|50|83400;YB-WH1500×500×18×25|1500|500|18|25|51100;YB-WH1500×500×18×30|1500|500|18|30|55900;YB-WH1500×500×18×36|1500|500|18|36|61700;YB-WH1500×500×18×40|1500|500|18|40|65500;YB-WH1500×500×20×45|1500|500|20|45|73200;YB-WH1500×550×18×30|1500|550|18|30|58900;YB-WH1500×550×18×36|1500|550|18|36|65300;YB-WH1500×550×18×40|1500|550|18|40|69500;YB-WH1500×550×20×45|1500|550|20|45|77700;YB-WH1500×600×18×30|1500|600|18|30|61900;YB-WH1500×600×18×36|1500|600|18|36|68900;YB-WH1500×600×18×40|1500|600|18|40|73500;YB-WH1500×600×20×45|1500|600|20|45|82200;YB-WH1500×600×20×50|1500|600|20|50|88000;YB-WH1600×600×18×30|1600|600|18|30|63700;YB-WH1600×600×18×36|1600|600|18|36|70700;YB-WH1600×600×18×40|1600|600|18|40|75300;YB-WH1600×600×20×45|1600|600|20|45|84200;YB-WH1600×600×20×50|1600|600|20|50|90000;YB-WH1600×650×18×30|1600|650|18|30|66700;YB-WH1600×650×18×36|1600|650|18|36|74300;YB-WH1600×650×18×40|1600|650|18|40|79300;YB-WH1600×650×20×45|1600|650|20|45|88700;YB-WH1600×650×20×50|1600|650|20|50|95000;YB-WH1600×700×18×30|1600|700|18|30|69700;YB-WH1600×700×18×36|1600|700|18|36|77900;YB-WH1600×700×18×40|1600|700|18|40|83300;YB-WH1600×700×20×45|1600|700|20|45|93200;YB-WH1600×700×20×50|1600|700|20|50|100000;YB-WH1700×600×18×30|1700|600|18|30|65500;YB-WH1700×600×18×36|1700|600|18|36|72500;YB-WH1700×600×18×40|1700|600|18|40|77100;YB-WH1700×600×20×45|1700|600|20|45|86200;YB-WH1700×600×20×50|1700|600|20|50|92000;YB-WH1700×650×18×30|1700|650|18|30|68500;YB-WH1700×650×18×36|1700|650|18|36|76100;YB-WH1700×650×18×40|1700|650|18|40|81100;YB-WH1700×650×20×45|1700|650|20|45|90700;YB-WH1700×650×20×50|1700|650|20|50|97000;YB-WH1700×700×18×32|1700|700|18|32|74200;YB-WH1700×700×18×36|1700|700|18|36|79700;YB-WH1700×700×18×40|1700|700|18|40|85100;YB-WH1700×700×20×45|1700|700|20|45|95200;YB-WH1700×700×20×50|1700|700|20|50|102000;YB-WH1700×750×18×32|1700|750|18|32|77400;YB-WH1700×750×18×36|1700|750|18|36|83300;YB-WH1700×750×18×40|1700|750|18|40|89100;YB-WH1700×750×20×45|1700|750|20|45|99700;YB-WH1700×750×20×50|1700|750|20|50|107000;YB-WH1800×600×18×30|1800|600|18|30|67300;YB-WH1800×600×18×36|1800|600|18|36|74300;YB-WH1800×600×18×40|1800|600|18|40|78900;YB-WH1800×600×20×45|1800|600|20|45|88200;YB-WH1800×600×20×50|1800|600|20|50|94000;YB-WH1800×650×18×30|1800|650|18|30|70300;YB-WH1800×650×18×36|1800|650|18|36|77900;YB-WH1800×650×18×40|1800|650|18|40|82900;YB-WH1800×650×20×45|1800|650|20|45|92700;YB-WH1800×650×20×50|1800|650|20|50|99000;YB-WH1800×700×18×32|1800|700|18|32|76000;YB-WH1800×700×18×36|1800|700|18|36|81500;YB-WH1800×700×18×40|1800|700|18|40|86900;YB-WH1800×700×20×45|1800|700|20|45|97200;YB-WH1800×700×20×50|1800|700|20|50|104000;YB-WH1800×750×18×32|1800|750|18|32|79200;YB-WH1800×750×18×36|1800|750|18|36|85100;YB-WH1800×750×18×40|1800|750|18|40|90900;YB-WH1800×750×20×45|1800|750|20|45|101700;YB-WH1800×750×20×50|1800|750|20|50|109000;YB-WH1900×650×18×30|1900|650|18|30|72100;YB-WH1900×650×18×36|1900|650|18|36|79700;YB-WH1900×650×18×40|1900|650|18|40|84700;YB-WH1900×650×20×45|1900|650|20|45|94700;YB-WH1900×650×20×50|1900|650|20|50|101000;YB-WH1900×700×18×32|1900|700|18|32|77800;YB-WH1900×700×18×36|1900|700|18|36|83300;YB-WH1900×700×18×40|1900|700|18|40|88700;YB-WH1900×700×20×45|1900|700|20|45|99200;YB-WH1900×700×20×50|1900|700|20|50|106000;YB-WH1900×750×18×34|1900|750|18|34|83900;YB-WH1900×750×18×36|1900|750|18|36|86900;YB-WH1900×750×18×40|1900|750|18|40|92700;YB-WH1900×750×20×45|1900|750|20|45|103700;YB-WH1900×750×20×50|1900|750|20|50|111000;YB-WH1900×800×18×34|1900|800|18|34|87300;YB-WH1900×800×18×36|1900|800|18|36|90500;YB-WH1900×800×18×40|1900|800|18|40|96700;YB-WH1900×800×20×45|1900|800|20|45|108200;YB-WH1900×800×20×50|1900|800|20|50|116000;YB-WH2000×650×18×30|2000|650|18|30|73900;YB-WH2000×650×18×36|2000|650|18|36|81500;YB-WH2000×650×18×40|2000|650|18|40|86500;YB-WH2000×650×20×45|2000|650|20|45|96700;YB-WH2000×650×20×50|2000|650|20|50|103000;YB-WH2000×700×18×32|2000|700|18|32|79600;YB-WH2000×700×18×36|2000|700|18|36|85100;YB-WH2000×700×18×40|2000|700|18|40|90500;YB-WH2000×700×20×45|2000|700|20|45|101200;YB-WH2000×700×20×50|2000|700|20|50|108000;YB-WH2000×750×18×34|2000|750|18|34|85700;YB-WH2000×750×18×36|2000|750|18|36|88700;YB-WH2000×750×18×40|2000|750|18|40|94500;YB-WH2000×750×20×45|2000|750|20|45|105700;YB-WH2000×750×20×50|2000|750|20|50|113000;YB-WH2000×800×18×34|2000|800|18|34|89100;YB-WH2000×800×18×36|2000|800|18|36|92300;YB-WH2000×800×20×40|2000|800|20|40|102400;YB-WH2000×800×20×45|2000|800|20|45|110200;YB-WH2000×800×20×50|2000|800|20|50|118000;YB-WH2000×850×18×36|2000|850|18|36|95900;YB-WH2000×850×18×40|2000|850|18|40|102500;YB-WH2000×850×20×45|2000|850|20|45|114700;YB-WH2000×850×20×50|2000|850|20|50|123000;YB-WH2000×850×20×55|2000|850|20|55|131300",
  "CN-LH": "LH100×50×2.3×3.2|100|50|2.3|3.2|535;LH100×50×3.2×4.5|100|50|3.2|4.5|741;LH100×100×4.5×6|100|100|4.5|6|1596;LH100×100×6×8|100|100|6|8|2104;LH120×120×3.2×4.5|120|120|3.2|4.5|1435;LH120×120×4.5×6|120|120|4.5|6|1926;LH150×75×3.2×4.5|150|75|3.2|4.5|1126;LH150×75×4.5×6|150|75|4.5|6|1521;LH150×100×3.2×4.5|150|100|3.2|4.5|1351;LH150×100×3.2×6|150|100|3.2|6|1642;LH150×100×4.5×6|150|100|4.5|6|1821;LH150×150×3.2×6|150|150|3.2|6|2242;LH150×150×4.5×6|150|150|4.5|6|2421;LH150×150×6×8|150|150|6|8|3204;LH200×100×3×3|200|100|3|3|1182;LH200×100×3.2×4.5|200|100|3.2|4.5|1511;LH200×100×3.2×6|200|100|3.2|6|1802;LH200×100×4.5×6|200|100|4.5|6|2046;LH200×100×6×8|200|100|6|8|2704;LH200×150×3.2×4.5|200|150|3.2|4.5|1961;LH200×150×3.2×6|200|150|3.2|6|2402;LH200×150×4.5×6|200|150|4.5|6|2646;LH200×150×6×8|200|150|6|8|3504;LH200×200×6×8|200|200|6|8|4304;LH250×125×3×3|250|125|3|3|1482;LH250×125×3.2×4.5|250|125|3.2|4.5|1896;LH250×125×3.2×6|250|125|3.2|6|2262;LH250×125×4.5×6|250|125|4.5|6|2571;LH250×125×4.5×8|250|125|4.5|8|3053;LH250×125×6×8|250|125|6|8|3404;LH250×150×3.2×4.5|250|150|3.2|4.5|2121;LH250×150×3.2×6|250|150|3.2|6|2562;LH250×150×4.5×6|250|150|4.5|6|2871;LH250×150×4.5×8|250|150|4.5|8|3453;LH250×150×4.5×9|250|150|4.5|9|3744;LH250×150×6×8|250|150|6|8|3804;LH250×150×6×9|250|150|6|9|4092;LH250×200×4.5×8|250|200|4.5|8|4253;LH250×200×4.5×9|250|200|4.5|9|4644;LH250×200×4.5×10|250|200|4.5|10|5035;LH250×200×6×8|250|200|6|8|4604;LH250×200×6×9|250|200|6|9|4992;LH250×200×6×10|250|200|6|10|5380;LH250×250×4.5×8|250|250|4.5|8|5053;LH250×250×4.5×9|250|250|4.5|9|5544;LH250×250×4.5×10|250|250|4.5|10|6035;LH250×250×6×8|250|250|6|8|5404;LH250×250×6×9|250|250|6|9|5892;LH250×250×6×10|250|250|6|10|6380;LH300×150×3.2×4.5|300|150|3.2|4.5|2281;LH300×150×3.2×6|300|150|3.2|6|2722;LH300×150×4.5×6|300|150|4.5|6|3096;LH300×150×4.5×8|300|150|4.5|8|3678;LH300×150×4.5×9|300|150|4.5|9|3669;LH300×150×4.5×10|300|150|4.5|10|4260;LH300×150×6×8|300|150|6|8|4104;LH300×150×6×9|300|150|6|9|4392;LH300×150×6×10|300|150|6|10|4680;LH300×200×4.5×8|300|200|4.5|8|4478;LH300×200×4.5×9|300|200|4.5|9|4869;LH300×200×4.5×10|300|200|4.5|10|5260;LH300×200×6×8|300|200|6|8|4904;LH300×200×6×9|300|200|6|9|5292;LH300×200×6×10|300|200|6|10|5680;LH300×250×4.5×8|300|250|4.5|8|5278;LH300×250×4.5×9|300|250|4.5|9|5769;LH300×250×4.5×10|300|250|4.5|10|6260;LH300×250×6×8|300|250|6|8|5704;LH300×250×6×9|300|250|6|9|6192;LH300×250×6×10|300|250|6|10|6680;LH350×150×3.2×4.5|350|150|3.2|4.5|2441;LH350×150×3.2×6|350|150|3.2|6|2882;LH350×150×4.5×6|350|150|4.5|6|3321;LH350×150×4.5×8|350|150|4.5|8|3903;LH350×150×4.5×9|350|150|4.5|9|4194;LH350×150×4.5×10|350|150|4.5|10|4485;LH350×150×6×8|350|150|6|8|4404;LH350×150×6×9|350|150|6|9|4692;LH350×150×6×10|350|150|6|10|4980;LH350×175×4.5×6|350|175|4.5|6|3621;LH350×175×4.5×8|350|175|4.5|8|4303;LH350×175×4.5×9|350|175|4.5|9|4644;LH350×175×4.5×10|350|175|4.5|10|4985;LH350×175×6×8|350|175|6|8|4804;LH350×175×6×9|350|175|6|9|5142;LH350×175×6×10|350|175|6|10|5480;LH350×200×4.5×8|350|200|4.5|8|4703;LH350×200×4.5×9|350|200|4.5|9|5094;LH350×200×4.5×10|350|200|4.5|10|5485;LH350×200×6×8|350|200|6|8|5204;LH350×200×6×9|350|200|6|9|5592;LH350×200×6×10|350|200|6|10|5980;LH350×250×4.5×8|350|250|4.5|8|5503;LH350×250×4.5×9|350|250|4.5|9|5994;LH350×250×4.5×10|350|250|4.5|10|6485;LH350×250×6×8|350|250|6|8|6004;LH350×250×6×9|350|250|6|9|6492;LH350×250×6×10|350|250|6|10|6980;LH400×150×4.5×8|400|150|4.5|8|4128;LH400×150×4.5×9|400|150|4.5|9|4419;LH400×150×4.5×10|400|150|4.5|10|4710;LH400×150×6×8|400|150|6|8|4704;LH400×150×6×9|400|150|6|9|4992;LH400×150×6×10|400|150|6|10|5280;LH400×200×4.5×8|400|200|4.5|8|4928;LH400×200×4.5×9|400|200|4.5|9|5319;LH400×200×4.5×10|400|200|4.5|10|5710;LH400×200×6×8|400|200|6|8|5504;LH400×200×6×9|400|200|6|9|5892;LH400×200×6×10|400|200|6|10|6280;LH400×250×4.5×8|400|250|4.5|8|5728;LH400×250×4.5×9|400|250|4.5|9|6219;LH400×250×4.5×10|400|250|4.5|10|6710;LH400×250×6×8|400|250|6|8|6304;LH400×250×6×9|400|250|6|9|6792;LH400×250×6×10|400|250|6|10|7280;LH450×200×4.5×8|450|200|4.5|8|5153;LH450×200×4.5×9|450|200|4.5|9|5544;LH450×200×4.5×10|450|200|4.5|10|5935;LH450×200×6×8|450|200|6|8|5804;LH450×200×6×9|450|200|6|9|6192;LH450×200×6×10|450|200|6|10|6580;LH450×250×4.5×8|450|250|4.5|8|5953;LH450×250×4.5×9|450|250|4.5|9|6444;LH450×250×4.5×10|450|250|4.5|10|6935;LH450×250×6×8|450|250|6|8|6604;LH450×250×6×9|450|250|6|9|7092;LH450×250×6×10|450|250|6|10|7580;LH500×200×4.5×8|500|200|4.5|8|5378;LH500×200×4.5×9|500|200|4.5|9|5769;LH500×200×4.5×10|500|200|4.5|10|6160;LH500×200×6×8|500|200|6|8|6104;LH500×200×6×9|500|200|6|9|6492;LH500×200×6×10|500|200|6|10|6880;LH500×250×4.5×8|500|250|4.5|8|6178;LH500×250×4.5×9|500|250|4.5|9|6669;LH500×250×4.5×10|500|250|4.5|10|7160;LH500×250×6×8|500|250|6|8|6904;LH500×250×6×9|500|250|6|9|7392;LH500×250×6×10|500|250|6|10|7880",
  "CN-I": "I10|100|68|4.5|7.6|1434.5;I12|120|74|5|8.4|1781.8;I12.6|126|74|5|8.4|1811.8;I14|140|80|5.5|9.1|2151.6;I16|160|88|6|9.9|2613.1;I18|180|94|6.5|10.7|3075.6;I20a|200|100|7|11.4|3557.8;I20b|200|102|9|11.4|3957.8;I22a|220|110|7.5|12.3|4212.8;I22b|220|112|9.5|12.3|4652.8;I24a|240|116|8|13|4774.1;I24b|240|118|10|13|5254.1;I25a|250|116|8|13|4854.1;I25b|250|118|10|13|5354.1;I27a|270|122|8.5|13.7|5455.4;I27b|270|124|10.5|13.7|5995.4;I28a|280|122|8.5|13.7|5540.4;I28b|280|124|10.5|13.7|6100.4;I30a|300|126|9|14.4|6125.4;I30b|300|128|11|14.4|6725.4;I30c|300|130|13|14.4|7325.4;I32a|320|130|9.5|15|6715.6;I32b|320|132|11.5|15|7355.6;I32c|320|134|13.5|15|7995.6;I36a|360|136|10|15.8|7648;I36b|360|138|12|15.8|8368;I36c|360|140|14|15.8|9088;I40a|400|142|10.5|16.5|8611.2;I40b|400|141|12.5|16.5|9411.2;I40c|400|146|14.5|16.5|10211.2;I45a|450|150|11.5|18|10244.6;I45b|450|152|13.5|18|11144.6;I45c|450|154|15.5|18|12044.6;I50a|500|158|12|20|11930.4;I50b|500|160|14|20|12930.4;I50c|500|162|16|20|13930.4;I55a|550|166|12.5|21|13418.5;I55b|550|168|14.5|21|14518.5;I55c|550|170|16.5|21|15618.5;I56a|560|166|12.5|21|13543.5;I56b|560|166|14.5|21|14663.5;I56c|560|168|16.5|21|15783.5;I63a|630|176|13|22|15465.8;I63b|630|178|15|22|16725.8;I63c|630|180|17|22|17985.8",
  "CN-YB-I": "YB-I10|100|55|4.5|7.2|1200;YB-I12|120|64|4.8|7.3|1470;YB-I14|140|73|4.9|7.5|1740;YB-I16|160|81|5|7.8|2020;YB-I18|180|90|5.1|8.1|2340;YB-I18a|180|100|5.1|8.3|2540;YB-I20|200|100|5.2|8.4|2680;YB-I20a|200|110|5.2|8.6|2890;YB-I22|220|110|5.4|8.7|3060;YB-I22a|220|120|5.4|8.9|3280;YB-I24|240|115|5.6|9.5|3480;YB-I24a|240|125|5.6|9.8|3750;YB-I27|270|125|6|9.8|4020;YB-I27a|270|135|6|10.2|4320;YB-I30|300|135|6.5|10.2|4650;YB-I30a|300|145|6.5|10.7|4990;YB-I33|330|140|7|11.2|5380;YB-I36|360|145|7.5|12.3|6190;YB-I40|400|155|8|13|7140;YB-I45|450|160|8.6|14.2|8300;YB-I50|500|170|9.5|15.2|9780;YB-I55|550|180|10.3|16.5|11400;YB-I60|600|190|11.1|17.8|13200;YB-I65|650|200|12|19.2|15300;YB-I70|700|210|13|20.8|17600;YB-I70a|700|210|15|24|20200;YB-I70b|700|210|17.5|28.2|23400",
  "CN-C": "[5|50|37|4.5|7|692.8;[6.3|63|40|4.8|7.5|845.1;[6.5|65|40|4.3|7.5|854.7;[8|80|43|5|8|1024.8;[10|100|48|5.3|8.5|1274.8;[12|120|53|5.5|9|1536.2;[12.6|126|53|5.5|9|1569.2;[14a|140|58|6|9.5|1851.6;[14b|140|60|8|9.5|2131.6;[16a|160|63|6.5|10|2196.2;[16b|160|65|8.5|10|2516.2;[18a|180|68|7|10.5|2569.9;[18b|180|70|9|10.5|2929.9;[20a|200|73|7|11|2883.7;[20b|200|75|9|11|3283.7;[22a|220|77|7|11.5|3184.6;[22b|220|79|9|11.5|3624.6;[24a|240|78|7|12|3421.7;[24b|240|80|9|12|3901.7;[24c|240|82|11|12|4381.7;[25a|250|78|7|12|3491.7;[25b|250|80|9|12|3991.7;[25c|250|82|11|12|4491.7;[27a|270|82|7.5|12.5|3928.4;[27b|270|84|9.5|12.5|4468.4;[27c|270|86|11.5|12.5|5008.4;[28a|280|82|7.5|12.5|4003.4;[28b|280|84|9.5|12.5|4563.4;[28c|280|86|11.5|12.5|5123.4;[30a|300|85|7.5|13.5|4390.2;[30b|300|87|9.5|13.5|4990.2;[30c|300|89|11.5|13.5|5590.2;[32a|320|88|8|14|4851.3;[32b|320|90|10|14|5491.3;[32c|320|92|12|14|6131.3;[36a|360|96|9|16|6091;[36b|360|98|11|16|6811;[36c|360|100|13|16|7531;[40a|400|100|10.5|18|7506.8;[40b|400|102|12.5|18|8306.8;[40c|400|104|14.5|18|9106.8",
  "CN-YB-C": "YB-C5|50|32|4.4|7|616;YB-C6.5|65|36|4.4|7.2|751;YB-C8|80|40|4.5|7.4|898;YB-C10|100|46|4.5|7.6|1090;YB-C12|120|52|4.8|7.8|1330;YB-C14|140|58|4.9|8.1|1560;YB-C14a|140|62|4.9|8.7|1700;YB-C16|160|64|5|8.4|1810;YB-C16a|160|68|5|9|1950;YB-C18|180|70|5.1|8.7|2070;YB-C18a|180|74|5.1|9.3|2220;YB-C20|200|76|5.2|9|2340;YB-C20a|200|80|5.2|9.7|2520;YB-C22|220|82|5.4|9.5|2670;YB-C22a|220|87|5.4|10.2|2880;YB-C24|240|90|5.6|10|3060;YB-C24a|240|95|5.6|10.7|3290;YB-C27|270|95|6|10.5|3520;YB-C30|300|100|6.5|11|4050;YB-C33|330|105|7|11.7|4650;YB-C36|360|110|7.5|12.6|5340;YB-C40|400|115|8|13.5|6150",
  "CN-TW": "TW50×100×6×8|50|100|6|8|1079;TW62.5×125×6.5×9|62.5|125|6.5|9|1500;TW75×150×7×10|75|150|7|10|1982;TW87.5×175×7.5×11|87.5|175|7.5|11|2571;TW100×200×8×12|100|200|8|12|3177;TW100×200×12×12|100|204|12|12|3577;TW125×250×9×14|125|250|9|14|4572;TW125×250×14×14|125|255|14|14|5197;TW150×300×12×12|150|300|12|12|5317;TW150×300×10×15|150|300|10|15|5923;TW150×300×15×15|150|305|15|15|6673;TW175×350×10×16|172|348|10|16|7201;TW175×350×12×19|175|350|12|19|8595;TW200×400×15×15|194|402|15|15|8923;TW200×400×11×18|197|398|11|18|9341;TW200×400×13×21|200|400|13|21|10935;TW200×400×21×21|200|408|21|21|12535;TW200×400×18×28|207|405|18|28|14770;TW200×400×20×35|214|407|20|35|18033",
  "CN-TM": "TM75×100×6×9|74|100|6|9|1317;TM100×150×6×9|97|150|6|9|1905;TM125×175×7×11|122|175|7|11|2775;TM150×200×8×12|147|200|8|12|3553;TM175×250×9×14|170|250|9|14|4977;TM200×300×10×16|195|300|10|16|6663;TM225×300×11×18|220|300|11|18|7695;TM250×300×11×15|241|300|11|15|7059;TM250×300×11×18|244|300|11|18|7959;TM275×300×11×15|272|300|11|15|7400;TM275×300×11×18|275|300|11|18|8300;TM300×300×12×17|291|300|12|17|8461;TM300×300×12×20|294|300|12|20|9361;TM300×300×14×23|297|302|14|23|10855",
  "CN-TN": "TN50×50×5×7|50|50|5|7|592;TN62.5×60×6×8|62.5|60|6|8|834;TN75×75×5×7|75|75|5|7|892;TN87.5×90×5×8|87.5|90|5|8|1145;TN100×100×4.5×7|99|99|4.5|7|1134;TN100×100×5.5×8|100|100|5.5|8|1333;TN125×125×5×8|124|124|5|8|1599;TN125×125×6×9|125|125|6|9|1848;TN150×150×5.5×8|149|149|5.5|8|2040;TN150×150×6.5×9|150|150|6.5|9|2339;TN175×175×6×9|173|174|6|9|2623;TN175×175×7×11|175|175|7|11|3146;TN200×200×7×11|198|199|7|11|3571;TN200×200×8×13|200|200|8|13|4169;TN225×200×8×12|223|199|8|12|4149;TN225×200×9×14|225|200|9|14|4772;TN250×200×9×14|248|199|9|14|4965;TN250×200×10×16|250|200|10|16|5613;TN250×200×11×19|253|201|11|19|6466;TN275×200×9×14|273|199|9|14|5190;TN275×200×10×16|275|200|10|16|5863;TN300×200×10×15|298|199|10|15|5888;TN300×200×11×17|300|200|11|17|6586;TN300×200×12×20|303|201|12|20|7489;TN325×300×10×15|323|299|10|15|7627;TN325×300×11×17|325|300|11|17|8561;TN325×300×12×20|328|301|12|20|9789;TN350×300×13×20|346|300|13|20|10311;TN350×300×13×24|350|300|13|24|11511;TN400×300×14×22|396|300|14|22|11975;TN400×300×14×26|400|300|14|26|13175;TN450×300×15×23|445|299|15|23|13346;TN450×300×16×28|450|300|16|28|15291;TN450×300×18×34|456|302|18|34|18003",
  "CN-LN": "25|16|3|116.2;32|20|3|149.2;32|20|4|193.9;40|25|3|189;40|25|4|246.7;45|28|4|280.6;50|32|3|243.1;50|32|4|317.7;56|36|3|274.3;56|36|4|359;56|36|5|441.5;63|40|4|405.8;63|40|5|499.3;63|40|6|590.8;63|40|7|680.2;70|45|4|454.7;70|45|5|560.9;70|45|6|664.7;70|45|7|765.7;75|50|5|612.5;75|50|6|726;75|50|8|946.7;75|50|10|1159;80|50|5|637.5;80|50|6|756;80|50|7|872.4;80|50|8|986.7;90|56|5|721.2;90|56|6|855.7;90|56|7|988;90|56|8|1118.3;100|63|6|961.7;100|63|7|1111.1;100|63|8|1258.4;100|63|10|1546.7;100|80|6|1063.7;100|80|7|1230.1;100|80|8|1394.4;100|80|10|1716.7;110|70|6|1063.7;110|70|7|1230.1;110|70|8|1394.4;110|70|10|1716.7;125|80|7|1409.6;125|80|8|1598.9;125|80|10|1971.2;125|80|12|2335.1;140|90|8|1803.8;140|90|10|2226.1;140|90|12|2640;140|90|14|3045.6;150|90|8|1883.9;150|90|10|2326.1;150|90|12|2760;150|90|14|3185.6;150|90|15|3395.2;150|90|16|3602.7;160|100|10|2531.5;160|100|12|3005.4;160|100|14|3470.9;160|100|16|3928.1;180|110|10|2837.3;180|110|12|3371.2;180|110|14|3896.7;180|110|16|4413.9;200|125|12|3791.2;200|125|14|4386.7;200|125|16|4973.9;200|125|18|5552.6",
  "CN-L": "20|20|3|113.2;20|20|4|145.9;25|25|3|143.2;25|25|4|185.9;30|30|3|174.9;30|30|4|227.6;36|36|3|210.9;36|36|4|275.6;36|36|5|338.2;40|40|3|235.9;40|40|4|308.6;40|40|5|379.1;45|45|3|265.9;45|45|4|348.6;45|45|5|429.2;45|45|6|507.6;50|50|3|297.1;50|50|4|389.7;50|50|5|480.3;50|50|6|568.8;56|56|3|334.3;56|56|4|439;56|56|5|541.5;56|56|6|642;56|56|7|740.4;56|56|8|836.7;60|60|5|582.9;60|60|6|691.4;60|60|7|797.7;60|60|8|902;63|63|4|497.8;63|63|5|614.3;63|63|6|728.8;63|63|7|841.2;63|63|8|951.5;63|63|10|1165.7;70|70|4|557;70|70|5|687.5;70|70|6|816;70|70|7|942.4;70|70|8|1066.7;75|75|5|741.2;75|75|6|879.7;75|75|7|1016;75|75|8|1150.3;75|75|9|1282.5;75|75|10|1412.6;80|80|5|791.2;80|80|6|939.7;80|80|7|1086;80|80|8|1230.3;80|80|9|1372.5;80|80|10|1512.6;90|90|6|1063.7;90|90|7|1230.1;90|90|8|1394.4;90|90|9|1556.6;90|90|10|1716.7;90|90|12|2030.6;100|100|6|1193.2;100|100|7|1379.6;100|100|8|1563.8;100|100|9|1746.2;100|100|10|1926.1;100|100|12|2280;100|100|14|2625.6;100|100|16|2962.7;110|110|7|1519.6;110|110|8|1723.8;110|110|10|2126.1;110|110|12|2520;110|110|14|2905.6;125|125|8|1975;125|125|10|2437.4;125|125|12|2891.2;125|125|14|3336.7;125|125|16|3773.9;140|140|10|2737.3;140|140|12|3251.2;140|140|14|3756.7;140|140|16|4253.9;150|150|8|2375;150|150|10|2937.3;150|150|12|3491.2;150|150|14|4036.7;150|150|15|4306.3;150|150|16|4573.9;160|160|10|3150.2;160|160|12|3741.1;160|160|14|4329.6;160|160|16|4906.7;180|180|12|4224.1;180|180|14|4889.6;180|180|16|5546.7;180|180|18|6195.5;200|200|14|5464.2;200|200|16|6201.3;200|200|18|6930.1;200|200|20|7650.5;200|200|24|9066.1;220|220|16|6866.4;220|220|18|7675.2;220|220|20|8475.6;220|220|22|9267.6;220|220|24|10051.2;220|220|26|10826.4;250|250|18|8784.2;250|250|20|9704.5;250|250|24|11520.1;250|250|26|12415.4;250|250|28|13302.2;250|250|30|14180.7;250|250|32|15050.8;250|250|35|16340.2",
  "CN-GB-SSP": "32|2.5|232;32|3|273;32|3.5|313;32|4|352;38|2.5|279;38|3|330;38|3.5|379;38|4|427;42|2.5|310;42|3|368;42|3.5|423;42|4|478;45|2.5|334;45|3|396;45|3.5|456;45|4|515;50|2.5|373;50|3|443;50|3.5|511;50|4|578;50|4.5|643;50|5|707;54|3|481;54|3.5|555;54|4|628;54|4.5|700;54|5|770;54|5.5|838;54|6|905;57|3|509;57|3.5|588;57|4|666;57|4.5|742;57|5|817;57|5.5|890;57|6|961;60|3|537;60|3.5|621;60|4|704;60|4.5|785;60|5|864;60|5.5|942;60|6|1018;63.5|3|570;63.5|3.5|660;63.5|4|748;63.5|4.5|834;63.5|5|919;63.5|5.5|1002;63.5|6|1084;68|3|613;68|3.5|709;68|4|804;68|4.5|898;68|5|990;68|5.5|1080;68|6|1169;70|3|631;70|3.5|731;70|4|829;70|4.5|926;70|5|1021;70|5.5|1114;70|6|1206;73|3|660;73|3.5|764;73|4|867;73|4.5|968;73|5|1068;73|5.5|1166;73|6|1263;76|3|688;76|3.5|797;76|4|905;76|4.5|1011;76|5|1115;76|5.5|1218;76|6|1319;83|3.5|874;83|4|993;83|4.5|1110;83|5|1225;83|5.5|1339;83|6|1451;83|6.5|1562;83|7|1671;89|3.5|940;89|4|1068;89|4.5|1195;89|5|1319;89|5.5|1443;89|6|1575;89|6.5|1685;89|7|1803;95|3.5|1006;95|4|1144;95|4.5|1279;95|5|1414;95|5.5|1546;95|6|1678;95|6.5|1807;95|7|1935;102|3.5|1083;102|4|1232;102|4.5|1378;102|5|1524;102|5.5|1667;102|6|1810;102|6.5|1950;102|7|2089;108|4|1306;108|4.5|1462;108|5|1617;108|5.5|1770;108|6|1922;108|6.5|2072;108|7|2220;108|7.5|2367;108|8|2512;114|4|1382;114|4.5|1548;114|5|1712;114|5.5|1875;114|6|2036;114|6.5|2195;114|7|2353;114|7.5|2509;114|8|2664;121|4|1470;121|4.5|1647;121|5|1822;121|5.5|1996;121|6|2168;121|6.5|2338;121|7|2507;121|7.5|2674;121|8|2840;127|4|1546;127|4.5|1732;127|5|1916;127|5.5|2099;127|6|2281;127|6.5|2461;127|7|2639;127|7.5|2816;127|8|2991;133|4|1621;133|4.5|1817;133|5|2011;133|5.5|2203;133|6|2394;133|6.5|2583;133|7|2771;133|7.5|2957;133|8|3142;140|4.5|1916;140|5|2121;140|5.5|2324;140|6|2526;140|6.5|2726;140|7|2925;140|7.5|3122;140|8|3318;140|9|3704;140|10|4084;146|4.5|2000;146|5|2215;146|5.5|2428;146|6|2639;146|6.5|2849;146|7|3057;146|7.5|3263;146|8|3468;146|9|3874;146|10|4273;152|4.5|2085;152|5|2309;152|5.5|2531;152|6|2752;152|6.5|2971;152|7|3189;152|7.5|3405;152|8|3619;152|9|4043;152|10|4461;159|4.5|2184;159|5|2419;159|5.5|2652;159|6|2884;159|6.5|3114;159|7|3343;159|7.5|3570;159|8|3795;159|9|4221;159|10|4681;168|4.5|2311;168|5|2560;168|5.5|2808;168|6|3054;168|6.5|3298;168|7|3541;168|7.5|3782;168|8|4021;168|9|4496;168|10|4964;180|5|2749;180|5.5|3015;180|6|3280;180|6.5|3543;180|7|3804;180|7.5|4064;180|8|4323;180|9|4835;180|10|5341;180|12|6333;194|5|2969;194|5.5|3257;194|6|3544;194|6.5|3829;194|7|4112;194|7.5|4394;194|8|4675;194|9|5231;194|10|5781;194|12|6861;203|6|3713;203|6.5|4013;203|7|4310;203|7.5|4606;203|8|4901;203|9|5485;203|10|6063;203|12|7201;203|14|8313;203|16|9400;219|6|4015;219|6.5|4339;219|7|4662;219|7.5|4983;219|8|5303;219|9|5938;219|10|6566;219|12|7804;219|14|9016;219|16|10204;245|6.5|4870;245|7|5234;245|7.5|5596;245|8|5956;245|9|6673;245|10|7383;245|12|8784;245|14|10160;245|16|11511;273|6.5|5442;273|7|5850;273|7.5|6256;273|8|6660;273|9|7464;273|10|8262;273|12|9839;273|14|11491;273|16|12918;299|7.5|6868;299|8|7314;299|9|8200;299|10|9079;299|12|10820;299|14|12535;299|16|14225;325|7.5|7481;325|8|7967;325|9|8935;325|10|9896;325|12|11800;325|14|13678;325|16|15532;351|8|8621;351|9|9670;351|10|10713;351|12|12780;351|14|14822;351|16|16839;377|9|10400;377|10|11524;377|11|12642;377|12|13753;377|13|14859;377|14|15958;377|15|17050;377|16|18137;402|9|11106;402|10|12309;402|11|13505;402|12|14695;402|13|15879;402|14|17056;402|15|18228;402|16|19393;426|9|11784;426|10|13062;426|11|14334;426|12|15600;426|13|16859;426|14|18112;426|15|19358;426|16|20598;450|9|12463;450|10|13861;450|11|15163;450|12|16504;450|13|17838;450|14|19167;450|15|20489;450|16|21804;465|9|12887;465|10|14287;465|11|15681;465|12|17069;465|13|18451;465|14|19826;465|15|21195;465|16|22558;480|9|13311;480|10|14758;480|11|16199;480|12|17634;480|13|19063;480|14|20485;480|15|21902;480|16|23311;500|9|13876;500|10|15386;500|11|16890;500|12|18388;500|13|19879;500|14|21365;500|15|22844;500|16|24316;530|9|14723;530|10|16328;530|11|17926;530|12|19518;530|13|21104;530|14|22683;530|15|24257;530|16|25823;550|9|15289;550|10|16956;550|11|18617;550|12|20272;550|13|21920;550|14|23563;550|15|25199;550|16|26828;560|9|15571;560|10|17270;560|11|18962;560|12|20649;560|13|22329;560|14|24002;560|15|25670;560|16|27331;600|9|16702;600|10|18526;600|11|20344;600|12|22156;600|13|23961;600|14|25761;600|15|27554;600|16|29340;630|9|17550;630|10|19468;630|11|21380;630|12|23286;630|13|25186;630|14|27079;630|15|28967;630|16|30847",
  "CN-GB-SPWSP": "219.1|5|3361;219.1|6|4015;219.1|7|4662;219.1|8|5303;244.5|5|3760;244.5|6|4493;244.5|7|5220;244.5|8|5941;273|6|5030;273|7|5847;273|8|6657;323.9|6|5989;323.9|7|6965;323.9|8|7935;325|6|6010;325|7|6990;325|8|7963;355.6|6|6587;355.6|7|7662;355.6|8|8732;377|6|6990;377|7|8133;377|8|9269;377|9|10400;406.4|6|7544;406.4|7|8779;406.4|8|10009;406.4|9|11231;406.4|10|12447;426|6|7913;426|7|9210;426|8|10500;426|9|11784;426|10|13062;457|6|8497;457|7|9891;457|8|11279;457|9|12660;457|10|14036;457|11|15405;457|12|16768;478|6|8893;478|7|10353;478|8|11806;478|9|13254;478|10|14695;478|11|16130;478|12|17559;508|6|9458;508|7|11012;508|8|12560;508|9|14102;508|10|15637;508|11|17166;508|12|18689;559|6|10419;559|7|12133;559|8|13841;559|9|15543;559|10|17239;559|11|18928;559|12|20611;559|13|22288;610|6|11379;610|7|13254;610|8|15122;610|9|16984;610|10|18840;610|11|20689;610|12|22533;610|13|24370;630|6|11756;630|7|13694;630|8|15625;630|9|17550;630|10|19468;630|11|21380;630|12|23286;630|13|25186;660|6|12321;660|7|14353;660|8|16378;660|9|18397;660|10|20410;660|11|22416;660|12|24417;660|13|26411;711|6|13282;711|7|15474;711|8|17659;711|9|19839;711|10|22011;711|11|24178;711|12|26338;711|13|28492;720|6|13452;720|7|15672;720|8|17785;720|9|20093;720|10|22294;720|11|24489;720|12|26677;720|13|28860;762|7|16595;762|8|18940;762|9|21280;762|10|23613;762|11|25940;762|12|28260;762|13|30574;762|14|32882;813|7|17716;813|8|20222;813|9|22721;813|10|25241;813|11|27701;813|12|30182;813|13|32656;813|14|35124;820|7|17870;820|8|20397;820|9|22919;820|10|25434;820|11|27943;820|12|30445;820|13|32942;820|14|35432;820|15|37916;820|16|41393;914|8|22759;914|9|25575;914|10|28386;914|11|31190;914|12|33987;914|13|36779;914|14|39564;914|15|42343;914|16|45116;920|8|22909;920|9|25745;920|10|28574;920|11|31397;920|12|34213;920|13|37024;920|14|39828;920|15|42626;920|16|45417;1020|8|25421;1020|9|28571;1020|10|31714;1020|11|34851;1020|12|37981;1020|13|41106;1020|14|44224;1020|15|47336;1020|16|50441;1120|8|27933;1120|9|31397;1120|10|34854;1120|11|38305;1120|12|41749;1120|13|45188;1120|14|48620;1120|15|52046;1120|16|55465;1220|10|37994;1220|11|41759;1220|12|45517;1220|13|49270;1220|14|53016;1220|15|56756;1220|16|60489;1420|10|44274;1420|11|48667;1420|12|53053;1420|13|57434;1420|14|61808;1420|15|66176;1420|16|70537",
  "CN-YB-STWSP": "32|2|188;32|2.5|232;38|2|226;38|2.5|279;40|2|239;40|2.5|295;42|2|251;42|2.5|310;45|2|270;45|2.5|334;45|3|396;51|2|308;51|2.5|381;51|3|452;51|3.5|522;53|2|320;53|2.5|397;53|3|471;53|3.5|544;57|2|346;57|2.5|428;57|3|509;57|3.5|588;60|2|364;60|2.5|452;60|3|537;60|3.5|621;63.5|2|386;63.5|2.5|479;63.5|3|570;63.5|3.5|660;70|2|427;70|2.5|530;70|3|631;70|3.5|731;70|4.5|926;76|2|465;76|2.5|577;76|3|688;76|3.5|797;76|4|905;76|4.6|1011;83|2|509;83|2.5|632;83|3|754;83|3.5|874;83|4|993;83|4.5|1110;89|2|547;89|2.5|679;89|3|811;89|3.5|940;89|4|1068;89|4.5|1195;95|2|584;95|2.5|726;95|3|867;95|3.5|1006;102|2|628;102|2.5|781;102|3|933;102|3.5|1083;102|4|1232;102|4.5|1378;102|5|1524;108|3|990;108|3.5|1149;108|4|1307;114|3|1046;114|3.5|1215;114|4|1382;114|4.5|1548;114|5|1712;121|3|1112;121|3.5|1292;121|4|1470;127|3|1169;127|3.5|1358;127|4|1546;127|4.5|1732;127|5|1916;133|3.5|1424;133|4|1621;133|4.5|1817;133|5|2011;140|3.5|1501;140|4|1709;140|4.5|1916;140|5|2121;140|5.5|2324;152|3.5|1633;152|4|1860;152|4.5|2085;152|5|2309;152|5.5|2531",
  "JP-H": "JIS H1000×400×19×40|1000|400|19|40|49760;JIS H1000×400×19×36|1000|400|19|36|46710;JIS H1000×400×19×32|1000|400|19|32|43660;JIS H1000×400×19×28|1000|400|19|28|40610;JIS H1000×400×19×25|1000|400|19|25|38330;JIS H1000×400×16×32|1000|400|16|32|40850;JIS H1000×400×16×28|1000|400|16|28|37780;JIS H1000×400×16×25|1000|400|16|25|35480;JIS H1000×400×16×22|1000|400|16|22|33170;JIS H1000×350×19×40|1000|350|19|40|45760;JIS H1000×350×19×36|1000|350|19|36|43110;JIS H1000×350×19×32|1000|350|19|32|40460;JIS H1000×350×19×28|1000|350|19|28|37810;JIS H1000×350×19×25|1000|350|19|25|35830;JIS H1000×350×16×32|1000|350|16|32|37650;JIS H1000×350×16×28|1000|350|16|28|34980;JIS H1000×350×16×25|1000|350|16|25|32980;JIS H1000×350×16×22|1000|350|16|22|30970;JIS H1000×300×19×40|1000|300|19|40|41760;JIS H1000×300×19×36|1000|300|19|36|39510;JIS H1000×300×19×32|1000|300|19|32|37260;JIS H1000×300×19×28|1000|300|19|28|35010;JIS H1000×300×19×25|1000|300|19|25|33330;JIS H1000×300×16×32|1000|300|16|32|34450;JIS H1000×300×16×28|1000|300|16|28|32180;JIS H1000×300×16×25|1000|300|16|25|30480;JIS H1000×300×16×22|1000|300|16|22|28770;JIS H1000×250×19×36|1000|250|19|36|35910;JIS H1000×250×19×32|1000|250|19|32|34060;JIS H1000×250×19×28|1000|250|19|28|32210;JIS H1000×250×19×25|1000|250|19|25|30830;JIS H1000×250×16×32|1000|250|16|32|31250;JIS H1000×250×16×28|1000|250|16|28|29380;JIS H1000×250×16×25|1000|250|16|25|27980;JIS H1000×250×16×22|1000|250|16|22|26570;JIS H950×400×19×40|950|400|19|40|48810;JIS H950×400×19×36|950|400|19|36|45760;JIS H950×400×19×32|950|400|19|32|42710;JIS H950×400×19×28|950|400|19|28|39660;JIS H950×400×19×25|950|400|19|25|37380;JIS H950×400×16×32|950|400|16|32|40050;JIS H950×400×16×28|950|400|16|28|36980;JIS H950×400×16×25|950|400|16|25|34680;JIS H950×400×16×22|950|400|16|22|32370;JIS H950×350×19×40|950|350|19|40|44810;JIS H950×350×19×36|950|350|19|36|42160;JIS H950×350×19×32|950|350|19|32|39510;JIS H950×350×19×28|950|350|19|28|36860;JIS H950×350×19×25|950|350|19|25|34880;JIS H950×350×16×32|950|350|16|32|36850;JIS H950×350×16×28|950|350|16|28|34180;JIS H950×350×16×25|950|350|16|25|32180;JIS H950×350×16×22|950|350|16|22|30170;JIS H950×300×19×40|950|300|19|40|40810;JIS H950×300×19×36|950|300|19|36|38560;JIS H950×300×19×32|950|300|19|32|36310;JIS H950×300×19×28|950|300|19|28|34060;JIS H950×300×19×25|950|300|19|25|32380;JIS H950×300×16×32|950|300|16|32|33650;JIS H950×300×16×28|950|300|16|28|31380;JIS H950×300×16×25|950|300|16|25|29680;JIS H950×300×16×22|950|300|16|22|27970;JIS H950×250×19×36|950|250|19|36|34960;JIS H950×250×19×32|950|250|19|32|33110;JIS H950×250×19×28|950|250|19|28|31260;JIS H950×250×19×25|950|250|19|25|29880;JIS H950×250×16×32|950|250|16|32|30450;JIS H950×250×16×28|950|250|16|28|28580;JIS H950×250×16×25|950|250|16|25|27180;JIS H950×250×16×22|950|250|16|22|25770;JIS H900×400×19×40|900|400|19|40|47860;JIS H900×400×19×36|900|400|19|36|44810;JIS H900×400×19×32|900|400|19|32|41760;JIS H900×400×19×28|900|400|19|28|38710;JIS H900×400×16×32|900|400|16|32|39250;JIS H900×400×16×28|900|400|16|28|36180;JIS H900×400×16×25|900|400|16|25|33880;JIS H900×350×19×40|900|350|19|40|43860;JIS H900×350×19×36|900|350|19|36|41210;JIS H900×350×19×32|900|350|19|32|38560;JIS H900×350×19×28|900|350|19|28|35910;JIS H900×350×19×25|900|350|19|25|33930;JIS H900×350×16×32|900|350|16|32|36050;JIS H900×350×16×28|900|350|16|28|33380;JIS H900×350×16×25|900|350|16|25|31380;JIS H900×300×19×32|900|300|19|32|35360;JIS H900×300×19×28|900|300|19|28|33110;JIS H900×300×19×25|900|300|19|25|31430;JIS H900×300×19×22|900|300|19|22|29740;JIS H900×300×16×32|900|300|16|32|32850;JIS H900×300×16×28|900|300|16|28|30580;JIS H900×300×16×25|900|300|16|25|28880;JIS H900×300×16×22|900|300|16|22|27170;JIS H900×300×16×19|900|300|16|19|25470;JIS H900×250×16×28|900|250|16|28|27780;JIS H900×250×16×25|900|250|16|25|26380;JIS H900×250×16×22|900|250|16|22|24970;JIS H900×250×16×19|900|250|16|19|23570;JIS H850×400×19×40|850|400|19|40|46910;JIS H850×400×19×36|850|400|19|36|43860;JIS H850×400×19×32|850|400|19|32|40810;JIS H850×400×19×28|850|400|19|28|37760;JIS H850×400×16×32|850|400|16|32|38450;JIS H850×400×16×28|850|400|16|28|35380;JIS H850×400×16×25|850|400|16|25|33080;JIS H850×350×19×40|850|350|19|40|42910;JIS H850×350×19×36|850|350|19|36|40260;JIS H850×350×19×32|850|350|19|32|37610;JIS H850×350×19×28|850|350|19|28|34960;JIS H850×350×16×32|850|350|16|32|35250;JIS H850×350×16×28|850|350|16|28|32580;JIS H850×350×16×25|850|350|16|25|30580;JIS H850×300×16×32|850|300|16|32|32050;JIS H850×300×16×28|850|300|16|28|29780;JIS H850×300×16×25|850|300|16|25|28080;JIS H850×300×16×22|850|300|16|22|26370;JIS H850×250×16×28|850|250|16|28|26980;JIS H850×250×16×25|850|250|16|25|25580;JIS H850×250×14×25|850|250|14|25|23980;JIS H850×250×14×22|850|250|14|22|22560;JIS H800×400×19×40|800|400|19|40|45960;JIS H800×400×19×36|800|400|19|36|42910;JIS H800×400×19×32|800|400|19|32|39860;JIS H800×400×19×28|800|400|19|28|36810;JIS H800×400×16×36|800|400|16|36|40730;JIS H800×400×16×32|800|400|16|32|37650;JIS H800×400×16×28|800|400|16|28|34580;JIS H800×400×16×25|800|400|16|25|32280;JIS H800×400×14×28|800|400|14|28|33090;JIS H800×400×14×25|800|400|14|25|30780;JIS H800×350×19×40|800|350|19|40|41960;JIS H800×350×19×36|800|350|19|36|39310;JIS H800×350×19×32|800|350|19|32|36660;JIS H800×350×19×28|800|350|19|28|34010;JIS H800×350×19×25|800|350|19|25|32030;JIS H800×350×16×36|800|350|16|36|37130;JIS H800×350×16×32|800|350|16|32|34450;JIS H800×350×16×28|800|350|16|28|31780;JIS H800×350×16×25|800|350|16|25|29780;JIS H800×350×14×28|800|350|14|28|30290;JIS H800×350×14×25|800|350|14|25|28280;JIS H800×300×16×32|800|300|16|32|31250;JIS H800×300×16×28|800|300|16|28|28980;JIS H800×300×16×25|800|300|16|25|27280;JIS H800×300×16×22|800|300|16|22|25570;JIS H800×300×14×28|800|300|14|28|27490;JIS H800×300×14×25|800|300|14|25|25780;JIS H800×300×14×22|800|300|14|22|24060;JIS H800×250×16×32|800|250|16|32|28050;JIS H800×250×16×28|800|250|16|28|26180;JIS H800×250×16×25|800|250|16|25|24780;JIS H800×250×14×25|800|250|14|25|23280;JIS H800×250×14×22|800|250|14|22|21860;JIS H750×350×16×36|750|350|16|36|36330;JIS H750×350×16×32|750|350|16|32|33650;JIS H750×350×16×28|750|350|16|28|30980;JIS H750×350×14×32|750|350|14|32|32280;JIS H750×350×14×28|750|350|14|28|29590;JIS H750×350×14×25|750|350|14|25|27580;JIS H750×300×16×32|750|300|16|32|30450;JIS H750×300×16×28|750|300|16|28|28180;JIS H750×300×16×25|750|300|16|25|26480;JIS H750×300×14×28|750|300|14|28|26790;JIS H750×300×14×25|750|300|14|25|25080;JIS H750×300×14×22|750|300|14|22|23360;JIS H750×250×14×28|750|250|14|28|23990;JIS H750×250×14×25|750|250|14|25|22580;JIS H750×250×12×25|750|250|12|25|21180;JIS H750×250×12×22|750|250|12|22|19750;JIS H750×250×12×19|750|250|12|19|18320;JIS H700×350×16×36|700|350|16|36|35530;JIS H700×350×16×32|700|350|16|32|32850;JIS H700×350×16×28|700|350|16|28|30180;JIS H700×350×16×25|700|350|16|25|28180;JIS H700×350×14×32|700|350|14|32|31580;JIS H700×350×14×28|700|350|14|28|28890;JIS H700×350×14×25|700|350|14|25|26880;JIS H700×350×12×25|700|350|12|25|25580;JIS H700×350×12×22|700|350|12|22|23550;JIS H700×300×16×32|700|300|16|32|29650;JIS H700×300×16×28|700|300|16|28|27380;JIS H700×300×16×25|700|300|16|25|25680;JIS H700×300×14×32|700|300|14|32|28380;JIS H700×300×14×28|700|300|14|28|26090;JIS H700×300×14×25|700|300|14|25|24380;JIS H700×300×12×25|700|300|12|25|23080;JIS H700×300×12×22|700|300|12|22|21350;JIS H700×300×12×19|700|300|12|19|19620;JIS H700×250×14×28|700|250|14|28|23290;JIS H700×250×14×25|700|250|14|25|21880;JIS H700×250×12×25|700|250|12|25|20580;JIS H700×250×12×22|700|250|12|22|19150;JIS H700×250×12×19|700|250|12|19|17720;JIS H700×250×9×19|700|250|9|19|15740;JIS H700×250×9×16|700|250|9|16|14290;JIS H700×200×12×28|700|200|12|28|19210;JIS H700×200×12×25|700|200|12|25|18080;JIS H700×200×12×22|700|200|12|22|16950;JIS H700×200×9×22|700|200|9|22|14980;JIS H700×200×9×19|700|200|9|19|13840;JIS H700×200×9×16|700|200|9|16|12690;JIS H700×200×9×12|700|200|9|12|11160;JIS H650×300×16×32|650|300|16|32|28720;JIS H650×300×16×28|650|300|16|28|26450;JIS H650×300×16×25|650|300|16|25|24750;JIS H650×300×16×22|650|300|16|22|23040;JIS H650×300×12×25|650|300|12|25|22350;JIS H650×300×12×22|650|300|12|22|20620;JIS H650×300×12×19|650|300|12|19|18890;JIS H650×300×12×16|650|300|12|16|17160;JIS H650×250×16×28|650|250|16|28|23650;JIS H650×250×12×28|650|250|12|28|21270;JIS H650×250×12×25|650|250|12|25|19850;JIS H650×250×12×22|650|250|12|22|18420;JIS H650×250×12×19|650|250|12|19|16990;JIS H650×200×12×28|650|200|12|28|18470;JIS H650×200×12×25|650|200|12|25|17350;JIS H650×200×12×22|650|200|12|22|16220;JIS H650×200×12×19|650|200|12|19|15090;JIS H650×200×9×22|650|200|9|22|14400;JIS H650×200×9×19|650|200|9|19|13250;JIS H650×200×9×16|650|200|9|16|12110;JIS H650×200×9×12|650|200|9|12|10580;JIS H600×300×16×32|600|300|16|32|27920;JIS H600×300×16×28|600|300|16|28|25650;JIS H600×300×12×28|600|300|12|28|23470;JIS H600×300×12×25|600|300|12|25|21750;JIS H600×300×12×22|600|300|12|22|20020;JIS H600×300×12×19|600|300|12|19|18290;JIS H600×250×16×32|600|250|16|32|24720;JIS H600×250×16×28|600|250|16|28|22850;JIS H600×250×12×28|600|250|12|28|20670;JIS H600×250×12×25|600|250|12|25|19250;JIS H600×250×12×22|600|250|12|22|17820;JIS H600×250×12×19|600|250|12|19|16390;JIS H600×250×9×19|600|250|9|19|14700;JIS H600×250×9×16|600|250|9|16|13260;JIS H600×200×12×28|600|200|12|28|17870;JIS H600×200×12×25|600|200|12|25|16750;JIS H600×200×12×22|600|200|12|22|15620;JIS H600×200×12×19|600|200|12|19|14490;JIS H600×200×9×22|600|200|9|22|13950;JIS H600×200×9×19|600|200|9|19|12800;JIS H600×200×9×16|600|200|9|16|11660;JIS H600×200×9×12|600|200|9|12|10130;JIS H550×300×16×28|550|300|16|28|24850;JIS H550×300×16×25|550|300|16|25|23150;JIS H550×300×16×22|550|300|16|22|21440;JIS H550×300×12×25|550|300|12|25|21150;JIS H550×300×12×22|550|300|12|22|19420;JIS H550×300×12×19|550|300|12|19|17690;JIS H550×300×12×16|550|300|12|16|15960;JIS H550×250×12×28|550|250|12|28|20070;JIS H550×250×12×25|550|250|12|25|18650;JIS H550×250×12×22|550|250|12|22|17220;JIS H550×250×9×22|550|250|9|22|15700;JIS H550×250×9×19|550|250|9|19|14250;JIS H550×250×9×16|550|250|9|16|12810;JIS H550×200×12×25|550|200|12|25|16150;JIS H550×200×12×22|550|200|12|22|15020;JIS H550×200×12×19|550|200|12|19|13890;JIS H550×200×9×22|550|200|9|22|13500;JIS H550×200×9×19|550|200|9|19|12350;JIS H550×200×9×16|550|200|9|16|11210;JIS H550×200×9×12|550|200|9|12|9679;JIS H500×300×16×28|500|300|16|28|24050;JIS H500×300×16×25|500|300|16|25|22350;JIS H500×300×16×22|500|300|16|22|20640;JIS H500×300×12×25|500|300|12|25|20550;JIS H500×300×12×22|500|300|12|22|18820;JIS H500×300×12×19|500|300|12|19|17090;JIS H500×300×12×16|500|300|12|16|15360;JIS H500×250×12×28|500|250|12|28|19470;JIS H500×250×12×25|500|250|12|25|18050;JIS H500×250×12×22|500|250|12|22|16620;JIS H500×250×9×22|500|250|9|22|15250;JIS H500×250×9×19|500|250|9|19|13800;JIS H500×250×9×16|500|250|9|16|12360;JIS H500×200×12×25|500|200|12|25|15550;JIS H500×200×12×22|500|200|12|22|14420;JIS H500×200×12×19|500|200|12|19|13290;JIS H500×200×9×22|500|200|9|22|13050;JIS H500×200×9×19|500|200|9|19|11900;JIS H500×200×9×16|500|200|9|16|10760;JIS H500×200×9×12|500|200|9|12|9229;JIS H450×250×12×28|450|250|12|28|18870;JIS H450×250×12×25|450|250|12|25|17450;JIS H450×250×12×22|450|250|12|22|16020;JIS H450×250×9×22|450|250|9|22|14800;JIS H450×250×9×19|450|250|9|19|13350;JIS H450×250×9×16|450|250|9|16|11910;JIS H450×200×12×25|450|200|12|25|14950;JIS H450×200×12×22|450|200|12|22|13820;JIS H450×200×12×19|450|200|12|19|12690;JIS H450×200×9×22|450|200|9|22|12600;JIS H450×200×9×19|450|200|9|19|11450;JIS H450×200×9×16|450|200|9|16|10310;JIS H450×200×9×12|450|200|9|12|8779;JIS H400×200×12×22|400|200|12|22|13220;JIS H400×200×9×22|400|200|9|22|12150;JIS H400×200×9×19|400|200|9|19|11000;JIS H400×200×9×16|400|200|9|16|9857;JIS H400×200×9×12|400|200|9|12|8329",
  "IN-ISJB": "ISJB 150|150|50|3|4.6|901;ISJB 175|175|50|3|4.6|1028;ISJB 200|200|60|3.4|5|1264;ISJB 225|225|80|3.7|5|1628",
  "IN-ISLB": "ISLB 75|75|50|3.7|5|771;ISLB 100|100|50|4|6.4|1021;ISLB 125|125|75|4.4|6.5|1512;ISLB 150|150|80|4.8|6.8|1808;ISLB 175|175|90|5.1|6.9|2130;ISLB 200|200|100|5.4|7.3|2527;ISLB 225|225|100|5.8|8.6|2992;ISLB 250|250|125|6.1|8.2|3553;ISLB 275|275|140|6.4|8.8|4202;ISLB 300|300|150|6.7|9.4|4808;ISLB 325|325|165|7|9.8|5490;ISLB 350|350|165|7.4|11.4|6301;ISLB 400|400|165|8|12.5|7243;ISLB 450|450|170|8.6|13.4|8314;ISLB 500|500|180|9.2|14.1|9550;ISLB 550|550|190|9.9|15|10997;ISLB 600|600|210|10.5|15.5|12607",
  "IN-ISMB": "ISMB 100|100|50|4.2|7|1140;ISMB 125|125|75|4.4|7.6|1660;ISMB 150|150|80|4.8|7.6|1900;ISMB 175|175|90|5.5|8.6|2462;ISMB 200|200|100|5.7|10.8|3233;ISMB 225|225|110|6.5|11.8|3972;ISMB 250|250|125|6.9|12.5|4755;ISMB 300|300|140|7.5|12.4|5626;ISMB 350|350|140|8.1|14.2|6670;ISMB 400|400|140|8.9|16|7840;ISMB 450|450|150|9.4|17.4|9227;ISMB 500|500|180|10.2|17.2|11074;ISMB 550|550|190|11.2|19.3|13211;ISMB 600|600|210|12|20.8|15621",
  "IN-ISWB": "ISWB 150|150|100|5.4|7|2167;ISWB 175|175|125|5.8|7.4|2811;ISWB 200|200|140|6.1|9|3671;ISWB 225|225|150|6.4|9.9|4324;ISWB 250|250|200|6.7|9|5205;ISWB 300|300|200|7.4|10|6133;ISWB 350|350|200|8|11.4|7250;ISWB 400|400|200|8.6|13|8501;ISWB 450|450|200|9.2|15.4|10115;ISWB 500|500|250|9.9|14.7|12122;ISWB 550|550|250|10.5|17.6|14334;ISWB 600-1|600|250|11.2|21.3|17038;ISWB 600-2|600|250|11.8|23.6|18486",
  "IN-ISHB": "ISHB 150-1|150|150|5.4|9|3448;ISHB 150-2|150|150|8.4|9|3898;ISHB 150-3|150|150|11.8|9|4408;ISHB 200-1|200|200|6.1|9|4754;ISHB 200-2|200|200|7.8|9|5094;ISHB 225-1|225|225|6.5|9.1|5494;ISHB 225-2|225|225|8.6|9.1|5966;ISHB 250-1|250|250|6.9|9.7|6496;ISHB 250-2|250|250|8.8|9.7|6971;ISHB 300-1|300|250|7.6|10.6|7485;ISHB 300-2|300|250|9.4|10.6|8025;ISHB 350-1|350|250|8.3|11.6|8591;ISHB 350-2|350|250|10.1|11.6|9221;ISHB 400-1|400|250|9.1|12.7|9866;ISHB 400-2|400|250|10.6|12.7|10466;ISHB 450-1|450|250|9.8|13.7|11114;ISHB 450-2|450|250|11.3|13.7|11789",
  "IN-ISJC": "ISJC 100|100|45|3|5.1|741;ISJC 125|125|50|3|6.6|1007;ISJC 150|150|55|3.6|6.9|1265;ISJC 175|175|60|3.6|6.9|1424;ISJC 200|200|70|4.1|7.1|1780",
  "IN-ISLC": "ISLC 75|75|40|3.7|6|726;ISLC 100|100|50|4|6.4|1002;ISLC 125|125|65|4.4|6.6|1367;ISLC 150|150|75|4.8|7.8|1836;ISLC 175|175|75|5.1|9.5|2240;ISLC 200|200|75|5.5|10.8|2622;ISLC 225|225|90|5.8|10.2|3053;ISLC 250|250|100|6.1|10.7|3565;ISLC 300|300|100|6.7|11.6|4211;ISLC 350|350|100|7.4|12.5|4947;ISLC 400|400|100|8|14|5825",
  "IN-ISMC": "ISMC 75|75|40|4.4|7.3|867;ISMC 100|100|50|4.7|7.5|1170;ISMC 125|125|65|5|8.1|1619;ISMC 150|150|75|5.4|9|2088;ISMC 175|175|75|5.7|10.2|2438;ISMC 200|200|75|6.1|11.4|2821;ISMC 225|225|80|6.4|12.4|3301;ISMC 250|250|80|7.1|14.1|3867;ISMC 300|300|90|7.6|13.6|4564;ISMC 350|350|100|8.1|13.5|5366;ISMC 400|400|100|8.6|15.3|6293",
  "IN-ISNT": "ISNT 20|20|20|3|3|113;ISNT 30|30|30|3|3|175;ISNT 40|40|40|6|6|448;ISNT 50|50|50|6|6|570;ISNT 60|60|60|6|6|690;ISNT 80|80|80|8|8|1225;ISNT 100|100|100|10|10|1910;ISNT 150|150|150|10|10|2908",
  "IN-ISHT": "ISHT 75|75|150|8.4|9|1949;ISHT 100|100|200|7.8|9|2547;ISHT 125|125|250|8.8|9.7|3485;ISHT 150|150|250|7.6|10.6|3742",
  "IN-ISST": "ISST 100|100|50|5.8|10|1037;ISST 150|150|75|8|11.6|1996;ISST 200|200|165|8|12.5|3622;ISST 250|250|180|9.2|14.1|4775",
  "IN-ISLT": "ISLT 50|50|50|4|6.4|511;ISLT 75|75|80|4.8|6.8|904;ISLT 100|100|100|5.7|10.8|1616",
  "IN-ISJT": "ISJT 75|75|50|3|4.6|450;ISJT 57.5|87.5|50|3.2|4.8|514;ISJT 100|100|60|3.4|5|632;ISJT 112.5|112.5|80|3.7|5|814",
  "IN-ISA": "20|20|3|112;20|20|4|145;25|25|3|141;25|25|4|184;25|25|5|225;30|30|3|173;30|30|4|226;30|30|5|277;35|35|3|203;35|35|4|266;35|35|5|327;35|35|6|386;40|40|3|234;40|40|4|307;40|40|5|378;40|40|6|447;45|45|3|264;45|45|4|347;45|45|5|428;45|45|6|507;50|50|3|295;50|50|4|388;50|50|5|479;50|50|6|568;55|55|5|527;55|55|6|626;55|55|8|818;55|55|10|1002;60|60|5|575;60|60|6|684;60|60|8|896;60|60|10|1100;65|65|6|744;65|65|8|976;65|65|10|1200;70|70|5|677;70|70|6|806;70|70|8|1058;70|70|10|1302;75|75|5|727;75|75|6|866;75|75|8|1138;75|75|10|1402;80|80|6|929;80|80|8|1221;80|80|10|1505;80|80|12|1781;90|90|6|1047;90|90|8|1379;90|90|10|1703;90|90|12|2019;100|100|6|1167;100|100|8|1539;100|100|10|1903;100|100|12|2259;110|110|8|1702;110|110|10|2106;110|110|12|2502;110|110|15|3081;130|130|8|2022;130|130|10|2506;130|130|12|2982;130|130|15|3681;150|150|10|2903;150|150|12|3459;150|150|15|4278;150|150|18|5079;200|200|12|4661;200|200|15|5780;200|200|18|6881;200|200|25|9380;30|20|3|141;30|20|4|184;30|20|5|225;40|25|3|188;40|25|4|246;40|25|5|302;40|25|6|356;45|30|3|218;45|30|4|286;45|30|5|352;45|30|6|416;50|30|3|234;50|30|4|307;50|30|5|378;50|30|6|447;60|40|5|476;60|40|6|565;60|40|8|737;65|45|5|526;65|45|6|625;65|45|8|817;70|45|5|552;70|45|6|656;70|45|8|858;70|45|10|1052;75|50|5|602;75|50|6|716;75|50|8|938;75|50|10|1152;80|50|5|627;80|50|6|746;80|50|8|978;80|50|10|1202;90|60|6|865;90|60|8|1137;90|60|10|1401;90|60|12|1657;100|65|6|955;100|65|8|1257;100|65|10|1551;100|75|6|1014;100|75|8|1336;100|75|10|1650;100|75|12|1956;125|75|6|1166;125|75|8|1538;125|75|10|1902;125|95|6|1286;125|95|8|1698;125|95|10|2102;125|95|12|2498;150|75|8|1742;150|75|10|2156;150|75|12|2562;150|115|8|2058;150|115|10|2552;150|115|12|3038;150|115|15|3752;200|100|10|2903;200|100|12|3459;200|100|15|4278;200|150|10|3400;200|150|12|4056;200|150|15|5025;200|150|18|5976",
  "IN-ISB": "25|25|2.6|216;25|25|3.2|253;38|38|2.6|349;38|38|3.2|416;38|38|4|503;49.5|49.5|2.6|470;49.5|49.5|2.9|519;49.5|49.5|3.6|628;49.5|49.5|4.5|758;72|72|3.2|834;72|72|4|1047;72|72|4.8|1231;91.5|91.5|3.6|1232;91.5|91.5|4.5|1514;91.5|91.5|5.4|1785;113.5|113.5|4.8|2028;113.5|113.5|5.4|2260;132|132|4.8|2383;132|132|5.4|2659;50|25|2.6|349;50|25|3.2|416;66|33|2.6|470;66|33|2.9|519;66|33|3.6|628;66|33|4.5|758;96|48|3.2|854;96|48|4|1047;96|48|4.8|1231;122|61|3.6|1232;122|61|4.5|1514;122|61|5.4|1785;145|82|4.8|2028;145|82|5.4|2260;172|92|4.8|2383;172|92|5.4|2659",
  "IN-ISNB": "ISNB 15L|21.3|2|121;ISNB 15M|21.3|2.6|153;ISNB 15H|21.3|3.2|182;ISNB 20L|26.9|2.3|178;ISNB 20M|26.9|2.6|198;ISNB 20H|26.9|3.2|238;ISNB 25L|33.7|2.6|254;ISNB 25M|33.7|3.2|306;ISNB 25H|33.7|4|373;ISNB 32L|42.4|2.6|325;ISNB 32M|42.4|3.2|394;ISNB 32H|42.4|4|482;ISNB 40L|48.3|2.9|413;ISNB 40M|48.3|3.2|453;ISNB 40H|48.3|4|556;ISNB 50L|60.3|2.9|523;ISNB 50M|60.3|3.6|641;ISNB 50H|60.3|4.5|788;ISNB 65L|76.1|3.2|732;ISNB 65M|76.1|3.6|820;ISNB 65H|76.1|4.5|1010;ISNB 80L|88.9|3.2|861;ISNB 80M|88.9|4|1070;ISNB 80H|88.9|4.8|1270;ISNB 90L|101.6|3.6|1110;ISNB 90M|101.6|4|1230;ISNB 90H|101.6|4.8|1460;ISNB 100L|114.3|3.6|1250;ISNB 100M|114.3|4.5|1550;ISNB 100H|114.3|5.4|1850;ISNB 110L|127|4.5|1730;ISNB 110M|127|4.8|1840;ISNB 110H|127|5.4|2060;ISNB 125L|139.7|4.5|1910;ISNB 125M|139.7|4.8|2030;ISNB 125H|139.7|5.4|2280;ISNB 135L|152.4|4.5|2090;ISNB 135M|152.4|4.8|2220;ISNB 135H|152.4|5.4|2500;ISNB 150L|165.1|4.5|2270;ISNB 150M|165.1|4.8|2420;ISNB 150H|165.1|5.4|2710;ISNB 160L|168.3|4.5|2310;ISNB 160M|168.3|4.8|2470;ISNB 160H1|168.3|5.4|2760;ISNB 160H2|168.3|6.3|3200;ISNB 175L|193.7|4.8|2850;ISNB 175M|193.7|5.4|3200;ISNB 175H|193.7|5.9|3480;ISNB 200L|219.1|4.8|3230;ISNB 200M|219.1|5.6|3750;ISNB 200H|219.1|5.9|3950;ISNB 225H|244.5|5.9|4420;ISNB 250H|273|5.9|4950;ISNB 300H|323.9|6.3|6280;ISNB 350H|355.6|8|8730"
};

let _INTL = null;
function intlKey(s) {
  return String(s == null ? "" : s).replace(/№/g, " ").normalize("NFKC").toUpperCase()
    .replace(/[×✕✖*∗XХ]/g, "X")
    .replace(/(\d),(\d)/g, "$1.$2")
    .replace(/[ΦФ⌀]/g, "Ø").replace(/口/g, "□").replace(/∟/g, "∠")
    .replace(/[\s_\-–—‐‑−()]/g, "");
}
function intlData() {
  if (_INTL) return _INTL;
  const rows = [], byKey = new Map(), ruShort = new Map(), meta = {};
  const add = (k, r) => { if (!k) return; const a = byKey.get(k); if (!a) byKey.set(k, [r]); else if (a.indexOf(r) < 0) a.push(r); };
  for (const [key, cc, shape, std, code, lab, tpl, dec, src] of INTL_FAMS) {
    meta[key] = { key, cc, shape, std, code, lab, src, n: 0 };
    const pack = INTL_PACK[key] || "";
    for (const line of pack.split(";")) {
      if (!line) continue;
      const f = line.split("|");
      let name, d = {};
      if (tpl) {
        const v = f.slice(0, -1), ph = shape === "L" ? ["a", "b", "t"] : shape === "B" ? ["h", "b", "t"] : ["D", "t"];
        name = tpl; ph.forEach((p, i) => { d[p] = +v[i]; name = name.replace("{" + p + "}", dec === "," ? v[i].replace(".", ",") : v[i]); });
      } else if (shape === "P") { name = f[0]; d = { D: +f[1], t: +f[2] }; }
      else { name = f[0]; d = { h: +f[1], b: +f[2], tw: +f[3], tf: +f[4] }; }
      const A = +f[f.length - 1];
      const pm = shape === "I" || shape === "C" ? 2 * d.h + 4 * d.b : shape === "T" ? 2 * (d.h + d.b)
        : shape === "L" ? 2 * (d.a + d.b) : shape === "B" ? 2 * (d.h + d.b) : Math.PI * d.D;
      const basis = shape === "I" || shape === "C" ? "ih" : shape === "T" ? "tee" : shape === "L" ? "angle" : shape === "B" ? "box" : "round";
      const r = { name, type: key, kgm: Math.round(A * 0.785) / 100, intl: true, cc, shape, dims: d, A, pm, basis, std, src, key: intlKey(name) };
      rows.push(r); meta[key].n++;
      add(r.key, r);
      if (key === "RU-8239") ruShort.set(intlKey(name.replace(/\s.*$/, "")), r);  // bare "I20" = ГОСТ 8239-89 only in a Russian interface
      else if (key === "CN-I") add("GB" + r.key, r);                                  // SAP2000 spelling GB-I20a
      else if (key === "CN-C") add("GBC" + r.key.slice(1), r);                        // GB-C20a
      else if (key === "CN-LN") add("LN" + r.key.slice(1), r);                        // Ln100X63X8
      else if (cc === "CN" && shape === "P") add(intlKey(std + d.D + "X" + d.t), r);   // GB-SSP108X4
    }
  }
  _INTL = { rows, byKey, ruShort, meta };
  return _INTL;
}
function intlPick(list) {
  if (!list || !list.length) return null;
  if (list.length === 1) return list[0];
  const cc = { ru: "RU", zh: "CN" }[SO_LANG];
  const pref = cc ? list.filter(r => r.cc === cc) : [];
  return pref.length === 1 ? pref[0] : null;                    // still ambiguous → not guessed
}
function intlByName(n) {
  const D = intlData(), k = intlKey(n);
  if (SO_LANG === "ru" && D.ruShort.has(k)) return D.ruShort.get(k);
  return intlPick(D.byKey.get(k));
}

/* ── native designations → a library row (or a name the old search knows) ─── */
const INTL_CYR = /[А-Яа-яЁё]/, INTL_CJK = /[㐀-鿿豈-﫿]/;
const INTL_NUM = "(\\d+(?:\\.\\d+)?)";
function intlDims(s) {                                          // "100x100x5" → [100,100,5]
  const m = s.match(new RegExp(`${INTL_NUM}(?:\\s*x\\s*${INTL_NUM})(?:\\s*x\\s*${INTL_NUM})?`));
  return m ? m.slice(1).filter(v => v != null).map(Number) : [];
}
function intlPrep(raw) {                                         // unify ×/х/*, decimal commas, spaces
  return String(raw).replace(/№/g, " ").normalize("NFKC").toUpperCase()
    .replace(/(\d)\s*[×✕✖*∗XХ]\s*(?=[\dØΦФ])/g, "$1x")
    .replace(/(\d),(\d)/g, "$1.$2").replace(/\s+/g, " ").trim();
}
function ruStd(s) {
  if (/57837/.test(s)) return "57837";
  if (/26020/.test(s)) return "26020";
  if (/АСЧМ|ASCHM|СТО\s*20-93/.test(s)) return "STO";
  if (/8239/.test(s)) return "8239";
  if (/8240/.test(s)) return "8240";
  if (/19425/.test(s)) return "19425";
  if (/5267/.test(s)) return "5267";
  if (/8509/.test(s)) return "8509";
  if (/8510/.test(s)) return "8510";
  if (/32931/.test(s)) return "32931";
  if (/30245/.test(s)) return "30245";
  if (/54157/.test(s)) return "54157";
  if (/2287/.test(s)) return "TU80";
  if (/ГОСТ|GOST|ТУ\s*\d|СТО/.test(s)) return "other";
  return null;
}
const RU_TUBE_STD = { "32931": "ГОСТ 32931-2015", "30245": "ГОСТ 30245-2003", "54157": "ГОСТ Р 54157-2010", "TU80": "ТУ 36-2287-80" };
const RU_LAT = { B: "Б", SH: "Ш", W: "Ш", K: "К", S: "С", DB: "ДБ", DK: "ДК", DSH: "ДШ", M: "М" };
function ruSection(raw, latinOk) {
  const s = intlPrep(raw);
  const std = ruStd(s);
  const body = s.replace(/(?:ГОСТ|GOST|СТО|STO|ТУ|TU)\s*[РP]?\s*[\d.\s-]+(?:-\d{2,4})?/g, " ").replace(/АСЧМ\s*20-93|АСЧМ/g, " ").replace(/\s+/g, " ").trim();
  let m;
  if (/^(?:ЛИСТ|ПОЛОСА|ПЛАСТИНА|ФАСОНКА|РЕБРО|НАКЛАДКА|ПРОКЛАДКА|ОПОРНАЯ ПЛИТА|ПЛИТА)/.test(body) || /^[-—–]\s*\d/.test(body)) return { plate: true };
  if (/^(?:КРУГ|ПРУТ|АРМАТУРА|КВАДРАТ\s+\d+$)/.test(body)) return null;
  const tube = /(?:ТРУБ|ТР\.|ПРОФТРУБ|ЗАМКНУТ|^□)/.test(body), sq = /(?:КВ\.?|КВАДРАТ|^□)/.test(body), rnd = /(?:^Ø|ØD|КРУГЛ|ЭЛЕКТРОСВАРН|БЕСШОВН)/.test(body);
  const angle = /(?:УГОЛ|УГ\.|^L\s*\d|^∟|^∠)/.test(body), chan = /(?:ШВЕЛЛЕР|ШВ\.|^\[)/.test(body);
  const beam = /(?:ДВУТАВР|БАЛКА|ДВ\.|^I\s*\d)/.test(body), tee = /(?:ТАВР|^Т\s*\d)/.test(body);
  const d = intlDims(body);
  if (tube || (sq && d.length)) {
    if (!d.length) return null;
    let h, b, t, round = false;
    if (d.length === 3) [h, b, t] = d;
    else if (d.length === 2 && (sq || /ПРОФИЛ/.test(body)) && !rnd) { h = b = d[0]; t = d[1]; }
    else if (d.length === 2) { round = true; [h, t] = d; }
    if (!(t > 0 && t < Math.min(h, b || h) / 2)) return null;
    const stdName = RU_TUBE_STD[std] || (std === "other" ? null : "ГОСТ 32931-2015");
    if (round) {
      const row = stdName && std !== null ? intlByName(`Ø${h}x${t} ${stdName}`) : null;
      return row ? { row } : { name: `CHS ${h}x${t}` };
    }
    const row = stdName ? intlByName(`□${h}x${b}x${t} ${stdName}`) : null;
    return row ? { row } : { name: `${h === b ? "SHS" : "RHS"} ${h}x${b}x${t}` };
  }
  if (angle) {
    if (d.length < 2) return null;
    const [a, b, t] = d.length === 3 ? d : [d[0], d[0], d[1]];
    if (!(t > 0 && t < Math.min(a, b) / 2)) return null;
    const row = a === b ? (std === null || std === "8509" ? intlByName(`L${a}x${t} ГОСТ 8509-93`) : null)
      : (std === null || std === "8510" ? intlByName(`L${a}x${b}x${t} ГОСТ 8510-86`) : null);
    return row ? { row } : { name: `L ${a}x${b}x${t}` };
  }
  if (tee && (m = body.match(/(\d+(?:\.\d+)?)\s*(КТ|ШТ|БТ)\s*(\d+)/))) { const row = intlByName(m[1] + m[2] + m[3]); return row ? { row } : null; }
  if (chan) {
    if ((m = body.match(/(\d+(?:\.\d+)?)\s*([ПУЭЛСВ])\s*([АБ]|-\d)?(?![А-Я])/))) {
      const des = m[1] + m[2] + (m[3] || "").replace(/^([АБ])$/, c => c.toLowerCase());
      if (std === "19425") { const row = intlByName(`[${des} ГОСТ 19425-74`); return row ? { row } : null; }
      if (std === "5267" || m[2] === "В") { const row = intlByName(`${des} ГОСТ 5267.1-90`); return row ? { row } : null; }
      const row = std === null || std === "8240" ? intlByName(des) : null;
      return row ? { row } : null;
    }
    if ((m = body.match(/(?:^|[^\d.])(\d+(?:\.\d+)?)\s*([АA])?(?:\s|$)/))) return { name: `GOST C${m[1]}${m[2] ? "a" : ""}` };   // 8240-72 style «Швеллер 20», «14а»
    return null;
  }
  // I-beams: 20Б1 · 30Ш2 · 40К1 · 25С1 · 20ДБ1 · 10ДК1 · 24М · КХБ-515 · 20Д1А
  const des = (m = body.match(/(\d+) ?(ДБ|ДК|ДШ|Б|Ш|К|С|М) ?(\d{0,2})(?![\dА-Я])/)) ? m[1] + m[2] + m[3]
    : (m = body.match(/(\d+)\s*(Д|К)\s*(\d)\s*А/)) ? m[1] + m[2] + m[3] + "А"
    : (m = body.match(/КХБ\s*-?\s*(\d+)/)) ? "КХБ-" + m[1]
    : latinOk && (m = body.match(/^(?:I\s*)?(\d+)\s*(DB|DK|DSH|SH|B|W|K|S|M)\s*(\d*)$/)) ? m[1] + RU_LAT[m[2]] + m[3] : null;
  if (des) {
    const suf = std === "26020" ? " ГОСТ 26020-83" : std === "STO" ? " СТО АСЧМ 20-93" : std === "19425" ? " ГОСТ 19425-74" : "";
    if (std === "19425") { const row = intlByName(`I${des}${suf}`); return row ? { row } : null; }
    let row = intlByName(des + suf);
    if (!row && /\dА$/.test(des)) row = intlByName(des + " R40-93");
    if (!row && std === null) row = intlByName(des + " СТО АСЧМ 20-93") || intlByName(des + " ГОСТ 26020-83");
    return row ? { row } : null;
  }
  if (beam && (m = body.match(/(?:^|[^\d.])(\d+)(?:\s|$)/)) && (std === null || std === "8239")) {
    const row = intlByName(`I${m[1]} ГОСТ 8239-89`); return row ? { row } : null;
  }
  return null;
}
function zhSection(raw) {
  const s = intlPrep(raw).replace(/^Ⅰ/, "I").replace(/口/g, "□");
  let m;
  if (/^(?:钢板|板|扁钢|花纹板|钢带|加劲板|连接板|节点板|垫板)/.test(s) || /^[-—－]\s*\d/.test(s) || /δ\s*=?\s*\d/.test(s)) return { plate: true };
  if (/^(?:圆钢|钢筋|螺纹钢|方钢\s*\d+$)/.test(s)) return null;
  const body = s.replace(/^(?:热轧|焊接|高频焊接|冷弯)?(?:H型钢|H形钢|工字钢|槽钢|等边角钢|不等边角钢|角钢|T型钢|剖分T型钢|无缝钢管|焊接钢管|直缝焊管|螺旋焊管|钢管|圆管|方管|方钢管|矩形管|矩形钢管)\s*/, "").trim();
  const kind = /H型钢|H形钢/.test(s) ? "H" : /工字钢/.test(s) ? "I" : /槽钢/.test(s) ? "C" : /角钢/.test(s) ? "L" : /T型钢/.test(s) ? "T"
    : /钢管|圆管|焊管/.test(s) ? "P" : /方管|方钢管|矩形管|矩形钢管/.test(s) ? "B" : null;
  let row = intlByName(body);
  if (row) return { row };
  const d = intlDims(body);
  if (kind === "B" || /^□/.test(body)) {
    const [h, b, t] = d.length === 3 ? d : d.length === 2 ? [d[0], d[0], d[1]] : [];
    return t > 0 && t < Math.min(h, b) / 2 ? { name: `${h === b ? "SHS" : "RHS"} ${h}x${b}x${t}` } : null;
  }
  if (kind === "P" || /^[ØΦФD]\s*\d/.test(body)) return d.length === 2 && d[1] < d[0] / 2 ? { name: `CHS ${d[0]}x${d[1]}` } : null;
  if (kind === "L" || /^[∠L]\s*\d/.test(body)) {
    if (d.length < 2) return null;
    const [a, b, t] = d.length === 3 ? d : [d[0], d[0], d[1]];
    if (!(t > 0 && t < Math.min(a, b) / 2)) return null;
    row = intlByName(a === b ? `∠${a}x${t}` : `∠${a}x${b}x${t}`);
    return row ? { row } : { name: `L ${a}x${b}x${t}` };
  }
  if ((m = body.match(/^(?:I|\[|C)?\s*(\d+(?:\.\d+)?)\s*([ABC])?$/)) && (kind === "I" || kind === "C" || /^[I\[]/.test(body))) {
    const isC = kind === "C" || /^\[/.test(body) || (/^C/.test(body) && kind !== "I");
    row = intlByName((isC ? "[" : "I") + m[1] + (m[2] || "").toLowerCase());
    return row ? { row } : null;
  }
  if (kind === "H" && d.length >= 2) return { name: `H${d.join("x")}` };
  return null;
}
const ES_WORDS = /\b(?:PERFIL(?:ES)?|VIGA|VIGUETA|TUBO|TUBULAR|ANGULAR|[AÁ]NGULO|CANAL|CUADRADO|RECTANGULAR|REDONDO|CIRCULAR|ESTRUCTURAL|LAMINADO|CHAPA|PLACA|PLANCHA|L[AÁ]MINA|PLETINA|LLANTA|FLEJE|REDONDO|BARRA)\b/;
function esSection(raw, esCtx) {
  const s = intlPrep(raw).replace(/[ÁÀ]/g, "A").replace(/[ÉÈ]/g, "E").replace(/[ÍÌ]/g, "I").replace(/[ÓÒ]/g, "O").replace(/[ÚÙ]/g, "U");
  const hasWord = ES_WORDS.test(s.replace(/Á/g, "A"));
  let m;
  if (!hasWord) {
    if (esCtx && (m = s.match(/^L\s*-?\s*(\d{2,3})\.(\d{1,2})$/)) && +m[2] <= 25 && +m[1] >= 20) return { name: `L ${m[1]}x${m[1]}x${m[2]}` };   // Spanish «L 50.5» = 50×50×5
    return null;
  }
  if (/^(?:CHAPA|PLACA|PLANCHA|LAMINA|PLETINA|LLANTA|FLEJE)\b/.test(s)) return { plate: true };
  if (/^(?:REDONDO|BARRA)\b/.test(s) && !/TUBO/.test(s)) return null;
  const d = intlDims(s);
  if (/TUBO|TUBULAR/.test(s)) {
    const round = /REDONDO|CIRCULAR|Ø/.test(s), sq = /CUADRADO/.test(s);
    if (round || (d.length === 2 && !sq)) return d.length >= 2 && d[1] < d[0] / 2 ? { name: `CHS ${d[0]}x${d[1]}` } : null;
    const [h, b, t] = d.length === 3 ? d : d.length === 2 ? [d[0], d[0], d[1]] : [];
    return t > 0 && t < Math.min(h, b) / 2 ? { name: `${h === b ? "SHS" : "RHS"} ${h}x${b}x${t}` } : null;
  }
  if (/ANGULAR|ANGULO/.test(s)) {
    if ((m = s.match(/(\d{2,3})\.(\d{1,2})$/)) && d.length < 2 && +m[2] <= 25) return { name: `L ${m[1]}x${m[1]}x${m[2]}` };
    if (d.length < 2) return null;
    const [a, b, t] = d.length === 3 ? d : [d[0], d[0], d[1]];
    return t > 0 && t < Math.min(a, b) / 2 ? { name: `L ${a}x${b}x${t}` } : null;
  }
  const rest = s.replace(/\b(?:PERFIL(?:ES)?|VIGA|VIGUETA|CANAL|ESTRUCTURAL|LAMINADO|EN CALIENTE|DE ACERO|ACERO|TIPO)\b/g, " ").replace(/\s+/g, " ").trim();
  return rest && rest !== s ? { name: rest } : null;
}
/* the section name in any supported language → row / old-library name / plate marker / null */
function intlSectionName(raw, ctx) {
  const s = String(raw == null ? "" : raw);
  if (!s.trim() || !/[A-Za-zА-Яа-яЁё㐀-鿿\[□口∠∟ØΦФ]/.test(s)) return null;
  const ruLatin = SO_LANG === "ru" && /^\s*(?:I\s*)?\d+\s*(?:DB|DK|DSH|SH|B|W|K|S|M)\s*\d{0,2}\s*$/i.test(s);
  if (INTL_CYR.test(s) || ctx === "ru" || ruLatin) return ruSection(s, ctx === "ru" || SO_LANG === "ru");
  if (INTL_CJK.test(s) || /[∠口□Φ]/.test(s) || /^\s*[\[Ⅰ]/.test(s) || ctx === "zh") return zhSection(s);
  return esSection(s, ctx === "es" || SO_LANG === "es");
}
const _intlMemo = new Map();
let _intlDepth = 0;
function intlFind(raw) {
  if (!INTL_SECTIONS_ENABLED || raw == null || _intlDepth > 1) return null;
  const s = String(raw).trim();
  if (!s || !/[^\d\s.,x×*\-+/]/.test(s)) return null;           // pure numbers are never sections
  const mk = SO_LANG + "\u0001" + s;
  if (_intlMemo.has(mk)) return _intlMemo.get(mk);
  let out = null;
  try {
    out = intlByName(s);
    if (!out) {
      const c = intlSectionName(s);
      if (c && c.row) out = c.row;
      else if (c && c.name && c.name.toUpperCase() !== s.toUpperCase()) { _intlDepth++; try { out = findSection(c.name); } finally { _intlDepth--; } }
    }
  } catch (e) { try { console.warn("[intl-sections add-on]", e); } catch { /* no console */ } out = null; }
  if (_intlMemo.size > 5000) _intlMemo.clear();
  _intlMemo.set(mk, out);
  return out;
}
// Extra autocomplete rows — only for native-script queries, IS/JIS/GB/YB/JG-style
// prefixes, or when the interface is Russian, Chinese or Spanish.
function intlSearch(q, n) {
  if (!INTL_SECTIONS_ENABLED || !(n > 0) || !q) return [];
  const ok = /[^\x00-\x7F]/.test(q) || /^\s*(?:IS[A-Z]|JIS|GB-|YB-|JG-|LH\d|HT\d|T[WMN]\d)/i.test(q) || SO_LANG === "ru" || SO_LANG === "zh" || SO_LANG === "es";
  if (!ok) return [];
  const K = intlKey(q); if (!K) return [];
  const starts = [], contains = [];
  for (const r of intlData().rows) { if (r.key.startsWith(K)) { starts.push(r); if (starts.length >= n) break; } else if (contains.length < n && r.key.includes(K)) contains.push(r); }
  return [...starts, ...contains].slice(0, n);
}
function intlFamilies() {
  if (!INTL_SECTIONS_ENABLED) return [];
  const { meta } = intlData();
  return INTL_FAMS.map(f => meta[f[0]]).filter(m => m && m.n > 0);
}
function intlFamRows(key) { return intlData().rows.filter(r => r.type === key).sort((x, y) => x.kgm - y.kgm); }
/* ╚══ end of INTERNATIONAL SECTIONS add-on ═══════════════════════════════════╝ */

/* ─── L-SHAPE OFFCUT HELPERS ─────────────────────────────────────────────
   An L-offcut = full rectangle A (width) × B (length) with a corner notch
   C (width) × D (length) removed. The largest axis-aligned rectangle that
   fits inside the L is the better of the two leftover rectangles:
     • (A − C) × B   (full-height strip beside the notch)
     • A × (B − D)   (full-width strip below the notch)
   We nest only into that usable rectangle — never claims metal that isn't there. */
function lUsableRect(A, B, C, D) {
  const r1 = { w: A - C, h: B };       // beside the notch
  const r2 = { w: A, h: B - D };       // below the notch
  return (r1.w * r1.h >= r2.w * r2.h) ? r1 : r2;
}

/* Pack parts (with qty) into ONE usable rectangle. Returns {placements, leftover}.
   leftover = the same parts list with reduced quantities for whatever didn't fit. */
function nestIntoRect(rectW, rectH, parts, kerf, allowRotation) {
  // expand
  let pieces = [];
  parts.forEach(p => { for (let q = 0; q < p.qty; q++) pieces.push({ w: p.width, h: p.length, id: p.id, label: p.label }); });
  const run = packOneSheetBest(rectW, rectH, pieces, kerf, allowRotation);
  const placements = run.placed.map(pl => ({ x: pl.x, y: pl.y, w: pl.w, h: pl.h, rotated: pl.rotated, id: pl.src.id, label: pl.src.label }));
  const placedCount = {};
  placements.forEach(pl => { placedCount[pl.id] = (placedCount[pl.id] || 0) + 1; });
  const notPlaced = run.left.slice();
  const leftover = parts.map(p => ({ ...p, qty: p.qty - (placedCount[p.id] || 0) })).filter(p => p.qty > 0);
  return { placements, leftover, placedAny: placements.length > 0 };
}

/* Find reusable leftover on a sheet by scanning the FREE area. Uses a y-band
   sweep: for every horizontal band bounded by part edges, find the empty x-gaps.
   This is guaranteed never to overlap a placed part. Reports the biggest empty
   rectangle, or an L when there's a clean right + bottom corner region. */
function findOffcuts(placements, m, usableW, usableH, reuseMin) {
  const L = m, R = m + usableW, T = m, B = m + usableH;
  const yTops = [...new Set([T, ...placements.map(p => p.y + p.h)])].filter(y => y >= T && y < B).sort((a, b) => a - b);
  const yBots = [...new Set([B, ...placements.map(p => p.y)])].filter(y => y > T && y <= B).sort((a, b) => a - b);
  let best = null;
  for (const y1 of yTops) for (const y2 of yBots) {
    const h = y2 - y1; if (h < reuseMin) continue;
    // x-intervals blocked by any part overlapping this y-band
    const blockers = placements
      .filter(p => p.y < y2 - 0.5 && p.y + p.h > y1 + 0.5)
      .map(p => [Math.max(L, p.x), Math.min(R, p.x + p.w)])
      .filter(([a, b]) => b > a)
      .sort((a, b) => a[0] - b[0]);
    let cursor = L;
    for (const [bx0, bx1] of blockers) {
      if (bx0 - cursor >= reuseMin) { const w = bx0 - cursor, area = w * h; if (!best || area > best.area) best = { x: cursor, y: y1, w, h, area }; }
      cursor = Math.max(cursor, bx1);
    }
    if (R - cursor >= reuseMin) { const w = R - cursor, area = w * h; if (!best || area > best.area) best = { x: cursor, y: y1, w, h, area }; }
  }
  // L-shape = sheet minus the top-left used bounding box (both strips guaranteed empty).
  const usedRight = placements.length ? Math.max(...placements.map(p => p.x + p.w)) : L;
  const usedBottom = placements.length ? Math.max(...placements.map(p => p.y + p.h)) : T;
  const rightW = R - usedRight, bottomH = B - usedBottom, notchW = usedRight - L, notchH = usedBottom - T;
  if (rightW >= reuseMin && bottomH >= reuseMin && notchW > 0 && notchH > 0) {
    const Larea = usableW * usableH - notchW * notchH;
    const r1 = { w: rightW, h: usableH }, r2 = { w: usableW, h: bottomH };
    const usable = (r1.w * r1.h >= r2.w * r2.h) ? r1 : r2;
    if (!best || Larea >= best.area) return [{ shape: "L", x: L, y: T, A: Math.round(usableW), B: Math.round(usableH), notchW: Math.round(notchW), notchH: Math.round(notchH), w: Math.round(usable.w), h: Math.round(usable.h) }];
  }
  if (best) return [{ shape: "rect", x: best.x, y: best.y, A: Math.round(best.w), B: Math.round(best.h), w: Math.round(best.w), h: Math.round(best.h) }];
  return [];
}


/* ════════════════════════════════════════════════════════════════════════
   OPTIMISATION CORE v2 — waste minimisation
   ------------------------------------------------------------------------
   1D (bars):  Sequential Heuristic Procedure (Gradisar et al.) — repeatedly
               cut the single best pattern by solving a bounded knapsack
               exactly, so each bar is filled as fully as remaining demand
               allows instead of committing piece by piece. The knapsack is
               the same structure Gilmore & Gomory (Oper. Res. 9:849, 1961)
               use as their pricing subproblem, but this is NOT true column
               generation: there is no LP master and no dual prices, so it
               carries no LP optimality guarantee. Followed by bar-elimination
               local search. Bounds below say how close it actually got.
   2D (plates): every candidate is GUILLOTINE-FEASIBLE (edge-to-edge cuts
               only) because Gulf shops cut on shears and band saws, not
               contour CNC. MaxRects would pack tighter but is not cuttable
               on that equipment, so it is deliberately excluded.
   Selection:  primary = fewest stock items; tie-break = largest single
               reusable offcut, since scattered scrap has no resale value.
   Guarantee:  the previously shipped heuristic is always run as one of the
               candidates, so v2 can never return a worse answer than v1.
════════════════════════════════════════════════════════════════════════ */

/* ── proven lower bounds — used for the honesty readout, not for packing ──
   A lower bound says "no algorithm can beat this". When the layout equals the
   bound the job is provably optimal and no further search is worthwhile
   (Korf, AAAI-02). Dual Feasible Functions (Fekete & Schepers, Math. Prog.
   91:11–31, 2001) dominate the plain area/material bound: verified here on
   480,000 random feasible sets with zero violations before being trusted. */
function dffU(x, k) {                      // classic DFF family u^(k)
  if (x <= 0) return 0;
  const t = x * (k + 1);
  if (Math.abs(t - Math.round(t)) < 1e-9) return x;
  return Math.floor(t) / k;
}
function lbDFF1(lengths, C) {              // 1D
  let best = 0;
  for (let k = 1; k <= 12; k++) {
    let s = 0; for (const L of lengths) s += dffU(L / C, k);
    best = Math.max(best, Math.ceil(s - 1e-9));
  }
  return best;
}
function lbDFF2(items, W, H, rot) {         // 2D — DFF applied to each axis
  // A DFF is only defined on [0,1], so a dimension larger than the sheet must
  // be clamped or the "bound" inflates and stops being a bound at all. When
  // rotation is allowed the packer may choose either orientation, so a valid
  // bound has to take the SMALLER of the two contributions.
  let best = 0;
  for (let k1 = 1; k1 <= 10; k1++) for (let k2 = 1; k2 <= 10; k2++) {
    let s = 0;
    for (const p of items) {
      const a = dffU(Math.min(p.w, W) / W, k1) * dffU(Math.min(p.h, H) / H, k2);
      const b = rot ? dffU(Math.min(p.h, W) / W, k1) * dffU(Math.min(p.w, H) / H, k2) : a;
      s += Math.min(a, b);
    }
    best = Math.max(best, Math.ceil(s - 1e-9));
  }
  return best;
}
/* Sheets that provably cannot be avoided: plain area bound, lifted by DFF. */
function lbSheets(queue, W, H, rot) {
  const area = queue.reduce((s, p) => s + Math.min(p.w, W) * Math.min(p.h, H), 0);
  return Math.max(Math.ceil(area / (W * H) - 1e-9), lbDFF2(queue, W, H, rot));
}

function lbMaterial(pieces, kerf) {
  const byStock = new Map();
  for (const p of pieces) byStock.set(p.stock, (byStock.get(p.stock) || 0) + p.length + kerf);
  let n = 0; for (const [stk, tot] of byStock) n += Math.ceil(tot / stk); return n;
}
function lbL2(pieces, kerf) {                       // Martello & Toth (1990)
  const byStock = new Map();
  for (const p of pieces) { if (!byStock.has(p.stock)) byStock.set(p.stock, []); byStock.get(p.stock).push(p.length + kerf); }
  let total = 0;
  for (const [C, ws] of byStock) {
    const w = [...ws].sort((a, b) => b - a); let best = 0;
    const cands = new Set([0]); for (const x of w) if (x <= C / 2) cands.add(x);
    for (const K of cands) {
      const N1 = w.filter(x => x > C - K).length;
      const N2 = w.filter(x => x > C / 2 && x <= C - K);
      const N3 = w.filter(x => x >= K && x <= C / 2);
      const s2 = N2.reduce((a, b) => a + b, 0), s3 = N3.reduce((a, b) => a + b, 0);
      const bnd = N1 + N2.length + Math.max(0, Math.ceil((s3 - (N2.length * C - s2)) / C));
      if (bnd > best) best = bnd;
    }
    total += best;
  }
  return total;
}

/* ── 1D: exact bounded knapsack → one fully-loaded bar ─────────────────── */
function knapPattern(items, stock, kerf) {
  const C = Math.max(0, Math.round(stock));
  if (!C || !items.length) return [];
  const dp = new Float64Array(C + 1);
  const from = new Int32Array(C + 1).fill(-1);
  const back = new Int32Array(C + 1).fill(-1);
  for (let i = 0; i < items.length; i++) {
    const w = Math.round(items[i].len + kerf), v = items[i].len;
    if (w <= 0 || w > C) continue;
    // A single bar physically cannot hold more than floor(C/w) copies, so
    // iterating the whole demand is wasted work on large jobs.
    const q = Math.min(items[i].qty, Math.floor(C / w));
    for (let t = 0; t < q; t++)
      for (let c = C; c >= w; c--) {
        const cand = dp[c - w] + v;
        if (cand > dp[c] + 1e-9) { dp[c] = cand; from[c] = i; back[c] = c - w; }
      }
  }
  let bc = 0; for (let c = 0; c <= C; c++) if (dp[c] > dp[bc]) bc = c;
  const take = new Map(); let c = bc, g = 0;
  while (c > 0 && from[c] >= 0 && g++ < 20000) { const i = from[c]; take.set(i, (take.get(i) || 0) + 1); c = back[c]; }
  return [...take.entries()].map(([i, n]) => ({ i, n: Math.min(n, items[i].qty) })).filter(p => p.n > 0);
}
// Build bins from a plain length multiset. `repeat` reuses an identical
// pattern while demand allows (fast); otherwise every bar is re-optimised.
function binsByKnapsack(lenList, stock, kerf, repeat) {
  const cnt = new Map(); for (const L of lenList) cnt.set(L, (cnt.get(L) || 0) + 1);
  let items = [...cnt.entries()].map(([len, qty]) => ({ len, qty })).sort((a, b) => b.len - a.len);
  const bins = []; let g = 0;
  while (items.length && g++ < 20000) {
    const pat = knapPattern(items, stock, kerf);
    if (!pat.length) { for (const it of items) for (let k = 0; k < it.qty; k++) bins.push({ lens: [it.len], remaining: Math.max(0, stock - it.len - kerf) }); break; }
    let reps = 1;
    if (repeat) { reps = Infinity; for (const { i, n } of pat) reps = Math.min(reps, Math.floor(items[i].qty / n)); reps = Math.max(1, reps); }
    for (let r = 0; r < reps; r++) {
      const lens = []; let used = 0;
      for (const { i, n } of pat) for (let k = 0; k < n; k++) { lens.push(items[i].len); used += items[i].len + kerf; }
      bins.push({ lens, remaining: stock - used });
    }
    for (const { i, n } of pat) items[i].qty -= n * reps;
    items = items.filter(it => it.qty > 0);
  }
  return bins;
}
function binsByFit(lenList, stock, kerf, firstFit) {
  const ls = [...lenList].sort((a, b) => b - a); const bins = [];
  for (const L of ls) {
    let best = null, bestRem = Infinity;
    for (const b of bins) {
      const need = L + kerf;
      if (b.remaining < need) continue;
      if (firstFit) { best = b; break; }
      if (b.remaining - need < bestRem) { bestRem = b.remaining - need; best = b; }
    }
    if (best) { best.lens.push(L); best.remaining -= L + kerf; }
    else bins.push({ lens: [L], remaining: stock - L - kerf });
  }
  return bins;
}
// Dissolve the lightest bar and redistribute — removes a whole bar when it works.
function eliminateBars(bins, stock, kerf, maxPasses) {
  let B = bins.map(b => ({ lens: [...b.lens], remaining: b.remaining }));
  for (let pass = 0; pass < (maxPasses || 60); pass++) {
    if (B.length < 2) break;
    let ti = 0, tLoad = Infinity;
    for (let i = 0; i < B.length; i++) { const load = B[i].lens.reduce((a, c) => a + c, 0); if (load < tLoad) { tLoad = load; ti = i; } }
    const target = B[ti];
    const others = B.filter((_, i) => i !== ti).map(b => ({ lens: [...b.lens], remaining: b.remaining }));
    let allFit = true;
    for (const L of [...target.lens].sort((a, b) => b - a)) {
      let best = null, bestRem = Infinity;
      for (const b of others) { const need = L + kerf; if (b.remaining >= need && b.remaining - need < bestRem) { bestRem = b.remaining - need; best = b; } }
      if (!best) { allFit = false; break; }
      best.lens.push(L); best.remaining -= L + kerf;
    }
    if (allFit) B = others; else break;
  }
  return B;
}
// Best of every construction, then local search. Tie-break favours ONE big
// reusable offcut over the same waste scattered across many bars.
function bestBins(lenList, stock, kerf) {
  if (!lenList.length) return [];
  // Budget: the per-bar knapsack re-optimises every single bar, which is the
  // strongest option but scales with piece count. Above a few hundred pieces
  // it costs seconds for almost no gain, so it is dropped there.
  const heavy = lenList.length <= 400;
  const cands = [
    binsByFit(lenList, stock, kerf, false),   // BFD — what v1 shipped
    binsByFit(lenList, stock, kerf, true),    // FFD
    binsByKnapsack(lenList, stock, kerf, true),
  ];
  if (heavy) cands.push(binsByKnapsack(lenList, stock, kerf, false));
  const passes = lenList.length > 800 ? 15 : 60;
  const improved = cands.map(c => eliminateBars(c, stock, kerf, passes));
  let best = improved[0];
  const maxRem = a => a.length ? Math.max.apply(null, a.map(b => b.remaining)) : 0;
  for (const c of improved) {
    if (c.length < best.length) best = c;
    else if (c.length === best.length && maxRem(c) > maxRem(best)) best = c;
  }
  return best;
}

/* ── 2D: guillotine-feasible packers ──────────────────────────────────── */
function packShelfFF(W, H, q, kerf, rot) {           // First-Fit Decreasing Height
  const placed = [], left = [], shelves = [];
  for (const p of q) {
    let done = false;
    for (const sh of shelves) {
      for (const o of (rot ? [[p.w, p.h, false], [p.h, p.w, true]] : [[p.w, p.h, false]])) {
        if (o[1] <= sh.rowH && sh.usedW + o[0] + (sh.usedW > 0 ? kerf : 0) <= W) {
          placed.push({ x: sh.usedW, y: sh.y, w: o[0], h: o[1], rotated: o[2], src: p }); sh.usedW += o[0] + kerf; done = true; break;
        }
      }
      if (done) break;
    }
    if (done) continue;
    const last = shelves[shelves.length - 1];
    const yTop = last ? last.y + last.rowH + kerf : 0;
    for (const o of (rot ? [[p.w, p.h, false], [p.h, p.w, true]] : [[p.w, p.h, false]])) {
      if (o[0] <= W && yTop + o[1] <= H) {
        shelves.push({ y: yTop, usedW: o[0] + kerf, rowH: o[1] });
        placed.push({ x: 0, y: yTop, w: o[0], h: o[1], rotated: o[2], src: p }); done = true; break;
      }
    }
    if (!done) left.push(p);
  }
  return { placed, left };
}
function packGuil(W, H, q, kerf, rot, split, fit) {  // free-rectangle, guillotine splits
  let free = [{ x: 0, y: 0, w: W, h: H }];
  const placed = [], left = [];
  for (const p of q) {
    let bi = -1, bw = 0, bh = 0, brot = false, bs = Infinity;
    for (let i = 0; i < free.length; i++) {
      const f = free[i];
      for (const o of (rot ? [[p.w, p.h, false], [p.h, p.w, true]] : [[p.w, p.h, false]])) {
        if (o[0] > f.w || o[1] > f.h) continue;
        const sc = fit === "area" ? (f.w * f.h - o[0] * o[1]) : Math.min(f.w - o[0], f.h - o[1]);
        if (sc < bs) { bs = sc; bi = i; bw = o[0]; bh = o[1]; brot = o[2]; }
      }
    }
    if (bi < 0) { left.push(p); continue; }
    const f = free[bi];
    placed.push({ x: f.x, y: f.y, w: bw, h: bh, rotated: brot, src: p });
    const rw = f.w - bw - kerf, rh = f.h - bh - kerf;
    const horiz = split === "SAS" ? (f.w - bw < f.h - bh) : (f.w - bw >= f.h - bh);
    free.splice(bi, 1);
    if (horiz) {
      if (rw > 0) free.push({ x: f.x + bw + kerf, y: f.y, w: rw, h: bh });
      if (rh > 0) free.push({ x: f.x, y: f.y + bh + kerf, w: f.w, h: rh });
    } else {
      if (rw > 0) free.push({ x: f.x + bw + kerf, y: f.y, w: rw, h: f.h });
      if (rh > 0) free.push({ x: f.x, y: f.y + bh + kerf, w: bw, h: rh });
    }
  }
  return { placed, left };
}
const SHEET_SORTS = [
  (a, b) => b.w * b.h - a.w * a.h,
  (a, b) => (b.h - a.h) || (b.w - a.w),
  (a, b) => (b.w - a.w) || (b.h - a.h),
  (a, b) => (b.w + b.h) - (a.w + a.h),
  (a, b) => Math.max(b.w, b.h) - Math.max(a.w, a.h),
];
/* v1 shelf pass, kept verbatim as a candidate so v2 can never lose to it. */
function packShelfNF(W, H, q, kerf, rot) {
  const placed = [], left = [];
  let usedW = 0, rowH = 0, y = 0;
  for (const p of q) {
    const tryP = (w, h, r) => {
      if (usedW + w + kerf <= W + kerf && h <= (H - y)) { placed.push({ x: usedW, y, w, h, rotated: r, src: p }); usedW += w + kerf; rowH = Math.max(rowH, h); return true; }
      return false;
    };
    let ok = tryP(p.w, p.h, false) || (rot && tryP(p.h, p.w, true));
    if (!ok) { y += rowH + kerf; usedW = 0; rowH = 0; if (y < H) ok = tryP(p.w, p.h, false) || (rot && tryP(p.h, p.w, true)); }
    if (!ok) left.push(p);
  }
  return { placed, left };
}

/* Pack the WHOLE job with one strategy, sheet after sheet, to completion. */
function runStrategy(W, H, queue, kerf, rot, packer, cmp, split, fit) {
  let q = [...queue].sort(cmp);
  const sheets = [];
  let guard = 0;
  while (q.length && guard++ < 5000) {
    const r = packer(W, H, q, kerf, rot, split, fit);
    if (!r.placed.length) break;                 // nothing fits an empty sheet
    sheets.push(r.placed);
    q = r.left;
  }
  return { sheets, stuck: q };
}

/* Choose between COMPLETE solutions, not sheet by sheet. Greedily filling the
   current sheet as hard as possible can leave a worse remainder for the next
   one, so per-sheet choices are not safe — measured, not assumed.
   Primary  : fewest sheets.
   Tie-break: least material committed to the final sheet, which concentrates
              the leftover into one large reusable offcut instead of scattering
              small unsellable scrap across several sheets.                  */
function packAllSheetsBest(W, H, queue, kerf, rot) {
  const runs = [];
  for (const cmp of SHEET_SORTS) {
    runs.push(runStrategy(W, H, queue, kerf, rot, packShelfNF, cmp));
    runs.push(runStrategy(W, H, queue, kerf, rot, packShelfFF, cmp));
    for (const sp of ["SAS", "LAS"]) for (const ft of ["area", "short"])
      runs.push(runStrategy(W, H, queue, kerf, rot, packGuil, cmp, sp, ft));
  }
  const lastArea = r => r.sheets.length ? r.sheets[r.sheets.length - 1].reduce((s, p) => s + p.w * p.h, 0) : 0;
  let best = null;
  for (const r of runs) {
    if (!best) { best = r; continue; }
    if (r.stuck.length !== best.stuck.length) { if (r.stuck.length < best.stuck.length) best = r; continue; }
    if (r.sheets.length < best.sheets.length) best = r;
    else if (r.sheets.length === best.sheets.length && lastArea(r) < lastArea(best)) best = r;
  }
  return best || { sheets: [], stuck: queue };
}

/* Single-sheet variant, used when re-nesting into one reusable offcut. */
function packOneSheetBest(W, H, queue, kerf, rot) {
  let best = null, bestArea = -1, bestN = -1;
  for (const cmp of SHEET_SORTS) {
    const sorted = [...queue].sort(cmp);
    const runs = [packShelfNF(W, H, sorted, kerf, rot), packShelfFF(W, H, sorted, kerf, rot)];
    for (const sp of ["SAS", "LAS"]) for (const ft of ["area", "short"]) runs.push(packGuil(W, H, sorted, kerf, rot, sp, ft));
    for (const r of runs) {
      const area = r.placed.reduce((s, p) => s + p.w * p.h, 0);
      if (area > bestArea || (area === bestArea && r.placed.length > bestN)) { bestArea = area; bestN = r.placed.length; best = r; }
    }
  }
  return best || { placed: [], left: queue };
}

/* ─── PLATE NESTING ENGINE (with reusable offcuts + oversize detection) ──── */
function splitPart(W, L, uw, uh, allowRotation, pref) {
  // Split a part too big for the sheet into welded sub-pieces, divided EQUALLY so
  // the weld lands sensibly and there are never tiny slivers.
  //   pref "welds" → fewest pieces / fewest welds (bigger pieces, may pack looser)
  //   pref "pack"  → smaller equal pieces that tile the sheet height better (less scrap)
  const layout = (a, b, colDiv, rowDiv) => {
    const cols = Math.max(1, colDiv(a), Math.ceil(a / uw)), rows = Math.max(1, rowDiv(b), Math.ceil(b / uh));
    const cw = a / cols, rh = b / rows, list = [];
    for (let c = 0; c < cols; c++) for (let r = 0; r < rows; r++) list.push({ w: Math.round(cw), h: Math.round(rh) });
    return list;
  };
  const minPieces = (a, b) => layout(a, b, () => 1, () => 1);
  // packing layout: choose row/col counts so each piece ≈ half the sheet (tiles better)
  const packLayout = (a, b) => layout(a, b, x => Math.ceil(x / (uw / 2)), y => Math.ceil(y / (uh / 2)));
  const make = (a, b) => pref === "pack" ? packLayout(a, b) : minPieces(a, b);
  const o1 = make(W, L);
  if (allowRotation) { const o2 = make(L, W); if (o2.length < o1.length) return o2; }
  return o1;
}

function nestPlates(sheetW, sheetH, parts, kerf, margin, allowRotation, reuseMin, splicePref) {
  const sheets = [];
  const fitsRaw = (w, h) => (w <= sheetW && h <= sheetH) || (allowRotation && h <= sheetW && w <= sheetH);

  // Auto-clamp the edge margin from the parts that already fit, so a near-sheet part isn't blocked.
  let m = margin;
  parts.forEach(p => { if (fitsRaw(p.width, p.length)) { const longP = Math.max(p.width, p.length), shortP = Math.min(p.width, p.length); const longS = Math.max(sheetW, sheetH), shortS = Math.min(sheetW, sheetH); if (longP <= longS) m = Math.min(m, Math.floor((longS - longP) / 2)); if (shortP <= shortS) m = Math.min(m, Math.floor((shortS - shortP) / 2)); } });
  m = Math.max(0, m);
  const usableW = sheetW - m * 2, usableH = sheetH - m * 2;

  // Build the piece queue, SPLICING any part bigger than the sheet into welded sub-pieces.
  let queue = [];
  const spliceMap = {}; // label -> {label, W, L, count, qty}
  parts.forEach(p => {
    for (let q = 0; q < p.qty; q++) {
      if (fitsRaw(p.width, p.length)) { queue.push({ w: p.width, h: p.length, id: p.id, label: p.label }); }
      else {
        const subs = splitPart(p.width, p.length, usableW, usableH, allowRotation, splicePref);
        subs.forEach((s, idx) => queue.push({ w: s.w, h: s.h, id: p.id, label: `${p.label}·${idx + 1}/${subs.length}`, spliced: true }));
        const key = p.label; if (!spliceMap[key]) spliceMap[key] = { label: p.label, W: p.width, L: p.length, count: subs.length, qty: 0 }; spliceMap[key].qty += 1;
      }
    }
  });

  queue.sort((a, b) => (b.w * b.h) - (a.w * a.h));
  const queue0 = queue.slice();            // kept for the lower-bound readout
  // v2: solve the whole job under every guillotine-feasible strategy and keep
  // the best complete solution. The v1 shelf pass is one of the strategies, so
  // the result can never be worse than what v1 produced.
  const solved = packAllSheetsBest(usableW, usableH, queue, kerf, allowRotation);
  solved.sheets.forEach(pl => {
    const placements = pl.map(q => ({
      x: m + q.x, y: m + q.y, w: q.w, h: q.h, rotated: q.rotated,
      id: q.src.id, label: q.src.label, spliced: q.src.spliced,
    }));
    sheets.push({ placements, offcuts: findOffcuts(placements, m, usableW, usableH, reuseMin) });
  });
  // Anything that cannot fit even an empty sheet gets its own clamped sheet,
  // exactly as before, so oversize parts stay visible instead of vanishing.
  solved.stuck.forEach(p => {
    sheets.push({ placements: [{ x: m, y: m, w: Math.min(p.w, usableW), h: Math.min(p.h, usableH), rotated: false, id: p.id, label: p.label, spliced: p.spliced }], offcuts: [] });
  });
  sheets.splices = Object.values(spliceMap);
  sheets.marginUsed = m;
  // Optimality certificate: when the sheet count equals a proven lower bound,
  // no nesting algorithm can do better on this cut list.
  sheets.lowerBound = lbSheets(queue0, usableW, usableH, allowRotation);
  sheets.proven = sheets.length <= sheets.lowerBound;
  return sheets;
}


/* Run the v2 bin packer on labelled pieces and map the chosen patterns back
   onto the original labels. Grouped by stock length: a cutting pattern is
   only valid within one market bar size. */
function buildBinsV2(pieces, kerf) {
  const byStock = new Map();
  for (const p of pieces) { if (!byStock.has(p.stock)) byStock.set(p.stock, []); byStock.get(p.stock).push(p); }
  const out = [];
  for (const [stock, ps] of byStock) {
    const packed = bestBins(ps.map(p => p.length), stock, kerf);
    // pool of real pieces keyed by length, so labels follow their cut
    const pool = new Map();
    for (const p of ps) { if (!pool.has(p.length)) pool.set(p.length, []); pool.get(p.length).push(p); }
    for (const b of packed) {
      const cuts = [];
      for (const L of b.lens) {
        const bucket = pool.get(L);
        const src = (bucket && bucket.length) ? bucket.pop() : { length: L, label: "", spliced: false };
        cuts.push({ length: src.length, label: src.label, spliced: src.spliced });
      }
      out.push({ stockLength: stock, cuts, remaining: b.remaining });
    }
  }
  return out;
}

/* ─── SECTION (1D BAR) NESTING ENGINE ────────────────────────────────────── */
function nestBars(items, kerf) {
  const pieces = []; const spliceMap = {}; let fullBarsFromSplices = 0;
  items.forEach(it => {
    for (let i = 0; i < it.qty; i++) {
      if (it.length > it.stock) {
        const bars = Math.ceil(it.length / it.stock); const fullBars = bars - 1; fullBarsFromSplices += fullBars;
        const remainder = it.length - fullBars * it.stock;
        if (remainder > 0) pieces.push({ length: remainder, label: it.label || "", stock: it.stock, spliced: true });
        const key = `${it.length}`; if (!spliceMap[key]) spliceMap[key] = { length: it.length, qty: 0, bars, stock: it.stock }; spliceMap[key].qty += 1;
      } else pieces.push({ length: it.length, label: it.label || "", stock: it.stock });
    }
  });
  pieces.sort((a, b) => b.length - a.length);
  const bins = [];
  for (let i = 0; i < fullBarsFromSplices; i++) { const stock = items[0].stock; bins.push({ stockLength: stock, cuts: [{ length: stock, label: "(spliced run)", spliced: true }], remaining: 0, spliceBar: true }); }
  buildBinsV2(pieces, kerf).forEach(b => bins.push(b));
  const totalStock = bins.reduce((s, b) => s + b.stockLength, 0);
  const totalNet = bins.reduce((s, b) => s + b.cuts.reduce((ss, c) => ss + c.length, 0), 0);
  const totalWaste = bins.reduce((s, b) => s + b.remaining, 0);
  const wastePct = totalStock ? ((totalWaste / totalStock) * 100).toFixed(1) : "0";
  // Optimality certificate (Korf, AAAI-02): if the bars used equals a valid
  // lower bound, no algorithm on earth can do better on this cut list.
  const lbBars = Math.max(
    lbMaterial(pieces, kerf),
    lbL2(pieces, kerf),
    lbDFF1(pieces.map(p => p.length + kerf), pieces.length ? pieces[0].stock : 12000)
  ) + fullBarsFromSplices;
  const proven = bins.length <= lbBars;
  return { bins, splices: Object.values(spliceMap), summary: { stockCount: bins.length, totalStock, totalNet, totalWaste, wastePct, utilPct: (100 - parseFloat(wastePct)).toFixed(1), lowerBound: lbBars, proven } };
}

/* ─── SECTION NESTING WITH OWNED LEFTOVERS ────────────────────────────────
   Cut from leftover bars you already own FIRST (free), then buy the minimum
   new market bars for whatever is left. `leftovers` = [{length, qty}].
   Returns the same shape as nestBars plus `reusedBins` and reused summary. */
function nestBarsLO(items, kerf, leftovers = []) {
  const pieces = []; const spliceMap = {}; let fullBarsFromSplices = 0;
  const stock0 = items.length ? items[0].stock : 12000;
  items.forEach(it => {
    for (let i = 0; i < it.qty; i++) {
      if (it.length > it.stock) {
        const bars = Math.ceil(it.length / it.stock); const fullBars = bars - 1; fullBarsFromSplices += fullBars;
        const remainder = it.length - fullBars * it.stock;
        if (remainder > 0) pieces.push({ length: remainder, label: it.label || "", stock: it.stock, spliced: true });
        const key = `${it.length}`; if (!spliceMap[key]) spliceMap[key] = { length: it.length, qty: 0, bars, stock: it.stock }; spliceMap[key].qty += 1;
      } else pieces.push({ length: it.length, label: it.label || "", stock: it.stock });
    }
  });
  pieces.sort((a, b) => b.length - a.length);

  // Owned leftover bars become pre-seeded "free" bins, longest first.
  const reusedBins = [];
  leftovers.forEach(lo => { const n = Math.max(0, Math.round(lo.qty || 0)); const L = Math.round(lo.length || 0); if (L <= 0) return; for (let i = 0; i < n; i++) reusedBins.push({ stockLength: L, cuts: [], remaining: L, reused: true }); });
  reusedBins.sort((a, b) => b.stockLength - a.stockLength);

  // PASS 1 — fill owned leftovers first (tightest remaining wins).
  const leftPieces = [];
  for (const p of pieces) {
    let best = null, bestRem = Infinity;
    for (const b of reusedBins) { const need = p.length + kerf; if (b.remaining >= need && (b.remaining - need) < bestRem) { bestRem = b.remaining - need; best = b; } }
    if (best) { best.cuts.push({ length: p.length, label: p.label, spliced: p.spliced }); best.remaining -= p.length + kerf; }
    else leftPieces.push(p);
  }

  // PASS 2 — remaining pieces go onto NEW market bars (the minimum to buy).
  const bins = [];
  for (let i = 0; i < fullBarsFromSplices; i++) bins.push({ stockLength: stock0, cuts: [{ length: stock0, label: "(spliced run)", spliced: true }], remaining: 0, spliceBar: true });
  buildBinsV2(leftPieces, kerf).forEach(b => bins.push(b));

  const newStock = bins.reduce((s, b) => s + b.stockLength, 0);
  const newNet = bins.reduce((s, b) => s + b.cuts.reduce((ss, c) => ss + c.length, 0), 0);
  const newWaste = bins.reduce((s, b) => s + b.remaining, 0);
  const reusedNet = reusedBins.reduce((s, b) => s + b.cuts.reduce((ss, c) => ss + c.length, 0), 0);
  const reusedInput = reusedBins.reduce((s, b) => s + b.stockLength, 0);
  const reusedRemain = reusedBins.reduce((s, b) => s + b.remaining, 0);
  const usedReused = reusedBins.filter(b => b.cuts.length > 0);
  const wastePct = newStock ? ((newWaste / newStock) * 100).toFixed(1) : "0";
  return {
    bins, reusedBins, splices: Object.values(spliceMap),
    summary: {
      stockCount: bins.length, totalStock: newStock, totalNet: newNet, totalWaste: newWaste,
      wastePct, utilPct: (100 - parseFloat(wastePct)).toFixed(1),
      reusedCount: usedReused.length, reusedProvided: reusedBins.length, reusedNet, reusedInput, reusedRemain,
    }
  };
}

/* ─── L-SHAPE FROM 6 CORNER POINTS ────────────────────────────────────────
   A rectilinear L is a 6-vertex polygon. From the points we recover the
   overall A×B box and the missing-corner notch (C,D), then the largest
   usable rectangle (same convention as lUsableRect). Works for any corner. */
function defaultLPoints() {
  // A=1200 wide, B=2400 long, notch 400×600 at the TOP-RIGHT, traced from origin.
  return [
    { x: 0, y: 0 }, { x: 1200, y: 0 }, { x: 1200, y: 1800 },
    { x: 800, y: 1800 }, { x: 800, y: 2400 }, { x: 0, y: 2400 },
  ];
}
function pointInPoly(pt, poly) {
  let inside = false;
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
    const xi = poly[i].x, yi = poly[i].y, xj = poly[j].x, yj = poly[j].y;
    const hit = ((yi > pt.y) !== (yj > pt.y)) && (pt.x < (xj - xi) * (pt.y - yi) / (yj - yi) + xi);
    if (hit) inside = !inside;
  }
  return inside;
}
function analyzeL(points) {
  const xs = [...new Set(points.map(p => Math.round(p.x)))].sort((a, b) => a - b);
  const ys = [...new Set(points.map(p => Math.round(p.y)))].sort((a, b) => a - b);
  const minX = xs[0], maxX = xs[xs.length - 1], minY = ys[0], maxY = ys[ys.length - 1];
  const A = maxX - minX, B = maxY - minY;
  let C = 0, D = 0, corner = null;
  if (xs.length >= 3 && ys.length >= 3 && A > 0 && B > 0) {
    const x1 = xs[1], y1 = ys[1];
    const cells = [
      { cx: (minX + x1) / 2, cy: (minY + y1) / 2, w: x1 - minX, h: y1 - minY, corner: "BL" },
      { cx: (x1 + maxX) / 2, cy: (minY + y1) / 2, w: maxX - x1, h: y1 - minY, corner: "BR" },
      { cx: (minX + x1) / 2, cy: (y1 + maxY) / 2, w: x1 - minX, h: maxY - y1, corner: "TL" },
      { cx: (x1 + maxX) / 2, cy: (y1 + maxY) / 2, w: maxX - x1, h: maxY - y1, corner: "TR" },
    ];
    const notch = cells.find(c => !pointInPoly({ x: c.cx, y: c.cy }, points));
    if (notch) { C = Math.round(notch.w); D = Math.round(notch.h); corner = notch.corner; }
  }
  const usable = lUsableRect(A, B, C, D);
  return { A, B, C, D, corner, usable, valid: xs.length === 3 && ys.length === 3 };
}

/* ─── RE-IMPORTABLE LEFTOVER FILES ────────────────────────────────────────
   Every Leftover PDF/Excel we export carries a machine-readable token so the
   same file can be dropped back in next job. Token survives printing to PDF
   (it's small visible text) and is also embedded as JSON for the HTML/Excel. */
const b64enc = s => { try { return btoa(unescape(encodeURIComponent(s))); } catch { return ""; } };
const b64dec = s => { try { return decodeURIComponent(escape(atob(s))); } catch { return ""; } };
function encodeLeftoverToken(payload) { return `STEELOPT<<${b64enc(JSON.stringify(payload))}>>END`; }
function decodeLeftoverToken(text) {
  // PDF/HTML extraction can sprinkle whitespace anywhere, so flatten first, then
  // match the unique anchors around a pure base64 body.
  const flat = String(text || "").replace(/\s+/g, "");
  const m = flat.match(/STEELOPT<<([A-Za-z0-9+/=]*)>>END/);
  if (!m) return null;
  try { return JSON.parse(b64dec(m[1])); } catch { return null; }
}
let _pdfjs = null;
async function loadPdfJs() {
  if (_pdfjs) return _pdfjs;
  if (typeof window !== "undefined" && window.pdfjsLib) { _pdfjs = window.pdfjsLib; return _pdfjs; }
  await new Promise((res, rej) => { const s = document.createElement("script"); s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"; s.onload = res; s.onerror = () => rej(new Error(RT("pdf reader unavailable", "قارئ PDF غير متاح"))); document.head.appendChild(s); });
  _pdfjs = window.pdfjsLib;
  _pdfjs.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  return _pdfjs;
}
async function pdfTextOf(file) {
  const pdfjs = await loadPdfJs();
  const buf = await file.arrayBuffer();
  const doc = await pdfjs.getDocument({ data: buf }).promise;
  let txt = "";
  for (let i = 1; i <= doc.numPages; i++) { const pg = await doc.getPage(i); const c = await pg.getTextContent(); txt += c.items.map(it => it.str).join(" ") + "\n"; }
  return txt;
}
// Read any leftover file we produced (or a plain spreadsheet) → {type:'sections'|'plates', items:[...]}.
async function parseLeftoverFile(file) {
  const name = (file.name || "").toLowerCase();
  // 1) PDF — pull the embedded token from the page text.
  if (name.endsWith(".pdf")) {
    const txt = await pdfTextOf(file);
    const tok = decodeLeftoverToken(txt);
    if (tok && tok.items) return { type: tok.t === "plate" ? "plates" : "sections", items: tok.items };
    throw new Error(RT("No Steel Optimizer data found in this PDF. Re-export the Leftover PDF and try again.", "لا توجد بيانات Steel Optimizer في ملف PDF هذا. صدّر ملف البواقي PDF من جديد وحاول مرة أخرى."));
  }
  // 2) HTML — token or embedded JSON island.
  if (name.endsWith(".html") || name.endsWith(".htm")) {
    const txt = await file.text();
    let tok = decodeLeftoverToken(txt);
    if (!tok) { const m = txt.match(/<script[^>]*id="steelopt"[^>]*>([\s\S]*?)<\/script>/); if (m) { try { tok = JSON.parse(m[1]); } catch { /* ignore */ } } }
    if (tok && tok.items) return { type: tok.t === "plate" ? "plates" : "sections", items: tok.items };
    throw new Error(RT("This HTML file has no Steel Optimizer leftover data.", "ملف HTML هذا لا يحتوي بيانات بواقٍ من Steel Optimizer."));
  }
  // 3) Excel / CSV — token sheet first, then human columns.
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  for (const sn of wb.SheetNames) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "" });
    for (const row of aoa) for (const cell of row) { const tok = decodeLeftoverToken(cell); if (tok && tok.items) return { type: tok.t === "plate" ? "plates" : "sections", items: tok.items }; }
  }
  // Fallback: parse the first sheet by recognising headers.
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "" });
  let hr = -1, H = [];
  for (let i = 0; i < aoa.length; i++) { const cells = aoa[i].map(c => String(c).toLowerCase()); if (cells.some(c => /profile/.test(c)) || cells.some(c => /thickness|offcut/.test(c))) { hr = i; H = aoa[i].map(c => String(c).toLowerCase()); break; } }
  if (hr === -1) throw new Error(RT("Couldn't recognise this file as a leftover list.", "تعذّر التعرّف على هذا الملف كقائمة بواقٍ."));
  const col = (...names) => { for (const n of names) { const k = H.findIndex(h => h.includes(n)); if (k !== -1) return k; } return -1; };
  const body = aoa.slice(hr + 1).filter(r => r.some(c => String(c).trim() !== "") && !/^total/i.test(String(r[0])));
  const isPlate = H.some(h => /thickness|offcut/.test(h));
  if (isPlate) {
    const tc = col("thickness"), wc = col("offcut w", "width", "w (mm)"), lc = col("offcut l", "length", "l (mm)"), sc = col("shape"), qc = col("qty", "pieces"), nwc = col("notch w"), nhc = col("notch h");
    const items = [];
    body.forEach(r => { const W = +r[wc] || 0, L = +r[lc] || 0; if (W <= 0 || L <= 0) return; const shape = sc !== -1 && /l/i.test(String(r[sc])) ? "L" : "rect"; items.push({ label: "Imported", thickness: +r[tc] || 8, shape, A: W, B: L, C: shape === "L" && nwc !== -1 ? +r[nwc] || 0 : 0, D: shape === "L" && nhc !== -1 ? +r[nhc] || 0 : 0, qty: qc !== -1 ? Math.max(1, +r[qc] || 1) : 1 }); });
    if (!items.length) throw new Error(RT("No plate leftovers found in this file.", "لا توجد بواقي ألواح في هذا الملف."));
    return { type: "plates", items };
  } else {
    const pc = col("profile"), gc = col("grade"), lc = col("reusable leftover", "leftover", "length"), qc = col("qty", "pieces");
    const agg = {};
    body.forEach(r => { const len = Math.round(+r[lc] || 0); if (len <= 0) return; const profile = String(r[pc] || "").trim().toUpperCase() || "(unspecified)"; const grade = gc !== -1 ? String(r[gc] || "").trim().toUpperCase() : ""; const q = qc !== -1 ? Math.max(1, +r[qc] || 1) : 1; const k = `${profile}||${grade}||${len}`; if (!agg[k]) agg[k] = { profile, grade, length: len, qty: 0 }; agg[k].qty += q; });
    const items = Object.values(agg);
    if (!items.length) throw new Error(RT("No bar leftovers found in this file.", "لا توجد بواقي أعواد في هذا الملف."));
    return { type: "sections", items };
  }
}


/* ════════════════════════════════════════════════════════════════════════
   SCAN IMPORT — PDF (text layer or scanned) and photos
   ------------------------------------------------------------------------
   Best-effort extraction only. Nothing here reaches the optimiser until the
   user has confirmed it in the review table: a misread digit is a wrongly
   cut bar, so the confirm step is deliberate and cannot be skipped.
   Everything runs in the browser — no file is uploaded anywhere.
════════════════════════════════════════════════════════════════════════ */
let _tesseract = null;
async function loadTesseract() {
  if (_tesseract) return _tesseract;
  if (typeof window !== "undefined" && window.Tesseract) { _tesseract = window.Tesseract; return _tesseract; }
  await new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/tesseract.js/5.1.0/tesseract.min.js";
    s.onload = res; s.onerror = () => rej(new Error("text reader unavailable"));
    document.head.appendChild(s);
  });
  _tesseract = window.Tesseract;
  if (!_tesseract) throw new Error(RT("text reader unavailable", "قارئ النص غير متاح"));
  return _tesseract;
}
async function ocrOne(src, onPct) {
  const T = await loadTesseract();
  const r = await T.recognize(src, "eng", {
    logger: m => { if (m && m.status === "recognizing text" && onPct) onPct(Math.round((m.progress || 0) * 100)); },
  });
  return (r && r.data && r.data.text) || "";
}
// Rasterise a scanned PDF so OCR has something to look at. Capped for speed.
async function pdfPageImages(file, maxPages = 5) {
  const pdfjs = await loadPdfJs();
  const buf = await file.arrayBuffer();
  const doc = await pdfjs.getDocument({ data: buf }).promise;
  const out = [];
  const n = Math.min(doc.numPages, maxPages);
  for (let i = 1; i <= n; i++) {
    const pg = await doc.getPage(i);
    const vp = pg.getViewport({ scale: 2 });
    const cv = document.createElement("canvas");
    cv.width = Math.round(vp.width); cv.height = Math.round(vp.height);
    await pg.render({ canvasContext: cv.getContext("2d"), viewport: vp }).promise;
    out.push(cv.toDataURL("image/png"));
  }
  return out;
}
async function scanFileToText(file, onStage) {
  const name = (file.name || "").toLowerCase();
  const type = file.type || "";
  const isImg = /\.(png|jpe?g|webp|bmp|gif)$/.test(name) || type.indexOf("image/") === 0;
  if (isImg) { if (onStage) onStage("ocr", 0); return await ocrOne(file, p => onStage && onStage("ocr", p)); }
  if (name.endsWith(".pdf") || type === "application/pdf") {
    let txt = "";
    try { txt = await pdfTextOf(file); } catch { txt = ""; }
    if (txt.replace(/\s/g, "").length > 80) return txt;         // real text layer — no OCR needed
    if (onStage) onStage("ocr", 0);                              // scanned PDF → rasterise + OCR
    const pages = await pdfPageImages(file);
    let all = "";
    for (let i = 0; i < pages.length; i++) {
      all += await ocrOne(pages[i], p => onStage && onStage("ocr", Math.round(((i + p / 100) / pages.length) * 100))) + "\n";
    }
    return all;
  }
  return await file.text();
}

const _snum = s => { const v = parseFloat(String(s).replace(/,/g, "")); return isFinite(v) ? v : NaN; };
const _SKIP_LINE = /^(total|subtotal|grand|page|sheet|drawing|project|client|rev\b|date\b|scale\b|notes?\b)/i;
// Rolled-section prefixes: those rows belong to the Sections module, not here.
const _SCAN_HOT = /\b(SHS|RHS|CHS|IPE|IPN|HEA|HEB|HEM|UBP|UB|UC|PFC|UPN|UPE|HSS|RSA|RSJ)\b/i;
function scanLines(text) {
  return String(text || "")
    .replace(/[×✕✖⨯]/g, "x")
    .replace(/[|]/g, " ")
    .split(/\r?\n/)
    .map(l => l.replace(/\s+/g, " ").trim())
    .filter(l => l.length >= 3);
}
// Sections: a profile the library recognises, plus a length, plus a quantity.
function scanRowsSections(text) {
  const out = [];
  for (const raw of scanLines(text)) {
    if (_SKIP_LINE.test(raw)) continue;
    // Pull the grade out first so it can never be swallowed into the profile.
    const gm = raw.match(/\bS\s?(235|275|355|420|460)\s?[A-Z]{0,3}\b/i);
    const line = gm ? raw.replace(gm[0], " ").replace(/\s+/g, " ").trim() : raw;
    const toks = line.split(" ");
    let profile = "", after = -1;
    for (let i = 0; i < toks.length && !profile; i++) {
      for (let take = 3; take >= 1; take--) {
        const cand = toks.slice(i, i + take).join(" ").replace(/[^A-Za-z0-9x.\/ -]/g, "").trim();
        if (cand.length < 3 || !/[A-Za-z]/.test(cand) || !/\d/.test(cand)) continue;
        if (findSection(cand)) { profile = cand.toUpperCase(); after = i + take; break; }
      }
    }
    if (!profile) continue;
    const nums = toks.slice(after).map(_snum).filter(v => isFinite(v) && v > 0);
    const len = nums.find(v => v >= 150 && v <= 30000);
    if (!len) continue;
    const qty = nums.find(v => v !== len && v >= 1 && v <= 999 && Number.isInteger(v)) || 1;
    const grade = gm ? gm[0].replace(/\s/g, "").toUpperCase() : "";
    out.push({ profile, grade, length: Math.round(len), qty: Math.round(qty) });
  }
  return out;
}
// Plates: L x W x T on one line, or three numbers where one reads as a thickness.
function scanRowsPlates(text) {
  const out = [];
  for (const line of scanLines(text)) {
    if (_SKIP_LINE.test(line)) continue;
    if (_SCAN_HOT.test(line)) continue;
    let L = 0, W = 0, T = 0, rest = line;
    const triple = line.match(/(\d{2,5}(?:\.\d+)?)\s*x\s*(\d{2,5}(?:\.\d+)?)\s*x\s*(\d{1,3}(?:\.\d+)?)/i);
    if (triple) {
      const a = _snum(triple[1]), b = _snum(triple[2]), c = _snum(triple[3]);
      L = Math.max(a, b); W = Math.min(a, b); T = c;
      rest = line.replace(triple[0], " ");
    } else {
      const nums = line.split(" ").map(_snum).filter(v => isFinite(v) && v > 0);
      const thick = nums.find(v => v >= 2 && v <= 120);
      const big = nums.filter(v => v !== thick && v >= 50 && v <= 12000).sort((x, y) => y - x);
      if (!thick || big.length < 2) continue;
      L = big[0]; W = big[1]; T = thick;
    }
    if (!(L > 0 && W > 0 && T > 0) || T > L || T > W) continue;
    const qn = rest.split(" ").map(_snum).filter(v => isFinite(v) && Number.isInteger(v) && v >= 1 && v <= 999);
    out.push({ length: Math.round(L), width: Math.round(W), thickness: T, qty: qn.length ? Math.round(qn[qn.length - 1]) : 1 });
  }
  return out.map((r, i) => ({ ...r, label: "P" + (i + 1) }));
}

function ScanImporter({ mode, onUse }) {
  const t = useT();
  const { lang } = useLang();
  const pls = n => ((lang === "en" || lang === "es") && n > 1) ? "s" : "";
  const [busy, setBusy] = useState(false);
  const [stage, setStage] = useState("");
  const [pct, setPct] = useState(0);
  const [err, setErr] = useState("");
  const [rows, setRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const inRef = useRef(null);

  const run = async file => {
    if (!file) return;
    setErr(""); setRows([]); setFileName(file.name || ""); setBusy(true); setStage("read"); setPct(0);
    try {
      const text = await scanFileToText(file, (st, p) => { setStage(st); setPct(p || 0); });
      const found = (mode === "plates" ? scanRowsPlates(text) : scanRowsSections(text)).slice(0, 200);
      if (!found.length) setErr(t("scanNoRows"));
      setRows(found);
    } catch (e) {
      setErr(t("scanFail", { msg: (e && e.message) || "unknown" }));
    }
    setBusy(false); setStage(""); setPct(0);
  };

  const up = (i, k, v) => setRows(rs => rs.map((r, j) => j === i
    ? { ...r, [k]: (k === "profile" || k === "grade" || k === "label") ? v : (+v || 0) } : r));
  const del = i => setRows(rs => rs.filter((_, j) => j !== i));
  const add = () => setRows(rs => [...rs, mode === "plates"
    ? { label: "P" + (rs.length + 1), length: 0, width: 0, thickness: 8, qty: 1 }
    : { profile: "", grade: "", length: 0, qty: 1 }]);

  const valid = rows.filter(r => mode === "plates"
    ? r.length > 0 && r.width > 0 && r.thickness > 0 && r.qty > 0
    : String(r.profile || "").trim() && r.length > 0 && r.qty > 0);

  const DROP = {
    border: "2px dashed #2d3748", borderRadius: 10, padding: "40px 22px", textAlign: "center",
    cursor: busy ? "default" : "pointer", background: "rgba(15,19,24,.5)", transition: "border-color .2s",
  };

  return (
    <div>
      <div style={{ marginBottom: 14, padding: "11px 14px", borderRadius: 6, background: "rgba(245,158,11,.08)", border: "1px solid rgba(245,158,11,.25)", color: "#fcd34d", fontSize: 14, lineHeight: 1.6 }}>{t("scanWarn")}</div>

      {rows.length === 0 && (
        <div
          style={DROP}
          onClick={() => { if (!busy) inRef.current && inRef.current.click(); }}
          onDragOver={e => e.preventDefault()}
          onDrop={e => { e.preventDefault(); if (!busy) run(e.dataTransfer.files && e.dataTransfer.files[0]); }}
          onMouseEnter={e => { if (!busy) e.currentTarget.style.borderColor = "#f59e0b"; }}
          onMouseLeave={e => { e.currentTarget.style.borderColor = "#2d3748"; }}
        >
          <div style={{ fontSize: 46, marginBottom: 12 }}>{busy ? "⏳" : "🔍"}</div>
          {busy ? (
            <div style={{ color: "#fcd34d", fontSize: 16, fontFamily: "'Space Mono', monospace" }}>
              {stage === "ocr" ? t("scanOcr", { pct }) : t("scanReading")}
            </div>
          ) : (
            <>
              <div style={{ color: "#cbd5e1", fontSize: 17, marginBottom: 6 }}>{t("scanDrop")}</div>
              <div style={{ color: "#94a3b8", fontSize: 14 }}>{t("scanAccept")}</div>
              <div style={{ color: "#475569", fontSize: 13, marginTop: 12, lineHeight: 1.7 }}>{t("scanTip")}</div>
            </>
          )}
          <input ref={inRef} type="file" accept=".pdf,.png,.jpg,.jpeg,.webp,image/*" style={{ display: "none" }}
                 onChange={e => { const f = e.target.files && e.target.files[0]; e.target.value = ""; run(f); }} />
        </div>
      )}

      {err && <div style={{ marginTop: 12, padding: "9px 13px", borderRadius: 5, background: "rgba(245,158,11,.14)", border: "1px solid rgba(245,158,11,.4)", color: "#fcd34d", fontSize: 14, lineHeight: 1.6 }}>{err}</div>}

      {rows.length > 0 && (
        <div style={{ marginTop: 4 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 12 }}>
            <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 14, color: "#22c55e" }}>{t("scanFound", { n: rows.length, s: pls(rows.length) })}</span>
            <span style={{ fontSize: 13, color: "#475569" }}>{fileName}</span>
            <button onClick={() => { setRows([]); setErr(""); }} style={{ marginInlineStart: "auto", background: "none", border: "1px solid #2d3748", borderRadius: 4, color: "#94a3b8", cursor: "pointer", fontSize: 13, padding: "5px 12px" }}>{t("scanRedo")}</button>
          </div>

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 14 }}>
              <thead><tr style={{ borderBottom: "1px solid #2d3748" }}>
                {(mode === "plates"
                  ? [t("lblId"), t("lblLength"), t("lblWidth"), t("lblThickness"), t("lblQty"), ""]
                  : [t("lblProfile"), t("lblGradeShort"), t("lblLength"), t("thQty"), ""]
                ).map((h, hi) => <th key={hi} style={{ textAlign: "start", padding: "6px 8px", color: "#64748b", fontWeight: 600, fontSize: 12, letterSpacing: 1, fontFamily: "'Space Mono', monospace" }}>{h}</th>)}
              </tr></thead>
              <tbody>
                {rows.map((r, i) => (
                  <tr key={i} style={{ borderBottom: "1px solid #1a2230" }}>
                    {mode === "plates" ? (
                      <>
                        <td style={{ padding: "5px 4px" }}><input value={r.label} onChange={e => up(i, "label", e.target.value)} style={{ ...CI, width: 56 }} /></td>
                        <td style={{ padding: "5px 4px" }}><input type="number" value={r.length} onChange={e => up(i, "length", e.target.value)} style={{ ...CI, width: 82 }} /></td>
                        <td style={{ padding: "5px 4px" }}><input type="number" value={r.width} onChange={e => up(i, "width", e.target.value)} style={{ ...CI, width: 82 }} /></td>
                        <td style={{ padding: "5px 4px" }}><input type="number" value={r.thickness} onChange={e => up(i, "thickness", e.target.value)} style={{ ...CI, width: 78, borderColor: "#d9770699" }} /></td>
                        <td style={{ padding: "5px 4px" }}><input type="number" value={r.qty} onChange={e => up(i, "qty", e.target.value)} style={{ ...CI, width: 58 }} /></td>
                      </>
                    ) : (
                      <>
                        <td style={{ padding: "5px 4px" }}><input value={r.profile} onChange={e => up(i, "profile", e.target.value)} style={{ ...CI, width: 140 }} /></td>
                        <td style={{ padding: "5px 4px" }}><input value={r.grade} onChange={e => up(i, "grade", e.target.value)} style={{ ...CI, width: 84 }} /></td>
                        <td style={{ padding: "5px 4px" }}><input type="number" value={r.length} onChange={e => up(i, "length", e.target.value)} style={{ ...CI, width: 90 }} /></td>
                        <td style={{ padding: "5px 4px" }}><input type="number" value={r.qty} onChange={e => up(i, "qty", e.target.value)} style={{ ...CI, width: 58 }} /></td>
                      </>
                    )}
                    <td style={{ padding: "5px 4px" }}><button onClick={() => del(i)} style={{ background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 16 }}>✕</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 14 }}>
            <button onClick={add} style={{ padding: "8px 18px", background: "rgba(245,158,11,.15)", border: "1px dashed #d97706", color: "#fbbf24", borderRadius: 4, cursor: "pointer", fontSize: 14 }}>{t("scanAddRow")}</button>
            <button onClick={() => valid.length && onUse(valid)} disabled={!valid.length}
              style={{ padding: "10px 24px", borderRadius: 6, border: "none", cursor: valid.length ? "pointer" : "not-allowed", fontSize: 15, fontWeight: 700, background: valid.length ? "linear-gradient(135deg,#10b981,#059669)" : "#1a2230", color: valid.length ? "#04140d" : "#475569" }}>
              {t("scanUse")}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════════════
   3D COVER — rotating I-beam + plate (Canvas 2D perspective, amber theme)
════════════════════════════════════════════════════════════════════════ */
/* ─── LIVE CUT SHOWCASE — real-case nesting demo + feature chips, below hero ── */
function LiveCutShowcase() {
  const T = useT();
  const Tref = useRef(T); Tref.current = T;          // canvas text follows the language without restarting the animation
  const ref = useRef(null), raf = useRef(null);
  useEffect(() => {
    const cv = ref.current; if (!cv) return; const ctx = cv.getContext("2d");
    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    const fit = () => { const w = cv.offsetWidth, h = cv.offsetHeight; if (cv.width !== Math.round(w * dpr)) { cv.width = w * dpr; cv.height = h * dpr; ctx.setTransform(dpr, 0, 0, dpr, 0, 0); } };
    // Real case: 1220×2440 sheet, 6 mm kerf (gaps visible), 8 parts in vivid colors,
    // leftover = one big smooth L (right strip + bottom strip) ≈ 40% of the sheet ≈ 56 kg.
    const parts = [
      [.012, .018, .235, .350, "#38bdf8"], [.258, .018, .235, .350, "#a78bfa"],
      [.525, .018, .223, .230, "#fb923c"], [.525, .270, .223, .230, "#f472b6"],
      [.012, .390, .160, .382, "#2dd4d0"], [.183, .390, .160, .382, "#fbbf24"],
      [.354, .390, .160, .382, "#60a5fa"], [.525, .522, .223, .250, "#f87171"],
    ];
    const Lpoly = [[.762, 0], [1, 0], [1, 1], [0, 1], [0, .788], [.762, .788]];
    const per = 450, nP = parts.length, tCut = nP * per + 250, tL = tCut + 650, P = 9200;
    let start = performance.now();
    const render = now => {
      fit();
      const CW = cv.offsetWidth, CH = cv.offsetHeight;
      const tt = (now - start) % P;
      const pad = 10, sx = pad, sy = pad, sw = CW - pad * 2, sh = CH - pad * 2 - 18;
      ctx.clearRect(0, 0, CW, CH);
      // sheet
      ctx.fillStyle = "rgba(12,17,25,.92)"; ctx.fillRect(sx, sy, sw, sh);
      ctx.strokeStyle = "rgba(245,158,11,.7)"; ctx.lineWidth = 1.4; ctx.strokeRect(sx, sy, sw, sh);
      ctx.strokeStyle = "rgba(148,163,184,.06)"; ctx.lineWidth = .5;
      for (let g = 1; g < 12; g++) { const gx = sx + g * sw / 12; ctx.beginPath(); ctx.moveTo(gx, sy); ctx.lineTo(gx, sy + sh); ctx.stroke(); }
      for (let g = 1; g < 6; g++) { const gy = sy + g * sh / 6; ctx.beginPath(); ctx.moveTo(sx, gy); ctx.lineTo(sx + sw, gy); ctx.stroke(); }
      // parts nest in, kerf gaps clearly visible
      parts.forEach((p, i) => {
        const t0 = i * per; if (tt < t0) return;
        const k = Math.min(1, (tt - t0) / 280), ease = 1 - Math.pow(1 - k, 3);
        const [px, py, pw, ph, col] = p;
        const x = sx + px * sw, y = sy + py * sh, w = pw * sw, h = ph * sh;
        const cx2 = x + w / 2, cy2 = y + h / 2, w2 = w * (.6 + .4 * ease), h2 = h * (.6 + .4 * ease);
        ctx.globalAlpha = ease;
        ctx.fillStyle = col + "70"; ctx.fillRect(cx2 - w2 / 2, cy2 - h2 / 2, w2, h2);
        ctx.strokeStyle = col; ctx.lineWidth = 1.4; ctx.strokeRect(cx2 - w2 / 2, cy2 - h2 / 2, w2, h2);
        if (k < 1) { ctx.strokeStyle = "rgba(255,255,255," + (.7 * (1 - k)).toFixed(2) + ")"; ctx.lineWidth = 2; ctx.strokeRect(cx2 - w2 / 2, cy2 - h2 / 2, w2, h2); }
        ctx.globalAlpha = 1;
      });
      // cutting-head sweep before the reveal
      if (tt > tCut && tt < tL) {
        const k = (tt - tCut) / (tL - tCut), lx = sx + k * sw;
        const lg = ctx.createLinearGradient(lx - 26, 0, lx, 0);
        lg.addColorStop(0, "rgba(245,158,11,0)"); lg.addColorStop(1, "rgba(255,200,90,.55)");
        ctx.fillStyle = lg; ctx.fillRect(lx - 26, sy, 26, sh);
        ctx.strokeStyle = "#ffd87a"; ctx.lineWidth = 1.6; ctx.beginPath(); ctx.moveTo(lx, sy); ctx.lineTo(lx, sy + sh); ctx.stroke();
      }
      // THE MONEY SHOT: the big smooth L-offcut, glowing green
      if (tt > tL) {
        const k = Math.min(1, (tt - tL) / 450), pulse = .72 + .28 * Math.sin((tt - tL) / 320);
        ctx.globalAlpha = k;
        ctx.beginPath(); Lpoly.forEach(([qx, qy], i) => i === 0 ? ctx.moveTo(sx + qx * sw, sy + qy * sh) : ctx.lineTo(sx + qx * sw, sy + qy * sh)); ctx.closePath();
        ctx.save(); ctx.shadowColor = "rgba(16,185,129,.8)"; ctx.shadowBlur = 18 * pulse;
        ctx.fillStyle = `rgba(16,185,129,${(.26 * pulse).toFixed(3)})`; ctx.fill(); ctx.restore();
        ctx.strokeStyle = "rgba(52,211,153,.95)"; ctx.lineWidth = 1.8; ctx.setLineDash([6, 4]); ctx.stroke(); ctx.setLineDash([]);
        // labels: the saving is the message
        const rcx = sx + .881 * sw;
        ctx.textAlign = "center"; ctx.fillStyle = "#a7f3d0"; ctx.font = "700 10px 'Space Mono', monospace";
        ctx.fillText(Tref.current("canvLoffcut"), rcx, sy + .30 * sh);
        ctx.fillStyle = "#34d399"; ctx.font = "800 15px 'Space Mono', monospace";
        ctx.fillText("≈ 56 " + Tref.current("kg"), rcx, sy + .30 * sh + 20);
        ctx.fillStyle = "#a7f3d0"; ctx.font = "700 10px 'Space Mono', monospace";
        ctx.fillText(Tref.current("scSaved"), rcx, sy + .30 * sh + 36);
        ctx.fillStyle = "rgba(167,243,208,.85)"; ctx.font = "700 10px 'Space Mono', monospace";
        ctx.fillText(Tref.current("scYours"), sx + .38 * sw, sy + .91 * sh);
        ctx.globalAlpha = 1;
      }
      if (tt > P - 600) { ctx.fillStyle = `rgba(7,10,15,${(((tt - (P - 600)) / 600) * .96).toFixed(3)})`; ctx.fillRect(0, 0, CW, CH); }
      // caption strip
      ctx.textAlign = "left"; ctx.fillStyle = "#94a3b8"; ctx.font = "10px 'Space Mono', monospace";
      ctx.fillText(Tref.current("scCaption"), sx, CH - 4);
      ctx.textAlign = "right";
      if (tt > tL) { ctx.fillStyle = "#34d399"; ctx.fillText(Tref.current("scKept"), sx + sw, CH - 4); }
      else { ctx.fillStyle = "#fbbf24"; ctx.fillText(Tref.current("scNesting", { p: Math.round(52 * Math.min(1, tt / (nP * per))) }), sx + sw, CH - 4); }
      raf.current = requestAnimationFrame(render);
    };
    const reduced = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (reduced) { start = performance.now() - (tL + 900); render(performance.now()); cancelAnimationFrame(raf.current); }
    else raf.current = requestAnimationFrame(render);
    return () => cancelAnimationFrame(raf.current);
  }, []);
  const chips = [
    ["📥", T("chipImport")], ["🔩", T("chipDb")],
    ["♻", T("chipOffcut")],
    ["⚖", T("chipTon")], ["📄", T("chipReports")],
  ];
  return (
    <div style={{ position: "relative", maxWidth: 980, margin: "0 auto", padding: "34px 20px 4px", textAlign: "center" }}>
      <h2 style={{ margin: "0 0 6px", fontFamily: "'Playfair Display', Georgia, serif", fontWeight: 900, fontSize: "clamp(26px,3.8vw,42px)", color: "#f8fafc" }}>{T("showTitle")}</h2>
      <div style={{ width: "min(720px, 94vw)", margin: "0 auto 18px", padding: "10px 12px 6px", borderRadius: 16, background: "rgba(13,18,26,.6)", backdropFilter: "blur(12px)", WebkitBackdropFilter: "blur(12px)", border: "1px solid rgba(245,158,11,.28)", boxShadow: "0 18px 60px -18px rgba(0,0,0,.65)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <span style={{ fontSize: 12, letterSpacing: 2.5, color: "rgba(251,191,36,.8)", fontFamily: "'Space Mono', monospace" }}>{T("liveNesting")}</span>
        </div>
        <canvas ref={ref} style={{ width: "100%", height: "min(42vw, 330px)", display: "block" }} />
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, justifyContent: "center", maxWidth: 860, margin: "0 auto" }}>
        {chips.map(([ic, txt]) => (
          <span key={txt} style={{ display: "inline-flex", alignItems: "center", gap: 7, padding: "6px 14px", borderRadius: 22, background: "rgba(19,25,32,.7)", backdropFilter: "blur(8px)", WebkitBackdropFilter: "blur(8px)", border: "1px solid rgba(148,163,184,.18)", fontSize: 14, color: "#cbd5e1", fontFamily: "'Space Mono', monospace" }}><span>{ic}</span>{txt}</span>
        ))}
      </div>
    </div>
  );
}

function Cover3D({ onStart }) {
  const T = useT();
  const ref = useRef(null), raf = useRef(null);
  useEffect(() => {
    const canvas = ref.current; if (!canvas) return; const ctx = canvas.getContext("2d");
    let t = 0;
    const resize = () => { const dpr = Math.min(window.devicePixelRatio || 1, 2); canvas.width = canvas.offsetWidth * dpr; canvas.height = canvas.offsetHeight * dpr; ctx.setTransform(dpr, 0, 0, dpr, 0, 0); };
    resize(); window.addEventListener("resize", resize);
    // rotate around Y only (pure turntable, no tilt); keep z for depth sorting
    const rot = (x, y, z, ry) => { const c = Math.cos(ry), s = Math.sin(ry); return { x: x * c + z * s, y, z: -x * s + z * c }; };
    const proj = (p, cx, cy) => { const f = 1200, s = f / (f + p.z + 460); return { x: cx + p.x * s, y: cy + p.y * s, z: p.z }; };
    const beamProfile = [[-1.1,-1.5],[1.1,-1.5],[1.1,-1.15],[0.15,-1.15],[0.15,1.15],[1.1,1.15],[1.1,1.5],[-1.1,1.5],[-1.1,1.15],[-0.15,1.15],[-0.15,-1.15],[-1.1,-1.15]];
    const render = () => {
      const W = canvas.offsetWidth, H = canvas.offsetHeight;
      ctx.clearRect(0, 0, W, H);
      // deep backdrop + warm overhead spotlight
      const bg = ctx.createLinearGradient(0, 0, 0, H); bg.addColorStop(0, "#0c1018"); bg.addColorStop(.55, "#10151d"); bg.addColorStop(1, "#070a0f"); ctx.fillStyle = bg; ctx.fillRect(0, 0, W, H);
      const cone = ctx.createRadialGradient(W * .5, -H * .25, 0, W * .5, -H * .25, H * 1.25);
      cone.addColorStop(0, "rgba(255,196,90,.16)"); cone.addColorStop(.55, "rgba(245,158,11,.05)"); cone.addColorStop(1, "rgba(0,0,0,0)");
      ctx.fillStyle = cone; ctx.fillRect(0, 0, W, H);
      ctx.strokeStyle = "rgba(245,158,11,.045)"; ctx.lineWidth = 1; const gs = 46;
      for (let x = 0; x < W; x += gs) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, H); ctx.stroke(); }
      for (let y = 0; y < H; y += gs) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(W, y); ctx.stroke(); }
      // rising forge embers (deterministic flicker)
      for (let i = 0; i < 38; i++) {
        const seed = i * 137.5, px = (seed + Math.sin(t * .22 + i * 1.7) * 24 + W) % W;
        const py = H - ((t * (14 + (i % 5) * 7) + seed * 3.1) % (H + 60)) + 30;
        const tw = Math.sin(t * 2.2 + i * 2.3) * .5 + .5, r = .8 + (i % 3) * .5;
        ctx.fillStyle = `rgba(252,${165 + (i % 3) * 25},60,${(tw * .4).toFixed(3)})`;
        ctx.beginPath(); ctx.arc(px, py, r, 0, 7); ctx.fill();
      }
      const ry = t * 0.5; // pure Y-axis spin, no tilt
      const bcx = W * 0.5, bcy = H * 0.47, sc = Math.min(W, H) * 0.115, depth = 9;
      const F = beamProfile.map(([x, y]) => rot(x * sc, y * sc, depth * sc * .5, ry));
      const Bk = beamProfile.map(([x, y]) => rot(x * sc, y * sc, -depth * sc * .5, ry));
      const drawBeam = (cy, mirror) => {
        const f2 = F.map(p => proj(p, bcx, cy)), b2 = Bk.map(p => proj(p, bcx, cy));
        const faces = [];
        for (let i = 0; i < beamProfile.length; i++) { const j = (i + 1) % beamProfile.length; faces.push({ kind: "side", pts: [f2[i], f2[j], b2[j], b2[i]], z: (f2[i].z + f2[j].z + b2[j].z + b2[i].z) / 4 }); }
        const zF = f2.reduce((s, p) => s + p.z, 0) / f2.length, zB = b2.reduce((s, p) => s + p.z, 0) / b2.length;
        faces.push({ kind: "cap", pts: f2, z: zF, near: zF <= zB });
        faces.push({ kind: "cap", pts: b2, z: zB, near: zB < zF });
        faces.sort((a, b) => b.z - a.z); // painter's algorithm: far first
        const lum = 0.5 + 0.5 * Math.abs(Math.cos(ry));
        faces.forEach(fc => {
          ctx.beginPath(); fc.pts.forEach((p, k) => k === 0 ? ctx.moveTo(p.x, p.y) : ctx.lineTo(p.x, p.y)); ctx.closePath();
          if (fc.kind === "side") {
            const base = [104, 112, 124].map(ch => Math.min(255, Math.round(ch * (0.5 + lum * 0.7))));
            const hi = [186, 196, 208].map(ch => Math.min(255, Math.round(ch * (0.55 + lum * 0.55))));
            const g = ctx.createLinearGradient(fc.pts[0].x, fc.pts[0].y, fc.pts[2].x, fc.pts[2].y);
            g.addColorStop(0, `rgb(${base.join(",")})`); g.addColorStop(.5, `rgb(${hi.join(",")})`); g.addColorStop(1, `rgb(${base.map(c => Math.round(c * .6)).join(",")})`);
            ctx.fillStyle = g; ctx.fill(); ctx.strokeStyle = "rgba(16,22,30,.5)"; ctx.lineWidth = .6; ctx.stroke();
          } else if (!fc.near) {
            ctx.fillStyle = `rgb(${Math.round(70 * lum + 30)},${Math.round(76 * lum + 32)},${Math.round(86 * lum + 36)})`; ctx.fill();
          } else {
            // near cap: brushed steel + rotation-tied specular sweep + amber rim
            ctx.save(); ctx.clip();
            const fg = ctx.createLinearGradient(0, cy - sc * 1.7, 0, cy + sc * 1.7);
            fg.addColorStop(0, "#d4dce5"); fg.addColorStop(.3, "#f0f4f8"); fg.addColorStop(.5, "#b6c0cc"); fg.addColorStop(.72, "#e0e6ed"); fg.addColorStop(1, "#96a1ae");
            ctx.fillStyle = fg; ctx.fillRect(bcx - sc * 6, cy - sc * 2, sc * 12, sc * 4);
            ctx.globalAlpha = .12; ctx.strokeStyle = "#fff"; ctx.lineWidth = .6;
            for (let k = 0; k < 36; k++) { const yy = cy - sc * 1.6 + k * (sc * 3.2 / 36); ctx.beginPath(); ctx.moveTo(bcx - sc * 5, yy); ctx.lineTo(bcx + sc * 5, yy + Math.sin(k * 12.9898) * 1.2); ctx.stroke(); }
            ctx.globalAlpha = 1;
            const sweep = Math.sin(ry) * .5 + .5;
            const sp = ctx.createLinearGradient(bcx - sc * 1.6, 0, bcx + sc * 1.6, 0);
            sp.addColorStop(Math.max(0, sweep - .2), "rgba(255,255,255,0)"); sp.addColorStop(sweep, "rgba(255,244,214,.6)"); sp.addColorStop(Math.min(1, sweep + .2), "rgba(255,255,255,0)");
            ctx.fillStyle = sp; ctx.fillRect(bcx - sc * 6, cy - sc * 2, sc * 12, sc * 4);
            ctx.restore();
            if (!mirror) { ctx.strokeStyle = "rgba(255,214,130,.3)"; ctx.lineWidth = 4.5; ctx.stroke(); }
            ctx.strokeStyle = mirror ? "rgba(245,158,11,.35)" : "#f59e0b"; ctx.lineWidth = 1.7; ctx.stroke();
          }
        });
      };
      // floor shadow anchors the beam
      const allPts = [...F, ...Bk].map(p => proj(p, bcx, bcy));
      const minX = Math.min(...allPts.map(p => p.x)), maxX = Math.max(...allPts.map(p => p.x));
      const floorY = bcy + sc * 2.35;
      const shW = (maxX - minX) * .58 + sc * .4;
      const sh = ctx.createRadialGradient(bcx, floorY, 0, bcx, floorY, shW);
      sh.addColorStop(0, "rgba(0,0,0,.55)"); sh.addColorStop(.6, "rgba(0,0,0,.25)"); sh.addColorStop(1, "rgba(0,0,0,0)");
      ctx.save(); ctx.translate(bcx, floorY); ctx.scale(1, .14); ctx.translate(-bcx, -floorY);
      ctx.fillStyle = sh; ctx.beginPath(); ctx.arc(bcx, floorY, shW, 0, 7); ctx.fill(); ctx.restore();
      // ground reflection: same projected beam, mirrored below the floor line, faded out
      ctx.save(); ctx.translate(0, 2 * floorY + sc * .3); ctx.scale(1, -1); ctx.globalAlpha = .16; drawBeam(bcy, true); ctx.restore();
      const fade = ctx.createLinearGradient(0, floorY, 0, floorY + sc * 2.4);
      fade.addColorStop(0, "rgba(12,16,24,.25)"); fade.addColorStop(1, "rgba(7,10,15,1)");
      ctx.fillStyle = fade; ctx.fillRect(0, floorY, W, sc * 2.6);
      // the beam
      drawBeam(bcy, false);
      t += 0.016; raf.current = requestAnimationFrame(render);
    };
    const reduced = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (reduced) { t = 0.9; render(); cancelAnimationFrame(raf.current); }
    else render();
    return () => { cancelAnimationFrame(raf.current); window.removeEventListener("resize", resize); };
  }, []);
  return (
    <div style={{ position: "relative", width: "100%", height: "min(82vh, 720px)", minHeight: 520, overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <canvas ref={ref} style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }} />
      <div style={{ position: "absolute", inset: 0, background: "radial-gradient(ellipse 52% 44% at 50% 50%, rgba(7,10,15,.55) 0%, rgba(7,10,15,.22) 48%, rgba(7,10,15,0) 74%)", zIndex: 1, pointerEvents: "none" }} />
      <div style={{ position: "relative", zIndex: 2, textAlign: "center", padding: "0 24px", pointerEvents: "none" }}>
        <div style={{ display: "inline-block", background: "rgba(245,158,11,.12)", border: "1px solid rgba(245,158,11,.4)", borderRadius: 30, padding: "4px 16px", marginBottom: 14, fontFamily: "'Space Mono', monospace", fontSize: 13, letterSpacing: 3, color: "#fbbf24", textTransform: "uppercase", backdropFilter: "blur(6px)", WebkitBackdropFilter: "blur(6px)" }}>{T("heroBadge")}</div>
        <h1 style={{ margin: "0 0 10px", lineHeight: 1.04, fontFamily: "'Playfair Display', Georgia, serif", fontWeight: 900, fontSize: "clamp(36px,6.4vw,76px)", color: "#f8fafc", textShadow: "0 4px 30px rgba(0,0,0,.65)" }}>{T("heroTitleA")} <span style={{ background: "linear-gradient(120deg,#fbbf24,#f59e0b 55%,#fb923c)", WebkitBackgroundClip: "text", backgroundClip: "text", color: "transparent" }}>{T("heroTitleB")}</span></h1>
        <p style={{ margin: "0 0 26px", color: "#cbd5e1", fontSize: "clamp(15px,2.1vw,20px)", fontFamily: "'Playfair Display', Georgia, serif", fontWeight: 600 }}>{T("heroSub")}</p>
        <button onClick={onStart} style={{ pointerEvents: "auto", background: "linear-gradient(135deg,#fbbf24,#d97706)", border: "none", color: "#1a1206", padding: "13px 42px", borderRadius: 10, fontSize: 17, fontWeight: 800, cursor: "pointer", letterSpacing: 1, fontFamily: "'Space Mono', monospace", animation: "ctaPulse 3.2s ease-in-out infinite" }}>{T("heroCta")}</button>
      </div>
      <div style={{ position: "absolute", bottom: 14, left: "50%", color: "rgba(245,200,120,.55)", fontSize: 13, letterSpacing: 2, fontFamily: "'Space Mono', monospace", textTransform: "uppercase", zIndex: 2, animation: "hintBounce 2.4s ease-in-out infinite" }}>{T("heroScroll")}</div>
    </div>
  );
}

/* ─── PLATE NESTING CANVAS ───────────────────────────────────────────────── */
function Mini({ label, value, accent }) { return <div style={{ textAlign: "center" }}><div style={{ fontSize: 12, color: "#64748b", marginBottom: 2, letterSpacing: 1 }}>{label}</div><div style={{ fontSize: 16, color: accent || "#cbd5e1", fontWeight: 700 }}>{value}</div></div>; }

/* ─── SECTION CUT BAR ────────────────────────────────────────────────────── */

/* ─── L-SHAPE INTERACTIVE DIAGRAM (labeled, eliminates ambiguity) ─────────── */
function LShapeDiagram({ A, B, C, D, t }) {
  const TT = useT();
  const W = 260, H = 240, pad = 46;
  const a = Math.max(A, 1), b = Math.max(B, 1);
  const isL = C > 0 && D > 0;
  const c = isL ? Math.min(C, a - 1) : 0, d = isL ? Math.min(D, b - 1) : 0;
  const sc = Math.min((W - pad * 2) / a, (H - pad * 2) / b);
  const ox = pad, oy = pad, aw = a * sc, bh = b * sc, cw = c * sc, dh = d * sc;
  const pts = isL ? [[ox, oy + dh], [ox + aw - cw, oy + dh], [ox + aw - cw, oy], [ox + aw, oy], [ox + aw, oy + bh], [ox, oy + bh]] : [[ox, oy], [ox + aw, oy], [ox + aw, oy + bh], [ox, oy + bh]];
  const usable = lUsableRect(a, b, c, d);
  let ux, uy, uw, uh;
  if (!isL) { ux = ox; uy = oy; uw = aw; uh = bh; }
  else if (usable.w === a) { ux = ox; uy = oy + dh; uw = aw; uh = (b - d) * sc; }
  else { ux = ox; uy = oy; uw = (a - c) * sc; uh = bh; }
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", maxWidth: 300, background: "linear-gradient(160deg, rgba(13,18,24,.9), rgba(10,14,20,.6))", borderRadius: 12, border: "1px solid rgba(148,163,184,.16)" }}>
      <defs><pattern id="usableHatch" width="7" height="7" patternTransform="rotate(45)" patternUnits="userSpaceOnUse"><line x1="0" y1="0" x2="0" y2="7" stroke="rgba(16,185,129,.4)" strokeWidth="1.5" /></pattern></defs>
      <polygon points={pts.map(p => p.join(",")).join(" ")} fill="rgba(245,158,11,.14)" stroke="#f59e0b" strokeWidth="1.6" strokeLinejoin="round" />
      <rect x={ux} y={uy} width={uw} height={uh} fill="url(#usableHatch)" stroke="#10b981" strokeWidth="1.6" strokeDasharray="6 3" rx="2" />
      <text x={ux + uw / 2} y={uy + uh / 2 + 3} fill="#6ee7b7" fontSize="10" fontWeight="700" textAnchor="middle" fontFamily="'Space Mono', monospace">{TT("usableWord")}</text>
      {/* A — bottom edge */}
      <text x={ox + aw / 2} y={oy + bh + 20} fill="#fcd34d" fontSize="11" fontWeight="700" textAnchor="middle" fontFamily="'Space Mono', monospace">{Math.round(A)} {TT("widthArrow")}</text>
      {/* B — left edge */}
      <text x={ox - 14} y={oy + bh / 2} fill="#fcd34d" fontSize="11" fontWeight="700" textAnchor="middle" fontFamily="'Space Mono', monospace" transform={`rotate(-90 ${ox - 14} ${oy + bh / 2})`}>{Math.round(B)} {TT("lengthArrow")}</text>
      {isL && <>
        {/* C — along the top, above the cut-out span only */}
        <text x={ox + aw - cw / 2} y={oy - 8} fill="#93c5fd" fontSize="10" fontWeight="700" textAnchor="middle" fontFamily="'Space Mono', monospace">{Math.round(C)}</text>
        <line x1={ox + aw - cw} y1={oy - 4} x2={ox + aw} y2={oy - 4} stroke="#93c5fd" strokeWidth="1" />
        {/* D — down the right edge, beside the cut-out span only, offset clear of C */}
        <text x={ox + aw + 16} y={oy + dh / 2} fill="#93c5fd" fontSize="10" fontWeight="700" textAnchor="middle" fontFamily="'Space Mono', monospace" transform={`rotate(90 ${ox + aw + 16} ${oy + dh / 2})`}>{Math.round(D)}</text>
        <line x1={ox + aw + 4} y1={oy} x2={ox + aw + 4} y2={oy + dh} stroke="#93c5fd" strokeWidth="1" />
        <text x={ox + aw - cw / 2} y={oy + dh / 2 + 3} fill="#64748b" fontSize="8" textAnchor="middle" fontFamily="'Space Mono', monospace">{TT("cutoutWord")}</text>
      </>}
      <text x={W / 2} y={H - 8} fill="#94a3b8" fontSize="9" textAnchor="middle" fontFamily="'Space Mono', monospace">{TT("lThickNote", { t: t })}</text>
    </svg>
  );
}

/* ─── L-SHAPE POINT BUILDER ────────────────────────────────────────────────
   The easy way to enter an L-shaped offcut: stand at one corner (0,0), walk
   the outline, type each corner you reach. The picture redraws after every
   number so you always see exactly what you're describing. PLATES ONLY. */
function LPointBuilder({ points, onChange }) {
  const TT = useT();
  const pts = points && points.length === 6 ? points : defaultLPoints();
  const info = analyzeL(pts);
  const W = 340, H = 320, pad = 52;
  const maxX = Math.max(...pts.map(p => p.x), 1), maxY = Math.max(...pts.map(p => p.y), 1);
  const sc = Math.min((W - pad * 2) / maxX, (H - pad * 2) / maxY);
  const ox = pad, oy = H - pad; // origin at bottom-left, y grows UP
  const SX = x => ox + x * sc;
  const SY = y => oy - y * sc;
  const setPt = (i, axis, v) => { const next = pts.map((p, pi) => pi === i ? { ...p, [axis]: Math.max(0, Math.round(+v || 0)) } : p); onChange(next, analyzeL(next)); };
  // usable rectangle position for shading
  let uRect = null;
  if (info.A > 0 && info.B > 0) {
    const fullH = info.usable.h === info.B, fullW = info.usable.w === info.A;
    if (fullH) { const onLeft = info.corner === "BR" || info.corner === "TR"; uRect = { x: onLeft ? 0 : info.C, y: 0, w: info.A - info.C, h: info.B }; }
    else if (fullW) { const onBottom = info.corner === "TL" || info.corner === "TR"; uRect = { x: 0, y: onBottom ? 0 : info.D, w: info.A, h: info.B - info.D }; }
  }
  const corners = ["①", "②", "③", "④", "⑤", "⑥"];
  const hints = [TT("lpHint0"), TT("lpHint1"), TT("lpHint2"), TT("lpHint3"), TT("lpHint4"), TT("lpHint5")];
  return (
    <div style={{ display: "flex", gap: 18, flexWrap: "wrap", alignItems: "flex-start" }}>
      <div style={{ flex: "1 1 220px", minWidth: 210 }}>
        <div style={{ fontSize: 14, color: "#6ee7b7", lineHeight: 1.6, marginBottom: 12, background: "rgba(16,185,129,.08)", border: "1px solid rgba(16,185,129,.25)", borderRadius: 8, padding: "9px 12px" }}>{TT("lpIntro")}</div>
        <table style={{ borderCollapse: "separate", borderSpacing: "0 7px" }}>
          <thead><tr><th></th><th style={{ fontSize: 11, color: "#64748b", letterSpacing: 1, textAlign: "left", paddingLeft: 6 }}>{TT("lpColX")}</th><th style={{ fontSize: 11, color: "#64748b", letterSpacing: 1, textAlign: "left", paddingLeft: 6 }}>{TT("lpColY")}</th></tr></thead>
          <tbody>
            {pts.map((p, i) => (
              <tr key={i}>
                <td style={{ paddingRight: 6, whiteSpace: "nowrap" }}>
                  <span style={{ fontSize: 17, color: "#fbbf24", fontWeight: 700 }}>{corners[i]}</span>
                </td>
                <td><input type="number" value={p.x} disabled={i === 0} onChange={e => setPt(i, "x", e.target.value)} style={{ ...CI, width: 78, opacity: i === 0 ? .5 : 1 }} /></td>
                <td><input type="number" value={p.y} disabled={i === 0} onChange={e => setPt(i, "y", e.target.value)} style={{ ...CI, width: 78, opacity: i === 0 ? .5 : 1 }} /></td>
                <td style={{ paddingLeft: 8, fontSize: 12, color: "#475569", fontFamily: "'Space Mono', monospace", whiteSpace: "nowrap" }}>{hints[i]}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ marginTop: 12, display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button onClick={() => { const d = defaultLPoints(); onChange(d, analyzeL(d)); }} style={{ ...CI, cursor: "pointer", borderColor: "#334155", padding: "6px 12px" }}>{TT("lpReset")}</button>
        </div>
      </div>
      <div style={{ flex: "0 0 auto" }}>
        <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", maxWidth: 360, background: "linear-gradient(160deg, rgba(13,18,24,.92), rgba(10,14,20,.6))", borderRadius: 12, border: "1px solid rgba(148,163,184,.16)" }}>
          <defs><pattern id="lbHatch" width="7" height="7" patternTransform="rotate(45)" patternUnits="userSpaceOnUse"><line x1="0" y1="0" x2="0" y2="7" stroke="rgba(16,185,129,.45)" strokeWidth="1.5" /></pattern></defs>
          {/* axes */}
          <line x1={ox} y1={oy} x2={W - 14} y2={oy} stroke="#f59e0b" strokeWidth="1.4" />
          <line x1={ox} y1={oy} x2={ox} y2={14} stroke="#f59e0b" strokeWidth="1.4" />
          <polygon points={`${W - 14},${oy} ${W - 21},${oy - 4} ${W - 21},${oy + 4}`} fill="#f59e0b" />
          <polygon points={`${ox},${14} ${ox - 4},${21} ${ox + 4},${21}`} fill="#f59e0b" />
          <text x={W - 18} y={oy + 16} fill="#fbbf24" fontSize="9" textAnchor="end" fontFamily="'Space Mono', monospace">{TT("lpAxisX")}</text>
          <text x={ox - 8} y={20} fill="#fbbf24" fontSize="9" textAnchor="start" fontFamily="'Space Mono', monospace" transform={`rotate(-90 ${ox - 8} 20)`}>{TT("lpAxisY")}</text>
          {/* usable rectangle */}
          {uRect && uRect.w > 0 && uRect.h > 0 && (
            <rect x={SX(uRect.x)} y={SY(uRect.y + uRect.h)} width={uRect.w * sc} height={uRect.h * sc} fill="url(#lbHatch)" stroke="#10b981" strokeWidth="1.4" strokeDasharray="6 3" rx="2" />
          )}
          {/* the outline so far */}
          <polygon points={pts.map(p => `${SX(p.x)},${SY(p.y)}`).join(" ")} fill="rgba(245,158,11,.12)" stroke="#f59e0b" strokeWidth="1.8" strokeLinejoin="round" />
          {/* corner dots + numbers */}
          {pts.map((p, i) => (
            <g key={i}>
              <circle cx={SX(p.x)} cy={SY(p.y)} r="9" fill="#0c1016" stroke="#fbbf24" strokeWidth="1.6" />
              <text x={SX(p.x)} y={SY(p.y) + 3.5} fill="#fbbf24" fontSize="9" fontWeight="700" textAnchor="middle" fontFamily="'Space Mono', monospace">{i + 1}</text>
            </g>
          ))}
          {uRect && uRect.w > 0 && uRect.h > 0 && (
            <text x={SX(uRect.x + uRect.w / 2)} y={SY(uRect.y + uRect.h / 2)} fill="#6ee7b7" fontSize="10" fontWeight="700" textAnchor="middle" fontFamily="'Space Mono', monospace">{TT("usableWord")}</text>
          )}
        </svg>
        <div style={{ marginTop: 8, textAlign: "center", fontSize: 13, color: info.valid ? "#6ee7b7" : "#fbbf24", fontFamily: "'Space Mono', monospace" }}>
          {TT("lpOverall")} <b>{Math.round(info.A)} × {Math.round(info.B)}</b> · {TT("lpUsable")} <b style={{ color: "#6ee7b7" }}>{Math.round(info.usable.w)} × {Math.round(info.usable.h)}</b>
          {!info.valid && <div style={{ color: "#fbbf24", marginTop: 3 }}>{TT("lpTip")}</div>}
        </div>
      </div>
    </div>
  );
}

/* ─── REUSE-A-LEFTOVER-FILE BUTTON (shared by Plates + Sections) ──────────── */

/* ════════════════════════════════════════════════════════════════════════
   MAIN APP
════════════════════════════════════════════════════════════════════════ */
/* ════════════════════════════════════════════════════════════════════════════
   EMAIL GATE — waitlist capture at the download moment (highest-intent action)
   ───────────────────────────────────────────────────────────────────────────
   HOW IT WORKS:
   • On-screen results (nesting, waste %, procurement summary) stay 100% free.
   • The FIRST time someone clicks any Download button, this asks for an email.
   • The email is POSTed to your Formspree inbox, then the file downloads.
   • It's remembered in the browser, so returning users are NEVER asked again.

   ► TO ACTIVATE: create a free form at https://formspree.io , copy the form ID
     from your endpoint URL (https://formspree.io/f/XXXXXXXX  →  the XXXXXXXX),
     and paste it below. Until you do, downloads still work (stored locally only).
══════════════════════════════════════════════════════════════════════════════ */
const FORMSPREE_ID = "xqerwkze"; // ◄── live Formspree form: https://formspree.io/f/xqerwkze
if (FORMSPREE_ID === "YOUR_FORM_ID" && typeof console !== "undefined") {
  console.warn("[steel-optimizer] FORMSPREE_ID is still the placeholder — emails are NOT being received. Nothing is sent anywhere.");
}
const GATE_KEY = "steelopt_email_v2";

// module-level bridge so any download button can open the one shared modal
const gateBus = { fn: null, open: null };
function safeExport(action) {
  // Exports build large HTML/XLSX strings and could throw on odd data. A throw
  // here is invisible to the React error boundary (event-handler context), so
  // catch it and tell the user in both languages instead of failing silently.
  try { action(); }
  catch (err) {
    console.error("Export failed:", err);
    try { alert(soTb("MISC.reportFail", "The report could not be generated. Please try again.", "تعذّر إنشاء التقرير. حاول مرة أخرى.")); } catch { /* no-op */ }
  }
}

function requestDownload(action) {
  track("download_results", { action: String(action || "") });
  const guarded = () => safeExport(action);
  let saved = null;
  try { saved = localStorage.getItem(GATE_KEY); } catch { saved = "skip"; }
  if (saved) { guarded(); return; }         // already captured → download immediately
  gateBus.fn = guarded;
  if (gateBus.open) gateBus.open();          // open modal, download runs after submit
  else guarded();                            // safety net: never block a download
}



/* ─── SECTION LIBRARY (browse families → sections → data card) ───────────── */
const SECTION_DIMS = {
  "127x76x13 UB":[127,76,4,7.6],"152x152x23 UC":[152.4,152.2,5.8,6.8],"152x152x30 UC":[157.6,152.9,6.5,9.4],"152x152x37 UC":[161.8,154.4,8,11.5],
  "152x152x44 UC":[166,155.9,9.5,13.6],"152x89x16 UB":[152.4,88.7,4.5,7.7],"178x102x19 UB":[177.8,101.2,4.8,7.9],"203x102x23 UB":[203.2,101.8,5.4,9.3],
  "203x133x25 UB":[203.2,133.2,5.7,7.8],"203x133x30 UB":[206.8,133.9,6.4,9.6],"203x203x100 UC":[228.6,210.3,14.5,23.7],"203x203x46 UC":[203.2,203.6,7.2,11],
  "203x203x52 UC":[206.2,204.3,7.9,12.5],"203x203x60 UC":[209.6,205.8,9.4,14.2],"203x203x71 UC":[215.8,206.4,10,17.3],"203x203x86 UC":[222.2,209.1,12.7,20.5],
  "254x102x22 UB":[254,101.6,5.7,6.8],"254x102x25 UB":[257.2,101.9,6,8.4],"254x102x28 UB":[260.4,102.2,6.3,10],"254x146x31 UB":[251.4,146.1,6,8.6],
  "254x146x37 UB":[256,146.4,6.3,10.9],"254x146x43 UB":[259.6,147.3,7.2,12.7],"254x254x107 UC":[266.7,258.8,12.8,20.5],"254x254x132 UC":[276.3,261.3,15.3,25.3],
  "254x254x167 UC":[289.1,265.2,19.2,31.7],"254x254x73 UC":[254.1,254.6,8.6,14.2],"254x254x89 UC":[260.3,256.3,10.3,17.3],"305x102x25 UB":[305.1,101.6,5.8,7],
  "305x102x28 UB":[308.7,101.8,6,8.8],"305x102x33 UB":[312.7,102.4,6.6,10.8],"305x127x37 UB":[304.4,123.4,7.1,10.7],"305x127x42 UB":[307.2,124.3,8,12.1],
  "305x127x48 UB":[311,125.3,9,14],"305x165x40 UB":[303.4,165,6,10.2],"305x165x46 UB":[306.6,165.7,6.7,11.8],"305x165x54 UB":[310.4,166.9,7.9,13.7],
  "305x305x118 UC":[314.5,307.4,12,18.7],"305x305x137 UC":[320.5,309.2,13.8,21.7],"305x305x158 UC":[327.1,311.2,15.8,25],"305x305x198 UC":[339.9,314.5,19.1,31.4],
  "305x305x240 UC":[352.5,318.4,23,37.7],"305x305x283 UC":[365.3,322.2,26.8,44.1],"305x305x97 UC":[307.9,305.3,9.9,15.4],"356x127x33 UB":[349,125.4,6,8.5],
  "356x127x39 UB":[353.4,126,6.6,10.7],"356x171x45 UB":[351.4,171.1,7,9.7],"356x171x51 UB":[355,171.5,7.4,11.5],"356x171x57 UB":[358,172.2,8.1,13],
  "356x171x67 UB":[363.4,173.2,9.1,15.7],"356x368x129 UC":[355.6,368.6,10.4,17.5],"356x368x153 UC":[362,370.5,12.3,20.7],"356x368x177 UC":[368.2,372.6,14.4,23.8],
  "356x368x202 UC":[374.6,374.7,16.5,27],"356x406x235 UC":[381,394.8,18.4,30.2],"356x406x287 UC":[393.6,399,22.6,36.5],"356x406x340 UC":[406.4,403,26.6,42.9],
  "356x406x393 UC":[419,407,30.6,49.2],"356x406x467 UC":[436.6,412.2,35.8,58],"356x406x551 UC":[455.6,418.5,42.1,67.5],"356x406x634 UC":[474.6,424,47.6,77],
  "406x140x39 UB":[398,141.8,6.4,8.6],"406x140x46 UB":[403.2,142.2,6.8,11.2],"406x178x54 UB":[402.6,177.7,7.7,10.9],"406x178x60 UB":[406.4,177.9,7.9,12.8],
  "406x178x67 UB":[409.4,178.8,8.8,14.3],"406x178x74 UB":[412.8,179.5,9.5,16],"457x152x52 UB":[449.8,152.4,7.6,10.9],"457x152x60 UB":[454.6,152.9,8.1,13.3],
  "457x152x67 UB":[458,153.8,9,15],"457x152x74 UB":[462,154.4,9.6,17],"457x152x82 UB":[465.8,155.3,10.5,18.9],"457x191x67 UB":[453.4,189.9,8.5,12.7],
  "457x191x74 UB":[457,190.4,9,14.5],"457x191x82 UB":[460,191.3,9.9,16],"457x191x89 UB":[463.4,191.9,10.5,17.7],"457x191x98 UB":[467.2,192.8,11.4,19.6],
  "533x210x101 UB":[536.7,210,10.8,17.4],"533x210x109 UB":[539.5,210.8,11.6,18.8],"533x210x122 UB":[544.5,211.9,12.7,21.3],"533x210x82 UB":[528.3,208.8,9.6,13.2],
  "533x210x92 UB":[533.1,209.3,10.1,15.6],"610x229x101 UB":[602.6,227.6,10.5,14.8],"610x229x113 UB":[607.6,228.2,11.1,17.3],"610x229x125 UB":[612.2,229,11.9,19.6],
  "610x229x140 UB":[617.2,230.2,13.1,22.1],"610x305x149 UB":[612.4,304.8,11.8,19.7],"610x305x179 UB":[620.2,307.1,14.1,23.6],"610x305x238 UB":[635.8,311.4,18.4,31.4],
  "686x254x125 UB":[677.9,253,11.7,16.2],"686x254x140 UB":[683.5,253.7,12.4,19],"686x254x152 UB":[687.5,254.5,13.2,21],"686x254x170 UB":[692.9,255.8,14.5,23.7],
  "762x267x134 UB":[750,264.4,12,15.5],"762x267x147 UB":[754,265.2,12.8,17.5],"762x267x173 UB":[762.2,266.7,14.3,21.6],"762x267x197 UB":[769.8,268,15.6,25.4],
  "838x292x176 UB":[834.9,291.7,14,18.8],"838x292x194 UB":[840.7,292.4,14.7,21.7],"838x292x226 UB":[850.9,293.8,16.1,26.8],"914x305x201 UB":[903,303.3,15.1,20.2],
  "914x305x224 UB":[910.4,304.1,15.9,23.9],"914x305x253 UB":[918.4,305.5,17.3,27.9],"914x305x289 UB":[926.6,307.7,19.5,32],"C 100x50x5x7.5":[100,50,5,7.5],
  "C 125x65x6x8":[125,65,6,8],"C 150x75x6.5x10":[150,75,6.5,10],"C 150x75x9x12.5":[150,75,9,12.5],"C 180x75x7x10.5":[180,75,7,10.5],
  "C 200x80x7.5x11":[200,80,7.5,11],"C 200x90x8x13.5":[200,90,8,13.5],"C 250x90x11x14.5":[250,90,11,14.5],"C 250x90x9x13":[250,90,9,13],
  "C 300x90x10x15.5":[300,90,10,15.5],"C 300x90x12x16":[300,90,12,16],"C 300x90x9x13":[300,90,9,13],"C 380x100x10.5x16":[380,100,10.5,16],
  "C 380x100x13x16.5":[380,100,13,16.5],"C 380x100x13x20":[380,100,13,20],"C 75x40x5x7":[75,40,5,7],"C10x15.3":[254,66,6.1,11.07],
  "C10x20":[254,69.6,9.63,11.07],"C10x25":[254,73.4,13.36,11.07],"C10x30":[254,77,17.09,11.07],"C12x20.7":[304.8,74.7,7.16,12.73],
  "C12x25":[304.8,77.5,9.83,12.73],"C12x30":[304.8,80.5,12.95,12.73],"C15x33.9":[381,86.4,10.16,16.51],"C15x40":[381,89.4,13.21,16.51],
  "C15x50":[381,94.5,18.19,16.51],"C3x4.1":[76.2,35.8,4.32,6.93],"C3x5":[76.2,38.1,6.55,6.93],"C3x6":[76.2,40.6,9.04,6.93],
  "C4x5.4":[101.6,40.1,4.67,7.52],"C4x7.25":[101.6,43.7,8.15,7.52],"C5x6.7":[127,44.4,4.83,8.13],"C5x9":[127,48,8.25,8.13],
  "C6x10.5":[152.4,51.6,7.98,8.71],"C6x13":[152.4,54.9,11.1,8.71],"C6x8.2":[152.4,48.8,5.08,8.71],"C7x12.25":[177.8,55.6,7.98,9.3],
  "C7x14.75":[177.8,58.4,10.64,9.3],"C7x9.8":[177.8,53.1,5.33,9.3],"C8x11.5":[203.2,57.4,5.59,9.91],"C8x13.75":[203.2,59.4,7.7,9.91],
  "C8x18.75":[203.2,64.3,12.37,9.91],"C9x13.4":[228.6,61.7,5.92,10.49],"C9x15":[228.6,63.2,7.24,10.49],"C9x20":[228.6,67.3,11.38,10.49],
  "HEA 100":[96,100,5,8],"HEA 1000":[990,300,16.5,31],"HEA 120":[114,120,5,8],"HEA 140":[133,140,5.5,8.5],
  "HEA 160":[152,160,6,9],"HEA 180":[171,180,6,9.5],"HEA 200":[190,200,6.5,10],"HEA 220":[210,220,7,11],
  "HEA 240":[230,240,7.5,12],"HEA 260":[250,260,7.5,12.5],"HEA 280":[270,280,8,13],"HEA 300":[290,300,8.5,14],
  "HEA 320":[310,300,9,15.5],"HEA 340":[330,300,9.5,16.5],"HEA 360":[350,300,10,17.5],"HEA 400":[390,300,11,19],
  "HEA 450":[440,300,11.5,21],"HEA 500":[490,300,12,23],"HEA 550":[540,300,12.5,24],"HEA 600":[590,300,13,25],
  "HEA 650":[640,300,13.5,26],"HEA 700":[690,300,14.5,27],"HEA 800":[790,300,15,28],"HEA 900":[890,300,16,30],
  "HEB 100":[100,100,6,10],"HEB 1000":[1000,300,19,36],"HEB 120":[120,120,6.5,11],"HEB 140":[140,140,7,12],
  "HEB 160":[160,160,8,13],"HEB 180":[180,180,8.5,14],"HEB 200":[200,200,9,15],"HEB 220":[220,220,9.5,16],
  "HEB 240":[240,240,10,17],"HEB 260":[260,260,10,17.5],"HEB 280":[280,280,10.5,18],"HEB 300":[300,300,11,19],
  "HEB 320":[320,300,11.5,20.5],"HEB 340":[340,300,12,21.5],"HEB 360":[360,300,12.5,22.5],"HEB 400":[400,300,13.5,24],
  "HEB 450":[450,300,14,26],"HEB 500":[500,300,14.5,28],"HEB 550":[550,300,15,29],"HEB 600":[600,300,15.5,30],
  "HEB 650":[650,300,16,31],"HEB 700":[700,300,17,32],"HEB 800":[800,300,17.5,33],"HEB 900":[900,300,18.5,35],
  "HEM 100":[120,106,12,20],"HEM 1000":[1008,302,21,40],"HEM 120":[140,126,12.5,21],"HEM 140":[160,146,13,22],
  "HEM 160":[180,166,14,23],"HEM 180":[200,186,14.5,24],"HEM 200":[220,206,15,25],"HEM 220":[240,226,15.5,26],
  "HEM 240":[270,248,18,32],"HEM 260":[290,268,18,32.5],"HEM 280":[310,288,18.5,33],"HEM 300":[340,310,21,39],
  "HEM 320":[359,309,21,40],"HEM 340":[377,309,21,40],"HEM 360":[395,308,21,40],"HEM 400":[432,307,21,40],
  "HEM 450":[478,307,21,40],"HEM 500":[524,306,21,40],"HEM 550":[572,306,21,40],"HEM 600":[620,305,21,40],
  "HEM 650":[668,305,21,40],"HEM 700":[716,304,21,40],"HEM 800":[814,303,21,40],"HEM 900":[910,302,21,40],
  "HM 150x100":[148,100,6,9],"HM 200x150":[194,150,6,9],"HM 250x175":[244,175,7,11],"HM 300x200":[294,200,8,12],
  "HM 350x250":[340,250,9,14],"HM 400x300":[390,300,10,16],"HM 450x300":[440,300,11,18],"HM 600x302":[594,302,14,23],
  "HN 100x50":[100,50,5,7],"HN 125x60":[125,60,6,8],"HN 150x75":[150,75,5,7],"HN 175x90":[175,90,5,8],
  "HN 198x99":[198,99,4.5,7],"HN 200x100":[200,100,5.5,8],"HN 248x124":[248,124,5,8],"HN 250x125":[250,125,6,9],
  "HN 298x149":[298,149,5.5,8],"HN 300x150":[300,150,6.5,9],"HN 346x174":[346,174,6,9],"HN 350x175":[350,175,7,11],
  "HN 396x199":[396,199,7,11],"HN 400x200":[400,200,8,13],"HN 446x199":[446,199,8,12],"HN 450x200":[450,200,9,14],
  "HN 496x199":[496,199,9,14],"HN 500x200":[500,200,10,16],"HN 596x199":[596,199,10,15],"HN 606x201":[606,201,12,20],
  "HN 692x300":[692,300,13,20],"HN 700x300":[692,300,13,20],"HN 792x300":[792,300,14,22],"HN 800x300":[792,300,14,22],
  "HN 892x299":[892,299,15,23],"HN 900x300":[892,299,15,23],"HW 100x100":[100,100,6,8],"HW 125x125":[125,125,6.5,9],
  "HW 150x150":[150,150,7,10],"HW 175x175":[175,175,7.5,11],"HW 200x200":[200,200,8,12],"HW 200x204":[200,204,12,12],
  "HW 250x250":[250,250,9,14],"HW 250x255":[250,255,14,14],"HW 300x300":[300,300,10,15],"HW 300x305":[300,305,15,15],
  "HW 350x350":[350,350,12,19],"HW 400x400":[400,400,13,21],"HW 400x408":[400,408,21,21],"HW 414x405":[414,405,18,28],
  "HW 428x407":[428,407,20,35],"HW 458x417":[458,417,30,50],"HW 498x432":[498,432,45,70],"I 200x100x7x10":[200,100,7,10],
  "I 200x150x9x16":[200,150,9,16],"I 250x125x10x19":[250,125,10,19],"I 250x125x7.5x12.5":[250,125,7.5,12.5],"I 300x150x10x18.5":[300,150,10,18.5],
  "I 300x150x11.5x22":[300,150,11.5,22],"I 300x150x8x13":[300,150,8,13],"I 350x150x12x24":[350,150,12,24],"I 350x150x9x15":[350,150,9,15],
  "I 400x150x10x18":[400,150,10,18],"I 400x150x12.5x25":[400,150,12.5,25],"I 450x175x11x20":[450,175,11,20],"I 450x175x13x26":[450,175,13,26],
  "I 600x190x13x25":[600,190,13,25],"I 600x190x16x35":[600,190,16,35],"IPE 100":[100,55,4.1,5.7],"IPE 120":[120,64,4.4,6.3],
  "IPE 140":[140,73,4.7,6.9],"IPE 160":[160,82,5,7.4],"IPE 180":[180,91,5.3,8],"IPE 200":[200,100,5.6,8.5],
  "IPE 220":[220,110,5.9,9.2],"IPE 240":[240,120,6.2,9.8],"IPE 270":[270,135,6.6,10.2],"IPE 300":[300,150,7.1,10.7],
  "IPE 330":[330,160,7.5,11.5],"IPE 360":[360,170,8,12.7],"IPE 400":[400,180,8.6,13.5],"IPE 450":[450,190,9.4,14.6],
  "IPE 500":[500,200,10.2,16],"IPE 550":[550,210,11.1,17.2],"IPE 600":[600,220,12,19],"IPE 80":[80,46,3.8,5.2],
  "MC10x22":[254,84.3,7.37,14.6],"MC10x25":[254,86.6,9.65,14.6],"MC12x31":[304.8,93.2,9.4,17.78],"MC12x35":[304.8,95.8,11.81,17.78],
  "MC18x42.7":[457.2,100.3,11.43,15.88],"MC18x58":[457.2,106.7,17.78,15.88],"MC6x12":[152.4,63.5,7.87,9.52],"MC6x15.1":[152.4,74.7,8.03,12.06],
  "MC8x18.7":[203.2,75.7,8.97,12.7],"MC8x21.4":[203.2,87.6,9.52,13.33],"UPE 100":[100,55,4.5,7.5],"UPE 120":[120,60,5,8],
  "UPE 140":[140,65,5,9],"UPE 160":[160,70,5.5,9.5],"UPE 180":[180,75,5.5,10.5],"UPE 200":[200,80,6,11],
  "UPE 220":[220,85,6.5,12],"UPE 240":[240,90,7,12.5],"UPE 270":[270,95,7.5,13.5],"UPE 300":[300,100,9.5,15],
  "UPE 330":[330,105,11,16],"UPE 360":[360,110,12,17],"UPE 400":[400,115,13.5,18],"UPE 80":[80,50,4,7],
  "UPN 100":[100,50,6,8.5],"UPN 120":[120,55,7,9],"UPN 140":[140,60,7,10],"UPN 160":[160,65,7.5,10.5],
  "UPN 180":[180,70,8,11],"UPN 200":[200,75,8.5,11.5],"UPN 220":[220,80,9,12.5],"UPN 240":[240,85,9.5,13],
  "UPN 260":[260,90,10,14],"UPN 280":[280,95,10,15],"UPN 300":[300,100,10,16],"UPN 320":[320,100,14,17.5],
  "UPN 350":[350,100,14,16],"UPN 380":[380,102,13.5,16],"UPN 400":[400,110,14,18],"UPN 80":[80,45,6,8],
  "W10x100":[281.9,261.6,17.27,28.45],"W10x112":[289.6,264.2,19.18,31.75],"W10x12":[250.7,100.6,4.83,5.33],"W10x15":[253.7,101.6,5.84,6.86],
  "W10x17":[256.5,101.9,6.1,8.38],"W10x19":[259.1,102.1,6.35,10.03],"W10x22":[259.1,146,6.1,9.14],"W10x26":[261.6,146.6,6.6,11.18],
  "W10x30":[266.7,147.6,7.62,12.95],"W10x33":[247.1,202.2,7.37,11.05],"W10x39":[252,202.9,8,13.46],"W10x45":[256.5,203.7,8.89,15.75],
  "W10x49":[254,254,8.64,14.22],"W10x54":[256.5,254,9.4,15.62],"W10x60":[259.1,256.5,10.67,17.27],"W10x68":[264.2,256.5,11.94,19.56],
  "W10x77":[269.2,259.1,13.46,22.1],"W10x88":[274.3,261.6,15.37,25.15],"W12x106":[327.7,309.9,15.49,25.15],"W12x120":[332.7,312.4,18.03,28.19],
  "W12x136":[340.4,315,20.07,31.75],"W12x14":[302.3,100.8,5.08,5.71],"W12x152":[348,317.5,22.1,35.56],"W12x16":[304.8,101.3,5.59,6.73],
  "W12x170":[355.6,320,24.38,39.62],"W12x19":[309.9,101.9,5.97,8.89],"W12x190":[365.8,322.6,26.92,44.2],"W12x210":[373.4,325.1,29.97,48.26],
  "W12x22":[312.4,102.4,6.6,10.79],"W12x230":[383.5,327.7,32.77,52.58],"W12x252":[391.2,330.2,35.56,57.15],"W12x26":[309.9,164.8,5.84,9.65],
  "W12x279":[403.9,332.7,38.86,62.74],"W12x30":[312.4,165.6,6.6,11.18],"W12x305":[414,335.3,41.4,68.83],"W12x336":[426.7,340.4,45.21,75.18],
  "W12x35":[317.5,166.6,7.62,13.21],"W12x40":[302.3,203.5,7.49,13.08],"W12x45":[307.3,204.5,8.51,14.6],"W12x50":[309.9,205.2,9.4,16.26],
  "W12x53":[307.3,254,8.76,14.6],"W12x58":[309.9,254,9.14,16.26],"W12x65":[307.3,304.8,9.91,15.37],"W12x72":[312.4,304.8,10.92,17.02],
  "W12x79":[315,307.3,11.94,18.67],"W12x87":[317.5,307.3,13.08,20.57],"W12x96":[322.6,309.9,13.97,22.86],"W14x109":[363.2,370.8,13.33,21.84],
  "W14x120":[368.3,373.4,14.99,23.88],"W14x132":[373.4,373.4,16.38,26.16],"W14x145":[375.9,393.7,17.27,27.69],"W14x159":[381,396.2,18.92,30.23],
  "W14x176":[386.1,398.8,21.08,33.27],"W14x193":[393.7,398.8,22.61,36.58],"W14x211":[398.8,401.3,24.89,39.62],"W14x22":[348,127,5.84,8.51],
  "W14x233":[406.4,403.9,27.18,43.69],"W14x257":[416.6,406.4,29.97,48.01],"W14x26":[353.1,127.8,6.48,10.67],"W14x283":[424.2,408.9,32.77,52.58],
  "W14x30":[350.5,170.9,6.86,9.78],"W14x311":[434.3,411.5,35.81,57.4],"W14x34":[355.6,171.4,7.24,11.56],"W14x342":[444.5,416.6,39.12,62.74],
  "W14x370":[454.7,419.1,42.16,67.56],"W14x38":[358.1,172,7.87,13.08],"W14x398":[464.8,421.6,44.96,72.39],"W14x426":[475,424.2,47.75,77.22],
  "W14x43":[348,203.2,7.75,13.46],"W14x455":[482.6,426.7,51.31,81.53],"W14x48":[350.5,204,8.64,15.11],"W14x500":[497.8,431.8,55.63,88.9],
  "W14x53":[353.1,204.7,9.4,16.76],"W14x550":[513.1,436.9,60.45,97.03],"W14x605":[530.9,442,66.04,105.66],"W14x61":[353.1,254,9.52,16.38],
  "W14x665":[548.6,449.6,71.88,114.81],"W14x68":[355.6,254,10.54,18.29],"W14x730":[569,454.7,77.98,124.71],"W14x74":[360.7,256.5,11.43,19.94],
  "W14x82":[363.2,256.5,12.95,21.72],"W14x90":[355.6,368.3,11.18,18.03],"W14x99":[360.7,370.8,12.32,19.81],"W16x100":[431.8,264.2,14.86,25.02],
  "W16x26":[398.8,139.7,6.35,8.76],"W16x31":[403.9,140.5,6.99,11.18],"W16x36":[403.9,177.5,7.49,10.92],"W16x40":[406.4,177.8,7.75,12.83],
  "W16x45":[408.9,178.8,8.76,14.35],"W16x50":[414,179.6,9.65,16],"W16x57":[416.6,180.8,10.92,18.16],"W16x67":[414,259.1,10.03,16.89],
  "W16x77":[419.1,261.6,11.56,19.3],"W16x89":[426.7,264.2,13.33,22.22],"W18x106":[475,284.5,14.99,23.88],"W18x119":[482.6,287,16.64,26.92],
  "W18x130":[490.2,284.5,17.02,30.48],"W18x143":[495.3,284.5,18.54,33.53],"W18x158":[500.4,287,20.57,36.58],"W18x35":[449.6,152.4,7.62,10.79],
  "W18x40":[454.7,152.9,8,13.33],"W18x46":[459.7,153.9,9.14,15.37],"W18x50":[457.2,190.5,9.02,14.48],"W18x55":[459.7,191.3,9.91,16],
  "W18x60":[462.3,192,10.54,17.65],"W18x65":[467.4,192.8,11.43,19.05],"W18x71":[469.9,194.1,12.57,20.57],"W18x76":[462.3,279.4,10.79,17.27],
  "W18x86":[467.4,281.9,12.19,19.56],"W18x97":[472.4,281.9,13.59,22.1],"W21x101":[543.6,312.4,12.7,20.32],"W21x111":[546.1,312.4,13.97,22.22],
  "W21x122":[551.2,315,15.24,24.38],"W21x132":[553.7,315,16.51,26.42],"W21x147":[561.3,317.5,18.29,29.21],"W21x44":[525.8,165.1,8.89,11.43],
  "W21x50":[528.3,165.9,9.65,13.59],"W21x57":[535.9,166.6,10.29,16.51],"W21x62":[533.4,209.3,10.16,15.62],"W21x68":[535.9,210.1,10.92,17.4],
  "W21x73":[538.5,210.8,11.56,18.8],"W21x83":[543.6,212.3,13.08,21.21],"W21x93":[548.6,213.9,14.73,23.62],"W24x103":[622.3,228.6,13.97,24.89],
  "W24x104":[612.1,325.1,12.7,19.05],"W24x117":[617.2,325.1,13.97,21.59],"W24x131":[622.3,327.7,15.37,24.38],"W24x146":[627.4,327.7,16.51,27.69],
  "W24x162":[635,330.2,17.91,30.99],"W24x55":[599.4,178.1,10.03,12.83],"W24x62":[602,178.8,10.92,14.99],"W24x68":[602,227.8,10.54,14.86],
  "W24x76":[607.1,228.3,11.18,17.27],"W24x84":[612.1,229.1,11.94,19.56],"W24x94":[617.2,230.4,13.08,22.22],"W27x102":[688.3,254,13.08,21.08],
  "W27x114":[693.4,256.5,14.48,23.62],"W27x129":[701,254,15.49,27.94],"W27x146":[696,355.6,15.37,24.76],"W27x161":[701,355.6,16.76,27.43],
  "W27x178":[706.1,358.1,18.41,30.23],"W27x84":[678.2,254,11.68,16.26],"W27x94":[683.3,254,12.45,18.92],"W30x108":[756.9,266.7,13.84,19.3],
  "W30x116":[762,266.7,14.35,21.59],"W30x124":[767.1,266.7,14.86,23.62],"W30x132":[769.6,266.7,15.62,25.4],"W30x148":[779.8,266.7,16.51,29.97],
  "W30x173":[772.2,381,16.64,27.18],"W30x191":[779.8,381,18.03,30.23],"W30x211":[784.9,383.5,19.68,33.53],"W30x90":[749.3,264.2,11.94,15.49],
  "W30x99":[754.4,266.7,13.21,17.02],"W33x118":[835.7,292.1,13.97,18.8],"W33x130":[840.7,292.1,14.73,21.72],"W33x141":[845.8,292.1,15.37,24.38],
  "W33x152":[850.9,294.6,16.13,26.92],"W33x169":[858.5,292.1,17.02,30.99],"W33x201":[856,398.8,18.16,29.21],"W33x221":[861.1,401.3,19.68,32.51],
  "W33x241":[868.7,403.9,21.08,35.56],"W36x135":[904.2,304.8,15.24,20.07],"W36x150":[911.9,304.8,15.88,23.88],"W36x160":[914.4,304.8,16.51,25.91],
  "W36x170":[919.5,304.8,17.27,27.94],"W36x182":[922,307.3,18.41,29.97],"W36x194":[927.1,307.3,19.43,32],"W36x210":[932.2,309.9,21.08,34.54],
  "W36x231":[927.1,419.1,19.3,32],"W36x232":[942.3,307.3,22.1,39.88],"W36x247":[932.2,419.1,20.32,34.29],"W36x256":[950,309.9,24.38,43.94],
  "W36x262":[937.3,421.6,21.34,36.58],"W36x282":[942.3,421.6,22.48,39.88],"W36x302":[947.4,424.2,24,42.67],"W36x361":[965.2,424.2,28.45,51.05],
  "W6x12":[153.2,101.6,5.84,7.11],"W6x15":[152.1,152.1,5.84,6.6],"W6x16":[159.5,102.4,6.6,10.29],"W6x20":[157.5,152.9,6.6,9.27],
  "W6x25":[162.1,154.4,8.13,11.56],"W6x9":[149.9,100.1,4.32,5.46],"W8x10":[200.4,100.1,4.32,5.21],"W8x13":[202.9,101.6,5.84,6.48],
  "W8x15":[206,102.1,6.22,8],"W8x18":[206.8,133.3,5.84,8.38],"W8x21":[210.3,133.9,6.35,10.16],"W8x24":[201.4,165.1,6.22,10.16],
  "W8x28":[204.7,166.1,7.24,11.81],"W8x31":[203.2,203.2,7.24,11.05],"W8x35":[206.2,203.7,7.87,12.57],"W8x40":[209.5,205,9.14,14.22],
  "W8x48":[215.9,206,10.16,17.4],"W8x58":[222.2,208.8,12.95,20.57],"W8x67":[228.6,210.3,14.48,23.75]
};
const SL_DIM_SRC = { en: "Dimensions: AISC Shapes DB v16.0 \u00b7 EN 10365 \u00b7 British Steel (BS EN 10365:2017) \u00b7 JIS G 3192 / KS / GB series. Every value cross-checked against unit mass; formula verified against 379 published section areas (max error 0.75%).", ar: "\u0627\u0644\u0623\u0628\u0639\u0627\u062f: AISC v16.0 \u00b7 EN 10365 \u00b7 British Steel \u00b7 JIS G 3192 / KS / GB. \u062a\u0645 \u0627\u0644\u062a\u062d\u0642\u0642 \u0645\u0646 \u0643\u0644 \u0642\u064a\u0645\u0629 \u0645\u0642\u0627\u0628\u0644 \u0627\u0644\u0648\u0632\u0646\u060c \u0648\u0645\u0646 \u0627\u0644\u0645\u0639\u0627\u062f\u0644\u0629 \u0645\u0642\u0627\u0628\u0644 379 \u0645\u0633\u0627\u062d\u0629 \u0645\u0642\u0637\u0639 \u0645\u0646\u0634\u0648\u0631\u0629." };

const SL_TXT = {
  btn:      { en: "Section Library", ar: "مكتبة المقاطع" },
  title:    { en: "Steel Section Library", ar: "مكتبة المقاطع الحديدية" },
  pickFam:  { en: "Choose a section family", ar: "اختر عائلة المقاطع" },
  search:   { en: "Search sections…", ar: "ابحث عن مقطع…" },
  back:     { en: "Back", ar: "رجوع" },
  close:    { en: "Close", ar: "إغلاق" },
  sections: { en: "sections", ar: "مقطع" },
  designation: { en: "Designation", ar: "التسمية" },
  family:   { en: "Family", ar: "العائلة" },
  unitMass: { en: "Unit mass", ar: "الوزن للمتر" },
  area:     { en: "Cross-section area", ar: "مساحة المقطع" },
  perBar:   { en: "Mass per stock bar", ar: "وزن العود الواحد" },
  perTonne: { en: "Bars per tonne", ar: "عدد الأعواد بالطن" },
  barsOf:   { en: "bars of", ar: "عود بطول" },
  note:     { en: "Area derived from unit mass at 7850 kg/m³. Use the optimizer for cutting plans and tonnage.", ar: "المساحة محسوبة من الوزن عند كثافة 7850 كجم/م³. استخدم البرنامج لخطط القص والتونية." },
  useIt:    { en: "Optimize with this profile ↓", ar: "استخدم هذا المقطع في البرنامج ↓" },
};
const SL_FAM_LABEL = {
  "IPE": { en: "IPE — European I-beams", ar: "IPE — كمرات أوروبية" },
  "HEA": { en: "HEA — Wide flange (light)", ar: "HEA — عريضة الشفة (خفيفة)" },
  "HEB": { en: "HEB — Wide flange (medium)", ar: "HEB — عريضة الشفة (متوسطة)" },
  "HEM": { en: "HEM — Wide flange (heavy)", ar: "HEM — عريضة الشفة (ثقيلة)" },
  "IPN": { en: "IPN — Tapered I-beams", ar: "IPN — كمرات مائلة الشفة" },
  "UB":  { en: "UB — UK beams", ar: "UB — كمرات بريطانية" },
  "UC":  { en: "UC — UK columns", ar: "UC — أعمدة بريطانية" },
  "UBP": { en: "UBP — UK bearing piles", ar: "UBP — خوازيق بريطانية" },
  "PFC": { en: "PFC — UK channels", ar: "PFC — قنوات بريطانية" },
  "UPN": { en: "UPN — Tapered channels", ar: "UPN — قنوات مائلة الشفة" },
  "UPE": { en: "UPE — Parallel channels", ar: "UPE — قنوات متوازية الشفة" },
  "W":   { en: "W — American wide flange", ar: "W — أمريكية عريضة الشفة" },
  "C-AMER": { en: "C — American channels", ar: "C — قنوات أمريكية" },
  "MC":  { en: "MC — American misc. channels", ar: "MC — قنوات أمريكية متنوعة" },
  "L":   { en: "L — Angles", ar: "L — زوايا" },
  "SHS": { en: "SHS — Square hollow", ar: "SHS — مربعة مجوفة" },
  "RHS": { en: "RHS — Rectangular hollow", ar: "RHS — مستطيلة مجوفة" },
  "CHS": { en: "CHS — Circular hollow", ar: "CHS — دائرية مجوفة" },
  "C-COLD": { en: "C — Cold-formed purlins", ar: "C — مدادات مشكلة على البارد" },
  "Z-COLD": { en: "Z — Cold-formed purlins", ar: "Z — مدادات مشكلة على البارد" },
  "JIS-HW": { en: "HW — Wide-flange H (GB/T 11263 · KS D 3502)", ar: "HW — كمرات H عريضة (صيني/كوري)" },
  "JIS-HM": { en: "HM — Medium-flange H (GB/T 11263 · KS D 3502)", ar: "HM — كمرات H متوسطة (صيني/كوري)" },
  "JIS-HN": { en: "HN — Narrow-flange H (GB/T 11263 · KS D 3502)", ar: "HN — كمرات H ضيقة (صيني/كوري)" },
  "JIS-I":  { en: "I — Tapered-flange I-beams (JIS G 3192)", ar: "I — كمرات يابانية مائلة الشفة (JIS G 3192)" },
  "JIS-C":  { en: "C — Tapered-flange channels (JIS G 3192)", ar: "C — قنوات يابانية مائلة الشفة (JIS G 3192)" },
  "JIS-LIP":{ en: "LC — Cold-formed lipped channels (JIS G 3350)", ar: "LC — قنوات مشكّلة على البارد (JIS G 3350)" },
  "GOST-I": { en: "GOST I-beams", ar: "GOST — كمرات روسية" },
  "GOST-C": { en: "GOST channels", ar: "GOST — قنوات روسية" },
  "GOST-B": { en: "GOST Б wide flange", ar: "GOST Б — عريضة الشفة" },
  "GOST-K": { en: "GOST К columns", ar: "GOST К — أعمدة" },
  "GB-I":   { en: "GB I-beams (China)", ar: "GB — كمرات صينية" },
  "GB-C":   { en: "GB channels (China)", ar: "GB — قنوات صينية" },
  "HSS":    { en: "HSS — American hollow structural sections", ar: "HSS — مقاطع مجوّفة أمريكية" },
  "HP":     { en: "HP — American bearing piles", ar: "HP — خوازيق أمريكية" },
  "S-AMER": { en: "S — American standard I-beams", ar: "S — كمرات أمريكية قياسية" },
  "M-AMER": { en: "M — American miscellaneous I-beams", ar: "M — كمرات أمريكية متنوعة" },
};
/* ╔══════════════════════════════════════════════════════════════════════════╗
   ║  PAINT AREA + SECTION PAGES — ISOLATED ADD-ON (everything lives here)     ║
   ╚══════════════════════════════════════════════════════════════════════════╝
   The optimisation engine, tonnage, waste and offcut logic are NOT touched.
   The rest of the app only calls the small helpers below; each one is wrapped
   so an error can never stop an optimisation or an export — it just shows "—".

   Switches:
     PAINT_AREA_ENABLED = false   → paint area disappears from results & reports
     SECTION_PAGE_LINKS = false   → no links to the /steel-sections pages
   Links also stay hidden automatically until the pages are actually deployed
   (the app checks that /steel-sections/s.css exists before showing them).
──────────────────────────────────────────────────────────────────────────── */
const PAINT_AREA_ENABLED = true;    // ◄ paint area on/off
const SECTION_PAGE_LINKS = true;    // ◄ links to /steel-sections on/off

/*@@SHARED-CORE:BEGIN@@*/
/* ─── PAINT AREA + SECTION PAGES — shared core ────────────────────────────────
   Plain JavaScript only (no JSX, no React). build-pages.mjs reads everything
   between the two @@SHARED-CORE markers at build time and uses it to generate
   the /steel-sections pages, so the website and the app always show the same
   numbers. Keep the markers.

   PAINT AREA RULE (conservative — never under-estimates):
     • outside faces only for closed sections (SHS/RHS/CHS/HSS/Pipe)
     • square corners, root radii and corner radii ignored → larger perimeter
     • rolled I/H and channels: 2h + 4b  (the exact sharp-corner value is
       2h + 4b − 2tw, so dropping −2tw adds a further 1–3 %)
     • angles: 2(a + b) (exact for sharp corners) · built-up: 2D + 4B − 2tw
     • no dimensions in the library → null (shown as "—"). Never guessed.
   Result is typically 2–5 % above published catalogue surface areas.
──────────────────────────────────────────────────────────────────────────── */

// ASME B36.10M outside diameter (mm) by DN, for AISC "PIPE{DN}STD/XS/XXS" labels.
const PIPE_OD_MM = { 15: 21.3, 20: 26.7, 25: 33.4, 32: 42.2, 40: 48.3, 50: 60.3, 65: 73.0, 80: 88.9, 90: 101.6, 100: 114.3, 125: 141.3, 150: 168.3, 200: 219.1, 250: 273.0, 300: 323.8, 350: 355.6, 400: 406.4, 450: 457.0, 500: 508.0, 600: 610.0, 650: 660.4 };

// AISC metric designation depth (mm) for each imperial nominal depth (in).
const AISC_DEPTH_MM = {
  W:  { 4: 100, 5: 130, 6: 150, 8: 200, 10: 250, 12: 310, 14: 360, 16: 410, 18: 460, 21: 530, 24: 610, 27: 690, 30: 760, 33: 840, 36: 920, 40: 1000, 44: 1100 },
  C:  { 3: 75, 4: 100, 5: 130, 6: 150, 7: 180, 8: 200, 9: 230, 10: 250, 12: 310, 15: 380 },
  MC: { 6: 150, 7: 180, 8: 200, 9: 230, 10: 250, 12: 310, 13: 330, 18: 460 },
};

/* Same physical section, two labels in the library:
     W310X39 ≡ W12x26 (AISC metric vs imperial) · UB127X76X13 ≡ 127x76x13 UB
   A metric AISC label is matched to its imperial twin only when exactly one
   imperial row has the same depth class AND the same weight (lb/ft × 1.48816
   within 1.5 %), one-to-one. Anything ambiguous stays unmatched. */
let _SO_ALIAS = null;
function sectionAliasOf(name) {
  if (!_SO_ALIAS) {
    _SO_ALIAS = {};
    const byName = {};
    STEEL_DB.forEach(s => { byName[s.name] = s; });
    const imp = {}, claim = {};
    STEEL_DB.forEach(s => {
      const m = s.name.match(/^(W|C|MC)(\d+)x(\d+(?:\.\d+)?)$/);
      if (m) (imp[m[1]] = imp[m[1]] || []).push({ name: s.name, d: +m[2], w: +m[3] });
    });
    STEEL_DB.forEach(s => {
      let m = s.name.match(/^(W|C|MC)(\d+)X(\d+(?:\.\d+)?)$/);
      if (m) {
        const map = AISC_DEPTH_MM[m[1]], dmm = +m[2], kg = +m[3];
        const hits = (imp[m[1]] || []).filter(i => map[i.d] === dmm && Math.abs(i.w * 1.48816 - kg) <= kg * 0.015);
        if (hits.length === 1) {                                     // one-to-one: keep the closer claimant
          const d = Math.abs(hits[0].w * 1.48816 - kg), prev = claim[hits[0].name];
          if (!prev || d < prev.d) { if (prev) delete _SO_ALIAS[prev.name]; claim[hits[0].name] = { name: s.name, d }; _SO_ALIAS[s.name] = hits[0].name; }
        }
        return;
      }
      m = s.name.match(/^(UB|UC|UBP)(\d+)X(\d+)X(\d+(?:\.\d+)?)$/);
      if (m) { const k = `${m[2]}x${m[3]}x${m[4]} ${m[1]}`; if (byName[k]) _SO_ALIAS[s.name] = k; }
    });
  }
  return _SO_ALIAS[name] || null;
}

// [h, b, tw, tf] in mm from the library (directly or through the twin label).
function sectionDims(name) {
  if (!name) return null;
  return SECTION_DIMS[name] || SECTION_DIMS[sectionAliasOf(name)] || null;
}

/* Paint area in m² per metre run, or null when the geometry is not known.
   Returns { m2pm, basis }  basis ∈ ih | box | round | angle | lipped | builtup */
function sectionPaintM2pm(sec) {
  if (!sec || !sec.name) return null;
  const name = String(sec.name).trim(), type = sec.type || "";
  const U = name.toUpperCase().replace(/×/g, "X").replace(/[*∗]/g, "X").replace(/\s+/g, "");
  const out = (mm, basis) => (mm > 0 && isFinite(mm) ? { m2pm: mm / 1000, basis } : null);
  const N = "(\\d+(?:\\.\\d+)?)";
  let m;
  if (sec.builtUp || type === "BUILT-UP") {                      // D x B x tf x tw
    m = U.match(new RegExp(`^(?:BU|PG)?${N}X${N}X${N}X${N}$`));
    return m ? out(2 * +m[1] + 4 * +m[2] - 2 * +m[4], "builtup") : null;
  }
  if (sec.tekla && sec.pm > 0) return out(sec.pm, sec.basis || "tekla");   // Tekla parametric C / ZZ: both faces of the wall
  if (sec.intl && sec.pm > 0) return out(sec.pm, sec.basis || "intl");   // international SAP2000 rows (own perimeter)
  const dm = sectionDims(name);
  if (dm) return out(2 * dm[0] + 4 * dm[1], "ih");
  if (type === "C-COLD" || type === "Z-COLD") return null;      // lip size not in the designation
  // Hollow sections: the perimeter implied by the library mass is also checked and the larger
  // value is used, so a label that is not "outside size × wall" can never under-state the area.
  const kgm = +sec.kgm || 0, A = kgm > 0 ? kgm / 7.85e-3 : 0;            // mm² at 7850 kg/m³
  if ((m = U.match(new RegExp(`^(?:SHS|RHS|HSS)${N}X${N}X${N}$`)))) { const t = +m[3]; return out(Math.max(2 * (+m[1] + +m[2]), A && t ? A / t + 4 * t : 0), "box"); }
  if ((m = U.match(new RegExp(`^(?:CHS|HSS)${N}X${N}$`)))) { const t = +m[2]; return out(Math.PI * Math.max(+m[1], A && t ? A / (Math.PI * t) + t : 0), "round"); }
  if ((m = U.match(/^PIPE(\d+)(?:STD|XS|XXS)$/)) && PIPE_OD_MM[+m[1]]) return out(Math.PI * PIPE_OD_MM[+m[1]], "round");
  if ((m = U.match(new RegExp(`^L${N}X${N}X${N}$`)))) return out(2 * (+m[1] + +m[2]), "angle");
  if (type === "PFC" && (m = U.match(new RegExp(`^${N}X${N}X${N}PFC$`)))) return out(2 * +m[1] + 4 * +m[2], "ih");
  if (/^JIS-(HW|HM|HN|I|C)$/.test(type) && (m = U.match(new RegExp(`^(?:HW|HM|HN|I|C)${N}X${N}`)))) return out(2 * +m[1] + 4 * +m[2], "ih");
  if (type === "JIS-LIP" && (m = U.match(new RegExp(`^LC${N}X${N}X${N}X${N}$`)))) return out(2 * +m[1] + 4 * +m[2] + 4 * +m[3], "lipped");
  return null;
}

// Display helpers — always round UP so the shown figure stays conservative.
const fmtM2pm = v => (Math.ceil(v * 1000 - 1e-9) / 1000).toFixed(3);
const fmtM2 = v => (v >= 1000 ? Math.ceil(v - 1e-9).toLocaleString("en-US") : (Math.ceil(v * 10 - 1e-9) / 10).toFixed(1));
const upM2 = (v, d) => Math.ceil(v * 10 ** d - 1e-9) / 10 ** d;          // numeric, for Excel cells

/* ─── section page addresses: /steel-sections/<family>/<section> ─────────── */
const SO_FAM_SLUG = {
  "IPE": "ipe", "HEA": "hea", "HEB": "heb", "HEM": "hem", "IPN": "ipn", "UB": "ub", "UC": "uc", "UBP": "ubp",
  "PFC": "pfc", "UPN": "upn", "UPE": "upe", "W": "w", "C-AMER": "c-american", "MC": "mc", "HP": "hp",
  "S-AMER": "s-american", "M-AMER": "m-american", "L": "angles", "SHS": "shs", "RHS": "rhs", "CHS": "chs",
  "HSS": "hss", "C-COLD": "c-cold-formed", "Z-COLD": "z-cold-formed", "JIS-HW": "hw", "JIS-HM": "hm",
  "JIS-HN": "hn", "JIS-I": "jis-i", "JIS-C": "jis-c", "JIS-LIP": "jis-lc", "GOST-I": "gost-i",
  "GOST-C": "gost-c", "GOST-B": "gost-b", "GOST-K": "gost-k", "GB-I": "gb-i", "GB-C": "gb-c",
};
function sectionFamSlug(type) { return SO_FAM_SLUG[type] || String(type || "other").toLowerCase().replace(/[^a-z0-9]+/g, "-"); }
function sectionSlug(name) {
  let s = String(name || "").trim();
  const uk = s.match(/^(\d[\d.xX×]*)\s+(UB|UC|UBP|PFC)$/i);            // 127x76x13 UB → UB 127x76x13
  if (uk) s = uk[2] + " " + uk[1];
  return s.toLowerCase().replace(/×/g, "x").replace(/\*/g, "x")
    .replace(/([a-z])\s+(?=[a-z])/g, "$1-")                           // GOST I20 → gost-i20
    .replace(/(\d)\.0+(?=\D|$)/g, "$1")                                // 5.0 → 5 (one page for CHS 114.3x5 / 114.3x5.0)
    .replace(/\s+/g, "").replace(/^([a-z]+)(?=\d)/, "$1-")
    .replace(/\./g, "-").replace(/[^a-z0-9x-]+/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
}
/* Library rows whose mass does not match their own label are kept off the public
   pages until the data is checked (the app still uses them for tonnage):
     • PIPE…STD/XS/XXS — mass on the AISC design wall (0.93 t), ≈ 7 % below nominal
     • round/box/angle rows whose mass ≠ the "size × wall" in the label          */
function sectionDataIssue(sec) {
  if (!sec || !sec.name) return "no name";
  const kgm = +sec.kgm; if (!(kgm > 0)) return "no mass";
  const U = String(sec.name).toUpperCase().replace(/×/g, "X").replace(/\s+/g, "");
  const N = "(\\d+(?:\\.\\d+)?)", rho = 7.85e-3;
  let m, r;
  if (/^PIPE\d+(?:STD|XS|XXS)$/.test(U)) return "pipe mass on AISC design wall (0.93 t)";
  if ((m = U.match(new RegExp(`^(?:CHS|HSS)${N}X${N}$`)))) { const D = +m[1], t = +m[2]; r = kgm / (Math.PI * (D - t) * t * rho); if (!(t < D / 2) || r < 0.97 || r > 1.03) return `mass is ${Math.round(r * 100)}% of OD × wall`; }
  else if ((m = U.match(new RegExp(`^(?:SHS|RHS|HSS)${N}X${N}X${N}$`)))) { const h = +m[1], b = +m[2], t = +m[3]; r = kgm / ((2 * (h + b) - 4 * t) * t * rho); if (!(2 * t < Math.min(h, b)) || r < 0.8 || r > 1.03) return `mass is ${Math.round(r * 100)}% of size × wall`; }
  else if ((m = U.match(new RegExp(`^L${N}X${N}X${N}$`)))) { const a = +m[1], b = +m[2], t = +m[3]; r = kgm / ((a + b - t) * t * rho); if (r < 0.97 || r > 1.03) return `mass is ${Math.round(r * 100)}% of legs × thickness`; }
  return null;
}

// The page a library row lives on (twin labels share one page). null = no page.
function sectionPagePath(sec, lang) {
  if (!sec || !sec.name || sec.computed || sec.builtUp) return null;
  const primary = sectionAliasOf(sec.name) || sec.name;
  const row = STEEL_DB.find(s => s.name === primary);
  if (!row || sectionDataIssue(row)) return null;
  return `${lang === "ar" ? "/ar" : ""}/steel-sections/${sectionFamSlug(row.type)}/${sectionSlug(row.name)}`;
}
/*@@SHARED-CORE:END@@*/

/* ─── safe wrappers used by the app ──────────────────────────────────────── */
function soSafe(fn, fallback = null) {
  try { const v = fn(); return (v == null || (typeof v === "number" && !isFinite(v))) ? fallback : v; }
  catch (e) { try { console.warn("[paint-area add-on]", e); } catch { /* no console */ } return fallback; }
}
function sectionPaintSafe(sec) { return PAINT_AREA_ENABLED ? soSafe(() => sectionPaintM2pm(sec)) : null; }
// Sections: net cut lengths of the user's own list × m²/m of the profile.
// Any error → paintErr, and the whole paint area is then hidden (never a partial total).
function sectionListPaint(sec, items) {
  const inputMm = soSafe(() => (items || []).reduce((s, c) => s + (+c.length || 0) * (+c.qty || 0), 0));
  const none = { inputMm, paintM2pm: null, paintM2: null };
  if (!PAINT_AREA_ENABLED) return none;
  return soSafe(() => {
    const pa = sec ? sectionPaintM2pm(sec) : null;
    return { inputMm, paintM2pm: pa ? pa.m2pm : null, paintM2: pa && inputMm != null ? (pa.m2pm * inputMm) / 1000 : null };
  }, { ...none, paintErr: true });
}
function sectionPaintTotals(groups) {
  const off = { paintM2: null, paintMissing: [] };
  if (!PAINT_AREA_ENABLED || groups.some(g => g.paintErr)) return off;
  return soSafe(() => ({
    paintM2: groups.reduce((s, g) => s + (g.paintM2 || 0), 0),
    paintMissing: groups.filter(g => g.paintM2 == null).map(g => g.profile),
  }), off);
}
// Plates: ONE face of the user's parts, length × width × qty (offcuts and scrap excluded).
function plateFacePaintM2(parts) {
  if (!PAINT_AREA_ENABLED) return null;
  return soSafe(() => (parts || []).reduce((s, p) => s + (+p.length || 0) * (+p.width || 0) * (+p.qty || 0), 0) / 1e6);
}
function addPaint(a, b) { return (a == null || b == null) ? null : a + b; }

/* ─── text (EN / AR) ─────────────────────────────────────────────────────── */
Object.assign(LANG_DICT, {
  paintHd:      { en: "🎨 PAINT AREA", ar: "🎨 مساحة الدهان" },
  paintNoteS:   { en: "Net cut lengths from your list × outside surface per metre. Square corners, root radii ignored → conservative (typically 2–5% above catalogue). Offcuts and scrap not included.",
                  ar: "أطوال القص الصافية من قائمتك × السطح الخارجي لكل متر. الزوايا محسوبة حادّة بدون أنصاف أقطار التدوير ← رقم متحفّظ (عادةً أعلى من الكتالوج بـ ٢–٥٪). البواقي والهالك غير محسوبة." },
  paintNoteP:   { en: "One face only · your parts (length × width × qty) · offcuts and scrap not included.",
                  ar: "وجه واحد فقط · القطع من قائمتك (الطول × العرض × العدد) · البواقي والهالك غير محسوبة." },
  paintMissing: { en: "Not included — no dimensions in the library: {list}", ar: "غير محسوبة — لا توجد أبعاد في المكتبة: {list}" },
  stPaint:      { en: "Paint Area", ar: "مساحة الدهان" },
  paintOneFace: { en: "Paint Area (1 face)", ar: "مساحة الدهان (وجه واحد)" },
  m2:           { en: "m²", ar: "م²" },
  siteTables:   { en: "Steel section tables — weight & paint area per metre", ar: "جداول المقاطع الحديدية — الوزن ومساحة الدهان للمتر" },
});
Object.assign(SL_TXT, {
  paint: { en: "Paint area (conservative)", ar: "مساحة الدهان (متحفّظة)" },
  page:  { en: "Full data page ↗", ar: "صفحة البيانات الكاملة ↗" },
});

/* ─── links to the static pages: shown only once they really exist ───────── */
let _soPagesCheck = null;
function useSectionPagesLive() {
  const [live, setLive] = useState(false);
  useEffect(() => {
    if (!SECTION_PAGE_LINKS || typeof fetch !== "function") return;
    if (!_soPagesCheck) _soPagesCheck = fetch("/steel-sections/s.css", { cache: "force-cache" })
      .then(r => r.ok && /css/i.test(r.headers.get("content-type") || ""))
      .catch(() => false);
    let on = true;
    _soPagesCheck.then(v => { if (on) setLive(!!v); });
    return () => { on = false; };
  }, []);
  return live;
}
function SiteLinks() {
  const { t, lang } = useLang();
  const live = useSectionPagesLive();
  if (!live) return null;
  return (
    <div style={{ maxWidth: 1180, margin: "0 auto", padding: "0 20px 44px", textAlign: "center" }}>
      <a href={lang === "ar" ? "/ar/steel-sections" : "/steel-sections"} style={{ color: "#fbbf24", fontFamily: "'Space Mono', monospace", fontSize: 15, textDecoration: "none", borderBottom: "1px dashed rgba(245,158,11,.5)" }}>📐 {t("siteTables")} →</a>
    </div>
  );
}

/* ─── the paint-area band shown in both results pages ────────────────────── */
function PaintBlock({ total, note, missing, count }) {
  const { t } = useLang();
  if (!PAINT_AREA_ENABLED || total == null) return null;
  const miss = (missing || []).filter(Boolean);
  const none = miss.length > 0 && miss.length === count;
  return (
    <div style={{ background: "rgba(56,189,248,.07)", border: "1px solid rgba(56,189,248,.35)", borderRadius: 12, padding: "16px 22px", marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
      <div style={{ flex: "1 1 320px" }}>
        <div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#7dd3fc", letterSpacing: ".12em" }}>{t("paintHd")}</div>
        <div style={{ fontSize: 14, color: "#94a3b8", marginTop: 6, lineHeight: 1.6 }}>{note}</div>
        {miss.length > 0 && <div style={{ fontSize: 14, color: "#fbbf24", marginTop: 6, lineHeight: 1.6 }}>{t("paintMissing", { list: miss.join(", ") })}</div>}
      </div>
      <div style={{ fontFamily: "'Space Mono', monospace", fontSize: 34, fontWeight: 800, color: "#e0f2fe", whiteSpace: "nowrap" }}>{none ? "—" : fmtM2(total)} <span style={{ fontSize: 17 }}>{t("m2")}</span></div>
    </div>
  );
}
/* ╚══ end of PAINT AREA + SECTION PAGES add-on ══════════════════════════════╝ */

/* ── Section Library: international groups (isolated add-on) ─────────────────
   Russia, China, Japan and India (SAP2000 databases) are listed under their
   own headers after the existing families; each region opens with one click.
   The existing families, their cards and their pages are not touched.      */
function soFamLabel(f, lang) {
  const e = SL_FAM_LABEL[f];
  if (e) return lang === "en" || lang === "ar" ? e[lang] : soUiTx(lang, "FAM." + f, e.en);
  const m = soSafe(() => intlData().meta[f]);
  return m && m.lab ? (m.lab[lang] || m.lab.en || f) : f;
}
const soIntlFamRows = (key) => soSafe(() => intlFamRows(key), []) || [];
const soIntlMatch = (r, q) => { const k = soSafe(() => intlKey(q), ""); return !k || (r.key || "").includes(k); };
function soIntlShort(r) {                                     // the family already names the standard
  const s = r.std ? " " + r.std : "";
  return s && r.name.endsWith(s) ? r.name.slice(0, -s.length) : r.name;
}
function SoIntlGroups({ lang, chip, word, open, setOpen, onPick }) {
  const fams = useMemo(() => soSafe(() => intlFamilies(), []) || [], []);
  if (!fams.length) return null;
  const regions = ["RU", "CN", "JP", "IN"].map(cc => ({ cc, list: fams.filter(m => m.cc === cc) })).filter(r => r.list.length);
  return (
    <div style={{ marginTop: 18, display: "grid", gap: 10 }}>
      {regions.map(({ cc, list }) => {
        const on = open === cc, n = list.reduce((a, m) => a + m.n, 0), lab = INTL_REGION[cc] || {};
        return (
          <div key={cc}>
            <button onClick={() => setOpen(on ? "" : cc)} aria-expanded={on}
              style={{ ...chip, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, borderColor: on ? "rgba(245,158,11,.45)" : chip.borderColor }}>
              <span style={{ color: "#fbbf24", fontFamily: "'Space Mono', monospace", fontWeight: 700, fontSize: 14.5 }}>{lab[lang] || lab.en || cc}</span>
              <span style={{ fontSize: 13, color: "#64748b", whiteSpace: "nowrap" }}>{n} {word} <span aria-hidden style={{ display: "inline-block", marginInlineStart: 6, color: "#fbbf24" }}>{on ? "▲" : "▼"}</span></span>
            </button>
            {on && (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(190px,1fr))", gap: 10, marginTop: 10 }}>
                {list.map(m => (
                  <button key={m.key} onClick={() => onPick(m.key)} style={chip}>
                    <div style={{ color: "#fbbf24", fontFamily: "'Space Mono', monospace", fontWeight: 700, fontSize: 15 }}>{m.code}{m.std ? <span style={{ color: "#94a3b8", fontWeight: 400, fontSize: 13 }}> · {m.std}</span> : null}</div>
                    <div style={{ fontSize: 13.5, color: "#94a3b8", marginTop: 2 }}>{(m.lab && (m.lab[lang] || m.lab.en)) || m.key}</div>
                    <div style={{ fontSize: 13, color: "#475569", marginTop: 4 }}>{m.n} {word}</div>
                  </button>))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
function SoIntlCard({ sec, lang, tx, U, famLabel }) {
  const d = sec.dims || {}, sh = sec.shape;
  const L = (key, ar, en) => (lang === "ar" ? ar : soUiTx(lang, key, en));
  const mm = v => `${v} ${U.mm}`;
  const dims = sh === "L" ? [[L("MISC.dimLegA", "الضلع a", "Leg a"), d.a], [L("MISC.dimLegB", "الضلع b", "Leg b"), d.b], [L("MISC.dimThick", "السماكة t", "Thickness t"), d.t]]
    : sh === "B" ? [[L("MISC.dimHeight", "الارتفاع h", "Height h"), d.h], [L("MISC.dimWidth", "العرض b", "Width b"), d.b], [L("MISC.dimWall", "سماكة الجدار t", "Wall thickness t"), d.t]]
    : sh === "P" ? [[L("MISC.dimOD", "القطر الخارجي D", "Outside diameter D"), d.D], [L("MISC.dimWall", "سماكة الجدار t", "Wall thickness t"), d.t]]
    : [[L("MISC.dimDepth", "الارتفاع h", "Depth h"), d.h], [L("MISC.dimFlangeW", "عرض الشفة b", "Flange width b"), d.b], [L("MISC.dimWebT", "سماكة الجذع tw", "Web thickness tw"), d.tw], [L("MISC.dimFlangeT", "سماكة الشفة tf", "Flange thickness tf"), d.tf]];
  const pa = sectionPaintSafe(sec);
  const A = (sec.A || 0) / 100;                                 // mm² → cm²
  const rows = [
    [tx("designation"), sec.name],
    ...dims.filter(([, v]) => v > 0).map(([k, v]) => [k, mm(v)]),
    ...(sec.std ? [[L("MISC.dimStd", "المواصفة", "Standard"), sec.std]] : []),
    [tx("family"), famLabel(sec.type)],
    [tx("unitMass"), `${sec.kgm} ${U.kgm}`],
    [tx("area"), `${A < 10 ? A.toFixed(2) : A.toFixed(1)} ${U.cm2}`],
    ...(PAINT_AREA_ENABLED ? [[tx("paint"), pa ? `${fmtM2pm(pa.m2pm)} ${U.m2pm}` : "—"]] : []),
    [`${tx("perBar")} — 6 ${U.m}`, `${(sec.kgm * 6).toFixed(1)} ${U.kg}`],
    [`${tx("perBar")} — 12 ${U.m}`, `${(sec.kgm * 12).toFixed(1)} ${U.kg}`],
    [`${tx("perTonne")} (12 ${U.m})`, `${(1000 / (sec.kgm * 12)).toFixed(1)} ${tx("barsOf")} 12 ${U.m}`],
  ];
  const note = (lang === "ar"
    ? "المصدر: قاعدة مقاطع SAP2000 ‏({file}) · {std}. الوزن = مساحة المقطع × 7850 كجم/م³. استُبعدت الصفوف التي لا تتطابق مساحتها مع أبعادها."
    : soUiTx(lang, "MISC.intlSrc", "Source: SAP2000 section database ({file}) · {std}. Mass = cross-section area × 7850 kg/m³. Rows whose area and dimensions disagree were left out."))
    .replace("{file}", sec.src || "SAP2000").replace("{std}", sec.std || ((INTL_REGION[sec.cc] || {})[lang] || ""));
  return (
    <div>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 15.5 }}>
        <tbody>{rows.map(([k, v], i) => (
          <tr key={i} style={{ borderBottom: "1px solid #1a2230" }}>
            <td style={{ padding: "10px 12px", color: "#94a3b8" }}>{k}</td>
            <td style={{ padding: "10px 12px", color: "#e2e8f0", fontFamily: "'Space Mono', monospace", textAlign: "end" }}>{v}</td>
          </tr>))}
        </tbody>
      </table>
      <div style={{ marginTop: 12, fontSize: 13.5, color: "#64748b", lineHeight: 1.7 }}>{note}</div>
    </div>
  );
}
/* ── end of Section Library international groups ── */

function SectionLibrary() {
  const { lang, dir } = useLang();
  const tx = (k) => (lang === "en" || lang === "ar" ? SL_TXT[k][lang] : soUiTx(lang, "SL." + k, SL_TXT[k].en));
  const U = lang === "ar" ? { mm: "مم", kgm: "كجم/م", cm2: "سم²", kg: "كجم", m: "م", m2pm: "م²/م" } : lang === "ru" ? { mm: "мм", kgm: "кг/м", cm2: "см²", kg: "кг", m: "м", m2pm: "м²/м" } : { mm: "mm", kgm: "kg/m", cm2: "cm²", kg: "kg", m: "m", m2pm: "m²/m" };
  const [open, setOpen] = useState(false);
  const [fam, setFam] = useState(null);
  const [sec, setSec] = useState(null);
  const pagesLive = useSectionPagesLive();
  const [q, setQ] = useState("");
  const [soRgn, setSoRgn] = useState(null);   // open region of the international groups: add-on
  const data = useMemo(() => {
    const seen = new Set(); const byFam = {};
    STEEL_DB.forEach(r => {
      if (seen.has(r.name)) return; seen.add(r.name);
      (byFam[r.type] = byFam[r.type] || []).push(r);
    });
    Object.values(byFam).forEach(a => a.sort((x, y) => x.kgm - y.kgm));
    return byFam;
  }, []);
  const fams = useMemo(() => Object.keys(data).sort((a, b) => (data[b].length - data[a].length)), [data]);
  const famLabel = (f) => soFamLabel(f, lang);   // + international families: add-on
  const listRaw = fam ? (data[fam] || soIntlFamRows(fam)) : [];
  const list = q ? listRaw.filter(r => (r.intl ? soIntlMatch(r, q) : r.name.toLowerCase().includes(q.toLowerCase()))) : listRaw;
  const closeAll = () => { setOpen(false); setFam(null); setSec(null); setQ(""); };
  const CARD = { background: "#0d1420", border: "1px solid #1e293b", borderRadius: 12 };
  const CHIP = { ...CARD, padding: "12px 14px", cursor: "pointer", textAlign: "start", color: "#cbd5e1", fontSize: 15, lineHeight: 1.5, width: "100%" };
  const areaCm2 = sec ? (sec.kgm / 7850 * 1e4) : 0;
  return (
    <>
      <button onClick={() => setOpen(true)} title={tx("title")}
        style={{ position: "fixed", top: 16, left: 16, zIndex: 10000,
          display: "inline-flex", alignItems: "center", gap: 8, padding: "8px 16px", borderRadius: 30, cursor: "pointer",
          background: "rgba(19,25,32,.82)", backdropFilter: "blur(10px)", WebkitBackdropFilter: "blur(10px)",
          border: "1px solid rgba(245,158,11,.45)", color: "#fbbf24",
          fontFamily: "'Space Mono', monospace", fontSize: 12, fontWeight: 700, letterSpacing: 1,
          boxShadow: "0 8px 24px -8px rgba(0,0,0,.6)" }}>
        <span aria-hidden>📐</span>{tx("btn")}
      </button>
      {open && (
        <div onClick={closeAll} style={{ position: "fixed", inset: 0, zIndex: 10001, background: "rgba(4,7,12,.78)", backdropFilter: "blur(6px)", display: "flex", alignItems: "flex-start", justifyContent: "center", overflowY: "auto", padding: "6vh 14px 40px" }}>
          <div dir={dir} onClick={e => e.stopPropagation()} style={{ ...CARD, width: "100%", maxWidth: 680, padding: 22, boxShadow: "0 30px 80px -20px rgba(0,0,0,.8)" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <div style={{ fontFamily: "'Space Mono', monospace", fontWeight: 700, color: "#fbbf24", fontSize: 17 }}>
                {sec ? sec.name : fam ? famLabel(fam) : tx("title")}
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                {(fam || sec) && <button onClick={() => sec ? setSec(null) : (setFam(null), setQ(""))} style={{ ...CHIP, width: "auto", padding: "6px 14px", color: "#fbbf24", borderColor: "rgba(245,158,11,.45)" }}>← {tx("back")}</button>}
                <button onClick={closeAll} style={{ ...CHIP, width: "auto", padding: "6px 14px" }}>{tx("close")}</button>
              </div>
            </div>
            {!fam && !sec && (<>
              <div style={{ color: "#64748b", fontSize: 14, marginBottom: 12 }}>{tx("pickFam")}</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(190px,1fr))", gap: 10 }}>
                {fams.map(f => (
                  <button key={f} onClick={() => setFam(f)} style={CHIP}>
                    <div style={{ color: "#fbbf24", fontFamily: "'Space Mono', monospace", fontWeight: 700, fontSize: 15 }}>{f}</div>
                    <div style={{ fontSize: 13.5, color: "#94a3b8", marginTop: 2 }}>{soCap(famLabel(f).replace(/^[^—]+—\s*/, ""), lang)}</div>
                    <div style={{ fontSize: 13, color: "#475569", marginTop: 4 }}>{data[f].length} {tx("sections")}</div>
                  </button>))}
              </div>
              <SoIntlGroups lang={lang} chip={CHIP} word={tx("sections")} open={soRgn === null ? (({ ru: "RU", zh: "CN" })[lang] || "") : soRgn} setOpen={setSoRgn} onPick={setFam} />
            </>)}
            {fam && !sec && (<>
              <input value={q} onChange={e => setQ(e.target.value)} placeholder={tx("search")}
                style={{ width: "100%", boxSizing: "border-box", marginBottom: 12, padding: "10px 14px", borderRadius: 8, background: "#0a0f18", border: "1px solid #1e293b", color: "#e2e8f0", fontSize: 16, outline: "none" }} />
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(150px,1fr))", gap: 8, maxHeight: "52vh", overflowY: "auto", paddingInlineEnd: 4 }}>
                {list.map(r => (
                  <button key={r.name} onClick={() => setSec(r)} style={{ ...CHIP, padding: "10px 12px", display: "flex", justifyContent: "space-between", gap: 8 }}>
                    <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 14.5, color: "#e2e8f0" }}>{r.intl ? soIntlShort(r) : r.name}</span>
                    <span style={{ fontSize: 13.5, color: "#f59e0b", whiteSpace: "nowrap" }}>{r.kgm} {U.kgm}</span>
                  </button>))}
              </div>
            </>)}
            {sec && (() => {
              if (sec.intl) return <SoIntlCard sec={sec} lang={lang} tx={tx} U={U} famLabel={famLabel} />;   // international sections: add-on
              const dm = soSafe(() => sectionDims(sec.name)) || SECTION_DIMS[sec.name];
              const pa = sectionPaintSafe(sec);
              const pagePath = pagesLive ? soSafe(() => sectionPagePath(sec, lang)) : null;
              const asianH = ["JIS-HW","JIS-HM","JIS-HN"].indexOf(sec.type) >= 0;
              const rows = [
                [tx("designation"), sec.name],
                ...(dm ? [
                  [lang === "ar" ? "الارتفاع h" : soUiTx(lang, "MISC.dimDepth", "Depth h"), `${dm[0]} ${U.mm}`],
                  [lang === "ar" ? "عرض الشفة b" : soUiTx(lang, "MISC.dimFlangeW", "Flange width b"), `${dm[1]} ${U.mm}`],
                  [lang === "ar" ? "سماكة الجذع tw" : soUiTx(lang, "MISC.dimWebT", "Web thickness tw"), `${dm[2]} ${U.mm}`],
                  [lang === "ar" ? "سماكة الشفة tf" : soUiTx(lang, "MISC.dimFlangeT", "Flange thickness tf"), `${dm[3]} ${U.mm}`],
                ] : []),
                [tx("family"), famLabel(sec.type)],
                [tx("unitMass"), `${sec.kgm} ${U.kgm}`],
                [tx("area"), `${areaCm2.toFixed(1)} ${U.cm2}`],
                ...(PAINT_AREA_ENABLED ? [[tx("paint"), pa ? `${fmtM2pm(pa.m2pm)} ${U.m2pm}` : "—"]] : []),
                [`${tx("perBar")} — 6 ${U.m}`, `${(sec.kgm * 6).toFixed(1)} ${U.kg}`],
                [`${tx("perBar")} — 12 ${U.m}`, `${(sec.kgm * 12).toFixed(1)} ${U.kg}`],
                [`${tx("perTonne")} (12 ${U.m})`, `${(1000 / (sec.kgm * 12)).toFixed(1)} ${tx("barsOf")} 12 ${U.m}`],
              ];
              return (
                <div>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 15.5 }}>
                    <tbody>{rows.map(([k, v], i) => (
                      <tr key={i} style={{ borderBottom: "1px solid #1a2230" }}>
                        <td style={{ padding: "10px 12px", color: "#94a3b8" }}>{k}</td>
                        <td style={{ padding: "10px 12px", color: "#e2e8f0", fontFamily: "'Space Mono', monospace", textAlign: "end" }}>{v}</td>
                      </tr>))}
                    </tbody>
                  </table>
                  {pagePath && <a href={pagePath} target="_blank" rel="noopener" style={{ display: "inline-block", marginTop: 12, color: "#fbbf24", fontSize: 14.5, fontFamily: "'Space Mono', monospace", textDecoration: "none", borderBottom: "1px dashed rgba(245,158,11,.5)" }}>{tx("page")}</a>}
                  <div style={{ marginTop: 12, fontSize: 13.5, color: "#64748b", lineHeight: 1.7 }}>{tx("note")}{dm ? <><br />{lang === "en" || lang === "ar" ? SL_DIM_SRC[lang] : soUiTx(lang, "DIMSRC", SL_DIM_SRC.en)}</> : null}{asianH ? <><br />{lang === "ar" ? "\u0645\u0644\u0627\u062d\u0638\u0629: h\u060c b\u060c tw\u060c tf \u0645\u062a\u0637\u0627\u0628\u0642\u0629 \u0641\u064a JIS G 3192 \u0648 KS D 3502 \u0648 GB/T 11263 \u2014 \u064a\u062e\u062a\u0644\u0641 \u0641\u0642\u0637 \u0646\u0635\u0641 \u0642\u0637\u0631 \u0627\u0644\u062a\u062f\u0648\u064a\u0631\u060c \u0641\u064a\u0646\u062a\u062c \u0641\u0631\u0642 \u0648\u0632\u0646 \u062d\u062a\u0649 2%. \u0627\u0644\u0623\u0648\u0632\u0627\u0646 \u0647\u0646\u0627 \u062a\u062a\u0628\u0639 KS/GB." : soUiTx(lang, "MISC.asianNote", "Note: h, b, tw, tf are identical across JIS G 3192, KS D 3502 and GB/T 11263 \u2014 only the root radius differs, giving up to 2% mass difference. Masses here follow KS/GB.")}</> : null}</div>
                </div>
              );
            })()}
          </div>
        </div>
      )}
    </>
  );
}


/* ╔══════════════════════════════════════════════════════════════════════════╗
   ║  CONTACT — fill these three lines in, or the block hides itself.          ║
   ╚══════════════════════════════════════════════════════════════════════════╝
   WhatsApp matters more than email in the Gulf: an engineer will photograph a
   cutting list and send it in seconds, but will not open an email client to
   write to a stranger. Put the number in international format, digits only.

   Use a dedicated number or WhatsApp Business — not the personal number tied
   to your employer. It keeps the product separate from your day job and keeps
   your private line out of reach of scrapers.
──────────────────────────────────────────────────────────────────────────── */
const CONTACT = {
  email: "support@steeloptimizer.com",
  whatsapp: "",           // digits only, country code first, e.g. "966XXXXXXXXX"
};

/* Thin bar at the very top: the moment an import looks wrong, the way out is
   already on screen. A failed file mailed to us is worth more than a silent
   bounce — it is the only way to find out what the parser cannot read yet. */
function ContactTopBar() {
  const t = useT();
  if (!CONTACT.email) return null;
  return (
    <div style={{
      position: "relative", zIndex: 5, textAlign: "center",
      padding: "8px 16px", fontSize: 14.5, lineHeight: 1.6,
      background: "rgba(245,158,11,.08)", borderBottom: "1px solid rgba(245,158,11,.18)",
      color: "#94a3b8", fontFamily: "'Space Mono', monospace",
    }}>
      {t("ctTop")}{" "}
      <a href={`mailto:${CONTACT.email}?subject=Steel%20Optimizer%20-%20file%20did%20not%20import`}
         onClick={() => track("contact_click", { via: "topbar" })}
         style={{ color: "#fbbf24", fontWeight: 700, textDecoration: "none", borderBottom: "1px solid rgba(251,191,36,.4)" }}>
        {CONTACT.email}
      </a>
    </div>
  );
}

function ContactFooter() {
  const t = useT();
  const has = CONTACT.email || CONTACT.whatsapp;
  if (!has) return null;                       // nothing configured — stay hidden
  const waMsg = encodeURIComponent(RT("Hi, I'm using Steel Optimizer and I need help with my material list.", "مرحباً، أستخدم Steel Optimizer وأحتاج مساعدة في قائمة المواد."));
  const btn = {
    display: "inline-flex", alignItems: "center", gap: 8, padding: "11px 20px",
    borderRadius: 8, textDecoration: "none", fontSize: 15.5, fontWeight: 600,
    fontFamily: "'Space Mono', monospace", whiteSpace: "nowrap",
  };
  return (
    <div style={{ maxWidth: 1180, margin: "0 auto", padding: "0 20px 72px" }}>
      <div style={{
        padding: "26px 26px 24px", borderRadius: 14,
        background: "linear-gradient(150deg, rgba(245,158,11,.07), rgba(10,14,20,.4))",
        border: "1px solid rgba(245,158,11,.20)",
      }}>
        <div style={{ fontSize: 20, fontWeight: 800, color: "#f8fafc", fontFamily: "'Playfair Display', serif", marginBottom: 9 }}>
          {t("ctTitle")}
        </div>
        <div style={{ fontSize: 15.5, color: "#94a3b8", lineHeight: 1.7, marginBottom: 18, maxWidth: 720 }}>
          {t("ctBody")}
        </div>

        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          {CONTACT.whatsapp && (
            <a href={`https://wa.me/${CONTACT.whatsapp}?text=${waMsg}`} target="_blank" rel="noreferrer"
               onClick={() => track("contact_click", { via: "whatsapp" })}
               style={{ ...btn, background: "#25D366", color: "#062b12" }}>
              ✆ {t("ctWhats")}
            </a>
          )}
          {CONTACT.email && (
            <a href={`mailto:${CONTACT.email}?subject=Steel%20Optimizer%20-%20help%20with%20my%20list`}
               onClick={() => track("contact_click", { via: "email" })}
               style={{ ...btn, background: "rgba(245,158,11,.15)", border: "1px solid #d97706", color: "#fbbf24" }}>
              ✉ {t("ctEmail")}
            </a>
          )}
          {CONTACT.whatsapp && (
            <span style={{ fontSize: 13.5, color: "#64748b", lineHeight: 1.6 }}>{t("ctWhatsHint")}</span>
          )}
        </div>

        {CONTACT.email && (
          <div style={{ marginTop: 14, fontSize: 14, color: "#64748b", fontFamily: "'Space Mono', monospace", wordBreak: "break-all" }}>
            {CONTACT.email}
          </div>
        )}

        <div style={{ marginTop: 18, paddingTop: 16, borderTop: "1px solid rgba(148,163,184,.12)", fontSize: 13.5, color: "#64748b", lineHeight: 1.7 }}>
          {t("ctBlocked")}
        </div>
        <div style={{ marginTop: 12, fontSize: 13.5, color: "#475569" }}>{t("ctBy")}</div>
      </div>
    </div>
  );
}

function AppInner() {
  // Re-check the subscription once per load: a cancelled plan re-locks itself.
  useEffect(() => {
    loadGA();                                        // page_view fires automatically
  }, []);
  const { dir } = useLang();
  const wsRef = useRef(null);
  const preSection = useMemo(() => { try { const v = new URLSearchParams(window.location.search).get("section"); const hit = v ? findSection(v) : null; return hit ? hit.name : null; } catch { return null; } }, []);
  const [module, setModule] = useState(preSection ? "sections" : null);
  const scrollWs = () => wsRef.current?.scrollIntoView({ behavior: "smooth" });
  useEffect(() => { if (!preSection) return; track("from_section_page", { section: preSection }); const tm = setTimeout(scrollWs, 500); return () => clearTimeout(tm); }, [preSection]);
  return (
    <div dir={dir} style={{ position: "relative", background: "#070a0f", minHeight: "100vh", color: "#cbd5e1", fontFamily: "'DM Sans', system-ui, sans-serif", overflow: "hidden" }}>
      <ContactTopBar />
      {I18N_EXTRA_ENABLED ? <SoLangMenu /> : <LangToggle />}
      <SectionLibrary />
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=Space+Mono:wght@400;700&family=Playfair+Display:wght@700;900&display=swap" rel="stylesheet" />
      <style>{`
        @keyframes floatGlow { 0%,100%{ transform:translate(0,0) scale(1); opacity:.55 } 50%{ transform:translate(2%,-3%) scale(1.08); opacity:.8 } }
        @keyframes floatGlow2 { 0%,100%{ transform:translate(0,0) scale(1); opacity:.4 } 50%{ transform:translate(-3%,2%) scale(1.12); opacity:.65 } }
        @keyframes gridDrift { from{ background-position:0 0 } to{ background-position:0 64px } }
        @keyframes ctaPulse { 0%,100%{ box-shadow:0 10px 34px rgba(245,158,11,.35) } 50%{ box-shadow:0 14px 52px rgba(245,158,11,.6) } }
        @keyframes hintBounce { 0%,100%{ transform:translate(-50%,0) } 50%{ transform:translate(-50%,7px) } }
        ::-webkit-scrollbar{ width:10px; height:10px } ::-webkit-scrollbar-track{ background:#0b0f15 } ::-webkit-scrollbar-thumb{ background:#1e293b; border-radius:5px } ::-webkit-scrollbar-thumb:hover{ background:#f59e0b66 }
        @media (prefers-reduced-motion: reduce){ *{ animation:none !important } }
      `}</style>
      {/* ── atmospheric backdrop ── */}
      <div style={{ position: "fixed", inset: 0, zIndex: 0, pointerEvents: "none" }}>
        <div style={{ position: "absolute", inset: 0, background: "radial-gradient(120% 80% at 50% -10%, #11202e 0%, #0a0e14 45%, #070a0f 100%)" }} />
        <div style={{ position: "absolute", top: "-12%", left: "-8%", width: "55vw", height: "55vw", background: "radial-gradient(circle, rgba(245,158,11,.22), transparent 62%)", filter: "blur(36px)", animation: "floatGlow 16s ease-in-out infinite" }} />
        <div style={{ position: "absolute", bottom: "-18%", right: "-10%", width: "50vw", height: "50vw", background: "radial-gradient(circle, rgba(20,160,180,.18), transparent 62%)", filter: "blur(40px)", animation: "floatGlow2 20s ease-in-out infinite" }} />
        <div style={{ position: "absolute", inset: 0, backgroundImage: "linear-gradient(rgba(148,163,184,.05) 1px, transparent 1px), linear-gradient(90deg, rgba(148,163,184,.05) 1px, transparent 1px)", backgroundSize: "64px 64px", maskImage: "radial-gradient(120% 90% at 50% 0%, #000 35%, transparent 85%)", WebkitMaskImage: "radial-gradient(120% 90% at 50% 0%, #000 35%, transparent 85%)", animation: "gridDrift 24s linear infinite" }} />
        <div style={{ position: "absolute", inset: 0, boxShadow: "inset 0 0 240px 40px rgba(0,0,0,.7)" }} />
      </div>
      <div style={{ position: "relative", zIndex: 1 }}>
        <Cover3D onStart={() => soRequireMember(scrollWs, false)} />
        <LiveCutShowcase />
        <div ref={wsRef} style={{ maxWidth: 1180, margin: "0 auto", padding: "32px 20px 64px" }}>
          {module === null && <ModuleChooser onPick={setModule} />}
          {module === "plates" && <PlatesModule onBack={() => setModule(null)} />}
          {module === "sections" && <SectionsModule onBack={() => setModule(null)} preSection={preSection} />}
        </div>
        <ContactFooter />
        <SiteLinks />
      </div>
      <EmailGate />
    </div>
  );
}


/* ─── PLATES MODULE ──────────────────────────────────────────────────────── */



/* ─── SECTIONS MODULE ────────────────────────────────────────────────────── */
function SectionsModule({ onBack, preSection }) {
  const { t, lang } = useLang();
  const pls = n => ((lang === "en" || lang === "es") && n > 1) ? "s" : "";
  const pricing = usePricing();
  const [inputMode, setInputMode] = useState(preSection ? "manual" : null);
  const [DEFAULT_STOCK, setDefStock] = useState(12000);   // one stock length for every profile (settings card)
  const [rows, setRows] = useState([{ id: genId(), profile: preSection || "", grade: "", lengths: "", stock: "" }]);
  const [activeAuto, setActiveAuto] = useState(null);
  const [gradeAuto, setGradeAuto] = useState(null);
  const [cutList, setCutList] = useState([]); const [kerf, setKerf] = useState(3); const [results, setResults] = useState(null); const [error, setError] = useState(""); const [uploadInfo, setUploadInfo] = useState(null);
  const [leftovers, setLeftovers] = useState([]); // owned bar offcuts {id, profile, grade, length, qty}
  const [loAuto, setLoAuto] = useState(null);
  const optBtnRef = useRef(null);
  const scrollToOpt = () => setTimeout(() => optBtnRef.current?.scrollIntoView({ behavior: "smooth", block: "center" }), 120);
  const fileRef = useRef(null);

  const upRow = (id, f, v) => setRows(p => p.map(r => r.id === id ? { ...r, [f]: v } : r));
  const addRow = () => setRows(p => [...p, { id: genId(), profile: "", grade: "", lengths: "", stock: "" }]);
  const delRow = id => setRows(p => p.filter(r => r.id !== id));
  const addLeftover = () => setLeftovers(p => [...p, { id: genId(), profile: "", grade: "", length: 6000, qty: 1 }]);
  const upLeftover = (id, f, v) => setLeftovers(p => p.map(r => r.id === id ? { ...r, [f]: (f === "profile" || f === "grade") ? v : (+v || 0) } : r));
  const delLeftover = id => setLeftovers(p => p.filter(r => r.id !== id));
  const importLeftovers = items => setLeftovers(p => [...p, ...items.map(it => ({ id: genId(), profile: (it.profile || "").toUpperCase(), grade: (it.grade || "").toUpperCase(), length: Math.round(it.length || 0), qty: Math.max(1, Math.round(it.qty || 1)) }))]);
  const parseLengths = raw => { if (!raw) return []; const toks = String(raw).trim().split(/[\s,;]+/).filter(Boolean); const out = []; for (const t of toks) { const m = t.match(/^(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+)$/i); if (m) { const len = +m[1], q = +m[2]; if (len > 0 && q > 0) for (let i = 0; i < q; i++) out.push(len); } else { const len = parseFloat(t); if (len > 0) out.push(len); } if (out.length >= 200) break; } return out.slice(0, 200); };
  const buildManual = () => { const list = []; rows.forEach(r => { const lens = parseLengths(r.lengths); const stock = parseInt(r.stock) > 0 ? parseInt(r.stock) : DEFAULT_STOCK; const counts = {}; lens.forEach(l => counts[l] = (counts[l] || 0) + 1); Object.entries(counts).forEach(([len, qty]) => list.push({ profile: (r.profile || "(unspecified)").toUpperCase(), grade: r.grade || "", length: +len, qty, stock })); }); return list; };

  const SECTION_RX = /\b(?:IPE|IPN|HEA|HEB|HEM|HE|UBP|UB|UC|HW|HM|HN|RHS|SHS|CHS|UPN|UPE|PFC|PIPE)\s?\d+(?:\s?[x×*]\s?\d+(?:\.\d+)?)*\b|\b\d+x\d+x\d+\s?(?:UB|UC|UBP|PFC)\b|\b(?:HW|HM|HN|C|L|I|H|LC)\s?\d+[x×*]\d+(?:[x×*]\d+(?:\.\d+)?){0,2}\b|\bW\d+x\d+\b|\b(?:GOST|GB)\s?[ICKB]?\s?\d+(?:\.\d+)?[a-cAB12]?\b/i;
  const looksProfile = c => { if (!c) return false; if (isExcludedProfile(c)) return false; if (findSection(c)) return true; return SECTION_RX.test(c); };
  const num = v => { if (v == null || v === "") return null; let s = String(v).trim().replace(/\s/g, ""); if (s.includes(",") && s.includes(".")) { if (s.lastIndexOf(",") > s.lastIndexOf(".")) s = s.replace(/\./g, "").replace(",", "."); else s = s.replace(/,/g, ""); } else if (s.includes(",")) s = /,\d{1,2}$/.test(s) ? s.replace(",", ".") : s.replace(/,/g, ""); const n = parseFloat(s); return isFinite(n) ? n : null; };
  function autoRead(wb) {
    const out = []; let skipped = 0; const unknownRows = new Map();
    wb.SheetNames.forEach(name => {
      const data = intlPrepSheet(arPrepSheet(XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, defval: "" })));   // Arabic sheets: isolated add-on
      let hr = -1, cm = null, lenUnit = null;
      for (let i = 0; i < Math.min(data.length, 30); i++) {
        const cells = (data[i] || []).map(c => normHeader(c).toLowerCase());
        const f = re => cells.findIndex(c => re.test(c));
        const strict = f(/profile|section|designation|shape/);
        // Advance Steel and several fabrication tools put the profile under "Description".
        const pi = strict !== -1 ? strict : f(/description|\bsize\b|\bdesc\b|\bmember\b/);
        const li = f(/length|^len\b|^len[ (]|^l ?\(mm\)|^l$/);
        if (pi !== -1 && li !== -1) {
          hr = i;
          const below = (data[i + 1] || []).map(c => normHeader(c).toLowerCase());
          // Units frequently sit on a second header row, e.g. Length / (inch).
          lenUnit = detectLenUnit(cells[li]) || detectLenUnit(below[li]) || null;
          cm = { profile: pi, grade: f(/grade|material/), qty: f(/\bqty\b|quantity|\bnos?\b|count|number|pcs/), length: li, ...soWeightCols(cells) };   // weight columns: Tekla add-on
          break;
        }
      }
      if (cm) {
        for (let i = hr + 1; i < data.length; i++) { const r = data[i] || []; if (r.some(c => /^\s*(grand\s+)?total\b/i.test(String(c == null ? "" : c)))) continue; let rawP = String(r[cm.profile] == null ? "" : r[cm.profile]).trim(); if (rawP && !looksProfile(rawP) && !isExcludedProfile(rawP)) { for (let ci = 0; ci < r.length; ci++) { if (ci === cm.length || ci === cm.qty) continue; const cand = String(r[ci] == null ? "" : r[ci]).trim(); if (cand && findSection(cand)) { rawP = cand; break; } } } if (!rawP) continue; if (isExcludedProfile(rawP)) { skipped++; continue; } if (!looksProfile(rawP)) { const q0 = cm.qty !== -1 ? num(r[cm.qty]) : 1; soUnknownAdd(unknownRows, rawP, q0, cm.grade !== -1 ? r[cm.grade] : "", soRowKg(r, cm, q0, num)); continue; } const length = cm.length !== -1 ? parseLengthToMm(r[cm.length], lenUnit) : null; if (!length || length < 5 || length > 40000) { const q0 = cm.qty !== -1 ? num(r[cm.qty]) : 1; if (cm.length !== -1 && String(r[cm.length] == null ? "" : r[cm.length]).trim() !== "") soUnknownAdd(unknownRows, `${rawP.toUpperCase()} (L = ${String(r[cm.length]).trim()})`, q0, cm.grade !== -1 ? r[cm.grade] : "", soRowKg(r, cm, q0, num)); continue; } const qty = (cm.qty !== -1 && num(r[cm.qty])) ? Math.round(num(r[cm.qty])) : 1; const grade = cm.grade !== -1 ? String(r[cm.grade] || "").trim() : ""; const fkgm = soRowKgm(r, cm, length, qty, num); out.push({ profile: rawP.toUpperCase(), grade: grade.toUpperCase(), length, qty, ...(fkgm ? { fkgm } : {}) }); }
      } else {
        for (const r of data) { if (!r || !r.length) continue; const cells = r.map(c => c == null ? "" : String(c).trim()); if (cells.some(c => /^total/i.test(c))) continue; let profile = "", pIdx = -1; for (let ci = 0; ci < cells.length; ci++) { const c = cells[ci]; if (!profile) { if (isExcludedProfile(c)) { profile = "__PLATE__"; break; } if (findSection(c)) { profile = c.replace(/\s+/g, " ").toUpperCase(); pIdx = ci; } else { const sm = c.match(SECTION_RX); if (sm) { profile = sm[0].replace(/\s+/g, " ").toUpperCase(); pIdx = ci; } } } } if (profile === "__PLATE__") { skipped++; continue; } const lengths = [], smalls = []; cells.forEach((c, ci) => { if (ci === pIdx) return; if (/^(?:A|S|SS|GR|GRADE|ASTM|EN|Q)\s?\d{2,4}[A-Z]*$/i.test(c)) return; if (/['"]/.test(c) || /\d\s*\/\s*\d/.test(c)) { const im = parseLengthToMm(c, null); if (im != null && im >= 200 && im <= 40000) lengths.push(im); return; } const m = c.replace(/,/g, "").match(/-?\d+(\.\d+)?/g); if (m) m.forEach(x => { const v = parseFloat(x); if (v >= 200 && v <= 40000) lengths.push(v); else if (Number.isInteger(v) && v >= 1 && v <= 999) smalls.push(v); }); }); if (lengths.length && profile) out.push({ profile, grade: "", length: lengths[lengths.length - 1], qty: smalls.length ? smalls[smalls.length - 1] : 1 }); }
      }
    });
    const merged = {}; out.forEach(r => { const k = `${r.profile}||${r.grade}||${r.length}`; if (!merged[k]) merged[k] = { ...r }; else merged[k].qty += r.qty; }); const res = Object.values(merged); res._skipped = skipped; res._unknown = soUnknownList(unknownRows); return res;
  }
  const handleFile = async f => { setError(""); setUploadInfo(null); track("upload_excel", { module: "sections", ext: (f.name || "").split(".").pop() }); alertMe("SOMEONE UPLOADED A FILE (sections)", { ext: (f.name || "").split(".").pop(), sizeKB: Math.round((f.size || 0) / 1024) }); try { const buf = await f.arrayBuffer(); const wb = XLSX.read(buf, { type: "array" }); const det = autoRead(wb); const skipped = det._skipped || 0; const list = det.map(r => ({ ...r, stock: DEFAULT_STOCK, stockAuto: true })); if (!list.length) { setError(t("secUpErr")); return; } setCutList(list); setUploadInfo({ name: f.name, rows: list.length, pieces: list.reduce((s, r) => s + r.qty, 0), skipped, unknown: det._unknown || [] }); scrollToOpt(); } catch (e) { setError(t("secReadErr", { msg: e.message })); } };
  const canonProfileKey = raw => { const sec = findSection(raw); if (sec) return sec.name.toUpperCase().replace(/\s+/g, ""); return String(raw || "").toUpperCase().replace(/×/g, "X").replace(/\*/g, "X").replace(/\s+/g, "").replace(/(\d)\.0+(?=\D|$)/g, "$1"); };
  const runOpt = list0 => {
    const list = list0.map(c => (c.stockAuto ? { ...c, stock: DEFAULT_STOCK } : c));   // Excel / scan rows: stock length from the settings card
    setError("");
    const valid = list.filter(c => c.length > 0 && c.qty > 0);
    if (!valid.length) { setError(t("secNoValid")); return; }
    const groups = {};
    valid.forEach(c => { const k = `${canonProfileKey(c.profile)}||${(c.grade || "").toUpperCase().trim()}`; if (!groups[k]) groups[k] = { items: [], displayProfile: c.profile, grade: (c.grade || "").trim() || "Steel (grade not specified)" }; groups[k].items.push(c); });
    // Owned leftover bars, keyed identically so each profile+grade reuses only its own offcuts.
    const loGroups = {};
    leftovers.forEach(lo => { const len = Math.round(lo.length || 0), qty = Math.max(0, Math.round(lo.qty || 0)); if (len <= 0 || qty <= 0) return; const k = `${canonProfileKey(lo.profile)}||${(lo.grade || "").toUpperCase().trim()}`; (loGroups[k] = loGroups[k] || []).push({ length: len, qty }); });
    let grandKg = 0, reusedGrandKg = 0;
    const gres = Object.entries(groups).map(([k, { items, displayProfile, grade }]) => {
      const sec = findSection(displayProfile); const profile = sec ? sec.name : displayProfile; const fk = soFileKgm(items), km = soKgm(sec, fk), kgm = km.kgm;   // file weight check: Tekla add-on
      const o = nestBarsLO(items, kerf, loGroups[k] || []);
      const netKg = kgm ? ((o.summary.totalNet + o.summary.reusedNet) / 1000) * kgm : null; // steel cut: new + reused
      const stockKg = kgm ? (o.summary.totalStock / 1000) * kgm : null;                    // weight to BUY (new bars only)
      const reusedKg = kgm ? (o.summary.reusedNet / 1000) * kgm : null;                     // value pulled from leftovers
      const reusedRemainKg = kgm ? (o.summary.reusedRemain / 1000) * kgm : null;
      if (stockKg) grandKg += stockKg; if (reusedKg) reusedGrandKg += reusedKg;
      return { profile, grade, kgm, netKg, stockKg, reusedKg, reusedRemainKg, computed: sec?.computed && !km.fromFile, kgmFromFile: km.fromFile || undefined, fileKgm: fk || undefined, fileDiff: km.diff != null ? km.diff : undefined, srcName: sec ? soSrcName(displayProfile, profile) : null, ...o, ...sectionListPaint(sec, items) };   // paint area: isolated add-on
    });
    if (inputMode === "manual") soNotifyManual("sections", { profiles: gres.length, rows: list.length, pieces: valid.reduce((s, c) => s + (+c.qty || 0), 0), via: preSection ? "section page" : "app" });
    else alertMe("SOMEONE RAN AN OPTIMISATION (sections)", { profiles: gres.length, rows: list.length });
    track("optimization_completed", { module: "sections", profiles: gres.length, tonnes: grandKg ? Math.round(grandKg / 1000 * 100) / 100 : 0 });
    setResults({ groups: gres, grandKg, reusedGrandKg, ...sectionPaintTotals(gres) });
  };

  if (results) return <SectionResults results={results} kerf={kerf} pricing={pricing} onBack={() => setResults(null)} />;

  return (
    <>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}><SectionTitle>{t("sectionsWS")}</SectionTitle><BackChip onClick={onBack} labelKey="backModules" /></div>
      {inputMode === null && (
        <div>
          <div style={{ textAlign: "center", color: "#94a3b8", fontSize: 17, marginBottom: 18 }}>{t("howEnter")}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(260px,1fr))", gap: 20 }}>
            <Chooser icon="✏️" titleKey="manualEntry" descKey="manualDescS" onClick={() => setInputMode("manual")} />
            <Chooser icon="📊" titleKey="uploadTitle" descKey="uploadDescS" onClick={() => setInputMode("excel")} />
            <Chooser icon="🔍" titleKey="scanTitle" descKey="scanDescS" onClick={() => setInputMode("scan")} />
          </div>
        </div>
      )}
      {inputMode !== null && (
        <Card title={t("stockSettings")}>
          <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end" }}>
            <div><Label>{t("secStockLbl")}</Label><input type="number" value={DEFAULT_STOCK} onChange={e => setDefStock(+e.target.value || 0)} style={{ ...IN, width: 130 }} /></div>
            <div><Label>{t("lblKerf")}</Label><input type="number" value={kerf} onChange={e => setKerf(+e.target.value)} style={{ ...IN, width: 90 }} /></div>
          </div>
          <div style={{ marginTop: 8, fontSize: 13, color: "#94a3b8", fontFamily: "'Space Mono', monospace" }}>{t("secStockHint")}</div>
        </Card>
      )}
      {inputMode === "manual" && (
        <Card title={t("manualCutList")} style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} labelKey="changeMethod" />}>
          <div style={{ overflowX: "visible" }}><table style={{ width: "100%", borderCollapse: "separate", borderSpacing: "0 10px" }}><thead><tr>{[t("lblProfile"), t("lblGradeShort"), t("thLengths"), t("thMarketLen"), ""].map((h, i) => <th key={h} style={{ textAlign: "left", padding: "0 8px", color: "#64748b", fontSize: 13, letterSpacing: 1, width: i === 0 ? "24%" : i === 1 ? "14%" : i === 3 ? "18%" : i === 4 ? 40 : "auto" }}>{h}</th>)}</tr></thead>
            <tbody>{rows.map(r => { const count = parseLengths(r.lengths).length; const matches = activeAuto === r.id ? searchSections(r.profile) : []; const sec = findSection(r.profile); return (
              <tr key={r.id}>
                <td style={{ position: "relative", padding: "0 8px" }}><input style={INL} placeholder="IPE 300" value={r.profile} onChange={e => { upRow(r.id, "profile", e.target.value); setActiveAuto(r.id); }} onFocus={() => setActiveAuto(r.id)} onBlur={() => setTimeout(() => setActiveAuto(a => a === r.id ? null : a), 150)} />{sec && <div style={{ fontSize: 13, color: sec.computed ? "#60a5fa" : "#22c55e", fontFamily: "'Space Mono', monospace", marginTop: 5 }}>{sec.kgm} {t("thKgM")} · {sec.intl ? soIntlBadge(sec) : sec.type}{sec.computed ? " " + t("computedSuffix") : ""}</div>}{matches.length > 0 && <div style={{ position: "absolute", top: "100%", left: 8, right: 8, zIndex: 50, background: "#0f1318", border: "1px solid #2d3748", borderRadius: 6, maxHeight: 260, overflowY: "auto", boxShadow: "0 8px 24px rgba(0,0,0,.5)" }}>{matches.map(m => <div key={m.name} onMouseDown={() => { upRow(r.id, "profile", m.name); setActiveAuto(null); }} style={{ padding: "11px 14px", cursor: "pointer", display: "flex", justifyContent: "space-between", borderBottom: "1px solid #1a2230", fontFamily: "'Space Mono', monospace", fontSize: 15 }}><span style={{ color: "#cbd5e1" }}>{m.name}</span><span style={{ color: "#f59e0b" }}>{m.kgm} {t("thKgM")}</span></div>)}</div>}</td>
                <td style={{ position: "relative", padding: "0 8px" }}><input style={INL} placeholder="S355JR" value={r.grade} onChange={e => { upRow(r.id, "grade", e.target.value); setGradeAuto(r.id); }} onFocus={() => setGradeAuto(r.id)} onBlur={() => setTimeout(() => setGradeAuto(a => a === r.id ? null : a), 160)} />{gradeAuto === r.id && (() => { const groups = gradeMatches(r.grade); return groups.length > 0 && <div style={{ position: "absolute", top: "100%", left: 8, right: 8, zIndex: 60, background: "#0f1318", border: "1px solid #2d3748", borderRadius: 8, maxHeight: 320, overflowY: "auto", boxShadow: "0 10px 30px rgba(0,0,0,.55)" }}>{groups.map(([grp, list]) => <div key={grp}><div style={{ padding: "7px 12px", fontSize: 12, letterSpacing: 1, color: "#f59e0b", background: "rgba(245,158,11,.07)", fontFamily: "'Space Mono', monospace", position: "sticky", top: 0 }}>{lang === "ar" ? (GRADE_GROUP_AR[grp] || grp) : soGradeGroup(grp, lang)}</div>{list.map(g => <div key={g} onMouseDown={() => { upRow(r.id, "grade", g); setGradeAuto(null); }} style={{ padding: "9px 14px", cursor: "pointer", borderBottom: "1px solid #161d27", fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#cbd5e1" }} onMouseEnter={e => e.currentTarget.style.background = "rgba(245,158,11,.1)"} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>{g}</div>)}</div>)}</div>; })()}</td>
                <td style={{ padding: "0 8px" }}><textarea style={{ ...INL, minHeight: 52, resize: "vertical", lineHeight: 1.6 }} rows={2} placeholder="e.g. 3000 4500 6000x3" value={r.lengths} onChange={e => upRow(r.id, "lengths", e.target.value)} /><div style={{ fontSize: 13, color: count > 200 ? "#ef4444" : "#475569", fontFamily: "'Space Mono', monospace", marginTop: 5 }}>{count > 0 ? t("piecesCount", { n: count }) : t("lengthsHint")}</div></td>
                <td style={{ padding: "0 8px" }}><input style={INL} type="number" placeholder={String(DEFAULT_STOCK)} value={r.stock} onChange={e => upRow(r.id, "stock", e.target.value)} /><div style={{ fontSize: 13, color: "#475569", fontFamily: "'Space Mono', monospace", marginTop: 5 }}>{t("blankEq", { len: fmtMm(DEFAULT_STOCK) })}</div></td>
                <td style={{ padding: "0 8px", verticalAlign: "top" }}>{rows.length > 1 && <button style={{ background: "#1a2230", border: "1px solid #2d3748", borderRadius: 4, color: "#94a3b8", cursor: "pointer", fontSize: 20, width: 40, height: 48 }} onClick={() => delRow(r.id)}>×</button>}</td>
              </tr>); })}</tbody>
          </table></div>
          <button style={{ marginTop: 8, padding: "12px 22px", background: "rgba(245,158,11,.15)", border: "1px dashed #d97706", color: "#fbbf24", borderRadius: 4, cursor: "pointer", fontSize: 15 }} onClick={addRow}>{t("addProfile")}</button>
          <p style={{ fontSize: 13, color: "#475569", fontFamily: "'Space Mono', monospace", marginTop: 14, lineHeight: 1.6 }}>{t("secManualHelp")}</p>
        </Card>
      )}
      {inputMode === "excel" && (
        <Card title={t("uploadTeklaTtl")} style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} labelKey="changeMethod" />}>
          <div onClick={() => fileRef.current?.click()} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleFile(f); }} style={{ border: "2px dashed #2d3748", borderRadius: 10, padding: "44px 24px", textAlign: "center", cursor: "pointer", background: "rgba(15,19,24,.5)", transition: "border-color .2s" }} onMouseEnter={e => e.currentTarget.style.borderColor = "#f59e0b"} onMouseLeave={e => e.currentTarget.style.borderColor = "#2d3748"}><div style={{ fontSize: 48, marginBottom: 14 }}>📊</div><div style={{ fontFamily: "'Playfair Display', serif", fontSize: 20, fontWeight: 700, color: "#f8fafc", marginBottom: 8 }}>{t("upTitleS")}</div><div style={{ fontSize: 16, color: "#cbd5e1", lineHeight: 1.6, maxWidth: 460, margin: "0 auto 6px" }}>{t("upBodyS")}</div><div style={{ fontSize: 14, color: "#94a3b8", marginTop: 6 }}>{t("upAcceptS")}</div><div style={{ fontSize: 13, color: "#475569", marginTop: 12, lineHeight: 1.7 }}>{t("upHintS")}</div><input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.txt" style={{ display: "none" }} onChange={e => { const f = e.target.files[0]; if (f) handleFile(f); }} /></div>
          {uploadInfo && <div style={{ marginTop: 16 }}><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 14, color: "#22c55e", marginBottom: 12 }}>{t("uploadedOk", { name: uploadInfo.name, rows: uploadInfo.rows, pieces: uploadInfo.pieces })}{uploadInfo.skipped ? <span style={{ color: "#94a3b8" }}>{t("platesIgnored", { n: uploadInfo.skipped, s: pls(uploadInfo.skipped) })}</span> : null}</div>{uploadInfo.unknown && uploadInfo.unknown.length > 0 && <div style={{ fontFamily: "'Space Mono', monospace", fontSize: 14, color: "#fbbf24", marginBottom: 12, lineHeight: 1.6 }}>{t("unknownRows", { list: uploadInfo.unknown.slice(0, 12).map(u => `${u.name} ×${u.qty}${u.kg > 0 ? ` (${fmtKg(u.kg)})` : ""}`).join(" · ") + (uploadInfo.unknown.length > 12 ? " …" : "") })}</div>}<div style={{ marginBottom: 14 }}><Label>{t("lblKerf")}</Label><input type="number" value={kerf} onChange={e => setKerf(+e.target.value)} style={{ ...IN, width: 90 }} /></div><table style={{ width: "100%", borderCollapse: "collapse" }}><thead><tr>{[t("lblProfile"), t("lblGradeShort"), t("thLengths"), t("thQty"), t("thKgM")].map(h => <th key={h} style={{ textAlign: "left", padding: "8px 10px", borderBottom: "1px solid #2d3748", color: "#475569", fontSize: 12, letterSpacing: 1, fontFamily: "'Space Mono', monospace" }}>{h}</th>)}</tr></thead><tbody>{cutList.slice(0, 40).map((r, i) => { const sec = findSection(r.profile); return <tr key={i}><td style={TD}>{r.profile}</td><td style={TD}>{r.grade || "—"}</td><td style={TD}>{fmtMm(r.length)}</td><td style={TD}>{r.qty}</td><td style={{ ...TD, color: sec ? "#22c55e" : "#475569" }}>{sec ? sec.kgm : "—"}{r.fkgm && (!sec || Math.abs(r.fkgm - sec.kgm) / r.fkgm > 0.02) ? <span style={{ color: "#94a3b8" }}> · {t("fileKgmCell", { v: r.fkgm.toFixed(2) })}</span> : null}</td></tr>; })}</tbody></table>{cutList.length > 40 && <div style={{ fontSize: 13, color: "#475569", fontFamily: "'Space Mono', monospace", marginTop: 8 }}>{t("andMore", { n: cutList.length - 40 })}</div>}</div>}
        </Card>
      )}
      {inputMode !== null && <PriceCard pricing={pricing} />}
      {inputMode !== null && (
        <Card title={t("reuseBarsTtl")} style={{ marginTop: 20 }}>
          <div style={{ fontSize: 15, color: "#6ee7b7", marginBottom: 14, lineHeight: 1.6 }}>{t("reuseBarsLead")}</div>
          <LeftoverImporter expects="sections" onImported={importLeftovers} />
          {leftovers.length === 0 && (
            <div style={{ padding: "20px 18px", textAlign: "center", border: "1px dashed rgba(16,185,129,.35)", borderRadius: 12, background: "rgba(16,185,129,.04)" }}>
              <div style={{ fontSize: 15, color: "#94a3b8", marginBottom: 12 }}>{t("noBarsYet")}</div>
              <button onClick={addLeftover} style={{ padding: "10px 22px", background: "linear-gradient(135deg,#10b981,#059669)", border: "none", color: "#04140d", borderRadius: 8, cursor: "pointer", fontSize: 15, fontWeight: 700 }}>{t("addLeftBar")}</button>
            </div>
          )}
          {leftovers.length > 0 && (
            <>
              <div style={{ overflowX: "visible" }}><table style={{ width: "100%", borderCollapse: "separate", borderSpacing: "0 10px" }}><thead><tr>{[t("lblProfile"), t("lblGradeShort"), t("loLeftoverLen"), t("lblQtyHow"), ""].map((h, i) => <th key={h} style={{ textAlign: "left", padding: "0 8px", color: "#64748b", fontSize: 13, letterSpacing: 1, width: i === 0 ? "30%" : i === 1 ? "18%" : i === 4 ? 40 : "auto" }}>{h}</th>)}</tr></thead>
                <tbody>{leftovers.map(r => { const matches = loAuto === r.id ? searchSections(r.profile) : []; const sec = findSection(r.profile); return (
                  <tr key={r.id}>
                    <td style={{ position: "relative", padding: "0 8px" }}><input style={{ ...INL, borderColor: "#10b98155" }} placeholder="IPE 300" value={r.profile} onChange={e => { upLeftover(r.id, "profile", e.target.value); setLoAuto(r.id); }} onFocus={() => setLoAuto(r.id)} onBlur={() => setTimeout(() => setLoAuto(a => a === r.id ? null : a), 150)} />{sec && <div style={{ fontSize: 13, color: sec.computed ? "#60a5fa" : "#22c55e", fontFamily: "'Space Mono', monospace", marginTop: 5 }}>{sec.kgm} {t("thKgM")} · {sec.intl ? soIntlBadge(sec) : sec.type}</div>}{matches.length > 0 && <div style={{ position: "absolute", top: "100%", left: 8, right: 8, zIndex: 50, background: "#0f1318", border: "1px solid #2d3748", borderRadius: 6, maxHeight: 260, overflowY: "auto", boxShadow: "0 8px 24px rgba(0,0,0,.5)" }}>{matches.map(m => <div key={m.name} onMouseDown={() => { upLeftover(r.id, "profile", m.name); setLoAuto(null); }} style={{ padding: "11px 14px", cursor: "pointer", display: "flex", justifyContent: "space-between", borderBottom: "1px solid #1a2230", fontFamily: "'Space Mono', monospace", fontSize: 15 }}><span style={{ color: "#cbd5e1" }}>{m.name}</span><span style={{ color: "#f59e0b" }}>{m.kgm} {t("thKgM")}</span></div>)}</div>}</td>
                    <td style={{ padding: "0 8px" }}><input style={INL} placeholder="S355JR" value={r.grade} onChange={e => upLeftover(r.id, "grade", e.target.value)} /></td>
                    <td style={{ padding: "0 8px" }}><input style={INL} type="number" placeholder="6000" value={r.length} onChange={e => upLeftover(r.id, "length", e.target.value)} /><div style={{ fontSize: 13, color: "#475569", fontFamily: "'Space Mono', monospace", marginTop: 5 }}>{t("loBarLeft")}</div></td>
                    <td style={{ padding: "0 8px" }}><input style={{ ...INL, borderColor: "#10b98199" }} type="number" min="1" placeholder="1" value={r.qty} onChange={e => upLeftover(r.id, "qty", e.target.value)} /><div style={{ fontSize: 13, color: "#475569", fontFamily: "'Space Mono', monospace", marginTop: 5 }}>{t("loHowManyLike")}</div></td>
                    <td style={{ padding: "0 8px", verticalAlign: "top" }}><button style={{ background: "#1a2230", border: "1px solid #2d3748", borderRadius: 4, color: "#94a3b8", cursor: "pointer", fontSize: 20, width: 40, height: 48 }} onClick={() => delLeftover(r.id)}>×</button></td>
                  </tr>); })}</tbody>
              </table></div>
              <button onClick={addLeftover} style={{ marginTop: 4, padding: "9px 20px", background: "rgba(16,185,129,.12)", border: "1px dashed #10b981", color: "#6ee7b7", borderRadius: 8, cursor: "pointer", fontSize: 15, fontWeight: 700 }}>{t("addLeftMore")}</button>
              <p style={{ fontSize: 13, color: "#475569", fontFamily: "'Space Mono', monospace", marginTop: 12, lineHeight: 1.6 }}>{t("loSameProfile")}</p>
            </>
          )}
        </Card>
      )}
      {error && <div style={{ background: "#450a0a", border: "1px solid #7f1d1d", borderRadius: 6, padding: "12px 16px", color: "#fca5a5", fontSize: 14, fontFamily: "'Space Mono', monospace", margin: "12px 0" }}>{error}</div>}
      {inputMode === "manual" && <div style={{ textAlign: "center", marginTop: 24 }}><button onClick={() => { const l = buildManual(); if (!l.length) { setError(t("secEnterLen")); return; } soRequireMember(() => runOpt(l), true); }} style={OPT_BTN(false)}>{t("optSections")}</button></div>}
      {inputMode === "scan" && (
        <Card title={t("scanCardTtl")} style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} labelKey="changeMethod" />}>
          <ScanImporter mode="sections" onUse={rs => {
            const list = rs.map(r => ({ profile: String(r.profile || "").trim().toUpperCase(), grade: r.grade || "", length: r.length, qty: r.qty, stock: DEFAULT_STOCK, stockAuto: true }));
            setCutList(list);
            setUploadInfo({ name: t("scanCardTtl"), rows: list.length, pieces: list.reduce((s, r) => s + r.qty, 0), skipped: 0 });
            scrollToOpt();
          }} />
        </Card>
      )}
      {(inputMode === "excel" || inputMode === "scan") && uploadInfo && <div ref={optBtnRef} style={{ textAlign: "center", marginTop: 24, scrollMarginTop: 80 }}><button onClick={() => soRequireMember(() => runOpt(cutList), true)} style={OPT_BTN(false)}>{t("optSections")}</button></div>}
    </>
  );
}

function SectionResults({ results, kerf, pricing, onBack }) {
  const { t, lang } = useLang();
  const pls = n => ((lang === "en" || lang === "es") && n > 1) ? "s" : "";
  const topRef = useRef(null);
  useEffect(() => { const t = setTimeout(() => topRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }), 60); return () => clearTimeout(t); }, []);
  const totalStock = results.groups.reduce((s, g) => s + g.summary.stockCount, 0);
  const stockMm = results.groups.reduce((s, g) => s + g.summary.totalStock, 0);
  const wasteMm = results.groups.reduce((s, g) => s + g.summary.totalWaste, 0);
  const wastePct = stockMm ? ((wasteMm / stockMm) * 100).toFixed(1) : "0";
  const wasteKg = results.groups.reduce((s, g) => s + (g.kgm ? (g.summary.totalWaste / 1000) * g.kgm : 0), 0);
  const hasLeftovers = results.groups.some(g => g.bins.some(b => b.remaining >= 1000) || (g.reusedBins || []).some(b => b.cuts.length && b.remaining >= 1000));
  const reusedBarsUsed = results.groups.reduce((s, g) => s + (g.summary.reusedCount || 0), 0);
  const reusedSavedKg = results.groups.reduce((s, g) => s + (g.reusedKg || 0), 0);
  const offcutKg = sectionLeftovers(results).reduce((s, r) => s + (r.kg || 0), 0);
  return (
    <>
      <div ref={topRef} style={{ scrollMarginTop: 12 }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24, flexWrap: "wrap", gap: 12 }}><SectionTitle>{t("optResults")}</SectionTitle><BackChip onClick={onBack} labelKey="backInput" /></div>
      {/* ── SUMMARY: what you need to buy (directly below header) ── */}
      <div style={{ background: "#1c1600", border: "2px solid #f59e0b", borderRadius: 12, padding: "20px 24px", marginBottom: 16 }}>
        <div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#f59e0b", letterSpacing: ".15em", marginBottom: 14 }}>{t("summaryBuy")}</div>
        {results.groups.map((g, gi) => { const byLen = g.bins.reduce((a, b) => { a[b.stockLength] = (a[b.stockLength] || 0) + 1; return a; }, {}); const parts = Object.entries(byLen).map(([len, qty]) => `${qty} × ${fmtMm(parseInt(len))}`).join(", "); return <div key={gi} style={{ display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap", padding: "8px 0", borderBottom: gi < results.groups.length - 1 ? "1px dashed #3a2e0a" : "none" }}><span style={{ fontFamily: "'Playfair Display', serif", fontWeight: 900, fontSize: 29, color: "#f59e0b", lineHeight: 1 }}>{g.summary.stockCount}</span><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#cbd5e1" }}>{soCountWord(t, g.summary.stockCount, "bar", "bars")} {soIsExtra(SO_LANG) ? soUiTx(SO_LANG, "MISC.barOf", t("sheetOf")) : t("sheetOf")}</span><span style={{ fontFamily: "'Playfair Display', serif", fontWeight: 700, fontSize: 22, color: "#f8fafc" }}>{g.profile}</span>{g.grade && <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 14, color: "#94a3b8" }}>({soGrade(g.grade)})</span>}<span style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#64748b" }}>— {parts}</span>{g.stockKg != null && <span style={{ marginLeft: "auto", fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#22c55e", fontWeight: 700 }}>{fmtKg(g.stockKg)}</span>}</div>; })}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginTop: 14, paddingTop: 12, borderTop: "1px solid #3a2e0a" }}><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#cbd5e1", fontWeight: 700 }}>{t("totalAcrossProf", { n: totalStock, g: results.groups.length, s: pls(results.groups.length) })}</span><span style={{ fontFamily: "'Playfair Display', serif", fontSize: 25, fontWeight: 900, color: "#f59e0b" }}>{fmtKg(results.grandKg)}</span></div>
      </div>
      {/* ── WASTE AFTER CUTTING (directly below summary) ── */}
      <div style={{ background: parseFloat(wastePct) <= 8 ? "linear-gradient(135deg,#16a34a,#15803d)" : parseFloat(wastePct) <= 20 ? "linear-gradient(135deg,#f59e0b,#d97706)" : "linear-gradient(135deg,#dc2626,#991b1b)", borderRadius: 12, padding: "20px 24px", marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}><div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#fff", opacity: .85, letterSpacing: ".1em" }}>{t("wasteAfterCut")}</div><div style={{ fontFamily: "'Playfair Display', serif", fontSize: 44, fontWeight: 900, color: "#fff" }}>{wastePct}%</div></div><div style={{ textAlign: "right" }}><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#fff", opacity: .85 }}>{t("offcutLength")}</div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 20, color: "#fff", fontWeight: 700 }}>{fmtMm(Math.round(wasteMm))}</div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#fff", opacity: .9, marginTop: 4 }}>{fmtKg(wasteKg)} {t("scrapReuse")}</div></div></div>
      <CostBlock buyKg={results.grandKg} netKg={Math.max(results.grandKg - wasteKg, 0)} wasteKg={wasteKg} offcutKg={offcutKg} pricing={pricing} />
      <PaintBlock total={results.paintM2} note={t("paintNoteS")} missing={results.paintMissing} count={results.groups.length} />
      {reusedBarsUsed > 0 && (
        <div style={{ background: "linear-gradient(135deg,#065f46,#047857)", borderRadius: 12, padding: "16px 22px", marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
          <div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#d1fae5", letterSpacing: ".1em" }}>{t("leftoversUsedHd")}</div><div style={{ fontFamily: "'Playfair Display', serif", fontSize: 25, fontWeight: 900, color: "#fff" }}>{t("leftoversUsedBd", { n: reusedBarsUsed, s: pls(reusedBarsUsed) })}</div></div>
          {reusedSavedKg > 0 && <div style={{ textAlign: "right" }}><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#d1fae5" }}>{t("steelNotBought")}</div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 22, color: "#fff", fontWeight: 700 }}>{fmtKg(reusedSavedKg)}</div></div>}
        </div>
      )}
      {results.groups.map((g, gi) => (
        <div key={gi} style={{ marginBottom: 32 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16, paddingBottom: 8, borderBottom: "1px solid #1e293b", flexWrap: "wrap" }}><span style={{ fontFamily: "'Playfair Display', serif", fontSize: 25, fontWeight: 700, color: "#f8fafc" }}>{g.profile}</span>{g.srcName && <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#64748b" }}>{t("fromFile", { name: g.srcName })}</span>}{g.grade && <span style={TAG}>{soGrade(g.grade)}</span>}{g.kgm && <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#94a3b8" }}>{g.kgm} {t("thKgMu")}{g.computed ? t("computedTag") : ""}{g.kgmFromFile ? t("fromFileTag") : ""}{g.paintM2pm != null ? ` · ${fmtM2pm(g.paintM2pm)} ${t("m2pmU")}` : ""}</span>}{g.fileDiff != null && Math.abs(g.fileDiff) > 0.02 && <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: Math.abs(g.fileDiff) > 0.05 ? "#fbbf24" : "#64748b" }}>{t(Math.abs(g.fileDiff) > 0.05 ? "fileKgmWarn" : "fileKgmNote", { v: g.fileKgm, d: (g.fileDiff > 0 ? "+" : "") + (g.fileDiff * 100).toFixed(1) })}</span>}{g.stockKg != null && <span style={{ marginLeft: "auto", fontFamily: "'Space Mono', monospace", fontSize: 16, color: "#f59e0b", fontWeight: 700 }}>{fmtKg(g.stockKg)}</span>}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(110px,1fr))", gap: 8, marginBottom: 16 }}>{[{ v: g.summary.stockCount, l: t("stBarsBuy") }, { v: fmtMm(g.summary.totalStock), l: t("stTotalLen") }, { v: fmtMm(g.summary.totalWaste), l: t("stOffcut") }, { v: `${g.summary.wastePct}%`, l: t("stWaste") }, { v: `${g.summary.utilPct}%`, l: t("stUtil") }, { v: g.stockKg != null ? fmtKg(g.stockKg) : "—", l: t("stWeight") }, ...(results.paintM2 != null ? [{ v: g.paintM2 != null ? `${fmtM2(g.paintM2)} ${t("m2")}` : "—", l: t("stPaint") }] : [])].map(s => <div key={s.l} style={ST}><div style={STV}>{s.v}</div><div style={STL}>{s.l}</div></div>)}</div>
          <div style={{ background: "#1c1600", border: "1px solid #f59e0b", borderRadius: 10, padding: 18, marginBottom: 16 }}><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#f59e0b", letterSpacing: ".15em", marginBottom: 14 }}>{t("whatToOrderTtl")}</div>{Object.entries(g.bins.reduce((a, b) => { a[b.stockLength] = (a[b.stockLength] || 0) + 1; return a; }, {})).map(([len, qty]) => { const kg = g.kgm ? (parseInt(len) / 1000) * g.kgm * qty : null; return <div key={len} style={{ display: "flex", alignItems: "baseline", gap: 10, flexWrap: "wrap", marginBottom: 10, paddingBottom: 10, borderBottom: "1px dashed #3a2e0a" }}><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 14, color: "#94a3b8" }}>{t("youNeedWord")}</span><span style={{ fontFamily: "'Playfair Display', serif", fontWeight: 900, fontSize: 34, color: "#f59e0b", lineHeight: 1 }}>{qty}</span><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 14, color: "#94a3b8" }}>×</span><span style={{ fontFamily: "'Playfair Display', serif", fontWeight: 700, fontSize: 25, color: "#f8fafc" }}>{g.profile}</span>{g.srcName && <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 14, color: "#64748b" }}>({g.srcName})</span>}<span style={{ fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#cbd5e1" }}>@ {fmtMm(parseInt(len))}</span>{g.grade && <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 14, color: "#94a3b8" }}>{t("gradeWord")} {soGrade(g.grade)}</span>}{kg && <span style={{ marginLeft: "auto", fontFamily: "'Space Mono', monospace", fontSize: 15, color: "#22c55e", fontWeight: 700 }}>{fmtKg(kg)}</span>}</div>; })}</div>
          {g.splices && g.splices.length > 0 && <div style={{ background: "#0c1f14", border: "1px solid #15803d", borderRadius: 10, padding: 18, marginBottom: 16 }}><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#4ade80", letterSpacing: ".1em", marginBottom: 10 }}>{t("longMembersHd", { len: fmtMm(g.splices[0].stock) })}</div>{g.splices.map((sp, oi) => <div key={oi} style={{ fontSize: 15, color: "#bbf7d0", marginBottom: 4 }}><span style={{ fontFamily: "'Space Mono', monospace", fontWeight: 700 }}>{sp.qty} × </span>{fmtMm(sp.length)}<span style={{ color: "#4ade80", fontSize: 13 }}>{t("eachFromBars", { n: sp.bars, len: fmtMm(sp.stock) })}</span></div>)}</div>}
          {g.reusedBins && g.reusedBins.some(b => b.cuts.length) && <Card title={t("cutFromLeftTtl")} style={{ marginBottom: 16 }}><div style={{ fontSize: 14, color: "#6ee7b7", marginBottom: 12, background: "rgba(16,185,129,.08)", border: "1px solid rgba(16,185,129,.25)", borderRadius: 6, padding: "8px 12px" }}>{t("cutFromLeftLead")}{g.reusedRemainKg ? t("leftRemaining", { kg: fmtKg(g.reusedRemainKg) }) : null}</div>{g.reusedBins.filter(b => b.cuts.length).map((b, i) => <CutBar key={i} bin={b} index={i} />)}</Card>}
          <Card title={t("visualCutPlan")}>{g.bins.length ? g.bins.map((b, i) => <CutBar key={i} bin={b} index={i} />) : <div style={{ color: "#64748b", fontFamily: "'Space Mono', monospace", fontSize: 14 }}>{g.reusedBins && g.reusedBins.some(x => x.cuts.length) ? t("allFromLeft") : t("noPlan")}</div>}</Card>
        </div>
      ))}
      {(() => {
        const REUSE_MIN = 1000; // a leftover bar >= 1 m is worth keeping
        const rows = [];
        results.groups.forEach(g => {
          g.bins.forEach((b, i) => { if (b.remaining >= REUSE_MIN) rows.push({ profile: g.profile, grade: g.grade, bar: t("barNum", { n: i + 1 }), stock: b.stockLength, len: Math.round(b.remaining), kg: g.kgm ? (b.remaining / 1000) * g.kgm : null }); });
          (g.reusedBins || []).forEach((b, i) => { if (b.cuts.length && b.remaining >= REUSE_MIN) rows.push({ profile: g.profile, grade: g.grade, bar: t("leftoverNum", { n: i + 1 }), stock: b.stockLength, len: Math.round(b.remaining), kg: g.kgm ? (b.remaining / 1000) * g.kgm : null }); });
        });
        const totalKg = rows.reduce((s, r) => s + (r.kg || 0), 0);
        if (!rows.length) return null;
        return (
          <Card title={t("reusableBarTtl")} style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 14, color: "#6ee7b7", marginBottom: 12, background: "rgba(16,185,129,.08)", border: "1px solid rgba(16,185,129,.25)", borderRadius: 6, padding: "8px 12px" }}>{t("reusableBarLead", { n: rows.length, s: pls(rows.length), kg: fmtKg(totalKg) })}</div>
            <div style={{ overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 15, fontFamily: "'Space Mono', monospace" }}><thead><tr style={{ background: "rgba(16,185,129,.12)", borderBottom: "1px solid rgba(16,185,129,.3)" }}>{[t("lblProfile"), t("lblGradeShort"), t("thFromBar"), t("thStockLength"), t("thReusableLeft"), t("thWeight"), t("thStatus")].map((h, i) => <th key={h} style={{ padding: "10px 12px", textAlign: i > 4 ? "right" : "left", color: "#6ee7b7", fontSize: 13, letterSpacing: 1, fontWeight: 700 }}>{h}</th>)}</tr></thead>
              <tbody>{rows.map((r, i) => <tr key={i} style={{ borderBottom: "1px solid #1a2230" }}><td style={{ padding: "9px 12px", color: "#cbd5e1" }}>{r.profile}</td><td style={{ padding: "9px 12px", color: "#94a3b8" }}>{soGrade(r.grade) || "—"}</td><td style={{ padding: "9px 12px", color: "#94a3b8" }}>{r.bar}</td><td style={{ padding: "9px 12px", color: "#94a3b8" }}>{fmtMm(r.stock)}</td><td style={{ padding: "9px 12px", color: "#6ee7b7", fontWeight: 700 }}>{fmtMm(r.len)}</td><td style={{ padding: "9px 12px", textAlign: "right", color: "#94a3b8" }}>{r.kg != null ? fmtKg(r.kg) : "—"}</td><td style={{ padding: "9px 12px", textAlign: "right", color: "#10B981", fontWeight: 700 }}>{t("keepStatus")}</td></tr>)}</tbody>
            </table></div>
          </Card>
        );
      })()}
      <Card title={t("dlReports")}>
        <div style={{ fontSize: 14, color: "#94a3b8", marginBottom: 12, lineHeight: 1.5 }}>{t("dlTwoDeliv")}</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button onClick={() => requestDownload(() => exportSectionPDF(results, kerf, { totalStock, wastePct, wasteMm, wasteKg }, pricing))} style={EXP}>{t("dlProcPDF")}</button>
          <button onClick={() => requestDownload(() => exportSectionExcel(results, pricing))} style={EXP}>{t("dlProcXLS")}</button>
          {hasLeftovers && <button onClick={() => requestDownload(() => exportSectionLeftoverPDF(results))} style={{ ...EXP, background: "linear-gradient(135deg,#10b981,#059669)", color: "#04140d" }}>{t("dlLeftPDF")}</button>}
          {hasLeftovers && <button onClick={() => requestDownload(() => exportSectionLeftoverExcel(results))} style={{ ...EXP, background: "linear-gradient(135deg,#10b981,#059669)", color: "#04140d" }}>{t("dlLeftXLS")}</button>}
        </div>
      </Card>
    </>
  );
}

/* ─── SHARED UI ──────────────────────────────────────────────────────────── */
function Card({ title, children, style, action }) { return <div style={{ background: "rgba(19,25,32,.72)", backdropFilter: "blur(12px)", WebkitBackdropFilter: "blur(12px)", border: "1px solid rgba(148,163,184,.14)", borderRadius: 14, overflow: "hidden", boxShadow: "0 12px 40px -12px rgba(0,0,0,.55)", ...style }}><div style={{ padding: "13px 18px", borderBottom: "1px solid rgba(148,163,184,.12)", background: "linear-gradient(90deg, rgba(245,158,11,.08), rgba(245,158,11,.01))", display: "flex", alignItems: "center", justifyContent: "space-between" }}><span style={{ fontSize: 15, fontWeight: 700, color: "#fbbf24", letterSpacing: .5, fontFamily: "'Space Mono', monospace" }}>{title}</span>{action}</div><div style={{ padding: 18 }}>{children}</div></div>; }
function Label({ children }) { return <div style={{ fontSize: 12, color: "#64748b", letterSpacing: 1, textTransform: "uppercase", marginBottom: 5, fontFamily: "'Space Mono', monospace" }}>{children}</div>; }
function StatCard({ l, v, i, a }) { return <div style={{ padding: "16px 14px", borderRadius: 8, background: "#0f1318", border: "1px solid #1e293b", textAlign: "center" }}><div style={{ fontSize: 22, marginBottom: 4 }}>{i}</div><div style={{ fontSize: 22, fontWeight: 800, color: a || "#cbd5e1", fontFamily: "'Space Mono', monospace" }}>{v}</div><div style={{ fontSize: 12, color: "#64748b", marginTop: 4, letterSpacing: 1 }}>{l}</div></div>; }

const IN = { background: "#0f1318", border: "1px solid #2d3748", color: "#cbd5e1", padding: "7px 10px", borderRadius: 4, fontSize: 15, width: 110, fontFamily: "'Space Mono', monospace", outline: "none" };
const ZB = { width: 28, height: 28, borderRadius: 4, border: "1px solid #2d3748", background: "rgba(15,19,24,.8)", color: "#94a3b8", cursor: "pointer", fontSize: 18, lineHeight: 1 };
const INL = { background: "#0f1318", border: "1.5px solid #334155", color: "#cbd5e1", padding: "14px 14px", borderRadius: 4, fontSize: 17, width: "100%", boxSizing: "border-box", fontFamily: "'Space Mono', monospace", outline: "none" };
const SEL = { background: "#0f1318", border: "1px solid #2d3748", color: "#cbd5e1", padding: "7px 10px", borderRadius: 4, fontSize: 15, fontFamily: "'Space Mono', monospace", cursor: "pointer", outline: "none" };
const CI = { background: "#0f1318", border: "1px solid #2d3748", color: "#cbd5e1", padding: "5px 7px", borderRadius: 3, fontSize: 14, fontFamily: "'Space Mono', monospace", outline: "none" };
const BTN = { padding: "8px 18px", borderRadius: 4, border: "1px solid #d97706", background: "rgba(245,158,11,.15)", color: "#fbbf24", cursor: "pointer", fontSize: 14, marginTop: 6 };
const EXP = { padding: "11px 24px", borderRadius: 6, border: "none", background: "linear-gradient(135deg,#f59e0b,#d97706)", color: "#1a1206", cursor: "pointer", fontSize: 15, fontWeight: 700, fontFamily: "'Space Mono', monospace" };
const TAG = { display: "inline-block", padding: "4px 10px", borderRadius: 20, fontSize: 13, fontFamily: "'Space Mono', monospace", background: "rgba(245,158,11,.15)", border: "1px solid #f59e0b", color: "#f59e0b" };
const TD = { padding: "9px 10px", borderBottom: "1px solid #1a2230", color: "#cbd5e1", fontFamily: "'Space Mono', monospace", fontSize: 14 };
const ST = { background: "#0f1318", border: "1px solid #1e293b", borderRadius: 8, padding: "16px 14px", textAlign: "center" };
const STV = { fontSize: 22, fontWeight: 900, fontFamily: "'Space Mono', monospace", color: "#f59e0b", lineHeight: 1.1 };
const STL = { fontSize: 11, color: "#64748b", letterSpacing: ".08em", marginTop: 4, textTransform: "uppercase" };
const OPT_BTN = isOpt => ({ padding: "16px 60px", fontSize: 18, fontWeight: 800, background: isOpt ? "rgba(120,80,10,.4)" : "linear-gradient(135deg,#f59e0b,#d97706)", border: "none", color: "#1a1206", borderRadius: 8, cursor: isOpt ? "default" : "pointer", boxShadow: "0 8px 30px rgba(245,158,11,.4)", letterSpacing: 2, fontFamily: "'Space Mono', monospace" });

/* ─── EXPORTS ────────────────────────────────────────────────────────────── */
function downloadHTML(html, name) { const win = window.open("", "_blank"); if (win && win.document) { win.document.open(); win.document.write(html); win.document.close(); } else { const blob = new Blob([html], { type: "text/html" }); const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = name; document.body.appendChild(a); a.click(); document.body.removeChild(a); setTimeout(() => URL.revokeObjectURL(url), 4000); } }
function saveWorkbook(wb, name) { rtWorkbook(wb); try { XLSX.writeFile(wb, name); } catch { try { const out = XLSX.write(wb, { bookType: "xlsx", type: "array" }); const blob = new Blob([out], { type: "application/octet-stream" }); const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = name; document.body.appendChild(a); a.click(); document.body.removeChild(a); setTimeout(() => URL.revokeObjectURL(url), 4000); } catch (e) { alert(RT("Export failed: ", "تعذّر التصدير: ") + (e.message || e)); } } }

function exportPlateExcel(results, material, reuseMin, t, pricing) {
  const wb = XLSX.utils.book_new();
  const summary = [[RT("Steel Optimizer — Plate Procurement", "Steel Optimizer — توريد الألواح")], [], [RT("Material", "المعدن"), material], [],
    [RT("Thickness (mm)", "السماكة (مم)"), RT("Sheet Size", "مقاس اللوح"), RT("Sheets Req.", "الألواح المطلوبة"), RT("Total Wt (t)", "الوزن الإجمالي (طن)"), RT("Parts", "القطع"), RT("Utilization %", "نسبة الاستغلال %"), RT("Net Wt (t)", "الوزن الصافي (طن)"), RT("Scrap Wt (t)", "وزن السكراب (طن)"), RT("Paint Area 1 face (m²)", "مساحة الدهان وجه واحد (م²)")],
    ...results.groups.map(g => [g.thickness, `${g.sw}x${g.sh}`, g.sheetCount, (g.sheetWeight / 1000).toFixed(2), g.partCount, g.utilPct, (g.partWeight / 1000).toFixed(2), (g.wasteWeight / 1000).toFixed(2), g.paintM2 != null ? upM2(g.paintM2, 1) : "—"]),
    [RT("TOTAL", "الإجمالي"), RT(`${results.groups.length} thicknesses`, `${results.groups.length} سماكة`, (L) => L === "ru" ? `толщин: ${results.groups.length}` : L === "zh" ? `${results.groups.length} 种厚度` : `espesores: ${results.groups.length}`), results.totals.sheets, (results.totals.sheetWeight / 1000).toFixed(2), results.totals.parts, results.totals.utilPct, (results.totals.partWeight / 1000).toFixed(2), (results.totals.wasteWeight / 1000).toFixed(2), results.totals.paintM2 != null ? upM2(results.totals.paintM2, 1) : "—"],
    [], [RT("TOTAL PURCHASE WEIGHT - gross full sheets (Ton)", "إجمالي وزن الشراء — ألواح كاملة (طن)"), (results.totals.sheetWeight / 1000).toFixed(2)],
    [RT("NET PARTS WEIGHT - finished plates only (Ton)", "الوزن الصافي للقطع — القطع النهائية فقط (طن)"), (results.totals.partWeight / 1000).toFixed(2)],
    [RT("SCRAP / OFFCUT WEIGHT (Ton)", "وزن السكراب / القصاصات (طن)"), (results.totals.wasteWeight / 1000).toFixed(2)],
    [RT("REUSABLE OFFCUTS", "القصاصات القابلة لإعادة الاستخدام"), `${results.totals.offcutCount} ${RT("pieces", "قطعة")}`, `${(results.totals.offcutWeight / 1000).toFixed(2)} ${RT("t", "طن")}`],
    [RT("PAINT AREA - one face of your parts, L x W x qty (m²)", "مساحة الدهان — وجه واحد لقطعك، الطول × العرض × العدد (م²)"), results.totals.paintM2 != null ? upM2(results.totals.paintM2, 1) : "—", RT("offcuts and scrap not included", "بدون البواقي والهالك")]];
  if (pricing && pricing.active) {
    const P = pricing.pricePerTon, M = v => `${Math.round(v).toLocaleString("en-US")} ${pricing.currency}`;
    summary.push([], [RT("STEEL PRICE (per ton)", "سعر الحديد (للطن)"), `${Math.round(P).toLocaleString("en-US")} ${pricing.currency}`, RT("your entered price", "السعر الذي أدخلته")],
      [RT("TOTAL PURCHASE COST - gross full sheets", "تكلفة الشراء الإجمالية — ألواح كاملة"), M(results.totals.sheetWeight / 1000 * P)],
      [RT("NET USED VALUE - finished plates only", "قيمة الحديد المستخدم — القطع النهائية فقط"), M(results.totals.partWeight / 1000 * P)],
      [RT("SCRAP COST - money lost in waste", "تكلفة الهالك — المال الضائع في الهدر"), M(results.totals.wasteWeight / 1000 * P)],
      [RT("REUSABLE OFFCUT VALUE - recoverable", "قيمة القصاصات الصالحة — قابلة للاسترداد"), M(results.totals.offcutWeight / 1000 * P)]);
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), RT("Summary", "الملخص"));
  const off = [[RT(`Reusable Offcut Register (each side >= ${reuseMin} mm)`, `سجل القصاصات القابلة لإعادة الاستخدام (كل ضلع ≥ ${reuseMin} مم)`, (L) => L === "ru" ? `Реестр деловых остатков (каждая сторона >= ${reuseMin} мм)` : L === "zh" ? `可再利用余料清单（每边 >= ${reuseMin} mm）` : `Registro de sobrantes reutilizables (cada lado >= ${reuseMin} mm)`)], [], [RT("Thickness (mm)", "السماكة (مم)"), RT("From Sheet", "من اللوح"), RT("Offcut W (mm)", "عرض القصاصة (مم)"), RT("Offcut L (mm)", "طول القصاصة (مم)"), RT("Weight (kg)", "الوزن (كجم)"), RT("Status", "الحالة")],
    ...results.groups.flatMap(g => g.offcuts.map(o => [g.thickness, RT(`Sheet ${o.sheet}`, `لوح ${o.sheet}`, (L) => L === "ru" ? `Лист ${o.sheet}` : L === "zh" ? `第 ${o.sheet} 张` : `Chapa ${o.sheet}`), o.w, o.h, o.weight.toFixed(2), RT("keep for future jobs", "احتفظ بها لمشاريع قادمة")]))];
  if (results.totals.offcutCount === 0) off.push([RT("No reusable offcuts — all leftovers below the minimum size.", "لا توجد قصاصات صالحة — كل البواقي أصغر من الحد الأدنى.")]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(off), RT("Reusable Offcuts", "القصاصات الصالحة"));
  // Cutting plan (placement coordinates per sheet) — the 2D cut plan as data.
  // A to-scale graphic of this same layout is in the PDF report.
  const plan = [[RT("CUTTING PLAN — part placements per sheet (origin = top-left corner, mm)", "خطة القص — مواقع القطع في كل لوح (نقطة الأصل = الزاوية العلوية اليسرى، مم)")], [],
    [RT("Thickness (mm)", "السماكة (مم)"), RT("Sheet #", "رقم اللوح"), RT("Part Label", "رمز القطعة"), RT("X (mm)", "X (مم)"), RT("Y (mm)", "Y (مم)"), RT("Width (mm)", "العرض (مم)"), RT("Length (mm)", "الطول (مم)"), RT("Rotated 90°", "مُدارة 90°"), RT("Spliced/Welded", "موصولة/ملحومة")]];
  results.groups.forEach(g => (g.sheets || []).forEach((s, si) => (s.placements || []).forEach(p =>
    plan.push([g.thickness, si + 1, p.label || "—", Math.round(p.x), Math.round(p.y), Math.round(p.w), Math.round(p.h), p.rotated ? RT("YES", "نعم") : RT("no", "لا"), p.spliced ? RT("YES", "نعم") : RT("no", "لا")]))));
  if (plan.length <= 3) plan.push([RT("No placements to list.", "لا توجد مواقع لعرضها.")]);
  const wsPlan = XLSX.utils.aoa_to_sheet(plan);
  wsPlan["!cols"] = [{ wch: 14 }, { wch: 8 }, { wch: 16 }, { wch: 9 }, { wch: 9 }, { wch: 11 }, { wch: 12 }, { wch: 11 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, wsPlan, RT("Cutting Plan", "خطة القص"));
  saveWorkbook(wb, "SteelOptimizer_Plates.xlsx");
}

// Re-import payload for plate offcuts: each leftover carries thickness, overall A×B and notch (C,D).
function plateLeftoverPayload(results) {
  const items = [];
  results.groups.forEach(g => g.offcuts.forEach(o => { items.push({ label: `${g.thickness}mm offcut`, thickness: g.thickness, shape: o.shape === "L" ? "L" : "rect", A: Math.round(o.A), B: Math.round(o.B), C: o.shape === "L" ? Math.round(o.notchW) : 0, D: o.shape === "L" ? Math.round(o.notchH) : 0, qty: 1 }); }));
  return { t: "plate", items };
}
function exportOffcutExcel(results, material, reuseMin) {
  const wb = XLSX.utils.book_new();
  const rows = [[RT("Reusable Offcut Register", "سجل القصاصات القابلة لإعادة الاستخدام")], [RT("Material", "المعدن"), material], [RT(`Minimum size kept (each side)`, `أقل مقاس يُحتفظ به (كل ضلع)`), `${reuseMin} ${RT("mm", "مم")}`], [],
    ["#", RT("Thickness (mm)", "السماكة (مم)"), RT("From Sheet", "من اللوح"), RT("Shape", "الشكل"), RT("Overall W (mm)", "العرض الكلي (مم)"), RT("Overall L (mm)", "الطول الكلي (مم)"), RT("Notch W (mm)", "عرض الركن (مم)"), RT("Notch L (mm)", "طول الركن (مم)"), RT("Usable Rect W (mm)", "عرض المستطيل القابل للاستخدام (مم)"), RT("Usable Rect L (mm)", "طول المستطيل القابل للاستخدام (مم)"), RT("Weight (kg)", "الوزن (كجم)"), RT("Status", "الحالة")]];
  let n = 0; results.groups.forEach(g => g.offcuts.forEach(o => { n++; rows.push([n, g.thickness, RT(`Sheet ${o.sheet}`, `لوح ${o.sheet}`, (L) => L === "ru" ? `Лист ${o.sheet}` : L === "zh" ? `第 ${o.sheet} 张` : `Chapa ${o.sheet}`), o.shape === "L" ? RT("L-shape", "شكل L") : RT("Rectangle", "مستطيل"), o.A, o.B, o.shape === "L" ? o.notchW : 0, o.shape === "L" ? o.notchH : 0, o.w, o.h, o.weight.toFixed(2), RT("keep for future jobs", "احتفظ بها لمشاريع قادمة")]); }));
  if (n === 0) rows.push(["—", RT("No reusable offcuts above the minimum size.", "لا توجد قصاصات صالحة أكبر من الحد الأدنى.")]);
  rows.push([], [RT("TOTAL REUSABLE OFFCUTS", "إجمالي القصاصات الصالحة"), `${results.totals.offcutCount} ${RT("pieces", "قطعة")}`, "", "", "", (results.totals.offcutWeight).toFixed(1) + RT(" kg", " كجم"), `${fmtTon(results.totals.offcutWeight)} ${RT("t", "طن")}`]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), RT("Reusable Offcuts", "القصاصات الصالحة"));
  // Machine-readable sheet so this exact file can be dropped back into the optimizer.
  const token = encodeLeftoverToken(plateLeftoverPayload(results));
  const tws = XLSX.utils.aoa_to_sheet([[RT("Steel Optimizer re-import data — keep this sheet.", "بيانات إعادة الاستيراد لـ Steel Optimizer — لا تحذف هذه الورقة.")], [RT("Drop this file into the “Reuse a previous Leftover file” button to load these leftovers again.", "أفلت هذا الملف في زر «إعادة استخدام ملف بواقٍ سابق» لتحميل هذه البواقي مرة أخرى.")], [token]]);
  tws["!cols"] = [{ wch: 120 }];
  XLSX.utils.book_append_sheet(wb, tws, "_reimport");
  saveWorkbook(wb, "SteelOptimizer_Offcuts.xlsx");
}

function offcutSVG(o) {
  // Draw the real leftover shape (L or rectangle) with the usable rectangle shaded.
  const W = 150, H = 130, pad = 18;
  const A = o.A, B = o.B, sc = Math.min((W - pad * 2) / A, (H - pad * 2) / B);
  const aw = A * sc, bh = B * sc, ox = (W - aw) / 2, oy = (H - bh) / 2;
  let body;
  if (o.shape === "L") {
    const nW = o.notchW * sc, nH = o.notchH * sc;
    const pts = [[ox + nW, oy], [ox + aw, oy], [ox + aw, oy + bh], [ox, oy + bh], [ox, oy + nH], [ox + nW, oy + nH]];
    body = `<polygon points="${pts.map(p => p.map(v => v.toFixed(1)).join(",")).join(" ")}" fill="#d8f3e6" stroke="#0a7a52" stroke-width="1.5"/>`;
    // usable rect: bottom strip or right strip
    const useBottom = o.w === A;
    const ux = ox, uy = useBottom ? oy + nH : oy, uw = useBottom ? aw : (aw - nW), uh = useBottom ? (bh - nH) : bh;
    body += `<rect x="${ux.toFixed(1)}" y="${uy.toFixed(1)}" width="${uw.toFixed(1)}" height="${uh.toFixed(1)}" fill="#10b981" fill-opacity="0.35" stroke="#10b981" stroke-width="1" stroke-dasharray="4 2"/>`;
  } else {
    body = `<rect x="${ox.toFixed(1)}" y="${oy.toFixed(1)}" width="${aw.toFixed(1)}" height="${bh.toFixed(1)}" fill="#10b981" fill-opacity="0.3" stroke="#0a7a52" stroke-width="1.5"/>`;
  }
  return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${body}<text x="${W / 2}" y="${H - 4}" font-size="9" text-anchor="middle" fill="#0a7a52" font-family="monospace">${A}×${B}${o.shape === "L" ? " L" : ""}</text></svg>`;
}

function exportOffcutPDF(results, material, reuseMin) {
  const rowsHtml = results.groups.flatMap(g => g.offcuts.map((o, k) => `<tr><td>${g.thickness} ${RT("mm", "مم")}</td><td>${RT("Sheet", "لوح")} ${o.sheet}</td><td>${o.shape === "L" ? RT("L-shape", "شكل L") : RT("Rectangle", "مستطيل")}</td><td>${o.shape === "L" ? `${o.A} x ${o.B} (${RT("notch", "ركن")} ${o.notchW} x ${o.notchH})` : `${o.A} x ${o.B}`}</td><td><b>${o.w} x ${o.h} ${RT("mm", "مم")}</b></td><td>${o.weight >= 1 ? o.weight.toFixed(1) + RT(" kg", " كجم") : (o.weight * 1000).toFixed(0) + RT(" g", " جم")}</td><td>${RT("keep for future jobs", "احتفظ بها لمشاريع قادمة")}</td></tr>`)).join("");
  const cards = results.groups.flatMap(g => g.offcuts.map(o => `<div class="card">${offcutSVG(o)}<div class="cap"><b>${g.thickness} ${RT("mm", "مم")} · ${RT("Sheet", "لوح")} ${o.sheet}</b><br/>${o.shape === "L" ? RT("L-shape", "شكل L") : RT("Rectangle", "مستطيل")} — ${RT("usable", "القابل للاستخدام")} <b>${o.w}×${o.h}</b><br/>${o.weight >= 1 ? o.weight.toFixed(1) + RT(" kg", " كجم") : (o.weight * 1000).toFixed(0) + RT(" g", " جم")}</div></div>`)).join("");
  const payload = plateLeftoverPayload(results);
  const token = encodeLeftoverToken(payload);
  const html = `<html${RT_HTML()}><head><title>${RT("Steel Optimizer — Reusable Offcuts", "Steel Optimizer — القصاصات الصالحة")}</title><style>
  body{font-family:Georgia,serif;padding:32px;color:#04140d}
  h1{color:#0a7a52;border-bottom:2px solid #bfe6d2;padding-bottom:10px}
  h2{color:#0a7a52;font-size:15px;margin-top:26px}
  table{width:100%;border-collapse:collapse;margin-top:14px;font-family:monospace;font-size:13px}
  th{background:#e8f8ef;color:#0a7a52;padding:10px;text-align:left;font-size:11px}
  td{padding:10px;border-bottom:1px solid #eee}
  tr.total td{background:#e8f8ef;font-weight:bold;border-top:2px solid #10b981}
  .banner{margin-top:14px;padding:14px 16px;background:#effaf3;border:2px solid #0a7a52;border-radius:8px;font-size:14px}
  .grid{display:flex;flex-wrap:wrap;gap:14px;margin-top:12px}
  .card{border:1px solid #cdeadd;border-radius:8px;padding:10px;text-align:center;background:#fbfffd;width:170px}
  .cap{font-family:monospace;font-size:10px;color:#234;margin-top:4px;line-height:1.5}
  .reimport{margin-top:18px;padding:10px 12px;background:#f8fafc;border:1px dashed #94a3b8;border-radius:6px;font-size:11px;color:#475569}
  .tok{font-family:monospace;font-size:7px;color:#9ca3af;word-break:break-all;line-height:1.3;margin-top:6px}
  .btnrow{margin-top:24px;display:flex;gap:12px;flex-wrap:wrap}
  button{padding:12px 26px;border-radius:6px;font-size:14px;cursor:pointer;font-family:inherit;line-height:1.2}
  .b-save{background:#0a7a52;color:#fff;border:2px solid #0a7a52}
  .b-print{background:#fff;color:#0a7a52;border:2px solid #0a7a52}
  .hint{margin-top:10px;font-size:11px;color:#64748b;line-height:1.7}
  @media print{button,.btnrow,.hint{display:none}}${RT_CSS()}
  </style></head><body>
  <h1>&#9851; ${RT("Reusable Offcut Register", "سجل القصاصات القابلة لإعادة الاستخدام")}</h1>
  <p>${RT("Material", "المعدن")}: <b>${material}</b> &middot; ${new Date().toLocaleDateString(soDateLoc("en-GB"))} &middot; ${RT("minimum kept side", "أقل ضلع يُحتفظ به")} &ge; ${reuseMin} ${RT("mm", "مم")}</p>
  <div class="banner">${RT(`${results.totals.offcutCount} reusable offcut${results.totals.offcutCount !== 1 ? "s" : ""} &asymp; <b>${fmtTon(results.totals.offcutWeight)} t</b> to keep for future jobs. Green dashed = the largest rectangle you can cut from each leftover.`, `${results.totals.offcutCount} قصاصة صالحة &asymp; <b>${fmtTon(results.totals.offcutWeight)} طن</b> للاحتفاظ بها لمشاريع قادمة. المتقطع الأخضر = أكبر مستطيل يمكن قصّه من كل قطعة متبقية.`, (L) => L === "ru" ? `Деловых остатков: ${results.totals.offcutCount} &asymp; <b>${fmtTon(results.totals.offcutWeight)} т</b> — сохраните для будущих заказов. Зелёный пунктир — наибольший прямоугольник, который можно вырезать из каждого остатка.` : L === "zh" ? `${results.totals.offcutCount} 块可再利用余料 &asymp; <b>${fmtTon(results.totals.offcutWeight)} t</b>，留待后续项目使用。绿色虚线 = 每块余料中可切出的最大矩形。` : `Sobrantes reutilizables: ${results.totals.offcutCount} &asymp; <b>${fmtTon(results.totals.offcutWeight)} t</b> para conservar en futuros trabajos. Discontinuo verde = el mayor rectángulo que se puede cortar de cada sobrante.`)}</div>
  <h2>${RT("Leftover shapes", "أشكال البواقي")}</h2>
  <div class="grid">${cards || `<p>${RT("No reusable offcuts above the minimum size.", "لا توجد قصاصات صالحة أكبر من الحد الأدنى.")}</p>`}</div>
  <h2>${RT("Register", "السجل")}</h2>
  <table><tr><th>${RT("Thickness", "السماكة")}</th><th>${RT("From Sheet", "من اللوح")}</th><th>${RT("Shape", "الشكل")}</th><th>${RT("Overall", "الأبعاد الكلية")}</th><th>${RT("Usable Rect", "المستطيل القابل للاستخدام")}</th><th>${RT("Weight", "الوزن")}</th><th>${RT("Status", "الحالة")}</th></tr>
  ${rowsHtml || `<tr><td colspan="7">${RT("No reusable offcuts above the minimum size.", "لا توجد قصاصات صالحة أكبر من الحد الأدنى.")}</td></tr>`}
  <tr class="total"><td colspan="5">${RT("TOTAL", "الإجمالي")} — ${results.totals.offcutCount} ${RT("pieces", "قطعة")}</td><td>${fmtTon(results.totals.offcutWeight)} ${RT("t", "طن")}</td><td></td></tr></table>
  <div class="reimport">&#9851; ${RT("Re-import code — drop this PDF back into the optimizer's <b>“Reuse a previous Leftover file”</b> button to load these leftovers again.", "رمز إعادة الاستيراد — أفلت ملف PDF هذا في زر <b>«إعادة استخدام ملف بواقٍ سابق»</b> لتحميل هذه البواقي مرة أخرى.")}<div class="tok">${token.replace(/</g, "&lt;").replace(/>/g, "&gt;")}</div></div>
  <script type="application/json" id="steelopt">${JSON.stringify(payload).replace(/</g, "\\u003c")}</script>
  <div class="btnrow">
  <button class="b-save" onclick="savePDF()">&#8681; ${soTb("MISC.tbSave", "Save as PDF", "حفظ PDF")}</button>
  <button class="b-print" onclick="window.print()">&#128424; ${soTb("MISC.tbPrint", "Print", "طباعة")}</button>
  </div>
  <div class="hint">${soTb("MISC.tbHint", "Both open your browser's print dialog. For a PDF, set <b>Destination</b> to &ldquo;Save as PDF&rdquo;.", "كلا الزرين يفتحان نافذة الطباعة. لحفظ ملف PDF اختر من قائمة <b>الوجهة</b> خيار الحفظ بصيغة PDF.")}</div>
  <script>function savePDF(){var t=document.title,d=new Date(),p=function(n){return(n<10?"0":"")+n};document.title="SteelOptimizer_Offcuts_"+d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate());var r=function(){document.title=t;window.removeEventListener("afterprint",r)};window.addEventListener("afterprint",r);window.print();setTimeout(r,4000);}<\/script></body></html>`;
  downloadHTML(html, "SteelOptimizer_Offcuts.html");
}

function sheetSVG(group, sObj, partColors) {
  const sw = group.sw, sh = group.sh, W = 150, H = Math.round(150 * sh / sw);
  const sc = W / sw;
  let body = `<rect x="0" y="0" width="${W}" height="${H}" fill="#0f1318" stroke="#b45309" stroke-width="1"/>`;
  (sObj.placements || []).forEach((p, i) => {
    const x = p.x * sc, y = p.y * sc, w = p.w * sc, h = p.h * sc, col = partColors[p.id] || "#3b82f6";
    body += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${h.toFixed(1)}" fill="${col}" fill-opacity="0.45" stroke="${col}" stroke-width="0.6"/>`;
    if (p.spliced) body += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${h.toFixed(1)}" fill="none" stroke="#f59e0b" stroke-width="0.8" stroke-dasharray="3 2"/>`;
  });
  (sObj.offcuts || []).forEach(o => {
    if (o.shape === "L") {
      const ox = o.x * sc, oy = o.y * sc, A = o.A * sc, B = o.B * sc, nW = o.notchW * sc, nH = o.notchH * sc;
      const pts = [[ox + nW, oy], [ox + A, oy], [ox + A, oy + B], [ox, oy + B], [ox, oy + nH], [ox + nW, oy + nH]];
      body += `<polygon points="${pts.map(p => p.map(v => v.toFixed(1)).join(",")).join(" ")}" fill="#10b981" fill-opacity="0.22" stroke="#10b981" stroke-width="0.8"/>`;
    } else {
      body += `<rect x="${(o.x * sc).toFixed(1)}" y="${(o.y * sc).toFixed(1)}" width="${(o.A * sc).toFixed(1)}" height="${(o.B * sc).toFixed(1)}" fill="#10b981" fill-opacity="0.22" stroke="#10b981" stroke-width="0.8"/>`;
    }
  });
  return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${body}</svg>`;
}

function exportPlatePDF(results, material, reuseMin, t, pricing) {
  let costBox = "";
  if (pricing && pricing.active) {
    const P = pricing.pricePerTon, M = v => `${Math.round(v).toLocaleString("en-US")} ${pricing.currency}`;
    costBox = `<div class="purchase" style="background:#eef6ff;border-color:#2563eb;font-size:15px">${RT("Total purchase cost:", "تكلفة الشراء الإجمالية:")} <b style="font-size:26px;color:#1d4ed8">${M(results.totals.sheetWeight / 1000 * P)}</b> &nbsp;&middot;&nbsp; ${RT("Net used:", "قيمة المستخدم:")} <b style="font-size:15px;color:#1a1206">${M(results.totals.partWeight / 1000 * P)}</b> &nbsp;&middot;&nbsp; <span style="color:#b91c1c">${RT("Scrap cost:", "تكلفة الهالك:")} <b>${M(results.totals.wasteWeight / 1000 * P)}</b></span> &nbsp;&middot;&nbsp; <span style="color:#0a7a52">${RT("Reusable offcut value:", "قيمة القصاصات الصالحة:")} <b>${M(results.totals.offcutWeight / 1000 * P)}</b></span><div style="font-size:11px;color:#777;margin-top:6px">${RT("Based on", "بناءً على")} ${Math.round(P).toLocaleString("en-US")} ${pricing.currency}/${RT("ton", "طن")} (${RT("your entered price", "السعر الذي أدخلته")})</div></div>`;
  }
  const offRows = results.groups.flatMap(g => g.offcuts.map(o => `<tr><td>${g.thickness} ${RT("mm", "مم")}</td><td>${RT("Sheet", "لوح")} ${o.sheet}</td><td>${o.shape === "L" ? RT("L-shape", "شكل L") : RT("Rect", "مستطيل")}</td><td>${o.shape === "L" ? `${o.A}×${o.B} (${RT("notch", "ركن")} ${o.notchW}×${o.notchH})` : `${o.A}×${o.B}`}</td><td><b>${o.w} × ${o.h} ${RT("mm", "مم")}</b></td><td>${o.weight.toFixed(1)} ${RT("kg", "كجم")}</td><td>${RT("keep for future jobs", "احتفظ بها لمشاريع قادمة")}</td></tr>`)).join("");
  const PCOL = ["#3b82f6", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6", "#ec4899", "#14b8a6", "#f97316", "#06b6d4", "#a855f7"];
  const layouts = results.groups.map(g => { let ci = 0; const cmap = {}; (g.sheets || []).forEach(s => (s.placements || []).forEach(p => { if (!(p.id in cmap)) cmap[p.id] = PCOL[ci++ % PCOL.length]; })); const tiles = (g.sheets || []).map((s, si) => `<div class="tile">${sheetSVG(g, s, cmap)}<div class="tcap">${RT("Sheet", "لوح")} ${si + 1} · ${g.thickness}${RT("mm", " مم", (L) => (L === "ru" ? " мм" : "mm"))}</div></div>`).join(""); return `<h2>${g.thickness} ${RT("mm", "مم")} — ${g.sheetCount} ${RT(`sheet${g.sheetCount > 1 ? "s" : ""}`, "لوح", (L) => L === "ru" ? `${soRuPlural(g.sheetCount, "лист", "листа", "листов")}` : L === "zh" ? `张` : `${g.sheetCount > 1 ? "chapas" : "chapa"}`)} (${g.sw}×${g.sh})</h2><div class="tiles">${tiles}</div>`; }).join("");
  const html = `<html${RT_HTML()}><head><title>${RT("Steel Optimizer — Plate Report", "Steel Optimizer — تقرير الألواح")}</title><style>
  body{font-family:Georgia,serif;padding:32px;color:#1a1206}
  h1{color:#b45309;border-bottom:2px solid #f0d9b0;padding-bottom:10px}
  h2{color:#0a7a52;font-size:15px;margin-top:24px;border-bottom:1px solid #e6dcc8;padding-bottom:6px}
  table{width:100%;border-collapse:collapse;margin-top:14px;font-family:monospace;font-size:13px}
  th{background:#fdf4e3;color:#b45309;padding:10px;text-align:left;font-size:11px}
  td{padding:10px;border-bottom:1px solid #eee}
  tr.total td{background:#fdf4e3;font-weight:bold;border-top:2px solid #f59e0b}
  table.off th{background:#e8f8ef;color:#0a7a52}
  .purchase{margin-top:20px;padding:18px;background:#fdf4e3;border:2px solid #f59e0b;border-radius:8px;font-size:18px}
  .purchase b{font-size:30px;color:#b45309}
  .buylist{margin:14px 0;padding:16px 18px;background:#fffaf0;border:1px solid #f0d9b0;border-radius:8px}
  .buylist h3{margin:0 0 10px;color:#b45309;font-size:13px;letter-spacing:1px;text-transform:uppercase}
  .buylist .row{font-size:16px;margin:6px 0}.buylist .n{font-weight:bold;color:#92400e;font-size:19px}
  .reuse{margin-top:12px;padding:14px 16px;background:#effaf3;border:2px solid #0a7a52;border-radius:8px;font-size:14px}
  .tiles{display:flex;flex-wrap:wrap;gap:10px;margin-top:10px}
  .tile{text-align:center}.tcap{font-family:monospace;font-size:9px;color:#555;margin-top:3px}
  .legend{font-size:11px;color:#555;margin-top:6px}.legend b{color:#0a7a52}
  .btnrow{margin-top:24px;display:flex;gap:12px;flex-wrap:wrap}
  button{padding:12px 26px;border-radius:6px;font-size:14px;cursor:pointer;font-family:inherit;line-height:1.2}
  .b-save{background:#b45309;color:#fff;border:2px solid #b45309}
  .b-print{background:#fff;color:#b45309;border:2px solid #b45309}
  .hint{margin-top:10px;font-size:11px;color:#64748b;line-height:1.7}
  @media print{button,.btnrow,.hint{display:none}}${RT_CSS()}
  </style></head><body>
  <h1>${RT("Steel Optimizer — Plate Procurement Summary", "Steel Optimizer — ملخص توريد الألواح")}</h1>
  <p>${RT("Material", "المعدن")}: <b>${material}</b> &nbsp;&middot;&nbsp; ${new Date().toLocaleDateString(soDateLoc("en-GB"))}</p>
  <div class="buylist"><h3>${RT("What you need to buy", "ما تحتاج شراءه")}</h3>
  ${results.groups.map(g => RT(`<div class="row">&#10003; You need <span class="n">${g.sheetCount}</span> sheet${g.sheetCount > 1 ? "s" : ""} of <b>${g.thickness} mm</b> (${g.sw}x${g.sh} mm, ${material})</div>`, `<div class="row">&#10003; تحتاج <span class="n">${g.sheetCount}</span> لوح بسماكة <b>${g.thickness} مم</b> (${g.sw}x${g.sh} مم، ${material})</div>`, (L) => L === "ru" ? `<div class="row">&#10003; Требуется листов: <span class="n">${g.sheetCount}</span>, толщина <b>${g.thickness} мм</b> (${g.sw}x${g.sh} мм, ${material})</div>` : L === "zh" ? `<div class="row">&#10003; 需要 <span class="n">${g.sheetCount}</span> 张 <b>${g.thickness} mm</b> 钢板（${g.sw}x${g.sh} mm，${material}）</div>` : `<div class="row">&#10003; Chapas necesarias: <span class="n">${g.sheetCount}</span> de <b>${g.thickness} mm</b> (${g.sw}x${g.sh} mm, ${material})</div>`)).join("")}</div>
  <table><tr><th>${RT("Thickness", "السماكة")}</th><th>${RT("Sheet Size", "مقاس اللوح")}</th><th>${RT("Sheets", "الألواح")}</th><th>${RT("Total Wt", "الوزن الإجمالي")}</th><th>${RT("Parts", "القطع")}</th><th>${RT("Utilization", "الاستغلال")}</th><th>${RT("Net Wt", "الوزن الصافي")}</th><th>${RT("Scrap Wt", "وزن السكراب")}</th><th>${RT("Paint (1 face)", "الدهان (وجه واحد)")}</th></tr>
  ${results.groups.map(g => `<tr><td>${g.thickness} ${RT("mm", "مم")}</td><td>${g.sw}x${g.sh}</td><td>${g.sheetCount}</td><td>${fmtTon(g.sheetWeight)} ${RT("t", "طن")}</td><td>${g.partCount}</td><td>${g.utilPct}%</td><td>${fmtTon(g.partWeight)} ${RT("t", "طن")}</td><td>${fmtTon(g.wasteWeight)} ${RT("t", "طن")}</td><td>${g.paintM2 != null ? fmtM2(g.paintM2) + RT(" m&sup2;", " م&sup2;") : "&mdash;"}</td></tr>`).join("")}
  <tr class="total"><td>${RT("TOTAL", "الإجمالي")}</td><td>${results.groups.length} ${RT("thk", "سماكة")}</td><td>${results.totals.sheets}</td><td>${fmtTon(results.totals.sheetWeight)} ${RT("t", "طن")}</td><td>${results.totals.parts}</td><td>${results.totals.utilPct}%</td><td>${fmtTon(results.totals.partWeight)} ${RT("t", "طن")}</td><td>${fmtTon(results.totals.wasteWeight)} ${RT("t", "طن")}</td><td>${results.totals.paintM2 != null ? fmtM2(results.totals.paintM2) + RT(" m&sup2;", " م&sup2;") : "&mdash;"}</td></tr></table>
  <p style="font-size:12px;color:#777">${RT("Total Wt = full sheets purchased &middot; Net Wt = steel used in project &middot; Scrap Wt = offcut (Total - Net) &middot; Paint = one face of your parts (L &times; W &times; qty), offcuts and scrap excluded", "الوزن الإجمالي = الألواح الكاملة المشتراة &middot; الوزن الصافي = الحديد المستخدم في المشروع &middot; وزن السكراب = القصاصات (الإجمالي − الصافي) &middot; الدهان = وجه واحد لقطعك (الطول &times; العرض &times; العدد) بدون البواقي والهالك")}</p>
  <div class="purchase">${RT("Total Purchase Weight (gross):", "إجمالي وزن الشراء:")} <b>${fmtTon(results.totals.sheetWeight)} ${RT("Ton", "طن")}</b> &nbsp;(${results.totals.sheets} ${RT("sheets", "لوح")}) &middot; ${RT("Net:", "الصافي:")} ${fmtTon(results.totals.partWeight)} ${RT("t", "طن")} &middot; ${RT("Scrap:", "السكراب:")} ${fmtTon(results.totals.wasteWeight)} ${RT("t", "طن")}${results.totals.paintM2 != null ? `<div style="font-size:15px;margin-top:8px">${RT("Paint area (one face):", "مساحة الدهان (وجه واحد):")} <b style="font-size:20px;color:#0369a1">${fmtM2(results.totals.paintM2)} ${RT("m&sup2;", "م&sup2;")}</b></div>` : ""}</div>
  ${costBox}
  <h2>${RT("Nesting layouts", "مخططات التوزيع")}</h2>
  <div class="legend">${RT("Amber dashed = welded/spliced piece &middot; <b>Green = reusable leftover</b> (L-shape or rectangle).", "المتقطع البرتقالي = قطعة ملحومة/موصولة &middot; <b>الأخضر = بواقٍ قابلة لإعادة الاستخدام</b> (شكل L أو مستطيل).")}</div>
  ${layouts}
  <h2>&#9851; ${RT(`Reusable Offcut Register (each side >= ${reuseMin} mm)`, `سجل القصاصات القابلة لإعادة الاستخدام (كل ضلع ≥ ${reuseMin} مم)`, (L) => L === "ru" ? `Реестр деловых остатков (каждая сторона >= ${reuseMin} мм)` : L === "zh" ? `可再利用余料清单（每边 >= ${reuseMin} mm）` : `Registro de sobrantes reutilizables (cada lado >= ${reuseMin} mm)`)}</h2>
  ${results.totals.offcutCount > 0 ? `<div class="reuse">${RT(`${results.totals.offcutCount} reusable offcut${results.totals.offcutCount > 1 ? "s" : ""} &asymp; ${fmtTon(results.totals.offcutWeight)} t to keep for future jobs.`, `${results.totals.offcutCount} قصاصة صالحة &asymp; ${fmtTon(results.totals.offcutWeight)} طن للاحتفاظ بها لمشاريع قادمة.`, (L) => L === "ru" ? `Деловых остатков: ${results.totals.offcutCount} &asymp; ${fmtTon(results.totals.offcutWeight)} т — сохраните для будущих заказов.` : L === "zh" ? `${results.totals.offcutCount} 块可再利用余料 &asymp; ${fmtTon(results.totals.offcutWeight)} t，留待后续项目使用。` : `Sobrantes reutilizables: ${results.totals.offcutCount} &asymp; ${fmtTon(results.totals.offcutWeight)} t para conservar en futuros trabajos.`)}</div><table class="off"><tr><th>${RT("Thickness", "السماكة")}</th><th>${RT("From Sheet", "من اللوح")}</th><th>${RT("Shape", "الشكل")}</th><th>${RT("Overall", "الأبعاد الكلية")}</th><th>${RT("Usable Rect", "المستطيل القابل للاستخدام")}</th><th>${RT("Weight", "الوزن")}</th><th>${RT("Status", "الحالة")}</th></tr>${offRows}</table>` : `<p style="color:#777">${RT("No reusable offcuts — all leftovers below the minimum size.", "لا توجد قصاصات صالحة — كل البواقي أصغر من الحد الأدنى.")}</p>`}
  <div class="btnrow">
  <button class="b-save" onclick="savePDF()">&#8681; ${soTb("MISC.tbSave", "Save as PDF", "حفظ PDF")}</button>
  <button class="b-print" onclick="window.print()">&#128424; ${soTb("MISC.tbPrint", "Print", "طباعة")}</button>
  </div>
  <div class="hint">${soTb("MISC.tbHint", "Both open your browser's print dialog. For a PDF, set <b>Destination</b> to &ldquo;Save as PDF&rdquo;.", "كلا الزرين يفتحان نافذة الطباعة. لحفظ ملف PDF اختر من قائمة <b>الوجهة</b> خيار الحفظ بصيغة PDF.")}</div>
  <script>function savePDF(){var t=document.title,d=new Date(),p=function(n){return(n<10?"0":"")+n};document.title="SteelOptimizer_Plates_"+d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate());var r=function(){document.title=t;window.removeEventListener("afterprint",r)};window.addEventListener("afterprint",r);window.print();setTimeout(r,4000);}<\/script></body></html>`;
  downloadHTML(html, "SteelOptimizer_Plates.html");
}

function exportSectionExcel(results, pricing) {
  const wb = XLSX.utils.book_new();
  const today = new Date().toLocaleDateString(soDateLoc());
  const totalBars = results.groups.reduce((s, g) => s + g.summary.stockCount, 0);
  const netKg = results.groups.reduce((s, g) => s + (g.netKg || 0), 0);
  const wasteKg = results.groups.reduce((s, g) => s + (g.kgm ? (g.summary.totalWaste / 1000) * g.kgm : 0), 0);
  const wsAdd = (name, aoa, cols) => { const ws = XLSX.utils.aoa_to_sheet(aoa); if (cols) ws["!cols"] = cols; XLSX.utils.book_append_sheet(wb, ws, name); };

  // ── 1) SUMMARY ──
  const sum = [
    [RT("STEEL OPTIMIZER — SECTION OPTIMIZATION REPORT", "STEEL OPTIMIZER — تقرير تحسين قص المقاطع")],
    [RT(`Generated: ${today}`, `تاريخ الإنشاء: ${today}`, (L) => L === "ru" ? `Сформирован: ${today}` : L === "zh" ? `生成日期：${today}` : `Generado: ${today}`)],
    [],
    [RT("KEY FIGURES", "الأرقام الرئيسية"), ""],
    [RT("Total Weight to Purchase (t)", "الوزن الإجمالي للشراء (طن)"), (results.grandKg / 1000).toFixed(3)],
    [RT("Net Weight used in Project (t)", "الوزن الصافي المستخدم في المشروع (طن)"), (netKg / 1000).toFixed(3)],
    [RT("Scrap / Offcut Weight (t)", "وزن السكراب / القصاصات (طن)"), (wasteKg / 1000).toFixed(3)],
    [RT("Total Bars to Buy", "إجمالي الأعواد للشراء"), totalBars],
    [RT("Profile Groups", "عدد المقاطع"), results.groups.length],
    [RT("Paint Area (m²) — net cut lengths, outside faces, square corners (conservative)", "مساحة الدهان (م²) — أطوال القص الصافية، الأوجه الخارجية، زوايا حادّة (متحفّظة)"), results.paintM2 == null || results.paintMissing.length === results.groups.length ? "—" : upM2(results.paintM2, 1)],
    [],
    [RT("PER-PROFILE BREAKDOWN", "التفصيل حسب المقطع")],
    [RT("Profile", "المقطع"), RT("Grade", "الرتبة"), RT("kg/m", "كجم/م"), RT("Bars to Buy", "أعواد للشراء"), RT("Total Length (m)", "الطول الإجمالي (م)"), RT("Net Length (m)", "الطول الصافي (م)"), RT("Offcut (m)", "القصاصة (م)"), RT("Utilization %", "نسبة الاستغلال %"), RT("Total Weight (t)", "الوزن الإجمالي (طن)"), RT("Paint m²/m", "الدهان م²/م"), RT("Paint Area (m²)", "مساحة الدهان (م²)")],
  ];
  if (pricing && pricing.active) {
    const P = pricing.pricePerTon, M = v => `${Math.round(v).toLocaleString("en-US")} ${pricing.currency}`;
    const offKg = sectionLeftovers(results).reduce((s, r) => s + (r.kg || 0), 0);
    sum.splice(10, 0,
      [RT("STEEL PRICE (per ton)", "سعر الحديد (للطن)"), `${Math.round(P).toLocaleString("en-US")} ${pricing.currency} — ${RT("your entered price", "السعر الذي أدخلته")}`],
      [RT("TOTAL PURCHASE COST - all raw material incl. waste", "تكلفة الشراء الإجمالية — كل المواد الخام شاملة الهدر"), M(results.grandKg / 1000 * P)],
      [RT("NET USED VALUE", "قيمة الحديد المستخدم"), M(Math.max(results.grandKg - wasteKg, 0) / 1000 * P)],
      [RT("SCRAP COST - money in waste", "تكلفة الهالك — المال في الهدر"), M(wasteKg / 1000 * P)],
      [RT("REUSABLE OFFCUT VALUE - recoverable", "قيمة القصاصات الصالحة — قابلة للاسترداد"), M(offKg / 1000 * P)]);
  }
  results.groups.forEach(g => sum.push([g.profile, soGrade(g.grade) || "—", g.kgm || "—", g.summary.stockCount, (g.summary.totalStock / 1000).toFixed(2), (g.summary.totalNet / 1000).toFixed(2), (g.summary.totalWaste / 1000).toFixed(2), g.summary.utilPct, g.stockKg != null ? (g.stockKg / 1000).toFixed(3) : "—", g.paintM2pm != null ? upM2(g.paintM2pm, 3) : "—", g.paintM2 != null ? upM2(g.paintM2, 1) : "—"]));
  sum.push([], [RT("Total Weight = full bars purchased · Net Weight = steel used in project · Offcut = Total − Net", "الوزن الإجمالي = الأعواد الكاملة المشتراة · الوزن الصافي = الحديد المستخدم في المشروع · القصاصة = الإجمالي − الصافي")]);
  if (results.paintM2 != null) sum.push([RT("Paint Area = net cut lengths from your list × outside surface per metre (square corners, root radii ignored — typically 2–5% above catalogue). Offcuts and scrap not included.", "مساحة الدهان = أطوال القص الصافية من قائمتك × السطح الخارجي لكل متر (زوايا حادّة بدون أنصاف أقطار التدوير — عادةً أعلى من الكتالوج بـ 2–5٪). البواقي والهالك غير محسوبة.")]);
  if (results.paintMissing && results.paintMissing.length) sum.push([RT(`Paint area not included — no dimensions in the library: ${results.paintMissing.join(", ")}`, `مساحة الدهان غير محسوبة — لا توجد أبعاد في المكتبة: ${results.paintMissing.join("، ")}`, (L) => L === "ru" ? `Площадь окраски не включена — в сортаменте нет размеров: ${results.paintMissing.join(", ")}` : L === "zh" ? `未计入涂装面积 — 截面库中无尺寸：${results.paintMissing.join(", ")}` : `Superficie de pintura no incluida — sin dimensiones en la biblioteca: ${results.paintMissing.join(", ")}`)]);
  wsAdd(RT("Summary", "الملخص"), sum, [{ wch: 34 }, { wch: 12 }, { wch: 8 }, { wch: 12 }, { wch: 16 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 16 }, { wch: 11 }, { wch: 15 }]);

  // ── 2) PROCUREMENT (what to order) ──
  const proc = [[RT("WHAT TO ORDER", "ما يجب طلبه")], [], [RT("Profile", "المقطع"), RT("Grade", "الرتبة"), RT("Market Length (mm)", "الطول بالسوق (مم)"), RT("Qty to Order", "العدد المطلوب"), RT("kg/m", "كجم/م"), RT("Weight (t)", "الوزن (طن)")]];
  results.groups.forEach(g => { const byLen = g.bins.reduce((a, b) => { a[b.stockLength] = (a[b.stockLength] || 0) + 1; return a; }, {}); Object.entries(byLen).forEach(([len, qty]) => { const kg = g.kgm ? (parseInt(len) / 1000) * g.kgm * qty : 0; proc.push([g.profile, soGrade(g.grade) || "", parseInt(len), qty, g.kgm || "", (kg / 1000).toFixed(3)]); }); });
  proc.push([], [RT("TOTAL STEEL TO PROCURE (t)", "إجمالي الحديد المطلوب توريده (طن)"), "", "", "", "", (results.grandKg / 1000).toFixed(3)]);
  wsAdd(RT("Procurement", "التوريد"), proc, [{ wch: 18 }, { wch: 12 }, { wch: 18 }, { wch: 14 }, { wch: 10 }, { wch: 12 }]);

  // ── 3) CUTTING PLAN (bar by bar, with a to-scale text cut map) ──
  const cutMap = (b, width = 40) => {
    let s = "";
    b.cuts.forEach((c, idx) => { const n = Math.max(1, Math.round((c.length / b.stockLength) * width)); s += (idx ? "|" : "") + "\u2588".repeat(n); });
    const ln = Math.round((b.remaining / b.stockLength) * width);
    if (ln > 0) s += (b.remaining >= 1000 ? "\u2592".repeat(ln) : "\u00b7".repeat(ln));
    return s;
  };
  const plan = [[RT("CUTTING PLAN — bar by bar", "خطة القص — عودًا بعود")], [], [RT("Cut Map key:  \u2588 = cut piece   |  = saw cut   \u2592 = reusable leftover (\u22651 m)   \u00b7 = scrap", "مفتاح الخريطة:  \u2588 = قطعة مقصوصة   |  = قطع المنشار   \u2592 = بواقٍ صالحة (\u22651 م)   \u00b7 = هالك")], [], [RT("Profile", "المقطع"), RT("Grade", "الرتبة"), RT("Bar #", "رقم العود"), RT("Available Length (mm)", "الطول المتوفر (مم)"), RT("Cuts (mm)", "القطع (مم)"), RT("Pieces", "عدد القطع"), RT("Leftover (mm)", "المتبقي (مم)"), RT("Leftover Reusable?", "صالح لإعادة الاستخدام؟"), RT("Utilization %", "نسبة الاستغلال %"), RT("Cut Map (to scale)", "خريطة القص (بمقياس)")]];
  results.groups.forEach(g => g.bins.forEach((b, i) => plan.push([g.profile, soGrade(g.grade) || "", i + 1, b.stockLength, b.cuts.map(c => Math.round(c.length)).join(" + "), b.cuts.length, Math.round(b.remaining), b.remaining >= 1000 ? RT("YES (≥1 m)", "نعم (≥1 م)") : RT("no", "لا"), Math.round(((b.stockLength - b.remaining) / b.stockLength) * 1000) / 10, cutMap(b)])));
  wsAdd(RT("Cutting Plan", "خطة القص"), plan, [{ wch: 18 }, { wch: 12 }, { wch: 7 }, { wch: 20 }, { wch: 40 }, { wch: 8 }, { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 46 }]);

  // ── 4) REUSABLE LEFTOVERS (the offcut register) ──
  const reuse = [[RT("REUSABLE BAR OFFCUTS — leftover ends ≥ 1 m, worth keeping for the next job", "قصاصات الأعواد الصالحة — نهايات متبقية ≥ 1 م تستحق الاحتفاظ بها للمشروع القادم")], [], [RT("Profile", "المقطع"), RT("Grade", "الرتبة"), RT("From Bar #", "من العود رقم"), RT("Available Length (mm)", "الطول المتوفر (مم)"), RT("Reusable Leftover (mm)", "البواقي الصالحة (مم)"), RT("Weight (kg)", "الوزن (كجم)")]];
  let rk = 0, reuseKg = 0; results.groups.forEach(g => g.bins.forEach((b, i) => { if (b.remaining >= 1000) { rk++; const kg = g.kgm ? (b.remaining / 1000) * g.kgm : 0; reuseKg += kg; reuse.push([g.profile, soGrade(g.grade) || "", i + 1, b.stockLength, Math.round(b.remaining), kg ? kg.toFixed(1) : ""]); } }));
  if (rk === 0) reuse.push([RT("No reusable bar offcuts ≥ 1 m — material was used efficiently.", "لا توجد قصاصات أعواد صالحة ≥ 1 م — تم استغلال المادة بكفاءة.")]);
  else reuse.push([], [RT("TOTAL REUSABLE", "إجمالي الصالح"), "", "", "", `${rk} ${RT("pieces", "قطعة")}`, reuseKg.toFixed(1)]);
  wsAdd(RT("Reusable Offcuts", "القصاصات الصالحة"), reuse, [{ wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 20 }, { wch: 22 }, { wch: 12 }]);

  // ── 5) SPLICED LONG MEMBERS (if any) ──
  const hasSplice = results.groups.some(g => g.splices && g.splices.length);
  if (hasSplice) {
    const sp = [[RT("LONG MEMBERS — SPLICED FROM SHORTER BARS", "عناصر طويلة — موصولة من أعواد أقصر")], [], [RT("Profile", "المقطع"), RT("Grade", "الرتبة"), RT("Member Length (mm)", "طول العنصر (مم)"), RT("Qty", "العدد"), RT("Bars per Member", "أعواد لكل عنصر"), RT("From Bar (mm)", "من عود (مم)")]];
    results.groups.forEach(g => (g.splices || []).forEach(s => sp.push([g.profile, soGrade(g.grade) || "", Math.round(s.length), s.qty, s.bars, Math.round(s.stock)])));
    wsAdd(RT("Spliced Members", "العناصر الموصولة"), sp, [{ wch: 18 }, { wch: 12 }, { wch: 18 }, { wch: 8 }, { wch: 16 }, { wch: 16 }]);
  }
  saveWorkbook(wb, "SteelOptimizer_Sections.xlsx");
}

/* Proportional visual cut bar for the PDF (light theme to match the cream report).
   Scale is shared per group (pxPerMm) so every bar is visually comparable. */
function barCutSVG(bin, pxPerMm) {
  const H = 30, W = Math.max(40, bin.stockLength * pxPerMm);
  const COL = ["#2563eb", "#0891b2", "#7c3aed", "#db2777", "#ea580c", "#0d9488", "#4f46e5", "#c026d3", "#0284c7", "#9333ea"];
  let x = 0, ci = 0, body = "";
  bin.cuts.forEach(c => {
    const w = c.length * pxPerMm, col = c.spliced ? "#b45309" : COL[ci++ % COL.length];
    body += `<rect x="${x.toFixed(1)}" y="0" width="${Math.max(0, w - 0.6).toFixed(1)}" height="${H}" fill="${col}" fill-opacity="0.88"/>`;
    if (w > 30) body += `<text x="${(x + w / 2).toFixed(1)}" y="${H / 2 + 3.5}" text-anchor="middle" font-size="9.5" fill="#fff" font-family="monospace">${Math.round(c.length)}</text>`;
    x += w;
  });
  const leftW = bin.remaining * pxPerMm;
  if (leftW > 0.5) {
    const reuse = bin.remaining >= 1000;
    body += `<rect x="${x.toFixed(1)}" y="0" width="${leftW.toFixed(1)}" height="${H}" fill="${reuse ? "#16a34a" : "#cbd5e1"}" fill-opacity="${reuse ? 0.35 : 0.55}" stroke="${reuse ? "#16a34a" : "#94a3b8"}" stroke-width="0.8" stroke-dasharray="3 2"/>`;
    if (leftW > 36) body += `<text x="${(x + leftW / 2).toFixed(1)}" y="${H / 2 + 3.5}" text-anchor="middle" font-size="9" fill="${reuse ? "#15803d" : "#64748b"}" font-family="monospace">${Math.round(bin.remaining)}${reuse ? RT(" keep", " احتفظ") : ""}</text>`;
  }
  return `<svg width="${W.toFixed(0)}" height="${H}" viewBox="0 0 ${W.toFixed(1)} ${H}" style="background:#faf7f0;border:1px solid #e6dcc8;border-radius:3px;vertical-align:middle">${body}</svg>`;
}

function exportSectionPDF(results, kerf, ow, pricing) {
  let costBox = "";
  if (pricing && pricing.active) {
    const P = pricing.pricePerTon, M = v => `${Math.round(v).toLocaleString("en-US")} ${pricing.currency}`;
    const offKg = sectionLeftovers(results).reduce((s, r) => s + (r.kg || 0), 0);
    costBox = `<div class="purchase" style="background:#eef6ff;border-color:#2563eb;font-size:15px">${RT("Total purchase cost:", "تكلفة الشراء الإجمالية:")} <b style="font-size:26px;color:#1d4ed8">${M(results.grandKg / 1000 * P)}</b> &nbsp;&middot;&nbsp; ${RT("Net used:", "قيمة المستخدم:")} <b style="font-size:15px;color:#1a1206">${M(Math.max(results.grandKg - ow.wasteKg, 0) / 1000 * P)}</b> &nbsp;&middot;&nbsp; <span style="color:#b91c1c">${RT("Scrap cost:", "تكلفة الهالك:")} <b>${M(ow.wasteKg / 1000 * P)}</b></span> &nbsp;&middot;&nbsp; <span style="color:#0a7a52">${RT("Reusable offcut value:", "قيمة القصاصات الصالحة:")} <b>${M(offKg / 1000 * P)}</b></span><div style="font-size:11px;color:#777;margin-top:6px">${RT("Based on", "بناءً على")} ${Math.round(P).toLocaleString("en-US")} ${pricing.currency}/${RT("ton", "طن")} (${RT("your entered price", "السعر الذي أدخلته")}) &middot; ${RT("purchase cost covers all raw material bought, including the waste portion.", "تكلفة الشراء تشمل كل المواد الخام المشتراة بما فيها جزء الهدر.")}</div></div>`;
  }
  const procRows = results.groups.map(g => { const byLen = g.bins.reduce((a, b) => { a[b.stockLength] = (a[b.stockLength] || 0) + 1; return a; }, {}); return Object.entries(byLen).map(([len, qty]) => { const kg = g.kgm ? (parseInt(len) / 1000) * g.kgm * qty : null; return `<tr><td>${g.profile}</td><td>${soGrade(g.grade) || "—"}</td><td>${fmtMm(parseInt(len))}</td><td>${qty}</td><td>${g.kgm || "—"}</td><td>${kg ? fmtKg(kg) : "—"}</td></tr>`; }).join(""); }).join("");
  const planRows = results.groups.map(g => g.bins.map((b, i) => `<tr><td>${g.profile}</td><td>#${i + 1}</td><td>${fmtMm(b.stockLength)}</td><td>${b.cuts.map(c => fmtMm(c.length)).join(" + ")}</td><td>${fmtMm(b.remaining)}</td><td>${(((b.stockLength - b.remaining) / b.stockLength) * 100).toFixed(1)}%</td></tr>`).join("")).join("");
  // Visual cut plan — one proportional bar diagram per stock bar, scaled per group.
  const visual = results.groups.map(g => {
    const longest = Math.max(...g.bins.map(b => b.stockLength), 1);
    const pxPerMm = 640 / longest; // longest bar in the group spans ~640px
    const bars = g.bins.map((b, i) => {
      const util = (((b.stockLength - b.remaining) / b.stockLength) * 100).toFixed(0);
      return `<div class="vbar"><div class="vlab">${RT("Bar", "عود")} ${i + 1} · ${fmtMm(b.stockLength)} · ${util}% ${RT("used", "مستخدم")}</div>${barCutSVG(b, pxPerMm)}</div>`;
    }).join("");
    return `<h3 class="vgrp">${g.profile}${g.grade ? ` · ${soGrade(g.grade)}` : ""} — ${g.bins.length} ${RT(`bar${g.bins.length > 1 ? "s" : ""}`, "عود", (L) => L === "ru" ? `${soRuPlural(g.bins.length, "хлыст", "хлыста", "хлыстов")}` : L === "zh" ? `根` : `${g.bins.length > 1 ? "barras" : "barra"}`)}</h3>${bars}`;
  }).join("");
  const html = `<html${RT_HTML()}><head><title>${RT("Steel Optimizer — Section Report", "Steel Optimizer — تقرير المقاطع")}</title><style>
  body{font-family:Georgia,serif;padding:32px;color:#1a1206}
  h1{color:#b45309;border-bottom:2px solid #f0d9b0;padding-bottom:10px}
  h2{color:#92400e;font-size:15px;margin-top:24px;border-bottom:1px solid #e6dcc8;padding-bottom:6px}
  table{width:100%;border-collapse:collapse;margin-top:12px;font-family:monospace;font-size:13px}
  th{background:#fdf4e3;color:#b45309;padding:9px;text-align:left;font-size:11px}
  td{padding:9px;border-bottom:1px solid #eee}
  .purchase{margin-top:18px;padding:18px;background:#fdf4e3;border:2px solid #f59e0b;border-radius:8px;font-size:18px}
  .purchase b{font-size:30px;color:#b45309}
  .vgrp{color:#92400e;font-size:13px;margin:18px 0 8px}
  .vbar{margin:7px 0}
  .vlab{font-family:monospace;font-size:10px;color:#555;margin-bottom:2px}
  .vleg{font-size:11px;color:#555;margin:6px 0 4px}.vleg b{color:#15803d}
  .btnrow{margin-top:24px;display:flex;gap:12px;flex-wrap:wrap}
  button{padding:12px 26px;border-radius:6px;font-size:14px;cursor:pointer;font-family:inherit;line-height:1.2}
  .b-save{background:#b45309;color:#fff;border:2px solid #b45309}
  .b-print{background:#fff;color:#b45309;border:2px solid #b45309}
  .hint{margin-top:10px;font-size:11px;color:#64748b;line-height:1.7}
  @media print{button,.btnrow,.hint{display:none}}${RT_CSS()}
  </style></head><body>
  <h1>${RT("Steel Optimizer — Section Procurement & Cutting", "Steel Optimizer — توريد وقص المقاطع")}</h1>
  <p>${new Date().toLocaleDateString(soDateLoc("en-GB"))} &nbsp;&middot;&nbsp; ${RT("kerf", "عرض القص")} ${kerf}${RT("mm", " مم", (L) => (L === "ru" ? " мм" : "mm"))}</p>
  <div class="purchase">${RT("Total Steel to Procure:", "إجمالي الحديد المطلوب:")} <b>${fmtKg(results.grandKg)}</b> &nbsp;&middot; ${ow.totalStock} ${RT("bars", "عود")} &middot; ${RT("waste", "الهدر")} ${ow.wastePct}% (${fmtKg(ow.wasteKg)})${results.paintM2 == null || results.paintMissing.length === results.groups.length ? "" : `<div style="font-size:15px;margin-top:8px">${RT("Paint area:", "مساحة الدهان:")} <b style="font-size:20px;color:#0369a1">${fmtM2(results.paintM2 || 0)} ${RT("m&sup2;", "م&sup2;")}</b></div>`}</div>
  ${costBox}
  <h2>${RT("Procurement — What to Order", "التوريد — ما يجب طلبه")}</h2>
  <table><tr><th>${RT("Profile", "المقطع")}</th><th>${RT("Grade", "الرتبة")}</th><th>${RT("Market Length", "الطول بالسوق")}</th><th>${RT("Qty", "العدد")}</th><th>${RT("kg/m", "كجم/م")}</th><th>${RT("Total Weight", "الوزن الإجمالي")}</th></tr>${procRows}</table>
  ${results.paintM2 == null ? "" : `<h2>${RT("Paint Area", "مساحة الدهان")}</h2>
  <table><tr><th>${RT("Profile", "المقطع")}</th><th>${RT("m&sup2;/m", "م&sup2;/م")}</th><th>${RT("Net cut length", "طول القص الصافي")}</th><th>${RT("Paint area", "مساحة الدهان")}</th></tr>${results.groups.map(g => `<tr><td>${g.profile}</td><td>${g.paintM2pm != null ? fmtM2pm(g.paintM2pm) : "&mdash;"}</td><td>${g.inputMm != null ? (g.inputMm / 1000).toFixed(2) + RT(" m", " م") : "&mdash;"}</td><td>${g.paintM2 != null ? fmtM2(g.paintM2) + RT(" m&sup2;", " م&sup2;") : "&mdash;"}</td></tr>`).join("")}<tr><td><b>${RT("TOTAL", "الإجمالي")}</b></td><td></td><td></td><td><b>${fmtM2(results.paintM2 || 0)} ${RT("m&sup2;", "م&sup2;")}</b></td></tr></table>
  <p style="font-size:12px;color:#777">${RT("Net cut lengths from your list &times; outside surface per metre &middot; square corners, root radii ignored (typically 2&ndash;5% above catalogue) &middot; offcuts and scrap not included", "أطوال القص الصافية من قائمتك &times; السطح الخارجي لكل متر &middot; زوايا حادّة بدون أنصاف أقطار التدوير (عادةً أعلى من الكتالوج بـ 2&ndash;5٪) &middot; البواقي والهالك غير محسوبة")}${results.paintMissing && results.paintMissing.length ? ` &middot; <b>${RT(`not included (no dimensions in the library): ${results.paintMissing.join(", ")}`, `غير محسوبة (لا توجد أبعاد في المكتبة): ${results.paintMissing.join("، ")}`, (L) => L === "ru" ? `не включено (в сортаменте нет размеров): ${results.paintMissing.join(", ")}` : L === "zh" ? `未计入（截面库中无尺寸）：${results.paintMissing.join(", ")}` : `no incluido (sin dimensiones en la biblioteca): ${results.paintMissing.join(", ")}`)}</b>` : ""}</p>`}
  <h2>${RT("Visual Cut Plan", "خطة القص المرئية")}</h2>
  <div class="vleg">${RT("Each bar drawn to scale &middot; coloured blocks = cut pieces (length in mm) &middot; <b>green dashed = reusable leftover &ge; 1 m</b> &middot; grey dashed = scrap &middot; amber = spliced run.", "كل عود مرسوم بمقياس &middot; الكتل الملوّنة = القطع (الطول بالمم) &middot; <b>المتقطع الأخضر = بواقٍ صالحة &ge; 1 م</b> &middot; المتقطع الرمادي = هالك &middot; البرتقالي = طول موصول.")}</div>
  ${visual}
  <h2>${RT("Cutting Plan — Bar by Bar", "خطة القص — عودًا بعود")}</h2>
  <table><tr><th>${RT("Profile", "المقطع")}</th><th>${RT("Bar #", "رقم العود")}</th><th>${RT("Stock Length", "الطول القياسي")}</th><th>${RT("Cuts", "القطع")}</th><th>${RT("Offcut", "القصاصة")}</th><th>${RT("Util", "الاستغلال")}</th></tr>${planRows}</table>
  <div class="btnrow">
  <button class="b-save" onclick="savePDF()">&#8681; ${soTb("MISC.tbSave", "Save as PDF", "حفظ PDF")}</button>
  <button class="b-print" onclick="window.print()">&#128424; ${soTb("MISC.tbPrint", "Print", "طباعة")}</button>
  </div>
  <div class="hint">${soTb("MISC.tbHint", "Both open your browser's print dialog. For a PDF, set <b>Destination</b> to &ldquo;Save as PDF&rdquo;.", "كلا الزرين يفتحان نافذة الطباعة. لحفظ ملف PDF اختر من قائمة <b>الوجهة</b> خيار الحفظ بصيغة PDF.")}</div>
  <script>function savePDF(){var t=document.title,d=new Date(),p=function(n){return(n<10?"0":"")+n};document.title="SteelOptimizer_Sections_"+d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate());var r=function(){document.title=t;window.removeEventListener("afterprint",r)};window.addEventListener("afterprint",r);window.print();setTimeout(r,4000);}<\/script></body></html>`;
  downloadHTML(html, "SteelOptimizer_Sections.html");
}

function sectionLeftovers(results) {
  const rows = [];
  results.groups.forEach(g => {
    g.bins.forEach((b, i) => { if (b.remaining >= 1000) rows.push({ profile: g.profile, grade: g.grade || "", bar: RT(`Bar ${i + 1}`, `عود ${i + 1}`, (L) => L === "ru" ? `Хлыст ${i + 1}` : L === "zh" ? `第 ${i + 1} 根` : `Barra ${i + 1}`), stock: b.stockLength, len: Math.round(b.remaining), kg: g.kgm ? (b.remaining / 1000) * g.kgm : 0 }); });
    (g.reusedBins || []).forEach((b, i) => { if (b.cuts.length && b.remaining >= 1000) rows.push({ profile: g.profile, grade: g.grade || "", bar: RT(`Leftover ${i + 1}`, `بواقٍ ${i + 1}`, (L) => L === "ru" ? `Остаток ${i + 1}` : L === "zh" ? `余料 ${i + 1}` : `Sobrante ${i + 1}`), stock: b.stockLength, len: Math.round(b.remaining), kg: g.kgm ? (b.remaining / 1000) * g.kgm : 0 }); });
  });
  return rows;
}
// Build the re-import payload for section leftovers: aggregate by profile+grade+length → qty.
function sectionLeftoverPayload(rows) {
  const agg = {};
  rows.forEach(r => { const grade = (r.grade && r.grade !== "Steel (grade not specified)") ? r.grade : ""; const k = `${r.profile}||${grade}||${r.len}`; if (!agg[k]) agg[k] = { profile: r.profile, grade, length: r.len, qty: 0 }; agg[k].qty += 1; });
  return { t: "sec", items: Object.values(agg) };
}
function exportSectionLeftoverExcel(results) {
  const wb = XLSX.utils.book_new();
  const rows = sectionLeftovers(results);
  const totalKg = rows.reduce((s, r) => s + (r.kg || 0), 0);
  const aoa = [
    [RT("STEEL OPTIMIZER — REUSABLE BAR OFFCUTS", "STEEL OPTIMIZER — قصاصات الأعواد الصالحة")],
    [RT(`Generated: ${new Date().toLocaleDateString()}`, `تاريخ الإنشاء: ${new Date().toLocaleDateString()}`, (L) => L === "ru" ? `Сформирован: ${new Date().toLocaleDateString(soDateLoc())}` : L === "zh" ? `生成日期：${new Date().toLocaleDateString(soDateLoc())}` : `Generado: ${new Date().toLocaleDateString(soDateLoc())}`)],
    [RT("Leftover bar ends ≥ 1 m — keep these to cut from on your next job instead of buying new.", "نهايات أعواد متبقية ≥ 1 م — احتفظ بها للقص منها في مشروعك القادم بدل الشراء.")],
    [],
    [RT("Profile", "المقطع"), RT("Grade", "الرتبة"), RT("From Bar #", "من العود رقم"), RT("Available Length (mm)", "الطول المتوفر (مم)"), RT("Reusable Leftover (mm)", "البواقي الصالحة (مم)"), RT("Weight (kg)", "الوزن (كجم)"), RT("Keep?", "احتفظ؟")],
  ];
  if (rows.length) { rows.forEach(r => aoa.push([r.profile, soGrade(r.grade) || "—", r.bar, r.stock, r.len, r.kg ? r.kg.toFixed(1) : "—", RT("YES", "نعم")])); aoa.push([], [RT("TOTAL REUSABLE", "إجمالي الصالح"), "", "", "", `${rows.length} ${RT("pieces", "قطعة")}`, totalKg.toFixed(1), ""]); }
  else aoa.push([RT("No reusable bar offcuts ≥ 1 m — material was used efficiently.", "لا توجد قصاصات أعواد صالحة ≥ 1 م — تم استغلال المادة بكفاءة.")]);
  const ws = XLSX.utils.aoa_to_sheet(aoa); ws["!cols"] = [{ wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 20 }, { wch: 22 }, { wch: 12 }, { wch: 8 }];
  XLSX.utils.book_append_sheet(wb, ws, RT("Reusable Offcuts", "القصاصات الصالحة"));
  // Machine-readable sheet so this exact file can be dropped back into the optimizer.
  const token = encodeLeftoverToken(sectionLeftoverPayload(rows));
  const tws = XLSX.utils.aoa_to_sheet([[RT("Steel Optimizer re-import data — keep this sheet.", "بيانات إعادة الاستيراد لـ Steel Optimizer — لا تحذف هذه الورقة.")], [RT("Drop this file into the “Reuse a previous Leftover file” button to load these offcuts again.", "أفلت هذا الملف في زر «إعادة استخدام ملف بواقٍ سابق» لتحميل هذه القصاصات مرة أخرى.")], [token]]);
  tws["!cols"] = [{ wch: 120 }];
  XLSX.utils.book_append_sheet(wb, tws, "_reimport");
  saveWorkbook(wb, "SteelOptimizer_Section_Leftovers.xlsx");
}
function exportSectionLeftoverPDF(results) {
  const rows = sectionLeftovers(results);
  const totalKg = rows.reduce((s, r) => s + (r.kg || 0), 0);
  const body = rows.length ? rows.map(r => `<tr><td>${r.profile}</td><td>${soGrade(r.grade) || "—"}</td><td>${r.bar}</td><td>${fmtMm(r.stock)}</td><td><b>${fmtMm(r.len)}</b></td><td>${r.kg ? fmtKg(r.kg) : "—"}</td><td>&#10003; ${RT("keep", "احتفظ")}</td></tr>`).join("") : `<tr><td colspan="7">${RT("No reusable bar offcuts &ge; 1 m — material was used efficiently.", "لا توجد قصاصات أعواد صالحة &ge; 1 م — تم استغلال المادة بكفاءة.")}</td></tr>`;
  const payload = sectionLeftoverPayload(rows);
  const token = encodeLeftoverToken(payload);
  const html = `<html${RT_HTML()}><head><title>${RT("Steel Optimizer — Reusable Offcuts", "Steel Optimizer — القصاصات الصالحة")}</title><style>
  body{font-family:Georgia,serif;padding:32px;color:#04140d}
  h1{color:#047857;border-bottom:2px solid #a7f3d0;padding-bottom:10px}
  table{width:100%;border-collapse:collapse;margin-top:12px;font-family:monospace;font-size:13px}
  th{background:#ecfdf5;color:#047857;padding:9px;text-align:left;font-size:11px}
  td{padding:9px;border-bottom:1px solid #eee}
  .save{margin-top:18px;padding:18px;background:#ecfdf5;border:2px solid #10b981;border-radius:8px;font-size:18px}
  .save b{font-size:30px;color:#047857}
  .reimport{margin-top:18px;padding:10px 12px;background:#f8fafc;border:1px dashed #94a3b8;border-radius:6px;font-size:11px;color:#475569}
  .tok{font-family:monospace;font-size:7px;color:#9ca3af;word-break:break-all;line-height:1.3;margin-top:6px}
  .btnrow{margin-top:24px;display:flex;gap:12px;flex-wrap:wrap}
  button{padding:12px 26px;border-radius:6px;font-size:14px;cursor:pointer;font-family:inherit;line-height:1.2}
  .b-save{background:#047857;color:#fff;border:2px solid #047857}
  .b-print{background:#fff;color:#047857;border:2px solid #047857}
  .hint{margin-top:10px;font-size:11px;color:#64748b;line-height:1.7}
  @media print{button,.btnrow,.hint{display:none}}${RT_CSS()}
  </style></head><body>
  <h1>&#9851; ${RT("Reusable Bar Offcuts — leftover steel for the next project", "قصاصات الأعواد الصالحة — حديد متبقٍ للمشروع القادم")}</h1>
  <p>${new Date().toLocaleDateString(soDateLoc("en-GB"))}</p>
  <div class="save">${RT("Reusable leftover:", "البواقي الصالحة:")} <b>${fmtKg(totalKg)}</b> &nbsp;&middot; ${rows.length} ${RT(`bar end${rows.length === 1 ? "" : "s"}`, "نهاية عود", (L) => L === "ru" ? `шт.` : L === "zh" ? `根料头` : `${rows.length !== 1 ? "extremos de barra" : "extremo de barra"}`)} &ge; 1 ${RT("m", "م")}</div>
  <table><tr><th>${RT("Profile", "المقطع")}</th><th>${RT("Grade", "الرتبة")}</th><th>${RT("From Bar", "من العود")}</th><th>${RT("Available Length", "الطول المتوفر")}</th><th>${RT("Reusable Leftover", "البواقي الصالحة")}</th><th>${RT("Weight", "الوزن")}</th><th>${RT("Status", "الحالة")}</th></tr>${body}</table>
  <div class="reimport">&#9851; ${RT("Re-import code — drop this PDF back into the optimizer's <b>“Reuse a previous Leftover file”</b> button to load these offcuts again.", "رمز إعادة الاستيراد — أفلت ملف PDF هذا في زر <b>«إعادة استخدام ملف بواقٍ سابق»</b> لتحميل هذه القصاصات مرة أخرى.")}<div class="tok">${token.replace(/</g, "&lt;").replace(/>/g, "&gt;")}</div></div>
  <script type="application/json" id="steelopt">${JSON.stringify(payload).replace(/</g, "\\u003c")}</script>
  <div class="btnrow">
  <button class="b-save" onclick="savePDF()">&#8681; ${soTb("MISC.tbSave", "Save as PDF", "حفظ PDF")}</button>
  <button class="b-print" onclick="window.print()">&#128424; ${soTb("MISC.tbPrint", "Print", "طباعة")}</button>
  </div>
  <div class="hint">${soTb("MISC.tbHint", "Both open your browser's print dialog. For a PDF, set <b>Destination</b> to &ldquo;Save as PDF&rdquo;.", "كلا الزرين يفتحان نافذة الطباعة. لحفظ ملف PDF اختر من قائمة <b>الوجهة</b> خيار الحفظ بصيغة PDF.")}</div>
  <script>function savePDF(){var t=document.title,d=new Date(),p=function(n){return(n<10?"0":"")+n};document.title="SteelOptimizer_Section_Leftovers_"+d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate());var r=function(){document.title=t;window.removeEventListener("afterprint",r)};window.addEventListener("afterprint",r);window.print();setTimeout(r,4000);}<\/script></body></html>`;
  downloadHTML(html, "SteelOptimizer_Section_Leftovers.html");
}

/* Last line of defense: an unexpected render error currently yields a blank
   white page. This catches it and offers a one-click reload instead. Saved
   preferences (language, price, license) live in localStorage and survive. */
class ErrorBoundary extends Component {
  constructor(props) { super(props); this.state = { err: null }; }
  static getDerivedStateFromError(err) { return { err }; }
  render() {
    if (!this.state.err) return this.props.children;
    return (
      <div style={{ minHeight: "100vh", background: "#0b0f17", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
        <div style={{ maxWidth: 480, textAlign: "center", fontFamily: "'Space Mono', monospace" }}>
          <div style={{ fontSize: 46, marginBottom: 14 }}>⚠️</div>
          <div style={{ color: "#f8fafc", fontSize: 18, fontWeight: 700, marginBottom: 14 }}>{soTb("MISC.errTitle", "An unexpected error occurred — please reload the page.", "حدث خطأ غير متوقع — أعد تحميل الصفحة.")}</div>
          <div style={{ color: "#64748b", fontSize: 13, marginBottom: 20, direction: "ltr", wordBreak: "break-word" }}>{String((this.state.err && this.state.err.message) || this.state.err)}</div>
          <button onClick={() => window.location.reload()} style={{ padding: "10px 24px", borderRadius: 8, border: "1px solid #f59e0b", background: "rgba(245,158,11,.15)", color: "#fbbf24", fontWeight: 700, fontSize: 16, cursor: "pointer", fontFamily: "inherit" }}>{soTb("MISC.errReload", "Reload", "إعادة التحميل")}</button>
        </div>
      </div>
    );
  }
}

/* ╔══════════════════════════════════════════════════════════════════════════╗
   ║  MEMBERSHIP — email + password accounts and a paid plan (isolated add-on) ║
   ╚══════════════════════════════════════════════════════════════════════════╝
     START OPTIMIZING  → sign in / create account (email + password)
     ⚡ OPTIMIZE        → signed in + active plan; otherwise: choose plan → pay → it runs

   Accounts: Supabase Auth (passwords are stored hashed by Supabase, never here).
   Payment:  Lemon Squeezy checkout — cards (Visa, Mastercard, Amex…), Apple Pay
             (Safari), Google Pay (Chrome), PayPal.
   The plan: Lemon Squeezy calls /api/lemon-webhook (Vercel) → Supabase table
             «memberships» → the app reads the signed-in user's own row.

   OFF until supabaseUrl, supabaseAnonKey and one checkout link are filled in:
   with any of them empty the app behaves exactly as before.
──────────────────────────────────────────────────────────────────────────── */
const MEMBERSHIP = {
  supabaseUrl: "https://cjeqlypqoxiwnleimwlx.supabase.co",   // ◄ Supabase project URL
  supabaseAnonKey: "sb_publishable_neg7qlt_WHT67FgGweLc1A_V0lFd5yo",   // ◄ publishable (public) key — NEVER put the secret key here
  checkoutMonthly: "",      // ◄ Lemon Squeezy → Products → monthly variant → Share → checkout link
  checkoutYearly: "",       // ◄ same for the yearly variant (leave empty for monthly only)
  priceMonthly: "$29",      // shown on the plan buttons
  priceYearly: "$228",
  requirePaid: false,       // ◄ true once the Lemon Squeezy checkout links above are filled in (then payment is required)
};
// ◄ PAYMENTS: false = the app is FREE. Login (email + password) stays; no plans, no prices, no checkout,
//   no Lemon Squeezy script, no subscription look-ups. true + checkout links + requirePaid = paid plan.
const PAYMENTS_ENABLED = false;
const MEMBERSHIP_ON = !!(MEMBERSHIP.supabaseUrl && MEMBERSHIP.supabaseAnonKey && (!(PAYMENTS_ENABLED && MEMBERSHIP.requirePaid) || MEMBERSHIP.checkoutMonthly || MEMBERSHIP.checkoutYearly));

const SO_MT = {
  signIn:     { en: "Sign in", ar: "تسجيل الدخول", ru: "Войти", zh: "登录", es: "Iniciar sesión" },
  signUp:     { en: "Create account", ar: "إنشاء حساب", ru: "Создать аккаунт", zh: "注册账号", es: "Crear cuenta" },
  titleIn:    { en: "Sign in to Steel Optimizer", ar: "سجّل الدخول إلى Steel Optimizer", ru: "Вход в Steel Optimizer", zh: "登录 Steel Optimizer", es: "Inicie sesión en Steel Optimizer" },
  titleUp:    { en: "Create your account", ar: "أنشئ حسابك", ru: "Создайте аккаунт", zh: "创建您的账号", es: "Cree su cuenta" },
  leadFree:   { en: "Free — just an email and a password.", ar: "مجاناً — فقط بريد إلكتروني وكلمة مرور.", ru: "Бесплатно — нужны только email и пароль.", zh: "免费 — 只需邮箱和密码。", es: "Gratis: solo un correo y una contraseña." },
  lead:       { en: "One account for your plan, on any device.", ar: "حساب واحد لاشتراكك على أي جهاز.", ru: "Один аккаунт для вашей подписки — на любом устройстве.", zh: "一个账号，任何设备都能使用您的订阅。", es: "Una cuenta para su plan, en cualquier dispositivo." },
  email:      { en: "Email", ar: "البريد الإلكتروني", ru: "Эл. почта", zh: "邮箱", es: "Correo electrónico" },
  password:   { en: "Password", ar: "كلمة المرور", ru: "Пароль", zh: "密码", es: "Contraseña" },
  passHint:   { en: "At least 8 characters", ar: "8 أحرف على الأقل", ru: "Не менее 8 символов", zh: "至少 8 个字符", es: "Al menos 8 caracteres" },
  toUp:       { en: "New here? Create an account", ar: "جديد هنا؟ أنشئ حسابًا", ru: "Впервые здесь? Создайте аккаунт", zh: "第一次使用？注册账号", es: "¿Es nuevo? Cree una cuenta" },
  toIn:       { en: "Already have an account? Sign in", ar: "لديك حساب؟ سجّل الدخول", ru: "Уже есть аккаунт? Войдите", zh: "已有账号？登录", es: "¿Ya tiene cuenta? Inicie sesión" },
  forgot:     { en: "Forgot password?", ar: "نسيت كلمة المرور؟", ru: "Забыли пароль?", zh: "忘记密码？", es: "¿Olvidó su contraseña?" },
  sendLink:   { en: "Send reset link", ar: "أرسل رابط إعادة التعيين", ru: "Отправить ссылку", zh: "发送重置链接", es: "Enviar enlace" },
  resetSent:  { en: "If this email has an account, a link to set a new password is on its way.", ar: "إذا كان لهذا البريد حساب، فسيصلك رابط لتعيين كلمة مرور جديدة.", ru: "Если этот адрес зарегистрирован, на него придёт ссылка для смены пароля.", zh: "如果该邮箱已注册，我们会发送设置新密码的链接。", es: "Si este correo tiene una cuenta, recibirá un enlace para crear una nueva contraseña." },
  confirmSent:{ en: "Check your email and confirm your account, then sign in.", ar: "افتح بريدك وأكّد حسابك، ثم سجّل الدخول.", ru: "Проверьте почту и подтвердите аккаунт, затем войдите.", zh: "请查收邮件并确认账号，然后登录。", es: "Revise su correo y confirme su cuenta; luego inicie sesión." },
  newPass:    { en: "Set a new password", ar: "عيّن كلمة مرور جديدة", ru: "Задайте новый пароль", zh: "设置新密码", es: "Cree una nueva contraseña" },
  savePass:   { en: "Save password", ar: "احفظ كلمة المرور", ru: "Сохранить пароль", zh: "保存密码", es: "Guardar contraseña" },
  errBad:     { en: "Wrong email or password.", ar: "البريد أو كلمة المرور غير صحيحة.", ru: "Неверная почта или пароль.", zh: "邮箱或密码错误。", es: "Correo o contraseña incorrectos." },
  errConfirm: { en: "Please confirm your email first — check your inbox.", ar: "أكّد بريدك الإلكتروني أولًا — افتح صندوق الوارد.", ru: "Сначала подтвердите почту — проверьте входящие.", zh: "请先确认邮箱 — 请查看收件箱。", es: "Primero confirme su correo: revise su bandeja de entrada." },
  errExists:  { en: "This email already has an account — sign in instead.", ar: "لهذا البريد حساب مسبقًا — سجّل الدخول بدلًا من ذلك.", ru: "Этот адрес уже зарегистрирован — войдите.", zh: "该邮箱已注册，请直接登录。", es: "Este correo ya tiene una cuenta: inicie sesión." },
  errWeak:    { en: "Password must be at least 8 characters.", ar: "يجب أن تكون كلمة المرور 8 أحرف على الأقل.", ru: "Пароль должен содержать не менее 8 символов.", zh: "密码至少需要 8 个字符。", es: "La contraseña debe tener al menos 8 caracteres." },
  errEmail:   { en: "Enter a valid email address.", ar: "أدخل بريدًا إلكترونيًا صحيحًا.", ru: "Введите корректный адрес эл. почты.", zh: "请输入有效的邮箱地址。", es: "Introduzca un correo válido." },
  errNet:     { en: "Could not reach the server. Check your connection and try again.", ar: "تعذّر الاتصال بالخادم. تحقق من الاتصال وحاول مجددًا.", ru: "Не удалось связаться с сервером. Проверьте подключение и повторите попытку.", zh: "无法连接服务器，请检查网络后重试。", es: "No se pudo conectar con el servidor. Compruebe su conexión e inténtelo de nuevo." },
  payTitle:   { en: "Choose your plan", ar: "اختر اشتراكك", ru: "Выберите тариф", zh: "选择您的套餐", es: "Elija su plan" },
  payLead:    { en: "Unlimited optimizations, PDF & Excel reports and leftover files.", ar: "تحسين بلا حدود، وتقارير PDF و Excel، وملفات البواقي.", ru: "Неограниченные расчёты, отчёты PDF и Excel, файлы деловых остатков.", zh: "无限次优化、PDF 与 Excel 报告、余料文件。", es: "Optimizaciones ilimitadas, informes PDF y Excel y archivos de sobrantes." },
  monthly:    { en: "Monthly", ar: "شهري", ru: "Помесячно", zh: "按月", es: "Mensual" },
  yearly:     { en: "Yearly", ar: "سنوي", ru: "На год", zh: "按年", es: "Anual" },
  perMo:      { en: "/month", ar: "/شهر", ru: "/мес.", zh: "/月", es: "/mes" },
  perYr:      { en: "/year", ar: "/سنة", ru: "/год", zh: "/年", es: "/año" },
  secure:     { en: "Secure checkout by Lemon Squeezy — card details never touch this site.", ar: "دفع آمن عبر Lemon Squeezy — بيانات البطاقة لا تمر عبر هذا الموقع.", ru: "Безопасная оплата через Lemon Squeezy — данные карты не проходят через этот сайт.", zh: "由 Lemon Squeezy 提供安全支付 — 卡信息不会经过本网站。", es: "Pago seguro con Lemon Squeezy: los datos de la tarjeta nunca pasan por este sitio." },
  waiting:    { en: "Confirming your payment…", ar: "جارٍ تأكيد الدفع…", ru: "Подтверждаем оплату…", zh: "正在确认付款…", es: "Confirmando su pago…" },
  checkAgain: { en: "I've paid — check again", ar: "دفعت — تحقّق مرة أخرى", ru: "Я оплатил — проверить снова", zh: "我已付款 — 再次检查", es: "Ya pagué: comprobar de nuevo" },
  notYet:     { en: "Payment not confirmed yet. If you have paid, wait a minute and check again.", ar: "لم يُؤكَّد الدفع بعد. إذا كنت دفعت، انتظر دقيقة ثم تحقّق مجددًا.", ru: "Оплата ещё не подтверждена. Если вы оплатили, подождите минуту и проверьте снова.", zh: "付款尚未确认。如果您已付款，请稍等一分钟后再检查。", es: "El pago aún no está confirmado. Si ya pagó, espere un minuto y compruebe de nuevo." },
  backPlans:  { en: "Back to plans", ar: "رجوع إلى الاشتراكات", ru: "Назад к тарифам", zh: "返回套餐", es: "Volver a los planes" },
  signedAs:   { en: "Signed in as {email}", ar: "مسجّل الدخول باسم {email}", ru: "Вы вошли как {email}", zh: "已登录：{email}", es: "Sesión iniciada como {email}" },
  signOut:    { en: "Sign out", ar: "تسجيل الخروج", ru: "Выйти", zh: "退出登录", es: "Cerrar sesión" },
  manage:     { en: "Manage subscription", ar: "إدارة الاشتراك", ru: "Управление подпиской", zh: "管理订阅", es: "Gestionar suscripción" },
  close:      { en: "Close", ar: "إغلاق", ru: "Закрыть", zh: "关闭", es: "Cerrar" },
  checking:   { en: "Checking your plan…", ar: "جارٍ التحقق من اشتراكك…", ru: "Проверяем подписку…", zh: "正在检查您的订阅…", es: "Comprobando su plan…" },
};
const soMT = (lang, k, vars) => {
  const e = SO_MT[k]; let s = e ? (e[lang] != null ? e[lang] : e.en) : k;
  if (vars) for (const v in vars) s = s.replace("{" + v + "}", vars[v]);
  return s;
};

/* ── session + plan store (module level, shared by the modal and the menu) ── */
const SO_SESSION_KEY = "steelopt_session_v1";
const soMem = { session: undefined, member: null, checkedAt: 0, listeners: new Set() };
const soMemEmit = () => soMem.listeners.forEach(f => { try { f(); } catch { /* ignore */ } });
function useSoMem() {
  const [, force] = useState(0);
  useEffect(() => { const f = () => force(x => x + 1); soMem.listeners.add(f); return () => { soMem.listeners.delete(f); }; }, []);
  return soMem;
}
async function soApi(path, { body, token, method } = {}) {
  const r = await fetch(MEMBERSHIP.supabaseUrl.replace(/\/+$/, "") + path, {
    method: method || (body ? "POST" : "GET"),
    headers: { apikey: MEMBERSHIP.supabaseAnonKey, "Content-Type": "application/json",
      ...(token ? { Authorization: "Bearer " + token } : MEMBERSHIP.supabaseAnonKey.split(".").length === 3 ? { Authorization: "Bearer " + MEMBERSHIP.supabaseAnonKey } : {}) },
    body: body ? JSON.stringify(body) : undefined,
  });
  let data = null; try { data = await r.json(); } catch { data = null; }
  if (!r.ok) {
    const e = new Error(String((data && (data.error_description || data.msg || data.message || data.error)) || "HTTP " + r.status));
    e.status = r.status; e.code = String((data && (data.error_code || data.code || data.error)) || "");
    e.cfg = (r.status === 401 || r.status === 403) && /api ?key|apikey/i.test(e.message + " " + e.code);   // project key wrong → configuration, not the user
    throw e;
  }
  return data;
}
// Supabase-side problems (not the visitor's fault): offline, 5xx, rate limits, e-mail sending / not-authorised address, unconfirmed e-mail
function soSvcProblem(e) {
  if (!e || e.uiMsg) return false;
  if (e.cfg || !e.status || e.status >= 500 || e.status === 429) return true;
  return /email_address_not_authorized|over_email_send_rate_limit|email_provider_disabled|signup_disabled|email_not_confirmed/i.test(e.code || "")
    || /error sending|not authorized|rate limit|not confirmed|signups? not allowed|disabled/i.test(e.message || "");
}
function soSetSession(d) {
  if (!d || !d.access_token) {
    soMem.session = null; soMem.member = null; soMem.checkedAt = 0;
    try { localStorage.removeItem(SO_SESSION_KEY); } catch { /* private mode */ }
  } else {
    const user = d.user || (soMem.session && soMem.session.user) || {};
    soMem.session = {
      access_token: d.access_token, refresh_token: d.refresh_token || (soMem.session && soMem.session.refresh_token),
      expires_at: d.expires_at ? d.expires_at * 1000 : Date.now() + (+d.expires_in || 3600) * 1000,
      user: { id: user.id, email: user.email },
    };
    try {
      localStorage.setItem(SO_SESSION_KEY, JSON.stringify(soMem.session));
      if (user.email && !localStorage.getItem("steelopt_email_v2")) localStorage.setItem("steelopt_email_v2", user.email);   // download e-mail step already answered
    } catch { /* private mode */ }
  }
  soMemEmit();
}
async function soSession() {
  if (soMem.session === undefined) { try { soMem.session = JSON.parse(localStorage.getItem(SO_SESSION_KEY) || "null"); } catch { soMem.session = null; } }
  const s = soMem.session;
  if (!s || !s.access_token) return null;
  if (s.expires_at - Date.now() > 60000 && s.user && s.user.id) return s;
  try { soSetSession(await soApi("/auth/v1/token?grant_type=refresh_token", { body: { refresh_token: s.refresh_token } })); return soMem.session; }
  catch (e) { if (e.status >= 400 && e.status < 500) soSetSession(null); return e.status >= 400 && e.status < 500 ? null : s; }
}
function soHasPlan(m) {
  if (!m) return false;
  if (m.status === "active" || m.status === "on_trial" || m.status === "past_due") return true;
  return m.status === "cancelled" && m.ends_at && Date.parse(m.ends_at) > Date.now();   // paid until the period ends
}
async function soLoadPlan() {
  if (!PAYMENTS_ENABLED) return null;   // free app: nothing to look up
  const s = await soSession();
  if (!s) { soMem.member = null; soMemEmit(); return null; }
  try {
    const rows = await soApi("/rest/v1/memberships?select=status,ends_at,renews_at,trial_ends_at,portal_url,updated_at&order=updated_at.desc", { token: s.access_token });
    soMem.member = Array.isArray(rows) && rows.length ? (rows.find(soHasPlan) || rows[0]) : null;
    soMem.checkedAt = Date.now();
  } catch { /* keep what we had */ }
  soMemEmit();
  return soMem.member;
}
function soSignOut() {
  const s = soMem.session;
  if (s && s.access_token) soApi("/auth/v1/logout", { body: {}, token: s.access_token }).catch(() => {});
  soSetSession(null);
}
let _soLemon = null;
function soLoadLemon() {
  if (!PAYMENTS_ENABLED) return Promise.reject(new Error("payments off"));
  if (window.LemonSqueezy) return Promise.resolve(window.LemonSqueezy);
  if (_soLemon) return _soLemon;
  _soLemon = new Promise((resolve, reject) => {
    const sc = document.createElement("script");
    sc.src = "https://app.lemonsqueezy.com/js/lemon.js"; sc.defer = true;
    sc.onload = () => { try { window.createLemonSqueezy && window.createLemonSqueezy(); } catch { /* ignore */ } window.LemonSqueezy ? resolve(window.LemonSqueezy) : reject(new Error("lemon.js")); };
    sc.onerror = () => { _soLemon = null; reject(new Error("lemon.js")); };
    document.head.appendChild(sc);
  });
  return _soLemon;
}

/* ── the gate: every protected button calls this ── */
const soGateBus = { open: null };
function soRequireMember(action, needPaid) {
  if (!MEMBERSHIP_ON || !soGateBus.open || soMem.cfgBroken) { action(); return; }   // not configured (or project key rejected) → exactly as before
  soGateBus.open(action, PAYMENTS_ENABLED && !!needPaid && MEMBERSHIP.requirePaid);
}

/* ── the modal (sign in · create account · reset password · plans · payment) ── */
function SoMemberHost() {
  const { lang, dir } = useLang();
  const L = (k, v) => soMT(lang, k, v);
  const [view, setView] = useState(null);      // null | in | up | forgot | reset | check | pay | wait
  const [email, setEmail] = useState(""), [pass, setPass] = useState("");
  const [busy, setBusy] = useState(false), [err, setErr] = useState(""), [msg, setMsg] = useState("");
  const pending = useRef(null), poll = useRef(null);
  const stopPoll = () => { if (poll.current) { clearInterval(poll.current); poll.current = null; } };
  const finish = () => { stopPoll(); const p = pending.current; pending.current = null; setView(null); setErr(""); setMsg(""); if (p && p.action) setTimeout(p.action, 0); };
  const afterAuth = async () => {
    const p = pending.current;
    if (!p) { setView(null); return; }
    if (!p.needPaid) { finish(); return; }
    if (soHasPlan(soMem.member) && Date.now() - soMem.checkedAt < 10 * 60000) { finish(); return; }
    setView("check");
    const m = await soLoadPlan();
    if (soHasPlan(m)) finish(); else { setErr(""); setView("pay"); }
  };
  const open = async (action, needPaid) => {
    if (pending.current && view) return;
    pending.current = { action, needPaid }; setErr(""); setMsg("");
    const s = await soSession();
    if (!s) { setView("in"); return; }
    await afterAuth();
  };
  useEffect(() => {
    if (!MEMBERSHIP_ON) return undefined;
    soGateBus.open = (a, n) => open(a, n);
    return () => { soGateBus.open = null; };
  });
  useEffect(() => {
    if (!MEMBERSHIP_ON) return undefined;
    // links from Supabase e-mails: #access_token=…&type=signup|recovery
    try {
      const h = new URLSearchParams((window.location.hash || "").replace(/^#/, ""));
      if (h.get("access_token")) {
        const tok = h.get("access_token"), type = h.get("type");
        soSetSession({ access_token: tok, refresh_token: h.get("refresh_token"), expires_in: +h.get("expires_in") || 3600, user: {} });
        soApi("/auth/v1/user", { token: tok }).then(u => soSetSession({ access_token: tok, refresh_token: h.get("refresh_token"), expires_in: +h.get("expires_in") || 3600, user: u })).catch(() => {});
        window.history.replaceState(null, "", window.location.pathname + window.location.search);
        if (type === "recovery") { setPass(""); setView("reset"); }
      }
    } catch { /* ignore */ }
    soSession().then(s => { if (s) soLoadPlan(); });
    return () => stopPoll();
  }, []);
  if (!MEMBERSHIP_ON || !view) return null;

  const okEmail = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email.trim());
  const freeStep = () => !!pending.current && !pending.current.needPaid;
  const run = async fn => {
    setBusy(true); setErr(""); setMsg("");
    try { await fn(); }
    catch (e) {
      if (e && e.cfg) { soMem.cfgBroken = true; try { console.warn("[membership] Supabase key rejected — check MEMBERSHIP.supabaseAnonKey; letting the user through"); } catch { /* ignore */ } finish(); return; }
      // a problem on the Supabase side (e-mail not deliverable: the built-in mailer reaches only the project team,
      // 2 per hour · rate limit · outage · offline · unconfirmed e-mail) must never lock a visitor out of a FREE step
      if (e && !e.uiMsg && freeStep() && soSvcProblem(e)) {
        soMem.cfgBroken = true;
        try { alertMe("MEMBERSHIP PROBLEM — visitor let through", { status: e.status || 0, code: String(e.code || "").slice(0, 40) }); } catch { /* ignore */ }
        finish(); return;
      }
      setErr(e && e.uiMsg ? e.uiMsg : L("errNet"));
    } finally { setBusy(false); }
  };
  const fail = k => { const e = new Error(k); e.uiMsg = L(k); throw e; };
  const signIn = () => run(async () => {
    if (!okEmail) fail("errEmail");
    try { soSetSession(await soApi("/auth/v1/token?grant_type=password", { body: { email: email.trim(), password: pass } })); }
    catch (e) { if (freeStep() && /not.?confirmed/i.test(e.message + " " + e.code)) throw e; if (!e.cfg && e.status >= 400 && e.status < 500) fail(/confirm/i.test(e.message + e.code) ? "errConfirm" : "errBad"); throw e; }
    setPass(""); await afterAuth();
  });
  const signUp = () => run(async () => {
    if (!okEmail) fail("errEmail");
    if (pass.length < 8) fail("errWeak");
    let d;
    try { d = await soApi("/auth/v1/signup", { body: { email: email.trim(), password: pass } }); }
    catch (e) { if (!e.cfg && /already|registered|exists/i.test(e.message + e.code)) fail("errExists"); if (!e.cfg && /password/i.test(e.message + e.code)) fail("errWeak"); throw e; }
    setPass("");
    if (d && d.access_token) { soSetSession(d); await afterAuth(); }
    else if (freeStep()) { soMem.cfgBroken = true; finish(); }   // "Confirm email" is on: account saved, don't make them wait for the e-mail
    else { setView("in"); setMsg(L("confirmSent")); }
  });
  const forgot = () => run(async () => {
    if (!okEmail) fail("errEmail");
    const back = window.location.origin + window.location.pathname;
    await soApi("/auth/v1/recover?redirect_to=" + encodeURIComponent(back), { body: { email: email.trim() } });
    setMsg(L("resetSent"));
  });
  const savePass = () => run(async () => {
    if (pass.length < 8) fail("errWeak");
    const s = await soSession(); if (!s) { setView("in"); return; }
    await soApi("/auth/v1/user", { body: { password: pass }, token: s.access_token, method: "PUT" });
    setPass(""); if (pending.current) await afterAuth(); else setView(null);
  });
  const checkNow = async () => { setBusy(true); const m = await soLoadPlan(); setBusy(false); if (soHasPlan(m)) finish(); else setErr(L("notYet")); };
  const pay = plan => run(async () => {
    const s = await soSession(); if (!s) { setView("in"); return; }
    const u = new URL(plan === "y" ? MEMBERSHIP.checkoutYearly : MEMBERSHIP.checkoutMonthly);
    u.searchParams.set("checkout[email]", s.user.email || email.trim());
    u.searchParams.set("checkout[custom][user_id]", s.user.id);
    setView("wait");
    stopPoll(); let n = 0;
    poll.current = setInterval(async () => { n++; const m = await soLoadPlan(); if (soHasPlan(m)) finish(); else if (n >= 60) stopPoll(); }, 4000);
    try {
      const LS = await soLoadLemon();
      try { LS.Setup({ eventHandler: e => { if (e && e.event === "Checkout.Success") soLoadPlan().then(m => { if (soHasPlan(m)) finish(); }); } }); } catch { /* ignore */ }
      u.searchParams.set("embed", "1");
      LS.Url.Open(u.toString());
    } catch { window.open(u.toString(), "_blank", "noopener"); }
  });
  const close = () => { stopPoll(); pending.current = null; setView(null); setErr(""); setMsg(""); };

  const ACC = "#fbbf24";
  const INP = { width: "100%", boxSizing: "border-box", padding: "13px 14px", borderRadius: 10, background: "#0a0f18", border: "1px solid #243044", color: "#e2e8f0", fontSize: 16, outline: "none", marginTop: 6 };
  const LBL = { display: "block", fontSize: 12.5, color: "#94a3b8", letterSpacing: 0.6, marginTop: 16, fontWeight: 600 };
  const BTN = { width: "100%", marginTop: 22, padding: "14px 18px", borderRadius: 10, border: "none", cursor: busy ? "default" : "pointer", background: busy ? "rgba(120,80,10,.45)" : "linear-gradient(135deg,#f59e0b,#d97706)", color: "#1a1206", fontSize: 16.5, fontWeight: 800, letterSpacing: 0.4, boxShadow: "0 10px 30px -8px rgba(245,158,11,.5)" };
  const LINK = { background: "none", border: "none", color: ACC, cursor: "pointer", fontSize: 14, padding: 0, textDecoration: "underline", textUnderlineOffset: 3 };
  const s = soMem.session;
  const title = view === "up" ? L("titleUp") : view === "forgot" ? L("forgot") : view === "reset" ? L("newPass")
    : view === "pay" || view === "wait" ? L("payTitle") : view === "check" ? L("checking") : L("titleIn");
  const plan = (k, price, per) => (
    <button key={k} disabled={busy} onClick={() => pay(k)} style={{ flex: "1 1 160px", padding: "18px 16px", borderRadius: 12, cursor: "pointer", textAlign: "start",
      background: k === "y" ? "rgba(245,158,11,.12)" : "#0a0f18", border: `1px solid ${k === "y" ? "rgba(245,158,11,.55)" : "#243044"}`, color: "#e2e8f0" }}>
      <div style={{ fontSize: 13, color: "#94a3b8", fontWeight: 700, letterSpacing: 0.6 }}>{L(k === "y" ? "yearly" : "monthly")}</div>
      <div style={{ marginTop: 6 }}><span style={{ fontSize: 30, fontWeight: 900, color: "#f8fafc" }}>{price}</span><span style={{ fontSize: 14, color: "#94a3b8" }}> {per}</span></div>
    </button>);
  return (
    <div role="dialog" aria-modal="true" onClick={close} style={{ position: "fixed", inset: 0, zIndex: 10050, background: "rgba(3,6,11,.8)", backdropFilter: "blur(8px)", WebkitBackdropFilter: "blur(8px)", display: "flex", alignItems: "flex-start", justifyContent: "center", overflowY: "auto", padding: "8vh 16px 40px" }}>
      <div dir={dir} onClick={e => e.stopPropagation()} style={{ width: "100%", maxWidth: view === "pay" || view === "wait" ? 520 : 440, background: "linear-gradient(180deg,#101828,#0b111c)", border: "1px solid rgba(245,158,11,.25)", borderRadius: 18, padding: "28px 26px 24px", boxShadow: "0 40px 90px -20px rgba(0,0,0,.85)", fontFamily: "'DM Sans', system-ui, sans-serif", color: "#e2e8f0" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
          <div>
            <div style={{ fontSize: 12, letterSpacing: 2, color: ACC, fontWeight: 800 }}>STEEL OPTIMIZER</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: "#f8fafc", marginTop: 6, lineHeight: 1.2 }}>{title}</div>
          </div>
          <button onClick={close} aria-label={L("close")} style={{ background: "none", border: "1px solid #243044", color: "#94a3b8", borderRadius: 10, width: 38, height: 38, cursor: "pointer", fontSize: 16, flex: "none" }}>✕</button>
        </div>
        {(view === "in" || view === "up") && <div style={{ fontSize: 15, color: "#94a3b8", marginTop: 8 }}>{L(PAYMENTS_ENABLED ? "lead" : "leadFree")}</div>}
        {msg && <div style={{ marginTop: 16, padding: "11px 14px", borderRadius: 10, background: "rgba(16,185,129,.12)", border: "1px solid rgba(16,185,129,.35)", color: "#6ee7b7", fontSize: 14.5 }}>{msg}</div>}
        {err && <div style={{ marginTop: 16, padding: "11px 14px", borderRadius: 10, background: "rgba(239,68,68,.12)", border: "1px solid rgba(239,68,68,.35)", color: "#fca5a5", fontSize: 14.5 }}>{err}</div>}

        {(view === "in" || view === "up" || view === "forgot") && (
          <form onSubmit={e => { e.preventDefault(); if (busy) return; view === "in" ? signIn() : view === "up" ? signUp() : forgot(); }}>
            <label style={LBL}>{L("email")}<input type="email" autoComplete="email" dir="ltr" value={email} onChange={e => setEmail(e.target.value)} style={INP} required /></label>
            {view !== "forgot" && <label style={LBL}>{L("password")}<input type="password" autoComplete={view === "up" ? "new-password" : "current-password"} dir="ltr" value={pass} onChange={e => setPass(e.target.value)} style={INP} required />
              {view === "up" && <span style={{ display: "block", marginTop: 6, fontSize: 12.5, color: "#64748b", fontWeight: 400 }}>{L("passHint")}</span>}</label>}
            <button type="submit" disabled={busy} style={BTN}>{busy ? "…" : view === "in" ? L("signIn") : view === "up" ? L("signUp") : L("sendLink")}</button>
          </form>)}
        {view === "in" && <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginTop: 16 }}>
          <button onClick={() => { setErr(""); setMsg(""); setView("up"); }} style={LINK}>{L("toUp")}</button>
          <button onClick={() => { setErr(""); setMsg(""); setView("forgot"); }} style={{ ...LINK, color: "#94a3b8" }}>{L("forgot")}</button></div>}
        {(view === "up" || view === "forgot") && <div style={{ marginTop: 16 }}><button onClick={() => { setErr(""); setMsg(""); setView("in"); }} style={LINK}>{L("toIn")}</button></div>}

        {view === "reset" && (
          <form onSubmit={e => { e.preventDefault(); if (!busy) savePass(); }}>
            <label style={LBL}>{L("password")}<input type="password" autoComplete="new-password" dir="ltr" value={pass} onChange={e => setPass(e.target.value)} style={INP} required />
              <span style={{ display: "block", marginTop: 6, fontSize: 12.5, color: "#64748b", fontWeight: 400 }}>{L("passHint")}</span></label>
            <button type="submit" disabled={busy} style={BTN}>{busy ? "…" : L("savePass")}</button>
          </form>)}

        {view === "check" && <div style={{ marginTop: 18, fontSize: 15, color: "#94a3b8" }}>⏳ {L("checking")}</div>}

        {view === "pay" && (<>
          <div style={{ fontSize: 15.5, color: "#cbd5e1", marginTop: 8 }}>{L("payLead")}</div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginTop: 18 }}>
            {MEMBERSHIP.checkoutMonthly && plan("m", MEMBERSHIP.priceMonthly, L("perMo"))}
            {MEMBERSHIP.checkoutYearly && plan("y", MEMBERSHIP.priceYearly, L("perYr"))}
          </div>
          <div style={{ marginTop: 16, display: "flex", gap: 8, flexWrap: "wrap" }}>
            {["Apple Pay", "Visa", "Mastercard", "Amex", "Google Pay"].map(b => <span key={b} style={{ padding: "5px 10px", borderRadius: 8, border: "1px solid #243044", background: "#0a0f18", fontSize: 12.5, color: "#cbd5e1", fontWeight: 700 }}>{b}</span>)}
          </div>
          <div style={{ marginTop: 12, fontSize: 13, color: "#64748b" }}>🔒 {L("secure")}</div>
        </>)}
        {view === "wait" && (<>
          <div style={{ marginTop: 18, fontSize: 16, color: "#cbd5e1" }}>⏳ {L("waiting")}</div>
          <button disabled={busy} onClick={checkNow} style={BTN}>{busy ? "…" : L("checkAgain")}</button>
          <div style={{ marginTop: 14 }}><button onClick={() => { setErr(""); stopPoll(); setView("pay"); }} style={LINK}>{L("backPlans")}</button></div>
        </>)}
        {s && s.user && s.user.email && view !== "in" && view !== "up" && view !== "forgot" && (
          <div style={{ marginTop: 20, paddingTop: 14, borderTop: "1px solid #1e293b", display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap", fontSize: 13, color: "#64748b" }}>
            <span dir="auto" style={{ wordBreak: "break-all" }}>{L("signedAs", { email: s.user.email })}</span>
            <button onClick={() => { soSignOut(); setView("in"); }} style={{ ...LINK, fontSize: 13, color: "#94a3b8" }}>{L("signOut")}</button>
          </div>)}
      </div>
    </div>
  );
}
/* account lines at the bottom of the language menu */
function SoAccountBox({ onDone }) {
  const { lang } = useLang();
  const mem = useSoMem();
  if (!MEMBERSHIP_ON) return null;
  const s = mem.session;
  const IT = { display: "block", width: "100%", padding: "9px 12px", borderRadius: 9, border: "none", background: "transparent", color: "#cbd5e1", fontSize: 14, textAlign: "start", cursor: "pointer", textDecoration: "none", fontFamily: "inherit", boxSizing: "border-box" };
  return (
    <div style={{ borderTop: "1px solid #1e293b", marginTop: 6, paddingTop: 6 }}>
      {s && s.user && s.user.email ? (<>
        <div dir="auto" style={{ padding: "6px 12px", fontSize: 12, color: "#64748b", wordBreak: "break-all" }}>{soMT(lang, "signedAs", { email: s.user.email })}</div>
        {PAYMENTS_ENABLED && mem.member && mem.member.portal_url && <a href={mem.member.portal_url} target="_blank" rel="noopener noreferrer" style={IT} onClick={onDone}>{soMT(lang, "manage")}</a>}
        <button onClick={() => { soSignOut(); onDone(); }} style={IT}>{soMT(lang, "signOut")}</button>
      </>) : <button onClick={() => { onDone(); soRequireMember(() => {}, false); }} style={{ ...IT, color: "#fbbf24", fontWeight: 700 }}>👤 {soMT(lang, "signIn")}</button>}
    </div>
  );
}
/* ╚══ end of MEMBERSHIP add-on ═══════════════════════════════════════════════╝ */

export default function App() {
  return (
    <LangProvider>
      <ErrorBoundary>
        <AppInner />
        <SoMemberHost />
      </ErrorBoundary>
    </LangProvider>
  );
}
