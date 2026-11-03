/**
 * يسحب أوامر العمل من TFMP بكل حالاتها ويحفظها في data/tfmp/work_orders/،
 * بنفس شكل data/tfmp/systems.json الحالي: { fetched_at, source, count, headers, rows }
 *
 * هذا سكربت مستقل تمامًا عن scripts/sync-tfmp-systems.js (سحب "CA Scorecard")
 * ولا يغيّره ولا يمسّ data/tfmp/systems.json — الاثنان يعملان جنبًا إلى جنب.
 *
 * ★ 2026-11-02 (بناءً على تجربة فعلية مع واجهة التصدير في TFMP، بما فيها
 * شاشة "Download Work Orders" نفسها — راجع لقطة الشاشة اللي بعتها): لقينا
 * إن فيه طريقتين تحميل مختلفتين في TFMP:
 *   1) رابط export.xlsx?lang=en بدون أي فلتر status خالص — ده بيرجع كل
 *      أوامر العمل بكل حالاتها في شيت "Work Orders" واحد، مش بس COMPLETED
 *      (جُرِّب فعليًا: 58,336 صف يشمل COMPLETED + EXPIRED + ASSIGNED +
 *      CANCELLED — لاحظنا كمان وجود حالة "CANCELLED" ما كناش ناخدها
 *      بالحسبان قبل كده). الطلب ده بسيط ومباشر ومفيهوش قيد الفلتر خالص.
 *   2) تفعيل "Include submission details" (= include_inspection=true) من
 *      نفس الشاشة — ده اللي بيفرض فلتر تاريخ/شهر إجباري (جرّبناه وطلع
 *      HTTP 422: "Exporting submission details needs a filter: choose an
 *      assignment month, or a completion date from which to start" من غيره).
 *      فبنستخدمه بس مع status=COMPLETED + completion_from/completion_to
 *      عشان نجيب شيتات التفاصيل (Inspection/CA/IR/Checklist) اللي أصلًا
 *      مالهاش معنى غير للأوامر المكتملة.
 *
 * فبقينا نسحب طلبين:
 *   • ALL      → export.xlsx?lang=en (بدون فلتر) لشيت "Work Orders" بكل
 *                الحالات — ده المصدر الوحيد لملف work-orders.json
 *   • DETAILS  → export.xlsx?status=COMPLETED&completion_from=…&
 *                completion_to=…&include_inspection=true&lang=en، ومنه
 *                بناخد بس شيتات Inspection/CA/IR/Checklist (شيت Work
 *                Orders بتاعه نفسه بنتجاهله لأن عندنا نسخة أشمل من ALL)
 *
 * ★★ 2026-11-02 (تحديث ثاني — إصلاح فشل فعلي في GitHub Actions):
 * طلب DETAILS لمدى زمني طويل دفعة واحدة (من يناير لليوم، ~29 ألف صف ×
 * 5 شيتات) بيخلي السيرفر ياخد وقت طويل جدًا في توليد الملف، لحد ما nginx
 * يقطع الاتصال بـ "504 Gateway Time-out" (حصل فعليًا في تشغيل حقيقي).
 * الحل: تقسيم طلب DETAILS شهرًا بشهر (completion_from/completion_to لكل
 * شهر على حدة) — كل طلب صغير وبيرجع بسرعة، وبعدين بندمج نتايج الشيتات
 * الأربعة من كل الشهور في نفس ملفات الخروج النهائية (مفيش تغيير في شكل
 * الملفات الناتجة ولا في dashboard.js، التقسيم داخلي في السكربت بس).
 * (المشروع بدأ فعليًا من يناير 2026، فمفيش داعي نرجع لتاريخ أبعد من كده
 * في completion_from — راجع WO_COMPLETED_FROM تحت لو احتجت تغييره.)
 *
 * ★★★ 2026-11-02 (تحديث ثالث — إصلاح فشل فعلي تاني في GitHub Actions بعد
 * التقسيم الشهري): شهر يناير (أول شهر في المشروع) رجّع ملف من غير شيت
 * "CA Details" خالص (مش شيت فاضي — الشيت مش موجود في الملف أصلًا)، لأنه
 * في هذا الشهر مفيش ولا صف واحد بتقييم CA. يعني TFMP بيحذف الشيت كله من
 * ملف الإكسل لو مفيهوش ولا صف، مش بس بيسيبه فاضي بعنوانه. قبل كده كنا
 * بنتعامل مع "allowEmpty" على إنه شيت موجود لكن صفوفه صفر — دلوقتي
 * بقى الشيت ممكن "مش موجود خالص" برضه حالة طبيعية (مش خطأ) ويتعامل
 * معاه بنفس المعنى: 0 صف لهذا الشهر لهذا الشيت بالذات. كمان أضفنا:
 *   • مهلة زمنية (timeout) 90 ثانية لكل طلب تنزيل — لو السيرفر علّق
 *     (hang) من غير ما يرجع حتى خطأ، بنوقف الطلب ونعيد المحاولة بدل ما
 *     نستنى لحد ما الـ Action نفسه ينتهي بالـ timeout العام بتاعه.
 *   • تأكيد إن رؤوس الأعمدة (headers) ثابتة بين كل شهور نفس الشيت —
 *     لو شهر رجّع رؤوس مختلفة عن باقي الشهور (يدل على تغيير في شكل
 *     التصدير من عند TFMP) بيوقف السكربت بخطأ واضح بدل ما يخلط بيانات
 *     غلط في أعمدة غلط بصمت.
 *
 * الخطوات: (1) تسجيل الدخول (2) تنزيل ملف ALL (3) تنزيل ملفات DETAILS
 * شهرًا بشهر ودمج شيتاتها الأربعة (شيت مفقود لشهر = 0 صف لهذا الشهر،
 * مش خطأ) (4) قراءة شيت Work Orders من ALL + الشيتات التفصيلية المدمجة
 * (5) فحص كل شيت على حدة (أعمدة إلزامية + ثبات الرؤوس بين الشهور)
 * (6) كتابة الملفات — ولا يُستبدل أي ملف إلا عند نجاح كل الفحوصات، فتبقى
 * آخر نسخة سليمة إن فشل التشغيل.
 *
 * GitHub Secrets (نفس أسرار سحب الـ FCA الحالي، لا تُكتب قيمها في أي ملف):
 *   TFMP_USERNAME , TFMP_PASSWORD
 * اختياري (GitHub Variables أو env):
 *   TFMP_BASE                 الافتراضي https://tfmp.meem-edgenta.tech
 *   WO_COMPLETED_FROM          الافتراضي 2026-01-01 (completion_from لأول شهر في DETAILS)
 *   WO_EXPORT_URL_ALL          لتجاوز بناء رابط ALL تمامًا عند الحاجة
 *   WO_EXPORT_URL_DETAILS      لتجاوز التقسيم الشهري كليًا وطلب رابط DETAILS
 *                              واحد يدويًا (للاختبار فقط — قد يرجع 504 لمدى طويل)
 * اختياري للاختبار فقط: XLSX_FILE (قراءة ملف محلي واحد بشيتاته الخمسة بدل
 * الطلبات كلها — مفيد لاختبار منطق قراءة الشيتات التفصيلية فقط)، OUT_DIR
 *
 * يحتاج مكتبة xlsx (SheetJS) — تُثبَّت في workflow بـ: npm install xlsx --no-save
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const BASE = (process.env.TFMP_BASE || 'https://tfmp.meem-edgenta.tech').replace(/\/$/, '');
const LOGIN_URL = BASE + '/api/v1/auth/login';
const TODAY = new Date().toISOString().slice(0, 10);
// المشروع بدأ من يناير 2026 — مفيش بيانات قبل كده، فده يغطي كل التاريخ فعليًا
const COMPLETED_FROM = process.env.WO_COMPLETED_FROM || '2026-01-01';
const EXPORT_URL_ALL =
  process.env.WO_EXPORT_URL_ALL ||
  BASE + '/api/v1/admin/work-orders/export.xlsx?lang=en';
// يُستخدم فقط لو WO_EXPORT_URL_DETAILS مضبوط يدويًا (تجاوز التقسيم الشهري)
const EXPORT_URL_DETAILS =
  process.env.WO_EXPORT_URL_DETAILS ||
  BASE + '/api/v1/admin/work-orders/export.xlsx?completion_from=' + COMPLETED_FROM + '&completion_to=' + TODAY + '&status=COMPLETED&include_inspection=true&lang=en';
const XLSX_FILE = process.env.XLSX_FILE || '';
const OUT_DIR = process.env.OUT_DIR || path.join(__dirname, '..', 'data', 'tfmp', 'work_orders');
const USER = process.env.TFMP_USERNAME || '';
const PASS = process.env.TFMP_PASSWORD || '';
let token = '';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function login() {
  if (!USER || !PASS) throw new Error('يجب ضبط TFMP_USERNAME و TFMP_PASSWORD في GitHub Secrets');
  const res = await fetch(LOGIN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify({ user_name: USER, password: PASS }),
  });
  if (!res.ok) throw new Error('فشل تسجيل الدخول: HTTP ' + res.status);
  const data = await res.json();
  if (!data.access_token) throw new Error('نجح تسجيل الدخول لكن لم يرد access_token');
  token = data.access_token;
  console.log('✔ تم تسجيل الدخول');
}

const FETCH_TIMEOUT_MS = Number(process.env.WO_FETCH_TIMEOUT_MS || 90000);

// مثل getText في سكربت systems، لكن يرجع Buffer ثنائي (الملف xlsx) بدل نص
async function getBinary(url, attempt = 1, relogins = 0) {
  let res;
  const ac = new AbortController();
  const timer = setTimeout(() => ac.abort(), FETCH_TIMEOUT_MS);
  try {
    try {
      res = await fetch(url, {
        headers: { Accept: 'application/octet-stream,*/*', Authorization: 'Bearer ' + token, Cookie: 'tfmp_session=' + token },
        signal: ac.signal,
      });
    } catch (e) {
      const isTimeout = e && e.name === 'AbortError';
      if (attempt < 3) { await sleep(1000 * attempt); return getBinary(url, attempt + 1, relogins); }
      throw isTimeout ? new Error('انتهت مهلة الطلب (' + Math.round(FETCH_TIMEOUT_MS / 1000) + 'ث) بعد 3 محاولات — السيرفر لم يرد خالص') : e;
    }
  } finally {
    clearTimeout(timer);
  }
  if (res.status === 401 || res.status === 403) {
    if (relogins >= 2) throw new Error('رفض الخادم التصريح (HTTP ' + res.status + ') حتى بعد إعادة تسجيل الدخول');
    await login();
    return getBinary(url, attempt, relogins + 1);
  }
  if ((res.status === 429 || res.status >= 500) && attempt < 3) {
    await sleep(2000 * attempt);
    return getBinary(url, attempt + 1, relogins);
  }
  if (!res.ok) {
    let bodyTxt = '';
    try { bodyTxt = (await res.clone().text()).slice(0, 500); } catch (e) {}
    throw new Error('HTTP ' + res.status + ' عند تنزيل ملف التصدير' + (bodyTxt ? ' — رد الخادم: ' + bodyTxt : ''));
  }
  const ct = String(res.headers.get('content-type') || '');
  const buf = Buffer.from(await res.arrayBuffer());
  if (ct.includes('text/html') || (buf.length < 2000 && buf.slice(0, 15).toString('utf8').trim().startsWith('<'))) {
    throw new Error('الرد صفحة HTML وليس ملف xlsx (غالبًا انتهت الجلسة أو الرابط غير صحيح)');
  }
  return buf;
}

// خلايا التاريخ تصل ككائن Date حقيقي (بفضل cellDates:true في القراءة)، أو كنص
// "2026-10-31 14:19:36.770000" في حالات نادرة — كلاهما يُطبَّع إلى "31/10/2026 14:19"
function normDate(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) {
    const p = (n) => String(n).padStart(2, '0');
    return p(v.getUTCDate()) + '/' + p(v.getUTCMonth() + 1) + '/' + v.getUTCFullYear() + ' ' + p(v.getUTCHours()) + ':' + p(v.getUTCMinutes());
  }
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
  return m ? m[3] + '/' + m[2] + '/' + m[1] + ' ' + m[4] + ':' + m[5] : String(v);
}

// وصف كل شيت: اسمه في الملف، الأعمدة الإلزامية للتحقق منه، وأعمدة التاريخ التي تُطبَّع
// ★★★★ 2026-11-03: كل شيت دلوقتي بيُقسَّم لملفات فصلية (quarter) بدل ملف
// واحد مجمّع يكبر للأبد — base هو اسم المجلد الفرعي، quarterField هو العمود
// اللي بيحدد ربع السنة (مع quarterFallbackField لو الأساسي فاضي)، نفس فكرة
// "Assignment Month" اللي الداشبورد بيستخدمها في فلتر الفترة بالفعل. راجع
// manifest.json وتعليق بناء الربع تحت لتفاصيل الآلية.
const SHEETS = [
  {
    sheet: 'Work Orders',
    file: 'work-orders.json',
    base: 'work-orders',
    source: 'tfmp-wo-export-work-orders',
    required: ['WO #', 'Status', 'Primary School', 'Ministry ID', 'Created', 'Completed'],
    dateCols: ['Created', 'Scheduled Start', 'Started', 'Completed'],
    quarterField: 'Assignment Month',
    quarterFallbackField: 'Created',
  },
  {
    sheet: 'Inspection Details',
    file: 'inspection-details.json',
    base: 'inspection-details',
    source: 'tfmp-wo-export-inspection-details',
    required: ['WO #', 'Section', 'Area Code', 'Area Name (EN)', 'Submitted At'],
    dateCols: ['Submitted At'],
    quarterField: 'Submitted At',
  },
  {
    sheet: 'CA Details',
    file: 'ca-details.json',
    base: 'ca-details',
    source: 'tfmp-wo-export-ca-details',
    required: ['WO #', 'Item Code', 'Item Name (EN)', 'Rating (1–5)', 'Submitted At'],
    dateCols: ['Submitted At'],
    quarterField: 'Submitted At',
  },
  {
    sheet: 'IR Details',
    file: 'ir-details.json',
    base: 'ir-details',
    source: 'tfmp-wo-export-ir-details',
    required: ['WO #', 'Section Code', 'Area Code', 'QA Outcome', 'Submitted At'],
    dateCols: ['Submitted At'],
    quarterField: 'Submitted At',
  },
  {
    sheet: 'Checklist Details',
    file: 'checklist-details.json',
    base: 'checklist-details',
    source: 'tfmp-wo-export-checklist-details',
    required: ['WO #', 'Checklist', 'Item Code', 'Answer', 'Submitted At'],
    dateCols: ['Submitted At'],
    quarterField: 'Submitted At',
  },
];

// أعمدة "Assignment Month" بشكل YYYY-MM، وأعمدة التاريخ التفصيلية بشكل
// dd/mm/yyyy (بعد normDate) — دالتين بسيطتين لاستخراج "YYYY-Qn" من كل شكل
function quarterKeyFromAssignmentMonth(v) {
  const m = String(v || '').match(/^(\d{4})-(\d{2})/);
  if (!m) return null;
  return m[1] + '-Q' + Math.ceil(Number(m[2]) / 3);
}
function quarterKeyFromDdMmYyyy(v) {
  const m = String(v || '').match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (!m) return null;
  return m[3] + '-Q' + Math.ceil(Number(m[2]) / 3);
}
// يقسّم صفوف شيت واحد على أرباع السنة — صف من غير تاريخ صالح (نادر) يروح
// لربع "unresolved" بدل ما يضيع أو يوقف التشغيل كله
function bucketByQuarter(cfg, table) {
  const fIdx = table.headers.indexOf(cfg.quarterField);
  const fbIdx = cfg.quarterFallbackField ? table.headers.indexOf(cfg.quarterFallbackField) : -1;
  const buckets = {};
  for (const r of table.rows) {
    let key = cfg.quarterField === 'Assignment Month'
      ? quarterKeyFromAssignmentMonth(r[fIdx])
      : quarterKeyFromDdMmYyyy(r[fIdx]);
    if (!key && fbIdx >= 0) key = quarterKeyFromDdMmYyyy(r[fbIdx]);
    if (!key) key = 'unresolved';
    (buckets[key] = buckets[key] || []).push(r);
  }
  return buckets;
}

function readSheet(wb, cfg, opts) {
  opts = opts || {};
  const ws = wb.Sheets[cfg.sheet];
  if (!ws) {
    // TFMP بيحذف الشيت كله من ملف الإكسل (مش بس بيسيبه فاضي) لو مفيش ولا صف
    // بيستحق هذا الشيت في المدى المطلوب (مثال فعلي: يناير 2026 من غير أي
    // صف CA Details). مع allowEmpty ده مش خطأ — يُعامل كـ 0 صف لهذا الطلب.
    if (opts.allowEmpty) return { headers: [], rows: [] };
    throw new Error('الشيت "' + cfg.sheet + '" غير موجود في الملف — تحقق من include_inspection=true في رابط التصدير');
  }
  // header:1 يعطي مصفوفة صفوف خام، أول صف هو رؤوس الأعمدة. raw:true يحافظ على خلايا
  // التاريخ ككائن Date حقيقي (بفضل cellDates:true عند القراءة) بدل نص منسّق قد يُفقد فيه اليوم/الشهر
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  if (!aoa.length) throw new Error('الشيت "' + cfg.sheet + '" فارغ تمامًا (بدون حتى صف العناوين)');
  const headers = aoa[0].map((h) => String(h == null ? '' : h).trim());
  const missing = cfg.required.filter((h) => !headers.includes(h));
  if (missing.length) throw new Error('الشيت "' + cfg.sheet + '": أعمدة ناقصة: ' + missing.join(' | '));

  const dateIdx = cfg.dateCols.map((h) => headers.indexOf(h)).filter((i) => i >= 0);
  const rows = aoa.slice(1).filter((r) => r.some((v) => v !== '' && v != null));
  rows.forEach((r) => { dateIdx.forEach((i) => { r[i] = normDate(r[i]); }); });

  // allowEmpty: بعض الحالات أو بعض الشهور ممكن يكون عددها صفر دلوقتي بشكل
  // طبيعي تمامًا (مفيش أوامر/صفوف في هذا الشهر أو هذه الحالة حاليًا) — ده مش خطأ
  if (!rows.length && !opts.allowEmpty) throw new Error('الشيت "' + cfg.sheet + '" بلا بيانات بعد التحقق');
  return { headers, rows };
}

// بدل ملف واحد لكل شيت، دلوقتي ملف لكل (شيت × ربع سنة) —
// data/tfmp/work_orders/<base>/<YYYY-Qn>.json — بنفس شكل { fetched_at,
// source, count, headers, rows } القديم تمامًا، فقط مقسّم. نفس فحص "مفيش
// انخفاض مفاجئ كبير" بس على مستوى كل ربع لوحده (مقارنة بنفس الربع في آخر
// سحب)، ونفس "لو مفيش تغيير متحفظش الملف" — يعني الأرباع المستقرة (شهور
// قديمة خلصت) هتفضل زي ما هي تلقائيًا من غير أي تدخل يدوي.
function prepareQuarter(cfg, quarterKey, rows, headers) {
  const dir = path.join(OUT_DIR, cfg.base);
  const file = path.join(dir, quarterKey + '.json');
  let prev = null;
  try { prev = JSON.parse(fs.readFileSync(file, 'utf8')); } catch (e) {}
  if (prev && prev.count && rows.length < prev.count * 0.9) {
    throw new Error(cfg.sheet + ' [' + quarterKey + ']: عدد الصفوف الجديد (' + rows.length + ') أقل بكثير من السابق (' + prev.count + ') — لن يُستبدل الملف');
  }
  if (prev && prev.count === rows.length && JSON.stringify(prev.headers) === JSON.stringify(headers) && JSON.stringify(prev.rows) === JSON.stringify(rows)) {
    return null; // مفيش تغيير — سيب الملف زي ما هو
  }
  return { dir, file, out: { fetched_at: new Date().toISOString(), source: cfg.source, count: rows.length, headers, rows } };
}

// manifest.json: قايمة الأرباع الموجودة فعليًا على القرص لكل شيت (بعد
// الكتابة)، عشان dashboard.js يعرف يجيب إيه من غير ما يخمّن أسماء ملفات —
// بيُبنى بمسح المجلدات مش من نتيجة هذا التشغيل بس، لأن أرباع قديمة مستقرة
// ممكن متتكتبش في هذا التشغيل بالذات لكن لازم تفضل في المانيفست
function buildManifest() {
  const manifest = { generated_at: new Date().toISOString(), sheets: {} };
  for (const cfg of SHEETS) {
    const dir = path.join(OUT_DIR, cfg.base);
    let files = [];
    try { files = fs.readdirSync(dir).filter((f) => f.endsWith('.json')); } catch (e) {}
    let total = 0;
    const quarters = files.map((f) => f.replace(/\.json$/, '')).sort();
    for (const f of files) {
      try { total += JSON.parse(fs.readFileSync(path.join(dir, f), 'utf8')).count || 0; } catch (e) {}
    }
    manifest.sheets[cfg.base] = { file: cfg.file, quarters, total_count: total, quarterField: cfg.quarterField };
  }
  return manifest;
}

const WO_SHEET_CFG = SHEETS[0]; // "Work Orders" — أول عنصر في SHEETS
const DETAIL_SHEETS = SHEETS.slice(1); // Inspection/CA/IR/Checklist — من طلبات DETAILS بس

async function fetchWorkbook(url, label) {
  if (!url.startsWith(BASE + '/')) throw new Error('رابط ' + label + ' يجب أن يكون على نفس نطاق TFMP (' + BASE + ') حتى لا تُرسَل بيانات الدخول لجهة أخرى');
  console.log('  ← ' + label + ' …');
  const buf = await getBinary(url);
  return XLSX.read(buf, { type: 'buffer', cellDates: true });
}

// يبني مدايات شهرية [from, to] (نص "YYYY-MM-DD") تغطي بالكامل المدى من
// fromStr إلى toStr، كل مدى محصور داخل حدود شهر تقويمي واحد (ومقصوص عند
// طرفي المدى الكلي). الهدف: كل طلب DETAILS يغطي شهر واحد بس فيرجع بسرعة
// (تجنب 504 Gateway Timeout من السيرفر عند توليد تصدير لمدى طويل دفعة واحدة)
function monthRanges(fromStr, toStr) {
  const ranges = [];
  const [fy, fm] = fromStr.split('-').map(Number);
  const fromDate = new Date(fromStr + 'T00:00:00Z');
  const toDate = new Date(toStr + 'T00:00:00Z');
  let cursor = new Date(Date.UTC(fy, fm - 1, 1));
  while (cursor <= toDate) {
    const y = cursor.getUTCFullYear();
    const m = cursor.getUTCMonth();
    const monthStart = new Date(Date.UTC(y, m, 1));
    const monthEndExclusive = new Date(Date.UTC(y, m + 1, 1));
    const rangeFromDate = monthStart < fromDate ? fromDate : monthStart;
    const rangeToDate = monthEndExclusive <= toDate ? new Date(monthEndExclusive.getTime() - 86400000) : toDate;
    ranges.push({ from: rangeFromDate.toISOString().slice(0, 10), to: rangeToDate.toISOString().slice(0, 10) });
    cursor = monthEndExclusive;
  }
  return ranges;
}

async function main() {
  let tables;

  if (XLSX_FILE) {
    console.log('(وضع الاختبار) قراءة الملف المحلي — بيتعامل معاملة ملف DETAILS الكامل (5 شيتات)');
    const buf = fs.readFileSync(XLSX_FILE);
    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
    tables = [{ cfg: WO_SHEET_CFG, table: readSheet(wb, WO_SHEET_CFG) }];
    for (const cfg of DETAIL_SHEETS) tables.push({ cfg, table: readSheet(wb, cfg, { allowEmpty: true }) });
  } else {
    console.log('… سحب أوامر العمل: طلب واحد بكل الحالات + طلبات تفاصيل التنفيذ للمكتملة (مقسّمة شهريًا)');
    await login();
    const allWb = await fetchWorkbook(EXPORT_URL_ALL, 'ALL (كل الحالات)');
    const woTable = readSheet(allWb, WO_SHEET_CFG);
    console.log('    ' + woTable.rows.length + ' صف (كل الحالات)');
    tables = [{ cfg: WO_SHEET_CFG, table: woTable }];

    if (process.env.WO_EXPORT_URL_DETAILS) {
      // تجاوز يدوي بالرابط كامل (بدون تقسيم شهري) — للاختبار فقط، قد يرجع 504 لمدى طويل
      const detailsWb = await fetchWorkbook(EXPORT_URL_DETAILS, 'DETAILS (COMPLETED + تفاصيل التنفيذ)');
      for (const cfg of DETAIL_SHEETS) tables.push({ cfg, table: readSheet(detailsWb, cfg, { allowEmpty: true }) });
    } else {
      const ranges = monthRanges(COMPLETED_FROM, TODAY);
      console.log('    تقسيم DETAILS على ' + ranges.length + ' شهر لتجنب 504 Gateway Timeout من السيرفر');
      const merged = {}; // cfg.sheet -> { headers, rows, headerSourceLabel }
      for (const range of ranges) {
        const url = BASE + '/api/v1/admin/work-orders/export.xlsx?completion_from=' + range.from + '&completion_to=' + range.to + '&status=COMPLETED&include_inspection=true&lang=en';
        const label = 'DETAILS ' + range.from + ' → ' + range.to;
        const wb = await fetchWorkbook(url, label);
        for (const cfg of DETAIL_SHEETS) {
          const t = readSheet(wb, cfg, { allowEmpty: true });
          if (!merged[cfg.sheet]) merged[cfg.sheet] = { headers: [], rows: [], headerSourceLabel: '' };
          if (t.headers.length) {
            if (!merged[cfg.sheet].headers.length) {
              merged[cfg.sheet].headers = t.headers;
              merged[cfg.sheet].headerSourceLabel = label;
            } else if (JSON.stringify(merged[cfg.sheet].headers) !== JSON.stringify(t.headers)) {
              // رؤوس أعمدة مختلفة بين شهرين لنفس الشيت = تغيير في شكل التصدير
              // من عند TFMP مش متوقّع — نوقف بدل ما نخلط أعمدة مختلفة بصمت
              throw new Error(
                'الشيت "' + cfg.sheet + '": رؤوس الأعمدة في ' + label + ' مختلفة عنها في ' + merged[cfg.sheet].headerSourceLabel +
                ' — (' + t.headers.join(' | ') + ') مقابل (' + merged[cfg.sheet].headers.join(' | ') + ')'
              );
            }
          }
          merged[cfg.sheet].rows.push(...t.rows);
        }
        await sleep(300); // مجرد إراحة بسيطة للسيرفر بين طلب وطلب
      }
      for (const cfg of DETAIL_SHEETS) tables.push({ cfg, table: merged[cfg.sheet] || { headers: [], rows: [] } });
    }
  }

  const ready = [];
  for (const { cfg, table } of tables) {
    const buckets = bucketByQuarter(cfg, table);
    const unresolvedCount = (buckets.unresolved || []).length;
    if (unresolvedCount) console.log('  تنبيه: ' + unresolvedCount + ' صف في "' + cfg.sheet + '" بدون تاريخ صالح — في ربع "unresolved"');
    for (const quarterKey of Object.keys(buckets).sort()) {
      const r = prepareQuarter(cfg, quarterKey, buckets[quarterKey], table.headers);
      if (r) ready.push(r);
    }
  }
  fs.mkdirSync(OUT_DIR, { recursive: true });
  for (const { dir, file, out } of ready) {
    fs.mkdirSync(dir, { recursive: true });
    const tmp = file + '.tmp';
    fs.writeFileSync(tmp, JSON.stringify(out));
    fs.renameSync(tmp, file);
    console.log('✔ حُفظ ' + out.count + ' صفًا و' + out.headers.length + ' عمودًا في', file);
  }
  if (!ready.length) console.log('لا تغيير في أي شيت — لم يُحفظ شيء');
  // manifest.json يُعاد بناؤه دايمًا (حتى لو مفيش تغيير في البيانات) عشان
  // يفضل يعكس كل الأرباع الموجودة فعليًا على القرص
  const manifest = buildManifest();
  fs.writeFileSync(path.join(OUT_DIR, 'manifest.json'), JSON.stringify(manifest, null, 2));
  console.log('✔ تحديث manifest.json');

  // أول تشغيل بعد التحديث للتقسيم الفصلي: امسح الملفات المجمّعة القديمة
  // (work-orders.json القديم... إلخ) لو موجودة — بقت كل بياناتها موجودة في
  // الملفات الفصلية الجديدة، وmanifest.json بقى المرجع الوحيد اللي
  // dashboard.js بيعتمد عليه (مع fallback للملف القديم لو manifest.json
  // مش موجود لأي سبب — راجع _woEnsureLoaded في dashboard.js)
  for (const cfg of SHEETS) {
    const oldFile = path.join(OUT_DIR, cfg.file);
    if (fs.existsSync(oldFile)) {
      fs.unlinkSync(oldFile);
      console.log('🗑 حُذف الملف المجمّع القديم:', oldFile);
    }
  }
}

main().catch((e) => { console.error('❌ فشل السحب:', e.message); process.exit(1); });
