/**
 * يسحب تصدير "Work Orders" من TFMP (ملف xlsx بخمسة شيتات: Work Orders, Inspection
 * Details, CA Details, IR Details, Checklist Details) ويحفظ كل شيت كملف JSON منفصل
 * في data/tfmp/work_orders/، بنفس شكل data/tfmp/systems.json الحالي:
 *   { fetched_at, source, count, headers, rows }
 *
 * هذا سكربت مستقل تمامًا عن scripts/sync-tfmp-systems.js (سحب "CA Scorecard")
 * ولا يغيّره ولا يمسّ data/tfmp/systems.json — الاثنان يعملان جنبًا إلى جنب.
 *
 * ★ 2026-11-01 (بناءً على طلب صريح): شِلنا فلتر status=COMPLETED وفلتر
 * التواريخ completion_from/completion_to خالص — الرابط بقى بدون أي فلتر
 * حالة أو تاريخ (بالظبط زي الرابط اللي اتبعت)، فبيرجع كل أوامر العمل
 * بكل حالاتها (ASSIGNED, IN PROGRESS, EXPIRED, COMPLETED, ...) مش
 * المكتملة بس. ده عشان الداشبورد يقدر يحسب "المخطط" (أوامر لسه مفتوحة:
 * Assigned/In Progress/Expired) مقابل "المكتمل" (Completed) لكل مدينة،
 * ونسبة إنجاز حقيقية = المكتمل ÷ المخطط.
 *
 * الخطوات: (1) تسجيل الدخول (2) تنزيل ملف xlsx (3) قراءة الشيتات الخمسة
 * (4) فحص كل شيت على حدة (5) كتابة الملفات — ولا يُستبدل أي ملف إلا عند
 * نجاح كل الفحوصات، فتبقى آخر نسخة سليمة إن فشل التشغيل.
 *
 * GitHub Secrets (نفس أسرار سحب الـ FCA الحالي، لا تُكتب قيمها في أي ملف):
 *   TFMP_USERNAME , TFMP_PASSWORD
 * اختياري (GitHub Variables أو env):
 *   TFMP_BASE            الافتراضي https://tfmp.meem-edgenta.tech
 *   WO_EXPORT_URL         لتجاوز بناء الرابط تمامًا عند الحاجة (مثلاً
 *                         لإضافة فلتر تاريخ/حالة لاحقًا لو احتجنا)
 * اختياري للاختبار فقط: XLSX_FILE (قراءة ملف محلي بدل التنزيل)، OUT_DIR
 *
 * يحتاج مكتبة xlsx (SheetJS) — تُثبَّت في workflow بـ: npm install xlsx --no-save
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const BASE = (process.env.TFMP_BASE || 'https://tfmp.meem-edgenta.tech').replace(/\/$/, '');
const LOGIN_URL = BASE + '/api/v1/auth/login';
// بدون فلتر status أو تاريخ — بيرجع كل أوامر العمل بكل حالاتها (راجع الملاحظة فوق)
const EXPORT_URL =
  process.env.WO_EXPORT_URL ||
  BASE + '/api/v1/admin/work-orders/export.xlsx?include_inspection=true&lang=en';
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

// مثل getText في سكربت systems، لكن يرجع Buffer ثنائي (الملف xlsx) بدل نص
async function getBinary(url, attempt = 1, relogins = 0) {
  let res;
  try {
    res = await fetch(url, {
      headers: { Accept: 'application/octet-stream,*/*', Authorization: 'Bearer ' + token, Cookie: 'tfmp_session=' + token },
    });
  } catch (e) {
    if (attempt < 3) { await sleep(1000 * attempt); return getBinary(url, attempt + 1, relogins); }
    throw e;
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
const SHEETS = [
  {
    sheet: 'Work Orders',
    file: 'work-orders.json',
    source: 'tfmp-wo-export-work-orders',
    required: ['WO #', 'Status', 'Primary School', 'Ministry ID', 'Created', 'Completed'],
    dateCols: ['Created', 'Scheduled Start', 'Started', 'Completed'],
  },
  {
    sheet: 'Inspection Details',
    file: 'inspection-details.json',
    source: 'tfmp-wo-export-inspection-details',
    required: ['WO #', 'Section', 'Area Code', 'Area Name (EN)', 'Submitted At'],
    dateCols: ['Submitted At'],
  },
  {
    sheet: 'CA Details',
    file: 'ca-details.json',
    source: 'tfmp-wo-export-ca-details',
    required: ['WO #', 'Item Code', 'Item Name (EN)', 'Rating (1–5)', 'Submitted At'],
    dateCols: ['Submitted At'],
  },
  {
    sheet: 'IR Details',
    file: 'ir-details.json',
    source: 'tfmp-wo-export-ir-details',
    required: ['WO #', 'Section Code', 'Area Code', 'QA Outcome', 'Submitted At'],
    dateCols: ['Submitted At'],
  },
  {
    sheet: 'Checklist Details',
    file: 'checklist-details.json',
    source: 'tfmp-wo-export-checklist-details',
    required: ['WO #', 'Checklist', 'Item Code', 'Answer', 'Submitted At'],
    dateCols: ['Submitted At'],
  },
];

function readSheet(wb, cfg) {
  const ws = wb.Sheets[cfg.sheet];
  if (!ws) throw new Error('الشيت "' + cfg.sheet + '" غير موجود في الملف — تحقق من include_inspection=true في رابط التصدير');
  // header:1 يعطي مصفوفة صفوف خام، أول صف هو رؤوس الأعمدة. raw:true يحافظ على خلايا
  // التاريخ ككائن Date حقيقي (بفضل cellDates:true عند القراءة) بدل نص منسّق قد يُفقد فيه اليوم/الشهر
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  if (!aoa.length) throw new Error('الشيت "' + cfg.sheet + '" فارغ');
  const headers = aoa[0].map((h) => String(h == null ? '' : h).trim());
  const missing = cfg.required.filter((h) => !headers.includes(h));
  if (missing.length) throw new Error('الشيت "' + cfg.sheet + '": أعمدة ناقصة: ' + missing.join(' | '));

  const dateIdx = cfg.dateCols.map((h) => headers.indexOf(h)).filter((i) => i >= 0);
  const rows = aoa.slice(1).filter((r) => r.some((v) => v !== '' && v != null));
  rows.forEach((r) => { dateIdx.forEach((i) => { r[i] = normDate(r[i]); }); });

  if (!rows.length) throw new Error('الشيت "' + cfg.sheet + '" بلا بيانات بعد التحقق');
  return { headers, rows };
}

function prepare(cfg, table) {
  const file = path.join(OUT_DIR, cfg.file);
  let prev = null;
  try { prev = JSON.parse(fs.readFileSync(file, 'utf8')); } catch (e) {}
  // فحص سلامة عام: انخفاض كبير مفاجئ في عدد الصفوف عن آخر سحب ناجح مريب (ممكن يبقى خطأ في الرابط أو التصريح)
  if (prev && prev.count && table.rows.length < prev.count * 0.9) {
    throw new Error(cfg.sheet + ': عدد الصفوف الجديد (' + table.rows.length + ') أقل بكثير من السابق (' + prev.count + ') — لن يُستبدل الملف');
  }
  if (prev && prev.count === table.rows.length && JSON.stringify(prev.headers) === JSON.stringify(table.headers) && JSON.stringify(prev.rows) === JSON.stringify(table.rows)) {
    console.log('لا تغيير في "' + cfg.sheet + '" منذ آخر سحب — لن يُعاد حفظ الملف');
    return null;
  }
  return { file, out: { fetched_at: new Date().toISOString(), source: cfg.source, count: table.rows.length, headers: table.headers, rows: table.rows } };
}

async function main() {
  let buf;
  if (XLSX_FILE) {
    buf = fs.readFileSync(XLSX_FILE);
    console.log('(وضع الاختبار) قراءة الملف المحلي');
  } else {
    if (!EXPORT_URL.startsWith(BASE + '/')) throw new Error('رابط التصدير يجب أن يكون على نفس نطاق TFMP (' + BASE + ') حتى لا تُرسَل بيانات الدخول لجهة أخرى');
    console.log('… سحب كل أوامر العمل بكل حالاتها (بدون فلتر status أو تاريخ)');
    await login();
    buf = await getBinary(EXPORT_URL);
  }

  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });

  const tables = SHEETS.map((cfg) => ({ cfg, table: readSheet(wb, cfg) }));

  const ready = tables.map(({ cfg, table }) => prepare(cfg, table)).filter(Boolean);
  fs.mkdirSync(OUT_DIR, { recursive: true });
  for (const { file, out } of ready) {
    const tmp = file + '.tmp';
    fs.writeFileSync(tmp, JSON.stringify(out));
    fs.renameSync(tmp, file);
    console.log('✔ حُفظ ' + out.count + ' صفًا و' + out.headers.length + ' عمودًا في', file);
  }
  if (!ready.length) console.log('لا تغيير في أي شيت — لم يُحفظ شيء');
}

main().catch((e) => { console.error('❌ فشل السحب:', e.message); process.exit(1); });
