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
 * فبقينا نسحب طلبين بس:
 *   • ALL      → export.xlsx?lang=en (بدون فلتر) لشيت "Work Orders" بكل
 *                الحالات — ده المصدر الوحيد لملف work-orders.json
 *   • DETAILS  → export.xlsx?status=COMPLETED&completion_from=…&
 *                completion_to=…&include_inspection=true&lang=en، ومنه
 *                بناخد بس شيتات Inspection/CA/IR/Checklist (شيت Work
 *                Orders بتاعه نفسه بنتجاهله لأن عندنا نسخة أشمل من ALL)
 * (المشروع بدأ فعليًا من يناير 2026، فمفيش داعي نرجع لتاريخ أبعد من كده
 * في completion_from — راجع WO_COMPLETED_FROM تحت لو احتجت تغييره.)
 *
 * الخطوات: (1) تسجيل الدخول (2) تنزيل الملفين (3) قراءة شيت Work Orders من
 * ALL + الشيتات التفصيلية الأربعة من DETAILS (4) فحص كل شيت على حدة
 * (5) كتابة الملفات — ولا يُستبدل أي ملف إلا عند نجاح كل الفحوصات، فتبقى
 * آخر نسخة سليمة إن فشل التشغيل.
 *
 * GitHub Secrets (نفس أسرار سحب الـ FCA الحالي، لا تُكتب قيمها في أي ملف):
 *   TFMP_USERNAME , TFMP_PASSWORD
 * اختياري (GitHub Variables أو env):
 *   TFMP_BASE                 الافتراضي https://tfmp.meem-edgenta.tech
 *   WO_COMPLETED_FROM          الافتراضي 2026-01-01 (completion_from لطلب DETAILS)
 *   WO_EXPORT_URL_ALL          لتجاوز بناء رابط ALL تمامًا عند الحاجة
 *   WO_EXPORT_URL_DETAILS      لتجاوز بناء رابط DETAILS تمامًا عند الحاجة
 * اختياري للاختبار فقط: XLSX_FILE (قراءة ملف محلي واحد بشيتاته الخمسة بدل
 * الطلبين — مفيد لاختبار منطق قراءة الشيتات التفصيلية فقط)، OUT_DIR
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

function readSheet(wb, cfg, opts) {
  opts = opts || {};
  const ws = wb.Sheets[cfg.sheet];
  if (!ws) throw new Error('الشيت "' + cfg.sheet + '" غير موجود في الملف — تحقق من include_inspection=true في رابط التصدير');
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

  // allowEmpty: بعض الحالات (مثلًا "IN PROGRESS") ممكن يكون عددها صفر دلوقتي
  // بشكل طبيعي تمامًا (مفيش أوامر في هذه الحالة حاليًا) — ده مش خطأ
  if (!rows.length && !opts.allowEmpty) throw new Error('الشيت "' + cfg.sheet + '" بلا بيانات بعد التحقق');
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

const WO_SHEET_CFG = SHEETS[0]; // "Work Orders" — أول عنصر في SHEETS
const DETAIL_SHEETS = SHEETS.slice(1); // Inspection/CA/IR/Checklist — من طلب DETAILS بس

async function fetchWorkbook(url, label) {
  if (!url.startsWith(BASE + '/')) throw new Error('رابط ' + label + ' يجب أن يكون على نفس نطاق TFMP (' + BASE + ') حتى لا تُرسَل بيانات الدخول لجهة أخرى');
  console.log('  ← ' + label + ' …');
  const buf = await getBinary(url);
  return XLSX.read(buf, { type: 'buffer', cellDates: true });
}

async function main() {
  let woTable, detailsWb;

  if (XLSX_FILE) {
    console.log('(وضع الاختبار) قراءة الملف المحلي — بيتعامل معاملة ملف DETAILS الكامل (5 شيتات)');
    const buf = fs.readFileSync(XLSX_FILE);
    detailsWb = XLSX.read(buf, { type: 'buffer', cellDates: true });
    woTable = readSheet(detailsWb, WO_SHEET_CFG);
  } else {
    console.log('… سحب أوامر العمل: طلب واحد بكل الحالات + طلب تفاصيل التنفيذ للمكتملة');
    await login();
    const allWb = await fetchWorkbook(EXPORT_URL_ALL, 'ALL (كل الحالات)');
    woTable = readSheet(allWb, WO_SHEET_CFG);
    console.log('    ' + woTable.rows.length + ' صف (كل الحالات)');
    detailsWb = await fetchWorkbook(EXPORT_URL_DETAILS, 'DETAILS (COMPLETED + تفاصيل التنفيذ)');
  }

  const tables = [{ cfg: WO_SHEET_CFG, table: woTable }];
  for (const cfg of DETAIL_SHEETS) tables.push({ cfg, table: readSheet(detailsWb, cfg) });

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
