/**
 * يسحب ملف "CA Scorecard — all — wide" (نفس ملف CSV الذي يُحمَّل يدويًا من TFMP)
 * ويحفظه في المستودع كـ data/tfmp/systems.json ليقرأه تبويبا "الأنظمة الرئيسية FCA"
 * و"الأنظمة التفصيلية FCA" مباشرة، بدل نسخ الـ CSV يدويًا إلى شيت "المدارس_والأنظمة".
 *
 * الخطوات: (1) تسجيل الدخول  (2) تنزيل الـ CSV من رابط التصدير  (3) فحص سلامته
 * (4) كتابة الملف — ولا يُستبدل الملف إلا عند نجاح كل الفحوصات، فتبقى آخر نسخة سليمة إن فشل التشغيل.
 *
 * GitHub Secrets : TFMP_USERNAME , TFMP_PASSWORD   (لا تُكتب قيمهما في أي ملف)
 * رابط التصدير الافتراضي: /api/v1/admin/dashboards/ca-scorecard/export.csv?sort=score_asc (يمكن تغييره بمتغير TFMP_EXPORT_URL اختياريًا)
 * اختياري للاختبار فقط: CSV_FILE (قراءة ملف محلي بدل التنزيل) ، OUT_FILE ، TFMP_BASE
 */
const fs = require('fs');
const path = require('path');

const BASE = (process.env.TFMP_BASE || 'https://tfmp.meem-edgenta.tech').replace(/\/$/, '');
const LOGIN_URL = BASE + '/api/v1/auth/login';
const EXPORT_URL = process.env.TFMP_EXPORT_URL || (BASE + '/api/v1/admin/dashboards/ca-scorecard/export.csv?sort=score_asc');
const CSV_FILE = process.env.CSV_FILE || '';
const OUT_FILE = process.env.OUT_FILE || path.join(__dirname, '..', 'data', 'tfmp', 'systems.json');
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

async function getText(url, attempt = 1, relogins = 0) {
  let res;
  try {
    res = await fetch(url, {
      headers: { Accept: 'text/csv,*/*', Authorization: 'Bearer ' + token, Cookie: 'tfmp_session=' + token },
    });
  } catch (e) {
    if (attempt < 3) { await sleep(1000 * attempt); return getText(url, attempt + 1, relogins); }
    throw e;
  }
  if (res.status === 401 || res.status === 403) {
    if (relogins >= 2) throw new Error('رفض الخادم التصريح (HTTP ' + res.status + ') حتى بعد إعادة تسجيل الدخول');
    await login();
    return getText(url, attempt, relogins + 1);
  }
  if ((res.status === 429 || res.status >= 500) && attempt < 3) {
    await sleep(2000 * attempt);
    return getText(url, attempt + 1, relogins);
  }
  if (!res.ok) throw new Error('HTTP ' + res.status + ' عند تنزيل ملف التصدير');
  return res.text();
}

// قارئ CSV صغير يدعم الاقتباس والأسطر داخل الحقول و BOM
function parseCsv(text) {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const rows = [];
  let row = [], field = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false; }
      else field += c;
    } else if (c === '"') inQ = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field); field = '';
      if (row.length > 1 || row[0] !== '') rows.push(row);
      row = [];
    } else field += c;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

// "2026-05-23 11:47:23" أو "2026-05-23T11:47:23.000Z" → "23/05/2026 11:47" (الصيغة التي تفهمها لوحة البيانات)
function normDate(v) {
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
  return m ? m[3] + '/' + m[2] + '/' + m[1] + ' ' + m[4] + ':' + m[5] : v;
}

async function main() {
  let text;
  if (CSV_FILE) {
    text = fs.readFileSync(CSV_FILE, 'utf8');
    console.log('(وضع الاختبار) قراءة الملف المحلي');
  } else {
    if (!EXPORT_URL.startsWith(BASE + '/')) throw new Error('رابط التصدير يجب أن يكون على نفس نطاق TFMP (' + BASE + ') حتى لا تُرسَل بيانات الدخول لجهة أخرى');
    await login();
    text = await getText(EXPORT_URL);
  }
  if (/^\s*</.test(text)) throw new Error('الرد صفحة HTML وليس CSV (غالبًا انتهت الجلسة أو الرابط غير صحيح)');

  const table = parseCsv(text);
  const headers = table[0].map((h) => h.trim());
  const rows = table.slice(1).filter((r) => r.length === headers.length);
  const skipped = table.length - 1 - rows.length;

  // فحوصات السلامة
  const need = ['Work Order #', 'School Name', 'Ministry ID', 'Site', 'Completed At', 'Submission Total Score %', 'Overall Condition Band'];
  const missing = need.filter((h) => !headers.includes(h));
  if (missing.length) throw new Error('أعمدة ناقصة في الملف: ' + missing.join(' | '));
  if (!headers.some((h) => h.endsWith(' — Rating'))) throw new Error('لا يوجد أي عمود ينتهي بـ " — Rating" (تغيّر شكل الملف؟)');
  if (!rows.length) throw new Error('الملف بلا بيانات');
  if (skipped > rows.length * 0.02) throw new Error('عدد كبير من الصفوف غير مكتملة (' + skipped + ') — يُرجَّح أن الملف مقطوع');

  const iDate = headers.indexOf('Completed At');
  rows.forEach((r) => { r[iDate] = normDate(r[iDate]); });

  let prev = null;
  try { prev = JSON.parse(fs.readFileSync(OUT_FILE, 'utf8')); } catch (e) {}
  if (prev && prev.count && rows.length < prev.count * 0.9) {
    throw new Error('عدد الصفوف الجديد (' + rows.length + ') أقل بكثير من السابق (' + prev.count + ') — لن يُستبدل الملف');
  }
  if (prev && prev.count === rows.length && JSON.stringify(prev.rows) === JSON.stringify(rows)) {
    console.log('لا تغيير في البيانات منذ آخر سحب — لن يُعاد حفظ الملف');
    return;
  }

  const out = { fetched_at: new Date().toISOString(), source: 'tfmp-export-csv', count: rows.length, headers, rows };
  fs.mkdirSync(path.dirname(OUT_FILE), { recursive: true });
  const tmp = OUT_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(out));
  fs.renameSync(tmp, OUT_FILE);
  console.log('✔ حُفظ ' + rows.length + ' صفًا و' + headers.length + ' عمودًا في', OUT_FILE, skipped ? '(تجاهل ' + skipped + ' صف غير مكتمل)' : '');
}

main().catch((e) => { console.error('❌ فشل السحب:', e.message); process.exit(1); });
