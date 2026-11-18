/**
 * ثراء الرمال — لوحة طلبات الحاويات على Google Sheets
 * ------------------------------------------------------
 * 1) افتح جدول قوقل جديد ← الإضافات (Extensions) ← Apps Script
 * 2) من الإعدادات (⚙️ Project Settings) تأكد إن V8 مفعّل
 * 3) احذف أي كود موجود والصق هذا الملف كاملاً، ثم احفظ
 * 4) شغّل الدالة setup مرة وحدة (توافق على الصلاحيات)
 * 5) انشر: Deploy ← New deployment ← Web app
 *    Execute as: Me   |   Who has access: Anyone
 * 6) انسخ رابط النشر (ينتهي بـ /exec) وحطه في ENDPOINT داخل index.html على GitHub
 *
 * صفحة الطلب مضمّنة تحت (PAGE_B64) مشفرة base64 — لا تعدلها يدوياً.
 * الرابط اللي ترسله للعملاء: https://waseemanalytic-dotcom.github.io/thara-alrimal/
 */

const SHEET_ORDERS  = 'الطلبات';
const SHEET_PRICES  = 'الأسعار';
const SHEET_SUMMARY = 'الملخص';

const STATUSES = ['جديد', 'تم التأكيد', 'عند العميل', 'تم السحب', 'ملغي'];
const SERVICES = {NEW: 'حاوية جديدة', SWAP: 'استبدال حاوية', PICKUP: 'سحب حاوية'};
const CITIES = ['الخبر', 'الظهران', 'الدمام', 'سيهات', 'القطيف', 'مدينة أخرى'];

// كل الأسعار غير شاملة ضريبة القيمة المضافة
const VAT = 0.15;

// الأسعار الأولية (تُنسخ لورقة "الأسعار" عند setup، وبعدها تعدّل من الورقة مباشرة وتظهر في الصفحة)
const SEED_PRICES = [
  ['NEW',       'حاوية جديدة 20 ياردة (شاملة التوصيل والسحب)', 400, 'ريال قبل الضريبة'],
  ['SWAP',      'استبدال حاوية (نسحب الممتلئة ونحط فاضية)',     400, 'ريال قبل الضريبة'],
  ['PICKUP',    'سحب حاوية',                                     0,   '0 = مشمول في سعر الحاوية'],
  ['EXTRA_DAY', 'سعر اليوم الإضافي',                              30,  'ريال قبل الضريبة'],
  ['DAYS',      'عدد الأيام المشمولة',                            10,  'يوم']
];

/* ================= الإعداد لمرة وحدة ================= */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale('ar');
  ss.setSpreadsheetTimeZone('Asia/Riyadh');

  // الطلبات
  const o = getOrCreate_(ss, SHEET_ORDERS);
  const oh = ['رقم الطلب','وقت الطلب','الخدمة','العدد','رقم الحاوية عند العميل','الاسم','الجوال','نوع العميل',
              'المدينة','الحي','رابط الموقع','وصف المكان','الموعد المطلوب','الدفع',
              'المبلغ قبل الضريبة','الضريبة 15%','الإجمالي شامل الضريبة','ملاحظات العميل',
              'الحالة','المسؤول','تاريخ التسليم الفعلي','رقم الحاوية المسلّمة','ملاحظات داخلية'];
  o.getRange(1,1,1,oh.length).setValues([oh]).setFontWeight('bold')
   .setBackground('#0F2439').setFontColor('#FFC107').setWrap(true);
  o.setFrozenRows(1);
  o.setRightToLeft(true);
  o.setColumnWidths(1, oh.length, 120);
  o.setColumnWidth(11, 220);
  o.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  o.getRange('M2:M').setNumberFormat('yyyy-mm-dd');
  o.getRange('U2:U').setNumberFormat('yyyy-mm-dd');
  o.getRange('U2:U').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
  o.getRange('S2:S').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  const colors = {'جديد':'#FFF3C4','تم التأكيد':'#DCEBFF','عند العميل':'#FFE0C2','تم السحب':'#D5F2E3','ملغي':'#F4D6D6'};
  o.setConditionalFormatRules(Object.keys(colors).map(s =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colors[s])
      .setRanges([o.getRange('S2:S')]).build()));

  // الأسعار
  const p = getOrCreate_(ss, SHEET_PRICES);
  if (p.getLastRow() < 2) {
    const ph = ['الرمز','البند','القيمة','ملاحظة'];
    p.getRange(1,1,1,ph.length).setValues([ph]).setFontWeight('bold').setBackground('#FFC107').setFontColor('#0F2439');
    p.getRange(2,1,SEED_PRICES.length,ph.length).setValues(SEED_PRICES);
    p.setFrozenRows(1); p.setRightToLeft(true);
    p.setColumnWidth(2, 330); p.setColumnWidth(4, 220);
    p.getRange('A:A').setFontColor('#999999');
    p.getRange('C2:C').setBackground('#FFF8E1').setFontWeight('bold');
    p.getRange(SEED_PRICES.length + 3, 2).setValue('عدّل الأرقام في عمود "القيمة" فقط، وتتحدث الصفحة تلقائياً. لا تغيّر عمود الرمز.')
      .setFontColor('#5D6B78');
  }

  // الملخص
  const s = getOrCreate_(ss, SHEET_SUMMARY);
  s.clear(); s.setRightToLeft(true);
  const q = "'" + SHEET_ORDERS + "'";
  const pr = "'" + SHEET_PRICES + "'";
  s.getRange('A1').setValue('ملخص طلبات ثراء الرمال').setFontSize(16).setFontWeight('bold');
  const rows = [
    ['عدد الطلبات', `=COUNTA(${q}!A2:A)`],
    ['طلبات جديدة تنتظر التأكيد', `=COUNTIF(${q}!S2:S,"جديد")`],
    ['تم التأكيد (بانتظار التوصيل)', `=COUNTIF(${q}!S2:S,"تم التأكيد")`],
    ['طلبات حاوياتها عند العملاء', `=COUNTIF(${q}!S2:S,"عند العميل")`],
    ['تم السحب', `=COUNTIF(${q}!S2:S,"تم السحب")`],
    ['ملغية', `=COUNTIF(${q}!S2:S,"ملغي")`],
    ['مبيعات مؤكدة قبل الضريبة (ريال)', `=SUM(${q}!O2:O)-SUMIF(${q}!S2:S,"ملغي",${q}!O2:O)-SUMIF(${q}!S2:S,"جديد",${q}!O2:O)`],
    ['مبيعات مؤكدة شامل الضريبة (ريال)', `=SUM(${q}!Q2:Q)-SUMIF(${q}!S2:S,"ملغي",${q}!Q2:Q)-SUMIF(${q}!S2:S,"جديد",${q}!Q2:Q)`]
  ];
  s.getRange(3,1,rows.length,2).setValues(rows);
  s.getRange(3,1,rows.length,1).setFontWeight('bold');
  s.setColumnWidth(1, 260);

  s.getRange('D2').setValue('الطلبات حسب المدينة').setFontWeight('bold');
  s.getRange('D3').setFormula(`=IFERROR(QUERY(${q}!I2:Q,"select I, count(I), sum(Q) where I is not null group by I order by count(I) desc label count(I) 'العدد', sum(Q) 'القيمة شامل الضريبة'",0),"لا توجد طلبات بعد")`);
  s.getRange('H2').setValue('الطلبات حسب الخدمة').setFontWeight('bold');
  s.getRange('H3').setFormula(`=IFERROR(QUERY(${q}!C2:D,"select C, count(C), sum(D) where C is not null group by C label count(C) 'الطلبات', sum(D) 'الحاويات'",0),"")`);

  // الحاويات عند العملاء مع الأيام والمبلغ الإضافي (يعتمد على تاريخ التسليم الفعلي، وإذا فاضي يأخذ الموعد المطلوب)
  s.getRange('A13').setValue('الحاويات عند العملاء الآن').setFontSize(13).setFontWeight('bold');
  const hh = ['رقم الطلب','الاسم','الجوال','الموقع','رقم الحاوية','تاريخ التسليم','الأيام','أيام إضافية','مبلغ إضافي قبل الضريبة'];
  s.getRange(14,1,1,hh.length).setValues([hh]).setFontWeight('bold').setBackground('#0F2439').setFontColor('#FFC107');
  const del = `IF(${q}!U2:U="",${q}!M2:M,${q}!U2:U)`;
  const days = `(TODAY()-${del})`;
  const over = `IF(${days}>${pr}!C6,${days}-${pr}!C6,0)`;
  s.getRange('A15').setFormula(
    `=IFERROR(SORT(FILTER({${q}!A2:A,${q}!F2:F,${q}!G2:G,${q}!I2:I&" - "&${q}!J2:J,` +
    `IF(${q}!V2:V="",${q}!E2:E,${q}!V2:V),${del},${days},${over},${over}*${pr}!C5},${q}!S2:S="عند العميل"),7,FALSE),"لا توجد حاويات عند العملاء")`);
  s.getRange('F15:F').setNumberFormat('yyyy-mm-dd');
  s.getRange('A12').setValue('الصفوف اللي فيها "أيام إضافية" أكبر من صفر: الحاوية تجاوزت المدة المشمولة.').setFontColor('#5D6B78');
  s.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H15>0')
    .setBackground('#F4D6D6').setRanges([s.getRange('A15:I200')]).build()]);

  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('الورقة1') || ss.getSheetByName('ورقة1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  ss.setActiveSheet(o);
}

/* ================= واجهة الصفحة ================= */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'prices') {
    return json_({ok: true, prices: readPrices_()});
  }
  const html = Utilities.newBlob(Utilities.base64Decode(PAGE_B64)).getDataAsString('UTF-8');
  return HtmlService.createHtmlOutput(html)
    .setTitle('ثراء الرمال | اطلب حاوية مخلفات بناء')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// تُستدعى من الصفحة عند فتحها من رابط Apps Script
function getPrices() { return readPrices_(); }
function submitOrder(o) { return handleOrder_(o); }

// تُستخدم من صفحة GitHub Pages
function doPost(e) {
  let o;
  try { o = JSON.parse(e.postData.contents); } catch (_) { return json_({ok: false, error: 'bad_request'}); }
  return json_(handleOrder_(o));
}

function handleOrder_(o) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const err = validate_(o);
    if (err) return {ok: false, error: err};

    const P = readPrices_();
    const qty = Math.min(5, Math.max(1, Math.floor(Number(o.qty) || 1)));
    const unit = Number(P[o.service]) || 0;
    const sub = r2_(unit * qty);
    const vat = r2_(sub * VAT);
    const total = r2_(sub + vat);
    const now = new Date();
    const id = nextId_();
    const date = parseDate_(o.date);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSheetByName(SHEET_ORDERS).appendRow([
      id, now, SERVICES[o.service], qty, clean_(o.containerNo, 20), clean_(o.name, 60), "'" + o.phone,
      clean_(o.customerType, 20), clean_(o.city, 30), clean_(o.district, 40), clean_(o.mapLink, 300),
      clean_(o.address, 150), date || clean_(o.date, 12), clean_(o.payment, 20),
      sub, vat, total, clean_(o.notes, 300), 'جديد', '', '', '', ''
    ]);

    notify_(id, o, qty, sub, vat, total);
    return {ok: true, id: id, subtotal: sub, vat: vat, total: total};
  } catch (x) {
    return {ok: false, error: 'server_error'};
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ================= أدوات ================= */
function readPrices_() {
  const out = {};
  SEED_PRICES.forEach(r => out[r[0]] = r[2]);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRICES);
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
      const k = String(r[0]).trim();
      if (out.hasOwnProperty(k) && r[2] !== '' && !isNaN(Number(r[2]))) out[k] = Number(r[2]);
    });
  }
  return out;
}

function validate_(o) {
  if (!o || !SERVICES.hasOwnProperty(o.service)) return 'service';
  if (!o.name || String(o.name).trim().length < 2) return 'name';
  if (!/^05\d{8}$/.test(String(o.phone || ''))) return 'phone';
  if (CITIES.indexOf(o.city) < 0) return 'city';
  if (!o.district) return 'district';
  if (!o.mapLink && String(o.address || '').trim().length < 5) return 'location';
  if (!o.payment) return 'payment';
  return '';
}

function parseDate_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

// يمنع إدخال معادلات في الجدول من بيانات العميل
function clean_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function nextId_() {
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty('seq') || 1000) + 1;
  props.setProperty('seq', String(n));
  return 'TR-' + n;
}

function notify_(id, o, qty, sub, vat, total) {
  try {
    const to = Session.getEffectiveUser().getEmail();
    if (!to) return;
    const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    MailApp.sendEmail({
      to: to,
      name: 'طلبات ثراء الرمال',
      subject: `🟨 طلب حاوية ${id} — ${SERVICES[o.service]} × ${qty} — ${o.city} ${o.district} — ${total} ريال`,
      body: [`طلب جديد من صفحة ثراء الرمال`, `رقم الطلب: ${id}`, '',
             `الخدمة: ${SERVICES[o.service]} (20 ياردة) × ${qty}`,
             o.containerNo ? `رقم الحاوية عند العميل: ${o.containerNo}` : '',
             `المبلغ قبل الضريبة: ${sub} ريال`, `الضريبة 15%: ${vat} ريال`, `الإجمالي شامل الضريبة: ${total} ريال`, '',
             `الاسم: ${o.name}`, `الجوال: ${o.phone}`, `نوع العميل: ${o.customerType || '-'}`,
             `المدينة والحي: ${o.city} - ${o.district}`, `الموقع: ${o.mapLink || '-'}`,
             o.address ? `الوصف: ${o.address}` : '', `الموعد المطلوب: ${o.date}`, `الدفع: ${o.payment}`,
             o.notes ? `ملاحظات: ${o.notes}` : '', '', `لوحة الطلبات: ${url}`].filter(x => x !== null).join('\n')
    });
  } catch (_) {}
}

function r2_(n) { return Math.round(n * 100) / 100; }

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= صفحة الطلب (base64) ================= */
const PAGE_B64 = 'PCFkb2N0eXBlIGh0bWw+CjxodG1sIGxhbmc9ImFyIiBkaXI9InJ0bCI+CjxoZWFkPgo8bWV0YSBjaGFyc2V0PSJ1dGYtOCI+CjxtZXRhIG5hbWU9InZpZXdwb3J0IiBjb250ZW50PSJ3aWR0aD1kZXZpY2Utd2lkdGgsaW5pdGlhbC1zY2FsZT0xLHZpZXdwb3J0LWZpdD1jb3ZlciI+Cjx0aXRsZT7Yq9ix2KfYoSDYp9mE2LHZhdin2YQgfCDYp9i32YTYqCDYrdin2YjZitipINmF2K7ZhNmB2KfYqiDYqNmG2KfYoSDZgdmKINin2YTZhdmG2LfZgtipINin2YTYtNix2YLZitipPC90aXRsZT4KPG1ldGEgbmFtZT0iZGVzY3JpcHRpb24iIGNvbnRlbnQ9Itit2KfZiNmK2Kkg2YXYrtmE2YHYp9iqINio2YbYp9ihIDIwINmK2KfYsdiv2Kkg2YXZhiDYq9ix2KfYoSDYp9mE2LHZhdin2YQ6IDQwMCDYsdmK2KfZhCDZhNmF2K/YqSAxMCDYo9mK2KfZhSDYtNin2YXZhNipINin2YTYqtmI2LXZitmEINmI2KfZhNiz2K3YqNiMINmF2YYg2KfZhNiu2KjYsSDYpdmE2Ykg2KfZhNmC2LfZitmBLiDYp9i32YTYqCDYp9mE2KLZhiDZiNin2YTYqtmI2LXZitmEINiu2YTYp9mEIDI0INiz2KfYudipLiI+CjxtZXRhIHByb3BlcnR5PSJvZzp0aXRsZSIgY29udGVudD0i2KvYsdin2KEg2KfZhNix2YXYp9mEIHwg2K3Yp9mI2YrYqSDZhdiu2YTZgdin2Kog2KjZhtin2KEgMjAg2YrYp9ix2K/YqSI+CjxtZXRhIHByb3BlcnR5PSJvZzpkZXNjcmlwdGlvbiIgY29udGVudD0iNDAwINix2YrYp9mEINmE2YXYr9ipIDEwINij2YrYp9mFINi02KfZhdmE2Kkg2KfZhNiq2YjYtdmK2YQg2YjYp9mE2LPYrdio2Iwg2YXZhiDYp9mE2K7YqNixINil2YTZiSDYp9mE2YLYt9mK2YEuINin2LfZhNioINio2K/ZgtmK2YLYqS4iPgo8bWV0YSBwcm9wZXJ0eT0ib2c6aW1hZ2UiIGNvbnRlbnQ9Imh0dHBzOi8vd2FzZWVtYW5hbHl0aWMtZG90Y29tLmdpdGh1Yi5pby90aGFyYS1hbHJpbWFsL2Fzc2V0cy9vZy5qcGciPgo8bWV0YSBwcm9wZXJ0eT0ib2c6dXJsIiBjb250ZW50PSJodHRwczovL3dhc2VlbWFuYWx5dGljLWRvdGNvbS5naXRodWIuaW8vdGhhcmEtYWxyaW1hbC8iPgo8bWV0YSBwcm9wZXJ0eT0ib2c6dHlwZSIgY29udGVudD0id2Vic2l0ZSI+CjxtZXRhIG5hbWU9InRoZW1lLWNvbG9yIiBjb250ZW50PSIjMEYyNDM5Ij4KPGxpbmsgcmVsPSJpY29uIiBocmVmPSJodHRwczovL3dhc2VlbWFuYWx5dGljLWRvdGNvbS5naXRodWIuaW8vdGhhcmEtYWxyaW1hbC9hc3NldHMvbG9nby5qcGciPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20iPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ3N0YXRpYy5jb20iIGNyb3Nzb3JpZ2luPgo8bGluayBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tL2NzczI/ZmFtaWx5PVRhamF3YWw6d2dodEA0MDA7NTAwOzcwMDs4MDAmZGlzcGxheT1zd2FwIiByZWw9InN0eWxlc2hlZXQiPgo8c3R5bGU+Cjpyb290ewogIC0tbmF2eTojMEYyNDM5OyAtLW5hdnktMjojMUEzNjU0OyAtLXllbGxvdzojRkZDMTA3OyAtLXllbGxvdy0yOiNGRkQ1NEE7CiAgLS1iZzojRjVGNkY4OyAtLWNhcmQ6I0ZGRkZGRjsgLS1pbms6IzE0MjAyQjsgLS1tdXRlZDojNUQ2Qjc4OyAtLWxpbmU6I0UyRTZFQjsKICAtLW9rOiMxRThFNUE7IC0tZXJyOiNDMDM5MkI7IC0tcmFkaXVzOjE2cHg7CiAgLS1zdHJpcGVzOnJlcGVhdGluZy1saW5lYXItZ3JhZGllbnQoLTQ1ZGVnLHZhcigtLXllbGxvdykgMCAxNHB4LHZhcigtLW5hdnkpIDE0cHggMjhweCk7Cn0KKntib3gtc2l6aW5nOmJvcmRlci1ib3g7bWFyZ2luOjA7cGFkZGluZzowfQpodG1sLGJvZHl7YmFja2dyb3VuZDp2YXIoLS1iZyl9CmJvZHl7Zm9udC1mYW1pbHk6J1RhamF3YWwnLHN5c3RlbS11aSxzYW5zLXNlcmlmO2NvbG9yOnZhcigtLWluayk7bGluZS1oZWlnaHQ6MS42O3BhZGRpbmctYm90dG9tOjExNnB4Oy13ZWJraXQtdGFwLWhpZ2hsaWdodC1jb2xvcjp0cmFuc3BhcmVudH0KYnV0dG9uLGlucHV0LHNlbGVjdCx0ZXh0YXJlYXtmb250OmluaGVyaXQ7Y29sb3I6aW5oZXJpdH0KLndyYXB7bWF4LXdpZHRoOjcyMHB4O21hcmdpbjowIGF1dG87cGFkZGluZzowIDE2cHh9CgovKiDYp9mE2YfZitiv2LEgKi8KaGVhZGVye2JhY2tncm91bmQ6dmFyKC0tbmF2eSk7Y29sb3I6I2ZmZjtwYWRkaW5nOjE2cHggMCAwfQouYnJhbmR7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6MTJweH0KLmJyYW5kIGltZ3t3aWR0aDo2MnB4O2hlaWdodDphdXRvO2ZsZXg6bm9uZX0KLmJyYW5kIGgxe2ZvbnQtc2l6ZToyOHB4O2ZvbnQtd2VpZ2h0OjgwMDtsaW5lLWhlaWdodDoxLjE7Y29sb3I6dmFyKC0teWVsbG93KX0KLmJyYW5kIHB7Zm9udC1zaXplOjE0cHg7b3BhY2l0eTouOX0KLnRydXN0e2Rpc3BsYXk6ZmxleDtnYXA6OHB4O2ZsZXgtd3JhcDp3cmFwO21hcmdpbjoxNHB4IDAgMTZweH0KLnRydXN0IHNwYW57YmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC4wOCk7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDI1NSwyNTUsMjU1LC4xOCk7Ym9yZGVyLXJhZGl1czo5OTlweDtwYWRkaW5nOjRweCAxMnB4O2ZvbnQtc2l6ZToxM3B4fQouc3RyaXBlc3toZWlnaHQ6MTJweDtiYWNrZ3JvdW5kOnZhcigtLXN0cmlwZXMpfQoKLyog2KfZhNio2LfYp9mC2Kkg2KfZhNix2KbZitiz2YrYqSAqLwouaGVyb3twb3NpdGlvbjpyZWxhdGl2ZTttYXJnaW4tdG9wOjE2cHg7Ym9yZGVyLXJhZGl1czp2YXIoLS1yYWRpdXMpO292ZXJmbG93OmhpZGRlbjtiYWNrZ3JvdW5kOnZhcigtLW5hdnkpO2JveC1zaGFkb3c6MCA2cHggMjJweCByZ2JhKDE1LDM2LDU3LC4xOCl9Ci5oZXJvIGltZ3tkaXNwbGF5OmJsb2NrO3dpZHRoOjEwMCU7aGVpZ2h0OmF1dG87YXNwZWN0LXJhdGlvOjE2LzEwO29iamVjdC1maXQ6Y292ZXI7b2JqZWN0LXBvc2l0aW9uOmNlbnRlciA1NSV9Ci5iYWRnZXtwb3NpdGlvbjphYnNvbHV0ZTt0b3A6MTJweDtsZWZ0OjEycHg7d2lkdGg6MTE4cHg7aGVpZ2h0OjExOHB4O2JvcmRlci1yYWRpdXM6NTAlO2JhY2tncm91bmQ6dmFyKC0teWVsbG93KTtib3JkZXI6NXB4IHNvbGlkIHZhcigtLW5hdnkpOwogIGRpc3BsYXk6ZmxleDtmbGV4LWRpcmVjdGlvbjpjb2x1bW47YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXI7Y29sb3I6dmFyKC0tbmF2eSk7dGV4dC1hbGlnbjpjZW50ZXI7bGluZS1oZWlnaHQ6MS4wNX0KLmJhZGdlIGJ7Zm9udC1zaXplOjM4cHg7Zm9udC13ZWlnaHQ6ODAwfQouYmFkZ2Ugc21hbGx7Zm9udC1zaXplOjEzcHg7Zm9udC13ZWlnaHQ6NzAwfQouYmFkZ2UgaXtmb250LXN0eWxlOm5vcm1hbDtmb250LXNpemU6MTFweDtmb250LXdlaWdodDo1MDA7bWFyZ2luLXRvcDozcHh9Ci5oZXJvIC5pbmZve2NvbG9yOiNmZmY7cGFkZGluZzoxNHB4IDE2cHggMTZweH0KLmhlcm8gaDJ7Zm9udC1zaXplOjIycHg7Zm9udC13ZWlnaHQ6ODAwO2NvbG9yOnZhcigtLXllbGxvdyl9Ci5oZXJvIHB7Zm9udC1zaXplOjE1cHg7b3BhY2l0eTouOTJ9CgovKiDYp9mE2K7Yt9mI2KfYqiAqLwpzZWN0aW9uLnN0ZXB7YmFja2dyb3VuZDp2YXIoLS1jYXJkKTtib3JkZXItcmFkaXVzOnZhcigtLXJhZGl1cyk7cGFkZGluZzoxOHB4IDE2cHg7bWFyZ2luLXRvcDoxNnB4O2JvcmRlcjoxcHggc29saWQgdmFyKC0tbGluZSl9CnNlY3Rpb24uc3RlcCBoMntmb250LXNpemU6MTlweDtmb250LXdlaWdodDo4MDA7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6MTBweDttYXJnaW4tYm90dG9tOjEycHh9Ci5ue3dpZHRoOjI4cHg7aGVpZ2h0OjI4cHg7Ym9yZGVyLXJhZGl1czo1MCU7YmFja2dyb3VuZDp2YXIoLS1uYXZ5KTtjb2xvcjp2YXIoLS15ZWxsb3cpO2Rpc3BsYXk6aW5saW5lLWZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXI7Zm9udC1zaXplOjE1cHg7ZmxleDpub25lfQouaGludHtmb250LXNpemU6MTNweDtjb2xvcjp2YXIoLS1tdXRlZCk7bWFyZ2luLXRvcDo2cHh9Cgouc3Zje2Rpc3BsYXk6Z3JpZDtnYXA6MTBweH0KLnN2YyBidXR0b257ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtqdXN0aWZ5LWNvbnRlbnQ6c3BhY2UtYmV0d2VlbjtnYXA6MTJweDt0ZXh0LWFsaWduOnJpZ2h0O2JhY2tncm91bmQ6I2ZmZjtib3JkZXI6MnB4IHNvbGlkIHZhcigtLWxpbmUpO2JvcmRlci1yYWRpdXM6MTRweDtwYWRkaW5nOjEycHggMTRweDtjdXJzb3I6cG9pbnRlcn0KLnN2YyBidXR0b24gYntkaXNwbGF5OmJsb2NrO2ZvbnQtc2l6ZToxN3B4fQouc3ZjIGJ1dHRvbiBzcGFue2Rpc3BsYXk6YmxvY2s7Zm9udC1zaXplOjEzcHg7Y29sb3I6dmFyKC0tbXV0ZWQpfQouc3ZjIGJ1dHRvbiBlbXtmb250LXN0eWxlOm5vcm1hbDtmb250LXdlaWdodDo4MDA7Zm9udC1zaXplOjE3cHg7d2hpdGUtc3BhY2U6bm93cmFwO2NvbG9yOnZhcigtLW5hdnkpfQouc3ZjIGJ1dHRvblthcmlhLXByZXNzZWQ9InRydWUiXXtib3JkZXItY29sb3I6dmFyKC0tbmF2eSk7YmFja2dyb3VuZDojRkZGOEUxO2JveC1zaGFkb3c6aW5zZXQgMCAwIDAgMXB4IHZhcigtLW5hdnkpfQoKLnF0eXtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpzcGFjZS1iZXR3ZWVuO21hcmdpbi10b3A6MTRweDtnYXA6MTJweH0KLnF0eSBsYWJlbHtmb250LXdlaWdodDo3MDB9Ci5zdGVwcGVye2Rpc3BsYXk6ZmxleDthbGlnbi1pdGVtczpjZW50ZXI7Ym9yZGVyOjJweCBzb2xpZCB2YXIoLS1saW5lKTtib3JkZXItcmFkaXVzOjEycHg7b3ZlcmZsb3c6aGlkZGVufQouc3RlcHBlciBidXR0b257d2lkdGg6NDRweDtoZWlnaHQ6NDRweDtib3JkZXI6MDtiYWNrZ3JvdW5kOiNGMEYyRjU7Zm9udC1zaXplOjIycHg7Zm9udC13ZWlnaHQ6NzAwO2N1cnNvcjpwb2ludGVyfQouc3RlcHBlciBvdXRwdXR7bWluLXdpZHRoOjQ0cHg7dGV4dC1hbGlnbjpjZW50ZXI7Zm9udC13ZWlnaHQ6ODAwO2ZvbnQtc2l6ZToxOHB4fQoKLmNoaXBze2Rpc3BsYXk6ZmxleDtmbGV4LXdyYXA6d3JhcDtnYXA6OHB4fQouY2hpcHMgYnV0dG9ue2JvcmRlcjoycHggc29saWQgdmFyKC0tbGluZSk7YmFja2dyb3VuZDojZmZmO2JvcmRlci1yYWRpdXM6OTk5cHg7cGFkZGluZzo4cHggMTZweDtmb250LXdlaWdodDo3MDA7Y3Vyc29yOnBvaW50ZXJ9Ci5jaGlwcyBidXR0b25bYXJpYS1wcmVzc2VkPSJ0cnVlIl17YmFja2dyb3VuZDp2YXIoLS1uYXZ5KTtjb2xvcjojZmZmO2JvcmRlci1jb2xvcjp2YXIoLS1uYXZ5KX0KCi5maWVsZHttYXJnaW4tdG9wOjEycHh9Ci5maWVsZCBsYWJlbHtkaXNwbGF5OmJsb2NrO2ZvbnQtd2VpZ2h0OjcwMDttYXJnaW4tYm90dG9tOjZweDtmb250LXNpemU6MTVweH0KLmZpZWxkIGlucHV0LC5maWVsZCB0ZXh0YXJlYXt3aWR0aDoxMDAlO2JvcmRlcjoycHggc29saWQgdmFyKC0tbGluZSk7Ym9yZGVyLXJhZGl1czoxMnB4O3BhZGRpbmc6MTFweCAxMnB4O2JhY2tncm91bmQ6I2ZmZjtmb250LXNpemU6MTZweH0KLmZpZWxkIGlucHV0OmZvY3VzLC5maWVsZCB0ZXh0YXJlYTpmb2N1c3tvdXRsaW5lOm5vbmU7Ym9yZGVyLWNvbG9yOnZhcigtLW5hdnkpfQouZmllbGQgdGV4dGFyZWF7bWluLWhlaWdodDo3NnB4O3Jlc2l6ZTp2ZXJ0aWNhbH0KLnJvdzJ7ZGlzcGxheTpncmlkO2dyaWQtdGVtcGxhdGUtY29sdW1uczoxZnIgMWZyO2dhcDoxMHB4fQoubG9jYnRue3dpZHRoOjEwMCU7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtqdXN0aWZ5LWNvbnRlbnQ6Y2VudGVyO2dhcDo4cHg7YmFja2dyb3VuZDp2YXIoLS1uYXZ5KTtjb2xvcjojZmZmO2JvcmRlcjowO2JvcmRlci1yYWRpdXM6MTJweDtwYWRkaW5nOjEycHg7Zm9udC13ZWlnaHQ6NzAwO2N1cnNvcjpwb2ludGVyO21hcmdpbi10b3A6MTJweH0KLmxvY2J0bi5va3tiYWNrZ3JvdW5kOnZhcigtLW9rKX0KLm9ye3RleHQtYWxpZ246Y2VudGVyO2NvbG9yOnZhcigtLW11dGVkKTtmb250LXNpemU6MTNweDttYXJnaW4tdG9wOjhweH0KCi53YXN0ZXtkaXNwbGF5OmdyaWQ7Z3JpZC10ZW1wbGF0ZS1jb2x1bW5zOjFmciAxZnI7Z2FwOjEwcHh9Ci53YXN0ZSBkaXZ7Ym9yZGVyLXJhZGl1czoxMnB4O3BhZGRpbmc6MTJweH0KLndhc3RlIC55ZXN7YmFja2dyb3VuZDojRTdGNkVFO2JvcmRlcjoxcHggc29saWQgI0JGRTVEMH0KLndhc3RlIC5ub3tiYWNrZ3JvdW5kOiNGQkVBRUE7Ym9yZGVyOjFweCBzb2xpZCAjRjFDNkM2fQoud2FzdGUgaDN7Zm9udC1zaXplOjE1cHg7bWFyZ2luLWJvdHRvbTo0cHh9Ci53YXN0ZSB1bHtsaXN0LXN0eWxlOm5vbmU7Zm9udC1zaXplOjE0cHh9Ci53YXN0ZSBsaTo6YmVmb3Jle2NvbnRlbnQ6IuKAoiAifQouYWdyZWV7ZGlzcGxheTpmbGV4O2dhcDoxMHB4O2FsaWduLWl0ZW1zOmZsZXgtc3RhcnQ7bWFyZ2luLXRvcDoxMnB4O2ZvbnQtd2VpZ2h0OjcwMDtmb250LXNpemU6MTVweDtjdXJzb3I6cG9pbnRlcn0KLmFncmVlIGlucHV0e3dpZHRoOjIycHg7aGVpZ2h0OjIycHg7bWFyZ2luLXRvcDoycHg7YWNjZW50LWNvbG9yOnZhcigtLW5hdnkpO2ZsZXg6bm9uZX0KCi5zdW17Ym9yZGVyLXRvcDoxcHggZGFzaGVkIHZhcigtLWxpbmUpO21hcmdpbi10b3A6NHB4fQouc3VtIGRpdntkaXNwbGF5OmZsZXg7anVzdGlmeS1jb250ZW50OnNwYWNlLWJldHdlZW47cGFkZGluZzo2cHggMDtmb250LXNpemU6MTVweH0KLnN1bSAudG90e2ZvbnQtd2VpZ2h0OjgwMDtmb250LXNpemU6MThweDtib3JkZXItdG9wOjFweCBzb2xpZCB2YXIoLS1saW5lKTttYXJnaW4tdG9wOjRweDtwYWRkaW5nLXRvcDoxMHB4fQoubm90ZXtiYWNrZ3JvdW5kOiNGRkY4RTE7Ym9yZGVyLXJhZGl1czoxMnB4O3BhZGRpbmc6MTBweCAxMnB4O2ZvbnQtc2l6ZToxNHB4O21hcmdpbi10b3A6MTBweH0KLmVycntjb2xvcjp2YXIoLS1lcnIpO2ZvbnQtd2VpZ2h0OjcwMDttaW4taGVpZ2h0OjEuNmVtO21hcmdpbi10b3A6MTBweDt0ZXh0LWFsaWduOmNlbnRlcn0KCi8qINin2YTYtNix2YrYtyDYp9mE2LPZgdmE2YogKi8KLmJhcntwb3NpdGlvbjpmaXhlZDtpbnNldDphdXRvIDAgMCAwO2JhY2tncm91bmQ6I2ZmZjtib3JkZXItdG9wOjFweCBzb2xpZCB2YXIoLS1saW5lKTtib3gtc2hhZG93OjAgLTZweCAyMHB4IHJnYmEoMCwwLDAsLjA4KTtwYWRkaW5nOjEwcHggMCBjYWxjKDEwcHggKyBlbnYoc2FmZS1hcmVhLWluc2V0LWJvdHRvbSkpO3otaW5kZXg6MTB9Ci5iYXIgLndyYXB7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6MTJweH0KLmJhciAudHtmbGV4OjE7bGluZS1oZWlnaHQ6MS4yNX0KLmJhciAudCBzbWFsbHtkaXNwbGF5OmJsb2NrO2NvbG9yOnZhcigtLW11dGVkKTtmb250LXNpemU6MTJweH0KLmJhciAudCBie2ZvbnQtc2l6ZToyMHB4fQouYmFyIGJ1dHRvbntiYWNrZ3JvdW5kOnZhcigtLXllbGxvdyk7Y29sb3I6dmFyKC0tbmF2eSk7Ym9yZGVyOjA7Ym9yZGVyLXJhZGl1czoxNHB4O3BhZGRpbmc6MTNweCAyMnB4O2ZvbnQtd2VpZ2h0OjgwMDtmb250LXNpemU6MTdweDtjdXJzb3I6cG9pbnRlcjtib3gtc2hhZG93OjAgM3B4IDAgI0M5OTUwMH0KLmJhciBidXR0b246ZGlzYWJsZWR7b3BhY2l0eTouNn0KCi8qINio2LnYryDYp9mE2KXYsdiz2KfZhCAqLwouZG9uZXtkaXNwbGF5Om5vbmU7dGV4dC1hbGlnbjpjZW50ZXI7cGFkZGluZy10b3A6MjhweH0KLmRvbmUgLmlje3dpZHRoOjg0cHg7aGVpZ2h0Ojg0cHg7Ym9yZGVyLXJhZGl1czo1MCU7YmFja2dyb3VuZDp2YXIoLS1vayk7Y29sb3I6I2ZmZjtmb250LXNpemU6NDZweDtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXI7bWFyZ2luOjAgYXV0byAxMnB4fQouZG9uZSBoMntmb250LXNpemU6MjZweH0KLmRvbmUgcHtjb2xvcjp2YXIoLS1tdXRlZCk7bWFyZ2luLXRvcDo2cHh9Ci5kb25lIC5pZHtmb250LXdlaWdodDo4MDA7Zm9udC1zaXplOjIwcHg7Y29sb3I6dmFyKC0tbmF2eSk7bWFyZ2luLXRvcDo4cHh9Ci53YXtkaXNwbGF5OmlubGluZS1mbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6OHB4O2JhY2tncm91bmQ6IzI1RDM2Njtjb2xvcjojZmZmO3RleHQtZGVjb3JhdGlvbjpub25lO2ZvbnQtd2VpZ2h0OjgwMDtib3JkZXItcmFkaXVzOjE0cHg7cGFkZGluZzoxNHB4IDIycHg7bWFyZ2luLXRvcDoxOHB4O2ZvbnQtc2l6ZToxN3B4fQoKZm9vdGVye3RleHQtYWxpZ246Y2VudGVyO2NvbG9yOnZhcigtLW11dGVkKTtmb250LXNpemU6MTNweDtwYWRkaW5nOjIycHggMCA4cHh9CmZvb3RlciBhe2NvbG9yOnZhcigtLW5hdnkpO2ZvbnQtd2VpZ2h0OjcwMH0KQG1lZGlhIChtYXgtd2lkdGg6NDIwcHgpey5iYWRnZXt3aWR0aDoxMDBweDtoZWlnaHQ6MTAwcHh9LmJhZGdlIGJ7Zm9udC1zaXplOjMycHh9Lndhc3Rle2dyaWQtdGVtcGxhdGUtY29sdW1uczoxZnJ9fQo8L3N0eWxlPgo8L2hlYWQ+Cjxib2R5PgoKPGhlYWRlcj4KICA8ZGl2IGNsYXNzPSJ3cmFwIj4KICAgIDxkaXYgY2xhc3M9ImJyYW5kIj4KICAgICAgPGltZyBzcmM9Imh0dHBzOi8vd2FzZWVtYW5hbHl0aWMtZG90Y29tLmdpdGh1Yi5pby90aGFyYS1hbHJpbWFsL2Fzc2V0cy9lbWJsZW0uanBnIiBhbHQ9Iti02LnYp9ixINir2LHYp9ihINin2YTYsdmF2KfZhCIgd2lkdGg9IjYyIiBoZWlnaHQ9IjU5Ij4KICAgICAgPGRpdj4KICAgICAgICA8aDE+2KvYsdin2KEg2KfZhNix2YXYp9mEPC9oMT4KICAgICAgICA8cD7Yrdin2YjZitin2Kog2YXYrtmE2YHYp9iqINio2YbYp9ihIDIwINmK2KfYsdiv2Kkgwrcg2KfZhNmF2YbYt9mC2Kkg2KfZhNi02LHZgtmK2Kk8L3A+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJ0cnVzdCI+CiAgICAgIDxzcGFuPvCfmpog2KfZhNiq2YjYtdmK2YQg2K7ZhNin2YQgMjQg2LPYp9i52Kk8L3NwYW4+CiAgICAgIDxzcGFuPvCfk4UgPGIgY2xhc3M9InBEYXlzIj4xMDwvYj4g2KPZitin2YUg2LnZhtiv2YM8L3NwYW4+CiAgICAgIDxzcGFuPvCfkrMg2KfZhNiv2YHYuSDZiNmC2Kog2KfZhNiq2LPZhNmK2YU8L3NwYW4+CiAgICA8L2Rpdj4KICA8L2Rpdj4KICA8ZGl2IGNsYXNzPSJzdHJpcGVzIj48L2Rpdj4KPC9oZWFkZXI+Cgo8bWFpbiBjbGFzcz0id3JhcCIgaWQ9ImFwcCI+CiAgPGRpdiBjbGFzcz0iaGVybyI+CiAgICA8aW1nIHNyYz0iaHR0cHM6Ly93YXNlZW1hbmFseXRpYy1kb3Rjb20uZ2l0aHViLmlvL3RoYXJhLWFscmltYWwvYXNzZXRzL2NvbnRhaW5lci5qcGciIGFsdD0i2K3Yp9mI2YrYqSDZhdiu2YTZgdin2Kog2KjZhtin2KEgMjAg2YrYp9ix2K/YqSDZhdmGINir2LHYp9ihINin2YTYsdmF2KfZhCIgd2lkdGg9IjkwMCIgaGVpZ2h0PSI3MDciPgogICAgPGRpdiBjbGFzcz0iYmFkZ2UiPjxiIGNsYXNzPSJwTmV3Ij40MDA8L2I+PHNtYWxsPtix2YrYp9mEIC8gPHNwYW4gY2xhc3M9InBEYXlzIj4xMDwvc3Bhbj4g2KPZitin2YU8L3NtYWxsPjxpPisg2LbYsdmK2KjYqSAxNSU8L2k+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJpbmZvIj4KICAgICAgPGgyPtit2KfZiNmK2Kkg2YXYrtmE2YHYp9iqINio2YbYp9ihPC9oMj4KICAgICAgPHA+2LPYudipIDIwINmK2KfYsdiv2Kkgwrcg2K3ZiNin2YTZiiAxNSDZhcKzIMK3INi02KfZhdmE2Kkg2KfZhNiq2YjYtdmK2YQg2YjYp9mE2LPYrdioLiDYqtix2YXZitmFINij2Ygg2KjZhtin2KHYnyDZhtmI2LXZhNmH2Kcg2YTZhdmI2YLYudmDINmI2YbYs9it2KjZh9inINmI2YLYqiDZhdinINiq2YXYqtmE2YouPC9wPgogICAgPC9kaXY+CiAgPC9kaXY+CgogIDxzZWN0aW9uIGNsYXNzPSJzdGVwIiBpZD0iczEiPgogICAgPGgyPjxzcGFuIGNsYXNzPSJuIj4xPC9zcGFuPiDZiNi0INiq2K3Yqtin2KzYnzwvaDI+CiAgICA8ZGl2IGNsYXNzPSJzdmMiIGlkPSJzdmMiPjwvZGl2PgogICAgPGRpdiBjbGFzcz0icXR5Ij4KICAgICAgPGxhYmVsIGZvcj0icXR5T3V0Ij7Yudiv2K8g2KfZhNit2KfZiNmK2KfYqjwvbGFiZWw+CiAgICAgIDxkaXYgY2xhc3M9InN0ZXBwZXIiPgogICAgICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBpZD0icU1pbnVzIiBhcmlhLWxhYmVsPSLYo9mC2YQiPuKIkjwvYnV0dG9uPgogICAgICAgIDxvdXRwdXQgaWQ9InF0eU91dCI+MTwvb3V0cHV0PgogICAgICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBpZD0icVBsdXMiIGFyaWEtbGFiZWw9Itij2YPYq9ixIj4rPC9idXR0b24+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWVsZCIgaWQ9ImNvbnRGaWVsZCIgc3R5bGU9ImRpc3BsYXk6bm9uZSI+CiAgICAgIDxsYWJlbCBmb3I9ImNvbnRhaW5lck5vIj7YsdmC2YUg2KfZhNit2KfZiNmK2Kkg2KfZhNmE2Yog2LnZhtiv2YMgKNin2K7YqtmK2KfYsdmKKTwvbGFiZWw+CiAgICAgIDxpbnB1dCBpZD0iY29udGFpbmVyTm8iIGlucHV0bW9kZT0ibnVtZXJpYyIgbWF4bGVuZ3RoPSIyMCIgcGxhY2Vob2xkZXI9Itin2YTYsdmC2YUg2KfZhNmF2YPYqtmI2Kgg2LnZhNmJINin2YTYrdin2YjZitip2Iwg2YXYq9mEIDA5Ij4KICAgIDwvZGl2PgogICAgPHAgY2xhc3M9ImhpbnQiPtmD2YQg2YrZiNmFINil2LbYp9mB2Yog2KjYudivINin2YTZitmI2YUgPHNwYW4gY2xhc3M9InBEYXlzIj4xMDwvc3Bhbj46IDxiIGNsYXNzPSJwRXh0cmEiPjMwPC9iPiDYsdmK2KfZhCArINin2YTYttix2YrYqNipLjwvcD4KICA8L3NlY3Rpb24+CgogIDxzZWN0aW9uIGNsYXNzPSJzdGVwIiBpZD0iczIiPgogICAgPGgyPjxzcGFuIGNsYXNzPSJuIj4yPC9zcGFuPiDZiNmK2YYg2KfZhNmF2YjZgti52J88L2gyPgogICAgPGRpdiBjbGFzcz0iY2hpcHMiIGlkPSJjaXRpZXMiPjwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJkaXN0cmljdCI+2KfZhNit2Yo8L2xhYmVsPgogICAgICA8aW5wdXQgaWQ9ImRpc3RyaWN0IiBtYXhsZW5ndGg9IjQwIiBwbGFjZWhvbGRlcj0i2YXYq9mE2KfZizog2KfZhNi52LLZitiy2YrYqSIgYXV0b2NvbXBsZXRlPSJhZGRyZXNzLWxldmVsMyI+CiAgICA8L2Rpdj4KICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBjbGFzcz0ibG9jYnRuIiBpZD0ibG9jYXRlIj7wn5ONINit2K/YryDZhdmI2YLYudmKINin2YTYrdin2YTZijwvYnV0dG9uPgogICAgPHAgY2xhc3M9Im9yIj7Yo9mIINin2YTYtdmCINix2KfYqNi3INin2YTZhdmI2YLYuSDZhdmGINiu2LHYp9im2Lcg2YLZiNmC2YQ8L3A+CiAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+CiAgICAgIDxpbnB1dCBpZD0ibWFwTGluayIgbWF4bGVuZ3RoPSIzMDAiIHBsYWNlaG9sZGVyPSJodHRwczovL21hcHMuYXBwLmdvby5nbC8uLi4iIGlucHV0bW9kZT0idXJsIj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJhZGRyZXNzIj7ZiNi12YEg2KfZhNmF2YPYp9mGICjYp9iu2KrZitin2LHZiik8L2xhYmVsPgogICAgICA8aW5wdXQgaWQ9ImFkZHJlc3MiIG1heGxlbmd0aD0iMTUwIiBwbGFjZWhvbGRlcj0i2YXYq9mE2KfZizog2KzZhtioINin2YTZhdiz2KzYr9iMINmB2YrZhNinINio2YjYp9io2Kkg2LPZiNiv2KfYoSI+CiAgICA8L2Rpdj4KICAgIDxwIGNsYXNzPSJub3RlIj7Zhtix2KzZiCDYqtis2YfZitiyINmF2YPYp9mGINmF2YbYp9iz2Kgg2YTZhNit2KfZiNmK2Kkg2KjYudmK2K8g2LnZhiDZhdiv2KfYrtmEINin2YTYs9mK2KfYsdin2KouPC9wPgogIDwvc2VjdGlvbj4KCiAgPHNlY3Rpb24gY2xhc3M9InN0ZXAiIGlkPSJzMyI+CiAgICA8aDI+PHNwYW4gY2xhc3M9Im4iPjM8L3NwYW4+INmI2LQg2YrZhtit2Lcg2YHZiiDYp9mE2K3Yp9mI2YrYqdifPC9oMj4KICAgIDxkaXYgY2xhc3M9Indhc3RlIj4KICAgICAgPGRpdiBjbGFzcz0ieWVzIj48aDM+4pyFINmF2LPZhdmI2K06INin2YTZhtmB2KfZitin2Kog2KfZhNil2YbYtNin2KbZitipPC9oMz4KICAgICAgICA8dWw+PGxpPtiu2LHYs9in2YbYqTwvbGk+PGxpPti32YjYqDwvbGk+PGxpPtio2YTYp9i3PC9saT48bGk+2LHZhdmEPC9saT48bGk+2K7YtNioINin2YTYqNmG2KfYoTwvbGk+PC91bD48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ibm8iPjxoMz7im5Qg2YXZhdmG2YjYuTwvaDM+CiAgICAgICAgPHVsPjxsaT7Yp9mE2YbZgdin2YrYp9iqINin2YTZhdmG2LLZhNmK2Kk8L2xpPjxsaT7Yp9mE2YXZiNin2K8g2KfZhNmD2YrZhdmK2KfYptmK2Kk8L2xpPjwvdWw+PC9kaXY+CiAgICA8L2Rpdj4KICAgIDxsYWJlbCBjbGFzcz0iYWdyZWUiPjxpbnB1dCB0eXBlPSJjaGVja2JveCIgaWQ9ImFncmVlIj4g2KPYqti52YfYryDYpdmGINin2YTYrdin2YjZitipINmE2YTZhtmB2KfZitin2Kog2KfZhNil2YbYtNin2KbZitipINmB2YLYtzwvbGFiZWw+CiAgPC9zZWN0aW9uPgoKICA8c2VjdGlvbiBjbGFzcz0ic3RlcCIgaWQ9InM0Ij4KICAgIDxoMj48c3BhbiBjbGFzcz0ibiI+NDwvc3Bhbj4g2KjZitin2YbYp9iq2YMg2YjYp9mE2YXZiNi52K88L2gyPgogICAgPGRpdiBjbGFzcz0icm93MiI+CiAgICAgIDxkaXYgY2xhc3M9ImZpZWxkIj48bGFiZWwgZm9yPSJuYW1lIj7Yp9mE2KfYs9mFPC9sYWJlbD48aW5wdXQgaWQ9Im5hbWUiIG1heGxlbmd0aD0iNjAiIGF1dG9jb21wbGV0ZT0ibmFtZSI+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImZpZWxkIj48bGFiZWwgZm9yPSJwaG9uZSI+2KfZhNis2YjYp9mEPC9sYWJlbD48aW5wdXQgaWQ9InBob25lIiBpbnB1dG1vZGU9InRlbCIgbWF4bGVuZ3RoPSIxNCIgcGxhY2Vob2xkZXI9IjA1eHh4eHh4eHgiIGF1dG9jb21wbGV0ZT0idGVsIj48L2Rpdj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmllbGQiPgogICAgICA8bGFiZWw+2KPZhtiqPC9sYWJlbD4KICAgICAgPGRpdiBjbGFzcz0iY2hpcHMiIGlkPSJjdHlwZSI+PC9kaXY+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0iZGF0ZSI+2KfZhNmF2YjYudivINin2YTZhdi32YTZiNioPC9sYWJlbD4KICAgICAgPGlucHV0IHR5cGU9ImRhdGUiIGlkPSJkYXRlIj4KICAgICAgPHAgY2xhc3M9ImhpbnQiIGlkPSJkYXRlSGludCI+2KfZhNiq2YjYtdmK2YQg2K7ZhNin2YQgMjQg2LPYp9i52Kkg2YXZhiDYqtij2YPZitivINin2YTYt9mE2KguINin2YTYrNmF2LnYqSDYpdis2KfYstipLjwvcD4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmllbGQiPgogICAgICA8bGFiZWw+2LfYsdmK2YLYqSDYp9mE2K/Zgdi5ICjZiNmC2Kog2KfZhNiq2LPZhNmK2YUpPC9sYWJlbD4KICAgICAgPGRpdiBjbGFzcz0iY2hpcHMiIGlkPSJwYXkiPjwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9Im5vdGVzIj7ZhdmE2KfYrdi42KfYqiAo2KfYrtiq2YrYp9ix2YopPC9sYWJlbD4KICAgICAgPHRleHRhcmVhIGlkPSJub3RlcyIgbWF4bGVuZ3RoPSIzMDAiIHBsYWNlaG9sZGVyPSLYo9mKINiq2YHYp9i12YrZhCDYqtiz2KfYudiv2YbYpyI+PC90ZXh0YXJlYT4KICAgIDwvZGl2PgoKICAgIDxkaXYgY2xhc3M9InN1bSIgaWQ9InN1bSI+PC9kaXY+CiAgICA8cCBjbGFzcz0iZXJyIiBpZD0iZXJyIiByb2xlPSJhbGVydCI+PC9wPgogIDwvc2VjdGlvbj4KCiAgPGZvb3Rlcj4KICAgINmE2YTYp9iz2KrZgdiz2KfYsTogPGEgaHJlZj0iaHR0cHM6Ly93YS5tZS85NjY1NzkzMjIwMzAiPtmI2KfYqtiz2KfYqCAwNTc5MzIyMDMwPC9hPjxicj4KICAgINir2LHYp9ihINin2YTYsdmF2KfZhCDCtyDZhdmGINin2YTYrtio2LEg2KXZhNmJINin2YTZgti32YrZgSDCtyDYp9mE2KPYs9i52KfYsSDYutmK2LEg2LTYp9mF2YTYqSDYttix2YrYqNipINin2YTZgtmK2YXYqSDYp9mE2YXYttin2YHYqSAxNSUKICA8L2Zvb3Rlcj4KPC9tYWluPgoKPGRpdiBjbGFzcz0iYmFyIiBpZD0iYmFyIj4KICA8ZGl2IGNsYXNzPSJ3cmFwIj4KICAgIDxkaXYgY2xhc3M9InQiPjxzbWFsbD7Yp9mE2KXYrNmF2KfZhNmKINi02KfZhdmEINin2YTYttix2YrYqNipPC9zbWFsbD48YiBpZD0iYmFyVG90YWwiPuKAlDwvYj48L2Rpdj4KICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBpZD0ic3VibWl0Ij7Yo9ix2LPZhCDYp9mE2LfZhNioPC9idXR0b24+CiAgPC9kaXY+CjwvZGl2PgoKPHNlY3Rpb24gY2xhc3M9ImRvbmUgd3JhcCIgaWQ9ImRvbmUiIGFyaWEtbGl2ZT0icG9saXRlIj4KICA8ZGl2IGNsYXNzPSJpYyI+4pyTPC9kaXY+CiAgPGgyPtmI2LXZhNmG2Kcg2LfZhNio2YMhPC9oMj4KICA8cCBjbGFzcz0iaWQiIGlkPSJkb25lSWQiPjwvcD4KICA8cD7YqNmG2YHYqtitINmE2YMg2YjYp9iq2LPYp9ioINmI2KfZhNi32YTYqCDZhdmD2KrZiNio2Iwg2KfYtti62LcgPGI+2KXYsdiz2KfZhDwvYj4g2LnYtNin2YYg2YbYo9mD2K8g2YTZgyDYp9mE2YXZiNi52K8uPC9wPgogIDxhIGNsYXNzPSJ3YSIgaWQ9IndhTGluayIgaHJlZj0iIyIgdGFyZ2V0PSJfdG9wIiByZWw9Im5vb3BlbmVyIj7wn5KsINij2LHYs9mEINin2YTYt9mE2Kgg2LnZhNmJINmI2KfYqtiz2KfYqDwvYT4KICA8cCBzdHlsZT0ibWFyZ2luLXRvcDoxNnB4Ij7ZhdinINin2YbZgdiq2K0g2YjYp9iq2LPYp9io2J8g2KfYtti62Lcg2KfZhNiy2LEg2YHZiNmCLjwvcD4KPC9zZWN0aW9uPgoKPHNjcmlwdD4KKGZ1bmN0aW9uKCl7ICJ1c2Ugc3RyaWN0IjsKLyogPT09PT09PT09PT09INin2YTYpdi52K/Yp9iv2KfYqiA9PT09PT09PT09PT0gKi8KY29uc3QgQ09ORklHID0gewogIC8vINix2KfYqNi3IEdvb2dsZSBBcHBzIFNjcmlwdCDYqNi52K8g2KfZhNmG2LTYsS4g2KXYsNinINiq2Y/YsdmDINmB2KfYttmK2KfZiyDZitmP2LHYs9mEINin2YTYt9mE2Kgg2LnZhNmJINmI2KfYqtiz2KfYqCDZgdmC2LcuCiAgRU5EUE9JTlQ6ICIiLAogIFdIQVRTQVBQOiAiOTY2NTc5MzIyMDMwIiwKICBDSVRJRVM6IFsi2KfZhNiu2KjYsSIsItin2YTYuNmH2LHYp9mGIiwi2KfZhNiv2YXYp9mFIiwi2LPZitmH2KfYqiIsItin2YTZgti32YrZgSIsItmF2K/ZitmG2Kkg2KPYrtix2YkiXSwKICBDVVNUT01FUl9UWVBFUzogWyLZgdix2K8iLCLZhdmC2KfZiNmEINij2Ygg2LTYsdmD2KkiXSwKICBQQVlNRU5UUzogWyLYtNio2YPYqSIsItiq2K3ZiNmK2YQg2KjZhtmD2YoiXSwKICBWQVQ6IDAuMTUgLy8g2YPZhCDYp9mE2KPYs9i52KfYsSDYutmK2LEg2LTYp9mF2YTYqSDYttix2YrYqNipINin2YTZgtmK2YXYqSDYp9mE2YXYttin2YHYqQp9OwovLyDYp9mE2KPYs9i52KfYsSDYp9mE2KfZgdiq2LHYp9i22YrYqSAo2KrYqtit2K/YqyDYqtmE2YLYp9im2YrYp9mLINmF2YYg2YjYsdmC2KkgItin2YTYo9iz2LnYp9ixIiDZgdmKINis2K/ZiNmEINmC2YjZgtmEKQpsZXQgUCA9IHtORVc6NDAwLCBTV0FQOjQwMCwgUElDS1VQOjAsIEVYVFJBX0RBWTozMCwgREFZUzoxMH07CmNvbnN0IFNFUlZJQ0VTID0gWwogIHtpZDoiTkVXIiwgICAgbmFtZToi2K3Yp9mI2YrYqSDYrNiv2YrYr9ipIiwgICAgICBkZXNjOigpID0+IGDZhtmI2LXZhNmH2Kcg2YTZhdmI2YLYudmDINmI2KrYqNmC2Ykg2LnZhtiv2YMgJHtQLkRBWVN9INij2YrYp9mF2Iwg2LTYp9mF2YTYqSDYp9mE2LPYrdioYH0sCiAge2lkOiJTV0FQIiwgICBuYW1lOiLYp9iz2KrYqNiv2KfZhCDYrdin2YjZitipIiwgICAgZGVzYzooKSA9PiAi2YbYs9it2Kgg2KfZhNmF2YXYqtmE2KbYqSDZiNmG2K3YtyDZhNmDINmB2KfYttmK2KkifSwKICB7aWQ6IlBJQ0tVUCIsIG5hbWU6Itiz2K3YqCDYrdin2YjZitipIiwgICAgICAgIGRlc2M6KCkgPT4gItin2YTYrdin2YjZitipINin2YXYqtmE2Kog2KPZiCDYrtmE2LXYqiDZhdmG2YfYpyJ9Cl07Cgpjb25zdCAkID0gaWQgPT4gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoaWQpOwpjb25zdCBzdGF0ZSA9IHtzdmM6Ik5FVyIsIHF0eToxLCBjaXR5Om51bGwsIGN0eXBlOm51bGwsIHBheTpudWxsLCBsb2M6bnVsbCwgc2VuZGluZzpmYWxzZX07CmNvbnN0IGZtdCA9IG4gPT4gTnVtYmVyLmlzSW50ZWdlcihuKSA/IG4udG9Mb2NhbGVTdHJpbmcoImVuLVVTIikgOiBuLnRvTG9jYWxlU3RyaW5nKCJlbi1VUyIse21pbmltdW1GcmFjdGlvbkRpZ2l0czoyLG1heGltdW1GcmFjdGlvbkRpZ2l0czoyfSk7CmNvbnN0IHJvdW5kMiA9IG4gPT4gTWF0aC5yb3VuZChuKjEwMCkvMTAwOwpjb25zdCBlc2MgPSBzID0+IFN0cmluZyhzKS5yZXBsYWNlKC9bJjw+IiddL2csIGMgPT4gKHsiJiI6IiZhbXA7IiwiPCI6IiZsdDsiLCI+IjoiJmd0OyIsJyInOiImcXVvdDsiLCInIjoiJiMzOTsifVtjXSkpOwpjb25zdCBpbkdhcyA9ICgpID0+ICEhKHdpbmRvdy5nb29nbGUgJiYgZ29vZ2xlLnNjcmlwdCAmJiBnb29nbGUuc2NyaXB0LnJ1bik7CmNvbnN0IHN2Y1ByaWNlID0gaWQgPT4gTnVtYmVyKFBbaWRdKSB8fCAwOwpjb25zdCBwcmljZUxhYmVsID0gaWQgPT4gc3ZjUHJpY2UoaWQpID8gYCR7Zm10KHN2Y1ByaWNlKGlkKSl9INix2YrYp9mEYCA6ICLZhdi02YXZiNmEIjsKCmZ1bmN0aW9uIHJlbmRlclByaWNlcygpewogIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoIi5wTmV3IikuZm9yRWFjaChlID0+IGUudGV4dENvbnRlbnQgPSBmbXQoTnVtYmVyKFAuTkVXKSkpOwogIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoIi5wRGF5cyIpLmZvckVhY2goZSA9PiBlLnRleHRDb250ZW50ID0gUC5EQVlTKTsKICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCIucEV4dHJhIikuZm9yRWFjaChlID0+IGUudGV4dENvbnRlbnQgPSBmbXQoTnVtYmVyKFAuRVhUUkFfREFZKSkpOwp9CmZ1bmN0aW9uIHJlbmRlclN2YygpewogICQoInN2YyIpLmlubmVySFRNTCA9IFNFUlZJQ0VTLm1hcChzID0+IGA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1zPSIke3MuaWR9IiBhcmlhLXByZXNzZWQ9IiR7c3RhdGUuc3ZjPT09cy5pZH0iPgogICAgPGRpdj48Yj4ke2VzYyhzLm5hbWUpfTwvYj48c3Bhbj4ke2VzYyhzLmRlc2MoKSl9PC9zcGFuPjwvZGl2PjxlbT4ke3ByaWNlTGFiZWwocy5pZCl9PC9lbT48L2J1dHRvbj5gKS5qb2luKCIiKTsKICAkKCJjb250RmllbGQiKS5zdHlsZS5kaXNwbGF5ID0gc3RhdGUuc3ZjID09PSAiTkVXIiA/ICJub25lIiA6ICJibG9jayI7Cn0KZnVuY3Rpb24gY2hpcHMoZWwsIGxpc3QsIGtleSl7CiAgJChlbCkuaW5uZXJIVE1MID0gbGlzdC5tYXAodiA9PiBgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtdj0iJHtlc2Modil9IiBhcmlhLXByZXNzZWQ9IiR7c3RhdGVba2V5XT09PXZ9Ij4ke2VzYyh2KX08L2J1dHRvbj5gKS5qb2luKCIiKTsKfQpmdW5jdGlvbiB0b3RhbHMoKXsKICBjb25zdCBzdWIgPSByb3VuZDIoc3ZjUHJpY2Uoc3RhdGUuc3ZjKSAqIHN0YXRlLnF0eSk7CiAgY29uc3QgdmF0ID0gcm91bmQyKHN1YiAqIENPTkZJRy5WQVQpOwogIHJldHVybiB7c3ViLCB2YXQsIHRvdGFsOiByb3VuZDIoc3ViICsgdmF0KX07Cn0KZnVuY3Rpb24gcmVuZGVyU3VtKCl7CiAgY29uc3QgcyA9IFNFUlZJQ0VTLmZpbmQoeCA9PiB4LmlkID09PSBzdGF0ZS5zdmMpLCB0ID0gdG90YWxzKCk7CiAgJCgic3VtIikuaW5uZXJIVE1MID0gYAogICAgPGRpdj48c3Bhbj4ke2VzYyhzLm5hbWUpfSDDlyAke3N0YXRlLnF0eX08L3NwYW4+PHNwYW4+JHt0LnN1YiA/IGZtdCh0LnN1YikgKyAiINix2YrYp9mEIiA6ICLZhdi02YXZiNmEIn08L3NwYW4+PC9kaXY+CiAgICA8ZGl2PjxzcGFuPtin2YTYttix2YrYqNipIDE1JTwvc3Bhbj48c3Bhbj4ke2ZtdCh0LnZhdCl9INix2YrYp9mEPC9zcGFuPjwvZGl2PgogICAgPGRpdiBjbGFzcz0idG90Ij48c3Bhbj7Yp9mE2KXYrNmF2KfZhNmKINi02KfZhdmEINin2YTYttix2YrYqNipPC9zcGFuPjxzcGFuPiR7Zm10KHQudG90YWwpfSDYsdmK2KfZhDwvc3Bhbj48L2Rpdj5gOwogICQoImJhclRvdGFsIikudGV4dENvbnRlbnQgPSBgJHtmbXQodC50b3RhbCl9INix2YrYp9mEYDsKICAkKCJxdHlPdXQiKS50ZXh0Q29udGVudCA9IHN0YXRlLnF0eTsKfQpmdW5jdGlvbiByZW5kZXJBbGwoKXsKICByZW5kZXJQcmljZXMoKTsgcmVuZGVyU3ZjKCk7IHJlbmRlclN1bSgpOwogIGNoaXBzKCJjaXRpZXMiLCBDT05GSUcuQ0lUSUVTLCAiY2l0eSIpOyBjaGlwcygiY3R5cGUiLCBDT05GSUcuQ1VTVE9NRVJfVFlQRVMsICJjdHlwZSIpOyBjaGlwcygicGF5IiwgQ09ORklHLlBBWU1FTlRTLCAicGF5Iik7Cn0KCi8qID09PT09PT09PT09PSDYp9mE2KrZgdin2LnZhCA9PT09PT09PT09PT0gKi8KJCgic3ZjIikuYWRkRXZlbnRMaXN0ZW5lcigiY2xpY2siLCBlID0+IHsgY29uc3QgYiA9IGUudGFyZ2V0LmNsb3Nlc3QoImJ1dHRvbiIpOyBpZiAoIWIpIHJldHVybjsgc3RhdGUuc3ZjID0gYi5kYXRhc2V0LnM7IHJlbmRlclN2YygpOyByZW5kZXJTdW0oKTsgfSk7CiQoInFNaW51cyIpLmFkZEV2ZW50TGlzdGVuZXIoImNsaWNrIiwgKCkgPT4geyBzdGF0ZS5xdHkgPSBNYXRoLm1heCgxLCBzdGF0ZS5xdHkgLSAxKTsgcmVuZGVyU3VtKCk7IH0pOwokKCJxUGx1cyIpLmFkZEV2ZW50TGlzdGVuZXIoImNsaWNrIiwgKCkgPT4geyBzdGF0ZS5xdHkgPSBNYXRoLm1pbig1LCBzdGF0ZS5xdHkgKyAxKTsgcmVuZGVyU3VtKCk7IH0pOwpbWyJjaXRpZXMiLCJjaXR5Il0sWyJjdHlwZSIsImN0eXBlIl0sWyJwYXkiLCJwYXkiXV0uZm9yRWFjaCgoW2VsLGtleV0pID0+IHsKICAkKGVsKS5hZGRFdmVudExpc3RlbmVyKCJjbGljayIsIGUgPT4geyBjb25zdCBiID0gZS50YXJnZXQuY2xvc2VzdCgiYnV0dG9uIik7IGlmICghYikgcmV0dXJuOwogICAgc3RhdGVba2V5XSA9IGIuZGF0YXNldC52OyAkKGVsKS5xdWVyeVNlbGVjdG9yQWxsKCJidXR0b24iKS5mb3JFYWNoKHggPT4geC5zZXRBdHRyaWJ1dGUoImFyaWEtcHJlc3NlZCIsIHggPT09IGIpKTsgfSk7Cn0pOwokKCJsb2NhdGUiKS5hZGRFdmVudExpc3RlbmVyKCJjbGljayIsICgpID0+IHsKICBjb25zdCBidG4gPSAkKCJsb2NhdGUiKTsKICBpZiAoIW5hdmlnYXRvci5nZW9sb2NhdGlvbil7IGJ0bi50ZXh0Q29udGVudCA9ICLYrNmH2KfYstmDINmF2Kcg2YrYr9i52YUg2KrYrdiv2YrYryDYp9mE2YXZiNmC2LnYjCDYp9mE2LXZgiDYp9mE2LHYp9io2Lcg2KrYrdiqIjsgcmV0dXJuOyB9CiAgYnRuLnRleHRDb250ZW50ID0gItis2KfYsdmKINiq2K3Yr9mK2K8g2YXZiNmC2LnZg+KApiI7CiAgbmF2aWdhdG9yLmdlb2xvY2F0aW9uLmdldEN1cnJlbnRQb3NpdGlvbihwID0+IHsKICAgIHN0YXRlLmxvYyA9IHtsYXQ6IHAuY29vcmRzLmxhdGl0dWRlLnRvRml4ZWQoNiksIGxuZzogcC5jb29yZHMubG9uZ2l0dWRlLnRvRml4ZWQoNil9OwogICAgYnRuLnRleHRDb250ZW50ID0gIuKckyDYqtmFINiq2K3Yr9mK2K8g2YXZiNmC2LnZgyI7IGJ0bi5jbGFzc0xpc3QuYWRkKCJvayIpOwogIH0sICgpID0+IHsgYnRuLnRleHRDb250ZW50ID0gItmF2Kcg2YLYr9ix2YbYpyDZhtit2K/YryDZhdmI2YLYudmD2Iwg2KfZhNi12YIg2LHYp9io2Lcg2KfZhNmF2YjZgti5INiq2K3YqiI7IH0sIHtlbmFibGVIaWdoQWNjdXJhY3k6dHJ1ZSwgdGltZW91dDoxNTAwMH0pOwp9KTsKCmNvbnN0IHRvZGF5ID0gbmV3IERhdGUoKTsgdG9kYXkuc2V0TWludXRlcyh0b2RheS5nZXRNaW51dGVzKCkgLSB0b2RheS5nZXRUaW1lem9uZU9mZnNldCgpKTsKY29uc3QgaXNvID0gdG9kYXkudG9JU09TdHJpbmcoKS5zbGljZSgwLDEwKTsKJCgiZGF0ZSIpLm1pbiA9IGlzbzsgJCgiZGF0ZSIpLnZhbHVlID0gaXNvOwokKCJkYXRlIikuYWRkRXZlbnRMaXN0ZW5lcigiY2hhbmdlIiwgKCkgPT4gewogIGNvbnN0IGQgPSBuZXcgRGF0ZSgkKCJkYXRlIikudmFsdWUgKyAiVDEyOjAwOjAwIik7CiAgJCgiZGF0ZUhpbnQiKS50ZXh0Q29udGVudCA9IGQuZ2V0RGF5KCkgPT09IDUKICAgID8gIuKaoO+4jyDYp9mE2KzZhdi52Kkg2KXYrNin2LLYqdiMINio2YbYo9mD2K8g2YTZgyDYo9mC2LHYqCDZhdmI2LnYryDYqNi52K/Zh9inLiIKICAgIDogItin2YTYqtmI2LXZitmEINiu2YTYp9mEIDI0INiz2KfYudipINmF2YYg2KrYo9mD2YrYryDYp9mE2LfZhNioLiDYp9mE2KzZhdi52Kkg2KXYrNin2LLYqS4iOwp9KTsKCi8qID09PT09PT09PT09PSDYp9mE2KXYsdiz2KfZhCA9PT09PT09PT09PT0gKi8KZnVuY3Rpb24gbm9ybVBob25lKHYpewogIHYgPSB2LnJlcGxhY2UoL1vZoC3ZqV0vZywgZCA9PiAi2aDZodmi2aPZpNml2abZp9mo2akiLmluZGV4T2YoZCkpLnJlcGxhY2UoL1teXGRdL2csIiIpOwogIGlmICh2LnN0YXJ0c1dpdGgoIjk2NiIpKSB2ID0gIjAiICsgdi5zbGljZSgzKTsKICBpZiAodi5zdGFydHNXaXRoKCI1IikgJiYgdi5sZW5ndGggPT09IDkpIHYgPSAiMCIgKyB2OwogIHJldHVybiB2Owp9CmZ1bmN0aW9uIGNvbGxlY3QoKXsKICBjb25zdCBtYXBMaW5rID0gc3RhdGUubG9jID8gYGh0dHBzOi8vbWFwcy5nb29nbGUuY29tLz9xPSR7c3RhdGUubG9jLmxhdH0sJHtzdGF0ZS5sb2MubG5nfWAgOiAkKCJtYXBMaW5rIikudmFsdWUudHJpbSgpOwogIHJldHVybiB7CiAgICBzZXJ2aWNlOiBzdGF0ZS5zdmMsIHF0eTogc3RhdGUucXR5LCBjb250YWluZXJObzogJCgiY29udGFpbmVyTm8iKS52YWx1ZS50cmltKCksCiAgICBjaXR5OiBzdGF0ZS5jaXR5LCBkaXN0cmljdDogJCgiZGlzdHJpY3QiKS52YWx1ZS50cmltKCksIG1hcExpbmssIGFkZHJlc3M6ICQoImFkZHJlc3MiKS52YWx1ZS50cmltKCksCiAgICBuYW1lOiAkKCJuYW1lIikudmFsdWUudHJpbSgpLCBwaG9uZTogbm9ybVBob25lKCQoInBob25lIikudmFsdWUpLCBjdXN0b21lclR5cGU6IHN0YXRlLmN0eXBlLAogICAgZGF0ZTogJCgiZGF0ZSIpLnZhbHVlLCBwYXltZW50OiBzdGF0ZS5wYXksIG5vdGVzOiAkKCJub3RlcyIpLnZhbHVlLnRyaW0oKSwgYWdyZWU6ICQoImFncmVlIikuY2hlY2tlZAogIH07Cn0KZnVuY3Rpb24gdmFsaWRhdGUobyl7CiAgaWYgKCFvLmNpdHkpIHJldHVybiBbItin2K7YqtixINin2YTZhdiv2YrZhtipLiIsICJzMiJdOwogIGlmIChvLmRpc3RyaWN0Lmxlbmd0aCA8IDIpIHJldHVybiBbItin2YPYqtioINin2YTYrdmKLiIsICJzMiJdOwogIGlmICghby5tYXBMaW5rICYmIG8uYWRkcmVzcy5sZW5ndGggPCA1KSByZXR1cm4gWyLYrdiv2K8g2YXZiNmC2LnZgyDYo9mIINin2YTYtdmCINix2KfYqNi3INin2YTZhdmI2YLYuS4iLCAiczIiXTsKICBpZiAoby5zZXJ2aWNlICE9PSAiUElDS1VQIiAmJiAhby5hZ3JlZSkgcmV0dXJuIFsi2KPZg9ivINil2YYg2KfZhNit2KfZiNmK2Kkg2YTZhNmG2YHYp9mK2KfYqiDYp9mE2KXZhti02KfYptmK2Kkg2YHZgti3LiIsICJzMyJdOwogIGlmIChvLm5hbWUubGVuZ3RoIDwgMikgcmV0dXJuIFsi2KfZg9iq2Kgg2KfYs9mF2YMuIiwgInM0Il07CiAgaWYgKCEvXjA1XGR7OH0kLy50ZXN0KG8ucGhvbmUpKSByZXR1cm4gWyLYp9mD2KrYqCDYsdmC2YUg2KzZiNin2YQg2LXYrdmK2K0g2YrYqNiv2KMg2KjZgCAwNS4iLCAiczQiXTsKICBpZiAoIW8uY3VzdG9tZXJUeXBlKSByZXR1cm4gWyLYp9iu2KrYsTog2YHYsdivINij2Ygg2YXZgtin2YjZhC/YtNix2YPYqS4iLCAiczQiXTsKICBpZiAoIW8uZGF0ZSkgcmV0dXJuIFsi2KfYrtiq2LEg2KfZhNmF2YjYudivLiIsICJzNCJdOwogIGlmICghby5wYXltZW50KSByZXR1cm4gWyLYp9iu2KrYsSDYt9ix2YrZgtipINin2YTYr9mB2LkuIiwgInM0Il07CiAgcmV0dXJuIG51bGw7Cn0KZnVuY3Rpb24gd2FUZXh0KG8sIGlkKXsKICBjb25zdCBzID0gU0VSVklDRVMuZmluZCh4ID0+IHguaWQgPT09IG8uc2VydmljZSksIHQgPSB0b3RhbHMoKTsKICByZXR1cm4gWwogICAgYNi32YTYqCDYrNiv2YrYryDZhdmGINi12YHYrdipINir2LHYp9ihINin2YTYsdmF2KfZhCR7aWQgPyAiINix2YLZhSAiICsgaWQgOiAiIn1gLAogICAgYNin2YTYrtiv2YXYqTogJHtzLm5hbWV9ICgyMCDZitin2LHYr9ipKSDDlyAke28ucXR5fWAsCiAgICBvLmNvbnRhaW5lck5vID8gYNix2YLZhSDYp9mE2K3Yp9mI2YrYqTogJHtvLmNvbnRhaW5lck5vfWAgOiAiIiwKICAgIHQuc3ViID8gYNin2YTZhdio2YTYuiDZgtio2YQg2KfZhNi22LHZitio2Kk6ICR7Zm10KHQuc3ViKX0g2LHZitin2YRgIDogItin2YTZhdio2YTYujog2YXYtNmF2YjZhCIsCiAgICB0LnN1YiA/IGDYp9mE2LbYsdmK2KjYqSAxNSU6ICR7Zm10KHQudmF0KX0g2LHZitin2YRgIDogIiIsCiAgICB0LnN1YiA/IGDYp9mE2KXYrNmF2KfZhNmKINi02KfZhdmEINin2YTYttix2YrYqNipOiAke2ZtdCh0LnRvdGFsKX0g2LHZitin2YRgIDogIiIsCiAgICBg2KfZhNin2LPZhTogJHtvLm5hbWV9YCwgYNin2YTYrNmI2KfZhDogJHtvLnBob25lfWAsIGDYp9mE2LnZhdmK2YQ6ICR7by5jdXN0b21lclR5cGV9YCwKICAgIGDYp9mE2YXYr9mK2YbYqSDZiNin2YTYrdmKOiAke28uY2l0eX0gLSAke28uZGlzdHJpY3R9YCwKICAgIG8ubWFwTGluayA/IGDYp9mE2YXZiNmC2Lk6ICR7by5tYXBMaW5rfWAgOiAiIiwgby5hZGRyZXNzID8gYNin2YTZiNi12YE6ICR7by5hZGRyZXNzfWAgOiAiIiwKICAgIGDYp9mE2YXZiNi52K86ICR7by5kYXRlfWAsIGDYp9mE2K/Zgdi5OiAke28ucGF5bWVudH1gLCBvLm5vdGVzID8gYNmF2YTYp9it2LjYp9iqOiAke28ubm90ZXN9YCA6ICIiCiAgXS5maWx0ZXIoQm9vbGVhbikuam9pbigiXG4iKTsKfQpmdW5jdGlvbiBzaG93RG9uZShvLCBpZCl7CiAgJCgiYXBwIikuc3R5bGUuZGlzcGxheSA9ICJub25lIjsgJCgiYmFyIikuc3R5bGUuZGlzcGxheSA9ICJub25lIjsKICAkKCJkb25lIikuc3R5bGUuZGlzcGxheSA9ICJibG9jayI7CiAgJCgiZG9uZUlkIikudGV4dENvbnRlbnQgPSBpZCA/IGDYsdmC2YUg2LfZhNio2YM6ICR7aWR9YCA6ICIiOwogIGNvbnN0IHdhID0gYGh0dHBzOi8vd2EubWUvJHtDT05GSUcuV0hBVFNBUFB9P3RleHQ9JHtlbmNvZGVVUklDb21wb25lbnQod2FUZXh0KG8sIGlkKSl9YDsKICAkKCJ3YUxpbmsiKS5ocmVmID0gd2E7CiAgd2luZG93LnNjcm9sbFRvKDAsMCk7CiAgLy8g2KrYrdmI2YrZhCDYp9mE2LnZhdmK2YQg2KrZhNmC2KfYptmK2KfZiyDZhNmE2YjYp9iq2LPYp9ioINmI2KfZhNix2LPYp9mE2Kkg2KzYp9mH2LLYqdiMINmI2YXYpyDYudmE2YrZhyDYpdmE2Kcg2YrYtti62Lcg2KXYsdiz2KfZhAogIHNldFRpbWVvdXQoKCkgPT4gewogICAgaWYgKGluR2FzKCkpIHsgdHJ5IHsgd2luZG93Lm9wZW4od2EsICJfdG9wIik7IH0gY2F0Y2goXykge30gfQogICAgZWxzZSB3aW5kb3cubG9jYXRpb24uaHJlZiA9IHdhOwogIH0sIDE1MDApOwp9CmZ1bmN0aW9uIGZhaWwoKXsKICAkKCJlcnIiKS50ZXh0Q29udGVudCA9ICLYtdin2LEg2K7Yt9ijINmB2Yog2KfZhNil2LHYs9in2YQuINin2LbYuti3INmF2LHYqSDYq9in2YbZitip2Iwg2KPZiCDYsdin2LPZhNmG2Kcg2LnZhNmJINmI2KfYqtiz2KfYqCAwNTc5MzIyMDMwLiI7CiAgJCgic3VibWl0IikuZGlzYWJsZWQgPSBmYWxzZTsgJCgic3VibWl0IikudGV4dENvbnRlbnQgPSAi2KPYsdiz2YQg2KfZhNi32YTYqCI7IHN0YXRlLnNlbmRpbmcgPSBmYWxzZTsKfQokKCJzdWJtaXQiKS5hZGRFdmVudExpc3RlbmVyKCJjbGljayIsIGFzeW5jICgpID0+IHsKICBpZiAoc3RhdGUuc2VuZGluZykgcmV0dXJuOwogIGNvbnN0IG8gPSBjb2xsZWN0KCksIHYgPSB2YWxpZGF0ZShvKTsKICAkKCJlcnIiKS50ZXh0Q29udGVudCA9IHYgPyB2WzBdIDogIiI7CiAgaWYgKHYpeyAkKHZbMV0pLnNjcm9sbEludG9WaWV3KHtiZWhhdmlvcjoic21vb3RoIiwgYmxvY2s6InN0YXJ0In0pOyByZXR1cm47IH0KICBpZiAoIWluR2FzKCkgJiYgIUNPTkZJRy5FTkRQT0lOVCl7IHNob3dEb25lKG8sICIiKTsgcmV0dXJuOyB9IC8vINmI2LbYuSDZiNin2KrYs9in2Kgg2YHZgti3CiAgc3RhdGUuc2VuZGluZyA9IHRydWU7ICQoInN1Ym1pdCIpLmRpc2FibGVkID0gdHJ1ZTsgJCgic3VibWl0IikudGV4dENvbnRlbnQgPSAi2KzYp9ix2Yog2KfZhNil2LHYs9in2YTigKYiOwogIGlmIChpbkdhcygpKXsKICAgIGdvb2dsZS5zY3JpcHQucnVuLndpdGhTdWNjZXNzSGFuZGxlcihkID0+IGQgJiYgZC5vayA/IHNob3dEb25lKG8sIGQuaWQpIDogZmFpbCgpKS53aXRoRmFpbHVyZUhhbmRsZXIoZmFpbCkuc3VibWl0T3JkZXIobyk7CiAgICByZXR1cm47CiAgfQogIHRyeSB7CiAgICBjb25zdCByZXMgPSBhd2FpdCBmZXRjaChDT05GSUcuRU5EUE9JTlQsIHttZXRob2Q6IlBPU1QiLCBib2R5OkpTT04uc3RyaW5naWZ5KG8pfSk7CiAgICBjb25zdCBkYXRhID0gYXdhaXQgcmVzLmpzb24oKTsKICAgIGlmICghZGF0YS5vaykgdGhyb3cgbmV3IEVycm9yKGRhdGEuZXJyb3IgfHwgImZhaWxlZCIpOwogICAgc2hvd0RvbmUobywgZGF0YS5pZCk7CiAgfSBjYXRjaCAoXykgeyBmYWlsKCk7IH0KfSk7CgovKiA9PT09PT09PT09PT0g2KrYrdmF2YrZhCDYp9mE2KPYs9i52KfYsSDZhdmGINis2K/ZiNmEINmC2YjZgtmEID09PT09PT09PT09PSAqLwpmdW5jdGlvbiBhcHBseVByaWNlcyhwKXsgaWYgKHAgJiYgdHlwZW9mIHAgPT09ICJvYmplY3QiICYmIHAuTkVXICE9IG51bGwpeyBQID0gT2JqZWN0LmFzc2lnbih7fSwgUCwgcCk7IHJlbmRlckFsbCgpOyB9IH0KYXN5bmMgZnVuY3Rpb24gbG9hZFByaWNlcygpewogIGlmIChpbkdhcygpKXsgZ29vZ2xlLnNjcmlwdC5ydW4ud2l0aFN1Y2Nlc3NIYW5kbGVyKGFwcGx5UHJpY2VzKS5nZXRQcmljZXMoKTsgcmV0dXJuOyB9CiAgaWYgKCFDT05GSUcuRU5EUE9JTlQpIHJldHVybjsKICB0cnkgewogICAgY29uc3QgcmVzID0gYXdhaXQgZmV0Y2goQ09ORklHLkVORFBPSU5UICsgIj9hY3Rpb249cHJpY2VzIik7CiAgICBjb25zdCBkYXRhID0gYXdhaXQgcmVzLmpzb24oKTsKICAgIGlmIChkYXRhLm9rKSBhcHBseVByaWNlcyhkYXRhLnByaWNlcyk7CiAgfSBjYXRjaChfKSB7IC8qINmG2LPYqtiu2K/ZhSDYp9mE2KPYs9i52KfYsSDYp9mE2YXYttmF2YbYqSAqLyB9Cn0KcmVuZGVyQWxsKCk7IGxvYWRQcmljZXMoKTsKfSkoKTsKPC9zY3JpcHQ+CjwvYm9keT4KPC9odG1sPgo=';
