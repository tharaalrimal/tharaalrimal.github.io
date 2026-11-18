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
const SERVICES = {NEW: 'حاوية جديدة', SWAP: 'استبدال حاوية'};
const CITIES = ['الخبر', 'الظهران', 'الدمام', 'سيهات', 'القطيف', 'مدينة أخرى'];

// الأسعار نهائية بدون ضريبة
const VAT = 0;

// الأسعار الأولية (تُنسخ لورقة "الأسعار" عند setup، وبعدها تعدّل من الورقة مباشرة وتظهر في الصفحة)
const SEED_PRICES = [
  ['NEW',       'حاوية جديدة 20 ياردة (شاملة التوصيل والسحب)', 400, 'ريال'],
  ['SWAP',      'استبدال حاوية (نسحب الممتلئة ونحط فاضية)',     400, 'ريال'],
  ['EXTRA_DAY', 'سعر اليوم الإضافي',                              30,  'ريال'],
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
              'المبلغ','الضريبة','الإجمالي','ملاحظات العميل',
              'الحالة','المسؤول','تاريخ التسليم الفعلي','رقم الحاوية المسلّمة','ملاحظات داخلية'];
  o.getRange(1,1,1,oh.length).setValues([oh]).setFontWeight('bold')
   .setBackground('#0F2439').setFontColor('#FFC107').setWrap(true);
  o.setFrozenRows(1);
  o.setRightToLeft(true);
  o.setColumnWidths(1, oh.length, 120);
  o.setColumnWidth(11, 220);
  o.hideColumns(16); // عمود الضريبة (غير مستخدم)
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

  // حذف بند السحب إن وُجد (أُلغيت خدمة السحب)
  const pv = p.getRange(1, 1, Math.max(p.getLastRow(), 1), 1).getValues();
  for (let i = pv.length - 1; i >= 1; i--) if (String(pv[i][0]) === 'PICKUP') p.deleteRow(i + 1);

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
    ['مبيعات مؤكدة (ريال)', `=SUM(${q}!Q2:Q)-SUMIF(${q}!S2:S,"ملغي",${q}!Q2:Q)-SUMIF(${q}!S2:S,"جديد",${q}!Q2:Q)`]
  ];
  s.getRange(3,1,rows.length,2).setValues(rows);
  s.getRange(3,1,rows.length,1).setFontWeight('bold');
  s.setColumnWidth(1, 260);

  s.getRange('D2').setValue('الطلبات حسب المدينة').setFontWeight('bold');
  s.getRange('D3').setFormula(`=IFERROR(QUERY(${q}!I2:Q,"select I, count(I), sum(Q) where I is not null group by I order by count(I) desc label count(I) 'العدد', sum(Q) 'القيمة'",0),"لا توجد طلبات بعد")`);
  s.getRange('H2').setValue('الطلبات حسب الخدمة').setFontWeight('bold');
  s.getRange('H3').setFormula(`=IFERROR(QUERY(${q}!C2:D,"select C, count(C), sum(D) where C is not null group by C label count(C) 'الطلبات', sum(D) 'الحاويات'",0),"")`);

  // الحاويات عند العملاء مع الأيام والمبلغ الإضافي (يعتمد على تاريخ التسليم الفعلي، وإذا فاضي يأخذ الموعد المطلوب)
  s.getRange('A13').setValue('الحاويات عند العملاء الآن').setFontSize(13).setFontWeight('bold');
  const hh = ['رقم الطلب','الاسم','الجوال','الموقع','رقم الحاوية','تاريخ التسليم','الأيام','أيام إضافية','مبلغ إضافي'];
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
      id, now, SERVICES[o.service], qty, o.containerNo ? "'" + clean_(o.containerNo, 20) : '', clean_(o.name, 60), "'" + o.phone,
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
             `المبلغ: ${total} ريال`, '',
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
const PAGE_B64 = 'PCFkb2N0eXBlIGh0bWw+CjxodG1sIGxhbmc9ImFyIiBkaXI9InJ0bCI+CjxoZWFkPgo8bWV0YSBjaGFyc2V0PSJ1dGYtOCI+CjxtZXRhIG5hbWU9InZpZXdwb3J0IiBjb250ZW50PSJ3aWR0aD1kZXZpY2Utd2lkdGgsaW5pdGlhbC1zY2FsZT0xLHZpZXdwb3J0LWZpdD1jb3ZlciI+Cjx0aXRsZT7Yq9ix2KfYoSDYp9mE2LHZhdin2YQgfCDYp9i32YTYqCDYrdin2YjZitipINmF2K7ZhNmB2KfYqiDYqNmG2KfYoSDZgdmKINin2YTZhdmG2LfZgtipINin2YTYtNix2YLZitipPC90aXRsZT4KPG1ldGEgbmFtZT0iZGVzY3JpcHRpb24iIGNvbnRlbnQ9Itit2KfZiNmK2Kkg2YXYrtmE2YHYp9iqINio2YbYp9ihIDIwINmK2KfYsdiv2Kkg2YXZhiDYq9ix2KfYoSDYp9mE2LHZhdin2YQ6IDQwMCDYsdmK2KfZhCDZhNmF2K/YqSAxMCDYo9mK2KfZhSDYtNin2YXZhNipINin2YTYqtmI2LXZitmEINmI2KfZhNiz2K3YqNiMINmF2YYg2KfZhNiu2KjYsSDYpdmE2Ykg2KfZhNmC2LfZitmBLiDYp9i32YTYqCDYp9mE2KLZhiDZiNin2YTYqtmI2LXZitmEINiu2YTYp9mEIDI0INiz2KfYudipLiI+CjxtZXRhIHByb3BlcnR5PSJvZzp0aXRsZSIgY29udGVudD0i2KvYsdin2KEg2KfZhNix2YXYp9mEIHwg2K3Yp9mI2YrYqSDZhdiu2YTZgdin2Kog2KjZhtin2KEgMjAg2YrYp9ix2K/YqSI+CjxtZXRhIHByb3BlcnR5PSJvZzpkZXNjcmlwdGlvbiIgY29udGVudD0iNDAwINix2YrYp9mEINmE2YXYr9ipIDEwINij2YrYp9mFINi02KfZhdmE2Kkg2KfZhNiq2YjYtdmK2YQg2YjYp9mE2LPYrdio2Iwg2YXZhiDYp9mE2K7YqNixINil2YTZiSDYp9mE2YLYt9mK2YEuINin2LfZhNioINio2K/ZgtmK2YLYqS4iPgo8bWV0YSBwcm9wZXJ0eT0ib2c6aW1hZ2UiIGNvbnRlbnQ9Imh0dHBzOi8vd2FzZWVtYW5hbHl0aWMtZG90Y29tLmdpdGh1Yi5pby90aGFyYS1hbHJpbWFsL2Fzc2V0cy9vZy5qcGciPgo8bWV0YSBwcm9wZXJ0eT0ib2c6dXJsIiBjb250ZW50PSJodHRwczovL3dhc2VlbWFuYWx5dGljLWRvdGNvbS5naXRodWIuaW8vdGhhcmEtYWxyaW1hbC8iPgo8bWV0YSBwcm9wZXJ0eT0ib2c6dHlwZSIgY29udGVudD0id2Vic2l0ZSI+CjxtZXRhIG5hbWU9InRoZW1lLWNvbG9yIiBjb250ZW50PSIjMEYyNDM5Ij4KPGxpbmsgcmVsPSJpY29uIiBocmVmPSJodHRwczovL3dhc2VlbWFuYWx5dGljLWRvdGNvbS5naXRodWIuaW8vdGhhcmEtYWxyaW1hbC9hc3NldHMvbG9nby5qcGciPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20iPgo8bGluayByZWw9InByZWNvbm5lY3QiIGhyZWY9Imh0dHBzOi8vZm9udHMuZ3N0YXRpYy5jb20iIGNyb3Nzb3JpZ2luPgo8bGluayBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tL2NzczI/ZmFtaWx5PVRhamF3YWw6d2dodEA0MDA7NTAwOzcwMDs4MDAmZGlzcGxheT1zd2FwIiByZWw9InN0eWxlc2hlZXQiPgo8c3R5bGU+Cjpyb290ewogIC0tbmF2eTojMEYyNDM5OyAtLW5hdnktMjojMUEzNjU0OyAtLXllbGxvdzojRkZDMTA3OyAtLXllbGxvdy0yOiNGRkQ1NEE7CiAgLS1iZzojRjVGNkY4OyAtLWNhcmQ6I0ZGRkZGRjsgLS1pbms6IzE0MjAyQjsgLS1tdXRlZDojNUQ2Qjc4OyAtLWxpbmU6I0UyRTZFQjsKICAtLW9rOiMxRThFNUE7IC0tZXJyOiNDMDM5MkI7IC0tcmFkaXVzOjE2cHg7CiAgLS1zdHJpcGVzOnJlcGVhdGluZy1saW5lYXItZ3JhZGllbnQoLTQ1ZGVnLHZhcigtLXllbGxvdykgMCAxNHB4LHZhcigtLW5hdnkpIDE0cHggMjhweCk7Cn0KKntib3gtc2l6aW5nOmJvcmRlci1ib3g7bWFyZ2luOjA7cGFkZGluZzowfQpodG1sLGJvZHl7YmFja2dyb3VuZDp2YXIoLS1iZyl9CmJvZHl7Zm9udC1mYW1pbHk6J1RhamF3YWwnLHN5c3RlbS11aSxzYW5zLXNlcmlmO2NvbG9yOnZhcigtLWluayk7bGluZS1oZWlnaHQ6MS42O3BhZGRpbmctYm90dG9tOjExNnB4Oy13ZWJraXQtdGFwLWhpZ2hsaWdodC1jb2xvcjp0cmFuc3BhcmVudH0KYnV0dG9uLGlucHV0LHNlbGVjdCx0ZXh0YXJlYXtmb250OmluaGVyaXQ7Y29sb3I6aW5oZXJpdH0KLndyYXB7bWF4LXdpZHRoOjcyMHB4O21hcmdpbjowIGF1dG87cGFkZGluZzowIDE2cHh9CgovKiDYp9mE2YfZitiv2LEgKi8KaGVhZGVye2JhY2tncm91bmQ6dmFyKC0tbmF2eSk7Y29sb3I6I2ZmZjtwYWRkaW5nOjE2cHggMCAwfQouYnJhbmR7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6MTJweH0KLmJyYW5kIGltZ3t3aWR0aDo2MnB4O2hlaWdodDphdXRvO2ZsZXg6bm9uZX0KLmJyYW5kIGgxe2ZvbnQtc2l6ZToyOHB4O2ZvbnQtd2VpZ2h0OjgwMDtsaW5lLWhlaWdodDoxLjE7Y29sb3I6dmFyKC0teWVsbG93KX0KLmJyYW5kIHB7Zm9udC1zaXplOjE0cHg7b3BhY2l0eTouOX0KLnRydXN0e2Rpc3BsYXk6ZmxleDtnYXA6OHB4O2ZsZXgtd3JhcDp3cmFwO21hcmdpbjoxNHB4IDAgMTZweH0KLnRydXN0IHNwYW57YmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC4wOCk7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDI1NSwyNTUsMjU1LC4xOCk7Ym9yZGVyLXJhZGl1czo5OTlweDtwYWRkaW5nOjRweCAxMnB4O2ZvbnQtc2l6ZToxM3B4fQouc3RyaXBlc3toZWlnaHQ6MTJweDtiYWNrZ3JvdW5kOnZhcigtLXN0cmlwZXMpfQoKLyog2KfZhNio2LfYp9mC2Kkg2KfZhNix2KbZitiz2YrYqSAqLwouaGVyb3twb3NpdGlvbjpyZWxhdGl2ZTttYXJnaW4tdG9wOjE2cHg7Ym9yZGVyLXJhZGl1czp2YXIoLS1yYWRpdXMpO292ZXJmbG93OmhpZGRlbjtiYWNrZ3JvdW5kOnZhcigtLW5hdnkpO2JveC1zaGFkb3c6MCA2cHggMjJweCByZ2JhKDE1LDM2LDU3LC4xOCl9Ci5oZXJvIGltZ3tkaXNwbGF5OmJsb2NrO3dpZHRoOjEwMCU7aGVpZ2h0OmF1dG87YXNwZWN0LXJhdGlvOjE2LzEwO29iamVjdC1maXQ6Y292ZXI7b2JqZWN0LXBvc2l0aW9uOmNlbnRlciA1NSV9Ci5iYWRnZXtwb3NpdGlvbjphYnNvbHV0ZTt0b3A6MTJweDtsZWZ0OjEycHg7d2lkdGg6MTE4cHg7aGVpZ2h0OjExOHB4O2JvcmRlci1yYWRpdXM6NTAlO2JhY2tncm91bmQ6dmFyKC0teWVsbG93KTtib3JkZXI6NXB4IHNvbGlkIHZhcigtLW5hdnkpOwogIGRpc3BsYXk6ZmxleDtmbGV4LWRpcmVjdGlvbjpjb2x1bW47YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXI7Y29sb3I6dmFyKC0tbmF2eSk7dGV4dC1hbGlnbjpjZW50ZXI7bGluZS1oZWlnaHQ6MS4wNX0KLmJhZGdlIGJ7Zm9udC1zaXplOjM4cHg7Zm9udC13ZWlnaHQ6ODAwfQouYmFkZ2Ugc21hbGx7Zm9udC1zaXplOjEzcHg7Zm9udC13ZWlnaHQ6NzAwfQouYmFkZ2UgaXtmb250LXN0eWxlOm5vcm1hbDtmb250LXNpemU6MTFweDtmb250LXdlaWdodDo1MDA7bWFyZ2luLXRvcDozcHh9Ci5oZXJvIC5pbmZve2NvbG9yOiNmZmY7cGFkZGluZzoxNHB4IDE2cHggMTZweH0KLmhlcm8gaDJ7Zm9udC1zaXplOjIycHg7Zm9udC13ZWlnaHQ6ODAwO2NvbG9yOnZhcigtLXllbGxvdyl9Ci5oZXJvIHB7Zm9udC1zaXplOjE1cHg7b3BhY2l0eTouOTJ9CgovKiDYp9mE2K7Yt9mI2KfYqiAqLwpzZWN0aW9uLnN0ZXB7YmFja2dyb3VuZDp2YXIoLS1jYXJkKTtib3JkZXItcmFkaXVzOnZhcigtLXJhZGl1cyk7cGFkZGluZzoxOHB4IDE2cHg7bWFyZ2luLXRvcDoxNnB4O2JvcmRlcjoxcHggc29saWQgdmFyKC0tbGluZSl9CnNlY3Rpb24uc3RlcCBoMntmb250LXNpemU6MTlweDtmb250LXdlaWdodDo4MDA7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6MTBweDttYXJnaW4tYm90dG9tOjEycHh9Ci5ue3dpZHRoOjI4cHg7aGVpZ2h0OjI4cHg7Ym9yZGVyLXJhZGl1czo1MCU7YmFja2dyb3VuZDp2YXIoLS1uYXZ5KTtjb2xvcjp2YXIoLS15ZWxsb3cpO2Rpc3BsYXk6aW5saW5lLWZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXI7Zm9udC1zaXplOjE1cHg7ZmxleDpub25lfQouaGludHtmb250LXNpemU6MTNweDtjb2xvcjp2YXIoLS1tdXRlZCk7bWFyZ2luLXRvcDo2cHh9Cgouc3Zje2Rpc3BsYXk6Z3JpZDtnYXA6MTBweH0KLnN2YyBidXR0b257ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtqdXN0aWZ5LWNvbnRlbnQ6c3BhY2UtYmV0d2VlbjtnYXA6MTJweDt0ZXh0LWFsaWduOnJpZ2h0O2JhY2tncm91bmQ6I2ZmZjtib3JkZXI6MnB4IHNvbGlkIHZhcigtLWxpbmUpO2JvcmRlci1yYWRpdXM6MTRweDtwYWRkaW5nOjEycHggMTRweDtjdXJzb3I6cG9pbnRlcn0KLnN2YyBidXR0b24gYntkaXNwbGF5OmJsb2NrO2ZvbnQtc2l6ZToxN3B4fQouc3ZjIGJ1dHRvbiBzcGFue2Rpc3BsYXk6YmxvY2s7Zm9udC1zaXplOjEzcHg7Y29sb3I6dmFyKC0tbXV0ZWQpfQouc3ZjIGJ1dHRvbiBlbXtmb250LXN0eWxlOm5vcm1hbDtmb250LXdlaWdodDo4MDA7Zm9udC1zaXplOjE3cHg7d2hpdGUtc3BhY2U6bm93cmFwO2NvbG9yOnZhcigtLW5hdnkpfQouc3ZjIGJ1dHRvblthcmlhLXByZXNzZWQ9InRydWUiXXtib3JkZXItY29sb3I6dmFyKC0tbmF2eSk7YmFja2dyb3VuZDojRkZGOEUxO2JveC1zaGFkb3c6aW5zZXQgMCAwIDAgMXB4IHZhcigtLW5hdnkpfQoKLnF0eXtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpzcGFjZS1iZXR3ZWVuO21hcmdpbi10b3A6MTRweDtnYXA6MTJweH0KLnF0eSBsYWJlbHtmb250LXdlaWdodDo3MDB9Ci5zdGVwcGVye2Rpc3BsYXk6ZmxleDthbGlnbi1pdGVtczpjZW50ZXI7Ym9yZGVyOjJweCBzb2xpZCB2YXIoLS1saW5lKTtib3JkZXItcmFkaXVzOjEycHg7b3ZlcmZsb3c6aGlkZGVufQouc3RlcHBlciBidXR0b257d2lkdGg6NDRweDtoZWlnaHQ6NDRweDtib3JkZXI6MDtiYWNrZ3JvdW5kOiNGMEYyRjU7Zm9udC1zaXplOjIycHg7Zm9udC13ZWlnaHQ6NzAwO2N1cnNvcjpwb2ludGVyfQouc3RlcHBlciBvdXRwdXR7bWluLXdpZHRoOjQ0cHg7dGV4dC1hbGlnbjpjZW50ZXI7Zm9udC13ZWlnaHQ6ODAwO2ZvbnQtc2l6ZToxOHB4fQoKLmNoaXBze2Rpc3BsYXk6ZmxleDtmbGV4LXdyYXA6d3JhcDtnYXA6OHB4fQouY2hpcHMgYnV0dG9ue2JvcmRlcjoycHggc29saWQgdmFyKC0tbGluZSk7YmFja2dyb3VuZDojZmZmO2JvcmRlci1yYWRpdXM6OTk5cHg7cGFkZGluZzo4cHggMTZweDtmb250LXdlaWdodDo3MDA7Y3Vyc29yOnBvaW50ZXJ9Ci5jaGlwcyBidXR0b25bYXJpYS1wcmVzc2VkPSJ0cnVlIl17YmFja2dyb3VuZDp2YXIoLS1uYXZ5KTtjb2xvcjojZmZmO2JvcmRlci1jb2xvcjp2YXIoLS1uYXZ5KX0KCi5maWVsZHttYXJnaW4tdG9wOjEycHh9Ci5maWVsZCBsYWJlbHtkaXNwbGF5OmJsb2NrO2ZvbnQtd2VpZ2h0OjcwMDttYXJnaW4tYm90dG9tOjZweDtmb250LXNpemU6MTVweH0KLmZpZWxkIGlucHV0LC5maWVsZCB0ZXh0YXJlYXt3aWR0aDoxMDAlO2JvcmRlcjoycHggc29saWQgdmFyKC0tbGluZSk7Ym9yZGVyLXJhZGl1czoxMnB4O3BhZGRpbmc6MTFweCAxMnB4O2JhY2tncm91bmQ6I2ZmZjtmb250LXNpemU6MTZweH0KLmZpZWxkIGlucHV0OmZvY3VzLC5maWVsZCB0ZXh0YXJlYTpmb2N1c3tvdXRsaW5lOm5vbmU7Ym9yZGVyLWNvbG9yOnZhcigtLW5hdnkpfQouZmllbGQgdGV4dGFyZWF7bWluLWhlaWdodDo3NnB4O3Jlc2l6ZTp2ZXJ0aWNhbH0KLnJvdzJ7ZGlzcGxheTpncmlkO2dyaWQtdGVtcGxhdGUtY29sdW1uczoxZnIgMWZyO2dhcDoxMHB4fQoubG9jYnRue3dpZHRoOjEwMCU7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtqdXN0aWZ5LWNvbnRlbnQ6Y2VudGVyO2dhcDo4cHg7YmFja2dyb3VuZDp2YXIoLS1uYXZ5KTtjb2xvcjojZmZmO2JvcmRlcjowO2JvcmRlci1yYWRpdXM6MTJweDtwYWRkaW5nOjEycHg7Zm9udC13ZWlnaHQ6NzAwO2N1cnNvcjpwb2ludGVyO21hcmdpbi10b3A6MTJweH0KLmxvY2J0bi5va3tiYWNrZ3JvdW5kOnZhcigtLW9rKX0KLm9ye3RleHQtYWxpZ246Y2VudGVyO2NvbG9yOnZhcigtLW11dGVkKTtmb250LXNpemU6MTNweDttYXJnaW4tdG9wOjhweH0KCi53YXN0ZXtkaXNwbGF5OmdyaWQ7Z3JpZC10ZW1wbGF0ZS1jb2x1bW5zOjFmciAxZnI7Z2FwOjEwcHh9Ci53YXN0ZSBkaXZ7Ym9yZGVyLXJhZGl1czoxMnB4O3BhZGRpbmc6MTJweH0KLndhc3RlIC55ZXN7YmFja2dyb3VuZDojRTdGNkVFO2JvcmRlcjoxcHggc29saWQgI0JGRTVEMH0KLndhc3RlIC5ub3tiYWNrZ3JvdW5kOiNGQkVBRUE7Ym9yZGVyOjFweCBzb2xpZCAjRjFDNkM2fQoud2FzdGUgaDN7Zm9udC1zaXplOjE1cHg7bWFyZ2luLWJvdHRvbTo0cHh9Ci53YXN0ZSB1bHtsaXN0LXN0eWxlOm5vbmU7Zm9udC1zaXplOjE0cHh9Ci53YXN0ZSBsaTo6YmVmb3Jle2NvbnRlbnQ6IuKAoiAifQouYWdyZWV7ZGlzcGxheTpmbGV4O2dhcDoxMHB4O2FsaWduLWl0ZW1zOmZsZXgtc3RhcnQ7bWFyZ2luLXRvcDoxMnB4O2ZvbnQtd2VpZ2h0OjcwMDtmb250LXNpemU6MTVweDtjdXJzb3I6cG9pbnRlcn0KLmFncmVlIGlucHV0e3dpZHRoOjIycHg7aGVpZ2h0OjIycHg7bWFyZ2luLXRvcDoycHg7YWNjZW50LWNvbG9yOnZhcigtLW5hdnkpO2ZsZXg6bm9uZX0KCi5zdW17Ym9yZGVyLXRvcDoxcHggZGFzaGVkIHZhcigtLWxpbmUpO21hcmdpbi10b3A6NHB4fQouc3VtIGRpdntkaXNwbGF5OmZsZXg7anVzdGlmeS1jb250ZW50OnNwYWNlLWJldHdlZW47cGFkZGluZzo2cHggMDtmb250LXNpemU6MTVweH0KLnN1bSAudG90e2ZvbnQtd2VpZ2h0OjgwMDtmb250LXNpemU6MThweDtib3JkZXItdG9wOjFweCBzb2xpZCB2YXIoLS1saW5lKTttYXJnaW4tdG9wOjRweDtwYWRkaW5nLXRvcDoxMHB4fQoubm90ZXtiYWNrZ3JvdW5kOiNGRkY4RTE7Ym9yZGVyLXJhZGl1czoxMnB4O3BhZGRpbmc6MTBweCAxMnB4O2ZvbnQtc2l6ZToxNHB4O21hcmdpbi10b3A6MTBweH0KLmVycntjb2xvcjp2YXIoLS1lcnIpO2ZvbnQtd2VpZ2h0OjcwMDttaW4taGVpZ2h0OjEuNmVtO21hcmdpbi10b3A6MTBweDt0ZXh0LWFsaWduOmNlbnRlcn0KCi8qINin2YTYtNix2YrYtyDYp9mE2LPZgdmE2YogKi8KLmJhcntwb3NpdGlvbjpmaXhlZDtpbnNldDphdXRvIDAgMCAwO2JhY2tncm91bmQ6I2ZmZjtib3JkZXItdG9wOjFweCBzb2xpZCB2YXIoLS1saW5lKTtib3gtc2hhZG93OjAgLTZweCAyMHB4IHJnYmEoMCwwLDAsLjA4KTtwYWRkaW5nOjEwcHggMCBjYWxjKDEwcHggKyBlbnYoc2FmZS1hcmVhLWluc2V0LWJvdHRvbSkpO3otaW5kZXg6MTB9Ci5iYXIgLndyYXB7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6MTJweH0KLmJhciAudHtmbGV4OjE7bGluZS1oZWlnaHQ6MS4yNX0KLmJhciAudCBzbWFsbHtkaXNwbGF5OmJsb2NrO2NvbG9yOnZhcigtLW11dGVkKTtmb250LXNpemU6MTJweH0KLmJhciAudCBie2ZvbnQtc2l6ZToyMHB4fQouYmFyIGJ1dHRvbntiYWNrZ3JvdW5kOnZhcigtLXllbGxvdyk7Y29sb3I6dmFyKC0tbmF2eSk7Ym9yZGVyOjA7Ym9yZGVyLXJhZGl1czoxNHB4O3BhZGRpbmc6MTNweCAyMnB4O2ZvbnQtd2VpZ2h0OjgwMDtmb250LXNpemU6MTdweDtjdXJzb3I6cG9pbnRlcjtib3gtc2hhZG93OjAgM3B4IDAgI0M5OTUwMH0KLmJhciBidXR0b246ZGlzYWJsZWR7b3BhY2l0eTouNn0KCi8qINio2LnYryDYp9mE2KXYsdiz2KfZhCAqLwouZG9uZXtkaXNwbGF5Om5vbmU7dGV4dC1hbGlnbjpjZW50ZXI7cGFkZGluZy10b3A6MjhweH0KLmRvbmUgLmlje3dpZHRoOjg0cHg7aGVpZ2h0Ojg0cHg7Ym9yZGVyLXJhZGl1czo1MCU7YmFja2dyb3VuZDp2YXIoLS1vayk7Y29sb3I6I2ZmZjtmb250LXNpemU6NDZweDtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXI7bWFyZ2luOjAgYXV0byAxMnB4fQouZG9uZSBoMntmb250LXNpemU6MjZweH0KLmRvbmUgcHtjb2xvcjp2YXIoLS1tdXRlZCk7bWFyZ2luLXRvcDo2cHh9Ci5kb25lIC5pZHtmb250LXdlaWdodDo4MDA7Zm9udC1zaXplOjIwcHg7Y29sb3I6dmFyKC0tbmF2eSk7bWFyZ2luLXRvcDo4cHh9Ci53YXtkaXNwbGF5OmlubGluZS1mbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6OHB4O2JhY2tncm91bmQ6IzI1RDM2Njtjb2xvcjojZmZmO3RleHQtZGVjb3JhdGlvbjpub25lO2ZvbnQtd2VpZ2h0OjgwMDtib3JkZXItcmFkaXVzOjE0cHg7cGFkZGluZzoxNHB4IDIycHg7bWFyZ2luLXRvcDoxOHB4O2ZvbnQtc2l6ZToxN3B4fQoKZm9vdGVye3RleHQtYWxpZ246Y2VudGVyO2NvbG9yOnZhcigtLW11dGVkKTtmb250LXNpemU6MTNweDtwYWRkaW5nOjIycHggMCA4cHh9CmZvb3RlciBhe2NvbG9yOnZhcigtLW5hdnkpO2ZvbnQtd2VpZ2h0OjcwMH0KQG1lZGlhIChtYXgtd2lkdGg6NDIwcHgpey5iYWRnZXt3aWR0aDoxMDBweDtoZWlnaHQ6MTAwcHh9LmJhZGdlIGJ7Zm9udC1zaXplOjMycHh9Lndhc3Rle2dyaWQtdGVtcGxhdGUtY29sdW1uczoxZnJ9fQo8L3N0eWxlPgo8L2hlYWQ+Cjxib2R5PgoKPGhlYWRlcj4KICA8ZGl2IGNsYXNzPSJ3cmFwIj4KICAgIDxkaXYgY2xhc3M9ImJyYW5kIj4KICAgICAgPGltZyBzcmM9Imh0dHBzOi8vd2FzZWVtYW5hbHl0aWMtZG90Y29tLmdpdGh1Yi5pby90aGFyYS1hbHJpbWFsL2Fzc2V0cy9lbWJsZW0uanBnIiBhbHQ9Iti02LnYp9ixINir2LHYp9ihINin2YTYsdmF2KfZhCIgd2lkdGg9IjYyIiBoZWlnaHQ9IjU5Ij4KICAgICAgPGRpdj4KICAgICAgICA8aDE+2KvYsdin2KEg2KfZhNix2YXYp9mEPC9oMT4KICAgICAgICA8cD7Yrdin2YjZitin2Kog2YXYrtmE2YHYp9iqINio2YbYp9ihIDIwINmK2KfYsdiv2Kkgwrcg2KfZhNmF2YbYt9mC2Kkg2KfZhNi02LHZgtmK2Kk8L3A+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJ0cnVzdCI+CiAgICAgIDxzcGFuPvCfmpog2KfZhNiq2YjYtdmK2YQg2K7ZhNin2YQgMjQg2LPYp9i52Kk8L3NwYW4+CiAgICAgIDxzcGFuPvCfk4UgPGIgY2xhc3M9InBEYXlzIj4xMDwvYj4g2KPZitin2YUg2LnZhtiv2YM8L3NwYW4+CiAgICAgIDxzcGFuPvCfkrMg2KfZhNiv2YHYuSDZiNmC2Kog2KfZhNiq2LPZhNmK2YU8L3NwYW4+CiAgICA8L2Rpdj4KICA8L2Rpdj4KICA8ZGl2IGNsYXNzPSJzdHJpcGVzIj48L2Rpdj4KPC9oZWFkZXI+Cgo8bWFpbiBjbGFzcz0id3JhcCIgaWQ9ImFwcCI+CiAgPGRpdiBjbGFzcz0iaGVybyI+CiAgICA8aW1nIHNyYz0iaHR0cHM6Ly93YXNlZW1hbmFseXRpYy1kb3Rjb20uZ2l0aHViLmlvL3RoYXJhLWFscmltYWwvYXNzZXRzL2NvbnRhaW5lci5qcGciIGFsdD0i2K3Yp9mI2YrYqSDZhdiu2YTZgdin2Kog2KjZhtin2KEgMjAg2YrYp9ix2K/YqSDZhdmGINir2LHYp9ihINin2YTYsdmF2KfZhCIgd2lkdGg9IjkwMCIgaGVpZ2h0PSI3MDciPgogICAgPGRpdiBjbGFzcz0iYmFkZ2UiPjxiIGNsYXNzPSJwTmV3Ij40MDA8L2I+PHNtYWxsPtix2YrYp9mEIC8gPHNwYW4gY2xhc3M9InBEYXlzIj4xMDwvc3Bhbj4g2KPZitin2YU8L3NtYWxsPjwvZGl2PgogICAgPGRpdiBjbGFzcz0iaW5mbyI+CiAgICAgIDxoMj7Yrdin2YjZitipINmF2K7ZhNmB2KfYqiDYqNmG2KfYoTwvaDI+CiAgICAgIDxwPtiz2LnYqSAyMCDZitin2LHYr9ipIMK3INit2YjYp9mE2YogMTUg2YXCsyDCtyDYtNin2YXZhNipINin2YTYqtmI2LXZitmEINmI2KfZhNiz2K3YqC4g2KrYsdmF2YrZhSDYo9mIINio2YbYp9ih2J8g2YbZiNi12YTZh9inINmE2YXZiNmC2LnZgyDZiNmG2LPYrdio2YfYpyDZiNmC2Kog2YXYpyDYqtmF2KrZhNmKLjwvcD4KICAgIDwvZGl2PgogIDwvZGl2PgoKICA8c2VjdGlvbiBjbGFzcz0ic3RlcCIgaWQ9InMxIj4KICAgIDxoMj48c3BhbiBjbGFzcz0ibiI+MTwvc3Bhbj4g2YjYtCDYqtit2KrYp9is2J88L2gyPgogICAgPGRpdiBjbGFzcz0ic3ZjIiBpZD0ic3ZjIj48L2Rpdj4KICAgIDxkaXYgY2xhc3M9InF0eSI+CiAgICAgIDxsYWJlbCBmb3I9InF0eU91dCI+2LnYr9ivINin2YTYrdin2YjZitin2Ko8L2xhYmVsPgogICAgICA8ZGl2IGNsYXNzPSJzdGVwcGVyIj4KICAgICAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgaWQ9InFNaW51cyIgYXJpYS1sYWJlbD0i2KPZgtmEIj7iiJI8L2J1dHRvbj4KICAgICAgICA8b3V0cHV0IGlkPSJxdHlPdXQiPjE8L291dHB1dD4KICAgICAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgaWQ9InFQbHVzIiBhcmlhLWxhYmVsPSLYo9mD2KvYsSI+KzwvYnV0dG9uPgogICAgICA8L2Rpdj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmllbGQiIGlkPSJjb250RmllbGQiIHN0eWxlPSJkaXNwbGF5Om5vbmUiPgogICAgICA8bGFiZWwgZm9yPSJjb250YWluZXJObyI+2LHZgtmFINin2YTYrdin2YjZitipINin2YTZhNmKINi52YbYr9mDICjYp9iu2KrZitin2LHZiik8L2xhYmVsPgogICAgICA8aW5wdXQgaWQ9ImNvbnRhaW5lck5vIiBpbnB1dG1vZGU9Im51bWVyaWMiIG1heGxlbmd0aD0iMjAiIHBsYWNlaG9sZGVyPSLYp9mE2LHZgtmFINin2YTZhdmD2KrZiNioINi52YTZiSDYp9mE2K3Yp9mI2YrYqdiMINmF2KvZhCAwOSI+CiAgICA8L2Rpdj4KICAgIDxwIGNsYXNzPSJoaW50Ij7Zg9mEINmK2YjZhSDYpdi22KfZgdmKINio2LnYryDYp9mE2YrZiNmFIDxzcGFuIGNsYXNzPSJwRGF5cyI+MTA8L3NwYW4+OiA8YiBjbGFzcz0icEV4dHJhIj4zMDwvYj4g2LHZitin2YQuPC9wPgogIDwvc2VjdGlvbj4KCiAgPHNlY3Rpb24gY2xhc3M9InN0ZXAiIGlkPSJzMiI+CiAgICA8aDI+PHNwYW4gY2xhc3M9Im4iPjI8L3NwYW4+INmI2YrZhiDYp9mE2YXZiNmC2LnYnzwvaDI+CiAgICA8ZGl2IGNsYXNzPSJjaGlwcyIgaWQ9ImNpdGllcyI+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImRpc3RyaWN0Ij7Yp9mE2K3ZijwvbGFiZWw+CiAgICAgIDxpbnB1dCBpZD0iZGlzdHJpY3QiIG1heGxlbmd0aD0iNDAiIHBsYWNlaG9sZGVyPSLZhdir2YTYp9mLOiDYp9mE2LnYstmK2LLZitipIiBhdXRvY29tcGxldGU9ImFkZHJlc3MtbGV2ZWwzIj4KICAgIDwvZGl2PgogICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGNsYXNzPSJsb2NidG4iIGlkPSJsb2NhdGUiPvCfk40g2K3Yr9ivINmF2YjZgti52Yog2KfZhNit2KfZhNmKPC9idXR0b24+CiAgICA8cCBjbGFzcz0ib3IiPtij2Ygg2KfZhNi12YIg2LHYp9io2Lcg2KfZhNmF2YjZgti5INmF2YYg2K7Ysdin2KbYtyDZgtmI2YLZhDwvcD4KICAgIDxkaXYgY2xhc3M9ImZpZWxkIj4KICAgICAgPGlucHV0IGlkPSJtYXBMaW5rIiBtYXhsZW5ndGg9IjMwMCIgcGxhY2Vob2xkZXI9Imh0dHBzOi8vbWFwcy5hcHAuZ29vLmdsLy4uLiIgaW5wdXRtb2RlPSJ1cmwiPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+CiAgICAgIDxsYWJlbCBmb3I9ImFkZHJlc3MiPtmI2LXZgSDYp9mE2YXZg9in2YYgKNin2K7YqtmK2KfYsdmKKTwvbGFiZWw+CiAgICAgIDxpbnB1dCBpZD0iYWRkcmVzcyIgbWF4bGVuZ3RoPSIxNTAiIHBsYWNlaG9sZGVyPSLZhdir2YTYp9mLOiDYrNmG2Kgg2KfZhNmF2LPYrNiv2Iwg2YHZitmE2Kcg2KjZiNin2KjYqSDYs9mI2K/Yp9ihIj4KICAgIDwvZGl2PgogICAgPHAgY2xhc3M9Im5vdGUiPtmG2LHYrNmIINiq2KzZh9mK2LIg2YXZg9in2YYg2YXZhtin2LPYqCDZhNmE2K3Yp9mI2YrYqSDYqNi52YrYryDYudmGINmF2K/Yp9iu2YQg2KfZhNiz2YrYp9ix2KfYqi48L3A+CiAgPC9zZWN0aW9uPgoKICA8c2VjdGlvbiBjbGFzcz0ic3RlcCIgaWQ9InMzIj4KICAgIDxoMj48c3BhbiBjbGFzcz0ibiI+Mzwvc3Bhbj4g2YjYtCDZitmG2K3YtyDZgdmKINin2YTYrdin2YjZitip2J88L2gyPgogICAgPGRpdiBjbGFzcz0id2FzdGUiPgogICAgICA8ZGl2IGNsYXNzPSJ5ZXMiPjxoMz7inIUg2YXYs9mF2YjYrTog2KfZhNmG2YHYp9mK2KfYqiDYp9mE2KXZhti02KfYptmK2Kk8L2gzPgogICAgICAgIDx1bD48bGk+2K7Ysdiz2KfZhtipPC9saT48bGk+2LfZiNioPC9saT48bGk+2KjZhNin2Lc8L2xpPjxsaT7YsdmF2YQ8L2xpPjxsaT7Yrti02Kgg2KfZhNio2YbYp9ihPC9saT48L3VsPjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJubyI+PGgzPuKblCDZhdmF2YbZiNi5PC9oMz4KICAgICAgICA8dWw+PGxpPtin2YTZhtmB2KfZitin2Kog2KfZhNmF2YbYstmE2YrYqTwvbGk+PGxpPtin2YTZhdmI2KfYryDYp9mE2YPZitmF2YrYp9im2YrYqTwvbGk+PC91bD48L2Rpdj4KICAgIDwvZGl2PgogICAgPGxhYmVsIGNsYXNzPSJhZ3JlZSI+PGlucHV0IHR5cGU9ImNoZWNrYm94IiBpZD0iYWdyZWUiPiDYo9iq2LnZh9ivINil2YYg2KfZhNit2KfZiNmK2Kkg2YTZhNmG2YHYp9mK2KfYqiDYp9mE2KXZhti02KfYptmK2Kkg2YHZgti3PC9sYWJlbD4KICA8L3NlY3Rpb24+CgogIDxzZWN0aW9uIGNsYXNzPSJzdGVwIiBpZD0iczQiPgogICAgPGgyPjxzcGFuIGNsYXNzPSJuIj40PC9zcGFuPiDYqNmK2KfZhtin2KrZgyDZiNin2YTZhdmI2LnYrzwvaDI+CiAgICA8ZGl2IGNsYXNzPSJyb3cyIj4KICAgICAgPGRpdiBjbGFzcz0iZmllbGQiPjxsYWJlbCBmb3I9Im5hbWUiPtin2YTYp9iz2YU8L2xhYmVsPjxpbnB1dCBpZD0ibmFtZSIgbWF4bGVuZ3RoPSI2MCIgYXV0b2NvbXBsZXRlPSJuYW1lIj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iZmllbGQiPjxsYWJlbCBmb3I9InBob25lIj7Yp9mE2KzZiNin2YQ8L2xhYmVsPjxpbnB1dCBpZD0icGhvbmUiIGlucHV0bW9kZT0idGVsIiBtYXhsZW5ndGg9IjE0IiBwbGFjZWhvbGRlcj0iMDV4eHh4eHh4eCIgYXV0b2NvbXBsZXRlPSJ0ZWwiPjwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+CiAgICAgIDxsYWJlbD7Yo9mG2Ko8L2xhYmVsPgogICAgICA8ZGl2IGNsYXNzPSJjaGlwcyIgaWQ9ImN0eXBlIj48L2Rpdj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iZmllbGQiPgogICAgICA8bGFiZWwgZm9yPSJkYXRlIj7Yp9mE2YXZiNi52K8g2KfZhNmF2LfZhNmI2Kg8L2xhYmVsPgogICAgICA8aW5wdXQgdHlwZT0iZGF0ZSIgaWQ9ImRhdGUiPgogICAgICA8cCBjbGFzcz0iaGludCIgaWQ9ImRhdGVIaW50Ij7Yp9mE2KrZiNi12YrZhCDYrtmE2KfZhCAyNCDYs9in2LnYqSDZhdmGINiq2KPZg9mK2K8g2KfZhNi32YTYqC4g2KfZhNis2YXYudipINil2KzYp9iy2KkuPC9wPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJmaWVsZCI+CiAgICAgIDxsYWJlbD7Yt9ix2YrZgtipINin2YTYr9mB2LkgKNmI2YLYqiDYp9mE2KrYs9mE2YrZhSk8L2xhYmVsPgogICAgICA8ZGl2IGNsYXNzPSJjaGlwcyIgaWQ9InBheSI+PC9kaXY+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImZpZWxkIj4KICAgICAgPGxhYmVsIGZvcj0ibm90ZXMiPtmF2YTYp9it2LjYp9iqICjYp9iu2KrZitin2LHZiik8L2xhYmVsPgogICAgICA8dGV4dGFyZWEgaWQ9Im5vdGVzIiBtYXhsZW5ndGg9IjMwMCIgcGxhY2Vob2xkZXI9Itij2Yog2KrZgdin2LXZitmEINiq2LPYp9i52K/ZhtinIj48L3RleHRhcmVhPgogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0ic3VtIiBpZD0ic3VtIj48L2Rpdj4KICAgIDxwIGNsYXNzPSJlcnIiIGlkPSJlcnIiIHJvbGU9ImFsZXJ0Ij48L3A+CiAgPC9zZWN0aW9uPgoKICA8Zm9vdGVyPgogICAg2YTZhNin2LPYqtmB2LPYp9ixOiA8YSBocmVmPSJodHRwczovL3dhLm1lLzk2NjU3OTMyMjAzMCI+2YjYp9iq2LPYp9ioIDA1NzkzMjIwMzA8L2E+PGJyPgogICAg2KvYsdin2KEg2KfZhNix2YXYp9mEIMK3INmF2YYg2KfZhNiu2KjYsSDYpdmE2Ykg2KfZhNmC2LfZitmBCiAgPC9mb290ZXI+CjwvbWFpbj4KCjxkaXYgY2xhc3M9ImJhciIgaWQ9ImJhciI+CiAgPGRpdiBjbGFzcz0id3JhcCI+CiAgICA8ZGl2IGNsYXNzPSJ0Ij48c21hbGw+2KfZhNil2KzZhdin2YTZijwvc21hbGw+PGIgaWQ9ImJhclRvdGFsIj7igJQ8L2I+PC9kaXY+CiAgICA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgaWQ9InN1Ym1pdCI+2KPYsdiz2YQg2KfZhNi32YTYqDwvYnV0dG9uPgogIDwvZGl2Pgo8L2Rpdj4KCjxzZWN0aW9uIGNsYXNzPSJkb25lIHdyYXAiIGlkPSJkb25lIiBhcmlhLWxpdmU9InBvbGl0ZSI+CiAgPGRpdiBjbGFzcz0iaWMiPuKckzwvZGl2PgogIDxoMj7ZiNi12YTZhtinINi32YTYqNmDITwvaDI+CiAgPHAgY2xhc3M9ImlkIiBpZD0iZG9uZUlkIj48L3A+CiAgPHA+2KjZhtmB2KrYrSDZhNmDINmI2KfYqtiz2KfYqCDZiNin2YTYt9mE2Kgg2YXZg9iq2YjYqNiMINin2LbYuti3IDxiPtil2LHYs9in2YQ8L2I+INi52LTYp9mGINmG2KPZg9ivINmE2YMg2KfZhNmF2YjYudivLjwvcD4KICA8YSBjbGFzcz0id2EiIGlkPSJ3YUxpbmsiIGhyZWY9IiMiIHRhcmdldD0iX3RvcCIgcmVsPSJub29wZW5lciI+8J+SrCDYo9ix2LPZhCDYp9mE2LfZhNioINi52YTZiSDZiNin2KrYs9in2Kg8L2E+CiAgPHAgc3R5bGU9Im1hcmdpbi10b3A6MTZweCI+2YXYpyDYp9mG2YHYqtitINmI2KfYqtiz2KfYqNifINin2LbYuti3INin2YTYstixINmB2YjZgi48L3A+Cjwvc2VjdGlvbj4KCjxzY3JpcHQ+CihmdW5jdGlvbigpeyAidXNlIHN0cmljdCI7Ci8qID09PT09PT09PT09PSDYp9mE2KXYudiv2KfYr9in2KogPT09PT09PT09PT09ICovCmNvbnN0IENPTkZJRyA9IHsKICAvLyDYsdin2KjYtyBHb29nbGUgQXBwcyBTY3JpcHQg2KjYudivINin2YTZhti02LEuINil2LDYpyDYqtmP2LHZgyDZgdin2LbZitin2Ysg2YrZj9ix2LPZhCDYp9mE2LfZhNioINi52YTZiSDZiNin2KrYs9in2Kgg2YHZgti3LgogIEVORFBPSU5UOiAiaHR0cHM6Ly9zY3JpcHQuZ29vZ2xlLmNvbS9tYWNyb3Mvcy9BS2Z5Y2J6NUlzbUlSWHhPV3lOUm9JZmY4OUhaTWdBeFJmVUhlQzVzRlo2cENGckxzV0lYUElyQmNqSFl2NzBGTjBQaUNmVGJEdy9leGVjIiwKICBXSEFUU0FQUDogIjk2NjU3OTMyMjAzMCIsCiAgQ0lUSUVTOiBbItin2YTYrtio2LEiLCLYp9mE2LjZh9ix2KfZhiIsItin2YTYr9mF2KfZhSIsItiz2YrZh9in2KoiLCLYp9mE2YLYt9mK2YEiLCLZhdiv2YrZhtipINij2K7YsdmJIl0sCiAgQ1VTVE9NRVJfVFlQRVM6IFsi2YHYsdivIiwi2YXZgtin2YjZhCDYo9mIINi02LHZg9ipIl0sCiAgUEFZTUVOVFM6IFsi2LTYqNmD2KkiLCLYqtit2YjZitmEINio2YbZg9mKIl0KfTsKLy8g2KfZhNij2LPYudin2LEg2KfZhNin2YHYqtix2KfYttmK2KkgKNiq2KrYrdiv2Ksg2KrZhNmC2KfYptmK2KfZiyDZhdmGINmI2LHZgtipICLYp9mE2KPYs9i52KfYsSIg2YHZiiDYrNiv2YjZhCDZgtmI2YLZhCkKbGV0IFAgPSB7TkVXOjQwMCwgU1dBUDo0MDAsIEVYVFJBX0RBWTozMCwgREFZUzoxMH07CmNvbnN0IFNFUlZJQ0VTID0gWwogIHtpZDoiTkVXIiwgICAgbmFtZToi2K3Yp9mI2YrYqSDYrNiv2YrYr9ipIiwgICAgICBkZXNjOigpID0+IGDZhtmI2LXZhNmH2Kcg2YTZhdmI2YLYudmDINmI2KrYqNmC2Ykg2LnZhtiv2YMgJHtQLkRBWVN9INij2YrYp9mF2Iwg2LTYp9mF2YTYqSDYp9mE2LPYrdioYH0sCiAge2lkOiJTV0FQIiwgICBuYW1lOiLYp9iz2KrYqNiv2KfZhCDYrdin2YjZitipIiwgICAgZGVzYzooKSA9PiAi2YbYs9it2Kgg2KfZhNmF2YXYqtmE2KbYqSDZiNmG2K3YtyDZhNmDINmB2KfYttmK2KkifQpdOwoKY29uc3QgJCA9IGlkID0+IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlkKTsKY29uc3Qgc3RhdGUgPSB7c3ZjOiJORVciLCBxdHk6MSwgY2l0eTpudWxsLCBjdHlwZTpudWxsLCBwYXk6bnVsbCwgbG9jOm51bGwsIHNlbmRpbmc6ZmFsc2V9Owpjb25zdCBmbXQgPSBuID0+IE51bWJlci5pc0ludGVnZXIobikgPyBuLnRvTG9jYWxlU3RyaW5nKCJlbi1VUyIpIDogbi50b0xvY2FsZVN0cmluZygiZW4tVVMiLHttaW5pbXVtRnJhY3Rpb25EaWdpdHM6MixtYXhpbXVtRnJhY3Rpb25EaWdpdHM6Mn0pOwpjb25zdCByb3VuZDIgPSBuID0+IE1hdGgucm91bmQobioxMDApLzEwMDsKY29uc3QgZXNjID0gcyA9PiBTdHJpbmcocykucmVwbGFjZSgvWyY8PiInXS9nLCBjID0+ICh7IiYiOiImYW1wOyIsIjwiOiImbHQ7IiwiPiI6IiZndDsiLCciJzoiJnF1b3Q7IiwiJyI6IiYjMzk7In1bY10pKTsKY29uc3QgaW5HYXMgPSAoKSA9PiAhISh3aW5kb3cuZ29vZ2xlICYmIGdvb2dsZS5zY3JpcHQgJiYgZ29vZ2xlLnNjcmlwdC5ydW4pOwpjb25zdCBzdmNQcmljZSA9IGlkID0+IE51bWJlcihQW2lkXSkgfHwgMDsKY29uc3QgcHJpY2VMYWJlbCA9IGlkID0+IHN2Y1ByaWNlKGlkKSA/IGAke2ZtdChzdmNQcmljZShpZCkpfSDYsdmK2KfZhGAgOiAi2YXYtNmF2YjZhCI7CgpmdW5jdGlvbiByZW5kZXJQcmljZXMoKXsKICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCIucE5ldyIpLmZvckVhY2goZSA9PiBlLnRleHRDb250ZW50ID0gZm10KE51bWJlcihQLk5FVykpKTsKICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCIucERheXMiKS5mb3JFYWNoKGUgPT4gZS50ZXh0Q29udGVudCA9IFAuREFZUyk7CiAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgiLnBFeHRyYSIpLmZvckVhY2goZSA9PiBlLnRleHRDb250ZW50ID0gZm10KE51bWJlcihQLkVYVFJBX0RBWSkpKTsKfQpmdW5jdGlvbiByZW5kZXJTdmMoKXsKICAkKCJzdmMiKS5pbm5lckhUTUwgPSBTRVJWSUNFUy5tYXAocyA9PiBgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtcz0iJHtzLmlkfSIgYXJpYS1wcmVzc2VkPSIke3N0YXRlLnN2Yz09PXMuaWR9Ij4KICAgIDxkaXY+PGI+JHtlc2Mocy5uYW1lKX08L2I+PHNwYW4+JHtlc2Mocy5kZXNjKCkpfTwvc3Bhbj48L2Rpdj48ZW0+JHtwcmljZUxhYmVsKHMuaWQpfTwvZW0+PC9idXR0b24+YCkuam9pbigiIik7CiAgJCgiY29udEZpZWxkIikuc3R5bGUuZGlzcGxheSA9IHN0YXRlLnN2YyA9PT0gIk5FVyIgPyAibm9uZSIgOiAiYmxvY2siOwp9CmZ1bmN0aW9uIGNoaXBzKGVsLCBsaXN0LCBrZXkpewogICQoZWwpLmlubmVySFRNTCA9IGxpc3QubWFwKHYgPT4gYDxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLXY9IiR7ZXNjKHYpfSIgYXJpYS1wcmVzc2VkPSIke3N0YXRlW2tleV09PT12fSI+JHtlc2Modil9PC9idXR0b24+YCkuam9pbigiIik7Cn0KZnVuY3Rpb24gdG90YWxzKCl7CiAgY29uc3Qgc3ViID0gcm91bmQyKHN2Y1ByaWNlKHN0YXRlLnN2YykgKiBzdGF0ZS5xdHkpOwogIHJldHVybiB7c3ViLCB0b3RhbDogc3VifTsKfQpmdW5jdGlvbiByZW5kZXJTdW0oKXsKICBjb25zdCBzID0gU0VSVklDRVMuZmluZCh4ID0+IHguaWQgPT09IHN0YXRlLnN2YyksIHQgPSB0b3RhbHMoKTsKICAkKCJzdW0iKS5pbm5lckhUTUwgPSBgCiAgICA8ZGl2PjxzcGFuPiR7ZXNjKHMubmFtZSl9IMOXICR7c3RhdGUucXR5fTwvc3Bhbj48c3Bhbj4ke2ZtdCh0LnN1Yil9INix2YrYp9mEPC9zcGFuPjwvZGl2PgogICAgPGRpdiBjbGFzcz0idG90Ij48c3Bhbj7Yp9mE2KXYrNmF2KfZhNmKPC9zcGFuPjxzcGFuPiR7Zm10KHQudG90YWwpfSDYsdmK2KfZhDwvc3Bhbj48L2Rpdj5gOwogICQoImJhclRvdGFsIikudGV4dENvbnRlbnQgPSBgJHtmbXQodC50b3RhbCl9INix2YrYp9mEYDsKICAkKCJxdHlPdXQiKS50ZXh0Q29udGVudCA9IHN0YXRlLnF0eTsKfQpmdW5jdGlvbiByZW5kZXJBbGwoKXsKICByZW5kZXJQcmljZXMoKTsgcmVuZGVyU3ZjKCk7IHJlbmRlclN1bSgpOwogIGNoaXBzKCJjaXRpZXMiLCBDT05GSUcuQ0lUSUVTLCAiY2l0eSIpOyBjaGlwcygiY3R5cGUiLCBDT05GSUcuQ1VTVE9NRVJfVFlQRVMsICJjdHlwZSIpOyBjaGlwcygicGF5IiwgQ09ORklHLlBBWU1FTlRTLCAicGF5Iik7Cn0KCi8qID09PT09PT09PT09PSDYp9mE2KrZgdin2LnZhCA9PT09PT09PT09PT0gKi8KJCgic3ZjIikuYWRkRXZlbnRMaXN0ZW5lcigiY2xpY2siLCBlID0+IHsgY29uc3QgYiA9IGUudGFyZ2V0LmNsb3Nlc3QoImJ1dHRvbiIpOyBpZiAoIWIpIHJldHVybjsgc3RhdGUuc3ZjID0gYi5kYXRhc2V0LnM7IHJlbmRlclN2YygpOyByZW5kZXJTdW0oKTsgfSk7CiQoInFNaW51cyIpLmFkZEV2ZW50TGlzdGVuZXIoImNsaWNrIiwgKCkgPT4geyBzdGF0ZS5xdHkgPSBNYXRoLm1heCgxLCBzdGF0ZS5xdHkgLSAxKTsgcmVuZGVyU3VtKCk7IH0pOwokKCJxUGx1cyIpLmFkZEV2ZW50TGlzdGVuZXIoImNsaWNrIiwgKCkgPT4geyBzdGF0ZS5xdHkgPSBNYXRoLm1pbig1LCBzdGF0ZS5xdHkgKyAxKTsgcmVuZGVyU3VtKCk7IH0pOwpbWyJjaXRpZXMiLCJjaXR5Il0sWyJjdHlwZSIsImN0eXBlIl0sWyJwYXkiLCJwYXkiXV0uZm9yRWFjaCgoW2VsLGtleV0pID0+IHsKICAkKGVsKS5hZGRFdmVudExpc3RlbmVyKCJjbGljayIsIGUgPT4geyBjb25zdCBiID0gZS50YXJnZXQuY2xvc2VzdCgiYnV0dG9uIik7IGlmICghYikgcmV0dXJuOwogICAgc3RhdGVba2V5XSA9IGIuZGF0YXNldC52OyAkKGVsKS5xdWVyeVNlbGVjdG9yQWxsKCJidXR0b24iKS5mb3JFYWNoKHggPT4geC5zZXRBdHRyaWJ1dGUoImFyaWEtcHJlc3NlZCIsIHggPT09IGIpKTsgfSk7Cn0pOwokKCJsb2NhdGUiKS5hZGRFdmVudExpc3RlbmVyKCJjbGljayIsICgpID0+IHsKICBjb25zdCBidG4gPSAkKCJsb2NhdGUiKTsKICBpZiAoIW5hdmlnYXRvci5nZW9sb2NhdGlvbil7IGJ0bi50ZXh0Q29udGVudCA9ICLYrNmH2KfYstmDINmF2Kcg2YrYr9i52YUg2KrYrdiv2YrYryDYp9mE2YXZiNmC2LnYjCDYp9mE2LXZgiDYp9mE2LHYp9io2Lcg2KrYrdiqIjsgcmV0dXJuOyB9CiAgYnRuLnRleHRDb250ZW50ID0gItis2KfYsdmKINiq2K3Yr9mK2K8g2YXZiNmC2LnZg+KApiI7CiAgbmF2aWdhdG9yLmdlb2xvY2F0aW9uLmdldEN1cnJlbnRQb3NpdGlvbihwID0+IHsKICAgIHN0YXRlLmxvYyA9IHtsYXQ6IHAuY29vcmRzLmxhdGl0dWRlLnRvRml4ZWQoNiksIGxuZzogcC5jb29yZHMubG9uZ2l0dWRlLnRvRml4ZWQoNil9OwogICAgYnRuLnRleHRDb250ZW50ID0gIuKckyDYqtmFINiq2K3Yr9mK2K8g2YXZiNmC2LnZgyI7IGJ0bi5jbGFzc0xpc3QuYWRkKCJvayIpOwogIH0sICgpID0+IHsgYnRuLnRleHRDb250ZW50ID0gItmF2Kcg2YLYr9ix2YbYpyDZhtit2K/YryDZhdmI2YLYudmD2Iwg2KfZhNi12YIg2LHYp9io2Lcg2KfZhNmF2YjZgti5INiq2K3YqiI7IH0sIHtlbmFibGVIaWdoQWNjdXJhY3k6dHJ1ZSwgdGltZW91dDoxNTAwMH0pOwp9KTsKCmNvbnN0IHRvZGF5ID0gbmV3IERhdGUoKTsgdG9kYXkuc2V0TWludXRlcyh0b2RheS5nZXRNaW51dGVzKCkgLSB0b2RheS5nZXRUaW1lem9uZU9mZnNldCgpKTsKY29uc3QgaXNvID0gdG9kYXkudG9JU09TdHJpbmcoKS5zbGljZSgwLDEwKTsKJCgiZGF0ZSIpLm1pbiA9IGlzbzsgJCgiZGF0ZSIpLnZhbHVlID0gaXNvOwokKCJkYXRlIikuYWRkRXZlbnRMaXN0ZW5lcigiY2hhbmdlIiwgKCkgPT4gewogIGNvbnN0IGQgPSBuZXcgRGF0ZSgkKCJkYXRlIikudmFsdWUgKyAiVDEyOjAwOjAwIik7CiAgJCgiZGF0ZUhpbnQiKS50ZXh0Q29udGVudCA9IGQuZ2V0RGF5KCkgPT09IDUKICAgID8gIuKaoO+4jyDYp9mE2KzZhdi52Kkg2KXYrNin2LLYqdiMINio2YbYo9mD2K8g2YTZgyDYo9mC2LHYqCDZhdmI2LnYryDYqNi52K/Zh9inLiIKICAgIDogItin2YTYqtmI2LXZitmEINiu2YTYp9mEIDI0INiz2KfYudipINmF2YYg2KrYo9mD2YrYryDYp9mE2LfZhNioLiDYp9mE2KzZhdi52Kkg2KXYrNin2LLYqS4iOwp9KTsKCi8qID09PT09PT09PT09PSDYp9mE2KXYsdiz2KfZhCA9PT09PT09PT09PT0gKi8KZnVuY3Rpb24gbm9ybVBob25lKHYpewogIHYgPSB2LnJlcGxhY2UoL1vZoC3ZqV0vZywgZCA9PiAi2aDZodmi2aPZpNml2abZp9mo2akiLmluZGV4T2YoZCkpLnJlcGxhY2UoL1teXGRdL2csIiIpOwogIGlmICh2LnN0YXJ0c1dpdGgoIjk2NiIpKSB2ID0gIjAiICsgdi5zbGljZSgzKTsKICBpZiAodi5zdGFydHNXaXRoKCI1IikgJiYgdi5sZW5ndGggPT09IDkpIHYgPSAiMCIgKyB2OwogIHJldHVybiB2Owp9CmZ1bmN0aW9uIGNvbGxlY3QoKXsKICBjb25zdCBtYXBMaW5rID0gc3RhdGUubG9jID8gYGh0dHBzOi8vbWFwcy5nb29nbGUuY29tLz9xPSR7c3RhdGUubG9jLmxhdH0sJHtzdGF0ZS5sb2MubG5nfWAgOiAkKCJtYXBMaW5rIikudmFsdWUudHJpbSgpOwogIHJldHVybiB7CiAgICBzZXJ2aWNlOiBzdGF0ZS5zdmMsIHF0eTogc3RhdGUucXR5LCBjb250YWluZXJObzogJCgiY29udGFpbmVyTm8iKS52YWx1ZS50cmltKCksCiAgICBjaXR5OiBzdGF0ZS5jaXR5LCBkaXN0cmljdDogJCgiZGlzdHJpY3QiKS52YWx1ZS50cmltKCksIG1hcExpbmssIGFkZHJlc3M6ICQoImFkZHJlc3MiKS52YWx1ZS50cmltKCksCiAgICBuYW1lOiAkKCJuYW1lIikudmFsdWUudHJpbSgpLCBwaG9uZTogbm9ybVBob25lKCQoInBob25lIikudmFsdWUpLCBjdXN0b21lclR5cGU6IHN0YXRlLmN0eXBlLAogICAgZGF0ZTogJCgiZGF0ZSIpLnZhbHVlLCBwYXltZW50OiBzdGF0ZS5wYXksIG5vdGVzOiAkKCJub3RlcyIpLnZhbHVlLnRyaW0oKSwgYWdyZWU6ICQoImFncmVlIikuY2hlY2tlZAogIH07Cn0KZnVuY3Rpb24gdmFsaWRhdGUobyl7CiAgaWYgKCFvLmNpdHkpIHJldHVybiBbItin2K7YqtixINin2YTZhdiv2YrZhtipLiIsICJzMiJdOwogIGlmIChvLmRpc3RyaWN0Lmxlbmd0aCA8IDIpIHJldHVybiBbItin2YPYqtioINin2YTYrdmKLiIsICJzMiJdOwogIGlmICghby5tYXBMaW5rICYmIG8uYWRkcmVzcy5sZW5ndGggPCA1KSByZXR1cm4gWyLYrdiv2K8g2YXZiNmC2LnZgyDYo9mIINin2YTYtdmCINix2KfYqNi3INin2YTZhdmI2YLYuS4iLCAiczIiXTsKICBpZiAoIW8uYWdyZWUpIHJldHVybiBbItij2YPYryDYpdmGINin2YTYrdin2YjZitipINmE2YTZhtmB2KfZitin2Kog2KfZhNil2YbYtNin2KbZitipINmB2YLYty4iLCAiczMiXTsKICBpZiAoby5uYW1lLmxlbmd0aCA8IDIpIHJldHVybiBbItin2YPYqtioINin2LPZhdmDLiIsICJzNCJdOwogIGlmICghL14wNVxkezh9JC8udGVzdChvLnBob25lKSkgcmV0dXJuIFsi2KfZg9iq2Kgg2LHZgtmFINis2YjYp9mEINi12K3ZititINmK2KjYr9ijINio2YAgMDUuIiwgInM0Il07CiAgaWYgKCFvLmN1c3RvbWVyVHlwZSkgcmV0dXJuIFsi2KfYrtiq2LE6INmB2LHYryDYo9mIINmF2YLYp9mI2YQv2LTYsdmD2KkuIiwgInM0Il07CiAgaWYgKCFvLmRhdGUpIHJldHVybiBbItin2K7YqtixINin2YTZhdmI2LnYry4iLCAiczQiXTsKICBpZiAoIW8ucGF5bWVudCkgcmV0dXJuIFsi2KfYrtiq2LEg2LfYsdmK2YLYqSDYp9mE2K/Zgdi5LiIsICJzNCJdOwogIHJldHVybiBudWxsOwp9CmZ1bmN0aW9uIHdhVGV4dChvLCBpZCl7CiAgY29uc3QgcyA9IFNFUlZJQ0VTLmZpbmQoeCA9PiB4LmlkID09PSBvLnNlcnZpY2UpLCB0ID0gdG90YWxzKCk7CiAgcmV0dXJuIFsKICAgIGDYt9mE2Kgg2KzYr9mK2K8g2YXZhiDYtdmB2K3YqSDYq9ix2KfYoSDYp9mE2LHZhdin2YQke2lkID8gIiDYsdmC2YUgIiArIGlkIDogIiJ9YCwKICAgIGDYp9mE2K7Yr9mF2Kk6ICR7cy5uYW1lfSAoMjAg2YrYp9ix2K/YqSkgw5cgJHtvLnF0eX1gLAogICAgby5jb250YWluZXJObyA/IGDYsdmC2YUg2KfZhNit2KfZiNmK2Kk6ICR7by5jb250YWluZXJOb31gIDogIiIsCiAgICBg2KfZhNmF2KjZhNi6OiAke2ZtdCh0LnRvdGFsKX0g2LHZitin2YRgLAogICAgYNin2YTYp9iz2YU6ICR7by5uYW1lfWAsIGDYp9mE2KzZiNin2YQ6ICR7by5waG9uZX1gLCBg2KfZhNi52YXZitmEOiAke28uY3VzdG9tZXJUeXBlfWAsCiAgICBg2KfZhNmF2K/ZitmG2Kkg2YjYp9mE2K3ZijogJHtvLmNpdHl9IC0gJHtvLmRpc3RyaWN0fWAsCiAgICBvLm1hcExpbmsgPyBg2KfZhNmF2YjZgti5OiAke28ubWFwTGlua31gIDogIiIsIG8uYWRkcmVzcyA/IGDYp9mE2YjYtdmBOiAke28uYWRkcmVzc31gIDogIiIsCiAgICBg2KfZhNmF2YjYudivOiAke28uZGF0ZX1gLCBg2KfZhNiv2YHYuTogJHtvLnBheW1lbnR9YCwgby5ub3RlcyA/IGDZhdmE2KfYrdi42KfYqjogJHtvLm5vdGVzfWAgOiAiIgogIF0uZmlsdGVyKEJvb2xlYW4pLmpvaW4oIlxuIik7Cn0KZnVuY3Rpb24gc2hvd0RvbmUobywgaWQpewogICQoImFwcCIpLnN0eWxlLmRpc3BsYXkgPSAibm9uZSI7ICQoImJhciIpLnN0eWxlLmRpc3BsYXkgPSAibm9uZSI7CiAgJCgiZG9uZSIpLnN0eWxlLmRpc3BsYXkgPSAiYmxvY2siOwogICQoImRvbmVJZCIpLnRleHRDb250ZW50ID0gaWQgPyBg2LHZgtmFINi32YTYqNmDOiAke2lkfWAgOiAiIjsKICBjb25zdCB3YSA9IGBodHRwczovL3dhLm1lLyR7Q09ORklHLldIQVRTQVBQfT90ZXh0PSR7ZW5jb2RlVVJJQ29tcG9uZW50KHdhVGV4dChvLCBpZCkpfWA7CiAgJCgid2FMaW5rIikuaHJlZiA9IHdhOwogIHdpbmRvdy5zY3JvbGxUbygwLDApOwogIC8vINiq2K3ZiNmK2YQg2KfZhNi52YXZitmEINiq2YTZgtin2KbZitin2Ysg2YTZhNmI2KfYqtiz2KfYqCDZiNin2YTYsdiz2KfZhNipINis2KfZh9iy2KnYjCDZiNmF2Kcg2LnZhNmK2Ycg2KXZhNinINmK2LbYuti3INil2LHYs9in2YQKICBzZXRUaW1lb3V0KCgpID0+IHsKICAgIGlmIChpbkdhcygpKSB7IHRyeSB7IHdpbmRvdy5vcGVuKHdhLCAiX3RvcCIpOyB9IGNhdGNoKF8pIHt9IH0KICAgIGVsc2Ugd2luZG93LmxvY2F0aW9uLmhyZWYgPSB3YTsKICB9LCAxNTAwKTsKfQpmdW5jdGlvbiBmYWlsKCl7CiAgJCgiZXJyIikudGV4dENvbnRlbnQgPSAi2LXYp9ixINiu2LfYoyDZgdmKINin2YTYpdix2LPYp9mELiDYp9i22LrYtyDZhdix2Kkg2KvYp9mG2YrYqdiMINij2Ygg2LHYp9iz2YTZhtinINi52YTZiSDZiNin2KrYs9in2KggMDU3OTMyMjAzMC4iOwogICQoInN1Ym1pdCIpLmRpc2FibGVkID0gZmFsc2U7ICQoInN1Ym1pdCIpLnRleHRDb250ZW50ID0gItij2LHYs9mEINin2YTYt9mE2KgiOyBzdGF0ZS5zZW5kaW5nID0gZmFsc2U7Cn0KJCgic3VibWl0IikuYWRkRXZlbnRMaXN0ZW5lcigiY2xpY2siLCBhc3luYyAoKSA9PiB7CiAgaWYgKHN0YXRlLnNlbmRpbmcpIHJldHVybjsKICBjb25zdCBvID0gY29sbGVjdCgpLCB2ID0gdmFsaWRhdGUobyk7CiAgJCgiZXJyIikudGV4dENvbnRlbnQgPSB2ID8gdlswXSA6ICIiOwogIGlmICh2KXsgJCh2WzFdKS5zY3JvbGxJbnRvVmlldyh7YmVoYXZpb3I6InNtb290aCIsIGJsb2NrOiJzdGFydCJ9KTsgcmV0dXJuOyB9CiAgaWYgKCFpbkdhcygpICYmICFDT05GSUcuRU5EUE9JTlQpeyBzaG93RG9uZShvLCAiIik7IHJldHVybjsgfSAvLyDZiNi22Lkg2YjYp9iq2LPYp9ioINmB2YLYtwogIHN0YXRlLnNlbmRpbmcgPSB0cnVlOyAkKCJzdWJtaXQiKS5kaXNhYmxlZCA9IHRydWU7ICQoInN1Ym1pdCIpLnRleHRDb250ZW50ID0gItis2KfYsdmKINin2YTYpdix2LPYp9mE4oCmIjsKICBpZiAoaW5HYXMoKSl7CiAgICBnb29nbGUuc2NyaXB0LnJ1bi53aXRoU3VjY2Vzc0hhbmRsZXIoZCA9PiBkICYmIGQub2sgPyBzaG93RG9uZShvLCBkLmlkKSA6IGZhaWwoKSkud2l0aEZhaWx1cmVIYW5kbGVyKGZhaWwpLnN1Ym1pdE9yZGVyKG8pOwogICAgcmV0dXJuOwogIH0KICB0cnkgewogICAgY29uc3QgcmVzID0gYXdhaXQgZmV0Y2goQ09ORklHLkVORFBPSU5ULCB7bWV0aG9kOiJQT1NUIiwgYm9keTpKU09OLnN0cmluZ2lmeShvKX0pOwogICAgY29uc3QgZGF0YSA9IGF3YWl0IHJlcy5qc29uKCk7CiAgICBpZiAoIWRhdGEub2spIHRocm93IG5ldyBFcnJvcihkYXRhLmVycm9yIHx8ICJmYWlsZWQiKTsKICAgIHNob3dEb25lKG8sIGRhdGEuaWQpOwogIH0gY2F0Y2ggKF8pIHsgZmFpbCgpOyB9Cn0pOwoKLyogPT09PT09PT09PT09INiq2K3ZhdmK2YQg2KfZhNij2LPYudin2LEg2YXZhiDYrNiv2YjZhCDZgtmI2YLZhCA9PT09PT09PT09PT0gKi8KZnVuY3Rpb24gYXBwbHlQcmljZXMocCl7IGlmIChwICYmIHR5cGVvZiBwID09PSAib2JqZWN0IiAmJiBwLk5FVyAhPSBudWxsKXsgUCA9IE9iamVjdC5hc3NpZ24oe30sIFAsIHApOyByZW5kZXJBbGwoKTsgfSB9CmFzeW5jIGZ1bmN0aW9uIGxvYWRQcmljZXMoKXsKICBpZiAoaW5HYXMoKSl7IGdvb2dsZS5zY3JpcHQucnVuLndpdGhTdWNjZXNzSGFuZGxlcihhcHBseVByaWNlcykuZ2V0UHJpY2VzKCk7IHJldHVybjsgfQogIGlmICghQ09ORklHLkVORFBPSU5UKSByZXR1cm47CiAgdHJ5IHsKICAgIGNvbnN0IHJlcyA9IGF3YWl0IGZldGNoKENPTkZJRy5FTkRQT0lOVCArICI/YWN0aW9uPXByaWNlcyIpOwogICAgY29uc3QgZGF0YSA9IGF3YWl0IHJlcy5qc29uKCk7CiAgICBpZiAoZGF0YS5vaykgYXBwbHlQcmljZXMoZGF0YS5wcmljZXMpOwogIH0gY2F0Y2goXykgeyAvKiDZhtiz2KrYrtiv2YUg2KfZhNij2LPYudin2LEg2KfZhNmF2LbZhdmG2KkgKi8gfQp9CnJlbmRlckFsbCgpOyBsb2FkUHJpY2VzKCk7Cn0pKCk7Cjwvc2NyaXB0Pgo8L2JvZHk+CjwvaHRtbD4K';
