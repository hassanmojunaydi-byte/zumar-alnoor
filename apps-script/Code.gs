/**
 * منصة حلقات زمر النور — الخادم (Google Apps Script)
 * قاعدة البيانات: ملف Google Sheets «قاعدة بيانات منصة زمر النور».
 * الحضور يُقرأ مباشرة من ملفات التحضير الاثني عشر، والتسميع يُكتب فيها.
 */
var CFG = {
  MASTER_ID: '1vF_rbFEAbQCN-oTgUrOWjScOkAkTKAXJFD9hxS4j_5E',
  OWNERS: ['hassan.mo.junaydi@gmail.com'],
  APP: 'حلقات زمر النور',
  GRID_TTL: 90,          // ثوانٍ: مدة الاحتفاظ بقراءة ملف التحضير
  SESSION_DAYS: 60,
  TEACHER_FILES_FOLDER: '1aXXMQs45LlfPPXtDO6gtSYkQorhPXNsw' // «صور وملفات المعلمين»
};
var ROLE_AR = {system:'النظام', admin:'الإدارة', exams:'لجنة الاختبارات', teacher:'معلم'};
var ROLE_EN = {'النظام':'system','الإدارة':'admin','لجنة الاختبارات':'exams','معلم':'teacher'};
var ACC_HEAD = ['البريد','الاسم','الدور','اسم المعلم في الكشوف','الحالة','تاريخ الطلب','آخر دخول','ملاحظات'];

/* ================= HTTP ================= */
function doGet(e) {
  var out = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CFG.APP)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  var icon = PropertiesService.getScriptProperties().getProperty('ICON_URL');
  if (icon) { try { out.setFaviconUrl(icon); } catch (err) {} }
  return out;
}
function doPost(e) {
  var req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) {}
  return ContentService.createTextOutput(JSON.stringify(handle_(req))).setMimeType(ContentService.MimeType.JSON);
}
/** تُستدعى من الصفحة عبر google.script.run */
function api(req) { return JSON.stringify(handle_(req || {})); }

function handle_(req) {
  try {
    var a = req.action, p = req.payload || {};
    if (a === 'public') return ok_({app: CFG.APP, clientId: prop_('GOOGLE_CLIENT_ID') || ''});
    if (a === 'code') return sendCode_(p.email);
    if (a === 'verify') return verifyCode_(p.email, p.code);
    if (a === 'google') return googleLogin_(p.credential);
    var me = auth_(req.token);
    if (!me) return {ok: false, error: 'auth'};
    if (a === 'logout') { PropertiesService.getScriptProperties().deleteProperty('s_' + req.token); return ok_({}); }
    if (a === 'boot') return boot_(me, p);
    if (a === 'claim') return claim_(me, p);
    if (!me.active) return {ok: false, error: 'pending'};
    if (a === 'write') return write_(me, p);
    if (a === 'tasmee') return tasmee_(me, p);
    if (a === 'accounts') return accounts_(me, p);
    if (a === 'notify') return notifyCommittee_(me, p.date);
    if (a === 'photos') return photos_(me);
    if (a === 'photo') return setPhoto_(me, p);
    if (a === 'importPhotos') return importPhotos_(me);
    return {ok: false, error: 'unknown_action'};
  } catch (err) {
    return {ok: false, error: String(err && err.message || err)};
  }
}
function ok_(o) { o.ok = true; return o; }
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function norm_(s) { return String(s || '').replace(/^أ\.\s*/, '').replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/\s+/g, '').toLowerCase(); }
function tokens_(s) { return String(s || '').replace(/^أ\.\s*/, '').replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').trim().split(/\s+/); }
var TALIAS = {'حسن زكريا': 'حسن جنيدي', 'حسن زكريا جنيدي': 'حسن جنيدي'};
function canon_(n) { n = String(n || '').replace(/^أ\.\s*/, '').trim(); return TALIAS[n] || n; }
function today_() { return Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd'); }
function addDays_(d, n) { var x = new Date(d + 'T12:00:00+03:00'); x.setDate(x.getDate() + n); return Utilities.formatDate(x, 'Asia/Riyadh', 'yyyy-MM-dd'); }
function dow_(d) { return new Date(d + 'T12:00:00+03:00').getDay(); }
function appUrl_() { return prop_('APP_URL') || ScriptApp.getService().getUrl(); }
function master_() { return SpreadsheetApp.openById(CFG.MASTER_ID); }

/* ================= ACCOUNTS & AUTH ================= */
function accSheet_() {
  var ss = master_(), sh = ss.getSheetByName('الحسابات');
  if (!sh) {
    sh = ss.insertSheet('الحسابات'); sh.setRightToLeft(true);
    sh.getRange(1, 1, 1, ACC_HEAD.length).setValues([ACC_HEAD]).setFontWeight('bold'); sh.setFrozenRows(1);
  }
  return sh;
}
function readAccounts_() {
  var v = accSheet_().getDataRange().getValues(), out = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i]; if (!r[0] && !r[1]) continue;
    out.push({row: i + 1, email: String(r[0] || '').trim().toLowerCase(), name: String(r[1] || ''), role: ROLE_EN[String(r[2] || '').trim()] || '',
      tname: canon_(r[3]), status: String(r[4] || ''), requested: r[5] ? String(r[5]) : '', last: r[6] ? String(r[6]) : '', notes: String(r[7] || '')});
  }
  return out;
}
function findAccount_(email) { email = String(email || '').toLowerCase(); var all = readAccounts_(); for (var i = 0; i < all.length; i++) if (all[i].email === email) return all[i]; return null; }
function auth_(token) {
  if (!token) return null;
  var raw = prop_('s_' + token); if (!raw) return null;
  var s = JSON.parse(raw); if (s.exp < Date.now()) { PropertiesService.getScriptProperties().deleteProperty('s_' + token); return null; }
  var email = s.email, isOwner = CFG.OWNERS.indexOf(email) >= 0, acc = findAccount_(email);
  var role = isOwner ? 'system' : (acc && acc.status === 'مفعّل' ? acc.role : '');
  return {email: email, owner: isOwner, role: role, active: !!role, tname: acc ? acc.tname : '', name: acc ? acc.name : '', status: acc ? acc.status : (isOwner ? 'مفعّل' : ''), account: acc};
}
/* تنظيف الجلسات: حذف المنتهية، والإبقاء على آخر 5 أجهزة لكل بريد */
function purgeSessions_() {
  var ps = PropertiesService.getScriptProperties(), all = ps.getProperties(), now = Date.now(), byEmail = {};
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('s_') !== 0) return;
    var s; try { s = JSON.parse(all[k]); } catch (e) { ps.deleteProperty(k); return; }
    if (!s || s.exp < now) { ps.deleteProperty(k); return; }
    (byEmail[s.email] = byEmail[s.email] || []).push({k: k, exp: s.exp});
  });
  Object.keys(byEmail).forEach(function (e) { byEmail[e].sort(function (a, b) { return b.exp - a.exp; }).slice(5).forEach(function (x) { ps.deleteProperty(x.k); }); });
}
function newSession_(email) {
  try { purgeSessions_(); } catch (e) {}
  var token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  PropertiesService.getScriptProperties().setProperty('s_' + token, JSON.stringify({email: email, exp: Date.now() + CFG.SESSION_DAYS * 864e5}));
  var acc = findAccount_(email);
  if (acc) accSheet_().getRange(acc.row, 7).setValue(Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'));
  return token;
}
function sendCode_(email) {
  email = String(email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return {ok: false, error: 'اكتب بريداً صحيحاً'};
  var c = CacheService.getScriptCache(), k = 'otp_' + email, prev = c.get(k);
  if (prev && JSON.parse(prev).at > Date.now() - 45000) return {ok: false, error: 'انتظر دقيقة ثم اطلب رمزاً جديداً'};
  var code = String(Math.floor(100000 + Math.random() * 900000));
  c.put(k, JSON.stringify({code: code, tries: 0, at: Date.now()}), 600);
  MailApp.sendEmail({to: email, subject: 'رمز الدخول إلى ' + CFG.APP + ': ' + code, name: CFG.APP,
    htmlBody: mail_('رمز الدخول', '<p style="margin:0 0 12px">استخدم هذا الرمز للدخول إلى منصة ' + CFG.APP + '. صالح لمدة 10 دقائق.</p>' +
      '<div style="font-size:34px;font-weight:800;letter-spacing:8px;background:#f6eedb;color:#6a1820;border-radius:12px;padding:14px;text-align:center;direction:ltr">' + code + '</div>' +
      '<p style="margin:14px 0 0;color:#74645d;font-size:13px">إذا لم تطلب الرمز فتجاهل هذه الرسالة.</p>')});
  return ok_({});
}
function verifyCode_(email, code) {
  email = String(email || '').trim().toLowerCase();
  var c = CacheService.getScriptCache(), k = 'otp_' + email, raw = c.get(k);
  if (!raw) return {ok: false, error: 'انتهت صلاحية الرمز، اطلب رمزاً جديداً'};
  var o = JSON.parse(raw);
  if (o.tries >= 5) { c.remove(k); return {ok: false, error: 'محاولات كثيرة، اطلب رمزاً جديداً'}; }
  if (String(code).trim() !== o.code) { o.tries++; c.put(k, JSON.stringify(o), 600); return {ok: false, error: 'الرمز غير صحيح'}; }
  c.remove(k);
  return ok_({token: newSession_(email)});
}
function googleLogin_(cred) {
  var cid = prop_('GOOGLE_CLIENT_ID'); if (!cid || !cred) return {ok: false, error: 'الدخول بزر Google غير مفعّل'};
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(cred), {muteHttpExceptions: true});
  if (r.getResponseCode() !== 200) return {ok: false, error: 'تعذّر التحقق من حساب Google'};
  var t = JSON.parse(r.getContentText());
  if (t.aud !== cid || String(t.email_verified) !== 'true') return {ok: false, error: 'تعذّر التحقق من حساب Google'};
  return ok_({token: newSession_(String(t.email).toLowerCase())});
}
function claim_(me, p) {
  if (me.active) return ok_({});
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = accSheet_(), acc = findAccount_(me.email);
    var role = p.want === 'admin' ? 'الإدارة' : p.want === 'exams' ? 'لجنة الاختبارات' : 'معلم';
    var row = [me.email, p.name || canon_(p.tname) || '', role, canon_(p.tname) || '', 'بانتظار الاعتماد', Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'), '', ''];
    if (acc) sh.getRange(acc.row, 1, 1, 6).setValues([row.slice(0, 6)]); else sh.appendRow(row);
  } finally { lock.releaseLock(); }
  var link = appUrl_();
  CFG.OWNERS.forEach(function (o) {
    MailApp.sendEmail({to: o, subject: 'طلب انضمام جديد إلى ' + CFG.APP, name: CFG.APP,
      htmlBody: mail_('طلب انضمام جديد', '<p><b>' + esc_(p.name || canon_(p.tname) || me.email) + '</b> (' + esc_(me.email) + ') يطلب الدخول بصفة <b>' + esc_(p.want === 'admin' ? 'الإدارة' : p.want === 'exams' ? 'لجنة الاختبارات' : 'معلم') + '</b>.</p><p>اعتمده من: النظام ← بوابة الحسابات.</p>' + btn_(link, 'فتح المنصة'))});
  });
  return ok_({});
}
function accounts_(me, p) {
  if (me.role !== 'system') return {ok: false, error: 'forbidden'};
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  var mailTo = null, msg = '';
  try {
    var sh = accSheet_(), acc = p.email ? findAccount_(p.email) : null;
    if (p.op === 'invite' || p.op === 'save') {
      var email = String(p.email || '').trim().toLowerCase();
      var row = [email, p.name || '', ROLE_AR[p.role] || 'معلم', canon_(p.tname) || '', p.status || 'مفعّل', acc ? acc.requested : Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'), acc ? acc.last : '', p.notes || (acc ? acc.notes : '')];
      var target = acc || (p.row ? {row: +p.row} : null);
      if (target) sh.getRange(target.row, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
      if (email && (p.op === 'invite' || (acc && acc.status !== 'مفعّل' && row[4] === 'مفعّل'))) mailTo = {email: email, name: p.name || p.tname || '', role: row[2]};
    } else if (p.op === 'revoke' && acc) {
      sh.getRange(acc.row, 5).setValue('موقوف');
    } else if (p.op === 'delete' && (acc || p.row)) {
      sh.deleteRow(acc ? acc.row : +p.row);
    }
  } finally { lock.releaseLock(); }
  var link = appUrl_();
  if (mailTo) {
    msg = 'السلام عليكم ' + (mailTo.name ? 'أ. ' + mailTo.name : '') + '\nتمت إضافتك في منصة ' + CFG.APP + ' بصفة: ' + mailTo.role + '.\nادخل من الرابط واكتب بريدك ' + mailTo.email + ' ليصلك رمز الدخول:\n' + link;
    MailApp.sendEmail({to: mailTo.email, subject: 'دعوة إلى منصة ' + CFG.APP, name: CFG.APP,
      htmlBody: mail_('أهلاً بك في المنصة', '<p>' + esc_(mailTo.name ? 'أ. ' + mailTo.name + '،' : '') + ' تمت إضافتك في منصة ' + CFG.APP + ' بصفة <b>' + esc_(mailTo.role) + '</b>.</p><p>افتح الرابط، واكتب بريدك هذا، فيصلك رمز الدخول. ثم أضف المنصة إلى شاشة جوالك لتفتحها كتطبيق.</p>' + btn_(link, 'الدخول إلى المنصة'))});
  }
  return ok_({accounts: readAccounts_(), message: msg, link: link});
}
function esc_(s) { return String(s || '').replace(/[&<>"]/g, function (c) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]; }); }
function btn_(href, label) { return '<p style="margin:18px 0 0"><a href="' + href + '" style="display:inline-block;background:#6a1820;color:#fff;text-decoration:none;padding:11px 20px;border-radius:10px;font-weight:700">' + label + '</a></p>'; }
function mail_(title, body) {
  return '<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;background:#f7f4f1;padding:20px"><div style="max-width:520px;margin:auto;background:#fff;border-radius:14px;overflow:hidden;border:1px solid #e5ddd5">' +
    '<div style="background:#6a1820;color:#fbf3ea;padding:16px 20px;border-bottom:3px solid #c29a2d"><div style="font-size:19px;font-weight:800">' + CFG.APP + '</div><div style="font-size:12.5px;color:#e3c9bd">مجمع اليحيى القرآني · جامع اليحيى بحي النسيم</div></div>' +
    '<div style="padding:20px;color:#2b1a17;font-size:15px;line-height:1.7"><h2 style="margin:0 0 10px;font-size:18px">' + title + '</h2>' + body + '</div></div></div>';
}

/* ================= DATA STORE (_db) ================= */
function dbSheet_() {
  var ss = master_(), sh = ss.getSheetByName('_db');
  if (!sh) { sh = ss.insertSheet('_db'); sh.getRange(1, 1, 1, 5).setValues([['collection', 'id', 'json', 'updated_at', 'by']]); sh.setFrozenRows(1); sh.hideSheet(); }
  return sh;
}
function readDb_() {
  var c = CacheService.getScriptCache(), hit = c.get('db_v');
  var sh = dbSheet_(), n = sh.getLastRow();
  var v = n > 1 ? sh.getRange(2, 1, n - 1, 3).getValues() : [];
  var db = {}, idx = {};
  for (var i = 0; i < v.length; i++) {
    var coll = v[i][0], id = String(v[i][1]); if (!coll) continue;
    var doc; try { doc = JSON.parse(v[i][2]); } catch (e) { continue; }
    (db[coll] = db[coll] || {})[id] = doc; idx[coll + '|' + id] = i + 2;
  }
  return {db: db, idx: idx, sh: sh};
}
function deepMerge_(a, b) {
  for (var k in b) {
    if (b[k] && typeof b[k] === 'object' && !Array.isArray(b[k])) a[k] = deepMerge_(a[k] && typeof a[k] === 'object' ? a[k] : {}, b[k]);
    else a[k] = b[k];
  }
  return a;
}
function bumpVersion_() { CacheService.getScriptCache().put('ver', String(Date.now()), 21600); }

/* ================= ATTENDANCE FILES ================= */
function pickSheet_(ss, fid) {
  var c = CacheService.getScriptCache(), k = 'sel_' + fid, hit = c.get(k);
  if (hit) { var o = JSON.parse(hit); var s = ss.getSheetByName(o.sheet); if (s && addDays_(o.fs, 34) >= today_()) return {sheet: s, fs: o.fs}; }
  var best = null;
  ss.getSheets().forEach(function (s) {
    if (s.getMaxColumns() < 60) return;
    var raw = s.getRange('BA3').getDisplayValue(), m = /(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/.exec(raw); if (!m) return;
    var ds = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
    if (ds <= today_() && (!best || ds > best.fs)) best = {sheet: s, fs: ds};
  });
  if (best) c.put(k, JSON.stringify({sheet: best.sheet.getName(), fs: best.fs}), 21600);
  return best;
}
function readFile_(meta, force) {
  var c = CacheService.getScriptCache(), k = 'grid_' + meta.id;
  if (!force) { var hit = c.get(k); if (hit) return JSON.parse(hit); }
  var out = {id: meta.id, file_id: meta.file_id, halaqa: meta.halaqa, stage: meta.stage, order: meta.order, teachers: meta.teachers || [], students: meta.students || [], grid: null, sheet: meta.sheet, first_sunday: meta.first_sunday, synced_at: null};
  try {
    var ss = SpreadsheetApp.openById(meta.file_id), sel = pickSheet_(ss, meta.id);
    if (sel) {
      var g = sel.sheet.getRange('A1:AW60').getDisplayValues(), last = 7;
      while (last < g.length && String(g[last][2] || '').trim()) last++;
      var grid = g.slice(0, last);
      out.grid = grid; out.sheet = sel.sheet.getName(); out.first_sunday = sel.fs;
      out.teachers = String(g[4][1] || '').split(/\n/).map(function (s) { return canon_(s); }).filter(String);
      out.students = grid.slice(7).map(function (r) { return String(r[2] || '').trim(); }).filter(String);
      if (String(g[2][12] || '').trim()) out.halaqa = String(g[2][12]).trim();
      out.synced_at = new Date().toISOString();
    }
  } catch (err) { out.error = String(err.message || err); }
  try { c.put(k, JSON.stringify(out), CFG.GRID_TTL); } catch (e) {}
  return out;
}
function allFiles_(db) {
  var metas = Object.keys(db.files || {}).map(function (id) { var f = db.files[id]; return {id: id, file_id: f.file_id, halaqa: f.halaqa, stage: f.stage, order: f.order, teachers: f.teachers, students: f.students, sheet: f.sheet, first_sunday: f.first_sunday}; });
  metas.sort(function (a, b) { return (a.order || 99) - (b.order || 99); });
  return metas.map(function (m) { return readFile_(m); });
}
function teacherFiles_(files, tname) { var n = norm_(canon_(tname)); return files.filter(function (f) { return (f.teachers || []).some(function (t) { return norm_(canon_(t)) === n; }); }); }
function tidOf_(db, name) {
  if (!name) return null; var n = norm_(canon_(name)), tk = tokens_(canon_(name)), ts = db.teachers || {};
  for (var id in ts) { var t = ts[id]; if (norm_(t.name) === n || (t.aliases || []).some(function (a) { return norm_(a) === n; })) return id; }
  for (var id2 in ts) { var k = tokens_(ts[id2].name); if (k[0] === tk[0] && k[k.length - 1] === tk[tk.length - 1]) return id2; }
  return null;
}

/* ================= BOOT ================= */
function boot_(me, p) {
  var store = readDb_(), db = store.db;
  var pub = {email: me.email, role: me.role, active: me.active, owner: me.owner, tname: me.tname, name: me.name, status: me.status};
  var files = allFiles_(db);
  if (!me.active) {
    var names = []; files.forEach(function (f) { (f.teachers || []).forEach(function (t) { if (names.indexOf(t) < 0) names.push(t); }); });
    return ok_({me: pub, teacherNames: names, link: appUrl_()});
  }
  var data = {};
  var full = me.role === 'system' || me.role === 'admin';
  var tid = tidOf_(db, me.tname);
  var athAdmin = ((db.config || {}).athar || {}).admin_teacher_id;
  if (full) {
    ['teachers', 'students', 'exams', 'athar', 'meetings', 'config', 'athar_log', 'tasmee'].forEach(function (c) { data[c] = db[c] || {}; });
    data.files = {}; files.forEach(function (f) { data.files[f.id] = f; });
  } else {
    var mine = teacherFiles_(files, me.tname), myIds = mine.map(function (f) { return f.id; });
    data.files = {}; mine.forEach(function (f) { data.files[f.id] = f; });
    data.teachers = {}; Object.keys(db.teachers || {}).forEach(function (id) { var t = db.teachers[id]; data.teachers[id] = {name: t.name, aliases: t.aliases || [], order: t.order}; });
    data.meetings = db.meetings || {}; data.config = {athar: (db.config || {}).athar || {}, targets: (db.config || {}).targets || {}};
    data.exams = {};
    Object.keys(db.exams || {}).forEach(function (id) { var e = db.exams[id];
      if (e.kind === 'داخلي' && (me.role === 'exams' || myIds.indexOf(e.fid) >= 0 || e.requested_by === me.email)) data.exams[id] = e; });
    data.athar = {}; data.athar_log = {};
    if (tid && db.athar && db.athar[tid]) data.athar[tid] = db.athar[tid];
    Object.keys(db.athar_log || {}).forEach(function (id) { if (tid && (id.indexOf(tid + '_') === 0 || tid === athAdmin)) data.athar_log[id] = db.athar_log[id]; });
    if (tid && tid === athAdmin) { data.athar = db.athar || {}; data.athar_log = db.athar_log || {}; }
    data.tasmee = {}; Object.keys(db.tasmee || {}).forEach(function (id) { if (myIds.indexOf(id.split('_')[0]) >= 0) data.tasmee[id] = db.tasmee[id]; });
  }
  var out = {me: pub, data: data, link: appUrl_(), serverTime: new Date().toISOString(), photosVer: prop_('PHOTOS_VER') || '0'};
  if (me.role === 'system') out.accounts = readAccounts_();
  var etag = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(out.data) + JSON.stringify(out.me) + (out.accounts ? JSON.stringify(out.accounts) : '')));
  if (p.etag && p.etag === etag) return ok_({same: true, etag: etag});
  out.etag = etag;
  return ok_(out);
}

/* ================= WRITES ================= */
function nextExamDayOk_(date) {
  var w = dow_(date); if (w !== 1 && w !== 3) return false;
  return today_() < date; // الرفع قبل يوم الاختبار
}
function canWrite_(me, db, coll, id, op, data, cur) {
  var r = me.role;
  if (r === 'system') return true;
  if (coll === 'config') return false;
  if (r === 'admin') return true;
  var tid = tidOf_(db, me.tname), athAdmin = ((db.config || {}).athar || {}).admin_teacher_id;
  if (coll === 'athar_log') return !!tid && (id.indexOf(tid + '_') === 0 || tid === athAdmin);
  if (coll === 'meetings') return !!tid && tid === athAdmin && op === 'update';
  if (coll === 'athar') return !!tid && tid === athAdmin;
  if (coll === 'exams') {
    var files = allFiles_(db), mine = teacherFiles_(files, me.tname).map(function (f) { return f.id; });
    if (op === 'set' && !cur) {
      if (!data || data.kind !== 'داخلي') return false;
      if (!nextExamDayOk_(data.date)) throw new Error('انتهى وقت الرفع لهذا اليوم. الرفع يكون قبل يوم الاختبار.');
      return r === 'exams' || mine.indexOf(data.fid) >= 0;
    }
    if (!cur || cur.kind !== 'داخلي') return false;
    if (r === 'exams') return true;
    if (op === 'delete') return cur.requested_by === me.email && today_() < cur.date && (cur.score == null || cur.score === '');
    return false;
  }
  return false;
}
function write_(me, p) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var store = readDb_(), db = store.db, coll = p.coll, op = p.op, id = p.id || (coll.charAt(0) + Utilities.getUuid().slice(0, 8));
    var cur = (db[coll] || {})[id];
    if (!canWrite_(me, db, coll, id, op, p.data, cur)) return {ok: false, error: 'لا تملك صلاحية هذا الإجراء'};
    if (op === 'update' && !cur) return {ok: false, error: 'السجل غير موجود'};
    var doc = op === 'delete' ? null : op === 'set' ? p.data : deepMerge_(JSON.parse(JSON.stringify(cur)), p.data);
    if (coll === 'exams' && doc && op === 'set' && !cur) { doc.requested_by = doc.requested_by || me.email; doc.requested_at = doc.requested_at || new Date().toISOString(); }
    var row = store.idx[coll + '|' + id], sh = store.sh, stamp = new Date().toISOString();
    if (op === 'delete') { if (row) sh.deleteRow(row); }
    else if (row) sh.getRange(row, 3, 1, 3).setValues([[JSON.stringify(doc), stamp, me.email]]);
    else sh.appendRow([coll, id, JSON.stringify(doc), stamp, me.email]);
    SpreadsheetApp.flush();
    (db[coll] = db[coll] || {}); if (doc) db[coll][id] = doc; else delete db[coll][id];
    try { report_(coll, db); } catch (e) {}
    return ok_({id: id, doc: doc});
  } finally { lock.releaseLock(); }
}
function tasmee_(me, p) {
  var store = readDb_(), db = store.db;
  var meta = (db.files || {})[p.fid]; if (!meta) return {ok: false, error: 'الحلقة غير موجودة'};
  var f = readFile_({id: p.fid, file_id: meta.file_id, halaqa: meta.halaqa, stage: meta.stage, order: meta.order});
  var allowed = me.role === 'system' || me.role === 'admin' || teacherFiles_([f], me.tname).length > 0;
  if (!allowed) return {ok: false, error: 'لا تملك صلاحية الرصد في هذه الحلقة'};
  var fs = f.first_sunday, diff = Math.round((new Date(p.date + 'T12:00:00+03:00') - new Date(fs + 'T12:00:00+03:00')) / 864e5);
  var w = Math.floor(diff / 7), d = diff % 7;
  if (w < 0 || w > 4 || d > 3) return {ok: false, error: 'لا يوجد عمود تسميع لهذا اليوم'};
  var col = 4 + 9 * w + 2 * d + 1; // رقم العمود في الورقة (1-based)
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(meta.file_id), sh = ss.getSheetByName(f.sheet);
    var names = sh.getRange(8, 3, 53, 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); });
    var i = names.indexOf(String(p.name).trim()); if (i < 0) return {ok: false, error: 'لم أجد الطالب في ملف التحضير'};
    sh.getRange(8 + i, col).setValue(p.value || '');
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  CacheService.getScriptCache().remove('grid_' + p.fid);
  var log = master_().getSheetByName('سجل التسميع') || master_().insertSheet('سجل التسميع');
  if (log.getLastRow() === 0) { log.setRightToLeft(true); log.appendRow(['الوقت', 'التاريخ', 'الحلقة', 'الطالب', 'التسميع', 'رصده']); }
  log.appendRow([Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'), p.date, meta.halaqa, p.name, p.value || '(مسح)', me.email]);
  return ok_({file: readFile_({id: p.fid, file_id: meta.file_id, halaqa: meta.halaqa, stage: meta.stage, order: meta.order}, true)});
}

/* ================= TEACHER PHOTOS ================= */
/* الصور تُحفظ في مجلد «صور المعلمين» بجانب قاعدة البيانات، وتسجَّل في _db كرقم ملف فقط */
function photoFolder_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('PHOTO_FOLDER');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var master = DriveApp.getFileById(CFG.MASTER_ID), parents = master.getParents();
  var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var ex = parent.getFoldersByName('صور المعلمين'), folder = ex.hasNext() ? ex.next() : parent.createFolder('صور المعلمين');
  props.setProperty('PHOTO_FOLDER', folder.getId());
  return folder;
}
function setDoc_(coll, id, doc, by) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var store = readDb_(), row = store.idx[coll + '|' + id], sh = store.sh, stamp = new Date().toISOString();
    if (!doc) { if (row) sh.deleteRow(row); }
    else if (row) sh.getRange(row, 3, 1, 3).setValues([[JSON.stringify(doc), stamp, by || '']]);
    else sh.appendRow([coll, id, JSON.stringify(doc), stamp, by || '']);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}
function bumpPhotos_() { PropertiesService.getScriptProperties().setProperty('PHOTOS_VER', String(Date.now())); }
function photoSrc_(fileId) {
  var c = CacheService.getScriptCache(), k = 'ph_' + fileId, hit = c.get(k); if (hit) return hit;
  var f = DriveApp.getFileById(fileId), blob = f.getSize() <= 70000 ? f.getBlob() : f.getThumbnail();
  if (!blob) return null;
  var src = 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
  try { c.put(k, src, 21600); } catch (e) {}
  return src;
}
function photos_(me) {
  var db = readDb_().db, all = db.photos || {}, out = {}, ids = Object.keys(all);
  if (me.role !== 'system' && me.role !== 'admin') { var tid = tidOf_(db, me.tname); ids = tid && all[tid] ? [tid] : []; }
  ids.forEach(function (k) { try { var s = photoSrc_(all[k].file); if (s) out[k] = s; } catch (e) {} });
  return ok_({photos: out, ver: prop_('PHOTOS_VER') || '0'});
}
function setPhoto_(me, p) {
  if (me.role !== 'system' && me.role !== 'admin') return {ok: false, error: 'forbidden'};
  var db = readDb_().db, t = (db.teachers || {})[p.tid]; if (!t) return {ok: false, error: 'المعلم غير موجود'};
  if (!p.src) { setDoc_('photos', p.tid, null, me.email); bumpPhotos_(); return ok_({}); }
  if (String(p.src).indexOf('data:image/') !== 0 || p.src.length > 400000) return {ok: false, error: 'الصورة كبيرة أو غير صالحة'};
  var parts = p.src.split(','), type = parts[0].replace('data:', '').replace(';base64', '');
  var file = photoFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(parts[1]), type, t.name + '.jpg'));
  setDoc_('photos', p.tid, {file: file.getId(), from: 'upload', at: new Date().toISOString()}, me.email);
  bumpPhotos_(); return ok_({});
}
function importPhotos_(me) {
  if (me.role !== 'system' && me.role !== 'admin') return {ok: false, error: 'forbidden'};
  var db = readDb_().db, best = {}, unmatched = [];
  var isFace = function (f) { var n = f.getName(), m = String(f.getMimeType()); return m.indexOf('image/') === 0 && /صورة\s*شخصي|شخصية|personal|photo/i.test(n) && !/هوية|بطاقة|إقامة|اقامة/.test(n); };
  var consider = function (tid, f) { if (!tid) return; if (!best[tid] || f.getLastUpdated() > best[tid].getLastUpdated()) best[tid] = f; };
  // 1) مجلد «صور وملفات المعلمين»: مجلد لكل معلم، فيه ملف «صورة شخصية …»
  var walk = function (folder, depth) {
    var tidF = tidOf_(db, folder.getName().replace(/^\d+\s*/, ''));
    var files = folder.getFiles();
    while (files.hasNext()) { var f = files.next(); if (!isFace(f)) continue;
      var t = tidF || tidOf_(db, f.getName().replace(/\.[^.]+$/, '').replace(/صورة\s*شخصية\s*[-–]?\s*/, '').replace(/\(\d+\)/, '').trim());
      if (t) consider(t, f); else unmatched.push(f.getName()); }
    if (depth < 2) { var subs = folder.getFolders(); while (subs.hasNext()) walk(subs.next(), depth + 1); }
  };
  try { walk(DriveApp.getFolderById(CFG.TEACHER_FILES_FOLDER), 0); } catch (e) { unmatched.push('تعذّر فتح مجلد «صور وملفات المعلمين»'); }
  // 2) مجلد «صور المعلمين» الخاص بالمنصة: صور مسمّاة باسم المعلم
  var folder = photoFolder_(), fs = folder.getFiles();
  while (fs.hasNext()) { var g = fs.next(); if (String(g.getMimeType()).indexOf('image/') !== 0) continue;
    var tg = tidOf_(db, g.getName().replace(/\.[^.]+$/, '').replace(/^أ\.?\s*/, '')); if (tg) consider(tg, g); else unmatched.push(g.getName()); }
  var matched = [];
  Object.keys(best).forEach(function (tid) { var cur = (db.photos || {})[tid]; if (cur && cur.from === 'upload') return; // لا نستبدل صورة رُفعت يدوياً
    setDoc_('photos', tid, {file: best[tid].getId(), from: 'drive', at: new Date().toISOString()}, me.email); matched.push(db.teachers[tid].name); });
  bumpPhotos_();
  return ok_({matched: matched, unmatched: unmatched, folderUrl: folder.getUrl()});
}

/* ================= READABLE REPORT TABS ================= */
function writeTab_(name, rows) {
  var ss = master_(), sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.setRightToLeft(true); sh.clearContents();
  if (rows.length) sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold'); sh.setFrozenRows(1);
}
function grade_(s) { return s === null || s === '' || s === undefined ? '' : s >= 90 ? 'ممتاز' : s >= 80 ? 'جيد جداً' : s >= 70 ? 'جيد' : s >= 60 ? 'مقبول' : 'لم يجتز'; }
function report_(coll, db) {
  if (coll === 'students') writeTab_('الطلاب المسجلون', [['الرقم', 'الاسم', 'الصف', 'المرحلة', 'الحلقة الرسمية', 'الجنسية', 'الحالة', 'رقم الجمعية']].concat(Object.keys(db.students || {}).sort().map(function (id) { var s = db.students[id]; return [id.replace(/^s/, ''), s.name || '', s.grade || '', s.stage || '', s.halaqa || '', s.nationality || '', s.status || '', s.assoc_no || '']; })));
  if (coll === 'teachers') writeTab_('المعلمون', [['الاسم', 'الحلقة الرسمية', 'المرحلة', 'الصفة', 'بطاقة معلم متطوع', 'انتهاؤها', 'بطاقة البركة', 'انتهاؤها']].concat(Object.keys(db.teachers || {}).map(function (id) { var t = db.teachers[id]; return ['أ. ' + (t.name || ''), t.halaqa || '', t.stage || '', t.role || '', t.card_ilm || '', t.card_ilm_exp || '', t.card_baraka || '', t.card_baraka_exp || '']; })));
  if (coll === 'exams') writeTab_('الاختبارات', [['التاريخ', 'النوع', 'رقم الطالب', 'الطالب', 'الفرع', 'الحلقة', 'المرحلة', 'المختبر', 'الطريقة', 'الدرجة', 'التقدير', 'رفعه', 'ملاحظات']].concat(Object.keys(db.exams || {}).map(function (id) { return db.exams[id]; }).sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); }).map(function (e) { return [e.date || '', e.kind === 'داخلي' ? 'داخلي' : 'الجمعية', String(e.student_id || '').replace(/^s/, ''), e.student || '', e.branch || '', e.halaqa || '', e.stage || '', e.examiner || '', e.mode || '', e.score == null ? '' : e.score, e.grade || grade_(e.score), e.requested_by || '', e.notes || '']; })));
}

/* ================= EXAM COMMITTEE ================= */
function committeeSummary_(db, date) {
  var list = Object.keys(db.exams || {}).map(function (id) { return db.exams[id]; }).filter(function (e) { return e.kind === 'داخلي' && e.date === date && (e.score == null || e.score === ''); });
  var byBranch = {}, byHalaqa = {};
  list.forEach(function (e) { byBranch[e.branch] = (byBranch[e.branch] || 0) + 1; (byHalaqa[e.halaqa] = byHalaqa[e.halaqa] || []).push(e); });
  return {list: list, byBranch: byBranch, byHalaqa: byHalaqa};
}
function notifyCommittee_(me, date) {
  if (['system', 'admin', 'exams'].indexOf(me.role) < 0) return {ok: false, error: 'forbidden'};
  return ok_({sent: sendCommittee_(date)});
}
function sendCommittee_(date) {
  var db = readDb_().db, s = committeeSummary_(db, date);
  var to = readAccounts_().filter(function (a) { return (a.role === 'exams' || a.role === 'admin') && a.status === 'مفعّل' && a.email; }).map(function (a) { return a.email; });
  if (!to.length || !s.list.length) return 0;
  var day = Utilities.formatDate(new Date(date + 'T12:00:00+03:00'), 'Asia/Riyadh', 'yyyy/MM/dd');
  var body = '<p>عدد المختبرين: <b>' + s.list.length + '</b></p><p><b>حسب الفرع:</b> ' + Object.keys(s.byBranch).sort(function (a, b) { return a - b; }).map(function (b) { return 'الفرع ' + b + ': ' + s.byBranch[b]; }).join(' · ') + '</p>' +
    Object.keys(s.byHalaqa).map(function (h) { return '<p style="margin:10px 0 4px"><b>' + esc_(h) + '</b></p><ul style="margin:0;padding-right:18px">' + s.byHalaqa[h].map(function (e) { return '<li>' + esc_(e.student) + ' · الفرع ' + esc_(e.branch) + '</li>'; }).join('') + '</ul>'; }).join('');
  MailApp.sendEmail({to: to.join(','), subject: 'كشف الاختبار الداخلي ' + day + ' · ' + s.list.length + ' طالب', name: CFG.APP, htmlBody: mail_('كشف الاختبار الداخلي · ' + day, body + btn_(appUrl_(), 'فتح المنصة'))});
  return to.length;
}
/** يعمل تلقائياً كل مساء: إن كان الغد اثنين أو أربعاء يُرسل الكشف للجنة */
function eveningDigest() {
  var tomorrow = addDays_(today_(), 1), w = dow_(tomorrow);
  if (w === 1 || w === 3) sendCommittee_(tomorrow);
}

/* ================= SETUP (شغّلها مرة واحدة) ================= */
/** نقل البيانات الحالية (المعلمين، الطلاب، الاختبارات، أثر...) من ورقة _seed مرة واحدة */
function importSeed_() {
  var ss = master_(), seed = ss.getSheetByName('_seed'); if (!seed) return;
  var db = dbSheet_();
  if (db.getLastRow() <= 1) {
    var b64 = String(seed.getRange('A1').getValue() || '');
    var json = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString('UTF-8');
    var rows = JSON.parse(json);
    if (rows.length) db.getRange(2, 1, rows.length, 5).setValues(rows);
    SpreadsheetApp.flush();
    var d = readDb_().db; ['students', 'teachers', 'exams'].forEach(function (c) { try { report_(c, d); } catch (e) {} });
  }
  ss.deleteSheet(seed);
}
function setup() {
  accSheet_(); dbSheet_();
  var acc = readAccounts_(), sh = accSheet_();
  CFG.OWNERS.forEach(function (o) { if (!acc.some(function (a) { return a.email === o; })) sh.appendRow([o, 'مسؤول المنصة', 'النظام', '', 'مفعّل', Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'), '', '']); });
  [['أسامة بن محفوظ'], ['حذيفة باحداد']].forEach(function (n) { if (!acc.some(function (a) { return a.name === n[0]; })) sh.appendRow(['', n[0], 'لجنة الاختبارات', '', 'بانتظار البريد', Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'), '', 'أضف بريده من بوابة الحسابات']); });
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'eveningDigest') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('eveningDigest').timeBased().atHour(21).everyDays(1).inTimezone('Asia/Riyadh').create();
  importSeed_();
  try { photoFolder_(); } catch (e) { Logger.log('مجلد الصور: ' + e); }
  var n = dbSheet_().getLastRow() - 1;
  Logger.log('تم الإعداد. عدد السجلات في قاعدة البيانات: ' + n);
  // سطر للتأكد من صلاحيات البريد والتطبيق
  if (false) { UrlFetchApp.fetch('https://www.google.com'); MailApp.getRemainingDailyQuota(); }
}

/** تشغيل مرة واحدة لمنح صلاحية الدرايف وإنشاء مجلد «صور المعلمين» */
function authorizeDrive() {
  Logger.log('صلاحية الدرايف مفعّلة: ' + DriveApp.getRootFolder().getName());
  Logger.log('مجلد الصور: ' + photoFolder_().getUrl());
}
