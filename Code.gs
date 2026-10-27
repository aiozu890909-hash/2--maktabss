/**
 * SmartSchool Attendance — Google Apps Script backend (Web App)
 * Sheets: Students, Attendance, Events, Audit.  Drive: SmartSchool/Faces va SmartSchool/Attendance Archive/YYYY/MM/DD
 * O'rnatish: setup() ni bir marta ishga tushiring -> Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
 */
/* ====== SHU IKKITA QIYMATNI YOZING ======
 * SHEET_ID : Google Sheets havolasidagi  .../spreadsheets/d/<SHU YERDA>/edit
 * FOLDER_ID: Drive papka havolasidagi    .../folders/<SHU YERDA>
 * Bo'sh qoldirsangiz: jadvalning o'zi (getActive) va Drive'dagi "SmartSchool" papkasi ishlatiladi. */
const SHEET_ID = '';
const FOLDER_ID = '';

const P = PropertiesService.getScriptProperties();
const DEF = { START: '08:30', END: '13:00', TZ: 'Asia/Tashkent', MINGAP: '10', SCHOOL: 'Maktab' };
const cfg = k => P.getProperty(k) || DEF[k];
const HEAD = {
  Students: ['id', 'ism', 'familiya', 'sinf', 'photoId', 'extra', 'active', 'createdAt', 'updatedAt'],
  Attendance: ['Sana', 'ID', 'Ism', 'Familiya', 'Sinf', 'Kirish', 'Chiqish', 'Kechikish (daq)', 'Erta ketish (daq)', 'Holat', 'Yangilangan'],
  Events: ['eventId', 'ts', 'studentId', 'type', 'note'],
  Audit: ['ts', 'user', 'action', 'details']
};
let _ss;
const SS = () => _ss || (_ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive());
const sh = n => SS().getSheetByName(n);
const fmt = (ts, f) => Utilities.formatDate(new Date(ts), cfg('TZ'), f);
const mins = t => { const p = String(t).split(':'); return +p[0] * 60 + +p[1]; };
const nowStr = () => fmt(Date.now(), 'yyyy-MM-dd HH:mm:ss');
const hash = (pw) => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw + P.getProperty('SALT')));

/* ---------- O'rnatish ---------- */
function setup() {
  const ss = SS();
  Object.keys(HEAD).forEach(n => {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (s.getLastRow() === 0) { s.appendRow(HEAD[n]); s.setFrozenRows(1); s.getRange(1, 1, 1, HEAD[n].length).setFontWeight('bold'); }
  });
  const A = sh('Attendance');
  [1, 2, 6, 7].forEach(c => A.getRange(1, c, A.getMaxRows(), 1).setNumberFormat('@')); // sana/ID/vaqt matn bo'lib qolsin
  sh('Students').getRange(1, 1, sh('Students').getMaxRows(), 1).setNumberFormat('@');
  sh('Events').getRange(1, 3, sh('Events').getMaxRows(), 1).setNumberFormat('@');
  if (!P.getProperty('SALT')) P.setProperty('SALT', Utilities.getUuid());
  if (!P.getProperty('DEVICE_KEY')) P.setProperty('DEVICE_KEY', Utilities.getUuid() + Utilities.getUuid());
  if (!P.getProperty('ADMIN_USER')) {
    const pw = Utilities.getUuid().slice(0, 10);
    setAdmin('admin', pw);
    Logger.log('ADMIN login: admin / ' + pw + '  (parolni setAdmin(login, parol) bilan almashtiring)');
  }
  Logger.log('DEVICE_KEY (Python uchun): ' + P.getProperty('DEVICE_KEY'));
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyClose').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyClose').timeBased().everyDays(1).atHour(14).create(); // 13:00 dan keyin
  folder(['Faces']);
}
function setAdmin(user, pass) { P.setProperty('ADMIN_USER', user); P.setProperty('ADMIN_HASH', hash(pass)); }

/* ---------- Drive yordamchilari ---------- */
function ensure(parent, name) { const it = parent.getFoldersByName(String(name)); return it.hasNext() ? it.next() : parent.createFolder(String(name)); }
const ROOT = () => FOLDER_ID ? DriveApp.getFolderById(FOLDER_ID) : ensure(DriveApp.getRootFolder(), 'SmartSchool');
function folder(path) { return path.reduce((p, n) => ensure(p, n), ROOT()); }
function savePhoto(id, b64, oldId) {
  const blob = Utilities.newBlob(Utilities.base64Decode(String(b64).replace(/^data:.*?,/, '')), 'image/jpeg', id + '.jpg');
  const f = folder(['Faces']).createFile(blob); // shaxsiy (ochiq havolasiz)
  if (oldId) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) {} }
  return f.getId();
}

/* ---------- Qoidalar ---------- */
function compute(inT, outT) {
  const s = mins(cfg('START')), e = mins(cfg('END'));
  let late = 0, early = 0, st = "O'z vaqtida";
  if (inT) { late = Math.max(0, mins(inT) - s); if (late) st = 'Kechikdi'; }
  if (outT) { early = Math.max(0, e - mins(outT)); if (early) st = late ? 'Kechikdi + Erta ketdi' : 'Erta ketdi'; }
  return { late: late, early: early, st: st };
}

/* ---------- Ma'lumot o'qish ---------- */
function students() {
  const s = sh('Students'), n = s.getLastRow(); if (n < 2) return [];
  return s.getRange(2, 1, n - 1, 9).getValues().map((r, i) => ({
    row: i + 2, id: String(r[0]), ism: r[1], familiya: r[2], sinf: r[3], photoId: r[4], extra: r[5],
    active: r[6] === true || r[6] === 'TRUE', createdAt: String(r[7]), updatedAt: String(r[8])
  }));
}
function attRows(lastN) {
  const A = sh('Attendance'), n = A.getLastRow(); if (n < 2) return [];
  const start = lastN ? Math.max(2, n - lastN + 1) : 2;
  return A.getRange(start, 1, n - start + 1, 11).getValues().map((r, i) => ({ row: start + i, v: r.map(x => x instanceof Date ? fmt(x, 'yyyy-MM-dd') : String(x)) }));
}
function findRow(day, id) {
  const rows = attRows(2000);
  for (let i = rows.length - 1; i >= 0; i--) if (rows[i].v[0] === day && rows[i].v[1] === String(id)) return rows[i];
  return null;
}

/* ---------- Hodisa yozish (idempotent) ---------- */
function logEvent(d) {
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const ev = sh('Events'), n = ev.getLastRow();
    if (n > 1) {
      const st = Math.max(2, n - 2999);
      if (ev.getRange(st, 1, n - st + 1, 1).getValues().some(r => r[0] === d.eventId)) return { dup: true };
    }
    const day = fmt(d.ts, 'yyyy-MM-dd'), t = fmt(d.ts, 'HH:mm:ss');
    let type = 'UNKNOWN', note = '';
    const stu = d.studentId ? students().find(s => s.id === String(d.studentId)) : null;
    if (d.studentId && (!stu || !stu.active)) note = 'faol bo\'lmagan yoki noma\'lum ID';
    if (stu && stu.active) {
      const A = sh('Attendance'), r = findRow(day, stu.id);
      if (!r) {
        const c = compute(t);
        A.appendRow([day, stu.id, stu.ism, stu.familiya, stu.sinf, t, '', c.late, '', c.st, nowStr()]);
        type = 'IN';
      } else if (r.v[5] && !r.v[6]) {
        if (mins(t) - mins(r.v[5]) < +cfg('MINGAP')) { type = 'IGNORED'; note = 'juda yaqin takror'; }
        else {
          const c = compute(r.v[5], t);
          A.getRange(r.row, 7, 1, 5).setValues([[t, c.late, c.early, c.st, nowStr()]]);
          type = 'OUT';
        }
      } else { type = 'REVIEW'; note = 'kunning 3+ aniqlashi — admin tekshiradi'; }
    }
    ev.appendRow([d.eventId, d.ts, d.studentId || '', type, note]);
    return { type: type, time: t };
  } finally { lock.releaseLock(); }
}

/* ---------- Kunni yopish va arxivlash ---------- */
function closeDay(day, user) {
  const A = sh('Attendance'), have = {};
  attRows(3000).forEach(r => { if (r.v[0] === day) have[r.v[1]] = 1; });
  students().filter(s => s.active && !have[s.id]).forEach(s =>
    A.appendRow([day, s.id, s.ism, s.familiya, s.sinf, '', '', '', '', 'Kelmadi', nowStr()]));
  const rows = attRows(3000).filter(r => r.v[0] === day).map(r => r.v);
  const csv = '\ufeff' + [HEAD.Attendance].concat(rows).map(r => r.map(c => '"' + String(c).replace(/"/g, '""') + '"').join(',')).join('\n');
  const [y, m, dd] = day.split('-');
  const dir = folder(['Attendance Archive', y, m, dd]), name = 'davomat-' + day + '.csv';
  const old = dir.getFilesByName(name); while (old.hasNext()) old.next().setTrashed(true);
  const f = dir.createFile(name, csv, MimeType.CSV);
  audit(user, 'closeDay', { day: day, rows: rows.length });
  return { url: f.getUrl(), rows: rows.length };
}
function dailyClose() {
  const wd = +fmt(Date.now(), 'u'); if (wd > 5) return; // shanba-yakshanba
  closeDay(fmt(Date.now(), 'yyyy-MM-dd'), 'auto');
}
function walk(f, out) {
  const fi = f.getFiles(); while (fi.hasNext()) { const x = fi.next(); out.push({ name: x.getName(), url: x.getUrl(), date: fmt(x.getLastUpdated(), 'yyyy-MM-dd HH:mm') }); }
  const fo = f.getFolders(); while (fo.hasNext()) walk(fo.next(), out);
}

/* ---------- Auth ---------- */
function audit(user, action, details) { sh('Audit').appendRow([nowStr(), user, action, JSON.stringify(details || {}).slice(0, 500)]); }
function login(d) {
  if (d.user === P.getProperty('ADMIN_USER') && hash(String(d.pass || '')) === P.getProperty('ADMIN_HASH')) {
    const tok = Utilities.getUuid() + Utilities.getUuid();
    CacheService.getScriptCache().put('t:' + tok, d.user, 21600);
    audit(d.user, 'login', {});
    return { ok: true, token: tok };
  }
  Utilities.sleep(1500); throw new Error('Login yoki parol noto\'g\'ri');
}

/* ---------- Amallar ---------- */
const ACT = {
  faces: () => ({ students: students().filter(s => s.active && s.photoId).map(s => ({ id: s.id, ism: s.ism, familiya: s.familiya, sinf: s.sinf, photoId: s.photoId })) }),
  photo: d => {
    if (!students().some(s => s.photoId === d.photoId)) throw new Error('Rasm topilmadi');
    return { data: Utilities.base64Encode(DriveApp.getFileById(d.photoId).getBlob().getBytes()) };
  },
  logEvent: d => logEvent(d),
  heartbeat: d => { P.setProperty('HB', JSON.stringify({ t: Date.now(), fps: d.fps, cam: d.cam, host: d.host })); },
  listStudents: () => ({ students: students() }),
  addStudent: (d, u) => {
    const id = String(d.id || '').trim();
    if (!id || !d.ism || !d.familiya || !d.sinf) throw new Error("Majburiy maydonlar to'ldirilmagan");
    if (students().some(s => s.id === id)) throw new Error('Bu ID allaqachon mavjud');
    if (!d.photo) throw new Error('Yuz rasmi kerak');
    const pid = savePhoto(id, d.photo), now = nowStr();
    sh('Students').appendRow([id, d.ism.trim(), d.familiya.trim(), d.sinf.trim(), pid, d.extra || '', true, now, now]);
    audit(u, 'addStudent', { id: id });
  },
  updateStudent: (d, u) => {
    const s = students().find(x => x.id === String(d.id)); if (!s) throw new Error('Topilmadi');
    const pid = d.photo ? savePhoto(s.id, d.photo, s.photoId) : s.photoId;
    sh('Students').getRange(s.row, 1, 1, 9).setValues([[s.id, d.ism || s.ism, d.familiya || s.familiya, d.sinf || s.sinf, pid, d.extra != null ? d.extra : s.extra, s.active, s.createdAt, nowStr()]]);
    audit(u, 'updateStudent', { id: s.id });
  },
  deactivate: (d, u) => {
    const s = students().find(x => x.id === String(d.id)); if (!s) throw new Error('Topilmadi');
    sh('Students').getRange(s.row, 7).setValue(false); sh('Students').getRange(s.row, 9).setValue(nowStr());
    audit(u, 'deactivate', { id: s.id }); // tarixiy davomat saqlanadi
  },
  today: () => {
    const day = fmt(Date.now(), 'yyyy-MM-dd'), rows = attRows(2000).filter(r => r.v[0] === day).map(r => r.v);
    const total = students().filter(s => s.active).length, present = rows.filter(r => r[5]);
    const hb = JSON.parse(P.getProperty('HB') || 'null');
    let drive = true; try { folder([]); } catch (e) { drive = false; }
    const ev = sh('Events'), n = ev.getLastRow(), map = {}; students().forEach(s => map[s.id] = s);
    const recent = n < 2 ? [] : ev.getRange(Math.max(2, n - 29), 1, Math.min(n - 1, 30), 5).getValues().reverse().map(r => {
      const s = map[String(r[2])] || {};
      return { ts: r[1], time: fmt(r[1], 'HH:mm:ss'), id: String(r[2]), ism: s.ism || '', familiya: s.familiya || '', sinf: s.sinf || '', type: r[3], note: r[4] };
    });
    return {
      day: day, total: total, present: present.length, late: rows.filter(r => +r[7] > 0).length,
      early: rows.filter(r => +r[8] > 0).length, notYet: Math.max(0, total - present.length), recent: recent,
      hb: hb ? { online: Date.now() - hb.t < 90000, fps: hb.fps, cam: hb.cam, host: hb.host } : { online: false },
      sheets: true, drive: drive, hours: cfg('START') + ' - ' + cfg('END')
    };
  },
  history: d => {
    const rows = attRows(0).map(r => r.v).filter(v => (!d.from || v[0] >= d.from) && (!d.to || v[0] <= d.to) && (!d.sinf || v[4] === d.sinf));
    return { head: HEAD.Attendance, rows: rows.reverse() };
  },
  editAttendance: (d, u) => {
    const r = findRow(d.day, d.id), t1 = d.in || '', t2 = d.out || '';
    const c = t1 ? compute(t1, t2 || null) : { late: '', early: '', st: 'Kelmadi' };
    if (!r) {
      const s = students().find(x => x.id === String(d.id)); if (!s) throw new Error('O\'quvchi topilmadi');
      sh('Attendance').appendRow([d.day, s.id, s.ism, s.familiya, s.sinf, t1, t2, c.late, t2 ? c.early : '', c.st, nowStr()]);
    } else sh('Attendance').getRange(r.row, 6, 1, 6).setValues([[t1, t2, c.late, t2 ? c.early : '', c.st, nowStr()]]);
    audit(u, 'editAttendance', d);
  },
  review: () => {
    const ev = sh('Events'), n = ev.getLastRow(); if (n < 2) return { events: [] };
    const map = {}; students().forEach(s => map[s.id] = s);
    const all = ev.getRange(Math.max(2, n - 499), 1, Math.min(n - 1, 500), 5).getValues().reverse();
    return { events: all.filter(r => ['REVIEW', 'UNKNOWN', 'IGNORED'].indexOf(r[3]) > -1).slice(0, 100).map(r => {
      const s = map[String(r[2])] || {};
      return { ts: fmt(r[1], 'yyyy-MM-dd HH:mm:ss'), id: String(r[2]), name: (s.ism || '') + ' ' + (s.familiya || ''), sinf: s.sinf || '', type: r[3], note: r[4] };
    }) };
  },
  closeDay: (d, u) => closeDay(d.day, u),
  archive: () => { const out = []; walk(folder(['Attendance Archive']), out); out.sort((a, b) => b.name.localeCompare(a.name)); return { files: out }; },
  getSettings: () => ({ SCHOOL: cfg('SCHOOL'), START: cfg('START'), END: cfg('END'), TZ: cfg('TZ'), MINGAP: cfg('MINGAP') }),
  setSettings: (d, u) => { ['SCHOOL', 'START', 'END', 'TZ', 'MINGAP'].forEach(k => { if (d[k]) P.setProperty(k, String(d[k])); }); audit(u, 'setSettings', d); }
};
const DEVICE = ['faces', 'photo', 'logEvent', 'heartbeat'];

function route(d) {
  if (d.action === 'login') return login(d);
  let user = null;
  if (d.key && d.key === P.getProperty('DEVICE_KEY') && DEVICE.indexOf(d.action) > -1) user = 'device';
  else { user = CacheService.getScriptCache().get('t:' + d.token); if (!user) throw new Error('AUTH'); }
  const f = ACT[d.action]; if (!f) throw new Error("Noma'lum amal");
  return Object.assign({ ok: true }, f(d, user) || {});
}
function doPost(e) {
  let out;
  try { out = route(JSON.parse(e.postData.contents)); }
  catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return ContentService.createTextOutput(JSON.stringify({ ok: true, service: 'SmartSchool Attendance' })).setMimeType(ContentService.MimeType.JSON); }
