/**
 * SmartSchool Attendance — Google Apps Script backend (Web App)
 * Sheets: Students, Attendance, Events, Audit.  Drive: SmartSchool/Faces va SmartSchool/Attendance Archive/YYYY/MM/DD
 * Python (camera.py) davomatni Sheets'ga TO'G'RIDAN-TO'G'RI yozadi va Drive'dan rasm oladi; bu skript faqat SAYT uchun.
 * O'rnatish: setup() ni bir marta ishga tushiring -> Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
 */
/* ====== SHU IKKITA QIYMATNI YOZING ======
 * SHEET_ID : Google Sheets havolasidagi  .../spreadsheets/d/<SHU YERDA>/edit
 * FOLDER_ID: Drive papka havolasidagi    .../folders/<SHU YERDA>
 * Bo'sh qoldirsangiz: jadvalning o'zi (getActive) va Drive'dagi "SmartSchool" papkasi ishlatiladi. */
const SHEET_ID = '';
const FOLDER_ID = '';

const P = PropertiesService.getScriptProperties();
const DEF = { START: '08:30', END: '13:00', TZ: 'Asia/Tashkent', MINGAP: '2', SCHOOL: 'Maktab', CLOSE_HOUR: '0' };
let _cfg = null;
const cfg = k => {   // sozlamalar 'Settings' varag'ida (Python ham shuni o'qiydi)
  if (!_cfg) {
    _cfg = {}; const s = sh('Settings');
    if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0] !== '' && r[1] !== '') _cfg[r[0]] = String(r[1]); });
  }
  return _cfg[k] || DEF[k];
};
function setCfg(k, v) {
  const s = sh('Settings'), n = s.getLastRow();
  const keys = n > 1 ? s.getRange(2, 1, n - 1, 1).getValues().map(r => r[0]) : [];
  const i = keys.indexOf(k);
  if (i > -1) s.getRange(i + 2, 2).setValue(String(v)); else s.appendRow([k, String(v)]);
  _cfg = null;
}
const HEAD = {
  Students: ['id', 'ism', 'familiya', 'sinf', 'photoId', 'extra', 'active', 'createdAt', 'updatedAt'],
  Attendance: ['Sana', 'ID', 'Ism', 'Familiya', 'Sinf', 'Kirish', 'Chiqish', 'Kechikish (daq)', 'Erta ketish (daq)', 'Holat', 'Yangilangan'],
  Events: ['eventId', 'ts', 'studentId', 'type', 'note'],
  Audit: ['ts', 'user', 'action', 'details'],
  Settings: ['key', 'value']
};
let _ss;
const SS = () => _ss || (_ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive());
const sh = n => SS().getSheetByName(n);
const fmt = (ts, f) => Utilities.formatDate(new Date(ts), cfg('TZ'), f);
const secs = t => { const p = String(t).split(':'); return +p[0] * 3600 + +p[1] * 60 + (+p[2] || 0); };
const mins = t => { const p = String(t).split(':'); return +p[0] * 60 + +p[1]; };
const nowStr = () => fmt(Date.now(), 'yyyy-MM-dd HH:mm:ss');
const salt = () => { let s = P.getProperty('SALT'); if (!s) { s = Utilities.getUuid(); P.setProperty('SALT', s); } return s; };
const hash = (pw) => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw + salt()));

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
  const S = sh('Settings'); S.getRange(1, 2, S.getMaxRows(), 1).setNumberFormat('@'); // '08:30' vaqt bo'lib ketmasin
  const have = S.getLastRow() > 1 ? S.getRange(2, 1, S.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  ['SCHOOL', 'START', 'END', 'TZ', 'MINGAP', 'CLOSE_HOUR', 'HB_T', 'HB_FPS', 'HB_CAM', 'HB_HOST'].forEach(k => { if (have.indexOf(k) < 0) S.appendRow([k, DEF[k] || '']); });
  if (!P.getProperty('ADMIN_USER')) {
    const pw = Utilities.getUuid().slice(0, 10);
    setAdmin('admin', pw);
    Logger.log('ADMIN login: admin / ' + pw + '  (parolni setAdmin(login, parol) bilan almashtiring)');
  }
  if (!P.getProperty('DEVICE_KEY')) P.setProperty('DEVICE_KEY', Utilities.getUuid() + Utilities.getUuid());
  Logger.log('DEVICE_KEY (Python, JSON siz rejim uchun): ' + P.getProperty('DEVICE_KEY'));
  installTrigger();
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

/* ---------- Hodisa yozish (Python 'Apps Script orqali' rejimida; idempotent) ---------- */
function logEvent(d) {
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const ev = sh('Events'), n = ev.getLastRow();
    if (n > 1) {
      const st = Math.max(2, n - 2999);
      if (ev.getRange(st, 1, n - st + 1, 1).getValues().some(r => r[0] === d.eventId)) return { dup: true };
    }
    const day = fmt(d.ts, 'yyyy-MM-dd'), t = fmt(d.ts, 'HH:mm:ss');
    let type = 'UNKNOWN', note = d.studentId ? '' : String(d.note || '').slice(0, 120);
    const stu = d.studentId ? students().find(s => s.id === String(d.studentId)) : null;
    if (stu && stu.active) {
      const A = sh('Attendance'), r = findRow(day, stu.id), dir = String(d.dir || '');
      if (dir === 'IN') {                                  // KIRISH kamerasi
        if (!r) {
          const c = compute(t);
          A.appendRow([day, stu.id, stu.ism, stu.familiya, stu.sinf, t, '', c.late, '', c.st, nowStr()]); type = 'IN';
        } else if (r.v[5] && !r.v[6]) { type = 'IGNORED'; note = 'allaqachon kirgan'; }
        else if (r.v[5] && r.v[6]) { type = 'REVIEW'; note = 'chiqib ketgandan keyin qayta kirish — admin tekshiradi'; }
        else { type = 'REVIEW'; note = 'bu kunga yozuv bor (kirishsiz) — admin tekshiradi'; }
      } else if (dir === 'OUT') {                          // CHIQISH kamerasi
        if (!r || !r.v[5]) { type = 'REVIEW'; note = "kirish yozilmagan, chiqish kamerasi ko'rdi — admin tekshiradi"; }
        else if (r.v[6]) { type = 'IGNORED'; note = 'allaqachon chiqqan'; }
        else if (secs(t) - secs(r.v[5]) < +cfg('MINGAP') * 60) { type = 'IGNORED'; note = 'kirishdan keyin juda yaqin'; }
        else {
          const c = compute(r.v[5], t);
          A.getRange(r.row, 7, 1, 5).setValues([[t, c.late, c.early, c.st, nowStr()]]); type = 'OUT';
        }
      } else if (!r) {                                     // bitta kamera rejimi: 1-aniqlanish = kirish
        const c = compute(t);
        A.appendRow([day, stu.id, stu.ism, stu.familiya, stu.sinf, t, '', c.late, '', c.st, nowStr()]);
        type = 'IN';
      } else if (r.v[5] && !r.v[6]) {                      // 2-aniqlanish = chiqish
        if (secs(t) - secs(r.v[5]) < +cfg('MINGAP') * 60) { type = 'IGNORED'; note = 'juda yaqin takror'; }
        else {
          const c = compute(r.v[5], t);
          A.getRange(r.row, 7, 1, 5).setValues([[t, c.late, c.early, c.st, nowStr()]]);
          type = 'OUT';
        }
      } else { type = 'REVIEW'; note = 'kunning 3+ aniqlashi — admin tekshiradi'; }
      if ((type === 'IN' || type === 'OUT') && d.cam) note = String(d.cam) + ' kamerasi';
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
/** Kunni yopish vaqti (Settings > CLOSE_HOUR, soat 0-23). Trigger shu soatda ishga tushadi. */
function installTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyClose').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyClose').timeBased().everyDays(1).atHour(+cfg('CLOSE_HOUR')).create();
}
/** Faqat TUGAGAN kunlar yopiladi: dars 13:00 gacha, shuning uchun 13:00 dan oldin yopilsa kecha (va zaxira uchun undan oldingi kun),
 *  13:00 dan keyin yopilsa bugun va kecha. Joriy kunning yozuvlari hech qachon vaqtidan oldin yopilmaydi. */
function dailyClose() {
  const h = +cfg('CLOSE_HOUR'), offs = h >= 13 ? [0, 1] : [1, 2];
  offs.forEach(k => {
    const ts = Date.now() - k * 86400000, day = fmt(ts, 'yyyy-MM-dd');
    if (+fmt(ts, 'u') > 5) return;                                   // shanba-yakshanba
    if (!attRows(3000).some(r => r.v[0] === day && r.v[5])) return;  // hech kim kelmagan kun (bayram) yopilmaydi
    closeDay(day, 'auto');
  });
}
function walk(f, out) {
  const fi = f.getFiles(); while (fi.hasNext()) { const x = fi.next(); out.push({ name: x.getName(), url: x.getUrl(), date: fmt(x.getLastUpdated(), 'yyyy-MM-dd HH:mm') }); }
  const fo = f.getFolders(); while (fo.hasNext()) walk(fo.next(), out);
}

/* ---------- Auth ---------- */
function audit(user, action, details) { try { sh('Audit').appendRow([nowStr(), user, action, JSON.stringify(details || {}).slice(0, 500)]); } catch (e) {} } // jurnal xatosi asosiy amalni to'xtatmasin
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
  photo: d => {
    if (!students().some(s => s.photoId === d.photoId)) throw new Error('Rasm topilmadi');
    return { data: Utilities.base64Encode(DriveApp.getFileById(d.photoId).getBlob().getBytes()) };
  },
  devStudents: () => ({ students: students().filter(s => s.active && s.photoId).map(s => ({ id: s.id, ism: s.ism, familiya: s.familiya, sinf: s.sinf, photoId: s.photoId })) }),
  devPhoto: d => ACT.photo(d),
  devApply: d => logEvent(d),
  devHeartbeat: d => { P.setProperty('HB', JSON.stringify({ t: Date.now(), fps: d.fps, cam: d.cam, host: d.host })); }, // tez (jadvalga yozilmaydi)
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
    const cc = CacheService.getScriptCache(), hit = cc.get('today');
    if (hit) return JSON.parse(hit);                      // 5 soniya kesh: bir nechta oyna/so'rov bir xil ishni takrorlamaydi
    const day = fmt(Date.now(), 'yyyy-MM-dd'), rows = attRows(900).filter(r => r.v[0] === day).map(r => r.v);
    const stl = students(), total = stl.filter(s => s.active).length, present = rows.filter(r => r[5]);
    let hbo = null; try { hbo = JSON.parse(P.getProperty('HB') || 'null'); } catch (e) {}
    const hbT = +cfg('HB_T'); if (hbT && (!hbo || hbT > hbo.t)) hbo = { t: hbT, fps: cfg('HB_FPS'), cam: cfg('HB_CAM'), host: cfg('HB_HOST') };
    const hb = hbo;
    let drive = true;
    if (cc.get('driveok') !== '1') { try { folder([]); cc.put('driveok', '1', 300); } catch (e) { drive = false; } }
    const ev = sh('Events'), n = ev.getLastRow(), map = {}; stl.forEach(s => map[s.id] = s);
    const all = n < 2 ? [] : ev.getRange(Math.max(2, n - 149), 1, Math.min(n - 1, 150), 5).getValues().reverse();
    // Jonli ro'yxatda faqat haqiqiy KIRDI/CHIQDI; noma'lum shaxslar va takrorlar bu yerga yozilmaydi (ular 'Tekshiruv' sahifasida)
    const recent = all.filter(r => r[3] === 'IN' || r[3] === 'OUT').slice(0, 30).map(r => {
      const s = map[String(r[2])] || {};
      return { ts: r[1], time: fmt(r[1], 'HH:mm:ss'), id: String(r[2]), ism: s.ism || '', familiya: s.familiya || '', sinf: s.sinf || '', type: r[3], note: r[4] };
    });
    const unk = all.filter(r => (r[3] === 'UNKNOWN' || r[3] === 'REVIEW') && fmt(r[1], 'yyyy-MM-dd') === day).length;
    const res = {
      day: day, total: total, present: present.length, late: rows.filter(r => +r[7] > 0).length,
      early: rows.filter(r => +r[8] > 0).length, notYet: Math.max(0, total - present.length), recent: recent, unk: unk,
      hb: hb ? { online: Date.now() - hb.t < 90000, fps: hb.fps, cam: hb.cam, host: hb.host } : { online: false },
      sheets: true, drive: drive, hours: cfg('START') + ' - ' + cfg('END')
    };
    cc.put('today', JSON.stringify(res), 5);
    return res;
  },
  history: d => {
    const rows = attRows(+d.last || 0).map(r => r.v).filter(v => (!d.from || v[0] >= d.from) && (!d.to || v[0] <= d.to) && (!d.sinf || v[4] === d.sinf));
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
  getSettings: () => ({ CLOSE_HOUR: cfg('CLOSE_HOUR'), SCHOOL: cfg('SCHOOL'), START: cfg('START'), END: cfg('END'), TZ: cfg('TZ'), MINGAP: cfg('MINGAP') }),
  setSettings: (d, u) => {
    if (d.CLOSE_HOUR !== undefined && d.CLOSE_HOUR !== '' && !(Number.isInteger(+d.CLOSE_HOUR) && +d.CLOSE_HOUR >= 0 && +d.CLOSE_HOUR <= 23)) throw new Error("Kunni yopish soati 0 dan 23 gacha butun son bo'lsin");
    ['SCHOOL', 'START', 'END', 'TZ', 'MINGAP', 'CLOSE_HOUR'].forEach(k => { if (d[k] !== undefined && d[k] !== '') setCfg(k, d[k]); });
    if (d.CLOSE_HOUR !== undefined && d.CLOSE_HOUR !== '') installTrigger();
    audit(u, 'setSettings', d);
  }
};

function route(d) {
  if (d.action === 'login') return login(d);
  let user = null;   // dev* amallar faqat qurilma kaliti bilan (Python, JSON siz rejim); qolganlari faqat admin tokeni bilan
  if (String(d.action).indexOf('dev') === 0) { if (d.key && d.key === P.getProperty('DEVICE_KEY')) user = 'device'; }
  else user = CacheService.getScriptCache().get('t:' + d.token);
  if (!user) throw new Error('AUTH');
  const f = ACT[d.action]; if (!f) throw new Error("Noma'lum amal");
  const out = Object.assign({ ok: true }, f(d, user) || {});
  if (['today', 'history', 'listStudents', 'getSettings', 'archive', 'review', 'photo', 'devStudents', 'devPhoto', 'devHeartbeat'].indexOf(d.action) < 0) CacheService.getScriptCache().remove('today');
  return out;
}
function doPost(e) {
  let out;
  try { out = route(JSON.parse(e.postData.contents)); }
  catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return ContentService.createTextOutput(JSON.stringify({ ok: true, service: 'SmartSchool Attendance', admin: !!P.getProperty('ADMIN_HASH'), device: !!P.getProperty('DEVICE_KEY'), v: 4 })).setMimeType(ContentService.MimeType.JSON); }


/* ---------- PAROLNI O'RNATISH (vaqtincha ishlatiladi) ----------
 * 1) Pastdagi PAROL qiymatiga o'z parolingizni yozing (12+ belgi) — faqat shu bitta joyga
 * 2) Funksiyalar ro'yxatidan "parol" ni tanlab Run bosing
 * 3) Ishlagach, PAROL ni 'BU_YERGA_PAROLNI_YOZING' holatiga qaytaring va saqlang (parol kodda qolmasin) */
function parol() {
  const LOGIN = 'adminjon';
  const PAROL = 'BU_YERGA_PAROLNI_YOZING';
  if (PAROL.indexOf('BU_YERGA') === 0 || PAROL.length < 8) throw new Error('Avval PAROL qiymatini yozing (kamida 8 belgi)');
  setAdmin(LOGIN, PAROL);
  Logger.log('Admin yangilandi: ' + LOGIN);
  try { login({ user: LOGIN, pass: PAROL }); Logger.log('Kirish sinovi: MUVAFFAQIYATLI (login va parol shu loyihada to\'g\'ri saqlandi)'); }
  catch (e) { Logger.log('Kirish sinovi: XATO — ' + e.message); }
  Logger.log("Eslatma: haqiqiy web app havolasi Deploy > Manage deployments ichida (/exec bilan tugaydi). Editor ko'rsatadigan /dev havola sinov uchun, uni saytga yozmang.");
}

/* Tekshiruv: Run qiling va Execution log'ni ko'ring (parolning o'zi ko'rsatilmaydi) */
function adminTekshir() {
  Logger.log('Login: ' + P.getProperty('ADMIN_USER') + ' | salt: ' + (P.getProperty('SALT') ? 'bor' : "YO'Q") + ' | parol saqlangan: ' + (P.getProperty('ADMIN_HASH') ? 'ha' : "YO'Q") + ' | qurilma kaliti: ' + (P.getProperty('DEVICE_KEY') ? 'bor' : "YO'Q (setup() ni ishga tushiring)"));
  try {
    const ss = SS();
    Logger.log('Jadval: ' + ss.getName() + ' | varaqlar: ' + Object.keys(HEAD).map(n => n + (ss.getSheetByName(n) ? ' bor' : " YO'Q (setup() ni ishga tushiring)")).join(', '));
  } catch (e) { Logger.log("Jadval ochilmadi — SHEET_ID yoki ruxsatni tekshiring: " + e.message); }
  try { Logger.log('Drive papka: ' + ROOT().getName()); } catch (e) { Logger.log("Drive papka ochilmadi — FOLDER_ID yoki ruxsatni tekshiring: " + e.message); }
}

/* Python (JSON siz rejim) uchun kalitni logda ko'rsatadi */
function kalit() { Logger.log('DEVICE_KEY: ' + P.getProperty('DEVICE_KEY')); }
