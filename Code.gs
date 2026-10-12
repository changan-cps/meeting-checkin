/**
 * Meeting Check-in — ส่วนเก็บข้อมูล (Google Apps Script)
 * วางโค้ดนี้ใน Extensions > Apps Script ของ Google Sheet แล้ว Deploy เป็น Web app
 * (Execute as: Me, Who has access: Anyone) ดูขั้นตอนใน README.md
 *
 * ชีตที่ระบบสร้างให้อัตโนมัติ:
 *   Meetings     — ตารางการประชุม
 *   Departments  — รายชื่อส่วนงาน
 *   Members      — รายชื่อพนักงานแยกตามส่วนงาน (แก้ในชีตโดยตรงได้)
 *   Checkins     — บันทึกการเช็คอินทุกการประชุม
 */

/* ===================== ตั้งค่า ===================== */
const ADMIN_PIN = '1234';        // รหัสผู้จัด/ผู้ดูแล — เปลี่ยนก่อนใช้งานจริง
const QR_SECONDS = 60;           // QR เปลี่ยนทุกกี่วินาที
const QR_GRACE_SLOTS = 1;        // ยอมรับ QR รอบก่อนหน้ากี่รอบ (1 = QR แต่ละอันใช้ได้นานสุด 2 นาที)
const TICKET_MINUTES = 10;       // หลังสแกนแล้ว ต้องกดเช็คอินภายในกี่นาที
const DEFAULT_DEPARTMENTS = ['CPS team leader', 'Deputy Team Leader of CPS Construction Project', 'CPS', 'SOS/EOS', 'QOS', 'DOS', 'MOS', 'COS', 'POS', 'Body Workshop', 'Paint Workshop', 'General Assembly Workshop', 'Powertrain Workshop', 'KD'];

const TZ = 'Asia/Bangkok';
const HEAD = {
  Meetings: ['รหัสประชุม', 'ชื่อประชุม', 'วันที่', 'เปิดสแกน', 'สายหลัง', 'ปิดสแกน', 'ห้อง', 'ส่วนงาน', 'คีย์ผู้จัด', 'สร้างเมื่อ', 'ผู้จัด (Host)'],
  Checkins: ['เวลาเช็คอิน', 'รหัสประชุม', 'ชื่อประชุม', 'ชื่อ', 'ส่วนงาน', 'รหัสการสแกน'],
  Departments: ['ส่วนงาน'],
  Members: ['ส่วนงาน', 'ชื่อ', 'ที่มา', 'วันที่เพิ่ม', 'รหัสพนักงาน', 'ตำแหน่ง']
};
const SRC_AUTO = 'เพิ่มตอนเช็คอิน';
const SRC_ADMIN = 'ผู้ดูแล';

/* ---------- helpers ---------- */
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === 'Meetings') sh.getRange('A:K').setNumberFormat('@'); // เก็บวันที่/เวลาเป็นข้อความ
    sh.appendRow(HEAD[name]);
    sh.setFrozenRows(1);
    if (name === 'Departments') {
      sh.getRange(2, 1, DEFAULT_DEPARTMENTS.length, 1).setValues(DEFAULT_DEPARTMENTS.map(function (d) { return [d]; }));
    }
  } else if (sh.getLastColumn() < HEAD[name].length) {
    // ชีตเดิมจากเวอร์ชันก่อน -> เติมหัวคอลัมน์ใหม่ (เช่น ตำแหน่ง) ให้อัตโนมัติ
    sh.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]);
  }
  return sh;
}
function rows_(sh, ncol) {
  const last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, ncol).getValues();
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
// ตัดช่องว่าง/ความยาว และตัดอักขระนำหน้าที่ชีตจะตีความเป็นสูตร
function clean_(v, max) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().replace(/^[=+\-@]+/, '').trim().slice(0, max);
}
function key_(s) { return String(s).replace(/\s+/g, ' ').trim().toLowerCase(); }
function err_(code, msg) { return { ok: false, code: code, error: msg }; }
function hmac_(secret, msg) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(msg, secret)).replace(/=+$/, '');
}
function rand_(n) {
  let s = '';
  while (s.length < n) s += Utilities.getUuid().replace(/-/g, '');
  return s.slice(0, n);
}
function cellText_(v, fmt) { return v instanceof Date ? Utilities.formatDate(v, TZ, fmt) : String(v == null ? '' : v); }

/* ---------- ส่วนงาน / รายชื่อ ---------- */
function getData_() {
  const departments = rows_(sheet_('Departments'), 1)
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (d, i, a) { return d && a.map(key_).indexOf(key_(d)) === i; });
  const members = rows_(sheet_('Members'), 6)
    .filter(function (r) { return String(r[0]).trim() && String(r[1]).trim(); })
    .map(function (r) { return { dept: String(r[0]).trim(), name: String(r[1]).trim(), src: String(r[2] || ''), id: String(r[4] || '').trim(), pos: String(r[5] || '').trim() }; });
  return { departments: departments, members: members };
}
function findDept_(departments, dept) {
  const k = key_(dept);
  for (let i = 0; i < departments.length; i++) if (key_(departments[i]) === k) return departments[i];
  return null;
}
function invited_(mt, db) {
  if (!mt.g.length) return db.departments.slice();
  const want = mt.g.map(key_);
  return db.departments.filter(function (d) { return want.indexOf(key_(d)) >= 0; });
}

/* ---------- การประชุม ---------- */
function readMeetings_() {
  return rows_(sheet_('Meetings'), 11).filter(function (r) { return r[0]; }).map(function (r) {
    return {
      id: String(r[0]), t: String(r[1]), d: cellText_(r[2], 'yyyy-MM-dd'),
      s: cellText_(r[3], 'HH:mm'), l: cellText_(r[4], 'HH:mm'), e: cellText_(r[5], 'HH:mm'), r: String(r[6] || ''),
      g: String(r[7] || '').split('|').map(function (x) { return x.trim(); }).filter(Boolean), k: String(r[8]), h: String(r[10] || '')
    };
  });
}
// ช่วงเวลาที่สแกนได้ (เวลาไทย) — คืน null ถ้าอยู่ในช่วง หรือข้อความถ้าไม่อยู่ในช่วง
function at_(d, hm) { return new Date(d + 'T' + hm + ':00+07:00').getTime(); }
function windowError_(mt, now) {
  if (!mt.d) return null;
  const open = mt.s ? at_(mt.d, mt.s) : at_(mt.d, '00:00');
  const close = mt.e ? at_(mt.d, mt.e) : at_(mt.d, '23:59') + 59000;
  if (now < open) return 'ยังไม่เปิดให้เช็คอิน เปิดสแกนวันที่ ' + mt.d + ' เวลา ' + (mt.s || '00:00') + ' น.';
  if (now > close) return 'ปิดรับเช็คอินแล้ว (รับถึง ' + (mt.e || '23:59') + ' น. ของวันที่ ' + mt.d + ')';
  return null;
}
function findMeeting_(id) {
  return readMeetings_().filter(function (x) { return x.id === String(id || ''); })[0] || null;
}
function pub_(mt) { return { id: mt.id, t: mt.t, d: mt.d, s: mt.s, l: mt.l, e: mt.e, r: mt.r, g: mt.g, h: mt.h }; }

function readCheckins_(meetingId) {
  return rows_(sheet_('Checkins'), 6)
    .filter(function (r) { return String(r[1]) === meetingId; })
    .map(function (r) { return { time: new Date(r[0]).toISOString(), name: String(r[3]), dept: String(r[4]), tid: String(r[5] || '') }; });
}
function deleteWhere_(sh, ncol, pred) {
  const rows = rows_(sh, ncol);
  for (let i = rows.length - 1; i >= 0; i--) if (pred(rows[i])) sh.deleteRow(i + 2);
}

/* ---------- QR หมุนเวียน + ตั๋วใช้ครั้งเดียว ---------- */
function slot_(now) { return Math.floor(now / (QR_SECONDS * 1000)); }
function qrToken_(mt, slot) { return hmac_(mt.k, 'Q|' + mt.id + '|' + slot).slice(0, 12); }
function makeTicket_(mt, now) {
  const tid = rand_(10), exp = now + TICKET_MINUTES * 60000;
  return tid + '.' + exp + '.' + hmac_(mt.k, 'T|' + mt.id + '|' + tid + '|' + exp).slice(0, 22);
}
function readTicket_(mt, ticket, now) {
  const p = String(ticket || '').split('.');
  if (p.length !== 3) return null;
  if (hmac_(mt.k, 'T|' + mt.id + '|' + p[0] + '|' + p[1]).slice(0, 22) !== p[2]) return null;
  if (Number(p[1]) < now) return null;
  return p[0];
}

/* ---------- GET (ใช้ทดสอบว่า URL ถูกต้อง) ---------- */
function doGet() {
  return json_({ ok: true, app: 'meeting-checkin' });
}

/* ---------- POST ---------- */
const WRITE_ACTIONS = ['checkin', 'admin', 'createMeeting', 'deleteMeeting'];
function doPost(e) {
  let lock = null;
  try {
    const data = JSON.parse(e.postData.contents);
    if (WRITE_ACTIONS.indexOf(data.action) >= 0) { lock = LockService.getScriptLock(); lock.waitLock(20000); }
    const fn = { host: host_, claim: claim_, checkin: checkin_, admin: admin_, meetings: meetings_, createMeeting: createMeeting_, deleteMeeting: deleteMeeting_ }[data.action];
    return json_(fn ? fn(data) : err_('action', 'unknown action'));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    if (lock) try { lock.releaseLock(); } catch (x) {}
  }
}
function pinOk_(data) { return String(data.pin || '') === String(ADMIN_PIN); }

// ตารางการประชุม (ต้องใช้ PIN)
function meetings_(data) {
  if (!pinOk_(data)) return err_('pin', 'PIN ไม่ถูกต้อง');
  return { ok: true, meetings: withCounts_(readMeetings_()) };
}
// จำนวนผู้เช็คอิน / ผู้ที่ต้องเข้า / มาสาย ของแต่ละรอบการประชุม
function withCounts_(list) {
  const db = getData_();
  const rows = rows_(sheet_('Checkins'), 5);
  return list.map(function (mt) {
    const recs = rows.filter(function (r) { return String(r[1]) === mt.id; });
    const dk = invited_(mt, db).map(key_);
    const late = mt.l ? recs.filter(function (r) { return new Date(r[0]).getTime() > at_(mt.d, mt.l); }).length : 0;
    return Object.assign({}, mt, {
      n: recs.length, late: late,
      total: db.members.filter(function (x) { return dk.indexOf(key_(x.dept)) >= 0; }).length
    });
  });
}
function createMeeting_(data) {
  if (!pinOk_(data)) return err_('pin', 'PIN ไม่ถูกต้อง');
  const t = clean_(data.t, 80);
  const d = /^\d{4}-\d{2}-\d{2}$/.test(data.d) ? data.d : '';
  const hm = function (v) { return /^\d{2}:\d{2}$/.test(v) ? v : ''; };
  const s = hm(data.s), l = hm(data.l), e = hm(data.e);
  if (!t || !d) return err_('input', 'กรุณากรอกชื่อและวันที่');
  if (s && e && e <= s) return err_('input', 'เวลาปิดสแกนต้องอยู่หลังเวลาเปิดสแกน');
  const db = getData_();
  const g = (data.g || []).map(function (x) { return findDept_(db.departments, x); }).filter(Boolean);
  const mt = { id: 'M' + rand_(9), t: t, d: d, s: s, l: l, e: e, r: clean_(data.r, 60), g: g, k: rand_(32), h: clean_(data.h, 60) };
  sheet_('Meetings').appendRow([mt.id, mt.t, mt.d, mt.s, mt.l, mt.e, mt.r, mt.g.join('|'), mt.k, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), mt.h]);
  return { ok: true, meeting: mt };
}
function deleteMeeting_(data) {
  if (!pinOk_(data)) return err_('pin', 'PIN ไม่ถูกต้อง');
  deleteWhere_(sheet_('Meetings'), 1, function (r) { return String(r[0]) === String(data.m); });
  return { ok: true, meetings: withCounts_(readMeetings_()) };
}

// หน้าจอผู้จัด: QR ปัจจุบัน + รายชื่อ (ต้องใช้คีย์ผู้จัดของการประชุม)
function host_(data) {
  const mt = findMeeting_(data.m);
  if (!mt || String(data.k || '') !== mt.k) return err_('key', 'ลิงก์ผู้จัดไม่ถูกต้อง หรือการประชุมถูกลบแล้ว');
  const now = Date.now(), slot = slot_(now);
  const db = getData_();
  return {
    ok: true, meeting: pub_(mt), departments: invited_(mt, db), members: db.members,
    records: readCheckins_(mt.id).map(function (r) { return { time: r.time, name: r.name, dept: r.dept }; }),
    token: qrToken_(mt, slot), msLeft: (slot + 1) * QR_SECONDS * 1000 - now, qrSeconds: QR_SECONDS,
    closed: windowError_(mt, now), serverTime: now
  };
}

// ผู้เข้าร่วมสแกน QR -> ตรวจ QR แล้วออกตั๋วใช้ครั้งเดียว
function claim_(data) {
  const mt = findMeeting_(data.m);
  if (!mt) return err_('meeting', 'ไม่พบการประชุมนี้');
  const now = Date.now();
  const closed = windowError_(mt, now);
  if (closed) return Object.assign(err_('closed', closed), { meeting: pub_(mt) });
  let ticket = '';
  if (data.ticket && readTicket_(mt, data.ticket, now)) {
    ticket = data.ticket;
  } else if (data.q) {
    const cur = slot_(now);
    for (let i = 0; i <= QR_GRACE_SLOTS; i++) if (qrToken_(mt, cur - i) === String(data.q)) { ticket = makeTicket_(mt, now); break; }
  }
  if (!ticket) return err_('expired', 'QR หมดอายุ');
  const db = getData_();
  const depts = invited_(mt, db), dk = depts.map(key_);
  return {
    ok: true, meeting: pub_(mt), ticket: ticket, departments: depts,
    members: db.members.filter(function (x) { return dk.indexOf(key_(x.dept)) >= 0; }).map(function (x) { return { dept: x.dept, name: x.name, pos: x.pos }; }),
    checked: readCheckins_(mt.id).map(function (r) { return key_(r.dept) + '|' + key_(r.name); })
  };
}

// {action:'checkin', m, ticket, dept, name}
function checkin_(data) {
  const mt = findMeeting_(data.m);
  if (!mt) return err_('meeting', 'ไม่พบการประชุมนี้');
  const now = Date.now();
  const closed = windowError_(mt, now);
  if (closed) return err_('closed', closed);
  const tid = readTicket_(mt, data.ticket, now);
  if (!tid) return err_('expired', 'QR หมดอายุ กรุณาสแกนใหม่');

  let name = clean_(data.name, 80);
  if (!name) return err_('input', 'กรุณาระบุชื่อ');
  const db = getData_();
  const dept = findDept_(invited_(mt, db), clean_(data.dept, 60));
  if (!dept) return err_('input', 'ไม่พบส่วนงานนี้');

  const member = db.members.filter(function (x) { return key_(x.dept) === key_(dept) && key_(x.name) === key_(name); })[0];
  if (member) name = member.name;

  const recs = readCheckins_(mt.id);
  const existing = recs.filter(function (r) { return key_(r.dept) === key_(dept) && key_(r.name) === key_(name); })[0];
  if (existing) return { ok: true, duplicate: true, added: false, record: { time: existing.time, name: existing.name, dept: existing.dept } };
  if (recs.some(function (r) { return r.tid === tid; })) return err_('used', 'QR นี้ใช้เช็คอินไปแล้ว กรุณาสแกนใหม่');

  const at = new Date(now); // ใช้เวลาของเซิร์ฟเวอร์ ไม่ใช้เวลาจากมือถือ
  let added = false;
  if (!member) { // ชื่อใหม่จาก "อื่นๆ" -> เพิ่มเข้ารายชื่อส่วนงานอัตโนมัติ
    sheet_('Members').appendRow([dept, name, SRC_AUTO, at]);
    added = true;
  }
  sheet_('Checkins').appendRow([at, mt.id, mt.t, name, dept, tid]);
  return { ok: true, duplicate: false, added: added, record: { time: at.toISOString(), name: name, dept: dept } };
}

// จัดการส่วนงาน/รายชื่อ {action:'admin', pin, op, ...}
function admin_(data) {
  if (!pinOk_(data)) return err_('pin', 'PIN ไม่ถูกต้อง');
  const db = getData_();
  const now = new Date();

  switch (data.op) {
    case 'checkPin':
      break;

    case 'addDept': {
      const d = clean_(data.dept, 60);
      if (!d) return err_('input', 'กรุณากรอกชื่อส่วนงาน');
      if (!findDept_(db.departments, d)) sheet_('Departments').appendRow([d]);
      break;
    }

    case 'delDept': {
      const k = key_(data.dept);
      deleteWhere_(sheet_('Departments'), 1, function (r) { return key_(r[0]) === k; });
      deleteWhere_(sheet_('Members'), 1, function (r) { return key_(r[0]) === k; });
      break;
    }

    case 'addNames': {
      const dept = findDept_(db.departments, clean_(data.dept, 60));
      if (!dept) return err_('input', 'ไม่พบส่วนงานนี้');
      const have = {};
      db.members.forEach(function (x) { if (key_(x.dept) === key_(dept)) have[key_(x.name)] = 1; });
      const rows = [];
      const sh = sheet_('Members');
      const existing = rows_(sh, 2);
      (data.names || []).forEach(function (x) {
        const n = clean_(typeof x === 'object' ? x.name : x, 80);
        const id = clean_(typeof x === 'object' ? x.id : '', 30);
        const pos = clean_(typeof x === 'object' ? x.pos : '', 60);
        if (!n) return;
        if (have[key_(n)]) { // มีชื่อแล้ว -> อัปเดตรหัสพนักงาน/ตำแหน่งถ้าใส่มา
          existing.forEach(function (r, i) {
            if (key_(r[0]) !== key_(dept) || key_(r[1]) !== key_(n)) return;
            if (id) sh.getRange(i + 2, 5).setValue(id);
            if (pos) sh.getRange(i + 2, 6).setValue(pos);
          });
          return;
        }
        have[key_(n)] = 1; rows.push([dept, n, SRC_ADMIN, now, id, pos]);
      });
      if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
      break;
    }

    case 'delName': {
      const kd = key_(data.dept), kn = key_(data.name);
      deleteWhere_(sheet_('Members'), 2, function (r) { return key_(r[0]) === kd && key_(r[1]) === kn; });
      break;
    }

    case 'approveName': { // ยืนยันชื่อที่เพิ่มตอนเช็คอินให้เป็นรายชื่อปกติ
      const kd = key_(data.dept), kn = key_(data.name);
      const sh = sheet_('Members');
      rows_(sh, 3).forEach(function (r, i) {
        if (key_(r[0]) === kd && key_(r[1]) === kn) sh.getRange(i + 2, 3).setValue(SRC_ADMIN);
      });
      break;
    }

    default:
      return err_('op', 'unknown op');
  }
  return Object.assign({ ok: true }, getData_());
}
