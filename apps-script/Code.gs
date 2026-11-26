/**
 * ระบบแบ่งกลุ่มนักศึกษา ป.จว. รุ่นที่ 1 — Google Apps Script backend (มีระบบเข้าสู่ระบบตามบทบาท)
 *
 * บทบาทและรหัสเข้าใช้งาน:
 *   นิสิต          รหัสนิสิต 10 หลัก (เช่น 69xxxxxxxx) — ต้องตรงกับคอลัมน์ "รหัสนิสิต" ในแผ่นงานรายชื่อ
 *   อาจารย์ผู้สอน  ตัวเลข 4 หลัก (เช่น 6901)          — ต้องอยู่ในแผ่นงาน "ผู้ใช้ (อาจารย์-TA)"
 *   TA ผู้ช่วย      TA + ตัวเลข 4 หลัก (เช่น TA6901)  — ต้องอยู่ในแผ่นงาน "ผู้ใช้ (อาจารย์-TA)"
 *   ผู้ดูแลระบบ    รหัสผ่านผู้ดูแล (ตั้งครั้งแรกจาก INITIAL_ADMIN_PASSWORD ใน Seed.gs)
 *
 * สิทธิ์แต่ละบทบาทปรับได้ที่ค่าคงที่ PERMS ด้านล่าง
 * ทุกคำขอจากหน้าเว็บเป็น POST (Content-Type: text/plain) {action, token, ...}
 * GitHub Actions ใช้ API key: GET ?action=summary&key=... และ POST {key, action:"snapshot"}
 */

const SHEET_NAME = 'รายชื่อตามกลุ่ม';
const SUMMARY_NAME = 'สรุปแต่ละกลุ่ม';
const USERS_NAME = 'ผู้ใช้ (อาจารย์-TA)';
const CONFIG_NAME = 'ตั้งค่า';
const LOG_NAME = 'บันทึกการแก้ไข';
const BACKUP_PREFIX = 'สำรอง ';
const KEEP_BACKUPS = 14;
const SESSION_SECONDS = 6 * 60 * 60; // อยู่ในระบบได้ 6 ชั่วโมง (สูงสุดที่ CacheService รองรับ)

const KEYS = ['group', 'id', 'studentId', 'name', 'status', 'career',
  'practice', 'consult', 'occupation', 'updatedAt'];
const OPTION_KEYS = ['status', 'career', 'practice', 'consult'];
const EDITABLE = KEYS.filter(k => k !== 'updatedAt');

const ROLE_NAMES = { student: 'นิสิต', teacher: 'อาจารย์ผู้สอน', ta: 'TA ผู้ช่วย', admin: 'ผู้ดูแลระบบ' };

/** สิทธิ์ของแต่ละบทบาท */
const PERMS = {
  // นิสิต: ดูได้เฉพาะกลุ่มของตัวเอง แก้ไขไม่ได้
  student: { viewAll: false, move: false, edit: [], add: false, del: false, admin: false },
  // อาจารย์: ดูทั้งหมด ย้ายกลุ่ม และแก้ข้อมูลประกอบ (ไม่แก้ชื่อ/รหัส ไม่ลบ)
  teacher: { viewAll: true, move: true, edit: ['group', 'status', 'career', 'practice', 'consult', 'occupation'],
    add: false, del: false, admin: false },
  // TA: ดูทั้งหมด และย้ายกลุ่มได้
  ta: { viewAll: true, move: true, edit: [], add: false, del: false, admin: false },
  // ผู้ดูแล: เพิ่ม ลบ แก้ไขได้ทุกช่อง ทุกข้อความ ทุกส่วน
  admin: { viewAll: true, move: true, edit: EDITABLE, add: true, del: true, admin: true },
};

const DEFAULT_CONFIG = {
  title: 'แบ่งกลุ่มนักศึกษา ป.จว. รุ่นที่ 1',
  subtitle: 'หลักสูตรประกาศนียบัตรจิตวิทยาเชิงพุทธ',
  groups: ['กลุ่ม 1', 'กลุ่ม 2', 'กลุ่ม 3', 'กลุ่ม 4', 'กลุ่ม 5', 'กลุ่ม 6', 'กลุ่ม 7', 'กลุ่ม 8'],
  options: {
    status: ['พระภิกษุ', 'ฆราวาส'],
    career: ['พระภิกษุ', 'พนักงานเอกชน/ผู้บริหาร', 'ธุรกิจ/ผู้ประกอบการ', 'ราชการ/การศึกษา',
      'จิตวิทยา/โค้ช/ที่ปรึกษา', 'สุขภาพ/การแพทย์', 'อิสระ/เกษียณ/อื่น ๆ'],
    practice: ['สม่ำเสมอ', 'ครั้งคราว', 'ยังไม่เคย'],
    consult: ['ทำงานด้านนี้', 'เคยอบรม', 'ไม่มี'],
  },
  labels: {
    group: 'กลุ่ม', id: 'เลขที่', studentId: 'รหัสนิสิต', name: 'ชื่อ-นามสกุล', status: 'สถานะ',
    career: 'กลุ่มอาชีพ', practice: 'ประสบการณ์ปฏิบัติ', consult: 'ประสบการณ์การปรึกษา',
    occupation: 'อาชีพ (ตามใบสมัคร)', updatedAt: 'แก้ไขล่าสุด',
  },
};

/* ═════════════════════════ ติดตั้ง ═════════════════════════ */

function setup() {
  const ss = ss_();
  ss.setSpreadsheetLocale('th_TH');
  ss.setSpreadsheetTimeZone('Asia/Bangkok');

  // แผ่นงานตั้งค่า
  if (!ss.getSheetByName(CONFIG_NAME)) saveConfigSheet_(clone_(DEFAULT_CONFIG));
  const cfg = getConfig_();

  // แผ่นงานรายชื่อ
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
  // ย้ายจากเวอร์ชันแรก (ยังไม่มีคอลัมน์รหัสนิสิต): แทรกคอลัมน์ C
  if (sh.getLastColumn() === KEYS.length - 1 && String(sh.getRange(1, 3).getValue()) === 'ชื่อ-นามสกุล') {
    sh.insertColumnAfter(2);
  }
  writeHeaders_(cfg);
  sh.setFrozenRows(1);

  if (sh.getLastRow() < 2 && typeof SEED_DATA !== 'undefined') {
    const now = new Date();
    const rows = SEED_DATA.map(r => [r[0], r[1], ''].concat(r.slice(2), [now]));
    sh.getRange(2, col_('id'), rows.length, 2).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, KEYS.length).setValues(rows);
  }
  sh.getRange(2, col_('id'), sh.getMaxRows() - 1, 2).setNumberFormat('@'); // เลขที่ + รหัสนิสิต เป็นข้อความ
  sh.getRange(2, col_('updatedAt'), sh.getMaxRows() - 1).setNumberFormat('dd/MM/yyyy HH:mm');
  applyValidation_(cfg);
  sh.autoResizeColumns(1, KEYS.length);

  // แผ่นงานผู้ใช้
  let us = ss.getSheetByName(USERS_NAME);
  if (!us) {
    us = ss.insertSheet(USERS_NAME);
    us.getRange(1, 1, 1, 4).setValues([['รหัส', 'ชื่อ-นามสกุล', 'หมายเหตุ', 'บทบาท (คำนวณอัตโนมัติ)']])
      .setFontWeight('bold').setBackground('#1f3b57').setFontColor('#ffffff');
    us.setFrozenRows(1);
    us.getRange('A:A').setNumberFormat('@');
    if (typeof SEED_USERS !== 'undefined' && SEED_USERS.length) {
      us.getRange(2, 1, SEED_USERS.length, 3).setValues(SEED_USERS);
    }
    us.getRange('D2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="",,IF(REGEXMATCH(UPPER(A2:A),"^TA[0-9]{4}$"),"TA ผู้ช่วย",' +
      'IF(REGEXMATCH(A2:A&"","^[0-9]{4}$"),"อาจารย์ผู้สอน","⚠️ รูปแบบรหัสไม่ถูกต้อง"))))');
    us.setColumnWidths(1, 4, 180);
  }

  buildSummarySheet_(cfg);

  let log = ss.getSheetByName(LOG_NAME);
  if (!log) {
    log = ss.insertSheet(LOG_NAME);
    log.getRange(1, 1, 1, 5).setValues([['เวลา', 'ผู้ทำรายการ', 'การกระทำ', 'เลขที่', 'รายละเอียด']]).setFontWeight('bold');
    log.setFrozenRows(1);
  }

  // รหัสผ่านผู้ดูแล + API key (เก็บใน Script Properties ไม่อยู่ในโค้ด)
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PASSWORD')) {
    if (typeof INITIAL_ADMIN_PASSWORD === 'undefined' || !INITIAL_ADMIN_PASSWORD)
      throw new Error('ยังไม่ได้กำหนด INITIAL_ADMIN_PASSWORD ใน Seed.gs');
    props.setProperty('ADMIN_PASSWORD', INITIAL_ADMIN_PASSWORD);
  }
  if (!props.getProperty('API_KEY')) props.setProperty('API_KEY', Utilities.getUuid().replace(/-/g, ''));
  Logger.log('ติดตั้งเสร็จแล้ว');
  Logger.log('API_KEY (สำหรับ GitHub Secrets ชื่อ GAS_API_KEY) = ' + props.getProperty('API_KEY'));
}

/** เปลี่ยนรหัสผ่านผู้ดูแล: แก้ค่าในบรรทัดล่างแล้ว Run (แล้วลบค่าออกจากโค้ด) */
function setAdminPassword() {
  const NEW_PASSWORD = '';
  if (NEW_PASSWORD.length < 8) throw new Error('ใส่รหัสผ่านใหม่อย่างน้อย 8 ตัวอักษรใน NEW_PASSWORD ก่อน Run');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', NEW_PASSWORD);
  Logger.log('เปลี่ยนรหัสผ่านผู้ดูแลแล้ว');
}

function rotateApiKey() {
  const key = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('API_KEY', key);
  Logger.log('API_KEY ใหม่ = ' + key);
}

/* ═════════════════════════ Web App ═════════════════════════ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'ping') return json_({ ok: true, time: new Date().toISOString() });
    checkKey_(p.key);
    if (p.action === 'summary') return json_({ ok: true, summary: summarize_(readAll_(), getConfig_()) });
    throw new Error('ไม่รู้จัก action: ' + p.action);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'รูปแบบคำขอไม่ถูกต้อง' });
  }
  try {
    if (body.action === 'login') return json_(login_(body.code));

    if (body.key) { // สำหรับ GitHub Actions
      checkKey_(body.key);
      if (body.action === 'snapshot') return json_({ ok: true, backup: withLock_(snapshot_) });
      if (body.action === 'summary') return json_({ ok: true, summary: summarize_(readAll_(), getConfig_()) });
      throw new Error('ไม่รู้จัก action: ' + body.action);
    }

    const user = auth_(body.token);
    const p = PERMS[user.role];
    const need = ok => { if (!ok) throw new Error('คุณไม่มีสิทธิ์ทำรายการนี้'); };
    const done = () => json_(bootstrap_(auth_(body.token)));

    switch (body.action) {
      case 'list':
        return json_(bootstrap_(user));
      case 'logout':
        CacheService.getScriptCache().remove('sess:' + body.token);
        return json_({ ok: true });
      case 'move':
        need(p.move);
        withLock_(() => updateStudent_(body.id, { group: body.group }, user, ['group'], 'ย้ายกลุ่ม'));
        return done();
      case 'update':
        need(p.edit.length > 0);
        withLock_(() => updateStudent_(body.id, body.data || {}, user, p.edit, 'แก้ไข'));
        return done();
      case 'add':
        need(p.add);
        withLock_(() => addStudent_(body.data || {}, user));
        return done();
      case 'delete':
        need(p.del);
        withLock_(() => deleteStudent_(body.id, user));
        return done();
      case 'addGroup':
        need(p.admin);
        withLock_(() => addGroup_(body.name, user));
        return done();
      case 'renameGroup':
        need(p.admin);
        withLock_(() => renameGroup_(body.from, body.to, user));
        return done();
      case 'deleteGroup':
        need(p.admin);
        withLock_(() => deleteGroup_(body.name, user));
        return done();
      case 'moveGroup':
        need(p.admin);
        withLock_(() => moveGroup_(body.name, body.direction, user));
        return done();
      case 'saveConfig':
        need(p.admin);
        withLock_(() => saveConfig_(body.config || {}, user));
        return done();
      case 'saveUsers':
        need(p.admin);
        withLock_(() => saveUsers_(body.users || [], user));
        return done();
      case 'snapshot':
        need(p.admin);
        return json_({ ok: true, backup: withLock_(snapshot_) });
      default:
        throw new Error('ไม่รู้จัก action: ' + body.action);
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err), auth: !!err.auth });
  }
}

/* ═════════════════════════ เข้าสู่ระบบ ═════════════════════════ */

function login_(rawCode) {
  const code = String(rawCode || '').trim();
  if (!code) throw new Error('กรุณาใส่รหัส');
  const cache = CacheService.getScriptCache();
  const failKey = 'fail:' + hash_(code.toUpperCase());
  if (Number(cache.get('fail:all') || 0) >= 50) throw new Error('มีการใส่รหัสผิดจำนวนมาก ระบบปิดชั่วคราว 10 นาที');
  if (Number(cache.get(failKey) || 0) >= 5) throw new Error('ใส่รหัสนี้ผิดเกิน 5 ครั้ง กรุณารอ 15 นาที');

  const adminPw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  let user = null;
  if (adminPw && code === adminPw) {
    user = { role: 'admin', code: 'admin', name: ROLE_NAMES.admin };
  } else if (/^\d{10}$/.test(code)) {
    const s = readAll_().find(x => x.studentId === code);
    if (s) user = { role: 'student', code: code, name: s.name };
  } else if (/^\d{4}$/.test(code) || /^TA\d{4}$/i.test(code)) {
    const u = users_().find(x => x.code === code.toUpperCase());
    if (u) user = { role: u.role, code: u.code, name: u.name };
  }

  if (!user) {
    cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), 900);
    cache.put('fail:all', String(Number(cache.get('fail:all') || 0) + 1), 600);
    let hint = 'รหัสไม่ถูกต้อง';
    if (/^\d{10}$/.test(code)) hint = 'ไม่พบรหัสนิสิตนี้ในระบบ';
    else if (/^\d{4}$/.test(code)) hint = 'ไม่พบรหัสอาจารย์นี้ในระบบ';
    else if (/^TA\d{4}$/i.test(code)) hint = 'ไม่พบรหัส TA นี้ในระบบ';
    throw new Error(hint);
  }
  cache.remove(failKey);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache.put('sess:' + token, JSON.stringify(user), SESSION_SECONDS);
  log_(user, 'เข้าสู่ระบบ', '', '');
  const out = bootstrap_(auth_(token));
  out.token = token;
  return out;
}

function auth_(token) {
  const fail = msg => { const e = new Error(msg || 'กรุณาเข้าสู่ระบบ'); e.auth = true; return e; };
  if (!token) throw fail();
  const raw = CacheService.getScriptCache().get('sess:' + token);
  if (!raw) throw fail('หมดเวลาการใช้งาน กรุณาเข้าสู่ระบบใหม่');
  const user = JSON.parse(raw);
  // ตรวจทุกครั้งว่ายังมีสิทธิ์อยู่ (เผื่อถูกลบหรือเปลี่ยนรหัส)
  if (user.role === 'student') {
    const me = readAll_().find(s => s.studentId === user.code);
    if (!me) throw fail('ไม่พบรหัสนิสิตของคุณในระบบแล้ว');
    user.me = me;
    user.name = me.name;
  } else if (user.role === 'teacher' || user.role === 'ta') {
    const u = users_().find(x => x.code === user.code);
    if (!u) throw fail('รหัสของคุณถูกยกเลิกแล้ว');
    user.name = u.name;
  }
  return user;
}

function bootstrap_(user) {
  const cfg = getConfig_();
  const p = PERMS[user.role];
  let students = readAll_();
  let me = null;
  if (!p.viewAll) {
    me = user.me;
    students = students.filter(s => s.group === me.group)
      .map(s => ({ group: s.group, id: s.id, name: s.name, status: s.status, career: s.career }));
  }
  return {
    ok: true,
    user: { role: user.role, roleName: ROLE_NAMES[user.role], name: user.name, code: user.role === 'admin' ? '' : user.code },
    perms: p,
    config: cfg,
    students: students,
    me: me,
    users: p.admin ? users_() : undefined,
  };
}

/* ═════════════════════════ นักศึกษา ═════════════════════════ */

function readAll_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, KEYS.length).getValues()
    .map((row, i) => toObj_(row, i + 2))
    .filter(s => s.id || s.name);
}

function toObj_(row, rowNum) {
  const o = { row: rowNum };
  KEYS.forEach((k, i) => {
    const v = row[i];
    o[k] = v instanceof Date ? v.toISOString() : String(v == null ? '' : v).trim();
  });
  return o;
}

function findRow_(id) {
  if (!id) throw new Error('ต้องระบุเลขที่');
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, col_('id'), n, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === String(id).trim()) return i + 2;
  return -1;
}

function validate_(data, cfg) {
  const L = cfg.labels;
  const has = k => data[k] != null && String(data[k]).trim() !== '';
  if (has('group') && cfg.groups.indexOf(String(data.group).trim()) < 0)
    throw new Error(L.group + ' ต้องเป็นหนึ่งใน: ' + cfg.groups.join(', '));
  OPTION_KEYS.forEach(k => {
    if (has(k) && cfg.options[k].indexOf(String(data[k]).trim()) < 0)
      throw new Error(L[k] + ' ต้องเป็นหนึ่งใน: ' + cfg.options[k].join(', '));
  });
  if (has('studentId') && !/^\d{10}$/.test(String(data.studentId).trim()))
    throw new Error(L.studentId + ' ต้องเป็นตัวเลข 10 หลัก เช่น 6900000001');
}

function assertUnique_(key, value, exceptRow) {
  if (!value) return;
  const dup = readAll_().find(s => s[key] === value && s.row !== exceptRow);
  if (dup) throw new Error(getConfig_().labels[key] + ' ' + value + ' ซ้ำกับ ' + dup.name);
}

function updateStudent_(id, data, user, allowed, label) {
  const cfg = getConfig_();
  validate_(data, cfg);
  const row = findRow_(id);
  if (row < 0) throw new Error('ไม่พบเลขที่ ' + id);
  const range = sheet_().getRange(row, 1, 1, KEYS.length);
  const current = range.getValues()[0];
  const before = toObj_(current, row);
  const changes = [];
  Object.keys(data).forEach(k => {
    const i = KEYS.indexOf(k);
    if (i < 0 || k === 'updatedAt') return;
    const v = String(data[k] == null ? '' : data[k]).trim();
    if (v === before[k]) return;
    if (allowed.indexOf(k) < 0) throw new Error('คุณไม่มีสิทธิ์แก้ไข "' + cfg.labels[k] + '"');
    if (k === 'id' && !v) throw new Error('ต้องมี' + cfg.labels.id);
    if (k === 'name' && !v) throw new Error('ต้องมี' + cfg.labels.name);
    if (k === 'group' && !v) throw new Error('ต้องเลือก' + cfg.labels.group);
    if (k === 'id' || k === 'studentId') assertUnique_(k, v, row);
    current[i] = v;
    changes.push(cfg.labels[k] + ': ' + (before[k] || '-') + ' → ' + (v || '-'));
  });
  if (!changes.length) return before;
  current[KEYS.indexOf('updatedAt')] = new Date();
  range.setValues([current]);
  log_(user, label, before.id, changes.join(' | '));
  return toObj_(current, row);
}

function addStudent_(data, user) {
  const cfg = getConfig_();
  validate_(data, cfg);
  const v = k => String(data[k] == null ? '' : data[k]).trim();
  if (!v('id') || !v('name') || !v('group'))
    throw new Error('ต้องระบุ ' + [cfg.labels.group, cfg.labels.id, cfg.labels.name].join(', '));
  assertUnique_('id', v('id'), -1);
  assertUnique_('studentId', v('studentId'), -1);
  const sh = sheet_();
  const row = KEYS.map(k => k === 'updatedAt' ? new Date() : v(k));
  const r = sh.getLastRow() + 1;
  sh.getRange(r, col_('id'), 1, 2).setNumberFormat('@');
  sh.getRange(r, 1, 1, KEYS.length).setValues([row]);
  log_(user, 'เพิ่มนักศึกษา', v('id'), v('name') + ' → ' + v('group'));
}

function deleteStudent_(id, user) {
  const row = findRow_(id);
  if (row < 0) throw new Error('ไม่พบเลขที่ ' + id);
  const sh = sheet_();
  const before = toObj_(sh.getRange(row, 1, 1, KEYS.length).getValues()[0], row);
  sh.deleteRow(row);
  log_(user, 'ลบนักศึกษา', id, before.name + ' (' + before.group + ')');
}

/* ═════════════════════════ กลุ่ม / ตั้งค่า (ผู้ดูแล) ═════════════════════════ */

function addGroup_(name, user) {
  const cfg = getConfig_();
  name = String(name || '').trim();
  if (!name) throw new Error('กรุณาใส่ชื่อกลุ่ม');
  if (cfg.groups.indexOf(name) >= 0) throw new Error('มีกลุ่ม "' + name + '" อยู่แล้ว');
  cfg.groups.push(name);
  persistConfig_(cfg);
  log_(user, 'เพิ่มกลุ่ม', '', name);
}

function renameGroup_(from, to, user) {
  const cfg = getConfig_();
  from = String(from || '').trim(); to = String(to || '').trim();
  const i = cfg.groups.indexOf(from);
  if (i < 0) throw new Error('ไม่พบกลุ่ม "' + from + '"');
  if (!to) throw new Error('กรุณาใส่ชื่อกลุ่มใหม่');
  if (to === from) return;
  if (cfg.groups.indexOf(to) >= 0) throw new Error('มีกลุ่ม "' + to + '" อยู่แล้ว');
  cfg.groups[i] = to;
  // ปลด validation ก่อน แล้วเปลี่ยนชื่อกลุ่มของสมาชิกทุกคน
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    const rng = sh.getRange(2, col_('group'), n, 1);
    rng.clearDataValidations();
    rng.setValues(rng.getValues().map(r => [String(r[0]).trim() === from ? to : r[0]]));
  }
  persistConfig_(cfg);
  log_(user, 'เปลี่ยนชื่อกลุ่ม', '', from + ' → ' + to);
}

function deleteGroup_(name, user) {
  const cfg = getConfig_();
  name = String(name || '').trim();
  const i = cfg.groups.indexOf(name);
  if (i < 0) throw new Error('ไม่พบกลุ่ม "' + name + '"');
  const members = readAll_().filter(s => s.group === name).length;
  if (members) throw new Error('กลุ่ม "' + name + '" ยังมีสมาชิก ' + members + ' คน — ย้ายออกก่อนจึงจะลบได้');
  cfg.groups.splice(i, 1);
  persistConfig_(cfg);
  log_(user, 'ลบกลุ่ม', '', name);
}

function moveGroup_(name, direction, user) {
  const cfg = getConfig_();
  const i = cfg.groups.indexOf(String(name || '').trim());
  const j = direction === 'left' ? i - 1 : i + 1;
  if (i < 0 || j < 0 || j >= cfg.groups.length) return;
  const t = cfg.groups[i]; cfg.groups[i] = cfg.groups[j]; cfg.groups[j] = t;
  persistConfig_(cfg);
}

function saveConfig_(input, user) {
  const cfg = getConfig_();
  const before = JSON.stringify(cfg);
  if (input.title != null) {
    if (!String(input.title).trim()) throw new Error('ต้องมีชื่อหัวเรื่อง');
    cfg.title = String(input.title).trim();
  }
  if (input.subtitle != null) cfg.subtitle = String(input.subtitle).trim();
  if (input.labels) KEYS.forEach(k => {
    if (input.labels[k] == null) return;
    const v = String(input.labels[k]).trim();
    if (!v) throw new Error('ชื่อหัวคอลัมน์ต้องไม่ว่าง');
    cfg.labels[k] = v;
  });
  if (input.options) OPTION_KEYS.forEach(k => {
    if (input.options[k] == null) return;
    const list = uniq_(Array.isArray(input.options[k]) ? input.options[k] : lines_(input.options[k]));
    if (!list.length) throw new Error('ตัวเลือก "' + cfg.labels[k] + '" ต้องมีอย่างน้อย 1 รายการ');
    cfg.options[k] = list;
  });
  persistConfig_(cfg);
  log_(user, 'แก้ไขการตั้งค่า', '', before === JSON.stringify(cfg) ? '(ไม่มีการเปลี่ยนแปลง)' : 'บันทึกการตั้งค่าใหม่');
}

function persistConfig_(cfg) {
  saveConfigSheet_(cfg);
  writeHeaders_(cfg);
  applyValidation_(cfg);
  buildSummarySheet_(cfg);
}

/* ═════════════════════════ ผู้ใช้ (ผู้ดูแล) ═════════════════════════ */

function roleOfCode_(code) {
  const c = String(code || '').trim().toUpperCase();
  if (/^\d{4}$/.test(c)) return 'teacher';
  if (/^TA\d{4}$/.test(c)) return 'ta';
  return '';
}

function users_() {
  const sh = ss_().getSheetByName(USERS_NAME);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues()
    .map(r => ({ code: String(r[0]).trim().toUpperCase(), name: String(r[1]).trim(), note: String(r[2]).trim() }))
    .filter(u => u.code)
    .map(u => (u.role = roleOfCode_(u.code), u.roleName = ROLE_NAMES[u.role] || '⚠️ รูปแบบไม่ถูกต้อง', u))
    .filter(u => u.role);
}

function saveUsers_(list, user) {
  const seen = {};
  const rows = list.map(u => {
    const code = String(u.code || '').trim().toUpperCase();
    const name = String(u.name || '').trim();
    if (!code && !name) return null;
    if (!roleOfCode_(code)) throw new Error('รหัส "' + code + '" ไม่ถูกต้อง — อาจารย์ใช้ตัวเลข 4 หลัก (เช่น 6901), TA ใช้ TA + 4 หลัก (เช่น TA6901)');
    if (!name) throw new Error('รหัส ' + code + ' ยังไม่มีชื่อ');
    if (seen[code]) throw new Error('รหัส ' + code + ' ซ้ำกัน');
    seen[code] = true;
    return [code, name, String(u.note || '').trim()];
  }).filter(Boolean);
  const sh = ss_().getSheetByName(USERS_NAME);
  if (!sh) throw new Error('ไม่พบแผ่นงาน "' + USERS_NAME + '" — Run setup ก่อน');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  log_(user, 'แก้ไขรายชื่อผู้ใช้', '', 'อาจารย์/TA ทั้งหมด ' + rows.length + ' คน');
}

/* ═════════════════════════ แผ่นงาน ═════════════════════════ */

let CONFIG_CACHE = null;

function getConfig_() {
  if (CONFIG_CACHE) return clone_(CONFIG_CACHE);
  const cfg = clone_(DEFAULT_CONFIG);
  const sh = ss_().getSheetByName(CONFIG_NAME);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => {
      const k = String(r[0]).trim();
      const v = String(r[1] == null ? '' : r[1]);
      if (k === 'title' || k === 'subtitle') cfg[k] = v.trim();
      else if (k === 'groups') cfg.groups = uniq_(lines_(v));
      else if (k.indexOf('options.') === 0 && cfg.options[k.slice(8)]) cfg.options[k.slice(8)] = uniq_(lines_(v));
      else if (k.indexOf('label.') === 0 && cfg.labels[k.slice(6)] != null && v.trim()) cfg.labels[k.slice(6)] = v.trim();
    });
  }
  CONFIG_CACHE = cfg;
  return clone_(cfg);
}

function saveConfigSheet_(cfg) {
  const ss = ss_();
  const sh = ss.getSheetByName(CONFIG_NAME) || ss.insertSheet(CONFIG_NAME);
  const rows = [['คีย์', 'ค่า (ถ้าเป็นรายการ ให้ใส่บรรทัดละ 1 รายการ)'],
    ['title', cfg.title], ['subtitle', cfg.subtitle], ['groups', cfg.groups.join('\n')]]
    .concat(OPTION_KEYS.map(k => ['options.' + k, cfg.options[k].join('\n')]))
    .concat(KEYS.map(k => ['label.' + k, cfg.labels[k]]));
  sh.clearContents();
  sh.getRange(2, 2, rows.length - 1, 1).setNumberFormat('@').setWrap(true);
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 420);
  CONFIG_CACHE = clone_(cfg);
}

function writeHeaders_(cfg) {
  sheet_().getRange(1, 1, 1, KEYS.length).setValues([KEYS.map(k => cfg.labels[k])])
    .setFontWeight('bold').setBackground('#1f3b57').setFontColor('#ffffff');
}

function applyValidation_(cfg) {
  const sh = sheet_();
  const n = Math.max(sh.getMaxRows() - 1, 1);
  const setList = (key, list) => {
    const rng = sh.getRange(2, col_(key), n);
    if (!list.length) return rng.clearDataValidations();
    rng.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(list, true).setAllowInvalid(false).build());
  };
  setList('group', cfg.groups);
  OPTION_KEYS.forEach(k => setList(k, cfg.options[k]));
  sh.getRange(2, col_('studentId'), n).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=OR(' + colLetter_(col_('studentId')) + '2="",REGEXMATCH(' +
      colLetter_(col_('studentId')) + '2&"","^[0-9]{10}$"))')
    .setAllowInvalid(false).setHelpText('รหัสนิสิตต้องเป็นตัวเลข 10 หลัก').build());
}

function buildSummarySheet_(cfg) {
  const ss = ss_();
  const s = ss.getSheetByName(SUMMARY_NAME) || ss.insertSheet(SUMMARY_NAME);
  s.clear();
  const L = cfg.labels;
  const parts = [['status', cfg.options.status], ['practice', cfg.options.practice], ['consult', cfg.options.consult]];
  const head = [L.group, 'รวม'];
  parts.forEach(([k, list]) => list.forEach(v => head.push(L[k] + ': ' + v)));
  s.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#1f3b57').setFontColor('#ffffff');
  const src = "'" + SHEET_NAME + "'!";
  const G = colLetter_(col_('group'));
  if (cfg.groups.length) {
    const rows = cfg.groups.map((g, i) => {
      const r = i + 2;
      const out = ['=COUNTIF(' + src + '$' + G + ':$' + G + ',$A' + r + ')'];
      parts.forEach(([k, list]) => list.forEach(v => {
        const C = colLetter_(col_(k));
        out.push('=COUNTIFS(' + src + '$' + G + ':$' + G + ',$A' + r + ',' + src + '$' + C + ':$' + C + ',"' + v.replace(/"/g, '""') + '")');
      }));
      return out;
    });
    s.getRange(2, 1, cfg.groups.length, 1).setNumberFormat('@').setValues(cfg.groups.map(g => [g]));
    s.getRange(2, 2, rows.length, head.length - 1).setFormulas(rows);
    const totalRow = cfg.groups.length + 2;
    const totals = [];
    for (let c = 2; c <= head.length; c++) totals.push('=SUM(' + colLetter_(c) + '2:' + colLetter_(c) + (totalRow - 1) + ')');
    s.getRange(totalRow, 1).setValue('รวมทั้งหมด');
    s.getRange(totalRow, 2, 1, totals.length).setFormulas([totals]);
    s.getRange(totalRow, 1, 1, head.length).setFontWeight('bold').setBackground('#eef2f6');
  }
  s.setFrozenColumns(1);
  s.autoResizeColumns(1, head.length);
}

function summarize_(students, cfg) {
  return cfg.groups.map(g => {
    const m = students.filter(s => s.group === g);
    const by = k => cfg.options[k].reduce((o, v) => (o[v] = m.filter(s => s[k] === v).length, o), {});
    return { group: g, total: m.length, status: by('status'), practice: by('practice'), consult: by('consult'), career: by('career') };
  });
}

function snapshot_() {
  const ss = ss_();
  const name = BACKUP_PREFIX + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HHmm');
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  sheet_().copyTo(ss).setName(name).hideSheet();
  const backups = ss.getSheets().filter(s => s.getName().indexOf(BACKUP_PREFIX) === 0)
    .sort((a, b) => a.getName() < b.getName() ? -1 : 1);
  while (backups.length > KEEP_BACKUPS) ss.deleteSheet(backups.shift());
  log_({ role: 'system', name: 'ระบบสำรองอัตโนมัติ' }, 'สำรองข้อมูล', '', name);
  return name;
}

/* ═════════════════════════ ตัวช่วย ═════════════════════════ */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_() {
  const sh = ss_().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('ไม่พบแผ่นงาน "' + SHEET_NAME + '" — กรุณา Run ฟังก์ชัน setup ก่อน');
  return sh;
}

function col_(key) { return KEYS.indexOf(key) + 1; }

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function lines_(v) { return String(v || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean); }
function uniq_(list) { return list.map(s => String(s).trim()).filter((s, i, a) => s && a.indexOf(s) === i); }
function clone_(o) { return JSON.parse(JSON.stringify(o)); }

function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('ระบบกำลังบันทึกข้อมูลอื่น ลองใหม่อีกครั้ง');
  try { return fn(); } finally { lock.releaseLock(); }
}

function log_(user, action, id, detail) {
  const log = ss_().getSheetByName(LOG_NAME);
  if (!log) return;
  const who = user.role === 'system' ? user.name : (ROLE_NAMES[user.role] || '') + ' · ' + (user.name || '');
  log.appendRow([new Date(), who, action, id, detail]);
}

function checkKey_(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!expected) throw new Error('ยังไม่ได้ตั้ง API_KEY — กรุณา Run ฟังก์ชัน setup');
  if (!key || key !== expected) throw new Error('API key ไม่ถูกต้อง');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
