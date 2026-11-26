/**
 * ระบบจัดการการเรียนการสอน ป.จว. รุ่นที่ 1 (LMS) — ส่วนเซิร์ฟเวอร์
 *
 * หน้าเว็บ (docs/index.html) เรียกผ่าน doPost ใน Code.gs ด้วย action ที่ขึ้นต้นด้วย "lms"
 * ข้อมูลทั้งหมดเก็บในแท็บ "LMS …" ของ Google Sheet นี้ (คอลัมน์ json คือข้อมูลจริง คอลัมน์อื่นไว้อ่าน/แก้ง่าย)
 * แก้ค่าในคอลัมน์ที่อยู่ใน sheetEditable ได้โดยตรงใน Sheet (เช่น PIN, ชื่อ, ลิงก์คลิป, คำแนะนำอาจารย์)
 *
 * การเข้าสู่ระบบ (ตรวจที่เซิร์ฟเวอร์ทั้งหมด):
 *   นักศึกษา  เลขที่ (เช่น 01-043) หรือเลขลำดับ (043) + PIN (ตั้งต้น 1234 — ระบบจะชวนให้เปลี่ยนเมื่อเข้าครั้งแรก)
 *   อาจารย์ / TA  ชื่อผู้ใช้ + รหัสผ่าน (สุ่มให้ตอน setup ดูได้ใน Execution log และหน้า Admin)
 *   Admin  ชื่อผู้ใช้ admin + รหัสผ่านจาก INITIAL_ADMIN_PASSWORD ใน Seed.gs
 */

const LMS_SESSION_SECONDS = 6 * 60 * 60;
const LMS_MAX_JSON = 49000;

const LMS_COLL = {
  students: { sheet: 'LMS นักศึกษา', cols: ['code', 'seq', 'name', 'nickname', 'type', 'province', 'education', 'profession', 'status', 'pin', 'notes'],
    sheetEditable: ['name', 'nickname', 'type', 'province', 'education', 'profession', 'status', 'pin', 'notes'] },
  staff: { sheet: 'LMS บุคลากร', cols: ['role', 'username', 'password', 'displayName'], sheetEditable: ['username', 'password', 'displayName'] },
  sessions: { sheet: 'LMS ตารางเรียน', cols: ['date', 'time', 'title', 'type', 'instructor', 'location'],
    sheetEditable: ['date', 'time', 'title', 'type', 'instructor', 'location'] },
  replays: { sheet: 'LMS คลิปย้อนหลัง', cols: ['sessionId', 'date', 'title', 'url', 'notes'], sheetEditable: ['url', 'notes'] },
  attendance: { sheet: 'LMS เช็กชื่อ', cols: ['userId', 'sessionId', 'checkInTime', 'checkOutTime'], sheetEditable: ['checkInTime', 'checkOutTime'] },
  logs: { sheet: 'LMS บันทึกการปรึกษา', cols: ['studentId', 'studentCode', 'studentName', 'date', 'counselingMode', 'counselingDuration', 'instructorFeedback'],
    sheetEditable: ['instructorFeedback'] },
  selfAssessments: { sheet: 'LMS แบบประเมินตนเอง', cols: [], sheetEditable: [] },
  practiceByUser: { sheet: 'LMS บันทึกปฏิบัติธรรม', cols: ['daily30Rounds', 'daily30Minutes', 'classMinutes'], sheetEditable: [] },
  courseReflections: { sheet: 'LMS สะท้อนรายวิชา', cols: [], sheetEditable: [] },
  settings: { sheet: 'LMS ตั้งค่า', cols: [], sheetEditable: [] },
  // งานสอนของอาจารย์
  assignments: { sheet: 'LMS การบ้าน', cols: ['subject', 'title', 'dueDate', 'maxScore', 'createdBy'],
    sheetEditable: ['title', 'dueDate', 'maxScore'] },
  submissions: { sheet: 'LMS ส่งการบ้าน', cols: ['assignmentId', 'studentId', 'submittedAt', 'score', 'feedback', 'gradedBy'],
    sheetEditable: ['score', 'feedback'] },
  reviews: { sheet: 'LMS คะแนนอาจารย์', cols: ['studentId', 'subject', 'kind', 'score', 'maxScore', 'comment', 'by'],
    sheetEditable: ['score', 'maxScore', 'comment'] },
};

/** บทบาทที่ให้คะแนน/สั่งการบ้าน และบทบาทที่เช็กชื่อนักศึกษาได้ (เพิ่ม 'ta' ได้ถ้าต้องการ) */
const LMS_GRADER_ROLES = ['teacher', 'admin'];
const LMS_ATTENDANCE_ROLES = ['teacher', 'admin'];

const LMS_DEFAULT_COURSE_INFO = {
  name: 'ประกาศนียบัตรการปรึกษาเชิงจิตวิทยาแนววิปัสสนาภาวนา (ป.จว. รุ่น 1)',
  batch: 'รุ่นที่ 1 • ภาคเรียนที่ 1/2569',
  institution: 'ภาควิชาจิตวิทยา คณะมนุษยศาสตร์ มหาวิทยาลัยมหาจุฬาลงกรณราชวิทยาลัย',
  classroom: 'ห้องเรียน A 103 อาคารเรียนรวม โซน A อ.วังน้อย จ.พระนครศรีอยุธยา',
  zoomId: '954 258 8616',
  zoomPass: '2569',
  zoomUrl: 'https://zoom.us/j/9542588616',
  classHours: '18:00 - 21:00 น.'
};

/* ═════════════════════════ ติดตั้ง (เรียกจาก setup() ใน Code.gs) ═════════════════════════ */

function lmsSetup_() {
  Object.keys(LMS_COLL).forEach(name => lmsSheet_(name));
  const who = 'setup';
  const created = [];

  if (!Object.keys(lmsRead_('settings').map).length) {
    lmsWriteMany_('settings', { courseInfo: LMS_DEFAULT_COURSE_INFO, postAssessmentLocked: { value: true } }, who);
  }
  if (!Object.keys(lmsRead_('staff').map).length) {
    const staff = {};
    if (typeof INITIAL_ADMIN_PASSWORD === 'undefined' || !INITIAL_ADMIN_PASSWORD) throw new Error('ยังไม่ได้กำหนด INITIAL_ADMIN_PASSWORD ใน Seed.gs');
    staff.admin = { id: 'admin', role: 'admin', username: 'admin', password: INITIAL_ADMIN_PASSWORD, displayName: 'ผู้ดูแลระบบหลักสูตร ป.จว.' };
    (typeof LMS_SEED_TEACHERS !== 'undefined' ? LMS_SEED_TEACHERS : []).forEach(t => {
      staff[t.id] = { id: t.id, role: 'teacher', username: t.username, password: lmsRandomPass_(), displayName: t.displayName };
      created.push('อาจารย์ ' + t.username + ' / ' + staff[t.id].password);
    });
    (typeof LMS_SEED_TAS !== 'undefined' ? LMS_SEED_TAS : []).forEach(t => {
      staff[t.id] = { id: t.id, role: 'ta', username: t.username, password: lmsRandomPass_(), displayName: t.displayName };
      created.push('TA ' + t.username + ' / ' + staff[t.id].password);
    });
    lmsWriteMany_('staff', staff, who);
  }
  if (!Object.keys(lmsRead_('students').map).length) lmsSeedStudents_(who);
  if (!Object.keys(lmsRead_('sessions').map).length) lmsSeedSessions_(who);

  Logger.log('LMS พร้อมใช้งาน · Admin: ชื่อผู้ใช้ admin + รหัสผ่านจาก INITIAL_ADMIN_PASSWORD');
  if (created.length) Logger.log('รหัสผ่านอาจารย์/TA ที่สุ่มให้ (ดู/เปลี่ยนได้ในหน้า Admin):\n' + created.join('\n'));
}

function lmsSeedStudents_(who) {
  if (typeof LMS_SEED_STUDENTS === 'undefined') return;
  const m = {};
  LMS_SEED_STUDENTS.forEach(s => { m[s.id] = Object.assign({}, s, { pin: String(s.pin || '1234') }); });
  lmsWriteMany_('students', m, who);
}

function lmsSeedSessions_(who) {
  if (typeof LMS_SEED_SESSIONS === 'undefined') return;
  const ses = {}, rep = {};
  LMS_SEED_SESSIONS.forEach(s => {
    ses[s.id] = Object.assign({}, s);
    rep['REP-' + s.id] = { id: 'REP-' + s.id, sessionId: s.id, date: s.date, title: s.title, instructor: s.instructor, url: '', notes: '' };
  });
  lmsWriteMany_('sessions', ses, who);
  if (!Object.keys(lmsRead_('replays').map).length) lmsWriteMany_('replays', rep, who);
}

function lmsRandomPass_() {
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 6; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

/* ═════════════════════════ ตัวจัดการคำขอ ═════════════════════════ */

function lmsHandle_(body) {
  switch (body.action) {
    case 'lmsPublic': return lmsPublic_();
    case 'lmsLogin': return lmsLogin_(body);
  }
  const user = lmsAuth_(body.token);
  switch (body.action) {
    case 'lmsLoad': return Object.assign({ ok: true, session: lmsPublicSession_(user) }, lmsData_(user));
    case 'lmsSave': return withLock_(() => lmsSave_(user, body.changes || {}));
    case 'lmsChangePassword': return withLock_(() => lmsChangePassword_(user, body.oldPassword, body.newPassword));
    case 'lmsLogout':
      CacheService.getScriptCache().remove('lms:' + body.token);
      return { ok: true };
    case 'lmsClearTrial': return withLock_(() => lmsClear_(user, body.adminPassword, false));
    case 'lmsReset': return withLock_(() => lmsClear_(user, body.adminPassword, true));
    default: throw new Error('ไม่รู้จัก action: ' + body.action);
  }
}

/** ข้อมูลก่อนเข้าสู่ระบบ: ชื่อหลักสูตร + รายชื่อบัญชีอาจารย์/TA (ไม่มีรหัสผ่าน) สำหรับ dropdown */
function lmsPublic_() {
  const staff = lmsValues_(lmsRead_('staff').map);
  const pick = r => staff.filter(s => s.role === r).map(s => ({ id: s.id, username: s.username, displayName: s.displayName }));
  return { ok: true, courseInfo: lmsSettings_().courseInfo, teacherAccounts: pick('teacher'), taAccounts: pick('ta') };
}

function lmsLogin_(body) {
  const role = String(body.role || 'student');
  const code = String(body.code || '').trim();
  const pin = String(body.pin || '').trim();
  if (!code || !pin) throw new Error('กรุณากรอกรหัสประจำตัวและรหัสผ่าน');

  const cache = CacheService.getScriptCache();
  const failKey = 'lmsfail:' + hash_(role + ':' + code.toLowerCase());
  if (Number(cache.get('lmsfail:all') || 0) >= 60) throw new Error('มีการใส่รหัสผิดจำนวนมาก ระบบปิดการเข้าสู่ระบบชั่วคราว 10 นาที');
  if (Number(cache.get(failKey) || 0) >= 5) throw new Error('ใส่รหัสผ่านผิดเกิน 5 ครั้ง กรุณารอ 15 นาที');
  const fail = msg => {
    cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), 900);
    cache.put('lmsfail:all', String(Number(cache.get('lmsfail:all') || 0) + 1), 600);
    return new Error(msg);
  };

  let session = null, mustChangePin = false;
  if (role === 'student') {
    const lc = code.toLowerCase();
    const st = lmsValues_(lmsRead_('students').map).find(s => String(s.code).toLowerCase() === lc || String(s.seq).toLowerCase() === lc);
    if (!st) throw fail('ไม่พบรหัสนักศึกษาในระบบ กรุณาตรวจสอบอีกครั้ง');
    if (String(st.pin || '1234') !== pin) throw fail('รหัสผ่านไม่ถูกต้อง');
    session = { role: 'student', userId: st.id, username: st.code, displayName: st.name, subtitle: 'รหัส ' + st.code + ' • ' + st.type };
    mustChangePin = String(st.pin || '1234') === '1234';
  } else if (role === 'teacher' || role === 'ta' || role === 'admin') {
    const lc = code.toLowerCase();
    const acc = lmsValues_(lmsRead_('staff').map).find(s => s.role === role && String(s.username).toLowerCase() === lc);
    const label = { teacher: 'อาจารย์ผู้สอน', ta: 'ผู้ช่วยสอน (TA)', admin: 'ผู้ดูแลระบบ' }[role];
    if (!acc) throw fail('ไม่พบบัญชี' + label + 'ในระบบ');
    if (String(acc.password) !== pin) throw fail('รหัสผ่าน' + label + 'ไม่ถูกต้อง');
    const sub = { teacher: 'คณาจารย์ผู้สอน (' + acc.username + ')', ta: 'ผู้ช่วยสอน TA (' + acc.username + ')', admin: 'ผู้ดูแลระบบหลักสูตร ป.จว.' }[role];
    session = { role: role, userId: acc.id, username: acc.username, displayName: acc.displayName, subtitle: sub };
  } else {
    throw new Error('บทบาทไม่ถูกต้อง');
  }
  cache.remove(failKey);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache.put('lms:' + token, JSON.stringify(session), LMS_SESSION_SECONDS);
  log_({ role: session.role === 'admin' ? 'admin' : session.role, name: session.displayName }, 'LMS เข้าสู่ระบบ', session.username, '');
  return Object.assign({ ok: true, token: token, session: session, mustChangePin: mustChangePin }, lmsData_(session));
}

function lmsAuth_(token) {
  const fail = msg => { const e = new Error(msg); e.auth = true; return e; };
  if (!token) throw fail('กรุณาเข้าสู่ระบบ');
  const raw = CacheService.getScriptCache().get('lms:' + token);
  if (!raw) throw fail('หมดเวลาการใช้งาน กรุณาเข้าสู่ระบบใหม่');
  const s = JSON.parse(raw);
  const coll = s.role === 'student' ? 'students' : 'staff';
  if (!lmsRead_(coll).map[s.userId]) throw fail('บัญชีของคุณถูกยกเลิกแล้ว');
  return s;
}

function lmsPublicSession_(s) {
  return { role: s.role, userId: s.userId, username: s.username, displayName: s.displayName, subtitle: s.subtitle };
}

/* ═════════════════════════ ข้อมูลตามสิทธิ์ ═════════════════════════ */

function lmsData_(s) {
  const isAdmin = s.role === 'admin', isStaff = s.role !== 'student', uid = s.userId;
  const omit = (o, k) => { const c = Object.assign({}, o); delete c[k]; return c; };

  const students = lmsValues_(lmsRead_('students').map);
  const staff = lmsValues_(lmsRead_('staff').map);
  const settings = lmsSettings_();
  const att = lmsRead_('attendance').map;
  const assess = lmsRead_('selfAssessments').map;
  const practice = lmsRead_('practiceByUser').map;
  const refl = lmsRead_('courseReflections').map;
  const byDate = (a, b) => String(a.date || '').localeCompare(String(b.date || '')) || String(a.id).localeCompare(String(b.id));
  const logNum = l => Number(String(l.id || '').replace(/\D/g, '')) || 0;

  const staffOf = role => staff.filter(x => x.role === role)
    .map(x => isAdmin ? { id: x.id, username: x.username, password: x.password, displayName: x.displayName }
      : { id: x.id, username: x.username, displayName: x.displayName });
  const adminRow = staff.find(x => x.role === 'admin') || {};
  const filterKeys = (m, keep) => Object.keys(m).filter(keep).reduce((o, k) => (o[k] = m[k], o), {});

  return {
    data: {
      courseInfo: settings.courseInfo,
      postAssessmentLocked: settings.postAssessmentLocked,
      studentCount: students.length,
      students: isStaff ? students.map(x => isAdmin ? x : omit(x, 'pin')) : students.filter(x => x.id === uid).map(x => omit(x, 'pin')),
      teacherAccounts: staffOf('teacher'),
      taAccounts: staffOf('ta'),
      adminAuth: isAdmin ? { username: adminRow.username, password: adminRow.password, displayName: adminRow.displayName }
        : { displayName: adminRow.displayName || 'ผู้ดูแลระบบหลักสูตร ป.จว.' },
      sessions: lmsValues_(lmsRead_('sessions').map).sort(byDate),
      replays: lmsValues_(lmsRead_('replays').map).sort(byDate),
      attendance: isStaff ? att : filterKeys(att, k => k.indexOf(uid + '_') === 0),
      logs: lmsValues_(lmsRead_('logs').map).filter(l => isStaff || l.studentId === uid).sort((a, b) => logNum(b) - logNum(a)),
      selfAssessments: isStaff ? assess : filterKeys(assess, k => k === uid),
      // บุคลากรเห็นบันทึกปฏิบัติธรรมของนักศึกษาทุกคน (เพื่อตรวจ) · นักศึกษาเห็นของตนเอง
      practiceByUser: isStaff ? practice : filterKeys(practice, k => k === uid),
      courseReflections: isStaff ? refl : {},
      assignments: lmsValues_(lmsRead_('assignments').map).sort((a, b) =>
        String(a.subject || '').localeCompare(String(b.subject || '')) || String(a.dueDate || '').localeCompare(String(b.dueDate || ''))),
      submissions: isStaff ? lmsRead_('submissions').map : filterKeys(lmsRead_('submissions').map, k => lmsSubStudent_(k) === uid),
      reviews: isStaff ? lmsRead_('reviews').map : filterKeys(lmsRead_('reviews').map, k => k.indexOf(uid + '_') === 0),
    }
  };
}

/** key ของการส่งการบ้าน = "<assignmentId>_<studentId>" */
function lmsSubStudent_(key) { const i = String(key).indexOf('_'); return i < 0 ? '' : String(key).slice(i + 1); }
function lmsSubAssignment_(key) { const i = String(key).indexOf('_'); return i < 0 ? '' : String(key).slice(0, i); }

function lmsSettings_() {
  const m = lmsRead_('settings').map;
  return {
    courseInfo: Object.assign({}, LMS_DEFAULT_COURSE_INFO, m.courseInfo || {}),
    postAssessmentLocked: m.postAssessmentLocked ? !!m.postAssessmentLocked.value : true,
  };
}

/* ═════════════════════════ บันทึกการเปลี่ยนแปลง (ตรวจสิทธิ์ทีละรายการ) ═════════════════════════ */

function lmsSave_(s, changes) {
  const isAdmin = s.role === 'admin', uid = s.userId;
  const applied = [], skipped = [];
  const settingsNow = lmsSettings_();
  const writes = {}; // coll -> {key: obj|null}
  const put = (coll, key, val) => { (writes[coll] = writes[coll] || {})[key] = val; applied.push(coll + ':' + key); };
  const skip = (coll, key, why) => skipped.push(coll + ':' + key + (why ? ' (' + why + ')' : ''));
  const cache = {};
  const cur = coll => cache[coll] || (cache[coll] = lmsRead_(coll).map);

  Object.keys(changes).forEach(clientColl => {
    const items = changes[clientColl] || {};
    Object.keys(items).forEach(key => {
      const val = items[key];
      const isDel = val === null;
      try {
        switch (clientColl) {
          case 'students':
            if (!isAdmin) return skip(clientColl, key);
            if (isDel) return put('students', key, null);
            return put('students', key, Object.assign({}, val, { id: key, pin: String(val.pin || (cur('students')[key] || {}).pin || '1234') }));

          case 'teacherAccounts':
          case 'taAccounts': {
            if (!isAdmin) return skip(clientColl, key);
            const role = clientColl === 'teacherAccounts' ? 'teacher' : 'ta';
            const old = cur('staff')[key];
            if (old && old.role === 'admin') return skip(clientColl, key, 'admin');
            if (isDel) return put('staff', key, null);
            return put('staff', key, { id: key, role: role, username: String(val.username || ''), password: String(val.password || (old || {}).password || lmsRandomPass_()), displayName: String(val.displayName || '') });
          }

          case 'sessions':
          case 'courseReflections':
            if (!isAdmin) return skip(clientColl, key);
            return put(clientColl, key, isDel ? null : Object.assign({}, val, clientColl === 'sessions' ? { id: key } : {}));

          case 'settings':
            if (!isAdmin) return skip(clientColl, key);
            if (key === 'courseInfo') return put('settings', 'courseInfo', val || {});
            if (key === 'postAssessmentLocked') return put('settings', 'postAssessmentLocked', { value: !!(val && val.value) });
            return skip(clientColl, key);

          case 'replays': {
            if (!(isAdmin || s.role === 'ta')) return skip(clientColl, key);
            if (isDel) return isAdmin ? put('replays', key, null) : skip(clientColl, key);
            const old = cur('replays')[key] || {};
            const merged = isAdmin ? Object.assign({}, val) : Object.assign({}, old, {
              url: val.url, notes: val.notes, uploadedFileName: val.uploadedFileName, uploadedFileSize: val.uploadedFileSize, uploadedBy: val.uploadedBy });
            if (/^blob:/i.test(String(merged.url || ''))) merged.url = ''; // คลิปที่เลือกจากเครื่องเปิดได้เฉพาะเครื่องนั้น
            merged.id = key;
            return put('replays', key, merged);
          }

          case 'attendance': {
            if (isAdmin) return put('attendance', key, val);
            if (isDel) return skip(clientColl, key);
            if (key.indexOf(uid + '_') === 0) {
              // บันทึกของตนเอง — คงผลที่อาจารย์บันทึกไว้
              const old = cur('attendance')[key] || {};
              const mine = Object.assign({}, val);
              ['teacherStatus', 'teacherNote', 'markedBy', 'markedAt'].forEach(f => { if (old[f] !== undefined) mine[f] = old[f]; else delete mine[f]; });
              return put('attendance', key, mine);
            }
            if (LMS_ATTENDANCE_ROLES.indexOf(s.role) >= 0) {
              // อาจารย์เช็กชื่อนักศึกษา: แก้ได้เฉพาะสถานะที่อาจารย์บันทึก
              const old = cur('attendance')[key] || {};
              return put('attendance', key, Object.assign({}, old, {
                teacherStatus: String(val.teacherStatus || ''), teacherNote: String(val.teacherNote || ''),
                markedBy: s.displayName, markedAt: new Date().toISOString() }));
            }
            return skip(clientColl, key);
          }

          case 'assignments': {
            if (LMS_GRADER_ROLES.indexOf(s.role) < 0) return skip(clientColl, key);
            if (isDel) return put('assignments', key, null);
            const old = cur('assignments')[key] || {};
            return put('assignments', key, Object.assign({}, val, { id: key, createdBy: old.createdBy || s.displayName,
              maxScore: String(Number(val.maxScore) || 0) }));
          }

          case 'submissions': {
            const old = cur('submissions')[key];
            const sid = lmsSubStudent_(key), aid = lmsSubAssignment_(key);
            if (s.role === 'student') {
              if (sid !== uid || isDel) return skip(clientColl, key);
              if (!cur('assignments')[aid]) return skip(clientColl, key, 'ไม่มีการบ้านนี้');
              const keep = old ? { score: old.score, feedback: old.feedback, gradedBy: old.gradedBy, gradedAt: old.gradedAt } : {};
              return put('submissions', key, Object.assign({ assignmentId: aid, studentId: uid }, {
                text: String(val.text || ''), link: String(val.link || ''), submittedAt: new Date().toISOString() }, keep));
            }
            if (LMS_GRADER_ROLES.indexOf(s.role) < 0) return skip(clientColl, key);
            if (isDel) return isAdmin ? put('submissions', key, null) : skip(clientColl, key);
            return put('submissions', key, Object.assign({ assignmentId: aid, studentId: sid }, old || {}, {
              score: val.score === '' || val.score == null ? '' : String(Number(val.score) || 0),
              feedback: String(val.feedback || ''), gradedBy: s.displayName, gradedAt: new Date().toISOString() }));
          }

          case 'reviews': {
            if (LMS_GRADER_ROLES.indexOf(s.role) < 0) return skip(clientColl, key);
            if (isDel) return put('reviews', key, null);
            return put('reviews', key, Object.assign({}, val, { by: s.displayName, at: new Date().toISOString() }));
          }

          case 'practiceByUser':
            if (!isAdmin && key !== uid) return skip(clientColl, key);
            return put('practiceByUser', key, isDel ? null : val);

          case 'selfAssessments': {
            if (isAdmin) return put('selfAssessments', key, val);
            if (s.role !== 'student' || key !== uid || isDel) return skip(clientColl, key);
            const old = cur('selfAssessments')[key] || {};
            const merged = Object.assign({}, val);
            if (settingsNow.postAssessmentLocked) { if (old.post) merged.post = old.post; else delete merged.post; }
            return put('selfAssessments', key, merged);
          }

          case 'logs': {
            if (isDel) return isAdmin ? put('logs', key, null) : skip(clientColl, key);
            const old = cur('logs')[key];
            if (isAdmin) return put('logs', key, Object.assign({}, val, { id: key }));
            if (s.role === 'student') {
              if (old && old.studentId !== uid) return skip(clientColl, key);
              const me = cur('students')[uid] || {};
              return put('logs', key, Object.assign({}, val, { id: key, studentId: uid, studentCode: me.code || val.studentCode, studentName: me.name || val.studentName,
                instructorFeedback: old ? (old.instructorFeedback || '') : '',
                subject: old ? old.subject : undefined, score: old ? old.score : undefined, maxScore: old ? old.maxScore : undefined,
                reviewedBy: old ? old.reviewedBy : undefined, reviewedAt: old ? old.reviewedAt : undefined }));
            }
            if (!old) return put('logs', key, Object.assign({}, val, { id: key, instructorFeedback: s.role === 'teacher' ? (val.instructorFeedback || '') : '' }));
            if (LMS_GRADER_ROLES.indexOf(s.role) >= 0) return put('logs', key, Object.assign({}, old, {
              instructorFeedback: String(val.instructorFeedback || ''), subject: String(val.subject || old.subject || ''),
              score: val.score === undefined ? (old.score === undefined ? '' : old.score)
                : (val.score === '' || val.score === null ? '' : String(Number(val.score) || 0)),
              maxScore: String(Number(val.maxScore) || Number(old.maxScore) || 10),
              reviewedBy: s.displayName, reviewedAt: new Date().toISOString() }));
            return skip(clientColl, key);
          }

          default:
            return skip(clientColl, key, 'unknown');
        }
      } catch (err) {
        skip(clientColl, key, String(err.message || err));
      }
    });
  });

  const who = s.username + ' (' + s.role + ')';
  Object.keys(writes).forEach(coll => lmsWriteMany_(coll, writes[coll], who));
  return { ok: true, applied: applied.length, skipped: skipped };
}

function lmsChangePassword_(s, oldPass, newPass) {
  oldPass = String(oldPass || '').trim(); newPass = String(newPass || '').trim();
  if (newPass.length < 4) throw new Error('รหัสผ่านใหม่ควรมีความยาวอย่างน้อย 4 ตัวอักษร');
  if (s.role === 'student') {
    const st = lmsRead_('students').map[s.userId];
    if (!st || String(st.pin || '1234') !== oldPass) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
    lmsWriteMany_('students', { [s.userId]: Object.assign({}, st, { pin: newPass }) }, s.username);
  } else {
    const acc = lmsRead_('staff').map[s.userId];
    if (!acc || String(acc.password) !== oldPass) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
    lmsWriteMany_('staff', { [s.userId]: Object.assign({}, acc, { password: newPass }) }, s.username);
  }
  return Object.assign({ ok: true }, lmsData_(s));
}

/** ล้างข้อมูลทดลอง (เช็กชื่อ บันทึกการปรึกษา แบบประเมิน บันทึกปฏิบัติธรรม) / รีเซ็ตตั้งต้น (เพิ่มรายชื่อ ตารางเรียน คลิป ตั้งค่า) */
function lmsClear_(s, adminPassword, fullReset) {
  if (s.role !== 'admin') throw new Error('เฉพาะผู้ดูแลระบบ');
  const acc = lmsRead_('staff').map[s.userId];
  if (!acc || String(acc.password) !== String(adminPassword || '')) throw new Error('รหัสผ่านผู้ดูแลระบบ (Admin) ไม่ถูกต้อง');
  snapshotLms_();
  ['attendance', 'logs', 'selfAssessments', 'practiceByUser', 'courseReflections', 'submissions', 'reviews'].forEach(lmsClearSheet_);
  if (fullReset) {
    ['students', 'sessions', 'replays', 'settings', 'assignments'].forEach(lmsClearSheet_);
    lmsWriteMany_('settings', { courseInfo: LMS_DEFAULT_COURSE_INFO, postAssessmentLocked: { value: true } }, s.username);
    lmsSeedStudents_(s.username);
    lmsSeedSessions_(s.username);
  }
  log_({ role: 'admin', name: s.displayName }, fullReset ? 'LMS รีเซ็ตข้อมูลตั้งต้น' : 'LMS ล้างข้อมูลทดลอง', '', '');
  return Object.assign({ ok: true }, lmsData_(s));
}

/** สำรองแท็บ LMS ทั้งหมดก่อนล้าง (ซ่อนไว้ ชื่อ "สำรอง LMS …") */
function snapshotLms_() {
  const ss = ss_();
  const stamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HHmm');
  ['attendance', 'logs', 'selfAssessments', 'practiceByUser', 'students', 'assignments', 'submissions', 'reviews'].forEach(name => {
    const sh = ss.getSheetByName(LMS_COLL[name].sheet);
    if (sh && sh.getLastRow() > 1) sh.copyTo(ss).setName(('สำรอง LMS ' + stamp + ' ' + name).slice(0, 99)).hideSheet();
  });
}

/* ═════════════════════════ การอ่าน/เขียนแผ่นงาน ═════════════════════════ */

function lmsSheet_(name) {
  const def = LMS_COLL[name];
  const ss = ss_();
  let sh = ss.getSheetByName(def.sheet);
  if (!sh) {
    sh = ss.insertSheet(def.sheet);
    const head = ['key'].concat(def.cols, ['json', 'updatedAt', 'updatedBy']);
    sh.getRange(1, 1, 1, sh.getMaxColumns()).setNumberFormat('@');
    sh.getRange('A:' + colLetter_(head.length)).setNumberFormat('@'); // เก็บเป็นข้อความทั้งหมด กันวันที่/เลขถูกแปลง
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#1B5544').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setFrozenColumns(1);
  }
  return sh;
}

function lmsRead_(name) {
  const def = LMS_COLL[name];
  const sh = lmsSheet_(name);
  const map = {}, rows = {};
  const last = sh.getLastRow();
  if (last < 2) return { map: map, rows: rows, sheet: sh };
  const width = def.cols.length + 4;
  const vals = sh.getRange(2, 1, last - 1, width).getDisplayValues();
  const jsonIdx = 1 + def.cols.length;
  vals.forEach((r, i) => {
    const key = String(r[0]).trim();
    if (!key) return;
    let obj = {};
    try { obj = r[jsonIdx] ? JSON.parse(r[jsonIdx]) : {}; } catch (e) { obj = {}; }
    if (obj === null || typeof obj !== 'object') obj = { value: obj };
    def.sheetEditable.forEach(c => {
      const cell = r[1 + def.cols.indexOf(c)];
      if (cell !== '' && cell != null) obj[c] = String(cell);
    });
    if (['attendance', 'selfAssessments', 'practiceByUser', 'courseReflections', 'settings', 'submissions', 'reviews'].indexOf(name) < 0 && !obj.id) obj.id = key;
    map[key] = obj;
    rows[key] = i + 2;
  });
  return { map: map, rows: rows, sheet: sh };
}

function lmsValues_(map) { return Object.keys(map).map(k => map[k]); }

function lmsCell_(name, key, obj, col) {
  if (name === 'attendance' && (col === 'userId' || col === 'sessionId')) {
    const i = key.indexOf('_');
    return col === 'userId' ? key.slice(0, i) : key.slice(i + 1);
  }
  if (name === 'practiceByUser') {
    const d30 = Array.isArray(obj.daily30) ? obj.daily30 : [];
    const cls = Array.isArray(obj.classPractice) ? obj.classPractice : [];
    if (col === 'daily30Rounds') return String(d30.length);
    if (col === 'daily30Minutes') return String(d30.reduce((a, x) => a + (Number(x.durationMinutes) || 0), 0));
    if (col === 'classMinutes') return String(cls.reduce((a, x) => a + (Number(x.hours) || 0) * 60 + (Number(x.minutes) || 0), 0));
  }
  const v = obj[col];
  if (v == null) return '';
  return typeof v === 'object' ? JSON.stringify(v) : String(v);
}

/** เขียนหลายรายการ: {key: obj} = เพิ่ม/แทนที่, {key: null} = ลบ */
function lmsWriteMany_(name, items, who) {
  const def = LMS_COLL[name];
  const cur = lmsRead_(name);
  const sh = cur.sheet;
  const now = new Date().toISOString();
  const appends = [], deletes = [];
  Object.keys(items).forEach(key => {
    const obj = items[key];
    if (obj === null) { if (cur.rows[key]) deletes.push(cur.rows[key]); return; }
    const json = JSON.stringify(obj);
    if (json.length > LMS_MAX_JSON) throw new Error('ข้อมูลของ ' + key + ' ยาวเกินกว่าที่ Sheet เก็บได้ในช่องเดียว');
    const row = [key].concat(def.cols.map(c => lmsCell_(name, key, obj, c)), [json, now, who || '']);
    if (cur.rows[key]) sh.getRange(cur.rows[key], 1, 1, row.length).setValues([row]);
    else appends.push(row);
  });
  if (appends.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, appends.length, appends[0].length).setNumberFormat('@').setValues(appends);
  }
  // Sheets ไม่ยอมให้ลบแถวจนเหลือแต่แถวหัวที่ตรึงไว้ จึงล้างเนื้อหาแทนเมื่อเหลือแถวน้อย
  deletes.sort((a, b) => b - a).forEach(r => {
    if (sh.getMaxRows() > 2) sh.deleteRow(r);
    else sh.getRange(r, 1, 1, sh.getMaxColumns()).clearContent();
  });
}

function lmsClearSheet_(name) {
  const sh = lmsSheet_(name);
  if (sh.getMaxRows() > 1) sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns()).clearContent();
}
