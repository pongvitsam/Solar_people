/**
 * Solar People — ระบบลงทะเบียนเข้างาน (Backend API)
 *
 * Frontend (GitHub Pages) เรียก API นี้ผ่าน Web App URL
 *   GET  ?action=list                    → รายชื่อบริษัท + สรุป (ไม่ต้องล็อกอิน)
 *   POST {action:'login', username}      → รับ token
 *   POST {action:'checkin', token, id, count}
 *   POST {action:'undo', token, id}
 *   POST {action:'users' | 'addUser' | 'updateUser' | 'deleteUser', token, ...}  (แอดมินเท่านั้น)
 *
 * ครั้งแรก: รันฟังก์ชัน setup() ใน editor หนึ่งครั้ง เพื่อสร้าง Google Sheet และอนุญาตสิทธิ์
 */

const SHEET = { COMPANIES: 'Companies', USERS: 'Users', LOG: 'Log' };

const COMPANY_HEADERS = ['PEA No.', 'ชื่อบริษัท', 'เลขที่ มท.', 'ลงวันที่', 'สถานะ', 'จำนวนคน', 'เวลาเข้างาน', 'ผู้บันทึก'];
const C = { ID: 0, NAME: 1, MT: 2, DATE: 3, STATUS: 4, COUNT: 5, TIME: 6, BY: 7 };

const USER_HEADERS = ['Username', 'ชื่อที่แสดง', 'Role', 'Active', 'สร้างเมื่อ'];
const U = { USERNAME: 0, NAME: 1, ROLE: 2, ACTIVE: 3, CREATED: 4 };

const LOG_HEADERS = ['เวลา', 'Action', 'CompanyID', 'ชื่อบริษัท', 'จำนวนคน', 'ผู้ทำรายการ'];

const STATUS_ARRIVED = 'มาแล้ว';
const SESSION_TTL_SEC = 6 * 60 * 60; // 6 ชั่วโมง (สูงสุดของ CacheService)
const USERNAME_RE = /^[A-Za-z0-9._-]{2,30}$/;

/* ------------------------------------------------------------------ */
/* Setup                                                              */
/* ------------------------------------------------------------------ */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const existingId = props.getProperty('SPREADSHEET_ID');
  if (existingId) {
    try { ss = SpreadsheetApp.openById(existingId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Solar People - ลงทะเบียนเข้างาน');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }
  ss.setSpreadsheetTimeZone('Asia/Bangkok');

  const companies = ensureSheet_(ss, SHEET.COMPANIES, COMPANY_HEADERS);
  companies.getRange('A:D').setNumberFormat('@');
  companies.getRange('F:F').setNumberFormat('0');
  companies.getRange('G:G').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  // นำเข้ารายชื่อจากไฟล์ Excel (ชีต 20k_N) เฉพาะตอนที่ยังไม่มีข้อมูล
  if (companies.getLastRow() < 2 && typeof SEED_COMPANIES !== 'undefined' && SEED_COMPANIES.length) {
    companies.getRange(2, 1, SEED_COMPANIES.length, 4).setValues(SEED_COMPANIES);
    Logger.log('นำเข้ารายชื่อบริษัท ' + SEED_COMPANIES.length + ' รายการ');
  }
  companies.getRange('E2:E').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList([STATUS_ARRIVED], true).setAllowInvalid(true).build()
  );
  companies.setColumnWidth(C.NAME + 1, 320);

  const users = ensureSheet_(ss, SHEET.USERS, USER_HEADERS);
  if (users.getLastRow() < 2) {
    const now = new Date();
    users.getRange(2, 1, 2, USER_HEADERS.length).setValues([
      ['admin', 'ผู้ดูแลระบบ', 'admin', true, now],
      ['emd2', 'emd2', 'staff', true, now],
    ]);
  }
  users.getRange('D2:D').insertCheckboxes();

  const log = ensureSheet_(ss, SHEET.LOG, LOG_HEADERS);
  log.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');

  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  Logger.log('✅ Spreadsheet พร้อมใช้งาน: ' + ss.getUrl());
  return ss.getUrl();
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const first = sh.getRange(1, 1, 1, headers.length);
  if (first.getValues()[0].join('') === '') first.setValues([headers]);
  first.setFontWeight('bold').setBackground('#1f3b57').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new AppError('NOT_SETUP', 'ระบบยังไม่ได้ตั้งค่า (ให้รัน setup() ใน Apps Script editor)');
  return SpreadsheetApp.openById(id);
}

/* ------------------------------------------------------------------ */
/* HTTP entry points                                                  */
/* ------------------------------------------------------------------ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  return handle_(Object.assign({ action: 'list' }, p));
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, code: 'BAD_JSON', error: 'รูปแบบข้อมูลไม่ถูกต้อง' });
  }
  return handle_(body);
}

const ROUTES = {
  list: () => listCompanies_(),
  login: (p) => login_(p),
  me: (p) => ({ user: auth_(p.token) }),
  checkin: (p) => checkin_(p),
  undo: (p) => undo_(p),
  users: (p) => listUsers_(p),
  addUser: (p) => addUser_(p),
  updateUser: (p) => updateUser_(p),
  deleteUser: (p) => deleteUser_(p),
};

function handle_(p) {
  try {
    const fn = ROUTES[p.action];
    if (!fn) throw new AppError('BAD_ACTION', 'ไม่รู้จักคำสั่ง: ' + p.action);
    return json_({ ok: true, data: fn(p) });
  } catch (err) {
    if (err instanceof AppError) return json_({ ok: false, code: err.code, error: err.message });
    console.error(err);
    return json_({ ok: false, code: 'SERVER_ERROR', error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

class AppError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

/* ------------------------------------------------------------------ */
/* Companies                                                          */
/* ------------------------------------------------------------------ */

function readCompanies_() {
  const sh = ss_().getSheetByName(SHEET.COMPANIES);
  const last = sh.getLastRow();
  const values = last < 2 ? [] : sh.getRange(2, 1, last - 1, COMPANY_HEADERS.length).getValues();
  return { sh, values };
}

function toCompany_(r) {
  const time = r[C.TIME];
  return {
    id: String(r[C.ID]).trim(), // PEA No.
    name: String(r[C.NAME]).trim(),
    mt: String(r[C.MT]).trim(),
    date: String(r[C.DATE]).trim(),
    arrived: String(r[C.STATUS]).trim() === STATUS_ARRIVED,
    count: Number(r[C.COUNT]) || 0,
    time: time instanceof Date ? time.toISOString() : (time ? String(time) : ''),
    by: String(r[C.BY]).trim(),
  };
}

function listCompanies_() {
  const { values } = readCompanies_();
  // ใช้เฉพาะแถวที่มีทั้ง PEA No. และชื่อบริษัท
  const companies = values.map(toCompany_).filter((c) => c.id && c.name);
  const arrived = companies.filter((c) => c.arrived);
  return {
    companies,
    summary: {
      total: companies.length,
      arrived: arrived.length,
      remaining: companies.length - arrived.length,
      people: arrived.reduce((s, c) => s + c.count, 0),
    },
    serverTime: new Date().toISOString(),
  };
}

function findCompanyRow_(values, id) {
  const target = String(id || '').trim();
  if (!target) throw new AppError('BAD_INPUT', 'ไม่ได้ระบุบริษัท');
  const i = values.findIndex((r) => String(r[C.ID]).trim() === target);
  if (i < 0) throw new AppError('NOT_FOUND', 'ไม่พบบริษัทนี้ในรายชื่อ (อาจถูกลบหรือแก้ไขใน Sheet)');
  return i;
}

function checkin_(p) {
  const user = auth_(p.token);
  const count = Math.floor(Number(p.count));
  if (!(count >= 1 && count <= 999)) throw new AppError('BAD_INPUT', 'จำนวนคนต้องอยู่ระหว่าง 1–999');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const { sh, values } = readCompanies_();
    const i = findCompanyRow_(values, p.id);
    const current = toCompany_(values[i]);
    if (current.arrived) {
      throw new AppError('ALREADY', `"${current.name}" ลงทะเบียนไปแล้ว${current.by ? ' โดย ' + current.by : ''}`);
    }
    const now = new Date();
    sh.getRange(i + 2, C.STATUS + 1, 1, 4).setValues([[STATUS_ARRIVED, count, now, user.name]]);
    log_('CHECKIN', current, count, user);
  } finally {
    lock.releaseLock();
  }
  return listCompanies_();
}

function undo_(p) {
  const user = auth_(p.token);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const { sh, values } = readCompanies_();
    const i = findCompanyRow_(values, p.id);
    const current = toCompany_(values[i]);
    if (!current.arrived) throw new AppError('NOT_ARRIVED', `"${current.name}" ยังไม่ได้ลงทะเบียน`);
    sh.getRange(i + 2, C.STATUS + 1, 1, 4).clearContent();
    log_('UNDO', current, current.count, user);
  } finally {
    lock.releaseLock();
  }
  return listCompanies_();
}

function log_(action, company, count, user) {
  ss_().getSheetByName(SHEET.LOG).appendRow([new Date(), action, company.id, company.name, count, user.name + ' (' + user.username + ')']);
}

/* ------------------------------------------------------------------ */
/* Users & sessions                                                   */
/* ------------------------------------------------------------------ */

function readUsers_() {
  const sh = ss_().getSheetByName(SHEET.USERS);
  const last = sh.getLastRow();
  const values = last < 2 ? [] : sh.getRange(2, 1, last - 1, USER_HEADERS.length).getValues();
  return { sh, values };
}

function toUser_(r) {
  const active = r[U.ACTIVE];
  return {
    username: String(r[U.USERNAME]).trim(),
    name: String(r[U.NAME]).trim() || String(r[U.USERNAME]).trim(),
    role: String(r[U.ROLE]).trim().toLowerCase() === 'admin' ? 'admin' : 'staff',
    active: active === true || String(active).toUpperCase() === 'TRUE',
  };
}

function findUser_(values, username) {
  const key = String(username || '').trim().toLowerCase();
  return values.findIndex((r) => String(r[U.USERNAME]).trim().toLowerCase() === key);
}

function login_(p) {
  const { values } = readUsers_();
  const i = findUser_(values, p.username);
  if (i < 0 || !String(p.username || '').trim()) throw new AppError('LOGIN_FAILED', 'ไม่พบชื่อผู้ใช้นี้');
  const user = toUser_(values[i]);
  if (!user.active) throw new AppError('LOGIN_FAILED', 'ผู้ใช้นี้ถูกปิดการใช้งาน');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + token, user.username, SESSION_TTL_SEC);
  return { token, user };
}

/** ตรวจ token และสถานะผู้ใช้ล่าสุดใน Sheet (ถ้าถูกปิด/ลบ จะใช้งานต่อไม่ได้ทันที) */
function auth_(token, needAdmin) {
  const username = token && CacheService.getScriptCache().get('s_' + token);
  if (!username) throw new AppError('SESSION_EXPIRED', 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const { values } = readUsers_();
  const i = findUser_(values, username);
  const user = i >= 0 ? toUser_(values[i]) : null;
  if (!user || !user.active) {
    CacheService.getScriptCache().remove('s_' + token);
    throw new AppError('SESSION_EXPIRED', 'บัญชีนี้ถูกปิดการใช้งาน');
  }
  if (needAdmin && user.role !== 'admin') throw new AppError('FORBIDDEN', 'ต้องเป็นผู้ดูแลระบบ');
  return user;
}

function listUsers_(p) {
  auth_(p.token, true);
  const { values } = readUsers_();
  return {
    users: values.map(toUser_).filter((u) => u.username),
    sheetUrl: ss_().getUrl(),
  };
}

function addUser_(p) {
  auth_(p.token, true);
  const username = String(p.username || '').trim();
  const name = String(p.name || '').trim() || username;
  const role = p.role === 'admin' ? 'admin' : 'staff';
  if (!USERNAME_RE.test(username)) {
    throw new AppError('BAD_INPUT', 'Username ใช้ได้เฉพาะ a-z, 0-9, . _ - ความยาว 2–30 ตัว');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const { sh, values } = readUsers_();
    if (findUser_(values, username) >= 0) throw new AppError('DUPLICATE', 'มี Username นี้อยู่แล้ว');
    sh.appendRow([username, name, role, true, new Date()]);
    sh.getRange(sh.getLastRow(), U.ACTIVE + 1).insertCheckboxes().check();
  } finally {
    lock.releaseLock();
  }
  return listUsers_(p);
}

function updateUser_(p) {
  const me = auth_(p.token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const { sh, values } = readUsers_();
    const i = findUser_(values, p.username);
    if (i < 0) throw new AppError('NOT_FOUND', 'ไม่พบผู้ใช้');
    const isSelf = toUser_(values[i]).username.toLowerCase() === me.username.toLowerCase();
    if (isSelf && (p.active === false || p.role === 'staff')) {
      throw new AppError('FORBIDDEN', 'ไม่สามารถปิดหรือลดสิทธิ์บัญชีของตัวเองได้');
    }
    if (typeof p.name === 'string' && p.name.trim()) sh.getRange(i + 2, U.NAME + 1).setValue(p.name.trim());
    if (p.role === 'admin' || p.role === 'staff') sh.getRange(i + 2, U.ROLE + 1).setValue(p.role);
    if (typeof p.active === 'boolean') sh.getRange(i + 2, U.ACTIVE + 1).setValue(p.active);
  } finally {
    lock.releaseLock();
  }
  return listUsers_(p);
}

function deleteUser_(p) {
  const me = auth_(p.token, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const { sh, values } = readUsers_();
    const i = findUser_(values, p.username);
    if (i < 0) throw new AppError('NOT_FOUND', 'ไม่พบผู้ใช้');
    if (toUser_(values[i]).username.toLowerCase() === me.username.toLowerCase()) {
      throw new AppError('FORBIDDEN', 'ไม่สามารถลบบัญชีของตัวเองได้');
    }
    sh.deleteRow(i + 2);
  } finally {
    lock.releaseLock();
  }
  return listUsers_(p);
}
