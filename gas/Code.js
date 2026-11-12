/**
 * Solar People — ระบบลงทะเบียนเข้างาน (Backend API)
 *
 * Frontend (GitHub Pages) เรียก API นี้ผ่าน Web App URL
 *   GET  ?action=list                    → รายชื่อบริษัท + สรุป (ไม่ต้องล็อกอิน)
 *   POST {action:'login', username}      → รับ token
 *   POST {action:'checkin', token, id, count}
 *   POST {action:'undo', token, id}
 *   POST {action:'register', code, id, name, position, email, phone}  ผู้เข้างานลงทะเบียนเอง (ไม่ต้องล็อกอิน)
 *   POST {action:'users' | 'addUser' | 'updateUser' | 'deleteUser', token, ...}  (แอดมินเท่านั้น)
 *
 * ครั้งแรก: รันฟังก์ชัน setup() ใน editor หนึ่งครั้ง เพื่อสร้าง Google Sheet และอนุญาตสิทธิ์
 */

const SHEET = { COMPANIES: 'Companies', USERS: 'Users', LOG: 'Log', REG: 'Registrations' };
const BUILD = 'staff-person-1'; // ใช้ตรวจว่า deployment ใหม่ใช้งานครบทุกเครื่องแล้ว (?action=ping)

const COMPANY_HEADERS = ['PEA No.', 'ชื่อบริษัท', 'เลขที่ มท.', 'ลงวันที่', 'สถานะ', 'จำนวนคน', 'เวลาเข้างาน', 'ผู้บันทึก'];
const C = { ID: 0, NAME: 1, MT: 2, DATE: 3, STATUS: 4, COUNT: 5, TIME: 6, BY: 7 };

const USER_HEADERS = ['Username', 'ชื่อที่แสดง', 'Role', 'Active', 'สร้างเมื่อ'];
const U = { USERNAME: 0, NAME: 1, ROLE: 2, ACTIVE: 3, CREATED: 4 };

const LOG_HEADERS = ['เวลา', 'Action', 'CompanyID', 'ชื่อบริษัท', 'จำนวนคน', 'ผู้ทำรายการ'];

/** ผู้เข้างานที่ลงทะเบียนเอง (1 แถว = 1 คน, เพิ่มต่อท้ายอย่างเดียว ไม่ต้องใช้ lock) */
const REG_HEADERS = ['เวลา', 'PEA No.', 'ชื่อบริษัท', 'ชื่อ-สกุล', 'ตำแหน่ง', 'อีเมล', 'เบอร์โทร', 'ช่องทาง'];
const R = { TIME: 0, ID: 1, COMPANY: 2, NAME: 3, POSITION: 4, EMAIL: 5, PHONE: 6, CHANNEL: 7 };
const SELF_LABEL = 'ลงทะเบียนเอง'; // ช่องทาง: ผู้เข้างานสแกน QR เอง
const STAFF_CHANNEL_PREFIX = 'เจ้าหน้าที่: '; // ช่องทาง: เจ้าหน้าที่กรอกแทน
const PERSON_LABEL = 'ลงทะเบียนรายคน';
const TEST_CHANNELS = ['TEST', 'LOADTEST'];

const STATUS_ARRIVED = 'มาแล้ว';
const SESSION_TTL_SEC = 6 * 60 * 60; // 6 ชั่วโมง (สูงสุดของ CacheService)
const USERNAME_RE = /^[A-Za-z0-9._-]{2,30}$/;

/* ------------------------------------------------------------------ */
/* Setup                                                              */
/* ------------------------------------------------------------------ */

/** Google Sheet ที่ใช้เป็นฐานข้อมูล ("Solar People - ลงทะเบียนเข้างาน" ใน Drive ของเจ้าของสคริปต์) */
const TARGET_SPREADSHEET_ID = '1NKj9w01h5orGZ52fBKlm7zbon1O0fjLZgVklt_FUQ4E';

/**
 * เชื่อมระบบเข้ากับ Sheet ตาม TARGET_SPREADSHEET_ID แล้วสร้างแท็บ/ใส่รายชื่อให้
 * (รันใน editor หนึ่งครั้ง — ไม่ลบข้อมูลเดิมในแท็บอื่นของ Sheet)
 */
function connectSheet() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); // ถ้าไม่มีสิทธิ์แก้ไข จะ error ตรงนี้
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  return setup();
}

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

  regSheet_();

  // ลบแท็บว่างที่ Google สร้างให้อัตโนมัติ (เฉพาะกรณีว่างจริง ๆ)
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0 && defaultSheet.getLastColumn() === 0) {
    ss.deleteSheet(defaultSheet);
  }

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

let SS_CACHE_ = null;
/** เปิด Spreadsheet ครั้งเดียวต่อ request */
function ss_() {
  if (SS_CACHE_) return SS_CACHE_;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new AppError('NOT_SETUP', 'ระบบยังไม่ได้ตั้งค่า (ให้รัน setup() ใน Apps Script editor)');
  SS_CACHE_ = SpreadsheetApp.openById(id);
  return SS_CACHE_;
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
  list: (p) => compat_(p, listCompanies_(p.v)),
  login: (p) => login_(p),
  me: (p) => ({ user: auth_(p.token) }),
  checkin: (p) => compat_(p, checkin_(p)),
  undo: (p) => compat_(p, undo_(p)),
  register: (p) => register_(p),
  staffRegister: (p) => staffRegister_(p),
  companyPeople: (p) => companyPeople_(p),
  removePerson: (p) => removePerson_(p),
  ping: () => ({ build: BUILD }),
  getEventCode: (p) => { auth_(p.token, true); return { code: eventCode_() }; },
  setEventCode: (p) => setEventCode_(p),
  clearTest: (p) => {
    auth_(p.token, true);
    // ล้าง cache กันซ้ำของเบอร์ทดสอบที่ระบุ (เผื่อแถวทดสอบหายไปจาก Sheet แล้ว)
    if (Array.isArray(p.phones) && p.phones.length) {
      CacheService.getScriptCache().removeAll(p.phones.slice(0, 1000).map((x) => 'ph_' + normPhone_(x)));
    }
    return clearTestRegistrations();
  },
  users: (p) => listUsers_(p),
  addUser: (p) => addUser_(p),
  updateUser: (p) => updateUser_(p),
  deleteUser: (p) => deleteUser_(p),
};

/**
 * หน้าเว็บรุ่นใหม่ส่ง f=2 มาและรับรูปแบบย่อ
 * หน้าเว็บรุ่นเก่าที่ยังค้างอยู่ใน cache ของเบราว์เซอร์ (ไม่มี f) ได้รูปแบบเดิม จะได้ไม่พัง
 */
function compat_(p, r) {
  if (String(p.f) === '2') return r;
  const d = r.d || listCompanies_().d; // หน้าเก่าไม่ส่ง v จึงไม่ได้คำตอบแบบ same อยู่แล้ว
  return {
    companies: d.c.map((x) => ({
      id: x[0], name: x[1], mt: x[2], arrived: !!x[3], count: x[4],
      time: x[5] ? new Date(x[5]).toISOString() : '', by: x[6],
    })),
    summary: d.s,
    serverTime: new Date().toISOString(),
  };
}

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
    arrived: String(r[C.STATUS]).trim() === STATUS_ARRIVED,
    count: Number(r[C.COUNT]) || 0,
    time: time instanceof Date ? time.getTime() : 0,
    by: String(r[C.BY]).trim(),
  };
}

/* ---------- Cache รายชื่อ (ลดการอ่าน Sheet) ----------
 * เก็บ JSON แบบย่อใน CacheService แบ่งเป็นชิ้น (ค่าละไม่เกิน 100KB)
 * - เช็คอิน/ยกเลิก: เขียนทับ cache ทันที (write-through)
 * - ลงทะเบียนเอง / แก้ไขใน Sheet โดยตรง: เห็นผลภายใน LIST_TTL_SEC วินาที
 */
const LIST_TTL_SEC = 5; // การลงทะเบียนเองไม่ล้าง cache ทุกครั้ง จึงเห็นผลภายใน ~5 วินาที
const CHUNK = 90000;

/**
 * แปลงเป็นรูปแบบย่อ:
 * c = [[id, name, mt, arrived(0/1), count, timeMs, by, staff(0/1), self], ...]
 * - staff = เจ้าหน้าที่เช็คอินให้ (แท็บ Companies), self = จำนวนคนที่ลงทะเบียนเอง (แท็บ Registrations)
 * - ไม่มีข้อมูลส่วนบุคคลของผู้ลงทะเบียนในผลลัพธ์
 */
function buildList_(values, agg) {
  agg = agg || regAgg_();
  const rows = [];
  let arrived = 0, people = 0;
  values.forEach((r) => {
    const c = toCompany_(r);
    if (!c.id || !c.name) return; // ใช้เฉพาะแถวที่มีทั้ง PEA No. และชื่อบริษัท
    const reg = agg.by[c.id] || { n: 0, first: 0 };
    const isArrived = c.arrived || reg.n > 0;
    const count = (c.arrived ? c.count : 0) + reg.n;
    const times = [c.arrived ? c.time : 0, reg.first].filter(Boolean);
    const by = c.arrived ? (reg.n ? `${c.by} + ${PERSON_LABEL} ${reg.n}` : c.by) : (reg.n ? PERSON_LABEL : '');
    if (isArrived) { arrived++; people += count; }
    rows.push([c.id, c.name, c.mt, isArrived ? 1 : 0, count, times.length ? Math.min(...times) : 0, by, c.arrived ? 1 : 0, reg.n]);
  });
  const json = JSON.stringify({
    c: rows,
    s: { total: rows.length, arrived, remaining: rows.length - arrived, people, self: agg.total },
  });
  const v = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, json, Utilities.Charset.UTF_8)).slice(0, 12);
  return { v, json };
}

function cacheList_(list) {
  const cache = CacheService.getScriptCache();
  const parts = {};
  const n = Math.ceil(list.json.length / CHUNK) || 1;
  for (let i = 0; i < n; i++) parts['list_' + list.v + '_' + i] = list.json.slice(i * CHUNK, (i + 1) * CHUNK);
  cache.putAll(parts, LIST_TTL_SEC + 30);
  cache.put('list_head', JSON.stringify({ v: list.v, n }), LIST_TTL_SEC);
}

function cachedList_() {
  const cache = CacheService.getScriptCache();
  const head = cache.get('list_head');
  if (!head) return null;
  const { v, n } = JSON.parse(head);
  return { v, n, getJson: () => {
    const keys = Array.from({ length: n }, (_, i) => 'list_' + v + '_' + i);
    const got = cache.getAll(keys);
    if (keys.some((k) => got[k] == null)) return null;
    return keys.map((k) => got[k]).join('');
  } };
}

/** ถ้าลูกค้าส่ง v ที่ตรงกับข้อมูลล่าสุดมา ตอบสั้น ๆ ว่าไม่มีการเปลี่ยนแปลง */
function listCompanies_(clientV) {
  const hit = cachedList_();
  if (hit) {
    if (clientV && clientV === hit.v) return { v: hit.v, same: true };
    const json = hit.getJson();
    if (json) return { v: hit.v, d: JSON.parse(json) };
  }
  const list = buildList_(readCompanies_().values, regAgg_());
  cacheList_(list);
  if (clientV && clientV === list.v) return { v: list.v, same: true };
  return { v: list.v, d: JSON.parse(list.json) };
}

function findCompanyRow_(values, id) {
  const target = String(id || '').trim();
  if (!target) throw new AppError('BAD_INPUT', 'ไม่ได้ระบุบริษัท');
  const i = values.findIndex((r) => String(r[C.ID]).trim() === target);
  if (i < 0) throw new AppError('NOT_FOUND', 'ไม่พบบริษัทนี้ในรายชื่อ (อาจถูกลบหรือแก้ไขใน Sheet)');
  return i;
}

/** อ่าน-แก้-เขียน ภายใต้ lock แล้วอัปเดต cache จากข้อมูลในหน่วยความจำ (ไม่อ่าน Sheet ซ้ำ) */
function mutateCompany_(id, fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const { sh, values } = readCompanies_();
    const i = findCompanyRow_(values, id);
    const row = fn(toCompany_(values[i])); // [สถานะ, จำนวน, เวลา, ผู้บันทึก]
    sh.getRange(i + 2, C.STATUS + 1, 1, 4).setValues([row]);
    SpreadsheetApp.flush();
    values[i][C.STATUS] = row[0];
    values[i][C.COUNT] = row[1];
    values[i][C.TIME] = row[2];
    values[i][C.BY] = row[3];
    const list = buildList_(values);
    cacheList_(list);
    return { v: list.v, d: JSON.parse(list.json) };
  } finally {
    lock.releaseLock();
  }
}

function checkin_(p) {
  const user = auth_(p.token);
  const count = Math.floor(Number(p.count));
  if (!(count >= 1 && count <= 999)) throw new AppError('BAD_INPUT', 'จำนวนคนต้องอยู่ระหว่าง 1–999');
  return mutateCompany_(p.id, (current) => {
    if (current.arrived) {
      throw new AppError('ALREADY', `"${current.name}" ลงทะเบียนไปแล้ว${current.by ? ' โดย ' + current.by : ''}`);
    }
    log_('CHECKIN', current, count, user);
    return [STATUS_ARRIVED, count, new Date(), user.name];
  });
}

function undo_(p) {
  const user = auth_(p.token);
  return mutateCompany_(p.id, (current) => {
    if (!current.arrived) throw new AppError('NOT_ARRIVED', `"${current.name}" ยังไม่ได้ลงทะเบียน`);
    log_('UNDO', current, current.count, user);
    return ['', '', '', ''];
  });
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

/* ------------------------------------------------------------------ */
/* ผู้เข้างานลงทะเบียนเอง                                                */
/* ------------------------------------------------------------------ */

const PHONE_TTL_SEC = 6 * 60 * 60;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

/** แท็บ Registrations (สร้างให้อัตโนมัติถ้ายังไม่มี) */
function regSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET.REG);
  if (sh) return sh;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    sh = ss.getSheetByName(SHEET.REG) || ensureSheet_(ss, SHEET.REG, REG_HEADERS);
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sh.getRange('G:G').setNumberFormat('@');
    return sh;
  } finally {
    lock.releaseLock();
  }
}

/** เหลือแต่ตัวเลข และเติม 0 นำหน้าที่ Sheet อาจตัดทิ้ง (เช่น 812345678 → 0812345678) */
function normPhone_(v) {
  let d = String(v == null ? '' : v).replace(/\D/g, '');
  if (d.startsWith('66') && d.length === 11) d = '0' + d.slice(2);
  if (d.length === 9 && d[0] !== '0') d = '0' + d;
  return d;
}

/** สรุปจำนวนคนที่ลงทะเบียนเองต่อบริษัท + เก็บรายการเบอร์ไว้กันซ้ำ */
function regAgg_() {
  const sh = ss_().getSheetByName(SHEET.REG);
  const by = {};
  const phones = [];
  let total = 0;
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, REG_HEADERS.length).getValues().forEach((r) => {
      const id = String(r[R.ID]).trim();
      if (!id) return;
      const t = r[R.TIME] instanceof Date ? r[R.TIME].getTime() : 0;
      const a = by[id] || (by[id] = { n: 0, first: 0 });
      a.n++;
      total++;
      if (t && (!a.first || t < a.first)) a.first = t;
      const ph = normPhone_(r[R.PHONE]);
      if (ph) phones.push(ph);
    });
  }
  CacheService.getScriptCache().put('reg_phones', JSON.stringify(phones), 60);
  return { by, phones, total };
}

function eventCode_() {
  const props = PropertiesService.getScriptProperties();
  let code = props.getProperty('EVENT_CODE');
  if (!code) {
    code = Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
    props.setProperty('EVENT_CODE', code);
  }
  return code;
}

function setEventCode_(p) {
  auth_(p.token, true);
  const code = String(p.code || '').trim().toUpperCase();
  if (!/^[A-Z0-9]{4,20}$/.test(code)) throw new AppError('BAD_INPUT', 'รหัสงานใช้ A-Z และ 0-9 ความยาว 4–20 ตัว');
  PropertiesService.getScriptProperties().setProperty('EVENT_CODE', code);
  return { code };
}

function register_(p) {
  const code = String(p.code || '').trim().toUpperCase();
  if (!code || code !== eventCode_()) {
    throw new AppError('BAD_CODE', 'รหัสงานไม่ถูกต้อง กรุณาสแกน QR Code ที่จุดลงทะเบียนหน้างาน');
  }
  const channel = TEST_CHANNELS.indexOf(p.channel) >= 0 ? p.channel : SELF_LABEL;
  return addPerson_(p, { strict: true, channel });
}

/** เจ้าหน้าที่ลงทะเบียนรายคนแทนผู้เข้างาน (บังคับแค่ชื่อและเบอร์) */
function staffRegister_(p) {
  const user = auth_(p.token);
  const channel = TEST_CHANNELS.indexOf(p.channel) >= 0 ? p.channel : STAFF_CHANNEL_PREFIX + user.name;
  const r = addPerson_(p, { strict: false, channel });
  CacheService.getScriptCache().remove('list_head'); // ให้หน้าเจ้าหน้าที่/แดชบอร์ดเห็นทันที
  log_('ADD_PERSON', { id: r.id, name: r.company }, 1, user);
  return Object.assign(r, companyPeople_(p, user));
}

/**
 * บันทึกผู้เข้างาน 1 คนลงแท็บ Registrations
 * strict = true (ผู้เข้างานกรอกเอง): บังคับ ชื่อ ตำแหน่ง อีเมล เบอร์
 * strict = false (เจ้าหน้าที่กรอกแทน): บังคับ ชื่อ เบอร์ ส่วนตำแหน่ง/อีเมล ไม่บังคับ
 */
function addPerson_(p, opt) {
  const name = String(p.name || '').trim().replace(/\s+/g, ' ');
  const position = String(p.position || '').trim().replace(/\s+/g, ' ');
  const email = String(p.email || '').trim().toLowerCase();
  const phone = normPhone_(p.phone);
  if (name.length < 2 || name.length > 100) throw new AppError('BAD_INPUT', 'กรุณากรอกชื่อ-สกุล');
  if (position.length > 100 || (opt.strict && !position)) throw new AppError('BAD_INPUT', 'กรุณากรอกตำแหน่ง');
  if (email.length > 120 || ((opt.strict || email) && !EMAIL_RE.test(email))) throw new AppError('BAD_INPUT', 'รูปแบบอีเมลไม่ถูกต้อง');
  if (!/^0\d{8,9}$/.test(phone)) throw new AppError('BAD_INPUT', 'เบอร์โทรศัพท์ต้องเป็นตัวเลข 9–10 หลัก');

  // ตรวจบริษัทจากรายชื่อที่ cache ไว้ (ไม่ต้องอ่าน Sheet ถ้า cache ยังอยู่)
  const list = listCompanies_();
  const row = list.d.c.find((x) => x[0] === String(p.id || '').trim());
  if (!row) throw new AppError('NOT_FOUND', 'ไม่พบบริษัทนี้ในรายชื่อ กรุณาติดต่อโต๊ะลงทะเบียน');

  const cache = CacheService.getScriptCache();
  const dupMsg = (company) => `เบอร์นี้ลงทะเบียนแล้ว${company ? ` (${company})` : ''} หากต้องการแก้ไข กรุณาติดต่อโต๊ะลงทะเบียน`;
  // เช็คซ้ำรอบแรกนอก lock (ตอบเร็ว ไม่ต้องรอคิว)
  if (cache.get('ph_' + phone)) throw new AppError('DUPLICATE', dupMsg(JSON.parse(cache.get('ph_' + phone) || '{}').company));
  const phones = JSON.parse(cache.get('reg_phones') || 'null') || regAgg_().phones;
  if (phones.indexOf(phone) >= 0) throw new AppError('DUPLICATE', dupMsg());

  const sh = regSheet_();
  // appendRow พร้อมกันหลาย request เขียนทับแถวเดียวกันได้ (ข้อมูลหาย — พบจาก load test)
  // จึงล็อกเฉพาะช่วงเขียนแถว ซึ่งสั้นมาก (ไม่อ่านทั้งชีตเหมือนการเช็คอินของเจ้าหน้าที่)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new AppError('BUSY', 'ขณะนี้มีผู้ลงทะเบียนพร้อมกันจำนวนมาก กรุณากดลงทะเบียนอีกครั้ง');
  }
  let now;
  try {
    // เช็คซ้ำอีกครั้งภายใน lock กันกดส่งซ้ำพร้อมกัน
    if (cache.get('ph_' + phone)) throw new AppError('DUPLICATE', dupMsg(JSON.parse(cache.get('ph_' + phone) || '{}').company));
    now = new Date();
    // ใส่ ' นำหน้าเบอร์ เพื่อให้ Sheet เก็บเป็นข้อความ (ไม่ตัด 0 ข้างหน้า)
    sh.appendRow([now, row[0], row[1], name, position, email, "'" + phone, opt.channel]);
    SpreadsheetApp.flush();
    cache.put('ph_' + phone, JSON.stringify({ company: row[1], t: now.getTime() }), PHONE_TTL_SEC);
  } finally {
    lock.releaseLock();
  }
  return { time: now.getTime(), company: row[1], id: row[0], name };
}

/** รายชื่อผู้เข้างานของบริษัท (เฉพาะเจ้าหน้าที่ที่ล็อกอิน) */
function companyPeople_(p, user) {
  if (!user) auth_(p.token);
  const id = String(p.id || '').trim();
  const sh = ss_().getSheetByName(SHEET.REG);
  const people = [];
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, REG_HEADERS.length).getValues().forEach((r) => {
      if (String(r[R.ID]).trim() !== id) return;
      people.push({
        name: String(r[R.NAME]),
        position: String(r[R.POSITION]),
        email: String(r[R.EMAIL]),
        phone: normPhone_(r[R.PHONE]),
        channel: String(r[R.CHANNEL]),
        time: r[R.TIME] instanceof Date ? r[R.TIME].getTime() : 0,
      });
    });
  }
  return { people };
}

/** ลบผู้เข้างาน 1 คน (ระบุด้วยบริษัท + เบอร์) กรณีกรอกผิด */
function removePerson_(p) {
  const user = auth_(p.token);
  const id = String(p.id || '').trim();
  const phone = normPhone_(p.phone);
  const sh = regSheet_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new AppError('BUSY', 'ระบบไม่ว่าง กรุณาลองอีกครั้ง');
  let removed;
  try {
    const last = sh.getLastRow();
    const values = last < 2 ? [] : sh.getRange(2, 1, last - 1, REG_HEADERS.length).getValues();
    const i = values.findIndex((r) => String(r[R.ID]).trim() === id && normPhone_(r[R.PHONE]) === phone);
    if (i < 0) throw new AppError('NOT_FOUND', 'ไม่พบผู้เข้างานคนนี้ (อาจถูกลบไปแล้ว)');
    removed = values[i];
    sh.deleteRow(i + 2);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  const cache = CacheService.getScriptCache();
  cache.removeAll(['ph_' + phone, 'list_head', 'reg_phones']);
  log_('REMOVE_PERSON', { id, name: String(removed[R.COMPANY]) }, 1, user);
  return companyPeople_(p, user);
}

/** ลบข้อมูลทดสอบ (ช่องทาง TEST / LOADTEST) ออกจากแท็บ Registrations */
function clearTestRegistrations() {
  const sh = regSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const last = sh.getLastRow();
    if (last < 2) return { removed: 0, kept: 0 };
    const values = sh.getRange(2, 1, last - 1, REG_HEADERS.length).getValues();
    const isTest = (r) => TEST_CHANNELS.indexOf(String(r[R.CHANNEL])) >= 0;
    const keep = values.filter((r) => !isTest(r));
    const removedPhones = values.filter(isTest).map((r) => 'ph_' + normPhone_(r[R.PHONE]));
    sh.getRange(2, 1, last - 1, REG_HEADERS.length).clearContent();
    if (keep.length) {
      keep.forEach((r) => { r[R.PHONE] = "'" + normPhone_(r[R.PHONE]); });
      sh.getRange(2, 1, keep.length, REG_HEADERS.length).setValues(keep);
    }
    const cache = CacheService.getScriptCache();
    if (removedPhones.length) cache.removeAll(removedPhones);
    cache.remove('list_head');
    cache.remove('reg_phones');
    return { removed: values.length - keep.length, kept: keep.length };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Cache รายชื่อผู้ใช้ ----------
 * แก้ผ่านหน้า admin: มีผลทันที / แก้ในแท็บ Users โดยตรง: มีผลภายใน 2 นาที
 */
const USERS_TTL_SEC = 120;

function userValues_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('users');
  if (hit) return JSON.parse(hit);
  const values = readUsers_().values.map((r) => [r[U.USERNAME], r[U.NAME], r[U.ROLE], r[U.ACTIVE]]);
  cache.put('users', JSON.stringify(values), USERS_TTL_SEC);
  return values;
}

function invalidateUsers_() {
  CacheService.getScriptCache().remove('users');
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
  const values = userValues_();
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
  const values = userValues_();
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
    invalidateUsers_();
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
    invalidateUsers_();
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
    invalidateUsers_();
  } finally {
    lock.releaseLock();
  }
  return listUsers_(p);
}
