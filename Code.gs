// Code.gs: หน้าเว็บ + ฟอร์มติดต่อ + ระบบสมัคร/เข้าสู่ระบบ
// Deploy: Execute as = Me, Who has access = Anyone
const ITER = 500;            // จำนวนรอบแฮชรหัสผ่าน
const SESSION_SECS = 21600;  // อายุ session 6 ชั่วโมง (ค่าสูงสุดของ CacheService)

// ตั้งค่าแอป: แก้ 2 ค่านี้
const APP_NAME = 'แอปของฉัน';
const ICON_FILE_ID = '';  // ID ไฟล์ icon-512.png ที่อัปโหลดไว้ใน Drive (แชร์แบบ Anyone with the link)

function doGet() {
  const out = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (ICON_FILE_ID) out.setFaviconUrl('https://drive.google.com/uc?export=view&id=' + ICON_FILE_ID);
  return out;
}

// ---------- Spreadsheet ----------
function ss_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('Contacts');
  ss.getSheets()[0].appendRow(['เวลา', 'ชื่อ', 'อีเมล', 'ข้อความ']);
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}
function usersSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName('Users');
  if (!sh) { sh = ss.insertSheet('Users'); sh.appendRow(['สร้างเมื่อ', 'อีเมล', 'ชื่อ', 'salt', 'hash']); }
  return sh;
}

function saveContact(name, email, message) {
  ss_().getSheets()[0].appendRow([new Date(), name, email, message]);
  return true;
}

// ---------- Auth ----------
function hash_(pw, salt) {
  let h = salt + pw;
  for (let i = 0; i < ITER; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  }
  return h;
}
function findUser_(email) {
  const rows = usersSheet_().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === email) return { email: rows[i][1], name: rows[i][2], salt: rows[i][3], hash: rows[i][4] };
  }
  return null;
}
function newSession_(u) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('s:' + token, u.email, SESSION_SECS);
  return { token: token, name: u.name, email: u.email };
}
function requireUser_(token) {
  const email = token && CacheService.getScriptCache().get('s:' + token);
  const u = email && findUser_(email);
  if (!u) throw new Error('หมดเวลา กรุณาเข้าสู่ระบบใหม่');
  return u;
}

function register(name, email, password) {
  name = String(name).trim(); email = String(email).trim().toLowerCase(); password = String(password);
  if (!name) throw new Error('กรุณากรอกชื่อ');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('อีเมลไม่ถูกต้อง');
  if (password.length < 8) throw new Error('รหัสผ่านต้องมีอย่างน้อย 8 ตัว');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (findUser_(email)) throw new Error('อีเมลนี้สมัครแล้ว');
    const salt = Utilities.getUuid();
    usersSheet_().appendRow([new Date(), email, name, salt, hash_(password, salt)]);
    return newSession_({ email: email, name: name });
  } finally { lock.releaseLock(); }
}

function login(email, password) {
  email = String(email).trim().toLowerCase();
  const cache = CacheService.getScriptCache(), fk = 'f:' + email;
  if (Number(cache.get(fk) || 0) >= 5) throw new Error('ลองผิดหลายครั้ง กรุณารอ 10 นาที');
  const u = findUser_(email);
  if (!u || hash_(String(password), u.salt) !== u.hash) {
    cache.put(fk, String(Number(cache.get(fk) || 0) + 1), 600);
    throw new Error('อีเมลหรือรหัสผ่านไม่ถูกต้อง');
  }
  cache.remove(fk);
  return newSession_(u);
}

function getMe(token) {
  const email = token && CacheService.getScriptCache().get('s:' + token);
  const u = email && findUser_(email);
  return u ? { name: u.name, email: u.email } : null;
}
function logout(token) { if (token) CacheService.getScriptCache().remove('s:' + token); return true; }

// ตัวอย่างฟังก์ชันที่ต้องล็อกอินก่อน: เรียก requireUser_(token) ทุกครั้งที่ส่งข้อมูลส่วนตัว
function getMyData(token) {
  const u = requireUser_(token);
  return { message: 'ข้อมูลนี้ส่งให้เฉพาะ ' + u.email + ' เท่านั้น' };
}
