/**
 * ════════════════════════════════════════════════════════════════════
 *  ระบบส่งงานข้อมูลวิจัย · RESEARCH UPLINK — Google Apps Script Backend  (API v2)
 * ════════════════════════════════════════════════════════════════════
 *
 *  รองรับการ "ทยอยส่งทีละบท"
 *  ─────────────────────────
 *  • ส่งครั้งแรก  → แนบกี่บทก็ได้ (อย่างน้อย 1 ไฟล์) ระบบออก "รหัสโครงการ" เช่น RS-261106-1432-K7QX
 *  • ส่งครั้งถัดไป → ใช้รหัสโครงการเดิม (หรือค้นหาด้วยชื่อ–สกุล) ส่งบทที่เหลือ หรือส่งฉบับแก้ไขทับบทเดิม
 *  • 1 โครงการ = 1 แถวในชีต "รายการส่งงาน" (อัปเดตแถวเดิม) + ทุกครั้งที่ส่งบันทึกในชีต "ประวัติการส่ง"
 *  • ฉบับเก่าไม่ถูกลบ — ย้ายไปโฟลเดอร์ย่อย "ฉบับก่อนหน้า" และตั้งชื่อ _v1, _v2 ...
 *
 *  วิธีติดตั้ง / อัปเดต
 *  ──────────────────
 *  1) script.google.com → เปิดโปรเจกต์ → วางโค้ดนี้แทนของเดิมทั้งหมดใน Code.gs → Save
 *  2) เลือกฟังก์ชัน  setup  → Run (รันซ้ำได้ ข้อมูลเดิมไม่หาย — จะเพิ่มคอลัมน์/ชีตใหม่ให้)
 *  3) ครั้งแรก : Deploy → New deployment → Web app (Execute as: Me · Who has access: Anyone)
 *     อัปเดต  : Deploy → Manage deployments → ✎ Edit → Version: "New version" → Deploy  (URL เดิมใช้ต่อได้)
 *  4) วาง Web app URL (ลงท้าย /exec) ที่ CONFIG.GAS_URL ใน index.html
 *
 *  คำสั่งที่หน้าเว็บเรียก (POST, body = JSON)
 *  ──────────────────────────────────────
 *    lookup → ค้นหาโครงการด้วย code หรือ firstName + lastName
 *    begin  → เริ่มการส่ง 1 ครั้ง (โครงการใหม่ หรือ projectId เดิม) → submissionId, folderId, batch
 *    upload → รับไฟล์ทีละช่อง (base64) — ฉบับเดิมของช่องนั้นย้ายไป "ฉบับก่อนหน้า"
 *    finish → ตรวจไฟล์ + อัปเดตแถวโครงการ + บันทึกประวัติ → คืนใบรับ
 *    update → (Dashboard · ต้องใช้ ADMIN_KEY) แก้ไขข้อมูลโครงการ / สถานะการตรวจ / หมายเหตุ
 *    delete → (Dashboard · ต้องใช้ ADMIN_KEY) ลบไฟล์รายบท หรือลบทั้งโครงการ (ย้ายไปถังขยะ Drive)
 *  GET ?action=list → ข้อมูล Dashboard (+ id ไฟล์ หมายเหตุ ประวัติ เมื่อ DASHBOARD_DETAILS = true)
 *
 *  ปุ่ม "แก้ไข / ลบ" ใน Dashboard: เข้าสู่ระบบด้วย CONFIG.ADMIN_USER + ADMIN_KEY (เปลี่ยนแล้ว Deploy เวอร์ชันใหม่)
 * ════════════════════════════════════════════════════════════════════
 */

/* ─────────────────────────── การตั้งค่า ─────────────────────────── */

const CONFIG = {
  ROOT_FOLDER_ID: '1Ek01eP6x15VdhWcEzzsnxlqyvaDV10D_',   // โฟลเดอร์ Google Drive ปลายทาง

  SPREADSHEET_NAME: 'ระบบส่งงานข้อมูลวิจัย · ฐานข้อมูล',
  SHEET_SUBMISSIONS: 'รายการส่งงาน',
  SHEET_HISTORY: 'ประวัติการส่ง',
  SHEET_LOGS: 'บันทึกระบบ',
  ARCHIVE_FOLDER: 'ฉบับก่อนหน้า',

  TIMEZONE: 'Asia/Bangkok',
  MAX_FILE_MB: 20,                       // ต้องตรงกับ CONFIG.MAX_FILE_MB ใน index.html
  ALLOWED_EXT: ['pdf', 'doc', 'docx'],
  API_VERSION: 3,

  // รหัสผ่านหน้า Dashboard (ว่าง = เปิดดูได้ทุกคนที่มีลิงก์) — ถ้าตั้ง ให้ใส่ค่าเดียวกันที่ CONFIG.DASHBOARD_KEY ใน dashboard.html
  DASHBOARD_KEY: '',

  // หน้างานวิจัยรายโครงการใน Dashboard: ส่ง id ไฟล์ของแต่ละบท + หมายเหตุ + ประวัติการส่ง ให้เปิดอ่านไฟล์ในหน้าเว็บได้
  // (false = Dashboard แสดงแค่ความคืบหน้า ไม่มีลิงก์ไฟล์/หมายเหตุ)
  // ผู้อ่านต้องมีสิทธิ์ดูไฟล์ใน Drive: บัญชีเจ้าของ/ผู้ที่ได้รับแชร์โฟลเดอร์ หรือแชร์โฟลเดอร์แบบ "ทุกคนที่มีลิงก์" หรือเปิด SHARE_FILES_WITH_LINK
  DASHBOARD_DETAILS: true,

  // true = ไฟล์ที่ส่งเข้ามาตั้งเป็น "ทุกคนที่มีลิงก์ดูได้" อัตโนมัติ → ทุกคนที่เปิด Dashboard อ่านไฟล์ได้โดยไม่ต้องล็อกอิน
  //        ไฟล์ที่ส่งมาก่อนเปิดตัวเลือกนี้: รันฟังก์ชัน shareAllFiles() 1 ครั้ง · แนะนำให้ตั้ง DASHBOARD_KEY ด้วย
  SHARE_FILES_WITH_LINK: false,

  // บัญชีผู้ดูแล สำหรับปุ่ม "แก้ไข" และ "ลบ" ใน Dashboard (ADMIN_KEY ว่าง = ปิดการแก้ไข/ลบทั้งหมด)
  // กรอกในหน้า Dashboard ตอนกดปุ่ม "ผู้ดูแล" (ไม่ต้องใส่ใน HTML)
  // ⚠ '1234' เดาง่าย — ใครเปิด Dashboard ได้ก็ลองเดาได้ ควรเปลี่ยนเป็นรหัสยาว ๆ ก่อนใช้งานจริง แล้ว Deploy เวอร์ชันใหม่
  ADMIN_USER: 'admin',
  ADMIN_KEY: '1234',
};

// ช่องเอกสาร — key ต้องตรงกับ SLOTS ใน index.html (required = นับรวมเพื่อ "ครบถ้วน" ของโครงการ)
const SLOTS = [
  { key: 'title', label: 'ชื่อเรื่อง',     required: true },
  { key: 'ch1',   label: 'บทที่ 1',        required: true },
  { key: 'ch2',   label: 'บทที่ 2',        required: true },
  { key: 'ch3',   label: 'บทที่ 3',        required: true },
  { key: 'ch4',   label: 'บทที่ 4',        required: true },
  { key: 'ch5',   label: 'บทที่ 5',        required: true },
  { key: 'ref',   label: 'เอกสารอ้างอิง', required: true },
];

const MIME = {
  pdf:  'application/pdf',
  doc:  'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};

const REVIEW_STATUSES = ['รอตรวจ', 'กำลังตรวจ', 'ส่งกลับแก้ไข', 'ผ่าน'];
const ID_RE = /^RS-\d{6}-\d{4}-[A-Z0-9]{4}$/;
const DT_FORMAT = 'dd/mm/yyyy hh:mm:ss';

// ลำดับ 18 คอลัมน์แรกเหมือนเวอร์ชันก่อน → แถวข้อมูลเดิมใช้ต่อได้
const HEADERS = [].concat(
  ['รหัสโครงการ', 'ส่งครั้งแรก', 'ชื่อเรื่อง', 'ชื่อ', 'สกุล', 'สถานที่ทำงาน'],
  SLOTS.map(function (s) { return 'ไฟล์ · ' + s.label; }),
  ['โฟลเดอร์งาน', 'จำนวนไฟล์', 'ความครบถ้วน', 'สถานะการตรวจ', 'หมายเหตุ',
   'รูปแบบงานวิจัย', 'อัปเดตล่าสุด', 'จำนวนครั้งที่ส่ง', 'ข้อมูลระบบ']
);
const HISTORY_HEADERS = ['เวลา', 'รหัสโครงการ', 'ครั้งที่', 'ชื่อ–สกุล', 'ประเภทการส่ง', 'เอกสารที่ส่ง (คลิกเปิดไฟล์)', 'ชื่อไฟล์ต้นฉบับ'];
const LOG_HEADERS = ['เวลา', 'ระดับ', 'การทำงาน', 'รหัสโครงการ', 'รายละเอียด'];

function col_(name) {
  const i = HEADERS.indexOf(name);
  if (i < 0) throw new Error('ไม่พบคอลัมน์: ' + name);
  return i + 1;
}


/* ═══════════════════════════ SETUP ═══════════════════════════ */

/**
 * รันก่อน Deploy (และทุกครั้งหลังอัปเดตโค้ด) — รันซ้ำได้ ข้อมูลเดิมไม่หาย
 * - ตรวจสิทธิ์โฟลเดอร์ Google Drive
 * - สร้าง/เปิด Google Sheet + ชีต "รายการส่งงาน", "ประวัติการส่ง", "บันทึกระบบ"
 * - จัดรูปแบบหัวตาราง ความกว้าง Dropdown สถานะ สีตามเงื่อนไข และซ่อนคอลัมน์ข้อมูลระบบ
 */
function setup() {
  const props = PropertiesService.getScriptProperties();

  let folder;
  try {
    folder = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    folder.getName();
  } catch (err) {
    throw new Error('เข้าถึงโฟลเดอร์ Google Drive ไม่ได้ (' + CONFIG.ROOT_FOLDER_ID + ') — '
      + 'ตรวจสอบว่าบัญชีที่รันสคริปต์มีสิทธิ์ "ผู้แก้ไข" ในโฟลเดอร์นี้');
  }

  const ss = openOrCreateSpreadsheet_(props);
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  styleSubmissionsSheet_(ensureSheet_(ss, CONFIG.SHEET_SUBMISSIONS, 0));
  styleHistorySheet_(ensureSheet_(ss, CONFIG.SHEET_HISTORY, 1));
  styleLogSheet_(ensureSheet_(ss, CONFIG.SHEET_LOGS, 2));
  removeEmptyDefaultSheets_(ss);
  props.setProperty('SPREADSHEET_ID', ss.getId());

  log_('INFO', 'setup', '-', 'ติดตั้ง/อัปเดตระบบสำเร็จ (API v' + CONFIG.API_VERSION + ')');

  const msg = [
    '✅ ติดตั้งระบบส่งงานข้อมูลวิจัยสำเร็จ (API v' + CONFIG.API_VERSION + ' · รองรับการทยอยส่งทีละบท)',
    '📊 Google Sheet : ' + ss.getUrl(),
    '📁 โฟลเดอร์ไฟล์ : ' + folder.getName() + ' — ' + folder.getUrl(),
    '➡️  ขั้นต่อไป    : Deploy → Manage deployments → Edit → New version (หรือ New deployment ถ้ายังไม่เคย)',
  ].join('\n');
  Logger.log(msg);
  return msg;
}

function openOrCreateSpreadsheet_(props) {
  const savedId = props.getProperty('SPREADSHEET_ID');
  if (savedId) {
    try {
      const ss = SpreadsheetApp.openById(savedId);
      if (!DriveApp.getFileById(savedId).isTrashed()) return ss;
    } catch (e) { /* ไฟล์เดิมหาย — สร้างใหม่ */ }
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();   // กรณีสคริปต์ผูกกับ Google Sheet
  if (active) return active;
  return SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
}

function ensureSheet_(ss, name, index) {
  return ss.getSheetByName(name) || ss.insertSheet(name, index);
}

function ensureColumns_(sh, n) {
  const max = sh.getMaxColumns();
  if (max < n) sh.insertColumnsAfter(max, n - max);
}

function styleHeader_(range, bg, fg) {
  range.setBackground(bg).setFontColor(fg)
    .setFontFamily('Sarabun').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

/** ชีตจากเวอร์ชันก่อน (มี 'อัปเดตล่าสุด' ที่คอลัมน์ 19) → แทรกคอลัมน์ 'รูปแบบงานวิจัย' ข้อมูลเดิมเลื่อนตามอัตโนมัติ */
function migrateColumns_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 19) return;
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (head.indexOf('รูปแบบงานวิจัย') >= 0) return;
  if (head[18] === 'อัปเดตล่าสุด') sh.insertColumnBefore(19);
}

function styleSubmissionsSheet_(sh) {
  migrateColumns_(sh);
  const nCol = HEADERS.length;
  ensureColumns_(sh, nCol);
  const maxRows = sh.getMaxRows();
  const bodyRows = Math.max(1, maxRows - 1);

  styleHeader_(sh.getRange(1, 1, 1, nCol).setValues([HEADERS]), '#1b1f3b', '#ffffff');
  sh.getRange(1, col_('ไฟล์ · ' + SLOTS[0].label), 1, SLOTS.length).setBackground('#3b2f8f').setFontColor('#e4dcff');
  sh.setRowHeight(1, 46);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);

  sh.getRange(2, 1, bodyRows, nCol).setFontFamily('Sarabun').setFontSize(11).setVerticalAlignment('middle');
  sh.getRange(2, col_('รหัสโครงการ'), bodyRows, 1).setFontFamily('Roboto Mono').setFontWeight('bold');
  [col_('ส่งครั้งแรก'), col_('อัปเดตล่าสุด')].forEach(function (c) {
    sh.getRange(2, c, bodyRows, 1).setNumberFormat(DT_FORMAT).setHorizontalAlignment('center');
  });
  sh.getRange(2, col_('ชื่อเรื่อง'), bodyRows, 1).setWrap(true);
  sh.getRange(2, col_('สถานที่ทำงาน'), bodyRows, 1).setWrap(true);
  sh.getRange(2, col_('รูปแบบงานวิจัย'), bodyRows, 1).setWrap(true);
  sh.getRange(2, col_('ไฟล์ · ' + SLOTS[0].label), bodyRows, SLOTS.length + 1).setHorizontalAlignment('center');
  ['จำนวนไฟล์', 'สถานะการตรวจ', 'จำนวนครั้งที่ส่ง'].forEach(function (h) {
    sh.getRange(2, col_(h), bodyRows, 1).setHorizontalAlignment('center');
  });

  const widths = {
    'รหัสโครงการ': 175, 'ส่งครั้งแรก': 150, 'ชื่อเรื่อง': 380, 'ชื่อ': 120, 'สกุล': 140,
    'สถานที่ทำงาน': 230, 'โฟลเดอร์งาน': 115, 'จำนวนไฟล์': 85, 'ความครบถ้วน': 230,
    'สถานะการตรวจ': 130, 'หมายเหตุ': 260, 'รูปแบบงานวิจัย': 210, 'อัปเดตล่าสุด': 150, 'จำนวนครั้งที่ส่ง': 95, 'ข้อมูลระบบ': 120,
  };
  Object.keys(widths).forEach(function (h) { sh.setColumnWidth(col_(h), widths[h]); });
  SLOTS.forEach(function (s) { sh.setColumnWidth(col_('ไฟล์ · ' + s.label), 110); });
  sh.showColumns(1, nCol);
  sh.hideColumns(col_('ข้อมูลระบบ'));   // ระบบใช้เก็บสถานะไฟล์ — ห้ามแก้ไข

  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.getRange(1, 1, maxRows, nCol)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor('#1b1f3b').setFirstRowColor('#ffffff').setSecondRowColor('#f5f4ff');

  const reviewRange = sh.getRange(2, col_('สถานะการตรวจ'), bodyRows, 1);
  reviewRange.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(REVIEW_STATUSES, true).setAllowInvalid(false).build()
  );

  const completeRange = sh.getRange(2, col_('ความครบถ้วน'), bodyRows, 1);
  const palette = {
    'รอตรวจ':       ['#e8eefc', '#2b4a8b'],
    'กำลังตรวจ':    ['#fff1d6', '#8a5a00'],
    'ส่งกลับแก้ไข': ['#ffe3e8', '#a1123a'],
    'ผ่าน':         ['#e3f9d2', '#2f6b00'],
  };
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('ครบถ้วน')
      .setBackground('#e3f9d2').setFontColor('#2f6b00').setBold(true).setRanges([completeRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('ไม่ครบ')
      .setBackground('#fff1d6').setFontColor('#8a5a00').setBold(true).setRanges([completeRange]).build(),
  ];
  REVIEW_STATUSES.forEach(function (st) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(st)
      .setBackground(palette[st][0]).setFontColor(palette[st][1]).setBold(true)
      .setRanges([reviewRange]).build());
  });
  sh.setConditionalFormatRules(rules);
  sh.setTabColor('#6d4aff');
}

function styleHistorySheet_(sh) {
  const nCol = HISTORY_HEADERS.length;
  ensureColumns_(sh, nCol);
  styleHeader_(sh.getRange(1, 1, 1, nCol).setValues([HISTORY_HEADERS]), '#3b2f8f', '#ffffff');
  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  const body = Math.max(1, sh.getMaxRows() - 1);
  sh.getRange(2, 1, body, nCol).setFontFamily('Sarabun').setFontSize(11).setVerticalAlignment('middle');
  sh.getRange(2, 1, body, 1).setNumberFormat(DT_FORMAT);
  sh.getRange(2, 2, body, 1).setFontFamily('Roboto Mono').setFontWeight('bold');
  sh.getRange(2, 3, body, 1).setHorizontalAlignment('center');
  sh.getRange(2, 7, body, 1).setWrap(true);
  [150, 175, 70, 190, 160, 360, 320].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setTabColor('#e83e8c');
}

function styleLogSheet_(sh) {
  const nCol = LOG_HEADERS.length;
  ensureColumns_(sh, nCol);
  styleHeader_(sh.getRange(1, 1, 1, nCol).setValues([LOG_HEADERS]), '#2a2f45', '#dfe3f5');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat(DT_FORMAT);
  [150, 80, 110, 175, 560].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setTabColor('#6c7c99');
}

/** ลบเฉพาะชีตเริ่มต้นที่ว่างเปล่า (Sheet1 / แผ่น1 / ชีต1) */
function removeEmptyDefaultSheets_(ss) {
  ss.getSheets().forEach(function (s) {
    if (/^(Sheet|แผ่น|แผ่นงาน|ชีต)\s*1$/.test(s.getName()) && s.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(s);
    }
  });
}


/* ═══════════════════════════ WEB APP ═══════════════════════════ */

/** GET ?action=ping — ตรวจสถานะการเชื่อมต่อ · GET ?action=list[&key=…] — ข้อมูลหน้า Dashboard */
function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = params.action || 'ping';
  if (action === 'list') {
    try { return json_(handleList_(params)); }
    catch (err) { return json_({ ok: false, error: (err && err.message) || String(err) }); }
  }
  if (action === 'ping') {
    return json_({
      ok: true,
      service: 'research-uplink',
      api: CONFIG.API_VERSION,
      ready: !!PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'),
      maxFileMB: CONFIG.MAX_FILE_MB,
      time: new Date().toISOString(),
    });
  }
  return json_({ ok: false, error: 'ไม่รู้จักคำสั่ง: ' + action });
}

/** POST (Content-Type: text/plain, body = JSON) */
function doPost(e) {
  let req = {};
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('ไม่พบข้อมูลที่ส่งมา');
    req = JSON.parse(e.postData.contents);
    switch (req.action) {
      case 'lookup': return json_(handleLookup_(req));
      case 'begin':  return json_(handleBegin_(req));
      case 'upload': return json_(handleUpload_(req));
      case 'finish': return json_(handleFinish_(req));
      case 'update': return json_(handleUpdate_(req));
      case 'delete': return json_(handleDelete_(req));
      case 'auth':   requireAdmin_(req); return json_({ ok: true });
      default: throw new Error('ไม่รู้จักคำสั่ง: ' + req.action);
    }
  } catch (err) {
    const message = (err && err.message) || String(err);
    log_('ERROR', req.action || '-', req.submissionId || req.projectId || req.code || '-',
      message + (req.slot ? ' [ช่อง ' + req.slot + ']' : ''));
    return json_({ ok: false, error: message });
  }
}

/** ข้อมูลสรุปทุกโครงการสำหรับ Dashboard (อ่านจากชีตอย่างเดียว ไม่เรียก Drive) */
function handleList_(params) {
  if (CONFIG.DASHBOARD_KEY && String(params.key || '') !== CONFIG.DASHBOARD_KEY) {
    throw new Error('ต้องใช้รหัสผ่าน Dashboard ที่ถูกต้อง');
  }
  const details = !!CONFIG.DASHBOARD_DETAILS;
  const iso = function (d) { return d instanceof Date && !isNaN(d.getTime()) ? d.toISOString() : ''; };
  const history = details ? readHistoryForList_(iso) : {};
  const sh = getSheet_(CONFIG.SHEET_SUBMISSIONS);
  const n = sh.getLastRow() - 1;
  const projects = [];
  if (n > 0) {
    const values = sh.getRange(2, 1, n, HEADERS.length).getValues();
    const ix = {};
    HEADERS.forEach(function (h, i) { ix[h] = i; });
    values.forEach(function (r) {
      const code = String(r[ix['รหัสโครงการ']]);
      if (!ID_RE.test(code)) return;
      let state = null;
      try { state = JSON.parse(r[ix['ข้อมูลระบบ']] || 'null'); } catch (e) { state = null; }
      const slots = {};
      SLOTS.forEach(function (s) {
        const f = state && state.files && state.files[s.key];
        if (f) {
          const o = { v: f.v || 1, at: f.at || '' };
          if (details) { o.id = f.id; o.name = f.original || f.name || ''; }
          slots[s.key] = o;
          return;
        }
        const cell = String(r[ix['ไฟล์ · ' + s.label]] || '').trim();      // แถวจากเวอร์ชันก่อน
        if (cell && cell !== '—') slots[s.key] = { v: 1, at: '' };
      });
      const extra = details ? {
        note: String(r[ix['หมายเหตุ']] || ''),
        folderId: (state && state.folderId) || '',
        history: history[code] || [],
      } : {};
      projects.push(Object.assign({
        code: code,
        title: String(r[ix['ชื่อเรื่อง']]),
        firstName: String(r[ix['ชื่อ']]),
        lastName: String(r[ix['สกุล']]),
        workplace: String(r[ix['สถานที่ทำงาน']]),
        design: String(r[ix['รูปแบบงานวิจัย']] || ''),
        done: Object.keys(slots).length,
        total: SLOTS.length,
        slots: slots,
        review: String(r[ix['สถานะการตรวจ']] || ''),
        rounds: Number(r[ix['จำนวนครั้งที่ส่ง']]) || (state && state.rounds) || 1,
        firstAt: iso(r[ix['ส่งครั้งแรก']]),
        updatedAt: iso(r[ix['อัปเดตล่าสุด']]) || iso(r[ix['ส่งครั้งแรก']]),
      }, extra));
    });
  }
  return {
    ok: true,
    api: CONFIG.API_VERSION,
    details: details,
    canManage: !!CONFIG.ADMIN_KEY,
    generatedAt: new Date().toISOString(),
    slots: SLOTS.map(function (s) { return { key: s.key, label: s.label }; }),
    reviewStatuses: REVIEW_STATUSES,
    projects: projects,
  };
}

/** ค้นหาโครงการ: { code } หรือ { firstName, lastName } */
function handleLookup_(req) {
  const sh = getSheet_(CONFIG.SHEET_SUBMISSIONS);

  if (req.code) {
    const code = normCode_(req.code);
    if (!ID_RE.test(code)) throw new Error('รูปแบบรหัสโครงการไม่ถูกต้อง (ตัวอย่าง: RS-261106-1432-K7QX)');
    const row = findRow_(sh, code);
    if (!row) throw new Error('ไม่พบรหัสโครงการ ' + code + ' — ตรวจสอบรหัสอีกครั้ง หรือค้นหาด้วยชื่อ–สกุล');
    return { ok: true, project: publicProject_(readProject_(sh, row)) };
  }

  const first = normName_(req.firstName), last = normName_(req.lastName);
  if (!first || !last) throw new Error('กรุณากรอกทั้งชื่อและสกุลเพื่อค้นหา');
  const n = sh.getLastRow() - 1;
  const notFound = 'ไม่พบโครงการที่ตรงกับชื่อ–สกุลนี้ — ถ้ายังไม่เคยส่ง ให้เลือก "ส่งงานครั้งแรก"';
  if (n < 1) throw new Error(notFound);

  const values = sh.getRange(2, 1, n, HEADERS.length).getValues();
  const iC = col_('รหัสโครงการ') - 1, iT = col_('ชื่อเรื่อง') - 1, iF = col_('ชื่อ') - 1, iL = col_('สกุล') - 1,
        iN = col_('จำนวนไฟล์') - 1, iU = col_('อัปเดตล่าสุด') - 1, iS = col_('ส่งครั้งแรก') - 1;
  const matches = values
    .filter(function (r) {
      const f = normName_(r[iF]), l = normName_(r[iL]);
      return ID_RE.test(String(r[iC])) && l === last && f && (f === first || f.indexOf(first) >= 0 || first.indexOf(f) >= 0);
    })
    .reverse()
    .slice(0, 10)
    .map(function (r) {
      return {
        code: String(r[iC]),
        title: String(r[iT]),
        done: Number(r[iN]) || 0,
        total: SLOTS.length,
        updatedAt: shortThai_(r[iU] instanceof Date ? r[iU] : r[iS]),
      };
    });
  if (!matches.length) throw new Error(notFound);
  return { ok: true, matches: matches };
}

/** เริ่มการส่ง 1 ครั้ง — โครงการใหม่ หรือโครงการเดิม (projectId) */
function handleBegin_(req) {
  const info = readInfo_(req);
  const batch = newBatchId_();

  if (req.projectId) {
    const code = normCode_(req.projectId);
    if (!ID_RE.test(code)) throw new Error('รูปแบบรหัสโครงการไม่ถูกต้อง');
    const sh = getSheet_(CONFIG.SHEET_SUBMISSIONS);
    const row = findRow_(sh, code);
    if (!row) throw new Error('ไม่พบรหัสโครงการ ' + code);
    const p = readProject_(sh, row);
    const folder = openSubmissionFolder_(p.state.folderId, code);
    return { ok: true, submissionId: code, folderId: folder.getId(), batch: batch, isNew: false, round: (p.state.rounds || 0) + 1 };
  }

  getDb_(); // ตรวจว่ารัน setup แล้ว
  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const id = newSubmissionId_();
  const folder = root.createFolder(folderName_(id, info));
  folder.setDescription('ชื่อเรื่อง: ' + info.title + '\nผู้วิจัย: ' + info.firstName + ' ' + info.lastName
    + '\nสถานที่ทำงาน: ' + info.workplace + (info.design ? '\nรูปแบบงานวิจัย: ' + info.design : ''));
  return { ok: true, submissionId: id, folderId: folder.getId(), batch: batch, isNew: true, round: 1 };
}

/** รับไฟล์ทีละช่อง — ฉบับเดิมของช่องนี้ย้ายไป "ฉบับก่อนหน้า" */
function handleUpload_(req) {
  const info = readInfo_(req);
  const id = readId_(req);
  const batch = readBatch_(req);
  const folder = openSubmissionFolder_(req.folderId, id);

  const index = SLOTS.map(function (s) { return s.key; }).indexOf(String(req.slot));
  if (index < 0) throw new Error('ช่องเอกสารไม่ถูกต้อง: ' + req.slot);
  const slot = SLOTS[index];

  const original = String(req.fileName || '').slice(0, 180);
  const m = /\.([A-Za-z0-9]+)$/.exec(original);
  const ext = m ? m[1].toLowerCase() : '';
  if (CONFIG.ALLOWED_EXT.indexOf(ext) < 0) {
    throw new Error(slot.label + ': ไม่รองรับไฟล์ประเภทนี้ (รองรับ ' + CONFIG.ALLOWED_EXT.join(', ').toUpperCase() + ')');
  }

  const b64 = String(req.data || '').replace(/^data:[^,]*,/, '');
  if (!b64) throw new Error(slot.label + ': ไม่พบข้อมูลไฟล์');
  const bytes = Utilities.base64Decode(b64);
  if (!bytes.length) throw new Error(slot.label + ': ไฟล์ว่างเปล่า');
  if (bytes.length > CONFIG.MAX_FILE_MB * 1024 * 1024) {
    throw new Error(slot.label + ': ไฟล์ใหญ่เกิน ' + CONFIG.MAX_FILE_MB + ' MB');
  }

  // จัดการฉบับเดิมของช่องนี้
  let maxV = 0;
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const d = parseDesc_(f.getDescription());
    if (d.slot !== slot.key) continue;
    if (d.batch && d.batch === batch) { f.setTrashed(true); continue; }   // ส่งซ้ำในครั้งเดียวกัน (กดลองใหม่)
    maxV = Math.max(maxV, d.v);
    archiveFile_(f, d, folder);                                            // เก็บฉบับก่อนหน้า
  }
  const arch = findArchive_(folder);
  if (arch) {
    const ai = arch.getFiles();
    while (ai.hasNext()) {
      const d = parseDesc_(ai.next().getDescription());
      if (d.slot === slot.key) maxV = Math.max(maxV, d.v);
    }
  }
  const v = maxV + 1;

  const name = pad2_(index) + '_' + slot.label + '_' + safeName_(info.firstName + ' ' + info.lastName)
    + (v > 1 ? '_v' + v : '') + '.' + ext;
  const file = folder.createFile(Utilities.newBlob(bytes, MIME[ext], name));
  file.setDescription('slot=' + slot.key + ';v=' + v + ';batch=' + batch + ';original=' + original);
  if (CONFIG.SHARE_FILES_WITH_LINK) shareFile_(file, id);
  return { ok: true, slot: slot.key, fileId: file.getId(), name: name, v: v };
}

/** จบการส่ง — อัปเดตแถวโครงการ + บันทึกประวัติ */
function handleFinish_(req) {
  const info = readInfo_(req);
  const id = readId_(req);
  const batch = readBatch_(req);
  const folder = openSubmissionFolder_(req.folderId, id);
  const sent = (req.files && typeof req.files === 'object') ? req.files : {};
  const now = new Date();

  // ตรวจไฟล์ที่อัปโหลดในครั้งนี้: ต้องอยู่ในโฟลเดอร์โครงการ ตรงช่อง และตรง batch
  const fresh = {};
  SLOTS.forEach(function (s) {
    const fid = sent[s.key];
    if (!fid) return;
    try {
      const f = DriveApp.getFileById(String(fid));
      if (f.isTrashed() || !isInFolder_(f, folder.getId())) return;
      const d = parseDesc_(f.getDescription());
      if (d.slot !== s.key || d.batch !== batch) return;
      fresh[s.key] = { id: f.getId(), name: f.getName(), original: d.original || f.getName(), v: d.v, at: now.toISOString(), batch: batch };
    } catch (e) { /* ข้ามไฟล์ที่ไม่พบ */ }
  });
  if (!Object.keys(fresh).length) throw new Error('ไม่พบไฟล์ที่อัปโหลด กรุณาลองส่งใหม่อีกครั้ง');

  const cache = CacheService.getScriptCache();
  const doneKey = 'done_' + id + '_' + batch;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const cached = cache.get(doneKey);           // กันบันทึกซ้ำ
    if (cached) return JSON.parse(cached);

    const sh = getSheet_(CONFIG.SHEET_SUBMISSIONS);
    const row = findRow_(sh, id);
    const prev = row ? readProject_(sh, row) : null;
    const state = prev ? prev.state : { folderId: folder.getId(), rounds: 0, files: {} };
    state.folderId = folder.getId();
    state.files = state.files || {};

    const added = [], replaced = [];
    SLOTS.forEach(function (s) {
      if (!fresh[s.key]) return;
      (state.files[s.key] ? replaced : added).push(s.key);
      state.files[s.key] = fresh[s.key];
    });
    state.rounds = (state.rounds || 0) + 1;

    // ไฟล์ค้างในโฟลเดอร์ที่ไม่ใช่ฉบับปัจจุบัน (เช่น ส่งไม่สำเร็จแล้วผู้ส่งเอาออก) → ถังขยะ
    const keepIds = Object.keys(state.files).map(function (k) { return state.files[k].id; });
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (keepIds.indexOf(f.getId()) < 0 && parseDesc_(f.getDescription()).slot) f.setTrashed(true);
    }

    const done = SLOTS.filter(function (s) { return state.files[s.key]; }).length;
    const missing = SLOTS.filter(function (s) { return s.required && !state.files[s.key]; })
                         .map(function (s) { return s.label; });

    const cells = {
      'รหัสโครงการ': id,
      'ส่งครั้งแรก': (prev && prev.firstAt instanceof Date) ? prev.firstAt : now,
      'ชื่อเรื่อง': text_(info.title),
      'ชื่อ': text_(info.firstName),
      'สกุล': text_(info.lastName),
      'สถานที่ทำงาน': text_(info.workplace),
      'โฟลเดอร์งาน': link_(folder.getUrl(), 'เปิดโฟลเดอร์'),
      'จำนวนไฟล์': done,
      'ความครบถ้วน': missing.length
        ? 'ไม่ครบ ' + done + '/' + SLOTS.length + ' · ขาด ' + missing.join(', ')
        : 'ครบถ้วน ' + done + '/' + SLOTS.length,
      'สถานะการตรวจ': REVIEW_STATUSES[0],            // มีไฟล์ใหม่ → กลับเป็น "รอตรวจ"
      'หมายเหตุ': prev ? text_(prev.note) : '',
      'รูปแบบงานวิจัย': text_(info.design || (prev ? prev.design : '')),
      'อัปเดตล่าสุด': now,
      'จำนวนครั้งที่ส่ง': state.rounds,
      'ข้อมูลระบบ': JSON.stringify(state),
    };
    SLOTS.forEach(function (s) {
      const f = state.files[s.key];
      cells['ไฟล์ · ' + s.label] = f ? link_(fileUrl_(f.id), 'v' + f.v + ' · ' + numDate_(new Date(f.at))) : '—';
    });
    const values = HEADERS.map(function (h) { return Object.prototype.hasOwnProperty.call(cells, h) ? cells[h] : ''; });

    let r;
    if (prev) {
      r = row;
      sh.getRange(r, 1, 1, HEADERS.length).setValues([values]);
    } else {
      sh.appendRow(values);
      r = sh.getLastRow();
    }
    sh.getRange(r, col_('ส่งครั้งแรก')).setNumberFormat(DT_FORMAT);
    sh.getRange(r, col_('อัปเดตล่าสุด')).setNumberFormat(DT_FORMAT);

    const kind = !prev ? 'ส่งครั้งแรก'
      : (added.length && replaced.length) ? 'ส่งเพิ่ม + ฉบับแก้ไข'
      : replaced.length ? 'ส่งฉบับแก้ไข' : 'ส่งเพิ่ม';
    appendHistory_(now, id, state.rounds, info, kind, fresh);
    SpreadsheetApp.flush();

    const result = {
      ok: true,
      submissionId: id,
      isNew: !prev,
      round: state.rounds,
      kind: kind,
      timestamp: thaiDateTime_(now),
      sentNow: SLOTS.filter(function (s) { return fresh[s.key]; }).map(function (s) {
        return s.label + (fresh[s.key].v > 1 ? ' (ฉบับที่ ' + fresh[s.key].v + ')' : '');
      }),
      done: done,
      total: SLOTS.length,
      complete: missing.length === 0,
      missing: missing,
    };
    cache.put(doneKey, JSON.stringify(result), 21600);
    log_('INFO', 'finish', id, kind + ' · ครั้งที่ ' + state.rounds + ' · ' + result.sentNow.join(', ') + ' · แถวที่ ' + r);
    try { folder.setName(folderName_(id, info)); } catch (e) { /* ไม่สำคัญ */ }
    return result;
  } finally {
    lock.releaseLock();
  }
}


/* ═══════════════════════════ จัดการจาก DASHBOARD (ต้องใช้ ADMIN_KEY) ═══════════════════════════ */

/** แก้ไขข้อมูลโครงการ: { adminKey, submissionId, title, firstName, lastName, workplace, design, review, note } */
function handleUpdate_(req) {
  requireAdmin_(req);
  const id = readId_(req);
  const info = readInfo_(req);
  const review = String(req.review || '');
  if (review && REVIEW_STATUSES.indexOf(review) < 0) throw new Error('สถานะการตรวจไม่ถูกต้อง: ' + review);
  const note = String(req.note == null ? '' : req.note)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim().slice(0, 2000);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getSheet_(CONFIG.SHEET_SUBMISSIONS);
    const row = findRow_(sh, id);
    if (!row) throw new Error('ไม่พบโครงการ ' + id);
    const values = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
    const old = function (h) { return String(values[col_(h) - 1] == null ? '' : values[col_(h) - 1]); };
    const next = {
      'ชื่อเรื่อง': info.title, 'ชื่อ': info.firstName, 'สกุล': info.lastName, 'สถานที่ทำงาน': info.workplace,
      'รูปแบบงานวิจัย': info.design || old('รูปแบบงานวิจัย'), 'หมายเหตุ': note,
    };
    if (review) next['สถานะการตรวจ'] = review;
    const changed = [];
    Object.keys(next).forEach(function (h) {
      if (old(h).replace(/^'/, '') === next[h]) return;
      changed.push(h);
      sh.getRange(row, col_(h)).setValue(h === 'สถานะการตรวจ' ? next[h] : text_(next[h]));
    });
    SpreadsheetApp.flush();
    if (changed.indexOf('ชื่อ') >= 0 || changed.indexOf('สกุล') >= 0) {
      const st = readState_(sh, row);
      try { if (st && st.folderId) DriveApp.getFolderById(st.folderId).setName(folderName_(id, info)); } catch (e) { /* ไม่สำคัญ */ }
    }
    log_('INFO', 'update', id, 'แก้ไขจาก Dashboard: ' + (changed.join(', ') || 'ไม่มีการเปลี่ยนแปลง'));
    return { ok: true, code: id, changed: changed };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ลบ: { adminKey, submissionId, slot } = ลบไฟล์บทนั้น (ฉบับปัจจุบัน)
 *     { adminKey, submissionId, confirm: รหัสโครงการ } = ลบทั้งโครงการ (แถวในชีต + โฟลเดอร์)
 * ไฟล์/โฟลเดอร์ย้ายไปถังขยะของ Google Drive (กู้คืนได้ภายใน 30 วัน) · ประวัติการส่งยังเก็บไว้
 */
function handleDelete_(req) {
  requireAdmin_(req);
  const id = readId_(req);
  const slotKey = req.slot ? String(req.slot) : '';
  const slot = slotKey ? SLOTS.filter(function (s) { return s.key === slotKey; })[0] : null;
  if (slotKey && !slot) throw new Error('ช่องเอกสารไม่ถูกต้อง: ' + slotKey);
  if (!slot && normCode_(req.confirm) !== id) throw new Error('กรุณาพิมพ์รหัสโครงการให้ตรงเพื่อยืนยันการลบ');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getSheet_(CONFIG.SHEET_SUBMISSIONS);
    const row = findRow_(sh, id);
    if (!row) throw new Error('ไม่พบโครงการ ' + id);
    const state = readState_(sh, row);

    if (slot) {
      const f = state && state.files && state.files[slot.key];
      if (!f) throw new Error(slot.label + ' ยังไม่มีไฟล์ให้ลบ');
      try { DriveApp.getFileById(f.id).setTrashed(true); } catch (e) { /* ไฟล์อาจถูกลบไปแล้ว */ }
      delete state.files[slot.key];
      const done = SLOTS.filter(function (s) { return state.files[s.key]; }).length;
      const missing = SLOTS.filter(function (s) { return s.required && !state.files[s.key]; }).map(function (s) { return s.label; });
      const set = function (h, v) { sh.getRange(row, col_(h)).setValue(v); };
      set('ไฟล์ · ' + slot.label, '—');
      set('จำนวนไฟล์', done);
      set('ความครบถ้วน', missing.length ? 'ไม่ครบ ' + done + '/' + SLOTS.length + ' · ขาด ' + missing.join(', ') : 'ครบถ้วน ' + done + '/' + SLOTS.length);
      set('อัปเดตล่าสุด', new Date());
      set('ข้อมูลระบบ', JSON.stringify(state));
      SpreadsheetApp.flush();
      log_('WARN', 'delete', id, 'ลบไฟล์ ' + slot.label + ' (v' + f.v + ' · ' + (f.original || f.name) + ') จาก Dashboard → ถังขยะ Drive');
      return { ok: true, code: id, slot: slot.key, done: done };
    }

    const v = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
    let trashed = false;
    try { if (state && state.folderId) { DriveApp.getFolderById(state.folderId).setTrashed(true); trashed = true; } } catch (e) { /* ไม่พบโฟลเดอร์ */ }
    sh.deleteRow(row);
    SpreadsheetApp.flush();
    log_('WARN', 'delete', id, 'ลบโครงการจาก Dashboard · ' + v[col_('ชื่อเรื่อง') - 1] + ' · ' + v[col_('ชื่อ') - 1] + ' ' + v[col_('สกุล') - 1]
      + (trashed ? ' · โฟลเดอร์ย้ายไปถังขยะ Drive' : ' · ไม่พบโฟลเดอร์'));
    return { ok: true, code: id, deleted: true };
  } finally {
    lock.releaseLock();
  }
}

function requireAdmin_(req) {
  if (!CONFIG.ADMIN_KEY) throw new Error('ยังไม่เปิดใช้การแก้ไข/ลบ — ผู้ดูแลต้องตั้ง ADMIN_KEY ใน Code.gs แล้ว Deploy เวอร์ชันใหม่');
  const userOk = !CONFIG.ADMIN_USER ||
    String(req.adminUser || '').trim().toLowerCase() === String(CONFIG.ADMIN_USER).trim().toLowerCase();
  if (!userOk || String(req.adminKey || '') !== String(CONFIG.ADMIN_KEY)) {
    Utilities.sleep(800);                                 // ชะลอการสุ่มเดารหัส
    throw new Error('ชื่อผู้ใช้หรือรหัสผู้ดูแลไม่ถูกต้อง');
  }
}

/** สถานะไฟล์ของแถว (ไม่ throw) */
function readState_(sh, row) {
  try {
    const st = JSON.parse(sh.getRange(row, col_('ข้อมูลระบบ')).getValues()[0][0] || 'null');
    if (st && st.folderId) { st.files = st.files || {}; return st; }
  } catch (e) { /* ข้าม */ }
  try { return legacyState_(String(sh.getRange(row, 1).getValues()[0][0])); } catch (e) { return null; }
}

/** ประวัติการส่งของทุกโครงการ (สำหรับหน้างานวิจัยรายโครงการ) — รวม id ไฟล์ทุกฉบับจากลิงก์ในชีต */
function readHistoryForList_(iso) {
  const out = {};
  const sh = getSheet_(CONFIG.SHEET_HISTORY);
  const n = sh.getLastRow() - 1;
  if (n < 1) return out;
  const values = sh.getRange(2, 1, n, 6).getValues();
  let rich = [];
  try { rich = sh.getRange(2, 6, n, 1).getRichTextValues(); } catch (e) { rich = []; }
  values.forEach(function (r, i) {
    const code = String(r[1]);
    if (!ID_RE.test(code)) return;
    const docs = [];
    const rt = rich[i] && rich[i][0];
    if (rt) rt.getRuns().forEach(function (run) {
      const m = /\/d\/([\w-]+)/.exec(run.getLinkUrl() || '');
      if (m) docs.push({ label: run.getText().trim(), id: m[1] });
    });
    (out[code] = out[code] || []).push({ at: iso(r[0]), round: Number(r[2]) || 0, kind: String(r[4] || ''), text: String(r[5] || ''), docs: docs });
  });
  return out;
}

function shareFile_(file, id) {
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { log_('WARN', 'share', id || '-', 'ตั้งค่าแชร์ไฟล์ไม่สำเร็จ: ' + ((e && e.message) || e)); }
}

/** รันเองครั้งเดียว (ถ้าเปิด SHARE_FILES_WITH_LINK): ตั้งไฟล์ทุกไฟล์ในโฟลเดอร์งาน รวมฉบับก่อนหน้า เป็น "ทุกคนที่มีลิงก์ดูได้" */
function shareAllFiles() {
  let n = 0;
  const walk = function (folder) {
    const fi = folder.getFiles();
    while (fi.hasNext()) { shareFile_(fi.next(), folder.getName()); n++; }
    const fo = folder.getFolders();
    while (fo.hasNext()) walk(fo.next());
  };
  const it = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID).getFolders();
  while (it.hasNext()) walk(it.next());
  log_('INFO', 'share', '-', 'shareAllFiles: ตั้งค่าแชร์ ' + n + ' ไฟล์');
  console.log('ตั้งค่าแชร์แล้ว ' + n + ' ไฟล์');
}


/* ═══════════════════════════ PROJECT DATA ═══════════════════════════ */

function findRow_(sh, code) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const cell = sh.getRange(2, 1, last - 1, 1).createTextFinder(code).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : 0;
}

function readProject_(sh, row) {
  const values = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  const get = function (h) { return values[col_(h) - 1]; };
  const code = String(get('รหัสโครงการ'));
  let state = null;
  try { state = JSON.parse(get('ข้อมูลระบบ') || 'null'); } catch (e) { state = null; }
  if (!state || !state.folderId) state = legacyState_(code);   // แถวจากเวอร์ชันก่อน
  return {
    row: row,
    code: code,
    firstAt: get('ส่งครั้งแรก'),
    updatedAt: get('อัปเดตล่าสุด'),
    title: String(get('ชื่อเรื่อง')),
    firstName: String(get('ชื่อ')),
    lastName: String(get('สกุล')),
    workplace: String(get('สถานที่ทำงาน')),
    review: String(get('สถานะการตรวจ') || ''),
    note: String(get('หมายเหตุ') || ''),
    design: String(get('รูปแบบงานวิจัย') || ''),
    state: state,
  };
}

/** สร้างสถานะไฟล์จากโฟลเดอร์ (สำหรับแถวที่บันทึกด้วยโค้ดเวอร์ชันก่อน) */
function legacyState_(code) {
  const q = "title contains '" + code + "' and '" + CONFIG.ROOT_FOLDER_ID + "' in parents and trashed = false";
  const it = DriveApp.searchFolders(q);
  if (!it.hasNext()) throw new Error('ไม่พบโฟลเดอร์ของโครงการ ' + code);
  const folder = it.next();
  const files = {};
  const fi = folder.getFiles();
  while (fi.hasNext()) {
    const f = fi.next();
    const d = parseDesc_(f.getDescription());
    if (!d.slot) continue;
    files[d.slot] = { id: f.getId(), name: f.getName(), original: d.original || f.getName(), v: d.v,
      at: f.getDateCreated().toISOString(), batch: d.batch };
  }
  return { folderId: folder.getId(), rounds: 1, files: files };
}

/** ข้อมูลที่ส่งกลับให้หน้าเว็บ (ไม่มี id ไฟล์/โฟลเดอร์) */
function publicProject_(p) {
  const files = {};
  Object.keys(p.state.files || {}).forEach(function (k) {
    const f = p.state.files[k];
    files[k] = { name: f.name, original: f.original, v: f.v, at: shortThai_(new Date(f.at)) };
  });
  return {
    code: p.code,
    title: p.title,
    firstName: p.firstName,
    lastName: p.lastName,
    workplace: p.workplace,
    design: p.design,
    review: p.review,
    note: p.note,
    rounds: p.state.rounds || 0,
    updatedAt: shortThai_(p.updatedAt instanceof Date ? p.updatedAt : p.firstAt),
    files: files,
  };
}

function appendHistory_(now, id, round, info, kind, fresh) {
  const sh = getSheet_(CONFIG.SHEET_HISTORY);
  const parts = SLOTS.filter(function (s) { return fresh[s.key]; }).map(function (s) {
    return { text: s.label + ' v' + fresh[s.key].v, url: fileUrl_(fresh[s.key].id), original: s.label + ': ' + fresh[s.key].original };
  });
  const SEP = ' · ';
  const text = parts.map(function (p) { return p.text; }).join(SEP);
  sh.appendRow([now, id, round, text_(info.firstName + ' ' + info.lastName), kind, text,
    text_(parts.map(function (p) { return p.original; }).join('\n'))]);
  const r = sh.getLastRow();
  const rich = SpreadsheetApp.newRichTextValue().setText(text);
  let pos = 0;
  parts.forEach(function (p) {
    rich.setLinkUrl(pos, pos + p.text.length, p.url);
    pos += p.text.length + SEP.length;
  });
  sh.getRange(r, 6).setRichTextValue(rich.build());
  sh.getRange(r, 1).setNumberFormat(DT_FORMAT);
}

function findArchive_(folder) {
  const it = folder.getFoldersByName(CONFIG.ARCHIVE_FOLDER);
  return it.hasNext() ? it.next() : null;
}

function archiveFile_(file, d, folder) {
  const arch = findArchive_(folder) || folder.createFolder(CONFIG.ARCHIVE_FOLDER);
  const name = file.getName();
  if (!/_v\d+\.[^.]+$/.test(name)) file.setName(name.replace(/(\.[^.]+)$/, '_v' + d.v + '$1'));
  file.moveTo(arch);
}

/** description ของไฟล์: slot=ch1;v=2;batch=B…;original=ชื่อไฟล์เดิม.pdf (รองรับรูปแบบเก่า slot=ch1;original=…) */
function parseDesc_(s) {
  s = String(s || '');
  const out = { slot: '', v: 1, batch: '', original: '' };
  const m = /^slot=([a-z0-9]+);/.exec(s);
  if (!m) return out;
  out.slot = m[1];
  const v = /;v=(\d+);/.exec(s);           if (v) out.v = Number(v[1]);
  const b = /;batch=([A-Za-z0-9]*);/.exec(s); if (b) out.batch = b[1];
  const o = /;original=([\s\S]*)$/.exec(s); if (o) out.original = o[1];
  return out;
}


/* ═══════════════════════════ HELPERS ═══════════════════════════ */

function readInfo_(req) {
  const clean = function (v, max) {
    return String(v == null ? '' : v)
      .replace(/[\u0000-\u001F\u007F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, max);
  };
  const info = {
    title:     clean(req.title, 500),
    firstName: clean(req.firstName, 100),
    lastName:  clean(req.lastName, 100),
    workplace: clean(req.workplace, 300),
    design:    clean(req.design, 150),   // รูปแบบงานวิจัย (ไม่บังคับฝั่งเซิร์ฟเวอร์ เพื่อรองรับหน้าเว็บรุ่นเก่า)
  };
  const missing = [];
  if (!info.title) missing.push('ชื่อเรื่อง');
  if (!info.firstName) missing.push('ชื่อ');
  if (!info.lastName) missing.push('สกุล');
  if (!info.workplace) missing.push('สถานที่ทำงาน');
  if (missing.length) throw new Error('กรุณากรอก: ' + missing.join(', '));
  return info;
}

function readId_(req) {
  const id = normCode_(req.submissionId);
  if (!ID_RE.test(id)) throw new Error('รหัสโครงการไม่ถูกต้อง');
  return id;
}

function readBatch_(req) {
  return String(req.batch || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 24) || 'B0';
}

function normCode_(s) {
  return String(s || '').toUpperCase().replace(/\s+/g, '');
}

function normName_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();
}

/** เปิดโฟลเดอร์โครงการ — ต้องขึ้นต้นด้วยรหัสโครงการ และอยู่ใต้โฟลเดอร์หลักเท่านั้น */
function openSubmissionFolder_(folderId, id) {
  let folder;
  try {
    folder = DriveApp.getFolderById(String(folderId || ''));
  } catch (e) {
    throw new Error('ไม่พบโฟลเดอร์ของโครงการนี้');
  }
  if (folder.isTrashed() || folder.getName().indexOf(id) !== 0) {
    throw new Error('โฟลเดอร์ไม่ตรงกับรหัสโครงการ');
  }
  const parents = folder.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === CONFIG.ROOT_FOLDER_ID) return folder;
  }
  throw new Error('โฟลเดอร์ปลายทางไม่ถูกต้อง');
}

function isInFolder_(file, folderId) {
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === folderId) return true;
  }
  return false;
}

function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('ระบบยังไม่พร้อมใช้งาน — ผู้ดูแลต้องรันฟังก์ชัน setup() ก่อน');
  return SpreadsheetApp.openById(id);
}

function getSheet_(name) {
  const sh = getDb_().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต "' + name + '" — กรุณารัน setup() อีกครั้ง');
  return sh;
}

function log_(level, action, id, detail) {
  try {
    getSheet_(CONFIG.SHEET_LOGS).appendRow([new Date(), level, action, id || '-', String(detail || '').slice(0, 1500)]);
  } catch (e) {
    console.error('เขียนบันทึกระบบไม่สำเร็จ: ' + e);
  }
  (level === 'ERROR' ? console.error : console.log)('[' + action + '] ' + id + ' — ' + detail);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function randomChars_(n) {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let r = '';
  for (let i = 0; i < n; i++) r += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  return r;
}

function newSubmissionId_() {
  return 'RS-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd-HHmm') + '-' + randomChars_(4);
}

function newBatchId_() {
  return 'B' + Date.now().toString(36).toUpperCase() + randomChars_(3);
}

function folderName_(id, info) {
  return id + '_' + safeName_(info.firstName + ' ' + info.lastName);
}

function safeName_(s) {
  return String(s).replace(/[\\/:*?"<>|#%]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80) || 'ไม่ระบุชื่อ';
}

function fileUrl_(id) {
  return 'https://drive.google.com/file/d/' + id + '/view';
}

/** กันข้อความที่ขึ้นต้นด้วย = + - @ ถูกตีความเป็นสูตรใน Sheets */
function text_(s) {
  s = String(s == null ? '' : s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function link_(url, label) {
  return '=HYPERLINK("' + String(url).replace(/"/g, '') + '","' + String(label).replace(/"/g, '') + '")';
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

const TH_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const TH_MONTHS_SHORT = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

function dateParts_(d) {
  const tz = CONFIG.TIMEZONE;
  return {
    day: Number(Utilities.formatDate(d, tz, 'd')),
    month: Number(Utilities.formatDate(d, tz, 'M')) - 1,
    yearBE: Number(Utilities.formatDate(d, tz, 'yyyy')) + 543,
    time: Utilities.formatDate(d, tz, 'HH:mm'),
  };
}

function thaiDateTime_(d) {
  const p = dateParts_(d);
  return p.day + ' ' + TH_MONTHS[p.month] + ' ' + p.yearBE + ' เวลา ' + p.time + ' น.';
}

function shortThai_(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return '';
  const p = dateParts_(d);
  return p.day + ' ' + TH_MONTHS_SHORT[p.month] + ' ' + String(p.yearBE).slice(-2);
}

function numDate_(d) {
  const p = dateParts_(d);
  return p.day + '/' + (p.month + 1) + '/' + String(p.yearBE).slice(-2);
}
