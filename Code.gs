/**
 * ระบบบันทึกข้อมูลพนักงาน (Google Apps Script + Google Sheets)
 * ต้องมีชีตชื่อ "Staff" และ "Users" อยู่ในสเปรดชีตนี้
 *
 * โครงสร้างชีต "Staff" (แถวที่ 1 เป็นหัวตาราง):
 * No | Employee ID | Name in English | Name in Thai | Nick Name | Date of Birth |
 * Start Date | Phone Number | Emergency Contact Number | E mail | Address | Status | Resigned Date
 *
 * โครงสร้างชีต "Users":
 * EmpID | Pass | Name
 *
 * ชีต "Pending" จะถูกสร้างอัตโนมัติ (ไม่ต้องสร้างเอง) ใช้เก็บข้อมูลที่พนักงานทั่วไป
 * กรอกส่งเข้ามารอแอดมินตรวจสอบ/อนุมัติก่อนเข้าชีต "Staff":
 * Timestamp | Employee ID | Name in English | Name in Thai | Nick Name | Date of Birth |
 * Start Date | Phone Number | Emergency Contact Number | E mail | Address
 */

const SHEET_STAFF = 'Staff';
const SHEET_USERS = 'Users';
const SHEET_PENDING = 'Pending';
const SESSION_DURATION_SEC = 6 * 60 * 60; // อายุ session แอดมิน 6 ชั่วโมง

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ทะเบียนประวัติพนักงาน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getStaffSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STAFF);
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "Staff" กรุณาสร้างชีตนี้ก่อน');
  return sheet;
}

function getUsersSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "Users" กรุณาสร้างชีตนี้ก่อน');
  return sheet;
}

function getPendingSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PENDING);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PENDING);
    sheet.appendRow([
      'Timestamp', 'Employee ID', 'Name in English', 'Name in Thai', 'Nick Name',
      'Date of Birth', 'Start Date', 'Phone Number', 'Emergency Contact Number', 'E mail', 'Address'
    ]);
  }
  return sheet;
}

function fmtDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function isActiveStatus_(status) {
  return !status || status === 'active';
}

// คำนวณช่วงเวลาแบบ "X ปี Y เดือน Z วัน"
// toStr = null/undefined -> นับถึงวันนี้ (ใช้กับคนที่ยังทำงานอยู่)
// toStr = วันที่ (string) -> นับถึงวันนั้น (ใช้กับคนที่ลาออกแล้ว)
// toStr = '' (ว่างแต่ระบุมา) -> ไม่ทราบวันสิ้นสุด หยุดนับ ไม่เดาว่าเป็นวันนี้
function calcDuration_(fromStr, toStr) {
  if (!fromStr) return '-';
  const from = new Date(fromStr);
  if (isNaN(from.getTime())) return '-';
  let to;
  if (toStr === null || toStr === undefined) {
    to = new Date();
  } else if (toStr) {
    to = new Date(toStr);
  } else {
    return '-'; // ลาออกแล้วแต่ไม่มีวันที่บันทึกไว้ - ไม่นับเพิ่ม
  }
  let years = to.getFullYear() - from.getFullYear();
  let months = to.getMonth() - from.getMonth();
  let days = to.getDate() - from.getDate();
  if (days < 0) {
    months -= 1;
    const prevMonthLastDay = new Date(to.getFullYear(), to.getMonth(), 0);
    days += prevMonthLastDay.getDate();
  }
  if (months < 0) { years -= 1; months += 12; }
  if (years < 0) return '-';
  return years + ' ปี ' + months + ' เดือน ' + days + ' วัน';
}

function rowToObj_(row, rowNum) {
  return {
    row: rowNum,
    no: row[0],
    empId: row[1],
    nameEn: row[2],
    nameTh: row[3],
    nickname: row[4],
    dob: fmtDate_(row[5]),
    startDate: fmtDate_(row[6]),
    phone: row[7] ? String(row[7]) : '',
    emergencyPhone: row[8] ? String(row[8]) : '',
    email: row[9] || '',
    address: row[10] || '',
    status: row[11] || 'active',
    resignedDate: fmtDate_(row[12]),
  };
}

// เรียกได้ทุกคน (ไม่ต้องล็อกอิน) - เห็นเฉพาะ "ชื่อ" ของพนักงานที่ยังทำงานอยู่ ไม่มีข้อมูลส่วนตัวอื่น
function getPublicStaffList() {
  const sheet = getStaffSheet_();
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue; // ข้ามแถวว่าง (ไม่มี Employee ID)
    if (!isActiveStatus_(row[11])) continue; // แสดงเฉพาะคนที่ยังทำงานอยู่
    out.push({
      empId: row[1],
      nameEn: row[2],
      nameTh: row[3],
      nickname: row[4],
    });
  }
  return out;
}

// ต้องมีสิทธิ์แอดมิน - เห็นข้อมูลเต็มของทุกคน รวมคนที่ลาออกแล้ว
function getStaffList(token) {
  requireAdmin_(token);
  const sheet = getStaffSheet_();
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue; // ข้ามแถวว่าง (ไม่มี Employee ID)
    const obj = rowToObj_(row, i + 1);
    obj.age = calcDuration_(obj.dob, null);
    obj.tenure = calcDuration_(obj.startDate, isActiveStatus_(obj.status) ? null : obj.resignedDate);
    out.push(obj);
  }
  return out;
}

// เข้าสู่ระบบแอดมิน คืนค่า token ไว้ยืนยันตัวตนในคำขอถัดไป
function login(empId, pass) {
  const sheet = getUsersSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]).trim() === String(empId).trim() && String(row[1]).trim() === String(pass).trim()) {
      const token = Utilities.getUuid();
      CacheService.getScriptCache().put(
        'session_' + token,
        JSON.stringify({ empId: row[0], name: row[2] }),
        SESSION_DURATION_SEC
      );
      return { success: true, token: token, name: row[2] };
    }
  }
  return { success: false, message: 'รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง' };
}

function verifySession_(token) {
  if (!token) return null;
  const cached = CacheService.getScriptCache().get('session_' + token);
  return cached ? JSON.parse(cached) : null;
}

function requireAdmin_(token) {
  const session = verifySession_(token);
  if (!session) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบแอดมินใหม่');
  return session;
}

// ตรวจสอบว่า token ยังใช้ได้อยู่ไหม (เผื่อรีเฟรชหน้า)
function checkSession(token) {
  const session = verifySession_(token);
  return session ? { valid: true, name: session.name } : { valid: false };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('session_' + token);
  return { success: true };
}

function todayStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// กันเหตุการณ์ที่สถานะเป็น "ลาออกแล้ว" แต่ไม่มีวันที่ลาออกติดมาด้วย (เช่น เรียก API ตรง ๆ
// หรือแก้ไขชีตเอง) โดยเติมวันที่วันนี้ให้ทันที เพื่อไม่ให้อายุงานนับเพิ่มไปเรื่อย ๆ
function normalizeRecord_(record) {
  const out = Object.assign({}, record);
  if (!isActiveStatus_(out.status) && !out.resignedDate) {
    out.resignedDate = todayStr_();
  }
  return out;
}

// เพิ่มแถวพนักงานใหม่ลงชีต Staff จริง ๆ (ใช้ร่วมกันทั้งจากแอดมินเพิ่มตรง ๆ และจากการอนุมัติคำขอ)
function insertStaffRow_(record) {
  if (!record.empId || !record.nameEn) {
    throw new Error('กรุณากรอกรหัสพนักงานและชื่อ-นามสกุล');
  }
  record = normalizeRecord_(record);
  const sheet = getStaffSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(record.empId)) {
      throw new Error('รหัสพนักงานนี้มีอยู่แล้วในระบบ');
    }
  }
  const existingNos = data.slice(1).filter(r => r[1] && r[0]).map(r => Number(r[0]) || 0);
  const nextNo = existingNos.length ? Math.max.apply(null, existingNos) + 1 : 1;
  sheet.appendRow([
    nextNo, record.empId, record.nameEn, record.nameTh || '', record.nickname || '',
    record.dob || '', record.startDate || '', record.phone || '', record.emergencyPhone || '',
    record.email || '', record.address || '', record.status || 'active', record.resignedDate || ''
  ]);
}

// เพิ่มพนักงานใหม่ (ต้องมีสิทธิ์แอดมิน)
function addStaff(token, record) {
  requireAdmin_(token);
  insertStaffRow_(record);
  return getStaffList(token);
}

// แก้ไขข้อมูลพนักงาน (ต้องมีสิทธิ์แอดมิน)
function updateStaff(token, empId, record) {
  requireAdmin_(token);
  record = normalizeRecord_(record);
  const sheet = getStaffSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(empId)) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, 2, 1, 12).setValues([[
        record.empId, record.nameEn, record.nameTh || '', record.nickname || '',
        record.dob || '', record.startDate || '', record.phone || '', record.emergencyPhone || '',
        record.email || '', record.address || '', record.status || 'active', record.resignedDate || ''
      ]]);
      return getStaffList(token);
    }
  }
  throw new Error('ไม่พบข้อมูลพนักงานรหัสนี้');
}

// ลบข้อมูลพนักงาน (ต้องมีสิทธิ์แอดมิน)
function deleteStaff(token, empId) {
  requireAdmin_(token);
  const sheet = getStaffSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(empId)) {
      sheet.deleteRow(i + 1);
      return getStaffList(token);
    }
  }
  throw new Error('ไม่พบข้อมูลพนักงานรหัสนี้');
}

// ---------------------------------------------------------------------
// แบบฟอร์มสำหรับพนักงานทั่วไป (ไม่ต้องล็อกอิน) - ส่งข้อมูลเข้าคิวรออนุมัติ
// ---------------------------------------------------------------------

// ผู้ใช้ทั่วไปกรอกฟอร์มส่งเข้ามา -> เก็บไว้ในชีต "Pending" รอแอดมินตรวจสอบ
function submitPendingStaff(record) {
  if (!record.empId || !record.nameEn) {
    throw new Error('กรุณากรอกรหัสพนักงานและชื่อ-นามสกุล');
  }
  // กันซ้ำกับข้อมูลที่อนุมัติแล้ว
  const staffSheet = getStaffSheet_();
  const staffData = staffSheet.getDataRange().getValues();
  for (let i = 1; i < staffData.length; i++) {
    if (String(staffData[i][1]) === String(record.empId)) {
      throw new Error('รหัสพนักงานนี้มีข้อมูลอยู่ในระบบแล้ว');
    }
  }
  // กันซ้ำกับคำขอที่ส่งไปแล้วแต่ยังไม่อนุมัติ
  const pendingSheet = getPendingSheet_();
  const pendingData = pendingSheet.getDataRange().getValues();
  for (let i = 1; i < pendingData.length; i++) {
    if (String(pendingData[i][1]) === String(record.empId)) {
      throw new Error('มีการส่งข้อมูลรหัสพนักงานนี้รออนุมัติอยู่แล้ว');
    }
  }
  pendingSheet.appendRow([
    new Date(), record.empId, record.nameEn, record.nameTh || '', record.nickname || '',
    record.dob || '', record.startDate || '', record.phone || '', record.emergencyPhone || '',
    record.email || '', record.address || ''
  ]);
  return { success: true };
}

function pendingRowToObj_(row, rowNum) {
  return {
    row: rowNum,
    timestamp: fmtDate_(row[0]),
    empId: row[1],
    nameEn: row[2],
    nameTh: row[3],
    nickname: row[4],
    dob: fmtDate_(row[5]),
    startDate: fmtDate_(row[6]),
    phone: row[7] ? String(row[7]) : '',
    emergencyPhone: row[8] ? String(row[8]) : '',
    email: row[9] || '',
    address: row[10] || '',
  };
}

// รายการคำขอที่รออนุมัติ (ต้องมีสิทธิ์แอดมิน)
function getPendingList(token) {
  requireAdmin_(token);
  const sheet = getPendingSheet_();
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    out.push(pendingRowToObj_(data[i], i + 1));
  }
  return out;
}

// นับจำนวนคำขอที่รออนุมัติ (ต้องมีสิทธิ์แอดมิน) - ใช้แสดง badge
function getPendingCount(token) {
  return getPendingList(token).length;
}

// อนุมัติคำขอ -> ย้ายเข้าชีต Staff แล้วลบออกจาก Pending (ต้องมีสิทธิ์แอดมิน)
function approvePending(token, empId) {
  requireAdmin_(token);
  const sheet = getPendingSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(empId)) {
      const obj = pendingRowToObj_(data[i], i + 1);
      insertStaffRow_({
        empId: obj.empId, nameEn: obj.nameEn, nameTh: obj.nameTh, nickname: obj.nickname,
        dob: obj.dob, startDate: obj.startDate, phone: obj.phone, emergencyPhone: obj.emergencyPhone,
        email: obj.email, address: obj.address, status: 'active'
      });
      sheet.deleteRow(i + 1);
      return { staffList: getStaffList(token), pendingList: getPendingList(token) };
    }
  }
  throw new Error('ไม่พบคำขอนี้ อาจถูกดำเนินการไปแล้ว');
}

// ปฏิเสธคำขอ -> ลบออกจาก Pending (ต้องมีสิทธิ์แอดมิน)
function rejectPending(token, empId) {
  requireAdmin_(token);
  const sheet = getPendingSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(empId)) {
      sheet.deleteRow(i + 1);
      return getPendingList(token);
    }
  }
  throw new Error('ไม่พบคำขอนี้ อาจถูกดำเนินการไปแล้ว');
}
