/**
 * ระบบอภิบาลและการพินิจ
 * Server-side (Google Apps Script)
 */
const CONFIG = {
  SPREADSHEET_ID: '1Di0chS8rRVuErWH2IEN28THVLCSjGSdhC2iuXhuusAs',
  DRIVE_FOLDER_ID: '1_34bRCXbVE2f-ko2P4VmuzVqM7YYW-dA',
  TIMEZONE: Session.getScriptTimeZone() || 'Asia/Bangkok',
  SHEETS: {
    youth: ['youthId', 'fullName', 'gradeLevel', 'vocationUnit', 'advisorTeacher', 'createdAt', 'updatedAt'],
    youth_documents: ['docId', 'youthId', 'docCategory', 'docType', 'itemName', 'fileId', 'fileName', 'fileUrl', 'uploadedAt', 'updatedAt', 'note', 'sortOrder', 'status'],
    activity_schedule: ['activityId', 'title', 'eventType', 'personName', 'roleType', 'location', 'startDate', 'endDate', 'startTime', 'endTime', 'allDay', 'description', 'status', 'createdAt', 'updatedAt', 'youthId', 'youthNameSnapshot', 'advisorTeacherSnapshot', 'courtName', 'courtMode', 'hospitalName', 'staffNames', 'taskOwner', 'eventDate', 'reminderDay', 'reminderWeek', 'reminderMonth', 'reminderQuarter', 'createdBy'],
    visit_rights: ['visitRightId', 'youthId', 'gradeLevel', 'allowedVisits', 'usedVisits', 'remainingVisits'],
    visit_bookings: ['bookingId', 'youthId', 'bookingDate', 'visitorName', 'note', 'createdAt'],
    discipline_logs: ['disciplineId', 'youthId', 'disciplineDate', 'title', 'detail', 'createdAt'],
    inventory: ['itemId', 'itemName', 'qtyTotal', 'qtyRemaining', 'unit', 'updatedAt'],
    settings_vocation_units: ['id', 'name', 'status', 'createdAt', 'updatedAt'],
    settings_advisor_teachers: ['id', 'name', 'status', 'createdAt', 'updatedAt'],
    holiday_cache: ['holidayId', 'holidayDate', 'holidayName', 'holidayType', 'sourceCalendarId', 'updatedAt', 'status'],
    visit_normal_quota_rule: ['ruleId', 'gradeLevel', 'allowedPerMonth', 'status', 'createdAt', 'updatedAt'],
    visit_special_quota_rule: ['specialRuleId', 'ruleName', 'gradeLevelScope', 'extraQuota', 'startDate', 'endDate', 'note', 'status', 'createdAt', 'updatedAt'],
    visit_quota_rule: ['ruleId', 'gradeLevel', 'allowedPerMonth', 'status', 'createdAt', 'updatedAt'],
    visit_special_rule: ['specialRuleId', 'ruleName', 'gradeLevelScope', 'extraQuota', 'startDate', 'endDate', 'note', 'status', 'createdAt', 'updatedAt'],
    visit_booking: ['bookingId', 'youthId', 'bookingDate', 'quotaTypeUsed', 'specialRuleId', 'note', 'importedFrom', 'createdAt', 'updatedAt', 'status']
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Client').evaluate().setTitle('ระบบงานอภิบาลและการพินิจ').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initSystem() {
  const ss = getSpreadsheet();
  Object.keys(CONFIG.SHEETS).forEach(function(name) { upsertSheetWithHeader_(ss, name, CONFIG.SHEETS[name]); });
  const folderInfo = checkDriveFolder_();
  return { ok: true, message: 'ระบบพร้อมใช้งาน', spreadsheetId: CONFIG.SPREADSHEET_ID, folder: folderInfo, initializedAt: new Date().toISOString() };
}

function getSpreadsheet() { return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); }
function getDriveFolder() { return DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID); }

function getSystemInfo() {
  return { appName: 'ระบบงานอภิบาลและการพินิจ', version: '0.3.0-document-module', spreadsheetId: CONFIG.SPREADSHEET_ID, folderId: CONFIG.DRIVE_FOLDER_ID, now: new Date().toISOString() };
}

/** -------- Youth existing module -------- */
function ensureYouthSheet_() {
  var sh = upsertSheetWithHeader_(getSpreadsheet(), 'youth', CONFIG.SHEETS.youth);
  sh.getRange(2, 3, Math.max(sh.getMaxRows()-1, 1), 1).setNumberFormat('@');
  return sh;
}
function ensureSettingsSheets_() {
  const ss = getSpreadsheet();
  upsertSheetWithHeader_(ss, 'settings_vocation_units', CONFIG.SHEETS.settings_vocation_units);
  upsertSheetWithHeader_(ss, 'settings_advisor_teachers', CONFIG.SHEETS.settings_advisor_teachers);
}

function getYouthList(params) {
  ensureYouthSheet_();
  ensureSettingsSheets_();
  params = params || {};
  const query = normalizeText_(params.query || '');
  const limit = params.limit === 'all' ? 'all' : Number(params.limit || 25);
  const sh = getSpreadsheet().getSheetByName('youth');
  const data = sh.getDataRange().getValues();
  const idx = headerMap_(data[0] || CONFIG.SHEETS.youth);

  let items = data.slice(1).filter(function(r) { return safeString_(r[idx.youthId]); }).map(function(r) {
    return {
      youthId: safeString_(r[idx.youthId]),
      fullName: safeString_(r[idx.fullName]),
      gradeLevel: normalizeGradeLevel_(r[idx.gradeLevel]),
      vocationUnit: safeString_(r[idx.vocationUnit]),
      advisorTeacher: safeString_(r[idx.advisorTeacher]),
      createdAt: safeString_(r[idx.createdAt]),
      updatedAt: safeString_(r[idx.updatedAt])
    };
  });

  if (query) {
    items = items.filter(function(it) {
      return [it.fullName, it.gradeLevel, it.vocationUnit, it.advisorTeacher].some(function(v) { return normalizeText_(v).indexOf(query) > -1; });
    });
  }

  items.sort(function(a, b) { return dateNum_(b.updatedAt || b.createdAt) - dateNum_(a.updatedAt || a.createdAt); });
  const total = items.length;
  if (limit !== 'all' && Number.isFinite(limit)) items = items.slice(0, limit);
  return { items: items, total: total, shown: items.length };
}

function saveYouth(data) {
  ensureYouthSheet_();
  validateYouthPayload_(data);
  const sh = getSpreadsheet().getSheetByName('youth');
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = headerMap_(header);
  const now = formatDateTime_(new Date());
  const row = new Array(header.length).fill('');
  row[idx.youthId] = generateYouthId_();
  row[idx.fullName] = data.fullName.trim();
  row[idx.gradeLevel] = normalizeGradeLevel_(data.gradeLevel);
  row[idx.vocationUnit] = safeString_(data.vocationUnit).trim();
  row[idx.advisorTeacher] = safeString_(data.advisorTeacher).trim();
  row[idx.createdAt] = now;
  row[idx.updatedAt] = now;
  sh.appendRow(row);
  return { ok: true, message: 'เพิ่มข้อมูลสำเร็จ', youthId: row[idx.youthId] };
}

function updateYouth(data) {
  ensureYouthSheet_();
  if (!data || !safeString_(data.youthId)) throw new Error('ไม่พบ youthId');
  validateYouthPayload_(data);
  const sh = getSpreadsheet().getSheetByName('youth');
  const values = sh.getDataRange().getValues();
  const idx = headerMap_(values[0]);
  const id = safeString_(data.youthId).trim();
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) if (safeString_(values[i][idx.youthId]) === id) { rowIndex = i + 1; break; }
  if (rowIndex === -1) throw new Error('ไม่พบข้อมูลที่ต้องการแก้ไข');
  sh.getRange(rowIndex, idx.fullName + 1).setValue(data.fullName.trim());
  sh.getRange(rowIndex, idx.gradeLevel + 1).setNumberFormat('@').setValue(normalizeGradeLevel_(data.gradeLevel));
  sh.getRange(rowIndex, idx.vocationUnit + 1).setValue(safeString_(data.vocationUnit).trim());
  sh.getRange(rowIndex, idx.advisorTeacher + 1).setValue(safeString_(data.advisorTeacher).trim());
  sh.getRange(rowIndex, idx.updatedAt + 1).setValue(formatDateTime_(new Date()));
  return { ok: true, message: 'แก้ไขข้อมูลสำเร็จ', youthId: id };
}

function deleteYouthById(youthId) {
  const id = safeString_(youthId).trim();
  if (!id) throw new Error('ไม่พบ youthId');
  return deleteYouthByIds([id]);
}
function deleteYouthByIds(youthIds) {
  ensureYouthSheet_();
  if (!youthIds || !youthIds.length) throw new Error('ไม่มีรายการที่ต้องการลบ');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const unique = {};
    youthIds.forEach(function(id) { const v = safeString_(id).trim(); if (v) unique[v] = true; });
    const ids = Object.keys(unique);
    if (!ids.length) throw new Error('ไม่มี youthId ที่ถูกต้อง');
    const total = createDeleteSummary_();
    ids.forEach(function(id) {
      const one = deleteYouthCascade_(id);
      mergeDeleteSummary_(total, one);
    });
    total.ok = true;
    total.message = buildDeleteSummaryMessage_(total);
    return total;
  } finally {
    lock.releaseLock();
  }
}

function deleteAllYouthRows() {
  ensureYouthSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getSpreadsheet().getSheetByName('youth');
    const values = sh.getDataRange().getValues();
    if (values.length <= 1) {
      const emptySummary = createDeleteSummary_();
      emptySummary.ok = true;
      emptySummary.message = buildDeleteSummaryMessage_(emptySummary);
      return emptySummary;
    }
    const idx = headerMap_(values[0]);
    const unique = {};
    for (let i = 1; i < values.length; i++) {
      const id = safeString_(values[i][idx.youthId]).trim();
      if (id) unique[id] = true;
    }
    const ids = Object.keys(unique);
    const total = createDeleteSummary_();
    ids.forEach(function(id) {
      const one = deleteYouthCascade_(id);
      mergeDeleteSummary_(total, one);
    });
    total.ok = true;
    total.message = buildDeleteSummaryMessage_(total);
    return total;
  } finally {
    lock.releaseLock();
  }
}
function deleteAllYouth() { return deleteAllYouthRows(); }

function deleteYouthCascade_(youthId) {
  const id = safeString_(youthId).trim();
  const summary = createDeleteSummary_();
  if (!id) return summary;
  summary.youthIds.push(id);

  const docResult = deleteYouthDocumentFiles_(id);
  summary.docsDeleted += Number(docResult.rowsDeleted || 0);
  summary.driveDeleted += Number(docResult.driveDeleted || 0);
  summary.driveFailed += Number(docResult.driveFailed || 0);
  summary.driveSkipped += Number(docResult.driveSkipped || 0);

  const sheets = getRelatedSheetsWithYouthId_();
  sheets.forEach(function(item) {
    const r = deleteYouthRelatedRows_(item.sheetName, id, item.options);
    summary.rowsDeletedBySheet[item.sheetName] = (summary.rowsDeletedBySheet[item.sheetName] || 0) + Number(r.rowsDeleted || 0);
    if (item.sheetName === 'youth') summary.youthDeleted += Number(r.rowsDeleted || 0);
    else if (item.sheetName === 'activity_schedule') summary.activitiesDeleted += Number(r.rowsDeleted || 0);
    else if (item.sheetName === 'visit_booking' || item.sheetName === 'visit_bookings') summary.visitDeleted += Number(r.rowsDeleted || 0);
    else if (item.sheetName !== 'youth_documents') summary.otherDeleted += Number(r.rowsDeleted || 0);
    summary.errors = summary.errors.concat(r.errors || []);
  });
  return summary;
}
function deleteYouthRelatedRows_(sheetName, youthId, options) {
  const out = { sheetName: sheetName, rowsDeleted: 0, errors: [] };
  const sh = getSpreadsheet().getSheetByName(sheetName);
  if (!sh) return out;
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return out;
  const idx = headerMap_(values[0]);
  if (idx.youthId === undefined) return out;

  const rows = findRowsByHeaderValue_(sh, 'youthId', youthId);
  if (!rows.length) return out;
  const softDelete = !!(options && options.softDelete && idx.status !== undefined);
  const now = formatDateTime_(new Date());
  const updatedAtIdx = idx.updatedAt;
  if (softDelete) {
    rows.forEach(function(row) {
      try {
        sh.getRange(row, idx.status + 1).setValue('deleted');
        if (updatedAtIdx !== undefined) sh.getRange(row, updatedAtIdx + 1).setValue(now);
        out.rowsDeleted++;
      } catch (err) { out.errors.push('sheet=' + sheetName + ' row=' + row + ' ' + (err && err.message ? err.message : err)); }
    });
  } else {
    for (let i = rows.length - 1; i >= 0; i--) {
      const row = rows[i];
      try {
        if (row > 1) {
          sh.deleteRow(row);
          out.rowsDeleted++;
        }
      } catch (err) { out.errors.push('sheet=' + sheetName + ' row=' + row + ' ' + (err && err.message ? err.message : err)); }
    }
  }
  return out;
}
function deleteYouthDocumentFiles_(youthId) {
  const out = { rowsDeleted: 0, driveDeleted: 0, driveFailed: 0, driveSkipped: 0, errors: [] };
  const sh = getSpreadsheet().getSheetByName('youth_documents');
  if (!sh) return out;
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return out;
  const idx = headerMap_(values[0]);
  if (idx.youthId === undefined) return out;
  const rows = findRowsByHeaderValue_(sh, 'youthId', youthId);
  if (!rows.length) return out;

  const seen = {};
  rows.forEach(function(row) {
    const v = values[row - 1];
    const fileId = idx.fileId === undefined ? '' : safeString_(v[idx.fileId]).trim();
    if (fileId) seen[fileId] = true;
  });
  Object.keys(seen).forEach(function(fileId) {
    const r = deleteDriveFileSafe_(fileId);
    if (r.deleted) out.driveDeleted++;
    else if (r.skipped) out.driveSkipped++;
    else out.driveFailed++;
    if (r.error) out.errors.push('drive fileId=' + fileId + ' ' + r.error);
  });

  const modeSoft = idx.status !== undefined;
  const now = formatDateTime_(new Date());
  if (modeSoft) {
    rows.forEach(function(row) {
      try {
        sh.getRange(row, idx.status + 1).setValue('deleted');
        if (idx.updatedAt !== undefined) sh.getRange(row, idx.updatedAt + 1).setValue(now);
        out.rowsDeleted++;
      } catch (err) { out.errors.push('youth_documents row=' + row + ' ' + (err && err.message ? err.message : err)); }
    });
  } else {
    for (let i = rows.length - 1; i >= 0; i--) {
      const row = rows[i];
      try {
        if (row > 1) {
          sh.deleteRow(row);
          out.rowsDeleted++;
        }
      } catch (err) { out.errors.push('youth_documents row=' + row + ' ' + (err && err.message ? err.message : err)); }
    }
  }
  return out;
}
function deleteDriveFileSafe_(fileId) {
  const id = safeString_(fileId).trim();
  if (!id) return { deleted: false, failed: false, skipped: true };
  try {
    Drive.Files.remove(id, { supportsAllDrives: true });
    return { deleted: true, failed: false, skipped: false };
  } catch (err) {
    const msg = safeString_(err && err.message ? err.message : err);
    const notFound = msg.indexOf('File not found') > -1 || msg.indexOf('notFound') > -1 || msg.indexOf('404') > -1;
    if (notFound) return { deleted: false, failed: false, skipped: true, error: msg };
    return { deleted: false, failed: true, skipped: false, error: msg };
  }
}
function getRelatedSheetsWithYouthId_() {
  const ss = getSpreadsheet();
  return ss.getSheets().map(function(sh) {
    const name = sh.getName();
    const values = sh.getDataRange().getValues();
    if (!values.length) return null;
    const idx = headerMap_(values[0]);
    if (idx.youthId === undefined) return null;
    if (name === 'youth_documents') return { sheetName: name, options: { skip: true } };
    if (name === 'youth') return { sheetName: name, options: { softDelete: false } };
    if (name === 'activity_schedule') return { sheetName: name, options: { softDelete: idx.status !== undefined } };
    if (name === 'visit_booking' || name === 'visit_bookings') return { sheetName: name, options: { softDelete: idx.status !== undefined } };
    return { sheetName: name, options: { softDelete: idx.status !== undefined } };
  }).filter(function(it) { return !!it && !(it.options && it.options.skip); });
}
function findRowsByHeaderValue_(sheet, headerName, value) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const idx = headerMap_(values[0]);
  const col = idx[headerName];
  if (col === undefined) return [];
  const target = safeString_(value);
  const rows = [];
  for (let i = 1; i < values.length; i++) if (safeString_(values[i][col]) === target) rows.push(i + 1);
  return rows;
}
function createDeleteSummary_() {
  return { ok: false, message: '', youthIds: [], youthDeleted: 0, docsDeleted: 0, driveDeleted: 0, driveFailed: 0, driveSkipped: 0, visitDeleted: 0, activitiesDeleted: 0, otherDeleted: 0, rowsDeletedBySheet: {}, errors: [] };
}
function mergeDeleteSummary_(target, source) {
  target.youthDeleted += Number(source.youthDeleted || 0);
  target.docsDeleted += Number(source.docsDeleted || 0);
  target.driveDeleted += Number(source.driveDeleted || 0);
  target.driveFailed += Number(source.driveFailed || 0);
  target.driveSkipped += Number(source.driveSkipped || 0);
  target.visitDeleted += Number(source.visitDeleted || 0);
  target.activitiesDeleted += Number(source.activitiesDeleted || 0);
  target.otherDeleted += Number(source.otherDeleted || 0);
  target.youthIds = target.youthIds.concat(source.youthIds || []);
  target.errors = target.errors.concat(source.errors || []);
  const bySheet = source.rowsDeletedBySheet || {};
  Object.keys(bySheet).forEach(function(name) {
    target.rowsDeletedBySheet[name] = (target.rowsDeletedBySheet[name] || 0) + Number(bySheet[name] || 0);
  });
}
function buildDeleteSummaryMessage_(s) {
  return [
    'ลบข้อมูลเด็กและเยาวชน ' + Number(s.youthDeleted || 0) + ' รายการ',
    'ลบเอกสาร/metadata ' + Number(s.docsDeleted || 0) + ' รายการ',
    'ลบไฟล์ใน Drive สำเร็จ ' + Number(s.driveDeleted || 0) + ' ไฟล์ (ไม่สำเร็จ ' + Number(s.driveFailed || 0) + ', ข้าม ' + Number(s.driveSkipped || 0) + ')',
    'ลบรายการจองเยี่ยม ' + Number(s.visitDeleted || 0) + ' รายการ',
    'ลบกิจกรรมที่เกี่ยวข้อง ' + Number(s.activitiesDeleted || 0) + ' รายการ',
    'ลบข้อมูลอื่นที่เกี่ยวข้อง ' + Number(s.otherDeleted || 0) + ' รายการ',
    'ข้อผิดพลาด ' + Number((s.errors || []).length) + ' รายการ'
  ].join('\\n');
}

function getVocationUnits() { ensureSettingsSheets_(); return getActiveSettingsList_('settings_vocation_units'); }
function getAdvisorTeachers() { ensureSettingsSheets_(); return getActiveSettingsList_('settings_advisor_teachers'); }

function importYouthFromExcel(base64Data, fileName, mimeType) {
  ensureYouthSheet_();
  if (!base64Data || !fileName) throw new Error('ข้อมูลไฟล์ไม่ครบถ้วน');
  const ext = (fileName.split('.').pop() || '').toLowerCase();
  if (['xlsx', 'xls', 'csv'].indexOf(ext) === -1) throw new Error('รองรับเฉพาะไฟล์ .xlsx, .xls, .csv');

  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'application/octet-stream', fileName);
  const rows = ext === 'csv' ? Utilities.parseCsv(blob.getDataAsString('UTF-8')) : convertExcelToSheetRows_(blob, fileName).rows;
  if (!rows || !rows.length) throw new Error('ไม่พบข้อมูลในไฟล์');

  const m = mapImportHeader_(rows[0]);
  const sh = getSpreadsheet().getSheetByName('youth');
  const values = sh.getDataRange().getValues();
  const idx = headerMap_(values[0]);
  const byName = {};
  for (let i = 1; i < values.length; i++) {
    const k = normalizeText_(values[i][idx.fullName]); if (k) byName[k] = i + 1;
  }

  let inserted = 0, updated = 0, skipped = 0;
  const now = formatDateTime_(new Date());
  rows.slice(1).forEach(function(r) {
    const fullName = safeString_(r[m.fullName]).trim();
    const gradeLevel = normalizeGradeLevel_(r[m.gradeLevel]);
    const vocationUnit = safeString_(r[m.vocationUnit]).trim();
    const advisorTeacher = safeString_(r[m.advisorTeacher]).trim();
    if (!fullName && !gradeLevel && !vocationUnit && !advisorTeacher) { skipped++; return; }
    if (!fullName) { skipped++; return; }
    const k = normalizeText_(fullName);
    if (byName[k]) {
      const row = byName[k];
      sh.getRange(row, idx.gradeLevel + 1).setValue(gradeLevel);
      sh.getRange(row, idx.vocationUnit + 1).setValue(vocationUnit);
      sh.getRange(row, idx.advisorTeacher + 1).setValue(advisorTeacher);
      sh.getRange(row, idx.updatedAt + 1).setValue(now);
      updated++;
    } else {
      const out = new Array(values[0].length).fill('');
      out[idx.youthId] = generateYouthId_();
      out[idx.fullName] = fullName;
      out[idx.gradeLevel] = gradeLevel;
      out[idx.vocationUnit] = vocationUnit;
      out[idx.advisorTeacher] = advisorTeacher;
      out[idx.createdAt] = now;
      out[idx.updatedAt] = now;
      sh.appendRow(out);
      byName[k] = sh.getLastRow();
      inserted++;
    }
  });
  return { ok: true, message: 'นำเข้าข้อมูลสำเร็จ', summary: { inserted: inserted, updated: updated, skipped: skipped } };
}

/** -------- Document module -------- */
function ensureYouthDocumentsSheet_() {
  return upsertSheetWithHeader_(getSpreadsheet(), 'youth_documents', CONFIG.SHEETS.youth_documents);
}

function getYouthBasicMap_() {
  ensureYouthSheet_();
  const sh = getSpreadsheet().getSheetByName('youth');
  const values = sh.getDataRange().getValues();
  const idx = headerMap_(values[0]);
  return values.slice(1).filter(function(r) { return safeString_(r[idx.youthId]); }).map(function(r) {
    return {
      youthId: safeString_(r[idx.youthId]),
      fullName: safeString_(r[idx.fullName]),
      advisorTeacher: safeString_(r[idx.advisorTeacher]),
      updatedAt: safeString_(r[idx.updatedAt])
    };
  });
}

function getYouthDocumentTable(params) {
  ensureYouthDocumentsSheet_();
  const query = normalizeText_((params && params.query) || '');
  const limit = params && params.limit === 'all' ? 'all' : Number((params && params.limit) || 25);

  let youthList = getYouthBasicMap_();
  if (query) {
    youthList = youthList.filter(function(y) {
      return normalizeText_(y.fullName).indexOf(query) > -1 || normalizeText_(y.advisorTeacher).indexOf(query) > -1;
    });
  }

  const latestMap = getLatestDocumentMap_();

  let rows = youthList.map(function(y) {
    const aKey = y.youthId + '__ASSESSMENT';
    const pKey = y.youthId + '__PT1';
    const a = latestMap[aKey] || null;
    const p = latestMap[pKey] || null;
    const lastUpdated = maxDateText_(y.updatedAt, a && a.updatedAt, p && p.updatedAt);
    return {
      youthId: y.youthId,
      fullName: y.fullName,
      advisorTeacher: y.advisorTeacher,
      hasAssessment: !!a,
      hasPt1: !!p,
      latestUpdatedAt: lastUpdated
    };
  });

  rows.sort(function(a, b) { return dateNum_(b.latestUpdatedAt) - dateNum_(a.latestUpdatedAt); });
  const total = rows.length;
  if (limit !== 'all' && Number.isFinite(limit)) rows = rows.slice(0, limit);
  return { items: rows, total: total, shown: rows.length };
}

function getLatestDocumentMap_() {
  ensureYouthDocumentsSheet_();
  const sh = getSpreadsheet().getSheetByName('youth_documents');
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return {};
  const idx = headerMap_(values[0]);
  const map = {};
  values.slice(1).forEach(function(r) {
    const status = normalizeText_(r[idx.status] || 'active');
    if (status !== 'active') return;
    const key = safeString_(r[idx.youthId]) + '__' + safeString_(r[idx.docCategory]);
    const current = map[key];
    const item = {
      docId: safeString_(r[idx.docId]),
      youthId: safeString_(r[idx.youthId]),
      docCategory: safeString_(r[idx.docCategory]),
      updatedAt: safeString_(r[idx.updatedAt])
    };
    if (!current || dateNum_(item.updatedAt) >= dateNum_(current.updatedAt)) map[key] = item;
  });
  return map;
}

function getYouthDocumentDetail(youthId, category) {
  ensureYouthDocumentsSheet_();
  if (!youthId || !category) throw new Error('ข้อมูลไม่ครบ');
  const sh = getSpreadsheet().getSheetByName('youth_documents');
  const values = sh.getDataRange().getValues();
  const idx = headerMap_(values[0]);
  const items = values.slice(1).map(function(r, i) {
    return {
      rowIndex: i + 2,
      docId: safeString_(r[idx.docId]),
      youthId: safeString_(r[idx.youthId]),
      docCategory: safeString_(r[idx.docCategory]),
      docType: safeString_(r[idx.docType]),
      itemName: safeString_(r[idx.itemName]),
      fileId: safeString_(r[idx.fileId]),
      fileName: safeString_(r[idx.fileName]),
      fileUrl: safeString_(r[idx.fileUrl]),
      uploadedAt: safeString_(r[idx.uploadedAt]),
      updatedAt: safeString_(r[idx.updatedAt]),
      note: safeString_(r[idx.note]),
      sortOrder: Number(r[idx.sortOrder] || 0),
      status: safeString_(r[idx.status] || 'active')
    };
  }).filter(function(it) {
    return it.youthId === youthId && it.docCategory === category && normalizeText_(it.status) === 'active';
  });

  items.sort(function(a, b) {
    if (a.sortOrder !== b.sortOrder) return a.sortOrder - b.sortOrder;
    return dateNum_(b.updatedAt || b.uploadedAt) - dateNum_(a.updatedAt || a.uploadedAt);
  });

  const youth = getYouthBasicMap_().find(function(y) { return y.youthId === youthId; }) || { youthId: youthId, fullName: '-', advisorTeacher: '-' };
  return { youth: youth, category: category, items: items };
}

function saveAssessmentDocument(payload) {
  ensureYouthDocumentsSheet_();
  validateDocumentPayload_(payload, 'ASSESSMENT');
  const file = payload.file;
  if (!file || !file.base64Data) throw new Error('กรุณาแนบไฟล์ประเมินระดับการควบคุม');

  markRowsDeleted_(payload.youthId, 'ASSESSMENT', ['CONTROL_ASSESSMENT']);
  const uploaded = uploadDriveFile_(file.base64Data, file.fileName, file.mimeType || 'application/octet-stream');
  appendDocumentRow_({
    youthId: payload.youthId,
    docCategory: 'ASSESSMENT',
    docType: 'CONTROL_ASSESSMENT',
    itemName: 'ประเมินระดับการควบคุม',
    fileId: uploaded.fileId,
    fileName: uploaded.fileName,
    fileUrl: uploaded.fileUrl,
    note: safeString_(payload.note),
    sortOrder: 1,
    status: 'active'
  });

  return { ok: true, message: 'บันทึกเอกสารประเมินสำเร็จ' };
}

function updateAssessmentDocument(payload) {
  ensureYouthDocumentsSheet_();
  validateDocumentPayload_(payload, 'ASSESSMENT');
  const detail = getYouthDocumentDetail(payload.youthId, 'ASSESSMENT').items;
  const current = detail.find(function(it) { return it.docType === 'CONTROL_ASSESSMENT'; });
  if (!current && !(payload.file && payload.file.base64Data)) {
    throw new Error('ไม่พบไฟล์เดิม และยังไม่ได้แนบไฟล์ใหม่');
  }

  let fileId = current ? current.fileId : '';
  let fileName = current ? current.fileName : '';
  let fileUrl = current ? current.fileUrl : '';

  if (payload.file && payload.file.base64Data) {
    const uploaded = uploadDriveFile_(payload.file.base64Data, payload.file.fileName, payload.file.mimeType || 'application/octet-stream');
    fileId = uploaded.fileId;
    fileName = uploaded.fileName;
    fileUrl = uploaded.fileUrl;
  }

  markRowsDeleted_(payload.youthId, 'ASSESSMENT', ['CONTROL_ASSESSMENT']);
  appendDocumentRow_({
    youthId: payload.youthId,
    docCategory: 'ASSESSMENT',
    docType: 'CONTROL_ASSESSMENT',
    itemName: 'ประเมินระดับการควบคุม',
    fileId: fileId,
    fileName: fileName,
    fileUrl: fileUrl,
    note: safeString_(payload.note),
    sortOrder: 1,
    status: 'active'
  });

  return { ok: true, message: 'แก้ไขเอกสารประเมินสำเร็จ' };
}

function savePt1Documents(payload) {
  ensureYouthDocumentsSheet_();
  validateDocumentPayload_(payload, 'PT1');
  const items = sanitizePt1Items_(payload.items || [], false);
  markRowsDeleted_(payload.youthId, 'PT1');
  items.forEach(function(it, idx) {
    const uploaded = uploadDriveFile_(it.file.base64Data, it.file.fileName, it.file.mimeType || 'application/octet-stream');
    appendDocumentRow_({
      youthId: payload.youthId,
      docCategory: 'PT1',
      docType: it.docType,
      itemName: it.itemName,
      fileId: uploaded.fileId,
      fileName: uploaded.fileName,
      fileUrl: uploaded.fileUrl,
      note: safeString_(payload.note),
      sortOrder: idx + 1,
      status: 'active'
    });
  });
  return { ok: true, message: 'บันทึกเอกสาร ปท.1 สำเร็จ' };
}

function updatePt1Documents(payload) {
  ensureYouthDocumentsSheet_();
  validateDocumentPayload_(payload, 'PT1');
  const items = sanitizePt1Items_(payload.items || [], true);
  markRowsDeleted_(payload.youthId, 'PT1');
  items.forEach(function(it, idx) {
    let fileId = safeString_(it.existingFileId);
    let fileName = safeString_(it.existingFileName);
    let fileUrl = safeString_(it.existingFileUrl);
    if (it.file && it.file.base64Data) {
      const uploaded = uploadDriveFile_(it.file.base64Data, it.file.fileName, it.file.mimeType || 'application/octet-stream');
      fileId = uploaded.fileId;
      fileName = uploaded.fileName;
      fileUrl = uploaded.fileUrl;
    }
    if (!fileId) throw new Error('รายการ ' + it.itemName + ' ยังไม่มีไฟล์');

    appendDocumentRow_({
      youthId: payload.youthId,
      docCategory: 'PT1',
      docType: it.docType,
      itemName: it.itemName,
      fileId: fileId,
      fileName: fileName,
      fileUrl: fileUrl,
      note: safeString_(payload.note),
      sortOrder: idx + 1,
      status: 'active'
    });
  });
  return { ok: true, message: 'แก้ไขเอกสาร ปท.1 สำเร็จ' };
}

function uploadDriveFile_(base64Data, fileName, mimeType) {
  if (!base64Data || !fileName) throw new Error('ไฟล์ไม่สมบูรณ์');
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'application/octet-stream', fileName);
  const uploaded = uploadToDriveV3_(blob, fileName, mimeType || 'application/octet-stream');
  return {
    fileId: uploaded.id,
    fileName: uploaded.name || fileName,
    fileUrl: buildDriveFileUrl_(uploaded.id)
  };
}

function buildDriveFileUrl_(fileId) {
  return 'https://drive.google.com/file/d/' + encodeURIComponent(fileId) + '/view';
}

function getFileMetadata(fileId) {
  if (!fileId) throw new Error('กรุณาระบุ fileId');
  const file = Drive.Files.get(fileId, { fields: 'id,name,mimeType,size,createdTime,webViewLink,webContentLink,parents' });
  return {
    fileId: file.id,
    fileName: file.name,
    mimeType: file.mimeType,
    size: file.size,
    createdTime: file.createdTime,
    fileUrl: file.webViewLink || file.webContentLink,
    parents: file.parents || []
  };
}

function checkDriveFolder_() {
  const folder = Drive.Files.get(CONFIG.DRIVE_FOLDER_ID, { fields: 'id,name,mimeType,webViewLink' });
  if (!folder || folder.mimeType !== 'application/vnd.google-apps.folder') throw new Error('Folder ID ไม่ถูกต้อง');
  return { id: folder.id, name: folder.name, webViewLink: folder.webViewLink };
}

/** -------- internals -------- */
function appendDocumentRow_(doc) {
  const sh = ensureYouthDocumentsSheet_();
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = headerMap_(header);
  const now = formatDateTime_(new Date());
  const row = new Array(header.length).fill('');
  row[idx.docId] = generateDocId_();
  row[idx.youthId] = doc.youthId;
  row[idx.docCategory] = doc.docCategory;
  row[idx.docType] = doc.docType;
  row[idx.itemName] = doc.itemName;
  row[idx.fileId] = doc.fileId;
  row[idx.fileName] = doc.fileName;
  row[idx.fileUrl] = doc.fileUrl;
  row[idx.uploadedAt] = now;
  row[idx.updatedAt] = now;
  row[idx.note] = safeString_(doc.note);
  row[idx.sortOrder] = Number(doc.sortOrder || 0);
  row[idx.status] = safeString_(doc.status || 'active');
  sh.appendRow(row);
}

function markRowsDeleted_(youthId, category, docTypes) {
  const sh = ensureYouthDocumentsSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return;
  const idx = headerMap_(values[0]);
  const now = formatDateTime_(new Date());
  const typeSet = nullToSet_(docTypes);

  for (let i = 1; i < values.length; i++) {
    const isTarget = safeString_(values[i][idx.youthId]) === youthId && safeString_(values[i][idx.docCategory]) === category;
    const typeOk = !typeSet || typeSet[safeString_(values[i][idx.docType])];
    const active = normalizeText_(values[i][idx.status] || 'active') === 'active';
    if (isTarget && typeOk && active) {
      sh.getRange(i + 1, idx.status + 1).setValue('deleted');
      sh.getRange(i + 1, idx.updatedAt + 1).setValue(now);
    }
  }
}

function validateDocumentPayload_(payload, category) {
  if (!payload || !payload.youthId) throw new Error('ไม่พบ youthId');
  if (category && payload.docCategory && payload.docCategory !== category) throw new Error('ประเภทเอกสารไม่ถูกต้อง');
  const youthIds = getYouthBasicMap_().map(function(y) { return y.youthId; });
  if (youthIds.indexOf(payload.youthId) === -1) throw new Error('ไม่พบข้อมูลเด็กและเยาวชน');
}

function sanitizePt1Items_(items, allowExisting) {
  const requiredTypes = ['PT1_DOCUMENT', 'FRONT_PHOTO', 'BACK_PHOTO'];
  const byType = {};
  const out = [];

  items.forEach(function(it, i) {
    const docType = safeString_(it.docType).trim() || 'CUSTOM_ATTACHMENT';
    const itemName = safeString_(it.itemName).trim();
    if (!itemName) throw new Error('กรุณาระบุชื่อรายการเอกสารทุกแถว');
    if (requiredTypes.indexOf(docType) > -1) byType[docType] = true;

    const hasNewFile = it.file && it.file.base64Data;
    const hasExisting = allowExisting && safeString_(it.existingFileId);
    if (!hasNewFile && !hasExisting) throw new Error('รายการ "' + itemName + '" ยังไม่ได้แนบไฟล์');

    out.push({
      docType: docType,
      itemName: itemName,
      file: hasNewFile ? it.file : null,
      existingFileId: safeString_(it.existingFileId),
      existingFileName: safeString_(it.existingFileName),
      existingFileUrl: safeString_(it.existingFileUrl),
      order: i + 1
    });
  });

  requiredTypes.forEach(function(t) {
    if (!byType[t]) throw new Error('ต้องมีเอกสารหลักครบ: PT1_DOCUMENT, FRONT_PHOTO, BACK_PHOTO');
  });

  return out;
}
function uploadToDriveV3_(blob, fileName, mimeType) {
  const token = ScriptApp.getOAuthToken();
  const boundary = 'gas-file-' + new Date().getTime();
  const meta = { name: fileName, parents: [CONFIG.DRIVE_FOLDER_ID] };
  const delimiter = '\r\n--' + boundary + '\r\n';
  const close = '\r\n--' + boundary + '--';

  const payload = delimiter + 'Content-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) + delimiter +
    'Content-Type: ' + (mimeType || 'application/octet-stream') + '\r\n\r\n';

  const body = Utilities.newBlob(payload).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob(close).getBytes());

  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name', {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    headers: { Authorization: 'Bearer ' + token },
    payload: body,
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error('อัปโหลดไฟล์ไม่สำเร็จ: ' + res.getContentText());
  return JSON.parse(res.getContentText());
}

function uploadFilePlaceholder(payload) {
  if (!payload || !payload.fileName) throw new Error('ข้อมูลไฟล์ไม่ครบถ้วน');
  return { ok: false, message: 'กำลังพัฒนา: uploadFilePlaceholder', expectedFields: ['fileName', 'mimeType', 'base64Data'] };
}

function upsertSheetWithHeader_(ss, sheetName, header) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  sh.setFrozenRows(1);
  return sh;
}
function headerMap_(header) { const map = {}; header.forEach(function(name, i) { map[safeString_(name)] = i; }); return map; }
function safeString_(v) { return v === null || v === undefined ? '' : String(v); }
function normalizeText_(v) { return safeString_(v).trim().toLowerCase(); }
function normalizeGradeLevel_(v){
  if (v instanceof Date) return v.getDate() + '/' + (v.getMonth()+1);
  var t = safeString_(v).trim();
  if(!t) return '';
  if (/GMT|\d{4}/.test(t)) {
    var d = new Date(t);
    if (!isNaN(d.getTime())) return d.getDate() + '/' + (d.getMonth()+1);
  }
  return t;
}
function extractBaseGrade_(v){
  var t=normalizeGradeLevel_(v);
  if(!t) return '';
  return t.split('/')[0].trim();
}

function formatDateTime_(d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'); }
function dateNum_(v) { const d = v instanceof Date ? v : new Date(v || 0); return isNaN(d.getTime()) ? 0 : d.getTime(); }
function maxDateText_() {
  const values = Array.prototype.slice.call(arguments).filter(function(v) { return !!v; });
  if (!values.length) return '';
  return values.sort(function(a, b) { return dateNum_(b) - dateNum_(a); })[0];
}
function nullToSet_(arr) {
  if (!arr || !arr.length) return null;
  const out = {}; arr.forEach(function(v) { out[safeString_(v)] = true; }); return out;
}

function generateId(prefix){
  const now=new Date();
  const p=safeString_(prefix||'ID');
  return p+'-'+Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyyMMdd-HHmmss')+'-'+Math.floor(1000+Math.random()*9000);
}

function validateYouthPayload_(data) {
  if (!data) throw new Error('ไม่พบข้อมูล');
  if (!safeString_(data.fullName).trim()) throw new Error('กรุณากรอกชื่อนามสกุลเด็กและเยาวชน');
  if (!safeString_(data.gradeLevel).trim()) throw new Error('กรุณากรอกระดับชั้น');
}

function generateYouthId_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = ensureYouthSheet_();
    const datePart = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd');
    const prefix = 'YTH-' + datePart + '-';
    let maxSeq = 0;
    if (sh.getLastRow() > 1) {
      const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = headerMap_(header);
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(function(r) {
        const id = safeString_(r[idx.youthId]);
        if (id.indexOf(prefix) === 0) maxSeq = Math.max(maxSeq, Number(id.replace(prefix, '')) || 0);
      });
    }
    return prefix + ('0000' + (maxSeq + 1)).slice(-4);
  } finally {
    lock.releaseLock();
  }
}

function generateDocId_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = ensureYouthDocumentsSheet_();
    const datePart = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd');
    const prefix = 'DOC-' + datePart + '-';
    let maxSeq = 0;
    if (sh.getLastRow() > 1) {
      const idx = headerMap_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(function(r) {
        const id = safeString_(r[idx.docId]);
        if (id.indexOf(prefix) === 0) maxSeq = Math.max(maxSeq, Number(id.replace(prefix, '')) || 0);
      });
    }
    return prefix + ('0000' + (maxSeq + 1)).slice(-4);
  } finally {
    lock.releaseLock();
  }
}

function getActiveSettingsList_(sheetName) {
  const sh = getSpreadsheet().getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const idx = headerMap_(data[0]);
  return data.slice(1).map(function(r) {
    return { id: safeString_(r[idx.id]), name: safeString_(r[idx.name]), status: safeString_(r[idx.status] || 'active') };
  }).filter(function(it) { return normalizeText_(it.name) && normalizeText_(it.status) !== 'inactive'; });
}

function mapImportHeader_(headerRow) {
  const idxMap = {};
  (headerRow || []).forEach(function(h, i) { idxMap[normalizeText_(h)] = i; });
  const required = { fullName: 'ชื่อนามสกุลเด็กและเยาวชน', gradeLevel: 'ระดับชั้น', vocationUnit: 'หน่วยเรียนวิชาชีพ', advisorTeacher: 'ครูที่ปรึกษา' };
  const map = {
    fullName: idxMap[normalizeText_(required.fullName)],
    gradeLevel: idxMap[normalizeText_(required.gradeLevel)],
    vocationUnit: idxMap[normalizeText_(required.vocationUnit)],
    advisorTeacher: idxMap[normalizeText_(required.advisorTeacher)]
  };
  const missing = Object.keys(map).filter(function(k) { return map[k] === undefined; }).map(function(k) { return required[k]; });
  if (missing.length) throw new Error('หัวตารางไม่ครบ: ' + missing.join(', '));
  return map;
}

function convertExcelToSheetRows_(blob, fileName) {
  const token = ScriptApp.getOAuthToken();
  const boundary = 'gas-boundary-' + new Date().getTime();
  const metadata = { name: 'tmp-import-' + fileName, mimeType: 'application/vnd.google-apps.spreadsheet', parents: [CONFIG.DRIVE_FOLDER_ID] };
  const delimiter = '\r\n--' + boundary + '\r\n';
  const closeDelim = '\r\n--' + boundary + '--';
  const payload = delimiter + 'Content-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(metadata) + delimiter + 'Content-Type: ' + (blob.getContentType() || 'application/octet-stream') + '\r\n\r\n';
  const body = Utilities.newBlob(payload).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob(closeDelim).getBytes());

  const response = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true', {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    headers: { Authorization: 'Bearer ' + token },
    payload: body,
    muteHttpExceptions: true
  });
  if (response.getResponseCode() >= 300) throw new Error('ไม่สามารถแปลงไฟล์ Excel ได้: ' + response.getContentText());
  const created = JSON.parse(response.getContentText());
  const tempId = created.id;
  try {
    return { rows: SpreadsheetApp.openById(tempId).getSheets()[0].getDataRange().getValues() };
  } finally {
    if (tempId) {
      UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(tempId), {
        method: 'delete', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
      });
    }
  }
}

/** -------- Activity calendar module -------- */
function ensureActivityScheduleSheet_() {
  return upsertSheetWithHeader_(getSpreadsheet(), 'activity_schedule', CONFIG.SHEETS.activity_schedule);
}

function ensureHolidayCacheSheet_() {
  return upsertSheetWithHeader_(getSpreadsheet(), 'holiday_cache', CONFIG.SHEETS.holiday_cache);
}

function getYouthOptionsForActivity() {
  return getYouthBasicMap_().map(function(y) {
    return { youthId: y.youthId, fullName: y.fullName, advisorTeacher: y.advisorTeacher };
  });
}
function uploadToDriveV3_(blob, fileName, mimeType) {
  const token = ScriptApp.getOAuthToken();
  const boundary = 'gas-file-' + new Date().getTime();
  const meta = { name: fileName, parents: [CONFIG.DRIVE_FOLDER_ID] };
  const delimiter = '\r\n--' + boundary + '\r\n';
  const close = '\r\n--' + boundary + '--';

  const payload = delimiter + 'Content-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) + delimiter +
    'Content-Type: ' + (mimeType || 'application/octet-stream') + '\r\n\r\n';

  const body = Utilities.newBlob(payload).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob(close).getBytes());

  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name', {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    headers: { Authorization: 'Bearer ' + token },
    payload: body,
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error('อัปโหลดไฟล์ไม่สำเร็จ: ' + res.getContentText());
  return JSON.parse(res.getContentText());
}

function uploadFilePlaceholder(payload) {
  if (!payload || !payload.fileName) throw new Error('ข้อมูลไฟล์ไม่ครบถ้วน');
  return { ok: false, message: 'กำลังพัฒนา: uploadFilePlaceholder', expectedFields: ['fileName', 'mimeType', 'base64Data'] };
}

function upsertSheetWithHeader_(ss, sheetName, header) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  sh.setFrozenRows(1);
  return sh;
}
function headerMap_(header) { const map = {}; header.forEach(function(name, i) { map[safeString_(name)] = i; }); return map; }
function safeString_(v) { return v === null || v === undefined ? '' : String(v); }
function normalizeText_(v) { return safeString_(v).trim().toLowerCase(); }
function normalizeGradeLevel_(v){
  if (v instanceof Date) return v.getDate() + '/' + (v.getMonth()+1);
  var t = safeString_(v).trim();
  if(!t) return '';
  if (/GMT|\d{4}/.test(t)) {
    var d = new Date(t);
    if (!isNaN(d.getTime())) return d.getDate() + '/' + (d.getMonth()+1);
  }
  return t;
}
function extractBaseGrade_(v){
  var t=normalizeGradeLevel_(v);
  if(!t) return '';
  return t.split('/')[0].trim();
}

function formatDateTime_(d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'); }
function dateNum_(v) { const d = v instanceof Date ? v : new Date(v || 0); return isNaN(d.getTime()) ? 0 : d.getTime(); }
function maxDateText_() {
  const values = Array.prototype.slice.call(arguments).filter(function(v) { return !!v; });
  if (!values.length) return '';
  return values.sort(function(a, b) { return dateNum_(b) - dateNum_(a); })[0];
}
function nullToSet_(arr) {
  if (!arr || !arr.length) return null;
  const out = {}; arr.forEach(function(v) { out[safeString_(v)] = true; }); return out;
}

function generateId(prefix){
  const now=new Date();
  const p=safeString_(prefix||'ID');
  return p+'-'+Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyyMMdd-HHmmss')+'-'+Math.floor(1000+Math.random()*9000);
}

function validateYouthPayload_(data) {
  if (!data) throw new Error('ไม่พบข้อมูล');
  if (!safeString_(data.fullName).trim()) throw new Error('กรุณากรอกชื่อนามสกุลเด็กและเยาวชน');
  if (!safeString_(data.gradeLevel).trim()) throw new Error('กรุณากรอกระดับชั้น');
}

function generateYouthId_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = ensureYouthSheet_();
    const datePart = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd');
    const prefix = 'YTH-' + datePart + '-';
    let maxSeq = 0;
    if (sh.getLastRow() > 1) {
      const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = headerMap_(header);
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(function(r) {
        const id = safeString_(r[idx.youthId]);
        if (id.indexOf(prefix) === 0) maxSeq = Math.max(maxSeq, Number(id.replace(prefix, '')) || 0);
      });
    }
    return prefix + ('0000' + (maxSeq + 1)).slice(-4);
  } finally {
    lock.releaseLock();
  }
}

function generateDocId_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = ensureYouthDocumentsSheet_();
    const datePart = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd');
    const prefix = 'DOC-' + datePart + '-';
    let maxSeq = 0;
    if (sh.getLastRow() > 1) {
      const idx = headerMap_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(function(r) {
        const id = safeString_(r[idx.docId]);
        if (id.indexOf(prefix) === 0) maxSeq = Math.max(maxSeq, Number(id.replace(prefix, '')) || 0);
      });
    }
    return prefix + ('0000' + (maxSeq + 1)).slice(-4);
  } finally {
    lock.releaseLock();
  }
}

function getActiveSettingsList_(sheetName) {
  const sh = getSpreadsheet().getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const idx = headerMap_(data[0]);
  return data.slice(1).map(function(r) {
    return { id: safeString_(r[idx.id]), name: safeString_(r[idx.name]), status: safeString_(r[idx.status] || 'active') };
  }).filter(function(it) { return normalizeText_(it.name) && normalizeText_(it.status) !== 'inactive'; });
}

function mapImportHeader_(headerRow) {
  const idxMap = {};
  (headerRow || []).forEach(function(h, i) { idxMap[normalizeText_(h)] = i; });
  const required = { fullName: 'ชื่อนามสกุลเด็กและเยาวชน', gradeLevel: 'ระดับชั้น', vocationUnit: 'หน่วยเรียนวิชาชีพ', advisorTeacher: 'ครูที่ปรึกษา' };
  const map = {
    fullName: idxMap[normalizeText_(required.fullName)],
    gradeLevel: idxMap[normalizeText_(required.gradeLevel)],
    vocationUnit: idxMap[normalizeText_(required.vocationUnit)],
    advisorTeacher: idxMap[normalizeText_(required.advisorTeacher)]
  };
  const missing = Object.keys(map).filter(function(k) { return map[k] === undefined; }).map(function(k) { return required[k]; });
  if (missing.length) throw new Error('หัวตารางไม่ครบ: ' + missing.join(', '));
  return map;
}

function convertExcelToSheetRows_(blob, fileName) {
  const token = ScriptApp.getOAuthToken();
  const boundary = 'gas-boundary-' + new Date().getTime();
  const metadata = { name: 'tmp-import-' + fileName, mimeType: 'application/vnd.google-apps.spreadsheet', parents: [CONFIG.DRIVE_FOLDER_ID] };
  const delimiter = '\r\n--' + boundary + '\r\n';
  const closeDelim = '\r\n--' + boundary + '--';
  const payload = delimiter + 'Content-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(metadata) + delimiter + 'Content-Type: ' + (blob.getContentType() || 'application/octet-stream') + '\r\n\r\n';
  const body = Utilities.newBlob(payload).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob(closeDelim).getBytes());

  const response = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true', {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    headers: { Authorization: 'Bearer ' + token },
    payload: body,
    muteHttpExceptions: true
  });
  if (response.getResponseCode() >= 300) throw new Error('ไม่สามารถแปลงไฟล์ Excel ได้: ' + response.getContentText());
  const created = JSON.parse(response.getContentText());
  const tempId = created.id;
  try {
    return { rows: SpreadsheetApp.openById(tempId).getSheets()[0].getDataRange().getValues() };
  } finally {
    if (tempId) {
      UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(tempId), {
        method: 'delete', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
      });
    }
  }
}

/** -------- Activity calendar module -------- */
function ensureActivityScheduleSheet_() {
  return upsertSheetWithHeader_(getSpreadsheet(), 'activity_schedule', CONFIG.SHEETS.activity_schedule);
}

function ensureHolidayCacheSheet_() {
  return upsertSheetWithHeader_(getSpreadsheet(), 'holiday_cache', CONFIG.SHEETS.holiday_cache);
}

function getYouthOptionsForActivity() {
  return getYouthBasicMap_().map(function(y) {
    return { youthId: y.youthId, fullName: y.fullName, advisorTeacher: y.advisorTeacher };
  });
}
function syncHolidays() {
  const y = new Date().getFullYear();
  const a = syncThaiHolidaysFromGoogleCalendar(y);
  const b = syncThaiHolidaysFromGoogleCalendar(y + 1);
  return { ok: !!(a.ok || b.ok), message: (a.message || '') + ' ' + (b.message || ''), count: Number(a.count || 0) + Number(b.count || 0) };
}
function refreshHolidayCacheForYear(year) {
  const y = Number(year || new Date().getFullYear());
  if (!y) throw new Error('กรุณาระบุปี ค.ศ.');
  return syncThaiHolidaysFromGoogleCalendar(y);
}
function syncThaiHolidaysFromGoogleCalendar(year) {
  ensureHolidayCacheSheet_();
  const y = Number(year);
  if (!y) throw new Error('กรุณาระบุปี ค.ศ.');
  const start = new Date(y, 0, 1);
  const end = new Date(y, 11, 31, 23, 59, 59);
  const candidates = ['th.th#holiday@group.v.calendar.google.com', 'en.th#holiday@group.v.calendar.google.com'];
  let events = [];
  let calendarId = '';
  for (let i = 0; i < candidates.length; i++) {
    try {
      const cal = CalendarApp.getCalendarById(candidates[i]);
      if (!cal) continue;
      const items = cal.getEvents(start, end) || [];
      if (items.length) {
        events = items.map(function(ev) { return { date: formatDateOnly_(toDate_(ev.getAllDayStartDate())), name: ev.getTitle(), type: 'PUBLIC_HOLIDAY' }; });
        calendarId = candidates[i];
        break;
      }
    } catch (e) {}
  }
  if (!events.length) return { ok: false, message: 'ไม่สามารถดึงวันหยุดจาก Google Calendar ได้ ใช้ cache เดิม', count: 0, year: y };
  const sh = ensureHolidayCacheSheet_();
  const vals = sh.getDataRange().getValues();
  const idx = headerMap_(vals[0]);
  const now = formatDateTime_(new Date());
  for (let i = vals.length - 1; i >= 1; i--) {
    const d = toDate_(vals[i][idx.holidayDate]);
    if (d.getFullYear() === y) {
      if (idx.status !== undefined) sh.getRange(i + 1, idx.status + 1).setValue('deleted');
    }
  }
  const rows = events.map(function(it, i) {
    const row = new Array(vals[0].length).fill('');
    row[idx.holidayId] = 'HOL-' + y + '-' + ('0000' + (i + 1)).slice(-4);
    row[idx.holidayDate] = it.date;
    row[idx.holidayName] = it.name;
    row[idx.holidayType] = it.type;
    if (idx.sourceCalendarId !== undefined) row[idx.sourceCalendarId] = calendarId;
    row[idx.updatedAt] = now;
    if (idx.status !== undefined) row[idx.status] = 'active';
    return row;
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, vals[0].length).setValues(rows);
  return { ok: true, message: 'sync วันหยุดจาก Google Calendar สำเร็จ', count: rows.length, year: y };
}
function getHolidayCache(startDate, endDate) {
  try {
    ensureHolidayCacheSheet_();
    const sh = getSpreadsheet().getSheetByName('holiday_cache');
    if (!sh) return [];
    const vals = sh.getDataRange().getValues();
    if (vals.length <= 1) return [];
    const idx = headerMap_(vals[0]);
    const start = toDate_(startDate);
    const end = toDate_(endDate);
    return vals.slice(1).map(function(r) {
      return {
        holidayId: safeString_(r[idx.holidayId]),
        holidayDate: formatDateOnly_(toDate_(r[idx.holidayDate])),
        holidayName: safeString_(r[idx.holidayName]),
        holidayType: safeString_(r[idx.holidayType]),
        sourceCalendarId: idx.sourceCalendarId !== undefined ? safeString_(r[idx.sourceCalendarId]) : '',
        updatedAt: safeString_(r[idx.updatedAt]),
        status: idx.status !== undefined ? safeString_(r[idx.status] || 'active') : 'active'
      };
    }).filter(function(it) {
      if (normalizeText_(it.status) === 'deleted') return false;
      const d = toDate_(it.holidayDate);
      return d >= start && d <= end;
    });
  } catch (e) {
    Logger.log('[getHolidayCache] error: %s', e && e.message ? e.message : e);
    return [];
  }
}

function getHolidayMap_(startDate, endDate) {
  const list = getHolidayCache(startDate, endDate);
  const map = {};
  list.forEach(function(h) {
    if (!map[h.holidayDate]) map[h.holidayDate] = [];
    map[h.holidayDate].push(h);
  });
  return map;
}

function cleanupOldActivities() {
  ensureActivityScheduleSheet_();
  const sh = getSpreadsheet().getSheetByName('activity_schedule');
  const vals = sh.getDataRange().getValues();
  if (vals.length<=1) return { ok:true, affected:0 };
  const idx = headerMap_(vals[0]);
  const today = new Date();
  const minFiscal = getThaiFiscalYear_(today)-1;
  let affected = 0;
  for (let i=1;i<vals.length;i++) {
    const d = safeString_(vals[i][idx.eventDate]);
    if (!d) continue;
    const fy = getThaiFiscalYear_(toDate_(d));
    if (fy < minFiscal && normalizeText_(vals[i][idx.status]||'active') !== 'deleted') {
      sh.getRange(i+1, idx.status+1).setValue('deleted');
      sh.getRange(i+1, idx.updatedAt+1).setValue(formatDateTime_(new Date()));
      affected++;
    }
  }
  return { ok:true, affected:affected, minFiscal:minFiscal };
}

function getThaiFiscalYear_(dateObj) {
  const d = dateObj instanceof Date ? dateObj : toDate_(dateObj);
  const y = d.getFullYear();
  const m = d.getMonth()+1;
  const fyAD = m >= 10 ? y + 1 : y;
  return fyAD + 543;
}

function getUpcomingReminders_(rangeType) {
  const today = new Date();
  const start = formatDateOnly_(today);
  let end = new Date(today);
  if (rangeType==='DAY') end = today;
  else if (rangeType==='WEEK') end.setDate(end.getDate()+7);
  else if (rangeType==='MONTH') end.setMonth(end.getMonth()+1);
  else end.setMonth(end.getMonth()+3);
  const events = fetchActivityRows_(start, formatDateOnly_(end), {});
  return events.filter(function(e){
    if (rangeType==='DAY') return normalizeText_(e.reminderDay)==='true';
    if (rangeType==='WEEK') return normalizeText_(e.reminderWeek)==='true';
    if (rangeType==='MONTH') return normalizeText_(e.reminderMonth)==='true';
    return normalizeText_(e.reminderQuarter)==='true';
  });
}

function fetchActivityRows_(startDate, endDate, params) {
  ensureActivityScheduleSheet_();
  params = params || {};
  const sh = getSpreadsheet().getSheetByName('activity_schedule');
  const vals = sh.getDataRange().getValues();
  if (vals.length<=1) return [];
  const idx = headerMap_(vals[0]);
  const start = startDate ? toDate_(startDate) : new Date(1900,0,1);
  const end = endDate ? toDate_(endDate) : new Date(2500,0,1);
  const query = normalizeText_(params.query||'');
  const eventTypeFilter = safeString_(params.eventType||'');

  return vals.slice(1).map(function(r){ return activityRowToObj_(r, idx); }).filter(function(it){
    if (normalizeText_(it.status||'active')==='deleted') return false;
    const s = toDate_(it.startDate || it.eventDate);
    const e = toDate_(it.endDate || it.startDate || it.eventDate);
    if (e < start || s > end) return false;
    if (eventTypeFilter && it.eventType !== eventTypeFilter) return false;
    if (query) {
      const hay = [it.personName, it.roleType, it.youthNameSnapshot, it.courtName, it.hospitalName, it.title, it.taskOwner, it.description, it.staffNames, it.location].join(' | ');
      if (normalizeText_(hay).indexOf(query)===-1) return false;
    }
    return true;
  });
}

function normalizeActivityPayload_(payload, isUpdate) {
  payload = payload || {};
  const now = formatDateTime_(new Date());
  const youthIds = (payload.youthIds && payload.youthIds.length ? payload.youthIds : safeString_(payload.youthId).split(',')).map(function(v){return safeString_(v).trim();}).filter(Boolean);
  const out = {
    activityId: isUpdate ? safeString_(payload.activityId) : generateActivityId_(),
    eventType: safeString_(payload.eventType),
    title: safeString_(payload.title),
    youthId: youthIds.join(','),
    youthNameSnapshot: '',
    advisorTeacherSnapshot: '',
    courtName: safeString_(payload.courtName),
    courtMode: safeString_(payload.courtMode),
    hospitalName: safeString_(payload.hospitalName),
    staffNames: safeString_(payload.staffNames),
    taskOwner: safeString_(payload.taskOwner),
    personName: safeString_(payload.personName),
    roleType: safeString_(payload.roleType || 'อื่น ๆ'),
    startDate: safeString_(payload.startDate || payload.eventDate),
    endDate: safeString_(payload.endDate || payload.startDate || payload.eventDate),
    eventDate: safeString_(payload.startDate || payload.eventDate),
    startTime: safeString_(payload.startTime),
    endTime: safeString_(payload.endTime),
    allDay: payload.allDay ? 'TRUE' : 'FALSE',
    description: safeString_(payload.description),
    location: safeString_(payload.location),
    reminderDay: payload.reminderDay ? 'TRUE' : 'FALSE',
    reminderWeek: payload.reminderWeek ? 'TRUE' : 'FALSE',
    reminderMonth: payload.reminderMonth ? 'TRUE' : 'FALSE',
    reminderQuarter: payload.reminderQuarter ? 'TRUE' : 'FALSE',
    status: safeString_(payload.status||'active'),
    createdAt: isUpdate ? safeString_(payload.createdAt||'') : now,
    updatedAt: now,
    createdBy: safeString_(payload.createdBy||'system')
  };

  const youthMap = {};
  getYouthBasicMap_().forEach(function(y){ youthMap[y.youthId]=y; });
  const selected = youthIds.map(function(id){ return youthMap[id]; }).filter(Boolean);
  if (selected.length) {
    out.youthNameSnapshot = selected.map(function(y){ return y.fullName; }).join(', ');
    out.advisorTeacherSnapshot = selected.map(function(y){ return y.advisorTeacher; }).filter(Boolean).join(', ');
  }

  if (!out.title) throw new Error('กรุณากรอกชื่อกิจกรรม');
  if (!out.startDate) throw new Error('กรุณาเลือกวันที่เริ่ม');
  if (toDate_(out.endDate) < toDate_(out.startDate)) throw new Error('วันที่สิ้นสุดต้องไม่น้อยกว่าวันที่เริ่ม');
  if (normalizeText_(out.allDay) !== 'true' && out.startTime && out.endTime && out.endTime < out.startTime) throw new Error('เวลาเริ่มต้องมาก่อนเวลาสิ้นสุด');
  if (out.eventType==='COURT') {
    if (!out.youthId || !out.startDate || (!out.allDay && !out.startTime) || !out.courtName || !out.courtMode) throw new Error('นัดขึ้นศาลต้องเลือกเด็กอย่างน้อย 1 คน และกรอกวันที่/เวลาเริ่ม/ศาล/รูปแบบ');
    out.title = 'นัดขึ้นศาล - ' + (selected[0] ? selected[0].fullName : '') + (selected.length>1 ? ' และอีก '+(selected.length-1)+' คน' : '');
  } else if (out.eventType==='HOSPITAL') {
    if (!out.youthId || !out.startDate || (!out.allDay && !out.startTime) || !out.hospitalName) throw new Error('นัดโรงพยาบาลต้องเลือกเด็กอย่างน้อย 1 คน และกรอกวันที่/เวลา/โรงพยาบาล');
    out.title = 'นัดโรงพยาบาล - ' + (selected[0] ? selected[0].fullName : '') + (selected.length>1 ? ' และอีก '+(selected.length-1)+' คน' : '');
  } else if (out.eventType==='INTERNAL_YOUTH') {
    if (!out.startDate || !out.title) throw new Error('กิจกรรมภายในต้องกรอก วันที่/ชื่อกิจกรรม');
  } else if (out.eventType==='STAFF_ACTIVITY') {
    if (!out.startDate || !out.title || (!out.staffNames && !out.taskOwner && !out.personName)) throw new Error('กิจกรรมเจ้าหน้าที่ต้องมี วันที่/ชื่อกิจกรรม และผู้เกี่ยวข้อง');
  } else if (out.eventType==='DEADLINE') {
    if (!out.startDate || !out.title) throw new Error('กำหนดส่งงานต้องมี วันที่/ชื่องาน');
  } else {
    if (!out.startDate || !out.title) throw new Error('กรุณากรอกข้อมูลกิจกรรมให้ครบ');
  }
  if (!out.personName) out.personName = out.taskOwner || out.youthNameSnapshot || out.staffNames || '';
  out.eventDate = out.startDate;
  return out;
}
function appendActivityRow_(activity) {
  const sh = ensureActivityScheduleSheet_();
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row = buildActivityRowFromPayload_(activity, header);
  sh.appendRow(row);
}

function buildActivityRowFromPayload_(activity, header) {
  const idx = headerMap_(header);
  const row = new Array(header.length).fill('');
  Object.keys(idx).forEach(function(k){ if (activity[k] !== undefined) row[idx[k]] = activity[k]; });
  return row;
}

function activityRowToObj_(r, idx) {
  const obj = {};
  Object.keys(idx).forEach(function(k){ obj[k] = safeString_(r[idx[k]]); });
  obj.startDate = obj.startDate ? formatDateOnly_(toDate_(obj.startDate)) : formatDateOnly_(toDate_(obj.eventDate));
  obj.endDate = obj.endDate ? formatDateOnly_(toDate_(obj.endDate)) : obj.startDate;
  obj.eventDate = obj.startDate;
  return obj;
}

function generateActivityId_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = ensureActivityScheduleSheet_();
    const prefix = 'ACT-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd') + '-';
    let max = 0;
    if (sh.getLastRow()>1) {
      const vals = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
      const idx = headerMap_(sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]);
      vals.forEach(function(r){
        const id = safeString_(r[idx.activityId]);
        if (id.indexOf(prefix)===0) max = Math.max(max, Number(id.replace(prefix,''))||0);
      });
    }
    return prefix + ('0000'+(max+1)).slice(-4);
  } finally { lock.releaseLock(); }
}

function toDate_(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const d = new Date(v);
  if (isNaN(d.getTime())) return new Date(1900,0,1);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function parseDate_(v) { return toDate_(v); }

function formatDateOnly_(d) {
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}
function formatThaiDate_(v) {
  const d = toDate_(v);
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'd MMMM ') + (d.getFullYear() + 543);
}
function getThaiMonthName_(month) {
  const names = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  return names[Math.max(1, Math.min(12, Number(month || 1))) - 1];
}
function getDateRange_(startDate, endDate) {
  const start = toDate_(startDate);
  const end = toDate_(endDate);
  const out = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) out.push(formatDateOnly_(new Date(d)));
  return out;
}

/** -------- Visit module -------- */
function ensureVisitQuotaRuleSheet_(){ return upsertSheetWithHeader_(getSpreadsheet(),'visit_normal_quota_rule',CONFIG.SHEETS.visit_normal_quota_rule || CONFIG.SHEETS.visit_quota_rule); }
function ensureVisitSpecialRuleSheet_(){ return upsertSheetWithHeader_(getSpreadsheet(),'visit_special_quota_rule',CONFIG.SHEETS.visit_special_quota_rule || CONFIG.SHEETS.visit_special_rule); }
function ensureVisitBookingSheet_(){ return upsertSheetWithHeader_(getSpreadsheet(),'visit_booking',CONFIG.SHEETS.visit_booking); }

function getYouthOptionsForVisit(){
  return getYouthRowsForVisit_().map(function(y){
    return {youthId:y.youthId,fullName:y.fullName,gradeLevel:y.gradeLevel||'',advisorTeacher:y.advisorTeacher||''};
  });
}

function getVisitDashboardData(params){
  params=params||{};const month=Number(params.month),year=Number(params.year);
  const table=getVisitTableData({month:month,year:year,query:params.query||'',gradeLevel:params.gradeLevel||'',statusFilter:params.statusFilter||'ALL'});
  return { cards:getVisitSummaryCards({month:month,year:year,rows:table.items}), table:table };
}

function getVisitSummaryCards(params){
  const rows=params.rows||[];
  return {
    totalYouth: rows.length,
    canBook: rows.filter(function(r){return r.statusKey==='CAN_BOOK' || r.statusKey==='SPECIAL_AVAILABLE';}).length,
    fullUsed: rows.filter(function(r){return r.statusKey==='FULL';}).length,
    monthBookings: rows.reduce(function(a,b){return a+(b.usedTotal||0);},0),
    activeSpecialRules: getVisitSpecialRules().filter(function(r){return normalizeText_(r.status)==='active';}).length
  };
}

function getVisitTableData(params){
  ensureVisitQuotaRuleSheet_();ensureVisitSpecialRuleSheet_();ensureVisitBookingSheet_();ensureYouthSheet_();
  params=params||{};const month=Number(params.month),year=Number(params.year);const query=normalizeText_(params.query||'');
  const gradeFilter=safeString_(params.gradeLevel||'');const statusFilter=safeString_(params.statusFilter||'ALL');
  const youth=getYouthRowsForVisit_();
  let items=youth.map(function(y){
    const usage=getVisitUsageByMonth_(y.youthId,month,year);
    const base=getBaseQuotaForGrade_(y.gradeLevel);
    const extra=usage.extraAvailable;
    const total=base+extra;
    const remaining=total-usage.usedTotal;
    let statusKey='CAN_BOOK';let statusLabel='ยังจองได้';
    if (remaining<=0){statusKey='FULL';statusLabel='ใช้สิทธิครบแล้ว';}
    else if (extra>0){statusKey='SPECIAL_AVAILABLE';statusLabel='มีสิทธิพิเศษ';}
    if (usage.usedTotal===0 && remaining>0){statusKey=statusKey==='SPECIAL_AVAILABLE'?'SPECIAL_AVAILABLE':'NO_BOOKING';statusLabel=statusKey==='SPECIAL_AVAILABLE'?'มีสิทธิพิเศษ':'ยังไม่มีการจอง';}
    return {
      youthId:y.youthId,fullName:y.fullName,gradeLevel:y.gradeLevel,advisorTeacher:y.advisorTeacher,
      baseQuota:base,specialQuota:extra,normalUsed:usage.normalUsed,specialUsed:usage.specialUsed,normalRemain:Math.max(0,base-usage.normalUsed),specialRemain:Math.max(0,extra-usage.specialUsed),usedTotal:usage.usedTotal,remaining:remaining,
      bookingDates:usage.bookingDates,quotaTypes:usage.bookings.map(function(b){return {date:b.bookingDate,type:b.quotaTypeUsed};}),
      statusKey:statusKey,statusLabel:statusLabel,
      note:usage.note
    };
  });
  if (query) items=items.filter(function(r){return normalizeText_([r.fullName,r.gradeLevel,r.advisorTeacher].join(' ')).indexOf(query)>-1;});
  if (gradeFilter) items=items.filter(function(r){return safeString_(r.gradeLevel)===gradeFilter;});
  if (statusFilter!=='ALL') items=items.filter(function(r){return r.statusKey===statusFilter;});
  return { items:items, month:month, year:year };
}

function getVisitDetailByYouthId(youthId, month, year){
  const youth=getYouthRowsForVisit_().find(function(y){return y.youthId===youthId;});
  if(!youth) throw new Error('ไม่พบข้อมูลเด็ก');
  const usage=getVisitUsageByMonth_(youthId,Number(month),Number(year));
  const base=getBaseQuotaForGrade_(youth.gradeLevel);const extra=usage.extraAvailable;
  return { youth:youth, baseQuota:base, specialQuota:extra, specialUsed:usage.specialUsed, normalUsed:usage.normalUsed, usedTotal:usage.usedTotal, remaining:(base+extra-usage.usedTotal), bookings:usage.bookings, specialRules:usage.applicableRules };
}

function getVisitQuotaRules(){
  ensureVisitQuotaRuleSheet_();
  const sh=getSpreadsheet().getSheetByName('visit_normal_quota_rule');const vals=sh.getDataRange().getValues();if(vals.length<=1)return[];
  const idx=headerMap_(vals[0]);
  return vals.slice(1).map(function(r){return {ruleId:safeString_(r[idx.ruleId]),gradeLevel:normalizeGradeLevel_(r[idx.gradeLevel]),allowedPerMonth:Number(r[idx.allowedPerMonth]||0),status:safeString_(r[idx.status]||'active'),createdAt:safeString_(r[idx.createdAt]),updatedAt:safeString_(r[idx.updatedAt])};}).filter(function(r){return r.ruleId;});
}
function saveVisitQuotaRule(data){
  if(!safeString_(data.gradeLevel)) throw new Error('กรุณาระบุระดับชั้น');
  if(safeString_(data.allowedPerMonth)==='') throw new Error('กรุณาระบุสิทธิต่อเดือน');
  const sh=ensureVisitQuotaRuleSheet_();const existing=getVisitQuotaRules().filter(function(r){return normalizeText_(r.status)==='active';});if(existing.some(function(r){return safeString_(r.gradeLevel)===safeString_(data.gradeLevel);})){throw new Error('ระดับชั้นนี้ถูกตั้งค่าสิทธิไว้แล้ว');}const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];const idx=headerMap_(h);const now=formatDateTime_(new Date());const row=new Array(h.length).fill('');
  var grade=validateGradeLevel_(data.gradeLevel);row[idx.ruleId]=generateId('VQR');row[idx.gradeLevel]=grade;row[idx.allowedPerMonth]=Number(data.allowedPerMonth);row[idx.status]=safeString_(data.status||'active');row[idx.createdAt]=now;row[idx.updatedAt]=now;sh.appendRow(row);return {ok:true,message:'บันทึกกติกาสิทธิพื้นฐานสำเร็จ'};
}
function updateVisitQuotaRule(data){
  if(!data||!data.ruleId) throw new Error('ไม่พบ ruleId');
  const sh=ensureVisitQuotaRuleSheet_();const vals=sh.getDataRange().getValues();const idx=headerMap_(vals[0]);
  for(let i=1;i<vals.length;i++){if(safeString_(vals[i][idx.ruleId])===safeString_(data.ruleId)){var grade=validateGradeLevel_(data.gradeLevel);sh.getRange(i+1,idx.gradeLevel+1).setValue(grade);sh.getRange(i+1,idx.allowedPerMonth+1).setValue(Number(data.allowedPerMonth));sh.getRange(i+1,idx.status+1).setValue(data.status||'active');sh.getRange(i+1,idx.updatedAt+1).setValue(formatDateTime_(new Date()));return {ok:true,message:'แก้ไขกติกาสำเร็จ'};}}
  throw new Error('ไม่พบกติกา');
}
function deleteVisitQuotaRule(ruleId){ return updateVisitQuotaRule({ruleId:ruleId,gradeLevel:'',allowedPerMonth:0,status:'inactive'}); }

function getVisitSpecialRules(){
  ensureVisitSpecialRuleSheet_();
  const sh=getSpreadsheet().getSheetByName('visit_special_quota_rule');const vals=sh.getDataRange().getValues();if(vals.length<=1)return[];
  const idx=headerMap_(vals[0]);
  return vals.slice(1).map(function(r){return {specialRuleId:safeString_(r[idx.specialRuleId]),ruleName:safeString_(r[idx.ruleName]),gradeLevelScope:safeString_(r[idx.gradeLevelScope]),extraQuota:Number(r[idx.extraQuota]||0),startDate:formatDateOnly_(toDate_(r[idx.startDate])),endDate:formatDateOnly_(toDate_(r[idx.endDate])),note:safeString_(r[idx.note]),status:safeString_(r[idx.status]||'active'),createdAt:safeString_(r[idx.createdAt]),updatedAt:safeString_(r[idx.updatedAt])};}).filter(function(r){return r.specialRuleId;});
}
function saveVisitSpecialRule(data){
  if(!safeString_(data.ruleName)||safeString_(data.extraQuota)===''||!safeString_(data.startDate)||!safeString_(data.endDate)) throw new Error('กรุณากรอกข้อมูลสิทธิพิเศษให้ครบ');
  var allowed = getAvailableGradeLevels_();
  if(!allowed.length) throw new Error('ยังไม่มีระดับชั้นในระบบ จึงยังตั้งค่าสิทธิพิเศษไม่ได้');
  var scopeRaw=normalizeGradeLevelScope_(data.gradeLevelScope||'ALL');
  if(normalizeText_(scopeRaw)!=='all'){
    var scopes=scopeRaw.split(',').map(function(v){return safeString_(v).trim();}).filter(Boolean);
    var invalid=scopes.filter(function(g){return allowed.indexOf(g)===-1;});
    if(invalid.length) throw new Error('พบระดับชั้นที่ไม่มีในระบบ: '+invalid.join(','));
  }
  if(toDate_(data.endDate)<toDate_(data.startDate)) throw new Error('วันสิ้นสุดต้องไม่น้อยกว่าวันเริ่มต้น');
  const sh=ensureVisitSpecialRuleSheet_();const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];const idx=headerMap_(h);const now=formatDateTime_(new Date());const row=new Array(h.length).fill('');
  row[idx.specialRuleId]=generateId('VSP');row[idx.ruleName]=data.ruleName;row[idx.gradeLevelScope]=normalizeGradeLevelScope_(data.gradeLevelScope||'ALL');row[idx.extraQuota]=Number(data.extraQuota);row[idx.startDate]=formatDateOnly_(toDate_(data.startDate));row[idx.endDate]=formatDateOnly_(toDate_(data.endDate));row[idx.note]=data.note||'';row[idx.status]=data.status||'active';row[idx.createdAt]=now;row[idx.updatedAt]=now;sh.appendRow(row);return {ok:true,message:'บันทึกสิทธิพิเศษสำเร็จ'};
}
function updateVisitSpecialRule(data){
  if(!data||!data.specialRuleId) throw new Error('ไม่พบ specialRuleId');
  const sh=ensureVisitSpecialRuleSheet_();const vals=sh.getDataRange().getValues();const idx=headerMap_(vals[0]);
  for(let i=1;i<vals.length;i++){if(safeString_(vals[i][idx.specialRuleId])===safeString_(data.specialRuleId)){sh.getRange(i+1,idx.ruleName+1).setValue(data.ruleName);sh.getRange(i+1,idx.gradeLevelScope+1).setValue(normalizeGradeLevelScope_(data.gradeLevelScope||'ALL'));sh.getRange(i+1,idx.extraQuota+1).setValue(Number(data.extraQuota));sh.getRange(i+1,idx.startDate+1).setValue(formatDateOnly_(toDate_(data.startDate)));sh.getRange(i+1,idx.endDate+1).setValue(formatDateOnly_(toDate_(data.endDate)));sh.getRange(i+1,idx.note+1).setValue(data.note||'');sh.getRange(i+1,idx.status+1).setValue(data.status||'active');sh.getRange(i+1,idx.updatedAt+1).setValue(formatDateTime_(new Date()));return {ok:true,message:'แก้ไขสิทธิพิเศษสำเร็จ'};}}
  throw new Error('ไม่พบสิทธิพิเศษ');
}
function deleteVisitSpecialRule(specialRuleId){ return updateVisitSpecialRule({specialRuleId:specialRuleId,ruleName:'',extraQuota:0,startDate:formatDateOnly_(new Date()),endDate:formatDateOnly_(new Date()),status:'inactive'}); }

function getVisitUsageByMonth_(youthId, month, year){
  ensureVisitBookingSheet_();
  const rules=getMonthlyApplicableSpecialRules_(youthId, month, year);
  const sh=getSpreadsheet().getSheetByName('visit_booking');const vals=sh.getDataRange().getValues();const idx=headerMap_(vals[0]);
  const bookings=[];
  vals.slice(1).forEach(function(r){
    if(safeString_(r[idx.youthId])!==youthId) return;
    if(normalizeText_(r[idx.status]||'active')!=='active') return;
    const d=toDate_(r[idx.bookingDate]);
    if((d.getMonth()+1)!==month || d.getFullYear()!==year) return;
    bookings.push({bookingId:safeString_(r[idx.bookingId]),bookingDate:formatDateOnly_(d),quotaTypeUsed:safeString_(r[idx.quotaTypeUsed]),specialRuleId:safeString_(r[idx.specialRuleId]),note:safeString_(r[idx.note]),updatedAt:safeString_(r[idx.updatedAt])});
  });
  const specialUsed=bookings.filter(function(b){return b.quotaTypeUsed==='SPECIAL';}).length;
  const normalUsed=bookings.filter(function(b){return b.quotaTypeUsed==='NORMAL';}).length;
  const extraAvailable=rules.reduce(function(a,b){return a+Number(b.extraQuota||0);},0);
  return { bookings:bookings, usedTotal:bookings.length, specialUsed:specialUsed, normalUsed:normalUsed, extraAvailable:extraAvailable, applicableRules:rules, bookingDates:bookings.map(function(b){return b.bookingDate;}), note:bookings.length?'':'ยังไม่มีการจอง' };
}

function getMonthlyApplicableSpecialRules_(youthId, month, year){
  var youth=getYouthRowsForVisit_().find(function(y){return y.youthId===youthId;});
  if(!youth) return [];
  var grade=safeString_(youth.gradeLevel).trim();
  var start=new Date(year, month-1, 1);
  var end=new Date(year, month, 0);
  return getVisitSpecialRules().filter(function(r){
    if(normalizeText_(r.status)!=='active') return false;
    var rs=toDate_(r.startDate), re=toDate_(r.endDate);
    if (re < start || rs > end) return false;
    var scopeNorm=normalizeGradeLevelScope_(r.gradeLevelScope||'ALL');
    if(scopeNorm==='ALL') return true;
    var arr=scopeNorm.split(',').map(function(v){return safeString_(v).trim();});
    return arr.indexOf(grade)>-1;
  });
}

function getApplicableSpecialRules_(youthId, targetDate){
  const youth=getYouthRowsForVisit_().find(function(y){return y.youthId===youthId;});if(!youth)return[];
  const grade=youth.gradeLevel;const t=toDate_(targetDate);
  return getVisitSpecialRules().filter(function(r){
    if(normalizeText_(r.status)!=='active') return false;
    if(t<toDate_(r.startDate) || t>toDate_(r.endDate)) return false;
    const scopeNorm=normalizeGradeLevelScope_(r.gradeLevelScope||'ALL');
    if(scopeNorm==='ALL') return true;
    const arr=scopeNorm.split(',').map(function(v){return safeString_(v).trim();});
    return arr.indexOf(safeString_(grade).trim())>-1;
  });
}

function calculateVisitQuotaStatus_(youthId, targetDate){
  const d=toDate_(targetDate);const month=d.getMonth()+1;const year=d.getFullYear();
  const youth=getYouthRowsForVisit_().find(function(y){return y.youthId===youthId;});if(!youth) return {canBook:false,note:'ไม่พบเด็ก'};
  const usage=getVisitUsageByMonth_(youthId,month,year);const base=getBaseQuotaForGrade_(youth.gradeLevel);
  const specialRules=getApplicableSpecialRules_(youthId,targetDate);
  let specialRemaining=Math.max(0,specialRules.reduce(function(a,b){return a+Number(b.extraQuota||0);},0)-usage.specialUsed);
  let normalRemaining=Math.max(0,base-usage.normalUsed);
  if(specialRemaining>0) return {canBook:true,quotaTypeUsed:'SPECIAL',specialRuleId:specialRules[0]?specialRules[0].specialRuleId:'',remainingAfter:(specialRemaining-1)+normalRemaining,note:'จองได้ โดยใช้สิทธิพิเศษ "'+(specialRules[0]?specialRules[0].ruleName:'')+'"'};
  if(normalRemaining>0) return {canBook:true,quotaTypeUsed:'NORMAL',specialRuleId:'',remainingAfter:normalRemaining-1,note:'จองได้ โดยใช้สิทธิพื้นฐาน'};
  return {canBook:false,quotaTypeUsed:'',specialRuleId:'',remainingAfter:0,note:'จองไม่ได้ เพราะใช้สิทธิครบแล้ว'};
}

function createVisitBooking(data){
  if(!safeString_(data.youthId)) throw new Error('กรุณาเลือกเด็ก');
  if(!safeString_(data.bookingDate)) throw new Error('กรุณาเลือกวันที่จอง');
  const status=calculateVisitQuotaStatus_(data.youthId,data.bookingDate);
  if(!status.canBook) throw new Error(status.note);
  const sh=ensureVisitBookingSheet_();const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];const idx=headerMap_(h);const now=formatDateTime_(new Date());const row=new Array(h.length).fill('');
  row[idx.bookingId]=generateId('VBK');row[idx.youthId]=data.youthId;row[idx.bookingDate]=formatDateOnly_(toDate_(data.bookingDate));row[idx.quotaTypeUsed]=status.quotaTypeUsed;row[idx.specialRuleId]=status.specialRuleId||'';row[idx.note]=data.note||'';row[idx.importedFrom]=data.importedFrom||'MANUAL';row[idx.createdAt]=now;row[idx.updatedAt]=now;row[idx.status]='active';sh.appendRow(row);
  return {ok:true,message:'บันทึกการจองสำเร็จ',quota:status};
}

function updateVisitBooking(data){
  if(!data||!data.bookingId) throw new Error('ไม่พบ bookingId');
  const sh=ensureVisitBookingSheet_();const vals=sh.getDataRange().getValues();const idx=headerMap_(vals[0]);
  for(let i=1;i<vals.length;i++){
    if(safeString_(vals[i][idx.bookingId])===safeString_(data.bookingId)){
      const yid=data.youthId||safeString_(vals[i][idx.youthId]);const bdate=data.bookingDate||safeString_(vals[i][idx.bookingDate]);
      const quota=calculateVisitQuotaStatus_(yid,bdate);if(!quota.canBook) throw new Error(quota.note);
      sh.getRange(i+1,idx.youthId+1).setValue(yid);sh.getRange(i+1,idx.bookingDate+1).setValue(formatDateOnly_(toDate_(bdate)));sh.getRange(i+1,idx.quotaTypeUsed+1).setValue(quota.quotaTypeUsed);sh.getRange(i+1,idx.specialRuleId+1).setValue(quota.specialRuleId||'');sh.getRange(i+1,idx.note+1).setValue(data.note||'');sh.getRange(i+1,idx.updatedAt+1).setValue(formatDateTime_(new Date()));
      return {ok:true,message:'แก้ไขการจองสำเร็จ'};
    }
  }
  throw new Error('ไม่พบการจอง');
}

function deleteVisitBooking(bookingId){
  const sh=ensureVisitBookingSheet_();const vals=sh.getDataRange().getValues();const idx=headerMap_(vals[0]);
  for(let i=1;i<vals.length;i++) if(safeString_(vals[i][idx.bookingId])===safeString_(bookingId)){sh.getRange(i+1,idx.status+1).setValue('deleted');sh.getRange(i+1,idx.updatedAt+1).setValue(formatDateTime_(new Date()));return {ok:true,message:'ลบการจองสำเร็จ'};}
  throw new Error('ไม่พบการจอง');
}

function deleteAllVisitBookings(criteria){
  criteria=criteria||{};if(!criteria.deleteScope) throw new Error('ไม่พบเงื่อนไขการลบ');
  const sh=ensureVisitBookingSheet_();const vals=sh.getDataRange().getValues();const idx=headerMap_(vals[0]);
  const month=Number(criteria.month||0),year=Number(criteria.year||0);const yset=criteria.youthIds&&criteria.youthIds.length?nullToSet_(criteria.youthIds):null;
  let count=0;
  for(let i=1;i<vals.length;i++){
    if(normalizeText_(vals[i][idx.status]||'active')!=='active') continue;
    const d=toDate_(vals[i][idx.bookingDate]);
    let ok=false;
    if(criteria.deleteScope==='MONTH_YEAR') ok=((d.getMonth()+1)===month && d.getFullYear()===year);
    else if(criteria.deleteScope==='MONTH_YEAR_YOUTH') ok=((d.getMonth()+1)===month && d.getFullYear()===year && yset[safeString_(vals[i][idx.youthId])]);
    else if(criteria.deleteScope==='ALL') ok=true;
    if(ok){sh.getRange(i+1,idx.status+1).setValue('deleted');sh.getRange(i+1,idx.updatedAt+1).setValue(formatDateTime_(new Date()));count++;}
  }
  return {ok:true,message:'ลบการจองสำเร็จ',deletedCount:count,month:month,year:year,scope:criteria.deleteScope};
}

function importVisitBookingsPreview(base64Data,fileName,mimeType,defaultBookingDate){
  if(!base64Data||!fileName) throw new Error('ข้อมูลไฟล์ไม่ครบ');
  const ext=(fileName.split('.').pop()||'').toLowerCase();if(['xlsx','xls','csv'].indexOf(ext)===-1) throw new Error('รองรับ .xlsx .xls .csv');
  const blob=Utilities.newBlob(Utilities.base64Decode(base64Data),mimeType||'application/octet-stream',fileName);
  const rows=ext==='csv'?Utilities.parseCsv(blob.getDataAsString('UTF-8')):convertExcelToSheetRows_(blob,fileName).rows;
  if(!rows||rows.length<2) throw new Error('ไม่พบข้อมูลในไฟล์');
  const h=(rows[0]||[]).map(function(v){return normalizeText_(v);});
  const nameIdx=Math.max(h.indexOf(normalizeText_('ชื่อ-นามสกุลเด็กและเยาวชน')),h.indexOf('fullname'));
  const dateIdx=h.indexOf(normalizeText_('bookingDate'))>-1?h.indexOf(normalizeText_('bookingDate')):h.indexOf(normalizeText_('วันที่จอง'));
  if(nameIdx<0) throw new Error('ไม่พบคอลัมน์ชื่อ');
  const preview=[];
  rows.slice(1).forEach(function(r,i){
    const excelName=safeString_(r[nameIdx]).trim();if(!excelName)return;
    const matched=matchYouthName_(excelName);const bookingDate=safeString_(r[dateIdx]||defaultBookingDate||'').trim();
    let quota={canBook:false,note:'ยังไม่ระบุวันที่'};
    if(matched.youthId&&bookingDate) quota=calculateVisitQuotaStatus_(matched.youthId,bookingDate);
    preview.push({rowNo:i+2,excelName:excelName,youthId:matched.youthId||'',matchedName:matched.fullName||'',bookingDate:bookingDate,canBook:quota.canBook,quotaTypeUsed:quota.quotaTypeUsed||'',specialRuleId:quota.specialRuleId||'',note:quota.note||matched.note||''});
  });
  return {items:preview,summary:{total:preview.length,canBook:preview.filter(function(i){return i.canBook;}).length,noMatch:preview.filter(function(i){return !i.youthId;}).length,cannotBook:preview.filter(function(i){return i.youthId && !i.canBook;}).length}};
}

function confirmImportVisitBookings(payload){
  const items=(payload&&payload.items)||[];let inserted=0,skipped=0,errors=[];
  items.forEach(function(it){
    try{ if(!it.keep) {skipped++; return;} if(!it.youthId||!it.bookingDate){skipped++; return;} const q=calculateVisitQuotaStatus_(it.youthId,it.bookingDate); if(!q.canBook){skipped++; return;} createVisitBooking({youthId:it.youthId,bookingDate:it.bookingDate,note:it.note||'',importedFrom:'EXCEL'}); inserted++; }
    catch(e){errors.push({rowNo:it.rowNo,error:e.message});}
  });
  return {ok:true,inserted:inserted,skipped:skipped,errorCount:errors.length,errors:errors};
}

function matchYouthName_(inputName){
  const n=normalizeText_(inputName).replace(/\s+/g,' ');if(!n)return{youthId:'',fullName:'',note:'ไม่พบชื่อ'};
  const youth=getYouthRowsForVisit_();
  let exact=youth.find(function(y){return normalizeText_(y.fullName).replace(/\s+/g,' ')===n;});
  if(exact) return {youthId:exact.youthId,fullName:exact.fullName,note:'ตรงอัตโนมัติ'};
  let partial=youth.filter(function(y){return normalizeText_(y.fullName).indexOf(n)>-1 || n.indexOf(normalizeText_(y.fullName))>-1;});
  if(partial.length===1) return {youthId:partial[0].youthId,fullName:partial[0].fullName,note:'จับคู่แบบใกล้เคียง'};
  return {youthId:'',fullName:'',note:'ไม่พบชื่อในระบบ'};
}

function getDistinctGradeLevels_(){
  var set={};
  getYouthRowsForVisit_().forEach(function(y){ var g=safeString_(y.gradeLevel).trim(); if(g) set[g]=true; });
  return Object.keys(set).sort();
}
function normalizeGradeLevelScope_(scopeRaw){
  var t=safeString_(scopeRaw).trim();
  if(!t) return 'ALL';
  if(normalizeText_(t)==='all') return 'ALL';
  var arr=t.split(',').map(function(v){return safeString_(v).trim();}).filter(Boolean);
  var uniq={};
  arr.forEach(function(v){uniq[v]=true;});
  return Object.keys(uniq).join(',');
}

function validateGradeLevel_(value){
  var v=safeString_(value).trim();
  var allowed=getDistinctGradeLevels_();
  if(allowed.indexOf(v)===-1) throw new Error('ระดับชั้นไม่อยู่ในระบบ: '+v);
  return v;
}

function getAvailableGradeLevels_(){
  return getDistinctGradeLevels_();
}

function getYouthRowsForVisit_(){
  ensureYouthSheet_();
  const sh=getSpreadsheet().getSheetByName('youth');const vals=sh.getDataRange().getValues();if(vals.length<=1)return[];const idx=headerMap_(vals[0]);
  return vals.slice(1).filter(function(r){return safeString_(r[idx.youthId]);}).map(function(r){return{youthId:safeString_(r[idx.youthId]),fullName:safeString_(r[idx.fullName]),gradeLevel:normalizeGradeLevel_(r[idx.gradeLevel]),advisorTeacher:safeString_(r[idx.advisorTeacher])};});
}
function getBaseQuotaForGrade_(gradeLevel){
  const g=normalizeGradeLevel_(gradeLevel);const rule=getVisitQuotaRules().find(function(r){return normalizeGradeLevel_(r.gradeLevel)===g && normalizeText_(r.status)==='active';});
  return rule?Number(rule.allowedPerMonth||0):0;
}

function deleteNormalQuota(ruleId){ return deleteVisitQuotaRule(ruleId); }
function deleteSpecialQuota(ruleId){ return deleteVisitSpecialRule(ruleId); }
function deleteBooking(bookingId){ return deleteVisitBooking(bookingId); }
function deleteAllBookings(criteria){ return deleteAllVisitBookings(criteria); }
