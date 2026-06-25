// 광수중학교 2026 3학년 진로 희망분야 제출 - Apps Script 백엔드
const SS_ID = '1dNPXDzXoktxhGj3Hb0z0J1Si5kir4sOLDUkx3XTotXg';
const SHEET_NAME = '시트1';

const COL = { ID: 1, GRADE: 2, BAN: 3, NUM: 4, NAME: 5, FIELD: 6, JOB: 7, TIME: 8 };

function norm(s) {
  return String(s == null ? '' : s).normalize('NFC').replace(/\s+/g, '').trim();
}

function getSheet_() {
  return SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_NAME);
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  let data;
  try {
    if (action === 'verify') data = verifyStudent(e.parameter.id, e.parameter.name);
    else if (action === 'submit') data = submitChoice(e.parameter.id, e.parameter.name, e.parameter.field, e.parameter.job);
    else if (action === 'list') data = listRecords();
    else data = { error: 'unknown action' };
  } catch (err) {
    data = { error: err.toString() };
  }
  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(data) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function findRow_(id, name) {
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const targetId = norm(id);
  const targetName = norm(name);
  for (let i = 1; i < rows.length; i++) {
    const rowId = norm(rows[i][COL.ID - 1]);
    const rowName = norm(rows[i][COL.NAME - 1]);
    if (rowId === targetId && rowName === targetName) {
      return { rowIndex: i + 1, row: rows[i] };
    }
  }
  return null;
}

function verifyStudent(id, name) {
  const found = findRow_(id, name);
  if (!found) return { ok: false, error: '학번 또는 이름이 일치하지 않습니다. 다시 확인해주세요.' };
  const row = found.row;
  return {
    ok: true,
    grade: row[COL.GRADE - 1],
    ban: row[COL.BAN - 1],
    num: row[COL.NUM - 1],
    name: row[COL.NAME - 1],
    alreadySubmitted: !!row[COL.FIELD - 1],
    field: row[COL.FIELD - 1] || '',
    job: row[COL.JOB - 1] || ''
  };
}

function submitChoice(id, name, field, job) {
  const found = findRow_(id, name);
  if (!found) return { ok: false, error: '학번 또는 이름이 일치하지 않습니다.' };
  if (!field || !job) return { ok: false, error: '분야와 직업을 모두 선택해주세요.' };
  const sheet = getSheet_();
  sheet.getRange(found.rowIndex, COL.FIELD).setValue(field);
  sheet.getRange(found.rowIndex, COL.JOB).setValue(job);
  sheet.getRange(found.rowIndex, COL.TIME).setValue(new Date());
  return { ok: true };
}

function listRecords() {
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[COL.ID - 1]) continue;
    result.push({
      id: r[COL.ID - 1],
      grade: r[COL.GRADE - 1],
      ban: r[COL.BAN - 1],
      num: r[COL.NUM - 1],
      name: r[COL.NAME - 1],
      field: r[COL.FIELD - 1] || '',
      job: r[COL.JOB - 1] || '',
      time: r[COL.TIME - 1] ? Utilities.formatDate(new Date(r[COL.TIME - 1]), 'GMT+9', 'yyyy-MM-dd HH:mm') : ''
    });
  }
  return { ok: true, records: result };
}
