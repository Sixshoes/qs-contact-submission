/**
 * QS 聯絡人提交 → 共用試算表 Pool
 *
 * ★ 重要：若腳本是「獨立專案」（不是從試算表 → 擴充功能開啟），
 *   getActiveSpreadsheet() 會是空的，導致寫不進去。
 *   本版會自動記住／建立試算表；也可手動填 SHEET_ID。
 *
 * 設定：
 * 1. 把本檔貼到 Apps Script 並儲存
 * 2.（建議）把試算表網址中 /d/XXXX/edit 的 XXXX 貼到下方 SHEET_ID
 * 3. 部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署
 *    （執行身分：我；誰可以存取：所有人）
 */

/** 可選：試算表 ID（網址 https://docs.google.com/spreadsheets/d/【這裡】/edit） */
var SHEET_ID = '';

var NOTIFY_TO = [
  'yitingchen@mail.fgu.edu.tw',
  'chlchang@mail.fgu.edu.tw',
];

var SHEET_SUBMISSIONS = '提交紀錄';
var SHEET_ACADEMIC = '學術聯絡人';
var SHEET_EMPLOYER = '雇主聯絡人';
var TZ = 'Asia/Taipei';
var PROP_SHEET_ID = 'SHEET_ID';

function doPost(e) {
  try {
    var raw = extractPayload_(e);
    var data = JSON.parse(raw);
    var result = appendSubmission_(data);
    maybeNotify_(data, result);
    return jsonOut_({
      ok: true,
      submissionId: result.submissionId,
      timestamp: result.timestamp,
      spreadsheetUrl: result.spreadsheetUrl,
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doGet() {
  var ss = null;
  var url = '';
  try {
    ss = getSpreadsheet_();
    url = ss.getUrl();
  } catch (err) {
    /* ignore */
  }
  return jsonOut_({
    ok: true,
    service: 'QS contact pool receiver',
    spreadsheetUrl: url,
    hint: 'Anyone can POST; rows append to one shared spreadsheet pool.',
  });
}

function extractPayload_(e) {
  if (e && e.parameter && e.parameter.payload) {
    return String(e.parameter.payload);
  }
  if (e && e.postData && e.postData.contents) {
    var contents = String(e.postData.contents);
    // text/plain JSON
    if (contents.charAt(0) === '{' || contents.charAt(0) === '[') {
      return contents;
    }
    // application/x-www-form-urlencoded
    if (e.parameter && e.parameter.payload) {
      return String(e.parameter.payload);
    }
  }
  return '{}';
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function formatTimestamp_(date) {
  return Utilities.formatDate(date || new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

/**
 * 優先順序：程式內 SHEET_ID → ScriptProperties → 綁定試算表 → 新建一本
 */
function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = String(SHEET_ID || '').trim() || props.getProperty(PROP_SHEET_ID);

  if (id) {
    return SpreadsheetApp.openById(id);
  }

  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    props.setProperty(PROP_SHEET_ID, active.getId());
    return active;
  }

  var created = SpreadsheetApp.create('QS 2028 聯絡人 Pool');
  props.setProperty(PROP_SHEET_ID, created.getId());
  return created;
}

function appendSubmission_(data) {
  var ss = getSpreadsheet_();
  ensureSheets_(ss);

  var unit = String(data.unit || '').trim();
  var submitter = String(data.submitter || '').trim();
  var stamp = formatTimestamp_(new Date());
  var academic = Array.isArray(data.academic) ? data.academic : [];
  var employer = Array.isArray(data.employer) ? data.employer : [];
  var submissionId = String(data.submissionId || Utilities.getUuid());

  if (!unit || !submitter) {
    throw new Error('缺少提交單位或提交人姓名');
  }
  if (!academic.length && !employer.length) {
    throw new Error('至少需要一筆學術或雇主聯絡人');
  }

  ss.getSheetByName(SHEET_SUBMISSIONS).appendRow([
    submissionId,
    unit,
    submitter,
    academic.length,
    employer.length,
    stamp,
  ]);

  var academicSheet = ss.getSheetByName(SHEET_ACADEMIC);
  academic.forEach(function (row) {
    academicSheet.appendRow([
      submissionId,
      unit,
      submitter,
      row.Source || '',
      row.Title || '',
      row['First Name'] || '',
      row['Last Name'] || '',
      row['Job Title'] || '',
      row.Department || '',
      row.Institution || '',
      row['Country or Territory'] || '',
      row.Email || '',
      row.Subject || '',
      row['Phone (Optional)'] || '',
      stamp,
    ]);
  });

  var employerSheet = ss.getSheetByName(SHEET_EMPLOYER);
  employer.forEach(function (row) {
    employerSheet.appendRow([
      submissionId,
      unit,
      submitter,
      row.Source || '',
      row.Title || '',
      row['First Name'] || '',
      row['Last Name'] || '',
      row.Position || '',
      row.Industry || '',
      row['Company Name'] || '',
      row['Country or Territory'] || '',
      row.Email || '',
      row['Phone (Optional)'] || '',
      stamp,
    ]);
  });

  return {
    submissionId: submissionId,
    academicCount: academic.length,
    employerCount: employer.length,
    timestamp: stamp,
    spreadsheetUrl: ss.getUrl(),
  };
}

function ensureSheets_(ss) {
  ensureSheet_(ss, SHEET_SUBMISSIONS, [
    '提交編號',
    '提交單位',
    '提交人',
    '學術筆數',
    '雇主筆數',
    '時間戳',
  ]);
  ensureSheet_(ss, SHEET_ACADEMIC, [
    '提交編號',
    '提交單位',
    '提交人',
    'Source',
    'Title',
    'First Name',
    'Last Name',
    'Job Title',
    'Department',
    'Institution',
    'Country or Territory',
    'Email',
    'Subject',
    'Phone (Optional)',
    '時間戳',
  ]);
  ensureSheet_(ss, SHEET_EMPLOYER, [
    '提交編號',
    '提交單位',
    '提交人',
    'Source',
    'Title',
    'First Name',
    'Last Name',
    'Position',
    'Industry',
    'Company Name',
    'Country or Territory',
    'Email',
    'Phone (Optional)',
    '時間戳',
  ]);
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
}

function maybeNotify_(data, result) {
  if (!NOTIFY_TO || !NOTIFY_TO.length) return;
  var unit = String(data.unit || '');
  var submitter = String(data.submitter || '');
  var subject = '【QS聯絡人提報】' + unit + '－' + submitter;
  var body = [
    '已有單位寫入共用試算表 Pool。',
    '',
    '提交單位：' + unit,
    '提交人：' + submitter,
    '提交編號：' + result.submissionId,
    '學術筆數：' + result.academicCount,
    '雇主筆數：' + result.employerCount,
    '時間戳：' + result.timestamp,
    '試算表：' + (result.spreadsheetUrl || ''),
  ].join('\n');

  try {
    MailApp.sendEmail({
      to: NOTIFY_TO.join(','),
      subject: subject,
      body: body,
    });
  } catch (err) {
    // 通知失敗不影響寫入
  }
}
