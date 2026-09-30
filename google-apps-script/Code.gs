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
 *
 * 行為：
 * - 同年同類型（學術／雇主）Email 已在 Pool → 整批拒絕（前端亦會擋）
 * - POST { "action":"poolEmails" } → 回傳現有信箱（供表單／信箱檢測擋今年重複）
 */

/** 試算表 ID（網址 https://docs.google.com/spreadsheets/d/【這裡】/edit） */
var SHEET_ID = '1pgNg_zFue_RQfiqejRNVcEEVTgi3dOWrLMT2TTX5O0s';

/** 設成 [] 或註解掉就不寄信；需要通知時再填信箱 */
var NOTIFY_TO = [];

var SHEET_SUBMISSIONS = '提交紀錄';
var SHEET_ACADEMIC = '學術聯絡人';
var SHEET_EMPLOYER = '雇主聯絡人';
var TZ = 'Asia/Taipei';
var PROP_SHEET_ID = 'SHEET_ID';

function doPost(e) {
  try {
    var raw = extractPayload_(e);
    var data = JSON.parse(raw);
    if (data && data.action === 'poolEmails') {
      return poolEmailsOut_();
    }
    var result = appendSubmission_(data);
    maybeNotify_(data, result);
    return jsonOut_({
      ok: true,
      submissionId: result.submissionId,
      timestamp: result.timestamp,
      spreadsheetUrl: result.spreadsheetUrl,
      academicCount: result.academicCount,
      employerCount: result.employerCount,
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doGet(e) {
  var action = '';
  try {
    action = String((e && e.parameter && e.parameter.action) || '').trim();
  } catch (err) {
    action = '';
  }

  if (action === 'poolEmails') {
    return poolEmailsOut_();
  }

  var url = '';
  try {
    url = getSpreadsheet_().getUrl();
  } catch (err) {
    /* ignore */
  }
  return jsonOut_({
    ok: true,
    service: 'QS contact pool receiver',
    spreadsheetUrl: url,
    hint: 'POST JSON to append contacts, or POST/GET action=poolEmails for existing emails.',
  });
}

function poolEmailsOut_() {
  try {
    var ss = getSpreadsheet_();
    ensureSheets_(ss);
    return jsonOut_({
      ok: true,
      academic: listPoolEmails_(ss.getSheetByName(SHEET_ACADEMIC)),
      employer: listPoolEmails_(ss.getSheetByName(SHEET_EMPLOYER)),
    });
  } catch (err) {
    return jsonOut_({
      ok: false,
      error: String(err && err.message ? err.message : err),
      academic: [],
      employer: [],
    });
  }
}

function extractPayload_(e) {
  if (e && e.parameter && e.parameter.payload) {
    return String(e.parameter.payload);
  }
  if (e && e.postData && e.postData.contents) {
    var contents = String(e.postData.contents);
    if (contents.charAt(0) === '{' || contents.charAt(0) === '[') {
      return contents;
    }
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

/** 強制當文字寫入，避免電話 0978… 被試算表吃掉開頭 0 */
function asText_(value) {
  var s = String(value == null ? '' : value).trim();
  if (!s) return '';
  return "'" + s;
}

function toHalfWidthAscii_(v) {
  return String(v == null ? '' : v).replace(/[\uFF01-\uFF5E]/g, function (ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xfee0);
  }).replace(/\u3000/g, ' ');
}

/** 寬鬆比對：去空白／全形、去尾點、不分大小寫 */
function emailKeyLoose_(email) {
  return toHalfWidthAscii_(email)
    .replace(/\s+/g, '')
    .replace(/\.+$/g, '')
    .toLowerCase();
}

function cleanCellEmail_(value) {
  var s = String(value == null ? '' : value).trim();
  if (s.charAt(0) === "'") s = s.substring(1).trim();
  return s;
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

function colIndex_(headers, name) {
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim() === name) return i;
  }
  return -1;
}

/** @returns {Array<{email:string,unit:string,submitter:string,firstName:string,lastName:string}>} */
function listPoolEmails_(sheet) {
  var out = [];
  if (!sheet || sheet.getLastRow() < 2) return out;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var emailCol = colIndex_(headers, 'Email');
  var unitCol = colIndex_(headers, '提交單位');
  var submitterCol = colIndex_(headers, '提交人');
  var firstCol = colIndex_(headers, 'First Name');
  var lastCol = colIndex_(headers, 'Last Name');
  if (emailCol < 0) return out;

  var seen = {};
  for (var r = 1; r < data.length; r++) {
    var email = cleanCellEmail_(data[r][emailCol]);
    var key = emailKeyLoose_(email);
    if (!key || seen[key]) continue;
    seen[key] = true;
    out.push({
      email: email,
      unit: unitCol >= 0 ? String(data[r][unitCol] || '').trim() : '',
      submitter: submitterCol >= 0 ? String(data[r][submitterCol] || '').trim() : '',
      firstName: firstCol >= 0 ? String(data[r][firstCol] || '').trim() : '',
      lastName: lastCol >= 0 ? String(data[r][lastCol] || '').trim() : '',
    });
  }
  return out;
}

/** @returns {Object<string, {email:string,unit:string,submitter:string,firstName:string,lastName:string}>} */
function indexPoolEmails_(sheet) {
  var list = listPoolEmails_(sheet);
  var map = {};
  for (var i = 0; i < list.length; i++) {
    map[emailKeyLoose_(list[i].email)] = list[i];
  }
  return map;
}

function findPoolDupes_(rows, index, typeLabel) {
  var hits = [];
  var seenBatch = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var email = cleanCellEmail_(row.Email);
    var key = emailKeyLoose_(email);
    if (!key) continue;
    if (index[key]) {
      var who = index[key].unit || '';
      if (index[key].submitter) who += (who ? '／' : '') + '提交人 ' + index[key].submitter;
      hits.push(typeLabel + ' ' + email + (who ? '（今年已由「' + who + '」提交）' : '（今年已提交過）'));
    } else if (seenBatch[key]) {
      hits.push(typeLabel + ' ' + email + '（本批名單內重複）');
    } else {
      seenBatch[key] = true;
    }
  }
  return hits;
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

  var academicIndex = indexPoolEmails_(ss.getSheetByName(SHEET_ACADEMIC));
  var employerIndex = indexPoolEmails_(ss.getSheetByName(SHEET_EMPLOYER));
  var dupHits = []
    .concat(findPoolDupes_(academic, academicIndex, '學術'))
    .concat(findPoolDupes_(employer, employerIndex, '雇主'));
  if (dupHits.length) {
    throw new Error('Email 與今年 Pool 重複，未寫入：' + dupHits.join('；'));
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
      asText_(row['Phone (Optional)']),
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
      asText_(row['Phone (Optional)']),
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
  var phoneCol = headers.indexOf('Phone (Optional)') + 1;
  if (phoneCol > 0) {
    sheet.getRange(1, phoneCol, Math.max(sheet.getMaxRows(), 1), 1).setNumberFormat('@');
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
