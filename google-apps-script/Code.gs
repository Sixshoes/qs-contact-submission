/**
 * QS 聯絡人提交 → 共用試算表 Pool
 *
 * 設計：所有單位／人員都提交到「同一本」試算表，以追加列方式累積（pool）。
 * 每一列最後一欄為「時間戳」（台北時間）。
 *
 * 設定步驟（用 yitingchen@gm.fgu.edu.tw 登入）：
 * 1. 開啟 https://sheets.google.com 新增試算表，命名例如「QS 2028 聯絡人 Pool」
 * 2. 擴充功能 → Apps Script，清空預設程式碼，貼上本檔全部內容並儲存
 * 3. 部署 → 新增部署作業 → 網頁應用程式
 *    - 執行身分：我
 *    - 具有存取權的使用者：任何人
 * 4. 授權後複製網址，設到網站環境變數 PUBLIC_GOOGLE_SCRIPT_URL
 *
 * 若先前已建過舊版工作表（時間戳在最左），請刪除「提交紀錄／學術聯絡人／雇主聯絡人」
 * 三個工作表後再提交一次，讓腳本依新欄位順序重建。
 */

var NOTIFY_TO = [
  'yitingchen@mail.fgu.edu.tw',
  'chlchang@mail.fgu.edu.tw',
];

var SHEET_SUBMISSIONS = '提交紀錄';
var SHEET_ACADEMIC = '學術聯絡人';
var SHEET_EMPLOYER = '雇主聯絡人';
var TZ = 'Asia/Taipei';

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) || '{}';
    var data = JSON.parse(raw);
    var result = appendSubmission_(data);
    maybeNotify_(data, result);
    return jsonOut_({ ok: true, submissionId: result.submissionId, timestamp: result.timestamp });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doGet() {
  return jsonOut_({
    ok: true,
    service: 'QS contact pool receiver',
    hint: 'Anyone can POST; rows append to one shared spreadsheet pool.',
  });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

/** 台北時間戳，例如 2026-08-25 08:44:12 */
function formatTimestamp_(date) {
  return Utilities.formatDate(date || new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function appendSubmission_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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

  // Pool：一律 append，不覆蓋既有列
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
    '已有單位寫入共用試算表 Pool（追加一筆，不覆蓋舊資料）。',
    '',
    '提交單位：' + unit,
    '提交人：' + submitter,
    '提交編號：' + result.submissionId,
    '學術筆數：' + result.academicCount,
    '雇主筆數：' + result.employerCount,
    '時間戳：' + result.timestamp,
    '',
    '請開啟綁定本腳本的 Google 試算表查看「學術聯絡人」「雇主聯絡人」工作表。',
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
