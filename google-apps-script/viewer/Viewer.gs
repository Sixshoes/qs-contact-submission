/**
 * QS 聯絡人 Pool 瀏覽（永續辦內部後台）
 *
 * 與提交用 Code.gs 分開部署：此頁「誰可以存取」請設為只有我自己（或校內帳號）。
 * 前提：Pool 試算表已共用給此腳本「執行身分」帳號（至少檢視者）。
 *
 * 設定步驟：
 * 1. 用後台帳號登入 script.google.com，新增專案（勿與公開提交腳本混用）
 * 2. 貼上本檔到「程式碼.gs」
 * 3. 檔案 → 新增 → HTML，檔名填 Pool（不要 .html），貼上 Pool.html
 * 4. 下方 SHEET_ID 填 Pool 試算表 ID（網址 /d/XXXX/edit 的 XXXX）
 * 5. 部署 → 新增部署 → 網頁應用程式（執行身分：我；誰可以存取：只有我自己）
 * 6. 更新 Pool.html 後需「新版本」重新部署才會生效
 */

/** 試算表 ID（與提交 Pool 同一本） */
var SHEET_ID = '';

var SHEET_SUBMISSIONS = '提交紀錄';
var SHEET_ACADEMIC = '學術聯絡人';
var SHEET_EMPLOYER = '雇主聯絡人';
var TZ = 'Asia/Taipei';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Pool')
    .setTitle('QS 聯絡人 Pool')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 供 Pool.html 呼叫 */
function getPoolData() {
  var id = String(SHEET_ID || '').trim();
  if (!id) {
    return { ok: false, error: '請在 Viewer.gs 設定 SHEET_ID（試算表 ID）' };
  }

  try {
    var ss = SpreadsheetApp.openById(id);
    return {
      ok: true,
      submissions: sheetToMatrix_(ss.getSheetByName(SHEET_SUBMISSIONS)),
      academic: sheetToMatrix_(ss.getSheetByName(SHEET_ACADEMIC)),
      employer: sheetToMatrix_(ss.getSheetByName(SHEET_EMPLOYER)),
      updatedAt: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
      spreadsheetUrl: ss.getUrl(),
    };
  } catch (err) {
    return {
      ok: false,
      error: String(err && err.message ? err.message : err),
    };
  }
}

function sheetToMatrix_(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];
  return sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
}
