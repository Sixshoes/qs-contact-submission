/**
 * QS 聯絡人 Pool 瀏覽（yitingchen@gm.fgu.edu.tw 專用）
 *
 * 前提：wl03112816 那邊的 Pool 試算表已「共用」給 yitingchen@gm.fgu.edu.tw（至少檢視者）。
 *
 * 設定步驟：
 * 1. 用 yitingchen 登入 script.google.com，開啟「View」專案
 * 2. 把本檔內容貼到「程式碼.gs」（覆蓋 myFunction）
 * 3. 檔案 → 新增 → HTML，檔名填 Pool（不要 .html）
 * 4. 把 Pool.html 全部貼進 Pool 檔
 * 5. 下方 SHEET_ID 填 wl03112816 試算表 ID（網址 /d/XXXX/edit 的 XXXX）
 * 6. 儲存 → 部署 → 新增部署 → 網頁應用程式
 *    - 執行身分：我
 *    - 誰可以存取：只有我自己
 * 7. 用 yitingchen 登入的瀏覽器開啟部署網址
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
