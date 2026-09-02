import ExcelJS from 'exceljs';
import {
  applyExcelCellValue,
  defaultBodyFont,
  defaultGuideBodyFont,
  defaultGuideTitleFont,
  defaultHeaderFont,
} from './excel-fonts.js';

const HEADER_FONT = defaultHeaderFont();
const BODY_FONT = defaultBodyFont();
const HEADER_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF1D4ED8' },
};
const ZEBRA_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFF1F5F9' },
};
const THIN_BORDER = {
  top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
};

/** 等待瀏覽器開始下載後再 revoke，避免第二個檔被吃掉 */
const DOWNLOAD_REVOKE_MS = 2500;
/** 連續兩個檔案下載之間的間隔 */
export const DOWNLOAD_SEQUENCE_GAP_MS = 1800;

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/**
 * Build a styled worksheet that stays merge-friendly:
 * - no merged cells in the data grid
 * - frozen header + AutoFilter
 * - consistent rectangular table
 */
function styleSheet(ws, headers, rows) {
  ws.views = [{ state: 'frozen', ySplit: 1, activeCell: 'A2' }];

  const headerRow = ws.getRow(1);
  headers.forEach((text, i) => {
    const cell = headerRow.getCell(i + 1);
    applyExcelCellValue(cell, text, HEADER_FONT);
    cell.fill = HEADER_FILL;
    cell.border = THIN_BORDER;
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
  headerRow.height = 24;

  rows.forEach((rowValues, rowIdx) => {
    const row = ws.getRow(rowIdx + 2);
    rowValues.forEach((value, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      applyExcelCellValue(cell, value ?? '', BODY_FONT);
      cell.border = THIN_BORDER;
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      if (rowIdx % 2 === 1) cell.fill = ZEBRA_FILL;
    });
    row.height = 20;
  });

  headers.forEach((header, i) => {
    let maxLen = String(header).length;
    for (const row of rows) {
      const len = String(row[i] ?? '').length;
      if (len > maxLen) maxLen = len;
    }
    ws.getColumn(i + 1).width = Math.min(Math.max(maxLen + 2, 12), 36);
  });

  if (headers.length && rows.length >= 0) {
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: Math.max(1, rows.length + 1), column: headers.length },
    };
  }

  ws.properties.defaultRowHeight = 18;
}

/** @param {'academic'|'employer'} kind */
function addMergeGuideSheet(workbook, kind) {
  const label = kind === 'academic' ? '學術' : '雇主';
  const guide = workbook.addWorksheet('合併說明', {
    properties: { tabColor: { argb: 'FF64748B' } },
  });
  const guideLines = [
    ['合併使用說明 / How to merge'],
    [''],
    [`1. 各單位下載的${label}檔「English」分頁欄位順序完全相同，可直接往下貼上合併。`],
    ['2. 建議以 English 分頁作為最終上傳 QS Hub 的版本（欄位名稱符合官方範本）。'],
    ['3. 「中文」與「English」列數、順序一一對應，方便核對；請勿插入／刪除單一分頁的列。'],
    ['4. 合併時請保留第 1 列表頭，只複製第 2 列起的資料列。'],
    ['5. 表頭已開啟篩選（AutoFilter），可用篩選檢查缺漏後再合併。'],
    [''],
    [`1. English sheets in ${kind} files share the same column order — append rows to merge.`],
    ['2. Use the English sheet for QS Hub upload (official header names).'],
    ['3. Chinese and English sheets stay row-aligned for checking; do not insert/delete rows on only one sheet.'],
    ['4. Keep row 1 headers; copy data from row 2 downward when merging.'],
    ['5. AutoFilter is enabled on header row for quick QA before merge.'],
  ];
  guideLines.forEach((line, idx) => {
    const row = guide.getRow(idx + 1);
    applyExcelCellValue(row.getCell(1), line[0], idx === 0 ? defaultGuideTitleFont() : defaultGuideBodyFont());
  });
  guide.getColumn(1).width = 90;
}

async function workbookToBlob(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: XLSX_MIME });
}

/** @param {{ headersZh: string[], headersEn: string[], rowsZh: unknown[][], rowsEn: unknown[][] }} opts @param {'academic'|'employer'} kind */
async function createBilingualWorkbookBlob({ headersZh, headersEn, rowsZh, rowsEn }, kind) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Fo Guang University QS Contact Form';
  workbook.created = new Date();

  const wsZh = workbook.addWorksheet('中文', {
    properties: { tabColor: { argb: 'FFDC2626' } },
  });
  styleSheet(wsZh, headersZh, rowsZh);

  const wsEn = workbook.addWorksheet('English', {
    properties: { tabColor: { argb: 'FF1D4ED8' } },
  });
  styleSheet(wsEn, headersEn, rowsEn);

  addMergeGuideSheet(workbook, kind);
  return workbookToBlob(workbook);
}

/** @param {Blob} blob @param {string} filename */
export function downloadBlob(blob, filename) {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.rel = 'noopener';
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    window.setTimeout(() => {
      a.remove();
      URL.revokeObjectURL(url);
      resolve();
    }, DOWNLOAD_REVOKE_MS);
  });
}

/** @param {{ filename: string, blob: Blob }[]} entries */
export async function downloadBlobsSequential(entries) {
  for (let i = 0; i < entries.length; i++) {
    if (i > 0) {
      await new Promise((r) => window.setTimeout(r, DOWNLOAD_SEQUENCE_GAP_MS));
    }
    await downloadBlob(entries[i].blob, entries[i].filename);
  }
}

/** @param {{ headersZh: string[], headersEn: string[], rowsZh: unknown[][], rowsEn: unknown[][] }} opts */
export async function createAcademicWorkbookBlob(opts) {
  return createBilingualWorkbookBlob(opts, 'academic');
}

/** @param {{ headersZh: string[], headersEn: string[], rowsZh: unknown[][], rowsEn: unknown[][] }} opts */
export async function createEmployerWorkbookBlob(opts) {
  return createBilingualWorkbookBlob(opts, 'employer');
}
