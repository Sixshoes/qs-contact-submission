import ExcelJS from 'exceljs';

const HEADER_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF1D4ED8' },
};
const HEADER_FONT = {
  name: 'Calibri',
  size: 11,
  bold: true,
  color: { argb: 'FFFFFFFF' },
};
const BODY_FONT = {
  name: 'Calibri',
  size: 11,
  color: { argb: 'FF0F172A' },
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
    cell.value = text;
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.border = THIN_BORDER;
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
  headerRow.height = 24;

  rows.forEach((rowValues, rowIdx) => {
    const row = ws.getRow(rowIdx + 2);
    rowValues.forEach((value, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      cell.value = value ?? '';
      cell.font = BODY_FONT;
      cell.border = THIN_BORDER;
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      if (rowIdx % 2 === 1) cell.fill = ZEBRA_FILL;
    });
    row.height = 20;
  });

  // Column widths from header + sample content (capped for readability)
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

  // Mark used range so copy/paste & Power Query merges cleanly
  ws.properties.defaultRowHeight = 18;
}

async function workbookToDownload(workbook, filename) {
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

/** Single-sheet academic / generic export */
export async function downloadStyledSheet({ headers, rows, filename, sheetName = 'Contacts' }) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Fo Guang University QS Contact Form';
  workbook.created = new Date();
  const ws = workbook.addWorksheet(sheetName, {
    properties: { tabColor: { argb: 'FF1D4ED8' } },
  });
  styleSheet(ws, headers, rows);
  await workbookToDownload(workbook, filename);
}

/**
 * Employer export: Chinese + English sheets with identical column order,
 * so rows line up 1:1 for easy side-by-side merge / QA.
 */
export async function downloadEmployerStyledWorkbook({
  headersZh,
  headersEn,
  rowsZh,
  rowsEn,
  filename,
}) {
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

  // Guide sheet for merging multiple unit files
  const guide = workbook.addWorksheet('合併說明', {
    properties: { tabColor: { argb: 'FF64748B' } },
  });
  const guideLines = [
    ['合併使用說明 / How to merge'],
    [''],
    ['1. 各單位下載的雇主檔「English」分頁欄位順序完全相同，可直接往下貼上合併。'],
    ['2. 建議以 English 分頁作為最終上傳 QS Hub 的版本（欄位名稱符合官方範本）。'],
    ['3. 「中文」與「English」列數、順序一一對應，方便核對；請勿插入／刪除單一分頁的列。'],
    ['4. 合併時請保留第 1 列表頭，只複製第 2 列起的資料列。'],
    ['5. 表頭已開啟篩選（AutoFilter），可用篩選檢查缺漏後再合併。'],
    [''],
    ['1. English sheets across unit files share the same column order — append rows to merge.'],
    ['2. Use the English sheet for QS Hub upload (official header names).'],
    ['3. Chinese and English sheets stay row-aligned for checking; do not insert/delete rows on only one sheet.'],
    ['4. Keep row 1 headers; copy data from row 2 downward when merging.'],
    ['5. AutoFilter is enabled on header row for quick QA before merge.'],
  ];
  guideLines.forEach((line, idx) => {
    const row = guide.getRow(idx + 1);
    row.getCell(1).value = line[0];
    row.getCell(1).font =
      idx === 0
        ? { name: 'Calibri', size: 14, bold: true, color: { argb: 'FF0F172A' } }
        : { name: 'Calibri', size: 11, color: { argb: 'FF334155' } };
  });
  guide.getColumn(1).width = 90;

  await workbookToDownload(workbook, filename);
}
