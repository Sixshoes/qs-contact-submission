import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';
import {
  applyExcelCellValue,
  defaultBodyFont,
  defaultGuideBodyFont,
  defaultGuideTitleFont,
  defaultHeaderFont,
} from '../src/scripts/excel-fonts.js';
import {
  SOURCE_EN,
  TITLES,
  ACADEMIC_JOB_TITLES,
  EMPLOYER_JOB_TITLES,
  INDUSTRIES,
  COUNTRIES,
  SUBJECTS,
  DEFAULT_COUNTRY,
  ACADEMIC_EXCEL_HEADERS_EN,
  EMPLOYER_EXCEL_HEADERS_EN,
} from '../src/data/qs-options.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.resolve(__dirname, '../public/templates');
const outFile = path.join(outDir, 'QS_聯絡人匯入樣板.xlsx');

/** 資料列數（第 2 列起） */
const DATA_ROW_COUNT = 100;
const FIRST_DATA_ROW = 2;
const LAST_DATA_ROW = FIRST_DATA_ROW + DATA_ROW_COUNT - 1;

const HEADER_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF1D4ED8' },
};
const HEADER_FONT = defaultHeaderFont();
const BODY_FONT = defaultBodyFont();
const THIN_BORDER = {
  top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
};

function sheetListRange(sheetName, itemCount) {
  const safe = String(sheetName).replace(/'/g, "''");
  return `'${safe}'!$A$2:$A$${itemCount + 1}`;
}

function styleDataSheet(ws, headers, rows, { fillDataRows = 0 } = {}) {
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
    });
    row.height = 20;
  });

  if (fillDataRows > 0) {
    const firstEmptyRow = FIRST_DATA_ROW + rows.length;
    const lastRow = FIRST_DATA_ROW + fillDataRows - 1;
    for (let row = firstEmptyRow; row <= lastRow; row++) {
      headers.forEach((_, colIdx) => {
        const cell = ws.getCell(row, colIdx + 1);
        cell.value = '';
        cell.font = { ...BODY_FONT };
        cell.border = THIN_BORDER;
        cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      });
      ws.getRow(row).height = 20;
    }
  }

  headers.forEach((header, i) => {
    let maxLen = String(header).length;
    for (const row of rows) {
      const len = String(row[i] ?? '').length;
      if (len > maxLen) maxLen = len;
    }
    ws.getColumn(i + 1).width = Math.min(Math.max(maxLen + 2, 12), 42);
  });

  if (headers.length) {
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: Math.max(1, rows.length + 1), column: headers.length },
    };
  }
}

function applyColumnListValidation(ws, colIndex, rangeFormula, { allowBlank = true, prompt } = {}) {
  const validation = {
    type: 'list',
    allowBlank,
    formulae: [rangeFormula],
    showErrorMessage: false,
    showInputMessage: Boolean(prompt),
    promptTitle: prompt?.title || '請選擇或輸入',
    prompt: prompt?.text || '可從下拉選單選擇，或直接輸入；選單外文字會自動視為 Other',
  };

  for (let row = FIRST_DATA_ROW; row <= LAST_DATA_ROW; row++) {
    ws.getCell(row, colIndex).dataValidation = { ...validation };
  }
}

function addGuideSheet(workbook) {
  const guide = workbook.addWorksheet('填寫說明', {
    properties: { tabColor: { argb: 'FF64748B' } },
  });

  const lines = [
    ['QS 聯絡人匯入樣板 — 填寫說明'],
    [''],
    ['一、用途'],
    ['本檔供各單位先離線整理聯絡人資料，欄位與線上提交系統／QS 官方格式一致。'],
    ['填寫完成後，請至線上提交網站「從 Excel 匯入」上傳本檔；通過驗證後可匯入表單並提交。'],
    [''],
    ['二、工作表說明'],
    ['• 「學術聯絡人」「雇主聯絡人」：選項欄位有下拉提示（▼），也可直接輸入；選單外文字匯入時會自動視為 Other。'],
    ['• 第 2 列為範例，填寫前請刪除；可從第 2 列起填寫，最多 100 筆。'],
    ['• 其他分頁為選項對照表（下拉選單來源），亦可手動查閱。'],
    [''],
    ['三、共通規則'],
    ['• Source 請選 Fo Guang University。'],
    ['• Title／職稱／產業／學科／國家等：可從下拉選擇，或直接輸入自訂內容（系統會自動當作 Other）。'],
    ['• 姓名、系所、機構、Email、電話請手動輸入。'],
    ['• Email：個人信箱，勿用 team@、info@ 等共用信箱；不可重複。'],
    ['• Phone 可留空；若填寫請含開頭 0（例如 0912345678）。'],
    [''],
    ['四、提交資訊（網站另填，不在本檔）'],
    ['• 提交單位、提交人姓名：請在線上系統第一步填寫。'],
    [''],
    ['若有疑問請聯絡永續發展辦公室。'],
  ];

  lines.forEach((line, idx) => {
    const row = guide.getRow(idx + 1);
    row.getCell(1).value = line[0];
    row.getCell(1).font = idx === 0 ? defaultGuideTitleFont() : defaultGuideBodyFont();
  });
  guide.getColumn(1).width = 95;
}

/** @returns {{ name: string, count: number, range: string }} */
function addOptionSheet(workbook, name, items, tabColor) {
  const ws = workbook.addWorksheet(name, { properties: { tabColor: { argb: tabColor } } });
  const headers = ['英文（下拉選單值）', '中文對照'];
  styleDataSheet(
    ws,
    headers,
    items.map((item) => [item.en, item.zh]),
  );
  return { name, count: items.length, range: sheetListRange(name, items.length) };
}

function colIndex(headers, headerName) {
  const idx = headers.indexOf(headerName);
  if (idx < 0) throw new Error(`找不到欄位：${headerName}`);
  return idx + 1;
}

function applyAcademicValidations(ws, ranges) {
  const H = ACADEMIC_EXCEL_HEADERS_EN;
  applyColumnListValidation(ws, colIndex(H, 'Source'), `"${SOURCE_EN}"`, {
    allowBlank: false,
    prompt: { title: 'Source', text: '請選擇 Fo Guang University' },
  });
  applyColumnListValidation(ws, colIndex(H, 'Title'), ranges.titles.range, {
    prompt: { title: 'Title', text: '可選擇或輸入；選單外會自動視為 Other' },
  });
  applyColumnListValidation(ws, colIndex(H, 'Job Title'), ranges.academicJobs.range, {
    prompt: { title: 'Job Title', text: '可選擇或輸入；選單外會自動視為 Other' },
  });
  applyColumnListValidation(ws, colIndex(H, 'Country or Territory'), ranges.countries.range, {
    prompt: { title: 'Country', text: '可選擇或輸入；選單外會自動視為 Other' },
  });
  applyColumnListValidation(ws, colIndex(H, 'Subject'), ranges.subjects.range, {
    prompt: { title: 'Subject', text: '可選擇或輸入；選單外會自動視為 Other' },
  });
}

function applyEmployerValidations(ws, ranges) {
  const H = EMPLOYER_EXCEL_HEADERS_EN;
  applyColumnListValidation(ws, colIndex(H, 'Source'), `"${SOURCE_EN}"`, {
    allowBlank: false,
    prompt: { title: 'Source', text: '請選擇 Fo Guang University' },
  });
  applyColumnListValidation(ws, colIndex(H, 'Title'), ranges.titles.range, {
    prompt: { title: 'Title', text: '可選擇或輸入；選單外會自動視為 Other' },
  });
  applyColumnListValidation(ws, colIndex(H, 'Position'), ranges.employerJobs.range, {
    prompt: { title: 'Position', text: '可選擇或輸入；選單外會自動視為 Other' },
  });
  applyColumnListValidation(ws, colIndex(H, 'Industry'), ranges.industries.range, {
    prompt: { title: 'Industry', text: '可選擇或輸入；選單外會自動視為 Other' },
  });
  applyColumnListValidation(ws, colIndex(H, 'Country or Territory'), ranges.countries.range, {
    prompt: { title: 'Country', text: '可選擇或輸入；選單外會自動視為 Other' },
  });
}

async function main() {
  fs.mkdirSync(outDir, { recursive: true });

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Fo Guang University QS Contact Form';
  workbook.created = new Date();

  addGuideSheet(workbook);

  const ranges = {
    titles: addOptionSheet(workbook, '稱謂', TITLES, 'FF94A3B8'),
    academicJobs: addOptionSheet(workbook, '學術職稱', ACADEMIC_JOB_TITLES, 'FF94A3B8'),
    employerJobs: addOptionSheet(workbook, '雇主職位', EMPLOYER_JOB_TITLES, 'FF94A3B8'),
    industries: addOptionSheet(workbook, '產業', INDUSTRIES, 'FF94A3B8'),
    subjects: addOptionSheet(workbook, '學科領域', SUBJECTS, 'FF94A3B8'),
    countries: addOptionSheet(workbook, '國家或地區', COUNTRIES, 'FF94A3B8'),
  };

  const academicExample = [
    SOURCE_EN,
    'Dr',
    'Ming',
    'Chen',
    'Professor/Associate Professor',
    'Department of History',
    SOURCE_EN,
    DEFAULT_COUNTRY,
    '（請改為實際 email）',
    'History',
    '0912345678',
  ];

  const employerExample = [
    SOURCE_EN,
    'Ms',
    'Jane',
    'Doe',
    'Manager/Executive',
    'Education',
    'Example Company Ltd.',
    DEFAULT_COUNTRY,
    '（請改為實際 email）',
    '0912345678',
  ];

  const wsAcademic = workbook.addWorksheet('學術聯絡人', {
    properties: { tabColor: { argb: 'FFDC2626' } },
  });
  styleDataSheet(wsAcademic, ACADEMIC_EXCEL_HEADERS_EN, [academicExample], {
    fillDataRows: DATA_ROW_COUNT,
  });
  applyAcademicValidations(wsAcademic, ranges);

  const wsEmployer = workbook.addWorksheet('雇主聯絡人', {
    properties: { tabColor: { argb: 'FFF59E0B' } },
  });
  styleDataSheet(wsEmployer, EMPLOYER_EXCEL_HEADERS_EN, [employerExample], {
    fillDataRows: DATA_ROW_COUNT,
  });
  applyEmployerValidations(wsEmployer, ranges);

  const buffer = await workbook.xlsx.writeBuffer();
  fs.writeFileSync(outFile, Buffer.from(buffer));
  console.log(`Wrote ${outFile}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
