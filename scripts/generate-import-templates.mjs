import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';
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
const EXAMPLE_FONT = {
  name: 'Calibri',
  size: 11,
  italic: true,
  color: { argb: 'FF64748B' },
};
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

function styleDataSheet(ws, headers, rows, { exampleRowCount = 0 } = {}) {
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
    const isExample = rowIdx < exampleRowCount;
    rowValues.forEach((value, colIdx) => {
      const cell = row.getCell(colIdx + 1);
      cell.value = value ?? '';
      cell.font = isExample ? EXAMPLE_FONT : BODY_FONT;
      cell.border = THIN_BORDER;
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    });
    row.height = 20;
  });

  headers.forEach((header, i) => {
    let maxLen = String(header).length;
    for (const row of rows) {
      const len = String(row[i] ?? '').length;
      if (len > maxLen) maxLen = len;
    }
    ws.getColumn(i + 1).width = Math.min(Math.max(maxLen + 2, 12), 40);
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
    showErrorMessage: true,
    errorStyle: 'error',
    errorTitle: '選項不正確',
    error: '請從下拉選單選擇有效選項',
    showInputMessage: Boolean(prompt),
    promptTitle: prompt?.title || '請選擇',
    prompt: prompt?.text || '請從下拉選單選擇',
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
    ['• 「學術聯絡人」「雇主聯絡人」：已為選項欄位設定下拉選單（點儲存格右側 ▼ 選擇）。'],
    ['• 灰色斜體列為範例，填寫前請刪除；可從第 2 列起填寫，最多 100 筆。'],
    ['• 其他分頁為選項對照表（下拉選單來源），亦可手動查閱。'],
    [''],
    ['三、共通規則'],
    ['• Source 請選 Fo Guang University。'],
    ['• Title／職稱／產業／學科／國家等請用下拉選單；若選 Other，請改為 Others (說明) 或 Other (說明) 格式。'],
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
    row.getCell(1).font =
      idx === 0
        ? { name: 'Calibri', size: 14, bold: true, color: { argb: 'FF0F172A' } }
        : { name: 'Calibri', size: 11, color: { argb: 'FF334155' } };
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
  applyColumnListValidation(ws, colIndex(ACADEMIC_EXCEL_HEADERS_EN, 'Source'), `"${SOURCE_EN}"`, {
    allowBlank: false,
    prompt: { title: 'Source', text: '請選擇 Fo Guang University' },
  });
  applyColumnListValidation(ws, colIndex(ACADEMIC_EXCEL_HEADERS_EN, 'Title'), ranges.titles.range, {
    prompt: { title: 'Title', text: '請選擇稱謂' },
  });
  applyColumnListValidation(ws, colIndex(ACADEMIC_EXCEL_HEADERS_EN, 'Job Title'), ranges.academicJobs.range, {
    prompt: { title: 'Job Title', text: '請選擇學術職稱' },
  });
  applyColumnListValidation(ws, colIndex(ACADEMIC_EXCEL_HEADERS_EN, 'Country or Territory'), ranges.countries.range, {
    prompt: { title: 'Country', text: '請選擇國家或地區' },
  });
  applyColumnListValidation(ws, colIndex(ACADEMIC_EXCEL_HEADERS_EN, 'Subject'), ranges.subjects.range, {
    prompt: { title: 'Subject', text: '請選擇學科領域' },
  });
}

function applyEmployerValidations(ws, ranges) {
  applyColumnListValidation(ws, colIndex(EMPLOYER_EXCEL_HEADERS_EN, 'Source'), `"${SOURCE_EN}"`, {
    allowBlank: false,
    prompt: { title: 'Source', text: '請選擇 Fo Guang University' },
  });
  applyColumnListValidation(ws, colIndex(EMPLOYER_EXCEL_HEADERS_EN, 'Title'), ranges.titles.range, {
    prompt: { title: 'Title', text: '請選擇稱謂' },
  });
  applyColumnListValidation(ws, colIndex(EMPLOYER_EXCEL_HEADERS_EN, 'Position'), ranges.employerJobs.range, {
    prompt: { title: 'Position', text: '請選擇雇主職位' },
  });
  applyColumnListValidation(ws, colIndex(EMPLOYER_EXCEL_HEADERS_EN, 'Industry'), ranges.industries.range, {
    prompt: { title: 'Industry', text: '請選擇產業' },
  });
  applyColumnListValidation(ws, colIndex(EMPLOYER_EXCEL_HEADERS_EN, 'Country or Territory'), ranges.countries.range, {
    prompt: { title: 'Country', text: '請選擇國家或地區' },
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
  styleDataSheet(wsAcademic, ACADEMIC_EXCEL_HEADERS_EN, [academicExample], { exampleRowCount: 1 });
  applyAcademicValidations(wsAcademic, ranges);

  const wsEmployer = workbook.addWorksheet('雇主聯絡人', {
    properties: { tabColor: { argb: 'FFF59E0B' } },
  });
  styleDataSheet(wsEmployer, EMPLOYER_EXCEL_HEADERS_EN, [employerExample], { exampleRowCount: 1 });
  applyEmployerValidations(wsEmployer, ranges);

  const buffer = await workbook.xlsx.writeBuffer();
  fs.writeFileSync(outFile, Buffer.from(buffer));
  console.log(`Wrote ${outFile}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
