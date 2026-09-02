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
    ['• 「學術聯絡人」「雇主聯絡人」：請填實際資料；灰色斜體列為範例，填寫前請刪除。'],
    ['• 「稱謂」「學術職稱」「雇主職位」「產業」「學科領域」「國家或地區」：僅供查閱，請複製英文全名填入。'],
    [''],
    ['三、共通規則'],
    ['• Source（來源）請固定填：Fo Guang University'],
    ['• 欄位請填英文官方選項（與 QS 範本一致）；「其他」請在網站提交時另行說明。'],
    ['• Email：請填個人信箱或個人公務信箱，勿用 team@、info@ 等共用信箱；每位聯絡人僅一組、且不可重複。'],
    ['• Phone：可留空；若填寫請含開頭 0（例如 0912345678）。'],
    ['• Country or Territory 預設可填：Taiwan SAR'],
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

function addOptionSheet(workbook, name, items, tabColor) {
  const ws = workbook.addWorksheet(name, { properties: { tabColor: { argb: tabColor } } });
  const headers = ['英文（請複製此欄）', '中文對照'];
  styleDataSheet(
    ws,
    headers,
    items.map((item) => [item.en, item.zh]),
  );
}

async function main() {
  fs.mkdirSync(outDir, { recursive: true });

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Fo Guang University QS Contact Form';
  workbook.created = new Date();

  addGuideSheet(workbook);

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

  const wsEmployer = workbook.addWorksheet('雇主聯絡人', {
    properties: { tabColor: { argb: 'FFF59E0B' } },
  });
  styleDataSheet(wsEmployer, EMPLOYER_EXCEL_HEADERS_EN, [employerExample], { exampleRowCount: 1 });

  addOptionSheet(workbook, '稱謂', TITLES, 'FF94A3B8');
  addOptionSheet(workbook, '學術職稱', ACADEMIC_JOB_TITLES, 'FF94A3B8');
  addOptionSheet(workbook, '雇主職位', EMPLOYER_JOB_TITLES, 'FF94A3B8');
  addOptionSheet(workbook, '產業', INDUSTRIES, 'FF94A3B8');
  addOptionSheet(workbook, '學科領域', SUBJECTS, 'FF94A3B8');
  addOptionSheet(workbook, '國家或地區', COUNTRIES, 'FF94A3B8');

  const buffer = await workbook.xlsx.writeBuffer();
  fs.writeFileSync(outFile, Buffer.from(buffer));
  console.log(`Wrote ${outFile}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
