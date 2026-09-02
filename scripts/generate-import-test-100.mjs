/**
 * 產生 100 筆匯入測試用 Excel，並執行與網站相同的 parseImportWorkbook 驗證。
 *
 * 用法：npm run generate:import-test
 * 輸出：test-fixtures/QS_匯入測試_學術100筆.xlsx
 *       test-fixtures/QS_匯入測試_雇主100筆.xlsx
 *       test-fixtures/QS_匯入測試_學術雇主各100筆.xlsx
 */
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
import { parseImportWorkbook } from '../src/scripts/excel-import.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outDirs = [
  path.resolve(__dirname, '../test-fixtures'),
  path.resolve(__dirname, '../public/test-fixtures'),
];
const COUNT = 100;

function noOther(list) {
  return list.filter((item) => item.en !== 'Other');
}

const SAFE_TITLES = noOther(TITLES);
const SAFE_ACADEMIC_JOBS = noOther(ACADEMIC_JOB_TITLES);
const SAFE_EMPLOYER_JOBS = noOther(EMPLOYER_JOB_TITLES);
const SAFE_INDUSTRIES = noOther(INDUSTRIES);
const SAFE_COUNTRIES = noOther(COUNTRIES);
const SAFE_SUBJECTS = noOther(SUBJECTS);
function pad(n, width = 3) {
  return String(n).padStart(width, '0');
}

function pick(list, i) {
  return list[i % list.length].en;
}

function academicRow(i) {
  const n = pad(i);
  return [
    SOURCE_EN,
    pick(SAFE_TITLES, i),
    `Test${n}`,
    `Academic${n}`,
    pick(SAFE_ACADEMIC_JOBS, i),
    `Department of Test ${n}`,
    SOURCE_EN,
    DEFAULT_COUNTRY,
    `qs.test.academic.${n}@fgu.edu.tw`,
    pick(SAFE_SUBJECTS, i),
    `09${String(10000000 + i).slice(-8)}`,
  ];
}

function employerRow(i) {
  const n = pad(i);
  return [
    SOURCE_EN,
    pick(SAFE_TITLES, i + 3),
    `Test${n}`,
    `Employer${n}`,
    pick(SAFE_EMPLOYER_JOBS, i),
    pick(SAFE_INDUSTRIES, i),
    `Test Company ${n} Ltd.`,
    pick(SAFE_COUNTRIES, i),
    `qs.test.employer.${n}@fgu.edu.tw`,
    `09${String(20000000 + i).slice(-8)}`,
  ];
}

async function addSheet(workbook, sheetName, headers, rows) {
  const ws = workbook.addWorksheet(sheetName);
  ws.addRow(headers);
  rows.forEach((row) => ws.addRow(row));
}

async function writeWorkbook(filename, { academicRows = [], employerRows = [] }) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'QS Import Test Generator';

  await addSheet(workbook, '學術聯絡人', ACADEMIC_EXCEL_HEADERS_EN, academicRows);
  await addSheet(workbook, '雇主聯絡人', EMPLOYER_EXCEL_HEADERS_EN, employerRows);

  const buffer = await workbook.xlsx.writeBuffer();
  const written = [];
  for (const outDir of outDirs) {
    fs.mkdirSync(outDir, { recursive: true });
    const outFile = path.join(outDir, filename);
    fs.writeFileSync(outFile, Buffer.from(buffer));
    written.push(outFile);
  }
  return written[0];
}

async function validateFile(filePath) {
  const buffer = fs.readFileSync(filePath);
  const result = await parseImportWorkbook(buffer);
  return { filePath, ...result };
}

async function main() {
  for (const outDir of outDirs) fs.mkdirSync(outDir, { recursive: true });

  const academicRows = Array.from({ length: COUNT }, (_, i) => academicRow(i + 1));
  const employerRows = Array.from({ length: COUNT }, (_, i) => employerRow(i + 1));

  const files = [
    await writeWorkbook('QS_匯入測試_學術100筆.xlsx', { academicRows }),
    await writeWorkbook('QS_匯入測試_雇主100筆.xlsx', { employerRows }),
    await writeWorkbook('QS_匯入測試_學術雇主各100筆.xlsx', { academicRows, employerRows }),
  ];

  console.log('已產生測試檔：\n');
  for (const file of files) {
    const result = await validateFile(file);
    const status = result.ok ? '✓ 驗證通過' : '✗ 驗證失敗';
    console.log(`${status}  ${path.basename(file)}`);
    console.log(`  學術 ${result.summary.academicCount} 筆、雇主 ${result.summary.employerCount} 筆`);
    if (!result.ok) {
      result.errors.slice(0, 5).forEach((e) => {
        console.log(`  - [${e.sheet}] 第 ${e.row} 列 ${e.field}: ${e.message}`);
      });
      if (result.errors.length > 5) console.log(`  …另有 ${result.errors.length - 5} 項錯誤`);
    }
    console.log(`  ${file}\n`);
  }

  console.log('手動測試：至網站 Step 1「從 Excel 匯入」上傳上述檔案即可。');
  console.log('線上下載：https://sixshoes.github.io/qs-contact-submission/test-fixtures/');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
