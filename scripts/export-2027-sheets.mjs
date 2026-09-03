import ExcelJS from 'exceljs';
import fs from 'node:fs';
import path from 'node:path';
import { PRIOR_YEAR_CONTACTS } from '../src/data/prior-year-contacts.js';

const outDir = path.resolve(import.meta.dirname, '../google-apps-script/viewer');

const ACADEMIC_HEADERS = [
  '來源工作表',
  'Title',
  'First Name',
  'Last Name',
  'Job Title',
  'Department',
  'Institution',
  'Email',
];
const EMPLOYER_HEADERS = [
  '來源工作表',
  'Title',
  'First Name',
  'Last Name',
  'Position',
  'Industry',
  'Company Name',
  'Email',
];

async function writeBook(filename, headers, rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.addRow(headers);
  for (const row of rows) ws.addRow(row);
  ws.getRow(1).font = { bold: true };
  const dest = path.join(outDir, filename);
  await wb.xlsx.writeFile(dest);
  console.log('wrote', dest, rows.length);
}

const academic = PRIOR_YEAR_CONTACTS.filter((r) => r.pool === 'academic').map((r) => [
  r.sheet,
  r.title,
  r.firstName,
  r.lastName,
  r.role,
  r.unit,
  r.org,
  r.email,
]);
const employer = PRIOR_YEAR_CONTACTS.filter((r) => r.pool === 'employer').map((r) => [
  r.sheet,
  r.title,
  r.firstName,
  r.lastName,
  r.role,
  r.unit,
  r.org,
  r.email,
]);

await writeBook('2027學術聯絡人.xlsx', ACADEMIC_HEADERS, academic);
await writeBook('2027雇主聯絡人.xlsx', EMPLOYER_HEADERS, employer);
