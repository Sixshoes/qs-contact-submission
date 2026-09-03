import ExcelJS from 'exceljs';
import fs from 'node:fs';
import path from 'node:path';

const ROOT = path.resolve(import.meta.dirname, '..');
const OUT = path.join(ROOT, 'src/data/prior-year-contacts.js');

const FILES = [
  {
    pool: 'academic',
    file: 'e:/奕廷/永續辦業務/大學排名/QS/2027/01Institutions/02Statistics/03Contact-Lists/聯絡人資料/Academic Contacts Pool.xlsm',
    sheets: ['已上傳(含時間戳)', '替代名單'],
    roleHeader: 'Job Title',
    orgHeader: 'Institution',
    unitHeader: 'Department',
  },
  {
    pool: 'employer',
    file: 'e:/奕廷/永續辦業務/大學排名/QS/2027/01Institutions/02Statistics/03Contact-Lists/聯絡人資料/Employer Contacts Pool.xlsm',
    sheets: ['已上載名單', '交換名單'],
    roleHeader: 'Position',
    orgHeader: 'Company Name',
    unitHeader: 'Industry',
  },
];

function cellText(cell) {
  if (!cell || cell.value == null) return '';
  const v = cell.value;
  if (typeof v === 'object' && v !== null) {
    if ('text' in v && v.text != null) return String(v.text);
    if ('result' in v && v.result != null) return String(v.result);
    if ('richText' in v && Array.isArray(v.richText)) {
      return v.richText.map((r) => r.text || '').join('');
    }
    if (v instanceof Date) return '';
  }
  return String(v);
}

function clean(v) {
  return String(v ?? '')
    .replace(/^\uFEFF/, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeEmail(v) {
  return clean(v).replace(/\s+/g, '').toLowerCase();
}

function headerMap(row) {
  const map = {};
  row.eachCell({ includeEmpty: true }, (cell, col) => {
    const key = clean(cellText(cell)).replace(/^\uFEFF/, '');
    if (key) map[key] = col;
  });
  return map;
}

const records = [];
const seen = new Set();

for (const spec of FILES) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(spec.file);
  for (const sheetName of spec.sheets) {
    const ws = wb.getWorksheet(sheetName);
    if (!ws) {
      throw new Error(`找不到工作表「${sheetName}」：${spec.file}`);
    }
    const cols = headerMap(ws.getRow(1));
    const emailCol = cols.Email;
    if (!emailCol) throw new Error(`${sheetName} 沒有 Email 欄`);
    ws.eachRow({ includeEmpty: false }, (row, n) => {
      if (n === 1) return;
      const email = normalizeEmail(cellText(row.getCell(emailCol)));
      if (!email || !email.includes('@')) return;
      const rec = {
        email,
        pool: spec.pool,
        sheet: sheetName,
        title: clean(cellText(row.getCell(cols.Title))),
        firstName: clean(cellText(row.getCell(cols['First Name']))),
        lastName: clean(cellText(row.getCell(cols['Last Name']))),
        role: clean(cellText(row.getCell(cols[spec.roleHeader]))),
        unit: clean(cellText(row.getCell(cols[spec.unitHeader]))),
        org: clean(cellText(row.getCell(cols[spec.orgHeader]))),
      };
      const key = [rec.email, rec.pool, rec.firstName, rec.lastName, rec.org].join('|').toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      records.push(rec);
    });
  }
}

const uniqueEmails = new Set(records.map((r) => r.email)).size;
const body = `/** 由 scripts/extract-prior-year-contacts.mjs 產生，請勿手改。 */
export const PRIOR_YEAR_CONTACTS = ${JSON.stringify(records, null, 2)};
`;

fs.writeFileSync(OUT, body, 'utf8');

const html = `<script>
var PRIOR_YEAR_CONTACTS = ${JSON.stringify(records)};
</script>`;

const poolPath = path.join(ROOT, 'google-apps-script/viewer/Pool.html');
let pool = fs.readFileSync(poolPath, 'utf8');
const marker = '<!--PRIOR_YEAR_DATA-->';
if (!pool.includes(marker)) {
  throw new Error('Pool.html 缺少 <!--PRIOR_YEAR_DATA--> 標記');
}
pool = pool.replace(
  /<!--PRIOR_YEAR_DATA-->[\s\S]*?(?=\n    <script>\n      var SUBMIT_SITE)/,
  `${marker}\n    ${html}\n    `,
);
fs.writeFileSync(poolPath, pool, 'utf8');

console.log(`wrote ${records.length} records (${uniqueEmails} unique emails) → ${path.relative(ROOT, OUT)}`);
console.log('updated google-apps-script/viewer/Pool.html');

