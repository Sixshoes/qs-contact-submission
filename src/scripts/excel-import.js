import ExcelJS from 'exceljs';
import {
  ACADEMIC_EXCEL_HEADERS_EN,
  EMPLOYER_EXCEL_HEADERS_EN,
  ACADEMIC_IMPORT_TEMPLATE_HEADERS,
  EMPLOYER_IMPORT_TEMPLATE_HEADERS,
  TITLES,
  ACADEMIC_JOB_TITLES,
  EMPLOYER_JOB_TITLES,
  INDUSTRIES,
  COUNTRIES,
  SUBJECTS,
  SOURCE_EN,
} from '../data/qs-options.js';
import {
  trimVal,
  matchOption,
  validateImportContact,
  normalizeImportContact,
  normalizeEmail,
} from './contact-validation.js';

const MAX_CONTACTS = 100;

const SHEET_ACADEMIC = '學術聯絡人';
const SHEET_EMPLOYER = '雇主聯絡人';

const PLACEHOLDER_HINTS = ['請改為實際', 'example', '範例', '（請改', '(請改'];

const OPTS = {
  TITLES,
  ACADEMIC_JOB_TITLES,
  EMPLOYER_JOB_TITLES,
  INDUSTRIES,
  COUNTRIES,
  SUBJECTS,
  SOURCE_EN,
};

function cellText(cell) {
  if (!cell || cell.value == null) return '';
  const v = cell.value;
  if (typeof v === 'object' && v !== null) {
    if ('text' in v && v.text != null) return String(v.text);
    if ('result' in v && v.result != null) return String(v.result);
    if ('richText' in v && Array.isArray(v.richText)) {
      return v.richText.map((r) => r.text || '').join('');
    }
    if (v instanceof Date) return v.toISOString();
  }
  return String(v);
}

function normalizeHeader(h) {
  return trimVal(h).replace(/\s+/g, ' ').toLowerCase();
}

function headersMatch(actual, expected) {
  const a = actual.map(normalizeHeader);
  const e = expected.map(normalizeHeader);
  if (a.length < e.length) return false;
  return e.every((h, i) => a[i] === h);
}

function detectHeaderFormat(headers, type) {
  if (type === 'academic') {
    if (headersMatch(headers, ACADEMIC_IMPORT_TEMPLATE_HEADERS)) {
      return { headers: ACADEMIC_IMPORT_TEMPLATE_HEADERS, hasOtherCols: true };
    }
    if (headersMatch(headers, ACADEMIC_EXCEL_HEADERS_EN)) {
      return { headers: ACADEMIC_EXCEL_HEADERS_EN, hasOtherCols: false };
    }
  } else {
    if (headersMatch(headers, EMPLOYER_IMPORT_TEMPLATE_HEADERS)) {
      return { headers: EMPLOYER_IMPORT_TEMPLATE_HEADERS, hasOtherCols: true };
    }
    if (headersMatch(headers, EMPLOYER_EXCEL_HEADERS_EN)) {
      return { headers: EMPLOYER_EXCEL_HEADERS_EN, hasOtherCols: false };
    }
  }
  return null;
}

function sheetToMatrix(worksheet) {
  if (!worksheet) return [];
  const rows = [];
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    const vals = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      vals[colNumber - 1] = cellText(cell);
    });
    rows.push(vals);
  });
  if (!rows.length) return [];
  const width = Math.max(...rows.map((r) => r.length));
  return rows.map((r) => {
    const out = [...r];
    while (out.length < width) out.push('');
    return out;
  });
}

function isPlaceholderRow(values) {
  const joined = values.join(' ').toLowerCase();
  return PLACEHOLDER_HINTS.some((h) => joined.includes(h.toLowerCase()));
}

function isEmptyDataRow(map) {
  return !trimVal(map['First Name']) && !trimVal(map['Last Name']) && !trimVal(map.Email);
}

function rowToMap(headers, values) {
  const map = {};
  headers.forEach((h, i) => {
    map[h] = trimVal(values[i]);
  });
  return map;
}

/**
 * 下拉選 Other 時，優先讀取旁邊的 Other 欄；舊版仍支援 Others (說明) 寫在同一格。
 */
function resolveOtherField(map, mainKey, otherKey, list, { pluralOther = false, hasOtherCols = false } = {}) {
  const mainRaw = trimVal(map[mainKey]);
  const parsed = matchOption(list, mainRaw, { pluralOther });

  if (hasOtherCols && (mainRaw === 'Other' || parsed.en === 'Other')) {
    const detail = trimVal(map[otherKey]);
    if (detail) return { en: 'Other', other: detail, invalid: false, raw: mainRaw };
    if (parsed.other) return parsed;
    return { en: 'Other', other: '', invalid: false, raw: mainRaw };
  }

  if (parsed.invalid && mainRaw) return parsed;
  return parsed;
}

function parseAcademicRow(map, rowNum, { hasOtherCols }) {
  const title = resolveOtherField(map, 'Title', 'Title Other', TITLES, { hasOtherCols });
  const job = resolveOtherField(map, 'Job Title', 'Job Title Other', ACADEMIC_JOB_TITLES, {
    pluralOther: true,
    hasOtherCols,
  });
  const country = resolveOtherField(map, 'Country or Territory', 'Country Other', COUNTRIES, { hasOtherCols });
  const subject = resolveOtherField(map, 'Subject', 'Subject Other', SUBJECTS, {
    pluralOther: true,
    hasOtherCols,
  });

  const contact = {
    _source: map.Source,
    title: title.invalid ? '' : title.en,
    titleOther: title.other,
    _rawTitle: title.raw || map.Title,
    firstName: map['First Name'],
    lastName: map['Last Name'],
    jobTitle: job.invalid ? '' : job.en,
    jobOther: job.other,
    _rawJob: job.invalid ? job.raw : '',
    department: map.Department,
    industry: '',
    industryOther: '',
    institution: map.Institution,
    country: country.invalid ? '' : country.en,
    countryOther: country.other,
    _rawCountry: country.invalid ? country.raw : '',
    email: map.Email,
    subject: subject.invalid ? '' : subject.en,
    subjectOther: subject.other,
    _rawSubject: subject.invalid ? subject.raw : '',
    phone: map['Phone (Optional)'],
  };

  const errors = [];
  if (title.invalid) {
    errors.push({ sheet: SHEET_ACADEMIC, row: rowNum, field: 'Title', message: `「${title.raw}」不在選項清單內` });
  }
  if (job.invalid) {
    errors.push({ sheet: SHEET_ACADEMIC, row: rowNum, field: 'Job Title', message: `「${job.raw}」不在選項清單內` });
  }
  if (country.invalid) {
    errors.push({
      sheet: SHEET_ACADEMIC,
      row: rowNum,
      field: 'Country or Territory',
      message: `「${country.raw}」不在選項清單內`,
    });
  }
  if (subject.invalid) {
    errors.push({ sheet: SHEET_ACADEMIC, row: rowNum, field: 'Subject', message: `「${subject.raw}」不在選項清單內` });
  }
  if (errors.length) return { contact, errors };

  return { contact, errors: validateImportContact(contact, 'academic', rowNum, OPTS) };
}

function parseEmployerRow(map, rowNum, { hasOtherCols }) {
  const title = resolveOtherField(map, 'Title', 'Title Other', TITLES, { hasOtherCols });
  const job = resolveOtherField(map, 'Position', 'Position Other', EMPLOYER_JOB_TITLES, {
    pluralOther: true,
    hasOtherCols,
  });
  const industry = resolveOtherField(map, 'Industry', 'Industry Other', INDUSTRIES, { hasOtherCols });
  const country = resolveOtherField(map, 'Country or Territory', 'Country Other', COUNTRIES, { hasOtherCols });

  const contact = {
    _source: map.Source,
    title: title.invalid ? '' : title.en,
    titleOther: title.other,
    _rawTitle: title.raw || map.Title,
    firstName: map['First Name'],
    lastName: map['Last Name'],
    jobTitle: job.invalid ? '' : job.en,
    jobOther: job.other,
    _rawJob: job.invalid ? job.raw : '',
    department: '',
    industry: industry.invalid ? '' : industry.en,
    industryOther: industry.other,
    _rawIndustry: industry.invalid ? industry.raw : '',
    institution: map['Company Name'],
    country: country.invalid ? '' : country.en,
    countryOther: country.other,
    _rawCountry: country.invalid ? country.raw : '',
    email: map.Email,
    subject: '',
    subjectOther: '',
    phone: map['Phone (Optional)'],
  };

  const errors = [];
  if (title.invalid) {
    errors.push({ sheet: SHEET_EMPLOYER, row: rowNum, field: 'Title', message: `「${title.raw}」不在選項清單內` });
  }
  if (job.invalid) {
    errors.push({ sheet: SHEET_EMPLOYER, row: rowNum, field: 'Position', message: `「${job.raw}」不在選項清單內` });
  }
  if (industry.invalid) {
    errors.push({ sheet: SHEET_EMPLOYER, row: rowNum, field: 'Industry', message: `「${industry.raw}」不在選項清單內` });
  }
  if (country.invalid) {
    errors.push({
      sheet: SHEET_EMPLOYER,
      row: rowNum,
      field: 'Country or Territory',
      message: `「${country.raw}」不在選項清單內`,
    });
  }
  if (errors.length) return { contact, errors };

  return { contact, errors: validateImportContact(contact, 'employer', rowNum, OPTS) };
}

function parseSheet(workbook, sheetName, type) {
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) {
    return {
      contacts: [],
      errors: [{ sheet: sheetName, row: 0, field: '工作表', message: `找不到「${sheetName}」工作表` }],
    };
  }

  const matrix = sheetToMatrix(ws);
  if (matrix.length < 1) {
    return { contacts: [], errors: [] };
  }

  const headers = matrix[0];
  const format = detectHeaderFormat(headers, type);
  if (!format) {
    const expected =
      type === 'academic' ? ACADEMIC_IMPORT_TEMPLATE_HEADERS.join('、') : EMPLOYER_IMPORT_TEMPLATE_HEADERS.join('、');
    return {
      contacts: [],
      errors: [
        {
          sheet: sheetName,
          row: 1,
          field: '表頭',
          message: `表頭欄位不符，請下載最新匯入樣板（預期欄位：${expected}…）`,
        },
      ],
    };
  }

  const contacts = [];
  const errors = [];

  for (let i = 1; i < matrix.length; i++) {
    const rowNum = i + 1;
    const values = matrix[i];
    const map = rowToMap(format.headers, values);
    if (isEmptyDataRow(map) || isPlaceholderRow(values)) continue;

    const parsed = type === 'academic'
      ? parseAcademicRow(map, rowNum, { hasOtherCols: format.hasOtherCols })
      : parseEmployerRow(map, rowNum, { hasOtherCols: format.hasOtherCols });

    if (parsed.errors.length) errors.push(...parsed.errors);
    else contacts.push(parsed.contact);
  }

  return { contacts, errors };
}

function checkDuplicateEmails(academic, employer) {
  const errors = [];
  const seen = new Map();

  const check = (list, sheet) => {
    list.forEach((c, idx) => {
      const email = normalizeEmail(c.email);
      if (!email) return;
      const row = idx + 2;
      if (seen.has(email)) {
        const prev = seen.get(email);
        errors.push({
          sheet,
          row,
          field: 'Email',
          message: `Email 與「${prev.sheet}」第 ${prev.row} 列重複（${email}）`,
        });
      } else {
        seen.set(email, { sheet, row });
      }
    });
  };

  check(academic, SHEET_ACADEMIC);
  check(employer, SHEET_EMPLOYER);
  return errors;
}

export async function parseImportWorkbook(buffer, options = {}) {
  const maxContacts = options.maxContacts ?? MAX_CONTACTS;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const academicResult = parseSheet(workbook, SHEET_ACADEMIC, 'academic');
  const employerResult = parseSheet(workbook, SHEET_EMPLOYER, 'employer');

  const errors = [...academicResult.errors, ...employerResult.errors];

  if (academicResult.contacts.length > maxContacts) {
    errors.push({
      sheet: SHEET_ACADEMIC,
      row: 0,
      field: '筆數',
      message: `學術聯絡人超過 ${maxContacts} 筆（目前 ${academicResult.contacts.length} 筆）`,
    });
  }
  if (employerResult.contacts.length > maxContacts) {
    errors.push({
      sheet: SHEET_EMPLOYER,
      row: 0,
      field: '筆數',
      message: `雇主聯絡人超過 ${maxContacts} 筆（目前 ${employerResult.contacts.length} 筆）`,
    });
  }

  if (!academicResult.contacts.length && !employerResult.contacts.length && !errors.length) {
    errors.push({
      sheet: '—',
      row: 0,
      field: '資料',
      message: '檔案中沒有可匯入的聯絡人資料（請刪除範例列並填入實際資料）',
    });
  }

  errors.push(...checkDuplicateEmails(academicResult.contacts, employerResult.contacts));

  const academic = academicResult.contacts.slice(0, maxContacts).map((c) => normalizeImportContact(c, 'academic'));
  const employer = employerResult.contacts.slice(0, maxContacts).map((c) => normalizeImportContact(c, 'employer'));

  return {
    ok: errors.length === 0,
    academic,
    employer,
    errors,
    summary: {
      academicCount: academic.length,
      employerCount: employer.length,
    },
  };
}

export async function parseImportFile(file, options = {}) {
  if (!file) throw new Error('請選擇檔案');
  const name = file.name.toLowerCase();
  if (!name.endsWith('.xlsx')) {
    return {
      ok: false,
      academic: [],
      employer: [],
      errors: [{ sheet: '—', row: 0, field: '檔案', message: '僅支援 .xlsx 格式' }],
      summary: { academicCount: 0, employerCount: 0 },
    };
  }
  const buffer = await file.arrayBuffer();
  return parseImportWorkbook(buffer, options);
}
