import ExcelJS from 'exceljs';
import {
  ACADEMIC_EXCEL_HEADERS_EN,
  EMPLOYER_EXCEL_HEADERS_EN,
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

function isEmptyDataRow(values) {
  const [firstName = '', lastName = '', email = ''] = [
    values[2],
    values[3],
    values[8] ?? values[8],
  ];
  return !trimVal(firstName) && !trimVal(lastName) && !trimVal(email);
}

function rowToMap(headers, values) {
  const map = {};
  headers.forEach((h, i) => {
    map[h] = trimVal(values[i]);
  });
  return map;
}

function parseTitle(raw) {
  const m = matchOption(TITLES, raw);
  if (m.invalid) return { title: '', titleOther: '', rawTitle: raw, invalid: true };
  return { title: m.en, titleOther: m.other, rawTitle: raw, invalid: false };
}

function parseAcademicRow(map, rowNum) {
  const title = parseTitle(map.Title);
  const job = matchOption(ACADEMIC_JOB_TITLES, map['Job Title'], { pluralOther: true });
  const country = matchOption(COUNTRIES, map['Country or Territory']);
  const subject = matchOption(SUBJECTS, map.Subject, { pluralOther: true });

  const contact = {
    _source: map.Source,
    title: title.title,
    titleOther: title.titleOther,
    _rawTitle: title.rawTitle,
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

  if (title.invalid) {
    return {
      contact,
      errors: [{ sheet: SHEET_ACADEMIC, row: rowNum, field: 'Title', message: `「${title.rawTitle}」不在選項清單內` }],
    };
  }
  if (job.invalid) {
    return {
      contact,
      errors: [{ sheet: SHEET_ACADEMIC, row: rowNum, field: 'Job Title', message: `「${job.raw}」不在選項清單內` }],
    };
  }
  if (country.invalid) {
    return {
      contact,
      errors: [{ sheet: SHEET_ACADEMIC, row: rowNum, field: 'Country or Territory', message: `「${country.raw}」不在選項清單內` }],
    };
  }
  if (subject.invalid) {
    return {
      contact,
      errors: [{ sheet: SHEET_ACADEMIC, row: rowNum, field: 'Subject', message: `「${subject.raw}」不在選項清單內` }],
    };
  }

  return { contact, errors: validateImportContact(contact, 'academic', rowNum, OPTS) };
}

function parseEmployerRow(map, rowNum) {
  const title = parseTitle(map.Title);
  const job = matchOption(EMPLOYER_JOB_TITLES, map.Position, { pluralOther: true });
  const industry = matchOption(INDUSTRIES, map.Industry);
  const country = matchOption(COUNTRIES, map['Country or Territory']);

  const contact = {
    _source: map.Source,
    title: title.title,
    titleOther: title.titleOther,
    _rawTitle: title.rawTitle,
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

  if (title.invalid) {
    return {
      contact,
      errors: [{ sheet: SHEET_EMPLOYER, row: rowNum, field: 'Title', message: `「${title.rawTitle}」不在選項清單內` }],
    };
  }
  if (job.invalid) {
    return {
      contact,
      errors: [{ sheet: SHEET_EMPLOYER, row: rowNum, field: 'Position', message: `「${job.raw}」不在選項清單內` }],
    };
  }
  if (industry.invalid) {
    return {
      contact,
      errors: [{ sheet: SHEET_EMPLOYER, row: rowNum, field: 'Industry', message: `「${industry.raw}」不在選項清單內` }],
    };
  }
  if (country.invalid) {
    return {
      contact,
      errors: [{ sheet: SHEET_EMPLOYER, row: rowNum, field: 'Country or Territory', message: `「${country.raw}」不在選項清單內` }],
    };
  }

  return { contact, errors: validateImportContact(contact, 'employer', rowNum, OPTS) };
}

function parseSheet(workbook, sheetName, expectedHeaders, type) {
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
  if (!headersMatch(headers, expectedHeaders)) {
    return {
      contacts: [],
      errors: [
        {
          sheet: sheetName,
          row: 1,
          field: '表頭',
          message: `表頭欄位不符，請使用匯入樣板（預期：${expectedHeaders.join('、')}）`,
        },
      ],
    };
  }

  /** @type {import('./contact-validation.js').ImportContact[]} */
  const contacts = [];
  /** @type {{ sheet: string, row: number, field: string, message: string }[]} */
  const errors = [];

  for (let i = 1; i < matrix.length; i++) {
    const rowNum = i + 1;
    const values = matrix[i];
    if (isEmptyDataRow(values) || isPlaceholderRow(values)) continue;

    const map = rowToMap(expectedHeaders, values);
    const parsed = type === 'academic' ? parseAcademicRow(map, rowNum) : parseEmployerRow(map, rowNum);
    if (parsed.errors.length) {
      errors.push(...parsed.errors);
    } else {
      contacts.push(parsed.contact);
    }
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

/**
 * @param {ArrayBuffer} buffer
 * @param {{ maxContacts?: number }} [options]
 */
export async function parseImportWorkbook(buffer, options = {}) {
  const maxContacts = options.maxContacts ?? MAX_CONTACTS;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const academicResult = parseSheet(workbook, SHEET_ACADEMIC, ACADEMIC_EXCEL_HEADERS_EN, 'academic');
  const employerResult = parseSheet(workbook, SHEET_EMPLOYER, EMPLOYER_EXCEL_HEADERS_EN, 'employer');

  /** @type {{ sheet: string, row: number, field: string, message: string }[]} */
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

/** @param {File} file @param {{ maxContacts?: number }} [options] */
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
