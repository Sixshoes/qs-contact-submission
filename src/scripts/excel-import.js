import ExcelJS from 'exceljs';
import {
  ACADEMIC_EXCEL_HEADERS_EN,
  EMPLOYER_EXCEL_HEADERS_EN,
  ACADEMIC_FILL_HEADERS,
  EMPLOYER_FILL_HEADERS,
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
  resolveFlexibleOption,
  validateImportContact,
  normalizeImportContact,
  normalizeEmail,
  emailKey,
  priorYearEmailError,
} from './contact-validation.js';

const MAX_CONTACTS = 100;

const SHEET_ACADEMIC = '學術聯絡人';
const SHEET_EMPLOYER = '雇主聯絡人';

const PLACEHOLDER_HINTS = ['請改為實際', 'example', '範例', '參考範例', '（請改', '(請改'];

/** 表頭別名 → 內部標準欄名（填寫樣板／舊版 QS 樣板共用） */
const HEADER_ALIASES = [
  [/^(source|來源)$/i, 'Source'],
  [/^(title|稱謂)$/i, 'Title'],
  [/^(title\s*other)$/i, 'Title Other'],
  [/^(first\s*name|名字|名)$/i, 'First Name'],
  [/^(last\s*name|姓氏|姓)$/i, 'Last Name'],
  [/^(job\s*title|職稱)$/i, 'Job Title'],
  [/^(job\s*title\s*other)$/i, 'Job Title Other'],
  [/^(position|職位)$/i, 'Position'],
  [/^(position\s*other)$/i, 'Position Other'],
  [/^(department|系所)$/i, 'Department'],
  [/^(industry|產業)$/i, 'Industry'],
  [/^(industry\s*other)$/i, 'Industry Other'],
  [/^(institution|機構)$/i, 'Institution'],
  [/^(company\s*name|公司名稱|公司)$/i, 'Company Name'],
  [/^(country(\s*or\s*territory)?|國家或地區|國家)$/i, 'Country or Territory'],
  [/^(country\s*other)$/i, 'Country Other'],
  [/^email/i, 'Email'],
  [/^(subject|學術領域|學科)$/i, 'Subject'],
  [/^(subject\s*other)$/i, 'Subject Other'],
  [/^phone/i, 'Phone (Optional)'],
  [/^(電話)/i, 'Phone (Optional)'],
];

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

function canonicalizeHeader(h) {
  const raw = trimVal(h);
  if (!raw) return '';
  const n = normalizeHeader(raw);
  for (const [re, canon] of HEADER_ALIASES) {
    if (re.test(n) || re.test(raw)) return canon;
  }
  return raw;
}

function headersMatch(actual, expected) {
  const a = actual.map(normalizeHeader);
  const e = expected.map(normalizeHeader);
  if (a.length < e.length) return false;
  return e.every((h, i) => a[i] === h);
}

/**
 * 辨識表頭：填寫樣板（無 Source）／舊 QS 官方欄／含 Other 欄
 * @returns {{ canonicalHeaders: string[], hasOtherCols: boolean, defaultSource: boolean } | null}
 */
function detectHeaderFormat(headers, type) {
  const candidates =
    type === 'academic'
      ? [
          { headers: ACADEMIC_FILL_HEADERS, hasOtherCols: false, defaultSource: true },
          { headers: ACADEMIC_IMPORT_TEMPLATE_HEADERS, hasOtherCols: true, defaultSource: false },
          { headers: ACADEMIC_EXCEL_HEADERS_EN, hasOtherCols: false, defaultSource: false },
        ]
      : [
          { headers: EMPLOYER_FILL_HEADERS, hasOtherCols: false, defaultSource: true },
          { headers: EMPLOYER_IMPORT_TEMPLATE_HEADERS, hasOtherCols: true, defaultSource: false },
          { headers: EMPLOYER_EXCEL_HEADERS_EN, hasOtherCols: false, defaultSource: false },
        ];

  for (const c of candidates) {
    if (headersMatch(headers, c.headers)) {
      return {
        canonicalHeaders: c.headers.map(canonicalizeHeader),
        hasOtherCols: c.hasOtherCols,
        defaultSource: c.defaultSource,
      };
    }
  }

  // 模糊：只要必要欄位齊全即可（容忍 Email／Phone 提示文字差異）
  const canonical = headers.map(canonicalizeHeader);
  const required =
    type === 'academic'
      ? ['Title', 'First Name', 'Last Name', 'Job Title', 'Department', 'Institution', 'Country or Territory', 'Email', 'Subject']
      : ['Title', 'First Name', 'Last Name', 'Position', 'Industry', 'Company Name', 'Country or Territory', 'Email'];
  const present = new Set(canonical.filter(Boolean));
  if (required.every((k) => present.has(k))) {
    return {
      canonicalHeaders: canonical,
      hasOtherCols: present.has('Title Other') || present.has('Job Title Other') || present.has('Position Other'),
      defaultSource: !present.has('Source'),
    };
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

function rowToMap(canonicalHeaders, values) {
  const map = {};
  canonicalHeaders.forEach((h, i) => {
    if (!h) return;
    // 同名欄取第一個非空
    const v = trimVal(values[i]);
    if (!(h in map) || (!map[h] && v)) map[h] = v;
  });
  return map;
}

function pickFlexible(map, mainKey, otherKey, list, { pluralOther = false } = {}) {
  const flex = resolveFlexibleOption(list, map[mainKey], { pluralOther });
  if (flex.en === 'Other' && !flex.other && otherKey && trimVal(map[otherKey])) {
    return { ...flex, other: trimVal(map[otherKey]) };
  }
  return flex;
}

function parseAcademicRow(map, rowNum, { hasOtherCols = false } = {}) {
  const title = pickFlexible(map, 'Title', hasOtherCols ? 'Title Other' : null, TITLES);
  const job = pickFlexible(map, 'Job Title', hasOtherCols ? 'Job Title Other' : null, ACADEMIC_JOB_TITLES, {
    pluralOther: true,
  });
  const country = pickFlexible(map, 'Country or Territory', hasOtherCols ? 'Country Other' : null, COUNTRIES);
  const subject = pickFlexible(map, 'Subject', hasOtherCols ? 'Subject Other' : null, SUBJECTS, {
    pluralOther: true,
  });

  const contact = {
    _source: trimVal(map.Source) || SOURCE_EN,
    title: title.en,
    titleOther: title.other,
    _rawTitle: title.raw,
    firstName: map['First Name'],
    lastName: map['Last Name'],
    jobTitle: job.en,
    jobOther: job.other,
    _rawJob: job.raw,
    department: map.Department,
    industry: '',
    industryOther: '',
    institution: map.Institution,
    country: country.en,
    countryOther: country.other,
    _rawCountry: country.raw,
    email: map.Email,
    subject: subject.en,
    subjectOther: subject.other,
    _rawSubject: subject.raw,
    phone: map['Phone (Optional)'],
  };

  return { contact, errors: validateImportContact(contact, 'academic', rowNum, OPTS) };
}

function parseEmployerRow(map, rowNum, { hasOtherCols = false } = {}) {
  const title = pickFlexible(map, 'Title', hasOtherCols ? 'Title Other' : null, TITLES);
  const job = pickFlexible(map, 'Position', hasOtherCols ? 'Position Other' : null, EMPLOYER_JOB_TITLES, {
    pluralOther: true,
  });
  const industry = pickFlexible(map, 'Industry', hasOtherCols ? 'Industry Other' : null, INDUSTRIES);
  const country = pickFlexible(map, 'Country or Territory', hasOtherCols ? 'Country Other' : null, COUNTRIES);

  const contact = {
    _source: trimVal(map.Source) || SOURCE_EN,
    title: title.en,
    titleOther: title.other,
    _rawTitle: title.raw,
    firstName: map['First Name'],
    lastName: map['Last Name'],
    jobTitle: job.en,
    jobOther: job.other,
    _rawJob: job.raw,
    department: '',
    industry: industry.en,
    industryOther: industry.other,
    _rawIndustry: industry.raw,
    institution: map['Company Name'],
    country: country.en,
    countryOther: country.other,
    _rawCountry: country.raw,
    email: map.Email,
    subject: '',
    subjectOther: '',
    phone: map['Phone (Optional)'],
  };

  return { contact, errors: validateImportContact(contact, 'employer', rowNum, OPTS) };
}

function parseSheet(workbook, sheetName, type, { onRow } = {}) {
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
      type === 'academic'
        ? `${ACADEMIC_FILL_HEADERS.slice(0, 4).join('、')}…`
        : `${EMPLOYER_FILL_HEADERS.slice(0, 4).join('、')}…`;
    return {
      contacts: [],
      errors: [
        {
          sheet: sheetName,
          row: 1,
          field: '表頭',
          message: `表頭欄位不符，請下載最新「單位填寫」匯入樣板（預期欄位：${expected}…）`,
        },
      ],
    };
  }

  const contacts = [];
  const errors = [];

  const dataRowCount = Math.max(0, matrix.length - 1);

  for (let i = 1; i < matrix.length; i++) {
    const rowNum = i + 1;
    const values = matrix[i];
    const map = rowToMap(format.canonicalHeaders, values);
    if (format.defaultSource && !trimVal(map.Source)) map.Source = SOURCE_EN;
    if (isEmptyDataRow(map) || isPlaceholderRow(values)) continue;

    const parsed = type === 'academic'
      ? parseAcademicRow(map, rowNum, { hasOtherCols: format.hasOtherCols })
      : parseEmployerRow(map, rowNum, { hasOtherCols: format.hasOtherCols });

    // 即使驗證失敗仍保留列資料，方便載入表單後用浮標修正
    contacts.push(parsed.contact);
    if (parsed.errors.length) errors.push(...parsed.errors);

    if (onRow) onRow(i, dataRowCount);
  }

  return { contacts, errors };
}

function checkDuplicateEmails(academic, employer) {
  const errors = [];
  const seen = new Map();

  const check = (list, sheet) => {
    list.forEach((c, idx) => {
      const email = normalizeEmail(c.email);
      const key = emailKey(email);
      if (!key) return;
      const row = idx + 2;
      if (seen.has(key)) {
        const prev = seen.get(key);
        errors.push({
          sheet,
          row,
          field: 'Email',
          message: `Email 與「${prev.sheet}」第 ${prev.row} 列重複（${email}）`,
        });
      } else {
        seen.set(key, { sheet, row });
      }
    });
  };

  check(academic, SHEET_ACADEMIC);
  check(employer, SHEET_EMPLOYER);
  return errors;
}

function checkPriorYearEmails(academic, employer) {
  const errors = [];
  const check = (list, sheet) => {
    list.forEach((c, idx) => {
      const msg = priorYearEmailError(c.email);
      if (!msg) return;
      errors.push({
        sheet,
        row: idx + 2,
        field: 'Email',
        message: msg,
      });
    });
  };
  check(academic, SHEET_ACADEMIC);
  check(employer, SHEET_EMPLOYER);
  return errors;
}

export async function parseImportWorkbook(buffer, options = {}) {
  const maxContacts = options.maxContacts ?? MAX_CONTACTS;
  const onProgress = options.onProgress;
  const tick = (label, percent) => {
    onProgress?.({ label, percent: Math.min(100, Math.max(0, Math.round(percent))) });
  };

  tick('讀取 Excel 檔案…', 3);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  tick('讀取完成', 10);

  const academicResult = parseSheet(workbook, SHEET_ACADEMIC, 'academic', {
    onRow: (cur, tot) => {
      if (!tot) return;
      tick(`驗證學術聯絡人（${cur}/${tot}）…`, 10 + (cur / tot) * 38);
    },
  });
  const employerResult = parseSheet(workbook, SHEET_EMPLOYER, 'employer', {
    onRow: (cur, tot) => {
      if (!tot) return;
      tick(`驗證雇主聯絡人（${cur}/${tot}）…`, 48 + (cur / tot) * 38);
    },
  });

  tick('檢查筆數與 Email…', 90);

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
  errors.push(...checkPriorYearEmails(academicResult.contacts, employerResult.contacts));

  const academic = academicResult.contacts.slice(0, maxContacts).map((c) => normalizeImportContact(c, 'academic'));
  const employer = employerResult.contacts.slice(0, maxContacts).map((c) => normalizeImportContact(c, 'employer'));

  tick('驗證完成', 100);

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
