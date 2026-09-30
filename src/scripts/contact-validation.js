/**
 * 聯絡人欄位驗證（表單與 Excel 匯入共用）
 */

import { PRIOR_YEAR_CONTACTS } from '../data/prior-year-contacts.js';

export const EMAIL_FORBIDDEN = /[\[\]{}();:,<>'#~=+"¬`!]/;

/** 常見共用／團隊信箱前綴（匯入時擋下） */
export const GENERIC_EMAIL_PREFIXES = [
  'team@',
  'info@',
  'admissions@',
  'noreply@',
  'no-reply@',
  'contact@',
  'admin@',
  'office@',
  'hr@',
  'support@',
  'enquiry@',
  'inquiry@',
];

export function trimVal(v) {
  return String(v ?? '').trim();
}

export function collapseSpaces(v) {
  return trimVal(v).replace(/\s+/g, ' ');
}

/** 全形 ASCII（含 ＠、．）轉半形，避免中文輸入法造成 Email 格式誤判 */
export function toHalfWidthAscii(v) {
  return String(v ?? '')
    .replace(/[\uFF01-\uFF5E]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0))
    .replace(/\u3000/g, ' ');
}

export function normalizeEmail(v) {
  return toHalfWidthAscii(trimVal(v)).replace(/\s+/g, '');
}

/** 比對用：只清空白／全形，大小寫視為不同帳號 */
export function emailKey(v) {
  return normalizeEmail(v);
}

/** @type {Map<string, typeof PRIOR_YEAR_CONTACTS>} */
const priorYearByEmail = new Map();
for (const rec of PRIOR_YEAR_CONTACTS) {
  const key = emailKey(rec.email);
  const list = priorYearByEmail.get(key) || [];
  list.push(rec);
  priorYearByEmail.set(key, list);
}

export function formatPriorYearRecord(rec) {
  const poolLabel = rec.pool === 'academic' ? '學術' : '雇主';
  const name = [rec.title, rec.firstName, rec.lastName].filter(Boolean).join(' ');
  return [poolLabel + '／' + rec.sheet, name, rec.role, rec.unit, rec.org].filter(Boolean).join('｜');
}

/** @returns {typeof PRIOR_YEAR_CONTACTS} */
export function findPriorYearMatches(email) {
  const key = emailKey(email);
  if (!key) return [];
  return priorYearByEmail.get(key) || [];
}

/** 寬鬆比對鍵：去尾點、不分大小寫（貼上檢查／常見貼錯用） */
export function emailLookupKeyLoose(v) {
  return normalizeEmail(v).replace(/\.+$/g, '').toLowerCase();
}

/** @type {Map<string, typeof PRIOR_YEAR_CONTACTS>} */
const priorYearByEmailLoose = new Map();
for (const rec of PRIOR_YEAR_CONTACTS) {
  const key = emailLookupKeyLoose(rec.email);
  if (!key) continue;
  const list = priorYearByEmailLoose.get(key) || [];
  list.push(rec);
  priorYearByEmailLoose.set(key, list);
}

/** @returns {typeof PRIOR_YEAR_CONTACTS} */
export function findPriorYearMatchesLoose(email) {
  const key = emailLookupKeyLoose(email);
  if (!key) return [];
  return priorYearByEmailLoose.get(key) || [];
}

/** @returns {string|null} */
export function priorYearEmailError(email) {
  const matches = findPriorYearMatches(email);
  if (!matches.length) return null;
  return '此 Email 與去年已提交名單重複：' + matches.map(formatPriorYearRecord).join('；');
}

/**
 * 從文字拆出信箱候選（換行、逗號、分號、空白、頓號皆可）
 * @param {string} text
 * @returns {string[]}
 */
export function splitEmailCandidates(text) {
  return String(text ?? '')
    .split(/[\n\r\t,;，；、|]+/u)
    .flatMap((part) => part.split(/\s+/u))
    .map((s) => trimVal(s))
    .filter(Boolean);
}

/**
 * @param {{ email?: string, unit?: string, submitter?: string, firstName?: string, lastName?: string, type?: string }} rec
 */
export function formatPoolRecord(rec) {
  const typeLabel = rec.type === 'employer' ? '雇主' : rec.type === 'academic' ? '學術' : '';
  const name = [rec.firstName, rec.lastName].filter(Boolean).join(' ');
  return [typeLabel, rec.unit, name, rec.submitter ? `提交人 ${rec.submitter}` : '']
    .filter(Boolean)
    .join('｜');
}

/**
 * 單筆／多筆信箱檢查（格式＋去年重複＋今年 Pool＋清單內重複）
 * @param {string} text
 * @param {{ findPool?: (email: string) => null | { email?: string, unit?: string, submitter?: string, firstName?: string, lastName?: string, type?: string } }} [options]
 * @returns {{ total: number, ok: number, rows: { raw: string, email: string, status: 'ok'|'invalid'|'prior'|'pool'|'dup', message: string, priorDetail?: string }[] }}
 */
export function inspectEmailList(text, options = {}) {
  const findPool = options.findPool;
  const candidates = splitEmailCandidates(text);
  /** @type {{ raw: string, email: string, status: 'ok'|'invalid'|'prior'|'pool'|'dup', message: string, priorDetail?: string }[]} */
  const rows = [];
  const seenLoose = new Map();

  for (const raw of candidates) {
    const email = normalizeEmail(raw);
    const formatErr = validateEmail(email, { checkGeneric: true });
    if (formatErr) {
      rows.push({ raw, email, status: 'invalid', message: formatErr });
      continue;
    }
    if (/\.$/.test(email)) {
      const prior = findPriorYearMatchesLoose(email);
      rows.push({
        raw,
        email,
        status: prior.length ? 'prior' : 'invalid',
        message: prior.length
          ? '結尾多了句點，且與去年名單重複'
          : 'Email 結尾多了句點，請檢查',
        priorDetail: prior.length ? prior.map(formatPriorYearRecord).join('；') : undefined,
      });
      continue;
    }

    const loose = emailLookupKeyLoose(email);
    const prior = findPriorYearMatchesLoose(email);
    if (prior.length) {
      rows.push({
        raw,
        email,
        status: 'prior',
        message: '與去年名單重複',
        priorDetail: prior.map(formatPriorYearRecord).join('；'),
      });
      continue;
    }

    const poolHit = typeof findPool === 'function' ? findPool(email) : null;
    if (poolHit) {
      rows.push({
        raw,
        email,
        status: 'pool',
        message: '今年已提交過（送出時會自動略過）',
        priorDetail: formatPoolRecord(poolHit),
      });
      continue;
    }

    if (seenLoose.has(loose)) {
      rows.push({
        raw,
        email,
        status: 'dup',
        message: `與清單第 ${seenLoose.get(loose)} 筆重複`,
      });
      continue;
    }
    seenLoose.set(loose, rows.length + 1);
    rows.push({ raw, email, status: 'ok', message: '可使用' });
  }

  return {
    total: rows.length,
    ok: rows.filter((r) => r.status === 'ok').length,
    rows,
  };
}

export function normalizePhone(v) {
  return trimVal(v).replace(/\s+/g, '');
}

/** @returns {string|null} */
export function validateEmail(email, { checkGeneric = true } = {}) {
  const e = normalizeEmail(email);
  if (!e) return 'Email 為必填';
  if (EMAIL_FORBIDDEN.test(e)) return 'Email 含有不允許的字元';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) return 'Email 格式不正確';
  if (checkGeneric && GENERIC_EMAIL_PREFIXES.some((p) => e.toLowerCase().startsWith(p))) {
    return '請勿使用 team@、info@ 等共用或團隊信箱';
  }
  return null;
}

/**
 * @param {import('../data/qs-options.js').BiOpt[]} list
 * @param {string} value
 */
export function matchOption(list, value, { pluralOther = false } = {}) {
  const raw = trimVal(value);
  if (!raw) return { en: '', other: '', invalid: false };

  const exact = list.find((item) => item.en === raw);
  if (exact) return { en: exact.en, other: '', invalid: false };

  const ci = list.find((item) => item.en.toLowerCase() === raw.toLowerCase());
  if (ci) return { en: ci.en, other: '', invalid: false };

  const byZh = list.find((item) => item.zh === raw);
  if (byZh) return { en: byZh.en, other: '', invalid: false };

  const venMatch = /^other\s*\(\s*ven\s*\)$/i.test(raw);
  if (venMatch) return { en: 'Ven', other: '', invalid: false };

  const otherRe = pluralOther ? /^others?\s*\((.+)\)$/i : /^other\s*\((.+)\)$/i;
  const otherMatch = raw.match(otherRe);
  if (otherMatch) return { en: 'Other', other: trimVal(otherMatch[1]), invalid: false };

  return { en: '', other: '', invalid: true, raw };
}

/**
 * 可選下拉或自行輸入：選單內→該選項；選單外→自動視為 Other（內容為輸入值）。
 */
export function resolveFlexibleOption(list, value, { pluralOther = false } = {}) {
  const raw = trimVal(value);
  if (!raw) return { en: '', other: '', raw: '' };

  if (raw === 'Other' || raw === 'Others') {
    return { en: 'Other', other: '', raw };
  }

  const parsed = matchOption(list, raw, { pluralOther });
  if (!parsed.invalid) {
    return { en: parsed.en, other: parsed.other || '', raw };
  }

  return { en: 'Other', other: raw, raw };
}

/** 表單／datalist 顯示用：Other 顯示自訂文字，其餘顯示英文選項 */
export function displayOptionInput(en, other, list) {
  if (en === 'Other') return other || '';
  if (!en) return '';
  const found = list.find((item) => item.en === en);
  return found ? found.en : en;
}

/**
 * @typedef {Object} ImportContact
 * @property {string} title
 * @property {string} titleOther
 * @property {string} firstName
 * @property {string} lastName
 * @property {string} jobTitle
 * @property {string} jobOther
 * @property {string} department
 * @property {string} industry
 * @property {string} industryOther
 * @property {string} institution
 * @property {string} country
 * @property {string} countryOther
 * @property {string} email
 * @property {string} subject
 * @property {string} subjectOther
 * @property {string} phone
 */

/**
 * @param {ImportContact} c
 * @param {'academic'|'employer'} type
 * @param {number} rowNum Excel 列號（從 2 起）
 * @param {import('../data/qs-options.js')} opts
 */
export function validateImportContact(c, type, rowNum, opts) {
  const {
    TITLES,
    ACADEMIC_JOB_TITLES,
    EMPLOYER_JOB_TITLES,
    INDUSTRIES,
    COUNTRIES,
    SUBJECTS,
    SOURCE_EN,
  } = opts;

  /** @type {{ sheet: string, row: number, field: string, message: string }[]} */
  const errors = [];
  const sheet = type === 'academic' ? '學術聯絡人' : '雇主聯絡人';
  const add = (field, message) => errors.push({ sheet, row: rowNum, field, message });

  const source = trimVal(c._source);
  if (source && source !== SOURCE_EN && source !== '佛光大學' && source !== 'Fo Guang University') {
    add('Source', `來源應為 Fo Guang University（目前：${source}）`);
  }

  if (!c.title) add('Title', '請填寫稱謂（Title）');
  if (c.title === 'Other' && !trimVal(c.titleOther)) {
    add('Title', '請填寫稱謂（可選下拉或自行輸入）');
  }

  if (!trimVal(c.firstName)) add('First Name', '名字為必填');
  if (!trimVal(c.lastName)) add('Last Name', '姓氏為必填');

  const jobList = type === 'academic' ? ACADEMIC_JOB_TITLES : EMPLOYER_JOB_TITLES;
  const jobField = type === 'academic' ? 'Job Title' : 'Position';
  if (!c.jobTitle) add(jobField, `請填寫${jobField}`);
  if (c.jobTitle === 'Other' && !trimVal(c.jobOther)) {
    add(jobField, `請填寫${jobField}（可選下拉或自行輸入）`);
  }

  if (type === 'academic') {
    if (!trimVal(c.department)) add('Department', '系所為必填');
    if (!c.subject) add('Subject', '請填寫學術領域（Subject）');
    if (c.subject === 'Other' && !trimVal(c.subjectOther)) {
      add('Subject', '請填寫學科領域（可選下拉或自行輸入）');
    }
  } else {
    if (!c.industry) add('Industry', '請填寫產業（Industry）');
    if (c.industry === 'Other' && !trimVal(c.industryOther)) {
      add('Industry', '請填寫產業（可選下拉或自行輸入）');
    }
  }

  const instField = type === 'academic' ? 'Institution' : 'Company Name';
  if (!trimVal(c.institution)) add(instField, `${instField} 為必填`);

  if (!c.country) add('Country or Territory', '請填寫國家或地區');
  if (c.country === 'Other' && !trimVal(c.countryOther)) {
    add('Country or Territory', '請填寫國家或地區（可選下拉或自行輸入）');
  }

  const emailErr = validateEmail(c.email);
  if (emailErr) add('Email', emailErr);

  const phone = normalizePhone(c.phone);
  if (phone && !/^[+0-9\-().]{6,20}$/.test(phone)) {
    add('Phone (Optional)', '電話格式不正確');
  }

  return errors;
}

/** @param {ImportContact[]} list @param {'academic'|'employer'} type */
export function normalizeImportContact(c, type) {
  return {
    type,
    title: c.title,
    titleOther: c.title === 'Other' ? collapseSpaces(c.titleOther) : '',
    firstName: collapseSpaces(c.firstName),
    lastName: collapseSpaces(c.lastName),
    jobTitle: c.jobTitle,
    jobOther: c.jobTitle === 'Other' ? collapseSpaces(c.jobOther) : '',
    department: collapseSpaces(c.department),
    industry: type === 'employer' ? c.industry : '',
    industryOther: type === 'employer' && c.industry === 'Other' ? collapseSpaces(c.industryOther) : '',
    institution: collapseSpaces(c.institution),
    country: c.country,
    countryOther: c.country === 'Other' ? collapseSpaces(c.countryOther) : '',
    email: normalizeEmail(c.email),
    subject: type === 'academic' ? c.subject : '',
    subjectOther: type === 'academic' && c.subject === 'Other' ? collapseSpaces(c.subjectOther) : '',
    phone: normalizePhone(c.phone),
  };
}
