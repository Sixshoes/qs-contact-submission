/**
 * 聯絡人欄位驗證（表單與 Excel 匯入共用）
 */

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

export function normalizeEmail(v) {
  return trimVal(v).replace(/\s+/g, '').toLowerCase();
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
  if (checkGeneric && GENERIC_EMAIL_PREFIXES.some((p) => e.startsWith(p))) {
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
    add('Title Other', '已選 Other，請在「Title Other」欄填寫說明');
  }

  if (!trimVal(c.firstName)) add('First Name', '名字為必填');
  if (!trimVal(c.lastName)) add('Last Name', '姓氏為必填');

  const jobList = type === 'academic' ? ACADEMIC_JOB_TITLES : EMPLOYER_JOB_TITLES;
  const jobField = type === 'academic' ? 'Job Title' : 'Position';
  if (!c.jobTitle) add(jobField, `請填寫${jobField}`);
  if (c.jobTitle === 'Other' && !trimVal(c.jobOther)) {
    add(`${jobField} Other`, `已選 Other，請在「${jobField} Other」欄填寫說明`);
  } else if (c.jobTitle && !jobList.some((item) => item.en === c.jobTitle) && c.jobTitle !== 'Other') {
    add(jobField, `「${c._rawJob || c.jobTitle}」不在選項清單內`);
  }

  if (type === 'academic') {
    if (!trimVal(c.department)) add('Department', '系所為必填');
    if (!c.subject) add('Subject', '請填寫學術領域（Subject）');
    if (c.subject === 'Other' && !trimVal(c.subjectOther)) {
      add('Subject Other', '已選 Other，請在「Subject Other」欄填寫說明');
    } else if (c.subject && !SUBJECTS.some((item) => item.en === c.subject) && c.subject !== 'Other') {
      add('Subject', `「${c._rawSubject || c.subject}」不在選項清單內`);
    }
  } else {
    if (!c.industry) add('Industry', '請填寫產業（Industry）');
    if (c.industry === 'Other' && !trimVal(c.industryOther)) {
      add('Industry Other', '已選 Other，請在「Industry Other」欄填寫說明');
    } else if (c.industry && !INDUSTRIES.some((item) => item.en === c.industry) && c.industry !== 'Other') {
      add('Industry', `「${c._rawIndustry || c.industry}」不在選項清單內`);
    }
  }

  const instField = type === 'academic' ? 'Institution' : 'Company Name';
  if (!trimVal(c.institution)) add(instField, `${instField} 為必填`);

  if (!c.country) add('Country or Territory', '請填寫國家或地區');
  if (c.country === 'Other' && !trimVal(c.countryOther)) {
    add('Country Other', '已選 Other，請在「Country Other」欄填寫說明');
  } else if (c.country && !COUNTRIES.some((item) => item.en === c.country) && c.country !== 'Other') {
    add('Country or Territory', `「${c._rawCountry || c.country}」不在選項清單內`);
  }

  if (c.title && !TITLES.some((item) => item.en === c.title) && c.title !== 'Other' && c.title !== 'Ven') {
    add('Title', `「${c._rawTitle || c.title}」不在選項清單內`);
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
