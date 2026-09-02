/** Excel 字型：中文標楷體、英文 Times New Roman */
export const FONT_ZH = '標楷體';
export const FONT_EN = 'Times New Roman';

/** 內文／填寫區字級（與原本 Calibri 11pt 相同） */
export const FONT_SIZE_BODY = 11;
/** 說明頁標題字級 */
export const FONT_SIZE_GUIDE_TITLE = 14;

const CJK_RE = /[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]/;

function scriptKind(ch) {
  return CJK_RE.test(ch) ? 'zh' : 'en';
}

function fontNameForScript(kind) {
  return kind === 'zh' ? FONT_ZH : FONT_EN;
}

/** @param {string} text */
function splitScriptRuns(text) {
  const str = String(text ?? '');
  if (!str) return [];

  /** @type {{ kind: 'zh' | 'en', text: string }[]} */
  const runs = [];
  let kind = scriptKind(str[0]);
  let buf = str[0];

  for (let i = 1; i < str.length; i++) {
    const ch = str[i];
    const next = scriptKind(ch);
    if (next === kind) {
      buf += ch;
    } else {
      runs.push({ kind, text: buf });
      kind = next;
      buf = ch;
    }
  }
  runs.push({ kind, text: buf });
  return runs;
}

/**
 * @param {import('exceljs').Cell} cell
 * @param {unknown} value
 * @param {import('exceljs').Font} [baseStyle]
 */
export function applyExcelCellValue(cell, value, baseStyle = {}) {
  const text = value == null ? '' : String(value);
  const style = { size: FONT_SIZE_BODY, color: { argb: 'FF0F172A' }, ...baseStyle };

  if (!text) {
    cell.value = '';
    cell.font = { ...style, name: FONT_ZH };
    return;
  }

  const runs = splitScriptRuns(text);
  if (runs.length <= 1) {
    cell.value = text;
    cell.font = { ...style, name: fontNameForScript(runs[0]?.kind ?? 'en') };
    return;
  }

  cell.value = {
    richText: runs.map((run) => ({
      font: { ...style, name: fontNameForScript(run.kind) },
      text: run.text,
    })),
  };
}

/** @param {import('exceljs').Font} [overrides] */
export function defaultBodyFont(overrides = {}) {
  return { name: FONT_ZH, size: FONT_SIZE_BODY, color: { argb: 'FF0F172A' }, ...overrides };
}

export function defaultExampleFont(overrides = {}) {
  return defaultBodyFont({ italic: true, color: { argb: 'FF64748B' }, ...overrides });
}

export function defaultHeaderFont(overrides = {}) {
  return { name: FONT_EN, size: FONT_SIZE_BODY, bold: true, color: { argb: 'FFFFFFFF' }, ...overrides };
}

export function defaultGuideTitleFont(overrides = {}) {
  return { name: FONT_ZH, size: FONT_SIZE_GUIDE_TITLE, bold: true, color: { argb: 'FF0F172A' }, ...overrides };
}

export function defaultGuideBodyFont(overrides = {}) {
  return { name: FONT_ZH, size: FONT_SIZE_BODY, color: { argb: 'FF334155' }, ...overrides };
}
