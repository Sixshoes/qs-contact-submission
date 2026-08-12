import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { marked } from 'marked';
import { chromium } from 'playwright';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, '..');
const sourceDir = path.join(rootDir, 'source-docs');
const publicDocsDir = path.join(rootDir, 'public', 'docs');

const enSourcePdf = path.join(sourceDir, 'Academic and Employer Survey Guidelines.pdf');
const zhSourceMd = path.join(sourceDir, 'Academic and Employer Survey Guidelines（中文譯本）.md');
const enOutPdf = path.join(publicDocsDir, 'qs-guidelines-en.pdf');
const zhOutPdf = path.join(publicDocsDir, 'qs-guidelines-zh.pdf');

marked.setOptions({ gfm: true, breaks: false });

function buildPrintHtml(title, bodyHtml) {
  return `<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8" />
  <title>${title}</title>
  <style>
    @page { size: A4; margin: 18mm 16mm; }
    body {
      font-family: "Microsoft JhengHei", "PingFang TC", "Noto Sans TC", sans-serif;
      font-size: 10.5pt;
      line-height: 1.55;
      color: #1e293b;
    }
    h1 { font-size: 20pt; margin: 0 0 12pt; color: #1b4332; }
    h2 { font-size: 14pt; margin: 18pt 0 8pt; color: #2d6a4f; page-break-after: avoid; }
    h3 { font-size: 12pt; margin: 14pt 0 6pt; color: #40916c; page-break-after: avoid; }
    h4 { font-size: 11pt; margin: 12pt 0 4pt; page-break-after: avoid; }
    p, li { margin: 0 0 6pt; }
    ul, ol { margin: 0 0 8pt 18pt; padding: 0; }
    blockquote {
      margin: 0 0 12pt;
      padding: 8pt 12pt;
      border-left: 3pt solid #52b788;
      background: #f8fafc;
      color: #475569;
    }
    code {
      font-family: Consolas, "Courier New", monospace;
      font-size: 9.5pt;
      background: #f1f5f9;
      padding: 0 3pt;
      border-radius: 3pt;
    }
    hr { border: none; border-top: 1px solid #cbd5e1; margin: 14pt 0; }
    strong { color: #0f172a; }
  </style>
</head>
<body>${bodyHtml}</body>
</html>`;
}

async function markdownToPdf(mdPath, pdfPath, title) {
  const markdown = fs.readFileSync(mdPath, 'utf8');
  const bodyHtml = marked.parse(markdown);
  const html = buildPrintHtml(title, bodyHtml);

  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle' });
    await page.pdf({
      path: pdfPath,
      format: 'A4',
      printBackground: true,
      margin: { top: '18mm', right: '16mm', bottom: '18mm', left: '16mm' },
    });
  } finally {
    await browser.close();
  }
}

fs.mkdirSync(publicDocsDir, { recursive: true });

if (!fs.existsSync(enSourcePdf)) {
  console.error('Missing English PDF:', enSourcePdf);
  process.exit(1);
}
fs.copyFileSync(enSourcePdf, enOutPdf);
console.log('Copied:', enOutPdf);

if (!fs.existsSync(zhSourceMd)) {
  console.error('Missing Chinese markdown:', zhSourceMd);
  console.error('Run: node scripts/generate-guidelines-zh.mjs');
  process.exit(1);
}

await markdownToPdf(
  zhSourceMd,
  zhOutPdf,
  'QS 全球學術與雇主問卷指南（中文譯本）',
);
console.log('Generated:', zhOutPdf);
