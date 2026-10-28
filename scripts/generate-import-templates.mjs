/**
 * 產生／同步「單位填寫／線上匯入」易讀樣板。
 *
 * 優先直接複製 source-templates 內已調好顏色／字型／下拉的原檔，
 * 以保留 Excel 原本的文字顏色與格式（勿用程式重建覆蓋）。
 *
 * 提交後下載給 QS Hub 的檔案仍由網站 excel-export 使用官方欄位（含 Source）。
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const curatedFile = path.resolve(__dirname, '../source-templates/QS_聯絡人匯入樣板_易讀版.xlsx');
const outDir = path.resolve(__dirname, '../public/templates');
const outFile = path.join(outDir, 'QS_聯絡人匯入樣板.xlsx');

function main() {
  if (!fs.existsSync(curatedFile)) {
    console.error('找不到已調色的易讀樣板：', curatedFile);
    console.error('請將調整好顏色的 Excel 放到 source-templates/QS_聯絡人匯入樣板_易讀版.xlsx');
    process.exit(1);
  }

  fs.mkdirSync(outDir, { recursive: true });
  fs.copyFileSync(curatedFile, outFile);
  const st = fs.statSync(outFile);
  console.log(`Copied curated fill template → ${outFile} (${st.size} bytes)`);
  console.log('（保留原檔文字顏色／字型／下拉驗證；勿改以程式重建）');
}

main();
