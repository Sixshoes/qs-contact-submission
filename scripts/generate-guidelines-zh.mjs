import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import {
  TITLES,
  ACADEMIC_JOB_TITLES,
  EMPLOYER_JOB_TITLES,
  INDUSTRIES,
  COUNTRIES,
  SUBJECTS,
} from '../src/data/qs-options.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.resolve(__dirname, '../source-docs');

const listItems = (items) =>
  items
    .map((item, i) => {
      const en = typeof item === 'string' ? item : item.en;
      const zh = typeof item === 'string' ? item : item.zh;
      return `${i + 1}. ${zh}（${en}）`;
    })
    .join('\n');

const md = `# QS 全球學術與雇主問卷指南（中文譯本）

> 原文：*Guidelines for QS Global Academic and Employer Survey*  
> 譯本說明：本文件為 QS 官方聯絡人名單填寫指南之繁體中文譯本，供佛光大學內部提報參考。**上傳 QS 之 CSV／Excel 欄位值仍請使用英文官方選項**（括號內英文），與 QS 範本一致。

---

## 前言

QS 全球學術與雇主問卷（QS Global Academic and Employer Surveys）之結果，用於 QS 排名中的**學術聲譽**與**雇主聲譽**指標。

為此，QS 邀請全球學術界與雇主界人士參與問卷並分享意見。

每年 QS 均要求各大學邀請學術與雇主聯絡人參與聲譽調查。參與流程詳見 QS 支援網站。  
同意邀請之 email 範本，以及透過聯絡人名單提供聯絡資訊時所用範本，亦可在 QS 支援網站取得。

---

## 1. 聯絡人名單（Contact List）

大學若要向 QS 提交聯絡人名單，須依下列步驟：

1.1. 自 QS 平台下載 CSV 格式之聯絡人名單範本。  
1.2. 在取得聯絡人同意後，依本指南第 2 節說明填寫範本。  
1.3. 將聯絡人名單以 **CSV UTF-8** 格式儲存，以便安全上傳至 Hub。  
1.4. 至 QS Hub 2.0 登入頁面，輸入帳密登入。  
1.5. 前往：**Institutions → Statistics → Contacts → Upload**  
1.6. 自下拉選單選擇 **Academic（學術）** 或 **Employer（雇主）**；亦可自按鈕下載範本與上傳說明。  
1.7. 點選上傳並選擇檔案，勾選同意（consent）後儲存提交。

---

## 2. 聯絡人名單填寫說明

名單中每位聯絡人之各欄位，須依下列規範填寫：

### 2.1. Source（來源）

聯絡人之「來源」為向 QS 提交提名名單之機構。  
請以**機構正式全名**填寫，**不可使用縮寫或簡稱**。  
（佛光大學填寫：**Fo Guang University**）

### 2.2. Title（稱謂）

聯絡人希望搭配姓名使用之正式稱謂。

**可選稱謂：**

${listItems(TITLES)}

若選 **Other（其他）**，請於括號內說明，例如：\`Other (Ven)\`。

### 2.3. First and Last Name（名字與姓氏）

名字（First Name）與姓氏（Last Name）須**分欄填寫**，使用正常文字格式，**前後不可有空格**。

### 2.4. Job Title / Position（職稱）

為求一致，QS 對學術與雇主聯絡人各訂有職稱清單。請盡量從清單中選擇。  
若聯絡人職務不在清單內，可選 **Other / Others（其他）**，並於括號內說明，例如：\`Others (xyz)\`。

#### 2.4.1 學術職稱（Academic Designations）

${listItems(ACADEMIC_JOB_TITLES)}

#### 2.4.2 雇主職稱（Employer Designations）

${listItems(EMPLOYER_JOB_TITLES)}

### 2.5. Department / Industry（系所／產業）

**學術聯絡人：** 各校系所名稱不一，請填寫**完整系所名稱**。  

**雇主聯絡人：** QS 採固定產業清單，請自下列選項選擇。若不在清單內，選 **Other（其他）** 並於括號內說明。

**產業清單（List of Industries）：**

${listItems(INDUSTRIES)}

### 2.6. Institution / Company Name（機構／公司名稱）

- **學術聯絡人：** 填寫聯絡人所屬**機構**（Institution）之正式名稱。  
- **雇主聯絡人：** 填寫聯絡人所屬**公司**（Company Name）之正式名稱。  

請與正式文件上之寫法一致。

### 2.7. Country or Territory（國家或地區）

QS 使用固定之國家／地區清單分類機構，填寫時請自清單選擇，**請勿填寫街道、城市或州／省地址**。

**國家／地區清單（填寫時請用英文括號內名稱）：**

${listItems(COUNTRIES)}

### 2.8. Email（電子信箱）

建議學術與雇主聯絡人均提供**個人公務／官方信箱**。請注意：

i. **請勿使用**通用或團隊信箱，例如 \`team@xyz.com\`、\`admissions@xyz.com\`、\`info@xyz.com\`。  
ii. 每個儲存格**僅能填一組** email；若同一人有多組信箱，請**複製整列**另填。  
iii. 若多人填寫**相同 email**，QS **僅會寄送一次**問卷邀請，請使用**不同信箱**。  
iv. email 不可含下列字元：\`[ ] { } ( ) ; : , < > ' # ~ = + ! " " ¬ \\\`\`  
v. email **不可含空白**（前、後或中間皆不可）。

### 2.9. Subject（學術領域，僅學術聯絡人）

為求一致，QS 對學術聯絡人訂有學科清單。請自清單選擇；若不在清單內，選 **Others（其他）** 並於括號內說明，例如：\`Others (xyz)\`。

**學科清單（List of subjects）：**

${listItems(SUBJECTS)}

### 2.10. Phone（電話，選填）

學術或雇主聯絡人之電話為**選填**；若涉隱私可不留。  
若填寫，請含**國際碼**，例如：\`+886...\`。

---

## 3. 問卷邀請（Survey Invitation）

QS 於每年 **2 月至 3 月**寄出學術與雇主聲譽問卷邀請。每位聯絡人最多收到 **3 封 email**（1 封邀請函 + 2 封提醒）。  
寄件者為：**rankings@qs.com**

為確保邀請信順利送達，建議：

i. 將 **rankings@qs.com** 加入安全寄件者／白名單，避免退信或進入垃圾郵件。  
ii. 若信箱使用中央伺服器，請透過資訊單位確認該 email 已白名單，未被機構封鎖。  
iii. 問卷邀請期間（2–3 月），請**每週檢查**垃圾郵件匣，以免漏信。

---

*譯本完成日期：2026 年 8 月｜佛光大學永續辦公室內部參考用*
`;

const outPath = path.join(outDir, 'Academic and Employer Survey Guidelines（中文譯本）.md');
fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(outPath, md, 'utf8');
console.log('Written:', outPath);
