# Pool 現況頁（Google 登入）設定

1. 開啟 https://console.cloud.google.com/ → 建立（或選取）專案
2. API 和服務 → 啟用「Google Sheets API」「Google Drive API」
3. 憑證 → 建立憑證 → OAuth 用戶端 ID
   - 應用程式類型：網頁應用程式
   - 已授權的 JavaScript 來源：
     - http://localhost:4321
     - https://sixshoes.github.io
   - 已授權的重新導向 URI：可留空（使用 token client）
4. 複製「用戶端 ID」
5. 本機：寫入 code/.env
   PUBLIC_GOOGLE_CLIENT_ID=你的用戶端ID.apps.googleusercontent.com
   （可選）PUBLIC_GOOGLE_SHEET_ID=試算表網址 /d/ 與 /edit 之間那串
6. GitHub：Settings → Secrets and variables → Actions → Variables
   新增 PUBLIC_GOOGLE_CLIENT_ID（與可選 PUBLIC_GOOGLE_SHEET_ID）
7. 重新部署後開啟 /pool/ ，用 yitingchen@gm.fgu.edu.tw 登入
