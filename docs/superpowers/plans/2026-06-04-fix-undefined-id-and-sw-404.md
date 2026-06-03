# 修復事工系統 undefined ID 與 SW 404 實現計劃

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計劃。步驟使用複選框（`- [ ]`）語法來跟踪進度。

**目標：** 修正後端 `findGroupByCode` 以提供 `encryptedCode`，使前端跳轉時攜帶正確的加密代碼，並推送正式版 GAS 及上傳 PWA ServiceWorker 相關檔案，解決 `undefined` ID 和 SW 404 報錯。

**建築：** 
1. 修正小組點名後端的 `findGroupByCode` 函數，回傳包含 `encryptedCode` 屬性的物件。
2. 使用 `clasp` 部署更新到正式版「事工管理」與「小組點名」後端。
3. 將本地 `service-worker.js` 與 `manifest.json` 納入 Git 追蹤，commit 並 push 到 GitHub 遠端。

**技術棧：** JavaScript (Google Apps Script), git, clasp

---

### 任務 1：修正小組點名正式版後端 `findGroupByCode`

**文件：**
- 修改：`d:/program/LKC/小組點名/Core.js`

- [ ] **步驟 1：修改代碼加入 `encryptedCode` 回傳**
  在 `d:/program/LKC/小組點名/Core.js` 的 `findGroupByCode` 函數內（約第 150-158 行），在比對 `rowCode === decryptedCode` 成功回傳的結果中，加入 `encryptedCode: encryptGroupCode(decryptedCode)`。
  修改後的程式碼片段：
  ```javascript
  if (rowCode === decryptedCode) {
    // 找到了！回傳小組名稱 (A 欄)
    return { 
      success: true, 
      groupName: String(data[i][0]).trim(), 
      isAdmin: false,
      encryptedCode: encryptGroupCode(decryptedCode)
    };
  }
  ```

- [ ] **步驟 2：人工代碼核對**
  核對 `d:/program/LKC/小組點名/Core.js` 中是否正確定義了 `encryptGroupCode` 與 `decryptedCode`，確保其作用域與拼字無誤。

---

### 任務 2：修正主日出席測試版後端 `findGroupByCode`

**文件：**
- 修改：`d:/program/LKC/主日出席_測試版/GroupCore.js`

- [ ] **步驟 1：修改代碼加入 `encryptedCode` 回傳**
  在 `d:/program/LKC/主日出席_測試版/GroupCore.js` 的 `findGroupByCode` 函數內（約第 151-158 行），在比對 `rowCode === decryptedCode` 成功回傳的結果中，加入 `encryptedCode: encryptGroupCode(decryptedCode)`。
  修改後的程式碼片段：
  ```javascript
  if (rowCode === decryptedCode) {
    return {
      success: true,
      groupName: String(data[i][0]).trim(),
      isAdmin: false,
      encryptedCode: encryptGroupCode(decryptedCode)
    };
  }
  ```

- [ ] **步驟 2：人工代碼核對**
  核對 `d:/program/LKC/主日出席_測試版/GroupCore.js` 中是否正確定義了 `encryptGroupCode` 與 `decryptedCode`，確保無拼字錯誤。

- [ ] **步驟 3：部署測試版 GAS**
  在 `d:/program/LKC/主日出席_測試版/` 目錄執行 `npx @google/clasp push --force`。

---

### 任務 3：部署正式版事工管理與小組點名 GAS

**文件：**
- 修改：無（僅執行 clasp 部署）

- [ ] **步驟 1：部署正式版事工管理 GAS**
  進入 `d:/program/LKC/事工管理/` 目錄，執行部署命令：
  ```powershell
  npx @google/clasp push --force
  ```
  預期輸出包含：`Pushed 3 files.` 或是類似成功的提示，代表正式版事工管理最新的加密解密邏輯已部署上線。

- [ ] **步驟 2：部署正式版小組點名 GAS**
  進入 `d:/program/LKC/小組點名/` 目錄，執行部署命令：
  ```powershell
  npx @google/clasp push --force
  ```
  預期輸出包含：`Pushed 5 files.` 或是類似成功的提示，代表正式版小組點名最新的回傳加密邏輯已部署上線。

---

### 任務 4：發布 PWA ServiceWorker 檔案並推送前端變更

**文件：**
- 新增/追蹤：`d:/program/Github/LKC1958_June_1.github.io/service-worker.js`
- 新增/追蹤：`d:/program/Github/LKC1958_June_1.github.io/manifest.json`

- [ ] **步驟 1：將 SW 與 manifest 納入 Git 追蹤**
  在 `d:/program/Github/LKC1958_June_1.github.io` 目錄下，執行命令：
  ```powershell
  git add service-worker.js manifest.json
  ```

- [ ] **步驟 2：提交變更並推送至 GitHub**
  執行 commit 與 push 動作：
  ```powershell
  git commit -m "feat: add service-worker.js and manifest.json to fix PWA 404"; git push origin main
  ```
  預期執行成功且順利推送到遠端伺服器。

---

### 任務 5：線上驗證

- [ ] **步驟 1：清除瀏覽器快取並驗證專屬連結跳轉**
  使用瀏覽器開啟專屬網址，如：`https://jirehwang.github.io/LKC1958_June_1.github.io/apps/LKC_Group/?id=enc_xxxxx`（使用實際小組加密金鑰或測試金鑰），確認成功進入該小組，且轉跳後的 URL 中 `code` 不再是 `"undefined"`，而是正確的加密字串。

- [ ] **步驟 2：驗證事工排班跳轉**
  在小組點名頁面點擊「前往事工排班」，驗證其開啟之新視窗中 URL 的 `id` 為正確的加密代碼，且頁面能成功呈現資料，不再出現 API 錯誤。

- [ ] **步驟 3：驗證 ServiceWorker 404 是否消失**
  重新載入頁面並開啟 Console，檢查有無 ServiceWorker 404 報錯。預期應為 `ServiceWorker 註冊成功`。
