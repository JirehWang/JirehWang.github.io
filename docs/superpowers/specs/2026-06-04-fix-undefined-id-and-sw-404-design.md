# 設計規格：修復事工系統 undefined ID 與 PWA ServiceWorker 404 報錯

本規格說明如何解決事工系統載入時拋出 `找不到 ID 對應的分頁：undefined` 錯誤，以及線上 PWA ServiceWorker 404 報錯問題。

## 1. 根本原因分析

### 1.1 `undefined` ID 錯誤
在小組點名系統中，若使用者透過專屬連結登入（URL 帶有 `?id=enc_xxxx` 參數），前端 [index.js](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_Group/index.js) 會呼叫後端 API `findGroupByCode`。
然而，後端 [Core.js](file:///d:/program/LKC/小組點名/Core.js#L130) 及 [GroupCore.js](file:///d:/program/LKC/主日出席_測試版/GroupCore.js#L137) 的 `findGroupByCode` 函數在驗證成功後，只回傳了 `success`、`groupName` 與 `isAdmin`，**漏掉了 `encryptedCode` 欄位**。
這導致前端在跳轉到事工排班系統時，將其編碼為 `"undefined"` 字串（即 `?id=undefined`）。當事工系統嘗試載入配置並發送 `getPageConfig` 請求時，後端即拋出 `找不到 ID 對應的分頁：undefined` 的 APIError。

此外，正式版的「事工管理」與「小組點名」後端 GAS 專案尚未執行 `clasp push`，這導致最新版本的解密代碼尚未部署，即使前端傳入正確的加密代碼也可能無法在正式環境中正常解析。

### 1.2 PWA ServiceWorker 404 報錯
本地的 `service-worker.js` 與 `manifest.json` 在專案中已經存在，但處於 Git `untracked`（未追蹤）狀態，因此 GitHub Pages 上並沒有部署這兩個檔案，導致 `config.js` 自動嘗試註冊 ServiceWorker 時收到 404 響應。

---

## 2. 解決方案設計

### 2.1 後端 `findGroupByCode` 修改
在後端的 `findGroupByCode` 比對代碼成功時，必須將加密後的代碼回傳給前端。

- **修改檔案**：
  1. [小組點名正式版 Core.js](file:///d:/program/LKC/小組點名/Core.js#L130)
  2. [主日出席測試版 GroupCore.js](file:///d:/program/LKC/主日出席_測試版/GroupCore.js#L137)

- **具體修改**：
  在 `rowCode === decryptedCode` 條件成立時，回傳物件新增 `encryptedCode`：
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

### 2.2 後端 GAS 部署
修改完成後，使用 `clasp` 將以下兩個正式版後端專案部署至雲端：
1. **事工管理** (`d:/program/LKC/事工管理/`)
2. **小組點名** (`d:/program/LKC/小組點名/`)

### 2.3 PWA 資源追蹤與發布
將 `service-worker.js` 與 `manifest.json` 新增至 Git，並執行 commit 與 push，使其在 `https://jirehwang.github.io/LKC1958_June_1.github.io/` 生效。

---

## 3. 規格自檢與驗證計劃

### 3.1 自檢清單
- **無占位符**：文件中沒有任何待定或 TODO 項。
- **一致性**：測試版與正式版後端的修改完全對齊，加密密鑰與前綴維持 `LKC-Secure-2026` / `enc_` 不變。
- **邊界清晰**：加解密仍在後端完成，前端僅接收、保存與透傳加密後的 code。

### 3.2 驗證步驟
1. 部署測試版/正式版後端，確認使用專屬連結進入小組後，瀏覽器網址的 `code` 參數不再是 `undefined`，而是正確的加密字串 `enc_xxxx`。
2. 點擊「前往事工排班」按鈕，確認事工排班頁面的 URL 參數為正確的加密 ID，且頁面能成功加載、無 `找不到 ID` 報錯。
3. 部署 GitHub Pages 後，清除瀏覽器快取，檢查 Console，確認 PWA ServiceWorker 成功註冊、不再出現 404 警告。
