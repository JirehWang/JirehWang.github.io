# 教會系統 UX 優化設計規格書 - 本地驗證與快取機制

為了提升使用者操作系統時的流暢度，避免因為 Google Apps Script (GAS) 的網路延遲與 Cold Start 導致使用者誤以為「操作失敗」而重複執行，我們將進行以下前端 UX 優化設計。

## 1. 編輯模式本地解密驗證 (`LKC_MinistrySchedule`)

### 目前問題
進入編輯模式（點擊「進入編輯 (需輸入 ID)」）時，前端使用原生 `prompt()` 彈窗取得使用者輸入，並向 GAS 後端發送 `verifyPageId` API 請求。因為 GAS 沒有此安全操作的快取，每次驗證需花費 2-4 秒，且沒有明顯的視覺進度反馈，容易導致使用者重複點擊或誤判。

### 優化設計
1. **取消 `prompt()` 並替換為 Bootstrap Modal**：
   - 在 `LKC_MinistrySchedule\index.html` 新增 `#unlockVerifyModal`。
   - 使用者點擊解鎖時，彈出此 Modal，提供獨立的 Password 輸入框。
2. **移植解密演算法至前端**：
   - 在 `script.js` 實現本地 `decryptGroupCode(str)` 函式，利用金鑰 `LKC-Secure-2026` 進行 XOR 解密。
   - 當使用者在 Modal 中點擊「確認」或按下 Enter 時，前端直接將輸入內容與解密後的 `currentId` 比對。
   - 如果匹配，立刻更新 `isEditorUnlocked = true` 並調用 `sessionManager.setUnlocked(currentId)` 儲存至 Session 快取，關閉彈窗，實現 **0ms 延遲** 解鎖。
   - 如果不匹配，直接在 Modal 內部顯示紅字錯誤訊息 `❌ ID 輸入錯誤！`，不關閉彈窗，使用者能立刻修正。

---

## 2. 小組點名首頁與統計中心快取優化 (`LKC_Group`)

### 目前問題
- 每次手動進入小組時，系統都會彈出 `prompt()` 要求使用者輸入密碼，並發送 `verifyGroup` 請求到後端，產生 2-4 秒的卡頓。
- 統計中心 (`stats.html`) 的即時驗證使用 1000ms 的防抖延遲。使用者輸入完代碼後，若立即點擊「查詢」，會因為驗證尚未完成而跳出錯誤警告 `請先輸入正確的編號並等待識別`。

### 優化設計
1. **小組進入 Modal 替換與本地記住代碼 (localStorage)**：
   - 在 `LKC_Group\index.html` 新增 `#verifyModal` 替換原生 `prompt()`。
   - 驗證成功後，將小組名稱與加密代碼存入裝置快取：`localStorage.setItem('group_code_' + groupName, encryptedCode)`。
   - 下次使用者在首頁點擊該小組時，若偵測到快取存在，**直接讀取快取並瞬間跳轉至小組頁面，跳過輸入代碼與後端驗證步驟**。
   - 若後端代碼異動，進入小組頁面載入資料失敗時，會自動清除該快取並引導使用者重新輸入。
2. **縮短防抖時間與排隊查詢機制 (`stats.js`)**：
   - 將代碼輸入框的 Debounce 時間由 1000ms 縮短至 **400ms**，提供旋轉 Spinner 動態反饋（`⏳ 驗證中...`）。
   - 新增 `pendingVerificationPromise`。當使用者輸入完立即按「查詢」或 Enter 鍵時，若驗證仍在進行，系統會顯示 `正在驗證編號，請稍候...`，等驗證 Promise 完成後，自動執行數據查詢，避免報錯。

---

## 3. AI 排班表載入遮罩 (`LKC_MinistrySchedule`)

### 目前問題
點擊 AI 解析/排班時，後端需要執行 Gemini API 呼叫，通常需要 5-15 秒。此時畫面上僅右上角顯示「AI 運行中...」且按鈕未鎖定，使用者極易重複點擊。

### 優化設計
1. **新增全螢幕載入遮罩**：
   - 在 `LKC_MinistrySchedule\index.html` 新增與小組系統樣式一致的 `#loading-overlay`。
   - 當呼叫 `processAI()` 時，啟用此遮罩並顯示 `🤖 AI 運算中，請稍候...`，將整個頁面鎖定。
   - 停用 AI 彈窗內的「AI 解析」按鈕與 Textarea，運算結束後再解鎖，完全防止重複提交。
