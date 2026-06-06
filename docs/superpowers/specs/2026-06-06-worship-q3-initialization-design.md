# 設計規格：敬拜團 Q3 季度框架初始化與 PWA 協定快取修正

本規格說明如何實現當敬拜團本地排班表為空時，自動從「教會行事曆」的「講道資訊」事項載入初始日期框架，以利使用者進行排班。同時修復 PWA Service Worker 攔截 `chrome-extension` 請求導致的快取報錯與連線失敗問題。

## 1. 根本原因分析

### 1.1 第三季（Q3）無法初始化排班
當使用者在敬拜團排班後台（`admin.html`）選擇未排班的季度（如 `2026-Q3`）並點擊「讀取服事表」時，前端會調用後端 API `getSchedule`。
後端 [api_schedule.js](file:///d:/program/LKC/敬拜團/api_schedule.js#L30) 中的 `getMergedSchedule(year, quarter)` 僅從本地的 `服事表總表` 讀取資料。如果該季度完全沒有歷史排班存檔，本地查詢回傳空陣列 `[]`。
由於 `getMergedSchedule` 只會針對本地已存在排班列的日期去對照行事曆以補全講道明細，這導致在季度初始無資料時，後端無法主動從「教會行事曆」讀取該季度的聚會日期，直接回傳 `[]`，使用者無法取得任何日期框架來設定請假與生成班表。

### 1.2 PWA 快取錯誤與連線失敗 (TypeEror)
當瀏覽器註冊了 PWA Service Worker 之後，[service-worker.js](file:///d:/program/Github/LKC1958_June_1.github.io/service-worker.js#L32) 會攔截頁面內發出的所有 `GET` 請求。
如果使用者安裝了 Chrome 瀏覽器擴充功能（如安全工具、翻譯工具或調試工具），這些擴充功能運作時可能在頁面上下文中發送 `chrome-extension://` 協定的資源請求。
因為 Service Worker 攔截了這些請求，並在背景調用 `cache.put(event.request, ...)` 進行快取，但 Cache API 限制僅能存儲 `http:` 與 `https:` 協定的請求，進而拋出 `TypeError: Failed to execute 'put' on 'Cache': Request scheme 'chrome-extension' is unsupported`。此未捕獲的 Promise 拒絕可能導致 SW 快取邏輯崩潰，干擾其他靜態資源的正常載入，並在前端引發 `❌ 讀取失敗，請確認網路連線。` 警告。

---

## 2. 解決方案設計

### 2.1 後端 `getMergedSchedule` 自動對接行事曆初始化

當本地的排班表查詢結果 `localData` 長度為 0 時，後端將自動從教會行事曆拉取該季度的所有主日聚會作為排班的初始列。

- **影響檔案**：
  1. [敬拜團專案 api_schedule.js](file:///d:/program/LKC/敬拜團/api_schedule.js)
  2. [主日出席測試版 WorshipSchedule.js](file:///d:/program/LKC/主日出席_測試版/WorshipSchedule.js)

- **核心變更**：
  在 `getMergedSchedule(year, quarter)` 內部，如果 `localData` 為空，則：
  1. **計算季度日期區間**：
     * Q1: `01-01` 至 `03-31`
     * Q2: `04-01` 至 `06-30`
     * Q3: `07-01` 至 `09-30`
     * Q4: `10-01` 至 `12-31`
  2. **讀取行事曆事項**：
     * 調用 `_readCalendarSheet('事項')` 獲取行事曆所有事件。
     * 調用 `getCalendarLinkConfig()` 獲取已配置的講道子類型（如「台語」、「華語」等）。
     * 將事件篩選在計算出的季度區間內，且其 `typeId` 必須在講道子類型 ID 集合中。
  3. **讀取職位人員配置**：
     * 調用 `getPositions()` 獲取所有目前設定的位置（如主領、配唱1、吉他等）。
  4. **生成初始框架**：
     * 對於每個篩選出的行事曆事件，建立一筆排班資料列，將職位欄位依據 `isRequired === '是'` 預設為 `"【待定】"` 或 `""`。
     * 設定 `年度`、`季度`、`日期`、`聚會名稱`（設為事件標題）以及 `聚會類別`（設為子類型名稱）。
     * 將 `leaves` 陣列初始化為空 `[]`。
     * 排列日期並以此作為 `localData` 繼續執行後續與行事曆講道題目、牧師、經文的合併動作。

### 2.2 PWA Service Worker 協定過濾修正

- **影響檔案**：
  1. [service-worker.js](file:///d:/program/Github/LKC1958_June_1.github.io/service-worker.js)

- **核心變更**：
  在 `service-worker.js` 的 `fetch` 監聽器最上方，解析請求網址的 protocol。若非 `http:` 且非 `https:`，則直接退出（不呼叫 `event.respondWith`，將控制權交還瀏覽器原生處理）：
  ```javascript
  self.addEventListener('fetch', event => {
    const url = new URL(event.request.url);

    // 🛡️ 安全防護：只處理 http 和 https 協定，避免 chrome-extension 等協定觸發 Cache API 報錯
    if (url.protocol !== 'http:' && url.protocol !== 'https:') {
      return;
    }
    // ...
  ```

---

## 3. 規格自檢與驗證計劃

### 3.1 自檢項目
- **無占位符**：文件中無任何待定項，具體職位與日期區間均已明定。
- **後端一致性**：`getPositions` 中的必排設定直接對應初始化時的 `"【待定】"`，非必排對應 `""`，維持與前端「智慧產生」邏輯的一致。
- **無副作用**：PWA SW 修改僅過濾協定，不影響正常站內 `GET` 資源的快取。

### 3.2 驗證步驟
1. **本地執行測試函數**：
   在 GAS 後端執行 `testMergedScheduleQ3()`，檢查 Logger 輸出是否正確列出 2026-Q3 的主日聚會清單，且職位均已成功填入 `【待定】`。
2. **前端測試**：
   * 在排班後台選擇 2026-Q3，點擊「讀取服事表」，確認不再顯示連線失敗或無資料，而是呈現 2026-Q3 的聚會列。
   * 開啟瀏覽器 DevTools Console，確認無 `chrome-extension` 快取報錯。
