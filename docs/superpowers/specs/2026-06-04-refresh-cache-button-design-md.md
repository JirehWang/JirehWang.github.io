# 系統快取重新整理按鈕設計規格書 (Refresh Cache Button Design Spec)

- **日期**：2026-06-04
- **狀態**：已批准 (Approved)
- **目標**：在門戶首頁 (`index.html`) 新增一個右下角懸浮按鈕 (FAB)，點擊後能夠一鍵清除瀏覽器的 Service Worker 快取，並重新載入頁面，確保會友與同工能即時獲取最新網頁資源。

---

## 1. 視覺與互動設計 (UI/UX)

### 1.1 懸浮按鈕 (Floating Action Button, FAB)
* **位置**：固定於畫面右下角 (`position: fixed; bottom: 24px; right: 24px; z-index: 1000;`)。
* **形狀**：圓形，直徑 50px。
* **背景色**：使用品牌色漸層，從 `#006030` (林口教會深綠色) 漸變至 `#30759f` (藍色)，與頁首漸層呼應。
* **圖示**：白色 SVG 重新整理 (Refresh) 圖示。
* **陰影**：`box-shadow: 0 4px 14px rgba(0, 0, 0, 0.2);`

### 1.2 微互動與動畫
* **Hover (懸停)**：
  - 按鈕微幅放大 `transform: scale(1.08)`。
  - 陰影變深 `box-shadow: 0 6px 20px rgba(0, 0, 0, 0.25)`。
  - 圖示順時針旋轉 180 度。
* **Active (點擊中)**：
  - 按鈕縮小 `transform: scale(0.95)`。
* **Spinning (清除中)**：
  - 圖示持續順時針 360 度旋轉，表示背景正在處理中。
* **Tooltip (提示氣泡)**：
  - 滑鼠移入時，在按鈕上方顯示簡約氣泡：「清除快取並強制更新」。

### 1.3 Toast 提示框
* 點擊按鈕後，在畫面下方中央顯示半透明深色 Toast：「🔄 正在清除系統快取並重新載入...」。
* 樣式使用現代玻璃擬態 (Glassmorphism)，背景為 `rgba(0, 0, 0, 0.75)`，文字為白色，邊角半徑 8px。

---

## 2. 功能行為與邏輯 (Behavior & Logic)

當按鈕被點擊時，將觸發以下非同步流程：

1. **鎖定按鈕**：
   - 將按鈕設為禁用，避免使用者重複點擊。
   - 為圖示加入 `.spinning` class 以顯示旋轉動畫。
2. **顯示 Toast**：
   - 插入 Toast HTML 到 DOM 中，展示清除快取狀態。
3. **清除快取與註銷**：
   - 呼叫 `caches.keys()` 遍歷所有快取儲存，並使用 `caches.delete(name)` 進行刪除。
   - 呼叫 `navigator.serviceWorker.getRegistrations()` 獲取所有已註冊的 Service Worker，並執行 `registration.unregister()`。
   - 動態載入 `./firebase/firebase-config.js` 的 `rtdb` 實例，並呼叫 Firebase Database SDK 的 `remove(ref(rtdb, 'cache'))` 直接刪除雲端快取節點。
4. **硬性重新整理**：
   - 延遲 1000 毫秒（提供使用者動畫感知時間），隨後呼叫 `window.location.reload()` 完成更新。

---

## 3. 變更點 (Proposed Changes)

### 3.1 [MODIFY] [index.html](file:///d:/program/Github/LKC1958_June_1.github.io/index.html)
* 在 `<body>` 結束標籤前，新增懸浮按鈕的 HTML 結構、CSS 樣式以及 JS 快取清理與 Firebase 快取重置邏輯。

---

## 4. 驗證計劃 (Verification Plan)

### 手動測試 (Manual Verification)
1. 在瀏覽器中載入網頁，確認右下角是否出現懸浮按鈕。
2. 懸停在按鈕上，驗證按鈕是否微幅放大、陰影加深，且 Refresh 圖示旋轉。
3. 點擊按鈕，確認是否彈出 Toast 提示，且圖示開始持續旋轉。
4. 確認頁面是否在 1 秒後成功重新整理。
5. 檢查瀏覽器的開發者工具 (Developer Tools -> Application -> Cache Storage)，確認快取已被成功清除。
6. 檢查 Firebase Realtime Database 控制台，確認點擊按鈕後 `cache` 節點被成功刪除。
