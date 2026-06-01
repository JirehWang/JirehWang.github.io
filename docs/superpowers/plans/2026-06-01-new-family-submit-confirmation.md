# 新家人系統 - 提交確認提示實作計劃

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計劃。步驟使用複選框（`- [ ]`）語法來跟蹤進度。

**目標：** 在新家人管理系統表單提交前加入 `confirm(...)` 提示，阻斷未經確認的 Enter 鍵或誤觸提交行為。

**架構：** 攔截前端 `submit` 事件，於觸發 API 請求前呼叫 `confirm` 對話框。使用者確認後繼續送出，取消則阻斷非同步 API 請求。

**技術棧：** Vanilla JavaScript (ES6)

---

### 任務 1：修改前端表單提交邏輯

**文件：**
- 修改：[script.js](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/script.js) (修改範圍為 submit 監聽器)

- [ ] **步驟 1：修改 `script.js` 表單 submit 監聽器**
  尋找 [script.js](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/script.js) 中 `form.addEventListener('submit', ...)` 區塊，並於最頂部 `event.preventDefault();` 之後加入防誤觸確認邏輯。

  **修改目標代碼：**
  ```javascript
  form.addEventListener('submit', async event => {
    event.preventDefault();
    
    // 新增確認對話框
    if (!confirm('確認要送出此筆新家人資料嗎？')) {
      return;
    }
    
    setNotice(formNotice, '送出中...');
    submitBtn.disabled = true;
  ```

- [ ] **步驟 2：人工驗證邏輯與防誤觸**
  在瀏覽器中開啟新家人系統前端頁面，進行以下手動測試：
  1. 填寫「新家人姓名」等必填欄位。
  2. 在姓名輸入框中按下「Enter」鍵觸發提交。
  3. 預期：畫面應顯示「確認要送出此筆新家人資料嗎？」的對話框。
  4. 點擊「取消」：對話框關閉，表單內容保留，且**不發送** API 網路請求。
  5. 點擊「確認」：對話框關閉，按鈕停用，顯示「送出中...」，且資料成功寫入 Google Sheet。

- [ ] **步驟 3：Commit 變更**
  確認測試皆符合預期後，將代碼變更 commit 至 Git。

  運行命令：
  ```powershell
  git add apps/LKC_NewFamily/script.js
  git commit -m "feat(new-family): add submit confirmation dialog to prevent accidental creation"
  ```
