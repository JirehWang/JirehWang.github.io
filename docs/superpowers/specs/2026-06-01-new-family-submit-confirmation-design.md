# 新家人系統 - 提交確認提示設計規格書 (New Family Form Submit Confirmation Spec)

本規格書設計在「新家人管理系統」前端表單提交時，加入瀏覽器內建確認視窗，以避免使用者因誤按 Enter 鍵或操作失誤導致在未完成填寫時自動建立資料庫記錄。

## 1. 變更範圍

本變更僅影響前端網頁元件，不涉及 Google Apps Script 後端或試算表結構。

*   **修改檔案**：
    *   [script.js](file:///d:/program/Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/script.js) (前端主要邏輯)

## 2. 詳細設計

### 2.1 表單提交事件攔截
在使用者點擊「送出表單」或於輸入框按下 Enter 鍵觸發表單提交事件時，在呼叫後端 API 之前彈出瀏覽器內建的確認對話方塊。

#### 邏輯流程：
1. 攔截表單的 `submit` 事件。
2. 呼叫 `confirm('確認要送出此筆新家人資料嗎？')`。
3. 如果使用者選擇「取消」：
   - 阻斷後續 API 呼叫。
   - 保持表單狀態不變（按鈕維持啟用，輸入內容保留）。
4. 如果使用者選擇「確認」：
   - 執行原先的 `preventDefault()`。
   - 將提交按鈕設為停用 (`disabled = true`)。
   - 顯示「送出中...」提示。
   - 透過 `callApi('submitNewFamily', ...)` 向 GAS 後端傳遞資料。

### 2.2 程式碼變更預覽

```javascript
// Github/LKC1958_June_1.github.io/apps/LKC_NewFamily/script.js
form.addEventListener('submit', async event => {
  event.preventDefault();
  
  // 新增：防誤觸確認對話框
  if (!confirm('確認要送出此筆新家人資料嗎？')) {
    return;
  }
  
  setNotice(formNotice, '送出中...');
  submitBtn.disabled = true;

  try {
    const result = await callApi('submitNewFamily', Object.fromEntries(new FormData(form).entries()));
    setNotice(formNotice, `${result.message}，表單號：${result.formNumber}`, 'success');
    form.reset();
    dateField.valueAsDate = new Date();
  } catch (error) {
    setNotice(formNotice, error.message || String(error), 'error');
  } finally {
    submitBtn.disabled = false;
  }
});
```

## 3. 驗證計劃 (Verification Plan)

為確保此防誤觸邏輯正確生效且不影響正常功能，將進行以下手動驗證步驟：

1. **正常送出流程驗證**：
   - 填寫表單後點擊「送出表單」按鈕。
   - 預期：彈出「確認要送出此筆新家人資料嗎？」提示。
   - 點擊「確認」後：按鈕停用，畫面顯示「送出中...」，成功送出後表單重設，資料正常寫入試算表。

2. **取消送出流程驗證**：
   - 填寫表單後點擊「送出表單」按鈕。
   - 預期：彈出「確認要送出此筆新家人資料嗎？」提示。
   - 點擊「取消」後：對話框關閉，表單維持原樣，內容不被清除，按鈕維持啟用狀態，**且後端資料庫（試算表）中沒有新增任何記錄**。

3. **Enter 鍵防誤觸驗證**：
   - 填寫「新家人姓名」後，在姓名輸入框中按下「Enter」鍵。
   - 預期：彈出「確認要送出此筆新家人資料嗎？」提示。
   - 點擊「取消」後：對話框關閉，表單內容保留，且不發送 API 請求。
