# 同工服事狀態管理系統

這個頁面是會友、出席、小組、事工與敬拜服事資料的只讀監控介面。API 由根目錄 `config.js` 將 action 加上 `memberStatus_` 前綴，再交給合併主 GAS 的 `MemberStatusCore.js`。

## 載入與互動

- 初次載入只呼叫 `getMembers`，先呈現摘要、篩選器、事工參與分布與會友列表。
- 不會在初次載入時自動呼叫第一位會友的 `getProfile`；使用者點選會友後才讀取詳細資料，避免多一次阻塞首屏的 GAS 往返。
- 事工參與分布預設顯示排序前 24 位。超過 24 位時提供「顯示完整 N 人」，點擊後在目前篩選條件下呈現完整名單，亦可收起。
- 重新整理按鈕先呼叫 `refreshCaches`，再重新讀取 `getMembers`。

## 主要檔案

- `index.html`：頁面結構與完整名單切換按鈕。
- `script.js`：載入、篩選、事工參與排序/展開、會友詳細資料讀取。
- `style.css`：桌面與行動版樣式。

前端行為測試位於 `tests/member-status-frontend.test.js`；後端聚合測試位於 `tests/member-status-core.test.js`。
