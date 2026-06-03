# 設計文檔：HTML config.js 引用路徑改為相對路徑

- **日期**：2026-06-03
- **狀態**：已批准 (Approved)
- **專案**：LKC 教會管理系統 (GitHub Pages)

## 1. 目的
將所有子應用程式 HTML 檔案中引用線上 `config.js` 的寫死路徑，改為本機相對路徑 `../../config.js`。
這能確保無論是在 GitHub Pages 線上環境，還是在本機 `localhost` 開發環境，均能載入最新版本的 `config.js`，進而觸發 PWA Service Worker 的註冊與更新邏輯。

## 2. 影響範圍
所有受影響的 HTML 檔案均位於 `D:/program/Github/LKC1958_June_1.github.io/apps/` 下的二級目錄。
因此，相對路徑 `../../config.js` 均可正確指向專案根目錄下的 `config.js`。

### 檔案清單：
1. `apps/LKC_Group/group.html`
2. `apps/LKC_Group/index.html`
3. `apps/LKC_Group/manage.html`
4. `apps/LKC_Group/stats.html`
5. `apps/LKC_MasterSchedule/calendar.html`
6. `apps/LKC_MasterSchedule/types.html`
7. `apps/LKC_MinistrySchedule/groupboard.html`
8. `apps/LKC_MinistrySchedule/index.html`
9. `apps/LKC_SundayserviceAttendance/index.html`
10. `apps/LKC_WhosCar/index.html`
11. `apps/LKC_worship/admin.html`
12. `apps/LKC_worship/index.html`
13. `apps/LKC_worship/worship_songs.html`

## 3. 修改對策
將：
```html
<script src="https://jirehwang.github.io/LKC1958_June_1.github.io/config.js"></script>
```
替換為：
```html
<script src="../../config.js"></script>
```

## 4. 驗證計畫
修改完成後，使用 `git diff` 驗證所有檔案皆修改正確。
並可用 grep 搜尋確認 `apps/` 目錄下不再含有該線上 URL 的引用。
