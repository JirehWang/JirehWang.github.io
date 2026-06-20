# ⛪ 教會主日出席系統

從 Google Apps Script 遷移至 GitHub Pages 的靜態前端。

## 📁 專案結構

```
church-attendance/
├── index.html          # 主頁面（導覽選單）
├── config.js           # ⚠️ 設定檔（填入你的 GAS URL）
├── js/
│   └── api.js          # GAS API 橋接層
└── pages/
    ├── attendance.html # 點名系統
    ├── members.html    # 會友名單管理
    ├── STATS.html      # 出席統計查詢
    └── Chart.html      # 趨勢分析圖表
```

## 🚀 部署步驟

### 1. 設定 GAS Web App URL

打開 `config.js`，將 `YOUR_SCRIPT_ID` 替換為你的 GAS 部署網址：

```javascript
window.GAS_CONFIG = {
  apiUrl: 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec',
};
```

**取得方式：**
- GAS 編輯器 → 右上角「部署」→「管理部署」→ 複製「網路應用程式」URL

### 2. 確認 GAS 後端有 doPost 入口

GAS 後端需要一個統一的 `doPost` 函式作為 API 路由入口，格式如下：

```javascript
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  const payload = body.payload;

  const result = (function() {
    switch (action) {
      case 'getAllMembers':      return getAllMembers();
      case 'updateMember':      return updateMember(payload[0], payload[1]);
      case 'deleteMember':      return deleteMember(payload);
      case 'addMember':         return addMember(payload);
      case 'getGroupConfig':    return getGroupConfig();
      case 'getSmartAttendanceList': return getSmartAttendanceList(payload[0], payload[1]);
      case 'syncClickToServer': return syncClickToServer(payload[0], payload[1], payload[2], payload[3]);
      case 'saveAttendance':    return saveAttendance(payload[0], payload[1], payload[2], payload[3], payload[4]);
      case 'revokeAttendance':  return revokeAttendance(payload[0], payload[1], payload[2]);
      case 'createAttendanceGroup': return createAttendanceGroup(payload[0], payload[1]);
      case 'updateDeviceMode':  return updateDeviceMode(payload[0], payload[1]);
      case 'getQuickSyncData':  return getQuickSyncData(payload[0], payload[1]);
      case 'getAttendanceStats': return getAttendanceStats(payload);
      case 'getCategoryChartData': return getCategoryChartData(payload[0], payload[1], payload[2]);
      case 'previewMemberCard': return previewMemberCard(payload);
      case 'generateMemberCard': return generateMemberCard(payload);
      default: throw new Error('Unknown action: ' + action);
    }
  })();

  return ContentService
    .createTextOutput(JSON.stringify({ data: result }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

### 3. 部署到 GitHub Pages

```bash
git init
git add .
git commit -m "init: migrate from GAS"
git remote add origin https://github.com/YOUR_USERNAME/church-attendance.git
git push -u origin main
```

在 GitHub Repository Settings → Pages → Source 選 `main` branch。

### 4. 本地測試

```bash
# Python 3
python -m http.server 8000
# 開啟 http://localhost:8000
```

## ⚠️ 注意事項

- `config.js` 中的 GAS URL 是明文，不要把敏感資料放在前端
- GAS Web App 部署時「存取權限」需設為「任何人」才能跨域呼叫
- 子頁面採用動態載入（fetch pages/*.html），需要透過 HTTP 伺服器訪問，不能直接雙擊開啟 index.html
