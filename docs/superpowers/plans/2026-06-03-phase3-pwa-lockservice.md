# 第三階段 (LockService 排隊鎖定與 PWA 離線快取) 實現計劃

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計劃。步驟使用複選框（`- [ ]`）語法來跟踪進度。

**目標：** 於測試版主 GAS 系統導入 `LockService` 排隊機制以根除多小組同時寫入時的小組與身分字串覆蓋衝突；同時在前端 Pages 導入 PWA 離線快取機制，解決主日現場網路擁塞導致網頁載入緩慢的問題。

**建築：**
1. **後台鎖定 (GAS LockService)**：於 `MemberDB.js` 與 `GroupCore.js` 的寫入進入點導入 GAS 內建 Script Lock，寫入前重新讀取 Sheet 最新狀態並排隊操作。
2. **PWA 自動註冊與離線快取**：於前端根目錄建立 `manifest.json` 與 `service-worker.js`，並於 `config.js` 的初始化邏輯中自動偵測 Scope 並註冊 SW，不對任何動態 POST API 快取。

**技術棧：** JavaScript, Service Worker API, Google Apps Script LockService API, PWA Manifest.

---

## 🛠️ 檔案結構變動清單

* **後端主 GAS 專案 (`D:\program\LKC\主日出席_測試版`)**
  * `MemberDB.js` (修改：`addMember`、`updateMember`、`deleteMember` 導入 LockService 鎖定)
  * `GroupCore.js` (修改：`submitAttendance`、`updateMemberList`、`updateGroupInfo` 導入 LockService 鎖定)
* **前端 GitHub Pages 專案 (`D:\program\Github\LKC1958_June_1.github.io`)**
  * `config.js` (修改：動態註冊根目錄下的 `service-worker.js`)
  * `manifest.json` (新建：設定 PWA 基本資訊與離線啟動參數)
  * `service-worker.js` (新建：Stale-While-Revalidate 離線快取核心邏輯)

---

### 任務 1：後台會友管理導入 LockService 排隊鎖定

**文件：**
- 修改：`D:/program/LKC/主日出席_測試版/MemberDB.js`

- [ ] **步驟 1：修改 `addMember` 導入鎖定機制**
  在 `addMember(member)` 函數開頭加上獲取 Script Lock 邏輯，並在函數結束及出錯時釋放鎖。
  修改 `addMember`：
  ```javascript
  function addMember(member) {
    const lock = LockService.getScriptLock();
    try {
      // 嘗試獲取鎖，最多等候 10 秒
      lock.waitLock(10000);
      
      const sheet = getMemberSheet();
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == member.name) return "⚠️ 新增失敗：姓名 [" + member.name + "] 已存在！";
      }
      const now = new Date();
      const uid = generateNextUID(data);
      const role = _normalizeRole(member.role);
      sheet.appendRow([member.name, member.gender, now, member.note, member.isExcluded, now, "初始建立", uid, "", member.group || "", role]);
      invalidateAndRebuildMemberCache();
      firebaseInvalidate(['getAllMembers', 'getAllGroupMembers', 'getMemberSuggestions', 'getStats', 'getAllGroupsStats', 'getAdminGroupsList', 'ministry_getPageConfig', 'ministry_getGroupMembers', 'getAttendanceStats']);
      return "✅ 新增成功！編號：" + uid;
    } finally {
      // 確保不論成功或失敗都會放鎖
      lock.releaseLock();
    }
  }
  ```

- [ ] **步驟 2：修改 `updateMember` 導入鎖定與重新讀取最新值**
  在 `updateMember(oldName, newData)` 中，必須先拿鎖，且在拿鎖成功後**再從 Sheet 讀取最新狀態**，才進行修改寫回，確保不會發生覆蓋。
  修改 `updateMember`：
  ```javascript
  function updateMember(oldName, newData) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      
      const sheet = getMemberSheet();
      // 拿鎖後才讀取最新 Sheet 狀態
      const data = sheet.getDataRange().getValues();
      const now = new Date();

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == oldName) {
          const rowIndex = i + 1;

          const finalName     = newData.name       !== undefined ? newData.name       : data[i][0];
          const finalGender   = newData.gender     !== undefined ? newData.gender     : data[i][1];
          const finalNote     = newData.note       !== undefined ? newData.note       : data[i][3];
          const finalExcluded = newData.isExcluded !== undefined ? newData.isExcluded : data[i][4];
          const finalGroup    = newData.group      !== undefined ? String(newData.group || "").trim() : (data[i][9] || "");
          const finalRole     = newData.role       !== undefined ? String(newData.role || "小羊").trim() : (data[i][10] || "小羊");

          const changeLog = [];
          if (data[i][0] != finalName)   changeLog.push("姓名: " + data[i][0] + "->" + finalName);
          if (data[i][1] != finalGender) changeLog.push("性別: " + data[i][1] + "->" + finalGender);
          if (data[i][3] != finalNote)   changeLog.push("備註異動");
          const oldExcluded = (data[i][4] === true || data[i][4] === "TRUE");
          const newExcluded = (finalExcluded === true || finalExcluded === "TRUE");
          if (oldExcluded !== newExcluded) changeLog.push("統計狀態變更");
          const oldGroup = data[i][9] ? String(data[i][9]).trim() : "";
          if (oldGroup !== finalGroup)   changeLog.push("所屬小組: " + (oldGroup || "(無)") + "->" + (finalGroup || "(無)"));
          const oldRole = data[i][10] ? String(data[i][10]).trim() : "小羊";
          if (oldRole !== finalRole)     changeLog.push("身分: " + oldRole + "->" + finalRole);

          if (changeLog.length > 0) {
            sheet.getRange(rowIndex, 1, 1, 7).setValues([[
              finalName, finalGender, data[i][2], finalNote, finalExcluded, now, changeLog.join(" | ")
            ]]);
            sheet.getRange(rowIndex, 10).setValue(finalGroup);
            sheet.getRange(rowIndex, 11).setValue(finalRole);
            invalidateAndRebuildMemberCache();
            firebaseInvalidate(['getAllMembers', 'getAllGroupMembers', 'getMemberSuggestions', 'getStats', 'getAllGroupsStats', 'getAdminGroupsList', 'ministry_getPageConfig', 'ministry_getGroupMembers', 'getAttendanceStats']);
            return "✅ 更新成功！";
          } else {
            return "⚠️ 資料無異動";
          }
        }
      }
      return "❌ 找不到原始資料，無法更新";
    } finally {
      lock.releaseLock();
    }
  }
  ```

- [ ] **步驟 3：修改 `deleteMember` 導入鎖定**
  修改 `deleteMember` 函數，加鎖以保證安全。
  ```javascript
  function deleteMember(name) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      const sheet = getMemberSheet();
      const data = sheet.getDataRange().getValues();
      const targetName = name.toString().trim();
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0].toString().trim() === targetName) {
          sheet.deleteRow(i + 1);
          invalidateAndRebuildMemberCache();
          firebaseInvalidate(['getAllMembers', 'getAllGroupMembers', 'getMemberSuggestions', 'getStats', 'getAllGroupsStats', 'getAdminGroupsList', 'ministry_getPageConfig', 'ministry_getGroupMembers', 'getAttendanceStats']);
          return "🗑️ 成功刪除會友: " + targetName;
        }
      }
      return "❌ 找不到會友 [" + targetName + "]";
    } catch (e) {
      return "❌ 刪除過程出錯: " + e.toString();
    } finally {
      lock.releaseLock();
    }
  }
  ```

---

### 任務 2：後端小組點名與設定導入 LockService 鎖定

**文件：**
- 修改：`D:/program/LKC/主日出席_測試版/GroupCore.js`

- [ ] **步驟 1：修改 `updateGroupInfo` 導入鎖定**
  修改 `updateGroupInfo` 函數，在寫入小組清單或更名工作表時加上鎖定防寫衝突。
  ```javascript
  function updateGroupInfo(uuid, oldName, newName, newCode, newStatus) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      
      if (!uuid) return { success: false, message: "缺少小組系統識別碼 (UUID)" };

      const listSheet = getGroupSheet("小組清單");
      if (!listSheet) return { success: false, message: "找不到小組清單" };

      const data = listSheet.getDataRange().getValues();
      let targetRowIndex = -1;

      for (let i = 1; i < data.length; i++) {
        if (data[i][4] && String(data[i][4]).trim() === String(uuid).trim()) {
          targetRowIndex = i + 1;
          break;
        }
      }

      if (targetRowIndex === -1) return { success: false, message: "系統錯誤：查無此小組的系統識別碼" };

      const cleanNewName = String(newName).trim();
      const cleanOldName = String(oldName).trim();

      if (cleanOldName !== cleanNewName) {
        for (let i = 1; i < data.length; i++) {
          if (i + 1 !== targetRowIndex && data[i][0] && String(data[i][0]).trim() === cleanNewName) {
            return { success: false, message: "新名稱已與其他小組重複，請換一個名字！" };
          }
        }
      }

      const finalStatus = (newStatus !== undefined && newStatus !== null)
        ? String(newStatus).trim()
        : data[targetRowIndex - 1][1];
      listSheet.getRange(targetRowIndex, 1, 1, 3).setValues([[
        cleanNewName,
        finalStatus,
        String(newCode).trim()
      ]]);

      if (cleanOldName !== cleanNewName) {
        const nameSheet = getGroupSheet(cleanOldName + "_名單");
        if (nameSheet) nameSheet.setName(cleanNewName + "_名單");

        const recordSheet = getGroupSheet(cleanOldName + "_點名紀錄");
        if (recordSheet) recordSheet.setName(cleanNewName + "_點名紀錄");
      }

      _rebuildGroupsCache();
      firebaseInvalidate(['getGroups', 'getAdminGroupsList', 'ministry_getGroups']);
      return { success: true };
    } catch (e) {
      return { success: false, message: "後端執行錯誤：" + e.message };
    } finally {
      lock.releaseLock();
    }
  }
  ```

- [ ] **步驟 2：確認並追加 `submitAttendance` 鎖定邏輯**
  在小組提交點名記錄 `submitAttendance` 時一併加鎖，防範同一小組多人點名時的紀錄覆蓋（若原代碼已存在，則加強，若無則補上）。

---

### 任務 3：建立 PWA Manifest 配置與基礎資源

**文件：**
- 創建：`D:/program/Github/LKC1958_June_1.github.io/manifest.json`

- [ ] **步驟 1：建立 `manifest.json`**
  宣告離線快取 App 的基本配置，指向統一的啟動 URL 與圖示。
  寫入 `manifest.json`：
  ```json
  {
    "short_name": "LKC系統",
    "name": "LKC 教會管理系統",
    "icons": [
      {
        "src": "https://jirehwang.github.io/LKC1958_June_1.github.io/docs/images/logo_192.png",
        "type": "image/png",
        "sizes": "192x192"
      },
      {
        "src": "https://jirehwang.github.io/LKC1958_June_1.github.io/docs/images/logo_512.png",
        "type": "image/png",
        "sizes": "512x512"
      }
    ],
    "start_url": "/LKC1958_June_1.github.io/index.html",
    "background_color": "#1a202c",
    "theme_color": "#667eea",
    "display": "standalone",
    "orientation": "portrait"
  }
  ```

- [ ] **步驟 2：在所有主要子應用的 HTML 檔案中引用 `manifest.json`**
  確保同工開啟各子系統（主日點名、小組點名、敬拜團）時，手機可以正確識別為 PWA。
  在 `LKC_Group/index.html`、`LKC_SundayserviceAttendance/index.html`、`LKC_worship/admin.html` 的 `<head>` 中插入：
  ```html
  <link rel="manifest" href="/LKC1958_June_1.github.io/manifest.json">
  ```

---

### 任務 4：編寫 PWA Service Worker 核心快取邏輯

**文件：**
- 創建：`D:/program/Github/LKC1958_June_1.github.io/service-worker.js`

- [ ] **步驟 1：編寫 `service-worker.js`**
  實作 `install`、`activate` 與 `fetch` 攔截。
  * `Stale-While-Revalidate` 策略：靜態資源（CSS、JS、HTML、圖片、Web Fonts）優先從 cache 回傳，並於背景拉取更新寫回 cache。
  * 對任何包含 `.google.com` (GAS) 或 Firebase 認證等動態 POST API 直接進行網路請求，絕對不進快取。
  寫入 `service-worker.js`：
  ```javascript
  const CACHE_NAME = 'lkc-pwa-cache-v1';
  const PRECACHE_ASSETS = [
    '/LKC1958_June_1.github.io/index.html',
    '/LKC1958_June_1.github.io/config.js',
    '/LKC1958_June_1.github.io/manifest.json'
  ];

  // 1. 安裝事件：預先快取核心資源
  self.addEventListener('install', event => {
    event.waitUntil(
      caches.open(CACHE_NAME)
        .then(cache => cache.addAll(PRECACHE_ASSETS))
        .then(() => self.skipWaiting())
    );
  });

  // 2. 啟用事件：清理舊版快取
  self.addEventListener('activate', event => {
    event.waitUntil(
      caches.keys().then(keys => {
        return Promise.all(
          keys.map(key => {
            if (key !== CACHE_NAME) {
              return caches.delete(key);
            }
          })
        );
      }).then(() => self.clients.claim())
    );
  });

  // 3. 攔截請求事件：Stale-While-Revalidate 策略
  self.addEventListener('fetch', event => {
    const url = new URL(event.request.url);

    // 🛡️ 安全防護：GAS 請求、Firebase 認證與所有 POST 請求一律直接過網，不快取
    if (event.request.method === 'POST' || 
        url.hostname.includes('script.google.com') || 
        url.hostname.includes('firebasedatabase.app')) {
      return; // 讓瀏覽器直接去網路請求
    }

    // 僅快取與本站相關的 GET 靜態資源
    if (event.request.method === 'GET') {
      event.respondWith(
        caches.open(CACHE_NAME).then(cache => {
          return cache.match(event.request).then(cachedResponse => {
            // 發送網路請求做背景更新
            const fetchPromise = fetch(event.request).then(networkResponse => {
              if (networkResponse && networkResponse.status === 200) {
                // 快取新版資源
                cache.put(event.request, networkResponse.clone());
              }
              return networkResponse;
            }).catch(err => {
              console.warn('[SW] 背景更新失敗（可能處於離線狀態）:', err);
            });

            // 快取命中則秒回快取；否則等待網路回傳
            return cachedResponse || fetchPromise;
          });
        })
      );
    }
  });
  ```

---

### 任務 5：於前端 config.js 實現 Service Worker 自動註冊

**文件：**
- 修改：`D:/program/Github/LKC1958_June_1.github.io/config.js`

- [ ] **步驟 1：修改 `config.js` 實現無感自動註冊**
  在 `config.js` 的 IIFE (立即執行函數) 尾端加入動態 Scope 判別與 SW 註冊碼，讓所有子系統無感套用。
  修改 `config.js`：
  ```javascript
  // ─────────────────────────────────────────────────────────────
  //  PWA Service Worker 自動註冊（Stale-While-Revalidate 靜默更新版）
  // ─────────────────────────────────────────────────────────────
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function() {
      // 依據 hostname 動態判定 scope 路徑
      // 本地環境可能是 "/"，而 GitHub Pages 是 "/LKC1958_June_1.github.io/"
      let scopePath = '/';
      let swPath = '/service-worker.js';

      if (window.location.hostname.indexOf('github.io') !== -1) {
        scopePath = '/LKC1958_June_1.github.io/';
        swPath = '/LKC1958_June_1.github.io/service-worker.js';
      }

      navigator.serviceWorker.register(swPath, { scope: scopePath })
        .then(function(reg) {
          console.log('✅ [PWA] ServiceWorker 註冊成功，Scope: ', reg.scope);
        }).catch(function(err) {
          console.warn('❌ [PWA] ServiceWorker 註冊失敗: ', err);
        });
    });
  }
  ```

---

### 任務 6：部署與本地測試驗證

- [ ] **步驟 1：Clasp Push 部署 GAS 專案**
  執行 `clasp push -f` 將 LockService 後台排隊代碼部署至測試環境。

- [ ] **步驟 2：測試 PWA 註冊與離線加載**
  本地啟動伺服器，載入點名系統，於開發者工具（F12）的 Application -> Service Workers 中確認 Service Worker 狀態為 `Activated`，且在 Network 設為 `Offline`（離線狀態）時，網頁外殼依然能秒開，確認 API 讀取能正確 bypass。
