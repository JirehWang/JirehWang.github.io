<!--
  📄 ARCHITECTURE.md
  教會系統當前架構說明 (測試版)

  本文件記錄方案 C（核心 3 合 1 整合）+ Firebase RTDB 快取 完成後的
  完整架構設計，並對比原始（test repo 既有 README 描述的）舊架構，
  方便日後維護與正式版上線的決策參考。

  最後更新：2026-05-16
  本次大更新：
    - 主日 members.html 簡化（移除所屬小組/身分欄）
    - 小組統計新增「總小組成員清單」(admin)
    - 小組成員拖曳排序 + 暱稱機制
    - 新增成員 datalist 自動完成（從主日抓會友候選）
    - Firebase Service Account OAuth2 整合 + 安全強化
    - 三層 invalidation 全面就緒
-->

# 🏛️ 教會系統架構文件（測試版）

## 📦 三層式快取架構（最終形態）

```
┌──────────────────────────────────────────────────────────────┐
│                      使用者瀏覽器                              │
│                  (GitHub Pages 前端)                          │
└────────────────────────┬─────────────────────────────────────┘
                         │ ① churchAPI(action, data)
                         ▼
┌──────────────────────────────────────────────────────────────┐
│             第一道：Firebase Realtime Database               │
│             (cache 命中即返回，~50ms)                         │
│             路徑：cache/{action}/{dataHash}                  │
└────────────────────────┬─────────────────────────────────────┘
                         │ ② cache miss → 打 GAS
                         ▼
┌──────────────────────────────────────────────────────────────┐
│             第二道：Google Apps Script (主 GAS)               │
│             (業務邏輯 + Sheet I/O + Firebase 反向更新)        │
└────────────────────────┬─────────────────────────────────────┘
                         │ ③ 讀寫
                         ▼
┌──────────────────────────────────────────────────────────────┐
│             真實來源：Google Sheets                           │
│             (主日 / 小組 / 事工 / 教會行事曆 4 份試算表)       │
└──────────────────────────────────────────────────────────────┘
```

---

## 🌳 子系統與 GAS 對應

| 前端子系統 | URL key | 後端 GAS | 整合狀態 |
|---|---|---|---|
| 主日點名 | `LKC_SundayserviceAttendance_TEST` | **主日_測試版 GAS** | ✓ 主 |
| 小組點名 | `LKC_Group_TEST` | **主日_測試版 GAS** | ✓ 已併入 |
| 事工管理 | `LKC_MinistrySchedule_TEST` | **主日_測試版 GAS**（action 加 `ministry_` 前綴） | ✓ 已併入 |
| 教會行事曆 | `LKC_MasterSchedule` | 教會行事曆 GAS（獨立） | ⏸ 待整合 |
| 敬拜團 | `LKC_worship` | 敬拜團 GAS（獨立） | ⏸ 待整合 |
| 車號查詢 | `LKC_WhosCar` | 車號查詢 GAS（獨立） | ⏸ 暫擱置 |
| 週報管理 | `LKC_SundayBulletin` | 週報 GAS（獨立） | ⏸ 暫擱置 |

---

## 🗂️ 試算表角色

### 1. 主日試算表 (master)
- **會友名單**（**單一真實來源**，11 欄含「身分」）
- 各場次點名紀錄（台語 / 華語 / 聯合 / 主日學A班/B班）
- 點名系統清單、SYNC_TEMP、系統監控

### 2. 小組試算表
- 小組清單（含 UUID + 代碼）
- `{小組名}_名單` 為**主日的快取鏡像 + 小組私有設定**（6 欄）
  - A 姓名 / B 建立日期 / C 身分（由主日同步）
  - D 系統編號（UID，由主日同步）
  - E 排序（小組私有，拖曳結果）
  - F 暱稱（小組私有，使用者自訂）
- `{小組名}_點名紀錄`（**全 UID 化**）

### 3. 事工試算表（測試副本）
- Config（含 UUID 7 欄結構）
- 各小組排班分頁、模板、審計日誌

### 4. 教會行事曆試算表
- 聚會資料、事工細項、講道資訊
- 事工系統 openById 唯讀

---

## 🏷️ 前端路由機制

`config.js` 統一中央路由：
- 各頁面 HTML 宣告 `window._GAS_KEY = 'LKC_XXX_TEST'`
- 中央 `_URL_ROUTER` 對應到實際 GAS URL
- `_ACTION_PREFIX` 自動為某些子系統加 action 前綴（例：`ministry_` → 後端能區別）

---

## 🔥 Firebase 快取設計

### 結構
```
cache/
├── getGroups/_default
├── getStats/<dataHash>
├── ministry_getPageConfig/<dataHash>
└── ...
```

### 啟用快取的 15 個 API（TTL 統一 6 小時）
| 類別 | API |
|---|---|
| 全域 | getGroups, getGroupConfig, getWeeklyReport, getAllMembers, getAdminGroupsList |
| 統計 | getStats, getAllGroupsStats, getAttendanceStats, getAttendanceTrend |
| 小組（管理員 / 輔助） | **getAllGroupMembers**（admin 總清單）、**getMemberSuggestions**（datalist 自動完成） |
| 事工 | ministry_getGroups, ministry_getTemplates, ministry_getAggregatedReport, ministry_getPageConfig |

### 三層 Invalidation（讓資料即時同步）
1. **前端**：寫入 action 自動清相關 read cache (`_INVALIDATE_ON_WRITE`)
2. **GAS 端**：CRUD 函式內呼叫 `firebaseInvalidate([...])`
3. **試算表手動編輯**：`onEdit` trigger 自動清

### 認證
- 前端：Firebase Web SDK
  - apiKey 公開於 firebase-config.js（Firebase 官方做法）
  - 受 **HTTP Referrer 限制**（GCP Console 設定）保護：只能從 jirehwang.github.io 呼叫
  - Firestore/RTDB Rules：`cache.read = true` / `cache.write = "auth != null"`
- 後端：Service Account JSON → JWT → OAuth2 Bearer Token
  - JSON 內容存在 GAS Script Properties (`FIREBASE_SERVICE_ACCOUNT`)
  - JSON 檔本身在 `D:\program\LKC\` 根目錄，不在任何 GAS / git 推送範圍
  - `.gitignore` 規則：`*firebase-adminsdk*.json` 預防意外提交

---

## 📝 與原始架構的差異對照表

> 以下「原始」指 test repo 內各 README 描述的舊狀態（GAS 個別獨立、無快取層、姓名做 key）

| 面向 | 原始架構 | 當前架構 | 變化原因 |
|---|---|---|---|
| **GAS 數量** | 每個子系統各一個（共 7 個 GAS） | 主日/小組/事工三系統合一（其餘待整合） | 消除跨 GAS UrlFetch、降冷啟動成本 |
| **跨 GAS 呼叫** | 小組 → 主日（取會友名單）每次 6-15 秒 | 同 GAS 內 function call 0.5-1 秒 | 砍掉雙重冷啟動 |
| **前端路由** | 每子系統各自有 config.js + api.js | 中央 `config.js` 統一管理路由與 churchAPI | 統一維護、避免 URL 散落 |
| **點名紀錄 key** | 姓名（含性別後綴 `(男)/(女)`） | **系統編號 (LK00001)** | 避免重名混淆、姓名修改不影響歷史 |
| **會友身分** | 存在小組 `_名單` 的身分欄 | 存在主日「會友名單」的「身分」欄（支援多組格式） | 主日為單一真實來源 |
| **多組支援** | 一人只能屬一組 | **所屬小組**「、」分隔 / **身分**「核心同工(A組)、一般同工(B組)」 | 反映實際情況（一人多組） |
| **GAS 冷啟動延遲** | 5-7 秒（每次首頁開啟） | < 1 秒（keepWarm 5 分鐘 trigger） | 確保 runtime 永遠熱機 |
| **快取層** | 無 | Firebase RTDB（13 API 走 cache，TTL 6h，3 層 invalidation） | 毫秒級回應、減少 GAS quota |
| **GAS 內部 cache** | 無 | CacheService（成員名單 / 點名計數 / 群組設定） | 同次請求內共用、降 Sheet I/O |
| **事工 → 小組名單** | openById 跨 SS 讀取（~500ms） | 直接呼叫 `getCachedMembers()` 過濾 | 整合後同 GAS function call |
| **Gemini AI 程式** | 教會行事曆 + 事工管理各寫一份 | 共用 `GeminiHelper.js` | 維護成本減半 |
| **每日 GAS 額度** | 估 60-100 分鐘 / 天 | 估 10-20 分鐘 / 天 | quota 釋放給真正的業務操作 |
| **個資儲存位置** | 全程在 GAS / Sheets | 部分快取在 Firebase（read-only 開放） | 前端讀取速度 vs 安全性權衡（內部使用可接受） |

---

## 🛡️ 安全分層

| 層 | 手段 |
|---|---|
| 中央 config.js | `AUTH_TOKEN = "ChurchApp-2026"` 隨請求送出 |
| GAS doPost | 小組/事工 action 強制 token 比對（主日歷史用 payload 格式） |
| Firebase Rules | `cache.read = true` / `cache.write = auth != null`（防匿名寫入垃圾） |
| Service Account | JSON 在 GAS Script Properties，**不在 git、不在 clasp push 範圍** |
| `.gitignore` | `*firebase-adminsdk*.json` / `.clasprc.json` 禁止意外上傳 |

---

## 🚀 部署位置一覽

| 項目 | 位置 | 更新方式 |
|---|---|---|
| 前端網頁 | GitHub Pages（`jirehwang.github.io/LKC1958_June_1.github.io/`） | `git push` 觸發自動部署 |
| 主 GAS | Apps Script 雲端（測試 deployment ID `AKfycbxBOFeLiX...`） | `clasp push` + `clasp deploy` |
| 教會行事曆 GAS | 獨立 Apps Script | `clasp push`（個別） |
| 敬拜團 GAS | 獨立 Apps Script | 同上 |
| 車號 / 週報 GAS | 獨立 Apps Script | 同上 |
| Firebase RTDB | `lkc1958june1-default-rtdb.asia-southeast1` | 由前端/GAS 自動寫入 |

---

## 🧰 GAS 主 Project 檔案清單

```
主日出席_測試版/
├── Core.js                # doPost 三路由（主日/小組/事工）
├── MemberDB.js            # 會友 CRUD + UID/Name lookups + parseAttendanceList
├── AttendanceDB.js        # 主日點名（全 UID 化）
├── CacheManager.js        # GAS 內部 CacheService 管理（成員名單 + 群組設定）
├── ChartDB.js             # 主日點名圖表
├── ReportService.js       # 主日統計（getAttendanceStats / Trend）
├── GroupCore.js           # 小組管理（getGroups / verifyGroup / getMemberSuggestions / getAllGroupMembers）
├── GroupAttendance.js     # 小組點名（含 checkGroupStatus / 排序 / 暱稱寫入）
├── GroupStatistics.js     # 小組統計
├── MinistryCore.js        # 事工管理整套
├── GeminiHelper.js        # 共用 AI helper
├── FirebaseSync.js        # GAS 端 Firebase RTDB 寫入（Service Account OAuth2 JWT）
│                          #   + onEditMain/Group/Ministry 三個 sheet 編輯偵測 trigger
├── MigrationTools.js      # 一次性遷移：身分搬到主日 + 找重名/同音
├── MigrationAttendance.js # 一次性遷移：點名紀錄姓名 → UID
├── QRCODE.js              # QR 掃描處理
├── QRcodeMaker.js         # QR / 卡片產生
├── Monitor.js             # 額度監控
└── 各 HTML（attendance/members/STATS/Chart/index）
```

> 📝 共享給維護人的 GAS 函式（手動執行一次）：
> - `setupKeepWarmTrigger`：5 分鐘 keep-warm trigger（防冷啟動）
> - `setupAllOnEditTriggers`：3 個試算表的 onEdit trigger
> - `setupMinistryAutoSyncTrigger`：事工小組清單自動同步
> - `setupAuditLogFlushTrigger`：審計日誌 5 分鐘批次寫
> - `testFirebaseAuth`：驗證 Service Account 設定
> - `firebaseCacheClearAll`：緊急清空所有 Firebase cache

---

## 📈 演進路線圖

| 階段 | 狀態 | 備註 |
|---|---|---|
| Phase 1：方案 B 整合（小組併入主日） | ✅ 完成 | 消除跨 GAS UrlFetch |
| Phase 2：方案 C 整合（事工併入主日） | ✅ 完成 | 共用 master 名單 |
| Phase 3：身分多組格式 + UID 化點名 | ✅ 完成 | 避免重名、改名造成混淆 |
| Phase 4：Firebase RTDB 快取層 | ✅ 完成 | 三層 invalidation |
| Phase 4.5：成員管理 UX 強化 | ✅ 完成 | 拖曳排序 / 暱稱 / datalist / admin 總清單 |
| Phase 4.6：安全強化 | ✅ 完成 | Service Account / HTTP Referrer / .gitignore |
| Phase 5：教會行事曆併入主 GAS | ⏸ 待測試副本 | 預計併入後再消滅一個獨立 GAS |
| Phase 6：敬拜團併入主 GAS | ⏸ 規劃中 | 與事工共用 GeminiHelper |
| Phase 7：正式版切換 | ⏸ 待測試穩定 | 把 _TEST 改為正式版 GAS / 試算表 ID |

---

## 📜 更新紀錄 (Changelog)

### 2026-05-16（Phase 4.5 + 4.6 大批更新）

#### 🆕 新功能
| 模組 | 改動 |
|---|---|
| **主日 members.html** | 拿掉「所屬小組」「身分」欄位（移交小組系統管理）；表格 6 欄、modal 簡化 |
| **小組統計（admin）** | 新增「📋 總小組成員清單」報表，顯示：姓名 / 性別 / 系統編號 / 所屬小組 / 身分；僅列出有歸組的會友 |
| **小組「管理名單與身分」modal** | 1. 拖曳排序（Sortable.js ⋮⋮ 手把）<br>2. 新增成員改成 datalist 自動完成（從主日載入候選）<br>3. 每筆加暱稱輸入欄 |
| **小組點名介面顯示** | 從「姓名」變成「姓名 (暱稱)」；無暱稱就只顯示姓名 |
| **修改歷史紀錄 modal** | 顯示風格同上 |
| **datalist 智慧顯示** | 唯一同名 → 純姓名；同名多人 → 補 `(LK00001)` 區分 |

#### 🔧 後端 API
| API | 用途 | TTL | 權限 |
|---|---|---|---|
| `getAllGroupMembers` | 取所有有歸組的會友（管理員專用） | 6h | ADMIN_CODE |
| `getMemberSuggestions` | 取所有會友 name+uid（給小組 datalist） | 6h | 一般 |

#### 🗄️ Schema 變化
- **主日「會友名單」**：原本 10 欄（K 欄「身分」是 phase 3 加的）— 此次無新欄位
- **小組 `_名單` sheet**：4 欄 → 6 欄
  - 原：A 姓名 / B 建立日期 / C 身分 / D 系統編號
  - 新增：E **排序**（拖曳結果）/ F **暱稱**（小組私有別名）

#### 🛡️ 安全強化
1. **Firebase Service Account 整合**
   - 從 Database Secret（legacy）改成 Service Account OAuth2 JWT
   - JSON 存在 `GAS Script Properties: FIREBASE_SERVICE_ACCOUNT`
   - GAS 端 RTDB 寫入流程：JWT → Token API → Bearer Token
2. **HTTP Referrer 限制**
   - Firebase Web API Key 在 GCP Console 加上 referrer 白名單
   - 只接受 `jirehwang.github.io/*` 的呼叫
   - 即使 Key 外洩，攻擊者無法從其他網域使用
3. **`.gitignore` 預防**
   - 加上 `*firebase-adminsdk*.json` / `.clasprc.json` 黑名單
   - 教會行事曆的 Gemini API Key 從硬編碼搬到 Script Properties

#### 🐛 修補
- **事工管理 fetchAPI 整合 churchAPI**：原本沒走中央 churchAPI、缺 `ministry_` 前綴、body 格式不對，整合方案 C 後失效 → 修復後自動走前綴 + Firebase cache
- **點名紀錄遷移工具**：`migrateAttendanceRecordsToUid` 將舊「姓名(性別)」格式自動轉成 LK 代碼

#### 📊 cache 數量
- 12 → **15 個 API** 走 Firebase 快取
- TTL 全部統一為 **6 小時**（依賴 invalidation 為主，TTL 兜底）

#### 🧹 invalidation 同步
- 各 CRUD 寫入後 + onEdit trigger 都同步清掉相關 cache topic
- 包含新增的 `getAllGroupMembers` / `getMemberSuggestions`

#### 📝 文件
- 本文件 `ARCHITECTURE.md` 全面更新（試算表結構 / API 清單 / 安全層 / Changelog）

---

### 2026-05-15（Phase 1-4 完成）
- 方案 B 整合：小組系統併入主日 GAS（消除跨 GAS UrlFetch）
- 方案 C 整合：事工管理併入主日 GAS（共用 master 名單）
- UID 化點名紀錄：避免重名混淆
- 身分多組格式：核心同工(A組)、一般同工(B組)
- Firebase RTDB 快取層上線：13 API + 三層 invalidation

---

> 📝 **本文件對應 commit hash**：請以 git log 的最新 ARCHITECTURE.md 修改為準
> 🔗 **舊架構文件**：各 `apps/*/README.md`（多為原始 GAS 獨立部署的描述）
> 💡 **共同維護者請特別注意**：本文件「更新紀錄」段是最常變動的部分，做完功能請隨手記錄
