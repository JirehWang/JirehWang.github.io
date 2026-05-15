<!--
  📄 ARCHITECTURE.md
  教會系統當前架構說明 (測試版)

  本文件記錄方案 C（核心 3 合 1 整合）+ Firebase RTDB 快取 完成後的
  完整架構設計，並對比原始（test repo 既有 README 描述的）舊架構，
  方便日後維護與正式版上線的決策參考。

  最後更新：2026-05-15（含 Firebase Service Account OAuth2 整合）
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
- `{小組名}_名單` 為**主日的快取鏡像**（4 欄含 UID）
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

### 啟用快取的 13 個 API（TTL 6 小時）
| 類別 | API |
|---|---|
| 全域 | getGroups, getGroupConfig, getWeeklyReport, getAllMembers, getAdminGroupsList |
| 統計 | getStats, getAllGroupsStats, getAttendanceStats, getAttendanceTrend |
| 事工 | ministry_getGroups, ministry_getTemplates, ministry_getAggregatedReport, ministry_getPageConfig |

### 三層 Invalidation（讓資料即時同步）
1. **前端**：寫入 action 自動清相關 read cache (`_INVALIDATE_ON_WRITE`)
2. **GAS 端**：CRUD 函式內呼叫 `firebaseInvalidate([...])`
3. **試算表手動編輯**：`onEdit` trigger 自動清

### 認證
- 前端：Firebase Web SDK + 開放 read（無個資範圍）
- 後端：Service Account JSON → JWT → OAuth2 Bearer Token

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
├── GroupCore.js           # 小組管理（getGroups / verifyGroup / etc.）
├── GroupAttendance.js     # 小組點名（含 checkGroupStatus）
├── GroupStatistics.js     # 小組統計
├── MinistryCore.js        # 事工管理整套
├── GeminiHelper.js        # 共用 AI helper
├── FirebaseSync.js        # GAS 端 Firebase RTDB 寫入 (Service Account OAuth2)
├── MigrationTools.js      # 一次性遷移：身分搬到主日 + 找重名/同音
├── MigrationAttendance.js # 一次性遷移：點名紀錄姓名 → UID
├── QRCODE.js              # QR 掃描處理
├── QRcodeMaker.js         # QR / 卡片產生
├── Monitor.js             # 額度監控
└── 各 HTML（attendance/members/STATS/Chart/index）
```

---

## 📈 演進路線圖

| 階段 | 狀態 | 備註 |
|---|---|---|
| Phase 1：方案 B 整合（小組併入主日） | ✅ 完成 | 消除跨 GAS UrlFetch |
| Phase 2：方案 C 整合（事工併入主日） | ✅ 完成 | 共用 master 名單 |
| Phase 3：身分多組格式 + UID 化點名 | ✅ 完成 | 避免重名、改名造成混淆 |
| Phase 4：Firebase RTDB 快取層 | ✅ 完成 | 三層 invalidation |
| Phase 5：教會行事曆併入主 GAS | ⏸ 待測試副本 | 預計併入後再消滅一個獨立 GAS |
| Phase 6：敬拜團併入主 GAS | ⏸ 規劃中 | 與事工共用 GeminiHelper |
| Phase 7：正式版切換 | ⏸ 待測試穩定 | 把 _TEST 改為正式版 GAS / 試算表 ID |

---

> 📝 **本文件對應 commit hash**：請以 git log 的最新 ARCHITECTURE.md 修改為準
> 🔗 **舊架構文件**：各 `apps/*/README.md`（多為原始 GAS 獨立部署的描述）
