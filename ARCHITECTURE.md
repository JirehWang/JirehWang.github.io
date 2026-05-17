<!--
  📄 ARCHITECTURE.md
  教會系統當前架構說明 (測試版)

  本文件記錄方案 C（核心 3 合 1 整合）+ Firebase RTDB 快取 完成後的
  完整架構設計，並對比原始（test repo 既有 README 描述的）舊架構，
  方便日後維護與正式版上線的決策參考。

  最後更新：2026-05-17
  本次大更新：
    - 主日點名 5 個 GET API 納入 Firebase 快取（點名介面 / 小組首頁 / 圖表 / 登入驗證）
    - 敬拜團 5 個 GET API 全部納入 Firebase 快取（公佈欄/服事表/位置/團員/曲目）
    - 敬拜團：新增「敬拜團員名單」分頁（從主日會友拉選，正式/實習狀態）
    - 敬拜團：「位置與同工」改為標籤式多選，資料來源綁定敬拜團員名單
    - 共用可搜尋浮動下拉元件（替代 HTML5 datalist，UX 強化）
    - 小組移除成員時同步更新主日的「所屬小組 / 身分」
    - 小組統計（單日）補入「當日有出席但已不在組」的歷史成員
    - 快取 API 數量：15 → 25
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
| 敬拜團 | `LKC_worship` | 敬拜團 GAS（獨立，含「敬拜團員名單」） | ⏸ 待整合（已強化前端） |
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

### 啟用快取的 25 個 API（TTL 統一 6 小時）
| 類別 | API |
|---|---|
| 全域 | getGroups, getGroupConfig, getWeeklyReport, getAllMembers, getAdminGroupsList |
| **主日點名介面** | **getSmartAttendanceList**（點名介面首載）、**checkGroupStatus**（小組首頁）、**findGroupByCode** / **verifyGroup**（登入驗證） |
| 統計 / 圖表 | getStats, getAllGroupsStats, getAttendanceStats, getAttendanceTrend, **getCategoryChartData** |
| 小組（管理員 / 輔助） | **getAllGroupMembers**（admin 總清單）、**getMemberSuggestions**（datalist 自動完成） |
| 事工 | ministry_getGroups, ministry_getTemplates, ministry_getAggregatedReport, ministry_getPageConfig, ministry_getGroupMembers |
| 敬拜團 | getSchedule, getScheduleByDateRange, getPositions, getTeamMembers, getSongs |

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
| Phase 4.7：小組移除同步 + 統計補歷史 | ✅ 完成 | 主日身分/組別與點名歷史一致 |
| Phase 4.8：敬拜團員名單 + 位置選人改造 | ✅ 完成 | 敬拜團員清單為單一來源、可搜尋下拉 |
| Phase 5：教會行事曆併入主 GAS | ⏸ 待測試副本 | 預計併入後再消滅一個獨立 GAS |
| Phase 6：敬拜團併入主 GAS | ⏸ 規劃中 | 與事工共用 GeminiHelper |
| Phase 7：正式版切換 | ⏸ 待測試穩定 | 把 _TEST 改為正式版 GAS / 試算表 ID |

---

## 📜 更新紀錄 (Changelog)

### 2026-05-17（後續：主日點名介面 5 個 GET 納入快取）

#### 🚀 新增 cacheable actions（TTL 6h）
| Action | 用途 | 加速效果 |
|---|---|---|
| `getSmartAttendanceList` | 點名介面開啟時抓會友 + 出席計數 + 同步狀態 | 🔥 主日點名首載 |
| `checkGroupStatus` | 小組點名首頁抓組員 + 啟用狀態 | 🔥 小組首載 |
| `getCategoryChartData` | 趨勢圖表（依分類） | Chart 頁面 |
| `findGroupByCode` | 小組代碼 → 名稱對應（登入） | 登入流程 |
| `verifyGroup` | 小組名 + 代碼驗證 | 同上 |

#### 🧹 對應的寫入失效規則（已併入既有規則）
| 觸發 action | 新增失效 topic |
|---|---|
| `addMember` / `updateMember` / `deleteMember` | + `getSmartAttendanceList` |
| `createGroup` / `updateGroupInfo` | + `findGroupByCode`、`verifyGroup` |
| `updateMemberList` / `ministry_updateGroupMemberRoles` | + `checkGroupStatus` |
| `initGroup`（新增條目） | `checkGroupStatus`、`getGroups`、`getAllGroupMembers` |
| `saveAttendance` / `revokeAttendance` | + `getSmartAttendanceList`、`getCategoryChartData` |
| `submitAttendance` / `updateAttendanceRecord` / `deleteAttendanceRecord` | + `checkGroupStatus`、`getCategoryChartData` |

#### 🛡️ 為何不會影響操作
- **`getSmartAttendanceList`**：點名介面開啟時抓一次，之後是 client-side 即時更新 +  `getQuickSyncData`（**未快取**）持續輪詢補差量
- **`checkGroupStatus`**：小組首頁開啟時抓一次，之後動作完全走 client state
- **不快取的 `getQuickSyncData`、`syncClickToServer`** 確保多裝置即時同步不受影響

#### 📊 快取 API 數量：20 → **25 個**

#### 📁 影響檔案
- `config.js` — _CACHEABLE_ACTIONS 加 5 條、_INVALIDATE_ON_WRITE 加/擴 8 條

---

### 2026-05-17（後續：敬拜團納入 Firebase 快取層）

#### 🚀 敬拜團 GET API 全部走 Firebase 快取
| 模組 | 對應 action | TTL |
|---|---|---|
| 公佈欄總表 | `getSchedule` | 6h |
| 服事表安排（季度） | `getSchedule` | 6h |
| 服事表安排（區間） | `getScheduleByDateRange` | 6h |
| 位置與同工 | `getPositions` | 6h |
| 敬拜團員名單 | `getTeamMembers` | 6h |
| 敬拜曲目 | `getSongs` | 6h |
| 主日會友候選（datalist） | `getMemberSuggestions` | 6h（與主日共享 cache topic，因為兩者讀的是同一份主日「會友名單」） |

#### 🧹 對應的寫入失效規則
| 寫入 action | 失效的 cache topic |
|---|---|
| `saveSchedule` | `getSchedule`、`getScheduleByDateRange` |
| `savePositions` | `getPositions` |
| `saveTeamMembers` | `getTeamMembers` |
| `saveSongs` | `getSongs`、`getSchedule`、`getScheduleByDateRange`（曲目會出現在班表「敬拜曲目」欄） |

#### 📝 注意事項
- 敬拜團 GAS 與主系統 GAS 共用同一個 Firebase RTDB cache topic 命名空間（`cache/{action}/{subkey}`）— 因為敬拜團的 action 名稱（`getSchedule` 等）目前**沒有被任何其他 GAS 使用**，所以不會衝突。
- 敬拜團 GAS 目前**尚未整合 `FirebaseSync.js`**，所以：
  - 前端寫入 → ✅ 立即失效（透過 `_INVALIDATE_ON_WRITE`）
  - 試算表手動編輯 → ⚠️ **無 onEdit invalidation**，最多等 6 小時 TTL 自然到期
  - 若需強一致性，未來可把 FirebaseSync 移植到敬拜團 GAS（Phase 6 整合時順便做）
- 快取 API 數量：15 → **20 個**

#### 📁 影響檔案
- `config.js` — 中央路由設定加入 5 個新 cacheable action + 4 條 invalidate 規則

---

### 2026-05-17（敬拜團 UX 改造 + 小組同步修補）

#### 🆕 敬拜團：可搜尋下拉選單 + 標籤式多選
| 模組 | 改動 |
|---|---|
| **敬拜團員名單分頁（新增區）** | 輸入框右側加「▼ 選擇」按鈕，可開啟主日會友的可搜尋下拉清單；點 input focus 也會自動展開；已加入的會自動過濾 |
| **位置與同工分頁（同工名單）** | 由純文字逗號分隔輸入框 → 改為**標籤式多選**；資料來源**綁定敬拜團員名單**（`getTeamMembers`）；正式 = 藍色徽章、實習 = 黃色徽章；舊資料中不在名單的人會顯示灰色 `⚠️`，提醒整理 |
| **共用可搜尋浮動下拉元件** | `_showFloatingDropdown(anchorEl, items, onPick, opts)` 通用元件：自帶搜尋框、即時關鍵字過濾、外部點擊/捲動/縮放自動關閉；支援 `disabled` 項目（用來顯示「已選」狀態） |
| **儲存後自動失效** | `saveTeamMembersToServer` 成功後清掉 `_worshipTeamCache`，「位置與同工」下次開啟即拉到最新名單 |

**對後端 API 0 影響**：同工名單仍以逗號字串存在 hidden input `.pos-personnel`，savePositions 後端不用任何改動。

#### 🛠️ 小組移除成員時，主日同步更新
- 之前：A 組從 `_名單` 移除某人時，**主日「會友名單」的「所屬小組 / 身分」沒被更新**，導致該人仍被視為 A 組成員
- 修正：`updateMemberList` 函式新增「先計算被移除者」流程：
  1. 找出原本在這組、但新名單沒有的人
  2. 用 `parseGroupRoles` 解析其多組身分 → 移除這組 → `formatGroupRoles` 回寫
  3. 透過 `updateMember` 更新主日的 `所屬小組 / 身分` 欄
- `_saveMemberLocalData` 改為**完全重寫** `_名單`（只保留原建立日期），確保被移除者也從本地快取消失
- 回傳訊息加上「移除 N」統計

#### 🛠️ 小組統計（單日）顯示歷史出席
- 問題：點某日統計時，**只列出目前還在組的人**，若有人已被移除但當天有出席，會被完全忽略
- 原則：**「請保留他原本的點名紀錄，舊的是如何就是如何，不可移除」**
- 修正：`getStats` singleDay 分支也加上補歷史成員邏輯（區間模式之前已有）
  - 把 presentUidSet 中目前不在組的人補進清單，role 標記為 `(歷史)`
- `_點名紀錄` sheet 本身**完全未動**，只是統計呈現補回來

#### 📁 影響檔案
**前端**（已 push）
- `apps/LKC_worship/admin.html` — 新增區下拉按鈕
- `apps/LKC_worship/script.js` — 共用下拉元件、標籤式多選、快取失效

**後端 GAS**（主日_測試版）
- `GroupAttendance.js` — `updateMemberList` 移除同步、`_saveMemberLocalData` 全量覆寫
- `GroupStatistics.js` — `getStats` singleDay 補入歷史成員

---

### 2026-05-16（後續補丁）
- **事工管理：直接編輯小組組員身分** — 小組/團契模板新增「🧑‍🤝‍🧑 設定組員身分」按鈕，免跳系統。
  後端新增 2 個橋接 API：`ministry_getGroupMembers` / `ministry_updateGroupMemberRoles`
- **URL 整合** — 所有寫死的舊獨立 repo 路徑統一指向 `LKC1958_June_1.github.io`：
  - GAS Core.js `doGet`：QR 跳轉的場次頁
  - GAS attendance.html：scannerUrl
  - 前端 attendance.js：scannerUrl + 場次 QR 的 finalUrl
  - 前端 group.js `goToFullStats`：原指 LKGroup.github.io → 改指本 repo apps/LKC_Group/stats.html
  - 目的：未來測試區直接升級為正式區，URL 不需再改

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
