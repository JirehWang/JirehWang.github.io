# LKC New Family

新家人管理系統的 GitHub Pages 前端，對應本機 GAS 來源：

`D:\program\LKC\新家人管理系統`

## 現況

- 前端位置：`apps/LKC_NewFamily`
- GAS 後端：獨立 Google Apps Script
- 目標試算表：`1ZSixQ9-T8_sNkYviTfP_GXgVk_8k_Lgs6YWStFRtFK8`
- 認證 token：由 `api-config.js` 的 `NEW_FAMILY_AUTH_TOKEN` 帶入
- 主日會友查核：透過主日點名 GAS / 主日會友名單同步「會友名單狀態」、「點名系統代碼」、「主日點名小組」

## 功能

- 新家人資料送出後寫入「追蹤中」分頁。
- 同工可查詢追蹤中案件，依姓名與日期篩選。
- 可編輯追蹤案件欄位。
- 可同步既有主日會友代碼與小組資訊。
- 勾選案件後可結案，資料移至「已結案」分頁。
- 已結案案件可查詢與做分析匯出。

## API actions

| Action | 用途 |
|---|---|
| `submitNewFamily` | 新增新家人追蹤案件 |
| `getTrackingCases` | 讀取追蹤中案件 |
| `getClosedCases` | 讀取已結案案件 |
| `updateTrackingCase` | 更新追蹤中案件 |
| `markTrackingMemberStatuses` | 標記會友名單狀態、點名系統代碼與主日小組 |
| `syncExistingMemberCodes` | 從主日會友名單同步既有代碼 |
| `closeCases` | 將追蹤中案件移到已結案 |

## Firebase cache

GAS 端會將完整清單寫入 Firebase RTDB：

- `cache/getTrackingCases/_default`
- `cache/getClosedCases/_default`

快取 TTL 為 5.5 小時；`setupNewFamilyCacheTriggers()` 會建立：

- `keepWarmNewFamilyCaches`：每 4 小時刷新完整清單。
- `onEditNewFamilySheet`：當「追蹤中」或「已結案」被手動編輯時刷新清單。

新增、編輯、同步會友狀態與結案時，GAS 也會清除舊 cache 並重建完整清單。

## 部署注意

1. 本資料夾只放 GitHub Pages 前端。
2. GAS 原始碼在 `D:\program\LKC\新家人管理系統`，更新後需另外 `clasp push`。
3. Web App URL 設定在 `api-config.js` 的 `NEW_FAMILY_API_URL`。
4. `FIREBASE_SERVICE_ACCOUNT` 必須放在 GAS Script Properties，不要提交 service account JSON。
