# 台語禮拜 PPT Firebase 內容同步契約

前端已改為 Firebase-first；節點不存在或讀取失敗時，才回退既有 GAS。這份變更不修改或部署 GAS。

## RTDB 路徑

所有內容位於 `worshipPpt/content`：

- `services/{YYYY-MM-DD}/calendar`：`cal_getEvents` 的完整回應，例如 `{ "success": true, "data": [...] }`
- `services/{YYYY-MM-DD}/reports`：週報 `reports_YYYY-MM-DD` 的資料物件
- `services/{YYYY-MM-DD}/praise`：讚美 `praise_songs_YYYY-MM-DD` 的資料物件
- `library/index`：`cal_getPptLibraryIndex` 的完整回應，例如 `{ "success": true, "data": [...] }`
- `bible/{version}/{book}/{chapter}/{verses}`：`cal_queryBible` 的完整回應，例如 `{ "success": true, "records": [...] }`

`verses` 使用查詢值（例如 `1-2`）；整章使用 `_all`。Firebase key 不允許的 `. # $ / [ ]` 與控制字元須替換為 `_`。

## Firebase Storage

PPTX 二進位檔放在 `worshipPpt/library/`，不要以 Base64 存入 RTDB。`library/index` 每筆資料保留既有 `fileId`，並增加：

```json
{
  "downloadUrl": "https://firebasestorage.googleapis.com/..."
}
```

前端看到 `downloadUrl` 或 `storageUrl` 時會直接讀 Firebase Storage；沒有時才回退 `cal_getPptLibraryFile`。

## 權限與部署

- RTDB `worshipPpt/content`：公開唯讀、瀏覽器禁止寫入。
- Storage `worshipPpt/library/**`：公開唯讀、瀏覽器禁止寫入。
- 日後 GAS 應以既有 Service Account／管理權限同步，管理身分不依賴瀏覽器規則取得寫入權。
- 部署規則時必須把範本合併進目前正式規則，不可覆蓋既有 `cache`、`logs` 等節點。

## 目前實際盤點（2026-07-15）

- 已有：`worshipPpt/layoutConfig/shared` 的版面群組與頁面歸屬。
- 尚無：當日行事曆、PPT Library index、PPTX Storage URL、聖經查詢、報告、讚美內容。
- `cache/cal_getEvents` 只有其他日期的短期快取，2026-07-15 沒有命中；它不能取代穩定同步資料。
- 正式 layoutState 尚未看到 `hymnOpacityBySection` 與 `outputScale`；規則範本已補上這兩種欄位的驗證。
