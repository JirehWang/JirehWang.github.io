# Firebase Realtime Database 快取整合

把 GAS 回傳的 JSON 暫存到 Firebase Realtime Database (RTDB),降低呼叫頻率、加快讀取。

## 一、初次設定

### ✅ 已完成
- Web App 已註冊
- `firebase/firebase-config.js` 已填入你的設定值
- 專案:`lkc1958june1`,RTDB 區域:`asia-southeast1` (新加坡)

### 你還要做的:確認 Realtime Database 已啟用

1. 打開 Firebase Console → 進入 `lkc1958june1` 專案
2. 左側 **建構** → **Realtime Database**
3. 如果還沒建立,點「建立資料庫」→ 選 `asia-southeast1` → **以測試模式啟動**
4. 如果看得到「資料」分頁,代表已經 OK ✅

### (上線前) 調整 Security Rules

測試模式 30 天後會自動鎖住。到 RTDB → **規則** 分頁,建議設定:

```json
{
  "rules": {
    "cache": {
      ".read": true,
      ".write": true,
      "$key": {
        ".validate": "newData.hasChildren(['value','expiresAt','updatedAt'])"
      }
    }
  }
}
```

> 上線後若要更嚴格,可把 `.write` 改成 `false`,改由 GAS 透過 Admin SDK 寫入。

## 二、在程式中使用

```html
<script type="module">
  import { cacheGet, cacheSet, cacheGetOrFetch }
    from '../firebase/firebase-cache.js';

  // 寫入快取 (TTL 300 秒)
  await cacheSet('songs-2026', { list: [...] }, 300);

  // 讀取快取 (過期或不存在會回傳 null)
  const data = await cacheGet('songs-2026');

  // 最常用:沒命中才呼叫 GAS
  const songs = await cacheGetOrFetch(
    'worship-songs-2026',
    () => window.churchAPI('getSongs'),
    600   // 快取 10 分鐘
  );
</script>
```

## 三、測試

打開 `firebase/example.html` 點按鈕:
- **寫入快取** → RTDB 的 `/cache/demo-key` 會新增一筆
- **讀取快取** → 頁面顯示資料內容
- **刪除快取** → 節點被移除

到 Firebase Console → Realtime Database → **資料** 分頁可以即時看到變化(RTDB 比 Firestore 即時)。

## 檔案說明

| 檔案 | 用途 |
|------|------|
| `firebase-config.js` | 初始化 Firebase App + RTDB |
| `firebase-cache.js` | 快取 API:`cacheGet` / `cacheSet` / `cacheDelete` / `cacheGetOrFetch` |
| `example.html` | 手動測試頁,含三個按鈕 |

## 為什麼選 RTDB 而不是 Firestore?

| | Realtime Database | Firestore |
|---|---|---|
| 資料模型 | 一棵 JSON 樹 | 文件/集合 |
| 適合 | 簡單 JSON 快取、即時同步 | 複雜查詢、大規模資料 |
| 免費額度 | 1GB 儲存 + 10GB/月傳輸 | 1GiB 儲存 + 50K 讀/20K 寫/天 |

你的用途是「JSON 快取存取點」→ RTDB 完美對應。
