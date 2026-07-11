# Firebase Firestore 快取整合

## 快取併發與失效策略（2026-07）

- `cacheGetOrFetch` 與 `cacheGetOrFetchWithMeta` 在同一頁面內會共用相同 `topic/subkey` 的進行中請求，避免元件同時初始化時重複讀取 GAS。
- 進行中的 Promise 無論成功或失敗都會移除，因此後續重新整理仍會正常執行。
- 管理入口先清 RTDB topics，再呼叫 GAS 重建後端快取；GAS 批次 invalidation 使用一次 root PATCH，失敗時退回逐筆刪除。

把 GAS 回傳的 JSON 暫存到 Firestore，降低呼叫頻率、加快讀取。

## 一、初次設定（你只需要做一次）

### 1. 在 Firebase Console 取得設定值

1. 打開 https://console.firebase.google.com/ 進入你的專案
2. 左上角齒輪 ⚙️ → **專案設定**
3. 「一般設定」分頁 → 滑到最下方「您的應用程式」
4. 如果還沒有 Web App：點 **`</>`** 圖示新增一個（不需要勾 Firebase Hosting）
5. 找到 `firebaseConfig`，會長這樣：

   ```js
   const firebaseConfig = {
     apiKey: "AIzaSy...",
     authDomain: "your-project.firebaseapp.com",
     projectId: "your-project",
     storageBucket: "your-project.appspot.com",
     messagingSenderId: "1234567890",
     appId: "1:1234567890:web:abcdef"
   };
   ```

6. 把這六個欄位的值複製到 `firebase/firebase-config.js`，覆蓋 `TODO_*` 的部分

> ⚠️ Web 端的 `apiKey` 不是機密金鑰，可以公開。真正的權限控制要設定下面的 Firestore Security Rules。

### 2. 建立 Firestore 資料庫

1. Firebase Console 左側 → **建構** → **Firestore Database**
2. 點「建立資料庫」
3. 第一次選 **以測試模式啟動**（30 天內可任意讀寫）
4. 區域選擇 `asia-east1`（台灣）或 `asia-northeast1`（東京）

### 3. （上線前）調整 Security Rules

測試模式 30 天後會自動鎖住。上線前到 Firestore → **規則** 分頁修改，例如只允許讀、寫入由 GAS 透過 Admin SDK：

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /cache/{key} {
      allow read: if true;          // 讀取公開
      allow write: if false;        // 前端不可直接寫
    }
  }
}
```

## 二、在程式中使用

```html
<script type="module">
  import { cacheGet, cacheSet, cacheGetOrFetch }
    from '../firebase/firebase-cache.js';

  // 寫入快取（TTL 300 秒）
  await cacheSet('songs-2026', { list: [...] }, 300);

  // 讀取快取（過期或不存在會回傳 null）
  const data = await cacheGet('songs-2026');

  // 最常用：沒命中才呼叫 GAS
  const songs = await cacheGetOrFetch(
    'worship-songs-2026',
    () => window.churchAPI('getSongs'),
    600   // 快取 10 分鐘
  );
</script>
```

## 三、測試

打開 `firebase/example.html` 點按鈕：
- **寫入快取** → Firestore 會新增 `cache/demo-key` 文件
- **讀取快取** → 控制台 / 頁面顯示資料
- **刪除快取** → 文件被移除

到 Firebase Console 的 Firestore → **資料** 分頁可以看到實際儲存的內容。

## 檔案說明

| 檔案 | 用途 |
|------|------|
| `firebase-config.js` | 初始化 Firebase App + Firestore（**請填入你的設定值**） |
| `firebase-cache.js` | 快取 API：`cacheGet` / `cacheSet` / `cacheDelete` / `cacheGetOrFetch` |
| `example.html` | 手動測試頁，含三個按鈕 |
