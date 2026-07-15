# 台語禮拜 PPT 共用版面 Firebase 設定

版面群組的雲端來源為 Realtime Database：

```text
worshipPpt/layoutConfig/shared
```

瀏覽器的 `localStorage` 只作首次遷移來源與離線備份。背景圖片及每週內容不會寫進此 RTDB 節點。

## 1. 建立版面編輯帳號

1. Firebase Console → Authentication → Sign-in method。
2. 啟用 Email/Password provider。
3. 在 Users 建立帳號 `worship-layout@lkc1958.org`。
4. 密碼使用專案負責人約定的版面解鎖密碼；不要把密碼寫入 Git、JavaScript 或本文件。

編輯器登入採 `inMemoryPersistence`，所以重新整理頁面後會再次鎖定。

## 2. 合併 Realtime Database Rules

`database.rules.worship-layout.json` 是 `worshipPpt` 節點的規則範本。請把其中的 `worshipPpt` 節點合併到 Firebase Console 目前正在使用的完整規則，保留既有的 `cache` 與其他正式節點，切勿直接覆蓋未備份的正式規則。

規則效果：

- 所有人可讀取共用版面，因此投影工作站不必先登入。
- 只有 Firebase Auth email 為 `worship-layout@lkc1958.org` 的使用者可以寫入。
- 寫入內容限制為 schema version、layout state、更新時間與更新者 UID。

## 3. 首次遷移

1. 使用原本保存版面參數的瀏覽器開啟編輯器。
2. 等待畫面顯示已載入本機版面備份。
3. 按「版面設定已鎖定」，輸入約定密碼。
4. 若雲端尚無配置，現有的 `layoutState.groups` 與 `layoutState.pageAssignments` 會自動寫入共用節點。
5. 用另一個瀏覽器重新開啟頁面，確認顯示「已載入全教會共用版面配置」。

## 4. 回復策略

雲端讀取失敗時，編輯器會使用最後一次成功同步到 `localStorage` 的版面備份。雲端寫入失敗時，畫面會明確顯示失敗，不會宣稱已完成同步。
