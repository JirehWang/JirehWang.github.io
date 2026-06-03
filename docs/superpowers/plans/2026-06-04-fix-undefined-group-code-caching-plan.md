# 修復小組代碼 undefined 快取問題 實作計劃

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計劃。步驟使用複選框（`- [ ]`）語法來跟踪進度。

**目標：** 移除 `config.js` 中登入和驗證 API 的快取，清除 Firebase 中的過期/錯誤快取，從而徹底解決 `res.encryptedCode` 為 `undefined` 導致小組頁面報錯的問題。

**架構：** 
1. 編輯 `config.js`，將 `findGroupByCode` 與 `verifyGroup` 移出快取名單 `_CACHEABLE_ACTIONS`。
2. 執行一次性 Python 腳本，透過 Firebase Realtime Database REST API 刪除 `cache/findGroupByCode` 和 `cache/verifyGroup` 下的所有 stale 快取。
3. 透過自動化腳本驗證 `config.js` 修改是否生效。

**技術棧：** Vanilla JavaScript, Firebase REST API, Python (用於快取清理與驗證)

---

### 任務 1：修改 `config.js` 移出登入快取

**文件：**
- 修改：`Github/LKC1958_June_1.github.io/config.js`
- 驗證：`Github/LKC1958_June_1.github.io/config.js` 的語法與變更。

- [ ] **步驟 1：修改 `config.js` 內容**
  移除 `_CACHEABLE_ACTIONS` 物件中的 `'findGroupByCode'` 與 `'verifyGroup'` 兩行。

  修改對比：
  ```diff
  -    'findGroupByCode':              _SIX_HOURS,  // 小組代碼 → 名稱（登入）
  -    'verifyGroup':                  _SIX_HOURS,  // 小組名 + 代碼驗證（登入）
  ```

- [ ] **步驟 2：執行靜態檢查**
  確保 `config.js` 無 JavaScript 語法錯誤。可在終端運行 node 執行它（或僅檢查語法）。
  運行：`node -c Github/LKC1958_June_1.github.io/config.js`
  預期：無錯誤輸出。

- [ ] **步驟 3：Commit 變更**
  ```bash
  git add config.js
  git commit -m "fix: remove findGroupByCode and verifyGroup from cached actions in config.js"
  ```

---

### 任務 2：清理 Firebase 中已存在的 Stale 快取

**文件：**
- 創建：`D:\program\purge_stale_cache.py` (一次性腳本)

- [ ] **步驟 1：編寫清理腳本**
  創建並編寫 `purge_stale_cache.py`，透過 REST API 發送 DELETE 請求清除 `findGroupByCode` 和 `verifyGroup` 快取節點。
  ```python
  import urllib.request
  import sys

  def purge_cache(topic):
      url = f"https://lkc1958june1-default-rtdb.asia-southeast1.firebasedatabase.app/cache/{topic}.json"
      req = urllib.request.Request(url, method='DELETE')
      try:
          with urllib.request.urlopen(req) as response:
              print(f"Successfully purged cache for: {topic} (Status: {response.status})")
      except Exception as e:
          print(f"Error purging cache for {topic}: {e}")

  if __name__ == '__main__':
      purge_cache("findGroupByCode")
      purge_cache("verifyGroup")
  ```

- [ ] **步驟 2：執行清理腳本**
  運行：`.venv\Scripts\python purge_stale_cache.py`
  預期：輸出 `Successfully purged cache for: findGroupByCode` 與 `Successfully purged cache for: verifyGroup`。

- [ ] **步驟 3：驗證快取是否被清空**
  執行：`.venv\Scripts\python inspect_rtdb_cache.py`
  預期：輸出為：
  ```
  === No cache found for findGroupByCode ===
  === No cache found for verifyGroup ===
  ```

---

### 任務 3：完成前自查與 Git 提交

**文件：**
- 刪除：`purge_stale_cache.py`
- 驗證：`git status` 及 `git diff`。

- [ ] **步驟 1：清理一次性腳本**
  刪除本地的 `purge_stale_cache.py` 及 `inspect_rtdb_cache.py`。
  運行：`Remove-Item -Path "purge_stale_cache.py", "inspect_rtdb_cache.py" -Force`

- [ ] **步驟 2：最終語法與狀態檢查**
  運行：`git status`
  確認僅有 `config.js` 和本實作計畫/設計文件的變更。
