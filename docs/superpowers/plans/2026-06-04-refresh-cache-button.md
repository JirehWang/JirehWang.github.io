# 刷新快取懸浮按鈕 實現計劃

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:executing-plans 逐任務實現此計劃。步驟使用複選框（`- [ ]`）語法來跟蹤進度。

**目標：** 在主門戶首頁 (`index.html`) 新增一個右下角懸浮按鈕 (FAB)，點擊後一鍵清除瀏覽器的 Service Worker 快取，並重新載入頁面。

**建築：** 
1. 在 `index.html` 中新增一個固定定位的懸浮按鈕，採用品牌漸層色。
2. 點擊後，以 CSS 動畫持續旋轉 SVG 圖示，並彈出玻璃擬態 Toast 提示。
3. 利用 JavaScript 的 `caches` API 及 `navigator.serviceWorker` API 清理快取並登出，隨後執行硬性重新整理。

**技術棧：** Vanilla HTML, CSS, JavaScript (Browser APIs: Caches, ServiceWorker)

---

## 1. 變更文件結構與職責

* **[MODIFY] [index.html](file:///d:/program/Github/LKC1958_June_1.github.io/index.html)**:
  - 在 `<body>` 中新增懸浮按鈕與 Toast 的 HTML 結構。
  - 在 `<style>` 中新增按鈕、動畫及 Toast 的 CSS 樣式。
  - 在 `<script>` 中新增快取清理與重新載入的 JS 互動邏輯。

---

## 2. 任務清單

### 任務 1：在 `index.html` 新增按鈕與 Toast 結構及 CSS 樣式

**文件：**
- 修改：`index.html`

- [ ] **步驟 1：在 `index.html` 底部新增按鈕與 Toast 結構**

在 `index.html` 的 `<footer>` 之後、`</body>` 之前加入以下 HTML：

```html
<!-- 刷新快取懸浮按鈕 (FAB) -->
<button id="btn-refresh-cache" class="fab-refresh" title="清除快取並強制更新" aria-label="清除快取並強制更新">
    <svg class="refresh-icon" viewBox="0 0 24 24" width="24" height="24">
        <path d="M17.65 6.35A7.958 7.958 0 0 0 12 4c-4.42 0-7.99 3.58-7.99 8s3.57 8 7.99 8c3.73 0 6.84-2.55 7.73-6h-2.08A5.99 5.99 0 0 1 12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6c1.66 0 3.14.69 4.22 1.78L13 11h7V4l-2.35 2.35z"/>
    </svg>
</button>

<!-- 懸浮提示 Toast -->
<div id="cache-toast" class="cache-toast">
    🔄 正在清除系統快取並重新載入...
</div>
```

- [ ] **步驟 2：在 `index.html` 的 `<style>` 中加入 CSS 樣式**

將以下 CSS 加入至 `index.html` 中的 `<style>` 區塊：

```css
/* 懸浮按鈕樣式 */
.fab-refresh {
    position: fixed;
    bottom: 24px;
    right: 24px;
    width: 50px;
    height: 50px;
    border-radius: 50%;
    background: linear-gradient(135deg, #006030 0%, #30759f 100%);
    border: none;
    cursor: pointer;
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.2);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 1000;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
    outline: none;
}

.fab-refresh:hover {
    transform: scale(1.08);
    box-shadow: 0 6px 20px rgba(0, 0, 0, 0.25);
}

.fab-refresh:active {
    transform: scale(0.95);
}

.fab-refresh[disabled] {
    cursor: not-allowed;
    opacity: 0.8;
}

.refresh-icon {
    fill: white;
    transition: transform 0.4s cubic-bezier(0.4, 0, 0.2, 1);
}

.fab-refresh:hover .refresh-icon {
    transform: rotate(180deg);
}

/* 旋轉動畫 */
@keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}

.refresh-icon.spinning {
    animation: spin 1s linear infinite !important;
}

/* Tooltip 樣式 */
.fab-refresh::after {
    content: attr(title);
    position: absolute;
    bottom: 60px;
    right: 0;
    background: rgba(45, 55, 72, 0.95);
    color: white;
    padding: 6px 12px;
    border-radius: 6px;
    font-size: 12px;
    white-space: nowrap;
    opacity: 0;
    visibility: hidden;
    transition: opacity 0.2s ease, visibility 0.2s ease;
    box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    pointer-events: none;
    font-family: inherit;
}

.fab-refresh:hover::after {
    opacity: 1;
    visibility: visible;
}

/* Glassmorphism Toast 提示框 */
.cache-toast {
    position: fixed;
    bottom: 90px;
    right: 24px;
    background: rgba(45, 55, 72, 0.9);
    backdrop-filter: blur(8px);
    -webkit-backdrop-filter: blur(8px);
    color: white;
    padding: 12px 20px;
    border-radius: 8px;
    font-size: 14px;
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.2);
    z-index: 1000;
    opacity: 0;
    transform: translateY(10px);
    transition: opacity 0.3s ease, transform 0.3s ease;
    pointer-events: none;
    font-family: inherit;
    display: flex;
    align-items: center;
    gap: 8px;
}

.cache-toast.show {
    opacity: 1;
    transform: translateY(0);
}
```

---

### 任務 2：新增清除快取與重新載入的 JS 互動邏輯

**文件：**
- 修改：`index.html`

- [ ] **步驟 1：在 `index.html` 的 `<script>` 中加入 JS 互動邏輯**

將以下 JavaScript 加入至 `index.html` 中的 `<script>` 區塊：

```javascript
    // 取得懸浮按鈕、圖示與 Toast 元素
    const btnRefresh = document.getElementById('btn-refresh-cache');
    const refreshIcon = btnRefresh ? btnRefresh.querySelector('.refresh-icon') : null;
    const cacheToast = document.getElementById('cache-toast');

    if (btnRefresh) {
        btnRefresh.addEventListener('click', async function() {
            // 1. 停用按鈕並加入旋轉動畫
            btnRefresh.disabled = true;
            if (refreshIcon) {
                refreshIcon.classList.add('spinning');
            }

            // 2. 顯示 Toast 提示
            if (cacheToast) {
                cacheToast.classList.add('show');
            }

            // 3. 清理快取邏輯
            try {
                // 清除 caches API 儲存的快取
                if ('caches' in window) {
                    const cacheKeys = await caches.keys();
                    await Promise.all(
                        cacheKeys.map(key => caches.delete(key))
                    );
                    console.log('✅ ServiceWorker cache storage cleared.');
                }

                // 登出所有已註冊的 Service Workers
                if ('serviceWorker' in navigator) {
                    const registrations = await navigator.serviceWorker.getRegistrations();
                    await Promise.all(
                        registrations.map(reg => reg.unregister())
                    );
                    console.log('✅ ServiceWorkers unregistered.');
                }
            } catch (err) {
                console.warn('❌ Cache cleanup encountered an error:', err);
            }

            // 4. 延遲 1 秒以提供視覺回饋後重新載入
            setTimeout(() => {
                window.location.reload(true);
            }, 1000);
        });
    }
```

- [ ] **步驟 2：驗證 HTML 的變更並進行 Git Commit**

```bash
git diff index.html
git add index.html
git commit -m "feat: add refresh cache FAB to portal index.html"
```
