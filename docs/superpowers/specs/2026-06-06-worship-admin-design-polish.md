# 敬拜團服事管理系統 — 網頁美編優化規格說明書 (Modern Glassmorphic)

本文件定義「敬拜團服事管理系統管理後台」的視覺美化設計系統與規格。我們將在維持原有 **LKC 綠 (#006030)** 與 **Portal 藍 (#30759f)** 色系的基礎上，引入「現代微光摩登風」的視覺美學。

## 設計核心目標 (Design Summary)
1. **毛玻璃與層次感 (Glassmorphism & Depth)**：卡片元件導入半透明背景、微邊框與大圓角，搭配模糊背板，大幅提升介面層次。
2. **現代化字型 (Typography Upgrade)**：引進 Google Fonts (Inter + Noto Sans TC)，取代微軟正黑體，解決預設字體粗細不均的問題。
3. **優雅的微互動 (Micro-interactions)**：對按鈕與互動式卡片添加精緻的 Hover 浮空與陰影漸變效果。
4. **表格與操作介面精緻化 (Table & UI Polish)**：優化表格的凍結列陰影，使滾動時的視覺層次更加清晰，並將系統按鈕統一為現代漸層圓角按鈕。

---

## 視覺規範 (Visual System)

### 1. 色彩系統 (Color Tokens)
保持原有品牌色彩，但透過透明度與漸層進行豐富：
- **主品牌綠 (LKC Green)**: `#006030`
- **品牌藍 (Portal Blue)**: `#30759f`
- **主漸層色 (Active Gradient)**: `linear-gradient(135deg, #006030 0%, #30759f 100%)`
- **頁面背景色 (Page Background)**: `#f4f7f6`
- **卡片背景色 (Card Background)**: `rgba(255, 255, 255, 0.85)` (backdrop-filter: blur(10px))
- **微光邊框色 (Subtle Border)**: `rgba(0, 96, 48, 0.08)`
- **陰影規範 (Shadow)**: `0 8px 32px 0 rgba(0, 96, 48, 0.04)`

### 2. 字型系統 (Typography)
- **字體匯入**: `@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Noto+Sans+TC:wght@400;500;700&display=swap');`
- **字型套用**: `font-family: 'Inter', 'Noto Sans TC', -apple-system, sans-serif;`

### 3. 卡片與區塊 (Cards & Containers)
- 所有資訊容器 (`.container`, `.card`, `.tab-content`) 統一採用：
  - `border-radius: 16px`
  - `border: 1px solid rgba(0, 96, 48, 0.08)`
  - `box-shadow: 0 8px 32px 0 rgba(0, 96, 48, 0.04)`

### 4. 導覽列與標籤頁 (Nav & Tab Pills)
- 導覽列背景改用極淡的藍綠半透明背景：`rgba(229, 236, 233, 0.6)`。
- 行動端優化圓角，活動頁籤套用 `linear-gradient(135deg, #006030 0%, #30759f 100%)`，並附帶微投影。

---

## 具體實施細節 (Proposed Changes)

### 1. `style.css` 優化
- 匯入 Google Fonts。
- 更新 `body` 與字體設定。
- 更新 `.container` 與 `.card` 的毛玻璃樣式。
- 優化 `.modern-table` 與 `.bulletin-table`：
  - 凍結的第一、第二欄右側加上細緻的陰影，突顯滾動時的黏滯感。
  - Header 的背景維持綠色，但加入微光亮條。
- 重構按鈕樣式，加入 `.btn-primary` 的漸層效果與 hover 浮空動畫。

### 2. `admin.html` 結構優化
- 在 `<head>` 中引入 Google Fonts 連結。
- 調整部分 Bootstrap 的排版層次，移除硬性邊框 style，改用 CSS 設計系統控制。
