# 台語禮拜 PPT 編輯器

`index.html` 是獨立的主日投影製作工作台。畫面依禮拜流程排列，左側切換流程項目，中間編輯內容，右側預覽 16:9 投影片。

## 目前資料流

```text
使用者輸入 → app.js model → 即時預覽
                    └→ localStorage 草稿
行事曆帶入 → calendar-integration.js → config.js / churchAPI
                                      → LKC_MasterSchedule GAS `cal_getEvents`
                                      → calendar-adapter.js → app.js model
                                      → GAS `cal_getPptLibraryIndex`
                                      → GAS `cal_getPptLibraryFile`
                                      → Google Drive 聖詩／啟應文 PPTX Base64
                                      → pptx-library.js → 分層投影片物件
```

- 固定禮文：使徒信經與主禱文各以單一全文欄位編輯，再依原 PPT 三頁的文字量比例與空白行即時重排；奉獻詩與阿們頌以預設內容建立。
- 週報雲端內容：依禮拜日期向週報 GAS 讀取 `reports_YYYY-MM-DD` 與 `praise_songs_YYYY-MM-DD`。報告依來源 PPT 分為「本會消息」（每頁兩則）與「關懷代禱」，並保留「報告」標題頁；讚美帶入詩名、演唱者與歌詞，歌詞仍以空白行分頁。兩者載入後都可手動修改；雲端無資料時不清除現有內容。
- 行事曆：以日期呼叫 `cal_getEvents`，精準選取 `講道資訊 - 台語`。行事曆值先保存為 `sourceValue`，不直接視為投影片內文。
- 雲端資料庫：聖詩以 `第{編號}首 {名稱}.pptx`、啟應文以 `{編號}.pptx` 配對。GAS 先列出 Drive 檔案索引，再由 `cal_getPptLibraryFile` 只讀回傳索引內 PPTX 的 Base64；瀏覽器在記憶體內還原並解析，避免 GitHub Pages 直接下載 Drive 檔案遭瀏覽器阻擋。
- 講道 PPT 可選擇檔案，但尚未在第一版讀取或合併。
- PPTX 匯出按鈕暫為端口。後續應重用 `apps/LKC_ppt_generator` 的 PptxGenJS 輸出方式。

## 行事曆欄位契約

`calendar-adapter.js` 接受行事曆事件的 `values[]`，以 `fieldName` 對應欄位。聖詩欄位可使用 `聖詩第一首`／`聖詩一`、`聖詩第二首`／`聖詩二`。本工具設定 `_GAS_KEY = 'LKC_MasterSchedule'`，沿用既有 Router，使用唯讀 action `cal_getPptLibraryIndex` 與 `cal_getPptLibraryFile`；不新增寫入行為。

### 台語講道資料的選取與經文產生器

- 查詢固定使用禮拜日期的同一天：`{ startDate, endDate }`。
- 只接受 `typeName === "台語"` 且 `typeFullName === "講道資訊 - 台語"` 的事項，不再以標題模糊比對或退回同日第一筆事項。
- 正式欄位包含：`講題`、`講員`、`經文`、`宣召`、`金句`、`聖詩一`、`啟應文`、`聖詩二`、`頌榮`。
- `宣召`、`經文`、`金句` 的 `sourceValue` 交給共用 `LKC_ppt_generator/bible-service.js`，以 `tghg` 取得台語經文全文，再由 `slide-production.js` 依每頁兩節建立投影片。
- `聖詩一`、`聖詩二`、`頌榮` 與 `啟應文` 的 `sourceValue` 只作資料庫索引。`ppt-library-integration.js` 以編號配對 Drive 索引，再由 `pptx-library.js` 在瀏覽器解壓縮 OOXML。
- PPTX 不會整頁轉圖。解析結果保留三層：全份共用背景圖、聖詩頁專用白色色塊、原 PPT 內嵌譜面圖片與文字方塊。圖片與文字都保留來源座標；白色色塊提供 40–80% 透明度，文字與譜面仍可依「標題／內文」版面群組調色及縮放。
- 所有聖詩相關段落（會前／正式聖詩、261、306B、頌榮、522）的白色色塊透明度預設同步調整；上方工具列提供預設勾選的「聖詩頁白底透明度同步」。取消勾選後只更新目前段落，其他聖詩保留各自透明度。
- `vendor-jszip.min.js` 固定隨 app 提供，避免外部 CDN 未載入時阻塞整個編輯器。
- `read-api.js` 正常情況沿用 `churchAPI` POST；若頁面由 `file://` 開啟，或 POST 因瀏覽器跨來源政策失敗，則改用同一 GAS 網址的唯讀 JSONP。JSONP 僅開放 `cal_getEvents`、`cal_getPptLibraryIndex`、`cal_getPptLibraryFile`、`cal_queryBible`，不提供任何寫入 action。`content-generators.js` 先在瀏覽器解析經文範圍，再由 GAS 伺服器查詢台語聖經，避免 `file://` 直接 fetch 信望愛 API。
- `layout-groups.js` 將所有流程攤平成一份連續 PPT，並直接把 PowerPoint 式「章節」整合到左側禮拜流程：流程編號左側可勾選整章，展開後以子項目顯示並勾選個別頁面。整個工具只使用右側既有的 16:9 共用預覽畫布，不再建立第二個版面畫布。
- 勾選章節或頁面只更新選取範圍；使用者由頂端「版面參數」按鈕開啟小型橫向懸浮面板。面板以「標題／內文」切換調整目前共用畫布，開啟後常駐至手動關閉，並可拖曳避開畫布。儲存後將同一組參數套用到所有勾選頁並立即寫入瀏覽器草稿。
- 背景模板由全份 PPT 共用：可使用純色 `backgroundColor`，或從頂端按鈕上傳 PNG／JPG／WebP 背景圖。背景圖以 16:9 畫布滿版裁切並可隨草稿保存；背景層不放置固定頁首或教會名稱。版面參數群組分別保存 `titleColor` 與 `contentColor`，讓真正的標題、內文可獨立選色。
- 開啟懸浮面板時，起始參數由目前共用畫布的實際 computed layout 換算，包含字級、X/Y、寬高、對齊與行距；因此不同投影片模板不再共用一組與畫布不符的固定起始值。
- `ppt-format-preview.js` 允許在非輸入欄位使用左右／上下方向鍵依整份 PPT 順序跨章切換；輸入欄位內的方向鍵仍保留游標操作。
- 版面群組與頁面歸屬存入 `localStorage` 草稿；解除群組後，該頁可保留獨立 `page.layout` 微調值。
- 本 app 設定 `window._FORCE_PRODUCTION_GAS = true`，因此 localhost 也讀正式 `LKC_MasterSchedule`；其他 app 仍保留原有 localhost 自動切換 `_TEST` 的行為。

## 2026-07-14 與來源 PPT 對齊

- 流程保留來源 PPT 的會前段落：首頁、`會前領唱`、第一首聖詩、第二首聖詩，再進入第二張禮拜首頁與 `靜默一分鐘`。
- 行事曆的兩個聖詩編號會同時供應會前與正式禮拜的對應聖詩；相同 Drive 檔案依 file ID 共用解析快取。
- 固定從 Drive 載入主禱文後的第 `261` 首（副歌）、奉獻第 `306B` 首、阿們頌第 `522` 首。只有 `306B` 另有奉獻標題頁，對齊來源 PPT。
- PPTX 解析支援 PowerPoint 主題色 `schemeClr` 與母片色彩對應，因此啟應文的「啓／應」文字顏色不會被統一成黑色。
- 聖詩、奉獻、頌榮與報告依來源 PPT 明確保留各自應有的標題頁。
