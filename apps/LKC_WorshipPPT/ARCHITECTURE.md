# 禮拜PPT產生器

`index.html` 是獨立的禮拜投影製作工作台。產品名稱與路徑不綁定單一語言；目前第一個內建流程是台語主日禮拜。畫面依所選禮拜流程排列，左側切換流程項目，中間編輯內容，右側預覽 16:9 投影片。

## 目前資料流

```text
使用者輸入 → app.js model → 即時預覽
                    ├→ localStorage 內容草稿／版面離線備份
                    └→ Firebase RTDB 全教會共用版面配置
行事曆帶入 → calendar-integration.js → config.js / churchAPI
                                      → LKC_MasterSchedule GAS `cal_getEvents`
                                      → calendar-adapter.js → app.js model
                                      → GAS `cal_getPptLibraryIndex`
                                      → GAS `cal_getPptLibraryFile`
                                      → Google Drive 聖詩／啟應文 PPTX Base64
                                      → pptx-library.js → 分層物件解析 → 樂譜／啟應文透明 PNG
```

- 固定禮文：使徒信經與主禱文各以單一全文欄位編輯，再依原 PPT 三頁的文字量比例與空白行即時重排；奉獻詩與阿們頌以預設內容建立。
- 週報雲端內容：依禮拜日期向週報 GAS 讀取 `reports_YYYY-MM-DD` 與 `praise_songs_YYYY-MM-DD`。報告依序分為「本會消息」（每頁兩則）、「教界消息」（每頁兩則）與「關懷代禱」，並保留「報告」標題頁；三類報告都取自 Sunday Bulletin `reports.html` 儲存的同一筆資料。讚美帶入詩名、演唱者與歌詞，歌詞仍以空白行分頁。兩者載入後都可手動修改；雲端無資料時不清除現有內容。
- 行事曆：以日期呼叫 `cal_getEvents`，精準選取 `講道資訊 - 台語`。行事曆值先保存為 `sourceValue`，不直接視為投影片內文。
- 雲端資料庫：聖詩以 `第{編號}首 {名稱}.pptx`、啟應文以 `{編號}.pptx` 配對。GAS 先列出 Drive 檔案索引；只有真正的 Firebase Storage／Google Cloud Storage URL 才由瀏覽器直接下載，Drive `downloadUrl` 則固定改走 `cal_getPptLibraryFile`，由 GAS 只讀回傳索引內 PPTX 的 Base64，避免瀏覽器 CORS 阻擋。瀏覽器在記憶體內還原並解析 OOXML；只有樂譜與啟應文頁會點陣化成透明整頁 PNG，避免不同 PowerPoint 環境重新排字。首頁、標題、報告、禮文、經文、讚美與其餘產生頁在匯出檔中仍是 PowerPoint 原生文字物件。
- 樂譜白底透明度按流程段落存入共用 `layoutState.hymnOpacityBySection`，與版面群組共用 Firebase Auth 解鎖權限；鎖定時不可調整，變更後寫入全教會共用 RTDB 文件。
- PPTX 使用原始素材相同的寬螢幕尺寸 `13.333 × 7.5`，避免 10 吋畫布造成字級相對放大。
- 講道 PPT 可選擇檔案，但尚未在第一版讀取或合併。
- `ppt-export.js` 使用 PptxGenJS 匯出完整 PPTX；畫布與匯出端共同讀取 `slide-production.js` 的預設版面、使用者版面群組、文字斷行與輸出比例。

## 行事曆欄位契約

`calendar-adapter.js` 接受行事曆事件的 `values[]`，以 `fieldName` 對應欄位。聖詩欄位可使用 `聖詩第一首`／`聖詩一`、`聖詩第二首`／`聖詩二`。本工具設定 `_GAS_KEY = 'LKC_MasterSchedule'`，沿用既有 Router，使用唯讀 action `cal_getPptLibraryIndex` 與 `cal_getPptLibraryFile`；不新增寫入行為。

### 台語講道資料的選取與經文產生器

- 查詢固定使用禮拜日期的同一天：`{ startDate, endDate }`。
- 只接受 `typeName === "台語"` 且 `typeFullName === "講道資訊 - 台語"` 的事項，不再以標題模糊比對或退回同日第一筆事項。
- 正式欄位包含：`講題`、`講員`、`經文`、`宣召`、`金句`、`聖詩一`、`啟應文`、`聖詩二`、`頌榮`。
- `宣召`、`經文`、`金句` 的 `sourceValue` 交給共用 `LKC_ppt_generator/bible-service.js`，以 `tghg` 取得台語經文全文，再由 `slide-production.js` 依每頁兩節建立投影片。
- `聖詩一`、`聖詩二`、`頌榮` 與 `啟應文` 的 `sourceValue` 只作資料庫索引。`ppt-library-integration.js` 以編號配對 Drive 索引，再由 `pptx-library.js` 在瀏覽器解壓縮 OOXML。
- 僅聖詩樂譜與啟應文資料庫頁會合成透明 PNG；這些 PNG 不含使用者背景，匯出時仍依序疊在「全份共用背景 → 聖詩頁白色色塊 → 樂譜／啟應文內容」之上。白色色塊提供 40–80% 透明度。其餘頁面不得用整頁圖取代文字。
- 所有聖詩相關段落（會前／正式聖詩、261、306B、頌榮、522）的白色色塊透明度預設同步調整；上方工具列提供預設勾選的「聖詩頁白底透明度同步」。取消勾選後只更新目前段落，其他聖詩保留各自透明度。
- `vendor-jszip.min.js` 固定隨 app 提供，避免外部 CDN 未載入時阻塞整個編輯器。
- `read-api.js` 正常情況沿用 `churchAPI` POST；若頁面由 `file://` 開啟，或 POST 因瀏覽器跨來源政策失敗，則改用同一 GAS 網址的唯讀 JSONP。JSONP 僅開放 `cal_getEvents`、`cal_getPptLibraryIndex`、`cal_getPptLibraryFile`、`cal_queryBible`，不提供任何寫入 action。`content-generators.js` 先在瀏覽器解析經文範圍，再由 GAS 伺服器查詢台語聖經，避免 `file://` 直接 fetch 信望愛 API。
- `layout-groups.js` 將所有流程攤平成一份連續 PPT，並直接把 PowerPoint 式「章節」整合到左側禮拜流程：流程編號左側可勾選整章，展開後以子項目顯示並勾選個別頁面。整個工具只使用右側既有的 16:9 共用預覽畫布，不再建立第二個版面畫布。
- 點頁面列只切換預覽，不改變版面選取；勾選頁面會加入待修改範圍並跳到該頁，勾選章節則批次選取並跳到章節第一頁。使用者由頂端「版面參數」按鈕開啟小型橫向懸浮面板。面板以「標題／內文」切換調整目前共用畫布，開啟後常駐至手動關閉，並可拖曳避開畫布。儲存後將同一組參數套用到所有勾選頁並立即寫入瀏覽器草稿。
- 背景模板由全份 PPT 共用：可使用純色 `backgroundColor`，或從頂端按鈕上傳 PNG／JPG／WebP 背景圖。背景圖以 16:9 畫布滿版裁切並可隨草稿保存；背景層不放置固定頁首或教會名稱。版面參數群組分別保存 `titleColor` 與 `contentColor`，讓真正的標題、內文可獨立選色。
- 開啟懸浮面板時，起始參數由目前共用畫布的實際 computed layout 換算，包含字級、X/Y、寬高、對齊與行距；因此不同投影片模板不再共用一組與畫布不符的固定起始值。
- `ppt-format-preview.js` 允許在非輸入欄位使用左右／上下方向鍵依整份 PPT 順序跨章切換；輸入欄位內的方向鍵仍保留游標操作。
- 版面群組與頁面歸屬以 Firebase RTDB `worshipPpt/layoutConfig/shared` 為全教會共用來源；`localStorage` 只保留首次遷移來源與離線備份。版面編輯預設鎖定，須經 Firebase Email/Password Authentication 解鎖後才可寫入；登入只保存在記憶體，重新整理即恢復鎖定。解除群組後，該頁可保留獨立 `page.layout` 微調值。
- 本 app 設定 `window._FORCE_PRODUCTION_GAS = true`，因此 localhost 也讀正式 `LKC_MasterSchedule`；其他 app 仍保留原有 localhost 自動切換 `_TEST` 的行為。

## 2026-07-14 與來源 PPT 對齊

- 流程保留來源 PPT 的會前段落：首頁、`會前領唱`、第一首聖詩、第二首聖詩，再進入第二張禮拜首頁與 `靜默一分鐘`。
- 行事曆的兩個聖詩編號會同時供應會前與正式禮拜的對應聖詩；相同 Drive 檔案依 file ID 共用解析快取。
- 固定從 Drive 載入主禱文後的第 `261` 首（副歌）、奉獻第 `306B` 首、阿們頌第 `522` 首。只有 `306B` 另有奉獻標題頁，對齊來源 PPT。
- PPTX 解析支援 PowerPoint 主題色 `schemeClr` 與母片色彩對應，因此啟應文的「啓／應」文字顏色不會被統一成黑色。
- 聖詩、奉獻、頌榮與報告依來源 PPT 明確保留各自應有的標題頁。

## 2026-07-15 畫布／匯出一致性

- `slide-production.js` 是畫布與匯出的單一版面來源。`resolvedLayoutForPage()` 依序合併投影片類型預設值、共用版面群組與頁面個別值；兩端不得各自維護另一套預設座標。
- PowerPoint 16:9 畫布寬為 960pt，所以預覽字級固定使用 `1pt = 1 / 9.6cqw`；版面面板反向量測也使用同一換算。文字與圖片輸出比例會同時反映在預覽與匯出。
- 原生文字頁先由 `wrapTextForBox()` 產生共用的明確換行，再由瀏覽器與 PptxGenJS 分別排版，因此使徒信經、主禱文、經文、報告與讚美不會因兩個文字引擎的字寬差異而換成不同的行。
- 所有產生頁使用微軟正黑粗體、零文字框邊界與頂端錨定；聖詩標題頁的 60pt 標題／副標題位置依來源 PPT 校準。
- 聖詩標題頁是原生文字頁，不套白色色塊；只有 `ppt-import`／`score` 樂譜頁套用白色色塊。啟應文點陣化前會統一標題方塊的垂直置中，避免第一張與後續有標題的頁面位置不同。
