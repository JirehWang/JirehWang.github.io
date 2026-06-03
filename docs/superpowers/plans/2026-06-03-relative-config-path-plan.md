# HTML config.js 引用路徑改為相對路徑 實現計畫

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推薦）或 superpowers:executing-plans 逐任務實現此計畫。步驟使用複選框（`- [ ]`）語法來跟踪進度。

**目標：** 將 13 個 HTML 檔案中寫死的 `config.js` 網址替換為相對路徑 `../../config.js`。

**建築：** 透過直接字串替換，將 `<script src="https://jirehwang.github.io/LKC1958_June_1.github.io/config.js"></script>` 改為相對路徑引用，以支持本地開發與分支測試。

**技術棧：** HTML, Git.

---

## 任務 1：修改 LKC_Group 子應用程式的 HTML 檔案

**文件：**
- 修改：`apps/LKC_Group/group.html`
- 修改：`apps/LKC_Group/index.html`
- 修改：`apps/LKC_Group/manage.html`
- 修改：`apps/LKC_Group/stats.html`

- [ ] **步驟 1：修改 `apps/LKC_Group/group.html`**
  將引用改為相對路徑。
- [ ] **步驟 2：修改 `apps/LKC_Group/index.html`**
  將引用改為相對路徑。
- [ ] **步驟 3：修改 `apps/LKC_Group/manage.html`**
  將引用改為相對路徑。
- [ ] **步驟 4：修改 `apps/LKC_Group/stats.html`**
  將引用改為相對路徑。

---

## 任務 2：修改 LKC_MasterSchedule 與 LKC_MinistrySchedule 子應用程式的 HTML 檔案

**文件：**
- 修改：`apps/LKC_MasterSchedule/calendar.html`
- 修改：`apps/LKC_MasterSchedule/types.html`
- 修改：`apps/LKC_MinistrySchedule/groupboard.html`
- 修改：`apps/LKC_MinistrySchedule/index.html`

- [ ] **步驟 1：修改 `apps/LKC_MasterSchedule/calendar.html`**
  將引用改為相對路徑。
- [ ] **步驟 2：修改 `apps/LKC_MasterSchedule/types.html`**
  將引用改為相對路徑。
- [ ] **步驟 3：修改 `apps/LKC_MinistrySchedule/groupboard.html`**
  將引用改為相對路徑。
- [ ] **步驟 4：修改 `apps/LKC_MinistrySchedule/index.html`**
  將引用改為相對路徑。

---

## 任務 3：修改 LKC_SundayserviceAttendance、LKC_WhosCar 與 LKC_worship 子應用程式的 HTML 檔案

**文件：**
- 修改：`apps/LKC_SundayserviceAttendance/index.html`
- 修改：`apps/LKC_WhosCar/index.html`
- 修改：`apps/LKC_worship/admin.html`
- 修改：`apps/LKC_worship/index.html`
- 修改：`apps/LKC_worship/worship_songs.html`

- [ ] **步驟 1：修改 `apps/LKC_SundayserviceAttendance/index.html`**
  將引用改為相對路徑。
- [ ] **步驟 2：修改 `apps/LKC_WhosCar/index.html`**
  將引用改為相對路徑。
- [ ] **步驟 3：修改 `apps/LKC_worship/admin.html`**
  將引用改為相對路徑。
- [ ] **步驟 4：修改 `apps/LKC_worship/index.html`**
  將引用改為相對路徑。
- [ ] **步驟 5：修改 `apps/LKC_worship/worship_songs.html`**
  將引用改為相對路徑。

---

## 任務 4：自審與完成前驗證

- [ ] **步驟 1：檢查 Git Diff**
  使用 `git diff` 確認修改符合預期。
- [ ] **步驟 2：使用 PowerShell 或 Grep 確認無遺留線上網址**
  確認 `apps/` 底下無 `https://jirehwang.github.io/LKC1958_June_1.github.io/config.js` 的存在。
