// 主應用控制器 - 教會週報管理系統

const App = {
  _autoSaveTimer: null,

  _formatLocalDate(d) {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  },

  async init() {
    console.log('[App] 初始化教會週報管理系統...');

    const today = new Date();
    const daysToSunday = today.getDay() === 0 ? 0 : 7 - today.getDay();
    const nextSunday = new Date(today);
    nextSunday.setDate(today.getDate() + daysToSunday);
    const sundayStr = this._formatLocalDate(nextSunday);

    document.getElementById('bulletinDate').value = sundayStr;
    BulletinModel.init(sundayStr);

    document.getElementById('bulletinDate').addEventListener('change', e => {
      BulletinModel.set('date', e.target.value);
      this.updateDateDisplay();
    });

    this.initTabs();
    this.initFormFields();
    this.initButtons();
    this.updateDateDisplay();
    this.syncFormFromModel();
    DraftManager.startAutoSave(() => BulletinModel.get());
    this.showToast('系統已就緒，歡迎使用教會週報管理系統', 'success');
  },

  updateDateDisplay() {
    const date = document.getElementById('bulletinDate').value;
    if (date) {
      const d = new Date(date + 'T00:00:00');
      document.getElementById('dateDisplay').textContent =
        `${d.getFullYear()}年${d.getMonth()+1}月${d.getDate()}日`;

      const nextD = new Date(d);
      nextD.setDate(d.getDate() + 7);
      const nextDate = this._formatLocalDate(nextD);
      BulletinModel.set('ministry.thisWeek.date', date);
      BulletinModel.set('ministry.nextWeek.date', nextDate);
    }
  },

  initTabs() {
    document.querySelectorAll('.tab-btn').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
        tab.classList.add('active');
        document.getElementById(`tab-${tab.dataset.tab}`)?.classList.add('active');
      });
    });
  },

  initFormFields() {
    document.addEventListener('input',  e => { const f = e.target.dataset.field; if (f) BulletinModel.set(f, e.target.value); });
    document.addEventListener('change', e => { const f = e.target.dataset.field; if (f) BulletinModel.set(f, e.target.value); });
  },

  initButtons() {
    document.getElementById('btnFetchAll')   ?.addEventListener('click', () => this.fetchAll());
    document.getElementById('btnSaveDraft')  ?.addEventListener('click', () => this.saveDraft());
    document.getElementById('btnLoadDraft')  ?.addEventListener('click', () => this.showDraftModal());
    document.getElementById('btnExportWord') ?.addEventListener('click', () => this.exportWord());
    document.getElementById('btnAddOffering')?.addEventListener('click', () => this.addOfferingRow());
    document.getElementById('btnAddEvent')   ?.addEventListener('click', () => this.addEventRow());
    document.getElementById('modalClose')    ?.addEventListener('click', () => this.hideDraftModal());
    document.getElementById('modalOverlay')  ?.addEventListener('click', e => { if (e.target === e.currentTarget) this.hideDraftModal(); });
  },

  async fetchAll() {
    const date = document.getElementById('bulletinDate').value;
    if (!date) { this.showToast('請先選擇日期', 'error'); return; }
    this.showLoading(true);
    this.showToast('正在從各系統帶入資料...', 'info');
    try {
      const results = await ChurchAPI.fetchAll(date);
      BulletinModel.applyAPIData(results);
      this.syncFormFromModel();
      const failed = Object.entries(results)
        .filter(([, v]) => !v.success)
        .map(([k]) => ({ calendar:'行事曆', service:'服事排班', worship:'敬拜團', worshipSongs:'敬拜曲目', attendance:'點名(手動)', smallGroups:'小組' }[k] || k));
      this.showToast(failed.length ? `帶入完成，請手動填入：${failed.join('、')}` : '全部資料帶入完成', failed.length ? 'warning' : 'success');
    } catch (err) {
      this.showToast('帶入失敗：' + err.message, 'error');
    } finally {
      this.showLoading(false);
    }
  },

  async fetchServiceProgram() {
    const date = document.getElementById('bulletinDate').value;
    if (!date) { this.showToast('請先選擇日期', 'error'); return; }
    this.showLoading(true);
    this.showToast('正在帶入主日程序資料...', 'info');
    try {
      const [calR, svcR, songsR] = await Promise.allSettled([
        ChurchAPI.fetchCalendarForDate(date),
        ChurchAPI.fetchServiceSchedule(date),
        ChurchAPI.fetchWorshipSongs(date)
      ]);
      const v = s => s.status === 'fulfilled' ? s.value : { success: false, error: s.reason?.message };
      const calResult   = v(calR);
      const svcResult   = v(svcR);
      const songsResult = v(songsR);

      BulletinModel.applyAPIData({ calendar: calResult, service: svcResult, worshipSongs: songsResult });
      this.syncFormFromModel();
      const msgs = [];
      if (!calResult.success) msgs.push(`行事曆失敗: ${calResult.error}`);
      else if (!calResult.data?.taiwanese && !calResult.data?.mandarin) msgs.push(`找不到 ${date} 的講道資訊`);
      if (!svcResult.success) msgs.push(`服事排班失敗: ${svcResult.error}`);
      if (!songsResult.success) msgs.push(`敬拜曲目失敗: ${songsResult.error}`);
      this.showToast(msgs.length ? msgs.join('；') : '主日程序資料帶入完成', msgs.length ? 'warning' : 'success');
    } catch (err) {
      this.showToast('帶入失敗：' + err.message, 'error');
    } finally {
      this.showLoading(false);
    }
  },

  async fetchMinistry() {
    const date = document.getElementById('bulletinDate').value;
    if (!date) { this.showToast('請先選擇日期', 'error'); return; }

    const nextD = new Date(date + 'T00:00:00');
    nextD.setDate(nextD.getDate() + 7);
    const nextDate = this._formatLocalDate(nextD);

    this.showLoading(true);
    this.showToast('正在帶入服事人員資料（本週 + 下週）...', 'info');
    try {
      const [calS, svcS, worS, nextCalS, nextSvcS, nextWorS] = await Promise.allSettled([
        ChurchAPI.fetchCalendarForDate(date),
        ChurchAPI.fetchServiceSchedule(date),
        ChurchAPI.fetchWorshipSchedule(date),
        ChurchAPI.fetchCalendarForDate(nextDate),
        ChurchAPI.fetchServiceSchedule(nextDate, true),
        ChurchAPI.fetchWorshipSchedule(nextDate, true)
      ]);
      const v = s => s.status === 'fulfilled' ? s.value : { success: false, error: s.reason?.message };

      BulletinModel.applyAPIData(      { calendar: v(calS),     service: v(svcS),     worship: v(worS)     });
      BulletinModel.applyNextWeekAPIData({ calendar: v(nextCalS), service: v(nextSvcS), worship: v(nextWorS) });
      this.syncFormFromModel();

      const failed = [
        !v(svcS).success    && `本週服事排班（${v(svcS).error    || ''}）`,
        !v(worS).success    && `本週敬拜團（${v(worS).error      || ''}）`,
        !v(nextSvcS).success && `下週服事排班（${v(nextSvcS).error || ''}）`,
        !v(nextWorS).success && `下週敬拜團（${v(nextWorS).error  || ''}）`
      ].filter(Boolean);
      this.showToast(
        failed.length ? `帶入完成，請手動確認：${failed.join('、')}` : '服事人員資料帶入完成（本週＋下週）',
        failed.length ? 'warning' : 'success'
      );
    } catch (err) {
      this.showToast('帶入失敗：' + err.message, 'error');
    } finally {
      this.showLoading(false);
    }
  },

  async fetchSection(section) {
    const date = document.getElementById('bulletinDate').value;
    if (!date) { this.showToast('請先選擇日期', 'error'); return; }
    this.showLoading(true);
    try {
      const map = {
        calendar:    () => ChurchAPI.fetchCalendarForDate(date),
        service:     () => ChurchAPI.fetchServiceSchedule(date),
        worship:     () => ChurchAPI.fetchWorshipSchedule(date),
        attendance:  () => ChurchAPI.fetchAttendance(date),
        smallGroups: () => ChurchAPI.fetchSmallGroups(date)
      };
      const result = await map[section]();
      if (result.success) {
        BulletinModel.applyAPIData({ [section]: result });
        this.syncFormFromModel();
        this.showToast('資料帶入完成', 'success');
      } else {
        this.showToast('帶入失敗：' + (result.error || '未知錯誤'), 'error');
      }
    } catch (err) {
      this.showToast('帶入失敗：' + err.message, 'error');
    } finally {
      this.showLoading(false);
    }
  },

  // 聚會統計 tab：同時帶入出席人數 + 小組人數，皆以主日日期為參照
  async fetchAttendanceTab() {
    const date = document.getElementById('bulletinDate').value;
    if (!date) { this.showToast('請先選擇日期', 'error'); return; }
    this.showLoading(true);
    this.showToast(`正在帶入聚會人數（參照主日 ${date}）...`, 'info');
    try {
      const [attR, sgR] = await Promise.allSettled([
        ChurchAPI.fetchAttendance(date),
        ChurchAPI.fetchSmallGroups(date)
      ]);
      const v = s => s.status === 'fulfilled' ? s.value : { success: false, error: s.reason?.message };
      const attendance  = v(attR);
      const smallGroups = v(sgR);

      BulletinModel.applyAPIData({ attendance, smallGroups });
      this.syncFormFromModel();

      const failed = [
        !attendance.success  && `出席人數（${attendance.error  || '未知錯誤'}）`,
        !smallGroups.success && `小組人數（${smallGroups.error || '未知錯誤'}）`
      ].filter(Boolean);
      this.showToast(
        failed.length ? `帶入完成，請手動確認：${failed.join('、')}` : '聚會人數資料帶入完成',
        failed.length ? 'warning' : 'success'
      );
    } catch (err) {
      this.showToast('帶入失敗：' + err.message, 'error');
    } finally {
      this.showLoading(false);
    }
  },

  async saveDraft() {
    try {
      const saved = await DraftManager.save(BulletinModel.get());
      this.showToast(saved ? '草稿已儲存' : '儲存失敗，請確認已選擇日期', saved ? 'success' : 'error');
    } catch (err) {
      this.showToast('草稿儲存失敗：' + err.message, 'error');
    }
  },

  async showDraftModal() {
    const list = document.getElementById('draftList');
    list.innerHTML = '<div class="empty-state">載入中...</div>';
    document.getElementById('modalOverlay').classList.add('show');
    try {
      const drafts = await DraftManager.list();
      list.innerHTML = '';
      if (!drafts || drafts.length === 0) {
        list.innerHTML = '<div class="empty-state">尚無已儲存的草稿</div>';
        return;
      }
      drafts.forEach(draft => {
        const item = document.createElement('div');
        item.className = 'draft-item';
        item.innerHTML = `
          <div class="draft-info">
            <strong>${draft.date}</strong>
            <span class="draft-preview">${draft.preview || ''}</span>
            <small>最後儲存：${new Date(draft.updatedAt).toLocaleString('zh-TW')}</small>
          </div>
          <div class="draft-actions">
            <button class="btn-sm btn-primary" onclick="App.loadDraft('${draft.date}')"> 載入</button>
            <button class="btn-sm btn-danger"  onclick="App.deleteDraft('${draft.date}')"> 刪除</button>
          </div>`;
        list.appendChild(item);
      });
    } catch (err) {
      list.innerHTML = `<div class="empty-state">載入失敗：${err.message}</div>`;
    }
  },

  hideDraftModal() { document.getElementById('modalOverlay').classList.remove('show'); },

  async loadDraft(date) {
    try {
      const data = await DraftManager.load(date);
      if (!data) { this.showToast('載入失敗', 'error'); return; }
      BulletinModel._current = data;
      document.getElementById('bulletinDate').value = data.date;
      this.syncFormFromModel();
      this.hideDraftModal();
      this.showToast(`草稿 ${date} 已載入`, 'success');
    } catch (err) { this.showToast('載入草稿失敗：' + err.message, 'error'); }
  },

  async deleteDraft(date) {
    if (!confirm(`確定要刪除 ${date} 的草稿？`)) return;
    try {
      await DraftManager.delete(date);
      this.showDraftModal();
      this.showToast('草稿已刪除', 'success');
    } catch (err) { this.showToast('刪除失敗：' + err.message, 'error'); }
  },

  async exportWord() {
    const data = BulletinModel.get();
    if (!data.date) { this.showToast('請先選擇日期', 'error'); return; }
    this.showLoading(true);
    this.showToast('正在產生 Word 文件...', 'info');
    try {
      const filename = await BulletinExport.generate(data);
      this.showToast(`Word 文件已下載：${filename}`, 'success');
    } catch (err) {
      this.showToast('匯出失敗：' + err.message, 'error');
    } finally {
      this.showLoading(false);
    }
  },

  syncFormFromModel() {
    const data = BulletinModel.get();
    document.querySelectorAll('[data-field]').forEach(el => {
      const v = el.dataset.field.split('.').reduce((o, k) => o?.[k], data);
      if (v !== undefined && v !== null) el.value = String(v);
    });
    this.syncSmallGroupsUI(data.attendance?.smallGroups || {});
    this.syncEventsUI(data.events || []);
    this.syncOfferingUI(data.offeringReport?.monthlyItems || []);
  },

  // 動態渲染小組欄位（內容由 API 或 model 決定）
  syncSmallGroupsUI(groups) {
    const container = document.getElementById('smallGroupsContainer');
    if (!container) return;
    const entries = Object.entries(groups);
    if (entries.length === 0) {
      container.innerHTML = '<div style="padding:16px;color:#999;text-align:center;width:100%">點擊「⬇ 自動帶入」從 LKC_Attendance 與 LKGroup 自動載入</div>';
      return;
    }
    container.innerHTML = '';
    entries.forEach(([name, count]) => {
      const div = document.createElement('div');
      div.className = 'small-group-item';
      div.innerHTML = `<label>${name}</label><input type="number" data-group="${name}" min="0" value="${parseInt(count) || 0}" onchange="App._updateGroup('${name}', this.value)">`;
      container.appendChild(div);
    });
  },

  syncEventsUI(events) {
    const c = document.getElementById('eventsContainer'); if (!c) return;
    c.innerHTML = ''; events.forEach((ev, i) => this.addEventRow(ev, i));
  },

  syncOfferingUI(items) {
    const c = document.getElementById('offeringContainer'); if (!c) return;
    c.innerHTML = ''; items.forEach((item, i) => this.addOfferingRow(item, i));
  },

  addEventRow(ev = null, idx = null) {
    const c = document.getElementById('eventsContainer'); if (!c) return;
    const i = idx !== null ? idx : c.children.length;
    const div = document.createElement('div'); div.className = 'dynamic-row'; div.dataset.idx = i;
    div.innerHTML = `
      <input type="date" class="form-input" value="${ev?.date||""}" onchange="App._updateEvent(${i},'date',this.value)">
      <input type="text" class="form-input" value="${ev?.name||""}" oninput="App._updateEvent(${i},'name',this.value)">
      <input type="text" class="form-input flex-2" value="${ev?.description||""}" oninput="App._updateEvent(${i},'description',this.value)">
      <button class="btn-icon btn-danger" onclick="App._removeEvent(${i})">&#x2715;</button>`;
    c.appendChild(div);
  },
  _updateEvent(i, f, v) { const d = BulletinModel.get(); if (!d.events[i]) d.events[i]={}; d.events[i][f]=v; },
  _removeEvent(i) { const d = BulletinModel.get(); d.events.splice(i,1); this.syncEventsUI(d.events); },

  addOfferingRow(item = null, idx = null) {
    const c = document.getElementById('offeringContainer'); if (!c) return;
    const i = idx !== null ? idx : c.children.length;
    const div = document.createElement('div'); div.className = 'dynamic-row'; div.dataset.idx = i;
    div.innerHTML = `
      <input type="text" class="form-input" value="${item?.name||""}"   oninput="App._updateOffering(${i},'name',this.value)">
      <input type="text" class="form-input" value="${item?.amount||""}" oninput="App._updateOffering(${i},'amount',this.value)">
      <input type="text" class="form-input" value="${item?.note||""}"   oninput="App._updateOffering(${i},'note',this.value)">
      <button class="btn-icon btn-danger" onclick="App._removeOffering(${i})">&#x2715;</button>`;
    c.appendChild(div);
  },
  _updateOffering(i, f, v) { const d = BulletinModel.get(); if (!d.offeringReport.monthlyItems[i]) d.offeringReport.monthlyItems[i]={}; d.offeringReport.monthlyItems[i][f]=v; },
  _removeOffering(i) { const d = BulletinModel.get(); d.offeringReport.monthlyItems.splice(i,1); this.syncOfferingUI(d.offeringReport.monthlyItems); },

  showToast(message, type = 'info') {
    const c = document.getElementById('toastContainer'); if (!c) return;
    const t = document.createElement('div'); t.className = `toast toast-${type}`; t.textContent = message; c.appendChild(t);
    setTimeout(() => t.classList.add('show'), 10);
    setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 300); }, 4500);
  },

  showLoading(show) {
    const o = document.getElementById('loadingOverlay');
    if (o) o.style.display = show ? 'flex' : 'none';
  }
};

document.addEventListener('DOMContentLoaded', () => App.init());
