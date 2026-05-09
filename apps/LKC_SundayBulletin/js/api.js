// API 客戶端 - 教會週報管理系統

const ChurchAPI = {

  // 共用 fetch helper：統一處理 HTTP 錯誤與 JSON 解析
  async _postJSON(url, init) {
    const res = await fetch(url, init);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return await res.json();
  },

  // 通用 GAS：text/plain，body = { action, token, data }
  callGAS(url, action, data = {}) {
    return this._postJSON(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action, token: CONFIG.SHARED_TOKEN, data })
    });
  },

  // LKC1958：form-urlencoded，避開 GAS preflight；body 包在 payload 欄位
  callLKC1958(action, data = {}) {
    const payload = { action, token: CONFIG.SHARED_TOKEN, data };
    return this._postJSON(CONFIG.LKC1958_GAS_URL, {
      method:   'POST',
      headers:  { 'Content-Type': 'application/x-www-form-urlencoded' },
      body:     'payload=' + encodeURIComponent(JSON.stringify(payload)),
      redirect: 'follow'
    });
  },

  // Attendance：text/plain，body 形狀為 { action, payload }（不同 schema）
  callAttendance(action, payload = null) {
    return this._postJSON(CONFIG.LKC_ATTENDANCE_GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({ action, payload })
    });
  },

  _unwrap(result) {
    if (Array.isArray(result)) return result;
    if (result && result.status === 'success' && result.data !== undefined) return result.data;
    if (result && result.success === true && result.data !== undefined) return result.data;
    return null;
  },

  _dateMatch(cellDate, targetDate) {
    const d = String(cellDate || '');
    if (!d || !targetDate) return false;
    if (d.startsWith(targetDate) || d.includes(targetDate)) return true;
    const cellClean   = d.replace(/[^\d]/g, '').substring(0, 8);
    const targetClean = targetDate.replace(/[^\d]/g, '');
    return cellClean === targetClean;
  },

  async fetchServiceSchedule(sundayDate, requireMatch = false) {
    try {
      const result = await this.callLKC1958('getAggregatedReport', { type: 'others' });
      const rawData = this._unwrap(result);
      if (!Array.isArray(rawData) || rawData.length === 0) throw new Error('資料格式不符');

      const headers  = rawData[0] || [];
      const rows     = rawData.slice(1);
      const dateIdx  = headers.indexOf('日期');
      debug('[LKC1958] headers:', headers, '| dateIdx:', dateIdx, '| looking for:', sundayDate);

      const matchingRows = dateIdx !== -1
        ? rows.filter(row => this._dateMatch(row[dateIdx], sundayDate))
        : rows.filter(row => this._dateMatch(row[0], sundayDate));

      debug('[LKC1958] matchingRows:', matchingRows.length);

      let sourceRows;
      if (matchingRows.length > 0) {
        sourceRows = matchingRows;
      } else if (requireMatch) {
        return { success: false, source: 'LKC1958_June_1', error: '查無資料' };
      } else {
        sourceRows = [rows[rows.length - 1] || []];
      }

      const r = {};
      headers.forEach((h, i) => {
        if (!h) return;
        for (const row of sourceRows) {
          const v = String(row[i] || '').trim();
          if (v !== '') { r[h] = v; break; }
        }
        if (r[h] === undefined) r[h] = '';
      });

      const newcomer = [r['新家人同工1'] || '', r['新家人同工2'] || ''].filter(Boolean).join('、');

      return {
        success: true, source: 'LKC1958_June_1',
        data: {
          mc:           r['台語司會'] || r['司會'] || r['司儀'] || '',
          zhMc:         r['華語司會'] || '',
          pianist:      r['司琴'] || '',
          choir:        r['詩班'] || '',
          usher:        r['招待/停車'] || r['招待'] || '',
          chairman:     r['主席'] || r['主理'] || '',
          songLeader:   r['領詩'] || '',
          soundControl: r['音控同工'] || r['音控'] || '',
          newcomerCare: newcomer,
          raw: r
        }
      };
    } catch (err) {
      console.error('[LKC1958]', err);
      return { success: false, source: 'LKC1958_June_1', error: err.message };
    }
  },

  // ==========================================
  // LKworship - 敬拜曲目（華語禮拜）
  //
  // getSongs 動作支援 { startDate, endDate } 或 { year, quarter }
  // 回傳每列含：日期 / 聚會名稱 / 聚會類別 / 主領 / 敬拜曲目
  // ==========================================
  async fetchWorshipSongs(date, requireMatch = false) {
    try {
      const result = await this.callGAS(CONFIG.LKWORSHIP_GAS_URL, 'getSongs', {
        startDate: date,
        endDate:   date
      });
      const rows = this._unwrap(result);
      debug('[LKworship-songs]', date, '| rows:', Array.isArray(rows) ? rows.length : 0);

      if (!Array.isArray(rows) || rows.length === 0) {
        return requireMatch
          ? { success: false, source: 'LKworship-songs', error: '查無曲目資料' }
          : { success: true,  source: 'LKworship-songs', data: { songs: '', leader: '', raw: null } };
      }

      // 在當天的列中找華語場次（聚會類別或聚會名稱含「華語」）
      const dayRows = rows.filter(r => this._dateMatch(r['日期'] || r['date'], date));
      const zhRow =
        dayRows.find(r => String(r['聚會類別'] || '').includes('華語')) ||
        dayRows.find(r => String(r['聚會名稱'] || '').includes('華語')) ||
        dayRows[0] || null;

      if (!zhRow) {
        return { success: true, source: 'LKworship-songs', data: { songs: '', leader: '', raw: null } };
      }

      return {
        success: true, source: 'LKworship-songs',
        data: {
          songs:  zhRow['敬拜曲目'] || '',
          leader: zhRow['主領']     || '',
          raw:    zhRow
        }
      };
    } catch (err) {
      console.error('[LKworship-songs]', err);
      return { success: false, source: 'LKworship-songs', error: err.message };
    }
  },

  async fetchWorshipSchedule(date, requireMatch = false) {
    try {
      const d = new Date(date);
      const year    = d.getFullYear();
      const quarter = 'Q' + Math.ceil((d.getMonth() + 1) / 3);

      const result = await this.callGAS(CONFIG.LKWORSHIP_GAS_URL, 'getSchedule', { year, quarter });
      const scheduleData = this._unwrap(result);
      let row = null;
      if (Array.isArray(scheduleData)) {
        row = scheduleData.find(r => this._dateMatch(r['日期'] || r[0], date)) || null;
        if (!row && !requireMatch) row = scheduleData[scheduleData.length - 1] || null;
      }

      let singers = '';
      if (row) {
        const parts = [row['配唱1'] || '', row['配唱2'] || '', row['配唱3'] || ''].filter(Boolean);
        singers = parts.length > 0 ? parts.join('、') : (row['配唱'] || '');
      }

      return {
        success: true, source: 'LKworship',
        data: { leader: row ? (row['主領'] || '') : '', singers, raw: row || {} }
      };
    } catch (err) {
      console.error('[LKworship]', err);
      return { success: false, source: 'LKworship', error: err.message };
    }
  },

  async fetchCalendar() {
    try {
      const result = await this.callGAS(CONFIG.LKCSCHEDULE_GAS_URL, 'load', {});
      debug('[LKCschedule] raw (first 800):', JSON.stringify(result).substring(0, 800));

      let events = [];
      if (result && result.success && Array.isArray(result.events)) {
        events = result.events;
      } else {
        const raw = this._unwrap(result);
        events = Array.isArray(raw) ? raw : [];
      }

      return {
        success: true, source: 'LKCschedule',
        data: events.map(e => {
          const sermons = e.sermons || [];
          const twS = sermons.find(s => s.type === '台語/聯合') || null;
          const zhS = sermons.find(s => s.type === '華語') || null;
          return {
            date:         e['日期']     || e.date         || '',
            name:         e['聚會名稱'] || e.name         || '',
            category:     e['講道類別'] || e['聚會類別'] || e.category || '',
            sermonTitle:  twS?.title          || e['講題']  || e.sermonTitle  || '',
            speaker:      twS?.speaker        || e['講員']  || e.speaker      || '',
            scripture:    twS?.scripture      || e['經文']  || e.scripture    || '',
            callToWorship:twS?.callToWorship  || e['宣召']  || e.callToWorship|| '',
            goldenVerse:  twS?.goldenVerse    || e['金句']  || e.goldenVerse  || '',
            hymn:         twS?.hymns          || e['詩歌']  || e['聖詩'] || e.hymn || '',
            notes:        e['備註']           || e.notes        || '',
            zhSermonTitle: zhS?.title     || '',
            zhSpeaker:     zhS?.speaker   || '',
            zhScripture:   zhS?.scripture || '',
            hasTwSermon: !!twS,
            hasZhSermon: !!zhS,
            raw: e
          };
        })
      };
    } catch (err) {
      console.error('[LKCschedule]', err);
      return { success: false, source: 'LKCschedule', error: err.message };
    }
  },

  async fetchCalendarForDate(date) {
    const result = await this.fetchCalendar();
    if (!result.success) return result;
    const events = result.data;

    const dayEvents = events.filter(e => this._dateMatch(e.date, date));
    debug('[LKCschedule] date:', date, '| dayEvents:', dayEvents.length);

    const twEvent = dayEvents.find(e => e.hasTwSermon)
      || dayEvents.find(e => e.category.includes('台語') || e.name.includes('台語') || e.category === '主日');
    const zhEvent = dayEvents.find(e => e.hasZhSermon)
      || dayEvents.find(e => e.category.includes('華語') || e.name.includes('華語'));

    const twService = twEvent || null;
    const zhService = zhEvent ? {
      date:        zhEvent.date,
      name:        zhEvent.name,
      sermonTitle: zhEvent.zhSermonTitle || zhEvent.sermonTitle || '',
      speaker:     zhEvent.zhSpeaker     || zhEvent.speaker     || '',
      scripture:   zhEvent.zhScripture   || zhEvent.scripture   || ''
    } : null;

    const today = new Date(date);
    const limit = new Date(today); limit.setMonth(limit.getMonth() + 3);
    const upcoming = events
      .filter(e => { const d = new Date(e.date); return d >= today && d <= limit; })
      .sort((a, b) => new Date(a.date) - new Date(b.date));

    return { success: true, source: 'LKCschedule', data: { taiwanese: twService, mandarin: zhService, upcoming } };
  },

  // ==========================================
  // LKC_Attendance - 主日禮拜出席人數
  //
  // 報告的是「上一個週日」的人數（週報日期 - 7 天）
  // ==========================================
  async fetchAttendance(date) {
    try {
      // 主日禮拜人數報告上一個週日的資料
      const d = new Date(date + 'T00:00:00');
      d.setDate(d.getDate() - 7);
      const prevSunday = formatYMD(d);
      debug('[LKC_Attendance] 查詢日期:', prevSunday, '(週報日期', date, '-7天)');

      const req = (type) => ({ type, mode: 'single', date: prevSunday, baseSheet: '會友名單', targetGroups: [] });

      const [twRaw, zhRaw] = await Promise.allSettled([
        this.callAttendance('getAttendanceStats', req('台語')),
        this.callAttendance('getAttendanceStats', req('華語'))
      ]);

      debug('[LKC_Attendance] 台語 raw:', twRaw.status === 'fulfilled' ? twRaw.value : twRaw.reason);
      debug('[LKC_Attendance] 華語 raw:', zhRaw.status === 'fulfilled' ? zhRaw.value : zhRaw.reason);

      const parseTotal = (settled) => {
        if (settled.status !== 'fulfilled') return 0;
        const data = settled.value?.data ?? settled.value;
        if (!data) return 0;
        return (Number(data.presentCount) || 0) + (Number(data.nfMale) || 0) + (Number(data.nfFemale) || 0);
      };

      return {
        success: true, source: 'LKC_Attendance',
        data: {
          taiwanese: { total: parseTotal(twRaw) },
          mandarin:  { total: parseTotal(zhRaw) }
        }
      };
    } catch (err) {
      console.error('[LKC_Attendance]', err);
      return { success: false, source: 'LKC_Attendance', error: err.message };
    }
  },

  // ==========================================
  // LKGroup - 小組人數
  //
  // 1. getGroups        → { success, groups: [{name}] }  完整小組列表
  // 2. getWeeklyReport  → { success, data: [{groupName, total, newFriends}], dateRange }
  //    統計區間：以主日日期為參照，往前推一週（上一週日 ~ 主日前一日週六）
  // ==========================================
  async fetchSmallGroups(date) {
    try {
      // 計算以主日日期為參照的小組統計區間：上一週日 ~ 主日前一日（週六）
      const sundayD   = new Date(date + 'T00:00:00');
      const prevSunD  = new Date(sundayD); prevSunD.setDate(sundayD.getDate() - 7);
      const prevSatD  = new Date(sundayD); prevSatD.setDate(sundayD.getDate() - 1);
      const startDate = formatYMD(prevSunD);
      const endDate   = formatYMD(prevSatD);
      debug('[LKGroup] 參照主日:', date, '統計區間:', startDate, '~', endDate);

      // Step 1: 取得小組列表
      let groupNames = [];
      try {
        const groupsRes = await this.callGAS(CONFIG.LKGROUP_GAS_URL, 'getGroups', {});
        debug('[LKGroup] getGroups response:', JSON.stringify(groupsRes).substring(0, 500));
        if (groupsRes?.success && Array.isArray(groupsRes.groups)) {
          groupNames = groupsRes.groups.map(g => g.name || g).filter(Boolean);
          debug('[LKGroup] 小組列表:', groupNames);
        }
      } catch (e) {
        debug('[LKGroup] getGroups 失敗，改用預設列表:', e.message);
      }
      if (groupNames.length === 0) groupNames = CONFIG.TW_GROUPS;

      // Step 2: 建立小組 map，預設人數 0
      const results = {};
      groupNames.forEach(name => { results[name] = { date: '', attendance: 0, newFriends: '' }; });

      // Step 3: 取得以主日日期為參照的小組人數
      try {
        const weeklyRes = await this.callGAS(CONFIG.LKGROUP_GAS_URL, 'getWeeklyReport', {
          referenceDate: date,
          startDate,
          endDate
        });
        debug('[LKGroup] getWeeklyReport response:', JSON.stringify(weeklyRes).substring(0, 500));
        if (weeklyRes?.success && Array.isArray(weeklyRes.data)) {
          weeklyRes.data.forEach(g => {
            if (Object.prototype.hasOwnProperty.call(results, g.groupName)) {
              results[g.groupName] = {
                date:       weeklyRes.dateRange || '',
                attendance: Number(g.total)     || 0,
                newFriends: g.newFriends        || ''
              };
            }
          });
        }
      } catch (e) {
        debug('[LKGroup] getWeeklyReport 失敗:', e.message);
      }

      return { success: true, source: 'LKGroup', data: results };
    } catch (err) {
      console.error('[LKGroup]', err);
      return { success: false, source: 'LKGroup', error: err.message };
    }
  },

  async fetchAll(date) {
    const [calendar, service, worship, worshipSongs, attendance, smallGroups] = await Promise.allSettled([
      this.fetchCalendarForDate(date),
      this.fetchServiceSchedule(date),
      this.fetchWorshipSchedule(date),
      this.fetchWorshipSongs(date),
      this.fetchAttendance(date),
      this.fetchSmallGroups(date)
    ]);
    const val = r => r.status === 'fulfilled' ? r.value : { success: false, error: r.reason?.message };
    return {
      calendar:     val(calendar),
      service:      val(service),
      worship:      val(worship),
      worshipSongs: val(worshipSongs),
      attendance:   val(attendance),
      smallGroups:  val(smallGroups)
    };
  }
};
