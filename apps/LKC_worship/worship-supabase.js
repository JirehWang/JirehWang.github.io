// ⚡ apps/LKC_worship/worship-supabase.js
// 敬拜團 Supabase 熱響應服務模組 (含冷熱資料自動分流)

(function(window) {
  const HOT_YEAR_THRESHOLD = 2025; // 2025 年及以後走 Supabase (熱響應 <50ms)

  let _supabaseClient = null;

  // 初始化 Supabase Client
  function getSupabase() {
    if (_supabaseClient) return _supabaseClient;
    const config = window._SUPABASE_CONFIG || {};
    if (!config.url || !config.anonKey) {
      console.warn('⚠️ Supabase 設定尚未載入');
      return null;
    }
    if (typeof window.supabase === 'undefined' && typeof createClient === 'undefined') {
      console.warn('⚠️ Supabase JS SDK 尚未載入');
      return null;
    }
    const create = (window.supabase && window.supabase.createClient) || createClient;
    _supabaseClient = create(config.url, config.anonKey);
    return _supabaseClient;
  }

  // 判斷是否為熱資料年份
  function isHotYear(year) {
    const y = parseInt(year, 10);
    return !isNaN(y) && y >= HOT_YEAR_THRESHOLD;
  }

  function isHotDate(dateStr) {
    if (!dateStr) return true;
    const y = parseInt(dateStr.slice(0, 4), 10);
    return isHotYear(y);
  }

  // 將 Supabase 資料庫中的 schedule row 轉換為前端預期的物件格式
  function transformDbScheduleToClient(row) {
    if (!row) return null;
    const base = {
      '日期': row.date,
      '聚會名稱': row.meeting_name || '',
      '聚會類別': row.meeting_category || '',
      '年度': row.year || '',
      '季度': row.quarter || '',
      '牧師': row.preacher || '',
      '題目': row.topic || '',
      '經文': row.scripture || '',
      '敬拜曲目': row.songs || '',
      'leaves': Array.isArray(row.leaves) ? row.leaves : []
    };

    // 展開動態崗位人員
    const assignments = (typeof row.positions_assignment === 'object' && row.positions_assignment) || {};
    Object.keys(assignments).forEach(k => {
      base[k] = assignments[k];
    });

    // 若有 raw_data 中的額外欄位也保留
    if (row.raw_data && typeof row.raw_data === 'object') {
      Object.keys(row.raw_data).forEach(k => {
        if (base[k] === undefined) base[k] = row.raw_data[k];
      });
    }

    return base;
  }

  // 將前端 schedule 物件轉換為 Supabase DB row 格式
  function transformClientScheduleToDb(item) {
    const date = item['日期'];
    if (!date) return null;

    const fixedKeys = new Set(['日期', '聚會名稱', '聚會類別', '年度', '季度', '牧師', '題目', '經文', '敬拜曲目', 'leaves', 'hasWarning', 'warningMessage']);
    const assignments = {};
    Object.keys(item).forEach(k => {
      if (!fixedKeys.has(k)) {
        assignments[k] = item[k];
      }
    });

    const [yearPart] = date.split('-');
    const m = parseInt(date.split('-')[1] || '1', 10);
    const q = 'Q' + (Math.floor((m - 1) / 3) + 1);

    return {
      date: date,
      meeting_name: item['聚會名稱'] || '',
      meeting_category: item['聚會類別'] || '',
      year: item['年度'] || yearPart,
      quarter: item['季度'] || q,
      preacher: item['牧師'] || '',
      topic: item['題目'] || '',
      scripture: item['經文'] || '',
      songs: item['敬拜曲目'] || '',
      positions_assignment: assignments,
      leaves: Array.isArray(item.leaves) ? item.leaves : [],
      raw_data: item,
      updated_at: new Date().toISOString()
    };
  }

  // ── 核心 API 實作 ──────────────────────────────────────────

  const WorshipSupabaseService = {
    // 1. 取得崗位設定
    async getPositions() {
      const sb = getSupabase();
      if (!sb) throw new Error('Supabase 未初始化');
      const { data, error } = await sb
        .from('worship_positions')
        .select('*')
        .order('sort_order', { ascending: true });

      if (error) throw error;
      return {
        status: 'success',
        data: (data || []).map(p => ({
          positionName: p.position_name,
          personnel: p.personnel || '',
          isRequired: p.is_required || '是'
        }))
      };
    },

    // 2. 儲存崗位設定
    async savePositions(payload) {
      const sb = getSupabase();
      if (!sb) throw new Error('Supabase 未初始化');
      const list = payload.positionsData || [];

      // 先清空再批次寫入
      await sb.from('worship_positions').delete().neq('id', '00000000-0000-0000-0000-000000000000');
      if (list.length > 0) {
        const rows = list.map((p, idx) => ({
          position_name: p.positionName,
          personnel: p.personnel || '',
          is_required: p.isRequired || '是',
          sort_order: idx + 1
        }));
        const { error } = await sb.from('worship_positions').insert(rows);
        if (error) throw error;
      }

      return { status: 'success', message: '位置設定已儲存！' };
    },

    // 3. 取得敬拜團員名單
    async getTeamMembers() {
      const sb = getSupabase();
      if (!sb) throw new Error('Supabase 未初始化');
      const { data, error } = await sb
        .from('worship_team_members')
        .select('*')
        .order('name', { ascending: true });

      if (error) throw error;
      return {
        status: 'success',
        data: (data || []).map(m => ({
          name: m.name,
          uid: m.uid || '',
          status: m.status || '正式',
          joinDate: m.join_date || ''
        }))
      };
    },

    // 4. 儲存敬拜團員名單
    async saveTeamMembers(payload) {
      const sb = getSupabase();
      if (!sb) throw new Error('Supabase 未初始化');
      const members = payload.members || [];

      await sb.from('worship_team_members').delete().neq('id', '00000000-0000-0000-0000-000000000000');
      if (members.length > 0) {
        const rows = members.map(m => ({
          name: m.name,
          uid: m.uid || '',
          status: m.status || '正式',
          join_date: m.joinDate || new Date().toISOString()
        }));
        const { error } = await sb.from('worship_team_members').insert(rows);
        if (error) throw error;
      }

      return { status: 'success', message: '敬拜團員名單已儲存！' };
    },

    // 5. 取得季度排班 (支援冷熱分流)
    async getSchedule(payload) {
      const year = payload && payload.year;
      const quarter = payload && payload.quarter;

      // 冷資料判定：若早於閾值年份，走 GAS 歷史存檔
      if (!isHotYear(year)) {
        console.log(`[Worship] 查詢 ${year} ${quarter} 為久遠歷史資料，切換至 GAS 讀取...`);
        return await window.churchAPI('getSchedule', payload);
      }

      // 熱資料：直接讀 Supabase (<50ms)
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('getSchedule', payload);

      const { data, error } = await sb
        .from('worship_schedules')
        .select('*')
        .eq('year', String(year))
        .eq('quarter', String(quarter))
        .order('date', { ascending: true });

      if (error) throw error;
      return {
        status: 'success',
        data: (data || []).map(transformDbScheduleToClient)
      };
    },

    // 6. 依日期區間取得排班 (支援冷熱分流)
    async getScheduleByDateRange(payload) {
      const start = payload && payload.startDate;
      const end = payload && payload.endDate;

      if (!isHotDate(start)) {
        console.log(`[Worship] 查詢區間 ${start}~${end} 含久遠歷史，切換至 GAS 讀取...`);
        return await window.churchAPI('getScheduleByDateRange', payload);
      }

      const sb = getSupabase();
      if (!sb) return await window.churchAPI('getScheduleByDateRange', payload);

      let query = sb.from('worship_schedules').select('*');
      if (start) query = query.gte('date', start);
      if (end) query = query.lte('date', end);
      query = query.order('date', { ascending: true });

      const { data, error } = await query;
      if (error) throw error;
      return {
        status: 'success',
        data: (data || []).map(transformDbScheduleToClient)
      };
    },

    // 7. 儲存排班表 (熱響應寫入)
    async saveSchedule(payload) {
      const sb = getSupabase();
      if (!sb) throw new Error('Supabase 未初始化');

      const items = payload.scheduleData || [];
      if (items.length === 0) return { status: 'success', message: '無資料需要儲存' };

      const dbRows = items.map(transformClientScheduleToDb).filter(Boolean);
      const { error } = await sb
        .from('worship_schedules')
        .upsert(dbRows, { onConflict: 'date' });

      if (error) throw error;
      return { status: 'success', message: '服事表儲存成功！' };
    },

    // 8. 取得曲目 (轉呼叫 getSchedule / getScheduleByDateRange)
    async getSongs(payload) {
      if (payload && payload.startDate) {
        return await this.getScheduleByDateRange(payload);
      }
      return await this.getSchedule(payload);
    },

    // 9. 儲存曲目
    async saveSongs(payload) {
      const sb = getSupabase();
      if (!sb) throw new Error('Supabase 未初始化');

      const songsData = payload.songsData || [];
      if (songsData.length === 0) return { status: 'success', message: '無資料' };

      for (const item of songsData) {
        const date = item['日期'];
        const songs = item['敬拜曲目'] || '';
        if (date) {
          await sb
            .from('worship_schedules')
            .update({ songs: songs, updated_at: new Date().toISOString() })
            .eq('date', date);
        }
      }

      return { status: 'success', message: '曲目已成功儲存！' };
    },

    // 10. 行事曆連結設定
    async getCalendarLinkConfig() {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('getCalendarLinkConfig', {});

      const { data } = await sb
        .from('worship_calendar_links')
        .select('*')
        .eq('id', 'default')
        .single();

      // 從 GAS 獲取可用的 sermonSubTypes（跨系統對齊）
      let subTypes = [];
      try {
        const gasRes = await window.churchAPI('getCalendarLinkConfig', {});
        if (gasRes && gasRes.status === 'success' && gasRes.data) {
          subTypes = gasRes.data.sermonSubTypes || [];
        }
      } catch (e) {}

      const defaultSub = (data && data.default_sermon_subtype_id) || '';
      const overrides = (data && data.overrides) || {};

      return {
        status: 'success',
        data: {
          defaultSermonSubTypeId: defaultSub,
          overrides: overrides,
          sermonSubTypes: subTypes,
          calendarReachable: subTypes.length > 0,
          defaultIsValid: subTypes.some(t => t.typeId === defaultSub)
        }
      };
    },

    async setDefaultSermonSubType(payload) {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('setDefaultSermonSubType', payload);

      const typeId = payload.typeId || '';
      const { error } = await sb
        .from('worship_calendar_links')
        .upsert({ id: 'default', default_sermon_subtype_id: typeId, updated_at: new Date().toISOString() });

      if (error) throw error;
      return { status: 'success', message: '預設講道子類型已儲存' };
    },

    async setDateOverride(payload) {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('setDateOverride', payload);

      const { date, typeId } = payload;
      const { data } = await sb.from('worship_calendar_links').select('overrides').eq('id', 'default').single();
      const overrides = (data && data.overrides) || {};

      if (typeId) overrides[date] = typeId;
      else delete overrides[date];

      const { error } = await sb
        .from('worship_calendar_links')
        .upsert({ id: 'default', overrides: overrides, updated_at: new Date().toISOString() });

      if (error) throw error;
      return { status: 'success', message: '日期覆寫已更新' };
    }
  };

  window.WorshipSupabaseService = WorshipSupabaseService;
})(window);
