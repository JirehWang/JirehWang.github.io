// ⚡ apps/LKC_MinistrySchedule/ministry-supabase.js
// 事工排班管理系統 Supabase 本地熱響應服務模組 (<50ms)
// 包含排班分頁配置、動態欄位定義、季度排班二維矩陣讀寫、同工名單快搜與佈告欄總表

(function() {
  function getSupabase() {
    return window._supabase || (typeof supabase !== 'undefined' && window.SUPABASE_CONFIG ? 
      supabase.createClient(window.SUPABASE_CONFIG.url, window.SUPABASE_CONFIG.anonKey) : null);
  }

  const MinistrySupabaseService = {
    // ── 1. 取得排班分頁與二維矩陣配置 (getPageConfig) ──────────────
    async getPageConfig(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const pageId = String(payload.id || payload.groupCode || '').trim();
      const pageName = String(payload.groupName || payload.name || '').trim();

      let pageQuery = sb.from('ministry_pages').select('*');
      if (pageId) {
        pageQuery = pageQuery.or(`page_id.eq.${pageId},uuid.eq.${pageId},page_name.eq.${pageId}`);
      } else if (pageName) {
        pageQuery = pageQuery.eq('page_name', pageName);
      }

      const { data: pageList, error: pageErr } = await pageQuery;
      if (pageErr) throw pageErr;

      let page = (pageList && pageList.length > 0) ? pageList[0] : null;

      // 若未找到但有指定 pageName 或 pageId，檢查 groups 表
      if (!page && (pageName || pageId)) {
        const { data: group } = await sb.from('groups')
          .select('*')
          .or(`code.eq.${pageId},name.eq.${pageName || pageId}`)
          .maybeSingle();

        if (group) {
          page = {
            uuid: group.uuid,
            page_name: group.name,
            page_id: group.code,
            template_type: group.group_type === '事工' ? 'ministry' : 'gathering',
            status: group.status || '顯示',
            schedule_target: '',
            custom_members: []
          };
        }
      }

      if (!page) {
        return { status: 'error', message: '查無此排班分頁' };
      }

      const effectivePageName = page.page_name;

      // 平行查詢欄位、排班矩陣與組員名冊
      const [fieldsRes, schedulesRes, membersRes] = await Promise.all([
        sb.from('ministry_fields').select('*').eq('page_name', effectivePageName).order('sort_order', { ascending: true }),
        sb.from('ministry_schedules').select('*').eq('page_name', effectivePageName).order('date', { ascending: true }),
        sb.from('group_members').select('*').eq('group_name', effectivePageName).order('sort_order', { ascending: true })
      ]);

      const fields = (fieldsRes.data || []).map(f => ({
        name: f.field_name,
        key: f.field_key || f.field_name,
        enabled: f.enabled !== false,
        custom: f.is_custom === true,
        useMemberList: f.use_member_list !== false
      }));

      const events = (schedulesRes.data || []).map(s => ({
        date: String(s.date).slice(0, 10),
        ...(s.assignments || {})
      }));

      const groupMembers = membersRes.data || [];
      const coreMembers = groupMembers.filter(m => m.role === '核心同工' || m.role === '福長').map(m => m.name);
      const generalMembers = groupMembers.map(m => m.name);

      return {
        status: 'success',
        data: {
          id: page.page_id || page.uuid,
          groupName: effectivePageName,
          pageName: effectivePageName,
          templateType: page.template_type || 'gathering',
          status: page.status || '顯示',
          prompt: page.prompt || '',
          scheduleTarget: page.schedule_target || '',
          customMembers: page.custom_members || [],
          fields: fields.length > 0 ? fields : undefined,
          events: events,
          members: generalMembers,
          coreMembers: coreMembers,
          generalMembers: generalMembers
        }
      };
    },

    // ── 2. 取得所有分頁清單 (getGroups) ──────────────────────────
    async getGroups(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const { data, error } = await sb.from('ministry_pages').select('*').order('page_name');
      if (error) throw error;

      const groups = (data || []).map(p => ({
        uuid: p.uuid,
        name: p.page_name,
        id: p.page_id || p.uuid,
        templateType: p.template_type || 'gathering',
        status: p.status || '顯示',
        scheduleTarget: p.schedule_target || ''
      }));

      return { status: 'success', data: { groups } };
    },

    // ── 3. 儲存排班二維表格資料 (saveSheetData) ───────────────────
    async saveSheetData(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || payload.pageName || '').trim();
      const matrix = Array.isArray(payload.matrix) ? payload.matrix : [];

      // matrix: [ [ '日期', '破冰', '敬拜', ... ], [ '2026-09-06', '小明', '小華' ], ... ]
      if (matrix.length > 1) {
        const headers = matrix[0];
        const dateIdx = 0;

        const scheduleRows = [];
        for (let i = 1; i < matrix.length; i++) {
          const row = matrix[i];
          const rawDate = row[dateIdx];
          if (!rawDate) continue;
          const dateStr = String(rawDate).slice(0, 10);

          const assignments = {};
          for (let col = 1; col < headers.length; col++) {
            const h = headers[col];
            if (h) {
              assignments[h] = row[col] || '';
            }
          }

          scheduleRows.push({
            page_name: groupName,
            date: dateStr,
            assignments: assignments,
            updated_at: new Date().toISOString()
          });
        }

        for (const row of scheduleRows) {
          await sb.from('ministry_schedules').upsert(row, { onConflict: 'page_name,date' });
        }
      }

      // 背景非同步雙寫同步至 Google Sheets
      setTimeout(() => {
        try {
          if (typeof window.churchAPI_original === 'function') {
            window.churchAPI_original('ministry_saveSheetData', payload).catch(e => console.warn('[Ministry Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { status: 'success', data: { message: '排班表已成功儲存' } };
    },

    // ── 4. 儲存欄位配置 (savePageFieldConfig) ────────────────────
    async savePageFieldConfig(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const pageId = String(payload.id || '').trim();
      const pageFieldConfig = payload.pageFieldConfig || {};
      const fields = Array.isArray(pageFieldConfig.fields) ? pageFieldConfig.fields : [];

      const { data: page } = await sb.from('ministry_pages')
        .select('page_name')
        .or(`page_id.eq.${pageId},uuid.eq.${pageId},page_name.eq.${pageId}`)
        .maybeSingle();

      const pageName = page ? page.page_name : pageId;

      if (pageName && fields.length > 0) {
        // 先清除原欄位再寫入
        await sb.from('ministry_fields').delete().eq('page_name', pageName);

        const fieldRows = fields.map((f, idx) => ({
          page_name: pageName,
          field_name: f.name || f.key,
          field_key: f.key || f.name,
          enabled: f.enabled !== false,
          is_custom: f.custom === true,
          use_member_list: f.useMemberList !== false,
          sort_order: idx + 1,
          updated_at: new Date().toISOString()
        }));

        await sb.from('ministry_fields').insert(fieldRows);
      }

      setTimeout(() => {
        try {
          if (typeof window.churchAPI_original === 'function') {
            window.churchAPI_original('ministry_savePageFieldConfig', payload).catch(e => console.warn('[Ministry Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { status: 'success', data: { message: '欄位配置已更新' } };
    },

    // ── 5. 會友大名單建議 (getMemberSuggestions) ──────────────────
    async getMemberSuggestions(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const { data, error } = await sb
        .from('church_members')
        .select('uid, name, phone, group_name')
        .order('name');

      if (error) throw error;

      return {
        status: 'success',
        data: {
          members: (data || []).map(m => ({
            uid: m.uid,
            name: m.name,
            phone: m.phone,
            groupName: m.group_name
          }))
        }
      };
    },

    // ── 6. 聚合未來 7 天服事總表 (getAggregatedReport) ────────────
    async getAggregatedReport(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const todayStr = new Date().toISOString().slice(0, 10);

      const { data: schedules, error } = await sb
        .from('ministry_schedules')
        .select('*')
        .gte('date', todayStr)
        .order('date', { ascending: true })
        .limit(50);

      if (error) throw error;

      const report = (schedules || []).map(s => ({
        groupName: s.page_name,
        date: String(s.date).slice(0, 10),
        assignments: s.assignments || {}
      }));

      return { status: 'success', data: { reports: report } };
    }
  };

  // 🎯 自動劫持 / 增強 window.churchAPI（支援 ministry_ 前綴）
  function setupMinistryRouter() {
    if (typeof window.churchAPI === 'function' && !window.churchAPI_original) {
      window.churchAPI_original = window.churchAPI;
      window.churchAPI = async function(action, data = {}) {
        const cleanAction = action.replace(/^ministry_/, '');
        if (MinistrySupabaseService[cleanAction] && typeof MinistrySupabaseService[cleanAction] === 'function') {
          try {
            const res = await MinistrySupabaseService[cleanAction](data);
            if (res !== null) return res;
          } catch (err) {
            console.warn(`[MinistrySupabase] Action ${action} handling error, falling back to GAS:`, err);
          }
        }
        return await window.churchAPI_original(action, data);
      };
    }
  }

  window.MinistrySupabaseService = MinistrySupabaseService;
  setupMinistryRouter();
  window.addEventListener('DOMContentLoaded', setupMinistryRouter);
})();
