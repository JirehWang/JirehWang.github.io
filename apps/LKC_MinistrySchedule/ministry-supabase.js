// ⚡ apps/LKC_MinistrySchedule/ministry-supabase.js
// 事工排班管理系統 Supabase 本地熱響應服務模組 (<50ms)
// 包含排班分頁配置、動態欄位定義、季度排班二維矩陣讀寫、同工名單快搜與佈告欄總表

(function() {
  function getSupabase() {
    if (window._supabase) return window._supabase;
    const config = window._SUPABASE_CONFIG || window.SUPABASE_CONFIG;
    const create = (window.supabase && window.supabase.createClient) || (typeof supabase !== 'undefined' && supabase.createClient);
    if (config && create) {
      window._supabase = create(config.url, config.anonKey);
      return window._supabase;
    }
    return null;
  }

  const OBFUSCATION_KEY = 'LKC-Secure-2026';
  const ENC_PREFIX = 'enc_';

  function decryptId(str) {
    if (typeof window.decryptGroupCode === 'function') {
      return window.decryptGroupCode(str);
    }
    const safeStr = String(str || '');
    if (!safeStr || safeStr.indexOf(ENC_PREFIX) !== 0) return safeStr;
    try {
      var hex = safeStr.substring(ENC_PREFIX.length);
      var plainText = '';
      for (var i = 0; i < hex.length; i += 2) {
        var charCode = parseInt(hex.substring(i, i + 2), 16);
        var decCharCode = charCode ^ OBFUSCATION_KEY.charCodeAt((i / 2) % OBFUSCATION_KEY.length);
        plainText += String.fromCharCode(decCharCode);
      }
      return plainText;
    } catch (e) {
      return safeStr;
    }
  }

  const MinistrySupabaseService = {
    // ── 1. 取得排班分頁與二維矩陣配置 (getPageConfig) ──────────────
    async getPageConfig(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const rawId = String(payload.id || payload.groupCode || '').trim();
      const pageId = decryptId(rawId).trim();
      const rawName = String(payload.groupName || payload.name || '').trim();
      const pageName = decryptId(rawName).trim();

      let pageQuery = sb.from('ministry_pages').select('*');
      if (pageId && pageName) {
        pageQuery = pageQuery.or(`page_id.eq.${pageId},uuid.eq.${pageId},page_name.eq.${pageName},page_name.eq.${pageId}`);
      } else if (pageId) {
        pageQuery = pageQuery.or(`page_id.eq.${pageId},uuid.eq.${pageId},page_name.eq.${pageId},page_id.eq.${rawId}`);
      } else if (pageName) {
        pageQuery = pageQuery.or(`page_name.eq.${pageName},page_id.eq.${pageName}`);
      }

      const { data: pageList, error: pageErr } = await pageQuery;
      if (pageErr) throw pageErr;

      let page = (pageList && pageList.length > 0) ? pageList[0] : null;

      // 若未找到但有指定 pageName 或 pageId，檢查 groups 表
      if (!page && (pageName || pageId)) {
        const { data: group } = await sb.from('groups')
          .select('*')
          .or(`code.eq.${pageId},name.eq.${pageName || pageId},uuid.eq.${pageId}`)
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
          await sb.from('ministry_pages').upsert(page, { onConflict: 'page_name' });
        }
      }

      // 若仍然未找到且允許 autoCreate
      if (!page && payload.autoCreate && (pageName || pageId)) {
        const newPageName = pageName || pageId;
        const newPageId = pageId || pageName;
        page = {
          uuid: (typeof crypto !== 'undefined' && crypto.randomUUID) ? crypto.randomUUID() : ('page_' + Date.now()),
          page_name: newPageName,
          page_id: newPageId,
          template_type: 'gathering',
          status: '顯示',
          schedule_target: '',
          custom_members: []
        };
        await sb.from('ministry_pages').upsert(page, { onConflict: 'page_name' });
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

      // 建構 2D Matrix (matrix[0] 為標題列，matrix[1..N] 為各日期資料列)
      let headerNames = fields.map(f => f.name);
      if (headerNames.length === 0) {
        headerNames = ['破冰', '敬拜', '話語分享', '主題', '經文', '地點', '套用講道'];
      }
      if (!headerNames.includes('套用講道')) {
        headerNames.push('套用講道');
      }

      const fullHeaders = ['日期', ...headerNames];
      const matrix = [fullHeaders];
      const eventData = {};

      (schedulesRes.data || []).forEach(s => {
        const dateStr = String(s.date).slice(0, 10);
        const assignments = s.assignments || {};
        eventData[dateStr] = assignments;

        const row = [dateStr];
        for (let i = 1; i < fullHeaders.length; i++) {
          const h = fullHeaders[i];
          row.push(assignments[h] || '');
        }
        matrix.push(row);
      });

      const groupMembers = membersRes.data || [];
      const coreMembers = groupMembers.filter(m => m.role === '核心同工' || m.role === '福長').map(m => m.name);
      const generalMembers = groupMembers.map(m => m.name);

      const templateName = page.template_type || '小組聚會表模板';

      return {
        status: 'success',
        data: {
          id: page.page_id || page.uuid,
          groupName: effectivePageName,
          pageName: effectivePageName,
          template: templateName,
          templateType: templateName,
          status: page.status || '顯示',
          prompt: page.prompt || '',
          groupPrompt: page.prompt || '',
          scheduleTarget: page.schedule_target || 'members',
          customMembers: page.custom_members || [],
          pageFieldConfig: {
            fields: fields.length > 0 ? fields : undefined,
            fieldTemplateType: templateName.includes('事工') ? '事工型模板' : '聚會型模板',
            requiredFields: ['日期']
          },
          matrix: matrix,
          eventData: eventData,
          events: (schedulesRes.data || []).map(s => ({
            date: String(s.date).slice(0, 10),
            ...(s.assignments || {})
          })),
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
        template: p.template_type === 'ministry' ? '事工型模板' : (p.template_type || '聚會型模板'),
        templateType: p.template_type === 'ministry' ? '事工型模板' : '聚會型模板',
        status: p.status || '顯示',
        scheduleTarget: p.schedule_target || ''
      }));

      return { status: 'success', data: groups, groups: groups };
    },

    // ── 3. 取得模板清單 (getTemplates) ───────────────────────────
    async getTemplates(payload) {
      const templates = ['聚會型模板', '事工型模板'];
      return { status: 'success', data: templates, templates: templates };
    },

    // ── 4. 取得牧區與群組 (getDistrictsAndClusters) ──────────────
    async getDistrictsAndClusters(payload) {
      return {
        status: 'success',
        data: { districts: [], clusters: [] }
      };
    },

    // ── 5. 儲存排班二維表格資料 (saveSheetData) ───────────────────
    async saveSheetData(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || payload.pageName || '').trim();
      const matrix = Array.isArray(payload.matrix) ? payload.matrix : [];

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

    // ── 6. 儲存欄位配置 (savePageFieldConfig) ────────────────────
    async savePageFieldConfig(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const rawId = String(payload.id || '').trim();
      const pageId = decryptId(rawId).trim();
      const pageFieldConfig = payload.pageFieldConfig || {};
      const fields = Array.isArray(pageFieldConfig.fields) ? pageFieldConfig.fields : [];

      const { data: page } = await sb.from('ministry_pages')
        .select('page_name')
        .or(`page_id.eq.${pageId},uuid.eq.${pageId},page_name.eq.${pageId}`)
        .maybeSingle();

      const pageName = page ? page.page_name : pageId;

      if (pageName && fields.length > 0) {
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

    // ── 7. 切換分頁狀態 (toggleGroupStatus) ──────────────────────
    async toggleGroupStatus(payload) {
      const sb = getSupabase();
      if (!sb) return null;
      const rawId = String(payload.id || '');
      const id = decryptId(rawId).trim();
      const newStatus = payload.status === '顯示' ? '停用' : '顯示';
      await sb.from('ministry_pages')
        .update({ status: newStatus, updated_at: new Date().toISOString() })
        .or(`page_id.eq.${id},uuid.eq.${id},page_name.eq.${id}`);
      return { status: 'success', data: { status: newStatus } };
    },

    // ── 8. 儲存 AI 提示詞 (saveGroupPrompt) ──────────────────────
    async saveGroupPrompt(payload) {
      const sb = getSupabase();
      if (!sb) return null;
      const rawId = String(payload.id || '');
      const id = decryptId(rawId).trim();
      const prompt = payload.prompt || '';
      await sb.from('ministry_pages')
        .update({ prompt: prompt, updated_at: new Date().toISOString() })
        .or(`page_id.eq.${id},uuid.eq.${id},page_name.eq.${id}`);
      return { status: 'success', data: { prompt } };
    },

    // ── 9. 會友大名單建議 (getMemberSuggestions) ──────────────────
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

    // ── 10. 聚合未來 7 天服事總表 (getAggregatedReport) ───────────
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
