// ⚡ apps/LKC_NewFamily/new-family-supabase.js
// 新家人管理系統 Supabase 本地熱響應服務模組 (<50ms)
// 包含追蹤中/已結案個案讀取、留名卡登錄、個案更新、轉會友標記與結案操作

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

  function mapDbToCase(r, idx) {
    return {
      id: r.id,
      rowNumber: r.row_number || (idx + 2),
      '姓名': r.name || '',
      '聚會別': r.service_type || '',
      '性別': r.gender || '',
      '職業': r.occupation || '',
      '年齡': r.age_group || '',
      '是否曾接觸教會': r.contacted_church_before || '',
      '來訪原因': r.visit_reason || '',
      '表單號': r.form_number ? Number(r.form_number) || r.form_number : '',
      '關懷同工': r.assigned_staff || '',
      '地址': r.address || '',
      '市話': r.tel || '',
      '手機': r.phone || '',
      '首次來訪日': r.first_visit_date ? String(r.first_visit_date).slice(0, 10) : '',
      '結案日期': r.closed_date ? String(r.closed_date).slice(0, 10) : '',
      '落戶狀態': r.settlement_status || '',
      '邀約人': r.inviter || '',
      '備註': r.notes || '',
      '會友狀態': r.member_status || '',
      '點名編號': r.member_code || '',
      '現行小組': r.current_group || ''
    };
  }

  const NewFamilySupabaseService = {
    // ── 1. 讀取追蹤中案件 (getTrackingCases) ──────────────────────
    async getTrackingCases() {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('getTrackingCases');

      const { data, error } = await sb
        .from('new_family_cases')
        .select('*')
        .eq('status', 'tracking')
        .order('form_number', { ascending: false });

      if (error) throw error;
      return { success: true, data: (data || []).map(mapDbToCase) };
    },

    // ── 2. 讀取已結案案件 (getClosedCases) ───────────────────────
    async getClosedCases() {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('getClosedCases');

      const { data, error } = await sb
        .from('new_family_cases')
        .select('*')
        .eq('status', 'closed')
        .order('form_number', { ascending: false });

      if (error) throw error;
      return { success: true, data: (data || []).map(mapDbToCase) };
    },

    // ── 3. 新增新家人留名卡 (submitNewFamily) ───────────────────
    async submitNewFamily(payload) {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('submitNewFamily', payload);

      const now = new Date();
      const datePrefix = now.toISOString().slice(0, 10).replace(/-/g, '');
      const randomSuffix = Math.floor(100 + Math.random() * 900);
      const formNumber = payload['表單號'] || `${datePrefix}${randomSuffix}`;

      const newRecord = {
        form_number: String(formNumber),
        name: String(payload['姓名'] || '').trim(),
        gender: String(payload['性別'] || '').trim(),
        service_type: String(payload['聚會別'] || '').trim(),
        occupation: String(payload['職業'] || '').trim(),
        age_group: String(payload['年齡'] || '').trim(),
        contacted_church_before: String(payload['是否曾接觸教會'] || '').trim(),
        visit_reason: String(payload['來訪原因'] || '').trim(),
        assigned_staff: String(payload['關懷同工'] || '').trim(),
        address: String(payload['地址'] || '').trim(),
        tel: String(payload['市話'] || '').trim(),
        phone: String(payload['手機'] || '').trim(),
        first_visit_date: payload['首次來訪日'] ? String(payload['首次來訪日']).slice(0, 10) : now.toISOString().slice(0, 10),
        settlement_status: String(payload['落戶狀態'] || '').trim(),
        inviter: String(payload['邀約人'] || '').trim(),
        notes: String(payload['備註'] || '').trim(),
        member_status: String(payload['會友狀態'] || '').trim(),
        member_code: String(payload['點名編號'] || '').trim(),
        current_group: String(payload['現行小組'] || '').trim(),
        status: 'tracking',
        updated_at: new Date().toISOString()
      };

      const { error } = await sb
        .from('new_family_cases')
        .upsert(newRecord, { onConflict: 'form_number' });

      if (error) throw error;

      // 背景雙寫至 Google Sheets
      setTimeout(() => {
        try {
          if (typeof window.churchAPI === 'function') {
            window.churchAPI('submitNewFamily', payload).catch(e => console.warn('[NewFamily Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '新家人登錄成功', formNumber };
    },

    // ── 4. 更新個案 (updateTrackingCase / updateClosedCase) ─────
    async updateTrackingCase(payload) {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('updateTrackingCase', payload);

      const vals = payload.values || payload;
      const id = payload.id;
      const formNumber = String(payload['表單號'] || payload.formNumber || payload.form_number || (vals && (vals.formNumber || vals['表單號'])) || '').trim();
      const rowNumber = Number(payload.rowNumber || payload.row_number || (vals && vals.rowNumber) || 0);
      const name = String(payload.name || payload['姓名'] || (vals && vals['姓名']) || '').trim();

      const updateData = {
        updated_at: new Date().toISOString()
      };

      const fieldMap = {
        '姓名': 'name',
        '性別': 'gender',
        '聚會別': 'service_type',
        '職業': 'occupation',
        '年齡': 'age_group',
        '是否曾接觸教會': 'contacted_church_before',
        '來訪原因': 'visit_reason',
        '關懷同工': 'assigned_staff',
        '地址': 'address',
        '市話': 'tel',
        '手機': 'phone',
        '首次來訪日': 'first_visit_date',
        '結案日期': 'closed_date',
        '落戶狀態': 'settlement_status',
        '邀約人': 'inviter',
        '備註': 'notes',
        '會友狀態': 'member_status',
        '點名編號': 'member_code',
        '現行小組': 'current_group'
      };

      Object.entries(fieldMap).forEach(([formKey, dbCol]) => {
        if (vals[formKey] !== undefined) updateData[dbCol] = vals[formKey];
      });

      let query = sb.from('new_family_cases').update(updateData);

      if (id) {
        query = query.eq('id', id);
      } else if (formNumber) {
        query = query.eq('form_number', formNumber);
      } else if (name) {
        query = query.eq('name', name);
      } else if (rowNumber) {
        query = query.eq('row_number', rowNumber);
      }

      const { error } = await query;
      if (error) throw error;

      // 背景雙寫
      setTimeout(() => {
        try {
          if (typeof window.churchAPI === 'function') {
            window.churchAPI('updateTrackingCase', payload).catch(e => console.warn('[NewFamily Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '更新成功' };
    },

    async updateClosedCase(payload) {
      return this.updateTrackingCase(payload);
    },

    // ── 5. 刪除追蹤中案件 (deleteTrackingCase) ───────────────────
    async deleteTrackingCase(payload) {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('deleteTrackingCase', payload);

      const id = payload.id;
      const formNumber = String(payload.formNumber || payload.form_number || payload['表單號'] || '').trim();
      const rowNumber = Number(payload.rowNumber || 0);
      const name = String(payload.name || payload['姓名'] || '').trim();

      let query = sb.from('new_family_cases').delete();
      if (id) {
        query = query.eq('id', id);
      } else if (formNumber) {
        query = query.eq('form_number', formNumber);
      } else if (name) {
        query = query.eq('name', name);
      } else if (rowNumber) {
        query = query.eq('row_number', rowNumber);
      } else {
        throw new Error('未提供欲刪除案件之識別資料');
      }

      const { error } = await query;
      if (error) throw error;

      setTimeout(() => {
        try {
          if (typeof window.churchAPI === 'function') {
            window.churchAPI('deleteTrackingCase', payload).catch(e => console.warn('[NewFamily Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '刪除成功' };
    },

    // ── 6. 批次標記會友狀態 (markTrackingMemberStatuses) ─────────
    async markTrackingMemberStatuses(payload) {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('markTrackingMemberStatuses', payload);

      const items = Array.isArray(payload.items) ? payload.items : (Array.isArray(payload) ? payload : []);
      for (const item of items) {
        const updateData = {
          updated_at: new Date().toISOString()
        };
        if (item.memberCode) updateData.member_code = item.memberCode;
        if (item.sundayGroup) updateData.current_group = item.sundayGroup;
        if (item.status) updateData.member_status = item.status;

        let query = sb.from('new_family_cases').update(updateData);
        if (item.id) query = query.eq('id', item.id);
        else if (item.formNumber) query = query.eq('form_number', String(item.formNumber));
        else if (item.name) query = query.eq('name', item.name);

        await query;
      }

      setTimeout(() => {
        try {
          if (typeof window.churchAPI === 'function') {
            window.churchAPI('markTrackingMemberStatuses', payload).catch(e => console.warn('[NewFamily Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '標記完成' };
    },

    // ── 7. 批次結案 (closeCases) ──────────────────────────────
    async closeCases(payload) {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('closeCases', payload);

      const rowNumbers = Array.isArray(payload.rowNumbers) ? payload.rowNumbers : [];
      const formNumbers = Array.isArray(payload.formNumbers) ? payload.formNumbers : [];
      const ids = Array.isArray(payload.ids) ? payload.ids : [];
      const items = Array.isArray(payload.items) ? payload.items : [];

      const allFormNumbers = [...formNumbers];
      const allIds = [...ids];
      items.forEach(it => {
        if (it.formNumber || it['表單號']) allFormNumbers.push(String(it.formNumber || it['表單號']));
        if (it.id) allIds.push(it.id);
      });

      const todayStr = new Date().toISOString().slice(0, 10);

      let query = sb.from('new_family_cases').update({
        status: 'closed',
        closed_date: todayStr,
        updated_at: new Date().toISOString()
      });

      if (allIds.length > 0) {
        query = query.in('id', allIds);
      } else if (allFormNumbers.length > 0) {
        query = query.in('form_number', allFormNumbers);
      } else if (rowNumbers.length > 0) {
        query = query.in('row_number', rowNumbers);
      }

      const { error } = await query;
      if (error) throw error;

      setTimeout(() => {
        try {
          if (typeof window.churchAPI === 'function') {
            window.churchAPI('closeCases', payload).catch(e => console.warn('[NewFamily Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '已成功結案' };
    },

    // ── 8. 小組與教區選單 (getDistrictsAndClusters / getGroups) ──
    async getDistrictsAndClusters() {
      const sb = getSupabase();
      if (!sb) return null;

      const { data: pages } = await sb.from('ministry_pages').select('page_name, template_type');
      const groups = (pages || []).map(p => ({
        name: p.page_name,
        cluster: p.template_type
      }));
      return { success: true, clusters: groups, groups: groups };
    },

    async getGroups() {
      const sb = getSupabase();
      if (!sb) return null;

      const { data: pages } = await sb.from('ministry_pages').select('page_name, template_type');
      const groups = (pages || []).map(p => ({
        name: p.page_name,
        cluster: p.template_type
      }));
      return { success: true, groups: groups };
    }
  };

  // 🎯 自動劫持 / 增強 window.churchAPI（支援 newfamily 路由）
  function setupNewFamilyRouter() {
    if (typeof window.churchAPI === 'function' && !window.churchAPI_original_nf) {
      window.churchAPI_original_nf = window.churchAPI;
      window.churchAPI = async function(action, data = {}) {
        if (NewFamilySupabaseService[action] && typeof NewFamilySupabaseService[action] === 'function') {
          try {
            const res = await NewFamilySupabaseService[action](data);
            if (res !== null) return res;
          } catch (err) {
            console.warn(`[NewFamilySupabase] Action ${action} handling error, falling back to GAS:`, err);
          }
        }
        return await window.churchAPI_original_nf(action, data);
      };
    }
  }

  window.NewFamilySupabaseService = NewFamilySupabaseService;
  setupNewFamilyRouter();
  window.addEventListener('DOMContentLoaded', setupNewFamilyRouter);
})();
