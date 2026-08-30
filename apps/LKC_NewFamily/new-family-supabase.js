// ⚡ apps/LKC_NewFamily/new-family-supabase.js
// 新家人管理系統 Supabase 本地熱響應服務模組 (<50ms)
// 包含追蹤中/已結案個案讀取、留名卡登錄、個案更新、轉會友標記與結案操作

(function() {
  function getSupabase() {
    return window._supabase || (typeof supabase !== 'undefined' && window.SUPABASE_CONFIG ? 
      supabase.createClient(window.SUPABASE_CONFIG.url, window.SUPABASE_CONFIG.anonKey) : null);
  }

  function mapDbToCase(r, idx) {
    return {
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
        .order('first_visit_date', { ascending: false });

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
        .order('closed_date', { ascending: false });

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

    // ── 4. 更新追蹤中案件 (updateTrackingCase) ───────────────────
    async updateTrackingCase(payload) {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('updateTrackingCase', payload);

      const formNumber = String(payload['表單號'] || payload.formNumber || '').trim();
      const rowNumber = Number(payload.rowNumber || payload.row_number || 0);

      let query = sb.from('new_family_cases').update({
        name: payload['姓名'],
        gender: payload['性別'],
        service_type: payload['聚會別'],
        occupation: payload['職業'],
        age_group: payload['年齡'],
        contacted_church_before: payload['是否曾接觸教會'],
        visit_reason: payload['來訪原因'],
        assigned_staff: payload['關懷同工'],
        address: payload['地址'],
        tel: payload['市話'],
        phone: payload['手機'],
        first_visit_date: payload['首次來訪日'],
        settlement_status: payload['落戶狀態'],
        inviter: payload['邀約人'],
        notes: payload['備註'],
        member_status: payload['會友狀態'],
        member_code: payload['點名編號'],
        current_group: payload['現行小組'],
        updated_at: new Date().toISOString()
      });

      if (formNumber) {
        query = query.eq('form_number', formNumber);
      } else if (rowNumber) {
        query = query.eq('row_number', rowNumber);
      } else {
        query = query.eq('name', payload['姓名']);
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

    // ── 5. 刪除追蹤中案件 (deleteTrackingCase) ───────────────────
    async deleteTrackingCase(payload) {
      const sb = getSupabase();
      if (!sb) return await window.churchAPI('deleteTrackingCase', payload);

      const formNumber = String(payload.formNumber || payload['表單號'] || '').trim();
      const rowNumber = Number(payload.rowNumber || 0);

      let query = sb.from('new_family_cases').delete();
      if (formNumber) query = query.eq('form_number', formNumber);
      else if (rowNumber) query = query.eq('row_number', rowNumber);

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
        if (item.formNumber) query = query.eq('form_number', String(item.formNumber));
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
      const todayStr = new Date().toISOString().slice(0, 10);

      let query = sb.from('new_family_cases').update({
        status: 'closed',
        closed_date: todayStr,
        updated_at: new Date().toISOString()
      });

      if (formNumbers.length > 0) {
        query = query.in('form_number', formNumbers.map(String));
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
    }
  };

  window.NewFamilySupabaseService = NewFamilySupabaseService;
})();
