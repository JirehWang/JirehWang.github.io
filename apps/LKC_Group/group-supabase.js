// ⚡ apps/LKC_Group/group-supabase.js
// 小組點名與名冊系統 Supabase 本地熱響應服務模組 (<50ms)
// 包含小組登入驗證、組員名冊讀取、每週點名送出、名冊拖曳排序、週報與統計中心

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

  const GroupSupabaseService = {
    // ── 1. 取得小組清單 (getGroups) ──────────────────────────────
    async getGroups(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const { data, error } = await sb
        .from('groups')
        .select('*')
        .order('name');

      if (error) throw error;

      const groups = (data || []).map(g => ({
        uuid: g.uuid,
        name: g.name,
        code: g.code,
        encryptedCode: g.encrypted_code || g.code,
        status: g.status || '顯示',
        type: g.group_type || '一般小組',
        groupType: g.group_type || '一般小組',
        associatedGroup: g.associated_group || '',
        districtUuid: g.district_uuid || '',
        clusterUuid: g.cluster_uuid || '',
        date: g.date || ''
      }));

      return { success: true, groups };
    },

    // ── 2. 代碼反查小組 (findGroupByCode) ────────────────────────
    async findGroupByCode(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const code = String(payload.groupCode || payload.code || '').trim().toUpperCase();
      if (!code) return { success: false, message: '請提供小組代碼' };

      const { data, error } = await sb
        .from('groups')
        .select('*');

      if (error) throw error;

      const match = (data || []).find(g => 
        (g.code && g.code.toUpperCase() === code) || 
        (g.encrypted_code && g.encrypted_code === code)
      );

      if (!match) {
        return { success: false, message: '查無此小組代碼' };
      }

      return {
        success: true,
        groupName: match.name,
        encryptedCode: match.encrypted_code || match.code,
        groupType: match.group_type || '一般小組'
      };
    },

    // ── 3. 驗證小組代碼 (verifyGroup) ───────────────────────────
    async verifyGroup(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();
      const code = String(payload.groupCode || payload.code || '').trim().toUpperCase();

      const { data, error } = await sb
        .from('groups')
        .select('*')
        .eq('name', groupName)
        .maybeSingle();

      if (error || !data) return { success: false, message: '查無此小組' };

      const isMatch = (data.code && data.code.toUpperCase() === code) || 
                      (data.encrypted_code && data.encrypted_code === code);

      if (!isMatch) {
        return { success: false, message: '小組代碼錯誤' };
      }

      return {
        success: true,
        encryptedCode: data.encrypted_code || data.code
      };
    },

    // ── 4. 檢查小組狀態與名單 (checkGroupStatus) ──────────────────
    async checkGroupStatus(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();

      const [groupRes, membersRes] = await Promise.all([
        sb.from('groups').select('*').eq('name', groupName).maybeSingle(),
        sb.from('group_members').select('*').eq('group_name', groupName).order('sort_order', { ascending: true })
      ]);

      const group = groupRes.data || {};
      const members = (membersRes.data || []).map(m => ({
        uid: m.uid || '',
        name: m.name || '',
        role: m.role || '小羊',
        nickname: m.nickname || ''
      }));

      return {
        success: true,
        isInitialized: members.length > 0,
        type: group.group_type || '一般小組',
        groupType: group.group_type || '一般小組',
        members
      };
    },

    // ── 5. 每週點名送出 (submitAttendance) ──────────────────────
    async submitAttendance(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();
      const date = String(payload.date || '').slice(0, 10);
      const present = Array.isArray(payload.present) ? payload.present : [];
      const absent = Array.isArray(payload.absent) ? payload.absent : [];
      const newFriends = payload.newFriends || '';
      const offering = Number(payload.offering || 0) || 0;
      const notes = payload.notes || '';

      const record = {
        group_name: groupName,
        date: date,
        present_uids: present,
        absent_uids: absent,
        new_friends: newFriends,
        offering: offering,
        notes: notes,
        updated_at: new Date().toISOString()
      };

      const { error } = await sb
        .from('group_attendance_records')
        .upsert(record, { onConflict: 'group_name,date' });

      if (error) throw error;

      // 背景非同步雙寫同步至 Google Sheets
      setTimeout(() => {
        try {
          if (typeof window.churchAPI_original === 'function') {
            window.churchAPI_original('submitAttendance', payload).catch(e => console.warn('[Group Backup] GAS sync error:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '點名紀錄已成功儲存' };
    },

    // ── 6. 取得點名統計與歷史紀錄 (getStats) ────────────────────
    async getStats(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();

      const [recordsRes, membersRes] = await Promise.all([
        sb.from('group_attendance_records').select('*').eq('group_name', groupName).order('date', { ascending: false }).limit(20),
        sb.from('group_members').select('*').eq('group_name', groupName).order('sort_order', { ascending: true })
      ]);

      const records = recordsRes.data || [];
      const members = membersRes.data || [];
      const memberMap = new Map();
      members.forEach(m => {
        if (m.uid) memberMap.set(m.uid, m.name);
      });

      const formattedRecords = records.map(r => {
        const presentNames = (r.present_uids || []).map(uid => memberMap.get(uid) || uid);
        const absentNames = (r.absent_uids || []).map(uid => memberMap.get(uid) || uid);
        return {
          date: r.date,
          count: presentNames.length,
          presentNames: presentNames.join(', '),
          absentNames: absentNames.join(', '),
          presentUids: r.present_uids || [],
          absentUids: r.absent_uids || [],
          newFriends: r.new_friends || '',
          offering: r.offering || 0,
          notes: r.notes || ''
        };
      });

      return {
        success: true,
        groupName,
        headers: ['日期', '出席人數', '出席名單', '缺席名單', '新朋友', '奉獻金額'],
        records: formattedRecords
      };
    },

    // ── 7. 更新出席紀錄 (updateAttendanceRecord) ────────────────
    async updateAttendanceRecord(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();
      const originalDate = String(payload.originalDate || '').slice(0, 10);
      const newDate = String(payload.newDate || originalDate).slice(0, 10);
      const present = Array.isArray(payload.present) ? payload.present : [];
      const absent = Array.isArray(payload.absent) ? payload.absent : [];
      const newFriends = payload.newFriends || '';

      const { error } = await sb
        .from('group_attendance_records')
        .update({
          date: newDate,
          present_uids: present,
          absent_uids: absent,
          new_friends: newFriends,
          updated_at: new Date().toISOString()
        })
        .eq('group_name', groupName)
        .eq('date', originalDate);

      if (error) throw error;

      setTimeout(() => {
        try {
          if (typeof window.churchAPI_original === 'function') {
            window.churchAPI_original('updateAttendanceRecord', payload).catch(e => console.warn('[Group Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '紀錄已更新' };
    },

    // ── 8. 刪除出席紀錄 (deleteAttendanceRecord) ────────────────
    async deleteAttendanceRecord(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();
      const originalDate = String(payload.originalDate || '').slice(0, 10);

      const { error } = await sb
        .from('group_attendance_records')
        .delete()
        .eq('group_name', groupName)
        .eq('date', originalDate);

      if (error) throw error;

      setTimeout(() => {
        try {
          if (typeof window.churchAPI_original === 'function') {
            window.churchAPI_original('deleteAttendanceRecord', payload).catch(e => console.warn('[Group Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '紀錄已刪除' };
    },

    // ── 9. 組員名冊更新 (updateMemberList / initGroup) ───────────
    async updateMemberList(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();
      const members = Array.isArray(payload.members) ? payload.members : [];

      // 刪除原組員並重新寫入排定順序的組員名單
      await sb.from('group_members').delete().eq('group_name', groupName);

      const rows = members.map((m, idx) => ({
        group_name: groupName,
        uid: m.uid || '',
        name: m.name || '',
        role: m.role || '小羊',
        nickname: m.nickname || '',
        sort_order: idx + 1,
        updated_at: new Date().toISOString()
      }));

      if (rows.length > 0) {
        const { error } = await sb.from('group_members').insert(rows);
        if (error) throw error;
      }

      setTimeout(() => {
        try {
          if (typeof window.churchAPI_original === 'function') {
            window.churchAPI_original('updateMemberList', payload).catch(e => console.warn('[Group Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '小組名單已更新' };
    },

    initGroup(payload) {
      return this.updateMemberList(payload);
    },

    // ── 10. 會友大名單快搜建議 (getMemberSuggestions) ─────────────
    async getMemberSuggestions(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const { data, error } = await sb
        .from('church_members')
        .select('uid, name, phone, group_name')
        .order('name');

      if (error) throw error;

      return {
        success: true,
        members: (data || []).map(m => ({
          uid: m.uid,
          name: m.name,
          phone: m.phone,
          groupName: m.group_name
        }))
      };
    },

    // ── 11. 本週出席人數週報 (getWeeklyReport) ───────────────────
    async getWeeklyReport(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      // 取得所有小組以及最近一週各組點名紀錄
      const [groupsRes, recordsRes] = await Promise.all([
        sb.from('groups').select('name, group_type, status').order('name'),
        sb.from('group_attendance_records').select('*').order('date', { ascending: false }).limit(100)
      ]);

      const groups = groupsRes.data || [];
      const records = recordsRes.data || [];

      // 找出各組最新一筆紀錄
      const latestByGroup = new Map();
      records.forEach(r => {
        if (!latestByGroup.has(r.group_name)) {
          latestByGroup.set(r.group_name, r);
        }
      });

      const report = groups.map(g => {
        const rec = latestByGroup.get(g.name);
        return {
          groupName: g.name,
          date: rec ? rec.date : '',
          presentCount: rec ? (rec.present_uids || []).length : 0,
          newFriendsCount: rec ? (rec.new_friends ? rec.new_friends.split(',').length : 0) : 0,
          offering: rec ? rec.offering || 0 : 0
        };
      });

      return { success: true, report };
    }
  };

  // 🎯 自動劫持 / 增強 window.churchAPI
  function setupGroupRouter() {
    if (typeof window.churchAPI === 'function' && !window.churchAPI_original) {
      window.churchAPI_original = window.churchAPI;
      window.churchAPI = async function(action, data = {}) {
        if (GroupSupabaseService[action] && typeof GroupSupabaseService[action] === 'function') {
          try {
            const res = await GroupSupabaseService[action](data);
            if (res !== null) return res;
          } catch (err) {
            console.warn(`[GroupSupabase] Action ${action} local handling error, falling back to GAS:`, err);
          }
        }
        return await window.churchAPI_original(action, data);
      };
    }
  }

  window.GroupSupabaseService = GroupSupabaseService;
  setupGroupRouter();
  window.addEventListener('DOMContentLoaded', setupGroupRouter);
})();
