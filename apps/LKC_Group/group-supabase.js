// ⚡ apps/LKC_Group/group-supabase.js
// 小組點名與名冊系統 Supabase 本地熱響應服務模組 (<50ms)
// 包含小組登入驗證、組員名冊讀取、每週點名送出、名冊拖曳排序、週報與統計中心

(function() {
  const ADMIN_CODE = 'LK31';
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

  function encryptId(str) {
    if (typeof window.encryptGroupCode === 'function') {
      return window.encryptGroupCode(str);
    }
    const safeStr = String(str || '');
    if (!safeStr || safeStr.indexOf(ENC_PREFIX) === 0) return safeStr;
    var hex = '';
    for (var i = 0; i < safeStr.length; i++) {
      var charCode = safeStr.charCodeAt(i) ^ OBFUSCATION_KEY.charCodeAt(i % OBFUSCATION_KEY.length);
      var hexByte = charCode.toString(16).padStart(2, '0');
      hex += hexByte;
    }
    return ENC_PREFIX + hex;
  }

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
        encryptedCode: g.encrypted_code || encryptId(g.code),
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

      const rawCode = String(payload.groupCode || payload.code || '').trim();
      const code = decryptId(rawCode).trim().toUpperCase();
      if (!code) return { success: false, message: '請提供小組代碼' };

      if (code === ADMIN_CODE) {
        return { success: true, groupName: "ADMIN", isAdmin: true, encryptedCode: encryptId(ADMIN_CODE) };
      }

      const { data, error } = await sb
        .from('groups')
        .select('*');

      if (error) throw error;

      const match = (data || []).find(g => 
        (g.code && g.code.toUpperCase() === code) || 
        (g.encrypted_code && g.encrypted_code === rawCode) ||
        (g.code && g.code.toUpperCase() === rawCode.toUpperCase())
      );

      if (!match) {
        return { success: false, message: '查無此小組代碼' };
      }

      return {
        success: true,
        groupName: match.name,
        encryptedCode: match.encrypted_code || encryptId(match.code),
        groupType: match.group_type || '一般小組'
      };
    },

    // ── 3. 驗證小組代碼 (verifyGroup) ───────────────────────────
    async verifyGroup(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();
      const rawCode = String(payload.groupCode || payload.code || '').trim();
      const code = decryptId(rawCode).trim().toUpperCase();

      if (code === ADMIN_CODE) {
        return { success: true, message: '管理員授權', isAdmin: true, encryptedCode: encryptId(ADMIN_CODE) };
      }

      const res = await this.findGroupByCode(payload);
      if (res.success && (res.groupName === groupName || res.isAdmin)) {
        return { success: true, message: '驗證成功', encryptedCode: res.encryptedCode };
      }

      return { success: false, message: res.message || '驗證失敗' };
    },

    // ── 4. 後台管理清單 (getAdminGroupsList) ─────────────────────
    async getAdminGroupsList(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const rawCode = String(payload.authCode || payload.groupCode || '').trim();
      const code = decryptId(rawCode).trim().toUpperCase();
      const isAdmin = (code === ADMIN_CODE);

      const { data, error } = await sb.from('groups').select('*').order('name');
      if (error) throw error;

      let groups = (data || []).map(g => ({
        uuid: g.uuid,
        name: g.name,
        code: g.code,
        status: g.status || '顯示',
        type: g.group_type || '一般小組',
        associatedGroup: g.associated_group || '',
        districtUuid: g.district_uuid || '',
        clusterUuid: g.cluster_uuid || '',
        date: g.date || ''
      }));

      if (!isAdmin) {
        groups = groups.filter(g => (g.code && g.code.toUpperCase() === code));
        if (groups.length === 0) {
          return { success: false, message: '權限不足或輸入代碼錯誤' };
        }
      }

      return {
        success: true,
        groups: groups,
        isAdmin: isAdmin,
        districts: [],
        clusters: []
      };
    },

    // ── 5. 更新小組資訊 (updateGroupInfo) ────────────────────────
    async updateGroupInfo(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();
      const uuid = payload.uuid;
      const status = payload.status;
      const type = payload.type;
      const associatedGroup = payload.associatedGroup;

      let query = sb.from('groups').update({
        status: status,
        group_type: type,
        associated_group: associatedGroup,
        updated_at: new Date().toISOString()
      });

      if (uuid) {
        query = query.eq('uuid', uuid);
      } else {
        query = query.eq('name', groupName);
      }

      await query;

      setTimeout(() => {
        try {
          if (typeof window.churchAPI_original === 'function') {
            window.churchAPI_original('updateGroupInfo', payload).catch(e => console.warn('[Group Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '小組資訊已成功更新' };
    },

    // ── 6. 檢查小組狀態與名單 (checkGroupStatus) ──────────────────
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
        name: m.name,
        role: m.role || '組員',
        sortOrder: m.sort_order || 0
      }));

      return {
        success: true,
        status: group.status || '顯示',
        type: group.group_type || '一般小組',
        associatedGroup: group.associated_group || '',
        members: members
      };
    },

    // ── 7. 取得統計報表 (getStats / getAllGroupsStats / getAllGroupMembers) ─
    async getStats(payload = {}) {
      const sb = getSupabase();
      if (!sb) {
        const gasFn = window.churchAPI_original || window.churchAPI;
        if (typeof gasFn === 'function') return await gasFn('getStats', payload);
        return null;
      }

      const groupName = String(payload.groupName || '').trim();
      const isRawMode = (payload.startDate === 'RAW_MODE');
      const isAll = (!groupName || groupName === 'ALL' || groupName === '小組清單');

      let query = sb.from('group_attendance_records').select('*').order('date', { ascending: false });
      if (!isAll) {
        query = query.eq('group_name', groupName);
      }

      const { data: records } = await query;

      // 若 Supabase 尚無該小組的歷史紀錄，回退至 GAS 讀取歷史試算表
      if (!records || records.length === 0) {
        const gasFn = window.churchAPI_original || window.churchAPI;
        if (typeof gasFn === 'function') {
          try {
            const gasRes = await gasFn('getStats', payload);
            if (gasRes && gasRes.success) return gasRes;
          } catch (e) {
            console.warn('[GroupSupabase] getStats GAS fallback:', e);
          }
        }
      }

      // 建立會友 UID -> 姓名反查表
      const { data: mems } = await sb.from('church_members').select('uid, name, group_name, role');
      const nameDirectory = {};
      (mems || []).forEach(m => {
        if (m.uid) nameDirectory[m.uid.toUpperCase()] = m.name;
      });

      if (isRawMode) {
        return {
          success: true,
          groupName: groupName,
          isSingleDay: false,
          data: (records || []).map(r => [
            r.date ? String(r.date).slice(0, 10).replace(/-/g, '/') : '',
            Array.isArray(r.present_members) ? r.present_members.join(', ') : String(r.present_members || ''),
            r.offering || 0,
            Array.isArray(r.new_friends) ? r.new_friends.join(', ') : String(r.new_friends || '')
          ]),
          nameDirectory: nameDirectory
        };
      }

      // 組員出席率分析模式
      const sDate = payload.startDate ? new Date(payload.startDate) : null;
      const eDate = payload.endDate ? new Date(payload.endDate) : null;
      const isSingleDay = (payload.startDate === payload.endDate && payload.startDate !== '' && payload.startDate !== undefined);

      const groupMembers = (mems || []).filter(m => {
        if (isAll) return Boolean(m.group_name);
        return String(m.group_name || '').includes(groupName);
      });

      const filteredRecords = (records || []).filter(r => {
        if (!r.date) return false;
        const t = new Date(r.date).getTime();
        if (sDate && t < sDate.getTime()) return false;
        if (eDate && t > eDate.getTime()) return false;
        return true;
      });

      const { data: sundayRecs } = await sb.from('attendance_records').select('*');
      const filteredSunday = (sundayRecs || []).filter(sr => {
        if (!sr.date) return false;
        const t = new Date(sr.date).getTime();
        if (sDate && t < sDate.getTime()) return false;
        if (eDate && t > eDate.getTime()) return false;
        return true;
      });

      const totalCellSessions = filteredRecords.length;
      const worshipSessions = filteredSunday.filter(sr => ['台語', '華語', '聯合'].includes(sr.service_type)).length;
      const schoolSessions = filteredSunday.filter(sr => String(sr.service_type || '').includes('主日學')).length;

      const data = groupMembers.map(m => {
        const uid = m.uid;
        const cellCount = filteredRecords.filter(r => (r.present_members || []).includes(uid) || (r.present_members || []).includes(m.name)).length;
        const sundayCount = filteredSunday.filter(sr => ['台語', '華語', '聯合'].includes(sr.service_type) && (sr.present_uids || []).includes(uid)).length;
        const schoolCount = filteredSunday.filter(sr => String(sr.service_type || '').includes('主日學') && (sr.present_uids || []).includes(uid)).length;

        if (isSingleDay) {
          return {
            name: m.name,
            uid: m.uid,
            group: m.group_name || groupName,
            cell: cellCount > 0,
            sunday: sundayCount > 0,
            school: schoolCount > 0
          };
        }

        return {
          name: m.name,
          uid: m.uid,
          group: m.group_name || groupName,
          cellRate: totalCellSessions > 0 ? ((cellCount / totalCellSessions) * 100).toFixed(1) : 0,
          cellStr: `${cellCount}/${totalCellSessions}`,
          sundayRate: worshipSessions > 0 ? ((sundayCount / worshipSessions) * 100).toFixed(1) : 0,
          sundayStr: `${sundayCount}/${worshipSessions}`,
          schoolRate: schoolSessions > 0 ? ((schoolCount / schoolSessions) * 100).toFixed(1) : 0,
          schoolStr: `${schoolCount}/${schoolSessions}`
        };
      });

      if (!isSingleDay) {
        data.sort((a, b) => parseFloat(b.cellRate) - parseFloat(a.cellRate));
      }

      return {
        success: true,
        groupName: groupName,
        isSingleDay: isSingleDay,
        data: data
      };
    },

    async getAllGroupsStats(payload = {}) {
      return this.getStats({ ...payload, groupName: 'ALL' });
    },

    async getAllGroupMembers(payload = {}) {
      const sb = getSupabase();
      if (!sb) {
        const gasFn = window.churchAPI_original || window.churchAPI;
        if (typeof gasFn === 'function') return await gasFn('getAllGroupMembers', payload);
        return null;
      }

      const { data: mems, error } = await sb
        .from('church_members')
        .select('uid, name, gender, group_name, role')
        .order('uid', { ascending: true });

      if (error) throw error;

      const filtered = (mems || []).filter(m => Boolean(m.group_name));

      return {
        success: true,
        data: filtered.map(m => ({
          name: m.name,
          gender: m.gender || '',
          uid: m.uid || '',
          group: m.group_name || '',
          role: m.role || '小羊'
        }))
      };
    },

    // ── 8. 儲存小組點名 (submitAttendance) ───────────────────────
    async submitAttendance(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();
      const dateStr = String(payload.date || '').slice(0, 10);
      const attendees = Array.isArray(payload.attendees) ? payload.attendees : [];
      const newFriends = Array.isArray(payload.newFriends) ? payload.newFriends : [];
      const offering = Number(payload.offering || 0);
      const notes = String(payload.notes || '').trim();

      if (!groupName || !dateStr) {
        return { success: false, message: '小組名稱與日期不可為空' };
      }

      const record = {
        group_name: groupName,
        date: dateStr,
        present_members: attendees,
        new_friends: newFriends,
        offering: offering,
        notes: notes,
        updated_at: new Date().toISOString()
      };

      const { error } = await sb.from('group_attendance_records').upsert(record, {
        onConflict: 'group_name,date'
      });

      if (error) throw error;

      setTimeout(() => {
        try {
          if (typeof window.churchAPI_original === 'function') {
            window.churchAPI_original('submitAttendance', payload).catch(e => console.warn('[Group Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '小組點名已成功送出！' };
    },

    // ── 9. 更新點名紀錄 (updateAttendanceRecord) ─────────────────
    async updateAttendanceRecord(payload) {
      return this.submitAttendance(payload);
    },

    // ── 10. 刪除點名紀錄 (deleteAttendanceRecord) ────────────────
    async deleteAttendanceRecord(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();
      const dateStr = String(payload.date || '').slice(0, 10);

      const { error } = await sb
        .from('group_attendance_records')
        .delete()
        .eq('group_name', groupName)
        .eq('date', dateStr);

      if (error) throw error;

      setTimeout(() => {
        try {
          if (typeof window.churchAPI_original === 'function') {
            window.churchAPI_original('deleteAttendanceRecord', payload).catch(e => console.warn('[Group Backup] GAS sync:', e.message));
          }
        } catch (e) {}
      }, 10);

      return { success: true, message: '點名紀錄已成功刪除' };
    },

    // ── 11. 更新組員名冊 (updateMemberList) ──────────────────────
    async updateMemberList(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();
      const members = Array.isArray(payload.members) ? payload.members : [];

      if (!groupName) return { success: false, message: '小組名稱不可為空' };

      await sb.from('group_members').delete().eq('group_name', groupName);

      if (members.length > 0) {
        const rows = members.map((m, idx) => ({
          group_name: groupName,
          name: typeof m === 'string' ? m : m.name,
          role: (typeof m === 'object' && m.role) ? m.role : '組員',
          sort_order: idx + 1,
          updated_at: new Date().toISOString()
        }));

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

      return { success: true, message: '名冊已成功更新！' };
    },

    // ── 12. 初始化新小組 (initGroup) ──────────────────────────────
    async initGroup(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const groupName = String(payload.groupName || '').trim();
      const type = payload.type || '一般小組';
      const associatedGroup = payload.associatedGroup || '';

      const { data, error } = await sb.from('groups').insert([{
        name: groupName,
        code: ('LK' + Math.floor(10 + Math.random() * 90)),
        group_type: type,
        associated_group: associatedGroup,
        status: '顯示',
        updated_at: new Date().toISOString()
      }]).select().single();

      if (error) throw error;

      return { success: true, group: data };
    },

    // ── 13. 會友名冊建議 (getMemberSuggestions) ──────────────────
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

    // ── 14. 每週統計週報 (getWeeklyReport) ────────────────────────
    async getWeeklyReport(payload) {
      const sb = getSupabase();
      if (!sb) return null;

      const dateStr = String(payload.date || '').slice(0, 10);

      const [groupsRes, recordsRes] = await Promise.all([
        sb.from('groups').select('*').eq('status', '顯示').order('name'),
        sb.from('group_attendance_records').select('*').eq('date', dateStr)
      ]);

      const groups = groupsRes.data || [];
      const records = recordsRes.data || [];
      const recordMap = {};
      records.forEach(r => { recordMap[r.group_name] = r; });

      const report = groups.map(g => {
        const r = recordMap[g.name];
        return {
          groupName: g.name,
          groupType: g.group_type || '一般小組',
          isSubmitted: Boolean(r),
          attendeeCount: r ? (r.present_members || []).length : 0,
          newFriendCount: r ? (r.new_friends || []).length : 0,
          offering: r ? Number(r.offering || 0) : 0,
          notes: r ? r.notes : ''
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
