/**
 * CacheManager.js — JSON 快取管理（會友名單 + 點名系統清單）
 *
 * ── 會友名單 ──────────────────────────────────────────────────────────────
 *  讀取：getCachedMembers()  → CacheService → cache miss 才讀 Sheet
 *  寫入：addMember / updateMember / deleteMember → Sheet → 立即重建快取
 *  保鮮：keepWarm（每 4 小時）無條件重建，確保 5.5h TTL 到期前已刷新
 *  欄位順序（index 0–8）：
 *    0:姓名  1:性別  2:建立日期  3:備註  4:不列入統計
 *    5:異動日期  6:異動紀錄  7:系統編號  8:所屬小組
 *
 * ── 點名系統清單 ──────────────────────────────────────────────────────────
 *  讀取：getCachedGroupConfig() → CacheService → cache miss 才讀 Sheet
 *  寫入：createAttendanceGroup → Sheet → 立即重建快取
 *  保鮮：同上，由 keepWarm 兜底
 *  格式：{ "類別": ["群組A", "群組B", ...], ... }
 */

const MEMBER_CACHE_KEY      = 'MEMBER_LIST_V2';
const MEMBER_CACHE_TTL      = 19800;     // 5.5 小時（CacheService 單鍵 TTL 上限 6 小時）
const MEMBER_CACHE_CHUNK    = 'MEMBER_LIST_CHK_';
const CHUNK_SIZE            = 90000;     // 每片上限 90 KB（CacheService 單鍵限 100 KB）

const GROUP_CONFIG_CACHE_KEY = 'GROUP_CONFIG_V1';
const GROUP_CONFIG_CACHE_TTL = 19800;    // 5.5 小時；寫入時即時重建，keepWarm 兜底

// ─────────────────────────────────────────────
//  公開 API
// ─────────────────────────────────────────────

/**
 * 取得會友名單陣列（快取優先）
 * cache miss 時自動從 Sheet 重建並寫回快取
 *
 * @returns {Array[]} 每列為 [姓名,性別,建立日期,備註,不列入統計,異動日期,異動紀錄,系統編號,所屬小組]
 */
function getCachedMembers() {
  const raw = _cacheGet(MEMBER_CACHE_KEY);
  if (raw !== null) {
    try { return JSON.parse(raw); } catch (e) { /* 損毀，繼續往下重建 */ }
  }
  return _rebuildMemberCache();
}

/**
 * 資料異動後（addMember / updateMember / deleteMember）呼叫
 * 立刻從 Sheet 重建快取，確保快取不帶舊值
 */
function invalidateAndRebuildMemberCache() {
  return _rebuildMemberCache();
}

// ─────────────────────────────────────────────
//  點名系統清單快取（寫入即同步，無 trigger）
// ─────────────────────────────────────────────

/**
 * 取得點名系統清單（快取優先）
 * cache miss 時自動從 Sheet 重建
 *
 * @returns {{ [category: string]: string[] }}
 */
function getCachedGroupConfig() {
  const raw = _cacheGet(GROUP_CONFIG_CACHE_KEY);
  if (raw !== null) {
    try { return JSON.parse(raw); } catch (e) { /* 損毀，往下重建 */ }
  }
  return _rebuildGroupConfigCache();
}

/**
 * createAttendanceGroup 寫入 Sheet 後呼叫，立即同步快取
 * @returns {{ [category: string]: string[] }} 最新的群組設定（與 getGroupConfig 相同格式）
 */
function invalidateAndRebuildGroupConfigCache() {
  return _rebuildGroupConfigCache();
}

function refreshAttendanceCaches() {
  try {
    const members = invalidateAndRebuildMemberCache();
    const groupConfig = invalidateAndRebuildGroupConfigCache();
    firebaseInvalidate([
      'getAllMembers',
      'getSmartAttendanceList',
      'getGroupConfig',
      'getWeeklyReport',
      'getAttendanceStats',
      'getAttendanceTrend',
      'getCategoryChartData',
      'getStats',
      'getAllGroupsStats'
    ]);
    return {
      success: true,
      message: 'Attendance caches refreshed',
      memberCount: Array.isArray(members) ? members.length : 0,
      groupConfigCategories: groupConfig ? Object.keys(groupConfig).length : 0
    };
  } catch (e) {
    return { success: false, message: 'Attendance cache refresh failed: ' + e.message };
  }
}

// ─────────────────────────────────────────────
//  內部實作
// ─────────────────────────────────────────────

function _rebuildGroupConfigCache() {
  const config = _readGroupConfigFromSheet();
  _cacheSet(GROUP_CONFIG_CACHE_KEY, JSON.stringify(config), GROUP_CONFIG_CACHE_TTL);
  return config;
}

/**
 * 從 Sheet 讀取點名系統清單，回傳與 getGroupConfig() 相同結構
 */
function _readGroupConfigFromSheet() {
  const ss = getSS();
  const listSheet = ss.getSheetByName('點名系統清單');
  if (!listSheet) return {};
  const data = listSheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const category  = data[i][0] ? data[i][0].toString().trim() : '';
    const groupName = data[i][1] ? data[i][1].toString().trim() : '';
    if (category && groupName) {
      if (!config[category]) config[category] = [];
      config[category].push(groupName);
    }
  }
  return config;
}

function _rebuildMemberCache() {
  const members = _readMembersFromSheet();
  _cacheSet(MEMBER_CACHE_KEY, JSON.stringify(members), MEMBER_CACHE_TTL);
  return members;
}

/**
 * 從 Sheet 讀取會友名單（完整原始讀取，不走快取）
 * 回傳固定欄位順序陣列，日期統一轉為 "yyyy/MM/dd" 字串
 *
 * 快取陣列固定 index：
 *   0:姓名 1:性別 2:建立日期 3:備註 4:不列入統計
 *   5:異動日期 6:異動紀錄 7:系統編號 8:所屬小組 9:身分
 */
function _readMembersFromSheet() {
  const sheet = getMemberSheet();
  const fullData = sheet.getDataRange().getValues();
  if (fullData.length <= 1) return [];

  const headers = fullData[0];
  const rows = fullData.slice(1);
  const targetHeaders = [
    '姓名', '性別', '建立日期', '備註', '不列入統計',
    '異動日期', '異動紀錄', '系統編號', '所屬小組', '身分'
  ];

  const colIndex = {};
  targetHeaders.forEach(h => { colIndex[h] = headers.indexOf(h); });

  return rows.map(row =>
    targetHeaders.map(h => {
      const idx = colIndex[h];
      if (idx === -1 || idx >= row.length) return '';
      const cell = row[idx];
      if (cell instanceof Date) return Utilities.formatDate(cell, 'GMT+8', 'yyyy/MM/dd');
      return cell;
    })
  );
}

// ─────────────────────────────────────────────
//  分片存取（支援超過 100 KB 的大型名單）
// ─────────────────────────────────────────────

/**
 * 寫入快取，超過 CHUNK_SIZE 自動分片
 */
function _cacheSet(key, value, ttl) {
  const cache = CacheService.getScriptCache();

  if (value.length <= CHUNK_SIZE) {
    // 單片直接寫入，清除舊分片計數
    cache.put(key, value, ttl);
    cache.remove(key + '_CNT');
  } else {
    // 分片：先清除舊的非分片版本，再批次寫入各片 + 片數
    const chunks = [];
    for (let i = 0; i < value.length; i += CHUNK_SIZE) {
      chunks.push(value.slice(i, i + CHUNK_SIZE));
    }
    const entries = {};
    entries[key + '_CNT'] = String(chunks.length);
    chunks.forEach((chunk, i) => { entries[MEMBER_CACHE_CHUNK + i] = chunk; });
    cache.putAll(entries, ttl);
    cache.remove(key);
  }
}

/**
 * 讀取快取，自動組裝分片
 * @returns {string|null} JSON 字串，或 null（cache miss）
 */
function _cacheGet(key) {
  const cache = CacheService.getScriptCache();

  // 先嘗試非分片版本（最常見路徑）
  const single = cache.get(key);
  if (single !== null) return single;

  // 再嘗試分片版本
  const cntStr = cache.get(key + '_CNT');
  if (cntStr === null) return null;

  const count = parseInt(cntStr, 10);
  const keys = Array.from({ length: count }, (_, i) => MEMBER_CACHE_CHUNK + i);
  const parts = cache.getAll(keys);
  const assembled = keys.map(k => parts[k] || '').join('');
  return assembled || null;
}
