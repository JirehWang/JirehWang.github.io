// 📦 中央安全路由設定 (多專案版)
//
// 使用方式：HTML 在載入此檔之前先宣告自己的 key，例如：
//   <script>window._GAS_KEY = 'LKC_worship';</script>
//   <script src="https://jirehwang.github.io/LKC1958_June_1.github.io/config.js"></script>
//
// 若沒有宣告 _GAS_KEY，會 fallback 到 pathname / hostname 推測。
//
// 提供的共用 API：
//   window.churchAPI(action, data)        — 呼叫 GAS
//   window.ensureAPIReady()                — 等待路由就緒（事件式）
//   window.showLoading(msg) / hideLoading()— 顯示遮罩（自動偵測 DOM）
//   window.userNotification                — toast 通知 + 載入指示
//   window.uiState                         — 防重複提交鎖
//   window.sessionManager                  — sessionStorage 管理
//   window.APIError                        — 自訂錯誤類別
(function() {
  // 📝 子系統 → GAS 部署網址 對應表
  const _URL_ROUTER = {
    "LKC_worship":                      "https://script.google.com/macros/s/AKfycbyk_6tUucVg-U4rRQjYHvk632teZyxufDkNX_X1WRUXPMGgsTaemVXD_mv9kBDjuSwOnA/exec",
    "LKC_MasterSchedule":               "https://script.google.com/macros/s/AKfycbwiYYWgKxmLRAEaE_pbp_kWyAzlRPcwYVQfvmJVamRJvosvt5wTTkvwebbFBkP8rMqX/exec",
    "LKC_MinistrySchedule":             "https://script.google.com/macros/s/AKfycbx4268IkgwQm2Es0gjDHLU_U9nKJrRMR1-xzbbtuaq08lePLgAQ2wnDRrCeHdy9jNhh/exec",
    "LKC_Group":                        "https://script.google.com/macros/s/AKfycbxBOFeLiXu23kBMGU8iSvRyJci6fruTfk7HdahhcQFY777sCPSgasuNM7Z1CeuzuS-r/exec",
    "LKC_WhosCar":                      "https://script.google.com/macros/s/AKfycbxOkoaNquIx_V8n_7eS_5ULmoqxPVly_Bezx9_QsmWSzNOcojrCI9Oa6UNd5hOD2euS/exec",
    "LKC_SundayserviceAttendance":      "https://script.google.com/macros/s/AKfycbxBOFeLiXu23kBMGU8iSvRyJci6fruTfk7HdahhcQFY777sCPSgasuNM7Z1CeuzuS-r/exec",
    "LKC_SundayserviceAttendance_TEST": "https://script.google.com/macros/s/AKfycbxBOFeLiXu23kBMGU8iSvRyJci6fruTfk7HdahhcQFY777sCPSgasuNM7Z1CeuzuS-r/exec",
    // 🔀 方案 B 整合：小組系統共用主日 GAS
    "LKC_Group_TEST":                   "https://script.google.com/macros/s/AKfycbxBOFeLiXu23kBMGU8iSvRyJci6fruTfk7HdahhcQFY777sCPSgasuNM7Z1CeuzuS-r/exec",
    // 🔀 方案 C 整合：事工管理共用主日 GAS（action 自動加 ministry_ 前綴）
    "LKC_MinistrySchedule_TEST":        "https://script.google.com/macros/s/AKfycbxBOFeLiXu23kBMGU8iSvRyJci6fruTfk7HdahhcQFY777sCPSgasuNM7Z1CeuzuS-r/exec",
    "LKC_NewFamily":                    "https://script.google.com/macros/s/AKfycbzU4f0XKtniINXQMbIK5QDPuT3ub2HeyiEYI60oUM3YHipdf-02uvuP3lp963dogxml/exec",
    // 🔀 Phase 5 & 6 整合測試版：行事曆與敬拜團併入主 GAS
    "LKC_MasterSchedule_TEST":          "https://script.google.com/macros/s/AKfycbxBOFeLiXu23kBMGU8iSvRyJci6fruTfk7HdahhcQFY777sCPSgasuNM7Z1CeuzuS-r/exec",
    "LKC_worship_TEST":                 "https://script.google.com/macros/s/AKfycbxBOFeLiXu23kBMGU8iSvRyJci6fruTfk7HdahhcQFY777sCPSgasuNM7Z1CeuzuS-r/exec",
  };

  // 📝 子系統 → 後端 action 自動前綴（避免不同系統 action 名稱衝突）
  const _ACTION_PREFIX = {
    "LKC_MinistrySchedule_TEST": "ministry_",
    "LKC_worship_TEST":          "worship_",
  };

  const _AUTH_TOKEN = "ChurchApp-2026";
  const _SESSION_TTL_MS = 3600000; // 1 小時

  // 🌟 路由判斷：_GAS_KEY 優先，其次 pathname / hostname
  const rawPath = window.location.pathname.split('/')[1] || "";
  const repoName = rawPath.replace(/\.github\.io$/i, '');
  const hostname = window.location.hostname.split('.')[0];

  // 🌟 自動切換測試路由（若在 localhost, 127.0.0.1、本地檔案執行，或帶有 ?test=1/?env=test，自動強制走 _TEST 路由）
  const isTestQuery = /[?&](env=test|test=1)\b/.test(window.location.search);
  const isLocalEnv = window.location.hostname === 'localhost' || 
                     window.location.hostname === '127.0.0.1' || 
                     window.location.protocol === 'file:' ||
                     isTestQuery;

  if (isLocalEnv && window._GAS_KEY && !window._GAS_KEY.endsWith('_TEST')) {
    const testKey = window._GAS_KEY + '_TEST';
    if (_URL_ROUTER[testKey]) {
      console.warn(`⚠️ [開發環境] 偵測到本地/測試執行，已將 ${window._GAS_KEY} 自動切換為測試版 ${testKey}`);
      window._GAS_KEY = testKey;
    }
  }

  let currentKey = null;
  if (window._GAS_KEY && _URL_ROUTER[window._GAS_KEY]) {
    currentKey = window._GAS_KEY;
  } else if (_URL_ROUTER[repoName]) {
    currentKey = repoName;
  } else if (_URL_ROUTER[hostname]) {
    currentKey = hostname;
  }

  if (!currentKey) {
    console.error(`🚨 [路由錯誤] 找不到對應的 GAS：未宣告 window._GAS_KEY，且 pathname='${repoName}' / hostname='${hostname}' 都不在 _URL_ROUTER 中`);
    window.GAS_URL = null;
  } else {
    window.GAS_URL = _URL_ROUTER[currentKey];
    console.log(`✅ [${currentKey}] 中央路由系統已就緒`);
  }

  window.AUTH_TOKEN = _AUTH_TOKEN;

  // ============================================================
  //  🎯 自訂 APIError
  // ============================================================
  class APIError extends Error {
    constructor(message, httpStatus = null, type = 'UNKNOWN_ERROR') {
      super(message);
      this.name = 'APIError';
      this.httpStatus = httpStatus;
      this.type = type;
    }
  }
  window.APIError = APIError;

  // ============================================================
  //  🚀 中央 API 呼叫（含 Firebase RTDB 快取整合）
  //    若當前 _GAS_KEY 在 _ACTION_PREFIX 中，自動加前綴
  //    若 action 在 _CACHEABLE_ACTIONS 中，自動經 Firebase 快取
  //    若 action 在 _INVALIDATE_ON_WRITE 中，呼叫後自動清除相關快取
  // ============================================================
  const _actionPrefix = _ACTION_PREFIX[currentKey] || "";

  // 📋 快取設定（key = 完整 action 名稱含前綴；value = TTL 秒）
  // 統一 21600 = 6 小時。實際更新依靠 invalidation 機制（前端 + GAS + onEdit）
  // TTL 只是「兜底」：若 invalidation 全部漏接，最壞 6 小時後自動刷新
  const _SIX_HOURS = 21600;
  const _CACHEABLE_ACTIONS = {
    // 主日系統 — 全域資料
    'getGroups':                    _SIX_HOURS,
    'getGroupConfig':               _SIX_HOURS,
    'getWeeklyReport':              _SIX_HOURS,
    'getAllMembers':                _SIX_HOURS,
    'getAdminGroupsList':           _SIX_HOURS,
    'getAllGroupMembers':           _SIX_HOURS,
    'getMemberSuggestions':         _SIX_HOURS,

    // 主日系統 — 點名介面（冷啟動加速）
    'getSmartAttendanceList':       _SIX_HOURS,  // 點名介面：會友 + 出席計數 + 同步狀態（首次載入）
    'checkGroupStatus':             _SIX_HOURS,  // 小組點名首頁：組員 + 初始化狀態

    // 主日系統 — 統計 / 圖表
    'getStats':                     _SIX_HOURS,
    'getAllGroupsStats':            _SIX_HOURS,
    'getAttendanceStats':           _SIX_HOURS,
    'getAttendanceTrend':           _SIX_HOURS,
    'getCategoryChartData':         _SIX_HOURS,  // 趨勢圖表（依分類）

    // 事工管理
    'ministry_getGroups':           _SIX_HOURS,
    'ministry_getTemplates':        _SIX_HOURS,
    'ministry_getAggregatedReport': _SIX_HOURS,
    'ministry_getPageConfig':       _SIX_HOURS,
    'ministry_getGroupMembers':     _SIX_HOURS,

    // 敬拜團 (無前綴版本，供舊獨立專案相容)
    'getSchedule':                  _SIX_HOURS,  // 公佈欄總表 + 服事表安排（季度）
    'getScheduleByDateRange':       _SIX_HOURS,  // 服事表安排（區間）
    'getPositions':                 _SIX_HOURS,  // 位置與同工
    'getTeamMembers':               _SIX_HOURS,  // 敬拜團員名單
    'getSongs':                     _SIX_HOURS,  // 敬拜曲目（順帶加上）

    // 敬拜團 (worship_ 前綴版本，供合併後的主 GAS 專案使用)
    'worship_getSchedule':              _SIX_HOURS,
    'worship_getScheduleByDateRange':   _SIX_HOURS,
    'worship_getPositions':             _SIX_HOURS,
    'worship_getTeamMembers':           _SIX_HOURS,
    'worship_getSongs':                 _SIX_HOURS,
    'worship_getMemberSuggestions':     _SIX_HOURS,

    // 教會行事曆
    'cal_getTypes':                 _SIX_HOURS,  // 事項類型（含 children 樹）
    'cal_getFields':                _SIX_HOURS,  // 欄位定義（含繼承解析）
    'cal_getEvents':                _SIX_HOURS,  // 事項清單（依日期/類型篩選）
    'cal_getEvent':                 _SIX_HOURS   // 單一事項詳情
  };

  // 寫入時要連帶清除的 read-cache（key = 寫入 action，value = 要清掉的 read action 陣列）
  // 使用 cacheDeleteAll 清整個 topic（包含所有 subkey），所以不分 data 變體
  const _INVALIDATE_ON_WRITE = {
    // ── 會友名單異動 ──
    'addMember':              ['getAllMembers', 'getAllGroupMembers', 'getMemberSuggestions', 'getStats', 'getAllGroupsStats', 'getAdminGroupsList', 'ministry_getPageConfig', 'getSmartAttendanceList'],
    'updateMember':           ['getAllMembers', 'getAllGroupMembers', 'getMemberSuggestions', 'getStats', 'getAllGroupsStats', 'getAdminGroupsList', 'ministry_getPageConfig', 'getSmartAttendanceList'],
    'deleteMember':           ['getAllMembers', 'getAllGroupMembers', 'getMemberSuggestions', 'getStats', 'getAllGroupsStats', 'getAdminGroupsList', 'ministry_getPageConfig', 'getSmartAttendanceList'],

    // ── 小組異動 ──
    'createGroup':            ['getGroups', 'getAdminGroupsList', 'findGroupByCode', 'verifyGroup'],
    'updateGroupInfo':        ['getGroups', 'getAdminGroupsList', 'findGroupByCode', 'verifyGroup'],
    'updateMemberList':       ['getAllMembers', 'getAllGroupMembers', 'getMemberSuggestions', 'getStats', 'getAllGroupsStats', 'ministry_getPageConfig', 'ministry_getGroupMembers', 'checkGroupStatus'],
    'ministry_updateGroupMemberRoles': ['getAllMembers', 'getAllGroupMembers', 'getMemberSuggestions', 'getStats', 'getAllGroupsStats', 'ministry_getPageConfig', 'ministry_getGroupMembers', 'checkGroupStatus'],
    'initGroup':              ['checkGroupStatus', 'getGroups', 'getAllGroupMembers'],

    // ── 主日點名異動 ──
    'createAttendanceGroup':  ['getGroupConfig'],
    'saveAttendance':         ['getWeeklyReport', 'getAttendanceStats', 'getAttendanceTrend', 'getStats', 'getAllGroupsStats', 'getSmartAttendanceList', 'getCategoryChartData'],
    'revokeAttendance':       ['getWeeklyReport', 'getAttendanceStats', 'getAttendanceTrend', 'getStats', 'getAllGroupsStats', 'getSmartAttendanceList', 'getCategoryChartData'],

    // ── 小組點名異動 ──
    'submitAttendance':       ['getWeeklyReport', 'getStats', 'getAllGroupsStats', 'checkGroupStatus', 'getCategoryChartData'],
    'updateAttendanceRecord': ['getWeeklyReport', 'getStats', 'getAllGroupsStats', 'checkGroupStatus', 'getCategoryChartData'],
    'deleteAttendanceRecord': ['getWeeklyReport', 'getStats', 'getAllGroupsStats', 'checkGroupStatus', 'getCategoryChartData'],

    // ── 事工異動 ──
    'ministry_createGroup':       ['ministry_getGroups', 'ministry_getAggregatedReport'],
    'ministry_toggleGroupStatus': ['ministry_getGroups'],
    'ministry_saveSheetData':     ['ministry_getAggregatedReport', 'ministry_getPageConfig'],
    'ministry_saveGroupPrompt':   ['ministry_getPageConfig'],
    'ministry_saveGroupMembers':  ['ministry_getPageConfig'],
    'ministry_saveSermonSettings': ['ministry_getPageConfig'],
    'ministry_forceRefreshEvents': ['ministry_getPageConfig'],

    // ── 敬拜團異動 ──
    'saveSchedule':           ['getSchedule', 'getScheduleByDateRange'],
    'savePositions':          ['getPositions'],
    'saveTeamMembers':        ['getTeamMembers'],
    'saveSongs':              ['getSongs', 'getSchedule', 'getScheduleByDateRange'],
    // 行事曆連結設定改變 → 公佈欄 / 服事表的「聚會名稱、聚會類別、講道資訊」會跟著變
    'setDefaultSermonSubType': ['getSchedule', 'getScheduleByDateRange'],
    'setDateOverride':         ['getSchedule', 'getScheduleByDateRange'],
    'clearCalendarLinkCache':  ['getSchedule', 'getScheduleByDateRange'],

    // ── 敬拜團異動 (worship_ 前綴版本) ──
    'worship_saveSchedule':           ['worship_getSchedule', 'worship_getScheduleByDateRange'],
    'worship_savePositions':          ['worship_getPositions'],
    'worship_saveTeamMembers':        ['worship_getTeamMembers'],
    'worship_saveSongs':              ['worship_getSongs', 'worship_getSchedule', 'worship_getScheduleByDateRange'],
    'worship_setDefaultSermonSubType': ['worship_getSchedule', 'worship_getScheduleByDateRange'],
    'worship_setDateOverride':         ['worship_getSchedule', 'worship_getScheduleByDateRange'],
    'worship_clearCalendarLinkCache':  ['worship_getSchedule', 'worship_getScheduleByDateRange'],

    // ── 教會行事曆異動 ──
    // 類型 / 排除欄位改變 → 影響事項本身的子類型/欄位顯示 + 敬拜團公佈欄
    'cal_addType':             ['cal_getTypes', 'cal_getFields', 'cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange'],
    'cal_updateType':          ['cal_getTypes', 'cal_getFields', 'cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange'],
    'cal_deleteType':          ['cal_getTypes', 'cal_getFields', 'cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange'],
    // 欄位改變 → 事項顯示 + 敬拜團公佈欄
    'cal_addField':            ['cal_getFields', 'cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange'],
    'cal_updateField':         ['cal_getFields', 'cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange'],
    'cal_deleteField':         ['cal_getFields', 'cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange'],
    'cal_reorderFields':       ['cal_getFields', 'cal_getEvents', 'cal_getEvent'],
    // 事項 CRUD → 公佈欄要重抓
    'cal_addEvent':            ['cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange'],
    'cal_updateEvent':         ['cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange'],
    'cal_deleteEvent':         ['cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange'],
    'cal_addEventsBatch':      ['cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange'],
    // 整批操作
    'cal_setupSchema':         ['cal_getTypes', 'cal_getFields', 'cal_getEvents', 'cal_getEvent'],
    'cal_migrateOldData':      ['cal_getTypes', 'cal_getFields', 'cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange'],
    'cal_clearNewData':        ['cal_getEvents', 'cal_getEvent', 'getSchedule', 'getScheduleByDateRange']
  };

  // Lazy-load Firebase cache module（僅在需要時 import；失敗則自動 fallback 至直接 GAS 呼叫）
  let _firebaseCachePromise = null;
  function _getFirebaseCache() {
    if (_firebaseCachePromise) return _firebaseCachePromise;
    _firebaseCachePromise = import('https://jirehwang.github.io/LKC1958_June_1.github.io/firebase/firebase-cache.js')
      .catch(err => { console.warn('[firebase-cache] 載入失敗，將以直接呼叫 GAS 模式運作:', err); return null; });
    return _firebaseCachePromise;
  }

  let _firebaseLoggerPromise = null;
  function _getFirebaseLogger() {
    if (_firebaseLoggerPromise) return _firebaseLoggerPromise;
    _firebaseLoggerPromise = import('https://jirehwang.github.io/LKC1958_June_1.github.io/firebase/firebase-logger.js')
      .catch(err => { console.warn('[firebase-logger] 載入失敗，略過遠端紀錄:', err); return null; });
    return _firebaseLoggerPromise;
  }

  function _logEvent(level, action, message, meta = {}) {
    if (!currentKey) return;
    _getFirebaseLogger().then(logger => {
      if (!logger || !logger.writeLog) return;
      return logger.writeLog({
        system: currentKey,
        level,
        action,
        message,
        durationMs: meta.durationMs,
        source: 'config.js',
        meta
      });
    }).catch(err => {
      console.warn('[church-log] 寫入失敗:', err);
    });
  }

  window.churchLog = function(entry = {}) {
    _logEvent(entry.level || 'info', entry.action || '', entry.message || '', entry.meta || {});
  };

  // 為帶 data 的 cache 計算 subkey（無 data → null 走 _default）
  // ⚠️ 不可截斷！之前用 substring(0, 24) 會讓 {year:'2026',quarter:'Q1'}
  //    與 {year:'2026',quarter:'Q2'} 產生同樣 subkey（差異在末尾），導致 cache 撞鍵
  function _makeSubkey(data) {
    if (!data || Object.keys(data).length === 0) return null;
    const json = JSON.stringify(data);
    return btoa(unescape(encodeURIComponent(json))).replace(/[^a-zA-Z0-9]/g, '');
  }

  function _isInvalidApiResponse(value) {
    return value &&
      typeof value === 'object' &&
      ((Object.prototype.hasOwnProperty.call(value, 'status') && value.status !== 'success') ||
       (Object.prototype.hasOwnProperty.call(value, 'success') && value.success === false));
  }

  // 直接打 GAS（不走 cache）
  async function _doDirectCall(realAction, data) {
    const resp = await fetch(window.GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: realAction, token: window.AUTH_TOKEN, data: data })
    });
    return await resp.json();
  }

  window.churchAPI = async function(action, data = {}) {
    if (!window.GAS_URL) {
      throw new APIError("系統尚未就緒：GAS_URL 為空", null, 'CONFIG_ERROR');
    }
    const startedAt = Date.now();
    // 已有前綴的 action 不重複加
    const realAction = (_actionPrefix && action.indexOf(_actionPrefix) !== 0)
      ? _actionPrefix + action
      : action;
    const ttl = _CACHEABLE_ACTIONS[realAction];
    const isWriteAction = Object.prototype.hasOwnProperty.call(_INVALIDATE_ON_WRITE, realAction);

    function logApi(level, message, extra = {}) {
      _logEvent(level, realAction, message, Object.assign({
        requestedAction: action,
        durationMs: Date.now() - startedAt,
        cacheable: Boolean(ttl),
        writeAction: isWriteAction
      }, extra));
    }

    try {
      // 🔥 可快取 → 走 Firebase（cache 路徑：cache/{action}/{dataHash}）
      if (ttl) {
        const fb = await _getFirebaseCache();
        if (fb && fb.cacheGetOrFetch) {
          const subkey = _makeSubkey(data);
          try {
            const cacheLoader = () => _doDirectCall(realAction, data);
            const cacheResult = fb.cacheGetOrFetchWithMeta
              ? await fb.cacheGetOrFetchWithMeta(realAction, subkey, cacheLoader, ttl)
              : { value: await fb.cacheGetOrFetch(realAction, subkey, cacheLoader, ttl), source: 'unknown' };
            const result = cacheResult.value;
            if (_isInvalidApiResponse(result)) {
              if (fb.cacheDelete) {
                fb.cacheDelete(realAction, subkey).catch(err => {
                  console.warn('[firebase-cache] delete invalid cache failed:', realAction, err);
                });
              } else if (fb.cacheDeleteAll) {
                fb.cacheDeleteAll(realAction).catch(err => {
                  console.warn('[firebase-cache] delete invalid cache topic failed:', realAction, err);
                });
              }
              logApi('warn', 'invalid cache response, fallback to GAS', {
                cacheSource: cacheResult.source,
                subkey,
                responseStatus: result.status
              });
              const fallbackResult = await _doDirectCall(realAction, data);
              logApi('warn', 'fallback GAS completed after invalid cache', {
                subkey,
                responseStatus: fallbackResult && fallbackResult.status
              });
              return fallbackResult;
            }
            const durationMs = Date.now() - startedAt;
            if (cacheResult.source === 'fresh' || durationMs > 3000) {
              logApi(durationMs > 5000 ? 'warn' : 'info', 'cacheable API completed', {
                cacheSource: cacheResult.source,
                subkey,
                responseStatus: result && result.status
              });
            }
            return result;
          } catch (e) {
            console.warn('[firebase-cache]', realAction, '失敗，回退直接呼叫:', e);
            logApi('warn', 'firebase cache failed, fallback to GAS', {
              error: e && e.message ? e.message : String(e),
              subkey
            });
          }
        }
      }

      // 直接呼叫 GAS
      const result = await _doDirectCall(realAction, data);
      if (_isInvalidApiResponse(result)) {
        logApi('error', 'GAS returned non-success response', {
          responseStatus: result.status,
          message: result.message || ''
        });
      } else if (isWriteAction || Date.now() - startedAt > 3000) {
        logApi(Date.now() - startedAt > 5000 ? 'warn' : 'info', 'direct GAS API completed', {
          responseStatus: result && result.status
        });
      }

      // 🗑️ 寫入 action 觸發整個 topic 的 cache 失效（不分 data 變體）
      // ⚠️ 改為 await：避免使用者寫入後立刻去讀（例：改完覆寫立刻看公佈欄），
      //    invalidation 還沒完成 → 讀到舊 cache → 看不到變動
      const toInvalidate = _INVALIDATE_ON_WRITE[realAction];
      if (toInvalidate && toInvalidate.length > 0) {
        const fb = await _getFirebaseCache();
        if (fb && fb.cacheDeleteAll) {
          try {
            await Promise.all(toInvalidate.map(topic =>
              fb.cacheDeleteAll(topic).catch(err => {
                console.warn('[invalidate]', topic, err);
              })
            ));
            console.log('[invalidate] cleared:', toInvalidate.join(', '));
            logApi('info', 'cache invalidated after write', {
              topics: toInvalidate
            });
          } catch (e) { /* 已個別 catch */ }
        }
      }

      return result;
    } catch (err) {
      console.error("📡 API 通訊失敗:", err);
      logApi('error', 'API request failed', {
        error: err && err.message ? err.message : String(err),
        errorName: err && err.name,
        httpStatus: err && err.httpStatus,
        type: err && err.type
      });
      throw err;
    }
  };

  // 對外暴露：手動清除整個 topic 的 cache
  window.churchAPIInvalidate = async function(topic) {
    const fb = await _getFirebaseCache();
    if (fb && fb.cacheDeleteAll) await fb.cacheDeleteAll(topic);
  };

  // ============================================================
  //  🐛 debug — 受 ?debug=1 或 window.DEBUG 控制的條件式 log
  //  使用方式：debug('[Module] 訊息', data) 取代瑣碎的 console.log
  // ============================================================
  const _debugEnabled = /[?&]debug=1\b/.test(window.location.search) || window.DEBUG === true;
  window.debug = _debugEnabled
    ? console.log.bind(console, '%c[debug]', 'color:#888')
    : function() {};

  // ============================================================
  //  🗓️ 共用日期工具
  // ============================================================
  window.WEEKDAY_NAMES = ['日', '一', '二', '三', '四', '五', '六'];

  // 格式化為 YYYY-MM-DD（使用本地時區，避免 toISOString 的 UTC 偏移）
  window.formatYMD = function(date) {
    const d = (date instanceof Date) ? date : new Date(date);
    if (isNaN(d.getTime())) return '';
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  };

  // ============================================================
  //  🛡️ ensureAPIReady — 由於 <script> 標籤同步載入順序保證 config.js
  //  在 app script 之前完成，這裡幾乎都是立即 resolve；保留 API 以兼容
  //  舊呼叫方式，並在極端情況提供 microtask 等待。
  // ============================================================
  window.ensureAPIReady = function() {
    if (typeof window.churchAPI === 'function' && window.GAS_URL) {
      return Promise.resolve();
    }
    return new Promise((resolve, reject) => {
      if (typeof window.churchAPI === 'function' && window.GAS_URL) {
        return resolve();
      }
      const timer = setTimeout(() => {
        document.removeEventListener('churchAPIReady', onReady);
        reject(new APIError("安全路由載入逾時，請確認網路連線或檔案路徑。", null, 'CONFIG_ERROR'));
      }, 5000);
      const onReady = () => {
        clearTimeout(timer);
        resolve();
      };
      document.addEventListener('churchAPIReady', onReady, { once: true });
    });
  };

  // ============================================================
  //  💡 載入指示器：自動偵測 DOM 元素
  //    - LKC_Group / LKC_MasterSchedule 用 #loading-overlay + #overlay-text
  //    - LKC_MinistrySchedule 用 #globalLoading
  // ============================================================
  function showLoading(msg = "處理中...") {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
      const textEl = document.getElementById('overlay-text');
      if (textEl) textEl.innerText = msg;
      overlay.style.display = 'flex';
      return;
    }
    const global = document.getElementById('globalLoading');
    if (global) {
      global.innerText = msg;
      global.classList.remove('hidden');
    }
  }

  function hideLoading() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) overlay.style.display = 'none';
    const global = document.getElementById('globalLoading');
    if (global) global.classList.add('hidden');
  }

  window.showLoading = showLoading;
  window.hideLoading = hideLoading;

  // ============================================================
  //  🔔 userNotification — toast 通知（Bootstrap-free，使用 inline style）
  // ============================================================
  const _toastStyle = {
    success: { bg: '#28a745', color: '#fff' },
    warning: { bg: '#ffc107', color: '#212529' },
    danger:  { bg: '#dc3545', color: '#fff' },
    info:    { bg: '#17a2b8', color: '#fff' }
  };

  let _toastStack = 0;
  function showToast(message, type = 'info', duration = 3000) {
    const style = _toastStyle[type] || _toastStyle.info;
    const offset = 16 + (_toastStack * 64);
    _toastStack++;

    const toast = document.createElement('div');
    toast.style.cssText = [
      'position:fixed', `bottom:${offset}px`, 'right:16px', 'z-index:99999',
      `background:${style.bg}`, `color:${style.color}`,
      'padding:12px 16px 12px 18px', 'border-radius:8px',
      'box-shadow:0 4px 14px rgba(0,0,0,0.18)',
      'font-family:"Microsoft JhengHei","Noto Sans TC",sans-serif',
      'font-size:14px', 'max-width:90vw', 'min-width:220px',
      'display:flex', 'align-items:center', 'gap:12px',
      'animation:lkc-toast-in 0.18s ease-out'
    ].join(';');
    toast.textContent = String(message);

    const closeBtn = document.createElement('span');
    closeBtn.textContent = '×';
    closeBtn.style.cssText = `cursor:pointer;font-size:20px;line-height:1;opacity:0.85;color:${style.color};margin-left:auto`;
    closeBtn.addEventListener('click', () => removeToast());
    toast.appendChild(closeBtn);

    document.body.appendChild(toast);

    function removeToast() {
      if (!toast.parentNode) return;
      toast.remove();
      _toastStack = Math.max(0, _toastStack - 1);
    }
    if (duration > 0) setTimeout(removeToast, duration);
  }

  // toast 進場動畫（一次性注入 keyframes）
  if (typeof document !== 'undefined' && !document.getElementById('lkc-toast-style')) {
    const style = document.createElement('style');
    style.id = 'lkc-toast-style';
    style.textContent = '@keyframes lkc-toast-in{from{transform:translateY(20px);opacity:0}to{transform:translateY(0);opacity:1}}';
    document.head && document.head.appendChild(style);
  }

  window.userNotification = {
    success: (msg, d = 3000) => showToast(msg, 'success', d),
    warning: (msg, d = 5000) => showToast(msg, 'warning', d),
    error:   (msg, d = 5000) => showToast(msg, 'danger',  d),
    info:    (msg, d = 3000) => showToast(msg, 'info',    d),
    showLoading: showLoading,
    hideLoading: hideLoading
  };

  // ============================================================
  //  🔐 uiState — 防重複提交鎖
  // ============================================================
  window.uiState = (function() {
    const _locks = {};
    return {
      lock:     (k) => { _locks[k] = true; },
      unlock:   (k) => { _locks[k] = false; },
      isLocked: (k) => _locks[k] === true
    };
  })();

  // ============================================================
  //  💾 sessionManager — sessionStorage 帶過期檢查
  // ============================================================
  window.sessionManager = {
    setUnlocked(id) {
      sessionStorage.setItem(`session_${id}`, JSON.stringify({ unlocked: true, timestamp: Date.now() }));
    },
    isUnlocked(id) {
      const raw = sessionStorage.getItem(`session_${id}`);
      if (!raw) return false;
      try {
        const data = JSON.parse(raw);
        if (Date.now() - data.timestamp > _SESSION_TTL_MS) {
          sessionStorage.removeItem(`session_${id}`);
          return false;
        }
        return data.unlocked === true;
      } catch (e) {
        return false;
      }
    },
    clear(id) {
      sessionStorage.removeItem(`session_${id}`);
    }
  };

  // ─────────────────────────────────────────────────────────────
  //  PWA Service Worker 自動註冊（Stale-While-Revalidate 靜默更新版）
  // ─────────────────────────────────────────────────────────────
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function() {
      // 依據 hostname 動態判定 scope 路徑
      // 本地環境可能是 "/"，而 GitHub Pages 是 "/LKC1958_June_1.github.io/"
      let scopePath = '/';
      let swPath = '/service-worker.js';

      if (window.location.hostname.indexOf('github.io') !== -1) {
        scopePath = '/LKC1958_June_1.github.io/';
        swPath = '/LKC1958_June_1.github.io/service-worker.js';
      }

      navigator.serviceWorker.register(swPath, { scope: scopePath })
        .then(function(reg) {
          console.log('✅ [PWA] ServiceWorker 註冊成功，Scope: ', reg.scope);
        }).catch(function(err) {
          console.warn('❌ [PWA] ServiceWorker 註冊失敗: ', err);
        });
    });
  }

  // 通知就緒（給可能等待的 ensureAPIReady listener）
  if (typeof document !== 'undefined') {
    document.dispatchEvent(new Event('churchAPIReady'));
  }
})();
