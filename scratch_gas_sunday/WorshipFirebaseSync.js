// ==========================================
// WorshipFirebaseSync.js — 敬拜團 GAS 端 Realtime Database 快取失效
// ==========================================

/**
 * 📊 onEdit trigger — 偵測手動編輯敬拜團 Sheet，自動清除快取
 * ⚠️ 此函數在 setupAllOnEditTriggers 中需要被綁定到敬拜團試算表上
 */
function onEditWorshipSheet(e) {
  if (!e || !e.range) return;
  const sheetName = e.range.getSheet().getName();
  try {
    if (sheetName === '服事表總表') {
      firebaseInvalidate(['getSchedule', 'getScheduleByDateRange']);
    } else if (sheetName === '位置人員清單') {
      firebaseInvalidate(['getPositions']);
    } else if (sheetName === '敬拜團員名單') {
      firebaseInvalidate(['getTeamMembers', 'getPositions']);
    } else if (sheetName === '敬拜曲目') {
      firebaseInvalidate(['getSongs', 'getSchedule', 'getScheduleByDateRange']);
    } else if (sheetName === '行事曆連結設定') {
      firebaseInvalidate(['getSchedule', 'getScheduleByDateRange']);
    }
  } catch (err) {
    console.log('[onEditWorshipSheet] 失敗: ' + err.message);
  }
}

/**
 * 清空敬拜團相關 Firebase RTDB 快取
 */
function firebaseCacheClearWorship() {
  try {
    firebaseInvalidate([
      'getSchedule',
      'getScheduleByDateRange',
      'getPositions',
      'getTeamMembers',
      'getSongs',
      'worship_getSchedule',
      'worship_getScheduleByDateRange',
      'worship_getPositions',
      'worship_getTeamMembers',
      'worship_getSongs'
    ]);
    Logger.log('✅ 已清空敬拜團相關 Firebase cache');
  } catch (e) {
    Logger.log('❌ 清空失敗：' + e.message);
  }
}

function worship_refreshCaches() {
  try {
    if (typeof clearCalendarLinkCache === 'function') clearCalendarLinkCache();
    firebaseCacheClearWorship();
    return { success: true, message: 'Worship caches refreshed' };
  } catch (e) {
    return { success: false, message: 'Worship cache refresh failed: ' + e.message };
  }
}
