/**
 * FirebaseSync.js — GAS 端寫入 Firebase Realtime Database
 *
 * 認證方式：Service Account（OAuth2 with JWT）— Google 推薦
 *
 * ─── 一次性設定（必做）─────────────────────────────────
 *   1. GAS 編輯器 → ⚙ 專案設定 → 「指令碼屬性」 → 新增屬性：
 *      Key:   FIREBASE_SERVICE_ACCOUNT
 *      Value: <貼上整份 service account JSON 內容>
 *
 *   2. 函式選 setupAllOnEditTriggers → ▶ 執行
 */

const FIREBASE_DB_URL = 'https://lkc1958june1-default-rtdb.asia-southeast1.firebasedatabase.app';

// ═══════════════════════════════════════════════════════════
//  OAuth2 token 取得（含 cache，1 小時內共用同一個 token）
// ═══════════════════════════════════════════════════════════
function _getFirebaseAccessToken() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('FB_ACCESS_TOKEN');
  if (cached) return cached;

  const props = PropertiesService.getScriptProperties();
  const saJson = props.getProperty('FIREBASE_SERVICE_ACCOUNT');
  if (!saJson) {
    throw new Error('FIREBASE_SERVICE_ACCOUNT 未設定（指令碼屬性）');
  }

  let sa;
  try {
    sa = JSON.parse(saJson);
  } catch (e) {
    throw new Error('FIREBASE_SERVICE_ACCOUNT JSON 格式錯誤：' + e.message);
  }

  const now = Math.floor(Date.now() / 1000);
  const header = { alg: 'RS256', typ: 'JWT' };
  const claim = {
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/firebase.database https://www.googleapis.com/auth/userinfo.email',
    aud: sa.token_uri || 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  };

  // Base64URL encode (Web-safe，去掉尾端 =)
  const b64url = (s) => Utilities.base64EncodeWebSafe(s).replace(/=+$/, '');
  const encHeader = b64url(JSON.stringify(header));
  const encClaim = b64url(JSON.stringify(claim));
  const toSign = encHeader + '.' + encClaim;

  // 用 service account 私鑰簽 RS256
  const signature = Utilities.computeRsaSha256Signature(toSign, sa.private_key);
  const encSig = Utilities.base64EncodeWebSafe(signature).replace(/=+$/, '');
  const jwt = toSign + '.' + encSig;

  // 換 access token
  const resp = UrlFetchApp.fetch(claim.aud, {
    method: 'post',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  const result = JSON.parse(resp.getContentText());
  if (result.error) {
    throw new Error('Firebase OAuth 失敗：' + JSON.stringify(result));
  }

  const token = result.access_token;
  // Token 有效 1 小時，cache 存 58 分鐘留 buffer
  cache.put('FB_ACCESS_TOKEN', token, 3500);
  return token;
}

// ═══════════════════════════════════════════════════════════
//  Firebase RTDB 寫/刪 API
// ═══════════════════════════════════════════════════════════

/**
 * 寫入 Firebase cache 路徑：cache/{topic}/{subkey}
 */
function firebaseCacheSet(topic, subkey, value, ttlSeconds) {
  try {
    const token = _getFirebaseAccessToken();
    const sk = subkey || '_default';
    const url = FIREBASE_DB_URL + '/cache/' + encodeURIComponent(topic) + '/' + encodeURIComponent(sk) + '.json';
    const expiresAt = ttlSeconds ? Date.now() + ttlSeconds * 1000 : null;
    UrlFetchApp.fetch(url, {
      method: 'put',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ value: value, expiresAt: expiresAt, updatedAt: Date.now() }),
      muteHttpExceptions: true
    });
  } catch (e) { console.log('[firebase] set(' + topic + ') 失敗: ' + e.message); }
}

/**
 * 清除整個 topic 下所有 subkey
 */
function firebaseCacheDeleteAll(topic) {
  try {
    const token = _getFirebaseAccessToken();
    const url = FIREBASE_DB_URL + '/cache/' + encodeURIComponent(topic) + '.json';
    UrlFetchApp.fetch(url, {
      method: 'delete',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
  } catch (e) { console.log('[firebase] deleteAll(' + topic + ') 失敗: ' + e.message); }
}

/**
 * 清除單一 subkey
 */
function firebaseCacheDelete(topic, subkey) {
  try {
    const token = _getFirebaseAccessToken();
    const sk = subkey || '_default';
    const url = FIREBASE_DB_URL + '/cache/' + encodeURIComponent(topic) + '/' + encodeURIComponent(sk) + '.json';
    UrlFetchApp.fetch(url, {
      method: 'delete',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
  } catch (e) { console.log('[firebase] delete(' + topic + ',' + sk + ') 失敗: ' + e.message); }
}

/**
 * 批次清除多個 topic
 */
function firebaseInvalidate(topics) {
  if (!topics || !topics.length) return { invalidatedCount: 0, mode: 'none' };
  const uniqueTopics = Array.from(new Set(topics.map(t => String(t || '').trim()).filter(Boolean)));

  try {
    const token = _getFirebaseAccessToken();
    const updates = {};
    uniqueTopics.forEach(topic => { updates[topic] = null; });
    const response = UrlFetchApp.fetch(FIREBASE_DB_URL + '/cache.json', {
      method: 'patch',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify(updates),
      muteHttpExceptions: true
    });
    const status = response.getResponseCode();
    if (status < 200 || status >= 300) {
      throw new Error('RTDB batch invalidation HTTP ' + status + ': ' + response.getContentText());
    }
    return { invalidatedCount: uniqueTopics.length, mode: 'batch' };
  } catch (e) {
    console.log('[firebase] batch invalidate 失敗，改用逐筆刪除: ' + e.message);
    uniqueTopics.forEach(topic => firebaseCacheDeleteAll(topic));
    return { invalidatedCount: uniqueTopics.length, mode: 'fallback' };
  }
}

// ═══════════════════════════════════════════════════════════
//  📊 onEdit triggers — 偵測手動 Sheet 編輯，自動 invalidate
// ═══════════════════════════════════════════════════════════

function onEditMain(e) {
  const sheetName = e.range.getSheet().getName();
  try {
    if (sheetName === '會友名單') {
      invalidateAndRebuildMemberCache();
      firebaseInvalidate(['getAllMembers', 'getAllGroupMembers', 'getMemberSuggestions', 'getStats', 'getAllGroupsStats', 'getAdminGroupsList', 'ministry_getPageConfig', 'ministry_getGroupMembers', 'getAttendanceStats']);
    } else if (sheetName === '點名系統清單') {
      invalidateAndRebuildGroupConfigCache();
      firebaseInvalidate(['getGroupConfig']);
    } else if (/點名紀錄$/.test(sheetName)) {
      firebaseInvalidate(['getAttendanceStats', 'getAttendanceTrend', 'getWeeklyReport']);
    }
  } catch (err) { console.log('[onEditMain] ' + err); }
}

function onEditGroup(e) {
  const sheetName = e.range.getSheet().getName();
  try {
    if (sheetName === '小組清單') {
      _rebuildGroupsCache();
      firebaseInvalidate(['getGroups', 'getAdminGroupsList', 'ministry_getGroups']);
    } else if (/_點名紀錄$/.test(sheetName)) {
      firebaseInvalidate(['getStats', 'getAllGroupsStats', 'getWeeklyReport']);
    } else if (/_名單$/.test(sheetName)) {
      firebaseInvalidate(['getStats', 'ministry_getPageConfig']);
    }
  } catch (err) { console.log('[onEditGroup] ' + err); }
}

function onEditMinistry(e) {
  const sheetName = e.range.getSheet().getName();
  try {
    if (sheetName === 'Config') {
      _invalidateMinistryGroupsCache();
      _invalidateConfigDataCache();
      firebaseInvalidate(['ministry_getGroups', 'ministry_getPageConfig']);
    } else if (sheetName === '模板名稱') {
      firebaseInvalidate(['ministry_getTemplates']);
    } else if (sheetName !== '審計日誌') {
      firebaseInvalidate(['ministry_getPageConfig', 'ministry_getAggregatedReport']);
    }
  } catch (err) { console.log('[onEditMinistry] ' + err); }
}

/**
 * 一次性設定：為 4 個試算表建立 onEdit trigger
 * 在 GAS 編輯器手動執行一次即可
 */
function setupAllOnEditTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'onEditMain' || fn === 'onEditGroup' || fn === 'onEditMinistry' || fn === 'onEditWorshipSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('onEditMain').forSpreadsheet(getSS()).onEdit().create();
  ScriptApp.newTrigger('onEditGroup').forSpreadsheet(getGroupSS()).onEdit().create();
  ScriptApp.newTrigger('onEditMinistry').forSpreadsheet(getMinistrySS()).onEdit().create();
  try {
    ScriptApp.newTrigger('onEditWorshipSheet').forSpreadsheet(getWorshipSS()).onEdit().create();
    Logger.log('✅ 已建立 4 個 onEdit trigger（主日 / 小組 / 事工 / 敬拜團）');
  } catch (err) {
    Logger.log('⚠️ 建立敬拜團 onEdit trigger 失敗（可能尚未宣告 getWorshipSS 或無權限）：' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════
//  測試 / 維運工具
// ═══════════════════════════════════════════════════════════

/**
 * 測試 Service Account 是否設定正確
 * 在 GAS 編輯器選此函式 ▶ 執行，看 Logger 結果
 */
function testFirebaseAuth() {
  try {
    const token = _getFirebaseAccessToken();
    Logger.log('✅ OAuth token 取得成功（前 30 字元）：' + token.substring(0, 30) + '...');

    // 試寫一筆測試資料
    firebaseCacheSet('_test', 'gas-ping', { msg: 'hello from GAS', ts: new Date().toISOString() }, 60);
    Logger.log('✅ 已寫入測試資料到 cache/_test/gas-ping');
    Logger.log('   去 https://lkc1958june1-default-rtdb.asia-southeast1.firebasedatabase.app/cache/_test/gas-ping.json 看看');
  } catch (e) {
    Logger.log('❌ 認證或寫入失敗：' + e.message);
  }
}

/**
 * 清空所有 Firebase cache（緊急用）
 */
function firebaseCacheClearAll() {
  try {
    const token = _getFirebaseAccessToken();
    UrlFetchApp.fetch(FIREBASE_DB_URL + '/cache.json', {
      method: 'delete',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    Logger.log('✅ 已清空所有 Firebase cache');
  } catch (e) { Logger.log('❌ 清空失敗：' + e.message); }
}

/**
 * 重新初始化並建立全系統所有觸發器（4小時 keepWarm + 4個編輯快取失效觸發器）
 * 在 GAS 編輯器選擇此函數執行即可。
 */
function setupAllTriggers() {
  setupKeepWarmTrigger();
  setupAllOnEditTriggers();
  Logger.log('🎉 測試版主 GAS 系統所有觸發器已重新初始化設定完成！');
}
