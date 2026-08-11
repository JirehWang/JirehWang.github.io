// 🗂️ Firebase Realtime Database 通用 JSON 快取（topic / subkey 雙層結構）
//
// 路徑結構：cache/{topic}/{subkey}
//   - topic    = action 名稱，例：getGroups / getStats
//   - subkey   = data 的 hash；無 data 時用 _default
//   - 子節點的 value: 任意 JSON-safe 物件
//   - 子節點的 expiresAt: unix-ms（null = 永不過期）
//
// 為什麼用兩層？
//   - cacheDeleteAll(topic) 可以用 1 次 remove 清掉整個 topic 下所有 subkey
//     方便寫入時 invalidate（例：addMember 時清掉所有 getStats 的 cache）

import { rtdb } from './firebase-config.js';
import {
  ref, get, set, remove
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-database.js";
import { CACHE_SCHEMA_VERSION, evaluateCacheEntry } from './cache-entry-contract.mjs';
import { createReadThrough } from './cache-read-through.mjs';

const ROOT = 'cache';
const _cacheLoads = createReadThrough();

function _path(topic, subkey) {
  return `${ROOT}/${topic}/${subkey || '_default'}`;
}

function _isInvalidApiResponse(value) {
  return value &&
    typeof value === 'object' &&
    Object.prototype.hasOwnProperty.call(value, 'status') &&
    value.status !== 'success';
}

// 🛡️ Firebase RTDB key 限制：不可空字串、不可含 . # $ / [ ]
//    遞迴清理物件：把不合法的 key 換掉（保留值），陣列照樣往內走
const _BAD_KEY_RE = /[.#$/\[\]\u0000-\u001f\u007f]/g;
function _sanitizeForFirebase(val) {
  if (val === null || val === undefined) return val;
  if (Array.isArray(val)) return val.map(_sanitizeForFirebase);
  if (typeof val !== 'object') return val;

  const out = {};
  for (const k in val) {
    if (!Object.prototype.hasOwnProperty.call(val, k)) continue;
    let safeKey = String(k);
    if (safeKey === '') safeKey = '_empty';
    if (_BAD_KEY_RE.test(safeKey)) safeKey = safeKey.replace(_BAD_KEY_RE, '_');
    // 若清理後與其他 key 衝突，加序號避免覆寫
    let finalKey = safeKey, i = 1;
    while (Object.prototype.hasOwnProperty.call(out, finalKey)) {
      finalKey = safeKey + '_' + (i++);
    }
    out[finalKey] = _sanitizeForFirebase(val[k]);
  }
  return out;
}

// 取得快取；只接受 GAS 寫入的目前 schema/generation。
// 舊格式安全視為 miss，交由 GAS 讀取並回寫，絕不由瀏覽器修復或刪除。
export async function cacheGet(topic, subkey) {
  const path = _path(topic, subkey);
  const snap = await get(ref(rtdb, path));
  if (!snap.exists()) return null;
  const data = snap.val();
  const evaluated = evaluateCacheEntry(data);
  return evaluated.hit ? evaluated.value : null;
}

// Browser clients are read-only.  Kept as a backwards-compatible no-op so
// legacy pages fail safe instead of bypassing GAS ownership of shared cache.
export async function cacheSet(topic, subkey, value, ttlSeconds = 300) {
  console.warn('[firebase-cache] Client cache writes are disabled; GAS owns shared cache:', topic);
  return false;
}

// 刪除單一 subkey 的快取
export async function cacheDelete(topic, subkey) {
  console.warn('[firebase-cache] Client cache writes are disabled; GAS owns shared cache:', topic);
  return false;
}

// 清除整個 topic 下所有 subkey 的快取（給寫入時 invalidate 用）
export async function cacheDeleteAll(topic) {
  console.warn('[firebase-cache] Client cache writes are disabled; GAS owns shared cache:', topic);
  return false;
}

// 高階：先讀 Firebase；miss 時只呼叫 loader 一次。loader 必須是 GAS，
// 由 GAS 在成功取得資料後 write-through，瀏覽器絕不自行回寫 Firebase。
export async function cacheGetOrFetch(topic, subkey, loader, ttlSeconds = 300) {
  const result = await _cacheLoads.getOrLoad(
    _path(topic, subkey),
    () => cacheGet(topic, subkey),
    loader
  );
  return result.value;
}

export async function cacheGetOrFetchWithMeta(topic, subkey, loader, ttlSeconds = 300) {
  return _cacheLoads.getOrLoad(
    _path(topic, subkey),
    () => cacheGet(topic, subkey),
    loader
  );
}
