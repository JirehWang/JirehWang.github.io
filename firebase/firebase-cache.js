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

const ROOT = 'cache';

function _path(topic, subkey) {
  return `${ROOT}/${topic}/${subkey || '_default'}`;
}

// 🛡️ Firebase RTDB key 限制：不可空字串、不可含 . # $ / [ ]
//    遞迴清理物件：把不合法的 key 換掉（保留值），陣列照樣往內走
const _BAD_KEY_RE = /[.#$/\[\]]/g;
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

// 取得快取；過期或不存在時回傳 null
// ♻️ 自我清理：讀到過期 entry 順手 remove，避免長期累積殭屍資料
//    （不 await 刪除，不影響回應速度）
export async function cacheGet(topic, subkey) {
  const path = _path(topic, subkey);
  const snap = await get(ref(rtdb, path));
  if (!snap.exists()) return null;
  const data = snap.val();
  if (data && data.expiresAt && data.expiresAt < Date.now()) {
    // 過期 → 順手刪掉
    remove(ref(rtdb, path)).catch(() => {});
    return null;
  }
  return data ? data.value : null;
}

// 寫入快取；ttlSeconds 為存活秒數（預設 300，傳 0 或 null 表示永久）
// 寫入前自動清理不合法的 key（防止 GAS 回傳含空字串 key / 含 . # $ / [ ] 的物件）
export async function cacheSet(topic, subkey, value, ttlSeconds = 300) {
  const expiresAt = ttlSeconds ? Date.now() + ttlSeconds * 1000 : null;
  await set(ref(rtdb, _path(topic, subkey)), {
    value: _sanitizeForFirebase(value),
    expiresAt: expiresAt,
    updatedAt: Date.now()
  });
}

// 刪除單一 subkey 的快取
export async function cacheDelete(topic, subkey) {
  await remove(ref(rtdb, _path(topic, subkey)));
}

// 清除整個 topic 下所有 subkey 的快取（給寫入時 invalidate 用）
export async function cacheDeleteAll(topic) {
  await remove(ref(rtdb, `${ROOT}/${topic}`));
}

// 高階：先讀快取，沒有或過期時呼叫 loader() 並寫回
export async function cacheGetOrFetch(topic, subkey, loader, ttlSeconds = 300) {
  const hit = await cacheGet(topic, subkey);
  if (hit !== null) return hit;
  const fresh = await loader();
  await cacheSet(topic, subkey, fresh, ttlSeconds);
  return fresh;
}
