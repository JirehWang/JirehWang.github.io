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

// 取得快取；過期或不存在時回傳 null
export async function cacheGet(topic, subkey) {
  const snap = await get(ref(rtdb, _path(topic, subkey)));
  if (!snap.exists()) return null;
  const data = snap.val();
  if (data && data.expiresAt && data.expiresAt < Date.now()) return null;
  return data ? data.value : null;
}

// 寫入快取；ttlSeconds 為存活秒數（預設 300，傳 0 或 null 表示永久）
export async function cacheSet(topic, subkey, value, ttlSeconds = 300) {
  const expiresAt = ttlSeconds ? Date.now() + ttlSeconds * 1000 : null;
  await set(ref(rtdb, _path(topic, subkey)), {
    value: value,
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
