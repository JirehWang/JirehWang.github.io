// 🗂️ Firebase Realtime Database 通用 JSON 快取
//
// 用途：把 GAS 回傳的 JSON 暫存到 Realtime Database 的 /cache 路徑，
//      降低 GAS 呼叫頻率、加快讀取速度（RTDB 全球低延遲）。
//
// 資料結構（/cache/{key}）：
//   {
//     value:     <任意 JSON-safe 物件>,
//     expiresAt: <unix-ms>  // null 表示永不過期
//     updatedAt: <unix-ms>
//   }

import { rtdb } from './firebase-config.js';
import {
  ref, get, set, remove
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-database.js";

const ROOT = 'cache';

// 取得快取；過期或不存在時回傳 null
export async function cacheGet(key) {
  const snap = await get(ref(rtdb, `${ROOT}/${key}`));
  if (!snap.exists()) return null;
  const data = snap.val();
  if (data && data.expiresAt && data.expiresAt < Date.now()) {
    return null;
  }
  return data ? data.value : null;
}

// 寫入快取；ttlSeconds 為存活秒數（預設 300 秒，傳 0 或 null 表示永久）
export async function cacheSet(key, value, ttlSeconds = 300) {
  const expiresAt = ttlSeconds ? Date.now() + ttlSeconds * 1000 : null;
  await set(ref(rtdb, `${ROOT}/${key}`), {
    value: value,
    expiresAt: expiresAt,
    updatedAt: Date.now()
  });
}

// 刪除快取
export async function cacheDelete(key) {
  await remove(ref(rtdb, `${ROOT}/${key}`));
}

// 高階：先讀快取，沒有或過期時呼叫 loader() 並寫回
export async function cacheGetOrFetch(key, loader, ttlSeconds = 300) {
  const hit = await cacheGet(key);
  if (hit !== null) return hit;
  const fresh = await loader();
  await cacheSet(key, fresh, ttlSeconds);
  return fresh;
}
