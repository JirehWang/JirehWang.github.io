// 🗂️ Firebase Realtime Database 通用 JSON 快取
//
// 用途：把 GAS 回傳的 JSON 暫存到 RTDB 的 /cache 節點，
//      降低 GAS 呼叫頻率、加快讀取速度。
//
// 資料結構 (/cache/{key}):
//   {
//     value:     <任意 JSON-safe 物件>,
//     expiresAt: <number ms epoch>,  // 0 表示永不過期
//     updatedAt: <number ms epoch>
//   }
//
// 注意：RTDB 的 key 不能包含 . # $ [ ] /
//      若 key 來自外部輸入，請用 encodeKey() 編碼。

import { db } from './firebase-config.js';
import {
  ref, get, set, remove
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-database.js";

const ROOT = 'cache';

// 把不合法字元轉成底線，避免 RTDB key 限制
export function encodeKey(key) {
  return String(key).replace(/[.#$\[\]\/]/g, '_');
}

// 取得快取；過期或不存在時回傳 null
export async function cacheGet(key) {
  const snap = await get(ref(db, `${ROOT}/${encodeKey(key)}`));
  if (!snap.exists()) return null;
  const data = snap.val();
  if (data.expiresAt && data.expiresAt < Date.now()) {
    return null;
  }
  return data.value;
}

// 寫入快取；ttlSeconds 為存活秒數（預設 300 秒，傳 0 或 null 表示永久）
export async function cacheSet(key, value, ttlSeconds = 300) {
  const now = Date.now();
  await set(ref(db, `${ROOT}/${encodeKey(key)}`), {
    value,
    expiresAt: ttlSeconds ? now + ttlSeconds * 1000 : 0,
    updatedAt: now
  });
}

// 刪除快取
export async function cacheDelete(key) {
  await remove(ref(db, `${ROOT}/${encodeKey(key)}`));
}

// 高階：先讀快取，沒有或過期時呼叫 loader() 並寫回
export async function cacheGetOrFetch(key, loader, ttlSeconds = 300) {
  const hit = await cacheGet(key);
  if (hit !== null) return hit;
  const fresh = await loader();
  await cacheSet(key, fresh, ttlSeconds);
  return fresh;
}
