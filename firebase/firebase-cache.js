// 🗂️ Firebase Firestore 通用 JSON 快取
//
// 用途：把 GAS 回傳的 JSON 暫存到 Firestore 的 `cache` collection，
//      降低 GAS 呼叫頻率、加快讀取速度（Firestore 全球 CDN）。
//
// 資料結構（cache/{key} 文件）：
//   {
//     value:     <任意 JSON-safe 物件>,
//     expiresAt: <Timestamp>,  // null 表示永不過期
//     updatedAt: <Timestamp>
//   }

import { db } from './firebase-config.js';
import {
  doc, getDoc, setDoc, deleteDoc,
  serverTimestamp, Timestamp
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore.js";

const COLLECTION = 'cache';

// 取得快取；過期或不存在時回傳 null
export async function cacheGet(key) {
  const snap = await getDoc(doc(db, COLLECTION, key));
  if (!snap.exists()) return null;
  const data = snap.data();
  if (data.expiresAt && data.expiresAt.toMillis() < Date.now()) {
    return null;
  }
  return data.value;
}

// 寫入快取；ttlSeconds 為存活秒數（預設 300 秒，傳 0 或 null 表示永久）
export async function cacheSet(key, value, ttlSeconds = 300) {
  const expiresAt = ttlSeconds
    ? Timestamp.fromMillis(Date.now() + ttlSeconds * 1000)
    : null;
  await setDoc(doc(db, COLLECTION, key), {
    value,
    expiresAt,
    updatedAt: serverTimestamp()
  });
}

// 刪除快取
export async function cacheDelete(key) {
  await deleteDoc(doc(db, COLLECTION, key));
}

// 高階：先讀快取，沒有或過期時呼叫 loader() 並寫回
export async function cacheGetOrFetch(key, loader, ttlSeconds = 300) {
  const hit = await cacheGet(key);
  if (hit !== null) return hit;
  const fresh = await loader();
  await cacheSet(key, fresh, ttlSeconds);
  return fresh;
}
