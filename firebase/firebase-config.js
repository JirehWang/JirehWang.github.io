// 🔥 Firebase 初始化設定（lkc1958june1 專案）
//
// 此設定由前端載入；apiKey 公開於前端是 Firebase 官方允許的做法。
// 真正的存取控制由 Firestore Security Rules / Realtime DB Rules 把關。
//
// 此專案同時啟用 Firestore + Realtime Database，
//  - Firestore 主要做 cache 層（cacheGet/cacheSet）
//  - Realtime DB 預留給未來即時推播 / 多裝置即時同步用

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-app.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore.js";
import { getDatabase } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-database.js";

const firebaseConfig = {
  apiKey:            "AIzaSyCyi5nWpuNpFcUmNY6WmpmGpf6J1Bi06UY",
  authDomain:        "lkc1958june1.firebaseapp.com",
  databaseURL:       "https://lkc1958june1-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId:         "lkc1958june1",
  storageBucket:     "lkc1958june1.firebasestorage.app",
  messagingSenderId: "245519602141",
  appId:             "1:245519602141:web:73537df7c2dc6485e5a634",
  measurementId:     "G-Y26JGTG9WH"
};

const app = initializeApp(firebaseConfig);
export const db  = getFirestore(app);   // Firestore 實例（cache 用）
export const rtdb = getDatabase(app);    // Realtime DB 實例（預留）
export { app };
