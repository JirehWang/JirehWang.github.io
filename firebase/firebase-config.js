// 🔥 Firebase 初始化設定 (使用 Realtime Database)
//
// 提醒：Web app 的 apiKey 並非機密金鑰，公開在前端是 Firebase 官方允許的做法。
//      真正的存取控制請靠 Realtime Database Rules 設定。

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-app.js";
import { getDatabase } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-database.js";

const firebaseConfig = {
  apiKey: "AIzaSyCyi5nWpuNpFcUmNY6WmpmGpf6J1Bi06UY",
  authDomain: "lkc1958june1.firebaseapp.com",
  databaseURL: "https://lkc1958june1-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "lkc1958june1",
  storageBucket: "lkc1958june1.firebasestorage.app",
  messagingSenderId: "245519602141",
  appId: "1:245519602141:web:73537df7c2dc6485e5a634",
  measurementId: "G-Y26JGTG9WH"
};

const app = initializeApp(firebaseConfig);
export const db = getDatabase(app);
export { app };
