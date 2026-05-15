// 🔥 Firebase 初始化設定
//
// 步驟：
//   1. 到 Firebase Console (https://console.firebase.google.com/) 開啟你的專案
//   2. 點 ⚙️ 專案設定 → 「一般設定」分頁 → 滑到最下方「您的應用程式」
//   3. 若還沒有 Web App，點 </> 圖示新增一個（不需要勾 Hosting）
//   4. 複製 firebaseConfig 物件的內容，貼到下方覆蓋 TODO 欄位
//   5. 在 Firebase Console → 建構 → Firestore Database → 建立資料庫
//      （第一次使用請選「以測試模式啟動」，之後再到「規則」分頁調整）
//
// 提醒：Web app 的 apiKey 並非機密金鑰，公開在前端是 Firebase 官方允許的做法。
//      真正的存取控制請靠 Firestore Security Rules 設定。

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-app.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "TODO_FILL_IN_apiKey",
  authDomain: "TODO_PROJECT_ID.firebaseapp.com",
  projectId: "TODO_PROJECT_ID",
  storageBucket: "TODO_PROJECT_ID.appspot.com",
  messagingSenderId: "TODO_FILL_IN_messagingSenderId",
  appId: "TODO_FILL_IN_appId"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export { app };
