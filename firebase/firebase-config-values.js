(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.LKCFirebaseBootstrap = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  const config = Object.freeze({
    apiKey: 'AIzaSyCyi5nWpuNpFcUmNY6WmpmGpf6J1Bi06UY',
    authDomain: 'lkc1958june1.firebaseapp.com',
    databaseURL: 'https://lkc1958june1-default-rtdb.asia-southeast1.firebasedatabase.app',
    projectId: 'lkc1958june1',
    storageBucket: 'lkc1958june1.firebasestorage.app',
    messagingSenderId: '245519602141',
    appId: '1:245519602141:web:73537df7c2dc6485e5a634',
    measurementId: 'G-Y26JGTG9WH'
  });

  function getOrInitializeApp(appSdk) {
    if (!appSdk || typeof appSdk.getApps !== 'function' || typeof appSdk.initializeApp !== 'function') {
      throw new Error('Firebase App SDK 尚未載入');
    }
    return appSdk.getApps().length
      ? appSdk.getApp()
      : appSdk.initializeApp(config);
  }

  return { config, getOrInitializeApp };
});
