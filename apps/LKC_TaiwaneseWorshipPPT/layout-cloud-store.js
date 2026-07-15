(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipLayoutCloud = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  const AUTH_EMAIL = 'worship-layout@lkc1958.org';
  const SHARED_LAYOUT_PATH = 'worshipPpt/layoutConfig/shared';

  function jsonObject(value) {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return {};
    return JSON.parse(JSON.stringify(value));
  }

  function normalizeLayoutState(value) {
    const source = value && typeof value === 'object' ? value : {};
    const normalized = {
      groups: jsonObject(source.groups),
      pageAssignments: jsonObject(source.pageAssignments)
    };
    const hymnOpacityBySection = {};
    Object.entries(jsonObject(source.hymnOpacityBySection)).forEach(([sectionId, value]) => {
      const opacity = Number(value);
      if (/^[a-z0-9-]+$/i.test(sectionId) && opacity >= 40 && opacity <= 80) {
        hymnOpacityBySection[sectionId] = opacity;
      }
    });
    if (Object.keys(hymnOpacityBySection).length) normalized.hymnOpacityBySection = hymnOpacityBySection;
    const outputScale = {};
    ['text', 'image'].forEach(key => {
      const scale = Number(source.outputScale && source.outputScale[key]);
      if (scale >= 80 && scale <= 120) outputScale[key] = scale;
    });
    if (Object.keys(outputScale).length) normalized.outputScale = outputScale;
    return normalized;
  }

  async function defaultFirebaseLoader() {
    const [config, authSdk, databaseSdk] = await Promise.all([
      import('../../firebase/firebase-config.js'),
      import('https://www.gstatic.com/firebasejs/10.13.0/firebase-auth.js'),
      import('https://www.gstatic.com/firebasejs/10.13.0/firebase-database.js')
    ]);
    return {
      auth: authSdk.getAuth(config.app),
      database: config.rtdb,
      inMemoryPersistence: authSdk.inMemoryPersistence,
      setPersistence: authSdk.setPersistence,
      signInWithEmailAndPassword: authSdk.signInWithEmailAndPassword,
      signOut: authSdk.signOut,
      ref: databaseSdk.ref,
      get: databaseSdk.get,
      set: databaseSdk.set,
      serverTimestamp: databaseSdk.serverTimestamp
    };
  }

  function createLayoutCloudStore(options = {}) {
    const loadFirebase = options.loadFirebase || defaultFirebaseLoader;
    let firebasePromise = null;
    const firebase = () => {
      if (!firebasePromise) {
        firebasePromise = Promise.resolve().then(loadFirebase).catch(error => {
          firebasePromise = null;
          throw error;
        });
      }
      return firebasePromise;
    };

    async function load() {
      const sdk = await firebase();
      const snapshot = await sdk.get(sdk.ref(sdk.database, SHARED_LAYOUT_PATH));
      if (!snapshot.exists()) return null;
      const value = snapshot.val();
      if (!value || value.schemaVersion !== 1 || !value.layoutState) return null;
      return normalizeLayoutState(value.layoutState);
    }

    async function isUnlocked() {
      const sdk = await firebase();
      return Boolean(sdk.auth.currentUser && sdk.auth.currentUser.email === AUTH_EMAIL);
    }

    async function unlock(password) {
      const sdk = await firebase();
      try {
        await sdk.setPersistence(sdk.auth, sdk.inMemoryPersistence);
        await sdk.signInWithEmailAndPassword(sdk.auth, AUTH_EMAIL, String(password || ''));
        return isUnlocked();
      } catch (error) {
        const code = String(error && error.code || '');
        if (/invalid-credential|wrong-password|invalid-login-credentials|user-not-found/.test(code)) {
          throw new Error('版面配置解鎖密碼錯誤');
        }
        throw error;
      }
    }

    async function save(layoutState) {
      const sdk = await firebase();
      const user = sdk.auth.currentUser;
      if (!user || user.email !== AUTH_EMAIL) throw new Error('版面配置尚未解鎖');
      await sdk.set(sdk.ref(sdk.database, SHARED_LAYOUT_PATH), {
        schemaVersion: 1,
        layoutState: normalizeLayoutState(layoutState),
        updatedAt: sdk.serverTimestamp(),
        updatedBy: user.uid
      });
    }

    async function lock() {
      const sdk = await firebase();
      await sdk.signOut(sdk.auth);
    }

    return { load, save, unlock, lock, isUnlocked };
  }

  return {
    AUTH_EMAIL,
    SHARED_LAYOUT_PATH,
    normalizeLayoutState,
    createLayoutCloudStore
  };
});
