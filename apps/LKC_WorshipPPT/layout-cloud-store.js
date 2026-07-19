(function(root, factory) {
  const api = factory(root);
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipLayoutCloud = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function(root) {
  const AUTH_EMAIL = 'worship-layout@lkc1958.org';
  const SHARED_LAYOUT_PATH = 'worshipPpt/layoutConfig/shared';

  function layoutPathForTemplate(templateId) {
    const safeTemplateId = String(templateId || 'taiwanese').trim();
    if (!/^[a-z0-9-]+$/i.test(safeTemplateId) || safeTemplateId === 'taiwanese') return SHARED_LAYOUT_PATH;
    return `worshipPpt/layoutConfig/templates/${safeTemplateId}`;
  }

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

  function chooseLayoutStateForLoad(localLayoutState, cloudLayoutState, localSyncPending) {
    if (localSyncPending) {
      return {
        layoutState: normalizeLayoutState(localLayoutState),
        source: 'local-pending'
      };
    }
    if (cloudLayoutState) {
      return {
        layoutState: normalizeLayoutState(cloudLayoutState),
        source: 'cloud'
      };
    }
    return {
      layoutState: normalizeLayoutState(localLayoutState),
      source: 'local'
    };
  }

  async function defaultFirebaseLoader() {
    const [appSdk, authSdk, databaseSdk] = await Promise.all([
      import('https://www.gstatic.com/firebasejs/10.13.0/firebase-app.js'),
      import('https://www.gstatic.com/firebasejs/10.13.0/firebase-auth.js'),
      import('https://www.gstatic.com/firebasejs/10.13.0/firebase-database.js')
    ]);
    const bootstrap = root.LKCFirebaseBootstrap;
    if (!bootstrap) throw new Error('Firebase 共用設定尚未載入');
    const app = bootstrap.getOrInitializeApp(appSdk);
    return {
      auth: authSdk.getAuth(app),
      database: databaseSdk.getDatabase(app),
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
    const layoutPath = layoutPathForTemplate(options.templateId);
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
      const snapshot = await sdk.get(sdk.ref(sdk.database, layoutPath));
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
      await sdk.set(sdk.ref(sdk.database, layoutPath), {
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
    layoutPathForTemplate,
    chooseLayoutStateForLoad,
    normalizeLayoutState,
    createLayoutCloudStore
  };
});
