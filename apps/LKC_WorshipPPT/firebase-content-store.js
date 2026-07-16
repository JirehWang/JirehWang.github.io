(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipFirebaseContent = api;
  if (typeof document !== 'undefined') root.worshipFirebaseContent = api.createFirebaseContentStore();
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  const CONTENT_ROOT = 'worshipPpt/content';

  const safeKey = value => String(value == null ? '' : value)
    .trim()
    .replace(/[.#$/\[\]\u0000-\u001f\u007f]/g, '_');

  function pathForAction(action, data = {}) {
    if (action === 'cal_getEvents') {
      const startDate = safeKey(data.startDate);
      const endDate = safeKey(data.endDate);
      return startDate && startDate === endDate
        ? `${CONTENT_ROOT}/services/${startDate}/calendar`
        : null;
    }
    if (action === 'cal_getPptLibraryIndex') return `${CONTENT_ROOT}/library/index`;
    if (action === 'cal_queryBible') {
      const version = safeKey(data.version || 'tghg');
      const book = safeKey(data.book);
      const chapter = safeKey(data.chap);
      const verses = safeKey(data.sec || '_all');
      return version && book && chapter
        ? `${CONTENT_ROOT}/bible/${version}/${book}/${chapter}/${verses}`
        : null;
    }
    return null;
  }

  async function defaultFirebaseLoader() {
    const [config, databaseSdk] = await Promise.all([
      import('../../firebase/firebase-config.js'),
      import('https://www.gstatic.com/firebasejs/10.13.0/firebase-database.js')
    ]);
    return {
      database: config.rtdb,
      ref: databaseSdk.ref,
      get: databaseSdk.get
    };
  }

  function createFirebaseContentStore(options = {}) {
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

    async function readPath(path) {
      if (!path) return null;
      const sdk = await firebase();
      const snapshot = await sdk.get(sdk.ref(sdk.database, path));
      return snapshot.exists() ? snapshot.val() : null;
    }

    return {
      readAction(action, data) {
        return readPath(pathForAction(action, data));
      },
      readServiceRecord(kind, date) {
        if (!['reports', 'praise'].includes(kind)) return Promise.resolve(null);
        const safeDate = safeKey(date);
        return readPath(safeDate ? `${CONTENT_ROOT}/services/${safeDate}/${kind}` : null);
      }
    };
  }

  return { CONTENT_ROOT, pathForAction, createFirebaseContentStore };
});
