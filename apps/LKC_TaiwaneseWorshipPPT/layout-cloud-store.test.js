const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const {
  AUTH_EMAIL,
  SHARED_LAYOUT_PATH,
  createLayoutCloudStore,
  normalizeLayoutState
} = require('./layout-cloud-store.js');

function firebaseFixture(initialValue = null) {
  const writes = [];
  const auth = { currentUser: null };
  const api = {
    auth,
    database: { name: 'test-db' },
    inMemoryPersistence: { name: 'memory' },
    ref: (_database, path) => ({ path }),
    get: async reference => ({
      exists: () => initialValue !== null,
      val: () => initialValue,
      ref: reference
    }),
    set: async (reference, value) => writes.push({ reference, value }),
    serverTimestamp: () => 123456789,
    setPersistence: async (_auth, persistence) => {
      assert.equal(persistence, api.inMemoryPersistence);
    },
    signInWithEmailAndPassword: async (_auth, email, password) => {
      assert.equal(email, AUTH_EMAIL);
      if (password !== 'test-secret') throw Object.assign(new Error('wrong password'), { code: 'auth/invalid-credential' });
      auth.currentUser = { uid: 'editor-1', email };
      return { user: auth.currentUser };
    },
    signOut: async () => { auth.currentUser = null; }
  };
  return { api, writes };
}

test('normalizes the shared cloud payload to layout groups and assignments only', () => {
  assert.deepEqual(normalizeLayoutState({
    groups: { scripture: { id: 'scripture', pageIds: ['scripture:1'], params: { contentSize: 42 } } },
    pageAssignments: { 'scripture:1': 'scripture' },
    backgroundImage: 'data:image/png;base64,too-large-for-layout-state'
  }), {
    groups: { scripture: { id: 'scripture', pageIds: ['scripture:1'], params: { contentSize: 42 } } },
    pageAssignments: { 'scripture:1': 'scripture' }
  });
});

test('loads the one church-wide layout from the dedicated RTDB path', async () => {
  const state = { groups: { shared: { id: 'shared' } }, pageAssignments: {} };
  const fixture = firebaseFixture({ schemaVersion: 1, layoutState: state });
  const store = createLayoutCloudStore({ loadFirebase: async () => fixture.api });

  assert.deepEqual(await store.load(), state);
  assert.equal(SHARED_LAYOUT_PATH, 'worshipPpt/layoutConfig/shared');
});

test('refuses cloud writes while the layout editor is locked', async () => {
  const fixture = firebaseFixture();
  const store = createLayoutCloudStore({ loadFirebase: async () => fixture.api });

  await assert.rejects(() => store.save({ groups: {}, pageAssignments: {} }), /尚未解鎖/);
  assert.equal(fixture.writes.length, 0);
});

test('passes the entered password to Firebase Auth and keeps auth in memory', async () => {
  const fixture = firebaseFixture();
  const store = createLayoutCloudStore({ loadFirebase: async () => fixture.api });

  await assert.rejects(() => store.unlock('wrong'), /密碼錯誤/);
  assert.equal(await store.unlock('test-secret'), true);
  assert.equal(await store.isUnlocked(), true);

  await store.save({ groups: { shared: { id: 'shared' } }, pageAssignments: {} });
  assert.deepEqual(fixture.writes, [{
    reference: { path: SHARED_LAYOUT_PATH },
    value: {
      schemaVersion: 1,
      layoutState: { groups: { shared: { id: 'shared' } }, pageAssignments: {} },
      updatedAt: 123456789,
      updatedBy: 'editor-1'
    }
  }]);

  await store.lock();
  assert.equal(await store.isUnlocked(), false);
});

test('allows a Firebase dependency load to be retried after a temporary failure', async () => {
  const fixture = firebaseFixture();
  let attempts = 0;
  const store = createLayoutCloudStore({
    loadFirebase: async () => {
      attempts += 1;
      if (attempts === 1) throw new Error('temporary network failure');
      return fixture.api;
    }
  });

  await assert.rejects(() => store.load(), /temporary network failure/);
  assert.equal(await store.load(), null);
  assert.equal(attempts, 2);
});

test('loads the cloud store before the locked layout UI and uses a password field', () => {
  const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
  assert.ok(html.indexOf('layout-cloud-store.js') < html.indexOf('layout-groups.js'));
  assert.match(html, /id="layout-unlock-password" type="password"/);
  assert.match(html, /id="layout-lock-toggle"/);
});
