const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const bootstrap = require('../../firebase/firebase-config-values.js');

test('provides one shared Firebase config and reuses the default app', () => {
  assert.equal(bootstrap.config.projectId, 'lkc1958june1');
  assert.match(bootstrap.config.databaseURL, /asia-southeast1\.firebasedatabase\.app$/);

  const existing = { name: '[DEFAULT]' };
  let initialized = 0;
  assert.equal(bootstrap.getOrInitializeApp({
    getApps: () => [existing],
    getApp: () => existing,
    initializeApp: () => { initialized += 1; }
  }), existing);
  assert.equal(initialized, 0);

  const created = { name: '[DEFAULT]', created: true };
  assert.equal(bootstrap.getOrInitializeApp({
    getApps: () => [],
    getApp: () => { throw new Error('not initialized'); },
    initializeApp: config => {
      initialized += 1;
      assert.equal(config, bootstrap.config);
      return created;
    }
  }), created);
  assert.equal(initialized, 1);
});

test('loads the classic Firebase bootstrap before stores and avoids relative dynamic imports', () => {
  const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
  const contentStore = fs.readFileSync(path.join(__dirname, 'firebase-content-store.js'), 'utf8');
  const layoutStore = fs.readFileSync(path.join(__dirname, 'layout-cloud-store.js'), 'utf8');

  assert.ok(html.indexOf('../../firebase/firebase-config-values.js') < html.indexOf('firebase-content-store.js'));
  assert.doesNotMatch(contentStore, /import\(['"]\.\.\/\.\.\/firebase\/firebase-config\.js['"]\)/);
  assert.doesNotMatch(layoutStore, /import\(['"]\.\.\/\.\.\/firebase\/firebase-config\.js['"]\)/);
  assert.match(contentStore, /root\.LKCFirebaseBootstrap/);
  assert.match(contentStore, /bootstrap\.getOrInitializeApp/);
  assert.match(layoutStore, /root\.LKCFirebaseBootstrap/);
  assert.match(layoutStore, /bootstrap\.getOrInitializeApp/);
});
