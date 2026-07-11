const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function loadFirebaseSync(fetchImpl) {
  const source = fs.readFileSync(
    path.join(__dirname, '..', 'scratch_gas_sunday', 'FirebaseSync.js'),
    'utf8'
  );
  const cacheValues = new Map([['FB_ACCESS_TOKEN', 'cached-token']]);
  const context = {
    console,
    Date,
    JSON,
    CacheService: {
      getScriptCache: () => ({
        get: key => cacheValues.get(key) || null,
        put: (key, value) => cacheValues.set(key, value),
        remove: key => cacheValues.delete(key)
      })
    },
    UrlFetchApp: { fetch: fetchImpl },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Utilities: {}
  };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context;
}

test('firebaseInvalidate removes multiple topics with one RTDB PATCH request', () => {
  const calls = [];
  const context = loadFirebaseSync((url, options) => {
    calls.push({ url, options });
    return { getResponseCode: () => 200, getContentText: () => '{}' };
  });

  const result = context.firebaseInvalidate(['getGroups', 'getStats', 'getGroups']);

  assert.equal(calls.length, 1);
  assert.equal(calls[0].options.method, 'patch');
  assert.deepEqual(JSON.parse(calls[0].options.payload), {
    getGroups: null,
    getStats: null
  });
  assert.equal(result.invalidatedCount, 2);
  assert.equal(result.mode, 'batch');
});

test('firebaseInvalidate falls back to individual deletes when batch request fails', () => {
  const calls = [];
  const context = loadFirebaseSync((url, options) => {
    calls.push({ url, options });
    if (options.method === 'patch') {
      return { getResponseCode: () => 500, getContentText: () => 'failure' };
    }
    return { getResponseCode: () => 200, getContentText: () => '{}' };
  });

  const result = context.firebaseInvalidate(['getGroups', 'getStats']);

  assert.equal(calls.length, 3);
  assert.deepEqual(calls.map(call => call.options.method), ['patch', 'delete', 'delete']);
  assert.equal(result.invalidatedCount, 2);
  assert.equal(result.mode, 'fallback');
});
