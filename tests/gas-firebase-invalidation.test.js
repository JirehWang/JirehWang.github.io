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
    PropertiesService: { getScriptProperties: () => ({
      getProperty: key => cacheValues.get(`property:${key}`) || null,
      setProperty: (key, value) => cacheValues.set(`property:${key}`, value),
      getProperties: () => Object.fromEntries(
        Array.from(cacheValues.entries())
          .filter(([key]) => key.startsWith('property:'))
          .map(([key, value]) => [key.slice('property:'.length), value])
      ),
      deleteProperty: key => cacheValues.delete(`property:${key}`)
    }) },
    Utilities: {
      Charset: { UTF_8: 'UTF_8' },
      base64Encode: value => Buffer.from(value, 'utf8').toString('base64')
    }
  };
  vm.createContext(context);
  vm.runInContext(source, context);
  context.__cacheValues = cacheValues;
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

test('successful GAS read write-through creates a current long-lived Firebase entry', () => {
  const calls = [];
  const context = loadFirebaseSync((url, options) => {
    calls.push({ url, options });
    return { getResponseCode: () => 200, getContentText: () => '{}' };
  });

  const result = context.firebaseCacheWriteThrough('getGroups', {}, {
    status: 'success', data: ['fresh-from-gas']
  });

  assert.equal(result.ok, true);
  assert.equal(result.cacheRefreshPending, false);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].options.method, 'put');
  const entry = JSON.parse(calls[0].options.payload);
  assert.equal(entry.schemaVersion, 2);
  assert.equal(entry.generation, 1);
  assert.equal(entry.expiresAt, null);
  assert.deepEqual(entry.value.data, ['fresh-from-gas']);
});

test('a slow pre-mutation read cannot publish over a newer generation', () => {
  const calls = [];
  const context = loadFirebaseSync((url, options) => {
    calls.push({ url, options });
    return { getResponseCode: () => 200, getContentText: () => '{}' };
  });
  const staleSnapshotRevision = context.firebaseCaptureCacheRevision('getGroups');

  // Simulates a successful mutation committing while the read above is still
  // building its response. It invalidates the old Firebase topic and bumps
  // its generation before the slow request reaches write-through.
  context.firebaseInvalidate(['getGroups']);
  const stale = context.firebaseCacheWriteThrough('getGroups', {}, {
    status: 'success', data: ['old-sheet-snapshot']
  }, staleSnapshotRevision);

  assert.equal(stale.ok, true);
  assert.equal(stale.stale, true);
  assert.equal(stale.cacheRefreshPending, true);
  assert.equal(calls.length, 1, 'stale payload must not issue an RTDB PUT');
  assert.equal(calls[0].options.method, 'patch');
  assert.match(
    context.__cacheValues.get('property:FB_CACHE_REFRESH_PENDING_getGroups'),
    /stale source revision/
  );

  const latest = context.firebaseCacheWriteThrough('getGroups', {}, {
    status: 'success', data: ['latest-sheet-snapshot']
  }, context.firebaseCaptureCacheRevision('getGroups'));
  assert.equal(latest.ok, true);
  assert.equal(calls.length, 2);
  assert.equal(calls[1].options.method, 'put');
  assert.deepEqual(JSON.parse(calls[1].options.payload).value.data, ['latest-sheet-snapshot']);
  assert.equal(context.__cacheValues.has('property:FB_CACHE_REFRESH_PENDING_getGroups'), false);
});

test('Firebase write-back failure is recorded and does not throw or perform a second read', () => {
  const calls = [];
  const context = loadFirebaseSync((url, options) => {
    calls.push({ url, options });
    return { getResponseCode: () => 500, getContentText: () => 'unavailable' };
  });

  const result = context.firebaseCacheWriteThrough('getGroups', {}, {
    status: 'success', data: ['fresh-from-gas']
  });

  assert.equal(result.ok, false);
  assert.equal(result.cacheRefreshPending, true);
  assert.equal(calls.length, 1);
});

test('daily reconciliation retries only topics marked by an earlier Firebase failure', () => {
  let shouldFail = true;
  const calls = [];
  const context = loadFirebaseSync((url, options) => {
    calls.push({ url, options });
    const failed = shouldFail;
    shouldFail = false;
    return {
      getResponseCode: () => failed ? 500 : 200,
      getContentText: () => failed ? 'unavailable' : '{}'
    };
  });

  context.firebaseCacheWriteThrough('ministry_getGroups', {}, { status: 'success', data: [] });
  const result = context.firebaseReconcilePendingTopics();

  assert.equal(result.attempted, 1);
  assert.equal(result.repaired, 1);
  assert.equal(calls.length, 2);
  assert.equal(calls[1].options.method, 'patch');
  assert.deepEqual(JSON.parse(calls[1].options.payload), { ministry_getGroups: null });
});

test('pending reconciliation is bounded, deduplicated, and leaves remaining markers for later', () => {
  const calls = [];
  const context = loadFirebaseSync((url, options) => {
    calls.push({ url, options });
    return { getResponseCode: () => 200, getContentText: () => '{}' };
  });
  const topics = [
    'getGroups', 'getGroupConfig', 'getWeeklyReport', 'getAllMembers', 'getAdminGroupsList',
    'getAllGroupMembers', 'getMemberSuggestions', 'getSmartAttendanceList', 'checkGroupStatus',
    'getStats', 'getAllGroupsStats', 'getAttendanceStats', 'getAttendanceTrend',
    'getCategoryChartData', 'ministry_getGroups', 'ministry_getTemplates',
    'ministry_getAggregatedReport', 'ministry_getPageConfig', 'ministry_getGroupMembers',
    'ministry_getMemberSuggestions', 'getSchedule'
  ];
  for (const topic of topics) context._recordFirebaseSyncFailure(topic, 'retry');

  const first = context.firebaseReconcilePendingTopics();
  assert.equal(first.attempted, 20);
  assert.equal(first.repaired, 20);
  assert.equal(calls.length, 1);
  assert.equal(Object.keys(JSON.parse(calls[0].options.payload)).length, 20);

  const second = context.firebaseReconcilePendingTopics();
  assert.equal(second.attempted, 1);
  assert.equal(second.repaired, 1);
  assert.equal(calls.length, 2);
});

test('pending reconciliation drops expired and unapproved markers without a Firebase request', () => {
  const calls = [];
  const context = loadFirebaseSync((url, options) => {
    calls.push({ url, options });
    return { getResponseCode: () => 200, getContentText: () => '{}' };
  });
  const staleKey = 'property:FB_CACHE_REFRESH_PENDING_getGroups';
  const unsafeKey = 'property:FB_CACHE_REFRESH_PENDING_unknown';
  context.__cacheValues.set(staleKey, JSON.stringify({
    topic: 'getGroups', updatedAt: Date.now() - (8 * 24 * 60 * 60 * 1000), error: 'old'
  }));
  context.__cacheValues.set(unsafeKey, JSON.stringify({
    topic: 'unknownTopic', updatedAt: Date.now(), error: 'unsafe'
  }));

  const result = context.firebaseReconcilePendingTopics();
  assert.equal(result.attempted, 0);
  assert.equal(calls.length, 0);
  assert.equal(context.__cacheValues.has(staleKey), false);
  assert.equal(context.__cacheValues.has(unsafeKey), false);
});
