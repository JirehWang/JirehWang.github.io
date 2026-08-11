const assert = require('node:assert/strict');
const test = require('node:test');

const {
  refreshCacheGroup,
  refreshAllCacheGroups,
  uniqueValues
} = require('../admin-cache-coordinator.js');

test('uniqueValues keeps first occurrence and removes empty values', () => {
  assert.deepEqual(uniqueValues(['a', '', 'a', null, 'b', undefined]), ['a', 'b']);
});

test('single-group refresh delegates cache ownership to the backend', async () => {
  const calls = [];
  const result = await refreshCacheGroup({
    topics: ['getTrackingCases', 'getClosedCases'],
    backendRefresh: 'newfamily_refreshCaches'
  }, {
    deleteTopics: async topics => {
      calls.push(['delete', ...topics]);
      return topics.length;
    },
    refreshBackend: async action => calls.push(['refresh', action])
  });

  assert.deepEqual(calls, [['refresh', 'newfamily_refreshCaches']]);
  assert.equal(result.topicCount, 2);
});

test('all-group refresh warms backends with bounded concurrency without browser deletion', async () => {
  const calls = [];
  let active = 0;
  let maxActive = 0;
  const groups = {
    one: { topics: ['a', 'shared'], backendRefresh: 'refreshOne' },
    two: { topics: ['shared', 'b'], backendRefresh: 'refreshTwo' },
    three: { topics: ['c'], backendRefresh: 'refreshThree' }
  };

  const result = await refreshAllCacheGroups(groups, {
    deleteTopics: async topics => calls.push(['delete', ...topics]),
    refreshBackend: async action => {
      active += 1;
      maxActive = Math.max(maxActive, active);
      calls.push(['refresh-start', action]);
      await new Promise(resolve => setTimeout(resolve, 5));
      calls.push(['refresh-end', action]);
      active -= 1;
    }
  }, { concurrency: 2 });

  assert.equal(maxActive, 2);
  assert.equal(result.topicCount, 4);
  assert.equal(result.backendCount, 3);
});

test('backend refresh failure is surfaced without any browser-side deletion', async () => {
  await assert.rejects(
    refreshCacheGroup({ topics: ['a'], backendRefresh: 'refreshOne' }, {
      deleteTopics: async () => { throw new Error('must not be called'); },
      refreshBackend: async () => { throw new Error('backend unavailable'); }
    }),
    /backend unavailable/
  );
});
