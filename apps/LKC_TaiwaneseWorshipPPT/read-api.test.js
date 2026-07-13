const test = require('node:test');
const assert = require('node:assert/strict');
const { buildJsonpUrl, read } = require('./read-api.js');

test('builds a JSONP URL for read-only GAS actions from file pages', () => {
  const url = new URL(buildJsonpUrl(
    'https://script.google.com/macros/s/example/exec',
    'cal_getEvents',
    { startDate: '2026-07-12', endDate: '2026-07-12' },
    'ChurchApp-2026',
    '__lkcCallback1'
  ));
  assert.equal(url.searchParams.get('action'), 'cal_getEvents');
  assert.equal(url.searchParams.get('token'), 'ChurchApp-2026');
  assert.equal(url.searchParams.get('callback'), '__lkcCallback1');
  assert.deepEqual(JSON.parse(url.searchParams.get('data')), {
    startDate: '2026-07-12', endDate: '2026-07-12'
  });
});

test('falls back to JSONP when API readiness fails with a network error', async () => {
  const previous = {
    ensureAPIReady: global.ensureAPIReady,
    churchAPI: global.churchAPI,
    GAS_URL: global.GAS_URL,
    AUTH_TOKEN: global.AUTH_TOKEN,
    location: global.location,
    document: global.document
  };

  global.ensureAPIReady = async () => { throw new TypeError('Failed to fetch'); };
  global.churchAPI = async () => { throw new Error('churchAPI should not run'); };
  global.GAS_URL = 'https://script.google.com/macros/s/example/exec';
  global.AUTH_TOKEN = 'ChurchApp-2026';
  global.location = { protocol: 'https:' };
  global.document = {
    createElement() {
      return { remove() {} };
    },
    head: {
      appendChild(script) {
        const callback = new URL(script.src).searchParams.get('callback');
        queueMicrotask(() => global[callback]({ success: true, data: ['fallback'] }));
      }
    }
  };

  try {
    const result = await read('cal_getEvents', { startDate: '2026-07-12' });
    assert.deepEqual(result, { success: true, data: ['fallback'] });
  } finally {
    Object.assign(global, previous);
  }
});
