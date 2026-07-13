const test = require('node:test');
const assert = require('node:assert/strict');
const { buildJsonpUrl } = require('./read-api.js');

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
