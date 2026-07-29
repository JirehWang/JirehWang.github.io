const assert = require('node:assert/strict');
const test = require('node:test');

const AttendanceSearchScroll = require('../apps/LKC_SundayserviceAttendance/attendance-search-scroll.js');

function makeStorage() {
  const values = new Map();
  return {
    getItem(key) {
      return values.has(key) ? values.get(key) : null;
    },
    setItem(key, value) {
      values.set(key, String(value));
    },
    removeItem(key) {
      values.delete(key);
    },
    has(key) {
      return values.has(key);
    }
  };
}

test('search scroll cache keys are scoped by attendance type and date', () => {
  assert.equal(
    AttendanceSearchScroll.getKey('禮拜', '2026-07-29'),
    'LKC:attendance-search-scroll:禮拜:2026-07-29'
  );
  assert.notEqual(
    AttendanceSearchScroll.getKey('禮拜', '2026-07-29'),
    AttendanceSearchScroll.getKey('主日學', '2026-07-29')
  );
});

test('saved pre-search anchor is consumed once when the search is cleared', () => {
  const storage = makeStorage();
  const key = AttendanceSearchScroll.getKey('禮拜', '2026-07-29');
  const anchor = { key: 'LK00123', offset: 96, scrollTop: 1280 };

  assert.equal(AttendanceSearchScroll.save(storage, key, anchor), true);
  assert.deepEqual(AttendanceSearchScroll.consume(storage, key), anchor);
  assert.equal(AttendanceSearchScroll.consume(storage, key), null);
  assert.equal(storage.has(key), false);
});

test('storage failures do not throw or create a false restore anchor', () => {
  const brokenStorage = {
    getItem() { throw new Error('storage blocked'); },
    setItem() { throw new Error('storage blocked'); },
    removeItem() { throw new Error('storage blocked'); }
  };
  const key = AttendanceSearchScroll.getKey('禮拜', '2026-07-29');

  assert.equal(AttendanceSearchScroll.save(brokenStorage, key, { scrollTop: 10 }), false);
  assert.equal(AttendanceSearchScroll.consume(brokenStorage, key), null);
});
