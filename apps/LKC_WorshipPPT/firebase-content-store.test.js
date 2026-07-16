const test = require('node:test');
const assert = require('node:assert/strict');
const {
  CONTENT_ROOT,
  pathForAction,
  createFirebaseContentStore
} = require('./firebase-content-store.js');

test('maps PPT read actions to stable Firebase content paths', () => {
  assert.equal(CONTENT_ROOT, 'worshipPpt/content');
  assert.equal(pathForAction('cal_getEvents', {
    startDate: '2026-07-15', endDate: '2026-07-15'
  }), 'worshipPpt/content/services/2026-07-15/calendar');
  assert.equal(pathForAction('cal_getPptLibraryIndex', {}), 'worshipPpt/content/library/index');
  assert.equal(pathForAction('cal_queryBible', {
    book: '太', chap: 13, sec: '1-2', version: 'tghg'
  }), 'worshipPpt/content/bible/tghg/太/13/1-2');
  assert.equal(pathForAction('cal_getPptLibraryFile', { fileId: 'large-file' }), null);
});

test('returns Firebase content when present and null when it has not been synchronized', async () => {
  const values = new Map([
    ['worshipPpt/content/services/2026-07-15/calendar', { success: true, data: [{ title: '講道資訊－台語' }] }],
    ['worshipPpt/content/services/2026-07-15/reports', { announcements: ['報告一'] }]
  ]);
  const sdk = {
    database: {},
    ref: (_database, path) => ({ path }),
    get: async reference => ({
      exists: () => values.has(reference.path),
      val: () => values.get(reference.path)
    })
  };
  const store = createFirebaseContentStore({ loadFirebase: async () => sdk });

  assert.deepEqual(await store.readAction('cal_getEvents', {
    startDate: '2026-07-15', endDate: '2026-07-15'
  }), { success: true, data: [{ title: '講道資訊－台語' }] });
  assert.deepEqual(await store.readServiceRecord('reports', '2026-07-15'), { announcements: ['報告一'] });
  assert.equal(await store.readAction('cal_queryBible', {
    book: '太', chap: 13, sec: '1-2', version: 'tghg'
  }), null);
});
