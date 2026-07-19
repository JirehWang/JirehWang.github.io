const test = require('node:test');
const assert = require('node:assert/strict');
const { queryBibleViaReadApi } = require('./content-generators.js');

test('queries parsed Taiwanese Bible references through the GAS read API', async () => {
  const calls = [];
  const bibleService = {
    parseQuery() {
      return [{ short: '太', chap: 13, sec: '1-2' }];
    }
  };
  const readApi = async (action, data) => {
    calls.push({ action, data });
    return { success: true, records: [{ chap: 13, sec: 1, text: '測試經文' }] };
  };
  const records = await queryBibleViaReadApi('馬太福音13:1-2', bibleService, readApi);
  assert.deepEqual(calls, [{
    action: 'cal_queryBible',
    data: { book: '太', chap: 13, sec: '1-2', version: 'tghg' }
  }]);
  assert.deepEqual(records, [{ chap: 13, sec: 1, text: '測試經文', bible_text: '測試經文' }]);
});

test('uses the template Bible version for Mandarin scripture', async () => {
  const calls = [];
  const bibleService = { parseQuery: () => [{ short: '太', chap: 13, sec: '47-50' }] };
  const readApi = async (action, data) => {
    calls.push({ action, data });
    return { records: [{ sec: 47, bible_text: '天國又好像網撒在海裏' }] };
  };
  const records = await queryBibleViaReadApi('馬太福音13:47-50', bibleService, readApi, 'unv');
  assert.equal(calls[0].data.version, 'unv');
  assert.equal(records[0].bible_text, '天國又好像網撒在海裏');
});
