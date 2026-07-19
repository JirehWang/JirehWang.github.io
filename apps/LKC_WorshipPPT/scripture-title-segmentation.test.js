const test = require('node:test');
const assert = require('node:assert/strict');
const { queryBibleViaReadApi } = require('./content-generators.js');
const { buildBiblePages } = require('./slide-production.js');

test('keeps each parsed scripture segment attached to its returned verses', async () => {
  const bibleService = {
    parseQuery: () => [
      { short: '創', bookName: '創世記', chap: 1, sec: '1-2' },
      { short: '約', bookName: '約翰福音', chap: 3, sec: '16' }
    ]
  };
  const readApi = async (_action, data) => ({
    records: data.book === '創'
      ? [{ chap: 1, sec: 1, text: '起初' }, { chap: 1, sec: 2, text: '地是空虛混沌' }]
      : [{ chap: 3, sec: 16, text: '神愛世人' }]
  });

  const records = await queryBibleViaReadApi('創世記1:1-2；約翰福音3:16', bibleService, readApi, 'unv');

  assert.deepEqual(records.map(record => ({
    bookName: record.queryBookName,
    chap: record.queryChap,
    sec: record.querySec,
    group: record.queryGroupKey
  })), [
    { bookName: '創世記', chap: 1, sec: '1-2', group: '創世記_1_1-2' },
    { bookName: '創世記', chap: 1, sec: '1-2', group: '創世記_1_1-2' },
    { bookName: '約翰福音', chap: 3, sec: '16', group: '約翰福音_3_16' }
  ]);
});

test('separates scripture query groups and titles each page from its actual verse range', () => {
  const pages = buildBiblePages('scripture', '聖經', '創世記1:1；約翰福音3:16-17', [
    { chap: 1, sec: 1, bible_text: '起初', queryBookName: '創世記', queryChap: 1, queryGroupKey: '創世記_1_1' },
    { chap: 3, sec: 16, bible_text: '神愛世人', queryBookName: '約翰福音', queryChap: 3, queryGroupKey: '約翰福音_3_16-17' },
    { chap: 3, sec: 17, bible_text: '不是定世人的罪', queryBookName: '約翰福音', queryChap: 3, queryGroupKey: '約翰福音_3_16-17' }
  ]);

  assert.equal(pages.length, 2);
  assert.equal(pages[0].title, '聖經－創世記 1:1');
  assert.equal(pages[0].body, '1 起初');
  assert.equal(pages[1].title, '聖經－約翰福音 3:16-17');
  assert.equal(pages[1].body, '16 神愛世人\n\n17 不是定世人的罪');
});
