const test = require('node:test');
const assert = require('node:assert/strict');
const {
  buildBiblePages,
  buildDeckEntries,
  paginateFixedText,
  createLayoutGroup,
  updateLayoutGroup,
  detachPagesFromLayoutGroup,
  normalizeColor
} = require('./slide-production.js');

test('normalizes solid background and text colors to safe hex values', () => {
  assert.equal(normalizeColor('#2f5d50', '#ffffff'), '#2f5d50');
  assert.equal(normalizeColor('#ABC', '#ffffff'), '#aabbcc');
  assert.equal(normalizeColor('not-a-color', '#ffffff'), '#ffffff');
});

test('flattens PPT chapters into one continuous deck order', () => {
  const deck = buildDeckEntries([
    { sectionId: 'cover', label: '首頁', pages: [{ id: 'cover:1' }] },
    { sectionId: 'creed', label: '使徒信經', pages: [{ id: 'creed:1' }, { id: 'creed:2' }] }
  ]);
  assert.deepEqual(deck.map(item => item.id), ['cover:1', 'creed:1', 'creed:2']);
  assert.deepEqual(deck.map(item => item.deckNumber), [1, 2, 3]);
  assert.equal(deck[2].sectionLabel, '使徒信經');
});

test('keeps fixed liturgy as one source field while rebuilding the original page count', () => {
  const source = '第一段。\n\n第二段較長，仍應完整保留。\n\n第三段。\n\n第四段結束。';
  const pages = paginateFixedText(source, [20, 35, 45]);
  assert.equal(pages.length, 3);
  assert.equal(pages.map(page => page.body).join('\n\n'), source);
  assert.ok(pages.every(page => page.body.length > 0));
});

test('turns a scripture query result into two-verses-per-slide pages', () => {
  const pages = buildBiblePages('scripture', '聖經', '馬太福音13:1-5', [
    { chap: 13, sec: 1, bible_text: '第一節' },
    { chap: 13, sec: 2, bible_text: '第二節' },
    { chap: 13, sec: 3, bible_text: '第三節' },
    { chap: 13, sec: 4, bible_text: '第四節' },
    { chap: 13, sec: 5, bible_text: '第五節' }
  ]);
  assert.equal(pages.length, 3);
  assert.equal(pages[0].id, 'scripture:1');
  assert.equal(pages[0].title, '聖經－馬太福音13:1-5');
  assert.equal(pages[0].body, '1 第一節\n\n2 第二節');
  assert.equal(pages[2].body, '5 第五節');
});

test('remembers named layout groups and keeps different batches independent', () => {
  const state = { groups: {}, pageAssignments: {} };
  createLayoutGroup(state, 'scripture-layout', ['scripture:1', 'scripture:2'], { contentSize: 42 });
  createLayoutGroup(state, 'liturgy-layout', ['creed:1'], { contentSize: 48, contentAlign: 'left' });
  updateLayoutGroup(state, 'scripture-layout', { contentSize: 40, contentAlign: 'center' });
  assert.deepEqual(state.groups['scripture-layout'].params, { contentSize: 40, contentAlign: 'center' });
  assert.deepEqual(state.groups['liturgy-layout'].params, { contentSize: 48, contentAlign: 'left' });
  assert.equal(state.pageAssignments['scripture:2'], 'scripture-layout');
  assert.equal(state.pageAssignments['creed:1'], 'liturgy-layout');
});

test('reassigns checked pages to the latest group and supports detaching for single-page tuning', () => {
  const state = { groups: {}, pageAssignments: {} };
  createLayoutGroup(state, 'group-a', ['scripture:1', 'scripture:2'], { contentSize: 42 });
  createLayoutGroup(state, 'group-b', ['scripture:2'], { contentSize: 36 });
  assert.deepEqual(state.groups['group-a'].pageIds, ['scripture:1']);
  assert.equal(state.pageAssignments['scripture:2'], 'group-b');
  detachPagesFromLayoutGroup(state, ['scripture:2']);
  assert.equal(state.pageAssignments['scripture:2'], undefined);
  assert.deepEqual(state.groups['group-b'].pageIds, []);
});
