const assert = require('node:assert/strict');
const production = require('./slide-production.js');

assert.equal(production.DEFAULT_LAYOUT_PARAMS.titleColor, '#111827',
  'the default title must remain readable on the white PrayerPPT background');
assert.equal(production.DEFAULT_LAYOUT_PARAMS.contentColor, '#1F2937',
  'the default content must remain readable on the white PrayerPPT background');

const model = {
  silence: { title: '請安靜心、等候神', label: '請安靜心、等候神' },
  thanksgiving: { title: '獻上感謝讚美', label: '獻上感謝讚美' },
  repentance: { title: '悔改認罪', label: '悔改認罪' },
  world: { title: '為世界 pray', label: '為世界 pray' },
  oneself: { title: '為自己 pray', label: '為自己 pray' }
};

const sections = production.parseRecognizedSections([
  '4. 獻上感謝讚美\na. 第一項\n\n6. 為世界 pray\na. 世界禱告\n\n2',
  '2026年7月19日 9:00～9:50\n林口長老教會\n禱告會\n\n1.請安靜心，等候神\n請放下身心重擔。',
  '10. 為自己 pray\na. 身心靈剛強\n\n5.'
], model);

assert.deepEqual(sections.world.lines, ['a. 世界禱告'],
  'a later image preamble must not leak into the previous image section');
assert.doesNotMatch(sections.world.lines.join('\n'), /2026年|林口長老教會|禱告會/,
  'date and church headings from a later image must be ignored');
assert.deepEqual(sections.silence.lines, ['請放下身心重擔。']);
assert.deepEqual(sections.oneself.lines, ['a. 身心靈剛強']);
assert.equal(sections.repentance, undefined,
  'a trailing handwritten page number such as "5." must not create a section');

const compactListPages = production.generateSectionPages('world', {
  type: 'list',
  title: '為世界 pray',
  body: 'a. 第一個小點\nb. 第二個小點\nc. 第三個小點'
});
assert.equal(compactListPages.length, 1,
  'small sub-points under the same major section must share one slide when they fit');
assert.equal(compactListPages[0].body, 'a. 第一個小點\nb. 第二個小點\nc. 第三個小點');

const overflowListPages = production.generateSectionPages('members', {
  type: 'list',
  title: '為教會肢體 pray',
  body: 'a. 第一點\nb. 第二點\nc. 第三點\nd. 第四點\ne. 第五點\nf. 第六點'
});
assert.equal(overflowListPages.length, 2,
  'a major section must continue onto another slide only after the available lines are filled');
assert.equal(overflowListPages[0].body.split('\n').length, 5);
assert.equal(overflowListPages[1].body, 'f. 第六點');

const keepPointTogetherPages = production.generateSectionPages('church', {
  type: 'list',
  title: '為教會 pray',
  body: 'a. 第一個小點\n第一點續行一\n第一點續行二\n第一點續行三\nb. 第二個小點\n第二點續行'
});
assert.equal(keepPointTogetherPages.length, 2);
assert.doesNotMatch(keepPointTogetherPages[0].body, /^b\./m,
  'a sub-point that fits on the next slide must not be split across a page boundary');
assert.match(keepPointTogetherPages[1].body, /^b\. 第二個小點/,
  'the complete sub-point must move to the next slide');

const layoutState = { groups: {}, pageAssignments: {} };
production.createLayoutGroup(layoutState, 'prayer-body', ['world:1'], { contentSize: 36 });
assert.equal(production.layoutForPage(layoutState, { id: 'world:1' }).contentSize, 36,
  'saved PrayerPPT layout groups must affect the resolved slide layout');
production.detachPagesFromLayoutGroup(layoutState, ['world:1']);
assert.equal(layoutState.pageAssignments['world:1'], undefined,
  'PrayerPPT layout groups must allow a page to be detached again');

assert.deepEqual(
  production.extractBibleReferences([
    '羅 8:26「閣聖神也親像按呢扶持咱的軟弱」',
    '弗 6:18「用逐樣的祈禱及懇求」',
    '提摩太前書 2:1「所以我所勸勉的」',
    '路 4:38「耶穌對會堂起來」'
  ]),
  ['羅 8:26', '弗 6:18', '提摩太前書 2:1', '路 4:38'],
  'Bible text must be reduced to references before FHL Bible API lookup'
);

console.log('PrayerPPT slide production checks passed');
