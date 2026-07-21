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

console.log('PrayerPPT slide production checks passed');
