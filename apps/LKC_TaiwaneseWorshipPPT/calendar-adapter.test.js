const test = require('node:test');
const assert = require('node:assert/strict');
const {
  applyCalendarEvent,
  selectTaiwaneseSermonEvent,
  getCalendarValue
} = require('./calendar-adapter.js');

test('maps calendar values into Taiwanese worship fields and keeps hymn numbers only', () => {
  const model = {
    call: {}, sermon: {}, scripture: {}, verse: {}, response: {}, 'hymn-1': {}, 'hymn-2': {}, doxology: {}
  };
  applyCalendarEvent({
    values: [
      { fieldName: '宣召', value: '詩篇 136:1' },
      { fieldName: '講道題目', value: '建造百倍成長的生命' },
      { fieldName: '講員', value: '陳志聰牧師' },
      { fieldName: '經文', value: '馬太福音 13:1-9' },
      { fieldName: '金句', value: '馬太福音 13:23' },
      { fieldName: '啟應文', value: '第 3 篇' },
      { fieldName: '聖詩第一首', value: '65' },
      { fieldName: '聖詩第二首', value: '474' },
      { fieldName: '頌榮', value: '510' }
    ]
  }, model);
  assert.equal(model.call.sourceValue, '詩篇 136:1');
  assert.equal(model.sermon.title, '建造百倍成長的生命');
  assert.equal(model.sermon.kicker, '陳志聰牧師');
  assert.equal(model.scripture.sourceValue, '馬太福音 13:1-9');
  assert.equal(model.verse.sourceValue, '馬太福音 13:23');
  assert.equal(model.response.sourceValue, '第 3 篇');
  assert.equal(model['hymn-1'].sourceValue, '65');
  assert.equal(model['hymn-2'].sourceValue, '474');
  assert.equal(model.doxology.sourceValue, '510');
});

test('selects only the same-date 講道資訊 - 台語 event', () => {
  const selected = selectTaiwaneseSermonEvent([
    { date: '2026-07-12', typeName: '聚會名稱', typeFullName: '聚會名稱' },
    { date: '2026-07-12', typeName: '華語', typeFullName: '講道資訊 - 華語' },
    { date: '2026-07-12', typeName: '台語', typeFullName: '講道資訊 - 台語', eventId: 'target' },
    { date: '2026-07-19', typeName: '台語', typeFullName: '講道資訊 - 台語', eventId: 'wrong-date' }
  ], '2026-07-12');
  assert.equal(selected.eventId, 'target');
});

test('keeps an alphanumeric hymn suffix used by Drive filenames', () => {
  const model = {
    call: {}, sermon: {}, scripture: {}, verse: {}, response: {}, 'hymn-1': {}, 'hymn-2': {}, doxology: {}
  };
  applyCalendarEvent({ values: [{ fieldName: '聖詩一', value: '第306B首' }] }, model);
  assert.equal(model['hymn-1'].sourceValue, '306B');
});

test('maps production fields as generator inputs instead of finished slide content', () => {
  const model = {
    call: {}, sermon: {}, scripture: {}, verse: {}, response: {}, 'hymn-1': {}, 'hymn-2': {}, doxology: {}
  };
  const event = {
    typeName: '台語',
    typeFullName: '講道資訊 - 台語',
    values: [
      { fieldName: '講題', value: '建造百倍成長的生命' },
      { fieldName: '講員', value: '陳志聰' },
      { fieldName: '經文', value: '馬太福音13:1-9' }
    ]
  };
  applyCalendarEvent(event, model);
  assert.equal(model.sermon.title, '建造百倍成長的生命');
  assert.equal(model.sermon.kicker, '陳志聰');
  assert.equal(model.scripture.sourceValue, '馬太福音13:1-9');
  assert.equal(model.scripture.body, undefined);
  assert.equal(getCalendarValue(event, 'scripture'), '馬太福音13:1-9');
});
