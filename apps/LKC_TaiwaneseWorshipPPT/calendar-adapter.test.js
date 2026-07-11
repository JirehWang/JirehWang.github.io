const test = require('node:test');
const assert = require('node:assert/strict');
const { applyCalendarEvent } = require('./calendar-adapter.js');

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
  assert.equal(model.call.body, '詩篇 136:1');
  assert.equal(model.sermon.title, '建造百倍成長的生命');
  assert.equal(model.sermon.kicker, '陳志聰牧師');
  assert.equal(model.scripture.body, '馬太福音 13:1-9');
  assert.equal(model.verse.body, '馬太福音 13:23');
  assert.equal(model.response.body, '第 3 篇');
  assert.equal(model['hymn-1'].title, '65');
  assert.equal(model['hymn-2'].title, '474');
  assert.equal(model.doxology.title, '510');
});
