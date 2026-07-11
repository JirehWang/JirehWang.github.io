const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function loadNormalizer() {
  const sourcePath = 'D:/program/LKC/主日出席_測試版/MinistryCore.js';
  const source = fs.readFileSync(sourcePath, 'utf8') +
    '\nthis.__test = { normalize: _ministryNormalizePageFieldConfig };';
  const context = { console, JSON, Date, Set, Array, Object };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context.__test.normalize;
}

test('page field config retains the cluster target and member-list flags', () => {
  const normalize = loadNormalizer();
  const config = normalize({
    fieldTemplateType: '事工型模板',
    scheduleMode: 'schedule',
    scheduleTarget: 'clusters',
    fields: [
      { name: '日期', enabled: true, custom: false, useMemberList: false },
      { name: '帶領小組群', enabled: true, custom: true, useMemberList: true }
    ],
    requiredFields: ['日期']
  }, 'M01', '事工型模板');

  assert.equal(config.scheduleTarget, 'clusters');
  assert.equal(config.fields.find(field => field.name === '帶領小組群').useMemberList, true);
});
