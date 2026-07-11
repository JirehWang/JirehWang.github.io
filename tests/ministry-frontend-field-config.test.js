const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const scriptPath = path.join(__dirname, '..', 'apps', 'LKC_MinistrySchedule', 'script.js');
const source = fs.readFileSync(scriptPath, 'utf8');
const configSlice = source.slice(
  source.indexOf('const initialFieldTemplates'),
  source.indexOf('//  🛡️ API 呼叫核心')
);

function createContext(storedConfig) {
  const storage = new Map();
  storage.set('ministry.pageFieldConfig.SVCH', JSON.stringify(storedConfig));
  const context = {
    currentId: 'SVCH',
    localStorage: {
      getItem: key => storage.get(key) || null,
      setItem: (key, value) => storage.set(key, String(value)),
      removeItem: key => storage.delete(key)
    }
  };
  vm.createContext(context);
  vm.runInContext(configSlice, context, { filename: scriptPath });
  return context;
}

test('backend field list setting wins over stale local field configuration', () => {
  const context = createContext({
    scheduleTarget: 'members',
    fields: [
      { name: '日期', enabled: true, useMemberList: false },
      { name: '小組', enabled: true, useMemberList: false },
      { name: '地點', enabled: true, useMemberList: false }
    ]
  });
  const data = {
    template: '新家人服事表模板',
    scheduleTarget: 'clusters',
    pageFieldConfig: {
      scheduleTarget: 'clusters',
      fields: [
        { name: '日期', enabled: true, useMemberList: false },
        { name: '小組', enabled: true, useMemberList: true },
        { name: '地點', enabled: true, useMemberList: false }
      ]
    }
  };

  const result = context.buildPageFieldConfig(data, ['日期', '小組', '地點']);

  assert.equal(result.scheduleTarget, 'clusters');
  assert.equal(result.fields.find(field => field.name === '小組').useMemberList, true);
});

test('new-family templates provide the shared list used by enabled custom fields', () => {
  const datalistSection = source.slice(
    source.indexOf('let datalistHTML = "";'),
    source.indexOf('const gridTemplate', source.indexOf('let datalistHTML = "";'))
  );
  const newFamilyStart = datalistSection.indexOf('if (currentTemplate === "新家人服事表模板")');
  const newFamilyBlock = datalistSection.slice(
    newFamilyStart,
    datalistSection.indexOf('\n    } else {', newFamilyStart)
  );
  assert.match(
    newFamilyBlock,
    /datalist id="customMembersList"/
  );
});
