const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const scriptPath = path.join(__dirname, '..', 'apps', 'LKC_MemberStatus', 'script.js');
const source = fs.readFileSync(scriptPath, 'utf8');

function loadMemberStatusScript(churchAPI) {
  const overlay = { classList: { add() {}, remove() {} } };
  const loadingText = { textContent: '' };
  const context = {
    console,
    document: {
      getElementById(id) {
        if (id === 'loadingOverlay') return overlay;
        if (id === 'loadingText') return loadingText;
        throw new Error(`Unexpected element lookup: ${id}`);
      }
    },
    window: {
      addEventListener() {},
      ensureAPIReady: async () => {},
      churchAPI
    }
  };
  vm.createContext(context);
  vm.runInContext(source, context, { filename: scriptPath });
  return context;
}

test('initial member load does not request the first profile before the user clicks', async () => {
  const calls = [];
  const context = loadMemberStatusScript(async action => {
    calls.push(action);
    return {
      success: true,
      members: [{ uid: 'LK00001', name: '王小明' }],
      filters: { groups: [], ministries: [] },
      unresolvedParticipants: []
    };
  });

  vm.runInContext(`
    selectedUids = [];
    populateFilters = () => {};
    renderMetrics = () => {};
    renderMemberList = () => {};
    selectMember = async uid => { selectedUids.push(uid); };
  `, context);

  await vm.runInContext('loadMembers()', context);
  assert.deepEqual(calls, ['getMembers']);
  assert.deepEqual(Array.from(context.selectedUids), []);
});

test('participation rows stay compact until expanded, then include every matching member', () => {
  const context = loadMemberStatusScript(async () => ({ success: true }));
  const rows = Array.from({ length: 30 }, (_, index) => ({
    uid: `LK${String(index + 1).padStart(5, '0')}`,
    name: `Member ${index + 1}`,
    participation: { ministryCount: index }
  }));
  context.testRows = rows;

  const compact = vm.runInContext('getVisibleParticipationRows(testRows, false)', context);
  const expanded = vm.runInContext('getVisibleParticipationRows(testRows, true)', context);

  assert.equal(compact.length, 24);
  assert.equal(expanded.length, 30);
  assert.equal(expanded[0].participation.ministryCount, 29);
});
