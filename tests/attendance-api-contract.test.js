const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repoRoot = path.join(__dirname, '..');

for (const app of ['LKC_SundayserviceAttendance', 'LKC_ChildrenAttendance']) {
  test(`${app} API retries only read requests and checks HTTP responses`, () => {
    const source = fs.readFileSync(path.join(repoRoot, 'apps', app, 'api.js'), 'utf8');
    assert.match(source, /READ_ONLY_ACTIONS/);
    assert.match(source, /getGroupConfig/);
    assert.match(source, /getSmartAttendanceList/);
    assert.match(source, /response\.ok/);
    assert.match(source, /AbortController/);
    assert.match(source, /attempts/);
    assert.match(source, /withFailureHandler/);
  });
}

test('Sunday attendance shows a retry action when group config cannot be loaded', () => {
  const source = fs.readFileSync(
    path.join(repoRoot, 'apps', 'LKC_SundayserviceAttendance', 'attendance.js'),
    'utf8'
  );
  const block = source.slice(
    source.indexOf('function loadGroupConfig'),
    source.indexOf('function renderCategorySelect')
  );
  assert.match(block, /withFailureHandler/);
  assert.match(block, /重試|重新載入|retry/i);
});
