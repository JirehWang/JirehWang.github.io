const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const source = fs.readFileSync(path.join(__dirname, 'bulletin-integration.js'), 'utf8');

test('reflows loaded and manually edited reports with the current effective layout', () => {
  const calls = source.match(/root\.reflowReportPagesForLayout\(\)/g) || [];
  assert.equal(calls.length, 2);
  assert.match(source, /applyReportsToModel\(model, reportsResult\.data\);[\s\S]{0,120}reflowReportPagesForLayout\(\)/);
  assert.match(source, /element\.oninput[\s\S]*applyReportsToModel\(model,[\s\S]*reflowReportPagesForLayout\(\);[\s\S]*preview\(\)/);
});

test('loads reports and praise from the existing bulletin API without a Firebase content mirror', () => {
  assert.doesNotMatch(source, /readServiceRecord/);
  assert.match(source, /loadCloudRecord\(endpoint, kind, date, root\.fetch\.bind\(root\)\)/);
});
