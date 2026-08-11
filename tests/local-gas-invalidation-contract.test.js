const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');

const gasFiles = [
  'D:/program/LKC/主日出席_測試版/FirebaseSync.js',
  'D:/program/LKC/兒童出席_GAS/FirebaseSync.js',
  'D:/program/LKC/敬拜團/FirebaseSync.js'
];

for (const file of gasFiles) {
  test(`GAS batch invalidation contract: ${file}`, { skip: !fs.existsSync(file) }, () => {
    const source = fs.readFileSync(file, 'utf8');
    assert.match(source, /function firebaseInvalidate\(topics\)/);
    assert.match(source, /method:\s*'patch'/);
    assert.match(
      source,
      /uniqueTopics\.forEach\(topic => (?:firebaseCacheDeleteAll\(topic\)|\{\s*firebaseCacheDeleteAll\(topic\);)/
    );
    assert.match(source, /mode:\s*'batch'/);
    assert.match(source, /mode:\s*'fallback'/);
  });
}

test('one-command GAS rollback assets exist', {
  skip: !fs.existsSync('restore-current-state.ps1')
}, () => {
  assert.equal(fs.existsSync('restore-current-state.ps1'), true);
  for (const name of ['main-FirebaseSync.js', 'children-FirebaseSync.js', 'worship-FirebaseSync.js', 'MemberStatusCore.js', 'main-ARCHITECTURE.md']) {
    assert.equal(fs.existsSync(`.rollback/gas/${name}`), true, name);
  }
});
