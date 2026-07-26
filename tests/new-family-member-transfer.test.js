const fs = require('fs');
const path = require('path');
const test = require('node:test');
const assert = require('node:assert/strict');

const scriptPath = path.join(
  __dirname,
  '..',
  'apps',
  'LKC_NewFamily',
  'script.js'
);
const source = fs.readFileSync(scriptPath, 'utf8');

test('新家人加入主日會友名單時使用追蹤資料的性別欄位', () => {
  assert.match(
    source,
    /callSundayAttendancePayloadApi\('addMember',\s*\{\s*name,\s*gender:\s*item\['性別'\]\s*\|\|\s*'',/,
    'addMember payload 應將追蹤資料的「性別」映射到 gender'
  );
  assert.doesNotMatch(
    source,
    /gender:\s*item\['新家人性別'\]/,
    '不存在的「新家人性別」欄位不應用於 addMember payload'
  );
});
