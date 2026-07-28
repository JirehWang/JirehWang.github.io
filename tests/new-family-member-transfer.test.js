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

test('新家人加入主日會友名單時只轉拋姓名與性別', () => {
  const payloadMatch = source.match(
    /callSundayAttendancePayloadApi\('addMember',\s*\{([\s\S]*?)\}\)/
  );
  assert.ok(payloadMatch, '應呼叫 addMember API');
  const payloadSource = payloadMatch[1];

  assert.match(
    payloadSource,
    /^\s*name,\s*gender:\s*item\['性別'\]\s*\|\|\s*'',?\s*$/,
    'addMember payload 應將追蹤資料的「性別」映射到 gender'
  );
  assert.doesNotMatch(
    payloadSource,
    /gender:\s*item\['新家人性別'\]/,
    '不存在的「新家人性別」欄位不應用於 addMember payload'
  );
  assert.doesNotMatch(payloadSource, /\bnote\s*:/, '備註不得轉拋至會友名單');
  assert.doesNotMatch(payloadSource, /\bisExcluded\s*:/, '除姓名、性別外不得轉拋其他欄位');
});

test('新家人表單僅在比對到既有會友時標記已加入', () => {
  assert.match(
    source,
    /delete\s+payload\['會友狀態'\];\s*if\s*\(selectedMemberRecord\)/,
    '送出前應先移除表單中的會友狀態'
  );
  assert.match(
    source,
    /if\s*\(selectedMemberRecord\)\s*\{[\s\S]*?payload\['會友狀態'\]\s*=\s*'已加入';/,
    '比對到既有會友時應標記「已加入」'
  );
  assert.doesNotMatch(
    source,
    /payload\['會友狀態'\]\s*=\s*'未加入';/,
    '未比對到會友時應保持空白，不應標記「未加入」'
  );
});
