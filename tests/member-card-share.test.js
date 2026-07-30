const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const projectRoot = path.resolve(__dirname, '..');
const memberPagePath = path.join(
  projectRoot,
  'apps',
  'LKC_SundayserviceAttendance',
  'members.html'
);
const cardPagePath = path.join(
  projectRoot,
  'apps',
  'LKC_SundayserviceAttendance',
  'card.html'
);
const backendRoot = path.join('D:/program', 'LKC', '主日出席_測試版');
const memberDbPath = path.join(backendRoot, 'MemberDB.js');
const corePath = path.join(backendRoot, 'Core.js');

function loadMemberDb() {
  const properties = new Map();
  const scriptProperties = {
    getProperty(key) {
      return properties.has(key) ? properties.get(key) : null;
    },
    setProperty(key, value) {
      properties.set(key, String(value));
    },
  };
  const context = vm.createContext({
    console,
    PropertiesService: { getScriptProperties: () => scriptProperties },
    Utilities: {
      getUuid: () => '12345678-1234-4234-8234-123456789abc',
    },
    getCachedMembers: () => [
      ['王小明', '男', '', '', false, '', '', 'LK00001', '', '小羊'],
    ],
  });

  vm.runInContext(fs.readFileSync(memberDbPath, 'utf8'), context, {
    filename: memberDbPath,
  });
  return context;
}

test('member card share contract is wired through the UI, GAS route, and public page', () => {
  const memberPage = fs.readFileSync(memberPagePath, 'utf8');
  const coreSource = fs.readFileSync(corePath, 'utf8');

  assert.match(memberPage, /id="typeShare"/);
  assert.match(memberPage, /getMemberCardShareLink/);
  assert.match(coreSource, /case 'getMemberCardShareLink'/);
  assert.match(coreSource, /case 'getMemberCardByShareToken'/);
  assert.equal(fs.existsSync(cardPagePath), true);

  const cardPage = fs.readFileSync(cardPagePath, 'utf8');
  assert.match(cardPage, /getMemberCardByShareToken/);
  assert.match(cardPage, /下載卡片圖檔/);
  assert.match(cardPage, /URL\.createObjectURL/);
  assert.match(cardPage, /window\.open/);
  assert.match(cardPage, /長按圖片/);
});

test('share tokens are stable per member and resolve to the current member card', () => {
  const context = loadMemberDb();
  const first = context.getMemberCardShareLink({ uid: 'LK00001' });
  const second = context.getMemberCardShareLink({ uid: 'LK00001' });

  assert.equal(first.success, true);
  assert.equal(first.shareToken, second.shareToken);
  assert.match(first.shareUrl, /apps\/LKC_SundayserviceAttendance\/card\.html\?share=/);
  assert.doesNotMatch(first.shareUrl, /LK00001/);

  context.previewMemberCard = ({ name, uid }) => ({
    success: true,
    base64: 'data:image/jpeg;base64,card-data',
    name,
    uid,
  });
  const resolved = context.getMemberCardByShareToken({ token: first.shareToken });

  assert.equal(resolved.success, true);
  assert.equal(resolved.base64, 'data:image/jpeg;base64,card-data');
  assert.equal(resolved.name, '王小明');
});

test('unknown or malformed share tokens never expose a card', () => {
  const context = loadMemberDb();

  assert.equal(context.getMemberCardByShareToken({ token: '' }).success, false);
  assert.equal(context.getMemberCardByShareToken({ token: 'not-issued' }).success, false);
});
