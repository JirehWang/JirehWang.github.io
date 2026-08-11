const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const gasPath = path.join(__dirname, '..', '..', '..', 'LKC', '主日出席_測試版', 'AttendanceTemp.js');
const source = fs.readFileSync(gasPath, 'utf8');

function loadContext() {
  const context = {
    Date,
    JSON,
    Math,
    Set,
    String,
    Number,
    Array,
    FIREBASE_DB_URL: 'https://example.invalid',
    UrlFetchApp: {},
    LockService: {},
    getSS() { throw new Error('not used'); }
  };
  vm.createContext(context);
  vm.runInContext(source, context, { filename: gasPath });
  return context;
}

test('GAS attendance planner carries source, lease, owner, and revision metadata', () => {
  const context = loadContext();
  const now = 1_700_000_000_000;
  const rows = [
    ['UID', '狀態', '類別', '時間', '操作者', '來源', 'revision', 'lockedUntil', 'ownerId'],
    ['LK100', 'checked', '主日', new Date(now), 'device-a', 'manual', 3, now + 600000, 'device-a']
  ];
  const plan = context._buildAttendanceTempMutationPlan(rows, {
    LK100: {
      uid: 'LK100',
      checked: true,
      source: 'qr',
      operatorId: 'device-b',
      ownerId: 'device-b',
      revision: 4,
      updatedAt: now + 1000,
      lockedUntil: now + 601000,
      expiresAt: now + 21600000,
      requestId: 'qr-4'
    }
  }, '主日', now);

  assert.equal(plan.upserts.length, 1);
  const values = plan.upserts[0].values;
  assert.equal(values[0], 'checked');
  assert.equal(values[1], '主日');
  assert.equal(new Date(values[2]).getTime(), now + 1000);
  assert.equal(values[3], 'device-b');
  assert.equal(values[4], 'qr');
  assert.equal(values[5], 4);
  assert.equal(values[6], now + 601000);
  assert.equal(values[7], 'device-b');
});

test('GAS attendance planner ignores a stale revision instead of overwriting SYNC_TEMP', () => {
  const context = loadContext();
  const now = 1_700_000_000_000;
  const rows = [
    ['UID', '狀態', '類別', '時間', '操作者', '來源', 'revision', 'lockedUntil', 'ownerId'],
    ['LK100', 'checked', '主日', new Date(now), 'device-b', 'qr', 8, now + 600000, 'device-b']
  ];
  const plan = context._buildAttendanceTempMutationPlan(rows, {
    LK100: {
      uid: 'LK100',
      checked: false,
      source: 'manual',
      operatorId: 'device-a',
      ownerId: '',
      revision: 7,
      updatedAt: now + 1000,
      lockedUntil: 0,
      expiresAt: now + 21600000,
      requestId: 'stale-7'
    }
  }, '主日', now);

  assert.equal(plan.upserts.length, 0);
  assert.equal(plan.appends.length, 0);
  assert.equal(plan.ackKeys.length, 1);
  assert.equal(plan.ackKeys[0], 'LK100');
});

test('GAS reads Firebase only after acquiring the scope lock', () => {
  const lockIndex = source.indexOf('lock.waitLock(30000)');
  const readIndex = source.indexOf('snapshot = _readAttendanceTempSnapshot(scope)');
  assert.ok(lockIndex >= 0);
  assert.ok(readIndex > lockIndex);
});
