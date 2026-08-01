const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const sourcePath = path.join(
  __dirname,
  '..',
  'apps',
  'LKC_SundayserviceAttendance',
  'agm_attendance.js'
);
const source = fs.readFileSync(sourcePath, 'utf8');
const iifeStart = source.indexOf('\n(function() {');

function loadAgmHelpers() {
  assert.notEqual(iifeStart, -1, 'AGM page should keep pure helpers before the page IIFE');
  const context = { console };
  vm.createContext(context);
  vm.runInContext(source.slice(0, iifeStart), context, { filename: sourcePath });
  return context;
}

function makeMembers(count) {
  return Array.from({ length: count }, (_, index) => ({
    name: `Member ${index + 1}`,
    uid: `LK${String(index + 1).padStart(5, '0')}`,
    categoryCode: 'CAT_1'
  }));
}

function stateFor(members) {
  return Object.fromEntries(members.map(member => [member.uid, true]));
}

test('AGM quorum requires strictly more than half of 100 eligible members', () => {
  const { getAgmQuorumStats } = loadAgmHelpers();
  const members = makeMembers(100);
  const checked = stateFor(members.slice(0, 50));

  const stats = getAgmQuorumStats(members, checked, {});

  assert.equal(stats.effectiveTotal, 100);
  assert.equal(stats.threshold, 51);
  assert.equal(stats.presentCount, 50);
  assert.equal(stats.isQuorumMet, false);

  const quorumStats = getAgmQuorumStats(members, stateFor(members.slice(0, 51)), {});
  assert.equal(quorumStats.isQuorumMet, true);
});

test('AGM leave members are removed from the denominator and cannot count as present', () => {
  const { getAgmQuorumStats, buildAgmAttendancePayload } = loadAgmHelpers();
  const members = makeMembers(100);
  const leaveState = stateFor(members.slice(0, 10));
  const checked = stateFor(members.slice(0, 50));

  const stats = getAgmQuorumStats(members, checked, leaveState);

  assert.equal(stats.leaveCount, 10);
  assert.equal(stats.effectiveTotal, 90);
  assert.equal(stats.threshold, 46);
  assert.equal(stats.presentCount, 40);
  assert.equal(stats.isQuorumMet, false);

  const payload = buildAgmAttendancePayload(
    members,
    checked,
    '2026 AGM',
    'AGM:2026 AGM',
    leaveState
  );
  assert.equal(payload.cat1Total, 90);
  assert.equal(payload.cat1Present, 40);
  assert.equal(payload.cat1Leave, 10);
  assert.equal(
    JSON.stringify(payload.leaveUids),
    JSON.stringify(members.slice(0, 10).map(member => member.uid))
  );
});

test('AGM session QR uses a stable session scope instead of the editable title', () => {
  const { normalizeAgmSessionRecord, buildAgmSessionQrUrl } = loadAgmHelpers();
  const session = normalizeAgmSessionRecord({ sessionId: 'sabc123', sessionName: '2026 會員大會' });

  assert.equal(session.sessionId, 'SABC123');
  assert.equal(session.meetingTitle, '2026 會員大會');
  assert.equal(session.scope, 'AGM:SABC123');
  assert.equal(
    buildAgmSessionQrUrl('sabc123', 'https://example.test/apps/?old=1'),
    'https://example.test/apps/?agmSession=SABC123&agmRole=scanner'
  );
});

test('AGM scanner QR entry is a locked attendance-only mode', () => {
  const { isAgmScannerQrEntry } = loadAgmHelpers();

  assert.equal(isAgmScannerQrEntry({ sessionId: 'SABC123', role: 'scanner' }), true);
  assert.equal(isAgmScannerQrEntry({ sessionId: 'SABC123', role: 'viewer' }), false);
  assert.equal(isAgmScannerQrEntry({ sessionId: '', role: 'scanner' }), false);
});

test('inactive official members are excluded from active attendance data', () => {
  const { isOfficialMemberActive } = loadAgmHelpers();

  assert.equal(isOfficialMemberActive({ name: 'Active', isActive: true }), true);
  assert.equal(isOfficialMemberActive({ name: 'Legacy blank status' }), true);
  assert.equal(isOfficialMemberActive({ name: 'Deceased', isActive: false }), false);
  assert.equal(isOfficialMemberActive({ name: 'Disabled', isActive: '停用' }), false);
});
