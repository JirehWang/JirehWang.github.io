const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repoRoot = path.join(__dirname, '..');
const gasRoot = path.join(repoRoot, '..', '..', 'LKC', '主日出席_測試版');

test('Firebase attendance temp exposes deterministic idempotent keys and realtime writes', () => {
  const source = fs.readFileSync(path.join(repoRoot, 'firebase', 'attendance-temp.js'), 'utf8');
  const stateSource = fs.readFileSync(path.join(repoRoot, 'firebase', 'attendance-temp-state.mjs'), 'utf8');
  assert.match(source, /ATTENDANCE_TEMP_ROOT\s*=\s*['"]attendanceTemp['"]/);
  assert.match(stateSource, /ATTENDANCE_TEMP_TTL_MS\s*=\s*6\s*\*\s*60\s*\*\s*60\s*\*\s*1000/);
  assert.match(stateSource, /ATTENDANCE_PENDING_LOCK_MS\s*=\s*10\s*\*\s*60\s*\*\s*1000/);
  assert.match(source, /export function attendanceTempKey/);
  assert.match(source, /export async function writeAttendanceTemp/);
  assert.match(source, /runTransaction/);
  assert.match(source, /source/);
  assert.match(source, /lockedUntil/);
  assert.match(source, /ownerId/);
  assert.match(source, /onValue/);
  assert.match(source, /updatedAt/);
  assert.match(source, /requestId/);
});

test('QR scanner keeps unacknowledged writes in a bounded retry queue and batches backend flushes', () => {
  const source = fs.readFileSync(
    path.join(repoRoot, 'apps', 'qrcodescanner.github.io', 'index.html'),
    'utf8'
  );
  assert.match(source, /ATTENDANCE_QR_QUEUE_KEY/);
  assert.match(source, /flushPendingScans/);
  assert.match(source, /MAX_PENDING_SCANS/);
  assert.match(source, /writeAttendanceTemp/);
  assert.match(source, /source: 'qr'/);
  assert.match(source, /mode/);
  assert.match(source, /setInterval\(flushPendingScans/);
  assert.doesNotMatch(source, /fetch\(targetUrl,\s*\{\s*mode:\s*'no-cors'/);
});

test('main attendance page writes temp state to Firebase and schedules a batch flush', () => {
  const source = fs.readFileSync(
    path.join(repoRoot, 'apps', 'LKC_SundayserviceAttendance', 'attendance.js'),
    'utf8'
  );
  assert.match(source, /writeAttendanceTemp/);
  assert.match(source, /flushAttendanceTemp/);
  assert.match(source, /ATTENDANCE_TEMP_FLUSH_INTERVAL_MS/);
  assert.match(source, /ATTENDANCE_TEMP_FLUSH_INTERVAL_MS\s*=\s*30000/);
  assert.match(source, /subscribeAttendanceTemp/);
  assert.match(source, /source: 'manual'/);
  assert.match(source, /updateAttendanceTempQueueItem/);
  assert.match(source, /remoteStatusSequence/);
  assert.match(source, /requestSequence !== remoteStatusSequence/);
  assert.doesNotMatch(source, /flushAttendanceTempToBackend\(currentAttType\)/);
  assert.match(source, /mode=/);
});

test('GAS status polling does not overwrite Firebase-owned pending attendance', () => {
  const source = fs.readFileSync(
    path.join(repoRoot, 'apps', 'LKC_SundayserviceAttendance', 'attendance.js'),
    'utf8'
  );
  assert.match(source, /realtimeAttendanceTempEntries/);
  assert.match(source, /realtimeAttendanceTempEntries\s*=\s*entries/);
  assert.match(source, /else if\s*\(realtimeAttendanceTempReady\s*\|\|/);
  assert.match(source, /hasOwnProperty\.call\(realtimeAttendanceTempEntries,\s*memKey\)/);
});

test('attendance cards expose manual and QR pending source styles', () => {
  const source = fs.readFileSync(
    path.join(repoRoot, 'apps', 'LKC_SundayserviceAttendance', 'attendance.html'),
    'utf8'
  );
  assert.match(source, /\.att-item\.pending-manual/);
  assert.match(source, /\.att-item\.pending-qr/);
  assert.match(source, /#2563eb/);
  assert.match(source, /content:\s*["']QR["']/);
});

test('GAS exposes a 30-second Firebase temp flush and idempotent mutation planner', () => {
  const source = fs.readFileSync(path.join(gasRoot, 'AttendanceTemp.js'), 'utf8');
  assert.match(source, /ATTENDANCE_TEMP_FLUSH_INTERVAL_MS\s*=\s*30000/);
  assert.match(source, /ATTENDANCE_TEMP_TTL_MS\s*=\s*6\s*\*\s*60\s*\*\s*60\s*\*\s*1000/);
  assert.match(source, /ATTENDANCE_PENDING_LOCK_MS\s*=\s*10\s*\*\s*60\s*\*\s*1000/);
  assert.match(source, /function flushAttendanceTemp\s*\(/);
  assert.match(source, /function _buildAttendanceTempMutationPlan\s*\(/);
  assert.match(source, /LockService/);
  assert.match(source, /requestId/);
  assert.match(source, /X-Firebase-ETag/);
  assert.match(source, /If-Match/);
  assert.match(source, /status === 412/);
  assert.match(source, /source/);
  assert.match(source, /lockedUntil/);
  assert.match(source, /revision/);
  assert.match(source, /_ensureAttendanceTempHeader/);
  assert.match(source, /lock\.waitLock[\s\S]*_readAttendanceTempSnapshot/);
  assert.match(fs.readFileSync(path.join(gasRoot, 'Core.js'), 'utf8'), /flushAttendanceTemp/);
});
