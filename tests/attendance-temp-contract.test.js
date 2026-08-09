const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repoRoot = path.join(__dirname, '..');
const gasRoot = path.join(repoRoot, '..', '..', 'LKC', '主日出席_測試版');

test('Firebase attendance temp exposes deterministic idempotent keys and realtime writes', () => {
  const source = fs.readFileSync(path.join(repoRoot, 'firebase', 'attendance-temp.js'), 'utf8');
  assert.match(source, /ATTENDANCE_TEMP_ROOT\s*=\s*['"]attendanceTemp['"]/);
  assert.match(source, /export function attendanceTempKey/);
  assert.match(source, /export async function writeAttendanceTemp/);
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
  assert.match(source, /mode=/);
});

test('GAS exposes a 5-second Firebase temp flush and idempotent mutation planner', () => {
  const source = fs.readFileSync(path.join(gasRoot, 'AttendanceTemp.js'), 'utf8');
  assert.match(source, /ATTENDANCE_TEMP_FLUSH_INTERVAL_MS\s*=\s*5000/);
  assert.match(source, /function flushAttendanceTemp\s*\(/);
  assert.match(source, /function _buildAttendanceTempMutationPlan\s*\(/);
  assert.match(source, /LockService/);
  assert.match(source, /requestId/);
  assert.match(source, /X-Firebase-ETag/);
  assert.match(source, /If-Match/);
  assert.match(source, /status === 412/);
  assert.match(fs.readFileSync(path.join(gasRoot, 'Core.js'), 'utf8'), /flushAttendanceTemp/);
});
