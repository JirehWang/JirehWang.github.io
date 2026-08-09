const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const attendanceDir = path.join(__dirname, '..', 'apps', 'LKC_SundayserviceAttendance');
const attendanceHtml = fs.readFileSync(path.join(attendanceDir, 'attendance.html'), 'utf8');
const attendanceJs = fs.readFileSync(path.join(attendanceDir, 'attendance.js'), 'utf8');

test('Sunday attendance add-member modal handlers are defined for inline HTML actions', () => {
  assert.match(attendanceHtml, /onclick="openAttendanceAddModal\(\)"/);
  assert.match(attendanceHtml, /onclick="closeAttendanceAddModal\(\)"/);
  assert.match(attendanceJs, /function openAttendanceAddModal\s*\(/);
  assert.match(attendanceJs, /function closeAttendanceAddModal\s*\(/);
});
