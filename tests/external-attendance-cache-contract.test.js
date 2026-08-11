const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const externalRoot = 'D:\\program\\LKC';

function source(relativePath) {
  return fs.readFileSync(path.join(externalRoot, relativePath), 'utf8');
}

test('external attendance cache is explicitly privacy scoped', { skip: !fs.existsSync(externalRoot) }, () => {
  const sundayFirebase = source('主日出席_測試版/FirebaseSync.js');
  const sundayCore = source('主日出席_測試版/Core.js');
  const sundayAttendance = source('主日出席_測試版/AttendanceDB.js');
  const childrenFirebase = source('兒童出席_GAS/FirebaseSync.js');
  const childrenCore = source('兒童出席_GAS/Core.js');
  const childrenAttendance = source('兒童出席_GAS/AttendanceDB.js');
  const config = fs.readFileSync(path.join(__dirname, '..', 'config.js'), 'utf8');

  assert.match(sundayFirebase, /new Set\(\['getAttendanceSortIndex'\]\)/);
  assert.match(childrenFirebase, /new Set\(\['children_getAttendanceSortIndex'\]\)/);
  assert.match(sundayFirebase, /expiresAt: Date\.now\(\) \+ FIREBASE_ATTENDANCE_CACHE_TTL_MS/);
  assert.match(childrenFirebase, /expiresAt: Date\.now\(\) \+ FIREBASE_ATTENDANCE_CACHE_TTL_MS/);
  assert.match(sundayCore, /memberKey: String\(member\.id \|\| ''\),/);
  assert.match(childrenCore, /memberKey: String\(member\.id \|\| ''\), rank: index, attendanceCount:/);
  assert.doesNotMatch(sundayCore, /firebaseCacheWriteThrough\('getSmartAttendanceList'/);
  assert.doesNotMatch(childrenCore, /firebaseCacheWriteThrough\('children_getSmartAttendanceList'/);
  assert.match(config, /const _NINETY_DAYS = 90 \* 24 \* 60 \* 60;/);
  assert.match(config, /'getAttendanceSortIndex':\s+_NINETY_DAYS/);
  assert.match(config, /'children_getAttendanceSortIndex':\s+_NINETY_DAYS/);
  assert.match(config, /realAction !== 'children_getAttendanceSortIndex'/);
  assert.doesNotMatch(config, /'getSmartAttendanceList':\s+_SIX_HOURS/);
  assert.doesNotMatch(config, /'children_getSmartAttendanceList':\s+_SIX_HOURS/);

  // A Firebase hit supplies only the historical count map, so the normal
  // member source still provides display names without scanning 90-day
  // attendance rows. The existing same-day status read is intentionally kept.
  assert.match(sundayFirebase, /function firebaseReadAttendanceSortIndex\(topic, requestData\)/);
  assert.match(childrenFirebase, /function firebaseReadAttendanceSortIndex\(topic, requestData\)/);
  assert.match(sundayFirebase, /key !== 'memberKey' && key !== 'rank' && key !== 'attendanceCount'/);
  assert.match(childrenFirebase, /key !== 'memberKey' && key !== 'rank' && key !== 'attendanceCount'/);
  assert.match(sundayAttendance, /const cachedAttendanceMap =[^;]+firebaseReadAttendanceSortIndex\('getAttendanceSortIndex', attendanceCacheRequest\)/s);
  assert.match(childrenAttendance, /const cachedAttendanceMap =[^;]+firebaseReadAttendanceSortIndex\('children_getAttendanceSortIndex', attendanceCacheRequest\)/s);
  assert.match(sundayAttendance, /cachedAttendanceMap !== null\s*\? cachedAttendanceMap\s*:\s*getAttendanceCountMap\(ss, type\)/);
  assert.match(childrenAttendance, /cachedAttendanceMap !== null\s*\? cachedAttendanceMap\s*:\s*getAttendanceCountMap\(ss, type\)/);
  assert.match(sundayAttendance, /firebaseCacheWriteThrough\('getAttendanceSortIndex', attendanceCacheRequest, cacheResponse, attendanceCacheRevision\)/);
  assert.match(childrenAttendance, /firebaseCacheWriteThrough\('children_getAttendanceSortIndex', attendanceCacheRequest, cacheResponse, attendanceCacheRevision\)/);
  assert.doesNotMatch(sundayAttendance, /cacheResponse[\s\S]{0,500}name:/);
  assert.doesNotMatch(childrenAttendance, /cacheResponse[\s\S]{0,500}name:/);
});
