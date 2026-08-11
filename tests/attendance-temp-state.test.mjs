import test from 'node:test';
import assert from 'node:assert/strict';
import {
  ATTENDANCE_PENDING_LOCK_MS,
  ATTENDANCE_TEMP_TTL_MS,
  buildAttendanceTempEntry,
  isAttendanceTempLockActive
} from '../firebase/attendance-temp-state.mjs';

const NOW = 1_700_000_000_000;

test('pending manual check claims a ten-minute lease and records its source', () => {
  const result = buildAttendanceTempEntry({
    current: null,
    uid: 'LK100',
    operatorId: 'device-a',
    source: 'manual',
    checked: true,
    requestId: 'manual-1',
    now: NOW
  });

  assert.equal(result.accepted, true);
  assert.equal(result.entry.checked, true);
  assert.equal(result.entry.source, 'manual');
  assert.equal(result.entry.ownerId, 'device-a');
  assert.equal(result.entry.revision, 1);
  assert.equal(result.entry.lockedUntil, NOW + ATTENDANCE_PENDING_LOCK_MS);
  assert.equal(result.entry.expiresAt, NOW + ATTENDANCE_TEMP_TTL_MS);
  assert.equal(isAttendanceTempLockActive(result.entry, NOW + 1), true);
});

test('duplicate QR check during an active lease is idempotent and cannot steal ownership', () => {
  const current = {
    uid: 'LK100',
    checked: true,
    source: 'manual',
    ownerId: 'device-a',
    operatorId: 'device-a',
    revision: 7,
    lockedUntil: NOW + 5 * 60 * 1000,
    expiresAt: NOW + ATTENDANCE_TEMP_TTL_MS
  };

  const result = buildAttendanceTempEntry({
    current,
    uid: 'LK100',
    operatorId: 'device-b',
    source: 'qr',
    checked: true,
    requestId: 'qr-1',
    now: NOW
  });

  assert.equal(result.accepted, false);
  assert.equal(result.reason, 'locked-by-other');
  assert.equal(result.entry.revision, 7);
  assert.equal(result.entry.ownerId, 'device-a');
  assert.equal(result.entry.source, 'manual');
});

test('owner can cancel pending attendance immediately without confirmation', () => {
  const current = {
    uid: 'LK100',
    checked: true,
    source: 'qr',
    ownerId: 'device-a',
    operatorId: 'device-a',
    revision: 7,
    lockedUntil: NOW + 5 * 60 * 1000,
    expiresAt: NOW + ATTENDANCE_TEMP_TTL_MS
  };

  const result = buildAttendanceTempEntry({
    current,
    uid: 'LK100',
    operatorId: 'device-a',
    source: 'qr',
    checked: false,
    requestId: 'qr-cancel-1',
    now: NOW
  });

  assert.equal(result.accepted, true);
  assert.equal(result.entry.checked, false);
  assert.equal(result.entry.ownerId, '');
  assert.equal(result.entry.lastActionBy, 'device-a');
  assert.equal(result.entry.revision, 8);
  assert.equal(result.entry.expiresAt, NOW + ATTENDANCE_TEMP_TTL_MS);
});

test('another device cannot cancel an active pending lease', () => {
  const current = {
    uid: 'LK100',
    checked: true,
    source: 'manual',
    ownerId: 'device-a',
    operatorId: 'device-a',
    revision: 7,
    lockedUntil: NOW + 5 * 60 * 1000,
    expiresAt: NOW + ATTENDANCE_TEMP_TTL_MS
  };

  const result = buildAttendanceTempEntry({
    current,
    uid: 'LK100',
    operatorId: 'device-b',
    source: 'manual',
    checked: false,
    requestId: 'manual-cancel-1',
    now: NOW
  });

  assert.equal(result.accepted, false);
  assert.equal(result.reason, 'locked-by-other');
  assert.equal(result.entry.checked, true);
  assert.equal(result.entry.revision, 7);
});

test('expired lease can be claimed by a different source', () => {
  const current = {
    uid: 'LK100',
    checked: true,
    source: 'manual',
    ownerId: 'device-a',
    operatorId: 'device-a',
    revision: 7,
    lockedUntil: NOW - 1,
    expiresAt: NOW + ATTENDANCE_TEMP_TTL_MS
  };

  const result = buildAttendanceTempEntry({
    current,
    uid: 'LK100',
    operatorId: 'device-b',
    source: 'qr',
    checked: true,
    requestId: 'qr-2',
    now: NOW
  });

  assert.equal(result.accepted, true);
  assert.equal(result.entry.revision, 8);
  assert.equal(result.entry.ownerId, 'device-b');
  assert.equal(result.entry.source, 'qr');
  assert.equal(result.entry.lockedUntil, NOW + ATTENDANCE_PENDING_LOCK_MS);
});
