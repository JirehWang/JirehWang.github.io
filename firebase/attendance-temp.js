// Realtime attendance staging channel.
// The child key (scope + UID) is the idempotency key. Retrying the same event
// therefore replaces the same record instead of appending another attendance.
import { app, rtdb } from './firebase-config.js';
import {
  ATTENDANCE_PENDING_LOCK_MS,
  ATTENDANCE_TEMP_SCHEMA_VERSION,
  ATTENDANCE_TEMP_TTL_MS,
  buildAttendanceTempEntry,
  normalizeAttendanceTempEntry
} from './attendance-temp-state.mjs';
import { getAuth, signInAnonymously } from 'https://www.gstatic.com/firebasejs/10.13.0/firebase-auth.js';
import {
  get,
  onValue,
  ref,
  runTransaction
} from 'https://www.gstatic.com/firebasejs/10.13.0/firebase-database.js';

export const ATTENDANCE_TEMP_ROOT = 'attendanceTemp';
export {
  ATTENDANCE_PENDING_LOCK_MS,
  ATTENDANCE_TEMP_SCHEMA_VERSION,
  ATTENDANCE_TEMP_TTL_MS
};
const auth = getAuth(app);
let authReady = null;

function ensureAttendanceAuth() {
  if (auth.currentUser) return Promise.resolve(auth.currentUser);
  if (!authReady) {
    authReady = signInAnonymously(auth).then(result => result.user).catch(error => {
      authReady = null;
      throw error;
    });
  }
  return authReady;
}

function encodeKey(value) {
  const normalized = String(value == null ? '' : value).trim();
  if (!normalized) throw new Error('attendance temp key is required');
  return encodeURIComponent(normalized);
}

export function attendanceTempKey(scope, uid) {
  return `${encodeKey(scope)}/${encodeKey(String(uid).toUpperCase())}`;
}

export function attendanceTempScopePath(scope) {
  return `${ATTENDANCE_TEMP_ROOT}/${encodeKey(scope)}`;
}

export function attendanceTempPath(scope, uid) {
  return `${attendanceTempScopePath(scope)}/${encodeKey(String(uid).toUpperCase())}`;
}

export function makeAttendanceRequestId(scope, uid, operatorId, now = Date.now()) {
  const device = String(operatorId || 'anonymous').trim() || 'anonymous';
  return `${encodeKey(scope)}:${encodeKey(String(uid).toUpperCase())}:${encodeKey(device)}:${now}`;
}

function normalizeEntry(entry, key) {
  if (!entry || typeof entry !== 'object') return null;
  const fallbackUid = (() => {
    try { return decodeURIComponent(key || ''); } catch (error) { return key || ''; }
  })();
  const normalized = normalizeAttendanceTempEntry({ ...entry, uid: entry.uid || fallbackUid });
  if (!normalized.uid) return null;
  return {
    ...normalized,
    ownerAuthId: String(entry.ownerAuthId || ''),
    lastActionAuthId: String(entry.lastActionAuthId || '')
  };
}

function activeEntries(value, now = Date.now()) {
  const result = {};
  Object.keys(value || {}).forEach(key => {
    const entry = normalizeEntry(value[key], key);
    if (!entry) return;
    if (entry.expiresAt && entry.expiresAt <= now) return;
    result[entry.uid] = entry;
  });
  return result;
}

export async function writeAttendanceTemp({
  scope,
  uid,
  checked,
  operatorId,
  source = 'manual',
  requestId,
  updatedAt = Date.now(),
  expiresAt = updatedAt + ATTENDANCE_TEMP_TTL_MS
}) {
  const normalizedUid = String(uid || '').trim().toUpperCase();
  if (!/^LK\d+$/i.test(normalizedUid)) throw new Error('attendance temp UID is invalid');
  const actor = String(operatorId || '').trim();
  const eventId = String(requestId || makeAttendanceRequestId(scope, normalizedUid, actor, updatedAt));
  const user = await ensureAttendanceAuth();
  const writeNow = Date.now();
  const target = ref(rtdb, attendanceTempPath(scope, normalizedUid));
  const result = await runTransaction(target, currentRaw => {
    const current = currentRaw ? normalizeEntry(currentRaw, normalizedUid) : null;
    const decision = buildAttendanceTempEntry({
      current,
      uid: normalizedUid,
      checked: checked === true,
      operatorId: actor,
      source,
      requestId: eventId,
      now: writeNow,
      expiresAt: Math.max(Number(expiresAt) || 0, writeNow + ATTENDANCE_TEMP_TTL_MS)
    });
    if (!decision.accepted) return;
    const ownerChanged = decision.entry.checked
      && (!current || current.ownerId !== decision.entry.ownerId || current.lockedUntil <= decision.entry.updatedAt);
    return {
      ...decision.entry,
      ownerAuthId: decision.entry.checked
        ? (ownerChanged ? user.uid : String(current && current.ownerAuthId || user.uid))
        : '',
      lastActionAuthId: user.uid
    };
  });
  const committed = result.committed === true;
  const finalValue = normalizeEntry(result.snapshot && result.snapshot.exists() ? result.snapshot.val() : null, normalizedUid);
  return finalValue ? { ...finalValue, committed } : { uid: normalizedUid, committed };
}

export async function readAttendanceTemp(scope) {
  await ensureAttendanceAuth();
  const snapshot = await get(ref(rtdb, attendanceTempScopePath(scope)));
  return activeEntries(snapshot.exists() ? snapshot.val() : {});
}

export function subscribeAttendanceTemp(scope, onChange, onError) {
  let unsubscribe = () => {};
  ensureAttendanceAuth()
    .then(() => {
      unsubscribe = onValue(
        ref(rtdb, attendanceTempScopePath(scope)),
        snapshot => onChange(activeEntries(snapshot.exists() ? snapshot.val() : {})),
        onError
      );
    })
    .catch(onError);
  return () => unsubscribe();
}
