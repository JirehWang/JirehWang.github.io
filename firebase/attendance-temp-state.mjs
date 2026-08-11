export const ATTENDANCE_PENDING_LOCK_MS = 10 * 60 * 1000;
export const ATTENDANCE_TEMP_TTL_MS = 6 * 60 * 60 * 1000;
export const ATTENDANCE_TEMP_SCHEMA_VERSION = 2;

const VALID_SOURCES = new Set(['manual', 'qr']);

function normalizeSource(source) {
  const value = String(source || '').trim().toLowerCase();
  return VALID_SOURCES.has(value) ? value : 'manual';
}

function positiveNumber(value, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : fallback;
}

export function normalizeAttendanceTempEntry(entry, now = Date.now()) {
  const raw = entry && typeof entry === 'object' ? entry : {};
  const updatedAt = positiveNumber(raw.updatedAt, now);
  const expiresAt = positiveNumber(raw.expiresAt, updatedAt + ATTENDANCE_TEMP_TTL_MS);
  const checked = raw.checked === true || raw.state === 'pending';
  const operatorId = String(raw.operatorId || raw.ownerId || '').trim();
  const ownerId = checked ? String(raw.ownerId || operatorId).trim() : '';
  return {
    schemaVersion: Number(raw.schemaVersion || ATTENDANCE_TEMP_SCHEMA_VERSION),
    uid: String(raw.uid || '').trim().toUpperCase(),
    checked,
    state: checked ? 'pending' : 'none',
    source: normalizeSource(raw.source),
    ownerId,
    lastActionBy: String(raw.lastActionBy || operatorId).trim(),
    operatorId,
    requestId: String(raw.requestId || '').trim(),
    revision: Math.max(0, Math.floor(Number(raw.revision) || 0)),
    updatedAt,
    lockedUntil: positiveNumber(raw.lockedUntil, 0),
    expiresAt
  };
}

export function isAttendanceTempLockActive(entry, now = Date.now()) {
  const normalized = normalizeAttendanceTempEntry(entry, now);
  return normalized.checked
    && normalized.lockedUntil > now
    && normalized.expiresAt > now;
}

export function buildAttendanceTempEntry({
  current,
  uid,
  checked,
  operatorId,
  source = 'manual',
  requestId = '',
  now = Date.now(),
  expiresAt = now + ATTENDANCE_TEMP_TTL_MS
}) {
  const currentEntry = current ? normalizeAttendanceTempEntry(current, now) : null;
  const actor = String(operatorId || '').trim();
  const nextSource = normalizeSource(source);
  const lockActive = currentEntry && isAttendanceTempLockActive(currentEntry, now);
  const sameOwner = Boolean(actor && currentEntry && currentEntry.ownerId === actor);

  // An active lease belongs to its owner. Duplicate checks from other devices
  // are safe no-ops, while an uncheck from another device is rejected.
  if (lockActive && !sameOwner) {
    return {
      accepted: false,
      reason: 'locked-by-other',
      entry: currentEntry
    };
  }

  const revision = (currentEntry ? currentEntry.revision : 0) + 1;
  const base = {
    schemaVersion: ATTENDANCE_TEMP_SCHEMA_VERSION,
    uid: String(uid || (currentEntry && currentEntry.uid) || '').trim().toUpperCase(),
    operatorId: actor,
    lastActionBy: actor,
    requestId: String(requestId || '').trim(),
    revision,
    updatedAt: now,
    expiresAt: Math.max(Number(expiresAt) || 0, now + ATTENDANCE_TEMP_TTL_MS)
  };

  if (checked === true) {
    const preserveOwner = Boolean(currentEntry && currentEntry.checked && sameOwner);
    return {
      accepted: true,
      entry: {
        ...base,
        checked: true,
        state: 'pending',
        source: preserveOwner ? currentEntry.source : nextSource,
        ownerId: preserveOwner ? currentEntry.ownerId : actor,
        lockedUntil: now + ATTENDANCE_PENDING_LOCK_MS
      }
    };
  }

  return {
    accepted: true,
    entry: {
      ...base,
      checked: false,
      state: 'none',
      source: currentEntry ? currentEntry.source : nextSource,
      ownerId: '',
      lockedUntil: 0
    }
  };
}
