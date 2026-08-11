// Shared cache-entry contract.  Browser readers accept only entries written by
// the GAS cache writer; legacy entries safely become a normal cache miss.
export const CACHE_SCHEMA_VERSION = 2;

function isInvalidApiResponse(value) {
  return value && typeof value === 'object' && (
    (Object.prototype.hasOwnProperty.call(value, 'status') && value.status !== 'success') ||
    (Object.prototype.hasOwnProperty.call(value, 'success') && value.success === false)
  );
}

export function evaluateCacheEntry(entry, now = Date.now()) {
  if (!entry || typeof entry !== 'object') {
    return { hit: false, value: null, reason: 'missing' };
  }
  if (entry.expiresAt && entry.expiresAt < now) {
    return { hit: false, value: null, reason: 'expired' };
  }
  if (entry.schemaVersion !== CACHE_SCHEMA_VERSION ||
      !Number.isInteger(entry.generation) || entry.generation < 1) {
    return { hit: false, value: null, reason: 'legacy' };
  }
  if (isInvalidApiResponse(entry.value)) {
    return { hit: false, value: null, reason: 'invalid-response' };
  }
  return { hit: true, value: entry.value, reason: 'cache' };
}
