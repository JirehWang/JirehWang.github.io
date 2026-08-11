import assert from 'node:assert/strict';
import test from 'node:test';

import {
  CACHE_SCHEMA_VERSION,
  evaluateCacheEntry
} from '../firebase/cache-entry-contract.mjs';

test('current Firebase cache entry is accepted without calling the backend', () => {
  const entry = {
    value: { status: 'success', data: ['cached'] },
    schemaVersion: CACHE_SCHEMA_VERSION,
    generation: 3,
    updatedAt: 1
  };

  assert.deepEqual(evaluateCacheEntry(entry, 100), {
    hit: true,
    value: entry.value,
    reason: 'cache'
  });
});

test('legacy, expired, and invalid entries are cache misses', () => {
  const current = {
    value: { status: 'success' },
    schemaVersion: CACHE_SCHEMA_VERSION,
    generation: 1
  };

  assert.equal(evaluateCacheEntry({ value: current.value }, 100).hit, false, 'legacy entry');
  assert.equal(evaluateCacheEntry({ ...current, generation: 0 }, 100).hit, false, 'missing generation');
  assert.equal(evaluateCacheEntry({ ...current, expiresAt: 99 }, 100).hit, false, 'expired entry');
  assert.equal(evaluateCacheEntry({ ...current, value: { status: 'error' } }, 100).hit, false, 'error response');
});
