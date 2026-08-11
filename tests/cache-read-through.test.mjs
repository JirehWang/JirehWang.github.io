import assert from 'node:assert/strict';
import test from 'node:test';

import { createReadThrough } from '../firebase/cache-read-through.mjs';

test('Firebase hit returns before GAS loader is called', async () => {
  const cache = createReadThrough();
  let gasCalls = 0;
  const result = await cache.getOrLoad(
    'getGroups/_default',
    async () => ({ status: 'success', data: ['firebase'] }),
    async () => { gasCalls += 1; return { status: 'success', data: ['gas'] }; }
  );

  assert.equal(gasCalls, 0);
  assert.equal(result.source, 'cache');
  assert.deepEqual(result.value.data, ['firebase']);
});

test('Firebase miss shares exactly one GAS load and has no browser write step', async () => {
  const cache = createReadThrough();
  let gasCalls = 0;
  const read = async () => null;
  const load = async () => {
    gasCalls += 1;
    await new Promise(resolve => setTimeout(resolve, 5));
    return { status: 'success', data: ['gas'] };
  };

  const [one, two] = await Promise.all([
    cache.getOrLoad('getGroups/_default', read, load),
    cache.getOrLoad('getGroups/_default', read, load)
  ]);

  assert.equal(gasCalls, 1);
  assert.equal(one.source, 'fresh');
  assert.strictEqual(one, two);
});

test('Firebase read failure reaches GAS once and does not retry the loader', async () => {
  const cache = createReadThrough();
  let gasCalls = 0;
  await assert.rejects(
    cache.getOrLoad(
      'getGroups/_default',
      async () => { throw new Error('firebase unavailable'); },
      async () => { gasCalls += 1; return { status: 'success' }; }
    ),
    /firebase unavailable/
  );
  assert.equal(gasCalls, 0, 'caller owns one direct-GAS fallback after a read failure');
});

test('a rejected GAS loader is invoked once and is not retried by single-flight', async () => {
  const cache = createReadThrough();
  let gasCalls = 0;
  await assert.rejects(
    cache.getOrLoad(
      'getGroups/_default',
      async () => null,
      async () => {
        gasCalls += 1;
        throw new Error('GAS JSON failure');
      }
    ),
    /GAS JSON failure/
  );
  assert.equal(gasCalls, 1);
});
