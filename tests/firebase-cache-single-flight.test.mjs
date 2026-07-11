import assert from 'node:assert/strict';
import test from 'node:test';

import { createSingleFlight } from '../firebase/cache-single-flight.mjs';

test('same cache key shares one in-flight loader and clears after completion', async () => {
  const singleFlight = createSingleFlight();
  let calls = 0;
  const loader = async () => {
    calls += 1;
    await new Promise(resolve => setTimeout(resolve, 5));
    return { value: calls };
  };

  const [first, second] = await Promise.all([
    singleFlight.run('getGroups/_default', loader),
    singleFlight.run('getGroups/_default', loader)
  ]);
  const third = await singleFlight.run('getGroups/_default', loader);

  assert.equal(calls, 2);
  assert.strictEqual(first, second);
  assert.deepEqual(third, { value: 2 });
});

test('different keys can load independently', async () => {
  const singleFlight = createSingleFlight();
  const calls = [];
  await Promise.all([
    singleFlight.run('a', async () => calls.push('a')),
    singleFlight.run('b', async () => calls.push('b'))
  ]);
  assert.deepEqual(calls.sort(), ['a', 'b']);
});

test('failed loaders are removed so the next request can retry', async () => {
  const singleFlight = createSingleFlight();
  let attempts = 0;
  await assert.rejects(singleFlight.run('a', async () => {
    attempts += 1;
    throw new Error('temporary');
  }), /temporary/);
  const result = await singleFlight.run('a', async () => {
    attempts += 1;
    return 'ok';
  });
  assert.equal(result, 'ok');
  assert.equal(attempts, 2);
});
