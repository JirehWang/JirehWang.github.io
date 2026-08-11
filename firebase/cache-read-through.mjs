import { createSingleFlight } from './cache-single-flight.mjs';

// Coordinates one browser request per cache key. It deliberately has no write
// callback: the loader is GAS and GAS is responsible for Firebase write-through.
export function createReadThrough(singleFlight = createSingleFlight()) {
  async function getOrLoad(key, read, loader) {
    return singleFlight.run(key, async () => {
      const cached = await read();
      if (cached !== null && cached !== undefined) {
        return { value: cached, source: 'cache' };
      }
      return { value: await loader(), source: 'fresh' };
    });
  }

  return { getOrLoad };
}
