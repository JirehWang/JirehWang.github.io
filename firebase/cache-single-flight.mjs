export function createSingleFlight() {
  const pending = new Map();

  function run(key, loader) {
    if (pending.has(key)) return pending.get(key);
    const promise = Promise.resolve().then(loader);
    pending.set(key, promise);
    promise.finally(() => {
      if (pending.get(key) === promise) pending.delete(key);
    }).catch(() => {});
    return promise;
  }

  return { run };
}
