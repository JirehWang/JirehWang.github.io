(function(root, factory) {
  const api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root) root.AdminCacheCoordinator = api;
})(typeof window !== 'undefined' ? window : globalThis, function() {
  'use strict';

  function uniqueValues(values) {
    return Array.from(new Set((values || []).filter(Boolean)));
  }

  async function mapWithConcurrency(items, concurrency, worker) {
    const queue = Array.from(items || []);
    if (queue.length === 0) return [];
    const limit = Math.max(1, Math.min(Number(concurrency) || 1, queue.length));
    const results = new Array(queue.length);
    let cursor = 0;

    async function runWorker() {
      while (cursor < queue.length) {
        const index = cursor++;
        results[index] = await worker(queue[index], index);
      }
    }

    await Promise.all(Array.from({ length: limit }, runWorker));
    return results;
  }

  async function refreshCacheGroup(config, dependencies) {
    const topics = uniqueValues(config && config.topics);
    const backendRefresh = config && config.backendRefresh;
    // Firebase is owned by GAS. Backend maintenance actions invalidate/rebuild
    // only after their authoritative Sheet work succeeds; the admin browser
    // must never delete shared cache itself.
    if (backendRefresh) await dependencies.refreshBackend(backendRefresh);
    return { topicCount: topics.length, backendCount: backendRefresh ? 1 : 0 };
  }

  async function refreshAllCacheGroups(groups, dependencies, options) {
    const configs = Object.values(groups || {});
    const topics = uniqueValues(configs.flatMap(config => config.topics || []));
    const backendRefreshes = uniqueValues(configs.map(config => config.backendRefresh));
    await mapWithConcurrency(
      backendRefreshes,
      options && options.concurrency ? options.concurrency : 2,
      dependencies.refreshBackend
    );
    return { topicCount: topics.length, backendCount: backendRefreshes.length };
  }

  return {
    uniqueValues,
    mapWithConcurrency,
    refreshCacheGroup,
    refreshAllCacheGroups
  };
});
