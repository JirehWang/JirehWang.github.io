(function(root, factory) {
  const api = factory(root);
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipReadApi = api;
  root.worshipReadAPI = api.read;
})(typeof globalThis !== 'undefined' ? globalThis : this, function(root) {
  let callbackSequence = 0;

  function buildJsonpUrl(endpoint, action, data, token, callbackName) {
    const url = new URL(endpoint);
    url.searchParams.set('action', action);
    url.searchParams.set('token', token || '');
    url.searchParams.set('data', JSON.stringify(data || {}));
    url.searchParams.set('callback', callbackName);
    return url.toString();
  }

  function jsonp(endpoint, action, data, token) {
    if (typeof document === 'undefined') return Promise.reject(new Error('JSONP 只能在瀏覽器執行'));
    return new Promise((resolve, reject) => {
      const callbackName = `__lkcWorshipJsonp_${Date.now()}_${++callbackSequence}`;
      const script = document.createElement('script');
      const cleanup = () => {
        clearTimeout(timer);
        script.remove();
        try { delete root[callbackName]; } catch (_) { root[callbackName] = undefined; }
      };
      const timer = setTimeout(() => {
        cleanup();
        reject(new Error('雲端行事曆讀取逾時'));
      }, 60000);
      root[callbackName] = result => {
        cleanup();
        if (result && result.success === false) reject(new Error(result.message || '雲端資料讀取失敗'));
        else resolve(result);
      };
      script.async = true;
      script.onerror = () => {
        cleanup();
        reject(new Error('無法連線至雲端行事曆'));
      };
      script.src = buildJsonpUrl(endpoint, action, data, token, callbackName);
      document.head.appendChild(script);
    });
  }

  async function read(action, data) {
    if (root.ensureAPIReady) await root.ensureAPIReady();
    const useJsonpFirst = root.location && root.location.protocol === 'file:';
    if (!useJsonpFirst && typeof root.churchAPI === 'function') {
      try {
        return await root.churchAPI(action, data || {});
      } catch (error) {
        if (!/failed to fetch|network|load failed/i.test(String(error && error.message))) throw error;
      }
    }
    if (!root.GAS_URL) throw new Error('行事曆雲端網址尚未就緒');
    return jsonp(root.GAS_URL, action, data || {}, root.AUTH_TOKEN || 'ChurchApp-2026');
  }

  return { buildJsonpUrl, jsonp, read };
});
