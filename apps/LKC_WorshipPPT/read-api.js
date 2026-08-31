(function(root, factory) {
  const api = factory(root);
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipReadApi = api;
  root.worshipReadAPI = api.read;
})(typeof globalThis !== 'undefined' ? globalThis : this, function(root) {
  let callbackSequence = 0;
  const JSONP_READ_ACTIONS = new Set([
    'cal_getEvents',
    'cal_getPptLibraryIndex',
    'cal_getPptLibraryFile',
    'cal_queryBible'
  ]);

  function buildJsonpUrl(endpoint, action, data, token, callbackName) {
    const url = new URL(endpoint);
    url.searchParams.set('action', action);
    url.searchParams.set('token', token || '');
    url.searchParams.set('data', JSON.stringify(data || {}));
    url.searchParams.set('callback', callbackName);
    url.searchParams.set('_lkc', `${Date.now()}_${callbackName}`);
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

  function isGithubPages() {
    const hostname = String(root.location && root.location.hostname || '').toLowerCase();
    return hostname === 'github.io' || hostname.endsWith('.github.io');
  }

  function shouldPreferJsonp(action) {
    if (!JSONP_READ_ACTIONS.has(action) || !root.location) return false;
    return root.location.protocol === 'file:' || isGithubPages();
  }

  function isJsonTransportError(error) {
    if (!error) return false;
    if (error.name === 'SyntaxError' || error.type === 'INVALID_RESPONSE') return true;
    const message = String(error.message || error).toLowerCase();
    return /failed to fetch|network|load failed|unexpected token|not valid json|http\s+4\d\d|http\s+5\d\d/.test(message);
  }

  async function read(action, data) {
    if (root.WorshipPptSupabaseService && typeof root.WorshipPptSupabaseService[action] === 'function') {
      try {
        const res = await root.WorshipPptSupabaseService[action](data || {});
        if (res !== null) return res;
      } catch (err) {
        console.warn(`[WorshipSupabase] Error calling ${action}, falling back:`, err);
      }
    }

    if (!root.GAS_URL) throw new Error('行事曆雲端網址尚未就緒');
    const useJsonpFirst = shouldPreferJsonp(action);
    let jsonpError = null;
    if (useJsonpFirst) {
      try {
        return await jsonp(root.GAS_URL, action, data || {}, root.AUTH_TOKEN || 'ChurchApp-2026');
      } catch (error) {
        jsonpError = error;
        if (root.location && root.location.protocol === 'file:') throw error;
      }
    }
    if (!useJsonpFirst || jsonpError) {
      try {
        if (root.ensureAPIReady) await root.ensureAPIReady();
        if (typeof root.churchAPI === 'function') {
          return await root.churchAPI(action, data || {});
        }
      } catch (error) {
        if (!isJsonTransportError(error)) throw error;
        if (jsonpError) throw jsonpError;
      }
    }
    if (jsonpError) throw jsonpError;
    return jsonp(root.GAS_URL, action, data || {}, root.AUTH_TOKEN || 'ChurchApp-2026');
  }

  return { buildJsonpUrl, jsonp, read };
});
