// 草稿管理 - 教會週報管理系統
// 雲端（GAS）為主要儲存，localStorage 為本地備援快取
// 設定雲端位置：在 js/config.js 的 GAS_SYNC_URL 填入網址即可

const DraftManager = {

  // ============================================================
  // 公開 API
  // ============================================================

  // 儲存草稿（雲端優先，失敗則僅存本地）
  async save(data) {
    if (!data.date) {
      console.warn('[Draft] 無日期，無法儲存');
      return { success: false, error: '請先選擇週報日期' };
    }

    const key = CONFIG.DRAFT_KEY_PREFIX + data.date;
    const payload = { ...data, updatedAt: new Date().toISOString() };

    // 先存本地快取（確保離線也有備份）
    this._saveLocal(key, payload);

    // 若有雲端 URL，以雲端為主
    if (CONFIG.GAS_SYNC_URL) {
      try {
        const result = await this._saveCloud(key, payload);
        console.log('[Draft] 雲端儲存成功:', key);
        return { success: true, location: 'cloud', key, updatedAt: result.updatedAt };
      } catch (err) {
        console.warn('[Draft] 雲端儲存失敗，已存入本地快取:', err.message);
        return { success: true, location: 'local-only', key, warning: '雲端暫時無法連線，草稿已存入本地快取' };
      }
    }

    return { success: true, location: 'local', key };
  },

  // 載入草稿（雲端優先，失敗則從本地快取取）
  async load(date) {
    const key = CONFIG.DRAFT_KEY_PREFIX + date;

    if (CONFIG.GAS_SYNC_URL) {
      try {
        const cloudData = await this._loadCloud(key);
        if (cloudData) {
          // 同步回本地快取
          this._saveLocal(key, cloudData);
          console.log('[Draft] 從雲端載入:', key);
          return { success: true, location: 'cloud', data: cloudData };
        }
      } catch (err) {
        console.warn('[Draft] 雲端載入失敗，嘗試本地快取:', err.message);
      }
    }

    // 從本地快取取
    const localData = this._loadLocal(key);
    if (localData) {
      console.log('[Draft] 從本地快取載入:', key);
      return { success: true, location: 'local', data: localData };
    }

    return { success: false, error: '找不到此日期的草稿' };
  },

  // 列出草稿（雲端優先）
  async list() {
    if (CONFIG.GAS_SYNC_URL) {
      try {
        const cloudList = await this._listCloud();
        console.log('[Draft] 從雲端取得草稿列表');
        return { success: true, location: 'cloud', drafts: cloudList };
      } catch (err) {
        console.warn('[Draft] 雲端列表失敗，改用本地快取:', err.message);
      }
    }

    const localList = this._listLocal();
    return { success: true, location: 'local', drafts: localList };
  },

  // 刪除草稿（同時刪除雲端與本地）
  async delete(date) {
    const key = CONFIG.DRAFT_KEY_PREFIX + date;
    this._deleteLocal(key);

    if (CONFIG.GAS_SYNC_URL) {
      try {
        await this._deleteCloud(key);
        console.log('[Draft] 雲端刪除成功:', key);
      } catch (err) {
        console.warn('[Draft] 雲端刪除失敗:', err.message);
      }
    }

    return { success: true };
  },

  // ============================================================
  // 雲端操作（GAS Web App）
  // ============================================================

  async _saveCloud(key, data) {
    const res = await fetch(CONFIG.GAS_SYNC_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' }, // GAS 需用 text/plain 避免 CORS preflight
      body: JSON.stringify({ key, data })
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (!json.success) throw new Error(json.error || '雲端回傳失敗');
    return json;
  },

  async _loadCloud(key) {
    const url = `${CONFIG.GAS_SYNC_URL}?action=load&key=${encodeURIComponent(key)}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (!json.success) return null;
    return json.data;
  },

  async _listCloud() {
    const url = `${CONFIG.GAS_SYNC_URL}?action=list`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (!json.success) throw new Error(json.error);
    return json.drafts || [];
  },

  async _deleteCloud(key) {
    const res = await fetch(CONFIG.GAS_SYNC_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({ action: 'delete', key })
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (!json.success) throw new Error(json.error || '刪除失敗');
    return json;
  },

  // ============================================================
  // 本地快取操作（localStorage）
  // ============================================================

  _saveLocal(key, data) {
    try {
      localStorage.setItem(key, JSON.stringify(data));
      this._updateLocalIndex(data.date);
    } catch (err) {
      console.warn('[Draft] localStorage 寫入失敗:', err.message);
    }
  },

  _loadLocal(key) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  },

  _deleteLocal(key) {
    try {
      const date = key.replace(CONFIG.DRAFT_KEY_PREFIX, '');
      localStorage.removeItem(key);
      this._removeFromLocalIndex(date);
    } catch (err) {
      console.warn('[Draft] localStorage 刪除失敗:', err.message);
    }
  },

  _listLocal() {
    const index = this._getLocalIndex();
    return index.map(date => {
      const draft = this._loadLocal(CONFIG.DRAFT_KEY_PREFIX + date);
      return {
        key: CONFIG.DRAFT_KEY_PREFIX + date,
        updatedAt: draft?.updatedAt || '',
        preview: draft?.taiwanese?.sermonTitle || draft?.mandarin?.sermonTitle || ''
      };
    }).filter(Boolean);
  },

  _getLocalIndex() {
    try {
      const raw = localStorage.getItem('bulletin_draft_index');
      return raw ? JSON.parse(raw) : [];
    } catch {
      return [];
    }
  },

  _updateLocalIndex(date) {
    let index = this._getLocalIndex().filter(d => d !== date);
    index.unshift(date);
    if (index.length > CONFIG.MAX_DRAFTS) {
      const removed = index.splice(CONFIG.MAX_DRAFTS);
      removed.forEach(d => localStorage.removeItem(CONFIG.DRAFT_KEY_PREFIX + d));
    }
    localStorage.setItem('bulletin_draft_index', JSON.stringify(index));
  },

  _removeFromLocalIndex(date) {
    const index = this._getLocalIndex().filter(d => d !== date);
    localStorage.setItem('bulletin_draft_index', JSON.stringify(index));
  },

  // ============================================================
  // 自動儲存
  // ============================================================

  startAutoSave(getDataFn) {
    this.stopAutoSave();
    this._autoSaveTimer = setInterval(async () => {
      const data = getDataFn();
      if (data?.date) {
        const result = await this.save(data);
        if (result.success) {
          const loc = result.location === 'cloud' ? '☁️ 雲端' : '💾 本地';
          console.log(`[Draft] 自動儲存完成 (${loc}):`, data.date);
        }
      }
    }, CONFIG.AUTO_SAVE_INTERVAL);

    // 頁面卸載時停止 timer，避免關閉/切頁後仍在背景發 fetch
    if (!this._unloadHooked) {
      window.addEventListener('pagehide', () => this.stopAutoSave(), { once: true });
      this._unloadHooked = true;
    }
  },

  stopAutoSave() {
    if (this._autoSaveTimer) {
      clearInterval(this._autoSaveTimer);
      this._autoSaveTimer = null;
    }
  }
};
