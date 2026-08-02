(function(root) {
  'use strict';

  const FIELD_CONFIGS = [
    { selector: '.ann-input', prefix: 'announcements' },
    { selector: '.church-news-input', prefix: 'churchNews' },
    { selector: '#prayerHome', id: 'prayer.homeRest' },
    { selector: '#prayerHospital', id: 'prayer.hospital' },
    { selector: '#prayerOther', id: 'prayer.other' }
  ];

  function buildProofreadingPayload(fields) {
    return (Array.isArray(fields) ? fields : [])
      .filter(field => field && field.id != null && String(field.text || '').trim())
      .map(field => ({
        id: String(field.id),
        text: String(field.text || '').trim()
      }));
  }

  function getResponseItems(response) {
    if (Array.isArray(response)) return response;
    if (response && Array.isArray(response.suggestions)) return response.suggestions;
    if (response && Array.isArray(response.data)) return response.data;
    return [];
  }

  function normalizeProofreadingResponse(response, requestedFields) {
    const requested = buildProofreadingPayload(requestedFields);
    const requestedById = new Map(requested.map(field => [field.id, field]));
    const responseById = new Map();

    getResponseItems(response).forEach(item => {
      const id = item && item.id != null ? String(item.id) : '';
      if (requestedById.has(id) && !responseById.has(id)) responseById.set(id, item);
    });

    return requested.map(field => {
      const item = responseById.get(field.id) || {};
      const suggestion = typeof item.suggestion === 'string' && item.suggestion.trim()
        ? item.suggestion.trim()
        : field.text;
      const changed = typeof item.changed === 'boolean'
        ? item.changed
        : suggestion !== field.text;
      return {
        id: field.id,
        text: field.text,
        suggestion,
        changed,
        note: typeof item.note === 'string' ? item.note.trim() : ''
      };
    });
  }

  function getFieldId(textarea, config, index) {
    if (config.id) return config.id;
    return `${config.prefix}.${index}`;
  }

  function createSuggestionSlot(document, textarea) {
    const layout = document.createElement('div');
    layout.className = 'ai-field-layout';

    const originalColumn = document.createElement('div');
    originalColumn.className = 'ai-original-column';
    const originalLabel = document.createElement('div');
    originalLabel.className = 'ai-field-label';
    originalLabel.textContent = '原文';
    originalColumn.appendChild(originalLabel);

    const suggestionColumn = document.createElement('div');
    suggestionColumn.className = 'ai-suggestion-column';
    const suggestionLabel = document.createElement('div');
    suggestionLabel.className = 'ai-field-label ai-suggestion-label';
    suggestionLabel.textContent = 'AI 建議修改';

    const suggestionInput = document.createElement('textarea');
    suggestionInput.className = 'ai-suggestion-input';
    suggestionInput.rows = textarea.rows || 2;
    suggestionInput.readOnly = true;
    suggestionInput.placeholder = '按「AI 檢查錯字」後顯示建議';
    suggestionInput.setAttribute('aria-label', 'AI 建議修改');

    const suggestionNote = document.createElement('div');
    suggestionNote.className = 'ai-suggestion-note';

    suggestionColumn.append(suggestionLabel, suggestionInput, suggestionNote);
    const parent = textarea.parentNode;
    parent.insertBefore(layout, textarea);
    layout.append(originalColumn, suggestionColumn);
    originalColumn.appendChild(textarea);

    textarea._aiSuggestionInput = suggestionInput;
    textarea._aiSuggestionNote = suggestionNote;
    textarea._aiSuggestionLayout = layout;
  }

  function getEditableTextareas(document) {
    const fields = [];
    FIELD_CONFIGS.forEach(config => {
      const textareas = Array.from(document.querySelectorAll(config.selector));
      textareas.forEach((textarea, index) => {
        const id = getFieldId(textarea, config, index);
        textarea.dataset.aiFieldId = id;
        if (!textarea.dataset.aiInitialized) {
          createSuggestionSlot(document, textarea);
          textarea.dataset.aiInitialized = 'true';
          textarea.addEventListener('input', () => {
            textarea._aiSuggestionInput.value = '';
            textarea._aiSuggestionNote.textContent = '原文已變更，請重新檢查';
            textarea._aiSuggestionLayout.classList.remove('ai-suggestion-ready');
          });
        }
        fields.push({ id, text: textarea.value });
      });
    });
    return fields;
  }

  function init(document) {
    if (!document) return;
    getEditableTextareas(document);
  }

  function clearSuggestions(document) {
    if (!document) return;
    getEditableTextareas(document).forEach(field => {
      const textarea = document.querySelector(`[data-ai-field-id="${field.id}"]`);
      if (!textarea || !textarea._aiSuggestionInput) return;
      textarea._aiSuggestionInput.value = '';
      textarea._aiSuggestionNote.textContent = '';
      textarea._aiSuggestionLayout.classList.remove('ai-suggestion-ready');
    });
  }

  async function checkAll(options) {
    const opts = options || {};
    const document = opts.document || root.document;
    const button = opts.button || (document && document.getElementById('btnAiProofread'));
    const status = opts.onStatus || function() {};
    const notify = opts.onNotify || function() {};
    const callApi = opts.callApi;
    if (!document || typeof callApi !== 'function') throw new Error('AI 校對服務尚未就緒');

    const requested = buildProofreadingPayload(getEditableTextareas(document));
    if (!requested.length) {
      status('請先輸入至少一個要檢查的欄位');
      notify('請先輸入至少一個要檢查的欄位', 'info');
      return [];
    }

    if (button) button.disabled = true;
    status('AI 檢查中，請稍候…');
    try {
      const response = await callApi({ fields: requested });
      const suggestions = normalizeProofreadingResponse(response, requested);
      suggestions.forEach(item => {
        const textarea = document.querySelector(`[data-ai-field-id="${item.id}"]`);
        if (!textarea || !textarea._aiSuggestionInput) return;
        textarea._aiSuggestionInput.value = item.suggestion;
        textarea._aiSuggestionNote.textContent = item.changed
          ? (item.note || '請人工確認後再採用')
          : (item.note || '原文無明顯錯字，可維持原文');
        textarea._aiSuggestionLayout.classList.add('ai-suggestion-ready');
      });
      status(`AI 檢查完成：${suggestions.length} 個欄位`);
      notify('AI 建議已產生，原文未被修改', 'success');
      return suggestions;
    } catch (error) {
      status('AI 檢查失敗，請稍後重試');
      notify(`AI 檢查失敗：${error.message || error}`, 'error');
      throw error;
    } finally {
      if (button) button.disabled = false;
    }
  }

  const api = {
    buildProofreadingPayload,
    normalizeProofreadingResponse,
    getEditableTextareas,
    init,
    clearSuggestions,
    checkAll
  };

  root.SundayBulletinAI = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
