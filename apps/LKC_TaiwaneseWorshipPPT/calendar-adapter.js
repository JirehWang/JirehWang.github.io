(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TaiwaneseWorshipCalendarAdapter = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  const aliases = {
    call: ['宣召'],
    sermonTitle: ['講道題目', '講道'],
    sermonSpeaker: ['講員', '講道者'],
    scripture: ['經文', '講道經文'],
    verse: ['金句'],
    responsiveReading: ['啟應文', '啟應'],
    hymn1: ['聖詩第一首', '聖詩一', '聖詩1'],
    hymn2: ['聖詩第二首', '聖詩二', '聖2'],
    doxology: ['頌榮']
  };
  const clean = value => String(value || '').trim();
  const getValue = (event, names) => {
    const values = Array.isArray(event && event.values) ? event.values : [];
    const found = values.find(item => names.includes(clean(item.fieldName)));
    return found ? clean(found.value) : '';
  };
  const hymnNumber = value => (clean(value).match(/\d+/) || [''])[0];
  function applyCalendarEvent(event, model) {
    const put = (id, key, value) => { if (value && model[id]) model[id][key] = value; };
    put('call', 'body', getValue(event, aliases.call));
    put('sermon', 'title', getValue(event, aliases.sermonTitle));
    put('sermon', 'kicker', getValue(event, aliases.sermonSpeaker));
    put('scripture', 'body', getValue(event, aliases.scripture));
    put('verse', 'body', getValue(event, aliases.verse));
    put('response', 'body', getValue(event, aliases.responsiveReading));
    put('hymn-1', 'title', hymnNumber(getValue(event, aliases.hymn1)));
    put('hymn-2', 'title', hymnNumber(getValue(event, aliases.hymn2)));
    put('doxology', 'title', hymnNumber(getValue(event, aliases.doxology)));
    return model;
  }
  return { applyCalendarEvent };
});
