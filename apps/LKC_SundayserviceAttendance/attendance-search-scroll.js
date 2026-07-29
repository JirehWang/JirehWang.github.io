(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.AttendanceSearchScroll = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  var PREFIX = 'LKC:attendance-search-scroll:';

  function getKey(attendanceType, dateValue) {
    return PREFIX + String(attendanceType || '') + ':' + String(dateValue || '');
  }

  function save(storage, key, anchor) {
    if (!storage || !key || !anchor) return false;
    try {
      storage.setItem(key, JSON.stringify(anchor));
      return true;
    } catch (error) {
      return false;
    }
  }

  function consume(storage, key) {
    if (!storage || !key) return null;
    try {
      var raw = storage.getItem(key);
      if (!raw) return null;
      storage.removeItem(key);
      var anchor = JSON.parse(raw);
      return anchor && typeof anchor === 'object' ? anchor : null;
    } catch (error) {
      try { storage.removeItem(key); } catch (removeError) { /* storage unavailable */ }
      return null;
    }
  }

  return {
    getKey: getKey,
    save: save,
    consume: consume,
  };
});
