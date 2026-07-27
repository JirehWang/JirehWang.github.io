(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.ListScrollAnchor = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  function getItems(container, selector) {
    if (!container || typeof container.querySelectorAll !== 'function') return [];
    return Array.prototype.slice.call(container.querySelectorAll(selector));
  }

  function isVisible(item) {
    return item && (!item.style || item.style.display !== 'none');
  }

  function capture(container, selector, preferredKey) {
    if (!container || typeof container.getBoundingClientRect !== 'function') return null;
    var items = getItems(container, selector).filter(isVisible);
    if (!items.length) {
      return { key: '', offset: 0, scrollTop: Number(container.scrollTop) || 0 };
    }

    var preferred = preferredKey === undefined || preferredKey === null
      ? ''
      : String(preferredKey);
    var anchorItem = preferred
      ? items.find(function (item) {
          return item.dataset && item.dataset.scrollKey === preferred;
        })
      : null;
    var containerTop = container.getBoundingClientRect().top;

    if (!anchorItem) {
      anchorItem = items.reduce(function (nearest, item) {
        if (!nearest) return item;
        var itemDistance = Math.abs(item.getBoundingClientRect().top - containerTop);
        var nearestDistance = Math.abs(nearest.getBoundingClientRect().top - containerTop);
        return itemDistance < nearestDistance ? item : nearest;
      }, null);
    }

    return {
      key: anchorItem && anchorItem.dataset ? anchorItem.dataset.scrollKey || '' : '',
      offset: anchorItem ? anchorItem.getBoundingClientRect().top - containerTop : 0,
      scrollTop: Number(container.scrollTop) || 0,
    };
  }

  function restore(container, selector, anchor) {
    if (!container || !anchor || typeof container.getBoundingClientRect !== 'function') return false;
    var key = anchor.key === undefined || anchor.key === null ? '' : String(anchor.key);
    var anchorItem = key
      ? getItems(container, selector).find(function (item) {
          return item.dataset && item.dataset.scrollKey === key;
        })
      : null;

    if (!anchorItem) {
      container.scrollTop = Number(anchor.scrollTop) || 0;
      return false;
    }

    var currentOffset = anchorItem.getBoundingClientRect().top - container.getBoundingClientRect().top;
    container.scrollTop += currentOffset - (Number(anchor.offset) || 0);
    return true;
  }

  return {
    capture: capture,
    restore: restore,
  };
});
