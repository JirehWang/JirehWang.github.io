const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const AttendanceSearchScroll = require('../apps/LKC_SundayserviceAttendance/attendance-search-scroll.js');

const projectRoot = path.resolve(__dirname, '..');
const ListScrollAnchor = require(path.join(
  projectRoot,
  'apps',
  'LKC_SundayserviceAttendance',
  'list-scroll-anchor.js'
));

function makeItem(key, top, display = 'flex') {
  return {
    dataset: { scrollKey: key },
    style: { display },
    getBoundingClientRect() {
      return { top };
    },
  };
}

function makeContainer(items, top = 100, scrollTop = 240) {
  return {
    scrollTop,
    querySelectorAll() {
      return items;
    },
    getBoundingClientRect() {
      return { top };
    },
  };
}

test('capture prefers the requested visible member and records its viewport offset', () => {
  const items = [
    makeItem('LK00001', 80, 'none'),
    makeItem('LK00002', 132),
    makeItem('LK00003', 220),
  ];
  const container = makeContainer(items);

  assert.deepEqual(
    ListScrollAnchor.capture(container, '[data-scroll-key]', 'LK00003'),
    { key: 'LK00003', offset: 120, scrollTop: 240 }
  );
});

test('capture uses the visible item nearest the scroll viewport when no key is preferred', () => {
  const items = [
    makeItem('hidden', 105, 'none'),
    makeItem('above', 72),
    makeItem('nearest', 118),
    makeItem('later', 260),
  ];
  const container = makeContainer(items);

  assert.equal(
    ListScrollAnchor.capture(container, '[data-scroll-key]').key,
    'nearest'
  );
});

test('restore keeps the anchored row at the same viewport offset after reflow', () => {
  const items = [makeItem('LK00003', 520)];
  const container = makeContainer(items);

  const restored = ListScrollAnchor.restore(
    container,
    '[data-scroll-key]',
    { key: 'LK00003', offset: 120, scrollTop: 240 }
  );

  assert.equal(restored, true);
  assert.equal(container.scrollTop, 540);
});

test('attendance and member-list flows both use stable scroll keys and restore anchors', () => {
  const attendanceSource = fs.readFileSync(
    path.join(projectRoot, 'apps', 'LKC_SundayserviceAttendance', 'attendance.js'),
    'utf8'
  );
  const memberSource = fs.readFileSync(
    path.join(projectRoot, 'apps', 'LKC_SundayserviceAttendance', 'members.html'),
    'utf8'
  );
  const indexSource = fs.readFileSync(
    path.join(projectRoot, 'apps', 'LKC_SundayserviceAttendance', 'index.html'),
    'utf8'
  );

  assert.match(attendanceSource, /dataset\.scrollKey/);
  assert.match(attendanceSource, /ListScrollAnchor\.capture/);
  assert.match(attendanceSource, /ListScrollAnchor\.restore/);
  assert.match(memberSource, /data-scroll-key/);
  assert.match(memberSource, /ListScrollAnchor\.capture/);
  assert.match(memberSource, /ListScrollAnchor\.restore/);
  assert.match(indexSource, /list-scroll-anchor\.js[\s\S]*attendance-search-scroll\.js/);
});

test('attendance search stores the pre-filter anchor and consumes it when clearing search', () => {
  const attendanceSource = fs.readFileSync(
    path.join(projectRoot, 'apps', 'LKC_SundayserviceAttendance', 'attendance.js'),
    'utf8'
  );
  const filterStart = attendanceSource.indexOf('function filterAttList');
  const captureIndex = attendanceSource.indexOf('ListScrollAnchor.capture', filterStart);
  const filterMutationIndex = attendanceSource.indexOf('item.style.display =', filterStart);
  const consumeIndex = attendanceSource.indexOf('AttendanceSearchScroll.consume', filterStart);

  assert.ok(filterStart >= 0, 'filterAttList should exist');
  assert.ok(captureIndex > filterStart && captureIndex < filterMutationIndex);
  assert.ok(consumeIndex > filterStart && consumeIndex < filterMutationIndex);
  assert.match(attendanceSource, /AttendanceSearchScroll\.save/);
  assert.match(attendanceSource, /AttendanceSearchScroll\.getKey/);
});

test('filterAttList restores the scroll position from before the first search keystroke', () => {
  const attendanceSource = fs.readFileSync(
    path.join(projectRoot, 'apps', 'LKC_SundayserviceAttendance', 'attendance.js'),
    'utf8'
  );
  const start = attendanceSource.indexOf('function filterAttList');
  const end = attendanceSource.indexOf('function toggleScanner', start);
  const searchInput = { value: '' };
  const dateInput = { value: '2026-07-29' };
  const items = ['甲一', '乙二', '丙三'].map(name => ({
    style: { display: 'flex' },
    querySelector() { return { innerText: name }; }
  }));
  const scrollArea = {
    scrollTop: 480,
    querySelectorAll() { return items; }
  };
  const storage = new Map();
  const localStorage = {
    getItem(key) { return storage.has(key) ? storage.get(key) : null; },
    setItem(key, value) { storage.set(key, String(value)); },
    removeItem(key) { storage.delete(key); }
  };
  let capturedScrollTop = null;
  let restoredAnchor = null;
  const context = {
    window: {
      localStorage,
      AttendanceSearchScroll,
      ListScrollAnchor: {
        capture(container) {
          capturedScrollTop = container.scrollTop;
          return { key: 'member-乙二', offset: 12, scrollTop: container.scrollTop };
        },
        restore(container, selector, anchor) {
          restoredAnchor = anchor;
          container.scrollTop = anchor.scrollTop;
          return true;
        }
      }
    },
    document: {
      getElementById(id) {
        return id === 'attSearchInput' ? searchInput : (id === 'attendanceDateInput' ? dateInput : null);
      },
      querySelector() { return scrollArea; },
      querySelectorAll() { return items; }
    },
    currentAttType: '禮拜',
    attSearchCacheKey: '',
    attSearchCacheActive: false,
    attSearchMemoryAnchor: null,
    requestAnimationFrame(callback) { callback(); }
  };
  vm.createContext(context);
  vm.runInContext(attendanceSource.slice(start, end), context, { filename: 'attendance.js' });

  searchInput.value = '乙';
  context.filterAttList();
  assert.equal(capturedScrollTop, 480);
  assert.equal(items[0].style.display, 'none');
  assert.equal(items[1].style.display, 'flex');

  scrollArea.scrollTop = 0;
  searchInput.value = '';
  context.filterAttList();
  assert.equal(scrollArea.scrollTop, 480);
  assert.equal(restoredAnchor.scrollTop, 480);
  assert.equal(items.every(item => item.style.display === 'flex'), true);
  assert.equal(storage.size, 0);
});
