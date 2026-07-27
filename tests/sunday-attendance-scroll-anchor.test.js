const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

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

  assert.match(attendanceSource, /dataset\.scrollKey/);
  assert.match(attendanceSource, /ListScrollAnchor\.capture/);
  assert.match(attendanceSource, /ListScrollAnchor\.restore/);
  assert.match(memberSource, /data-scroll-key/);
  assert.match(memberSource, /ListScrollAnchor\.capture/);
  assert.match(memberSource, /ListScrollAnchor\.restore/);
});
