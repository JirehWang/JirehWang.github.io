const test = require('node:test');
const assert = require('node:assert/strict');
const library = require('./pptx-library.js');

test('recognizes hymn and responsive-reading database filenames', () => {
  assert.deepEqual(library.parseLibraryFilename('第65首 為著美麗的地面.pptx', 'hymn'), {
    kind: 'hymn', number: '65', title: '為著美麗的地面'
  });
  assert.deepEqual(library.parseLibraryFilename('第306B首 我的生命獻給祢.pptx', 'hymn'), {
    kind: 'hymn', number: '306B', title: '我的生命獻給祢'
  });
  assert.deepEqual(library.parseLibraryFilename('21.pptx', 'response'), {
    kind: 'response', number: '21', title: '啟應文 21'
  });
  assert.equal(library.parseLibraryFilename('說明文件.pptx', 'hymn'), null);
});

test('normalizes calendar numbers without losing a hymn suffix', () => {
  assert.equal(library.normalizeLibraryNumber('第 065 首'), '65');
  assert.equal(library.normalizeLibraryNumber('306b'), '306B');
  assert.equal(library.normalizeLibraryNumber('第 21 篇'), '21');
});

test('matches library entries by kind and normalized number', () => {
  const entries = [
    { kind: 'hymn', number: '65', fileId: 'h65' },
    { kind: 'hymn', number: '306B', fileId: 'h306b' },
    { kind: 'response', number: '21', fileId: 'r21' }
  ];
  assert.equal(library.findLibraryEntry(entries, 'hymn', '第065首').fileId, 'h65');
  assert.equal(library.findLibraryEntry(entries, 'hymn', '306b').fileId, 'h306b');
  assert.equal(library.findLibraryEntry(entries, 'response', '第 21 篇').fileId, 'r21');
});

test('composes PowerPoint group transforms into slide coordinates', () => {
  const parent = library.groupTransform({
    offX: 100, offY: 200, extX: 400, extY: 300,
    chOffX: 10, chOffY: 20, chExtX: 200, chExtY: 150
  }, { sx: 1, sy: 1, tx: 0, ty: 0 });
  assert.deepEqual(parent, { sx: 2, sy: 2, tx: 80, ty: 160 });
  assert.deepEqual(library.mapRect({ x: 20, y: 30, w: 50, h: 25 }, parent), {
    x: 120, y: 220, w: 100, h: 50
  });
});

test('turns EMU rectangles into stable percentages', () => {
  assert.deepEqual(
    library.rectToPercent({ x: 0, y: 685800, w: 6096000, h: 3429000 }, 12192000, 6858000),
    { x: 0, y: 10, w: 50, h: 50 }
  );
});
