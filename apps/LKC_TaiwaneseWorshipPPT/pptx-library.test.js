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

test('reports the PPTX download stage when the browser fetch fails', async () => {
  const previousFetch = global.fetch;
  global.fetch = async () => { throw new TypeError('Failed to fetch'); };
  try {
    await assert.rejects(
      library.downloadAndParse({ fileId: 'test-file' }, {}),
      /PPTX 下載失敗.*Failed to fetch/
    );
  } finally {
    global.fetch = previousFetch;
  }
});

test('decodes a GAS Base64 PPTX response without using Drive fetch', async () => {
  const bytes = library.base64ToArrayBuffer(Buffer.from([80, 75, 3, 4]).toString('base64'));
  assert.deepEqual(Array.from(new Uint8Array(bytes)), [80, 75, 3, 4]);
});

test('resolves PowerPoint theme text colors so responsive-reading roles stay distinct', () => {
  const theme = { dk1: '#000000', dk2: '#0E2841' };
  const colorMap = { tx1: 'dk1', tx2: 'dk2' };
  assert.equal(library.resolveSchemeColor('tx1', theme, colorMap), '#000000');
  assert.equal(library.resolveSchemeColor('tx2', theme, colorMap), '#0E2841');
});

test('prefers a Firebase Storage download URL over the GAS Base64 endpoint', async () => {
  const previousFetch = global.fetch;
  let fetchedUrl = '';
  let gasCalls = 0;
  global.fetch = async url => {
    fetchedUrl = url;
    return { ok: true, arrayBuffer: async () => new ArrayBuffer(4) };
  };
  const jszip = { loadAsync: async () => { throw new Error('storage payload reached parser'); } };
  try {
    await assert.rejects(
      library.downloadAndParse(
        { fileId: 'h65', downloadUrl: 'https://firebasestorage.googleapis.com/example.pptx' },
        jszip,
        async () => { gasCalls += 1; throw new Error('GAS must not run'); }
      )
    );
    assert.equal(fetchedUrl, 'https://firebasestorage.googleapis.com/example.pptx');
    assert.equal(gasCalls, 0);
  } finally {
    global.fetch = previousFetch;
  }
});

test('uses an inherited slide-layout font size when a placeholder run omits sz', () => {
  assert.deepEqual(library.inheritRunStyle({ bold: true }, 60), { bold: true, fontSize: 60 });
  assert.deepEqual(library.inheritRunStyle({ fontSize: 48 }, 60), { fontSize: 48 });
});

test('rasterizes an imported library page into one transparent full-slide image', async () => {
  const calls = [];
  const context = {
    clearRect: (...args) => calls.push(['clearRect', ...args]),
    drawImage: (...args) => calls.push(['drawImage', ...args]),
    fillText: (...args) => calls.push(['fillText', ...args]),
    measureText: text => ({ width: String(text).length * 20 }),
    save() {}, restore() {},
    set font(value) { calls.push(['font', value]); },
    set fillStyle(value) { calls.push(['fillStyle', value]); },
    set textBaseline(value) { calls.push(['textBaseline', value]); }
  };
  const canvas = {
    width: 0, height: 0,
    getContext: () => context,
    toDataURL: type => `data:${type};base64,rasterized`
  };
  const pages = [{
    id: 'imported:1', kind: 'ppt-import', sourceWidth: 12192000, sourceHeight: 6858000,
    objects: [
      { type: 'image', src: 'data:image/png;base64,score', x: 5, y: 10, w: 90, h: 35 },
      { type: 'text', x: 10, y: 50, w: 80, h: 20, align: 'center', verticalAlign: 'center', fontSize: 36, runs: [{ text: '歌詞', fontSize: 36, color: '#ff0000' }] }
    ]
  }];

  const result = await library.rasterizeImportedPages(pages, {
    width: 1600,
    createCanvas: () => canvas,
    loadImage: async src => ({ src })
  });

  assert.equal(canvas.width, 1600);
  assert.equal(canvas.height, 900);
  assert.deepEqual(result[0].objects, [{
    type: 'image', src: 'data:image/png;base64,rasterized', x: 0, y: 0, w: 100, h: 100
  }]);
  assert.equal(result[0].rasterized, true);
  assert.ok(calls.some(call => call[0] === 'drawImage'));
  assert.ok(calls.some(call => call[0] === 'fillText' && call[1] === '歌詞'));
});
