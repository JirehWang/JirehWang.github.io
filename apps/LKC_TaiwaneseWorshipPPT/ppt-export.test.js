const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { exportWorshipPPTX, getImportedSlideObjects, cleanParagraphProperties } = require('./ppt-export.js');

test('exports slides correctly with mock pptxgenjs', async () => {
  const slides = [];
  let createdPptx;
  class MockPptx {
    constructor() {
      createdPptx = this;
      this.layout = '';
    }
    addSlide() {
      const slide = {
        background: null,
        shapes: [],
        images: [],
        texts: [],
        addShape(type, opts) {
          this.shapes.push({ type, opts });
        },
        addImage(opts) {
          this.images.push(opts);
        },
        addText(text, opts) {
          this.texts.push({ text, opts });
        }
      };
      slides.push(slide);
      return slide;
    }
    writeFile(opts) {
      this.writtenFileOpts = opts;
      return Promise.resolve();
    }
  }

  MockPptx.ShapeType = { rect: 'rect' };

  const mockDeck = [
    { kind: 'cover', sectionId: 'cover', sectionLabel: '台語主日禮拜' },
    { kind: 'score', sectionId: 'hymn-1', sectionLabel: '聖詩一', title: '聖詩第1首', kicker: '第一首' },
    { kind: 'ppt-import', sectionId: 'hymn-1', sectionLabel: '聖詩一', objects: [
      { type: 'image', src: 'data:image/png;base64,123', x: 10, y: 10, w: 20, h: 20 },
      { type: 'text', role: 'content', text: '歌詞內容', x: 15, y: 25, w: 70, h: 50, runs: [{ text: '歌詞內容', fontSize: 18 }] }
    ]}
  ];

  const mockLayoutState = { groups: {}, pageAssignments: {} };
  const mockProduction = {
    layoutForPage: () => ({
      titleSize: 60, titleX: 10, titleY: 6, titleW: 80, titleH: 16, titleAlign: 'center', titleColor: '#111111',
      contentSize: 48, contentX: 8, contentY: 24, contentW: 84, contentH: 68, contentAlign: 'left', contentColor: '#111111', lineSpacing: 1.5
    })
  };

  const mockModel = {
    'hymn-1': { opacity: 50 }
  };

  // Run export
  const exportOptions = {
    PptxGenJS: MockPptx,
    getDeckEntries: () => mockDeck,
    layoutState: mockLayoutState,
    production: mockProduction,
    model: mockModel,
    backgroundColor: '#ffffff',
    backgroundImage: '',
    serviceDate: '2026-07-14'
  };

  // Mock global/root variables
  globalThis.hymnOpacitySectionIds = ['hymn-1'];

  await exportWorshipPPTX(exportOptions);

  assert.equal(createdPptx.layout, 'LAYOUT_WIDE');
  assert.equal(slides.length, 3);
  
  // Verify slide 1: cover
  assert.equal(slides[0].texts.length, 2);
  assert.equal(slides[0].texts[0].text, '台語主日禮拜');
  assert.equal(slides[0].texts[1].text, '主後2026年07月14日');

  // Verify slide 2: score (hymn-1) has white overlay shape
  assert.equal(slides[1].shapes.length, 2);
  assert.equal(slides[1].shapes[0].type, 'rect');
  assert.equal(slides[1].shapes[0].opts.fill.transparency, 50); // 100 - opacity(50) = 50
  assert.equal(slides[1].shapes[0].opts.w, 13.333);
  assert.equal(slides[1].shapes[0].opts.h, 7.5);
  assert.equal(slides[1].texts[0].text, '聖詩第1首');

  // Verify slide 3: ppt-import
  assert.equal(slides[2].images.length, 1);
  assert.equal(slides[2].images[0].data, 'data:image/png;base64,123');
  assert.equal(slides[2].texts.length, 1);
  assert.equal(slides[2].texts[0].text[0].text, '歌詞內容');
});

test('uses safe default bounds when an ungrouped page has no stored layout parameters', async () => {
  const slides = [];
  class MockPptx {
    addSlide() {
      const slide = {
        texts: [],
        addText(text, opts) { this.texts.push({ text, opts }); },
        addShape() {},
        addImage() {}
      };
      slides.push(slide);
      return slide;
    }
    writeFile() { return Promise.resolve(); }
  }

  await exportWorshipPPTX({
    PptxGenJS: MockPptx,
    getDeckEntries: () => [{ kind: 'cover', sectionId: 'cover', sectionLabel: '台語主日禮拜' }],
    layoutState: { groups: {}, pageAssignments: {} },
    production: { layoutForPage: () => ({}) },
    model: {},
    backgroundColor: '#ffffff',
    serviceDate: '2026-07-12'
  });

  assert.equal(slides[0].texts.length, 2);
  slides[0].texts.forEach(({ opts }) => {
    for (const key of ['x', 'y', 'w', 'h']) assert.equal(Number.isFinite(opts[key]), true, `${key} must be finite`);
    assert.ok(opts.w > 0, 'text width must be positive');
    assert.ok(opts.h > 0, 'text height must be positive');
  });
});

test('preserves source bounds for an ungrouped imported PowerPoint slide', async () => {
  const slides = [];
  class MockPptx {
    addSlide() {
      const slide = {
        texts: [],
        addText(text, opts) { this.texts.push({ text, opts }); },
        addShape() {},
        addImage() {}
      };
      slides.push(slide);
      return slide;
    }
    writeFile() { return Promise.resolve(); }
  }

  await exportWorshipPPTX({
    PptxGenJS: MockPptx,
    getDeckEntries: () => [{
      kind: 'ppt-import',
      sectionId: 'hymn-1',
      objects: [{
        type: 'text',
        role: 'content',
        text: 'source text',
        x: 15,
        y: 25,
        w: 70,
        h: 50,
        fontSize: 18,
        runs: [{ text: 'source text', fontSize: 18 }]
      }]
    }],
    layoutState: { groups: {}, pageAssignments: {} },
    production: { layoutForPage: () => ({}) },
    model: {},
    backgroundColor: '#ffffff',
    serviceDate: '2026-07-12'
  });

  const [{ text, opts }] = slides[0].texts;
  assert.equal(opts.x, 1.99995);
  assert.equal(opts.y, 1.875);
  assert.equal(opts.w, 9.3331);
  assert.equal(opts.h, 3.75);
  assert.equal(text[0].options.fontSize, 18);
});

test('the export button forwards the active background and model state', () => {
  const appSource = fs.readFileSync(path.join(__dirname, 'app.js'), 'utf8').replace(/\s+/g, '');
  assert.match(
    appSource,
    /exportWorshipPPTX\(\{model,backgroundColor,backgroundImage\}\)/,
    'app.js must pass its lexical state to the standalone exporter'
  );
});

test('cleanParagraphProperties removes duplicate a:pPr tags from a:p paragraphs', () => {
  const originalXml = `
    <a:p>
      <a:pPr indent="0" marL="0"><a:buNone/></a:pPr>
      <a:r><a:t>聖詩 </a:t></a:r>
      <a:pPr indent="0" marL="0"><a:buNone/></a:pPr>
      <a:r><a:t>065</a:t></a:r>
      <a:pPr indent="0" marL="0"><a:buNone/></a:pPr>
      <a:r><a:t> 首</a:t></a:r>
    </a:p>
  `;
  const expectedXml = `
    <a:p>
      <a:pPr indent="0" marL="0"><a:buNone/></a:pPr>
      <a:r><a:t>聖詩 </a:t></a:r>
      
      <a:r><a:t>065</a:t></a:r>
      
      <a:r><a:t> 首</a:t></a:r>
    </a:p>
  `;
  const result = cleanParagraphProperties(originalXml);
  assert.equal(result.replace(/\s+/g, ''), expectedXml.replace(/\s+/g, ''));
});
